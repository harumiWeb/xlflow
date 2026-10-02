package compiler

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// TestGenerateExcelNewFormArtifact is opt-in: it only authors files; the local
// PowerShell gate opens them in Excel. Ordinary tests never start Excel.
func TestGenerateExcelNewFormArtifact(t *testing.T) {
	input, output, expected := os.Getenv("XLFLOW_GENERATION_WORKBOOK"), os.Getenv("XLFLOW_GENERATION_OUTPUT"), os.Getenv("XLFLOW_GENERATION_EXPECTED")
	if input == "" || output == "" || expected == "" {
		t.Skip("set XLFLOW_GENERATION_WORKBOOK/OUTPUT/EXPECTED for the local gate")
	}
	input, _ = filepath.Abs(input)
	output, _ = filepath.Abs(output)
	expected, _ = filepath.Abs(expected)
	if strings.EqualFold(input, output) || strings.EqualFold(expected, input) || strings.EqualFold(expected, output) {
		t.Fatal("outputs must be new, distinct files")
	}
	for _, path := range []string{output, expected} {
		if _, err := os.Stat(path); !os.IsNotExist(err) {
			t.Fatalf("output already exists or cannot be checked: %s", path)
		}
	}
	reader, err := zip.OpenReader(input)
	if err != nil {
		t.Fatal(err)
	}
	defer func() {
		if err := reader.Close(); err != nil {
			t.Error(err)
		}
	}()
	var projectBytes []byte
	for _, entry := range reader.File {
		if entry.Name == "xl/vbaProject.bin" {
			projectBytes = readZipEntry(t, entry)
		}
	}
	project, err := vbaproject.Read(projectBytes)
	if err != nil {
		t.Fatal(err)
	}
	for i := range project.Modules {
		if project.Modules[i].Name == "Main" {
			project.Modules[i].Source = `Attribute VB_Name = "Main"
Option Explicit
Public Sub RunGenerationSentinel()
    On Error GoTo Failed
    Dim generated As Object, blankInstance As Object
    Set generated = VBA.UserForms.Add("GeneratedForm")
    Set blankInstance = VBA.UserForms.Add("GeneratedEmptyForm")
    If Not generated.VerifyGeneration() Or blankInstance.Controls.Count <> 0 Then Err.Raise 5
    ThisWorkbook.Worksheets(1).Range("B1").Value2 = generated.InsideWidth
    ThisWorkbook.Worksheets(1).Range("B2").Value2 = generated.InsideHeight
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-883-ok"
    Unload generated
    Unload blankInstance
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "error:" & Err.Number & ":" & Err.Description
End Sub
`
		}
	}
	// Retain both Excel-authored forms and add independently named forms.
	inputSpec := newSpec()
	inputSpec.Form.Caption = new("issue883-日本語😀")
	// Use a whole export-grid size so VBIDE's twip/grid quantization does not
	// obscure the independently verified binary client-size contract.
	inputSpec.Form.Build = &spec.FormSpecBuildForm{ClientWidth: new(324.0), ClientHeight: new(282.0)}
	for i, kind := range []string{"Label", "TextBox", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "ComboBox", "ListBox", "SpinButton", "ScrollBar", "Image"} {
		control := spec.FormSpecControl{Name: kind + "Main", Type: kind, Left: new(float64(12 + (i%2)*160)), Top: new(float64(12 + (i/2)*44)), Enabled: new(false), Visible: new(true)}
		if _, ok := spec.LookupControlProperty(kind, "caption"); ok {
			control.Caption = new(control.Name + "-日本語😀")
		}
		switch kind {
		case "TextBox":
			control.Text = new("text-日本語😀")
		case "ComboBox":
			control.Value = "combo-text"
		case "CheckBox", "OptionButton", "ToggleButton":
			control.Value = true
		case "SpinButton":
			control.Value = float64(12)
		case "ScrollBar":
			control.Value = float64(34)
		}
		inputSpec.Controls = append(inputSpec.Controls, control)
	}
	code := `Option Explicit
Public Function VerifyGeneration() As Boolean
    VerifyGeneration = Me.Controls.Count = 11 And CBool(Me.ToggleButtonMain.Value) And CLng(Me.SpinButtonMain.Value) = 12 And CLng(Me.ScrollBarMain.Value) = 34
End Function
`
	var results []spec.FormSpec
	for _, authoring := range []spec.FormSpec{inputSpec, {SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: spec.FormSpecForm{Name: "GeneratedEmptyForm"}, Controls: []spec.FormSpecControl{}}} {
		form, err := CompileNew(authoring, project.Props.CodePage)
		if err != nil {
			t.Fatal(err)
		}
		body := "Option Explicit\n"
		if authoring.Form.Name == "GeneratedForm" {
			body = code
		}
		module, err := vbaproject.NewUserFormModule(form, body, project.Props.CodePage)
		if err != nil {
			t.Fatal(err)
		}
		project, err = vbaproject.WithNewUserForm(project, form, module)
		if err != nil {
			t.Fatal(err)
		}
		result, err := projection.Project(form)
		if err != nil {
			t.Fatal(err)
		}
		// These runtime defaults were observed in the saved/reopened Excel
		// baseline. Keep them in the generic observation bag, not build intent:
		// neither property is part of the supported generation input surface.
		for i := range result.Controls {
			control := &result.Controls[i]
			if control.Type != "SpinButton" && control.Type != "ScrollBar" {
				continue
			}
			if control.Observed == nil {
				control.Observed = &spec.FormSpecObservedControl{}
			}
			if control.Observed.Properties == nil {
				control.Observed.Properties = make(map[string]any)
			}
			control.Observed.Properties["Delay"] = 50
			if control.Type == "ScrollBar" {
				control.Observed.Properties["ProportionalThumb"] = true
			}
		}
		results = append(results, result)
	}
	newProject, err := vbaproject.Write(project)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := vbaproject.Read(newProject); err != nil {
		t.Fatal(err)
	}
	var workbook bytes.Buffer
	writer := zip.NewWriter(&workbook)
	for _, entry := range reader.File {
		if entry.Name != "xl/vbaProject.bin" {
			if err := writer.Copy(entry); err != nil {
				t.Fatal(err)
			}
			continue
		}
		stream, err := writer.CreateHeader(&entry.FileHeader)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := stream.Write(newProject); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(output, workbook.Bytes(), 0o600); err != nil {
		t.Fatal(err)
	}
	encoded, err := json.MarshalIndent(struct {
		Forms []spec.FormSpec `json:"forms"`
	}{results}, "", "  ")
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(expected, append(encoded, '\n'), 0o600); err != nil {
		t.Fatal(err)
	}
	t.Logf("generation artifact=%s expected=%s", output, expected)
}

func TestReadExcelNormalizedNewForms(t *testing.T) {
	path, expectedPath := os.Getenv("XLFLOW_GENERATION_NORMALIZED_PROJECT"), os.Getenv("XLFLOW_GENERATION_EXPECTED")
	if path == "" && expectedPath == "" {
		path = filepath.Join("testdata", "generation-excel-generated", "normalized.bin")
		expectedPath = filepath.Join("testdata", "generation-excel-generated", "expected.json")
	} else if path == "" || expectedPath == "" {
		t.Fatal("set both XLFLOW_GENERATION_NORMALIZED_PROJECT and XLFLOW_GENERATION_EXPECTED")
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(data)
	if err != nil {
		t.Fatal(err)
	}
	data, err = os.ReadFile(expectedPath)
	if err != nil {
		t.Fatal(err)
	}
	var expected struct {
		Forms []spec.FormSpec `json:"forms"`
	}
	if err := json.Unmarshal(data, &expected); err != nil {
		t.Fatal(err)
	}
	for _, want := range expected.Forms {
		index := slices.IndexFunc(project.Forms, func(f *oforms.Form) bool { return f.Name == want.Form.Name })
		if index < 0 {
			t.Fatalf("form missing: %s", want.Form.Name)
		}
		got, err := projection.Project(project.Forms[index])
		if err != nil {
			t.Fatal(err)
		}
		if got.Form.Caption == nil || *got.Form.Caption != *want.Form.Caption || len(got.Controls) != len(want.Controls) {
			t.Fatalf("normalized form differs: %s", want.Form.Name)
		}
		if !reflect.DeepEqual(got.Form.Build, want.Form.Build) {
			t.Fatalf("persisted client dimensions differ: %s got=%+v want=%+v", want.Form.Name, got.Form.Build, want.Form.Build)
		}
		for i, control := range want.Controls {
			actual := got.Controls[i]
			wantValue, _ := json.Marshal(control.Value)
			gotValue, _ := json.Marshal(actual.Value)
			if control.Name != actual.Name || control.Type != actual.Type || !bytes.Equal(wantValue, gotValue) || !reflect.DeepEqual(control.Enabled, actual.Enabled) || !reflect.DeepEqual(control.Visible, actual.Visible) || !reflect.DeepEqual(control.Caption, actual.Caption) || !reflect.DeepEqual(control.Text, actual.Text) || !reflect.DeepEqual(control.TabIndex, actual.TabIndex) {
				t.Fatalf("normalized control differs: %s got=%+v", control.Name, actual)
			}
			for _, pair := range [][2]*float64{{actual.Left, control.Left}, {actual.Top, control.Top}, {actual.Width, control.Width}, {actual.Height, control.Height}} {
				if !reflect.DeepEqual(pair[0], pair[1]) {
					t.Fatalf("normalized geometry differs: %s", control.Name)
				}
			}
		}
	}
}

func TestExcelAuthoredGenerationEvidence(t *testing.T) {
	read := func(name string) *vbaproject.Project {
		t.Helper()
		data, err := os.ReadFile(filepath.Join("testdata", "generation-excel-authored", name))
		if err != nil {
			t.Fatal(err)
		}
		project, err := vbaproject.Read(data)
		if err != nil {
			t.Fatal(err)
		}
		return project
	}
	baseline := read("baseline.bin")
	index := slices.IndexFunc(baseline.Forms, func(f *oforms.Form) bool { return f.Name == "GenerationBaseline" })
	if index < 0 || len(baseline.Forms[index].Controls) != 11 {
		t.Fatal("11-control Excel baseline missing")
	}
	if !bytes.Equal(baseline.Forms[index].CompObj.Raw, mustCompileEmpty(t).CompObj.Raw) {
		t.Fatal("generated CompObj differs from Excel")
	}
	for _, kind := range []string{"ToggleButton", "SpinButton", "ScrollBar", "Image"} {
		t.Run(kind, func(t *testing.T) {
			project := read("enabled-" + kind + "Main-False.bin")
			formIndex := slices.IndexFunc(project.Forms, func(f *oforms.Form) bool { return f.Name == "GenerationBaseline" })
			if formIndex < 0 {
				t.Fatal("baseline form missing")
			}
			form := project.Forms[formIndex]
			controlIndex := slices.IndexFunc(form.Controls, func(c *oforms.Control) bool { return c.Name == kind+"Main" })
			if controlIndex < 0 {
				t.Fatal("control missing")
			}
			wantBits := int64(0x19)
			if kind == "ToggleButton" {
				wantBits = 0x2c800819
			}
			if bits := form.Controls[controlIndex].Record.Values["VariousPropertyBits"]; bits != wantBits {
				t.Fatalf("disabled flags=%#x, want %#x", bits, wantBits)
			}
		})
	}
}

func mustCompileEmpty(t *testing.T) *oforms.Form {
	t.Helper()
	form, err := CompileNew(newSpec(), 932)
	if err != nil {
		t.Fatal(err)
	}
	return form
}
