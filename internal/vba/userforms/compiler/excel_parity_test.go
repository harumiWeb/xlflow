package compiler

import (
	"archive/zip"
	"bytes"
	"io"
	"math"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// TestGenerateExcelMutationArtifact prepares the workbook used by the local
// scripts/test-userform-mutation-e2e.ps1 verify gate. Ordinary tests skip it.
func TestGenerateExcelMutationArtifact(t *testing.T) {
	input := os.Getenv("XLFLOW_MUTATION_WORKBOOK")
	output := os.Getenv("XLFLOW_MUTATION_OUTPUT")
	expected := os.Getenv("XLFLOW_MUTATION_EXPECTED")
	if input == "" || output == "" || expected == "" {
		t.Skip("set XLFLOW_MUTATION_WORKBOOK, XLFLOW_MUTATION_OUTPUT, XLFLOW_MUTATION_EXPECTED for local Excel verification")
	}
	input, err := filepath.Abs(input)
	if err != nil {
		t.Fatal(err)
	}
	output, err = filepath.Abs(output)
	if err != nil {
		t.Fatal(err)
	}
	if strings.EqualFold(input, output) {
		t.Fatal("output must not replace the Excel baseline")
	}
	reader, err := zip.OpenReader(input)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := reader.Close(); err != nil {
			t.Error(err)
		}
	})
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
	if len(project.Forms) != 1 {
		t.Fatalf("want one form, got %d", len(project.Forms))
	}
	before, err := projection.Project(project.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	after := snapshotCopy(t, before)
	after.Form.Caption = new("issue882-after-日本語")
	for i := range after.Controls {
		control := &after.Controls[i]
		control.Left = new(*control.Left + 2.25)
		control.Top = new(*control.Top + 1.5)
		control.Width = new(*control.Width + 3.25)
		control.Height = new(*control.Height + 1.5)
		control.Visible = new(false)
		control.Enabled = new(false)
		switch control.Name {
		case "TextMain":
			control.TabIndex = new(2)
		case "ButtonMain":
			control.TabIndex = new(1)
		}
		control.Properties = map[string]any{"Tag": "issue882-tag-日本語", "ControlTipText": "issue882-tip"}
		control.Properties["BackColor"] = float64(0x112233)
		control.Properties["ForeColor"] = float64(0x334455)
		if control.Type == "Label" || control.Type == "TextBox" || control.Type == "ComboBox" || control.Type == "ListBox" || control.Type == "Frame" {
			control.Properties["BorderColor"] = float64(0x556677)
			control.Properties["BorderStyle"] = float64(1)
		}
		if control.Type == "TextBox" || control.Type == "ComboBox" {
			control.Properties["MaxLength"] = float64(128)
		}
		if control.Type == "OptionButton" {
			control.Properties["GroupName"] = "issue882-group"
		}
		switch control.Type {
		case "Label", "CommandButton", "CheckBox", "OptionButton", "Frame":
			control.Caption = new("after-" + control.Name + "-日本語")
		}
		switch control.Type {
		case "TextBox":
			control.Text = new("after-" + control.Name + "-日本語 🐚")
		case "CheckBox", "OptionButton":
			control.Value = true
		case "ComboBox":
			control.Value = "after-combo"
		}
	}
	compiled, err := CompileEdits(project.Forms[0], before, after, project.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	project.Forms[0] = compiled
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
		header := entry.FileHeader
		stream, err := writer.CreateHeader(&header)
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
	projected, err := projection.Project(compiled)
	if err != nil {
		t.Fatal(err)
	}
	for i := range projected.Controls {
		projected.Controls[i].Properties = after.Controls[i].Properties
	}
	body, err := spec.MarshalSnapshot("json", projected)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(expected, body, 0o600); err != nil {
		t.Fatal(err)
	}
	t.Logf("Excel gate artifact=%s expected=%s", output, expected)
}

func readZipEntry(t *testing.T, entry *zip.File) []byte {
	t.Helper()
	reader, err := entry.Open()
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := reader.Close(); err != nil {
			t.Error(err)
		}
	})
	body, err := io.ReadAll(reader)
	if err != nil {
		t.Fatal(err)
	}
	return body
}

// TestReadExcelNormalizedArtifact verifies the model after Excel SaveAs/reopen,
// including UTF-16 text and all explicitly edited property-bag fields.
func TestReadExcelNormalizedArtifact(t *testing.T) {
	path := os.Getenv("XLFLOW_MUTATION_NORMALIZED_PROJECT")
	expectedPath := os.Getenv("XLFLOW_MUTATION_EXPECTED")
	if path == "" || expectedPath == "" {
		t.Skip("set XLFLOW_MUTATION_NORMALIZED_PROJECT and XLFLOW_MUTATION_EXPECTED after the Excel gate")
	}
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(body)
	if err != nil {
		t.Fatal(err)
	}
	expected, err := spec.LoadFormSpec(spec.SpecInput{Path: expectedPath, Format: "json"})
	if err != nil {
		t.Fatal(err)
	}
	actual, err := projection.Project(project.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(expected.Form.Caption, actual.Form.Caption) || len(expected.Controls) != len(actual.Controls) {
		t.Fatal("normalized form differs")
	}
	controls := flattenControls(project.Forms[0].Controls)
	for i, want := range expected.Controls {
		got := actual.Controls[i]
		if want.Name != got.Name || want.ParentID != got.ParentID || want.Type != got.Type {
			t.Fatalf("normalized identity differs for %s", want.Name)
		}
		for _, field := range []string{"caption", "text", "value", "left", "top", "width", "height", "tabindex", "enabled", "visible"} {
			w, _ := snapshotProperty(want, field)
			g, _ := snapshotProperty(got, field)
			if w == nil {
				continue
			}
			if field == "left" || field == "top" || field == "width" || field == "height" {
				if g == nil || math.Abs(w.(float64)-g.(float64)) > 0.05 {
					t.Fatalf("normalized %s.%s geometry differs", want.Name, field)
				}
			} else if !reflect.DeepEqual(w, g) {
				t.Fatalf("normalized %s.%s differs: %v != %v", want.Name, field, g, w)
			}
		}
		for name, value := range want.Properties {
			key := strings.ToLower(name)
			w, err := bagValue(want.Type, key, value)
			if err != nil {
				t.Fatal(err)
			}
			g, found := editedProperty(controls[i], propertyNames[key])
			if !found || !reflect.DeepEqual(w, g) {
				t.Fatalf("normalized %s.%s differs: %v != %v", want.Name, name, g, w)
			}
		}
	}
}
