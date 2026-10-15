package compiler

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"os"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func excelTypedFixture(t *testing.T) (*vbaproject.Project, spec.FormSpec) {
	t.Helper()
	body, err := os.ReadFile("testdata/typed-properties-excel/vbaProject.bin")
	if err != nil {
		t.Fatal(err)
	}
	sum := sha256.Sum256(body)
	if hex.EncodeToString(sum[:]) != "b2b3849bc01f7375cfbc6445ebbba85d04f7ba05a583941c377e3734078ae37c" {
		t.Fatal("Excel evidence fixture changed; recapture and document runtime validation explicitly")
	}
	project, err := vbaproject.Read(body)
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Forms) != 1 || project.Forms[0].Name != "TypedProperties" {
		t.Fatalf("unexpected forms: %+v", project.Forms)
	}
	assertions := 0
	for _, module := range project.Modules {
		if module.Name == "Main" {
			assertions = strings.Count(module.Source, "AssertValue ")
		}
	}
	if assertions < 120 {
		t.Fatalf("fixture lost its original runtime assertion macro: %d assertions", assertions)
	}
	got, err := projection.Project(project.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	return project, got
}

func TestRealExcelSavedTypedAppearanceEvidence(t *testing.T) {
	project, got := excelTypedFixture(t)
	body, err := os.ReadFile("testdata/typed-properties-excel/generation-source.json")
	if err != nil {
		t.Fatal(err)
	}
	sourceSum := sha256.Sum256(body)
	if hex.EncodeToString(sourceSum[:]) != "3fd36ebdd9b97f08ce5aef4ebb147c0d0e611d6119b9027aa47a91b3540d7265" {
		t.Fatal("runtime generation source changed without recapturing Excel evidence")
	}
	var want spec.FormSpec
	if err = json.Unmarshal(body, &want); err != nil {
		t.Fatal(err)
	}
	if got.Form.Font == nil || !fontMatches(persistenceFont(want.Form.Font), persistenceFont(got.Form.Font)) {
		t.Fatalf("Excel-saved root font differs: %+v", got.Form.Font)
	}
	back, fore := got.Form.BackColor, got.Form.ForeColor
	if back == nil {
		back = &spec.FormSpecColor{Kind: "system", Value: "buttonFace"}
	}
	if fore == nil {
		fore = &spec.FormSpecColor{Kind: "system", Value: "windowText"}
	}
	if !reflect.DeepEqual(back, want.Form.BackColor) || !reflect.DeepEqual(fore, want.Form.ForeColor) {
		t.Fatal("Excel-saved root colors differ")
	}
	for _, expected := range want.Controls {
		var actual spec.FormSpecControl
		for _, control := range got.Controls {
			if control.Name == expected.Name {
				actual = control
				break
			}
		}
		if actual.Name == "" {
			t.Fatalf("missing Excel-saved %s", expected.Name)
		}
		if expected.Font != nil && !fontMatches(persistenceFont(expected.Font), persistenceFont(actual.Font)) {
			t.Fatalf("%s Excel-saved font differs: %+v", expected.Name, actual.Font)
		}
		border := actual.BorderStyle
		if actual.Type == "Image" && border == nil {
			border = new(1)
		}
		if !reflect.DeepEqual(expected.BackColor, actual.BackColor) || !reflect.DeepEqual(expected.ForeColor, actual.ForeColor) || !reflect.DeepEqual(expected.TextAlign, actual.TextAlign) || !reflect.DeepEqual(expected.BorderStyle, border) {
			t.Fatalf("%s Excel-saved appearance differs: %+v", expected.Name, actual)
		}
		model := findControl(project.Forms[0].Controls, expected.Name)
		if model.MultiPage != nil {
			hidden := model.MultiPage.Hidden.Record
			if hidden.Values["ForeColor"] != 0x563412 || hidden.Values["BackColor"] != 0xffeedd {
				t.Fatal("real Excel MultiPage hidden TabStrip color evidence changed")
			}
			model = model.MultiPage.Hidden
		}
		if expected.Font != nil && model.Record != nil && model.Record.TextProps != nil {
			weight, found := model.Record.TextProps.Values["FontWeight"]
			if !found || weight != 700 || actual.Font.Bold == nil || *actual.Font.Bold != (weight > 550) {
				t.Fatalf("%s real Excel bold weight evidence differs: %d", expected.Name, weight)
			}
		}
	}
}

func TestRealExcelSavedTypedPartialTemplatePreservesAppearance(t *testing.T) {
	project, before := excelTypedFixture(t)
	desired := spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: spec.FormSpecForm{Name: "TypedProperties", Font: &spec.FormSpecFont{Underline: new(true)}}}
	for _, control := range before.Controls {
		item := spec.FormSpecControl{ID: control.ID, Name: control.Name, Type: control.Type, ParentID: control.ParentID, ZIndex: control.ZIndex}
		if control.Name == "LabelTyped" {
			item.Font = &spec.FormSpecFont{Italic: new(false)}
		}
		desired.Controls = append(desired.Controls, item)
	}
	result, err := CompileTemplate(project.Forms[0], desired, project.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	after, err := projection.Project(result)
	if err != nil {
		t.Fatal(err)
	}
	after = snapshotCopy(t, after)
	expected := snapshotCopy(t, before)
	expected.Form.Font.Underline = new(true)
	for i := range expected.Controls {
		if expected.Controls[i].Name == "LabelTyped" {
			expected.Controls[i].Font.Italic = new(false)
		}
	}
	if !reflect.DeepEqual(expected.Form, after.Form) {
		t.Fatalf("partial root font replaced another appearance value: got=%+v", after.Form)
	}
	if !reflect.DeepEqual(expected.Controls, after.Controls) {
		for i := range expected.Controls {
			if !reflect.DeepEqual(expected.Controls[i], after.Controls[i]) {
				a, _ := json.Marshal(expected.Controls[i])
				b, _ := json.Marshal(after.Controls[i])
				t.Errorf("control[%d] before=%s after=%s", i, a, b)
			}
		}
		t.Fatal("partial Label font changed another appearance or hierarchy value")
	}
	for _, original := range flattenControls(project.Forms[0].Controls) {
		if original.Name == "LabelTyped" {
			continue
		}
		retained := findControl(result.Controls, original.Name)
		if retained == nil || !reflect.DeepEqual(original.Record, retained.Record) {
			t.Fatalf("partial font edit changed unrelated binary control %s", original.Name)
		}
		if original.MultiPage != nil && !reflect.DeepEqual(original.MultiPage.Hidden.Record, retained.MultiPage.Hidden.Record) {
			t.Fatal("partial font edit changed hidden MultiPage appearance")
		}
	}
	original, _ := oforms.SerializeForm(project.Forms[0], project.Props.CodePage)
	noop, err := CompileEdits(project.Forms[0], before, before, project.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	unchanged, _ := oforms.SerializeForm(noop, project.Props.CodePage)
	assertDesignerBytes(t, original, unchanged)
	repeated, err := CompileTemplate(result, desired, project.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	first, _ := oforms.SerializeForm(result, project.Props.CodePage)
	second, _ := oforms.SerializeForm(repeated, project.Props.CodePage)
	assertDesignerBytes(t, first, second)
}
