package compiler

import (
	"errors"
	"math"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func newSpec() spec.FormSpec {
	return spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: spec.FormSpecForm{Name: "GeneratedForm"}, Controls: []spec.FormSpecControl{}}
}

func TestCompileNewEmptyAndCommonControls(t *testing.T) {
	for _, codePage := range []uint16{1252, 932} {
		input := newSpec()
		input.Form.Caption = new("Generated Caption")
		input.Form.Build = &spec.FormSpecBuildForm{ClientWidth: new(320.0), ClientHeight: new(240.0)}
		empty, err := CompileNew(input, codePage)
		if err != nil {
			t.Fatal(err)
		}
		if len(empty.Controls) != 0 || empty.Levels[0].Record.Values["NextAvailableID"] != 1 {
			t.Fatal("empty form topology")
		}
		for _, kind := range []string{"Label", "TextBox", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "ComboBox", "ListBox", "SpinButton", "ScrollBar", "Image"} {
			input.Controls = append(input.Controls, spec.FormSpecControl{Name: kind + "Main", Type: kind, Left: new(12.25), Top: new(8.5), Enabled: new(false), Visible: new(false)})
		}
		input.Controls[1].Text = new("text-日本語😀")
		input.Controls[1].Properties = map[string]any{"MaxLength": float64(128), "Tag": "tag-😀", "ControlTipText": "tip"}
		input.Controls[3].Value = true
		input.Controls[4].Value = "True"
		input.Controls[5].Value = true
		input.Controls[8].Value = float64(12)
		input.Controls[9].Value = float64(34)
		input.Controls[2].TabIndex = new(7)
		before := snapshotCopy(t, input)
		first, err := CompileNew(input, codePage)
		if err != nil {
			t.Fatal(err)
		}
		second, err := CompileNew(input, codePage)
		if err != nil {
			t.Fatal(err)
		}
		a, err := oforms.SerializeForm(first, codePage)
		if err != nil {
			t.Fatal(err)
		}
		b, err := oforms.SerializeForm(second, codePage)
		if err != nil {
			t.Fatal(err)
		}
		assertDesignerBytes(t, a, b)
		if !reflect.DeepEqual(before, input) {
			t.Fatal("authoring input changed")
		}
		projected, err := projection.Project(first)
		if err != nil {
			t.Fatal(err)
		}
		if projected.Form.Caption == nil || *projected.Form.Caption != *input.Form.Caption || len(projected.Controls) != 11 {
			t.Fatal("form read-back differs")
		}
		if first.Levels[0].Record.Values["NextAvailableID"] != 12 {
			t.Fatal("invalid next site ID")
		}
		var total uint32
		for i, control := range first.Controls {
			if control.ID != int32(i+1) || control.Depth != 0 || control.SiteType != 1 {
				t.Fatal("invalid site identity")
			}
			total += control.ObjectStreamSize
			got := projected.Controls[i]
			if got.Name != input.Controls[i].Name || got.Type != input.Controls[i].Type || got.Enabled == nil || *got.Enabled || got.Visible == nil || *got.Visible {
				t.Fatalf("control read-back differs: %+v", got)
			}
			if math.Abs(*got.Left-12.25) > 0.02 || math.Abs(*got.Top-8.5) > 0.02 {
				t.Fatal("HIMETRIC conversion differs")
			}
		}
		if int(total) != len(first.Levels[0].ORaw) {
			t.Fatal("object extents differ")
		}
		if projected.Controls[1].Value != *input.Controls[1].Text || projected.Controls[5].Value != "True" {
			t.Fatal("string/boolean read-back differs")
		}
		if first.Controls[8].Record.Values["Position"] != 12 || first.Controls[9].Record.Values["Position"] != 34 {
			t.Fatal("numeric control values differ")
		}
		if *projected.Controls[2].TabIndex != 7 {
			t.Fatal("explicit tab order lost")
		}
	}
}

func TestCompileNewRejectsUnsupportedAndInvalidInput(t *testing.T) {
	for _, tc := range []struct {
		name   string
		code   string
		change func(*spec.FormSpec)
	}{
		{"outer-size", GenerationUnsupported, func(s *spec.FormSpec) { s.Form.Width = new(300.0) }},
		{"size-overflow", GenerationInvalid, func(s *spec.FormSpec) { s.Form.Build = &spec.FormSpecBuildForm{ClientWidth: new(1e100)} }},
		{"nan", GenerationInvalid, func(s *spec.FormSpec) { s.Controls[0].Left = new(math.NaN()) }},
		{"container", GenerationUnsupported, func(s *spec.FormSpec) { s.Controls[0].Type = "Frame" }},
		{"picture", GenerationUnsupported, func(s *spec.FormSpec) {
			s.Controls[0].Type = "Image"
			s.Controls[0].Properties = map[string]any{"Picture": "data"}
		}},
		{"list", GenerationUnsupported, func(s *spec.FormSpec) { s.Controls[0].Type = "ComboBox"; s.Controls[0].List = []string{"x"} }},
		{"selected", GenerationUnsupported, func(s *spec.FormSpec) { s.Controls[0].Type = "ListBox"; s.Controls[0].SelectedIndex = new(0) }},
		{"alias", GenerationConflict, func(s *spec.FormSpec) {
			s.Controls[0].Caption = new("one")
			s.Controls[0].Properties = map[string]any{"Caption": "two"}
		}},
		{"text-value-alias", GenerationConflict, func(s *spec.FormSpec) {
			s.Controls[0].Type = "TextBox"
			s.Controls[0].Text = new("one")
			s.Controls[0].Value = "two"
		}},
		{"numeric-fraction", GenerationInvalid, func(s *spec.FormSpec) { s.Controls[0].Type = "SpinButton"; s.Controls[0].Value = 1.5 }},
		{"invalid-cp", GenerationInvalid, func(s *spec.FormSpec) { s.Form.Name = "日本語" }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			input := newSpec()
			input.Controls = []spec.FormSpecControl{{Name: "ControlMain", Type: "Label"}}
			tc.change(&input)
			result, err := CompileNew(input, 1252)
			detail, ok := errors.AsType[*Error](err)
			if result != nil || !ok || detail.Code != tc.code {
				t.Fatalf("got result=%v error=%v; want %s", result, err, tc.code)
			}
		})
	}
}

func TestCompileNewReportsRootCaptionCodePageContext(t *testing.T) {
	for _, tc := range []struct {
		name     string
		form     spec.FormSpecForm
		property string
	}{
		{name: "legacy", form: spec.FormSpecForm{Name: "GeneratedForm", Caption: new("😀")}, property: "form.caption"},
		{name: "build", form: spec.FormSpecForm{Name: "GeneratedForm", Build: &spec.FormSpecBuildForm{Caption: new("😀")}}, property: "form.build.caption"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			input := newSpec()
			input.Form = tc.form
			_, err := CompileNew(input, 932)
			detail, ok := errors.AsType[*Error](err)
			if !ok || detail.Code != GenerationInvalid || detail.Property != tc.property {
				t.Fatalf("error = %v, want invalid %s", err, tc.property)
			}
		})
	}
}

func TestCompileEditsRejectsClientSizeChanges(t *testing.T) {
	base, before := fixture(t, "p4_form.bin")
	after := snapshotCopy(t, before)
	if after.Form.Build == nil {
		after.Form.Build = &spec.FormSpecBuildForm{}
	}
	after.Form.Build.ClientWidth = new(200.0)
	_, err := CompileEdits(base, before, after, 932)
	if detail, ok := errors.AsType[*Error](err); !ok || detail.Code != Unsupported {
		t.Fatalf("client size edit was discarded: %v", err)
	}
}
