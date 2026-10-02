package pack

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestBlankFormsJapaneseAndAllCommonControlsReadback(t *testing.T) {
	formSpec := spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: spec.FormSpecForm{Name: "JapaneseForm", Caption: new("日本語😀")}}
	for i, kind := range []string{"Label", "TextBox", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "ComboBox", "ListBox", "SpinButton", "ScrollBar", "Image"} {
		control := spec.FormSpecControl{ID: kind, Name: kind + "Main", Type: kind, TabIndex: new(i)}
		if kind == "TextBox" {
			control.Text = new("日本語😀")
		}
		formSpec.Controls = append(formSpec.Controls, control)
	}
	sources := []SourceModule{
		{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: "Option Explicit"},
		{Name: "Sheet1", Type: ModuleTypeDocument, Source: "Option Explicit"},
		{Name: "JapaneseForm", Type: ModuleTypeForm, FormSpec: &formSpec, Source: "Option Explicit\n' 日本語\n"},
	}
	body, meta, err := BuildBlankWorkbook(sources, BlankOptions{CodePage: 932})
	if err != nil {
		t.Fatal(err)
	}
	if meta.Form != 1 {
		t.Fatalf("forms=%d", meta.Form)
	}
	p, err := vbaproject.Read(mustEntry(t, body, "xl/vbaProject.bin"))
	if err != nil {
		t.Fatal(err)
	}
	if len(p.Forms) != 1 {
		t.Fatalf("forms=%d", len(p.Forms))
	}
	projected, err := projection.Project(p.Forms[0])
	if err != nil {
		t.Fatal(err)
	}
	if projected.Form.Caption == nil || *projected.Form.Caption != "日本語😀" || len(projected.Controls) != 11 {
		t.Fatalf("projection=%+v", projected)
	}
	for _, module := range p.Modules {
		if module.Name == "JapaneseForm" && !strings.Contains(module.Source, "日本語") {
			t.Fatal("code lost Japanese text")
		}
	}
	// The caller-owned spec is still authoring input, not normalized output.
	if formSpec.Controls[0].Width != nil {
		t.Fatal("generation changed authoring input")
	}
}
