package pack

import (
	"bytes"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestPackFramesGenerateAndEditHierarchy(t *testing.T) {
	source := authoredForm("NestedForm")
	source.FormSpec.Controls = []spec.FormSpecControl{
		{ID: "left", Name: "LeftFrame", Type: "Frame", Controls: []spec.FormSpecControl{{ID: "input", Name: "Input", Type: "TextBox", Text: new("preserve")}}},
		{ID: "right", Name: "RightFrame", Type: "Frame"},
	}
	created, err := GenerateVBAProject(readTestFile(t, "corpus", "p1_compiled.bin"), []SourceModule{source})
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(created)
	if err != nil {
		t.Fatal(err)
	}
	before := bytes.Clone(project.Forms[0].Controls[0].Children[0].Record.Raw)
	source.FormSpec.Controls = []spec.FormSpecControl{
		{ID: "right", Name: "RightFrame", Type: "Frame", ZIndex: new(0)},
		{ID: "input", Name: "Input", Type: "TextBox", ParentID: "right"},
		{ID: "new", Name: "NewFrame", Type: "Frame", ZIndex: new(1), Controls: []spec.FormSpecControl{{Name: "NewLabel", Type: "Label", Caption: new("nested")}}},
	}
	updated, err := GenerateVBAProject(created, []SourceModule{source})
	if err != nil {
		t.Fatal(err)
	}
	got, err := vbaproject.Read(updated)
	if err != nil {
		t.Fatal(err)
	}
	form := got.Forms[0]
	if len(form.Controls) != 2 || form.Controls[0].Name != "RightFrame" || form.Controls[0].Children[0].Name != "Input" || !bytes.Equal(before, form.Controls[0].Children[0].Record.Raw) {
		t.Fatal("retained payload or hierarchy changed")
	}
	state, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if len(state.Controls) != 4 || state.Controls[1].ParentID != state.Controls[0].ID || state.Controls[3].ParentID != state.Controls[2].ID {
		t.Fatal("pull projection hierarchy")
	}
	stored, err := oforms.SerializeForm(form, got.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	if _, orphan := stored.Storages["NestedForm/i01"]; orphan {
		t.Fatal("deleted Frame orphan")
	}
	if got.Modules[len(got.Modules)-1].Type != vbaproject.ModuleForm {
		t.Fatal("module identity changed")
	}
}
