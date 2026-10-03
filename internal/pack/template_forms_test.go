package pack

import (
	"bytes"
	"errors"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func authoredForm(name string) SourceModule {
	return SourceModule{Name: name, Type: ModuleTypeForm, Source: "Option Explicit\n' updated form code", FormSpec: &spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: spec.FormSpecForm{Name: name, Caption: new("日本語")}, Controls: []spec.FormSpecControl{{ID: "custom", Name: "Input", Type: "TextBox", Text: new("first")}}}}
}

func TestTemplateFormsAddEditRemoveAndPreserve(t *testing.T) {
	base := readTestFile(t, "corpus", "p1_compiled.bin")
	created, err := GenerateVBAProject(base, []SourceModule{authoredForm("Login")})
	if err != nil {
		t.Fatal(err)
	}
	p, err := vbaproject.Read(created)
	if err != nil {
		t.Fatal(err)
	}
	before, err := oforms.SerializeForm(p.Forms[0], p.Props.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	kept, err := GenerateVBAProject(created, nil)
	if err != nil {
		t.Fatal(err)
	}
	readKept, err := vbaproject.Read(kept)
	if err != nil {
		t.Fatal(err)
	}
	after, err := oforms.SerializeForm(readKept.Forms[0], readKept.Props.CodePage)
	if err != nil || !reflect.DeepEqual(before, after) {
		t.Fatalf("Designer drift: %v", err)
	}
	update := authoredForm("Login")
	update.FormSpec.Form.Caption = new("Changed")
	update.FormSpec.Controls[0].Text = new("second")
	updated, err := GenerateVBAProject(created, []SourceModule{update})
	if err != nil {
		t.Fatal(err)
	}
	readUpdated, err := vbaproject.Read(updated)
	if err != nil {
		t.Fatal(err)
	}
	state, err := projection.Project(readUpdated.Forms[0])
	if err != nil || *state.Form.Caption != "Changed" || *state.Controls[0].Text != "second" {
		t.Fatalf("update failed: %+v %v", state, err)
	}
	if !strings.Contains(projectModuleSource(readUpdated, "Login"), "updated form code") {
		t.Fatal("code not updated")
	}
	removed, err := GenerateVBAProject(updated, nil, TemplateOptions{UserFormTopology: "source"})
	if err != nil {
		t.Fatal(err)
	}
	readRemoved, err := vbaproject.Read(removed)
	if err != nil {
		t.Fatal(err)
	}
	if len(readRemoved.Forms) != 0 || projectModuleSource(readRemoved, "Login") != "" || strings.Contains(string(readRemoved.ProjectStreamRaw), "BaseClass=Login") {
		t.Fatal("form not removed completely")
	}
	if !bytes.Equal(readUpdated.ReferencesRaw, readRemoved.ReferencesRaw) {
		t.Fatal("removal changed references")
	}
}

func TestTemplatePlanFormsIsAtomicAndDeterministic(t *testing.T) {
	p, err := vbaproject.Read(readTestFile(t, "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	before, err := vbaproject.Clone(p)
	if err != nil {
		t.Fatal(err)
	}
	bad := authoredForm("Bad")
	bad.FormSpec.Controls[0].Type = "Frame"
	bad.FormSpec.Controls[0].Properties = map[string]any{"picture": "unsupported"}
	plan, err := PlanProject(p, []SourceModule{authoredForm("Good"), bad})
	if detail, ok := errors.AsType[*compiler.Error](err); !ok || detail.Code != compiler.GenerationUnsupported {
		t.Fatalf("capability error: %v", err)
	}
	if plan.project != nil || !reflect.DeepEqual(before, p) {
		t.Fatal("failed plan changed input")
	}
	plan, err = PlanProject(p, []SourceModule{authoredForm("Zed"), authoredForm("Alpha")})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(before, p) {
		t.Fatal("successful plan changed input")
	}
	var names []string
	for _, m := range plan.modules {
		if m.Type == vbaproject.ModuleForm {
			names = append(names, m.Name)
		}
	}
	if !reflect.DeepEqual(names, []string{"Alpha", "Zed"}) {
		t.Fatalf("form order: %v", names)
	}
}

func TestReadbackRejectsOrphanDesigner(t *testing.T) {
	body, err := GenerateVBAProject(readTestFile(t, "corpus", "p1_compiled.bin"), []SourceModule{authoredForm("Login")})
	if err != nil {
		t.Fatal(err)
	}
	p, err := vbaproject.Read(body)
	if err != nil {
		t.Fatal(err)
	}
	p.Modules[len(p.Modules)-1].Type = vbaproject.ModuleClass
	bad, err := vbaproject.Write(p)
	if err == nil && validateProjectReadback(p, bad) == nil {
		t.Fatal("orphan Designer accepted")
	}
}

func TestReadbackRejectsExtraOpaqueStreamAndStorage(t *testing.T) {
	body, err := GenerateVBAProject(readTestFile(t, "corpus", "p1_compiled.bin"), nil)
	if err != nil {
		t.Fatal(err)
	}
	expected, err := vbaproject.Read(body)
	if err != nil {
		t.Fatal(err)
	}
	for _, scenario := range []string{"extra stream", "extra empty storage", "changed metadata"} {
		t.Run(scenario, func(t *testing.T) {
			changed, err := vbaproject.Clone(expected)
			if err != nil {
				t.Fatal(err)
			}
			switch scenario {
			case "extra stream":
				changed.RawStreams["OpaqueExtra"] = []byte("unexpected")
			case "extra empty storage":
				changed.StorageMetadata["OpaqueEmpty"] = cfb.StorageMeta{}
			case "changed metadata":
				meta := changed.StorageMetadata[""]
				meta.CLSID[0] ^= 1
				changed.StorageMetadata[""] = meta
			}
			bad, err := vbaproject.Write(changed)
			if err != nil {
				t.Fatal(err)
			}
			if validateProjectReadback(expected, bad) == nil {
				t.Fatal("unplanned opaque state accepted")
			}
		})
	}
}

func TestTemplateNoOpDesignerStillCountsCarriedStreams(t *testing.T) {
	body, err := GenerateVBAProject(readTestFile(t, "corpus", "p1_compiled.bin"), []SourceModule{authoredForm("Login")})
	if err != nil {
		t.Fatal(err)
	}
	_, baseline, err := generateVBAProject(body, nil)
	if err != nil {
		t.Fatal(err)
	}
	_, noOp, err := generateVBAProject(body, []SourceModule{authoredForm("Login")})
	if err != nil {
		t.Fatal(err)
	}
	if noOp.CarriedStreams != baseline.CarriedStreams {
		t.Fatalf("no-op undercount: %d vs %d", noOp.CarriedStreams, baseline.CarriedStreams)
	}
}
