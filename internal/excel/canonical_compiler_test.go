package excel

import (
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/pack"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestPushPackSourcesCarriesCanonicalSpecAndCodeSidecarMetadata(t *testing.T) {
	spec := &forms.FormSpec{Form: forms.FormSpecForm{Name: "CanonicalForm"}}
	rawSpec := []byte("form:\n  name: CanonicalForm\n")
	code := []byte("Private Sub Button_Click()\nEnd Sub\n")
	picture := []byte{0, 1, 2, 3}
	components := []sourceinventory.Component{
		{Type: sourceinventory.ComponentStandard, Name: "Module1", SourcePath: "src/modules/Module1.bas", Source: []byte("Option Explicit\n")},
		{Type: sourceinventory.ComponentClass, Name: "Class1", SourcePath: "src/classes/Class1.cls", Source: []byte("Option Explicit\n")},
		{
			Type:                 sourceinventory.ComponentForm,
			Name:                 "CanonicalForm",
			SourcePath:           "src/forms/specs/CanonicalForm.yaml",
			Source:               rawSpec,
			FormSpec:             spec,
			PreserveExistingCode: true,
			Related: []sourceinventory.Artifact{
				{Path: "src/forms/specs/CanonicalForm.yaml", Source: rawSpec},
				{Path: "src/forms/code/CanonicalForm.bas", Source: code},
				{Path: "src/forms/assets/CanonicalForm.bas", Source: picture, Role: sourceinventory.ArtifactRolePicture},
			},
		},
		{Type: sourceinventory.ComponentForm, Name: "LegacyForm", SourcePath: "src/forms/LegacyForm.frm", Source: []byte("legacy"), FormSpec: nil},
	}

	got := pushPackSources(components, "sidecar")
	if len(got) != 3 {
		t.Fatalf("source count = %d, want standard, class, canonical form", len(got))
	}
	if got[0].Type != pack.ModuleTypeStandard || got[1].Type != pack.ModuleTypeClass {
		t.Fatalf("baseline module types were not preserved: %+v", got[:2])
	}
	form := got[2]
	if form.Type != pack.ModuleTypeForm || form.Name != "CanonicalForm" {
		t.Fatalf("canonical form source = %+v", form)
	}
	if form.SourcePath != "src/forms/code/CanonicalForm.bas" || form.FormSpecPath != "src/forms/specs/CanonicalForm.yaml" {
		t.Fatalf("source paths = source %q, spec %q", form.SourcePath, form.FormSpecPath)
	}
	if !reflect.DeepEqual(form.FormSpecSource, rawSpec) || form.Source != string(code) {
		t.Fatalf("spec bytes/source were not retained: spec=%q source=%q", form.FormSpecSource, form.Source)
	}
	if !form.PreserveExistingCode || form.FormSpec != spec {
		t.Fatalf("preservation/spec metadata was lost: %+v", form)
	}
	if !reflect.DeepEqual(form.RelatedPaths, []string{"src/forms/specs/CanonicalForm.yaml", "src/forms/code/CanonicalForm.bas", "src/forms/assets/CanonicalForm.bas"}) {
		t.Fatalf("related paths = %v", form.RelatedPaths)
	}
}

func TestDiscoverPushCompilerInputsIgnoresStaleCanonicalCompatibilityArtifacts(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Src.Forms = filepath.Join("src", "forms")
	cfg.UserForm.CodeSource = "sidecar"
	write := func(relative string, body []byte) {
		t.Helper()
		path := filepath.Join(root, relative)
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, body, 0o600); err != nil {
			t.Fatal(err)
		}
	}
	write(filepath.Join("src", "forms", "specs", "Login.json"), []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[]}`))
	write(filepath.Join("src", "forms", "Login.frm"), []byte("stale malformed compatibility artifact"))
	write(filepath.Join("src", "forms", "Login.frx"), []byte("stale binary compatibility artifact"))

	components, names, err := discoverPushCompilerInputs(root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(names, []string{"Login"}) {
		t.Fatalf("canonical names = %v", names)
	}
	if len(components) != 1 || components[0].FormSpec == nil || !components[0].PreserveExistingCode {
		t.Fatalf("canonical component/preserve-code metadata = %+v", components)
	}
	for _, path := range components[0].RelatedPaths() {
		if filepath.Ext(path) == ".frm" || filepath.Ext(path) == ".frx" {
			t.Fatalf("stale compatibility artifact entered compiler inventory: %v", components[0].RelatedPaths())
		}
	}
}
