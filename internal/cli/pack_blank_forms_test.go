package cli

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
)

const blankFormSpec = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"text1","name":"TextBox1","type":"TextBox"}]}`

func TestBlankPackCanonicalFormsCodeAuthority(t *testing.T) {
	for _, tc := range []struct {
		name       string
		codeSource string
		withFRM    bool
		withCode   bool
		wantCode   string
		wantPath   string
	}{
		{name: "sidecar", codeSource: "sidecar", withFRM: true, withCode: true, wantCode: "FromSidecar", wantPath: "src/forms/code/Login.bas"},
		{name: "explicit frm", codeSource: "frm", withFRM: true, withCode: true, wantCode: "FromFRM", wantPath: "src/forms/Login.frm"},
		{name: "sidecar ignores compatibility frm", codeSource: "sidecar", withFRM: true},
		{name: "omitted code is empty", codeSource: "sidecar"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			root := t.TempDir()
			writePackSourceTree(t, root, false)
			writePackSourceModule(t, root, "src/forms/specs/Login.json", []byte(blankFormSpec))
			if tc.withFRM {
				writePackSourceModule(t, root, "src/forms/Login.frm", []byte("VERSION 5.00\nAttribute VB_Name = \"Login\"\nOption Explicit\nPublic Sub FromFRM()\nEnd Sub\n"))
			}
			if tc.withCode {
				writePackSourceModule(t, root, "src/forms/code/Login.bas", []byte("Option Explicit\nPublic Sub FromSidecar()\nEnd Sub\n"))
			}
			cfg := config.Default()
			cfg.UserForm.CodeSource = tc.codeSource
			sources, err := collectPackSourceModulesForMode(root, cfg, true)
			if err != nil {
				t.Fatal(err)
			}
			var form packpkg.SourceModule
			for _, s := range sources {
				if s.Type == packpkg.ModuleTypeForm {
					form = s
				}
			}
			if form.FormSpec == nil {
				t.Fatal("missing canonical spec")
			}
			if tc.wantCode != "" && !strings.Contains(form.Source, tc.wantCode) {
				t.Fatalf("code source = %q, want %q", form.Source, tc.wantCode)
			}
			if tc.wantCode == "" && form.Source != "" {
				t.Fatalf("omitted code should remain empty, got %q", form.Source)
			}
			if tc.wantPath != "" && form.SourcePath != tc.wantPath {
				t.Errorf("code source path = %q, want %q", form.SourcePath, tc.wantPath)
			}
			if _, meta, err := packpkg.BuildBlankWorkbook(sources, packpkg.BlankOptions{}); err != nil || meta.Form != 1 {
				t.Fatalf("blank generation: %+v %v", meta, err)
			}
		})
	}
}

func TestPackBlankFormsCLIAndPublicationSafety(t *testing.T) {
	root := t.TempDir()
	writePackConfig(t, root)
	writePackSourceTree(t, root, false)
	writePackSourceModule(t, root, "src/forms/specs/Login.json", []byte(blankFormSpec))
	out, err := runPackCommandForTest(root, "--json", "pack", "--blank", "--out", "dist/Fresh.xlsm")
	if err != nil {
		t.Fatalf("%v\n%s", err, out)
	}
	artifact := filepath.Join(root, "dist", "Fresh.xlsm")
	before, err := os.ReadFile(artifact)
	if err != nil {
		t.Fatal(err)
	}
	// Unsupported container must leave the existing output unchanged.
	writePackSourceModule(t, root, "src/forms/specs/Login.json", []byte(strings.Replace(blankFormSpec, `"type":"TextBox"`, `"type":"Frame","properties":{"picture":"unsupported"}`, 1)))
	out, err = runPackCommandForTest(root, "--json", "pack", "--blank", "--out", "dist/Fresh.xlsm")
	if err == nil || errorCodeFromJSON(t, out) != "pack_userform_generation_unsupported" {
		t.Fatalf("unsupported result: %v\n%s", err, out)
	}
	after, err := os.ReadFile(artifact)
	if err != nil || !bytes.Equal(before, after) {
		t.Fatal("rejected pack changed output")
	}
}
