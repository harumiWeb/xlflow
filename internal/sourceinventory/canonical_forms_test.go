package sourceinventory

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

const canonicalTestSpec = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[]}`

func TestCanonicalFormDiscoveryRejectsAmbiguousInputs(t *testing.T) {
	for _, test := range []struct {
		name  string
		files map[string]string
		mode  string
	}{
		{"duplicate specs", map[string]string{"specs/Login.json": canonicalTestSpec, "specs/Login.yaml": canonicalTestSpec}, "sidecar"},
		{"wrong form name", map[string]string{"specs/Login.json": strings.Replace(canonicalTestSpec, "Login", "Other", 1)}, "sidecar"},
		{"orphan code", map[string]string{"code/Login.bas": "Option Explicit"}, "sidecar"},
		{"nested code", map[string]string{"specs/Login.json": canonicalTestSpec, "code/nested/Login.bas": "Option Explicit"}, "sidecar"},
		{"sidecar attributes", map[string]string{"specs/Login.json": canonicalTestSpec, "code/Login.bas": `Attribute VB_Name = "Login"`}, "sidecar"},
		{"frm required", map[string]string{"specs/Login.json": canonicalTestSpec}, "frm"},
		{"fake attribute in comment", map[string]string{"specs/Login.json": canonicalTestSpec, "Login.frm": `' Attribute VB_Name = "Login"`}, "frm"},
	} {
		t.Run(test.name, func(t *testing.T) {
			root := t.TempDir()
			cfg := config.Default()
			cfg.UserForm.CodeSource = test.mode
			for path, body := range test.files {
				writeInventoryFile(t, root, "src/forms", path, body)
			}
			if _, err := collectCanonicalForms(root, filepath.Join(root, "src/forms"), Options{Config: cfg}); err == nil {
				t.Fatal("accepted invalid canonical layout")
			}
		})
	}
}

func TestCanonicalFormsIgnoreUnusedDesignerAndAllowExternalRoots(t *testing.T) {
	root := t.TempDir()
	base := t.TempDir()
	cfg := config.Default()
	writeInventoryFile(t, base, "specs", "Login.json", canonicalTestSpec)
	writeInventoryFile(t, base, "code", "Login.bas", "Option Explicit")
	// The stale compatibility designer is deliberately not authoring input.
	writeInventoryFile(t, base, "", "Login.frm", "stale Designer bytes")
	writeInventoryFile(t, base, "", "Login.frx", "opaque resource")
	items, err := collectCanonicalForms(root, base, Options{Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || string(items[0].Source) != "Option Explicit" {
		t.Fatalf("items=%+v", items)
	}
	if _, err := os.Stat(filepath.Join(base, "Login.frm")); err != nil {
		t.Fatal(err)
	}
}

func TestCanonicalLegacyFormsRemainRecursiveAndAttachSidecars(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	frm := "VERSION 5.00\nAttribute VB_Name = \"Legacy\"\nOption Explicit\n"
	writeInventoryFile(t, root, "src/forms/legacy", "Legacy.frm", frm)
	writeInventoryFile(t, root, "src/forms/code", "Legacy.bas", "Option Explicit\nPublic Sub Run()\nEnd Sub\n")
	items, err := collectCanonicalForms(root, base, Options{Config: cfg, AllowLegacyFormArtifacts: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || items[0].FormSpec != nil || !strings.HasSuffix(items[0].SourcePath, "legacy/Legacy.frm") {
		t.Fatalf("items = %+v", items)
	}
	if _, ok := items[0].Artifact("src/forms/code/Legacy.bas"); !ok {
		t.Fatalf("legacy sidecar was not attached: %+v", items[0].RelatedPaths())
	}
}

func TestCanonicalLegacySidecarRejectsAmbiguousNestedForms(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	cfg := config.Default()
	frm := "VERSION 5.00\nAttribute VB_Name = \"Legacy\"\n"
	writeInventoryFile(t, root, "src/forms/one", "Legacy.frm", frm)
	writeInventoryFile(t, root, "src/forms/two", "Legacy.frm", frm)
	writeInventoryFile(t, root, "src/forms/code", "Legacy.bas", "Option Explicit\n")
	if _, err := collectCanonicalForms(root, base, Options{Config: cfg, AllowLegacyFormArtifacts: true}); err == nil {
		t.Fatal("accepted sidecar for ambiguous nested legacy forms")
	}
}

func TestCanonicalFormsRejectEmptyNestedDirectory(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(filepath.Join(base, "empty"), 0o755); err != nil {
		t.Fatal(err)
	}
	if _, err := collectCanonicalForms(root, base, Options{Config: config.Default()}); err == nil {
		t.Fatal("accepted empty nested canonical form directory")
	}
}

func TestCanonicalFormsReportFirstInvalidDirectoryDeterministically(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	for _, name := range []string{"Zulu", "Alpha"} {
		if err := os.MkdirAll(filepath.Join(base, name), 0o755); err != nil {
			t.Fatal(err)
		}
	}
	for range 64 {
		_, err := collectCanonicalForms(root, base, Options{Config: config.Default()})
		if err == nil || !strings.Contains(err.Error(), "src/forms/Alpha") {
			t.Fatalf("expected first directory Alpha, got %v", err)
		}
	}
}

func TestCanonicalFormsReportFirstAmbiguousLegacyFormDeterministically(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	for _, name := range []string{"Zulu", "Alpha"} {
		for _, dir := range []string{"one", "two"} {
			writeInventoryFile(t, root, "src/forms/"+dir, name+".frm", "VERSION 5.00\n")
		}
		writeInventoryFile(t, root, "src/forms/code", name+".bas", "Option Explicit\n")
	}
	for range 64 {
		_, err := collectCanonicalForms(root, base, Options{Config: config.Default(), AllowLegacyFormArtifacts: true})
		if err == nil || !strings.Contains(err.Error(), "ambiguous legacy form Alpha:") {
			t.Fatalf("expected first ambiguous form Alpha, got %v", err)
		}
	}
}

func TestCanonicalFormSpecPreservesAuthoredRootOmission(t *testing.T) {
	root := t.TempDir()
	base := filepath.Join(root, "src", "forms")
	body := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login","observed":{"caption":"Observed","width":300,"height":200}},"controls":[{"id":"label1","name":"Label1","type":"Label","caption":"Authored"}]}`
	writeInventoryFile(t, root, "src/forms/specs", "Login.json", body)
	cfg := config.Default()
	items, err := collectCanonicalForms(root, base, Options{Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 1 || items[0].FormSpec == nil {
		t.Fatalf("items = %+v", items)
	}
	spec := items[0].FormSpec
	if spec.Form.Caption != nil || spec.Form.Width != nil || spec.Form.Height != nil || spec.Form.Build != nil {
		t.Fatalf("inferred root fields survived: %+v", spec.Form)
	}
	if spec.Form.Observed == nil || spec.Form.Observed.Width == nil || *spec.Form.Observed.Width != 300 {
		t.Fatalf("explicit observed root fields were lost: %+v", spec.Form.Observed)
	}
	if len(spec.Controls) != 1 || spec.Controls[0].Caption == nil || *spec.Controls[0].Caption != "Authored" || spec.Controls[0].Observed != nil {
		t.Fatalf("control authored fields = %+v", spec.Controls)
	}
}
