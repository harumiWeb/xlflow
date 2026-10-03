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
