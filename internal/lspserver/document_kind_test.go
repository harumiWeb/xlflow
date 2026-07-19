package lspserver

import (
	"path/filepath"
	"testing"
)

func TestDetectDocumentKind(t *testing.T) {
	root := t.TempDir()
	specs := filepath.Join(root, "custom", "forms", "specs")
	tests := []struct {
		name   string
		path   string
		source string
		want   DocumentKind
	}{
		{"yaml", filepath.Join(specs, "Login.yaml"), "kind: xlflow.userform\n", DocumentKindUserFormYAML},
		{"yml", filepath.Join(specs, "Login.yml"), "kind: xlflow.userform\n", DocumentKindUserFormYAML},
		{"json", filepath.Join(specs, "Login.json"), `{"kind":"xlflow.userform"}`, DocumentKindUserFormJSON},
		{"incomplete yaml", filepath.Join(specs, "Login.yaml"), "controls:\n  - type: TextBox\n    ca\n", DocumentKindUserFormYAML},
		{"other yaml kind", filepath.Join(specs, "Other.yaml"), "kind: example.other\n", DocumentKindUnknown},
		{"other json kind", filepath.Join(specs, "Other.json"), `{"kind":"example.other"}`, DocumentKindUnknown},
		{"outside specs", filepath.Join(root, "notes.yaml"), "kind: xlflow.userform\n", DocumentKindUnknown},
		{"nested specs", filepath.Join(specs, "nested", "Login.yaml"), "kind: xlflow.userform\n", DocumentKindUnknown},
		{"vba", filepath.Join(root, "src", "modules", "Main.bas"), "Option Explicit\n", DocumentKindVBA},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			if got := DetectDocumentKind(root, filepath.Join("custom", "forms"), tc.path, tc.source); got != tc.want {
				t.Fatalf("DetectDocumentKind() = %v, want %v", got, tc.want)
			}
		})
	}
}
