package symbols

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestProjectDiscoveryDeduplicatesConfiguredTestSources(t *testing.T) {
	root := t.TempDir()
	tests := filepath.Join(root, "tests")
	if err := os.MkdirAll(tests, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(tests, "Checks.bas")
	if err := os.WriteFile(path, []byte("Public Sub TestPass()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Src.Modules = "tests"
	files, err := DiscoverProjectSourceFilesContext(t.Context(), root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 1 || files[0].Path != path || files[0].ModuleKind != "standard" {
		t.Fatalf("deduplicated discovery: %+v", files)
	}
}

func TestProjectDiscoveryHonorsCancellation(t *testing.T) {
	ctx, cancel := context.WithCancel(t.Context())
	cancel()
	if _, err := DiscoverProjectSourceFilesContext(ctx, t.TempDir(), config.Default()); !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled discovery: %v", err)
	}
}

func TestProjectDiscoverySelectsOneTestUserFormCodeSource(t *testing.T) {
	for _, mode := range []string{"frm", "sidecar"} {
		t.Run(mode, func(t *testing.T) {
			root := t.TempDir()
			forms := filepath.Join(root, "tests", "forms")
			if err := os.MkdirAll(filepath.Join(forms, "code"), 0o755); err != nil {
				t.Fatal(err)
			}
			form := filepath.Join(forms, "Panel.frm")
			code := filepath.Join(forms, "code", "Panel.bas")
			for _, path := range []string{form, code} {
				if err := os.WriteFile(path, []byte("Attribute VB_Name = \"Panel\"\n"), 0o644); err != nil {
					t.Fatal(err)
				}
			}
			cfg := config.Default()
			cfg.UserForm.CodeSource = mode
			files, err := DiscoverProjectSourceFilesContext(t.Context(), root, cfg)
			if err != nil {
				t.Fatal(err)
			}
			want := form
			if mode == "sidecar" {
				want = code
			}
			if len(files) != 1 || files[0].Path != want || files[0].ModuleKind != "form" {
				t.Fatalf("%s selection: %+v", mode, files)
			}
		})
	}
}
