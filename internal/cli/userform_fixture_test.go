package cli

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/excel"
)

func TestBuildFormWriteOptionsAcceptsRepresentativeUserFormFixture(t *testing.T) {
	body, err := os.ReadFile(filepath.Join("..", "excel", "forms", "intel", "testdata", "representative-userform.yaml"))
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	path := filepath.Join(root, "src", "forms", "specs", "AllControlsForm.yaml")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, body, 0o644); err != nil {
		t.Fatal(err)
	}
	opts, err := buildFormWriteOptions("build", path, true, true, true, excel.CommandOptions{}, root)
	if err != nil {
		t.Fatal(err)
	}
	if opts.Spec.Form.Name != "AllControlsForm" || len(opts.Spec.Controls) != 8 {
		t.Fatalf("fixture options = %#v", opts)
	}
}
