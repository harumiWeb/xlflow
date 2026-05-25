package scripts

import (
	"os"
	"path/filepath"
	"testing"
)

func TestMaterializeUsesIndependentTempDirs(t *testing.T) {
	path1, cleanup1, err := Materialize("run")
	if err != nil {
		t.Fatal(err)
	}
	if cleanup1 == nil {
		t.Fatal("expected cleanup for first materialization")
	}
	cleaned1 := false
	defer func() {
		if !cleaned1 {
			cleanup1()
		}
	}()

	path2, cleanup2, err := Materialize("run")
	if err != nil {
		t.Fatal(err)
	}
	if cleanup2 == nil {
		t.Fatal("expected cleanup for second materialization")
	}
	defer cleanup2()

	dir1 := filepath.Dir(path1)
	dir2 := filepath.Dir(path2)
	if dir1 == dir2 {
		t.Fatalf("expected independent temp dirs, got shared dir %q", dir1)
	}

	cleanup1()
	cleaned1 = true

	if _, err := os.Stat(dir1); !os.IsNotExist(err) {
		t.Fatalf("expected first cleanup to remove %q, got %v", dir1, err)
	}
	if _, err := os.Stat(path2); err != nil {
		t.Fatalf("expected second script to remain after first cleanup: %v", err)
	}
	if _, err := os.Stat(filepath.Join(dir2, "common.ps1")); err != nil {
		t.Fatalf("expected second bundle common.ps1 to remain after first cleanup: %v", err)
	}
}

func TestMaterializeProcessScriptIsBundled(t *testing.T) {
	path, cleanup, err := Materialize("process")
	if err != nil {
		t.Fatal(err)
	}
	if cleanup == nil {
		t.Fatal("expected cleanup for materialized bundle")
	}
	defer cleanup()
	dir := filepath.Dir(path)
	if filepath.Base(path) != "process.ps1" {
		t.Fatalf("script path = %q, want bundled process.ps1", path)
	}
	if _, err := os.Stat(filepath.Join(dir, "common.ps1")); err != nil {
		t.Fatalf("expected bundled common.ps1 alongside process.ps1: %v", err)
	}
}
