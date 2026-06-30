package typedb

import (
	"os"
	"path/filepath"
	"testing"
)

func TestStatusForMissingManifestReportsStale(t *testing.T) {
	dir := t.TempDir()

	status, err := StatusFor(Options{Dir: dir, GeneratorVersion: "1.2.3"})
	if err != nil {
		t.Fatal(err)
	}
	if status.ManifestExists {
		t.Fatal("manifest should be missing")
	}
	if !status.Stale || status.Reason != "manifest_missing" {
		t.Fatalf("status stale = %v reason = %q", status.Stale, status.Reason)
	}
	if status.Dir != dir {
		t.Fatalf("dir = %q, want %q", status.Dir, dir)
	}
}

func TestStatusForManifestChecksOutputFiles(t *testing.T) {
	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, "excel.generated.json"), []byte(`{"types":[]}`), 0o644); err != nil {
		t.Fatal(err)
	}
	manifest := Manifest{
		GeneratorVersion: "1.2.3",
		Libraries: []ManifestLibrary{{
			Name:   "Excel",
			LibID:  "{00020813-0000-0000-C000-000000000046}",
			Major:  1,
			Minor:  9,
			LCID:   0,
			Source: "registry",
			Output: "excel.generated.json",
		}},
	}
	if err := WriteManifest(dir, manifest); err != nil {
		t.Fatal(err)
	}

	status, err := StatusFor(Options{Dir: dir, GeneratorVersion: "1.2.3"})
	if err != nil {
		t.Fatal(err)
	}
	if !status.ManifestExists || status.Stale {
		t.Fatalf("unexpected status: %+v", status)
	}
	if len(status.Libraries) != 1 || !status.Libraries[0].Exists {
		t.Fatalf("library output not detected: %+v", status.Libraries)
	}
	if len(status.GeneratedFiles) != 1 {
		t.Fatalf("generated files = %+v", status.GeneratedFiles)
	}
}

func TestCleanRemovesResolvedDir(t *testing.T) {
	parent := t.TempDir()
	dir := filepath.Join(parent, "typelib")
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "manifest.json"), []byte(`{}`), 0o644); err != nil {
		t.Fatal(err)
	}

	cleaned, err := Clean(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cleaned != dir {
		t.Fatalf("cleaned = %q, want %q", cleaned, dir)
	}
	if _, err := os.Stat(dir); !os.IsNotExist(err) {
		t.Fatalf("dir still exists or unexpected error: %v", err)
	}
}
