package coordination

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestPublishFileCreatesNewArtifact(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "Book.xlsm")

	result, err := PublishFile(target, []byte("packed-workbook"), nil)
	if err != nil {
		t.Fatalf("PublishFile: %v", err)
	}
	if result.ReplacedExisting || result.Publication != "atomic_create" {
		t.Fatalf("publication = %#v, want atomic_create without replace", result)
	}
	if result.Cleanup.Status != "clean" {
		t.Fatalf("cleanup = %#v, want clean", result.Cleanup)
	}
	body, err := os.ReadFile(target)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(body, []byte("packed-workbook")) {
		t.Fatalf("published bytes = %q", body)
	}
	assertNoStagingResidual(t, dir)
}

func TestPublishFileReplacesExistingArtifact(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "Book.xlsm")
	if err := os.WriteFile(target, []byte("previous-valid-output"), 0o644); err != nil {
		t.Fatal(err)
	}

	result, err := PublishFile(target, []byte("packed-workbook-v2"), nil)
	if err != nil {
		t.Fatalf("PublishFile: %v", err)
	}
	if !result.ReplacedExisting || result.Publication != "atomic_replace" {
		t.Fatalf("publication = %#v, want atomic_replace", result)
	}
	body, err := os.ReadFile(target)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(body, []byte("packed-workbook-v2")) {
		t.Fatalf("published bytes = %q", body)
	}
	assertNoStagingResidual(t, dir)
}

func TestPublishFileValidationFailurePreservesOutput(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "Book.xlsm")
	sentinel := []byte("previous-valid-output")
	if err := os.WriteFile(target, sentinel, 0o644); err != nil {
		t.Fatal(err)
	}

	_, err := PublishFile(target, []byte("replacement"), func(string) error {
		return errors.New("staged artifact rejected")
	})
	if err == nil || !strings.Contains(err.Error(), "staged artifact rejected") {
		t.Fatalf("err = %v, want validation failure", err)
	}
	body, readErr := os.ReadFile(target)
	if readErr != nil {
		t.Fatal(readErr)
	}
	if !bytes.Equal(body, sentinel) {
		t.Fatalf("existing output was modified: %q", body)
	}
	assertNoStagingResidual(t, dir)
}

func TestPublishFileValidationFailureLeavesMissingOutputAbsent(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "Book.xlsm")

	_, err := PublishFile(target, []byte("replacement"), func(string) error {
		return errors.New("staged artifact rejected")
	})
	if err == nil {
		t.Fatal("expected validation failure")
	}
	if _, statErr := os.Stat(target); !errors.Is(statErr, os.ErrNotExist) {
		t.Fatalf("output should remain absent, stat error = %v", statErr)
	}
	assertNoStagingResidual(t, dir)
}

func TestPublishFileReplaceIntoMissingParentFailsWithoutResidual(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "missing", "Book.xlsm")

	_, err := PublishFile(target, []byte("packed-workbook"), nil)
	if err == nil {
		t.Fatal("expected failure for a missing destination directory")
	}
	assertNoStagingResidual(t, filepath.Join(dir, "missing"))
}

func TestSameFileIdentityResolvesSymlinkAlias(t *testing.T) {
	dir := t.TempDir()
	build := filepath.Join(dir, "build")
	if err := os.MkdirAll(build, 0o755); err != nil {
		t.Fatal(err)
	}
	target := filepath.Join(build, "Book.xlsm")
	if err := os.WriteFile(target, []byte("template"), 0o644); err != nil {
		t.Fatal(err)
	}
	link := filepath.Join(dir, "build-alias")
	if err := os.Symlink(build, link); err != nil {
		t.Skipf("symlink creation unavailable: %v", err)
	}
	defer func() { _ = os.Remove(link) }()

	same, err := SameFileIdentity(dir, filepath.Join(link, "Book.xlsm"), target)
	if err != nil {
		t.Fatalf("SameFileIdentity: %v", err)
	}
	if !same {
		t.Fatal("symlinked directory alias was not recognized as the same file")
	}
	other, err := SameFileIdentity(dir, filepath.Join(link, "Other.xlsm"), target)
	if err != nil {
		t.Fatalf("SameFileIdentity: %v", err)
	}
	if other {
		t.Fatal("distinct filenames must not share an identity")
	}
}

func TestSameFileIdentityRejectsMissingDistinctFiles(t *testing.T) {
	dir := t.TempDir()
	same, err := SameFileIdentity(dir,
		filepath.Join(dir, "dist", "Book.xlsm"),
		filepath.Join(dir, "build", "Book.xlsm"))
	if err != nil {
		t.Fatalf("SameFileIdentity: %v", err)
	}
	if same {
		t.Fatal("distinct missing paths must not share an identity")
	}
}

func assertNoStagingResidual(t *testing.T, dir string) {
	t.Helper()
	entries, err := os.ReadDir(dir)
	if err != nil {
		if !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("read staging dir: %v", err)
		}
		return
	}
	for _, entry := range entries {
		if strings.HasPrefix(entry.Name(), ".xlflow-publish-") {
			t.Fatalf("temporary artifact was not cleaned: %s", entry.Name())
		}
	}
}
