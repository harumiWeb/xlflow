package sourceedit

import (
	"bytes"
	"context"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestCaptureClonesPublicSourceAndApplyPreservesMode(t *testing.T) {
	path := writeTestFile(t, t.TempDir(), "source.bas", []byte("old source\n"), 0o640)
	before, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	snapshot := captureTestSnapshot(t, path)
	if !filepath.IsAbs(snapshot.Path) {
		t.Fatalf("Snapshot.Path = %q, want absolute path", snapshot.Path)
	}
	if !bytes.Equal(snapshot.Source, []byte("old source\n")) {
		t.Fatalf("Snapshot.Source = %q", snapshot.Source)
	}
	snapshot.Source[0] = 'X'

	if err := snapshot.Apply(t.Context(), []byte("new source\n")); err != nil {
		t.Fatalf("Apply() error = %v", err)
	}
	got, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(got) != "new source\n" {
		t.Fatalf("published contents = %q", got)
	}
	info, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if info.Mode().Perm() != before.Mode().Perm() {
		t.Fatalf("published mode = %04o, want captured mode %04o", info.Mode().Perm(), before.Mode().Perm())
	}
}

func TestApplyRejectsStaleContent(t *testing.T) {
	dir := t.TempDir()
	path := writeTestFile(t, dir, "source.bas", []byte("captured\n"), 0o600)
	snapshot := captureTestSnapshot(t, path)
	if err := os.WriteFile(path, []byte("external edit\n"), 0o600); err != nil {
		t.Fatal(err)
	}

	err := snapshot.Apply(t.Context(), []byte("replacement\n"))
	if !errors.Is(err, ErrStale) {
		t.Fatalf("Apply() error = %v, want ErrStale", err)
	}
	assertFileContents(t, path, []byte("external edit\n"))
	assertNoTemporaryFiles(t, dir)
}

func TestApplyRejectsReplacedFileIdentity(t *testing.T) {
	dir := t.TempDir()
	path := writeTestFile(t, dir, "source.bas", []byte("captured\n"), 0o600)
	snapshot := captureTestSnapshot(t, path)
	backup := filepath.Join(dir, "moved-away.bas")
	if err := os.Rename(path, backup); err != nil {
		t.Fatal(err)
	}
	writeTestFile(t, dir, filepath.Base(path), []byte("replacement identity\n"), 0o600)

	err := snapshot.Apply(t.Context(), []byte("new contents\n"))
	if !errors.Is(err, ErrStale) {
		t.Fatalf("Apply() error = %v, want ErrStale", err)
	}
	assertFileContents(t, path, []byte("replacement identity\n"))
	assertNoTemporaryFiles(t, dir)
}

func TestApplyRejectsChangedPermissionBits(t *testing.T) {
	path := writeTestFile(t, t.TempDir(), "source.bas", []byte("captured\n"), 0o600)
	snapshot := captureTestSnapshot(t, path)
	if err := os.Chmod(path, 0o400); err != nil {
		t.Fatal(err)
	}

	err := snapshot.Apply(t.Context(), []byte("new contents\n"))
	if !errors.Is(err, ErrStale) {
		t.Fatalf("Apply() error = %v, want ErrStale", err)
	}
	assertFileContents(t, path, []byte("captured\n"))
}

func TestCaptureRejectsSymlinkTarget(t *testing.T) {
	dir := t.TempDir()
	target := writeTestFile(t, dir, "target.bas", []byte("source"), 0o600)
	link := filepath.Join(dir, "link.bas")
	if err := os.Symlink(target, link); err != nil {
		t.Skipf("symlink creation unavailable: %v", err)
	}
	if _, err := Capture(t.Context(), link); err == nil {
		t.Fatal("Capture() succeeded for a symlink target")
	}
}

func TestCaptureRejectsSymlinkAncestor(t *testing.T) {
	root := t.TempDir()
	targetDir := filepath.Join(root, "target-dir")
	if err := os.Mkdir(targetDir, 0o700); err != nil {
		t.Fatal(err)
	}
	writeTestFile(t, targetDir, "source.bas", []byte("source"), 0o600)
	linkDir := filepath.Join(root, "link-dir")
	if err := os.Symlink(targetDir, linkDir); err != nil {
		t.Skipf("directory symlink creation unavailable: %v", err)
	}
	if _, err := Capture(t.Context(), filepath.Join(linkDir, "source.bas")); err == nil {
		t.Fatal("Capture() succeeded through a symlink ancestor")
	}
}

func TestApplyRejectsReadOnlyOriginal(t *testing.T) {
	path := writeTestFile(t, t.TempDir(), "source.bas", []byte("source"), 0o600)
	if err := os.Chmod(path, 0o444); err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = os.Chmod(path, 0o600) })
	info, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if info.Mode().Perm()&0o222 != 0 {
		t.Skipf("filesystem does not expose read-only permission bits: %04o", info.Mode().Perm())
	}
	snapshot := captureTestSnapshot(t, path)

	err = snapshot.Apply(t.Context(), []byte("replacement"))
	if !errors.Is(err, ErrReadOnly) {
		t.Fatalf("Apply() error = %v, want ErrReadOnly", err)
	}
	assertFileContents(t, path, []byte("source"))
}

func TestApplyCancellationBeforePublicationCleansTemporaryFile(t *testing.T) {
	dir := t.TempDir()
	path := writeTestFile(t, dir, "source.bas", []byte("source"), 0o600)
	snapshot := captureTestSnapshot(t, path)
	ctx, cancel := context.WithCancel(t.Context())
	snapshot.ops.beforePublish = cancel

	err := snapshot.Apply(ctx, []byte("replacement"))
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("Apply() error = %v, want context.Canceled", err)
	}
	assertFileContents(t, path, []byte("source"))
	assertNoTemporaryFiles(t, dir)
}

func TestApplyNoOpDoesNotReplaceFile(t *testing.T) {
	dir := t.TempDir()
	path := writeTestFile(t, dir, "source.bas", []byte("same source"), 0o600)
	snapshot := captureTestSnapshot(t, path)
	before, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	snapshot.ops.publish = func(string, string) error {
		t.Fatal("no-op called publish")
		return nil
	}

	if err := snapshot.Apply(t.Context(), []byte("same source")); err != nil {
		t.Fatalf("Apply() error = %v", err)
	}
	after, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if !os.SameFile(before, after) {
		t.Fatal("no-op replaced the source file")
	}
	assertNoTemporaryFiles(t, dir)
}

func TestApplyPublishFailureLeavesOriginalAndCleansTemporaryFile(t *testing.T) {
	dir := t.TempDir()
	path := writeTestFile(t, dir, "source.bas", []byte("source"), 0o600)
	snapshot := captureTestSnapshot(t, path)
	wantErr := errors.New("injected publish failure")
	snapshot.ops.publish = func(tempPath, targetPath string) error {
		if filepath.Dir(tempPath) != filepath.Dir(targetPath) {
			t.Errorf("temporary file directory = %q, target directory = %q", filepath.Dir(tempPath), filepath.Dir(targetPath))
		}
		return wantErr
	}

	err := snapshot.Apply(t.Context(), []byte("replacement"))
	if !errors.Is(err, wantErr) {
		t.Fatalf("Apply() error = %v, want injected error", err)
	}
	assertFileContents(t, path, []byte("source"))
	assertNoTemporaryFiles(t, dir)
}

func TestApplyRechecksAfterStaging(t *testing.T) {
	dir := t.TempDir()
	path := writeTestFile(t, dir, "source.bas", []byte("captured"), 0o600)
	snapshot := captureTestSnapshot(t, path)
	snapshot.ops.beforePublish = func() {
		if err := os.WriteFile(path, []byte("external edit after staging"), 0o600); err != nil {
			t.Fatal(err)
		}
	}
	if err := snapshot.Apply(t.Context(), []byte("proposed")); !errors.Is(err, ErrStale) {
		t.Fatalf("Apply() error = %v, want ErrStale", err)
	}
	assertFileContents(t, path, []byte("external edit after staging"))
	assertNoTemporaryFiles(t, dir)
}

func captureTestSnapshot(t *testing.T, path string) *Snapshot {
	t.Helper()
	snapshot, err := Capture(t.Context(), path)
	if err != nil {
		t.Fatalf("Capture() error = %v", err)
	}
	return snapshot
}

func writeTestFile(t *testing.T, dir, name string, contents []byte, mode os.FileMode) string {
	t.Helper()
	path := filepath.Join(dir, name)
	if err := os.WriteFile(path, contents, mode); err != nil {
		t.Fatal(err)
	}
	if err := os.Chmod(path, mode); err != nil {
		t.Fatal(err)
	}
	return path
}

func assertFileContents(t *testing.T, path string, want []byte) {
	t.Helper()
	got, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, want) {
		t.Fatalf("file contents = %q, want %q", got, want)
	}
}

func assertNoTemporaryFiles(t *testing.T, dir string) {
	t.Helper()
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	for _, entry := range entries {
		if strings.HasPrefix(entry.Name(), ".sourceedit-") {
			t.Errorf("temporary file remains: %s", entry.Name())
		}
	}
}
