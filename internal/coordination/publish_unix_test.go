//go:build !windows

package coordination

import (
	"os"
	"path/filepath"
	"syscall"
	"testing"
)

func TestPublishFileUnixPermissions(t *testing.T) {
	oldUmask := syscall.Umask(0o027)
	t.Cleanup(func() { syscall.Umask(oldUmask) })

	dir := t.TempDir()
	created := filepath.Join(dir, "created.xlsm")
	if _, err := PublishFile(created, []byte("new"), nil); err != nil {
		t.Fatalf("create: %v", err)
	}
	assertPermissionMode(t, created, 0o640)

	replaced := filepath.Join(dir, "replaced.xlsm")
	if err := os.WriteFile(replaced, []byte("old"), 0o604); err != nil {
		t.Fatal(err)
	}
	if err := os.Chmod(replaced, 0o604); err != nil {
		t.Fatal(err)
	}
	if _, err := PublishFile(replaced, []byte("replacement"), nil); err != nil {
		t.Fatalf("replace: %v", err)
	}
	assertPermissionMode(t, replaced, 0o604)
}

func TestPublishFileUnixReplacesReadOnlyTarget(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "Book.xlsm")
	if err := os.WriteFile(target, []byte("old"), 0o444); err != nil {
		t.Fatal(err)
	}
	if err := os.Chmod(target, 0o444); err != nil {
		t.Fatal(err)
	}

	if _, err := PublishFile(target, []byte("replacement"), nil); err != nil {
		t.Fatalf("PublishFile: %v", err)
	}
	body, err := os.ReadFile(target)
	if err != nil {
		t.Fatal(err)
	}
	if string(body) != "replacement" {
		t.Fatalf("published bytes = %q", body)
	}
	assertPermissionMode(t, target, 0o444)
}

func assertPermissionMode(t *testing.T, path string, want os.FileMode) {
	t.Helper()
	info, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if got := info.Mode().Perm(); got != want {
		t.Fatalf("mode = %04o, want %04o", got, want)
	}
}
