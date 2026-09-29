//go:build windows

package coordination

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"testing"

	"golang.org/x/sys/windows"
)

func TestPublishFileReportsBusyDestination(t *testing.T) {
	dir := t.TempDir()
	target := filepath.Join(dir, "Book.xlsm")
	sentinel := []byte("previous-valid-output")
	if err := os.WriteFile(target, sentinel, 0o644); err != nil {
		t.Fatal(err)
	}

	handle := holdExclusiveHandle(t, target)

	_, err := PublishFile(target, []byte("replacement"), nil)
	if !errors.Is(err, ErrPublishTargetBusy) {
		t.Fatalf("err = %v, want ErrPublishTargetBusy", err)
	}
	// The exclusive handle denies reads as well, so release it before
	// verifying the destination was preserved.
	if err := windows.CloseHandle(handle); err != nil {
		t.Fatalf("CloseHandle: %v", err)
	}
	body, readErr := os.ReadFile(target)
	if readErr != nil {
		t.Fatal(readErr)
	}
	if !bytes.Equal(body, sentinel) {
		t.Fatalf("busy output was modified: %q", body)
	}
	assertNoStagingResidual(t, dir)
}

func holdExclusiveHandle(t *testing.T, path string) windows.Handle {
	t.Helper()
	name, err := windows.UTF16PtrFromString(path)
	if err != nil {
		t.Fatal(err)
	}
	// share mode 0 denies every sharing mode, including the delete sharing that
	// MoveFileEx requires to replace the destination.
	handle, err := windows.CreateFile(name, windows.GENERIC_READ|windows.GENERIC_WRITE, 0, nil, windows.OPEN_EXISTING, windows.FILE_ATTRIBUTE_NORMAL, 0)
	if err != nil {
		t.Fatalf("CreateFile exclusive: %v", err)
	}
	return handle
}
