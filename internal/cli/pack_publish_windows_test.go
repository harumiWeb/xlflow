//go:build windows

package cli

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"golang.org/x/sys/windows"

	"github.com/harumiWeb/xlflow/internal/output"
)

func TestPackCommandReportsBusyOutput(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)

	if err := os.MkdirAll(filepath.Join(dir, "dist"), 0o755); err != nil {
		t.Fatal(err)
	}
	outPath := filepath.Join(dir, "dist", "Book.xlsm")
	sentinel := []byte("previous-valid-output")
	if err := os.WriteFile(outPath, sentinel, 0o644); err != nil {
		t.Fatal(err)
	}

	// Deny every sharing mode, including the delete sharing that MoveFileEx
	// requires to replace the destination.
	name, err := windows.UTF16PtrFromString(outPath)
	if err != nil {
		t.Fatal(err)
	}
	handle, err := windows.CreateFile(name, windows.GENERIC_READ|windows.GENERIC_WRITE, 0, nil, windows.OPEN_EXISTING, windows.FILE_ATTRIBUTE_NORMAL, 0)
	if err != nil {
		t.Fatalf("CreateFile exclusive: %v", err)
	}

	stdout, runErr := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if runErr == nil || output.ExitCode(runErr) != output.ExitEnvironment {
		t.Fatalf("err=%v exit=%d, want environment failure", runErr, output.ExitCode(runErr))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_output_busy" {
		t.Fatalf("error code = %q, want pack_output_busy\n%s", got, stdout)
	}
	// The exclusive handle denies reads as well, so release it before
	// verifying the destination was preserved.
	if err := windows.CloseHandle(handle); err != nil {
		t.Fatalf("CloseHandle: %v", err)
	}
	body, readErr := os.ReadFile(outPath)
	if readErr != nil {
		t.Fatal(readErr)
	}
	if !bytes.Equal(body, sentinel) {
		t.Fatalf("busy output was modified: %q", body)
	}
	assertNoPackStagingResidual(t, filepath.Join(dir, "dist"))
}
