//go:build windows

package workbookuse

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"testing"
	"unsafe"
)

func TestRestartManagerProcessInfoLayout(t *testing.T) {
	if size := unsafe.Sizeof(rmProcessInfo{}); size != 668 {
		t.Fatalf("rmProcessInfo size = %d, want 668", size)
	}
}

func TestDetectAlwaysEndsRestartManagerSession(t *testing.T) {
	previousCanonical := canonicalWorkbook
	previousStart := startSession
	previousRegister := registerWorkbook
	previousEnd := endSession
	previousQuery := queryProcesses
	t.Cleanup(func() {
		canonicalWorkbook = previousCanonical
		startSession = previousStart
		registerWorkbook = previousRegister
		endSession = previousEnd
		queryProcesses = previousQuery
	})
	canonicalWorkbook = func(path string) (string, error) { return path, nil }
	startSession = func() (uint32, error) { return 99, nil }
	ended := false
	endSession = func(handle uint32) {
		if handle != 99 {
			t.Fatalf("handle = %d", handle)
		}
		ended = true
	}
	registerWorkbook = func(uint32, string) error { return fmt.Errorf("register failed") }
	queryProcesses = func(uint32) ([]rmProcessInfo, error) {
		t.Fatal("query must not run after registration failure")
		return nil, nil
	}
	if _, err := Detect("Book.xlsm"); err == nil {
		t.Fatal("Detect succeeded after registration failure")
	}
	if !ended {
		t.Fatal("RmEndSession equivalent was not called")
	}
}

func TestDetectReportsOnlyExcelProcesses(t *testing.T) {
	previousCanonical := canonicalWorkbook
	previousStart := startSession
	previousRegister := registerWorkbook
	previousEnd := endSession
	previousQuery := queryProcesses
	previousIdentify := identifyExcel
	t.Cleanup(func() {
		canonicalWorkbook = previousCanonical
		startSession = previousStart
		registerWorkbook = previousRegister
		endSession = previousEnd
		queryProcesses = previousQuery
		identifyExcel = previousIdentify
	})
	canonicalWorkbook = func(path string) (string, error) { return path, nil }
	startSession = func() (uint32, error) { return 7, nil }
	registerWorkbook = func(uint32, string) error { return nil }
	endSession = func(uint32) {}
	queryProcesses = func(uint32) ([]rmProcessInfo, error) {
		return []rmProcessInfo{
			{Process: rmUniqueProcess{ProcessID: 101}},
			{Process: rmUniqueProcess{ProcessID: 202}},
			{Process: rmUniqueProcess{ProcessID: 202}},
		}, nil
	}
	identifyExcel = func(pid uint32) (bool, error) { return pid == 202, nil }

	state, err := Detect("Book.xlsm")
	if err != nil {
		t.Fatal(err)
	}
	if !state.OpenInExcel {
		t.Fatal("Excel process was not reported")
	}
	if len(state.ExcelPIDs) != 1 || state.ExcelPIDs[0] != 202 {
		t.Fatalf("ExcelPIDs = %v, want [202]", state.ExcelPIDs)
	}
}

func TestDetectIgnoresUnrelatedProcesses(t *testing.T) {
	previousCanonical := canonicalWorkbook
	previousStart := startSession
	previousRegister := registerWorkbook
	previousEnd := endSession
	previousQuery := queryProcesses
	previousIdentify := identifyExcel
	t.Cleanup(func() {
		canonicalWorkbook = previousCanonical
		startSession = previousStart
		registerWorkbook = previousRegister
		endSession = previousEnd
		queryProcesses = previousQuery
		identifyExcel = previousIdentify
	})
	canonicalWorkbook = func(path string) (string, error) { return path, nil }
	startSession = func() (uint32, error) { return 7, nil }
	registerWorkbook = func(uint32, string) error { return nil }
	endSession = func(uint32) {}
	queryProcesses = func(uint32) ([]rmProcessInfo, error) {
		return []rmProcessInfo{{Process: rmUniqueProcess{ProcessID: 101}}}, nil
	}
	identifyExcel = func(uint32) (bool, error) { return false, nil }

	state, err := Detect("Book.xlsm")
	if err != nil {
		t.Fatal(err)
	}
	if state.OpenInExcel || len(state.ExcelPIDs) != 0 {
		t.Fatalf("unrelated process produced state %+v", state)
	}
}

func TestDetectClosedUnicodeFile(t *testing.T) {
	path := filepath.Join(t.TempDir(), "日本語 workbook.xlsm")
	if err := os.WriteFile(path, []byte("fixture"), 0o600); err != nil {
		t.Fatal(err)
	}
	state, err := Detect(path)
	if err != nil {
		t.Fatal(err)
	}
	if state.OpenInExcel || len(state.ExcelPIDs) != 0 {
		t.Fatalf("closed workbook state = %+v", state)
	}
}

func TestDetectMissingFile(t *testing.T) {
	_, err := Detect(filepath.Join(t.TempDir(), "missing.xlsm"))
	if !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("error = %v, want os.ErrNotExist", err)
	}
}
