package cli

import (
	"bytes"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/harumiWeb/xlflow/internal/filepush"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/workbookuse"
)

func stubClosedWorkbook(t *testing.T) {
	t.Helper()
	if runtime.GOOS != "windows" {
		return
	}
	t.Cleanup(stubWorkbookUseDetector(t, workbookuse.State{}))
}

func writePushSourceTree(t *testing.T, dir string) {
	t.Helper()
	writePackSourceModule(t, dir, filepath.Join("src", "modules", "Module1.bas"), readPackFixture(t, "testdata", "disk", "p1", "modules", "Module1.bas"))
	writePackSourceModule(t, dir, filepath.Join("src", "classes", "Class1.cls"), readPackFixture(t, "testdata", "disk", "p1", "classes", "Class1.cls"))
	writePackSourceModule(t, dir, filepath.Join("src", "workbook", "Sheet1.bas"), readPackFixture(t, "testdata", "disk", "p1", "workbook", "Sheet1.bas"))
	writePackSourceModule(t, dir, filepath.Join("src", "workbook", "ThisWorkbook.bas"), readPackFixture(t, "testdata", "disk", "p1", "workbook", "ThisWorkbook.bas"))
}

func TestPushDefaultsToExcelBackend(t *testing.T) {
	// The default must stay the Excel bridge; the file backend is opt-in only.
	// Assert the flag contract directly instead of executing a push that would
	// need Excel.
	a := &app{cwd: t.TempDir()}
	cmd := a.pushCommand()
	flag := cmd.Flags().Lookup("backend")
	if flag == nil || flag.DefValue != "excel" {
		t.Fatalf("push --backend default = %+v, want excel", flag)
	}
}

func TestPushRejectsInvalidBackendAndSessionCombos(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	for _, args := range [][]string{
		{"push", "--backend", "bogus"},
		{"push", "--backend", "auto"},
		{"push", "--backend", "file", "--session"},
		{"push", "--backend", "file", "--no-save"},
	} {
		stdout, err := runBuildCommandForTest(dir, append([]string{"--json"}, args...)...)
		if err == nil || output.ExitCode(err) != output.ExitConfig || jsonErrorCode(t, stdout) != "push_args_invalid" {
			t.Fatalf("%v: stdout=%s err=%v", args, stdout, err)
		}
	}
}

func TestPushFileBackendPublishesSavedWorkbook(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePushSourceTree(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	stubClosedWorkbook(t)

	stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file")
	if err != nil {
		t.Fatalf("push --backend file: %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	push := cliObjectMap(env.Push)
	if push["backend"] != "file" || push["target"] != "saved_workbook" || push["vbe_validation"] != "not_performed" {
		t.Fatalf("push authority = %#v", push)
	}
	if target := cliObjectMap(env.Target); target["kind"] != "file" {
		t.Fatalf("target = %#v, want kind=file", target)
	}
	if !jsonHasWarningCode(t, stdout, "vbe_validation_skipped") {
		t.Fatalf("missing vbe_validation_skipped warning: %s", stdout)
	}
	if env.Backup == nil {
		t.Fatal("default push did not record a backup")
	}
	if _, err := os.Stat(filepath.Join(dir, ".xlflow", "state", "push.json")); err != nil {
		t.Fatalf("push state not written: %v", err)
	}
}

func TestPushFileBackendChangedOnlySkips(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePushSourceTree(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	stubClosedWorkbook(t)

	if stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file", "--changed-only"); err != nil {
		t.Fatalf("first push: %v\n%s", err, stdout)
	}
	stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file", "--changed-only")
	if err != nil {
		t.Fatalf("second push: %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	source := cliObjectMap(env.Source)
	if source["changed"] != false {
		t.Fatalf("unchanged push reported changed: %s", stdout)
	}
	if env.Backup != nil {
		t.Fatalf("skipped push created a backup: %s", stdout)
	}
}

func TestPushFileBackendRejectsOpenWorkbook(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePushSourceTree(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	lock := filepath.Join(dir, "build", "~$Book.xlsm")
	if err := os.WriteFile(lock, []byte{}, 0o644); err != nil {
		t.Fatal(err)
	}
	before, err := os.ReadFile(filepath.Join(dir, "build", "Book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}

	stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file")
	if err == nil || output.ExitCode(err) != output.ExitConfig || jsonErrorCode(t, stdout) != "push_workbook_open" {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
	after, readErr := os.ReadFile(filepath.Join(dir, "build", "Book.xlsm"))
	if readErr != nil || !bytes.Equal(before, after) {
		t.Fatal("workbook changed despite open-workbook rejection")
	}
}

func TestPushFileBackendRejectsActiveSession(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePushSourceTree(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), []byte(`{"pid":123,"workbook_path":"build/Book.xlsm"}`), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file")
	if err == nil || output.ExitCode(err) != output.ExitConfig || jsonErrorCode(t, stdout) != "push_active_session" {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
}

func TestPushFileBackendRejectsDuplicateModules(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePushSourceTree(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	writePackSourceModule(t, dir, filepath.Join("src", "modules", "nested", "module1.bas"), []byte("Attribute VB_Name = \"module1\"\nOption Explicit\n"))
	stubClosedWorkbook(t)

	stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file")
	if err == nil || output.ExitCode(err) != output.ExitValidation || jsonErrorCode(t, stdout) != "duplicate_module_name" {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
}

func TestPushFileBackendRejectsProtectedProject(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePushSourceTree(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p3_protected.bin"))
	stubClosedWorkbook(t)

	stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file")
	if err == nil || output.ExitCode(err) != output.ExitValidation || jsonErrorCode(t, stdout) != "push_protected_project" {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
}

func TestPushFileBackendRejectsNonXlsm(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	// Point the configured workbook at a non-macro-enabled extension.
	body, err := os.ReadFile(filepath.Join(dir, "xlflow.toml"))
	if err != nil {
		t.Fatal(err)
	}
	body = bytes.ReplaceAll(body, []byte("Book.xlsm"), []byte("Book.xlsx"))
	if err := os.WriteFile(filepath.Join(dir, "xlflow.toml"), body, 0o644); err != nil {
		t.Fatal(err)
	}
	stdout, err := runBuildCommandForTest(dir, "--json", "push", "--backend", "file")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
}

func TestFilePushErrorCodeMapping(t *testing.T) {
	for _, tc := range []struct {
		err  error
		want string
	}{
		{filepush.ErrDuplicateModule, "duplicate_module_name"},
		{filepush.ErrLineNumberSafety, "vba_line_number_safety_failed"},
		{filepush.ErrBackup, "push_backup_failed"},
		{errors.Join(filepush.ErrPublish, errors.New("io")), "push_write_failed"},
		{errors.New("other"), "push_failed"},
	} {
		if got := filePushErrorCode(tc.err); got != tc.want {
			t.Fatalf("filePushErrorCode(%v) = %q, want %q", tc.err, got, tc.want)
		}
	}
}
