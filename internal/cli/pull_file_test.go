package cli

import (
	"bytes"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/filepull"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/workbookuse"
	"github.com/harumiWeb/xlflow/internal/wsl"
)

func TestPullDefaultsToAutoFileBackend(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	restore := stubWorkbookUseDetector(t, workbookuse.State{})
	defer restore()

	stdout, err := runBuildCommandForTest(dir, "--json", "pull")
	if err != nil {
		t.Fatalf("auto pull: %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	pull := cliObjectMap(env.Pull)
	if pull["backend"] != "file" || pull["backend_selection"] != "auto" || pull["selection_reason"] != "file_backend_supported" || pull["source"] != "saved_workbook" {
		t.Fatalf("auto pull metadata = %#v", pull)
	}
}

func TestAutoPullSelectionMatrix(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows selection matrix")
	}
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: dir}

	restore := stubWorkbookUseDetector(t, workbookuse.State{OpenInExcel: true, ExcelPIDs: []uint32{42}})
	selection, err := a.selectPullBackend(cfg, "auto", false)
	restore()
	if err != nil || selection.Backend != "excel" || selection.Reason != "workbook_open_in_excel" || !selection.AttachOpen {
		t.Fatalf("open selection = %+v, %v", selection, err)
	}

	restore = stubWorkbookUseDetectorError(t, errors.New("restart manager unavailable"))
	selection, err = a.selectPullBackend(cfg, "auto", false)
	restore()
	if err != nil || selection.Backend != "excel" || selection.Reason != "open_state_probe_failed" || !selection.AttachOpen || selection.Warning == nil {
		t.Fatalf("failed probe selection = %+v, %v", selection, err)
	}

	selection, err = a.selectPullBackend(cfg, "auto", true)
	if err != nil || selection.Backend != "excel" || selection.Reason != "session_requested" {
		t.Fatalf("session selection = %+v, %v", selection, err)
	}
	selection, err = a.selectPullBackend(cfg, "file", false)
	if err != nil || selection.Mode != "explicit" || selection.Reason != "explicit_backend" {
		t.Fatalf("explicit selection = %+v, %v", selection, err)
	}
}

func TestDelegatedAutoPullPreservesWSLProbeFailure(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("delegated auto selection runs in Windows xlflow")
	}
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	t.Setenv(wsl.EnvDelegated, "1")
	t.Setenv(envPullAutoProbeWarning, "WSL session status unavailable")
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}

	selection, err := (&app{cwd: dir}).selectPullBackend(cfg, "auto", false)
	if err != nil {
		t.Fatal(err)
	}
	if selection.Backend != "excel" || selection.Reason != "open_state_probe_failed" || !selection.AttachOpen || selection.Warning == nil || selection.Warning.Error() != "WSL session status unavailable" {
		t.Fatalf("selection = %+v", selection)
	}
}

func TestAutoPullDelegatesWhenWSLSessionStateCannotBeVerified(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), []byte("{"), 0o600); err != nil {
		t.Fatal(err)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	previousIsWSL := isWSL
	isWSL = func() bool { return true }
	t.Cleanup(func() { isWSL = previousIsWSL })

	selection, err := (&app{cwd: dir}).selectPullBackend(cfg, "auto", false)
	if err != nil {
		t.Fatal(err)
	}
	if selection.Backend != "excel" || selection.Reason != "open_state_probe_failed" || selection.Warning == nil || !selection.Delegate {
		t.Fatalf("selection = %+v", selection)
	}
}

func TestAutoPullMatchingLiveSessionWinsOverOpenWorkbook(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows session selection")
	}
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), []byte(`{"pid":42,"workbook_path":"build/Book.xlsm"}`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	restore := stubWorkbookUseDetector(t, workbookuse.State{OpenInExcel: true, ExcelPIDs: []uint32{42}})
	defer restore()
	selection, err := (&app{cwd: dir}).selectPullBackend(cfg, "auto", false)
	if err != nil || selection.Reason != "matching_live_session" || selection.AttachOpen {
		t.Fatalf("selection = %+v, %v", selection, err)
	}
}

func TestAutoPullPreservesProtectedProjectValidation(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p3_protected.bin"))
	restore := stubWorkbookUseDetector(t, workbookuse.State{})
	defer restore()
	stdout, err := runBuildCommandForTest(dir, "--json", "pull")
	if err == nil || output.ExitCode(err) != output.ExitValidation || jsonErrorCode(t, stdout) != "pull_protected_project" {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
}

func TestAutoPullSelectsExcelForUserForm(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows UserForm fallback")
	}
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p4_form.bin"))
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	restore := stubWorkbookUseDetector(t, workbookuse.State{})
	defer restore()
	selection, err := (&app{cwd: dir}).selectPullBackend(cfg, "auto", false)
	if err != nil || selection.Backend != "excel" || selection.Reason != "file_backend_unsupported_userform" {
		t.Fatalf("selection = %+v, %v", selection, err)
	}
}

func stubWorkbookUseDetector(t *testing.T, state workbookuse.State) func() {
	t.Helper()
	previous := detectWorkbookUse
	detectWorkbookUse = func(string) (workbookuse.State, error) { return state, nil }
	return func() { detectWorkbookUse = previous }
}

func stubWorkbookUseDetectorError(t *testing.T, detectErr error) func() {
	t.Helper()
	previous := detectWorkbookUse
	detectWorkbookUse = func(string) (workbookuse.State, error) { return workbookuse.State{}, detectErr }
	return func() { detectWorkbookUse = previous }
}

func TestPullFileBackendPublishesSavedWorkbookSource(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))

	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file")
	if err != nil {
		t.Fatalf("pull --backend file: %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	pull := cliObjectMap(env.Pull)
	if pull["backend"] != "file" || pull["backend_selection"] != "explicit" || pull["selection_reason"] != "explicit_backend" || pull["source"] != "saved_workbook" {
		t.Fatalf("pull authority = %#v", pull)
	}
	if target := cliObjectMap(env.Target); target["kind"] != "file" {
		t.Fatalf("target = %#v, want kind=file", target)
	}
	for _, path := range []string{
		filepath.Join(dir, "src", "modules", "Module1.bas"),
		filepath.Join(dir, "src", "classes", "Class1.cls"),
		filepath.Join(dir, "src", "workbook", "Sheet1.bas"),
		filepath.Join(dir, "src", "workbook", "ThisWorkbook.bas"),
	} {
		body, readErr := os.ReadFile(path)
		if readErr != nil || strings.HasPrefix(string(body), "\xef\xbb\xbf") || !strings.HasSuffix(string(body), "\n") {
			t.Fatalf("source %s was not published as UTF-8 without BOM and final newline: %v", path, readErr)
		}
	}
}

func TestPullFileBackendWarnsWhenMatchingSessionIsIgnored(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), []byte(`{"pid":123,"workbook_path":"build/Book.xlsm"}`), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file")
	if err != nil {
		t.Fatal(err)
	}
	if !jsonHasWarningCode(t, stdout, "file_pull_live_session_ignored") {
		t.Fatalf("missing live-session warning: %s", stdout)
	}
}

func TestPullFileBackendRejectsUserFormBeforeMutation(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p4_form.bin"))
	existing := filepath.Join(dir, "src", "modules", "keep.bas")
	if err := os.WriteFile(existing, []byte("keep"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("error = %v, exit = %d; want validation failure", err, output.ExitCode(err))
	}
	if jsonErrorCode(t, stdout) != "pull_userform_unsupported" {
		t.Fatalf("unexpected failure: %s", stdout)
	}
	body, readErr := os.ReadFile(existing)
	if readErr != nil || string(body) != "keep" {
		t.Fatalf("source mutated after rejection: %q, %v", body, readErr)
	}
}

func TestPullFileBackendRejectsSessionFlag(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file", "--session")
	if err == nil || output.ExitCode(err) != output.ExitConfig || jsonErrorCode(t, stdout) != "pull_args_invalid" {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
}

func TestFilePullPublicationFailureKeepsPublicationErrorCode(t *testing.T) {
	err := errors.Join(filepull.ErrPublish, &os.PathError{Op: "write", Path: "src/modules/Main.bas", Err: os.ErrPermission})
	if got := filePullErrorCode(err); got != "pull_source_publish_failed" {
		t.Fatalf("filePullErrorCode() = %q", got)
	}
}

func TestPullFileBackendCoordinatesSourceTreeInsteadOfWorkbook(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}

	workbookIdentity, err := coordination.NewWorkbookIdentity(dir, filepath.Join(dir, "build", "Book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	workbookLease, err := manager.Acquire(t.Context(), coordination.AcquireRequest{
		Identity: workbookIdentity, Command: "push", OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	stdout, runErr := runPullCommandWithManager(dir, manager, "--json", "pull", "--backend", "file")
	if runErr != nil {
		t.Fatalf("file pull contended on workbook lease: %v\n%s", runErr, stdout)
	}
	stdout, runErr = runPullCommandWithManager(dir, manager, "--json", "pull")
	if releaseErr := workbookLease.Release(); releaseErr != nil {
		t.Fatal(releaseErr)
	}
	if runErr != nil {
		t.Fatalf("auto file pull contended on workbook lease: %v\n%s", runErr, stdout)
	}

	sourceIdentity, err := coordination.NewSourceTreeIdentity(dir, filepath.Join(dir, "src", "modules"))
	if err != nil {
		t.Fatal(err)
	}
	sourceLease, err := manager.Acquire(t.Context(), coordination.AcquireRequest{
		Identity: sourceIdentity, Command: "pull", OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceSourceTree,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = sourceLease.Release() }()
	stdout, runErr = runPullCommandWithManager(dir, manager, "--json", "pull", "--backend", "file")
	if runErr == nil || output.ExitCode(runErr) != output.ExitEnvironment || jsonErrorCode(t, stdout) != coordination.SourceTreeBusyCode {
		t.Fatalf("source-tree contention stdout=%s err=%v", stdout, runErr)
	}
}

func TestAutoExcelPullRetainsWorkbookCoordination(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows workbook coordination and Excel fallback")
	}
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p4_form.bin"))
	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(dir, filepath.Join(dir, "build", "Book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	lease, err := manager.Acquire(t.Context(), coordination.AcquireRequest{
		Identity: identity, Command: "push", OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = lease.Release() }()
	restore := stubWorkbookUseDetector(t, workbookuse.State{})
	defer restore()

	stdout, runErr := runPullCommandWithManager(dir, manager, "--json", "pull")
	if runErr == nil || output.ExitCode(runErr) != output.ExitEnvironment || jsonErrorCode(t, stdout) != coordination.WorkbookBusyCode {
		t.Fatalf("auto Excel contention stdout=%s err=%v", stdout, runErr)
	}
}

func runPullCommandWithManager(dir string, manager *coordination.Manager, args ...string) (string, error) {
	stdout := new(bytes.Buffer)
	a := &app{cwd: dir, stdout: stdout, stderr: new(bytes.Buffer), stdoutTerminal: func() bool { return false }, coordination: manager}
	root := a.rootCommand()
	root.SetArgs(args)
	err := root.Execute()
	return stdout.String(), err
}

func jsonErrorCode(t *testing.T, body string) string {
	t.Helper()
	var envelope struct {
		Error *output.Error `json:"error"`
	}
	if err := json.Unmarshal([]byte(body), &envelope); err != nil {
		t.Fatal(err)
	}
	if envelope.Error == nil {
		return ""
	}
	return envelope.Error.Code
}

func jsonHasWarningCode(t *testing.T, body, code string) bool {
	t.Helper()
	var envelope struct {
		Warnings []struct {
			Code string `json:"code"`
		} `json:"warnings"`
	}
	if err := json.Unmarshal([]byte(body), &envelope); err != nil {
		t.Fatal(err)
	}
	for _, warning := range envelope.Warnings {
		if warning.Code == code {
			return true
		}
	}
	return false
}
