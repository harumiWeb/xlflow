package excel

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestExternalScriptPathFindsRepositoryScripts(t *testing.T) {
	path, ok := externalScriptPath(t.TempDir(), "run")
	if !ok {
		t.Fatal("expected repository script path")
	}
	if path == "" {
		t.Fatal("expected script path")
	}
}

func TestScriptPathPrefersRootScriptsDirectory(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	want := filepath.Join(scriptsDir, "run.ps1")
	if err := os.WriteFile(want, []byte("Write-Output 'override'\n"), 0o600); err != nil {
		t.Fatal(err)
	}

	path, cleanup, err := scriptPath(root, "run")
	if err != nil {
		t.Fatal(err)
	}
	if cleanup != nil {
		t.Fatal("expected on-disk override without cleanup")
	}
	if path != want {
		t.Fatalf("script path = %q, want %q", path, want)
	}
}

func TestMaterializeBundledScriptWritesCompleteBundle(t *testing.T) {
	path, cleanup, err := materializeBundledScript("ui")
	if err != nil {
		t.Fatal(err)
	}
	if cleanup == nil {
		t.Fatal("expected cleanup for materialized bundle")
	}
	dir := filepath.Dir(path)
	if filepath.Base(path) != "ui.ps1" {
		t.Fatalf("script path = %q, want bundled ui.ps1", path)
	}
	if _, err := os.Stat(filepath.Join(dir, "common.ps1")); err != nil {
		t.Fatalf("expected bundled common.ps1: %v", err)
	}
	cleanup()
	if _, err := os.Stat(dir); !os.IsNotExist(err) {
		t.Fatalf("expected cleanup to remove %q, got %v", dir, err)
	}
}

func TestScriptResultAcceptsScalarLogString(t *testing.T) {
	var result ScriptResult
	body := []byte(`{"status":"ok","command":"session","logs":"stopped xlflow Excel session"}`)
	if err := json.Unmarshal(body, &result); err != nil {
		t.Fatal(err)
	}
	if len(result.Logs) != 1 || result.Logs[0] != "stopped xlflow Excel session" {
		t.Fatalf("unexpected logs: %+v", result.Logs)
	}
}

func TestBuildUIButtonAddScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet:       "Menu",
		Cell:        "B2",
		Text:        "Run",
		Macro:       "Main.Run",
		ID:          "run",
		Width:       160,
		Height:      40,
		CreateSheet: true,
		VerifyMacro: true,
	})
	if args["Action"] != "add" {
		t.Fatalf("action = %q, want add", args["Action"])
	}
	if args["WorkbookPath"] != filepath.Join(root, "build", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["Sheet"] != "Menu" || args["Cell"] != "B2" || args["Text"] != "Run" || args["Macro"] != "Main.Run" || args["Id"] != "run" {
		t.Fatalf("unexpected args: %+v", args)
	}
	if args["Width"] != "160" || args["Height"] != "40" || args["CreateSheet"] != "true" || args["VerifyMacro"] != "true" {
		t.Fatalf("unexpected numeric/bool args: %+v", args)
	}
}

func TestUIValidationFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"sheet_not_found", "button_not_found", "ui_button_args_invalid"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitValidation {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitValidation)
			}
		})
	}
}

func TestTestFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"test_failed", "no_tests_found", "test_not_found", "duplicate_test_name"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitValidation {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitValidation)
			}
		})
	}
}

func TestRunnerTestReturnsEnvironmentFailureOnNonWindows(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("non-Windows behavior")
	}
	env, code, err := Runner{RootDir: t.TempDir()}.Test(config.Default(), "")
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	if env.Command != "test" {
		t.Fatalf("command = %q, want test", env.Command)
	}
}

func TestBuildRunScriptArgsSerializesArgumentsAndSaveAs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:        "Report.Generate",
		WorkbookPath: filepath.Join("fixtures", "Book.xlsm"),
		Args: []RunArgument{
			{Type: "string", Value: "fixtures\\sample.xlsx"},
			{Type: "int", Value: "3"},
			{Type: "bool", Value: "true"},
		},
		SaveAs: filepath.Join("build", "Result.xlsm"),
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["MacroName"] != "Report.Generate" {
		t.Fatalf("macro name = %q", args["MacroName"])
	}
	if args["WorkbookPath"] != filepath.Join(root, "fixtures", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["SaveWorkbook"] != "false" {
		t.Fatalf("save flag = %q", args["SaveWorkbook"])
	}
	if args["Direct"] != "false" || args["UseSession"] != "false" {
		t.Fatalf("unexpected direct/session defaults: %+v", args)
	}
	if args["SaveAsPath"] != filepath.Join(root, "build", "Result.xlsm") {
		t.Fatalf("save-as path = %q", args["SaveAsPath"])
	}
	wantJSON := `[{"type":"string","value":"fixtures\\sample.xlsx"},{"type":"int","value":"3"},{"type":"bool","value":"true"}]`
	wantJSON64 := base64.StdEncoding.EncodeToString([]byte(wantJSON))
	if args["MacroArgsJSON"] != wantJSON64 {
		t.Fatalf("macro args json base64 = %s, want %s", args["MacroArgsJSON"], wantJSON64)
	}
}

func TestBuildRunScriptArgsPassesFastDirectAndSession(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:               "Main.Run",
		Fast:                true,
		Session:             true,
		SuppressModalErrors: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Direct"] != "true" || args["UseSession"] != "true" {
		t.Fatalf("unexpected args: %+v", args)
	}
	if args["SuppressModalErrors"] != "true" {
		t.Fatalf("SuppressModalErrors = %q, want true", args["SuppressModalErrors"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("metadata path = %q", args["MetadataPath"])
	}
}

func TestBuildRunScriptArgsDiagnosticDisablesFastDirect(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:      "Main.Run",
		Fast:       true,
		Diagnostic: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Diagnostic"] != "true" {
		t.Fatalf("Diagnostic = %q, want true", args["Diagnostic"])
	}
	if args["Direct"] != "false" {
		t.Fatalf("Direct = %q, want false for diagnostic fast run", args["Direct"])
	}
}

func TestBuildRunScriptArgsPropagatesSuppressModalErrors(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:               "Main.Run",
		SuppressModalErrors: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["SuppressModalErrors"] != "true" {
		t.Fatalf("SuppressModalErrors = %q, want true", args["SuppressModalErrors"])
	}
}

func TestMacroFailureIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "macro_failed", Message: "boom"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(macro_failed) = %d, want %d", got, output.ExitValidation)
	}
}

func TestMacroDisabledIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "macro_disabled", Message: "disabled"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(macro_disabled) = %d, want %d", got, output.ExitValidation)
	}
}

func TestBuildRunScriptArgsNormalizesNilArgsToEmptyArray(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:        "Sheet1.Main",
		WorkbookPath: "Book.xlsm",
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["MacroArgsJSON"] != base64.StdEncoding.EncodeToString([]byte("[]")) {
		t.Fatalf("macro args json = %q, want base64 of []", args["MacroArgsJSON"])
	}
}

func TestBuildRunScriptArgsEnablesTrace(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro: "Main.Run",
		Trace: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["TraceEnabled"] != "true" {
		t.Fatalf("trace enabled = %q, want true", args["TraceEnabled"])
	}
	if args["TraceFile"] == "" {
		t.Fatal("expected trace file path")
	}
	if filepath.Base(filepath.Dir(args["TraceFile"])) != "traces" || filepath.Base(filepath.Dir(filepath.Dir(args["TraceFile"]))) != ".xlflow" {
		t.Fatalf("trace file path = %q, expected .xlflow traces directory", args["TraceFile"])
	}
}

func TestTraceNotInjectedIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "trace_not_injected", Message: "trace missing"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(trace_not_injected) = %d, want %d", got, output.ExitValidation)
	}
}

func TestVBACompileFailedIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "vba_compile_failed", Message: "compile failed"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(vba_compile_failed) = %d, want %d", got, output.ExitValidation)
	}
}

func TestMacroNotFoundIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "macro_not_found", Message: "missing"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(macro_not_found) = %d, want %d", got, output.ExitValidation)
	}
}

func TestDuplicateModuleNameIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "duplicate_module_name", Message: "duplicate"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(duplicate_module_name) = %d, want %d", got, output.ExitValidation)
	}
}

func TestPullScriptArgsIncludeFolderConfig(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildPullScriptArgs(root, cfg, SessionCommandOptions{})
	if args["Folders"] != "true" || args["FolderAnnotation"] != "update" || args["DefaultComponentFolders"] != "true" {
		t.Fatalf("unexpected folder config args: %+v", args)
	}
}

func TestTraceInjectScriptArgsIncludeModulesDirForConfiguredWorkbook(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTraceInjectScriptArgs(root, cfg, "")
	if args["ModulesDir"] != filepath.Join(root, "src", "modules") {
		t.Fatalf("modules dir = %q", args["ModulesDir"])
	}
}

func TestTraceScriptArgsPassSessionFlag(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTraceScriptArgs(root, cfg, TraceOptions{Action: "status", Session: true})
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("metadata path = %q", args["MetadataPath"])
	}
}

func TestTraceInjectScriptArgsOmitModulesDirForStandaloneWorkbook(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTraceInjectScriptArgs(root, cfg, "other.xlsm")
	if _, ok := args["ModulesDir"]; ok {
		t.Fatalf("standalone workbook should not receive ModulesDir: %+v", args)
	}
}

func TestStartKeepaliveWritesImmediateAndPeriodicHeartbeat(t *testing.T) {
	var stderr bytes.Buffer
	stop := startKeepalive("run", CommandOptions{
		Keepalive:         true,
		KeepaliveInterval: 10 * time.Millisecond,
		Stderr:            &stderr,
	})
	deadline := time.Now().Add(500 * time.Millisecond)
	for time.Now().Before(deadline) {
		if strings.Count(stderr.String(), "xlflow: run still running...") >= 2 {
			break
		}
		time.Sleep(10 * time.Millisecond)
	}
	stop()

	got := stderr.String()
	if !strings.Contains(got, "xlflow: run still running... elapsed=0s") {
		t.Fatalf("missing immediate heartbeat:\n%s", got)
	}
	if strings.Count(got, "xlflow: run still running...") < 2 {
		t.Fatalf("expected periodic heartbeat after immediate line:\n%s", got)
	}
}

func TestWriteDoneMarkerWritesSuccessAndFailure(t *testing.T) {
	var stderr bytes.Buffer
	writeDoneMarker("push", output.New("push"), CommandOptions{Keepalive: true, Stderr: &stderr})
	writeDoneMarker("run", output.Failure("run", output.Error{Code: "macro_timeout", Message: "timed out"}), CommandOptions{Keepalive: true, Stderr: &stderr})

	got := stderr.String()
	if !strings.Contains(got, "XLFLOW_DONE status=success command=push\n") {
		t.Fatalf("missing success marker:\n%s", got)
	}
	if !strings.Contains(got, "XLFLOW_DONE status=failed command=run code=macro_timeout\n") {
		t.Fatalf("missing failure marker:\n%s", got)
	}
	if strings.Count(got, "XLFLOW_DONE") != 2 {
		t.Fatalf("expected exactly two done markers:\n%s", got)
	}
}

func TestKeepaliveDoesNotWriteWhenDisabled(t *testing.T) {
	var stderr bytes.Buffer
	stop := startKeepalive("push", CommandOptions{Stderr: &stderr})
	stop()
	writeDoneMarker("push", output.New("push"), CommandOptions{Stderr: &stderr})
	if got := stderr.String(); got != "" {
		t.Fatalf("disabled keepalive wrote output:\n%s", got)
	}
}
