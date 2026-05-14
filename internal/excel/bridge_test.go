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
	"github.com/harumiWeb/xlflow/internal/excel/forms"
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

func TestBuildExportImageScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildExportImageScriptArgs(root, cfg, ExportImageOptions{
		WorkbookPath: filepath.Join("fixtures", "Book.xlsm"),
		Sheet:        "QR",
		Range:        "A1:AE31",
		OutputDir:    filepath.Join("artifacts", "images"),
		Name:         "qr-demo",
		Format:       "png",
		Overwrite:    true,
		Session:      true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["WorkbookPath"] != filepath.Join(root, "fixtures", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["Sheet"] != "QR" || args["RangeAddress"] != "A1:AE31" {
		t.Fatalf("unexpected sheet/range args: %+v", args)
	}
	if args["OutputPath"] != filepath.Join(root, "artifacts", "images", "qr-demo.png") {
		t.Fatalf("output path = %q", args["OutputPath"])
	}
	if args["ImageFormat"] != "png" || args["Overwrite"] != "true" || args["UseSession"] != "true" {
		t.Fatalf("unexpected export args: %+v", args)
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("metadata path = %q", args["MetadataPath"])
	}
}

func TestBuildFormExportImageScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildFormExportImageScriptArgs(root, cfg, FormExportImageOptions{
		Name:        "UserForm1",
		OutPath:     filepath.Join("artifacts", "forms", "UserForm1.png"),
		Initializer: "InitializeForm",
		Overwrite:   true,
		Session:     true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["WorkbookPath"] != filepath.Join(root, "build", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["FormName"] != "UserForm1" || args["Initializer"] != "InitializeForm" {
		t.Fatalf("unexpected form args: %+v", args)
	}
	if args["OutputPath"] != filepath.Join(root, "artifacts", "forms", "UserForm1.png") {
		t.Fatalf("output path = %q", args["OutputPath"])
	}
	if args["Overwrite"] != "true" || args["UseSession"] != "true" {
		t.Fatalf("unexpected overwrite/session args: %+v", args)
	}
}

func TestBuildEditCellScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	value := "ABC123"
	args, err := buildEditCellScriptArgs(root, cfg, EditCellOptions{
		WorkbookPath: filepath.Join("fixtures", "Book.xlsm"),
		Sheet:        "Input",
		Cell:         "B2",
		Value:        &value,
		Events:       EditEventOn,
		Session:      true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Action"] != "cell" || args["WorkbookPath"] != filepath.Join(root, "fixtures", "Book.xlsm") {
		t.Fatalf("unexpected edit cell args: %+v", args)
	}
	if args["Sheet"] != "Input" || args["Cell"] != "B2" || args["Value"] != "ABC123" || args["Events"] != "on" {
		t.Fatalf("unexpected edit cell payload: %+v", args)
	}
	if args["UseSession"] != "true" || args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("unexpected session args: %+v", args)
	}
}

func TestBuildEditRangeRowsAndColumnsScriptArgs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	rangeArgs, err := buildEditRangeScriptArgs(root, cfg, EditRangeOptions{
		Sheet:   "QR",
		Range:   "A1:AE31",
		Fill:    "#FFFFFF",
		Session: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if rangeArgs["Action"] != "range" || rangeArgs["RangeAddress"] != "A1:AE31" || rangeArgs["Fill"] != "#FFFFFF" {
		t.Fatalf("unexpected edit range args: %+v", rangeArgs)
	}
	rowsArgs := buildEditRowsScriptArgs(root, cfg, EditRowsOptions{
		Sheet:   "QR",
		Rows:    "1:31",
		Height:  12,
		Session: true,
	})
	if rowsArgs["Action"] != "rows" || rowsArgs["Rows"] != "1:31" || rowsArgs["Height"] != "12" {
		t.Fatalf("unexpected edit rows args: %+v", rowsArgs)
	}
	columnArgs := buildEditColumnsScriptArgs(root, cfg, EditColumnsOptions{
		Sheet:   "QR",
		Columns: "A:AE",
		Width:   2.2,
		Session: true,
	})
	if columnArgs["Action"] != "columns" || columnArgs["Columns"] != "A:AE" || columnArgs["Width"] != "2.2" {
		t.Fatalf("unexpected edit columns args: %+v", columnArgs)
	}
}

func TestBuildEditScriptArgsRejectInvalidCombinations(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	value := "ABC123"
	formula := "=A1+B1"
	if _, err := buildEditCellScriptArgs(root, cfg, EditCellOptions{
		Sheet:   "Input",
		Cell:    "B2",
		Value:   &value,
		Formula: &formula,
		Session: true,
	}); err == nil || !strings.Contains(err.Error(), "exactly one") {
		t.Fatalf("expected cell mutation conflict, got %v", err)
	}
	if _, err := buildEditRangeScriptArgs(root, cfg, EditRangeOptions{
		Sheet:   "QR",
		Range:   "A1:B2",
		Fill:    "#FFFFFF",
		Clear:   "all",
		Session: true,
	}); err == nil || !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected range mutation conflict, got %v", err)
	}
}

func TestResolveExportImageOutputDefaultPath(t *testing.T) {
	root := t.TempDir()
	resolved, err := resolveExportImageOutput(root, filepath.Join("build", "Book.xlsm"), ExportImageOptions{
		Sheet:  "Sheet 1",
		Range:  "A1:AE31",
		Format: "png",
	})
	if err != nil {
		t.Fatal(err)
	}
	if !resolved.Default {
		t.Fatal("expected default output path")
	}
	wantDir := filepath.Join(root, ".xlflow", "artifacts", "images", "Book")
	if filepath.Dir(resolved.Path) != wantDir {
		t.Fatalf("dir = %q, want %q", filepath.Dir(resolved.Path), wantDir)
	}
	base := filepath.Base(resolved.Path)
	if !strings.HasPrefix(base, "Sheet_1_A1-AE31_") || !strings.HasSuffix(base, ".png") {
		t.Fatalf("unexpected default filename %q", base)
	}
}

func TestResolveFormExportImageOutputRejectsUnsupportedExtension(t *testing.T) {
	_, err := resolveFormExportImageOutput(t.TempDir(), FormExportImageOptions{OutPath: "artifacts\\UserForm1.webp"})
	if err == nil || !strings.Contains(err.Error(), "supported formats: png") {
		t.Fatalf("expected unsupported format error, got %v", err)
	}
}

func TestResolveFormExportImageOutputRejectsMissingExtension(t *testing.T) {
	_, err := resolveFormExportImageOutput(t.TempDir(), FormExportImageOptions{OutPath: "artifacts\\UserForm1"})
	if err == nil || !strings.Contains(err.Error(), `unsupported image format ""; supported formats: png`) {
		t.Fatalf("expected missing extension error, got %v", err)
	}
}

func TestNormalizeExportImagePathRejectsUnsupportedExtension(t *testing.T) {
	_, _, err := normalizeExportImagePath(t.TempDir(), "artifacts\\qr.webp", "png")
	if err == nil || !strings.Contains(err.Error(), "supported formats: png") {
		t.Fatalf("expected unsupported format error, got %v", err)
	}
}

func TestRunnerExportImageReturnsValidationFailureForUnsupportedFormat(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.ExportImage(config.Default(), ExportImageOptions{
		Sheet:  "QR",
		Range:  "A1:B2",
		Format: "webp",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerExportImageReturnsValidationFailureForUnsupportedExtension(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.ExportImage(config.Default(), ExportImageOptions{
		Sheet:   "QR",
		Range:   "A1:B2",
		OutPath: "artifacts\\qr.webp",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerFormExportImageReturnsValidationFailureForUnsupportedExtension(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.FormExportImage(config.Default(), FormExportImageOptions{
		Name:    "UserForm1",
		OutPath: "artifacts\\UserForm1.webp",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerFormExportImageReturnsValidationFailureForMissingExtension(t *testing.T) {
	env, code, err := Runner{RootDir: t.TempDir()}.FormExportImage(config.Default(), FormExportImageOptions{
		Name:    "UserForm1",
		OutPath: "artifacts\\UserForm1",
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitValidation {
		t.Fatalf("exit code = %d, want %d", code, output.ExitValidation)
	}
	if env.Error == nil || env.Error.Code != "unsupported_image_format" {
		t.Fatalf("unexpected error: %+v", env.Error)
	}
}

func TestRunnerInspectFormNormalizesCommandName(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$WorkbookPath,[string]$FormName,[string]$Basis,[string]$Initializer,[string]$Visible,[string]$UseSession,[string]$MetadataPath)
@{ status = "ok"; command = "inspect-form"; logs = @(); error = $null } | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "inspect-form.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.InspectForm(config.Default(), InspectFormOptions{Name: "UserForm1", Basis: "designer"})
	if err != nil {
		t.Fatal(err)
	}
	if env.Command != "inspect" {
		t.Fatalf("command = %q, want inspect", env.Command)
	}
}

func TestOutputFileExistsIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "output_file_exists", Message: "exists"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(output_file_exists) = %d, want %d", got, output.ExitValidation)
	}
}

func TestUnsupportedImageFormatIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "unsupported_image_format", Message: "bad format"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(unsupported_image_format) = %d, want %d", got, output.ExitValidation)
	}
}

func TestFormExportImageFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"runtime_form_load_failed", "form_initializer_failed", "window_not_found", "image_capture_failed"} {
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

func TestEditValidationFailureCodesAreValidationFailures(t *testing.T) {
	for _, code := range []string{"session_required", "invalid_color", "invalid_cell_address", "invalid_row_selector", "invalid_column_selector", "vba_event_error"} {
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

func TestRunnerEditReturnsConfigFailureForInvalidMutations(t *testing.T) {
	value := "ABC123"
	formula := "=A1+B1"
	env, code, err := Runner{RootDir: t.TempDir()}.EditCell(config.Default(), EditCellOptions{
		Sheet:   "Input",
		Cell:    "B2",
		Value:   &value,
		Formula: &formula,
		Session: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitConfig {
		t.Fatalf("exit code = %d, want %d", code, output.ExitConfig)
	}
	if env.Error == nil || env.Error.Code != "edit_args_invalid" {
		t.Fatalf("unexpected error: %+v", env.Error)
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

func TestListFormsScriptArgsIncludeFolderAndSessionConfig(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildListFormsScriptArgs(root, cfg, SessionCommandOptions{Session: true})
	if args["Action"] != "forms" {
		t.Fatalf("Action = %q, want forms", args["Action"])
	}
	if args["ProjectRoot"] != root {
		t.Fatalf("ProjectRoot = %q, want %q", args["ProjectRoot"], root)
	}
	if args["FormsDir"] != filepath.Join(root, "src", "forms") {
		t.Fatalf("FormsDir = %q", args["FormsDir"])
	}
	if args["FolderAnnotation"] != "update" || args["Folders"] != "true" || args["DefaultComponentFolders"] != "true" {
		t.Fatalf("unexpected folder config args: %+v", args)
	}
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestInspectFormScriptArgsIncludeSessionAndInitializer(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildInspectFormScriptArgs(root, cfg, InspectFormOptions{
		Name:           "UserForm1",
		Basis:          "runtime",
		Initializer:    "InitializeForm",
		StrictDesigner: true,
		Session:        true,
	})
	if args["WorkbookPath"] != filepath.Join(root, "build", "Book.xlsm") {
		t.Fatalf("WorkbookPath = %q", args["WorkbookPath"])
	}
	if args["FormName"] != "UserForm1" {
		t.Fatalf("FormName = %q", args["FormName"])
	}
	if args["Basis"] != "runtime" {
		t.Fatalf("Basis = %q", args["Basis"])
	}
	if args["Initializer"] != "InitializeForm" {
		t.Fatalf("Initializer = %q", args["Initializer"])
	}
	if args["StrictDesigner"] != "true" {
		t.Fatalf("StrictDesigner = %q", args["StrictDesigner"])
	}
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestFormWriteScriptArgsIncludeSpecPayloadAndSessionFlags(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildFormWriteScriptArgs(root, cfg, FormWriteOptions{
		Action:   "apply",
		SpecPath: "src/forms/UserForm1.form.yaml",
		Spec: forms.FormSpec{
			SchemaVersion: 1,
			Kind:          "xlflow.userform",
			Basis:         "designer",
			Form:          forms.FormSpecForm{Name: "UserForm1"},
			Controls: []forms.FormSpecControl{{
				Name: "txtCustomer",
				Type: "TextBox",
			}},
			Warnings: []forms.FormSpecWarning{},
		},
		Session: true,
		NoSave:  true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["Action"] != "apply" {
		t.Fatalf("Action = %q", args["Action"])
	}
	if args["SpecPath"] != "src/forms/UserForm1.form.yaml" {
		t.Fatalf("SpecPath = %q", args["SpecPath"])
	}
	if args["UseSession"] != "true" || args["NoSave"] != "true" {
		t.Fatalf("unexpected session flags: %+v", args)
	}
	if args["SpecJson64"] == "" {
		t.Fatalf("SpecJson64 should be set: %+v", args)
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
