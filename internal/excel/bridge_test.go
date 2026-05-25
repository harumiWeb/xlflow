package excel

import (
	"encoding/base64"
	"encoding/json"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

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

func TestBuildRunScriptArgsPassesRuntimeMode(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:         "Main.Run",
		RuntimeMode:   RuntimeModeAgent,
		RuntimeSource: RuntimeSourceEnvironment,
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["RuntimeMode"] != RuntimeModeAgent {
		t.Fatalf("RuntimeMode = %q, want %q", args["RuntimeMode"], RuntimeModeAgent)
	}
	if args["RuntimeSource"] != RuntimeSourceEnvironment {
		t.Fatalf("RuntimeSource = %q, want %q", args["RuntimeSource"], RuntimeSourceEnvironment)
	}
}

func TestBuildRunScriptArgsPassesUIResponses(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro: "Main.Run",
		UIResponses: UIResponses{
			MsgBox:     map[string]string{"confirm-save": "yes"},
			Input:      map[string]string{"customer-name": "John"},
			FileDialog: []FileDialogResponse{{Kind: "get-open", DialogID: "source_files", Values: []string{"C:\\tmp\\a.txt", "C:\\tmp\\b.txt"}}},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if got, want := args["MsgBoxResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"confirm-save":"yes"}`)); got != want {
		t.Fatalf("MsgBoxResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["InputResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"customer-name":"John"}`)); got != want {
		t.Fatalf("InputResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["FileDialogResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`[{"kind":"get-open","dialog_id":"source_files","values":["C:\\tmp\\a.txt","C:\\tmp\\b.txt"]}]`)); got != want {
		t.Fatalf("FileDialogResponsesJSON = %q, want %q", got, want)
	}
}

func TestBuildRunScriptArgsPassesUIStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:    "Main.Run",
		UIStream: UIStreamOptions{Enabled: true, RedactInput: true},
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["UIStreamEnabled"] != "true" {
		t.Fatalf("UIStreamEnabled = %q, want true", args["UIStreamEnabled"])
	}
	if args["UIStreamRedactInput"] != "true" {
		t.Fatalf("UIStreamRedactInput = %q, want true", args["UIStreamRedactInput"])
	}
}

func TestBuildRunScriptArgsPassesDebugStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:       "Main.Run",
		DebugStream: DebugStreamOptions{Enabled: true, PipeName: `\\.\pipe\xlflow-debug-test`},
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["DebugStreamEnabled"] != "true" {
		t.Fatalf("DebugStreamEnabled = %q, want true", args["DebugStreamEnabled"])
	}
	if args["DebugStreamPipeName"] != `\\.\pipe\xlflow-debug-test` {
		t.Fatalf("DebugStreamPipeName = %q, want debug pipe name", args["DebugStreamPipeName"])
	}
}

func TestBuildTestScriptArgsPassesRuntimeMode(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{RuntimeMode: RuntimeModeTest, RuntimeSource: RuntimeSourceCommand})
	if args["RuntimeMode"] != RuntimeModeTest {
		t.Fatalf("RuntimeMode = %q, want %q", args["RuntimeMode"], RuntimeModeTest)
	}
	if args["RuntimeSource"] != RuntimeSourceCommand {
		t.Fatalf("RuntimeSource = %q, want %q", args["RuntimeSource"], RuntimeSourceCommand)
	}
}

func TestBuildTestScriptArgsPassesUIResponses(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{
		UIResponses: UIResponses{
			MsgBox:     map[string]string{"confirm-save": "no"},
			Input:      map[string]string{"customer-name": "Jane"},
			FileDialog: []FileDialogResponse{{Kind: "folder", DialogID: "target_dir", Cancelled: true}},
		},
	})
	if got, want := args["MsgBoxResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"confirm-save":"no"}`)); got != want {
		t.Fatalf("MsgBoxResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["InputResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`{"customer-name":"Jane"}`)); got != want {
		t.Fatalf("InputResponsesJSON = %q, want %q", got, want)
	}
	if got, want := args["FileDialogResponsesJSON"], base64.StdEncoding.EncodeToString([]byte(`[{"kind":"folder","dialog_id":"target_dir","cancelled":true}]`)); got != want {
		t.Fatalf("FileDialogResponsesJSON = %q, want %q", got, want)
	}
}

func TestBuildTestScriptArgsPassesUIStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{UIStream: UIStreamOptions{Enabled: true, RedactInput: true}})
	if args["UIStreamEnabled"] != "true" {
		t.Fatalf("UIStreamEnabled = %q, want true", args["UIStreamEnabled"])
	}
	if args["UIStreamRedactInput"] != "true" {
		t.Fatalf("UIStreamRedactInput = %q, want true", args["UIStreamRedactInput"])
	}
}

func TestBuildTestScriptArgsPassesDebugStreamOptions(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildTestScriptArgs(root, cfg, "", TestOptions{DebugStream: DebugStreamOptions{Enabled: true, PipeName: `\\.\pipe\xlflow-debug-test`}})
	if args["DebugStreamEnabled"] != "true" {
		t.Fatalf("DebugStreamEnabled = %q, want true", args["DebugStreamEnabled"])
	}
	if args["DebugStreamPipeName"] != `\\.\pipe\xlflow-debug-test` {
		t.Fatalf("DebugStreamPipeName = %q, want debug pipe name", args["DebugStreamPipeName"])
	}
}

func TestMergeUIResultAppendsStreamEvents(t *testing.T) {
	existing := map[string]any{"events": []any{map[string]any{"kind": "msgbox", "dialog_id": "existing"}}}
	merged := mergeUIResult(existing, []map[string]any{{"kind": "inputbox", "dialog_id": "customer-name"}})
	mergedMap, ok := merged.(map[string]any)
	if !ok {
		t.Fatalf("merged UI = %#v", merged)
	}
	events, ok := mergedMap["events"].([]any)
	if !ok || len(events) != 2 {
		t.Fatalf("merged events = %#v", mergedMap["events"])
	}
}

func TestMergeDebugResultPreservesExistingEvents(t *testing.T) {
	existing := map[string]any{
		"events":    []any{map[string]any{"message": "existing"}},
		"count":     4,
		"truncated": true,
	}
	streamed := map[string]any{
		"events": []any{map[string]any{"message": "streamed"}},
		"count":  3,
	}

	merged := mergeDebugResult(existing, streamed)
	mergedMap, ok := merged.(map[string]any)
	if !ok {
		t.Fatalf("merged debug = %#v", merged)
	}

	events := debugEventList(mergedMap["events"])
	if len(events) != 2 {
		t.Fatalf("merged debug events = %#v", mergedMap["events"])
	}
	if got := events[0]["message"]; got != "existing" {
		t.Fatalf("first merged debug event = %#v, want existing", got)
	}
	if got := events[1]["message"]; got != "streamed" {
		t.Fatalf("second merged debug event = %#v, want streamed", got)
	}
	if got := mergedMap["count"]; got != 4 {
		t.Fatalf("merged debug count = %#v, want 4", got)
	}
	if got := mergedMap["truncated"]; got != true {
		t.Fatalf("merged debug truncated = %#v, want true", got)
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
	if args["CodeSource"] != "sidecar" {
		t.Fatalf("CodeSource = %q", args["CodeSource"])
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
	if args["FormsDir"] != filepath.Join(root, "src", "forms") {
		t.Fatalf("FormsDir = %q", args["FormsDir"])
	}
	if args["CodeSource"] != "sidecar" {
		t.Fatalf("CodeSource = %q", args["CodeSource"])
	}
	if args["FolderAnnotation"] != "update" || args["Folders"] != "true" || args["DefaultComponentFolders"] != "true" {
		t.Fatalf("unexpected folder args: %+v", args)
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

func TestBuildUIButtonAddScriptArgsIncludesSessionMetadata(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet: "Menu",
		Cell:  "B2",
		Text:  "Run",
		Macro: "Main.Run",
	})
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonListScriptArgsIncludesSessionMetadata(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonListScriptArgs(root, cfg, UIButtonListOptions{Sheet: "Menu"})
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonRemoveScriptArgsIncludesSessionMetadata(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonRemoveScriptArgs(root, cfg, UIButtonRemoveOptions{ID: "run"})
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonAddScriptArgsPassSessionFlag(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet:   "Menu",
		Cell:    "B2",
		Text:    "Run",
		Macro:   "Main.Run",
		Session: true,
	})
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestBuildUIButtonListScriptArgsPassSessionFlag(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonListScriptArgs(root, cfg, UIButtonListOptions{Sheet: "Menu", Session: true})
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestBuildUIButtonRemoveScriptArgsPassSessionFlag(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonRemoveScriptArgs(root, cfg, UIButtonRemoveOptions{ID: "run", Session: true})
	if args["UseSession"] != "true" {
		t.Fatalf("UseSession = %q, want true", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q", args["MetadataPath"])
	}
}

func TestBuildUIButtonAddScriptArgsDefaultUseSessionFalse(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonAddScriptArgs(root, cfg, UIButtonAddOptions{
		Sheet: "Menu",
		Cell:  "B2",
		Text:  "Run",
		Macro: "Main.Run",
	})
	if args["UseSession"] != "false" {
		t.Fatalf("UseSession = %q, want false when Session option is not set", args["UseSession"])
	}
	if args["MetadataPath"] != filepath.Join(root, ".xlflow", "session.json") {
		t.Fatalf("MetadataPath = %q, want %q (always set for auto-detection)", args["MetadataPath"], filepath.Join(root, ".xlflow", "session.json"))
	}
}

func TestBuildUIButtonListScriptArgsDefaultUseSessionFalse(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonListScriptArgs(root, cfg, UIButtonListOptions{Sheet: "Menu"})
	if args["UseSession"] != "false" {
		t.Fatalf("UseSession = %q, want false when Session option is not set", args["UseSession"])
	}
}

func TestBuildUIButtonRemoveScriptArgsDefaultUseSessionFalse(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args := buildUIButtonRemoveScriptArgs(root, cfg, UIButtonRemoveOptions{ID: "run"})
	if args["UseSession"] != "false" {
		t.Fatalf("UseSession = %q, want false when Session option is not set", args["UseSession"])
	}
}

func TestProcessArgsInvalidIsConfigFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "process_args_invalid", Message: "invalid process args"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitConfig {
		t.Fatalf("exitCodeForScriptResult(process_args_invalid) = %d, want %d", got, output.ExitConfig)
	}
}

func TestProcessNotFoundIsConfigFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "process_not_found", Message: "process not found"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitConfig {
		t.Fatalf("exitCodeForScriptResult(process_not_found) = %d, want %d", got, output.ExitConfig)
	}
}

func TestProcessEnvironmentFailureCodesAreEnvironmentFailures(t *testing.T) {
	for _, code := range []string{"process_enumeration_failed", "process_termination_failed", "process_cleanup_failed"} {
		t.Run(code, func(t *testing.T) {
			result := ScriptResult{
				Status: output.StatusFailed,
				Error:  &output.Error{Code: code, Message: code},
			}
			if got := exitCodeForScriptResult(result); got != output.ExitEnvironment {
				t.Fatalf("exitCodeForScriptResult(%s) = %d, want %d", code, got, output.ExitEnvironment)
			}
		})
	}
}

func TestProcessListSerializesActionArg(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action)
$result = @{ status="ok"; command="process"; error=$null; logs=@(); process=@(@{action=$Action}) }
$result | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if env.Command != "process list" {
		t.Fatalf("Command = %q, want process list", env.Command)
	}
	processes, ok := env.Process.([]interface{})
	if !ok || len(processes) != 1 {
		t.Fatalf("Process = %v, want one-element array", env.Process)
	}
	m, ok := processes[0].(map[string]interface{})
	if !ok || m["action"] != "list" {
		t.Fatalf("Process[0].action = %v, want list", m["action"])
	}
}

func TestProcessListDefaultsActionArg(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action)
$result = @{ status="ok"; command="process"; error=$null; logs=@(); process=@(@{action=$Action}) }
$result | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.ProcessList(ProcessListOptions{})
	if err != nil {
		t.Fatal(err)
	}
	processes, ok := env.Process.([]interface{})
	if !ok || len(processes) != 1 {
		t.Fatalf("Process = %v, want one-element array", env.Process)
	}
	m, ok := processes[0].(map[string]interface{})
	if !ok || m["action"] != "list" {
		t.Fatalf("Process[0].action = %v, want list", m["action"])
	}
}

func TestProcessCleanupSerializesPidModeArgs(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action,[string]$TargetPid="",[string]$Auto="false",[string]$All="false")
$result = @{ status="ok"; command="process"; error=$null; logs=@(); process=@(@{action=$Action;targetPid=$TargetPid;auto=$Auto;all=$All}) }
$result | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.ProcessCleanup(ProcessCleanupOptions{Action: "cleanup", PID: 1234})
	if err != nil {
		t.Fatal(err)
	}
	if env.Command != "process cleanup" {
		t.Fatalf("Command = %q, want process cleanup", env.Command)
	}
	processes, ok := env.Process.([]interface{})
	if !ok || len(processes) != 1 {
		t.Fatalf("Process = %v, want one-element array", env.Process)
	}
	m, ok := processes[0].(map[string]interface{})
	if !ok {
		t.Fatalf("Process[0] is not a map: %T", processes[0])
	}
	if m["action"] != "cleanup" {
		t.Fatalf("action = %v, want cleanup", m["action"])
	}
	if m["targetPid"] != "1234" {
		t.Fatalf("targetPid = %v, want 1234", m["targetPid"])
	}
	if m["auto"] != "false" {
		t.Fatalf("auto = %v, want false", m["auto"])
	}
	if m["all"] != "false" {
		t.Fatalf("all = %v, want false", m["all"])
	}
}

func TestProcessCleanupDefaultsActionArg(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action,[string]$TargetPid="",[string]$Auto="false",[string]$All="false")
$result = @{ status="ok"; command="process"; error=$null; logs=@(); process=@(@{action=$Action;targetPid=$TargetPid;auto=$Auto;all=$All}) }
$result | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.ProcessCleanup(ProcessCleanupOptions{PID: 1234})
	if err != nil {
		t.Fatal(err)
	}
	processes, ok := env.Process.([]interface{})
	if !ok || len(processes) != 1 {
		t.Fatalf("Process = %v, want one-element array", env.Process)
	}
	m, ok := processes[0].(map[string]interface{})
	if !ok {
		t.Fatalf("Process[0] is not a map: %T", processes[0])
	}
	if m["action"] != "cleanup" {
		t.Fatalf("action = %v, want cleanup", m["action"])
	}
	if m["targetPid"] != "1234" {
		t.Fatalf("targetPid = %v, want 1234", m["targetPid"])
	}
}

func TestProcessCleanupSerializesAutoModeArgs(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action,[string]$TargetPid="",[string]$Auto="false",[string]$All="false")
$result = @{ status="ok"; command="process"; error=$null; logs=@(); process=@(@{action=$Action;targetPid=$TargetPid;auto=$Auto;all=$All}) }
$result | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.ProcessCleanup(ProcessCleanupOptions{Action: "cleanup", Auto: true})
	if err != nil {
		t.Fatal(err)
	}
	processes, ok := env.Process.([]interface{})
	if !ok || len(processes) != 1 {
		t.Fatalf("Process = %v, want one-element array", env.Process)
	}
	m, ok := processes[0].(map[string]interface{})
	if !ok {
		t.Fatalf("Process[0] is not a map: %T", processes[0])
	}
	if m["auto"] != "true" {
		t.Fatalf("auto = %v, want true", m["auto"])
	}
	if m["all"] != "false" {
		t.Fatalf("all = %v, want false", m["all"])
	}
	if m["targetPid"] != "" {
		t.Fatalf("targetPid = %v, want empty string", m["targetPid"])
	}
}

func TestProcessCleanupSerializesAllModeArgs(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action,[string]$TargetPid="",[string]$Auto="false",[string]$All="false")
$result = @{ status="ok"; command="process"; error=$null; logs=@(); process=@(@{action=$Action;targetPid=$TargetPid;auto=$Auto;all=$All}) }
$result | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, _, err := Runner{RootDir: root}.ProcessCleanup(ProcessCleanupOptions{Action: "cleanup", All: true})
	if err != nil {
		t.Fatal(err)
	}
	processes, ok := env.Process.([]interface{})
	if !ok || len(processes) != 1 {
		t.Fatalf("Process = %v, want one-element array", env.Process)
	}
	m, ok := processes[0].(map[string]interface{})
	if !ok {
		t.Fatalf("Process[0] is not a map: %T", processes[0])
	}
	if m["all"] != "true" {
		t.Fatalf("all = %v, want true", m["all"])
	}
	if m["auto"] != "false" {
		t.Fatalf("auto = %v, want false", m["auto"])
	}
}

func TestProcessListDecodesListResult(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action)
$result = @{
	status="ok"
	command="process"
	error=$null
	logs=@("found 2 Excel process(es)")
	process=@(
		@{pid=1234;has_workbook=$true}
		@{pid=5678;has_workbook=$false}
		@{pid=9012;has_workbook=$null}
	)
}
$result | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, code, err := Runner{RootDir: root}.ProcessList(ProcessListOptions{Action: "list"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Status != output.StatusOK {
		t.Fatalf("Status = %q, want ok", env.Status)
	}
	if env.Command != "process list" {
		t.Fatalf("Command = %q", env.Command)
	}
	processes, ok := env.Process.([]interface{})
	if !ok {
		t.Fatalf("Process is not an array: %T", env.Process)
	}
	if len(processes) != 3 {
		t.Fatalf("len(Process) = %d, want 3", len(processes))
	}
	p0, ok := processes[0].(map[string]interface{})
	if !ok {
		t.Fatalf("Process[0] is not a map: %T", processes[0])
	}
	if pid, ok := p0["pid"]; !ok || pid.(float64) != 1234 {
		t.Fatalf("Process[0].pid = %v", p0["pid"])
	}
	if hw, ok := p0["has_workbook"]; !ok || hw != true {
		t.Fatalf("Process[0].has_workbook = %v", p0["has_workbook"])
	}
}

func TestProcessCleanupDecodesCleanupResult(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action,[string]$TargetPid="",[string]$Auto="false",[string]$All="false")
'{"status":"ok","command":"process","error":null,"logs":["terminated 1 Excel process(es)"],"process":{"action":"cleanup","mode":"pid","total":1,"results":[{"pid":5678,"terminated":true,"method":"graceful"}]}}'
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, code, err := Runner{RootDir: root}.ProcessCleanup(ProcessCleanupOptions{Action: "cleanup", PID: 5678})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want %d", code, output.ExitSuccess)
	}
	if env.Status != output.StatusOK {
		t.Fatalf("Status = %q, want ok", env.Status)
	}
	if env.Command != "process cleanup" {
		t.Fatalf("Command = %q", env.Command)
	}
	p, ok := env.Process.(map[string]interface{})
	if !ok {
		t.Fatalf("Process is not a map: %T", env.Process)
	}
	if p["action"] != "cleanup" {
		t.Fatalf("action = %v, want cleanup", p["action"])
	}
	if p["mode"] != "pid" {
		t.Fatalf("mode = %v, want pid", p["mode"])
	}
	if total, ok := p["total"]; !ok || total.(float64) != 1 {
		t.Fatalf("total = %v", p["total"])
	}
	results, ok := p["results"].([]interface{})
	if !ok || len(results) != 1 {
		t.Fatalf("results = %v", p["results"])
	}
	r0, ok := results[0].(map[string]interface{})
	if !ok {
		t.Fatalf("results[0] is not a map: %T", results[0])
	}
	if r0["method"] != "graceful" {
		t.Fatalf("method = %v, want graceful", r0["method"])
	}
}

func TestProcessCleanupDecodesTerminationFailedResult(t *testing.T) {
	root := t.TempDir()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param([string]$Action,[string]$TargetPid="",[string]$Auto="false",[string]$All="false")
'{"status":"failed","command":"process","error":{"code":"process_termination_failed","message":"1 of 3 Excel process(es) failed to terminate","source":"","number":0,"line":0,"phase":""},"logs":null,"process":{"action":"cleanup","mode":"all","total":3,"results":[{"pid":1234,"terminated":true,"method":"force"},{"pid":5678,"terminated":true,"method":"force"},{"pid":9012,"terminated":false,"method":"none"}]}}'
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "process.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
	env, code, err := Runner{RootDir: root}.ProcessCleanup(ProcessCleanupOptions{Action: "cleanup", All: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want ExitEnvironment (%d)", code, output.ExitEnvironment)
	}
	if env.Status != output.StatusFailed {
		t.Fatalf("Status = %q, want failed", env.Status)
	}
	if env.Command != "process cleanup" {
		t.Fatalf("Command = %q", env.Command)
	}
	if env.Error == nil {
		t.Fatal("Error is nil, want process_termination_failed")
	}
	if env.Error.Code != "process_termination_failed" {
		t.Fatalf("Error.Code = %q, want process_termination_failed", env.Error.Code)
	}
	p, ok := env.Process.(map[string]interface{})
	if !ok {
		t.Fatalf("Process is not a map: %T", env.Process)
	}
	if p["mode"] != "all" {
		t.Fatalf("mode = %v, want all", p["mode"])
	}
	results, ok := p["results"].([]interface{})
	if !ok || len(results) != 3 {
		t.Fatalf("results = %v", p["results"])
	}
}
