package excel

import (
	"encoding/base64"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestScriptPathFindsRepositoryScripts(t *testing.T) {
	path, err := scriptPath(t.TempDir(), "run")
	if err != nil {
		t.Fatal(err)
	}
	if path == "" {
		t.Fatal("expected script path")
	}
}

func TestRunnerTestFindsRepositoryScript(t *testing.T) {
	path, err := scriptPath(t.TempDir(), "test")
	if err != nil {
		t.Fatal(err)
	}
	if path == "" {
		t.Fatal("expected test script path")
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
		WorkbookPath: "fixtures\\Book.xlsm",
		Args: []RunArgument{
			{Type: "string", Value: "fixtures\\sample.xlsx"},
			{Type: "int", Value: "3"},
			{Type: "bool", Value: "true"},
		},
		SaveAs: "build\\Result.xlsm",
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
	if args["SaveAsPath"] != filepath.Join(root, "build", "Result.xlsm") {
		t.Fatalf("save-as path = %q", args["SaveAsPath"])
	}
	wantJSON := `[{"type":"string","value":"fixtures\\sample.xlsx"},{"type":"int","value":"3"},{"type":"bool","value":"true"}]`
	wantJSON64 := base64.StdEncoding.EncodeToString([]byte(wantJSON))
	if args["MacroArgsJSON"] != wantJSON64 {
		t.Fatalf("macro args json base64 = %s, want %s", args["MacroArgsJSON"], wantJSON64)
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
