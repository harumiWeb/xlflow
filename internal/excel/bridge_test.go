package excel

import (
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
