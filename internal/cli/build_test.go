package cli

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestRootCommandIncludesBuildCommand(t *testing.T) {
	root := (&app{}).rootCommand()
	cmd, _, err := root.Find([]string{"build"})
	if err != nil || cmd == nil || cmd.Name() != "build" {
		t.Fatalf("build command = %#v, err = %v", cmd, err)
	}
	for _, name := range []string{"base", "out", "dry-run"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("build command missing --%s", name)
		}
	}
}

func TestCapabilitiesPublishesBuildCoordinationPolicy(t *testing.T) {
	stdout := new(bytes.Buffer)
	a := &app{stdout: stdout, stderr: new(bytes.Buffer), stdoutTerminal: func() bool { return false }}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "capabilities"})
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatal(err)
	}
	capabilities := cliObjectMap(env.Capabilities)
	commands := cliObjectMap(capabilities["commands"])
	build := cliObjectMap(commands["build"])
	if build["resource_scope"] != "none" || build["operation_kind"] != "read" || build["parallel_safe"] != true || build["retryable_when_busy"] != false || build["recovery_behavior"] != "not_applicable" {
		t.Fatalf("build capability = %#v", build)
	}
	buildCommand, _, err := root.Find([]string{"build"})
	if err != nil || shouldDelegateCommand(buildCommand, topLevelCommandName(buildCommand)) {
		t.Fatalf("build dry-run must stay local: command=%#v, err=%v", buildCommand, err)
	}
}

func TestBuildDryRunReportsPlanWithoutWriting(t *testing.T) {
	dir := writeBuildProject(t, "Book.xlsm")
	before := readBuildTree(t, dir)
	stdout, err := runBuildCommandForTest(dir, "--json", "build", "--dry-run")
	if err != nil {
		t.Fatal(err)
	}
	if after := readBuildTree(t, dir); !reflect.DeepEqual(before, after) {
		t.Fatal("dry run mutated the project")
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	build := cliObjectMap(env.Build)
	if build["dry_run"] != true || build["base"] != "build/Book.xlsm" || build["output"] != "build/Release/Book.xlsm" {
		t.Fatalf("build payload = %#v", build)
	}
	included, ok := build["included"].([]any)
	if !ok || len(included) != 1 {
		t.Fatalf("included = %#v", build["included"])
	}
	component, _ := included[0].(map[string]any)
	if component["name"] != "Main" || component["source_path"] != "src/modules/Main.bas" {
		t.Fatalf("included component = %#v", component)
	}
	if _, err := os.Stat(filepath.Join(dir, "build", "Release")); !os.IsNotExist(err) {
		t.Fatalf("dry run created output directory: %v", err)
	}
	human, err := runBuildCommandForTest(dir, "build", "--dry-run")
	if err != nil || !strings.Contains(human, "Base:\n  build/Book.xlsm") || !strings.Contains(human, "Included:\n  Main (src/modules/Main.bas)") {
		t.Fatalf("human output = %q, err = %v", human, err)
	}
}

func TestBuildAcceptsProjectWorkbookFormatsAndOverrides(t *testing.T) {
	for _, ext := range []string{".xlsm", ".xlam", ".xlsb"} {
		t.Run(ext, func(t *testing.T) {
			dir := writeBuildProject(t, "Base"+ext)
			if err := os.WriteFile(filepath.Join(dir, "templates", "Override"+ext), []byte("workbook"), 0o644); err != nil {
				t.Fatal(err)
			}
			stdout, err := runBuildCommandForTest(dir, "--json", "build", "--dry-run", "--base", "templates/Override"+ext, "--out", "dist/Product"+ext)
			if err != nil {
				t.Fatal(err)
			}
			var env output.Envelope
			if json.Unmarshal([]byte(stdout), &env) != nil {
				t.Fatalf("invalid JSON: %s", stdout)
			}
			build := cliObjectMap(env.Build)
			if build["base"] != "templates/Override"+ext || build["output"] != "dist/Product"+ext {
				t.Fatalf("build payload = %#v", build)
			}
		})
	}
}

func TestBuildValidationAndNonDryRunBoundary(t *testing.T) {
	dir := writeBuildProject(t, "Book.xlsm")
	for _, tc := range []struct {
		name, want string
		args       []string
		code       int
	}{
		{name: "extension mismatch", args: []string{"--json", "build", "--dry-run", "--out", "dist/Book.xlam"}, want: "build_args_invalid", code: output.ExitConfig},
		{name: "same file", args: []string{"--json", "build", "--dry-run", "--out", "build/Book.xlsm"}, want: "build_plan_invalid", code: output.ExitConfig},
		{name: "missing base", args: []string{"--json", "build", "--dry-run", "--base", "missing.xlsm"}, want: "build_args_invalid", code: output.ExitConfig},
		{name: "pipeline pending", args: []string{"--json", "build"}, want: "build_not_implemented", code: output.ExitEnvironment},
	} {
		t.Run(tc.name, func(t *testing.T) {
			stdout, err := runBuildCommandForTest(dir, tc.args...)
			if err == nil || output.ExitCode(err) != tc.code {
				t.Fatalf("err=%v, exit=%d, output=%s", err, output.ExitCode(err), stdout)
			}
			if got := errorCodeFromJSON(t, stdout); got != tc.want {
				t.Fatalf("code=%q, want %q: %s", got, tc.want, stdout)
			}
		})
	}
}

func runBuildCommandForTest(dir string, args ...string) (string, error) {
	stdout := new(bytes.Buffer)
	a := &app{cwd: dir, stdout: stdout, stderr: new(bytes.Buffer), stdoutTerminal: func() bool { return false }}
	root := a.rootCommand()
	root.SetArgs(args)
	err := root.Execute()
	return stdout.String(), err
}

func writeBuildProject(t *testing.T, workbook string) string {
	t.Helper()
	dir := t.TempDir()
	cfg := config.Default()
	cfg.Excel.Path = filepath.ToSlash(filepath.Join("build", workbook))
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{"build", "templates", "src/modules", "src/classes", "src/forms", "src/workbook"} {
		if err := os.MkdirAll(filepath.Join(dir, path), 0o755); err != nil {
			t.Fatal(err)
		}
	}
	if err := os.WriteFile(filepath.Join(dir, "build", workbook), []byte("workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "src", "modules", "Main.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	return dir
}

func readBuildTree(t *testing.T, root string) map[string]string {
	t.Helper()
	result := map[string]string{}
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		body, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		rel, _ := filepath.Rel(root, path)
		result[filepath.ToSlash(rel)] = string(body)
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	return result
}
