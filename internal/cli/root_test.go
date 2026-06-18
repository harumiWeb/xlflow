package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"reflect"
	"runtime"
	"strings"
	"testing"
	"time"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/backup"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/vbafmt"
	"github.com/xuri/excelize/v2"
)

type stubReleaseChecker struct {
	release latestRelease
	err     error
}

type runOptionsInput struct {
	Macro              string
	Workbook           string
	Args               []string
	MsgBox             []string
	InputBox           []string
	FileDialog         []string
	Save               bool
	SaveAs             string
	Headless           bool
	Interactive        bool
	Direct             bool
	Fast               bool
	Diagnostic         bool
	DiagnosticExplicit bool
	GUICompileErrors   bool
	Session            bool
	Timeout            time.Duration
	CommandOptions     excel.CommandOptions
	UIStream           bool
}

func buildRunOptionsForTest(cfg config.Config, in runOptionsInput) (excel.RunOptions, error) {
	if in.Timeout == 0 {
		in.Timeout = 5 * time.Minute
	}
	return buildRunOptionsWithUIStream(
		cfg,
		in.Macro,
		in.Workbook,
		in.Args,
		in.MsgBox,
		in.InputBox,
		in.FileDialog,
		in.Save,
		in.SaveAs,
		in.Headless,
		in.Interactive,
		in.Direct,
		in.Fast,
		in.Diagnostic,
		in.DiagnosticExplicit,
		in.GUICompileErrors,
		in.Session,
		in.Timeout,
		in.CommandOptions,
		in.UIStream,
	)
}

func (s stubReleaseChecker) LatestRelease(ctx context.Context) (latestRelease, error) {
	return s.release, s.err
}

func TestRootCommandIncludesTestCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"test"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "test" {
		t.Fatalf("expected test command, got %#v", cmd)
	}
	if flag := cmd.Flags().Lookup("filter"); flag == nil {
		t.Fatal("expected test command to define --filter")
	}
	if flag := cmd.Flags().Lookup("ui-stream"); flag == nil {
		t.Fatal("expected test command to define --ui-stream")
	}
}

func TestTestCommandWritesProgressToStderrForNonInteractiveJSONRuns(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)

	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	writeTestTestScript(t, dir)

	var stdout bytes.Buffer
	var stderr bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &stderr,
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("--json", "test"))
	if err := root.Execute(); err != nil {
		t.Fatalf("test command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if got := stderr.String(); got != "xlflow: Running VBA tests...\n" {
		t.Fatalf("stderr progress = %q", got)
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout.String())
	}
	if env["command"] != "test" {
		t.Fatalf("expected command=test, got %v", env["command"])
	}
}

func TestRootCommandIncludesVersionCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"version"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "version" {
		t.Fatalf("expected version command, got %#v", cmd)
	}
}

func TestRootCommandIncludesBridgeFlag(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	flag := root.PersistentFlags().Lookup("bridge")
	if flag == nil {
		t.Fatal("expected persistent --bridge flag")
	}
}

func TestLoadConfigAllowsValidBridgeOverrideWhenConfigBridgeIsInvalid(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
bridge = "broken"
`)
	if err := os.WriteFile(filepath.Join(dir, config.FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: dir, bridge: "powershell"}
	cfg, err := a.loadConfig("fmt")
	if err != nil {
		t.Fatalf("loadConfig returned error: %v", err)
	}
	if cfg.Excel.Bridge != "broken" {
		t.Fatalf("excel.bridge = %q, want broken", cfg.Excel.Bridge)
	}
}

func TestRootCommandIncludesBackupAndRollbackCommands(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	backupCmd, _, err := root.Find([]string{"backup", "list"})
	if err != nil {
		t.Fatal(err)
	}
	if backupCmd == nil || backupCmd.Name() != "list" {
		t.Fatalf("expected backup list command, got %#v", backupCmd)
	}
	rollbackCmd, _, err := root.Find([]string{"rollback"})
	if err != nil {
		t.Fatal(err)
	}
	if rollbackCmd == nil || rollbackCmd.Name() != "rollback" {
		t.Fatalf("expected rollback command, got %#v", rollbackCmd)
	}
	for _, name := range []string{"latest", "backup"} {
		if rollbackCmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected rollback command to define --%s", name)
		}
	}
}

func TestVersionCommandWritesBuildInfoJSON(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{
		stdout:    &stdout,
		stderr:    &bytes.Buffer{},
		buildInfo: BuildInfo{Version: "1.2.3", Commit: "abc123", Date: "2026-05-02T00:00:00Z"},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "version"})

	if err := root.Execute(); err != nil {
		t.Fatalf("version command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Status  string    `json:"status"`
		Command string    `json:"command"`
		Version BuildInfo `json:"version"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Status != output.StatusOK {
		t.Fatalf("status = %q, want %q", got.Status, output.StatusOK)
	}
	if got.Command != "version" {
		t.Fatalf("command = %q, want version", got.Command)
	}
	if got.Version.Version != "1.2.3" || got.Version.Commit != "abc123" || got.Version.Date != "2026-05-02T00:00:00Z" {
		t.Fatalf("unexpected version payload: %#v", got.Version)
	}
}

func TestVersionCommandUsesDefaultBuildInfo(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{
		stdout:    &stdout,
		stderr:    &bytes.Buffer{},
		buildInfo: BuildInfo{}.withDefaults(),
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "version"})

	if err := root.Execute(); err != nil {
		t.Fatalf("version command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Version BuildInfo `json:"version"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Version.Version != "dev" || got.Version.Commit != "none" || got.Version.Date != "unknown" {
		t.Fatalf("unexpected default version payload: %#v", got.Version)
	}
}

func TestLintCommandJSONIncludesConfigWarnings(t *testing.T) {
	dir := writeCLIWarningLintProject(t, `forbid_select = false`)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "lint"})

	if err := root.Execute(); err != nil {
		t.Fatalf("lint command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Warnings []map[string]any `json:"warnings"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if !hasCLIWarning(got.Warnings, "deprecated_lint_rule_config", "VB002") {
		t.Fatalf("expected config deprecation warning, got %+v", got.Warnings)
	}
}

func TestLintCommandPlainIncludesConfigWarnings(t *testing.T) {
	dir := writeCLIWarningLintProject(t, `forbid_public_module_fields = true
disabled_rules = ["VB006"]`)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"lint"})

	if err := root.Execute(); err != nil {
		t.Fatalf("lint command error = %v, exit = %d", err, output.ExitCode(err))
	}
	text := stdout.String()
	for _, want := range []string{"Warnings:", "[lint].forbid_public_module_fields is deprecated", "lint rule VB006 is enabled", "[lint].disabled_rules takes precedence"} {
		if !strings.Contains(text, want) {
			t.Fatalf("expected lint output to contain %q:\n%s", want, text)
		}
	}
}

func TestAnalyzeCommandJSONIncludesConfigWarnings(t *testing.T) {
	dir := writeCLIWarningAnalyzeProject(t, `detect_byref_argument_mismatch = true`)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "analyze"})

	if err := root.Execute(); err != nil {
		t.Fatalf("analyze command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Warnings []map[string]any `json:"warnings"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if !hasCLIWarning(got.Warnings, "deprecated_analyze_rule_config", "VBA206") {
		t.Fatalf("expected analyze config deprecation warning, got %+v", got.Warnings)
	}
}

func writeCLIWarningLintProject(t *testing.T, lintConfig string) string {
	t.Helper()
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte("Option Explicit\nPublic Sub Run()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	body := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
` + lintConfig + "\n"
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	return dir
}

func writeCLIWarningAnalyzeProject(t *testing.T, analyzeConfig string) string {
	t.Helper()
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte("Option Explicit\nPublic Sub Run()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	body := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
` + analyzeConfig + "\n"
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	return dir
}

func hasCLIWarning(warnings []map[string]any, code string, rule string) bool {
	for _, warning := range warnings {
		if warning["code"] == code && warning["rule"] == rule {
			return true
		}
	}
	return false
}

func TestVersionCommandVerboseIncludesExecutableAndFeatures(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{
		stdout:    &stdout,
		stderr:    &bytes.Buffer{},
		buildInfo: BuildInfo{Version: "1.2.3", Commit: "abc123", Date: "2026-05-02T00:00:00Z"},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "version", "--verbose"})

	if err := root.Execute(); err != nil {
		t.Fatalf("version verbose command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Version struct {
			Version        string `json:"version"`
			ExecutablePath string `json:"executable_path"`
			Features       []struct {
				Name string `json:"name"`
			} `json:"features"`
		} `json:"version"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Version.Version != "1.2.3" {
		t.Fatalf("version = %q", got.Version.Version)
	}
	if got.Version.ExecutablePath == "" {
		t.Fatal("expected executable path in verbose version payload")
	}
	if len(got.Version.Features) == 0 {
		t.Fatal("expected supported features in verbose version payload")
	}
}

func TestResolvedVersionScriptsIncludesNewUserFormScripts(t *testing.T) {
	scripts := resolvedVersionScripts(t.TempDir())
	names := make([]string, 0, len(scripts))
	for _, script := range scripts {
		names = append(names, script.Command)
	}
	for _, want := range []string{"list", "inspect-form", "form-export-image"} {
		found := false
		for _, name := range names {
			if name == want {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("resolvedVersionScripts missing %q in %#v", want, names)
		}
	}
}

func TestRootCommandIncludesRunFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"run"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"arg", "msgbox", "inputbox", "filedialog", "input", "save", "no-save", "save-as", "headless", "interactive", "direct", "fast", "diagnostic", "gui-compile-errors", "session", "timeout", "ui-stream"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected run command to define --%s", name)
		}
	}
}

func TestRunCommandDiagnosticDefaultsTrue(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"run"})
	if err != nil {
		t.Fatal(err)
	}
	flag := cmd.Flags().Lookup("diagnostic")
	if flag == nil {
		t.Fatal("expected --diagnostic flag")
		return
	}
	if flag.DefValue != "true" {
		t.Fatalf("diagnostic default = %q, want true", flag.DefValue)
	}
}

func TestRunCommandIncludesVerboseFlag(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"run"})
	if err != nil {
		t.Fatal(err)
	}
	flag := cmd.Flags().Lookup("verbose")
	if flag == nil {
		t.Fatal("expected --verbose flag")
		return
	}
	if flag.DefValue != "false" {
		t.Fatalf("verbose default = %q, want false", flag.DefValue)
	}
}

func TestHeadlessGUIBoundaryLogsExplainProjectWideScanAndLintOverride(t *testing.T) {
	logs := headlessGUIBoundaryLogs(config.Default())
	for _, want := range []string{
		"Headless preflight scans the configured source tree, not the target macro call graph.",
		"Use xlflow run --interactive if a human can operate Excel dialogs.",
		"replace raw MsgBox/InputBox/file dialog calls with XlflowUI wrappers",
		"--msgbox/--inputbox/--filedialog",
		"[lint].forbid_interactive_input = false",
	} {
		found := false
		for _, line := range logs {
			if strings.Contains(line, want) {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("expected headless GUI boundary log containing %q in %#v", want, logs)
		}
	}

	cfg := config.Default()
	cfg.Lint.ForbidInteractiveInput = false
	logs = headlessGUIBoundaryLogs(cfg)
	for _, line := range logs {
		if strings.Contains(line, "[lint].forbid_interactive_input = false") {
			t.Fatalf("did not expect lint override hint when interactive-input lint is already disabled: %#v", logs)
		}
	}
}

func TestRootCommandIncludesPushFastFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"push"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"backup", "fast", "changed-only", "session", "no-save"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected push command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesSessionSaveAndRunnerCommands(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	for _, args := range [][]string{
		{"session", "start"},
		{"session", "status"},
		{"session", "stop"},
		{"save"},
		{"runner", "install"},
		{"runner", "remove"},
		{"runner", "status"},
	} {
		cmd, _, err := root.Find(args)
		if err != nil {
			t.Fatal(err)
		}
		if cmd == nil || cmd.Name() != args[len(args)-1] {
			t.Fatalf("expected command %v, got %#v", args, cmd)
		}
	}
}

func TestRootCommandIncludesInspectGUICommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"inspect-gui"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "inspect-gui" {
		t.Fatalf("expected inspect-gui command, got %#v", cmd)
	}
}

func TestRootCommandIncludesInspectSubcommands(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	for _, args := range [][]string{
		{"inspect", "workbook"},
		{"inspect", "sheets"},
		{"inspect", "form"},
		{"inspect", "range"},
		{"inspect", "used-range"},
		{"inspect", "cell"},
	} {
		cmd, _, err := root.Find(args)
		if err != nil {
			t.Fatal(err)
		}
		if cmd == nil || cmd.Name() != args[len(args)-1] {
			t.Fatalf("expected inspect command %v, got %#v", args, cmd)
		}
	}
}

func TestRootCommandIncludesAttachActiveCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"attach"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "attach" {
		t.Fatalf("expected attach command, got %#v", cmd)
	}
	if cmd.Flags().Lookup("active") == nil {
		t.Fatal("expected attach command to define --active")
	}
}

func TestRootCommandRejectsRemovedTraceCommand(t *testing.T) {
	a := &app{
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		cwd:            t.TempDir(),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"trace"})
	err := root.Execute()
	if err == nil || !strings.Contains(err.Error(), "unknown command") {
		t.Fatalf("expected unknown command error, got %v", err)
	}
}

func TestRootCommandIncludesAnalyzeAndCheckCommands(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	for _, name := range []string{"analyze", "check"} {
		cmd, _, err := root.Find([]string{name})
		if err != nil {
			t.Fatal(err)
		}
		if cmd == nil || cmd.Name() != name {
			t.Fatalf("expected %s command, got %#v", name, cmd)
		}
	}
}

func TestRootCommandIncludesFmtCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"fmt"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "fmt" {
		t.Fatalf("expected fmt command, got %#v", cmd)
	}
	for _, name := range []string{"write", "check", "diff", "stdin", "line-numbers"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected fmt command to define --%s", name)
		}
	}
}

func TestFmtCommandStdinConflictsWithWrite(t *testing.T) {
	a := &app{
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		cwd:            t.TempDir(),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	cmd := a.fmtCommand()
	cmd.SetArgs([]string{"--stdin", "--write"})
	_, err := cmd.ExecuteC()
	if err == nil {
		t.Fatal("expected error for --stdin combined with --write")
	}
	if !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected conflict error, got %v", err)
	}
}

func TestFmtCommandModeConflict(t *testing.T) {
	a := &app{
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		cwd:            t.TempDir(),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	cmd := a.fmtCommand()
	cmd.SetArgs([]string{"--write", "--check"})
	_, err := cmd.ExecuteC()
	if err == nil {
		t.Fatal("expected error for --write combined with --check")
	}
	if !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected conflict error, got %v", err)
	}
}

func TestFmtCommandStdinConflictsWithLineNumbers(t *testing.T) {
	a := &app{
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		cwd:            t.TempDir(),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	cmd := a.fmtCommand()
	cmd.SetArgs([]string{"--stdin", "--line-numbers", "add"})
	_, err := cmd.ExecuteC()
	if err == nil {
		t.Fatal("expected error for --stdin combined with --line-numbers")
	}
	if !strings.Contains(err.Error(), "--line-numbers") {
		t.Fatalf("expected --line-numbers conflict error, got %v", err)
	}
}

func setupFmtProjectDir(t *testing.T, dir string) {
	t.Helper()
	if err := os.WriteFile(filepath.Join(dir, "xlflow.toml"), []byte{}, 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "modules"), 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "classes"), 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "workbook"), 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "tests"), 0755); err != nil {
		t.Fatal(err)
	}
}

func TestFmtWriteViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "test.bas")
	original := "Sub Main()\nx=1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(original), 0644); err != nil {
		t.Fatal(err)
	}
	a := &app{
		cwd:            dir,
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", "--write", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --write error = %v, exit = %d", err, output.ExitCode(err))
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(data) == original {
		t.Fatal("expected file to be written with formatted content")
	}
}

func TestFmtCheckViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "test.bas")
	original := "Sub Main()\nx=1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(original), 0644); err != nil {
		t.Fatal(err)
	}
	a := &app{
		cwd:            dir,
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", "--check", path})
	err := root.Execute()
	if err == nil {
		t.Fatal("expected non-0 exit for unformatted file with --check")
	}
	if code := output.ExitCode(err); code != output.ExitValidation {
		t.Fatalf("expected exit code %d for --check, got %d", output.ExitValidation, code)
	}
}

func TestFmtDiffViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "test.bas")
	original := "Sub Main()\nx=1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(original), 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"fmt", "--diff", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --diff error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	if !strings.Contains(got, "would be reformatted") {
		t.Fatalf("expected diff summary in output:\n%s", got)
	}
	if !strings.Contains(got, "--- a/") {
		t.Fatalf("expected unified diff old-file header in output:\n%s", got)
	}
	if !strings.Contains(got, "+++ b/") {
		t.Fatalf("expected unified diff new-file header in output:\n%s", got)
	}
	if !strings.Contains(got, "@@") {
		t.Fatalf("expected unified diff hunk header in output:\n%s", got)
	}
	if strings.Contains(got, "- +++ b/") || strings.Contains(got, "- @@") {
		t.Fatalf("diff lines should not be list-prefixed:\n%s", got)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(data) != original {
		t.Fatal("--diff mode should not modify files")
	}
}

func TestFmtJSONEnvelopeViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "test.bas")
	formatted := "Sub Main()\n    x = 1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(formatted), 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout.String())
	}
	if env["command"] != "fmt" {
		t.Fatalf("expected command=fmt, got %v", env["command"])
	}
	outputMap, ok := env["output"].(map[string]any)
	if !ok {
		t.Fatal("expected output map in envelope")
	}
	for _, key := range []string{"mode", "changed", "unchanged", "skipped", "total"} {
		if _, ok := outputMap[key]; !ok {
			t.Fatalf("expected output.%s in JSON envelope", key)
		}
	}
	lineNumbers, ok := outputMap["line_numbers"].(map[string]any)
	if !ok {
		t.Fatal("expected output.line_numbers map in JSON envelope")
	}
	if lineNumbers["mode"] != string(vbafmt.LineNumberModePreserve) {
		t.Fatalf("expected default line-number mode preserve, got %v", lineNumbers["mode"])
	}
	if applied, _ := lineNumbers["applied"].(bool); applied {
		t.Fatalf("expected default inspect line-number payload to report applied=false, got %v", lineNumbers["applied"])
	}
	if changed := outputMap["changed"].(float64); changed != 0 {
		t.Fatalf("expected changed=0 for already formatted file, got %v", changed)
	}
}

func TestFmtLineNumbersJSONViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "src", "modules", "Sample.bas")
	input := "Public Sub Sample()\n    Dim x As Integer\n    x = 1 / 0\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(input), 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", "--line-numbers", "add", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --line-numbers add --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout.String())
	}
	outputMap, ok := env["output"].(map[string]any)
	if !ok {
		t.Fatal("expected output map in envelope")
	}
	lineNumbers, ok := outputMap["line_numbers"].(map[string]any)
	if !ok {
		t.Fatal("expected output.line_numbers map in envelope")
	}
	if lineNumbers["mode"] != string(vbafmt.LineNumberModeAdd) {
		t.Fatalf("expected mode add, got %v", lineNumbers["mode"])
	}
	if applied, _ := lineNumbers["applied"].(bool); applied {
		t.Fatalf("expected inspect line-number payload to report applied=false, got %v", lineNumbers["applied"])
	}
	if got := lineNumbers["files_to_change"].(float64); got != 1 {
		t.Fatalf("expected files_to_change=1, got %v", got)
	}
	if got := lineNumbers["lines_to_add"].(float64); got != 2 {
		t.Fatalf("expected lines_to_add=2, got %v", got)
	}
}

func TestFmtLineNumbersWriteJSONViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "src", "modules", "Sample.bas")
	input := "Public Sub Sample()\n    Dim x As Integer\n    x = 1 / 0\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(input), 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", "--line-numbers", "add", "--write", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --line-numbers add --write --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout.String())
	}
	outputMap, ok := env["output"].(map[string]any)
	if !ok {
		t.Fatal("expected output map in envelope")
	}
	lineNumbers, ok := outputMap["line_numbers"].(map[string]any)
	if !ok {
		t.Fatal("expected output.line_numbers map in envelope")
	}
	if applied, _ := lineNumbers["applied"].(bool); !applied {
		t.Fatalf("expected write line-number payload to report applied=true, got %v", lineNumbers["applied"])
	}
	if got := lineNumbers["files_changed"].(float64); got != 1 {
		t.Fatalf("expected files_changed=1, got %v", got)
	}
	if got := lineNumbers["lines_added"].(float64); got != 2 {
		t.Fatalf("expected lines_added=2, got %v", got)
	}
	if _, ok := lineNumbers["lines_to_add"]; ok {
		t.Fatalf("did not expect prospective inspect key lines_to_add in write mode: %#v", lineNumbers)
	}
}

func TestFmtLineNumbersWarningsUsePathInJSONViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "src", "modules", "Sample.bas")
	input := "10  ' legacy comment\nPublic Sub Sample()\n    x = 1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(input), 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", "--line-numbers", "add", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --line-numbers warning --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout.String())
	}
	outputMap := env["output"].(map[string]any)
	lineNumbers := outputMap["line_numbers"].(map[string]any)
	warnings, ok := lineNumbers["warnings"].([]any)
	if !ok || len(warnings) == 0 {
		t.Fatalf("expected line_numbers warnings, got %#v", lineNumbers["warnings"])
	}
	warning := warnings[0].(map[string]any)
	if warning["path"] == nil || warning["path"] == "" {
		t.Fatalf("expected warning.path, got %#v", warning)
	}
	if _, exists := warning["file"]; exists {
		t.Fatalf("did not expect warning.file key, got %#v", warning)
	}
}

func TestFmtJSONTargetContractDefaultScope(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	a := &app{
		cwd:            dir,
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt"})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(a.stdout.(*bytes.Buffer).Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v", err)
	}
	target, ok := env["target"].(map[string]any)
	if !ok {
		t.Fatal("expected target in JSON envelope")
	}
	if target["kind"] != "source" {
		t.Fatalf("expected target.kind=source, got %v", target["kind"])
	}
	path, ok := target["path"].(string)
	if !ok {
		t.Fatal("expected target.path string")
	}
	if !strings.Contains(path, "tests") {
		t.Fatalf("expected default target.path to include tests/, got %q", path)
	}
	if strings.Contains(path, "src") {
		// ok: default scope includes src/*
	} else {
		t.Fatalf("expected default target.path to include src dirs, got %q", path)
	}
}

func TestFmtJSONTargetContractExplicitPath(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "src", "modules", "test.bas")
	formatted := "Sub Main()\n    x = 1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(formatted), 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v", err)
	}
	target, ok := env["target"].(map[string]any)
	if !ok {
		t.Fatal("expected target in JSON envelope")
	}
	if target["kind"] != "source" {
		t.Fatalf("expected target.kind=source, got %v", target["kind"])
	}
	pathStr, ok := target["path"].(string)
	if !ok {
		t.Fatal("expected target.path string")
	}
	if pathStr == "" {
		t.Fatal("expected non-empty target.path for explicit path")
	}
}

func TestFmtStdinViaCLI(t *testing.T) {
	input := "Sub Main()\nx=1\nEnd Sub\n"
	stdinReader, stdinWriter, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	go func() {
		_, _ = stdinWriter.Write([]byte(input))
		_ = stdinWriter.Close()
	}()
	dir := t.TempDir()
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	oldStdin := os.Stdin
	os.Stdin = stdinReader
	defer func() { os.Stdin = oldStdin }()
	root := a.rootCommand()
	root.SetArgs([]string{"fmt", "--stdin"})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --stdin error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	if strings.TrimSpace(got) == "" {
		t.Fatal("expected non-empty formatted output from --stdin")
	}
}

func TestFmtStdinRejectsPathArgsViaCLI(t *testing.T) {
	dir := t.TempDir()
	a := &app{
		cwd:            dir,
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"fmt", "--stdin", "src/modules/Test.bas"})
	err := root.Execute()
	if err == nil {
		t.Fatal("expected error for --stdin combined with path args")
	}
	if !strings.Contains(err.Error(), "path arguments") {
		t.Fatalf("expected path-argument validation error, got %v", err)
	}
}

func TestFmtStdinDetectsClassModuleViaCLI(t *testing.T) {
	input := "VERSION 1.0 CLASS\nBEGIN\nMultiUse = -1\nEND\nAttribute VB_Name = \"Test\"\nOption Explicit\nPublic Sub Run()\nx=1\nEnd Sub\n"
	stdinReader, stdinWriter, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	go func() {
		_, _ = stdinWriter.Write([]byte(input))
		_ = stdinWriter.Close()
	}()
	dir := t.TempDir()
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	oldStdin := os.Stdin
	os.Stdin = stdinReader
	defer func() { os.Stdin = oldStdin }()
	root := a.rootCommand()
	root.SetArgs([]string{"fmt", "--stdin"})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --stdin error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	if !strings.Contains(got, "VERSION 1.0 CLASS") || !strings.Contains(got, `Attribute VB_Name = "Test"`) {
		t.Fatalf("expected class header to be preserved for stdin class input:\n%s", got)
	}
}

func TestFmtJSONWarningsUseSkipReasonViaCLI(t *testing.T) {
	dir := t.TempDir()
	setupFmtProjectDir(t, dir)
	path := filepath.Join(dir, "UserForm1.frm")
	if err := os.WriteFile(path, []byte("VERSION 5.00\n"), 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", path})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout.String())
	}
	warnings, ok := env["warnings"].([]any)
	if !ok || len(warnings) != 1 {
		t.Fatalf("expected one warning, got %v", env["warnings"])
	}
	warning, ok := warnings[0].(map[string]any)
	if !ok {
		t.Fatalf("expected warning map, got %T", warnings[0])
	}
	if warning["code"] != "unsupported extension: .frm" {
		t.Fatalf("expected warning code to use skip reason, got %v", warning["code"])
	}
}

func TestFmtStdinJSONViaCLI(t *testing.T) {
	input := "Sub Main()\nx=1\nEnd Sub\n"
	stdinReader, stdinWriter, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	go func() {
		_, _ = stdinWriter.Write([]byte(input))
		_ = stdinWriter.Close()
	}()
	dir := t.TempDir()
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	oldStdin := os.Stdin
	os.Stdin = stdinReader
	defer func() { os.Stdin = oldStdin }()
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt", "--stdin"})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt --stdin --json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("failed to parse JSON output: %v\nraw: %s", err, stdout.String())
	}
	if env.Command != "fmt" {
		t.Fatalf("expected command 'fmt', got %q", env.Command)
	}
	if env.Status != "ok" {
		t.Fatalf("expected status 'ok', got %q", env.Status)
	}
	if env.Output == nil {
		t.Fatal("expected non-nil output field in JSON envelope")
	}
	outMap, ok := env.Output.(map[string]any)
	if !ok {
		t.Fatalf("expected output to be a map, got %T", env.Output)
	}
	if mode, _ := outMap["mode"].(string); mode != "inspect" {
		t.Fatalf("expected mode 'inspect', got %q", mode)
	}
	if changed, _ := outMap["changed"].(float64); changed != 1 {
		t.Fatalf("expected changed=1, got %v", changed)
	}
	if total, _ := outMap["total"].(float64); total != 1 {
		t.Fatalf("expected total=1, got %v", total)
	}
}

func TestFmtDiffCheckModeConflictViaCLI(t *testing.T) {
	dir := t.TempDir()
	a := &app{
		cwd:            dir,
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"fmt", "--diff", "--check"})
	err := root.Execute()
	if err == nil {
		t.Fatal("expected error for --diff combined with --check")
	}
	if !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected conflict error, got %v", err)
	}
}

func TestFmtEmptyProjectViaCLI(t *testing.T) {
	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, "xlflow.toml"), []byte{}, 0644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         new(bytes.Buffer),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "fmt"})
	if err := root.Execute(); err != nil {
		t.Fatalf("fmt in empty project error = %v, exit = %d", err, output.ExitCode(err))
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout.String())
	}
	outputMap, ok := env["output"].(map[string]any)
	if !ok {
		t.Fatal("expected output map in envelope for empty project")
	}
	if total, ok := outputMap["total"]; !ok {
		t.Fatal("expected output.total in empty project envelope")
	} else {
		totalFloat, _ := total.(float64)
		if totalFloat != 0 {
			t.Fatalf("expected total=0 for empty project, got %v", total)
		}
	}
}

func TestRootCommandIncludesMacrosCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"macros"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "macros" {
		t.Fatalf("expected macros command, got %#v", cmd)
	}
	if cmd.Flags().Lookup("session") == nil {
		t.Fatal("expected macros command to define --session")
	}
}

func TestRootCommandIncludesSessionFlagsForWorkbookReaders(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	for _, args := range [][]string{
		{"list", "forms"},
		{"pull"},
		{"export-image"},
		{"inspect", "workbook"},
		{"inspect", "sheets"},
		{"inspect", "range"},
		{"inspect", "used-range"},
		{"inspect", "cell"},
		{"test"},
	} {
		cmd, _, err := root.Find(args)
		if err != nil {
			t.Fatal(err)
		}
		if cmd.Flags().Lookup("session") == nil {
			t.Fatalf("expected %v command to define --session", args)
		}
	}
}

func TestRootCommandIncludesListFormsCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"list", "forms"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "forms" {
		t.Fatalf("expected list forms command, got %#v", cmd)
	}
	for _, name := range []string{"session"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected list forms command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesFormSnapshotCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"form", "snapshot"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "snapshot" {
		t.Fatalf("expected form snapshot command, got %#v", cmd)
	}
	for _, name := range []string{"out", "session"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected form snapshot command to define --%s", name)
		}
	}
	if cmd.Flags().Lookup("designer") != nil {
		t.Fatal("form snapshot command should not expose --designer")
	}
}

func TestRootCommandIncludesFormBuildCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"form", "build"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "build" {
		t.Fatalf("expected form build command, got %#v", cmd)
	}
	for _, name := range []string{"overwrite", "session", "no-save"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected form build command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesFormApplyCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"form", "apply"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "apply" {
		t.Fatalf("expected form apply command, got %#v", cmd)
	}
	if !cmd.Hidden {
		t.Fatal("expected form apply command to be hidden")
	}
	for _, name := range []string{"session", "no-save"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected form apply command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesFormExportImageCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"form", "export-image"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "export-image" {
		t.Fatalf("expected form export-image command, got %#v", cmd)
	}
	for _, name := range []string{"out", "initializer", "overwrite", "session"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected form export-image command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesInspectFormCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"inspect", "form"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "form" {
		t.Fatalf("expected inspect form command, got %#v", cmd)
	}
	for _, name := range []string{"runtime", "designer", "both", "initializer", "session"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected inspect form command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesExportImageCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"export-image"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "export-image" {
		t.Fatalf("expected export-image command, got %#v", cmd)
	}
	for _, name := range []string{"sheet", "range", "out", "output-dir", "name", "format", "overwrite", "session"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected export-image command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesEditCommands(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	for _, args := range [][]string{
		{"edit", "cell"},
		{"edit", "range"},
		{"edit", "rows"},
		{"edit", "columns"},
	} {
		cmd, _, err := root.Find(args)
		if err != nil {
			t.Fatal(err)
		}
		if cmd == nil || cmd.Name() != args[len(args)-1] {
			t.Fatalf("expected command %v, got %#v", args, cmd)
		}
	}

	cell, _, err := root.Find([]string{"edit", "cell"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"sheet", "cell", "value", "formula", "fill", "events", "session"} {
		if cell.Flags().Lookup(name) == nil {
			t.Fatalf("expected edit cell to define --%s", name)
		}
	}
}

func TestBuildExportImageOptionsValidatesAndNormalizes(t *testing.T) {
	opts, err := buildExportImageOptions(" build\\Book.xlsm ", " QR ", "ae31:a1", "", " artifacts\\images ", " qr-demo ", " png ", true, true, excel.CommandOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if opts.WorkbookPath != "build\\Book.xlsm" {
		t.Fatalf("workbook = %q", opts.WorkbookPath)
	}
	if opts.Sheet != "QR" {
		t.Fatalf("sheet = %q", opts.Sheet)
	}
	if opts.Range != "A1:AE31" {
		t.Fatalf("range = %q, want A1:AE31", opts.Range)
	}
	if opts.OutputDir != "artifacts\\images" || opts.Name != "qr-demo" || opts.Format != "png" {
		t.Fatalf("unexpected output opts: %#v", opts)
	}
	if !opts.Overwrite || !opts.Session {
		t.Fatalf("expected overwrite/session to be preserved: %#v", opts)
	}
}

func TestBuildExportImageOptionsRejectsInvalidCombinations(t *testing.T) {
	_, err := buildExportImageOptions("", "", "A1:B2", "", "", "", "png", false, false, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "--sheet is required") {
		t.Fatalf("expected missing sheet error, got %v", err)
	}
	_, err = buildExportImageOptions("", "QR", "bad", "", "", "", "png", false, false, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "--range") {
		t.Fatalf("expected range validation error, got %v", err)
	}
	_, err = buildExportImageOptions("", "QR", "A1:B2", "out.png", "artifacts", "", "png", false, false, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "--out cannot be combined") {
		t.Fatalf("expected out/output-dir conflict, got %v", err)
	}
	_, err = buildExportImageOptions("", "QR", "A1:B2", "", "", "nested\\qr.png", "png", false, false, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "--name must be a filename") {
		t.Fatalf("expected name validation error, got %v", err)
	}
	_, err = buildExportImageOptions("", "QR", "A1:B2", "", "", "qr:demo", "png", false, false, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "invalid Windows characters") {
		t.Fatalf("expected invalid Windows character error, got %v", err)
	}
	_, err = buildExportImageOptions("", "QR", "A1:B2", "", "", "qr\x1fdemo", "png", false, false, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "invalid Windows characters") {
		t.Fatalf("expected control-character validation error, got %v", err)
	}
}

func TestBuildFormSnapshotOptionsValidatesAndNormalizes(t *testing.T) {
	stderr := bytes.Buffer{}
	opts, err := buildFormSnapshotOptions(" UserForm1 ", " artifacts\\UserForm1.form.yaml ", true, excel.CommandOptions{Stderr: &stderr})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Inspect.Name != "UserForm1" || opts.Inspect.Basis != "designer" {
		t.Fatalf("inspect opts = %#v", opts.Inspect)
	}
	if opts.Inspect.StrictDesigner {
		t.Fatalf("form snapshot should use non-executing designer opts = %#v", opts.Inspect)
	}
	if !opts.Inspect.Session || opts.Inspect.Keepalive.Stderr != &stderr {
		t.Fatalf("command/session opts = %#v", opts.Inspect)
	}
	if opts.OutPath != "artifacts\\UserForm1.form.yaml" {
		t.Fatalf("out path = %q", opts.OutPath)
	}
}

func TestBuildFormSnapshotOptionsRejectsMissingRequirements(t *testing.T) {
	if _, err := buildFormSnapshotOptions("UserForm1", "", false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--out is required") {
		t.Fatalf("expected out requirement error, got %v", err)
	}
	if _, err := buildFormSnapshotOptions("", "artifacts\\UserForm1.form.yaml", false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "form name is required") {
		t.Fatalf("expected form name error, got %v", err)
	}
}

func TestBuildFormExportImageOptionsValidatesAndNormalizes(t *testing.T) {
	stderr := bytes.Buffer{}
	opts, err := buildFormExportImageOptions(" UserForm1 ", " artifacts\\UserForm1.png ", " InitializeForm ", true, true, excel.CommandOptions{Stderr: &stderr})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Name != "UserForm1" || opts.OutPath != "artifacts\\UserForm1.png" || opts.Initializer != "InitializeForm" {
		t.Fatalf("unexpected form export-image opts: %#v", opts)
	}
	if !opts.Overwrite || !opts.Session || opts.Keepalive.Stderr != &stderr {
		t.Fatalf("unexpected command/session opts: %#v", opts)
	}
}

func TestBuildFormExportImageOptionsRejectsMissingRequirements(t *testing.T) {
	if _, err := buildFormExportImageOptions("", "artifacts\\UserForm1.png", "", false, false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "form name is required") {
		t.Fatalf("expected form name error, got %v", err)
	}
	if _, err := buildFormExportImageOptions("UserForm1", "", "", false, false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--out is required") {
		t.Fatalf("expected out requirement error, got %v", err)
	}
}

func TestBuildFormWriteOptionsValidatesAndNormalizes(t *testing.T) {
	root := t.TempDir()
	specPath := filepath.Join(root, "specs", "UserForm1.form.yaml")
	if err := os.MkdirAll(filepath.Dir(specPath), 0o755); err != nil {
		t.Fatal(err)
	}
	body := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: UserForm1\ncontrols:\n  - name: txtCustomer\n    type: TextBox\nwarnings: []\n"
	if err := os.WriteFile(specPath, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	stderr := bytes.Buffer{}
	opts, err := buildFormWriteOptions(" build ", specPath, true, true, true, excel.CommandOptions{Stderr: &stderr}, root)
	if err != nil {
		t.Fatal(err)
	}
	if opts.Action != "build" || opts.Spec.Form.Name != "UserForm1" {
		t.Fatalf("unexpected form write opts: %#v", opts)
	}
	if !opts.Overwrite || !opts.Session || !opts.NoSave {
		t.Fatalf("expected overwrite/session/no-save to be preserved: %#v", opts)
	}
	if opts.Keepalive.Stderr != &stderr {
		t.Fatalf("unexpected command opts: %#v", opts.Keepalive)
	}
	if !strings.HasSuffix(opts.SpecInput.DisplayPath, "specs/UserForm1.form.yaml") {
		t.Fatalf("unexpected display path: %q", opts.SpecInput.DisplayPath)
	}
}

func TestBuildFormWriteOptionsRejectsInvalidRequirements(t *testing.T) {
	root := t.TempDir()
	if _, err := buildFormWriteOptions("apply", "missing.form.yaml", false, false, true, excel.CommandOptions{}, root); err == nil || !strings.Contains(err.Error(), "--no-save requires --session") {
		t.Fatalf("expected no-save/session error, got %v", err)
	}
	if _, err := buildFormWriteOptions("noop", "missing.form.yaml", false, false, false, excel.CommandOptions{}, root); err == nil || !strings.Contains(err.Error(), "unsupported form action") {
		t.Fatalf("expected action error, got %v", err)
	}
}

func TestBuildEditCellOptionsValidatesAndNormalizes(t *testing.T) {
	value := "ABC123"
	opts, err := buildEditCellOptions(" build\\Book.xlsm ", " Input ", " b2 ", "", " on ", &value, nil, true, excel.CommandOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if opts.WorkbookPath != "build\\Book.xlsm" || opts.Sheet != "Input" || opts.Cell != "B2" {
		t.Fatalf("unexpected normalized edit cell opts: %#v", opts)
	}
	if opts.Value == nil || *opts.Value != "ABC123" {
		t.Fatalf("value = %#v", opts.Value)
	}
	if opts.Events != excel.EditEventOn || !opts.Session {
		t.Fatalf("unexpected event/session opts: %#v", opts)
	}
}

func TestBuildEditCellOptionsRejectsInvalidCombinations(t *testing.T) {
	value := "ABC123"
	formula := "=A1+B1"
	if _, err := buildEditCellOptions("", "Input", "B2", "", "keep", &value, nil, false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--session") {
		t.Fatalf("expected session requirement error, got %v", err)
	}
	if _, err := buildEditCellOptions("", "Input", "B2", "", "keep", nil, nil, true, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "exactly one") {
		t.Fatalf("expected missing mutation error, got %v", err)
	}
	if _, err := buildEditCellOptions("", "Input", "B2", "", "keep", &value, &formula, true, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "exactly one") {
		t.Fatalf("expected multi mutation error, got %v", err)
	}
	if _, err := buildEditCellOptions("", "Input", "B2", "#12", "keep", nil, nil, true, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "#RGB") {
		t.Fatalf("expected invalid color error, got %v", err)
	}
	if _, err := buildEditCellOptions("", "Input", "B2", "#fff", "off", nil, nil, true, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--events applies") {
		t.Fatalf("expected fill/events conflict, got %v", err)
	}
}

func TestBuildEditRangeOptionsValidatesAndNormalizes(t *testing.T) {
	opts, err := buildEditRangeOptions(" build\\Book.xlsm ", " QR ", "ae31:a1", "#fff", "", true, excel.CommandOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if opts.WorkbookPath != "build\\Book.xlsm" || opts.Sheet != "QR" || opts.Range != "A1:AE31" || opts.Fill != "#FFFFFF" {
		t.Fatalf("unexpected edit range opts: %#v", opts)
	}
	if _, err := buildEditRangeOptions("", "QR", "A1:B2", "#fff", "", false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--session") {
		t.Fatalf("expected session requirement error, got %v", err)
	}
}

func TestBuildEditRowsAndColumnsOptionsValidateSelectors(t *testing.T) {
	rows, err := buildEditRowsOptions("", "QR", "31:1", 12, true, excel.CommandOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if rows.Rows != "1:31" {
		t.Fatalf("rows selector = %q", rows.Rows)
	}
	columns, err := buildEditColumnsOptions("", "QR", "ae:a", 2.2, true, excel.CommandOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if columns.Columns != "A:AE" {
		t.Fatalf("columns selector = %q", columns.Columns)
	}
	if _, err := buildEditRowsOptions("", "QR", "1:31", 12, false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--session") {
		t.Fatalf("expected session requirement error, got %v", err)
	}
	if _, err := buildEditColumnsOptions("", "QR", "A:AE", 2.2, false, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--session") {
		t.Fatalf("expected session requirement error, got %v", err)
	}
	if _, err := buildEditRowsOptions("", "QR", "0", 12, true, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--rows") {
		t.Fatalf("expected invalid rows error, got %v", err)
	}
	if _, err := buildEditColumnsOptions("", "QR", "A:3", 2.2, true, excel.CommandOptions{}); err == nil || !strings.Contains(err.Error(), "--columns") {
		t.Fatalf("expected invalid columns error, got %v", err)
	}
}

func TestRootCommandIncludesUIButtonCommands(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	for _, args := range [][]string{
		{"ui", "button", "add"},
		{"ui", "button", "list"},
		{"ui", "button", "remove"},
	} {
		cmd, _, err := root.Find(args)
		if err != nil {
			t.Fatal(err)
		}
		if cmd == nil || cmd.Name() != args[len(args)-1] {
			t.Fatalf("expected command %v, got %#v", args, cmd)
		}
	}

	add, _, err := root.Find([]string{"ui", "button", "add"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"sheet", "cell", "text", "macro", "id", "width", "height", "create-sheet", "verify-macro", "session"} {
		if add.Flags().Lookup(name) == nil {
			t.Fatalf("expected ui button add to define --%s", name)
		}
	}

	list, _, err := root.Find([]string{"ui", "button", "list"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"sheet", "session"} {
		if list.Flags().Lookup(name) == nil {
			t.Fatalf("expected ui button list to define --%s", name)
		}
	}

	remove, _, err := root.Find([]string{"ui", "button", "remove"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"id", "sheet", "session"} {
		if remove.Flags().Lookup(name) == nil {
			t.Fatalf("expected ui button remove to define --%s", name)
		}
	}
}

func TestBuildUIButtonAddOptionsDefaultsAndNormalizesID(t *testing.T) {
	opts, err := buildUIButtonAddOptions(excel.UIButtonAddOptions{
		Sheet:       " Menu ",
		Cell:        " B2 ",
		Text:        " Run ",
		Macro:       " Main.RunAggregation ",
		Width:       160,
		Height:      40,
		CreateSheet: true,
		VerifyMacro: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if opts.ID != "main-runaggregation" {
		t.Fatalf("id = %q, want main-runaggregation", opts.ID)
	}
	if opts.Sheet != "Menu" || opts.Cell != "B2" || opts.Text != "Run" || opts.Macro != "Main.RunAggregation" {
		t.Fatalf("unexpected trimmed opts: %#v", opts)
	}
	if !opts.CreateSheet || !opts.VerifyMacro {
		t.Fatalf("expected boolean flags to be preserved: %#v", opts)
	}
}

func TestBuildUIButtonAddOptionsValidatesRequiredFields(t *testing.T) {
	_, err := buildUIButtonAddOptions(excel.UIButtonAddOptions{Sheet: "Menu", Cell: "B2", Text: "Run", Width: 160, Height: 40})
	if err == nil || !strings.Contains(err.Error(), "--macro is required") {
		t.Fatalf("expected macro required error, got %v", err)
	}
	_, err = buildUIButtonAddOptions(excel.UIButtonAddOptions{Sheet: "Menu", Cell: "B2", Text: "Run", Macro: "Main.Run", Width: 0, Height: 40})
	if err == nil || !strings.Contains(err.Error(), "--width") {
		t.Fatalf("expected width error, got %v", err)
	}
}

func TestBuildUIButtonAddOptionsPreservesSessionFlag(t *testing.T) {
	opts, err := buildUIButtonAddOptions(excel.UIButtonAddOptions{
		Sheet:   "Menu",
		Cell:    "B2",
		Text:    "Run",
		Macro:   "Main.Run",
		Width:   160,
		Height:  40,
		Session: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if !opts.Session {
		t.Fatalf("expected Session to be preserved as true, got false")
	}
}

func TestBuildUIButtonAddOptionsDefaultsSessionFalse(t *testing.T) {
	opts, err := buildUIButtonAddOptions(excel.UIButtonAddOptions{
		Sheet:  "Menu",
		Cell:   "B2",
		Text:   "Run",
		Macro:  "Main.Run",
		Width:  160,
		Height: 40,
	})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Session {
		t.Fatalf("expected Session to default to false, got true")
	}
}

func TestBuildUIButtonRemoveOptionsPreservesSessionFlag(t *testing.T) {
	opts, err := buildUIButtonRemoveOptions(excel.UIButtonRemoveOptions{
		ID:      "mybutton",
		Session: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	if !opts.Session {
		t.Fatalf("expected Session to be preserved as true, got false")
	}
}

func TestBuildUIButtonRemoveOptionsDefaultsSessionFalse(t *testing.T) {
	opts, err := buildUIButtonRemoveOptions(excel.UIButtonRemoveOptions{
		ID: "mybutton",
	})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Session {
		t.Fatalf("expected Session to default to false, got true")
	}
}

func TestRootCommandIncludesDiffCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"diff"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "diff" {
		t.Fatalf("expected diff command, got %#v", cmd)
	}
	for _, name := range []string{"vba-before", "vba-after"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected diff command to define --%s", name)
		}
	}
}

func TestRootCommandIncludesSkillInstallCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"skill", "install"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "install" {
		t.Fatalf("expected skill install command, got %#v", cmd)
	}
	for _, name := range []string{"agent", "target", "force"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected skill install command to define --%s", name)
		}
	}
	if usage := cmd.Flags().Lookup("agent").Usage; strings.Contains(usage, "copilot") {
		t.Fatalf("agent flag usage should not mention copilot: %q", usage)
	}
}

func TestRootCommandIncludesModuleInstallCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"module", "install"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "install" {
		t.Fatalf("expected module install command, got %#v", cmd)
	}
	if cmd.Flags().Lookup("push") == nil {
		t.Fatal("expected module install command to define --push")
	}
}

func TestNewAndInitIncludeWithSkillFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	for _, command := range []string{"new", "init"} {
		cmd, _, err := root.Find([]string{command})
		if err != nil {
			t.Fatal(err)
		}
		for _, name := range []string{"with-skill", "agent"} {
			if cmd.Flags().Lookup(name) == nil {
				t.Fatalf("expected %s command to define --%s", command, name)
			}
		}
	}
}

func TestBuildStatusWarningsAndHintsSessionDirty(t *testing.T) {
	session := map[string]any{
		"active":               true,
		"save_required":        true,
		"live_newer_than_disk": true,
	}
	state := map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": true,
	}
	warnings, hints := buildStatusWarningsAndHints(session, state)

	foundDirty := false
	foundLiveNewer := false
	for _, w := range warnings {
		switch w["code"] {
		case "session_dirty":
			foundDirty = true
		case "live_session_newer_than_disk":
			foundLiveNewer = true
		}
	}
	if !foundDirty {
		t.Fatal("expected session_dirty warning")
	}
	if !foundLiveNewer {
		t.Fatal("expected live_session_newer_than_disk warning")
	}

	foundSave := false
	foundSaveBeforePush := false
	for _, h := range hints {
		switch h["code"] {
		case "save_session":
			foundSave = true
		case "save_before_push":
			foundSaveBeforePush = true
		}
	}
	if !foundSave {
		t.Fatal("expected save_session hint")
	}
	if !foundSaveBeforePush {
		t.Fatal("expected save_before_push hint")
	}
}

func TestBuildStatusWarningsAndHintsSourceNewer(t *testing.T) {
	session := map[string]any{"active": false}
	state := map[string]any{
		"src_newer_than_workbook":      true,
		"live_session_newer_than_disk": false,
	}
	warnings, hints := buildStatusWarningsAndHints(session, state)

	found := false
	for _, w := range warnings {
		if w["code"] == "source_newer_than_workbook" {
			found = true
		}
	}
	if !found {
		t.Fatal("expected source_newer_than_workbook warning")
	}

	found = false
	for _, h := range hints {
		if h["code"] == "push_source" {
			found = true
		}
	}
	if !found {
		t.Fatal("expected push_source hint")
	}
}

func TestBuildStatusWarningsAndHintsSessionInactive(t *testing.T) {
	session := map[string]any{"active": false}
	state := map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": false,
	}
	warnings, hints := buildStatusWarningsAndHints(session, state)

	if len(warnings) != 0 {
		t.Fatalf("expected no warnings for inactive session, got %v", warnings)
	}
	if len(hints) != 0 {
		t.Fatalf("expected no hints for inactive session, got %v", hints)
	}
}

func TestInitCommandIncludesWithModuleFlag(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	cmd, _, err := root.Find([]string{"init"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd.Flags().Lookup("with-module") == nil {
		t.Fatal("expected init command to define --with-module")
	}
}

func TestTestCommandIncludesModuleAndTagFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	cmd, _, err := root.Find([]string{"test"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd.Flags().Lookup("module") == nil {
		t.Fatal("expected test command to define --module")
	}
	if cmd.Flags().Lookup("tag") == nil {
		t.Fatal("expected test command to define --tag")
	}
}

func TestSkillInstallCommandInstallsProviderSkill(t *testing.T) {
	dir := t.TempDir()
	a := &app{cwd: dir, bridge: "powershell"}
	root := a.rootCommand()
	root.SetArgs([]string{"skill", "install", "--agent", "codex"})
	if err := root.Execute(); err != nil {
		t.Fatalf("skill install command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if _, err := os.Stat(filepath.Join(dir, ".codex", "skills", "xlflow", "SKILL.md")); err != nil {
		t.Fatal(err)
	}
}

func TestInitWithSkillInstallsProviderSkill(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook, "--with-skill", "--agent", "codex"))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if _, err := os.Stat(filepath.Join(dir, ".codex", "skills", "xlflow", "SKILL.md")); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(dir, "prompts", "agent.md")); !os.IsNotExist(err) {
		t.Fatalf("expected prompts/agent.md not to be scaffolded, got %v", err)
	}
}

func TestInitCommandRendersWelcomeForInteractiveTerminal(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
		buildInfo:      BuildInfo{Version: "1.2.3"},
		updateChecker:  stubReleaseChecker{},
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	for _, want := range []string{"Welcome to", "Docs: https://harumiweb.github.io/xlflow/commands/", "Version: 1.2.3", "copied workbook to build/Input.xlsm", "pulled workbook VBA into source"} {
		if !strings.Contains(got, want) {
			t.Fatalf("interactive init output missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "+-") {
		t.Fatalf("interactive init output should not include badge borders:\n%s", got)
	}
	if strings.Index(got, "Welcome to") >= strings.Index(got, " ██╗  ██╗ ██╗      ███████╗ ██╗       ██████╗  ██╗    ██╗") ||
		strings.Index(got, "Docs: https://harumiweb.github.io/xlflow/commands/") >= strings.Index(got, "Version: 1.2.3") {
		t.Fatalf("expected welcome heading and meta order before command summary:\n%s", got)
	}
	if strings.Index(got, "Version: 1.2.3") > strings.Index(got, "xlflow init") {
		t.Fatalf("expected welcome UI before command summary:\n%s", got)
	}
}

func TestInitCommandSkipsWelcomeForJSONOutput(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("--json", "init", workbook))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if strings.Contains(stdout.String(), "Welcome to xlflow") {
		t.Fatalf("json output should not include welcome UI:\n%s", stdout.String())
	}
	var env map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("json output should remain valid: %v\n%s", err, stdout.String())
	}
	if env["command"] != "init" {
		t.Fatalf("json command = %#v, want init", env["command"])
	}
}

func TestInitCommandShowsUpdateNoticeWhenNewReleaseIsAvailable(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
		buildInfo:      BuildInfo{Version: "1.2.3"},
		updateChecker:  stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	if !strings.Contains(got, "Update available: v1.2.4") {
		t.Fatalf("interactive init output missing update notice:\n%s", got)
	}
}

func TestInitCommandSilentlySkipsFailedUpdateCheck(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
		buildInfo:      BuildInfo{Version: "1.2.3"},
		updateChecker:  stubReleaseChecker{err: errors.New("network down")},
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	if strings.Contains(got, "Update available:") {
		t.Fatalf("interactive init output should skip failed update checks:\n%s", got)
	}
}

func TestInitCommandSkipsUpdateCheckWithFlag(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
		buildInfo:      BuildInfo{Version: "1.2.3"},
		updateChecker:  stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook, "--no-update-check"))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if strings.Contains(stdout.String(), "Update available:") {
		t.Fatalf("interactive init output should skip update notice when --no-update-check is set:\n%s", stdout.String())
	}
}

func TestInitCommandSkipsUpdateCheckWithEnv(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	t.Setenv(noUpdateCheckEnvVar, "1")

	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
		buildInfo:      BuildInfo{Version: "1.2.3"},
		updateChecker:  stubReleaseChecker{release: latestRelease{Version: "v1.2.4"}},
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if strings.Contains(stdout.String(), "Update available:") {
		t.Fatalf("interactive init output should skip update notice when %s is set:\n%s", noUpdateCheckEnvVar, stdout.String())
	}
}

func TestInitCommandAutoPullsWorkbookSource(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, true)

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}

	modulePath := filepath.Join(dir, "src", "modules", "Imported.bas")
	body, err := os.ReadFile(modulePath)
	if err != nil {
		t.Fatalf("expected auto-pulled module at %s: %v", modulePath, err)
	}
	if !strings.Contains(string(body), `Attribute VB_Name = "Imported"`) {
		t.Fatalf("unexpected pulled module body:\n%s", string(body))
	}
}

func TestInitCommandWithModuleAutoPushesHelperSource(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestPullScript(t, dir, false)
	writeTestPushScript(t, dir)

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("init", workbook, "--with-module"))
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}

	for _, path := range []string{
		filepath.Join(dir, "src", "modules", "XlflowAssert.bas"),
		filepath.Join(dir, "src", "modules", "XlflowRuntime.bas"),
		filepath.Join(dir, "src", "modules", "XlflowUI.bas"),
		filepath.Join(dir, "src", "modules", "XlflowDebug.bas"),
	} {
		if _, err := os.Stat(path); err != nil {
			t.Fatalf("expected helper module at %s: %v", path, err)
		}
	}
	markerPath := filepath.Join(dir, ".xlflow", "push.called")
	body, err := os.ReadFile(markerPath)
	if err != nil {
		t.Fatalf("expected auto-push marker at %s: %v", markerPath, err)
	}
	text := string(body)
	for _, want := range []string{"XlflowAssert.bas", "XlflowRuntime.bas", "XlflowUI.bas", "XlflowDebug.bas"} {
		if !strings.Contains(text, want) {
			t.Fatalf("expected push marker to mention %s:\n%s", want, text)
		}
	}
}

func TestNewCommandAutoPushesScaffoldSource(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	writeTestNewScript(t, dir)
	writeTestPushScript(t, dir)

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--bridge", "powershell", "new", "Book.xlsm"})
	if err := root.Execute(); err != nil {
		t.Fatalf("new command error = %v, exit = %d", err, output.ExitCode(err))
	}

	markerPath := filepath.Join(dir, ".xlflow", "push.called")
	body, err := os.ReadFile(markerPath)
	if err != nil {
		t.Fatalf("expected auto-push marker at %s: %v", markerPath, err)
	}
	text := string(body)
	for _, want := range []string{"Main.bas", "App.bas", "Ui.bas", "XlflowAssert.bas"} {
		if !strings.Contains(text, want) {
			t.Fatalf("expected push marker to mention %s:\n%s", want, text)
		}
	}
}

func TestNewCommandRendersWelcomeForInteractiveTerminal(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	writeTestNewScript(t, dir)
	writeTestPushScript(t, dir)
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
		buildInfo:      BuildInfo{Version: "1.2.3"},
		updateChecker:  stubReleaseChecker{},
	}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("new", "Book.xlsm"))
	if err := root.Execute(); err != nil {
		t.Fatalf("new command error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	for _, want := range []string{"Welcome to", "Docs: https://harumiweb.github.io/xlflow/commands/", "Version: 1.2.3", "created xlflow.toml"} {
		if !strings.Contains(got, want) {
			t.Fatalf("interactive new output missing %q:\n%s", want, got)
		}
	}
	if !strings.Contains(got, "Welcome to\n\n ██╗  ██╗") {
		t.Fatalf("expected one blank line between heading and logo:\n%s", got)
	}
	if strings.Index(got, "Docs: https://harumiweb.github.io/xlflow/commands/") > strings.Index(got, "Version: 1.2.3") {
		t.Fatalf("expected command reference before version:\n%s", got)
	}
}

func TestModuleInstallCommandInstallsHelperModules(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"module", "install"})
	if err := root.Execute(); err != nil {
		t.Fatalf("module install command error = %v, exit = %d", err, output.ExitCode(err))
	}
	for _, path := range []string{
		filepath.Join(dir, "src", "modules", "XlflowAssert.bas"),
		filepath.Join(dir, "src", "modules", "XlflowRuntime.bas"),
		filepath.Join(dir, "src", "modules", "XlflowUI.bas"),
		filepath.Join(dir, "src", "modules", "XlflowDebug.bas"),
	} {
		if _, err := os.Stat(path); err != nil {
			t.Fatalf("expected installed helper module at %s: %v", path, err)
		}
	}
}

func TestModuleInstallCommandWithPushAutoPushesHelperSource(t *testing.T) {
	skipWindowsPowerShellOnlyTest(t)
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	writeTestPushScript(t, dir)
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs(withPowerShellBridge("module", "install", "--push"))
	if err := root.Execute(); err != nil {
		t.Fatalf("module install --push command error = %v, exit = %d", err, output.ExitCode(err))
	}
	markerPath := filepath.Join(dir, ".xlflow", "push.called")
	body, err := os.ReadFile(markerPath)
	if err != nil {
		t.Fatalf("expected push marker at %s: %v", markerPath, err)
	}
	text := string(body)
	for _, want := range []string{"XlflowAssert.bas", "XlflowRuntime.bas", "XlflowUI.bas", "XlflowDebug.bas"} {
		if !strings.Contains(text, want) {
			t.Fatalf("expected push marker to mention %s:\n%s", want, text)
		}
	}
}

func TestSkillInstallCommandRefusesOverwriteUnlessForced(t *testing.T) {
	dir := t.TempDir()
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"skill", "install", "--agent", "codex"})
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}

	root = a.rootCommand()
	root.SetArgs([]string{"skill", "install", "--agent", "codex"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("expected overwrite refusal, got err=%v exit=%d", err, output.ExitCode(err))
	}

	root = a.rootCommand()
	root.SetArgs([]string{"skill", "install", "--agent", "codex", "--force"})
	if err := root.Execute(); err != nil {
		t.Fatalf("forced skill install command error = %v, exit = %d", err, output.ExitCode(err))
	}
}

func TestSkillInstallJSONRequiresExplicitTarget(t *testing.T) {
	dir := t.TempDir()
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "skill", "install"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("expected config error, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestBuildDiffOptionsRejectsPartialVBADirs(t *testing.T) {
	_, err := buildDiffOptions(".", "before.xlsx", "after.xlsx", "src1", "")
	if err == nil || !strings.Contains(err.Error(), "--vba-before and --vba-after") {
		t.Fatalf("expected vba dir pair error, got %v", err)
	}
}

func TestBuildDiffOptionsRejectsUnsupportedWorkbookExtensions(t *testing.T) {
	_, err := buildDiffOptions(".", "before.xls", "after.xlsx", "", "")
	if err == nil || !strings.Contains(err.Error(), "unsupported extension") {
		t.Fatalf("expected extension error, got %v", err)
	}
}

func TestBuildInspectCellSelectorRejectsMixedSelectorForms(t *testing.T) {
	_, err := buildInspectCellSelector([]string{"Sheet1!A1"}, "Sheet1", "A1", false)
	if err == nil || !strings.Contains(err.Error(), "cannot combine") {
		t.Fatalf("expected selector conflict, got %v", err)
	}
}

func TestBuildInspectCellSelectorParsesQuotedSheetNames(t *testing.T) {
	got, err := buildInspectCellSelector([]string{"'World News'!A1:F3"}, "", "", true)
	if err != nil {
		t.Fatal(err)
	}
	if got.Sheet != "World News" || got.Address != "A1:F3" {
		t.Fatalf("selector = %#v, want World News A1:F3", got)
	}
}

func TestBuildInspectCellSelectorRejectsInvalidAddressSyntax(t *testing.T) {
	_, err := buildInspectCellSelector(nil, "Visible", "nope", false)
	if err == nil || !strings.Contains(err.Error(), "invalid address") {
		t.Fatalf("expected invalid address error, got %v", err)
	}
}

func TestCollectSourceUserFormNamesFindsRecursiveFrmFiles(t *testing.T) {
	dir := t.TempDir()
	formsDir := filepath.Join(dir, "src", "forms", "Nested")
	if err := os.MkdirAll(formsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "CustomerForm.frm"), []byte("VERSION 5.00"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "OrderForm.frm"), []byte("VERSION 5.00"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "OrderForm.frx"), []byte{0x00}, 0o644); err != nil {
		t.Fatal(err)
	}

	got := collectSourceUserFormNames(filepath.Join(dir, "src", "forms"))
	want := []string{"CustomerForm", "OrderForm"}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("collectSourceUserFormNames() = %#v, want %#v", got, want)
	}
}

func TestInspectSourceUserFormMessagesReturnsWarningAndHint(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := os.MkdirAll(filepath.Join(dir, cfg.Src.Forms), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, cfg.Src.Forms, "UserForm1.frm"), []byte("VERSION 5.00"), 0o644); err != nil {
		t.Fatal(err)
	}

	warnings, hints := inspectSourceUserFormMessages(dir, cfg)
	if len(warnings) != 1 || warnings[0]["code"] != "userform_inspect_saved_file" {
		t.Fatalf("warnings = %#v", warnings)
	}
	if got := fmt.Sprint(warnings[0]["message"]); !strings.Contains(got, "UserForm1") {
		t.Fatalf("warning message = %q", got)
	}
	if len(hints) != 1 || hints[0]["code"] != "userform_planned_commands" {
		t.Fatalf("hints = %#v", hints)
	}
}

func TestDiffCommandReturnsSuccessWhenDifferencesExist(t *testing.T) {
	dir := t.TempDir()
	beforePath := filepath.Join(dir, "before.xlsx")
	afterPath := filepath.Join(dir, "after.xlsx")
	before := excelize.NewFile()
	if err := before.SetCellValue("Sheet1", "A1", "old"); err != nil {
		t.Fatal(err)
	}
	if err := before.SaveAs(beforePath); err != nil {
		t.Fatal(err)
	}
	after := excelize.NewFile()
	if err := after.SetCellValue("Sheet1", "A1", "new"); err != nil {
		t.Fatal(err)
	}
	if err := after.SaveAs(afterPath); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"diff", beforePath, afterPath})
	if err := root.Execute(); err != nil {
		t.Fatalf("diff command error = %v, exit = %d", err, output.ExitCode(err))
	}
}

func TestInspectRangeCommandWritesJSONEnvelope(t *testing.T) {
	dir := t.TempDir()
	createInspectCommandFixture(t, dir)

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "inspect", "range", "--sheet", "Visible", "--address", "A1:C2"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect range error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Command string         `json:"command"`
		Inspect map[string]any `json:"inspect"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Command != "inspect" {
		t.Fatalf("command = %q, want inspect", got.Command)
	}
	if got.Inspect["target"] != "range" {
		t.Fatalf("target = %#v, want range", got.Inspect["target"])
	}
	rangeMap, ok := got.Inspect["range"].(map[string]any)
	if !ok {
		t.Fatalf("range payload = %#v", got.Inspect["range"])
	}
	if rangeMap["range"] != "A1:C2" {
		t.Fatalf("range = %#v, want A1:C2", rangeMap["range"])
	}
}

func TestInspectRangeCommandIncludesStyleMetadataInJSONEnvelope(t *testing.T) {
	dir := t.TempDir()
	createInspectCommandFixture(t, dir)

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "inspect", "range", "--sheet", "Visible", "--address", "A1:C2", "--include-style"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect range include-style error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Inspect map[string]any `json:"inspect"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Inspect["target_info"] == nil {
		t.Fatalf("target_info missing from inspect payload: %s", stdout.String())
	}
	rangeMap, ok := got.Inspect["range"].(map[string]any)
	if !ok {
		t.Fatalf("range payload = %#v", got.Inspect["range"])
	}
	if rangeMap["style_included"] != true {
		t.Fatalf("style_included = %#v, want true", rangeMap["style_included"])
	}
	if _, ok := rangeMap["cells"].([]any); !ok {
		t.Fatalf("cells payload missing from include-style output: %s", stdout.String())
	}
	if _, ok := rangeMap["rows"].([]any); !ok {
		t.Fatalf("rows payload missing from include-style output: %s", stdout.String())
	}
	if _, ok := rangeMap["columns"].([]any); !ok {
		t.Fatalf("columns payload missing from include-style output: %s", stdout.String())
	}
	if _, ok := rangeMap["merged_ranges"].([]any); !ok {
		t.Fatalf("merged_ranges payload missing from include-style output: %s", stdout.String())
	}
}

func TestInspectRangeCommandWritesMarkdown(t *testing.T) {
	dir := t.TempDir()
	createInspectCommandFixture(t, dir)

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"inspect", "range", "--sheet", "Visible", "--address", "A1:C2", "--format", "markdown"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect range markdown error = %v, exit = %d", err, output.ExitCode(err))
	}

	text := stdout.String()
	if !strings.Contains(text, "Snapshot: saved workbook file") {
		t.Fatalf("markdown output = %q, want snapshot note", text)
	}
	if !strings.Contains(text, "Sheet: Visible") {
		t.Fatalf("markdown output = %q, want sheet header", text)
	}
	if !strings.Contains(text, "| C1 | C2 | C3 |") {
		t.Fatalf("markdown output = %q, want markdown table", text)
	}
}

func TestInspectWorkbookJSONIncludesTargetAndSessionState(t *testing.T) {
	dir := t.TempDir()
	createInspectCommandFixture(t, dir)

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "inspect", "workbook"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect workbook json error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Target  map[string]any `json:"target"`
		Session map[string]any `json:"session"`
		Inspect map[string]any `json:"inspect"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Target["kind"] != "file" {
		t.Fatalf("target = %#v", got.Target)
	}
	if got.Session["active"] != false {
		t.Fatalf("session = %#v", got.Session)
	}
	if _, ok := got.Session["running"]; ok {
		t.Fatalf("inspect session must not contain running field (status-specific contract)")
	}
	if _, ok := got.Session["workbook_open"]; ok {
		t.Fatalf("inspect session must not contain workbook_open field (status-specific contract)")
	}
	if _, ok := got.Session["metadata"]; ok {
		t.Fatalf("inspect session must not contain metadata field (status-specific contract)")
	}
	if _, ok := got.Inspect["target_info"].(map[string]any); !ok {
		t.Fatalf("inspect target_info missing: %s", stdout.String())
	}
}

func TestBuildInspectFormBasisDefaultsToRuntime(t *testing.T) {
	got, err := buildInspectFormBasis(false, false, false)
	if err != nil {
		t.Fatal(err)
	}
	if got != "runtime" {
		t.Fatalf("basis = %q, want runtime", got)
	}
}

func TestBuildInspectFormBasisRejectsConflictingModes(t *testing.T) {
	_, err := buildInspectFormBasis(true, true, false)
	if err == nil || !strings.Contains(err.Error(), "choose only one") {
		t.Fatalf("expected conflicting mode error, got %v", err)
	}
}

func TestBuildInspectFormOptionsRejectsInitializerForDesigner(t *testing.T) {
	_, err := buildInspectFormOptions("UserForm1", "designer", "InitializeForm", false, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "--initializer can only be used") {
		t.Fatalf("expected initializer validation error, got %v", err)
	}
}

func TestStaleFileInspectHintsQuoteSelectors(t *testing.T) {
	hints := staleFileInspectHints("range", "World News", "A1:B2")
	if len(hints) == 0 {
		t.Fatal("expected hints")
	}
	message, _ := hints[0]["message"].(string)
	if !strings.Contains(message, "--sheet \"World News\" --address \"A1:B2\"") {
		t.Fatalf("unexpected inspect hint message: %q", message)
	}

	usedRangeHints := staleFileInspectHints("used-range", "Quarterly Report")
	usedRangeMessage, _ := usedRangeHints[0]["message"].(string)
	if !strings.Contains(usedRangeMessage, "--sheet \"Quarterly Report\"") {
		t.Fatalf("unexpected used-range hint message: %q", usedRangeMessage)
	}
}

func TestInspectStateForWorkbookIncludesWorkbookNameFallback(t *testing.T) {
	a := &app{}
	workbookPath := filepath.Join("build", "World News.xlsm")
	_, session, _ := a.inspectStateForWorkbook(config.Config{}, workbookPath)
	if got := session["workbook_name"]; got != filepath.Base(workbookPath) {
		t.Fatalf("workbook_name = %v, want %q", got, filepath.Base(workbookPath))
	}
}

func TestInspectStateForWorkbookExcludesStatusOnlyFields(t *testing.T) {
	_, session, _ := new(app).inspectStateForWorkbook(config.Config{}, filepath.Join("build", "Book.xlsm"))

	statusOnly := []string{"running", "workbook_open", "metadata"}
	for _, field := range statusOnly {
		if _, ok := session[field]; ok {
			t.Errorf("inspectStateForWorkbook must not include field %q (leaks session-status-specific fields into inspect output)", field)
		}
	}
}

func TestInspectFormCommandUsesInspectFormArgsInvalidCode(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{
		cwd:    t.TempDir(),
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "inspect", "form", "UserForm1", "--designer", "--initializer", "InitializeForm"})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("expected config failure, got err=%v exit=%d", err, output.ExitCode(err))
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse inspect form error output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Error.Code != "inspect_form_args_invalid" {
		t.Fatalf("unexpected inspect form error payload: %+v", got)
	}
}

func TestFormSnapshotCommandUsesFormSnapshotArgsInvalidCode(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{
		cwd:    t.TempDir(),
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "form", "snapshot", "UserForm1", "--out", "artifacts\\UserForm1.form.txt"})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("expected config failure, got err=%v exit=%d", err, output.ExitCode(err))
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse form snapshot error output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Error.Code != "form_snapshot_args_invalid" {
		t.Fatalf("unexpected form snapshot error payload: %+v", got)
	}
}

func TestFormBuildCommandReturnsSpecParseMetadata(t *testing.T) {
	dir := t.TempDir()
	specPath := filepath.Join(dir, "src", "forms", "specs", "UserForm1.yaml")
	if err := os.MkdirAll(filepath.Dir(specPath), 0o755); err != nil {
		t.Fatal(err)
	}
	body := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: UserForm1\n  caption: -\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(specPath, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "form", "build", specPath})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected validation failure, got err=%v exit=%d", err, output.ExitCode(err))
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code string `json:"code"`
		} `json:"error"`
		Spec map[string]any `json:"spec"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse form build error output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Error.Code != "spec_parse_failed" {
		t.Fatalf("unexpected form build error payload: %+v", got)
	}
	if got.Spec["format"] != "yaml" || got.Spec["path"] != "src/forms/specs/UserForm1.yaml" {
		t.Fatalf("unexpected spec metadata: %+v", got.Spec)
	}
	if suggestion, _ := got.Spec["suggestion"].(string); !strings.Contains(suggestion, `caption: ""`) {
		t.Fatalf("unexpected suggestion: %+v", got.Spec)
	}
}

func TestFormBuildSidecarModeRunsSourcePreflightBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "sidecar"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	specPath := filepath.Join(dir, "src", "forms", "specs", "UserForm1.yaml")
	specBody := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: UserForm1\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(specPath, []byte(specBody), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecarBody := "Option Explicit\n\nPublic Sub BreakAnalyzer()\n  Dim ws As Worksheet\n  Set ws = ThisWorkbook.Worksheets(1)\n  ws.DisplayGridlines = True\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "UserForm1.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "form", "build", "src/forms/specs/UserForm1.yaml"})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected validation failure, got err=%v exit=%d", err, output.ExitCode(err))
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code  string `json:"code"`
			Phase string `json:"phase"`
		} `json:"error"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse form build preflight output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Error.Phase != "preflight" {
		t.Fatalf("unexpected form build preflight payload: %+v", got)
	}
	if got.Error.Code != "analyze_failed" && got.Error.Code != "source_preflight_failed" {
		t.Fatalf("unexpected preflight error code: %+v", got)
	}
}

func TestFormBuildSidecarModeSyncsEmbeddedCodeBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "sidecar"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	specPath := filepath.Join(dir, "src", "forms", "specs", "UserForm1.yaml")
	specBody := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: UserForm1\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(specPath, []byte(specBody), 0o644); err != nil {
		t.Fatal(err)
	}
	frmBody := "VERSION 5.00\nBegin {GUID} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n\nPrivate Sub UserForm_Initialize()\n    version = \"frm\"\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "UserForm1.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecarBody := "Option Explicit\n\nPrivate Sub UserForm_Initialize()\n    version = \"sidecar\"\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "UserForm1.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if err := a.runUserFormCodeSourcePreflight("form build", cfg, map[string]bool{"UserForm1": true}); err != nil {
		t.Fatalf("runUserFormCodeSourcePreflight() error = %v, exit = %d", err, output.ExitCode(err))
	}
	rewritten, err := os.ReadFile(filepath.Join(dir, "src", "forms", "UserForm1.frm"))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(rewritten), `version = "sidecar"`) || strings.Contains(string(rewritten), `version = "frm"`) {
		t.Fatalf("frm artifact was not synchronized from sidecar:\n%s", string(rewritten))
	}
}

func TestFormBuildSidecarModeRejectsAttributeContaminatedSidecarBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "sidecar"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	specPath := filepath.Join(dir, "src", "forms", "specs", "UserForm1.yaml")
	specBody := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: UserForm1\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(specPath, []byte(specBody), 0o644); err != nil {
		t.Fatal(err)
	}
	frmBody := "VERSION 5.00\nBegin {GUID} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "UserForm1.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecarBody := "Attribute VB_Name = \"UserForm1\"\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "UserForm1.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "form", "build", "src/forms/specs/UserForm1.yaml"})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected validation failure, got err=%v exit=%d", err, output.ExitCode(err))
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code  string `json:"code"`
			Phase string `json:"phase"`
		} `json:"error"`
		Issues []struct {
			Code       string `json:"code"`
			File       string `json:"file"`
			Line       int    `json:"line"`
			Suggestion string `json:"suggestion"`
		} `json:"issues"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse form build contaminated-sidecar output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Error.Code != "source_preflight_failed" || got.Error.Phase != "preflight" {
		t.Fatalf("unexpected contaminated-sidecar preflight payload: %+v", got)
	}
	if len(got.Issues) != 1 || got.Issues[0].Code != "FRM202" || got.Issues[0].Line != 1 {
		t.Fatalf("unexpected contaminated-sidecar issues: %+v", got.Issues)
	}
	if !strings.Contains(got.Issues[0].Suggestion, "Attribute VB_*") {
		t.Fatalf("unexpected suggestion: %+v", got.Issues[0])
	}
}

func TestFormApplySidecarModeRunsPreflightBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "sidecar"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	specPath := filepath.Join(dir, "src", "forms", "specs", "UserForm1.yaml")
	specBody := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: UserForm1\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(specPath, []byte(specBody), 0o644); err != nil {
		t.Fatal(err)
	}
	frmBody := "VERSION 5.00\nBegin {GUID} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n\nPrivate Sub UserForm_Initialize()\n    version = \"frm\"\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "UserForm1.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecarBody := "Option Explicit\n\nPrivate Sub UserForm_Initialize()\n    version = \"sidecar\"\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "UserForm1.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	opts := formWriteCommandOptions{
		Action: "apply",
		Spec: forms.FormSpec{
			Form: forms.FormSpecForm{Name: "UserForm1"},
		},
	}
	if err := a.runFormWritePreflight("form apply", cfg, opts); err != nil {
		t.Fatalf("runFormWritePreflight() error = %v, exit = %d", err, output.ExitCode(err))
	}
	rewritten, err := os.ReadFile(filepath.Join(dir, "src", "forms", "UserForm1.frm"))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(rewritten), `version = "sidecar"`) || strings.Contains(string(rewritten), `version = "frm"`) {
		t.Fatalf("frm artifact was not synchronized from sidecar:\n%s", string(rewritten))
	}
}

func TestInspectSymbolsJSONEnvelope(t *testing.T) {
	dir := t.TempDir()
	writeInspectSymbolsFixture(t, dir, filepath.Join("src", "modules"))

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "inspect", "symbols"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect symbols json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var got struct {
		Status  string `json:"status"`
		Command string `json:"command"`
		Inspect struct {
			Target  string `json:"target"`
			Source  string `json:"source"`
			Root    string `json:"root"`
			Summary struct {
				Files   int `json:"files"`
				Symbols int `json:"symbols"`
			} `json:"summary"`
			Files []struct {
				Path       string `json:"path"`
				ModuleName string `json:"moduleName"`
				Symbols    []struct {
					Name string `json:"name"`
					Kind string `json:"kind"`
				} `json:"symbols"`
			} `json:"files"`
		} `json:"inspect"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("failed to parse inspect symbols output: %v\n%s", err, stdout.String())
	}
	if got.Status != output.StatusOK || got.Command != "inspect" || got.Inspect.Target != "symbols" || got.Inspect.Source != "tree_sitter_vba" {
		t.Fatalf("unexpected envelope: %+v", got)
	}
	if got.Inspect.Root != "src" || got.Inspect.Summary.Files != 1 || len(got.Inspect.Files) != 1 {
		t.Fatalf("unexpected inspect summary: %+v", got.Inspect)
	}
	if got.Inspect.Files[0].Path != "src/modules/Main.bas" || got.Inspect.Files[0].ModuleName != "Main" {
		t.Fatalf("unexpected file result: %+v", got.Inspect.Files[0])
	}
	found := false
	for _, symbol := range got.Inspect.Files[0].Symbols {
		if symbol.Name == "Run" && symbol.Kind == "sub" {
			found = true
		}
	}
	if !found {
		t.Fatalf("Run symbol missing: %+v", got.Inspect.Files[0].Symbols)
	}
}

func TestInspectSymbolsStandaloneFormatJSON(t *testing.T) {
	dir := t.TempDir()
	writeInspectSymbolsFixture(t, dir, filepath.Join("src", "modules"))

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"inspect", "--format", "json", "symbols"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect symbols format json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var got map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("failed to parse standalone inspect output: %v\n%s", err, stdout.String())
	}
	if got["status"] != nil || got["command"] != nil {
		t.Fatalf("standalone inspect output should not be envelope: %+v", got)
	}
	if got["target"] != "symbols" || got["source"] != "tree_sitter_vba" {
		t.Fatalf("unexpected standalone payload: %+v", got)
	}
}

func TestInspectSymbolsTextAndPath(t *testing.T) {
	dir := t.TempDir()
	writeInspectSymbolsFixture(t, dir, "custom-src")

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"inspect", "symbols", "--path", "custom-src"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect symbols text error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	for _, want := range []string{"custom-src/Main.bas", "Module Main", "Public Sub Run()"} {
		if !strings.Contains(got, want) {
			t.Fatalf("inspect symbols text missing %q:\n%s", want, got)
		}
	}
}

func TestInspectCallsJSONEnvelope(t *testing.T) {
	dir := t.TempDir()
	writeInspectCallsFixture(t, dir, filepath.Join("src", "modules"))

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "inspect", "calls"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect calls json error = %v, exit = %d", err, output.ExitCode(err))
	}
	var got struct {
		Status  string `json:"status"`
		Command string `json:"command"`
		Inspect struct {
			Target  string `json:"target"`
			Source  string `json:"source"`
			Root    string `json:"root"`
			Summary struct {
				Files int `json:"files"`
				Calls int `json:"calls"`
			} `json:"summary"`
			Calls []struct {
				File   string `json:"file"`
				Module string `json:"module"`
				Caller struct {
					QualifiedName string `json:"qualifiedName"`
				} `json:"caller"`
				Callee struct {
					Text     string `json:"text"`
					BaseName string `json:"baseName"`
				} `json:"callee"`
				Resolution struct {
					Status string `json:"status"`
				} `json:"resolution"`
			} `json:"calls"`
		} `json:"inspect"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("failed to parse inspect calls output: %v\n%s", err, stdout.String())
	}
	if got.Status != output.StatusOK || got.Command != "inspect" || got.Inspect.Target != "calls" || got.Inspect.Source != "tree_sitter_vba" {
		t.Fatalf("unexpected envelope: %+v", got)
	}
	if got.Inspect.Root != "src" || got.Inspect.Summary.Files != 1 || got.Inspect.Summary.Calls == 0 {
		t.Fatalf("unexpected inspect summary: %+v", got.Inspect)
	}
	found := false
	for _, call := range got.Inspect.Calls {
		if call.Callee.Text == "BuildReport" && call.Caller.QualifiedName == "Main.Run" && call.Resolution.Status == "matched" {
			found = true
		}
	}
	if !found {
		t.Fatalf("BuildReport call missing: %+v", got.Inspect.Calls)
	}
}

func TestInspectCallsTextPathAndFilters(t *testing.T) {
	dir := t.TempDir()
	writeInspectCallsFixture(t, dir, "custom-src")

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"inspect", "calls", "--path", "custom-src", "--from", "Main.Run", "--to", "BuildReport"})

	if err := root.Execute(); err != nil {
		t.Fatalf("inspect calls text error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	for _, want := range []string{"custom-src/Main.bas", "Main.Run", "-> BuildReport"} {
		if !strings.Contains(got, want) {
			t.Fatalf("inspect calls text missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "Debug.Print") {
		t.Fatalf("inspect calls filter included Debug.Print:\n%s", got)
	}
}

func writeInspectSymbolsFixture(t *testing.T, dir, sourceDir string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Join(dir, sourceDir), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
End Sub
`
	if err := os.WriteFile(filepath.Join(dir, sourceDir, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func writeInspectCallsFixture(t *testing.T, dir, sourceDir string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Join(dir, sourceDir), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    BuildReport 1, 2
    Debug.Print "done"
End Sub

Public Sub BuildReport(ByVal first As Long, ByVal second As Long)
End Sub
`
	if err := os.WriteFile(filepath.Join(dir, sourceDir, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func TestPushRejectsAttributeContaminatedUserFormSidecarBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "modules"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "classes"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "workbook"), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "sidecar"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	frmBody := "VERSION 5.00\nBegin {GUID} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "UserForm1.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecarBody := "Attribute VB_Name = \"UserForm1\"\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "UserForm1.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected validation failure, got err=%v exit=%d", err, output.ExitCode(err))
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code  string `json:"code"`
			Phase string `json:"phase"`
		} `json:"error"`
		Issues []struct {
			Code string `json:"code"`
			File string `json:"file"`
			Line int    `json:"line"`
		} `json:"issues"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse push contaminated-sidecar output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Error.Code != "source_preflight_failed" || got.Error.Phase != "preflight" {
		t.Fatalf("unexpected push contaminated-sidecar payload: %+v", got)
	}
	if len(got.Issues) != 1 || got.Issues[0].Code != "FRM202" || got.Issues[0].Line != 1 {
		t.Fatalf("unexpected push contaminated-sidecar issues: %+v", got.Issues)
	}
}

func TestPushRejectsSpecDrivenUserFormArtifactNameMismatchBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "modules"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "classes"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "workbook"), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "sidecar"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	specBody := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: RegistrationForm\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "specs", "RegistrationForm.yaml"), []byte(specBody), 0o644); err != nil {
		t.Fatal(err)
	}
	frmBody := "VERSION 5.00\nBegin {GUID} RegistrationForm\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "RegistrationForm.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecarBody := "Option Explicit\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "RegistrationForm.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected validation failure, got err=%v exit=%d", err, output.ExitCode(err))
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code  string `json:"code"`
			Phase string `json:"phase"`
		} `json:"error"`
		Issues []struct {
			Code string `json:"code"`
			File string `json:"file"`
			Line int    `json:"line"`
		} `json:"issues"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse push spec/artifact mismatch output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Error.Code != "source_preflight_failed" || got.Error.Phase != "preflight" {
		t.Fatalf("unexpected push spec/artifact mismatch payload: %+v", got)
	}
	if len(got.Issues) != 1 || got.Issues[0].Code != "FRM201" || got.Issues[0].Line != 4 {
		t.Fatalf("unexpected push spec/artifact mismatch issues: %+v", got.Issues)
	}
}

func TestFormBuildSidecarModePreflightIgnoresOtherForms(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	configBody := `[project]
entry = "Main.Run"

[excel]
path = "build/missing.xlsm"

[userform]
code_source = "sidecar"
`
	if err := os.WriteFile(filepath.Join(dir, config.FileName), []byte(configBody), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "specs", "UserForm1.yaml"), []byte("schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: UserForm1\ncontrols: []\nwarnings: []\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "UserForm1.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	staleFrm := "VERSION 5.00\nBegin {GUID} UserForm2\nEnd\nAttribute VB_Name = \"UserForm2\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n\nPublic Sub BreakAnalyzer()\n  Dim ws As Worksheet\n  Set ws = ThisWorkbook.Worksheets(1)\n  ws.DisplayGridlines = True\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "UserForm2.frm"), []byte(staleFrm), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", "code", "UserForm2.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "form", "build", "src/forms/specs/UserForm1.yaml"})

	err := root.Execute()
	if err == nil {
		t.Fatal("expected form build to fail after preflight because workbook path is missing")
	}

	var got struct {
		Status string `json:"status"`
		Error  struct {
			Code  string `json:"code"`
			Phase string `json:"phase"`
		} `json:"error"`
	}
	if unmarshalErr := json.Unmarshal(stdout.Bytes(), &got); unmarshalErr != nil {
		t.Fatalf("failed to parse form build output: %v\n%s", unmarshalErr, stdout.String())
	}
	if got.Error.Phase == "preflight" && (got.Error.Code == "analyze_failed" || got.Error.Code == "source_preflight_failed") {
		t.Fatalf("unrelated UserForm2 should not block UserForm1 build preflight: %+v", got)
	}
}

func TestBuildRunOptionsRejectsConflictingSaveFlags(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Args: []string{"string:hello"}, Save: true, SaveAs: "build\\result.xlsm"})
	if err == nil || !strings.Contains(err.Error(), "--save and --save-as cannot be combined") {
		t.Fatalf("expected save conflict error, got %v", err)
	}
}

func TestRunCommandRejectsNoSaveCombinedWithSaveFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	var stdout bytes.Buffer
	var stderr bytes.Buffer
	root.SetOut(&stdout)
	root.SetErr(&stderr)
	root.SetArgs([]string{"--json", "run", "Main.Run", "--no-save", "--save"})

	err := root.Execute()
	if err == nil {
		t.Fatal("expected run command to reject --no-save with --save")
	}
}

func TestBuildRunOptionsParsesTypedArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Workbook: "fixtures\\Book.xlsm", Args: []string{"string:hello", "int:7", "double:3.5", "bool:true"}, Headless: true})
	if err != nil {
		t.Fatal(err)
	}

	want := []excel.RunArgument{
		{Type: "string", Value: "hello"},
		{Type: "int", Value: "7"},
		{Type: "double", Value: "3.5"},
		{Type: "bool", Value: "true"},
	}
	if opts.Macro != "Main.Run" {
		t.Fatalf("macro = %q, want Main.Run", opts.Macro)
	}
	if opts.WorkbookPath != "fixtures\\Book.xlsm" {
		t.Fatalf("workbook path = %q", opts.WorkbookPath)
	}
	if opts.Mode != "headless" {
		t.Fatalf("mode = %q, want headless", opts.Mode)
	}
	if opts.RuntimeMode != excel.RuntimeModeHeadless || opts.RuntimeSource != excel.RuntimeSourceCommand {
		t.Fatalf("runtime = (%q, %q), want (%q, %q)", opts.RuntimeMode, opts.RuntimeSource, excel.RuntimeModeHeadless, excel.RuntimeSourceCommand)
	}
	if opts.Timeout != 5*time.Minute {
		t.Fatalf("timeout = %s", opts.Timeout)
	}
	if opts.Keepalive.Stderr != nil {
		t.Fatalf("command stderr = %#v, want nil", opts.Keepalive.Stderr)
	}
	if !reflect.DeepEqual(opts.Args, want) {
		t.Fatalf("run args = %#v, want %#v", opts.Args, want)
	}
}

func TestBuildRunOptionsParsesUIResponses(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", MsgBox: []string{"confirm-save=yes"}, InputBox: []string{"customer-name=John"}})
	if err != nil {
		t.Fatal(err)
	}
	if got := opts.UIResponses.MsgBox["confirm_save"]; got != "yes" {
		t.Fatalf("msgbox response = %q, want yes", got)
	}
	if got := opts.UIResponses.Input["customer_name"]; got != "John" {
		t.Fatalf("input response = %q, want John", got)
	}
}

func TestBuildRunOptionsParsesFileDialogResponses(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", FileDialog: []string{"get-open:source-files=C:\\tmp\\a.txt", "get-open:source-files=C:\\tmp\\b.txt", "save-as:result-file=C:\\tmp\\out.xlsx"}})
	if err != nil {
		t.Fatal(err)
	}
	if len(opts.UIResponses.FileDialog) != 2 {
		t.Fatalf("file dialog responses = %#v, want 2 entries", opts.UIResponses.FileDialog)
	}
	openDialog := opts.UIResponses.FileDialog[0]
	if openDialog.Kind != "get-open" || openDialog.DialogID != "source_files" {
		t.Fatalf("open dialog = %#v", openDialog)
	}
	if !reflect.DeepEqual(openDialog.Values, []string{"C:\\tmp\\a.txt", "C:\\tmp\\b.txt"}) {
		t.Fatalf("open dialog values = %#v", openDialog.Values)
	}
	saveDialog := opts.UIResponses.FileDialog[1]
	if saveDialog.Kind != "save-as" || saveDialog.DialogID != "result_file" {
		t.Fatalf("save dialog = %#v", saveDialog)
	}
	if !reflect.DeepEqual(saveDialog.Values, []string{"C:\\tmp\\out.xlsx"}) {
		t.Fatalf("save dialog values = %#v", saveDialog.Values)
	}
	if saveDialog.Cancelled {
		t.Fatalf("save dialog cancelled = true, want false: %#v", saveDialog)
	}
}

func TestBuildRunOptionsParsesCancelledFileDialogResponses(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", FileDialog: []string{"folder:target-dir=@cancel"}})
	if err != nil {
		t.Fatal(err)
	}
	if len(opts.UIResponses.FileDialog) != 1 {
		t.Fatalf("file dialog responses = %#v, want 1 entry", opts.UIResponses.FileDialog)
	}
	if !opts.UIResponses.FileDialog[0].Cancelled {
		t.Fatalf("cancelled = false, want true: %#v", opts.UIResponses.FileDialog[0])
	}
}

func TestTestCommandParsesFileDialogFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	cmd, _, err := root.Find([]string{"test"})
	if err != nil {
		t.Fatal(err)
	}
	if err := cmd.ParseFlags([]string{
		"--filedialog", "get-open:source-files=C:\\tmp\\a.txt",
		"--filedialog", "get-open:source-files=C:\\tmp\\b.txt",
		"--filedialog", "save-as:result-file=C:\\tmp\\out.xlsx",
		"--filedialog", "folder:target-dir=@cancel",
	}); err != nil {
		t.Fatal(err)
	}
	literals, err := cmd.Flags().GetStringArray("filedialog")
	if err != nil {
		t.Fatal(err)
	}
	parsed, err := parseFileDialogResponseLiterals(literals)
	if err != nil {
		t.Fatal(err)
	}
	if len(parsed) != 3 {
		t.Fatalf("file dialog responses = %#v, want 3 entries", parsed)
	}
	if got := parsed[0]; got.Kind != "get-open" || got.DialogID != "source_files" || !reflect.DeepEqual(got.Values, []string{"C:\\tmp\\a.txt", "C:\\tmp\\b.txt"}) || got.Cancelled {
		t.Fatalf("open dialog = %#v", got)
	}
	if got := parsed[1]; got.Kind != "save-as" || got.DialogID != "result_file" || !reflect.DeepEqual(got.Values, []string{"C:\\tmp\\out.xlsx"}) || got.Cancelled {
		t.Fatalf("save dialog = %#v", got)
	}
	if got := parsed[2]; got.Kind != "folder" || got.DialogID != "target_dir" || !got.Cancelled || len(got.Values) != 0 {
		t.Fatalf("folder dialog = %#v", got)
	}
}

func TestBuildRunOptionsRejectsInvalidFileDialogResponses(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", FileDialog: []string{"unknown:pick=C:\\tmp\\a.txt"}})
	if err == nil || !strings.Contains(err.Error(), "unsupported filedialog kind") {
		t.Fatalf("expected filedialog kind error, got %v", err)
	}
	_, err = buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", FileDialog: []string{"save-as:result=C:\\tmp\\a.txt", "save-as:result=C:\\tmp\\b.txt"}})
	if err == nil || !strings.Contains(err.Error(), "accepts only one scripted path") {
		t.Fatalf("expected single-path error, got %v", err)
	}
	_, err = buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", FileDialog: []string{"folder:target=@cancel", "folder:target=C:\\tmp"}})
	if err == nil || !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected cancel/path conflict, got %v", err)
	}
	_, err = buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", FileDialog: []string{"get-open:confirm save=C:\\tmp\\a.txt", "get-open:confirm-save=C:\\tmp\\b.txt"}})
	if err == nil || !strings.Contains(err.Error(), "collides with") {
		t.Fatalf("expected normalized collision error, got %v", err)
	}
}

func TestBuildRunOptionsPreservesInputBoxWhitespaceAndNormalizesMsgBox(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", MsgBox: []string{"confirm-save= YES "}, InputBox: []string{"customer-name=  Alice  ", "single-space= "}})
	if err != nil {
		t.Fatal(err)
	}
	if got := opts.UIResponses.MsgBox["confirm_save"]; got != "yes" {
		t.Fatalf("msgbox response = %q, want yes", got)
	}
	if got := opts.UIResponses.Input["customer_name"]; got != "  Alice  " {
		t.Fatalf("input response = %q, want preserved whitespace", got)
	}
	if got := opts.UIResponses.Input["single_space"]; got != " " {
		t.Fatalf("single-space input response = %q, want one space", got)
	}
}

func TestBuildRunOptionsRejectsNonStableDialogIDs(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", MsgBox: []string{"!!!=yes"}})
	if err == nil || !strings.Contains(err.Error(), "must contain at least one ASCII letter or digit") {
		t.Fatalf("expected invalid dialog id error, got %v", err)
	}
	_, err = buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", MsgBox: []string{"confirm save=yes", "confirm-save=no"}})
	if err == nil || !strings.Contains(err.Error(), "collides with") {
		t.Fatalf("expected normalized collision error, got %v", err)
	}
}

func TestBuildRunOptionsUsesEnvironmentRuntimeOverrideByDefault(t *testing.T) {
	t.Setenv("XLFLOW_MODE", excel.RuntimeModeAgent)
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run"})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Mode != excel.RuntimeModeHeadless {
		t.Fatalf("mode = %q, want %q operational mode", opts.Mode, excel.RuntimeModeHeadless)
	}
	if opts.RuntimeMode != excel.RuntimeModeAgent || opts.RuntimeSource != excel.RuntimeSourceEnvironment {
		t.Fatalf("runtime = (%q, %q), want (%q, %q)", opts.RuntimeMode, opts.RuntimeSource, excel.RuntimeModeAgent, excel.RuntimeSourceEnvironment)
	}
}

func TestBuildRunOptionsAllowsEmptyStringArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Args: []string{"string:"}})
	if err != nil {
		t.Fatal(err)
	}
	if len(opts.Args) != 1 || opts.Args[0].Type != "string" || opts.Args[0].Value != "" {
		t.Fatalf("run args = %#v", opts.Args)
	}
}

func TestBuildRunOptionsRejectsConflictingRunModes(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Headless: true, Interactive: true})
	if err == nil || !strings.Contains(err.Error(), "--headless and --interactive") {
		t.Fatalf("expected run mode conflict error, got %v", err)
	}
}

func TestBuildRunOptionsRejectsDirectWithArgs(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Args: []string{"string:hello"}, Direct: true})
	if err == nil || !strings.Contains(err.Error(), "--direct cannot be used with --arg") {
		t.Fatalf("expected direct arg conflict, got %v", err)
	}
}

func TestRunCommandRejectsRemovedTraceFlag(t *testing.T) {
	a := &app{
		stdout:         new(bytes.Buffer),
		stderr:         new(bytes.Buffer),
		cwd:            t.TempDir(),
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"run", "--trace"})
	err := root.Execute()
	if err == nil || !strings.Contains(err.Error(), "unknown flag: --trace") {
		t.Fatalf("expected unknown flag error, got %v", err)
	}
}

func TestBuildRunOptionsRejectsDirectWithDiagnostic(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Direct: true, Diagnostic: true, DiagnosticExplicit: true})
	if err == nil || !strings.Contains(err.Error(), "--gui-compile-errors") {
		t.Fatalf("expected direct diagnostic conflict, got %v", err)
	}
}

func TestBuildRunOptionsAutoDisablesDefaultDiagnosticForDirect(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Direct: true, Diagnostic: true})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Diagnostic {
		t.Fatalf("diagnostic = true, want false for default direct run: %#v", opts)
	}
	if !opts.Direct {
		t.Fatalf("direct = false, want true: %#v", opts)
	}
	if !opts.SuppressModalErrors {
		t.Fatalf("SuppressModalErrors = false, want true for default direct run: %#v", opts)
	}
}

func TestBuildRunOptionsAutoDisablesDefaultDiagnosticForFast(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Fast: true, Diagnostic: true})
	if err != nil {
		t.Fatal(err)
	}
	if !opts.Fast {
		t.Fatalf("fast = false, want true: %#v", opts)
	}
	if opts.Diagnostic {
		t.Fatalf("diagnostic = true, want false for default fast run: %#v", opts)
	}
}

func TestBuildRunOptionsAllowsExplicitFastDiagnostic(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Fast: true, Diagnostic: true, DiagnosticExplicit: true})
	if err != nil {
		t.Fatal(err)
	}
	if !opts.Fast || !opts.Diagnostic {
		t.Fatalf("unexpected fast diagnostic options: %#v", opts)
	}
}

func TestBuildRunOptionsAllowsDirectWhenGUICompileErrorsOptOutIsSet(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Direct: true, Diagnostic: true, GUICompileErrors: true})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Diagnostic {
		t.Fatalf("diagnostic = true, want false with gui compile error opt-out: %#v", opts)
	}
	if !opts.Direct {
		t.Fatalf("direct = false, want true: %#v", opts)
	}
	if opts.SuppressModalErrors {
		t.Fatalf("SuppressModalErrors = true, want false with gui error opt-out: %#v", opts)
	}
}

func TestBuildRunOptionsWithUIStreamEnablesRedactedStreamByDefault(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", UIStream: true})
	if err != nil {
		t.Fatal(err)
	}
	if !opts.UIStream.Enabled {
		t.Fatalf("UIStream.Enabled = false, want true: %#v", opts)
	}
	if !opts.UIStream.RedactInput {
		t.Fatalf("UIStream.RedactInput = false, want true: %#v", opts)
	}
	if !opts.DebugStream.Enabled {
		t.Fatalf("DebugStream.Enabled = false, want true: %#v", opts)
	}
}

func TestBuildRunOptionsWithUIStreamRejectsDirect(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Direct: true, UIStream: true})
	if err == nil || !strings.Contains(err.Error(), "--direct cannot be combined with --ui-stream") {
		t.Fatalf("expected direct ui-stream conflict, got %v", err)
	}
}

func TestBuildRunOptionsWithUIStreamRejectsFast(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Fast: true, UIStream: true})
	if err == nil || !strings.Contains(err.Error(), "--fast cannot be combined with --ui-stream") {
		t.Fatalf("expected fast ui-stream conflict, got %v", err)
	}
}

func TestBuildPushOptionsExpandsFastAndRequiresSessionForNoSave(t *testing.T) {
	opts, err := buildPushOptions("always", true, false, false, false, excel.CommandOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if opts.BackupMode != "never" || !opts.ChangedOnly || !opts.Fast {
		t.Fatalf("unexpected fast push options: %#v", opts)
	}
	_, err = buildPushOptions("always", false, false, false, true, excel.CommandOptions{})
	if err == nil || !strings.Contains(err.Error(), "--no-save requires --session") {
		t.Fatalf("expected no-save session error, got %v", err)
	}
}

func TestBuildRollbackTargetRequiresExactlyOneSelector(t *testing.T) {
	if _, err := buildRollbackTarget(false, ""); err == nil {
		t.Fatal("expected selector error")
	}
	if _, err := buildRollbackTarget(true, "20260518-100000-push"); err == nil {
		t.Fatal("expected exclusive selector error")
	}
	target, err := buildRollbackTarget(true, "")
	if err != nil {
		t.Fatal(err)
	}
	if !target.Latest {
		t.Fatalf("target = %#v, want latest", target)
	}
}

func TestBackupListCommandReturnsWorkbookBackupsOnly(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	workbookPath := filepath.Join(dir, "build", "Book.xlsm")
	if err := os.MkdirAll(filepath.Dir(workbookPath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(workbookPath, []byte("book"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := backup.Create(dir, workbookPath, "before-push", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC)); err != nil {
		t.Fatal(err)
	}
	legacyDir := filepath.Join(dir, ".xlflow", "backups", "legacy")
	if err := os.MkdirAll(legacyDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(legacyDir, "Module1.bas"), []byte("legacy"), 0o644); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "backup", "list"})
	if err := root.Execute(); err != nil {
		t.Fatalf("backup list error = %v", err)
	}
	var got struct {
		Backups []map[string]any `json:"backups"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if len(got.Backups) != 1 {
		t.Fatalf("backups = %#v, want 1", got.Backups)
	}
}

func TestRollbackCommandRequiresBackupSelector(t *testing.T) {
	dir := t.TempDir()
	if err := config.Write(filepath.Join(dir, config.FileName), config.Default()); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "rollback"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("rollback err=%v exit=%d", err, output.ExitCode(err))
	}
}

func createInspectCommandFixture(t *testing.T, dir string) {
	t.Helper()

	cfg := config.Default()
	cfg.Excel.Path = filepath.Join("build", "Book.xlsx")
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	buildDir := filepath.Join(dir, "build")
	if err := os.MkdirAll(buildDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(buildDir, "Book.xlsx")
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "Visible"); err != nil {
		t.Fatal(err)
	}
	if _, err := f.NewSheet("Hidden"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Visible", "A1", "A1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Visible", "C1", "C1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Visible", "B2", "B2"); err != nil {
		t.Fatal(err)
	}
	styleID, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"000000"}},
		Font: &excelize.Font{Family: "Calibri", Size: 11, Bold: true, Color: "FFFFFF"},
		Border: []excelize.Border{
			{Type: "right", Style: 1, Color: "D9D9D9"},
		},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		NumFmt:    10,
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle("Visible", "A1", "A1", styleID); err != nil {
		t.Fatal(err)
	}
	if err := f.SetRowHeight("Visible", 2, 25); err != nil {
		t.Fatal(err)
	}
	if err := f.SetColWidth("Visible", "B", "B", 20); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSheetVisible("Hidden", false); err != nil {
		t.Fatal(err)
	}
	f.SetActiveSheet(0)
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
}

func TestRunHeadlessPreflightRejectsGUIBoundariesBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  MsgBox \"stop\"\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "run", "Main.Run", "--headless"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestRunHeadlessPreflightRejectsFullyQualifiedRawDialogsBeforeExcel(t *testing.T) {
	for _, tc := range []struct {
		name string
		expr string
	}{
		{name: "msgbox", expr: `VBA.Interaction.MsgBox("stop")`},
		{name: "inputbox", expr: `VBA.Interaction.InputBox("name?")`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			cfg := config.Default()
			if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
				t.Fatal(err)
			}
			src := filepath.Join(dir, "src", "modules")
			if err := os.MkdirAll(src, 0o755); err != nil {
				t.Fatal(err)
			}
			body := "Option Explicit\nPublic Sub Run()\n  Dim value As Variant\n  value = " + tc.expr + "\nEnd Sub\n"
			if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
				t.Fatal(err)
			}

			a := &app{cwd: dir}
			root := a.rootCommand()
			root.SetArgs([]string{"--json", "run", "Main.Run", "--headless"})
			err := root.Execute()
			if err == nil || output.ExitCode(err) != output.ExitValidation {
				t.Fatalf("expected validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
			}
		})
	}
}

func TestPushRejectsTypographicQuotesBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  If Mid$(text, index, 1) <> “\"\" Then\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected source validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestPushRejectsLikelyCStyleQuoteEscapesBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  If Mid$(text, index, 1) <> \"\\\"\" Then\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected source validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestPushRejectsVBAProcedureSyntaxBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\nEnd Function\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected source validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestPushRejectsMissingLineContinuationWhitespaceBeforeExcel(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  Debug.Print \"hello\"_\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected source validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestAnalyzeCommandReturnsValidationForFindings(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  Dim ws As Worksheet\n  ws = ThisWorkbook.Worksheets(1)\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "analyze"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected analysis validation failure, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestPushCommandReturnsValidationForBlockingAnalysisFindings(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  Dim ws As Worksheet\n  Set ws = ThisWorkbook.Worksheets(1)\n  With ws\n    .DisplayGridlines = False\n  End With\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected source validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestRunCommandReturnsValidationForBlockingAnalysisFindings(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  Dim ws As Worksheet\n  Set ws = ThisWorkbook.Worksheets(1)\n  ws.DisplayGridlines = False\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "run", "Main.Run", "--interactive"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("expected run preflight validation failure before Excel, got err=%v exit=%d", err, output.ExitCode(err))
	}
}

func TestRunCommandRejectsDiagnosticAndGUICompileErrorsTogether(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}

	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "run", "Main.Run", "--diagnostic", "--gui-compile-errors"})
	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("expected run flag conflict, got err=%v exit=%d", err, output.ExitCode(err))
	}
	if !strings.Contains(err.Error(), "--gui-compile-errors") {
		t.Fatalf("error = %v, want gui compile error conflict", err)
	}
}

func TestBuildRunDiagnosticPreservesExistingScriptDiagnostic(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	a := &app{cwd: dir}
	env := output.Failure("run", output.Error{Code: "macro_failed", Message: "runtime modal", Source: "Main", Number: 438, Line: 12, Phase: "invoke_macro"})
	env.RunDiagnostic = map[string]any{
		"kind":     "runtime",
		"message":  []string{"Run-time error '438':", "Object doesn't support this property or method."},
		"dialog":   map[string]any{"title": "Microsoft Visual Basic"},
		"location": map[string]any{"module": "Main", "line": 12},
	}

	diag := a.buildRunDiagnostic(cfg, env)
	if got := cliObjectMap(diag["dialog"]); got["title"] != "Microsoft Visual Basic" {
		t.Fatalf("dialog metadata was not preserved: %#v", diag)
	}
	if got := cliObjectMap(diag["location"]); got["module"] != "Main" {
		t.Fatalf("location metadata was not preserved: %#v", diag)
	}
	if got := diag["kind"]; got != "runtime" {
		t.Fatalf("kind = %#v, want runtime", got)
	}
	if got := diag["likely_cause"]; got == nil || got == "" {
		t.Fatalf("likely_cause was not populated: %#v", diag)
	}
}

func TestBuildRunDiagnosticBackfillsBlankScriptLocation(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	a := &app{cwd: dir}
	env := output.Failure("run", output.Error{Code: "macro_failed", Message: "runtime modal", Source: "Main", Number: 438, Line: 12, Phase: "invoke_macro"})
	env.RunDiagnostic = map[string]any{
		"kind":     "runtime",
		"message":  []string{"Run-time error '438':", "Object doesn't support this property or method."},
		"dialog":   map[string]any{"title": "Microsoft Visual Basic"},
		"location": map[string]any{"module": "", "line": 0},
	}

	diag := a.buildRunDiagnostic(cfg, env)
	location := cliObjectMap(diag["location"])
	if got := location["module"]; got != "Main" {
		t.Fatalf("module = %#v, want Main: %#v", got, diag)
	}
	if got := location["line"]; got != 12 {
		t.Fatalf("line = %#v, want 12: %#v", got, diag)
	}
}

func TestBuildRunOptionsRejectsMalformedTypedArguments(t *testing.T) {
	cfg := config.Default()
	tests := []struct {
		literal string
		wantErr string
	}{
		{"int:not-a-number", "int values must parse"},
		{"bool:maybe", "bool values must be true or false"},
		{"hello", "expected type:value"},
		{"float:3.14", "unsupported --arg type prefix"},
		{"double:not-a-number", "double values must parse"},
		{"double:NaN", "double values must parse"},
		{"double:Inf", "double values must parse"},
		{"double:", "double values cannot be empty"},
		{"int:", "int values cannot be empty"},
		{"bool:", "bool values cannot be empty"},
	}
	for _, tt := range tests {
		t.Run(tt.literal, func(t *testing.T) {
			_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", Args: []string{tt.literal}})
			if err == nil {
				t.Fatalf("expected %q to fail", tt.literal)
			}
			if !strings.Contains(err.Error(), tt.wantErr) {
				t.Fatalf("error = %v, want substring %q", err, tt.wantErr)
			}
		})
	}
}

func TestBuildRunOptionsRejectsMalformedUIResponses(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", MsgBox: []string{"missing-delimiter"}})
	if err == nil || !strings.Contains(err.Error(), "expected id=value") {
		t.Fatalf("expected malformed msgbox response error, got %v", err)
	}
	_, err = buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", MsgBox: []string{"confirm-save=maybe"}})
	if err == nil || !strings.Contains(err.Error(), "unsupported msgbox result") {
		t.Fatalf("expected unsupported msgbox result error, got %v", err)
	}
	_, err = buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", MsgBox: []string{"confirm-save=yes", "confirm-save=no"}})
	if err == nil || !strings.Contains(err.Error(), "duplicate dialog id") {
		t.Fatalf("expected duplicate dialog id error, got %v", err)
	}
	_, err = buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", InputBox: []string{"customer name=John", "customer-name=Jane"}})
	if err == nil || !strings.Contains(err.Error(), "collides with") {
		t.Fatalf("expected normalized inputbox collision error, got %v", err)
	}
}

func TestBuildRunOptionsPreservesCommandStderr(t *testing.T) {
	cfg := config.Default()
	stderr := bytes.Buffer{}
	opts, err := buildRunOptionsForTest(cfg, runOptionsInput{Macro: "Main.Run", CommandOptions: excel.CommandOptions{Stderr: &stderr}})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Keepalive.Stderr != &stderr {
		t.Fatalf("command stderr = %#v, want %p", opts.Keepalive.Stderr, &stderr)
	}
}

func TestFilterAnalysisFindingsKeepsLegacyTraceFindings(t *testing.T) {
	findings := []analyze.Finding{
		{Code: "VBA104", Severity: "error"},
		{Code: "VBA105", Severity: "error"},
		{Code: "VBA106", Severity: "error"},
	}

	filtered := filterAnalysisFindings(findings, nil)
	if len(filtered) != len(findings) {
		t.Fatalf("expected findings to remain unchanged, got %+v", filtered)
	}
}

func skipWindowsPowerShellOnlyTest(t *testing.T) {
	t.Helper()
	if runtime.GOOS != "windows" {
		t.Skip("requires Windows PowerShell bridge")
	}
	if _, err := exec.LookPath("powershell"); err != nil {
		t.Skip("powershell not available")
	}
}

func withPowerShellBridge(args ...string) []string {
	return append([]string{"--bridge", "powershell"}, args...)
}

func writeTestPullScript(t *testing.T, root string, createModule bool) {
	t.Helper()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	createModuleLiteral := "$false"
	if createModule {
		createModuleLiteral = "$true"
	}
	script := fmt.Sprintf(`param(
  [string]$WorkbookPath,
  [string]$ModulesDir,
  [string]$ClassesDir,
  [string]$FormsDir,
  [string]$WorkbookDir
)
New-Item -ItemType Directory -Force -Path $ModulesDir, $ClassesDir, $FormsDir, $WorkbookDir | Out-Null
Get-ChildItem -LiteralPath $ModulesDir, $ClassesDir, $FormsDir, $WorkbookDir -File -ErrorAction SilentlyContinue | Remove-Item -Force
if (%s) {
  Set-Content -LiteralPath (Join-Path $ModulesDir 'Imported.bas') -Value "Attribute VB_Name = ""Imported""" -Encoding UTF8
}
@{
  status = 'ok'
  command = 'pull'
  logs = @('stub pull ok')
  workbook = @{ path = $WorkbookPath }
} | ConvertTo-Json -Compress
`, createModuleLiteral)
	if err := os.WriteFile(filepath.Join(scriptsDir, "pull.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
}

func writeTestNewScript(t *testing.T, root string) {
	t.Helper()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param(
  [string]$WorkbookPath
)
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $WorkbookPath) | Out-Null
Set-Content -LiteralPath $WorkbookPath -Value 'fake workbook' -Encoding UTF8
@{
  status = 'ok'
  command = 'new'
  logs = @('stub new ok')
  workbook = @{ path = $WorkbookPath }
} | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "new.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
}

func writeTestPushScript(t *testing.T, root string) {
	t.Helper()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param(
  [string]$WorkbookPath,
  [string]$ModulesDir,
  [string]$StatePath
)
$markerPath = Join-Path (Split-Path -Parent $StatePath) '..\push.called'
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $markerPath) | Out-Null
$names = @()
if (Test-Path -LiteralPath $ModulesDir) {
  $names = @(Get-ChildItem -LiteralPath $ModulesDir -File | Select-Object -ExpandProperty Name)
}
Set-Content -LiteralPath $markerPath -Value ($names -join [Environment]::NewLine) -Encoding UTF8
@{
  status = 'ok'
  command = 'push'
  logs = @('stub push ok')
  workbook = @{ path = $WorkbookPath }
} | ConvertTo-Json -Compress
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "push.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
}

func writeTestTestScript(t *testing.T, root string) {
	t.Helper()
	scriptsDir := filepath.Join(root, "scripts")
	if err := os.MkdirAll(scriptsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	script := `param(
  [string]$WorkbookPath,
  [string]$Filter = "",
  [string]$ModuleFilter = "",
  [string]$TagFilter = "",
  [string]$Visible = "false",
  [string]$RuntimeMode = "test",
  [string]$RuntimeSource = "command",
  [string]$MsgBoxResponsesJSON = "",
  [string]$InputResponsesJSON = "",
  [string]$FileDialogResponsesJSON = "",
  [string]$DebugStreamEnabled = "false",
  [string]$DebugStreamPipeName = "",
  [string]$UIStreamEnabled = "false",
  [string]$UIStreamRedactInput = "true",
  [string]$UIStreamPipeName = "",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)
@{
  status = 'ok'
  command = 'test'
  logs = @('stub test ok')
  workbook = @{ path = $WorkbookPath }
  tests = @()
} | ConvertTo-Json -Compress -Depth 5
`
	if err := os.WriteFile(filepath.Join(scriptsDir, "test.ps1"), []byte(script), 0o600); err != nil {
		t.Fatal(err)
	}
}

func TestRootCommandIncludesStatusCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"status"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "status" {
		t.Fatalf("expected status command, got %#v", cmd)
	}
}

func TestStatusJSONBaseline(t *testing.T) {
	dir := t.TempDir()
	createStatusCommandFixture(t, dir)

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "status"})

	if err := root.Execute(); err != nil {
		t.Fatalf("status command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Status  string         `json:"status"`
		Command string         `json:"command"`
		Project map[string]any `json:"project"`
		Session map[string]any `json:"session"`
		State   map[string]any `json:"state"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Status != output.StatusOK {
		t.Fatalf("status = %q, want %q", got.Status, output.StatusOK)
	}
	if got.Command != "status" {
		t.Fatalf("command = %q, want status", got.Command)
	}
	if got.Project == nil {
		t.Fatal("expected project in JSON envelope")
	}
	if got.Project["root"] == nil {
		t.Fatal("expected project.root in JSON envelope")
	}
	if got.Project["workbook_path"] == nil {
		t.Fatal("expected project.workbook_path in JSON envelope")
	}
	if got.State == nil {
		t.Fatal("expected state in JSON envelope")
	}
	if got.State["src_newer_than_workbook"] == nil {
		t.Fatal("expected state.src_newer_than_workbook in JSON envelope")
	}
	if got.Session == nil {
		t.Fatal("expected session in JSON envelope")
	}
}

func TestStatusJSONSourceNewerThanWorkbook(t *testing.T) {
	dir := t.TempDir()
	createStatusCommandFixture(t, dir)

	srcDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(srcDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\nEnd Sub\n"
	srcFile := filepath.Join(srcDir, "Main.bas")
	if err := os.WriteFile(srcFile, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	workbookPath := filepath.Join(dir, "build", "Book.xlsm")
	wbInfo, err := os.Stat(workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	wbMtime := wbInfo.ModTime()
	sourceLater := wbMtime.Add(time.Minute)
	if err := os.Chtimes(srcFile, sourceLater, sourceLater); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "status"})

	if err := root.Execute(); err != nil {
		t.Fatalf("status command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		State    map[string]any   `json:"state"`
		Warnings []map[string]any `json:"warnings"`
		Hints    []map[string]any `json:"hints"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.State == nil {
		t.Fatal("expected state in JSON envelope")
	}
	srcNewer, ok := got.State["src_newer_than_workbook"].(bool)
	if !ok || !srcNewer {
		t.Fatalf("expected src_newer_than_workbook=true, got %v", got.State["src_newer_than_workbook"])
	}
	if got.State["latest_source_modified_at"] == nil || got.State["latest_source_modified_at"].(string) == "" {
		t.Fatal("expected latest_source_modified_at in state")
	}

	foundSourceNewerWarning := false
	for _, w := range got.Warnings {
		if code, _ := w["code"].(string); code == "source_newer_than_workbook" {
			foundSourceNewerWarning = true
			break
		}
	}
	if !foundSourceNewerWarning {
		t.Fatalf("expected source_newer_than_workbook warning in envelope: %s", stdout.String())
	}

	foundPushHint := false
	for _, h := range got.Hints {
		if code, _ := h["code"].(string); code == "push_source" {
			foundPushHint = true
			break
		}
	}
	if !foundPushHint {
		t.Fatalf("expected push_source hint in envelope: %s", stdout.String())
	}
}

func createStatusCommandFixture(t *testing.T, dir string) {
	t.Helper()

	cfg := config.Default()
	cfg.Excel.Path = filepath.Join("build", "Book.xlsm")
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}

	buildDir := filepath.Join(dir, "build")
	if err := os.MkdirAll(buildDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(buildDir, "Book.xlsm")
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Data", "A1", "hello"); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}

	srcModules := filepath.Join(dir, "src", "modules")
	srcClasses := filepath.Join(dir, "src", "classes")
	srcForms := filepath.Join(dir, "src", "forms")
	srcWorkbook := filepath.Join(dir, "src", "workbook")
	for _, d := range []string{srcModules, srcClasses, srcForms, srcWorkbook} {
		if err := os.MkdirAll(d, 0o755); err != nil {
			t.Fatal(err)
		}
	}
}

func TestStatusJSONSessionFieldShape(t *testing.T) {
	dir := t.TempDir()
	createStatusCommandFixture(t, dir)

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "status"})

	if err := root.Execute(); err != nil {
		t.Fatalf("status command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Session map[string]any `json:"session"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if got.Session == nil {
		t.Fatal("expected session in JSON envelope")
	}

	if got.Session["running"] == nil {
		t.Fatal("expected session.running in status JSON (must match session status contract)")
	}
	running, ok := got.Session["running"].(bool)
	if !ok {
		t.Fatalf("session.running must be bool, got %T: %v", got.Session["running"], got.Session["running"])
	}
	if running {
		t.Log("session.running=true is valid when a real session is active")
	}

	if got.Session["workbook_open"] == nil {
		t.Fatal("expected session.workbook_open in status JSON (must match session status contract)")
	}
	wbOpen, ok := got.Session["workbook_open"].(bool)
	if !ok {
		t.Fatalf("session.workbook_open must be bool, got %T: %v", got.Session["workbook_open"], got.Session["workbook_open"])
	}
	if wbOpen {
		t.Log("session.workbook_open=true is valid when a real session is active")
	}

	if _, ok := got.Session["metadata"]; !ok {
		t.Fatal("expected session.metadata key in status JSON (must match session status contract)")
	}
}

func TestStatusWarningsExcludeInspectSpecificMessages(t *testing.T) {
	dir := t.TempDir()
	createStatusCommandFixture(t, dir)

	srcDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(srcDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(srcDir, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	a := &app{
		cwd:    dir,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "status"})

	if err := root.Execute(); err != nil {
		t.Fatalf("status command error = %v, exit = %d", err, output.ExitCode(err))
	}

	var got struct {
		Warnings []map[string]any `json:"warnings"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}

	forbiddenCodes := map[string]bool{
		"command_reads_saved_file": true,
	}
	for _, w := range got.Warnings {
		code, _ := w["code"].(string)
		if forbiddenCodes[code] {
			t.Fatalf("status output must not contain inspect-specific warning %q: have %+v", code, w)
		}
		if code == "live_session_dirty" {
			msg, _ := w["message"].(string)
			if strings.Contains(msg, "inspect") || strings.Contains(msg, "inspected") {
				t.Fatalf("status output must not contain inspect-specific wording in live_session_dirty: %s", msg)
			}
		}
	}
}

func TestBuildStatusWarningsAndHintsProducesStatusSpecificCodes(t *testing.T) {
	session := map[string]any{
		"active":        true,
		"save_required": true,
	}
	state := map[string]any{
		"src_newer_than_workbook":      true,
		"live_session_newer_than_disk": true,
	}

	warnings, hints := buildStatusWarningsAndHints(session, state)

	expectedWarnings := map[string]bool{
		"session_dirty":                false,
		"source_newer_than_workbook":   false,
		"live_session_newer_than_disk": false,
	}
	for _, w := range warnings {
		code, _ := w["code"].(string)
		if _, expected := expectedWarnings[code]; expected {
			expectedWarnings[code] = true
		}
	}
	for code, found := range expectedWarnings {
		if !found {
			t.Errorf("expected warning %q not found in %v", code, warnings)
		}
	}

	expectedHints := map[string]bool{
		"save_session":     false,
		"push_source":      false,
		"save_before_push": false,
	}
	for _, h := range hints {
		code, _ := h["code"].(string)
		if _, expected := expectedHints[code]; expected {
			expectedHints[code] = true
		}
	}
	for code, found := range expectedHints {
		if !found {
			t.Errorf("expected hint %q not found in %v", code, hints)
		}
	}
}

func TestBuildStatusStateBaselineDefaults(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Excel.Path = filepath.Join("build", "Book.xlsm")

	state := buildStatusState(root, cfg, filepath.Join(root, "build", "Book.xlsm"))

	if v, ok := state["workbook_saved"]; !ok {
		t.Error("workbook_saved missing from baseline state")
	} else if !v.(bool) {
		t.Error("workbook_saved default should be true")
	}

	if v, ok := state["src_newer_than_workbook"]; !ok {
		t.Error("src_newer_than_workbook missing from baseline state")
	} else if v.(bool) {
		t.Error("src_newer_than_workbook default should be false")
	}

	if v, ok := state["live_session_newer_than_disk"]; !ok {
		t.Error("live_session_newer_than_disk missing from baseline state")
	} else if v.(bool) {
		t.Error("live_session_newer_than_disk default should be false")
	}

	if v, ok := state["source_of_truth"]; !ok {
		t.Error("source_of_truth missing from baseline state")
	} else if v.(string) != "saved_workbook" {
		t.Errorf("source_of_truth default should be saved_workbook, got %v", v)
	}
}

func TestStatusWorkbookSavedDerivesFromSaveRequired(t *testing.T) {
	tests := []struct {
		name      string
		session   map[string]any
		wantSaved bool
	}{
		{
			name: "save_required true, dirty absent",
			session: map[string]any{
				"active":        true,
				"save_required": true,
			},
			wantSaved: false,
		},
		{
			name: "save_required false, dirty absent",
			session: map[string]any{
				"active":        true,
				"save_required": false,
			},
			wantSaved: true,
		},
		{
			name: "save_required true and dirty true",
			session: map[string]any{
				"active":        true,
				"save_required": true,
				"dirty":         true,
			},
			wantSaved: false,
		},
		{
			name: "save_required false and dirty false",
			session: map[string]any{
				"active":        true,
				"save_required": false,
				"dirty":         false,
			},
			wantSaved: true,
		},
		{
			name: "save_required true but session not active",
			session: map[string]any{
				"active":        false,
				"save_required": true,
				"dirty":         true,
			},
			wantSaved: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			active := boolValueForCLI(tt.session, "active")
			saved := true
			if active {
				saved = !boolValueForCLI(tt.session, "save_required")
			}
			if saved != tt.wantSaved {
				t.Errorf("workbook_saved = %v, want %v (save_required=%v, dirty=%v, active=%v)",
					saved, tt.wantSaved,
					boolValueForCLI(tt.session, "save_required"),
					boolValueForCLI(tt.session, "dirty"),
					active)
			}
		})
	}
}

func TestBuildStatusWarningsAndHintsSaveRequiredDirtyAbsent(t *testing.T) {
	session := map[string]any{
		"active":        true,
		"save_required": true,
	}
	state := map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": true,
	}

	warnings, hints := buildStatusWarningsAndHints(session, state)

	foundDirty := false
	for _, w := range warnings {
		if code, _ := w["code"].(string); code == "session_dirty" {
			foundDirty = true
			break
		}
	}
	if !foundDirty {
		t.Fatal("expected session_dirty warning when save_required=true (buildStatusWarningsAndHints must check save_required)")
	}

	foundSave := false
	for _, h := range hints {
		if code, _ := h["code"].(string); code == "save_session" {
			foundSave = true
			break
		}
	}
	if !foundSave {
		t.Fatal("expected save_session hint when save_required=true")
	}
}

func TestBuildStatusSessionBaselineStableFields(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("on Windows, buildStatusSession calls the probe; baseline only reached on non-Windows")
	}

	a := &app{cwd: t.TempDir()}
	cfg := config.Default()

	session := a.buildStatusSession(cfg, filepath.Join("path", "to", "workbook.xlsm"))

	if _, ok := session["running"]; !ok {
		t.Error("baseline session must include 'running' field for stable JSON schema on all platforms")
	}
	if _, ok := session["workbook_open"]; !ok {
		t.Error("baseline session must include 'workbook_open' field for stable JSON schema on all platforms")
	}
	if _, ok := session["metadata"]; !ok {
		t.Error("baseline session must include 'metadata' field for stable JSON schema on all platforms")
	}
}

func TestRootCommandIncludesProcessCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"process"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "process" {
		t.Fatalf("expected process command, got %#v", cmd)
	}
}

func TestRootCommandIncludesProcessListCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"process", "list"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "list" {
		t.Fatalf("expected process list command, got %#v", cmd)
	}
}

func TestRootCommandIncludesProcessCleanupCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"process", "cleanup"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "cleanup" {
		t.Fatalf("expected process cleanup command, got %#v", cmd)
	}
}

func TestProcessCleanupCommandDefinesFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"process", "cleanup"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"auto", "all", "yes"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected process cleanup command to define --%s", name)
		}
	}
}

func TestProcessListCommandArgCount(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"process", "list"})
	if err != nil {
		t.Fatal(err)
	}
	expected := cobra.NoArgs
	if cmd.Args == nil {
		t.Fatal("expected process list to define Args validator")
	}
	dummy := &cobra.Command{Args: expected}
	if reflect.TypeOf(cmd.Args) != reflect.TypeOf(dummy.Args) {
		t.Fatalf("expected process list to use NoArgs validator")
	}
}

func TestProcessCleanupValidatesExclusiveModes(t *testing.T) {
	err := validateProcessCleanupArgs("", false, false, false)
	if err == nil || !strings.Contains(err.Error(), "requires a PID") {
		t.Fatalf("expected PID requirement error, got %v", err)
	}

	err = validateProcessCleanupArgs("123", true, false, false)
	if err == nil || !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected pid/auto conflict error, got %v", err)
	}

	err = validateProcessCleanupArgs("123", false, true, false)
	if err == nil || !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected pid/all conflict error, got %v", err)
	}

	err = validateProcessCleanupArgs("", true, true, false)
	if err == nil || !strings.Contains(err.Error(), "cannot be combined") {
		t.Fatalf("expected auto/all conflict error, got %v", err)
	}
}

func TestProcessCleanupYesRequiresAllFlag(t *testing.T) {
	err := validateProcessCleanupArgs("", false, false, true)
	if err == nil || !strings.Contains(err.Error(), "--yes requires --all") {
		t.Fatalf("expected --yes requires --all error, got %v", err)
	}

	err = validateProcessCleanupArgs("123", false, false, true)
	if err == nil || !strings.Contains(err.Error(), "--yes requires --all") {
		t.Fatalf("expected --yes requires --all error with pid, got %v", err)
	}
}

func TestProcessCleanupValidPID(t *testing.T) {
	err := validateProcessCleanupArgs("0", false, false, false)
	if err == nil || !strings.Contains(err.Error(), "positive integer") {
		t.Fatalf("expected positive integer error for PID 0, got %v", err)
	}

	err = validateProcessCleanupArgs("-1", false, false, false)
	if err == nil || !strings.Contains(err.Error(), "positive integer") {
		t.Fatalf("expected positive integer error for negative PID, got %v", err)
	}

	err = validateProcessCleanupArgs("abc", false, false, false)
	if err == nil || !strings.Contains(err.Error(), "positive integer") {
		t.Fatalf("expected positive integer error for non-numeric PID, got %v", err)
	}

	err = validateProcessCleanupArgs("1", false, false, false)
	if err != nil {
		t.Fatalf("expected valid PID 1 to pass, got %v", err)
	}

	err = validateProcessCleanupArgs(" 123 ", false, false, false)
	if err != nil {
		t.Fatalf("expected whitespace-padded PID to pass validation, got %v", err)
	}
}

func TestResolvedVersionScriptsIncludesProcess(t *testing.T) {
	scripts := resolvedVersionScripts(t.TempDir())
	names := make([]string, 0, len(scripts))
	for _, script := range scripts {
		names = append(names, script.Command)
	}
	found := false
	for _, name := range names {
		if name == "process" {
			found = true
			break
		}
	}
	if !found {
		t.Fatalf("resolvedVersionScripts missing %q in %#v", "process", names)
	}
}

func TestConfirmPromptWritesToProvidedWriter(t *testing.T) {
	var buf bytes.Buffer
	result := confirmPrompt(strings.NewReader(""), &buf, "Test prompt? [y/N] ")
	if result {
		t.Fatalf("confirmPrompt should return false with no input")
	}
	got := buf.String()
	if got != "Test prompt? [y/N] " {
		t.Fatalf("confirmPrompt output = %q, want %q", got, "Test prompt? [y/N] ")
	}
}

func TestConfirmPromptAcceptsYes(t *testing.T) {
	var buf bytes.Buffer
	result := confirmPrompt(strings.NewReader("yes"), &buf, "Prompt ")
	if !result {
		t.Fatalf("confirmPrompt should return true for 'yes'")
	}
}

func TestConfirmPromptAcceptsY(t *testing.T) {
	var buf bytes.Buffer
	result := confirmPrompt(strings.NewReader("y"), &buf, "Prompt ")
	if !result {
		t.Fatalf("confirmPrompt should return true for 'y'")
	}
}

func TestConfirmPromptRejectsNo(t *testing.T) {
	var buf bytes.Buffer
	result := confirmPrompt(strings.NewReader("n"), &buf, "Prompt ")
	if result {
		t.Fatalf("confirmPrompt should return false for 'n'")
	}
}
