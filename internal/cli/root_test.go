package cli

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

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
}

func TestRootCommandIncludesRunFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"run"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"arg", "input", "save", "save-as", "trace", "headless", "interactive", "direct", "fast", "session", "timeout", "keepalive", "keepalive-interval"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected run command to define --%s", name)
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

func TestRootCommandIncludesExcelCommandKeepaliveFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	for _, args := range [][]string{
		{"new"},
		{"doctor"},
		{"attach"},
		{"pull"},
		{"push"},
		{"trace", "inject"},
		{"macros"},
		{"test"},
		{"ui", "button", "add"},
		{"ui", "button", "list"},
		{"ui", "button", "remove"},
		{"check"},
	} {
		t.Run(strings.Join(args, " "), func(t *testing.T) {
			cmd, _, err := root.Find(args)
			if err != nil {
				t.Fatal(err)
			}
			for _, name := range []string{"keepalive", "keepalive-interval"} {
				if cmd.Flags().Lookup(name) == nil {
					t.Fatalf("expected %v command to define --%s", args, name)
				}
			}
		})
	}
}

func TestRootCommandDoesNotAddKeepaliveToNonExcelCommands(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	for _, args := range [][]string{
		{"init"},
		{"lint"},
		{"analyze"},
		{"diff"},
		{"inspect-gui"},
		{"skill", "install"},
	} {
		t.Run(strings.Join(args, " "), func(t *testing.T) {
			cmd, _, err := root.Find(args)
			if err != nil {
				t.Fatal(err)
			}
			if cmd.Flags().Lookup("keepalive") != nil || cmd.Flags().Lookup("keepalive-interval") != nil {
				t.Fatalf("expected %v command not to define keepalive flags", args)
			}
		})
	}
}

func TestAddKeepaliveFlagsUsesDefaultInterval(t *testing.T) {
	var flags keepaliveFlags
	cmd := &cobra.Command{Use: "sample"}
	addKeepaliveFlags(cmd, &flags)

	if err := cmd.Flags().Set("keepalive", "true"); err != nil {
		t.Fatal(err)
	}
	if err := cmd.Flags().Set("keepalive-interval", "3s"); err != nil {
		t.Fatal(err)
	}
	if !flags.enabled {
		t.Fatal("expected keepalive to be enabled")
	}
	if flags.interval != 3*time.Second {
		t.Fatalf("interval = %s, want 3s", flags.interval)
	}
}

func TestPullCommandRejectsInvalidKeepaliveIntervalBeforeConfigLoad(t *testing.T) {
	dir := t.TempDir()
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"pull", "--keepalive", "--keepalive-interval", "0s"})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("expected config exit for invalid interval, got err=%v exit=%d", err, output.ExitCode(err))
	}
	if !strings.Contains(err.Error(), "--keepalive-interval") {
		t.Fatalf("error = %v, want keepalive interval message", err)
	}
}

func TestUIButtonCommandRejectsInvalidKeepaliveIntervalWithUIButtonError(t *testing.T) {
	dir := t.TempDir()
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{
		"ui", "button", "add",
		"--sheet", "Menu",
		"--cell", "B2",
		"--text", "Run",
		"--macro", "Main.Run",
		"--keepalive",
		"--keepalive-interval", "0s",
	})

	err := root.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("expected config exit for invalid interval, got err=%v exit=%d", err, output.ExitCode(err))
	}
	if !strings.Contains(err.Error(), "--keepalive-interval") {
		t.Fatalf("error = %v, want keepalive interval message", err)
	}
}

func TestKeepaliveFlagsAreNotDuplicatedOnRun(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"run"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"keepalive", "keepalive-interval"} {
		if flag := cmd.Flags().Lookup(name); flag == nil {
			t.Fatalf("expected run command to define --%s", name)
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

func TestRootCommandIncludesTraceInjectCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	for _, name := range []string{"enable", "disable", "status", "clean", "inject"} {
		cmd, _, err := root.Find([]string{"trace", name})
		if err != nil {
			t.Fatal(err)
		}
		if cmd == nil || cmd.Name() != name {
			t.Fatalf("expected trace %s command, got %#v", name, cmd)
		}
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
	for _, name := range []string{"sheet", "cell", "text", "macro", "id", "width", "height", "create-sheet", "verify-macro"} {
		if add.Flags().Lookup(name) == nil {
			t.Fatalf("expected ui button add to define --%s", name)
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

func TestSkillInstallCommandInstallsProviderSkill(t *testing.T) {
	dir := t.TempDir()
	a := &app{cwd: dir}
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
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: dir}
	root := a.rootCommand()
	root.SetArgs([]string{"init", workbook, "--with-skill", "--agent", "codex"})
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

func TestBuildRunOptionsRejectsConflictingSaveFlags(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:hello"}, true, "build\\result.xlsm", false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--save and --save-as cannot be combined") {
		t.Fatalf("expected save conflict error, got %v", err)
	}
}

func TestBuildRunOptionsParsesTypedArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "", "fixtures\\Book.xlsm", []string{"string:hello", "int:7", "bool:true"}, false, "", true, true, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err != nil {
		t.Fatal(err)
	}

	want := []excel.RunArgument{
		{Type: "string", Value: "hello"},
		{Type: "int", Value: "7"},
		{Type: "bool", Value: "true"},
	}
	if opts.Macro != "Main.Run" {
		t.Fatalf("macro = %q, want Main.Run", opts.Macro)
	}
	if opts.WorkbookPath != "fixtures\\Book.xlsm" {
		t.Fatalf("workbook path = %q", opts.WorkbookPath)
	}
	if !opts.Trace {
		t.Fatal("expected trace flag to be enabled")
	}
	if opts.Mode != "headless" {
		t.Fatalf("mode = %q, want headless", opts.Mode)
	}
	if opts.Timeout != 5*time.Minute {
		t.Fatalf("timeout = %s", opts.Timeout)
	}
	if opts.Keepalive.Keepalive {
		t.Fatal("keepalive should default to disabled")
	}
	if opts.Keepalive.KeepaliveInterval != defaultKeepaliveInterval {
		t.Fatalf("keepalive interval = %s, want %s", opts.Keepalive.KeepaliveInterval, defaultKeepaliveInterval)
	}
	if !reflect.DeepEqual(opts.Args, want) {
		t.Fatalf("run args = %#v, want %#v", opts.Args, want)
	}
}

func TestBuildRunOptionsAllowsEmptyStringArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:"}, false, "", false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err != nil {
		t.Fatal(err)
	}
	if len(opts.Args) != 1 || opts.Args[0].Type != "string" || opts.Args[0].Value != "" {
		t.Fatalf("run args = %#v", opts.Args)
	}
}

func TestBuildRunOptionsRejectsConflictingRunModes(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, true, true, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--headless and --interactive") {
		t.Fatalf("expected run mode conflict error, got %v", err)
	}
}

func TestBuildRunOptionsRejectsDirectWithTraceOrArgs(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", true, false, false, true, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--direct cannot be combined with --trace") {
		t.Fatalf("expected direct trace conflict, got %v", err)
	}
	_, err = buildRunOptions(cfg, "Main.Run", "", []string{"string:hello"}, false, "", false, false, false, true, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--direct cannot be used with --arg") {
		t.Fatalf("expected direct arg conflict, got %v", err)
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
		{"int:", "int values cannot be empty"},
		{"bool:", "bool values cannot be empty"},
	}
	for _, tt := range tests {
		t.Run(tt.literal, func(t *testing.T) {
			_, err := buildRunOptions(cfg, "Main.Run", "", []string{tt.literal}, false, "", false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
			if err == nil {
				t.Fatalf("expected %q to fail", tt.literal)
			}
			if !strings.Contains(err.Error(), tt.wantErr) {
				t.Fatalf("error = %v, want substring %q", err, tt.wantErr)
			}
		})
	}
}

func TestBuildRunOptionsEnablesKeepalive(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, false, false, false, false, false, 5*time.Minute, true, 3*time.Second)
	if err != nil {
		t.Fatal(err)
	}
	if !opts.Keepalive.Keepalive {
		t.Fatal("expected keepalive to be enabled")
	}
	if opts.Keepalive.KeepaliveInterval != 3*time.Second {
		t.Fatalf("keepalive interval = %s, want 3s", opts.Keepalive.KeepaliveInterval)
	}
}

func TestBuildKeepaliveOptionsRejectsNonPositiveIntervalWhenEnabled(t *testing.T) {
	_, err := buildKeepaliveOptions(true, 0)
	if err == nil || !strings.Contains(err.Error(), "--keepalive-interval") {
		t.Fatalf("expected keepalive interval error, got %v", err)
	}
}
