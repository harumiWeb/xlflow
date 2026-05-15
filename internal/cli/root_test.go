package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

type stubReleaseChecker struct {
	release latestRelease
	err     error
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
	for _, name := range []string{"arg", "input", "save", "save-as", "trace", "headless", "interactive", "direct", "fast", "diagnostic", "gui-compile-errors", "session", "timeout", "keepalive", "keepalive-interval"} {
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
	}
	if flag.DefValue != "true" {
		t.Fatalf("diagnostic default = %q, want true", flag.DefValue)
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
		{"list", "forms"},
		{"pull"},
		{"push"},
		{"export-image"},
		{"edit", "cell"},
		{"edit", "range"},
		{"edit", "rows"},
		{"edit", "columns"},
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
		{"inspect", "workbook"},
		{"inspect", "range"},
		{"inspect", "used-range"},
		{"inspect", "cell"},
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
		{"test"},
		{"trace", "enable"},
		{"trace", "disable"},
		{"trace", "status"},
		{"trace", "inject"},
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
	for _, name := range []string{"session", "keepalive", "keepalive-interval"} {
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
	for _, name := range []string{"out", "session", "keepalive", "keepalive-interval"} {
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
	for _, name := range []string{"overwrite", "session", "no-save", "keepalive", "keepalive-interval"} {
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
	for _, name := range []string{"session", "no-save", "keepalive", "keepalive-interval"} {
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
	for _, name := range []string{"out", "initializer", "overwrite", "session", "keepalive", "keepalive-interval"} {
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
	for _, name := range []string{"runtime", "designer", "both", "initializer", "session", "keepalive", "keepalive-interval"} {
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
	for _, name := range []string{"sheet", "range", "out", "output-dir", "name", "format", "overwrite", "session", "keepalive", "keepalive-interval"} {
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
	for _, name := range []string{"sheet", "cell", "value", "formula", "fill", "events", "session", "keepalive", "keepalive-interval"} {
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
	opts, err := buildFormSnapshotOptions(" UserForm1 ", " artifacts\\UserForm1.form.yaml ", true, keepaliveFlags{enabled: true, interval: 7 * time.Second})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Inspect.Name != "UserForm1" || opts.Inspect.Basis != "designer" {
		t.Fatalf("inspect opts = %#v", opts.Inspect)
	}
	if !opts.Inspect.StrictDesigner {
		t.Fatalf("expected strict designer snapshot opts = %#v", opts.Inspect)
	}
	if !opts.Inspect.Session || !opts.Inspect.Keepalive.Keepalive || opts.Inspect.Keepalive.KeepaliveInterval != 7*time.Second {
		t.Fatalf("keepalive/session opts = %#v", opts.Inspect)
	}
	if opts.OutPath != "artifacts\\UserForm1.form.yaml" {
		t.Fatalf("out path = %q", opts.OutPath)
	}
}

func TestBuildFormSnapshotOptionsRejectsMissingRequirements(t *testing.T) {
	if _, err := buildFormSnapshotOptions("UserForm1", "", false, keepaliveFlags{}); err == nil || !strings.Contains(err.Error(), "--out is required") {
		t.Fatalf("expected out requirement error, got %v", err)
	}
	if _, err := buildFormSnapshotOptions("", "artifacts\\UserForm1.form.yaml", false, keepaliveFlags{}); err == nil || !strings.Contains(err.Error(), "form name is required") {
		t.Fatalf("expected form name error, got %v", err)
	}
}

func TestBuildFormExportImageOptionsValidatesAndNormalizes(t *testing.T) {
	opts, err := buildFormExportImageOptions(" UserForm1 ", " artifacts\\UserForm1.png ", " InitializeForm ", true, true, keepaliveFlags{enabled: true, interval: 7 * time.Second})
	if err != nil {
		t.Fatal(err)
	}
	if opts.Name != "UserForm1" || opts.OutPath != "artifacts\\UserForm1.png" || opts.Initializer != "InitializeForm" {
		t.Fatalf("unexpected form export-image opts: %#v", opts)
	}
	if !opts.Overwrite || !opts.Session || !opts.Keepalive.Keepalive || opts.Keepalive.KeepaliveInterval != 7*time.Second {
		t.Fatalf("unexpected keepalive/session opts: %#v", opts)
	}
}

func TestBuildFormExportImageOptionsRejectsMissingRequirements(t *testing.T) {
	if _, err := buildFormExportImageOptions("", "artifacts\\UserForm1.png", "", false, false, keepaliveFlags{}); err == nil || !strings.Contains(err.Error(), "form name is required") {
		t.Fatalf("expected form name error, got %v", err)
	}
	if _, err := buildFormExportImageOptions("UserForm1", "", "", false, false, keepaliveFlags{}); err == nil || !strings.Contains(err.Error(), "--out is required") {
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

	opts, err := buildFormWriteOptions(" build ", specPath, true, true, true, keepaliveFlags{enabled: true, interval: 7 * time.Second}, root)
	if err != nil {
		t.Fatal(err)
	}
	if opts.Action != "build" || opts.Spec.Form.Name != "UserForm1" {
		t.Fatalf("unexpected form write opts: %#v", opts)
	}
	if !opts.Overwrite || !opts.Session || !opts.NoSave {
		t.Fatalf("expected overwrite/session/no-save to be preserved: %#v", opts)
	}
	if !opts.Keepalive.Keepalive || opts.Keepalive.KeepaliveInterval != 7*time.Second {
		t.Fatalf("unexpected keepalive opts: %#v", opts.Keepalive)
	}
	if !strings.HasSuffix(opts.SpecInput.DisplayPath, "specs/UserForm1.form.yaml") {
		t.Fatalf("unexpected display path: %q", opts.SpecInput.DisplayPath)
	}
}

func TestBuildFormWriteOptionsRejectsInvalidRequirements(t *testing.T) {
	root := t.TempDir()
	if _, err := buildFormWriteOptions("apply", "missing.form.yaml", false, false, true, keepaliveFlags{}, root); err == nil || !strings.Contains(err.Error(), "--no-save requires --session") {
		t.Fatalf("expected no-save/session error, got %v", err)
	}
	if _, err := buildFormWriteOptions("noop", "missing.form.yaml", false, false, false, keepaliveFlags{}, root); err == nil || !strings.Contains(err.Error(), "unsupported form action") {
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

func TestInitCommandRendersWelcomeForInteractiveTerminal(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
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
	root.SetArgs([]string{"init", workbook})
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	for _, want := range []string{"Welcome to xlflow", "Version: 1.2.3", "copied workbook to build/Input.xlsm"} {
		if !strings.Contains(got, want) {
			t.Fatalf("interactive init output missing %q:\n%s", want, got)
		}
	}
	if strings.Index(got, "Welcome to xlflow") > strings.Index(got, "xlflow init") {
		t.Fatalf("expected welcome before command summary:\n%s", got)
	}
}

func TestInitCommandSkipsWelcomeForJSONOutput(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return true },
		stderrTerminal: func() bool { return true },
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "init", workbook})
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
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
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
	root.SetArgs([]string{"init", workbook})
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	if !strings.Contains(got, "Update available: v1.2.4") {
		t.Fatalf("interactive init output missing update notice:\n%s", got)
	}
}

func TestInitCommandSilentlySkipsFailedUpdateCheck(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
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
	root.SetArgs([]string{"init", workbook})
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	got := stdout.String()
	if strings.Contains(got, "Update available:") {
		t.Fatalf("interactive init output should skip failed update checks:\n%s", got)
	}
}

func TestInitCommandSkipsUpdateCheckWithFlag(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
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
	root.SetArgs([]string{"init", workbook, "--no-update-check"})
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if strings.Contains(stdout.String(), "Update available:") {
		t.Fatalf("interactive init output should skip update notice when --no-update-check is set:\n%s", stdout.String())
	}
}

func TestInitCommandSkipsUpdateCheckWithEnv(t *testing.T) {
	t.Setenv(noUpdateCheckEnvVar, "1")

	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
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
	root.SetArgs([]string{"init", workbook})
	if err := root.Execute(); err != nil {
		t.Fatalf("init command error = %v, exit = %d", err, output.ExitCode(err))
	}
	if strings.Contains(stdout.String(), "Update available:") {
		t.Fatalf("interactive init output should skip update notice when %s is set:\n%s", noUpdateCheckEnvVar, stdout.String())
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
	_, err := buildInspectFormOptions("UserForm1", "designer", "InitializeForm", false, keepaliveFlags{})
	if err == nil || !strings.Contains(err.Error(), "--initializer can only be used") {
		t.Fatalf("expected initializer validation error, got %v", err)
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

func TestBuildRunOptionsRejectsConflictingSaveFlags(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:hello"}, true, "build\\result.xlsm", false, false, false, false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--save and --save-as cannot be combined") {
		t.Fatalf("expected save conflict error, got %v", err)
	}
}

func TestBuildRunOptionsParsesTypedArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "", "fixtures\\Book.xlsm", []string{"string:hello", "int:7", "bool:true"}, false, "", true, true, false, false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
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
	opts, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:"}, false, "", false, false, false, false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err != nil {
		t.Fatal(err)
	}
	if len(opts.Args) != 1 || opts.Args[0].Type != "string" || opts.Args[0].Value != "" {
		t.Fatalf("run args = %#v", opts.Args)
	}
}

func TestBuildRunOptionsRejectsConflictingRunModes(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, true, true, false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--headless and --interactive") {
		t.Fatalf("expected run mode conflict error, got %v", err)
	}
}

func TestBuildRunOptionsRejectsDirectWithTraceOrArgs(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", true, false, false, true, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--direct cannot be combined with --trace") {
		t.Fatalf("expected direct trace conflict, got %v", err)
	}
	_, err = buildRunOptions(cfg, "Main.Run", "", []string{"string:hello"}, false, "", false, false, false, true, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--direct cannot be used with --arg") {
		t.Fatalf("expected direct arg conflict, got %v", err)
	}
}

func TestBuildRunOptionsRejectsDirectWithDiagnostic(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, false, false, true, false, true, true, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err == nil || !strings.Contains(err.Error(), "--gui-compile-errors") {
		t.Fatalf("expected direct diagnostic conflict, got %v", err)
	}
}

func TestBuildRunOptionsAutoDisablesDefaultDiagnosticForDirect(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, false, false, true, false, true, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
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

func TestBuildRunOptionsAllowsFastDiagnostic(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, false, false, false, true, true, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
	if err != nil {
		t.Fatal(err)
	}
	if !opts.Fast || !opts.Diagnostic {
		t.Fatalf("unexpected fast diagnostic options: %#v", opts)
	}
}

func TestBuildRunOptionsAllowsDirectWhenGUICompileErrorsOptOutIsSet(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, false, false, true, false, true, false, true, false, 5*time.Minute, false, defaultKeepaliveInterval)
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
		{"int:", "int values cannot be empty"},
		{"bool:", "bool values cannot be empty"},
	}
	for _, tt := range tests {
		t.Run(tt.literal, func(t *testing.T) {
			_, err := buildRunOptions(cfg, "Main.Run", "", []string{tt.literal}, false, "", false, false, false, false, false, false, false, false, false, 5*time.Minute, false, defaultKeepaliveInterval)
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
	opts, err := buildRunOptions(cfg, "Main.Run", "", nil, false, "", false, false, false, false, false, false, false, false, false, 5*time.Minute, true, 3*time.Second)
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

func TestFilterAnalysisFindingsIgnoresTraceHelperCodesForTraceRun(t *testing.T) {
	findings := []analyze.Finding{
		{Code: "VBA104", Severity: "error"},
		{Code: "VBA105", Severity: "error"},
		{Code: "VBA106", Severity: "error"},
	}

	filtered := filterAnalysisFindings(findings, ignoredRunPreflightAnalysisCodes(excel.RunOptions{Trace: true}))
	if len(filtered) != 1 || filtered[0].Code != "VBA104" {
		t.Fatalf("unexpected filtered findings: %+v", filtered)
	}
}

func TestIgnoredRunPreflightAnalysisCodesEmptyWithoutTrace(t *testing.T) {
	if got := ignoredRunPreflightAnalysisCodes(excel.RunOptions{}); got != nil {
		t.Fatalf("expected nil ignored codes without trace, got %+v", got)
	}
}
