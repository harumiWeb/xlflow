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
	"strconv"
	"strings"
	"testing"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/wsl"
)

func TestDelegatedTopLevelCommands(t *testing.T) {
	delegated := []string{
		"attach", "check", "doctor", "edit", "export-image", "form", "init",
		"list", "macros", "new", "process", "pull", "push", "rollback",
		"recovery", "run", "runner", "save", "session", "status", "test", "type", "ui",
	}
	local := []string{
		"backup", "completion", "diff", "fmt", "generate", "impact", "inspect-gui", "lint",
		"inspect", "lsp", "module", "skill", "version",
	}
	for _, name := range delegated {
		if !shouldDelegateTopLevelCommand(name) {
			t.Errorf("%s should be delegated", name)
		}
	}
	for _, name := range local {
		if shouldDelegateTopLevelCommand(name) {
			t.Errorf("%s should remain local", name)
		}
	}
	if len(delegatedTopLevelCommands) != len(delegated) {
		t.Fatalf("delegated command count = %d, want %d", len(delegatedTopLevelCommands), len(delegated))
	}
}

func TestShouldDelegateInspectCommand(t *testing.T) {
	cases := []struct {
		name    string
		session bool
		want    bool
	}{
		{name: "form", want: true},
		{name: "workbook", want: false},
		{name: "workbook", session: true, want: true},
		{name: "sheets", want: false},
		{name: "sheets", session: true, want: true},
		{name: "range", want: false},
		{name: "range", session: true, want: true},
		{name: "used-range", want: false},
		{name: "used-range", session: true, want: true},
		{name: "cell", want: false},
		{name: "cell", session: true, want: true},
		{name: "unknown", want: false},
	}
	for _, tt := range cases {
		t.Run(fmt.Sprintf("%s/session=%v", tt.name, tt.session), func(t *testing.T) {
			cmd := &cobra.Command{Use: tt.name}
			cmd.Flags().Bool("session", false, "")
			if tt.session {
				if err := cmd.Flags().Set("session", "true"); err != nil {
					t.Fatalf("set session flag: %v", err)
				}
			}
			if got := shouldDelegateInspectCommand(cmd); got != tt.want {
				t.Fatalf("shouldDelegateInspectCommand() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestShouldDelegateTestListCommand(t *testing.T) {
	root := &cobra.Command{Use: "xlflow"}
	testCmd := &cobra.Command{Use: "test"}
	listCmd := &cobra.Command{Use: "list"}
	root.AddCommand(testCmd)
	testCmd.AddCommand(listCmd)

	if shouldDelegateCommand(listCmd, topLevelCommandName(listCmd)) {
		t.Fatal("test list should remain local because it only inspects source files")
	}
	if !shouldDelegateCommand(testCmd, topLevelCommandName(testCmd)) {
		t.Fatal("test should remain delegated because it runs workbook-backed tests")
	}
}

func TestShouldKeepFilePullLocalUnderWSL(t *testing.T) {
	root := (&app{}).rootCommand()
	pull, _, err := root.Find([]string{"pull"})
	if err != nil {
		t.Fatal(err)
	}
	if shouldDelegateCommand(pull, "pull") {
		t.Fatal("default auto pull should remain local until backend selection")
	}
	if err := pull.Flags().Set("backend", "file"); err != nil {
		t.Fatal(err)
	}
	if shouldDelegateCommand(pull, "pull") {
		t.Fatal("file pull should remain local")
	}
	if err := pull.Flags().Set("backend", "excel"); err != nil {
		t.Fatal(err)
	}
	if !shouldDelegateCommand(pull, "pull") {
		t.Fatal("explicit Excel pull should delegate")
	}
	if err := pull.Flags().Set("backend", "auto"); err != nil {
		t.Fatal(err)
	}
	if err := pull.Flags().Set("session", "true"); err != nil {
		t.Fatal(err)
	}
	if !shouldDelegateCommand(pull, "pull") {
		t.Fatal("auto pull with explicit session intent should delegate")
	}
}

func TestMatchingWSLLiveSessionUsesWindowsStatusProof(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()
	t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
	t.Setenv("WSL_HELPER_MODE", "session-status")
	isWSL = func() bool { return true }
	translateWSLPath = func(context.Context, string) (string, error) { return `C:\dev\project\build\Book.xlsm`, nil }
	resolveWindowsExecutable = func(context.Context) (string, string, error) { return "ignored.exe", `C:\tools\xlflow.exe`, nil }
	newDelegatedCommand = delegatedHelperCommand
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), []byte(`{"pid":42,"workbook_path":"C:\\dev\\project\\build\\Book.xlsm"}`), 0o600); err != nil {
		t.Fatal(err)
	}
	matched, err := (&app{cwd: dir}).matchingWSLLiveSession(t.Context(), filepath.Join(dir, "build", "Book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	if !matched {
		t.Fatal("matching Windows session status was not accepted")
	}
}

func TestDelegateAutoPullForwardsProbeWarning(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()
	t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
	t.Setenv("WSL_HELPER_MODE", "echo")
	isWSL = func() bool { return true }
	translateWSLPath = func(context.Context, string) (string, error) { return `C:\dev\project`, nil }
	resolveWindowsExecutable = func(context.Context) (string, string, error) { return "ignored.exe", `C:\tools\xlflow.exe`, nil }
	translateWSLArgs = func(_ context.Context, args []string) ([]string, error) { return append([]string{}, args...), nil }
	newDelegatedCommand = delegatedHelperCommand
	var stdout bytes.Buffer
	a := &app{cwd: t.TempDir(), rawArgs: []string{"--json", "pull"}, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	pull, _, err := root.Find([]string{"pull"})
	if err != nil {
		t.Fatal(err)
	}

	err = a.delegateAutoPullToWindows(pull, errors.New("session status unavailable"))
	if !errors.Is(err, errWSLDelegated) {
		t.Fatalf("delegation error = %v", err)
	}
	if !strings.Contains(stdout.String(), "PULL_WARNING=session status unavailable") {
		t.Fatalf("probe warning was not forwarded:\n%s", stdout.String())
	}
	if !strings.Contains(stdout.String(), envPullAutoProbeWarning+"/w") {
		t.Fatalf("WSLENV did not forward probe warning:\n%s", stdout.String())
	}
}

func TestShouldNotDelegateBackupSubcommands(t *testing.T) {
	root := &cobra.Command{Use: "xlflow"}
	backupCmd := &cobra.Command{Use: "backup"}
	pruneCmd := &cobra.Command{Use: "prune"}
	deleteCmd := &cobra.Command{Use: "delete"}
	root.AddCommand(backupCmd)
	backupCmd.AddCommand(pruneCmd, deleteCmd)

	for _, cmd := range []*cobra.Command{backupCmd, pruneCmd, deleteCmd} {
		if shouldDelegateCommand(cmd, topLevelCommandName(cmd)) {
			t.Fatalf("%s should remain local under WSL", cmd.CommandPath())
		}
	}
}

func TestShouldDelegateWorkbookBackedUserFormCommandsOnly(t *testing.T) {
	root := (&app{}).rootCommand()
	tests := []struct {
		path     []string
		delegate bool
	}{
		{path: []string{"form", "migrate", "sidecar"}, delegate: true},
		{path: []string{"form", "snapshot"}, delegate: true},
		{path: []string{"form", "build"}, delegate: true},
		{path: []string{"form", "apply"}, delegate: true},
		{path: []string{"form", "export-image"}, delegate: true},
		{path: []string{"inspect", "form"}, delegate: true},
		{path: []string{"list", "forms"}, delegate: true},
		{path: []string{"form", "new"}, delegate: false},
		{path: []string{"module", "rename"}, delegate: false},
		{path: []string{"module", "remove"}, delegate: false},
		{path: []string{"encoding", "convert"}, delegate: false},
	}
	for _, tt := range tests {
		cmd, _, err := root.Find(tt.path)
		if err != nil {
			t.Fatalf("find %v: %v", tt.path, err)
		}
		if got := shouldDelegateCommand(cmd, topLevelCommandName(cmd)); got != tt.delegate {
			t.Errorf("%s delegation = %t, want %t", cmd.CommandPath(), got, tt.delegate)
		}
	}
}

func TestUnsafeWorkbookPoliciesDelegateEvenWhenTopLevelWasHistoricallyLocal(t *testing.T) {
	root := (&app{}).rootCommand()
	for _, path := range [][]string{{"pack"}, {"diff"}, {"formulas", "pull"}, {"inspect", "workbook"}, {"recovery", "clear"}} {
		cmd, _, err := root.Find(path)
		if err != nil {
			t.Fatalf("find %v: %v", path, err)
		}
		if !shouldDelegateCommand(cmd, topLevelCommandName(cmd)) {
			t.Errorf("%s should delegate so the Windows process owns coordination", cmd.CommandPath())
		}
	}
}

func TestTopLevelCommandName(t *testing.T) {
	root := &cobra.Command{Use: "xlflow"}
	form := &cobra.Command{Use: "form"}
	build := &cobra.Command{Use: "build"}
	root.AddCommand(form)
	form.AddCommand(build)
	if got := topLevelCommandName(build); got != "form" {
		t.Fatalf("topLevelCommandName() = %q, want form", got)
	}
}

func TestDelegateWSLCommandPreservesStreamsArgumentsAndExitCode(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
	t.Setenv("WSL_HELPER_MODE", "echo")
	t.Setenv("WSL_HELPER_EXIT", "7")

	isWSL = func() bool { return true }
	translateWSLPath = func(context.Context, string) (string, error) { return `C:\dev\project`, nil }
	resolveWindowsExecutable = func(context.Context) (string, string, error) {
		return "ignored.exe", `C:\tools\xlflow.exe`, nil
	}
	translateWSLArgs = func(_ context.Context, args []string) ([]string, error) {
		got := append([]string{}, args...)
		got[len(got)-1] = `C:\dev\Book.xlsm`
		return got, nil
	}
	newDelegatedCommand = delegatedHelperCommand

	dir := t.TempDir()
	var stdout bytes.Buffer
	var stderr bytes.Buffer
	a := &app{
		cwd:       dir,
		rawArgs:   []string{"--json", "run", "/mnt/c/dev/Book.xlsm"},
		stdout:    &stdout,
		stderr:    &stderr,
		buildInfo: BuildInfo{Version: "1.2.3"},
	}
	root := &cobra.Command{Use: "xlflow"}
	run := &cobra.Command{Use: "run"}
	root.AddCommand(run)

	err := a.delegateWSLCommand(run)
	if code := output.ExitCode(err); code != 7 {
		t.Fatalf("exit code = %d, want 7; err=%v", code, err)
	}
	if !strings.Contains(stdout.String(), `ARGS=--json|run|C:\dev\Book.xlsm`) {
		t.Fatalf("stdout did not preserve translated args:\n%s", stdout.String())
	}
	if !strings.Contains(stdout.String(), "DELEGATED=1") {
		t.Fatalf("stdout did not report delegation marker:\n%s", stdout.String())
	}
	for _, want := range []string{"XLFLOW_WINDOWS_DELEGATED/w", "XLFLOW_EXCEL_BRIDGE/w", "XLFLOW_MODE/w", "XLFLOW_NO_UPDATE_CHECK/w"} {
		if !strings.Contains(stdout.String(), want) {
			t.Fatalf("stdout did not report WSLENV entry %q:\n%s", want, stdout.String())
		}
	}
	if !strings.Contains(stdout.String(), "CWD="+dir) {
		t.Fatalf("stdout did not report delegated cwd:\n%s", stdout.String())
	}
	if !strings.Contains(stderr.String(), "helper stderr") {
		t.Fatalf("stderr was not preserved:\n%s", stderr.String())
	}
}

func TestDelegateWSLCommandPreservesCoordinationWaitOptions(t *testing.T) {
	for _, args := range [][]string{
		{"--json", "--wait", "--wait-timeout", "45s", "push"},
		{"--json", "--wait", "--wait-timeout=45s", "push"},
	} {
		t.Run(strings.Join(args, "_"), func(t *testing.T) {
			restore := stubWSLDelegationGlobals(t)
			defer restore()
			t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
			t.Setenv("WSL_HELPER_MODE", "echo")
			isWSL = func() bool { return true }
			translateWSLPath = func(context.Context, string) (string, error) { return `C:\dev\project`, nil }
			resolveWindowsExecutable = func(context.Context) (string, string, error) { return "ignored.exe", `C:\tools\xlflow.exe`, nil }
			translateWSLArgs = func(_ context.Context, got []string) ([]string, error) { return append([]string{}, got...), nil }
			newDelegatedCommand = delegatedHelperCommand
			var stdout bytes.Buffer
			a := &app{cwd: t.TempDir(), rawArgs: args, stdout: &stdout, stderr: &bytes.Buffer{}, buildInfo: BuildInfo{Version: "1.2.3"}}
			root := a.rootCommand()
			root.SetArgs(args)
			err := root.Execute()
			if !errors.Is(err, errWSLDelegated) || output.ExitCode(err) != output.ExitSuccess {
				t.Fatalf("delegation error = %v, exit=%d", err, output.ExitCode(err))
			}
			if want := "ARGS=" + strings.Join(args, "|"); !strings.Contains(stdout.String(), want) {
				t.Fatalf("wait args were not preserved, want %q:\n%s", want, stdout.String())
			}
		})
	}
}

func TestDelegateWSLCommandStopsLocalRunEAfterSuccessfulDelegation(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
	t.Setenv("WSL_HELPER_MODE", "echo")

	isWSL = func() bool { return true }
	translateWSLPath = func(context.Context, string) (string, error) { return `C:\dev\project`, nil }
	resolveWindowsExecutable = func(context.Context) (string, string, error) {
		return "ignored.exe", `C:\tools\xlflow.exe`, nil
	}
	translateWSLArgs = func(_ context.Context, args []string) ([]string, error) { return args, nil }
	newDelegatedCommand = delegatedHelperCommand

	var stdout bytes.Buffer
	a := &app{
		cwd:       t.TempDir(),
		rawArgs:   []string{"new", "Book.xlsm", "--json"},
		stdout:    &stdout,
		buildInfo: BuildInfo{Version: "1.2.3"},
	}
	root := a.rootCommand()
	root.SetArgs([]string{"new", "Book.xlsm", "--json"})

	err := root.Execute()
	if !errors.Is(err, errWSLDelegated) {
		t.Fatalf("root.Execute() error = %v, want delegated sentinel", err)
	}
	if code := output.ExitCode(err); code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want 0", code)
	}
	if strings.Contains(stdout.String(), "refusing to overwrite") {
		t.Fatalf("local RunE appears to have continued:\n%s", stdout.String())
	}
	if !strings.Contains(stdout.String(), "ARGS=new|Book.xlsm|--json") {
		t.Fatalf("delegated helper did not run:\n%s", stdout.String())
	}
}

func TestDelegateWSLCommandSkipsWhenRawArgsUnavailable(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	isWSL = func() bool { return true }
	translateWSLPath = func(context.Context, string) (string, error) {
		t.Fatal("translateWSLPath must not be called without rawArgs")
		return "", nil
	}

	a := &app{cwd: "/tmp/xlflow-test"}
	root := &cobra.Command{Use: "xlflow"}
	run := &cobra.Command{Use: "run"}
	root.AddCommand(run)

	if err := a.delegateWSLCommand(run); err != nil {
		t.Fatalf("delegateWSLCommand() error = %v", err)
	}
}

func TestDelegateWSLCommandWritesStructuredResolutionFailure(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	isWSL = func() bool { return true }
	translateWSLPath = func(context.Context, string) (string, error) { return `C:\dev\project`, nil }
	resolveWindowsExecutable = func(context.Context) (string, string, error) {
		return "", "", &wsl.Error{Code: "windows_xlflow_not_found", Message: "not found"}
	}

	var stdout bytes.Buffer
	a := &app{
		json:      true,
		cwd:       "/mnt/c/dev/project",
		rawArgs:   []string{"--json", "run"},
		stdout:    &stdout,
		buildInfo: BuildInfo{Version: "1.2.3"},
	}
	root := &cobra.Command{Use: "xlflow"}
	run := &cobra.Command{Use: "run"}
	root.AddCommand(run)

	err := a.delegateWSLCommand(run)
	if code := output.ExitCode(err); code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	var env output.Envelope
	if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
		t.Fatalf("invalid JSON: %v\n%s", decodeErr, stdout.String())
	}
	if env.Error == nil || env.Error.Code != "windows_xlflow_not_found" {
		t.Fatalf("error = %+v", env.Error)
	}
}

func TestDelegateWSLDoctorReportsMissingWindowsExecutable(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	isWSL = func() bool { return true }
	wslDistroName = func() string { return "Ubuntu" }
	translateWSLPath = func(context.Context, string) (string, error) { return `C:\dev\project`, nil }
	resolveWindowsExecutable = func(context.Context) (string, string, error) {
		return "", "", &wsl.Error{Code: "windows_xlflow_not_found", Message: "not found"}
	}

	var stdout bytes.Buffer
	a := &app{
		json:      true,
		cwd:       "/mnt/c/dev/project",
		rawArgs:   []string{"doctor", "--json"},
		stdout:    &stdout,
		buildInfo: BuildInfo{Version: "1.2.3"},
	}
	root := &cobra.Command{Use: "xlflow"}
	doctor := &cobra.Command{Use: "doctor"}
	root.AddCommand(doctor)

	err := a.delegateWSLCommand(doctor)
	if code := output.ExitCode(err); code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	var env output.Envelope
	if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
		t.Fatalf("invalid JSON: %v\n%s", decodeErr, stdout.String())
	}
	windows := mapValue(mapValue(env.Diagnostics)["windows"])
	if windows["xlflow_found"] != false || windows["bridge_found"] != false || windows["excel_available"] != false {
		t.Fatalf("windows diagnostics = %#v", windows)
	}
	pathTranslation := mapValue(mapValue(env.Diagnostics)["path_translation"])
	if pathTranslation["supported"] != true {
		t.Fatalf("path translation diagnostics = %#v", pathTranslation)
	}
}

func TestRunDelegatedDoctorMergesWSLDiagnostics(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
	t.Setenv("WSL_HELPER_MODE", "doctor-success")
	newDelegatedCommand = delegatedHelperCommand
	wslDistroName = func() string { return "Ubuntu-24.04" }

	dir := t.TempDir()
	var stdout bytes.Buffer
	a := &app{
		json:      true,
		cwd:       dir,
		stdout:    &stdout,
		buildInfo: BuildInfo{Version: "1.2.3"},
	}
	err := a.runDelegatedDoctor(
		context.Background(),
		"ignored.exe",
		`C:\tools\xlflow.exe`,
		`C:\dev\project`,
		[]string{"doctor"},
	)
	if !errors.Is(err, errWSLDelegated) {
		t.Fatalf("runDelegatedDoctor() error = %v, want delegated sentinel", err)
	}
	if code := output.ExitCode(err); code != output.ExitSuccess {
		t.Fatalf("exit code = %d, want 0", code)
	}

	var env output.Envelope
	if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
		t.Fatalf("invalid JSON: %v\n%s", decodeErr, stdout.String())
	}
	diagnostics := mapValue(env.Diagnostics)
	host := mapValue(diagnostics["host"])
	if host["is_wsl"] != true || host["distro"] != "Ubuntu-24.04" {
		t.Fatalf("host diagnostics = %#v", host)
	}
	windows := mapValue(diagnostics["windows"])
	if windows["xlflow_found"] != true || windows["bridge_found"] != true || windows["excel_available"] != true {
		t.Fatalf("windows diagnostics = %#v", windows)
	}
	if windows["xlflow_version"] != "1.2.4" {
		t.Fatalf("windows xlflow version = %#v", windows["xlflow_version"])
	}
	pathTranslation := mapValue(diagnostics["path_translation"])
	if pathTranslation["windows_path"] != `C:\dev\project` || pathTranslation["supported"] != true {
		t.Fatalf("path translation diagnostics = %#v", pathTranslation)
	}
	warnings := anySlice(env.Warnings)
	if len(warnings) != 1 || mapValue(warnings[0])["code"] != "wsl_windows_version_mismatch" {
		t.Fatalf("warnings = %#v", warnings)
	}
}

func TestRunDelegatedDoctorPreservesWindowsFailureAndExitCode(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
	t.Setenv("WSL_HELPER_MODE", "doctor-failure")
	newDelegatedCommand = delegatedHelperCommand

	dir := t.TempDir()
	var stdout bytes.Buffer
	a := &app{
		json:      true,
		cwd:       dir,
		stdout:    &stdout,
		buildInfo: BuildInfo{Version: "dev"},
	}
	err := a.runDelegatedDoctor(context.Background(), "ignored.exe", `C:\tools\xlflow.exe`, `C:\dev\project`, []string{"--json", "doctor"})
	if code := output.ExitCode(err); code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d; err=%v", code, output.ExitEnvironment, err)
	}
	var env output.Envelope
	if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
		t.Fatalf("invalid JSON: %v\n%s", decodeErr, stdout.String())
	}
	if env.Error == nil || env.Error.Code != "excel_com_failure" {
		t.Fatalf("error = %+v", env.Error)
	}
	windows := mapValue(mapValue(env.Diagnostics)["windows"])
	if windows["excel_available"] != false {
		t.Fatalf("windows diagnostics = %#v", windows)
	}
}

func TestRunDelegatedDoctorRejectsInvalidJSON(t *testing.T) {
	restore := stubWSLDelegationGlobals(t)
	defer restore()

	t.Setenv("GO_WANT_WSL_HELPER_PROCESS", "1")
	t.Setenv("WSL_HELPER_MODE", "doctor-invalid")
	newDelegatedCommand = delegatedHelperCommand

	dir := t.TempDir()
	var stdout bytes.Buffer
	a := &app{
		json:      true,
		cwd:       dir,
		stdout:    &stdout,
		buildInfo: BuildInfo{Version: "dev"},
	}
	err := a.runDelegatedDoctor(context.Background(), "ignored.exe", `C:\tools\xlflow.exe`, `C:\dev\project`, []string{"doctor"})
	if code := output.ExitCode(err); code != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d", code, output.ExitEnvironment)
	}
	var env output.Envelope
	if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
		t.Fatalf("invalid JSON: %v\n%s", decodeErr, stdout.String())
	}
	if env.Error == nil || env.Error.Code != "windows_xlflow_execution_failed" {
		t.Fatalf("error = %+v", env.Error)
	}
	windows := mapValue(mapValue(env.Diagnostics)["windows"])
	if windows["xlflow_found"] != true || windows["excel_available"] != false {
		t.Fatalf("windows diagnostics = %#v", windows)
	}
}

func TestWSLDelegationHelperProcess(t *testing.T) {
	if os.Getenv("GO_WANT_WSL_HELPER_PROCESS") != "1" {
		return
	}
	args := helperArgs(os.Args)
	mode := os.Getenv("WSL_HELPER_MODE")
	switch {
	case containsArgument(args, "version"):
		fmt.Print(`{"status":"ok","command":"version","error":null,"logs":[],"version":{"version":"1.2.4","commit":"test","date":"today"}}`)
	case mode == "doctor-success":
		fmt.Print(`{"status":"ok","command":"doctor","error":null,"logs":[],"diagnostics":{"selected_bridge":"dotnet","excel":{"com_activation":true,"vbide_access":true}},"bridge":{"name":"xlflow-excel-bridge"}}`)
	case mode == "doctor-failure":
		fmt.Print(`{"status":"failed","command":"doctor","error":{"code":"excel_com_failure","message":"Excel COM activation failed"},"logs":[]}`)
		os.Exit(output.ExitEnvironment)
	case mode == "doctor-invalid":
		fmt.Print(`not-json`)
	case mode == "session-status":
		fmt.Print(`{"status":"ok","command":"session status","session":{"active":true,"running":true,"workbook_open":true,"workbook_path":"C:\\dev\\project\\build\\Book.xlsm"}}`)
	case mode == "echo":
		fmt.Printf("ARGS=%s\n", strings.Join(args, "|"))
		fmt.Printf("DELEGATED=%s\n", os.Getenv(wsl.EnvDelegated))
		fmt.Printf("WSLENV=%s\n", os.Getenv("WSLENV"))
		fmt.Printf("PULL_WARNING=%s\n", os.Getenv(envPullAutoProbeWarning))
		cwd, _ := os.Getwd()
		fmt.Printf("CWD=%s\n", cwd)
		fmt.Fprintln(os.Stderr, "helper stderr")
		exitCode, _ := strconv.Atoi(os.Getenv("WSL_HELPER_EXIT"))
		os.Exit(exitCode)
	default:
		fmt.Printf(`{"status":"failed","command":"doctor","error":{"code":"unexpected_helper_mode","message":%q},"logs":[]}`, mode)
		os.Exit(output.ExitEnvironment)
	}
	os.Exit(0)
}

func delegatedHelperCommand(ctx context.Context, _ string, args ...string) *exec.Cmd {
	helperArgs := []string{"-test.run=TestWSLDelegationHelperProcess", "--"}
	helperArgs = append(helperArgs, args...)
	return exec.CommandContext(ctx, os.Args[0], helperArgs...)
}

func helperArgs(args []string) []string {
	for i, arg := range args {
		if arg == "--" {
			return args[i+1:]
		}
	}
	return nil
}

func containsArgument(args []string, target string) bool {
	for _, arg := range args {
		if arg == target {
			return true
		}
	}
	return false
}

func stubWSLDelegationGlobals(t *testing.T) func() {
	t.Helper()
	originalIsWSL := isWSL
	originalDistro := wslDistroName
	originalResolve := resolveWindowsExecutable
	originalTranslatePath := translateWSLPath
	originalTranslateArgs := translateWSLArgs
	originalCommand := newDelegatedCommand
	return func() {
		isWSL = originalIsWSL
		wslDistroName = originalDistro
		resolveWindowsExecutable = originalResolve
		translateWSLPath = originalTranslatePath
		translateWSLArgs = originalTranslateArgs
		newDelegatedCommand = originalCommand
	}
}

func TestDelegationTestHelpers(t *testing.T) {
	if got := helperArgs([]string{"test", "--", "run", "Main"}); !reflect.DeepEqual(got, []string{"run", "Main"}) {
		t.Fatalf("helperArgs() = %#v", got)
	}
	if got := filepath.Base(os.Args[0]); got == "" {
		t.Fatal(errors.New("test executable path is empty"))
	}
}

func TestMergeWSLEnvPreservesExistingEntries(t *testing.T) {
	got := mergeWSLEnv("GOPATH/l:XLFLOW_MODE", "XLFLOW_MODE", "XLFLOW_EXCEL_BRIDGE")
	if got != "GOPATH/l:XLFLOW_MODE/w:XLFLOW_EXCEL_BRIDGE/w" {
		t.Fatalf("mergeWSLEnv() = %q", got)
	}
}

func TestMergeWSLEnvUpgradesExistingOneWayEntries(t *testing.T) {
	got := mergeWSLEnv("XLFLOW_MODE/u:XLFLOW_EXCEL_BRIDGE/up:XLFLOW_NO_UPDATE_CHECK/w", "XLFLOW_MODE", "XLFLOW_EXCEL_BRIDGE", "XLFLOW_NO_UPDATE_CHECK", "XLFLOW_WINDOWS_DELEGATED")
	want := "XLFLOW_MODE/uw:XLFLOW_EXCEL_BRIDGE/upw:XLFLOW_NO_UPDATE_CHECK/w:XLFLOW_WINDOWS_DELEGATED/w"
	if got != want {
		t.Fatalf("mergeWSLEnv() = %q", got)
	}
}
