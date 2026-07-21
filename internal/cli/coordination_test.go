package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestWorkbookContentionStopsCLILeafWithStructuredBusyFailure(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(rootDir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(rootDir, cfg.Excel.Path)
	if err != nil {
		t.Fatal(err)
	}
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity:      identity,
		Command:       "run",
		OperationKind: coordination.OperationExecute,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = lease.Release() }()

	var stdout bytes.Buffer
	a := &app{
		cwd:          rootDir,
		rawArgs:      []string{"--json", "push"},
		stdout:       &stdout,
		stderr:       &bytes.Buffer{},
		coordination: manager,
	}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "push"})
	err = root.Execute()
	if output.ExitCode(err) != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d (err=%v)", output.ExitCode(err), output.ExitEnvironment, err)
	}
	if !errors.Is(err, coordination.ErrWorkbookBusy) {
		t.Fatalf("error = %v, want workbook busy", err)
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatalf("decode JSON: %v\n%s", err, stdout.String())
	}
	if env.Error == nil || env.Error.Code != coordination.WorkbookBusyCode {
		t.Fatalf("error payload = %#v", env.Error)
	}
	details, ok := env.Error.Details.(map[string]any)
	if !ok {
		t.Fatalf("error details = %#v", env.Error.Details)
	}
	if details["workbook"] != identity.CanonicalPath || details["operation"] != "push" || details["retryable"] != true {
		t.Fatalf("busy details = %#v", details)
	}
	owner, ok := details["owner"].(map[string]any)
	if !ok || owner["command"] != "run" {
		t.Fatalf("owner details = %#v", details["owner"])
	}
}

func TestWorkbookContentionStopsEveryUserFormLeafBeforeHandler(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	tests := []struct {
		name      string
		operation coordination.CommandID
		args      func(string) []string
		output    func(string) string
	}{
		{name: "migrate sidecar", operation: "form.migrate.sidecar", args: func(string) []string { return []string{"--json", "form", "migrate", "sidecar"} }},
		{name: "snapshot", operation: "form.snapshot", args: func(root string) []string {
			return []string{"--json", "form", "snapshot", "UserForm1", "--out", filepath.Join(root, "snapshot.yaml")}
		}, output: func(root string) string { return filepath.Join(root, "snapshot.yaml") }},
		{name: "build", operation: "form.build", args: func(root string) []string {
			return []string{"--json", "form", "build", filepath.Join(root, "missing.yaml")}
		}},
		{name: "hidden apply", operation: "form.apply", args: func(root string) []string {
			return []string{"--json", "form", "apply", filepath.Join(root, "missing.yaml")}
		}},
		{name: "export image", operation: "form.export-image", args: func(root string) []string {
			return []string{"--json", "form", "export-image", "UserForm1", "--out", filepath.Join(root, "form.png")}
		}, output: func(root string) string { return filepath.Join(root, "form.png") }},
		{name: "inspect form", operation: "inspect.form", args: func(string) []string { return []string{"--json", "inspect", "form", "UserForm1", "--designer"} }},
		{name: "list forms", operation: "list.forms", args: func(string) []string { return []string{"--json", "list", "forms"} }},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			rootDir := t.TempDir()
			cfg := config.Default()
			if err := config.Write(filepath.Join(rootDir, config.FileName), cfg); err != nil {
				t.Fatal(err)
			}
			manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
			if err != nil {
				t.Fatal(err)
			}
			identity, err := coordination.NewWorkbookIdentity(rootDir, cfg.Excel.Path)
			if err != nil {
				t.Fatal(err)
			}
			owner, err := manager.Acquire(context.Background(), coordination.AcquireRequest{Identity: identity, Command: "run", OperationKind: coordination.OperationExecute, ResourceScope: coordination.ResourceWorkbook})
			if err != nil {
				t.Fatal(err)
			}
			defer func() { _ = owner.Release() }()

			args := tt.args(rootDir)
			var stdout bytes.Buffer
			a := &app{cwd: rootDir, rawArgs: args, stdout: &stdout, stderr: &bytes.Buffer{}, coordination: manager}
			root := a.rootCommand()
			root.SetArgs(args)
			runErr := root.Execute()
			if output.ExitCode(runErr) != output.ExitEnvironment || !errors.Is(runErr, coordination.ErrWorkbookBusy) {
				t.Fatalf("result = exit %d, err %v; want workbook_busy exit 3", output.ExitCode(runErr), runErr)
			}
			var env output.Envelope
			if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
				t.Fatalf("decode JSON: %v\n%s", err, stdout.String())
			}
			if env.Error == nil || env.Error.Code != coordination.WorkbookBusyCode {
				t.Fatalf("error payload = %#v", env.Error)
			}
			details, ok := env.Error.Details.(map[string]any)
			if !ok || details["operation"] != string(tt.operation) {
				t.Fatalf("busy details = %#v", env.Error.Details)
			}
			ownerDetails, ok := details["owner"].(map[string]any)
			if !ok || ownerDetails["command"] != "run" || ownerDetails["operation_kind"] != string(coordination.OperationExecute) {
				t.Fatalf("owner details = %#v", details["owner"])
			}
			if tt.output != nil {
				if _, statErr := os.Stat(tt.output(rootDir)); !os.IsNotExist(statErr) {
					t.Fatalf("handler side effect exists after contention: %v", statErr)
				}
			}
		})
	}
}

func TestCoordinationWaitOptionsValidation(t *testing.T) {
	tests := []struct {
		name     string
		args     []string
		wantCode string
	}{
		{name: "timeout requires wait", args: []string{"--json", "--wait-timeout", "1s", "push"}, wantCode: coordinationWaitArgsInvalidCode},
		{name: "timeout must be positive", args: []string{"--json", "--wait", "--wait-timeout", "0s", "push"}, wantCode: coordinationWaitArgsInvalidCode},
		{name: "timeout cannot be negative", args: []string{"--json", "--wait", "--wait-timeout", "-1s", "push"}, wantCode: coordinationWaitArgsInvalidCode},
		{name: "source command cannot wait", args: []string{"--json", "--wait", "lint"}, wantCode: coordinationWaitUnsupportedCode},
		{name: "source-only form new cannot wait", args: []string{"--json", "--wait", "form", "new", "UserForm1"}, wantCode: coordinationWaitUnsupportedCode},
		{name: "parallel observer cannot wait", args: []string{"--json", "--wait", "session", "status"}, wantCode: coordinationWaitUnsupportedCode},
		{name: "excel instance command cannot wait", args: []string{"--json", "--wait", "doctor"}, wantCode: coordinationWaitUnsupportedCode},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var stdout bytes.Buffer
			a := &app{cwd: t.TempDir(), rawArgs: tt.args, stdout: &stdout, stderr: &bytes.Buffer{}}
			root := a.rootCommand()
			root.SetArgs(tt.args)
			err := root.Execute()
			if output.ExitCode(err) != output.ExitConfig {
				t.Fatalf("exit code = %d, want %d (err=%v)", output.ExitCode(err), output.ExitConfig, err)
			}
			var env output.Envelope
			if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
				t.Fatalf("decode JSON: %v\n%s", decodeErr, stdout.String())
			}
			if env.Error == nil || env.Error.Code != tt.wantCode {
				t.Fatalf("error = %#v, want %s", env.Error, tt.wantCode)
			}
		})
	}
}

func TestCoordinationWaitDefaultIsThirtySeconds(t *testing.T) {
	a := &app{}
	root := a.rootCommand()
	if a.waitTimeout != 30*time.Second {
		t.Fatalf("wait timeout = %s, want 30s", a.waitTimeout)
	}
	if got := root.PersistentFlags().Lookup("wait-timeout").DefValue; got != "30s" {
		t.Fatalf("wait-timeout default = %q, want 30s", got)
	}
}

func TestCoordinationWaitOptionsAllowGeneratedCobraCommands(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{cwd: t.TempDir(), rawArgs: []string{"--wait", "help"}, stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetOut(&stdout)
	root.SetArgs(a.rawArgs)
	if err := root.Execute(); err != nil {
		t.Fatalf("generated help command rejected --wait: %v", err)
	}
}

func TestWorkbookCoordinationWaitsThenRunsHandlerOnce(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir, manager, identity, owner := setupHeldCoordinationLock(t, "run")
	var stdout, stderr bytes.Buffer
	a := &app{cwd: rootDir, json: true, wait: true, waitTimeout: time.Second, stdout: &stdout, stderr: &stderr, coordination: manager}
	time.AfterFunc(150*time.Millisecond, func() { _ = owner.Release() })
	runs := 0
	err := a.withWorkbookCoordination(context.Background(), "push", []string{identity.CanonicalPath}, func() error {
		runs++
		return nil
	})
	if err != nil {
		t.Fatalf("waited command failed: %v", err)
	}
	if runs != 1 {
		t.Fatalf("handler runs = %d, want 1", runs)
	}
	if stdout.Len() != 0 || stderr.Len() != 0 {
		t.Fatalf("JSON wait emitted progress: stdout=%q stderr=%q", stdout.String(), stderr.String())
	}
}

func TestDesignerCoordinationWaitsThenPublishesDesignerOwner(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir, manager, identity, owner := setupHeldCoordinationLock(t, "form.build")
	var stdout, stderr bytes.Buffer
	a := &app{cwd: rootDir, json: true, wait: true, waitTimeout: time.Second, stdout: &stdout, stderr: &stderr, coordination: manager}
	time.AfterFunc(150*time.Millisecond, func() { _ = owner.Release() })
	runs := 0
	err := a.withWorkbookCoordination(context.Background(), "form.snapshot", []string{identity.CanonicalPath}, func() error {
		runs++
		probe, probeErr := manager.Probe(context.Background(), identity)
		if probeErr != nil {
			return probeErr
		}
		if !probe.Busy || probe.Owner == nil || probe.Owner.Command != "form.snapshot" || probe.Owner.OperationKind != coordination.OperationDesigner {
			t.Fatalf("designer owner = %#v", probe.Owner)
		}
		return nil
	})
	if err != nil {
		t.Fatalf("waited Designer command failed: %v", err)
	}
	if runs != 1 || stdout.Len() != 0 || stderr.Len() != 0 {
		t.Fatalf("runs=%d stdout=%q stderr=%q", runs, stdout.String(), stderr.String())
	}
}

func TestDesignerCoordinationConflictsWithDesignerExecutionAndSourceUpdates(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	for _, ownerCommand := range []coordination.CommandID{"form.build", "run", "test", "push"} {
		t.Run(string(ownerCommand), func(t *testing.T) {
			rootDir, manager, identity, owner := setupHeldCoordinationLock(t, ownerCommand)
			defer func() { _ = owner.Release() }()
			var stdout bytes.Buffer
			a := &app{cwd: rootDir, json: true, stdout: &stdout, stderr: &bytes.Buffer{}, coordination: manager}
			runs := 0
			err := a.withWorkbookCoordination(context.Background(), "form.snapshot", []string{identity.CanonicalPath}, func() error { runs++; return nil })
			if output.ExitCode(err) != output.ExitEnvironment || !errors.Is(err, coordination.ErrWorkbookBusy) {
				t.Fatalf("result = exit %d, err %v; want workbook_busy exit 3", output.ExitCode(err), err)
			}
			if runs != 0 {
				t.Fatalf("handler runs = %d, want 0", runs)
			}
			var env output.Envelope
			if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
				t.Fatalf("decode JSON: %v\n%s", decodeErr, stdout.String())
			}
			if env.Error == nil || env.Error.Code != coordination.WorkbookBusyCode {
				t.Fatalf("error payload = %#v", env.Error)
			}
			details, ok := env.Error.Details.(map[string]any)
			ownerDetails, ownerOK := details["owner"].(map[string]any)
			if !ok || !ownerOK || details["operation"] != "form.snapshot" || ownerDetails["command"] != string(ownerCommand) {
				t.Fatalf("busy details = %#v", env.Error.Details)
			}
		})
	}
}

func TestDesignerCoordinationTimeoutAndCancellationNeverRunHandler(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	tests := []struct {
		name    string
		code    string
		context func() context.Context
		timeout time.Duration
	}{
		{name: "timeout", code: coordination.WorkbookBusyTimeoutCode, context: func() context.Context { return context.Background() }, timeout: 250 * time.Millisecond},
		{name: "cancel", code: coordination.WorkbookBusyCancelledCode, context: func() context.Context {
			ctx, cancel := context.WithCancel(context.Background())
			time.AfterFunc(150*time.Millisecond, cancel)
			return ctx
		}, timeout: time.Second},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			rootDir, manager, identity, owner := setupHeldCoordinationLock(t, "run")
			defer func() { _ = owner.Release() }()
			var stdout, stderr bytes.Buffer
			a := &app{cwd: rootDir, json: true, wait: true, waitTimeout: tt.timeout, stdout: &stdout, stderr: &stderr, coordination: manager}
			runs := 0
			err := a.withWorkbookCoordination(tt.context(), "form.snapshot", []string{identity.CanonicalPath}, func() error { runs++; return nil })
			assertCoordinationWaitFailure(t, err, stdout.Bytes(), tt.code, tt.timeout.String())
			if runs != 0 || stderr.Len() != 0 {
				t.Fatalf("handler runs=%d stderr=%q", runs, stderr.String())
			}
		})
	}
}

func TestWorkbookCoordinationWaitTimeoutIsStructuredAndJSONPure(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir, manager, identity, owner := setupHeldCoordinationLock(t, "run")
	defer func() { _ = owner.Release() }()
	var stdout, stderr bytes.Buffer
	a := &app{cwd: rootDir, json: true, wait: true, waitTimeout: 75 * time.Millisecond, stdout: &stdout, stderr: &stderr, coordination: manager}
	err := a.withWorkbookCoordination(context.Background(), "push", []string{identity.CanonicalPath}, func() error {
		t.Fatal("handler must not run after timeout")
		return nil
	})
	assertCoordinationWaitFailure(t, err, stdout.Bytes(), coordination.WorkbookBusyTimeoutCode, "75ms")
	if stderr.Len() != 0 {
		t.Fatalf("JSON wait emitted stderr progress: %q", stderr.String())
	}
}

func TestWorkbookCoordinationWaitCancellationIsStructured(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir, manager, identity, owner := setupHeldCoordinationLock(t, "run")
	defer func() { _ = owner.Release() }()
	var stdout, stderr bytes.Buffer
	a := &app{cwd: rootDir, json: true, wait: true, waitTimeout: time.Second, stdout: &stdout, stderr: &stderr, coordination: manager}
	ctx, cancel := context.WithCancel(context.Background())
	time.AfterFunc(50*time.Millisecond, cancel)
	err := a.withWorkbookCoordination(ctx, "push", []string{identity.CanonicalPath}, func() error {
		t.Fatal("handler must not run after cancellation")
		return nil
	})
	assertCoordinationWaitFailure(t, err, stdout.Bytes(), coordination.WorkbookBusyCancelledCode, "1s")
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("error = %v, want context cancellation", err)
	}
	if stderr.Len() != 0 {
		t.Fatalf("JSON cancellation emitted stderr progress: %q", stderr.String())
	}
}

func TestWorkbookCoordinationHumanWaitMessageAppearsOnlyAfterContention(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir, manager, identity, owner := setupHeldCoordinationLock(t, "run")
	defer func() { _ = owner.Release() }()
	var stdout, stderr bytes.Buffer
	a := &app{cwd: rootDir, wait: true, waitTimeout: 50 * time.Millisecond, stdout: &stdout, stderr: &stderr, coordination: manager}
	_ = a.withWorkbookCoordination(context.Background(), "push", []string{identity.CanonicalPath}, func() error { return nil })
	if !strings.Contains(stderr.String(), "Waiting up to 50ms") || !strings.Contains(stderr.String(), identity.CanonicalPath) {
		t.Fatalf("wait message = %q", stderr.String())
	}
	if strings.Count(stderr.String(), "\n") != 1 {
		t.Fatalf("wait message should be exactly one line: %q", stderr.String())
	}
}

func TestWorkbookCoordinationWithoutContentionEmitsNoWaitMessage(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	var stderr bytes.Buffer
	// The wait deadline also covers lock metadata publication, which can be slow
	// on loaded Windows CI runners even without workbook contention.
	a := &app{cwd: rootDir, wait: true, waitTimeout: 5 * time.Second, stdout: &bytes.Buffer{}, stderr: &stderr, coordination: manager}
	if err := a.withWorkbookCoordination(context.Background(), "push", []string{filepath.Join(rootDir, "book.xlsm")}, func() error { return nil }); err != nil {
		t.Fatal(err)
	}
	if stderr.Len() != 0 {
		t.Fatalf("uncontended wait emitted progress: %q", stderr.String())
	}
}

func TestWorkbookCoordinationNonWaitCancellationKeepsSetupFailureCode(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{cwd: rootDir, json: true, stdout: &stdout, stderr: &bytes.Buffer{}, coordination: manager}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	runErr := a.withWorkbookCoordination(ctx, "push", []string{filepath.Join(rootDir, "book.xlsm")}, func() error { return nil })
	if output.ExitCode(runErr) != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d (err=%v)", output.ExitCode(runErr), output.ExitEnvironment, runErr)
	}
	var env output.Envelope
	if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
		t.Fatalf("decode JSON: %v\n%s", decodeErr, stdout.String())
	}
	if env.Error == nil || env.Error.Code != "coordination_acquire_failed" {
		t.Fatalf("non-wait cancellation error = %#v", env.Error)
	}
}

func TestWorkbookCoordinationTimeoutReleasesEarlierMultiTargetLease(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	firstPath := filepath.Join(rootDir, "first.xlsm")
	secondPath := filepath.Join(rootDir, "second.xlsm")
	firstIdentity, _ := coordination.NewWorkbookIdentity(rootDir, firstPath)
	secondIdentity, _ := coordination.NewWorkbookIdentity(rootDir, secondPath)
	if firstIdentity.LockID > secondIdentity.LockID {
		firstPath, secondPath = secondPath, firstPath
		firstIdentity, secondIdentity = secondIdentity, firstIdentity
	}
	descriptor, _ := coordination.Lookup("diff")
	owner, err := manager.Acquire(context.Background(), coordination.AcquireRequest{Identity: secondIdentity, Command: "run", OperationKind: coordination.OperationExecute, ResourceScope: coordination.ResourceWorkbook})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = owner.Release() }()
	var stdout bytes.Buffer
	a := &app{cwd: rootDir, json: true, wait: true, waitTimeout: 75 * time.Millisecond, stdout: &stdout, stderr: &bytes.Buffer{}, coordination: manager}
	err = a.withWorkbookCoordination(context.Background(), descriptor.ID, []string{firstPath, secondPath}, func() error { return nil })
	assertCoordinationWaitFailure(t, err, stdout.Bytes(), coordination.WorkbookBusyTimeoutCode, "75ms")
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{Identity: firstIdentity, Command: descriptor.ID, OperationKind: descriptor.Policy.OperationKind, ResourceScope: descriptor.Policy.ResourceScope})
	if err != nil {
		t.Fatalf("earlier target remained locked: %v", err)
	}
	_ = lease.Release()
}

func TestWorkbookCoordinationUsesOneTimeoutBudgetAcrossMultipleTargets(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	firstPath := filepath.Join(rootDir, "first.xlsm")
	secondPath := filepath.Join(rootDir, "second.xlsm")
	firstIdentity, _ := coordination.NewWorkbookIdentity(rootDir, firstPath)
	secondIdentity, _ := coordination.NewWorkbookIdentity(rootDir, secondPath)
	if firstIdentity.LockID > secondIdentity.LockID {
		firstPath, secondPath = secondPath, firstPath
		firstIdentity, secondIdentity = secondIdentity, firstIdentity
	}
	request := func(identity coordination.WorkbookIdentity) coordination.AcquireRequest {
		return coordination.AcquireRequest{Identity: identity, Command: "run", OperationKind: coordination.OperationExecute, ResourceScope: coordination.ResourceWorkbook}
	}
	firstOwner, err := manager.Acquire(context.Background(), request(firstIdentity))
	if err != nil {
		t.Fatal(err)
	}
	secondOwner, err := manager.Acquire(context.Background(), request(secondIdentity))
	if err != nil {
		_ = firstOwner.Release()
		t.Fatal(err)
	}
	defer func() { _ = firstOwner.Release() }()
	defer func() { _ = secondOwner.Release() }()

	// Keep a wide enough gap between one shared deadline (about 1.5s) and a
	// mistakenly restarted second deadline (about 2.4s) for Windows CI
	// scheduling delays. The former 220ms budget was too narrow and flaky.
	const waitBudget = 1500 * time.Millisecond
	time.AfterFunc(900*time.Millisecond, func() { _ = firstOwner.Release() })
	var stdout bytes.Buffer
	a := &app{cwd: rootDir, json: true, wait: true, waitTimeout: waitBudget, stdout: &stdout, stderr: &bytes.Buffer{}, coordination: manager}
	started := time.Now()
	err = a.withWorkbookCoordination(context.Background(), "diff", []string{firstPath, secondPath}, func() error {
		t.Fatal("handler must not run after the shared timeout expires")
		return nil
	})
	elapsed := time.Since(started)
	assertCoordinationWaitFailure(t, err, stdout.Bytes(), coordination.WorkbookBusyTimeoutCode, waitBudget.String())
	if elapsed < 1200*time.Millisecond {
		t.Fatalf("multi-target acquisition timed out too early after %s", elapsed)
	}
	if elapsed >= 2100*time.Millisecond {
		t.Fatalf("multi-target acquisition took %s; timeout budget appears to have restarted per target", elapsed)
	}
}

func TestWorkbookCoordinationTimeoutDoesNotCoverHandlerRuntime(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: rootDir, wait: true, waitTimeout: 500 * time.Millisecond, stdout: &bytes.Buffer{}, stderr: &bytes.Buffer{}, coordination: manager}
	err = a.withWorkbookCoordination(context.Background(), "push", []string{filepath.Join(rootDir, "book.xlsm")}, func() error {
		time.Sleep(600 * time.Millisecond)
		return nil
	})
	if err != nil {
		t.Fatalf("handler runtime was covered by wait timeout: %v", err)
	}
}

func TestWorkbookRecoveryBlocksWithoutWaitingAndDoesNotRunHandler(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(rootDir, filepath.Join(rootDir, "book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	owner, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity:      identity,
		Command:       "run",
		OperationKind: coordination.OperationExecute,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := owner.PublishRecovery(coordination.RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
		ExcelPID:  12345,
	}); err != nil {
		t.Fatal(err)
	}
	if err := owner.Release(); err != nil {
		t.Fatal(err)
	}

	var stdout bytes.Buffer
	var stderr bytes.Buffer
	a := &app{
		cwd:          rootDir,
		json:         true,
		wait:         true,
		waitTimeout:  5 * time.Second,
		stdout:       &stdout,
		stderr:       &stderr,
		coordination: manager,
	}
	runs := 0
	err = a.withWorkbookCoordination(context.Background(), "push", []string{identity.CanonicalPath}, func() error {
		runs++
		return nil
	})
	if output.ExitCode(err) != output.ExitEnvironment || runs != 0 {
		t.Fatalf("err=%v exit=%d runs=%d", err, output.ExitCode(err), runs)
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatal(err)
	}
	if env.Error == nil || env.Error.Code != coordination.WorkbookRecoveryRequiredCode || env.Error.Phase != "coordination.recovery" {
		t.Fatalf("error = %#v", env.Error)
	}
	details, ok := env.Error.Details.(map[string]any)
	if !ok || details["reason"] != "vba_may_still_be_running" || details["retryable"] != false || details["wait_will_resolve"] != false {
		t.Fatalf("details = %#v", env.Error.Details)
	}
	if strings.Contains(stderr.String(), "Waiting up to") {
		t.Fatalf("recovery unexpectedly entered wait path: %s", stderr.String())
	}
}

func TestWorkbookRecoveryReadFailureIsFailClosed(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(rootDir, filepath.Join(rootDir, "book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.Mkdir(filepath.Join(manager.StateDir(), identity.LockID+".recovery.json"), 0o700); err != nil {
		t.Fatal(err)
	}
	var stdout bytes.Buffer
	a := &app{
		cwd:          rootDir,
		json:         true,
		stdout:       &stdout,
		stderr:       &bytes.Buffer{},
		coordination: manager,
	}
	runs := 0
	err = a.withWorkbookCoordination(context.Background(), "push", []string{identity.CanonicalPath}, func() error {
		runs++
		return nil
	})
	if output.ExitCode(err) != output.ExitEnvironment || runs != 0 {
		t.Fatalf("err=%v exit=%d runs=%d", err, output.ExitCode(err), runs)
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatal(err)
	}
	if env.Error == nil || env.Error.Code != coordination.RecoveryCheckFailedCode || env.Error.Phase != "coordination.recovery" {
		t.Fatalf("error = %#v", env.Error)
	}
}

func setupHeldCoordinationLock(t *testing.T, command coordination.CommandID) (string, *coordination.Manager, coordination.WorkbookIdentity, *coordination.Lease) {
	t.Helper()
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(rootDir, filepath.Join(rootDir, "book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	descriptor, err := coordination.Lookup(command)
	if err != nil {
		t.Fatal(err)
	}
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{Identity: identity, Command: command, OperationKind: descriptor.Policy.OperationKind, ResourceScope: descriptor.Policy.ResourceScope})
	if err != nil {
		t.Fatal(err)
	}
	return rootDir, manager, identity, lease
}

func assertCoordinationWaitFailure(t *testing.T, err error, body []byte, code, timeout string) {
	t.Helper()
	if output.ExitCode(err) != output.ExitEnvironment {
		t.Fatalf("exit code = %d, want %d (err=%v)", output.ExitCode(err), output.ExitEnvironment, err)
	}
	var env output.Envelope
	if decodeErr := json.Unmarshal(body, &env); decodeErr != nil {
		t.Fatalf("decode JSON: %v\n%s", decodeErr, body)
	}
	if env.Error == nil || env.Error.Code != code || env.Error.Phase != "coordination.acquire" {
		t.Fatalf("error payload = %#v", env.Error)
	}
	details, ok := env.Error.Details.(map[string]any)
	if !ok || details["wait_timeout"] != timeout || details["retryable"] != true {
		t.Fatalf("wait details = %#v", env.Error.Details)
	}
}

func TestCoordinationTargetsResolveMultiWorkbookCommands(t *testing.T) {
	rootDir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(rootDir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	a := &app{cwd: rootDir}
	root := a.rootCommand()

	diffCmd, _, err := root.Find([]string{"diff"})
	if err != nil {
		t.Fatal(err)
	}
	targets, ok := a.coordinationTargets(diffCmd, []string{"before.xlsm", "after.xlsm"}, "diff")
	if !ok || len(targets) != 2 || targets[0] != filepath.Join(rootDir, "before.xlsm") || targets[1] != filepath.Join(rootDir, "after.xlsm") {
		t.Fatalf("diff targets = %#v, resolved=%v", targets, ok)
	}

	initCmd, _, err := root.Find([]string{"init"})
	if err != nil {
		t.Fatal(err)
	}
	source := filepath.Join(rootDir, "imports", "Sample.xlsm")
	targets, ok = a.coordinationTargets(initCmd, []string{source}, "init")
	if !ok || len(targets) != 2 || targets[0] != source || targets[1] != filepath.Join(rootDir, "build", "Sample.xlsm") {
		t.Fatalf("init targets = %#v, resolved=%v", targets, ok)
	}

	editCmd, _, err := root.Find([]string{"edit", "cell"})
	if err != nil {
		t.Fatal(err)
	}
	override := filepath.Join(rootDir, "other", "Edit.xlsm")
	targets, ok = a.coordinationTargets(editCmd, []string{override}, "edit.cell")
	if !ok || len(targets) != 1 || targets[0] != override {
		t.Fatalf("edit targets = %#v, resolved=%v", targets, ok)
	}
}

func TestCoordinatedLeafReleasesAllTargetsWhenHandlerFails(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	cfg := config.Default()
	if err := config.Write(filepath.Join(rootDir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	a := &app{
		cwd:          rootDir,
		rawArgs:      []string{"--json", "pack", "--experimental", "--out", "artifact.xlsm"},
		stdout:       &bytes.Buffer{},
		stderr:       &bytes.Buffer{},
		coordination: manager,
	}
	root := a.rootCommand()
	root.SetArgs(a.rawArgs)
	if err := root.Execute(); err == nil {
		t.Fatal("pack unexpectedly succeeded without a template workbook")
	}

	descriptor, err := coordination.Lookup("pack")
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{workbookArgPath(rootDir, cfg.Excel.Path), filepath.Join(rootDir, "artifact.xlsm")} {
		identity, err := coordination.NewWorkbookIdentity(rootDir, path)
		if err != nil {
			t.Fatal(err)
		}
		lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
			Identity:      identity,
			Command:       descriptor.ID,
			OperationKind: descriptor.Policy.OperationKind,
			ResourceScope: descriptor.Policy.ResourceScope,
		})
		if err != nil {
			t.Fatalf("target remained locked after handler failure: %s: %v", path, err)
		}
		if err := lease.Release(); err != nil {
			t.Fatal(err)
		}
	}
}
