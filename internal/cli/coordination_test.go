package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"path/filepath"
	"runtime"
	"testing"

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
