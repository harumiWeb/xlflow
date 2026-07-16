package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestRecoveryClearVerifiesStoppedExcelPIDAndRemovesMarker(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
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
	publishRecoveryForTest(t, manager, identity, 2147483000)

	var stdout bytes.Buffer
	a := &app{
		cwd:          rootDir,
		rawArgs:      []string{"--json", "recovery", "clear"},
		stdout:       &stdout,
		stderr:       &bytes.Buffer{},
		coordination: manager,
	}
	root := a.rootCommand()
	root.SetArgs(a.rawArgs)
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatal(err)
	}
	recovery := cliObjectMap(env.Recovery)
	if recovery["cleared"] != true || recovery["forced"] != false {
		t.Fatalf("recovery = %#v", recovery)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if state.Required {
		t.Fatalf("recovery marker remains: %+v", state)
	}
}

func TestRecoveryClearRequiresForceWhenExcelPIDIsUnknown(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
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
	identity, _ := coordination.NewWorkbookIdentity(rootDir, cfg.Excel.Path)
	publishRecoveryForTest(t, manager, identity, 0)

	var stdout bytes.Buffer
	a := &app{
		cwd:          rootDir,
		rawArgs:      []string{"--json", "recovery", "clear"},
		stdout:       &stdout,
		stderr:       &bytes.Buffer{},
		coordination: manager,
	}
	root := a.rootCommand()
	root.SetArgs(a.rawArgs)
	err = root.Execute()
	if output.ExitCode(err) != output.ExitEnvironment {
		t.Fatalf("exit=%d err=%v", output.ExitCode(err), err)
	}
	var env output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatal(err)
	}
	if env.Error == nil || env.Error.Code != coordination.WorkbookRecoveryVerificationFailedCode {
		t.Fatalf("error = %#v", env.Error)
	}

	stdout.Reset()
	a.rawArgs = []string{"--json", "recovery", "clear", "--force"}
	root = a.rootCommand()
	root.SetArgs(a.rawArgs)
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	if err := json.Unmarshal(stdout.Bytes(), &env); err != nil {
		t.Fatal(err)
	}
	recovery := cliObjectMap(env.Recovery)
	if recovery["cleared"] != true || recovery["forced"] != true {
		t.Fatalf("recovery = %#v", recovery)
	}
	var warnings []map[string]any
	body, err := json.Marshal(env.Warnings)
	if err != nil {
		t.Fatal(err)
	}
	if err := json.Unmarshal(body, &warnings); err != nil {
		t.Fatal(err)
	}
	if len(warnings) != 1 || warnings[0]["code"] != "recovery_force_cleared" {
		t.Fatalf("warnings = %#v", warnings)
	}
}

func TestTasklistPIDRunningIgnoresLocalizedNoMatchMessage(t *testing.T) {
	running, err := tasklistPIDRunning([]byte("情報: 指定された条件に一致するタスクは実行されていません。\r\n"), 24680)
	if err != nil {
		t.Fatal(err)
	}
	if running {
		t.Fatal("localized tasklist no-match message was treated as a running process")
	}
	running, err = tasklistPIDRunning([]byte("\"EXCEL.EXE\",\"24680\",\"Console\",\"1\",\"10,000 K\"\r\n"), 24680)
	if err != nil {
		t.Fatal(err)
	}
	if !running {
		t.Fatal("tasklist CSV row was not recognized")
	}
}

func publishRecoveryForTest(t *testing.T, manager *coordination.Manager, identity coordination.WorkbookIdentity, excelPID int) {
	t.Helper()
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity:      identity,
		Command:       "run",
		OperationKind: coordination.OperationExecute,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := lease.PublishRecovery(coordination.RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
		ExcelPID:  excelPID,
	}); err != nil {
		_ = lease.Release()
		t.Fatal(err)
	}
	if err := lease.Release(); err != nil {
		t.Fatal(err)
	}
}
