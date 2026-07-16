package excel

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestProcessListPassesRecoveryPIDsToBridge(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	manager, identity := setupExcelRecoveryMarker(t, root, 24680)
	_ = identity
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	var requests []excelbridge.Request
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     "dotnet",
			supports: true,
			requests: &requests,
			response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"process","logs":[],"process":[]}`)},
		}
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).ProcessList(ProcessListOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil || len(requests) != 1 {
		t.Fatalf("result code=%d error=%#v requests=%d", code, env.Error, len(requests))
	}
	var pids []int
	if err := json.Unmarshal([]byte(requests[0].Args["SkipWorkbookProbePids"]), &pids); err != nil {
		t.Fatal(err)
	}
	if len(pids) != 1 || pids[0] != 24680 {
		t.Fatalf("pids = %#v", pids)
	}
}

func TestProcessListFailsClosedOnInvalidRecoveryMetadata(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, filepath.Join(root, "Book.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	markerPath := filepath.Join(manager.StateDir(), identity.LockID+".recovery.json")
	if err := os.WriteFile(markerPath, []byte(`{`), 0o600); err != nil {
		t.Fatal(err)
	}
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	callCount := 0
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{name: "dotnet", supports: true, callCount: &callCount}
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).ProcessList(ProcessListOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment || env.Error == nil ||
		env.Error.Code != coordination.RecoveryCheckFailedCode ||
		env.Error.Phase != "coordination.recovery" {
		t.Fatalf("result code=%d error=%#v", code, env.Error)
	}
	if callCount != 0 {
		t.Fatalf("bridge calls = %d, want 0", callCount)
	}
}

func TestProcessCleanupClearsOnlyConfirmedTerminatedPID(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	manager, identity := setupExcelRecoveryMarker(t, root, 24680)
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     "dotnet",
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{
				"protocol_version":1,
				"status":"ok",
				"command":"process",
				"logs":["terminated 1 Excel process(es)"],
				"process":{"action":"cleanup","mode":"pid","total":1,"results":[{"pid":24680,"terminated":true,"method":"force"}]}
			}`)},
		}
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).ProcessCleanup(ProcessCleanupOptions{PID: 24680})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil {
		t.Fatalf("result code=%d env=%#v", code, env)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if state.Required {
		t.Fatalf("marker remains: %+v", state)
	}
	recovery, ok := env.Recovery.(map[string]any)
	if !ok || recovery["count"] != 1 {
		t.Fatalf("recovery = %#v", env.Recovery)
	}
}

func TestProcessCleanupFailureKeepsRecoveryMarker(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	manager, identity := setupExcelRecoveryMarker(t, root, 24680)
	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     "dotnet",
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{
				"protocol_version":1,
				"status":"failed",
				"command":"process",
				"error":{"code":"process_termination_failed","message":"failed"},
				"logs":[],
				"process":{"action":"cleanup","mode":"pid","total":1,"results":[{"pid":24680,"terminated":false,"method":"none"}]}
			}`)},
		}
	}
	_, _, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).ProcessCleanup(ProcessCleanupOptions{PID: 24680})
	if err != nil {
		t.Fatal(err)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required {
		t.Fatal("failed cleanup cleared recovery marker")
	}
}

func TestProcessCleanupAllClearsUnknownPIDMarkerAfterZeroProcessProof(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	manager, identity := setupExcelRecoveryMarker(t, root, 0)
	originalProvider := bridgeProviderForMode
	originalProcessCheck := anyExcelProcessRunningFunc
	t.Cleanup(func() {
		bridgeProviderForMode = originalProvider
		anyExcelProcessRunningFunc = originalProcessCheck
	})
	anyExcelProcessRunningFunc = func() (bool, error) { return false, nil }
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     "dotnet",
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{
				"protocol_version":1,
				"status":"ok",
				"command":"process",
				"logs":["0 Excel processes found"],
				"process":{"action":"cleanup","mode":"all","total":0,"results":[]}
			}`)},
		}
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).ProcessCleanup(ProcessCleanupOptions{All: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil {
		t.Fatalf("result code=%d env=%#v", code, env)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if state.Required {
		t.Fatalf("unknown-PID marker remains: %+v", state)
	}
}

func TestTasklistRowsTreatLocalizedNoMatchAsNonProcessRow(t *testing.T) {
	rows, err := tasklistRows([]byte("情報: 指定された条件に一致するタスクは実行されていません。\r\n"))
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) != 1 || len(rows[0]) != 1 {
		t.Fatalf("rows = %#v", rows)
	}
	rows, err = tasklistRows([]byte("\"EXCEL.EXE\",\"24680\",\"Console\",\"1\",\"10,000 K\"\r\n"))
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) != 1 || len(rows[0]) < 2 || rows[0][0] != "EXCEL.EXE" || rows[0][1] != "24680" {
		t.Fatalf("rows = %#v", rows)
	}
}

func TestManagedSessionDiscardClearsRecoveryMarker(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	cfg := config.Default()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, cfg.Excel.Path)
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
	if _, err := lease.PublishRecovery(coordination.RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
		Session:   coordination.RecoverySession{Active: true, Owner: "managed"},
		ExcelPID:  24680,
	}); err != nil {
		t.Fatal(err)
	}
	if err := lease.Release(); err != nil {
		t.Fatal(err)
	}

	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	var requests []excelbridge.Request
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     "dotnet",
			supports: true,
			requests: &requests,
			response: excelbridge.Response{Stdout: []byte(`{
				"protocol_version":1,
				"status":"ok",
				"command":"session",
				"logs":["discarded unsafe unsaved workbook changes"],
				"recovery":{"required":false,"reason":"managed_session_discarded","operation":"session.stop","excel_pid":24680,"cleanup_confirmed":true,"session":{"active":false,"owner":"managed"}}
			}`)},
		}
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).Session(cfg, "stop", SessionCommandOptions{Discard: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil {
		t.Fatalf("result code=%d error=%#v", code, env.Error)
	}
	if len(requests) != 1 || requests[0].Args["Discard"] != "true" {
		t.Fatalf("requests = %#v", requests)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if state.Required {
		t.Fatalf("marker remains: %+v", state)
	}
}

func TestExternalSessionDiscardRetainsValidRecoveryMarker(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	cfg := config.Default()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, cfg.Excel.Path)
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
	if _, err := lease.PublishRecovery(coordination.RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
		Session:   coordination.RecoverySession{Active: true, Owner: "external"},
		ExcelPID:  24680,
	}); err != nil {
		t.Fatal(err)
	}
	if err := lease.Release(); err != nil {
		t.Fatal(err)
	}

	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{
			name:     "dotnet",
			supports: true,
			response: excelbridge.Response{Stdout: []byte(`{
				"protocol_version":1,
				"status":"ok",
				"command":"session",
				"logs":["detached external Excel session"],
				"recovery":{"required":true,"reason":"external_session_detached","operation":"session.stop","excel_pid":24680,"cleanup_confirmed":false,"session":{"active":false,"owner":"external"}}
			}`)},
		}
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).Session(cfg, "stop", SessionCommandOptions{Discard: true})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil {
		t.Fatalf("result code=%d env=%#v", code, env)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required || state.Invalid || state.Metadata == nil ||
		state.Metadata.Session.Active ||
		state.Metadata.Session.Owner != "external" {
		t.Fatalf("state = %+v", state)
	}
	actions := coordination.RecoveryActions(state)
	if len(actions) == 0 || actions[0] != "close the workbook in Excel without saving" {
		t.Fatalf("actions = %#v", actions)
	}
	statusEnv, statusCode, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).Session(cfg, "status")
	if err != nil {
		t.Fatal(err)
	}
	if statusCode != output.ExitSuccess || statusEnv.Error != nil {
		t.Fatalf("status code=%d error=%#v", statusCode, statusEnv.Error)
	}
	session, ok := statusEnv.Session.(map[string]any)
	if !ok || session["active"] != false || session["discard_required"] != true || session["owner"] != "external" {
		t.Fatalf("session = %#v", statusEnv.Session)
	}
}

func TestSessionStatusUsesRecoveryMarkerWhenProjectSessionMetadataIsMissing(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	cfg := config.Default()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, cfg.Excel.Path)
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
	if _, err := lease.PublishRecovery(coordination.RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
		Session:   coordination.RecoverySession{Active: true, Owner: "external"},
		ExcelPID:  2147483000,
	}); err != nil {
		t.Fatal(err)
	}
	if err := lease.Release(); err != nil {
		t.Fatal(err)
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).Session(cfg, "status")
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil {
		t.Fatalf("result code=%d error=%#v", code, env.Error)
	}
	session, ok := env.Session.(map[string]any)
	if !ok || session["discard_required"] != true || session["source_of_truth"] != "uncertain" || session["owner"] != "external" {
		t.Fatalf("session = %#v", env.Session)
	}
}

func TestSessionStatusForNonSessionRecoveryDoesNotInventManagedOwner(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	manager, _ := setupExcelRecoveryMarker(t, root, 0)
	cfg := config.Default()
	cfg.Excel.Path = "Book.xlsm"
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).Session(cfg, "status")
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil {
		t.Fatalf("result code=%d error=%#v", code, env.Error)
	}
	session, ok := env.Session.(map[string]any)
	if !ok ||
		session["active"] != false ||
		session["discard_required"] != true ||
		session["source_of_truth"] != "uncertain" ||
		session["owner"] != "none" {
		t.Fatalf("session = %#v", env.Session)
	}
}

func TestSessionStatusUsesMetadataOnlyWhileRecoveryIsRequired(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows recovery coordination")
	}
	root := t.TempDir()
	cfg := config.Default()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, cfg.Excel.Path)
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
	if _, err := lease.PublishRecovery(coordination.RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
		Session:   coordination.RecoverySession{Active: true, Owner: "managed"},
		ExcelPID:  2147483000,
	}); err != nil {
		t.Fatal(err)
	}
	if err := lease.Release(); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(root, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	sessionBody, _ := json.Marshal(map[string]any{
		"pid":           2147483000,
		"workbook_path": identity.CanonicalPath,
		"owner":         "managed",
	})
	if err := os.WriteFile(filepath.Join(root, ".xlflow", "session.json"), sessionBody, 0o600); err != nil {
		t.Fatal(err)
	}

	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	callCount := 0
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{name: "dotnet", supports: true, callCount: &callCount}
	}
	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).Session(cfg, "status")
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil || callCount != 0 {
		t.Fatalf("result code=%d error=%#v calls=%d", code, env.Error, callCount)
	}
	session, ok := env.Session.(map[string]any)
	if !ok || session["source_of_truth"] != "uncertain" || session["recovery_required"] != true {
		t.Fatalf("session = %#v", env.Session)
	}
	coord, ok := env.Coordination.(map[string]any)
	if !ok || coord["recovery_required"] != true {
		t.Fatalf("coordination = %#v", env.Coordination)
	}
}

func setupExcelRecoveryMarker(t *testing.T, root string, excelPID int) (*coordination.Manager, coordination.WorkbookIdentity) {
	t.Helper()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, filepath.Join(root, "Book.xlsm"))
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
	return manager, identity
}
