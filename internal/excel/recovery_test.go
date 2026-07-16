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
		t.Fatalf("result code=%d error=%#v", code, env.Error)
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
