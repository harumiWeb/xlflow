package excel

import (
	"context"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestRunnerWorkbookBusyStopsBeforeBridgeExecution(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	root := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
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
	defer func() { _ = lease.Release() }()

	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	callCount := 0
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{name: "dotnet", supports: true, callCount: &callCount}
	}

	env, code, err := (Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}).PushWithOptions(cfg, PushOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitEnvironment || env.Error == nil || env.Error.Code != coordination.WorkbookBusyCode {
		t.Fatalf("result = code %d, error %#v", code, env.Error)
	}
	if callCount != 0 {
		t.Fatalf("bridge calls = %d, want 0", callCount)
	}
	details, ok := env.Error.Details.(map[string]any)
	if !ok || details["operation"] != coordination.CommandID("push") || details["retryable"] != true {
		t.Fatalf("details = %#v", env.Error.Details)
	}
}
