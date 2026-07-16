package excel

import (
	"context"
	"os"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestRunnerWorkbookBusyStopsUserFormCommandsBeforeBridgeExecution(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	spec := forms.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: forms.FormSpecForm{Name: "UserForm1"}}
	tests := []struct {
		name      string
		operation coordination.CommandID
		invoke    func(Runner, config.Config, string) (output.Envelope, int, error)
	}{
		{name: "list forms", operation: "list.forms", invoke: func(r Runner, cfg config.Config, _ string) (output.Envelope, int, error) {
			return r.ListForms(cfg, SessionCommandOptions{})
		}},
		{name: "inspect form", operation: "inspect.form", invoke: func(r Runner, cfg config.Config, _ string) (output.Envelope, int, error) {
			return r.InspectForm(cfg, InspectFormOptions{Name: "UserForm1", Basis: "designer"})
		}},
		{name: "form build", operation: "form.build", invoke: func(r Runner, cfg config.Config, _ string) (output.Envelope, int, error) {
			return r.FormWrite(cfg, FormWriteOptions{Action: "build", SpecPath: "UserForm1.yaml", Spec: spec})
		}},
		{name: "form apply", operation: "form.apply", invoke: func(r Runner, cfg config.Config, _ string) (output.Envelope, int, error) {
			return r.FormWrite(cfg, FormWriteOptions{Action: "apply", SpecPath: "UserForm1.yaml", Spec: spec})
		}},
		{name: "form export image", operation: "form.export-image", invoke: func(r Runner, cfg config.Config, root string) (output.Envelope, int, error) {
			return r.FormExportImage(cfg, FormExportImageOptions{Name: "UserForm1", OutPath: filepath.Join(root, "UserForm1.png")})
		}},
		{name: "pull", operation: "pull", invoke: func(r Runner, cfg config.Config, _ string) (output.Envelope, int, error) {
			return r.PullWithOptions(cfg, SessionCommandOptions{})
		}},
		{name: "push", operation: "push", invoke: func(r Runner, cfg config.Config, _ string) (output.Envelope, int, error) {
			return r.PushWithOptions(cfg, PushOptions{})
		}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
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
			lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{Identity: identity, Command: "run", OperationKind: coordination.OperationExecute, ResourceScope: coordination.ResourceWorkbook})
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

			env, code, err := tt.invoke(Runner{RootDir: root, BridgeMode: "dotnet", Coordination: manager}, cfg, root)
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
			if !ok || details["operation"] != tt.operation || details["retryable"] != true {
				t.Fatalf("details = %#v", env.Error.Details)
			}
		})
	}
}

func TestRunnerUserFormCoordinationDoesNotConflictAcrossWorkbooks(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	rootA := t.TempDir()
	rootB := t.TempDir()
	identityA, err := coordination.NewWorkbookIdentity(rootA, cfg.Excel.Path)
	if err != nil {
		t.Fatal(err)
	}
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{Identity: identityA, Command: "form.build", OperationKind: coordination.OperationDesigner, ResourceScope: coordination.ResourceWorkbook})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = lease.Release() }()

	original := bridgeProviderForMode
	t.Cleanup(func() { bridgeProviderForMode = original })
	callCount := 0
	bridgeProviderForMode = func(string, excelbridge.Mode) excelbridge.Provider {
		return trackingBridgeProvider{name: "dotnet", supports: true, callCount: &callCount, response: excelbridge.Response{Stdout: []byte(`{"protocol_version":1,"status":"ok","command":"inspect-form","logs":[]}`)}}
	}
	env, code, err := (Runner{RootDir: rootB, BridgeMode: "dotnet", Coordination: manager}).InspectForm(cfg, InspectFormOptions{Name: "UserForm1", Basis: "designer"})
	if err != nil {
		t.Fatal(err)
	}
	if code != output.ExitSuccess || env.Error != nil || callCount != 1 {
		t.Fatalf("different workbook result = code %d, error %#v, bridge calls %d", code, env.Error, callCount)
	}
}

func TestRunnerWorkbookRecoveryStopsBeforeBridgeExecution(t *testing.T) {
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
	if _, err := lease.PublishRecovery(coordination.RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
	}); err != nil {
		t.Fatal(err)
	}
	if err := lease.Release(); err != nil {
		t.Fatal(err)
	}

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
	if code != output.ExitEnvironment || env.Error == nil || env.Error.Code != coordination.WorkbookRecoveryRequiredCode {
		t.Fatalf("result = code %d, error %#v", code, env.Error)
	}
	if callCount != 0 {
		t.Fatalf("bridge calls = %d, want 0", callCount)
	}
}

func TestRunnerRecoveryReadFailureIsFailClosed(t *testing.T) {
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
	if err := os.Mkdir(filepath.Join(manager.StateDir(), identity.LockID+".recovery.json"), 0o700); err != nil {
		t.Fatal(err)
	}
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
	if code != output.ExitEnvironment || env.Error == nil ||
		env.Error.Code != coordination.RecoveryCheckFailedCode ||
		env.Error.Phase != "coordination.recovery" {
		t.Fatalf("result = code %d, error %#v", code, env.Error)
	}
	details, ok := env.Error.Details.(map[string]any)
	if !ok ||
		details["workbook"] != identity.CanonicalPath ||
		details["attempted_operation"] != coordination.CommandID("push") ||
		details["retryable"] != false ||
		details["cause"] == nil {
		t.Fatalf("details = %#v", env.Error.Details)
	}
	if callCount != 0 {
		t.Fatalf("bridge calls = %d, want 0", callCount)
	}
}
