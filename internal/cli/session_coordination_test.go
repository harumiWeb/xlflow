package cli

import (
	"context"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"runtime"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestSessionCoordinationStatusPayload(t *testing.T) {
	startedAt := time.Date(2026, time.July, 15, 9, 30, 0, 123456789, time.FixedZone("JST", 9*60*60))
	tests := []struct {
		name   string
		probe  coordination.ProbeResult
		fields map[string]any
	}{
		{
			name:   "idle ignores owner",
			probe:  coordination.ProbeResult{Owner: &coordination.OwnerMetadata{Command: "run"}},
			fields: map[string]any{"busy": false, "recovery_required": false},
		},
		{
			name:   "busy without owner",
			probe:  coordination.ProbeResult{Busy: true},
			fields: map[string]any{"busy": true, "recovery_required": false},
		},
		{
			name: "busy with public owner fields",
			probe: coordination.ProbeResult{Busy: true, Owner: &coordination.OwnerMetadata{
				SchemaVersion: 99,
				Generation:    "internal-generation",
				Workbook:      `C:\private\Book.xlsm`,
				PID:           12345,
				Command:       "run",
				OperationKind: coordination.OperationExecute,
				ResourceScope: coordination.ResourceWorkbook,
				StartedAt:     startedAt,
			}},
			fields: map[string]any{
				"busy":              true,
				"recovery_required": false,
				"resource_scope":    "workbook",
				"operation_kind":    "execute",
				"command":           "run",
				"pid":               float64(12345),
				"started_at":        "2026-07-15T00:30:00.123456789Z",
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			body, err := json.Marshal(sessionCoordinationStatusFromProbe(tt.probe))
			if err != nil {
				t.Fatal(err)
			}
			var got map[string]any
			if err := json.Unmarshal(body, &got); err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(got, tt.fields) {
				t.Fatalf("coordination payload = %#v, want %#v", got, tt.fields)
			}
			for _, internal := range []string{"schema_version", "generation", "workbook", "argv"} {
				if _, ok := got[internal]; ok {
					t.Fatalf("coordination payload exposed internal field %q: %#v", internal, got)
				}
			}
		})
	}
}

func TestRunSessionStatusProbesBeforeBridgeAndPreservesSession(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Excel.Path = filepath.Join("build", "Book.xlsm")
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

	wantSession := map[string]any{"active": true, "running": true, "metadata": map[string]any{"owner": "managed"}}
	runs := 0
	a := &app{cwd: rootDir, coordination: manager}
	env, code, err := a.runSessionStatus(context.Background(), cfg, func() (output.Envelope, int, error) {
		runs++
		if err := lease.Release(); err != nil {
			t.Fatal(err)
		}
		env := output.New("session")
		env.Session = wantSession
		return env, output.ExitSuccess, nil
	})
	if err != nil || code != output.ExitSuccess {
		t.Fatalf("runSessionStatus() = code %d, err %v", code, err)
	}
	if runs != 1 {
		t.Fatalf("bridge runs = %d, want 1", runs)
	}
	if !reflect.DeepEqual(env.Session, wantSession) {
		t.Fatalf("session payload changed: %#v", env.Session)
	}
	status := cliObjectMap(env.Coordination)
	if status["busy"] != true || status["command"] != "run" || status["operation_kind"] != "execute" || status["resource_scope"] != "workbook" {
		t.Fatalf("coordination payload = %#v", status)
	}
	if status["pid"] != float64(os.Getpid()) || status["started_at"] == "" {
		t.Fatalf("coordination owner = %#v", status)
	}
	probeAfter, err := manager.Probe(context.Background(), identity)
	if err != nil {
		t.Fatal(err)
	}
	if probeAfter.Busy {
		t.Fatal("test bridge callback did not release workbook lease")
	}
}

func TestRunSessionStatusKeepsFailureEnvelopeWhenProbeUnavailable(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	rootDir := t.TempDir()
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	wantError := &output.Error{Code: "session_failed", Message: "bridge status failed"}
	a := &app{cwd: rootDir, coordination: manager}
	env, code, err := a.runSessionStatus(ctx, cfg, func() (output.Envelope, int, error) {
		env := output.Failure("session", *wantError)
		env.Session = map[string]any{"active": false}
		env.Warnings = []map[string]any{{"code": "existing_warning", "message": "preserve me"}}
		return env, output.ExitEnvironment, nil
	})
	if err != nil || code != output.ExitEnvironment {
		t.Fatalf("runSessionStatus() = code %d, err %v", code, err)
	}
	if env.Status != output.StatusFailed || !reflect.DeepEqual(env.Error, wantError) {
		t.Fatalf("failure envelope changed: %#v", env)
	}
	if env.Coordination != nil {
		t.Fatalf("unavailable probe reported coordination: %#v", env.Coordination)
	}
	warnings := anySlice(env.Warnings)
	if len(warnings) != 2 || cliObjectMap(warnings[0])["code"] != "existing_warning" || cliObjectMap(warnings[1])["code"] != coordinationStatusUnavailableCode {
		t.Fatalf("warnings = %#v", warnings)
	}

	appendUniqueMessage(&env.Warnings, coordinationStatusUnavailableCode, "duplicate")
	if got := len(anySlice(env.Warnings)); got != 2 {
		t.Fatalf("coordination warning was duplicated: %#v", env.Warnings)
	}
}

func TestObserveSessionCoordinationTreatsMissingWorkbookAsIdle(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Excel.Path = filepath.Join("missing", "Book.xlsm")
	a := &app{cwd: t.TempDir(), coordination: manager}
	status, unavailable := a.observeSessionCoordination(context.Background(), cfg)
	if unavailable || status == nil || status.Busy {
		t.Fatalf("missing workbook observation = status %#v, unavailable %v", status, unavailable)
	}
}

func TestRunSessionStatusReturnsBridgeExecutionErrorAfterObservation(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	manager, err := coordination.NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	wantErr := errors.New("bridge execution failed")
	a := &app{cwd: t.TempDir(), coordination: manager}
	env, code, err := a.runSessionStatus(context.Background(), config.Default(), func() (output.Envelope, int, error) {
		return output.Envelope{}, output.ExitEnvironment, wantErr
	})
	if !errors.Is(err, wantErr) || code != output.ExitEnvironment {
		t.Fatalf("runSessionStatus() = code %d, err %v", code, err)
	}
	status := cliObjectMap(env.Coordination)
	if status["busy"] != false {
		t.Fatalf("pre-bridge observation was not retained: %#v", env.Coordination)
	}
}
