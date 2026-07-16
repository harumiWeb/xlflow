package excel

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

type coordinationFailure struct {
	env      output.Envelope
	exitCode int
}

func (r Runner) observeRecoveryBeforeBridge(commandName string, args map[string]string, descriptor coordination.Descriptor) (*coordinationFailure, bool) {
	if runtime.GOOS != "windows" || descriptor.Policy.RecoveryBehavior != coordination.RecoveryObserve {
		return nil, false
	}
	workbook := strings.TrimSpace(args["WorkbookPath"])
	if workbook == "" {
		return nil, false
	}
	baseDir := strings.TrimSpace(r.RootDir)
	if baseDir == "" {
		var err error
		baseDir, err = os.Getwd()
		if err != nil {
			return newCoordinationSetupFailure(commandName, "coordination_identity_failed", err), true
		}
	}
	if !filepath.IsAbs(baseDir) {
		absolute, err := filepath.Abs(baseDir)
		if err != nil {
			return newCoordinationSetupFailure(commandName, "coordination_identity_failed", err), true
		}
		baseDir = absolute
	}
	identity, err := coordination.NewWorkbookIdentity(baseDir, workbook)
	if err != nil {
		return newCoordinationSetupFailure(commandName, "coordination_identity_failed", err), true
	}
	manager := r.Coordination
	if manager == nil {
		manager, err = coordination.NewDefaultManager()
		if err != nil {
			return newCoordinationSetupFailure(commandName, "coordination_init_failed", err), true
		}
	}
	observation, err := manager.Observe(context.Background(), identity)
	if err != nil {
		return newCoordinationSetupFailure(commandName, coordination.RecoveryCheckFailedCode, err), true
	}
	if !observation.Recovery.Required {
		return nil, false
	}
	if descriptor.ID != "session.status" {
		return nil, false
	}
	env := recoverySessionStatusEnvelope(commandName, args, identity, observation)
	return &coordinationFailure{env: env, exitCode: output.ExitSuccess}, true
}

func recoverySessionStatusEnvelope(commandName string, args map[string]string, identity coordination.WorkbookIdentity, observation coordination.Observation) output.Envelope {
	type sessionMetadata struct {
		PID          int    `json:"pid"`
		WorkbookPath string `json:"workbook_path"`
		Owner        string `json:"owner"`
		Poisoned     bool   `json:"poisoned"`
		PoisonReason string `json:"poison_reason"`
	}
	var metadata sessionMetadata
	metadataPath := strings.TrimSpace(args["MetadataPath"])
	if body, err := os.ReadFile(metadataPath); err == nil {
		_ = json.Unmarshal(body, &metadata)
	}
	matches := metadata.WorkbookPath != "" && coordination.SamePath(metadata.WorkbookPath, identity.CanonicalPath)
	running := false
	if matches && metadata.PID > 0 {
		running, _ = processExistsByPID(metadata.PID)
	}
	owner := strings.TrimSpace(metadata.Owner)
	if owner == "" {
		owner = "managed"
	}
	session := map[string]any{
		"active":               matches && running,
		"running":              running,
		"workbook_open":        nil,
		"workbook_path":        identity.CanonicalPath,
		"owner":                owner,
		"dirty":                nil,
		"save_required":        false,
		"live_newer_than_disk": false,
		"source_of_truth":      "uncertain",
		"discard_required":     matches,
		"recovery_required":    true,
	}
	if matches {
		session["metadata"] = map[string]any{
			"pid":           metadata.PID,
			"workbook_path": metadata.WorkbookPath,
			"owner":         owner,
			"poisoned":      metadata.Poisoned,
			"poison_reason": metadata.PoisonReason,
		}
	}
	env := output.New(commandName)
	env.Session = session
	env.Target = map[string]any{"kind": map[bool]string{true: "live_session", false: "file"}[matches && running], "path": identity.CanonicalPath}
	env.Workbook = map[string]any{
		"path":       identity.CanonicalPath,
		"session":    matches && running,
		"saved":      false,
		"dirty":      nil,
		"needs_save": false,
	}
	env.Coordination = recoveryCoordinationPayload(observation)
	env.Warnings = []map[string]any{{
		"code":    coordination.WorkbookRecoveryRequiredCode,
		"message": "The workbook is quarantined. Session status used persisted metadata and did not access the uncertain Excel instance.",
	}}
	env.Hints = []map[string]any{{
		"code":    "recover_workbook",
		"message": strings.Join(coordination.RecoveryActions(observation.Recovery), "; "),
	}}
	env.Logs = []string{"reported recovery-aware session status without Excel COM access"}
	return env
}

func recoveryCoordinationPayload(observation coordination.Observation) map[string]any {
	payload := map[string]any{
		"busy":              observation.Busy,
		"recovery_required": observation.Recovery.Required,
	}
	if observation.Recovery.Required {
		recovery := coordination.RecoveryDetails(coordination.WorkbookIdentity{}, observation.Recovery)
		for _, key := range []string{"workbook", "retryable", "wait_will_resolve", "recovery_actions"} {
			delete(recovery, key)
		}
		payload["recovery"] = recovery
	}
	if observation.Busy && observation.Owner != nil {
		payload["resource_scope"] = observation.Owner.ResourceScope
		payload["operation_kind"] = observation.Owner.OperationKind
		payload["command"] = observation.Owner.Command
		payload["pid"] = observation.Owner.PID
		payload["started_at"] = observation.Owner.StartedAt.UTC().Format(time.RFC3339Nano)
	}
	return payload
}

func processExistsByPID(pid int) (bool, error) {
	if pid <= 0 {
		return false, nil
	}
	cmd := exec.Command("tasklist", "/FI", fmt.Sprintf("PID eq %d", pid), "/FO", "CSV", "/NH")
	out, err := cmd.Output()
	if err != nil {
		return false, err
	}
	text := strings.TrimSpace(string(out))
	return text != "" && !strings.Contains(strings.ToLower(text), "no tasks are running"), nil
}

func (r Runner) acquireBridgeCoordination(commandName string, args map[string]string, descriptor coordination.Descriptor) (*coordination.Lease, bool, *coordinationFailure) {
	policy := descriptor.Policy
	if runtime.GOOS != "windows" || !runnerRequiresWorkbookLease(policy) {
		return nil, false, nil
	}
	workbook := strings.TrimSpace(args["WorkbookPath"])
	if workbook == "" {
		// Preserve the bridge command's established argument validation contract.
		return nil, false, nil
	}
	baseDir := strings.TrimSpace(r.RootDir)
	if baseDir == "" {
		var err error
		baseDir, err = os.Getwd()
		if err != nil {
			return nil, false, newCoordinationSetupFailure(commandName, "coordination_identity_failed", err)
		}
	}
	if !filepath.IsAbs(baseDir) {
		absolute, err := filepath.Abs(baseDir)
		if err != nil {
			return nil, false, newCoordinationSetupFailure(commandName, "coordination_identity_failed", err)
		}
		baseDir = absolute
	}
	identity, err := coordination.NewWorkbookIdentity(baseDir, workbook)
	if err != nil {
		return nil, false, newCoordinationSetupFailure(commandName, "coordination_identity_failed", err)
	}
	if borrowed := r.BorrowedLeases.Lease(identity); borrowed != nil {
		if failure := requireBridgeRecoveryAllowed(commandName, descriptor, borrowed, bridgeRecoveryIntent(descriptor.ID, args)); failure != nil {
			return nil, false, failure
		}
		return borrowed, false, nil
	}
	manager := r.Coordination
	if manager == nil {
		manager, err = coordination.NewDefaultManager()
		if err != nil {
			return nil, false, newCoordinationSetupFailure(commandName, "coordination_init_failed", err)
		}
	}
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity:      identity,
		Command:       descriptor.ID,
		OperationKind: policy.OperationKind,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err == nil {
		if failure := requireBridgeRecoveryAllowed(commandName, descriptor, lease, bridgeRecoveryIntent(descriptor.ID, args)); failure != nil {
			_ = lease.Release()
			return nil, false, failure
		}
		return lease, true, nil
	}
	var busy *coordination.BusyError
	if errors.As(err, &busy) {
		return nil, false, newWorkbookBusyFailure(commandName, descriptor, busy)
	}
	return nil, false, newCoordinationSetupFailure(commandName, "coordination_acquire_failed", err)
}

func runnerRequiresWorkbookLease(policy coordination.Policy) bool {
	if policy.ResourceScope == coordination.ResourceWorkbook && !policy.ParallelSafe {
		return true
	}
	return policy.RecoveryBehavior == coordination.RecoveryBlock
}

func bridgeRecoveryIntent(commandID coordination.CommandID, args map[string]string) bool {
	if commandID == "session.stop" {
		value, _ := strconv.ParseBool(strings.TrimSpace(args["Discard"]))
		return value
	}
	return commandID == "recovery.clear"
}

func requireBridgeRecoveryAllowed(commandName string, descriptor coordination.Descriptor, lease *coordination.Lease, recoveryIntent bool) *coordinationFailure {
	err := lease.RequireRecoveryAllowed(descriptor.Policy.RecoveryBehavior, recoveryIntent)
	if err == nil {
		return nil
	}
	var required *coordination.RecoveryRequiredError
	if errors.As(err, &required) {
		return newWorkbookRecoveryFailure(commandName, descriptor, required)
	}
	return newCoordinationSetupFailure(commandName, coordination.RecoveryCheckFailedCode, err)
}

func newCoordinationSetupFailure(commandName, code string, err error) *coordinationFailure {
	return &coordinationFailure{
		env: output.Failure(commandName, output.Error{
			Code:    code,
			Message: err.Error(),
			Source:  "xlflow",
			Phase:   "coordination.acquire",
		}),
		exitCode: output.ExitEnvironment,
	}
}

func newWorkbookBusyFailure(commandName string, descriptor coordination.Descriptor, busy *coordination.BusyError) *coordinationFailure {
	details := map[string]any{
		"workbook":       busy.Identity.CanonicalPath,
		"operation":      descriptor.ID,
		"resource_scope": descriptor.Policy.ResourceScope,
		"retryable":      descriptor.Policy.RetryableWhenBusy,
	}
	if busy.Owner != nil {
		details["owner"] = busy.Owner
	}
	message := fmt.Sprintf("Another xlflow operation is currently using this workbook: %s. Retry after it completes.", busy.Identity.CanonicalPath)
	return &coordinationFailure{
		env: output.Failure(commandName, output.Error{
			Code:    coordination.WorkbookBusyCode,
			Message: message,
			Source:  "xlflow",
			Phase:   "coordination.acquire",
			Details: details,
		}),
		exitCode: output.ExitEnvironment,
	}
}

func newWorkbookRecoveryFailure(commandName string, descriptor coordination.Descriptor, required *coordination.RecoveryRequiredError) *coordinationFailure {
	details := coordination.RecoveryDetails(required.Identity, required.State)
	details["attempted_operation"] = descriptor.ID
	return &coordinationFailure{
		env: output.Failure(commandName, output.Error{
			Code:    coordination.WorkbookRecoveryRequiredCode,
			Message: "The workbook is in an uncertain Excel state after a previous operation. Explicit recovery is required before this command can run; --wait will not resolve it.",
			Source:  "xlflow",
			Phase:   "coordination.recovery",
			Details: details,
		}),
		exitCode: output.ExitEnvironment,
	}
}
