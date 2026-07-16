package excel

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

type recoveryLease struct {
	lease    *coordination.Lease
	metadata coordination.RecoveryMetadata
}

var anyExcelProcessRunningFunc = anyExcelProcessRunning

func (r Runner) recoveryManager() (*coordination.Manager, error) {
	if r.Coordination != nil {
		return r.Coordination, nil
	}
	return coordination.NewDefaultManager()
}

func (r Runner) recoveryProbePIDs() ([]int, error) {
	if runtime.GOOS != "windows" {
		return nil, nil
	}
	manager, err := r.recoveryManager()
	if err != nil {
		return nil, err
	}
	entries, err := manager.ListRecoveries()
	if err != nil {
		return nil, err
	}
	seen := map[int]struct{}{}
	for _, entry := range entries {
		if entry.State.Invalid || entry.State.Metadata == nil {
			return nil, fmt.Errorf("recovery metadata %s is invalid", entry.LockID)
		}
		if entry.State.Metadata != nil && entry.State.Metadata.ExcelPID > 0 {
			seen[entry.State.Metadata.ExcelPID] = struct{}{}
		}
	}
	result := make([]int, 0, len(seen))
	for pid := range seen {
		result = append(result, pid)
	}
	return result, nil
}

func (r Runner) acquireProcessRecoveryLeases(opts ProcessCleanupOptions) ([]recoveryLease, *coordinationFailure) {
	if runtime.GOOS != "windows" {
		return nil, nil
	}
	manager, err := r.recoveryManager()
	if err != nil {
		return nil, newCoordinationSetupFailure("process cleanup", "coordination_init_failed", err)
	}
	entries, err := manager.ListRecoveries()
	if err != nil {
		return nil, newCoordinationSetupFailure("process cleanup", coordination.RecoveryCheckFailedCode, err)
	}
	leases := make([]recoveryLease, 0, len(entries))
	release := func() {
		for i := len(leases) - 1; i >= 0; i-- {
			_ = leases[i].lease.Release()
		}
	}
	for _, entry := range entries {
		metadata := entry.State.Metadata
		if metadata == nil {
			continue
		}
		if opts.PID > 0 && metadata.ExcelPID != opts.PID {
			continue
		}
		identity, identityErr := coordination.NewWorkbookIdentity(filepath.Dir(metadata.Workbook), metadata.Workbook)
		if identityErr != nil || identity.LockID != entry.LockID {
			continue
		}
		lease, acquireErr := manager.Acquire(context.Background(), coordination.AcquireRequest{
			Identity:      identity,
			Command:       "process.cleanup",
			OperationKind: coordination.OperationMutate,
			ResourceScope: coordination.ResourceWorkbook,
		})
		if acquireErr != nil {
			release()
			var busy *coordination.BusyError
			if errors.As(acquireErr, &busy) {
				descriptor, _ := coordination.Lookup("process.cleanup")
				descriptor.Policy.ResourceScope = coordination.ResourceWorkbook
				return nil, newWorkbookBusyFailure("process cleanup", descriptor, busy)
			}
			return nil, newCoordinationSetupFailure("process cleanup", "coordination_acquire_failed", acquireErr)
		}
		state, stateErr := lease.RecoveryState()
		if stateErr != nil {
			_ = lease.Release()
			release()
			return nil, newCoordinationSetupFailure("process cleanup", coordination.RecoveryCheckFailedCode, stateErr)
		}
		if state.Metadata == nil || state.Metadata.Generation != metadata.Generation {
			_ = lease.Release()
			continue
		}
		leases = append(leases, recoveryLease{lease: lease, metadata: *state.Metadata})
	}
	return leases, nil
}

func releaseRecoveryLeases(leases []recoveryLease) {
	for i := len(leases) - 1; i >= 0; i-- {
		_ = leases[i].lease.Release()
	}
}

func clearRecoveredProcessMarkers(env *output.Envelope, opts ProcessCleanupOptions, leases []recoveryLease) {
	if env == nil || len(leases) == 0 {
		return
	}
	terminated := terminatedProcessIDs(env.Process)
	clearUnknown := false
	if opts.All {
		if any, err := anyExcelProcessRunningFunc(); err == nil {
			clearUnknown = !any
		}
	}
	cleared := make([]map[string]any, 0)
	clearFailures := make([]map[string]any, 0)
	for _, held := range leases {
		pid := held.metadata.ExcelPID
		_, pidTerminated := terminated[pid]
		if clearUnknown {
			pidTerminated = true
		} else if pid <= 0 {
			pidTerminated = clearUnknown
		}
		if !pidTerminated {
			continue
		}
		ok, err := held.lease.ClearRecovery(held.metadata.Generation)
		if err != nil {
			clearFailures = append(clearFailures, map[string]any{
				"code":     "workbook_recovery_clear_failed",
				"message":  fmt.Sprintf("Excel process cleanup succeeded, but recovery quarantine could not be cleared for %s.", held.metadata.Workbook),
				"workbook": held.metadata.Workbook,
				"cause":    err.Error(),
			})
			continue
		}
		if !ok {
			clearFailures = append(clearFailures, map[string]any{
				"code":     "workbook_recovery_generation_changed",
				"message":  fmt.Sprintf("Recovery quarantine changed while process cleanup was running and was not cleared for %s.", held.metadata.Workbook),
				"workbook": held.metadata.Workbook,
			})
			continue
		}
		cleared = append(cleared, map[string]any{
			"workbook": held.metadata.Workbook,
			"excel_pid": func() any {
				if pid > 0 {
					return pid
				}
				return nil
			}(),
		})
	}
	if len(cleared) > 0 {
		env.Recovery = map[string]any{
			"cleared": cleared,
			"count":   len(cleared),
		}
		env.Logs = append(env.Logs, fmt.Sprintf("cleared recovery quarantine for %d workbook(s)", len(cleared)))
	}
	if len(clearFailures) > 0 {
		env.Warnings = appendEnvelopeObjects(env.Warnings, clearFailures)
	}
}

func appendEnvelopeObjects(existing any, additions []map[string]any) any {
	switch values := existing.(type) {
	case nil:
		return additions
	case []map[string]any:
		return append(values, additions...)
	case []any:
		for _, addition := range additions {
			values = append(values, addition)
		}
		return values
	default:
		result := []any{existing}
		for _, addition := range additions {
			result = append(result, addition)
		}
		return result
	}
}

func terminatedProcessIDs(process any) map[int]struct{} {
	result := map[int]struct{}{}
	payload, ok := process.(map[string]any)
	if !ok {
		return result
	}
	items := make([]map[string]any, 0)
	switch raw := payload["results"].(type) {
	case []any:
		for _, item := range raw {
			if row, ok := item.(map[string]any); ok {
				items = append(items, row)
			}
		}
	case []map[string]any:
		items = append(items, raw...)
	}
	for _, row := range items {
		terminated, _ := row["terminated"].(bool)
		if !terminated {
			continue
		}
		switch pid := row["pid"].(type) {
		case float64:
			if pid > 0 {
				result[int(pid)] = struct{}{}
			}
		case int:
			if pid > 0 {
				result[pid] = struct{}{}
			}
		case json.Number:
			if parsed, err := strconv.Atoi(pid.String()); err == nil && parsed > 0 {
				result[parsed] = struct{}{}
			}
		}
	}
	return result
}

func anyExcelProcessRunning() (bool, error) {
	cmd := exec.Command("tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH")
	out, err := cmd.Output()
	if err != nil {
		return false, err
	}
	rows, err := tasklistRows(out)
	if err != nil {
		return false, err
	}
	for _, row := range rows {
		if len(row) < 2 || !strings.EqualFold(strings.TrimSpace(row[0]), "EXCEL.EXE") {
			continue
		}
		if pid, parseErr := strconv.Atoi(strings.TrimSpace(row[1])); parseErr == nil && pid > 0 {
			return true, nil
		}
	}
	return false, nil
}
