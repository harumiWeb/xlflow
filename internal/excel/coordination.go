package excel

import (
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

type coordinationFailure struct {
	env      output.Envelope
	exitCode int
}

func (r Runner) acquireBridgeCoordination(commandName string, args map[string]string, descriptor coordination.Descriptor) (*coordination.Lease, *coordinationFailure) {
	policy := descriptor.Policy
	if r.SkipCoordination || runtime.GOOS != "windows" || policy.ResourceScope != coordination.ResourceWorkbook || policy.ParallelSafe {
		return nil, nil
	}
	workbook := strings.TrimSpace(args["WorkbookPath"])
	if workbook == "" {
		// Preserve the bridge command's established argument validation contract.
		return nil, nil
	}
	baseDir := strings.TrimSpace(r.RootDir)
	if baseDir == "" {
		var err error
		baseDir, err = os.Getwd()
		if err != nil {
			return nil, newCoordinationSetupFailure(commandName, "coordination_identity_failed", err)
		}
	}
	if !filepath.IsAbs(baseDir) {
		absolute, err := filepath.Abs(baseDir)
		if err != nil {
			return nil, newCoordinationSetupFailure(commandName, "coordination_identity_failed", err)
		}
		baseDir = absolute
	}
	identity, err := coordination.NewWorkbookIdentity(baseDir, workbook)
	if err != nil {
		return nil, newCoordinationSetupFailure(commandName, "coordination_identity_failed", err)
	}
	manager := r.Coordination
	if manager == nil {
		manager, err = coordination.NewDefaultManager()
		if err != nil {
			return nil, newCoordinationSetupFailure(commandName, "coordination_init_failed", err)
		}
	}
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity:      identity,
		Command:       descriptor.ID,
		OperationKind: policy.OperationKind,
		ResourceScope: policy.ResourceScope,
	})
	if err == nil {
		return lease, nil
	}
	var busy *coordination.BusyError
	if errors.As(err, &busy) {
		return nil, newWorkbookBusyFailure(commandName, descriptor, busy)
	}
	return nil, newCoordinationSetupFailure(commandName, "coordination_acquire_failed", err)
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
