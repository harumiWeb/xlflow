package cli

import (
	"context"
	"errors"
	"fmt"
	"os"
	"os/signal"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/workbookformat"
)

const (
	coordinationWaitArgsInvalidCode   = "coordination_wait_args_invalid"
	coordinationWaitUnsupportedCode   = "coordination_wait_unsupported"
	coordinationStatusUnavailableCode = "coordination_status_unavailable"
)

type sessionCoordinationStatus struct {
	Busy          bool   `json:"busy"`
	ResourceScope string `json:"resource_scope,omitempty"`
	OperationKind string `json:"operation_kind,omitempty"`
	Command       string `json:"command,omitempty"`
	PID           int    `json:"pid,omitempty"`
	StartedAt     string `json:"started_at,omitempty"`
}

func (a *app) validateCoordinationWaitOptions(cmd *cobra.Command) error {
	timeoutChanged := false
	if cmd != nil {
		if flag := cmd.Flags().Lookup("wait-timeout"); flag != nil {
			timeoutChanged = flag.Changed
		}
	}
	commandName := coordinationCommandName(cmd)
	if a.waitTimeout <= 0 {
		return a.writeFailure(commandName, output.ExitConfig, coordinationWaitArgsInvalidCode, fmt.Errorf("--wait-timeout must be greater than zero"))
	}
	if timeoutChanged && !a.wait {
		return a.writeFailure(commandName, output.ExitConfig, coordinationWaitArgsInvalidCode, fmt.Errorf("--wait-timeout requires --wait"))
	}
	if !a.wait || isGeneratedCobraCommand(cmd) {
		return nil
	}
	descriptor, err := coordination.LookupCLI(cmd.CommandPath())
	if err != nil {
		return a.writeFailure(commandName, output.ExitEnvironment, coordination.MissingPolicyCode, err)
	}
	policy := descriptor.Policy
	if policy.ResourceScope != coordination.ResourceWorkbook || policy.ParallelSafe || !policy.RetryableWhenBusy {
		return a.writeFailure(commandName, output.ExitConfig, coordinationWaitUnsupportedCode, fmt.Errorf("--wait is not supported for %s", commandName))
	}
	return nil
}

func coordinationCommandName(cmd *cobra.Command) string {
	if cmd == nil {
		return "xlflow"
	}
	return strings.TrimSpace(strings.TrimPrefix(cmd.CommandPath(), "xlflow"))
}

func (a *app) wrapCoordinatedLeaves(root *cobra.Command) {
	var walk func(*cobra.Command)
	walk = func(command *cobra.Command) {
		for _, child := range command.Commands() {
			walk(child)
		}
		if command.RunE == nil {
			return
		}
		descriptor, err := coordination.LookupCLI(command.CommandPath())
		if err != nil || descriptor.Policy.ResourceScope != coordination.ResourceWorkbook || descriptor.Policy.ParallelSafe {
			return
		}
		original := command.RunE
		command.RunE = func(cmd *cobra.Command, args []string) error {
			targets, resolved := a.coordinationTargets(cmd, args, descriptor.ID)
			if !resolved {
				return original(cmd, args)
			}
			return a.withWorkbookCoordination(cmd.Context(), descriptor.ID, targets, func() error {
				return original(cmd, args)
			})
		}
	}
	walk(root)
}

func (a *app) coordinationTargets(cmd *cobra.Command, args []string, commandID coordination.CommandID) ([]string, bool) {
	switch commandID {
	case "new":
		name := ""
		if len(args) > 0 {
			name = args[0]
		}
		normalized, err := workbookformat.NormalizeProjectWorkbookName(name)
		if err != nil {
			return nil, false
		}
		return []string{filepath.Join(a.cwd, "build", normalized)}, true
	case "init":
		if len(args) != 1 || strings.TrimSpace(args[0]) == "" {
			return nil, false
		}
		source := workbookArgPath(a.cwd, args[0])
		return []string{source, filepath.Join(a.cwd, "build", filepath.Base(source))}, true
	case "formulas.pull":
		if value, ok := commandFlagString(cmd, "src"); ok && value != "" {
			return []string{workbookArgPath(a.cwd, value)}, true
		}
	case "diff":
		if len(args) == 2 {
			return []string{workbookArgPath(a.cwd, args[0]), workbookArgPath(a.cwd, args[1])}, true
		}
		return nil, false
	case "pack":
		out, outOK := commandFlagString(cmd, "out")
		experimental, experimentalErr := cmd.Flags().GetBool("experimental")
		if !outOK || out == "" || experimentalErr != nil || !experimental {
			return nil, false
		}
		cfg, ok := a.coordinationConfig()
		if !ok {
			return nil, false
		}
		configured := workbookArgPath(a.cwd, cfg.Excel.Path)
		template, _ := commandFlagString(cmd, "template")
		if template == "" {
			template = configured
		} else {
			template = workbookArgPath(a.cwd, template)
		}
		return []string{configured, template, workbookArgPath(a.cwd, out)}, true
	case "run":
		cfg, ok := a.coordinationConfig()
		if !ok {
			return nil, false
		}
		workbook := workbookArgPath(a.cwd, cfg.Excel.Path)
		if input, exists := commandFlagString(cmd, "input"); exists && input != "" {
			workbook = workbookArgPath(a.cwd, input)
		}
		targets := []string{workbook}
		if saveAs, exists := commandFlagString(cmd, "save-as"); exists && saveAs != "" {
			targets = append(targets, workbookArgPath(a.cwd, saveAs))
		}
		return targets, true
	case "export-image", "edit.sheet.add", "edit.cell", "edit.formula", "edit.range", "edit.rows", "edit.columns":
		if len(args) == 1 && strings.TrimSpace(args[0]) != "" {
			return []string{workbookArgPath(a.cwd, args[0])}, true
		}
	}

	cfg, ok := a.coordinationConfig()
	if !ok || strings.TrimSpace(cfg.Excel.Path) == "" {
		return nil, false
	}
	return []string{workbookArgPath(a.cwd, cfg.Excel.Path)}, true
}

func (a *app) coordinationConfig() (config.Config, bool) {
	cfg, err := config.Load(a.cwd)
	if err != nil && errors.Is(err, config.ErrInvalidExcelBridge) && a.hasValidBridgeOverride() {
		cfg, err = config.LoadAllowInvalidExcelBridge(a.cwd)
	}
	return cfg, err == nil
}

func commandFlagString(cmd *cobra.Command, name string) (string, bool) {
	if cmd == nil || cmd.Flags().Lookup(name) == nil {
		return "", false
	}
	value, err := cmd.Flags().GetString(name)
	return strings.TrimSpace(value), err == nil
}

func (a *app) withWorkbookCoordination(ctx context.Context, commandID coordination.CommandID, workbookPaths []string, run func() error) error {
	release, err := a.acquireWorkbookCoordination(ctx, commandID, workbookPaths)
	if err != nil {
		return err
	}
	defer release()
	return run()
}

func (a *app) acquireWorkbookCoordination(ctx context.Context, commandID coordination.CommandID, workbookPaths []string) (func(), error) {
	descriptor, err := coordination.Lookup(commandID)
	if err != nil {
		return nil, a.writeFailure(string(commandID), output.ExitEnvironment, coordination.MissingPolicyCode, err)
	}
	policy := descriptor.Policy
	if runtime.GOOS != "windows" || policy.ResourceScope != coordination.ResourceWorkbook || policy.ParallelSafe {
		return func() {}, nil
	}
	if ctx == nil {
		ctx = context.Background()
	}
	identities := make(map[string]coordination.WorkbookIdentity, len(workbookPaths))
	for _, workbookPath := range workbookPaths {
		if strings.TrimSpace(workbookPath) == "" {
			continue
		}
		identity, identityErr := coordination.NewWorkbookIdentity(a.cwd, workbookPath)
		if identityErr != nil {
			return nil, a.writeFailure(string(commandID), output.ExitEnvironment, "coordination_identity_failed", identityErr)
		}
		identities[identity.LockID] = identity
	}
	if len(identities) == 0 {
		return nil, a.writeFailure(string(commandID), output.ExitEnvironment, "coordination_identity_failed", errors.New("workbook path is required for coordination"))
	}

	lockIDs := make([]string, 0, len(identities))
	for lockID := range identities {
		lockIDs = append(lockIDs, lockID)
	}
	sort.Strings(lockIDs)
	manager, err := a.coordinationManager()
	if err != nil {
		return nil, a.writeFailure(string(commandID), output.ExitEnvironment, "coordination_init_failed", err)
	}

	acquireCtx := ctx
	stopSignal := func() {}
	cancelTimeout := func() {}
	if a.wait {
		acquireCtx, stopSignal = signal.NotifyContext(ctx, os.Interrupt)
		acquireCtx, cancelTimeout = context.WithTimeout(acquireCtx, a.waitTimeout)
	}
	defer stopSignal()
	defer cancelTimeout()

	leases := make([]*coordination.Lease, 0, len(lockIDs))
	waitingAnnounced := false
	release := func() {
		for i := len(leases) - 1; i >= 0; i-- {
			_ = leases[i].Release()
		}
	}
	for _, lockID := range lockIDs {
		identity := identities[lockID]
		request := coordination.AcquireRequest{
			Identity:      identity,
			Command:       descriptor.ID,
			OperationKind: policy.OperationKind,
			ResourceScope: policy.ResourceScope,
		}
		lease, acquireErr := manager.Acquire(acquireCtx, request)
		var busy *coordination.BusyError
		if a.wait && errors.As(acquireErr, &busy) {
			if !waitingAnnounced && !a.json {
				_, _ = fmt.Fprintf(a.stderrWriter(), "Waiting up to %s for the workbook to become available: %s\n", a.waitTimeout, busy.Identity.CanonicalPath)
				waitingAnnounced = true
			}
			request.Wait = true
			lease, acquireErr = manager.Acquire(acquireCtx, request)
		}
		if acquireErr == nil && a.wait && acquireCtx.Err() != nil {
			_ = lease.Release()
			acquireErr = acquireCtx.Err()
		}
		if acquireErr != nil {
			release()
			if a.wait && errors.Is(acquireErr, context.DeadlineExceeded) {
				return nil, a.writeWorkbookWaitFailure(descriptor, identity, coordination.WorkbookBusyTimeoutCode, fmt.Sprintf("The workbook did not become available within %s.", a.waitTimeout), acquireErr)
			}
			if a.wait && errors.Is(acquireErr, context.Canceled) {
				return nil, a.writeWorkbookWaitFailure(descriptor, identity, coordination.WorkbookBusyCancelledCode, "Waiting for the workbook was cancelled.", acquireErr)
			}
			if errors.As(acquireErr, &busy) {
				return nil, a.writeWorkbookBusyFailure(descriptor, busy)
			}
			return nil, a.writeFailure(string(commandID), output.ExitEnvironment, "coordination_acquire_failed", acquireErr)
		}
		leases = append(leases, lease)
	}
	return release, nil
}

func (a *app) coordinationManager() (*coordination.Manager, error) {
	if a.coordination != nil {
		return a.coordination, nil
	}
	return coordination.NewDefaultManager()
}

func (a *app) runSessionStatus(ctx context.Context, cfg config.Config, run func() (output.Envelope, int, error)) (output.Envelope, int, error) {
	status, unavailable := a.observeSessionCoordination(ctx, cfg)
	env, code, err := run()
	if status != nil {
		env.Coordination = status
	}
	if unavailable {
		appendUniqueMessage(&env.Warnings, coordinationStatusUnavailableCode, "Workbook coordination status could not be observed; session status is still available.")
	}
	return env, code, err
}

func (a *app) observeSessionCoordination(ctx context.Context, cfg config.Config) (*sessionCoordinationStatus, bool) {
	if runtime.GOOS != "windows" {
		return nil, false
	}
	identity, err := coordination.NewWorkbookIdentity(a.cwd, cfg.Excel.Path)
	if err != nil {
		return nil, true
	}
	manager, err := a.coordinationManager()
	if err != nil {
		return nil, true
	}
	result, err := manager.Probe(ctx, identity)
	if err != nil {
		return nil, true
	}
	return sessionCoordinationStatusFromProbe(result), false
}

func sessionCoordinationStatusFromProbe(result coordination.ProbeResult) *sessionCoordinationStatus {
	status := &sessionCoordinationStatus{Busy: result.Busy}
	if !result.Busy || result.Owner == nil {
		return status
	}
	status.ResourceScope = string(result.Owner.ResourceScope)
	status.OperationKind = string(result.Owner.OperationKind)
	status.Command = string(result.Owner.Command)
	status.PID = result.Owner.PID
	status.StartedAt = result.Owner.StartedAt.UTC().Format(time.RFC3339Nano)
	return status
}

func (a *app) writeWorkbookWaitFailure(descriptor coordination.Descriptor, identity coordination.WorkbookIdentity, code, message string, cause error) error {
	details := map[string]any{
		"workbook":       identity.CanonicalPath,
		"operation":      descriptor.ID,
		"resource_scope": descriptor.Policy.ResourceScope,
		"retryable":      descriptor.Policy.RetryableWhenBusy,
		"wait_timeout":   a.waitTimeout.String(),
	}
	commandName := string(descriptor.ID)
	if len(descriptor.CLI) > 0 && strings.TrimSpace(descriptor.CLI[0].Path) != "" {
		commandName = descriptor.CLI[0].Path
	}
	env := output.Failure(commandName, output.Error{
		Code:    code,
		Message: message,
		Source:  "xlflow",
		Phase:   "coordination.acquire",
		Details: details,
	})
	a.addConfigWarnings(&env)
	if err := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); err != nil {
		return output.WithExitCode(output.ExitEnvironment, err)
	}
	return output.WithExitCode(output.ExitEnvironment, cause)
}

func (a *app) writeWorkbookBusyFailure(descriptor coordination.Descriptor, busy *coordination.BusyError) error {
	details := map[string]any{
		"workbook":       busy.Identity.CanonicalPath,
		"operation":      descriptor.ID,
		"resource_scope": descriptor.Policy.ResourceScope,
		"retryable":      descriptor.Policy.RetryableWhenBusy,
	}
	if busy.Owner != nil {
		details["owner"] = busy.Owner
	}
	commandName := string(descriptor.ID)
	if len(descriptor.CLI) > 0 && strings.TrimSpace(descriptor.CLI[0].Path) != "" {
		commandName = descriptor.CLI[0].Path
	}
	env := output.Failure(commandName, output.Error{
		Code:    coordination.WorkbookBusyCode,
		Message: fmt.Sprintf("Another xlflow operation is currently using this workbook: %s. Retry after it completes.", busy.Identity.CanonicalPath),
		Source:  "xlflow",
		Phase:   "coordination.acquire",
		Details: details,
	})
	a.addConfigWarnings(&env)
	if err := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); err != nil {
		return output.WithExitCode(output.ExitEnvironment, err)
	}
	return output.WithExitCode(output.ExitEnvironment, busy)
}
