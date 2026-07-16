package cli

import (
	"errors"
	"fmt"
	"strings"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
)

func (a *app) recoveryCommand() *cobra.Command {
	recovery := &cobra.Command{
		Use:   "recovery",
		Short: "Inspect or clear workbook recovery quarantine",
	}
	recovery.AddCommand(a.recoveryClearCommand())
	return recovery
}

func (a *app) recoveryClearCommand() *cobra.Command {
	var force bool
	cmd := &cobra.Command{
		Use:   "clear",
		Short: "Clear recovery quarantine for the configured workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("recovery clear")
			if err != nil {
				return err
			}
			identity, err := coordination.NewWorkbookIdentity(a.cwd, cfg.Excel.Path)
			if err != nil {
				return a.writeFailure("recovery clear", output.ExitEnvironment, "coordination_identity_failed", err)
			}
			lease := a.activeLeases.Lease(identity)
			if lease == nil {
				return a.writeFailure("recovery clear", output.ExitEnvironment, "coordination_acquire_failed", errors.New("workbook recovery lease is unavailable"))
			}
			state, err := lease.RecoveryState()
			if err != nil {
				return a.writeFailure("recovery clear", output.ExitEnvironment, coordination.RecoveryCheckFailedCode, err)
			}
			if !state.Required {
				env := output.New("recovery clear")
				env.Recovery = map[string]any{
					"required": false,
					"cleared":  false,
					"forced":   force,
					"workbook": identity.CanonicalPath,
				}
				env.Logs = []string{"workbook recovery was not required"}
				return a.write(env, output.ExitSuccess)
			}
			if !force {
				if state.Invalid || state.Metadata == nil || state.Metadata.ExcelPID <= 0 {
					return a.writeRecoveryVerificationFailure(identity, state, "The recovery marker does not contain a verifiable Excel process ID. Use --force only after confirming Excel and VBA are no longer active.")
				}
				running, verifyErr := processRunning(state.Metadata.ExcelPID)
				if verifyErr != nil {
					return a.writeRecoveryVerificationFailure(identity, state, fmt.Sprintf("Excel process verification failed: %v", verifyErr))
				}
				if running {
					return a.writeRecoveryVerificationFailure(identity, state, fmt.Sprintf("Excel process %d is still running.", state.Metadata.ExcelPID))
				}
			}
			expectedGeneration := ""
			if state.Metadata != nil {
				expectedGeneration = state.Metadata.Generation
			}
			cleared, err := lease.ClearRecovery(expectedGeneration)
			if err != nil {
				return a.writeFailure("recovery clear", output.ExitEnvironment, "workbook_recovery_clear_failed", err)
			}
			env := output.New("recovery clear")
			env.Recovery = map[string]any{
				"required": false,
				"cleared":  cleared,
				"forced":   force,
				"workbook": identity.CanonicalPath,
			}
			env.Logs = []string{"cleared workbook recovery quarantine"}
			if force {
				env.Warnings = []map[string]any{{
					"code":    "recovery_force_cleared",
					"message": "The recovery marker was force-cleared. This did not terminate Excel or VBA and did not repair or verify workbook state.",
				}}
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&force, "force", false, "clear the recovery marker without verifying that the affected Excel process stopped")
	return cmd
}

func (a *app) writeRecoveryVerificationFailure(identity coordination.WorkbookIdentity, state coordination.RecoveryState, message string) error {
	details := coordination.RecoveryDetails(identity, state)
	details["retryable"] = false
	env := output.Failure("recovery clear", output.Error{
		Code:    coordination.WorkbookRecoveryVerificationFailedCode,
		Message: strings.TrimSpace(message),
		Source:  "xlflow",
		Phase:   "coordination.recovery",
		Details: details,
	})
	a.addConfigWarnings(&env)
	if err := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); err != nil {
		return output.WithExitCode(output.ExitEnvironment, err)
	}
	return output.WithExitCode(output.ExitEnvironment, errors.New(coordination.WorkbookRecoveryVerificationFailedCode))
}
