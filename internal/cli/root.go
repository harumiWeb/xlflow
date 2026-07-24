package cli

import (
	"bytes"
	"cmp"
	"context"
	"encoding/csv"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"runtime/debug"
	"slices"
	"strconv"
	"strings"
	"time"
	"unicode"
	"unicode/utf8"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"

	"github.com/harumiWeb/xlflow/internal/agentskill"
	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/backup"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/diff"
	"github.com/harumiWeb/xlflow/internal/excel"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	formulaspkg "github.com/harumiWeb/xlflow/internal/formulas"
	"github.com/harumiWeb/xlflow/internal/gui"
	workbookinspect "github.com/harumiWeb/xlflow/internal/inspect"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/lspserver"
	"github.com/harumiWeb/xlflow/internal/output"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
	"github.com/harumiWeb/xlflow/internal/project"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vba/testdiscover"
	"github.com/harumiWeb/xlflow/internal/vbafmt"
	"github.com/harumiWeb/xlflow/internal/workbookformat"
)

type app struct {
	json           bool
	bridge         string
	wait           bool
	waitTimeout    time.Duration
	cwd            string
	rawArgs        []string
	stdout         io.Writer
	stderr         io.Writer
	stdoutTerminal func() bool
	stderrTerminal func() bool
	configWarnings []map[string]any
	buildInfo      BuildInfo
	updateChecker  releaseChecker
	coordination   *coordination.Manager
	activeLeases   *coordination.LeaseSet
}

var automaticBackupPrune = backup.Prune

var (
	errFormMigrateArgs     = errors.New("userform migration arguments invalid")
	errFormMigrateConflict = errors.New("userform migration conflict")
	errFormMigrateInspect  = errors.New("userform designer inspection failed")
)

type BuildInfo struct {
	Version string `json:"version"`
	Commit  string `json:"commit"`
	Date    string `json:"date"`
}

type versionFeature struct {
	Name        string `json:"name"`
	Description string `json:"description"`
}

type versionBuildSetting struct {
	Key   string `json:"key"`
	Value string `json:"value"`
}

type versionVerbosePayload struct {
	BuildInfo
	ExecutablePath string                `json:"executable_path,omitempty"`
	GoVersion      string                `json:"go_version,omitempty"`
	ModulePath     string                `json:"module_path,omitempty"`
	BuildSettings  []versionBuildSetting `json:"build_settings,omitempty"`
	Features       []versionFeature      `json:"features,omitempty"`
}

type updateCheckPayload struct {
	CurrentVersion  string `json:"current_version"`
	LatestVersion   string `json:"latest_version,omitempty"`
	UpdateAvailable bool   `json:"update_available"`
	ReleaseURL      string `json:"release_url,omitempty"`
}

type formSnapshotCommandOptions struct {
	Inspect excel.InspectFormOptions
	OutPath string
}

type formWriteCommandOptions struct {
	Action    string
	SpecInput forms.SpecInput
	Spec      forms.FormSpec
	Overwrite bool
	Session   bool
	NoSave    bool
	Keepalive excel.CommandOptions
}

type formMigrationFile struct {
	Name     string
	FRMPath  string
	FRXPath  string
	CodePath string
	SpecPath string
	Code     string
	Spec     forms.FormSpec
}

type formMigrationRollback struct {
	Path    string
	Existed bool
	Body    []byte
}

type inspectBridgeOptions struct {
	Session   bool
	Keepalive excel.CommandOptions
}

func Execute() error {
	return ExecuteWithBuildInfo(BuildInfo{})
}

func ExecuteWithBuildInfo(info BuildInfo) error {
	cwd, err := os.Getwd()
	if err != nil {
		return output.WithExitCode(output.ExitEnvironment, err)
	}
	a := &app{cwd: cwd, rawArgs: append([]string{}, os.Args[1:]...), buildInfo: info.withDefaults()}
	root := a.rootCommand()
	return a.executeRoot(root)
}

func (info BuildInfo) withDefaults() BuildInfo {
	if strings.TrimSpace(info.Version) == "" {
		info.Version = "dev"
	}
	if strings.TrimSpace(info.Commit) == "" {
		info.Commit = "none"
	}
	if strings.TrimSpace(info.Date) == "" {
		info.Date = "unknown"
	}
	return info
}

func (a *app) rootCommand() *cobra.Command {
	a.buildInfo = a.buildInfo.withDefaults()
	if a.updateChecker == nil {
		a.updateChecker = newGitHubReleaseChecker(nil)
	}
	root := &cobra.Command{
		Use:           "xlflow",
		Short:         "Agent-ready VBA development framework",
		SilenceUsage:  true,
		SilenceErrors: true,
		PersistentPreRunE: func(cmd *cobra.Command, args []string) error {
			if err := a.requireCoordinationPolicy(cmd); err != nil {
				return err
			}
			if err := a.validateCoordinationWaitOptions(cmd); err != nil {
				return err
			}
			return a.delegateWSLCommand(cmd)
		},
	}
	root.PersistentFlags().BoolVar(&a.json, "json", false, "write machine-readable JSON output")
	root.PersistentFlags().StringVar(&a.bridge, "bridge", "", "Excel bridge mode: auto, dotnet")
	root.PersistentFlags().BoolVar(&a.wait, "wait", false, "wait for a busy workbook before starting the command")
	root.PersistentFlags().DurationVar(&a.waitTimeout, "wait-timeout", 30*time.Second, "maximum time to wait for workbook coordination")
	root.AddCommand(
		a.capabilitiesCommand(),
		a.newCommand(),
		a.initCommand(),
		a.doctorCommand(),
		a.attachCommand(),
		a.backupCommand(),
		a.listCommand(),
		a.formCommand(),
		a.formulasCommand(),
		a.pullCommand(),
		a.pushCommand(),
		a.buildCommand(),
		a.packCommand(),
		a.rollbackCommand(),
		a.sessionCommand(),
		a.saveCommand(),
		a.statusCommand(),
		a.runnerCommand(),
		a.runCommand(),
		a.exportImageCommand(),
		a.editCommand(),
		a.macrosCommand(),
		a.uiCommand(),
		a.testCommand(),
		a.typeCommand(),
		a.diffCommand(),
		a.inspectCommand(),
		a.inspectGUICommand(),
		a.lintCommand(),
		a.lspCommand(),
		a.fmtCommand(),
		a.analyzeCommand(),
		a.checkCommand(),
		a.generateCommand(),
		a.moduleCommand(),
		a.skillCommand(),
		a.versionCommand(),
		a.updateCommand(),
		a.processCommand(),
		a.recoveryCommand(),
	)
	a.wrapCoordinatedLeaves(root)
	return root
}

func (a *app) capabilitiesCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "capabilities",
		Short: "Show machine-readable command coordination capabilities",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			capabilities := coordination.PublicCapabilities()
			env := output.New("capabilities")
			env.Capabilities = capabilities
			env.Logs = []string{fmt.Sprintf("capability version: %d", capabilities.CapabilityVersion), fmt.Sprintf("commands: %d", len(capabilities.Commands))}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func (a *app) requireCoordinationPolicy(cmd *cobra.Command) error {
	if isGeneratedCobraCommand(cmd) {
		return nil
	}
	if _, err := coordination.LookupCLI(cmd.CommandPath()); err != nil {
		command := strings.TrimSpace(strings.TrimPrefix(cmd.CommandPath(), "xlflow"))
		return a.writeFailure(command, output.ExitEnvironment, coordination.MissingPolicyCode, err)
	}
	return nil
}

func isGeneratedCobraCommand(cmd *cobra.Command) bool {
	if cmd == nil {
		return false
	}
	current := cmd
	for current != nil {
		switch current.Name() {
		case "completion", "help", cobra.ShellCompRequestCmd, cobra.ShellCompNoDescRequestCmd:
			return true
		}
		current = current.Parent()
	}
	return false
}

func (a *app) executeRoot(root *cobra.Command) error {
	err := root.Execute()
	if err == nil {
		return nil
	}
	if name, ok := unknownCommandName(err); ok {
		return a.writeUnknownCommandFailure(root, name, err)
	}
	return err
}

func unknownCommandName(err error) (string, bool) {
	if err == nil {
		return "", false
	}
	const prefix = "unknown command "
	message := err.Error()
	if !strings.HasPrefix(message, prefix) {
		return "", false
	}
	rest := strings.TrimPrefix(message, prefix)
	quoted, _, ok := strings.Cut(rest, " for ")
	if !ok {
		return "", false
	}
	name, quoteErr := strconv.Unquote(quoted)
	if quoteErr != nil || name == "" {
		return "", false
	}
	return name, true
}

func (a *app) writeUnknownCommandFailure(root *cobra.Command, name string, err error) error {
	suggestions := root.SuggestionsFor(name)
	env := output.Failure("xlflow", output.Error{
		Code:        "unknown_command",
		Message:     fmt.Sprintf("unknown command %q", name),
		Suggestions: suggestions,
	})
	if jsonOutput := a.json || rawArgsRequestJSON(a.rawArgs); jsonOutput {
		opts := a.outputOptions()
		opts.JSON = true
		if writeErr := output.WriteWithOptions(a.stdoutWriter(), env, opts); writeErr != nil {
			return output.WithExitCode(output.ExitConfig, writeErr)
		}
		return output.WithExitCode(output.ExitConfig, err)
	}
	if writeErr := writeUnknownCommandText(a.stderrWriter(), root.CommandPath(), name, suggestions); writeErr != nil {
		return output.WithExitCode(output.ExitConfig, writeErr)
	}
	return output.WithExitCode(output.ExitConfig, err)
}

func writeUnknownCommandText(w io.Writer, commandPath string, name string, suggestions []string) error {
	if commandPath == "" {
		commandPath = "xlflow"
	}
	if _, err := fmt.Fprintf(w, "Error: unknown command %q for %q\n", name, commandPath); err != nil {
		return err
	}
	if len(suggestions) > 0 {
		if _, err := fmt.Fprintln(w, "\nDid you mean this?"); err != nil {
			return err
		}
		for _, suggestion := range suggestions {
			if _, err := fmt.Fprintf(w, "  %s\n", suggestion); err != nil {
				return err
			}
		}
		if _, err := fmt.Fprintln(w); err != nil {
			return err
		}
	}
	_, err := fmt.Fprintf(w, "Run %q for usage.\n", commandPath+" --help")
	return err
}

func rawArgsRequestJSON(args []string) bool {
	for i := 0; i < len(args); i++ {
		arg := args[i]
		if arg == "--json" {
			return true
		}
		if strings.HasPrefix(arg, "--json=") {
			value := strings.TrimPrefix(arg, "--json=")
			parsed, err := strconv.ParseBool(value)
			return err == nil && parsed
		}
		if arg == "--bridge" {
			i++
			continue
		}
		if strings.HasPrefix(arg, "--bridge=") {
			continue
		}
		if strings.HasPrefix(arg, "-") {
			continue
		}
		return false
	}
	return false
}

func (a *app) versionCommand() *cobra.Command {
	var verbose bool
	cmd := &cobra.Command{
		Use:   "version",
		Short: "Show xlflow build information",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			env := output.New("version")
			env.Version = a.versionPayload(verbose)
			env.Logs = a.versionLogs(verbose)
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&verbose, "verbose", false, "show executable path, build settings, script resolution, and supported features")
	return cmd
}

func (a *app) updateCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "update",
		Short: "Check for xlflow updates",
	}
	cmd.AddCommand(a.updateCheckCommand())
	return cmd
}

func (a *app) updateCheckCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "check",
		Short: "Check the latest xlflow GitHub release",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			env, exitCode := a.updateCheckEnvelope(cmd.Context())
			return a.write(env, exitCode)
		},
	}
}

func (a *app) updateCheckEnvelope(ctx context.Context) (output.Envelope, int) {
	info := a.buildInfo.withDefaults()
	payload := updateCheckPayload{
		CurrentVersion: info.Version,
	}
	if a.updateChecker == nil {
		a.updateChecker = newGitHubReleaseChecker(nil)
	}
	checkCtx, cancel := context.WithTimeout(ctx, updateCheckTimeout)
	defer cancel()
	update, err := checkForUpdate(checkCtx, a.updateChecker, info.Version)
	if err != nil {
		env := output.Failure("update check", output.Error{
			Code:    "update_check_failed",
			Message: "failed to check for xlflow updates: " + err.Error(),
			Phase:   "update_check",
		})
		env.Update = payload
		return env, output.ExitEnvironment
	}
	payload.LatestVersion = update.LatestVersion
	payload.UpdateAvailable = update.LatestVersion != ""
	payload.ReleaseURL = update.ReleaseURL
	env := output.New("update check")
	env.Update = payload
	if payload.UpdateAvailable {
		env.Logs = []string{
			"current version: " + payload.CurrentVersion,
			"latest version: " + payload.LatestVersion,
		}
		if payload.ReleaseURL != "" {
			env.Logs = append(env.Logs, "release: "+payload.ReleaseURL)
		}
	} else {
		env.Logs = []string{"xlflow is up to date"}
	}
	return env, output.ExitSuccess
}

func (a *app) versionPayload(verbose bool) any {
	info := a.buildInfo.withDefaults()
	if !verbose {
		return info
	}
	payload := versionVerbosePayload{
		BuildInfo:      info,
		ExecutablePath: resolvedExecutablePath(),
		Features:       supportedVersionFeatures(),
	}
	if buildInfo, ok := debug.ReadBuildInfo(); ok {
		payload.GoVersion = buildInfo.GoVersion
		payload.ModulePath = buildInfo.Main.Path
		payload.BuildSettings = buildSettingsFromInfo(buildInfo)
	}
	return payload
}

func (a *app) versionLogs(verbose bool) []string {
	info := a.buildInfo.withDefaults()
	logs := []string{
		"version: " + info.Version,
		"commit: " + info.Commit,
		"date: " + info.Date,
	}
	if !verbose {
		return logs
	}
	if exe := resolvedExecutablePath(); exe != "" {
		logs = append(logs, "executable: "+exe)
	}
	if features := supportedVersionFeatures(); len(features) > 0 {
		logs = append(logs, fmt.Sprintf("features: %d supported", len(features)))
	}
	return logs
}

func resolvedExecutablePath() string {
	exe, err := os.Executable()
	if err != nil {
		return ""
	}
	return exe
}

func supportedVersionFeatures() []versionFeature {
	return []versionFeature{
		{Name: "version-verbose", Description: "Show executable path, build settings, script resolution, and supported features."},
		{Name: "vba-lsp-stdio", Description: "Run a reusable VBA language server over stdio for editor and agent integrations."},
		{Name: "auto-session-reuse", Description: "Reuse a matching active xlflow session workbook when --session is omitted for workbook-backed commands."},
		{Name: "save-state-visibility", Description: "Return structured save-required state when a live session workbook differs from disk."},
		{Name: "push-save-default", Description: "Save workbook changes by default after push unless --no-save opts out during a session."},
		{Name: "run-entry-fallback", Description: "Use project.entry from xlflow.toml when xlflow run is invoked without a macro argument."},
		{Name: "diagnostic-run", Description: "Compile before run and return structured VBA compile diagnostics by default."},
		{Name: "range-image-export", Description: "Export a worksheet range to a PNG image for visual verification."},
		{Name: "form-image-export", Description: "Export a runtime UserForm to a PNG image for visual verification."},
		{Name: "form-build-overwrite", Description: "Create or replace Designer-backed UserForms from persisted xlflow.userform specs."},
		{Name: "workbook-edit-helpers", Description: "Mutate a live session workbook for agent-driven test setup, event triggering, and visual tuning."},
	}
}

func buildSettingsFromInfo(info *debug.BuildInfo) []versionBuildSetting {
	if info == nil || len(info.Settings) == 0 {
		return nil
	}
	settings := make([]versionBuildSetting, 0, len(info.Settings))
	for _, setting := range info.Settings {
		switch setting.Key {
		case "vcs", "vcs.revision", "vcs.time", "vcs.modified", "GOARCH", "GOOS":
			settings = append(settings, versionBuildSetting{Key: setting.Key, Value: setting.Value})
		}
	}
	slices.SortFunc(settings, func(a, b versionBuildSetting) int {
		return cmp.Compare(a.Key, b.Key)
	})
	return settings
}

func (a *app) macrosCommand() *cobra.Command {
	var session bool
	var runnable bool
	cmd := &cobra.Command{
		Use:   "macros",
		Short: "Discover runnable workbook macros",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("macros")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Reading VBA project", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).MacrosWithOptions(cfg, excel.MacrosOptions{Session: session, Entry: cfg.Project.Entry, RunnableOnly: runnable, Keepalive: commandOpts})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	cmd.Flags().BoolVar(&runnable, "runnable", false, "show only runnable macros")
	return cmd
}

func (a *app) listCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "list",
		Short: "List workbook resources",
	}
	cmd.AddCommand(a.listFormsCommand())
	return cmd
}

func (a *app) backupCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "backup",
		Short: "Manage workbook backups",
	}
	cmd.AddCommand(a.backupListCommand(), a.backupPruneCommand(), a.backupDeleteCommand())
	return cmd
}

func (a *app) backupListCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "list",
		Short: "List available workbook backups",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("backup list")
			if err != nil {
				return err
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			scan, err := backup.Scan(a.cwd, workbookPath)
			if err != nil {
				return a.writeFailure("backup list", output.ExitEnvironment, "backup_scan_failed", err)
			}
			env := output.New("backup list")
			env.Backups = renderBackupRecords(a.cwd, scan.Records)
			env.Warnings = renderInvalidBackupWarnings(a.cwd, scan.Invalid)
			env.Logs = []string{fmt.Sprintf("found %d backup(s)", len(scan.Records))}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func (a *app) backupPruneCommand() *cobra.Command {
	var keepLast int
	var olderThan string
	var maxTotalSize string
	var dryRun bool
	var allWorkbooks bool
	var includeInvalid bool
	var includeLegacy bool
	cmd := &cobra.Command{
		Use:   "prune",
		Short: "Preview or delete old workbook backups",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			opts := backup.PruneOptions{
				DryRun:         dryRun,
				AllWorkbooks:   allWorkbooks,
				IncludeInvalid: includeInvalid,
				IncludeLegacy:  includeLegacy,
			}
			if cmd.Flags().Changed("keep-last") {
				opts.KeepLast = &keepLast
			}
			if strings.TrimSpace(olderThan) != "" {
				duration, err := backup.ParseRetentionDuration(olderThan)
				if err != nil {
					return a.writeFailure("backup prune", output.ExitConfig, backup.ErrPruneArgsInvalid, err)
				}
				opts.OlderThan = duration
				opts.OlderThanSet = true
			}
			if strings.TrimSpace(maxTotalSize) != "" {
				size, err := backup.ParseSize(maxTotalSize)
				if err != nil {
					return a.writeFailure("backup prune", output.ExitConfig, backup.ErrPruneArgsInvalid, err)
				}
				opts.MaxTotalSize = size
				opts.MaxTotalSizeSet = true
			}
			cfg, err := a.loadConfig("backup prune")
			if err != nil {
				return err
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			result, err := backup.Prune(a.cwd, workbookPath, opts)
			env := output.New("backup prune")
			env.BackupPrune = renderBackupPruneResult(a.cwd, result)
			if err != nil {
				code := output.ExitEnvironment
				if backupErrorCode(err) == backup.ErrPruneArgsInvalid {
					code = output.ExitConfig
				}
				env.Status = output.StatusFailed
				env.Error = &output.Error{Code: backupErrorCode(err), Message: err.Error()}
				return a.write(env, code)
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().IntVar(&keepLast, "keep-last", 0, "retain the newest N backups")
	cmd.Flags().StringVar(&olderThan, "older-than", "", "delete backups older than a duration such as 30d")
	cmd.Flags().StringVar(&maxTotalSize, "max-total-size", "", "delete oldest backups until total storage is below a size such as 2GB")
	cmd.Flags().BoolVar(&dryRun, "dry-run", false, "preview candidates without deleting files")
	cmd.Flags().BoolVar(&allWorkbooks, "all-workbooks", false, "evaluate backups for all managed workbooks")
	cmd.Flags().BoolVar(&includeInvalid, "include-invalid", false, "include invalid managed backup directories")
	cmd.Flags().BoolVar(&includeLegacy, "include-legacy", false, "include legacy managed backup directories without metadata")
	return cmd
}

func (a *app) backupDeleteCommand() *cobra.Command {
	var backupID string
	cmd := &cobra.Command{
		Use:   "delete",
		Short: "Delete one managed workbook backup",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			if strings.TrimSpace(backupID) == "" {
				return a.writeFailure("backup delete", output.ExitConfig, backup.ErrDeleteArgsInvalid, fmt.Errorf("--backup is required"))
			}
			cfg, err := a.loadConfig("backup delete")
			if err != nil {
				return err
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			result, err := backup.Delete(a.cwd, workbookPath, backupID)
			if err != nil {
				return a.writeFailure("backup delete", backupDeleteExitCode(err), backupErrorCode(err), err)
			}
			env := output.New("backup delete")
			env.BackupDelete = map[string]any{
				"id":          result.ID,
				"path":        displayPath(a.cwd, result.Path),
				"freed_bytes": result.FreedBytes,
			}
			env.Logs = []string{"deleted backup " + result.ID}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&backupID, "backup", "", "delete a specific backup ID")
	return cmd
}

func (a *app) rollbackCommand() *cobra.Command {
	var latest bool
	var backupID string
	cmd := &cobra.Command{
		Use:   "rollback",
		Short: "Restore the workbook from a saved backup",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			target, err := buildRollbackTarget(latest, backupID)
			if err != nil {
				return a.writeFailure("rollback", output.ExitConfig, "rollback_args_invalid", err)
			}
			cfg, err := a.loadConfig("rollback")
			if err != nil {
				return err
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			blocked, err := a.rollbackBlockedBySession(workbookPath)
			if err != nil {
				return a.writeFailure("rollback", output.ExitEnvironment, "rollback_session_check_failed", err)
			}
			if blocked {
				return a.writeFailure("rollback", output.ExitValidation, "workbook_in_use", fmt.Errorf("the configured workbook is attached to an active xlflow session; stop the session before rollback"))
			}
			record, err := resolveRollbackRecord(a.cwd, workbookPath, target)
			if err != nil {
				return a.writeFailure("rollback", output.ExitValidation, "backup_not_found", err)
			}
			safety, err := backup.Create(a.cwd, workbookPath, "pre-rollback", time.Now())
			if err != nil {
				return a.writeFailure("rollback", output.ExitEnvironment, "rollback_backup_failed", err)
			}
			if err := backup.Restore(workbookPath, record); err != nil {
				code := "rollback_failed"
				exitCode := output.ExitEnvironment
				if looksLikeWorkbookInUse(err) {
					code = "workbook_in_use"
					exitCode = output.ExitValidation
				}
				return a.writeFailure("rollback", exitCode, code, err)
			}
			env := output.New("rollback")
			env.Target = map[string]any{"kind": "file", "path": displayPath(a.cwd, workbookPath)}
			env.Rollback = map[string]any{
				"restored_from": map[string]any{
					"id":         record.ID,
					"path":       displayPath(a.cwd, record.BackupFileAbsPath),
					"reason":     record.Reason,
					"created_at": record.CreatedAt.Format(time.RFC3339),
				},
				"safety_backup": map[string]any{
					"id":   safety.ID,
					"path": displayPath(a.cwd, safety.BackupFileAbsPath),
				},
				"target": map[string]any{
					"path": displayPath(a.cwd, workbookPath),
				},
			}
			env.Warnings = []map[string]any{
				{
					"code":    "source_out_of_sync",
					"message": "Rollback restored only the workbook file. Source files under `src/` were not changed and may now be out of sync.",
				},
			}
			env.Hints = []map[string]any{
				{"code": "verify_workbook", "message": "Run `xlflow inspect --json` to verify the restored workbook state."},
				{"code": "sync_source", "message": "Run `xlflow pull --json` if you want source files to match the restored workbook."},
			}
			env.Logs = []string{"restored workbook from backup"}
			a.applyAutomaticBackupRetention(cfg, workbookPath, &env, true)
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&latest, "latest", false, "restore the latest backup for the configured workbook")
	cmd.Flags().StringVar(&backupID, "backup", "", "restore a specific backup ID")
	return cmd
}

func (a *app) formCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "form",
		Short: "Manage workbook UserForms",
	}
	cmd.AddCommand(a.formNewCommand(), a.formMigrateCommand(), a.formSnapshotCommand(), a.formBuildCommand(), a.formApplyCommand(), a.formExportImageCommand())
	return cmd
}

func (a *app) formMigrateCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "migrate",
		Short: "Migrate UserForm source layouts",
	}
	cmd.AddCommand(a.formMigrateSidecarCommand())
	return cmd
}

func (a *app) formMigrateSidecarCommand() *cobra.Command {
	var overwrite bool
	var session bool
	cmd := &cobra.Command{
		Use:   "sidecar [FormName]",
		Short: "Migrate frm-mode UserForms to sidecar code and Designer specs",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			target := ""
			if len(args) > 0 {
				target = strings.TrimSpace(args[0])
			}
			if target != "" && !isBareComponentName(target) {
				return a.writeFailure("form migrate sidecar", output.ExitConfig, "form_migrate_args_invalid", fmt.Errorf("form name must be a bare VBA component name"))
			}
			cfg, err := a.loadConfig("form migrate sidecar")
			if err != nil {
				return err
			}
			before := cfg.UserForm.CodeSource
			if before != "frm" && before != "sidecar" {
				return a.writeFailure("form migrate sidecar", output.ExitConfig, "form_migrate_args_invalid", fmt.Errorf("userform.code_source must be one of frm, sidecar"))
			}
			if err := a.rejectStaleSourceForFormMigration(cfg); err != nil {
				return a.writeFailure("form migrate sidecar", output.ExitValidation, "form_migrate_conflict", err)
			}
			result, err := a.prepareUserFormSidecarMigration(cfg, target, overwrite, session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				if errors.Is(err, errFormMigrateArgs) {
					return a.writeFailure("form migrate sidecar", output.ExitConfig, "form_migrate_args_invalid", err)
				}
				if errors.Is(err, errFormMigrateConflict) {
					return a.writeFailure("form migrate sidecar", output.ExitValidation, "form_migrate_conflict", err)
				}
				if errors.Is(err, errFormMigrateInspect) {
					return a.writeFailure("form migrate sidecar", output.ExitEnvironment, "form_migrate_inspect_failed", err)
				}
				return a.writeFailure("form migrate sidecar", output.ExitEnvironment, "form_migrate_failed", err)
			}
			created, updated, skipped, err := a.writeUserFormSidecarMigration(before, result, overwrite)
			if err != nil {
				if errors.Is(err, errFormMigrateConflict) {
					return a.writeFailure("form migrate sidecar", output.ExitValidation, "form_migrate_conflict", err)
				}
				return a.writeFailure("form migrate sidecar", output.ExitEnvironment, "form_migrate_failed", err)
			}
			env := output.New("form migrate sidecar")
			env.Source = map[string]any{
				"operation":          "userform.migrate_sidecar",
				"code_source_before": before,
				"code_source_after":  "sidecar",
				"forms":              renderFormMigrationFiles(a.cwd, result),
				"created":            created,
				"updated":            updated,
				"skipped":            skipped,
				"config_path":        config.FileName,
				"requires_push":      false,
			}
			env.Logs = []string{fmt.Sprintf("migrated %d UserForm(s) to sidecar mode", len(result))}
			if before != "sidecar" {
				env.Logs = append(env.Logs, "updated [userform].code_source to sidecar")
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&overwrite, "overwrite", false, "replace existing sidecar code or Designer spec files")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) formulasCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "formulas",
		Short: "Manage workbook formula snapshots",
	}
	cmd.AddCommand(a.formulasPullCommand())
	cmd.AddCommand(a.formulasInspectCommand())
	return cmd
}

func (a *app) formulasPullCommand() *cobra.Command {
	var srcPath string
	var outDir string
	cmd := &cobra.Command{
		Use:   "pull",
		Short: "Extract workbook formulas into region JSONL snapshots",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			workbookPath := strings.TrimSpace(srcPath)
			if workbookPath == "" {
				cfg, err := a.loadConfig("formulas pull")
				if err != nil {
					return err
				}
				workbookPath = cfg.Excel.Path
			}
			workbookPath = workbookArgPath(a.cwd, workbookPath)
			if err := workbookformat.ValidateFormulaSnapshotWorkbook(workbookPath); err != nil {
				var unsupported workbookformat.UnsupportedError
				if errors.As(err, &unsupported) {
					return a.writeUnsupportedWorkbookFormat("formulas pull", unsupported)
				}
				return a.writeFailure("formulas pull", output.ExitConfig, "formulas_pull_args_invalid", err)
			}
			outputDir := workbookArgPath(a.cwd, strings.TrimSpace(outDir))
			if outputDir == "" {
				outputDir = filepath.Join(a.cwd, "formulas")
			}
			result, err := formulaspkg.Pull(workbookPath, outputDir)
			if err != nil {
				return a.writeFailure("formulas pull", output.ExitEnvironment, "formulas_pull_failed", err)
			}
			env := output.New("formulas pull")
			env.Workbook = map[string]any{
				"path": displayPath(a.cwd, workbookPath),
			}
			env.Output = formulaResultPayload(result, a.cwd)
			env.Logs = []string{fmt.Sprintf("extracted %d formula region(s) from %d sheet(s)", result.FormulaRegionCount, len(result.Manifest.Sheets))}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&srcPath, "src", "", "source workbook path; when omitted, use [excel].path from xlflow.toml")
	cmd.Flags().StringVar(&outDir, "out", "formulas", "output directory for formula snapshots")
	return cmd
}

func (a *app) formulasInspectCommand() *cobra.Command {
	var dir string
	var summary bool
	var sheet string
	var cell string
	var cellRange string
	cmd := &cobra.Command{
		Use:   "inspect",
		Short: "Inspect formula snapshot regions",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			view, err := buildFormulaInspectView(summary, sheet, cell, cellRange)
			if err != nil {
				return a.writeFailure("formulas inspect", output.ExitConfig, "formulas_inspect_args_invalid", err)
			}
			snapshotDir := workbookArgPath(a.cwd, strings.TrimSpace(dir))
			if snapshotDir == "" {
				snapshotDir = filepath.Join(a.cwd, "formulas")
			}
			var result formulaspkg.InspectResult
			switch view.kind {
			case "summary":
				result, err = formulaspkg.InspectSummary(snapshotDir)
			case "sheet":
				result, err = formulaspkg.InspectSheet(snapshotDir, view.value)
			case "cell":
				result, err = formulaspkg.InspectCell(snapshotDir, view.value)
			case "range":
				result, err = formulaspkg.InspectRange(snapshotDir, view.value)
			default:
				err = fmt.Errorf("unsupported formula inspect view %q", view.kind)
			}
			if err != nil {
				code := output.ExitEnvironment
				if formulaspkg.IsInspectArgumentError(err) {
					code = output.ExitConfig
				}
				errCode := "formulas_inspect_failed"
				if code == output.ExitConfig {
					errCode = "formulas_inspect_args_invalid"
				}
				return a.writeFailure("formulas inspect", code, errCode, err)
			}
			result.Dir = displayPath(a.cwd, snapshotDir)
			env := output.New("formulas inspect")
			env.Output = map[string]any{"formulas_inspect": result}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&dir, "dir", "formulas", "formula snapshot directory")
	cmd.Flags().BoolVar(&summary, "summary", false, "show workbook-level formula summary")
	cmd.Flags().StringVar(&sheet, "sheet", "", "show formula regions for one sheet")
	cmd.Flags().StringVar(&cell, "cell", "", "show the formula region containing a cell such as Invoice!E500")
	cmd.Flags().StringVar(&cellRange, "range", "", "show formula regions overlapping a range such as Invoice!D2:F1001")
	return cmd
}

type formulaInspectView struct {
	kind  string
	value string
}

func buildFormulaInspectView(summary bool, sheet, cell, cellRange string) (formulaInspectView, error) {
	selected := 0
	if summary {
		selected++
	}
	if strings.TrimSpace(sheet) != "" {
		selected++
	}
	if strings.TrimSpace(cell) != "" {
		selected++
	}
	if strings.TrimSpace(cellRange) != "" {
		selected++
	}
	if selected > 1 {
		return formulaInspectView{}, fmt.Errorf("choose only one of --summary, --sheet, --cell, or --range")
	}
	switch {
	case strings.TrimSpace(sheet) != "":
		return formulaInspectView{kind: "sheet", value: strings.TrimSpace(sheet)}, nil
	case strings.TrimSpace(cell) != "":
		return formulaInspectView{kind: "cell", value: strings.TrimSpace(cell)}, nil
	case strings.TrimSpace(cellRange) != "":
		return formulaInspectView{kind: "range", value: strings.TrimSpace(cellRange)}, nil
	default:
		return formulaInspectView{kind: "summary"}, nil
	}
}

func (a *app) formNewCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "new <name>",
		Short: "Create source files for a sidecar UserForm",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("form new")
			if err != nil {
				return err
			}
			result, err := project.NewUserForm(a.cwd, args[0], cfg)
			if err != nil {
				if errors.Is(err, project.ErrUserFormRequiresSidecar) {
					return a.writeFailure("form new", output.ExitConfig, "form_new_requires_sidecar", err)
				}
				if errors.Is(err, project.ErrInvalidComponentName) {
					return a.writeFailure("form new", output.ExitConfig, "form_new_args_invalid", err)
				}
				if errors.Is(err, project.ErrScaffoldExists) {
					return a.writeFailure("form new", output.ExitValidation, "form_new_failed", err)
				}
				return a.writeFailure("form new", output.ExitEnvironment, "form_new_failed", err)
			}
			env := output.New("form new")
			env.Source = map[string]any{
				"created":     result.Created,
				"kind":        "form",
				"name":        result.Name,
				"code_path":   result.CodePath,
				"spec_path":   result.SpecPath,
				"code_source": result.CodeSource,
			}
			env.Logs = []string{fmt.Sprintf("created UserForm source: %s", result.Name)}
			return a.write(env, output.ExitSuccess)
		},
	}
	return cmd
}

func (a *app) formSnapshotCommand() *cobra.Command {
	var session bool
	var outPath string
	cmd := &cobra.Command{
		Use:   "snapshot <name>",
		Short: "Write a designer UserForm snapshot spec to a file",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			opts, err := buildFormSnapshotOptions(args[0], outPath, session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("form snapshot", output.ExitConfig, "form_snapshot_args_invalid", err)
			}
			resolvedOutput, err := forms.ResolveSnapshotOutput(a.cwd, opts.OutPath)
			if err != nil {
				return a.writeFailure("form snapshot", output.ExitConfig, "form_snapshot_args_invalid", err)
			}
			cfg, err := a.loadConfig("form snapshot")
			if err != nil {
				return err
			}
			var scriptEnv output.Envelope
			var code int
			err = a.withExcelProgress("Snapshotting workbook form", opts.Inspect.Keepalive, func() error {
				var runErr error
				scriptEnv, code, runErr = a.excelRunnerForConfig(cfg).InspectForm(cfg, opts.Inspect)
				return runErr
			})
			if err != nil {
				return err
			}
			if code != output.ExitSuccess {
				scriptEnv.Command = "form snapshot"
				return a.write(scriptEnv, code)
			}
			spec, err := forms.FormSpecFromInspectSnapshot(scriptEnv.Forms)
			if err != nil {
				return a.writeFailure("form snapshot", output.ExitEnvironment, "form_snapshot_write_failed", err)
			}
			if err := forms.WriteSnapshot(resolvedOutput, spec); err != nil {
				return a.writeFailure("form snapshot", output.ExitEnvironment, "form_snapshot_write_failed", err)
			}
			env := output.New("form snapshot")
			env.Target = scriptEnv.Target
			env.Session = scriptEnv.Session
			env.Workbook = scriptEnv.Workbook
			env.Warnings = scriptEnv.Warnings
			env.Hints = scriptEnv.Hints
			env.Output = map[string]any{
				"path":   resolvedOutput.DisplayPath,
				"format": resolvedOutput.Format,
			}
			formSummary := map[string]any{
				"name":              spec.Form.Name,
				"basis":             spec.Basis,
				"coordinate_system": spec.CoordinateSystem,
				"control_count":     len(spec.Controls),
			}
			if spec.Form.Caption != nil {
				formSummary["caption"] = *spec.Form.Caption
			}
			if spec.Form.Width != nil {
				formSummary["width"] = *spec.Form.Width
			}
			if spec.Form.Height != nil {
				formSummary["height"] = *spec.Form.Height
			}
			env.Forms = formSummary
			env.Logs = []string{fmt.Sprintf("wrote %s UserForm snapshot for %s to %s", spec.Basis, spec.Form.Name, resolvedOutput.DisplayPath)}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&outPath, "out", "", "write the snapshot spec to a .json, .yaml, or .yml file")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) formBuildCommand() *cobra.Command {
	var session bool
	var overwrite bool
	var noSave bool
	cmd := &cobra.Command{
		Use:   "build <spec.json|spec.yaml|spec.yml>",
		Short: "Create a UserForm from a persisted spec",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			opts, err := buildFormWriteOptions("build", args[0], overwrite, session, noSave, buildCommandOptions(a.stderrWriter()), a.cwd)
			if err != nil {
				var specErr *forms.SpecError
				if errors.As(err, &specErr) {
					return a.writeFormSpecFailure("form build", specErr)
				}
				return a.writeFailure("form build", output.ExitConfig, "form_build_args_invalid", err)
			}
			cfg, err := a.loadConfig("form build")
			if err != nil {
				return err
			}
			if err := a.runFormWritePreflight("form build", cfg, opts); err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Building workbook form", opts.Keepalive, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).FormWrite(cfg, excel.FormWriteOptions{
					Action:    opts.Action,
					SpecPath:  opts.SpecInput.DisplayPath,
					Spec:      opts.Spec,
					Overwrite: opts.Overwrite,
					Session:   opts.Session,
					NoSave:    opts.NoSave,
					Keepalive: opts.Keepalive,
				})
				return runErr
			})
			if err != nil {
				return err
			}
			appendFormSpecValidationWarnings(&env, opts.Spec.ValidationWarnings)
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&overwrite, "overwrite", false, "replace an existing UserForm with the same spec form name")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	cmd.Flags().BoolVar(&noSave, "no-save", false, "leave session-backed workbook changes unsaved until xlflow save")
	return cmd
}

func (a *app) formApplyCommand() *cobra.Command {
	var session bool
	var noSave bool
	cmd := &cobra.Command{
		Use:    "apply <spec.json|spec.yaml|spec.yml>",
		Short:  "Apply a persisted spec to an existing UserForm",
		Hidden: true,
		Args:   cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			opts, err := buildFormWriteOptions("apply", args[0], false, session, noSave, buildCommandOptions(a.stderrWriter()), a.cwd)
			if err != nil {
				var specErr *forms.SpecError
				if errors.As(err, &specErr) {
					return a.writeFormSpecFailure("form apply", specErr)
				}
				return a.writeFailure("form apply", output.ExitConfig, "form_apply_args_invalid", err)
			}
			cfg, err := a.loadConfig("form apply")
			if err != nil {
				return err
			}
			if err := a.runFormWritePreflight("form apply", cfg, opts); err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Applying workbook form spec", opts.Keepalive, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).FormWrite(cfg, excel.FormWriteOptions{
					Action:    opts.Action,
					SpecPath:  opts.SpecInput.DisplayPath,
					Spec:      opts.Spec,
					Overwrite: opts.Overwrite,
					Session:   opts.Session,
					NoSave:    opts.NoSave,
					Keepalive: opts.Keepalive,
				})
				return runErr
			})
			if err != nil {
				return err
			}
			appendFormSpecValidationWarnings(&env, opts.Spec.ValidationWarnings)
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	cmd.Flags().BoolVar(&noSave, "no-save", false, "leave session-backed workbook changes unsaved until xlflow save")
	return cmd
}

func (a *app) formExportImageCommand() *cobra.Command {
	var session bool
	var overwrite bool
	var outPath string
	var initializer string
	cmd := &cobra.Command{
		Use:   "export-image <name>",
		Short: "Export a runtime UserForm as a PNG image",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			opts, err := buildFormExportImageOptions(args[0], outPath, initializer, overwrite, session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("form export-image", output.ExitConfig, "form_export_image_args_invalid", err)
			}
			cfg, err := a.loadConfig("form export-image")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Exporting workbook form image", opts.Keepalive, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).FormExportImage(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&outPath, "out", "", "write the PNG image to an explicit file path")
	cmd.Flags().StringVar(&initializer, "initializer", "", "optional public form method to invoke with ThisWorkbook before capture")
	cmd.Flags().BoolVar(&overwrite, "overwrite", false, "replace an existing output file")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) listFormsCommand() *cobra.Command {
	var session bool
	cmd := &cobra.Command{
		Use:   "forms",
		Short: "List workbook UserForms",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("list")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Listing workbook forms", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).ListForms(cfg, excel.SessionCommandOptions{Session: session, Keepalive: commandOpts})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) uiCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "ui",
		Short: "Manage workbook UI controls",
	}
	cmd.AddCommand(a.uiButtonCommand())
	return cmd
}

func (a *app) uiButtonCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "button",
		Short: "Manage xlflow workbook buttons",
	}
	cmd.AddCommand(
		a.uiButtonAddCommand(),
		a.uiButtonListCommand(),
		a.uiButtonRemoveCommand(),
	)
	return cmd
}

func (a *app) uiButtonAddCommand() *cobra.Command {
	var opts excel.UIButtonAddOptions
	cmd := &cobra.Command{
		Use:   "add",
		Short: "Add or update a workbook form-control button",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			built, err := buildUIButtonAddOptions(opts)
			if err != nil {
				return a.writeFailure("ui", output.ExitConfig, "ui_button_args_invalid", err)
			}
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("ui")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Adding workbook button", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).UIButtonAdd(cfg, built, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&opts.Sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&opts.Cell, "cell", "", "top-left cell address such as B2")
	cmd.Flags().StringVar(&opts.Text, "text", "", "button caption")
	cmd.Flags().StringVar(&opts.Macro, "macro", "", "macro assigned to the button, such as Main.Run")
	cmd.Flags().StringVar(&opts.ID, "id", "", "stable xlflow button id")
	cmd.Flags().IntVar(&opts.Width, "width", 160, "button width in points")
	cmd.Flags().IntVar(&opts.Height, "height", 40, "button height in points")
	cmd.Flags().BoolVar(&opts.CreateSheet, "create-sheet", false, "create the target worksheet when it does not exist")
	cmd.Flags().BoolVar(&opts.VerifyMacro, "verify-macro", false, "verify that the assigned macro exists before saving")
	cmd.Flags().BoolVar(&opts.Session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) uiButtonListCommand() *cobra.Command {
	var opts excel.UIButtonListOptions
	cmd := &cobra.Command{
		Use:   "list",
		Short: "List xlflow-managed workbook buttons",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("ui")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Listing workbook buttons", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).UIButtonList(cfg, opts, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&opts.Sheet, "sheet", "", "worksheet name")
	cmd.Flags().BoolVar(&opts.Session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) uiButtonRemoveCommand() *cobra.Command {
	var opts excel.UIButtonRemoveOptions
	cmd := &cobra.Command{
		Use:   "remove",
		Short: "Remove an xlflow-managed workbook button",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			built, err := buildUIButtonRemoveOptions(opts)
			if err != nil {
				return a.writeFailure("ui", output.ExitConfig, "ui_button_args_invalid", err)
			}
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("ui")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Removing workbook button", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).UIButtonRemove(cfg, built, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&opts.ID, "id", "", "stable xlflow button id")
	cmd.Flags().StringVar(&opts.Sheet, "sheet", "", "worksheet name")
	cmd.Flags().BoolVar(&opts.Session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func buildUIButtonAddOptions(opts excel.UIButtonAddOptions) (excel.UIButtonAddOptions, error) {
	if strings.TrimSpace(opts.Sheet) == "" {
		return opts, fmt.Errorf("--sheet is required")
	}
	if strings.TrimSpace(opts.Cell) == "" {
		return opts, fmt.Errorf("--cell is required")
	}
	if strings.TrimSpace(opts.Text) == "" {
		return opts, fmt.Errorf("--text is required")
	}
	if strings.TrimSpace(opts.Macro) == "" {
		return opts, fmt.Errorf("--macro is required")
	}
	if opts.Width <= 0 {
		return opts, fmt.Errorf("--width must be greater than 0")
	}
	if opts.Height <= 0 {
		return opts, fmt.Errorf("--height must be greater than 0")
	}
	opts.Sheet = strings.TrimSpace(opts.Sheet)
	opts.Cell = strings.TrimSpace(opts.Cell)
	opts.Text = strings.TrimSpace(opts.Text)
	opts.Macro = strings.TrimSpace(opts.Macro)
	opts.ID = normalizeUIButtonID(opts.ID)
	if opts.ID == "" {
		opts.ID = normalizeUIButtonID(opts.Macro)
	}
	if opts.ID == "" {
		return opts, fmt.Errorf("--id could not be derived from --macro")
	}
	return opts, nil
}

func buildUIButtonRemoveOptions(opts excel.UIButtonRemoveOptions) (excel.UIButtonRemoveOptions, error) {
	opts.ID = normalizeUIButtonID(opts.ID)
	opts.Sheet = strings.TrimSpace(opts.Sheet)
	if opts.ID == "" {
		return opts, fmt.Errorf("--id is required")
	}
	return opts, nil
}

func normalizeUIButtonID(value string) string {
	value = strings.ToLower(strings.TrimSpace(value))
	var b strings.Builder
	lastDash := false
	for _, r := range value {
		valid := (r >= 'a' && r <= 'z') || (r >= '0' && r <= '9')
		if valid {
			b.WriteRune(r)
			lastDash = false
			continue
		}
		if !lastDash && b.Len() > 0 {
			b.WriteByte('-')
			lastDash = true
		}
	}
	return strings.Trim(b.String(), "-")
}

func (a *app) newCommand() *cobra.Command {
	var withSkill bool
	var skillAgent string
	var noUpdateCheck bool

	cmd := &cobra.Command{
		Use:   "new [workbook]",
		Short: "Create a new xlflow project and macro workbook or add-in",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			if err := a.writeScaffoldWelcome("new", noUpdateCheck); err != nil {
				return output.WithExitCode(output.ExitEnvironment, err)
			}
			commandOpts := buildCommandOptions(a.stderrWriter())
			var skillOpts agentskill.InstallOptions
			if withSkill {
				opts, err := a.resolveSkillInstallOptions(skillAgent, "", false)
				if err != nil {
					return a.writeFailure("new", output.ExitConfig, "skill_agent_required", err)
				}
				skillOpts = opts
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			{
				runOpts := commandOpts
				var excelEnv output.Envelope
				var excelCode int
				result, err := project.New(a.cwd, workbook, func(path string) error {
					env, code, err := a.runExcelWithProgress("Creating workbook", runOpts, func() (output.Envelope, int, error) {
						return a.excelRunner().New(path, runOpts)
					})
					excelEnv = env
					excelCode = code
					if err != nil {
						return err
					}
					if code != output.ExitSuccess {
						if env.Error != nil {
							return errors.New(env.Error.Message)
						}
						return errors.New("workbook creation failed")
					}
					return nil
				})
				if err != nil {
					if excelCode != 0 {
						return a.write(excelEnv, excelCode)
					}
					return a.writeFailure("new", output.ExitConfig, "new_failed", err)
				}
				bootstrapEnv, bootstrapCode, bootstrapErr := a.bootstrapScaffoldPush(runOpts)
				if bootstrapErr != nil {
					return bootstrapErr
				}
				if bootstrapCode != output.ExitSuccess {
					return a.write(bootstrapEnv, bootstrapCode)
				}
				var skillResult agentskill.InstallResult
				if withSkill {
					skillResult, err = agentskill.Install(skillOpts)
					if err != nil {
						return a.writeFailure("new", output.ExitConfig, "skill_install_failed", err)
					}
				}
				env := output.New("new")
				env.Workbook = result.Workbook
				env.Logs = []string{
					"created " + result.ConfigPath,
					"created " + result.Workbook,
					"pushed scaffolded VBA source to workbook",
				}
				if withSkill {
					env.Logs = append(env.Logs, "installed xlflow skill to "+skillResult.Path)
				}
				a.attachTypeDBBootstrap(&env)
				return a.write(env, output.ExitSuccess)
			}
		},
	}
	cmd.Flags().BoolVar(&withSkill, "with-skill", false, "install the bundled xlflow AI agent skill")
	cmd.Flags().StringVar(&skillAgent, "agent", "", "skill provider target: agents, codex, claude, cursor, or gemini")
	cmd.Flags().BoolVar(&noUpdateCheck, "no-update-check", false, "skip the interactive GitHub release update check during project scaffolding")
	return cmd
}

func (a *app) initCommand() *cobra.Command {
	var withSkill bool
	var withModule bool
	var skillAgent string
	var noUpdateCheck bool
	var userFormCodeSource string

	cmd := &cobra.Command{
		Use:   "init <workbook>",
		Short: "Create an xlflow project from an existing macro workbook or add-in",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			if err := a.writeScaffoldWelcome("init", noUpdateCheck); err != nil {
				return output.WithExitCode(output.ExitEnvironment, err)
			}
			commandOpts := buildCommandOptions(a.stderrWriter())
			var skillOpts agentskill.InstallOptions
			if withSkill {
				opts, err := a.resolveSkillInstallOptions(skillAgent, "", false)
				if err != nil {
					return a.writeFailure("init", output.ExitConfig, "skill_agent_required", err)
				}
				skillOpts = opts
			}
			{
				runOpts := commandOpts
				codeSource := strings.TrimSpace(userFormCodeSource)
				if codeSource == "" {
					codeSource = "frm"
				}
				result, err := project.InitWithOptions(a.cwd, args[0], project.InitOptions{UserFormCodeSource: codeSource})
				if err != nil {
					return a.writeFailure("init", output.ExitConfig, "init_failed", err)
				}
				bootstrapEnv, bootstrapCode, bootstrapErr := a.bootstrapScaffoldPull(runOpts)
				if bootstrapErr != nil {
					return bootstrapErr
				}
				if bootstrapCode != output.ExitSuccess {
					return a.write(bootstrapEnv, bootstrapCode)
				}
				var sidecarSpecs []string
				if codeSource == "sidecar" {
					cfg, err := a.loadConfig("init")
					if err != nil {
						return err
					}
					sidecarSpecs, err = a.writeImportedUserFormSpecs(cfg, runOpts)
					if err != nil {
						return a.writeFailure("init", output.ExitEnvironment, "init_failed", err)
					}
				}
				var installedModules project.InstallModulesResult
				if withModule {
					cfg, err := a.loadConfig("init")
					if err != nil {
						return err
					}
					installedModules, err = project.InstallHelperModules(a.cwd, cfg.Src)
					if err != nil {
						return a.writeFailure("init", output.ExitConfig, "module_install_failed", err)
					}
					pushEnv, pushCode, pushErr := a.pushSource("init", cfg, excel.PushOptions{
						BackupMode: "always",
						Keepalive:  runOpts,
						SourceRoot: project.ResolveModuleRoot(a.cwd, cfg.Src),
					}, "Importing bundled helper modules")
					if pushErr != nil {
						return pushErr
					}
					if pushCode != output.ExitSuccess {
						return a.write(pushEnv, pushCode)
					}
				}
				var skillResult agentskill.InstallResult
				if withSkill {
					skillResult, err = agentskill.Install(skillOpts)
					if err != nil {
						return a.writeFailure("init", output.ExitConfig, "skill_install_failed", err)
					}
				}
				env := output.New("init")
				env.Workbook = result.Workbook
				env.Logs = []string{
					"created " + result.ConfigPath,
					"copied workbook to " + result.Workbook,
					"pulled workbook VBA into source",
				}
				if len(sidecarSpecs) > 0 {
					env.Source = map[string]any{"created": sidecarSpecs}
					env.Logs = append(env.Logs, fmt.Sprintf("wrote %d UserForm Designer spec(s)", len(sidecarSpecs)))
				}
				if withModule {
					env.Source = mergeCreatedSource(env.Source, installedModules.Created)
					env.Logs = append(env.Logs,
						fmt.Sprintf("installed %d bundled helper module(s) into source", len(installedModules.Created)),
						"pushed bundled helper modules to workbook",
					)
				}
				if withSkill {
					env.Logs = append(env.Logs, "installed xlflow skill to "+skillResult.Path)
				}
				a.attachTypeDBBootstrap(&env)
				return a.write(env, output.ExitSuccess)
			}
		},
	}
	cmd.Flags().BoolVar(&withSkill, "with-skill", false, "install the bundled xlflow AI agent skill")
	cmd.Flags().BoolVar(&withModule, "with-module", false, "install bundled xlflow helper modules and push them to the workbook")
	cmd.Flags().StringVar(&skillAgent, "agent", "", "skill provider target: agents, codex, claude, cursor, or gemini")
	cmd.Flags().BoolVar(&noUpdateCheck, "no-update-check", false, "skip the interactive GitHub release update check during project scaffolding")
	cmd.Flags().StringVar(&userFormCodeSource, "userform-code-source", "frm", "UserForm code source for imported projects: frm or sidecar")
	return cmd
}

func (a *app) bootstrapScaffoldPush(keepaliveOpts excel.CommandOptions) (output.Envelope, int, error) {
	cfg, err := config.Load(a.cwd)
	if err != nil {
		return output.Envelope{}, 0, a.writeFailure("new", output.ExitConfig, "config_error", err)
	}
	return a.runExcelWithProgress("Importing scaffolded VBA source", keepaliveOpts, func() (output.Envelope, int, error) {
		return a.excelRunnerForConfig(cfg).PushWithOptions(cfg, excel.PushOptions{
			BackupMode: "never",
			Keepalive:  keepaliveOpts,
		})
	})
}

func (a *app) bootstrapScaffoldPull(keepaliveOpts excel.CommandOptions) (output.Envelope, int, error) {
	cfg, err := config.Load(a.cwd)
	if err != nil {
		return output.Envelope{}, 0, a.writeFailure("init", output.ExitConfig, "config_error", err)
	}
	return a.runExcelWithProgress("Exporting workbook VBA source", keepaliveOpts, func() (output.Envelope, int, error) {
		return a.excelRunnerForConfig(cfg).PullWithOptions(cfg, excel.SessionCommandOptions{
			Keepalive: keepaliveOpts,
		})
	})
}

func (a *app) writeImportedUserFormSpecs(cfg config.Config, commandOpts excel.CommandOptions) ([]string, error) {
	formsDir := workbookArgPath(a.cwd, cfg.Src.Forms)
	candidates, err := collectUserFormMigrationCandidates(a.cwd, formsDir, "")
	if err != nil {
		return nil, err
	}
	var created []string
	for _, item := range candidates {
		if _, err := os.Stat(item.SpecPath); err == nil {
			return nil, fmt.Errorf("refusing to overwrite existing Designer spec %s", item.SpecPath)
		} else if !errors.Is(err, os.ErrNotExist) {
			return nil, err
		}
		spec, err := a.inspectUserFormSpecForMigration(cfg, item.Name, false, commandOpts)
		if err != nil {
			return nil, err
		}
		output := forms.SnapshotOutput{Path: item.SpecPath, DisplayPath: displayPath(a.cwd, item.SpecPath), Format: "yaml"}
		if err := forms.WriteSnapshot(output, spec); err != nil {
			return nil, err
		}
		created = append(created, displayPath(a.cwd, item.SpecPath))
	}
	return created, nil
}

func mergeCreatedSource(source any, created []string) map[string]any {
	out := map[string]any{}
	for key, value := range cliObjectMap(source) {
		out[key] = value
	}
	existing := stringSliceForCLI(out["created"])
	existing = append(existing, created...)
	out["created"] = existing
	return out
}

func stringSliceForCLI(value any) []string {
	if value == nil {
		return nil
	}
	switch v := value.(type) {
	case []string:
		return append([]string{}, v...)
	case []any:
		out := make([]string, 0, len(v))
		for _, item := range v {
			if s, ok := item.(string); ok {
				out = append(out, s)
			}
		}
		return out
	default:
		return nil
	}
}

func (a *app) attachTypeDBBootstrap(env *output.Envelope) {
	if env == nil {
		return
	}
	env.Logs = append(env.Logs, "Type database: built-in DB ok")
	status, err := typedb.StatusFor(typedb.Options{GeneratorVersion: a.buildInfo.withDefaults().Version})
	if err != nil {
		appendTypeDBBootstrapWarning(env, "type_db_status_failed", "Generated TypeLib DB status could not be inspected: "+err.Error())
		return
	}
	if status.ManifestExists && !status.Stale {
		env.Logs = append(env.Logs, "Type database: generated TypeLib DB already exists")
		env.TypeDB = status
		return
	}
	if status.ManifestExists && status.Stale {
		reason := status.Reason
		if reason == "" {
			reason = "unknown reason"
		}
		appendTypeDBBootstrapWarning(env, "type_db_stale", "Generated TypeLib DB is stale: "+reason)
	}
	resolvedDir, err := typedb.ResolveDir("")
	if err != nil {
		appendTypeDBBootstrapWarning(env, "type_db_dir_failed", "Generated TypeLib DB directory could not be resolved: "+err.Error())
		return
	}
	typeDBEnv, code, err := a.excelRunner().TypeDBImport(excel.TypeDBImportOptions{
		OutputDir:        resolvedDir,
		GeneratorVersion: a.buildInfo.withDefaults().Version,
		Libraries:        []string{"excel"},
		Keepalive:        buildCommandOptions(a.stderrWriter()),
	})
	if err != nil {
		appendTypeDBBootstrapWarning(env, "type_db_init_skipped", "Generated TypeLib DB was skipped: "+err.Error())
		appendTypeDBBootstrapHint(env)
		return
	}
	if code != output.ExitSuccess {
		reason := "unknown error"
		if typeDBEnv.Error != nil && typeDBEnv.Error.Message != "" {
			reason = typeDBEnv.Error.Message
		}
		appendTypeDBBootstrapWarning(env, "type_db_init_skipped", "Generated TypeLib DB was skipped: "+reason)
		appendTypeDBBootstrapHint(env)
		return
	}
	env.Logs = append(env.Logs, "Type database: generated TypeLib DB created at "+resolvedDir)
	env.TypeDB = typeDBEnv.TypeDB
}

func appendTypeDBBootstrapWarning(env *output.Envelope, code string, message string) {
	warnings := anySlice(env.Warnings)
	warnings = append(warnings, map[string]any{"code": code, "message": message})
	env.Warnings = warnings
}

func appendTypeDBBootstrapHint(env *output.Envelope) {
	appendEnvelopeHint(env, "type_db_init_later", "Run `xlflow type db init` after installing Excel to enable richer COM completions.")
}

func appendEnvelopeHint(env *output.Envelope, code string, message string) {
	hints := anySlice(env.Hints)
	hints = append(hints, map[string]any{"code": code, "message": message})
	env.Hints = hints
}

func (a *app) doctorCommand() *cobra.Command {
	var checkWorkbook bool
	cmd := &cobra.Command{
		Use:   "doctor",
		Short: "Diagnose Excel COM and VBIDE access",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			doctorConfig, err := a.loadDoctorConfig(checkWorkbook)
			if err != nil {
				return err
			}
			cfg := doctorConfig.Config
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Checking Excel automation", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).DoctorWithOptions(cfg, excel.DoctorOptions{
					CheckWorkbook: doctorConfig.CheckWorkbook,
					Keepalive:     commandOpts,
				})
				return runErr
			})
			if err != nil {
				return err
			}
			attachDoctorProjectConfigDiagnostics(&env, doctorConfig)
			appendDoctorConfigMessages(&env, doctorConfig)
			if doctorConfig.Found {
				boundaries, analyzeErr := gui.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
				if analyzeErr == nil && len(boundaries) > 0 {
					env.GUIBoundaries = boundaries
					env.Diagnostics = withGUIBoundarySummary(env.Diagnostics, boundaries)
					env.Logs = append(env.Logs, fmt.Sprintf("detected %d GUI boundary candidate(s) in source", len(boundaries)))
				}
			}
			a.attachTypeDBDoctorStatus(&env)
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&checkWorkbook, "workbook", false, "open the configured workbook as part of doctor diagnostics")
	return cmd
}

type doctorConfigLoadResult struct {
	Config        config.Config
	Found         bool
	Path          string
	CheckWorkbook bool
	Warnings      []map[string]any
	Hints         []map[string]any
}

func (a *app) loadDoctorConfig(requestedCheckWorkbook bool) (doctorConfigLoadResult, error) {
	path := filepath.Join(a.cwd, config.FileName)
	cfg, err := config.Load(a.cwd)
	if err != nil && errors.Is(err, config.ErrInvalidExcelBridge) && a.hasValidBridgeOverride() {
		cfg, err = config.LoadAllowInvalidExcelBridge(a.cwd)
	}
	if err == nil {
		if len(cfg.Warnings) > 0 {
			a.configWarnings = append(a.configWarnings, cfg.Warnings...)
		}
		return doctorConfigLoadResult{
			Config:        cfg,
			Found:         true,
			Path:          path,
			CheckWorkbook: requestedCheckWorkbook,
		}, nil
	}
	if !errors.Is(err, config.ErrConfigNotFound) {
		return doctorConfigLoadResult{}, a.writeFailure("doctor", output.ExitConfig, "config_error", err)
	}

	result := doctorConfigLoadResult{
		Config:        config.Default(),
		Found:         false,
		Path:          path,
		CheckWorkbook: false,
		Warnings: []map[string]any{
			{
				"code":    "project_config_missing",
				"message": fmt.Sprintf("%s was not found; running project-independent diagnostics only.", config.FileName),
			},
		},
		Hints: []map[string]any{
			{
				"code":    "project_create",
				"message": "Run `xlflow new` to create a new xlflow project.",
			},
			{
				"code":    "project_init",
				"message": "Run `xlflow init <workbook>` to convert an existing workbook into an xlflow project.",
			},
		},
	}
	if requestedCheckWorkbook {
		result.Warnings = append(result.Warnings, map[string]any{
			"code":    "doctor_workbook_skipped",
			"message": "`xlflow doctor --workbook` was requested, but no configured workbook is available without xlflow.toml.",
		})
		result.Hints = append(result.Hints, map[string]any{
			"code":    "doctor_workbook_requires_project",
			"message": "Create or initialize an xlflow project before using `xlflow doctor --workbook`.",
		})
	}
	return result, nil
}

func attachDoctorProjectConfigDiagnostics(env *output.Envelope, result doctorConfigLoadResult) {
	if env == nil {
		return
	}
	diag := map[string]any{}
	for key, item := range cliObjectMap(env.Diagnostics) {
		diag[key] = item
	}
	diag["project_config"] = map[string]any{
		"found": result.Found,
		"path":  result.Path,
	}
	env.Diagnostics = diag
}

func appendDoctorConfigMessages(env *output.Envelope, result doctorConfigLoadResult) {
	if env == nil {
		return
	}
	if len(result.Warnings) > 0 {
		warnings := anySlice(env.Warnings)
		for _, warning := range result.Warnings {
			warnings = append(warnings, warning)
		}
		env.Warnings = warnings
	}
	if len(result.Hints) > 0 {
		hints := anySlice(env.Hints)
		for _, hint := range result.Hints {
			hints = append(hints, hint)
		}
		env.Hints = hints
	}
}

const vbaObjectModelAccessWarningCode = "vba_object_model_access_disabled"

func appendVBAObjectModelAccessMessages(env *output.Envelope) {
	if env == nil || !vbaObjectModelAccessLooksDisabled(env) {
		return
	}
	appendUniqueMessage(&env.Warnings, vbaObjectModelAccessWarningCode, "Excel VBA object model access is disabled or unavailable. xlflow needs this setting to import, export, inspect, and run VBA components.")
	appendUniqueMessage(&env.Hints, "enable_vba_object_model_access", "In Excel, open Trust Center -> Macro Settings and enable \"Trust access to the VBA project object model\", then rerun the xlflow command.")
}

func vbaObjectModelAccessLooksDisabled(env *output.Envelope) bool {
	diag := cliObjectMap(env.Diagnostics)
	excelDiag := cliObjectMap(diag["excel"])
	if value, ok := boolValueInMap(excelDiag, "trust_vba_access"); ok && !value {
		return true
	}
	if value, ok := boolValueInMap(excelDiag, "vbide_access"); ok && !value {
		return true
	}
	if value, ok := boolValueInMap(excelDiag, "vbproject_access"); ok && !value {
		return true
	}
	if env.Error == nil {
		return false
	}
	code := strings.ToLower(strings.TrimSpace(env.Error.Code))
	if strings.Contains(code, "vbproject_access_denied") || strings.Contains(code, "vbide") {
		return true
	}
	message := strings.ToLower(env.Error.Message)
	return strings.Contains(message, "vbproject access is denied") ||
		strings.Contains(message, "vbide access") ||
		strings.Contains(message, "trust access to the vba project object model") ||
		strings.Contains(message, "get_vbproject failed") ||
		strings.Contains(message, "import_vba_components failed") ||
		strings.Contains(env.Error.Message, "プログラミングによる Visual Basic プロジェクトへのアクセス")
}

func boolValueInMap(m map[string]any, key string) (bool, bool) {
	value, ok := m[key]
	if !ok || value == nil {
		return false, false
	}
	switch v := value.(type) {
	case bool:
		return v, true
	case string:
		if strings.EqualFold(v, "true") {
			return true, true
		}
		if strings.EqualFold(v, "false") {
			return false, true
		}
	}
	return false, false
}

func appendUniqueMessage(target *any, code string, message string) {
	items := anySlice(*target)
	for _, item := range items {
		if cliObjectMap(item)["code"] == code {
			*target = items
			return
		}
	}
	items = append(items, map[string]any{"code": code, "message": message})
	*target = items
}

func (a *app) attachTypeDBDoctorStatus(env *output.Envelope) {
	if env == nil {
		return
	}
	status, err := typedb.StatusFor(typedb.Options{GeneratorVersion: a.buildInfo.withDefaults().Version})
	if err != nil {
		appendTypeDBBootstrapWarning(env, "type_db_status_failed", "Generated TypeLib DB status could not be inspected: "+err.Error())
		return
	}
	env.TypeDB = status
	switch {
	case !status.ManifestExists:
		appendTypeDBBootstrapWarning(env, "type_db_missing", "Generated TypeLib DB has not been initialized.")
		appendEnvelopeHint(env, "type_db_init", "Run `xlflow type db init` or `xlflow type db refresh --library all` to enable richer COM completions.")
	case status.Stale:
		reason := status.Reason
		if reason == "" {
			reason = "unknown reason"
		}
		appendTypeDBBootstrapWarning(env, "type_db_stale", "Generated TypeLib DB is stale: "+reason)
		appendEnvelopeHint(env, "type_db_refresh", "Run `xlflow type db refresh --library all` to regenerate the TypeLib database.")
	}
}

func (a *app) attachCommand() *cobra.Command {
	var active bool
	cmd := &cobra.Command{
		Use:   "attach --active",
		Short: "Inspect the active Excel workbook connection",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("attach")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Inspecting active workbook", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).Attach(cfg, active, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			env.Warnings = append(anySlice(env.Warnings), map[string]any{
				"code":    "attach_active_deprecated",
				"message": "`xlflow attach --active` is deprecated and only validates the active workbook. Use `xlflow session attach` when you want xlflow commands to operate on an already-open workbook.",
			})
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&active, "active", false, "attach to the active Excel workbook")
	return cmd
}

func (a *app) pullCommand() *cobra.Command {
	var session bool
	var withFormulas bool
	cmd := &cobra.Command{
		Use:   "pull",
		Short: "Export VBA components from the configured workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("pull")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Exporting VBA source", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).PullWithOptions(cfg, excel.SessionCommandOptions{Session: session, Keepalive: commandOpts})
				return runErr
			})
			if err != nil {
				return err
			}
			if code == output.ExitSuccess && withFormulas {
				formulaResult, formulaErr := formulaspkg.Pull(workbookArgPath(a.cwd, cfg.Excel.Path), filepath.Join(a.cwd, "formulas"))
				if formulaErr != nil {
					attachFormulaPullError(&env, formulaErr)
				}
				if formulaErr == nil {
					attachFormulaPullResult(&env, formulaResult, a.cwd)
					env.Logs = append(env.Logs, fmt.Sprintf("extracted %d formula region(s) from %d sheet(s)", formulaResult.FormulaRegionCount, len(formulaResult.Manifest.Sheets)))
				}
				if session && formulaErr == nil {
					env.Warnings = append(anySlice(env.Warnings), map[string]any{
						"code":    "formula_snapshot_saved_file",
						"message": "Formula snapshots were extracted from the saved workbook file. If the live session workbook has unsaved formula changes, run `xlflow save --json` and `xlflow formulas pull --json` again.",
					})
				}
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	cmd.Flags().BoolVar(&withFormulas, "formulas", false, "also extract worksheet formula snapshots into formulas/")
	return cmd
}

func attachFormulaPullResult(env *output.Envelope, result formulaspkg.Result, root string) {
	if env == nil {
		return
	}
	outputPayload := cliObjectMap(env.Output)
	outputPayload["formulas"] = formulaResultPayload(result, root)
	env.Output = outputPayload
}

func formulaResultPayload(result formulaspkg.Result, root string) map[string]any {
	return map[string]any{
		"dir":                  displayPath(root, result.OutputDir),
		"manifest":             displayPath(root, result.ManifestPath),
		"sheet_count":          len(result.Manifest.Sheets),
		"formula_region_count": result.FormulaRegionCount,
		"parse_status_summary": result.Manifest.ParseStatusSummary,
		"defined_name_count":   len(result.Names),
	}
}

func attachFormulaPullError(env *output.Envelope, err error) {
	if env == nil || err == nil {
		return
	}
	outputPayload := cliObjectMap(env.Output)
	outputPayload["formulas_error"] = map[string]any{
		"code":    "pull_formulas_failed",
		"message": err.Error(),
	}
	env.Output = outputPayload
	env.Warnings = append(anySlice(env.Warnings), map[string]any{
		"code":    "pull_formulas_failed",
		"message": "VBA source was pulled, but formula snapshot extraction failed: " + err.Error(),
	})
}

func (a *app) packCommand() *cobra.Command {
	var outPath string
	var templatePath string
	var experimental bool
	cmd := &cobra.Command{
		Use:   "pack --out <path.xlsm> --experimental",
		Short: "Build an .xlsm artifact from source and a workbook template",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			if !experimental {
				return a.writeFailure("pack", output.ExitConfig, "pack_experimental_required", errors.New("pack requires --experimental"))
			}
			if strings.TrimSpace(outPath) == "" || !strings.EqualFold(filepath.Ext(outPath), ".xlsm") {
				return a.writeFailure("pack", output.ExitConfig, "pack_args_invalid", errors.New("--out is required and must end in .xlsm"))
			}

			cfg, err := a.loadConfig("pack")
			if err != nil {
				return err
			}
			configuredWorkbook := workbookArgPath(a.cwd, cfg.Excel.Path)
			if err := workbookformat.ValidatePackTemplate(configuredWorkbook); err != nil {
				var unsupported workbookformat.UnsupportedError
				if errors.As(err, &unsupported) {
					return a.writeUnsupportedWorkbookFormat("pack", unsupported)
				}
				return a.writeFailure("pack", output.ExitConfig, "pack_args_invalid", err)
			}
			resolvedTemplate := strings.TrimSpace(templatePath)
			if resolvedTemplate == "" {
				resolvedTemplate = configuredWorkbook
			} else {
				resolvedTemplate = workbookArgPath(a.cwd, resolvedTemplate)
			}
			if strings.TrimSpace(resolvedTemplate) == "" {
				return a.writeFailure("pack", output.ExitConfig, "pack_template_not_found", errors.New("template workbook path is empty"))
			}
			if _, err := os.Stat(resolvedTemplate); err != nil {
				return a.writeFailure("pack", output.ExitConfig, "pack_template_not_found", err)
			}
			if err := workbookformat.ValidatePackTemplate(resolvedTemplate); err != nil {
				var unsupported workbookformat.UnsupportedError
				if errors.As(err, &unsupported) {
					return a.writeUnsupportedWorkbookFormat("pack", unsupported)
				}
				return a.writeFailure("pack", output.ExitConfig, "pack_args_invalid", err)
			}

			resolvedOut := workbookArgPath(a.cwd, outPath)
			if samePath(resolvedOut, resolvedTemplate) || samePath(resolvedOut, configuredWorkbook) {
				return a.writeFailure("pack", output.ExitConfig, "pack_in_place_overwrite", fmt.Errorf("--out must differ from template and configured workbook: %s", resolvedOut))
			}
			candidates := uniqueNonEmptyPaths(resolvedTemplate, configuredWorkbook, resolvedOut)
			for _, candidate := range candidates {
				if lockPath, locked := officeLockFilePresent(candidate); locked {
					return a.writeFailure("pack", output.ExitConfig, "pack_active_session", fmt.Errorf("workbook appears to be open: %s", lockPath))
				}
			}
			if wb, active, err := a.packActiveSession(candidates); err != nil {
				return a.writeFailure("pack", output.ExitConfig, "pack_active_session", fmt.Errorf("could not check for an active session: %w", err))
			} else if active {
				return a.writeFailure("pack", output.ExitConfig, "pack_active_session", fmt.Errorf("an xlflow session is active for %s", wb))
			}

			sources, err := collectPackSourceModules(a.cwd, cfg)
			if err != nil {
				if errors.Is(err, packpkg.ErrAmbiguousLayout) {
					return a.writePackEngineFailure(err)
				}
				return a.writeFailure("pack", output.ExitEnvironment, "pack_source_read_failed", err)
			}
			templateBytes, err := os.ReadFile(resolvedTemplate)
			if err != nil {
				return a.writeFailure("pack", output.ExitConfig, "pack_template_not_found", err)
			}
			workbookBytes, meta, err := packpkg.BuildWorkbook(templateBytes, sources)
			if err != nil {
				return a.writePackEngineFailure(err)
			}
			createdParentDirs, err := writePackOutput(resolvedOut, workbookBytes)
			if err != nil {
				return a.writeFailure("pack", output.ExitEnvironment, "pack_write_failed", err)
			}

			env := output.New("pack")
			outputPayload := map[string]any{
				"path":   displayPath(a.cwd, resolvedOut),
				"format": "xlsm",
			}
			if createdParentDirs {
				outputPayload["created_parent_dirs"] = true
			}
			env.Output = outputPayload
			env.Pack = map[string]any{
				"backend":        "pure-go",
				"experimental":   true,
				"vbe_validation": "not_performed",
				"template":       displayPath(a.cwd, resolvedTemplate),
				"modules": map[string]any{
					"standard":        meta.Standard,
					"class":           meta.Class,
					"document":        meta.Document,
					"form":            meta.Form,
					"carried_streams": meta.CarriedStreams,
				},
			}
			env.Warnings = []map[string]any{{
				"code":    "vbe_validation_skipped",
				"message": "pack did not open Excel; no VBE compile or runtime validation was performed.",
			}}
			env.Logs = []string{"packed " + displayPath(a.cwd, resolvedOut)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&outPath, "out", "", "destination .xlsm artifact path")
	cmd.Flags().StringVar(&templatePath, "template", "", "workbook template path")
	cmd.Flags().BoolVar(&experimental, "experimental", false, "enable experimental pure-Go pack")
	return cmd
}

func (a *app) writePackEngineFailure(err error) error {
	switch {
	case errors.Is(err, packpkg.ErrProtectedProject):
		return a.writeFailure("pack", output.ExitValidation, "pack_protected_project", err)
	case errors.Is(err, packpkg.ErrSignedProject):
		return a.writeFailure("pack", output.ExitValidation, "pack_signed_project", err)
	case errors.Is(err, packpkg.ErrUserFormGenerationUnsupported):
		return a.writeFailure("pack", output.ExitValidation, "pack_userform_generation_unsupported", err)
	case errors.Is(err, packpkg.ErrAmbiguousLayout):
		return a.writeFailure("pack", output.ExitValidation, "pack_ambiguous_layout", err)
	default:
		return a.writeFailure("pack", output.ExitEnvironment, "pack_failed", err)
	}
}

func (a *app) pushCommand() *cobra.Command {
	var backupMode string
	var fast bool
	var changedOnly bool
	var session bool
	var noSave bool

	cmd := &cobra.Command{
		Use:   "push",
		Short: "Import source VBA components into the configured workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("push")
			if err != nil {
				return err
			}
			pushOpts, err := buildPushOptions(backupMode, fast, changedOnly, session, noSave, commandOpts)
			if err != nil {
				return a.writeFailure("push", output.ExitConfig, "push_args_invalid", err)
			}
			env, code, err := a.pushSource("push", cfg, pushOpts, "Importing VBA source")
			if err != nil {
				return err
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			a.applyAutomaticBackupRetention(cfg, workbookPath, &env, code == output.ExitSuccess && env.Backup != nil)
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&backupMode, "backup", "always", "backup policy: always or never")
	cmd.Flags().BoolVar(&fast, "fast", false, "use development-oriented fast push defaults")
	cmd.Flags().BoolVar(&changedOnly, "changed-only", false, "skip workbook updates when source state has not changed")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	cmd.Flags().BoolVar(&noSave, "no-save", false, "do not save after session push")
	return cmd
}

func (a *app) generateCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "generate",
		Short: "Generate project artifacts",
	}
	cmd.AddCommand(a.generateTestCommand())
	return cmd
}

func (a *app) generateTestCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "test <module-name>",
		Short: "Generate a new VBA test module",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("generate test")
			if err != nil {
				return err
			}
			result, err := project.GenerateTestModule(a.cwd, args[0], cfg.Src)
			if err != nil {
				return a.writeFailure("generate test", output.ExitConfig, "generate_test_failed", err)
			}
			env := output.New("generate test")
			env.Source = map[string]any{"path": result.Path, "created": result.Created}
			env.Logs = []string{fmt.Sprintf("created test module: %s", result.Path)}
			return a.write(env, output.ExitSuccess)
		},
	}
	return cmd
}

func (a *app) moduleCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "module",
		Short: "Manage VBA source modules",
	}
	cmd.AddCommand(a.moduleNewCommand(), a.moduleRemoveCommand(), a.moduleRenameCommand(), a.moduleInstallCommand())
	return cmd
}

func (a *app) moduleNewCommand() *cobra.Command {
	var moduleType string
	cmd := &cobra.Command{
		Use:   "new <name>",
		Short: "Create a new standard or class module source file",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			if strings.TrimSpace(moduleType) == "" {
				return a.writeFailure("module new", output.ExitConfig, "module_new_args_invalid", errors.New("--type is required"))
			}
			cfg, err := a.loadConfig("module new")
			if err != nil {
				return err
			}
			result, err := project.NewModule(a.cwd, args[0], moduleType, cfg.Src)
			if err != nil {
				if errors.Is(err, project.ErrInvalidComponentName) || errors.Is(err, project.ErrInvalidModuleType) {
					return a.writeFailure("module new", output.ExitConfig, "module_new_args_invalid", err)
				}
				if errors.Is(err, project.ErrScaffoldExists) {
					return a.writeFailure("module new", output.ExitValidation, "module_new_failed", err)
				}
				return a.writeFailure("module new", output.ExitEnvironment, "module_new_failed", err)
			}
			env := output.New("module new")
			env.Source = map[string]any{
				"created": result.Created,
				"kind":    result.Kind,
				"name":    result.Name,
				"path":    result.Path,
			}
			env.Logs = []string{fmt.Sprintf("created %s module source: %s", result.Kind, result.Path)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&moduleType, "type", "", "module type: standard or class")
	return cmd
}

func (a *app) moduleRemoveCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "remove <module-name>",
		Short: "Remove a source module from the project",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("module remove")
			if err != nil {
				return err
			}
			result, err := project.RemoveModule(a.cwd, args[0], cfg.Src)
			if err != nil {
				return a.writeModuleMutationFailure("module remove", err)
			}
			env := output.New("module remove")
			env.Source = map[string]any{
				"operation":     result.Operation,
				"module":        result.Module,
				"kind":          result.Kind,
				"removed":       result.Removed,
				"requires_push": result.RequiresPush,
			}
			env.Logs = []string{
				fmt.Sprintf("Removed module %q.", result.Module),
				`Run "xlflow push" to apply the change to the workbook.`,
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	return cmd
}

func (a *app) moduleRenameCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "rename <old-name> <new-name>",
		Short: "Rename a source module in the project",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("module rename")
			if err != nil {
				return err
			}
			result, err := project.RenameModule(a.cwd, args[0], args[1], cfg.Src)
			if err != nil {
				return a.writeModuleMutationFailure("module rename", err)
			}
			env := output.New("module rename")
			env.Source = map[string]any{
				"operation":     result.Operation,
				"old_name":      result.OldName,
				"new_name":      result.NewName,
				"kind":          result.Kind,
				"renamed":       result.Renamed,
				"requires_push": result.RequiresPush,
			}
			env.Logs = []string{
				fmt.Sprintf("Renamed module %q to %q.", result.OldName, result.NewName),
				`Run "xlflow push" to apply the change to the workbook.`,
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	return cmd
}

func (a *app) writeModuleMutationFailure(command string, err error) error {
	switch {
	case errors.Is(err, project.ErrInvalidComponentName):
		return a.writeFailure(command, output.ExitValidation, "module_name_invalid", err)
	case errors.Is(err, project.ErrProtectedModule):
		return a.writeFailure(command, output.ExitValidation, "protected_module", err)
	case errors.Is(err, project.ErrModuleNotFound):
		return a.writeFailure(command, output.ExitValidation, "module_not_found", err)
	case errors.Is(err, project.ErrModuleAlreadyExists):
		return a.writeFailure(command, output.ExitValidation, "module_already_exists", err)
	case errors.Is(err, project.ErrModuleAmbiguous):
		return a.writeFailure(command, output.ExitValidation, "module_ambiguous", err)
	default:
		return a.writeFailure(command, output.ExitEnvironment, "module_mutation_failed", err)
	}
}

func (a *app) moduleInstallCommand() *cobra.Command {
	var push bool
	cmd := &cobra.Command{
		Use:   "install",
		Short: "Install bundled xlflow helper modules into the configured source tree",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("module install")
			if err != nil {
				return err
			}
			result, err := project.InstallHelperModules(a.cwd, cfg.Src)
			if err != nil {
				return a.writeFailure("module install", output.ExitConfig, "module_install_failed", err)
			}
			env := output.New("module install")
			env.Source = map[string]any{"created": result.Created}
			env.Logs = []string{fmt.Sprintf("installed %d bundled helper module(s) into source", len(result.Created))}
			if !push {
				return a.write(env, output.ExitSuccess)
			}
			commandOpts := buildCommandOptions(a.stderrWriter())
			pushEnv, pushCode, pushErr := a.pushSource("module install", cfg, excel.PushOptions{
				BackupMode: "always",
				Keepalive:  commandOpts,
				SourceRoot: project.ResolveModuleRoot(a.cwd, cfg.Src),
			}, "Importing bundled helper modules")
			if pushErr != nil {
				return pushErr
			}
			if pushCode != output.ExitSuccess {
				return a.write(pushEnv, pushCode)
			}
			env.Workbook = cfg.Excel.Path
			env.Logs = append(env.Logs, "pushed bundled helper modules to workbook")
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&push, "push", false, "push the bundled helper modules to the configured workbook after installation")
	return cmd
}

func buildPushOptions(backupMode string, fast bool, changedOnly bool, session bool, noSave bool, keepalive excel.CommandOptions) (excel.PushOptions, error) {
	if backupMode == "" {
		backupMode = "always"
	}
	if backupMode != "always" && backupMode != "never" {
		return excel.PushOptions{}, fmt.Errorf("--backup must be always or never")
	}
	if noSave && !session {
		return excel.PushOptions{}, fmt.Errorf("--no-save requires --session")
	}
	if fast {
		backupMode = "never"
		changedOnly = true
	}
	return excel.PushOptions{
		BackupMode:  backupMode,
		Fast:        fast,
		ChangedOnly: changedOnly,
		Session:     session,
		NoSave:      noSave,
		Keepalive:   keepalive,
	}, nil
}

func (a *app) pushSource(command string, cfg config.Config, pushOpts excel.PushOptions, progressLabel string) (output.Envelope, int, error) {
	if err := a.runUserFormCodeSourcePreflight(command, cfg, nil); err != nil {
		return output.Envelope{}, 0, err
	}
	if err := a.runUserFormArtifactPreflight(command, cfg, nil); err != nil {
		return output.Envelope{}, 0, err
	}
	if err := a.runSourcePreflight(command, cfg, "pushing to Excel", nil, nil); err != nil {
		return output.Envelope{}, 0, err
	}
	var env output.Envelope
	var code int
	run := func() error {
		var runErr error
		env, code, runErr = a.excelRunnerForConfig(cfg).PushWithOptions(cfg, pushOpts)
		return runErr
	}
	err := a.withExcelProgress(progressLabel, pushOpts.Keepalive, run)
	if err != nil {
		return output.Envelope{}, 0, err
	}
	return env, code, nil
}

type rollbackTarget struct {
	Latest   bool
	BackupID string
}

func buildRollbackTarget(latest bool, backupID string) (rollbackTarget, error) {
	backupID = strings.TrimSpace(backupID)
	switch {
	case latest && backupID != "":
		return rollbackTarget{}, fmt.Errorf("--latest and --backup cannot be combined")
	case !latest && backupID == "":
		return rollbackTarget{}, fmt.Errorf("exactly one of --latest or --backup is required")
	}
	return rollbackTarget{Latest: latest, BackupID: backupID}, nil
}

func resolveRollbackRecord(rootDir, workbookPath string, target rollbackTarget) (backup.Record, error) {
	if target.Latest {
		return backup.Latest(rootDir, workbookPath)
	}
	return backup.Find(rootDir, workbookPath, target.BackupID)
}

func (a *app) applyAutomaticBackupRetention(cfg config.Config, workbookPath string, env *output.Envelope, trigger bool) {
	if env == nil || !trigger || !cfg.Backup.Retention.Enabled {
		return
	}
	opts := automaticBackupPruneOptions(cfg.Backup.Retention)
	result, err := automaticBackupPrune(a.cwd, workbookPath, opts)
	if shouldAttachAutomaticBackupPrune(result, err) {
		env.BackupPrune = renderAutomaticBackupPruneResult(a.cwd, result)
	}
	for _, warning := range automaticBackupSkippedWarnings(a.cwd, result) {
		env.Warnings = append(anySlice(env.Warnings), warning)
	}
	if err != nil {
		env.Warnings = append(anySlice(env.Warnings), map[string]any{
			"code":    backup.ErrPruneFailed,
			"message": "The workbook operation succeeded, but old backups could not be pruned.",
		})
		env.Logs = append(env.Logs, "workbook operation succeeded, but automatic backup pruning failed")
		return
	}
	if result.Deleted > 0 {
		env.Logs = append(env.Logs, fmt.Sprintf("pruned %d old backup(s), freed %s", result.Deleted, formatBackupPruneBytes(result.FreedBytes)))
	}
	if len(result.SkippedInvalid) > 0 || len(result.SkippedLegacy) > 0 {
		env.Logs = append(env.Logs, "automatic backup pruning skipped invalid or legacy backup entries")
	}
}

func automaticBackupPruneOptions(retention config.BackupRetentionConfig) backup.PruneOptions {
	opts := backup.PruneOptions{
		MinKeep:           retention.MinKeep,
		AllowNoConditions: true,
	}
	if retention.MaxCount > 0 {
		maxCount := retention.MaxCount
		opts.MaxCount = &maxCount
	}
	if retention.MaxAgeDays > 0 {
		opts.OlderThan = time.Duration(retention.MaxAgeDays) * 24 * time.Hour
		opts.OlderThanSet = true
	}
	if retention.MaxTotalSizeMB > 0 {
		opts.MaxTotalSize = int64(retention.MaxTotalSizeMB) * 1000 * 1000
		opts.MaxTotalSizeSet = true
	}
	return opts
}

func shouldAttachAutomaticBackupPrune(result backup.PruneResult, err error) bool {
	return err != nil || result.Deleted > 0 || result.Failed > 0 || len(result.SkippedInvalid) > 0 || len(result.SkippedLegacy) > 0
}

func automaticBackupSkippedWarnings(rootDir string, result backup.PruneResult) []map[string]any {
	warnings := make([]map[string]any, 0, len(result.SkippedInvalid)+len(result.SkippedLegacy))
	for _, entry := range result.SkippedInvalid {
		warnings = append(warnings, map[string]any{
			"code":    "invalid_backup_entry",
			"message": "An invalid backup entry was skipped during automatic pruning.",
			"path":    displayPath(rootDir, entry.Directory),
			"reason":  entry.Code,
			"detail":  entry.Message,
		})
	}
	for _, entry := range result.SkippedLegacy {
		warnings = append(warnings, map[string]any{
			"code":    "legacy_backup_entry",
			"message": "A legacy backup entry without metadata was skipped during automatic pruning.",
			"path":    displayPath(rootDir, entry.Directory),
		})
	}
	return warnings
}

func formatBackupPruneBytes(bytes int64) string {
	if bytes >= 1000*1000 {
		return fmt.Sprintf("%.1f MB", float64(bytes)/1000/1000)
	}
	return fmt.Sprintf("%d bytes", bytes)
}

func renderBackupRecords(rootDir string, records []backup.Record) []map[string]any {
	rendered := make([]map[string]any, 0, len(records))
	for _, record := range records {
		rendered = append(rendered, map[string]any{
			"id":         record.ID,
			"created_at": record.CreatedAt.Format(time.RFC3339),
			"reason":     record.Reason,
			"workbook":   displayPath(rootDir, record.OriginalWorkbookPath),
			"path":       displayPath(rootDir, record.BackupFileAbsPath),
			"size_bytes": record.SizeBytes,
			"status":     "valid",
		})
	}
	return rendered
}

func renderBackupPruneResult(rootDir string, result backup.PruneResult) map[string]any {
	return map[string]any{
		"dry_run":                 result.DryRun,
		"matched":                 result.Matched,
		"deleted":                 result.Deleted,
		"failed":                  result.Failed,
		"freed_bytes":             result.FreedBytes,
		"candidates":              renderBackupCandidates(rootDir, result.Candidates),
		"deleted_entries":         renderDeletedBackupEntries(rootDir, result.DeletedEntries),
		"failed_entries":          renderFailedBackupEntries(rootDir, result.FailedEntries),
		"skipped_invalid_entries": renderSkippedInvalidBackupEntries(rootDir, result.SkippedInvalid),
		"skipped_legacy_entries":  renderSkippedLegacyBackupEntries(rootDir, result.SkippedLegacy),
	}
}

func renderAutomaticBackupPruneResult(rootDir string, result backup.PruneResult) map[string]any {
	rendered := renderBackupPruneResult(rootDir, result)
	rendered["automatic"] = true
	return rendered
}

func renderBackupCandidates(rootDir string, entries []backup.CandidateEntry) []map[string]any {
	rendered := make([]map[string]any, 0, len(entries))
	for _, entry := range entries {
		item := map[string]any{
			"id":         entry.ID,
			"path":       displayPath(rootDir, entry.Directory),
			"size_bytes": entry.SizeBytes,
			"reasons":    entry.Reasons,
			"status":     entry.Status,
		}
		if !entry.CreatedAt.IsZero() {
			item["created_at"] = entry.CreatedAt.Format(time.RFC3339)
		}
		if entry.Reason != "" {
			item["reason"] = entry.Reason
		}
		if entry.Code != "" {
			item["code"] = entry.Code
		}
		if entry.Message != "" {
			item["message"] = entry.Message
		}
		rendered = append(rendered, item)
	}
	return rendered
}

func renderDeletedBackupEntries(rootDir string, entries []backup.DeletedEntry) []map[string]any {
	rendered := make([]map[string]any, 0, len(entries))
	for _, entry := range entries {
		rendered = append(rendered, map[string]any{
			"id":          entry.ID,
			"path":        displayPath(rootDir, entry.Directory),
			"freed_bytes": entry.FreedBytes,
		})
	}
	return rendered
}

func renderFailedBackupEntries(rootDir string, entries []backup.FailedEntry) []map[string]any {
	rendered := make([]map[string]any, 0, len(entries))
	for _, entry := range entries {
		rendered = append(rendered, map[string]any{
			"id":      entry.ID,
			"path":    displayPath(rootDir, entry.Directory),
			"code":    entry.Code,
			"message": entry.Message,
		})
	}
	return rendered
}

func renderSkippedInvalidBackupEntries(rootDir string, entries []backup.InvalidEntry) []map[string]any {
	rendered := make([]map[string]any, 0, len(entries))
	for _, entry := range entries {
		rendered = append(rendered, map[string]any{
			"path":    displayPath(rootDir, entry.Directory),
			"code":    entry.Code,
			"message": entry.Message,
			"status":  "invalid",
		})
	}
	return rendered
}

func renderSkippedLegacyBackupEntries(rootDir string, entries []backup.LegacyEntry) []map[string]any {
	rendered := make([]map[string]any, 0, len(entries))
	for _, entry := range entries {
		rendered = append(rendered, map[string]any{
			"path":   displayPath(rootDir, entry.Directory),
			"status": "legacy",
		})
	}
	return rendered
}

func backupErrorCode(err error) string {
	var backupErr *backup.Error
	if errors.As(err, &backupErr) && backupErr.Code != "" {
		return backupErr.Code
	}
	return backup.ErrDeleteFailed
}

func backupDeleteExitCode(err error) int {
	switch backupErrorCode(err) {
	case backup.ErrDeleteArgsInvalid:
		return output.ExitConfig
	case backup.ErrNotFound, backup.ErrDeleteScopeMismatch, backup.ErrDeleteUnsafePath:
		return output.ExitValidation
	default:
		return output.ExitEnvironment
	}
}

func renderInvalidBackupWarnings(rootDir string, entries []backup.InvalidEntry) []map[string]any {
	if len(entries) == 0 {
		return nil
	}
	warnings := make([]map[string]any, 0, len(entries))
	for _, entry := range entries {
		warnings = append(warnings, map[string]any{
			"code":    "invalid_backup_entry",
			"message": "A backup entry could not be read.",
			"path":    displayPath(rootDir, entry.Directory),
			"reason":  entry.Code,
			"detail":  entry.Message,
		})
	}
	return warnings
}

func displayPath(rootDir, path string) string {
	rel, err := filepath.Rel(rootDir, path)
	if err != nil {
		return filepath.ToSlash(path)
	}
	if strings.HasPrefix(rel, "..") {
		return filepath.ToSlash(path)
	}
	return filepath.ToSlash(rel)
}

type sessionMetadata struct {
	PID          int    `json:"pid"`
	WorkbookPath string `json:"workbook_path"`
}

// readSessionMetadata loads .xlflow/session.json. found is false when no session
// file exists; a present-but-unreadable or malformed file returns an error.
func (a *app) readSessionMetadata() (sessionMetadata, bool, error) {
	body, err := os.ReadFile(filepath.Join(a.cwd, ".xlflow", "session.json"))
	if errors.Is(err, os.ErrNotExist) {
		return sessionMetadata{}, false, nil
	}
	if err != nil {
		return sessionMetadata{}, false, err
	}
	body = bytes.TrimPrefix(body, []byte{0xEF, 0xBB, 0xBF})
	var metadata sessionMetadata
	if err := json.Unmarshal(body, &metadata); err != nil {
		return sessionMetadata{}, false, err
	}
	return metadata, true, nil
}

// packActiveSession reports whether .xlflow/session.json records a session for any
// of the candidate workbooks. pack must not read possibly-dirty live state, so a
// recorded session blocks it regardless of process liveness; the Office lock-file
// check covers the open-in-Excel case separately.
func (a *app) packActiveSession(candidates []string) (string, bool, error) {
	metadata, found, err := a.readSessionMetadata()
	if err != nil || !found {
		return "", false, err
	}
	for _, candidate := range candidates {
		if samePath(metadata.WorkbookPath, candidate) {
			return metadata.WorkbookPath, true, nil
		}
	}
	return "", false, nil
}

func (a *app) rollbackBlockedBySession(workbookPath string) (bool, error) {
	metadata, found, err := a.readSessionMetadata()
	if err != nil || !found {
		return false, err
	}
	if !samePath(metadata.WorkbookPath, workbookPath) {
		return false, nil
	}
	if metadata.PID <= 0 {
		return true, nil
	}
	if runtime.GOOS != "windows" {
		return false, nil
	}
	running, err := processRunning(metadata.PID)
	if err != nil {
		return false, err
	}
	return running, nil
}

func processRunning(pid int) (bool, error) {
	cmd := exec.Command("tasklist", "/FI", fmt.Sprintf("PID eq %d", pid), "/FO", "CSV", "/NH")
	out, err := cmd.Output()
	if err != nil {
		return false, err
	}
	return tasklistPIDRunning(out, pid)
}

func tasklistPIDRunning(out []byte, pid int) (bool, error) {
	reader := csv.NewReader(bytes.NewReader(out))
	reader.FieldsPerRecord = -1
	for {
		record, readErr := reader.Read()
		if errors.Is(readErr, io.EOF) {
			return false, nil
		}
		if readErr != nil {
			return false, fmt.Errorf("parse tasklist output: %w", readErr)
		}
		if len(record) < 2 {
			continue
		}
		recordedPID, parseErr := strconv.Atoi(strings.TrimSpace(record[1]))
		if parseErr == nil && recordedPID == pid {
			return true, nil
		}
	}
}

func looksLikeWorkbookInUse(err error) bool {
	if err == nil {
		return false
	}
	message := strings.ToLower(err.Error())
	return strings.Contains(message, "being used by another process") ||
		strings.Contains(message, "used by another process") ||
		strings.Contains(message, "permission denied") ||
		strings.Contains(message, "access is denied")
}

func samePath(a, b string) bool {
	return coordination.SamePath(a, b)
}

func buildExportImageOptions(workbook, sheet, cellRange, outPath, outputDir, name, format string, overwrite bool, session bool, keepalive excel.CommandOptions) (excel.ExportImageOptions, error) {
	sheet = strings.TrimSpace(sheet)
	if sheet == "" {
		return excel.ExportImageOptions{}, fmt.Errorf("--sheet is required")
	}
	normalizedRange, err := validateInspectRangeAddress(cellRange)
	if err != nil {
		return excel.ExportImageOptions{}, fmt.Errorf("--range %w", err)
	}
	outPath = strings.TrimSpace(outPath)
	outputDir = strings.TrimSpace(outputDir)
	name = strings.TrimSpace(name)
	if outPath != "" && (outputDir != "" || name != "") {
		return excel.ExportImageOptions{}, fmt.Errorf("--out cannot be combined with --output-dir or --name")
	}
	if err := validateWindowsFilename(name); err != nil {
		return excel.ExportImageOptions{}, err
	}
	return excel.ExportImageOptions{
		WorkbookPath: strings.TrimSpace(workbook),
		Sheet:        sheet,
		Range:        normalizedRange,
		OutPath:      outPath,
		OutputDir:    outputDir,
		Name:         name,
		Format:       strings.TrimSpace(format),
		Overwrite:    overwrite,
		Session:      session,
		Keepalive:    keepalive,
	}, nil
}

func buildEditCellOptions(workbook, sheet, cell, fill, events string, value *string, formula *string, session bool, keepalive excel.CommandOptions) (excel.EditCellOptions, error) {
	if !session {
		return excel.EditCellOptions{}, fmt.Errorf("`xlflow edit` requires --session")
	}
	sheet = strings.TrimSpace(sheet)
	if sheet == "" {
		return excel.EditCellOptions{}, fmt.Errorf("--sheet is required")
	}
	cellAddress, err := validateInspectCellAddress(cell)
	if err != nil {
		return excel.EditCellOptions{}, fmt.Errorf("--cell %w", err)
	}
	fill = strings.TrimSpace(fill)
	events = strings.ToLower(strings.TrimSpace(events))
	if events == "" {
		events = string(excel.EditEventKeep)
	}
	if events != string(excel.EditEventKeep) && events != string(excel.EditEventOn) && events != string(excel.EditEventOff) {
		return excel.EditCellOptions{}, fmt.Errorf("--events must be keep, on, or off")
	}
	mutations := 0
	if value != nil {
		mutations++
	}
	if formula != nil {
		mutations++
	}
	if fill != "" {
		mutations++
	}
	if mutations != 1 {
		return excel.EditCellOptions{}, fmt.Errorf("exactly one of --value, --formula, or --fill is required")
	}
	if fill != "" {
		normalized, err := normalizeEditColor(fill)
		if err != nil {
			return excel.EditCellOptions{}, err
		}
		fill = normalized
	}
	if fill != "" && events != string(excel.EditEventKeep) {
		return excel.EditCellOptions{}, fmt.Errorf("--events applies only to --value or --formula edits")
	}
	return excel.EditCellOptions{
		WorkbookPath: strings.TrimSpace(workbook),
		Sheet:        sheet,
		Cell:         cellAddress,
		Value:        value,
		Formula:      formula,
		Fill:         fill,
		Events:       excel.EditEventMode(events),
		Session:      session,
		Keepalive:    keepalive,
	}, nil
}

func buildEditRangeOptions(workbook, sheet, cellRange, fill, clear string, session bool, keepalive excel.CommandOptions) (excel.EditRangeOptions, error) {
	if !session {
		return excel.EditRangeOptions{}, fmt.Errorf("`xlflow edit` requires --session")
	}
	sheet = strings.TrimSpace(sheet)
	if sheet == "" {
		return excel.EditRangeOptions{}, fmt.Errorf("--sheet is required")
	}
	normalizedRange, err := validateInspectRangeAddress(cellRange)
	if err != nil {
		return excel.EditRangeOptions{}, fmt.Errorf("--range %w", err)
	}
	fill = strings.TrimSpace(fill)
	clear = strings.ToLower(strings.TrimSpace(clear))
	if fill != "" && clear != "" {
		return excel.EditRangeOptions{}, fmt.Errorf("--fill and --clear cannot be combined")
	}
	if fill == "" && clear == "" {
		return excel.EditRangeOptions{}, fmt.Errorf("one of --fill or --clear is required")
	}
	if fill != "" {
		normalized, err := normalizeEditColor(fill)
		if err != nil {
			return excel.EditRangeOptions{}, err
		}
		fill = normalized
	}
	if clear != "" && clear != "contents" && clear != "formats" && clear != "all" {
		return excel.EditRangeOptions{}, fmt.Errorf("--clear must be contents, formats, or all")
	}
	return excel.EditRangeOptions{
		WorkbookPath: strings.TrimSpace(workbook),
		Sheet:        sheet,
		Range:        normalizedRange,
		Fill:         fill,
		Clear:        clear,
		Session:      session,
		Keepalive:    keepalive,
	}, nil
}

func buildEditFormulaOptions(workbook, sheet, cellRange, events string, formula *string, formulaR1C1 *string, calculate bool, session bool, keepalive excel.CommandOptions) (excel.EditFormulaOptions, error) {
	if !session {
		return excel.EditFormulaOptions{}, fmt.Errorf("`xlflow edit` requires --session")
	}
	sheet = strings.TrimSpace(sheet)
	if sheet == "" {
		return excel.EditFormulaOptions{}, fmt.Errorf("--sheet is required")
	}
	normalizedRange, err := validateInspectRangeAddress(cellRange)
	if err != nil {
		return excel.EditFormulaOptions{}, fmt.Errorf("--range %w", err)
	}
	events = strings.ToLower(strings.TrimSpace(events))
	if events == "" {
		events = string(excel.EditEventKeep)
	}
	if events != string(excel.EditEventKeep) && events != string(excel.EditEventOn) && events != string(excel.EditEventOff) {
		return excel.EditFormulaOptions{}, fmt.Errorf("--events must be keep, on, or off")
	}
	mutations := 0
	if formula != nil {
		mutations++
	}
	if formulaR1C1 != nil {
		mutations++
	}
	if mutations != 1 {
		return excel.EditFormulaOptions{}, fmt.Errorf("exactly one of --formula or --formula-r1c1 is required")
	}
	return excel.EditFormulaOptions{
		WorkbookPath: strings.TrimSpace(workbook),
		Sheet:        sheet,
		Range:        normalizedRange,
		Formula:      formula,
		FormulaR1C1:  formulaR1C1,
		Events:       excel.EditEventMode(events),
		Calculate:    calculate,
		Session:      session,
		Keepalive:    keepalive,
	}, nil
}

func buildEditRowsOptions(workbook, sheet, rows string, height float64, session bool, keepalive excel.CommandOptions) (excel.EditRowsOptions, error) {
	if !session {
		return excel.EditRowsOptions{}, fmt.Errorf("`xlflow edit` requires --session")
	}
	sheet = strings.TrimSpace(sheet)
	if sheet == "" {
		return excel.EditRowsOptions{}, fmt.Errorf("--sheet is required")
	}
	normalizedRows, err := validateEditRowsSelector(rows)
	if err != nil {
		return excel.EditRowsOptions{}, err
	}
	if height <= 0 {
		return excel.EditRowsOptions{}, fmt.Errorf("--height must be greater than 0")
	}
	return excel.EditRowsOptions{
		WorkbookPath: strings.TrimSpace(workbook),
		Sheet:        sheet,
		Rows:         normalizedRows,
		Height:       height,
		Session:      session,
		Keepalive:    keepalive,
	}, nil
}

func buildEditColumnsOptions(workbook, sheet, columns string, width float64, session bool, keepalive excel.CommandOptions) (excel.EditColumnsOptions, error) {
	if !session {
		return excel.EditColumnsOptions{}, fmt.Errorf("`xlflow edit` requires --session")
	}
	sheet = strings.TrimSpace(sheet)
	if sheet == "" {
		return excel.EditColumnsOptions{}, fmt.Errorf("--sheet is required")
	}
	normalizedColumns, err := validateEditColumnsSelector(columns)
	if err != nil {
		return excel.EditColumnsOptions{}, err
	}
	if width <= 0 {
		return excel.EditColumnsOptions{}, fmt.Errorf("--width must be greater than 0")
	}
	return excel.EditColumnsOptions{
		WorkbookPath: strings.TrimSpace(workbook),
		Sheet:        sheet,
		Columns:      normalizedColumns,
		Width:        width,
		Session:      session,
		Keepalive:    keepalive,
	}, nil
}

var errInvalidWorksheetName = errors.New("invalid worksheet name")

func buildEditSheetAddOptions(workbook, name, before, after string, ifMissing, session bool, keepalive excel.CommandOptions) (excel.EditSheetAddOptions, error) {
	if !session {
		return excel.EditSheetAddOptions{}, fmt.Errorf("`xlflow edit` requires --session")
	}
	name = strings.TrimSpace(name)
	if name == "" {
		return excel.EditSheetAddOptions{}, fmt.Errorf("--name is required")
	}
	if err := validateExcelWorksheetName(name); err != nil {
		return excel.EditSheetAddOptions{}, err
	}
	before = strings.TrimSpace(before)
	after = strings.TrimSpace(after)
	if before != "" && after != "" {
		return excel.EditSheetAddOptions{}, fmt.Errorf("--before and --after cannot be combined")
	}
	return excel.EditSheetAddOptions{
		WorkbookPath: strings.TrimSpace(workbook),
		Name:         name,
		Before:       before,
		After:        after,
		IfMissing:    ifMissing,
		Session:      session,
		Keepalive:    keepalive,
	}, nil
}

func validateExcelWorksheetName(name string) error {
	if utf8.RuneCountInString(name) > 31 {
		return fmt.Errorf("%w: --name must be 31 characters or fewer", errInvalidWorksheetName)
	}
	if strings.HasPrefix(name, "'") || strings.HasSuffix(name, "'") {
		return fmt.Errorf("%w: --name cannot start or end with an apostrophe", errInvalidWorksheetName)
	}
	for _, r := range name {
		if unicode.IsControl(r) {
			return fmt.Errorf("%w: --name cannot contain control characters", errInvalidWorksheetName)
		}
		if strings.ContainsRune(`:\/?*[]`, r) {
			return fmt.Errorf("%w: --name cannot contain : \\ / ? * [ ]", errInvalidWorksheetName)
		}
	}
	return nil
}

func normalizeEditColor(value string) (string, error) {
	value = strings.TrimSpace(value)
	if value == "" {
		return "", fmt.Errorf("--fill is required")
	}
	if !strings.HasPrefix(value, "#") {
		return "", fmt.Errorf("--fill must use #RGB or #RRGGBB")
	}
	hex := strings.ToUpper(strings.TrimPrefix(value, "#"))
	switch len(hex) {
	case 3:
		for _, r := range hex {
			if !isHexDigit(r) {
				return "", fmt.Errorf("--fill must use #RGB or #RRGGBB")
			}
		}
		return fmt.Sprintf("#%c%c%c%c%c%c", hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]), nil
	case 6:
		for _, r := range hex {
			if !isHexDigit(r) {
				return "", fmt.Errorf("--fill must use #RGB or #RRGGBB")
			}
		}
		return "#" + hex, nil
	default:
		return "", fmt.Errorf("--fill must use #RGB or #RRGGBB")
	}
}

func isHexDigit(r rune) bool {
	return (r >= '0' && r <= '9') || (r >= 'A' && r <= 'F') || (r >= 'a' && r <= 'f')
}

func validateEditRowsSelector(value string) (string, error) {
	value = strings.TrimSpace(value)
	if value == "" {
		return "", fmt.Errorf("--rows is required")
	}
	parts := strings.SplitN(value, ":", 2)
	start, err := parsePositiveRow(parts[0])
	if err != nil {
		return "", err
	}
	end := start
	if len(parts) == 2 {
		end, err = parsePositiveRow(parts[1])
		if err != nil {
			return "", err
		}
	}
	if end < start {
		start, end = end, start
	}
	if start == end {
		return strconv.Itoa(start), nil
	}
	return fmt.Sprintf("%d:%d", start, end), nil
}

func parsePositiveRow(value string) (int, error) {
	value = strings.TrimSpace(value)
	row, err := strconv.Atoi(value)
	if err != nil || row <= 0 {
		return 0, fmt.Errorf("--rows must use a positive row selector such as 1 or 1:31")
	}
	return row, nil
}

func validateEditColumnsSelector(value string) (string, error) {
	value = strings.TrimSpace(value)
	if value == "" {
		return "", fmt.Errorf("--columns is required")
	}
	parts := strings.SplitN(value, ":", 2)
	start, err := parseColumnSelector(parts[0])
	if err != nil {
		return "", err
	}
	end := start
	if len(parts) == 2 {
		end, err = parseColumnSelector(parts[1])
		if err != nil {
			return "", err
		}
	}
	if end < start {
		start, end = end, start
	}
	first, err := excelize.ColumnNumberToName(start)
	if err != nil {
		return "", fmt.Errorf("--columns must use Excel column letters such as A or A:AE")
	}
	if start == end {
		return first, nil
	}
	last, err := excelize.ColumnNumberToName(end)
	if err != nil {
		return "", fmt.Errorf("--columns must use Excel column letters such as A or A:AE")
	}
	return first + ":" + last, nil
}

func parseColumnSelector(value string) (int, error) {
	value = strings.ToUpper(strings.TrimSpace(value))
	if value == "" {
		return 0, fmt.Errorf("--columns must use Excel column letters such as A or A:AE")
	}
	for _, r := range value {
		if r < 'A' || r > 'Z' {
			return 0, fmt.Errorf("--columns must use Excel column letters such as A or A:AE")
		}
	}
	column, err := excelize.ColumnNameToNumber(value)
	if err != nil || column <= 0 {
		return 0, fmt.Errorf("--columns must use Excel column letters such as A or A:AE")
	}
	return column, nil
}

func validateWindowsFilename(name string) error {
	if name == "" {
		return nil
	}
	if strings.Contains(name, "/") || strings.Contains(name, "\\") || filepath.Base(name) != name {
		return fmt.Errorf("--name must be a filename without path separators or invalid Windows characters")
	}
	if strings.ContainsAny(name, `<>:"|?*`) {
		return fmt.Errorf("--name must be a filename without path separators or invalid Windows characters")
	}
	for _, r := range name {
		if unicode.IsControl(r) {
			return fmt.Errorf("--name must be a filename without path separators or invalid Windows characters")
		}
	}
	return nil
}

func sessionUsageHint() string {
	return "reuse the matching xlflow session workbook when available"
}

func buildCommandOptions(stderr io.Writer) excel.CommandOptions {
	if stderr == nil {
		stderr = os.Stderr
	}
	return excel.CommandOptions{Stderr: stderr, Progress: true}
}

var allowedMsgBoxResults = map[string]struct{}{
	"abort":  {},
	"cancel": {},
	"ignore": {},
	"no":     {},
	"ok":     {},
	"retry":  {},
	"yes":    {},
}

var allowedFileDialogKinds = map[string]struct{}{
	"get-open":  {},
	"file-open": {},
	"save-as":   {},
	"folder":    {},
}

func normalizeUIResponseID(value string) string {
	trimmed := strings.ToLower(strings.TrimSpace(value))
	var builder strings.Builder
	lastSeparator := false
	for _, r := range trimmed {
		isValid := (r >= 'a' && r <= 'z') || (r >= '0' && r <= '9')
		if isValid {
			builder.WriteRune(r)
			lastSeparator = false
			continue
		}
		if builder.Len() > 0 && !lastSeparator {
			builder.WriteByte('_')
			lastSeparator = true
		}
	}
	return strings.Trim(builder.String(), "_")
}

func parseUIResponseLiterals(flagName string, literals []string, normalizer func(string) (string, error)) (map[string]string, error) {
	responses := make(map[string]string, len(literals))
	normalizedIDs := make(map[string]string, len(literals))
	for _, literal := range literals {
		parts := strings.SplitN(literal, "=", 2)
		if len(parts) != 2 {
			return nil, fmt.Errorf("invalid --%s %q: expected id=value", flagName, literal)
		}
		id := strings.TrimSpace(parts[0])
		if id == "" {
			return nil, fmt.Errorf("invalid --%s %q: dialog id cannot be empty", flagName, literal)
		}
		normalizedID := normalizeUIResponseID(id)
		if normalizedID == "" {
			return nil, fmt.Errorf("invalid --%s %q: dialog id must contain at least one ASCII letter or digit", flagName, literal)
		}
		if existing, exists := normalizedIDs[normalizedID]; exists {
			if existing == id {
				return nil, fmt.Errorf("invalid --%s %q: duplicate dialog id %q", flagName, literal, id)
			}
			return nil, fmt.Errorf("invalid --%s %q: dialog id %q collides with %q after normalization", flagName, literal, id, existing)
		}
		value := parts[1]
		if normalizer != nil {
			normalizedValue, err := normalizer(value)
			if err != nil {
				return nil, fmt.Errorf("invalid --%s %q: %w", flagName, literal, err)
			}
			value = normalizedValue
		}
		responses[normalizedID] = value
		normalizedIDs[normalizedID] = id
	}
	return responses, nil
}

func normalizeMsgBoxResponseToken(value string) (string, error) {
	trimmed := strings.ToLower(strings.TrimSpace(value))
	if trimmed == "" {
		return "", fmt.Errorf("msgbox result cannot be empty")
	}
	if _, ok := allowedMsgBoxResults[trimmed]; !ok {
		return "", fmt.Errorf("unsupported msgbox result %q", value)
	}
	return trimmed, nil
}

func normalizeFileDialogKind(value string) (string, error) {
	trimmed := strings.ToLower(strings.TrimSpace(value))
	if _, ok := allowedFileDialogKinds[trimmed]; !ok {
		return "", fmt.Errorf("unsupported filedialog kind %q", value)
	}
	return trimmed, nil
}

func parseFileDialogResponseLiterals(literals []string) ([]excel.FileDialogResponse, error) {
	responses := make(map[string]*excel.FileDialogResponse, len(literals))
	originalIDs := make(map[string]string, len(literals))
	order := make([]string, 0, len(literals))
	for _, literal := range literals {
		parts := strings.SplitN(literal, "=", 2)
		if len(parts) != 2 {
			return nil, fmt.Errorf("invalid --filedialog %q: expected kind:id=value", literal)
		}
		selector := strings.SplitN(strings.TrimSpace(parts[0]), ":", 2)
		if len(selector) != 2 {
			return nil, fmt.Errorf("invalid --filedialog %q: expected kind:id=value", literal)
		}
		kind, err := normalizeFileDialogKind(selector[0])
		if err != nil {
			return nil, fmt.Errorf("invalid --filedialog %q: %w", literal, err)
		}
		id := strings.TrimSpace(selector[1])
		if id == "" {
			return nil, fmt.Errorf("invalid --filedialog %q: dialog id cannot be empty", literal)
		}
		normalizedID := normalizeUIResponseID(id)
		if normalizedID == "" {
			return nil, fmt.Errorf("invalid --filedialog %q: dialog id must contain at least one ASCII letter or digit", literal)
		}
		key := kind + ":" + normalizedID
		if existing, exists := originalIDs[key]; exists && existing != id {
			return nil, fmt.Errorf("invalid --filedialog %q: dialog id %q collides with %q after normalization", literal, id, existing)
		}
		response, exists := responses[key]
		if !exists {
			response = &excel.FileDialogResponse{Kind: kind, DialogID: normalizedID}
			responses[key] = response
			originalIDs[key] = id
			order = append(order, key)
		}
		value := parts[1]
		if strings.TrimSpace(value) == "@cancel" {
			if response.Cancelled || len(response.Values) > 0 {
				return nil, fmt.Errorf("invalid --filedialog %q: @cancel cannot be combined with path values for the same dialog", literal)
			}
			response.Cancelled = true
			continue
		}
		if strings.TrimSpace(value) == "" {
			return nil, fmt.Errorf("invalid --filedialog %q: scripted path cannot be empty", literal)
		}
		if response.Cancelled {
			return nil, fmt.Errorf("invalid --filedialog %q: path values cannot be combined with @cancel for the same dialog", literal)
		}
		response.Values = append(response.Values, value)
		if (kind == "save-as" || kind == "folder") && len(response.Values) > 1 {
			return nil, fmt.Errorf("invalid --filedialog %q: kind %q accepts only one scripted path", literal, kind)
		}
	}
	parsed := make([]excel.FileDialogResponse, 0, len(order))
	for _, key := range order {
		parsed = append(parsed, *responses[key])
	}
	return parsed, nil
}

func buildRunOptions(cfg config.Config, macro, input string, argLiterals []string, msgBoxLiterals []string, inputBoxLiterals []string, fileDialogLiterals []string, save bool, saveAs string, headless bool, interactive bool, direct bool, fast bool, diagnostic bool, diagnosticExplicit bool, guiCompileErrors bool, session bool, timeout time.Duration, commandOpts excel.CommandOptions) (excel.RunOptions, error) {
	return buildRunOptionsWithUIStream(cfg, macro, input, argLiterals, msgBoxLiterals, inputBoxLiterals, fileDialogLiterals, save, saveAs, headless, interactive, direct, fast, diagnostic, diagnosticExplicit, guiCompileErrors, session, timeout, commandOpts, false)
}

func buildRunOptionsWithUIStream(cfg config.Config, macro, input string, argLiterals []string, msgBoxLiterals []string, inputBoxLiterals []string, fileDialogLiterals []string, save bool, saveAs string, headless bool, interactive bool, direct bool, fast bool, diagnostic bool, diagnosticExplicit bool, guiCompileErrors bool, session bool, timeout time.Duration, commandOpts excel.CommandOptions, uiStream bool) (excel.RunOptions, error) {
	if save && saveAs != "" {
		return excel.RunOptions{}, fmt.Errorf("--save and --save-as cannot be combined")
	}
	if headless && interactive {
		return excel.RunOptions{}, fmt.Errorf("--headless and --interactive cannot be combined")
	}
	if guiCompileErrors && diagnosticExplicit && diagnostic {
		return excel.RunOptions{}, fmt.Errorf("--diagnostic and --gui-compile-errors cannot be combined")
	}
	if guiCompileErrors {
		diagnostic = false
	}
	if fast && uiStream {
		return excel.RunOptions{}, fmt.Errorf("--fast cannot be combined with --ui-stream")
	}
	if direct && uiStream {
		return excel.RunOptions{}, fmt.Errorf("--direct cannot be combined with --ui-stream")
	}
	if direct && diagnostic {
		if diagnosticExplicit {
			return excel.RunOptions{}, fmt.Errorf("--direct cannot be combined with diagnostic mode; omit --diagnostic or use --gui-compile-errors")
		}
		diagnostic = false
	}
	if fast && diagnostic && !diagnosticExplicit {
		diagnostic = false
	}
	if macro == "" {
		macro = cfg.Project.Entry
	}
	args := make([]excel.RunArgument, 0, len(argLiterals))
	for _, literal := range argLiterals {
		parts := strings.SplitN(literal, ":", 2)
		if len(parts) != 2 {
			return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: expected type:value", literal)
		}
		switch parts[0] {
		case "string":
		case "int", "double", "bool":
			if parts[1] == "" {
				return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: %s values cannot be empty", literal, parts[0])
			}
			if parts[0] == "int" {
				if _, err := strconv.Atoi(parts[1]); err != nil {
					return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: int values must parse as base-10 integers", literal)
				}
			}
			if parts[0] == "bool" && parts[1] != "true" && parts[1] != "false" {
				return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: bool values must be true or false", literal)
			}
			if parts[0] == "double" {
				value, err := strconv.ParseFloat(parts[1], 64)
				if err != nil || math.IsNaN(value) || math.IsInf(value, 0) {
					return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: double values must parse as finite invariant-culture numbers", literal)
				}
			}
		default:
			return excel.RunOptions{}, fmt.Errorf("unsupported --arg type prefix %q", parts[0])
		}
		args = append(args, excel.RunArgument{Type: parts[0], Value: parts[1]})
	}
	if direct && len(args) > 0 {
		return excel.RunOptions{}, fmt.Errorf("--direct cannot be used with --arg")
	}
	msgBoxResponses, err := parseUIResponseLiterals("msgbox", msgBoxLiterals, normalizeMsgBoxResponseToken)
	if err != nil {
		return excel.RunOptions{}, err
	}
	inputResponses, err := parseUIResponseLiterals("inputbox", inputBoxLiterals, nil)
	if err != nil {
		return excel.RunOptions{}, err
	}
	fileDialogResponses, err := parseFileDialogResponseLiterals(fileDialogLiterals)
	if err != nil {
		return excel.RunOptions{}, err
	}
	mode := ""
	if headless {
		mode = "headless"
	}
	if interactive {
		mode = "interactive"
	}
	runtime := excel.ResolveRunRuntimeInfo(mode)
	if mode == "" && runtime.Mode != excel.RuntimeModeInteractive {
		mode = excel.RuntimeModeHeadless
	}
	return excel.RunOptions{
		Macro:               macro,
		WorkbookPath:        input,
		Args:                args,
		UIResponses:         excel.UIResponses{MsgBox: msgBoxResponses, Input: inputResponses, FileDialog: fileDialogResponses},
		DebugStream:         excel.DebugStreamOptions{Enabled: true},
		UIStream:            excel.UIStreamOptions{Enabled: uiStream, RedactInput: true},
		Save:                save,
		SaveAs:              saveAs,
		Mode:                mode,
		RuntimeMode:         runtime.Mode,
		RuntimeSource:       runtime.Source,
		Direct:              direct,
		Fast:                fast,
		Diagnostic:          diagnostic,
		SuppressModalErrors: !guiCompileErrors,
		Session:             session,
		Timeout:             timeout,
		Keepalive:           commandOpts,
	}, nil
}

func (a *app) sessionCommand() *cobra.Command {
	session := &cobra.Command{
		Use:   "session",
		Short: "Manage an xlflow Excel session",
	}
	for _, action := range []string{"start", "status", "stop"} {
		action := action
		var discard bool
		cmd := &cobra.Command{
			Use:   action,
			Short: action + " the xlflow Excel session",
			Args:  cobra.NoArgs,
			RunE: func(cmd *cobra.Command, args []string) error {
				cfg, err := a.loadConfig("session")
				if err != nil {
					return err
				}
				run := func() (output.Envelope, int, error) {
					return a.excelRunnerForConfig(cfg).Session(cfg, action, excel.SessionCommandOptions{Discard: discard})
				}
				var env output.Envelope
				var code int
				if action == "status" {
					env, code, err = a.runSessionStatus(cmd.Context(), cfg, run)
				} else {
					env, code, err = run()
				}
				if err != nil {
					return err
				}
				return a.write(env, code)
			},
		}
		if action == "stop" {
			cmd.Flags().BoolVar(&discard, "discard", false, "close a managed session without saving unsaved workbook changes")
		}
		session.AddCommand(cmd)
	}
	attach := &cobra.Command{
		Use:   "attach",
		Short: "Attach xlflow to an already-open Excel workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("session attach")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Attaching open workbook", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).SessionAttach(cfg, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	session.AddCommand(attach)
	return session
}

func (a *app) saveCommand() *cobra.Command {
	var session bool
	cmd := &cobra.Command{
		Use:   "save",
		Short: "Save the current xlflow session workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("save")
			if err != nil {
				return err
			}
			env, code, err := a.excelRunnerForConfig(cfg).SaveSession(cfg, excel.SessionCommandOptions{Session: session})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) statusCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "status",
		Short: "Show project, source, workbook, and session state",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("status")
			if err != nil {
				return err
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			env := output.New("status")
			projectPayload := buildStatusProject(a.cwd, cfg, workbookPath)
			statePayload := buildStatusState(a.cwd, cfg, workbookPath)
			sessionState := a.buildStatusSession(cfg, workbookPath)
			if liveNewer, ok := sessionState["live_newer_than_disk"]; ok {
				statePayload["live_session_newer_than_disk"] = liveNewer
			}
			if sourceOfTruth, ok := sessionState["source_of_truth"]; ok {
				statePayload["source_of_truth"] = sourceOfTruth
			}
			if active := boolValueForCLI(sessionState, "active"); active {
				statePayload["workbook_saved"] = !boolValueForCLI(sessionState, "save_required")
			}
			env.Project = projectPayload
			env.Session = sessionState
			env.State = statePayload
			if coordinationStatus, unavailable := a.observeSessionCoordination(cmd.Context(), cfg); coordinationStatus != nil {
				env.Coordination = coordinationStatus
			} else if unavailable {
				markStatusRecoveryCheckFailed(sessionState, statePayload, &env)
				appendUniqueMessage(&env.Warnings, coordinationStatusUnavailableCode, "Workbook coordination and recovery status could not be observed.")
			}
			warnings, hints := buildStatusWarningsAndHints(sessionState, statePayload)
			env.Warnings = appendObjectMessages(env.Warnings, warnings)
			env.Hints = hints
			env.Logs = []string{"status reported"}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func markStatusRecoveryCheckFailed(session, state map[string]any, env *output.Envelope) {
	session["dirty"] = nil
	session["workbook_open"] = nil
	session["source_of_truth"] = "uncertain"
	session["discard_required"] = true
	session["recovery_required"] = nil
	session["recovery_check_failed"] = true
	state["source_of_truth"] = "uncertain"
	state["workbook_saved"] = nil
	if env != nil {
		env.Coordination = map[string]any{
			"busy":                  nil,
			"recovery_required":     nil,
			"recovery_check_failed": true,
		}
	}
}

func buildStatusProject(root string, cfg config.Config, workbookPath string) map[string]any {
	return map[string]any{
		"root":          displayPath(root, root),
		"workbook_path": displayPath(root, workbookPath),
		"src_paths": []string{
			displayPath(root, workbookArgPath(root, cfg.Src.Modules)),
			displayPath(root, workbookArgPath(root, cfg.Src.Classes)),
			displayPath(root, workbookArgPath(root, cfg.Src.Forms)),
			displayPath(root, workbookArgPath(root, cfg.Src.Workbook)),
		},
		"project_name": cfg.Project.Name,
	}
}

func buildStatusState(root string, cfg config.Config, workbookPath string) map[string]any {
	state := map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": false,
		"workbook_saved":               true,
		"source_of_truth":              "saved_workbook",
	}
	workbookInfo, wbErr := os.Stat(workbookPath)
	if wbErr == nil {
		state["workbook_last_modified_at"] = workbookInfo.ModTime().Format(time.RFC3339)
	}
	srcPaths := []string{
		workbookArgPath(root, cfg.Src.Modules),
		workbookArgPath(root, cfg.Src.Classes),
		workbookArgPath(root, cfg.Src.Forms),
		workbookArgPath(root, cfg.Src.Workbook),
	}
	latestMtime := latestSourceModTime(srcPaths)
	if !latestMtime.IsZero() {
		state["latest_source_modified_at"] = latestMtime.Format(time.RFC3339)
		if wbErr == nil && latestMtime.After(workbookInfo.ModTime()) {
			state["src_newer_than_workbook"] = true
		}
	}
	pushStatePath := filepath.Join(root, ".xlflow", "state", "push.json")
	if pushInfo, err := os.Stat(pushStatePath); err == nil {
		state["push_state_last_modified_at"] = pushInfo.ModTime().Format(time.RFC3339)
	}
	return state
}

func latestSourceModTime(srcPaths []string) time.Time {
	var latest time.Time
	for _, srcPath := range srcPaths {
		_ = filepath.WalkDir(srcPath, func(path string, d os.DirEntry, err error) error {
			if err != nil || d.IsDir() {
				return nil
			}
			ext := strings.ToLower(filepath.Ext(d.Name()))
			switch ext {
			case ".bas", ".cls", ".frm", ".frx":
			default:
				return nil
			}
			info, err := d.Info()
			if err != nil {
				return nil
			}
			if info.ModTime().After(latest) {
				latest = info.ModTime()
			}
			return nil
		})
	}
	return latest
}

func buildStatusWarningsAndHints(session, state map[string]any) ([]map[string]any, []map[string]any) {
	var warnings []map[string]any
	var hints []map[string]any
	if boolValueForCLI(session, "active") {
		if boolValueForCLI(session, "save_required") {
			warnings = append(warnings, map[string]any{
				"code":    "session_dirty",
				"message": "The live session workbook has unsaved changes.",
			})
			hints = append(hints, map[string]any{
				"code":    "save_session",
				"message": "Run `xlflow save --session` to persist the live workbook to disk.",
			})
		}
	}
	if boolValueForCLI(state, "src_newer_than_workbook") {
		warnings = append(warnings, map[string]any{
			"code":    "source_newer_than_workbook",
			"message": "Source files are newer than the saved workbook.",
		})
		hints = append(hints, map[string]any{
			"code":    "push_source",
			"message": "Run `xlflow push` to import the latest source into the workbook.",
		})
	}
	if boolValueForCLI(state, "live_session_newer_than_disk") {
		warnings = append(warnings, map[string]any{
			"code":    "live_session_newer_than_disk",
			"message": "The live session workbook is newer than the saved workbook on disk.",
		})
		hints = append(hints, map[string]any{
			"code":    "save_before_push",
			"message": "Run `xlflow save --session` before `xlflow push` when the live session has unsaved changes you want to keep.",
		})
	}
	return warnings, hints
}

func (a *app) buildStatusSession(cfg config.Config, workbookPath string) map[string]any {
	session := map[string]any{
		"active":               false,
		"workbook_path":        workbookPath,
		"workbook_name":        filepath.Base(workbookPath),
		"dirty":                false,
		"save_required":        false,
		"live_newer_than_disk": false,
		"source_of_truth":      "saved_workbook",
		"running":              false,
		"workbook_open":        false,
		"metadata":             nil,
	}
	if runtime.GOOS != "windows" {
		return session
	}
	status, ok := a.inspectSessionStatus(cfg)
	if !ok {
		return session
	}
	statusWorkbookPath := stringValueForCLI(status, "workbook_path")
	if strings.TrimSpace(statusWorkbookPath) == "" || !samePath(statusWorkbookPath, workbookPath) {
		return session
	}
	active := boolValueForCLI(status, "active") || (boolValueForCLI(status, "running") && boolValueForCLI(status, "workbook_open"))
	saveRequired := boolValueForCLI(status, "save_required") || boolValueForCLI(status, "needs_save")
	session["active"] = active
	if rawDirty, exists := status["dirty"]; exists {
		session["dirty"] = rawDirty
	}
	if sourceOfTruth := stringValueForCLI(status, "source_of_truth"); sourceOfTruth != "" {
		session["source_of_truth"] = sourceOfTruth
	}
	session["save_required"] = saveRequired
	session["live_newer_than_disk"] = saveRequired
	if saveRequired && stringValueForCLI(status, "source_of_truth") == "" {
		session["source_of_truth"] = "live_workbook"
	}
	if running, ok := status["running"]; ok {
		session["running"] = running
	}
	if open, ok := status["workbook_open"]; ok {
		session["workbook_open"] = open
	}
	if metadata, ok := status["metadata"]; ok {
		session["metadata"] = metadata
	}
	if mode := stringValueForCLI(status, "mode"); mode != "" {
		session["mode"] = mode
	}
	if name := stringValueForCLI(status, "workbook_name"); strings.TrimSpace(name) != "" {
		session["workbook_name"] = name
	}
	if present, ok := status["userforms_present"]; ok {
		session["userforms_present"] = present
	}
	if count, ok := status["userform_count"]; ok {
		session["userform_count"] = count
	}
	if known, ok := status["userforms_known"]; ok {
		session["userforms_known"] = known
	}
	if discardRequired, ok := status["discard_required"]; ok {
		session["discard_required"] = discardRequired
	}
	if recoveryRequired, ok := status["recovery_required"]; ok {
		session["recovery_required"] = recoveryRequired
	}
	return session
}

func (a *app) runnerCommand() *cobra.Command {
	runner := &cobra.Command{
		Use:   "runner",
		Short: "Manage the persistent xlflow runner module",
	}
	for _, action := range []string{"install", "remove", "status"} {
		action := action
		cmd := &cobra.Command{
			Use:   action,
			Short: action + " the persistent xlflow runner module",
			Args:  cobra.NoArgs,
			RunE: func(cmd *cobra.Command, args []string) error {
				cfg, err := a.loadConfig("runner")
				if err != nil {
					return err
				}
				env, code, err := a.excelRunnerForConfig(cfg).RunnerModule(cfg, action)
				if err != nil {
					return err
				}
				return a.write(env, code)
			},
		}
		runner.AddCommand(cmd)
	}
	return runner
}

func headlessGUIBoundaryLogs(cfg config.Config) []string {
	logs := []string{
		"Headless preflight scans the configured source tree, not the target macro call graph.",
		"Use xlflow run --interactive if a human can operate Excel dialogs.",
		"For simple dialogs, replace raw MsgBox/InputBox/file dialog calls with XlflowUI wrappers and drive them with --msgbox/--inputbox/--filedialog.",
		"For repeatable automation, refactor GUI entrypoints into parameterized headless procedures.",
	}
	if cfg.Lint.ForbidInteractiveInput {
		logs = append(logs, "If this project is intentionally interactive, set [lint].forbid_interactive_input = false to suppress VB007 warnings; headless preflight still blocks GUI boundaries.")
	}
	return logs
}

func (a *app) runCommand() *cobra.Command {
	var argLiterals []string
	var input string
	var msgBoxLiterals []string
	var inputBoxLiterals []string
	var fileDialogLiterals []string
	var save bool
	var noSave bool
	var saveAs string
	var headless bool
	var interactive bool
	var direct bool
	var fast bool
	var verbose bool
	var diagnostic bool
	var guiCompileErrors bool
	var session bool
	var timeout time.Duration
	var uiStream bool
	var push bool

	cmd := &cobra.Command{
		Use:   "run [macro]",
		Short: "Run a workbook macro",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("run")
			if err != nil {
				return err
			}
			if noSave && (save || strings.TrimSpace(saveAs) != "") {
				return a.writeFailure("run", output.ExitConfig, "run_args_invalid", fmt.Errorf("--no-save cannot be combined with --save or --save-as"))
			}
			macro := ""
			if len(args) == 1 {
				macro = args[0]
			}
			var opts excel.RunOptions
			commandOpts := buildCommandOptions(a.stderrWriter())
			commandOpts.Progress = false
			if uiStream {
				opts, err = buildRunOptionsWithUIStream(cfg, macro, input, argLiterals, msgBoxLiterals, inputBoxLiterals, fileDialogLiterals, save, saveAs, headless, interactive, direct, fast, diagnostic, cmd.Flags().Changed("diagnostic"), guiCompileErrors, session, timeout, commandOpts, true)
			} else {
				opts, err = buildRunOptions(cfg, macro, input, argLiterals, msgBoxLiterals, inputBoxLiterals, fileDialogLiterals, save, saveAs, headless, interactive, direct, fast, diagnostic, cmd.Flags().Changed("diagnostic"), guiCompileErrors, session, timeout, commandOpts)
			}
			if err != nil {
				return a.writeFailure("run", output.ExitConfig, "run_args_invalid", err)
			}
			if push {
				if strings.TrimSpace(input) != "" {
					return a.writeFailure("run", output.ExitConfig, "run_args_invalid", fmt.Errorf("--push cannot be combined with --input"))
				}
				pushOpts, err := buildRunPushOptions(session, fast, buildCommandOptions(a.stderrWriter()))
				if err != nil {
					return a.writeFailure("run", output.ExitConfig, "run_args_invalid", err)
				}
				pushEnv, pushCode, pushErr := a.pushSource("run", cfg, pushOpts, "Importing VBA source")
				if pushErr != nil {
					return pushErr
				}
				if pushCode != output.ExitSuccess {
					return a.write(pushEnv, pushCode)
				}
			}
			if a.shouldRunSourcePreflight(cfg, opts) {
				if err := a.runSourcePreflight("run", cfg, "running macros", nil, nil); err != nil {
					return err
				}
			}
			if opts.Mode == "headless" {
				boundaries, err := gui.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
				if err != nil {
					return a.writeFailure("run", output.ExitEnvironment, "gui_preflight_failed", err)
				}
				if len(boundaries) > 0 {
					env := output.Failure("run", output.Error{
						Code:    "gui_boundary_detected",
						Message: "Cannot run in headless mode because this project contains GUI interaction boundaries.",
						Source:  "xlflow",
						Phase:   "preflight",
					})
					env.GUIBoundaries = boundaries
					env.Logs = headlessGUIBoundaryLogs(cfg)
					return a.write(env, output.ExitValidation)
				}
			}
			var env output.Envelope
			var code int
			run := func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).Run(cfg, opts)
				return runErr
			}
			err = a.withSpinner("Running macro", run)
			if err != nil {
				return err
			}
			if shouldAttachRunDiagnostic(env) {
				env.RunDiagnostic = a.buildRunDiagnostic(cfg, env)
			}
			return a.writeWithOutputOptions(env, code, a.outputOptionsWithVerbose(verbose))
		},
	}
	cmd.Flags().StringArrayVar(&argLiterals, "arg", nil, "pass a typed macro argument such as string:hello, int:7, double:3.5, or bool:true")
	cmd.Flags().StringArrayVar(&msgBoxLiterals, "msgbox", nil, "provide a scripted MsgBox response as dialog-id=result")
	cmd.Flags().StringArrayVar(&inputBoxLiterals, "inputbox", nil, "provide a scripted InputBox response as dialog-id=value")
	cmd.Flags().StringArrayVar(&fileDialogLiterals, "filedialog", nil, "provide a scripted file dialog response as kind:dialog-id=path or kind:dialog-id=@cancel")
	cmd.Flags().StringVar(&input, "input", "", "override workbook path for this run")
	cmd.Flags().BoolVar(&save, "save", false, "save the opened workbook after a successful run")
	cmd.Flags().BoolVar(&noSave, "no-save", false, "leave the workbook unchanged on disk after the run")
	cmd.Flags().StringVar(&saveAs, "save-as", "", "write the successful workbook result to a new path")
	cmd.Flags().BoolVar(&headless, "headless", false, "reject GUI interaction boundaries before running the macro")
	cmd.Flags().BoolVar(&interactive, "interactive", false, "run with Excel visible and alerts enabled for human interaction")
	cmd.Flags().BoolVar(&direct, "direct", false, "run an argument-free macro without injecting a temporary harness")
	cmd.Flags().BoolVar(&fast, "fast", false, "use development-oriented fast run defaults")
	cmd.Flags().BoolVar(&verbose, "verbose", false, "include verbose diagnostic JSON fields for xlflow debugging")
	cmd.Flags().BoolVar(&diagnostic, "diagnostic", true, "compile VBA before running and return structured compile diagnostics (default true)")
	cmd.Flags().BoolVar(&guiCompileErrors, "gui-compile-errors", false, "allow VBA compile and runtime error dialogs to surface via the GUI instead of structured diagnostics")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	cmd.Flags().DurationVar(&timeout, "timeout", 5*time.Minute, "maximum macro runtime before xlflow reports a timeout")
	cmd.Flags().BoolVar(&uiStream, "ui-stream", false, "stream headless XlflowUI dialog events to stderr in real time")
	cmd.Flags().BoolVar(&push, "push", false, "import source VBA into the workbook before running the macro")
	return cmd
}

func buildRunPushOptions(session bool, fast bool, keepalive excel.CommandOptions) (excel.PushOptions, error) {
	pushOpts, err := buildPushOptions("always", fast, false, session, session, keepalive)
	if err != nil {
		return excel.PushOptions{}, err
	}
	return pushOpts, nil
}

func shouldAttachRunDiagnostic(env output.Envelope) bool {
	if env.Status != output.StatusFailed || env.Error == nil || env.Error.Phase != "invoke_macro" {
		return false
	}
	switch env.Error.Code {
	case "macro_failed", "macro_not_found", "macro_disabled":
		return true
	default:
		return false
	}
}

func (a *app) exportImageCommand() *cobra.Command {
	var sheet string
	var cellRange string
	var outPath string
	var outputDir string
	var name string
	var format string
	var overwrite bool
	var session bool

	cmd := &cobra.Command{
		Use:   "export-image [workbook]",
		Short: "Export a worksheet range as an image",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("export-image")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildExportImageOptions(workbook, sheet, cellRange, outPath, outputDir, name, format, overwrite, session, commandOpts)
			if err != nil {
				return a.writeFailure("export-image", output.ExitConfig, "export_image_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Exporting worksheet range image", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).ExportImage(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&cellRange, "range", "", "range address such as A1:AE31")
	cmd.Flags().StringVar(&outPath, "out", "", "write the image to an explicit file path")
	cmd.Flags().StringVar(&outputDir, "output-dir", "", "write the image into this directory using a generated filename")
	cmd.Flags().StringVar(&name, "name", "", "output filename only; uses the default image directory unless --output-dir is set")
	cmd.Flags().StringVar(&format, "format", "png", "image format; only png is currently supported")
	cmd.Flags().BoolVar(&overwrite, "overwrite", false, "replace an existing output file")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) editCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "edit",
		Short: "Mutate a live session workbook for development and testing",
	}
	cmd.AddCommand(
		a.editCellCommand(),
		a.editRangeCommand(),
		a.editFormulaCommand(),
		a.editRowsCommand(),
		a.editColumnsCommand(),
		a.editSheetCommand(),
	)
	return cmd
}

func (a *app) editSheetCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "sheet",
		Short: "Mutate worksheets in a live-session workbook",
	}
	cmd.AddCommand(a.editSheetAddCommand())
	return cmd
}

func (a *app) editSheetAddCommand() *cobra.Command {
	var name string
	var before string
	var after string
	var ifMissing bool
	var session bool

	cmd := &cobra.Command{
		Use:   "add [workbook]",
		Short: "Add a worksheet to a live-session workbook",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildEditSheetAddOptions(workbook, name, before, after, ifMissing, session, commandOpts)
			if err != nil {
				code := "edit_args_invalid"
				exitCode := output.ExitConfig
				if errors.Is(err, errInvalidWorksheetName) {
					code = "invalid_sheet_name"
					exitCode = output.ExitValidation
				}
				return a.writeFailure("edit", exitCode, code, err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Adding worksheet", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).EditSheetAdd(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&name, "name", "", "worksheet name to create")
	cmd.Flags().StringVar(&before, "before", "", "insert before this worksheet")
	cmd.Flags().StringVar(&after, "after", "", "insert after this worksheet")
	cmd.Flags().BoolVar(&ifMissing, "if-missing", false, "treat an existing worksheet with the same name as success")
	cmd.Flags().BoolVar(&session, "session", false, "require a matching active xlflow session workbook")
	return cmd
}

func (a *app) editCellCommand() *cobra.Command {
	var sheet string
	var cell string
	var value string
	var formula string
	var fill string
	var events string
	var session bool

	cmd := &cobra.Command{
		Use:   "cell [workbook]",
		Short: "Edit one live-session cell",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			var valuePtr *string
			if cmd.Flags().Changed("value") {
				valuePtr = &value
			}
			var formulaPtr *string
			if cmd.Flags().Changed("formula") {
				formulaPtr = &formula
			}
			opts, err := buildEditCellOptions(workbook, sheet, cell, fill, events, valuePtr, formulaPtr, session, commandOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook cell", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).EditCell(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&cell, "cell", "", "single cell address such as B2")
	cmd.Flags().StringVar(&value, "value", "", "set a text value")
	cmd.Flags().StringVar(&formula, "formula", "", "set a formula such as =A1+B1")
	cmd.Flags().StringVar(&fill, "fill", "", "set fill color using #RGB or #RRGGBB")
	cmd.Flags().StringVar(&events, "events", string(excel.EditEventKeep), "event mode for value/formula edits: keep, on, or off")
	cmd.Flags().BoolVar(&session, "session", false, "require a matching active xlflow session workbook")
	return cmd
}

func (a *app) editFormulaCommand() *cobra.Command {
	var sheet string
	var cellRange string
	var formula string
	var formulaR1C1 string
	var events string
	var calculate bool
	var session bool

	cmd := &cobra.Command{
		Use:   "formula [workbook]",
		Short: "Edit formulas in one live-session range",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			var formulaPtr *string
			if cmd.Flags().Changed("formula") {
				formulaPtr = &formula
			}
			var formulaR1C1Ptr *string
			if cmd.Flags().Changed("formula-r1c1") {
				formulaR1C1Ptr = &formulaR1C1
			}
			opts, err := buildEditFormulaOptions(workbook, sheet, cellRange, events, formulaPtr, formulaR1C1Ptr, calculate, session, commandOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook formulas", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).EditFormula(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&cellRange, "range", "", "range address such as D2:D1001")
	cmd.Flags().StringVar(&formula, "formula", "", "set an A1-style formula such as =B2*C2")
	cmd.Flags().StringVar(&formulaR1C1, "formula-r1c1", "", "set an R1C1-style formula such as =RC[-2]*RC[-1]")
	cmd.Flags().StringVar(&events, "events", string(excel.EditEventKeep), "event mode for formula edits: keep, on, or off")
	cmd.Flags().BoolVar(&calculate, "calculate", false, "calculate the target range after editing formulas")
	cmd.Flags().BoolVar(&session, "session", false, "require a matching active xlflow session workbook")
	return cmd
}

func (a *app) editRangeCommand() *cobra.Command {
	var sheet string
	var cellRange string
	var fill string
	var clear string
	var session bool

	cmd := &cobra.Command{
		Use:   "range [workbook]",
		Short: "Edit one live-session range",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildEditRangeOptions(workbook, sheet, cellRange, fill, clear, session, commandOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook range", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).EditRange(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&cellRange, "range", "", "range address such as A1:AE31")
	cmd.Flags().StringVar(&fill, "fill", "", "set fill color using #RGB or #RRGGBB")
	cmd.Flags().StringVar(&clear, "clear", "", "clear contents, formats, or all")
	cmd.Flags().BoolVar(&session, "session", false, "require a matching active xlflow session workbook")
	return cmd
}

func (a *app) editRowsCommand() *cobra.Command {
	var sheet string
	var rows string
	var height float64
	var session bool

	cmd := &cobra.Command{
		Use:   "rows [workbook]",
		Short: "Set row height on a live-session worksheet",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildEditRowsOptions(workbook, sheet, rows, height, session, commandOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook row heights", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).EditRows(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&rows, "rows", "", "row selector such as 1 or 1:31")
	cmd.Flags().Float64Var(&height, "height", 0, "row height in points")
	cmd.Flags().BoolVar(&session, "session", false, "require a matching active xlflow session workbook")
	return cmd
}

func (a *app) editColumnsCommand() *cobra.Command {
	var sheet string
	var columns string
	var width float64
	var session bool

	cmd := &cobra.Command{
		Use:   "columns [workbook]",
		Short: "Set column width on a live-session worksheet",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildEditColumnsOptions(workbook, sheet, columns, width, session, commandOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook column widths", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).EditColumns(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&columns, "columns", "", "column selector such as A or A:AE")
	cmd.Flags().Float64Var(&width, "width", 0, "column width in Excel character units")
	cmd.Flags().BoolVar(&session, "session", false, "require a matching active xlflow session workbook")
	return cmd
}

func (a *app) testCommand() *cobra.Command {
	var filter string
	var moduleFilter string
	var tagFilter string
	var isolation string
	var noSave bool
	var msgBoxLiterals []string
	var inputBoxLiterals []string
	var fileDialogLiterals []string
	var session bool
	var uiStream bool
	var failFast bool
	var maxFailures int
	var rerunFailed int
	cmd := &cobra.Command{
		Use:   "test",
		Short: "Run workbook VBA tests",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("test")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			runtime := excel.ResolveTestRuntimeInfo()
			msgBoxResponses, err := parseUIResponseLiterals("msgbox", msgBoxLiterals, normalizeMsgBoxResponseToken)
			if err != nil {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", err)
			}
			inputResponses, err := parseUIResponseLiterals("inputbox", inputBoxLiterals, nil)
			if err != nil {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", err)
			}
			fileDialogResponses, err := parseFileDialogResponseLiterals(fileDialogLiterals)
			if err != nil {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", err)
			}
			isolation = strings.ToLower(strings.TrimSpace(isolation))
			if isolation == "" {
				isolation = "none"
			}
			if isolation != "none" && isolation != "module" && isolation != "test" {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", fmt.Errorf("unsupported isolation mode %q; expected none, module, or test", isolation))
			}
			maxFailuresSet := cmd.Flags().Changed("max-failures")
			if failFast && maxFailuresSet {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", fmt.Errorf("--fail-fast is equivalent to --max-failures 1 and cannot be combined with --max-failures"))
			}
			if maxFailuresSet && maxFailures <= 0 {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", fmt.Errorf("--max-failures must be greater than zero"))
			}
			if rerunFailed < 0 {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", fmt.Errorf("--rerun-failed must be zero or greater"))
			}
			resolvedMaxFailures := maxFailures
			if failFast {
				resolvedMaxFailures = 1
			}
			if session && isolation != "none" {
				return a.writeFailure("test", output.ExitConfig, "unsupported_test_isolation", fmt.Errorf("isolation mode %q is not supported with --session", isolation))
			}
			if session && rerunFailed > 0 {
				return a.writeFailure("test", output.ExitConfig, "unsupported_test_rerun", fmt.Errorf("--rerun-failed cannot be combined with --session because retries require a fresh workbook baseline"))
			}
			err = a.withExcelProgress("Running VBA tests", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).TestWithOptions(cfg, filter, excel.TestOptions{Session: session, Isolation: isolation, NoSave: noSave, FailFast: failFast, MaxFailures: resolvedMaxFailures, RerunFailed: rerunFailed, Keepalive: commandOpts, RuntimeMode: runtime.Mode, RuntimeSource: runtime.Source, UIResponses: excel.UIResponses{MsgBox: msgBoxResponses, Input: inputResponses, FileDialog: fileDialogResponses}, DebugStream: excel.DebugStreamOptions{Enabled: true}, UIStream: excel.UIStreamOptions{Enabled: uiStream, RedactInput: true}, ModuleFilter: moduleFilter, TagFilter: tagFilter})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&filter, "filter", "", "run only the test whose qualified or unique procedure name exactly matches filter")
	cmd.Flags().StringVar(&moduleFilter, "module", "", "run only tests in the module whose name exactly matches filter")
	cmd.Flags().StringVar(&tagFilter, "tag", "", "run only tests tagged with the given tag")
	cmd.Flags().StringVar(&isolation, "isolation", "none", "workbook isolation mode: none, module, or test")
	cmd.Flags().BoolVar(&noSave, "no-save", false, "close the test workbook without explicitly saving changes")
	cmd.Flags().StringArrayVar(&msgBoxLiterals, "msgbox", nil, "provide a scripted MsgBox response as dialog-id=result")
	cmd.Flags().StringArrayVar(&inputBoxLiterals, "inputbox", nil, "provide a scripted InputBox response as dialog-id=value")
	cmd.Flags().StringArrayVar(&fileDialogLiterals, "filedialog", nil, "provide a scripted file dialog response as kind:dialog-id=path or kind:dialog-id=@cancel")
	cmd.Flags().BoolVar(&uiStream, "ui-stream", false, "stream headless XlflowUI dialog events to stderr in real time")
	cmd.Flags().BoolVar(&failFast, "fail-fast", false, "stop scheduling tests after the first final failure")
	cmd.Flags().IntVar(&maxFailures, "max-failures", 0, "stop scheduling tests after N final failures")
	cmd.Flags().IntVar(&rerunFailed, "rerun-failed", 0, "rerun failed tests up to N additional attempts")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	cmd.AddCommand(a.testListCommand())
	return cmd
}

func (a *app) testListCommand() *cobra.Command {
	var path string
	var module string
	cmd := &cobra.Command{
		Use:   "list",
		Short: "List source-defined VBA tests",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("test list")
			if err != nil {
				return err
			}
			result, err := testdiscover.Discover(testdiscover.Options{
				RootDir: a.cwd,
				Config:  cfg,
				Path:    path,
				Module:  module,
			})
			if err != nil {
				var duplicateErr testdiscover.DuplicateTestError
				if errors.As(err, &duplicateErr) {
					return a.writeFailure("test list", output.ExitValidation, "duplicate_test_name", err)
				}
				var metadataErr testdiscover.InvalidMetadataError
				if errors.As(err, &metadataErr) {
					return a.writeFailure("test list", output.ExitValidation, "invalid_test_metadata", err)
				}
				var testCaseErr testdiscover.InvalidTestCaseError
				if errors.As(err, &testCaseErr) {
					return a.writeFailure("test list", output.ExitValidation, "invalid_test_case", err)
				}
				return a.writeFailure("test list", output.ExitEnvironment, "test_discovery_failed", err)
			}
			env := output.New("test list")
			env.Target = map[string]any{
				"kind":        "source",
				"path":        result.Root,
				"description": "VBA source tests",
			}
			env.Tests = result
			env.Logs = []string{fmt.Sprintf("discovered %d VBA test(s) in %d source file(s)", result.Summary.Tests, result.Summary.Files)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&path, "path", "", "source directory or file to inspect (default: configured source tree)")
	cmd.Flags().StringVar(&module, "module", "", "list only tests in the module whose name exactly matches filter")
	return cmd
}

func (a *app) typeCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "type",
		Short: "Manage VBA type intelligence data",
	}
	cmd.AddCommand(a.typeDBCommand())
	return cmd
}

func (a *app) typeDBCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "db",
		Short: "Manage generated TypeLib type databases",
	}
	cmd.AddCommand(
		a.typeDBStatusCommand(),
		a.typeDBInitCommand(),
		a.typeDBRefreshCommand(),
		a.typeDBCleanCommand(),
	)
	return cmd
}

func (a *app) typeDBStatusCommand() *cobra.Command {
	var dir string
	cmd := &cobra.Command{
		Use:   "status",
		Short: "Show generated TypeLib type database status",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			status, err := typedb.StatusFor(typedb.Options{
				Dir:              dir,
				GeneratorVersion: a.buildInfo.withDefaults().Version,
			})
			if err != nil {
				return a.writeFailure("type db status", output.ExitEnvironment, "type_db_status_failed", err)
			}
			env := output.New("type db status")
			env.TypeDB = status
			if status.ManifestExists {
				env.Logs = []string{"generated type database status loaded"}
			} else {
				env.Logs = []string{"generated type database has not been initialized"}
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&dir, "dir", "", "override generated type database directory")
	return cmd
}

func (a *app) typeDBInitCommand() *cobra.Command {
	var dir string
	var libraries []string
	cmd := &cobra.Command{
		Use:   "init",
		Short: "Generate TypeLib type databases when missing",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			status, err := typedb.StatusFor(typedb.Options{Dir: dir, GeneratorVersion: a.buildInfo.withDefaults().Version})
			if err != nil {
				return a.writeFailure("type db init", output.ExitEnvironment, "type_db_status_failed", err)
			}
			if status.ManifestExists && !status.Stale {
				env := output.New("type db init")
				env.TypeDB = status
				env.Logs = []string{"generated type database already exists"}
				return a.write(env, output.ExitSuccess)
			}
			return a.generateTypeDB("type db init", dir, libraries)
		},
	}
	cmd.Flags().StringVar(&dir, "dir", "", "override generated type database directory")
	cmd.Flags().StringSliceVar(&libraries, "library", []string{"excel"}, "TypeLib library to import (repeat or comma-separate; use all for every known library present; default: excel)")
	return cmd
}

func (a *app) typeDBRefreshCommand() *cobra.Command {
	var dir string
	var libraries []string
	var force bool
	cmd := &cobra.Command{
		Use:   "refresh",
		Short: "Refresh generated TypeLib type databases",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			return a.generateTypeDB("type db refresh", dir, libraries)
		},
	}
	cmd.Flags().StringVar(&dir, "dir", "", "override generated type database directory")
	cmd.Flags().StringSliceVar(&libraries, "library", []string{"excel"}, "TypeLib library to import (repeat or comma-separate; use all for every known library present; default: excel)")
	cmd.Flags().BoolVar(&force, "force", false, "deprecated compatibility flag; refresh always regenerates")
	return cmd
}

func (a *app) generateTypeDB(command string, dir string, libraries []string) error {
	resolvedDir, err := typedb.ResolveDir(dir)
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "type_db_dir_failed", err)
	}
	env, code, err := a.excelRunner().TypeDBImport(excel.TypeDBImportOptions{
		OutputDir:        resolvedDir,
		GeneratorVersion: a.buildInfo.withDefaults().Version,
		Libraries:        libraries,
		Keepalive:        buildCommandOptions(a.stderrWriter()),
	})
	if err != nil {
		return err
	}
	env.Command = command
	return a.write(env, code)
}

func (a *app) typeDBCleanCommand() *cobra.Command {
	var dir string
	cmd := &cobra.Command{
		Use:   "clean",
		Short: "Delete generated TypeLib type databases",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cleaned, err := typedb.Clean(dir)
			if err != nil {
				return a.writeFailure("type db clean", output.ExitEnvironment, "type_db_clean_failed", err)
			}
			env := output.New("type db clean")
			env.TypeDB = map[string]any{"dir": cleaned}
			env.Logs = []string{"deleted generated type database directory"}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&dir, "dir", "", "override generated type database directory")
	return cmd
}

func (a *app) processCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "process",
		Short: "Manage local Excel processes",
	}
	cmd.AddCommand(a.processListCommand(), a.processCleanupCommand())
	return cmd
}

func (a *app) processListCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "list",
		Short: "List all local Excel processes",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			env, code, err := a.excelRunner().ProcessList(excel.ProcessListOptions{Action: "list"})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
}

func (a *app) processCleanupCommand() *cobra.Command {
	var auto bool
	var all bool
	var yes bool

	cmd := &cobra.Command{
		Use:   "cleanup [pid]",
		Short: "Terminate Excel processes",
		Long: `Terminate Excel processes by PID, by auto-detection (empty workbooks only), or by force-kill of all processes.

WARNING: cleanup --all forcibly terminates ALL Excel processes regardless of
unsaved workbooks or active work. Use with extreme caution.`,
		Args: cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			pid := ""
			if len(args) == 1 {
				pid = args[0]
			}
			if err := validateProcessCleanupArgs(pid, auto, all, yes); err != nil {
				return a.writeFailure("process cleanup", output.ExitConfig, "process_args_invalid", err)
			}
			if all && !yes {
				if a.json {
					return a.writeFailure("process cleanup", output.ExitConfig, "process_args_invalid", fmt.Errorf("--json with cleanup --all requires --yes for non-interactive safety"))
				}
				if !confirmPrompt(os.Stdin, a.stderrWriter(), "This will forcibly terminate ALL Excel processes. Unsaved work will be lost. Continue? [y/N] ") {
					env := output.New("process cleanup")
					env.Error = &output.Error{Code: "process_cancelled", Message: "cleanup --all cancelled by user"}
					env.Status = output.StatusFailed
					env.Logs = []string{"cleanup --all cancelled"}
					return a.write(env, output.ExitSuccess)
				}
			}
			opts := excel.ProcessCleanupOptions{Action: "cleanup", Auto: auto, All: all}
			if pid != "" {
				pidInt, err := strconv.Atoi(strings.TrimSpace(pid))
				if err != nil || pidInt <= 0 {
					return a.writeFailure("process cleanup", output.ExitConfig, "process_args_invalid", fmt.Errorf("PID must be a positive integer"))
				}
				opts.PID = pidInt
			}
			env, code, err := a.excelRunner().ProcessCleanup(opts)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&auto, "auto", false, "terminate only Excel processes with no open workbooks")
	cmd.Flags().BoolVar(&all, "all", false, "force-terminate ALL Excel processes (dangerous - prompts for confirmation)")
	cmd.Flags().BoolVar(&yes, "yes", false, "skip confirmation prompt for --all (use with caution)")
	return cmd
}

func validateProcessCleanupArgs(pid string, auto bool, all bool, yes bool) error {
	pid = strings.TrimSpace(pid)
	if yes && !all {
		return fmt.Errorf("--yes requires --all")
	}
	modeCount := 0
	if pid != "" {
		modeCount++
	}
	if auto {
		modeCount++
	}
	if all {
		modeCount++
	}
	if modeCount == 0 {
		return fmt.Errorf("process cleanup requires a PID, --auto, or --all")
	}
	if modeCount > 1 {
		return fmt.Errorf("PID, --auto, and --all cannot be combined")
	}
	if pid != "" {
		pidInt, err := strconv.Atoi(pid)
		if err != nil || pidInt <= 0 {
			return fmt.Errorf("PID must be a positive integer")
		}
	}
	return nil
}

func confirmPrompt(r io.Reader, w io.Writer, prompt string) bool {
	if _, err := fmt.Fprint(w, prompt); err != nil {
		return false
	}
	var response string
	_, err := fmt.Fscanln(r, &response)
	if err != nil {
		return false
	}
	response = strings.ToLower(strings.TrimSpace(response))
	return response == "y" || response == "yes"
}

func (a *app) diffCommand() *cobra.Command {
	var vbaBefore string
	var vbaAfter string
	cmd := &cobra.Command{
		Use:   "diff <before-workbook> <after-workbook>",
		Short: "Compare workbook state and exported VBA source",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			opts, err := buildDiffOptions(a.cwd, args[0], args[1], vbaBefore, vbaAfter)
			if err != nil {
				var unsupported workbookformat.UnsupportedError
				if errors.As(err, &unsupported) {
					return a.writeUnsupportedWorkbookFormat("diff", unsupported)
				}
				return a.writeFailure("diff", output.ExitConfig, "diff_args_invalid", err)
			}
			result, err := diff.Compare(opts)
			if err != nil {
				return a.writeFailure("diff", output.ExitEnvironment, "diff_failed", err)
			}
			env := output.New("diff")
			env.Diff = result
			env.Logs = result.Logs()
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&vbaBefore, "vba-before", "", "compare exported VBA source from this before directory")
	cmd.Flags().StringVar(&vbaAfter, "vba-after", "", "compare exported VBA source from this after directory")
	return cmd
}

type inspectSharedFlags struct {
	format string
}

func (a *app) inspectCommand() *cobra.Command {
	flags := &inspectSharedFlags{}
	cmd := &cobra.Command{
		Use:   "inspect",
		Short: "Inspect saved workbook state",
	}
	cmd.PersistentFlags().StringVar(&flags.format, "format", "text", "output format: text, json, or markdown")
	cmd.AddCommand(
		a.inspectWorkbookCommand(flags),
		a.inspectSheetsCommand(flags),
		a.inspectFormCommand(flags),
		a.inspectRangeCommand(flags),
		a.inspectUsedRangeCommand(flags),
		a.inspectCellCommand(flags),
		a.inspectCallsCommand(flags),
		a.inspectSymbolsCommand(flags),
	)
	return cmd
}

func (a *app) inspectCallsCommand(flags *inspectSharedFlags) *cobra.Command {
	var path string
	var from string
	var to string
	var includeMembers bool
	var includeBuiltins bool
	cmd := &cobra.Command{
		Use:   "calls",
		Short: "Inspect VBA source call sites",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			result, err := calls.Inspect(calls.Options{
				RootDir: a.cwd,
				Config:  cfg,
				Path:    path,
				From:    from,
				To:      to,
			})
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			env := output.New("inspect")
			env.Target = map[string]any{
				"kind":        "source",
				"path":        result.Root,
				"description": "VBA source files",
			}
			env.Inspect = workbookinspect.Payload{
				Target:  "calls",
				Format:  format,
				Source:  "tree_sitter_vba",
				Root:    result.Root,
				Calls:   result.Calls,
				Summary: result.Summary,
			}
			env.Logs = []string{fmt.Sprintf("inspected %d VBA source file(s), found %d call site(s)", result.Summary.Files, result.Summary.Calls)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&path, "path", "", "source directory or file to inspect (default: configured source tree)")
	cmd.Flags().StringVar(&from, "from", "", "show only calls made from a module or procedure")
	cmd.Flags().StringVar(&to, "to", "", "show only call sites whose callee matches this name")
	cmd.Flags().BoolVar(&includeMembers, "include-members", false, "compatibility no-op; member calls are included by default")
	cmd.Flags().BoolVar(&includeBuiltins, "include-builtins", false, "compatibility no-op; built-in-looking calls are included by default")
	return cmd
}

func (a *app) inspectSymbolsCommand(flags *inspectSharedFlags) *cobra.Command {
	var path string
	var includePrivate bool
	var includeLabels bool
	var module string
	cmd := &cobra.Command{
		Use:   "symbols",
		Short: "Inspect VBA source symbols",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			result, err := symbols.Inspect(symbols.Options{
				RootDir:        a.cwd,
				Config:         cfg,
				Path:           path,
				IncludePrivate: includePrivate,
				IncludeLabels:  includeLabels,
				Module:         module,
			})
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			env := output.New("inspect")
			env.Target = map[string]any{
				"kind":        "source",
				"path":        result.Root,
				"description": "VBA source files",
			}
			env.Inspect = workbookinspect.Payload{
				Target:  "symbols",
				Format:  format,
				Source:  "tree_sitter_vba",
				Root:    result.Root,
				Files:   result.Files,
				Summary: result.Summary,
			}
			env.Logs = []string{fmt.Sprintf("inspected %d VBA source file(s), found %d symbol(s)", result.Summary.Files, result.Summary.Symbols)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&path, "path", "", "source directory or file to inspect (default: configured source tree)")
	cmd.Flags().BoolVar(&includePrivate, "include-private", false, "include private and local symbols")
	cmd.Flags().BoolVar(&includeLabels, "include-labels", false, "include labels and line-number labels")
	cmd.Flags().StringVar(&module, "module", "", "inspect only one module")
	return cmd
}

func (a *app) inspectWorkbookCommand(flags *inspectSharedFlags) *cobra.Command {
	var session bool
	cmd := &cobra.Command{
		Use:   "workbook",
		Short: "Inspect workbook summary information",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			bridgeOpts, err := buildInspectBridgeOptions(session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			if bridgeOpts.Session {
				env, code, err := a.runSessionInspect(cfg, excel.InspectOptions{
					Target:    "workbook",
					Session:   true,
					Keepalive: bridgeOpts.Keepalive,
				}, format)
				if err != nil {
					return err
				}
				return a.write(env, code)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			if err := a.rejectUnsupportedFileInspectWorkbook("inspect", "inspect workbook", workbookPath); err != nil {
				return err
			}
			workbook, err := workbookinspect.Workbook(workbookPath)
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			target, sessionState, warnings := a.inspectStateForWorkbook(cfg, workbook.Path)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env := output.New("inspect")
			env.Target = target
			env.Session = sessionState
			env.Warnings = append(warnings, formWarnings...)
			hints := append([]map[string]any{}, formHints...)
			if len(warnings) > 0 {
				hints = append(hints, staleFileInspectHints("workbook")...)
			}
			env.Hints = hints
			env.Inspect = workbookinspect.Payload{
				Target:     "workbook",
				TargetInfo: workbookinspect.SavedFileTargetInfo(workbook.Path),
				Format:     format,
				Source:     "file",
				Workbook:   &workbook,
			}
			env.Logs = []string{fmt.Sprintf("inspected workbook %s", workbook.Path)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "inspect the live workbook through the managed xlflow session")
	return cmd
}

func (a *app) inspectSheetsCommand(flags *inspectSharedFlags) *cobra.Command {
	var session bool
	cmd := &cobra.Command{
		Use:   "sheets",
		Short: "Inspect workbook worksheets",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			bridgeOpts, err := buildInspectBridgeOptions(session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			if bridgeOpts.Session {
				env, code, err := a.runSessionInspect(cfg, excel.InspectOptions{
					Target:    "sheets",
					Session:   true,
					Keepalive: bridgeOpts.Keepalive,
				}, format)
				if err != nil {
					return err
				}
				return a.write(env, code)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			if err := a.rejectUnsupportedFileInspectWorkbook("inspect", "inspect sheets", workbookPath); err != nil {
				return err
			}
			sheets, err := workbookinspect.Sheets(workbookPath)
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			target, sessionState, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env := output.New("inspect")
			env.Target = target
			env.Session = sessionState
			env.Warnings = append(warnings, formWarnings...)
			hints := append([]map[string]any{}, formHints...)
			if len(warnings) > 0 {
				hints = append(hints, staleFileInspectHints("sheets")...)
			}
			env.Hints = hints
			env.Inspect = workbookinspect.Payload{
				Target:     "sheets",
				TargetInfo: workbookinspect.SavedFileTargetInfo(workbookPath),
				Format:     format,
				Source:     "file",
				Sheets:     sheets,
			}
			env.Logs = []string{fmt.Sprintf("inspected %d worksheet(s)", len(sheets))}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "inspect the live workbook through the managed xlflow session")
	return cmd
}

func (a *app) inspectFormCommand(flags *inspectSharedFlags) *cobra.Command {
	var session bool
	var runtime bool
	var designer bool
	var both bool
	var initializer string
	cmd := &cobra.Command{
		Use:   "form <name>",
		Short: "Inspect a workbook UserForm through Excel COM",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_form_args_invalid", err)
			}
			basis, err := buildInspectFormBasis(runtime, designer, both)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_form_args_invalid", err)
			}
			opts, err := buildInspectFormOptions(args[0], basis, initializer, session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_form_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			var scriptEnv output.Envelope
			var code int
			err = a.withExcelProgress("Inspecting workbook form", opts.Keepalive, func() error {
				var runErr error
				scriptEnv, code, runErr = a.excelRunnerForConfig(cfg).InspectForm(cfg, opts)
				return runErr
			})
			if err != nil {
				return err
			}
			if code != output.ExitSuccess {
				return a.write(scriptEnv, code)
			}
			env := output.New("inspect")
			env.Target = scriptEnv.Target
			env.Session = scriptEnv.Session
			env.Workbook = scriptEnv.Workbook
			env.Warnings = scriptEnv.Warnings
			env.Hints = scriptEnv.Hints
			payload := workbookinspect.Payload{
				Target: "form",
				Format: format,
				Source: "excel_com",
			}
			if opts.Basis == "both" {
				payload.Forms = scriptEnv.Forms
			} else {
				payload.Form = scriptEnv.Forms
			}
			env.Inspect = payload
			env.Logs = []string{fmt.Sprintf("inspected %s UserForm %s", opts.Basis, opts.Name)}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&runtime, "runtime", false, "inspect the loaded runtime form state")
	cmd.Flags().BoolVar(&designer, "designer", false, "inspect the design-time VBIDE designer state")
	cmd.Flags().BoolVar(&both, "both", false, "inspect both runtime and designer state")
	cmd.Flags().StringVar(&initializer, "initializer", "", "optional public form method invoked with ThisWorkbook during runtime inspection")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
}

func (a *app) inspectRangeCommand(flags *inspectSharedFlags) *cobra.Command {
	var sheet string
	var address string
	var maxRows int
	var maxCols int
	var includeStyle bool
	var session bool
	cmd := &cobra.Command{
		Use:   "range [<sheet!A1:B2>]",
		Short: "Inspect a worksheet range",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			limits, err := buildInspectLimits(maxRows, maxCols)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			selector, err := buildInspectCellSelector(args, sheet, address, true)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			bridgeOpts, err := buildInspectBridgeOptions(session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			if bridgeOpts.Session {
				env, code, err := a.runSessionInspect(cfg, excel.InspectOptions{
					Target:       "range",
					Sheet:        selector.Sheet,
					Address:      selector.Address,
					Limits:       inspectLimitsMap(limits),
					IncludeStyle: includeStyle,
					Session:      true,
					Keepalive:    bridgeOpts.Keepalive,
				}, format)
				if err != nil {
					return err
				}
				return a.write(env, code)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			if err := a.rejectUnsupportedFileInspectWorkbook("inspect", "inspect range", workbookPath); err != nil {
				return err
			}
			snapshot, err := workbookinspect.Range(workbookPath, selector.Sheet, selector.Address, workbookinspect.RangeOptions{
				Limits:       limits,
				IncludeStyle: includeStyle,
			})
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			env := output.New("inspect")
			target, sessionState, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env.Target = target
			env.Session = sessionState
			env.Warnings = append(warnings, formWarnings...)
			hints := append([]map[string]any{}, formHints...)
			if len(warnings) > 0 {
				hints = append(hints, staleFileInspectHints("range", selector.Sheet, selector.Address)...)
			}
			env.Hints = hints
			env.Inspect = workbookinspect.Payload{
				Target:     "range",
				TargetInfo: workbookinspect.SavedFileTargetInfo(workbookPath),
				Format:     format,
				Source:     "file",
				Range:      &snapshot,
			}
			env.Logs = []string{fmt.Sprintf("inspected range %s!%s", selector.Sheet, snapshot.Range)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&address, "address", "", "range address such as A1:F20")
	cmd.Flags().IntVar(&maxRows, "max-rows", 100, "maximum rows returned")
	cmd.Flags().IntVar(&maxCols, "max-cols", 30, "maximum columns returned")
	cmd.Flags().BoolVar(&includeStyle, "include-style", false, "include style, row, column, and merge metadata in the result")
	cmd.Flags().BoolVar(&session, "session", false, "inspect the live workbook through the managed xlflow session")
	return cmd
}

func buildInspectFormBasis(runtime, designer, both bool) (string, error) {
	selected := 0
	if runtime {
		selected++
	}
	if designer {
		selected++
	}
	if both {
		selected++
	}
	if selected == 0 {
		return "runtime", nil
	}
	if selected > 1 {
		return "", fmt.Errorf("choose only one of --runtime, --designer, or --both")
	}
	switch {
	case runtime:
		return "runtime", nil
	case designer:
		return "designer", nil
	default:
		return "both", nil
	}
}

func buildInspectFormOptions(name, basis, initializer string, session bool, commandOpts excel.CommandOptions) (excel.InspectFormOptions, error) {
	if strings.TrimSpace(name) == "" {
		return excel.InspectFormOptions{}, fmt.Errorf("form name is required")
	}
	trimmedBasis := strings.ToLower(strings.TrimSpace(basis))
	if trimmedBasis != "runtime" && trimmedBasis != "designer" && trimmedBasis != "both" {
		return excel.InspectFormOptions{}, fmt.Errorf("unsupported inspect form basis %q", basis)
	}
	trimmedInitializer := strings.TrimSpace(initializer)
	if trimmedInitializer != "" && trimmedBasis == "designer" {
		return excel.InspectFormOptions{}, fmt.Errorf("--initializer can only be used with runtime or both inspection")
	}
	return excel.InspectFormOptions{
		Name:        strings.TrimSpace(name),
		Basis:       trimmedBasis,
		Initializer: trimmedInitializer,
		Session:     session,
		Keepalive:   commandOpts,
	}, nil
}

func buildInspectBridgeOptions(session bool, commandOpts excel.CommandOptions) (inspectBridgeOptions, error) {
	return inspectBridgeOptions{
		Session:   session,
		Keepalive: commandOpts,
	}, nil
}

func inspectLimitsMap(limits workbookinspect.Limits) map[string]int {
	return map[string]int{
		"max_rows": limits.MaxRows,
		"max_cols": limits.MaxCols,
	}
}

func staleFileInspectHints(target string, args ...string) []map[string]any {
	command := "xlflow inspect " + target
	switch target {
	case "range", "cell":
		if len(args) >= 2 {
			command += " --sheet " + strconv.Quote(args[0]) + " --address " + strconv.Quote(args[1])
		}
	case "used-range":
		if len(args) >= 1 && strings.TrimSpace(args[0]) != "" {
			command += " --sheet " + strconv.Quote(args[0])
		}
	}
	return []map[string]any{
		{
			"code":    "inspect_live_session",
			"message": command + " --session --json reads the live workbook currently open in Excel.",
		},
		{
			"code":    "save_session_before_file_inspect",
			"message": "Run `xlflow save --session --json` before disk-backed inspect when the live workbook is newer than disk.",
		},
	}
}

func buildFormSnapshotOptions(name, outPath string, session bool, commandOpts excel.CommandOptions) (formSnapshotCommandOptions, error) {
	inspectOpts, err := buildInspectFormOptions(name, "designer", "", session, commandOpts)
	if err != nil {
		return formSnapshotCommandOptions{}, err
	}
	trimmedOut := strings.TrimSpace(outPath)
	if trimmedOut == "" {
		return formSnapshotCommandOptions{}, fmt.Errorf("--out is required")
	}
	return formSnapshotCommandOptions{
		Inspect: inspectOpts,
		OutPath: trimmedOut,
	}, nil
}

func buildFormWriteOptions(action, specPath string, overwrite, session, noSave bool, commandOpts excel.CommandOptions, root string) (formWriteCommandOptions, error) {
	normalizedAction := strings.ToLower(strings.TrimSpace(action))
	if normalizedAction != "build" && normalizedAction != "apply" {
		return formWriteCommandOptions{}, fmt.Errorf("unsupported form action %q", action)
	}
	if noSave && !session {
		return formWriteCommandOptions{}, fmt.Errorf("--no-save requires --session")
	}
	specInput, err := forms.ResolveSpecInput(root, specPath)
	if err != nil {
		return formWriteCommandOptions{}, err
	}
	spec, err := forms.LoadFormSpec(specInput)
	if err != nil {
		return formWriteCommandOptions{}, err
	}
	return formWriteCommandOptions{
		Action:    normalizedAction,
		SpecInput: specInput,
		Spec:      spec,
		Overwrite: overwrite,
		Session:   session,
		NoSave:    noSave,
		Keepalive: commandOpts,
	}, nil
}

func buildFormExportImageOptions(name, outPath, initializer string, overwrite bool, session bool, commandOpts excel.CommandOptions) (excel.FormExportImageOptions, error) {
	if strings.TrimSpace(name) == "" {
		return excel.FormExportImageOptions{}, fmt.Errorf("form name is required")
	}
	trimmedOut := strings.TrimSpace(outPath)
	if trimmedOut == "" {
		return excel.FormExportImageOptions{}, fmt.Errorf("--out is required")
	}
	return excel.FormExportImageOptions{
		Name:        strings.TrimSpace(name),
		OutPath:     trimmedOut,
		Initializer: strings.TrimSpace(initializer),
		Overwrite:   overwrite,
		Session:     session,
		Keepalive:   commandOpts,
	}, nil
}

func (a *app) inspectUsedRangeCommand(flags *inspectSharedFlags) *cobra.Command {
	var sheet string
	var maxRows int
	var maxCols int
	var includeStyle bool
	var session bool
	cmd := &cobra.Command{
		Use:   "used-range [<sheet>]",
		Short: "Inspect the lightweight used range for a worksheet",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			limits, err := buildInspectLimits(maxRows, maxCols)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			targetSheet, err := buildInspectSheetSelector(args, sheet)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			bridgeOpts, err := buildInspectBridgeOptions(session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			if bridgeOpts.Session {
				env, code, err := a.runSessionInspect(cfg, excel.InspectOptions{
					Target:       "used-range",
					Sheet:        targetSheet,
					Limits:       inspectLimitsMap(limits),
					IncludeStyle: includeStyle,
					Session:      true,
					Keepalive:    bridgeOpts.Keepalive,
				}, format)
				if err != nil {
					return err
				}
				return a.write(env, code)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			if err := a.rejectUnsupportedFileInspectWorkbook("inspect", "inspect used-range", workbookPath); err != nil {
				return err
			}
			snapshot, err := workbookinspect.UsedRange(workbookPath, targetSheet, workbookinspect.RangeOptions{
				Limits:       limits,
				IncludeStyle: includeStyle,
			})
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			env := output.New("inspect")
			target, sessionState, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env.Target = target
			env.Session = sessionState
			env.Warnings = append(warnings, formWarnings...)
			hints := append([]map[string]any{}, formHints...)
			if len(warnings) > 0 {
				hints = append(hints, staleFileInspectHints("used-range", targetSheet)...)
			}
			env.Hints = hints
			env.Inspect = workbookinspect.Payload{
				Target:     "used-range",
				TargetInfo: workbookinspect.SavedFileTargetInfo(workbookPath),
				Format:     format,
				Source:     "file",
				Range:      &snapshot,
			}
			env.Logs = []string{fmt.Sprintf("inspected used range for %s", targetSheet)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().IntVar(&maxRows, "max-rows", 100, "maximum rows returned")
	cmd.Flags().IntVar(&maxCols, "max-cols", 30, "maximum columns returned")
	cmd.Flags().BoolVar(&includeStyle, "include-style", false, "include style, row, column, and merge metadata in the result")
	cmd.Flags().BoolVar(&session, "session", false, "inspect the live workbook through the managed xlflow session")
	return cmd
}

func (a *app) inspectCellCommand(flags *inspectSharedFlags) *cobra.Command {
	var sheet string
	var address string
	var session bool
	cmd := &cobra.Command{
		Use:   "cell [<sheet!A1>]",
		Short: "Inspect a single worksheet cell",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			format, err := validateInspectFormat(flags.format)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			selector, err := buildInspectCellSelector(args, sheet, address, false)
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			bridgeOpts, err := buildInspectBridgeOptions(session, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return a.writeFailure("inspect", output.ExitConfig, "inspect_args_invalid", err)
			}
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			if bridgeOpts.Session {
				env, code, err := a.runSessionInspect(cfg, excel.InspectOptions{
					Target:    "cell",
					Sheet:     selector.Sheet,
					Address:   selector.Address,
					Session:   true,
					Keepalive: bridgeOpts.Keepalive,
				}, format)
				if err != nil {
					return err
				}
				return a.write(env, code)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			if err := a.rejectUnsupportedFileInspectWorkbook("inspect", "inspect cell", workbookPath); err != nil {
				return err
			}
			cell, err := workbookinspect.Cell(workbookPath, selector.Sheet, selector.Address)
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			target, sessionState, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env := output.New("inspect")
			env.Target = target
			env.Session = sessionState
			env.Warnings = append(warnings, formWarnings...)
			hints := append([]map[string]any{}, formHints...)
			if len(warnings) > 0 {
				hints = append(hints, staleFileInspectHints("cell", selector.Sheet, selector.Address)...)
			}
			env.Hints = hints
			env.Inspect = workbookinspect.Payload{
				Target:     "cell",
				TargetInfo: workbookinspect.SavedFileTargetInfo(workbookPath),
				Format:     format,
				Source:     "file",
				Cell:       &cell,
			}
			env.Logs = []string{fmt.Sprintf("inspected cell %s!%s", selector.Sheet, selector.Address)}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&sheet, "sheet", "", "worksheet name")
	cmd.Flags().StringVar(&address, "address", "", "cell address such as A1")
	cmd.Flags().BoolVar(&session, "session", false, "inspect the live workbook through the managed xlflow session")
	return cmd
}

func (a *app) runSessionInspect(cfg config.Config, opts excel.InspectOptions, format string) (output.Envelope, int, error) {
	var scriptEnv output.Envelope
	var code int
	err := a.withExcelProgress("Inspecting live workbook", opts.Keepalive, func() error {
		var runErr error
		scriptEnv, code, runErr = a.excelRunnerForConfig(cfg).Inspect(cfg, opts)
		return runErr
	})
	if err != nil {
		return output.Envelope{}, 0, err
	}
	if code != output.ExitSuccess {
		return scriptEnv, code, nil
	}
	env := output.New("inspect")
	env.Target = scriptEnv.Target
	env.Session = scriptEnv.Session
	env.Workbook = scriptEnv.Workbook
	env.Warnings = scriptEnv.Warnings
	env.Hints = scriptEnv.Hints
	env.Logs = scriptEnv.Logs
	payload := cliObjectMap(scriptEnv.Inspect)
	if len(payload) > 0 {
		payload["format"] = format
		env.Inspect = payload
	}
	return env, output.ExitSuccess, nil
}

func (a *app) inspectStateForWorkbook(cfg config.Config, workbookPath string) (map[string]any, map[string]any, []map[string]any) {
	target := map[string]any{
		"kind":        "file",
		"path":        workbookPath,
		"description": "Saved workbook file on disk",
	}
	session := map[string]any{
		"active":               false,
		"workbook_path":        workbookPath,
		"workbook_name":        filepath.Base(workbookPath),
		"dirty":                false,
		"save_required":        false,
		"live_newer_than_disk": false,
		"source_of_truth":      "saved_workbook",
	}
	if runtime.GOOS != "windows" {
		return target, session, nil
	}
	status, ok := a.inspectSessionStatus(cfg)
	if !ok {
		return target, session, nil
	}
	statusWorkbookPath := stringValueForCLI(status, "workbook_path")
	if strings.TrimSpace(statusWorkbookPath) == "" || !samePath(statusWorkbookPath, workbookPath) {
		return target, session, nil
	}
	active := boolValueForCLI(status, "active") || (boolValueForCLI(status, "running") && boolValueForCLI(status, "workbook_open"))
	saveRequired := boolValueForCLI(status, "save_required") || boolValueForCLI(status, "needs_save")
	session["active"] = active
	if rawDirty, exists := status["dirty"]; exists {
		session["dirty"] = rawDirty
	}
	session["save_required"] = saveRequired
	session["live_newer_than_disk"] = saveRequired
	if saveRequired {
		session["source_of_truth"] = "live_workbook"
	}
	if mode := stringValueForCLI(status, "mode"); mode != "" {
		session["mode"] = mode
	}
	if name := stringValueForCLI(status, "workbook_name"); strings.TrimSpace(name) != "" {
		session["workbook_name"] = name
	}
	if present, ok := status["userforms_present"]; ok {
		session["userforms_present"] = present
	}
	if count, ok := status["userform_count"]; ok {
		session["userform_count"] = count
	}
	if known, ok := status["userforms_known"]; ok {
		session["userforms_known"] = known
	}
	if !active || !saveRequired {
		return target, session, nil
	}
	warnings := []map[string]any{
		{
			"code":    "live_session_dirty",
			"message": "A live session exists and has unsaved changes. This command inspected the saved workbook file, so the result may be stale until `xlflow save --session` persists the live workbook.",
		},
		{
			"code":    "command_reads_saved_file",
			"message": "This inspect command read the saved workbook file on disk, not the live workbook in Excel.",
		},
	}
	if knownValue, ok := session["userforms_known"]; ok && knownValue != nil && !boolValueForCLI(session, "userforms_known") {
		warnings = append(warnings, map[string]any{
			"code":    "userform_detection_unavailable",
			"message": "xlflow could not determine whether the live workbook contains UserForms. Save before relying on disk-backed inspect, pull, or source review for form state.",
		})
	}
	return target, session, warnings
}

func (a *app) inspectSessionStatus(cfg config.Config) (map[string]any, bool) {
	env, _, err := a.excelRunnerForConfig(cfg).Session(cfg, "status")
	if err == nil {
		if status := cliObjectMap(env.Session); len(status) > 0 {
			return status, true
		}
	}

	exe := resolvedExecutablePath()
	if strings.TrimSpace(exe) == "" || !strings.EqualFold(filepath.Ext(exe), ".exe") {
		return nil, false
	}
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	cmd := exec.CommandContext(ctx, exe, "--json", "session", "status")
	cmd.Dir = a.cwd
	var stdout bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = io.Discard
	if err := cmd.Run(); err != nil {
		return nil, false
	}
	var envOut output.Envelope
	if err := json.Unmarshal(stdout.Bytes(), &envOut); err != nil {
		return nil, false
	}
	status := cliObjectMap(envOut.Session)
	if len(status) == 0 {
		return nil, false
	}
	return status, true
}

func inspectSourceUserFormMessages(root string, cfg config.Config) ([]map[string]any, []map[string]any) {
	names := collectSourceUserFormNames(workbookArgPath(root, cfg.Src.Forms))
	if len(names) == 0 {
		return nil, nil
	}
	warnings := []map[string]any{{
		"code":    "userform_inspect_saved_file",
		"message": fmt.Sprintf("UserForms detected in source (%s). File-based inspect reads the saved workbook and cannot verify live UserForm Designer/runtime state from `.frm` text alone.", strings.Join(names, ", ")),
	}}
	hints := []map[string]any{{
		"code":    "userform_planned_commands",
		"message": "UserForm workflow: `xlflow inspect form <name> --designer --json`, `xlflow form snapshot <name> --out src/forms/specs/<name>.yaml`, edit the spec, then `xlflow form build src/forms/specs/<name>.yaml --overwrite` and verify with `xlflow form export-image <name> --out <path>`.",
	}}
	return warnings, hints
}

func collectSourceUserFormNames(formsDir string) []string {
	if strings.TrimSpace(formsDir) == "" {
		return nil
	}
	names := map[string]struct{}{}
	_ = filepath.WalkDir(formsDir, func(path string, d os.DirEntry, err error) error {
		if err != nil || d == nil || d.IsDir() {
			return nil
		}
		if !strings.EqualFold(filepath.Ext(d.Name()), ".frm") {
			return nil
		}
		name := strings.TrimSpace(strings.TrimSuffix(d.Name(), filepath.Ext(d.Name())))
		if name != "" {
			names[name] = struct{}{}
		}
		return nil
	})
	if len(names) == 0 {
		return nil
	}
	result := make([]string, 0, len(names))
	for name := range names {
		result = append(result, name)
	}
	slices.Sort(result)
	return result
}

func (a *app) prepareUserFormSidecarMigration(cfg config.Config, target string, overwrite bool, session bool, commandOpts excel.CommandOptions) ([]formMigrationFile, error) {
	formsDir := workbookArgPath(a.cwd, cfg.Src.Forms)
	candidates, err := collectUserFormMigrationCandidates(a.cwd, formsDir, target)
	if err != nil {
		return nil, err
	}
	if len(candidates) == 0 {
		if target != "" {
			return nil, fmt.Errorf("%w: UserForm %q was not found under %s", errFormMigrateArgs, target, displayPath(a.cwd, formsDir))
		}
		return nil, fmt.Errorf("%w: no UserForm .frm files found under %s", errFormMigrateArgs, displayPath(a.cwd, formsDir))
	}
	for i := range candidates {
		body, err := os.ReadFile(candidates[i].FRMPath)
		if err != nil {
			return nil, fmt.Errorf("read %s: %w", candidates[i].FRMPath, err)
		}
		code := forms.NormalizeUserFormCodeText(forms.ExtractUserFormCodeFromFRM(string(body)))
		if firstAttributeLine(code) > 0 {
			return nil, fmt.Errorf("%w: extracted sidecar code for %s contains Attribute VB_* header lines", errFormMigrateConflict, candidates[i].Name)
		}
		candidates[i].Code = code
		if err := validateFormMigrationConflicts(candidates[i], overwrite); err != nil {
			return nil, err
		}
	}
	for i := range candidates {
		spec, err := a.inspectUserFormSpecForMigration(cfg, candidates[i].Name, session, commandOpts)
		if err != nil {
			return nil, err
		}
		if !strings.EqualFold(spec.Form.Name, candidates[i].Name) {
			return nil, fmt.Errorf("UserForm spec name %q does not match .frm basename %q", spec.Form.Name, candidates[i].Name)
		}
		vbName, err := userFormVBNameFromFRM(candidates[i].FRMPath)
		if err != nil {
			return nil, err
		}
		if !strings.EqualFold(vbName, candidates[i].Name) {
			return nil, fmt.Errorf("UserForm artifact %s declares Attribute VB_Name = %q, want %q", candidates[i].FRMPath, vbName, candidates[i].Name)
		}
		candidates[i].Spec = spec
	}
	return candidates, nil
}

func (a *app) rejectStaleSourceForFormMigration(cfg config.Config) error {
	workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
	state := buildStatusState(a.cwd, cfg, workbookPath)
	if !boolValueForCLI(state, "src_newer_than_workbook") {
		return nil
	}
	message := "source files are newer than the configured workbook; run xlflow push or xlflow pull before migrating UserForms to sidecar mode"
	if latest := stringValueForCLI(state, "latest_source_modified_at"); latest != "" {
		message += "; latest source modified at " + latest
	}
	if workbook := stringValueForCLI(state, "workbook_last_modified_at"); workbook != "" {
		message += "; workbook modified at " + workbook
	}
	return fmt.Errorf("%w: %s", errFormMigrateConflict, message)
}

func collectUserFormMigrationCandidates(root string, formsDir string, target string) ([]formMigrationFile, error) {
	if strings.TrimSpace(formsDir) == "" {
		return nil, nil
	}
	if _, err := os.Stat(formsDir); err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return nil, nil
		}
		return nil, err
	}
	codeDir := filepath.Join(formsDir, "code")
	specDir := filepath.Join(formsDir, "specs")
	var out []formMigrationFile
	seen := map[string]string{}
	err := filepath.WalkDir(formsDir, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			if sameCLIPath(path, codeDir) || sameCLIPath(path, specDir) {
				return filepath.SkipDir
			}
			return nil
		}
		if !strings.EqualFold(filepath.Ext(d.Name()), ".frm") {
			return nil
		}
		name := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		if target != "" && !strings.EqualFold(name, target) {
			return nil
		}
		normalizedName := strings.ToLower(name)
		if first, ok := seen[normalizedName]; ok {
			return fmt.Errorf("%w: duplicate UserForm basename %q at %s and %s", errFormMigrateConflict, name, first, path)
		}
		seen[normalizedName] = path
		frxPath := filepath.Join(filepath.Dir(path), name+".frx")
		if _, err := os.Stat(frxPath); err != nil {
			if errors.Is(err, os.ErrNotExist) {
				frxPath = ""
			} else {
				return err
			}
		}
		out = append(out, formMigrationFile{
			Name:     name,
			FRMPath:  path,
			FRXPath:  frxPath,
			CodePath: filepath.Join(formsDir, "code", name+".bas"),
			SpecPath: filepath.Join(formsDir, "specs", name+".yaml"),
		})
		return nil
	})
	if err != nil {
		return nil, err
	}
	slices.SortFunc(out, func(a, b formMigrationFile) int {
		return cmp.Compare(strings.ToLower(a.Name), strings.ToLower(b.Name))
	})
	return out, nil
}

func validateFormMigrationConflicts(item formMigrationFile, overwrite bool) error {
	if body, err := os.ReadFile(item.CodePath); err == nil {
		if forms.NormalizeUserFormCodeText(string(body)) != item.Code && !overwrite {
			return fmt.Errorf("%w: refusing to overwrite existing sidecar code %s", errFormMigrateConflict, item.CodePath)
		}
	} else if !errors.Is(err, os.ErrNotExist) {
		return err
	}
	if _, err := os.Stat(item.SpecPath); err == nil {
		if !overwrite {
			return fmt.Errorf("%w: refusing to overwrite existing Designer spec %s", errFormMigrateConflict, item.SpecPath)
		}
	} else if !errors.Is(err, os.ErrNotExist) {
		return err
	}
	return nil
}

func (a *app) inspectUserFormSpecForMigration(cfg config.Config, name string, session bool, commandOpts excel.CommandOptions) (forms.FormSpec, error) {
	opts, err := buildInspectFormOptions(name, "designer", "", session, commandOpts)
	if err != nil {
		return forms.FormSpec{}, err
	}
	var env output.Envelope
	var code int
	err = a.withExcelProgress("Inspecting workbook form", opts.Keepalive, func() error {
		var runErr error
		env, code, runErr = a.excelRunnerForConfig(cfg).InspectForm(cfg, opts)
		return runErr
	})
	if err != nil {
		return forms.FormSpec{}, fmt.Errorf("%w: %v", errFormMigrateInspect, err)
	}
	if code != output.ExitSuccess {
		message := "inspect form failed"
		if env.Error != nil && strings.TrimSpace(env.Error.Message) != "" {
			message = env.Error.Message
		}
		return forms.FormSpec{}, fmt.Errorf("%w: %s", errFormMigrateInspect, message)
	}
	spec, err := forms.FormSpecFromInspectSnapshot(env.Forms)
	if err != nil {
		return forms.FormSpec{}, fmt.Errorf("%w: %v", errFormMigrateInspect, err)
	}
	return spec, nil
}

func (a *app) writeUserFormSidecarMigration(before string, items []formMigrationFile, overwrite bool) ([]string, []string, []string, error) {
	created := make([]string, 0)
	updated := make([]string, 0)
	skipped := make([]string, 0)
	rollback := make([]formMigrationRollback, 0)
	fail := func(err error) ([]string, []string, []string, error) {
		if rollbackErr := rollbackFormMigrationWrites(rollback); rollbackErr != nil {
			err = fmt.Errorf("%w; rollback failed: %v", err, rollbackErr)
		}
		return nil, nil, nil, err
	}
	for _, item := range items {
		if err := os.MkdirAll(filepath.Dir(item.CodePath), 0o755); err != nil {
			return fail(err)
		}
		if body, err := os.ReadFile(item.CodePath); err == nil && forms.NormalizeUserFormCodeText(string(body)) == item.Code {
			skipped = append(skipped, displayPath(a.cwd, item.CodePath))
		} else {
			existed := false
			var oldBody []byte
			if _, err := os.Stat(item.CodePath); err == nil {
				existed = true
				oldBody, err = os.ReadFile(item.CodePath)
				if err != nil {
					return fail(err)
				}
			} else if !errors.Is(err, os.ErrNotExist) {
				return fail(err)
			}
			if existed && !overwrite {
				return fail(fmt.Errorf("%w: refusing to overwrite existing sidecar code %s", errFormMigrateConflict, item.CodePath))
			}
			body := item.Code
			if body != "" {
				body += "\n"
			}
			if err := os.WriteFile(item.CodePath, []byte(body), 0o644); err != nil {
				return fail(err)
			}
			rollback = append(rollback, formMigrationRollback{Path: item.CodePath, Existed: existed, Body: oldBody})
			if existed {
				updated = append(updated, displayPath(a.cwd, item.CodePath))
			} else {
				created = append(created, displayPath(a.cwd, item.CodePath))
			}
		}
		specOutput := forms.SnapshotOutput{Path: item.SpecPath, DisplayPath: displayPath(a.cwd, item.SpecPath), Format: "yaml"}
		specExisted := false
		var oldSpecBody []byte
		if _, err := os.Stat(item.SpecPath); err == nil {
			specExisted = true
			oldSpecBody, err = os.ReadFile(item.SpecPath)
			if err != nil {
				return fail(err)
			}
		} else if !errors.Is(err, os.ErrNotExist) {
			return fail(err)
		}
		if specExisted && !overwrite {
			return fail(fmt.Errorf("%w: refusing to overwrite existing Designer spec %s", errFormMigrateConflict, item.SpecPath))
		}
		if err := forms.WriteSnapshot(specOutput, item.Spec); err != nil {
			return fail(err)
		}
		rollback = append(rollback, formMigrationRollback{Path: item.SpecPath, Existed: specExisted, Body: oldSpecBody})
		if specExisted {
			updated = append(updated, displayPath(a.cwd, item.SpecPath))
		} else {
			created = append(created, displayPath(a.cwd, item.SpecPath))
		}
	}
	if before != "sidecar" {
		if err := config.UpdateUserFormCodeSource(filepath.Join(a.cwd, config.FileName), "sidecar"); err != nil {
			return fail(err)
		}
		updated = append(updated, config.FileName)
	}
	return created, updated, skipped, nil
}

func rollbackFormMigrationWrites(items []formMigrationRollback) error {
	var errs []error
	for i := len(items) - 1; i >= 0; i-- {
		item := items[i]
		if item.Existed {
			if err := os.WriteFile(item.Path, item.Body, 0o644); err != nil {
				errs = append(errs, err)
			}
			continue
		}
		if err := os.Remove(item.Path); err != nil && !errors.Is(err, os.ErrNotExist) {
			errs = append(errs, err)
		}
	}
	return errors.Join(errs...)
}

func renderFormMigrationFiles(root string, items []formMigrationFile) []map[string]any {
	out := make([]map[string]any, 0, len(items))
	for _, item := range items {
		entry := map[string]any{
			"name":      item.Name,
			"frm_path":  displayPath(root, item.FRMPath),
			"code_path": displayPath(root, item.CodePath),
			"spec_path": displayPath(root, item.SpecPath),
		}
		if item.FRXPath != "" {
			entry["frx_path"] = displayPath(root, item.FRXPath)
		}
		out = append(out, entry)
	}
	return out
}

func firstAttributeLine(text string) int {
	lines := strings.Split(strings.ReplaceAll(strings.ReplaceAll(text, "\r\n", "\n"), "\r", "\n"), "\n")
	for i, line := range lines {
		if strings.HasPrefix(strings.TrimSpace(line), "Attribute VB_") {
			return i + 1
		}
	}
	return 0
}

func userFormVBNameFromFRM(path string) (string, error) {
	body, err := os.ReadFile(path)
	if err != nil {
		return "", err
	}
	lines := strings.Split(strings.ReplaceAll(strings.ReplaceAll(string(body), "\r\n", "\n"), "\r", "\n"), "\n")
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if !strings.HasPrefix(strings.ToLower(trimmed), "attribute vb_name") {
			continue
		}
		_, value, ok := strings.Cut(trimmed, "=")
		if !ok {
			break
		}
		name := strings.Trim(strings.TrimSpace(value), `"`)
		if name != "" {
			return name, nil
		}
	}
	return "", fmt.Errorf("read %s: Attribute VB_Name was not found", path)
}

func isBareComponentName(name string) bool {
	if name == "" || name != filepath.Base(name) || strings.Contains(name, "..") || strings.ContainsAny(name, `/\`) {
		return false
	}
	for i, r := range name {
		if i == 0 {
			if (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z') || r == '_' {
				continue
			}
			return false
		}
		if (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z') || (r >= '0' && r <= '9') || r == '_' {
			continue
		}
		return false
	}
	return true
}

func sameCLIPath(a, b string) bool {
	return strings.EqualFold(filepath.Clean(a), filepath.Clean(b))
}

func collectPackSourceModules(root string, cfg config.Config) ([]packpkg.SourceModule, error) {
	var sources []packpkg.SourceModule
	collect := func(dir string, typ packpkg.ModuleType, exts ...string) error {
		base := workbookArgPath(root, dir)
		if strings.TrimSpace(base) == "" {
			return nil
		}
		if _, err := os.Stat(base); err != nil {
			if errors.Is(err, os.ErrNotExist) {
				return nil
			}
			return err
		}
		allowed := map[string]bool{}
		for _, ext := range exts {
			allowed[strings.ToLower(ext)] = true
		}
		return filepath.WalkDir(base, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if d == nil || d.IsDir() {
				return nil
			}
			ext := strings.ToLower(filepath.Ext(d.Name()))
			if !allowed[ext] {
				return nil
			}
			body, err := os.ReadFile(path)
			if err != nil {
				return err
			}
			name := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
			sources = append(sources, packpkg.SourceModule{
				Name:   name,
				Type:   typ,
				Source: string(body),
			})
			return nil
		})
	}
	if err := collect(cfg.Src.Modules, packpkg.ModuleTypeStandard, ".bas"); err != nil {
		return nil, err
	}
	if err := collect(cfg.Src.Classes, packpkg.ModuleTypeClass, ".cls"); err != nil {
		return nil, err
	}
	if err := collect(cfg.Src.Workbook, packpkg.ModuleTypeDocument, ".bas", ".cls"); err != nil {
		return nil, err
	}
	formSources, err := collectFormSources(root, cfg)
	if err != nil {
		return nil, err
	}
	sources = append(sources, formSources...)
	return sources, nil
}

// collectFormSources reads UserForm sources honoring [userform].code_source. In frm mode the .frm is the
// code authority; in sidecar mode the authoritative code-behind is src/forms/code/<FormName>.bas and is
// merged into the .frm text IN MEMORY (the on-disk .frm is never modified — pack must not dirty sources).
// Mismatches (a sidecar carrying Attribute VB_* headers, or a sidecar with no matching .frm) fail loud as
// ErrAmbiguousLayout. The canonical merge/validation live in internal/excel/forms (used by push/pull); pack
// reuses them read-only rather than the disk-writing runUserFormCodeSourcePreflight.
func collectFormSources(root string, cfg config.Config) ([]packpkg.SourceModule, error) {
	formsDir := workbookArgPath(root, cfg.Src.Forms)
	if strings.TrimSpace(formsDir) == "" {
		return nil, nil
	}
	if _, err := os.Stat(formsDir); err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return nil, nil
		}
		return nil, err
	}
	sidecar := strings.EqualFold(cfg.UserForm.CodeSource, "sidecar")
	if sidecar {
		issues, err := forms.ValidateUserFormCodeSidecars(formsDir, nil)
		if err != nil {
			return nil, err
		}
		if len(issues) > 0 {
			return nil, fmt.Errorf("%w: %d UserForm sidecar issue(s) under %s; first issue: %s", packpkg.ErrAmbiguousLayout, len(issues), filepath.Join(cfg.Src.Forms, "code"), issues[0].Error())
		}
		if err := rejectOrphanFormSidecars(formsDir); err != nil {
			return nil, err
		}
	}
	codeDir := filepath.Join(formsDir, "code")
	var sources []packpkg.SourceModule
	walkErr := filepath.WalkDir(formsDir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			if samePath(path, codeDir) {
				return filepath.SkipDir
			}
			return nil
		}
		if !strings.EqualFold(filepath.Ext(d.Name()), ".frm") {
			return nil
		}
		formName := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		frmBody, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		source := string(frmBody)
		if sidecar {
			basBody, err := os.ReadFile(filepath.Join(codeDir, formName+".bas"))
			switch {
			case err == nil:
				source = forms.MergeUserFormCodeIntoFRM(string(frmBody), string(basBody))
			case errors.Is(err, os.ErrNotExist):
				// no sidecar for this form: keep the .frm code as-is
			default:
				return err
			}
		}
		sources = append(sources, packpkg.SourceModule{
			Name:   formName,
			Type:   packpkg.ModuleTypeForm,
			Source: source,
		})
		return nil
	})
	if walkErr != nil {
		return nil, walkErr
	}
	return sources, nil
}

// rejectOrphanFormSidecars fails loud when src/forms/code/<X>.bas has no matching <X>.frm — a sidecar
// pointing at a form pack cannot place is an ambiguous layout, not a silent skip.
func rejectOrphanFormSidecars(formsDir string) error {
	codeDir := filepath.Join(formsDir, "code")
	// Discover form names the same way collectFormSources does — every .frm under
	// formsDir recursively, excluding the code/ sidecar dir — so a sidecar for a form
	// kept in a subdirectory is matched, not mistaken for an orphan.
	formNames := map[string]bool{}
	walkErr := filepath.WalkDir(formsDir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			if samePath(path, codeDir) {
				return filepath.SkipDir
			}
			return nil
		}
		if strings.EqualFold(filepath.Ext(d.Name()), ".frm") {
			formNames[strings.ToLower(strings.TrimSuffix(d.Name(), filepath.Ext(d.Name())))] = true
		}
		return nil
	})
	if walkErr != nil {
		return walkErr
	}
	entries, err := os.ReadDir(codeDir)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return nil
		}
		return err
	}
	for _, e := range entries {
		if e.IsDir() || !strings.EqualFold(filepath.Ext(e.Name()), ".bas") {
			continue
		}
		formName := strings.TrimSuffix(e.Name(), filepath.Ext(e.Name()))
		if !formNames[strings.ToLower(formName)] {
			return fmt.Errorf("%w: UserForm sidecar %q has no matching %s.frm", packpkg.ErrAmbiguousLayout, formName, formName)
		}
	}
	return nil
}

func uniqueNonEmptyPaths(paths ...string) []string {
	seen := map[string]bool{}
	var out []string
	for _, path := range paths {
		if strings.TrimSpace(path) == "" {
			continue
		}
		key := strings.ToLower(filepath.Clean(path))
		if seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, path)
	}
	return out
}

func officeLockFilePresent(workbookPath string) (string, bool) {
	lockPath := filepath.Join(filepath.Dir(workbookPath), "~$"+filepath.Base(workbookPath))
	if _, err := os.Stat(lockPath); err == nil {
		return lockPath, true
	}
	return lockPath, false
}

func writePackOutput(path string, body []byte) (bool, error) {
	parent := filepath.Dir(path)
	createdParentDirs := false
	if parent != "." && parent != "" {
		if _, err := os.Stat(parent); err != nil {
			if !errors.Is(err, os.ErrNotExist) {
				return false, err
			}
			createdParentDirs = true
		}
		if err := os.MkdirAll(parent, 0o755); err != nil {
			return false, err
		}
	}
	return createdParentDirs, os.WriteFile(path, body, 0o644)
}

func buildDiffOptions(root, beforeWorkbook, afterWorkbook, vbaBefore, vbaAfter string) (diff.Options, error) {
	if (vbaBefore == "") != (vbaAfter == "") {
		return diff.Options{}, fmt.Errorf("--vba-before and --vba-after must be provided together")
	}
	if err := validateWorkbookDiffExt(beforeWorkbook); err != nil {
		return diff.Options{}, fmt.Errorf("before workbook: %w", err)
	}
	if err := validateWorkbookDiffExt(afterWorkbook); err != nil {
		return diff.Options{}, fmt.Errorf("after workbook: %w", err)
	}
	return diff.Options{
		BeforeWorkbook: workbookArgPath(root, beforeWorkbook),
		AfterWorkbook:  workbookArgPath(root, afterWorkbook),
		VBABeforeDir:   workbookArgPath(root, vbaBefore),
		VBAAfterDir:    workbookArgPath(root, vbaAfter),
	}, nil
}

func validateWorkbookDiffExt(path string) error {
	return workbookformat.ValidateDiffWorkbook(path)
}

func workbookArgPath(root, path string) string {
	if path == "" || filepath.IsAbs(path) {
		return path
	}
	return filepath.Join(root, path)
}

type inspectCellSelector struct {
	Sheet   string
	Address string
}

func validateInspectFormat(format string) (string, error) {
	switch normalized := strings.ToLower(strings.TrimSpace(format)); normalized {
	case "", "text":
		return "text", nil
	case "json", "markdown":
		return normalized, nil
	default:
		return "", fmt.Errorf("unsupported inspect format %q; expected text, json, or markdown", format)
	}
}

func buildInspectLimits(maxRows, maxCols int) (workbookinspect.Limits, error) {
	if maxRows <= 0 {
		return workbookinspect.Limits{}, fmt.Errorf("--max-rows must be greater than 0")
	}
	if maxCols <= 0 {
		return workbookinspect.Limits{}, fmt.Errorf("--max-cols must be greater than 0")
	}
	return workbookinspect.Limits{MaxRows: maxRows, MaxCols: maxCols}, nil
}

func buildInspectSheetSelector(args []string, sheet string) (string, error) {
	if len(args) > 0 {
		if strings.TrimSpace(sheet) != "" {
			return "", fmt.Errorf("cannot combine a positional sheet selector with --sheet")
		}
		return parseInspectSheetLiteral(args[0]), nil
	}
	if strings.TrimSpace(sheet) == "" {
		return "", fmt.Errorf("worksheet name is required")
	}
	return parseInspectSheetLiteral(sheet), nil
}

func buildInspectCellSelector(args []string, sheet, address string, allowRange bool) (inspectCellSelector, error) {
	if len(args) > 0 {
		if strings.TrimSpace(sheet) != "" || strings.TrimSpace(address) != "" {
			return inspectCellSelector{}, fmt.Errorf("cannot combine a positional selector with --sheet or --address")
		}
		return parseInspectSelectorLiteral(args[0], allowRange)
	}
	if strings.TrimSpace(sheet) == "" || strings.TrimSpace(address) == "" {
		return inspectCellSelector{}, fmt.Errorf("--sheet and --address are required when no positional selector is provided")
	}
	return parseInspectSelectorLiteral(
		fmt.Sprintf("%s!%s", parseInspectSheetLiteral(sheet), strings.TrimSpace(address)),
		allowRange,
	)
}

func parseInspectSelectorLiteral(literal string, allowRange bool) (inspectCellSelector, error) {
	index := strings.LastIndex(literal, "!")
	if index <= 0 || index == len(literal)-1 {
		return inspectCellSelector{}, fmt.Errorf("expected selector in the form Sheet!A1")
	}
	selector := inspectCellSelector{
		Sheet:   parseInspectSheetLiteral(literal[:index]),
		Address: strings.TrimSpace(literal[index+1:]),
	}
	if selector.Sheet == "" {
		return inspectCellSelector{}, fmt.Errorf("worksheet name is required")
	}
	if selector.Address == "" {
		return inspectCellSelector{}, fmt.Errorf("address is required")
	}
	if allowRange {
		normalized, err := validateInspectRangeAddress(selector.Address)
		if err != nil {
			return inspectCellSelector{}, err
		}
		selector.Address = normalized
		return selector, nil
	}
	normalized, err := validateInspectCellAddress(selector.Address)
	if err != nil {
		return inspectCellSelector{}, err
	}
	selector.Address = normalized
	return selector, nil
}

func validateInspectCellAddress(address string) (string, error) {
	clean := strings.ToUpper(strings.TrimSpace(strings.ReplaceAll(address, "$", "")))
	if clean == "" {
		return "", fmt.Errorf("address is required")
	}
	if strings.Contains(clean, ":") {
		return "", fmt.Errorf("expected a single cell address, got %q", address)
	}
	if _, _, err := excelize.CellNameToCoordinates(clean); err != nil {
		return "", fmt.Errorf("invalid address %q: %w", address, err)
	}
	return clean, nil
}

func validateInspectRangeAddress(address string) (string, error) {
	parts := strings.SplitN(strings.TrimSpace(address), ":", 2)
	if len(parts) == 0 || strings.TrimSpace(parts[0]) == "" {
		return "", fmt.Errorf("address is required")
	}
	first, err := validateInspectCellAddress(parts[0])
	if err != nil {
		return "", err
	}
	if len(parts) == 1 {
		return first, nil
	}
	last, err := validateInspectCellAddress(parts[1])
	if err != nil {
		return "", err
	}
	firstCol, firstRow, _ := excelize.CellNameToCoordinates(first)
	lastCol, lastRow, _ := excelize.CellNameToCoordinates(last)
	if lastCol < firstCol {
		firstCol, lastCol = lastCol, firstCol
	}
	if lastRow < firstRow {
		firstRow, lastRow = lastRow, firstRow
	}
	start, err := excelize.CoordinatesToCellName(firstCol, firstRow)
	if err != nil {
		return "", err
	}
	end, err := excelize.CoordinatesToCellName(lastCol, lastRow)
	if err != nil {
		return "", err
	}
	return start + ":" + end, nil
}

func parseInspectSheetLiteral(sheet string) string {
	trimmed := strings.TrimSpace(sheet)
	if len(trimmed) >= 2 && trimmed[0] == '\'' && trimmed[len(trimmed)-1] == '\'' {
		trimmed = trimmed[1 : len(trimmed)-1]
		trimmed = strings.ReplaceAll(trimmed, "''", "'")
	}
	return trimmed
}

func (a *app) fmtCommand() *cobra.Command {
	var write bool
	var check bool
	var diff bool
	var stdin bool
	var lineNumbers string

	cmd := &cobra.Command{
		Use:   "fmt [path...]",
		Short: "Format VBA source files (.bas, .cls)",
		Args:  cobra.ArbitraryArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			if stdin {
				if len(args) > 0 {
					return a.writeFailure("fmt", output.ExitConfig, "fmt_args_invalid", fmt.Errorf("--stdin cannot be combined with path arguments"))
				}
				if write || check || diff {
					return a.writeFailure("fmt", output.ExitConfig, "fmt_args_invalid", fmt.Errorf("--stdin cannot be combined with --write, --check, or --diff"))
				}
				if cmd.Flags().Changed("line-numbers") {
					return a.writeFailure("fmt", output.ExitConfig, "fmt_args_invalid", fmt.Errorf("--stdin cannot be combined with --line-numbers"))
				}
				return a.runFmtStdin()
			}
			modeCount := 0
			if write {
				modeCount++
			}
			if check {
				modeCount++
			}
			if diff {
				modeCount++
			}
			if modeCount > 1 {
				return a.writeFailure("fmt", output.ExitConfig, "fmt_args_invalid", fmt.Errorf("--write, --check, and --diff cannot be combined"))
			}
			lineNumberMode := vbafmt.LineNumberMode(lineNumbers)
			switch lineNumberMode {
			case vbafmt.LineNumberModePreserve, vbafmt.LineNumberModeAdd, vbafmt.LineNumberModeRemove, vbafmt.LineNumberModeRenumber:
			default:
				return a.writeFailure("fmt", output.ExitConfig, "fmt_args_invalid", fmt.Errorf("invalid --line-numbers %q: expected preserve, add, remove, or renumber", lineNumbers))
			}
			cfg, err := a.loadFmtConfig("fmt", false)
			if err != nil {
				return err
			}
			formatCfg := fmtFormatConfigFromConfig(cfg)
			opts := vbafmt.FmtOptions{
				Write:                 write,
				Check:                 check,
				Diff:                  diff,
				Paths:                 args,
				Root:                  a.cwd,
				Cfg:                   cfg,
				LineNumbers:           lineNumberMode,
				OperatorSpacing:       formatCfg.OperatorSpacing,
				OperatorSpacingSet:    formatCfg.OperatorSpacingSet,
				DeclarationSpacing:    formatCfg.DeclarationSpacing,
				DeclarationSpacingSet: formatCfg.DeclarationSpacingSet,
				KeywordCasing:         formatCfg.KeywordCasing,
				KeywordCasingSet:      formatCfg.KeywordCasingSet,
				BuiltinCasing:         formatCfg.BuiltinCasing,
				BuiltinCasingSet:      formatCfg.BuiltinCasingSet,
			}
			result, err := vbafmt.Run(opts)
			if err != nil {
				return a.writeFailure("fmt", output.ExitEnvironment, "fmt_failed", err)
			}

			env := output.New("fmt")
			targetPath := buildFmtTargetPath(cfg, args)
			env.Target = map[string]any{
				"kind":        "source",
				"path":        targetPath,
				"description": "source files",
			}
			mode := "inspect"
			if write {
				mode = "write"
			} else if check {
				mode = "check"
			} else if diff {
				mode = "diff"
			}
			outputPayload := map[string]any{
				"mode":         mode,
				"changed":      result.Changed,
				"unchanged":    result.Unchanged,
				"skipped":      result.Skipped,
				"total":        result.Total,
				"line_numbers": buildFmtLineNumbersPayload(mode, result.LineNumbers),
			}
			if len(result.LineNumbers.Warnings) > 0 {
				lineWarnings := make([]map[string]any, 0, len(result.LineNumbers.Warnings))
				for _, warning := range result.LineNumbers.Warnings {
					item := map[string]any{
						"message": warning.Message,
					}
					if warning.Path != "" {
						item["path"] = warning.Path
					}
					if warning.Line > 0 {
						item["line"] = warning.Line
					}
					lineWarnings = append(lineWarnings, item)
				}
				outputPayload["line_numbers"].(map[string]any)["warnings"] = lineWarnings
			}
			if len(result.ChangedPaths) > 0 {
				outputPayload["changed_paths"] = pathsToAny(result.ChangedPaths)
			}
			if len(result.SkippedPaths) > 0 {
				outputPayload["skipped_paths"] = pathsToAny(result.SkippedPaths)
			}
			if len(result.SkippedReasons) > 0 {
				reasons := make([]map[string]any, 0, len(result.SkippedReasons))
				for _, sr := range result.SkippedReasons {
					reasons = append(reasons, map[string]any{
						"path":   sr.Path,
						"reason": sr.Reason,
					})
				}
				outputPayload["skipped_reasons"] = reasons
			}
			env.Output = outputPayload

			var warnings []map[string]any
			for _, sr := range result.SkippedReasons {
				code := sr.Reason
				if code == "" {
					code = "fmt_skipped_unsupported_extension"
				}
				warnings = append(warnings, map[string]any{
					"code":    code,
					"message": fmt.Sprintf("Skipped file: %s", sr.Path),
				})
			}
			if len(warnings) > 0 {
				env.Warnings = warnings
			}
			if len(result.LineNumbers.Warnings) > 0 {
				for _, warning := range result.LineNumbers.Warnings {
					item := map[string]any{
						"code":    "fmt_line_numbers_warning",
						"message": warning.Message,
					}
					if warning.Path != "" {
						item["file"] = warning.Path
					}
					if warning.Line > 0 {
						item["line"] = warning.Line
					}
					warnings = append(warnings, item)
				}
				env.Warnings = warnings
			}

			if check && result.Changed > 0 {
				env.Hints = []map[string]any{
					{"code": "fmt_write_hint", "message": "Run `xlflow fmt --write` to apply formatting changes."},
				}
				env.Status = output.StatusFailed
				env.Error = &output.Error{
					Code:    "fmt_check_failed",
					Message: fmt.Sprintf("%d file(s) not formatted", result.Changed),
				}
				return a.write(env, output.ExitValidation)
			}

			if diff && result.Changed > 0 {
				env.Hints = []map[string]any{
					{"code": "fmt_write_hint", "message": "Run `xlflow fmt --write` to apply formatting changes."},
				}
				env.Logs = a.buildFmtDiffLogs(result)
				return a.write(env, output.ExitSuccess)
			}

			if result.Changed > 0 {
				env.Logs = a.buildFmtWriteLogs(write, result)
			} else {
				env.Logs = []string{fmt.Sprintf("%d file(s) already formatted", result.Total)}
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&write, "write", false, "write formatted source back to files")
	cmd.Flags().BoolVar(&check, "check", false, "check whether source files are formatted without modifying them")
	cmd.Flags().BoolVar(&diff, "diff", false, "show unified diff of formatting changes without modifying files")
	cmd.Flags().BoolVar(&stdin, "stdin", false, "read VBA source from stdin and write formatted output to stdout")
	cmd.Flags().StringVar(&lineNumbers, "line-numbers", string(vbafmt.LineNumberModePreserve), "line-number policy: preserve, add, remove, or renumber")
	return cmd
}

func (a *app) runFmtStdin() error {
	var buf bytes.Buffer
	if _, err := io.Copy(&buf, os.Stdin); err != nil {
		return a.writeFailure("fmt", output.ExitEnvironment, "fmt_stdin_read_failed", err)
	}
	input := buf.String()
	if strings.TrimSpace(input) == "" {
		if a.json {
			env := output.New("fmt")
			env.Target = map[string]any{
				"kind":        "source",
				"description": "stdin input",
			}
			env.Output = map[string]any{
				"mode":      "inspect",
				"changed":   0,
				"unchanged": 1,
				"skipped":   0,
				"total":     1,
			}
			env.Logs = []string{"0 file(s) formatted"}
			return a.write(env, output.ExitSuccess)
		}
		_, _ = fmt.Fprintln(a.stdoutWriter())
		return nil
	}
	cfg, err := a.loadFmtConfig("fmt", true)
	if err != nil {
		return err
	}
	formatCfg := fmtFormatConfigFromConfig(cfg)
	formatted, err := vbafmt.FormatTextWithOptions(input, looksLikeClassModule(input), formatCfg)
	if err != nil {
		return a.writeFailure("fmt", output.ExitEnvironment, "fmt_failed", err)
	}
	if a.json {
		env := output.New("fmt")
		env.Target = map[string]any{
			"kind":        "source",
			"description": "stdin input",
		}
		changed := 0
		if formatted != input {
			changed = 1
		}
		env.Output = map[string]any{
			"mode":      "inspect",
			"changed":   changed,
			"unchanged": 1 - changed,
			"skipped":   0,
			"total":     1,
		}
		env.Logs = []string{fmt.Sprintf("%d file(s) would be formatted", changed)}
		return a.write(env, output.ExitSuccess)
	}
	_, err = fmt.Fprint(a.stdoutWriter(), formatted)
	if err != nil {
		return a.writeFailure("fmt", output.ExitEnvironment, "fmt_stdin_write_failed", err)
	}
	return nil
}

func (a *app) loadFmtConfig(command string, allowMissing bool) (config.Config, error) {
	cfg, err := config.Load(a.cwd)
	if err != nil && errors.Is(err, config.ErrInvalidExcelBridge) && a.hasValidBridgeOverride() {
		cfg, err = config.LoadAllowInvalidExcelBridge(a.cwd)
	}
	if err != nil {
		if allowMissing && errors.Is(err, config.ErrConfigNotFound) {
			return config.Default(), nil
		}
		return cfg, a.writeFailure(command, output.ExitConfig, "config_error", err)
	}
	if len(cfg.Warnings) > 0 {
		a.configWarnings = append(a.configWarnings, cfg.Warnings...)
	}
	return cfg, nil
}

func fmtFormatConfigFromConfig(cfg config.Config) vbafmt.FormatConfig {
	return vbafmt.FormatConfig{
		OperatorSpacing:       cfg.Fmt.OperatorSpacing,
		OperatorSpacingSet:    true,
		DeclarationSpacing:    cfg.Fmt.DeclarationSpacing,
		DeclarationSpacingSet: true,
		KeywordCasing:         cfg.Fmt.KeywordCasing,
		KeywordCasingSet:      true,
		BuiltinCasing:         cfg.Fmt.BuiltinCasing,
		BuiltinCasingSet:      true,
	}
}

func looksLikeClassModule(input string) bool {
	for _, line := range strings.Split(input, "\n") {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" {
			continue
		}
		upper := strings.ToUpper(trimmed)
		if strings.HasPrefix(upper, "VERSION ") && strings.Contains(upper, "CLASS") {
			return true
		}
		if strings.HasPrefix(upper, "ATTRIBUTE VB_") {
			return true
		}
		return false
	}
	return false
}

func (a *app) buildFmtWriteLogs(wasWritten bool, result *vbafmt.Result) []string {
	var logs []string
	if wasWritten {
		logs = append(logs, fmt.Sprintf("%d file(s) formatted", result.Changed))
	} else {
		logs = append(logs, fmt.Sprintf("%d file(s) would be formatted", result.Changed))
	}
	for _, path := range result.ChangedPaths {
		if wasWritten {
			logs = append(logs, fmt.Sprintf("formatted %s", displayPath(a.cwd, path)))
		} else {
			logs = append(logs, fmt.Sprintf("would format %s", displayPath(a.cwd, path)))
		}
	}
	for _, path := range result.SkippedPaths {
		logs = append(logs, fmt.Sprintf("skipped %s", displayPath(a.cwd, path)))
	}
	return logs
}

func buildFmtLineNumbersPayload(mode string, summary vbafmt.LineNumberSummary) map[string]any {
	payload := map[string]any{
		"mode":    string(summary.Mode),
		"applied": mode == "write",
	}
	if mode == "write" {
		payload["files_changed"] = summary.FilesChanged
		payload["lines_added"] = summary.LinesAdded
		payload["lines_removed"] = summary.LinesRemoved
		payload["lines_renumbered"] = summary.LinesRenumbered
		return payload
	}
	payload["files_to_change"] = summary.FilesChanged
	payload["lines_to_add"] = summary.LinesAdded
	payload["lines_to_remove"] = summary.LinesRemoved
	payload["lines_to_renumber"] = summary.LinesRenumbered
	return payload
}

func (a *app) buildFmtDiffLogs(result *vbafmt.Result) []string {
	var logs []string
	logs = append(logs, fmt.Sprintf("%d file(s) would be reformatted", result.Changed))
	for _, path := range result.ChangedPaths {
		formatted, ok := result.FormattedByPath[path]
		if !ok {
			logs = append(logs, fmt.Sprintf("error reading %s: formatted content not found", displayPath(a.cwd, path)))
			continue
		}
		original, err := os.ReadFile(path)
		if err != nil {
			logs = append(logs, fmt.Sprintf("error reading %s: %v", displayPath(a.cwd, path), err))
			continue
		}
		if diffText := vbafmt.Diff(displayPath(a.cwd, path), string(original), formatted); diffText != "" {
			logs = append(logs, diffText)
		}
	}
	return logs
}

func pathsToAny(paths []string) []any {
	result := make([]any, 0, len(paths))
	for _, p := range paths {
		result = append(result, p)
	}
	return result
}

func buildFmtTargetPath(cfg config.Config, args []string) string {
	if len(args) > 0 {
		return strings.Join(args, ", ")
	}
	modules := cfg.Src.Modules
	if modules == "" {
		modules = filepath.ToSlash(filepath.Join("src", "modules"))
	}
	classes := cfg.Src.Classes
	if classes == "" {
		classes = filepath.ToSlash(filepath.Join("src", "classes"))
	}
	workbook := cfg.Src.Workbook
	if workbook == "" {
		workbook = filepath.ToSlash(filepath.Join("src", "workbook"))
	}
	dirs := []string{modules, classes, workbook, "tests"}
	return strings.Join(dirs, ", ")
}

func (a *app) lintCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "lint",
		Short: "Lint VBA source files",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("lint")
			if err != nil {
				return err
			}
			lintResult, err := lint.Linter{RootDir: a.cwd, Config: cfg}.RunResult()
			if err != nil {
				return a.writeFailure("lint", output.ExitEnvironment, "lint_failed", err)
			}
			issues := lintResult.Issues
			env := output.New("lint")
			env.Issues = issues
			env.Warnings = lintResult.Warnings
			if len(issues) > 0 {
				env.Status = output.StatusFailed
				env.Error = &output.Error{Code: "lint_failed", Message: fmt.Sprintf("%d lint issue(s) found", len(issues))}
				return a.write(env, output.ExitValidation)
			}
			env.Logs = []string{"no lint issues found"}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func (a *app) lspCommand() *cobra.Command {
	var stdio bool
	var check bool
	var showVersion bool
	var logFile string
	var performanceLog bool
	cmd := &cobra.Command{
		Use:   "lsp",
		Short: "Run the VBA language server",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			modeCount := 0
			for _, enabled := range []bool{stdio, check, showVersion} {
				if enabled {
					modeCount++
				}
			}
			if modeCount != 1 {
				return a.writeFailure("lsp", output.ExitConfig, "lsp_args_invalid", fmt.Errorf("exactly one of --stdio, --check, or --version is required"))
			}
			if showVersion {
				env := output.New("lsp")
				env.Version = map[string]any{
					"name":    "xlflow-vba-lsp",
					"version": a.buildInfo.withDefaults().Version,
					"commit":  a.buildInfo.withDefaults().Commit,
					"date":    a.buildInfo.withDefaults().Date,
				}
				env.Logs = []string{"xlflow-vba-lsp " + a.buildInfo.withDefaults().Version}
				return a.write(env, output.ExitSuccess)
			}
			cfg, err := a.loadLSPConfig()
			if err != nil {
				return err
			}
			opts := lspserver.Options{
				RootDir: a.cwd,
				Config:  cfg,
				Build: lspserver.BuildInfo{
					Version: a.buildInfo.withDefaults().Version,
					Commit:  a.buildInfo.withDefaults().Commit,
					Date:    a.buildInfo.withDefaults().Date,
				},
				LogFile:        logFile,
				PerformanceLog: performanceLog,
				Stderr:         a.stderrWriter(),
			}
			if check {
				a.ensureLSPTypeDBGenerated()
				if err := lspserver.Check(opts); err != nil {
					return a.writeFailure("lsp", output.ExitEnvironment, "lsp_check_failed", err)
				}
				typeDBStatus, statusErr := typedb.StatusFor(typedb.Options{GeneratorVersion: a.buildInfo.withDefaults().Version})
				typeDatabase := "builtin"
				if statusErr == nil && typeDBStatus.ManifestExists && !typeDBStatus.Stale {
					typeDatabase = "builtin+global_generated"
				}
				env := output.New("lsp")
				env.Diagnostics = map[string]any{
					"server":        "xlflow-vba-lsp",
					"transport":     "stdio",
					"type_database": typeDatabase,
					"sync":          "full",
				}
				if statusErr == nil {
					env.TypeDB = typeDBStatus
				}
				env.Logs = []string{"lsp pre-launch check passed"}
				return a.write(env, output.ExitSuccess)
			}
			a.ensureLSPTypeDBGenerated()
			return lspserver.RunStdio(opts)
		},
	}
	cmd.Flags().BoolVar(&stdio, "stdio", false, "run the language server over stdio")
	cmd.Flags().BoolVar(&check, "check", false, "validate LSP prerequisites and exit")
	cmd.Flags().BoolVar(&showVersion, "version", false, "show LSP server version and exit")
	cmd.Flags().StringVar(&logFile, "log-file", "", "write LSP logs to this file instead of stderr")
	cmd.Flags().BoolVar(&performanceLog, "performance-log", false, "log LSP operation performance metrics")
	return cmd
}

func (a *app) ensureLSPTypeDBGenerated() {
	status, err := typedb.StatusFor(typedb.Options{GeneratorVersion: a.buildInfo.withDefaults().Version})
	if err != nil {
		a.writeLSPStderr("xlflow-lsp: generated TypeLib DB status could not be inspected: %v\n", err)
		return
	}
	if status.ManifestExists && !status.Stale {
		return
	}
	resolvedDir, err := typedb.ResolveDir("")
	if err != nil {
		a.writeLSPStderr("xlflow-lsp: generated TypeLib DB directory could not be resolved: %v\n", err)
		return
	}
	state := "missing"
	if status.ManifestExists && status.Stale {
		state = "stale"
	}
	a.writeLSPStderr("xlflow-lsp: generated TypeLib DB %s; attempting best-effort generation at %s\n", state, resolvedDir)
	typeDBEnv, code, err := a.excelRunner().TypeDBImport(excel.TypeDBImportOptions{
		OutputDir:        resolvedDir,
		GeneratorVersion: a.buildInfo.withDefaults().Version,
		Libraries:        []string{"all"},
		Keepalive:        buildCommandOptions(a.stderrWriter()),
	})
	if err != nil {
		a.writeLSPStderr("xlflow-lsp: generated TypeLib DB generation skipped: %v\n", err)
		return
	}
	if code != output.ExitSuccess {
		reason := "unknown error"
		if typeDBEnv.Error != nil && typeDBEnv.Error.Message != "" {
			reason = typeDBEnv.Error.Message
		}
		a.writeLSPStderr("xlflow-lsp: generated TypeLib DB generation skipped: %s\n", reason)
		return
	}
	a.writeLSPStderr("xlflow-lsp: generated TypeLib DB created at %s\n", resolvedDir)
}

func (a *app) writeLSPStderr(format string, args ...any) {
	_, _ = fmt.Fprintf(a.stderrWriter(), format, args...)
}

func (a *app) analyzeCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "analyze",
		Short: "Analyze VBA source for runtime-risk patterns",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("analyze")
			if err != nil {
				return err
			}
			analyzeResult, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg}.RunResult()
			if err != nil {
				return a.writeFailure("analyze", output.ExitEnvironment, "analyze_failed", err)
			}
			findings := analyzeResult.Findings
			env := output.New("analyze")
			env.Analysis = findings
			env.Warnings = analyzeResult.Warnings
			if len(findings) > 0 {
				env.Status = output.StatusFailed
				env.Error = &output.Error{Code: "analyze_failed", Message: fmt.Sprintf("%d analysis finding(s) found", len(findings))}
				return a.write(env, output.ExitValidation)
			}
			env.Logs = []string{"no analysis findings found"}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func (a *app) checkCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "check",
		Short: "Run lint, analyze, and doctor",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("check")
			if err != nil {
				return err
			}
			env := output.New("check")
			check := map[string]any{}
			lintResult, err := lint.Linter{RootDir: a.cwd, Config: cfg}.RunResult()
			if err != nil {
				return a.writeFailure("check", output.ExitEnvironment, "lint_failed", err)
			}
			issues := lintResult.Issues
			check["lint"] = map[string]any{"status": statusForCount(len(issues)), "count": len(issues)}
			analyzeResult, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg}.RunResult()
			if err != nil {
				return a.writeFailure("check", output.ExitEnvironment, "analyze_failed", err)
			}
			findings := analyzeResult.Findings
			check["analyze"] = map[string]any{"status": statusForCount(len(findings)), "count": len(findings)}
			env.Warnings = mergeWarningsUnique(lintResult.Warnings, analyzeResult.Warnings)
			var doctor output.Envelope
			var doctorCode int
			err = a.withExcelProgress("Checking Excel automation", commandOpts, func() error {
				var runErr error
				doctor, doctorCode, runErr = a.excelRunnerForConfig(cfg).Doctor(cfg, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			check["doctor"] = map[string]any{"status": doctor.Status, "code": doctorCode}
			env.Check = check
			env.Issues = issues
			env.Analysis = findings
			env.Diagnostics = doctor.Diagnostics
			env.Workbook = doctor.Workbook
			if doctor.Status == output.StatusFailed {
				env.Status = output.StatusFailed
				env.Error = doctor.Error
				return a.write(env, output.ExitEnvironment)
			}
			if len(issues) > 0 || len(findings) > 0 {
				env.Status = output.StatusFailed
				env.Error = &output.Error{Code: "check_failed", Message: "lint or analysis findings found", Source: "xlflow"}
				return a.write(env, output.ExitValidation)
			}
			env.Logs = []string{"all checks passed"}
			return a.write(env, output.ExitSuccess)
		},
	}
	return cmd
}

func statusForCount(count int) string {
	if count == 0 {
		return output.StatusOK
	}
	return output.StatusFailed
}

func mergeWarningsUnique(warningSets ...any) []any {
	var merged []any
	seen := map[string]bool{}
	for _, warningSet := range warningSets {
		for _, raw := range anySlice(warningSet) {
			warning := mapValue(raw)
			key := warningKey(warning)
			if seen[key] {
				continue
			}
			seen[key] = true
			merged = append(merged, raw)
		}
	}
	return merged
}

func warningKey(warning map[string]any) string {
	return fmt.Sprint(warning["code"]) + "|" +
		fmt.Sprint(warning["rule"]) + "|" +
		fmt.Sprint(warning["file"]) + "|" +
		fmt.Sprint(warning["line"]) + "|" +
		fmt.Sprint(warning["target_line"])
}

func (a *app) inspectGUICommand() *cobra.Command {
	return &cobra.Command{
		Use:   "inspect-gui",
		Short: "Report VBA GUI interaction boundaries",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("inspect-gui")
			if err != nil {
				return err
			}
			boundaries, err := gui.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
			if err != nil {
				return a.writeFailure("inspect-gui", output.ExitEnvironment, "inspect_gui_failed", err)
			}
			env := output.New("inspect-gui")
			env.GUIBoundaries = boundaries
			if len(boundaries) == 0 {
				env.Logs = []string{"no GUI boundaries found"}
			} else {
				env.Logs = []string{fmt.Sprintf("detected %d GUI boundary candidate(s)", len(boundaries))}
			}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func withGUIBoundarySummary(value any, boundaries []gui.Boundary) any {
	diag := map[string]any{}
	if value != nil {
		for key, item := range cliObjectMap(value) {
			diag[key] = item
		}
	}
	counts := map[string]int{}
	for _, boundary := range boundaries {
		counts[boundary.Kind]++
	}
	diag["gui_boundaries"] = map[string]any{
		"count":    len(boundaries),
		"by_kind":  counts,
		"detected": len(boundaries) > 0,
	}
	return diag
}

func (a *app) buildRunDiagnostic(cfg config.Config, env output.Envelope) map[string]any {
	diag := map[string]any{}
	for key, item := range cliObjectMap(env.RunDiagnostic) {
		diag[key] = item
	}
	state := buildStatusState(a.cwd, cfg, workbookArgPath(a.cwd, cfg.Excel.Path))
	sourceNewer := boolValueForCLI(state, "src_newer_than_workbook")
	if _, ok := diag["kind"]; !ok {
		diag["kind"] = "runtime"
	}
	if sourceNewer {
		diag["likely_cause"] = "The macro may not have been pushed yet. Source files are newer than the saved workbook."
		diag["suggestion"] = "Run `xlflow push` first, or rerun as `xlflow run --push ...` to import source before executing the macro."
		sourceState := map[string]any{
			"src_newer_than_workbook": true,
		}
		if latest := stringValueForCLI(state, "latest_source_modified_at"); latest != "" {
			sourceState["latest_source_modified_at"] = latest
		}
		if workbook := stringValueForCLI(state, "workbook_last_modified_at"); workbook != "" {
			sourceState["workbook_last_modified_at"] = workbook
		}
		diag["source_state"] = sourceState
	} else if _, ok := diag["likely_cause"]; !ok {
		diag["likely_cause"] = "The macro failed while running user VBA code."
	}
	if _, ok := diag["suggestion"]; !ok {
		diag["suggestion"] = "Inspect the failing procedure, add targeted XlflowDebug.Log calls around the suspected block, and rerun with --json."
	}
	if env.Error != nil {
		if !cliLocationHasMeaningfulData(diag["location"]) {
			diag["location"] = map[string]any{
				"module": env.Error.Source,
				"line":   env.Error.Line,
			}
		}
		if env.Error.Number == 450 && diag["likely_cause"] == "The macro failed while running user VBA code." {
			diag["likely_cause"] = "VBA runtime error 450 often means an invalid property assignment, wrong argument count, or missing Set for an object assignment."
			diag["suggestion"] = "Check object assignments near the reported line; use Set when assigning Workbook, Worksheet, Range, or other object references."
		}
	}
	findings, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
	if err == nil && env.Error != nil && !sourceNewer {
		for _, finding := range findings {
			if finding.Module != "" && env.Error.Source != "" && !strings.EqualFold(finding.Module, env.Error.Source) {
				continue
			}
			if env.Error.Line > 0 && finding.Line > 0 && absInt(finding.Line-env.Error.Line) > 3 {
				continue
			}
			diag["location"] = map[string]any{
				"file":      finding.File,
				"module":    finding.Module,
				"procedure": finding.Procedure,
				"line":      finding.Line,
			}
			diag["nearby_code"] = finding.NearbyCode
			diag["likely_cause"] = finding.Reason
			diag["suggestion"] = finding.Suggestion
			break
		}
	}
	return diag
}

func (a *app) runSourcePreflight(command string, cfg config.Config, action string, ignoredAnalysisCodes map[string]bool, pathFilter func(string) bool) error {
	lintResult, err := lint.Linter{RootDir: a.cwd, Config: cfg, PathFilter: pathFilter}.RunResult()
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "lint_failed", err)
	}
	issues := lintResult.Issues
	blockingIssues := lint.PushBlockingIssues(issues)
	if len(blockingIssues) > 0 {
		env := output.Failure(command, output.Error{
			Code:    "lint_failed",
			Message: fmt.Sprintf("%d source issue(s) must be fixed before %s to avoid a VBA editor dialog", len(blockingIssues), action),
			Source:  "xlflow",
			Phase:   "preflight",
		})
		env.Issues = blockingIssues
		env.Warnings = lintResult.Warnings
		env.Logs = []string{"blocked before Excel automation to avoid a VBA editor dialog"}
		return a.write(env, output.ExitValidation)
	}
	analyzeResult, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg, PathFilter: pathFilter}.RunResult()
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "analyze_failed", err)
	}
	findings := analyzeResult.Findings
	blockingFindings := filterAnalysisFindings(analyze.BlockingFindings(findings), ignoredAnalysisCodes)
	if len(blockingIssues) == 0 && len(blockingFindings) == 0 {
		return nil
	}
	count := len(blockingIssues) + len(blockingFindings)
	errorCode := "source_preflight_failed"
	if len(blockingFindings) == 0 {
		errorCode = "lint_failed"
	} else if len(blockingIssues) == 0 {
		errorCode = "analyze_failed"
	}
	env := output.Failure(command, output.Error{
		Code:    errorCode,
		Message: fmt.Sprintf("%d source issue(s) must be fixed before %s to avoid a VBA editor dialog", count, action),
		Source:  "xlflow",
		Phase:   "preflight",
	})
	if len(blockingIssues) > 0 {
		env.Issues = blockingIssues
	}
	if len(blockingFindings) > 0 {
		env.Analysis = blockingFindings
	}
	env.Warnings = mergeWarningsUnique(lintResult.Warnings, analyzeResult.Warnings)
	env.Logs = []string{"blocked before Excel automation to avoid a VBA editor dialog"}
	return a.write(env, output.ExitValidation)
}

func buildUserFormSourcePathFilter(root string, cfg config.Config, targetForms map[string]bool) func(string) bool {
	if len(targetForms) == 0 {
		return nil
	}
	formsRoot := filepath.Clean(filepath.Join(root, cfg.Src.Forms))
	codeRoot := filepath.Clean(filepath.Join(formsRoot, "code"))
	return func(path string) bool {
		cleanPath := filepath.Clean(path)
		ext := strings.ToLower(filepath.Ext(cleanPath))
		base := strings.TrimSuffix(filepath.Base(cleanPath), filepath.Ext(cleanPath))
		if ext == ".frm" && isPathInsideRoot(cleanPath, formsRoot) {
			return targetForms[base]
		}
		if ext == ".bas" && isPathInsideRoot(cleanPath, codeRoot) {
			return targetForms[base]
		}
		return true
	}
}

func isPathInsideRoot(path string, root string) bool {
	cleanPath := filepath.Clean(path)
	cleanRoot := filepath.Clean(root)
	if strings.EqualFold(cleanPath, cleanRoot) {
		return true
	}
	return strings.HasPrefix(strings.ToLower(cleanPath), strings.ToLower(cleanRoot)+strings.ToLower(string(os.PathSeparator)))
}

func (a *app) runUserFormCodeSourcePreflight(command string, cfg config.Config, targetForms map[string]bool) error {
	if cfg.UserForm.CodeSource != "sidecar" {
		return nil
	}
	issues, err := forms.ValidateUserFormCodeSidecars(filepath.Join(a.cwd, cfg.Src.Forms), targetForms)
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "source_preflight_failed", err)
	}
	if len(issues) > 0 {
		rendered := make([]lint.Issue, 0, len(issues))
		for _, issue := range issues {
			rendered = append(rendered, issue.LintIssue(filepath.Join(a.cwd, cfg.Src.Forms)))
		}
		env := output.Failure(command, output.Error{
			Code:    "source_preflight_failed",
			Message: fmt.Sprintf("%d source issue(s) must be fixed before %s to avoid corrupting UserForm artifacts", len(rendered), command),
			Source:  "xlflow",
			Phase:   "preflight",
		})
		env.Issues = rendered
		env.Logs = []string{"blocked before Excel automation because a UserForm sidecar contains Attribute VB_* header lines"}
		return a.write(env, output.ExitValidation)
	}
	updated, err := forms.SyncUserFormCodeSidecars(filepath.Join(a.cwd, cfg.Src.Forms), targetForms)
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "userform_code_sync_failed", err)
	}
	if len(updated) == 0 {
		return nil
	}
	return nil
}

func (a *app) runUserFormArtifactPreflight(command string, cfg config.Config, targetForms map[string]bool) error {
	if cfg.UserForm.CodeSource != "sidecar" {
		return nil
	}
	issues, err := forms.ValidateUserFormArtifactsAgainstSpecs(filepath.Join(a.cwd, cfg.Src.Forms), targetForms)
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "source_preflight_failed", err)
	}
	if len(issues) == 0 {
		return nil
	}
	rendered := make([]lint.Issue, 0, len(issues))
	for _, issue := range issues {
		rendered = append(rendered, issue.LintIssue(filepath.Join(a.cwd, cfg.Src.Forms)))
	}
	env := output.Failure(command, output.Error{
		Code:    "source_preflight_failed",
		Message: fmt.Sprintf("%d UserForm spec/artifact issue(s) must be fixed before %s so push cannot import the wrong Designer-backed form", len(rendered), command),
		Source:  "xlflow",
		Phase:   "preflight",
	})
	env.Issues = rendered
	env.Logs = []string{"blocked before Excel automation because spec-driven UserForm artifacts are missing or inconsistent with src/forms/specs"}
	return a.write(env, output.ExitValidation)
}

func (a *app) runFormWritePreflight(command string, cfg config.Config, opts formWriteCommandOptions) error {
	targetForms := map[string]bool{opts.Spec.Form.Name: true}
	if err := a.runUserFormCodeSourcePreflight(command, cfg, targetForms); err != nil {
		return err
	}
	if cfg.UserForm.CodeSource == "sidecar" {
		if err := a.runSourcePreflight(command, cfg, "writing workbook forms", nil, buildUserFormSourcePathFilter(a.cwd, cfg, targetForms)); err != nil {
			return err
		}
	}
	return nil
}

func filterAnalysisFindings(findings []analyze.Finding, ignoredCodes map[string]bool) []analyze.Finding {
	if len(ignoredCodes) == 0 {
		return findings
	}
	filtered := make([]analyze.Finding, 0, len(findings))
	for _, finding := range findings {
		if ignoredCodes[finding.Code] {
			continue
		}
		filtered = append(filtered, finding)
	}
	return filtered
}

func (a *app) shouldRunSourcePreflight(cfg config.Config, opts excel.RunOptions) bool {
	if opts.Session {
		return true
	}
	workbook := cfg.Excel.Path
	if opts.WorkbookPath != "" {
		workbook = opts.WorkbookPath
	}
	if workbook == "" || cfg.Excel.Path == "" {
		return false
	}
	return strings.EqualFold(filepath.Clean(workbookArgPath(a.cwd, workbook)), filepath.Clean(workbookArgPath(a.cwd, cfg.Excel.Path)))
}

func absInt(v int) int {
	if v < 0 {
		return -v
	}
	return v
}

func cliLocationHasMeaningfulData(value any) bool {
	location := cliObjectMap(value)
	if len(location) == 0 {
		return false
	}
	if strings.TrimSpace(stringValueForCLI(location, "module")) != "" ||
		strings.TrimSpace(stringValueForCLI(location, "file")) != "" ||
		strings.TrimSpace(stringValueForCLI(location, "procedure")) != "" {
		return true
	}
	line := strings.TrimSpace(stringValueForCLI(location, "line"))
	return line != "" && line != "0"
}

func cliObjectMap(value any) map[string]any {
	if value == nil {
		return map[string]any{}
	}
	switch v := value.(type) {
	case map[string]any:
		return v
	case map[string]string:
		out := make(map[string]any, len(v))
		for key, item := range v {
			out[key] = item
		}
		return out
	default:
		data, err := json.Marshal(value)
		if err != nil {
			return map[string]any{}
		}
		var out map[string]any
		if err := json.Unmarshal(data, &out); err != nil {
			return map[string]any{}
		}
		if out == nil {
			return map[string]any{}
		}
		return out
	}
}

func stringValueForCLI(m map[string]any, key string) string {
	value, ok := m[key]
	if !ok || value == nil {
		return ""
	}
	return fmt.Sprint(value)
}

func boolValueForCLI(m map[string]any, key string) bool {
	value, ok := m[key]
	if !ok || value == nil {
		return false
	}
	switch v := value.(type) {
	case bool:
		return v
	case string:
		return strings.EqualFold(v, "true")
	default:
		return false
	}
}

func (a *app) skillCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "skill",
		Short: "Manage bundled AI agent skills",
	}
	cmd.AddCommand(a.skillInstallCommand())
	return cmd
}

func (a *app) skillInstallCommand() *cobra.Command {
	var agent string
	var target string
	var force bool
	cmd := &cobra.Command{
		Use:   "install",
		Short: "Install the bundled xlflow AI agent skill",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			opts, err := a.resolveSkillInstallOptions(agent, target, force)
			if err != nil {
				return a.writeFailure("skill install", output.ExitConfig, "skill_agent_required", err)
			}
			result, err := agentskill.Install(opts)
			if err != nil {
				return a.writeFailure("skill install", output.ExitConfig, "skill_install_failed", err)
			}
			env := output.New("skill install")
			env.Logs = []string{"installed xlflow skill to " + result.Path}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&agent, "agent", "", "skill provider target: agents, codex, claude, cursor, or gemini")
	cmd.Flags().StringVar(&target, "target", "", "directory that should receive the xlflow skill folder")
	cmd.Flags().BoolVar(&force, "force", false, "replace an existing xlflow skill")
	return cmd
}

func (a *app) resolveSkillInstallOptions(agent, target string, force bool) (agentskill.InstallOptions, error) {
	if agent == "" && target == "" {
		if a.json || !stdinIsTerminal() {
			return agentskill.InstallOptions{}, fmt.Errorf("--agent or --target is required when interactive selection is unavailable")
		}
		provider, err := runSkillSelector()
		if err != nil {
			return agentskill.InstallOptions{}, err
		}
		agent = provider.Name
	}
	return agentskill.InstallOptions{
		RootDir: a.cwd,
		Agent:   agent,
		Target:  target,
		Force:   force,
	}, nil
}

func stdinIsTerminal() bool {
	info, err := os.Stdin.Stat()
	if err != nil {
		return false
	}
	return info.Mode()&os.ModeCharDevice != 0
}

func stdoutIsTerminal() bool {
	info, err := os.Stdout.Stat()
	if err != nil {
		return false
	}
	return info.Mode()&os.ModeCharDevice != 0
}

func stderrIsTerminal() bool {
	info, err := os.Stderr.Stat()
	if err != nil {
		return false
	}
	return info.Mode()&os.ModeCharDevice != 0
}

func (a *app) stdoutWriter() io.Writer {
	if a.stdout != nil {
		return a.stdout
	}
	return os.Stdout
}

func (a *app) stderrWriter() io.Writer {
	if a.stderr != nil {
		return a.stderr
	}
	return os.Stderr
}

func (a *app) stdoutIsInteractive() bool {
	if a.stdoutTerminal != nil {
		return a.stdoutTerminal()
	}
	return stdoutIsTerminal()
}

func (a *app) stderrIsInteractive() bool {
	if a.stderrTerminal != nil {
		return a.stderrTerminal()
	}
	return stderrIsTerminal()
}

func (a *app) loadConfig(command string) (config.Config, error) {
	cfg, err := config.Load(a.cwd)
	if err != nil && errors.Is(err, config.ErrInvalidExcelBridge) && a.hasValidBridgeOverride() {
		cfg, err = config.LoadAllowInvalidExcelBridge(a.cwd)
	}
	if err != nil {
		return cfg, a.writeFailure(command, output.ExitConfig, "config_error", err)
	}
	if len(cfg.Warnings) > 0 {
		a.configWarnings = append(a.configWarnings, cfg.Warnings...)
	}
	return cfg, nil
}

func (a *app) loadLSPConfig() (config.Config, error) {
	cfg, err := config.Load(a.cwd)
	if err != nil && errors.Is(err, config.ErrInvalidExcelBridge) && a.hasValidBridgeOverride() {
		cfg, err = config.LoadAllowInvalidExcelBridge(a.cwd)
	}
	if err != nil {
		if errors.Is(err, config.ErrConfigNotFound) {
			return cfg, nil
		}
		return cfg, a.writeFailure("lsp", output.ExitConfig, "config_error", err)
	}
	if len(cfg.Warnings) > 0 {
		a.configWarnings = append(a.configWarnings, cfg.Warnings...)
	}
	return cfg, nil
}

func (a *app) hasValidBridgeOverride() bool {
	for _, candidate := range []string{a.bridge, os.Getenv(excelbridge.EnvBridge)} {
		if strings.TrimSpace(candidate) == "" {
			continue
		}
		_, err := excelbridge.ParseMode(candidate)
		return err == nil
	}
	return false
}

func (a *app) excelRunner() excel.Runner {
	return excel.Runner{RootDir: a.cwd, BridgeMode: a.bridge, Coordination: a.coordination, BorrowedLeases: a.activeLeases}
}

func (a *app) excelRunnerForConfig(cfg config.Config) excel.Runner {
	return excel.Runner{RootDir: a.cwd, BridgeMode: a.bridge, ConfigBridgeMode: cfg.Excel.Bridge, Coordination: a.coordination, BorrowedLeases: a.activeLeases}
}

func (a *app) writeScaffoldWelcome(command string, skipUpdateCheck bool) error {
	opts := a.outputOptions()
	if !shouldRenderScaffoldWelcome(command, opts) {
		return nil
	}
	_, err := fmt.Fprint(a.stdoutWriter(), renderScaffoldWelcome(a.scaffoldWelcomeModel(skipUpdateCheck), opts.Color))
	return err
}

func (a *app) writeFailure(command string, code int, errCode string, err error) error {
	env := output.Failure(command, output.Error{Code: errCode, Message: err.Error()})
	a.addConfigWarnings(&env)
	if writeErr := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); writeErr != nil {
		return output.WithExitCode(code, writeErr)
	}
	return output.WithExitCode(code, err)
}

func (a *app) writeUnsupportedWorkbookFormat(command string, err workbookformat.UnsupportedError) error {
	env := output.Failure(command, output.Error{
		Code:    workbookformat.UnsupportedErrorCode,
		Message: err.Error(),
	})
	env.Workbook = map[string]any{
		"format":     workbookformat.Format(err.Extension),
		"capability": err.Capability,
	}
	a.addConfigWarnings(&env)
	if writeErr := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); writeErr != nil {
		return output.WithExitCode(output.ExitConfig, writeErr)
	}
	return output.WithExitCode(output.ExitConfig, err)
}

func (a *app) rejectUnsupportedFileInspectWorkbook(command, capability, workbookPath string) error {
	if err := workbookformat.ValidateFileInspectWorkbook(workbookPath, capability); err != nil {
		var unsupported workbookformat.UnsupportedError
		if errors.As(err, &unsupported) {
			return a.writeUnsupportedWorkbookFormat(command, unsupported)
		}
		return a.writeFailure(command, output.ExitConfig, "inspect_args_invalid", err)
	}
	return nil
}

func (a *app) writeFormSpecFailure(command string, specErr *forms.SpecError) error {
	env := output.Failure(command, output.Error{Code: specErr.Code, Message: specErr.Message})
	a.addConfigWarnings(&env)
	spec := map[string]any{}
	if specErr.Path != "" {
		spec["path"] = specErr.Path
	}
	if specErr.Format != "" {
		spec["format"] = specErr.Format
	}
	if specErr.Line > 0 {
		spec["line"] = specErr.Line
	}
	if specErr.Column > 0 {
		spec["column"] = specErr.Column
	}
	if specErr.Field != "" {
		spec["field"] = specErr.Field
	}
	if specErr.Suggestion != "" {
		spec["suggestion"] = specErr.Suggestion
	}
	if len(specErr.Issues) > 0 {
		spec["issues"] = formValidationIssuesPayload(specErr.Issues)
	}
	if len(spec) > 0 {
		env.Spec = spec
	}
	if writeErr := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); writeErr != nil {
		return output.WithExitCode(output.ExitValidation, writeErr)
	}
	return output.WithExitCode(output.ExitValidation, specErr)
}

func appendFormSpecValidationWarnings(env *output.Envelope, issues []forms.ValidationIssue) {
	if env == nil || len(issues) == 0 {
		return
	}
	warnings := anySlice(env.Warnings)
	for _, issue := range issues {
		if issue.Severity != forms.SeverityWarning {
			continue
		}
		warnings = append(warnings, formValidationIssuePayload(issue))
	}
	if len(warnings) > 0 {
		env.Warnings = warnings
	}
}

func formValidationIssuesPayload(issues []forms.ValidationIssue) []map[string]any {
	out := make([]map[string]any, 0, len(issues))
	for _, issue := range issues {
		out = append(out, formValidationIssuePayload(issue))
	}
	return out
}

func formValidationIssuePayload(issue forms.ValidationIssue) map[string]any {
	item := map[string]any{}
	if issue.Code != "" {
		item["code"] = issue.Code
	}
	if issue.Severity != "" {
		item["severity"] = string(issue.Severity)
	}
	if issue.Message != "" {
		item["message"] = issue.Message
	}
	if issue.Field != "" {
		item["field"] = issue.Field
	}
	if issue.Suggestion != "" {
		item["suggestion"] = issue.Suggestion
	}
	if issue.Support != "" {
		item["support"] = string(issue.Support)
	}
	return item
}

func (a *app) write(env output.Envelope, code int) error {
	return a.writeWithOutputOptions(env, code, a.outputOptions())
}

func (a *app) writeWithOutputOptions(env output.Envelope, code int, opts output.Options) error {
	a.addConfigWarnings(&env)
	appendVBAObjectModelAccessMessages(&env)
	if err := output.WriteWithOptions(a.stdoutWriter(), env, opts); err != nil {
		return output.WithExitCode(code, err)
	}
	if code != output.ExitSuccess {
		return output.WithExitCode(code, fmt.Errorf("%s failed", env.Command))
	}
	return nil
}

func (a *app) addConfigWarnings(env *output.Envelope) {
	if env == nil || len(a.configWarnings) == 0 {
		return
	}
	warnings := anySlice(env.Warnings)
	for _, warning := range a.configWarnings {
		warnings = append(warnings, warning)
	}
	env.Warnings = warnings
}

func (a *app) outputOptions() output.Options {
	return a.outputOptionsWithVerbose(false)
}

func (a *app) outputOptionsWithVerbose(verbose bool) output.Options {
	interactive := !a.json && a.stdoutIsInteractive()
	return output.Options{
		JSON:        a.json,
		Interactive: interactive,
		Color:       interactive,
		Verbose:     verbose,
	}
}

func (a *app) runExcelWithProgress(label string, opts excel.CommandOptions, fn func() (output.Envelope, int, error)) (output.Envelope, int, error) {
	var env output.Envelope
	var code int
	err := a.withExcelProgress(label, opts, func() error {
		var runErr error
		env, code, runErr = fn()
		return runErr
	})
	return env, code, err
}

func (a *app) withExcelProgress(label string, opts excel.CommandOptions, fn func() error) error {
	if !opts.Progress {
		return fn()
	}
	return a.withSpinner(label, fn)
}

func (a *app) withSpinner(label string, fn func() error) error {
	if !a.stderrIsInteractive() {
		_, _ = fmt.Fprintf(a.stderrWriter(), "xlflow: %s...\n", label)
		return fn()
	}
	return runSpinner(a.stderrWriter(), label, fn)
}
