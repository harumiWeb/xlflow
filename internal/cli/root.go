package cli

import (
	"bytes"
	"cmp"
	"context"
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

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"

	"github.com/harumiWeb/xlflow/internal/agentskill"
	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/backup"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/diff"
	"github.com/harumiWeb/xlflow/internal/excel"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	"github.com/harumiWeb/xlflow/internal/gui"
	workbookinspect "github.com/harumiWeb/xlflow/internal/inspect"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/output"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
	"github.com/harumiWeb/xlflow/internal/project"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vbafmt"
)

type app struct {
	json           bool
	bridge         string
	cwd            string
	rawArgs        []string
	stdout         io.Writer
	stderr         io.Writer
	stdoutTerminal func() bool
	stderrTerminal func() bool
	configWarnings []map[string]any
	buildInfo      BuildInfo
	updateChecker  releaseChecker
}

type BuildInfo struct {
	Version string `json:"version"`
	Commit  string `json:"commit"`
	Date    string `json:"date"`
}

type versionFeature struct {
	Name        string `json:"name"`
	Description string `json:"description"`
}

type versionScriptInfo struct {
	Command string `json:"command"`
	Source  string `json:"source"`
	Path    string `json:"path,omitempty"`
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
	Scripts        []versionScriptInfo   `json:"scripts,omitempty"`
	Features       []versionFeature      `json:"features,omitempty"`
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
	return root.Execute()
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
			return a.delegateWSLCommand(cmd)
		},
	}
	root.PersistentFlags().BoolVar(&a.json, "json", false, "write machine-readable JSON output")
	root.PersistentFlags().StringVar(&a.bridge, "bridge", "", "Excel bridge mode: auto, dotnet; powershell is deprecated explicit opt-in until v0.16.0")
	root.AddCommand(
		a.newCommand(),
		a.initCommand(),
		a.doctorCommand(),
		a.attachCommand(),
		a.backupCommand(),
		a.listCommand(),
		a.formCommand(),
		a.pullCommand(),
		a.pushCommand(),
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
		a.diffCommand(),
		a.inspectCommand(),
		a.inspectGUICommand(),
		a.lintCommand(),
		a.fmtCommand(),
		a.analyzeCommand(),
		a.checkCommand(),
		a.generateCommand(),
		a.moduleCommand(),
		a.skillCommand(),
		a.versionCommand(),
		a.processCommand(),
	)
	return root
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

func (a *app) versionPayload(verbose bool) any {
	info := a.buildInfo.withDefaults()
	if !verbose {
		return info
	}
	payload := versionVerbosePayload{
		BuildInfo:      info,
		ExecutablePath: resolvedExecutablePath(),
		Features:       supportedVersionFeatures(),
		Scripts:        resolvedVersionScripts(a.cwd),
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

func resolvedVersionScripts(root string) []versionScriptInfo {
	commands := []string{"run", "push", "pull", "macros", "test", "session", "list", "inspect-form", "form-write", "export-image", "form-export-image", "edit", "process"}
	scripts := make([]versionScriptInfo, 0, len(commands))
	for _, command := range commands {
		info := versionScriptInfo{Command: command, Source: "embedded"}
		if path, ok := resolvedVersionScriptPath(root, command); ok {
			info.Source = "override"
			info.Path = path
		}
		scripts = append(scripts, info)
	}
	return scripts
}

func resolvedVersionScriptPath(root, command string) (string, bool) {
	name := command + ".ps1"
	candidates := []string{}
	if root != "" {
		candidates = append(candidates, filepath.Join(root, "scripts", name))
	}
	if _, file, _, ok := runtime.Caller(0); ok {
		candidates = append(candidates, filepath.Join(filepath.Dir(file), "scripts", name))
	}
	if exe, err := os.Executable(); err == nil {
		candidates = append(candidates, filepath.Join(filepath.Dir(exe), "scripts", name))
	}
	for _, candidate := range candidates {
		clean := filepath.Clean(candidate)
		if _, err := os.Stat(clean); err == nil {
			return clean, true
		}
	}
	return "", false
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
	cmd.AddCommand(a.backupListCommand())
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
			records, err := backup.List(a.cwd, workbookPath)
			if err != nil {
				return a.writeFailure("backup list", output.ExitEnvironment, "backup_list_failed", err)
			}
			env := output.New("backup list")
			env.Backups = renderBackupRecords(a.cwd, records)
			env.Logs = []string{fmt.Sprintf("found %d backup(s)", len(records))}
			return a.write(env, output.ExitSuccess)
		},
	}
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
	cmd.AddCommand(a.formSnapshotCommand(), a.formBuildCommand(), a.formApplyCommand(), a.formExportImageCommand())
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
		Short: "Create a new xlflow project and macro workbook",
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

	cmd := &cobra.Command{
		Use:   "init <workbook>",
		Short: "Create an xlflow project from an existing macro workbook",
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
				result, err := project.Init(a.cwd, args[0])
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
				if withModule {
					env.Source = map[string]any{"created": installedModules.Created}
					env.Logs = append(env.Logs,
						fmt.Sprintf("installed %d bundled helper module(s) into source", len(installedModules.Created)),
						"pushed bundled helper modules to workbook",
					)
				}
				if withSkill {
					env.Logs = append(env.Logs, "installed xlflow skill to "+skillResult.Path)
				}
				return a.write(env, output.ExitSuccess)
			}
		},
	}
	cmd.Flags().BoolVar(&withSkill, "with-skill", false, "install the bundled xlflow AI agent skill")
	cmd.Flags().BoolVar(&withModule, "with-module", false, "install bundled xlflow helper modules and push them to the workbook")
	cmd.Flags().StringVar(&skillAgent, "agent", "", "skill provider target: agents, codex, claude, cursor, or gemini")
	cmd.Flags().BoolVar(&noUpdateCheck, "no-update-check", false, "skip the interactive GitHub release update check during project scaffolding")
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

func (a *app) doctorCommand() *cobra.Command {
	var checkWorkbook bool
	cmd := &cobra.Command{
		Use:   "doctor",
		Short: "Diagnose Excel COM and VBIDE access",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := a.loadConfig("doctor")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Checking Excel automation", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).DoctorWithOptions(cfg, excel.DoctorOptions{
					CheckWorkbook: checkWorkbook,
					Keepalive:     commandOpts,
				})
				return runErr
			})
			if err != nil {
				return err
			}
			boundaries, analyzeErr := gui.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
			if analyzeErr == nil && len(boundaries) > 0 {
				env.GUIBoundaries = boundaries
				env.Diagnostics = withGUIBoundarySummary(env.Diagnostics, boundaries)
				env.Logs = append(env.Logs, fmt.Sprintf("detected %d GUI boundary candidate(s) in source", len(boundaries)))
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&checkWorkbook, "workbook", false, "open the configured workbook as part of doctor diagnostics")
	return cmd
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
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&active, "active", false, "attach to the active Excel workbook")
	return cmd
}

func (a *app) pullCommand() *cobra.Command {
	var session bool
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
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
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
		Short: "Manage bundled xlflow helper modules",
	}
	cmd.AddCommand(a.moduleInstallCommand())
	return cmd
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

func renderBackupRecords(rootDir string, records []backup.Record) []map[string]any {
	rendered := make([]map[string]any, 0, len(records))
	for _, record := range records {
		rendered = append(rendered, map[string]any{
			"id":         record.ID,
			"created_at": record.CreatedAt.Format(time.RFC3339),
			"reason":     record.Reason,
			"workbook":   displayPath(rootDir, record.OriginalWorkbookPath),
			"path":       displayPath(rootDir, record.BackupFileAbsPath),
		})
	}
	return rendered
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
	text := strings.TrimSpace(string(out))
	if text == "" {
		return false, nil
	}
	if strings.Contains(strings.ToLower(text), "no tasks are running") {
		return false, nil
	}
	return true, nil
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
	left, leftErr := canonicalPath(a)
	right, rightErr := canonicalPath(b)
	if leftErr == nil && rightErr == nil {
		return strings.EqualFold(left, right)
	}
	if leftInfo, err := os.Stat(a); err == nil {
		if rightInfo, err := os.Stat(b); err == nil {
			return os.SameFile(leftInfo, rightInfo)
		}
	}
	return strings.EqualFold(filepath.Clean(a), filepath.Clean(b))
}

func canonicalPath(path string) (string, error) {
	if strings.TrimSpace(path) == "" {
		return "", fmt.Errorf("path is required")
	}
	absPath, err := filepath.Abs(path)
	if err != nil {
		return "", err
	}
	resolvedPath, err := filepath.EvalSymlinks(absPath)
	if err == nil {
		return filepath.Clean(resolvedPath), nil
	}
	if errors.Is(err, os.ErrNotExist) {
		return filepath.Clean(absPath), nil
	}
	return "", err
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
		cmd := &cobra.Command{
			Use:   action,
			Short: action + " the xlflow Excel session",
			Args:  cobra.NoArgs,
			RunE: func(cmd *cobra.Command, args []string) error {
				cfg, err := a.loadConfig("session")
				if err != nil {
					return err
				}
				env, code, err := a.excelRunnerForConfig(cfg).Session(cfg, action)
				if err != nil {
					return err
				}
				return a.write(env, code)
			},
		}
		session.AddCommand(cmd)
	}
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
			warnings, hints := buildStatusWarningsAndHints(sessionState, statePayload)
			env.Warnings = warnings
			env.Hints = hints
			env.Logs = []string{"status reported"}
			return a.write(env, output.ExitSuccess)
		},
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
	session["save_required"] = saveRequired
	session["live_newer_than_disk"] = saveRequired
	if saveRequired {
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
			if env.Status == output.StatusFailed && env.Error != nil && env.Error.Code == "macro_failed" && env.Error.Phase == "invoke_macro" {
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
	return cmd
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
		a.editRowsCommand(),
		a.editColumnsCommand(),
	)
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
	var msgBoxLiterals []string
	var inputBoxLiterals []string
	var fileDialogLiterals []string
	var session bool
	var uiStream bool
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
			err = a.withExcelProgress("Running VBA tests", commandOpts, func() error {
				var runErr error
				env, code, runErr = a.excelRunnerForConfig(cfg).TestWithOptions(cfg, filter, excel.TestOptions{Session: session, Keepalive: commandOpts, RuntimeMode: runtime.Mode, RuntimeSource: runtime.Source, UIResponses: excel.UIResponses{MsgBox: msgBoxResponses, Input: inputResponses, FileDialog: fileDialogResponses}, DebugStream: excel.DebugStreamOptions{Enabled: true}, UIStream: excel.UIStreamOptions{Enabled: uiStream, RedactInput: true}, ModuleFilter: moduleFilter, TagFilter: tagFilter})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&filter, "filter", "", "run only the test whose procedure name exactly matches filter")
	cmd.Flags().StringVar(&moduleFilter, "module", "", "run only tests in the module whose name exactly matches filter")
	cmd.Flags().StringVar(&tagFilter, "tag", "", "run only tests tagged with the given tag")
	cmd.Flags().StringArrayVar(&msgBoxLiterals, "msgbox", nil, "provide a scripted MsgBox response as dialog-id=result")
	cmd.Flags().StringArrayVar(&inputBoxLiterals, "inputbox", nil, "provide a scripted InputBox response as dialog-id=value")
	cmd.Flags().StringArrayVar(&fileDialogLiterals, "filedialog", nil, "provide a scripted file dialog response as kind:dialog-id=path or kind:dialog-id=@cancel")
	cmd.Flags().BoolVar(&uiStream, "ui-stream", false, "stream headless XlflowUI dialog events to stderr in real time")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
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
	cmd.Flags().BoolVar(&includeMembers, "include-members", false, "include member calls (included by default)")
	cmd.Flags().BoolVar(&includeBuiltins, "include-builtins", false, "include built-in-looking calls (included by default)")
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
			workbook, err := workbookinspect.Workbook(workbookArgPath(a.cwd, cfg.Excel.Path))
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
			sheets, err := workbookinspect.Sheets(workbookArgPath(a.cwd, cfg.Excel.Path))
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
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
			cell, err := workbookinspect.Cell(workbookArgPath(a.cwd, cfg.Excel.Path), selector.Sheet, selector.Address)
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
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
			return nil, fmt.Errorf("%w: %d UserForm sidecar(s) under %s carry Attribute VB_* header lines; a sidecar must hold code-behind only (starting at Option Explicit)", packpkg.ErrAmbiguousLayout, len(issues), filepath.Join(cfg.Src.Forms, "code"))
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
		if _, err := os.Stat(filepath.Join(formsDir, formName+".frm")); err != nil {
			if errors.Is(err, os.ErrNotExist) {
				return fmt.Errorf("%w: UserForm sidecar %q has no matching %s.frm", packpkg.ErrAmbiguousLayout, formName, formName)
			}
			return err
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
	switch strings.ToLower(filepath.Ext(path)) {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return nil
	default:
		return fmt.Errorf("unsupported extension %q; expected .xlsx, .xlsm, .xltx, or .xltm", filepath.Ext(path))
	}
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
			cfg, err := a.loadConfig("fmt")
			if err != nil {
				return err
			}
			opts := vbafmt.FmtOptions{
				Write:       write,
				Check:       check,
				Diff:        diff,
				Paths:       args,
				Root:        a.cwd,
				Cfg:         cfg,
				LineNumbers: lineNumberMode,
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
	formatted, err := vbafmt.FormatText(input, looksLikeClassModule(input))
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
	if _, ok := diag["kind"]; !ok {
		diag["kind"] = "runtime"
	}
	if _, ok := diag["likely_cause"]; !ok {
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
	if err == nil && env.Error != nil {
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
	return excel.Runner{RootDir: a.cwd, BridgeMode: a.bridge}
}

func (a *app) excelRunnerForConfig(cfg config.Config) excel.Runner {
	return excel.Runner{RootDir: a.cwd, BridgeMode: a.bridge, ConfigBridgeMode: cfg.Excel.Bridge}
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
	if len(spec) > 0 {
		env.Spec = spec
	}
	if writeErr := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); writeErr != nil {
		return output.WithExitCode(output.ExitValidation, writeErr)
	}
	return output.WithExitCode(output.ExitValidation, specErr)
}

func (a *app) write(env output.Envelope, code int) error {
	return a.writeWithOutputOptions(env, code, a.outputOptions())
}

func (a *app) writeWithOutputOptions(env output.Envelope, code int, opts output.Options) error {
	a.addConfigWarnings(&env)
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
