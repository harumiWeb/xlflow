package cli

import (
	"bytes"
	"cmp"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
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
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	"github.com/harumiWeb/xlflow/internal/gui"
	workbookinspect "github.com/harumiWeb/xlflow/internal/inspect"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/project"
)

type app struct {
	json           bool
	cwd            string
	stdout         io.Writer
	stderr         io.Writer
	stdoutTerminal func() bool
	stderrTerminal func() bool
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
	a := &app{cwd: cwd, buildInfo: info.withDefaults()}
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
	}
	root.PersistentFlags().BoolVar(&a.json, "json", false, "write machine-readable JSON output")
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
		a.rollbackCommand(),
		a.sessionCommand(),
		a.saveCommand(),
		a.runnerCommand(),
		a.traceCommand(),
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
		a.analyzeCommand(),
		a.checkCommand(),
		a.moduleCommand(),
		a.skillCommand(),
		a.versionCommand(),
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
		{Name: "trace-lifecycle", Description: "Enable, disable, inspect, or temporarily inject XlflowTrace support."},
		{Name: "range-image-export", Description: "Export a worksheet range to a PNG image for visual verification."},
		{Name: "form-image-export", Description: "Export a runtime UserForm to a PNG image for visual verification."},
		{Name: "form-build-overwrite", Description: "Create or replace Designer-backed UserForms from persisted xlflow.userform specs."},
		{Name: "workbook-edit-helpers", Description: "Mutate a live session workbook for agent-driven test setup, event triggering, and visual tuning."},
	}
}

func resolvedVersionScripts(root string) []versionScriptInfo {
	commands := []string{"run", "push", "pull", "macros", "test", "trace", "session", "list", "inspect-form", "form-write", "export-image", "form-export-image", "edit"}
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.MacrosWithOptions(cfg, excel.SessionCommandOptions{Session: session, Keepalive: commandOpts})
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
		Short: "Write a strict designer UserForm snapshot spec to a file",
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
				scriptEnv, code, runErr = excel.Runner{RootDir: a.cwd}.InspectForm(cfg, opts.Inspect)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.FormWrite(cfg, excel.FormWriteOptions{
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.FormWrite(cfg, excel.FormWriteOptions{
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.FormExportImage(cfg, opts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.ListForms(cfg, excel.SessionCommandOptions{Session: session, Keepalive: commandOpts})
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.UIButtonAdd(cfg, built, commandOpts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.UIButtonList(cfg, opts, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&opts.Sheet, "sheet", "", "worksheet name")
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.UIButtonRemove(cfg, built, commandOpts)
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
						return excel.Runner{RootDir: a.cwd}.New(path, runOpts)
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
		return excel.Runner{RootDir: a.cwd}.PushWithOptions(cfg, excel.PushOptions{
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
		return excel.Runner{RootDir: a.cwd}.PullWithOptions(cfg, excel.SessionCommandOptions{
			Keepalive: keepaliveOpts,
		})
	})
}

func (a *app) doctorCommand() *cobra.Command {
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.Doctor(cfg, commandOpts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.Attach(cfg, active, commandOpts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.PullWithOptions(cfg, excel.SessionCommandOptions{Session: session, Keepalive: commandOpts})
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
		env, code, runErr = excel.Runner{RootDir: a.cwd}.PushWithOptions(cfg, pushOpts)
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

func (a *app) rollbackBlockedBySession(workbookPath string) (bool, error) {
	metadataPath := filepath.Join(a.cwd, ".xlflow", "session.json")
	body, err := os.ReadFile(metadataPath)
	if errors.Is(err, os.ErrNotExist) {
		return false, nil
	}
	if err != nil {
		return false, err
	}
	body = bytes.TrimPrefix(body, []byte{0xEF, 0xBB, 0xBF})
	var metadata sessionMetadata
	if err := json.Unmarshal(body, &metadata); err != nil {
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

func buildRunOptions(cfg config.Config, macro, input string, argLiterals []string, msgBoxLiterals []string, inputBoxLiterals []string, fileDialogLiterals []string, save bool, saveAs string, trace bool, headless bool, interactive bool, direct bool, fast bool, diagnostic bool, diagnosticExplicit bool, guiCompileErrors bool, session bool, timeout time.Duration, commandOpts excel.CommandOptions) (excel.RunOptions, error) {
	return buildRunOptionsWithUIStream(cfg, macro, input, argLiterals, msgBoxLiterals, inputBoxLiterals, fileDialogLiterals, save, saveAs, trace, headless, interactive, direct, fast, diagnostic, diagnosticExplicit, guiCompileErrors, session, timeout, commandOpts, false)
}

func buildRunOptionsWithUIStream(cfg config.Config, macro, input string, argLiterals []string, msgBoxLiterals []string, inputBoxLiterals []string, fileDialogLiterals []string, save bool, saveAs string, trace bool, headless bool, interactive bool, direct bool, fast bool, diagnostic bool, diagnosticExplicit bool, guiCompileErrors bool, session bool, timeout time.Duration, commandOpts excel.CommandOptions, uiStream bool) (excel.RunOptions, error) {
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
	if direct && trace {
		return excel.RunOptions{}, fmt.Errorf("--direct cannot be combined with --trace")
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
		case "int", "bool":
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
		Trace:               trace,
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
				env, code, err := excel.Runner{RootDir: a.cwd}.Session(cfg, action)
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
			env, code, err := excel.Runner{RootDir: a.cwd}.SaveSession(cfg, excel.SessionCommandOptions{Session: session})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
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
				env, code, err := excel.Runner{RootDir: a.cwd}.RunnerModule(cfg, action)
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
	var saveAs string
	var trace bool
	var headless bool
	var interactive bool
	var direct bool
	var fast bool
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
			macro := ""
			if len(args) == 1 {
				macro = args[0]
			}
			var opts excel.RunOptions
			commandOpts := buildCommandOptions(a.stderrWriter())
			commandOpts.Progress = false
			if uiStream {
				opts, err = buildRunOptionsWithUIStream(cfg, macro, input, argLiterals, msgBoxLiterals, inputBoxLiterals, fileDialogLiterals, save, saveAs, trace, headless, interactive, direct, fast, diagnostic, cmd.Flags().Changed("diagnostic"), guiCompileErrors, session, timeout, commandOpts, true)
			} else {
				opts, err = buildRunOptions(cfg, macro, input, argLiterals, msgBoxLiterals, inputBoxLiterals, fileDialogLiterals, save, saveAs, trace, headless, interactive, direct, fast, diagnostic, cmd.Flags().Changed("diagnostic"), guiCompileErrors, session, timeout, commandOpts)
			}
			if err != nil {
				return a.writeFailure("run", output.ExitConfig, "run_args_invalid", err)
			}
			if a.shouldRunSourcePreflight(cfg, opts) {
				if err := a.runSourcePreflight("run", cfg, "running macros", ignoredRunPreflightAnalysisCodes(opts), nil); err != nil {
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.Run(cfg, opts)
				return runErr
			}
			err = a.withSpinner("Running macro", run)
			if err != nil {
				return err
			}
			if env.Status == output.StatusFailed && env.Error != nil && env.Error.Code == "macro_failed" && env.Error.Phase == "invoke_macro" {
				env.RunDiagnostic = a.buildRunDiagnostic(cfg, env)
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringArrayVar(&argLiterals, "arg", nil, "pass a typed macro argument such as string:hello, int:7, or bool:true")
	cmd.Flags().StringArrayVar(&msgBoxLiterals, "msgbox", nil, "provide a scripted MsgBox response as dialog-id=result")
	cmd.Flags().StringArrayVar(&inputBoxLiterals, "inputbox", nil, "provide a scripted InputBox response as dialog-id=value")
	cmd.Flags().StringArrayVar(&fileDialogLiterals, "filedialog", nil, "provide a scripted file dialog response as kind:dialog-id=path or kind:dialog-id=@cancel")
	cmd.Flags().StringVar(&input, "input", "", "override workbook path for this run")
	cmd.Flags().BoolVar(&save, "save", false, "save the opened workbook after a successful run")
	cmd.Flags().StringVar(&saveAs, "save-as", "", "write the successful workbook result to a new path")
	cmd.Flags().BoolVar(&trace, "trace", false, "collect XlflowTrace log events during the run")
	cmd.Flags().BoolVar(&headless, "headless", false, "reject GUI interaction boundaries before running the macro")
	cmd.Flags().BoolVar(&interactive, "interactive", false, "run with Excel visible and alerts enabled for human interaction")
	cmd.Flags().BoolVar(&direct, "direct", false, "run an argument-free macro without injecting a temporary harness")
	cmd.Flags().BoolVar(&fast, "fast", false, "use development-oriented fast run defaults")
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.ExportImage(cfg, opts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.EditCell(cfg, opts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.EditRange(cfg, opts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.EditRows(cfg, opts)
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.EditColumns(cfg, opts)
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

func (a *app) traceCommand() *cobra.Command {
	trace := &cobra.Command{
		Use:   "trace",
		Short: "Manage workbook trace logging support",
	}
	trace.AddCommand(
		a.traceLifecycleCommand("enable", "Enable the XlflowTrace VBA module"),
		a.traceLifecycleCommand("disable", "Disable the XlflowTrace VBA module"),
		a.traceLifecycleCommand("status", "Report XlflowTrace status"),
		a.traceLifecycleCommand("clean", "Remove xlflow trace log files"),
		a.traceInjectCommand(),
	)
	return trace
}

func (a *app) traceInjectCommand() *cobra.Command {
	cmd := a.traceLifecycleCommand("inject", "Deprecated alias for trace enable")
	cmd.Use = "inject [workbook]"
	return cmd
}

func (a *app) traceLifecycleCommand(action, short string) *cobra.Command {
	var force bool
	var session bool
	cmd := &cobra.Command{
		Use:   action + " [workbook]",
		Short: short,
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			commandOpts := buildCommandOptions(a.stderrWriter())
			cfg, err := config.Load(a.cwd)
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			if err != nil {
				if workbook == "" && action != "clean" {
					return a.writeFailure("trace", output.ExitConfig, "config_error", err)
				}
				cfg = config.Default()
			}
			traceAction := action
			if traceAction == "inject" {
				traceAction = "enable"
			}
			var env output.Envelope
			var code int
			label := "Managing trace module"
			if traceAction == "clean" {
				label = "Cleaning trace logs"
			}
			err = a.withExcelProgress(label, commandOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.Trace(cfg, excel.TraceOptions{Action: traceAction, Workbook: workbook, Force: force, Session: session}, commandOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	if action == "disable" {
		cmd.Flags().BoolVar(&force, "force", false, "remove modified trace helper source")
	}
	if action != "clean" {
		cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	}
	return cmd
}

func (a *app) testCommand() *cobra.Command {
	var filter string
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
			commandOpts.Progress = false
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
				env, code, runErr = excel.Runner{RootDir: a.cwd}.TestWithOptions(cfg, filter, excel.TestOptions{Session: session, Keepalive: commandOpts, RuntimeMode: runtime.Mode, RuntimeSource: runtime.Source, UIResponses: excel.UIResponses{MsgBox: msgBoxResponses, Input: inputResponses, FileDialog: fileDialogResponses}, DebugStream: excel.DebugStreamOptions{Enabled: true}, UIStream: excel.UIStreamOptions{Enabled: uiStream, RedactInput: true}})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&filter, "filter", "", "run only the test whose procedure name exactly matches filter")
	cmd.Flags().StringArrayVar(&msgBoxLiterals, "msgbox", nil, "provide a scripted MsgBox response as dialog-id=result")
	cmd.Flags().StringArrayVar(&inputBoxLiterals, "inputbox", nil, "provide a scripted InputBox response as dialog-id=value")
	cmd.Flags().StringArrayVar(&fileDialogLiterals, "filedialog", nil, "provide a scripted file dialog response as kind:dialog-id=path or kind:dialog-id=@cancel")
	cmd.Flags().BoolVar(&uiStream, "ui-stream", false, "stream headless XlflowUI dialog events to stderr in real time")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	return cmd
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
	)
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
				scriptEnv, code, runErr = excel.Runner{RootDir: a.cwd}.InspectForm(cfg, opts)
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
	inspectOpts.StrictDesigner = true
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
		scriptEnv, code, runErr = excel.Runner{RootDir: a.cwd}.Inspect(cfg, opts)
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
	if strings.TrimSpace(statusWorkbookPath) == "" || !strings.EqualFold(filepath.Clean(statusWorkbookPath), filepath.Clean(workbookPath)) {
		return target, session, nil
	}
	active := boolValueForCLI(status, "active") || (boolValueForCLI(status, "running") && boolValueForCLI(status, "workbook_open"))
	dirty := boolValueForCLI(status, "dirty")
	saveRequired := boolValueForCLI(status, "save_required") || boolValueForCLI(status, "needs_save")
	session["active"] = active
	session["dirty"] = dirty
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
	if !active || !dirty {
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
	env, _, err := excel.Runner{RootDir: a.cwd}.Session(cfg, "status")
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
			issues, err := lint.Linter{RootDir: a.cwd, Config: cfg}.Run()
			if err != nil {
				return a.writeFailure("lint", output.ExitEnvironment, "lint_failed", err)
			}
			env := output.New("lint")
			env.Issues = issues
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
			findings, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
			if err != nil {
				return a.writeFailure("analyze", output.ExitEnvironment, "analyze_failed", err)
			}
			env := output.New("analyze")
			env.Analysis = findings
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
			issues, err := lint.Linter{RootDir: a.cwd, Config: cfg}.Run()
			if err != nil {
				return a.writeFailure("check", output.ExitEnvironment, "lint_failed", err)
			}
			check["lint"] = map[string]any{"status": statusForCount(len(issues)), "count": len(issues)}
			findings, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
			if err != nil {
				return a.writeFailure("check", output.ExitEnvironment, "analyze_failed", err)
			}
			check["analyze"] = map[string]any{"status": statusForCount(len(findings)), "count": len(findings)}
			var doctor output.Envelope
			var doctorCode int
			err = a.withExcelProgress("Checking Excel automation", commandOpts, func() error {
				var runErr error
				doctor, doctorCode, runErr = excel.Runner{RootDir: a.cwd}.Doctor(cfg, commandOpts)
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
		diag["suggestion"] = "Inspect the failing procedure and rerun with --trace if the last successful step is unclear."
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
	if trace := cliObjectMap(env.Trace); len(trace) > 0 {
		events := cliListOfObjects(trace["events"])
		if len(events) > 0 {
			last := events[len(events)-1]
			diag["trace_context"] = map[string]any{
				"last_event": stringValueForCLI(last, "message"),
				"timestamp":  stringValueForCLI(last, "timestamp"),
			}
		}
	}
	return diag
}

func (a *app) runSourcePreflight(command string, cfg config.Config, action string, ignoredAnalysisCodes map[string]bool, pathFilter func(string) bool) error {
	issues, err := lint.Linter{RootDir: a.cwd, Config: cfg, PathFilter: pathFilter}.Run()
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "lint_failed", err)
	}
	findings, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg, PathFilter: pathFilter}.Run()
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "analyze_failed", err)
	}
	blockingIssues := lint.PushBlockingIssues(issues)
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

func ignoredRunPreflightAnalysisCodes(opts excel.RunOptions) map[string]bool {
	if !opts.Trace {
		return nil
	}
	return map[string]bool{
		"VBA105": true,
		"VBA106": true,
	}
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

func cliListOfObjects(value any) []map[string]any {
	switch v := value.(type) {
	case []map[string]any:
		return v
	case []any:
		out := make([]map[string]any, 0, len(v))
		for _, item := range v {
			out = append(out, cliObjectMap(item))
		}
		return out
	default:
		data, err := json.Marshal(value)
		if err != nil {
			return nil
		}
		var out []map[string]any
		if err := json.Unmarshal(data, &out); err != nil {
			return nil
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
	if err != nil {
		return cfg, a.writeFailure(command, output.ExitConfig, "config_error", err)
	}
	return cfg, nil
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
	if writeErr := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); writeErr != nil {
		return output.WithExitCode(code, writeErr)
	}
	return output.WithExitCode(code, err)
}

func (a *app) writeFormSpecFailure(command string, specErr *forms.SpecError) error {
	env := output.Failure(command, output.Error{Code: specErr.Code, Message: specErr.Message})
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
	if err := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); err != nil {
		return output.WithExitCode(code, err)
	}
	if code != output.ExitSuccess {
		return output.WithExitCode(code, fmt.Errorf("%s failed", env.Command))
	}
	return nil
}

func (a *app) outputOptions() output.Options {
	interactive := !a.json && a.stdoutIsInteractive()
	return output.Options{
		JSON:        a.json,
		Interactive: interactive,
		Color:       interactive,
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
