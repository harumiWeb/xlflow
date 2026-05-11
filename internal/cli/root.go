package cli

import (
	"cmp"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
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
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/diff"
	"github.com/harumiWeb/xlflow/internal/excel"
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

const defaultKeepaliveInterval = 5 * time.Second

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

type keepaliveFlags struct {
	enabled  bool
	interval time.Duration
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
		a.listCommand(),
		a.pullCommand(),
		a.pushCommand(),
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
		{Name: "workbook-edit-helpers", Description: "Mutate a live session workbook for agent-driven test setup, event triggering, and visual tuning."},
	}
}

func resolvedVersionScripts(root string) []versionScriptInfo {
	commands := []string{"run", "push", "pull", "macros", "test", "trace", "session", "export-image", "edit"}
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
	var keepalive keepaliveFlags
	var session bool
	cmd := &cobra.Command{
		Use:   "macros",
		Short: "Discover runnable workbook macros",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("macros", output.ExitConfig, "macros_args_invalid", err)
			}
			cfg, err := a.loadConfig("macros")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Reading VBA project", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.MacrosWithOptions(cfg, excel.SessionCommandOptions{Session: session, Keepalive: keepaliveOpts})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	addKeepaliveFlags(cmd, &keepalive)
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

func (a *app) listFormsCommand() *cobra.Command {
	var keepalive keepaliveFlags
	var session bool
	cmd := &cobra.Command{
		Use:   "forms",
		Short: "List workbook UserForms",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("list", output.ExitConfig, "list_args_invalid", err)
			}
			cfg, err := a.loadConfig("list")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Listing workbook forms", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.ListForms(cfg, excel.SessionCommandOptions{Session: session, Keepalive: keepaliveOpts})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	addKeepaliveFlags(cmd, &keepalive)
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
	var keepalive keepaliveFlags
	cmd := &cobra.Command{
		Use:   "add",
		Short: "Add or update a workbook form-control button",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			built, err := buildUIButtonAddOptions(opts)
			if err != nil {
				return a.writeFailure("ui", output.ExitConfig, "ui_button_args_invalid", err)
			}
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("ui", output.ExitConfig, "ui_button_args_invalid", err)
			}
			cfg, err := a.loadConfig("ui")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Adding workbook button", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.UIButtonAdd(cfg, built, keepaliveOpts)
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
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) uiButtonListCommand() *cobra.Command {
	var opts excel.UIButtonListOptions
	var keepalive keepaliveFlags
	cmd := &cobra.Command{
		Use:   "list",
		Short: "List xlflow-managed workbook buttons",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("ui", output.ExitConfig, "ui_button_args_invalid", err)
			}
			cfg, err := a.loadConfig("ui")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Listing workbook buttons", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.UIButtonList(cfg, opts, keepaliveOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&opts.Sheet, "sheet", "", "worksheet name")
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) uiButtonRemoveCommand() *cobra.Command {
	var opts excel.UIButtonRemoveOptions
	var keepalive keepaliveFlags
	cmd := &cobra.Command{
		Use:   "remove",
		Short: "Remove an xlflow-managed workbook button",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			built, err := buildUIButtonRemoveOptions(opts)
			if err != nil {
				return a.writeFailure("ui", output.ExitConfig, "ui_button_args_invalid", err)
			}
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("ui", output.ExitConfig, "ui_button_args_invalid", err)
			}
			cfg, err := a.loadConfig("ui")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Removing workbook button", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.UIButtonRemove(cfg, built, keepaliveOpts)
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
	addKeepaliveFlags(cmd, &keepalive)
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
	var keepalive keepaliveFlags

	cmd := &cobra.Command{
		Use:   "new [workbook]",
		Short: "Create a new xlflow project and macro workbook",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			if err := a.writeScaffoldWelcome("new", noUpdateCheck); err != nil {
				return output.WithExitCode(output.ExitEnvironment, err)
			}
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("new", output.ExitConfig, "new_args_invalid", err)
			}
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
			var excelEnv output.Envelope
			var excelCode int
			result, err := project.New(a.cwd, workbook, func(path string) error {
				env, code, err := a.runExcelWithProgress("Creating workbook", keepaliveOpts, func() (output.Envelope, int, error) {
					return excel.Runner{RootDir: a.cwd}.New(path, keepaliveOpts)
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
			}
			if withSkill {
				env.Logs = append(env.Logs, "installed xlflow skill to "+skillResult.Path)
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&withSkill, "with-skill", false, "install the bundled xlflow AI agent skill")
	cmd.Flags().StringVar(&skillAgent, "agent", "", "skill provider target: agents, codex, claude, cursor, or gemini")
	cmd.Flags().BoolVar(&noUpdateCheck, "no-update-check", false, "skip the interactive GitHub release update check during project scaffolding")
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) initCommand() *cobra.Command {
	var withSkill bool
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
			var skillOpts agentskill.InstallOptions
			if withSkill {
				opts, err := a.resolveSkillInstallOptions(skillAgent, "", false)
				if err != nil {
					return a.writeFailure("init", output.ExitConfig, "skill_agent_required", err)
				}
				skillOpts = opts
			}
			result, err := project.Init(a.cwd, args[0])
			if err != nil {
				return a.writeFailure("init", output.ExitConfig, "init_failed", err)
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
			}
			if withSkill {
				env.Logs = append(env.Logs, "installed xlflow skill to "+skillResult.Path)
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().BoolVar(&withSkill, "with-skill", false, "install the bundled xlflow AI agent skill")
	cmd.Flags().StringVar(&skillAgent, "agent", "", "skill provider target: agents, codex, claude, cursor, or gemini")
	cmd.Flags().BoolVar(&noUpdateCheck, "no-update-check", false, "skip the interactive GitHub release update check during project scaffolding")
	return cmd
}

func (a *app) doctorCommand() *cobra.Command {
	var keepalive keepaliveFlags
	cmd := &cobra.Command{
		Use:   "doctor",
		Short: "Diagnose Excel COM and VBIDE access",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("doctor", output.ExitConfig, "doctor_args_invalid", err)
			}
			cfg, err := a.loadConfig("doctor")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Checking Excel automation", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.Doctor(cfg, keepaliveOpts)
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
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) attachCommand() *cobra.Command {
	var active bool
	var keepalive keepaliveFlags
	cmd := &cobra.Command{
		Use:   "attach --active",
		Short: "Inspect the active Excel workbook connection",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("attach", output.ExitConfig, "attach_args_invalid", err)
			}
			cfg, err := a.loadConfig("attach")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Inspecting active workbook", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.Attach(cfg, active, keepaliveOpts)
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&active, "active", false, "attach to the active Excel workbook")
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) pullCommand() *cobra.Command {
	var keepalive keepaliveFlags
	var session bool
	cmd := &cobra.Command{
		Use:   "pull",
		Short: "Export VBA components from the configured workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("pull", output.ExitConfig, "pull_args_invalid", err)
			}
			cfg, err := a.loadConfig("pull")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Exporting VBA source", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.PullWithOptions(cfg, excel.SessionCommandOptions{Session: session, Keepalive: keepaliveOpts})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) pushCommand() *cobra.Command {
	var keepalive keepaliveFlags
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
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("push", output.ExitConfig, "push_args_invalid", err)
			}
			cfg, err := a.loadConfig("push")
			if err != nil {
				return err
			}
			pushOpts, err := buildPushOptions(backupMode, fast, changedOnly, session, noSave, keepaliveOpts)
			if err != nil {
				return a.writeFailure("push", output.ExitConfig, "push_args_invalid", err)
			}
			if err := a.runSourcePreflight("push", cfg, "pushing to Excel", nil); err != nil {
				return err
			}
			var env output.Envelope
			var code int
			run := func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.PushWithOptions(cfg, pushOpts)
				return runErr
			}
			err = a.withExcelProgress("Importing VBA source", keepaliveOpts, run)
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
	addKeepaliveFlags(cmd, &keepalive)
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

func addKeepaliveFlags(cmd *cobra.Command, keepalive *keepaliveFlags) {
	cmd.Flags().BoolVar(&keepalive.enabled, "keepalive", false, "write periodic progress heartbeat lines to stderr")
	cmd.Flags().DurationVar(&keepalive.interval, "keepalive-interval", defaultKeepaliveInterval, "interval between keepalive heartbeat lines")
}

func buildKeepaliveOptions(keepalive bool, interval time.Duration) (excel.CommandOptions, error) {
	if keepalive && interval <= 0 {
		return excel.CommandOptions{}, fmt.Errorf("--keepalive-interval must be greater than 0")
	}
	return excel.CommandOptions{
		Keepalive:         keepalive,
		KeepaliveInterval: interval,
		Stderr:            os.Stderr,
	}, nil
}

func buildRunOptions(cfg config.Config, macro, input string, argLiterals []string, save bool, saveAs string, trace bool, headless bool, interactive bool, direct bool, fast bool, diagnostic bool, diagnosticExplicit bool, guiCompileErrors bool, session bool, timeout time.Duration, keepalive bool, keepaliveInterval time.Duration) (excel.RunOptions, error) {
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
	if direct && diagnostic {
		if diagnosticExplicit {
			return excel.RunOptions{}, fmt.Errorf("--direct cannot be combined with diagnostic mode; omit --diagnostic or use --gui-compile-errors")
		}
		diagnostic = false
	}
	keepaliveOpts, err := buildKeepaliveOptions(keepalive, keepaliveInterval)
	if err != nil {
		return excel.RunOptions{}, err
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
	mode := ""
	if headless {
		mode = "headless"
	}
	if interactive {
		mode = "interactive"
	}
	return excel.RunOptions{
		Macro:               macro,
		WorkbookPath:        input,
		Args:                args,
		Save:                save,
		SaveAs:              saveAs,
		Trace:               trace,
		Mode:                mode,
		Direct:              direct,
		Fast:                fast,
		Diagnostic:          diagnostic,
		SuppressModalErrors: !guiCompileErrors,
		Session:             session,
		Timeout:             timeout,
		Keepalive:           keepaliveOpts,
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

func (a *app) runCommand() *cobra.Command {
	var argLiterals []string
	var input string
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
	var keepalive bool
	var keepaliveInterval time.Duration

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
			opts, err := buildRunOptions(cfg, macro, input, argLiterals, save, saveAs, trace, headless, interactive, direct, fast, diagnostic, cmd.Flags().Changed("diagnostic"), guiCompileErrors, session, timeout, keepalive, keepaliveInterval)
			if err != nil {
				return a.writeFailure("run", output.ExitConfig, "run_args_invalid", err)
			}
			if a.shouldRunSourcePreflight(cfg, opts) {
				if err := a.runSourcePreflight("run", cfg, "running macros", ignoredRunPreflightAnalysisCodes(opts)); err != nil {
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
					env.Logs = []string{
						"Use xlflow run --interactive if a human can operate Excel dialogs.",
						"For repeatable automation, refactor GUI entrypoints into parameterized headless procedures.",
					}
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
			if opts.Keepalive.Keepalive {
				err = run()
			} else {
				err = a.withSpinner("Running macro", run)
			}
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
	cmd.Flags().BoolVar(&keepalive, "keepalive", false, "write periodic progress heartbeat lines to stderr")
	cmd.Flags().DurationVar(&keepaliveInterval, "keepalive-interval", defaultKeepaliveInterval, "interval between keepalive heartbeat lines")
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
	var keepalive keepaliveFlags

	cmd := &cobra.Command{
		Use:   "export-image [workbook]",
		Short: "Export a worksheet range as an image",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("export-image", output.ExitConfig, "export_image_args_invalid", err)
			}
			cfg, err := a.loadConfig("export-image")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildExportImageOptions(workbook, sheet, cellRange, outPath, outputDir, name, format, overwrite, session, keepaliveOpts)
			if err != nil {
				return a.writeFailure("export-image", output.ExitConfig, "export_image_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Exporting worksheet range image", keepaliveOpts, func() error {
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
	addKeepaliveFlags(cmd, &keepalive)
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
	var keepalive keepaliveFlags

	cmd := &cobra.Command{
		Use:   "cell [workbook]",
		Short: "Edit one live-session cell",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
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
			opts, err := buildEditCellOptions(workbook, sheet, cell, fill, events, valuePtr, formulaPtr, session, keepaliveOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook cell", keepaliveOpts, func() error {
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
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) editRangeCommand() *cobra.Command {
	var sheet string
	var cellRange string
	var fill string
	var clear string
	var session bool
	var keepalive keepaliveFlags

	cmd := &cobra.Command{
		Use:   "range [workbook]",
		Short: "Edit one live-session range",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildEditRangeOptions(workbook, sheet, cellRange, fill, clear, session, keepaliveOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook range", keepaliveOpts, func() error {
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
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) editRowsCommand() *cobra.Command {
	var sheet string
	var rows string
	var height float64
	var session bool
	var keepalive keepaliveFlags

	cmd := &cobra.Command{
		Use:   "rows [workbook]",
		Short: "Set row height on a live-session worksheet",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildEditRowsOptions(workbook, sheet, rows, height, session, keepaliveOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook row heights", keepaliveOpts, func() error {
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
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) editColumnsCommand() *cobra.Command {
	var sheet string
	var columns string
	var width float64
	var session bool
	var keepalive keepaliveFlags

	cmd := &cobra.Command{
		Use:   "columns [workbook]",
		Short: "Set column width on a live-session worksheet",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			cfg, err := a.loadConfig("edit")
			if err != nil {
				return err
			}
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			opts, err := buildEditColumnsOptions(workbook, sheet, columns, width, session, keepaliveOpts)
			if err != nil {
				return a.writeFailure("edit", output.ExitConfig, "edit_args_invalid", err)
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Editing workbook column widths", keepaliveOpts, func() error {
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
	addKeepaliveFlags(cmd, &keepalive)
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
	var keepalive keepaliveFlags
	var force bool
	var session bool
	cmd := &cobra.Command{
		Use:   action + " [workbook]",
		Short: short,
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("trace", output.ExitConfig, "trace_args_invalid", err)
			}
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
			err = a.withExcelProgress(label, keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.Trace(cfg, excel.TraceOptions{Action: traceAction, Workbook: workbook, Force: force, Session: session}, keepaliveOpts)
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
	addKeepaliveFlags(cmd, &keepalive)
	return cmd
}

func (a *app) testCommand() *cobra.Command {
	var filter string
	var keepalive keepaliveFlags
	var session bool
	cmd := &cobra.Command{
		Use:   "test",
		Short: "Run workbook VBA tests",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("test", output.ExitConfig, "test_args_invalid", err)
			}
			cfg, err := a.loadConfig("test")
			if err != nil {
				return err
			}
			var env output.Envelope
			var code int
			err = a.withExcelProgress("Running VBA tests", keepaliveOpts, func() error {
				var runErr error
				env, code, runErr = excel.Runner{RootDir: a.cwd}.TestWithOptions(cfg, filter, excel.SessionCommandOptions{Session: session, Keepalive: keepaliveOpts})
				return runErr
			})
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&filter, "filter", "", "run only the test whose procedure name exactly matches filter")
	cmd.Flags().BoolVar(&session, "session", false, "force "+sessionUsageHint())
	addKeepaliveFlags(cmd, &keepalive)
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
		a.inspectRangeCommand(flags),
		a.inspectUsedRangeCommand(flags),
		a.inspectCellCommand(flags),
	)
	return cmd
}

func (a *app) inspectWorkbookCommand(flags *inspectSharedFlags) *cobra.Command {
	return &cobra.Command{
		Use:   "workbook",
		Short: "Inspect workbook summary information",
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
			workbook, err := workbookinspect.Workbook(workbookArgPath(a.cwd, cfg.Excel.Path))
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			target, session, warnings := a.inspectStateForWorkbook(cfg, workbook.Path)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env := output.New("inspect")
			env.Target = target
			env.Session = session
			env.Warnings = append(warnings, formWarnings...)
			env.Hints = formHints
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
}

func (a *app) inspectSheetsCommand(flags *inspectSharedFlags) *cobra.Command {
	return &cobra.Command{
		Use:   "sheets",
		Short: "Inspect workbook worksheets",
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
			sheets, err := workbookinspect.Sheets(workbookArgPath(a.cwd, cfg.Excel.Path))
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			target, session, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env := output.New("inspect")
			env.Target = target
			env.Session = session
			env.Warnings = append(warnings, formWarnings...)
			env.Hints = formHints
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
}

func (a *app) inspectRangeCommand(flags *inspectSharedFlags) *cobra.Command {
	var sheet string
	var address string
	var maxRows int
	var maxCols int
	var includeStyle bool
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
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
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
			target, session, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env.Target = target
			env.Session = session
			env.Warnings = append(warnings, formWarnings...)
			env.Hints = formHints
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
	return cmd
}

func (a *app) inspectUsedRangeCommand(flags *inspectSharedFlags) *cobra.Command {
	var sheet string
	var maxRows int
	var maxCols int
	var includeStyle bool
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
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
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
			target, session, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env.Target = target
			env.Session = session
			env.Warnings = append(warnings, formWarnings...)
			env.Hints = formHints
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
	return cmd
}

func (a *app) inspectCellCommand(flags *inspectSharedFlags) *cobra.Command {
	var sheet string
	var address string
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
			cfg, err := a.loadConfig("inspect")
			if err != nil {
				return err
			}
			cell, err := workbookinspect.Cell(workbookArgPath(a.cwd, cfg.Excel.Path), selector.Sheet, selector.Address)
			if err != nil {
				return a.writeFailure("inspect", output.ExitEnvironment, "inspect_failed", err)
			}
			workbookPath := workbookArgPath(a.cwd, cfg.Excel.Path)
			target, session, warnings := a.inspectStateForWorkbook(cfg, workbookPath)
			formWarnings, formHints := inspectSourceUserFormMessages(a.cwd, cfg)
			env := output.New("inspect")
			env.Target = target
			env.Session = session
			env.Warnings = append(warnings, formWarnings...)
			env.Hints = formHints
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
	return cmd
}

func (a *app) inspectStateForWorkbook(cfg config.Config, workbookPath string) (map[string]any, map[string]any, []map[string]any) {
	target := map[string]any{
		"kind":        "file",
		"path":        workbookPath,
		"description": "Saved workbook file on disk",
	}
	session := map[string]any{
		"active":        false,
		"workbook_path": workbookPath,
		"dirty":         false,
		"save_required": false,
	}
	if runtime.GOOS != "windows" {
		return target, session, nil
	}
	metadataPath := filepath.Join(a.cwd, ".xlflow", "session.json")
	if !inspectSessionMetadataMatchesWorkbook(metadataPath, workbookPath) {
		return target, session, nil
	}
	env, _, err := excel.Runner{RootDir: a.cwd}.Session(cfg, "status")
	if err != nil {
		return target, session, nil
	}
	status := cliObjectMap(env.Session)
	active := boolValueForCLI(status, "active") || (boolValueForCLI(status, "running") && boolValueForCLI(status, "workbook_open"))
	dirty := boolValueForCLI(status, "dirty")
	saveRequired := boolValueForCLI(status, "save_required") || boolValueForCLI(status, "needs_save")
	session["active"] = active
	session["dirty"] = dirty
	session["save_required"] = saveRequired
	if mode := stringValueForCLI(status, "mode"); mode != "" {
		session["mode"] = mode
	}
	if !active || !dirty {
		return target, session, nil
	}
	return target, session, []map[string]any{
		{
			"code":    "live_session_dirty",
			"message": "A live session exists and has unsaved changes. This command inspected the saved file, not the live workbook.",
		},
		{
			"code":    "command_reads_saved_file",
			"message": "This inspect command read the saved workbook file on disk.",
		},
	}
}

func inspectSessionMetadataMatchesWorkbook(metadataPath, workbookPath string) bool {
	data, err := os.ReadFile(metadataPath)
	if err != nil {
		return false
	}
	var metadata struct {
		WorkbookPath string `json:"workbook_path"`
	}
	if err := json.Unmarshal(data, &metadata); err != nil {
		return false
	}
	if strings.TrimSpace(metadata.WorkbookPath) == "" {
		return false
	}
	return strings.EqualFold(filepath.Clean(metadata.WorkbookPath), filepath.Clean(workbookPath))
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
		"message": "Planned/future commands for deeper UserForm inspection include `xlflow form snapshot <name> --designer`, `xlflow inspect form <name> --runtime --json`, and `xlflow export-form-image <name>`.",
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
	var keepalive keepaliveFlags
	cmd := &cobra.Command{
		Use:   "check",
		Short: "Run lint, analyze, and doctor",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			keepaliveOpts, err := buildKeepaliveOptions(keepalive.enabled, keepalive.interval)
			if err != nil {
				return a.writeFailure("check", output.ExitConfig, "check_args_invalid", err)
			}
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
			err = a.withExcelProgress("Checking Excel automation", keepaliveOpts, func() error {
				var runErr error
				doctor, doctorCode, runErr = excel.Runner{RootDir: a.cwd}.Doctor(cfg, keepaliveOpts)
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
	addKeepaliveFlags(cmd, &keepalive)
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

func (a *app) runSourcePreflight(command string, cfg config.Config, action string, ignoredAnalysisCodes map[string]bool) error {
	issues, err := lint.Linter{RootDir: a.cwd, Config: cfg}.Run()
	if err != nil {
		return a.writeFailure(command, output.ExitEnvironment, "lint_failed", err)
	}
	findings, err := analyze.Analyzer{RootDir: a.cwd, Config: cfg}.Run()
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
	if opts.Keepalive {
		return fn()
	}
	return a.withSpinner(label, fn)
}

func (a *app) withSpinner(label string, fn func() error) error {
	if a.json || !a.stdoutIsInteractive() || !a.stderrIsInteractive() {
		return fn()
	}
	return runSpinner(a.stderrWriter(), label, fn)
}
