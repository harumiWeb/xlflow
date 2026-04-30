package cli

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/agentskill"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/diff"
	"github.com/harumiWeb/xlflow/internal/excel"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/project"
)

type app struct {
	json bool
	cwd  string
}

func Execute() error {
	cwd, err := os.Getwd()
	if err != nil {
		return output.WithExitCode(output.ExitEnvironment, err)
	}
	a := &app{cwd: cwd}
	root := a.rootCommand()
	return root.Execute()
}

func (a *app) rootCommand() *cobra.Command {
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
		a.pullCommand(),
		a.pushCommand(),
		a.traceCommand(),
		a.runCommand(),
		a.macrosCommand(),
		a.testCommand(),
		a.diffCommand(),
		a.lintCommand(),
		a.skillCommand(),
	)
	return root
}

func (a *app) macrosCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "macros",
		Short: "Discover runnable workbook macros",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("macros")
			if err != nil {
				return err
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.Macros(cfg)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
}

func (a *app) newCommand() *cobra.Command {
	var withSkill bool
	var skillAgent string

	cmd := &cobra.Command{
		Use:   "new [workbook]",
		Short: "Create a new xlflow project and macro workbook",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
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
				env, code, err := excel.Runner{RootDir: a.cwd}.New(path)
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
	cmd.Flags().StringVar(&skillAgent, "agent", "", "skill provider target: agents, codex, claude, cursor, gemini, or copilot")
	return cmd
}

func (a *app) initCommand() *cobra.Command {
	var withSkill bool
	var skillAgent string

	cmd := &cobra.Command{
		Use:   "init <workbook>",
		Short: "Create an xlflow project from an existing macro workbook",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
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
	cmd.Flags().StringVar(&skillAgent, "agent", "", "skill provider target: agents, codex, claude, cursor, gemini, or copilot")
	return cmd
}

func (a *app) doctorCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "doctor",
		Short: "Diagnose Excel COM and VBIDE access",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("doctor")
			if err != nil {
				return err
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.Doctor(cfg)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
}

func (a *app) pullCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "pull",
		Short: "Export VBA components from the configured workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("pull")
			if err != nil {
				return err
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.Pull(cfg)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
}

func (a *app) pushCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "push",
		Short: "Import source VBA components into the configured workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("push")
			if err != nil {
				return err
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.Push(cfg)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
}

func buildRunOptions(cfg config.Config, macro, input string, argLiterals []string, save bool, saveAs string, trace bool) (excel.RunOptions, error) {
	if save && saveAs != "" {
		return excel.RunOptions{}, fmt.Errorf("--save and --save-as cannot be combined")
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
	return excel.RunOptions{
		Macro:        macro,
		WorkbookPath: input,
		Args:         args,
		Save:         save,
		SaveAs:       saveAs,
		Trace:        trace,
	}, nil
}

func (a *app) runCommand() *cobra.Command {
	var argLiterals []string
	var input string
	var save bool
	var saveAs string
	var trace bool

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
			opts, err := buildRunOptions(cfg, macro, input, argLiterals, save, saveAs, trace)
			if err != nil {
				return a.writeFailure("run", output.ExitConfig, "run_args_invalid", err)
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.Run(cfg, opts)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringArrayVar(&argLiterals, "arg", nil, "pass a typed macro argument such as string:hello, int:7, or bool:true")
	cmd.Flags().StringVar(&input, "input", "", "override workbook path for this run")
	cmd.Flags().BoolVar(&save, "save", false, "save the opened workbook after a successful run")
	cmd.Flags().StringVar(&saveAs, "save-as", "", "write the successful workbook result to a new path")
	cmd.Flags().BoolVar(&trace, "trace", false, "collect XlflowTrace log events during the run")
	return cmd
}

func (a *app) traceCommand() *cobra.Command {
	trace := &cobra.Command{
		Use:   "trace",
		Short: "Manage workbook trace logging support",
	}
	trace.AddCommand(a.traceInjectCommand())
	return trace
}

func (a *app) traceInjectCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "inject [workbook]",
		Short: "Inject the XlflowTrace VBA module into a workbook",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := config.Load(a.cwd)
			workbook := ""
			if len(args) == 1 {
				workbook = args[0]
			}
			if err != nil {
				if workbook == "" {
					return a.writeFailure("trace", output.ExitConfig, "config_error", err)
				}
				cfg = config.Default()
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.TraceInject(cfg, workbook)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
}

func (a *app) testCommand() *cobra.Command {
	var filter string
	cmd := &cobra.Command{
		Use:   "test",
		Short: "Run workbook VBA tests",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("test")
			if err != nil {
				return err
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.Test(cfg, filter)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringVar(&filter, "filter", "", "run only the test whose procedure name exactly matches filter")
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
	cmd.Flags().StringVar(&agent, "agent", "", "skill provider target: agents, codex, claude, cursor, gemini, or copilot")
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

func (a *app) loadConfig(command string) (config.Config, error) {
	cfg, err := config.Load(a.cwd)
	if err != nil {
		return cfg, a.writeFailure(command, output.ExitConfig, "config_error", err)
	}
	return cfg, nil
}

func (a *app) writeFailure(command string, code int, errCode string, err error) error {
	env := output.Failure(command, output.Error{Code: errCode, Message: err.Error()})
	if writeErr := output.Write(os.Stdout, env, a.json); writeErr != nil {
		return output.WithExitCode(code, writeErr)
	}
	return output.WithExitCode(code, err)
}

func (a *app) write(env output.Envelope, code int) error {
	if err := output.Write(os.Stdout, env, a.json); err != nil {
		return output.WithExitCode(code, err)
	}
	if code != output.ExitSuccess {
		return output.WithExitCode(code, fmt.Errorf("%s failed", env.Command))
	}
	return nil
}
