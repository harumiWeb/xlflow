package cli

import (
	"errors"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/config"
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
		a.runCommand(),
		a.testCommand(),
		a.lintCommand(),
	)
	return root
}

func (a *app) newCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "new [workbook]",
		Short: "Create a new xlflow project and macro workbook",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
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
			env := output.New("new")
			env.Workbook = result.Workbook
			env.Logs = []string{
				"created " + result.ConfigPath,
				"created " + result.Workbook,
			}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func (a *app) initCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "init <workbook>",
		Short: "Create an xlflow project from an existing macro workbook",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := project.Init(a.cwd, args[0])
			if err != nil {
				return a.writeFailure("init", output.ExitConfig, "init_failed", err)
			}
			env := output.New("init")
			env.Workbook = result.Workbook
			env.Logs = []string{
				"created " + result.ConfigPath,
				"copied workbook to " + result.Workbook,
			}
			return a.write(env, output.ExitSuccess)
		},
	}
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

func buildRunOptions(cfg config.Config, macro, input string, argLiterals []string, save bool, saveAs string) (excel.RunOptions, error) {
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
	}, nil
}

func (a *app) runCommand() *cobra.Command {
	var argLiterals []string
	var input string
	var save bool
	var saveAs string

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
			opts, err := buildRunOptions(cfg, macro, input, argLiterals, save, saveAs)
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
	return cmd
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
