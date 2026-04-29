package cli

import (
	"fmt"
	"os"

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
		a.initCommand(),
		a.doctorCommand(),
		a.pullCommand(),
		a.pushCommand(),
		a.runCommand(),
		a.lintCommand(),
	)
	return root
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

func (a *app) runCommand() *cobra.Command {
	return &cobra.Command{
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
			env, code, err := excel.Runner{RootDir: a.cwd}.Run(cfg, macro)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
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
