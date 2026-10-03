package cli

import (
	"errors"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/vba/architecture"
)

func (a *app) architectureCommand() *cobra.Command {
	var opts architecture.Options
	cmd := &cobra.Command{
		Use:   "architecture",
		Short: "Report VBA project architecture without opening Excel",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("architecture")
			if err != nil {
				return err
			}
			report, warnings, err := architecture.CollectContext(cmd.Context(), a.cwd, cfg, opts)
			if err != nil {
				code, errCode := output.ExitEnvironment, "architecture_failed"
				switch {
				case errors.Is(err, architecture.ErrInvalidScope):
					code, errCode = output.ExitConfig, "architecture_scope_invalid"
				case errors.Is(err, architecture.ErrParse):
					code, errCode = output.ExitValidation, "architecture_parse_failed"
				}
				return a.writeFailure("architecture", code, errCode, err)
			}
			env := output.New("architecture")
			env.Architecture = report
			env.Diagnostics = []any{}
			env.Warnings = warnings
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&opts.Path, "path", "", "display a source file or directory after resolving the whole project")
	cmd.Flags().StringVar(&opts.Module, "module", "", "display one exact module name after resolving the whole project")
	return cmd
}
