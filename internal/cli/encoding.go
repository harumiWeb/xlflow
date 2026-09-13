package cli

import (
	"context"
	"errors"
	"fmt"
	"strings"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/vba/sourceencoding"
)

func (a *app) encodingCommand() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "encoding",
		Short: "Check and convert VBA source encodings",
	}
	cmd.AddCommand(a.encodingCheckCommand(), a.encodingConvertCommand())
	return cmd
}

func (a *app) encodingCheckCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "check [path...]",
		Short: "Check VBA source files for UTF-8 without BOM",
		Args:  cobra.ArbitraryArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("encoding check")
			if err != nil {
				return err
			}
			result, checkErr := sourceencoding.Check(cmd.Context(), sourceEncodingOptions(a.cwd, cfg, args))
			env := output.New("encoding check")
			env.Source = result
			env.Logs = sourceEncodingLogs(result)
			if checkErr != nil {
				return a.writeSourceEncodingResult(env, sourceEncodingExitCode(checkErr), "encoding_check_failed", checkErr)
			}
			if result.Summary.Total == 0 {
				env.Logs = []string{"0 source files found"}
			}
			return a.write(env, output.ExitSuccess)
		},
	}
}

// runSourceEncodingPreflight is used by commands that may eventually invoke
// Excel. It deliberately scans the complete managed source scope, including
// files that a particular analyzer mode may otherwise skip (for example a
// sidecar form designer file).
func (a *app) runSourceEncodingPreflight(ctx context.Context, command string, cfg config.Config) error {
	result, err := sourceencoding.Check(ctx, sourceEncodingOptions(a.cwd, cfg, nil))
	if err == nil {
		return nil
	}
	env := output.New(command)
	env.Source = result
	env.Logs = sourceEncodingLogs(result)
	return a.writeSourceEncodingResult(env, sourceEncodingExitCode(err), "encoding_check_failed", err)
}

func (a *app) encodingConvertCommand() *cobra.Command {
	var from string
	cmd := &cobra.Command{
		Use:   "convert --from cp932 [path...]",
		Short: "Convert CP932 VBA source files to UTF-8 without BOM",
		Args:  cobra.ArbitraryArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			if strings.TrimSpace(from) == "" {
				return a.writeFailure("encoding convert", output.ExitConfig, "encoding_args_invalid", errors.New("--from is required"))
			}
			cfg, err := a.loadConfig("encoding convert")
			if err != nil {
				return err
			}
			result, convertErr := sourceencoding.Convert(cmd.Context(), sourceEncodingOptions(a.cwd, cfg, args), from)
			env := output.New("encoding convert")
			env.Source = result
			env.Logs = sourceEncodingLogs(result)
			if convertErr != nil {
				return a.writeSourceEncodingResult(env, sourceEncodingExitCode(convertErr), "encoding_convert_failed", convertErr)
			}
			return a.write(env, output.ExitSuccess)
		},
	}
	cmd.Flags().StringVar(&from, "from", "", "source encoding: cp932")
	return cmd
}

func sourceEncodingOptions(root string, cfg config.Config, paths []string) sourceencoding.Options {
	return sourceencoding.Options{
		RootDir: root,
		Roots: sourceencoding.Roots{
			Modules:  cfg.Src.Modules,
			Classes:  cfg.Src.Classes,
			Forms:    cfg.Src.Forms,
			Workbook: cfg.Src.Workbook,
		},
		Paths: paths,
	}
}

func sourceEncodingLogs(result sourceencoding.Result) []string {
	logs := make([]string, 0, len(result.Files)+1)
	for _, file := range result.Files {
		logs = append(logs, fmt.Sprintf("%s: %s", file.Path, file.Status))
	}
	logs = append(logs, fmt.Sprintf("%d source file(s): %d valid, %d invalid, %d converted, %d unchanged", result.Summary.Total, result.Summary.Valid, result.Summary.Invalid, result.Summary.Converted, result.Summary.Unchanged))
	return logs
}

func sourceEncodingExitCode(err error) int {
	var invalid *sourceencoding.ValidationErrors
	if invalid, _ = errors.AsType[*sourceencoding.ValidationErrors](err); invalid != nil {
		return output.ExitValidation
	}
	var scope *sourceencoding.ScopeError
	if scope, _ = errors.AsType[*sourceencoding.ScopeError](err); scope != nil {
		return output.ExitConfig
	}
	return output.ExitEnvironment
}

func (a *app) writeSourceEncodingResult(env output.Envelope, code int, fallbackCode string, err error) error {
	if projection, ok := sourceEncodingErrorProjection(err); ok {
		env.Status = output.StatusFailed
		env.Error = &output.Error{
			Code:        "source_encoding_invalid",
			Message:     projection.Message,
			Source:      projection.Source,
			Details:     projection.Details,
			Suggestions: projection.Suggestions,
		}
		return a.write(env, output.ExitValidation)
	}
	var transaction *sourceencoding.TransactionError
	if transaction, _ = errors.AsType[*sourceencoding.TransactionError](err); transaction != nil && transaction.Rollback != nil {
		env.Status = output.StatusFailed
		env.Error = &output.Error{
			Code:    "encoding_convert_rollback_failed",
			Message: transaction.Error(),
			Source:  transaction.Path,
			Details: map[string]any{
				"commit_error":   transaction.Cause.Error(),
				"rollback_error": transaction.Rollback.Error(),
			},
		}
		return a.write(env, output.ExitEnvironment)
	}
	var scope *sourceencoding.ScopeError
	if scope, _ = errors.AsType[*sourceencoding.ScopeError](err); scope != nil {
		return a.writeFailure(env.Command, output.ExitConfig, "encoding_args_invalid", err)
	}
	return a.writeFailure(env.Command, code, fallbackCode, err)
}

type sourceEncodingProjection struct {
	Message     string
	Source      string
	Details     any
	Suggestions []string
}

func sourceEncodingErrorProjection(err error) (sourceEncodingProjection, bool) {
	var validation *sourceencoding.ValidationErrors
	if validation, _ = errors.AsType[*sourceencoding.ValidationErrors](err); validation != nil && len(validation.Errors) > 0 {
		first := validation.Errors[0]
		details := first.Details()
		if len(validation.Errors) > 1 {
			files := make([]map[string]any, 0, len(validation.Errors))
			for _, item := range validation.Errors {
				files = append(files, map[string]any{
					"path":        item.Path,
					"status":      string(item.Status),
					"reason":      item.Reason,
					"offset":      item.Position.Offset,
					"line":        item.Position.Line,
					"byte_column": item.Position.ByteColumn,
				})
			}
			details["files"] = files
		}
		return sourceEncodingProjection{
			Message:     validation.Error(),
			Source:      first.Path,
			Details:     details,
			Suggestions: sourceEncodingSuggestions(validation.Errors),
		}, true
	}
	var single *sourceencoding.Error
	if single, _ = errors.AsType[*sourceencoding.Error](err); single != nil {
		return sourceEncodingProjection{
			Message:     single.Error(),
			Source:      single.Path,
			Details:     single.Details(),
			Suggestions: sourceEncodingSuggestions([]*sourceencoding.Error{single}),
		}, true
	}
	return sourceEncodingProjection{}, false
}

func sourceEncodingSuggestions(violations []*sourceencoding.Error) []string {
	suggestions := []string{"encoding check"}
	canConvert := len(violations) > 0
	for _, violation := range violations {
		switch violation.Status {
		case sourceencoding.StatusUTF8BOM, sourceencoding.StatusUTF16BOM:
			return append(suggestions, "save the file as UTF-8 without BOM")
		case sourceencoding.StatusInvalidCP932:
			canConvert = false
		}
	}
	if canConvert {
		return append(suggestions, "encoding convert --from cp932")
	}
	return suggestions
}
