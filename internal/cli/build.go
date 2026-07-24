package cli

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"

	buildpkg "github.com/harumiWeb/xlflow/internal/build"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/workbookformat"
)

// buildCommand currently owns the stable planning contract only. The Excel
// mutation pipeline intentionally remains separate so planning cannot modify a
// development workbook or publish a partial release artifact.
func (a *app) buildCommand() *cobra.Command {
	var basePath, outPath string
	var dryRun bool
	cmd := &cobra.Command{
		Use:   "build",
		Short: "Plan an Excel-backed release workbook build",
		Long:  "Plan an Excel-backed release workbook build. The base workbook is never modified; build writes a separate artifact. Excel is required for the future build pipeline. Use pack for the Excel-independent alternative.",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("build")
			if err != nil {
				return err
			}
			plan, err := buildpkg.Plan(buildpkg.Options{Root: a.cwd, Config: cfg, BaseWorkbook: strings.TrimSpace(basePath), OutputPath: strings.TrimSpace(outPath)})
			if err != nil {
				return a.writeFailure("build", output.ExitConfig, "build_plan_invalid", err)
			}
			if err := validateBuildPaths(a.cwd, plan); err != nil {
				return a.writeFailure("build", output.ExitConfig, "build_args_invalid", err)
			}

			payload := buildPayload(plan, dryRun)
			if dryRun {
				env := output.New("build")
				env.Build = payload
				env.Logs = []string{"build plan resolved without opening Excel or writing an artifact"}
				return a.write(env, output.ExitSuccess)
			}
			env := output.Failure("build", output.Error{
				Code:    "build_not_implemented",
				Message: "The Excel-backed build pipeline is not implemented yet. Use --dry-run to inspect the validated build plan.",
				Source:  "xlflow",
			})
			env.Build = payload
			return a.write(env, output.ExitEnvironment)
		},
	}
	cmd.Flags().StringVar(&basePath, "base", "", "base workbook path (defaults to [excel].path)")
	cmd.Flags().StringVar(&outPath, "out", "", "complete output workbook path")
	cmd.Flags().BoolVar(&dryRun, "dry-run", false, "resolve and validate the build plan without opening Excel or writing files")
	return cmd
}

func validateBuildPaths(root string, plan buildpkg.BuildPlan) error {
	base := workbookArgPath(root, plan.BaseWorkbook)
	outputPath := workbookArgPath(root, plan.OutputPath)
	if err := workbookformat.ValidateProjectWorkbookPath(base); err != nil {
		return fmt.Errorf("base workbook: %w", err)
	}
	if err := workbookformat.ValidateProjectWorkbookPath(outputPath); err != nil {
		return fmt.Errorf("output workbook: %w", err)
	}
	if !strings.EqualFold(filepath.Ext(base), filepath.Ext(outputPath)) {
		return fmt.Errorf("output workbook extension must match base workbook format (%s)", filepath.Ext(base))
	}
	info, err := os.Stat(base)
	if err != nil {
		return fmt.Errorf("base workbook does not exist: %w", err)
	}
	if info.IsDir() {
		return fmt.Errorf("base workbook is a directory: %s", plan.BaseWorkbook)
	}
	if info, err := os.Stat(outputPath); err == nil && info.IsDir() {
		return fmt.Errorf("output workbook path is a directory: %s", plan.OutputPath)
	} else if err != nil && !errors.Is(err, os.ErrNotExist) {
		return fmt.Errorf("inspect output workbook path: %w", err)
	}
	if _, err := existingOutputAncestor(outputPath); err != nil {
		return err
	}
	return nil
}

func existingOutputAncestor(path string) (string, error) {
	for current := filepath.Dir(path); ; current = filepath.Dir(current) {
		info, err := os.Stat(current)
		if err == nil {
			if !info.IsDir() {
				return "", fmt.Errorf("output parent is not a directory: %s", current)
			}
			return current, nil
		}
		if !errors.Is(err, os.ErrNotExist) {
			return "", fmt.Errorf("output location is inaccessible: %w", err)
		}
		parent := filepath.Dir(current)
		if parent == current {
			return "", fmt.Errorf("output location has no accessible parent: %s", path)
		}
	}
}

func buildPayload(plan buildpkg.BuildPlan, dryRun bool) map[string]any {
	included := plan.Included
	if included == nil {
		included = []buildpkg.BuildComponent{}
	}
	excluded := plan.Excluded
	if excluded == nil {
		excluded = []buildpkg.BuildComponent{}
	}
	warnings := plan.Warnings
	if warnings == nil {
		warnings = []buildpkg.BuildWarning{}
	}
	return map[string]any{
		"dry_run":  dryRun,
		"base":     plan.BaseWorkbook,
		"output":   plan.OutputPath,
		"included": included,
		"excluded": excluded,
		"warnings": warnings,
	}
}
