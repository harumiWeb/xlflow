package cli

import (
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"

	buildpkg "github.com/harumiWeb/xlflow/internal/build"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/workbookformat"
)

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
			planJSON, err := json.Marshal(plan)
			if err != nil {
				return a.writeFailure("build", output.ExitEnvironment, "build_plan_encode_failed", err)
			}
			base := workbookArgPath(a.cwd, plan.BaseWorkbook)
			out := workbookArgPath(a.cwd, plan.OutputPath)
			env, code, err := a.excelRunnerForConfig(cfg).Build(cfg, base64.StdEncoding.EncodeToString(planJSON), base, out, buildCommandOptions(a.stderrWriter()))
			if err != nil {
				return err
			}
			if env.Status == output.StatusOK {
				env.Build = mergeBuildPayload(payload, env.Build)
			}
			return a.write(env, code)
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
	rootIdentity, err := coordination.NewWorkbookIdentity(root, root)
	if err != nil {
		return fmt.Errorf("resolve project root identity: %w", err)
	}
	outputIdentity, err := coordination.NewWorkbookIdentity(root, outputPath)
	if err != nil {
		return fmt.Errorf("resolve output identity: %w", err)
	}
	relative, err := filepath.Rel(rootIdentity.CanonicalPath, outputIdentity.CanonicalPath)
	if err != nil || relative == ".." || strings.HasPrefix(relative, ".."+string(filepath.Separator)) || filepath.IsAbs(relative) {
		return fmt.Errorf("output workbook resolves outside the project root: %s", plan.OutputPath)
	}
	return nil
}

func mergeBuildPayload(plan map[string]any, bridge any) map[string]any {
	merged := make(map[string]any, len(plan)+4)
	for key, value := range plan {
		merged[key] = value
	}
	if fields, ok := bridge.(map[string]any); ok {
		for key, value := range fields {
			merged[key] = value
		}
		merged["validation"] = map[string]any{
			"source_applied":     fields["source_applied"],
			"components_applied": fields["components_applied"],
			"vbe_compile":        fields["vbe_compile"],
			"workbook_saved":     fields["workbook_saved"],
			"workbook_closed":    fields["workbook_closed"],
			"excel_cleanup":      fields["excel_cleanup"],
		}
	}
	return merged
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
		"schema_version":      1,
		"command":             "build",
		"backend":             "excel",
		"dry_run":             dryRun,
		"base":                plan.BaseWorkbook,
		"output":              plan.OutputPath,
		"included_components": included,
		"excluded_components": excluded,
		"validation": map[string]any{
			"source_applied":  false,
			"vbe_compile":     "not_run",
			"workbook_saved":  false,
			"workbook_closed": false,
		},
		"publication": map[string]any{
			"replaced_existing": false,
			"method":            "not_run",
		},
		"manifest": map[string]any{
			"path":      plan.OutputPath + ".build.json",
			"published": false,
			"error":     nil,
		},
		// Retain the original field names during the v1 transition. They remain
		// convenient for current human renderers and older integrations.
		"included": included,
		"excluded": excluded,
		"warnings": warnings,
	}
}
