package cli

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/bmatcuk/doublestar/v4"
	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// metricsCommand is intentionally source-only. It builds procedure IR and
// project call resolution directly, without invoking lint/analyze findings.
func (a *app) metricsCommand() *cobra.Command {
	return &cobra.Command{
		Use:   "metrics",
		Short: "Calculate deterministic VBA procedure complexity metrics",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("metrics")
			if err != nil {
				return err
			}
			result, warnings, err := collectProcedureMetrics(cmd.Context(), a.cwd, cfg)
			if err != nil {
				code := output.ExitEnvironment
				if strings.Contains(strings.ToLower(err.Error()), "parse") {
					code = output.ExitValidation
				}
				return a.writeFailure("metrics", code, "metrics_failed", err)
			}
			thresholds := procedureMetricThresholds(cfg.Metrics.Thresholds)
			violations, err := proceduremetrics.EvaluateThresholds(result, thresholds)
			if err != nil {
				return a.writeFailure("metrics", output.ExitConfig, "metrics_thresholds_invalid", err)
			}
			env := output.New("metrics")
			env.Metrics = map[string]any{
				"schema_version": proceduremetrics.JSONSchemaVersion,
				"procedures":     result,
			}
			if len(violations) > 0 {
				env.Diagnostics = violations
				env.Status = output.StatusFailed
				env.Error = &output.Error{
					Code:    "metrics_threshold_exceeded",
					Message: metricsThresholdFailureMessage(len(violations)),
				}
			}
			env.Warnings = warnings
			env.Logs = []string{fmt.Sprintf("calculated metrics for %d procedure(s)", len(result))}
			if len(violations) > 0 {
				return a.write(env, output.ExitValidation)
			}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func metricsThresholdFailureMessage(count int) string {
	if count == 1 {
		return "1 procedure complexity threshold exceeded"
	}
	return fmt.Sprintf("%d procedure complexity thresholds exceeded", count)
}

// collectProcedureMetrics discovers, parses, resolves, and collects all
// configured source procedures. The resolver is assembled from the complete
// project before any call metrics are evaluated, making call fan-out stable
// regardless of filesystem enumeration order.
func collectProcedureMetrics(ctx context.Context, root string, cfg config.Config) ([]proceduremetrics.ProcedureMetrics, []map[string]any, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	files, err := discoverMetricsSourceFiles(root, cfg)
	if err != nil {
		return nil, nil, err
	}

	type sourceDocument struct {
		doc        procedureir.DocumentIR
		relative   string
		moduleKind string
	}
	documents := make([]sourceDocument, 0, len(files))
	warnings := make([]map[string]any, 0)
	matchedExcludes := make([]bool, len(cfg.Metrics.Exclude))
	for _, file := range files {
		if err := ctx.Err(); err != nil {
			return nil, nil, err
		}
		relative, err := filepath.Rel(root, file.Path)
		if err != nil {
			return nil, nil, err
		}
		relative = filepath.ToSlash(filepath.Clean(relative))
		if excluded, matched := metricsExcluded(relative, cfg.Metrics.Exclude); excluded {
			for i := range matched {
				if matched[i] {
					matchedExcludes[i] = true
				}
			}
			continue
		}
		source, err := os.ReadFile(file.Path)
		if err != nil {
			return nil, nil, err
		}
		doc, err := procedureir.BuildSourceContext(ctx, procedureir.BuildOptions{
			RootDir: root, Path: relative, ModuleKind: file.ModuleKind,
		}, source)
		if err != nil {
			return nil, nil, err
		}
		if doc.Parse.HasError || doc.Parse.HasMissing {
			return nil, nil, fmt.Errorf("parse recovery in %s; metrics require complete source", relative)
		}
		if strings.TrimSpace(doc.ModuleName) == "" {
			doc.ModuleName = strings.TrimSuffix(filepath.Base(relative), filepath.Ext(relative))
		}
		doc.Path = relative
		documents = append(documents, sourceDocument{doc: doc, relative: relative, moduleKind: file.ModuleKind})
	}
	for i, pattern := range cfg.Metrics.Exclude {
		if !matchedExcludes[i] {
			warnings = append(warnings, map[string]any{
				"code": "metrics_exclude_unmatched", "severity": "warning",
				"pattern": pattern, "message": fmt.Sprintf("metrics exclude pattern %q matched no source files", pattern),
			})
		}
	}

	resolverSymbols := make([]procedureir.ResolverSymbol, 0)
	for _, item := range documents {
		for _, procedure := range item.doc.Procedures {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: procedure.Symbol.Name, Module: item.doc.ModuleName, ModuleKind: item.moduleKind,
				Kind: string(procedure.Symbol.Kind), Visibility: procedure.Symbol.Visibility,
				File: item.relative, Line: procedure.Symbol.DeclarationRange.StartLine,
			})
		}
	}
	resolver := procedureir.NewResolver(resolverSymbols)
	result := make([]proceduremetrics.ProcedureMetrics, 0)
	for _, item := range documents {
		if err := ctx.Err(); err != nil {
			return nil, nil, err
		}
		resolved := procedureir.Resolve(item.doc, resolver)
		graphs, err := vbacfg.BuildDocumentContext(ctx, resolved)
		if err != nil {
			return nil, nil, err
		}
		result = append(result, proceduremetrics.CollectDocument(resolved, graphs)...)
	}
	proceduremetrics.Sort(result)
	return result, warnings, nil
}

// discoverMetricsSourceFiles extends the shared configured-root discovery with
// the repository's literal tests/ tree. The latter is intentionally not part
// of symbols.DiscoverSourceFiles because symbols inspection only reports
// configured production roots.
func discoverMetricsSourceFiles(root string, projectCfg config.Config) ([]symbols.SourceFile, error) {
	files, err := symbols.DiscoverSourceFiles(symbols.Options{RootDir: root, Config: projectCfg})
	if err != nil {
		return nil, err
	}
	seen := make(map[string]bool, len(files))
	for _, file := range files {
		abs, absErr := filepath.Abs(file.Path)
		if absErr == nil {
			seen[strings.ToLower(filepath.Clean(abs))] = true
		}
	}
	testsRoot := filepath.Join(root, "tests")
	if _, statErr := os.Stat(testsRoot); statErr != nil {
		if os.IsNotExist(statErr) {
			return files, nil
		}
		return nil, statErr
	}
	err = filepath.WalkDir(testsRoot, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			return nil
		}
		ext := strings.ToLower(filepath.Ext(path))
		if ext != ".bas" && ext != ".cls" && ext != ".frm" {
			return nil
		}
		abs, absErr := filepath.Abs(path)
		if absErr != nil {
			return absErr
		}
		key := strings.ToLower(filepath.Clean(abs))
		if seen[key] {
			return nil
		}
		seen[key] = true
		moduleKind := "standard"
		switch ext {
		case ".cls":
			moduleKind = "class"
		case ".frm":
			moduleKind = "form"
		}
		files = append(files, symbols.SourceFile{Path: abs, ModuleKind: moduleKind})
		return nil
	})
	if err != nil {
		return nil, err
	}
	sort.SliceStable(files, func(i, j int) bool { return files[i].Path < files[j].Path })
	return files, nil
}

func metricsExcluded(path string, patterns []string) (bool, []bool) {
	matched := make([]bool, len(patterns))
	for i, pattern := range patterns {
		ok, err := doublestar.Match(pattern, path)
		if err == nil && ok {
			matched[i] = true
		}
	}
	for _, ok := range matched {
		if ok {
			return true, matched
		}
	}
	return false, matched
}

func procedureMetricThresholds(cfg config.Thresholds) proceduremetrics.Thresholds {
	return proceduremetrics.Thresholds{
		proceduremetrics.MetricCyclomaticComplexity: cfg.CyclomaticComplexity,
		proceduremetrics.MetricMaxNestingDepth:      cfg.MaxNestingDepth,
		proceduremetrics.MetricStatementCount:       cfg.StatementCount,
		proceduremetrics.MetricSourceLineCount:      cfg.SourceLineCount,
		proceduremetrics.MetricBranchCount:          cfg.BranchCount,
		proceduremetrics.MetricLoopCount:            cfg.LoopCount,
		proceduremetrics.MetricGotoCount:            cfg.GoToCount,
		proceduremetrics.MetricExitPointCount:       cfg.ExitPointCount,
		proceduremetrics.MetricParameterCount:       cfg.ParameterCount,
		proceduremetrics.MetricByRefParameterCount:  cfg.ByRefParameterCount,
		proceduremetrics.MetricLocalVariableCount:   cfg.LocalVariableCount,
		proceduremetrics.MetricCallFanOut:           cfg.CallFanOut,
	}
}
