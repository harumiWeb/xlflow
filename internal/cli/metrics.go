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
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
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
			hotspotReport := buildHotspotReport(result)
			hotspotFindings := hotspotFindingsForConfig(hotspotReport, cfg.Metrics.Hotspots)
			env := output.New("metrics")
			env.Metrics = map[string]any{
				"schema_version": proceduremetrics.JSONSchemaVersion,
				"procedures":     result,
				"hotspots":       hotspotReport,
			}
			if len(violations) > 0 {
				env.Diagnostics = violations
				env.Status = output.StatusFailed
				env.Error = &output.Error{
					Code:    "metrics_threshold_exceeded",
					Message: metricsThresholdFailureMessage(len(violations)),
				}
			}
			if len(hotspotFindings) > 0 {
				if existing, ok := env.Diagnostics.([]proceduremetrics.ThresholdViolation); ok {
					merged := make([]any, 0, len(existing)+len(hotspotFindings))
					for _, item := range existing {
						merged = append(merged, item)
					}
					for _, item := range hotspotFindings {
						merged = append(merged, item)
					}
					env.Diagnostics = merged
				} else {
					env.Diagnostics = hotspotFindings
				}
				thresholdFindings := hotspotThresholdFindingCount(hotspotFindings)
				if thresholdFindings > 0 {
					env.Status = output.StatusFailed
					if len(violations) == 0 {
						env.Error = &output.Error{Code: "metrics_hotspot_threshold_exceeded", Message: fmt.Sprintf("%d architectural hotspot threshold finding(s)", thresholdFindings)}
					}
				}
			}
			env.Warnings = warnings
			env.Logs = []string{fmt.Sprintf("calculated metrics for %d procedure(s)", len(result))}
			if len(violations) > 0 || hotspotThresholdFindingCount(hotspotFindings) > 0 {
				return a.write(env, output.ExitValidation)
			}
			return a.write(env, output.ExitSuccess)
		},
	}
}

func buildHotspotReport(procedures []proceduremetrics.ProcedureMetrics) hotspots.Report {
	procedureInputs := make([]hotspots.Input, 0, len(procedures))
	moduleMap := map[string]hotspots.Input{}
	callersByCallee := map[string]map[string]bool{}
	procedureIDs := map[string]string{}
	moduleIDsByProcedure := map[string]string{}
	procedureAdjacency := map[string]map[string]bool{}
	moduleOut := map[string]map[string]bool{}
	moduleIn := map[string]map[string]bool{}
	for _, procedure := range procedures {
		qualified := strings.ToLower(procedure.Module + "." + procedure.Name)
		moduleID := procedure.File + "|" + procedure.Module
		id := procedure.File + "|" + procedure.Module + "|" + procedure.Name + "|" + string(procedure.Kind)
		procedureIDs[qualified] = id
		moduleIDsByProcedure[id] = moduleID
	}
	for _, procedure := range procedures {
		callerID := procedure.File + "|" + procedure.Module + "|" + procedure.Name + "|" + string(procedure.Kind)
		callerModuleID := procedure.File + "|" + procedure.Module
		for _, callee := range procedure.ResolvedCallees {
			if callersByCallee[callee] == nil {
				callersByCallee[callee] = map[string]bool{}
			}
			callersByCallee[callee][callerID] = true
			calleeID, ok := procedureIDs[callee]
			if !ok {
				continue
			}
			if procedureAdjacency[callerID] == nil {
				procedureAdjacency[callerID] = map[string]bool{}
			}
			procedureAdjacency[callerID][calleeID] = true
			calleeModuleID := moduleIDsByProcedure[calleeID]
			if callerModuleID == calleeModuleID {
				continue
			}
			if moduleOut[callerModuleID] == nil {
				moduleOut[callerModuleID] = map[string]bool{}
			}
			if moduleIn[calleeModuleID] == nil {
				moduleIn[calleeModuleID] = map[string]bool{}
			}
			moduleOut[callerModuleID][calleeModuleID] = true
			moduleIn[calleeModuleID][callerModuleID] = true
		}
	}
	cycleParticipation := cycleParticipationCounts(procedureAdjacency)
	for _, procedure := range procedures {
		id := procedure.File + "|" + procedure.Module + "|" + procedure.Name + "|" + string(procedure.Kind)
		moduleID := procedure.File + "|" + procedure.Module
		raw := map[string]int{
			string(hotspots.SignalComplexity):       procedure.CyclomaticComplexity,
			string(hotspots.SignalCallFanOut):       procedure.CallFanOut,
			string(hotspots.SignalCallFanIn):        len(callersByCallee[strings.ToLower(procedure.Module+"."+procedure.Name)]),
			string(hotspots.SignalAffectedModules):  0,
			string(hotspots.SignalExcelEffects):     procedure.ExcelEffectCount,
			string(hotspots.SignalMutableReads):     procedure.MutableStateReads,
			string(hotspots.SignalMutableWrites):    procedure.MutableStateWrites,
			string(hotspots.SignalMutableMutations): procedure.MutableStateMutations,
			string(hotspots.SignalCycleCount):       cycleParticipation[id],
			string(hotspots.SignalExternalDeps):     procedure.ExternalDependencyCount,
			string(hotspots.SignalErrorHandling):    procedure.ErrorHandlingCount,
			string(hotspots.SignalResourceOwned):    procedure.ResourceOwnershipCount,
		}
		uncertainty := map[string]int{
			string(hotspots.UncertaintyAmbiguous):  procedure.AmbiguousCallCount,
			string(hotspots.UncertaintyUnresolved): procedure.UnresolvedCallCount,
			string(hotspots.UncertaintyDynamic):    procedure.DynamicCallCount,
		}
		affected := reachableProcedureModules(id, procedureAdjacency, moduleIDsByProcedure)
		raw[string(hotspots.SignalAffectedModules)] = len(affected)
		procedureInputs = append(procedureInputs, hotspots.Input{ID: id, Kind: "procedure", File: procedure.File, Module: procedure.Module, ModuleKind: procedure.ModuleKind, Name: procedure.Name, ProcedureKind: string(procedure.Kind), DeclarationByte: procedure.DeclarationRange.StartByte, Line: procedure.DeclarationRange.StartLine, RawSignals: raw, Uncertainty: uncertainty})
		module := moduleMap[moduleID]
		if module.ID == "" {
			module = hotspots.Input{ID: moduleID, Kind: "module", File: procedure.File, Module: procedure.Module, ModuleKind: procedure.ModuleKind, Name: procedure.Module, Line: procedure.DeclarationRange.StartLine, RawSignals: map[string]int{}}
			for _, signal := range []hotspots.SignalName{hotspots.SignalComplexity, hotspots.SignalComplexityMax, hotspots.SignalCallFanIn, hotspots.SignalCallFanOut, hotspots.SignalAffectedModules, hotspots.SignalExcelEffects, hotspots.SignalMutableReads, hotspots.SignalMutableWrites, hotspots.SignalMutableMutations, hotspots.SignalCycleCount, hotspots.SignalExternalDeps, hotspots.SignalErrorHandling, hotspots.SignalResourceOwned, hotspots.SignalPublicProcedures} {
				module.RawSignals[string(signal)] = 0
			}
		}
		module.RawSignals[string(hotspots.SignalComplexity)] += procedure.CyclomaticComplexity
		if procedure.CyclomaticComplexity > module.RawSignals[string(hotspots.SignalComplexityMax)] {
			module.RawSignals[string(hotspots.SignalComplexityMax)] = procedure.CyclomaticComplexity
		}
		module.RawSignals[string(hotspots.SignalCallFanOut)] += procedure.CallFanOut
		module.RawSignals[string(hotspots.SignalExcelEffects)] += procedure.ExcelEffectCount
		module.RawSignals[string(hotspots.SignalMutableReads)] += procedure.MutableStateReads
		module.RawSignals[string(hotspots.SignalMutableWrites)] += procedure.MutableStateWrites
		module.RawSignals[string(hotspots.SignalMutableMutations)] += procedure.MutableStateMutations
		module.RawSignals[string(hotspots.SignalCycleCount)] += cycleParticipation[id]
		module.RawSignals[string(hotspots.SignalExternalDeps)] += procedure.ExternalDependencyCount
		module.RawSignals[string(hotspots.SignalErrorHandling)] += procedure.ErrorHandlingCount
		module.RawSignals[string(hotspots.SignalResourceOwned)] += procedure.ResourceOwnershipCount
		if module.Uncertainty == nil {
			module.Uncertainty = map[string]int{}
		}
		for key, value := range uncertainty {
			module.Uncertainty[key] += value
		}
		if !strings.EqualFold(strings.TrimSpace(procedure.Visibility), "private") {
			module.RawSignals[string(hotspots.SignalPublicProcedures)]++
		}
		moduleMap[moduleID] = module
	}
	for moduleID, module := range moduleMap {
		module.RawSignals[string(hotspots.SignalCallFanIn)] = len(moduleIn[moduleID])
		module.RawSignals[string(hotspots.SignalCallFanOut)] = len(moduleOut[moduleID])
		module.RawSignals[string(hotspots.SignalAffectedModules)] = len(reachableModules(moduleID, moduleOut))
		moduleMap[moduleID] = module
	}
	modules := make([]hotspots.Input, 0, len(moduleMap))
	for _, module := range moduleMap {
		modules = append(modules, module)
	}
	return hotspots.BuildReport(procedureInputs, modules)
}

// The traversal starts at each procedure intentionally. Reusing a module-level
// closure would overstate affected modules when a module contains procedures
// with different reachable callees.
func reachableProcedureModules(start string, adjacency map[string]map[string]bool, moduleByProcedure map[string]string) map[string]bool {
	modules := map[string]bool{}
	if module := moduleByProcedure[start]; module != "" {
		modules[module] = true
	}
	seen := map[string]bool{start: true}
	queue := []string{start}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for next := range adjacency[current] {
			if !seen[next] {
				seen[next] = true
				queue = append(queue, next)
			}
			if module := moduleByProcedure[next]; module != "" {
				modules[module] = true
			}
		}
	}
	return modules
}

func reachableModules(start string, adjacency map[string]map[string]bool) map[string]bool {
	seen := map[string]bool{start: true}
	queue := []string{start}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for next := range adjacency[current] {
			if !seen[next] {
				seen[next] = true
				queue = append(queue, next)
			}
		}
	}
	return seen
}

const cycleEnumerationWorkBudget = 100_000

// cycleParticipationCounts enumerates elementary cycles in the confirmed
// procedure graph. Restricting traversal to nodes >= the cycle's start node
// makes the lexicographically smallest node the canonical start, so each
// directed cycle is counted once while keeping the result deterministic. A
// fixed work budget prevents dense graphs from making metrics unbounded; when
// the budget is exhausted, counts are a conservative lower bound.
func cycleParticipationCounts(adjacency map[string]map[string]bool) map[string]int {
	result := map[string]int{}
	nodesSet := map[string]bool{}
	for from, targets := range adjacency {
		nodesSet[from] = true
		for to := range targets {
			nodesSet[to] = true
		}
	}
	nodes := make([]string, 0, len(nodesSet))
	for node := range nodesSet {
		nodes = append(nodes, node)
	}
	sort.Strings(nodes)
	work := 0
	exhausted := false
	for _, start := range nodes {
		if exhausted {
			break
		}
		path := []string{start}
		visited := map[string]bool{start: true}
		var visit func(string)
		visit = func(current string) {
			if exhausted {
				return
			}
			work++
			if work > cycleEnumerationWorkBudget {
				exhausted = true
				return
			}
			nexts := make([]string, 0, len(adjacency[current]))
			for next := range adjacency[current] {
				nexts = append(nexts, next)
			}
			sort.Strings(nexts)
			for _, next := range nexts {
				work++
				if work > cycleEnumerationWorkBudget {
					exhausted = true
					return
				}
				if next == start {
					for _, member := range path {
						result[member]++
					}
					continue
				}
				if next < start || visited[next] {
					continue
				}
				visited[next] = true
				path = append(path, next)
				visit(next)
				path = path[:len(path)-1]
				delete(visited, next)
			}
		}
		visit(start)
	}
	return result
}

func hotspotFindingsForConfig(report hotspots.Report, cfg config.HotspotsConfig) []map[string]any {
	findings := make([]map[string]any, 0)
	appendFindings := func(entities []hotspots.Entity, topN int, threshold float64) {
		for _, entity := range hotspots.Select(entities, topN, threshold) {
			severity := "information"
			if entity.SelectedBy != nil && entity.SelectedBy.Threshold {
				severity = "warning"
			}
			findings = append(findings, map[string]any{"code": "MX002", "severity": severity, "kind": entity.Kind, "file": entity.File, "line": entity.Line, "module": entity.Module, "procedure": entity.Name, "rank": entity.Rank, "score": entity.Score, "score_model": entity.ScoreModel, "raw_signals": entity.RawSignals, "normalized_signals": entity.NormalizedSignals, "uncertainty": entity.Uncertainty, "active_signal_count": entity.ActiveSignalCount, "selected_by": entity.SelectedBy, "message": fmt.Sprintf("%s is an architectural hotspot candidate (score %.2f)", entity.Name, entity.Score)})
		}
	}
	appendFindings(report.Procedures, cfg.ProcedureTopN, cfg.ProcedureScoreThreshold)
	appendFindings(report.Modules, cfg.ModuleTopN, cfg.ModuleScoreThreshold)
	return findings
}

func hotspotThresholdFindingCount(findings []map[string]any) int {
	count := 0
	for _, finding := range findings {
		switch selected := finding["selected_by"].(type) {
		case *hotspots.Selection:
			if selected != nil && selected.Threshold {
				count++
			}
		case hotspots.Selection:
			if selected.Threshold {
				count++
			}
		}
	}
	return count
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
				IsArray: procedure.Symbol.IsArray, ValueShape: procedure.Symbol.ValueShape,
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
		moduleKind, include := metricsTestSourceKind(projectCfg, abs)
		if !include {
			return nil
		}
		seen[key] = true
		files = append(files, symbols.SourceFile{Path: abs, ModuleKind: moduleKind})
		return nil
	})
	if err != nil {
		return nil, err
	}
	sort.SliceStable(files, func(i, j int) bool { return files[i].Path < files[j].Path })
	return files, nil
}

// metricsTestSourceKind applies the same UserForm source-of-truth rule to the
// literal tests/ tree that symbols.DiscoverSourceFiles applies to configured
// source roots. A tests form pair is conventionally laid out as
// tests/<forms-dir>/<Name>.frm and tests/<forms-dir>/code/<Name>.bas. Ordinary
// .bas/.cls files retain their existing standard/class classification.
func metricsTestSourceKind(cfg config.Config, path string) (string, bool) {
	ext := strings.ToLower(filepath.Ext(path))
	switch ext {
	case ".cls":
		return "class", true
	case ".frm":
		sidecar := filepath.Join(filepath.Dir(path), "code", strings.TrimSuffix(filepath.Base(path), ext)+".bas")
		if strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
			if _, err := os.Stat(sidecar); err == nil {
				return "", false
			}
		}
		return "form", true
	case ".bas":
		if strings.EqualFold(filepath.Base(filepath.Dir(path)), "code") {
			formsDir := filepath.Dir(filepath.Dir(path))
			form := filepath.Join(formsDir, strings.TrimSuffix(filepath.Base(path), ext)+".frm")
			if _, err := os.Stat(form); err == nil {
				if strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
					return "form", true
				}
				return "", false
			}
		}
		return "standard", true
	default:
		return "", false
	}
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
