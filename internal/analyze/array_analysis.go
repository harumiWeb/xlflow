package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// ArrayAnalysisResult is the immutable, procedure-local semantic result used
// by array diagnostic projections.  It intentionally has no public fields:
// the result is an implementation detail of one analysis revision and must
// not become a cross-revision cache or a mutable rule API.
//
// The slices and maps are populated before the result is handed to any
// projector.  Projectors only read them and create Finding values, which makes
// the result safe for concurrent read-only consumers.
type ArrayAnalysisResult struct {
	variables      map[string]arrayVariable
	constants      map[string]int
	capacityGuards []arrayResumeNextCapacityGuard
	// cfgWalks counts only fixed-point worklists owned by this result. The
	// ReDim-in-loop inspection is intentionally not included; Range.Value
	// shape is the only explicitly separate secondary lattice.
	cfgWalks       uint64
	projectionRuns uint64

	lifecycleFindings []Finding
	runtimeFindings   []Finding
	redimFindings     []Finding
	rangeFindings     []Finding
}

func (r *ArrayAnalysisResult) findings(kind string) []Finding {
	if r == nil {
		return nil
	}
	var source []Finding
	switch kind {
	case "lifecycle":
		source = r.lifecycleFindings
	case "runtime":
		source = r.runtimeFindings
	case "redim":
		source = r.redimFindings
	case "range":
		source = r.rangeFindings
	}
	return append([]Finding(nil), source...)
}

func arrayAnalysisEnabled(cfg config.AnalyzeConfig) bool {
	if cfg.DetectArrayLifecycleSafety || cfg.DetectRedimPreserveDimension || cfg.DetectObjectArrayComparison || cfg.DetectRedimPreserveInLoops || cfg.DetectRangeValueArrayShape {
		return true
	}
	enabled, known := config.AnalyzeRuleEnabled(cfg, "VBA249")
	return (!known || enabled) && cfg.DetectDeterministicRuntimeErrors
}

// buildArrayAnalysisResult materializes all procedure-local array preparation
// once.  Project-wide ByRef/module/return summaries are supplied by ctx and
// remain owned by the existing analysis-context fixed points.
func (a Analyzer) buildArrayAnalysisResult(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) *ArrayAnalysisResult {
	variables := arrayVariables(file, proc, moduleDecls)
	objectArrayApplicable := arrayObjectArrayApplicable(variables)
	if !arrayAnalysisEnabled(a.Config.Analyze) && !objectArrayApplicable {
		return nil
	}
	result := &ArrayAnalysisResult{
		variables:      variables,
		constants:      arrayIntegerConstants(file, proc, a.visibleConstantValues, a.visibleConstants),
		capacityGuards: arrayResumeNextCapacityGuards(file, proc, variables),
	}
	entryState := arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, variables)

	runtimeEnabled, knownRuntimeRule := config.AnalyzeRuleEnabled(a.Config.Analyze, "VBA249")
	hasArrayVariable := false
	for _, variable := range variables {
		if variable.isArray {
			hasArrayVariable = true
			break
		}
	}
	runtimeRequested := hasArrayVariable && a.Config.Analyze.DetectDeterministicRuntimeErrors && (!knownRuntimeRule || runtimeEnabled)
	comparisonRequested := a.Config.Analyze.DetectObjectArrayComparison && hasArrayVariable && arrayHasComparisonExpression(proc)
	// A procedure can contain an object array even when the explicit lifecycle
	// switches are disabled. The shared transfer owns the always-on VBA101/
	// VBA102 projection for that case.
	coreFlowRequested := a.Config.Analyze.DetectArrayLifecycleSafety ||
		a.Config.Analyze.DetectRedimPreserveDimension ||
		objectArrayApplicable ||
		runtimeRequested
	if coreFlowRequested || comparisonRequested {
		var runtimeSink *[]Finding
		if runtimeRequested {
			runtimeSink = &result.runtimeFindings
		}
		if coreFlowRequested {
			result.lifecycleFindings = a.arrayLifecycleFindingsPreparedWithRuntimeEntry(file, proc, ctx, moduleDecls, result.variables, result.constants, result.capacityGuards, runtimeSink, entryState)
		} else {
			// VBA209 is an expression-only projection. Keep it out of the
			// allocation worklist when no other core lane is applicable; this
			// makes array_cfg_walks reflect actual fixed-point work.
			result.lifecycleFindings = a.arrayComparisonFindings(file, proc, result.variables)
		}
		if proc.Graph != nil && coreFlowRequested {
			result.cfgWalks++
		}
		result.projectionRuns += arrayCoreProjectionRuns(a.Config.Analyze, proc, variables, objectArrayApplicable, runtimeRequested)
	}
	if a.Config.Analyze.DetectRedimPreserveInLoops {
		var redimApplicable bool
		result.redimFindings, redimApplicable = a.redimPreserveLoopFindingsPreparedWithApplicability(file, proc, moduleDecls, result.variables)
		if redimApplicable {
			result.projectionRuns++
		}
	}
	if a.Config.Analyze.DetectRangeValueArrayShape && rangeValueShapeApplicable(file, proc) {
		result.rangeFindings = a.rangeValueShapeFindings(file, proc)
		result.projectionRuns++
		// An empty statement projection uses the source-line fallback and does
		// not start the graph worklist, even when a recovered CFG object exists.
		if proc.Graph != nil && len(proc.Statements) > 0 && !rangeValueProjectionUnknown(proc) {
			result.cfgWalks++
		}
	}
	return result
}

func arrayObjectArrayApplicable(variables map[string]arrayVariable) bool {
	for _, variable := range variables {
		if variable.isArray && variable.isObject {
			return true
		}
	}
	return false
}

// arrayCoreProjectionRuns counts the compatible projections fed by the shared
// core lane. The lane itself is one worklist, but telemetry reports the
// enabled, applicable diagnostic projections rather than treating that lane as
// one diagnostic. This keeps the counter independent from finding multiplicity.
func arrayCoreProjectionRuns(cfg config.AnalyzeConfig, proc sourceProcedure, variables map[string]arrayVariable, objectArrayApplicable, runtimeRequested bool) uint64 {
	var runs uint64
	hasArray := false
	for _, variable := range variables {
		if variable.isArray {
			hasArray = true
			break
		}
	}
	if cfg.DetectArrayLifecycleSafety && arrayLifecycleProjectionApplicable(proc, variables) {
		runs++ // VBA227 lifecycle and For Each projection
	}
	if runtimeRequested {
		runs++ // deterministic array VBA249 projection
	}
	if cfg.DetectRedimPreserveDimension && arrayHasRedimPreserveOperation(proc) {
		runs++ // VBA208
	}
	if cfg.DetectObjectArrayComparison && hasArray && arrayHasComparisonExpression(proc) {
		runs++ // VBA209
	}
	if objectArrayApplicable {
		runs++ // VBA101/VBA102
	}
	return runs
}

func arrayLifecycleProjectionApplicable(proc sourceProcedure, variables map[string]arrayVariable) bool {
	if len(variables) > 0 {
		for _, statement := range proc.Statements {
			lower := strings.ToLower(statement.Text)
			if strings.Contains(lower, "redim") || strings.Contains(lower, "erase ") || strings.Contains(lower, "lbound(") || strings.Contains(lower, "ubound(") || strings.Contains(lower, "for each") {
				return true
			}
			if len(arrayIndexedUses(statement.Text, variables)) > 0 {
				return true
			}
		}
	}
	return false
}

func arrayHasRedimPreserveOperation(proc sourceProcedure) bool {
	for _, statement := range proc.Statements {
		if strings.Contains(strings.ToLower(statement.Text), "redim preserve") {
			return true
		}
	}
	return false
}

func arrayHasComparisonExpression(proc sourceProcedure) bool {
	for _, expression := range proc.Expressions {
		if expression.SyntaxKind == "comparison_expression" && !expression.Recovered {
			return true
		}
	}
	return false
}

// rangeValueShapeApplicable is the cheap procedure-level gate for the
// deliberately independent Range.Value lattice.  The planner can schedule
// the array domain for another rule (for example ReDim), but that must not
// start a Range.Value worklist when the procedure has no Range.Value-shaped
// source. Unknown/recovered statements are left applicable so the existing
// conservative scanner remains fail-open.
func rangeValueShapeApplicable(file parsedFile, proc sourceProcedure) bool {
	if rangeValueProjectionUnknown(proc) {
		// A partial/recovered projection is not proof that the source lacks a
		// Range.Value operation. Keep the domain enabled so the source fallback
		// can make a conservative attempt.
		return true
	}
	// A recovered/empty IR projection is not proof that the procedure has no
	// Range.Value operation. Use the source lines in that case; the fallback is
	// deliberately limited to procedures with no statement projection so the
	// normal IR gate remains unchanged for complete procedures.
	if len(proc.Statements) == 0 {
		for _, expression := range proc.Expressions {
			if expression.Recovered {
				return true
			}
			lower := strings.ToLower(expression.Text)
			if strings.Contains(lower, ".value") || strings.Contains(lower, "range(") || strings.Contains(lower, "resize(") || strings.Contains(lower, "cells(") {
				return true
			}
		}
		return rangeValueSourceLinesApplicable(file, proc)
	}
	for _, statement := range proc.Statements {
		if statement.Recovered {
			return true
		}
		lower := strings.ToLower(statement.Text)
		if strings.Contains(lower, ".value") || strings.Contains(lower, "range(") || strings.Contains(lower, "resize(") || strings.Contains(lower, "cells(") {
			return true
		}
	}
	for _, expression := range proc.Expressions {
		if expression.Recovered {
			return true
		}
		lower := strings.ToLower(expression.Text)
		if strings.Contains(lower, ".value") || strings.Contains(lower, "range(") || strings.Contains(lower, "resize(") || strings.Contains(lower, "cells(") {
			return true
		}
	}
	return false
}

func rangeValueProjectionUnknown(proc sourceProcedure) bool {
	if proc.Features.unknown&featureRangeArray == 0 {
		return false
	}
	if len(proc.Statements) == 0 || len(proc.Expressions) == 0 || proc.Graph == nil {
		return true
	}
	for _, statement := range proc.Statements {
		if statement.Recovered || statement.Kind == procedureir.StatementRecovered || statement.Kind == procedureir.StatementUnknown {
			return true
		}
	}
	return false
}

func (r *ArrayAnalysisResult) lifecycle() []Finding  { return r.findings("lifecycle") }
func (r *ArrayAnalysisResult) runtime() []Finding    { return r.findings("runtime") }
func (r *ArrayAnalysisResult) redim() []Finding      { return r.findings("redim") }
func (r *ArrayAnalysisResult) rangeShape() []Finding { return r.findings("range") }

func arrayEntryStateForProcedure(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable) arrayFlowState {
	state := arrayInitialState(variables)
	state = applyArrayInternalStorageConfiguration(state, file, proc, variables, moduleDecls, ctx.arrayModuleConfigurations[file.Path])
	state = applyArrayByRefEntryStates(state, proc, variables, ctx.arrayByRefEntryStates, ctx.arrayByRefEntryConditions)
	return applyArrayModuleEntryState(state, file, proc, variables, moduleDecls, ctx.arrayModuleEntryStates)
}
