package analyze

import (
	"context"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// ArrayAnalysisResult is the immutable, procedure-local semantic result used
// by array diagnostic projections. It intentionally has no public fields and
// is safe for the process-local semantic query store to retain while the key's
// revision inputs remain valid.
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
	if source == nil {
		return nil
	}
	findings := make([]Finding, len(source))
	for index, finding := range source {
		findings[index] = cloneFinding(finding)
	}
	return findings
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
func (a Analyzer) buildArrayAnalysisResultContext(cancelCtx context.Context, file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, plan procedureAnalysisPlan) (*ArrayAnalysisResult, error) {
	if err := cancelCtx.Err(); err != nil {
		return nil, err
	}
	variables := arrayVariables(file, proc, moduleDecls)
	objectArrayApplicable := arrayObjectArrayApplicable(variables)
	if !arrayAnalysisEnabled(a.Config.Analyze) && !objectArrayApplicable {
		return nil, nil
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
	runtimeRequested := plan.runsProjection(procedureProjectionRuntime) && hasArrayVariable && a.Config.Analyze.DetectDeterministicRuntimeErrors && (!knownRuntimeRule || runtimeEnabled)
	comparisonRequested := plan.runsProjection(procedureProjectionArrayComparison) && a.Config.Analyze.DetectObjectArrayComparison && hasArrayVariable && arrayHasComparisonExpression(proc)
	// A procedure can contain an object array even when the explicit lifecycle
	// switches are disabled. The shared transfer owns the always-on VBA101/
	// VBA102 projection for that case.
	coreFlowRequested := (plan.runsProjection(procedureProjectionArrayLifecycle) && (a.Config.Analyze.DetectArrayLifecycleSafety || objectArrayApplicable)) ||
		(plan.runsProjection(procedureProjectionArrayRedim) && a.Config.Analyze.DetectRedimPreserveDimension) ||
		runtimeRequested
	if coreFlowRequested || comparisonRequested {
		kernelAnalyzer := a
		kernelAnalyzer.Config.Analyze.DetectArrayLifecycleSafety = plan.runsProjection(procedureProjectionArrayLifecycle) && a.Config.Analyze.DetectArrayLifecycleSafety
		kernelAnalyzer.Config.Analyze.DetectRedimPreserveDimension = plan.runsProjection(procedureProjectionArrayRedim) && a.Config.Analyze.DetectRedimPreserveDimension
		kernelAnalyzer.Config.Analyze.DetectObjectArrayComparison = plan.runsProjection(procedureProjectionArrayComparison) && a.Config.Analyze.DetectObjectArrayComparison
		var runtimeSink *[]Finding
		if runtimeRequested {
			runtimeSink = &result.runtimeFindings
		}
		if coreFlowRequested {
			var err error
			result.lifecycleFindings, err = kernelAnalyzer.arrayLifecycleFindingsPreparedWithRuntimeEntryContext(cancelCtx, file, proc, ctx, moduleDecls, result.variables, result.constants, result.capacityGuards, runtimeSink, entryState)
			if err != nil {
				return nil, err
			}
		} else {
			// VBA209 is an expression-only projection. Keep it out of the
			// allocation worklist when no other core lane is applicable; this
			// makes array_cfg_walks reflect actual fixed-point work.
			result.lifecycleFindings = a.arrayComparisonFindings(file, proc, result.variables)
		}
		if proc.Graph != nil && coreFlowRequested {
			result.cfgWalks++
		}
		result.projectionRuns += arrayCoreProjectionRuns(kernelAnalyzer.Config.Analyze, file, proc, variables, objectArrayApplicable && plan.runsProjection(procedureProjectionArrayLifecycle), runtimeRequested)
	}
	if plan.runsProjection(procedureProjectionArrayRedimLoop) && a.Config.Analyze.DetectRedimPreserveInLoops {
		var redimApplicable bool
		result.redimFindings, redimApplicable = a.redimPreserveLoopFindingsPreparedWithApplicability(file, proc, moduleDecls, result.variables)
		if redimApplicable {
			result.projectionRuns++
		}
	}
	if plan.runsProjection(procedureProjectionArrayRangeShape) && a.Config.Analyze.DetectRangeValueArrayShape && rangeValueShapeApplicable(file, proc) {
		result.rangeFindings = a.rangeValueShapeFindings(file, proc)
		result.projectionRuns++
		// An empty statement projection uses the source-line fallback and does
		// not start the graph worklist, even when a recovered CFG object exists.
		if proc.Graph != nil && proc.Statements.Len() > 0 && !rangeValueProjectionUnknown(proc) {
			result.cfgWalks++
		}
	}
	return result, cancelCtx.Err()
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
func arrayCoreProjectionRuns(cfg config.AnalyzeConfig, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, objectArrayApplicable, runtimeRequested bool) uint64 {
	var runs uint64
	hasArray := false
	for _, variable := range variables {
		if variable.isArray {
			hasArray = true
			break
		}
	}
	if cfg.DetectArrayLifecycleSafety && arrayLifecycleProjectionApplicable(file, proc, variables) {
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

func arrayLifecycleProjectionApplicable(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable) bool {
	if len(variables) == 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		if arrayLifecycleTextApplicable(statement.Text, variables) {
			return true
		}
	}
	if proc.Features.unknown&featureArray != 0 {
		return true
	}
	start, end := proc.StartLine, proc.EndLine
	if start < 1 {
		start = 1
	}
	if end <= 0 || end > len(file.Lines) {
		end = len(file.Lines)
	}
	if start <= end {
		for line := start; line <= end; line++ {
			if arrayLifecycleTextApplicable(normalizedCodeLine(file.Lines[line-1]), variables) {
				return true
			}
		}
	}
	return false
}

func arrayLifecycleTextApplicable(text string, variables map[string]arrayVariable) bool {
	lower := strings.ToLower(text)
	if strings.Contains(lower, "redim") || strings.Contains(lower, "erase ") || strings.Contains(lower, "lbound(") || strings.Contains(lower, "ubound(") || strings.Contains(lower, "for each") {
		return true
	}
	return len(arrayIndexedUses(text, variables)) > 0
}

func arrayHasRedimPreserveOperation(proc sourceProcedure) bool {
	for statement := range proc.Statements.All() {
		if strings.Contains(strings.ToLower(statement.Text), "redim preserve") {
			return true
		}
	}
	return false
}

func arrayHasComparisonExpression(proc sourceProcedure) bool {
	for expression := range proc.Expressions.All() {
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
	if proc.Statements.Len() == 0 {
		for expression := range proc.Expressions.All() {
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
	for statement := range proc.Statements.All() {
		if statement.Recovered {
			return true
		}
		lower := strings.ToLower(statement.Text)
		if strings.Contains(lower, ".value") || strings.Contains(lower, "range(") || strings.Contains(lower, "resize(") || strings.Contains(lower, "cells(") {
			return true
		}
	}
	for expression := range proc.Expressions.All() {
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
	if rangeValueHasUnknownExpression(proc) {
		return true
	}
	if proc.Features.unknown&featureRangeArray == 0 {
		return false
	}
	if proc.Statements.Len() == 0 || proc.Expressions.Len() == 0 {
		return true
	}
	// A graphless procedure can still have a complete statement/expression
	// projection. Keep the existing linear IR transfer for that case; only an
	// actually incomplete projection should fall back to source text.
	for statement := range proc.Statements.All() {
		if statement.Recovered || statement.Kind == procedureir.StatementRecovered || statement.Kind == procedureir.StatementUnknown {
			return true
		}
	}
	return false
}

func rangeValueHasUnknownExpression(proc sourceProcedure) bool {
	rangeStatements := map[int]bool{}
	for statement := range proc.Statements.All() {
		if rangeValueTextLooksRelevant(statement.Text) {
			rangeStatements[statement.ID] = true
		}
	}
	for expression := range proc.Expressions.All() {
		if !expression.Recovered && expression.Kind != procedureir.ExpressionUnknown {
			continue
		}
		if rangeValueTextLooksRelevant(expression.Text) || (expression.StatementID != 0 && rangeStatements[expression.StatementID]) {
			return true
		}
	}
	return false
}

func rangeValueTextLooksRelevant(text string) bool {
	lower := strings.ToLower(text)
	return strings.Contains(lower, ".value") || strings.Contains(lower, "range(") || strings.Contains(lower, "resize(") || strings.Contains(lower, "cells(")
}

func (r *ArrayAnalysisResult) lifecycle() []Finding  { return r.findings("lifecycle") }
func (r *ArrayAnalysisResult) runtime() []Finding    { return r.findings("runtime") }
func (r *ArrayAnalysisResult) redim() []Finding      { return r.findings("redim") }
func (r *ArrayAnalysisResult) rangeShape() []Finding { return r.findings("range") }

func arrayEntryStateForProcedure(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable) arrayFlowState {
	state := arrayInitialState(variables)
	state = applyArrayStaticInitializationState(state, file, proc, ctx, variables)
	state = applyArrayInternalStorageConfiguration(state, file, proc, variables, moduleDecls, ctx.arrayModuleConfigurations[file.Path])
	state = applyArrayByRefEntryStates(state, proc, variables, ctx.arrayByRefEntryStates, ctx.arrayByRefEntryConditions)
	state = applyArrayModuleReadyGuardState(state, file, proc, variables, moduleDecls, ctx.arrayModuleReadyGuards)
	return applyArrayModuleEntryState(state, file, proc, variables, moduleDecls, ctx.arrayModuleEntryStates, ctx.arrayParticipantKeys)
}
