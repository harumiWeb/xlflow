package analyze

import (
	"context"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
	"github.com/harumiWeb/xlflow/internal/gui"
	"github.com/harumiWeb/xlflow/internal/lint"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type arrayCFGStrategy uint8

const (
	arrayCFGStrategyAuto arrayCFGStrategy = iota
	arrayCFGStrategyCompact
	arrayCFGStrategyLegacy
)

type arrayFallbackReason uint8

const (
	arrayFallbackUnsupported arrayFallbackReason = iota
	arrayFallbackEmptyState
	arrayFallbackIndex
)

// arrayAllocation is deliberately a three-point lattice.  In particular,
// unknown is not treated as allocated: a VBA runtime operation must be
// proven safe on every path before it is allowed through the analysis.
type arrayAllocation uint8

const (
	arrayUnknown arrayAllocation = iota
	arrayUnallocated
	arrayAllocated
)

type arrayBound struct {
	known      bool
	value      int
	expression string
}

type arrayDimension struct {
	lower arrayBound
	upper arrayBound
}

type arrayOrigin uint8

const (
	arrayOriginUnknown arrayOrigin = iota
	arrayOriginLocal
	arrayOriginRangeValue
)

type arrayVariable struct {
	name      string
	typ       string
	isArray   bool
	isVariant bool
	isObject  bool
	// knownScalar is intentionally narrower than "not an array".  Only
	// built-in scalar declarations are strong enough to prove that an array
	// operation or For Each source is invalid.  User-defined classes/UDTs and
	// unresolved external types remain unknown and therefore fail open.
	knownScalar bool
	fixed       bool
	parameter   bool
	static      bool
	paramArray  bool
	dimensions  []arrayDimension
}

type arrayValue struct {
	kind            arrayAllocation
	knownArray      bool
	mayBeEmpty      bool
	dimensions      []arrayDimension
	preserveShape   []arrayDimension
	origin          arrayOrigin
	allocationProbe string
	safeBoundProbe  string
	// allocationCountSource records a narrow conditional allocation contract:
	// the array is allocated when the named scalar is positive, or when the
	// named collection's Count is positive. The fact is refined only on a
	// matching control-flow branch; it is never treated as unconditional.
	allocationCountSource string
	// conditionalAllocationSource records that a direct ReDim occurred under
	// the true body of a simple scalar comparison. The fact survives the merge
	// with the branch that skipped the ReDim and is consumed only by a later
	// matching true branch.
	conditionalAllocationSource string
	// returnNonEmptyArrayParameter is set only on an interprocedural return
	// summary. It names the callee's ByRef array parameter whose non-empty value
	// makes the returned array allocation definite.
	returnNonEmptyArrayParameter string
	// returnPositiveScalarParameter is set only on an interprocedural return
	// summary. It names the callee's scalar parameter whose positive value makes
	// the returned array allocation definite.
	returnPositiveScalarParameter string
	// nonEmptySource records the caller-side array whose non-empty state makes
	// this returned array non-empty. It is consumed by a matching StrPtr guard.
	nonEmptySource string
	// returnDescriptor* records a narrow typed-array return whose SAFEARRAY
	// descriptor is populated from scalar function parameters. The metadata is
	// consumed at the caller so known arguments can recover the returned shape;
	// unknown arguments retain only the allocation proof.
	returnDescriptorSourceParameter string
	returnDescriptorStartParameter  string
	returnDescriptorLengthParameter string
	returnDescriptorLowerParameter  string
	boundsProof                     arrayBoundsProof
}

type arrayFlowState map[string]arrayValue

type arrayBoundsProof struct {
	loopEndLine                      int
	priorKind                        arrayAllocation
	priorKnownArray                  bool
	priorMayBeEmpty                  bool
	priorAllocationCount             string
	priorConditionalAllocationSource string
}

const arrayDictionaryCountSourcePrefix = "dictionary:"

// arrayResumeNextCapacityGuard describes the narrow growable-buffer idiom
// where an unallocated array is probed with UBound under Resume Next, a
// capacity fallback ReDim Preserve allocates it when data is present, and a
// following loop writes only the requested range. The guard is used only for
// the probe and that loop's indexed writes; it does not globally mark the
// array allocated.
type arrayResumeNextCapacityGuard struct {
	target         string
	probeLine      int
	indexStartLine int
	indexEndLine   int
}

type arrayUse struct {
	name string
	args []string
}

type directArrayRedimClause struct {
	name       string
	dimensions string
}

var (
	arrayRedimRe                      = regexp.MustCompile(`(?i)^\s*redim\s+(preserve\s+)?(.+)$`)
	arrayRedimClauseRe                = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?)?\s*$`)
	arrayRedimTypeSuffixRe            = regexp.MustCompile(`(?i)^as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?$`)
	arrayEraseRe                      = regexp.MustCompile(`(?i)^\s*erase\s+(.+)$`)
	arrayEraseNameRe                  = regexp.MustCompile(`(?i)^[A-Za-z_]\w*$`)
	arrayBoundCallRe                  = regexp.MustCompile(`(?i)\b(lbound|ubound)\s*\(\s*([^,)]*)\s*(?:,\s*([^)]*))?\)`)
	arrayBoundOperatorRe              = regexp.MustCompile(`(?i)\b(?:mod|and|or|not)\b`)
	arrayForBoundRe                   = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*([-+]?\d+)\s+to\s+(?:lbound|ubound)\s*\(\s*([A-Za-z_]\w*)`)
	arrayForUBoundRe                  = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*([-+]?\d+)\s+to\s+ubound\s*\(\s*([A-Za-z_]\w*)`)
	arrayReturnArrayDocRe             = regexp.MustCompile(`(?i)^@returns?\s+(?:(?:variant|object)\s*<)?array(?:<|\b)`)
	arrayTypeNameExpressionRe         = regexp.MustCompile(`(?i)^typename\s*\(\s*([A-Za-z_]\w*)\s*\)$`)
	arrayQuotedCaseRe                 = regexp.MustCompile(`^\s*"([^"]*)"\s*$`)
	arrayEmptyGuardRe                 = regexp.MustCompile(`(?i)^\s*\(?\s*not\s+([A-Za-z_]\w*)\s*\)?\s*=\s*-1\s*$`)
	arrayForScalarBoundRe             = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*[-+]?\d+\s+to\s+([A-Za-z_]\w*)\s*$`)
	arrayForCountRe                   = regexp.MustCompile(`(?i)^\s*for\s+([A-Za-z_]\w*)\s*=\s*(0|1)\s+to\s+([A-Za-z_]\w*)\s*\.\s*count(\s*-\s*1)?\s*$`)
	arrayDimensionCountLoopRe         = regexp.MustCompile(`(?i)^\s*for\s+([A-Za-z_]\w*)\s*=\s*1\s+to\s+[A-Za-z_]\w*\s*$`)
	arrayForEachRe                    = regexp.MustCompile(`(?i)^\s*for\s+each\s+[A-Za-z_]\w*\s+in\s+([^\r\n]+)`)
	arrayIndexedSourceRe              = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\(`)
	arrayGuardCallRe                  = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*(?:(=|<>|>=|<=|>|<)\s*(-?\d+))?\s*$`)
	arrayGuardReversedRe              = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>|>=|<=|>|<)\s*([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*$`)
	arrayGuardValueRe                 = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*(=|<>|>=|<=|>|<)\s*(-?\d+)\s*$`)
	arrayGuardValueReversedRe         = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>|>=|<=|>|<)\s*([A-Za-z_]\w*)\s*$`)
	arrayIsArrayGuardRe               = regexp.MustCompile(`(?i)^\s*isarray\s*\(\s*(.+)\s*\)\s*$`)
	arrayByteArrayGuardRe             = regexp.MustCompile(`(?i)^\s*(?:vartypeof|vartype)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*=\s*\(?\s*vbarray\s+or\s+vbbyte\s*\)?\s*$`)
	arrayStrPtrGuardRe                = regexp.MustCompile(`(?i)^\s*strptr\s*\(\s*([A-Za-z_]\w*)\s*\)\s*(=|<>)\s*0\s*$`)
	arraySafeArrayZeroExitGuardRe     = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arraySafeArrayPointerCopyRe       = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(?:[A-Za-z_]\w*\.)?copymemoryfromptr\s+([A-Za-z_]\w*)\s*,\s*([A-Za-z_]\w*)\s*,\s*lenb\s*\(\s*([A-Za-z_]\w*)\s*\)\s*$`)
	arrayByteArrayReadRe              = regexp.MustCompile(`(?i)^\s*(?:[A-Za-z_]\w*\.)*read\s*\(\s*-1\s*\)\s*$`)
	arraySetupGuardRe                 = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s+then\s+exit\s+sub\s*$`)
	arrayStaticReadyGuardRe           = regexp.MustCompile(`(?i)^\s*if\s+not\s+([A-Za-z_]\w*)\s*\.\s*isset\s+then\s*$`)
	arrayModuleReadyGuardRe           = regexp.MustCompile(`(?i)^\s*if\s+not\s+([A-Za-z_]\w*)\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayOnErrorGotoRe                = regexp.MustCompile(`(?i)^\s*on\s+error\s+goto\s+([A-Za-z_]\w*)\s*$`)
	arrayOnErrorResumeNextRe          = regexp.MustCompile(`(?i)^\s*on\s+error\s+resume\s+next\s*$`)
	arrayOnErrorResumeNextStatementRe = regexp.MustCompile(`(?i)(?:^|\bthen\s+)on\s+error\s+resume\s+next(?:\s+else\b.*)?$`)
	arrayOnErrorGotoZeroRe            = regexp.MustCompile(`(?i)^\s*on\s+error\s+goto\s+0\s*$`)
	arrayErrNumberFailureRe           = regexp.MustCompile(`(?i)^\s*if\s+err\.number\s*<>\s*0\s+then\s*$`)
	arrayCapacityProbeRe              = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayBoundsProbeRe                = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*-\s*lbound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayCheckedProbeExitRe           = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*(?:<=|=)\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayCapacityIfRe                 = regexp.MustCompile(`(?i)^\s*if\s+.+\s*>\s*([A-Za-z_]\w*)\s+then\s*$`)
	arrayForZeroToCountRe             = regexp.MustCompile(`(?i)^\s*for\s+[A-Za-z_]\w*\s*=\s*0\s+to\s+[A-Za-z_]\w*\s*-\s*1\s*$`)
	arrayLabelRe                      = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*:\s*$`)
	arrayCountComparisonRe            = regexp.MustCompile(`(?i)^\s*(.*?)\s*(=|<>|>=|<=|>|<)\s*(-?\d+)\s*$`)
	arrayConditionAndRe               = regexp.MustCompile(`(?i)\s+and\s+`)
	arrayConditionOrRe                = regexp.MustCompile(`(?i)\s+or\s+`)
	arrayScalarConditionRe            = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*(=|<>)\s*([A-Za-z_]\w*|-?\d+)\s*$`)
	arrayScalarConditionReversedRe    = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>)\s*([A-Za-z_]\w*)\s*$`)
	arraySelectCaseRe                 = regexp.MustCompile(`(?i)^select\s+case\s+(.+)$`)
	arrayPositiveCaseRe               = regexp.MustCompile(`(?i)^case\s+(-?\d+)\s*$`)
)

func (a Analyzer) arrayLifecycleFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []Finding {
	variables := arrayVariables(file, proc, moduleDecls)
	capacityGuards := arrayResumeNextCapacityGuards(file, proc, variables)
	constants := arrayIntegerConstants(file, proc, a.visibleConstantValues, a.visibleConstants)
	return a.arrayLifecycleFindingsPrepared(file, proc, ctx, moduleDecls, variables, constants, capacityGuards)
}

// arrayLifecycleFindingsPrepared is used by the shared procedure-local array
// result. Its preparation inputs are explicit so enabled projections do not
// rebuild the same variable catalog, constant environment, or capacity guard
// facts independently.
func (a Analyzer) arrayLifecycleFindingsPrepared(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) []Finding {
	return a.arrayLifecycleFindingsPreparedWithRuntime(file, proc, ctx, moduleDecls, variables, constants, capacityGuards, nil)
}

// arrayLifecycleFindingsPreparedWithRuntime optionally collects the
// deterministic array runtime projection while the shared array worklist is
// already visiting the procedure. A non-nil sink removes the historical
// second VBA249 array walk without changing the standalone helper contract.
func (a Analyzer) arrayLifecycleFindingsPreparedWithRuntime(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, runtimeSink *[]Finding) []Finding {
	return a.arrayLifecycleFindingsPreparedWithRuntimeEntry(file, proc, ctx, moduleDecls, variables, constants, capacityGuards, runtimeSink, nil)
}

func (a Analyzer) arrayLifecycleFindingsPreparedWithRuntimeEntry(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, runtimeSink *[]Finding, preparedEntry arrayFlowState) []Finding {
	findings, _ := a.arrayLifecycleFindingsPreparedWithRuntimeEntryContext(context.Background(), file, proc, ctx, moduleDecls, variables, constants, capacityGuards, runtimeSink, preparedEntry)
	return findings
}

func (a Analyzer) arrayLifecycleFindingsPreparedWithRuntimeEntryContext(cancelCtx context.Context, file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, runtimeSink *[]Finding, preparedEntry arrayFlowState) ([]Finding, error) {
	if err := cancelCtx.Err(); err != nil {
		return nil, err
	}
	objectArrayDiagnosticsApplicable := false
	for _, variable := range variables {
		if variable.isArray && variable.isObject {
			objectArrayDiagnosticsApplicable = true
			break
		}
	}
	if !a.Config.Analyze.DetectArrayLifecycleSafety && !a.Config.Analyze.DetectRedimPreserveDimension && !a.Config.Analyze.DetectObjectArrayComparison && !objectArrayDiagnosticsApplicable && runtimeSink == nil {
		return nil, nil
	}
	vba227Variables := variables
	var vba227ResumeNextBefore []bool
	if a.Config.Analyze.DetectArrayLifecycleSafety {
		vba227Variables = arrayVBA227Variables(variables, file, proc)
		vba227ResumeNextBefore = arrayVBA227ResumeNextPrefixes(file, proc)
	}
	comparisonFindings := a.arrayComparisonFindings(file, proc, variables)
	comparisonFindings = append(comparisonFindings, a.arrayForEachFindings(file, proc, variables, ctx)...)
	baseLaneRequested := a.Config.Analyze.DetectArrayLifecycleSafety || a.Config.Analyze.DetectRedimPreserveDimension || objectArrayDiagnosticsApplicable
	initial := preparedEntry
	if initial == nil {
		initial = arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, variables)
	}
	runtimeBase := arrayOptionBase(file)
	if proc.Graph == nil {
		findings := append([]Finding(nil), comparisonFindings...)
		if !baseLaneRequested && runtimeSink == nil {
			findings = append(findings, a.arrayLifecycleLinearFindings(file, proc, ctx, variables, initial, constants, capacityGuards)...)
			return uniqueArrayFindings(findings), nil
		}
		seen := map[string]bool{}
		for _, finding := range findings {
			seen[arrayFindingKey(finding)] = true
		}
		vba227Initial := arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, vba227Variables)
		baseState := initial
		vba227State := vba227Initial
		runtimeState := arrayInitialState(variables)
		probe := a
		probe.Config.Analyze.DetectArrayLifecycleSafety = true
		runtimeSeen := map[string]bool{}
		for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
			if line&255 == 0 {
				if err := cancelCtx.Err(); err != nil {
					return nil, err
				}
			}
			text := normalizedCodeLine(file.Lines[line-1])
			if baseLaneRequested {
				var baseIssues []Finding
				baseState, baseIssues = a.arrayTransfer(file, proc, ctx, variables, baseState, text, line, constants, capacityGuards)
				for _, finding := range baseIssues {
					if a.Config.Analyze.DetectArrayLifecycleSafety && finding.Code == "VBA227" {
						continue
					}
					key := arrayFindingKey(finding)
					if !seen[key] {
						seen[key] = true
						findings = append(findings, finding)
					}
				}
				if a.Config.Analyze.DetectArrayLifecycleSafety {
					var lifecycleIssues []Finding
					vba227State, lifecycleIssues = a.arrayVBA227Transfer(file, proc, ctx, vba227Variables, vba227State, text, line, constants, capacityGuards, vba227ResumeNextBefore)
					for _, finding := range lifecycleIssues {
						if finding.Code != "VBA227" {
							continue
						}
						key := arrayFindingKey(finding)
						if !seen[key] {
							seen[key] = true
							findings = append(findings, finding)
						}
					}
				}
			}
			if runtimeSink != nil {
				for _, issue := range deterministicArrayRuntimeIssues(text, line, runtimeState, variables, constants, proc, runtimeBase) {
					key := strconv.Itoa(issue.line) + ":" + issue.kind + ":" + issue.operationKey
					if runtimeSeen[key] {
						continue
					}
					runtimeSeen[key] = true
					message, reason, suggestion := deterministicRuntimeFailureText(issue.kind)
					finding := a.simpleFinding(file, proc, issue.line, "VBA249", "error", message, reason, suggestion)
					finding.RuntimeError = &RuntimeErrorContext{Kind: issue.kind}
					finding.arrayOperationKey = issue.operationKey
					*runtimeSink = append(*runtimeSink, finding)
				}
				runtimeState, _ = probe.arrayTransfer(file, proc, ctx, variables, runtimeState, text, line, constants, nil)
			}
		}
		sortFindings(findings)
		if runtimeSink != nil {
			sortFindings(*runtimeSink)
		}
		return uniqueArrayFindings(findings), nil
	}

	findings := append([]Finding(nil), comparisonFindings...)
	seen := map[string]bool{}
	for _, finding := range findings {
		seen[arrayFindingKey(finding)] = true
	}
	baseView := proc.Graph.View(vbacfg.EdgeFilter{})
	lanes := make([]arrayCFGWorklistLane, 0, 3)
	if baseLaneRequested {
		lanes = append(lanes, arrayCFGWorklistLane{
			Graph: &baseView, Initial: initial, Stats: ctx.arrayStats,
			Visit: func(text string, line int, in arrayFlowState) arrayFlowState {
				out, issues := a.arrayTransfer(file, proc, ctx, variables, in, text, line, constants, capacityGuards)
				for _, call := range arrayCallsAtLine(proc.Calls, line) {
					out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
					out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
				}
				for _, finding := range issues {
					if finding.Code == "VBA227" {
						continue
					}
					key := arrayFindingKey(finding)
					if !seen[key] {
						seen[key] = true
						findings = append(findings, finding)
					}
				}
				return out
			},
			EdgeState: func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
				out = applyArrayConditionalAllocationBranch(out, &baseView, block, edge)
				out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
				return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
			},
		})
	}
	if a.Config.Analyze.DetectArrayLifecycleSafety {
		vba227Graph := arrayVBA227Graph(proc, ctx)
		vba227Initial := arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, vba227Variables)
		lanes = append(lanes, arrayCFGWorklistLane{
			Graph: &vba227Graph, Initial: vba227Initial, Stats: ctx.arrayStats, SourceLines: true,
			ReliableExceptional: func(statement *procedureir.Statement, in, out arrayFlowState) bool {
				return arrayAllocationTransferIsReliable(statement, in, out)
			},
			Visit: func(text string, line int, in arrayFlowState) arrayFlowState {
				out, issues := a.arrayVBA227Transfer(file, proc, ctx, vba227Variables, in, text, line, constants, capacityGuards, vba227ResumeNextBefore)
				for _, call := range arrayCallsAtLine(proc.Calls, line) {
					out = applyArrayModuleCallEffects(out, file, proc, call, ctx, vba227Variables, moduleDecls)
					out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, vba227Variables, moduleDecls)
				}
				for _, finding := range issues {
					if finding.Code != "VBA227" {
						continue
					}
					key := arrayFindingKey(finding)
					if !seen[key] {
						seen[key] = true
						findings = append(findings, finding)
					}
				}
				return out
			},
			EdgeState: func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
				out = applyArrayConditionalAllocationBranch(out, &vba227Graph, block, edge)
				out = applyArrayVBA227ConditionalReDimBranch(out, block.Statement, edge, vba227Variables)
				out = arraySuccessfulConditionState(out, block.Statement, vba227Variables, vba227ResumeNextBefore, proc)
				out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, vba227Variables)
				out = applyArraySafeBoundGuard(out, block.Statement, edge, ctx.arraySafeBoundGuards, vba227Variables)
				out = applyArrayForBoundState(out, block.Statement, edge, vba227Variables)
				return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], vba227Variables, file, proc, moduleDecls)
			},
		})
	}
	if runtimeSink != nil {
		probe := a
		probe.Config.Analyze.DetectArrayLifecycleSafety = true
		runtimeSeen := map[string]bool{}
		runtimeState := arrayInitialState(variables)
		lanes = append(lanes, arrayCFGWorklistLane{
			Graph: &baseView, Initial: runtimeState, Stats: ctx.arrayStats,
			Visit: func(text string, line int, in arrayFlowState) arrayFlowState {
				for _, issue := range deterministicArrayRuntimeIssues(text, line, in, variables, constants, proc, runtimeBase) {
					key := strconv.Itoa(issue.line) + ":" + issue.kind + ":" + issue.operationKey
					if runtimeSeen[key] {
						continue
					}
					runtimeSeen[key] = true
					message, reason, suggestion := deterministicRuntimeFailureText(issue.kind)
					finding := a.simpleFinding(file, proc, issue.line, "VBA249", "error", message, reason, suggestion)
					finding.RuntimeError = &RuntimeErrorContext{Kind: issue.kind}
					finding.arrayOperationKey = issue.operationKey
					*runtimeSink = append(*runtimeSink, finding)
				}
				out, _ := probe.arrayTransfer(file, proc, ctx, variables, in, text, line, constants, nil)
				return out
			},
		})
	}
	if err := walkArrayCFGCombined(cancelCtx, file.Lines, lanes); err != nil {
		return nil, err
	}
	if runtimeSink != nil {
		sortFindings(*runtimeSink)
	}
	sortFindings(findings)
	return findings, nil
}

func (a Analyzer) arrayForEachFindings(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, ctx analysisContext) []Finding {
	if !a.Config.Analyze.DetectArrayLifecycleSafety {
		return nil
	}
	var findings []Finding
	for statement := range proc.Statements.All() {
		if statement.Recovered || statement.Kind != procedureir.StatementForEach {
			continue
		}
		source := ""
		// The grammar has used both `value` and `collection` for the source
		// expression across parser versions.  Prefer the unambiguous statement
		// text when it matches the canonical For Each form, then fall back to
		// the IR value expression for recovered/alternate forms.
		if match := arrayForEachRe.FindStringSubmatch(statement.Text); len(match) > 0 {
			source = match[1]
		} else if statement.Value != nil {
			source = statement.Value.Text
		}
		source = strings.TrimSpace(strings.SplitN(source, "'", 2)[0])
		if !iterableSourceKnownInvalid(source, variables, arrayInitialState(variables), ctx) {
			continue
		}
		findings = append(findings, a.simpleFinding(file, proc, statement.Range.StartLine, "VBA227", "warning", strings.TrimSpace(source)+" is not a collection or array and cannot be used as a For Each source.", "For Each requires an iterable Collection or array value; this source is a known scalar.", "Iterate an array or Collection, or change the source expression to an iterable value."))
	}
	return findings
}

func uniqueArrayFindings(findings []Finding) []Finding {
	seen := map[string]bool{}
	out := make([]Finding, 0, len(findings))
	for _, finding := range findings {
		key := arrayFindingKey(finding)
		if seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, finding)
	}
	sortFindings(out)
	return out
}

func arrayFindingKey(finding Finding) string {
	return finding.Code + ":" + strconv.Itoa(finding.Line) + ":" + finding.Message
}

// arrayComparisonFindings inspects the shared expression IR instead of the
// CFG block text. A comparison's direct operand is the only position where
// an array is being used as a scalar; an array nested inside a call, member
// access, or indexed expression is a different semantic construct.
func (a Analyzer) arrayComparisonFindings(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable) []Finding {
	if !a.Config.Analyze.DetectObjectArrayComparison {
		return nil
	}
	facts := proc.analysisFacts()

	arrayNames := make([]string, 0, len(variables))
	nonArrayNames := procedureIRNonArrayNames(proc)
	for name, variable := range variables {
		if variable.isArray && !nonArrayNames[name] {
			arrayNames = append(arrayNames, name)
		}
	}
	sort.Strings(arrayNames)

	var findings []Finding
	facts.forEachExpression(func(comparison procedureir.Expression) {
		if comparison.SyntaxKind != "comparison_expression" || comparison.Recovered {
			return
		}
		statement, ok := facts.Statement(comparison.StatementID)
		if !ok || comparisonAssignmentCarrier(statement, comparison) {
			return
		}
		matched := map[string]bool{}
		for _, childID := range comparison.Children {
			child, ok := directComparisonOperand(facts, childID)
			if !ok || child.Recovered || child.Kind != procedureir.ExpressionIdentifier {
				continue
			}
			name := strings.ToLower(cleanIdentifier(child.Text))
			if variable, exists := variables[name]; exists && variable.isArray {
				matched[name] = true
			}
		}
		for _, name := range arrayNames {
			if !matched[name] {
				continue
			}
			variable := variables[name]
			findings = append(findings, a.simpleFinding(file, proc, comparison.Range.StartLine, "VBA209", "warning", variable.name+" appears to be compared as a scalar value.", "VBA arrays cannot be compared directly to scalar values.", "Compare explicit elements or bounds instead of the array variable itself."))
		}
	})
	sortFindings(findings)
	return findings
}

func procedureIRNonArrayNames(proc sourceProcedure) map[string]bool {
	names := map[string]bool{}
	for declaration := range proc.Declarations.All() {
		if declaration.Recovered || declaration.Name == "" {
			continue
		}
		isArray := declaration.IsArray
		if declaration.Kind == "parameter" && strings.Contains(declaration.Type, "()") {
			isArray = true
		}
		if !isArray {
			names[strings.ToLower(declaration.Name)] = true
		}
	}
	return names
}

func directComparisonOperand(facts *procedureAnalysisFacts, id int) (procedureir.Expression, bool) {
	child, ok := facts.Expression(id)
	for ok && child.Kind == procedureir.ExpressionParentheses && len(child.Children) == 1 {
		child, ok = facts.Expression(child.Children[0])
	}
	return child, ok
}

func comparisonAssignmentCarrier(statement procedureir.Statement, expression procedureir.Expression) bool {
	if expression.ParentID != 0 {
		return false
	}
	return statement.Kind == procedureir.StatementAssignment || statement.Kind == procedureir.StatementSet
}

func (a Analyzer) arrayLifecycleLinearFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) []Finding {
	var findings []Finding
	seen := map[string]bool{}
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		var issues []Finding
		state, issues = a.arrayTransfer(file, proc, ctx, variables, state, normalizedCodeLine(file.Lines[line-1]), line, constants, capacityGuards)
		for _, finding := range issues {
			key := finding.Code + ":" + strconv.Itoa(finding.Line) + ":" + finding.Message
			if !seen[key] {
				seen[key] = true
				findings = append(findings, finding)
			}
		}
	}
	sortFindings(findings)
	return findings
}

func (a Analyzer) arrayVBA227Transfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, resumeNextBefore []bool) (arrayFlowState, []Finding) {
	state = arrayVBA227ClearLoopBodyBounds(state, line)
	state = arrayVBA227ClearConditionalAllocationGuards(state, proc, text, line, variables)
	transfer := func(input arrayFlowState, source string) (arrayFlowState, []Finding) {
		output, findings := a.arrayTransfer(file, proc, ctx, variables, input, source, line, constants, capacityGuards)
		output = arrayVBA227AttachConditionalReDimState(output, proc, source, line, variables)
		output = arrayVBA227AttachReturnProvenance(output, source, ctx, variables, constants)
		if resumeNextBefore == nil || arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
			return output, findings
		}
		return output, arrayVBA227FilterNestedBoundIndexFindings(findings, source, variables)
	}
	if line >= 1 && line <= len(file.Lines) && vbaLineContinues(file.Lines[line-1]) && arrayVBA227HasArrayFactoryAssignment(text) {
		text = arrayLogicalCodeLine(file.Lines, line)
	}
	inlineText := text
	if line >= 1 && line <= len(file.Lines) {
		inlineText = normalizedCodeLine(file.Lines[line-1])
	}
	if redim, ok := inlineArrayRedimText(inlineText); ok {
		text = redim
	} else if assignment, ok := inlineArrayFactoryAssignmentText(inlineText); ok {
		text = assignment
	} else if assignment, ok := inlineArrayStrConvAssignmentText(inlineText); ok {
		text = assignment
	} else if assignment, ok := inlineArraySafeBoundAssignmentText(inlineText, ctx.arraySafeBoundGuards); ok {
		text = assignment
	} else if assignment, ok := inlineArrayDictionaryAssignmentText(inlineText); ok {
		text = assignment
	} else if assignment, ok := inlineArrayReturnAssignmentText(inlineText, ctx.arrayReturns); ok {
		text = assignment
	} else if assignment, ok := inlineArrayQualifiedReturnAssignmentText(file, proc, line, inlineText, ctx.arrayReturnsQualified); ok {
		text = assignment
	} else if assignment, ok := inlineArrayAssignmentText(inlineText); ok {
		text = assignment
	}
	if condition, body, ok := arrayIfThenParts(text); ok {
		if body != "" && !arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) && !arrayProcedureHasErrorHandling(proc) && arrayVBA227StatementAlwaysRaises(body) {
			if guardedState, safe := arrayNonEmptyGuardState(state, condition, variables); safe {
				_, findings := transfer(state, condition)
				return guardedState, findings
			}
		}
		if body != "" && !arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
			if guardedState, safe := arraySafeBoundBranchState(state, condition, vbacfg.EdgeBranchTrue, ctx.arraySafeBoundGuards, variables); safe {
				conditionState, findings := transfer(state, condition)
				thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
				thenState := guardedState
				if thenBody != "" {
					var thenFindings []Finding
					thenState, thenFindings = transfer(thenState, thenBody)
					findings = append(findings, thenFindings...)
				}
				elseState := conditionState
				if hasElse && elseBody != "" {
					var elseFindings []Finding
					elseState, elseFindings = transfer(elseState, elseBody)
					findings = append(findings, elseFindings...)
				}
				if hasElse {
					state = meetArrayState(thenState, elseState)
				} else {
					state = meetArrayState(thenState, conditionState)
				}
				return state, findings
			}
		}
		if argument, _, guard := arrayIsArrayGuardCondition(condition); guard && arrayElementBaseName(argument) != "" {
			state, findings := transfer(state, condition)
			guardedState := arrayElementGuardState(state, argument, variables)
			thenState := cloneArrayState(guardedState)
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			if thenBody != "" {
				var bodyFindings []Finding
				thenState, bodyFindings = transfer(thenState, thenBody)
				findings = append(findings, bodyFindings...)
			}
			if hasElse {
				elseState := cloneArrayState(guardedState)
				var elseFindings []Finding
				if elseBody != "" {
					elseState, elseFindings = transfer(elseState, elseBody)
				}
				findings = append(findings, elseFindings...)
				state = meetArrayState(thenState, elseState)
			} else {
				state = thenState
			}
			return state, findings
		}
		// A normal multi-line If evaluates its condition before either branch
		// can run. If a typed array's bounds query returns normally, the array
		// is allocated on both the true and false paths. Keep this refinement
		// narrow: ElseIf merging and inline bodies retain their existing CFG
		// handling, and Resume Next may continue after a failed query.
		if body == "" && strings.HasPrefix(strings.ToLower(strings.TrimSpace(condition)), "if ") && arrayVBA227HasPureBoundsCondition(condition) && !arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
			state, findings := transfer(state, condition)
			return arraySuccessfulBoundsState(state, condition, variables, arrayVBA227LoopBodyEndLine(proc, line)), findings
		}
	}
	state, findings := transfer(state, text)
	findings = arrayVBA227FilterForBodyIndexFindings(findings, proc, line, state, variables, resumeNextBefore)
	if (arrayVBA227HasSuccessfulBoundsExpression(text) || arrayVBA227HasDictionaryBoundsExpression(text, state)) &&
		!arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) &&
		!strings.Contains(strings.ToLower(text), "on error resume next") {
		state = arraySuccessfulBoundsState(state, text, variables, arrayVBA227LoopBodyEndLine(proc, line))
	}
	// Source-line CFG blocks can contain an If condition and its body. Apply
	// the normal-path fact after the condition while the block is still being
	// processed so a nested element access in the body does not repeat the
	// condition's possible outer-array failure.
	if argument, _, ok := arrayIsArrayGuardCondition(text); ok {
		state = arrayElementGuardState(state, argument, variables)
	}
	if name, ok := arraySafeArrayPointerGuardTarget(file, proc, line, text, variables); ok {
		if value, known := state[name]; known {
			value.kind = arrayAllocated
			value.knownArray = true
			// A nonzero SAFEARRAY descriptor proves that bounds can be
			// queried, but it does not prove that the descriptor contains an
			// element. Retain a possible-empty state so indexed access remains
			// checked even when the incoming ByRef state had no shape facts.
			value.mayBeEmpty = true
			state[name] = value
		}
	}
	return state, findings
}

// arraySafeArrayPointerGuardTarget recognizes the narrow low-level VBA idiom
// used to inspect a dynamic Byte-array descriptor without calling LBound or
// UBound first:
//
//	ptr = VarPtrArray(values)
//	If ptr = 0 Then Exit Function
//	CopyMemoryFromPtr pSA, ptr, LenB(pSA)
//	If pSA = 0 Then Exit Function
//
// The final guard's normal path has a nonzero SAFEARRAY descriptor, which is
// enough to make later bounds queries valid. Keep the contract contiguous and
// structural; a pointer-slot check alone, a different memory-copy shape, or a
// missing descriptor check must remain conservative.
func arraySafeArrayPointerGuardTarget(file parsedFile, proc sourceProcedure, line int, text string, variables map[string]arrayVariable) (string, bool) {
	if line <= proc.StartLine || line > proc.EndLine || line > len(file.Lines) {
		return "", false
	}
	guard := strings.TrimSpace(normalizedCodeLine(text))
	match := arraySafeArrayZeroExitGuardRe.FindStringSubmatch(guard)
	if len(match) != 2 {
		return "", false
	}
	descriptorName := strings.ToLower(cleanIdentifier(match[1]))
	previous := make([]string, 0, 3)
	for index := line - 2; index >= max(proc.StartLine-1, 0) && len(previous) < 3; index-- {
		candidate := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		if candidate == "" || strings.HasPrefix(candidate, "'") || strings.HasPrefix(candidate, "#") {
			continue
		}
		previous = append(previous, candidate)
	}
	if len(previous) != 3 {
		return "", false
	}
	// The scan above is backwards from the descriptor guard.
	copyText := previous[0]
	ptrGuard := previous[1]
	pointerAssignment := previous[2]
	copyMatch := arraySafeArrayPointerCopyRe.FindStringSubmatch(copyText)
	if len(copyMatch) != 4 || !strings.EqualFold(copyMatch[1], descriptorName) || !strings.EqualFold(copyMatch[1], copyMatch[3]) {
		return "", false
	}
	ptrName := strings.ToLower(cleanIdentifier(copyMatch[2]))
	ptrGuardMatch := arraySafeArrayZeroExitGuardRe.FindStringSubmatch(ptrGuard)
	if len(ptrGuardMatch) != 2 || !strings.EqualFold(ptrGuardMatch[1], ptrName) {
		return "", false
	}
	lhs, rhs, indexed, assigned := arrayAssignment(pointerAssignment)
	if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), ptrName) || !strings.EqualFold(arrayCallName(rhs), "varptrarray") {
		return "", false
	}
	open := firstParenOutsideString(rhs)
	if open < 0 {
		return "", false
	}
	close := matchingParen(rhs, open)
	if close < 0 || strings.TrimSpace(rhs[close+1:]) != "" {
		return "", false
	}
	arguments := splitArgs(rhs[open+1 : close])
	if len(arguments) != 1 {
		return "", false
	}
	arrayName := directArrayArgumentName(arguments[0])
	variable, known := variables[arrayName]
	if !known || !variable.isArray || !isByteArrayVariable(variable) {
		return "", false
	}
	return arrayName, true
}

// arrayVBA227FilterForBodyIndexFindings removes only the unallocated/empty
// index observations for a loop body whose For bound necessarily succeeded
// before the body could run. Source-line CFG blocks include a For header and
// its nested body in one scan, so the edge refinement is applied too late for
// the first body visit; this narrow filter preserves the bound finding itself
// and any known lower/upper-bound violation.
func arrayVBA227FilterForBodyIndexFindings(findings []Finding, proc sourceProcedure, line int, state arrayFlowState, variables map[string]arrayVariable, resumeNextBefore []bool) []Finding {
	if line <= 0 || arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
		return findings
	}
	proven := map[string]bool{}
	provenNonEmpty := map[string]bool{}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementFor || line <= statement.Range.StartLine || line >= statement.Range.EndLine {
			continue
		}
		header := strings.TrimSpace(statement.Text)
		if newline := strings.IndexAny(header, "\r\n"); newline >= 0 {
			header = strings.TrimSpace(header[:newline])
		}
		header = strings.TrimSpace(normalizedCodeLine(header))
		argument := ""
		if _, _, countSource, _, ok := arrayForCountHeader(header); ok {
			for name, value := range state {
				if value.allocationCountSource == "" || !arrayCountExpressionMatches(countSource, value.allocationCountSource) {
					continue
				}
				variable, known := variables[name]
				if known && (variable.isArray || variable.isVariant) {
					proven[name] = true
					provenNonEmpty[name] = true
				}
			}
		}
		if match := arrayForScalarBoundRe.FindStringSubmatch(header); len(match) == 2 {
			bound, known := state[strings.ToLower(cleanIdentifier(match[1]))]
			if known {
				argument = bound.safeBoundProbe
			}
		}
		name := strings.ToLower(cleanIdentifier(argument))
		variable, known := variables[name]
		if name != "" && known && (variable.isArray || variable.isVariant) {
			proven[name] = true
			provenNonEmpty[name] = true
		}
		if match := arrayForUBoundRe.FindStringSubmatch(header); len(match) == 3 {
			name := strings.ToLower(cleanIdentifier(match[2]))
			variable, variableKnown := variables[name]
			_, valueKnown := state[name]
			start, startKnown := integerLiteral(match[1])
			if variableKnown && (variable.isArray || variable.isVariant) && valueKnown && startKnown && start >= 0 {
				// The For body is reachable only after UBound succeeded. When
				// the loop starts at a nonnegative index, an empty array cannot
				// reach the body (its upper bound is below the start). The known
				// lower bound, when available, additionally proves the index is
				// not below the array's lower bound.
				proven[name] = true
				provenNonEmpty[name] = true
			}
		}
	}
	if len(proven) == 0 && len(provenNonEmpty) == 0 {
		return findings
	}
	filtered := findings[:0]
	for _, finding := range findings {
		remove := false
		if finding.Code == "VBA227" {
			for name := range proven {
				if finding.arrayOperationKey == arrayIndexOperationKey(name, "unallocated") ||
					provenNonEmpty[name] && finding.arrayOperationKey == arrayIndexOperationKey(name, "empty") {
					remove = true
					break
				}
			}
		}
		if !remove {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

func arrayVBA227StatementAlwaysRaises(text string) bool {
	trimmed := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{"err.raise ", "err.raise(", "call err.raise ", "call err.raise("} {
		if strings.HasPrefix(trimmed, prefix) {
			return true
		}
	}
	return trimmed == "err.raise" || trimmed == "call err.raise"
}

func arrayNonEmptyGuardState(state arrayFlowState, condition string, variables map[string]arrayVariable) (arrayFlowState, bool) {
	condition = strings.TrimSpace(condition)
	if strings.HasPrefix(strings.ToLower(condition), "if ") {
		condition = strings.TrimSpace(condition[3:])
	}
	match := arrayEmptyGuardRe.FindStringSubmatch(condition)
	if len(match) != 2 {
		return state, false
	}
	name := strings.ToLower(cleanIdentifier(match[1]))
	variable, known := variables[name]
	if !known || !variable.isArray {
		return state, false
	}
	value, known := state[name]
	if !known {
		return state, false
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	value.mayBeEmpty = false
	updated[name] = value
	return updated, true
}

// arrayVBA227FilterNestedBoundIndexFindings removes only the redundant
// unallocated-index observation from an expression such as
// data(LBound(data)). The bound query itself remains a finding: if it fails,
// the indexed expression is never evaluated. Keep the proof conservative when
// the same array has another indexed use on the source line, because the
// normalized finding range cannot distinguish those uses.
func arrayVBA227FilterNestedBoundIndexFindings(findings []Finding, text string, variables map[string]arrayVariable) []Finding {
	indexedCounts := make(map[string]int)
	proofCandidates := make(map[string]bool)
	for _, use := range arrayIndexedUsesForSource(text, variables) {
		if len(use.args) == 0 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(use.name))
		if name == "" {
			continue
		}
		indexedCounts[name]++
		variable, known := variables[name]
		if known && variable.isArray && arrayUseHasSelfBoundsQuery(use) {
			proofCandidates[name] = true
		}
	}
	boundFailures := make(map[string]bool)
	for _, finding := range findings {
		if finding.Code != "VBA227" {
			continue
		}
		for _, kind := range []string{"lbound", "ubound"} {
			for name := range proofCandidates {
				if finding.arrayOperationKey == arrayBoundOperationKey(kind, name, "unallocated") {
					boundFailures[name] = true
				}
			}
		}
	}
	proofKeys := make(map[string]bool)
	for name := range proofCandidates {
		if indexedCounts[name] == 1 && boundFailures[name] {
			proofKeys[arrayIndexOperationKey(name, "unallocated")] = true
		}
	}
	if len(proofKeys) == 0 {
		return findings
	}
	filtered := findings[:0]
	for _, finding := range findings {
		if finding.Code == "VBA227" && proofKeys[finding.arrayOperationKey] {
			continue
		}
		filtered = append(filtered, finding)
	}
	return filtered
}

func arrayUseHasSelfBoundsQuery(use arrayUse) bool {
	for _, argument := range use.args {
		for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(argument, -1) {
			if strings.EqualFold(cleanIdentifier(bound[2]), use.name) {
				return true
			}
		}
	}
	return false
}

func arrayVBA227HasPureBoundsCondition(text string) bool {
	hasLower, hasUpper := false, false
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		switch strings.ToLower(bound[1]) {
		case "lbound":
			hasLower = true
		case "ubound":
			hasUpper = true
		}
	}
	if !hasLower || !hasUpper {
		return false
	}
	condition := strings.TrimSpace(text)
	if !strings.HasPrefix(strings.ToLower(condition), "if ") {
		return false
	}
	condition = strings.TrimSpace(condition[len("if "):])
	condition = arrayBoundCallRe.ReplaceAllString(condition, "")
	condition = arrayBoundOperatorRe.ReplaceAllString(condition, "")
	for _, char := range condition {
		if isIdentifierStart(byte(char)) {
			return false
		}
	}
	return true
}

func arrayVBA227HasSuccessfulBoundsExpression(text string) bool {
	trimmed := strings.ToLower(strings.TrimSpace(text))
	// Conditional expressions do not dominate the following source line:
	// an ElseIf condition can be skipped when an earlier branch is taken.
	// Branch-specific allocation facts belong to the CFG edge transfer, not
	// to this normal-path statement refinement.
	if strings.HasPrefix(trimmed, "if ") || strings.HasPrefix(trimmed, "elseif ") || strings.HasPrefix(trimmed, "else if ") {
		return false
	}
	seen := make(map[string]uint8)
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		if name == "" {
			continue
		}
		if strings.EqualFold(bound[1], "lbound") {
			seen[name] |= 1
		} else {
			seen[name] |= 2
		}
	}
	for _, bounds := range seen {
		if bounds != 0 {
			return true
		}
	}
	return false
}

func arrayVBA227LoopBodyEndLine(proc sourceProcedure, line int) int {
	if line <= 0 {
		return 0
	}
	endLine := 0
	for statement := range proc.Statements.All() {
		switch statement.Kind {
		case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile:
		default:
			continue
		}
		if line > statement.Range.StartLine && line < statement.Range.EndLine &&
			(endLine == 0 || statement.Range.EndLine < endLine) {
			endLine = statement.Range.EndLine
		}
	}
	return endLine
}

func arrayVBA227AttachConditionalReDimState(state arrayFlowState, proc sourceProcedure, text string, line int, variables map[string]arrayVariable) arrayFlowState {
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(state) == 0 || line <= 0 || len(match) == 0 {
		return state
	}
	guard, ok := arrayVBA227ConditionalReDimGuard(proc, line, variables)
	if !ok {
		return state
	}
	var updated arrayFlowState
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct {
			legacy := arrayRedimClauseRe.FindStringSubmatch(clause)
			if len(legacy) == 0 {
				continue
			}
			redim = directArrayRedimClause{name: legacy[1], dimensions: legacy[2]}
		}
		name := strings.ToLower(cleanIdentifier(redim.name))
		variable, knownVariable := variables[name]
		value, knownValue := state[name]
		if !knownVariable || !variable.isArray || !knownValue || value.kind != arrayAllocated || !value.knownArray {
			continue
		}
		if value.conditionalAllocationSource != "" && value.conditionalAllocationSource != guard {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.conditionalAllocationSource = guard
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227ConditionalReDimGuard(proc sourceProcedure, line int, variables map[string]arrayVariable) (string, bool) {
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementReDim || statement.Range.StartLine != line {
			continue
		}
		parent := procedureStatementByID(proc, statement.ParentID)
		if parent.Kind != procedureir.StatementIf && parent.Kind != procedureir.StatementElseIf {
			return "", false
		}
		return arrayVBA227ScalarConditionSource(parent, variables)
	}
	return "", false
}

// arrayVBA227ClearConditionalAllocationGuards forgets a conditional ReDim fact
// when the scalar that controls it is assigned outside a matching guard. A
// later equality check is only useful if the value being checked is still the
// value that controlled the allocation; retaining the fact across an
// unrelated assignment would turn a path-sensitive proof into an unsound one.
func arrayVBA227ClearConditionalAllocationGuards(state arrayFlowState, proc sourceProcedure, text string, line int, variables map[string]arrayVariable) arrayFlowState {
	assigned, _, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return state
	}
	assignedName := strings.ToLower(cleanIdentifier(assigned))
	if assignedName == "" {
		return state
	}
	var updated arrayFlowState
	for name, value := range state {
		condition := value.conditionalAllocationSource
		if condition == "" || arrayVBA227ScalarConditionLHS(condition) != assignedName {
			continue
		}
		if arrayVBA227LineWithinMatchingGuard(proc, line, condition, variables) {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.conditionalAllocationSource = ""
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227LineWithinMatchingGuard(proc sourceProcedure, line int, condition string, variables map[string]arrayVariable) bool {
	statement := procedureStatementAtLine(proc, line)
	visited := map[int]bool{}
	for statement.ParentID != 0 && !visited[statement.ParentID] {
		visited[statement.ParentID] = true
		parent := procedureStatementByID(proc, statement.ParentID)
		if parent.ID == 0 {
			break
		}
		if parent.Kind == procedureir.StatementIf || parent.Kind == procedureir.StatementElseIf {
			if parentCondition, ok := arrayVBA227ScalarConditionSource(parent, variables); ok && parentCondition == condition {
				return true
			}
		}
		statement = parent
	}
	return false
}

func arrayVBA227ScalarConditionLHS(condition string) string {
	if match := arrayScalarConditionRe.FindStringSubmatch(condition); len(match) == 4 {
		return strings.ToLower(cleanIdentifier(match[1]))
	}
	if match := arrayScalarConditionReversedRe.FindStringSubmatch(condition); len(match) == 4 {
		return strings.ToLower(cleanIdentifier(match[3]))
	}
	return ""
}

func procedureStatementAtLine(proc sourceProcedure, line int) procedureir.Statement {
	var best procedureir.Statement
	for statement := range proc.Statements.All() {
		if line < statement.Range.StartLine || line > statement.Range.EndLine {
			continue
		}
		if best.ID == 0 || statement.Range.StartLine > best.Range.StartLine || statement.Range.EndLine < best.Range.EndLine {
			best = statement
		}
	}
	return best
}

func arrayVBA227ScalarConditionSource(statement procedureir.Statement, variables map[string]arrayVariable) (string, bool) {
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	condition = strings.TrimSpace(condition)
	lower := strings.ToLower(condition)
	if strings.HasPrefix(lower, "if ") {
		condition = strings.TrimSpace(condition[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		condition = strings.TrimSpace(condition[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(condition), " then"); then >= 0 && strings.TrimSpace(condition[then+5:]) == "" {
		condition = strings.TrimSpace(condition[:then])
	}
	for len(condition) >= 2 && condition[0] == '(' && condition[len(condition)-1] == ')' {
		condition = strings.TrimSpace(condition[1 : len(condition)-1])
	}
	if arrayConditionAndRe.MatchString(condition) || arrayConditionOrRe.MatchString(condition) {
		return "", false
	}
	if match := arrayScalarConditionRe.FindStringSubmatch(condition); len(match) == 4 {
		return arrayVBA227NormalizeScalarCondition(match[1], match[2], match[3], variables)
	}
	if match := arrayScalarConditionReversedRe.FindStringSubmatch(condition); len(match) == 4 {
		return arrayVBA227NormalizeScalarCondition(match[3], match[2], match[1], variables)
	}
	return "", false
}

func arrayVBA227NormalizeScalarCondition(lhs, operator, rhs string, variables map[string]arrayVariable) (string, bool) {
	lhs = strings.ToLower(cleanIdentifier(lhs))
	rhs = strings.ToLower(strings.TrimSpace(rhs))
	variable, known := variables[lhs]
	if !known || variable.isArray || variable.isVariant || variable.isObject {
		return "", false
	}
	if rhsVariable, known := variables[rhs]; known && (rhsVariable.isArray || rhsVariable.isVariant || rhsVariable.isObject) {
		return "", false
	}
	return lhs + operator + rhs, true
}

func applyArrayVBA227ConditionalReDimBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue || statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf {
		return state
	}
	condition, ok := arrayVBA227ScalarConditionSource(*statement, variables)
	if !ok {
		return state
	}
	var updated arrayFlowState
	for name, value := range state {
		if value.conditionalAllocationSource != condition {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227HasDictionaryBoundsExpression(text string, state arrayFlowState) bool {
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		if !strings.EqualFold(bound[1], "ubound") {
			continue
		}
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		value, known := state[name]
		if known {
			if _, dictionary := arrayDictionaryCountSource(value.allocationCountSource); dictionary {
				return true
			}
		}
	}
	return false
}

func arraySuccessfulBoundsState(state arrayFlowState, text string, variables map[string]arrayVariable, loopEndLine int) arrayFlowState {
	// An explicitly declared Variant element array is still a known array, so a
	// successful bounds query establishes its allocation. An untyped Variant is
	// not marked isArray and remains conservative in the normal transfer path.
	var updated arrayFlowState
	dictionarySources := map[string]bool{}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		if strings.EqualFold(bound[1], "ubound") {
			if value, known := state[name]; known {
				if source, ok := arrayDictionaryCountSource(value.allocationCountSource); ok {
					dictionarySources[source] = true
				}
			}
		}
		variable, known := variables[name]
		if !known || !variable.isArray {
			continue
		}
		value, known := state[name]
		if !known {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value = arrayVBA227RecordBoundsProof(value, loopEndLine)
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	for name, value := range state {
		source, ok := arrayDictionaryCountSource(value.allocationCountSource)
		if !ok || !dictionarySources[source] {
			continue
		}
		variable, known := variables[name]
		if !known || !variable.isArray && !variable.isVariant {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value = updated[name]
		value = arrayVBA227RecordBoundsProof(value, loopEndLine)
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		value.allocationCountSource = ""
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227RecordBoundsProof(value arrayValue, loopEndLine int) arrayValue {
	if loopEndLine == 0 || value.boundsProof.loopEndLine != 0 || value.kind == arrayAllocated && value.knownArray {
		return value
	}
	value.boundsProof = arrayBoundsProof{
		loopEndLine:                      loopEndLine,
		priorKind:                        value.kind,
		priorKnownArray:                  value.knownArray,
		priorMayBeEmpty:                  value.mayBeEmpty,
		priorAllocationCount:             value.allocationCountSource,
		priorConditionalAllocationSource: value.conditionalAllocationSource,
	}
	return value
}

func arrayVBA227ClearLoopBodyBounds(state arrayFlowState, line int) arrayFlowState {
	if line <= 0 {
		return state
	}
	var updated arrayFlowState
	for name, value := range state {
		if value.boundsProof.loopEndLine != line {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.kind = value.boundsProof.priorKind
		value.knownArray = value.boundsProof.priorKnownArray
		value.mayBeEmpty = value.boundsProof.priorMayBeEmpty
		value.allocationCountSource = value.boundsProof.priorAllocationCount
		value.conditionalAllocationSource = value.boundsProof.priorConditionalAllocationSource
		value.boundsProof = arrayBoundsProof{}
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayValueKnownLowerBound(value arrayValue) (int, bool) {
	for _, dimensions := range [][]arrayDimension{value.dimensions, value.preserveShape} {
		if len(dimensions) == 0 || !dimensions[0].lower.known {
			continue
		}
		return dimensions[0].lower.value, true
	}
	return 0, false
}

func arraySuccessfulConditionState(state arrayFlowState, statement *procedureir.Statement, variables map[string]arrayVariable, resumeNextBefore []bool, proc sourceProcedure) arrayFlowState {
	if statement == nil || (statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf) {
		return state
	}
	line := statement.Range.StartLine
	if arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
		return state
	}
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if strings.TrimSpace(condition) == "" {
		return state
	}
	updated := state
	cloned := false
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(condition, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		variable, known := variables[name]
		if !known || !variable.isArray {
			continue
		}
		value, known := updated[name]
		if !known {
			continue
		}
		if !cloned {
			updated = cloneArrayState(state)
			cloned = true
		}
		value = arrayVBA227RecordBoundsProof(value, arrayVBA227LoopBodyEndLine(proc, line))
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return arraySuccessfulBoundsState(updated, condition, variables, arrayVBA227LoopBodyEndLine(proc, line))
}

// arrayVBA227ResumeNextPrefixes computes the conservative "may have seen
// Resume Next" fact once per procedure. A reset is intentionally not modeled
// here because the array worklist does not carry VBA's procedure-level error
// mode and a reset may be reachable only on one branch.
func arrayVBA227ResumeNextPrefixes(file parsedFile, proc sourceProcedure) []bool {
	prefixes := make([]bool, len(file.Lines)+1)
	mayHaveResumeNext := false
	start := max(1, proc.StartLine)
	end := min(len(file.Lines), proc.EndLine)
	for line := start; line <= end; line++ {
		prefixes[line] = mayHaveResumeNext
		for _, statement := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			if arrayOnErrorResumeNextStatementRe.MatchString(strings.TrimSpace(statement)) {
				mayHaveResumeNext = true
				break
			}
		}
	}
	return prefixes
}

func arrayVBA227ResumeNextBeforeLine(prefixes []bool, line int) bool {
	return line >= 0 && line < len(prefixes) && prefixes[line]
}

func arrayIfThenParts(text string) (condition, body string, ok bool) {
	text = strings.TrimSpace(text)
	lower := strings.ToLower(text)
	prefixLength := 0
	switch {
	case strings.HasPrefix(lower, "if "):
		prefixLength = len("if ")
	case strings.HasPrefix(lower, "elseif "):
		prefixLength = len("elseif ")
	default:
		return "", "", false
	}
	rest := strings.TrimSpace(text[prefixLength:])
	then := arrayTopLevelKeywordIndex(rest, "then")
	if then < 0 {
		return "", "", false
	}
	return strings.TrimSpace(text[:prefixLength] + rest[:then]), strings.TrimSpace(rest[then+len("then"):]), true
}

func arrayIfThenBodyParts(body string) (thenBody, elseBody string, hasElse bool) {
	elseIndex := arrayTopLevelKeywordIndex(body, "else")
	if elseIndex < 0 {
		return strings.TrimSpace(body), "", false
	}
	return strings.TrimSpace(body[:elseIndex]), strings.TrimSpace(body[elseIndex+len("else"):]), true
}

func arrayTopLevelKeywordIndex(text, keyword string) int {
	if keyword == "" {
		return -1
	}
	depth := 0
	inString := false
	for i := 0; i <= len(text)-len(keyword); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			continue
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		}
		if inString || depth != 0 || !strings.EqualFold(text[i:i+len(keyword)], keyword) {
			continue
		}
		if i > 0 && isIdentifierPart(text[i-1]) {
			continue
		}
		end := i + len(keyword)
		if end < len(text) && isIdentifierPart(text[end]) {
			continue
		}
		return i
	}
	return -1
}

func arrayVBA227HasArrayFactoryAssignment(text string) bool {
	_, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(rhs))
	return strings.HasPrefix(lower, "array(") || strings.HasPrefix(lower, "split(") || strings.HasPrefix(lower, "filter(")
}

func arrayLogicalCodeLine(lines []string, line int) string {
	if line < 1 || line > len(lines) {
		return ""
	}
	logical := ""
	for index := line; index <= len(lines); index++ {
		raw := lines[index-1]
		part := strings.TrimSpace(normalizedCodeLine(raw))
		if strings.HasSuffix(part, "_") {
			part = strings.TrimSpace(strings.TrimSuffix(part, "_"))
		}
		if part != "" {
			if logical != "" {
				logical += " "
			}
			logical += part
		}
		if !vbaLineContinues(raw) {
			break
		}
	}
	return logical
}

func arrayLogicalSourceLine(lines []string, line int) string {
	if line < 1 || line > len(lines) {
		return ""
	}
	logical := ""
	for index := line; index <= len(lines); index++ {
		part := strings.TrimSpace(arraySourceOrderStripComment(lines[index-1]))
		if strings.HasSuffix(part, "_") {
			part = strings.TrimSpace(strings.TrimSuffix(part, "_"))
		}
		if part != "" {
			if logical != "" {
				logical += " "
			}
			logical += part
		}
		if !vbaLineContinues(lines[index-1]) {
			break
		}
	}
	return logical
}

func inlineArrayRedimText(text string) (string, bool) {
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return "", false
	}
	prefix := strings.TrimSpace(strings.ToLower(text[:colon]))
	if !strings.HasPrefix(prefix, "dim ") && !strings.HasPrefix(prefix, "static ") {
		return "", false
	}
	redim := strings.TrimSpace(text[colon+1:])
	if next := strings.IndexByte(redim, ':'); next >= 0 {
		redim = strings.TrimSpace(redim[:next])
	}
	if !strings.HasPrefix(strings.ToLower(redim), "redim ") {
		return "", false
	}
	return redim, true
}

func inlineArrayFactoryAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	switch arrayCallName(rhs) {
	case "array", "split", "filter":
		return remainder, true
	default:
		return "", false
	}
}

func inlineArrayDeclarationRemainder(text string) (string, bool) {
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return "", false
	}
	prefix := strings.TrimSpace(strings.ToLower(text[:colon]))
	if !strings.HasPrefix(prefix, "dim ") && !strings.HasPrefix(prefix, "static ") {
		return "", false
	}
	remainder := strings.TrimSpace(text[colon+1:])
	if remainder == "" {
		return "", false
	}
	return remainder, true
}

func inlineArrayAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, _, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	return remainder, true
}

func inlineArrayStrConvAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed || !strings.EqualFold(arrayCallName(rhs), "strconv") {
		return "", false
	}
	return remainder, true
}

func inlineArraySafeBoundAssignmentText(text string, guards map[string]bool) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed || !guards[arrayCallName(rhs)] {
		return "", false
	}
	return remainder, true
}

func inlineArrayDictionaryAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	_, _, ok = arrayDictionaryMemberParts(rhs)
	if !ok {
		return "", false
	}
	return remainder, true
}

func inlineArrayReturnAssignmentText(text string, returns map[string]arrayValue) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	value, known := returns[arrayCallName(rhs)]
	if !known || value.kind != arrayAllocated || !value.knownArray {
		return "", false
	}
	return remainder, true
}

func inlineArrayQualifiedReturnAssignmentText(file parsedFile, proc sourceProcedure, line int, text string, returns map[string]arrayValue) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	lhs, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	receiver, member, ok := arrayMemberCallParts(rhs)
	if !ok {
		return "", false
	}
	typeName := arrayTypeNameCaseAtLine(file, proc, line, receiver)
	if typeName == "" {
		return "", false
	}
	value, known := returns[strings.ToLower(typeName+"."+member)]
	if !known || value.kind != arrayAllocated || !value.knownArray {
		return "", false
	}
	// The qualified summary proves only array allocation here. Replace the
	// member call with a recognized array factory so the ordinary transfer can
	// carry that fact without guessing the returned shape.
	return lhs + " = Array()", true
}

func arrayMemberCallParts(text string) (receiver, member string, ok bool) {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		close := matchingParen(trimmed, open)
		if close < 0 || strings.TrimSpace(trimmed[close+1:]) != "" {
			return "", "", false
		}
		trimmed = strings.TrimSpace(trimmed[:open])
	}
	dot := strings.LastIndexByte(trimmed, '.')
	if dot <= 0 || dot >= len(trimmed)-1 {
		return "", "", false
	}
	receiver = cleanIdentifier(strings.TrimSpace(trimmed[:dot]))
	member = cleanIdentifier(strings.TrimSpace(trimmed[dot+1:]))
	if !arrayEraseNameRe.MatchString(receiver) || !arrayEraseNameRe.MatchString(member) {
		return "", "", false
	}
	return receiver, member, true
}

func arrayDictionaryMemberParts(text string) (receiver, member string, ok bool) {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		close := matchingParen(trimmed, open)
		if close < 0 || strings.TrimSpace(trimmed[close+1:]) != "" {
			return "", "", false
		}
		trimmed = strings.TrimSpace(trimmed[:open])
	}
	dot := strings.LastIndexByte(trimmed, '.')
	if dot < 0 || dot >= len(trimmed)-1 {
		return "", "", false
	}
	receiver = strings.TrimSpace(trimmed[:dot])
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(trimmed[dot+1:])))
	if member != "keys" && member != "items" {
		return "", "", false
	}
	if receiver == "" {
		if !strings.HasPrefix(trimmed, ".") {
			return "", "", false
		}
		return "", member, true
	}
	for _, part := range strings.Split(receiver, ".") {
		if !arrayEraseNameRe.MatchString(strings.TrimSpace(part)) {
			return "", "", false
		}
	}
	return receiver, member, true
}

func arrayDictionaryMemberExpressionState(file parsedFile, proc sourceProcedure, line int, rhs string, variables map[string]arrayVariable) (arrayValue, bool) {
	receiver, _, ok := arrayDictionaryMemberParts(rhs)
	knownNonEmpty := false
	if source := arrayLogicalSourceLine(file.Lines, line); source != "" {
		if _, sourceRHS, indexed, assigned := arrayAssignment(source); assigned && !indexed {
			knownNonEmpty = arrayDictionaryMemberKnownNonEmpty(file, line, sourceRHS)
		}
	}
	if !knownNonEmpty {
		knownNonEmpty = arrayDictionaryMemberKnownNonEmpty(file, line, rhs)
	}
	if !ok {
		if !knownNonEmpty {
			return arrayValue{}, false
		}
		receiver, _, _, _ = arrayDictionarySnapshotParts(rhs)
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, proc, line)
	}
	if receiver == "" || !knownNonEmpty && !arrayDictionaryReceiverProven(file, proc, line, receiver, variables) {
		return arrayValue{}, false
	}
	source := canonicalArrayBoundExpression(receiver)
	kind := arrayUnknown
	if knownNonEmpty {
		kind = arrayAllocated
	}
	return arrayValue{
		kind:                  kind,
		knownArray:            true,
		mayBeEmpty:            !knownNonEmpty,
		origin:                arrayOriginLocal,
		allocationCountSource: arrayDictionaryCountSourcePrefix + source,
	}, true
}

// arrayDictionaryMemberKnownNonEmpty recognizes the outer dictionary returned
// by a helper such as CreateLookupDict. The helper creates fixed members before
// it consumes its input, so Keys and Items on that outer dictionary always
// contain those members even when the input pair array is empty.
func arrayDictionaryMemberKnownNonEmpty(file parsedFile, line int, rhs string) bool {
	receiver, key, _, ok := arrayDictionarySnapshotParts(rhs)
	if !ok {
		return false
	}
	for procedure := range file.procedureView().All() {
		if !strings.EqualFold(procedure.Name, "CreateLookupDict") {
			continue
		}
		if arrayProcedureReturnsNonEmptyObjectMemberSet(file, procedure) &&
			arrayDictionaryMemberAssignmentUsesHelper(file, line, receiver, key) {
			return true
		}
	}
	return false
}

func arrayDictionarySnapshotParts(text string) (receiver, key, member string, ok bool) {
	canonical := canonicalArrayBoundExpression(text)
	lower := strings.ToLower(canonical)
	for _, candidate := range []string{"keys", "items"} {
		suffix := "." + candidate + "()"
		if !strings.HasSuffix(lower, suffix) {
			continue
		}
		prefix := canonical[:len(canonical)-len(suffix)]
		if len(prefix) == 0 || prefix[len(prefix)-1] != ')' {
			continue
		}
		open := -1
		for index := 0; index < len(prefix)-1; index++ {
			if prefix[index] == '(' && matchingParen(prefix, index) == len(prefix)-1 {
				open = index
				break
			}
		}
		if open <= 0 {
			continue
		}
		return strings.TrimSpace(prefix[:open]), strings.TrimSpace(prefix[open+1 : len(prefix)-1]), candidate, true
	}
	return "", "", "", false
}

func arrayProcedureReturnsNonEmptyObjectMemberSet(file parsedFile, procedure sourceProcedure) bool {
	returnedObject := ""
	start := max(0, procedure.StartLine-1)
	end := min(procedure.EndLine, len(file.Lines))
	for line := start; line < end; line++ {
		text := arrayLogicalSourceLine(file.Lines, line+1)
		if text == "" {
			continue
		}
		if lhs, rhs, assigned := arrayAssignmentSides(text); assigned && !strings.Contains(lhs, "(") && strings.EqualFold(cleanIdentifier(lhs), procedure.Name) {
			returnedObject = cleanIdentifier(rhs)
		}
	}
	if returnedObject == "" {
		return false
	}
	members := map[string]bool{}
	for line := start; line < end; line++ {
		text := arrayLogicalSourceLine(file.Lines, line+1)
		if text == "" {
			continue
		}
		base, memberKey, memberRHS, assigned := arrayMemberAssignmentParts(text)
		if !assigned || arrayCallName(memberRHS) != "createobject" {
			continue
		}
		if returnedObject != "" && !strings.EqualFold(cleanIdentifier(base), returnedObject) {
			continue
		}
		members[canonicalArrayBoundExpression(memberKey)] = true
	}
	return len(members) >= 2
}

func arrayDictionaryMemberAssignmentUsesHelper(file parsedFile, line int, receiver, key string) bool {
	wantReceiver := canonicalArrayBoundExpression(receiver)
	wantKey := canonicalArrayBoundExpression(key)
	assignedByHelper := false
	invalidBeforeSnapshot := false
	for sourceLine := 1; sourceLine <= len(file.Lines); sourceLine++ {
		text := arrayLogicalSourceLine(file.Lines, sourceLine)
		base, memberKey, memberRHS, assigned := arrayMemberAssignmentParts(text)
		if !assigned || canonicalArrayBoundExpression(base) != wantReceiver || canonicalArrayBoundExpression(memberKey) != wantKey {
			continue
		}
		if arrayCallName(memberRHS) == "createlookupdict" {
			// The initializing call may be in a different procedure that is
			// textually later in the module.  Keep the summary module-wide, but
			// let an earlier direct reassignment invalidate the fact for this
			// snapshot.
			assignedByHelper = true
		} else if sourceLine < line {
			invalidBeforeSnapshot = true
		}
	}
	return assignedByHelper && !invalidBeforeSnapshot
}

func arrayMemberAssignmentParts(text string) (receiver, key, rhs string, ok bool) {
	lhs, rhs, assigned := arrayAssignmentSides(text)
	if !assigned {
		return "", "", "", false
	}
	open := firstParenOutsideString(lhs)
	if open <= 0 {
		return "", "", "", false
	}
	close := matchingParen(lhs, open)
	if close != len(lhs)-1 {
		return "", "", "", false
	}
	return strings.TrimSpace(lhs[:open]), strings.TrimSpace(lhs[open+1 : close]), rhs, true
}

func arrayAssignmentSides(text string) (lhs, rhs string, ok bool) {
	trimmed := strings.TrimSpace(text)
	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "set ") || strings.HasPrefix(lower, "let ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	inString := false
	for index := 0; index < len(trimmed); index++ {
		switch trimmed[index] {
		case '"':
			if inString && index+1 < len(trimmed) && trimmed[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case '=':
			if inString || index > 0 && (trimmed[index-1] == '<' || trimmed[index-1] == '>' || trimmed[index-1] == '=') {
				continue
			}
			lhs = strings.TrimSpace(trimmed[:index])
			rhs = strings.TrimSpace(trimmed[index+1:])
			return lhs, rhs, lhs != "" && rhs != ""
		}
	}
	return "", "", false
}

func arrayWithReceiverAtLine(file parsedFile, proc sourceProcedure, line int) string {
	stack := make([]string, 0, 2)
	start := max(1, proc.StartLine)
	end := min(line-1, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[current-1]))
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if lower == "end with" {
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
			continue
		}
		if strings.HasPrefix(lower, "with ") {
			receiver := strings.TrimSpace(text[len("with "):])
			if receiver != "" {
				stack = append(stack, receiver)
			}
		}
	}
	if len(stack) == 0 {
		return ""
	}
	return stack[len(stack)-1]
}

func arrayDictionaryReceiverProven(file parsedFile, proc sourceProcedure, line int, receiver string, variables map[string]arrayVariable) bool {
	receiver = strings.TrimSpace(receiver)
	if receiver == "" {
		return false
	}
	if variable, known := variables[strings.ToLower(cleanIdentifier(receiver))]; known && isDictionaryType(variable.typ) {
		return true
	}
	if strings.EqualFold(arrayTypeNameCaseAtLine(file, proc, line, receiver), "Dictionary") {
		return true
	}
	if !strings.EqualFold(receiver, "This.children") || !strings.EqualFold(arraySelectCaseValueAtLine(file, proc, line, "This.iType"), "eJSONObject") {
		return false
	}
	for _, rawLine := range file.Lines {
		text := canonicalArrayBoundExpression(gui.StripComment(rawLine))
		if strings.Contains(text, "setthis.children=createdictionary") {
			return true
		}
	}
	return false
}

func arraySelectCaseValueAtLine(file parsedFile, proc sourceProcedure, line int, expression string) string {
	type frame struct {
		expression string
		caseValue  string
	}
	frames := make([]frame, 0, 2)
	want := canonicalArrayBoundExpression(expression)
	start := max(1, proc.StartLine)
	end := min(line, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.Join(strings.Fields(gui.StripComment(file.Lines[current-1])), " ")
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "select case ") {
			frames = append(frames, frame{expression: canonicalArrayBoundExpression(text[len("select case "):])})
			continue
		}
		if lower == "end select" {
			if len(frames) > 0 {
				frames = frames[:len(frames)-1]
			}
			continue
		}
		if !strings.HasPrefix(lower, "case ") || len(frames) == 0 {
			continue
		}
		caseText := strings.TrimSpace(text[len("case "):])
		if comma := strings.IndexByte(caseText, ','); comma >= 0 {
			caseText = strings.TrimSpace(caseText[:comma])
		}
		if strings.EqualFold(caseText, "else") {
			caseText = ""
		}
		frames[len(frames)-1].caseValue = caseText
	}
	for index := len(frames) - 1; index >= 0; index-- {
		if frames[index].expression == want {
			return frames[index].caseValue
		}
	}
	return ""
}

func arrayTypeNameCaseAtLine(file parsedFile, proc sourceProcedure, line int, receiver string) string {
	if receiver == "" || line <= 0 || len(file.Lines) == 0 {
		return ""
	}
	type frame struct {
		receiver string
		caseName string
	}
	frames := make([]frame, 0, 2)
	start := max(1, proc.StartLine)
	end := min(line, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.Join(strings.Fields(gui.StripComment(file.Lines[current-1])), " ")
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "select case ") {
			expression := strings.TrimSpace(text[len("select case "):])
			selectedReceiver := ""
			if match := arrayTypeNameExpressionRe.FindStringSubmatch(expression); len(match) == 2 {
				selectedReceiver = cleanIdentifier(match[1])
			}
			frames = append(frames, frame{receiver: selectedReceiver})
			continue
		}
		if strings.HasPrefix(lower, "end select") {
			if len(frames) > 0 {
				frames = frames[:len(frames)-1]
			}
			continue
		}
		if !strings.HasPrefix(lower, "case ") || len(frames) == 0 {
			continue
		}
		caseText := strings.TrimSpace(text[len("case "):])
		if match := arrayQuotedCaseRe.FindStringSubmatch(caseText); len(match) == 2 {
			frames[len(frames)-1].caseName = match[1]
		} else {
			frames[len(frames)-1].caseName = ""
		}
	}
	for index := len(frames) - 1; index >= 0; index-- {
		if strings.EqualFold(frames[index].receiver, receiver) {
			return frames[index].caseName
		}
	}
	return ""
}

func arrayResumeNextCapacityGuards(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable) []arrayResumeNextCapacityGuard {
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), proc.EndLine)
	if start >= end {
		return nil
	}
	lineText := func(index int) string {
		return strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
	}
	nextNonEmpty := func(index int) (int, string, bool) {
		for index < end {
			text := lineText(index)
			if text != "" {
				return index, text, true
			}
			index++
		}
		return 0, "", false
	}

	var guards []arrayResumeNextCapacityGuard
	for index := start; index < end; index++ {
		if !arrayOnErrorResumeNextRe.MatchString(lineText(index)) {
			continue
		}
		probeIndex := -1
		var capacityName, targetName string
		for candidate := index + 1; candidate < end; candidate++ {
			text := lineText(candidate)
			if arrayOnErrorGotoZeroRe.MatchString(text) {
				break
			}
			if match := arrayCapacityProbeRe.FindStringSubmatch(text); len(match) == 3 {
				probeIndex = candidate
				capacityName = strings.ToLower(match[1])
				targetName = strings.ToLower(match[2])
				break
			}
			if match := arrayBoundsProbeRe.FindStringSubmatch(text); len(match) == 4 && strings.EqualFold(match[2], match[3]) {
				probeIndex = candidate
				capacityName = strings.ToLower(match[1])
				targetName = strings.ToLower(match[2])
				break
			}
		}
		if probeIndex < 0 {
			continue
		}
		variable, known := variables[targetName]
		if !known || !variable.isArray {
			continue
		}
		if arrayResumeNextCapacityStartsAtZero(file, proc, variables, index, capacityName) {
			if restoreIndex, restoreText, ok := nextNonEmpty(probeIndex + 1); ok &&
				(arrayOnErrorGotoZeroRe.MatchString(restoreText) || arrayOnErrorGotoRe.MatchString(restoreText)) {
				if checkIndex, checkText, ok := nextNonEmpty(restoreIndex + 1); ok {
					if match := arrayCheckedProbeExitRe.FindStringSubmatch(checkText); len(match) == 2 && strings.EqualFold(match[1], capacityName) {
						indexStartLine := checkIndex + 2
						indexEndLine := end
						for candidate := indexStartLine - 1; candidate < end; candidate++ {
							text := lineText(candidate)
							if erase := arrayEraseRe.FindStringSubmatch(text); len(erase) == 2 && strings.EqualFold(strings.TrimSpace(erase[1]), targetName) {
								indexEndLine = candidate
								break
							}
						}
						guards = append(guards, arrayResumeNextCapacityGuard{
							target:         targetName,
							probeLine:      probeIndex + 1,
							indexStartLine: indexStartLine,
							indexEndLine:   indexEndLine,
						})
						continue
					}
				}
			}
		}
		errIndex, errText, ok := nextNonEmpty(probeIndex + 1)
		if !ok || !arrayErrNumberFailureRe.MatchString(errText) {
			continue
		}
		errEnd := arraySourceIfEnd(file.Lines, errIndex, end)
		if errEnd < 0 {
			continue
		}
		capacityZero := false
		errCleared := false
		for candidate := errIndex + 1; candidate < errEnd; candidate++ {
			text := lineText(candidate)
			if strings.EqualFold(text, "err.clear") {
				errCleared = true
			}
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if assigned && !indexed && strings.EqualFold(lhs, capacityName) && strings.TrimSpace(rhs) == "0" {
				capacityZero = true
			}
		}
		if !capacityZero || !errCleared {
			continue
		}
		restoreIndex, restoreText, ok := nextNonEmpty(errEnd + 1)
		if !ok || !arrayOnErrorGotoZeroRe.MatchString(restoreText) {
			continue
		}
		capacityIfIndex, capacityIfText, ok := nextNonEmpty(restoreIndex + 1)
		if !ok {
			continue
		}
		capacityMatch := arrayCapacityIfRe.FindStringSubmatch(capacityIfText)
		if len(capacityMatch) != 2 || !strings.EqualFold(capacityMatch[1], capacityName) {
			continue
		}
		capacityIfEnd := arraySourceIfEnd(file.Lines, capacityIfIndex, end)
		if capacityIfEnd < 0 {
			continue
		}
		preserveTarget := false
		for candidate := capacityIfIndex + 1; candidate < capacityIfEnd; candidate++ {
			match := arrayRedimRe.FindStringSubmatch(lineText(candidate))
			if len(match) == 0 || strings.TrimSpace(match[1]) == "" {
				continue
			}
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct && strings.EqualFold(redim.name, targetName) {
					preserveTarget = true
					break
				}
			}
			if preserveTarget {
				break
			}
		}
		if !preserveTarget {
			continue
		}
		forIndex, forText, ok := nextNonEmpty(capacityIfEnd + 1)
		if !ok || !arrayForZeroToCountRe.MatchString(forText) {
			continue
		}
		targetIndexed := false
		nextIndex := -1
		for candidate := forIndex + 1; candidate < end; candidate++ {
			text := lineText(candidate)
			lower := strings.ToLower(text)
			if lower == "next" || strings.HasPrefix(lower, "next ") {
				nextIndex = candidate
				break
			}
			for _, use := range arrayIndexedUses(text, variables) {
				if strings.EqualFold(use.name, targetName) && len(use.args) > 0 {
					targetIndexed = true
					break
				}
			}
		}
		if !targetIndexed || nextIndex < 0 {
			continue
		}
		guards = append(guards, arrayResumeNextCapacityGuard{
			target:         targetName,
			probeLine:      probeIndex + 1,
			indexStartLine: forIndex + 2,
			indexEndLine:   nextIndex,
		})
	}
	return guards
}

// arrayResumeNextCapacityStartsAtZero proves the only state that makes a
// Resume Next bounds probe safe to use as a guard: a failed assignment must
// leave the capacity at zero. A stale positive value would make `If capacity
// <= 0 Then Exit ...` incorrectly accept an unallocated array.
func arrayResumeNextCapacityStartsAtZero(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, resumeIndex int, capacityName string) bool {
	variable, known := variables[strings.ToLower(cleanIdentifier(capacityName))]
	if known && variable.static {
		return false
	}
	start := max(0, proc.StartLine-1)
	lineText := func(index int) string {
		if index < 0 || index >= len(file.Lines) {
			return ""
		}
		return strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
	}
	for index := resumeIndex - 1; index >= start; index-- {
		text := lineText(index)
		if text == "" {
			continue
		}
		if rhs, assigned := arrayScalarAssignment(text, capacityName); assigned {
			return strings.TrimSpace(rhs) == "0"
		}
		break
	}

	if !known || variable.parameter || variable.isArray || variable.isVariant || !variable.knownScalar {
		return false
	}
	declarationLine := 0
	for declaration := range proc.Declarations.All() {
		if declaration.Scope != procedureir.ScopeLocal || !strings.EqualFold(cleanIdentifier(declaration.Name), cleanIdentifier(capacityName)) || declaration.IsArray || !arrayKnownScalarType(declaration.Type) {
			continue
		}
		declarationLine = declaration.Range.StartLine
		break
	}
	if declarationLine == 0 || declarationLine > resumeIndex+1 {
		return false
	}
	for index := start; index < resumeIndex; index++ {
		text := lineText(index)
		if text == "" || index+1 == declarationLine {
			continue
		}
		if _, assigned := arrayScalarAssignment(text, capacityName); assigned || strings.Contains(strings.ToLower(text), strings.ToLower(cleanIdentifier(capacityName))) {
			return false
		}
	}
	return true
}

func arrayScalarAssignment(text, name string) (string, bool) {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return "", false
	}
	statements := strings.Split(text, ":")
	for index := len(statements) - 1; index >= 0; index-- {
		statement := statements[index]
		if strings.TrimSpace(statement) == "" {
			continue
		}
		lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(statement))
		if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
			return rhs, true
		}
		break
	}
	return "", false
}

func arraySourceIfEnd(lines []string, start, end int) int {
	depth := 0
	for index := start; index < end; index++ {
		text := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[index])))
		if text == "" {
			continue
		}
		if depth == 0 {
			depth = 1
			continue
		}
		if strings.HasPrefix(text, "if ") && strings.HasSuffix(text, " then") {
			depth++
			continue
		}
		if text == "end if" {
			depth--
			if depth == 0 {
				return index
			}
		}
	}
	return -1
}

func arrayResumeNextCapacityProbeApplies(guards []arrayResumeNextCapacityGuard, name string, line int) bool {
	for _, guard := range guards {
		if guard.probeLine == line && strings.EqualFold(guard.target, name) {
			return true
		}
	}
	return false
}

func arrayResumeNextCapacityIndexApplies(guards []arrayResumeNextCapacityGuard, name string, line int) bool {
	for _, guard := range guards {
		if line >= guard.indexStartLine && line <= guard.indexEndLine && strings.EqualFold(guard.target, name) {
			return true
		}
	}
	return false
}

func arrayResumeNextCapacityProofApplies(guards []arrayResumeNextCapacityGuard, name string, line int) bool {
	return arrayResumeNextCapacityProbeApplies(guards, name, line) || arrayResumeNextCapacityIndexApplies(guards, name, line)
}

// walkArrayCFG owns the common allocation-state worklist used by both the
// procedure findings pass and the array-return summary pass. Exceptional and
// uncertain edges retain the predecessor's input state because the statement
// may not have completed before control leaves the block.
func walkArrayCFG(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState) {
	walkArrayCFGWorklist(graph, lines, initial, visit, nil, nil, false)
}

func walkArrayCFGWithEdgesStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stats *arrayInterproceduralStats) {
	walkArrayCFGWithStopStats(graph, lines, initial, visit, edgeState, nil, stats)
}

func walkArrayCFGWithStopStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStats(graph, lines, initial, visit, edgeState, stop, false, stats)
}

func walkArrayCFGWithSourceLinesReliableStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStatsWithReliable(graph, lines, initial, visit, edgeState, nil, true, stats, reliableExceptional)
}

type arrayCFGBlockVisit func(block vbacfg.Block, text string, line int, in arrayFlowState) arrayFlowState

// arrayCFGWorklistReachable follows the same edge set as the array worklist.
// CFGView.IsReachable also expands unknown-flow sources for conservative
// diagnostics, but those synthetic reachability results do not cause the
// worklist to visit a disconnected nested statement block. The distinction is
// needed when source-line recovery has to attribute a call to its container.
func arrayCFGWorklistReachable(graph *vbacfg.CFGView) map[vbacfg.BlockID]bool {
	reachable := map[vbacfg.BlockID]bool{}
	if graph == nil {
		return reachable
	}
	entry := graph.Entry()
	reachable[entry] = true
	queue := []vbacfg.BlockID{entry}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if reachable[edge.To] {
				return true
			}
			reachable[edge.To] = true
			queue = append(queue, edge.To)
			return true
		})
	}
	return reachable
}

func arrayCFGBlockOwnsNestedStatements(block vbacfg.Block) bool {
	if block.Statement == nil {
		return false
	}
	switch block.Statement.Kind {
	case procedureir.StatementIf, procedureir.StatementElseIf, procedureir.StatementElse,
		procedureir.StatementSelect, procedureir.StatementCase,
		procedureir.StatementFor, procedureir.StatementForEach,
		procedureir.StatementDo, procedureir.StatementWhile, procedureir.StatementWith:
		return true
	default:
		return false
	}
}

func walkArrayCFGWithSourceLinesReliableStatsAndBlock(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, visitBlock arrayCFGBlockVisit, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStatsWithReliableAndBlock(graph, lines, initial, visit, visitBlock, edgeState, nil, true, stats, reliableExceptional)
}

func walkArrayCFGWorklist(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
	walkArrayCFGWorklistStats(graph, lines, initial, visit, edgeState, stop, sourceLines, nil)
}

func walkArrayCFGWorklistStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStatsWithReliable(graph, lines, initial, visit, edgeState, stop, sourceLines, stats, nil)
}

func walkArrayCFGWorklistStatsWithReliable(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool) {
	walkArrayCFGWorklistStatsWithReliableAndBlock(graph, lines, initial, visit, nil, edgeState, stop, sourceLines, stats, reliableExceptional)
}

func walkArrayCFGWorklistStatsWithReliableAndBlock(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, visitBlock arrayCFGBlockVisit, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool) {
	strategy := arrayCFGStrategyAuto
	if stats != nil {
		strategy = stats.strategy
	}
	if strategy == arrayCFGStrategyLegacy {
		if stats != nil {
			stats.addLegacyWalk()
		}
		walkArrayCFGWorklistLegacyWithBlock(graph, lines, initial, visit, visitBlock, edgeState, stop, sourceLines)
		return
	}
	// The Array adapter owns the policy-specific source-line, edge, and stop
	// semantics at the indexed solver boundary. Only index/solver construction
	// incompatibility falls back to the legacy map worklist; a solve-time error
	// is not retried after analysis has begun.
	// A zero-slot state cannot carry the declaration-only scalar/object checks
	// that still consult the variables side table from the transfer callback.
	// Keep those paths on the legacy adapter; this is an intentional
	// compatibility boundary rather than an attempt to enlarge the compact
	// lattice with non-array declarations.
	if graph != nil && len(initial) > 0 {
		if _, err := semanticstate.NewIndexView(*graph); err == nil {
			if stats != nil {
				stats.addCompactWalk()
			}
			if stop == nil && !sourceLines && edgeState == nil && reliableExceptional == nil {
				_ = walkArrayCFGCompact(context.Background(), graph, lines, initial, visit, nil, false)
			} else {
				_ = walkArrayCFGCompactAdvancedWithBlock(context.Background(), graph, lines, initial, visit, visitBlock, edgeState, stop, reliableExceptional, sourceLines)
			}
			return
		}
	}
	if stats != nil {
		stats.addLegacyWalk()
		if graph != nil {
			// A forced compact run is a test/benchmark oracle. Unsupported graph
			// or participant layouts use this same compatibility fallback; do not
			// retry the compact solver after callbacks have started.
			reason := arrayFallbackIndex
			if len(initial) == 0 {
				reason = arrayFallbackEmptyState
			}
			stats.addFallbackReason(reason)
		}
	}
	walkArrayCFGWorklistLegacyWithBlock(graph, lines, initial, visit, visitBlock, edgeState, stop, sourceLines)
}

func walkArrayCFGWorklistLegacy(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
	walkArrayCFGWorklistLegacyWithBlock(graph, lines, initial, visit, nil, edgeState, stop, sourceLines)
}

func walkArrayCFGWorklistLegacyWithBlock(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, visitBlock arrayCFGBlockVisit, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
	if graph == nil {
		return
	}
	inStates := map[vbacfg.BlockID]arrayFlowState{graph.Entry(): initial}
	queued := map[vbacfg.BlockID]bool{graph.Entry(): true}
	for len(queued) > 0 {
		// Block IDs are ordered so the worklist cannot inherit Go map iteration
		// order and change the fixed-point path through the array state lattice.
		var id vbacfg.BlockID
		first := true
		for candidate := range queued {
			if first || candidate < id {
				id = candidate
				first = false
			}
		}
		delete(queued, id)
		in := cloneArrayState(inStates[id])
		block, ok := graph.BlockByID(id)
		if !ok {
			continue
		}
		// Keep the predecessor state intact for exceptional/uncertain edges.
		// Transfer functions mutate their input state, so give the current
		// block its own copy once instead of cloning again in every transfer.
		out := cloneArrayState(in)
		visitLine := func(text string, line int, state arrayFlowState) arrayFlowState {
			if visitBlock != nil {
				return visitBlock(block, text, line, state)
			}
			return visit(text, line, state)
		}
		if block.Statement != nil {
			if !sourceLines {
				line := block.Statement.Range.StartLine
				if line == 0 {
					line = block.Range.StartLine
				}
				text := block.Statement.Text
				if strings.TrimSpace(text) == "" && line >= 1 && line <= len(lines) {
					text = normalizedCodeLine(lines[line-1])
				}
				// The transfer callback owns the block-local copy.  Keeping the
				// predecessor input untouched is required for exceptional and
				// uncertain edges, which deliberately propagate `in` below.
				out = visitLine(text, line, out)
				if stop != nil && stop(text, line) {
					continue
				}
			} else {
				start := block.Statement.Range.StartLine
				if start == 0 {
					start = block.Range.StartLine
				}
				end := block.Statement.Range.EndLine
				if end < start {
					end = start
				}
				stopped := false
				if (block.Statement.Kind == procedureir.StatementSelect || block.Statement.Kind == procedureir.StatementCase) && start >= 1 && start <= len(lines) {
					// Select Case and Case own separate CFG blocks for each branch.
					// Visiting a Case's whole source range here would scan nested
					// statements once before their branch edge facts are applied,
					// making a branch-local allocation fact appear to be absent.
					// The nested blocks below own the remaining physical lines.
					text := normalizedCodeLine(lines[start-1])
					out = visitLine(text, start, out)
					stopped = stop != nil && stop(text, start)
				} else if start == end && start >= 1 && start <= len(lines) {
					// A single physical line can still have multiple logical CFG
					// statements, for example `If ... Then ReDim ...`. Preserve the
					// block's own text so the ReDim block is not mistaken for the
					// surrounding If statement.
					text := block.Statement.Text
					if strings.TrimSpace(text) == "" {
						text = normalizedCodeLine(lines[start-1])
					}
					out = visitLine(text, start, out)
					stopped = stop != nil && stop(text, start)
				} else if start >= 1 && end <= len(lines) {
					// CFG blocks may contain an entire multi-line loop or conditional
					// statement. Process its physical source lines in order so an
					// allocation earlier in the block is visible to later bounds or
					// element accesses. The CFG still owns path joins and exceptional
					// edges; this only restores statement order within a block.
					for line := start; line <= end; line++ {
						text := normalizedCodeLine(lines[line-1])
						if strings.TrimSpace(text) == "" {
							continue
						}
						out = visitLine(text, line, out)
						if stop != nil && stop(text, line) {
							stopped = true
							break
						}
					}
				} else {
					text := block.Statement.Text
					if strings.TrimSpace(text) == "" && start >= 1 && start <= len(lines) {
						text = normalizedCodeLine(lines[start-1])
					}
					out = visitLine(text, start, out)
					stopped = stop != nil && stop(text, start)
				}
				if stopped {
					continue
				}
			}
		}
		graph.ForEachOutgoing(id, func(edge vbacfg.Edge) bool {
			next := out
			if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				next = in
				// The source-line VBA227 pass still models exceptional edges
				// conservatively, except when the statement itself establishes a
				// deterministic allocation. A valid plain ReDim (or a recognized
				// built-in array factory assignment) cannot leave the target in its
				// pre-allocation state merely because the procedure has
				// `On Error Resume Next`; keeping that predecessor state would make
				// every following indexed write in the same branch look unsafe.
				if sourceLines && edge.Class == vbacfg.EdgeExceptional && arrayAllocationTransferIsReliable(block.Statement, in, out) {
					next = out
				}
			} else if edgeState != nil {
				next = edgeState(block, edge, next)
			}
			if mergeArrayState(inStates, edge.To, next) {
				queued[edge.To] = true
			}
			return true
		})
	}
}

// arrayCFGWorklistLane describes one array state policy that can be advanced
// by walkArrayCFGCombined.  Lanes deliberately own their state and transfer
// callbacks: sharing a queue must not force the block-level, source-line, and
// runtime policies to share a merge or edge interpretation.
//
// Graph is normally the same graph for every lane.  It is allowed to be a
// policy-specific copy, however (for example arrayVBA227Graph removes
// impossible normal continuations).  Block IDs are stable across those copies
// and the worklist remains shared even when an edge is absent from one lane.
type arrayCFGWorklistLane struct {
	Graph   *vbacfg.CFGView
	Initial arrayFlowState
	Stats   *arrayInterproceduralStats

	// Visit receives a private copy of the lane's block input state and returns
	// the state to propagate to outgoing edges.  In source-line mode it is
	// called once for each physical line owned by the block.
	Visit func(text string, line int, in arrayFlowState) arrayFlowState

	// EdgeState applies a lane-specific normal-edge refinement (guards,
	// Select Case facts, module configuration, and so on). It is not called for
	// exceptional or uncertain edges, which retain the predecessor state.
	EdgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState

	// Stop can terminate propagation for this lane after visiting a statement.
	// A stopped lane does not prevent other lanes from processing the same block.
	Stop func(text string, line int) bool

	// ReliableExceptional permits a source-line lane to carry a deterministic
	// allocation across an exceptional edge. The historical default is the
	// predecessor state; callers should set this only for the existing narrow
	// reliable-allocation contract.
	ReliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool

	// SourceLines restores physical source order inside a CFG block.  False
	// retains the historical single-statement block semantics.
	SourceLines bool
}

// walkArrayCFGCombined advances multiple array policies with one deterministic
// block/edge index and one queue. Each lane still has an independent state
// lattice and callback policy, preserving semantic compatibility while
// avoiding repeated graph scheduling and indexing work.
//
// Exceptional and uncertain edges retain each lane's predecessor input state.
// A lane may opt into the existing reliable-allocation exception through
// ReliableExceptional; this is intentionally independent from SourceLines so
// runtime or other block-level lanes cannot accidentally inherit it.
func walkArrayCFGCombined(ctx context.Context, lines []string, lanes []arrayCFGWorklistLane) error {
	if len(lanes) == 0 {
		return ctx.Err()
	}
	strategy := arrayCFGStrategyAuto
	for _, lane := range lanes {
		if lane.Stats == nil || lane.Stats.strategy == arrayCFGStrategyAuto {
			continue
		}
		strategy = lane.Stats.strategy
		break
	}
	if strategy != arrayCFGStrategyLegacy {
		if handled, err := walkArrayCFGCombinedCompact(ctx, lines, lanes); handled {
			return err
		}
	}
	// A forced compact run remains observable through the same compatibility
	// fallback when a lane cannot be indexed; no second compact attempt is
	// made after callbacks have started.
	fallbackReason := arrayFallbackUnsupported
	for _, lane := range lanes {
		if len(lane.Initial) == 0 {
			fallbackReason = arrayFallbackEmptyState
			break
		}
		if lane.Graph != nil {
			if _, err := semanticstate.NewIndexView(*lane.Graph); err != nil {
				fallbackReason = arrayFallbackIndex
				break
			}
		}
	}
	seenStats := map[*arrayInterproceduralStats]bool{}
	for _, lane := range lanes {
		if lane.Stats == nil || seenStats[lane.Stats] {
			continue
		}
		seenStats[lane.Stats] = true
		lane.Stats.addLegacyWalk()
		if strategy != arrayCFGStrategyLegacy {
			lane.Stats.addFallbackReason(fallbackReason)
		}
	}

	type graphIndex struct {
		graph *vbacfg.CFGView
	}
	type laneIndex struct {
		graphIndex
		inStates map[vbacfg.BlockID]arrayFlowState
	}
	indexes := make([]laneIndex, len(lanes))
	graphIndexes := make(map[*vbacfg.CFGView]graphIndex, len(lanes))
	queued := map[vbacfg.BlockID]bool{}
	for index, lane := range lanes {
		if lane.Graph == nil {
			continue
		}
		shared, ok := graphIndexes[lane.Graph]
		if !ok {
			shared = graphIndex{graph: lane.Graph}
			graphIndexes[lane.Graph] = shared
		}
		inStates := map[vbacfg.BlockID]arrayFlowState{
			lane.Graph.Entry(): cloneArrayState(lane.Initial),
		}
		indexes[index] = laneIndex{graphIndex: shared, inStates: inStates}
		queued[lane.Graph.Entry()] = true
	}

	for len(queued) > 0 {
		if err := ctx.Err(); err != nil {
			return err
		}
		// Block IDs are ordered so convergence does not depend on Go map
		// iteration order. A shared queue is safe even when a policy-specific
		// graph omits an edge: that lane simply has no outgoing edge to merge.
		var id vbacfg.BlockID
		first := true
		for candidate := range queued {
			if first || candidate < id {
				id = candidate
				first = false
			}
		}
		delete(queued, id)

		for index, lane := range lanes {
			stateIndex := &indexes[index]
			if lane.Graph == nil || lane.Visit == nil {
				continue
			}
			in, ok := stateIndex.inStates[id]
			if !ok {
				continue
			}
			block, ok := stateIndex.graph.BlockByID(id)
			if !ok {
				continue
			}

			in = cloneArrayState(in)
			out := cloneArrayState(in)
			stopped := false
			if block.Statement != nil {
				if !lane.SourceLines {
					line := block.Statement.Range.StartLine
					if line == 0 {
						line = block.Range.StartLine
					}
					text := block.Statement.Text
					if strings.TrimSpace(text) == "" && line >= 1 && line <= len(lines) {
						text = normalizedCodeLine(lines[line-1])
					}
					// Keep the predecessor input immutable to the block transfer;
					// exceptional and uncertain edges must receive that input state.
					out = lane.Visit(text, line, out)
					stopped = lane.Stop != nil && lane.Stop(text, line)
				} else {
					start := block.Statement.Range.StartLine
					if start == 0 {
						start = block.Range.StartLine
					}
					end := block.Statement.Range.EndLine
					if end < start {
						end = start
					}
					if (block.Statement.Kind == procedureir.StatementSelect || block.Statement.Kind == procedureir.StatementCase) && start >= 1 && start <= len(lines) {
						// Select Case and Case own separate CFG blocks for each
						// branch; do not scan all clause lines before applying the
						// edge fact.
						text := normalizedCodeLine(lines[start-1])
						out = lane.Visit(text, start, out)
						stopped = lane.Stop != nil && lane.Stop(text, start)
					} else if start == end && start >= 1 && start <= len(lines) {
						text := block.Statement.Text
						if strings.TrimSpace(text) == "" {
							text = normalizedCodeLine(lines[start-1])
						}
						out = lane.Visit(text, start, out)
						stopped = lane.Stop != nil && lane.Stop(text, start)
					} else if start >= 1 && end <= len(lines) {
						for line := start; line <= end; line++ {
							if line&255 == 0 {
								if err := ctx.Err(); err != nil {
									return err
								}
							}
							text := normalizedCodeLine(lines[line-1])
							if strings.TrimSpace(text) == "" {
								continue
							}
							out = lane.Visit(text, line, out)
							if lane.Stop != nil && lane.Stop(text, line) {
								stopped = true
								break
							}
						}
					} else {
						text := block.Statement.Text
						if strings.TrimSpace(text) == "" && start >= 1 && start <= len(lines) {
							text = normalizedCodeLine(lines[start-1])
						}
						out = lane.Visit(text, start, out)
						stopped = lane.Stop != nil && lane.Stop(text, start)
					}
				}
			}
			if stopped {
				continue
			}

			stateIndex.graph.ForEachOutgoing(id, func(edge vbacfg.Edge) bool {
				next := out
				if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
					next = in
					if lane.SourceLines && edge.Class == vbacfg.EdgeExceptional && lane.ReliableExceptional != nil && lane.ReliableExceptional(block.Statement, in, out) {
						next = out
					}
				} else if lane.EdgeState != nil {
					next = lane.EdgeState(block, edge, next)
				}
				if mergeArrayState(stateIndex.inStates, edge.To, next) {
					queued[edge.To] = true
				}
				return true
			})
		}
	}
	return nil
}

func arrayAllocationTransferIsReliable(statement *procedureir.Statement, in, out arrayFlowState) bool {
	if statement == nil || !arrayStateAddsAllocation(in, out) {
		return false
	}
	text := strings.TrimSpace(statement.Text)
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		// ReDim Preserve can fail when the prior allocation or shape is
		// unknown. A plain ReDim is the deterministic allocation boundary.
		return strings.TrimSpace(match[1]) == ""
	}
	_, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(rhs))
	return strings.HasPrefix(lower, "array(") || arrayCallName(rhs) == "split" || arrayCallName(rhs) == "filter"
}

func arrayStateAddsAllocation(in, out arrayFlowState) bool {
	for name, value := range out {
		if value.kind != arrayAllocated || !value.knownArray {
			continue
		}
		before, ok := in[name]
		if !ok || before.kind != arrayAllocated || !before.knownArray {
			return true
		}
	}
	return false
}

func arrayCountExpressionMatches(expression, source string) bool {
	if strings.TrimSpace(expression) == "" || strings.TrimSpace(source) == "" {
		return false
	}
	expression = canonicalArrayBoundExpression(expression)
	source = canonicalArrayBoundExpression(arrayCountSourceExpression(source))
	return expression == source || expression == source+".count"
}

func arrayCountSourceExpression(source string) string {
	source = strings.TrimSpace(source)
	if strings.HasPrefix(strings.ToLower(source), arrayDictionaryCountSourcePrefix) {
		return strings.TrimSpace(source[len(arrayDictionaryCountSourcePrefix):])
	}
	return source
}

func arrayDictionaryCountSource(source string) (string, bool) {
	source = strings.TrimSpace(source)
	if !strings.HasPrefix(strings.ToLower(source), arrayDictionaryCountSourcePrefix) {
		return "", false
	}
	expression := strings.TrimSpace(source[len(arrayDictionaryCountSourcePrefix):])
	if expression == "" {
		return "", false
	}
	return canonicalArrayBoundExpression(expression), true
}

func applyArrayConditionalAllocationBranch(state arrayFlowState, graph *vbacfg.CFGView, block vbacfg.Block, edge vbacfg.Edge) arrayFlowState {
	if block.Statement == nil {
		return state
	}
	if block.Statement.Kind == procedureir.StatementSelect && edge.Kind == vbacfg.EdgeCase && graph != nil {
		selectExpression := selectCaseExpression(block.Statement.Text)
		caseBlock, ok := graph.BlockByID(edge.To)
		if !ok {
			return state
		}
		caseOK := positiveSelectCaseValue(caseBlock.Statement)
		if !caseOK || strings.TrimSpace(selectExpression) == "" {
			return state
		}
		for name, value := range state {
			if value.allocationCountSource == "" || !arrayCountExpressionMatches(selectExpression, value.allocationCountSource) {
				continue
			}
			value.kind = arrayAllocated
			value.knownArray = true
			value.allocationCountSource = ""
			state[name] = value
		}
		return state
	}
	if block.Statement.Kind != procedureir.StatementIf && block.Statement.Condition == nil {
		return state
	}
	condition := ""
	if block.Statement.Condition != nil {
		condition = block.Statement.Condition.Text
	} else {
		condition = block.Statement.Text
	}
	comparisons := []string{condition}
	if arrayConditionAndRe.MatchString(condition) {
		comparisons = arrayConditionAndRe.Split(condition, -1)
	}
	for name, value := range state {
		for _, comparison := range comparisons {
			lhs, operator, literal, ok := arrayCountComparison(comparison)
			if !ok {
				continue
			}
			positiveBranch, ok := positiveArrayCountBranch(operator, literal)
			if !ok || edge.Kind != positiveBranch || value.allocationCountSource == "" || !arrayCountExpressionMatches(lhs, value.allocationCountSource) {
				continue
			}
			value.kind = arrayAllocated
			value.knownArray = true
			value.allocationCountSource = ""
			state[name] = value
			break
		}
	}
	return state
}

func selectCaseExpression(text string) string {
	first := strings.TrimSpace(strings.SplitN(text, "\n", 2)[0])
	match := arraySelectCaseRe.FindStringSubmatch(first)
	if len(match) != 2 {
		return ""
	}
	return strings.TrimSpace(match[1])
}

func positiveSelectCaseValue(statement *procedureir.Statement) bool {
	if statement == nil {
		return false
	}
	first := strings.TrimSpace(strings.SplitN(statement.Text, "\n", 2)[0])
	match := arrayPositiveCaseRe.FindStringSubmatch(first)
	if len(match) != 2 {
		return false
	}
	value, err := strconv.Atoi(match[1])
	return err == nil && value > 0
}

func arrayCountComparison(text string) (lhs, operator, literal string, ok bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "if ") {
		text = strings.TrimSpace(text[3:])
	}
	if then := strings.Index(strings.ToLower(text), " then"); then >= 0 {
		text = strings.TrimSpace(text[:then])
	}
	match := arrayCountComparisonRe.FindStringSubmatch(text)
	if len(match) != 4 {
		return "", "", "", false
	}
	return strings.TrimSpace(match[1]), match[2], match[3], true
}

func positiveArrayCountBranch(operator, literal string) (vbacfg.EdgeKind, bool) {
	value, err := strconv.Atoi(literal)
	if err != nil {
		return "", false
	}
	switch operator {
	case "=":
		if value == 0 {
			return vbacfg.EdgeBranchFalse, true
		}
	case "<>":
		if value == 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">":
		if value >= 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">=":
		if value >= 1 {
			return vbacfg.EdgeBranchTrue, true
		}
	case "<":
		if value <= 1 {
			return vbacfg.EdgeBranchFalse, true
		}
	case "<=":
		if value <= 0 {
			return vbacfg.EdgeBranchFalse, true
		}
	}
	return "", false
}

// applyArrayAllocationGuard refines only the branch where a proven array
// length helper returns a positive value. The true branch of IsArray is also
// enough to establish array-ness for a Variant assignment. The opposite
// branch is left at its existing lattice value because an arbitrary caller
// may have additional paths or side effects that this rule cannot prove.
func applyArrayAllocationGuard(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, guards map[string]bool, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	if statement.Condition == nil {
		return state
	}
	if updated, ok := arrayStrPtrGuardState(state, statement.Condition.Text, edge.Kind, variables); ok {
		return arrayVBA227PropagateNonEmptyReturnInputs(updated, statement.Condition.Text, variables)
	}
	if argument, arrayBranch, ok := arrayIsArrayGuardCondition(statement.Condition.Text); ok {
		if name := directArrayArgumentName(argument); name != "" {
			if edge.Kind != arrayBranch {
				return state
			}
			variable, known := variables[name]
			if !known || !variable.isVariant {
				return state
			}
			value, known := state[name]
			if !known {
				return state
			}
			updated := cloneArrayState(state)
			value.kind = arrayAllocated
			value.knownArray = true
			updated[name] = value
			return updated
		}

		// Nested element guards are refined by arrayVBA227Transfer while the
		// source-line block is still being processed. Refining here would apply
		// the fact after later statements in the same block, potentially
		// overwriting an Erase or unknown assignment before the next block.
		return state
	}
	argument, allocatedBranch, ok := arrayAllocationGuardCondition(statement.Condition.Text, guards, state)
	if !ok || edge.Kind != allocatedBranch {
		return state
	}
	name := strings.ToLower(cleanIdentifier(argument))
	variable, known := variables[name]
	if !known || !variable.isArray && !variable.isVariant {
		return state
	}
	value, known := state[name]
	if !known {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	value.mayBeEmpty = false
	updated[name] = value
	return updated
}

// applyArraySafeBoundGuard refines the branch where a helper that returns an
// upper bound (or -1 after a caught bounds failure) proves that its array is
// allocated and has a nonnegative upper bound. This is separate from the
// positive-length helper contract because zero is a successful upper-bound
// result, not its failure sentinel.
func applyArraySafeBoundGuard(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, guards map[string]bool, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	if statement.Condition == nil {
		return state
	}
	updated, ok := arraySafeBoundBranchState(state, statement.Condition.Text, edge.Kind, guards, variables)
	if !ok {
		return state
	}
	return updated
}

// applyArrayForBoundState refines the path that can enter a For body after a
// scalar safe-bound helper has returned a nonnegative result. Its -1 sentinel
// intentionally represents a zero-iteration loop, so no fact is propagated to
// the loop-exit path. A direct UBound expression is left to the existing
// bound diagnostic; for source-line CFG blocks, that finding is also the
// conservative evidence for body accesses when the bound can fail.
func applyArrayForBoundState(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeLoopBody {
		return state
	}
	text := strings.TrimSpace(statement.Text)
	if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
		text = strings.TrimSpace(text[:newline])
	}
	text = strings.TrimSpace(normalizedCodeLine(text))
	if _, _, countSource, _, ok := arrayForCountHeader(text); ok {
		if updated, changed := arrayForCountArrayState(state, countSource, variables); changed {
			return updated
		}
	}
	match := arrayForScalarBoundRe.FindStringSubmatch(text)
	if len(match) != 2 {
		return state
	}
	bound, known := state[strings.ToLower(cleanIdentifier(match[1]))]
	if !known || bound.safeBoundProbe == "" {
		return state
	}
	return arrayForBoundArrayState(state, bound.safeBoundProbe, variables)
}

func arrayForCountHeader(text string) (loopVariable, start, countSource string, hasMinusOne bool, ok bool) {
	match := arrayForCountRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) != 5 {
		return "", "", "", false, false
	}
	if match[2] == "0" && strings.TrimSpace(match[4]) == "" {
		// `For i = 0 To items.Count` also enters when i == Count, so it
		// does not prove that an indexed access in the body is in range.
		return "", "", "", false, false
	}
	return match[1], match[2], match[3], strings.TrimSpace(match[4]) != "", true
}

func arrayForCountArrayState(state arrayFlowState, countSource string, variables map[string]arrayVariable) (arrayFlowState, bool) {
	var updated arrayFlowState
	for name, value := range state {
		if value.allocationCountSource == "" || !arrayCountExpressionMatches(countSource, value.allocationCountSource) {
			continue
		}
		variable, known := variables[name]
		if !known || !variable.isArray && !variable.isVariant {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value = updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		value.allocationCountSource = ""
		updated[name] = value
	}
	if updated == nil {
		return state, false
	}
	return updated, true
}

func arrayForBoundArrayState(state arrayFlowState, argument string, variables map[string]arrayVariable) arrayFlowState {
	name := strings.ToLower(cleanIdentifier(argument))
	variable, known := variables[name]
	if !known || !variable.isArray && !variable.isVariant {
		return state
	}
	value, known := state[name]
	if !known {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	updated[name] = value
	return updated
}

func arraySafeBoundBranchState(state arrayFlowState, text string, branch vbacfg.EdgeKind, guards map[string]bool, variables map[string]arrayVariable) (arrayFlowState, bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "if ") {
		text = strings.TrimSpace(text[3:])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	functionName, argument, operator, literal, reversed, ok := parseArrayAllocationGuard(text)
	if ok {
		functionName = strings.ToLower(lastName(functionName))
		if !guards[functionName] {
			return state, false
		}
	} else {
		argument, operator, literal, reversed, ok = parseArraySafeBoundProbeVariable(text, state)
		if !ok {
			return state, false
		}
	}
	value, err := strconv.Atoi(literal)
	if err != nil {
		return state, false
	}
	allocatedBranch, ok := safeBoundNonnegativeBranch(operator, value, reversed)
	if !ok || branch != allocatedBranch {
		return state, false
	}
	name := strings.ToLower(cleanIdentifier(argument))
	variable, known := variables[name]
	if !known || !variable.isArray {
		return state, false
	}
	current, known := state[name]
	if !known {
		return state, false
	}
	updated := cloneArrayState(state)
	current.kind = arrayAllocated
	current.knownArray = true
	current.mayBeEmpty = false
	updated[name] = current
	return updated, true
}

func safeBoundNonnegativeBranch(operator string, value int, reversed bool) (vbacfg.EdgeKind, bool) {
	if reversed {
		switch operator {
		case ">":
			operator = "<"
		case ">=":
			operator = "<="
		case "<":
			operator = ">"
		case "<=":
			operator = ">="
		}
	}
	switch operator {
	case "=":
		if value >= 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">":
		if value >= -1 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">=":
		if value >= 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case "<":
		if value >= 0 {
			return vbacfg.EdgeBranchFalse, true
		}
	case "<=":
		if value >= -1 {
			return vbacfg.EdgeBranchFalse, true
		}
	}
	return "", false
}

func arrayIsArrayGuardCondition(text string) (string, vbacfg.EdgeKind, bool) {
	text = strings.TrimSpace(text)
	if condition, _, ok := arrayIfThenParts(text); ok {
		text = condition
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "if ") {
		text = strings.TrimSpace(text[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		text = strings.TrimSpace(text[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	negated := false
	if strings.HasPrefix(strings.ToLower(text), "not ") {
		negated = true
		text = strings.TrimSpace(text[4:])
	}
	match := arrayIsArrayGuardRe.FindStringSubmatch(text)
	if len(match) == 2 {
		branch := vbacfg.EdgeBranchTrue
		if negated {
			branch = vbacfg.EdgeBranchFalse
		}
		return match[1], branch, true
	}
	if negated {
		return "", "", false
	}
	if match := arrayByteArrayGuardRe.FindStringSubmatch(text); len(match) == 2 {
		return match[1], vbacfg.EdgeBranchTrue, true
	}
	return "", "", false
}

func arrayStrPtrGuardParts(text string, variables map[string]arrayVariable) ([]string, string, bool) {
	text = strings.TrimSpace(text)
	if condition, _, ok := arrayIfThenParts(text); ok {
		text = condition
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "if ") {
		text = strings.TrimSpace(text[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		text = strings.TrimSpace(text[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	if arrayConditionAndRe.MatchString(text) {
		return nil, "", false
	}
	parts := arrayConditionOrRe.Split(text, -1)
	if len(parts) == 0 {
		return nil, "", false
	}
	names := make([]string, 0, len(parts))
	singleOperator := ""
	for _, part := range parts {
		match := arrayStrPtrGuardRe.FindStringSubmatch(strings.TrimSpace(part))
		if len(match) != 3 {
			return nil, "", false
		}
		if len(parts) > 1 && match[2] != "=" {
			return nil, "", false
		}
		if len(parts) == 1 {
			singleOperator = match[2]
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		variable, knownVariable := variables[name]
		if !knownVariable || !isByteArrayVariable(variable) {
			return nil, "", false
		}
		names = append(names, name)
	}
	return names, singleOperator, true
}

// arrayStrPtrGuardState recognizes the established VBA Byte-array idiom
// `If StrPtr(values) = 0 Then ...`: the false branch is reached only when the
// array has a usable element. A compound zero-pointer guard joined by Or has
// the same property for every operand on its false branch. Keep this limited
// to declared arrays; StrPtr on arbitrary Variants or scalar expressions does
// not establish an array contract for this rule.
func arrayStrPtrGuardState(state arrayFlowState, text string, branch vbacfg.EdgeKind, variables map[string]arrayVariable) (arrayFlowState, bool) {
	names, singleOperator, ok := arrayStrPtrGuardParts(text, variables)
	if !ok {
		return state, false
	}
	for _, name := range names {
		if _, knownValue := state[name]; !knownValue {
			return state, false
		}
	}
	if len(names) > 1 {
		if branch != vbacfg.EdgeBranchFalse {
			return state, false
		}
	} else {
		requiredBranch := vbacfg.EdgeBranchFalse
		if singleOperator == "<>" {
			requiredBranch = vbacfg.EdgeBranchTrue
		}
		if branch != requiredBranch {
			return state, false
		}
	}
	updated := cloneArrayState(state)
	for _, name := range names {
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		updated[name] = value
	}
	return updated, true
}

// arrayVBA227PropagateNonEmptyReturnInputs transfers a caller-side input-array
// fact through a returned Byte-array value. The transfer is intentionally
// driven by the StrPtr branch that already proved the returned value non-empty;
// a bare return summary never establishes the input fact by itself.
func arrayVBA227PropagateNonEmptyReturnInputs(state arrayFlowState, text string, variables map[string]arrayVariable) arrayFlowState {
	names, _, ok := arrayStrPtrGuardParts(text, variables)
	if !ok {
		return state
	}
	updated := cloneArrayState(state)
	changed := false
	for _, name := range names {
		value, known := updated[name]
		if !known || value.nonEmptySource == "" {
			continue
		}
		source := strings.ToLower(cleanIdentifier(value.nonEmptySource))
		variable, knownVariable := variables[source]
		sourceValue, knownSource := updated[source]
		if !knownVariable || !variable.isArray || !knownSource {
			continue
		}
		sourceValue.kind = arrayAllocated
		sourceValue.knownArray = true
		sourceValue.mayBeEmpty = false
		updated[source] = sourceValue
		changed = true
	}
	if !changed {
		return state
	}
	return updated
}

func arrayElementBaseName(text string) string {
	text = strings.TrimSpace(text)
	open := strings.IndexByte(text, '(')
	if open <= 0 || matchingParen(text, open) != len(text)-1 {
		return ""
	}
	if strings.TrimSpace(text[open+1:len(text)-1]) == "" {
		return ""
	}
	return directArrayArgumentName(text[:open])
}

func arrayElementGuardState(state arrayFlowState, argument string, variables map[string]arrayVariable) arrayFlowState {
	name := arrayElementBaseName(argument)
	if name == "" {
		return state
	}
	variable, known := variables[name]
	if !known || !variable.isArray {
		return state
	}
	value, known := state[name]
	if !known {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	value.mayBeEmpty = false
	updated[name] = value
	return updated
}

// arrayVBA227Graph removes normal-flow edges after direct raises and after
// private helpers whose normal CFG has no path to the procedure exit.  The
// latter covers project-local error wrappers such as RaiseContractError:
// their call sites must not poison the normal allocation state with an
// impossible fall-through branch.
func arrayVBA227Graph(proc sourceProcedure, ctx analysisContext) vbacfg.CFGView {
	if proc.Graph == nil {
		return vbacfg.CFGView{}
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	removed := map[vbacfg.BlockID]bool{}
	for call := range proc.Calls.All() {
		_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !ok || !arrayProcedureAlwaysRaises(target) {
			continue
		}
		block, ok := graph.BlockForStatement(call.StatementID)
		if ok {
			removed[block.ID] = true
		}
	}
	if len(removed) == 0 {
		return graph
	}
	return graph.WithoutNormalContinuationsFrom(removed)
}

func arrayProcedureAlwaysRaises(proc sourceProcedure) bool {
	if proc.Graph == nil {
		return false
	}
	graph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true, WithoutNormalErrRaiseContinuation: true})
	return !graph.IsReachable(graph.NormalExit())
}

func arrayAllocationGuardCondition(text string, guards map[string]bool, state arrayFlowState) (string, vbacfg.EdgeKind, bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "if ") {
		text = strings.TrimSpace(text[3:])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}

	functionName, argument, operator, literal, reversed, ok := parseArrayAllocationGuard(text)
	if ok {
		functionName = strings.ToLower(lastName(functionName))
		if _, ok := guards[functionName]; !ok {
			return "", "", false
		}
	} else {
		argument, operator, literal, reversed, ok = parseArrayAllocationProbeVariable(text, state)
		if !ok {
			return "", "", false
		}
	}

	if operator == "" {
		return argument, vbacfg.EdgeBranchTrue, true
	}
	value, err := strconv.Atoi(literal)
	if err != nil {
		return "", "", false
	}
	if reversed {
		switch operator {
		case ">":
			operator = "<"
		case ">=":
			operator = "<="
		case "<":
			operator = ">"
		case "<=":
			operator = ">="
		}
	}

	// The helper's proven return domain is zero for the handled failure path
	// and a positive array length after successful bounds inspection.
	switch {
	case operator == "=" && value == 0:
		return argument, vbacfg.EdgeBranchFalse, true
	case operator == "<>" && value == 0:
		return argument, vbacfg.EdgeBranchTrue, true
	case operator == "<>" && value == 1:
		// A dimension-count probe returns zero for an unallocated value and
		// one for a one-dimensional array. Its `<> 1` rejection branch is
		// therefore safe to leave only on the false edge, just like the
		// ordinary positive-length probe's zero check.
		return argument, vbacfg.EdgeBranchFalse, true
	case operator == ">" && value >= 0:
		return argument, vbacfg.EdgeBranchTrue, true
	case operator == ">=" && value >= 1:
		return argument, vbacfg.EdgeBranchTrue, true
	case operator == "<" && value <= 0:
		return argument, vbacfg.EdgeBranchFalse, true
	case operator == "<=" && value < 0:
		return argument, vbacfg.EdgeBranchFalse, true
	default:
		return "", "", false
	}
}

func parseArrayAllocationGuard(text string) (functionName, argument, operator, literal string, reversed, ok bool) {
	if match := arrayGuardCallRe.FindStringSubmatch(text); len(match) == 5 {
		return match[1], match[2], match[3], match[4], false, true
	}
	if match := arrayGuardReversedRe.FindStringSubmatch(text); len(match) == 5 {
		return match[3], match[4], match[2], match[1], true, true
	}
	return "", "", "", "", false, false
}

func parseArrayAllocationProbeVariable(text string, state arrayFlowState) (argument, operator, literal string, reversed, ok bool) {
	if match := arrayGuardValueRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[1]))]
		if !exists || value.allocationProbe == "" {
			return "", "", "", false, false
		}
		return value.allocationProbe, match[2], match[3], false, true
	}
	if match := arrayGuardValueReversedRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[3]))]
		if !exists || value.allocationProbe == "" {
			return "", "", "", false, false
		}
		return value.allocationProbe, match[2], match[1], true, true
	}
	return "", "", "", false, false
}

func parseArraySafeBoundProbeVariable(text string, state arrayFlowState) (argument, operator, literal string, reversed, ok bool) {
	if match := arrayGuardValueRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[1]))]
		if !exists || value.safeBoundProbe == "" {
			return "", "", "", false, false
		}
		return value.safeBoundProbe, match[2], match[3], false, true
	}
	if match := arrayGuardValueReversedRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[3]))]
		if !exists || value.safeBoundProbe == "" {
			return "", "", "", false, false
		}
		return value.safeBoundProbe, match[2], match[1], true, true
	}
	return "", "", "", false, false
}

func arrayAllocationProbeArgument(text string, guards map[string]bool) (string, bool) {
	functionName, argument, operator, literal, _, ok := parseArrayAllocationGuard(text)
	if !ok || operator != "" || literal != "" {
		return "", false
	}
	if _, ok := guards[strings.ToLower(lastName(functionName))]; !ok {
		return "", false
	}
	return strings.ToLower(cleanIdentifier(argument)), true
}

func arraySafeBoundProbeArgument(text string, guards map[string]bool) (string, bool) {
	functionName, argument, operator, literal, _, ok := parseArrayAllocationGuard(text)
	if !ok || operator != "" || literal != "" {
		return "", false
	}
	if _, ok := guards[strings.ToLower(lastName(functionName))]; !ok {
		return "", false
	}
	return strings.ToLower(cleanIdentifier(argument)), true
}

func mergeArrayState(states map[vbacfg.BlockID]arrayFlowState, id vbacfg.BlockID, incoming arrayFlowState) bool {
	current, exists := states[id]
	if !exists {
		states[id] = cloneArrayState(incoming)
		return true
	}
	merged := meetArrayState(current, incoming)
	if arrayStateEqual(current, merged) {
		return false
	}
	states[id] = merged
	return true
}

func meetArrayState(left, right arrayFlowState) arrayFlowState {
	out := arrayFlowState{}
	keys := map[string]bool{}
	for key := range left {
		keys[key] = true
	}
	for key := range right {
		keys[key] = true
	}
	for key := range keys {
		l, lok := left[key]
		r, rok := right[key]
		if !lok || !rok {
			out[key] = arrayValue{kind: arrayUnknown}
			continue
		}
		out[key] = meetArrayValue(l, r)
	}
	return out
}

func meetArrayValue(left, right arrayValue) arrayValue {
	out := arrayValue{
		kind:                            left.kind,
		knownArray:                      left.knownArray,
		mayBeEmpty:                      left.mayBeEmpty,
		origin:                          left.origin,
		dimensions:                      append([]arrayDimension(nil), left.dimensions...),
		preserveShape:                   append([]arrayDimension(nil), left.preserveShape...),
		allocationCountSource:           left.allocationCountSource,
		conditionalAllocationSource:     left.conditionalAllocationSource,
		returnNonEmptyArrayParameter:    left.returnNonEmptyArrayParameter,
		returnPositiveScalarParameter:   left.returnPositiveScalarParameter,
		nonEmptySource:                  left.nonEmptySource,
		returnDescriptorSourceParameter: left.returnDescriptorSourceParameter,
		returnDescriptorStartParameter:  left.returnDescriptorStartParameter,
		returnDescriptorLengthParameter: left.returnDescriptorLengthParameter,
		returnDescriptorLowerParameter:  left.returnDescriptorLowerParameter,
		boundsProof:                     left.boundsProof,
	}
	if left.kind != right.kind {
		out.kind = arrayUnknown
	}
	if left.knownArray != right.knownArray {
		out.knownArray = false
	}
	if left.conditionalAllocationSource != right.conditionalAllocationSource {
		if left.conditionalAllocationSource == "" {
			out.conditionalAllocationSource = right.conditionalAllocationSource
		} else if right.conditionalAllocationSource == "" {
			out.conditionalAllocationSource = left.conditionalAllocationSource
		} else {
			out.conditionalAllocationSource = ""
		}
	}
	if left.returnNonEmptyArrayParameter != right.returnNonEmptyArrayParameter {
		out.returnNonEmptyArrayParameter = ""
	}
	if left.returnPositiveScalarParameter != right.returnPositiveScalarParameter {
		out.returnPositiveScalarParameter = ""
	}
	if left.nonEmptySource != right.nonEmptySource {
		out.nonEmptySource = ""
	}
	if left.returnDescriptorSourceParameter != right.returnDescriptorSourceParameter ||
		left.returnDescriptorStartParameter != right.returnDescriptorStartParameter ||
		left.returnDescriptorLengthParameter != right.returnDescriptorLengthParameter ||
		left.returnDescriptorLowerParameter != right.returnDescriptorLowerParameter {
		out.returnDescriptorSourceParameter = ""
		out.returnDescriptorStartParameter = ""
		out.returnDescriptorLengthParameter = ""
		out.returnDescriptorLowerParameter = ""
	}
	out.mayBeEmpty = left.mayBeEmpty || right.mayBeEmpty
	if left.origin != right.origin {
		out.origin = arrayOriginUnknown
	}
	if !arrayDimensionsEqual(left.dimensions, right.dimensions) {
		out.dimensions = nil
	}
	out.preserveShape = meetArrayDimensions(left.preserveShape, right.preserveShape)
	if left.allocationProbe != right.allocationProbe {
		out.allocationProbe = ""
	} else {
		out.allocationProbe = left.allocationProbe
	}
	if left.safeBoundProbe != right.safeBoundProbe {
		out.safeBoundProbe = ""
	} else {
		out.safeBoundProbe = left.safeBoundProbe
	}
	if left.allocationCountSource != right.allocationCountSource {
		out.allocationCountSource = ""
	}
	if left.boundsProof != right.boundsProof {
		out.boundsProof = arrayBoundsProof{}
	}
	return out
}

func meetArrayDimensions(left, right []arrayDimension) []arrayDimension {
	if len(left) == 0 || len(left) != len(right) {
		return nil
	}
	out := make([]arrayDimension, len(left))
	for i := range left {
		out[i] = arrayDimension{
			lower: meetArrayBound(left[i].lower, right[i].lower),
			upper: meetArrayBound(left[i].upper, right[i].upper),
		}
	}
	return out
}

func meetArrayBound(left, right arrayBound) arrayBound {
	if arrayBoundsEquivalent(left, right) {
		return left
	}
	return arrayBound{}
}

func cloneArrayState(state arrayFlowState) arrayFlowState {
	out := arrayFlowState{}
	for name, value := range state {
		value.dimensions = append([]arrayDimension(nil), value.dimensions...)
		value.preserveShape = append([]arrayDimension(nil), value.preserveShape...)
		out[name] = value
	}
	return out
}

func arrayStateEqual(left, right arrayFlowState) bool {
	if len(left) != len(right) {
		return false
	}
	for key, l := range left {
		r, ok := right[key]
		if !ok || l.kind != r.kind || l.knownArray != r.knownArray || l.mayBeEmpty != r.mayBeEmpty || l.origin != r.origin || l.allocationProbe != r.allocationProbe || l.safeBoundProbe != r.safeBoundProbe || l.allocationCountSource != r.allocationCountSource || l.conditionalAllocationSource != r.conditionalAllocationSource || l.returnNonEmptyArrayParameter != r.returnNonEmptyArrayParameter || l.returnPositiveScalarParameter != r.returnPositiveScalarParameter || l.nonEmptySource != r.nonEmptySource || l.returnDescriptorSourceParameter != r.returnDescriptorSourceParameter || l.returnDescriptorStartParameter != r.returnDescriptorStartParameter || l.returnDescriptorLengthParameter != r.returnDescriptorLengthParameter || l.returnDescriptorLowerParameter != r.returnDescriptorLowerParameter || l.boundsProof != r.boundsProof || !arrayDimensionsEqual(l.dimensions, r.dimensions) || !arrayDimensionsEqual(l.preserveShape, r.preserveShape) {
			return false
		}
	}
	return true
}

func arrayDimensionsEqual(left, right []arrayDimension) bool {
	if len(left) != len(right) {
		return false
	}
	for i := range left {
		if left[i] != right[i] {
			return false
		}
	}
	return true
}

func arrayInitialState(variables map[string]arrayVariable) arrayFlowState {
	state := arrayFlowState{}
	for name, variable := range variables {
		// Only arrays and Variants can contribute to the allocation lattice.
		// Scalar/object entries remain available in `variables` for declaration
		// diagnostics, but carrying them through every CFG state is prohibitively
		// expensive for large modules.
		if !variable.isArray && !variable.isVariant {
			continue
		}
		if variable.isVariant && !variable.isArray {
			state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
			continue
		}
		if !variable.isArray {
			state[name] = arrayValue{kind: arrayUnknown, knownArray: false, origin: arrayOriginLocal}
			continue
		}
		value := arrayValue{
			knownArray:    variable.isArray,
			origin:        arrayOriginLocal,
			dimensions:    append([]arrayDimension(nil), variable.dimensions...),
			preserveShape: append([]arrayDimension(nil), variable.dimensions...),
		}
		if variable.fixed {
			value.kind = arrayAllocated
		} else if variable.parameter {
			if variable.paramArray {
				// ParamArray is materialized as an array even when the caller
				// supplies no arguments.  Its rank and bounds may be unknown,
				// but allocation itself is guaranteed by the procedure contract.
				value.kind = arrayAllocated
			} else {
				value.kind = arrayUnknown
			}
		} else {
			value.kind = arrayUnallocated
		}
		state[name] = value
	}
	return state
}

// applyArrayStaticInitializationState recognizes the narrow one-time setup
// idiom used by procedures that keep a reusable backing array in a Static
// local. The static readiness flag is passed to a resolved ByRef helper that
// sets its flag field on every normal return, and the first call then performs
// a successful direct ReDim before any array use. This lets the entry state
// carry the normal-call invariant without treating arbitrary Static arrays as
// allocated.
func applyArrayStaticInitializationState(state arrayFlowState, file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable) arrayFlowState {
	var updated arrayFlowState
	for name, variable := range variables {
		if !variable.static || !variable.isArray || variable.fixed || !arrayStaticArrayInitializationProven(file, proc, ctx, variables, name) {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayStaticArrayInitializationProven(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, targetName string) bool {
	targetName = strings.ToLower(cleanIdentifier(targetName))
	target, ok := variables[targetName]
	if !ok || !target.static || !target.isArray || target.fixed || target.parameter {
		return false
	}
	if proc.StartLine < 1 || proc.EndLine < proc.StartLine || proc.StartLine > len(file.Lines) {
		return false
	}

	type staticReadyGuard struct {
		name  string
		index int
	}
	guards := make([]staticReadyGuard, 0, 1)
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), proc.EndLine)
	for index := start + 1; index < end; index++ {
		match := arrayStaticReadyGuardRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(file.Lines[index])))
		if len(match) != 2 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		ready, declared := variables[name]
		if !declared || !ready.static || ready.parameter || ready.isArray || ready.isVariant || ready.isObject || ready.knownScalar {
			// The readiness object must be a Static UDT-like local. A scalar
			// Boolean guard or a non-static value does not carry state across
			// calls and must remain on the ordinary CFG path.
			continue
		}
		guards = append(guards, staticReadyGuard{name: name, index: index})
	}
	if len(guards) != 1 {
		return false
	}
	guard := guards[0]

	redimIndex, ok := arrayStaticInitializationBlock(file, guard.index, end, targetName)
	if !ok {
		return false
	}
	// The target must not be used before the normal-path ReDim. Otherwise the
	// entry invariant would hide a genuine first-call access before setup.
	for index := start + 1; index < redimIndex; index++ {
		if arrayStaticSourceUsesTarget(file.Lines[index], targetName, variables) {
			return false
		}
	}
	for call := range proc.Calls.All() {
		if call.Range.StartLine-1 < redimIndex && arrayCallPassesDirectArrayArgument(proc, call, targetName) {
			return false
		}
	}

	// Admit exactly one direct ReDim for this target, and reject Erase or a
	// whole-array replacement anywhere in the procedure. Indexed writes after
	// the setup are harmless and are intentionally not filtered here.
	for index := start + 1; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
			usesTarget := false
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct && strings.EqualFold(cleanIdentifier(redim.name), targetName) {
					usesTarget = true
					if strings.TrimSpace(match[1]) != "" || index != redimIndex {
						return false
					}
				}
			}
			if usesTarget && index != redimIndex {
				return false
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 && strings.EqualFold(strings.TrimSpace(match[1]), targetName) {
			return false
		}
		if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), targetName) {
			return false
		}
	}

	initializerFound := false
	for call := range proc.Calls.All() {
		line := call.Range.StartLine - 1
		if line <= guard.index || line >= redimIndex {
			continue
		}
		if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
			return false
		}
		if initializerFound {
			return false
		}
		helper, parameter, resolved := arrayStaticReadyInitializer(file, proc, call, guard.name, ctx)
		if !resolved || helper.StartByte == proc.StartByte || !arrayStaticHelperSetsReadyFlag(file, helper, parameter) {
			return false
		}
		initializerFound = true
	}
	if !initializerFound {
		return false
	}

	// Keep the proof tied to the same straight-line pre-ReDim region. The
	// post-ReDim body may contain the implementation's indexed writes and
	// loops, but a conditional or loop before allocation would leave a bypass.
	for index := guard.index + 1; index < redimIndex; index++ {
		if arrayStaticPreRedimControlFlow(file.Lines[index]) {
			return false
		}
	}
	return true
}

func arrayStaticInitializationBlock(file parsedFile, guardIndex, end int, targetName string) (redimIndex int, ok bool) {
	if guardIndex < 0 || guardIndex >= end {
		return 0, false
	}
	ifDepth := 1
	redimIndex = -1
	for index := guardIndex + 1; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		lower := strings.ToLower(text)
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") {
			continue
		}
		if lower == "end if" {
			if ifDepth == 1 {
				return redimIndex, redimIndex >= 0
			}
			ifDepth--
			continue
		}
		if ifDepth == 1 && (lower == "else" || strings.HasPrefix(lower, "elseif ")) {
			return 0, false
		}
		if arrayStaticBlockIfStart(text) {
			if ifDepth == 1 && redimIndex < 0 {
				return 0, false
			}
			ifDepth++
			continue
		}
		if ifDepth == 1 && redimIndex < 0 && arrayStaticPreRedimControlFlow(file.Lines[index]) {
			return 0, false
		}
		if ifDepth != 1 {
			continue
		}
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 {
			continue
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), targetName) {
				if strings.TrimSpace(match[1]) != "" || redimIndex >= 0 {
					return 0, false
				}
				redimIndex = index
			}
		}
	}
	return 0, false
}

func arrayStaticReadyInitializer(file parsedFile, caller sourceProcedure, call procedureir.CallSite, readyName string, ctx analysisContext) (sourceProcedure, string, bool) {
	resolution := arrayCallResolution(ctx, call)
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return sourceProcedure{}, "", false
	}
	want := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
	var helper sourceProcedure
	found := false
	for _, candidate := range file.Procedures {
		if arrayProcedureKey(candidate) != want {
			continue
		}
		if found {
			return sourceProcedure{}, "", false
		}
		helper = candidate
		found = true
	}
	if !found {
		return sourceProcedure{}, "", false
	}
	bindings, ok := arrayCallArgumentBindings(caller, helper, call)
	if !ok {
		return sourceProcedure{}, "", false
	}
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= helper.Params.Len() || directArrayArgumentName(binding.text) != strings.ToLower(cleanIdentifier(readyName)) {
			continue
		}
		parameter := helper.Params.valueAt(binding.parameterIndex)
		if parameterIsByRefScalar(parameter) && !arrayKnownScalarType(parameter.Type) && !isObjectType(parameter.Type) {
			return helper, helper.Params.valueAt(binding.parameterIndex).Name, true
		}
	}
	return sourceProcedure{}, "", false
}

func arrayStaticHelperSetsReadyFlag(file parsedFile, helper sourceProcedure, parameter string) bool {
	want := canonicalArrayBoundExpression(parameter + ".isSet")
	count := 0
	setLine := -1
	lastExecutable := -1
	depth := 0
	start := max(0, helper.StartLine-1)
	end := min(len(file.Lines), helper.EndLine)
	for index := start; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		lower := strings.ToLower(text)
		if arrayStaticExecutableSourceLine(text) {
			lastExecutable = index
		}
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") {
			continue
		}
		if arrayStaticBlockEnd(text) {
			if depth == 0 {
				return false
			}
			depth--
			continue
		}
		if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed && canonicalArrayBoundExpression(lhs) == want {
			if depth != 0 || !strings.EqualFold(strings.TrimSpace(rhs), "true") {
				return false
			}
			count++
			setLine = index
		}
		if arrayStaticBlockStart(text) {
			depth++
		}
		if strings.HasPrefix(lower, "on error resume next") {
			return false
		}
	}
	return depth == 0 && count == 1 && setLine == lastExecutable && setLine >= start
}

func arrayStaticSourceUsesTarget(text, targetName string, variables map[string]arrayVariable) bool {
	for _, use := range arrayIndexedUsesForSource(text, variables) {
		if strings.EqualFold(cleanIdentifier(use.name), targetName) && len(use.args) > 0 {
			return true
		}
	}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		if len(bound) > 2 && strings.EqualFold(cleanIdentifier(bound[2]), targetName) {
			return true
		}
	}
	if match := arrayForEachRe.FindStringSubmatch(strings.TrimSpace(text)); len(match) == 2 && strings.EqualFold(cleanIdentifier(match[1]), targetName) {
		return true
	}
	return false
}

func arrayStaticExecutableSourceLine(text string) bool {
	text = strings.TrimSpace(normalizedCodeLine(text))
	if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") || isProcedureHeaderLine(strings.ToLower(text)) {
		return false
	}
	switch strings.ToLower(text) {
	case "end sub", "end function", "end property", "else":
		return false
	default:
		return !arrayStaticBlockEnd(text)
	}
}

func arrayStaticBlockIfStart(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return (strings.HasPrefix(lower, "if ") || strings.HasPrefix(lower, "elseif ")) && strings.HasSuffix(lower, " then")
}

func arrayStaticBlockStart(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if arrayStaticBlockIfStart(lower) {
		return true
	}
	for _, prefix := range []string{"for ", "do", "while ", "select ", "with "} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}

func arrayStaticBlockEnd(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{"end if", "end with", "end select", "next", "loop", "wend"} {
		if lower == prefix || strings.HasPrefix(lower, prefix+" ") {
			return true
		}
	}
	return false
}

func arrayStaticPreRedimControlFlow(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(text)))
	if lower == "" || strings.HasPrefix(lower, "'") || strings.HasPrefix(lower, "#") {
		return false
	}
	if arrayStaticBlockIfStart(lower) {
		return true
	}
	for _, prefix := range []string{"for ", "do", "while ", "select ", "with ", "else", "goto ", "on error ", "exit "} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}

func applyArrayByRefEntryStates(state arrayFlowState, proc sourceProcedure, variables map[string]arrayVariable, entries map[string]map[int]bool, conditions map[string]map[int]string) arrayFlowState {
	parameters := entries[arrayProcedureKey(proc)]
	conditionalParameters := conditions[arrayProcedureKey(proc)]
	if len(parameters) == 0 && len(conditionalParameters) == 0 {
		return state
	}
	updated := cloneArrayState(state)
	for index, allocated := range parameters {
		if !allocated || index < 0 || index >= proc.Params.Len() {
			continue
		}
		name := strings.ToLower(proc.Params.valueAt(index).Name)
		variable, known := variables[name]
		value, exists := updated[name]
		if !known || !exists || !variable.isArray {
			continue
		}
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	for index, source := range conditionalParameters {
		if source == "" || index < 0 || index >= proc.Params.Len() {
			continue
		}
		name := strings.ToLower(proc.Params.valueAt(index).Name)
		variable, known := variables[name]
		value, exists := updated[name]
		if !known || !exists || !variable.isArray || value.kind == arrayAllocated && value.knownArray {
			continue
		}
		if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, source) {
			continue
		}
		value.allocationCountSource = source
		updated[name] = value
	}
	return updated
}

func arrayProcedureKey(proc sourceProcedure) string {
	module := strings.TrimSpace(proc.Module)
	name := strings.TrimSpace(proc.Name)
	if module == "" {
		return strings.ToLower(name)
	}
	if name == "" {
		return strings.ToLower(module)
	}
	return strings.ToLower(module + "." + name)
}

func arrayParticipantProcedureIdentity(proc sourceProcedure) string {
	path := arrayProcedureSourcePath(proc)
	if path == "" {
		path = strings.ToLower(strings.TrimSpace(proc.Module))
	}
	return strings.Join([]string{
		path,
		strconv.Itoa(proc.Index),
		strconv.Itoa(proc.StartByte),
		strconv.Itoa(proc.StartLine),
		strings.ToLower(strings.TrimSpace(proc.Name)),
		strings.ToLower(string(proc.ProcedureKind)),
	}, "\x00")
}

func arrayParticipantDisambiguatedKey(proc sourceProcedure) string {
	return arrayProcedureKey(proc) + "|" + arrayParticipantProcedureIdentity(proc)
}

func arrayParticipantLookupKey(proc sourceProcedure, participantKeys map[string]string) string {
	if len(participantKeys) > 0 {
		if key := participantKeys[arrayParticipantProcedureIdentity(proc)]; key != "" {
			return key
		}
	}
	return arrayProcedureKey(proc)
}

func arrayProcedureSourcePath(proc sourceProcedure) string {
	if proc.Document == nil {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(proc.Document.Path))
}

func arrayProcedureLess(left, right sourceProcedure) bool {
	leftPath, rightPath := arrayProcedureSourcePath(left), arrayProcedureSourcePath(right)
	if leftPath != rightPath {
		return leftPath < rightPath
	}
	if left.StartByte != right.StartByte {
		return left.StartByte < right.StartByte
	}
	if left.StartLine != right.StartLine {
		return left.StartLine < right.StartLine
	}
	if !strings.EqualFold(left.Module, right.Module) {
		return strings.ToLower(left.Module) < strings.ToLower(right.Module)
	}
	if !strings.EqualFold(left.Name, right.Name) {
		return strings.ToLower(left.Name) < strings.ToLower(right.Name)
	}
	if left.ProcedureKind != right.ProcedureKind {
		return string(left.ProcedureKind) < string(right.ProcedureKind)
	}
	return arrayProcedureKey(left) < arrayProcedureKey(right)
}

type arrayByRefAllocationSummaries map[string]map[int]bool

// arrayByRefConditionalAllocations records a ByRef array output that is
// allocated only when a count-bearing input is positive. The outer key is the
// callee procedure; each entry maps the output array parameter index to the
// count-bearing input parameter index.
type arrayByRefConditionalAllocations map[string]map[int]int

// arrayByRefLengthAllocations records a ByRef array output whose paired
// ByRef scalar output is assigned a successful array length. A positive value
// of that scalar is therefore a conditional allocation proof for the array.
type arrayByRefLengthAllocations map[string]map[int]int

var (
	arrayByRefCountExitRe   = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*\.\s*count\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayByRefCountRedimRe  = regexp.MustCompile(`(?i)^\s*redim\s+([A-Za-z_]\w*)\s*\(\s*0\s+to\s+([A-Za-z_]\w*)\s*\.\s*count\s*-\s*1\s*\)\s*$`)
	arrayByRefLengthFullRe  = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*-\s*lbound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayByRefLengthUpperRe = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
)

type arrayModuleAllocationSummaries map[string]map[string]bool

// arrayModuleInvalidationSummaries records module arrays that may be
// unallocated or unknown when a project-local procedure returns normally.
// The summary starts from an allocated module-array state, so a fixed-size
// array and an Erase followed by a guaranteed ReDim remain allocated while a
// reachable conditional Erase is retained as an invalidation.
type arrayModuleInvalidationSummaries map[string]map[string]bool

func inferArrayModuleInvalidationSummaries(files []parsedFile, ctx analysisContext) arrayModuleInvalidationSummaries {
	summaries := arrayModuleInvalidationSummaries{}
	ctx.arrayModuleInvalidations = summaries
	ctx.arrayModuleInvalidationCacheWritable = true
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		moduleDecls := file.moduleDecls()
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			key := arrayProcedureKey(proc)
			if key == "" {
				continue
			}
			if _, cached := summaries[key]; cached {
				continue
			}
			summaries[key] = arrayPrivateModuleArrayInvalidationsWithVisiting(file, proc, moduleDecls, ctx, map[string]bool{})
		}
	}
	return summaries
}

type arrayProcedureDominators map[string]map[vbacfg.BlockID]bool

type arrayModuleConfigurationState struct {
	byProcedure       map[string]map[string]bool
	dataTable         map[string]bool
	genericCollection map[string]bool
}

// arrayModuleEntryStates records module-level arrays that are allocated at
// every known entry into a project-local helper. A private helper is analyzed
// independently from its callers, so without this summary an allocation made
// by a public entry procedure is lost as soon as the call crosses a procedure
// boundary.
type arrayModuleEntryStates map[string]map[string]bool

// arrayModuleReadyGuardStates records the stronger, source-owned invariant
// behind a module Boolean readiness guard. The implication is intentionally
// narrow: the guard has one source-owned True write, that write is reached
// only after the module array is allocated on every path, and direct array
// invalidation is paired with a dominating False write. This lets a public
// consumer prove its module array without trusting arbitrary caller state.
type arrayModuleReadyGuardStates map[string]map[string]map[string]bool

func arrayPrivateProcedureTargets(files []parsedFile) map[string]sourceProcedure {
	targets := map[string]sourceProcedure{}
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		facts := file.ModuleFacts
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			visibility := strings.TrimSpace(proc.Visibility)
			private := strings.EqualFold(visibility, "Private") || strings.EqualFold(visibility, "Friend")
			modulePrivate := strings.EqualFold(visibility, "Public") && facts.privateModulePresent()
			if !private && !modulePrivate {
				continue
			}
			targets[arrayProcedureKey(proc)] = proc
		}
	}
	return targets
}

type arrayParticipantGraph struct {
	all              map[string]sourceProcedure
	fileByKey        map[string]parsedFile
	byModule         map[string][]string
	keyByIdentity    map[string]string
	candidateIndex   arrayCandidateIndex
	adjacency        map[string]map[string]bool
	reverse          map[string]map[string]bool
	resolvedReverse  map[string]map[string]bool
	callAdjacency    map[string]map[string]bool
	knownSeeds       map[string]bool
	intrinsicSeeds   map[string]bool
	uncertainFacts   map[bool]map[string]bool
	uncertainCalls   map[string]bool
	moduleArrayUsers map[string][]string
}

// buildArrayParticipantGraph classifies procedures and resolves all
// project-local call edges once. The two participant boundaries (the local
// fail-open plan and the narrower fixed-point plan) share this immutable graph
// so a revision does not scan every procedure and call site twice.
func buildArrayParticipantGraph(files []parsedFile, ctx analysisContext) *arrayParticipantGraph {
	all := make(map[string]sourceProcedure)
	fileByKey := make(map[string]parsedFile)
	byModule := make(map[string][]string)
	keyByIdentity := make(map[string]string)
	type entry struct {
		file parsedFile
		proc sourceProcedure
		base string
	}
	entries := make([]entry, 0)
	baseCounts := make(map[string]int)
	for _, file := range files {
		procedures := file.procedureView()
		for index := 0; index < procedures.Len(); index++ {
			proc := procedures.valueAt(index)
			base := arrayProcedureKey(proc)
			if base == "" {
				continue
			}
			entries = append(entries, entry{file: file, proc: proc, base: base})
			baseCounts[base]++
		}
	}
	for _, item := range entries {
		key := item.base
		if baseCounts[item.base] > 1 {
			key = arrayParticipantDisambiguatedKey(item.proc)
		}
		if _, exists := all[key]; exists {
			// Synthetic focused projections may omit Document/Index and still
			// produce identical identity fields. Keep their source-order ordinal
			// as a deterministic final discriminator.
			key = key + "|" + strconv.Itoa(len(all))
		}
		all[key] = item.proc
		fileByKey[key] = item.file
		identity := arrayParticipantProcedureIdentity(item.proc)
		if _, exists := keyByIdentity[identity]; exists {
			identity = identity + "\x00" + strconv.Itoa(len(all))
		}
		keyByIdentity[identity] = key
		module := strings.ToLower(strings.TrimSpace(item.proc.Module))
		byModule[module] = append(byModule[module], key)
	}
	if len(all) == 0 {
		return &arrayParticipantGraph{
			all:              all,
			fileByKey:        fileByKey,
			byModule:         byModule,
			keyByIdentity:    keyByIdentity,
			candidateIndex:   buildArrayCandidateIndex(all),
			adjacency:        map[string]map[string]bool{},
			reverse:          map[string]map[string]bool{},
			resolvedReverse:  map[string]map[string]bool{},
			callAdjacency:    map[string]map[string]bool{},
			knownSeeds:       map[string]bool{},
			intrinsicSeeds:   map[string]bool{},
			uncertainFacts:   map[bool]map[string]bool{false: {}, true: {}},
			uncertainCalls:   map[string]bool{},
			moduleArrayUsers: map[string][]string{},
		}
	}
	candidateIndex := buildArrayCandidateIndex(all)

	adjacency := make(map[string]map[string]bool, len(all))
	knownSeeds := make(map[string]bool, len(all))
	intrinsicSeeds := make(map[string]bool, len(all))
	uncertainFacts := map[bool]map[string]bool{false: {}, true: {}}
	uncertainCalls := make(map[string]bool)
	moduleArrayUsers := make(map[string][]string)
	callAdjacency := make(map[string]map[string]bool, len(all))
	type resolvedEdge struct {
		caller string
		target string
	}
	resolvedEdges := make([]resolvedEdge, 0)
	for key, proc := range all {
		file := fileByKey[key]
		moduleDecls := moduleDeclarationsForProcedure(files, proc)
		arraySeed := procedureArraySeed(proc)
		moduleArrayUse := procedureUsesModuleArray(file, proc, moduleDecls)
		shapeSeed := procedureHasArrayParameter(proc) || procedureReturnsArray(proc) || moduleArrayUse
		arraySeed = arraySeed || procedureHasArrayForEach(proc) || procedureHasObjectComparison(proc)
		if arraySeed || shapeSeed {
			knownSeeds[key] = true
			intrinsicSeeds[key] = true
		}
		if moduleArrayUse {
			module := strings.ToLower(strings.TrimSpace(proc.Module))
			moduleArrayUsers[module] = append(moduleArrayUsers[module], key)
		}
		for _, ignoreFeatureUnknown := range []bool{false, true} {
			if arrayParticipantFactsUncertain(proc, ignoreFeatureUnknown) {
				// Recovered statements, missing CFGs, and conditional IR facts are
				// bounded uncertainty. Keep the known module-array cluster fail-open
				// until resolution can prove a smaller dependency boundary.
				uncertainFacts[ignoreFeatureUnknown][key] = true
			}
		}
		for call := range proc.Calls.All() {
			resolution := arrayCallResolution(ctx, call)
			addEdge := func(target string) {
				if target == "" || target == key {
					if target == key {
						adjacency[key] = ensureArrayKeySet(adjacency[key])
						adjacency[key][target] = true
					}
					return
				}
				adjacency[key] = ensureArrayKeySet(adjacency[key])
				adjacency[key][target] = true
			}
			addCallEdge := func(target string) {
				if target == "" {
					return
				}
				callAdjacency[key] = ensureArrayKeySet(callAdjacency[key])
				callAdjacency[key][target] = true
			}
			if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
				candidate := resolution.Candidates[0]
				target := arrayCandidateKey(candidate, all, candidateIndex)
				addCallEdge(target)
				// Defer resolved-edge filtering until every procedure's
				// intrinsic seed has been classified. The source map is not
				// ordered, so checking the target while this loop runs would
				// make participant membership depend on map iteration order.
				resolvedEdges = append(resolvedEdges, resolvedEdge{caller: key, target: target})
				continue
			}
			if resolution.Status == procedureir.ResolutionAmbiguous || resolution.Status == procedureir.ResolutionUnresolved || resolution.Status == procedureir.ResolutionDynamic || resolution.Status == procedureir.ResolutionIncomplete {
				for _, candidate := range resolution.Candidates {
					target := arrayCandidateKey(candidate, all, candidateIndex)
					addCallEdge(target)
					addEdge(target)
				}
				if len(resolution.Candidates) == 0 {
					// A candidate-bearing ambiguous/dynamic/unresolved call is
					// bounded by those project-local candidates. Only a boundary
					// with no target identity expands to its owning module.
					uncertainCalls[key] = true
				}
			}
		}
	}
	for _, edge := range resolvedEdges {
		if intrinsicSeeds[edge.target] {
			adjacency[edge.caller] = ensureArrayKeySet(adjacency[edge.caller])
			adjacency[edge.caller][edge.target] = true
		}
	}
	// Keep reverse links for every project-local resolved edge separately from
	// the direct semantic graph. A bounded caller extension below lets a chain
	// such as Top -> Wrapper -> ArrayWorker reach the seed transitively when
	// facts are complete; recovered/incomplete targets use the module-array
	// fallback below instead of opening an unbounded caller hub.
	reverse := make(map[string]map[string]bool, len(adjacency))
	resolvedReverse := make(map[string]map[string]bool, len(adjacency))
	for _, edge := range resolvedEdges {
		if edge.target == "" {
			continue
		}
		resolvedReverse[edge.target] = ensureArrayKeySet(resolvedReverse[edge.target])
		resolvedReverse[edge.target][edge.caller] = true
	}
	for caller, callees := range adjacency {
		for callee := range callees {
			reverse[callee] = ensureArrayKeySet(reverse[callee])
			reverse[callee][caller] = true
		}
	}
	return &arrayParticipantGraph{
		all:              all,
		fileByKey:        fileByKey,
		byModule:         byModule,
		keyByIdentity:    keyByIdentity,
		candidateIndex:   candidateIndex,
		adjacency:        adjacency,
		reverse:          reverse,
		resolvedReverse:  resolvedReverse,
		callAdjacency:    callAdjacency,
		knownSeeds:       knownSeeds,
		intrinsicSeeds:   intrinsicSeeds,
		uncertainFacts:   uncertainFacts,
		uncertainCalls:   uncertainCalls,
		moduleArrayUsers: moduleArrayUsers,
	}
}

func (graph *arrayParticipantGraph) participantSet(ignoreFeatureUnknown bool) map[string]bool {
	participants := make(map[string]bool, len(graph.all))
	for key := range graph.knownSeeds {
		participants[key] = true
	}
	for key := range graph.uncertainFacts[ignoreFeatureUnknown] {
		participants[key] = true
	}
	closeArrayParticipantClosure(participants, graph.adjacency, graph.reverse)
	addResolvedCallerBoundary(participants, graph.resolvedReverse, graph.all, graph.byModule, ignoreFeatureUnknown)
	closeArrayParticipantClosure(participants, graph.adjacency, graph.reverse)
	globalFallback := false
	for key := range graph.uncertainFacts[ignoreFeatureUnknown] {
		if !participants[key] {
			continue
		}
		module := strings.ToLower(strings.TrimSpace(graph.all[key].Module))
		if module == "" {
			globalFallback = true
			continue
		}
		// Recovered or incomplete facts are expanded to the smallest known
		// module-array cluster. This preserves conservative state propagation
		// without turning every procedure in a giant module into a candidate.
		for _, user := range graph.moduleArrayUsers[module] {
			participants[user] = true
		}
	}
	for key := range graph.uncertainCalls {
		if !participants[key] {
			continue
		}
		// An unknown-only procedure remains a local participant for fail-open
		// diagnostics, but it must not open an entire giant module before a
		// semantic array seed reaches the uncertainty boundary.
		if !graph.intrinsicSeeds[key] && !ignoreFeatureUnknown {
			continue
		}
		module := strings.ToLower(strings.TrimSpace(graph.all[key].Module))
		if module == "" {
			globalFallback = true
			continue
		}
		if users := graph.moduleArrayUsers[module]; len(users) > 0 {
			for _, user := range users {
				participants[user] = true
			}
			continue
		}
		for _, procedure := range graph.byModule[module] {
			participants[procedure] = true
		}
	}
	if globalFallback {
		for key := range graph.all {
			participants[key] = true
		}
	}
	closeArrayParticipantClosure(participants, graph.adjacency, graph.reverse)
	return participants
}

// buildArrayParticipantSet derives the bounded interprocedural closure used
// by the array capability. Module declarations are deliberately not seeds:
// only a procedure that observes an array locally, exposes an array-shaped
// parameter/return, or reaches an array through a resolved call participates.
func buildArrayParticipantSet(files []parsedFile, ctx analysisContext) map[string]bool {
	return buildArrayParticipantGraph(files, ctx).participantSet(ctx.arrayIgnoreFeatureUnknown)
}

func buildArrayParticipantSets(files []parsedFile, ctx analysisContext) (map[string]bool, map[string]bool, map[string]string) {
	graph := buildArrayParticipantGraph(files, ctx)
	participants := graph.participantSet(ctx.arrayIgnoreFeatureUnknown)
	return participants, buildArrayInterproceduralParticipantSetFromGraph(graph, participants), graph.keyByIdentity
}

// buildArrayInterproceduralParticipantSet keeps the local fail-open plan
// separate from the fixed-point scope. A complete procedure that has no
// semantic array seed but carries an unknown array capability must retain its
// local array kernel/projection; it must not, by itself, make every array
// summary and module-entry solver walk the surrounding module.
func buildArrayInterproceduralParticipantSet(files []parsedFile, ctx analysisContext, participants map[string]bool) map[string]bool {
	graph := buildArrayParticipantGraph(files, ctx)
	return buildArrayInterproceduralParticipantSetFromGraph(graph, participants)
}

func buildArrayInterproceduralParticipantSetFromGraph(graph *arrayParticipantGraph, participants map[string]bool) map[string]bool {
	if len(participants) == 0 {
		return map[string]bool{}
	}
	// Derive the fixed-point boundary with feature-unknown bits ignored from the
	// shared graph. This preserves the proven semantic closure used before an
	// unknown-only local participant was added, while the outer participant plan
	// still retains that procedure for local fail-open diagnostics.
	legacy := graph.participantSet(true)
	legacyResult := make(map[string]bool, len(legacy))
	for key := range legacy {
		if participants[key] {
			legacyResult[key] = true
		}
	}
	all := graph.all
	fileByKey := graph.fileByKey
	moduleSizes := make(map[string]int)
	for _, proc := range all {
		moduleSizes[strings.ToLower(strings.TrimSpace(proc.Module))]++
	}
	connected := make(map[string]bool)
	for key := range graph.callAdjacency {
		if !participants[key] {
			continue
		}
		for target := range graph.callAdjacency[key] {
			if !participants[target] {
				continue
			}
			if legacy[key] {
				connected[target] = true
			}
			if legacy[target] {
				connected[key] = true
			}
		}
	}
	for key := range connected {
		if legacyResult[key] {
			continue
		}
		proc := graph.all[key]
		if !procedureArrayFactsUncertain(proc) || !procedureHasCompleteArrayFacts(proc) || proc.Features.unknown&arrayParticipantUnknownFeatures == 0 {
			continue
		}
		module := strings.ToLower(strings.TrimSpace(proc.Module))
		if moduleSizes[module] > arrayResolvedCallerModuleLimit {
			if !procedureHasDirectModuleArrayOperation(fileByKey[key], proc, fileByKey[key].moduleDecls()) {
				continue
			}
		}
		legacyResult[key] = true
	}
	return legacyResult
}

const arrayResolvedCallerModuleLimit = 512

func addResolvedCallerBoundary(participants map[string]bool, reverse map[string]map[string]bool, all map[string]sourceProcedure, byModule map[string][]string, ignoreFeatureUnknown bool) {
	keys := make([]string, 0, len(participants))
	for key := range participants {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, target := range keys {
		procedure, ok := all[target]
		if !ok || arrayParticipantFactsUncertain(procedure, ignoreFeatureUnknown) || len(byModule[strings.ToLower(strings.TrimSpace(procedure.Module))]) > arrayResolvedCallerModuleLimit {
			continue
		}
		for caller := range reverse[target] {
			participants[caller] = true
		}
	}
}

const arrayParticipantUnknownFeatures = featureArray | featureRangeArray | featureObject

func procedureArraySeed(proc sourceProcedure) bool {
	if proc.Features.present&(featureArray|featureRangeArray) != 0 {
		return true
	}
	// Resolved scalar calls without array-shaped evidence remain excluded. Any
	// unresolved/external call that leaves an array capability unknown is a seed
	// and is bounded by the uncertainty policy below.
	return false
}

func procedureArrayFactsUncertain(proc sourceProcedure) bool {
	if proc.Features.unknown&arrayParticipantUnknownFeatures != 0 {
		return true
	}
	if proc.IR == nil {
		return proc.Features.unknown != 0
	}
	return !procedureHasCompleteArrayFacts(proc)
}

func arrayParticipantFactsUncertain(proc sourceProcedure, ignoreFeatureUnknown bool) bool {
	if !ignoreFeatureUnknown {
		return procedureArrayFactsUncertain(proc)
	}
	if proc.IR == nil {
		return proc.Features.unknown != 0
	}
	return !procedureHasCompleteArrayFacts(proc)
}

// procedureHasCompleteArrayFacts reports whether the procedure's structural
// IR facts are complete. Feature unknowns are intentionally checked by the
// caller so complete IR with an unknown array capability can remain a local
// fail-open participant without widening the interprocedural scope.
func procedureHasCompleteArrayFacts(proc sourceProcedure) bool {
	if proc.IR == nil {
		return false
	}
	if proc.Document == nil || proc.Document.Parse.HasError || proc.Document.Parse.HasMissing || proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 || proc.Graph == nil {
		return false
	}
	if len(proc.Graph.UnknownFlowSources) > 0 {
		return false
	}
	for declaration := range proc.Declarations.All() {
		if declaration.Recovered || len(declaration.ConditionalBranches) > 0 {
			return false
		}
	}
	for statement := range proc.Statements.All() {
		if statement.Recovered || statement.Kind == procedureir.StatementUnknown || statement.Kind == procedureir.StatementRecovered || len(statement.ConditionalBranches) > 0 {
			return false
		}
	}
	for expression := range proc.Expressions.All() {
		if expression.Recovered || expression.Kind == procedureir.ExpressionUnknown && !isKnownNonValueExpressionSyntax(expression.SyntaxKind) {
			return false
		}
	}
	return true
}

func closeArrayParticipantClosure(participants map[string]bool, adjacency, reverse map[string]map[string]bool) {
	keys := make([]string, 0, len(participants))
	for key := range participants {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	queue := append([]string(nil), keys...)
	seen := make(map[string]bool, len(keys))
	for _, key := range keys {
		seen[key] = true
	}
	for head := 0; head < len(queue); head++ {
		current := queue[head]
		neighbors := make([]string, 0, len(adjacency[current])+len(reverse[current]))
		for caller := range reverse[current] {
			if !seen[caller] {
				neighbors = append(neighbors, caller)
			}
		}
		for callee := range adjacency[current] {
			if !seen[callee] {
				neighbors = append(neighbors, callee)
			}
		}
		sort.Strings(neighbors)
		for _, neighbor := range neighbors {
			if seen[neighbor] {
				continue
			}
			seen[neighbor] = true
			participants[neighbor] = true
			queue = append(queue, neighbor)
		}
	}
}

func moduleDeclarationsForProcedure(files []parsedFile, proc sourceProcedure) map[string]sourceDeclaration {
	for _, file := range files {
		if strings.EqualFold(strings.TrimSpace(file.Module), strings.TrimSpace(proc.Module)) || strings.EqualFold(strings.TrimSpace(file.IR.ModuleName), strings.TrimSpace(proc.Module)) {
			return file.moduleDecls()
		}
	}
	return nil
}

func procedureHasArrayParameter(proc sourceProcedure) bool {
	for parameter := range proc.Params.All() {
		if parameterIsArray(parameter) {
			return true
		}
	}
	return false
}

func procedureReturnsArray(proc sourceProcedure) bool {
	return proc.ReturnValueShape == procedureir.ValueShapeFixedArray ||
		proc.ReturnValueShape == procedureir.ValueShapeDynamicArray ||
		strings.Contains(proc.ReturnType, "()")
}

func procedureUsesModuleArray(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) bool {
	if len(moduleDecls) == 0 {
		return false
	}
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeModule {
			continue
		}
		for name, declaration := range moduleDecls {
			if strings.EqualFold(strings.TrimSpace(name), strings.TrimSpace(access.Name)) && declaration.Array && !declaration.Parameter {
				return true
			}
		}
	}
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct {
					if declaration, ok := moduleDecls[strings.ToLower(cleanIdentifier(redim.name))]; ok && declaration.Array && !declaration.Parameter {
						return true
					}
				}
			}
		}
		if lhs, _, indexed, ok := arrayAssignment(text); ok && !indexed {
			if declaration, declared := moduleDecls[strings.ToLower(cleanIdentifier(lhs))]; declared && declaration.Array && !declaration.Parameter {
				return true
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
			if declaration, declared := moduleDecls[strings.ToLower(cleanIdentifier(match[1]))]; declared && declaration.Array && !declaration.Parameter {
				return true
			}
		}
	}
	if proc.StartLine < 1 || proc.EndLine < proc.StartLine || len(file.Lines) == 0 {
		return false
	}
	if facts := file.moduleAnalysisFacts(); facts != nil {
		for name, declaration := range moduleDecls {
			if !declaration.Array || declaration.Parameter {
				continue
			}
			used := false
			facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
				if operation.Line >= proc.StartLine && operation.Line <= proc.EndLine {
					used = true
				}
			})
			if used {
				return true
			}
		}
	}
	return false
}

func procedureHasDirectModuleArrayOperation(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) bool {
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		for name, declaration := range moduleDecls {
			if declaration.Array && !declaration.Parameter && moduleArrayIndexedIdentifier(text, name) {
				return true
			}
		}
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct {
					if declaration, ok := moduleDecls[strings.ToLower(cleanIdentifier(redim.name))]; ok && declaration.Array && !declaration.Parameter {
						return true
					}
				}
			}
		}
	}
	return false
}

func moduleArrayIndexedIdentifier(text, name string) bool {
	text = strings.ToLower(text)
	name = strings.ToLower(strings.TrimSpace(name))
	if name == "" {
		return false
	}
	for start := 0; start <= len(text)-len(name); {
		relative := strings.Index(text[start:], name)
		if relative < 0 {
			return false
		}
		index := start + relative
		end := index + len(name)
		if (index == 0 || !isIdentifierPart(text[index-1])) && (end == len(text) || !isIdentifierPart(text[end])) {
			for end < len(text) && (text[end] == ' ' || text[end] == '\t') {
				end++
			}
			if end < len(text) && text[end] == '(' {
				return true
			}
		}
		start = index + len(name)
	}
	return false
}

func procedureHasArrayForEach(proc sourceProcedure) bool {
	for statement := range proc.Statements.All() {
		if statement.Kind == procedureir.StatementForEach {
			return true
		}
	}
	return false
}

func procedureHasObjectComparison(proc sourceProcedure) bool {
	if proc.Features.present&featureObject == 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		text := strings.ToLower(strings.TrimSpace(statement.Text))
		if strings.Contains(text, "nothing") && strings.Contains(text, "=") && !strings.Contains(text, " is ") {
			return true
		}
	}
	return false
}

func ensureArrayKeySet(set map[string]bool) map[string]bool {
	if set == nil {
		return map[string]bool{}
	}
	return set
}

type arrayCandidateLineKey struct {
	line int
	kind string
}

type arrayCandidateQualifiedKindKey struct {
	qualified string
	kind      string
}

type arrayCandidateIndex struct {
	byName          map[string]string
	byQualified     map[string]string
	byQualifiedKind map[arrayCandidateQualifiedKindKey]string
	byLineAndKind   map[arrayCandidateLineKey]string
}

// buildArrayCandidateIndex preserves the old sorted-key tie breaking while
// avoiding a project-wide key collection and sort for every uncertain call.
func buildArrayCandidateIndex(all map[string]sourceProcedure) arrayCandidateIndex {
	index := arrayCandidateIndex{
		byName:          make(map[string]string, len(all)),
		byQualified:     make(map[string]string, len(all)),
		byQualifiedKind: make(map[arrayCandidateQualifiedKindKey]string, len(all)),
		byLineAndKind:   make(map[arrayCandidateLineKey]string),
	}
	keys := make([]string, 0, len(all))
	for key := range all {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, key := range keys {
		proc := all[key]
		name := strings.ToLower(strings.TrimSpace(proc.Name))
		if name == "" {
			continue
		}
		if _, exists := index.byName[name]; !exists {
			index.byName[name] = key
		}
		qualified := strings.ToLower(strings.TrimSpace(proc.Module + "." + proc.Name))
		if qualified != "." {
			if existing, exists := index.byQualified[qualified]; !exists {
				index.byQualified[qualified] = key
			} else if existing != key {
				index.byQualified[qualified] = ""
			}
			qualifiedKind := arrayCandidateQualifiedKindKey{qualified: qualified, kind: strings.ToLower(string(proc.ProcedureKind))}
			if existing, exists := index.byQualifiedKind[qualifiedKind]; !exists {
				index.byQualifiedKind[qualifiedKind] = key
			} else if existing != key {
				index.byQualifiedKind[qualifiedKind] = ""
			}
		}
		if proc.StartLine > 0 {
			lineKey := arrayCandidateLineKey{line: proc.StartLine, kind: strings.ToLower(string(proc.ProcedureKind))}
			if _, exists := index.byLineAndKind[lineKey]; !exists {
				index.byLineAndKind[lineKey] = key
			}
		}
	}
	return index
}

func arrayCandidateKey(candidate procedureir.Candidate, all map[string]sourceProcedure, index arrayCandidateIndex) string {
	qualified := strings.ToLower(strings.TrimSpace(candidate.QualifiedName))
	if proc, ok := all[qualified]; ok {
		return arrayProcedureKey(proc)
	}
	qualifiedKind := arrayCandidateQualifiedKindKey{qualified: qualified, kind: strings.ToLower(strings.TrimSpace(candidate.Kind))}
	if key := index.byQualifiedKind[qualifiedKind]; key != "" {
		return key
	}
	if key := index.byQualified[qualified]; key != "" {
		return key
	}
	if key, ok := index.byName[qualified]; ok {
		return key
	}
	if candidate.Line > 0 {
		lineKey := arrayCandidateLineKey{line: candidate.Line, kind: strings.ToLower(candidate.Kind)}
		if key, ok := index.byLineAndKind[lineKey]; ok {
			return key
		}
	}
	return ""
}

func arrayProcedureIsParticipant(ctx analysisContext, proc sourceProcedure) bool {
	participants := ctx.arrayInterproceduralParticipants
	if participants == nil {
		participants = ctx.arrayParticipants
	}
	if participants == nil {
		return true
	}
	return participants[arrayParticipantLookupKey(proc, ctx.arrayParticipantKeys)]
}

func inferArrayByRefAllocationSummaries(files []parsedFile, ctx analysisContext, targets map[string]sourceProcedure) arrayByRefAllocationSummaries {
	summaries := arrayByRefAllocationSummaries{}
	procedures := make([]struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
	}, 0)
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !procedureHasByRefArrayParameter(proc) || !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			procedures = append(procedures, struct {
				file        parsedFile
				proc        sourceProcedure
				moduleDecls map[string]sourceDeclaration
			}{file: file, proc: proc, moduleDecls: moduleDecls})
		}
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		for call := range procedure.proc.Calls.All() {
			if targetKey, _, ok := arrayPrivateTargetForCall(ctx, targets, call); ok {
				dependents[targetKey] = append(dependents[targetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	contributions := make(arrayByRefAllocationSummaries, len(procedures))
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		procedure := procedures[index]
		key := arrayProcedureKey(procedure.proc)
		if !arrayProcedureIsParticipant(ctx, procedure.proc) {
			continue
		}
		if ctx.arrayStats != nil && procedure.proc.Graph != nil {
			ctx.arrayStats.addCFGWalk()
		}
		value := arrayByRefAllocationSummaryForProcedure(procedure.file, procedure.proc, summaries, ctx)
		old := arrayByRefAllocationSummaries{key: contributions[key]}
		fresh := arrayByRefAllocationSummaries{key: value}
		if arrayByRefAllocationSummariesEqual(old, fresh) {
			continue
		}
		if len(value) == 0 {
			delete(contributions, key)
			delete(summaries, key)
		} else {
			contributions[key] = value
			summaries[key] = value
		}
		for _, dependent := range dependents[key] {
			if !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
		}
	}
	return summaries
}

func arrayByRefAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, summaries arrayByRefAllocationSummaries, ctx analysisContext) map[int]bool {
	if proc.Graph == nil {
		return nil
	}
	parameters := map[string]int{}
	for index, parameter := range proc.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) {
			parameters[strings.ToLower(parameter.Name)] = index
		}
	}
	if len(parameters) == 0 {
		return nil
	}
	flowCtx := ctx
	flowCtx.arrayByRefAllocations = summaries
	moduleDecls := file.moduleDecls()
	return arrayByRefFlowAllocations(file, proc, flowCtx, moduleDecls)
}

// arrayByRefFlowAllocations proves ByRef array outputs at normal procedure
// exits. Unlike the direct-assignment summary above, this pass keeps the
// branch-local state established by an IsArray/type guard and meets it at the
// normal exit. It therefore recognizes helpers that fill the same output on
// multiple accepted input branches while excluding paths that terminate in a
// direct or project-local error raiser.
func arrayByRefFlowAllocations(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) map[int]bool {
	if proc.Graph == nil {
		return nil
	}
	variables := arrayVariables(file, proc, moduleDecls)
	initial := arrayInitialState(variables)
	graph := arrayVBA227Graph(proc, ctx)
	moduleArrays := arrayModuleNamesForProcedure(file, proc, moduleDecls)
	localGoSubAllocations := arrayLocalGoSubAllocationSummaries(proc, &graph, variables, ctx, arrayOptionBase(file), arrayIntegerConstants(file, proc, nil, nil), moduleArrays)
	parameterNames := make([]string, 0, proc.Params.Len())
	for _, parameter := range proc.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) {
			parameterNames = append(parameterNames, strings.ToLower(parameter.Name))
		}
	}
	var normalExit map[string]arrayValue
	hasNormalExit := false
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		out, _ := (Analyzer{}).arrayTransfer(file, proc, ctx, variables, in, text, line, nil, nil)
		out = applyArrayLocalGoSubStatementEffects(out, text, localGoSubAllocations)
		for _, call := range arrayCallsAtLine(proc.Calls, line) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
				out = applyArrayConditionalByRefCallEffects(out, proc, call, ctx)
			} else {
				out = applyArrayByRefCallEffects(out, proc, call, ctx)
			}
			out = applyArrayLocalGoSubEffects(out, proc, call, localGoSubAllocations)
		}
		return out
	}
	edgeState := func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
		out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
		if edge.To == graph.NormalExit() {
			if !hasNormalExit {
				normalExit = make(map[string]arrayValue, len(parameterNames))
				for _, name := range parameterNames {
					normalExit[name] = out[name]
				}
				hasNormalExit = true
			} else {
				for _, name := range parameterNames {
					normalExit[name] = meetArrayValue(normalExit[name], out[name])
				}
			}
		}
		return out
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, arrayAllocationTransferIsReliable, ctx.arrayStats)
	if !hasNormalExit {
		return nil
	}
	allocated := map[int]bool{}
	for index, parameter := range proc.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) {
			continue
		}
		value, ok := normalExit[strings.ToLower(parameter.Name)]
		if ok && value.kind == arrayAllocated && value.knownArray {
			allocated[index] = true
		}
	}
	return allocated
}

// inferArrayByRefConditionalAllocations recognizes the narrow, common output
// helper contract used by argument adapters:
//
//	If items.Count = 0 Then Exit Sub
//	ReDim values(0 To items.Count - 1)
//
// The guard must dominate the ReDim, and every normal path from the guard's
// non-exit branch must pass through that ReDim. This keeps the summary tied to
// the helper's actual control flow and avoids turning an arbitrary conditional
// ReDim into an allocation guarantee.
func inferArrayByRefConditionalAllocations(files []parsedFile) arrayByRefConditionalAllocations {
	summaries := arrayByRefConditionalAllocations{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if proc.Graph == nil {
				continue
			}
			parameters := map[string]int{}
			for index, parameter := range proc.Params.AllIndexed() {
				parameters[strings.ToLower(cleanIdentifier(parameter.Name))] = index
			}
			if len(parameters) == 0 {
				continue
			}
			guards := map[string]struct {
				statementID int
				line        int
				parameter   int
			}{}
			for statement := range proc.Statements.All() {
				match := arrayByRefCountExitRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
				if len(match) != 2 {
					continue
				}
				parameter, ok := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				if !ok {
					continue
				}
				guards[strings.ToLower(cleanIdentifier(match[1]))] = struct {
					statementID int
					line        int
					parameter   int
				}{statementID: statement.ID, line: statement.Range.StartLine, parameter: parameter}
			}
			if len(guards) == 0 {
				continue
			}
			for statement := range proc.Statements.All() {
				match := arrayByRefCountRedimRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
				if len(match) != 3 || arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
					continue
				}
				output, outputOK := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				guard, guardOK := guards[strings.ToLower(cleanIdentifier(match[2]))]
				if !outputOK || !guardOK || statement.Range.StartLine <= guard.line {
					continue
				}
				if output < 0 || output >= proc.Params.Len() || !parameterIsByRefArray(proc.Params.valueAt(output)) {
					continue
				}
				guardBlock, guardBlockOK := proc.Graph.BlockForStatement(guard.statementID)
				redimBlock, redimBlockOK := proc.Graph.BlockForStatement(statement.ID)
				if !guardBlockOK || !redimBlockOK {
					continue
				}
				guardDominatesRedim := false
				for _, candidate := range proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[redimBlock.ID] {
					if candidate == guardBlock.ID {
						guardDominatesRedim = true
						break
					}
				}
				if !guardDominatesRedim {
					continue
				}
				if !arrayFalseBranchRequiresBlock(*proc.Graph, guardBlock.ID, redimBlock.ID) {
					continue
				}
				key := arrayProcedureKey(proc)
				if summaries[key] == nil {
					summaries[key] = map[int]int{}
				}
				// Conflicting output contracts remain unknown rather than being
				// overwritten by declaration order.
				if previous, exists := summaries[key][output]; exists && previous != guard.parameter {
					delete(summaries[key], output)
					continue
				}
				summaries[key][output] = guard.parameter
			}
			if len(summaries[arrayProcedureKey(proc)]) == 0 {
				delete(summaries, arrayProcedureKey(proc))
			}
		}
	}
	return summaries
}

// inferArrayByRefLengthAllocations recognizes a helper that returns the
// successful length of a ByRef array through a paired ByRef scalar output:
//
// \tbyteLength = UBound(bytes) - LBound(bytes) + 1
//
// The assignment must dominate the procedure's normal exit, or be the normal
// branch of an explicit zero-length guard whose sibling assigns zero. A
// positive length at a caller then proves that the array-bound query completed
// successfully, without making the array unconditionally allocated on other
// paths.
func inferArrayByRefLengthAllocations(files []parsedFile) arrayByRefLengthAllocations {
	summaries := arrayByRefLengthAllocations{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if proc.Graph == nil {
				continue
			}
			parameters := map[string]int{}
			for index, parameter := range proc.Params.AllIndexed() {
				parameters[strings.ToLower(cleanIdentifier(parameter.Name))] = index
			}
			if len(parameters) == 0 {
				continue
			}
			dominators := arrayProcedureNormalExitDominators(proc)
			for statement := range proc.Statements.All() {
				text := strings.TrimSpace(normalizedCodeLine(statement.Text))
				match := arrayByRefLengthFullRe.FindStringSubmatch(text)
				if len(match) == 4 && !strings.EqualFold(cleanIdentifier(match[2]), cleanIdentifier(match[3])) {
					match = nil
				}
				if len(match) == 0 {
					upper := arrayByRefLengthUpperRe.FindStringSubmatch(text)
					if len(upper) == 3 {
						match = []string{upper[0], upper[1], upper[2], upper[2]}
					}
				}
				if len(match) != 4 {
					continue
				}
				lengthIndex, lengthOK := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				arrayIndex, arrayOK := parameters[strings.ToLower(cleanIdentifier(match[2]))]
				if !lengthOK || !arrayOK || lengthIndex == arrayIndex || !parameterIsByRefScalar(proc.Params.valueAt(lengthIndex)) || !parameterIsByRefArray(proc.Params.valueAt(arrayIndex)) {
					continue
				}
				dominatesExit := arrayProcedureBlockDominatesNormalExit(proc, statement.ID, dominators)
				if !dominatesExit {
					parent, parentOK := arrayByRefLengthGuard(proc, statement.ID)
					if !parentOK || !arrayProcedureBlockDominatesNormalExit(proc, parent.ID, dominators) || !arrayByRefLengthHasZeroBranch(proc, parent.ID, statement.ID, match[1]) {
						continue
					}
				}
				key := arrayProcedureKey(proc)
				if summaries[key] == nil {
					summaries[key] = map[int]int{}
				}
				if previous, exists := summaries[key][arrayIndex]; exists && previous != lengthIndex {
					delete(summaries[key], arrayIndex)
					continue
				}
				summaries[key][arrayIndex] = lengthIndex
			}
			if len(summaries[arrayProcedureKey(proc)]) == 0 {
				delete(summaries, arrayProcedureKey(proc))
			}
		}
	}
	return summaries
}

func arrayProcedureStatementByID(proc sourceProcedure, id int) (procedureir.Statement, bool) {
	for statement := range proc.Statements.All() {
		if statement.ID == id {
			return statement, true
		}
	}
	return procedureir.Statement{}, false
}

func arrayByRefLengthGuard(proc sourceProcedure, statementID int) (procedureir.Statement, bool) {
	seen := map[int]bool{}
	for statementID != 0 && !seen[statementID] {
		seen[statementID] = true
		statement, ok := arrayProcedureStatementByID(proc, statementID)
		if !ok {
			return procedureir.Statement{}, false
		}
		if statement.Kind == procedureir.StatementIf {
			return statement, true
		}
		statementID = statement.ParentID
	}
	return procedureir.Statement{}, false
}

func arrayByRefLengthHasZeroBranch(proc sourceProcedure, parentID, formulaID int, lengthName string) bool {
	for statement := range proc.Statements.All() {
		if statement.ID == formulaID || statement.ParentID != parentID {
			continue
		}
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(statement.Text))
		if ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(lengthName)) && strings.TrimSpace(rhs) == "0" {
			return true
		}
	}
	return false
}

func arrayFalseBranchRequiresBlock(graph vbacfg.Graph, guardBlock, requiredBlock vbacfg.BlockID) bool {
	queue := make([]vbacfg.BlockID, 0, 1)
	for _, edge := range graph.Edges {
		if edge.From == guardBlock && edge.Kind == vbacfg.EdgeBranchFalse && edge.Class == vbacfg.EdgeNormal {
			queue = append(queue, edge.To)
		}
	}
	if len(queue) == 0 {
		return false
	}
	visited := map[vbacfg.BlockID]bool{}
	reachedRequired := false
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if visited[current] {
			continue
		}
		visited[current] = true
		if current == requiredBlock {
			reachedRequired = true
			continue
		}
		if current == graph.NormalExit {
			return false
		}
		for _, edge := range graph.Edges {
			if edge.From != current || edge.Class != vbacfg.EdgeNormal {
				continue
			}
			queue = append(queue, edge.To)
		}
	}
	return reachedRequired
}

func arrayByRefAllocationSummariesEqual(left, right arrayByRefAllocationSummaries) bool {
	if len(left) != len(right) {
		return false
	}
	for procedure, parameters := range left {
		other, ok := right[procedure]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index := range parameters {
			if !other[index] {
				return false
			}
		}
	}
	return true
}

func inferArrayModuleConfigurationStates(files []parsedFile, summaries arrayModuleAllocationSummaries) map[string]arrayModuleConfigurationState {
	states := map[string]arrayModuleConfigurationState{}
	for _, file := range files {
		state := arrayModuleConfigurationState{byProcedure: map[string]map[string]bool{}}
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(strings.TrimSpace(proc.Name))
			if !strings.HasPrefix(name, "configure") {
				continue
			}
			arrays := summaries[arrayProcedureKey(proc)]
			if len(arrays) == 0 {
				continue
			}
			state.byProcedure[name] = cloneArrayNameSet(arrays)
			switch name {
			case "configuredatatable":
				state.dataTable = mergeArrayNameSets(state.dataTable, arrays)
			case "configuregenericcollection":
				state.genericCollection = mergeArrayNameSets(state.genericCollection, arrays)
			}
		}
		if len(state.byProcedure) > 0 {
			states[file.Path] = state
		}
	}
	return states
}

func cloneArrayNameSet(values map[string]bool) map[string]bool {
	if len(values) == 0 {
		return nil
	}
	clone := make(map[string]bool, len(values))
	for name, allocated := range values {
		clone[name] = allocated
	}
	return clone
}

func mergeArrayNameSets(left, right map[string]bool) map[string]bool {
	if len(right) == 0 {
		return left
	}
	if left == nil {
		left = map[string]bool{}
	}
	for name := range right {
		left[name] = true
	}
	return left
}

func applyArrayModuleCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	conditional := arrayProcedureLineHasInlineConditional(file, call.Range.StartLine)
	if conditional && arrayProcedureLineInlineConditionIsFalse(file, call.Range.StartLine) {
		return state
	}
	key, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	// Callers pass a block-local state to this transfer callback. The legacy
	// walkers keep predecessor input separate before invoking it, while the
	// compact cursor owns the map outright; update that private state in place
	// so module/ByRef effects do not reintroduce whole-state copies.
	updated := state
	if updated == nil {
		updated = arrayFlowState{}
	}
	markArgument := func(name string) {
		name = strings.ToLower(cleanIdentifier(name))
		variable, known := variables[name]
		if !known || !variable.isArray {
			return
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	markModule := func(name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			return
		}
		declaration, declared := moduleDecls[name]
		if !declared || !declaration.Array || declaration.Parameter {
			return
		}
		markArgument(name)
	}
	if !conditional {
		for name := range ctx.arrayModuleAllocations[key] {
			markModule(name)
		}
	}
	arguments, mapped := arrayCallFormalArguments(proc, target, call)
	if mapped && !conditional {
		for index := range ctx.arrayByRefAllocations[key] {
			if index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
				continue
			}
			markArgument(arguments[index])
		}
		for outputIndex, countIndex := range ctx.arrayByRefConditionalAllocations[key] {
			if outputIndex < 0 || outputIndex >= target.Params.Len() || countIndex < 0 || countIndex >= target.Params.Len() {
				continue
			}
			outputName := directArrayArgumentName(arguments[outputIndex])
			countSource := directArrayArgumentName(arguments[countIndex])
			if outputName == "" || countSource == "" {
				continue
			}
			name := strings.ToLower(outputName)
			variable, known := variables[name]
			if !known || !variable.isArray {
				continue
			}
			value := updated[name]
			if value.kind == arrayAllocated && value.knownArray {
				continue
			}
			if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, countSource) {
				continue
			}
			value.allocationCountSource = countSource
			updated[name] = value
		}
		for outputIndex, lengthIndex := range ctx.arrayByRefLengthAllocations[key] {
			if outputIndex >= len(arguments) || lengthIndex < 0 || lengthIndex >= len(arguments) {
				continue
			}
			outputName := directArrayArgumentName(arguments[outputIndex])
			lengthSource := directArrayArgumentName(arguments[lengthIndex])
			if outputName == "" || lengthSource == "" {
				continue
			}
			name := strings.ToLower(outputName)
			variable, known := variables[name]
			if !known || !variable.isArray {
				continue
			}
			value := updated[name]
			if value.kind == arrayAllocated && value.knownArray {
				continue
			}
			if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, lengthSource) {
				continue
			}
			value.allocationCountSource = lengthSource
			updated[name] = value
		}
		for name := range arrayConfigurationArraysForGuard(file, target, arguments, ctx.arrayModuleConfigurations[file.Path]) {
			markModule(name)
		}
	}
	if !ctx.arraySkipModuleInvalidationEffects {
		invalidated := arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)
		for name := range invalidated {
			if declarations.shadowsModule(name) {
				continue
			}
			if _, tracked := updated[name]; tracked {
				updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
			}
		}
	}
	return updated
}

// arrayPrivateModuleArrayInvalidations returns the normal-return invalidation
// summary for a project-local helper. A precomputed summary is used by the
// production context; focused compatibility callers compute the same summary
// on demand.
func arrayPrivateModuleArrayInvalidations(file parsedFile, target sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	key := arrayProcedureKey(target)
	if ctx.arrayModuleInvalidations != nil {
		if summary, ok := ctx.arrayModuleInvalidations[key]; ok {
			return summary
		}
	}
	return arrayPrivateModuleArrayInvalidationsWithVisiting(file, target, moduleDecls, ctx, map[string]bool{})
}

// arrayPrivateModuleArrayInvalidationsWithVisiting identifies module arrays
// that are not proven allocated at a target's normal exit. The summary starts
// from allocated arrays so it models the effect on a caller that already has
// a valid module-array allocation. Direct operations are evaluated through the
// normal CFG, and resolved local calls are summarized recursively.
func arrayPrivateModuleArrayInvalidationsWithVisiting(file parsedFile, target sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, visiting map[string]bool) map[string]bool {
	key := arrayProcedureKey(target)
	if !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
		return nil
	}
	names := arrayModuleNamesForProcedure(file, target, moduleDecls)
	if len(names) == 0 {
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = nil
		}
		return nil
	}
	if visiting[key] {
		// A recursive effect cycle has no finite normal-return proof. Keep all
		// visible module arrays conservative rather than assuming the cycle is
		// read-only.
		return names
	}
	if ctx.arrayModuleInvalidations != nil {
		if summary, ok := ctx.arrayModuleInvalidations[key]; ok {
			return summary
		}
	}
	visiting[key] = true
	defer delete(visiting, key)

	variables := arrayVariables(file, target, moduleDecls)
	initial := arrayInitialState(variables)
	for name := range names {
		value := initial[name]
		value.kind = arrayAllocated
		value.knownArray = true
		initial[name] = value
	}
	if target.Graph == nil {
		state := initial
		constants := arrayIntegerConstants(file, target, nil, nil)
		if target.Statements.Len() > 0 {
			for statement := range target.Statements.All() {
				line := statement.Range.StartLine
				if line < 1 {
					line = target.StartLine
				}
				text := strings.TrimSpace(normalizedCodeLine(statement.Text))
				if text == "" && line >= 1 && line <= len(file.Lines) {
					text = normalizedCodeLine(file.Lines[line-1])
				}
				if text == "" {
					continue
				}
				state = arrayModuleSummaryTransfer(file, target, ctx, variables, state, text, line, constants, moduleDecls, names, visiting)
			}
		} else {
			for line := target.StartLine; line <= target.EndLine && line <= len(file.Lines); line++ {
				state = arrayModuleSummaryTransfer(file, target, ctx, variables, state, normalizedCodeLine(file.Lines[line-1]), line, constants, moduleDecls, names, visiting)
			}
		}
		result := arrayModuleInvalidationsFromState(names, state)
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = result
		}
		return result
	}

	graph := target.Graph.View(vbacfg.EdgeFilter{NormalOnly: true, WithoutNormalErrRaiseContinuation: true})
	constants := arrayIntegerConstants(file, target, nil, nil)
	var normalExit arrayFlowState
	hasNormalExit := false
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		return arrayModuleSummaryTransfer(file, target, ctx, variables, in, text, line, constants, moduleDecls, names, visiting)
	}
	edgeState := func(_ vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		if edge.To != graph.NormalExit() {
			return out
		}
		if !hasNormalExit {
			normalExit = cloneArrayState(out)
			hasNormalExit = true
		} else {
			normalExit = meetArrayState(normalExit, out)
		}
		return out
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, nil, ctx.arrayStats)
	if !hasNormalExit {
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = nil
		}
		return nil
	}
	result := arrayModuleInvalidationsFromState(names, normalExit)
	if ctx.arrayModuleInvalidationCacheWritable && key != "" {
		ctx.arrayModuleInvalidations[key] = result
	}
	return result
}

func arrayModuleInvalidationsFromState(names map[string]bool, state arrayFlowState) map[string]bool {
	invalidated := map[string]bool{}
	for name := range names {
		value, known := state[name]
		if !known || value.kind != arrayAllocated || !value.knownArray {
			invalidated[name] = true
		}
	}
	if len(invalidated) == 0 {
		return nil
	}
	return invalidated
}

func arrayModuleSummaryTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	if condition, body, ok := arrayIfThenParts(text); ok && strings.TrimSpace(body) != "" {
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		condition = strings.TrimSpace(condition)
		lowerCondition := strings.ToLower(condition)
		switch {
		case strings.HasPrefix(lowerCondition, "if "):
			condition = strings.TrimSpace(condition[len("if "):])
		case strings.HasPrefix(lowerCondition, "elseif "):
			condition = strings.TrimSpace(condition[len("elseif "):])
		}
		thenState := arrayModuleSummaryTransferParts(file, proc, ctx, variables, cloneArrayState(state), thenBody, line, constants, moduleDecls, moduleArrays, visiting)
		elseState := cloneArrayState(state)
		if hasElse {
			elseState = arrayModuleSummaryTransferParts(file, proc, ctx, variables, elseState, elseBody, line, constants, moduleDecls, moduleArrays, visiting)
		}
		result := meetArrayState(thenState, elseState)
		applyCalls := func(value arrayFlowState, conditionValue, conditionKnown bool) arrayFlowState {
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				if conditionKnown && !arrayInlineConditionalCallIsReachable(file, call, conditionValue, hasElse) {
					continue
				}
				value = applyArrayModuleSummaryCallEffects(value, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
			}
			return value
		}
		if value, known := arraySourceOrderConstantBoolean(condition, constants); known {
			if value {
				result = applyCalls(thenState, true, true)
			} else if hasElse {
				result = applyCalls(elseState, false, true)
			} else {
				result = state
			}
		} else if !hasElse {
			result = applyCalls(meetArrayState(thenState, state), false, false)
		} else {
			result = applyCalls(result, false, false)
		}
		return result
	}
	state = arrayModuleSummaryTransferParts(file, proc, ctx, variables, state, text, line, constants, moduleDecls, moduleArrays, visiting)
	callsAtLine := arrayCallsAtLine(proc.Calls, line)
	for _, call := range callsAtLine {
		state = applyArrayModuleSummaryCallEffects(state, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
	}
	return state
}

func arrayModuleSummaryTransferParts(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	for _, part := range splitRangeValueSourceStatements(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		if condition, body, ok := arrayIfThenParts(part); ok && strings.TrimSpace(body) != "" {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			condition = strings.TrimSpace(condition)
			lowerCondition := strings.ToLower(condition)
			switch {
			case strings.HasPrefix(lowerCondition, "if "):
				condition = strings.TrimSpace(condition[len("if "):])
			case strings.HasPrefix(lowerCondition, "elseif "):
				condition = strings.TrimSpace(condition[len("elseif "):])
			}
			thenState := arrayModuleSummaryTransferParts(file, proc, ctx, variables, cloneArrayState(state), thenBody, line, constants, moduleDecls, moduleArrays, visiting)
			elseState := cloneArrayState(state)
			if hasElse {
				elseState = arrayModuleSummaryTransferParts(file, proc, ctx, variables, elseState, elseBody, line, constants, moduleDecls, moduleArrays, visiting)
			}
			if value, known := arraySourceOrderConstantBoolean(condition, constants); known {
				if value {
					state = thenState
				} else if hasElse {
					state = elseState
				}
			} else if hasElse {
				state = meetArrayState(thenState, elseState)
			} else {
				state = meetArrayState(thenState, state)
			}
			continue
		}
		state, _ = (Analyzer{}).arrayTransfer(file, proc, ctx, variables, state, part, line, constants, nil)
	}
	return state
}

func applyArrayModuleSummaryCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	// The procedure IR also represents an indexed array expression such as
	// `result(index)` or `mZipWork(offset)` as a CallSite. Those expressions
	// are handled by arrayTransfer; they are not procedure calls whose module
	// effects belong in this summary. Without this guard, the resolver can
	// reinterpret the expression as a source-local target and recursively
	// invalidate every module array while summarizing an otherwise read-only
	// indexed assignment.
	if call.Callee.Receiver == nil {
		name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
		if variable, ok := variables[name]; ok && (variable.isArray || variable.isVariant) {
			return state
		}
	}
	key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !resolved {
		target, resolved = arraySourceModuleTargetForCall(file, call, ctx)
		if !resolved {
			return state
		}
		key = arrayProcedureKey(target)
	}
	if !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
		return state
	}
	updated := cloneArrayState(state)
	if !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
		for name := range ctx.arrayModuleAllocations[key] {
			name = strings.ToLower(cleanIdentifier(name))
			if moduleArrays[name] {
				value := updated[name]
				value.kind = arrayAllocated
				value.knownArray = true
				updated[name] = value
			}
		}
	}
	for name := range arrayPrivateModuleArrayInvalidationsWithVisiting(file, target, moduleDecls, ctx, visiting) {
		if moduleArrays[name] {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	if !procedureHasByRefArrayParameter(target) {
		return updated
	}
	bindings, mapped := arrayCallArgumentBindings(proc, target, call)
	if !mapped {
		return updated
	}
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		name := strings.ToLower(directArrayArgumentName(binding.text))
		if !moduleArrays[name] {
			continue
		}
		if ctx.arrayByRefAllocations[key][binding.parameterIndex] {
			updated[name] = arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

func arrayPrivateCallMayInvalidateModuleArray(file parsedFile, caller sourceProcedure, target sourceProcedure, call procedureir.CallSite, moduleDecls map[string]sourceDeclaration, ctx analysisContext) bool {
	if len(arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)) > 0 {
		return true
	}
	if !procedureHasByRefArrayParameter(target) {
		return false
	}
	bindings, mapped := arrayCallArgumentBindings(caller, target, call)
	if !mapped {
		return true
	}
	for _, binding := range bindings {
		name := strings.ToLower(directArrayArgumentName(binding.text))
		declaration, declared := moduleDecls[name]
		if !declared || !declaration.Array || declaration.Parameter || binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			return true
		}
	}
	return false
}

// applyArrayUnknownModuleCallEffects is used by the recovered source-order
// path, where a public call cannot be matched to a private module effect
// summary. A source-local public target can still be inspected for direct
// module-array effects, including when it receives no explicit array
// argument. Calls that do not identify a source-local target are left alone:
// treating every unresolved/external call as a mutation of every visible
// module array turns unrelated object construction and host calls into false
// positives.
func applyArrayUnknownModuleCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	if ctx.arraySkipModuleInvalidationEffects {
		return state
	}
	if arrayProcedureLineInlineConditionIsFalse(file, call.Range.StartLine) {
		return state
	}
	if _, _, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call); resolved {
		return state
	}
	target, ok := arraySourceModuleTargetForCall(file, call, ctx)
	if !ok || !procedureUsesModuleArray(file, target, moduleDecls) {
		return state
	}
	invalidated := arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)
	if len(invalidated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range invalidated {
		if declarations.shadowsModule(name) {
			continue
		}
		variable, known := variables[name]
		if !known || !variable.isArray {
			continue
		}
		if _, tracked := updated[name]; tracked {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

// arraySourceModuleTargetForCall recovers a source-local target for the
// source-order fallback. The project resolver may report a public target, but
// it may also be incomplete for a recovered call. In the latter case a unique
// same-module source procedure is sufficient evidence for inspecting that
// procedure's module-array accesses; an absent target remains an external or
// late-bound call and must not invalidate all module arrays.
func arraySourceModuleTargetForCall(file parsedFile, call procedureir.CallSite, ctx analysisContext) (sourceProcedure, bool) {
	procedures := file.procedureView()
	if procedures.Len() == 0 {
		return sourceProcedure{}, false
	}
	resolution := call.Resolution
	if ctx.procedureResolver != nil {
		resolution = ctx.procedureResolver.ResolveCall(call)
	}
	if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
		qualifiedName := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
		for index := 0; index < procedures.Len(); index++ {
			target := procedures.valueAt(index)
			if strings.EqualFold(arrayProcedureKey(target), qualifiedName) {
				return target, true
			}
		}
	}
	// A receiver-bearing call that was not resolved to the exact source target
	// belongs to another object or to a late-bound member. Do not reinterpret a
	// same-module procedure with the same member name as its target.
	if call.Callee.Receiver != nil {
		return sourceProcedure{}, false
	}

	baseName := cleanIdentifier(strings.TrimPrefix(strings.TrimSpace(call.Callee.BaseName), "New "))
	if baseName == "" {
		baseName = cleanIdentifier(strings.TrimPrefix(strings.TrimSpace(call.Callee.Text), "New "))
	}
	callerModule := strings.TrimSpace(call.Caller.QualifiedName)
	if dot := strings.IndexByte(callerModule, '.'); dot >= 0 {
		callerModule = callerModule[:dot]
	}
	if callerModule == "" {
		callerModule = strings.TrimSpace(call.Module)
	}
	if callerModule == "" {
		callerModule = strings.TrimSpace(file.Module)
	}
	var match sourceProcedure
	matched := 0
	for index := 0; index < procedures.Len(); index++ {
		candidate := procedures.valueAt(index)
		if !strings.EqualFold(strings.TrimSpace(candidate.Name), baseName) ||
			!strings.EqualFold(strings.TrimSpace(candidate.Module), callerModule) {
			continue
		}
		visibility := strings.TrimSpace(candidate.Visibility)
		if strings.EqualFold(visibility, "Private") || strings.EqualFold(visibility, "Friend") {
			continue
		}
		match = candidate
		matched++
	}
	if matched == 1 {
		return match, true
	}
	return sourceProcedure{}, false
}

// applyArrayByRefCallEffects carries the invalidating side of a private
// ByRef-array contract. The ordinary CFG walk records proven allocation
// outputs, but the recovered source-order fallback also has to account for a
// preceding helper that can Erase or otherwise replace the caller's array.
// Read-only helpers preserve the caller's allocation state; helpers whose
// parameter effect cannot be proven make it unknown.
func applyArrayByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	key, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return applyArrayUnknownByRefCallEffects(state, proc, call, ctx)
	}
	if !procedureHasByRefArrayParameter(target) {
		return state
	}
	bindings, ok := arrayCallArgumentBindings(proc, target, call)
	if !ok {
		return state
	}
	// The compact CFG solver may revisit sibling branch blocks with a state map
	// that shares storage with their predecessor. Keep this call's mutation
	// isolated so an invalidating true branch cannot poison its false sibling.
	updated := cloneArrayState(state)
	apply := func(index int, argument string) {
		if index < 0 || index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
			return
		}
		name := directArrayArgumentName(argument)
		if name == "" {
			return
		}
		if ctx.arrayByRefAllocations[key][index] {
			value := updated[name]
			value.kind = arrayAllocated
			value.knownArray = true
			updated[name] = value
			return
		}
		// Conditional and paired-length output contracts carry a stronger
		// caller-side fact than the generic invalidation scan. Their output may
		// be unallocated for the zero-length case, but the existing count/length
		// refinement will establish the successful branch without losing that
		// relation to an unknown state here.
		if _, conditional := ctx.arrayByRefConditionalAllocations[key][index]; conditional {
			return
		}
		if _, pairedLength := ctx.arrayByRefLengthAllocations[key][index]; pairedLength {
			return
		}
		if arrayByRefParameterMayInvalidate(target, index, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	for _, binding := range bindings {
		apply(binding.parameterIndex, binding.text)
	}
	return updated
}

func applyArrayUnknownByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	if _, _, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call); resolved {
		return state
	}
	updated := cloneArrayState(state)
	markUnknown := func(argument string) {
		name := directArrayArgumentName(argument)
		if name == "" {
			return
		}
		if _, tracked := updated[name]; tracked {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	for _, argument := range call.Arguments.Named {
		markUnknown(argument.ValueText)
	}
	for _, argument := range arrayCallArgumentTexts(proc, call) {
		markUnknown(argument)
	}
	return updated
}

// applyArrayConditionalByRefCallEffects joins the two possible outcomes of
// an inline conditional call without treating the call's allocation summary
// as unconditional. A helper that may invalidate its ByRef array makes the
// joined state unknown; a helper that only allocates leaves the prior state,
// which is the conservative meet of the conditional paths.
func applyArrayConditionalByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return applyArrayUnknownByRefCallEffects(state, proc, call, ctx)
	}
	if !procedureHasByRefArrayParameter(target) {
		return state
	}
	bindings, ok := arrayCallArgumentBindings(proc, target, call)
	if !ok {
		return state
	}
	updated := cloneArrayState(state)
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		name := directArrayArgumentName(binding.text)
		if name == "" {
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

// arrayByRefParameterMayInvalidate reports whether a private ByRef-array
// parameter can lose its allocation while the procedure returns normally.
// Element writes and reads preserve an allocated input; Erase, unknown whole
// array replacement, and an unproven nested ByRef call do not. The recursion
// guard keeps recursive helper cycles from being treated as an additional
// mutation. Any direct mutation in the current cycle member is still found by
// its own scan.
func arrayByRefParameterMayInvalidate(proc sourceProcedure, parameterIndex int, ctx analysisContext, visiting map[string]bool) bool {
	if parameterIndex < 0 || parameterIndex >= proc.Params.Len() || !parameterIsByRefArray(proc.Params.valueAt(parameterIndex)) {
		return true
	}
	key := strings.ToLower(arrayProcedureKey(proc)) + "#" + strconv.Itoa(parameterIndex)
	if visiting[key] {
		return false
	}
	visiting[key] = true
	defer delete(visiting, key)
	name := strings.ToLower(cleanIdentifier(proc.Params.valueAt(parameterIndex).Name))
	names := map[string]bool{name: true}
	for statement := range proc.Statements.All() {
		if !arrayByRefStatementReachable(proc, statement) {
			continue
		}
		if statement.Recovered {
			return true
		}
		for _, part := range splitRangeValueSourceStatements(statement.Text) {
			part = strings.TrimSpace(part)
			if part == "" {
				continue
			}
			if arrayByRefParameterInlineMutation(part, names, ctx) {
				return true
			}
		}
		for _, nested := range arrayCallsAtLine(proc.Calls, statement.Range.StartLine) {
			if arrayByRefCallIsReadOnly(nested) {
				continue
			}
			if !arrayCallPassesDirectArrayArgument(proc, nested, name) {
				continue
			}
			nestedKey, nestedTarget, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, nested)
			if !resolved {
				return true
			}
			bindings, mapped := arrayCallArgumentBindings(proc, nestedTarget, nested)
			if !mapped {
				return true
			}
			for _, binding := range bindings {
				if directArrayArgumentName(binding.text) != name || binding.parameterIndex >= nestedTarget.Params.Len() || !parameterIsByRefArray(nestedTarget.Params.valueAt(binding.parameterIndex)) {
					continue
				}
				if ctx.arrayByRefAllocations[nestedKey][binding.parameterIndex] {
					continue
				}
				if arrayByRefParameterMayInvalidate(nestedTarget, binding.parameterIndex, ctx, visiting) {
					return true
				}
			}
		}
	}
	return false
}

func arrayByRefCallIsReadOnly(call procedureir.CallSite) bool {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return true
	}
	switch strings.ToLower(cleanIdentifier(call.Callee.BaseName)) {
	case "lbound", "ubound":
		return true
	default:
		return false
	}
}

func arrayByRefStatementReachable(proc sourceProcedure, statement procedureir.Statement) bool {
	if proc.Graph == nil {
		return true
	}
	block, ok := proc.Graph.BlockForStatement(statement.ID)
	if !ok {
		return true
	}
	return proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).IsReachable(block.ID)
}

func arrayByRefParameterInlineMutation(text string, names map[string]bool, ctx analysisContext) bool {
	if condition, body, ok := arrayIfThenParts(text); ok && strings.TrimSpace(body) != "" {
		condition = strings.TrimSpace(condition)
		lowerCondition := strings.ToLower(condition)
		switch {
		case strings.HasPrefix(lowerCondition, "if "):
			condition = strings.TrimSpace(condition[len("if "):])
		case strings.HasPrefix(lowerCondition, "elseif "):
			condition = strings.TrimSpace(condition[len("elseif "):])
		}
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		branches := []string{thenBody}
		if hasElse {
			branches = append(branches, elseBody)
		}
		if value, known := arraySourceOrderConstantBoolean(condition, nil); known {
			if value {
				branches = []string{thenBody}
			} else if hasElse {
				branches = []string{elseBody}
			} else {
				branches = nil
			}
		}
		for _, branch := range branches {
			for _, statement := range splitRangeValueSourceStatements(branch) {
				if arrayByRefParameterInlineMutation(strings.TrimSpace(statement), names, ctx) {
					return true
				}
			}
		}
		return false
	}
	text = strings.TrimSpace(text)
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
		for _, target := range splitArgs(match[1]) {
			if names[strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))] {
				return true
			}
		}
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			name := strings.ToLower(cleanIdentifier(redim.name))
			if !direct || !names[name] {
				continue
			}
			// ReDim Preserve retains an already allocated input array. Its shape
			// may change, but that is not an allocation invalidation at the
			// caller boundary.
			if strings.TrimSpace(match[1]) != "" {
				continue
			}
			if !arrayStatementAllocatesName(text, name, ctx) {
				return true
			}
		}
	}
	if lhs, _, indexed, ok := arrayAssignment(text); ok && !indexed {
		name := strings.ToLower(cleanIdentifier(lhs))
		return names[name] && !arrayStatementAllocatesName(text, name, ctx)
	}
	return false
}

func arrayConfigurationArraysForGuard(file parsedFile, target sourceProcedure, arguments []string, configurations arrayModuleConfigurationState) map[string]bool {
	name := strings.ToLower(strings.TrimSpace(target.Name))
	if !strings.HasPrefix(name, "require") || !arrayGuardProcedureRejectsInvalidState(file, target) || !arrayGuardTargetsCurrentObject(target, arguments) {
		return nil
	}
	if arrays := configurations.byProcedure["configure"+strings.TrimPrefix(name, "require")]; len(arrays) > 0 {
		return arrays
	}
	if name == "requireerror" {
		if arrays := configurations.byProcedure["configureaggregateerror"]; len(arrays) > 0 {
			return arrays
		}
	}
	body := strings.ToLower(strings.Join(file.Lines[max(0, target.StartLine-1):min(len(file.Lines), target.EndLine)], "\n"))
	if strings.Contains(body, "role_data_table") {
		return configurations.dataTable
	}
	if arrayGuardUsesGenericCollectionConfiguration(body) {
		return configurations.genericCollection
	}
	return nil
}

// applyArrayInternalStorageConfiguration carries the class-instance array
// contract into Friend/Private storage members that are called through a
// configured receiver. These members intentionally do not repeat a public
// role guard: their callers have already established the owning collection,
// data-row, or aggregate-error configuration on that receiver.
func applyArrayInternalStorageConfiguration(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, configurations arrayModuleConfigurationState) arrayFlowState {
	if !strings.EqualFold(strings.TrimSpace(proc.ModuleKind), "class") {
		return state
	}
	arrays := arrayInternalStorageConfigurationArrays(proc, configurations)
	if len(arrays) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayInternalStorageConfigurationArrays(proc sourceProcedure, configurations arrayModuleConfigurationState) map[string]bool {
	name := strings.ToLower(strings.TrimSpace(proc.Name))
	switch name {
	case "internalcollectionitems", "internalcollectionkeys", "internalcollectionpriorities",
		"internaladdlookupgroup", "internalappendcollectionitem", "internalappendcollectionkey",
		"internalappendcollectionpriority", "internalqueuevalue", "internalpushvalue":
		return configurations.genericCollection
	case "internaladdwrapped":
		return mergeArrayNameSets(cloneArrayNameSet(configurations.byProcedure["configurelist"]), configurations.genericCollection)
	case "internaldatacolumns":
		return configurations.dataTable
	case "internaldatarows", "internalappendrowcell", "acceptrowchanges", "rejectrowchanges":
		return configurations.byProcedure["configuredatarow"]
	case "internalinnerexceptions":
		return configurations.byProcedure["configureaggregateerror"]
	default:
		return nil
	}
}

func arrayGuardUsesGenericCollectionConfiguration(body string) bool {
	for _, marker := range []string{
		"isgenericcollectionrole",
		"ispriorityqueuekind",
		"isdictionarycollection",
		"issetcollection",
		"mcollectionkind",
		"ROLE_DICTIONARY",
		"ROLE_HASH_SET",
		"ROLE_COLLECTION",
		"ROLE_IMMUTABLE",
		"ROLE_CONCURRENT",
	} {
		if strings.Contains(body, strings.ToLower(marker)) {
			return true
		}
	}
	return false
}

func arrayGuardProcedureRejectsInvalidState(file parsedFile, target sourceProcedure) bool {
	start := max(0, target.StartLine-1)
	end := min(len(file.Lines), target.EndLine)
	if start >= end {
		return false
	}
	body := strings.ToLower(strings.Join(file.Lines[start:end], "\n"))
	return strings.Contains(body, "err.raise") || strings.Contains(body, "raisecontracterror")
}

func arrayGuardTargetsCurrentObject(target sourceProcedure, arguments []string) bool {
	if target.Params.Len() <= 1 {
		return true
	}
	return len(arguments) > 0 && strings.EqualFold(strings.TrimSpace(arguments[0]), "me")
}

func applyArrayModuleConfigurationBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, configurations arrayModuleConfigurationState, variables map[string]arrayVariable, file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue || statement.Condition == nil {
		return state
	}
	condition := strings.ToLower(strings.TrimSpace(statement.Condition.Text))
	var arrays map[string]bool
	switch {
	case arrayPositiveGenericCollectionKindBranch(condition):
		arrays = configurations.genericCollection
	case strings.Contains(condition, "role_immutable") && !strings.Contains(condition, "<> role_immutable"):
		arrays = configurations.genericCollection
	case strings.Contains(condition, "role_list") && !strings.Contains(condition, "<> role_list"):
		arrays = configurations.byProcedure["configurelist"]
	case strings.Contains(condition, "role_data_row") && !strings.Contains(condition, "<> role_data_row"):
		arrays = configurations.byProcedure["configuredatarow"]
	case strings.Contains(condition, "role_data_table") && !strings.Contains(condition, "<> role_data_table"):
		arrays = configurations.dataTable
	}
	if len(arrays) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayPositiveGenericCollectionKindBranch(condition string) bool {
	for _, marker := range []string{
		"isgenericcollectionrole",
		"ispriorityqueuekind",
		"issortedmapkind",
		"issortedsetkind",
	} {
		if strings.Contains(condition, marker) && !strings.Contains(condition, "not "+marker) {
			return true
		}
	}
	return false
}

func inferArrayModuleAllocationSummaries(files []parsedFile, ctx analysisContext, targets map[string]sourceProcedure, byRefSummaries arrayByRefAllocationSummaries) arrayModuleAllocationSummaries {
	summaries := arrayModuleAllocationSummaries{}
	procedures := make([]struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
	}, 0)
	// Package-local synthetic callers may omit ModuleFacts. Attach one local
	// immutable instance before expanding the procedure list so compatibility
	// paths do not rebuild and rescan module source once per procedure.
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			procedures = append(procedures, struct {
				file        parsedFile
				proc        sourceProcedure
				moduleDecls map[string]sourceDeclaration
			}{file: file, proc: proc, moduleDecls: moduleDecls})
		}
	}
	if len(procedures) == 0 {
		return summaries
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	dominators := arrayProcedureDominators{}
	for _, procedure := range procedures {
		dominators[arrayProcedureKey(procedure.proc)] = arrayProcedureNormalExitDominators(procedure.proc)
	}

	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		for call := range procedure.proc.Calls.All() {
			if targetKey, _, ok := arrayPrivateTargetForCall(ctx, targets, call); ok {
				dependents[targetKey] = append(dependents[targetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	contributions := make(arrayModuleAllocationSummaries, len(procedures))
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		procedure := procedures[index]
		if !arrayProcedureIsParticipant(ctx, procedure.proc) {
			continue
		}
		key := arrayProcedureKey(procedure.proc)
		value := arrayModuleAllocationSummaryForProcedure(procedure.file, procedure.proc, procedure.moduleDecls, targets, summaries, byRefSummaries, ctx, dominators[key])
		old := arrayModuleAllocationSummaries{key: contributions[key]}
		fresh := arrayModuleAllocationSummaries{key: value}
		if arrayModuleAllocationSummariesEqual(old, fresh) {
			continue
		}
		if len(value) == 0 {
			delete(contributions, key)
			delete(summaries, key)
		} else {
			contributions[key] = value
			summaries[key] = value
		}
		for _, dependent := range dependents[key] {
			if !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
		}
	}
	return summaries
}

func arrayModuleAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, targets map[string]sourceProcedure, summaries arrayModuleAllocationSummaries, byRefSummaries arrayByRefAllocationSummaries, ctx analysisContext, dominators map[vbacfg.BlockID]bool) map[string]bool {
	moduleArrays := map[string]bool{}
	for name, declaration := range moduleDecls {
		if declaration.Array && !declaration.Parameter {
			moduleArrays[strings.ToLower(name)] = true
		}
	}
	if len(moduleArrays) == 0 || proc.Graph == nil {
		return nil
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	for name := range moduleArrays {
		if declarations.shadowsModule(name) {
			delete(moduleArrays, name)
		}
	}
	idempotentSetupArrays := arrayModuleIdempotentSetupArrays(file, proc, moduleDecls, ctx)
	allocated := map[string]bool{}
	addDirectAllocation := func(statementID int, name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if !moduleArrays[name] || (!arrayProcedureBlockDominatesNormalExit(proc, statementID, dominators) && !idempotentSetupArrays[name]) {
			return
		}
		allocated[name] = true
	}
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct && !arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
					addDirectAllocation(statement.ID, redim.name)
				}
			}
		}
		if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed {
			name := strings.ToLower(cleanIdentifier(lhs))
			if moduleArrays[name] {
				if value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx); known && value.kind == arrayAllocated && value.knownArray {
					if !arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
						addDirectAllocation(statement.ID, name)
					}
				}
			}
		}
	}
	for call := range proc.Calls.All() {
		key, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
		if !ok {
			continue
		}
		calleeArrays := summaries[key]
		calleeByRefArrays := byRefSummaries[key]
		if len(calleeArrays) == 0 && len(calleeByRefArrays) == 0 {
			continue
		}
		guaranteed := !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) && arrayProcedureBlockDominatesNormalExit(proc, call.StatementID, dominators)
		if !guaranteed && arrayProcedureHasIdempotentSetupGuard(file, proc, call.Range.StartLine, moduleDecls) {
			guaranteed = true
		}
		if !guaranteed {
			continue
		}
		for name := range calleeArrays {
			if moduleArrays[name] {
				allocated[name] = true
			}
		}
		if len(calleeByRefArrays) > 0 {
			arguments, mapped := arrayCallFormalArguments(proc, target, call)
			if mapped {
				for index := range calleeByRefArrays {
					if index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
						continue
					}
					name := strings.ToLower(directArrayArgumentName(arguments[index]))
					if moduleArrays[name] {
						allocated[name] = true
					}
				}
			}
		}
	}
	return allocated
}

// arrayModuleIdempotentSetupArrays recognizes the narrow one-time module
// initialization idiom used by private helper routines:
//
//	If ready Then Exit Sub
//	ReDim values(...)
//	ready = True
//
// The direct ReDim does not dominate the procedure's normal exit because the
// already-initialized branch exits early.  The summary can nevertheless carry
// the allocation when the Boolean guard is module-scoped, is written only to
// True by this procedure, is not written elsewhere in the module, and is the
// final executable statement.  These constraints keep an arbitrary Boolean
// branch from becoming an allocation proof.
func arrayModuleIdempotentSetupArrays(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	if proc.StartLine < 1 || proc.EndLine < proc.StartLine || proc.StartLine > len(file.Lines) {
		return nil
	}
	end := min(len(file.Lines), proc.EndLine)
	start := max(0, proc.StartLine-1)
	facts := file.moduleAnalysisFacts()
	type setupGuard struct {
		name    string
		checkAt int
	}
	guards := make([]setupGuard, 0)
	for index := start; index < end; index++ {
		name, ok := facts.sourceLineSetupGuard(index)
		if !ok {
			continue
		}
		declaration, ok := moduleDecls[name]
		if !ok || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			continue
		}
		guards = append(guards, setupGuard{name: name, checkAt: index})
	}
	if len(guards) == 0 {
		return nil
	}
	constants := arrayIntegerConstants(file, proc, nil, nil)

	lastExecutable := -1
	for index := start; index < end; index++ {
		if facts.sourceLineIsExecutable(index) {
			lastExecutable = index
		}
	}
	if lastExecutable < start {
		return nil
	}

	guardWrites := map[string][]struct {
		line int
		rhs  string
	}{}
	for _, guard := range guards {
		facts.forEachArrayOperationFor(guard.name, func(operation moduleArrayOperationFact) {
			if operation.Kind != moduleArrayWholeAssignment {
				return
			}
			guardWrites[guard.name] = append(guardWrites[guard.name], struct {
				line int
				rhs  string
			}{line: operation.Line, rhs: operation.RHS})
		})
	}

	result := map[string]bool{}
	for _, guard := range guards {
		writes := guardWrites[guard.name]
		if len(writes) != 1 || writes[0].line != lastExecutable || !strings.EqualFold(writes[0].rhs, "true") {
			continue
		}
		setAt := writes[0].line
		for index := guard.checkAt + 1; index < setAt; index++ {
			facts.forEachArrayOperationAt(index, func(operation moduleArrayOperationFact) {
				if operation.Kind != moduleArrayDirectRedim || operation.Preserve {
					return
				}
				name := operation.Name
				declaration, declared := moduleDecls[name]
				if !declared || !declaration.Array || declaration.Parameter {
					return
				}
				if moduleArrayOperationHasOtherWrite(facts, name, index) {
					return
				}
				if !arrayModuleSetupReDimIsReliable(file, proc, guard.checkAt, index, setAt, guard.name, name, constants, ctx, moduleDecls) {
					return
				}
				result[name] = true
			})
		}
	}
	if len(result) == 0 {
		return nil
	}
	return result
}

// inferArrayModuleReadyGuardStates recognizes the module-level lifecycle
// invariant used by consumers such as CSV readers:
//
//	moduleArray = Split(...)
//	ready = True
//
// and later:
//
//	If Not ready Then Exit Function
//	use moduleArray(...)
//
// The existing module allocation summary is caller-oriented and therefore
// cannot establish the state at an independently callable public procedure.
// This summary is deliberately source-owned and fail-closed so an arbitrary
// Boolean assignment never becomes an array allocation proof.
func inferArrayModuleReadyGuardStates(files []parsedFile, ctx analysisContext) arrayModuleReadyGuardStates {
	states := arrayModuleReadyGuardStates{}
	for _, file := range files {
		facts := file.moduleAnalysisFacts()
		if facts == nil {
			continue
		}
		moduleDecls := file.moduleDecls()
		for guardName, guardDeclaration := range moduleDecls {
			guardName = strings.ToLower(cleanIdentifier(guardName))
			if guardName == "" || guardDeclaration.Array || guardDeclaration.Parameter || !strings.EqualFold(strings.TrimSpace(guardDeclaration.Type), "Boolean") || !arrayModuleReadyGuardSourceOwned(file, guardDeclaration) {
				continue
			}

			var writes []moduleArrayOperationFact
			valid := true
			trueWrites := make([]moduleArrayOperationFact, 0, 1)
			facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
				if !valid {
					return
				}
				owner, owned := arrayModuleProcedureAtLine(file, operation.Line+1)
				if !owned {
					valid = false
					return
				}
				scope := newDeclarationScope(file, owner)
				scope.module = moduleDecls
				if scope.shadowsModule(guardName) {
					return
				}
				if operation.Kind != moduleArrayWholeAssignment {
					valid = false
					return
				}
				rhs := strings.ToLower(strings.TrimSpace(operation.RHS))
				if rhs != "true" && rhs != "false" {
					valid = false
					return
				}
				writes = append(writes, operation)
				if rhs == "true" {
					trueWrites = append(trueWrites, operation)
				}
			})
			if !valid || len(trueWrites) != 1 {
				continue
			}
			writer, ok := arrayModuleProcedureAtLine(file, trueWrites[0].Line+1)
			if !ok {
				continue
			}
			allocated := arrayModuleReadyGuardAllocationProof(file, writer, trueWrites[0].Line+1, moduleDecls, ctx)
			safeAllocated := map[string]bool{}
			for name := range allocated {
				if arrayModuleReadyGuardLifecycleSafe(file, guardName, map[string]bool{name: true}, facts, moduleDecls, ctx) {
					safeAllocated[name] = true
				}
			}
			if len(safeAllocated) == 0 {
				continue
			}
			if states[file.Path] == nil {
				states[file.Path] = map[string]map[string]bool{}
			}
			states[file.Path][guardName] = safeAllocated
		}
	}
	return states
}

func arrayModuleReadyGuardSourceOwned(file parsedFile, declaration sourceDeclaration) bool {
	if declaration.Line < 1 || declaration.Line > len(file.Lines) {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(file.Lines[declaration.Line-1])))
	return !strings.HasPrefix(lower, "public ") && !strings.HasPrefix(lower, "global ")
}

func arrayModuleProcedureAtLine(file parsedFile, line int) (sourceProcedure, bool) {
	if line < 1 {
		return sourceProcedure{}, false
	}
	procedures := file.procedureView()
	for index := 0; index < procedures.Len(); index++ {
		procedure := procedures.valueAt(index)
		if line >= procedure.StartLine && line <= procedure.EndLine {
			return procedure, true
		}
	}
	return sourceProcedure{}, false
}

func arrayModuleReadyGuardAllocationProof(file parsedFile, proc sourceProcedure, readyLine int, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	if proc.Graph == nil || readyLine < 1 || readyLine > len(file.Lines) {
		return nil
	}
	variables := arrayVariables(file, proc, moduleDecls)
	candidates := map[string]bool{}
	for name, declaration := range moduleDecls {
		name = strings.ToLower(cleanIdentifier(name))
		if name != "" && declaration.Array && !declaration.Parameter && arrayModuleReadyGuardSourceOwned(file, declaration) {
			if variable, known := variables[name]; known && variable.isArray {
				candidates[name] = true
			}
		}
	}
	if len(candidates) == 0 {
		return nil
	}

	graph := arrayVBA227Graph(proc, ctx)
	initial := arrayInitialState(variables)
	seenReady := false
	failed := map[string]bool{}
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		if line == readyLine {
			seenReady = true
			for name := range candidates {
				value, known := in[name]
				if !known || value.kind != arrayAllocated || !value.knownArray {
					failed[name] = true
				}
			}
		}
		out, _ := (Analyzer{}).arrayVBA227Transfer(file, proc, ctx, variables, in, text, line, nil, nil, nil)
		for _, call := range arrayCallsAtLine(proc.Calls, line) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
		}
		return out
	}
	edgeState := func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
		out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
		return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, arrayAllocationTransferIsReliable, nil)
	if !seenReady {
		return nil
	}
	result := map[string]bool{}
	for name := range candidates {
		if !failed[name] {
			result[name] = true
		}
	}
	return result
}

func arrayModuleReadyGuardLifecycleSafe(file parsedFile, guardName string, arrays map[string]bool, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration, ctx analysisContext) bool {
	if facts == nil {
		return false
	}
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		safe := true
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if !safe {
				return
			}
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
			if !ok {
				safe = false
				return
			}
			scope := newDeclarationScope(file, owner)
			scope.module = moduleDecls
			if scope.shadowsModule(name) {
				return
			}
			line := operation.Line + 1
			if line < 1 || line > len(file.Lines) {
				safe = false
				return
			}
			switch operation.Kind {
			case moduleArrayWholeAssignment:
				lhs, rhs, indexed, assigned := arrayAssignment(arrayLogicalCodeLine(file.Lines, line))
				if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), name) {
					safe = false
					return
				}
				value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx)
				if !known || value.kind != arrayAllocated || !value.knownArray {
					safe = false
				}
			case moduleArrayDirectRedim:
				if arraySummaryStatementAlwaysFails(arrayLogicalCodeLine(file.Lines, line), arrayOptionBase(file), arrayIntegerConstants(file, owner, nil, nil)) {
					safe = false
				}
			case moduleArrayErase:
				if !arrayModuleReadyGuardFalseWriteDominates(file, owner, guardName, line, facts, moduleDecls) {
					safe = false
				}
			default:
				safe = false
			}
		})
		if !safe {
			return false
		}
	}
	return true
}

func arrayModuleReadyGuardFalseWriteDominates(file parsedFile, proc sourceProcedure, guardName string, eraseLine int, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration) bool {
	if proc.Graph == nil || facts == nil || eraseLine <= proc.StartLine {
		return false
	}
	falseLines := make([]int, 0, 1)
	facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
		if operation.Kind != moduleArrayWholeAssignment || !strings.EqualFold(strings.TrimSpace(operation.RHS), "false") {
			return
		}
		line := operation.Line + 1
		owner, ok := arrayModuleProcedureAtLine(file, line)
		if !ok || owner.StartByte != proc.StartByte || owner.StartLine != proc.StartLine || owner.EndLine != proc.EndLine {
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if scope.shadowsModule(guardName) {
			return
		}
		if line < eraseLine {
			falseLines = append(falseLines, line)
		}
	})
	if len(falseLines) == 0 {
		return false
	}
	eraseStatement, eraseOK := arrayModuleStatementAtLine(proc, eraseLine)
	if !eraseOK {
		return false
	}
	normalGraph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	eraseBlock, eraseBlockOK := normalGraph.BlockForStatement(eraseStatement.ID)
	if !eraseBlockOK || !normalGraph.IsReachable(eraseBlock.ID) {
		return false
	}
	for _, falseLine := range falseLines {
		if arrayModuleReadyGuardHasWriteBetween(file, proc, guardName, falseLine, eraseLine, facts, moduleDecls) {
			continue
		}
		falseStatement, falseOK := arrayModuleStatementAtLine(proc, falseLine)
		if !falseOK {
			continue
		}
		falseBlock, falseBlockOK := normalGraph.BlockForStatement(falseStatement.ID)
		if !falseBlockOK {
			continue
		}
		for _, dominator := range normalGraph.DominatorsOf(eraseBlock.ID) {
			if dominator == falseBlock.ID {
				return true
			}
		}
	}
	return false
}

func arrayModuleReadyGuardHasWriteBetween(file parsedFile, proc sourceProcedure, guardName string, startLine, endLine int, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration) bool {
	written := false
	facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
		if written || operation.Kind != moduleArrayWholeAssignment {
			return
		}
		line := operation.Line + 1
		if line <= startLine || line >= endLine {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, line)
		if !ok || owner.StartByte != proc.StartByte || owner.StartLine != proc.StartLine || owner.EndLine != proc.EndLine {
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if !scope.shadowsModule(guardName) {
			written = true
		}
	})
	return written
}

func arrayModuleStatementAtLine(proc sourceProcedure, line int) (procedureir.Statement, bool) {
	var found procedureir.Statement
	matched := false
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine != line {
			continue
		}
		if matched {
			return procedureir.Statement{}, false
		}
		found = statement
		matched = true
	}
	return found, matched
}

func arrayModuleSetupReDimIsReliable(file parsedFile, proc sourceProcedure, guardLine, redimLine, setLine int, guardName, name string, constants map[string]int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) bool {
	if redimLine <= guardLine || redimLine >= setLine || redimLine < 0 || setLine >= len(file.Lines) {
		return false
	}
	sourceLine := redimLine + 1
	if sourceLine < 1 || sourceLine > len(file.Lines) || arraySummaryStatementAlwaysFails(normalizedCodeLine(file.Lines[redimLine]), arrayOptionBase(file), constants) {
		return false
	}
	variables := arrayVariables(file, proc, moduleDecls)
	// A call between the allocation and the ready flag can erase or replace the
	// module array without leaving a direct operation fact in this procedure.
	// A resolved private helper is admitted only when its direct and ByRef array
	// effects are known not to invalidate this module's arrays; public,
	// unresolved, and otherwise unmodelled calls remain conservative.
	for call := range proc.Calls.All() {
		line := call.Range.StartLine
		if line >= sourceLine && line < setLine+1 {
			if arrayCallIsIndexedArrayAccess(proc, call, variables) {
				continue
			}
			if call.IsRaiseEvent {
				return false
			}
			if call.Resolution.Status == procedureir.ResolutionBuiltinLike {
				continue
			}
			_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			if !resolved || arrayPrivateCallMayInvalidateModuleArray(file, proc, target, call, moduleDecls, ctx) {
				return false
			}
		}
	}

	if proc.Graph == nil {
		// Compatibility projections may not carry CFGs. Accept only a straight-
		// line source interval in that case; any visible control construct makes
		// the allocation conditional or permits an unmodelled bypass.
		for line := guardLine + 1; line < setLine; line++ {
			if line == redimLine {
				continue
			}
			if arrayModuleSetupLineHasControlFlow(file.Lines[line]) {
				return false
			}
		}
		return true
	}

	findStatement := func(line int, match func(procedureir.Statement) bool) (procedureir.Statement, bool) {
		var found procedureir.Statement
		matched := false
		for statement := range proc.Statements.All() {
			if statement.Range.StartLine != line || !match(statement) {
				continue
			}
			if matched {
				return procedureir.Statement{}, false
			}
			found = statement
			matched = true
		}
		return found, matched
	}
	guardStatement, guardOK := findStatement(guardLine+1, func(statement procedureir.Statement) bool {
		return len(arraySetupGuardRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))) == 2
	})
	redimStatement, redimOK := findStatement(sourceLine, func(statement procedureir.Statement) bool {
		match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
		if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
			return false
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), cleanIdentifier(name)) {
				return true
			}
		}
		return false
	})
	readyStatement, readyOK := findStatement(setLine+1, func(statement procedureir.Statement) bool {
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(statement.Text))
		return ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(guardName)) && strings.EqualFold(strings.TrimSpace(rhs), "true")
	})
	if !guardOK || !redimOK || !readyOK {
		return false
	}
	guardBlock, guardBlockOK := proc.Graph.BlockForStatement(guardStatement.ID)
	redimBlock, redimBlockOK := proc.Graph.BlockForStatement(redimStatement.ID)
	readyBlock, readyBlockOK := proc.Graph.BlockForStatement(readyStatement.ID)
	if !guardBlockOK || !redimBlockOK || !readyBlockOK {
		return false
	}
	normalGraph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	if !normalGraph.IsReachable(redimBlock.ID) || !normalGraph.IsReachable(readyBlock.ID) {
		return false
	}
	dominatesReady := false
	for _, dominator := range normalGraph.DominatorsOf(readyBlock.ID) {
		if dominator == redimBlock.ID {
			dominatesReady = true
			break
		}
	}
	if !dominatesReady || !arrayFalseBranchRequiresBlock(*proc.Graph, guardBlock.ID, redimBlock.ID) {
		return false
	}
	return arrayModuleSetupReachesNormalExit(normalGraph, readyBlock.ID)
}

// arrayCallIsIndexedArrayAccess filters the procedure IR's expression-call
// projection from actual procedure invocations. The VBA tree-sitter grammar
// represents an indexed array expression such as mPow2(index) as a CallSite;
// an unresolved array expression must not make an otherwise straight-line
// setup helper look like it contains an unknown side effect.
func arrayCallIsIndexedArrayAccess(proc sourceProcedure, call procedureir.CallSite, variables map[string]arrayVariable) bool {
	name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
	if name == "" || strings.Contains(name, ".") || call.StatementID <= 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		if statement.ID != call.StatementID {
			continue
		}
		for _, use := range arrayIndexedUses(statement.Text, variables) {
			if strings.EqualFold(cleanIdentifier(use.name), name) {
				return true
			}
		}
		return false
	}
	return false
}

func arrayModuleSetupLineHasControlFlow(line string) bool {
	code := strings.TrimSpace(normalizedCodeLine(line))
	if code == "" {
		return false
	}
	lower := strings.ToLower(code)
	for _, prefix := range []string{
		"if ", "elseif ", "else", "end if", "for ", "for each ", "next", "do", "loop", "while ", "wend",
		"select ", "case ", "goto ", "on error ", "with ", "end with", "exit ",
	} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return strings.HasSuffix(code, ":")
}

func arrayModuleSetupReachesNormalExit(graph vbacfg.CFGView, from vbacfg.BlockID) bool {
	if from == graph.NormalExit() {
		return true
	}
	seen := map[vbacfg.BlockID]bool{from: true}
	queue := []vbacfg.BlockID{from}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		found := false
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if edge.To == graph.NormalExit() {
				found = true
				return false
			}
			if !seen[edge.To] {
				seen[edge.To] = true
				queue = append(queue, edge.To)
			}
			return true
		})
		if found {
			return true
		}
	}
	return false
}

func moduleArrayOperationHasOtherWrite(facts *moduleAnalysisFacts, name string, setupLine int) bool {
	name = strings.ToLower(cleanIdentifier(name))
	otherWrite := false
	facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
		if otherWrite {
			return
		}
		if operation.Kind == moduleArrayDirectRedim && operation.Line == setupLine {
			return
		}
		otherWrite = true
	})
	return otherWrite
}

func arrayProcedureLineHasInlineConditional(file parsedFile, line int) bool {
	if line < 1 || line > len(file.Lines) {
		return false
	}
	text := strings.ToLower(strings.TrimSpace(normalizedCodeLine(file.Lines[line-1])))
	return strings.HasPrefix(text, "if ") && strings.Contains(text, " then ")
}

func arrayProcedureLineInlineConditionIsFalse(file parsedFile, line int) bool {
	if !arrayProcedureLineHasInlineConditional(file, line) {
		return false
	}
	condition, body, ok := arrayIfThenParts(normalizedCodeLine(file.Lines[line-1]))
	if !ok || strings.TrimSpace(body) == "" {
		return false
	}
	condition = strings.TrimSpace(condition)
	lowerCondition := strings.ToLower(condition)
	if strings.HasPrefix(lowerCondition, "if ") {
		condition = strings.TrimSpace(condition[len("if "):])
	}
	value, known := arraySourceOrderConstantBoolean(condition, nil)
	return known && !value
}

func arrayInlineConditionalCallIsReachable(file parsedFile, call procedureir.CallSite, conditionValue, hasElse bool) bool {
	if conditionValue && !hasElse {
		return true
	}
	line := call.Range.StartLine
	if line < 1 || line > len(file.Lines) || call.Range.StartColumn <= 0 {
		return conditionValue || hasElse
	}
	raw := gui.StripComment(file.Lines[line-1])
	trimmed := strings.TrimSpace(raw)
	leading := len(raw) - len(strings.TrimLeft(raw, " \t"))
	_, body, ok := arrayIfThenParts(trimmed)
	if !ok || strings.TrimSpace(body) == "" {
		return true
	}
	prefixLength := 0
	switch {
	case strings.HasPrefix(strings.ToLower(trimmed), "if "):
		prefixLength = len("if ")
	case strings.HasPrefix(strings.ToLower(trimmed), "elseif "):
		prefixLength = len("elseif ")
	default:
		return true
	}
	rest := strings.TrimSpace(trimmed[prefixLength:])
	thenIndex := arrayTopLevelKeywordIndex(rest, "then")
	if thenIndex < 0 {
		return true
	}
	bodyStart := leading + prefixLength + (len(trimmed[prefixLength:]) - len(rest)) + thenIndex + len("then")
	elseIndex := arrayTopLevelKeywordIndex(body, "else")
	if elseIndex < 0 {
		return conditionValue
	}
	elseStart := bodyStart + elseIndex
	callColumn := call.Range.StartColumn - 1
	if callColumn < bodyStart {
		return true
	}
	inThen := callColumn < elseStart
	return inThen == conditionValue
}

func arrayProcedureNormalExitDominators(proc sourceProcedure) map[vbacfg.BlockID]bool {
	if proc.Graph == nil {
		return nil
	}
	dominators := proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[proc.Graph.NormalExit]
	result := make(map[vbacfg.BlockID]bool, len(dominators))
	for _, id := range dominators {
		result[id] = true
	}
	return result
}

func arrayProcedureBlockDominatesNormalExit(proc sourceProcedure, statementID int, dominators map[vbacfg.BlockID]bool) bool {
	if proc.Graph == nil || statementID <= 0 || len(dominators) == 0 {
		return false
	}
	block, ok := proc.Graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	return dominators[block.ID]
}

func arrayProcedureHasIdempotentSetupGuard(file parsedFile, proc sourceProcedure, candidateLine int, moduleDecls map[string]sourceDeclaration) bool {
	if candidateLine <= proc.StartLine {
		return false
	}
	guard := ""
	for line := proc.StartLine; line < candidateLine && line <= len(file.Lines); line++ {
		match := arraySetupGuardRe.FindStringSubmatch(normalizedCodeLine(file.Lines[line-1]))
		if len(match) != 2 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		declaration, ok := moduleDecls[name]
		if !ok || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			continue
		}
		guard = name
	}
	if guard == "" {
		return false
	}
	for line := candidateLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(file.Lines[line-1]))
		if ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), guard) && strings.EqualFold(strings.TrimSpace(rhs), "true") {
			return true
		}
	}
	return false
}

func arrayModuleAllocationSummariesEqual(left, right arrayModuleAllocationSummaries) bool {
	if len(left) != len(right) {
		return false
	}
	for procedure, arrays := range left {
		other, ok := right[procedure]
		if !ok || len(arrays) != len(other) {
			return false
		}
		for name := range arrays {
			if !other[name] {
				return false
			}
		}
	}
	return true
}

func arrayModuleInitializationStates(files []parsedFile, summaries arrayModuleAllocationSummaries) map[string]map[string]bool {
	states := map[string]map[string]bool{}
	for _, file := range files {
		moduleKind := strings.ToLower(strings.TrimSpace(file.ModuleKind))
		if moduleKind != "form" && moduleKind != "class" {
			continue
		}
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		initializer := arrayModuleInitializerName(moduleKind)
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !strings.EqualFold(strings.TrimSpace(proc.Name), initializer) {
				continue
			}
			for name := range summaries[arrayProcedureKey(proc)] {
				if declaration, ok := moduleDecls[name]; ok && declaration.Array && !declaration.Parameter {
					if states[file.Path] == nil {
						states[file.Path] = map[string]bool{}
					}
					states[file.Path][name] = true
				}
			}
		}
	}
	return states
}

func applyArrayModuleInitializationState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, initializationStates map[string]map[string]bool) arrayFlowState {
	if len(initializationStates[file.Path]) == 0 {
		return state
	}
	moduleKind := strings.ToLower(strings.TrimSpace(file.ModuleKind))
	if moduleKind != "form" && moduleKind != "class" {
		return state
	}
	initializer := arrayModuleInitializerName(moduleKind)
	if strings.EqualFold(strings.TrimSpace(proc.Name), initializer) {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range initializationStates[file.Path] {
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, ok := moduleDecls[name]
		variable, known := variables[name]
		if !ok || !known || !declaration.Array || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayModuleInitializerName(moduleKind string) string {
	if moduleKind == "form" {
		return "userform" + "_initialize"
	}
	return "class" + "_initialize"
}

// inferArrayModuleEntryStates propagates a module-array allocation from a
// known caller to a project-local helper. Private procedures are analyzed as
// standalone procedures, so their initial state cannot otherwise reflect an
// assignment performed by the caller. A fact is retained only when every
// resolved call from the same module reaches the helper with that array
// allocated. The fixed point also covers chains of private helpers.
func inferArrayModuleEntryStates(a Analyzer, files []parsedFile, ctx analysisContext) arrayModuleEntryStates {
	if len(ctx.arrayPrivateTargets) == 0 {
		return arrayModuleEntryStates{}
	}

	type procedureInfo struct {
		file        parsedFile
		proc        sourceProcedure
		key         string
		moduleDecls map[string]sourceDeclaration
		variables   map[string]arrayVariable
	}
	procedures := make([]procedureInfo, 0)
	moduleArrays := map[string]map[string]bool{}
	moduleFiles := map[string]string{}
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			key := arrayParticipantLookupKey(proc, ctx.arrayParticipantKeys)
			procedures = append(procedures, procedureInfo{
				file: file, proc: proc, moduleDecls: moduleDecls,
				variables: arrayVariables(file, proc, moduleDecls),
				key:       key,
			})
			moduleArrays[key] = arrayModuleNamesForProcedure(file, proc, moduleDecls)
			moduleFiles[key] = file.Path
		}
	}
	if len(procedures) == 0 {
		return arrayModuleEntryStates{}
	}

	initializationStates := arrayModuleInitializationStates(files, ctx.arrayModuleAllocations)
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	indexByKey := make(map[string]int, len(procedures))
	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		key := procedure.key
		indexByKey[key] = index
		for call := range procedure.proc.Calls.All() {
			_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			participantTargetKey := arrayParticipantLookupKey(target, ctx.arrayParticipantKeys)
			if ok && moduleFiles[participantTargetKey] == procedure.file.Path {
				dependents[participantTargetKey] = append(dependents[participantTargetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}

	evaluate := func(procedure procedureInfo, entries arrayModuleEntryStates) map[string]map[string]bool {
		variables := procedure.variables
		initial := arrayInitialState(variables)
		initial = applyArrayModuleInitializationState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, initializationStates)
		initial = applyArrayModuleReadyGuardState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, ctx.arrayModuleReadyGuards)
		initial = applyArrayModuleEntryState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, entries, ctx.arrayParticipantKeys)
		initial = applyArrayInternalStorageConfiguration(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, ctx.arrayModuleConfigurations[procedure.file.Path])
		candidates := map[string]map[string]bool{}
		recordCall := func(call procedureir.CallSite, state arrayFlowState) {
			_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			key := arrayParticipantLookupKey(target, ctx.arrayParticipantKeys)
			if !ok || moduleFiles[key] != procedure.file.Path {
				return
			}
			names := moduleArrays[key]
			if len(names) == 0 {
				return
			}
			candidate := candidates[key]
			if candidate == nil {
				candidate = cloneArrayNameSet(names)
				candidates[key] = candidate
			}
			for name := range names {
				value, known := state[name]
				if !known || value.kind != arrayAllocated || !value.knownArray {
					candidate[name] = false
				}
			}
		}
		visit := func(text string, line int, in arrayFlowState) arrayFlowState {
			for _, call := range arrayCallsAtLine(procedure.proc.Calls, line) {
				recordCall(call, in)
			}
			out, _ := a.arrayTransfer(procedure.file, procedure.proc, ctx, variables, in, text, line, nil, nil)
			for _, call := range arrayCallsAtLine(procedure.proc.Calls, line) {
				out = applyArrayModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
				out = applyArrayUnknownModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
			}
			return out
		}
		if procedure.proc.Graph == nil {
			state := initial
			for line := procedure.proc.StartLine; line <= procedure.proc.EndLine && line <= len(procedure.file.Lines); line++ {
				state = visit(normalizedCodeLine(procedure.file.Lines[line-1]), line, state)
			}
			return candidates
		}
		graph := arrayVBA227Graph(procedure.proc, ctx)
		if ctx.arrayStats != nil {
			ctx.arrayStats.addCFGWalk()
		}
		walkArrayCFGWithEdgesStats(&graph, procedure.file.Lines, initial, visit, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
			out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[procedure.file.Path], variables, procedure.file, procedure.proc, procedure.moduleDecls)
		}, ctx.arrayStats)
		return candidates
	}

	contributions := make(map[string]map[string]map[string]bool, len(procedures))
	entries := arrayModuleEntryStates{}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		key := procedures[index].key
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		contribution := evaluate(procedures[index], entries)
		if arrayModuleEntryContributionsEqual(contributions[key], contribution) {
			continue
		}
		contributions[key] = contribution
		next := arrayModuleEntryStates{}
		for _, caller := range procedures {
			for target, names := range contributions[caller.key] {
				if next[target] == nil {
					next[target] = cloneArrayNameSet(names)
					continue
				}
				for name := range next[target] {
					if !names[name] {
						delete(next[target], name)
					}
				}
			}
		}
		changedTargetSet := make(map[string]bool, len(entries)+len(next))
		for target := range entries {
			changedTargetSet[target] = true
		}
		for target := range next {
			changedTargetSet[target] = true
		}
		changedTargets := make([]string, 0, len(changedTargetSet))
		for target := range changedTargetSet {
			if !arrayModuleEntryTargetEqual(entries, next, target) {
				changedTargets = append(changedTargets, target)
			}
		}
		if len(changedTargets) == 0 {
			continue
		}
		entries = next
		sortArrayProcedureKeys(changedTargets, indexByKey)
		for _, target := range changedTargets {
			for _, dependent := range dependents[target] {
				if !queued[dependent] {
					queued[dependent] = true
					queue = append(queue, dependent)
				}
			}
			if index, ok := indexByKey[target]; ok && !queued[index] {
				queued[index] = true
				queue = append(queue, index)
			}
		}
	}
	return entries
}

func arrayModuleEntryContributionsEqual(left, right map[string]map[string]bool) bool {
	if len(left) != len(right) {
		return false
	}
	for key, names := range left {
		other, ok := right[key]
		if !ok || len(names) != len(other) {
			return false
		}
		for name, allocated := range names {
			if other[name] != allocated {
				return false
			}
		}
	}
	return true
}

func arrayModuleEntryTargetEqual(left, right arrayModuleEntryStates, target string) bool {
	leftNames, leftOK := left[target]
	rightNames, rightOK := right[target]
	if !leftOK || !rightOK {
		return !leftOK && !rightOK
	}
	if len(leftNames) != len(rightNames) {
		return false
	}
	for name, allocated := range leftNames {
		if rightNames[name] != allocated {
			return false
		}
	}
	return true
}

func arrayModuleNamesForProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]bool {
	moduleArrays := map[string]bool{}
	for name, declaration := range moduleDecls {
		if declaration.Array && !declaration.Parameter {
			moduleArrays[strings.ToLower(name)] = true
		}
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	for name := range moduleArrays {
		if declarations.shadowsModule(name) {
			delete(moduleArrays, name)
		}
	}
	return moduleArrays
}

func applyArrayModuleEntryState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, entries arrayModuleEntryStates, participantKeys ...map[string]string) arrayFlowState {
	var keyIndex map[string]string
	if len(participantKeys) > 0 {
		keyIndex = participantKeys[0]
	}
	allocated := entries[arrayParticipantLookupKey(proc, keyIndex)]
	if len(allocated) == 0 && len(keyIndex) > 0 {
		// Focused compatibility callers may still provide the legacy base key.
		allocated = entries[arrayProcedureKey(proc)]
	}
	if len(allocated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name, isAllocated := range allocated {
		if !isAllocated {
			continue
		}
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || declaration.Parameter || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func applyArrayModuleReadyGuardState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, guards arrayModuleReadyGuardStates) arrayFlowState {
	byGuard := guards[file.Path]
	if len(byGuard) == 0 {
		return state
	}
	guardName, ok := arrayModuleReadyGuardAtEntry(file, proc, moduleDecls)
	if !ok {
		return state
	}
	allocated := byGuard[guardName]
	if len(allocated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range allocated {
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || declaration.Parameter || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayModuleReadyGuardAtEntry(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) (string, bool) {
	if proc.StartLine < 1 || proc.EndLine <= proc.StartLine || proc.StartLine > len(file.Lines) {
		return "", false
	}
	end := min(len(file.Lines), proc.EndLine)
	for line := proc.StartLine + 1; line < end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" || strings.HasPrefix(text, "'") {
			continue
		}
		if declRe.MatchString(text) || strings.HasPrefix(strings.ToLower(text), "on error ") || strings.HasSuffix(text, ":") {
			continue
		}
		match := arrayModuleReadyGuardRe.FindStringSubmatch(text)
		if len(match) != 2 {
			return "", false
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		declaration, declared := moduleDecls[name]
		if !declared || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			return "", false
		}
		return name, true
	}
	return "", false
}

type arrayByRefEntryEvidence struct {
	seen                bool
	allocated           bool
	conditionCompatible bool
	condition           string
}

type arrayByRefCallCandidate struct {
	key    string
	target sourceProcedure
	call   procedureir.CallSite
}

type arrayLocalGoSubSummary struct {
	guaranteedAllocated map[string]bool
	unknown             map[string]bool
}

type arrayLocalGoSubAllocations map[string]arrayLocalGoSubSummary

func arrayLocalGoSubAllocationSummaries(proc sourceProcedure, graph *vbacfg.CFGView, variables map[string]arrayVariable, ctx analysisContext, base int, constants map[string]int, moduleArrays map[string]bool) arrayLocalGoSubAllocations {
	statements := make([]procedureir.Statement, 0, proc.Statements.Len())
	for statement := range proc.Statements.All() {
		statements = append(statements, statement)
	}
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartLine != statements[j].Range.StartLine {
			return statements[i].Range.StartLine < statements[j].Range.StartLine
		}
		return statements[i].ID < statements[j].ID
	})
	summaries := arrayLocalGoSubAllocations{}
	if graph == nil {
		return summaries
	}
	for index, label := range statements {
		if label.Kind != procedureir.StatementLabel {
			continue
		}
		labelName := arrayLocalGoSubLabelName(label)
		if labelName == "" {
			continue
		}
		end := len(statements)
		for cursor := index + 1; cursor < len(statements); cursor++ {
			if arrayLocalGoSubIsReturn(statements[cursor]) {
				end = cursor
				break
			}
		}
		if end == len(statements) {
			continue
		}
		summary := arrayLocalGoSubSummary{
			guaranteedAllocated: map[string]bool{},
			unknown:             map[string]bool{},
		}
		for name, variable := range variables {
			if variable.isArray && arrayLocalGoSubAllocationInvariant(proc, graph, statements, index, end, name, ctx, base, constants, moduleArrays) {
				summary.guaranteedAllocated[name] = true
			} else if variable.isArray && !variable.fixed && arrayLocalGoSubMayMutateName(proc, graph, statements, index, end, name, ctx, moduleArrays) {
				summary.unknown[name] = true
			}
		}
		summaries[labelName] = summary
	}
	return summaries
}

func arrayLocalGoSubAllocationInvariant(proc sourceProcedure, graph *vbacfg.CFGView, statements []procedureir.Statement, labelIndex, end int, name string, ctx analysisContext, base int, constants map[string]int, moduleArrays map[string]bool) bool {
	if graph == nil || labelIndex < 0 || labelIndex >= end || end > len(statements) {
		return false
	}
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	allowed := map[vbacfg.BlockID]procedureir.Statement{}
	var labelBlock vbacfg.Block
	labelFound := false
	for index := labelIndex; index < end; index++ {
		block, ok := graph.BlockForStatement(statements[index].ID)
		if !ok {
			return false
		}
		allowed[block.ID] = statements[index]
		if index == labelIndex {
			labelBlock = block
			labelFound = true
		}
	}
	if !labelFound {
		return false
	}
	returnBlock, returnOK := graph.BlockForStatement(statements[end].ID)
	if !returnOK {
		return false
	}

	type stateAtBlock struct {
		id        vbacfg.BlockID
		allocated bool
	}
	seenStates := map[vbacfg.BlockID]map[bool]bool{labelBlock.ID: {false: true}}
	queue := []stateAtBlock{{id: labelBlock.ID}}
	for len(queue) > 0 {
		currentState := queue[0]
		queue = queue[1:]
		currentID := currentState.id
		if _, ok := graph.BlockByID(currentID); !ok {
			return false
		}
		failed := false
		graph.ForEachOutgoing(currentID, func(edge vbacfg.Edge) bool {
			target, targetOK := graph.BlockByID(edge.To)
			if !targetOK {
				failed = true
				return true
			}
			// An unknown edge can leave the GoSub body through a dynamic or
			// recovered transfer. It is not evidence of a successful Return,
			// even when the predecessor happens to be allocated.
			if edge.Kind == vbacfg.EdgeUnknown || target.Kind == vbacfg.BlockUnknownExit {
				failed = true
				return true
			}
			statement, inside := allowed[target.ID]
			if !inside {
				// The first Return is deliberately outside the allowed body. Any
				// other statement target leaves the GoSub body without returning to
				// its caller, so it cannot be treated as a successful summary.
				if target.ID != returnBlock.ID && !arrayLocalGoSubIsTerminalBlock(target) {
					failed = true
					return true
				}
				if !currentState.allocated {
					failed = true
				}
				// Keep visiting sibling edges. A terminal edge is safe only when
				// every other edge from this block also preserves the summary.
				return true
			}
			nextAllocated := arrayLocalGoSubStateAfterStatement(statement.Text, name, currentState.allocated, ctx, base, constants)
			unknownCall, guaranteedCall := arrayLocalGoSubArrayCallEffect(proc, statement, name, ctx, moduleArrays)
			if unknownCall {
				nextAllocated = false
			} else if guaranteedCall {
				nextAllocated = true
			}
			if !seenStates[target.ID][nextAllocated] {
				if seenStates[target.ID] == nil {
					seenStates[target.ID] = map[bool]bool{}
				}
				seenStates[target.ID][nextAllocated] = true
				queue = append(queue, stateAtBlock{id: target.ID, allocated: nextAllocated})
			}
			return true
		})
		if failed {
			return false
		}
	}
	return true
}

func arrayLocalGoSubStateAfterStatement(text, name string, allocated bool, ctx analysisContext, base int, constants map[string]int) bool {
	names := map[string]bool{name: true}
	for _, statement := range splitRangeValueSourceStatements(text) {
		statement = strings.TrimSpace(statement)
		if statement == "" {
			continue
		}
		if arrayLocalGoSubAllocationIsReliable(statement, name, ctx, base, constants) {
			allocated = true
			continue
		}
		if arrayLocalGoSubPreserveKeepsAllocation(statement, name, allocated, base, constants) {
			continue
		}
		if arraySourceOrderInlineArrayMutation(statement, names, ctx) || arraySourceOrderMutatesArrayStatement(statement, names, ctx) {
			allocated = false
		}
	}
	return allocated
}

func arrayLocalGoSubPreserveKeepsAllocation(text, name string, allocated bool, base int, constants map[string]int) bool {
	if !allocated {
		return false
	}
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 || strings.TrimSpace(match[1]) == "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if direct && strings.EqualFold(cleanIdentifier(redim.name), name) && !arraySummaryStatementAlwaysFails(text, base, constants) {
			return true
		}
	}
	return false
}

func arrayLocalGoSubMayMutateName(proc sourceProcedure, graph *vbacfg.CFGView, statements []procedureir.Statement, labelIndex, end int, name string, ctx analysisContext, moduleArrays map[string]bool) bool {
	if graph == nil || labelIndex < 0 || labelIndex >= end || end > len(statements) {
		return true
	}
	for index := labelIndex; index < end; index++ {
		statement := statements[index]
		for _, part := range splitRangeValueSourceStatements(statement.Text) {
			part = strings.TrimSpace(part)
			if part == "" {
				continue
			}
			if arrayStatementAllocatesName(part, name, ctx) || arraySourceOrderInlineArrayMutation(part, map[string]bool{name: true}, ctx) || arraySourceOrderMutatesArrayStatement(part, map[string]bool{name: true}, ctx) {
				return true
			}
		}
		unknownCall, _ := arrayLocalGoSubArrayCallEffect(proc, statement, name, ctx, moduleArrays)
		if unknownCall {
			return true
		}
	}
	return false
}

// arrayLocalGoSubArrayCallEffect keeps a local GoSub summary fail-closed at
// calls that pass the tracked array directly. Only an already-proven private
// ByRef allocation contract is allowed to establish the post-call state;
// builtin-like calls are treated as non-mutating, while every other call may
// erase or otherwise invalidate the array behind the analyzer's view.
func arrayLocalGoSubArrayCallEffect(proc sourceProcedure, statement procedureir.Statement, name string, ctx analysisContext, moduleArrays map[string]bool) (unknown, guaranteedAllocated bool) {
	name = strings.ToLower(cleanIdentifier(name))
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
			continue
		}
		if call.StatementID != statement.ID && (call.StatementID != 0 || call.Range.StartLine != statement.Range.StartLine) {
			continue
		}
		arguments := arrayCallArgumentTexts(proc, call)
		relevantArgument := false
		key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if resolved {
			bindings, mapped := arrayCallArgumentBindings(proc, target, call)
			if mapped {
				for _, binding := range bindings {
					if directArrayArgumentName(binding.text) != name {
						continue
					}
					if binding.parameterIndex >= 0 && binding.parameterIndex < target.Params.Len() && target.Params.valueAt(binding.parameterIndex).ParamArray {
						continue
					}
					relevantArgument = true
					if binding.parameterIndex < target.Params.Len() && parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) && ctx.arrayByRefAllocations[key][binding.parameterIndex] {
						guaranteedAllocated = true
						continue
					}
					unknown = true
				}
			} else if call.Arguments.Count > 0 {
				// A resolved call with an incomplete argument projection cannot
				// prove which formal parameter receives the tracked array.
				unknown = true
			}
		} else {
			// Keep unresolved calls conservative by checking the raw actual
			// expressions. Named arguments are handled here only because there
			// is no resolved formal signature to bind them to.
			for _, argument := range call.Arguments.Named {
				if directArrayArgumentName(argument.ValueText) == name {
					relevantArgument = true
					unknown = true
				}
			}
			for _, argument := range arguments {
				if directArrayArgumentName(argument) == name {
					relevantArgument = true
					unknown = true
				}
			}
		}
		if call.Arguments.Count > 0 && len(arguments) != call.Arguments.Count {
			// An incomplete argument projection cannot prove that this call did
			// not receive the tracked array by reference.
			unknown = true
		}
		if !relevantArgument && moduleArrays[name] {
			if arrayLocalGoSubCallHasProvenModuleContract(ctx, call, name) {
				guaranteedAllocated = true
			} else {
				unknown = true
			}
		}
	}
	return unknown, guaranteedAllocated
}

func arrayLocalGoSubCallHasProvenModuleContract(ctx analysisContext, call procedureir.CallSite, name string) bool {
	key, _, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return false
	}
	return ctx.arrayModuleAllocations[key][strings.ToLower(cleanIdentifier(name))]
}

func arrayLocalGoSubIsTerminalBlock(block vbacfg.Block) bool {
	switch block.Kind {
	case vbacfg.BlockNormalExit, vbacfg.BlockExceptionalExit, vbacfg.BlockTerminationExit, vbacfg.BlockUnknownExit:
		return true
	default:
		return false
	}
}

func arrayLocalGoSubAllocationIsReliable(text, name string, ctx analysisContext, base int, constants map[string]int) bool {
	if !arrayStatementAllocatesName(text, name, ctx) {
		return false
	}
	return !arraySummaryStatementAlwaysFails(text, base, constants)
}

func arrayLocalGoSubLabelName(statement procedureir.Statement) string {
	label := statement.Label
	if label == "" {
		label = statement.Text
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(label, ":")))
}

func arrayLocalGoSubIsReturn(statement procedureir.Statement) bool {
	return strings.EqualFold(strings.TrimSpace(statement.Text), "return")
}

func arrayLocalGoSubTarget(proc sourceProcedure, call procedureir.CallSite) string {
	statement, ok := arrayProcedureStatementByID(proc, call.StatementID)
	if !ok || !strings.EqualFold(call.Callee.BaseName, "gosub") {
		return ""
	}
	text := strings.TrimSpace(statement.Text)
	if len(text) < len("gosub") || !strings.EqualFold(text[:len("gosub")], "gosub") {
		return ""
	}
	fields := strings.Fields(strings.TrimSpace(text[len("gosub"):]))
	if len(fields) == 0 {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(fields[0], ":")))
}

func applyArrayLocalGoSubEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, summaries arrayLocalGoSubAllocations) arrayFlowState {
	target := arrayLocalGoSubTarget(proc, call)
	if target == "" {
		return state
	}
	return applyArrayLocalGoSubSummary(state, target, summaries)
}

func applyArrayLocalGoSubSummary(state arrayFlowState, target string, summaries arrayLocalGoSubAllocations) arrayFlowState {
	summary, known := summaries[target]
	if !known {
		// A local GoSub can mutate caller/module arrays without carrying an
		// explicit array argument. An absent summary therefore means unknown
		// state, not "no effect"; retaining allocated here would suppress
		// VBA227 after cleanup or another unmodelled side effect.
		updated := cloneArrayState(state)
		for name := range updated {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
		return updated
	}
	updated := cloneArrayState(state)
	for name := range summary.unknown {
		updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
	}
	for name := range summary.guaranteedAllocated {
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func applyArrayLocalGoSubStatementEffects(state arrayFlowState, text string, summaries arrayLocalGoSubAllocations) arrayFlowState {
	for _, statement := range splitRangeValueSourceStatements(text) {
		target, ok := arrayLocalGoSubTargetFromStatementText(statement)
		if !ok {
			continue
		}
		state = applyArrayLocalGoSubSummary(state, target, summaries)
	}
	return state
}

func arrayLocalGoSubTargetFromStatementText(text string) (string, bool) {
	fields := strings.Fields(strings.TrimSpace(text))
	if len(fields) < 2 || !strings.EqualFold(fields[0], "gosub") {
		return "", false
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(fields[1], ":"))), true
}

func inferArrayByRefEntryStates(a Analyzer, files []parsedFile, ctx analysisContext) (map[string]map[int]bool, map[string]map[int]string) {
	targets := ctx.arrayPrivateTargets
	if len(targets) == 0 {
		return map[string]map[int]bool{}, map[string]map[int]string{}
	}
	moduleAllocationSummaries := ctx.arrayModuleAllocations
	moduleInitializationStates := arrayModuleInitializationStates(files, moduleAllocationSummaries)

	type callerInfo struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
		variables   map[string]arrayVariable
		constants   map[string]int
	}
	callers := make([]callerInfo, 0)
	for _, file := range files {
		moduleDecls := file.moduleDecls()
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			caller := procedures.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, caller) {
				continue
			}
			eligibleCaller := false
			for call := range caller.Calls.All() {
				_, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
				if ok && procedureHasByRefArrayParameter(target) && arrayProcedureIsParticipant(ctx, target) {
					eligibleCaller = true
					break
				}
			}
			if !eligibleCaller {
				continue
			}
			callers = append(callers, callerInfo{
				file: file, proc: caller, moduleDecls: moduleDecls,
				variables: arrayVariables(file, caller, moduleDecls),
				constants: arrayIntegerConstants(file, caller, a.visibleConstantValues, a.visibleConstants),
			})
		}
	}
	sort.SliceStable(callers, func(i, j int) bool {
		return arrayProcedureLess(callers[i].proc, callers[j].proc)
	})

	evaluateCaller := func(caller callerInfo, entries map[string]map[int]bool, conditions map[string]map[int]string) map[string]map[int]arrayByRefEntryEvidence {
		file := caller.file
		proc := caller.proc
		moduleDecls := caller.moduleDecls
		variables := caller.variables
		constants := caller.constants
		localCtx := ctx
		localCtx.arrayByRefEntryConditions = conditions
		evidence := map[string]map[int]arrayByRefEntryEvidence{}
		initial := arrayInitialState(variables)
		initial = applyArrayModuleInitializationState(initial, file, proc, variables, moduleDecls, moduleInitializationStates)
		initial = applyArrayModuleReadyGuardState(initial, file, proc, variables, moduleDecls, ctx.arrayModuleReadyGuards)
		initial = applyArrayByRefEntryStates(initial, proc, variables, entries, conditions)
		initial = applyArrayModuleEntryState(initial, file, proc, variables, moduleDecls, ctx.arrayModuleEntryStates, ctx.arrayParticipantKeys)
		initial = applyArrayInternalStorageConfiguration(initial, file, proc, variables, moduleDecls, ctx.arrayModuleConfigurations[file.Path])
		var baseView vbacfg.CFGView
		var summaryGraph *vbacfg.CFGView
		var worklistReachable map[vbacfg.BlockID]bool
		worklistReachableLines := map[int]bool{}
		if proc.Graph != nil {
			baseView = proc.Graph.View(vbacfg.EdgeFilter{})
			summaryGraph = &baseView
			worklistReachable = arrayCFGWorklistReachable(&baseView)
			for statement := range proc.Statements.All() {
				line := statement.Range.StartLine
				owner, ownerOK := baseView.BlockForStatement(statement.ID)
				if line > 0 && ownerOK && worklistReachable[owner.ID] {
					worklistReachableLines[line] = true
				}
			}
		}
		moduleArrays := arrayModuleNamesForProcedure(file, proc, moduleDecls)
		localGoSubAllocations := arrayLocalGoSubAllocationSummaries(proc, summaryGraph, variables, localCtx, arrayOptionBase(file), constants, moduleArrays)
		visitForBlock := func(text string, line int, in arrayFlowState, ownerStatementID int, filterNestedCalls, skipNestedState bool) arrayFlowState {
			if skipNestedState {
				return in
			}
			var eligible []arrayByRefCallCandidate
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				if filterNestedCalls && ownerStatementID > 0 && call.StatementID != ownerStatementID {
					owner, ownerOK := baseView.BlockForStatement(call.StatementID)
					if ownerOK && worklistReachable[owner.ID] {
						continue
					}
				}
				key, target, ok := arrayPrivateTargetForCall(localCtx, targets, call)
				if !ok || !procedureHasByRefArrayParameter(target) || !arrayProcedureIsParticipant(localCtx, target) {
					continue
				}
				eligible = append(eligible, arrayByRefCallCandidate{key: key, target: target, call: call})
			}
			if len(eligible) > 0 {
				allSameTarget := true
				for _, entry := range eligible[1:] {
					if entry.key != eligible[0].key {
						allSameTarget = false
						break
					}
				}
				// Record and apply each call in source order. This matters when
				// one physical line invokes the same ByRef helper repeatedly: the
				// second call must see an Erase or other invalidation from the
				// first call rather than the original pre-line state.
				recordState := cloneArrayState(in)
				for _, entry := range eligible {
					record := allSameTarget
					if !record {
						// Nested calls on one source line are normally kept
						// conservative because the pre-line state cannot describe
						// mutations from an earlier, different helper. An outer
						// ByRef call whose array argument is a proven allocated
						// expression is independent of that ordering, however.
						allProven, hasExpression := arrayByRefCallHasProvenArrayArguments(file, entry.target, proc, entry.call, recordState, localCtx)
						record = allProven && (hasExpression || arrayByRefCallIsInnermostNested(entry.call, eligible))
					}
					if record {
						arrayRecordByRefCall(evidence, entry.key, entry.target, proc, entry.call, file, recordState, localCtx)
					}
					recordState = applyArrayModuleCallEffects(recordState, file, proc, entry.call, localCtx, variables, moduleDecls)
					if arrayProcedureLineHasInlineConditional(file, entry.call.Range.StartLine) {
						recordState = applyArrayConditionalByRefCallEffects(recordState, proc, entry.call, localCtx)
					} else {
						recordState = applyArrayByRefCallEffects(recordState, proc, entry.call, localCtx)
					}
					recordState = applyArrayLocalGoSubEffects(recordState, proc, entry.call, localGoSubAllocations)
				}
			}
			// ByRef entry proofs must use the same logical-line normalization as
			// VBA227 itself; otherwise a continued Split assignment can make a
			// later call-site argument look unallocated.
			out, _ := a.arrayVBA227Transfer(file, proc, localCtx, variables, in, text, line, constants, nil, nil)
			out = applyArrayLocalGoSubStatementEffects(out, text, localGoSubAllocations)
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				if filterNestedCalls && ownerStatementID > 0 && call.StatementID != ownerStatementID {
					// This callback is visiting the container statement's source
					// range. Nested calls are visited again by their own CFG block
					// (or by the source-order fallback when that block is recovered),
					// so applying their post-call effect here would leak one branch
					// into its siblings.
					continue
				}
				out = applyArrayModuleCallEffects(out, file, proc, call, localCtx, variables, moduleDecls)
				out = applyArrayUnknownModuleCallEffects(out, file, proc, call, localCtx, variables, moduleDecls)
				if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
					out = applyArrayConditionalByRefCallEffects(out, proc, call, localCtx)
				} else {
					out = applyArrayByRefCallEffects(out, proc, call, localCtx)
				}
				out = applyArrayLocalGoSubEffects(out, proc, call, localGoSubAllocations)
			}
			return out
		}
		visit := func(text string, line int, in arrayFlowState) arrayFlowState {
			return visitForBlock(text, line, in, 0, false, false)
		}
		visitBlock := func(block vbacfg.Block, text string, line int, in arrayFlowState) arrayFlowState {
			filterNestedCalls := arrayCFGBlockOwnsNestedStatements(block)
			skipNestedState := false
			if filterNestedCalls && block.Statement != nil {
				start := block.Statement.Range.StartLine
				if start == 0 {
					start = block.Range.StartLine
				}
				skipNestedState = line > start && worklistReachableLines[line]
			}
			return visitForBlock(text, line, in, block.StatementID, filterNestedCalls, skipNestedState)
		}
		if proc.Graph == nil {
			state := initial
			for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
				state = visit(normalizedCodeLine(file.Lines[line-1]), line, state)
			}
			return evidence
		}
		fallbackFacts := buildArraySourceOrderFallbackFacts(file, proc, &baseView, variables, localCtx, constants)
		fallbackFacts.unknownFlow = len(proc.Graph.UnknownFlowSources) > 0
		if ctx.arrayStats != nil {
			ctx.arrayStats.addCFGWalk()
		}
		walkArrayCFGWithSourceLinesReliableStatsAndBlock(&baseView, file.Lines, initial, visit, visitBlock, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			out = applyArrayConditionalAllocationBranch(out, &baseView, block, edge)
			out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
		}, nil, ctx.arrayStats)
		// A recovered source construct can make a call block unreachable in the
		// CFG even though the call is valid VBA source.  The parser currently
		// represents some colon-separated single-line statements this way.  For
		// that narrow boundary, retain a call-site allocation proof from the
		// lexical source order; only direct array arguments with an allocation
		// invariant across all reachable branch alternatives are admitted.
		for call := range proc.Calls.All() {
			block, ok := proc.Graph.BlockForStatement(call.StatementID)
			if !ok || worklistReachable[block.ID] {
				continue
			}
			if !arrayByRefSourceOrderFallbackApplies(file, proc, &baseView, fallbackFacts, call) {
				continue
			}
			key, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
			if !ok || !procedureHasByRefArrayParameter(target) || !arrayProcedureIsParticipant(ctx, target) {
				continue
			}
			if state, proven := a.arrayByRefCallSourceOrderProof(file, fallbackFacts, localGoSubAllocations, proc, target, call, initial, ctx, variables, constants); proven {
				arrayRecordByRefCall(evidence, key, target, proc, call, file, state, ctx)
			}
		}
		return evidence
	}

	dependents := make(map[string][]int)
	indexByKey := make(map[string]int, len(callers))
	for index, caller := range callers {
		indexByKey[arrayProcedureKey(caller.proc)] = index
		for call := range caller.proc.Calls.All() {
			if key, target, ok := arrayPrivateTargetForCall(ctx, targets, call); ok && procedureHasByRefArrayParameter(target) && arrayProcedureIsParticipant(ctx, target) {
				dependents[key] = append(dependents[key], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	contributions := make(map[string]map[string]map[int]arrayByRefEntryEvidence, len(callers))
	entries := map[string]map[int]bool{}
	conditions := map[string]map[int]string{}
	queue := make([]int, len(callers))
	queued := make([]bool, len(callers))
	for index := range callers {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(callers) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		caller := callers[index]
		key := arrayProcedureKey(caller.proc)
		evidence := evaluateCaller(caller, entries, conditions)
		if arrayByRefEvidenceMapsEqual(contributions[key], evidence) {
			continue
		}
		contributions[key] = evidence
		merged := map[string]map[int]arrayByRefEntryEvidence{}
		for _, callerEvidence := range contributions {
			mergeArrayByRefEntryEvidence(merged, callerEvidence)
		}
		result := map[string]map[int]bool{}
		conditionalResult := map[string]map[int]string{}
		for targetKey, parameters := range merged {
			for parameterIndex, fact := range parameters {
				if !fact.seen {
					continue
				}
				if fact.allocated {
					if result[targetKey] == nil {
						result[targetKey] = map[int]bool{}
					}
					result[targetKey][parameterIndex] = true
				}
				if !fact.allocated && fact.conditionCompatible && fact.condition != "" {
					if conditionalResult[targetKey] == nil {
						conditionalResult[targetKey] = map[int]string{}
					}
					conditionalResult[targetKey][parameterIndex] = fact.condition
				}
			}
		}
		changedTargets := arrayByRefEntryChangedTargets(entries, conditions, result, conditionalResult)
		if len(changedTargets) == 0 {
			continue
		}
		sortArrayProcedureKeys(changedTargets, indexByKey)
		entries = result
		conditions = conditionalResult
		for _, target := range changedTargets {
			if dependent, ok := indexByKey[target]; ok && !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
			for _, dependent := range dependents[target] {
				if !queued[dependent] {
					queued[dependent] = true
					queue = append(queue, dependent)
				}
			}
		}
	}
	return entries, conditions
}

func arrayByRefEntryStatesEqual(left, right map[string]map[int]bool) bool {
	if len(left) != len(right) {
		return false
	}
	for target, parameters := range left {
		other, ok := right[target]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index := range parameters {
			if !other[index] {
				return false
			}
		}
	}
	return true
}

func arrayByRefEntryConditionsEqual(left, right map[string]map[int]string) bool {
	if len(left) != len(right) {
		return false
	}
	for target, parameters := range left {
		other, ok := right[target]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index, condition := range parameters {
			if other[index] != condition {
				return false
			}
		}
	}
	return true
}

func arrayByRefEvidenceMapsEqual(left, right map[string]map[int]arrayByRefEntryEvidence) bool {
	if len(left) != len(right) {
		return false
	}
	for target, leftParameters := range left {
		rightParameters, ok := right[target]
		if !ok || len(leftParameters) != len(rightParameters) {
			return false
		}
		for index, leftFact := range leftParameters {
			if rightParameters[index] != leftFact {
				return false
			}
		}
	}
	return true
}

func mergeArrayByRefEntryEvidence(dst, src map[string]map[int]arrayByRefEntryEvidence) {
	for target, parameters := range src {
		for index, incoming := range parameters {
			if dst[target] == nil {
				dst[target] = map[int]arrayByRefEntryEvidence{}
			}
			current, exists := dst[target][index]
			if !exists {
				dst[target][index] = incoming
				continue
			}
			current.allocated = current.allocated && incoming.allocated
			if !incoming.allocated && incoming.condition == "" {
				current.conditionCompatible = false
			}
			if incoming.condition != "" {
				if current.condition == "" {
					current.condition = incoming.condition
				} else if !strings.EqualFold(current.condition, incoming.condition) {
					current.conditionCompatible = false
				}
			}
			current.conditionCompatible = current.conditionCompatible && incoming.conditionCompatible
			current.seen = current.seen || incoming.seen
			dst[target][index] = current
		}
	}
}

func arrayByRefEntryChangedTargets(oldEntries map[string]map[int]bool, oldConditions map[string]map[int]string, newEntries map[string]map[int]bool, newConditions map[string]map[int]string) []string {
	keys := map[string]bool{}
	for target := range oldEntries {
		keys[target] = true
	}
	for target := range oldConditions {
		keys[target] = true
	}
	for target := range newEntries {
		keys[target] = true
	}
	for target := range newConditions {
		keys[target] = true
	}
	changed := make([]string, 0, len(keys))
	for target := range keys {
		oldEntry := map[string]map[int]bool{}
		newEntry := map[string]map[int]bool{}
		if names := oldEntries[target]; len(names) > 0 {
			oldEntry[target] = names
		}
		if names := newEntries[target]; len(names) > 0 {
			newEntry[target] = names
		}
		oldCondition := map[string]map[int]string{}
		newCondition := map[string]map[int]string{}
		if conditions := oldConditions[target]; len(conditions) > 0 {
			oldCondition[target] = conditions
		}
		if conditions := newConditions[target]; len(conditions) > 0 {
			newCondition[target] = conditions
		}
		if !arrayByRefEntryStatesEqual(oldEntry, newEntry) || !arrayByRefEntryConditionsEqual(oldCondition, newCondition) {
			changed = append(changed, target)
		}
	}
	return changed
}

func sortArrayProcedureKeys(keys []string, indexByKey map[string]int) {
	sort.SliceStable(keys, func(i, j int) bool {
		left, leftOK := indexByKey[keys[i]]
		right, rightOK := indexByKey[keys[j]]
		if leftOK && rightOK && left != right {
			return left < right
		}
		if leftOK != rightOK {
			return leftOK
		}
		return keys[i] < keys[j]
	})
}

func arrayOptionPrivateModule(lines []string) bool {
	for _, line := range lines {
		if strings.EqualFold(strings.TrimSpace(normalizedCodeLine(line)), "option private module") {
			return true
		}
	}
	return false
}

func arrayCallsAtLine(calls readOnlySpan[procedureir.CallSite], line int) []procedureir.CallSite {
	matched := make([]procedureir.CallSite, 0, 1)
	for call := range calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line {
			continue
		}
		matched = append(matched, call)
	}
	return matched
}

func arrayPrivateTargetForCall(ctx analysisContext, targets map[string]sourceProcedure, call procedureir.CallSite) (string, sourceProcedure, bool) {
	resolution := arrayCallResolution(ctx, call)
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return "", sourceProcedure{}, false
	}
	key := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
	target, ok := targets[key]
	return key, target, ok
}

// arrayCallResolution rechecks the one parser shape that currently loses an
// unparenthesized call's procedure name when its implicit member argument is
// written with a leading dot, for example `Consume .values`.  The ordinary
// resolver remains authoritative; the retry is deliberately limited to a
// single dotted argument so an unresolved object member is not reinterpreted
// as a project-local procedure call.
func arrayCallResolution(ctx analysisContext, call procedureir.CallSite) procedureir.CallResolution {
	resolution := call.Resolution
	if ctx.procedureResolver == nil {
		return resolution
	}
	resolution = ctx.procedureResolver.ResolveCall(call)
	if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
		return resolution
	}
	name := arrayImplicitMemberArgumentCallName(call.Callee.Text)
	if name == "" || strings.EqualFold(name, call.Callee.BaseName) {
		return resolution
	}
	retried := call
	retried.Callee.Text = name
	retried.Callee.BaseName = name
	retried.Callee.Member = name
	retried.Callee.Receiver = nil
	return ctx.procedureResolver.ResolveCall(retried)
}

func arrayImplicitMemberArgumentCallName(text string) string {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "call ") {
		text = strings.TrimSpace(text[len("call "):])
	}
	if text == "" || !isIdentifierStart(text[0]) {
		return ""
	}
	end := 1
	for end < len(text) && isIdentifierPart(text[end]) {
		end++
	}
	rest := strings.TrimSpace(text[end:])
	if len(rest) < 2 || rest[0] != '.' || !isIdentifierStart(rest[1]) {
		return ""
	}
	memberEnd := 2
	for memberEnd < len(rest) && isIdentifierPart(rest[memberEnd]) {
		memberEnd++
	}
	if strings.TrimSpace(rest[memberEnd:]) != "" {
		return ""
	}
	return cleanIdentifier(text[:end])
}

func procedureHasByRefArrayParameter(proc sourceProcedure) bool {
	for parameter := range proc.Params.All() {
		if parameterIsByRefArray(parameter) {
			return true
		}
	}
	return false
}

func arrayRecordByRefCall(evidence map[string]map[int]arrayByRefEntryEvidence, targetKey string, target, caller sourceProcedure, call procedureir.CallSite, file parsedFile, state arrayFlowState, ctx analysisContext) {
	// A self-recursive ByRef helper preserves the entry array state supplied by
	// its caller. Treating the recursive edge as an independent unknown entry
	// would poison the evidence from the allocated external call and keep the
	// callee conservative forever.
	if strings.EqualFold(targetKey, arrayProcedureKey(caller)) {
		return
	}
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return
	}
	arguments := make([]string, target.Params.Len())
	bound := make(map[int]bool, len(bindings))
	for _, binding := range bindings {
		arguments[binding.parameterIndex] = binding.text
		bound[binding.parameterIndex] = true
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) || !bound[index] {
			continue
		}
		name := ""
		name = directArrayArgumentName(arguments[index])
		value, known := state[name]
		allocated := known && value.kind == arrayAllocated && value.knownArray
		if !allocated && arrayQualifiedArgumentProvenAllocated(file, caller, call, arguments[index], ctx) {
			value = arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}
			known = true
			allocated = true
		}
		_, _, qualifiedMember := arrayQualifiedMemberParts(arguments[index])
		if !allocated && !qualifiedMember && index < len(arguments) {
			// A function returning a dynamic array is a valid ByRef array
			// argument in VBA.  The identifier-only path above cannot attach
			// that expression to a caller state entry, so consult the existing
			// array-return summaries before treating the callee entry as
			// unallocated.  Unknown or conditionally allocated returns remain
			// conservative because arrayExpressionState returns no allocated
			// proof for them.
			if returned, returnedKnown := arrayExpressionState(arguments[index], state, ctx); returnedKnown && returned.kind == arrayAllocated && returned.knownArray {
				value = returned
				known = true
				allocated = true
			}
		}
		condition := ""
		if known && !allocated && value.allocationCountSource != "" {
			condition = arrayConditionalEntrySource(target, arguments, index, value.allocationCountSource)
		}
		if known && !allocated && condition == "" && arrayByRefCallArrayVacuouslyUnused(target, index, arguments) {
			// A call may intentionally pass an unallocated optional array to a
			// helper whose only uses are behind a false literal guard. It must not
			// poison the conditional-entry evidence collected from calls that do
			// reach those uses.
			continue
		}
		parameters := evidence[targetKey]
		if parameters == nil {
			parameters = map[int]arrayByRefEntryEvidence{}
		}
		fact := parameters[index]
		if !fact.seen {
			fact.allocated = allocated
			fact.conditionCompatible = allocated || condition != ""
			fact.condition = condition
		} else {
			fact.allocated = fact.allocated && allocated
			if !allocated && condition == "" {
				fact.conditionCompatible = false
			}
			if condition != "" {
				if fact.condition == "" {
					fact.condition = condition
				} else if !strings.EqualFold(fact.condition, condition) {
					fact.conditionCompatible = false
				}
			}
		}
		fact.seen = true
		parameters[index] = fact
		evidence[targetKey] = parameters
	}
}

type arraySourceOrderAllocation struct {
	line     int
	parentID int
}

type arraySourceOrderFallbackFacts struct {
	conditionalTransferLines   []int
	unconditionalTransferLines []int
	definiteExitLines          []int
	unknownFlow                bool
	parents                    map[int]procedureir.Statement
	allocations                map[string][]arraySourceOrderAllocation
	bypassTargetMin            map[int]int
	branchGroups               map[int]map[int]bool
	branchTransferBypass       map[int]map[string]int
	ambiguousTransferLines     map[int]bool
}

// buildArraySourceOrderFallbackFacts materializes the source-order facts once
// per caller. The fallback is only used for recovered CFG boundaries, but a
// caller can contain many such calls; keeping the statement and CFG scans out
// of the per-call proof avoids multiplying the recovery cost by call count.
func buildArraySourceOrderFallbackFacts(file parsedFile, proc sourceProcedure, graph *vbacfg.CFGView, variables map[string]arrayVariable, ctx analysisContext, constants map[string]int) arraySourceOrderFallbackFacts {
	facts := arraySourceOrderFallbackFacts{
		parents:                make(map[int]procedureir.Statement, proc.Statements.Len()),
		allocations:            map[string][]arraySourceOrderAllocation{},
		bypassTargetMin:        map[int]int{},
		branchGroups:           map[int]map[int]bool{},
		branchTransferBypass:   map[int]map[string]int{},
		ambiguousTransferLines: map[int]bool{},
	}
	if graph == nil {
		return facts
	}
	worklistReachable := arrayCFGWorklistReachable(graph)
	statements := make([]procedureir.Statement, 0, proc.Statements.Len())
	statementsByLine := map[int][]procedureir.Statement{}
	for statement := range proc.Statements.All() {
		statements = append(statements, statement)
		facts.parents[statement.ID] = statement
		line := statement.Range.StartLine
		if line > 0 {
			statementsByLine[line] = append(statementsByLine[line], statement)
		}
	}
	for _, statement := range statements {
		if statement.Kind != procedureir.StatementIf {
			continue
		}
		// A nested multi-line If owns an independent branch group even though
		// the IR links it to its containing If through ParentID. Register every
		// If root so allocations in all nested branches participate in the
		// source-order proof.
		facts.branchGroups[statement.ID] = map[int]bool{statement.ID: true}
	}
	hasElse := map[int]bool{}
	for _, statement := range statements {
		switch statement.Kind {
		case procedureir.StatementElseIf:
			if root, ok := arraySourceOrderIfRoot(facts.parents, statement.ID); ok {
				if branches := facts.branchGroups[root]; branches != nil {
					branches[statement.ID] = true
				}
			}
		case procedureir.StatementElse:
			if branches := facts.branchGroups[statement.ParentID]; branches != nil {
				branches[statement.ID] = true
				hasElse[statement.ParentID] = true
			}
		}
	}
	for root, branches := range facts.branchGroups {
		if !hasElse[root] {
			// An If/ElseIf chain without a final Else has an implicit path
			// that executes none of its branch bodies.
			branches[0] = true
		}
	}

	conditionalTransferLines := map[int]bool{}
	unconditionalTransferLines := map[int]bool{}
	definiteExitLines := map[int]bool{}
	for _, statement := range statements {
		line := statement.Range.StartLine
		if line <= proc.StartLine {
			continue
		}
		if statement.Kind == procedureir.StatementIf && statement.SyntaxKind == "single_line_if_statement" {
			if line >= 1 && line <= len(file.Lines) && arraySourceOrderInlineConditionalDefinitelyTerminates(normalizedCodeLine(file.Lines[line-1]), constants) {
				if block, ok := graph.BlockForStatement(statement.ID); ok && worklistReachable[block.ID] {
					definiteExitLines[line] = true
				}
			}
			for _, candidate := range statementsByLine[line] {
				if !arraySourceOrderProcedureTransfer(candidate) || !arraySourceOrderInlineConditionalTransfer(candidate, file.Lines) {
					continue
				}
				block, ok := graph.BlockForStatement(candidate.ID)
				if ok && worklistReachable[block.ID] {
					conditionalTransferLines[line] = true
					break
				}
			}
		}
		if arraySourceOrderProcedureTransfer(statement) && !arraySourceOrderInlineConditionalTransfer(statement, file.Lines) {
			unconditionalTransferLines[line] = true
		}
	}
	for line := range conditionalTransferLines {
		facts.conditionalTransferLines = append(facts.conditionalTransferLines, line)
	}
	for line := range unconditionalTransferLines {
		facts.unconditionalTransferLines = append(facts.unconditionalTransferLines, line)
	}
	for line := range definiteExitLines {
		facts.definiteExitLines = append(facts.definiteExitLines, line)
	}
	sort.Ints(facts.conditionalTransferLines)
	sort.Ints(facts.unconditionalTransferLines)
	sort.Ints(facts.definiteExitLines)

	allocationLines := map[int]bool{}
	for _, statement := range statements {
		line := statement.Range.StartLine
		if line < 1 {
			continue
		}
		block, ok := graph.BlockForStatement(statement.ID)
		if !ok || !worklistReachable[block.ID] {
			continue
		}
		for name, variable := range variables {
			if !variable.isArray || !arrayStatementAllocatesName(statement.Text, name, ctx) {
				continue
			}
			name = strings.ToLower(cleanIdentifier(name))
			facts.allocations[name] = append(facts.allocations[name], arraySourceOrderAllocation{line: line, parentID: statement.ParentID})
			allocationLines[line] = true
		}
	}
	orderedAllocationLines := make([]int, 0, len(allocationLines))
	for line := range allocationLines {
		orderedAllocationLines = append(orderedAllocationLines, line)
	}
	sort.Ints(orderedAllocationLines)
	branchAllocationLines := map[int]map[int]map[string][]int{}
	for name, allocations := range facts.allocations {
		for _, allocation := range allocations {
			root, branch, ok := facts.branchForAllocation(allocation.parentID)
			if !ok {
				continue
			}
			branches := branchAllocationLines[root]
			if branches == nil {
				branches = map[int]map[string][]int{}
			}
			names := branches[branch]
			if names == nil {
				names = map[string][]int{}
			}
			names[name] = append(names[name], allocation.line)
			branches[branch] = names
			branchAllocationLines[root] = branches
		}
	}
	for _, branches := range branchAllocationLines {
		for _, names := range branches {
			for name := range names {
				sort.Ints(names[name])
			}
		}
	}
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		segments := splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1]))
		hasTransfer := false
		hasAllocation := false
		for _, segment := range segments {
			if arraySourceOrderTextHasProcedureTransfer(segment.text) {
				hasTransfer = true
			}
			for name, variable := range variables {
				if variable.isArray && arrayStatementAllocatesName(segment.text, name, ctx) {
					hasAllocation = true
					break
				}
			}
			if hasTransfer && hasAllocation {
				facts.ambiguousTransferLines[line] = true
				break
			}
		}
	}

	// A source block with an edge over an allocation proves that the allocation
	// is not unconditional for calls at or after the edge target. Compute the
	// earliest such target for every candidate allocation in one graph pass;
	// individual calls then use a constant-time threshold check.
	graph.ForEachBlock(func(block vbacfg.Block) bool {
		if block.Kind != vbacfg.BlockStatement || !worklistReachable[block.ID] {
			return true
		}
		sourceLine := block.Range.StartLine
		if sourceLine <= proc.StartLine {
			return true
		}
		graph.ForEachOutgoing(block.ID, func(edge vbacfg.Edge) bool {
			target, ok := graph.BlockByID(edge.To)
			if !ok || target.Kind != vbacfg.BlockStatement {
				if block.Statement != nil && arraySourceOrderProcedureTransfer(*block.Statement) {
					facts.recordBranchTransferBypass(block, target, sourceLine, branchAllocationLines)
				}
				return true
			}
			targetLine := target.Range.StartLine
			for _, allocationLine := range orderedAllocationLines {
				if allocationLine <= sourceLine {
					continue
				}
				if allocationLine >= targetLine {
					break
				}
				if current, exists := facts.bypassTargetMin[allocationLine]; !exists || targetLine < current {
					facts.bypassTargetMin[allocationLine] = targetLine
				}
			}
			if block.Statement != nil && arraySourceOrderProcedureTransfer(*block.Statement) {
				facts.recordBranchTransferBypass(block, target, sourceLine, branchAllocationLines)
			}
			return true
		})
		return true
	})
	return facts
}

func (facts arraySourceOrderFallbackFacts) hasConditionalTransferBefore(line int) bool {
	index := sort.SearchInts(facts.conditionalTransferLines, line)
	return index > 0
}

func (facts arraySourceOrderFallbackFacts) hasDefiniteExitBefore(line int) bool {
	index := sort.SearchInts(facts.definiteExitLines, line)
	return index > 0
}

func (facts arraySourceOrderFallbackFacts) hasUnconditionalTransfer(afterLine, beforeLine int) bool {
	index := sort.Search(len(facts.unconditionalTransferLines), func(index int) bool {
		return facts.unconditionalTransferLines[index] > afterLine
	})
	return index < len(facts.unconditionalTransferLines) && facts.unconditionalTransferLines[index] < beforeLine
}

func (facts arraySourceOrderFallbackFacts) allocationDominatesCall(allocationLine, callLine int) bool {
	targetLine, bypassed := facts.bypassTargetMin[allocationLine]
	return !bypassed || targetLine > callLine
}

func (facts arraySourceOrderFallbackFacts) hasAmbiguousTransferBefore(line int) bool {
	for transferLine := range facts.ambiguousTransferLines {
		if transferLine < line {
			return true
		}
	}
	return false
}

func (facts arraySourceOrderFallbackFacts) allocationInvariant(name string, beforeLine int) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	if facts.hasAmbiguousTransferBefore(beforeLine) {
		return false
	}
	unconditional := false
	branchAllocations := map[int]map[int]bool{}
	for _, allocation := range facts.allocations[name] {
		if allocation.line < 1 || allocation.line >= beforeLine || facts.hasUnconditionalTransfer(allocation.line, beforeLine) {
			continue
		}
		if allocation.parentID == 0 {
			if facts.allocationDominatesCall(allocation.line, beforeLine) {
				unconditional = true
			}
			continue
		}
		root, branch, ok := facts.branchForAllocation(allocation.parentID)
		if !ok {
			continue
		}
		branches := branchAllocations[root]
		if branches == nil {
			branches = map[int]bool{}
		}
		branches[branch] = true
		branchAllocations[root] = branches
	}
	if unconditional {
		return true
	}
	for root := range facts.branchGroups {
		if facts.branchTransferBypassBefore(root, name, beforeLine) {
			return false
		}
	}
	requiredGroup := false
	for root, expected := range facts.branchGroups {
		observed := branchAllocations[root]
		if len(observed) == 0 {
			continue
		}
		requiredGroup = true
		for branch := range expected {
			if !observed[branch] {
				return false
			}
		}
	}
	return requiredGroup
}

func arraySourceOrderIfRoot(parents map[int]procedureir.Statement, statementID int) (int, bool) {
	seen := map[int]bool{}
	for statementID > 0 && !seen[statementID] {
		seen[statementID] = true
		statement, ok := parents[statementID]
		if !ok {
			return 0, false
		}
		switch statement.Kind {
		case procedureir.StatementIf:
			return statement.ID, true
		case procedureir.StatementElseIf:
			statementID = statement.ParentID
		default:
			return 0, false
		}
	}
	return 0, false
}

func (facts arraySourceOrderFallbackFacts) branchForAllocation(parentID int) (int, int, bool) {
	parent, ok := facts.parents[parentID]
	if !ok {
		return 0, 0, false
	}
	switch parent.Kind {
	case procedureir.StatementIf, procedureir.StatementElseIf:
		root, ok := arraySourceOrderIfRoot(facts.parents, parent.ID)
		return root, parent.ID, ok
	case procedureir.StatementElse:
		root, ok := arraySourceOrderIfRoot(facts.parents, parent.ParentID)
		return root, parent.ID, ok
	default:
		return 0, 0, false
	}
}

func (facts arraySourceOrderFallbackFacts) branchForStatement(statementID int) (int, int, bool) {
	seen := map[int]bool{}
	statement, ok := facts.parents[statementID]
	if !ok {
		return 0, 0, false
	}
	parentID := statement.ParentID
	for parentID > 0 && !seen[parentID] {
		seen[parentID] = true
		root, branch, branchOK := facts.branchForAllocation(parentID)
		if branchOK && facts.branchGroups[root] != nil {
			return root, branch, true
		}
		parent, parentOK := facts.parents[parentID]
		if !parentOK {
			break
		}
		parentID = parent.ParentID
	}
	return 0, 0, false
}

func (facts *arraySourceOrderFallbackFacts) recordBranchTransferBypass(block, target vbacfg.Block, sourceLine int, branchAllocationLines map[int]map[int]map[string][]int) {
	if block.Statement == nil || !arraySourceOrderProcedureTransfer(*block.Statement) {
		return
	}
	targetLine := target.Range.StartLine
	if target.Statement != nil && target.Statement.Range.StartLine > 0 {
		targetLine = target.Statement.Range.StartLine
	}
	sourceRoot, sourceBranch, sourceInBranch := facts.branchForStatement(block.Statement.ID)
	for root := range facts.branchGroups {
		rootStatement := facts.parents[root]
		if sourceLine > rootStatement.Range.EndLine {
			continue
		}
		targetAfterGroup := target.Kind != vbacfg.BlockStatement || targetLine <= 0 || targetLine > rootStatement.Range.EndLine
		if !targetAfterGroup {
			continue
		}
		if sourceLine < rootStatement.Range.StartLine || !sourceInBranch || sourceRoot != root {
			for _, names := range branchAllocationLines[root] {
				for name := range names {
					facts.saveBranchTransferBypass(root, name, sourceLine)
				}
			}
			continue
		}
		for name, lines := range branchAllocationLines[root][sourceBranch] {
			if len(lines) == 0 || lines[0] >= sourceLine {
				facts.saveBranchTransferBypass(root, name, sourceLine)
			}
		}
	}
}

func (facts *arraySourceOrderFallbackFacts) saveBranchTransferBypass(root int, name string, line int) {
	byName := facts.branchTransferBypass[root]
	if byName == nil {
		byName = map[string]int{}
		facts.branchTransferBypass[root] = byName
	}
	if previous, exists := byName[name]; !exists || line < previous {
		byName[name] = line
	}
}

func (facts arraySourceOrderFallbackFacts) branchTransferBypassBefore(root int, name string, beforeLine int) bool {
	line, ok := facts.branchTransferBypass[root][name]
	return ok && line < beforeLine
}

func arrayByRefSourceOrderFallbackApplies(file parsedFile, proc sourceProcedure, graph *vbacfg.CFGView, facts arraySourceOrderFallbackFacts, call procedureir.CallSite) bool {
	if graph == nil || facts.unknownFlow || call.Range.StartLine <= proc.StartLine || !arraySourceOrderCallLineIsSingleStatement(file.Lines, call.Range.StartLine) {
		return false
	}
	if facts.hasDefiniteExitBefore(call.Range.StartLine) {
		return false
	}
	return facts.hasConditionalTransferBefore(call.Range.StartLine)
}

func arraySourceOrderCallLineIsSingleStatement(lines []string, line int) bool {
	if line < 1 || line > len(lines) {
		return false
	}
	// The source-order fallback intentionally does not interpret statement
	// order within a colon-separated line.  Rejecting those calls keeps a
	// preceding Erase/assignment from being skipped when the recovered CFG
	// maps the whole line to one statement.
	code := strings.TrimSpace(normalizedCodeLine(arraySourceOrderStripComment(lines[line-1])))
	// VBA permits a Rem comment after a statement separator. The separator is
	// part of the comment boundary, not a second executable statement, so do
	// not let colons inside that comment disable the fallback.
	if strings.HasSuffix(code, ":") {
		code = strings.TrimSpace(strings.TrimSuffix(code, ":"))
	}
	return !arraySourceOrderHasStatementSeparator(code)
}

func arraySourceOrderHasStatementSeparator(code string) bool {
	inString := false
	for index := 0; index < len(code); index++ {
		switch code[index] {
		case '"':
			if inString && index+1 < len(code) && code[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case ':':
			if inString {
				continue
			}
			next := index + 1
			for next < len(code) && (code[next] == ' ' || code[next] == '\t') {
				next++
			}
			if next < len(code) && code[next] == '=' {
				continue
			}
			return true
		}
	}
	return false
}

func arraySourceOrderTextHasProcedureTransfer(text string) bool {
	text = strings.TrimSpace(text)
	if _, body, ok := arrayIfThenParts(text); ok {
		for _, part := range splitRangeValueSourceStatements(body) {
			if arraySourceOrderTextHasProcedureTransfer(part) {
				return true
			}
		}
		return false
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "goto ") {
		return true
	}
	if lower == "end" || lower == "end sub" || lower == "end function" || lower == "end property" {
		return true
	}
	for _, prefix := range []string{"exit sub", "exit function", "exit property"} {
		if lower == prefix || strings.HasPrefix(lower, prefix+" ") {
			return true
		}
	}
	return false
}

func arraySourceOrderProcedureTransfer(statement procedureir.Statement) bool {
	switch statement.Kind {
	case procedureir.StatementGoTo:
		_, isGoSub := arrayLocalGoSubTargetFromStatementText(statement.Text)
		return !isGoSub
	case procedureir.StatementEnd:
		return true
	case procedureir.StatementExit:
		if statement.Control == nil {
			lower := strings.ToLower(strings.TrimSpace(statement.Text))
			return strings.HasPrefix(lower, "exit sub") || strings.HasPrefix(lower, "exit function") || strings.HasPrefix(lower, "exit property")
		}
		switch statement.Control.Transfer {
		case procedureir.TransferExitSub, procedureir.TransferExitFunction, procedureir.TransferExitProperty, procedureir.TransferTerminate:
			return true
		default:
			return false
		}
	default:
		return false
	}
}

func arraySourceOrderInlineConditionalTransfer(statement procedureir.Statement, lines []string) bool {
	line := statement.Range.StartLine
	if line < 1 || line > len(lines) {
		return false
	}
	text := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[line-1])))
	if !strings.HasPrefix(text, "if ") {
		return false
	}
	then := strings.Index(text, " then")
	if then < 0 {
		return false
	}
	suffix := strings.TrimSpace(text[then+len(" then"):])
	if suffix == "" {
		return false
	}
	switch statement.Kind {
	case procedureir.StatementGoTo:
		return strings.Contains(suffix, "goto ")
	case procedureir.StatementEnd:
		return strings.HasPrefix(suffix, "end")
	case procedureir.StatementExit:
		return strings.Contains(suffix, "exit sub") || strings.Contains(suffix, "exit function") || strings.Contains(suffix, "exit property")
	default:
		return false
	}
}

func arraySourceOrderInlineConditionalDefinitelyTerminates(text string, constants map[string]int) bool {
	condition, body, ok := arrayIfThenParts(text)
	if !ok || strings.TrimSpace(body) == "" {
		return false
	}
	condition = strings.TrimSpace(condition)
	lowerCondition := strings.ToLower(condition)
	switch {
	case strings.HasPrefix(lowerCondition, "if "):
		condition = strings.TrimSpace(condition[len("if "):])
	case strings.HasPrefix(lowerCondition, "elseif "):
		condition = strings.TrimSpace(condition[len("elseif "):])
	default:
		return false
	}
	value, known := arraySourceOrderConstantBoolean(condition, constants)
	if !known {
		return false
	}
	thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
	selected := thenBody
	if !value {
		if !hasElse {
			return false
		}
		selected = elseBody
	}
	return arraySourceOrderInlineBodyDefinitelyTerminates(selected, constants)
}

func arraySourceOrderInlineBodyDefinitelyTerminates(text string, constants map[string]int) bool {
	for _, statement := range splitRangeValueSourceStatements(text) {
		statement = strings.TrimSpace(statement)
		if statement == "" {
			continue
		}
		if arraySourceOrderProcedureExitStatementText(statement) || arraySourceOrderInlineConditionalDefinitelyTerminates(statement, constants) {
			return true
		}
	}
	return false
}

func arraySourceOrderProcedureExitStatementText(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{"exit sub", "exit function", "exit property"} {
		if lower == prefix || strings.HasPrefix(lower, prefix+" ") {
			return true
		}
	}
	switch lower {
	case "end", "end sub", "end function", "end property":
		return true
	default:
		return false
	}
}

func arraySourceOrderConstantBoolean(expression string, constants map[string]int) (bool, bool) {
	values := make(map[string]constexpr.Value, len(constants))
	for name, value := range constants {
		values[name] = constexpr.Value{Kind: constexpr.ValueLongLong, Integer: int64(value)}
	}
	result := constexpr.Evaluate(expression, constexpr.NewValues(values))
	if result.Kind != constexpr.Known || result.Typed.Kind != constexpr.ValueBoolean {
		return false, false
	}
	return result.Typed.Boolean, true
}

func arraySourceOrderInlineArrayMutation(text string, names map[string]bool, ctx analysisContext) bool {
	condition, body, ok := arrayIfThenParts(text)
	if !ok || strings.TrimSpace(body) == "" {
		return false
	}
	condition = strings.TrimSpace(condition)
	lowerCondition := strings.ToLower(condition)
	switch {
	case strings.HasPrefix(lowerCondition, "if "):
		condition = strings.TrimSpace(condition[len("if "):])
	case strings.HasPrefix(lowerCondition, "elseif "):
		condition = strings.TrimSpace(condition[len("elseif "):])
	}
	thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
	branches := []string{thenBody}
	if hasElse {
		branches = append(branches, elseBody)
	}
	if value, known := arraySourceOrderConstantBoolean(condition, nil); known {
		if value {
			branches = []string{thenBody}
		} else if hasElse {
			branches = []string{elseBody}
		} else {
			branches = nil
		}
	}
	for _, branch := range branches {
		for _, statement := range splitRangeValueSourceStatements(branch) {
			statement = strings.TrimSpace(statement)
			if statement == "" {
				continue
			}
			if arraySourceOrderInlineArrayMutation(statement, names, ctx) {
				return true
			}
			if arraySourceOrderMutatesArrayStatement(statement, names, ctx) {
				return true
			}
		}
	}
	return false
}

func arraySourceOrderMutatesArrayStatement(text string, names map[string]bool, ctx analysisContext) bool {
	text = strings.TrimSpace(text)
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
		for _, target := range splitArgs(match[1]) {
			if names[strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))] {
				return true
			}
		}
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		if strings.TrimSpace(match[1]) != "" {
			// ReDim Preserve keeps an allocated input array allocated. The
			// source-order state solver handles the separate requirement that
			// the input must already be allocated.
			return false
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			name := strings.ToLower(cleanIdentifier(redim.name))
			if direct && names[name] && !arrayStatementAllocatesName(text, name, ctx) {
				return true
			}
		}
	}
	if lhs, _, indexed, ok := arrayAssignment(text); ok && !indexed {
		name := strings.ToLower(cleanIdentifier(lhs))
		return names[name] && !arrayStatementAllocatesName(text, name, ctx)
	}
	return false
}

// arraySourceOrderStripComment removes VBA apostrophe and Rem comments while
// preserving the byte offsets before the comment. The latter matters because
// recovered CallSite ranges are mapped back to colon-separated source
// segments below.
func arraySourceOrderStripComment(line string) string {
	inString := false
	for index := 0; index < len(line); index++ {
		switch line[index] {
		case '"':
			if inString && index+1 < len(line) && line[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case '\'':
			if !inString {
				return line[:index]
			}
		default:
			if !inString && arraySourceOrderRemCommentAt(line, index) {
				return line[:index]
			}
		}
	}
	return line
}

func arraySourceOrderRemCommentAt(line string, index int) bool {
	if index+3 > len(line) || !strings.EqualFold(line[index:index+3], "Rem") {
		return false
	}
	if index > 0 {
		previous := index - 1
		for previous >= 0 {
			switch line[previous] {
			case ' ', '\t', '\r', '\n':
				previous--
				continue
			}
			break
		}
		if previous >= 0 && line[previous] != ':' {
			return false
		}
	}
	if index+3 < len(line) {
		next := line[index+3]
		if next != ':' && next != ' ' && next != '\t' && next != '\r' && next != '\n' {
			return false
		}
	}
	return true
}

func arraySourceOrderLineStartByte(source []byte, line int) int {
	if line <= 1 {
		return 0
	}
	current := 1
	for index, value := range source {
		if value != '\n' {
			continue
		}
		current++
		if current == line {
			return index + 1
		}
	}
	return -1
}

func arraySourceOrderCallsBySegment(file parsedFile, line int, calls []procedureir.CallSite) ([][]procedureir.CallSite, []procedureir.CallSite) {
	if line < 1 || line > len(file.Lines) {
		return nil, calls
	}
	segments := splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1]))
	if len(segments) == 0 {
		return nil, calls
	}
	bySegment := make([][]procedureir.CallSite, len(segments))
	unassigned := make([]procedureir.CallSite, 0)
	lineStart := arraySourceOrderLineStartByte(file.Source, line)
	for _, call := range calls {
		segmentIndex := -1
		if lineStart >= 0 {
			relative := call.Range.StartByte - lineStart
			for index, segment := range segments {
				if relative >= segment.start && (relative < segment.end || index == len(segments)-1 && relative <= segment.end) {
					segmentIndex = index
					break
				}
			}
		}
		if segmentIndex < 0 {
			unassigned = append(unassigned, call)
			continue
		}
		bySegment[segmentIndex] = append(bySegment[segmentIndex], call)
	}
	return bySegment, unassigned
}

func (a Analyzer) arrayByRefCallSourceOrderProof(file parsedFile, facts arraySourceOrderFallbackFacts, localGoSubAllocations arrayLocalGoSubAllocations, caller, target sourceProcedure, call procedureir.CallSite, initial arrayFlowState, ctx analysisContext, variables map[string]arrayVariable, constants map[string]int) (arrayFlowState, bool) {
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return nil, false
	}
	arguments := make([]string, target.Params.Len())
	bound := make(map[int]bool, len(bindings))
	for _, binding := range bindings {
		arguments[binding.parameterIndex] = binding.text
		bound[binding.parameterIndex] = true
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) || !bound[index] {
			continue
		}
		if directArrayArgumentName(arguments[index]) == "" && !arrayQualifiedArgumentProvenAllocated(file, caller, call, arguments[index], ctx) {
			return nil, false
		}
	}
	arrayNames := map[string]bool{}
	initiallyAllocated := map[string]bool{}
	for index, parameter := range target.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) && index < len(arguments) {
			name := strings.ToLower(cleanIdentifier(directArrayArgumentName(arguments[index])))
			arrayNames[name] = true
			value, known := initial[name]
			initiallyAllocated[name] = known && value.kind == arrayAllocated && value.knownArray
		}
	}
	if call.Range.StartLine <= caller.StartLine || !arraySourceOrderCallLineIsSingleStatement(file.Lines, call.Range.StartLine) {
		return nil, false
	}
	state := cloneArrayState(initial)
	// Recovered CFG blocks can coalesce a colon-separated physical line. Apply
	// each logical segment in source order so Erase or another mutation before a
	// later call cannot be hidden by passing the whole physical line to transfer.
	for line := caller.StartLine; line < call.Range.StartLine && line <= len(file.Lines); line++ {
		segments := splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1]))
		callsBySegment, unassignedCalls := arraySourceOrderCallsBySegment(file, line, arrayCallsAtLine(caller.Calls, line))
		if len(unassignedCalls) > 0 && len(segments) > 1 {
			// A call without a trustworthy source offset cannot be placed among
			// colon-separated statements. Continuing would apply its side effect
			// after the whole line and could silently reverse an allocation and a
			// later Erase. The fallback is intentionally fail-closed here.
			return nil, false
		}
		for segmentIndex, segment := range segments {
			statement := normalizedCodeLine(segment.text)
			if strings.TrimSpace(statement) == "" {
				continue
			}
			if arraySourceOrderPreserveNeedsAllocatedInput(statement, arrayNames, state) {
				// A successful ReDim Preserve requires an already allocated
				// dynamic array. The source-order fallback has no exceptional
				// edge on which an On Error Resume Next failure could continue,
				// so an unallocated input cannot be used as an allocation proof.
				return nil, false
			}
			// Ordinary ReDim and array-factory assignments intentionally continue
			// through arrayTransfer below. Only an inline conditional mutation is
			// rejected because its branch state is not modeled by this fallback.
			if arraySourceOrderInlineArrayMutation(statement, arrayNames, ctx) {
				return nil, false
			}
			state, _ = a.arrayTransfer(file, caller, ctx, variables, state, statement, line, constants, nil)
			state = applyArrayLocalGoSubStatementEffects(state, statement, localGoSubAllocations)
			for _, previous := range callsBySegment[segmentIndex] {
				state = applyArrayModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
				state = applyArrayUnknownModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
				if arrayProcedureLineHasInlineConditional(file, previous.Range.StartLine) {
					state = applyArrayConditionalByRefCallEffects(state, caller, previous, ctx)
				} else {
					state = applyArrayByRefCallEffects(state, caller, previous, ctx)
				}
			}
		}
		for _, previous := range unassignedCalls {
			state = applyArrayModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
			state = applyArrayUnknownModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
			if arrayProcedureLineHasInlineConditional(file, previous.Range.StartLine) {
				state = applyArrayConditionalByRefCallEffects(state, caller, previous, ctx)
			} else {
				state = applyArrayByRefCallEffects(state, caller, previous, ctx)
			}
		}
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) || !bound[index] {
			continue
		}
		name := directArrayArgumentName(arguments[index])
		if name == "" {
			if !arrayQualifiedArgumentProvenAllocated(file, caller, call, arguments[index], ctx) {
				return nil, false
			}
			continue
		}
		value, known := state[name]
		if !known || value.kind != arrayAllocated || !value.knownArray || (!facts.allocationInvariant(name, call.Range.StartLine) && !initiallyAllocated[name]) {
			return nil, false
		}
	}
	return state, true
}

func arraySourceOrderPreserveNeedsAllocatedInput(text string, names map[string]bool, state arrayFlowState) bool {
	match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) == 0 || strings.TrimSpace(match[1]) == "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		name := strings.ToLower(cleanIdentifier(redim.name))
		if !direct || name == "" || !names[name] {
			continue
		}
		value, known := state[name]
		if !known || value.kind != arrayAllocated || !value.knownArray {
			return true
		}
	}
	return false
}

func arrayStatementAllocatesName(text, name string, ctx analysisContext) bool {
	if match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text)); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), name) {
				return true
			}
		}
	}
	if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
		value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx)
		return known && value.kind == arrayAllocated && value.knownArray
	}
	return false
}

func arrayByRefCallHasProvenArrayArguments(file parsedFile, target, caller sourceProcedure, call procedureir.CallSite, state arrayFlowState, ctx analysisContext) (bool, bool) {
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return false, false
	}
	foundExpression := false
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() {
			return false, false
		}
		parameter := target.Params.valueAt(binding.parameterIndex)
		if !parameterIsByRefArray(parameter) {
			continue
		}
		argument := binding.text
		if name := directArrayArgumentName(argument); name != "" {
			value, known := state[name]
			if !known || value.kind != arrayAllocated || !value.knownArray {
				return false, false
			}
			continue
		}
		if arrayQualifiedArgumentProvenAllocated(file, caller, call, argument, ctx) {
			foundExpression = true
			continue
		}
		if _, _, qualifiedMember := arrayQualifiedMemberParts(argument); qualifiedMember {
			return false, false
		}
		value, known := arrayExpressionState(argument, state, ctx)
		if !known || value.kind != arrayAllocated || !value.knownArray {
			return false, false
		}
		foundExpression = true
	}
	return true, foundExpression
}

func arrayByRefCallIsInnermostNested(call procedureir.CallSite, calls []arrayByRefCallCandidate) bool {
	nested := false
	for _, other := range calls {
		if arrayCallRangeContains(other.call, call) {
			nested = true
		}
		if arrayCallRangeContains(call, other.call) {
			return false
		}
	}
	return nested
}

func arrayCallRangeContains(outer, inner procedureir.CallSite) bool {
	if outer.ID == inner.ID {
		return false
	}
	if outer.Range.StartByte != 0 || outer.Range.EndByte != 0 || inner.Range.StartByte != 0 || inner.Range.EndByte != 0 {
		return outer.Range.StartByte <= inner.Range.StartByte && inner.Range.EndByte <= outer.Range.EndByte && (outer.Range.StartByte < inner.Range.StartByte || inner.Range.EndByte < outer.Range.EndByte)
	}
	return outer.Range.StartLine == inner.Range.StartLine && outer.Range.StartColumn <= inner.Range.StartColumn && inner.Range.EndColumn <= outer.Range.EndColumn && (outer.Range.StartColumn < inner.Range.StartColumn || inner.Range.EndColumn < outer.Range.EndColumn)
}

func arrayByRefCallArrayVacuouslyUnused(target sourceProcedure, arrayIndex int, arguments []string) bool {
	if arrayIndex < 0 || arrayIndex >= target.Params.Len() {
		return false
	}
	arrayName := strings.ToLower(cleanIdentifier(target.Params.valueAt(arrayIndex).Name))
	if arrayName == "" {
		return false
	}
	statements := make(map[int]procedureir.Statement, target.Statements.Len())
	for statement := range target.Statements.All() {
		statements[statement.ID] = statement
	}
	variables := map[string]arrayVariable{
		arrayName: {name: arrayName, isArray: true},
	}
	found := false
	for statement := range target.Statements.All() {
		if !arrayStatementReferencesArray(statement.Text, arrayName, variables) {
			continue
		}
		found = true
		if !arrayStatementHasFalseCallGuard(statement, statements, target, arguments) {
			return false
		}
	}
	return found
}

func arrayStatementReferencesArray(text, arrayName string, variables map[string]arrayVariable) bool {
	for _, use := range arrayIndexedUses(text, variables) {
		if strings.EqualFold(cleanIdentifier(use.name), arrayName) && len(use.args) > 0 {
			return true
		}
	}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		if strings.EqualFold(cleanIdentifier(bound[2]), arrayName) {
			return true
		}
	}
	return false
}

func arrayStatementHasFalseCallGuard(statement procedureir.Statement, statements map[int]procedureir.Statement, target sourceProcedure, arguments []string) bool {
	seen := map[int]bool{}
	for current := statement; current.ParentID != 0 && !seen[current.ParentID]; {
		seen[current.ParentID] = true
		parent, ok := statements[current.ParentID]
		if !ok {
			return false
		}
		if parent.Kind == procedureir.StatementIf && parent.Condition != nil && arrayConditionHasFalseCallArgument(parent.Condition.Text, target, arguments) {
			return true
		}
		current = parent
	}
	return false
}

func arrayConditionHasFalseCallArgument(condition string, target sourceProcedure, arguments []string) bool {
	for _, comparison := range arrayConditionAndRe.Split(condition, -1) {
		comparison = strings.TrimSpace(comparison)
		if index, ok := arrayProcedureParameterIndex(target, comparison); ok && index < len(arguments) && strings.EqualFold(strings.TrimSpace(arguments[index]), "false") {
			return true
		}
		lhs, operator, literal, ok := arrayCountComparison(comparison)
		if !ok {
			continue
		}
		index, ok := arrayProcedureParameterIndex(target, lhs)
		if !ok || index >= len(arguments) {
			continue
		}
		value, valueOK := integerLiteral(arguments[index])
		bound, boundOK := integerLiteral(literal)
		if valueOK && boundOK && arrayIntegerComparisonFalse(value, operator, bound) {
			return true
		}
	}
	return false
}

func arrayProcedureParameterIndex(proc sourceProcedure, name string) (int, bool) {
	name = strings.ToLower(cleanIdentifier(name))
	for index, parameter := range proc.Params.AllIndexed() {
		if strings.ToLower(cleanIdentifier(parameter.Name)) == name {
			return index, true
		}
	}
	return 0, false
}

func arrayIntegerComparisonFalse(value int, operator string, bound int) bool {
	switch operator {
	case "=":
		return value != bound
	case "<>":
		return value == bound
	case ">":
		return value <= bound
	case ">=":
		return value < bound
	case "<":
		return value >= bound
	case "<=":
		return value > bound
	default:
		return false
	}
}

func arrayConditionalEntrySource(target sourceProcedure, arguments []string, arrayParameterIndex int, source string) string {
	if source == "" {
		return ""
	}
	for index, argument := range arguments {
		if index == arrayParameterIndex || !arrayCountExpressionMatches(argument, source) {
			continue
		}
		if index < 0 || index >= target.Params.Len() || parameterIsByRefArray(target.Params.valueAt(index)) {
			continue
		}
		return strings.ToLower(target.Params.valueAt(index).Name)
	}
	return ""
}

func arrayCallArgumentTexts(proc sourceProcedure, call procedureir.CallSite) []string {
	if len(call.Arguments.ExpressionIDs) == 0 {
		return nil
	}
	facts := proc.analysisFacts()
	texts := make([]string, 0, len(call.Arguments.ExpressionIDs))
	for _, id := range call.Arguments.ExpressionIDs {
		expression, ok := facts.Expression(id)
		if !ok {
			return nil
		}
		texts = append(texts, strings.TrimSpace(expression.Text))
	}
	return texts
}

type arrayCallArgumentBinding struct {
	parameterIndex int
	text           string
}

// arrayCallArgumentBindings maps actual arguments to formal parameters while
// retaining VBA's source order for positional arguments. Named arguments are
// identified by their expression IDs, so a positional argument before a named
// argument is not mistaken for the named formal slot.
func arrayCallArgumentBindings(caller sourceProcedure, target sourceProcedure, call procedureir.CallSite) ([]arrayCallArgumentBinding, bool) {
	argumentCount := call.Arguments.Count
	// Some unparenthesized calls with an implicit member argument currently
	// expose the expression ID but leave Count at zero.  Prefer the concrete
	// expression projection when Count is absent; otherwise the ByRef transfer
	// would silently discard a real actual argument.
	if argumentCount == 0 {
		argumentCount = max(len(call.Arguments.ExpressionIDs), len(call.Arguments.Named))
	}
	if argumentCount == 0 {
		if argument, ok := arrayImplicitMemberArgument(call); ok {
			if target.Params.Len() != 1 {
				return nil, false
			}
			return []arrayCallArgumentBinding{{parameterIndex: 0, text: argument}}, true
		}
		return nil, true
	}
	if len(call.Arguments.ExpressionIDs) == 0 {
		if len(call.Arguments.Named) != argumentCount {
			return nil, false
		}
		bindings := make([]arrayCallArgumentBinding, 0, len(call.Arguments.Named))
		used := map[int]bool{}
		for _, named := range call.Arguments.Named {
			index, ok := arrayFormalParameterIndex(target, named.Name)
			if !ok || used[index] {
				return nil, false
			}
			used[index] = true
			bindings = append(bindings, arrayCallArgumentBinding{parameterIndex: index, text: strings.TrimSpace(named.ValueText)})
		}
		return bindings, true
	}
	texts := arrayCallArgumentTexts(caller, call)
	if len(texts) != argumentCount || len(call.Arguments.ExpressionIDs) != len(texts) {
		return nil, false
	}
	namedByExpressionID := make(map[int]int, len(call.Arguments.Named))
	for _, named := range call.Arguments.Named {
		if named.ExpressionID == 0 {
			return nil, false
		}
		index, ok := arrayFormalParameterIndex(target, named.Name)
		if !ok {
			return nil, false
		}
		if _, exists := namedByExpressionID[named.ExpressionID]; exists {
			return nil, false
		}
		namedByExpressionID[named.ExpressionID] = index
	}
	bindings := make([]arrayCallArgumentBinding, 0, len(texts))
	used := map[int]bool{}
	nextPositional := 0
	for actualIndex, text := range texts {
		formalIndex, named := namedByExpressionID[call.Arguments.ExpressionIDs[actualIndex]]
		if !named {
			for nextPositional < target.Params.Len() && used[nextPositional] {
				nextPositional++
			}
			if nextPositional < target.Params.Len() {
				formalIndex = nextPositional
				nextPositional++
			} else {
				last := target.Params.Len() - 1
				if last < 0 || !target.Params.valueAt(last).ParamArray {
					return nil, false
				}
				// Every extra positional argument belongs to the trailing
				// ParamArray formal. Keep one binding per actual so a direct
				// array argument among the extras still participates in the
				// ByRef effect analysis.
				formalIndex = last
			}
		}
		if used[formalIndex] && (formalIndex < 0 || formalIndex >= target.Params.Len() || !target.Params.valueAt(formalIndex).ParamArray) {
			return nil, false
		}
		used[formalIndex] = true
		bindings = append(bindings, arrayCallArgumentBinding{parameterIndex: formalIndex, text: text})
	}
	return bindings, true
}

func arrayImplicitMemberArgument(call procedureir.CallSite) (string, bool) {
	if call.Arguments.Count != 0 || len(call.Arguments.ExpressionIDs) != 0 || len(call.Arguments.Named) != 0 || call.Callee.Receiver == nil {
		return "", false
	}
	procedureName := arrayImplicitMemberArgumentCallName(call.Callee.Text)
	if procedureName == "" || !strings.EqualFold(procedureName, *call.Callee.Receiver) {
		return "", false
	}
	member := cleanIdentifier(call.Callee.Member)
	if member == "" {
		return "", false
	}
	return "." + member, true
}

func arrayCallFormalArguments(caller sourceProcedure, target sourceProcedure, call procedureir.CallSite) ([]string, bool) {
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return nil, false
	}
	arguments := make([]string, target.Params.Len())
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= len(arguments) {
			return nil, false
		}
		arguments[binding.parameterIndex] = binding.text
	}
	return arguments, true
}

func arrayFormalParameterIndex(target sourceProcedure, name string) (int, bool) {
	name = strings.TrimSpace(name)
	for index, parameter := range target.Params.AllIndexed() {
		if strings.EqualFold(strings.TrimSpace(parameter.Name), name) {
			return index, true
		}
	}
	return 0, false
}

func arrayCallPassesDirectArrayArgument(proc sourceProcedure, call procedureir.CallSite, name string) bool {
	name = strings.ToLower(cleanIdentifier(name))
	for _, argument := range call.Arguments.Named {
		if directArrayArgumentName(argument.ValueText) == name {
			return true
		}
	}
	for _, argument := range arrayCallArgumentTexts(proc, call) {
		if directArrayArgumentName(argument) == name {
			return true
		}
	}
	return false
}

func directArrayArgumentName(text string) string {
	text = strings.TrimSpace(text)
	if text == "" || !isIdentifierStart(text[0]) {
		return ""
	}
	for i := 1; i < len(text); i++ {
		if !isIdentifierPart(text[i]) {
			return ""
		}
	}
	return strings.ToLower(cleanIdentifier(text))
}

// arrayQualifiedArgumentProvenAllocated carries narrow allocation proofs for a
// qualified array member passed to a private ByRef array parameter. The
// ordinary array state is intentionally keyed by local identifiers, so a call
// such as `Consume holder.values` cannot otherwise reuse a dominating ReDim of
// the member. The additional proofs are limited to normal-path ReDim,
// descriptor-backed array setup, and dictionary snapshots whose non-empty
// range is established by the same caller.
func arrayQualifiedArgumentProvenAllocated(file parsedFile, caller sourceProcedure, call procedureir.CallSite, argument string, ctx analysisContext) bool {
	want := arrayQualifiedArgumentTarget(file, caller, call.Range.StartLine, argument)
	if want == "" || caller.Graph == nil {
		return false
	}
	if arrayQualifiedDescriptorArgumentProvenAllocated(file, caller, call, argument, ctx) ||
		arrayQualifiedDictionarySnapshotProvenAllocated(file, caller, call, argument) {
		return true
	}
	for statement := range caller.Statements.All() {
		line := statement.Range.StartLine
		dominates := arrayStatementDominatesCall(caller, statement.ID, statement.Range.StartLine, call)
		redim := arrayQualifiedRedimAllocatesTarget(file, caller, line, statement.Text, want)
		if line <= caller.StartLine || line >= call.Range.StartLine || !dominates {
			continue
		}
		if redim {
			return true
		}
	}
	return false
}

// arrayQualifiedDescriptorArgumentProvenAllocated recognizes the accessor
// pattern used when a VBA SAFEARRAY descriptor is projected onto a typed array
// member. The array itself is not visible to the source-level analyzer, but a
// dominating data pointer and element-count write, followed by a positive
// count guard, establish the same normal-path contract as ReDim. A projected
// `rgsabound()` array has one additional form: its caller can establish a
// successful `ReDim ... (0 To ub)` from the descriptor dimension count before
// passing the accessor through another private helper. That normal-path fact
// proves the descriptor count is positive without requiring a duplicated scalar
// count argument at every helper boundary.
func arrayQualifiedDescriptorArgumentProvenAllocated(file parsedFile, caller sourceProcedure, call procedureir.CallSite, argument string, ctx analysisContext) bool {
	receiver, member, ok := arrayQualifiedMemberParts(argument)
	if !ok || !strings.EqualFold(member, "arr") && !strings.EqualFold(member, "rgsabound") {
		return false
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, caller, call.Range.StartLine)
	}
	if receiver == "" {
		return false
	}
	arguments := arrayCallArgumentTexts(caller, call)
	if len(arguments) < 2 {
		if !strings.EqualFold(member, "rgsabound") {
			return false
		}
		normalPathNonEmpty := arrayQualifiedDescriptorHasSuccessfulShapeUse(file, caller, call, receiver)
		return arrayQualifiedDescriptorReceiverInitialized(file, caller, call, receiver, ctx, !normalPathNonEmpty, map[string]bool{})
	}
	count := canonicalArrayBoundExpression(arguments[1])
	if count != "" && arrayQualifiedDescriptorCountPositive(file, caller, call, arguments[1]) {
		wantReceiver := canonicalArrayBoundExpression(receiver)
		wantData := wantReceiver + ".sa.pvdata"
		wantBounds := wantReceiver + ".sa.rgsabound0.celements"
		hasData := false
		hasBounds := false
		for statement := range caller.Statements.All() {
			line := statement.Range.StartLine
			if line <= caller.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(caller, statement.ID, line, call) {
				continue
			}
			text := strings.TrimSpace(statement.Text)
			if text == "" {
				text = arrayLogicalSourceLine(file.Lines, line)
			}
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if !assigned || indexed {
				continue
			}
			target := canonicalArrayBoundExpression(lhs)
			switch target {
			case wantData:
				hasData = strings.TrimSpace(rhs) != ""
			case wantBounds:
				hasBounds = canonicalArrayBoundExpression(rhs) == count
			}
		}
		if hasData && hasBounds {
			return true
		}
	}
	if !strings.EqualFold(member, "rgsabound") || !arrayQualifiedDescriptorHasSuccessfulShapeUse(file, caller, call, receiver) {
		return false
	}
	return arrayQualifiedDescriptorReceiverInitialized(file, caller, call, receiver, ctx, false, map[string]bool{})
}

func arrayQualifiedDescriptorReceiverInitialized(file parsedFile, proc sourceProcedure, call procedureir.CallSite, receiver string, ctx analysisContext, requirePositive bool, visiting map[string]bool) bool {
	receiver = strings.ToLower(cleanIdentifier(strings.TrimSpace(receiver)))
	if receiver == "" {
		return false
	}
	parameterIndex, isParameter := arrayProcedureParameterIndex(proc, receiver)
	if !isParameter {
		return arrayQualifiedDescriptorLocalInitialized(file, proc, call, receiver, requirePositive)
	}
	visitKey := arrayProcedureKey(proc) + ":" + receiver
	if visiting[visitKey] {
		return false
	}
	visiting[visitKey] = true
	defer delete(visiting, visitKey)

	foundCaller := false
	for _, candidate := range file.Procedures {
		for nested := range candidate.Calls.All() {
			targetKey, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, nested)
			if !ok || targetKey != arrayProcedureKey(proc) {
				continue
			}
			bindings, bound := arrayCallArgumentBindings(candidate, target, nested)
			if !bound {
				return false
			}
			actual := ""
			for _, binding := range bindings {
				if binding.parameterIndex == parameterIndex {
					actual = strings.TrimSpace(binding.text)
					break
				}
			}
			if directArrayArgumentName(actual) == "" {
				return false
			}
			foundCaller = true
			if _, isParameter := arrayProcedureParameterIndex(candidate, actual); isParameter {
				if !arrayQualifiedDescriptorReceiverInitialized(file, candidate, nested, actual, ctx, requirePositive, visiting) {
					return false
				}
				continue
			}
			if !arrayQualifiedDescriptorLocalInitialized(file, candidate, nested, actual, requirePositive) {
				return false
			}
		}
	}
	return foundCaller
}

func arrayQualifiedDescriptorLocalInitialized(file parsedFile, proc sourceProcedure, call procedureir.CallSite, receiver string, requirePositive bool) bool {
	wantReceiver := canonicalArrayBoundExpression(receiver)
	wantData := wantReceiver + ".sa.pvdata"
	wantBounds := wantReceiver + ".sa.rgsabound0.celements"
	hasData := false
	count := ""
	for statement := range proc.Statements.All() {
		line := statement.Range.StartLine
		if line <= proc.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(proc, statement.ID, line, call) {
			continue
		}
		for _, text := range arrayQualifiedDescriptorStatementTexts(file, statement) {
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if !assigned || indexed {
				continue
			}
			switch canonicalArrayBoundExpression(lhs) {
			case wantData:
				canonical := canonicalArrayBoundExpression(rhs)
				hasData = canonical != "" && canonical != "0" && canonical != "nullptr"
			case wantBounds:
				count = canonicalArrayBoundExpression(rhs)
			}
		}
	}
	if !hasData || count == "" {
		return false
	}
	if !requirePositive {
		return true
	}
	if value, err := strconv.Atoi(count); err == nil {
		return value > 0
	}
	return arrayQualifiedDescriptorCountPositive(file, proc, call, count)
}

func arrayQualifiedDescriptorStatementTexts(file parsedFile, statement procedureir.Statement) []string {
	texts := make([]string, 0, 4)
	seen := map[string]bool{}
	add := func(source string) {
		for _, segment := range splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(source)) {
			text := strings.TrimSpace(segment.text)
			if text != "" && !seen[text] {
				seen[text] = true
				texts = append(texts, text)
			}
		}
	}
	add(statement.Text)
	if statement.Range.StartLine >= 1 && statement.Range.StartLine <= len(file.Lines) {
		add(arrayLogicalSourceLine(file.Lines, statement.Range.StartLine))
	}
	return texts
}

func arrayQualifiedDescriptorHasSuccessfulShapeUse(file parsedFile, proc sourceProcedure, call procedureir.CallSite, receiver string) bool {
	wantCount := canonicalArrayBoundExpression(receiver + ".sa.rgsabound0.cElements")
	upperNames := map[string]int{}
	for statement := range proc.Statements.All() {
		line := statement.Range.StartLine
		if line <= proc.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(proc, statement.ID, line, call) {
			continue
		}
		for _, text := range arrayQualifiedDescriptorStatementTexts(file, statement) {
			if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed && canonicalArrayBoundExpression(rhs) == wantCount+"-1" {
				upperNames[strings.ToLower(cleanIdentifier(lhs))] = line
			}
			match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text))
			if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
				continue
			}
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct || redim.name == "" {
					continue
				}
				dimensions := canonicalArrayBoundExpression(redim.dimensions)
				if !strings.HasPrefix(dimensions, "0to") {
					continue
				}
				upper := strings.TrimPrefix(dimensions, "0to")
				if assignmentLine, ok := upperNames[strings.ToLower(cleanIdentifier(upper))]; ok && assignmentLine < line {
					return true
				}
			}
		}
	}
	return false
}

// arrayQualifiedDescriptorCountPositive proves that the count used to build a
// descriptor-backed array is positive on the call path. It understands direct
// comparison branches and a zero guard whose then arm jumps to a label after
// the call (the normal-path shape used by parser entry points).
func arrayQualifiedDescriptorCountPositive(file parsedFile, proc sourceProcedure, call procedureir.CallSite, count string) bool {
	wantCount := canonicalArrayBoundExpression(count)
	if wantCount == "" {
		return false
	}
	for guard := range proc.Statements.All() {
		if guard.Kind != procedureir.StatementIf && guard.Kind != procedureir.StatementElseIf || !arrayStatementDominatesCall(proc, guard.ID, guard.Range.StartLine, call) {
			continue
		}
		condition := guard.Text
		if guard.Condition != nil && strings.TrimSpace(guard.Condition.Text) != "" {
			condition = guard.Condition.Text
		}
		lhs, operator, literal, ok := arrayQualifiedCountComparison(condition)
		if !ok || lhs != wantCount {
			continue
		}
		positiveBranch, positive := positiveArrayCountBranch(operator, literal)
		if !positive {
			continue
		}
		branch, underGuard := arrayQualifiedStatementBranch(proc, call.StatementID, guard.ID)
		if underGuard && arrayQualifiedBranchMatchesPositive(branch, positiveBranch) {
			return true
		}
		if operator == "=" && literal == "0" && !underGuard && arrayQualifiedZeroGuardSkipsCall(proc, guard, call) && arrayQualifiedCountHasNonNegativeOrigin(file, proc, call, count) {
			return true
		}
	}
	return false
}

func arrayQualifiedCountComparison(text string) (lhs, operator, literal string, ok bool) {
	text = strings.TrimSpace(text)
	if condition, _, hasThen := arrayIfThenParts(text); hasThen {
		text = condition
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "if ") {
		text = strings.TrimSpace(text[len("if "):])
	} else if strings.HasPrefix(lower, "elseif ") {
		text = strings.TrimSpace(text[len("elseif "):])
	}
	text = arrayQualifiedTrimOuterParens(text)
	lhs, operator, literal, ok = arrayCountComparison(text)
	if !ok {
		return "", "", "", false
	}
	lhs = canonicalArrayBoundExpression(arrayQualifiedTrimOuterParens(lhs))
	return lhs, operator, literal, lhs != ""
}

func arrayQualifiedTrimOuterParens(text string) string {
	text = strings.TrimSpace(text)
	for strings.HasPrefix(text, "(") {
		close := matchingParen(text, 0)
		if close != len(text)-1 {
			break
		}
		text = strings.TrimSpace(text[1:close])
	}
	return text
}

func arrayQualifiedStatementBranch(proc sourceProcedure, statementID, ancestorID int) (procedureir.BranchRole, bool) {
	seen := map[int]bool{}
	for statementID > 0 && !seen[statementID] {
		seen[statementID] = true
		statement, ok := arrayProcedureStatementByID(proc, statementID)
		if !ok {
			return "", false
		}
		if statement.ID == ancestorID {
			if statement.Kind == procedureir.StatementIf && statement.SyntaxKind == "single_line_if_statement" {
				return procedureir.BranchThen, true
			}
			return "", false
		}
		if statement.ParentID == ancestorID {
			if statement.Kind == procedureir.StatementElse {
				return procedureir.BranchElse, true
			}
			return procedureir.BranchThen, true
		}
		statementID = statement.ParentID
	}
	return "", false
}

func arrayQualifiedBranchMatchesPositive(branch procedureir.BranchRole, positive vbacfg.EdgeKind) bool {
	switch branch {
	case procedureir.BranchThen:
		return positive == vbacfg.EdgeBranchTrue
	case procedureir.BranchElse:
		return positive == vbacfg.EdgeBranchFalse
	default:
		return false
	}
}

func arrayQualifiedZeroGuardSkipsCall(proc sourceProcedure, guard procedureir.Statement, call procedureir.CallSite) bool {
	thenStatements := make([]procedureir.Statement, 0)
	for statement := range proc.Statements.All() {
		if statement.ParentID != guard.ID || statement.Kind == procedureir.StatementElseIf || statement.Kind == procedureir.StatementElse {
			continue
		}
		thenStatements = append(thenStatements, statement)
	}
	if len(thenStatements) == 0 {
		return false
	}
	sort.SliceStable(thenStatements, func(i, j int) bool {
		if thenStatements[i].Range.StartLine != thenStatements[j].Range.StartLine {
			return thenStatements[i].Range.StartLine < thenStatements[j].Range.StartLine
		}
		return thenStatements[i].ID < thenStatements[j].ID
	})
	for _, statement := range thenStatements[:len(thenStatements)-1] {
		if statement.Kind == procedureir.StatementIf || statement.Kind == procedureir.StatementElseIf || statement.Kind == procedureir.StatementElse || statement.Control != nil && statement.Control.Transfer != "" {
			return false
		}
	}
	return arrayQualifiedStatementLeavesBeforeCall(proc, thenStatements[len(thenStatements)-1], call)
}

func arrayQualifiedStatementLeavesBeforeCall(proc sourceProcedure, statement procedureir.Statement, call procedureir.CallSite) bool {
	if statement.Control != nil {
		switch statement.Control.Transfer {
		case procedureir.TransferExitSub, procedureir.TransferExitFunction, procedureir.TransferExitProperty, procedureir.TransferTerminate:
			return true
		case procedureir.TransferGoto:
			labelLine, ok := arrayQualifiedLabelLine(proc, statement.Control.Target)
			return ok && labelLine > call.Range.StartLine
		}
	}
	switch statement.Kind {
	case procedureir.StatementEnd:
		return true
	case procedureir.StatementExit:
		return arraySourceOrderProcedureExitStatementText(statement.Text)
	default:
		return false
	}
}

func arrayQualifiedLabelLine(proc sourceProcedure, target string) (int, bool) {
	target = strings.ToLower(strings.TrimSpace(strings.TrimSuffix(target, ":")))
	if target == "" {
		return 0, false
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementLabel || !strings.EqualFold(arrayLocalGoSubLabelName(statement), target) {
			continue
		}
		return statement.Range.StartLine, true
	}
	return 0, false
}

func arrayQualifiedCountHasNonNegativeOrigin(file parsedFile, proc sourceProcedure, call procedureir.CallSite, count string) bool {
	want := canonicalArrayBoundExpression(count)
	hasAssignment := false
	hasSafeOrigin := false
	statements := make([]procedureir.Statement, 0, proc.Statements.Len())
	for statement := range proc.Statements.All() {
		statements = append(statements, statement)
	}
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartLine != statements[j].Range.StartLine {
			return statements[i].Range.StartLine < statements[j].Range.StartLine
		}
		return statements[i].ID < statements[j].ID
	})
	for _, statement := range statements {
		line := statement.Range.StartLine
		if line <= proc.StartLine || line >= call.Range.StartLine {
			continue
		}
		texts := []string{strings.TrimSpace(statement.Text)}
		if source := arrayLogicalSourceLine(file.Lines, line); source != "" && source != texts[0] {
			texts = append(texts, source)
		}
		matched := false
		assignmentSafe := false
		for _, text := range texts {
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if !assigned || indexed || canonicalArrayBoundExpression(lhs) != want {
				continue
			}
			matched = true
			canonical := canonicalArrayBoundExpression(rhs)
			candidateSafe := false
			switch {
			case arrayQualifiedNonNegativeCountOrigin(canonical):
				candidateSafe = true
			case canonical == want+"*2", canonical == "2*"+want:
				// A byte count multiplied by two remains non-negative after the
				// masked or LenB origin has been seen.
				candidateSafe = hasSafeOrigin
			case canonical == "0":
				// The default zero value is also non-negative.
				candidateSafe = true
			default:
				continue
			}
			assignmentSafe = assignmentSafe || candidateSafe
		}
		if !matched {
			continue
		}
		hasAssignment = true
		if !assignmentSafe {
			// An unknown later assignment invalidates an earlier safe origin;
			// otherwise a count such as `sizeB = -1` could inherit proof from a
			// preceding LenB assignment.  The source fallback above prevents a
			// truncated IR expression from hiding a safe full-line assignment.
			return false
		}
		hasSafeOrigin = true
	}
	return hasAssignment && hasSafeOrigin
}

func arrayQualifiedNonNegativeCountOrigin(canonical string) bool {
	canonical = canonicalArrayBoundExpression(canonical)
	for strings.HasPrefix(canonical, "clng(") {
		close := matchingParen(canonical, len("clng"))
		if close != len(canonical)-1 {
			break
		}
		canonical = canonical[len("clng("):close]
	}
	if strings.HasPrefix(canonical, "lenb(") && matchingParen(canonical, len("lenb")) == len(canonical)-1 {
		return true
	}
	for _, mask := range []string{"and&h7fffffff", "and&hffff&", "and&hffff", "and&hff&", "and&hff"} {
		if strings.HasSuffix(canonical, mask) {
			return true
		}
	}
	return false
}

func arrayQualifiedDictionarySnapshotProvenAllocated(file parsedFile, caller sourceProcedure, call procedureir.CallSite, argument string) bool {
	receiver, member, ok := arrayQualifiedMemberParts(argument)
	if !ok || !strings.EqualFold(member, "arrkeys") && !strings.EqualFold(member, "arritems") {
		return false
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, caller, call.Range.StartLine)
	}
	if receiver == "" {
		return false
	}
	arguments := arrayCallArgumentTexts(caller, call)
	if len(arguments) == 0 || !arrayQualifiedUpperBoundProvenNonNegative(file, caller, call, arguments[len(arguments)-1]) {
		return false
	}
	want := canonicalArrayBoundExpression(receiver + "." + member)
	for line := max(1, caller.StartLine); line < call.Range.StartLine && line <= len(file.Lines); line++ {
		source := arrayLogicalSourceLine(file.Lines, line)
		if source == "" {
			continue
		}
		if _, body, ok := arrayIfThenParts(source); ok {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			if hasElse && arrayQualifiedSnapshotAssignmentArm(file, caller, line, thenBody, want, member) && arrayQualifiedSnapshotAssignmentArm(file, caller, line, elseBody, want, member) && arrayQualifiedSourceLineDominatesCall(caller, line, call) {
				return true
			}
			continue
		}
		if arrayQualifiedSnapshotAssignmentArm(file, caller, line, source, want, member) && arrayQualifiedSourceLineDominatesCall(caller, line, call) {
			return true
		}
	}
	return false
}

func arrayQualifiedSnapshotAssignmentArm(file parsedFile, caller sourceProcedure, line int, text, want, member string) bool {
	lhs, rhs, indexed, assigned := arrayAssignment(text)
	if !assigned || indexed || arrayQualifiedArgumentTarget(file, caller, line, lhs) != want {
		return false
	}
	_, snapshotMember, ok := arrayDictionaryMemberParts(rhs)
	return ok && strings.EqualFold(snapshotMember, member[len("arr"):])
}

func arrayQualifiedSourceLineDominatesCall(proc sourceProcedure, line int, call procedureir.CallSite) bool {
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine == line && arrayStatementDominatesCall(proc, statement.ID, line, call) {
			return true
		}
	}
	return false
}

func arrayQualifiedUpperBoundProvenNonNegative(file parsedFile, proc sourceProcedure, call procedureir.CallSite, argument string) bool {
	want := arrayQualifiedArgumentTarget(file, proc, call.Range.StartLine, argument)
	if want == "" {
		return false
	}
	for guard := range proc.Statements.All() {
		if guard.Kind != procedureir.StatementIf && guard.Kind != procedureir.StatementElseIf || !arrayStatementDominatesCall(proc, guard.ID, guard.Range.StartLine, call) {
			continue
		}
		condition := guard.Text
		if guard.Condition != nil && strings.TrimSpace(guard.Condition.Text) != "" {
			condition = guard.Condition.Text
		}
		lhs, operator, literal, ok := arrayQualifiedCountComparison(condition)
		if !ok {
			continue
		}
		positiveBranch, positive := positiveArrayCountBranch(operator, literal)
		if !positive {
			continue
		}
		callBranch, underGuard := arrayQualifiedStatementBranch(proc, call.StatementID, guard.ID)
		if !underGuard || !arrayQualifiedBranchMatchesPositive(callBranch, positiveBranch) {
			continue
		}
		for statement := range proc.Statements.All() {
			line := statement.Range.StartLine
			if line <= proc.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(proc, statement.ID, line, call) {
				continue
			}
			statementBranch, sameBranch := arrayQualifiedStatementBranch(proc, statement.ID, guard.ID)
			if !sameBranch || statementBranch != callBranch {
				continue
			}
			text := strings.TrimSpace(statement.Text)
			if text == "" {
				text = arrayLogicalSourceLine(file.Lines, line)
			}
			assignment, rhs, indexed, assigned := arrayAssignment(text)
			if assigned && !indexed && arrayQualifiedArgumentTarget(file, proc, line, assignment) == want && canonicalArrayBoundExpression(rhs) == lhs {
				return true
			}
		}
	}
	return false
}

func arrayQualifiedArgumentTarget(file parsedFile, caller sourceProcedure, line int, argument string) string {
	receiver, member, ok := arrayQualifiedMemberParts(argument)
	if !ok {
		return ""
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, caller, line)
	}
	if receiver == "" {
		return ""
	}
	return canonicalArrayBoundExpression(receiver + "." + member)
}

func arrayQualifiedMemberParts(text string) (receiver, member string, ok bool) {
	text = strings.TrimSpace(text)
	if text == "" {
		return "", "", false
	}
	dot := strings.LastIndexByte(text, '.')
	if dot < 0 || dot >= len(text)-1 {
		return "", "", false
	}
	member = cleanIdentifier(strings.TrimSpace(text[dot+1:]))
	if !arrayEraseNameRe.MatchString(member) {
		return "", "", false
	}
	receiver = strings.TrimSpace(text[:dot])
	if receiver == "" {
		if text[0] != '.' {
			return "", "", false
		}
		return "", member, true
	}
	for _, part := range strings.Split(receiver, ".") {
		if !arrayEraseNameRe.MatchString(strings.TrimSpace(part)) {
			return "", "", false
		}
	}
	return receiver, member, true
}

func arrayQualifiedRedimAllocatesTarget(file parsedFile, caller sourceProcedure, line int, text, want string) bool {
	match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		clause = strings.TrimSpace(clause)
		open := firstParenOutsideString(clause)
		if open <= 0 {
			continue
		}
		close := matchingParen(clause, open)
		if close < 0 {
			continue
		}
		remainder := strings.TrimSpace(clause[close+1:])
		if remainder != "" && !arrayRedimTypeSuffixRe.MatchString(remainder) {
			continue
		}
		target := arrayQualifiedArgumentTarget(file, caller, line, strings.TrimSpace(clause[:open]))
		if target == "" || target != want {
			continue
		}
		dimensions := parseArrayDimensionsWithConstants(clause[open+1:close], arrayOptionBase(file), nil)
		if arrayDimensionsKnownNonEmpty(dimensions) {
			return true
		}
	}
	return false
}

func arrayDimensionsKnownNonEmpty(dimensions []arrayDimension) bool {
	// A plain ReDim that reaches the following statement has established an
	// allocation. Unknown bounds may make ReDim fail before that continuation,
	// but they do not make the successfully reached array empty unless the
	// bounds are provably impossible. Preserve the existing conservative check
	// for the latter case while allowing runtime bounds such as `0 To .ub`.
	return len(dimensions) > 0 && !impossibleArrayBounds(dimensions)
}

func arrayStatementDominatesCall(proc sourceProcedure, statementID, statementLine int, call procedureir.CallSite) bool {
	if proc.Graph == nil || statementID <= 0 || call.StatementID <= 0 {
		return false
	}
	statementBlock, statementOK := proc.Graph.BlockForStatement(statementID)
	callBlock, callOK := proc.Graph.BlockForStatement(call.StatementID)
	if !statementOK || !callOK {
		return false
	}
	if statementBlock.ID == callBlock.ID {
		return statementLine < call.Range.StartLine
	}
	for _, dominator := range proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[callBlock.ID] {
		if dominator == statementBlock.ID {
			return true
		}
	}
	return false
}

func parameterIsByRefArray(parameter parameterInfo) bool {
	// ParamArray materializes a new Variant array in the callee. Its actual
	// arguments are elements of that array, not aliases to caller arrays.
	return !parameter.ParamArray && parameterIsArray(parameter) && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
}

func parameterIsByRefScalar(parameter parameterInfo) bool {
	return !parameterIsArray(parameter) && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
}

func (a Analyzer) arrayTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) (arrayFlowState, []Finding) {
	if state == nil {
		state = arrayFlowState{}
	}
	if assignment, ok := inlineArrayAssignmentText(text); ok {
		text = assignment
	}
	var findings []Finding
	addWithKey := func(operationKey, code, message, reason, suggestion string) {
		if code == "VBA227" && !a.Config.Analyze.DetectArrayLifecycleSafety {
			return
		}
		if code == "VBA208" && !a.Config.Analyze.DetectRedimPreserveDimension {
			return
		}
		if code == "VBA209" && !a.Config.Analyze.DetectObjectArrayComparison {
			return
		}
		finding := a.simpleFinding(file, proc, line, code, "warning", message, reason, suggestion)
		finding.arrayLifecycleFinding = code == "VBA227"
		finding.arrayOperationKey = operationKey
		findings = append(findings, finding)
	}
	add := func(code, message, reason, suggestion string) {
		addWithKey("", code, message, reason, suggestion)
	}

	lower := strings.ToLower(strings.TrimSpace(text))
	if declRe.MatchString(text) || isProcedureHeaderLine(lower) {
		return state, findings
	}
	allocationProbeParameter := ""
	if ctx.arrayAllocationGuards[strings.ToLower(proc.Name)] {
		allocationProbeParameter, _ = arrayAllocationGuardParameter(proc)
	}
	if allocationProbeParameter == "" && ctx.arraySafeBoundGuards[strings.ToLower(proc.Name)] {
		allocationProbeParameter, _ = arraySafeBoundGuardParameter(proc)
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		base := arrayOptionBase(file)
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct {
				// Keep the existing receiver-array state transition for nested
				// member ReDim statements. VBA208 must not attribute the member
				// shape to that receiver, but changing the shared VBA227 state is
				// a separate analyzer contract.
				legacy := arrayRedimClauseRe.FindStringSubmatch(clause)
				if len(legacy) == 0 {
					continue
				}
				redim = directArrayRedimClause{name: legacy[1], dimensions: legacy[2]}
			}
			name := strings.ToLower(redim.name)
			variable, known := variables[name]
			// An unresolved target may be an implicit Variant or an external
			// member.  ReDim is valid for some of those runtime shapes, so only
			// report a target mismatch once the shared declaration facts prove a
			// scalar or fixed array.
			if !known {
				continue
			}
			old := state[name]
			dimensions := parseArrayDimensionsWithConstants(redim.dimensions, base, constants)
			// Ordinary analysis keeps an unknown Variant conservative. During
			// documented array-return inference, VBE-confirmed non-Preserve
			// ReDim semantics allow the Variant to establish its array value.
			preserve := strings.TrimSpace(match[1]) != ""
			variantArray := variable.isVariant && (old.knownArray || ctx.arrayAllowVariantRedim && !preserve)
			resizable := (variable.isArray || variantArray) && !variable.fixed
			if variable.isVariant && !old.knownArray && (!ctx.arrayAllowVariantRedim || preserve) {
				continue
			}
			if !resizable {
				// An unresolved object/UDT/external declaration is not a proven
				// scalar.  Leave it unknown rather than guessing that ReDim is
				// invalid; the shared shape contract is deliberately fail-open.
				if !variable.isArray && !variable.isVariant && !variable.knownScalar {
					continue
				}
				add("VBA227", redim.name+" is not a dynamic array and cannot be resized with ReDim.", "ReDim requires a dynamic array; fixed-size arrays and scalar values have no resizable allocation state.", "Declare the value as a dynamic array, or remove ReDim and use its declared bounds.")
			} else if impossibleArrayBounds(dimensions) {
				add("VBA227", redim.name+" has impossible constant ReDim bounds.", "A ReDim lower bound cannot be greater than its upper bound.", "Use bounds whose lower value is less than or equal to the upper value, or keep the bounds dynamic.")
			} else if direct && match[1] != "" && !preserveDimensionsSafe(old.preserveShape, dimensions) {
				add("VBA208", "ReDim Preserve may change a non-final or unknown array dimension.", "VBA can only preserve an array while changing its final dimension, and that cannot be proven when the prior shape is unknown.", "Only change the final dimension, or copy values into a newly sized array explicitly.")
			}
			if resizable && !impossibleArrayBounds(dimensions) {
				next := arrayValue{kind: arrayAllocated, knownArray: true, dimensions: dimensions, origin: arrayOriginLocal}
				if direct {
					next.preserveShape = dimensions
				} else {
					next.preserveShape = append([]arrayDimension(nil), old.preserveShape...)
				}
				state[name] = next
			}
		}
		return state, findings
	}
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) > 0 {
		for _, target := range splitArgs(match[1]) {
			name := strings.ToLower(strings.TrimSpace(target))
			if !arrayEraseNameRe.MatchString(name) {
				continue
			}
			if variable, ok := variables[name]; ok {
				if variable.fixed {
					state[name] = arrayValue{
						kind:          arrayAllocated,
						knownArray:    true,
						dimensions:    append([]arrayDimension(nil), variable.dimensions...),
						preserveShape: append([]arrayDimension(nil), variable.dimensions...),
						origin:        arrayOriginLocal,
					}
				} else if variable.isArray {
					state[name] = arrayValue{kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}
				} else if variable.isVariant {
					state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
				} else if variable.knownScalar || variable.isObject {
					add("VBA227", strings.TrimSpace(target)+" is not an array and cannot be erased as an array.", "Erase applies to arrays; a scalar value has no array allocation to clear.", "Erase an array variable or remove the Erase statement for this scalar value.")
				}
			}
		}
		return state, findings
	}
	// Inline `If ... Then ReDim ...` has a branch-specific state that is not
	// represented by one transfer result. Keep it conservative and avoid
	// treating the ReDim bounds themselves as an array access.
	if strings.Contains(lower, "redim ") {
		return state, findings
	}

	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		argument := strings.TrimSpace(bound[2])
		name := strings.ToLower(argument)
		value, ok := state[name]
		variable, known := variables[name]
		if !known {
			if scalarExpressionKnown(argument) {
				addWithKey(arrayBoundOperationKey(bound[1], argument, "scalar"), "VBA227", bound[1]+" cannot be used on a known scalar expression.", "LBound and UBound require an array value; this argument is a statically known scalar.", "Pass an array value to the bound function or remove the bound query.")
			}
			continue
		}
		if !variable.isArray && !variable.isVariant {
			if !variable.knownScalar && !variable.isObject {
				continue
			}
			addWithKey(arrayBoundOperationKey(bound[1], argument, "scalar"), "VBA227", bound[1]+" cannot be used on non-array "+variable.name+".", "LBound and UBound require an array value; this target is a known scalar.", "Pass an array variable to the bound function or remove the bound query.")
			continue
		}
		if !ok || value.origin == arrayOriginRangeValue {
			continue
		}
		// A Variant has no statically proven array nature.  Keep this path
		// fail-open; only a proven array (or a proven scalar handled above) is
		// actionable here.
		if variable.isVariant && !value.knownArray {
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			if arrayResumeNextCapacityProofApplies(capacityGuards, name, line) {
				// A recognized Resume Next capacity probe deliberately catches
				// this bounds failure before its fallback allocation branch.
				continue
			}
			if allocationProbeParameter != "" && strings.EqualFold(argument, allocationProbeParameter) {
				// A recognized allocation probe deliberately catches this
				// bounds failure and returns zero from its recovery label.
				continue
			}
			addWithKey(arrayBoundOperationKey(bound[1], argument, "unallocated"), "VBA227", bound[1]+" is used before "+variable.name+" is proven to be allocated.", "LBound and UBound raise a runtime error for an unallocated dynamic array and are unsafe for an unknown Variant.", "Allocate the array on every path before querying its bounds, or guard the operation explicitly.")
			continue
		}
		dimension := 1
		if argument := strings.TrimSpace(bound[3]); argument != "" {
			parsed, err := strconv.Atoi(argument)
			if err != nil {
				// A variable or expression is an unknown dimension, not an
				// invalid one. No contradiction can be proven statically.
				continue
			}
			dimension = parsed
		}
		if dimension < 1 || len(value.dimensions) > 0 && dimension > len(value.dimensions) {
			addWithKey(arrayBoundOperationKey(bound[1], argument, "bounds"), "VBA227", bound[1]+" uses invalid dimension "+strconv.Itoa(dimension)+" for "+variable.name+".", "The requested dimension is outside the array dimensions known at this point.", "Use a valid dimension number for the array, or avoid assuming a shape that is not statically known.")
		}
	}

	if match := arrayForEachRe.FindStringSubmatch(text); len(match) > 0 {
		if iterableSourceKnownInvalid(match[1], variables, state, ctx) {
			add("VBA227", strings.TrimSpace(match[1])+" is not a collection or array and cannot be used as a For Each source.", "For Each requires an iterable Collection or array value; this source is a known scalar.", "Iterate an array or Collection, or change the source expression to an iterable value.")
		}
	}

	for _, use := range arrayIndexedUsesForSource(text, variables) {
		// An empty subscript pair passes the whole array to a procedure; it is
		// not an element access whose dimension or allocation should be checked
		// at the call site. The callee owns any element-access diagnostics.
		if len(use.args) == 0 {
			continue
		}
		if arrayResumeNextCapacityIndexApplies(capacityGuards, use.name, line) {
			continue
		}
		value := state[strings.ToLower(use.name)]
		if variable, ok := variables[strings.ToLower(use.name)]; ok && variable.isVariant && !value.knownArray {
			continue
		}
		if value.origin == arrayOriginRangeValue {
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			addWithKey(arrayIndexOperationKey(use.name, "unallocated"), "VBA227", use.name+" is indexed before its array allocation is guaranteed.", "An array access can fail after Erase, before ReDim, or on a branch where allocation is not established.", "Allocate the array on every path before indexing it, or guard the access with a proven allocation check.")
			continue
		}
		if value.mayBeEmpty {
			addWithKey(arrayIndexOperationKey(use.name, "empty"), "VBA227", use.name+" is indexed while its Byte array may be empty.", "A zero-length Byte array has valid bounds queries but no element that can be indexed.", "Guard the element access with a positive length or allocate a non-empty Byte array first.")
			continue
		}
		if len(value.dimensions) > 0 && len(use.args) != len(value.dimensions) {
			addWithKey(arrayIndexOperationKey(use.name, "dimension"), "VBA227", use.name+" is indexed with "+strconv.Itoa(len(use.args))+" dimension(s), but its known shape has "+strconv.Itoa(len(value.dimensions))+".", "The number of subscripts must match the array dimensions known to the analyzer.", "Use the correct number of subscripts or revise the declared array shape.")
			continue
		}
		for i, arg := range use.args {
			if i >= len(value.dimensions) {
				break
			}
			if literal, ok := integerLiteral(arg); ok {
				dimension := value.dimensions[i]
				if dimension.lower.known && literal < dimension.lower.value || dimension.upper.known && literal > dimension.upper.value {
					addWithKey(arrayIndexOperationKey(use.name, "bounds"), "VBA227", use.name+" is indexed outside its known bounds.", "The subscript contradicts the lower or upper bound established by the declaration or ReDim.", "Use an index within the declared bounds, or establish the bounds dynamically before access.")
				}
			}
		}
	}

	if match := arrayForBoundRe.FindStringSubmatch(text); len(match) > 0 {
		name := strings.ToLower(match[2])
		if variable, ok := variables[name]; ok {
			value := state[name]
			if value.kind == arrayAllocated && value.knownArray && len(value.dimensions) > 0 && value.dimensions[0].lower.known {
				if start, ok := integerLiteral(match[1]); ok && start < value.dimensions[0].lower.value {
					add("VBA227", "The loop range assumes an inconsistent lower bound for "+variable.name+".", "The loop starts at a value different from the known lower bound of the array.", "Use LBound("+variable.name+") as the loop start.")
				}
			}
		}
	}

	if lhs, rhs, indexed, ok := arrayAssignment(text); ok {
		name := strings.ToLower(lhs)
		if variable, exists := variables[name]; exists && !variable.isArray && !variable.isVariant {
			if argument, probe := arrayAllocationProbeArgument(rhs, ctx.arrayAllocationGuards); probe {
				state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginLocal, allocationProbe: argument}
			} else if argument, probe := arraySafeBoundProbeArgument(rhs, ctx.arraySafeBoundGuards); probe {
				state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginLocal, safeBoundProbe: argument}
			} else if value, tracked := state[name]; tracked && (value.allocationProbe != "" || value.safeBoundProbe != "") {
				value.allocationProbe = ""
				value.safeBoundProbe = ""
				state[name] = value
			}
		}
		if variable, exists := variables[name]; exists && variable.isArray && variable.isObject && indexed && !strings.HasPrefix(lower, "set ") {
			code := "VBA101"
			typ := variable.typ
			callee := arrayCallName(rhs)
			if returnType := ctx.functionReturns[callee]; isObjectType(returnType) {
				code = "VBA102"
				typ = returnType
			}
			if code == "VBA101" || code == "VBA102" {
				// These are existing missing-Set rules; the lifecycle rule only
				// supplies the indexed target that the old text matcher missed.
				findings = append(findings, a.objectSetFinding(file, proc, line, code, strings.TrimSpace(lhs), typ))
			}
		}
		if !indexed {
			if value, known := arrayDictionaryMemberExpressionState(file, proc, line, rhs, variables); known {
				if variable, exists := variables[name]; exists && (variable.isArray || variable.isVariant) {
					state[name] = value
				}
			} else if value, known := arrayExpressionState(rhs, state, ctx); known {
				if value.mayBeEmpty && arrayExpressionKnownNonEmpty(file, proc, line, rhs, variables) {
					value.mayBeEmpty = false
				}
				if variable, exists := variables[name]; exists && (variable.isArray || variable.isVariant) {
					state[name] = value
				}
			} else if variable, exists := variables[name]; exists {
				if value, assigned := byteArrayStringAssignment(file, proc, line, variable, rhs, variables); assigned {
					state[name] = value
				} else if variable.isArray || variable.isVariant {
					value := arrayValue{kind: arrayUnknown, knownArray: false, origin: arrayOriginUnknown}
					// An arbitrary source assigned to a Byte array may be a
					// zero-length byte array. Keep that possibility so a later
					// successful bounds query can prove allocation without
					// incorrectly proving that an element exists.
					if isByteArrayVariable(variable) {
						value.mayBeEmpty = true
					}
					state[name] = value
				}
			}
		}
	}
	return state, findings
}

func arrayVariables(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]arrayVariable {
	decls := newDeclarationScope(file, proc)
	decls.module = moduleDecls
	base := arrayOptionBase(file)
	// The legacy line scanner intentionally ignores Const and Enum syntax.
	// Fill those gaps from the shared procedure IR so a known scalar constant
	// can still be classified as a non-iterable/non-array operation target.
	for declaration := range proc.Declarations.All() {
		key := strings.ToLower(declaration.Name)
		if key == "" {
			continue
		}
		decls.addExtraIfMissing(key, sourceDeclaration{
			Name: declaration.Name, Type: declaration.Type, Line: declaration.Range.StartLine,
			Object: declaration.IsObject, Array: declaration.IsArray,
			Fixed:      declaration.ValueShape == procedureir.ValueShapeFixedArray,
			Dimensions: parameterArrayDimensions(declaration.ArrayBounds, base),
		})
	}
	for param := range proc.Params.All() {
		array := param.ParamArray || strings.Contains(param.Type, "()") || param.ValueShape == procedureir.ValueShapeFixedArray || param.ValueShape == procedureir.ValueShapeDynamicArray
		decls.parameters[strings.ToLower(param.Name)] = sourceDeclaration{
			Name: param.Name, Type: param.Type, Array: array,
			Fixed:      param.ValueShape == procedureir.ValueShapeFixedArray,
			Dimensions: parameterArrayDimensions(param.ArrayBounds, base),
			Object:     isObjectType(param.Type), Parameter: true, ParamArray: param.ParamArray,
		}
	}
	catalog := file.ArrayVariableCatalog
	if catalog == nil {
		catalog = buildArrayVariableCatalog(file, moduleDecls)
	}
	usedModuleNames := make(map[string]struct{}, proc.Accesses.Len())
	for access := range proc.Accesses.All() {
		if name := strings.ToLower(strings.TrimSpace(access.Name)); name != "" {
			usedModuleNames[name] = struct{}{}
		}
	}
	// An empty access projection can mean a genuinely independent procedure,
	// but it can also be the only reliable view left after parser recovery or
	// an incomplete IR build. Preserve the historical fail-open behavior for
	// that boundary instead of silently dropping module scalars from array
	// classification.
	includeAllModule := proc.Accesses.Len() == 0
	moduleCapacity := len(usedModuleNames)
	for _, variable := range catalog {
		if includeAllModule || variable.isArray {
			moduleCapacity++
		}
	}
	variables := make(map[string]arrayVariable, moduleCapacity+len(decls.extra)+len(decls.local)+len(decls.parameters))
	for key, variable := range catalog {
		if !includeAllModule && !variable.isArray {
			if _, used := usedModuleNames[key]; !used {
				continue
			}
		}
		variables[key] = variable
	}
	// Module declarations are immutable and already normalized in catalog.
	// Only the procedure overlays need to be materialized here. Overlay order
	// matches declarationScope.forEach: IR extras, source locals, parameters.
	for key, decl := range decls.extra {
		variables[key] = newArrayVariable(file, decl, base)
	}
	for key, decl := range decls.local {
		variables[key] = newArrayVariable(file, decl, base)
	}
	for key, decl := range decls.parameters {
		variables[key] = newArrayVariable(file, decl, base)
	}
	return variables
}

func buildArrayVariableCatalog(file parsedFile, moduleDecls map[string]sourceDeclaration) map[string]arrayVariable {
	variables := make(map[string]arrayVariable, len(moduleDecls))
	base := arrayOptionBase(file)
	for key, decl := range moduleDecls {
		variables[key] = newArrayVariable(file, decl, base)
	}
	return variables
}

func newArrayVariable(file parsedFile, decl sourceDeclaration, base int) arrayVariable {
	typeName := strings.TrimSpace(decl.Type)
	// VBA declarations without an As clause are implicit Variant values.
	// Treat them exactly like an explicit Variant so array-sensitive rules
	// do not turn an unresolved value into a false scalar proof.
	isVariant := typeName == "" || strings.EqualFold(typeName, "Variant")
	variable := arrayVariable{name: decl.Name, typ: decl.Type, isArray: decl.Array, isVariant: isVariant, isObject: decl.Object, knownScalar: !isVariant && arrayKnownScalarType(typeName), parameter: decl.Parameter, static: decl.Static, paramArray: decl.ParamArray}
	if decl.Array {
		variable.dimensions, variable.fixed = declarationDimensions(file.Lines, decl.Line, decl.Name, base)
		if len(decl.Dimensions) > 0 {
			variable.dimensions = append([]arrayDimension(nil), decl.Dimensions...)
		}
		if decl.Fixed {
			variable.fixed = true
		}
		if len(variable.dimensions) > 0 {
			variable.fixed = true
		}
	}
	return variable
}

func isByteArrayVariable(variable arrayVariable) bool {
	return variable.isArray && isByteArrayTypeName(variable.typ)
}

func isByteArrayTypeName(typeName string) bool {
	typeName = strings.TrimSpace(typeName)
	if colon := strings.IndexByte(typeName, ':'); colon >= 0 {
		typeName = strings.TrimSpace(typeName[:colon])
	}
	return strings.EqualFold(typeName, "Byte")
}

// arrayVBA227Variables overlays the narrow declaration facts needed by the
// source-line VBA227 pass. The legacy declaration scanner is intentionally
// retained for the historical array rules, but it can span a colon-separated
// `Dim ...: ReDim ...` line and mistake the ReDim bounds for declaration
// bounds. Procedure IR identifies the declaration's dynamic shape without
// that ambiguity.
func arrayVBA227Variables(baseVariables map[string]arrayVariable, file parsedFile, proc sourceProcedure) map[string]arrayVariable {
	overlays := make([]procedureir.Declaration, 0)
	base := arrayOptionBase(file)
	for declaration := range proc.Declarations.All() {
		key := strings.ToLower(strings.TrimSpace(declaration.Name))
		if key == "" || !declaration.IsArray || declaration.ValueShape != procedureir.ValueShapeDynamicArray {
			continue
		}
		inlineRedim := arrayDeclarationHasInlineRedim(file.Lines, declaration)
		inlineStrConv := isByteArrayTypeName(declaration.Type) && arrayDeclarationHasInlineStrConvAssignment(file.Lines, declaration)
		if !inlineRedim && !inlineStrConv {
			continue
		}
		overlays = append(overlays, declaration)
	}
	if len(overlays) == 0 {
		return baseVariables
	}
	variables := make(map[string]arrayVariable, len(baseVariables))
	for key, variable := range baseVariables {
		variable.dimensions = append([]arrayDimension(nil), variable.dimensions...)
		variables[key] = variable
	}
	for _, declaration := range overlays {
		key := strings.ToLower(strings.TrimSpace(declaration.Name))
		variable, ok := variables[key]
		if !ok {
			variable = arrayVariable{
				name:        declaration.Name,
				typ:         declaration.Type,
				isArray:     true,
				isVariant:   strings.EqualFold(strings.TrimSpace(declaration.Type), "Variant"),
				isObject:    declaration.IsObject,
				knownScalar: false,
			}
		}
		variable.name = declaration.Name
		variable.typ = declaration.Type
		variable.isArray = true
		variable.isVariant = strings.EqualFold(strings.TrimSpace(declaration.Type), "Variant")
		variable.isObject = declaration.IsObject
		variable.dimensions = parameterArrayDimensions(declaration.ArrayBounds, base)
		variable.fixed = declaration.ValueShape == procedureir.ValueShapeFixedArray || len(variable.dimensions) > 0
		variables[key] = variable
	}
	return variables
}

func arrayDeclarationHasInlineRedim(lines []string, declaration procedureir.Declaration) bool {
	line := declaration.Range.StartLine
	if line < 1 || line > len(lines) {
		return false
	}
	text := strings.ToLower(normalizedCodeLine(lines[line-1]))
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return false
	}
	return strings.HasPrefix(strings.TrimSpace(text[colon+1:]), "redim ")
}

func arrayDeclarationHasInlineStrConvAssignment(lines []string, declaration procedureir.Declaration) bool {
	line := declaration.Range.StartLine
	if line < 1 || line > len(lines) {
		return false
	}
	parts := splitRangeValueSourceStatements(normalizedCodeLine(lines[line-1]))
	for _, part := range parts[1:] {
		lhs, rhs, indexed, ok := arrayAssignment(part)
		if ok && !indexed && strings.EqualFold(lhs, declaration.Name) && strings.EqualFold(arrayCallName(rhs), "strconv") {
			return true
		}
	}
	return false
}

func parameterArrayDimensions(bounds []procedureir.ArrayBound, base int) []arrayDimension {
	if len(bounds) == 0 {
		return nil
	}
	dimensions := make([]arrayDimension, 0, len(bounds))
	for _, bound := range bounds {
		lowerText, upperText := strings.TrimSpace(bound.Lower), strings.TrimSpace(bound.Upper)
		if lowerText == "" && upperText == "" {
			upperText = strings.TrimSpace(bound.Expression)
			lowerText = strconv.Itoa(base)
		}
		dimensions = append(dimensions, arrayDimension{
			lower: integerBound(lowerText), upper: integerBound(upperText),
		})
	}
	return dimensions
}

func optionBase(lines []string) int {
	for _, line := range lines {
		fields := strings.Fields(strings.ToLower(normalizedCodeLine(line)))
		if len(fields) == 3 && fields[0] == "option" && fields[1] == "base" {
			if value, err := strconv.Atoi(fields[2]); err == nil && (value == 0 || value == 1) {
				return value
			}
		}
	}
	return 0
}

func arrayOptionBase(file parsedFile) int {
	if file.ArrayOptionBaseSet {
		return file.ArrayOptionBase
	}
	return optionBase(file.Lines)
}

func declarationDimensions(lines []string, line int, name string, base int) ([]arrayDimension, bool) {
	if line < 1 || line > len(lines) {
		return nil, false
	}
	stmt := normalizedCodeLine(lines[line-1])
	m := declRe.FindStringSubmatch(stmt)
	if len(m) == 0 {
		return nil, false
	}
	for _, part := range splitArgs(m[1]) {
		candidate, _, array, _ := declarationNameAndType(part)
		if !array || !strings.EqualFold(candidate, name) {
			continue
		}
		start := strings.Index(part, "(")
		end := strings.LastIndex(part, ")")
		if start < 0 || end < start {
			return nil, false
		}
		raw := strings.TrimSpace(part[start+1 : end])
		if raw == "" {
			return nil, false
		}
		return parseArrayDimensions(raw, base), true
	}
	return nil, false
}

func parseArrayDimensions(text string, base int) []arrayDimension {
	var dimensions []arrayDimension
	for _, part := range splitArgs(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		pieces := strings.SplitN(strings.ToLower(part), " to ", 2)
		if len(pieces) == 2 {
			dimensions = append(dimensions, arrayDimension{lower: integerBound(pieces[0]), upper: integerBound(pieces[1])})
			continue
		}
		upper := integerBound(part)
		dimensions = append(dimensions, arrayDimension{lower: integerBound(strconv.Itoa(base)), upper: upper})
	}
	return dimensions
}

func parseArrayDimensionsWithConstants(text string, base int, constants map[string]int) []arrayDimension {
	var dimensions []arrayDimension
	for _, part := range splitArgs(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		pieces := strings.SplitN(strings.ToLower(part), " to ", 2)
		if len(pieces) == 2 {
			dimensions = append(dimensions, arrayDimension{lower: integerBoundWithConstants(pieces[0], constants), upper: integerBoundWithConstants(pieces[1], constants)})
			continue
		}
		dimensions = append(dimensions, arrayDimension{lower: integerBoundWithConstants(strconv.Itoa(base), constants), upper: integerBoundWithConstants(part, constants)})
	}
	return dimensions
}

// arrayIntegerConstants extends the shared range-value constant table with
// the Enum members that are valid integer bound names in VBA.  The parser is
// intentionally small and fail-open: unresolved expressions simply do not
// enter the table, so dynamic bounds remain unclassified.
func arrayIntegerConstants(file parsedFile, proc sourceProcedure, projectValues map[string]constexpr.Value, visibleNames map[string]bool) map[string]int {
	constants := rangeValueIntegerConstants(arrayIntegerModuleConstants(file), proc)
	// Project constants fill names that are not already provided by the module
	// projection, preserving the historical module-over-project precedence.
	for name, value := range projectValues {
		visibleKey := strings.ToLower(cleanIdentifier(name))
		if visibleNames != nil && !visibleNames[visibleKey] {
			continue
		}
		if value.Kind != constexpr.ValueInteger && value.Kind != constexpr.ValueLong && value.Kind != constexpr.ValueLongLong {
			continue
		}
		if integer, ok := constexpr.IntegerAsInt(value); ok {
			key := strings.ToLower(name)
			if _, exists := constants[key]; !exists {
				constants[key] = integer
			}
		}
	}
	return constants
}

func arrayIntegerModuleConstants(file parsedFile) map[string]int {
	if file.ArrayIntegerModuleConstants != nil {
		return file.ArrayIntegerModuleConstants
	}
	base := file.RangeValueModuleConstants
	if base == nil {
		base = rangeValueModuleIntegerConstants(file.Lines, file.IR)
	}
	constants := make(map[string]int, len(base))
	for name, value := range base {
		constants[name] = value
	}
	sharedValues := file.ConstantValues
	if sharedValues == nil {
		sharedValues = lint.ConstantValuesFromSource(string(file.Source), &file.IR, nil)
	}
	for name, value := range sharedValues {
		if value.Kind != constexpr.ValueInteger && value.Kind != constexpr.ValueLong && value.Kind != constexpr.ValueLongLong {
			continue
		}
		if integer, ok := constexpr.IntegerAsInt(value); ok {
			constants[strings.ToLower(name)] = integer
		}
	}
	// Keep the legacy enum fallback for hand-built parsedFile values whose
	// ConstantValues projection is absent or incomplete.
	inEnum := false
	var next *int
	conditionalDepth := 0
	for _, line := range file.Lines {
		code := strings.TrimSpace(normalizedCodeLine(line))
		lower := strings.ToLower(code)
		if strings.HasPrefix(lower, "#if ") {
			conditionalDepth++
			continue
		}
		if strings.HasPrefix(lower, "#end if") {
			if conditionalDepth > 0 {
				conditionalDepth--
			}
			continue
		}
		if conditionalDepth > 0 || strings.HasPrefix(lower, "#elseif ") || strings.HasPrefix(lower, "#else") {
			continue
		}
		if strings.HasPrefix(lower, "enum ") || strings.HasPrefix(lower, "public enum ") || strings.HasPrefix(lower, "private enum ") || strings.HasPrefix(lower, "friend enum ") {
			inEnum = true
			next = nil
			continue
		}
		if inEnum && strings.HasPrefix(lower, "end enum") {
			inEnum = false
			next = nil
			continue
		}
		if !inEnum || code == "" {
			continue
		}
		parts := strings.SplitN(code, "=", 2)
		fields := strings.Fields(parts[0])
		if len(fields) == 0 {
			continue
		}
		name := strings.TrimSpace(fields[0])
		if name == "" {
			continue
		}
		key := strings.ToLower(cleanIdentifier(name))
		if len(parts) == 2 {
			value, err := constantIntegerExpression(strings.TrimSpace(parts[1]), constants)
			if err != nil {
				next = nil
				continue
			}
			constants[key] = value
			n := value + 1
			next = &n
			continue
		}
		if next == nil {
			continue
		}
		constants[key] = *next
		n := *next + 1
		next = &n
	}
	return constants
}

func integerBound(text string) arrayBound {
	value, ok := integerLiteral(text)
	return arrayBound{known: ok, value: value, expression: canonicalArrayBoundExpression(text)}
}

func integerBoundWithConstants(text string, constants map[string]int) arrayBound {
	value, err := constantIntegerExpression(text, constants)
	if err != nil {
		return arrayBound{expression: canonicalArrayBoundExpression(text)}
	}
	return arrayBound{known: true, value: value, expression: canonicalArrayBoundExpression(text)}
}

func impossibleArrayBounds(dimensions []arrayDimension) bool {
	for _, dimension := range dimensions {
		if dimension.lower.known && dimension.upper.known && dimension.lower.value > dimension.upper.value {
			return true
		}
	}
	return false
}

func iterableSourceKnownInvalid(source string, variables map[string]arrayVariable, state arrayFlowState, ctx analysisContext) bool {
	source = strings.TrimSpace(source)
	if source == "" {
		return false
	}
	// A uniquely resolved scalar Function/Property Get return is as strong as
	// a scalar declaration.  Unknown, external, or duplicate call targets are
	// deliberately absent from functionShapes and therefore remain fail-open.
	if strings.Contains(source, "(") {
		callee := arrayCallName(source)
		if returnType, ok := ctx.functionReturns[callee]; ok && isObjectType(returnType) {
			// Known Collection/Dictionary/Object-compatible returns are valid
			// iterable sources even though their value shape is non-array.
			return false
		}
		if shape, ok := ctx.functionShapes[callee]; ok {
			switch shape {
			case procedureir.ValueShapeScalar:
				return true
			case procedureir.ValueShapeFixedArray, procedureir.ValueShapeDynamicArray:
				return false
			}
		}
	}
	// An indexed element of a statically typed scalar array is itself a
	// scalar, even though the base expression is an allocated array. Variant
	// and object arrays remain conservative because their elements may hold an
	// array or implement the Collection iteration contract.
	if match := arrayIndexedSourceRe.FindStringSubmatch(source); len(match) > 0 {
		if variable, ok := variables[strings.ToLower(match[1])]; ok && variable.isArray && !variable.isVariant && !variable.isObject && variable.knownScalar {
			return true
		}
	}
	if value, known := arrayExpressionState(source, state, ctx); known {
		if value.knownArray {
			return false
		}
		// Unknown calls and Variant values remain conservative.  A known scalar
		// expression is handled below where its declaration is available.
		if strings.Contains(source, "(") {
			return false
		}
	}
	if scalarExpressionKnown(source) {
		return true
	}
	name := source
	if dot := strings.IndexAny(name, ".("); dot >= 0 {
		name = name[:dot]
	}
	name = strings.TrimSpace(name)
	if !arrayEraseNameRe.MatchString(name) {
		// Numeric, Boolean, date, and string literals are definitely scalar; all other
		// expressions remain unknown rather than guessing their result type.
		return false
	}
	variable, ok := variables[strings.ToLower(name)]
	if !ok || variable.isArray || variable.isVariant || variable.isObject || !variable.knownScalar || dcKindFromType(variable.typ) != dcUnknown {
		return false
	}
	return true
}

func arrayKnownScalarType(typ string) bool {
	switch strings.ToLower(cleanIdentifier(strings.TrimSpace(typ))) {
	case "byte", "boolean", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal", "date", "string":
		return true
	default:
		return false
	}
}

func scalarExpressionKnown(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" {
		return false
	}
	lower := strings.ToLower(text)
	switch lower {
	case "true", "false", "nothing", "empty", "null":
		return true
	}
	if strings.HasPrefix(text, `"`) || (strings.HasPrefix(text, "#") && strings.HasSuffix(text, "#")) {
		return true
	}
	if _, ok := integerLiteral(text); ok {
		return true
	}
	if _, err := strconv.ParseFloat(strings.TrimRight(text, "%&!@$^"), 64); err == nil {
		return true
	}
	_, err := constantIntegerExpression(text, nil)
	return err == nil
}

func integerLiteral(text string) (int, bool) {
	value, err := strconv.Atoi(strings.TrimSpace(text))
	return value, err == nil
}

func preserveDimensionsSafe(previous, next []arrayDimension) bool {
	if len(next) == 1 {
		return len(previous) <= 1
	}
	if len(previous) != len(next) || len(previous) == 0 {
		return false
	}
	for i := 0; i < len(previous)-1; i++ {
		if !arrayBoundsEquivalent(previous[i].lower, next[i].lower) || !arrayBoundsEquivalent(previous[i].upper, next[i].upper) {
			return false
		}
	}
	return true
}

func arrayBoundsEquivalent(left, right arrayBound) bool {
	if left.known && right.known {
		return left.value == right.value
	}
	return left.expression != "" && left.expression == right.expression
}

func canonicalArrayBoundExpression(text string) string {
	var out strings.Builder
	inString := false
	for i := 0; i < len(text); i++ {
		ch := text[i]
		if ch == '"' {
			out.WriteByte(ch)
			if inString && i+1 < len(text) && text[i+1] == '"' {
				out.WriteByte(text[i+1])
				i++
				continue
			}
			inString = !inString
			continue
		}
		if !inString {
			if ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' {
				continue
			}
			if ch >= 'A' && ch <= 'Z' {
				ch += 'a' - 'A'
			}
		}
		out.WriteByte(ch)
	}
	return out.String()
}

func parseDirectArrayRedimClause(clause string) (directArrayRedimClause, bool) {
	clause = strings.TrimSpace(clause)
	if clause == "" || !isIdentifierStart(clause[0]) {
		return directArrayRedimClause{}, false
	}
	endName := 1
	for endName < len(clause) && isIdentifierPart(clause[endName]) {
		endName++
	}
	open := endName
	for open < len(clause) && (clause[open] == ' ' || clause[open] == '\t') {
		open++
	}
	if open >= len(clause) || clause[open] != '(' {
		return directArrayRedimClause{}, false
	}
	close := matchingParen(clause, open)
	if close < 0 {
		return directArrayRedimClause{}, false
	}
	remainder := strings.TrimSpace(clause[close+1:])
	if remainder != "" && !arrayRedimTypeSuffixRe.MatchString(remainder) {
		return directArrayRedimClause{}, false
	}
	dimensions := strings.TrimSpace(clause[open+1 : close])
	if dimensions == "" {
		return directArrayRedimClause{}, false
	}
	return directArrayRedimClause{name: clause[:endName], dimensions: dimensions}, true
}

func arrayIndexedUses(text string, variables map[string]arrayVariable) []arrayUse {
	var uses []arrayUse
	for i := 0; i < len(text); i++ {
		if !isIdentifierStart(text[i]) || (i > 0 && isIdentifierPart(text[i-1])) {
			continue
		}
		start := i
		i++
		for i < len(text) && isIdentifierPart(text[i]) {
			i++
		}
		name := text[start:i]
		// A qualified member call such as Application.OnTime(...) can
		// coincide with a local scalar named OnTime.  The member name is not
		// an array variable access; only unqualified identifiers participate
		// in this source-level array scan.
		if start > 0 && (text[start-1] == '.' || text[start-1] == '!') {
			continue
		}
		key := strings.ToLower(name)
		variable, ok := variables[key]
		if !ok || !variable.isArray && !variable.isVariant {
			continue
		}
		for i < len(text) && (text[i] == ' ' || text[i] == '\t') {
			i++
		}
		if i >= len(text) || text[i] != '(' {
			continue
		}
		end := matchingParen(text, i)
		if end < 0 {
			continue
		}
		uses = append(uses, arrayUse{name: name, args: splitArgs(text[i+1 : end])})
		i = end
	}
	return uses
}

// arrayIndexedUsesForSource scans source-line CFG text after removing syntax
// that cannot perform an array access. Recovered blocks may retain string
// literals and comments, so feeding their raw text to the VBA227 lane would
// mistake prose such as "clipboard data (" for an indexed use. Keep the
// primitive scanner raw for runtime projections that have their own source
// normalization contract.
func arrayIndexedUsesForSource(text string, variables map[string]arrayVariable) []arrayUse {
	return arrayIndexedUses(maskStringLiterals(gui.StripComment(text)), variables)
}

func matchingParen(text string, start int) int {
	depth := 0
	inString := false
	for i := start; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString {
				depth--
				if depth == 0 {
					return i
				}
			}
		}
	}
	return -1
}

func isIdentifierStart(char byte) bool {
	return char == '_' || char >= 'A' && char <= 'Z' || char >= 'a' && char <= 'z'
}
func isIdentifierPart(char byte) bool { return isIdentifierStart(char) || char >= '0' && char <= '9' }

func arrayAssignment(text string) (lhs, rhs string, indexed, ok bool) {
	trimmed := strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(trimmed), "set ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	if strings.HasPrefix(strings.ToLower(trimmed), "let ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	for i := 0; i < len(trimmed); i++ {
		if trimmed[i] != '=' || i > 0 && (trimmed[i-1] == '<' || trimmed[i-1] == '>' || trimmed[i-1] == '=') {
			continue
		}
		lhs = strings.TrimSpace(trimmed[:i])
		rhs = strings.TrimSpace(trimmed[i+1:])
		if lhs == "" || strings.HasPrefix(strings.ToLower(lhs), "if ") {
			return "", "", false, false
		}
		if open := strings.Index(lhs, "("); open >= 0 && strings.HasSuffix(lhs, ")") {
			wholeArray := strings.TrimSpace(lhs[open+1:len(lhs)-1]) == ""
			return cleanIdentifier(lhs[:open]), rhs, !wholeArray, true
		}
		return cleanIdentifier(lhs), rhs, false, true
	}
	return "", "", false, false
}

func arrayExpressionState(rhs string, state arrayFlowState, ctx analysisContext) (arrayValue, bool) {
	lower := strings.ToLower(strings.TrimSpace(rhs))
	if strings.HasPrefix(lower, "array(") || lower == "array" {
		shape := []arrayDimension{{}}
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: shape, preserveShape: shape, origin: arrayOriginLocal}, true
	}
	if strings.Contains(lower, ".value") || strings.Contains(lower, ".value2") {
		return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginRangeValue}, true
	}
	name := arrayCallName(rhs)
	if name == "split" || name == "filter" {
		shape := []arrayDimension{{}}
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: shape, preserveShape: shape, origin: arrayOriginLocal}, true
	}
	// Keep StrConv-to-Byte-array assignments in the type-aware transfer below;
	// it can distinguish a known non-empty String from an unknown one.
	if name == "strconv" {
		return arrayValue{}, false
	}
	if arrayByteArrayReadRe.MatchString(rhs) {
		return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}, true
	}
	if value, ok := state[name]; ok && value.kind == arrayAllocated && value.knownArray {
		return value, true
	}
	if value, ok := ctx.arrayReturns[name]; ok {
		_, hasLowerBound := arrayValueKnownLowerBound(value)
		if value.knownArray || hasLowerBound {
			return value, true
		}
	}
	if name != "" && strings.Contains(rhs, "(") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	if strings.EqualFold(strings.TrimSpace(rhs), "empty") || strings.EqualFold(strings.TrimSpace(rhs), "nothing") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	return arrayValue{}, false
}

// arrayVBA227AttachReturnProvenance consumes the narrow conditional facts in
// an array-return summary at the call site. The summary records a formal
// parameter; this step maps it to the actual argument and keeps the fact
// path-sensitive until the caller proves the corresponding condition.
func arrayVBA227AttachReturnProvenance(state arrayFlowState, text string, ctx analysisContext, variables map[string]arrayVariable, constants map[string]int) arrayFlowState {
	lhs, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return state
	}
	callee := arrayCallName(rhs)
	if callee == "" || ctx.functionAmbiguous[callee] {
		return state
	}
	summary, ok := ctx.arrayReturns[callee]
	if !ok || summary.returnNonEmptyArrayParameter == "" && summary.returnPositiveScalarParameter == "" && summary.returnDescriptorSourceParameter == "" {
		return state
	}
	target := strings.ToLower(cleanIdentifier(lhs))
	variable, knownVariable := variables[target]
	value, knownValue := state[target]
	if !knownVariable || !knownValue || !variable.isArray && !variable.isVariant {
		return state
	}
	signature, ok := ctx.procedures[callee]
	if !ok {
		return state
	}
	arguments, ok := arrayReturnCallArguments(rhs, signature)
	if !ok {
		return state
	}
	updated := value
	changed := false
	if formal := summary.returnPositiveScalarParameter; formal != "" {
		if index, found := arrayFormalParameterIndexFromSignature(signature, formal); found && index < len(arguments) {
			if length, err := constantIntegerExpression(strings.TrimSpace(arguments[index]), constants); err == nil && length > 0 {
				updated.kind = arrayAllocated
				updated.knownArray = true
				updated.mayBeEmpty = false
				updated.returnPositiveScalarParameter = ""
				updated.returnNonEmptyArrayParameter = ""
				changed = true
			}
		}
	}
	if formal := summary.returnNonEmptyArrayParameter; formal != "" {
		if index, found := arrayFormalParameterIndexFromSignature(signature, formal); found && index < len(arguments) {
			if source := directArrayArgumentName(arguments[index]); source != "" {
				sourceVariable, sourceKnown := variables[source]
				if sourceKnown && sourceVariable.isArray {
					updated.nonEmptySource = source
					updated.returnNonEmptyArrayParameter = ""
					updated.returnPositiveScalarParameter = ""
					changed = true
				}
			}
		}
	}
	if summary.returnDescriptorSourceParameter != "" {
		updated = arrayDescriptorReturnCallValue(updated, summary, arguments, signature, constants)
		updated.returnDescriptorSourceParameter = ""
		updated.returnDescriptorStartParameter = ""
		updated.returnDescriptorLengthParameter = ""
		updated.returnDescriptorLowerParameter = ""
		changed = true
	}
	if !changed {
		return state
	}
	updatedState := cloneArrayState(state)
	updatedState[target] = updated
	return updatedState
}

func arraySimpleCallArguments(text string) ([]string, bool) {
	text = strings.TrimSpace(text)
	open := firstParenOutsideString(text)
	if open < 0 {
		return nil, false
	}
	close := matchingParen(text, open)
	if close < 0 || strings.TrimSpace(text[close+1:]) != "" {
		return nil, false
	}
	return splitArgs(text[open+1 : close]), true
}

var arrayNamedArgumentRe = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*:=\s*(.*?)\s*$`)

// arrayReturnCallArguments maps positional and named actuals to a procedure
// signature and fills omitted optional arguments with their declared defaults.
// Return summaries need this small binding layer because a call such as
// StringToIntegers("ABC", outLowBound:=5) cannot be interpreted by position.
func arrayReturnCallArguments(text string, signature procedureSignature) ([]string, bool) {
	raw, ok := arraySimpleCallArguments(text)
	if !ok {
		return nil, false
	}
	arguments := make([]string, signature.Params.Len())
	assigned := make([]bool, len(arguments))
	nextPositional := 0
	for _, rawArgument := range raw {
		argument := strings.TrimSpace(rawArgument)
		if match := arrayNamedArgumentRe.FindStringSubmatch(argument); len(match) == 3 {
			index, found := arrayFormalParameterIndexFromSignature(signature, match[1])
			if !found || assigned[index] {
				return nil, false
			}
			arguments[index] = strings.TrimSpace(match[2])
			assigned[index] = true
			continue
		}
		for nextPositional < len(assigned) && assigned[nextPositional] {
			nextPositional++
		}
		if nextPositional >= len(arguments) {
			return nil, false
		}
		if argument == "" {
			nextPositional++
			continue
		}
		arguments[nextPositional] = argument
		assigned[nextPositional] = true
		nextPositional++
	}
	for index, parameter := range signature.Params.AllIndexed() {
		if assigned[index] {
			continue
		}
		if !parameter.Optional {
			return nil, false
		}
		if parameter.HasDefault {
			arguments[index] = parameter.Default
		}
	}
	return arguments, true
}

func arrayDescriptorReturnCallValue(value, summary arrayValue, arguments []string, signature procedureSignature, constants map[string]int) arrayValue {
	argument := func(formal string) (string, bool) {
		index, ok := arrayFormalParameterIndexFromSignature(signature, formal)
		if !ok || index >= len(arguments) {
			return "", false
		}
		return strings.TrimSpace(arguments[index]), true
	}
	source, sourceOK := argument(summary.returnDescriptorSourceParameter)
	startText, startOK := argument(summary.returnDescriptorStartParameter)
	lengthText, lengthOK := argument(summary.returnDescriptorLengthParameter)
	lowerText, lowerOK := argument(summary.returnDescriptorLowerParameter)
	if !sourceOK || !startOK || !lengthOK || !lowerOK {
		return value
	}
	start, startErr := constantIntegerExpression(startText, constants)
	length, lengthErr := constantIntegerExpression(lengthText, constants)
	lower, lowerErr := constantIntegerExpression(lowerText, constants)
	if startErr != nil || lengthErr != nil || lowerErr != nil || start < 1 || length < -1 {
		return value
	}

	count, countKnown := 0, false
	if length == 0 {
		count, countKnown = 0, true
	} else if sourceLength, known := arrayStringExpressionKnownLength(source); known {
		count = length
		if length == -1 || start+length-1 > sourceLength {
			count = sourceLength - start + 1
			if count < 0 {
				count = 0
			}
		}
		countKnown = true
	}
	dimension := arrayDimension{lower: arrayBound{known: true, value: lower}}
	if countKnown {
		dimension.upper = arrayBound{known: true, value: lower + count - 1}
		value.mayBeEmpty = count == 0
	} else {
		// A descriptor-backed return always has a valid SAFEARRAY descriptor,
		// but an unknown input length may still produce zero elements.
		value.mayBeEmpty = true
	}
	value.dimensions = []arrayDimension{dimension}
	value.preserveShape = append([]arrayDimension(nil), value.dimensions...)
	return value
}

func arrayStringExpressionKnownLength(expression string) (int, bool) {
	expression = strings.TrimSpace(expression)
	if strings.EqualFold(expression, "vbNullString") {
		return 0, true
	}
	if length, ok := arrayStringLiteralLength(expression); ok {
		return length, true
	}
	if !strings.EqualFold(arrayCallName(expression), "strconv") {
		return 0, false
	}
	open := firstParenOutsideString(expression)
	close := matchingParen(expression, open)
	if open < 0 || close < 0 || strings.TrimSpace(expression[close+1:]) != "" {
		return 0, false
	}
	arguments := splitArgs(expression[open+1 : close])
	if len(arguments) == 0 {
		return 0, false
	}
	return arrayStringExpressionKnownLength(arguments[0])
}

func arrayStringLiteralLength(expression string) (int, bool) {
	if len(expression) < 2 || expression[0] != '"' || expression[len(expression)-1] != '"' {
		return 0, false
	}
	length := 0
	for index := 1; index < len(expression)-1; index++ {
		if expression[index] >= 0x80 {
			return 0, false
		}
		if expression[index] == '"' {
			if index+1 >= len(expression)-1 || expression[index+1] != '"' {
				return 0, false
			}
			index++
		}
		length++
	}
	return length, true
}

func arrayFormalParameterIndexFromSignature(signature procedureSignature, name string) (int, bool) {
	for index, parameter := range signature.Params.AllIndexed() {
		if strings.EqualFold(strings.TrimSpace(parameter.Name), strings.TrimSpace(name)) {
			return index, true
		}
	}
	return 0, false
}

var (
	arrayStringNonEmptyBlockRe = regexp.MustCompile(`(?i)^\s*if\s+(?:[a-z_]\w*\.)?len\s*\(\s*([a-z_]\w*)\s*\)\s*>\s*0\s+then\s*$`)
	arrayStringLengthAssignRe  = regexp.MustCompile(`(?i)^\s*([a-z_]\w*)\s*=\s*(?:[a-z_]\w*\.)?len\s*\(\s*([a-z_]\w*)\s*\)\s*$`)
	arrayStringEmptyExitRe     = regexp.MustCompile(`(?i)^\s*if\s+([a-z_]\w*)\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayStringEmptyLenExitRe  = regexp.MustCompile(`(?i)^\s*if\s+(?:[a-z_]\w*\.)?len\s*\(\s*([a-z_]\w*)\s*\)\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
)

func arrayExpressionKnownNonEmpty(file parsedFile, proc sourceProcedure, line int, rhs string, variables map[string]arrayVariable) bool {
	open := firstParenOutsideString(strings.TrimSpace(rhs))
	if open < 0 {
		return false
	}
	close := matchingParen(rhs, open)
	if close < 0 || strings.TrimSpace(rhs[close+1:]) != "" {
		return false
	}
	for _, argument := range splitArgs(rhs[open+1 : close]) {
		name := strings.ToLower(cleanIdentifier(strings.TrimSpace(argument)))
		variable, ok := variables[name]
		if !ok || variable.isArray || variable.isVariant || !strings.EqualFold(strings.TrimSpace(variable.typ), "String") {
			continue
		}
		if arrayStringIsKnownNonEmpty(file, proc, line, name) {
			return true
		}
	}
	return false
}

// VBA copies a String into a Byte array. A non-empty source establishes a
// usable allocation; vbNullString establishes a known empty allocation whose
// bounds may be queried but whose elements must not be indexed.
func byteArrayStringAssignment(file parsedFile, proc sourceProcedure, line int, variable arrayVariable, rhs string, variables map[string]arrayVariable) (arrayValue, bool) {
	allocated := func(mayBeEmpty bool) (arrayValue, bool) {
		return arrayValue{kind: arrayAllocated, knownArray: true, mayBeEmpty: mayBeEmpty, origin: arrayOriginLocal}, true
	}
	if !variable.isArray || !strings.EqualFold(strings.TrimSpace(variable.typ), "Byte") {
		return arrayValue{}, false
	}
	rhs = strings.TrimSpace(rhs)
	if strings.EqualFold(rhs, "vbNullString") {
		return allocated(true)
	}
	if strings.EqualFold(arrayCallName(rhs), "strconv") {
		open := firstParenOutsideString(rhs)
		close := matchingParen(rhs, open)
		if open < 0 || close < 0 || strings.TrimSpace(rhs[close+1:]) != "" {
			return arrayValue{}, false
		}
		arguments := splitArgs(rhs[open+1 : close])
		if len(arguments) == 0 {
			return arrayValue{}, false
		}
		if strings.EqualFold(strings.TrimSpace(arguments[0]), "vbNullString") {
			return allocated(true)
		}
		if arrayStringExpressionKnownNonEmpty(file, proc, line, arguments[0], variables) {
			return allocated(false)
		}
		return arrayValue{}, false
	}
	if strings.HasPrefix(rhs, `"`) {
		if len(rhs) <= 1 || strings.HasPrefix(rhs, `""`) {
			return arrayValue{}, false
		}
		return allocated(false)
	}
	if !arrayEraseNameRe.MatchString(rhs) {
		return arrayValue{}, false
	}
	source, ok := variables[strings.ToLower(rhs)]
	if !ok || source.isArray || source.isVariant || !strings.EqualFold(strings.TrimSpace(source.typ), "String") {
		return arrayValue{}, false
	}
	if !arrayStringIsKnownNonEmpty(file, proc, line, rhs) {
		return arrayValue{}, false
	}
	return allocated(false)
}

func arrayStringExpressionKnownNonEmpty(file parsedFile, proc sourceProcedure, line int, expression string, variables map[string]arrayVariable) bool {
	expression = strings.TrimSpace(expression)
	if expression == "" || strings.EqualFold(expression, "vbNullString") {
		return false
	}
	if arrayStringExpressionHasNonEmptyLiteral(expression) {
		return true
	}
	name := cleanIdentifier(expression)
	if name != expression {
		return false
	}
	variable, ok := variables[strings.ToLower(name)]
	if !ok || variable.isArray || variable.isVariant || !strings.EqualFold(strings.TrimSpace(variable.typ), "String") {
		return false
	}
	if arrayStringIsKnownNonEmpty(file, proc, line, name) {
		return true
	}
	return arrayStringVariableHasNonEmptyAssignment(file, proc, line, name)
}

func arrayStringExpressionHasNonEmptyLiteral(expression string) bool {
	expression = strings.TrimSpace(expression)
	for strings.HasPrefix(expression, "(") {
		close := matchingParen(expression, 0)
		if close != len(expression)-1 {
			break
		}
		expression = strings.TrimSpace(expression[1:close])
	}
	for _, operand := range splitStringConcatenation(expression) {
		if arrayStringLiteralHasValue(operand) || arrayStringNonEmptyConstant(operand) {
			return true
		}
	}
	return false
}

func arrayStringNonEmptyConstant(operand string) bool {
	switch strings.ToLower(strings.TrimSpace(operand)) {
	case "vbnullchar", "vbcrlf", "vbcr", "vblf", "vbtab", "vbverticaltab", "vbformfeed", "vbnewline":
		return true
	default:
		return false
	}
}

func arrayStringVariableHasNonEmptyAssignment(file parsedFile, proc sourceProcedure, line int, source string) bool {
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), line-1)
	depth := 0
	assigned := false
	for index := start; index < end; index++ {
		for _, statement := range splitRangeValueSourceStatements(strings.TrimSpace(normalizedCodeLine(file.Lines[index]))) {
			text := strings.TrimSpace(statement)
			if delta := arrayStringBlockBoundary(text); delta < 0 {
				if depth > 0 {
					depth--
				}
				continue
			}
			lhs, rhs, indexed, ok := arrayAssignment(text)
			if ok && !indexed && strings.EqualFold(lhs, source) {
				assigned = depth == 0 && arrayStringExpressionHasNonEmptyLiteral(rhs)
			}
			if delta := arrayStringBlockBoundary(text); delta > 0 {
				depth++
			}
		}
	}
	return assigned
}

func splitStringConcatenation(expression string) []string {
	var parts []string
	start := 0
	inString := false
	depth := 0
	for index := 0; index < len(expression); index++ {
		switch expression[index] {
		case '"':
			if inString && index+1 < len(expression) && expression[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '&':
			if !inString && depth == 0 {
				parts = append(parts, strings.TrimSpace(expression[start:index]))
				start = index + 1
			}
		}
	}
	return append(parts, strings.TrimSpace(expression[start:]))
}

func arrayStringLiteralHasValue(operand string) bool {
	operand = strings.TrimSpace(operand)
	if len(operand) < 2 || operand[0] != '"' || operand[len(operand)-1] != '"' {
		return false
	}
	length := 0
	for index := 1; index < len(operand)-1; index++ {
		if operand[index] == '"' {
			if index+1 < len(operand)-1 && operand[index+1] == '"' {
				length++
				index++
				continue
			}
			return false
		}
		length++
	}
	return length > 0
}

func arrayStringBlockBoundary(text string) int {
	lower := strings.ToLower(strings.TrimSpace(text))
	switch {
	case lower == "end if", lower == "end select", lower == "end with", lower == "loop", lower == "wend", strings.HasPrefix(lower, "next"):
		return -1
	case strings.HasPrefix(lower, "if ") && strings.HasSuffix(lower, " then"), strings.HasPrefix(lower, "for "), strings.HasPrefix(lower, "do "), lower == "do", strings.HasPrefix(lower, "select case "), strings.HasPrefix(lower, "with "), strings.HasPrefix(lower, "while "):
		return 1
	default:
		return 0
	}
}

func arrayStringIsKnownNonEmpty(file parsedFile, proc sourceProcedure, line int, source string) bool {
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), line-1)
	lengthVariables := map[string]bool{}
	for index := start; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		if match := arrayStringEmptyLenExitRe.FindStringSubmatch(text); len(match) == 2 && strings.EqualFold(match[1], source) {
			return true
		}
		if match := arrayStringLengthAssignRe.FindStringSubmatch(text); len(match) == 3 && strings.EqualFold(match[2], source) {
			lengthVariables[strings.ToLower(match[1])] = true
		}
		if match := arrayStringEmptyExitRe.FindStringSubmatch(text); len(match) == 2 && lengthVariables[strings.ToLower(match[1])] {
			return true
		}
	}

	depth := 0
	guardDepth := 0
	for index := start; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		lower := strings.ToLower(text)
		if lower == "end if" {
			if guardDepth == depth {
				guardDepth = 0
			}
			if depth > 0 {
				depth--
			}
			continue
		}
		if match := arrayStringNonEmptyBlockRe.FindStringSubmatch(text); len(match) == 2 && strings.EqualFold(match[1], source) {
			depth++
			guardDepth = depth
			continue
		}
		if strings.HasPrefix(lower, "if ") && strings.HasSuffix(lower, " then") {
			depth++
		}
	}
	return guardDepth > 0
}

func arrayCallName(text string) string {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		if close := matchingParen(trimmed, open); close >= 0 && strings.TrimSpace(trimmed[close+1:]) == "" {
			name := strings.TrimSpace(trimmed[:open])
			if dot := strings.LastIndex(name, "."); dot >= 0 {
				name = name[dot+1:]
			}
			return strings.ToLower(cleanIdentifier(name))
		}
	}
	name := strings.TrimSpace(lastName(trimmed))
	if open := strings.Index(name, "("); open >= 0 {
		name = name[:open]
	}
	return strings.ToLower(cleanIdentifier(name))
}

func firstParenOutsideString(text string) int {
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				return i
			}
		}
	}
	return -1
}

// inferArrayAllocationGuards recognizes a deliberately small helper
// contract: a scalar Function with one array or Variant parameter returns a
// positive UBound-based length on its normal path and returns zero from an
// On Error GoTo recovery label. That contract is enough to prove the positive
// branch of a direct call, while arbitrary helper functions remain unknown.
func inferArrayAllocationGuards(files []parsedFile) map[string]bool {
	candidates := map[string][]string{}
	procedureNames := map[string]int{}
	recognizedNames := map[string]int{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(proc.Name)
			if name != "" {
				procedureNames[name]++
			}
			parameter, ok := arrayAllocationGuardParameter(proc)
			if !ok {
				parameter, ok = arraySafeArrayPointerLengthGuardParameter(file, proc)
			}
			if !ok {
				parameter, ok = arrayDimensionCountGuardParameter(proc)
			}
			if !ok {
				continue
			}
			candidates[name] = append(candidates[name], parameter)
			recognizedNames[name]++
		}
	}
	guards := map[string]bool{}
	for name := range candidates {
		// Private procedures are module-scoped, so the same helper name can
		// legitimately occur in multiple modules. Keep a bare-name guard only
		// when every procedure with that name has the same narrow guard shape;
		// an unrelated duplicate remains conservative.
		if name == "" || recognizedNames[name] != procedureNames[name] {
			continue
		}
		guards[name] = true
	}
	return guards
}

// inferArraySafeBoundGuards recognizes helpers that return UBound(array) on
// their normal path and a negative sentinel after catching an unallocated
// array. A nonnegative result therefore proves both that UBound succeeded and
// that the array has an index at zero or above.
func inferArraySafeBoundGuards(files []parsedFile) map[string]bool {
	candidates := map[string][]string{}
	procedureNames := map[string]int{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(proc.Name)
			if name != "" {
				procedureNames[name]++
			}
			parameter, ok := arraySafeBoundGuardParameter(proc)
			if !ok {
				continue
			}
			candidates[name] = append(candidates[name], parameter)
		}
	}
	guards := map[string]bool{}
	for name, parameters := range candidates {
		if name == "" || procedureNames[name] != 1 || len(parameters) != 1 {
			continue
		}
		guards[name] = true
	}
	return guards
}

func arraySafeBoundGuardParameter(proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	variantParameter := strings.EqualFold(cleanIdentifier(strings.TrimSpace(parameter.Type)), "variant")
	if parameter.Name == "" || (!parameterIsArray(parameter) && !variantParameter) || proc.Name == "" {
		return "", false
	}

	errorLabel := ""
	recovery := false
	hasRecovery := false
	foundRecoveryLabel := false
	normalReturns := 0
	recoveryNegativeReturns := 0
	invalidReturn := false
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if match := arrayOnErrorGotoRe.FindStringSubmatch(text); len(match) == 2 {
			errorLabel = strings.ToLower(match[1])
			hasRecovery = true
		}
		if match := arrayLabelRe.FindStringSubmatch(text); len(match) == 2 {
			recovery = errorLabel != "" && strings.EqualFold(match[1], errorLabel)
			if recovery {
				foundRecoveryLabel = true
			}
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		switch {
		case arrayUpperBoundExpressionMatches(rhs, parameter.Name):
			normalReturns++
		case recovery && strings.TrimSpace(rhs) == "-1":
			recoveryNegativeReturns++
		default:
			invalidReturn = true
		}
	}
	if !hasRecovery || !foundRecoveryLabel || normalReturns != 1 || recoveryNegativeReturns != 1 || invalidReturn {
		return "", false
	}
	return strings.ToLower(parameter.Name), true
}

func arrayUpperBoundExpressionMatches(rhs, parameter string) bool {
	compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
	return compact == "ubound("+strings.ToLower(parameter)+")"
}

func arrayAllocationGuardParameter(proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	variantParameter := strings.EqualFold(cleanIdentifier(strings.TrimSpace(parameter.Type)), "variant")
	if parameter.Name == "" || (!parameterIsArray(parameter) && !variantParameter) || proc.Name == "" {
		return "", false
	}

	errorLabel := ""
	recovery := false
	hasRecovery := false
	foundRecoveryLabel := false
	positiveReturns := 0
	recoveryZeroReturns := 0
	invalidReturn := false
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if match := arrayOnErrorGotoRe.FindStringSubmatch(text); len(match) == 2 {
			errorLabel = strings.ToLower(match[1])
			hasRecovery = true
		}
		if match := arrayLabelRe.FindStringSubmatch(text); len(match) == 2 {
			recovery = errorLabel != "" && strings.EqualFold(match[1], errorLabel)
			if recovery {
				foundRecoveryLabel = true
			}
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		switch {
		case arrayLengthExpressionMatches(rhs, parameter.Name):
			positiveReturns++
		case recovery && strings.TrimSpace(rhs) == "0":
			recoveryZeroReturns++
		default:
			invalidReturn = true
		}
	}
	// A typed VBA Function defaults its return value to zero.  A recovery
	// label that falls through to End Function therefore has the same
	// allocation-probe contract as an explicit `FunctionName = 0` assignment,
	// provided the recovery label was actually found and no other return
	// assignment invalidated the shape above.
	if !hasRecovery || !foundRecoveryLabel || positiveReturns != 1 || recoveryZeroReturns > 1 || invalidReturn {
		return "", false
	}
	return strings.ToLower(parameter.Name), true
}

// arraySafeArrayPointerLengthGuardParameter recognizes a scalar helper that
// returns the length of a Byte array after the low-level SAFEARRAY descriptor
// guard. Unlike the ordinary On Error-based allocation probe, this form uses
// the function's default zero return on either early-exit path. Require the
// returned expression to be derived from preceding LBound and UBound
// assignments so an arbitrary pointer check cannot become an allocation
// contract.
func arraySafeArrayPointerLengthGuardParameter(file parsedFile, proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	if parameter.Name == "" || !parameterIsArray(parameter) || proc.Name == "" {
		return "", false
	}
	variables := arrayVariables(file, proc, file.moduleDecls())
	parameterName := strings.ToLower(cleanIdentifier(parameter.Name))
	variable, known := variables[parameterName]
	if !known || !variable.isArray || !isByteArrayVariable(variable) {
		return "", false
	}
	guardLine := 0
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		if target, ok := arraySafeArrayPointerGuardTarget(file, proc, line, file.Lines[line-1], variables); ok && target == parameterName {
			guardLine = line
			break
		}
	}
	if guardLine == 0 {
		return "", false
	}
	lowerName := ""
	upperName := ""
	returnCount := 0
	invalidReturn := false
	for line := guardLine + 1; line <= proc.EndLine && line <= len(file.Lines); line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") {
			continue
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed {
			continue
		}
		compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
		switch {
		case compact == "lbound("+parameterName+")":
			if lowerName != "" {
				return "", false
			}
			lowerName = strings.ToLower(cleanIdentifier(lhs))
		case compact == "ubound("+parameterName+")":
			if upperName != "" {
				return "", false
			}
			upperName = strings.ToLower(cleanIdentifier(lhs))
		case strings.EqualFold(cleanIdentifier(lhs), proc.Name):
			returnCount++
			expected := upperName + "-" + lowerName + "+1"
			if lowerName == "" || upperName == "" || compact != expected {
				invalidReturn = true
			}
		}
	}
	if lowerName == "" || upperName == "" || returnCount != 1 || invalidReturn {
		return "", false
	}
	return parameterName, true
}

// arrayDimensionCountGuardParameter recognizes the helper shape used by
// GetArrayDimsCount: it probes successive LBound dimensions under an error
// handler and returns the last successful dimension number minus one. The
// result is zero when the first probe fails, so a caller branch that proves
// the result is one has also proved that its input is an allocated 1D array.
// Keep this separate from the ordinary length probe so an arbitrary scalar
// helper returning `someValue - 1` cannot become an allocation contract.
func arrayDimensionCountGuardParameter(proc sourceProcedure) (string, bool) {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return "", false
	}
	if !arrayKnownScalarType(proc.ReturnType) || isObjectType(proc.ReturnType) || proc.Params.Len() != 1 {
		return "", false
	}
	parameter := proc.Params.valueAt(0)
	variantParameter := strings.EqualFold(cleanIdentifier(strings.TrimSpace(parameter.Type)), "variant")
	if parameter.Name == "" || (!parameterIsArray(parameter) && !variantParameter) || proc.Name == "" {
		return "", false
	}

	errorLabel := ""
	recovery := false
	foundRecoveryLabel := false
	loopVariable := ""
	hasDimensionProbe := false
	countReturns := 0
	invalidReturn := false
	for statement := range proc.Statements.All() {
		rawText := strings.TrimSpace(statement.Text)
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if match := arrayOnErrorGotoRe.FindStringSubmatch(text); len(match) == 2 {
			errorLabel = strings.ToLower(match[1])
		}
		if match := arrayLabelRe.FindStringSubmatch(text); len(match) == 2 {
			recovery = errorLabel != "" && strings.EqualFold(match[1], errorLabel)
			if recovery {
				foundRecoveryLabel = true
			}
		}
		loopText := rawText
		if newline := strings.IndexAny(loopText, "\r\n"); newline >= 0 {
			loopText = strings.TrimSpace(loopText[:newline])
		}
		if match := arrayDimensionCountLoopRe.FindStringSubmatch(loopText); len(match) == 2 {
			loopVariable = strings.ToLower(cleanIdentifier(match[1]))
		}
		for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
			if !strings.EqualFold(bound[1], "lbound") || !strings.EqualFold(cleanIdentifier(bound[2]), parameter.Name) || loopVariable == "" || !strings.EqualFold(cleanIdentifier(bound[3]), loopVariable) {
				continue
			}
			hasDimensionProbe = true
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		if recovery && arrayDimensionCountExpressionMatches(rhs, loopVariable) {
			countReturns++
		} else {
			invalidReturn = true
		}
	}
	if errorLabel == "" || !foundRecoveryLabel || loopVariable == "" || !hasDimensionProbe || countReturns != 1 || invalidReturn {
		return "", false
	}
	return strings.ToLower(parameter.Name), true
}

func arrayDimensionCountExpressionMatches(rhs, dimension string) bool {
	compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
	return compact == strings.ToLower(dimension)+"-1"
}

func parameterIsArray(parameter parameterInfo) bool {
	return parameter.ParamArray || strings.Contains(parameter.Type, "()") || parameter.ValueShape == procedureir.ValueShapeFixedArray || parameter.ValueShape == procedureir.ValueShapeDynamicArray
}

func arrayLengthExpressionMatches(rhs, parameter string) bool {
	compact := strings.Join(strings.Fields(strings.ToLower(rhs)), "")
	parameter = strings.ToLower(parameter)
	return compact == "ubound("+parameter+")-lbound("+parameter+")+1" || compact == "ubound("+parameter+")+1"
}

// inferArrayReturnSummaries intentionally summarizes only normal, directly
// observed return assignments. Recognized allocation guards refine the normal
// branch, while a definitely failing ReDim without local error handling does
// not contribute a normal path. A missing assignment, mixed assignment kinds,
// duplicate procedure names, and recursive/external calls remain unknown.
// Dependencies are solved to a fixed point so a private array-returning helper
// may be declared after its caller without turning an unproved VBA function
// contract into an allocation guarantee.
func arrayReturnSummaryDuplicateNames(procedures []sourceProcedure) map[string]bool {
	counts := make(map[string]int, len(procedures))
	for _, procedure := range procedures {
		name := strings.ToLower(strings.TrimSpace(procedure.Name))
		if name != "" {
			counts[name]++
		}
	}
	duplicates := make(map[string]bool)
	for name, count := range counts {
		if count > 1 {
			duplicates[name] = true
		}
	}
	return duplicates
}

func inferDocumentedArrayReturnSummaries(files []parsedFile) map[string]arrayValue {
	returns := map[string]arrayValue{}
	for _, file := range files {
		for procedure := range file.procedureView().All() {
			if procedure.ProcedureKind != procedureir.ProcedureFunction && procedure.ProcedureKind != procedureir.ProcedurePropertyGet || procedure.Name == "" {
				continue
			}
			if !arrayProcedureDocumentsArray(file, procedure) || !arrayProcedureHasReturnAllocation(file, procedure) {
				continue
			}
			returns[arrayProcedureKey(procedure)] = arrayValue{
				kind:       arrayAllocated,
				knownArray: true,
				mayBeEmpty: true,
				origin:     arrayOriginLocal,
			}
		}
	}
	return returns
}

func inferDocumentedNonEmptyArrayReturnSummaries(files []parsedFile) map[string]arrayValue {
	returns := map[string]arrayValue{}
	for _, file := range files {
		for procedure := range file.procedureView().All() {
			if procedure.ProcedureKind != procedureir.ProcedureFunction && procedure.ProcedureKind != procedureir.ProcedurePropertyGet || procedure.Name == "" {
				continue
			}
			if !arrayProcedureDocumentsArray(file, procedure) || !arrayProcedureHasNonEmptyReturnAllocation(file, procedure) {
				continue
			}
			returns[arrayProcedureKey(procedure)] = arrayValue{
				kind:       arrayAllocated,
				knownArray: true,
				origin:     arrayOriginLocal,
			}
		}
	}
	return returns
}

// inferDocumentedArrayReturnLowerBoundSummaries records a weaker contract for
// documented helpers that may return an unallocated array on an empty input,
// but consistently allocate that array from a known lower bound. The caller
// still receives a VBA227 for a direct UBound on the possibly-unallocated
// result; the lower bound is used only after that query has successfully
// reached a For body.
func inferDocumentedArrayReturnLowerBoundSummaries(files []parsedFile) map[string]arrayValue {
	returns := map[string]arrayValue{}
	for _, file := range files {
		for procedure := range file.procedureView().All() {
			if procedure.ProcedureKind != procedureir.ProcedureFunction && procedure.ProcedureKind != procedureir.ProcedurePropertyGet || procedure.Name == "" {
				continue
			}
			if !arrayProcedureDocumentsArray(file, procedure) || !arrayProcedureHasReturnAllocation(file, procedure) || arrayProcedureHasNonEmptyReturnAllocation(file, procedure) {
				continue
			}
			name, ok := arrayProcedureReturnSource(file, procedure)
			if !ok {
				continue
			}
			lower, ok := arrayProcedureReturnLowerBound(file, procedure, name)
			if !ok {
				continue
			}
			shape := []arrayDimension{{lower: lower}}
			returns[arrayProcedureKey(procedure)] = arrayValue{
				// Keep this as an unknown allocation state. The lower bound is a
				// separate proof used only after a successful UBound query; marking
				// it as an unallocated known array would make the deterministic
				// runtime lane report the same query and suppress the lifecycle
				// diagnostic that the caller still needs to see.
				kind:          arrayUnknown,
				knownArray:    false,
				dimensions:    shape,
				preserveShape: shape,
				origin:        arrayOriginLocal,
			}
		}
	}
	return returns
}

func arrayProcedureReturnSource(file parsedFile, proc sourceProcedure) (string, bool) {
	name := ""
	start := max(0, proc.StartLine-1)
	end := min(proc.EndLine, len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(normalizedCodeLine(rawLine)))
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		source := directArrayArgumentName(rhs)
		if source == "" || name != "" && !strings.EqualFold(name, source) {
			return "", false
		}
		name = source
	}
	return name, name != ""
}

func arrayProcedureReturnLowerBound(file parsedFile, proc sourceProcedure, source string) (arrayBound, bool) {
	var lower arrayBound
	hasLower := false
	start := max(0, proc.StartLine-1)
	end := min(proc.EndLine, len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		text := strings.TrimSpace(normalizedCodeLine(rawLine))
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct || !strings.EqualFold(cleanIdentifier(redim.name), source) {
					continue
				}
				dimensions := parseArrayDimensionsWithConstants(redim.dimensions, arrayOptionBase(file), arrayIntegerConstants(file, proc, nil, nil))
				if len(dimensions) != 1 || !dimensions[0].lower.known {
					return arrayBound{}, false
				}
				if hasLower && lower.value != dimensions[0].lower.value {
					return arrayBound{}, false
				}
				lower = dimensions[0].lower
				hasLower = true
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
			for _, target := range splitArgs(match[1]) {
				if strings.EqualFold(cleanIdentifier(target), source) {
					return arrayBound{}, false
				}
			}
		}
		if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), source) {
			return arrayBound{}, false
		}
	}
	return lower, hasLower
}

func arrayProcedureDocumentsArray(file parsedFile, proc sourceProcedure) bool {
	start := max(0, proc.StartLine-1-5)
	end := min(max(0, proc.StartLine-1), len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		comment := strings.TrimSpace(rawLine)
		if !strings.HasPrefix(comment, "'") {
			continue
		}
		comment = strings.TrimSpace(comment[1:])
		if arrayReturnArrayDocRe.MatchString(comment) {
			return true
		}
	}
	return false
}

func arrayProcedureHasReturnAllocation(file parsedFile, proc sourceProcedure) bool {
	hasAllocation := false
	hasReturn := false
	start := max(0, proc.StartLine-1)
	end := min(proc.EndLine, len(file.Lines))
	for _, rawLine := range file.Lines[start:end] {
		text := strings.TrimSpace(normalizedCodeLine(rawLine))
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "redim ") {
			hasAllocation = true
		}
		lhs, _, indexed, assigned := arrayAssignment(text)
		if assigned && !indexed && strings.EqualFold(lhs, proc.Name) {
			hasReturn = true
		}
	}
	return hasAllocation && hasReturn
}

// arrayDescriptorArrayReturnSummary recognizes a typed array Function that
// returns a view backed by a persistent SAFEARRAY descriptor. The contract is
// intentionally structural: a Static UDT-like accessor is initialized behind
// its readiness flag, the descriptor data/count/lower-bound fields come from
// scalar parameters, and the direct array member is returned on every normal
// path. Unknown call arguments retain a possible-empty allocation; known
// literal arguments can recover the returned one-dimensional shape later.
func arrayDescriptorArrayReturnSummary(file parsedFile, proc sourceProcedure, ctx analysisContext) (arrayValue, bool) {
	if proc.Graph == nil || proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return arrayValue{}, false
	}
	if proc.ReturnValueShape != procedureir.ValueShapeDynamicArray && !strings.Contains(strings.ReplaceAll(proc.ReturnType, " ", ""), "()") {
		return arrayValue{}, false
	}
	root, returnLine, ok := arrayDescriptorReturnSource(file, proc)
	if !ok {
		return arrayValue{}, false
	}
	variables := arrayVariables(file, proc, file.moduleDecls())
	ready, known := variables[strings.ToLower(root)]
	if !known || !ready.static || ready.isArray || ready.isVariant || ready.isObject || ready.knownScalar {
		return arrayValue{}, false
	}
	if !arrayDescriptorReadyInitialized(file, proc, root, ctx) {
		return arrayValue{}, false
	}
	source, start, length, lower, setupLines, ok := arrayDescriptorReturnSetup(file, proc, root)
	if !ok || !arrayDescriptorParameter(proc, source, "String") || !arrayDescriptorParameter(proc, start, "") || !arrayDescriptorParameter(proc, length, "") || !arrayDescriptorParameter(proc, lower, "") {
		return arrayValue{}, false
	}
	for _, line := range setupLines {
		if line >= returnLine || !arrayDescriptorLineDominatesNormalExit(proc, line) {
			return arrayValue{}, false
		}
	}
	if !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true}) {
		return arrayValue{}, false
	}
	return arrayValue{
		kind:                            arrayAllocated,
		knownArray:                      true,
		mayBeEmpty:                      true,
		origin:                          arrayOriginLocal,
		returnDescriptorSourceParameter: strings.ToLower(source),
		returnDescriptorStartParameter:  strings.ToLower(start),
		returnDescriptorLengthParameter: strings.ToLower(length),
		returnDescriptorLowerParameter:  strings.ToLower(lower),
	}, true
}

func arrayDescriptorReturnSource(file parsedFile, proc sourceProcedure) (string, int, bool) {
	root := ""
	returnLine := 0
	for line := max(0, proc.StartLine-1); line < min(proc.EndLine, len(file.Lines)); line++ {
		lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(normalizedCodeLine(file.Lines[line])))
		if !assigned || indexed || !strings.EqualFold(lhs, proc.Name) {
			continue
		}
		if returnLine != 0 {
			return "", 0, false
		}
		receiver, member, ok := arrayQualifiedMemberParts(rhs)
		if !ok || receiver == "" {
			return "", 0, false
		}
		parts := append(strings.Split(receiver, "."), member)
		if len(parts) != 3 || !strings.EqualFold(parts[1], "ac") || !strings.HasPrefix(strings.ToLower(parts[2]), "d") {
			return "", 0, false
		}
		root = parts[0]
		returnLine = line + 1
	}
	return root, returnLine, root != "" && returnLine != 0
}

func arrayDescriptorReadyInitialized(file parsedFile, proc sourceProcedure, root string, ctx analysisContext) bool {
	pattern := regexp.MustCompile(`(?i)(?:^|:)\s*if\s+not\s+` + regexp.QuoteMeta(root) + `\s*\.\s*isset\s+then\s+initmemoryaccessor\s+` + regexp.QuoteMeta(root) + `\s*$`)
	guardLine := -1
	for line := max(0, proc.StartLine-1); line < min(proc.EndLine, len(file.Lines)); line++ {
		if pattern.MatchString(strings.TrimSpace(normalizedCodeLine(file.Lines[line]))) {
			if guardLine >= 0 {
				return false
			}
			guardLine = line
		}
	}
	if guardLine < 0 {
		return false
	}
	initializerFound := false
	for call := range proc.Calls.All() {
		if call.Range.StartLine-1 != guardLine {
			continue
		}
		if initializerFound {
			return false
		}
		helper, parameter, resolved := arrayStaticReadyInitializer(file, proc, call, root, ctx)
		if !resolved || helper.StartByte == proc.StartByte || !arrayStaticHelperSetsReadyFlag(file, helper, parameter) {
			return false
		}
		initializerFound = true
	}
	return initializerFound
}

func arrayDescriptorReturnSetup(file parsedFile, proc sourceProcedure, root string) (string, string, string, string, []int, bool) {
	prefix := regexp.QuoteMeta(root)
	pvRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*pvdata\s*=\s*strptr\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*\(\s*([A-Za-z_]\w*)\s*-\s*1\s*\)\s*\*\s*[A-Za-z_]\w*\s*$`)
	cbRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*cbelements\s*=\s*\S.*$`)
	lowerRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*rgsabound0\s*\.\s*llbound\s*=\s*([A-Za-z_]\w*)\s*$`)
	countRe := regexp.MustCompile(`(?i)^\s*` + prefix + `\s*\.\s*sa\s*\.\s*rgsabound0\s*\.\s*celements\s*=\s*([A-Za-z_]\w*)\s*$`)
	source, start, length, lower := "", "", "", ""
	pvCount, cbCount, lowerCount, countCount := 0, 0, 0, 0
	lines := make([]int, 0, 4)
	for line := max(0, proc.StartLine-1); line < min(proc.EndLine, len(file.Lines)); line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line]))
		if match := pvRe.FindStringSubmatch(text); len(match) == 3 {
			pvCount++
			source, start = match[1], match[2]
			lines = append(lines, line+1)
		}
		if cbRe.MatchString(text) {
			cbCount++
			lines = append(lines, line+1)
		}
		if match := lowerRe.FindStringSubmatch(text); len(match) == 2 {
			lowerCount++
			lower = match[1]
			lines = append(lines, line+1)
		}
		if match := countRe.FindStringSubmatch(text); len(match) == 2 {
			countCount++
			length = match[1]
			lines = append(lines, line+1)
		}
	}
	if pvCount != 1 || cbCount != 1 || lowerCount != 1 || countCount != 1 {
		return "", "", "", "", nil, false
	}
	return source, start, length, lower, lines, true
}

func arrayDescriptorParameter(proc sourceProcedure, name, typeName string) bool {
	for parameter := range proc.Params.All() {
		if !strings.EqualFold(strings.TrimSpace(parameter.Name), strings.TrimSpace(name)) {
			continue
		}
		if parameter.IsArray || parameter.ValueShape == procedureir.ValueShapeFixedArray || parameter.ValueShape == procedureir.ValueShapeDynamicArray {
			return false
		}
		return typeName == "" || strings.EqualFold(strings.TrimSpace(parameter.Type), typeName)
	}
	return false
}

func arrayDescriptorLineDominatesNormalExit(proc sourceProcedure, line int) bool {
	dominators := arrayProcedureNormalExitDominators(proc)
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine == line && arrayProcedureBlockDominatesNormalExit(proc, statement.ID, dominators) {
			return true
		}
	}
	return false
}

// arrayConditionalReturnSummary recognizes a small, path-sensitive family of
// array factories. The returned array is intentionally kept conditional: the
// caller must prove either that the scalar length is positive or that the
// input Byte array is non-empty before this fact becomes an allocation proof.
// Requiring one direct return source, one guarded ReDim, and a definitely
// assigned function result keeps this summary fail-open for general helper
// functions and error paths.
func arrayConditionalReturnSummary(file parsedFile, proc sourceProcedure) (arrayValue, bool) {
	if proc.Graph == nil || proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return arrayValue{}, false
	}
	source, ok := arrayProcedureReturnSource(file, proc)
	if !ok {
		return arrayValue{}, false
	}
	source = strings.ToLower(cleanIdentifier(source))
	variables := arrayVariables(file, proc, file.moduleDecls())
	returnedVariable, ok := variables[source]
	if !ok || !returnedVariable.isArray || returnedVariable.fixed || returnedVariable.parameter {
		return arrayValue{}, false
	}
	if !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true}) {
		return arrayValue{}, false
	}

	returnLines := make([]int, 0, 1)
	foundConditionalOutput := false
	redimLine := 0
	returnNonEmptyArrayParameter := ""
	returnPositiveScalarParameter := ""
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), proc.Name) {
			returnLines = append(returnLines, statement.Range.StartLine)
			continue
		}
		if statement.Kind != procedureir.StatementReDim {
			if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), source) {
				condition, branch, guarded := arrayConditionalReturnGuard(proc, statement)
				input, proven := arrayConditionalReturnStrPtrInput(condition, branch, variables)
				if !guarded || !proven || returnPositiveScalarParameter != "" || returnNonEmptyArrayParameter != "" && !strings.EqualFold(returnNonEmptyArrayParameter, input) {
					return arrayValue{}, false
				}
				returnNonEmptyArrayParameter = input
				foundConditionalOutput = true
			}
			if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
				for _, target := range splitArgs(match[1]) {
					if strings.EqualFold(cleanIdentifier(target), source) {
						return arrayValue{}, false
					}
				}
			}
			continue
		}
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 {
			continue
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct || !strings.EqualFold(cleanIdentifier(redim.name), source) {
				continue
			}
			if strings.TrimSpace(match[1]) != "" {
				return arrayValue{}, false
			}
			condition, branch, guarded := arrayConditionalReturnGuard(proc, statement)
			if !guarded {
				return arrayValue{}, false
			}
			if input, proven := arrayConditionalReturnStrPtrInput(condition, branch, variables); proven {
				if returnPositiveScalarParameter != "" || returnNonEmptyArrayParameter != "" && !strings.EqualFold(returnNonEmptyArrayParameter, input) {
					return arrayValue{}, false
				}
				returnNonEmptyArrayParameter = input
			} else if scalar, proven := arrayConditionalReturnPositiveScalar(condition, branch, redim.dimensions, variables); proven {
				if returnNonEmptyArrayParameter != "" || returnPositiveScalarParameter != "" && !strings.EqualFold(returnPositiveScalarParameter, scalar) {
					return arrayValue{}, false
				}
				returnPositiveScalarParameter = scalar
			} else {
				return arrayValue{}, false
			}
			foundConditionalOutput = true
			if redimLine == 0 || statement.Range.StartLine < redimLine {
				redimLine = statement.Range.StartLine
			}
		}
	}
	if len(returnLines) != 1 || !foundConditionalOutput || redimLine > returnLines[0] {
		return arrayValue{}, false
	}
	return arrayValue{
		kind:                          arrayUnknown,
		knownArray:                    true,
		mayBeEmpty:                    true,
		origin:                        arrayOriginLocal,
		returnNonEmptyArrayParameter:  returnNonEmptyArrayParameter,
		returnPositiveScalarParameter: returnPositiveScalarParameter,
	}, true
}

func arrayConditionalReturnGuard(proc sourceProcedure, statement procedureir.Statement) (string, vbacfg.EdgeKind, bool) {
	current := statement
	visited := map[int]bool{}
	for current.ParentID != 0 && !visited[current.ParentID] {
		visited[current.ParentID] = true
		parent := procedureStatementByID(proc, current.ParentID)
		if parent.ID == 0 {
			return "", "", false
		}
		switch parent.Kind {
		case procedureir.StatementIf, procedureir.StatementElseIf:
			return arrayConditionalReturnCondition(parent), vbacfg.EdgeBranchTrue, true
		case procedureir.StatementElse:
			branch := procedureStatementByID(proc, parent.ParentID)
			if branch.Kind != procedureir.StatementIf && branch.Kind != procedureir.StatementElseIf {
				return "", "", false
			}
			return arrayConditionalReturnCondition(branch), vbacfg.EdgeBranchFalse, true
		}
		current = parent
	}
	return "", "", false
}

func arrayConditionalReturnCondition(statement procedureir.Statement) string {
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	condition = strings.TrimSpace(condition)
	lower := strings.ToLower(condition)
	if strings.HasPrefix(lower, "if ") {
		condition = strings.TrimSpace(condition[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		condition = strings.TrimSpace(condition[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(condition), " then"); then >= 0 && strings.TrimSpace(condition[then+5:]) == "" {
		condition = strings.TrimSpace(condition[:then])
	}
	for len(condition) >= 2 && condition[0] == '(' && condition[len(condition)-1] == ')' {
		condition = strings.TrimSpace(condition[1 : len(condition)-1])
	}
	return condition
}

func arrayConditionalReturnStrPtrInput(condition string, branch vbacfg.EdgeKind, variables map[string]arrayVariable) (string, bool) {
	match := arrayStrPtrGuardRe.FindStringSubmatch(strings.TrimSpace(condition))
	if len(match) != 3 {
		return "", false
	}
	name := strings.ToLower(cleanIdentifier(match[1]))
	variable, known := variables[name]
	if !known || !variable.parameter || !parameterIsByRefArrayForVariable(variable) || !isByteArrayVariable(variable) {
		return "", false
	}
	required := vbacfg.EdgeBranchFalse
	if match[2] == "<>" {
		required = vbacfg.EdgeBranchTrue
	}
	return name, branch == required
}

func parameterIsByRefArrayForVariable(variable arrayVariable) bool {
	return variable.parameter && variable.isArray && !variable.paramArray
}

func arrayConditionalReturnPositiveScalar(condition string, branch vbacfg.EdgeKind, dimensions string, variables map[string]arrayVariable) (string, bool) {
	lhs, operator, literal, ok := arrayCountComparison(condition)
	if !ok {
		return "", false
	}
	name := strings.ToLower(cleanIdentifier(lhs))
	variable, known := variables[name]
	if !known || !variable.parameter || variable.isArray || variable.isVariant || variable.isObject {
		return "", false
	}
	positiveBranch, ok := positiveArrayCountBranch(operator, literal)
	if !ok || branch != positiveBranch || !arrayRedimUsesPositiveScalar(dimensions, name) {
		return "", false
	}
	return name, true
}

func arrayRedimUsesPositiveScalar(dimensions, parameter string) bool {
	wanted := strings.ToLower(cleanIdentifier(parameter)) + "-1"
	for _, dimension := range splitArgs(dimensions) {
		compact := strings.ToLower(canonicalArrayBoundExpression(dimension))
		if strings.HasPrefix(compact, "0to") && strings.TrimPrefix(compact, "0to") == wanted {
			return true
		}
	}
	return false
}

// arrayProcedureHasNonEmptyReturnAllocation recognizes the stronger CFG-backed
// contract needed when a documented array return is used by a bare caller.
// A plain ReDim with fully known, non-empty bounds is different from a
// ReDim Preserve inside a loop: the latter can leave the returned array
// unallocated when the loop has no iterations. Every normal return path must
// therefore assign the documented return value from a known non-empty array.
func arrayProcedureHasNonEmptyReturnAllocation(file parsedFile, proc sourceProcedure) bool {
	if proc.Graph == nil {
		return false
	}
	variables := arrayVariables(file, proc, file.moduleDecls())
	constants := arrayIntegerConstants(file, proc, nil, nil)
	ctx := analysisContext{arrayAllowVariantRedim: true}
	type returnCandidate struct {
		value arrayValue
		ok    bool
	}
	returnCandidates := map[int]returnCandidate{}
	base := arrayOptionBase(file)
	procedureHasErrorHandling := arrayProcedureHasErrorHandling(proc)
	// An Err.Raise branch is not a normal return path. Keep exceptional edges
	// available for procedures with active error handling, but remove the
	// synthetic normal continuation so an error-only branch cannot make the
	// return assignment look non-definite.
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	walkArrayCFGWithStopStats(&graph, file.Lines, arrayInitialState(variables), func(text string, line int, in arrayFlowState) arrayFlowState {
		if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(lhs, proc.Name) {
			value, known := arrayExpressionState(rhs, in, ctx)
			returnCandidates[line] = returnCandidate{
				value: value,
				ok:    known && value.kind == arrayAllocated && value.knownArray && value.origin != arrayOriginRangeValue,
			}
		}
		out, _ := (Analyzer{}).arrayTransfer(file, proc, ctx, variables, in, text, line, constants, nil)
		return out
	}, nil, func(text string, _ int) bool {
		return !procedureHasErrorHandling && arraySummaryStatementAlwaysFails(text, base, constants)
	}, nil)
	if len(returnCandidates) == 0 || !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true, WithoutNormalErrRaiseContinuation: true}) {
		return false
	}
	for _, candidate := range returnCandidates {
		if !candidate.ok || candidate.value.mayBeEmpty {
			return false
		}
	}
	return true
}

type arrayReturnSummarySet struct {
	bare      map[string]arrayValue
	qualified map[string]arrayValue
}

func inferArrayReturnSummaries(files []parsedFile, arrayAllocationGuards map[string]bool, participantCtx analysisContext) map[string]arrayValue {
	return inferArrayReturnSummarySet(files, arrayAllocationGuards, participantCtx).bare
}

func inferArrayReturnSummarySet(files []parsedFile, arrayAllocationGuards map[string]bool, participantCtx analysisContext) arrayReturnSummarySet {
	type candidate struct {
		value arrayValue
		ok    bool
	}
	type returnProcedure struct {
		file      parsedFile
		proc      sourceProcedure
		variables map[string]arrayVariable
		constants map[string]int
	}
	procedures := make([]returnProcedure, 0)
	allReturnProcedures := make([]sourceProcedure, 0)
	for _, file := range files {
		fileProcedures := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < fileProcedures.Len(); procedureIndex++ {
			proc := fileProcedures.valueAt(procedureIndex)
			if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
				continue
			}
			if proc.Name == "" {
				continue
			}
			// Duplicate bare-name summaries are ambiguous even when one of the
			// same-named procedures is outside the array participant closure.
			// Collect names before filtering so narrowing the fixed-point scope
			// cannot turn an otherwise ambiguous call into a definite summary.
			allReturnProcedures = append(allReturnProcedures, proc)
			if !arrayProcedureIsParticipant(participantCtx, proc) {
				continue
			}
			procedures = append(procedures, returnProcedure{
				file:      file,
				proc:      proc,
				variables: arrayVariables(file, proc, moduleDecls),
				constants: arrayIntegerConstants(file, proc, nil, nil),
			})
		}
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	ambiguousReturnNames := arrayReturnSummaryDuplicateNames(allReturnProcedures)

	evaluate := func(procedure returnProcedure, summaries map[string]arrayValue) (candidate, bool) {
		proc := procedure.proc
		if proc.Graph == nil {
			// Without a CFG the scan cannot distinguish conditional from
			// unconditional return assignments, so leave the summary unknown.
			return candidate{}, false
		}
		arrayReturns := summaries
		// Never use a previous summary for the procedure currently being
		// inspected. This keeps direct and mutual recursion fail-open while
		// still allowing independent helpers to converge over later rounds.
		if _, self := summaries[strings.ToLower(proc.Name)]; self {
			arrayReturns = make(map[string]arrayValue, len(summaries)-1)
			for name, value := range summaries {
				if !strings.EqualFold(name, proc.Name) {
					arrayReturns[name] = value
				}
			}
		}
		ctx := analysisContext{
			arrayReturns:           arrayReturns,
			arrayAllowVariantRedim: arrayProcedureDocumentsArray(procedure.file, proc) && arrayProcedureHasReturnAllocation(procedure.file, proc),
		}
		returnCandidates := map[int]candidate{}
		base := arrayOptionBase(procedure.file)
		procedureHasErrorHandling := arrayProcedureHasErrorHandling(proc)
		if participantCtx.arrayStats != nil {
			participantCtx.arrayStats.addCFGWalk()
		}
		baseView := proc.Graph.View(vbacfg.EdgeFilter{})
		walkArrayCFGWithStopStats(&baseView, procedure.file.Lines, arrayInitialState(procedure.variables), func(text string, line int, in arrayFlowState) arrayFlowState {
			if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed && strings.EqualFold(lhs, proc.Name) {
				value, known := arrayExpressionState(rhs, in, ctx)
				returnCandidates[line] = candidate{value: value, ok: known && value.kind == arrayAllocated && value.knownArray && value.origin != arrayOriginRangeValue}
			}
			out, _ := (Analyzer{}).arrayTransfer(procedure.file, proc, ctx, procedure.variables, in, text, line, procedure.constants, nil)
			return out
		}, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			return applyArrayAllocationGuard(out, block.Statement, edge, arrayAllocationGuards, procedure.variables)
		}, func(text string, _ int) bool {
			return !procedureHasErrorHandling && arraySummaryStatementAlwaysFails(text, base, procedure.constants)
		}, participantCtx.arrayStats)
		conditionalValue, hasConditional := arrayConditionalReturnSummary(procedure.file, proc)
		descriptorValue, hasDescriptor := arrayDescriptorArrayReturnSummary(procedure.file, proc, participantCtx)
		returnLines := make([]int, 0, len(returnCandidates))
		for line := range returnCandidates {
			returnLines = append(returnLines, line)
		}
		sort.Ints(returnLines)
		returns := make([]candidate, 0, len(returnLines))
		for _, line := range returnLines {
			returns = append(returns, returnCandidates[line])
		}
		if len(returns) == 0 || !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true}) {
			if hasConditional {
				return candidate{value: conditionalValue, ok: true}, true
			}
			if hasDescriptor {
				return candidate{value: descriptorValue, ok: true}, true
			}
			return candidate{}, false
		}
		valid := returns[0].ok
		value := returns[0].value
		for _, returned := range returns[1:] {
			if !returned.ok || !arrayReturnValueCompatible(proc, value, returned.value) {
				valid = false
				break
			}
			value = meetArrayValue(value, returned.value)
		}
		if !valid && hasConditional {
			return candidate{value: conditionalValue, ok: true}, true
		}
		if !valid && hasDescriptor {
			return candidate{value: descriptorValue, ok: true}, true
		}
		return candidate{value: value, ok: valid}, true
	}

	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		for call := range procedure.proc.Calls.All() {
			resolution := call.Resolution
			if participantCtx.procedureResolver != nil {
				resolution = participantCtx.procedureResolver.ResolveCall(call)
			}
			for _, candidate := range resolution.Candidates {
				name := strings.ToLower(strings.TrimSpace(candidate.QualifiedName))
				if dot := strings.LastIndexByte(name, '.'); dot >= 0 {
					name = name[dot+1:]
				}
				if name != "" {
					dependents[name] = append(dependents[name], index)
				}
			}
		}
	}
	for name := range dependents {
		sort.Ints(dependents[name])
	}
	contributions := make(map[string]candidate, len(procedures))
	present := make(map[string]bool, len(procedures))
	groups := make(map[string]map[string]candidate)
	summaries := map[string]arrayValue{}
	qualifiedSummaries := map[string]arrayValue{}
	documentedSummaries := inferDocumentedArrayReturnSummaries(files)
	documentedBareSummaries := inferDocumentedNonEmptyArrayReturnSummaries(files)
	documentedLowerBoundSummaries := inferDocumentedArrayReturnLowerBoundSummaries(files)
	for key, value := range documentedSummaries {
		qualifiedSummaries[key] = value
	}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		procedure := procedures[index]
		name := strings.ToLower(strings.TrimSpace(procedure.proc.Name))
		key := arrayProcedureKey(procedure.proc)
		value, hasContribution := evaluate(procedure, summaries)
		if hasContribution && value.ok {
			qualifiedSummaries[key] = value.value
		} else {
			delete(qualifiedSummaries, key)
		}
		if ambiguousReturnNames[name] {
			// Summary lookups use bare names for compatibility with the existing
			// expression resolver. A duplicate bare name is therefore permanently
			// ambiguous for this revision. Keep it at the unknown bottom of the
			// lattice instead of allowing iteration order to delete and recreate a
			// summary while duplicate candidates are evaluated.
			continue
		}
		if head >= len(procedures) && participantCtx.arrayStats != nil {
			participantCtx.arrayStats.addRevisit()
		}
		if present[key] == hasContribution && (!hasContribution || arrayValueEqual(contributions[key].value, value.value) && contributions[key].ok == value.ok) {
			continue
		}
		if hasContribution {
			contributions[key] = value
			present[key] = true
			if groups[name] == nil {
				groups[name] = map[string]candidate{}
			}
			groups[name][key] = value
		} else {
			delete(contributions, key)
			delete(present, key)
			if group := groups[name]; group != nil {
				delete(group, key)
				if len(group) == 0 {
					delete(groups, name)
				}
			}
		}
		old, hadOld := summaries[name]
		group := groups[name]
		if len(group) == 1 {
			for _, value := range group {
				if value.ok {
					summaries[name] = value.value
				} else {
					delete(summaries, name)
				}
			}
		} else {
			delete(summaries, name)
		}
		fresh, hasFresh := summaries[name]
		changed := hadOld != hasFresh || (hasFresh && !arrayValueEqual(old, fresh))
		if !changed {
			continue
		}
		for _, dependent := range dependents[name] {
			if !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
		}
	}
	for key, value := range documentedSummaries {
		if _, exists := qualifiedSummaries[key]; !exists {
			qualifiedSummaries[key] = value
		}
	}
	for key, value := range documentedBareSummaries {
		name := key
		if dot := strings.LastIndexByte(name, '.'); dot >= 0 {
			name = name[dot+1:]
		}
		if !ambiguousReturnNames[name] {
			if _, exists := summaries[name]; !exists {
				summaries[name] = value
			}
		}
	}
	for key, value := range documentedLowerBoundSummaries {
		name := key
		if dot := strings.LastIndexByte(name, '.'); dot >= 0 {
			name = name[dot+1:]
		}
		if !ambiguousReturnNames[name] {
			// The lower-bound-only contract is deliberately more precise than
			// an inferred allocated summary: it preserves the possible empty
			// input path while still proving the loop's lower-side index bound.
			// Prefer it for the bare lookup so a generic fixed-point result
			// cannot erase the retained UBound diagnostic.
			summaries[name] = value
		}
	}
	return arrayReturnSummarySet{bare: summaries, qualified: qualifiedSummaries}
}

func arrayProcedureHasErrorHandling(proc sourceProcedure) bool {
	for statement := range proc.Statements.All() {
		if strings.HasPrefix(strings.ToLower(strings.TrimSpace(statement.Text)), "on error ") {
			return true
		}
	}
	return false
}

func arraySummaryStatementAlwaysFails(text string, base int, constants map[string]int) bool {
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct {
			continue
		}
		if impossibleArrayBounds(parseArrayDimensionsWithConstants(redim.dimensions, base, constants)) {
			return true
		}
	}
	return false
}

func arrayValueEqual(left, right arrayValue) bool {
	return arrayValueCompatible(left, right) && left.mayBeEmpty == right.mayBeEmpty
}

// arrayReturnValueCompatible keeps a return summary when every normal return
// path produces an array with the same allocation provenance. Return branches
// may legitimately produce different shapes (for example Array() for an
// empty result and ReDim 1 To length for a non-empty result); that shape
// mismatch must not erase the stronger fact that the returned value is an
// array. Callers still receive no shape bounds from such a summary.
func arrayReturnValueCompatible(proc sourceProcedure, left, right arrayValue) bool {
	if arrayValueCompatible(left, right) {
		return true
	}
	if proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return false
	}
	return left.kind == arrayAllocated && right.kind == arrayAllocated && left.knownArray && right.knownArray && left.origin == right.origin && left.allocationProbe == right.allocationProbe && left.safeBoundProbe == right.safeBoundProbe && left.allocationCountSource == right.allocationCountSource && left.returnNonEmptyArrayParameter == right.returnNonEmptyArrayParameter && left.returnPositiveScalarParameter == right.returnPositiveScalarParameter && left.nonEmptySource == right.nonEmptySource && left.returnDescriptorSourceParameter == right.returnDescriptorSourceParameter && left.returnDescriptorStartParameter == right.returnDescriptorStartParameter && left.returnDescriptorLengthParameter == right.returnDescriptorLengthParameter && left.returnDescriptorLowerParameter == right.returnDescriptorLowerParameter
}

func arrayValueCompatible(left, right arrayValue) bool {
	return left.kind == right.kind && left.knownArray == right.knownArray && left.origin == right.origin && left.allocationProbe == right.allocationProbe && left.allocationCountSource == right.allocationCountSource && left.returnNonEmptyArrayParameter == right.returnNonEmptyArrayParameter && left.returnPositiveScalarParameter == right.returnPositiveScalarParameter && left.nonEmptySource == right.nonEmptySource && left.returnDescriptorSourceParameter == right.returnDescriptorSourceParameter && left.returnDescriptorStartParameter == right.returnDescriptorStartParameter && left.returnDescriptorLengthParameter == right.returnDescriptorLengthParameter && left.returnDescriptorLowerParameter == right.returnDescriptorLowerParameter && arrayDimensionsEqual(left.dimensions, right.dimensions) && arrayDimensionsEqual(left.preserveShape, right.preserveShape)
}
