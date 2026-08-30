package analyze

import (
	"context"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
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
	// allocationCountSource records a narrow conditional allocation contract:
	// the array is allocated when the named scalar is positive, or when the
	// named collection's Count is positive. The fact is refined only on a
	// matching control-flow branch; it is never treated as unconditional.
	allocationCountSource string
}

type arrayFlowState map[string]arrayValue

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
	arrayRedimRe              = regexp.MustCompile(`(?i)^\s*redim\s+(preserve\s+)?(.+)$`)
	arrayRedimClauseRe        = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?)?\s*$`)
	arrayRedimTypeSuffixRe    = regexp.MustCompile(`(?i)^as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?$`)
	arrayEraseRe              = regexp.MustCompile(`(?i)^\s*erase\s+(.+)$`)
	arrayEraseNameRe          = regexp.MustCompile(`(?i)^[A-Za-z_]\w*$`)
	arrayBoundCallRe          = regexp.MustCompile(`(?i)\b(lbound|ubound)\s*\(\s*([^,)]*)\s*(?:,\s*([^)]*))?\)`)
	arrayForBoundRe           = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*([-+]?\d+)\s+to\s+(?:lbound|ubound)\s*\(\s*([A-Za-z_]\w*)`)
	arrayForEachRe            = regexp.MustCompile(`(?i)^\s*for\s+each\s+[A-Za-z_]\w*\s+in\s+([^\r\n]+)`)
	arrayIndexedSourceRe      = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\(`)
	arrayGuardCallRe          = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*(?:(=|<>|>=|<=|>|<)\s*(-?\d+))?\s*$`)
	arrayGuardReversedRe      = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>|>=|<=|>|<)\s*([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*$`)
	arrayGuardValueRe         = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*(=|<>|>=|<=|>|<)\s*(-?\d+)\s*$`)
	arrayGuardValueReversedRe = regexp.MustCompile(`(?i)^\s*(-?\d+)\s*(=|<>|>=|<=|>|<)\s*([A-Za-z_]\w*)\s*$`)
	arrayIsArrayGuardRe       = regexp.MustCompile(`(?i)^\s*isarray\s*\(\s*(.+)\s*\)\s*$`)
	arrayByteArrayGuardRe     = regexp.MustCompile(`(?i)^\s*(?:vartypeof|vartype)\s*\(\s*([A-Za-z_]\w*)\s*\)\s*=\s*\(?\s*vbarray\s+or\s+vbbyte\s*\)?\s*$`)
	arrayByteArrayReadRe      = regexp.MustCompile(`(?i)^\s*(?:[A-Za-z_]\w*\.)*read\s*\(\s*-1\s*\)\s*$`)
	arraySetupGuardRe         = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s+then\s+exit\s+sub\s*$`)
	arrayOnErrorGotoRe        = regexp.MustCompile(`(?i)^\s*on\s+error\s+goto\s+([A-Za-z_]\w*)\s*$`)
	arrayOnErrorResumeNextRe  = regexp.MustCompile(`(?i)^\s*on\s+error\s+resume\s+next\s*$`)
	arrayOnErrorGotoZeroRe    = regexp.MustCompile(`(?i)^\s*on\s+error\s+goto\s+0\s*$`)
	arrayErrNumberFailureRe   = regexp.MustCompile(`(?i)^\s*if\s+err\.number\s*<>\s*0\s+then\s*$`)
	arrayCapacityProbeRe      = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayBoundsProbeRe        = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*-\s*lbound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayCheckedProbeExitRe   = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*(?:<=|=)\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayCapacityIfRe         = regexp.MustCompile(`(?i)^\s*if\s+.+\s*>\s*([A-Za-z_]\w*)\s+then\s*$`)
	arrayForZeroToCountRe     = regexp.MustCompile(`(?i)^\s*for\s+[A-Za-z_]\w*\s*=\s*0\s+to\s+[A-Za-z_]\w*\s*-\s*1\s*$`)
	arrayLabelRe              = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*:\s*$`)
	arrayCountComparisonRe    = regexp.MustCompile(`(?i)^\s*(.*?)\s*(=|<>|>=|<=|>|<)\s*(-?\d+)\s*$`)
	arrayConditionAndRe       = regexp.MustCompile(`(?i)\s+and\s+`)
	arraySelectCaseRe         = regexp.MustCompile(`(?i)^select\s+case\s+(.+)$`)
	arrayPositiveCaseRe       = regexp.MustCompile(`(?i)^case\s+(-?\d+)\s*$`)
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
	if a.Config.Analyze.DetectArrayLifecycleSafety {
		vba227Variables = arrayVBA227Variables(variables, file, proc)
	}
	comparisonFindings := a.arrayComparisonFindings(file, proc, variables)
	comparisonFindings = append(comparisonFindings, a.arrayForEachFindings(file, proc, variables, ctx)...)
	baseLaneRequested := a.Config.Analyze.DetectArrayLifecycleSafety || a.Config.Analyze.DetectRedimPreserveDimension || objectArrayDiagnosticsApplicable
	initial := preparedEntry
	if initial == nil {
		initial = arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, variables)
	}
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
					vba227State, lifecycleIssues = a.arrayVBA227Transfer(file, proc, ctx, vba227Variables, vba227State, text, line, constants, capacityGuards)
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
				for _, issue := range deterministicArrayRuntimeIssues(text, line, runtimeState, variables, constants) {
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
				out, issues := a.arrayVBA227Transfer(file, proc, ctx, vba227Variables, in, text, line, constants, capacityGuards)
				for _, call := range arrayCallsAtLine(proc.Calls, line) {
					out = applyArrayModuleCallEffects(out, file, proc, call, ctx, vba227Variables, moduleDecls)
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
				out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, vba227Variables)
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
				for _, issue := range deterministicArrayRuntimeIssues(text, line, in, variables, constants) {
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

func (a Analyzer) arrayVBA227Transfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) (arrayFlowState, []Finding) {
	if line >= 1 && line <= len(file.Lines) && vbaLineContinues(file.Lines[line-1]) && arrayVBA227HasArrayFactoryAssignment(text) {
		text = arrayLogicalCodeLine(file.Lines, line)
	}
	if redim, ok := inlineArrayRedimText(text); ok {
		text = redim
	}
	if condition, body, ok := arrayIfThenParts(text); ok {
		if argument, _, guard := arrayIsArrayGuardCondition(condition); guard && arrayElementBaseName(argument) != "" {
			state, findings := a.arrayTransfer(file, proc, ctx, variables, state, condition, line, constants, capacityGuards)
			guardedState := arrayElementGuardState(state, argument, variables)
			thenState := cloneArrayState(guardedState)
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			if thenBody != "" {
				var bodyFindings []Finding
				thenState, bodyFindings = a.arrayTransfer(file, proc, ctx, variables, thenState, thenBody, line, constants, capacityGuards)
				findings = append(findings, bodyFindings...)
			}
			if hasElse {
				elseState := cloneArrayState(guardedState)
				var elseFindings []Finding
				if elseBody != "" {
					elseState, elseFindings = a.arrayTransfer(file, proc, ctx, variables, elseState, elseBody, line, constants, capacityGuards)
				}
				findings = append(findings, elseFindings...)
				state = meetArrayState(thenState, elseState)
			} else {
				state = thenState
			}
			return state, findings
		}
	}
	state, findings := a.arrayTransfer(file, proc, ctx, variables, state, text, line, constants, capacityGuards)
	// Source-line CFG blocks can contain an If condition and its body. Apply
	// the normal-path fact after the condition while the block is still being
	// processed so a nested element access in the body does not repeat the
	// condition's possible outer-array failure.
	if argument, _, ok := arrayIsArrayGuardCondition(text); ok {
		state = arrayElementGuardState(state, argument, variables)
	}
	return state, findings
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

func walkArrayCFGWorklist(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
	walkArrayCFGWorklistStats(graph, lines, initial, visit, edgeState, stop, sourceLines, nil)
}

func walkArrayCFGWorklistStats(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats) {
	walkArrayCFGWorklistStatsWithReliable(graph, lines, initial, visit, edgeState, stop, sourceLines, stats, nil)
}

func walkArrayCFGWorklistStatsWithReliable(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool, stats *arrayInterproceduralStats, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool) {
	strategy := arrayCFGStrategyAuto
	if stats != nil {
		strategy = stats.strategy
	}
	if strategy == arrayCFGStrategyLegacy {
		if stats != nil {
			stats.addLegacyWalk()
		}
		walkArrayCFGWorklistLegacy(graph, lines, initial, visit, edgeState, stop, sourceLines)
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
				_ = walkArrayCFGCompactAdvanced(context.Background(), graph, lines, initial, visit, edgeState, stop, reliableExceptional, sourceLines)
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
	walkArrayCFGWorklistLegacy(graph, lines, initial, visit, edgeState, stop, sourceLines)
}

func walkArrayCFGWorklistLegacy(graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) {
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
				out = visit(text, line, out)
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
				if block.Statement.Kind == procedureir.StatementSelect && start >= 1 && start <= len(lines) {
					// Select Case owns separate CFG blocks for each Case clause.
					// Visiting the whole source range here would scan every clause
					// once before its case-specific edge state is applied, making a
					// branch-local allocation fact appear to be absent. The clause
					// blocks below own the remaining physical lines.
					text := normalizedCodeLine(lines[start-1])
					out = visit(text, start, out)
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
					out = visit(text, start, out)
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
						out = visit(text, line, out)
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
					out = visit(text, start, out)
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
					if block.Statement.Kind == procedureir.StatementSelect && start >= 1 && start <= len(lines) {
						// Select Case owns separate CFG blocks for each Case clause;
						// do not scan all clause lines before applying the edge fact.
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
	source = canonicalArrayBoundExpression(source)
	return expression == source || expression == source+".count"
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
	updated[name] = value
	return updated
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
		kind:                  left.kind,
		knownArray:            left.knownArray,
		mayBeEmpty:            left.mayBeEmpty,
		origin:                left.origin,
		dimensions:            append([]arrayDimension(nil), left.dimensions...),
		preserveShape:         append([]arrayDimension(nil), left.preserveShape...),
		allocationCountSource: left.allocationCountSource,
	}
	if left.kind != right.kind {
		out.kind = arrayUnknown
	}
	if left.knownArray != right.knownArray {
		out.knownArray = false
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
	if left.allocationCountSource != right.allocationCountSource {
		out.allocationCountSource = ""
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
		if !ok || l.kind != r.kind || l.knownArray != r.knownArray || l.mayBeEmpty != r.mayBeEmpty || l.origin != r.origin || l.allocationProbe != r.allocationProbe || l.allocationCountSource != r.allocationCountSource || !arrayDimensionsEqual(l.dimensions, r.dimensions) || !arrayDimensionsEqual(l.preserveShape, r.preserveShape) {
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
			resolution := call.Resolution
			if ctx.procedureResolver != nil {
				resolution = ctx.procedureResolver.ResolveCall(call)
			}
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
		value := arrayByRefAllocationSummaryForProcedure(procedure.file, procedure.proc, targets, summaries, ctx, dominators[key])
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

func arrayByRefAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, targets map[string]sourceProcedure, summaries arrayByRefAllocationSummaries, ctx analysisContext, dominators map[vbacfg.BlockID]bool) map[int]bool {
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
	allocated := map[int]bool{}
	addAllocation := func(statementID int, name string) {
		index, ok := parameters[strings.ToLower(cleanIdentifier(name))]
		if !ok || arrayProcedureLineHasInlineConditional(file, statementLine(proc, statementID)) || !arrayProcedureBlockDominatesNormalExit(proc, statementID, dominators) {
			return
		}
		allocated[index] = true
	}
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct {
					addAllocation(statement.ID, redim.name)
				}
			}
		}
		if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed {
			if value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx); known && value.kind == arrayAllocated && value.knownArray {
				addAllocation(statement.ID, lhs)
			}
		}
	}
	for call := range proc.Calls.All() {
		if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) || !arrayProcedureBlockDominatesNormalExit(proc, call.StatementID, dominators) {
			continue
		}
		key, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
		if !ok {
			continue
		}
		calleeParameters := summaries[key]
		if len(calleeParameters) == 0 {
			continue
		}
		arguments := arrayCallArgumentTexts(proc, call)
		if len(call.Arguments.Named) > 0 || len(arguments) != call.Arguments.Count {
			continue
		}
		for index := range calleeParameters {
			if index >= len(arguments) || index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
				continue
			}
			if parameterIndex, ok := parameters[directArrayArgumentName(arguments[index])]; ok {
				allocated[parameterIndex] = true
			}
		}
	}
	flowCtx := ctx
	flowCtx.arrayByRefAllocations = summaries
	moduleDecls := file.moduleDecls()
	for index := range arrayByRefFlowAllocations(file, proc, flowCtx, moduleDecls) {
		allocated[index] = true
	}
	return allocated
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
		for _, call := range arrayCallsAtLine(proc.Calls, line) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
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

func statementLine(proc sourceProcedure, statementID int) int {
	for statement := range proc.Statements.All() {
		if statement.ID == statementID {
			return statement.Range.StartLine
		}
	}
	return 0
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
	if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
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
	for name := range ctx.arrayModuleAllocations[key] {
		markModule(name)
	}
	arguments := arrayCallArgumentTexts(proc, call)
	if len(call.Arguments.Named) == 0 && len(arguments) == call.Arguments.Count {
		for index := range ctx.arrayByRefAllocations[key] {
			if index >= len(arguments) || index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
				continue
			}
			markArgument(arguments[index])
		}
		for outputIndex, countIndex := range ctx.arrayByRefConditionalAllocations[key] {
			if outputIndex >= len(arguments) || countIndex < 0 || countIndex >= len(arguments) {
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
	}
	for name := range arrayConfigurationArraysForGuard(file, target, arguments, ctx.arrayModuleConfigurations[file.Path]) {
		markModule(name)
	}
	return updated
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
	idempotentSetupArrays := arrayModuleIdempotentSetupArrays(file, proc, moduleDecls)
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
			arguments := arrayCallArgumentTexts(proc, call)
			if len(call.Arguments.Named) == 0 && len(arguments) == call.Arguments.Count {
				for index := range calleeByRefArrays {
					if index >= len(arguments) || index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
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
func arrayModuleIdempotentSetupArrays(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]bool {
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
				result[name] = true
			})
		}
	}
	if len(result) == 0 {
		return nil
	}
	return result
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
		initial = applyArrayByRefEntryStates(initial, proc, variables, entries, conditions)
		initial = applyArrayModuleEntryState(initial, file, proc, variables, moduleDecls, ctx.arrayModuleEntryStates, ctx.arrayParticipantKeys)
		initial = applyArrayInternalStorageConfiguration(initial, file, proc, variables, moduleDecls, ctx.arrayModuleConfigurations[file.Path])
		visit := func(text string, line int, in arrayFlowState) arrayFlowState {
			var eligible []arrayByRefCallCandidate
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				key, target, ok := arrayPrivateTargetForCall(localCtx, targets, call)
				if !ok || !procedureHasByRefArrayParameter(target) || !arrayProcedureIsParticipant(localCtx, target) {
					continue
				}
				eligible = append(eligible, arrayByRefCallCandidate{key: key, target: target, call: call})
			}
			if len(eligible) > 0 {
				targetKey := eligible[0].key
				sameTarget := true
				for _, entry := range eligible[1:] {
					if entry.key != targetKey {
						sameTarget = false
						break
					}
				}
				if sameTarget {
					for _, entry := range eligible {
						arrayRecordByRefCall(evidence, entry.key, entry.target, proc, entry.call, in, localCtx)
					}
				} else {
					// Nested calls on one source line are normally kept
					// conservative because the pre-line state cannot describe
					// mutations from an earlier, different helper. An outer
					// ByRef call whose array argument is a proven allocated
					// expression is independent of that ordering, however.
					for _, entry := range eligible {
						allProven, hasExpression := arrayByRefCallHasProvenArrayArguments(entry.target, proc, entry.call, in, localCtx)
						if allProven && (hasExpression || arrayByRefCallIsInnermostNested(entry.call, eligible)) {
							arrayRecordByRefCall(evidence, entry.key, entry.target, proc, entry.call, in, localCtx)
						}
					}
				}
			}
			out, _ := a.arrayTransfer(file, proc, localCtx, variables, in, text, line, constants, nil)
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				out = applyArrayModuleCallEffects(out, file, proc, call, localCtx, variables, moduleDecls)
			}
			return out
		}
		if proc.Graph == nil {
			state := initial
			for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
				state = visit(normalizedCodeLine(file.Lines[line-1]), line, state)
			}
			return evidence
		}
		if ctx.arrayStats != nil {
			ctx.arrayStats.addCFGWalk()
		}
		baseView := proc.Graph.View(vbacfg.EdgeFilter{})
		walkArrayCFGWithEdgesStats(&baseView, file.Lines, initial, visit, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			out = applyArrayConditionalAllocationBranch(out, &baseView, block, edge)
			out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
		}, ctx.arrayStats)
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
	resolution := ctx.procedureResolver.ResolveCall(call)
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return "", sourceProcedure{}, false
	}
	key := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
	target, ok := targets[key]
	return key, target, ok
}

func procedureHasByRefArrayParameter(proc sourceProcedure) bool {
	for parameter := range proc.Params.All() {
		if parameterIsByRefArray(parameter) {
			return true
		}
	}
	return false
}

func arrayRecordByRefCall(evidence map[string]map[int]arrayByRefEntryEvidence, targetKey string, target, caller sourceProcedure, call procedureir.CallSite, state arrayFlowState, ctx analysisContext) {
	// A self-recursive ByRef helper preserves the entry array state supplied by
	// its caller. Treating the recursive edge as an independent unknown entry
	// would poison the evidence from the allocated external call and keep the
	// callee conservative forever.
	if strings.EqualFold(targetKey, arrayProcedureKey(caller)) {
		return
	}
	arguments := arrayCallArgumentTexts(caller, call)
	if len(call.Arguments.Named) > 0 || len(arguments) != call.Arguments.Count {
		return
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) {
			continue
		}
		name := ""
		if index < len(arguments) {
			name = directArrayArgumentName(arguments[index])
		}
		value, known := state[name]
		allocated := known && value.kind == arrayAllocated && value.knownArray
		if !allocated && index < len(arguments) {
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

func arrayByRefCallHasProvenArrayArguments(target, caller sourceProcedure, call procedureir.CallSite, state arrayFlowState, ctx analysisContext) (bool, bool) {
	arguments := arrayCallArgumentTexts(caller, call)
	if len(call.Arguments.Named) > 0 || len(arguments) != call.Arguments.Count {
		return false, false
	}
	foundExpression := false
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) {
			continue
		}
		if index >= len(arguments) {
			return false, false
		}
		argument := arguments[index]
		if name := directArrayArgumentName(argument); name != "" {
			value, known := state[name]
			if !known || value.kind != arrayAllocated || !value.knownArray {
				return false, false
			}
			continue
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

func parameterIsByRefArray(parameter parameterInfo) bool {
	return parameterIsArray(parameter) && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
}

func parameterIsByRefScalar(parameter parameterInfo) bool {
	return !parameterIsArray(parameter) && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
}

func (a Analyzer) arrayTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) (arrayFlowState, []Finding) {
	if state == nil {
		state = arrayFlowState{}
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
			// A Variant can be resized only after its array nature is proven by
			// an assignment such as Array(...).  An unproven Variant remains
			// unknown and must not be treated as a scalar ReDim misuse.
			variantArray := variable.isVariant && old.knownArray
			resizable := (variable.isArray || variantArray) && !variable.fixed
			if variable.isVariant && !old.knownArray {
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

	for _, use := range arrayIndexedUses(text, variables) {
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
				if start, ok := integerLiteral(match[1]); ok && start != value.dimensions[0].lower.value {
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
			} else if value, tracked := state[name]; tracked && value.allocationProbe != "" {
				value.allocationProbe = ""
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
			if value, known := arrayExpressionState(rhs, state, ctx); known {
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
					state[name] = arrayValue{kind: arrayUnknown, knownArray: false, origin: arrayOriginUnknown}
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
		if key == "" || !declaration.IsArray || declaration.ValueShape != procedureir.ValueShapeDynamicArray || !arrayDeclarationHasInlineRedim(file.Lines, declaration) {
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
	if arrayByteArrayReadRe.MatchString(rhs) {
		return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}, true
	}
	if value, ok := state[name]; ok && value.kind == arrayAllocated && value.knownArray {
		return value, true
	}
	if value, ok := ctx.arrayReturns[name]; ok && value.kind == arrayAllocated && value.knownArray {
		return value, true
	}
	if name != "" && strings.Contains(rhs, "(") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	if strings.EqualFold(strings.TrimSpace(rhs), "empty") || strings.EqualFold(strings.TrimSpace(rhs), "nothing") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	return arrayValue{}, false
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

func inferArrayReturnSummaries(files []parsedFile, arrayAllocationGuards map[string]bool, participantCtx analysisContext) map[string]arrayValue {
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
		ctx := analysisContext{arrayReturns: arrayReturns}
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
			return candidate{}, false
		}
		valid := returns[0].ok
		value := returns[0].value
		for _, returned := range returns[1:] {
			if !returned.ok || !arrayValueCompatible(value, returned.value) {
				valid = false
				break
			}
			value = meetArrayValue(value, returned.value)
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
		key := arrayProcedureKey(procedure.proc)
		value, hasContribution := evaluate(procedure, summaries)
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
	return summaries
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

func arrayValueCompatible(left, right arrayValue) bool {
	return left.kind == right.kind && left.knownArray == right.knownArray && left.origin == right.origin && left.allocationProbe == right.allocationProbe && left.allocationCountSource == right.allocationCountSource && arrayDimensionsEqual(left.dimensions, right.dimensions) && arrayDimensionsEqual(left.preserveShape, right.preserveShape)
}
