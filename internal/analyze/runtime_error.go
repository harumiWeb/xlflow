package analyze

import (
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var runtimeConstAssignmentRe = regexp.MustCompile(`(?i)^\s*(?:(?:public|private|friend|static)\s+)?const\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s+as\s+[A-Za-z_][A-Za-z0-9_.]*(?:\s*\([^)]*\))?)?\s*=\s*(.+)$`)
var runtimeConversionCallRe = regexp.MustCompile(`(?i)^\s*(cbyte|cint|clng|clnglng|csng|cdbl|ccur|cdec)\s*\(`)
var runtimeScalarArithmeticRe = regexp.MustCompile(`(?i)^\s*([A-Za-z_][A-Za-z0-9_]*)\s*([+-])\s*(\d+)\s*$`)
var runtimeForPositiveScalarBoundRe = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*1\s+to\s+([A-Za-z_][A-Za-z0-9_]*)\s*$`)

// RuntimeErrorContext is the stable, machine-readable evidence attached to a
// deterministic runtime finding.  The registry evidence class remains the
// source of truth for severity and preflight behavior; this context only
// identifies the proven runtime failure kind.
type RuntimeErrorContext struct {
	Kind string `json:"kind"`
}

type runtimeConstantState map[string]constexpr.Value

// deterministicRuntimeErrorFindingsWithArrayResult reports only failures that
// the shared constant evaluator can prove from a complete expression, projects
// the array facts from the procedure-local result when available, and retains
// a nil-result fallback for focused helper callers that do not run through
// procedure orchestration. Runtime values and unresolved identifiers stay
// unknown by construction.
func (a Analyzer) deterministicRuntimeErrorFindingsWithArrayResult(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, arrayResult *ArrayAnalysisResult) []Finding {
	if enabled, known := config.AnalyzeRuleEnabled(a.Config.Analyze, "VBA249"); !known || !enabled {
		return nil
	}

	facts := proc.analysisFacts()
	localNames := runtimeLocalNames(proc)
	var base constexpr.Environment
	if file.RuntimeConstantBase != nil {
		base = runtimeConstantScope{base: file.RuntimeConstantBase, hidden: localNames}
	} else {
		// Focused helper callers may omit the revision-scoped cache. Preserve the
		// compatibility path while production batch/realtime analysis normalizes
		// the large project environment only once per file.
		values := make(map[string]constexpr.Value, len(a.visibleConstantValues)+len(file.ConstantValues)+8)
		for name, value := range a.visibleConstantValues {
			values[name] = value
		}
		if file.ConstantValues == nil {
			for name, value := range lint.ConstantValuesFromSource(string(file.Source), &file.IR, values) {
				values[name] = value
			}
		} else {
			for name, value := range file.ConstantValues {
				values[name] = value
			}
		}
		for name := range localNames {
			delete(values, name)
		}
		base = constexpr.NewValues(values)
	}
	initial := runtimeLocalConstantState(file, proc, base)
	findings := make([]Finding, 0)
	seen := make(map[string]bool)
	if proc.Graph == nil {
		state := initial
		writes := runtimeWriteNames(proc.Accesses)
		// The canonical ProcedureIR preserves source order, so no projection
		// sort or temporary statement collection is needed here.
		for statement := range proc.Statements.All() {
			env := runtimeConstantEnvironment(base, state)
			findings = appendRuntimeStatementFindings(findings, seen, a, file, proc, statement, facts, env)
			state = runtimeTransfer(statement, state, env, writes[statement.ID])
		}
		if arrayResult != nil {
			return append(findings, arrayResult.runtime()...)
		}
		return append(findings, a.deterministicArrayRuntimeFindings(file, proc, ctx, moduleDecls)...)
	}

	states := runtimeCFGStates(proc.Graph, initial, base, proc.Accesses)
	blocks := append([]vbacfg.Block(nil), proc.Graph.Blocks...)
	sort.SliceStable(blocks, func(i, j int) bool { return blocks[i].ID < blocks[j].ID })
	for _, block := range blocks {
		if block.Statement == nil {
			continue
		}
		state, ok := states[block.ID]
		if !ok {
			continue
		}
		env := runtimeConstantEnvironment(base, state)
		findings = appendRuntimeStatementFindings(findings, seen, a, file, proc, *block.Statement, facts, env)
	}
	if arrayResult != nil {
		return append(findings, arrayResult.runtime()...)
	}
	return append(findings, a.deterministicArrayRuntimeFindings(file, proc, ctx, moduleDecls)...)
}

func runtimeLocalNames(proc sourceProcedure) map[string]bool {
	names := make(map[string]bool)
	for declaration := range proc.Declarations.All() {
		if name := runtimeSimpleIdentifier(declaration.Name); name != "" {
			names[name] = true
		}
	}
	for parameter := range proc.Params.All() {
		if name := runtimeSimpleIdentifier(parameter.Name); name != "" {
			names[name] = true
		}
	}
	for statement := range proc.Statements.All() {
		if name, ok := runtimeAssignmentTarget(statement); ok {
			names[name] = true
		}
	}
	return names
}

// deterministicArrayRuntimeFindings projects only the array facts that are
// already proven by the shared VBA227 allocation/shape lattice. It intentionally
// does not duplicate that rule's possible-failure checks: unknown allocation,
// unknown Variant shape, dynamic bounds, and external values remain silent.
func (a Analyzer) deterministicArrayRuntimeFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []Finding {
	variables := arrayVariables(file, proc, moduleDecls)
	hasArray := false
	for _, variable := range variables {
		if variable.isArray {
			hasArray = true
			break
		}
	}
	if !hasArray {
		return nil
	}
	constants := arrayIntegerConstants(file, proc, a.visibleConstantValues, a.visibleConstants)
	base := arrayOptionBase(file)
	state := arrayInitialState(variables)
	findings := make([]Finding, 0)
	seen := map[string]bool{}
	runtimeStopLines := map[int]bool{}
	probe := a
	probe.Config.Analyze.DetectArrayLifecycleSafety = true
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		lineIssues := a.deterministicArrayRuntimeIssues(file, text, line, in, variables, constants, proc, base, ctx, moduleDecls)
		for _, issue := range lineIssues {
			key := strconv.Itoa(issue.line) + ":" + issue.kind + ":" + issue.operationKey
			if seen[key] {
				continue
			}
			seen[key] = true
			message, reason, suggestion := deterministicRuntimeFailureText(issue.kind)
			finding := a.simpleFinding(file, proc, issue.line, "VBA249", "error", message, reason, suggestion)
			finding.RuntimeError = &RuntimeErrorContext{Kind: issue.kind}
			finding.arrayOperationKey = issue.operationKey
			findings = append(findings, finding)
		}
		_, compound := runtimeArrayTextShape(text)
		if compound {
			// The runtime-only fallback must retain the same source-order compound
			// proof as the shared array lane. Otherwise a block-level entry state
			// reports uses in a zero-count loop or after a successful bounds header,
			// even though those statements cannot execute on the failing path.
			out, _, reachable := probe.deterministicArrayRuntimeCompoundAnalysis(file, text, line, in, variables, constants, proc, base, ctx, moduleDecls)
			if !reachable {
				runtimeStopLines[line] = true
			}
			return out
		}
		out, _ := probe.arrayTransfer(file, proc, ctx, variables, in, text, line, constants, nil)
		forEachArrayCallAtLine(proc, line, func(call procedureir.CallSite) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			if _, _, conditional := arrayIfThenParts(text); conditional {
				out = applyArrayConditionalByRefCallEffects(out, proc, call, ctx)
			} else {
				out = applyArrayByRefCallEffects(out, proc, call, ctx)
			}
		}, ctx.arrayStats)
		if runtimeArrayLineHasFatalIssue(file, text, in, lineIssues, variables, constants, proc, line) {
			runtimeStopLines[line] = true
		}
		return applyArrayForBoundHeaderState(out, text, variables)
	}
	if proc.Graph != nil {
		view := proc.Graph.View(vbacfg.EdgeFilter{})
		walkArrayCFGWithStopStats(&view, file.Lines, state, visit, nil, func(_ string, line int) bool {
			return runtimeStopLines[line]
		}, nil)
	} else {
		start := proc.StartLine
		if start < 1 {
			start = 1
		}
		end := proc.EndLine
		if end > len(file.Lines) {
			end = len(file.Lines)
		}
		if start <= end {
			_ = visit(strings.Join(file.Lines[start-1:end], "\n"), start, state)
		}
	}
	sortFindings(findings)
	return findings
}

type deterministicArrayRuntimeIssue struct {
	line         int
	kind         string
	operationKey string
}

func runtimeArrayCallLines(proc sourceProcedure) map[int]bool {
	lines := make(map[int]bool)
	for call := range proc.Calls.All() {
		if call.Range.StartLine > 0 {
			lines[call.Range.StartLine] = true
		}
	}
	return lines
}

func arrayBoundOperationKey(kind, name, failure string) string {
	return "bound:" + strings.ToLower(strings.TrimSpace(kind)) + ":" + strings.ToLower(strings.TrimSpace(name)) + ":" + strings.ToLower(strings.TrimSpace(failure))
}

func arrayIndexOperationKey(name, failure string) string {
	return "index:" + strings.ToLower(strings.TrimSpace(name)) + ":" + strings.ToLower(strings.TrimSpace(failure))
}

func (a Analyzer) deterministicArrayRuntimeIssues(file parsedFile, text string, line int, state arrayFlowState, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, base int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []deterministicArrayRuntimeIssue {
	selectCase, compound := runtimeArrayTextShape(text)
	if selectCase {
		return a.deterministicArrayRuntimeBlockIssues(file, text, line, state, variables, constants, proc, base, ctx, moduleDecls)
	}
	if compound {
		return a.deterministicArrayRuntimeCompoundIssues(file, text, line, state, variables, constants, proc, base, ctx, moduleDecls)
	}
	return deterministicArrayRuntimeLineIssues(file, text, line, state, variables, constants, proc, base)
}

func runtimeArrayTextShape(text string) (selectCase, compound bool) {
	trimmed := strings.TrimSpace(text)
	if trimmed == "" {
		return false, false
	}
	if !strings.ContainsAny(trimmed, "\r\n") {
		lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(trimmed)))
		selectCase = strings.HasPrefix(lower, "select case ")
		return selectCase, selectCase
	}
	lines := normalizedSourceLines(text)
	for _, rawLine := range lines {
		lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(rawLine)))
		if lower == "" {
			continue
		}
		selectCase = strings.HasPrefix(lower, "select case ")
		break
	}
	return selectCase, len(lines) > 1
}

// deterministicArrayRuntimeBlockIssues evaluates the physical lines owned by
// a Select Case block with one independent state per top-level Case. The CFG
// can expose a whole Case tree as one statement; scanning that statement from
// its entry state mistakes branch-local ReDim and assignment facts for
// deterministic failures (notably ReDim expressions themselves). Nested
// control flow is deliberately skipped while preserving the state before it,
// so only straight-line facts are promoted to the high-precision runtime
// projection.
func (a Analyzer) deterministicArrayRuntimeBlockIssues(file parsedFile, text string, startLine int, entry arrayFlowState, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, base int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []deterministicArrayRuntimeIssue {
	lines := normalizedSourceLines(text)
	if len(lines) < 2 {
		return deterministicArrayRuntimeLineIssues(file, text, startLine, entry, variables, constants, proc, base)
	}
	probe := a
	probe.Config.Analyze.DetectArrayLifecycleSafety = true
	issues := make([]deterministicArrayRuntimeIssue, 0)
	processLine := func(state *arrayFlowState, raw string, line int) {
		text := normalizedCodeLine(raw)
		if strings.TrimSpace(text) == "" {
			return
		}
		issues = append(issues, deterministicArrayRuntimeLineIssues(file, text, line, *state, variables, constants, proc, base)...)
		transferText := text
		if redim, ok := inlineArrayRedimText(text); ok {
			transferText = redim
		}
		before := cloneArrayState(*state)
		out, _ := probe.arrayTransfer(file, proc, ctx, variables, *state, transferText, line, constants, nil)
		if strings.HasPrefix(strings.ToLower(strings.TrimSpace(transferText)), "redim preserve") {
			out = restoreRuntimePreserveState(out, before, transferText, variables, proc, line)
		}
		forEachArrayCallAtLine(proc, line, func(call procedureir.CallSite) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			// The compound scanner evaluates an If/ElseIf header once and then
			// starts its branches from that shared state. Do not apply a ByRef
			// invalidation from a call embedded in the header: doing so loses a
			// known-unallocated output on the true branch. Unconditional calls
			// retain the normal invalidation and Erase restoration handling.
			if _, _, conditional := arrayIfThenParts(text); !conditional {
				out = applyArrayByRefCallEffects(out, proc, call, ctx)
			}
			out = restoreDeterministicArrayEraseState(out, proc, call, ctx)
		}, ctx.arrayStats)
		*state = out
	}

	baseState := cloneArrayState(entry)
	firstCase := -1
	for index, rawLine := range lines {
		lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(rawLine)))
		if strings.HasPrefix(lower, "case ") {
			firstCase = index
			break
		}
	}
	if firstCase < 0 {
		return deterministicArrayRuntimeLineIssues(file, text, startLine, entry, variables, constants, proc, base)
	}
	for index := 0; index < firstCase; index++ {
		processLine(&baseState, lines[index], startLine+index)
	}

	active := false
	branchState := cloneArrayState(baseState)
	controlStack := make([]string, 0, 4)
	for index := firstCase; index < len(lines); index++ {
		line := normalizedCodeLine(lines[index])
		lower := strings.ToLower(strings.TrimSpace(line))
		if lower == "" {
			continue
		}
		if strings.HasPrefix(lower, "case ") && len(controlStack) == 0 {
			active = true
			branchState = cloneArrayState(baseState)
			continue
		}
		if isRuntimeArrayBlockControlEnd(lower) {
			if len(controlStack) > 0 {
				controlStack = controlStack[:len(controlStack)-1]
				continue
			}
			break
		}
		if !active {
			continue
		}
		if len(controlStack) > 0 {
			if kind, starts := runtimeArrayBlockControlStart(lower); starts {
				controlStack = append(controlStack, kind)
			}
			continue
		}
		processLine(&branchState, line, startLine+index)
		if kind, starts := runtimeArrayBlockControlStart(lower); starts {
			controlStack = append(controlStack, kind)
		}
	}
	return issues
}

type runtimeArrayCompoundScanner struct {
	analyzer    Analyzer
	file        parsedFile
	variables   map[string]arrayVariable
	constants   map[string]int
	proc        sourceProcedure
	base        int
	ctx         analysisContext
	moduleDecls map[string]sourceDeclaration
	probe       Analyzer
}

type runtimeScalarValue struct {
	known bool
	value int
}

type runtimeScalarState map[string]runtimeScalarValue

type runtimeArrayCompoundFlow struct {
	arrays  arrayFlowState
	scalars runtimeScalarState
	guards  map[string]bool
}

func cloneRuntimeScalarState(state runtimeScalarState) runtimeScalarState {
	out := runtimeScalarState{}
	for name, value := range state {
		out[name] = value
	}
	return out
}

func cloneRuntimeArrayCompoundFlow(flow runtimeArrayCompoundFlow) runtimeArrayCompoundFlow {
	return runtimeArrayCompoundFlow{
		arrays:  cloneArrayState(flow.arrays),
		scalars: cloneRuntimeScalarState(flow.scalars),
		guards:  cloneRuntimeConditionGuards(flow.guards),
	}
}

func cloneRuntimeConditionGuards(guards map[string]bool) map[string]bool {
	out := map[string]bool{}
	for condition, knownTrue := range guards {
		if knownTrue {
			out[condition] = true
		}
	}
	return out
}

func meetRuntimeConditionGuards(left, right map[string]bool) map[string]bool {
	out := map[string]bool{}
	for condition, knownTrue := range left {
		if knownTrue && right[condition] {
			out[condition] = true
		}
	}
	return out
}

// runtimeConditionKey deliberately normalizes only the syntax needed for
// matching a branch guard with a loop condition. It does not attempt to
// evaluate arbitrary VBA expressions.
func runtimeConditionKey(text string) string {
	lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(text)))
	condition := ""
	switch {
	case strings.HasPrefix(lower, "if "):
		condition = strings.TrimSpace(strings.TrimPrefix(lower, "if "))
		condition = strings.TrimSpace(strings.TrimSuffix(condition, " then"))
	case strings.HasPrefix(lower, "elseif "):
		condition = strings.TrimSpace(strings.TrimPrefix(lower, "elseif "))
		condition = strings.TrimSpace(strings.TrimSuffix(condition, " then"))
	case strings.HasPrefix(lower, "do while "):
		condition = strings.TrimSpace(strings.TrimPrefix(lower, "do while "))
	case strings.HasPrefix(lower, "while "):
		condition = strings.TrimSpace(strings.TrimPrefix(lower, "while "))
	default:
		return ""
	}
	condition = strings.ReplaceAll(condition, " ", "")
	condition = strings.ReplaceAll(condition, "\t", "")
	condition = strings.ReplaceAll(condition, "(", "")
	condition = strings.ReplaceAll(condition, ")", "")
	return condition
}

func runtimeConditionGuardsAfterLine(guards map[string]bool, text string, variables map[string]arrayVariable) map[string]bool {
	updated := cloneRuntimeConditionGuards(guards)
	for _, rawPart := range splitRangeValueSourceStatements(text) {
		lhs, _, indexed, assigned := arrayAssignment(strings.TrimSpace(rawPart))
		if !assigned || indexed {
			continue
		}
		name := strings.ToLower(runtimeSimpleIdentifier(lhs))
		if name == "" {
			continue
		}
		if variable, known := variables[name]; known && (variable.isArray || variable.isVariant) {
			continue
		}
		for condition := range updated {
			if runtimeConditionMentionsName(condition, name) {
				delete(updated, condition)
			}
		}
	}
	return updated
}

func runtimeConditionMentionsName(condition, name string) bool {
	condition = strings.ToLower(condition)
	name = strings.ToLower(name)
	for index := 0; index < len(condition); {
		if !isIdentifierStart(condition[index]) {
			index++
			continue
		}
		start := index
		index++
		for index < len(condition) && isIdentifierPart(condition[index]) {
			index++
		}
		if condition[start:index] == name {
			return true
		}
	}
	return false
}

type runtimeConditionScope struct {
	kind       string
	condition  string
	trueBranch bool
}

// runtimeArrayConditionGuardsBefore reconstructs only the lexical branch
// guards that enclose a CFG statement. The ProcedureIR may expose a nested
// loop as an independent statement, so the compound scanner cannot otherwise
// see the outer If that proves its first iteration executes. Assignment
// invalidation keeps a guard from surviving a mutation of its condition.
func runtimeArrayConditionGuardsBefore(file parsedFile, proc sourceProcedure, line int) map[string]bool {
	guards := map[string]bool{}
	stack := make([]runtimeConditionScope, 0, 4)
	start := proc.StartLine
	if start < 1 {
		start = 1
	}
	end := line
	if end > len(file.Lines)+1 {
		end = len(file.Lines) + 1
	}
	for sourceLine := start; sourceLine < end; sourceLine++ {
		text := normalizedCodeLine(file.Lines[sourceLine-1])
		lower := strings.ToLower(strings.TrimSpace(text))
		if lower == "" {
			continue
		}
		if endKind, ends := runtimeArrayCompoundBlockEndKind(lower); ends {
			for index := len(stack) - 1; index >= 0; index-- {
				if stack[index].kind != endKind {
					continue
				}
				if stack[index].trueBranch && stack[index].condition != "" {
					delete(guards, stack[index].condition)
				}
				stack = stack[:index]
				break
			}
			continue
		}
		if strings.HasPrefix(lower, "elseif ") {
			if len(stack) > 0 && stack[len(stack)-1].kind == "if" {
				previous := stack[len(stack)-1]
				if previous.trueBranch && previous.condition != "" {
					delete(guards, previous.condition)
				}
				condition := runtimeConditionKey(text)
				stack[len(stack)-1] = runtimeConditionScope{kind: "if", condition: condition, trueBranch: condition != ""}
				if condition != "" {
					guards[condition] = true
				}
			}
			continue
		}
		if lower == "else" || strings.HasPrefix(lower, "else ") {
			if len(stack) > 0 && stack[len(stack)-1].kind == "if" {
				previous := stack[len(stack)-1]
				if previous.trueBranch && previous.condition != "" {
					delete(guards, previous.condition)
				}
				stack[len(stack)-1].trueBranch = false
			}
			continue
		}
		if kind, starts := runtimeArrayBlockControlStart(lower); starts {
			condition := ""
			trueBranch := false
			if kind == "if" {
				condition = runtimeConditionKey(text)
				trueBranch = condition != ""
				if trueBranch {
					guards[condition] = true
				}
			}
			stack = append(stack, runtimeConditionScope{kind: kind, condition: condition, trueBranch: trueBranch})
		}
		guards = runtimeConditionGuardsAfterLine(guards, text, nil)
	}
	return guards
}

func meetRuntimeScalarState(left, right runtimeScalarState) runtimeScalarState {
	out := runtimeScalarState{}
	for name, leftValue := range left {
		rightValue, ok := right[name]
		if !ok {
			out[name] = runtimeScalarValue{}
			continue
		}
		if leftValue.known && rightValue.known && leftValue.value == rightValue.value {
			out[name] = leftValue
		} else {
			out[name] = runtimeScalarValue{}
		}
	}
	for name := range right {
		if _, ok := left[name]; !ok {
			out[name] = runtimeScalarValue{}
		}
	}
	return out
}

func runtimeScalarTransfer(state runtimeScalarState, text string, variables map[string]arrayVariable) runtimeScalarState {
	updated := cloneRuntimeScalarState(state)
	for _, rawPart := range splitRangeValueSourceStatements(text) {
		part := strings.TrimSpace(rawPart)
		lhs, rhs, indexed, ok := arrayAssignment(part)
		if !ok || indexed {
			continue
		}
		name := runtimeSimpleIdentifier(lhs)
		if name == "" {
			continue
		}
		if variable, known := variables[name]; known && (variable.isArray || variable.isVariant) {
			continue
		}
		if value, literal := integerLiteral(rhs); literal {
			updated[name] = runtimeScalarValue{known: true, value: value}
			continue
		}
		if match := runtimeScalarArithmeticRe.FindStringSubmatch(rhs); len(match) == 4 {
			baseName := runtimeSimpleIdentifier(match[1])
			base, known := updated[baseName]
			delta, parsed := strconv.Atoi(match[3])
			if known && base.known && parsed == nil {
				if match[2] == "-" {
					delta = -delta
				}
				updated[name] = runtimeScalarValue{known: true, value: base.value + delta}
				continue
			}
		}
		if sourceName := runtimeSimpleIdentifier(rhs); sourceName != "" {
			if source, known := updated[sourceName]; known {
				updated[name] = source
				continue
			}
		}
		updated[name] = runtimeScalarValue{}
	}
	return updated
}

func runtimeScalarForHeaderHasZeroUpper(text string, scalars runtimeScalarState) bool {
	match := runtimeForPositiveScalarBoundRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(text)))
	if len(match) != 2 {
		return false
	}
	value, known := scalars[runtimeSimpleIdentifier(match[1])]
	return known && value.known && value.value == 0
}

type runtimeArrayIfBranch struct {
	marker    int
	bodyStart int
	bodyEnd   int
}

func (a Analyzer) deterministicArrayRuntimeCompoundIssues(file parsedFile, text string, startLine int, entry arrayFlowState, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, base int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []deterministicArrayRuntimeIssue {
	_, issues := a.deterministicArrayRuntimeCompoundState(file, text, startLine, entry, variables, constants, proc, base, ctx, moduleDecls)
	return issues
}

func (a Analyzer) deterministicArrayRuntimeCompoundState(file parsedFile, text string, startLine int, entry arrayFlowState, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, base int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) (arrayFlowState, []deterministicArrayRuntimeIssue) {
	state, issues, _ := a.deterministicArrayRuntimeCompoundAnalysis(file, text, startLine, entry, variables, constants, proc, base, ctx, moduleDecls)
	return state, issues
}

func (a Analyzer) deterministicArrayRuntimeCompoundAnalysis(file parsedFile, text string, startLine int, entry arrayFlowState, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, base int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) (arrayFlowState, []deterministicArrayRuntimeIssue, bool) {
	lines := normalizedSourceLines(text)
	scanner := runtimeArrayCompoundScanner{
		analyzer:    a,
		file:        file,
		variables:   variables,
		constants:   constants,
		proc:        proc,
		base:        base,
		ctx:         ctx,
		moduleDecls: moduleDecls,
		probe:       a,
	}
	scanner.probe.Config.Analyze.DetectArrayLifecycleSafety = true
	flow, issues, reachable := scanner.scanSequence(lines, 0, len(lines), startLine, runtimeArrayCompoundFlow{
		arrays:  cloneArrayState(entry),
		scalars: runtimeScalarState{},
		guards:  runtimeArrayConditionGuardsBefore(file, proc, startLine),
	})
	return flow.arrays, issues, reachable
}

func (s *runtimeArrayCompoundScanner) processLine(flow runtimeArrayCompoundFlow, raw string, line int) (runtimeArrayCompoundFlow, []deterministicArrayRuntimeIssue, bool) {
	text := normalizedCodeLine(raw)
	if strings.TrimSpace(text) == "" {
		return flow, nil, false
	}
	issues := deterministicArrayRuntimeLineIssues(s.file, text, line, flow.arrays, s.variables, s.constants, s.proc, s.base)
	transferText := text
	if redim, ok := inlineArrayRedimText(text); ok {
		transferText = redim
	}
	before := cloneArrayState(flow.arrays)
	out, _ := s.probe.arrayTransfer(s.file, s.proc, s.ctx, s.variables, flow.arrays, transferText, line, s.constants, nil)
	if strings.HasPrefix(strings.ToLower(strings.TrimSpace(transferText)), "redim preserve") {
		out = restoreRuntimePreserveState(out, before, transferText, s.variables, s.proc, line)
	}
	forEachArrayCallAtLine(s.proc, line, func(call procedureir.CallSite) {
		out = applyArrayModuleCallEffects(out, s.file, s.proc, call, s.ctx, s.variables, s.moduleDecls)
		out = applyArrayUnknownModuleCallEffects(out, s.file, s.proc, call, s.ctx, s.variables, s.moduleDecls)
		// An If/ElseIf header is shared by all branches in this scanner. Join a
		// conditional ByRef result conservatively: a callee that can both leave
		// an unallocated output and produce an allocated one becomes unknown,
		// while a no-result helper that only preserves the unallocated state keeps
		// the deterministic VBA249 proof.
		if _, _, conditional := arrayIfThenParts(text); conditional {
			out = applyArrayConditionalByRefCallEffects(out, s.proc, call, s.ctx)
		} else {
			out = applyArrayByRefCallEffects(out, s.proc, call, s.ctx)
		}
		out = restoreDeterministicArrayEraseState(out, s.proc, call, s.ctx)
	}, s.ctx.arrayStats)
	if runtimeArrayLineHasFatalIssue(s.file, text, before, issues, s.variables, s.constants, s.proc, line) {
		return runtimeArrayCompoundFlow{arrays: out, scalars: cloneRuntimeScalarState(flow.scalars), guards: cloneRuntimeConditionGuards(flow.guards)}, issues, true
	}
	return runtimeArrayCompoundFlow{
		arrays:  out,
		scalars: runtimeScalarTransfer(flow.scalars, text, s.variables),
		guards:  runtimeConditionGuardsAfterLine(flow.guards, text, s.variables),
	}, issues, false
}

func (s *runtimeArrayCompoundScanner) scanSequence(lines []string, start, end, startLine int, entry runtimeArrayCompoundFlow) (runtimeArrayCompoundFlow, []deterministicArrayRuntimeIssue, bool) {
	flow := cloneRuntimeArrayCompoundFlow(entry)
	issues := make([]deterministicArrayRuntimeIssue, 0)
	for index := start; index < end; {
		line := normalizedCodeLine(lines[index])
		lower := strings.ToLower(strings.TrimSpace(line))
		if lower == "" {
			index++
			continue
		}
		if kind, starts := runtimeArrayBlockControlStart(lower); starts {
			var next int
			var nestedIssues []deterministicArrayRuntimeIssue
			var nestedReachable bool
			switch kind {
			case "if":
				next, flow, nestedIssues, nestedReachable = s.scanIf(lines, index, end, startLine, flow)
			case "select":
				next, flow, nestedIssues, nestedReachable = s.scanSelect(lines, index, end, startLine, flow)
			case "for", "do", "while", "with":
				next, flow, nestedIssues, nestedReachable = s.scanLoop(lines, index, end, startLine, flow, kind)
			default:
				next = index
			}
			if next > index {
				issues = append(issues, nestedIssues...)
				if !nestedReachable {
					return flow, issues, false
				}
				index = next
				continue
			}
		}
		var lineIssues []deterministicArrayRuntimeIssue
		var terminal bool
		flow, lineIssues, terminal = s.processLine(flow, line, startLine+index)
		issues = append(issues, lineIssues...)
		if terminal {
			return flow, issues, false
		}
		index++
	}
	return flow, issues, true
}

func (s *runtimeArrayCompoundScanner) scanIf(lines []string, start, limit, startLine int, entry runtimeArrayCompoundFlow) (int, runtimeArrayCompoundFlow, []deterministicArrayRuntimeIssue, bool) {
	headerState, issues, headerTerminal := s.processLine(entry, lines[start], startLine+start)
	if headerTerminal {
		return start, headerState, issues, false
	}
	branches, end, ok := runtimeArrayIfBranches(lines, start, limit)
	if !ok {
		return start, headerState, issues, true
	}
	hasElse := false
	for _, branch := range branches {
		if branch.marker >= 0 {
			marker := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[branch.marker])))
			if marker == "else" || strings.HasPrefix(marker, "else ") {
				hasElse = true
			}
		}
	}
	merged := cloneRuntimeArrayCompoundFlow(headerState)
	mergedReachable := !hasElse
	for _, branch := range branches {
		branchState := cloneRuntimeArrayCompoundFlow(headerState)
		branchReachable := true
		guard := ""
		if branch.marker < 0 {
			guard = runtimeConditionKey(lines[start])
		}
		if branch.marker >= 0 && strings.HasPrefix(strings.ToLower(strings.TrimSpace(lines[branch.marker])), "elseif ") {
			var markerIssues []deterministicArrayRuntimeIssue
			var markerTerminal bool
			branchState, markerIssues, markerTerminal = s.processLine(branchState, lines[branch.marker], startLine+branch.marker)
			issues = append(issues, markerIssues...)
			branchReachable = !markerTerminal
			guard = runtimeConditionKey(lines[branch.marker])
		}
		if branchReachable && guard != "" {
			branchState.guards[guard] = true
		}
		var branchIssues []deterministicArrayRuntimeIssue
		if branchReachable {
			branchState, branchIssues, branchReachable = s.scanSequence(lines, branch.bodyStart, branch.bodyEnd, startLine, branchState)
		}
		issues = append(issues, branchIssues...)
		if branchReachable {
			if mergedReachable {
				merged.arrays = meetArrayState(merged.arrays, branchState.arrays)
				merged.scalars = meetRuntimeScalarState(merged.scalars, branchState.scalars)
				merged.guards = meetRuntimeConditionGuards(merged.guards, branchState.guards)
			} else {
				merged = branchState
			}
			mergedReachable = true
		}
	}
	return end + 1, merged, issues, mergedReachable
}

func (s *runtimeArrayCompoundScanner) scanSelect(lines []string, start, limit, startLine int, entry runtimeArrayCompoundFlow) (int, runtimeArrayCompoundFlow, []deterministicArrayRuntimeIssue, bool) {
	headerState, issues, headerTerminal := s.processLine(entry, lines[start], startLine+start)
	if headerTerminal {
		return start, headerState, issues, false
	}
	branches, end, ok := runtimeArraySelectBranches(lines, start, limit)
	if !ok {
		return start, headerState, issues, true
	}
	hasCaseElse := false
	for _, branch := range branches {
		if branch.marker >= 0 {
			marker := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[branch.marker])))
			if marker == "case else" {
				hasCaseElse = true
			}
		}
	}
	merged := cloneRuntimeArrayCompoundFlow(headerState)
	mergedReachable := !hasCaseElse
	for _, branch := range branches {
		branchState := cloneRuntimeArrayCompoundFlow(headerState)
		var markerIssues []deterministicArrayRuntimeIssue
		var markerTerminal bool
		branchState, markerIssues, markerTerminal = s.processLine(branchState, lines[branch.marker], startLine+branch.marker)
		issues = append(issues, markerIssues...)
		branchState.arrays = applyArrayConditionalAllocationCaseState(branchState.arrays, lines[start], lines[branch.marker])
		var branchIssues []deterministicArrayRuntimeIssue
		branchReachable := !markerTerminal
		if branchReachable {
			branchState, branchIssues, branchReachable = s.scanSequence(lines, branch.bodyStart, branch.bodyEnd, startLine, branchState)
		}
		issues = append(issues, branchIssues...)
		if branchReachable {
			if mergedReachable {
				merged.arrays = meetArrayState(merged.arrays, branchState.arrays)
				merged.scalars = meetRuntimeScalarState(merged.scalars, branchState.scalars)
				merged.guards = meetRuntimeConditionGuards(merged.guards, branchState.guards)
			} else {
				merged = branchState
			}
			mergedReachable = true
		}
	}
	return end + 1, merged, issues, mergedReachable
}

func (s *runtimeArrayCompoundScanner) scanLoop(lines []string, start, limit, startLine int, entry runtimeArrayCompoundFlow, kind string) (int, runtimeArrayCompoundFlow, []deterministicArrayRuntimeIssue, bool) {
	headerState, issues, headerTerminal := s.processLine(entry, lines[start], startLine+start)
	if headerTerminal {
		return start, headerState, issues, false
	}
	end, ok := runtimeArrayCompoundBlockEnd(lines, start, limit, kind)
	if !ok {
		return start, headerState, issues, true
	}
	loopCondition := runtimeConditionKey(lines[start])
	guaranteedFirstIteration := loopCondition != "" && headerState.guards[loopCondition]
	if kind == "for" && runtimeScalarForHeaderHasZeroUpper(lines[start], headerState.scalars) {
		return end + 1, headerState, issues, true
	}
	bodyState := cloneRuntimeArrayCompoundFlow(headerState)
	if kind == "for" {
		bodyState.arrays = applyArrayForBoundHeaderState(bodyState.arrays, lines[start], s.variables)
	}
	bodyState, bodyIssues, bodyReachable := s.scanSequence(lines, start+1, end, startLine, bodyState)
	issues = append(issues, bodyIssues...)
	if kind == "with" {
		return end + 1, bodyState, issues, bodyReachable
	}
	if !bodyReachable {
		if guaranteedFirstIteration {
			return end + 1, headerState, issues, false
		}
		// For loops may execute zero times, so their header path remains a
		// possible normal continuation even when the first body iteration
		// terminates on a deterministic Preserve failure.
		return end + 1, headerState, issues, true
	}
	return end + 1, runtimeArrayCompoundFlow{
		arrays:  meetArrayState(headerState.arrays, bodyState.arrays),
		scalars: meetRuntimeScalarState(headerState.scalars, bodyState.scalars),
		guards:  meetRuntimeConditionGuards(headerState.guards, bodyState.guards),
	}, issues, true
}

func runtimeArrayIfBranches(lines []string, start, limit int) ([]runtimeArrayIfBranch, int, bool) {
	branches := make([]runtimeArrayIfBranch, 0, 2)
	bodyStart := start + 1
	marker := -1
	stack := make([]string, 0, 4)
	for index := start + 1; index < limit; index++ {
		lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[index])))
		if kind, starts := runtimeArrayBlockControlStart(lower); starts {
			stack = append(stack, kind)
			continue
		}
		if endKind, ends := runtimeArrayCompoundBlockEndKind(lower); ends {
			if len(stack) > 0 {
				if stack[len(stack)-1] == endKind {
					stack = stack[:len(stack)-1]
				}
				continue
			}
			if endKind != "if" {
				continue
			}
			branches = append(branches, runtimeArrayIfBranch{marker: marker, bodyStart: bodyStart, bodyEnd: index})
			return branches, index, true
		}
		if len(stack) == 0 && (strings.HasPrefix(lower, "elseif ") || lower == "else" || strings.HasPrefix(lower, "else ")) {
			branches = append(branches, runtimeArrayIfBranch{marker: marker, bodyStart: bodyStart, bodyEnd: index})
			marker = index
			bodyStart = index + 1
		}
	}
	return nil, start, false
}

func runtimeArraySelectBranches(lines []string, start, limit int) ([]runtimeArrayIfBranch, int, bool) {
	branches := make([]runtimeArrayIfBranch, 0, 2)
	marker := -1
	bodyStart := start + 1
	stack := make([]string, 0, 4)
	for index := start + 1; index < limit; index++ {
		lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[index])))
		if kind, starts := runtimeArrayBlockControlStart(lower); starts {
			stack = append(stack, kind)
			continue
		}
		if endKind, ends := runtimeArrayCompoundBlockEndKind(lower); ends {
			if len(stack) > 0 {
				if stack[len(stack)-1] == endKind {
					stack = stack[:len(stack)-1]
				}
				continue
			}
			if endKind != "select" {
				continue
			}
			if marker >= 0 {
				branches = append(branches, runtimeArrayIfBranch{marker: marker, bodyStart: bodyStart, bodyEnd: index})
			}
			return branches, index, true
		}
		if len(stack) == 0 && strings.HasPrefix(lower, "case ") {
			if marker >= 0 {
				branches = append(branches, runtimeArrayIfBranch{marker: marker, bodyStart: bodyStart, bodyEnd: index})
			}
			marker = index
			bodyStart = index + 1
		}
	}
	return nil, start, false
}

func runtimeArrayCompoundBlockEnd(lines []string, start, limit int, kind string) (int, bool) {
	stack := []string{kind}
	for index := start + 1; index < limit; index++ {
		lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[index])))
		if nested, starts := runtimeArrayBlockControlStart(lower); starts {
			stack = append(stack, nested)
			continue
		}
		endKind, ends := runtimeArrayCompoundBlockEndKind(lower)
		if !ends || len(stack) == 0 {
			continue
		}
		if stack[len(stack)-1] != endKind {
			continue
		}
		stack = stack[:len(stack)-1]
		if len(stack) == 0 {
			return index, true
		}
	}
	return start, false
}

func runtimeArrayCompoundBlockEndKind(lower string) (string, bool) {
	switch {
	case lower == "end if":
		return "if", true
	case lower == "next" || strings.HasPrefix(lower, "next "):
		return "for", true
	case lower == "loop" || strings.HasPrefix(lower, "loop "):
		return "do", true
	case lower == "wend":
		return "while", true
	case lower == "end with":
		return "with", true
	case lower == "end select":
		return "select", true
	default:
		return "", false
	}
}

// restoreDeterministicArrayEraseState narrows the generic ByRef invalidation
// state only for a helper whose reachable body contains an unconditional,
// direct Erase of the passed array. Other invalidations remain unknown so the
// deterministic VBA249 projection does not turn a possible failure into a
// false positive.
func restoreDeterministicArrayEraseState(state arrayFlowState, caller sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !resolved {
		return state
	}
	bindings, mapped := arrayCallArgumentBindings(caller, target, call)
	if !mapped {
		return state
	}
	updated := cloneArrayState(state)
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		if !arrayByRefParameterDefinitelyErases(target, binding.parameterIndex) {
			continue
		}
		name := directArrayArgumentName(binding.text)
		if name == "" {
			continue
		}
		value, tracked := updated[name]
		if !tracked || value.kind != arrayUnknown {
			continue
		}
		updated[name] = arrayValue{kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}
	}
	return updated
}

func arrayByRefParameterDefinitelyErases(proc sourceProcedure, parameterIndex int) bool {
	if parameterIndex < 0 || parameterIndex >= proc.Params.Len() || !parameterIsByRefArray(proc.Params.valueAt(parameterIndex)) {
		return false
	}
	name := strings.ToLower(cleanIdentifier(proc.Params.valueAt(parameterIndex).Name))
	sawErase := false
	for statement := range proc.Statements.All() {
		if !arrayByRefStatementReachable(proc, statement) {
			continue
		}
		for _, rawPart := range splitRangeValueSourceStatements(statement.Text) {
			part := strings.TrimSpace(rawPart)
			if part == "" {
				continue
			}
			lower := strings.ToLower(part)
			if strings.HasPrefix(lower, "if ") || strings.HasPrefix(lower, "elseif ") || lower == "else" || strings.HasPrefix(lower, "else ") ||
				strings.HasPrefix(lower, "select case ") || strings.HasPrefix(lower, "case ") || strings.HasPrefix(lower, "for ") ||
				strings.HasPrefix(lower, "do") || strings.HasPrefix(lower, "while ") || strings.HasPrefix(lower, "goto ") ||
				strings.HasPrefix(lower, "exit ") || strings.HasPrefix(lower, "on error") {
				return false
			}
			match := arrayEraseRe.FindStringSubmatch(part)
			if len(match) == 2 {
				parts := splitArgs(match[1])
				if len(parts) != 1 || !strings.EqualFold(cleanIdentifier(strings.TrimSpace(parts[0])), name) {
					return false
				}
				sawErase = true
			}
		}
	}
	return sawErase
}

func restoreRuntimePreserveState(out, before arrayFlowState, text string, variables map[string]arrayVariable, proc sourceProcedure, line int) arrayFlowState {
	// Resume Next keeps the source-order path alive, but its existing runtime
	// projection intentionally reports only the failed Preserve itself. A
	// labeled handler with Resume Next, on the other hand, transfers through a
	// distinct error path and must retain the pre-Preserve state so the resumed
	// indexed operation can be diagnosed as a second failure.
	if !runtimeArrayPreserveUsesResumableGoTo(proc, line) {
		return out
	}
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 || !strings.EqualFold(strings.TrimSpace(match[1]), "preserve") {
		return out
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct {
			continue
		}
		name := strings.ToLower(redim.name)
		variable, known := variables[name]
		if !known || !variable.isArray || variable.fixed {
			continue
		}
		old, hadOld := before[name]
		if hadOld && (old.kind != arrayAllocated || !old.knownArray) {
			out[name] = old
		}
	}
	return out
}

func runtimeArrayPreserveUsesResumableGoTo(proc sourceProcedure, line int) bool {
	facts := proc.analysisFacts()
	if facts == nil || !facts.runtimeErrorModeEventsBuilt {
		return false
	}
	resumable := false
	for _, event := range facts.runtimeErrorModeEvents {
		if event.line >= line {
			continue
		}
		mode := strings.TrimSpace(event.mode)
		switch {
		case mode == "goto 0", mode == "resume next":
			resumable = false
		case strings.HasPrefix(mode, "goto "):
			resumable = runtimeArrayErrorHandlerMayContinue(proc, event.line, strings.TrimSpace(strings.TrimPrefix(mode, "goto ")))
		default:
			resumable = false
		}
	}
	return resumable
}

func isRuntimeArrayBlockControlEnd(lower string) bool {
	return lower == "end if" || lower == "next" || strings.HasPrefix(lower, "next ") || lower == "loop" || strings.HasPrefix(lower, "loop ") || lower == "wend" || lower == "end with" || lower == "end select"
}

func runtimeArrayBlockControlStart(lower string) (string, bool) {
	switch {
	case strings.HasPrefix(lower, "select case "):
		return "select", true
	case strings.HasPrefix(lower, "if ") && runtimeArrayBlockIsBlockIf(lower):
		return "if", true
	case strings.HasPrefix(lower, "for "):
		return "for", true
	case lower == "do" || strings.HasPrefix(lower, "do while ") || strings.HasPrefix(lower, "do until "):
		return "do", true
	case strings.HasPrefix(lower, "while "):
		return "while", true
	case strings.HasPrefix(lower, "with "):
		return "with", true
	default:
		return "", false
	}
}

func deterministicArrayRuntimeLineIssues(file parsedFile, text string, line int, state arrayFlowState, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, base int) []deterministicArrayRuntimeIssue {
	var issues []deterministicArrayRuntimeIssue
	add := func(kind, operationKey string) {
		issues = append(issues, deterministicArrayRuntimeIssue{line: line, kind: kind, operationKey: operationKey})
	}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		variable, ok := variables[name]
		if !ok || !variable.isArray {
			continue
		}
		value := state[name]
		if value.kind == arrayUnallocated && value.knownArray {
			add("array_unallocated", arrayBoundOperationKey(bound[1], name, "unallocated"))
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			continue
		}
		dimension := 1
		if raw := strings.TrimSpace(bound[3]); raw != "" {
			parsed, err := strconv.Atoi(raw)
			if err != nil {
				continue
			}
			dimension = parsed
		}
		if dimension < 1 || len(value.dimensions) > 0 && dimension > len(value.dimensions) {
			add("array_subscript_out_of_bounds", arrayBoundOperationKey(bound[1], name, "bounds"))
		}
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.EqualFold(strings.TrimSpace(match[1]), "preserve") {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct {
				continue
			}
			name := strings.ToLower(strings.TrimSpace(redim.name))
			variable, ok := variables[name]
			if !ok || !variable.isArray {
				continue
			}
			value := state[name]
			if value.kind == arrayUnallocated && value.knownArray {
				if runtimeArrayUnallocatedUseIsInZeroCountLoop(file, proc, line, name) {
					continue
				}
				// The bound expression is evaluated before ReDim Preserve. If it
				// queries UBound on the same array, reaching Preserve already proves
				// that this particular bound query succeeded. The earlier UBound may
				// still be a deterministic failure on the entry state, but Preserve
				// itself cannot be the second failure on that same path.
				if runtimePreserveUsesSuccessfulBound(redim.dimensions, name) {
					continue
				}
				add("array_unallocated", arrayIndexOperationKey(name, "unallocated"))
			}
		}
	}
	lowerText := strings.ToLower(strings.TrimSpace(text))
	indexedStatement := !strings.HasPrefix(lowerText, "dim ") &&
		!strings.HasPrefix(lowerText, "static ") &&
		!strings.HasPrefix(lowerText, "private ") &&
		!strings.HasPrefix(lowerText, "public ") &&
		!strings.HasPrefix(lowerText, "friend ") &&
		!strings.HasPrefix(lowerText, "redim ") &&
		!strings.HasPrefix(lowerText, "erase ")
	if !indexedStatement {
		return issues
	}
	for _, use := range arrayIndexedUses(text, variables) {
		if len(use.args) == 0 {
			continue
		}
		name := strings.ToLower(use.name)
		variable, ok := variables[name]
		if !ok || !variable.isArray {
			continue
		}
		value := state[name]
		if value.kind == arrayUnallocated && value.knownArray {
			if runtimeArrayUnallocatedUseIsInZeroCountLoop(file, proc, line, name) {
				continue
			}
			if arrayBlockProvesIndexedUseAfterRedim(text, use.name, variables, constants, proc, line, base) {
				continue
			}
			add("array_unallocated", arrayIndexOperationKey(name, "unallocated"))
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			continue
		}
		if len(value.dimensions) > 0 && len(use.args) != len(value.dimensions) {
			add("array_subscript_out_of_bounds", arrayIndexOperationKey(name, "dimension"))
			continue
		}
		for index, argument := range use.args {
			if index >= len(value.dimensions) {
				break
			}
			result := constexpr.EvaluateIntegerEnvironment(argument, constexpr.IntegerValues(constants))
			if result.Kind != constexpr.Known {
				continue
			}
			bound := value.dimensions[index]
			if bound.lower.known && result.Value < bound.lower.value || bound.upper.known && result.Value > bound.upper.value {
				add("array_subscript_out_of_bounds", arrayIndexOperationKey(name, "bounds"))
			}
		}
	}
	return issues
}

func runtimeArrayLineHasFatalPreserve(text string, state arrayFlowState, variables map[string]arrayVariable, proc sourceProcedure, line int) bool {
	active := runtimeArrayBlockHasPriorErrorHandling(proc, line)
	if active {
		return false
	}
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 || !strings.EqualFold(strings.TrimSpace(match[1]), "preserve") {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct {
			continue
		}
		name := strings.ToLower(strings.TrimSpace(redim.name))
		variable, ok := variables[name]
		if !ok || !variable.isArray || variable.fixed {
			continue
		}
		value := state[name]
		if value.kind == arrayUnallocated && value.knownArray && !runtimePreserveUsesSuccessfulBound(redim.dimensions, name) {
			return true
		}
	}
	return false
}

func runtimeArrayLineHasFatalIndexedIssue(file parsedFile, issues []deterministicArrayRuntimeIssue, text string, proc sourceProcedure, line int) bool {
	if runtimeArrayBlockHasPriorErrorHandling(proc, line) || strings.Contains(strings.ToLower(text), "on error") {
		return false
	}
	if runtimeArrayLineFollowsMultiTargetErase(file, proc, line, issues) {
		// A single Erase can clear several independent dynamic arrays. Keep
		// visiting the following source lines so the runtime projection reports
		// each newly invalidated target, even though the first failing access
		// would terminate one concrete execution.
		return false
	}
	for _, issue := range issues {
		if strings.HasPrefix(issue.operationKey, "index:") {
			return true
		}
	}
	return false
}

func runtimeArrayLineFollowsMultiTargetErase(file parsedFile, proc sourceProcedure, line int, issues []deterministicArrayRuntimeIssue) bool {
	if line <= proc.StartLine || line <= 1 || line > len(file.Lines) {
		return false
	}
	match := arrayEraseRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(file.Lines[line-1-1])))
	if len(match) == 0 {
		return false
	}
	targets := make(map[string]bool)
	for _, target := range splitArgs(match[1]) {
		name := strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))
		if name != "" {
			targets[name] = true
		}
	}
	if len(targets) < 2 {
		return false
	}
	seenIssue := false
	for _, issue := range issues {
		if !strings.HasPrefix(issue.operationKey, "index:") {
			continue
		}
		parts := strings.Split(issue.operationKey, ":")
		if len(parts) < 2 || !targets[parts[1]] {
			return false
		}
		seenIssue = true
	}
	return seenIssue
}

// runtimeArrayLineHasFatalIssue reports a normal-flow terminal array failure.
// A later, successful plain ReDim of the same target is a narrow recovery
// boundary: the source-order projection must continue past the boundary so a
// second independent lifecycle failure (for example, after a subsequent
// Erase) is still observable. Without this exception the CFG stop edge would
// hide every later state transition in the procedure. Preserve and indexed
// operations remain terminal when another operation touches the target first.
func runtimeArrayLineHasFatalIssue(file parsedFile, text string, state arrayFlowState, issues []deterministicArrayRuntimeIssue, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, line int) bool {
	if len(issues) == 0 {
		return false
	}
	fatalPreserve := runtimeArrayLineHasFatalPreserve(text, state, variables, proc, line)
	fatalIndexed := runtimeArrayLineHasFatalIndexedIssue(file, issues, text, proc, line)
	if !fatalPreserve && !fatalIndexed {
		return false
	}
	return !runtimeArrayFailureHasImmediateReset(file, proc, line, text, issues, variables, constants)
}

func runtimeArrayFailureHasImmediateReset(file parsedFile, proc sourceProcedure, line int, text string, issues []deterministicArrayRuntimeIssue, variables map[string]arrayVariable, constants map[string]int) bool {
	targets := runtimeArrayFailureTargets(text, issues)
	if len(targets) == 0 || line >= len(file.Lines) {
		return false
	}
	base := arrayOptionBase(file)
	for target := range targets {
		if !runtimeArrayTargetHasNextReset(file, proc, line, target, variables, constants, base) {
			return false
		}
	}
	return true
}

func runtimeArrayFailureTargets(text string, issues []deterministicArrayRuntimeIssue) map[string]bool {
	targets := make(map[string]bool)
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct {
				name := strings.ToLower(cleanIdentifier(redim.name))
				if name != "" {
					targets[name] = true
				}
			}
		}
	}
	for _, issue := range issues {
		parts := strings.Split(issue.operationKey, ":")
		if len(parts) >= 2 && parts[0] == "index" {
			name := strings.ToLower(cleanIdentifier(parts[1]))
			if name != "" {
				targets[name] = true
			}
		}
	}
	return targets
}

func runtimeArrayTargetHasNextReset(file parsedFile, proc sourceProcedure, line int, target string, variables map[string]arrayVariable, constants map[string]int, base int) bool {
	target = strings.ToLower(cleanIdentifier(target))
	for candidateLine := line + 1; candidateLine <= proc.EndLine && candidateLine <= len(file.Lines); candidateLine++ {
		for _, source := range splitRangeValueSourceStatements(arraySourceOrderStripComment(file.Lines[candidateLine-1])) {
			source = strings.TrimSpace(source)
			if source == "" || runtimeArrayNonOperationLine(source) {
				continue
			}
			if runtimeArraySuccessfulPlainRedim(source, target, variables, constants, base) {
				return true
			}
			if runtimeArraySourceTouchesTarget(source, target, variables) {
				return false
			}
		}
	}
	return false
}

func runtimeArraySuccessfulPlainRedim(text, target string, variables map[string]arrayVariable, constants map[string]int, base int) bool {
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct || !strings.EqualFold(cleanIdentifier(redim.name), target) {
			continue
		}
		variable, known := variables[target]
		if !known || !variable.isArray || variable.fixed {
			return false
		}
		return !impossibleArrayBounds(parseArrayDimensionsWithConstants(redim.dimensions, base, constants))
	}
	return false
}

func runtimeArraySourceTouchesTarget(text, target string, variables map[string]arrayVariable) bool {
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), target) {
				return true
			}
		}
	}
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) > 0 {
		for _, name := range splitArgs(match[1]) {
			if strings.EqualFold(cleanIdentifier(name), target) {
				return true
			}
		}
	}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		if strings.EqualFold(cleanIdentifier(bound[2]), target) {
			return true
		}
	}
	for _, use := range arrayIndexedUses(text, variables) {
		if strings.EqualFold(cleanIdentifier(use.name), target) && len(use.args) > 0 {
			return true
		}
	}
	if lhs, _, indexed, assigned := arrayAssignment(text); assigned && strings.EqualFold(cleanIdentifier(lhs), target) {
		return true
	} else if assigned && indexed && strings.EqualFold(cleanIdentifier(arrayElementBaseName(lhs)), target) {
		return true
	}
	return false
}

func runtimeArrayNonOperationLine(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return lower == "else" || strings.HasPrefix(lower, "else ") || lower == "end if" || lower == "next" || strings.HasPrefix(lower, "next ") || lower == "loop" || strings.HasPrefix(lower, "loop ") || lower == "wend" || lower == "end with" || lower == "end select" || strings.HasPrefix(lower, "case ") || strings.HasPrefix(lower, "select case ") || strings.HasPrefix(lower, "if ") || strings.HasPrefix(lower, "elseif ") || strings.HasPrefix(lower, "do ") || lower == "do" || strings.HasPrefix(lower, "while ") || strings.HasPrefix(lower, "for ") || strings.HasPrefix(lower, "with ")
}

func runtimePreserveUsesSuccessfulBound(dimensions, target string) bool {
	target = strings.ToLower(cleanIdentifier(target))
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(dimensions, -1) {
		if !strings.EqualFold(bound[1], "ubound") || !strings.EqualFold(cleanIdentifier(strings.TrimSpace(bound[2])), target) {
			continue
		}
		return true
	}
	return false
}

type runtimeArrayBlockScope struct {
	kind   string
	id     int
	branch string
	line   int
}

type runtimeArrayBlockRedim struct {
	offset int
	path   []string
}

// runtimeArrayBlockProvesScopedRedim handles a multi-line conditional block
// that the procedure IR keeps as one CFG statement. A block-entry state is
// still the correct conservative state for the enclosing condition, but a
// local ReDim dominates a later use when both are in the same lexical branch.
// Only plain ReDim, source-order dominance, and direct array-call safety are
// accepted; branch joins, Erase, and unknown writes deliberately fall back to
// the entry-state warning.
func runtimeArrayBlockProvesScopedRedim(text, target string, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, startLine, base int) bool {
	lines := normalizedSourceLines(text)
	if len(lines) < 2 {
		return false
	}
	target = strings.ToLower(strings.TrimSpace(target))
	redims := map[string][]runtimeArrayBlockRedim{}
	scopes := make([]runtimeArrayBlockScope, 0, 8)
	nextScopeID := 0
	sawUse := false
	for offset, rawLine := range lines {
		line := normalizedCodeLine(rawLine)
		lower := strings.ToLower(strings.TrimSpace(line))
		if lower == "" {
			continue
		}
		if strings.HasPrefix(lower, "end if") {
			scopes = runtimeArrayBlockPopScope(scopes, "if")
			continue
		}
		if strings.HasPrefix(lower, "end select") {
			scopes = runtimeArrayBlockPopScope(scopes, "select")
			continue
		}
		if strings.HasPrefix(lower, "end with") {
			scopes = runtimeArrayBlockPopScope(scopes, "with")
			continue
		}
		if lower == "wend" {
			scopes = runtimeArrayBlockPopScope(scopes, "while")
			continue
		}
		if strings.HasPrefix(lower, "loop") {
			scopes = runtimeArrayBlockPopScope(scopes, "do")
			continue
		}
		if strings.HasPrefix(lower, "next") {
			scopes = runtimeArrayBlockPopScope(scopes, "for")
			continue
		}
		if strings.HasPrefix(lower, "elseif ") {
			runtimeArrayBlockSetBranch(&scopes, "if", "elseif:"+lower)
		} else if lower == "else" || strings.HasPrefix(lower, "else ") {
			runtimeArrayBlockSetBranch(&scopes, "if", "else")
		} else if strings.HasPrefix(lower, "case ") {
			runtimeArrayBlockSetBranch(&scopes, "select", "case:"+lower)
		}

		isRedim := arrayRedimRe.MatchString(line)
		if !isRedim {
			for _, use := range arrayIndexedUsesForSource(line, variables) {
				if len(use.args) == 0 || !strings.EqualFold(use.name, target) {
					continue
				}
				sawUse = true
				if !runtimeArrayBlockHasDominatingRedim(redims[target], offset, scopes, proc, startLine, target) {
					return false
				}
			}
		}

		if match := arrayRedimRe.FindStringSubmatch(line); len(match) > 0 {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct || !strings.EqualFold(redim.name, target) {
					continue
				}
				variable, known := variables[strings.ToLower(redim.name)]
				if !known || !variable.isArray || variable.fixed {
					continue
				}
				dimensions := parseArrayDimensionsWithConstants(redim.dimensions, base, constants)
				if impossibleArrayBounds(dimensions) {
					continue
				}
				plain := strings.TrimSpace(match[1]) == ""
				// A successful ReDim Preserve is an allocation boundary for
				// subsequent normal execution. If it fails, a procedure without
				// error handling cannot reach the later use; with Resume Next or a
				// recovery label the old unallocated state must remain possible.
				preserveNormalSuccess := !plain && !runtimeArrayBlockHasPriorErrorHandling(proc, startLine+offset)
				if plain || preserveNormalSuccess {
					redims[target] = append(redims[target], runtimeArrayBlockRedim{
						offset: offset,
						path:   runtimeArrayBlockScopePath(scopes),
					})
				}
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(line); len(match) > 0 {
			for _, erased := range splitArgs(match[1]) {
				if strings.EqualFold(strings.TrimSpace(erased), target) {
					redims[target] = nil
				}
			}
		}
		if lhs, _, indexed, assigned := arrayAssignment(line); assigned && !indexed && strings.EqualFold(lhs, target) {
			redims[target] = nil
		}

		if strings.HasPrefix(lower, "select case ") {
			nextScopeID++
			scopes = append(scopes, runtimeArrayBlockScope{kind: "select", id: nextScopeID, branch: "body"})
		} else if strings.HasPrefix(lower, "with ") {
			nextScopeID++
			scopes = append(scopes, runtimeArrayBlockScope{kind: "with", id: nextScopeID, branch: "body"})
		} else if lower == "do" || strings.HasPrefix(lower, "do while ") || strings.HasPrefix(lower, "do until ") {
			nextScopeID++
			scopes = append(scopes, runtimeArrayBlockScope{kind: "do", id: nextScopeID, branch: "body"})
		} else if strings.HasPrefix(lower, "while ") {
			nextScopeID++
			scopes = append(scopes, runtimeArrayBlockScope{kind: "while", id: nextScopeID, branch: "body"})
		} else if strings.HasPrefix(lower, "for ") {
			nextScopeID++
			scopes = append(scopes, runtimeArrayBlockScope{kind: "for", id: nextScopeID, branch: "body"})
		} else if strings.HasPrefix(lower, "if ") && runtimeArrayBlockIsBlockIf(lower) {
			nextScopeID++
			scopes = append(scopes, runtimeArrayBlockScope{kind: "if", id: nextScopeID, branch: "body"})
		}
	}
	return sawUse
}

func runtimeArrayBlockHasPriorErrorHandling(proc sourceProcedure, line int) bool {
	active := false
	if facts := proc.analysisFacts(); facts != nil && facts.runtimeErrorModeEventsBuilt {
		for _, event := range facts.runtimeErrorModeEvents {
			if event.line >= line {
				continue
			}
			switch {
			case event.mode == "resume next":
				active = true
			case event.mode == "goto 0":
				active = false
			case strings.HasPrefix(event.mode, "goto "):
				active = runtimeArrayErrorHandlerMayContinue(proc, event.line, strings.TrimSpace(strings.TrimPrefix(event.mode, "goto ")))
			default:
				active = true
			}
		}
		return active
	}
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine >= line {
			continue
		}
		for _, part := range splitRangeValueSourceStatements(statement.Text) {
			normalized := strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(part))), " ")
			if !strings.HasPrefix(normalized, "on error ") {
				continue
			}
			mode := strings.TrimSpace(strings.TrimPrefix(normalized, "on error "))
			switch {
			case mode == "resume next":
				// Resume Next keeps the source-order continuation after a failed
				// operation reachable, so retain the pre-operation array state.
				active = true
			case mode == "goto 0":
				active = false
			case strings.HasPrefix(mode, "goto "):
				// A label handler normally leaves the sequential path.  Keep it
				// conservative only when the handler can explicitly resume or jump
				// back into executable code.  An error handler that exits (or simply
				// reaches the procedure end) must not make statements after a fatal
				// Preserve appear reachable.
				active = runtimeArrayErrorHandlerMayContinue(proc, statement.Range.StartLine, strings.TrimSpace(strings.TrimPrefix(mode, "goto ")))
			default:
				// Unknown error-mode syntax remains fail-closed for the runtime
				// projection rather than assuming that a failed operation stops.
				active = true
			}
		}
	}
	return active
}

func runtimeArrayErrorHandlerMayContinue(proc sourceProcedure, directiveLine int, target string) bool {
	target = strings.ToLower(cleanIdentifier(target))
	if target == "" || target == "0" {
		return false
	}
	handlerLine := 0
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementLabel || statement.Range.StartLine <= directiveLine {
			continue
		}
		if strings.EqualFold(cleanIdentifier(statement.Label), target) {
			handlerLine = statement.Range.StartLine
			break
		}
	}
	if handlerLine == 0 {
		// An unresolved handler may be supplied by recovered source or an
		// incomplete procedure. Preserve the historical conservative result.
		return true
	}
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine < handlerLine {
			continue
		}
		if statement.Kind == procedureir.StatementLabel && statement.Range.StartLine != handlerLine {
			// A later label begins a distinct handler/control-flow region. The
			// current handler has no explicit continuation before it.
			break
		}
		if statement.Kind == procedureir.StatementResume || statement.Kind == procedureir.StatementGoTo {
			return true
		}
	}
	return false
}

func runtimeArrayProcedureNeedsCFGRefinement(proc sourceProcedure) bool {
	hasPreserve := false
	hasNestedCompound := false
	for statement := range proc.Statements.All() {
		match := arrayRedimRe.FindStringSubmatch(statement.Text)
		if len(match) > 0 && strings.EqualFold(strings.TrimSpace(match[1]), "preserve") {
			hasPreserve = true
		}
		switch statement.Kind {
		case procedureir.StatementCase, procedureir.StatementWith, procedureir.StatementDo:
			hasNestedCompound = true
		}
	}
	if !hasPreserve {
		return false
	}
	// A compound CFG block can own both the failing Preserve and statements
	// after it.  When the operation has no Resume/GoTo handler, a failed
	// Preserve terminates the normal path, so the source-ordered lane must be
	// enabled even without an explicit On Error directive.  Restrict this to
	// the compound kinds that the runtime lane expands below; ordinary
	// straight-line procedures keep the cheaper CFG path.
	if hasNestedCompound {
		return true
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementOnError || statement.Control == nil || statement.Control.Transfer != procedureir.TransferOnErrorGoto {
			continue
		}
		if !runtimeArrayErrorHandlerMayContinue(proc, statement.Range.StartLine, statement.Label) {
			return true
		}
	}
	return false
}

type runtimeScalarEnvironment struct {
	state     runtimeScalarState
	constants map[string]int
}

func (env runtimeScalarEnvironment) Resolve(name string) (constexpr.Value, bool) {
	key := strings.ToLower(strings.TrimSpace(name))
	if value, ok := env.state[key]; ok && value.known {
		return constexpr.Value{Kind: constexpr.ValueLongLong, Integer: int64(value.value)}, true
	}
	if value, ok := env.constants[key]; ok {
		return constexpr.Value{Kind: constexpr.ValueLongLong, Integer: int64(value)}, true
	}
	return constexpr.Value{}, false
}

func runtimeArrayCFGConditionValue(text string, state runtimeScalarState, constants map[string]int) (bool, bool) {
	condition, _, ok := arrayIfThenParts(text)
	if !ok {
		return false, false
	}
	lower := strings.ToLower(strings.TrimSpace(condition))
	if strings.HasPrefix(lower, "if ") || strings.HasPrefix(lower, "elseif ") {
		condition = strings.TrimSpace(condition[strings.IndexAny(lower, " ")+1:])
	}
	result := constexpr.Evaluate(condition, runtimeScalarEnvironment{state: state, constants: constants})
	if result.Kind != constexpr.Known || result.Typed.Kind != constexpr.ValueBoolean {
		return false, false
	}
	return result.Typed.Boolean, true
}

func runtimeArrayCFGInitialScalarState(proc sourceProcedure, variables map[string]arrayVariable) runtimeScalarState {
	state := runtimeScalarState{}
	for declaration := range proc.Declarations.All() {
		name := runtimeSimpleIdentifier(declaration.Name)
		if name == "" {
			continue
		}
		variable, ok := variables[name]
		if !ok || !variable.knownScalar || variable.isArray || variable.isVariant || variable.parameter || variable.static {
			continue
		}
		// VBA initializes local numeric variables to zero on procedure entry.
		// Parameters, static locals, and module variables remain unknown because
		// their value can come from outside this invocation.
		state[name] = runtimeScalarValue{known: true, value: 0}
	}
	return state
}

type runtimeArrayCFGFlow struct {
	arrays  arrayFlowState
	scalars runtimeScalarState
}

func (a Analyzer) runtimeArrayCFGTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, in runtimeArrayCFGFlow, text string, line int, constants map[string]int) runtimeArrayCFGFlow {
	refinementCtx := ctx
	refinementCtx.arrayStats = nil
	outArrays, _ := a.arrayTransfer(file, proc, refinementCtx, variables, in.arrays, text, line, constants, nil)
	forEachArrayCallAtLine(proc, line, func(call procedureir.CallSite) {
		outArrays = applyArrayModuleCallEffects(outArrays, file, proc, call, refinementCtx, variables, moduleDecls)
		outArrays = applyArrayUnknownModuleCallEffects(outArrays, file, proc, call, refinementCtx, variables, moduleDecls)
		if _, _, conditional := arrayIfThenParts(text); conditional {
			outArrays = applyArrayConditionalByRefCallEffects(outArrays, proc, call, refinementCtx)
		} else {
			outArrays = applyArrayByRefCallEffects(outArrays, proc, call, refinementCtx)
		}
		outArrays = restoreDeterministicArrayEraseState(outArrays, proc, call, refinementCtx)
	}, refinementCtx.arrayStats)
	return runtimeArrayCFGFlow{
		arrays:  outArrays,
		scalars: runtimeScalarTransfer(in.scalars, text, variables),
	}
}

func runtimeArrayCFGBlockText(file parsedFile, block vbacfg.Block) (string, int) {
	if block.Statement == nil {
		return "", block.Range.StartLine
	}
	line := block.Statement.Range.StartLine
	if line == 0 {
		line = block.Range.StartLine
	}
	text := block.Statement.Text
	if strings.TrimSpace(text) == "" && line >= 1 && line <= len(file.Lines) {
		text = normalizedCodeLine(file.Lines[line-1])
	}
	return text, line
}

func runtimeArrayCFGStates(a Analyzer, file parsedFile, proc sourceProcedure, graph *vbacfg.Graph, initial arrayFlowState, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, ctx analysisContext, constants map[string]int) map[vbacfg.BlockID]runtimeArrayCFGFlow {
	if graph == nil {
		return nil
	}
	blocks := make(map[vbacfg.BlockID]vbacfg.Block, len(graph.Blocks))
	outgoing := make(map[vbacfg.BlockID][]vbacfg.Edge, len(graph.Blocks))
	for _, block := range graph.Blocks {
		blocks[block.ID] = block
	}
	for _, edge := range graph.Edges {
		outgoing[edge.From] = append(outgoing[edge.From], edge)
	}
	states := map[vbacfg.BlockID]runtimeArrayCFGFlow{
		graph.Entry: {
			arrays:  cloneArrayState(initial),
			scalars: runtimeArrayCFGInitialScalarState(proc, variables),
		},
	}
	queued := map[vbacfg.BlockID]bool{graph.Entry: true}
	for len(queued) > 0 {
		var id vbacfg.BlockID
		first := true
		for candidate := range queued {
			if first || candidate < id {
				id = candidate
				first = false
			}
		}
		delete(queued, id)
		block, ok := blocks[id]
		if !ok {
			continue
		}
		in := states[id]
		// arrayTransfer updates its map in place. Keep the block input state
		// immutable so refinement observes the state before the statement (and
		// exceptional edges retain that same predecessor state).
		out := runtimeArrayCFGFlow{
			arrays:  cloneArrayState(in.arrays),
			scalars: cloneRuntimeScalarState(in.scalars),
		}
		text, line := runtimeArrayCFGBlockText(file, block)
		if block.Statement != nil {
			out = a.runtimeArrayCFGTransfer(file, proc, ctx, variables, moduleDecls, out, text, line, constants)
		}
		for _, edge := range outgoing[id] {
			next := out
			if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				next = in
			}
			current, exists := states[edge.To]
			if !exists {
				states[edge.To] = runtimeArrayCFGFlow{arrays: cloneArrayState(next.arrays), scalars: cloneRuntimeScalarState(next.scalars)}
				queued[edge.To] = true
				continue
			}
			merged := runtimeArrayCFGFlow{
				arrays:  meetArrayState(current.arrays, next.arrays),
				scalars: meetRuntimeScalarState(current.scalars, next.scalars),
			}
			if !arrayStateEqual(current.arrays, merged.arrays) || !runtimeScalarStateEqual(current.scalars, merged.scalars) {
				states[edge.To] = merged
				queued[edge.To] = true
			}
		}
	}
	return states
}

func runtimeScalarStateEqual(left, right runtimeScalarState) bool {
	if len(left) != len(right) {
		return false
	}
	for name, value := range left {
		if other, ok := right[name]; !ok || value != other {
			return false
		}
	}
	return true
}

func runtimeArrayCFGRefinedView(a Analyzer, file parsedFile, proc sourceProcedure, base vbacfg.CFGView, initial arrayFlowState, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, ctx analysisContext, constants map[string]int) vbacfg.CFGView {
	if !runtimeArrayProcedureNeedsCFGRefinement(proc) {
		return base
	}
	materialized := base.Materialize()
	originalEdges := append([]vbacfg.Edge(nil), materialized.Edges...)
	currentEdges := originalEdges
	for iteration := 0; iteration < 8; iteration++ {
		materialized.Edges = currentEdges
		states := runtimeArrayCFGStates(a, file, proc, &materialized, initial, variables, moduleDecls, ctx, constants)
		remove := map[vbacfg.EdgeID]bool{}
		for _, block := range materialized.Blocks {
			if block.Statement == nil {
				continue
			}
			flow, ok := states[block.ID]
			if !ok {
				continue
			}
			text, line := runtimeArrayCFGBlockText(file, block)
			if condition, known := runtimeArrayCFGConditionValue(text, flow.scalars, constants); known {
				for _, edge := range currentEdges {
					if edge.From != block.ID || edge.Class != vbacfg.EdgeNormal {
						continue
					}
					if edge.Kind == vbacfg.EdgeBranchTrue && !condition || edge.Kind == vbacfg.EdgeBranchFalse && condition {
						remove[edge.ID] = true
					}
				}
			}
			issues := a.deterministicArrayRuntimeIssues(file, text, line, flow.arrays, variables, constants, proc, arrayOptionBase(file), ctx, moduleDecls)
			if runtimeArrayLineHasFatalIssue(file, text, flow.arrays, issues, variables, constants, proc, line) {
				for _, edge := range currentEdges {
					if edge.From == block.ID && edge.Class == vbacfg.EdgeNormal {
						remove[edge.ID] = true
					}
				}
			}
		}
		// Edge refinement is monotonic: once a branch is proven impossible it
		// must not be reintroduced when a later state iteration has no new
		// removal for that predecessor.
		nextEdges := make([]vbacfg.Edge, 0, len(currentEdges)-len(remove))
		for _, edge := range currentEdges {
			if !remove[edge.ID] {
				nextEdges = append(nextEdges, edge)
			}
		}
		if runtimeArrayCFGEdgesEqual(currentEdges, nextEdges) {
			materialized.Edges = nextEdges
			return materialized.View(vbacfg.EdgeFilter{})
		}
		currentEdges = nextEdges
	}
	return base
}

func runtimeArrayCFGEdgesEqual(left, right []vbacfg.Edge) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		if left[index].ID != right[index].ID {
			return false
		}
	}
	return true
}

func runtimeArrayBlockHasDominatingRedim(redims []runtimeArrayBlockRedim, useOffset int, scopes []runtimeArrayBlockScope, proc sourceProcedure, startLine int, target string) bool {
	usePath := runtimeArrayBlockScopePath(scopes)
	for _, redim := range redims {
		if redim.offset >= useOffset || !runtimeArrayBlockPathPrefix(redim.path, usePath) {
			continue
		}
		for line := startLine + redim.offset + 1; line <= startLine+useOffset; line++ {
			if arrayBlockHasDirectArrayArgumentCall(proc, line, target) {
				return false
			}
		}
		return true
	}
	return false
}

func runtimeArrayBlockScopePath(scopes []runtimeArrayBlockScope) []string {
	path := make([]string, 0, len(scopes))
	for _, scope := range scopes {
		path = append(path, scope.kind+":"+strconv.Itoa(scope.id)+":"+scope.branch)
	}
	return path
}

func runtimeArrayBlockPathPrefix(prefix, path []string) bool {
	if len(prefix) > len(path) {
		return false
	}
	for index := range prefix {
		if prefix[index] != path[index] {
			return false
		}
	}
	return true
}

func runtimeArrayBlockPopScope(scopes []runtimeArrayBlockScope, kind string) []runtimeArrayBlockScope {
	for index := len(scopes) - 1; index >= 0; index-- {
		if scopes[index].kind == kind {
			return scopes[:index]
		}
	}
	return scopes
}

func runtimeArrayBlockSetBranch(scopes *[]runtimeArrayBlockScope, kind, branch string) {
	for index := len(*scopes) - 1; index >= 0; index-- {
		if (*scopes)[index].kind == kind {
			(*scopes)[index].branch = branch
			*scopes = (*scopes)[:index+1]
			return
		}
	}
}

func runtimeArrayBlockIsBlockIf(lower string) bool {
	thenIndex := strings.Index(lower, " then")
	if thenIndex < 0 {
		return false
	}
	after := strings.TrimSpace(lower[thenIndex+len(" then"):])
	return after == "" || after == "_"
}

type runtimeArraySourcePart struct {
	text string
	line int
	path []string
}

// runtimeArraySourcePartsBefore returns source-order statements together with
// their lexical branch path. It is intentionally source based because the
// runtime projection may receive several CFG blocks for one procedure and
// therefore cannot carry every scalar fact through its block-level state.
func runtimeArraySourcePartsBefore(file parsedFile, proc sourceProcedure, endLine int) []runtimeArraySourcePart {
	start := proc.StartLine
	if start < 1 {
		start = 1
	}
	end := endLine
	if end > len(file.Lines)+1 {
		end = len(file.Lines) + 1
	}
	if end <= start {
		return nil
	}
	parts := make([]runtimeArraySourcePart, 0)
	for line := start; line < end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" {
			continue
		}
		path := runtimeArrayBlockScopePath(runtimeArraySourceScopesBefore(file, proc, line))
		for _, rawPart := range splitRangeValueSourceStatements(text) {
			part := strings.TrimSpace(rawPart)
			if part == "" {
				continue
			}
			parts = append(parts, runtimeArraySourcePart{text: part, line: line, path: append([]string(nil), path...)})
		}
	}
	return parts
}

func runtimeArraySourceScopesBefore(file parsedFile, proc sourceProcedure, line int) []runtimeArrayBlockScope {
	start := proc.StartLine
	if start < 1 {
		start = 1
	}
	end := line
	if end > len(file.Lines)+1 {
		end = len(file.Lines) + 1
	}
	scopes := make([]runtimeArrayBlockScope, 0, 8)
	nextScopeID := 0
	for sourceLine := start; sourceLine < end; sourceLine++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[sourceLine-1]))
		lower := strings.ToLower(text)
		if lower == "" {
			continue
		}
		switch {
		case lower == "end if":
			scopes = runtimeArrayBlockPopScope(scopes, "if")
			continue
		case lower == "end select":
			scopes = runtimeArrayBlockPopScope(scopes, "select")
			continue
		case lower == "end with":
			scopes = runtimeArrayBlockPopScope(scopes, "with")
			continue
		case lower == "wend":
			scopes = runtimeArrayBlockPopScope(scopes, "while")
			continue
		case lower == "loop" || strings.HasPrefix(lower, "loop "):
			scopes = runtimeArrayBlockPopScope(scopes, "do")
			continue
		case lower == "next" || strings.HasPrefix(lower, "next "):
			scopes = runtimeArrayBlockPopScope(scopes, "for")
			continue
		}
		if strings.HasPrefix(lower, "elseif ") {
			runtimeArrayBlockSetBranch(&scopes, "if", "elseif:"+lower)
		} else if lower == "else" || strings.HasPrefix(lower, "else ") {
			runtimeArrayBlockSetBranch(&scopes, "if", "else")
		} else if strings.HasPrefix(lower, "case ") {
			runtimeArrayBlockSetBranch(&scopes, "select", "case:"+lower)
		}
		kind, starts := runtimeArrayBlockControlStart(lower)
		if !starts {
			continue
		}
		nextScopeID++
		scopes = append(scopes, runtimeArrayBlockScope{kind: kind, id: nextScopeID, branch: "body", line: sourceLine})
	}
	return scopes
}

func runtimeArrayUnallocatedUseIsInZeroCountLoop(file parsedFile, proc sourceProcedure, line int, target string) bool {
	if line <= proc.StartLine || line > len(file.Lines) {
		return false
	}
	target = strings.ToLower(cleanIdentifier(target))
	if target == "" {
		return false
	}
	scopes := runtimeArraySourceScopesBefore(file, proc, line)
	for _, scope := range scopes {
		if scope.kind != "for" || scope.line <= proc.StartLine || scope.line > len(file.Lines) {
			continue
		}
		header := strings.TrimSpace(normalizedCodeLine(file.Lines[scope.line-1]))
		if !runtimeArrayZeroCountLoopProven(file, proc, scope, header, target) {
			continue
		}
		return true
	}
	return false
}

func runtimeArrayZeroCountLoopProven(file parsedFile, proc sourceProcedure, loop runtimeArrayBlockScope, header, target string) bool {
	match := runtimeForPositiveScalarBoundRe.FindStringSubmatch(header)
	if len(match) != 2 {
		return false
	}
	upperName := runtimeSimpleIdentifier(match[1])
	if upperName == "" {
		return false
	}
	parts := runtimeArraySourcePartsBefore(file, proc, loop.line)
	upperIndex := -1
	countName := ""
	for index, part := range parts {
		lhs, rhs, indexed, assigned := arrayAssignment(part.text)
		if !assigned || indexed || runtimeSimpleIdentifier(lhs) != upperName {
			continue
		}
		if upperIndex >= 0 {
			return false
		}
		countName = runtimeSimpleIdentifier(rhs)
		if countName == "" {
			return false
		}
		upperIndex = index
	}
	if upperIndex < 0 || countName == "" {
		return false
	}

	initialized := false
	incremented := false
	for index, part := range parts {
		if index >= upperIndex {
			break
		}
		lhs, rhs, indexed, assigned := arrayAssignment(part.text)
		if !assigned || indexed || runtimeSimpleIdentifier(lhs) != countName {
			continue
		}
		if value, literal := integerLiteral(rhs); literal && value == 0 {
			if initialized {
				return false
			}
			initialized = true
			continue
		}
		arithmetic := runtimeScalarArithmeticRe.FindStringSubmatch(rhs)
		if !initialized || len(arithmetic) != 4 || runtimeSimpleIdentifier(arithmetic[1]) != countName || arithmetic[2] != "+" {
			return false
		}
		delta, err := strconv.Atoi(arithmetic[3])
		if err != nil || delta <= 0 || !runtimeArraySourceBranchHasFatalPreserve(parts, index, upperIndex, part.path, target, proc) {
			return false
		}
		incremented = true
	}
	if !initialized || !incremented {
		return false
	}
	for call := range proc.Calls.All() {
		if call.Range.StartLine < proc.StartLine || call.Range.StartLine >= loop.line {
			continue
		}
		if strings.EqualFold(cleanIdentifier(call.Callee.BaseName), target) {
			continue
		}
		for _, argument := range arrayCallArgumentTexts(proc, call) {
			if directArrayArgumentName(argument) == countName || directArrayArgumentName(argument) == target {
				return false
			}
		}
	}
	return true
}

func runtimeArraySourceBranchHasFatalPreserve(parts []runtimeArraySourcePart, assignmentIndex, upperIndex int, path []string, target string, proc sourceProcedure) bool {
	for index := 0; index < upperIndex; index++ {
		if index == assignmentIndex || !runtimeArrayPathsEqual(parts[index].path, path) {
			continue
		}
		match := arrayRedimRe.FindStringSubmatch(parts[index].text)
		if len(match) == 0 || !strings.EqualFold(strings.TrimSpace(match[1]), "preserve") || runtimeArrayBlockHasPriorErrorHandling(proc, parts[index].line) {
			continue
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), target) && !runtimePreserveUsesSuccessfulBound(redim.dimensions, target) {
				return true
			}
		}
	}
	return false
}

func runtimeArrayPathsEqual(left, right []string) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		if left[index] != right[index] {
			return false
		}
	}
	return true
}

// arrayBlockProvesIndexedUseAfterRedim handles the CFG representation of a
// Select Case clause. The procedure IR may expose the select statement or a
// Case body as one multi-line statement, so the runtime lane otherwise checks
// each indexed use against the block's entry state before arrayTransfer sees
// the earlier ReDim. Only straight-line Case bodies are accepted here; nested
// control flow remains on the conservative entry-state path.
func arrayBlockProvesIndexedUseAfterRedim(text string, target string, variables map[string]arrayVariable, constants map[string]int, proc sourceProcedure, startLine, base int) bool {
	if runtimeArrayBlockProvesScopedRedim(text, target, variables, constants, proc, startLine, base) {
		return true
	}
	lines := normalizedSourceLines(text)
	if len(lines) < 2 {
		return false
	}
	target = strings.ToLower(strings.TrimSpace(target))
	redimmed := false
	sawUse := false
	sawCaseLabel := false
	selectCaseCount := 0
	for offset, rawLine := range lines {
		line := normalizedCodeLine(rawLine)
		lower := strings.ToLower(strings.TrimSpace(line))
		if lower == "" {
			continue
		}
		caseLabel := strings.HasPrefix(lower, "case ")
		if strings.HasPrefix(lower, "select case ") {
			// A Select Case nested in a Case body needs its own path state. Do
			// not let the last inner Case allocation prove an outer use.
			if sawCaseLabel || selectCaseCount > 0 {
				return false
			}
			selectCaseCount++
		}
		if caseLabel {
			sawCaseLabel = true
			redimmed = false
		}
		if lower == "end select" {
			continue
		}
		if !caseLabel && runtimeArrayBlockHasNestedControlFlow(lower) {
			return false
		}
		if redimmed && arrayBlockHasDirectArrayArgumentCall(proc, startLine+offset, target) {
			return false
		}
		if match := arrayRedimRe.FindStringSubmatch(line); len(match) > 0 {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct {
					legacy := arrayRedimClauseRe.FindStringSubmatch(clause)
					if len(legacy) == 0 {
						continue
					}
					redim = directArrayRedimClause{name: legacy[1], dimensions: legacy[2]}
				}
				// ReDim Preserve can fail on an unallocated input, especially when
				// On Error Resume Next keeps execution in the branch. Only a plain
				// ReDim is a deterministic allocation boundary here.
				if strings.TrimSpace(match[1]) == "" &&
					!impossibleArrayBounds(parseArrayDimensionsWithConstants(redim.dimensions, base, constants)) &&
					strings.EqualFold(redim.name, target) {
					redimmed = true
				}
			}
			continue
		}
		for _, use := range arrayIndexedUses(line, variables) {
			if !strings.EqualFold(use.name, target) || len(use.args) == 0 {
				continue
			}
			sawUse = true
			if !redimmed {
				return false
			}
		}
	}
	return sawCaseLabel && sawUse
}

func arrayBlockHasDirectArrayArgumentCall(proc sourceProcedure, line int, target string) bool {
	target = strings.ToLower(strings.TrimSpace(target))
	for call := range proc.Calls.All() {
		if call.Range.StartLine != line {
			continue
		}
		for _, argument := range arrayCallArgumentTexts(proc, call) {
			if strings.EqualFold(directArrayArgumentName(argument), target) {
				return true
			}
		}
	}
	return false
}

func runtimeArrayBlockHasNestedControlFlow(line string) bool {
	for _, prefix := range []string{
		"if ", "elseif ", "else", "end if", "for ", "next ", "do", "loop", "while ", "wend", "with ", "end with", "on error", "goto ", "exit ", "erase ",
	} {
		if strings.HasPrefix(line, prefix) {
			return true
		}
	}
	return false
}

func suppressDeterministicArrayWarningDuplicates(findings []Finding) []Finding {
	if len(findings) == 0 {
		return findings
	}
	runtimeOperations := map[string]bool{}
	for _, finding := range findings {
		if finding.Code != "VBA249" || finding.RuntimeError == nil {
			continue
		}
		if (finding.RuntimeError.Kind == "array_unallocated" || finding.RuntimeError.Kind == "array_subscript_out_of_bounds") && finding.arrayOperationKey != "" {
			runtimeOperations[arrayOperationFindingKey(finding)] = true
		}
	}
	if len(runtimeOperations) == 0 {
		return findings
	}
	out := make([]Finding, 0, len(findings))
	for _, finding := range findings {
		if finding.Code == "VBA227" && finding.arrayLifecycleFinding && finding.arrayOperationKey != "" && runtimeOperations[arrayOperationFindingKey(finding)] {
			continue
		}
		out = append(out, finding)
	}
	return out
}

func arrayOperationFindingKey(finding Finding) string {
	return finding.File + ":" + strconv.Itoa(finding.Line) + ":" + finding.arrayOperationKey
}

func runtimeConstantEnvironment(base constexpr.Environment, state runtimeConstantState) constexpr.Environment {
	// Keep the immutable file/project environment shared and layer the
	// procedure-local transfer state over it. The previous implementation
	// copied every base constant for every CFG statement, which made a large
	// module's runtime-error pass allocate in proportion to
	// statements*constants. runtimeConstantState keys are normalized by the
	// producer; Resolve still normalizes callers to preserve constexpr.Values'
	// case-insensitive contract.
	return runtimeConstantOverlay{base: base, state: state}
}

// runtimeConstantOverlay is an immutable, read-only constexpr environment for
// one transfer state. The state map is never mutated after it is passed here;
// transfers clone before writing, so concurrent readers can safely share the
// base environment and the overlay value.
type runtimeConstantOverlay struct {
	base  constexpr.Environment
	state runtimeConstantState
}

type runtimeConstantScope struct {
	base   constexpr.Environment
	hidden map[string]bool
}

func (scope runtimeConstantScope) Resolve(name string) (constexpr.Value, bool) {
	key := strings.ToLower(strings.TrimSpace(name))
	if scope.hidden[key] || scope.base == nil {
		return constexpr.Value{}, false
	}
	return scope.base.Resolve(key)
}

func (overlay runtimeConstantOverlay) Resolve(name string) (constexpr.Value, bool) {
	key := strings.ToLower(strings.TrimSpace(name))
	if value, ok := overlay.state[key]; ok {
		return value, true
	}
	if overlay.base == nil {
		return constexpr.Value{}, false
	}
	return overlay.base.Resolve(key)
}

func runtimeCFGStates(graph *vbacfg.Graph, initial runtimeConstantState, base constexpr.Environment, accesses readOnlySpan[procedureir.VariableAccess]) map[vbacfg.BlockID]runtimeConstantState {
	if graph == nil {
		return nil
	}
	blocks := make(map[vbacfg.BlockID]vbacfg.Block, len(graph.Blocks))
	outgoing := make(map[vbacfg.BlockID][]vbacfg.Edge)
	for _, block := range graph.Blocks {
		blocks[block.ID] = block
	}
	for _, edge := range graph.Edges {
		outgoing[edge.From] = append(outgoing[edge.From], edge)
	}
	states := map[vbacfg.BlockID]runtimeConstantState{graph.Entry: cloneRuntimeState(initial)}
	// A recovered/unknown-flow source means that any statement may be reached
	// without the physically reconstructed predecessor.  Keep literal and
	// project-constant proofs (which do not depend on local state), but clear
	// local propagated values on those paths.  This mirrors cfg's conservative
	// definite-assignment query and prevents a stale zero from becoming a
	// universal proof after parser recovery.
	unknownFlow := false
	reachable := graph.Reachable(vbacfg.EdgeFilter{})
	reachableSet := make(map[vbacfg.BlockID]bool, len(reachable))
	for _, id := range reachable {
		reachableSet[id] = true
	}
	for _, source := range graph.UnknownFlowSources {
		if reachableSet[source] {
			unknownFlow = true
			break
		}
	}
	writes := runtimeWriteNames(accesses)
	queued := map[vbacfg.BlockID]bool{graph.Entry: true}
	for len(queued) > 0 {
		var id vbacfg.BlockID
		first := true
		for candidate := range queued {
			if first || candidate < id {
				id, first = candidate, false
			}
		}
		delete(queued, id)
		block, ok := blocks[id]
		if !ok {
			continue
		}
		in := cloneRuntimeState(states[id])
		out := in
		if block.Statement != nil {
			env := runtimeConstantEnvironment(base, in)
			out = runtimeTransfer(*block.Statement, in, env, writes[block.Statement.ID])
		}
		for _, edge := range outgoing[id] {
			if block.Statement != nil {
				env := runtimeConstantEnvironment(base, in)
				if !runtimeEdgeAllowed(*block.Statement, edge, env) {
					continue
				}
			}
			next := out
			if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				next = in
			}
			if unknownFlow && edge.To != graph.Entry {
				// Do not mutate the predecessor state: only the merged target
				// receives the conservative unknown-flow input.
				next = runtimeConstantState{}
			}
			if mergeRuntimeState(states, edge.To, next) {
				queued[edge.To] = true
			}
		}
	}
	return states
}

func mergeRuntimeState(states map[vbacfg.BlockID]runtimeConstantState, id vbacfg.BlockID, incoming runtimeConstantState) bool {
	current, exists := states[id]
	if !exists {
		states[id] = cloneRuntimeState(incoming)
		return true
	}
	merged := meetRuntimeState(current, incoming)
	if runtimeStateEqual(current, merged) {
		return false
	}
	states[id] = merged
	return true
}

func meetRuntimeState(left, right runtimeConstantState) runtimeConstantState {
	out := runtimeConstantState{}
	for name, value := range left {
		other, ok := right[name]
		if ok && other == value {
			out[name] = value
		}
	}
	return out
}

func cloneRuntimeState(state runtimeConstantState) runtimeConstantState {
	out := runtimeConstantState{}
	for name, value := range state {
		out[name] = value
	}
	return out
}

func runtimeStateEqual(left, right runtimeConstantState) bool {
	if len(left) != len(right) {
		return false
	}
	for name, value := range left {
		if other, ok := right[name]; !ok || other != value {
			return false
		}
	}
	return true
}

func runtimeTransfer(statement procedureir.Statement, in runtimeConstantState, env constexpr.Environment, writes []string) runtimeConstantState {
	out := cloneRuntimeState(in)
	// A procedure call can mutate any ByRef argument (and an unresolved or
	// external call can mutate state that the source-only resolver cannot see).
	// The caller-side IR does not always mark a ByRef argument as a write, so a
	// call is an explicit invalidation barrier. The expression at the call
	// site is evaluated against the incoming state before this transfer, which
	// preserves diagnostics such as `Debug.Print 10 / denominator`.
	if statement.Kind == procedureir.StatementCall || statement.Kind == procedureir.StatementRaiseEvent {
		out = runtimeConstantState{}
	}
	for _, name := range writes {
		delete(out, name)
	}
	if name, expression, ok := runtimeConstantAssignment(statement); ok {
		result := constexpr.Evaluate(expression, env)
		if result.Kind == constexpr.Known {
			out[name] = result.Typed
		} else {
			delete(out, name)
		}
	}
	return out
}

func runtimeWriteNames(accesses readOnlySpan[procedureir.VariableAccess]) map[int][]string {
	byStatement := map[int][]string{}
	seen := map[int]map[string]bool{}
	for access := range accesses.All() {
		if access.StatementID == 0 || (access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) {
			continue
		}
		name := runtimeSimpleIdentifier(access.Name)
		if name == "" {
			continue
		}
		if seen[access.StatementID] == nil {
			seen[access.StatementID] = map[string]bool{}
		}
		if seen[access.StatementID][name] {
			continue
		}
		seen[access.StatementID][name] = true
		byStatement[access.StatementID] = append(byStatement[access.StatementID], name)
	}
	for id := range byStatement {
		sort.Strings(byStatement[id])
	}
	return byStatement
}

func runtimeLocalConstantState(file parsedFile, proc sourceProcedure, base constexpr.Environment) runtimeConstantState {
	state := runtimeConstantState{}
	if proc.StartLine <= 0 || proc.EndLine <= 0 || len(file.Lines) == 0 {
		return state
	}
	start, end := proc.StartLine, proc.EndLine
	if start < 1 {
		start = 1
	}
	if end > len(file.Lines) {
		end = len(file.Lines)
	}
	for line := start; line <= end; line++ {
		match := runtimeConstAssignmentRe.FindStringSubmatch(normalizedCodeLine(file.Lines[line-1]))
		if len(match) == 0 {
			continue
		}
		name := runtimeSimpleIdentifier(match[1])
		if name == "" {
			continue
		}
		env := runtimeConstantEnvironment(base, state)
		if result := constexpr.Evaluate(strings.TrimSpace(match[2]), env); result.Kind == constexpr.Known {
			state[name] = result.Typed
		}
	}
	return state
}

func runtimeEdgeAllowed(statement procedureir.Statement, edge vbacfg.Edge, env constexpr.Environment) bool {
	if edge.Uncertain || statement.Condition == nil {
		return true
	}
	if edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return true
	}
	result := constexpr.Evaluate(statement.Condition.Text, env)
	if result.Kind != constexpr.Known || result.Typed.Kind != constexpr.ValueBoolean {
		return true
	}
	condition := result.Typed.Boolean
	if edge.Kind == vbacfg.EdgeBranchTrue {
		return condition
	}
	return !condition
}

func runtimeConstantAssignment(statement procedureir.Statement) (name, expression string, ok bool) {
	if statement.Kind == procedureir.StatementAssignment || statement.Kind == procedureir.StatementSet {
		if statement.Target != nil {
			name = runtimeSimpleIdentifier(statement.Target.Text)
		}
		if name == "" {
			if match := assignRe.FindStringSubmatch(statement.Text); len(match) > 0 {
				name = runtimeSimpleIdentifier(match[1])
			}
		}
		if name != "" {
			if index := strings.Index(statement.Text, "="); index >= 0 {
				expression = strings.TrimSpace(statement.Text[index+1:])
			}
			if expression != "" {
				return name, expression, true
			}
		}
	}
	if match := runtimeConstAssignmentRe.FindStringSubmatch(statement.Text); len(match) > 0 {
		return runtimeSimpleIdentifier(match[1]), strings.TrimSpace(match[2]), true
	}
	return "", "", false
}

func runtimeAssignmentTarget(statement procedureir.Statement) (string, bool) {
	if statement.Target != nil {
		if name := runtimeSimpleIdentifier(statement.Target.Text); name != "" {
			return name, true
		}
	}
	if match := assignRe.FindStringSubmatch(statement.Text); len(match) > 0 {
		if name := runtimeSimpleIdentifier(match[1]); name != "" {
			return name, true
		}
	}
	if match := runtimeConstAssignmentRe.FindStringSubmatch(statement.Text); len(match) > 0 {
		if name := runtimeSimpleIdentifier(match[1]); name != "" {
			return name, true
		}
	}
	return "", false
}

func runtimeSimpleIdentifier(text string) string {
	text = strings.TrimSpace(text)
	if text == "" || !isIdentifierStart(text[0]) {
		return ""
	}
	for i := 1; i < len(text); i++ {
		if !isIdentifierPart(text[i]) {
			return ""
		}
	}
	return strings.ToLower(text)
}

func appendRuntimeStatementFindings(findings []Finding, seen map[string]bool, analyzer Analyzer, file parsedFile, proc sourceProcedure, statement procedureir.Statement, facts *procedureAnalysisFacts, env constexpr.Environment) []Finding {
	add := func(expression procedureir.Expression, kind string) {
		line := expression.Range.StartLine
		if line <= 0 {
			line = statement.Range.StartLine
		}
		if line <= 0 {
			line = proc.StartLine
		}
		key := strconv.Itoa(line) + ":" + kind + ":" + expression.Text
		if seen[key] {
			return
		}
		seen[key] = true
		message, reason, suggestion := deterministicRuntimeFailureText(kind)
		finding := analyzer.simpleFinding(file, proc, line, "VBA249", "error", message, reason, suggestion)
		finding.RuntimeError = &RuntimeErrorContext{Kind: kind}
		findings = append(findings, finding)
	}
	for _, expression := range expressionsForRuntimeStatement(statement, facts) {
		if expression.Recovered || expression.Text == "" {
			continue
		}
		// A conversion can be nested inside a larger expression, while the
		// shared evaluator intentionally leaves the enclosing call unknown.
		// Walk the expression tree so supported conversions are checked at the
		// exact call site instead of only when they are the root expression.
		for _, candidate := range runtimeExpressionTree(expression, facts) {
			if candidate.Recovered || candidate.Text == "" {
				continue
			}
			if kind, ok := deterministicRuntimeConversionFailure(candidate.Text, env); ok {
				add(candidate, kind)
			}
		}
		if result := constexpr.Evaluate(expression.Text, env); result.Kind == constexpr.Invalid {
			if kind, ok := deterministicRuntimeFailureKind(result, expression, facts, env); ok {
				add(expression, kind)
			}
		}
	}
	return findings
}

func runtimeExpressionTree(root procedureir.Expression, facts *procedureAnalysisFacts) []procedureir.Expression {
	out := make([]procedureir.Expression, 0, 1+len(root.Children))
	visited := make(map[int]bool)
	var visit func(procedureir.Expression)
	visit = func(expression procedureir.Expression) {
		if visited[expression.ID] {
			return
		}
		visited[expression.ID] = true
		out = append(out, expression)
		for _, childID := range expression.Children {
			if child, ok := facts.Expression(childID); ok {
				visit(child)
			}
		}
	}
	visit(root)
	return out
}

// deterministicRuntimeConversionFailure models only the conversion domains
// whose target ranges are fixed by VBA and whose source value is already known
// to the shared constexpr evaluator. Locale-dependent numeric strings and
// unknown/Variant/call results remain silent.
func deterministicRuntimeConversionFailure(expression string, env constexpr.Environment) (string, bool) {
	match := runtimeConversionCallRe.FindStringSubmatch(expression)
	if len(match) == 0 {
		return "", false
	}
	trimmed := strings.TrimSpace(expression)
	open := strings.Index(trimmed, "(")
	if open < 0 || !strings.HasSuffix(trimmed, ")") {
		return "", false
	}
	close := len(trimmed) - 1
	if close <= open {
		return "", false
	}
	argument := strings.TrimSpace(trimmed[open+1 : close])
	if argument == "" || hasTopLevelComma(argument) {
		return "", false
	}
	result := constexpr.Evaluate(argument, env)
	if result.Kind != constexpr.Known {
		return "", false
	}
	if result.Typed.Kind == constexpr.ValueString {
		if definitelyNonnumericString(result.Typed.String) {
			return "conversion_type_mismatch", true
		}
		// Numeric string conversion is locale-sensitive in VBA. Keep this
		// conservative even when ParseFloat accepts the spelling.
		return "", false
	}
	if !runtimeNumericValue(result.Typed) {
		return "", false
	}
	value := runtimeNumericFloat(result.Typed)
	name := strings.ToLower(match[1])
	switch name {
	case "cbyte":
		value = math.RoundToEven(value)
		if value < 0 || value > 255 {
			return "conversion_overflow", true
		}
	case "cint":
		value = math.RoundToEven(value)
		if value < -32768 || value > 32767 {
			return "conversion_overflow", true
		}
	case "clng":
		value = math.RoundToEven(value)
		if value < -2147483648 || value > 2147483647 {
			return "conversion_overflow", true
		}
	case "clnglng":
		value = math.RoundToEven(value)
		if value < float64(math.MinInt64) || value > float64(math.MaxInt64) {
			return "conversion_overflow", true
		}
	case "csng":
		if value < -math.MaxFloat32 || value > math.MaxFloat32 {
			return "conversion_overflow", true
		}
	case "ccur":
		if value < -922337203685477.5807 || value > 922337203685477.5807 {
			return "conversion_overflow", true
		}
	case "cdec":
		if value < -79228162514264337593543950335.0 || value > 79228162514264337593543950335.0 {
			return "conversion_overflow", true
		}
	}
	return "", false
}

func hasTopLevelComma(text string) bool {
	depth := 0
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
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case ',':
			if !inString && depth == 0 {
				return true
			}
		}
	}
	return false
}

func runtimeNumericValue(value constexpr.Value) bool {
	switch value.Kind {
	case constexpr.ValueInteger, constexpr.ValueLong, constexpr.ValueLongLong, constexpr.ValueSingle, constexpr.ValueDouble, constexpr.ValueCurrency:
		return true
	default:
		return false
	}
}

func runtimeNumericFloat(value constexpr.Value) float64 {
	switch value.Kind {
	case constexpr.ValueInteger, constexpr.ValueLong, constexpr.ValueLongLong:
		return float64(value.Integer)
	case constexpr.ValueCurrency:
		return float64(value.Currency) / 10000
	default:
		return value.Float
	}
}

func expressionsForRuntimeStatement(statement procedureir.Statement, facts *procedureAnalysisFacts) []procedureir.Expression {
	ids := make([]int, 0, len(statement.ExpressionIDs))
	for _, id := range statement.ExpressionIDs {
		if expression, ok := facts.Expression(id); ok && expression.ParentID == 0 {
			ids = append(ids, id)
		}
	}
	if len(ids) == 0 {
		facts.forEachExpression(func(expression procedureir.Expression) {
			if expression.StatementID == statement.ID && expression.ParentID == 0 {
				ids = append(ids, expression.ID)
			}
		})
	}
	sort.Ints(ids)
	out := make([]procedureir.Expression, 0, len(ids))
	for _, id := range ids {
		if expression, ok := facts.Expression(id); ok {
			out = append(out, expression)
		}
	}
	return out
}

func deterministicRuntimeFailureKind(result constexpr.Result, expression procedureir.Expression, facts *procedureAnalysisFacts, env constexpr.Environment) (string, bool) {
	if result.Kind != constexpr.Invalid {
		return "", false
	}
	if result.Reason == "division by zero" {
		return "division_by_zero", true
	}
	// The evaluator rejects non-numeric arithmetic operands. Turn that into a
	// runtime diagnostic only when a child is a known string that cannot be
	// interpreted as a number. Numeric strings are accepted by VBA coercion.
	switch result.Reason {
	case "arithmetic requires numeric operands", "unary plus requires numeric operand", "unary minus requires numeric operand", "exponent requires numeric operands", "integer operator requires integral operands":
		if expressionHasKnownNonnumericString(expression, facts, env) {
			return "numeric_type_mismatch", true
		}
	}
	return "", false
}

func expressionHasKnownNonnumericString(expression procedureir.Expression, facts *procedureAnalysisFacts, env constexpr.Environment) bool {
	for _, childID := range expression.Children {
		child, ok := facts.Expression(childID)
		if !ok || child.Recovered {
			continue
		}
		if result := constexpr.Evaluate(child.Text, env); result.Kind == constexpr.Known && result.Typed.Kind == constexpr.ValueString && definitelyNonnumericString(result.Typed.String) {
			return true
		}
		if expressionHasKnownNonnumericString(child, facts, env) {
			return true
		}
	}
	return false
}

func numericString(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" {
		return false
	}
	parsed, err := strconv.ParseFloat(value, 64)
	return err == nil && !math.IsNaN(parsed) && !math.IsInf(parsed, 0)
}

func definitelyNonnumericString(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" || numericString(value) {
		return false
	}
	// Decimal separators, currency markers, and other punctuation can be
	// accepted by VBA's locale-sensitive coercion rules. Only alphabetic
	// content makes the failure independent of the host locale.
	for _, r := range value {
		if (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z') {
			return true
		}
	}
	return false
}

func deterministicRuntimeFailureText(kind string) (message, reason, suggestion string) {
	switch kind {
	case "division_by_zero":
		return "This expression is guaranteed to fail at runtime because it divides by zero.", "The divisor is a statically known zero value on this expression.", "Ensure the divisor is nonzero before evaluating the expression."
	case "numeric_type_mismatch":
		return "This expression is guaranteed to fail at runtime because numeric arithmetic uses a known nonnumeric string.", "The shared constant evaluator proves that a string operand cannot be coerced to a number for this operation.", "Validate or convert the string to a numeric value before using it in arithmetic."
	case "conversion_type_mismatch":
		return "This conversion is guaranteed to fail at runtime because its known value has an incompatible type.", "The conversion input is a known nonnumeric string and cannot be coerced by the supported VBA conversion.", "Validate the input before converting it, or use a compatible source type."
	case "conversion_overflow":
		return "This conversion is guaranteed to fail at runtime because its known value is outside the target type range.", "The shared constant evaluator proves that the value cannot be represented by the requested VBA numeric type.", "Use a wider target type or constrain the value before converting it."
	case "array_unallocated":
		return "This array access is guaranteed to fail at runtime because the dynamic array is unallocated.", "The shared array allocation facts prove that the array has not been allocated on every path reaching this access.", "Allocate the array with ReDim before accessing it, or guard the access with a proven allocation check."
	case "array_subscript_out_of_bounds":
		return "This array access is guaranteed to fail at runtime because its subscript is outside the known bounds.", "The shared array shape facts prove that the subscript or dimension cannot exist in the array at this point.", "Use a subscript within the established bounds or derive it from LBound/UBound."
	default:
		return "This expression is guaranteed to fail at runtime.", "Constant evaluation proves that the expression has an invalid runtime value.", "Replace the invalid value or guard the operation before evaluating it."
	}
}
