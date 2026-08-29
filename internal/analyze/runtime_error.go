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
	state := arrayInitialState(variables)
	findings := make([]Finding, 0)
	seen := map[string]bool{}
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		for _, issue := range deterministicArrayRuntimeIssues(text, line, in, variables, constants) {
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
		// Use the existing transfer, with the legacy warning gate forced on so
		// that this projection remains available even when users disable the
		// lower-confidence VBA227 warning rule.
		probe := a
		probe.Config.Analyze.DetectArrayLifecycleSafety = true
		out, _ := probe.arrayTransfer(file, proc, ctx, variables, in, text, line, constants, nil)
		return out
	}
	if proc.Graph != nil {
		view := proc.Graph.View(vbacfg.EdgeFilter{})
		walkArrayCFG(&view, file.Lines, state, visit)
	} else {
		for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
			state = visit(normalizedCodeLine(file.Lines[line-1]), line, state)
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

func arrayBoundOperationKey(kind, name, failure string) string {
	return "bound:" + strings.ToLower(strings.TrimSpace(kind)) + ":" + strings.ToLower(strings.TrimSpace(name)) + ":" + strings.ToLower(strings.TrimSpace(failure))
}

func arrayIndexOperationKey(name, failure string) string {
	return "index:" + strings.ToLower(strings.TrimSpace(name)) + ":" + strings.ToLower(strings.TrimSpace(failure))
}

func deterministicArrayRuntimeIssues(text string, line int, state arrayFlowState, variables map[string]arrayVariable, constants map[string]int) []deterministicArrayRuntimeIssue {
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
			result := constexpr.EvaluateInteger(argument, constants)
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
