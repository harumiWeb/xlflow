package dataflow

import (
	"container/heap"
	"fmt"
	"regexp"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// AnalyzeProcedure performs a conservative, intra-procedural forward
// may-analysis. A supplied graph is used as-is; when it is empty a graph is
// built from the procedure so small protocol adapters can use the API safely.
func AnalyzeProcedure(procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) Result {
	if len(graph.Blocks) == 0 {
		graph = cfg.Build(procedure)
	}
	a := newProcedureAnalyzer(procedure, graph, options)
	return a.run()
}

type provenance struct {
	source Source
	state  State
	path   []PathStep
	safe   map[SinkKind]bool
}

type value struct {
	origins map[string]provenance
}

type abstractState struct {
	vars map[string]value
}

type procedureAnalyzer struct {
	procedure         procedureir.ProcedureIR
	graph             cfg.Graph
	options           Options
	statements        map[int]procedureir.Statement
	expressions       map[int]procedureir.Expression
	callsByExpression map[int]procedureir.CallSite
	callsByStatement  map[int][]procedureir.CallSite
	declaredNames     map[string]bool
	knownShellObjects map[string]bool
	selectVariables   map[int]string
	findings          map[string]Finding
	blocksByID        map[cfg.BlockID]cfg.Block
	outgoingEdges     map[cfg.BlockID][]cfg.Edge
}

type analysisStats struct {
	transfers int
}

func newProcedureAnalyzer(procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) *procedureAnalyzer {
	a := &procedureAnalyzer{
		procedure:         procedure,
		graph:             graph,
		options:           options,
		statements:        map[int]procedureir.Statement{},
		expressions:       map[int]procedureir.Expression{},
		callsByExpression: map[int]procedureir.CallSite{},
		callsByStatement:  map[int][]procedureir.CallSite{},
		declaredNames:     map[string]bool{},
		knownShellObjects: map[string]bool{},
		selectVariables:   map[int]string{},
		findings:          map[string]Finding{},
		blocksByID:        map[cfg.BlockID]cfg.Block{},
		outgoingEdges:     map[cfg.BlockID][]cfg.Edge{},
	}
	for _, block := range graph.Blocks {
		a.blocksByID[block.ID] = block
	}
	for _, edge := range graph.Edges {
		a.outgoingEdges[edge.From] = append(a.outgoingEdges[edge.From], edge)
	}
	for from := range a.outgoingEdges {
		sort.SliceStable(a.outgoingEdges[from], func(i, j int) bool {
			left, right := a.outgoingEdges[from][i], a.outgoingEdges[from][j]
			if left.To != right.To {
				return left.To < right.To
			}
			return left.ID < right.ID
		})
	}
	for _, statement := range procedure.Statements {
		a.statements[statement.ID] = statement
	}
	for _, declaration := range procedure.Declarations {
		a.declaredNames[canonicalName(declaration.Name)] = true
	}
	for _, parameter := range procedure.Symbol.Parameters {
		a.declaredNames[canonicalName(parameter.Name)] = true
	}
	for _, expression := range procedure.Expressions {
		a.expressions[expression.ID] = expression
	}
	for _, call := range procedure.Calls {
		if call.ExpressionID != 0 {
			a.callsByExpression[call.ExpressionID] = call
		}
		a.callsByStatement[call.StatementID] = append(a.callsByStatement[call.StatementID], call)
	}
	for _, statement := range procedure.Statements {
		if statement.Kind != procedureir.StatementSelect {
			continue
		}
		if variable := selectorVariable(statement.Text); variable != "" {
			a.selectVariables[statement.ID] = variable
		}
	}
	a.scanKnownShellObjects()
	return a
}

func (a *procedureAnalyzer) run() Result {
	result, _ := a.runWithStats()
	return result
}

func (a *procedureAnalyzer) runWithStats() (Result, analysisStats) {
	return a.runWithStatsAndRank(nil)
}

func (a *procedureAnalyzer) runWithStatsAndRank(worklistRank map[cfg.BlockID]int) (Result, analysisStats) {
	stats := analysisStats{}
	entry := abstractState{vars: map[string]value{}}
	for _, parameter := range a.procedure.Symbol.Parameters {
		entry.vars[canonicalName(parameter.Name)] = valueFromSource(Source{
			Kind:        SourceParameter,
			Label:       parameter.Name,
			Range:       parameter.Range,
			StatementID: 0,
		}, PathStep{Kind: "source", Label: parameter.Name, Range: parameter.Range})
	}

	reachable := map[cfg.BlockID]bool{}
	for _, id := range a.graph.Reachable(cfg.EdgeFilter{}) {
		reachable[id] = true
	}
	if len(reachable) == 0 {
		return Result{Findings: nil, States: nil}, stats
	}
	inStates := map[cfg.BlockID]abstractState{a.graph.Entry: entry}
	if worklistRank == nil {
		worklistRank = a.reversePostOrderRank(reachable)
	}
	queue := rankedWorklist{{id: a.graph.Entry, rank: worklistRank[a.graph.Entry]}}
	heap.Init(&queue)
	inQueue := map[cfg.BlockID]bool{a.graph.Entry: true}
	for len(queue) > 0 {
		id := heap.Pop(&queue).(rankedBlock).id
		inQueue[id] = false
		state, ok := inStates[id]
		if !ok {
			continue
		}
		out := a.transfer(id, state, false)
		stats.transfers++
		for _, edge := range a.outgoing(id) {
			if !reachable[edge.To] {
				continue
			}
			next := cloneState(out)
			a.applyGuard(id, edge, next)
			merged, changed := joinState(inStates[edge.To], next, inStates[edge.To].vars != nil)
			if !changed {
				continue
			}
			inStates[edge.To] = merged
			if !inQueue[edge.To] {
				heap.Push(&queue, rankedBlock{id: edge.To, rank: worklistRank[edge.To]})
				inQueue[edge.To] = true
			}
		}
	}

	// Finding emission is a projection of the converged block-entry states, not
	// a side effect of the worklist schedule. This keeps diagnostics and their
	// representative paths stable when the fixed-point traversal order changes.
	a.findings = map[string]Finding{}
	blockIDs := make([]cfg.BlockID, 0, len(reachable))
	for id := range reachable {
		blockIDs = append(blockIDs, id)
	}
	sort.Slice(blockIDs, func(i, j int) bool { return blockIDs[i] < blockIDs[j] })
	for _, id := range blockIDs {
		state, ok := inStates[id]
		if !ok {
			continue
		}
		a.transfer(id, state, true)
	}

	findings := make([]Finding, 0, len(a.findings))
	for _, finding := range a.findings {
		findings = append(findings, finding)
	}
	sort.SliceStable(findings, func(i, j int) bool {
		if findings[i].Sink.Range.StartByte != findings[j].Sink.Range.StartByte {
			return findings[i].Sink.Range.StartByte < findings[j].Sink.Range.StartByte
		}
		if findings[i].Sink.Kind != findings[j].Sink.Kind {
			return findings[i].Sink.Kind < findings[j].Sink.Kind
		}
		return sourceKey(findings[i].Source) < sourceKey(findings[j].Source)
	})
	states := map[cfg.BlockID]map[string]State{}
	for id, state := range inStates {
		states[id] = map[string]State{}
		for name, variable := range state.vars {
			states[id][name] = variableState(variable)
		}
	}
	return Result{Findings: findings, States: states}, stats
}

func (a *procedureAnalyzer) transfer(id cfg.BlockID, input abstractState, collectFindings bool) abstractState {
	state := cloneState(input)
	block, ok := a.blocksByID[id]
	if !ok || block.Statement == nil {
		return state
	}
	statement := *block.Statement
	recovered := statement.Recovered || statement.Kind == procedureir.StatementRecovered
	if recovered {
		if target := assignmentTarget(statement); target != "" {
			state.vars[target] = a.unknownValue(statement.Range, statement.ID, "recovered statement")
		}
	}
	if target := assignmentTarget(statement); !recovered && target != "" && statement.Value != nil {
		assigned := a.evalExpression(*statement.Value, state)
		assigned = appendPath(assigned, PathStep{Kind: "assignment", Label: statement.Text, Range: statement.Range, StatementID: statement.ID})
		state.vars[target] = assigned
	}
	for _, call := range a.callsByStatement[statement.ID] {
		a.applySourceWrite(call, state)
		if collectFindings {
			a.inspectSink(call, state)
		}
	}
	return state
}

func (a *procedureAnalyzer) applySourceWrite(call procedureir.CallSite, state abstractState) {
	kind, label, ok := sourceCall(call)
	if !ok || kind != SourceFileInput || len(call.Arguments.ExpressionIDs) == 0 {
		return
	}
	// VBA's Input/Line Input/Get statements write their last argument(s),
	// whereas function-style reads are handled by evalCall on the assignment.
	for _, expressionID := range call.Arguments.ExpressionIDs[1:] {
		expression, exists := a.expressions[expressionID]
		if !exists || expression.Kind != procedureir.ExpressionIdentifier {
			continue
		}
		state.vars[canonicalName(expression.Text)] = valueFromSource(Source{
			Kind: kind, Label: label, Range: call.Range, StatementID: call.StatementID,
		}, PathStep{Kind: "source", Label: call.Callee.Text, Range: call.Range, StatementID: call.StatementID})
	}
}

func (a *procedureAnalyzer) evalExpression(expression procedureir.Expression, state abstractState) value {
	if canonical, ok := a.expressions[expression.ID]; ok {
		expression = canonical
	}
	if expression.Recovered || expression.Kind == procedureir.ExpressionUnknown {
		return a.unknownValue(expression.Range, expression.StatementID, "unsupported expression")
	}
	switch expression.Kind {
	case procedureir.ExpressionLiteral:
		return value{}
	case procedureir.ExpressionIdentifier:
		if variable, ok := state.vars[canonicalName(expression.Text)]; ok {
			return cloneValue(variable)
		}
		if a.declaredNames[canonicalName(expression.Text)] {
			return a.unknownValue(expression.Range, expression.StatementID, "unassigned declared value")
		}
		if kind, label, ok := sourceIdentifier(expression.Text); ok {
			return valueFromSource(Source{Kind: kind, Label: label, Range: expression.Range, StatementID: expression.StatementID}, PathStep{Kind: "source", Label: expression.Text, Range: expression.Range, StatementID: expression.StatementID})
		}
		return a.unknownValue(expression.Range, expression.StatementID, "unassigned value")
	case procedureir.ExpressionParentheses:
		return a.evalChildren(expression.Children, state)
	case procedureir.ExpressionBinary:
		result := a.evalChildren(expression.Children, state)
		if !isEmptyValue(result) {
			return appendPath(result, PathStep{Kind: "concatenation", Label: expression.Text, Range: expression.Range, StatementID: expression.StatementID})
		}
		return result
	case procedureir.ExpressionUnary:
		result := a.evalChildren(expression.Children, state)
		if !isEmptyValue(result) {
			return appendPath(result, PathStep{Kind: "transformation", Label: expression.Text, Range: expression.Range, StatementID: expression.StatementID})
		}
		return result
	case procedureir.ExpressionMember:
		if kind, label, ok := sourceMember(expression.Text); ok {
			return valueFromSource(Source{Kind: kind, Label: label, Range: expression.Range, StatementID: expression.StatementID}, PathStep{Kind: "source", Label: expression.Text, Range: expression.Range, StatementID: expression.StatementID})
		}
		result := a.evalMemberChildren(expression, state)
		if !isEmptyValue(result) {
			return a.unknownTransform(result, expression)
		}
		return result
	case procedureir.ExpressionCall:
		return a.evalCall(expression, state)
	case procedureir.ExpressionNew:
		return value{}
	default:
		return a.unknownValue(expression.Range, expression.StatementID, "unsupported expression")
	}
}

func (a *procedureAnalyzer) evalMemberChildren(expression procedureir.Expression, state abstractState) value {
	member := ""
	if dot := strings.LastIndex(expression.Text, "."); dot >= 0 {
		member = canonicalName(expression.Text[dot+1:])
	}
	var values []value
	for _, id := range expression.Children {
		child, ok := a.expressions[id]
		if !ok {
			continue
		}
		if member != "" && child.Kind == procedureir.ExpressionIdentifier && canonicalName(child.Text) == member {
			continue
		}
		values = append(values, a.evalExpression(child, state))
	}
	return joinValues(values)
}

func (a *procedureAnalyzer) evalCall(expression procedureir.Expression, state abstractState) value {
	call, ok := a.callsByExpression[expression.ID]
	if !ok {
		return a.unknownValue(expression.Range, expression.StatementID, "unresolved call")
	}
	args := a.callArguments(call, state)
	if kind, label, ok := sourceCall(call); ok {
		return valueFromSource(Source{Kind: kind, Label: label, Range: call.Range, StatementID: call.StatementID}, PathStep{Kind: "source", Label: call.Callee.Text, Range: call.Range, StatementID: call.StatementID})
	}
	name := strings.ToLower(call.Callee.BaseName)
	if name == "encodeurl" || strings.EqualFold(call.Callee.Member, "EncodeURL") {
		result := joinValues(args)
		if !isEmptyValue(result) {
			for key, origin := range result.origins {
				origin.safe = copySafe(origin.safe)
				origin.safe[SinkHTTPURL] = true
				origin.path = append(origin.path, PathStep{Kind: "sanitization", Label: call.Callee.Text, Range: call.Range, StatementID: call.StatementID})
				result.origins[key] = origin
			}
		}
		return result
	}
	if name == "trim" || name == "cstr" || name == "replace" {
		result := joinValues(args)
		if !isEmptyValue(result) {
			return appendPath(result, PathStep{Kind: "transformation", Label: call.Callee.Text, Range: call.Range, StatementID: call.StatementID})
		}
		return result
	}
	if name == "isnumeric" || name == "len" {
		return value{}
	}
	result := joinValues(args)
	if isEmptyValue(result) {
		return a.unknownValue(expression.Range, expression.StatementID, "unknown transformation")
	}
	return a.unknownTransform(result, expression)
}

func (a *procedureAnalyzer) evalChildren(ids []int, state abstractState) value {
	var values []value
	for _, id := range ids {
		if expression, ok := a.expressions[id]; ok {
			values = append(values, a.evalExpression(expression, state))
		}
	}
	return joinValues(values)
}

func (a *procedureAnalyzer) callArguments(call procedureir.CallSite, state abstractState) []value {
	args := make([]value, 0, len(call.Arguments.ExpressionIDs))
	for _, id := range call.Arguments.ExpressionIDs {
		if expression, ok := a.expressions[id]; ok {
			args = append(args, a.evalExpression(expression, state))
		}
	}
	return args
}

func (a *procedureAnalyzer) inspectSink(call procedureir.CallSite, state abstractState) {
	targets := sinkTargets(call, a.knownShellObjects)
	if len(targets) == 0 {
		return
	}
	args := a.callArguments(call, state)
	for _, target := range targets {
		if target.argument < 0 || target.argument >= len(args) {
			continue
		}
		origins := args[target.argument].origins
		keys := make([]string, 0, len(origins))
		for key := range origins {
			keys = append(keys, key)
		}
		sort.Strings(keys)
		for _, key := range keys {
			origin := origins[key]
			if origin.safe[target.kind] || (origin.state != StateTainted && origin.state != StateUnknown) {
				continue
			}
			sink := Sink{Kind: target.kind, Label: target.label, Range: call.Range, StatementID: call.StatementID, Argument: target.argument}
			finding := Finding{
				State:   origin.state,
				Source:  origin.source,
				Sink:    sink,
				Path:    append([]PathStep(nil), origin.path...),
				Message: fmt.Sprintf("Conservative analysis found %s flowing to %s.", origin.source.Label, target.label),
				Reason:  "The value is untrusted or crossed an unknown transformation before reaching a sensitive API.",
			}
			key := fmt.Sprintf("%s\x00%d\x00%s", sink.Kind, sink.StatementID, sourceKey(origin.source))
			if previous, ok := a.findings[key]; !ok {
				a.findings[key] = finding
			} else {
				if betterPath(finding.Path, previous.Path) {
					previous.Path = append([]PathStep(nil), finding.Path...)
				}
				previous.State = joinStateValue(previous.State, finding.State)
				a.findings[key] = previous
			}
		}
	}
}

type sinkTarget struct {
	kind     SinkKind
	label    string
	argument int
}

func sinkTargets(call procedureir.CallSite, knownShellObjects map[string]bool) []sinkTarget {
	base := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	receiver := ""
	if call.Callee.Receiver != nil {
		receiver = strings.ToLower(strings.TrimSpace(*call.Callee.Receiver))
	}
	var out []sinkTarget
	if base == "shell" && receiver == "" {
		out = append(out, sinkTarget{SinkShell, "Shell", 0})
	}
	if (member == "run" || member == "exec") && (strings.Contains(full, "wscript.shell") || knownShellObjects[receiver] || strings.Contains(receiver, "shell")) {
		kind, label := SinkWScriptShellRun, "WScript.Shell.Run"
		if member == "exec" {
			kind, label = SinkWScriptShellExec, "WScript.Shell.Exec"
		}
		out = append(out, sinkTarget{kind, label, 0})
	}
	if member == "execute" || base == "execute" || member == "executenonquery" || base == "executenonquery" || member == "openrecordset" || base == "openrecordset" || member == "runsql" || base == "runsql" || member == "executesql" || base == "executesql" {
		if looksDatabaseReceiver(receiver, full) {
			out = append(out, sinkTarget{SinkSQLExecution, "SQL execution", 0})
		}
	}
	if base == "kill" || member == "kill" || base == "rmdir" || member == "rmdir" || member == "deletefile" || base == "deletefile" || member == "deletefolder" || base == "deletefolder" {
		out = append(out, sinkTarget{SinkDestructiveFile, "destructive file operation", 0})
	}
	if member == "open" && (strings.Contains(full, "workbooks.open") || receiver == "workbooks") {
		out = append(out, sinkTarget{SinkWorkbooksOpen, "Workbooks.Open", 0})
	}
	if member == "saveas" || base == "saveas" {
		out = append(out, sinkTarget{SinkSaveAs, "SaveAs", 0})
	}
	if member == "open" && looksHTTPReceiver(receiver, full) {
		out = append(out, sinkTarget{SinkHTTPURL, "HTTP request URL", 1})
	}
	if member == "setrequestheader" && looksHTTPReceiver(receiver, full) {
		out = append(out, sinkTarget{SinkHTTPHeader, "HTTP request header", 0}, sinkTarget{SinkHTTPHeader, "HTTP request header", 1})
	}
	return out
}

func sourceCall(call procedureir.CallSite) (SourceKind, string, bool) {
	base := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	switch strings.TrimSuffix(base, "$") {
	case "inputbox":
		return SourceInputBox, "InputBox", true
	case "environ":
		return SourceEnvironment, "environment variable", true
	case "command":
		return SourceCommandLine, "command-line argument", true
	}
	if base == "readall" || base == "readtext" || base == "readline" || base == "input" || base == "lineinput" || base == "get" || strings.Contains(full, "opentextfile") {
		return SourceFileInput, "text/CSV file input", true
	}
	if strings.Contains(full, "wscript.arguments") {
		return SourceCommandLine, "command-line argument", true
	}
	if (member == "environment" || base == "environment") && strings.Contains(full, "wscript.shell") {
		return SourceEnvironment, "environment variable", true
	}
	if member == "responsetext" || base == "responsetext" || member == "responsebody" || base == "responsebody" || member == "response" || base == "response" || strings.Contains(full, "readresponse") {
		return SourceHTTPResponse, "HTTP response", true
	}
	if member == "fields" || base == "fields" || member == "field" || base == "field" || base == "getrows" || member == "getrows" || base == "openrecordset" || member == "openrecordset" || ((base == "execute" || member == "execute") && looksDatabaseReceiver(strings.ToLower(strings.TrimSpace(receiverText(call))), full)) {
		return SourceDatabaseResult, "database result", true
	}
	return "", "", false
}

func sourceIdentifier(text string) (SourceKind, string, bool) {
	switch strings.TrimSuffix(strings.ToLower(strings.TrimSpace(text)), "$") {
	case "command":
		return SourceCommandLine, "command-line argument", true
	}
	return "", "", false
}

func sourceMember(text string) (SourceKind, string, bool) {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.Contains(lower, "responsetext") || strings.Contains(lower, "responsebody") || (strings.Contains(lower, ".response") && (strings.Contains(lower, "http") || strings.Contains(lower, "xmlhttp") || strings.Contains(lower, "winhttp"))) {
		return SourceHTTPResponse, "HTTP response", true
	}
	if (strings.Contains(lower, "range(") || strings.Contains(lower, "cells(") || strings.Contains(lower, "worksheet")) && (strings.HasSuffix(lower, ".value") || strings.HasSuffix(lower, ".text") || strings.HasSuffix(lower, ".formula")) {
		return SourceWorksheetCell, "worksheet cell value", true
	}
	if strings.Contains(lower, "recordset") && (strings.Contains(lower, ".fields") || strings.Contains(lower, ".value")) {
		return SourceDatabaseResult, "database result", true
	}
	if strings.Contains(lower, "wscript.arguments") || (strings.Contains(lower, ".arguments") && strings.Contains(lower, "wscript")) {
		return SourceCommandLine, "command-line argument", true
	}
	if strings.Contains(lower, "wscript.shell.environment") {
		return SourceEnvironment, "environment variable", true
	}
	return "", "", false
}

func receiverText(call procedureir.CallSite) string {
	if call.Callee.Receiver == nil {
		return ""
	}
	return *call.Callee.Receiver
}

func looksDatabaseReceiver(receiver, full string) bool {
	text := strings.ToLower(receiver + " " + full)
	for _, marker := range []string{"conn", "connection", "command", "database", "currentdb", "recordset", "ado", "adodb", "docmd"} {
		if strings.Contains(text, marker) {
			return true
		}
	}
	for _, marker := range []string{"db", "sql"} {
		if containsIdentifierToken(text, marker) {
			return true
		}
	}
	for _, name := range []string{"cn", "cmd", "rs", "db"} {
		if receiver == name {
			return true
		}
	}
	return false
}

func containsIdentifierToken(text, token string) bool {
	for start := 0; start <= len(text)-len(token); start++ {
		if !strings.HasPrefix(text[start:], token) {
			continue
		}
		end := start + len(token)
		if (start == 0 || !isIdentifierByte(text[start-1])) && (end == len(text) || !isIdentifierByte(text[end])) {
			return true
		}
	}
	return false
}

func isIdentifierByte(value byte) bool {
	return value >= 'a' && value <= 'z' || value >= '0' && value <= '9' || value == '_'
}

func looksHTTPReceiver(receiver, full string) bool {
	text := strings.ToLower(receiver + " " + full)
	for _, marker := range []string{"http", "xmlhttp", "winhttp", "request"} {
		if strings.Contains(text, marker) {
			return true
		}
	}
	if receiver == "req" || receiver == "xhr" || receiver == "client" {
		return true
	}
	return false
}

func (a *procedureAnalyzer) scanKnownShellObjects() {
	assign := regexp.MustCompile(`(?i)\b(?:set\s+)?([a-z_][a-z0-9_]*)\s*=\s*(?:createobject\s*\(\s*"wscript\.shell"|new\s+wscript\.shell)`)
	for _, statement := range a.procedure.Statements {
		if match := assign.FindStringSubmatch(statement.Text); len(match) > 1 {
			a.knownShellObjects[canonicalName(match[1])] = true
		}
	}
}

func (a *procedureAnalyzer) unknownValue(r vbaast.Range, statementID int, label string) value {
	return valueFromSourceState(Source{Kind: SourceUnknown, Label: "unknown input", Range: r, StatementID: statementID}, StateUnknown, PathStep{Kind: "unknown_transformation", Label: label, Range: r, StatementID: statementID})
}

func (a *procedureAnalyzer) unknownTransform(input value, expression procedureir.Expression) value {
	result := cloneValue(input)
	for key, origin := range result.origins {
		origin.state = joinStateValue(origin.state, StateUnknown)
		origin.path = append(origin.path, PathStep{Kind: "unknown_transformation", Label: expression.Text, Range: expression.Range, StatementID: expression.StatementID})
		result.origins[key] = origin
	}
	return result
}

func (a *procedureAnalyzer) applyGuard(from cfg.BlockID, edge cfg.Edge, state abstractState) {
	if edge.Kind != cfg.EdgeBranchTrue && edge.Kind != cfg.EdgeCase {
		return
	}
	block, ok := a.blocksByID[from]
	if !ok || block.Statement == nil {
		return
	}
	statement := *block.Statement
	if edge.Kind == cfg.EdgeBranchTrue {
		if variable, ok := exactEqualityVariable(statement.Text); ok {
			validated := state.vars[variable]
			markSafe(&validated)
			state.vars[variable] = validated
		}
	}
	if edge.Kind == cfg.EdgeCase {
		caseBlock, ok := a.blocksByID[edge.To]
		if !ok || caseBlock.Statement == nil {
			return
		}
		variable := a.selectVariables[caseBlock.Statement.ParentID]
		if variable == "" {
			variable = a.nearestSelector(*caseBlock.Statement)
		}
		if variable != "" && caseLiteral(caseBlock.Statement.Text) {
			validated := state.vars[variable]
			markSafe(&validated)
			state.vars[variable] = validated
		}
	}
}

func (a *procedureAnalyzer) nearestSelector(statement procedureir.Statement) string {
	visited := map[int]bool{}
	for current := statement.ParentID; current != 0; {
		if visited[current] {
			break
		}
		visited[current] = true
		parent, ok := a.statements[current]
		if !ok {
			break
		}
		if variable := a.selectVariables[parent.ID]; variable != "" {
			return variable
		}
		current = parent.ParentID
	}
	return ""
}

func markSafe(variable *value) {
	if variable == nil {
		return
	}
	for key, origin := range variable.origins {
		origin.safe = copySafe(origin.safe)
		for _, sink := range SinkCatalog() {
			origin.safe[sink.Kind] = true
		}
		variable.origins[key] = origin
	}
}

func exactEqualityVariable(text string) (string, bool) {
	match := regexp.MustCompile(`(?i)\bif\s+([a-z_][a-z0-9_]*)\s*=\s*("[^"]*"|[-+]?[0-9]+)\s+then\b`).FindStringSubmatch(strings.TrimSpace(text))
	if len(match) == 0 {
		return "", false
	}
	return canonicalName(match[1]), true
}

func selectorVariable(text string) string {
	match := regexp.MustCompile(`(?i)\bselect\s+case\s+([a-z_][a-z0-9_]*)`).FindStringSubmatch(text)
	if len(match) == 0 {
		return ""
	}
	return canonicalName(match[1])
}

func caseLiteral(text string) bool {
	firstLine := strings.SplitN(text, "\n", 2)[0]
	return regexp.MustCompile(`(?i)^\s*case\s+("[^"]*"|[-+]?[0-9]+)(?:\s*,|\s*$)`).MatchString(strings.TrimSpace(firstLine))
}

func assignmentTarget(statement procedureir.Statement) string {
	if statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet {
		return ""
	}
	if statement.Target != nil && statement.Target.Kind == procedureir.ExpressionIdentifier {
		return canonicalName(statement.Target.Text)
	}
	text := strings.TrimSpace(statement.Text)
	match := regexp.MustCompile(`(?i)^(?:set\s+)?([a-z_][a-z0-9_]*)\s*=`).FindStringSubmatch(text)
	if len(match) > 1 {
		return canonicalName(match[1])
	}
	return ""
}

func (a *procedureAnalyzer) outgoing(from cfg.BlockID) []cfg.Edge {
	return a.outgoingEdges[from]
}

func (a *procedureAnalyzer) reversePostOrderRank(reachable map[cfg.BlockID]bool) map[cfg.BlockID]int {
	visited := map[cfg.BlockID]bool{}
	postorder := make([]cfg.BlockID, 0, len(reachable))
	type frame struct {
		id   cfg.BlockID
		next int
	}
	visited[a.graph.Entry] = true
	stack := []frame{{id: a.graph.Entry}}
	for len(stack) > 0 {
		current := &stack[len(stack)-1]
		edges := a.outgoing(current.id)
		if current.next < len(edges) {
			target := edges[current.next].To
			current.next++
			if reachable[target] && !visited[target] {
				visited[target] = true
				stack = append(stack, frame{id: target})
			}
			continue
		}
		postorder = append(postorder, current.id)
		stack = stack[:len(stack)-1]
	}
	rank := make(map[cfg.BlockID]int, len(postorder))
	for index := len(postorder) - 1; index >= 0; index-- {
		rank[postorder[index]] = len(postorder) - 1 - index
	}
	return rank
}

type rankedBlock struct {
	id   cfg.BlockID
	rank int
}

type rankedWorklist []rankedBlock

func (w rankedWorklist) Len() int { return len(w) }
func (w rankedWorklist) Less(i, j int) bool {
	if w[i].rank != w[j].rank {
		return w[i].rank < w[j].rank
	}
	return w[i].id < w[j].id
}
func (w rankedWorklist) Swap(i, j int) { w[i], w[j] = w[j], w[i] }
func (w *rankedWorklist) Push(value any) {
	*w = append(*w, value.(rankedBlock))
}
func (w *rankedWorklist) Pop() any {
	old := *w
	last := len(old) - 1
	value := old[last]
	*w = old[:last]
	return value
}

func canonicalName(name string) string {
	name = strings.TrimSpace(name)
	name = strings.Trim(name, "[]")
	if dot := strings.LastIndex(name, "."); dot >= 0 {
		name = name[dot+1:]
	}
	return strings.ToLower(name)
}

func valueFromSource(source Source, step PathStep) value {
	return valueFromSourceState(source, StateTainted, step)
}

func valueFromSourceState(source Source, state State, step PathStep) value {
	return value{origins: map[string]provenance{sourceKey(source): {source: source, state: state, path: []PathStep{step}, safe: map[SinkKind]bool{}}}}
}

func joinValues(values []value) value {
	result := value{origins: map[string]provenance{}}
	for _, current := range values {
		result, _ = joinValue(result, current, true)
	}
	return result
}

func joinValue(a, b value, initialized bool) (value, bool) {
	if !initialized {
		return cloneValue(b), !isSameValue(a, b)
	}
	result := cloneValue(a)
	changed := false
	for key, origin := range b.origins {
		if previous, ok := result.origins[key]; ok {
			merged := previous
			merged.state = joinStateValue(previous.state, origin.state)
			merged.path = betterPathValue(previous.path, origin.path)
			merged.safe = intersectSafe(previous.safe, origin.safe)
			if !sameProvenance(previous, merged) {
				result.origins[key] = merged
				changed = true
			}
			continue
		}
		result.origins[key] = cloneProvenance(origin)
		changed = true
	}
	return result, changed
}

func joinState(a, b abstractState, initialized bool) (abstractState, bool) {
	if !initialized {
		return cloneState(b), true
	}
	result := cloneState(a)
	changed := false
	keys := map[string]bool{}
	for key := range a.vars {
		keys[key] = true
	}
	for key := range b.vars {
		keys[key] = true
	}
	for key := range keys {
		left, leftOK := a.vars[key]
		right, rightOK := b.vars[key]
		if !leftOK {
			left = unknownStandaloneValue()
		}
		if !rightOK {
			right = unknownStandaloneValue()
		}
		merged, _ := joinValue(left, right, true)
		previous, hadPrevious := a.vars[key]
		if !hadPrevious || !isSameValue(previous, merged) {
			result.vars[key] = merged
			changed = true
		}
	}
	return result, changed
}

func unknownStandaloneValue() value {
	return valueFromSourceState(Source{Kind: SourceUnknown, Label: "unknown input"}, StateUnknown, PathStep{Kind: "unknown_transformation", Label: "possibly unassigned value"})
}

func cloneState(state abstractState) abstractState {
	result := abstractState{vars: map[string]value{}}
	for key, variable := range state.vars {
		result.vars[key] = cloneValue(variable)
	}
	return result
}

func cloneValue(input value) value {
	result := value{origins: map[string]provenance{}}
	for key, origin := range input.origins {
		result.origins[key] = cloneProvenance(origin)
	}
	return result
}

func cloneProvenance(input provenance) provenance {
	input.path = append([]PathStep(nil), input.path...)
	input.safe = copySafe(input.safe)
	return input
}

func copySafe(input map[SinkKind]bool) map[SinkKind]bool {
	result := map[SinkKind]bool{}
	for key, value := range input {
		result[key] = value
	}
	return result
}

func intersectSafe(a, b map[SinkKind]bool) map[SinkKind]bool {
	result := map[SinkKind]bool{}
	for key, value := range a {
		if value && b[key] {
			result[key] = true
		}
	}
	return result
}

func appendPath(input value, step PathStep) value {
	result := cloneValue(input)
	for key, origin := range result.origins {
		origin.path = append(origin.path, step)
		result.origins[key] = origin
	}
	return result
}

func variableState(input value) State {
	state := StateClean
	for _, origin := range input.origins {
		state = joinStateValue(state, origin.state)
	}
	return state
}

func joinStateValue(a, b State) State {
	if a == StateUnknown || b == StateUnknown {
		return StateUnknown
	}
	if a == StateTainted || b == StateTainted {
		return StateTainted
	}
	return StateClean
}

func isEmptyValue(input value) bool { return len(input.origins) == 0 }

func sameProvenance(a, b provenance) bool {
	return a.state == b.state && sourceKey(a.source) == sourceKey(b.source) && pathKey(a.path) == pathKey(b.path) && sameSafe(a.safe, b.safe)
}

func sameSafe(a, b map[SinkKind]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for key, value := range a {
		if b[key] != value {
			return false
		}
	}
	return true
}

func isSameValue(a, b value) bool {
	if len(a.origins) != len(b.origins) {
		return false
	}
	for key, left := range a.origins {
		right, ok := b.origins[key]
		if !ok || !sameProvenance(left, right) {
			return false
		}
	}
	return true
}

func betterPath(a, b []PathStep) bool {
	if len(a) != len(b) {
		return len(a) < len(b)
	}
	return pathKey(a) < pathKey(b)
}

func betterPathValue(a, b []PathStep) []PathStep {
	if betterPath(b, a) {
		return append([]PathStep(nil), b...)
	}
	return append([]PathStep(nil), a...)
}
