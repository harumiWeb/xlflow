package dataflow

import (
	"container/heap"
	"context"
	"fmt"
	"regexp"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var (
	cmdInterpreterPattern        = regexp.MustCompile(`(?i)\bcmd(?:\.exe)?\b`)
	cmdSwitchPattern             = regexp.MustCompile(`(?i)(?:^|\s)[/-][ck](?:\s|$)`)
	powershellInterpreterPattern = regexp.MustCompile(`(?i)\b(?:powershell|pwsh)(?:\.exe)?\b`)
	scriptHostInterpreterPattern = regexp.MustCompile(`(?i)\b(?:wscript|cscript|mshta)(?:\.exe)?\b`)
	powershellCommandPattern     = regexp.MustCompile(`(?i)(?:^|\s)[/-](?:command|c)(?:\s|$)`)
	credentialAssignmentPattern  = regexp.MustCompile(`(?i)(?:password|passwd|secret|token|api[_-]?key|authorization)\s*=`)
	credentialWordPattern        = regexp.MustCompile(`(?i)\b(?:password|passwd|secret|token|apikey|api_key)\b`)
	knownShellObjectPattern      = regexp.MustCompile(`(?i)\b(?:set\s+)?([a-z_][a-z0-9_]*)\s*=\s*(?:createobject\s*\(\s*"(wscript\.shell|shell\.application)"|new\s+(wscript\.shell|shell\.application))`)
)

// AnalyzeProcedure performs a conservative, intra-procedural forward
// may-analysis. A supplied graph is used as-is; when it is empty a graph is
// built from the procedure so small protocol adapters can use the API safely.
func AnalyzeProcedure(procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) Result {
	result, _ := AnalyzeProcedureContext(context.Background(), procedure, graph, options)
	return result
}

// AnalyzeProcedureContext performs the same analysis as AnalyzeProcedure but
// observes ctx and returns its cancellation or deadline error explicitly.
// Cancellation never produces a partial Result: callers must handle the
// returned error before using any findings.
func AnalyzeProcedureContext(ctx context.Context, procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) (Result, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return Result{}, err
	}
	if len(graph.Blocks) == 0 {
		var err error
		graph, err = cfg.BuildContext(ctx, procedure)
		if err != nil {
			return Result{}, err
		}
	}
	a, err := newProcedureAnalyzerContext(ctx, procedure, graph, options)
	if err != nil {
		return Result{}, err
	}
	return a.runContext()
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

type sqlObjectKind string

const (
	sqlObjectUnknown    sqlObjectKind = "unknown"
	sqlObjectCommand    sqlObjectKind = "command"
	sqlObjectRecordset  sqlObjectKind = "recordset"
	sqlObjectQueryDef   sqlObjectKind = "querydef"
	sqlObjectConnection sqlObjectKind = "connection"
)

// sqlObjectState is intentionally small. It records only the state needed to
// connect a CommandText/property assignment or an appended parameter with a
// later Execute/Open call in the same procedure.
type sqlObjectState struct {
	kind          sqlObjectKind
	identity      string
	commandText   value
	parameters    map[string]value
	parameterized bool
}

type abstractState struct {
	vars             map[string]value
	sqlObjects       map[string]sqlObjectState
	varsShared       bool
	sqlObjectsShared bool
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
	constantNames     map[string]bool
	knownShellObjects map[string]string
	selectVariables   map[int]string
	findings          map[string]Finding
	commandFindings   map[string]CommandFinding
	sqlFindings       map[string]SQLFinding
	blocksByID        map[cfg.BlockID]cfg.Block
	outgoingEdges     map[cfg.BlockID][]cfg.Edge
	ctx               context.Context
}

type analysisStats struct {
	transfers int
}

func newProcedureAnalyzer(procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) *procedureAnalyzer {
	a, _ := newProcedureAnalyzerContext(context.Background(), procedure, graph, options)
	return a
}

func newProcedureAnalyzerContext(ctx context.Context, procedure procedureir.ProcedureIR, graph cfg.Graph, options Options) (*procedureAnalyzer, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	a := &procedureAnalyzer{
		procedure:         procedure,
		graph:             graph,
		options:           options,
		statements:        map[int]procedureir.Statement{},
		expressions:       map[int]procedureir.Expression{},
		callsByExpression: map[int]procedureir.CallSite{},
		callsByStatement:  map[int][]procedureir.CallSite{},
		declaredNames:     map[string]bool{},
		constantNames:     map[string]bool{},
		knownShellObjects: map[string]string{},
		selectVariables:   map[int]string{},
		findings:          map[string]Finding{},
		commandFindings:   map[string]CommandFinding{},
		sqlFindings:       map[string]SQLFinding{},
		blocksByID:        map[cfg.BlockID]cfg.Block{},
		outgoingEdges:     map[cfg.BlockID][]cfg.Edge{},
		ctx:               ctx,
	}
	for i, block := range graph.Blocks {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		a.blocksByID[block.ID] = block
	}
	for i, edge := range graph.Edges {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		a.outgoingEdges[edge.From] = append(a.outgoingEdges[edge.From], edge)
	}
	for from := range a.outgoingEdges {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		sort.SliceStable(a.outgoingEdges[from], func(i, j int) bool {
			left, right := a.outgoingEdges[from][i], a.outgoingEdges[from][j]
			if left.To != right.To {
				return left.To < right.To
			}
			return left.ID < right.ID
		})
	}
	for i, statement := range procedure.Statements {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		a.statements[statement.ID] = statement
	}
	for i, declaration := range procedure.Declarations {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		name := canonicalName(declaration.Name)
		a.declaredNames[name] = true
		if declaration.Kind == "const" {
			a.constantNames[name] = true
		}
	}
	for i, parameter := range procedure.Symbol.Parameters {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		a.declaredNames[canonicalName(parameter.Name)] = true
	}
	for i, expression := range procedure.Expressions {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		a.expressions[expression.ID] = expression
	}
	for i, call := range procedure.Calls {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		if call.ExpressionID != 0 {
			a.callsByExpression[call.ExpressionID] = call
		}
		a.callsByStatement[call.StatementID] = append(a.callsByStatement[call.StatementID], call)
	}
	for i, statement := range procedure.Statements {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		if statement.Kind != procedureir.StatementSelect {
			continue
		}
		if variable := selectorVariable(statement.Text); variable != "" {
			a.selectVariables[statement.ID] = variable
		}
	}
	a.scanKnownShellObjects()
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	return a, nil
}

func (a *procedureAnalyzer) runWithStats() (Result, analysisStats) {
	result, stats, _ := a.runWithStatsContext()
	return result, stats
}

func (a *procedureAnalyzer) runWithStatsAndRank(worklistRank map[cfg.BlockID]int) (Result, analysisStats) {
	result, stats, _ := a.runWithStatsAndRankContext(worklistRank)
	return result, stats
}

func (a *procedureAnalyzer) runContext() (Result, error) {
	result, _, err := a.runWithStatsContext()
	return result, err
}

func (a *procedureAnalyzer) runWithStatsContext() (Result, analysisStats, error) {
	return a.runWithStatsAndRankContext(nil)
}

func (a *procedureAnalyzer) runWithStatsAndRankContext(worklistRank map[cfg.BlockID]int) (Result, analysisStats, error) {
	stats := analysisStats{}
	if err := a.contextErr(); err != nil {
		return Result{}, stats, err
	}
	entry := abstractState{vars: map[string]value{}, sqlObjects: map[string]sqlObjectState{}}
	for i, parameter := range a.procedure.Symbol.Parameters {
		if i&0xff == 0 {
			if err := a.contextErr(); err != nil {
				return Result{}, stats, err
			}
		}
		entry.vars[canonicalName(parameter.Name)] = valueFromSource(Source{
			Kind:        SourceParameter,
			Label:       parameter.Name,
			Range:       parameter.Range,
			StatementID: 0,
		}, PathStep{Kind: "source", Label: parameter.Name, Range: parameter.Range})
	}

	reachable, err := a.reachableContext()
	if err != nil {
		return Result{}, stats, err
	}
	if len(reachable) == 0 {
		return Result{Findings: nil, CommandFindings: nil, SQLFindings: nil, States: nil}, stats, nil
	}
	inStates := map[cfg.BlockID]abstractState{a.graph.Entry: entry}
	if worklistRank == nil {
		worklistRank, err = a.reversePostOrderRankContext(reachable)
		if err != nil {
			return Result{}, stats, err
		}
	}
	queue := rankedWorklist{{id: a.graph.Entry, rank: worklistRank[a.graph.Entry]}}
	heap.Init(&queue)
	inQueue := map[cfg.BlockID]bool{a.graph.Entry: true}
	for len(queue) > 0 {
		if err := a.contextErr(); err != nil {
			return Result{}, stats, err
		}
		id := heap.Pop(&queue).(rankedBlock).id
		inQueue[id] = false
		state, ok := inStates[id]
		if !ok {
			continue
		}
		out, err := a.transfer(id, state, false)
		if err != nil {
			return Result{}, stats, err
		}
		stats.transfers++
		for _, edge := range a.outgoing(id) {
			if err := a.contextErr(); err != nil {
				return Result{}, stats, err
			}
			if !reachable[edge.To] {
				continue
			}
			next := cloneState(&out)
			a.applyGuard(id, edge, &next)
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
	a.commandFindings = map[string]CommandFinding{}
	a.sqlFindings = map[string]SQLFinding{}
	blockIDs := make([]cfg.BlockID, 0, len(reachable))
	for id := range reachable {
		if err := a.contextErr(); err != nil {
			return Result{}, stats, err
		}
		blockIDs = append(blockIDs, id)
	}
	sort.Slice(blockIDs, func(i, j int) bool { return blockIDs[i] < blockIDs[j] })
	for _, id := range blockIDs {
		if err := a.contextErr(); err != nil {
			return Result{}, stats, err
		}
		state, ok := inStates[id]
		if !ok {
			continue
		}
		if _, err := a.transfer(id, state, true); err != nil {
			return Result{}, stats, err
		}
	}

	findings := make([]Finding, 0, len(a.findings))
	for _, finding := range a.findings {
		findings = append(findings, finding)
	}
	commandFindings := make([]CommandFinding, 0, len(a.commandFindings))
	for _, finding := range a.commandFindings {
		commandFindings = append(commandFindings, finding)
	}
	sqlFindings := make([]SQLFinding, 0, len(a.sqlFindings))
	for _, finding := range a.sqlFindings {
		sqlFindings = append(sqlFindings, finding)
	}
	sort.SliceStable(sqlFindings, func(i, j int) bool {
		left, right := sqlFindings[i], sqlFindings[j]
		if left.Execution.Range.StartByte != right.Execution.Range.StartByte {
			return left.Execution.Range.StartByte < right.Execution.Range.StartByte
		}
		if left.RiskKind != right.RiskKind {
			return left.RiskKind < right.RiskKind
		}
		return sourceKey(left.Source) < sourceKey(right.Source)
	})
	sort.SliceStable(commandFindings, func(i, j int) bool {
		left, right := commandFindings[i], commandFindings[j]
		if left.Execution.Range.StartByte != right.Execution.Range.StartByte {
			return left.Execution.Range.StartByte < right.Execution.Range.StartByte
		}
		if left.RiskKind != right.RiskKind {
			return left.RiskKind < right.RiskKind
		}
		return sourceKey(left.Source) < sourceKey(right.Source)
	})
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
		if err := a.contextErr(); err != nil {
			return Result{}, stats, err
		}
		states[id] = map[string]State{}
		for name, variable := range state.vars {
			states[id][name] = variableState(variable)
		}
	}
	return Result{Findings: findings, CommandFindings: commandFindings, SQLFindings: sqlFindings, States: states}, stats, nil
}

func (a *procedureAnalyzer) transfer(id cfg.BlockID, input abstractState, collectFindings bool) (abstractState, error) {
	if err := a.contextErr(); err != nil {
		return abstractState{}, err
	}
	state := cloneState(&input)
	block, ok := a.blocksByID[id]
	if !ok || block.Statement == nil {
		return state, nil
	}
	statement := *block.Statement
	recovered := statement.Recovered || statement.Kind == procedureir.StatementRecovered
	if recovered {
		if target := assignmentTarget(statement); target != "" {
			state.ensureVars()
			state.vars[target] = a.unknownValue(statement.Range, statement.ID, "recovered statement")
		}
	}
	if target := assignmentTarget(statement); !recovered && target != "" && statement.Value != nil {
		assigned := a.evalExpression(*statement.Value, state)
		if err := a.contextErr(); err != nil {
			return abstractState{}, err
		}
		assigned = appendPath(assigned, PathStep{Kind: "assignment", Label: statement.Text, Range: statement.Range, StatementID: statement.ID})
		state.ensureVars()
		state.vars[target] = assigned
	}
	if !recovered {
		a.applySQLAssignment(statement, &state)
	}
	for _, call := range a.callsByStatement[statement.ID] {
		if err := a.contextErr(); err != nil {
			return abstractState{}, err
		}
		a.applySourceWrite(call, &state)
		a.applySQLCallState(call, &state)
		if collectFindings {
			a.inspectSink(call, state)
			a.inspectSQL(call, state)
			a.inspectCommand(call, state)
		}
	}
	if err := a.contextErr(); err != nil {
		return abstractState{}, err
	}
	return state, nil
}

// applySQLAssignment records the small amount of object/property state needed
// to connect a CommandText assignment with a later Command.Execute call. It
// deliberately does not attempt whole-project aliasing or interprocedural
// propagation; those remain outside the procedure-local data-flow contract.
func (a *procedureAnalyzer) applySQLAssignment(statement procedureir.Statement, state *abstractState) {
	if state == nil {
		return
	}
	if state.sqlObjects == nil {
		state.sqlObjects = map[string]sqlObjectState{}
	}
	if statement.Target != nil && statement.Target.Kind == procedureir.ExpressionMember {
		receiver, member := sqlMemberTarget(statement.Target.Text)
		if strings.EqualFold(member, "commandtext") && receiver != "" {
			object := state.sqlObjects[receiver]
			if object.kind == sqlObjectUnknown {
				object.kind = a.sqlObjectKind(receiver)
			}
			if object.identity == "" {
				object.identity = receiver
			}
			if statement.Value != nil {
				object.commandText = a.evalExpression(*statement.Value, *state)
			}
			state.ensureSQLObjects()
			state.sqlObjects[receiver] = object
			propagateSQLAlias(state, receiver, object)
		}
		return
	}
	target := assignmentTarget(statement)
	if target == "" {
		return
	}
	rhs := ""
	if statement.Value != nil {
		rhs = statement.Value.Text
	}
	if rhs == "" {
		if _, after, ok := strings.Cut(statement.Text, "="); ok {
			rhs = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(after), "Set "))
		}
	}
	if kind := sqlObjectKindFromText(rhs); kind != sqlObjectUnknown {
		state.ensureSQLObjects()
		state.sqlObjects[target] = sqlObjectState{kind: kind, identity: target, parameters: map[string]value{}}
		return
	}
	if source := canonicalName(rhs); source != "" {
		if object, ok := state.sqlObjects[source]; ok {
			if object.identity == "" {
				object.identity = source
			}
			state.ensureSQLObjects()
			state.sqlObjects[target] = cloneSQLObjectState(object)
			return
		}
	}
	if kind := a.sqlObjectKind(target); kind != sqlObjectUnknown {
		object := state.sqlObjects[target]
		object.kind = kind
		if object.identity == "" {
			object.identity = target
		}
		if object.parameters == nil {
			object.parameters = map[string]value{}
		}
		state.ensureSQLObjects()
		state.sqlObjects[target] = object
	}
}

func (a *procedureAnalyzer) applySQLCallState(call procedureir.CallSite, state *abstractState) {
	if state == nil {
		return
	}
	if state.sqlObjects == nil {
		state.sqlObjects = map[string]sqlObjectState{}
	}
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	rawReceiver := strings.TrimSpace(receiverText(call))
	receiver := canonicalName(rawReceiver)
	if strings.HasSuffix(strings.ToLower(rawReceiver), ".parameters") {
		receiver = canonicalName(rawReceiver[:len(rawReceiver)-len(".parameters")])
	}
	if strings.EqualFold(call.Callee.BaseName, "CallByName") {
		if target, name, invocation, valueID, ok := callByNameSQLParts(call, a.expressions); ok {
			switch strings.ToLower(name) {
			case "commandtext":
				if strings.EqualFold(invocation, "vblet") {
					object := state.sqlObjects[target]
					if object.kind == sqlObjectUnknown {
						object.kind = a.sqlObjectKind(target)
					}
					if object.identity == "" {
						object.identity = target
					}
					if expression, exists := a.expressions[valueID]; exists {
						object.commandText = a.evalExpression(expression, *state)
					}
					state.ensureSQLObjects()
					state.sqlObjects[target] = object
					propagateSQLAlias(state, target, object)
				}
			case "append", "createparameter":
				markSQLParameterized(state, target)
			}
		}
		return
	}
	if member == "append" && strings.Contains(full, ".parameters") {
		markSQLParameterized(state, receiver)
	}
	if member == "createparameter" || strings.HasSuffix(full, ".createparameter") {
		markSQLParameterized(state, receiver)
	}
}

func propagateSQLAlias(state *abstractState, name string, object sqlObjectState) {
	if state == nil {
		return
	}
	state.ensureSQLObjects()
	for aliasName, alias := range state.sqlObjects {
		if aliasName == name || alias.identity != object.identity {
			continue
		}
		alias.commandText = cloneValue(object.commandText)
		alias.parameterized = object.parameterized
		state.sqlObjects[aliasName] = alias
	}
}

func markSQLParameterized(state *abstractState, command string) {
	if state == nil {
		return
	}
	object, ok := state.sqlObjects[command]
	if !ok {
		return
	}
	object.parameterized = true
	if object.identity == "" {
		object.identity = command
	}
	state.ensureSQLObjects()
	state.sqlObjects[command] = object
	for aliasName, alias := range state.sqlObjects {
		if aliasName == command || alias.identity != object.identity {
			continue
		}
		alias.parameterized = true
		state.sqlObjects[aliasName] = alias
	}
}

func callByNameSQLParts(call procedureir.CallSite, expressions map[int]procedureir.Expression) (target, member, invocation string, valueID int, ok bool) {
	ids := call.Arguments.ExpressionIDs
	if len(ids) < 3 {
		return "", "", "", 0, false
	}
	first, firstOK := expressions[ids[0]]
	second, secondOK := expressions[ids[1]]
	third, thirdOK := expressions[ids[2]]
	if !firstOK || !secondOK || !thirdOK || first.Kind != procedureir.ExpressionIdentifier || second.Kind != procedureir.ExpressionLiteral {
		return "", "", "", 0, false
	}
	target = canonicalName(first.Text)
	member = strings.Trim(strings.TrimSpace(second.Text), "\"")
	invocation = strings.Trim(strings.TrimSpace(third.Text), "\"")
	if len(ids) >= 4 {
		valueID = ids[3]
	}
	return target, member, invocation, valueID, true
}

func (a *procedureAnalyzer) inspectSQL(call procedureir.CallSite, state abstractState) {
	api, argument, receiver, commandExecute, ok := a.sqlCallTarget(call, state)
	if !ok {
		return
	}
	var input value
	var expressionText string
	parameterized := false
	if commandExecute {
		object, exists := state.sqlObjects[receiver]
		if !exists {
			return
		}
		input = object.commandText
		parameterized = object.parameterized
		if len(input.origins) == 0 {
			return
		}
		for _, origin := range input.origins {
			for index := len(origin.path) - 1; index >= 0; index-- {
				if origin.path[index].Kind == "concatenation" {
					expressionText = origin.path[index].Label
					break
				}
			}
			if expressionText != "" {
				break
			}
		}
	} else {
		args := a.callArguments(call, state)
		if argument < 0 || argument >= len(args) {
			return
		}
		input = args[argument]
		if argument < len(call.Arguments.ExpressionIDs) {
			if expression, exists := a.expressions[call.Arguments.ExpressionIDs[argument]]; exists {
				expressionText = expression.Text
			}
		}
	}
	if len(input.origins) == 0 {
		return
	}
	for _, origin := range input.origins {
		if origin.state != StateTainted && origin.state != StateUnknown {
			continue
		}
		role := classifySQLRole(expressionText, origin.path, origin.source.Label)
		risk := classifySQLRisk(a, origin, role, expressionText)
		finding := SQLFinding{
			State:         origin.state,
			Source:        origin.source,
			Execution:     SQLExecution{API: api, Role: role, Range: call.Range, Argument: argument},
			RiskKind:      risk,
			OriginState:   origin.state,
			Parameterized: parameterized,
			Path:          append([]PathStep(nil), origin.path...),
			Message:       fmt.Sprintf("Potential SQL construction risk: %s reaches %s.", origin.source.Label, api),
			Reason:        "Conservative analysis found external or unknown data in SQL text; this is not proof of a SQL-injection exploit.",
		}
		findingKey := fmt.Sprintf("%d:%s:%s:%s", call.Range.StartByte, risk, sourceKey(origin.source), api)
		if previous, exists := a.sqlFindings[findingKey]; exists {
			if betterPath(finding.Path, previous.Path) {
				previous.Path = finding.Path
			}
			a.sqlFindings[findingKey] = previous
		} else {
			a.sqlFindings[findingKey] = finding
		}
	}
}

func (a *procedureAnalyzer) sqlCallTarget(call procedureir.CallSite, state abstractState) (api string, argument int, receiver string, commandExecute bool, ok bool) {
	base := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	receiver = canonicalName(receiverText(call))
	if strings.EqualFold(call.Callee.BaseName, "CallByName") {
		if target, name, invocation, valueID, partsOK := callByNameSQLParts(call, a.expressions); partsOK {
			receiver = target
			switch strings.ToLower(name) {
			case "execute":
				if strings.EqualFold(invocation, "vbmethod") {
					switch sqlObjectKindForState(a, receiver, state) {
					case sqlObjectCommand:
						return "ADODB.Command.Execute", -1, receiver, true, true
					case sqlObjectConnection:
						if valueID != 0 {
							return "ADODB.Connection.Execute", 3, receiver, false, true
						}
					case sqlObjectQueryDef:
						if valueID != 0 {
							return "Database.Execute", 3, receiver, false, true
						}
					}
				}
			case "open":
				if strings.EqualFold(invocation, "vbmethod") && sqlObjectKindForState(a, receiver, state) == sqlObjectRecordset {
					if valueID == 0 {
						return "", -1, receiver, false, false
					}
					return "ADODB.Recordset.Open", 3, receiver, false, true
				}
			case "runsql":
				if strings.EqualFold(invocation, "vbmethod") && (receiver == "docmd" || receiver == "") {
					if valueID == 0 {
						return "", -1, receiver, false, false
					}
					return "DoCmd.RunSQL", 3, receiver, false, true
				}
			case "openrecordset":
				if strings.EqualFold(invocation, "vbmethod") && (receiver == "currentdb" || sqlObjectKindForState(a, receiver, state) == sqlObjectQueryDef) {
					if valueID == 0 {
						return "", -1, receiver, false, false
					}
					return "DAO.OpenRecordset", 3, receiver, false, true
				}
			}
		}
		return "", -1, receiver, false, false
	}
	if (member == "execute" || base == "execute" || member == "executenonquery" || base == "executenonquery") && sqlObjectKindForState(a, receiver, state) == sqlObjectCommand {
		return "ADODB.Command.Execute", -1, receiver, true, true
	}
	if member == "open" && sqlObjectKindForState(a, receiver, state) == sqlObjectRecordset {
		return "ADODB.Recordset.Open", 0, receiver, false, true
	}
	if member == "runsql" || base == "runsql" {
		if strings.Contains(full, "docmd") || receiver == "docmd" || receiver == "" {
			return "DoCmd.RunSQL", 0, receiver, false, true
		}
	}
	if member == "openrecordset" || base == "openrecordset" {
		if receiver == "currentdb" || sqlObjectKindForState(a, receiver, state) == sqlObjectQueryDef {
			return "DAO.OpenRecordset", 0, receiver, false, true
		}
	}
	if member == "execute" || base == "execute" || member == "executesql" || base == "executesql" {
		kind := sqlObjectKindForState(a, receiver, state)
		if receiver == "currentdb" || kind == sqlObjectQueryDef || kind == sqlObjectCommand {
			return "Database.Execute", 0, receiver, false, true
		}
		if kind == sqlObjectConnection {
			return "ADODB.Connection.Execute", 0, receiver, false, true
		}
	}
	return "", -1, receiver, false, false
}

func (a *procedureAnalyzer) sqlObjectKind(name string) sqlObjectKind {
	name = canonicalName(name)
	if name == "currentdb" {
		return sqlObjectQueryDef
	}
	for _, declaration := range a.procedure.Declarations {
		if canonicalName(declaration.Name) != name {
			continue
		}
		typ := strings.ToLower(strings.TrimSpace(declaration.Type))
		switch {
		case strings.Contains(typ, "command"):
			return sqlObjectCommand
		case strings.Contains(typ, "recordset"):
			return sqlObjectRecordset
		case strings.Contains(typ, "connection"):
			return sqlObjectConnection
		case strings.Contains(typ, "database"), strings.Contains(typ, "querydef"):
			return sqlObjectQueryDef
		}
	}
	return sqlObjectUnknown
}

func sqlObjectKindForState(a *procedureAnalyzer, name string, state abstractState) sqlObjectKind {
	if object, ok := state.sqlObjects[canonicalName(name)]; ok {
		if object.kind != sqlObjectUnknown {
			return object.kind
		}
	}
	return a.sqlObjectKind(name)
}

func sqlObjectKindFromText(text string) sqlObjectKind {
	lower := strings.ToLower(strings.TrimSpace(text))
	switch {
	case strings.Contains(lower, "adodb.command"), strings.Contains(lower, "createcommand"):
		return sqlObjectCommand
	case strings.Contains(lower, "adodb.recordset"), strings.Contains(lower, "openrecordset"):
		return sqlObjectRecordset
	case strings.Contains(lower, "querydef"):
		return sqlObjectQueryDef
	case strings.Contains(lower, "adodb.connection"), strings.Contains(lower, "currentdb"), strings.Contains(lower, "database"):
		if strings.Contains(lower, "adodb.connection") {
			return sqlObjectConnection
		}
		return sqlObjectQueryDef
	default:
		return sqlObjectUnknown
	}
}

func sqlMemberTarget(text string) (string, string) {
	text = strings.TrimSpace(text)
	index := strings.LastIndex(text, ".")
	if index < 0 {
		return "", ""
	}
	return canonicalName(text[:index]), canonicalName(text[index+1:])
}

func classifySQLRole(text string, path []PathStep, sourceLabel string) SQLInputRole {
	lowerText := strings.ToLower(strings.TrimSpace(text))
	lowerSource := strings.ToLower(strings.TrimSpace(sourceLabel))
	if lowerSource != "" {
		for index := len(path) - 1; index >= 0; index-- {
			if path[index].Kind != "concatenation" {
				continue
			}
			segment := strings.ToLower(path[index].Label)
			labels := []string{lowerSource}
			for sourceIndex := index - 1; sourceIndex >= 0; sourceIndex-- {
				if path[sourceIndex].Kind == "source" {
					label := strings.ToLower(strings.TrimSpace(path[sourceIndex].Label))
					if label != "" && label != lowerSource {
						labels = append(labels, label)
					}
					break
				}
				if path[sourceIndex].Kind == "assignment" {
					if equals := strings.Index(path[sourceIndex].Label, "="); equals > 0 {
						label := strings.ToLower(strings.TrimSpace(path[sourceIndex].Label[:equals]))
						if label != "" && label != lowerSource {
							labels = append(labels, label)
						}
					}
				}
			}
			// Prefer the last occurrence: short source names such as `id` or
			// `name` often also appear in constant SQL column names before the
			// concatenated variable token.
			position := -1
			for _, label := range labels {
				if candidate := strings.LastIndex(segment, label); candidate > position {
					position = candidate
				}
			}
			if position < 0 {
				continue
			}
			prefix := strings.TrimRight(strings.TrimSpace(segment[:position]), "&'\" ")
			if strings.Contains(prefix, " like ") {
				return SQLInputRoleLikePattern
			}
			for _, marker := range []string{" from ", " join ", " into ", " table ", " order by ", " group by ", " select ", " update ", " set "} {
				if strings.HasSuffix(prefix, strings.TrimSpace(marker)) {
					return SQLInputRoleIdentifier
				}
			}
			return SQLInputRoleValue
		}
	}
	if strings.Contains(lowerText, " like ") && (strings.Contains(lowerText, "%") || strings.Contains(lowerText, "*")) {
		return SQLInputRoleLikePattern
	}
	if lowerText == "" {
		return SQLInputRoleUnknown
	}
	return SQLInputRoleValue
}

func classifySQLRisk(a *procedureAnalyzer, origin provenance, role SQLInputRole, text string) SQLRiskKind {
	lower := strings.ToLower(text)
	for _, step := range origin.path {
		if step.Kind == "transformation" || step.Kind == "assignment" || step.Kind == "concatenation" {
			lower += " " + strings.ToLower(step.Label)
		}
	}
	if role == SQLInputRoleIdentifier {
		return SQLRiskDynamicIdentifier
	}
	if role == SQLInputRoleLikePattern || (strings.Contains(lower, "like") && (strings.Contains(lower, "%") || strings.Contains(lower, "*"))) {
		return SQLRiskWildcardLikeInput
	}
	if sqlLocaleSensitive(a, origin, lower) {
		return SQLRiskLocaleSensitiveValue
	}
	if strings.Contains(lower, "'") || strings.Contains(lower, "#") || strings.Contains(lower, "replace(") {
		return SQLRiskManualQuoting
	}
	if origin.state == StateUnknown {
		return SQLRiskUnknownOrigin
	}
	return SQLRiskExternalValueConcatenation
}

func sqlLocaleSensitive(a *procedureAnalyzer, origin provenance, text string) bool {
	if strings.Contains(text, "format(") || strings.Contains(text, "format$(") || strings.Contains(text, "cstr(") || strings.Contains(text, "cdbl(") || strings.Contains(text, "cdate(") {
		return true
	}
	label := canonicalName(origin.source.Label)
	for _, parameter := range a.procedure.Symbol.Parameters {
		if canonicalName(parameter.Name) != label {
			continue
		}
		return localeSensitiveTypeName(parameter.Type)
	}
	for _, step := range origin.path {
		if step.Kind != "assignment" {
			continue
		}
		target := step.Label
		if equals := strings.Index(target, "="); equals >= 0 {
			target = target[:equals]
		}
		target = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(target), "Set "))
		assignmentLabel := canonicalName(target)
		for _, declaration := range a.procedure.Declarations {
			if canonicalName(declaration.Name) != assignmentLabel {
				continue
			}
			return localeSensitiveTypeName(declaration.Type)
		}
	}
	for _, declaration := range a.procedure.Declarations {
		if canonicalName(declaration.Name) != label {
			continue
		}
		return localeSensitiveTypeName(declaration.Type)
	}
	return false
}

func localeSensitiveTypeName(typ string) bool {
	typ = strings.ToLower(strings.TrimSpace(typ))
	for _, marker := range []string{"date", "double", "single", "currency", "decimal"} {
		if strings.Contains(typ, marker) {
			return true
		}
	}
	return false
}

func (a *procedureAnalyzer) applySourceWrite(call procedureir.CallSite, state *abstractState) {
	if state == nil {
		return
	}
	kind, label, ok := sourceCall(call)
	if !ok || kind != SourceFileInput || len(call.Arguments.ExpressionIDs) == 0 {
		return
	}
	// VBA's Input/Line Input/Get statements write their last argument(s),
	// whereas function-style reads are handled by evalCall on the assignment.
	for _, expressionID := range call.Arguments.ExpressionIDs[1:] {
		if a.contextErr() != nil {
			return
		}
		expression, exists := a.expressions[expressionID]
		if !exists || expression.Kind != procedureir.ExpressionIdentifier {
			continue
		}
		state.ensureVars()
		state.vars[canonicalName(expression.Text)] = valueFromSource(Source{
			Kind: kind, Label: label, Range: call.Range, StatementID: call.StatementID,
		}, PathStep{Kind: "source", Label: call.Callee.Text, Range: call.Range, StatementID: call.StatementID})
	}
}

func (a *procedureAnalyzer) evalExpression(expression procedureir.Expression, state abstractState) value {
	if a.contextErr() != nil {
		return value{}
	}
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
		name := canonicalName(expression.Text)
		if variable, ok := state.vars[name]; ok {
			return cloneValue(variable)
		}
		if a.constantNames[name] {
			return value{}
		}
		if a.declaredNames[name] {
			return a.unknownValue(expression.Range, expression.StatementID, "unassigned declared value")
		}
		if a.options.IsKnownConstant != nil && a.options.IsKnownConstant(expression.Text) {
			return value{}
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
		if a.contextErr() != nil {
			return value{}
		}
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
	// Chr/Chr$ are deterministic formatting helpers.  A literal character
	// code (the common Windows quoting idiom Chr$(34)) contributes no unknown
	// provenance; a dynamic character code still carries any origins from its
	// argument through the transformation below.
	if name == "chr" || name == "chr$" {
		result := joinValues(args)
		if !isEmptyValue(result) {
			return appendPath(result, PathStep{Kind: "transformation", Label: call.Callee.Text, Range: call.Range, StatementID: call.StatementID})
		}
		return result
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
		if a.contextErr() != nil {
			return value{}
		}
		if expression, ok := a.expressions[id]; ok {
			values = append(values, a.evalExpression(expression, state))
		}
	}
	return joinValues(values)
}

func (a *procedureAnalyzer) callArguments(call procedureir.CallSite, state abstractState) []value {
	args := make([]value, 0, len(call.Arguments.ExpressionIDs))
	for _, id := range call.Arguments.ExpressionIDs {
		if a.contextErr() != nil {
			return args
		}
		if expression, ok := a.expressions[id]; ok {
			args = append(args, a.evalExpression(expression, state))
		}
	}
	return args
}

func (a *procedureAnalyzer) inspectSink(call procedureir.CallSite, state abstractState) {
	targets := a.sinkTargets(call, state)
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
			if a.contextErr() != nil {
				return
			}
			keys = append(keys, key)
		}
		sort.Strings(keys)
		for _, key := range keys {
			if a.contextErr() != nil {
				return
			}
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

// inspectCommand records process-launch-specific observations. It runs only
// after the CFG fixed point has converged, so aliases and branch guards use
// the same state as the generic source-to-sink analysis.
func (a *procedureAnalyzer) inspectCommand(call procedureir.CallSite, state abstractState) {
	targets := commandTargets(call, a.knownShellObjects)
	if len(targets) == 0 {
		return
	}
	for _, target := range targets {
		// Window-style and wait flags are retained in the classifier for role
		// awareness, but they do not make command text unsafe by themselves.
		// Hidden/unobserved execution is handled separately below.
		if target.role == CommandRoleWindowStyle || target.role == CommandRoleWait {
			continue
		}
		if target.argument < 0 || target.argument >= len(call.Arguments.ExpressionIDs) {
			continue
		}
		expression, ok := a.expressions[call.Arguments.ExpressionIDs[target.argument]]
		if !ok {
			continue
		}
		target.execution.Interpreter = commandInterpreter(expression.Text)
		// VBA Shell accepts one combined command string.  When the expression is
		// a concatenation without an explicit interpreter, dynamic fragments are
		// ordinary arguments rather than executable text.  Keep the executable
		// role for a direct value (for example Shell raw), and let the interpreter
		// handling below upgrade cmd/PowerShell/script-host command strings.
		if target.kind == SinkShell && target.role == CommandRoleExecutable && expression.Kind == procedureir.ExpressionBinary {
			target.role = CommandRoleArguments
			target.execution.Role = target.role
		}
		if target.kind == SinkShell && target.role == CommandRoleArguments && strings.Contains(strings.ToLower(expression.Text), "/c") {
			target.role = CommandRoleShellCommand
			target.execution.Role = target.role
		}
		if strings.HasPrefix(strings.ToLower(vbaStringContent(expression.Text)), "http") && (target.role == CommandRoleDocument || target.role == CommandRoleExecutable) {
			target.role = CommandRoleURL
			target.execution.Role = target.role
		}
		if target.execution.Interpreter != "" && target.role == CommandRoleExecutable {
			switch target.execution.Interpreter {
			case "cmd.exe":
				target.role = CommandRoleShellCommand
			case "powershell":
				if powershellUsesCommand(expression.Text) {
					target.role = CommandRoleShellCommand
				} else {
					target.role = CommandRoleArguments
				}
			case "script_host":
				// Script hosts receive a script path plus interpreter options in
				// one command string; without a structured argument list, keep
				// the conservative command-text role.
				target.role = CommandRoleShellCommand
			}
			target.execution.Role = target.role
		}
		value := a.evalExpression(expression, state)
		origins := make([]string, 0, len(value.origins))
		for key := range value.origins {
			origins = append(origins, key)
		}
		sort.Strings(origins)
		for _, key := range origins {
			origin := value.origins[key]
			if origin.safe[target.kind] || (origin.state != StateTainted && origin.state != StateUnknown) {
				continue
			}
			role := target.role
			if target.kind == SinkShell && role == CommandRoleExecutable && nonStringParameter(a.procedure, origin.source) {
				// Numeric and Boolean parameters cannot inject shell metacharacters
				// into an executable path.  They remain a process-launch observation
				// because the value still changes the launched command's arguments.
				role = CommandRoleArguments
			}
			execution := target.execution
			execution.Role = role
			class := CommandRiskProcessLaunch
			kind := CommandRiskUnknownOrigin
			if origin.state == StateTainted && (role == CommandRoleExecutable || role == CommandRoleShellCommand) {
				class = CommandRiskInjection
				kind = CommandRiskTaintedCommandText
			}
			a.recordCommand(CommandFinding{
				State: origin.state, Source: origin.source, Execution: execution,
				RiskClass: class, RiskKind: kind, OriginState: origin.state,
				Path:    append([]PathStep(nil), origin.path...),
				Message: fmt.Sprintf("Command input from %s reaches %s.", origin.source.Label, target.execution.Launcher),
				Reason:  "External or unknown input reaches a process-launch argument; the origin and command role must be validated before execution.",
			})
		}
		if len(origins) > 0 && looksCredentialArgument(expression.Text, target.role) {
			origin := value.origins[origins[0]]
			a.recordCommand(CommandFinding{
				State: origin.state, Source: origin.source, Execution: target.execution,
				RiskClass: CommandRiskProcessLaunch, RiskKind: CommandRiskCredentialExposure,
				OriginState: origin.state, Path: append([]PathStep(nil), origin.path...),
				Message: "A credential-like value reaches a process command or argument.",
				Reason:  "Secrets passed on a command line may be exposed through process listings, logs, or child-process diagnostics.",
			})
		}
		if len(value.origins) == 0 {
			if looksUnquotedExecutable(expression.Text, target.role) {
				a.recordCommand(staticCommandFinding(target, CommandRiskUnquotedExecutablePath, "The executable path contains spaces without an outer pair of quotes."))
			}
			if looksCredentialArgument(expression.Text, target.role) {
				a.recordCommand(staticCommandFinding(target, CommandRiskCredentialExposure, "A credential-like value appears in a process command or argument."))
			}
		}
	}
	if commandHidden(call, a.expressions) && commandResultDiscarded(call, a.statements) {
		a.recordCommand(staticCommandFinding(targets[0], CommandRiskObservability, "The launch is hidden or unobserved and its process result is discarded."))
	}
}

type commandTarget struct {
	execution CommandExecution
	kind      SinkKind
	argument  int
	role      CommandRole
}

func commandTargets(call procedureir.CallSite, knownShellObjects map[string]string) []commandTarget {
	base := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	receiver := ""
	if call.Callee.Receiver != nil {
		receiver = strings.ToLower(strings.TrimSpace(*call.Callee.Receiver))
	}
	var out []commandTarget
	add := func(kind SinkKind, label string, index int, role CommandRole) {
		if index >= 0 && index < call.Arguments.Count {
			out = append(out, commandTarget{kind: kind, argument: index, role: role, execution: CommandExecution{Launcher: label, Role: role, Argument: index, Range: call.Range}})
		}
	}
	if base == "shell" && receiver == "" && !userDefinedLauncher(call) {
		add(SinkShell, "Shell", 0, CommandRoleExecutable)
		add(SinkShell, "Shell", 1, CommandRoleWindowStyle)
	}
	if (member == "run" || member == "exec") && isWScriptShellReceiver(receiver, full, knownShellObjects) {
		kind, label := SinkWScriptShellRun, "WScript.Shell.Run"
		if member == "exec" {
			kind, label = SinkWScriptShellExec, "WScript.Shell.Exec"
		}
		add(kind, label, 0, CommandRoleExecutable)
		if member == "run" {
			add(kind, label, 1, CommandRoleWindowStyle)
			add(kind, label, 2, CommandRoleWait)
		}
	}
	if isShellExecuteCall(base, member, full, receiver, knownShellObjects) && !userDefinedLauncher(call) {
		primary := 2
		label := shellExecuteLabel(call, knownShellObjects)
		if strings.Contains(full, "shell.application") || knownShellObjects[receiver] == "shell.application" {
			primary = 0
		}
		role := CommandRoleDocument
		if primary == 2 {
			role = CommandRoleExecutable
		}
		add(SinkShellExecute, label, primary, role)
		if primary+1 < call.Arguments.Count {
			add(SinkShellExecute, label, primary+1, CommandRoleArguments)
		}
		show := 4
		if primary == 2 {
			show = 5
		}
		add(SinkShellExecute, label, show, CommandRoleWindowStyle)
	}
	return out
}

func staticCommandFinding(target commandTarget, kind CommandRiskKind, reason string) CommandFinding {
	return CommandFinding{State: StateClean, Execution: target.execution, RiskClass: CommandRiskProcessLaunch, RiskKind: kind, OriginState: StateClean, Reason: reason, Message: reason}
}

func isShellExecuteCall(base, member, full, receiver string, knownShellObjects map[string]string) bool {
	switch base {
	case "shellexecute", "shellexecutea", "shellexecutew":
		return true
	}
	if member != "shellexecute" {
		return false
	}
	return strings.Contains(full, "shell.application") || knownShellObjects[receiver] == "shell.application"
}

func isWScriptShellReceiver(receiver, full string, knownShellObjects map[string]string) bool {
	if knownShellObjects[receiver] == "wscript.shell" || strings.Contains(full, "wscript.shell") {
		return true
	}
	return strings.Contains(receiver, "createobject") &&
		(strings.Contains(receiver, "wscript") || strings.Contains(receiver, "wshell"))
}

func userDefinedLauncher(call procedureir.CallSite) bool {
	if call.Resolution.Status != procedureir.ResolutionMatched && call.Resolution.Status != procedureir.ResolutionAmbiguous {
		return false
	}
	if len(call.Resolution.Candidates) == 0 {
		return true
	}
	for _, candidate := range call.Resolution.Candidates {
		kind := strings.ToLower(strings.TrimSpace(candidate.Kind))
		if !strings.HasPrefix(kind, "declare") {
			return true
		}
	}
	return false
}

func (a *procedureAnalyzer) recordCommand(finding CommandFinding) {
	key := fmt.Sprintf("%d\x00%s\x00%d\x00%d", finding.Execution.Range.StartByte, finding.RiskKind, finding.Source.Range.StartByte, finding.Execution.Argument)
	if previous, ok := a.commandFindings[key]; ok {
		if betterPath(finding.Path, previous.Path) {
			previous.Path = append([]PathStep(nil), finding.Path...)
		}
		if previous.State == StateUnknown && finding.State == StateTainted {
			previous.State = finding.State
			previous.OriginState = finding.OriginState
		}
		a.commandFindings[key] = previous
		return
	}
	a.commandFindings[key] = finding
}

func commandInterpreter(text string) string {
	lower := strings.ToLower(strings.TrimSpace(text))
	switch {
	case cmdInterpreterPattern.MatchString(lower) && cmdSwitchPattern.MatchString(lower):
		return "cmd.exe"
	case powershellInterpreterPattern.MatchString(lower):
		return "powershell"
	case scriptHostInterpreterPattern.MatchString(lower):
		return "script_host"
	default:
		return ""
	}
}

func nonStringParameter(procedure procedureir.ProcedureIR, source Source) bool {
	if source.Kind != SourceParameter {
		return false
	}
	wanted := canonicalName(source.Label)
	for _, parameter := range procedure.Symbol.Parameters {
		if canonicalName(parameter.Name) != wanted {
			continue
		}
		typ := strings.ToLower(strings.TrimSpace(parameter.Type))
		if typ == "" || typ == "string" || strings.Contains(typ, "string") {
			return false
		}
		for _, numeric := range []string{"byte", "integer", "long", "single", "double", "currency", "decimal", "date", "boolean"} {
			if strings.Contains(typ, numeric) {
				return true
			}
		}
		return false
	}
	return false
}

func powershellUsesCommand(text string) bool {
	return powershellCommandPattern.MatchString(strings.TrimSpace(text))
}

func looksUnquotedExecutable(text string, role CommandRole) bool {
	if role != CommandRoleExecutable && role != CommandRoleShellCommand {
		return false
	}
	trimmed := vbaStringContent(text)
	if strings.HasPrefix(trimmed, `"`) {
		return false
	}
	lower := strings.ToLower(trimmed)
	for _, extension := range []string{".exe", ".com", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".wsf", ".hta"} {
		if index := strings.Index(lower, extension); index >= 0 && strings.Contains(trimmed[:index], " ") {
			return true
		}
	}
	return false
}

func vbaStringContent(text string) string {
	trimmed := strings.TrimSpace(text)
	if len(trimmed) >= 2 && strings.HasPrefix(trimmed, `"`) && strings.HasSuffix(trimmed, `"`) {
		trimmed = trimmed[1 : len(trimmed)-1]
		trimmed = strings.ReplaceAll(trimmed, `""`, `"`)
	}
	return trimmed
}

func looksCredentialArgument(text string, role CommandRole) bool {
	if role != CommandRoleArguments && role != CommandRoleExecutable && role != CommandRoleShellCommand {
		return false
	}
	return credentialAssignmentPattern.MatchString(text) ||
		credentialWordPattern.MatchString(strings.TrimSpace(text))
}

func commandHidden(call procedureir.CallSite, expressions map[int]procedureir.Expression) bool {
	base := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	if base != "shell" && member != "run" {
		return false
	}
	if len(call.Arguments.ExpressionIDs) < 2 {
		return false
	}
	expression, ok := expressions[call.Arguments.ExpressionIDs[1]]
	if !ok {
		return false
	}
	value := strings.ToLower(strings.TrimSpace(expression.Text))
	return value == "0" || value == "vbhide" || value == "vb_hide"
}

func commandResultDiscarded(call procedureir.CallSite, statements map[int]procedureir.Statement) bool {
	statement, ok := statements[call.StatementID]
	if !ok {
		return true
	}
	return assignmentTarget(statement) == ""
}

type sinkTarget struct {
	kind     SinkKind
	label    string
	argument int
}

func (a *procedureAnalyzer) sinkTargets(call procedureir.CallSite, state abstractState) []sinkTarget {
	knownShellObjects := a.knownShellObjects
	base := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	receiver := ""
	if call.Callee.Receiver != nil {
		receiver = strings.ToLower(strings.TrimSpace(*call.Callee.Receiver))
	}
	var out []sinkTarget
	if base == "shell" && receiver == "" && !userDefinedLauncher(call) {
		out = append(out, sinkTarget{SinkShell, "Shell", 0})
	}
	if (member == "run" || member == "exec") && isWScriptShellReceiver(receiver, full, knownShellObjects) {
		kind, label := SinkWScriptShellRun, "WScript.Shell.Run"
		if member == "exec" {
			kind, label = SinkWScriptShellExec, "WScript.Shell.Exec"
		}
		out = append(out, sinkTarget{kind, label, 0})
	}
	if isShellExecuteCall(base, member, full, receiver, knownShellObjects) && !userDefinedLauncher(call) {
		out = append(out, sinkTarget{SinkShellExecute, "ShellExecute", shellExecutePrimaryArgument(call, knownShellObjects)})
	}
	if member == "execute" || base == "execute" || member == "executenonquery" || base == "executenonquery" || member == "openrecordset" || base == "openrecordset" || member == "runsql" || base == "runsql" || member == "executesql" || base == "executesql" {
		// ADODB.Command.Execute receives recordsAffected/parameters/options;
		// its SQL text is stored in CommandText and must not be inferred from
		// the first output argument by the generic VBA224 fallback.
		if looksDatabaseReceiver(receiver, full) && !a.looksADODBCommandReceiver(receiver, full, state) {
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

func shellExecutePrimaryArgument(call procedureir.CallSite, knownShellObjects map[string]string) int {
	// Shell.Application.ShellExecute(File, Args, ...) starts at argument 0;
	// the Win32 ShellExecute signature has hWnd, operation, file, parameters,
	// directory, show, so the file is argument 2.
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	receiver := ""
	if call.Callee.Receiver != nil {
		receiver = strings.ToLower(strings.TrimSpace(*call.Callee.Receiver))
	}
	if strings.Contains(full, "shell.application") || knownShellObjects[receiver] == "shell.application" {
		return 0
	}
	return 2
}

func shellExecuteLabel(call procedureir.CallSite, knownShellObjects map[string]string) string {
	full := strings.ToLower(strings.TrimSpace(call.Callee.Text))
	receiver := ""
	if call.Callee.Receiver != nil {
		receiver = strings.ToLower(strings.TrimSpace(*call.Callee.Receiver))
	}
	if strings.Contains(full, "shell.application") || knownShellObjects[receiver] == "shell.application" {
		return "Shell.Application.ShellExecute"
	}
	base := strings.TrimSpace(call.Callee.BaseName)
	if base != "" {
		return base
	}
	return "ShellExecute"
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

func (a *procedureAnalyzer) looksADODBCommandReceiver(receiver, full string, state abstractState) bool {
	if sqlObjectKindFromText(receiver) == sqlObjectCommand || sqlObjectKindFromText(full) == sqlObjectCommand {
		return true
	}
	return sqlObjectKindForState(a, receiver, state) == sqlObjectCommand
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
	for _, statement := range a.procedure.Statements {
		if a.contextErr() != nil {
			return
		}
		if match := knownShellObjectPattern.FindStringSubmatch(statement.Text); len(match) > 1 {
			kind := ""
			if len(match) > 2 {
				kind = strings.ToLower(strings.TrimSpace(match[2]))
			}
			if kind == "" && len(match) > 3 {
				kind = strings.ToLower(strings.TrimSpace(match[3]))
			}
			if kind != "" {
				a.knownShellObjects[canonicalName(match[1])] = kind
			}
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

func (a *procedureAnalyzer) applyGuard(from cfg.BlockID, edge cfg.Edge, state *abstractState) {
	if state == nil {
		return
	}
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
			validated = cloneValue(validated)
			markSafe(&validated)
			state.ensureVars()
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
			validated = cloneValue(validated)
			markSafe(&validated)
			state.ensureVars()
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
	rank, _ := a.reversePostOrderRankContext(reachable)
	return rank
}

func (a *procedureAnalyzer) reversePostOrderRankContext(reachable map[cfg.BlockID]bool) (map[cfg.BlockID]int, error) {
	visited := map[cfg.BlockID]bool{}
	postorder := make([]cfg.BlockID, 0, len(reachable))
	type frame struct {
		id   cfg.BlockID
		next int
	}
	visited[a.graph.Entry] = true
	stack := []frame{{id: a.graph.Entry}}
	for len(stack) > 0 {
		if err := a.contextErr(); err != nil {
			return nil, err
		}
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
		if index&0xff == 0 {
			if err := a.contextErr(); err != nil {
				return nil, err
			}
		}
		rank[postorder[index]] = len(postorder) - 1 - index
	}
	return rank, nil
}

func (a *procedureAnalyzer) reachableContext() (map[cfg.BlockID]bool, error) {
	seen := map[cfg.BlockID]bool{}
	if err := a.contextErr(); err != nil {
		return nil, err
	}
	seen[a.graph.Entry] = true
	queue := []cfg.BlockID{a.graph.Entry}
	for len(queue) > 0 {
		if err := a.contextErr(); err != nil {
			return nil, err
		}
		current := queue[0]
		queue = queue[1:]
		for _, edge := range a.outgoing(current) {
			if err := a.contextErr(); err != nil {
				return nil, err
			}
			if seen[edge.To] {
				continue
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
		}
	}
	unknownReached := false
	for _, source := range a.graph.UnknownFlowSources {
		if err := a.contextErr(); err != nil {
			return nil, err
		}
		if seen[source] {
			unknownReached = true
			break
		}
	}
	if unknownReached {
		for _, block := range a.graph.Blocks {
			if err := a.contextErr(); err != nil {
				return nil, err
			}
			if block.Kind == cfg.BlockStatement {
				seen[block.ID] = true
			}
		}
		seen[a.graph.UnknownExit] = true
		queue = queue[:0]
		for id := range seen {
			queue = append(queue, id)
		}
		for len(queue) > 0 {
			if err := a.contextErr(); err != nil {
				return nil, err
			}
			current := queue[0]
			queue = queue[1:]
			for _, edge := range a.outgoing(current) {
				if err := a.contextErr(); err != nil {
					return nil, err
				}
				if seen[edge.To] {
					continue
				}
				seen[edge.To] = true
				queue = append(queue, edge.To)
			}
		}
	}
	reachable := make(map[cfg.BlockID]bool, len(seen))
	for _, block := range a.graph.Blocks {
		if seen[block.ID] {
			reachable[block.ID] = true
		}
	}
	return reachable, nil
}

func (a *procedureAnalyzer) contextErr() error {
	if a == nil || a.ctx == nil {
		return nil
	}
	return a.ctx.Err()
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
		return cloneState(&b), true
	}
	result := cloneState(&a)
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
			result.ensureVars()
			result.vars[key] = merged
			changed = true
		}
	}
	sqlKeys := map[string]bool{}
	for key := range a.sqlObjects {
		sqlKeys[key] = true
	}
	for key := range b.sqlObjects {
		sqlKeys[key] = true
	}
	for key := range sqlKeys {
		left, leftOK := a.sqlObjects[key]
		right, rightOK := b.sqlObjects[key]
		if !leftOK {
			result.ensureSQLObjects()
			result.sqlObjects[key] = cloneSQLObjectState(right)
			changed = true
			continue
		}
		if !rightOK {
			continue
		}
		merged := joinSQLObjectState(left, right)
		if !sameSQLObjectState(left, merged) {
			result.ensureSQLObjects()
			result.sqlObjects[key] = merged
			changed = true
		}
	}
	return result, changed
}

func unknownStandaloneValue() value {
	return valueFromSourceState(Source{Kind: SourceUnknown, Label: "unknown input"}, StateUnknown, PathStep{Kind: "unknown_transformation", Label: "possibly unassigned value"})
}

func cloneState(state *abstractState) abstractState {
	if state == nil {
		return abstractState{vars: map[string]value{}, sqlObjects: map[string]sqlObjectState{}}
	}
	state.varsShared = true
	state.sqlObjectsShared = true
	result := abstractState{
		vars:             state.vars,
		sqlObjects:       state.sqlObjects,
		varsShared:       true,
		sqlObjectsShared: true,
	}
	if result.vars == nil {
		result.vars = map[string]value{}
		result.varsShared = false
	}
	if result.sqlObjects == nil {
		result.sqlObjects = map[string]sqlObjectState{}
		result.sqlObjectsShared = false
	}
	return result
}

func (s *abstractState) ensureVars() {
	if s == nil || !s.varsShared {
		return
	}
	out := make(map[string]value, len(s.vars))
	for key, variable := range s.vars {
		out[key] = variable
	}
	s.vars = out
	s.varsShared = false
}

func (s *abstractState) ensureSQLObjects() {
	if s == nil || !s.sqlObjectsShared {
		return
	}
	out := make(map[string]sqlObjectState, len(s.sqlObjects))
	for key, object := range s.sqlObjects {
		out[key] = object
	}
	s.sqlObjects = out
	s.sqlObjectsShared = false
}

func cloneSQLObjectState(input sqlObjectState) sqlObjectState {
	result := input
	if input.parameters != nil {
		result.parameters = make(map[string]value, len(input.parameters))
		for key, parameter := range input.parameters {
			result.parameters[key] = parameter
		}
	}
	return result
}

func joinSQLObjectState(a, b sqlObjectState) sqlObjectState {
	result := cloneSQLObjectState(a)
	if result.identity == "" || b.identity == "" || result.identity != b.identity {
		result.identity = ""
	}
	if result.kind == sqlObjectUnknown || b.kind == sqlObjectUnknown || result.kind != b.kind {
		result.kind = sqlObjectUnknown
	}
	result.commandText, _ = joinValue(result.commandText, b.commandText, true)
	result.parameterized = a.parameterized && b.parameterized
	if result.parameters == nil {
		result.parameters = map[string]value{}
	}
	for key, parameter := range b.parameters {
		if previous, ok := result.parameters[key]; ok {
			result.parameters[key], _ = joinValue(previous, parameter, true)
		} else {
			result.parameters[key] = cloneValue(parameter)
		}
	}
	return result
}

func sameSQLObjectState(a, b sqlObjectState) bool {
	if a.kind != b.kind || a.identity != b.identity || a.parameterized != b.parameterized || !isSameValue(a.commandText, b.commandText) || len(a.parameters) != len(b.parameters) {
		return false
	}
	for key, value := range a.parameters {
		if other, ok := b.parameters[key]; !ok || !isSameValue(value, other) {
			return false
		}
	}
	return true
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
	return a.state == b.state && sourceKey(a.source) == sourceKey(b.source) && comparePathKeys(a.path, b.path) == 0 && sameSafe(a.safe, b.safe)
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
	return comparePathKeys(a, b) < 0
}

func betterPathValue(a, b []PathStep) []PathStep {
	if betterPath(b, a) {
		return append([]PathStep(nil), b...)
	}
	return append([]PathStep(nil), a...)
}
