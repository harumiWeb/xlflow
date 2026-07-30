package cfg

import (
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// BuildDocument builds one independent graph per procedure in source order.
func BuildDocument(in procedureir.DocumentIR) Document {
	out := Document{Path: in.Path, Graphs: make([]Graph, len(in.Procedures))}
	for i := range in.Procedures {
		out.Graphs[i] = Build(in.Procedures[i])
	}
	return out
}

type loopFrame struct {
	kind procedureir.StatementKind
	exit BlockID
}

type builder struct {
	procedure procedureir.ProcedureIR
	graph     Graph
	blockByID map[int]BlockID
	stmtByID  map[int]procedureir.Statement
	children  map[int][]int
	labels    map[string][]BlockID
	edges     []Edge
}

// Build constructs a deterministic conservative graph for one procedure.
func Build(in procedureir.ProcedureIR) Graph {
	in = procedureir.CloneProcedureIR(in)
	b := &builder{
		procedure: in,
		blockByID: map[int]BlockID{},
		stmtByID:  map[int]procedureir.Statement{},
		children:  map[int][]int{},
		labels:    map[string][]BlockID{},
	}
	b.initialize()
	top := b.children[0]
	entry := b.graph.NormalExit
	if len(top) > 0 {
		entry = b.buildSequence(top, b.graph.NormalExit, nil)
	}
	b.add(b.graph.Entry, entry, EdgeFallthrough, EdgeNormal, false, nil)
	b.addUnknownRecoveryEdges()
	b.addExceptionalEdges()
	b.finish()
	return b.graph
}

func (b *builder) initialize() {
	symbol := b.procedure.Symbol
	symbol.Parameters = append([]procedureir.Parameter(nil), b.procedure.Symbol.Parameters...)
	b.graph = Graph{
		Procedure: symbol,
		Entry:     1, NormalExit: 2, ExceptionalExit: 3, TerminationExit: 4, UnknownExit: 5,
		Blocks: []Block{
			{ID: 1, Kind: BlockEntry},
			{ID: 2, Kind: BlockNormalExit},
			{ID: 3, Kind: BlockExceptionalExit},
			{ID: 4, Kind: BlockTerminationExit},
			{ID: 5, Kind: BlockUnknownExit},
		},
	}
	statements := append([]procedureir.Statement(nil), b.procedure.Statements...)
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartByte != statements[j].Range.StartByte {
			return statements[i].Range.StartByte < statements[j].Range.StartByte
		}
		return statements[i].ID < statements[j].ID
	})
	assignments := map[int][]Variable{}
	for _, access := range b.procedure.Accesses {
		if access.StatementID == 0 ||
			(access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) {
			continue
		}
		assignments[access.StatementID] = appendVariable(assignments[access.StatementID],
			Variable{Scope: access.Scope, Name: access.Name})
	}
	for i := range statements {
		statement := statements[i]
		id := BlockID(len(b.graph.Blocks) + 1)
		copyStatement := statement
		copyStatement.ExpressionIDs = append([]int(nil), statement.ExpressionIDs...)
		copyStatement.Target = cloneExpression(statement.Target)
		copyStatement.Value = cloneExpression(statement.Value)
		copyStatement.Condition = cloneExpression(statement.Condition)
		if statement.Control != nil {
			control := *statement.Control
			copyStatement.Control = &control
		}
		b.graph.Blocks = append(b.graph.Blocks, Block{
			ID: id, Kind: BlockStatement, StatementID: statement.ID,
			Statement: &copyStatement, Range: statement.Range,
			Assignments: assignments[statement.ID],
		})
		b.blockByID[statement.ID] = id
		b.stmtByID[statement.ID] = statement
		b.children[statement.ParentID] = append(b.children[statement.ParentID], statement.ID)
		if statement.Kind == procedureir.StatementLabel {
			name := normalizedTarget(statement.Label)
			if statement.Control != nil && statement.Control.Target != "" {
				name = normalizedTarget(statement.Control.Target)
			}
			if name != "" {
				b.labels[name] = append(b.labels[name], id)
			}
		}
	}
}

func appendVariable(in []Variable, variable Variable) []Variable {
	variable = variable.canonical()
	for _, existing := range in {
		if existing == variable {
			return in
		}
	}
	return append(in, variable)
}

func (b *builder) buildSequence(ids []int, continuation BlockID, loops []loopFrame) BlockID {
	next := continuation
	for i := len(ids) - 1; i >= 0; i-- {
		next = b.buildStatement(ids[i], next, loops)
	}
	return next
}

func (b *builder) buildStatement(statementID int, next BlockID, loops []loopFrame) BlockID {
	statement := b.stmtByID[statementID]
	block := b.blockByID[statementID]
	control := statement.Control
	if control != nil {
		switch control.Transfer {
		case procedureir.TransferGoto:
			b.addTarget(block, control.Target, EdgeGoto, EdgeNormal, false, statement)
			return block
		case procedureir.TransferExitSub, procedureir.TransferExitFunction, procedureir.TransferExitProperty:
			if transferMatchesProcedure(control.Transfer, b.procedure.Symbol.Kind) {
				b.add(block, b.graph.NormalExit, EdgeProcedureExit, EdgeNormal, false, &statement)
			} else {
				b.add(block, b.graph.UnknownExit, EdgeUnknown, EdgeNormal, true, &statement)
			}
			return block
		case procedureir.TransferExitFor:
			target := nearestLoopExit(loops, procedureir.StatementFor, procedureir.StatementForEach, b.graph.UnknownExit)
			b.add(block, target, EdgeLoopExit, EdgeNormal, target == b.graph.UnknownExit, &statement)
			return block
		case procedureir.TransferExitDo:
			target := nearestLoopExit(loops, procedureir.StatementDo, "", b.graph.UnknownExit)
			b.add(block, target, EdgeLoopExit, EdgeNormal, target == b.graph.UnknownExit, &statement)
			return block
		case procedureir.TransferResumeRetry, procedureir.TransferResumeNext:
			b.add(block, b.graph.UnknownExit, EdgeResume, EdgeExceptional, true, &statement)
			return block
		case procedureir.TransferResumeLabel:
			b.addTarget(block, control.Target, EdgeResume, EdgeExceptional, false, statement)
			return block
		case procedureir.TransferTerminate:
			b.add(block, b.graph.TerminationExit, EdgeTermination, EdgeNormal, false, &statement)
			return block
		}
	}
	switch statement.Kind {
	case procedureir.StatementIf, procedureir.StatementElseIf:
		b.buildIf(statement, next, loops)
	case procedureir.StatementSelect:
		b.buildSelect(statement, next, loops)
	case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementWhile, procedureir.StatementDo:
		b.buildLoop(statement, next, loops)
	default:
		children := b.children[statement.ID]
		dest := next
		if len(children) > 0 {
			dest = b.buildSequence(children, next, loops)
		}
		b.add(block, dest, EdgeFallthrough, EdgeNormal, false, &statement)
	}
	return block
}

func transferMatchesProcedure(transfer procedureir.TransferKind, kind procedureir.ProcedureKind) bool {
	switch transfer {
	case procedureir.TransferExitSub:
		return kind == procedureir.ProcedureSub
	case procedureir.TransferExitFunction:
		return kind == procedureir.ProcedureFunction
	case procedureir.TransferExitProperty:
		return kind == procedureir.ProcedureProperty || kind == procedureir.ProcedurePropertyGet ||
			kind == procedureir.ProcedurePropertyLet || kind == procedureir.ProcedurePropertySet
	default:
		return false
	}
}

func (b *builder) buildIf(statement procedureir.Statement, next BlockID, loops []loopFrame) {
	block := b.blockByID[statement.ID]
	children := b.children[statement.ID]
	var thenIDs, alternatives []int
	hasRoles := false
	for _, id := range children {
		child := b.stmtByID[id]
		if child.Control != nil && child.Control.Branch != "" {
			hasRoles = true
			if child.Control.Branch == procedureir.BranchElse {
				alternatives = append(alternatives, id)
			} else {
				thenIDs = append(thenIDs, id)
			}
		}
	}
	if !hasRoles {
		for _, id := range children {
			child := b.stmtByID[id]
			if child.Kind == procedureir.StatementElseIf || child.Kind == procedureir.StatementElse {
				alternatives = append(alternatives, id)
			} else if len(alternatives) == 0 {
				thenIDs = append(thenIDs, id)
			}
		}
	}
	thenEntry := next
	if len(thenIDs) > 0 {
		thenEntry = b.buildSequence(thenIDs, next, loops)
	}
	elseEntry := next
	if len(alternatives) > 0 {
		if hasRoles {
			elseEntry = b.buildSequence(alternatives, next, loops)
		} else {
			elseEntry = b.buildAlternativeChain(alternatives, next, next, loops)
		}
	}
	b.add(block, thenEntry, EdgeBranchTrue, EdgeNormal, false, &statement)
	b.add(block, elseEntry, EdgeBranchFalse, EdgeNormal, false, &statement)
}

func (b *builder) buildAlternativeChain(ids []int, falseFallback, join BlockID, loops []loopFrame) BlockID {
	nextAlternative := falseFallback
	for i := len(ids) - 1; i >= 0; i-- {
		statement := b.stmtByID[ids[i]]
		if statement.Kind == procedureir.StatementElse {
			nextAlternative = b.buildStatement(statement.ID, join, loops)
			continue
		}
		if statement.Kind != procedureir.StatementElseIf {
			nextAlternative = b.buildStatement(statement.ID, nextAlternative, loops)
			continue
		}
		nextAlternative = b.buildElseIf(statement, nextAlternative, join, loops)
	}
	return nextAlternative
}

func (b *builder) buildElseIf(statement procedureir.Statement, falseEntry, continuation BlockID, loops []loopFrame) BlockID {
	var body, nestedAlternatives []int
	for _, id := range b.children[statement.ID] {
		child := b.stmtByID[id]
		if child.Kind == procedureir.StatementElseIf || child.Kind == procedureir.StatementElse {
			nestedAlternatives = append(nestedAlternatives, id)
		} else if len(nestedAlternatives) == 0 {
			body = append(body, id)
		} else {
			nestedAlternatives = append(nestedAlternatives, id)
		}
	}
	if len(nestedAlternatives) > 0 {
		falseEntry = b.buildAlternativeChain(nestedAlternatives, falseEntry, continuation, loops)
	}
	trueEntry := continuation
	if len(body) > 0 {
		trueEntry = b.buildSequence(body, continuation, loops)
	}
	block := b.blockByID[statement.ID]
	b.add(block, trueEntry, EdgeBranchTrue, EdgeNormal, false, &statement)
	b.add(block, falseEntry, EdgeBranchFalse, EdgeNormal, false, &statement)
	return block
}

func (b *builder) buildSelect(statement procedureir.Statement, next BlockID, loops []loopFrame) {
	block := b.blockByID[statement.ID]
	cases := b.children[statement.ID]
	hasElse := false
	for _, id := range cases {
		candidate := b.stmtByID[id]
		if candidate.Kind != procedureir.StatementCase {
			continue
		}
		caseEntry := b.buildStatement(id, next, loops)
		b.add(block, caseEntry, EdgeCase, EdgeNormal, false, &statement)
		hasElse = hasElse || candidate.Control != nil && candidate.Control.CaseElse
	}
	if !hasElse {
		b.add(block, next, EdgeBranchFalse, EdgeNormal, false, &statement)
	}
}

func (b *builder) buildLoop(statement procedureir.Statement, next BlockID, loops []loopFrame) {
	block := b.blockByID[statement.ID]
	frame := loopFrame{kind: statement.Kind, exit: next}
	body := b.children[statement.ID]
	if statement.Kind == procedureir.StatementDo {
		var testID int
		filtered := make([]int, 0, len(body))
		for _, id := range body {
			if b.stmtByID[id].SyntaxKind == "do_condition" {
				testID = id
			} else {
				filtered = append(filtered, id)
			}
		}
		body = filtered
		if testID != 0 {
			testBlock := b.blockByID[testID]
			bodyEntry := testBlock
			if len(body) > 0 {
				bodyEntry = b.buildSequence(body, testBlock, append(loops, frame))
			}
			b.markLoopBackEdges(statement.ID, testBlock)
			postTest := statement.Control != nil &&
				(statement.Control.Loop == procedureir.LoopPostWhile || statement.Control.Loop == procedureir.LoopPostUntil)
			if postTest {
				b.add(block, bodyEntry, EdgeLoopBody, EdgeNormal, false, &statement)
				b.add(testBlock, bodyEntry, EdgeLoopBack, EdgeNormal, false, b.statementPointer(testID))
			} else {
				b.add(block, testBlock, EdgeFallthrough, EdgeNormal, false, &statement)
				b.add(testBlock, bodyEntry, EdgeLoopBody, EdgeNormal, false, b.statementPointer(testID))
				if len(body) == 0 {
					b.add(testBlock, testBlock, EdgeLoopBack, EdgeNormal, false, b.statementPointer(testID))
				}
			}
			b.add(testBlock, next, EdgeLoopExit, EdgeNormal, false, b.statementPointer(testID))
			return
		}
		bodyEntry := block
		if len(body) > 0 {
			bodyEntry = b.buildSequence(body, block, append(loops, frame))
		}
		b.add(block, bodyEntry, EdgeLoopBody, EdgeNormal, false, &statement)
		if len(body) == 0 {
			b.add(block, block, EdgeLoopBack, EdgeNormal, false, &statement)
		}
		return
	}
	bodyEntry := block
	if len(body) > 0 {
		bodyEntry = b.buildSequence(body, block, append(loops, frame))
	}
	b.markLoopBackEdges(statement.ID, block)
	b.add(block, bodyEntry, EdgeLoopBody, EdgeNormal, false, &statement)
	if len(body) == 0 {
		b.add(block, block, EdgeLoopBack, EdgeNormal, false, &statement)
	}
	b.add(block, next, EdgeLoopExit, EdgeNormal, false, &statement)
}

func (b *builder) markLoopBackEdges(loopStatementID int, target BlockID) {
	for i := range b.edges {
		edge := &b.edges[i]
		if edge.To == target && edge.Class == EdgeNormal &&
			edge.Kind != EdgeGoto && b.isDescendantStatement(edge.StatementID, loopStatementID) {
			edge.Kind = EdgeLoopBack
		}
	}
}

func (b *builder) isDescendantStatement(statementID, ancestorID int) bool {
	for statementID != 0 {
		statement := b.stmtByID[statementID]
		if statement.ParentID == ancestorID {
			return true
		}
		statementID = statement.ParentID
	}
	return false
}

func (b *builder) statementPointer(id int) *procedureir.Statement {
	statement, ok := b.stmtByID[id]
	if !ok {
		return nil
	}
	return &statement
}

func nearestLoopExit(loops []loopFrame, first, second procedureir.StatementKind, fallback BlockID) BlockID {
	for i := len(loops) - 1; i >= 0; i-- {
		if loops[i].kind == first || second != "" && loops[i].kind == second {
			return loops[i].exit
		}
	}
	return fallback
}

func (b *builder) addTarget(from BlockID, target string, kind EdgeKind, class EdgeClass, uncertain bool, statement procedureir.Statement) {
	candidates := b.labels[normalizedTarget(target)]
	if len(candidates) == 1 {
		targetBlock := b.graph.Blocks[int(candidates[0])-1]
		targetRecovered := targetBlock.Statement != nil && targetBlock.Statement.Recovered
		b.add(from, candidates[0], kind, class, uncertain || targetRecovered, &statement)
		if targetRecovered {
			b.add(from, b.graph.UnknownExit, EdgeUnknown, class, true, &statement)
		}
		return
	}
	if len(candidates) == 0 {
		b.add(from, b.graph.UnknownExit, EdgeUnknown, class, true, &statement)
		return
	}
	for _, candidate := range candidates {
		b.add(from, candidate, kind, class, true, &statement)
	}
	b.add(from, b.graph.UnknownExit, EdgeUnknown, class, true, &statement)
}

func normalizedTarget(target string) string {
	return strings.ToLower(strings.Trim(strings.TrimSpace(target), "[]:"))
}

func (b *builder) add(from, to BlockID, kind EdgeKind, class EdgeClass, uncertain bool, statement *procedureir.Statement) {
	edge := Edge{From: from, To: to, Kind: kind, Class: class, Uncertain: uncertain}
	if statement != nil {
		edge.StatementID = statement.ID
		edge.Range = statement.Range
		if statement.Control != nil && statement.Control.Range.EndByte > statement.Control.Range.StartByte {
			edge.Range = statement.Control.Range
		}
	}
	b.edges = append(b.edges, edge)
}

func (b *builder) addUnknownRecoveryEdges() {
	for _, statement := range b.procedure.Statements {
		if statement.Kind != procedureir.StatementUnknown && statement.Kind != procedureir.StatementRecovered &&
			!statement.Recovered {
			continue
		}
		from := b.blockByID[statement.ID]
		b.add(from, b.graph.UnknownExit, EdgeUnknown, EdgeNormal, true, &statement)
		for _, block := range b.graph.Blocks {
			if block.Kind == BlockStatement && block.ID != from {
				b.add(from, block.ID, EdgeUnknown, EdgeNormal, true, &statement)
			}
		}
	}
}

type errorMode string

const (
	errorDisabled errorMode = "disabled"
	errorNext     errorMode = "next"
)

func (b *builder) addExceptionalEdges() {
	modes := map[BlockID]map[errorMode]bool{b.graph.Entry: {errorDisabled: true}}
	changed := true
	for changed {
		changed = false
		for _, edge := range b.edges {
			if edge.Class != EdgeNormal {
				continue
			}
			in := modes[edge.From]
			if len(in) == 0 {
				continue
			}
			out := b.transferErrorMode(edge.From, in)
			if mergeModes(modes, edge.To, out) {
				changed = true
			}
		}
	}
	baseEdges := append([]Edge(nil), b.edges...)
	for _, block := range b.graph.Blocks {
		if !b.isFaultSite(block) {
			continue
		}
		for mode := range modes[block.ID] {
			statement := block.Statement
			switch mode {
			case errorDisabled:
				b.add(block.ID, b.graph.ExceptionalExit, EdgeError, EdgeExceptional, true, statement)
			case errorNext:
				added := false
				for _, edge := range baseEdges {
					if edge.From == block.ID && edge.Class == EdgeNormal {
						b.add(block.ID, edge.To, EdgeError, EdgeExceptional, true, statement)
						added = true
					}
				}
				if !added {
					b.add(block.ID, b.graph.NormalExit, EdgeError, EdgeExceptional, true, statement)
				}
			default:
				target := strings.TrimPrefix(string(mode), "handler:")
				candidates := b.labels[target]
				if len(candidates) == 1 {
					b.add(block.ID, candidates[0], EdgeError, EdgeExceptional, true, statement)
				} else {
					b.add(block.ID, b.graph.UnknownExit, EdgeError, EdgeExceptional, true, statement)
				}
			}
		}
	}
	// Once a handler is active, a second fault does not re-enter that handler.
	// The CFG intentionally has no separate active-handler block state, so add
	// the conservative outward failure transition to every normal-flow block
	// reachable from a configured handler target.
	for _, statement := range b.procedure.Statements {
		if statement.Control == nil || statement.Control.Transfer != procedureir.TransferOnErrorGoto {
			continue
		}
		candidates := b.labels[normalizedTarget(statement.Control.Target)]
		if len(candidates) != 1 {
			continue
		}
		for blockID := range reachableFromEdges(candidates[0], baseEdges) {
			block := b.graph.Blocks[int(blockID)-1]
			if b.isFaultSite(block) {
				b.add(blockID, b.graph.ExceptionalExit, EdgeError, EdgeExceptional, true, block.Statement)
			}
		}
	}
}

func reachableFromEdges(entry BlockID, edges []Edge) map[BlockID]bool {
	seen := map[BlockID]bool{entry: true}
	queue := []BlockID{entry}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for _, edge := range edges {
			if edge.From != current || edge.Class != EdgeNormal || seen[edge.To] {
				continue
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
		}
	}
	return seen
}

func (b *builder) transferErrorMode(blockID BlockID, in map[errorMode]bool) map[errorMode]bool {
	block := b.graph.Blocks[int(blockID)-1]
	if block.Statement == nil || block.Statement.Control == nil {
		return in
	}
	switch block.Statement.Control.Transfer {
	case procedureir.TransferOnErrorDisable:
		return map[errorMode]bool{errorDisabled: true}
	case procedureir.TransferOnErrorResumeNext:
		return map[errorMode]bool{errorNext: true}
	case procedureir.TransferOnErrorGoto:
		target := normalizedTarget(block.Statement.Control.Target)
		if target == "" {
			return map[errorMode]bool{errorDisabled: true}
		}
		return map[errorMode]bool{errorMode("handler:" + target): true}
	default:
		return in
	}
}

func mergeModes(all map[BlockID]map[errorMode]bool, id BlockID, incoming map[errorMode]bool) bool {
	if all[id] == nil {
		all[id] = map[errorMode]bool{}
	}
	changed := false
	for mode := range incoming {
		if !all[id][mode] {
			all[id][mode] = true
			changed = true
		}
	}
	return changed
}

func (b *builder) isFaultSite(block Block) bool {
	if block.Statement == nil {
		return false
	}
	switch block.Statement.Kind {
	case procedureir.StatementLabel, procedureir.StatementOnError, procedureir.StatementResume,
		procedureir.StatementGoTo, procedureir.StatementExit, procedureir.StatementEnd:
		return false
	default:
		return true
	}
}

func (b *builder) finish() {
	sort.SliceStable(b.edges, func(i, j int) bool {
		a, c := b.edges[i], b.edges[j]
		if a.From != c.From {
			return a.From < c.From
		}
		if a.To != c.To {
			return a.To < c.To
		}
		if a.Class != c.Class {
			return a.Class < c.Class
		}
		if a.Kind != c.Kind {
			return a.Kind < c.Kind
		}
		if a.Uncertain != c.Uncertain {
			return !a.Uncertain
		}
		return a.StatementID < c.StatementID
	})
	out := make([]Edge, 0, len(b.edges))
	for _, edge := range b.edges {
		if len(out) > 0 {
			last := out[len(out)-1]
			last.ID, edge.ID = 0, 0
			if last == edge {
				continue
			}
		}
		edge.ID = EdgeID(len(out) + 1)
		out = append(out, edge)
	}
	b.graph.Edges = out
}
