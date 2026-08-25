package cfg

import "strings"

// WithoutNormalErrRaiseContinuation returns a graph where normal-flow edges
// leaving an Err.Raise or Error statement are removed. These statements never
// complete through normal VBA flow; exceptional edges remain available so
// active On Error modes, including Resume Next, are still represented.
func (g Graph) WithoutNormalErrRaiseContinuation() Graph {
	return g.View(EdgeFilter{WithoutNormalErrRaiseContinuation: true}).Materialize()
}

func isNonReturningRaiseBlock(block Block) bool {
	statement := block.Statement
	if statement == nil {
		return false
	}
	text := strings.ToLower(strings.TrimSpace(statement.Text))
	if strings.HasPrefix(text, "call ") || strings.HasPrefix(text, "call\t") {
		text = strings.TrimSpace(text[len("call"):])
	}
	if !strings.HasPrefix(text, "err.raise") {
		return strings.HasPrefix(text, "error ") || strings.HasPrefix(text, "error\t")
	}
	return len(text) == len("err.raise") || text[len("err.raise")] == ' ' || text[len("err.raise")] == '\t' || text[len("err.raise")] == '('
}

// BlockForStatement returns the block owning statementID.
func (g Graph) BlockForStatement(statementID int) (Block, bool) {
	if statementID <= 0 {
		return Block{}, false
	}
	blockIndex, ok := g.queryIndexes().blocksByStatement[statementID]
	if ok && blockIndex >= 0 && blockIndex < len(g.Blocks) {
		block := g.Blocks[blockIndex]
		if block.Kind == BlockStatement && block.StatementID == statementID {
			return block, true
		}
	}
	// Keep the original scan as a defensive fallback for graph values that
	// were changed after their index was built or have no indexed statement.
	for _, block := range g.Blocks {
		if block.Kind == BlockStatement && block.StatementID == statementID {
			return block, true
		}
	}
	return Block{}, false
}

// BlockByID returns the block identified by id. Block IDs are graph-local and
// are not required to be contiguous, so callers must not use an ID as a slice
// index. The lookup is backed by the graph query index and is O(1) on average.
func (g Graph) BlockByID(id BlockID) (Block, bool) {
	if id <= 0 {
		return Block{}, false
	}
	index := g.queryIndexes()
	blockIndex, ok := index.blocksByID[id]
	if ok && blockIndex >= 0 && blockIndex < len(g.Blocks) {
		block := g.Blocks[blockIndex]
		if block.ID == id {
			return block, true
		}
	}
	// Keep a defensive fallback for graph values whose public block slice was
	// changed after the index was built.
	for _, block := range g.Blocks {
		if block.ID == id {
			return block, true
		}
	}
	return Block{}, false
}

// OutgoingEdges returns edges leaving from. Edges retain their graph order and
// the returned slice is owned by the graph query index; callers must treat it
// as read-only. The lookup is O(1) on average.
func (g Graph) OutgoingEdges(from BlockID) []Edge {
	if from <= 0 {
		return nil
	}
	return g.queryIndexes().outgoing[from]
}

// Reachable returns reachable block IDs in stable ID order.
func (g Graph) Reachable(filter EdgeFilter) []BlockID {
	return g.View(filter).Reachable()
}

// IsReachable reports whether target is reachable from Entry.
func (g Graph) IsReachable(target BlockID, filter EdgeFilter) bool {
	return g.View(filter).IsReachable(target)
}

// Unreachable returns statement blocks that cannot be reached from Entry.
func (g Graph) Unreachable(filter EdgeFilter) []Block {
	seen := g.reachableWithout(filter, nil)
	var out []Block
	for _, block := range g.Blocks {
		if block.Kind == BlockStatement && !seen[block.ID] {
			out = append(out, block)
		}
	}
	return out
}

// ExitTransitions returns reachable edges into selected synthetic exits.
func (g Graph) ExitTransitions(selection ExitSelection, filter EdgeFilter) []Edge {
	view := g.View(filter)
	reachable := view.reachableSet()
	selected := idSet(selection.ids(g))
	var out []Edge
	view.ForEachEdge(func(edge Edge) bool {
		if reachable[edge.From] && selected[edge.To] {
			out = append(out, edge)
		}
		return true
	})
	return out
}

// Dominators computes conservative dominators. Uncertain edges are always
// included; NormalOnly may be used for explicitly normal-flow consumers.
func (g Graph) Dominators(filter EdgeFilter) map[BlockID][]BlockID {
	return g.View(filter).Dominators()
}

// DefiniteAssignments returns variables definitely assigned on entry to every
// reachable block. It is a must-analysis and always includes uncertain edges.
func (g Graph) DefiniteAssignments(filter EdgeFilter) map[BlockID][]Variable {
	return g.View(filter).DefiniteAssignments()
}

// IsDefinitelyAssigned is the case-insensitive convenience query.
func (g Graph) IsDefinitelyAssigned(block BlockID, variable Variable, filter EdgeFilter) bool {
	variable = variable.canonical()
	for _, assigned := range g.DefiniteAssignments(filter)[block] {
		if assigned == variable {
			return true
		}
	}
	return false
}

// CleanupGuaranteed reports whether every reachable selected exit path crosses
// at least one cleanup statement. Unknown/exceptional exits participate by
// default, and uncertain edges can never be excluded.
func (g Graph) CleanupGuaranteed(cleanupStatementIDs []int, selection ExitSelection, filter EdgeFilter) bool {
	removed := map[BlockID]bool{}
	for _, statementID := range cleanupStatementIDs {
		if block, ok := g.BlockForStatement(statementID); ok {
			removed[block.ID] = true
		}
	}
	if len(removed) == 0 {
		return false
	}
	reachable := g.reachableWithout(filter, removed)
	for _, exit := range selection.ids(g) {
		if reachable[exit] {
			return false
		}
	}
	return true
}

func (g Graph) reachableWithout(filter EdgeFilter, removed map[BlockID]bool) map[BlockID]bool {
	if removed == nil {
		return cloneBlockSet(g.View(filter).reachableSet())
	}
	return g.View(filter).reachableSetWithoutTargets(removed)
}

func (g Graph) predecessors(id BlockID, filter EdgeFilter, reachable map[BlockID]bool) []BlockID {
	return g.View(filter).predecessors(id, reachable)
}

func (g Graph) block(id BlockID) Block {
	block, _ := g.BlockByID(id)
	return block
}

func idSet(ids []BlockID) map[BlockID]bool {
	out := map[BlockID]bool{}
	for _, id := range ids {
		out[id] = true
	}
	return out
}

func copySet(in map[BlockID]bool) map[BlockID]bool {
	out := map[BlockID]bool{}
	for key := range in {
		out[key] = true
	}
	return out
}

func sameSet(a, b map[BlockID]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for key := range a {
		if !b[key] {
			return false
		}
	}
	return true
}

func copyVariableSet(in map[Variable]bool) map[Variable]bool {
	out := map[Variable]bool{}
	for key := range in {
		out[key] = true
	}
	return out
}

func sameVariableSet(a, b map[Variable]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for key := range a {
		if !b[key] {
			return false
		}
	}
	return true
}
