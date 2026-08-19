package cfg

import (
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// WithoutNormalErrRaiseContinuation returns a graph where normal-flow edges
// leaving an Err.Raise or Error statement are removed. These statements never
// complete through normal VBA flow; exceptional edges remain available so
// active On Error modes, including Resume Next, are still represented.
func (g Graph) WithoutNormalErrRaiseContinuation() Graph {
	g.Blocks = append([]Block(nil), g.Blocks...)
	edges := g.Edges
	g.Edges = make([]Edge, 0, len(edges))
	for _, edge := range edges {
		if edge.Class == EdgeNormal && g.isNonReturningRaiseBlock(edge.From) {
			continue
		}
		g.Edges = append(g.Edges, edge)
	}
	g.query = buildQueryIndex(g)
	return g
}

func (g Graph) isNonReturningRaiseBlock(id BlockID) bool {
	if id <= 0 || int(id) > len(g.Blocks) {
		return false
	}
	statement := g.Blocks[int(id)-1].Statement
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

// Reachable returns reachable block IDs in stable ID order.
func (g Graph) Reachable(filter EdgeFilter) []BlockID {
	seen := g.reachableWithout(filter, nil)
	out := make([]BlockID, 0, len(seen))
	for _, block := range g.Blocks {
		if seen[block.ID] {
			out = append(out, block.ID)
		}
	}
	return out
}

// IsReachable reports whether target is reachable from Entry.
func (g Graph) IsReachable(target BlockID, filter EdgeFilter) bool {
	index := g.queryIndexes()
	if reachable, ok := index.cachedReachable(filter); ok {
		return reachable[target]
	}
	return g.reachableWithout(filter, nil)[target]
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
	reachable := g.reachableWithout(filter, nil)
	selected := idSet(selection.ids(g))
	var out []Edge
	for _, edge := range g.Edges {
		if filter.accepts(edge) && reachable[edge.From] && selected[edge.To] {
			out = append(out, edge)
		}
	}
	return out
}

// Dominators computes conservative dominators. Uncertain edges are always
// included; NormalOnly may be used for explicitly normal-flow consumers.
func (g Graph) Dominators(filter EdgeFilter) map[BlockID][]BlockID {
	reachable := g.reachableWithout(filter, nil)
	all := map[BlockID]bool{}
	for id := range reachable {
		all[id] = true
	}
	dom := map[BlockID]map[BlockID]bool{}
	for id := range reachable {
		if id == g.Entry {
			dom[id] = map[BlockID]bool{id: true}
		} else {
			dom[id] = copySet(all)
		}
	}
	changed := true
	for changed {
		changed = false
		unknownDominators, hasUnknown := g.unknownDominatorInput(reachable, dom)
		for id := range reachable {
			if id == g.Entry {
				continue
			}
			preds := g.predecessors(id, filter, reachable)
			next := map[BlockID]bool{id: true}
			var intersection map[BlockID]bool
			if len(preds) > 0 {
				intersection = copySet(dom[preds[0]])
				for _, pred := range preds[1:] {
					for candidate := range intersection {
						if !dom[pred][candidate] {
							delete(intersection, candidate)
						}
					}
				}
			}
			if hasUnknown && g.block(id).Kind == BlockStatement {
				if intersection == nil {
					intersection = copySet(unknownDominators)
				} else {
					for candidate := range intersection {
						if !unknownDominators[candidate] {
							delete(intersection, candidate)
						}
					}
				}
			}
			for candidate := range intersection {
				next[candidate] = true
			}
			if !sameSet(dom[id], next) {
				dom[id] = next
				changed = true
			}
		}
	}
	out := map[BlockID][]BlockID{}
	for id, set := range dom {
		for candidate := range set {
			out[id] = append(out[id], candidate)
		}
		sort.Slice(out[id], func(i, j int) bool { return out[id][i] < out[id][j] })
	}
	return out
}

func (g Graph) unknownDominatorInput(
	reachable map[BlockID]bool,
	dominators map[BlockID]map[BlockID]bool,
) (map[BlockID]bool, bool) {
	var intersection map[BlockID]bool
	for _, source := range g.UnknownFlowSources {
		if !reachable[source] {
			continue
		}
		if intersection == nil {
			intersection = copySet(dominators[source])
			continue
		}
		for candidate := range intersection {
			if !dominators[source][candidate] {
				delete(intersection, candidate)
			}
		}
	}
	return intersection, intersection != nil
}

// DefiniteAssignments returns variables definitely assigned on entry to every
// reachable block. It is a must-analysis and always includes uncertain edges.
func (g Graph) DefiniteAssignments(filter EdgeFilter) map[BlockID][]Variable {
	reachable := g.reachableWithout(filter, nil)
	universe := map[Variable]bool{}
	parameters := map[Variable]bool{}
	for _, parameter := range g.Procedure.Parameters {
		variable := Variable{Scope: procedureir.ScopeParameter, Name: parameter.Name}.canonical()
		universe[variable] = true
		parameters[variable] = true
	}
	for _, block := range g.Blocks {
		for _, variable := range block.Assignments {
			universe[variable.canonical()] = true
		}
	}
	in := map[BlockID]map[Variable]bool{}
	outSet := map[BlockID]map[Variable]bool{}
	for id := range reachable {
		if id == g.Entry {
			in[id] = copyVariableSet(parameters)
		} else {
			in[id] = copyVariableSet(universe)
		}
		outSet[id] = withBlockAssignments(in[id], g.block(id))
	}
	changed := true
	for changed {
		changed = false
		unknownInput, hasUnknown := g.unknownAssignmentInput(reachable, in)
		for id := range reachable {
			if id == g.Entry {
				continue
			}
			incoming := g.assignmentInputs(id, filter, reachable, in, outSet)
			if hasUnknown && g.block(id).Kind == BlockStatement {
				incoming = append(incoming, unknownInput)
			}
			next := map[Variable]bool{}
			if len(incoming) > 0 {
				next = copyVariableSet(incoming[0])
				for _, predecessor := range incoming[1:] {
					for variable := range next {
						if !predecessor[variable] {
							delete(next, variable)
						}
					}
				}
			}
			if !sameVariableSet(in[id], next) {
				in[id] = next
				outSet[id] = withBlockAssignments(next, g.block(id))
				changed = true
			}
		}
	}
	result := map[BlockID][]Variable{}
	for id, set := range in {
		for variable := range set {
			result[id] = append(result[id], variable)
		}
		sort.Slice(result[id], func(i, j int) bool {
			if result[id][i].Scope != result[id][j].Scope {
				return result[id][i].Scope < result[id][j].Scope
			}
			return result[id][i].Name < result[id][j].Name
		})
	}
	return result
}

func (g Graph) unknownAssignmentInput(
	reachable map[BlockID]bool,
	in map[BlockID]map[Variable]bool,
) (map[Variable]bool, bool) {
	var intersection map[Variable]bool
	for _, source := range g.UnknownFlowSources {
		if !reachable[source] {
			continue
		}
		if intersection == nil {
			intersection = copyVariableSet(in[source])
			continue
		}
		for variable := range intersection {
			if !in[source][variable] {
				delete(intersection, variable)
			}
		}
	}
	return intersection, intersection != nil
}

func (g Graph) assignmentInputs(
	id BlockID,
	filter EdgeFilter,
	reachable map[BlockID]bool,
	in, out map[BlockID]map[Variable]bool,
) []map[Variable]bool {
	var inputs []map[Variable]bool
	for _, edge := range g.Edges {
		if edge.To != id || !filter.accepts(edge) || !reachable[edge.From] {
			continue
		}
		source := g.block(edge.From)
		forEachZeroIteration := source.Statement != nil &&
			source.Statement.Kind == procedureir.StatementForEach && edge.Kind == EdgeLoopExit
		unknownBeforeCompletion := edge.Kind == EdgeUnknown && edge.Uncertain
		if edge.Class == EdgeExceptional || forEachZeroIteration || unknownBeforeCompletion {
			// A fault can occur before the source statement's writes complete.
			// For Each may also take its zero-iteration exit before assigning
			// the iterator variable. Unknown control may diverge before a
			// recovered statement's writes complete.
			inputs = append(inputs, in[edge.From])
		} else {
			inputs = append(inputs, out[edge.From])
		}
	}
	return inputs
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
	index := g.queryIndexes()
	if removed == nil {
		if reachable, ok := index.cachedReachable(filter); ok {
			return cloneBlockSet(reachable)
		}
	}
	seen := physicalReachableWithIndex(g, index, filter, removed)
	if g.unknownFlowReached(seen) {
		for _, block := range g.Blocks {
			if block.Kind == BlockStatement && !removed[block.ID] {
				seen[block.ID] = true
			}
		}
		if !removed[g.UnknownExit] {
			seen[g.UnknownExit] = true
		}
		seen = expandReachableWithIndex(index, filter, removed, seen)
	}
	return seen
}

func (g Graph) unknownFlowReached(reachable map[BlockID]bool) bool {
	for _, source := range g.UnknownFlowSources {
		if reachable[source] {
			return true
		}
	}
	return false
}

func (g Graph) predecessors(id BlockID, filter EdgeFilter, reachable map[BlockID]bool) []BlockID {
	var out []BlockID
	seen := map[BlockID]bool{}
	for _, edge := range g.queryIndexes().incoming[id] {
		if filter.accepts(edge) && reachable[edge.From] && !seen[edge.From] {
			seen[edge.From] = true
			out = append(out, edge.From)
		}
	}
	sort.Slice(out, func(i, j int) bool { return out[i] < out[j] })
	return out
}

func (g Graph) block(id BlockID) Block {
	if id > 0 && int(id) <= len(g.Blocks) {
		return g.Blocks[int(id)-1]
	}
	return Block{}
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

func withBlockAssignments(in map[Variable]bool, block Block) map[Variable]bool {
	out := copyVariableSet(in)
	for _, variable := range block.Assignments {
		out[variable.canonical()] = true
	}
	return out
}
