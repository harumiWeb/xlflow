package cfg

import (
	"sort"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// CFGView is a read-only flow view over one immutable graph revision. It
// shares the graph's block/edge storage and query index; filtering never
// materializes a second graph. The callbacks exposed by this type must not
// retain or mutate the values they receive.
type CFGView struct {
	graph              Graph
	index              *queryIndex
	filter             EdgeFilter
	excludedNormalFrom map[BlockID]bool
	cache              *viewCache
}

// viewCache is owned by one graph revision/filter identity. sync.Once keeps
// lazy derived queries deterministic when analyzer readers arrive together.
type viewCache struct {
	reachableOnce  sync.Once
	reachable      map[BlockID]bool
	dominatorsOnce sync.Once
	dominators     map[BlockID][]BlockID
}

// View returns a read-only view for filter. Canonical filters share a bounded
// cache owned by the graph revision. A graph literal without a retained index
// still gets a stable cache for the lifetime of this view.
func (g Graph) View(filter EdgeFilter) CFGView {
	index := g.queryIndexes()
	return CFGView{
		graph:  g,
		index:  index,
		filter: filter,
		cache:  &index.viewCaches[filterCacheKey(filter)],
	}
}

// WithoutNormalErrRaiseContinuationView returns the allocation-free canonical
// view used by error-outcome consumers. Use Materialize only at compatibility
// boundaries that still require mutable Graph fields.
func (g Graph) WithoutNormalErrRaiseContinuationView() CFGView {
	return g.View(EdgeFilter{WithoutNormalErrRaiseContinuation: true})
}

// WithoutNormalContinuationsFrom returns a view that additionally removes
// normal edges leaving the supplied block IDs. The input is copied because a
// caller may reuse its temporary map after constructing the view. The custom
// mask has a view-local cache and therefore cannot grow the revision cache.
func (v CFGView) WithoutNormalContinuationsFrom(blocks map[BlockID]bool) CFGView {
	if len(blocks) == 0 {
		return v
	}
	removed := make(map[BlockID]bool, len(v.excludedNormalFrom)+len(blocks))
	for id, present := range v.excludedNormalFrom {
		if present {
			removed[id] = true
		}
	}
	for id, present := range blocks {
		if present {
			removed[id] = true
		}
	}
	if len(removed) == 0 {
		return v
	}
	return CFGView{
		graph:              v.graph,
		index:              v.index,
		filter:             v.filter,
		excludedNormalFrom: removed,
		cache:              &viewCache{},
	}
}

func filterCacheKey(filter EdgeFilter) int {
	key := 0
	if filter.NormalOnly {
		key |= 1
	}
	if filter.WithoutNormalErrRaiseContinuation {
		key |= 2
	}
	return key
}

// Materialize retains the legacy defensive-copy behavior for callers that
// still need mutable Graph fields. New analyzer code should stay on CFGView.
func (v CFGView) Materialize() Graph {
	out := v.graph
	out.Blocks = append([]Block(nil), v.graph.Blocks...)
	out.Edges = make([]Edge, 0, len(v.graph.Edges))
	v.ForEachEdge(func(edge Edge) bool {
		out.Edges = append(out.Edges, edge)
		return true
	})
	out.query = buildQueryIndex(out)
	return out
}

// ForEachBlock visits blocks in their stable source/storage order.
func (v CFGView) ForEachBlock(fn func(Block) bool) {
	if fn == nil {
		return
	}
	for _, block := range v.graph.Blocks {
		if !fn(block) {
			return
		}
	}
}

// BlockCount returns the number of blocks in this revision. The count is
// independent of the edge filter because views never remove block storage.
func (v CFGView) BlockCount() int { return len(v.graph.Blocks) }

// ForEachBlockOrdinal visits blocks with their stable dense revision-local
// ordinals. The ordinal is the base block-storage position, not the public
// sparse BlockID.
func (v CFGView) ForEachBlockOrdinal(fn func(BlockOrdinal, Block) bool) {
	if fn == nil {
		return
	}
	for ordinal, block := range v.graph.Blocks {
		if !fn(BlockOrdinal(ordinal), block) {
			return
		}
	}
}

// BlockAtOrdinal resolves one dense revision-local ordinal without exposing
// the underlying block slice.
func (v CFGView) BlockAtOrdinal(ordinal BlockOrdinal) (Block, bool) {
	if ordinal < 0 || int(ordinal) >= len(v.graph.Blocks) {
		return Block{}, false
	}
	return v.graph.Blocks[int(ordinal)], true
}

// Ordinal resolves a sparse public BlockID to its stable dense ordinal. It
// follows the first-match rule used by the base revision, including defensive
// directly-constructed block ID zero values.
func (v CFGView) Ordinal(id BlockID) (BlockOrdinal, bool) {
	if id < 0 {
		return 0, false
	}
	blockIndex, ok := v.index.blocksByID[id]
	if !ok || blockIndex < 0 || blockIndex >= len(v.graph.Blocks) || v.graph.Blocks[blockIndex].ID != id {
		return 0, false
	}
	return BlockOrdinal(blockIndex), true
}

// ForEachEdge visits accepted edges in base graph order.
func (v CFGView) ForEachEdge(fn func(Edge) bool) {
	if fn == nil {
		return
	}
	for _, edge := range v.graph.Edges {
		if !v.accepts(edge) {
			continue
		}
		if !fn(edge) {
			return
		}
	}
}

// ForEachOutgoing visits accepted outgoing edges in query-index order.
func (v CFGView) ForEachOutgoing(from BlockID, fn func(Edge) bool) {
	if fn == nil {
		return
	}
	for _, edge := range v.index.outgoing[from] {
		if !v.accepts(edge) {
			continue
		}
		if !fn(edge) {
			return
		}
	}
}

// ForEachIncoming visits accepted incoming edges in query-index order.
func (v CFGView) ForEachIncoming(to BlockID, fn func(Edge) bool) {
	if fn == nil {
		return
	}
	for _, edge := range v.index.incoming[to] {
		if !v.accepts(edge) {
			continue
		}
		if !fn(edge) {
			return
		}
	}
}

func (v CFGView) accepts(edge Edge) bool {
	if !v.filter.accepts(edge) {
		return false
	}
	if edge.Class != EdgeNormal {
		return true
	}
	if v.excludedNormalFrom[edge.From] {
		return false
	}
	return !v.filter.WithoutNormalErrRaiseContinuation || !v.index.nonReturningRaise[edge.From]
}

// BlockByID resolves a graph-local block through the revision index. Unlike
// the legacy Graph lookup, a view also addresses defensive directly-constructed
// ID zero values; Graph remains the compatibility boundary for its historical
// positive-ID behavior.
func (v CFGView) BlockByID(id BlockID) (Block, bool) { return v.blockByID(id) }

func (v CFGView) blockByID(id BlockID) (Block, bool) {
	if id < 0 {
		return Block{}, false
	}
	blockIndex, ok := v.index.blocksByID[id]
	if !ok || blockIndex < 0 || blockIndex >= len(v.graph.Blocks) || v.graph.Blocks[blockIndex].ID != id {
		return Block{}, false
	}
	return v.graph.Blocks[blockIndex], true
}

// BlockForStatement resolves a statement block using the base query index.
func (v CFGView) BlockForStatement(statementID int) (Block, bool) {
	return v.graph.BlockForStatement(statementID)
}

func (v CFGView) Entry() BlockID           { return v.graph.Entry }
func (v CFGView) NormalExit() BlockID      { return v.graph.NormalExit }
func (v CFGView) ExceptionalExit() BlockID { return v.graph.ExceptionalExit }
func (v CFGView) TerminationExit() BlockID { return v.graph.TerminationExit }
func (v CFGView) UnknownExit() BlockID     { return v.graph.UnknownExit }

// Reachable returns accepted reachable block IDs in stable graph order. The
// cached set is copied because the legacy materialized API is mutable.
func (v CFGView) Reachable() []BlockID {
	seen := v.reachableSet()
	out := make([]BlockID, 0, len(seen))
	for _, block := range v.graph.Blocks {
		if seen[block.ID] {
			out = append(out, block.ID)
		}
	}
	return out
}

// IsReachable reads cached membership without materializing a result slice.
func (v CFGView) IsReachable(target BlockID) bool { return v.reachableSet()[target] }

func (v CFGView) reachableSet() map[BlockID]bool {
	v.cache.reachableOnce.Do(func() { v.cache.reachable = v.computeReachable() })
	return v.cache.reachable
}

// reachableSetWithoutTargets computes a one-off reachability result while
// treating removed block IDs as cut targets. Cleanup guarantees use this
// overlay to ask whether an exit remains reachable after bypassing selected
// cleanup statements; it intentionally does not participate in the bounded
// canonical view cache.
func (v CFGView) reachableSetWithoutTargets(removed map[BlockID]bool) map[BlockID]bool {
	seen := map[BlockID]bool{}
	if removed[v.graph.Entry] {
		return seen
	}
	seen[v.graph.Entry] = true
	queue := []BlockID{v.graph.Entry}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		v.ForEachOutgoing(current, func(edge Edge) bool {
			if removed[edge.To] || seen[edge.To] {
				return true
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
			return true
		})
	}
	unknownReached := false
	for _, source := range v.graph.UnknownFlowSources {
		if seen[source] {
			unknownReached = true
			break
		}
	}
	if unknownReached {
		for _, block := range v.graph.Blocks {
			if block.Kind == BlockStatement && !removed[block.ID] {
				seen[block.ID] = true
			}
		}
		if !removed[v.graph.UnknownExit] {
			seen[v.graph.UnknownExit] = true
		}
		queue = queue[:0]
		for id := range seen {
			queue = append(queue, id)
		}
		for len(queue) > 0 {
			current := queue[0]
			queue = queue[1:]
			v.ForEachOutgoing(current, func(edge Edge) bool {
				if removed[edge.To] || seen[edge.To] {
					return true
				}
				seen[edge.To] = true
				queue = append(queue, edge.To)
				return true
			})
		}
	}
	return seen
}

func (v CFGView) computeReachable() map[BlockID]bool {
	seen := map[BlockID]bool{}
	seen[v.graph.Entry] = true
	queue := []BlockID{v.graph.Entry}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		v.ForEachOutgoing(current, func(edge Edge) bool {
			if !seen[edge.To] {
				seen[edge.To] = true
				queue = append(queue, edge.To)
			}
			return true
		})
	}
	unknownReached := false
	for _, source := range v.graph.UnknownFlowSources {
		if seen[source] {
			unknownReached = true
			break
		}
	}
	if unknownReached {
		for _, block := range v.graph.Blocks {
			if block.Kind == BlockStatement {
				seen[block.ID] = true
			}
		}
		seen[v.graph.UnknownExit] = true
		queue = queue[:0]
		for id := range seen {
			queue = append(queue, id)
		}
		for len(queue) > 0 {
			current := queue[0]
			queue = queue[1:]
			v.ForEachOutgoing(current, func(edge Edge) bool {
				if !seen[edge.To] {
					seen[edge.To] = true
					queue = append(queue, edge.To)
				}
				return true
			})
		}
	}
	return seen
}

// Dominators returns a defensive map copy of the cached conservative result.
// DominatorsOf is available to read a single cached entry without copying the
// complete result.
func (v CFGView) Dominators() map[BlockID][]BlockID {
	result := v.dominatorSet()
	out := make(map[BlockID][]BlockID, len(result))
	for id, values := range result {
		out[id] = append([]BlockID(nil), values...)
	}
	return out
}

// DominatorsOf returns an owned copy of one cached dominator list. Copying one
// entry keeps the view read-only without materializing the complete result.
func (v CFGView) DominatorsOf(block BlockID) []BlockID {
	return append([]BlockID(nil), v.dominatorSet()[block]...)
}

func (v CFGView) dominatorSet() map[BlockID][]BlockID {
	v.cache.dominatorsOnce.Do(func() { v.cache.dominators = v.computeDominators() })
	return v.cache.dominators
}

func (v CFGView) computeDominators() map[BlockID][]BlockID {
	reachable := v.reachableSet()
	all := map[BlockID]bool{}
	for id := range reachable {
		all[id] = true
	}
	dom := map[BlockID]map[BlockID]bool{}
	for id := range reachable {
		if id == v.graph.Entry {
			dom[id] = map[BlockID]bool{id: true}
		} else {
			dom[id] = copySet(all)
		}
	}
	changed := true
	for changed {
		changed = false
		unknownDominators, hasUnknown := v.unknownDominatorInput(reachable, dom)
		for id := range reachable {
			if id == v.graph.Entry {
				continue
			}
			preds := v.predecessors(id, reachable)
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
			if hasUnknown {
				block, ok := v.blockByID(id)
				if ok && block.Kind == BlockStatement {
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

func (v CFGView) unknownDominatorInput(
	reachable map[BlockID]bool,
	dominators map[BlockID]map[BlockID]bool,
) (map[BlockID]bool, bool) {
	var intersection map[BlockID]bool
	for _, source := range v.graph.UnknownFlowSources {
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

func (v CFGView) predecessors(id BlockID, reachable map[BlockID]bool) []BlockID {
	seen := map[BlockID]bool{}
	var out []BlockID
	v.ForEachIncoming(id, func(edge Edge) bool {
		if reachable[edge.From] && !seen[edge.From] {
			seen[edge.From] = true
			out = append(out, edge.From)
		}
		return true
	})
	sort.Slice(out, func(i, j int) bool { return out[i] < out[j] })
	return out
}

// DefiniteAssignments returns variables definitely assigned at every reachable
// block boundary. It shares the view's filtered reachability and adjacency.
func (v CFGView) DefiniteAssignments() map[BlockID][]Variable {
	return v.DefiniteAssignmentsWith(nil)
}

// DefiniteAssignmentsWith is the compatibility-preserving assignment overlay
// hook used by return-path analysis. Returning false for one block/variable
// suppresses that assignment without copying the block or edge slices.
func (v CFGView) DefiniteAssignmentsWith(allow func(Block, Variable) bool) map[BlockID][]Variable {
	reachable := v.reachableSet()
	universe := map[Variable]bool{}
	parameters := map[Variable]bool{}
	for _, parameter := range v.graph.Procedure.Parameters {
		variable := (Variable{Scope: procedureir.ScopeParameter, Name: parameter.Name}).canonical()
		universe[variable] = true
		parameters[variable] = true
	}
	for _, block := range v.graph.Blocks {
		for _, variable := range block.Assignments {
			variable = variable.canonical()
			if allow == nil || allow(block, variable) {
				universe[variable] = true
			}
		}
	}
	in := map[BlockID]map[Variable]bool{}
	outSet := map[BlockID]map[Variable]bool{}
	for id := range reachable {
		if id == v.graph.Entry {
			in[id] = copyVariableSet(parameters)
		} else {
			in[id] = copyVariableSet(universe)
		}
		block, _ := v.blockByID(id)
		outSet[id] = v.withBlockAssignments(in[id], block, allow)
	}
	changed := true
	for changed {
		changed = false
		unknownInput, hasUnknown := v.unknownAssignmentInput(reachable, in)
		for id := range reachable {
			if id == v.graph.Entry {
				continue
			}
			incoming := v.assignmentInputs(id, reachable, in, outSet)
			block, _ := v.blockByID(id)
			if hasUnknown && block.Kind == BlockStatement {
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
				outSet[id] = v.withBlockAssignments(next, block, allow)
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

func (v CFGView) withBlockAssignments(in map[Variable]bool, block Block, allow func(Block, Variable) bool) map[Variable]bool {
	out := copyVariableSet(in)
	for _, variable := range block.Assignments {
		variable = variable.canonical()
		if allow == nil || allow(block, variable) {
			out[variable] = true
		}
	}
	return out
}

func (v CFGView) unknownAssignmentInput(reachable map[BlockID]bool, in map[BlockID]map[Variable]bool) (map[Variable]bool, bool) {
	var intersection map[Variable]bool
	for _, source := range v.graph.UnknownFlowSources {
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

func (v CFGView) assignmentInputs(id BlockID, reachable map[BlockID]bool, in, out map[BlockID]map[Variable]bool) []map[Variable]bool {
	var inputs []map[Variable]bool
	v.ForEachIncoming(id, func(edge Edge) bool {
		if !reachable[edge.From] {
			return true
		}
		source, _ := v.blockByID(edge.From)
		forEachZeroIteration := source.Statement != nil &&
			source.Statement.Kind == procedureir.StatementForEach && edge.Kind == EdgeLoopExit
		unknownBeforeCompletion := edge.Kind == EdgeUnknown && edge.Uncertain
		if edge.Class == EdgeExceptional || forEachZeroIteration || unknownBeforeCompletion {
			inputs = append(inputs, in[edge.From])
		} else {
			inputs = append(inputs, out[edge.From])
		}
		return true
	})
	return inputs
}
