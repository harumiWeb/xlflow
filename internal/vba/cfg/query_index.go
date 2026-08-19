package cfg

// queryIndex contains immutable indexes for one graph revision. Graph values
// are copied frequently, so the index is shared by value copies while Clone
// and graph transformations rebuild it for their independent slice storage.
type queryIndex struct {
	blocksByStatement map[int]int
	outgoing          map[BlockID][]Edge
	incoming          map[BlockID][]Edge
	reachableAll      map[BlockID]bool
	reachableNormal   map[BlockID]bool
	blocksBase        *Block
	blocksLen         int
	edgesBase         *Edge
	edgesLen          int
}

func buildQueryIndex(g Graph) *queryIndex {
	index := &queryIndex{
		blocksByStatement: make(map[int]int),
		outgoing:          make(map[BlockID][]Edge),
		incoming:          make(map[BlockID][]Edge),
		blocksLen:         len(g.Blocks),
		edgesLen:          len(g.Edges),
	}
	if len(g.Blocks) > 0 {
		index.blocksBase = &g.Blocks[0]
	}
	if len(g.Edges) > 0 {
		index.edgesBase = &g.Edges[0]
	}
	for blockIndex, block := range g.Blocks {
		if block.Kind == BlockStatement && block.StatementID > 0 {
			if _, exists := index.blocksByStatement[block.StatementID]; !exists {
				index.blocksByStatement[block.StatementID] = blockIndex
			}
		}
	}
	for _, edge := range g.Edges {
		index.outgoing[edge.From] = append(index.outgoing[edge.From], edge)
		index.incoming[edge.To] = append(index.incoming[edge.To], edge)
	}
	index.reachableAll = computeReachable(g, index, EdgeFilter{})
	index.reachableNormal = computeReachable(g, index, EdgeFilter{NormalOnly: true})
	return index
}

func (g Graph) queryIndexes() *queryIndex {
	if g.query != nil && g.query.matches(g) {
		return g.query
	}
	// Some internal tests construct Graph literals directly, and callers may
	// replace the public block/edge slices on a copied graph. Rebuild when the
	// slice storage changes so the fallback remains correct without penalizing
	// stable graph revisions.
	return buildQueryIndex(g)
}

func (index *queryIndex) matches(g Graph) bool {
	if index.blocksLen != len(g.Blocks) || index.edgesLen != len(g.Edges) {
		return false
	}
	if len(g.Blocks) > 0 && index.blocksBase != &g.Blocks[0] {
		return false
	}
	if len(g.Edges) > 0 && index.edgesBase != &g.Edges[0] {
		return false
	}
	return true
}

func computeReachable(g Graph, index *queryIndex, filter EdgeFilter) map[BlockID]bool {
	seen := physicalReachableWithIndex(g, index, filter, nil)
	if unknownFlowReachedFor(g, seen) {
		for _, block := range g.Blocks {
			if block.Kind == BlockStatement {
				seen[block.ID] = true
			}
		}
		seen[g.UnknownExit] = true
		seen = expandReachableWithIndex(index, filter, nil, seen)
	}
	return seen
}

func unknownFlowReachedFor(g Graph, reachable map[BlockID]bool) bool {
	for _, source := range g.UnknownFlowSources {
		if reachable[source] {
			return true
		}
	}
	return false
}

func physicalReachableWithIndex(g Graph, index *queryIndex, filter EdgeFilter, removed map[BlockID]bool) map[BlockID]bool {
	seen := map[BlockID]bool{}
	if removed != nil && removed[g.Entry] {
		return seen
	}
	seen[g.Entry] = true
	return expandReachableWithIndex(index, filter, removed, seen)
}

func expandReachableWithIndex(index *queryIndex, filter EdgeFilter, removed map[BlockID]bool, seen map[BlockID]bool) map[BlockID]bool {
	queue := make([]BlockID, 0, len(seen))
	for id := range seen {
		queue = append(queue, id)
	}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for _, edge := range index.outgoing[current] {
			if !filter.accepts(edge) || removed != nil && removed[edge.To] {
				continue
			}
			if seen[edge.To] {
				continue
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
		}
	}
	return seen
}

func cloneBlockSet(in map[BlockID]bool) map[BlockID]bool {
	out := make(map[BlockID]bool, len(in))
	for id, present := range in {
		out[id] = present
	}
	return out
}
