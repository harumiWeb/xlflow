package cfg

// queryIndex contains immutable indexes for one graph revision. Graph values
// are copied frequently, so the index is shared by value copies while Clone
// and graph transformations rebuild it for their independent slice storage.
type queryIndex struct {
	blocksByStatement  map[int]int
	blocksByID         map[BlockID]int
	outgoing           map[BlockID][]Edge
	incoming           map[BlockID][]Edge
	nonReturningRaise  map[BlockID]bool
	viewCaches         [4]viewCache
	blocksBase         *Block
	blocksLen          int
	edgesBase          *Edge
	edgesLen           int
	entry              BlockID
	unknownExit        BlockID
	unknownFlowSources []BlockID
}

func buildQueryIndex(g Graph) *queryIndex {
	blockCapacity := len(g.Blocks)
	index := &queryIndex{
		blocksByStatement:  make(map[int]int, blockCapacity),
		blocksByID:         make(map[BlockID]int, blockCapacity),
		outgoing:           make(map[BlockID][]Edge, blockCapacity),
		incoming:           make(map[BlockID][]Edge, blockCapacity),
		nonReturningRaise:  make(map[BlockID]bool, blockCapacity),
		blocksLen:          len(g.Blocks),
		edgesLen:           len(g.Edges),
		entry:              g.Entry,
		unknownExit:        g.UnknownExit,
		unknownFlowSources: append([]BlockID(nil), g.UnknownFlowSources...),
	}
	if len(g.Blocks) > 0 {
		index.blocksBase = &g.Blocks[0]
	}
	if len(g.Edges) > 0 {
		index.edgesBase = &g.Edges[0]
	}
	for blockIndex, block := range g.Blocks {
		if _, exists := index.blocksByID[block.ID]; !exists {
			index.blocksByID[block.ID] = blockIndex
			index.nonReturningRaise[block.ID] = isNonReturningRaiseBlock(block)
		}
		if block.Kind == BlockStatement && block.StatementID > 0 {
			if _, exists := index.blocksByStatement[block.StatementID]; !exists {
				index.blocksByStatement[block.StatementID] = blockIndex
			}
		}
	}
	outgoingRemaining := make(map[BlockID]int, blockCapacity)
	incomingRemaining := make(map[BlockID]int, blockCapacity)
	for _, edge := range g.Edges {
		outgoingRemaining[edge.From]++
		incomingRemaining[edge.To]++
	}
	outgoingStorage := make([]Edge, len(g.Edges))
	incomingStorage := make([]Edge, len(g.Edges))
	outgoingOffset := 0
	for from, count := range outgoingRemaining {
		index.outgoing[from] = outgoingStorage[outgoingOffset : outgoingOffset+count : outgoingOffset+count]
		outgoingOffset += count
	}
	incomingOffset := 0
	for to, count := range incomingRemaining {
		index.incoming[to] = incomingStorage[incomingOffset : incomingOffset+count : incomingOffset+count]
		incomingOffset += count
	}
	for _, edge := range g.Edges {
		outgoing := index.outgoing[edge.From]
		outgoingIndex := len(outgoing) - outgoingRemaining[edge.From]
		outgoing[outgoingIndex] = edge
		outgoingRemaining[edge.From]--
		incoming := index.incoming[edge.To]
		incomingIndex := len(incoming) - incomingRemaining[edge.To]
		incoming[incomingIndex] = edge
		incomingRemaining[edge.To]--
	}
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
	return index.entry == g.Entry && index.unknownExit == g.UnknownExit &&
		sameBlockIDs(index.unknownFlowSources, g.UnknownFlowSources)
}

func sameBlockIDs(a, b []BlockID) bool {
	if len(a) != len(b) {
		return false
	}
	for i, id := range a {
		if b[i] != id {
			return false
		}
	}
	return true
}

func cloneBlockSet(in map[BlockID]bool) map[BlockID]bool {
	out := make(map[BlockID]bool, len(in))
	for id, present := range in {
		out[id] = present
	}
	return out
}
