package cfg

import (
	"strconv"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var (
	benchmarkBlockSink       Block
	benchmarkBlockByIDSink   Block
	benchmarkIsReachableSink bool
	benchmarkReachableSink   []BlockID
	benchmarkPredecessorSink []BlockID
	benchmarkOutgoingSink    []Edge
	benchmarkViewSink        CFGView
	benchmarkQueryIndexSink  *queryIndex
)

func BenchmarkCFGQuery(b *testing.B) {
	for _, size := range []int{100, 1000, 5000} {
		graph := benchmarkQueryGraph(size)
		b.Run("indexed/"+benchmarkSizeName(size), func(b *testing.B) {
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				benchmarkBlockSink, _ = graph.BlockForStatement(i%size + 1)
				benchmarkReachableSink = graph.Reachable(EdgeFilter{})
				benchmarkPredecessorSink = graph.predecessors(BlockID(size+5), EdgeFilter{}, benchmarkSet(benchmarkReachableSink))
			}
		})
		b.Run("legacy/"+benchmarkSizeName(size), func(b *testing.B) {
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				benchmarkBlockSink, _ = legacyBlockForStatement(graph, i%size+1)
				benchmarkReachableSink = legacyReachable(graph, EdgeFilter{})
				benchmarkPredecessorSink = legacyPredecessors(graph, BlockID(size+5), EdgeFilter{}, benchmarkSet(benchmarkReachableSink))
			}
		})
	}
}

func BenchmarkCFGIsReachable(b *testing.B) {
	for _, size := range []int{100, 1000, 5000} {
		graph := benchmarkQueryGraph(size)
		for _, test := range []struct {
			name   string
			filter EdgeFilter
		}{
			{name: "default", filter: EdgeFilter{}},
			{name: "normal-only", filter: EdgeFilter{NormalOnly: true}},
		} {
			test := test
			b.Run("indexed/"+test.name+"/"+benchmarkSizeName(size), func(b *testing.B) {
				b.ReportAllocs()
				b.ResetTimer()
				for i := 0; i < b.N; i++ {
					target := BlockID(i%size + 6)
					benchmarkIsReachableSink = graph.IsReachable(target, test.filter)
				}
			})
			b.Run("legacy-scan/"+test.name+"/"+benchmarkSizeName(size), func(b *testing.B) {
				b.ReportAllocs()
				b.ResetTimer()
				for i := 0; i < b.N; i++ {
					target := BlockID(i%size + 6)
					benchmarkIsReachableSink = legacyIsReachable(graph, target, test.filter)
				}
			})
		}
	}
}

func BenchmarkCFGBlockForStatement(b *testing.B) {
	const size = 5000
	graph := benchmarkQueryGraph(size)
	b.Run("indexed", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			benchmarkBlockSink, _ = graph.BlockForStatement(i%size + 1)
		}
	})
	b.Run("legacy", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			benchmarkBlockSink, _ = legacyBlockForStatement(graph, i%size+1)
		}
	})
}

func BenchmarkCFGIndexedLookup(b *testing.B) {
	const size = 5000
	graph := benchmarkQueryGraph(size)
	b.Run("block-by-id", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			benchmarkBlockByIDSink, _ = graph.BlockByID(BlockID(i%size + 6))
		}
	})
	b.Run("outgoing-edges", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			benchmarkOutgoingSink = graph.OutgoingEdges(BlockID(i%size + 6))
		}
	})
}

func BenchmarkBuildQueryIndex(b *testing.B) {
	for _, size := range []int{100, 1000, 5000} {
		graph := benchmarkQueryGraph(size)
		b.Run("build/"+benchmarkSizeName(size), func(b *testing.B) {
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				benchmarkQueryIndexSink = buildQueryIndex(graph)
			}
		})
	}
}

// BenchmarkCFGViewReuse contrasts the compatibility materialization boundary
// with the immutable view used by migrated analyzers. The view path performs
// one lazy reachability/dominator build and reuses it for every query; the
// materialized path copies graph storage and rebuilds its revision on every
// iteration.
func BenchmarkCFGViewReuse(b *testing.B) {
	graph := benchmarkRaiseQueryGraph(5000)
	b.Run("materialized", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			filtered := graph.WithoutNormalErrRaiseContinuation()
			benchmarkIsReachableSink = filtered.IsReachable(filtered.NormalExit, EdgeFilter{})
			benchmarkViewSink = filtered.View(EdgeFilter{})
		}
	})
	b.Run("shared-view", func(b *testing.B) {
		view := graph.WithoutNormalErrRaiseContinuationView()
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			benchmarkIsReachableSink = view.IsReachable(view.NormalExit())
			benchmarkViewSink = view
		}
	})
}

func benchmarkSizeName(size int) string {
	return strconv.Itoa(size)
}

func benchmarkQueryGraph(size int) Graph {
	blocks := make([]Block, 0, size+5)
	blocks = append(blocks,
		Block{ID: 1, Kind: BlockEntry},
		Block{ID: 2, Kind: BlockNormalExit},
		Block{ID: 3, Kind: BlockExceptionalExit},
		Block{ID: 4, Kind: BlockTerminationExit},
		Block{ID: 5, Kind: BlockUnknownExit},
	)
	for i := 0; i < size; i++ {
		blocks = append(blocks, Block{ID: BlockID(i + 6), Kind: BlockStatement, StatementID: i + 1})
	}
	edges := []Edge{
		{From: 1, To: 6, Class: EdgeNormal},
		{From: 6, To: 3, Class: EdgeExceptional},
	}
	for i := 0; i < size; i++ {
		from := BlockID(i + 6)
		if i+1 < size {
			edges = append(edges, Edge{From: from, To: from + 1, Class: EdgeNormal})
		}
		if i+2 < size {
			edges = append(edges, Edge{From: from, To: from + 2, Class: EdgeNormal})
		}
	}
	edges = append(edges, Edge{From: BlockID(size + 5), To: 2, Class: EdgeNormal})
	graph := Graph{Blocks: blocks, Edges: edges, Entry: 1, NormalExit: 2, ExceptionalExit: 3, TerminationExit: 4, UnknownExit: 5}
	graph.query = buildQueryIndex(graph)
	return graph
}

func benchmarkRaiseQueryGraph(size int) Graph {
	graph := benchmarkQueryGraph(size)
	if len(graph.Blocks) > 5 {
		graph.Blocks[5].Statement = &procedureir.Statement{Text: "Err.Raise 5"}
	}
	graph.query = buildQueryIndex(graph)
	return graph
}

func legacyIsReachable(g Graph, target BlockID, filter EdgeFilter) bool {
	visited := map[BlockID]bool{g.Entry: true}
	queue := []BlockID{g.Entry}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for _, edge := range g.Edges {
			if edge.From != current || !filter.accepts(edge) || visited[edge.To] {
				continue
			}
			visited[edge.To] = true
			queue = append(queue, edge.To)
		}
	}
	return visited[target]
}

func benchmarkSet(ids []BlockID) map[BlockID]bool {
	set := make(map[BlockID]bool, len(ids))
	for _, id := range ids {
		set[id] = true
	}
	return set
}

func legacyBlockForStatement(g Graph, statementID int) (Block, bool) {
	for _, block := range g.Blocks {
		if block.Kind == BlockStatement && block.StatementID == statementID {
			return block, true
		}
	}
	return Block{}, false
}

func legacyReachable(g Graph, filter EdgeFilter) []BlockID {
	visited := map[BlockID]bool{g.Entry: true}
	queue := []BlockID{g.Entry}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for _, edge := range g.Edges {
			if edge.From != current || !filter.accepts(edge) {
				continue
			}
			if visited[edge.To] {
				continue
			}
			visited[edge.To] = true
			queue = append(queue, edge.To)
		}
	}
	out := make([]BlockID, 0, len(visited))
	for _, block := range g.Blocks {
		if visited[block.ID] {
			out = append(out, block.ID)
		}
	}
	return out
}

func legacyPredecessors(g Graph, id BlockID, filter EdgeFilter, reachable map[BlockID]bool) []BlockID {
	seen := map[BlockID]bool{}
	for _, edge := range g.Edges {
		if edge.To == id && filter.accepts(edge) && reachable[edge.From] {
			seen[edge.From] = true
		}
	}
	out := make([]BlockID, 0, len(seen))
	for predecessor := range seen {
		out = append(out, predecessor)
	}
	return out
}
