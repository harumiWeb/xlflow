package dataflow

import (
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func parseCharsProcedure(tb testing.TB) (procedureir.ProcedureIR, cfg.Graph) {
	tb.Helper()
	path := filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "projects", "third_party", "vba-fast-json", "src", "LibJSON.bas")
	source, err := os.ReadFile(path)
	if err != nil {
		tb.Fatal(err)
	}
	document, err := procedureir.BuildSource(procedureir.BuildOptions{Path: path, ModuleKind: "standard"}, source)
	if err != nil {
		tb.Fatal(err)
	}
	for _, procedure := range document.Procedures {
		if procedure.Symbol.Name == "ParseChars" {
			return procedure, cfg.Build(procedure)
		}
	}
	tb.Fatal("ParseChars procedure not found")
	return procedureir.ProcedureIR{}, cfg.Graph{}
}

func TestAnalyzeProcedureParseCharsConvergesWithBoundedTransfers(t *testing.T) {
	t.Parallel()
	procedure, graph := parseCharsProcedure(t)
	analyzer := newProcedureAnalyzer(procedure, graph, Options{})
	result, stats := analyzer.runWithStats()
	reachableBlocks := len(graph.Reachable(cfg.EdgeFilter{}))
	if stats.transfers > 5*reachableBlocks {
		t.Fatalf("worklist transfers = %d for %d reachable blocks, want at most five transfers per block", stats.transfers, reachableBlocks)
	}
	if len(result.States) != reachableBlocks {
		t.Fatalf("analyzed states = %d, want one for each of %d reachable blocks", len(result.States), reachableBlocks)
	}
	repeated := newProcedureAnalyzer(procedure, graph, Options{})
	repeatedResult, repeatedStats := repeated.runWithStats()
	if repeatedStats.transfers != stats.transfers || !reflect.DeepEqual(repeatedResult, result) {
		t.Fatalf("repeated analysis was not deterministic: transfers %d then %d", stats.transfers, repeatedStats.transfers)
	}
}

func BenchmarkAnalyzeProcedureParseChars(b *testing.B) {
	procedure, graph := parseCharsProcedure(b)
	reachableBlocks := len(graph.Reachable(cfg.EdgeFilter{}))
	b.ReportAllocs()
	b.ResetTimer()
	for range b.N {
		result, stats := newProcedureAnalyzer(procedure, graph, Options{}).runWithStats()
		if stats.transfers > 5*reachableBlocks || len(result.States) != reachableBlocks {
			b.Fatalf("analysis changed during benchmark: transfers=%d states=%d reachable=%d", stats.transfers, len(result.States), reachableBlocks)
		}
	}
}

func TestRankedWorklistHandlesDeepWideGraphDeterministically(t *testing.T) {
	t.Parallel()
	const (
		depth = 8000
		width = 4000
	)
	blocks := make([]cfg.Block, 0, depth+width)
	edges := make([]cfg.Edge, 0, depth+width-1)
	reachable := make(map[cfg.BlockID]bool, depth+width)
	for id := 1; id <= depth+width; id++ {
		blockID := cfg.BlockID(id)
		blocks = append(blocks, cfg.Block{ID: blockID})
		reachable[blockID] = true
	}
	for id := 1; id < depth; id++ {
		edges = append(edges, cfg.Edge{ID: cfg.EdgeID(len(edges) + 1), From: cfg.BlockID(id), To: cfg.BlockID(id + 1)})
	}
	for offset := 1; offset <= width; offset++ {
		edges = append(edges, cfg.Edge{ID: cfg.EdgeID(len(edges) + 1), From: cfg.BlockID(depth), To: cfg.BlockID(depth + offset)})
	}
	graph := cfg.Graph{Blocks: blocks, Edges: edges, Entry: 1}
	analyzer := newProcedureAnalyzer(procedureir.ProcedureIR{}, graph, Options{})
	rank := analyzer.reversePostOrderRank(reachable)
	if len(rank) != depth+width || rank[graph.Entry] != 0 {
		t.Fatalf("reverse-postorder ranks = %d with entry rank %d", len(rank), rank[graph.Entry])
	}

	reversedGraph := graph
	reversedGraph.Edges = append([]cfg.Edge(nil), graph.Edges...)
	for left, right := 0, len(reversedGraph.Edges)-1; left < right; left, right = left+1, right-1 {
		reversedGraph.Edges[left], reversedGraph.Edges[right] = reversedGraph.Edges[right], reversedGraph.Edges[left]
	}
	reversed := newProcedureAnalyzer(procedureir.ProcedureIR{}, reversedGraph, Options{})
	if reversedRank := reversed.reversePostOrderRank(reachable); !reflect.DeepEqual(reversedRank, rank) {
		t.Fatal("reverse-postorder ranks depend on input edge order")
	}

	result, stats := analyzer.runWithStats()
	if stats.transfers != depth+width || len(result.States) != depth+width {
		t.Fatalf("analysis processed %d transfers and %d states, want %d each", stats.transfers, len(result.States), depth+width)
	}
}
