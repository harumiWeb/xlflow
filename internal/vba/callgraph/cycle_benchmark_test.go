package callgraph

import (
	"context"
	"fmt"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// BenchmarkCycleDetection compares the former exhaustive detector with the
// bounded detector on the same dense SCC and exercises larger linear-scale
// workloads through the bounded API only. Fixture construction is outside the
// timed region so the benchmark measures cycle detection.
func BenchmarkCycleDetection(b *testing.B) {
	b.Run("exhaustive-baseline/dense-scc-8", func(b *testing.B) {
		benchmarkCycleDetector(b, completeCycleSnapshot(8), func(ctx context.Context, g graph) (int, error) {
			nodes, edges := graphCycleInputs(g)
			cycles, err := enumerateCycles(ctx, nodes, edges, nil)
			return len(cycles), err
		}, -1)
	})
	b.Run("bounded-scc/dense-scc-8", func(b *testing.B) {
		benchmarkCycleDetector(b, completeCycleSnapshot(8), boundedDetector(1), 1)
	})
	for _, size := range []int{100, 250, 1000, 2000} {
		size := size
		b.Run(fmt.Sprintf("bounded-scc/dense-scc-%d", size), func(b *testing.B) {
			benchmarkCycleDetector(b, denseCycleSnapshot(size), boundedDetector(1), 1)
		})
	}
	for _, size := range []int{1000, 2000} {
		size := size
		b.Run(fmt.Sprintf("bounded-scc/ring-%d", size), func(b *testing.B) {
			benchmarkCycleDetector(b, ringSnapshot(size), boundedDetector(1), 1)
		})
		b.Run(fmt.Sprintf("bounded-scc/dag-%d", size), func(b *testing.B) {
			benchmarkCycleDetector(b, dagSnapshot(size), boundedDetector(0), 0)
		})
	}
}

func boundedDetector(wantComponents int) func(context.Context, graph) (int, error) {
	return func(ctx context.Context, g graph) (int, error) {
		components, err := findCyclicComponentsGraphContext(ctx, g)
		if err != nil {
			return 0, err
		}
		if len(components) != wantComponents {
			return 0, fmt.Errorf("bounded detector returned %d components, want %d", len(components), wantComponents)
		}
		return len(components), nil
	}
}

func benchmarkCycleDetector(b *testing.B, input Snapshot, detect func(context.Context, graph) (int, error), components int) {
	b.Helper()
	g := build(input)
	b.ReportAllocs()
	b.ResetTimer()
	var outputs int
	for i := 0; i < b.N; i++ {
		var err error
		outputs, err = detect(context.Background(), g)
		if err != nil {
			b.Fatal(err)
		}
	}
	b.StopTimer()
	b.ReportMetric(float64(len(input.Symbols)), "procedures/op")
	b.ReportMetric(float64(len(input.Calls)), "calls/op")
	if components >= 0 {
		b.ReportMetric(float64(components), "components/op")
	} else {
		b.ReportMetric(float64(outputs), "cycles/op")
	}
}

func graphCycleInputs(g graph) (map[string]ID, []Edge) {
	nodes := make(map[string]ID, len(g.nodes))
	allEdges := make([]Edge, 0)
	for key, node := range g.nodes {
		nodes[key] = node.ID
	}
	for _, edges := range g.out {
		allEdges = append(allEdges, edges...)
	}
	return nodes, allEdges
}

func completeCycleSnapshot(size int) Snapshot {
	symbolsList := make([]symbols.Symbol, 0, size)
	callsList := make([]calls.Call, 0, size*(size-1))
	for i := 0; i < size; i++ {
		name := fmt.Sprintf("P%04d", i)
		symbolsList = append(symbolsList, symbols.Symbol{Name: name, Kind: "sub", Module: "Dense", File: "Dense.bas", StartLine: i + 1, StartColumn: 1})
	}
	line := 1
	for from := 0; from < size; from++ {
		for to := 0; to < size; to++ {
			if from == to {
				continue
			}
			fromName := fmt.Sprintf("P%04d", from)
			toName := fmt.Sprintf("P%04d", to)
			callsList = append(callsList, matched("Dense", "Dense.bas", fromName, "Dense", "Dense.bas", toName, to+1, line))
			line++
		}
	}
	return SnapshotFromResult(&calls.Result{Symbols: symbolsList, Calls: callsList})
}

func denseCycleSnapshot(size int) Snapshot {
	const fanout = 8
	symbolsList := make([]symbols.Symbol, 0, size)
	callsList := make([]calls.Call, 0, size*fanout)
	line := 1
	for i := 0; i < size; i++ {
		name := fmt.Sprintf("P%04d", i)
		symbolsList = append(symbolsList, symbols.Symbol{Name: name, Kind: "sub", Module: "Dense", File: "Dense.bas", StartLine: i + 1, StartColumn: 1})
		for offset := 1; offset <= fanout; offset++ {
			to := (i + offset) % size
			callee := fmt.Sprintf("P%04d", to)
			callsList = append(callsList, matched("Dense", "Dense.bas", name, "Dense", "Dense.bas", callee, to+1, line))
			line++
		}
	}
	return SnapshotFromResult(&calls.Result{Symbols: symbolsList, Calls: callsList})
}

func ringSnapshot(size int) Snapshot {
	symbolsList := make([]symbols.Symbol, 0, size)
	callsList := make([]calls.Call, 0, size)
	for i := 0; i < size; i++ {
		name := fmt.Sprintf("P%04d", i)
		symbolsList = append(symbolsList, symbols.Symbol{Name: name, Kind: "sub", Module: "Ring", File: "Ring.bas", StartLine: i + 1, StartColumn: 1})
		callee := fmt.Sprintf("P%04d", (i+1)%size)
		callsList = append(callsList, matched("Ring", "Ring.bas", name, "Ring", "Ring.bas", callee, ((i+1)%size)+1, i+1))
	}
	return SnapshotFromResult(&calls.Result{Symbols: symbolsList, Calls: callsList})
}

func dagSnapshot(size int) Snapshot {
	symbolsList := make([]symbols.Symbol, 0, size)
	callsList := make([]calls.Call, 0, size-1)
	for i := 0; i < size; i++ {
		name := fmt.Sprintf("P%04d", i)
		symbolsList = append(symbolsList, symbols.Symbol{Name: name, Kind: "sub", Module: "DAG", File: "DAG.bas", StartLine: i + 1, StartColumn: 1})
		if i+1 < size {
			callee := fmt.Sprintf("P%04d", i+1)
			callsList = append(callsList, matched("DAG", "DAG.bas", name, "DAG", "DAG.bas", callee, i+2, i+1))
		}
	}
	return SnapshotFromResult(&calls.Result{Symbols: symbolsList, Calls: callsList})
}
