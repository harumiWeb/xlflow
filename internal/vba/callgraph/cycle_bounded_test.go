package callgraph

import (
	"context"
	"errors"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestFindCyclicComponentsReportsOneDeterministicWitnessPerSCC(t *testing.T) {
	input := SnapshotFromResult(&calls.Result{
		Symbols: []symbols.Symbol{
			symbol("M", "M.bas", "Self", 1),
			symbol("M", "M.bas", "A", 3),
			symbol("M", "M.bas", "B", 4),
			symbol("M", "M.bas", "C", 5),
			symbol("M", "M.bas", "Leaf", 7),
		},
		Calls: []calls.Call{
			matched("M", "M.bas", "Self", "M", "M.bas", "Self", 1, 1),
			matched("M", "M.bas", "A", "M", "M.bas", "B", 4, 1),
			matched("M", "M.bas", "B", "M", "M.bas", "C", 5, 1),
			matched("M", "M.bas", "C", "M", "M.bas", "A", 3, 1),
			matched("M", "M.bas", "A", "M", "M.bas", "C", 5, 1),
			matched("M", "M.bas", "C", "M", "M.bas", "A", 3, 1),
			matched("M", "M.bas", "Leaf", "M", "M.bas", "Leaf", 7, 1),
		},
	})
	components, err := FindCyclicComponentsContext(context.Background(), input)
	if err != nil {
		t.Fatal(err)
	}
	if len(components) != 3 {
		t.Fatalf("components = %#v, want self SCC, three-node SCC, and leaf self SCC", components)
	}
	if got := cycleKey(components[0].Witness.Nodes); got != "m|m.a|sub|M.bas|3|1,m|m.b|sub|M.bas|4|1,m|m.c|sub|M.bas|5|1" {
		t.Fatalf("first witness = %q", got)
	}
	if len(components[1].Witness.Nodes) != 1 || components[1].Witness.Nodes[0].QualifiedName != "M.Leaf" {
		t.Fatalf("leaf witness = %#v", components[1].Witness)
	}
	if len(components[2].Witness.Nodes) != 1 || components[2].Witness.Nodes[0].QualifiedName != "M.Self" {
		t.Fatalf("self witness = %#v", components[2].Witness)
	}
	if len(components[0].Nodes) != 3 || len(components[0].Witness.Edges) != len(components[0].Witness.Nodes) {
		t.Fatalf("component/witness shape = %#v", components[0])
	}
	for i, edge := range components[0].Witness.Edges {
		if edge.Caller != components[0].Witness.Nodes[i] || edge.Callee != components[0].Witness.Nodes[(i+1)%len(components[0].Witness.Nodes)] {
			t.Fatalf("witness edge %d not aligned: %#v", i, components[0].Witness)
		}
	}
}

func TestFindCyclicComponentsIsStableAcrossInputPermutationAndParallelEdges(t *testing.T) {
	input := SnapshotFromResult(&calls.Result{
		Symbols: []symbols.Symbol{symbol("M", "M.bas", "A", 1), symbol("M", "M.bas", "B", 2), symbol("M", "M.bas", "C", 3)},
		Calls: []calls.Call{
			matched("M", "M.bas", "A", "M", "M.bas", "B", 2, 1),
			matched("M", "M.bas", "A", "M", "M.bas", "B", 2, 2),
			matched("M", "M.bas", "B", "M", "M.bas", "C", 3, 3),
			matched("M", "M.bas", "C", "M", "M.bas", "A", 1, 4),
		},
	})
	first, err := FindCyclicComponentsContext(context.Background(), input)
	if err != nil {
		t.Fatal(err)
	}
	permuted := input
	permuted.Symbols = append([]Symbol(nil), input.Symbols...)
	permuted.Calls = append([]calls.Call(nil), input.Calls...)
	for left, right := 0, len(permuted.Symbols)-1; left < right; left, right = left+1, right-1 {
		permuted.Symbols[left], permuted.Symbols[right] = permuted.Symbols[right], permuted.Symbols[left]
	}
	for left, right := 0, len(permuted.Calls)-1; left < right; left, right = left+1, right-1 {
		permuted.Calls[left], permuted.Calls[right] = permuted.Calls[right], permuted.Calls[left]
	}
	second, err := FindCyclicComponentsContext(context.Background(), permuted)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first, second) {
		t.Fatalf("components changed after input permutation: first=%#v second=%#v", first, second)
	}
	if len(first) != 1 || len(first[0].Witness.Edges) != 3 || first[0].Witness.Edges[0].Location.StartLine != 1 {
		t.Fatalf("parallel endpoint/witness = %#v", first)
	}
}

func TestFindCyclicComponentsDiscardsPartialResultsOnCancellation(t *testing.T) {
	input := ringSnapshot(200)
	components, err := FindCyclicComponentsContext(&cancelAfterChecksContext{remaining: 3}, input)
	if !errors.Is(err, context.Canceled) || components != nil {
		t.Fatalf("canceled bounded result = (%#v, %v), want nil and context.Canceled", components, err)
	}
}

func TestFindCyclicComponentsHandlesLargeRingAndDAG(t *testing.T) {
	for _, test := range []struct {
		name       string
		input      Snapshot
		components int
	}{
		{name: "ring-2000", input: ringSnapshot(2000), components: 1},
		{name: "dag-2000", input: dagSnapshot(2000), components: 0},
	} {
		t.Run(test.name, func(t *testing.T) {
			components, err := FindCyclicComponentsContext(context.Background(), test.input)
			if err != nil {
				t.Fatal(err)
			}
			if len(components) != test.components {
				t.Fatalf("components = %d, want %d", len(components), test.components)
			}
		})
	}
}
