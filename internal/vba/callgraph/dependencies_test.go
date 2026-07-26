package callgraph

import (
	"reflect"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
)

func TestDependenciesProjectsCallsTypesAndUncertaintyDeterministically(t *testing.T) {
	snapshot := Snapshot{
		Symbols: []Symbol{
			{Name: "A", Kind: "module", Module: "A", ModuleKind: "standard", File: "A.bas", Line: 1, Column: 1},
			{Name: "Run", Kind: "sub", Module: "A", ModuleKind: "standard", File: "A.bas", Line: 2, Column: 1},
			{Name: "B", Kind: "module", Module: "B", ModuleKind: "standard", File: "B.bas", Line: 1, Column: 1},
			{Name: "Work", Kind: "sub", Module: "B", ModuleKind: "standard", File: "B.bas", Line: 2, Column: 1},
			{Name: "Service", Kind: "module", Module: "Service", ModuleKind: "class", File: "Service.cls", Line: 1, Column: 1},
			{Name: "Interfaces", Kind: "module", Module: "Interfaces", ModuleKind: "standard", File: "Interfaces.bas", Line: 1, Column: 1},
			{Name: "IFoo", Kind: "type", Module: "Interfaces", ModuleKind: "standard", File: "Interfaces.bas", Line: 2, Column: 1},
		},
		Calls: []calls.Call{
			dependencyCall("A", "A.bas", "Run", "B", "B.bas", "Work", "matched", 3),
			dependencyCall("B", "B.bas", "Work", "A", "A.bas", "Run", "matched", 3),
			dependencyCall("B", "B.bas", "Work", "B", "B.bas", "Work", "matched", 4),
			dependencyCall("A", "A.bas", "Run", "", "", "Missing", "unresolved", 4),
		},
		TypeReferences: []calls.TypeReference{
			{Kind: "uses_type", File: "A.bas", Module: "A", Target: "Service", Range: vbaast.Range{StartLine: 5, StartColumn: 12}},
			{Kind: "constructs", File: "A.bas", Module: "A", Target: "Service", Range: vbaast.Range{StartLine: 6, StartColumn: 19}},
			{Kind: "implements", File: "Service.cls", Module: "Service", Target: "IFoo", Range: vbaast.Range{StartLine: 2, StartColumn: 12}},
		},
	}

	first := Dependencies(snapshot, DependencyRequest{})
	if len(first.Nodes) != 8 || len(first.Edges) != 9 || len(first.UncertainEdges) != 1 {
		t.Fatalf("projection sizes = nodes:%d edges:%d uncertain:%d; result=%+v", len(first.Nodes), len(first.Edges), len(first.UncertainEdges), first)
	}
	if !hasDependencyEdge(first.Edges, "calls", "procedure|a|a.run|sub|A.bas|2|1", "procedure|b|b.work|sub|B.bas|2|1") {
		t.Fatalf("procedure call projection missing: %+v", first.Edges)
	}
	if !hasDependencyEdgeKind(first.Edges, "uses_type") || !hasDependencyEdgeKind(first.Edges, "constructs") || !hasDependencyEdgeKind(first.Edges, "implements") {
		t.Fatalf("type projections missing: %+v", first.Edges)
	}
	if !hasDependencyEdge(first.Edges, "calls", "module|b|B.bas", "module|b|B.bas") {
		t.Fatalf("self dependency missing: %+v", first.Edges)
	}
	if first.UncertainEdges[0].Status != "unresolved" || first.UncertainEdges[0].Callee != "Missing" {
		t.Fatalf("uncertainty = %+v", first.UncertainEdges)
	}
	for i := 0; i < 10; i++ {
		if got := Dependencies(snapshot, DependencyRequest{}); !reflect.DeepEqual(first, got) {
			t.Fatalf("projection order changed on run %d: first=%+v got=%+v", i, first, got)
		}
	}

	filtered := Dependencies(snapshot, DependencyRequest{Module: "a"})
	if len(filtered.Edges) != 4 || len(filtered.Nodes) != 5 {
		t.Fatalf("module filter = nodes:%d edges:%d result=%+v", len(filtered.Nodes), len(filtered.Edges), filtered)
	}
	if got := Dependencies(snapshot, DependencyRequest{Module: "missing"}); len(got.Nodes) != 0 || len(got.Edges) != 0 || len(got.UncertainEdges) != 0 {
		t.Fatalf("empty filter = %+v", got)
	}
}

func TestDependenciesRejectsQualifiedExternalTypeReference(t *testing.T) {
	result := Dependencies(Snapshot{
		Symbols: []Symbol{
			{Name: "Main", Kind: "module", Module: "Main", ModuleKind: "standard", File: "Main.bas", Line: 1, Column: 1},
			{Name: "Service", Kind: "module", Module: "Service", ModuleKind: "class", File: "Service.cls", Line: 1, Column: 1},
		},
		TypeReferences: []calls.TypeReference{{Kind: "uses_type", File: "Main.bas", Module: "Main", Target: "External.Service", Range: vbaast.Range{StartLine: 2, StartColumn: 12}}},
	}, DependencyRequest{})
	if len(result.Edges) != 0 || len(result.UncertainEdges) != 1 || result.UncertainEdges[0].Status != "external" {
		t.Fatalf("qualified external reference = %+v", result)
	}
}

func dependencyCall(callerModule, callerFile, callerName, calleeModule, calleeFile, calleeName, status string, line int) calls.Call {
	call := calls.Call{CallSite: calls.CallSite{
		File: callerFile, Module: callerModule, Caller: &calls.Caller{Name: callerName, Kind: "sub", QualifiedName: callerModule + "." + callerName},
		Callee: calls.Callee{Text: calleeName, BaseName: calleeName}, Range: vbaast.Range{StartLine: line, StartColumn: 2},
	}, Resolution: calls.Resolution{Status: status}}
	if status == "matched" {
		call.Resolution.Candidates = []calls.Candidate{{QualifiedName: calleeModule + "." + calleeName, Kind: "sub", File: calleeFile, Line: 2}}
	}
	return call
}

func hasDependencyEdge(edges []DependencyEdge, kind, from, to string) bool {
	for _, edge := range edges {
		if edge.Kind == kind && edge.From == from && edge.To == to {
			return true
		}
	}
	return false
}

func hasDependencyEdgeKind(edges []DependencyEdge, kind string) bool {
	for _, edge := range edges {
		if edge.Kind == kind {
			return true
		}
	}
	return false
}
