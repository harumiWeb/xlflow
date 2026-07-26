package callgraph

import (
	"errors"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestAnalyzeTraversesBothDirectionsCyclesAndUncertainty(t *testing.T) {
	input := &calls.Result{
		Symbols: []symbols.Symbol{symbol("A", "A.bas", "Run", 1), symbol("B", "B.cls", "Work", 3), symbol("C", "C.bas", "Finish", 5)},
		Calls: []calls.Call{
			matched("A", "A.bas", "Run", "B", "B.cls", "Work", 3, 2),
			matched("B", "B.cls", "Work", "C", "C.bas", "Finish", 5, 4),
			matched("C", "C.bas", "Finish", "C", "C.bas", "Finish", 5, 6),
			unresolved("B", "B.cls", "Work", 7),
		},
	}

	got, err := Analyze(input, Request{Target: "B.Work", Direction: DirectionBoth, Depth: 2})
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Nodes) != 3 || len(got.Edges) != 3 {
		t.Fatalf("confirmed graph = nodes:%d edges:%d", len(got.Nodes), len(got.Edges))
	}
	if len(got.DirectCallers) != 1 || got.DirectCallers[0].ID.QualifiedName != "A.Run" {
		t.Fatalf("direct callers = %#v", got.DirectCallers)
	}
	if len(got.DirectCallees) != 1 || got.DirectCallees[0].ID.QualifiedName != "C.Finish" {
		t.Fatalf("direct callees = %#v", got.DirectCallees)
	}
	if got.Uncertainty.Unresolved != 1 {
		t.Fatalf("uncertainty = %#v", got.Uncertainty)
	}
	if len(got.Cycles) != 1 || len(got.Cycles[0].Nodes) != 1 || got.Cycles[0].Nodes[0].QualifiedName != "C.Finish" {
		t.Fatalf("cycles = %#v", got.Cycles)
	}
	if len(got.AffectedModules) != 3 || got.AffectedModules[1].Kind != "class" {
		t.Fatalf("modules = %#v", got.AffectedModules)
	}
	callersOnly, err := Analyze(input, Request{Target: "B.Work", Direction: DirectionCallers, Depth: 1})
	if err != nil || len(callersOnly.DirectCallers) != 1 || len(callersOnly.DirectCallees) != 0 {
		t.Fatalf("callers-only impact = %#v, %v", callersOnly, err)
	}
}

func TestAnalyzeHonorsDirectionDepthAndAmbiguousTargets(t *testing.T) {
	input := &calls.Result{Symbols: []symbols.Symbol{symbol("A", "A.bas", "Run", 1), symbol("B", "B.bas", "Work", 1), symbol("B", "other/B.bas", "Work", 1)}, Calls: []calls.Call{matched("A", "A.bas", "Run", "B", "B.bas", "Work", 1, 2)}}
	got, err := Analyze(input, Request{Target: "A.Run", Direction: DirectionCallees, Depth: 0})
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Nodes) != 1 || len(got.Edges) != 0 || len(got.DirectCallees) != 0 {
		t.Fatalf("depth zero = %#v", got)
	}
	_, err = Analyze(input, Request{Target: "B.Work", Depth: 1})
	var ambiguous *AmbiguousTargetError
	if !errors.As(err, &ambiguous) || len(ambiguous.Candidates) != 2 {
		t.Fatalf("ambiguous error = %#v", err)
	}
	_, err = Analyze(input, Request{Target: "Missing.Run", Depth: 1})
	if !errors.Is(err, ErrTargetNotFound) {
		t.Fatalf("not found error = %#v", err)
	}
}

func TestAnalyzeDeduplicatesDiamondAndExcludesDisconnectedSymbols(t *testing.T) {
	input := &calls.Result{
		Symbols: []symbols.Symbol{
			symbol("A", "A.bas", "Run", 1), symbol("B", "B.bas", "Left", 1), symbol("C", "C.bas", "Right", 1),
			symbol("D", "D.bas", "Join", 1), symbol("Z", "Z.bas", "Disconnected", 1),
		},
		Calls: []calls.Call{
			matched("A", "A.bas", "Run", "B", "B.bas", "Left", 1, 2), matched("A", "A.bas", "Run", "C", "C.bas", "Right", 1, 3),
			matched("B", "B.bas", "Left", "D", "D.bas", "Join", 1, 2), matched("C", "C.bas", "Right", "D", "D.bas", "Join", 1, 2),
		},
	}
	got, err := Analyze(input, Request{Target: "A.Run", Direction: DirectionCallees, Depth: 2})
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Nodes) != 4 || len(got.Edges) != 4 || len(got.DownstreamCallees) != 1 || got.DownstreamCallees[0].ID.QualifiedName != "D.Join" {
		t.Fatalf("diamond impact = %#v", got)
	}
}

func symbol(module, file, name string, line int) symbols.Symbol {
	return symbols.Symbol{Name: name, Kind: "sub", Module: module, File: file, StartLine: line, StartColumn: 1}
}
func matched(callerModule, callerFile, callerName, calleeModule, calleeFile, calleeName string, calleeLine, callLine int) calls.Call {
	return calls.Call{CallSite: calls.CallSite{File: callerFile, Module: callerModule, Caller: &calls.Caller{Name: callerName, Kind: "sub", QualifiedName: callerModule + "." + callerName}, Range: vbaast.Range{StartLine: callLine, StartColumn: 1, EndLine: callLine, EndColumn: 8}}, Resolution: calls.Resolution{Status: "matched", Candidates: []calls.Candidate{{QualifiedName: calleeModule + "." + calleeName, Kind: "sub", File: calleeFile, Line: calleeLine}}}}
}
func unresolved(module, file, name string, callLine int) calls.Call {
	return calls.Call{CallSite: calls.CallSite{File: file, Module: module, Caller: &calls.Caller{Name: name, Kind: "sub", QualifiedName: module + "." + name}, Range: vbaast.Range{StartLine: callLine, StartColumn: 1}}, Resolution: calls.Resolution{Status: "unresolved"}}
}
