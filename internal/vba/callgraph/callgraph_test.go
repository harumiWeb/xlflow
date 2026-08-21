package callgraph

import (
	"context"
	"errors"
	"reflect"
	"testing"
	"time"

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

func TestAnalyzeKeepsAccessorCallersAndDeduplicatesDirectNodes(t *testing.T) {
	input := &calls.Result{
		Symbols: []symbols.Symbol{
			{Name: "Value", Kind: "property_get", Module: "Thing", File: "Thing.cls", StartLine: 1, StartColumn: 1},
			{Name: "Value", Kind: "property_let", Module: "Thing", File: "Thing.cls", StartLine: 5, StartColumn: 1},
			symbol("Thing", "Thing.cls", "Helper", 9),
		},
		Calls: []calls.Call{
			matchedKind("Thing", "Thing.cls", "Value", "property_get", "Thing", "Thing.cls", "Helper", "sub", 9, 2),
			matchedKind("Thing", "Thing.cls", "Value", "property_get", "Thing", "Thing.cls", "Helper", "sub", 9, 3),
		},
	}
	got, err := Analyze(input, Request{Target: "Thing.Helper", Direction: DirectionCallers, Depth: 1})
	if err != nil {
		t.Fatal(err)
	}
	if len(got.DirectCallers) != 1 || got.DirectCallers[0].ID.Kind != "property_get" || len(got.Edges) != 2 {
		t.Fatalf("accessor caller impact = %#v", got)
	}
}

func TestAnalyzeUsesConfiguredModuleKindsAndStableCycles(t *testing.T) {
	input := &calls.Result{
		ModuleKinds: map[string]string{"Sheet1.cls": "document", "FormCode.bas": "form"},
		Symbols: []symbols.Symbol{
			symbol("A", "A.bas", "Run", 1), symbol("B", "B.bas", "Work", 1), symbol("C", "C.bas", "Finish", 1),
			symbol("Sheet1", "Sheet1.cls", "Changed", 1), symbol("FormCode", "FormCode.bas", "Submit", 1),
		},
		Calls: []calls.Call{
			matched("A", "A.bas", "Run", "B", "B.bas", "Work", 1, 2), matched("B", "B.bas", "Work", "C", "C.bas", "Finish", 1, 2),
			matched("C", "C.bas", "Finish", "A", "A.bas", "Run", 1, 2), matched("A", "A.bas", "Run", "C", "C.bas", "Finish", 1, 3),
			matched("A", "A.bas", "Run", "Sheet1", "Sheet1.cls", "Changed", 1, 4), matched("Sheet1", "Sheet1.cls", "Changed", "FormCode", "FormCode.bas", "Submit", 1, 2),
		},
	}
	first, err := Analyze(input, Request{Target: "A.Run", Direction: DirectionCallees, Depth: 3})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []struct{ name, kind string }{{"Sheet1.Changed", "document"}, {"FormCode.Submit", "form"}} {
		found := false
		for _, node := range first.Nodes {
			if node.ID.QualifiedName == want.name && node.ModuleKind == want.kind {
				found = true
			}
		}
		if !found {
			t.Fatalf("configured module kind %s=%s missing from %#v", want.name, want.kind, first.Nodes)
		}
	}
	for i := 0; i < 20; i++ {
		again, err := Analyze(input, Request{Target: "A.Run", Direction: DirectionCallees, Depth: 3})
		if err != nil || !reflect.DeepEqual(first.Cycles, again.Cycles) {
			t.Fatalf("cycle output changed on run %d: first=%#v again=%#v err=%v", i, first.Cycles, again.Cycles, err)
		}
	}
}

func TestFindCyclesEnumeratesElementaryGraphWideCyclesDeterministically(t *testing.T) {
	names := []string{"Self", "MutualA", "MutualB", "OuterA", "OuterB", "OuterC", "IndependentA", "IndependentB"}
	symbolsList := make([]symbols.Symbol, 0, len(names))
	for i, name := range names {
		symbolsList = append(symbolsList, symbol("M", "M.bas", name, i+1))
	}
	input := SnapshotFromResult(&calls.Result{
		Symbols: symbolsList,
		Calls: []calls.Call{
			// Direct recursion.
			matched("M", "M.bas", "Self", "M", "M.bas", "Self", 1, 1),
			// Mutual recursion.
			matched("M", "M.bas", "MutualA", "M", "M.bas", "MutualB", 3, 2),
			matched("M", "M.bas", "MutualB", "M", "M.bas", "MutualA", 2, 3),
			// A three-procedure cycle and a nested two-procedure cycle sharing
			// the same entry point.
			matched("M", "M.bas", "OuterA", "M", "M.bas", "OuterB", 5, 4),
			matched("M", "M.bas", "OuterB", "M", "M.bas", "OuterA", 4, 5),
			matched("M", "M.bas", "OuterB", "M", "M.bas", "OuterC", 6, 6),
			matched("M", "M.bas", "OuterC", "M", "M.bas", "OuterA", 4, 7),
			// Independent cycle.
			matched("M", "M.bas", "IndependentA", "M", "M.bas", "IndependentB", 8, 8),
			matched("M", "M.bas", "IndependentB", "M", "M.bas", "IndependentA", 7, 9),
			// Parallel endpoint evidence: the earliest source call is retained
			// in the cycle while Analyze continues to expose both call sites.
			matched("M", "M.bas", "MutualA", "M", "M.bas", "MutualB", 2, 20),
			// Unresolved calls are uncertainty, not confirmed graph edges.
			unresolved("M", "M.bas", "Self", 21),
		},
	})

	cycles := FindCycles(input)
	if len(cycles) != 5 {
		t.Fatalf("cycles = %#v", cycles)
	}
	// cycleKey ordering is stable; compare the emitted qualified names and
	// verify no rotation duplicates are present.
	seen := map[string]bool{}
	for _, cycle := range cycles {
		if len(cycle.Nodes) != len(cycle.Edges) {
			t.Fatalf("cycle nodes/edges mismatch: %#v", cycle)
		}
		if len(cycle.Nodes) == 0 {
			t.Fatalf("empty cycle: %#v", cycle)
		}
		key := cycleKey(cycle.Nodes)
		if seen[key] {
			t.Fatalf("rotation duplicate: %#v", cycle)
		}
		seen[key] = true
		for i, edge := range cycle.Edges {
			if edge.Caller != cycle.Nodes[i] || edge.Callee != cycle.Nodes[(i+1)%len(cycle.Nodes)] {
				t.Fatalf("cycle edge %d not ordered with nodes: %#v", i, cycle)
			}
		}
		if cycle.Nodes[0].String() != minimumCycleNodeForTest(cycle.Nodes) {
			t.Fatalf("cycle is not canonically rotated: %#v", cycle)
		}
	}
	for _, expected := range []string{
		"m|m.independenta|sub|M.bas|7|1,m|m.independentb|sub|M.bas|8|1",
		"m|m.mutuala|sub|M.bas|2|1,m|m.mutualb|sub|M.bas|3|1",
		"m|m.outera|sub|M.bas|4|1,m|m.outerb|sub|M.bas|5|1",
		"m|m.outera|sub|M.bas|4|1,m|m.outerb|sub|M.bas|5|1,m|m.outerc|sub|M.bas|6|1",
		"m|m.self|sub|M.bas|1|1",
	} {
		if !seen[expected] {
			t.Fatalf("missing cycle %q in %#v", expected, seen)
		}
	}
	for _, cycle := range cycles {
		for _, edge := range cycle.Edges {
			if edge.Caller.QualifiedName == "M.MutualA" && edge.Callee.QualifiedName == "M.MutualB" && edge.Location.StartLine != 2 {
				t.Fatalf("parallel endpoint did not retain earliest edge: %#v", edge)
			}
		}
	}
	impact, err := AnalyzeSnapshot(input, Request{Target: "M.Self", Direction: DirectionBoth, Depth: 1})
	if err != nil || impact.Uncertainty.Unresolved != 1 {
		t.Fatalf("unresolved call uncertainty = %#v, err=%v", impact.Uncertainty, err)
	}
	again := FindCycles(input)
	if !reflect.DeepEqual(cycles, again) {
		t.Fatalf("cycle output changed: first=%#v again=%#v", cycles, again)
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
	if got := FindCycles(permuted); !reflect.DeepEqual(cycles, got) {
		t.Fatalf("cycle output changed after input permutation: first=%#v permuted=%#v", cycles, got)
	}
}

func minimumCycleNodeForTest(nodes []ID) string {
	minimum := nodes[0].String()
	for _, node := range nodes[1:] {
		if key := node.String(); key < minimum {
			minimum = key
		}
	}
	return minimum
}

func TestFindCyclesRetainsDistinctDirectedPathsWithSameNodes(t *testing.T) {
	input := SnapshotFromResult(&calls.Result{
		Symbols: []symbols.Symbol{
			symbol("M", "M.bas", "A", 1),
			symbol("M", "M.bas", "B", 2),
			symbol("M", "M.bas", "C", 3),
		},
		Calls: []calls.Call{
			matched("M", "M.bas", "A", "M", "M.bas", "B", 2, 1),
			matched("M", "M.bas", "B", "M", "M.bas", "C", 3, 1),
			matched("M", "M.bas", "C", "M", "M.bas", "A", 1, 1),
			matched("M", "M.bas", "A", "M", "M.bas", "C", 3, 2),
			matched("M", "M.bas", "C", "M", "M.bas", "B", 2, 2),
			matched("M", "M.bas", "B", "M", "M.bas", "A", 1, 2),
		},
	})
	cycles := FindCycles(input)
	var threeNode int
	paths := map[string]bool{}
	for _, cycle := range cycles {
		if len(cycle.Nodes) == 3 {
			threeNode++
			paths[cycleKey(cycle.Nodes)] = true
		}
	}
	if threeNode != 2 || len(paths) != 2 {
		t.Fatalf("cycles = %#v, want clockwise and reverse three-node paths", cycles)
	}
}

type cancelAfterChecksContext struct{ remaining int }

func (c *cancelAfterChecksContext) Deadline() (time.Time, bool) { return time.Time{}, false }
func (c *cancelAfterChecksContext) Done() <-chan struct{}       { return nil }
func (c *cancelAfterChecksContext) Value(any) any               { return nil }
func (c *cancelAfterChecksContext) Err() error {
	if c.remaining <= 0 {
		return context.Canceled
	}
	c.remaining--
	return nil
}

func TestFindCyclesContextDiscardsPartialResultsOnCancellation(t *testing.T) {
	input := SnapshotFromResult(&calls.Result{
		Symbols: []symbols.Symbol{
			symbol("M", "M.bas", "A", 1),
			symbol("M", "M.bas", "B", 2),
		},
		Calls: []calls.Call{
			matched("M", "M.bas", "A", "M", "M.bas", "B", 2, 1),
			matched("M", "M.bas", "B", "M", "M.bas", "A", 1, 1),
		},
	})
	cycles, err := FindCyclesContext(&cancelAfterChecksContext{remaining: 3}, input)
	if !errors.Is(err, context.Canceled) || cycles != nil {
		t.Fatalf("canceled cycle result = (%#v, %v), want nil and context.Canceled", cycles, err)
	}
}

func TestAnalyzeReachabilitySeparatesConfirmedPossibleAndUnreachableClusters(t *testing.T) {
	input := &calls.Result{
		Symbols: []symbols.Symbol{
			{Name: "Run", Kind: "sub", Visibility: "Public", Module: "Main", File: "Main.bas", StartLine: 1, StartColumn: 1},
			privateSymbol("Main", "Main.bas", "ConfirmedHelper", 4),
			privateSymbol("Main", "Main.bas", "DynamicOnly", 7),
			privateSymbol("Main", "Main.bas", "Isolated", 10),
			privateSymbol("Dead", "Dead.bas", "DeadCaller", 1),
			privateSymbol("Dead", "Dead.bas", "DeadTarget", 4),
		},
		Calls: []calls.Call{
			matched("Main", "Main.bas", "Run", "Main", "Main.bas", "ConfirmedHelper", 4, 2),
			matched("Dead", "Dead.bas", "DeadCaller", "Dead", "Dead.bas", "DeadTarget", 4, 2),
		},
		DynamicReferences: []calls.DynamicReference{
			{File: "Main.bas", Module: "Main", Caller: &calls.Caller{Name: "Run", Kind: "sub", QualifiedName: "Main.Run"}, API: "application.run", Target: "Main.DynamicOnly", Kind: "static"},
			{File: "Dead.bas", Module: "Dead", Caller: &calls.Caller{Name: "DeadCaller", Kind: "sub", QualifiedName: "Dead.DeadCaller"}, API: "application.onkey", Kind: "unknown"},
		},
	}

	got := AnalyzeReachability(SnapshotFromResult(input), ReachabilityRequest{
		Roots: []Root{{Target: "Main.Run", Confidence: RootConfirmed, Reason: "test"}},
	})
	if !hasReachabilityNode(got.Confirmed, "Main.Run") || !hasReachabilityNode(got.Confirmed, "Main.ConfirmedHelper") {
		t.Fatalf("confirmed reachability = %#v", got.Confirmed)
	}
	if !hasReachabilityNode(got.Possible, "Main.DynamicOnly") {
		t.Fatalf("possible dynamic target = %#v", got.Possible)
	}
	if hasReachabilityNode(got.Possible, "Dead.DeadTarget") || hasReachabilityNode(got.Confirmed, "Dead.DeadTarget") {
		t.Fatalf("unreachable dynamic caller leaked into reachability: possible=%#v confirmed=%#v", got.Possible, got.Confirmed)
	}
	if !hasReachabilityNode(got.Unreachable, "Main.Isolated") || !hasReachabilityNode(got.Unreachable, "Dead.DeadCaller") || !hasReachabilityNode(got.Unreachable, "Dead.DeadTarget") {
		t.Fatalf("unreachable private procedures = %#v", got.Unreachable)
	}
	if len(got.Clusters) != 2 || !hasClusterSize(got.Clusters, 1) || !hasClusterSize(got.Clusters, 2) {
		t.Fatalf("unreachable clusters = %#v", got.Clusters)
	}
}

func TestAnalyzeReachabilityTreatsAmbiguousRootsAsPossible(t *testing.T) {
	input := SnapshotFromResult(&calls.Result{Symbols: []symbols.Symbol{
		privateSymbol("One", "One.bas", "Run", 1),
		privateSymbol("Two", "Two.bas", "Run", 1),
	}})
	got := AnalyzeReachability(input, ReachabilityRequest{Roots: []Root{{Target: "Run", Confidence: RootConfirmed}}})
	if len(got.Confirmed) != 0 || len(got.Possible) != 2 || len(got.Unreachable) != 0 {
		t.Fatalf("ambiguous roots = confirmed:%#v possible:%#v unreachable:%#v", got.Confirmed, got.Possible, got.Unreachable)
	}
	input = SnapshotFromResult(&calls.Result{Symbols: []symbols.Symbol{
		privateSymbol("Only", "Only.bas", "Run", 1),
	}})
	got = AnalyzeReachability(input, ReachabilityRequest{Roots: []Root{{Target: "Run", Confidence: RootConfirmed}}})
	if len(got.Confirmed) != 1 || got.Confirmed[0].ID.QualifiedName != "Only.Run" || len(got.Possible) != 0 {
		t.Fatalf("unique unqualified root = confirmed:%#v possible:%#v", got.Confirmed, got.Possible)
	}
	input = SnapshotFromResult(&calls.Result{Symbols: []symbols.Symbol{
		privateSymbol("Other", "Other.bas", "Run", 1),
	}})
	got = AnalyzeReachability(input, ReachabilityRequest{Roots: []Root{{Target: "Missing.Run", Confidence: RootConfirmed}}})
	if len(got.Confirmed) != 0 || len(got.Possible) != 1 || got.Possible[0].ID.QualifiedName != "Other.Run" {
		t.Fatalf("unresolved qualified root = confirmed:%#v possible:%#v", got.Confirmed, got.Possible)
	}
}

func TestBuildIndexesCandidateIdentityWithoutChangingMatches(t *testing.T) {
	input := Snapshot{
		Symbols: []Symbol{
			{Name: "Caller", Kind: "sub", Module: "Main", File: "Main.bas", Line: 1, Column: 1},
			// Kelvin sign and ASCII k are EqualFold-equivalent but have different
			// lower-case representations. The index must retain that match.
			{Name: "Karget", Kind: "sub", Module: "M", File: "M.bas", Line: 5, Column: 1},
			// Same qualified name and declaration line in different files must
			// remain distinct candidate identities.
			{Name: "Target", Kind: "sub", Module: "M", File: "M.bas", Line: 8, Column: 1},
			{Name: "Target", Kind: "sub", Module: "M", File: "Other.bas", Line: 8, Column: 1},
			// Different columns make these distinct graph nodes but the candidate
			// identity is not unique, so no edge may be inferred.
			{Name: "Collision", Kind: "sub", Module: "M", File: "M.bas", Line: 11, Column: 1},
			{Name: "Collision", Kind: "sub", Module: "M", File: "M.bas", Line: 11, Column: 2},
		},
		Calls: []calls.Call{
			matched("Main", "Main.bas", "Caller", "M", "M.bas", "karget", 5, 2),
			matched("Main", "Main.bas", "Caller", "M", "Other.bas", "Target", 8, 3),
			// These candidates have the same qualified name but fail one of the
			// exact identity fields that the old graph scan checked.
			matchedKind("Main", "Main.bas", "Caller", "sub", "M", "M.bas", "Target", "function", 8, 4),
			matched("Main", "Main.bas", "Caller", "M", "Missing.bas", "Target", 8, 5),
			// File normalization is only an index aid; the historical exact file
			// comparison must still reject a cleaned-but-different candidate.
			matched("Main", "Main.bas", "Caller", "M", ".\\M.bas", "Target", 8, 6),
			matched("Main", "Main.bas", "Caller", "M", "M.bas", "Target", 9, 6),
			matched("Main", "Main.bas", "Caller", "M", "M.bas", "Collision", 11, 7),
			calls.Call{CallSite: calls.CallSite{File: "Main.bas", Module: "Main", Caller: &calls.Caller{Name: "Caller", Kind: "sub", QualifiedName: "Main.Caller"}, Range: vbaast.Range{StartLine: 8}}, Resolution: calls.Resolution{Status: "ambiguous", Candidates: []calls.Candidate{{QualifiedName: "M.Target", Kind: "sub", File: "M.bas", Line: 8}}}},
		},
	}

	g := build(input)
	callerKey, callerOK := g.lookupCallerKey("Main.Caller", "sub", "Main.bas")
	if !callerOK {
		t.Fatal("caller was not indexed")
	}
	if got := len(g.out[callerKey]); got != 2 {
		t.Fatalf("confirmed edges = %d, want 2", got)
	}
	if got := g.out[callerKey][0].Callee.QualifiedName; got != "M.Karget" {
		t.Fatalf("first edge callee = %q, want Unicode EqualFold target", got)
	}
	if got := g.out[callerKey][1].Callee.File; got != "Other.bas" {
		t.Fatalf("second edge file = %q, want Other.bas", got)
	}
	if got := len(g.uncertain[callerKey]); got != 1 {
		t.Fatalf("non-matched calls recorded as uncertain = %d, want 1", got)
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
	permutedGraph := build(permuted)
	permutedCallerKey, permutedCallerOK := permutedGraph.lookupCallerKey("Main.Caller", "sub", "Main.bas")
	if !permutedCallerOK {
		t.Fatal("caller was not indexed after input permutation")
	}
	if got := edgeKeys(permutedGraph.out[permutedCallerKey]); !reflect.DeepEqual(got, edgeKeys(g.out[callerKey])) {
		t.Fatalf("edge ordering changed after input permutation: first=%v permuted=%v", edgeKeys(g.out[callerKey]), got)
	}
}

func TestBuildSelectsExactCandidateFromNormalizedFileCollision(t *testing.T) {
	input := Snapshot{
		Symbols: []Symbol{
			{Name: "Caller", Kind: "sub", Module: "Main", File: "Main.bas", Line: 1, Column: 1},
			{Name: "Target", Kind: "sub", Module: "M", File: "src/M.bas", Line: 5, Column: 1},
			{Name: "Target", Kind: "sub", Module: "M", File: "src/./M.bas", Line: 5, Column: 2},
		},
		Calls: []calls.Call{
			matched("Main", "Main.bas", "Caller", "M", "src/M.bas", "Target", 5, 2),
			matched("Main", "Main.bas", "Caller", "M", "src/./M.bas", "Target", 5, 3),
		},
	}

	g := build(input)
	callerKey, ok := g.lookupCallerKey("Main.Caller", "sub", "Main.bas")
	if !ok {
		t.Fatal("caller was not indexed")
	}
	if got := len(g.out[callerKey]); got != 2 {
		t.Fatalf("confirmed edges = %d, want 2", got)
	}
	files := map[string]int{}
	for _, edge := range g.out[callerKey] {
		if edge.Callee.File != "src/M.bas" && edge.Callee.File != "src/./M.bas" {
			t.Fatalf("unexpected callee file = %q", edge.Callee.File)
		}
		files[edge.Callee.File]++
	}
	if files["src/M.bas"] != 1 || files["src/./M.bas"] != 1 {
		t.Fatalf("normalized collision targets = %#v, want one edge per raw file", files)
	}
}

func TestBuildIndexesPropertyAccessorKinds(t *testing.T) {
	input := Snapshot{
		Symbols: []Symbol{
			{Name: "Caller", Kind: "sub", Module: "Main", File: "Main.bas", Line: 1, Column: 1},
			{Name: "Value", Kind: "property_get", Module: "Thing", File: "Thing.cls", Line: 5, Column: 1},
			{Name: "Value", Kind: "property_let", Module: "Thing", File: "Thing.cls", Line: 10, Column: 1},
			{Name: "Value", Kind: "property_set", Module: "Thing", File: "Thing.cls", Line: 15, Column: 1},
		},
		Calls: []calls.Call{
			matchedKind("Main", "Main.bas", "Caller", "sub", "Thing", "Thing.cls", "Value", "property_get", 5, 2),
			matchedKind("Main", "Main.bas", "Caller", "sub", "Thing", "Thing.cls", "Value", "property_let", 10, 3),
			matchedKind("Main", "Main.bas", "Caller", "sub", "Thing", "Thing.cls", "Value", "property_set", 15, 4),
		},
	}

	g := build(input)
	callerKey, ok := g.lookupCallerKey("Main.Caller", "sub", "Main.bas")
	if !ok {
		t.Fatal("caller was not indexed")
	}
	if got := len(g.out[callerKey]); got != 3 {
		t.Fatalf("property accessor edges = %d, want 3", got)
	}
	gotKinds := make([]string, 0, len(g.out[callerKey]))
	for _, edge := range g.out[callerKey] {
		gotKinds = append(gotKinds, edge.Callee.Kind)
	}
	if want := []string{"property_get", "property_let", "property_set"}; !reflect.DeepEqual(gotKinds, want) {
		t.Fatalf("property accessor kinds = %v, want %v", gotKinds, want)
	}
}

func TestBuildKeepsNonMatchedResolutionStatusesUnconnected(t *testing.T) {
	statuses := []string{"ambiguous", "unresolved", "external", "builtin_like", "member_call", "dynamic", "incomplete"}
	input := Snapshot{
		Symbols: []Symbol{
			{Name: "Caller", Kind: "sub", Module: "Main", File: "Main.bas", Line: 1, Column: 1},
			{Name: "Target", Kind: "sub", Module: "M", File: "M.bas", Line: 2, Column: 1},
		},
		Calls: make([]calls.Call, 0, len(statuses)),
	}
	for i, status := range statuses {
		input.Calls = append(input.Calls, calls.Call{
			CallSite: calls.CallSite{
				File: "Main.bas", Module: "Main",
				Caller: &calls.Caller{Name: "Caller", Kind: "sub", QualifiedName: "Main.Caller"},
				Range:  vbaast.Range{StartLine: i + 2},
			},
			Resolution: calls.Resolution{
				Status:     status,
				Candidates: []calls.Candidate{{QualifiedName: "M.Target", Kind: "sub", File: "M.bas", Line: 2}},
			},
		})
	}

	g := build(input)
	callerKey, ok := g.lookupCallerKey("Main.Caller", "sub", "Main.bas")
	if !ok {
		t.Fatal("caller was not resolved")
	}
	if got := len(g.out[callerKey]); got != 0 {
		t.Fatalf("non-matched statuses created %d edges", got)
	}
	if got := len(g.uncertain[callerKey]); got != len(statuses) {
		t.Fatalf("uncertain calls = %d, want %d", got, len(statuses))
	}
}

func edgeKeys(edges []Edge) []string {
	keys := make([]string, len(edges))
	for i, edge := range edges {
		keys[i] = edgeKey(edge)
	}
	return keys
}

func TestAnalyzeReachabilityHandlesRecursiveDiamondWithoutDuplicates(t *testing.T) {
	input := &calls.Result{
		Symbols: []symbols.Symbol{
			{Name: "Run", Kind: "sub", Visibility: "Public", Module: "Main", File: "Main.bas", StartLine: 1, StartColumn: 1},
			privateSymbol("Main", "Main.bas", "Left", 4),
			privateSymbol("Main", "Main.bas", "Right", 7),
			privateSymbol("Main", "Main.bas", "Join", 10),
		},
		Calls: []calls.Call{
			matched("Main", "Main.bas", "Run", "Main", "Main.bas", "Left", 4, 2),
			matched("Main", "Main.bas", "Run", "Main", "Main.bas", "Right", 7, 3),
			matched("Main", "Main.bas", "Left", "Main", "Main.bas", "Right", 7, 5),
			matched("Main", "Main.bas", "Right", "Main", "Main.bas", "Left", 4, 8),
			matched("Main", "Main.bas", "Left", "Main", "Main.bas", "Join", 10, 6),
			matched("Main", "Main.bas", "Right", "Main", "Main.bas", "Join", 10, 9),
		},
	}
	got := AnalyzeReachability(SnapshotFromResult(input), ReachabilityRequest{Roots: []Root{{Target: "Main.Run", Confidence: RootConfirmed}}})
	if len(got.Confirmed) != 4 || len(got.Possible) != 0 || len(got.Unreachable) != 0 {
		t.Fatalf("recursive diamond reachability = confirmed:%#v possible:%#v unreachable:%#v", got.Confirmed, got.Possible, got.Unreachable)
	}
}

func TestAnalyzeReachabilityTerminatesWhenPossibleRootCallsConfirmedNode(t *testing.T) {
	input := &calls.Result{
		Symbols: []symbols.Symbol{
			{Name: "Run", Kind: "sub", Visibility: "Public", Module: "Main", File: "Main.bas", StartLine: 1, StartColumn: 1},
			{Name: "PublicApi", Kind: "function", Visibility: "Public", Module: "Api", File: "Api.bas", StartLine: 1, StartColumn: 1},
		},
		Calls: []calls.Call{
			matchedKind("Api", "Api.bas", "PublicApi", "function", "Main", "Main.bas", "Run", "sub", 1, 2),
		},
	}

	got := AnalyzeReachability(SnapshotFromResult(input), ReachabilityRequest{Roots: []Root{
		{Target: "Main.Run", Confidence: RootConfirmed},
		{Target: "Api.PublicApi", Confidence: RootPossible},
	}})
	if !hasReachabilityNode(got.Confirmed, "Main.Run") || len(got.Possible) != 1 || !hasReachabilityNode(got.Possible, "Api.PublicApi") {
		t.Fatalf("possible root reachability = confirmed:%#v possible:%#v unreachable:%#v", got.Confirmed, got.Possible, got.Unreachable)
	}
}

func privateSymbol(module, file, name string, line int) symbols.Symbol {
	sym := symbol(module, file, name, line)
	sym.Visibility = "Private"
	return sym
}

func hasReachabilityNode(nodes []Node, qualifiedName string) bool {
	for _, node := range nodes {
		if node.ID.QualifiedName == qualifiedName {
			return true
		}
	}
	return false
}

func hasClusterSize(clusters [][]Node, size int) bool {
	for _, cluster := range clusters {
		if len(cluster) == size {
			return true
		}
	}
	return false
}

func symbol(module, file, name string, line int) symbols.Symbol {
	return symbols.Symbol{Name: name, Kind: "sub", Module: module, File: file, StartLine: line, StartColumn: 1}
}
func matched(callerModule, callerFile, callerName, calleeModule, calleeFile, calleeName string, calleeLine, callLine int) calls.Call {
	return matchedKind(callerModule, callerFile, callerName, "sub", calleeModule, calleeFile, calleeName, "sub", calleeLine, callLine)
}
func matchedKind(callerModule, callerFile, callerName, callerKind, calleeModule, calleeFile, calleeName, calleeKind string, calleeLine, callLine int) calls.Call {
	return calls.Call{CallSite: calls.CallSite{File: callerFile, Module: callerModule, Caller: &calls.Caller{Name: callerName, Kind: callerKind, QualifiedName: callerModule + "." + callerName}, Range: vbaast.Range{StartLine: callLine, StartColumn: 1, EndLine: callLine, EndColumn: 8}}, Resolution: calls.Resolution{Status: "matched", Candidates: []calls.Candidate{{QualifiedName: calleeModule + "." + calleeName, Kind: calleeKind, File: calleeFile, Line: calleeLine}}}}
}
func unresolved(module, file, name string, callLine int) calls.Call {
	return calls.Call{CallSite: calls.CallSite{File: file, Module: module, Caller: &calls.Caller{Name: name, Kind: "sub", QualifiedName: module + "." + name}, Range: vbaast.Range{StartLine: callLine, StartColumn: 1}}, Resolution: calls.Resolution{Status: "unresolved"}}
}
