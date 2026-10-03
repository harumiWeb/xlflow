package callgraph

import "testing"

func TestAnalyzeReachabilityWithRootsRetainsResolvedConfidence(t *testing.T) {
	requested := []Root{
		{Target: "A.Only", Confidence: RootConfirmed, Reason: "configured"},
		{Target: "A.Duplicate", Confidence: RootConfirmed, Reason: "configured"},
		{Target: "Legacy.Probe", Confidence: RootConfirmed, Reason: "configured"},
		{Target: "Missing.Entry", Confidence: RootConfirmed, Reason: "configured"},
	}
	result, roots := AnalyzeReachabilityWithRoots(Snapshot{Symbols: []Symbol{
		{Name: "Only", Kind: "sub", Module: "A", ModuleKind: "standard", File: "A.bas", Line: 1, Visibility: "Public"},
		{Name: "Duplicate", Kind: "sub", Module: "A", ModuleKind: "standard", File: "A.bas", Line: 4, Visibility: "Public"},
		{Name: "Duplicate", Kind: "sub", Module: "A", ModuleKind: "standard", File: "A_Copy.bas", Line: 4, Visibility: "Public"},
		{Name: "Probe", Kind: "sub", Module: "B", ModuleKind: "standard", File: "B.bas", Line: 2, Visibility: "Public"},
	}}, ReachabilityRequest{Roots: requested})
	if len(roots) != len(requested) {
		t.Fatalf("root resolutions=%d, want %d: %+v", len(roots), len(requested), roots)
	}

	if roots[0].Status != "resolved" || roots[0].Confidence != RootConfirmed || len(roots[0].Nodes) != 1 {
		t.Errorf("exact root resolution=%+v", roots[0])
	}
	if roots[1].Status != "ambiguous" || roots[1].Confidence != RootPossible || len(roots[1].Nodes) != 2 {
		t.Errorf("ambiguous root resolution=%+v", roots[1])
	}
	if roots[2].Status != "fallback" || roots[2].Confidence != RootPossible || len(roots[2].Nodes) != 1 {
		t.Errorf("fallback root resolution=%+v", roots[2])
	}
	if roots[3].Status != "unresolved" || roots[3].Confidence != RootPossible || len(roots[3].Nodes) != 0 {
		t.Errorf("unresolved root resolution=%+v", roots[3])
	}
	for index, resolution := range roots {
		if resolution.Root != requested[index] {
			t.Errorf("root request %d changed in resolution: got %+v want %+v", index, resolution.Root, requested[index])
		}
	}
	if !hasReachabilityNode(result.Confirmed, "A.Only") {
		t.Errorf("confirmed exact root was not confirmed: %+v", result.Confirmed)
	}
	for _, target := range []string{"A.Duplicate", "B.Probe"} {
		if !hasReachabilityNode(result.Possible, target) {
			t.Errorf("downgraded root target %q is not possible: %+v", target, result.Possible)
		}
	}
}
