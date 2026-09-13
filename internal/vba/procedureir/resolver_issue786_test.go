package procedureir

import "testing"

func TestIssue786ReceiverlessCallPrefersSameModuleProcedure(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name       string
		moduleKind string
		kind       string
	}{
		{name: "standard sub", moduleKind: "standard", kind: "sub"},
		{name: "class private sub", moduleKind: "class", kind: "sub"},
		{name: "class private function", moduleKind: "class", kind: "function"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			resolver := NewResolver([]ResolverSymbol{
				{Name: "printMsg", Module: "PrintMsgOverrideRepro", ModuleKind: test.moduleKind, Kind: test.kind, Visibility: "Private", File: "PrintMsgOverrideRepro.cls"},
				{Name: "printMsg", Module: "CDPHelpers", ModuleKind: "standard", Kind: test.kind, Visibility: "Public", File: "CDPHelpers.bas"},
			})
			got := resolver.ResolveCall(CallSite{
				Module: "PrintMsgOverrideRepro",
				Caller: ProcedureRef{QualifiedName: "PrintMsgOverrideRepro.CallIt"},
				Callee: Callee{Text: "printMsg", BaseName: "printMsg", Member: "printMsg"},
			})
			if got.Status != ResolutionMatched || len(got.Candidates) != 1 || got.Candidates[0].QualifiedName != "PrintMsgOverrideRepro.printMsg" {
				t.Fatalf("same-module receiverless resolution = %#v, want local match", got)
			}
		})
	}
}

func TestIssue786ReceiverlessCallKeepsExternalStandardFallback(t *testing.T) {
	t.Parallel()
	resolver := NewResolver([]ResolverSymbol{{
		Name: "printMsg", Module: "CDPHelpers", ModuleKind: "standard", Kind: "sub", Visibility: "Public", File: "CDPHelpers.bas",
	}})
	got := resolver.ResolveCall(CallSite{
		Module: "PrintMsgOverrideRepro",
		Caller: ProcedureRef{QualifiedName: "PrintMsgOverrideRepro.CallIt"},
		Callee: Callee{Text: "printMsg", BaseName: "printMsg", Member: "printMsg"},
	})
	if got.Status != ResolutionMatched || len(got.Candidates) != 1 || got.Candidates[0].QualifiedName != "CDPHelpers.printMsg" {
		t.Fatalf("external standard fallback = %#v, want external match", got)
	}
}

func TestIssue786QualifiedCallStillSelectsExternalStandardProcedure(t *testing.T) {
	t.Parallel()
	resolver := NewResolver([]ResolverSymbol{
		{Name: "printMsg", Module: "PrintMsgOverrideRepro", ModuleKind: "class", Kind: "sub", Visibility: "Private", File: "PrintMsgOverrideRepro.cls"},
		{Name: "printMsg", Module: "CDPHelpers", ModuleKind: "standard", Kind: "sub", Visibility: "Public", File: "CDPHelpers.bas"},
	})
	receiver := "CDPHelpers"
	got := resolver.ResolveCall(CallSite{
		Module: "PrintMsgOverrideRepro",
		Caller: ProcedureRef{QualifiedName: "PrintMsgOverrideRepro.CallIt"},
		Callee: Callee{Text: "CDPHelpers.printMsg", BaseName: "printMsg", Member: "printMsg", Receiver: &receiver},
	})
	if got.Status != ResolutionMatched || len(got.Candidates) != 1 || got.Candidates[0].QualifiedName != "CDPHelpers.printMsg" {
		t.Fatalf("qualified external resolution = %#v, want external match", got)
	}
}

func TestIssue786MultipleSameModuleProceduresRemainAmbiguous(t *testing.T) {
	t.Parallel()
	resolver := NewResolver([]ResolverSymbol{
		{Name: "printMsg", Module: "PrintMsgOverrideRepro", ModuleKind: "class", Kind: "sub", Visibility: "Private", File: "PrintMsgOverrideRepro.cls", Line: 2},
		{Name: "printMsg", Module: "PrintMsgOverrideRepro", ModuleKind: "class", Kind: "sub", Visibility: "Private", File: "PrintMsgOverrideRepro.cls", Line: 3},
		{Name: "printMsg", Module: "CDPHelpers", ModuleKind: "standard", Kind: "sub", Visibility: "Public", File: "CDPHelpers.bas", Line: 2},
	})
	got := resolver.ResolveCall(CallSite{
		Module: "PrintMsgOverrideRepro",
		Caller: ProcedureRef{QualifiedName: "PrintMsgOverrideRepro.CallIt"},
		Callee: Callee{Text: "printMsg", BaseName: "printMsg", Member: "printMsg"},
	})
	if got.Status != ResolutionAmbiguous || len(got.Candidates) != 2 {
		t.Fatalf("same-module duplicate resolution = %#v, want two local candidates", got)
	}
	for _, candidate := range got.Candidates {
		if candidate.QualifiedName != "PrintMsgOverrideRepro.printMsg" {
			t.Fatalf("duplicate resolution retained external candidate: %#v", got.Candidates)
		}
	}
}

func TestIssue786ConditionalLocalProcedureDoesNotHideExternalFallback(t *testing.T) {
	t.Parallel()
	symbols := []ResolverSymbol{
		{
			Name: "printMsg", Module: "PrintMsgOverrideRepro", ModuleKind: "class", Kind: "sub", Visibility: "Private",
			File: "PrintMsgOverrideRepro.cls", ConditionalBranches: []ConditionalBranch{{Group: "issue786", Branch: 0}},
		},
		{Name: "printMsg", Module: "CDPHelpers", ModuleKind: "standard", Kind: "sub", Visibility: "Public", File: "CDPHelpers.bas"},
	}
	site := CallSite{
		Module: "PrintMsgOverrideRepro",
		Caller: ProcedureRef{QualifiedName: "PrintMsgOverrideRepro.CallIt"},
		Callee: Callee{Text: "printMsg", BaseName: "printMsg", Member: "printMsg"},
	}
	for _, test := range []struct {
		name     string
		complete bool
		status   ResolutionStatus
	}{
		{name: "incomplete project", complete: false, status: ResolutionIncomplete},
		{name: "complete project", complete: true, status: ResolutionAmbiguous},
	} {
		t.Run(test.name, func(t *testing.T) {
			got := NewResolverWithCompleteness(symbols, test.complete).ResolveCall(site)
			if got.Status != test.status || len(got.Candidates) != 2 {
				t.Fatalf("conditional local resolution = %#v, want %s with both candidates", got, test.status)
			}
		})
	}
}

func TestIssue786SoleConditionalLocalProcedureIsIncomplete(t *testing.T) {
	t.Parallel()
	resolver := NewResolver([]ResolverSymbol{
		{
			Name: "printMsg", Module: "PrintMsgOverrideRepro", ModuleKind: "class", Kind: "sub", Visibility: "Private",
			File: "PrintMsgOverrideRepro.cls", ConditionalBranches: []ConditionalBranch{{Group: "issue786", Branch: 0}},
		},
	})
	got := resolver.ResolveCall(CallSite{
		Module: "PrintMsgOverrideRepro",
		Caller: ProcedureRef{QualifiedName: "PrintMsgOverrideRepro.CallIt"},
		Callee: Callee{Text: "printMsg", BaseName: "printMsg", Member: "printMsg"},
	})
	if got.Status != ResolutionIncomplete || len(got.Candidates) != 1 || got.Candidates[0].QualifiedName != "PrintMsgOverrideRepro.printMsg" {
		t.Fatalf("sole conditional local resolution = %#v, want incomplete local candidate", got)
	}
}
