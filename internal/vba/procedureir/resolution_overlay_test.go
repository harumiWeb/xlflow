package procedureir

import (
	"encoding/json"
	"reflect"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

type overlayTestResolver struct{}

func (overlayTestResolver) ResolveCall(site CallSite) CallResolution {
	return CallResolution{Status: ResolutionMatched, Candidates: []Candidate{{QualifiedName: site.Callee.BaseName, Kind: "sub"}}}
}

func (overlayTestResolver) ResolveSymbol(ref SymbolReference) SymbolResolution {
	return SymbolResolution{Status: ResolutionAmbiguous, Scope: ScopeProject, Candidates: []Candidate{
		{QualifiedName: "A." + ref.Name, Kind: "enum_member"},
		{QualifiedName: "B." + ref.Name, Kind: "enum_member"},
	}}
}

func (overlayTestResolver) ResolveEvent(ref SymbolReference) SymbolResolution {
	return SymbolResolution{Status: ResolutionMatched, Scope: ScopeModule, Candidates: []Candidate{{QualifiedName: ref.Module + "." + ref.Name, Kind: "event"}}}
}

type overlayDiagnosticResolver struct{}

func (overlayDiagnosticResolver) ResolveCall(site CallSite) CallResolution {
	return CallResolution{
		Status:       ResolutionNonCallable,
		ProjectLocal: true,
		Candidates:   []Candidate{{QualifiedName: site.Callee.BaseName, Kind: "variable"}},
	}
}

func (overlayDiagnosticResolver) ResolveSymbol(SymbolReference) SymbolResolution {
	return SymbolResolution{Status: ResolutionAmbiguous, Scope: ScopeProject, Candidates: []Candidate{
		{QualifiedName: "A.Member", Kind: "enum_member"},
		{QualifiedName: "B.Member", Kind: "enum_member"},
	}}
}

func (overlayDiagnosticResolver) ResolveEvent(SymbolReference) SymbolResolution {
	return SymbolResolution{Scope: ScopeUnresolved, Status: ResolutionUnresolved}
}

func overlayTestDocument() DocumentIR {
	return DocumentIR{Path: "Main.bas", ModuleName: "Main", ModuleKind: "class", Procedures: []ProcedureIR{{
		Symbol:      ProcedureSymbol{Name: "Run", QualifiedName: "Main.Run", Kind: ProcedureSub, DeclarationRange: vbaast.Range{StartByte: 10}},
		Statements:  []Statement{{ID: 1, Kind: StatementCall, Text: "Helper()"}},
		Expressions: []Expression{{ID: 1, Kind: ExpressionIdentifier, Text: "value"}},
		Calls:       []CallSite{{ID: 0, Callee: Callee{Text: "Helper", BaseName: "Helper"}, Range: vbaast.Range{StartByte: 20, EndByte: 28}, StatementID: 1}},
		Accesses:    []VariableAccess{{ID: 0, Name: "value", Scope: ScopeProject, ExpressionID: 1}},
		RaiseEvents: []RaiseEventReference{{ID: 0, Name: "Changed", Module: "Main", Caller: ProcedureRef{Name: "Run"}}},
	}}}
}

func TestResolveViewUsesFactIDsWithoutMutatingInput(t *testing.T) {
	doc := overlayTestDocument()
	before := overlayTestDocument()
	view := ResolveView(doc, overlayTestResolver{})

	call, ok := view.ResolvedCall(0, 1)
	if !ok || call.Resolution.Status != ResolutionMatched {
		t.Fatalf("ResolvedCall = (%+v, %t)", call, ok)
	}
	access, ok := view.ResolvedAccess(0, 0)
	if !ok || access.Scope != ScopeProject || access.Resolution.Status != ResolutionAmbiguous {
		t.Fatalf("ResolvedAccess = (%+v, %t)", access, ok)
	}
	event, ok := view.ResolvedEvent(0, 0)
	if !ok || event.Resolution.Status != ResolutionMatched {
		t.Fatalf("ResolvedEvent = (%+v, %t)", event, ok)
	}
	if !reflect.DeepEqual(doc, before) {
		t.Fatalf("ResolveView mutated input: before=%+v after=%+v", before, doc)
	}

	call.Resolution.Candidates[0].QualifiedName = "mutated"
	again, ok := view.ResolvedCall(0, 1)
	if !ok || again.Resolution.Candidates[0].QualifiedName == "mutated" {
		t.Fatal("ResolvedCall exposed mutable overlay candidates")
	}
}

func TestResolvedProcedureProjectsFactsWithoutCloningSyntaxPayload(t *testing.T) {
	doc := overlayTestDocument()
	view := ResolveView(doc, overlayTestResolver{})
	sourceStatus := doc.Procedures[0].Calls[0].Resolution.Status
	projected, ok := view.ResolvedProcedure(0)
	if !ok {
		t.Fatal("ResolvedProcedure returned no procedure")
	}
	if projected.Symbol.Name != doc.Procedures[0].Symbol.Name {
		t.Fatalf("projected symbol = %q, want %q", projected.Symbol.Name, doc.Procedures[0].Symbol.Name)
	}
	if projected.Calls[0].Resolution.Status != ResolutionMatched || doc.Procedures[0].Calls[0].Resolution.Status != sourceStatus {
		t.Fatalf("projected/source resolutions = %+v / %+v", projected.Calls[0].Resolution, doc.Procedures[0].Calls[0].Resolution)
	}
	if len(projected.Statements) == 0 || &projected.Statements[0] != &doc.Procedures[0].Statements[0] {
		t.Fatal("syntax-local statement payload was unexpectedly copied")
	}
	projected.Calls[0].Resolution.Candidates[0].QualifiedName = "mutated"
	if got, _ := view.ResolvedCall(0, 1); got.Resolution.Candidates[0].QualifiedName == "mutated" {
		t.Fatal("projected call exposed mutable overlay storage")
	}
}

func TestResolveViewUsesOrdinalFallbackForDuplicateHandBuiltIDs(t *testing.T) {
	doc := overlayTestDocument()
	doc.Procedures[0].Calls = append(doc.Procedures[0].Calls,
		CallSite{ID: 1, Callee: Callee{Text: "Other", BaseName: "Other"}},
	)
	doc.Procedures[0].Accesses = append(doc.Procedures[0].Accesses,
		VariableAccess{ID: 0, Name: "other", Scope: ScopeProject},
	)
	view := ResolveView(doc, overlayTestResolver{})
	if _, ok := view.ResolvedCall(0, 1); !ok {
		t.Fatal("first duplicate call ID did not resolve")
	}
	if _, ok := view.ResolvedCall(0, 2); !ok {
		t.Fatal("ordinal fallback did not resolve second duplicate call ID")
	}
	if _, ok := view.ResolvedAccess(0, 1); !ok {
		t.Fatal("first zero access ID did not resolve")
	}
	if _, ok := view.ResolvedAccess(0, 2); !ok {
		t.Fatal("ordinal fallback did not resolve second zero access ID")
	}
}

func TestResolveMaterializeMatchesOverlayResolution(t *testing.T) {
	doc := overlayTestDocument()
	view := ResolveView(doc, overlayTestResolver{})
	materialized := view.Materialize()
	legacy := Resolve(doc, overlayTestResolver{})
	if !reflect.DeepEqual(materialized, legacy) {
		t.Fatalf("Materialize differs from Resolve:\nmaterialized=%+v\nlegacy=%+v", materialized, legacy)
	}
	if !reflect.DeepEqual(doc, overlayTestDocument()) {
		t.Fatal("materialization mutated source document")
	}
}

func TestResolveViewNilResolverPreservesCloneOnlyCompatibility(t *testing.T) {
	doc := overlayTestDocument()
	view := ResolveView(doc, nil)
	if _, ok := view.ResolvedCall(0, 1); ok {
		t.Fatal("nil-resolver view unexpectedly exposed a call resolution")
	}
	doc.Procedures[0].Calls[0].Resolution = CallResolution{Status: ResolutionNotAttempted}
	doc.Procedures[0].Accesses[0].Resolution = SymbolResolution{Scope: ScopeProject, Status: ResolutionNotAttempted}
	doc.Procedures[0].RaiseEvents[0].Resolution = SymbolResolution{Scope: ScopeUnresolved, Status: ResolutionNotAttempted}

	got := view.Materialize()
	want := Clone(doc)
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("nil-resolver materialization differs from Clone:\nmaterialized=%+v\nclone=%+v", got, want)
	}
	if !reflect.DeepEqual(Resolve(doc, nil), want) {
		t.Fatal("Resolve(nil) no longer has clone-only compatibility semantics")
	}
}

func TestResolveNilResolverPreservesInputFacts(t *testing.T) {
	doc := overlayTestDocument()
	doc.Procedures[0].Calls[0].NonCallableNames = []string{"existing"}
	if got, want := Resolve(doc, nil), Clone(doc); !reflect.DeepEqual(got, want) {
		t.Fatalf("Resolve(nil) changed facts: got=%+v want=%+v", got, want)
	}
}

func TestDiagnosticsViewMatchesMaterializedDiagnostics(t *testing.T) {
	doc := overlayTestDocument()
	view := ResolveView(doc, overlayTestResolver{})
	got := DiagnosticsView(view, true)
	want := Diagnostics(view.Materialize(), true)
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("DiagnosticsView = %+v, Diagnostics(Materialize) = %+v", got, want)
	}
}

func TestDiagnosticsViewPreservesVB052VB053VB054Parity(t *testing.T) {
	doc := DocumentIR{
		ModuleName: "Main",
		Procedures: []ProcedureIR{{
			Symbol:      ProcedureSymbol{Name: "Run", QualifiedName: "Main.Run", Kind: ProcedureSub},
			Statements:  []Statement{{ID: 1, Kind: StatementCall, Text: "Missing()"}},
			Expressions: []Expression{{ID: 1, Kind: ExpressionIdentifier, Text: "Member"}},
			Calls:       []CallSite{{ID: 1, Callee: Callee{Text: "Missing()", BaseName: "Missing"}, StatementID: 1}},
			RaiseEvents: []RaiseEventReference{{ID: 1, Name: "MissingEvent"}},
			Accesses:    []VariableAccess{{ID: 1, Name: "Member", Scope: ScopeProject, ExpressionID: 1}},
		}},
	}
	view := ResolveView(doc, overlayDiagnosticResolver{})
	got := DiagnosticsView(view, true)
	want := Diagnostics(view.Materialize(), true)
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("DiagnosticsView differs from materialized diagnostics:\nview=%+v\nmaterialized=%+v", got, want)
	}
	if len(got) != 3 || got[0].Code != "VB052" || got[1].Code != "VB054" || got[2].Code != "VB053" {
		t.Fatalf("diagnostics = %+v, want VB052, VB054, VB053 in source walker order", got)
	}
	if !reflect.DeepEqual(got[0].Candidates, []Candidate{{QualifiedName: "Missing", Kind: "variable"}}) {
		t.Fatalf("VB052 candidates = %+v", got[0].Candidates)
	}
}

func TestResolveViewAssignsProductionAccessAndEventIDs(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "IDs.bas", ModuleName: "IDs", ModuleKind: "class"}, []byte(`Public Sub Run()
  Dim value As Long
  value = value + 1
  RaiseEvent Changed
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 1 {
		t.Fatalf("procedures = %d, want 1", len(doc.Procedures))
	}
	procedure := doc.Procedures[0]
	for index, access := range procedure.Accesses {
		if access.ID != index+1 {
			t.Fatalf("access %d ID = %d, want %d", index, access.ID, index+1)
		}
	}
	for index, event := range procedure.RaiseEvents {
		if event.ID != index+1 {
			t.Fatalf("event %d ID = %d, want %d", index, event.ID, index+1)
		}
	}
}

func TestResolutionFactIDsSurviveCloneAndRebase(t *testing.T) {
	doc := overlayTestDocument()
	procedure := doc.Procedures[0]
	cloned := CloneProcedureIR(procedure)
	rebased := RebaseProcedure(procedure,
		vbaast.Range{StartLine: 1, StartByte: 10},
		vbaast.Range{StartLine: 5, StartByte: 100},
	)
	if cloned.Accesses[0].ID != procedure.Accesses[0].ID || cloned.RaiseEvents[0].ID != procedure.RaiseEvents[0].ID {
		t.Fatalf("clone changed resolution fact IDs: clone=%+v source=%+v", cloned, procedure)
	}
	if rebased.Accesses[0].ID != procedure.Accesses[0].ID || rebased.RaiseEvents[0].ID != procedure.RaiseEvents[0].ID {
		t.Fatalf("rebase changed resolution fact IDs: rebased=%+v source=%+v", rebased, procedure)
	}
}

func TestResolutionFactIDsAreNotSerialized(t *testing.T) {
	payload, err := json.Marshal(overlayTestDocument())
	if err != nil {
		t.Fatal(err)
	}
	encoded := string(payload)
	for _, field := range []string{"\"accesses\":[{\"id\"", "\"raiseEvents\":[{\"id\""} {
		if containsJSONField(encoded, field) {
			t.Fatalf("resolution fact ID leaked into JSON: %s", encoded)
		}
	}
}

func containsJSONField(encoded, field string) bool {
	for i := 0; i+len(field) <= len(encoded); i++ {
		if encoded[i:i+len(field)] == field {
			return true
		}
	}
	return false
}
