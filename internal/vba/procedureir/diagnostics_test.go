package procedureir

import "testing"

func TestDiagnosticsFailOpenWhenQualifiedExpressionIsMissing(t *testing.T) {
	procedure := ProcedureIR{
		Accesses: []VariableAccess{{
			Name:         "Ready",
			ExpressionID: 99,
			Resolution: SymbolResolution{
				Status: ResolutionAmbiguous,
				Candidates: []Candidate{
					{QualifiedName: "Main.Ready", Kind: "enum_member"},
					{QualifiedName: "Main.Ready", Kind: "enum_member"},
				},
			},
		}},
	}
	if got := Diagnostics(DocumentIR{Procedures: []ProcedureIR{procedure}}, true); len(got) != 0 {
		t.Fatalf("missing expression diagnostics = %#v, want none", got)
	}
}

func TestDiagnosticsIgnoreIdentifierSubstringInAssignment(t *testing.T) {
	procedure := ProcedureIR{
		Statements: []Statement{{ID: 1, Kind: StatementAssignment, Text: "MyHelperValue = Helper"}},
		Calls: []CallSite{{
			StatementID: 1,
			Callee:      Callee{Text: "Helper", BaseName: "Helper"},
			Resolution:  CallResolution{Status: ResolutionNonCallable, ProjectLocal: true},
		}},
	}
	if got := Diagnostics(DocumentIR{Procedures: []ProcedureIR{procedure}}, true); len(got) != 0 {
		t.Fatalf("identifier-substring diagnostics = %#v, want none", got)
	}
}

func TestSyntacticInvocationRequiresIdentifierBoundary(t *testing.T) {
	procedure := ProcedureIR{Statements: []Statement{{ID: 1, Kind: StatementCall, Text: "MyHelperValue"}}}
	call := CallSite{StatementID: 1, Callee: Callee{Text: "Helper", BaseName: "Helper"}}
	if syntacticInvocation(procedure, call) {
		t.Fatal("substring-only callee should not be treated as an invocation")
	}
}

func TestDiagnosticsFailOpenForIncompleteAndRecoveredIR(t *testing.T) {
	cases := []struct {
		name string
		edit func(*DocumentIR)
	}{
		{name: "parse incomplete", edit: func(doc *DocumentIR) { doc.Parse.HasMissing = true }},
		{name: "procedure recovered", edit: func(doc *DocumentIR) { doc.Procedures[0].Symbol.Recovered = true }},
		{name: "local recovered", edit: func(doc *DocumentIR) { doc.Procedures[0].Declarations[0].Recovered = true }},
		{name: "conditional local", edit: func(doc *DocumentIR) {
			doc.Procedures[0].Declarations[0].ConditionalBranches = []ConditionalBranch{{Group: "1"}}
		}},
		{name: "recovered statement", edit: func(doc *DocumentIR) { doc.Procedures[0].Statements[0].Recovered = true }},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			doc := DocumentIR{
				Procedures: []ProcedureIR{{
					Declarations: []Declaration{{Name: "value"}},
					Statements:   []Statement{{ID: 1, Kind: StatementCall, Text: "Missing()"}},
					Calls: []CallSite{{
						StatementID: 1,
						Callee:      Callee{Text: "Missing()", BaseName: "Missing"},
						Resolution:  CallResolution{Status: ResolutionNonCallable, ProjectLocal: true},
					}},
				}},
			}
			tc.edit(&doc)
			if got := Diagnostics(doc, true); len(got) != 0 {
				t.Fatalf("diagnostics = %#v, want none", got)
			}
		})
	}
}

func TestDiagnosticsFailOpenForRecoveredOrConditionalRaiseEvent(t *testing.T) {
	for _, recovered := range []bool{true, false} {
		doc := DocumentIR{Procedures: []ProcedureIR{{
			RaiseEvents: []RaiseEventReference{{
				Name:       "Missing",
				Recovered:  recovered,
				Resolution: SymbolResolution{Status: ResolutionUnresolved},
				ConditionalBranches: func() []ConditionalBranch {
					if recovered {
						return nil
					}
					return []ConditionalBranch{{Group: "1"}}
				}(),
			}},
		}}}
		if got := Diagnostics(doc, true); len(got) != 0 {
			t.Fatalf("recovered=%v event diagnostics = %#v, want none", recovered, got)
		}
	}
}
