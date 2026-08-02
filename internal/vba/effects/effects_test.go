package effects

import (
	"reflect"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestDirectEffectsUseReachableHighConfidenceIR(t *testing.T) {
	summary := buildSources(t, sourceFile{"Module1.bas", "Module1", `Public Sub Run()
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Range("A1").Value = 1
    MsgBox "done"
    Shell "calc.exe"
    Workbooks.Open "book.xlsx"
    ThisWorkbook.Close
    On Error Resume Next
    Err.Raise 5
    GoTo Finished
    Range("Z9").Value = 2
Finished:
End Sub
`})
	run := find(t, summary, "Module1.Run")
	want := map[EffectKind]bool{
		DisablesEvents: true, ChangesCalculation: true, WritesCells: true,
		ShowsDialog: true, LaunchesProcess: true,
		OpensWorkbook: true, ClosesWorkbook: true, SuppressesErrors: true,
		RaisesError: true,
	}
	for kind := range want {
		if !run.Has(kind) {
			t.Errorf("missing direct effect %q: %#v", kind, run.Direct)
		}
	}
	for _, evidence := range run.Direct {
		if evidence.Target == `Range("Z9").Value` {
			t.Fatal("unreachable assignment was classified")
		}
	}
}

func TestRestoresEventsAndCellMutations(t *testing.T) {
	summary := buildSources(t, sourceFile{"Helpers.bas", "Helpers", `Private savedEvents As Boolean
Public Sub Restore()
    Application.EnableEvents = savedEvents
    Worksheets(1).Range("A1").ClearContents
    ThisWorkbook.Save
End Sub
`})
	restore := find(t, summary, "Helpers.Restore")
	for _, kind := range []EffectKind{RestoresEvents, RestoresApplicationState, WritesCells, ChangesWorkbook} {
		if !restore.Has(kind) {
			t.Errorf("missing %s", kind)
		}
	}
}

func TestEffectsRecognizeEventTriggeringOperations(t *testing.T) {
	summary := buildSources(t, sourceFile{"Events.bas", "Events", `Public Sub Trigger()
    Application.Calculate
    Application.Goto Range("A1")
    Worksheets.Add
    Worksheets(1).Name = "Renamed"
End Sub
`})
	trigger := find(t, summary, "Events.Trigger")
	for _, kind := range []EffectKind{Recalculates, ChangesSelection, ChangesWorkbook} {
		if !trigger.Has(kind) {
			t.Errorf("missing event-triggering effect %s: %#v", kind, trigger.Direct)
		}
	}
}

func TestCellAndSelectionEffectsAvoidUnrelatedWorkbookOrObjectEffects(t *testing.T) {
	summary := buildSources(t, sourceFile{"Effects.bas", "Effects", `Public Sub Trigger()
    Dim worker As Object
    Range("A1").Value = 1
    worker.Activate
    Application.Goto Range("B2")
End Sub
`})
	trigger := find(t, summary, "Effects.Trigger")
	if got := count(trigger.Direct, WritesCells); got != 1 {
		t.Fatalf("cell write effects = %d, want 1: %#v", got, trigger.Direct)
	}
	if got := count(trigger.Direct, ChangesWorkbook); got != 0 {
		t.Fatalf("cell value assignment must not duplicate changes_workbook: %#v", trigger.Direct)
	}
	if got := count(trigger.Direct, ChangesSelection); got != 1 {
		t.Fatalf("only Application.Goto should change selection: %#v", trigger.Direct)
	}
}

func TestErrorFunctionDoesNotRaiseButErrorStatementDoes(t *testing.T) {
	summary := buildSources(t, sourceFile{"Errors.bas", "Errors", `Public Sub Run()
    Dim description As String
    description = Error(5)
End Sub

Public Sub RaiseError()
    Error 6
End Sub
`})
	run := find(t, summary, "Errors.Run")
	if got := count(run.Direct, RaisesError); got != 0 {
		t.Fatalf("Error() function raises_error count = %d, want 0: %#v", got, run.Direct)
	}
	raise := find(t, summary, "Errors.RaiseError")
	if got := count(raise.Direct, RaisesError); got != 1 {
		t.Fatalf("Error statement raises_error count = %d, want 1: %#v", got, raise.Direct)
	}
}

func TestGenericApplicationStateEvidencePreservesVBA203Properties(t *testing.T) {
	summary := buildSources(t, sourceFile{"State.bas", "State", `Private savedAlerts As Boolean
Public Sub PushState()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub
Public Sub PopState()
    Application.DisplayAlerts = savedAlerts
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
`})
	push := find(t, summary, "State.PushState")
	pop := find(t, summary, "State.PopState")
	if count(push.Direct, ChangesApplicationState) != 2 {
		t.Fatalf("changes = %#v", push.Direct)
	}
	if count(push.Direct, RestoresApplicationState) != 0 {
		t.Fatalf("state-setting assignments produced restore evidence: %#v", push.Direct)
	}
	if count(pop.Direct, RestoresApplicationState) != 3 {
		t.Fatalf("restores = %#v", pop.Direct)
	}
}

func TestGenericApplicationStateEvidenceCoversAllTrackedProperties(t *testing.T) {
	summary := buildSources(t, sourceFile{"State.bas", "State", `Private savedStatus As Variant
Private savedInteractive As Boolean
Private savedLinks As Boolean
Public Sub PushState()
    Application.StatusBar = "working"
    Application.Cursor = xlWait
    Application.Interactive = False
    Application.AskToUpdateLinks = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.CutCopyMode = xlCopy
End Sub
Public Sub PopState()
    Application.StatusBar = savedStatus
    Application.Cursor = xlDefault
    Application.Interactive = savedInteractive
    Application.AskToUpdateLinks = savedLinks
    Application.AutomationSecurity = msoAutomationSecurityByUI
    Application.CutCopyMode = False
End Sub
`})
	push := find(t, summary, "State.PushState")
	pop := find(t, summary, "State.PopState")
	if count(push.Direct, ChangesApplicationState) != 6 {
		t.Fatalf("push state changes = %#v", push.Direct)
	}
	if count(push.Direct, RestoresApplicationState) != 0 {
		t.Fatalf("state-setting assignments produced restore evidence: %#v", push.Direct)
	}
	if count(pop.Direct, ChangesApplicationState) != 6 || count(pop.Direct, RestoresApplicationState) != 6 {
		t.Fatalf("pop state evidence = %#v", pop.Direct)
	}
}

func TestGenericApplicationStateEvidenceRecognizesWithApplication(t *testing.T) {
	summary := buildSources(t, sourceFile{"State.bas", "State", `Private savedEvents As Boolean
Public Sub PushState()
    With Application
        .EnableEvents = False
    End With
End Sub
Public Sub PopState()
    With Application
        .EnableEvents = savedEvents
    End With
End Sub
`})
	push := find(t, summary, "State.PushState")
	pop := find(t, summary, "State.PopState")
	if count(push.Direct, ChangesApplicationState) != 1 || count(push.Direct, RestoresApplicationState) != 0 {
		t.Fatalf("push state evidence = %#v", push.Direct)
	}
	if count(pop.Direct, RestoresApplicationState) != 1 {
		t.Fatalf("pop state evidence = %#v", pop.Direct)
	}
}

func TestPropagationConvergesAndDeduplicatesDiamondAndCycles(t *testing.T) {
	summary := buildSources(t,
		sourceFile{"A.bas", "A", "Public Sub Root()\n Left\n Right\nEnd Sub\n"},
		sourceFile{"B.bas", "B", "Public Sub Left()\n Leaf\nEnd Sub\n"},
		sourceFile{"C.bas", "C", "Public Sub Right()\n Leaf\nEnd Sub\n"},
		sourceFile{"D.bas", "D", "Public Sub Leaf()\n Range(\"A1\").Value = 1\n Root\nEnd Sub\n"},
	)
	root := find(t, summary, "A.Root")
	if len(root.Direct) != 0 {
		t.Fatalf("root direct = %#v", root.Direct)
	}
	if count(root.Propagated, WritesCells) != 1 || count(root.Propagated, ChangesWorkbook) != 0 {
		t.Fatalf("diamond/cycle provenance not deduplicated: %#v", root.Propagated)
	}
	leaf := find(t, summary, "D.Leaf")
	if len(leaf.Propagated) != 0 {
		t.Fatalf("self-origin returned through cycle: %#v", leaf.Propagated)
	}
}

func TestCallUncertaintyClassificationAndPropagation(t *testing.T) {
	doc := procedureir.DocumentIR{Path: "Calls.bas", ModuleName: "Calls", Procedures: []procedureir.ProcedureIR{
		manualProcedure("Calls.Root", 1, []procedureir.CallSite{{ID: 1, StatementID: 1, Callee: procedureir.Callee{Text: "Child", BaseName: "Child"}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "Calls.Child", Kind: "sub", File: "Calls.bas", Line: 10}}}}}),
		manualProcedure("Calls.Child", 10, []procedureir.CallSite{
			manualCall(1, procedureir.ResolutionAmbiguous), manualCall(2, procedureir.ResolutionUnresolved),
			manualCall(3, procedureir.ResolutionExternal), manualCall(4, procedureir.ResolutionMemberCall),
			manualCall(5, procedureir.ResolutionBuiltinLike),
		}),
	}}
	project := Build([]Document{{IR: doc, CFG: cfg.BuildDocument(doc)}})
	child := find(t, project, "Calls.Child")
	if got := uncertaintyKinds(child.DirectUncertainty); !reflect.DeepEqual(got, []UncertaintyKind{UncertaintyAmbiguous, UncertaintyDynamic, UncertaintyExternal, UncertaintyUnresolved}) {
		t.Fatalf("direct uncertainty = %v", got)
	}
	root := find(t, project, "Calls.Root")
	if len(root.PropagatedUncertainty) != 4 {
		t.Fatalf("propagated uncertainty = %#v", root.PropagatedUncertainty)
	}
}

func TestBuildIsDeterministicAcrossDocumentOrder(t *testing.T) {
	a := sourceFile{"A.bas", "A", "Public Sub AProc()\n BProc\nEnd Sub\n"}
	b := sourceFile{"B.bas", "B", "Public Sub BProc()\n Application.EnableEvents = True\nEnd Sub\n"}
	first := buildSources(t, a, b)
	second := buildSources(t, b, a)
	if !reflect.DeepEqual(first.All(), second.All()) {
		t.Fatalf("order-dependent summaries:\n%#v\n%#v", first.All(), second.All())
	}
}

func TestProjectSummaryReturnsDefensiveCopies(t *testing.T) {
	project := buildSources(t, sourceFile{"State.bas", "State", "Public Sub Run()\n Application.EnableEvents = False\nEnd Sub\n"})
	all := project.All()
	id := all[0].Identity
	all[0].Identity.Name = "mutated"
	all[0].Direct[0].Target = "mutated"

	first, ok := project.Lookup(id)
	if !ok {
		t.Fatal("summary lookup failed")
	}
	first.Direct[0].Target = "mutated again"
	second, ok := project.Lookup(id)
	if !ok || second.Identity.Name != "Run" || second.Direct[0].Target == "mutated" || second.Direct[0].Target == "mutated again" {
		t.Fatalf("project summary was mutated through a returned value: %#v", second)
	}
}

func TestUnknownAliasesAndMatchedProjectNamesAreNotPositiveEffects(t *testing.T) {
	doc := procedureir.DocumentIR{Path: "FalsePositive.bas", ModuleName: "FalsePositive", Procedures: []procedureir.ProcedureIR{
		{
			Symbol: procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "FalsePositive.Run", Kind: procedureir.ProcedureSub, DeclarationRange: rangeAt(1)},
			Statements: []procedureir.Statement{
				{ID: 1, Kind: procedureir.StatementAssignment, Target: &procedureir.Expression{Text: `foo.Range("A1").Value`}, Value: &procedureir.Expression{Text: "1"}},
				{ID: 2, Kind: procedureir.StatementCall},
			},
			Calls: []procedureir.CallSite{
				{ID: 1, StatementID: 2, Callee: procedureir.Callee{Text: "MsgBox", BaseName: "MsgBox"}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "Other.MsgBox", Kind: "sub", File: "Other.bas", Line: 1}}}},
				{ID: 2, StatementID: 2, Callee: procedureir.Callee{Text: "InputBox", BaseName: "InputBox"}, Resolution: procedureir.CallResolution{Status: procedureir.ResolutionAmbiguous, Candidates: []procedureir.Candidate{{QualifiedName: "One.InputBox"}, {QualifiedName: "Two.InputBox"}}}},
			},
		},
		{Symbol: procedureir.ProcedureSymbol{Name: "MsgBox", QualifiedName: "Other.MsgBox", Kind: procedureir.ProcedureSub, DeclarationRange: rangeAt(1)}},
	}}
	project := Build([]Document{{IR: doc, CFG: cfg.BuildDocument(doc)}})
	run := find(t, project, "FalsePositive.Run")
	if run.Has(WritesCells) || run.Has(ShowsDialog) {
		t.Fatalf("unknown alias/project call classified: %#v", run)
	}
}

func TestRecoveredStatementsDoNotProduceEffects(t *testing.T) {
	proc := procedureir.ProcedureIR{
		Symbol:     procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "M.Run", Kind: procedureir.ProcedureSub, DeclarationRange: rangeAt(1)},
		Statements: []procedureir.Statement{{ID: 1, Kind: procedureir.StatementAssignment, Recovered: true, Target: &procedureir.Expression{Text: `Range("A1").Value`}, Value: &procedureir.Expression{Text: "1"}}},
	}
	doc := procedureir.DocumentIR{Path: "M.bas", ModuleName: "M", Procedures: []procedureir.ProcedureIR{proc}}
	run := find(t, Build([]Document{{IR: doc, CFG: cfg.BuildDocument(doc)}}), "M.Run")
	if len(run.Direct) != 0 {
		t.Fatalf("recovered effect = %#v", run.Direct)
	}
}

func TestMissingCFGDoesNotClaimStatementsAreUnreachable(t *testing.T) {
	proc := procedureir.ProcedureIR{
		Symbol:     procedureir.ProcedureSymbol{Name: "Run", QualifiedName: "M.Run", Kind: procedureir.ProcedureSub, DeclarationRange: rangeAt(1)},
		Statements: []procedureir.Statement{{ID: 1, Kind: procedureir.StatementAssignment, Target: &procedureir.Expression{Text: `Range("A1").Value`}, Value: &procedureir.Expression{Text: "1"}}},
	}
	doc := procedureir.DocumentIR{Path: "M.bas", ModuleName: "M", Procedures: []procedureir.ProcedureIR{proc}}
	run := find(t, Build([]Document{{IR: doc}}), "M.Run")
	if !run.Has(WritesCells) {
		t.Fatal("missing CFG was treated as proof of unreachability")
	}
}

type sourceFile struct{ path, module, source string }

func buildSources(t *testing.T, sources ...sourceFile) ProjectSummary {
	t.Helper()
	irs := make([]procedureir.DocumentIR, len(sources))
	var symbols []procedureir.ResolverSymbol
	for i, source := range sources {
		ir, err := procedureir.BuildSource(procedureir.BuildOptions{Path: source.path, ModuleName: source.module}, []byte(source.source))
		if err != nil {
			t.Fatal(err)
		}
		irs[i] = ir
		for _, proc := range ir.Procedures {
			symbols = append(symbols, procedureir.ResolverSymbol{Name: proc.Symbol.Name, Module: ir.ModuleName, ModuleKind: ir.ModuleKind, Kind: string(proc.Symbol.Kind), Visibility: proc.Symbol.Visibility, File: ir.Path, Line: proc.Symbol.DeclarationRange.StartLine})
		}
	}
	resolver := procedureir.NewSymbolResolver(symbols)
	documents := make([]Document, len(irs))
	for i, ir := range irs {
		resolved := procedureir.Resolve(ir, resolver)
		documents[i] = Document{IR: resolved, CFG: cfg.BuildDocument(resolved)}
	}
	return Build(documents)
}

func manualProcedure(name string, line int, calls []procedureir.CallSite) procedureir.ProcedureIR {
	statements := make([]procedureir.Statement, len(calls))
	for i := range calls {
		calls[i].StatementID = i + 1
		statements[i] = procedureir.Statement{ID: i + 1, Kind: procedureir.StatementCall}
	}
	return procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{Name: name, QualifiedName: name, Kind: procedureir.ProcedureSub, DeclarationRange: rangeAt(line)}, Statements: statements, Calls: calls}
}

func manualCall(id int, status procedureir.ResolutionStatus) procedureir.CallSite {
	return procedureir.CallSite{ID: id, StatementID: id, Callee: procedureir.Callee{Text: "Call"}, Resolution: procedureir.CallResolution{Status: status}}
}

func rangeAt(line int) vbaast.Range { return vbaast.Range{StartLine: line} }

func find(t *testing.T, project ProjectSummary, qualified string) ProcedureSummary {
	t.Helper()
	for _, summary := range project.All() {
		if summary.Identity.QualifiedName == qualified {
			return summary
		}
	}
	t.Fatalf("summary %q not found: %#v", qualified, project.All())
	return ProcedureSummary{}
}

func count(evidence []Evidence, kind EffectKind) int {
	n := 0
	for _, item := range evidence {
		if item.Effect == kind {
			n++
		}
	}
	return n
}

func uncertaintyKinds(items []CallUncertainty) []UncertaintyKind {
	out := make([]UncertaintyKind, len(items))
	for i := range items {
		out[i] = items[i].Kind
	}
	return out
}
