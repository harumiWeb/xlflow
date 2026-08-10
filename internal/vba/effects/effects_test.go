package effects

import (
	"path/filepath"
	"reflect"
	"strings"
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

func TestErrorSummaryClassifiesHandlerOutcomes(t *testing.T) {
	summary := buildSources(t, sourceFile{"Errors.bas", "Errors", `Public Sub Swallow()
    On Error GoTo Handler
    Workbooks.Open "missing.xlsx"
    Exit Sub
Handler:
    Debug.Print Err.Description
End Sub

Public Sub Rethrow()
    On Error GoTo Handler
    Workbooks.Open "missing.xlsx"
    Exit Sub
Handler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub PartialRethrow()
    On Error GoTo Handler
    Workbooks.Open "missing.xlsx"
    Exit Sub
Handler:
    If Err.Number = 53 Then
        Err.Raise Err.Number
    End If
End Sub

Public Sub CleanupRethrow()
    On Error GoTo Cleanup
    Workbooks.Open "missing.xlsx"
Cleanup:
    Debug.Print "cleanup"
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Public Sub ZeroGuardedRethrow()
    On Error GoTo Handler
    Workbooks.Open "missing.xlsx"
    Exit Sub
Handler:
    If Err.Number = 0 Then Exit Sub
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExpectedErrorRecovery()
    Dim number As Long
    On Error GoTo Handler
    Workbooks.Open "missing.xlsx"
    Exit Sub
Handler:
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    If number <> 9 Then
        Err.Raise number
    End If
End Sub

Public Function TryOpen() As Boolean
    On Error GoTo Failed
    Workbooks.Open "book.xlsx"
    TryOpen = True
    Exit Function
Failed:
    TryOpen = False
End Function

Public Function ValueOrEmpty() As String
    On Error GoTo Failed
    ValueOrEmpty = "value"
    Exit Function
Failed:
    ValueOrEmpty = vbNullString
End Function

Public Function ObjectOrNothing() As Object
    On Error GoTo Failed
    Set ObjectOrNothing = CreateObject("Scripting.Dictionary")
    Exit Function
Failed:
    Set ObjectOrNothing = Nothing
End Function

Public Sub Recover()
    On Error GoTo Failed
    Workbooks.Open "book.xlsx"
    Exit Sub
Failed:
    Resume Next
End Sub

Private Sub WriteDiagnostic(ByVal message As String)
    Debug.Print message
End Sub

Private Sub ForwardDiagnostic(ByVal message As String)
    WriteDiagnostic message
End Sub

Private Sub IgnoresSecondArgument(ByVal message As String, ByVal ignored As String)
    Debug.Print message
End Sub

Public Sub WrapperLog()
    On Error GoTo Handler
    Workbooks.Open "book.xlsx"
    Exit Sub
Handler:
    ForwardDiagnostic Err.Description
End Sub

Public Sub LoggerShapedNameOnly()
    On Error GoTo Handler
    Workbooks.Open "book.xlsx"
    Exit Sub
Handler:
    LogError Err.Description
End Sub


Public Sub ErrorPassedToUnusedLoggerArgument()
    On Error GoTo Handler
    Workbooks.Open "book.xlsx"
    Exit Sub
Handler:
    IgnoresSecondArgument "constant", Err.Description
End Sub


Public Sub CopiedErrorNotUsedBySink()
    Dim e As String
    On Error GoTo Handler
    Workbooks.Open "book.xlsx"
    Exit Sub
Handler:
    e = Err.Description
    Debug.Print "unrelated"
End Sub
`})

	swallow := find(t, summary, "Errors.Swallow")
	if !swallow.Error.HasErrorHandler || !swallow.Error.SuppressesErrors || !swallow.Error.LogsAndContinues {
		t.Fatalf("swallow summary = %#v", swallow.Error)
	}
	if evidence := firstErrorEvidence(swallow.Error.Direct, ErrorSuppresses); evidence.Target != errorTargetHandler || evidence.Value != "handler" {
		t.Fatalf("swallow evidence = %#v", evidence)
	}

	rethrow := find(t, summary, "Errors.Rethrow")
	if !rethrow.Error.RethrowsErrors || !rethrow.Error.MayRaise || rethrow.Error.SuppressesErrors {
		t.Fatalf("rethrow summary = %#v", rethrow.Error)
	}
	partial := find(t, summary, "Errors.PartialRethrow")
	if !partial.Error.RethrowsErrors || !partial.Error.SuppressesErrors {
		t.Fatalf("partial rethrow summary = %#v", partial.Error)
	}
	cleanup := find(t, summary, "Errors.CleanupRethrow")
	if !cleanup.Error.RethrowsErrors || cleanup.Error.SuppressesErrors {
		t.Fatalf("cleanup rethrow summary = %#v", cleanup.Error)
	}
	zeroGuard := find(t, summary, "Errors.ZeroGuardedRethrow")
	if !zeroGuard.Error.RethrowsErrors || zeroGuard.Error.SuppressesErrors {
		t.Fatalf("zero-guarded rethrow summary = %#v", zeroGuard.Error)
	}
	expectedRecovery := find(t, summary, "Errors.ExpectedErrorRecovery")
	if !expectedRecovery.Error.RethrowsErrors || expectedRecovery.Error.SuppressesErrors {
		t.Fatalf("expected error recovery summary = %#v", expectedRecovery.Error)
	}

	tryOpen := find(t, summary, "Errors.TryOpen")
	if !tryOpen.Error.ReturnsSuccessFlag || tryOpen.Error.SuppressesErrors {
		t.Fatalf("Try-style summary = %#v", tryOpen.Error)
	}

	fallback := find(t, summary, "Errors.ValueOrEmpty")
	if fallback.Error.ReturnsSuccessFlag || fallback.Error.SuppressesErrors {
		t.Fatalf("fallback summary = %#v", fallback.Error)
	}
	objectFallback := find(t, summary, "Errors.ObjectOrNothing")
	if objectFallback.Error.SuppressesErrors {
		t.Fatalf("object fallback summary = %#v", objectFallback.Error)
	}

	recover := find(t, summary, "Errors.Recover")
	if recover.Error.UsesResumeNext || recover.Error.SuppressesErrors {
		t.Fatalf("Resume recovery summary = %#v", recover.Error)
	}
	wrapper := find(t, summary, "Errors.WrapperLog")
	if !wrapper.Error.SuppressesErrors || !wrapper.Error.LogsAndContinues {
		t.Fatalf("local logger wrapper summary = %#v", wrapper.Error)
	}
	nameOnly := find(t, summary, "Errors.LoggerShapedNameOnly")
	if !nameOnly.Error.SuppressesErrors || nameOnly.Error.LogsAndContinues {
		t.Fatalf("logger-shaped unresolved call summary = %#v", nameOnly.Error)
	}
	logger := find(t, summary, "Errors.WriteDiagnostic")
	if logger.Error.LogsAndContinues {
		t.Fatalf("ordinary logger body is not itself logs-and-continues = %#v", logger.Error)
	}
	unused := find(t, summary, "Errors.ErrorPassedToUnusedLoggerArgument")
	if !unused.Error.SuppressesErrors || unused.Error.LogsAndContinues {
		t.Fatalf("error passed to unused logger argument = %#v", unused.Error)
	}
	copied := find(t, summary, "Errors.CopiedErrorNotUsedBySink")
	if !copied.Error.SuppressesErrors || copied.Error.LogsAndContinues {
		t.Fatalf("unrelated sink consumed copied error = %#v", copied.Error)
	}
}

func TestErrorSummaryDoesNotTreatOrdinaryBooleanPredicateAsSuccessContract(t *testing.T) {
	summary := buildSources(t, sourceFile{"Predicate.bas", "Predicate", `Public Function IsPositive(ByVal value As Long) As Boolean
    If value > 0 Then
        IsPositive = True
    Else
        IsPositive = False
    End If
End Function
`})
	predicate := find(t, summary, "Predicate.IsPositive")
	if predicate.Error.ReturnsSuccessFlag {
		t.Fatalf("ordinary predicate summary = %#v", predicate.Error)
	}
}

func TestErrorSummaryDistinguishesCheckedAndUncheckedResumeNext(t *testing.T) {
	summary := buildSources(t, sourceFile{"Probe.bas", "Probe", `Public Sub CheckedProbe()
    On Error Resume Next
    Workbooks.Open "optional.xlsx"
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
End Sub

Public Sub UncheckedProbe()
    On Error Resume Next
    Workbooks.Open "one.xlsx"
    Workbooks.Open "two.xlsx"
End Sub

Public Sub UnrelatedCheck()
    On Error Resume Next
    Workbooks.Open "optional.xlsx"
    If ThisWorkbook.Saved Then Debug.Print "saved"
    On Error GoTo 0
End Sub

Public Sub CapturedErrProbe()
    Dim failed As Boolean
    Dim detail As String
    On Error Resume Next
    Workbooks.Open "optional.xlsx"
    detail = CStr(Err.Number) & ": " & Err.Description
    failed = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    If failed Then Debug.Print detail
End Sub

Public Sub RestoredResultProbe()
    Dim value As Variant
    On Error Resume Next
    value = Workbooks("optional.xlsx")
    On Error GoTo 0
    If IsEmpty(value) Then Exit Sub
End Sub

Public Sub RestoredDerivedProbe()
    Dim value As Variant
    Dim kind As VbVarType
    On Error Resume Next
    value = Workbooks("optional.xlsx")
    On Error GoTo 0
    kind = VarType(value)
    If kind = vbEmpty Then Exit Sub
End Sub

Public Sub RestoredBranchDerivedProbe(ByVal input As Variant)
    Dim value As Variant
    Dim kind As VbVarType
    If IsObject(input) Then
        On Error Resume Next
        value = input
        On Error GoTo 0
        kind = VarType(value)
    Else
        kind = VarType(input)
    End If
    If kind = vbEmpty Then Exit Sub
End Sub

Public Sub AssertedErrProbe()
    On Error Resume Next
    Workbooks.Open "optional.xlsx"
    Debug.Assert Err.Number <> 0
    On Error GoTo 0
End Sub

Public Sub DerivedBooleanProbe(ByRef values() As Variant)
    Dim hasAny As Boolean
    On Error Resume Next
    hasAny = (UBound(values) >= LBound(values))
    On Error GoTo 0
    If Not hasAny Then Err.Raise 5
End Sub

Public Function ExplicitResumeFallback() As Boolean
    ExplicitResumeFallback = False
    On Error Resume Next
    ExplicitResumeFallback = (Workbooks.Count > 0)
    On Error GoTo 0
End Function
`})
	checked := find(t, summary, "Probe.CheckedProbe")
	if !checked.Error.UsesResumeNext || checked.Error.SuppressesErrors {
		t.Fatalf("checked probe summary = %#v", checked.Error)
	}
	unchecked := find(t, summary, "Probe.UncheckedProbe")
	if !unchecked.Error.UsesResumeNext || !unchecked.Error.SuppressesErrors {
		t.Fatalf("unchecked probe summary = %#v", unchecked.Error)
	}
	if evidence := firstErrorEvidence(unchecked.Error.Direct, ErrorSuppresses); evidence.Target != errorTargetResumeNext {
		t.Fatalf("unchecked probe evidence = %#v", evidence)
	}
	unrelated := find(t, summary, "Probe.UnrelatedCheck")
	if !unrelated.Error.SuppressesErrors {
		t.Fatalf("unrelated condition accepted as probe check: %#v", unrelated.Error)
	}
	captured := find(t, summary, "Probe.CapturedErrProbe")
	if !captured.Error.UsesResumeNext || captured.Error.SuppressesErrors {
		t.Fatalf("captured Err probe summary = %#v", captured.Error)
	}
	for _, name := range []string{"Probe.RestoredResultProbe", "Probe.RestoredDerivedProbe", "Probe.RestoredBranchDerivedProbe", "Probe.AssertedErrProbe", "Probe.DerivedBooleanProbe", "Probe.ExplicitResumeFallback"} {
		probe := find(t, summary, name)
		if probe.Error.SuppressesErrors {
			t.Fatalf("checked probe %s summary = %#v", name, probe.Error)
		}
	}
}

func TestErrorSummaryAcceptsRethrowWrapperAndExplicitFallbacks(t *testing.T) {
	summary := buildSources(t, sourceFile{"Fallbacks.bas", "Fallbacks", `Private Sub RaiseAgain()
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub StopNow()
    End
End Sub

Public Sub WrappedRethrow()
    On Error GoTo Failed
    Err.Raise 5
    Exit Sub
Failed:
    RaiseAgain
End Sub

Public Sub TerminalHandler()
    On Error GoTo Failed
    Err.Raise 5
    Exit Sub
Failed:
    StopNow
End Sub

Public Function Preinitialized() As Boolean
    Preinitialized = False
    On Error GoTo Failed
    Err.Raise 5
    Preinitialized = True
    Exit Function
Failed:
End Function

Public Function InitializedAfterSetup() As Boolean
    On Error GoTo Failed
    InitializedAfterSetup = False
    Err.Raise 5
    InitializedAfterSetup = True
    Exit Function
Failed:
End Function

Public Function Composite() As ShellResult
    On Error GoTo Failed
    Err.Raise 5
    Exit Function
Failed:
    Composite.ExitCode = Err.Number
End Function
`})
	for _, name := range []string{"Fallbacks.WrappedRethrow", "Fallbacks.TerminalHandler", "Fallbacks.Preinitialized", "Fallbacks.InitializedAfterSetup", "Fallbacks.Composite"} {
		got := find(t, summary, name)
		if got.Error.SuppressesErrors {
			t.Fatalf("explicit contract %s summary = %#v", name, got.Error)
		}
	}
	if got := find(t, summary, "Fallbacks.WrappedRethrow"); !got.Error.RethrowsErrors {
		t.Fatalf("wrapped rethrow summary = %#v", got.Error)
	}
	if got := find(t, summary, "Fallbacks.Preinitialized"); !got.Error.ReturnsSuccessFlag {
		t.Fatalf("preinitialized Boolean summary = %#v", got.Error)
	}
}

func TestErrorSummaryMayRaiseIncludesUnhandledCFGPath(t *testing.T) {
	summary := buildSources(t, sourceFile{"Unhandled.bas", "Unhandled", `Public Sub Run()
    Workbooks.Open "missing.xlsx"
End Sub
`})
	run := find(t, summary, "Unhandled.Run")
	if !run.Error.MayRaise {
		t.Fatalf("unhandled fault path summary = %#v", run.Error)
	}
	evidence := firstErrorEvidence(run.Error.Direct, ErrorMayRaise)
	if evidence.Target != "unhandled_fault" || evidence.StatementID == 0 {
		t.Fatalf("unhandled fault evidence = %#v", evidence)
	}
}

func TestErrorSummaryPropagatesProvenanceThroughUniqueCalls(t *testing.T) {
	summary := buildSources(t,
		sourceFile{"Entry.bas", "Entry", "Public Sub EntryPoint()\n Middle\nEnd Sub\n"},
		sourceFile{"Middle.bas", "Middle", "Public Sub Middle()\n Leaf\nEnd Sub\n"},
		sourceFile{"Leaf.bas", "Leaf", `Public Sub Leaf()
    On Error GoTo Cleanup
    Workbooks.Open "missing.xlsx"
    Exit Sub
Cleanup:
End Sub
`},
	)
	entry := find(t, summary, "Entry.EntryPoint")
	if !entry.Error.SuppressesErrors || len(entry.Error.Propagated) == 0 {
		t.Fatalf("entry error summary = %#v", entry.Error)
	}
	evidence := firstErrorEvidence(entry.Error.Propagated, ErrorSuppresses)
	if evidence.Origin.QualifiedName != "Leaf.Leaf" || evidence.Target != errorTargetCleanup {
		t.Fatalf("propagated provenance = %#v", evidence)
	}
	if len(evidence.CallChain) != 3 || evidence.CallChain[0].QualifiedName != "Entry.EntryPoint" ||
		evidence.CallChain[1].QualifiedName != "Middle.Middle" || evidence.CallChain[2].QualifiedName != "Leaf.Leaf" {
		t.Fatalf("representative call chain = %#v", evidence.CallChain)
	}
	entry.Error.Propagated[0].CallChain[0].QualifiedName = "mutated"
	fresh := find(t, summary, "Entry.EntryPoint")
	if fresh.Error.Propagated[0].CallChain[0].QualifiedName == "mutated" {
		t.Fatal("ProjectSummary.Lookup returned an aliased error call chain")
	}
}

func TestErrorSummaryPreservesExternalCallUncertainty(t *testing.T) {
	doc := procedureir.DocumentIR{Path: "Calls.bas", ModuleName: "Calls", Procedures: []procedureir.ProcedureIR{
		manualProcedure("Calls.Root", 1, []procedureir.CallSite{manualCall(1, procedureir.ResolutionExternal)}),
	}}
	root := find(t, Build([]Document{{IR: doc, CFG: cfg.BuildDocument(doc)}}), "Calls.Root")
	if root.Error.SuppressesErrors || !root.Error.MayRaise {
		t.Fatalf("external call error outcome = %#v", root.Error)
	}
	if len(root.DirectUncertainty) != 1 || root.DirectUncertainty[0].Kind != UncertaintyExternal {
		t.Fatalf("external uncertainty = %#v", root.DirectUncertainty)
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

func TestMembershipIndexComputesOneKeyPerAdd(t *testing.T) {
	const size = 4096
	keyCalls := 0
	index := newMembershipIndex(size, func(value int) string {
		keyCalls++
		return decimal(value)
	})
	for value := 0; value < size; value++ {
		if !index.add(value) {
			t.Fatalf("first add of %d was rejected", value)
		}
	}
	for value := 0; value < size; value++ {
		if index.add(value) {
			t.Fatalf("duplicate add of %d was accepted", value)
		}
	}
	if want := 2 * size; keyCalls != want {
		t.Fatalf("key calls = %d, want %d", keyCalls, want)
	}
}

func TestPropagationMembershipPreservesFirstSeenOrderAndExactSlices(t *testing.T) {
	callerID := ProcedureIdentity{File: "Caller.bas", Module: "Caller", QualifiedName: "Caller.Run", Kind: procedureir.ProcedureSub, DeclarationLine: 1}
	calleeID := ProcedureIdentity{File: "Callee.bas", Module: "Callee", QualifiedName: "Callee.Run", Kind: procedureir.ProcedureSub, DeclarationLine: 1}
	uncertainties := []CallUncertainty{
		{Kind: UncertaintyUnresolved, Origin: calleeID, CallID: 3, Callee: "Third"},
		{Kind: UncertaintyUnresolved, Origin: calleeID, CallID: 2, Callee: "Second"},
		{Kind: UncertaintyUnresolved, Origin: calleeID, CallID: 1, Callee: "First"},
	}
	summaries := map[string]*ProcedureSummary{
		callerID.Key(): {Identity: callerID},
		calleeID.Key(): {Identity: calleeID, DirectUncertainty: uncertainties},
	}
	propagate(summaries, []edge{
		{from: callerID.Key(), to: calleeID.Key()},
		{from: callerID.Key(), to: calleeID.Key()},
	})

	caller := summaries[callerID.Key()]
	if !reflect.DeepEqual(caller.PropagatedUncertainty, uncertainties) {
		t.Fatalf("propagated uncertainty changed order or contents:\ngot  %#v\nwant %#v", caller.PropagatedUncertainty, uncertainties)
	}
	if caller.DirectUncertainty != nil {
		t.Fatalf("direct uncertainty was mutated: %#v", caller.DirectUncertainty)
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
	project := buildSources(t, sourceFile{"State.bas", "State", "Public Sub Run()\n Application.EnableEvents = False\n On Error Resume Next\nEnd Sub\n"})
	all := project.All()
	id := all[0].Identity
	all[0].Identity.Name = "mutated"
	all[0].Direct[0].Target = "mutated"
	all[0].Error.Direct[0].Target = "mutated"

	first, ok := project.Lookup(id)
	if !ok {
		t.Fatal("summary lookup failed")
	}
	first.Direct[0].Target = "mutated again"
	second, ok := project.Lookup(id)
	if !ok || second.Identity.Name != "Run" || second.Direct[0].Target == "mutated" || second.Direct[0].Target == "mutated again" || second.Error.Direct[0].Target == "mutated" {
		t.Fatalf("project summary was mutated through a returned value: %#v", second)
	}
}

func TestProjectSummaryLookupCandidateMatchesResolverIdentity(t *testing.T) {
	path := filepath.Join("Dir", "State.bas")
	project := buildSources(t, sourceFile{path, "State", "Public Sub Run()\n Application.EnableEvents = False\nEnd Sub\n"})
	candidate := procedureir.Candidate{
		File:          filepath.FromSlash("DIR/STATE.BAS"),
		QualifiedName: "state.run",
		Kind:          "SUB",
		Line:          1,
	}

	first, ok := project.LookupCandidate(candidate)
	if !ok {
		t.Fatal("case-insensitive, slash-normalized candidate did not match")
	}
	first.Direct[0].Target = "mutated"
	second, ok := project.LookupCandidate(candidate)
	if !ok || second.Direct[0].Target == "mutated" {
		t.Fatalf("candidate lookup did not return a defensive copy: %#v", second)
	}
}

func TestProjectSummaryLookupCandidateRequiresCompleteIdentity(t *testing.T) {
	project := buildSources(t, sourceFile{"State.bas", "State", "Public Sub Run()\nEnd Sub\n"})
	matching := procedureir.Candidate{File: "State.bas", QualifiedName: "State.Run", Kind: "sub", Line: 1}
	tests := map[string]procedureir.Candidate{
		"file":           {File: "Other.bas", QualifiedName: matching.QualifiedName, Kind: matching.Kind, Line: matching.Line},
		"qualified name": {File: matching.File, QualifiedName: "State.Other", Kind: matching.Kind, Line: matching.Line},
		"kind":           {File: matching.File, QualifiedName: matching.QualifiedName, Kind: "function", Line: matching.Line},
		"line":           {File: matching.File, QualifiedName: matching.QualifiedName, Kind: matching.Kind, Line: 2},
	}
	for name, candidate := range tests {
		t.Run(name, func(t *testing.T) {
			if _, ok := project.LookupCandidate(candidate); ok {
				t.Fatalf("mismatched candidate unexpectedly resolved: %#v", candidate)
			}
		})
	}
}

func TestProjectSummaryLookupCandidatePreservesUnicodeCaseFolding(t *testing.T) {
	doc := procedureir.DocumentIR{
		Path:       "Σ.bas",
		ModuleName: "Σ",
		Procedures: []procedureir.ProcedureIR{{Symbol: procedureir.ProcedureSymbol{
			Name: "Run", QualifiedName: "Σ.Run", Kind: procedureir.ProcedureSub,
			DeclarationRange: rangeAt(1),
		}}},
	}
	project := Build([]Document{{IR: doc}})
	candidate := procedureir.Candidate{File: "ς.BAS", QualifiedName: "ς.run", Kind: "SUB", Line: 1}
	if _, ok := project.LookupCandidate(candidate); !ok {
		t.Fatal("candidate lookup lost strings.EqualFold Unicode semantics")
	}
}

func TestProjectSummaryLookupCandidateUsesFirstDeterministicDuplicate(t *testing.T) {
	procedure := func(module string) procedureir.ProcedureIR {
		return procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{
			Name: "Run", QualifiedName: "Shared.Run", Kind: procedureir.ProcedureSub,
			DeclarationRange: rangeAt(1),
		}}
	}
	documents := func(first, second string) []Document {
		return []Document{
			{IR: procedureir.DocumentIR{Path: "Shared.bas", ModuleName: first, Procedures: []procedureir.ProcedureIR{procedure(first)}}},
			{IR: procedureir.DocumentIR{Path: "Shared.bas", ModuleName: second, Procedures: []procedureir.ProcedureIR{procedure(second)}}},
		}
	}
	candidate := procedureir.Candidate{File: "Shared.bas", QualifiedName: "Shared.Run", Kind: "sub", Line: 1}
	for _, project := range []ProjectSummary{Build(documents("Z", "A")), Build(documents("A", "Z"))} {
		got, ok := project.LookupCandidate(candidate)
		if !ok || !strings.EqualFold(got.Identity.Module, "A") {
			t.Fatalf("duplicate candidate selected %q, want deterministic first module A", got.Identity.Module)
		}
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

func firstErrorEvidence(items []ErrorEvidence, behavior ErrorBehaviorKind) ErrorEvidence {
	for _, item := range items {
		if item.Behavior == behavior {
			return item
		}
	}
	return ErrorEvidence{}
}
