package analyze

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA237ReportsFailureAtOwningHandlerWithPublicContext(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Middle
End Sub

Private Sub Middle()
  Leaf
End Sub

Private Sub Leaf()
  On Error GoTo Handler
  Workbooks.Open "missing.xlsx"
  Exit Sub
Handler:
  Debug.Print Err.Description
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA237")
	if len(got) != 1 || got[0].Procedure != "Leaf" || got[0].Line != 14 {
		t.Fatalf("VBA237 handler findings = %+v, want Leaf handler only", got)
	}
	if !strings.Contains(got[0].Message, "logs a runtime error") || !strings.Contains(got[0].Reason, "Main.Run -> Main.Middle -> Main.Leaf") {
		t.Fatalf("VBA237 public-chain context = %+v", got[0])
	}
}

func TestVBA237IncludesHostEventInRepresentativeChain(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Leaf()
  On Error GoTo Handler
  Workbooks.Open "missing.xlsx"
  Exit Sub
Handler:
End Sub
`)
	workbook := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbook, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "Sheet1.cls"), []byte(`Attribute VB_Name = "Sheet1"
Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
  Main.Leaf
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA237")
	if len(got) != 1 || !strings.Contains(got[0].Reason, "Sheet1.Worksheet_Change -> Main.Leaf") {
		t.Fatalf("VBA237 host event chain = %+v", got)
	}
}

func TestVBA237AcceptsRethrowAndCheckedTryResultsButReportsIgnoredResult(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function TryWork() As Boolean
  On Error GoTo Failed
  Workbooks.Open "book.xlsx"
  TryWork = True
  Exit Function
Failed:
  TryWork = False
End Function

Public Sub CheckedDirect()
  If Not TryWork() Then Exit Sub
End Sub

Public Sub CheckedLocal()
  Dim ok As Boolean
  ok = TryWork()
  If Not ok Then Exit Sub
End Sub

Public Function Propagated() As Boolean
  Propagated = TryWork()
End Function

Public Sub Ignored()
  TryWork
End Sub

Public Sub PartiallyChecked(ByVal skip As Boolean)
  Dim ok As Boolean
  ok = TryWork()
  If skip Then Exit Sub
  If Not ok Then Exit Sub
End Sub

Private Sub Consume(ByVal value As Boolean)
End Sub

Public Sub ArgumentUseIsUncertain()
  Consume TryWork()
End Sub

Public Sub ContainerUseIsUncertain()
  Dim values(0 To 0) As Boolean
  values(0) = TryWork()
End Sub

Public Sub StoredButUnused()
  Dim ok As Boolean
  ok = TryWork()
End Sub

Public Sub Rethrow()
  On Error GoTo Handler
  Workbooks.Open "missing.xlsx"
  Exit Sub
Handler:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA237")
	if len(got) != 3 || got[0].Procedure != "Ignored" || got[0].Line != 26 ||
		got[1].Procedure != "PartiallyChecked" || got[1].Line != 31 || got[2].Procedure != "StoredButUnused" {
		t.Fatalf("VBA237 Try result findings = %+v, want ignored, partially checked, and unused stored results", got)
	}
	if !strings.Contains(got[0].Message, "ignores the Boolean success result") {
		t.Fatalf("VBA237 ignored-result message = %+v", got[0])
	}
}

func TestVBA237CanBeDisabled(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error GoTo Handler
  Workbooks.Open "missing.xlsx"
  Exit Sub
Handler:
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DisabledRules = []string{"VBA237"}
	cfg.Analyze.DetectErrorSuppressionPropagation = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA237"); len(got) != 0 {
		t.Fatalf("disabled VBA237 findings = %+v", got)
	}
}

func TestVBA237HonorsInlineSuppressionAtLossLocation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error GoTo Handler
  Workbooks.Open "missing.xlsx"
  Exit Sub
  ' xlflow:disable-next-line VBA237
Handler:
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA237"); len(got) != 0 {
		t.Fatalf("inline-suppressed VBA237 findings = %+v", got)
	}
}

func TestVBA237LeavesBroadResumeNextScopeToVBA214(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Probe()
  On Error Resume Next
  Workbooks.Open "one.xlsx"
  Workbooks.Open "two.xlsx"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA214"); len(got) != 1 {
		t.Fatalf("VBA214 findings = %+v, want the scope diagnostic", got)
	}
	if got := findingsByCode(findings, "VBA237"); len(got) != 0 {
		t.Fatalf("VBA237 duplicated VBA214 at Resume Next scope: %+v", got)
	}
}

func TestVBA237DoesNotTreatOrdinaryBooleanPredicateAsSuccessContract(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function IsPositive(ByVal value As Long) As Boolean
  If value > 0 Then
    IsPositive = True
  Else
    IsPositive = False
  End If
End Function

Public Sub Run()
  IsPositive 1
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA237"); len(got) != 0 {
		t.Fatalf("ordinary Boolean predicate produced VBA237: %+v", got)
	}
}

func TestVBA237AcceptsObservedByRefFailureOutputs(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function TryBounds(ByRef lower As Long) As Boolean
  On Error GoTo Failed
  lower = LBound(Array(1))
  TryBounds = True
  Exit Function
Failed:
  lower = -1
  TryBounds = False
End Function

Private Function TryText(ByRef result As String) As Boolean
  On Error GoTo Failed
  result = CStr(1)
  TryText = True
  Exit Function
Failed:
  result = ""
  TryText = False
End Function

Public Sub CheckedSentinel()
  Dim lower As Long
  TryBounds lower
  If lower < 0 Then Exit Sub
End Sub

Public Sub ConditionallyCheckedSentinels()
  Dim first As Long
  Dim second As Long
  TryBounds first
  TryBounds second
  If first < 0 Then Exit Sub
  If second < 0 Then Exit Sub
End Sub

Public Function ForwardFallback() As String
  TryText ForwardFallback
End Function

Public Sub IgnoredOutput()
  Dim lower As Long
  TryBounds lower
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA237")
	if len(got) != 1 || got[0].Procedure != "IgnoredOutput" {
		t.Fatalf("ByRef failure-output findings = %+v, want only IgnoredOutput", got)
	}
}
