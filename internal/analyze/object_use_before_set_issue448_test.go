package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA202Issue448TracksObjectStateAcrossCFGAndCalls(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Private Sub Init(ByRef target As Worksheet, ByVal assignIt As Boolean)
  If assignIt Then Set target = ThisWorkbook.Worksheets(1)
End Sub

Private Sub ClearTarget(ByRef target As Worksheet)
  Set target = Nothing
End Sub

Public Function MaybeSheet(ByVal returnIt As Boolean) As Worksheet
  If returnIt Then Set MaybeSheet = ThisWorkbook.Worksheets(1)
End Function

Public Sub Run(ByVal assignIt As Boolean)
  Dim ws As Worksheet
  Init ws, assignIt
  Debug.Print ws.Name
  Debug.Print sharedSheet.Name
End Sub

Public Sub ResetUse()
  Dim resetTarget As Worksheet
  Set resetTarget = ThisWorkbook.Worksheets(1)
  ClearTarget resetTarget
  Debug.Print resetTarget.Name
End Sub

Public Sub ReturnUse()
  Dim maybeTarget As Worksheet
  Set maybeTarget = MaybeSheet(False)
  Debug.Print maybeTarget.Name
End Sub

Public Sub EarlyExitUse(ByVal skip As Boolean)
  Dim earlyTarget As Worksheet
  If skip Then Exit Sub
  Debug.Print earlyTarget.Name
End Sub

Public Sub ErrorPathUse(ByVal fail As Boolean)
  Dim errorTarget As Worksheet
  On Error GoTo Handler
  If fail Then Err.Raise 5
  Set errorTarget = ThisWorkbook.Worksheets(1)
Handler:
  Debug.Print errorTarget.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	wantLines := []int{19, 20, 27, 33, 39, 48}
	if len(got) != len(wantLines) {
		t.Fatalf("VBA202 findings = %+v, want lines %v", got, wantLines)
	}
	for index, line := range wantLines {
		if got[index].Line != line {
			t.Fatalf("VBA202 finding %d = %+v, want line %d", index, got[index], line)
		}
	}
}

func TestVBA202Issue448PreservesIndexedDictionaryWrite(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub IndexedWrite()
  Dim dict As Object
  Set dict = CreateObject("Scripting.Dictionary")
  dict("key") = "value"
  Debug.Print dict("key")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("dominating construction must survive indexed assignment: %+v", got)
	}
}

func TestVBA202Issue448RecognizesAsNewAndConstructors(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim dict As New Scripting.Dictionary
  Dim created As Object
  Set created = CreateObject("Scripting.Dictionary")
  Debug.Print dict.Count
  Debug.Print created.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("constructor-backed objects should not produce VBA202: %+v", got)
	}
}

func TestVBA202Issue448ChecksCollectionAndDictionaryDefaultItems(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Unsafe()
  Dim values As Collection
  Debug.Print values(1)
End Sub

Public Sub Safe()
  Dim values As Object
  Set values = CreateObject("Scripting.Dictionary")
  Debug.Print values("key")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Unsafe" {
		t.Fatalf("default-item receiver findings = %+v, want only Unsafe", got)
	}
}

func TestVBA202Issue448PropagatesModuleFieldInitializationFromCallee(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Private Sub InitializeModuleState()
  Set sharedSheet = ThisWorkbook.Worksheets(1)
End Sub

Public Sub UseModuleState()
  InitializeModuleState
  Debug.Print sharedSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("module initializer should establish the field before use: %+v", got)
	}
}

func TestVBA202Issue448DoesNotAssumeUncalledInitializerRuns(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Private Sub Initialize()
  Set sharedSheet = ThisWorkbook.Worksheets(1)
End Sub

Public Sub UseWithoutCallingInitializer()
  Debug.Print sharedSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "UseWithoutCallingInitializer" {
		t.Fatalf("an uncalled procedure must not initialize module state: %+v", got)
	}
}

func TestVBA202Issue448RejectsConditionalModuleInitializer(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Private Sub Initialize(ByVal assignIt As Boolean)
  If assignIt Then Set sharedSheet = ThisWorkbook.Worksheets(1)
End Sub

Public Sub UseAfterConditionalInitializer()
  Initialize True
  Debug.Print sharedSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "UseAfterConditionalInitializer" {
		t.Fatalf("a conditional initializer must not establish module state: %+v", got)
	}
}

func TestVBA202Issue448DoesNotLiftUnrelatedModuleAssignmentsIntoEntryState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private mGuard As Worksheet
Private mTarget As Worksheet

Private Sub GuardOnly()
  If mGuard Is Nothing Then Exit Sub
End Sub

Private Sub UncalledAssignment()
  Set mTarget = ThisWorkbook.Worksheets(1)
End Sub

Public Sub UseTargetWithoutInitialization()
  Debug.Print mTarget.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "UseTargetWithoutInitialization" {
		t.Fatalf("unrelated guard and assignment must not initialize module state: %+v", got)
	}
}

func TestVBA202Issue448TracksModuleResetAfterLifecycleInitialization(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Private Sub InitializeModuleState()
  Set sharedSheet = ThisWorkbook.Worksheets(1)
  Set sharedSheet = Nothing
  Debug.Print sharedSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 7 {
		t.Fatalf("module reset should invalidate lifecycle initialization: %+v", got)
	}
}

func TestVBA202Issue448RefinesNotNothingGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Guard(ByVal candidate As Worksheet)
  If Not candidate Is Nothing Then
    Debug.Print candidate.Name
  End If
End Sub

Public Sub InlineGuard(ByVal candidate As Worksheet)
  If candidate Is Nothing Then Set candidate = ThisWorkbook.Worksheets(1)
  Debug.Print candidate.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Not-Is-Nothing guard should refine the true branch: %+v", got)
	}
}

func TestVBA202Issue448HonorsLocalShadowingOfModuleObject(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Public Sub Shadow()
  Dim sharedSheet As Worksheet
  Set Main.sharedSheet = ThisWorkbook.Worksheets(1)
  Debug.Print sharedSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Shadow" {
		t.Fatalf("shadowing local should remain nullable despite module initialization: %+v", got)
	}
}

func TestVBA202Issue448TracksExplicitNothingReset(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ResetThenUse()
  Dim target As Worksheet
  Set target = ThisWorkbook.Worksheets(1)
  Set target = Nothing
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 6 {
		t.Fatalf("explicit Nothing reset should invalidate later use: %+v", got)
	}
}

func TestVBA202Issue448PreservesStateAcrossByValObjectArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Observe(ByVal target As Worksheet)
  Debug.Print target.Name
End Sub

Public Sub UseByVal()
  Dim target As Worksheet
  Set target = ThisWorkbook.Worksheets(1)
  Observe target
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("ByVal object calls must preserve caller state: %+v", got)
	}
}

func TestVBA202Issue448PreservesStateAcrossReadOnlyByRefObjectArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Observe(ByRef target As Worksheet)
  Debug.Print target.Name
End Sub

Public Sub UseByRef()
  Dim target As Worksheet
  Set target = ThisWorkbook.Worksheets(1)
  Observe target
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("read-only ByRef object calls must preserve caller state: %+v", got)
	}
}

func TestVBA202Issue448PreservesByRefSemanticsForParenthesizedObjectArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub AssignTarget(ByRef target As Worksheet)
  Set target = ThisWorkbook.Worksheets(1)
End Sub

Public Sub UseParenthesized()
  Dim target As Worksheet
  AssignTarget (target)
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 9 {
		t.Fatalf("parenthesized ByRef actual must not mutate the caller's object state: %+v", got)
	}
}

func TestVBA202Issue448PreservesStateAcrossParenthesizedUnresolvedObjectArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub UseUnresolvedParenthesized()
  Dim target As Worksheet
  Set target = ThisWorkbook.Worksheets(1)
  ExternalHelper (target)
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an unresolved parenthesized object actual must preserve caller state: %+v", got)
	}
}

func TestVBA202Issue448KeepsPublicByValEntryConservative(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Observe(ByVal target As Worksheet)
  Debug.Print target.Name
End Sub

Public Sub UseByVal()
  Dim target As Worksheet
  Set target = ThisWorkbook.Worksheets(1)
  Observe target
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Observe" {
		t.Fatalf("public ByVal entry must remain conservative: %+v", got)
	}
}

func TestVBA202Issue448UsesQualifiedObjectFunctionSummary(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Helpers.bas", `Option Explicit
Public Function BuildSheet() As Worksheet
  Set BuildSheet = ThisWorkbook.Worksheets(1)
End Function
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim target As Worksheet
  Set target = Helpers.BuildSheet()
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("qualified object-returning function should use its summary: %+v", got)
	}
}

func TestVBA202Issue448UsesTypedReceiverObjectFunctionSummary(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Factory.cls", `Attribute VB_Name = "Factory"
Option Explicit
Public Function BuildDictionary() As Object
  Set BuildDictionary = CreateObject("Scripting.Dictionary")
End Function
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim factory As Factory
  Dim target As Object
  Set factory = New Factory
  Set target = factory.BuildDictionary()
  Debug.Print target.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("typed receiver object-returning function should use its summary: %+v", got)
	}
}

func TestVBA202Issue448TracksExcelMemberFactoriesAndTypeNameGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ResolveTarget(ByVal wb As Workbook) As Range
  Dim currentSelection As Object
  Dim selectedRange As Range
  Set currentSelection = wb.Application.Selection
  If TypeName(currentSelection) <> "Range" Then
    Err.Raise 5
  End If
  Set selectedRange = currentSelection
  Set ResolveTarget = selectedRange.Cells(1, 1)
End Function

Public Sub Run()
  Dim wb As Workbook
  Dim target As Range
  Set wb = Application.Workbooks(1)
  Set target = ResolveTarget(wb)
  Debug.Print target.Address
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Excel member factories and TypeName guard should establish object state: %+v", got)
	}
}

func TestVBA202Issue448TracksExcelFactoryAfterPublicBoundaryGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal wb As Workbook)
  Dim targetSheet As Worksheet
  If wb Is Nothing Then Exit Sub
  Set targetSheet = wb.Worksheets(1)
  PrepareSheet targetSheet
  Debug.Print targetSheet.Name
End Sub

Private Sub PrepareSheet(ByVal targetSheet As Worksheet)
  targetSheet.Range("A1").Value2 = "ready"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a guarded Excel receiver and its factory result should remain safe across a private ByVal helper: %+v", got)
	}
}

func TestVBA202Issue448DoesNotAssumeExcelFactoryUnderResumeNext(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal wb As Workbook)
  Dim targetSheet As Worksheet
  If wb Is Nothing Then Exit Sub
  On Error Resume Next
  Set targetSheet = wb.Worksheets(1)
  On Error GoTo 0
  Debug.Print targetSheet.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 8 {
		t.Fatalf("Resume Next may leave an Excel factory result Nothing: %+v", got)
	}
}

func TestVBA202Issue448UsesUserFormControlAsPrivateByValObject(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeFormSidecar(t, dir, "Dialog.bas", `Option Explicit
Private Sub ConfigureControl(ByVal targetControl As Object)
  targetControl.Visible = True
End Sub

Private Sub UserForm_Initialize()
  ConfigureControl Me.TextBox1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a UserForm control passed to a private ByVal helper is initialized by the form: %+v", got)
	}
}

func TestVBA202Issue448RecognizesBooleanCleanupGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Decode(ByVal responseBody As Variant) As String
  Dim stream As Object
  Dim streamOpened As Boolean
  On Error GoTo ErrHandler
  Set stream = CreateObject("ADODB.Stream")
  stream.Open
  streamOpened = True
  Decode = stream.ReadText

CleanExit:
  If streamOpened Then
    stream.Close
  End If
  Exit Function

ErrHandler:
  If Not streamOpened Then Err.Raise Err.Number
  On Error GoTo CloseFailed
  stream.Close
  On Error GoTo 0
  Err.Raise Err.Number

CloseFailed:
  On Error GoTo 0
  Err.Raise Err.Number
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("boolean cleanup guard should prove stream is open before Close: %+v", got)
	}
}

func TestVBA202Issue448CoversErrorAndEarlyExitPathsAndIndexedWrites(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ErrorPath()
  Dim target As Worksheet
  On Error GoTo Handler
  Err.Raise 5
  Set target = ThisWorkbook.Worksheets(1)
Handler:
  Debug.Print target.Name
End Sub

Public Sub EarlyGoto(ByVal skipSet As Boolean)
  Dim target As Worksheet
  If skipSet Then GoTo UseTarget
  Set target = ThisWorkbook.Worksheets(1)
UseTarget:
  Debug.Print target.Name
End Sub

Public Sub IndexedWrite()
  Dim dict As Object
  Set dict = CreateObject("Scripting.Dictionary")
  dict("key") = 1
  Debug.Print dict("key")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 2 || got[0].Procedure != "ErrorPath" || got[1].Procedure != "EarlyGoto" {
		t.Fatalf("error and early-exit paths should warn while initialized indexed writes stay safe: %+v", got)
	}
}
