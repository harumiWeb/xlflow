package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
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

func TestVBA202Issue448RecognizesChartObjectAndSeriesFactories(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub RenderChart(ByVal targetSheet As Worksheet)
  Dim chartObject As ChartObject
  Dim chart As Chart
  Dim chartSeries As Series

  Set chartObject = targetSheet.ChartObjects.Add(0, 0, 100, 100)
  chartObject.Placement = xlMoveAndSize
  Set chart = chartObject.Chart
  chart.ChartType = xlLine
  Set chartSeries = chart.SeriesCollection.NewSeries
  chartSeries.Name = "Trend"
End Sub

Public Sub Run()
  RenderChart ThisWorkbook.Worksheets(1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Excel ChartObject and Series factories should establish object state: %+v", got)
	}
}

func TestVBA202Issue448RecognizesUserFormControlsItem(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub TintGhostSprite(ByVal spriteContainer As Object, ByVal bodyColor As Long)
  Dim controlIndex As Long
  Dim pixelControl As Object

  If spriteContainer Is Nothing Then Exit Sub
  For controlIndex = 0 To spriteContainer.Controls.Count - 1
    Set pixelControl = spriteContainer.Controls.Item(controlIndex)
    If pixelControl.Tag = "body" Then
      pixelControl.BackColor = bodyColor
    End If
  Next controlIndex
End Sub

Public Sub Run()
  Dim frame As Object
  Set frame = CreateObject("Forms.Frame.1")
  TintGhostSprite frame, RGB(80, 120, 255)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a bounded UserForm Controls.Item result should establish object state: %+v", got)
	}
}

func TestVBA202Issue448PreservesInitializedByRefDictionaryAcrossRecursiveCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function Pass(ByVal container As Object, ByRef dict As Object) As Object
  If container Is Nothing Then
    Debug.Print dict.Exists("key")
  Else
    Set dict = Pass(Nothing, dict)
  End If
  Set Pass = dict
End Function

Public Sub Run()
  Dim dict As New Scripting.Dictionary
  Set dict = Pass(CreateObject("Scripting.Dictionary"), dict)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an initialized ByRef dictionary should remain available across recursive calls: %+v", got)
	}
}

func TestVBA202Issue448PreservesTypedInitializedByRefDictionaryAcrossRecursiveCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function PassTyped(ByVal container As Object, ByRef dict As Dictionary) As Dictionary
  If container Is Nothing Then
    Debug.Print dict.Item("key")
  Else
    Set dict = PassTyped(Nothing, dict)
  End If
  Set PassTyped = dict
End Function

Private Sub WalkTyped(ByVal items As Collection)
  Dim item As Object
  Dim dict As New Dictionary
  For Each item In items
    Set dict = PassTyped(item, dict)
  Next item
End Sub

Public Sub Run()
  Dim dict As New Dictionary
  dict.Add "key", 1
  Set dict = PassTyped(CreateObject("Scripting.Dictionary"), dict)
  WalkTyped New Collection
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a typed initialized ByRef dictionary should remain available across recursive calls: %+v", got)
	}
}

func TestVBA202Issue448DoesNotUseWrittenReturnParameterAlias(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ClearAndReturn(ByRef value As Object) As Object
  Set value = Nothing
  Set ClearAndReturn = value
End Function

Public Sub Run()
  Dim target As Object
  Set target = CreateObject("Scripting.Dictionary")
  Dim result As Object
  Set result = ClearAndReturn(target)
  Debug.Print result.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 12 {
		t.Fatalf("a return parameter cleared before return must not inherit the caller state: %+v", got)
	}
}

func TestVBA202Issue448PropagatesCollectionItemFunctionResult(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function EnsureRow(ByVal rows As Collection, ByVal rowIndex As Long) As Object
  Dim rowValues As Object
  Do While rows.Count < rowIndex
    Set rowValues = CreateObject("Scripting.Dictionary")
    rows.Add rowValues
  Loop
  Set EnsureRow = rows(rowIndex)
End Function

Public Sub Run()
  Dim rows As Collection
  Set rows = New Collection
  rows.Add CreateObject("Scripting.Dictionary")

  Dim directValue As Object
  Set directValue = rows(1)
  Debug.Print directValue.Exists("key")

  Dim helperValue As Object
  Set helperValue = EnsureRow(rows, 1)
  Debug.Print helperValue.Exists("key")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a Collection item and a helper returning one should establish object state: %+v", got)
	}
}

func TestVBA202Issue448RecognizesGuardedDictionaryCollectionItems(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub AddIndexRow(ByVal targetIndex As Object, ByVal indexKey As String, ByVal rowIndex As Long)
  Dim rows As Collection
  If targetIndex.Exists(indexKey) Then
    Set rows = targetIndex(indexKey)
  Else
    Set rows = New Collection
    targetIndex.Add indexKey, rows
  End If
  rows.Add rowIndex
End Sub

Private Function IndexedRows(ByVal targetIndex As Object, ByVal indexKey As String) As Collection
  If Not targetIndex Is Nothing Then
    If targetIndex.Exists(indexKey) Then
      Set IndexedRows = targetIndex(indexKey)
      Exit Function
    End If
  End If
  Set IndexedRows = New Collection
End Function

Public Sub Run()
  Dim targetIndex As Object
  Set targetIndex = CreateObject("Scripting.Dictionary")
  AddIndexRow targetIndex, "key", 1

  Dim rows As Collection
  Set rows = IndexedRows(targetIndex, "key")
  Debug.Print rows.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("guarded Dictionary Collection items should establish object state: %+v", got)
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

Public Sub ParenthesizedGuard(ByVal candidate As Worksheet)
  If Not (candidate Is Nothing) Then
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

func TestVBA202Issue448RecognizesTerminalGuardsAndLateBoundFactories(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub RaiseContractError()
  Err.Raise 5
End Sub

Private Sub RequireRangeTarget(ByVal target As Object)
  If target Is Nothing Then
    RaiseContractError
  End If
End Sub

Public Sub Run(ByVal target As Object, ByVal candidate As Object)
  If candidate Is Nothing Then
    RaiseContractError
  End If
  Debug.Print candidate.Name

  RequireRangeTarget target
  Debug.Print target.Name

  Dim fs As Object
  Dim textFile As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set textFile = fs.OpenTextFile("sample.txt", 1)
  Debug.Print textFile.ReadAll

  Dim binaryStream As Object
  Dim sourceStream As Object
  Set binaryStream = CreateObject("ADODB.Stream")
  Set sourceStream = CreateObject("ADODB.Stream")
  With sourceStream
    .CopyTo binaryStream
  End With
  Debug.Print binaryStream.Position
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("terminal guards and known late-bound object factories should establish object state: %+v", got)
	}
}

func TestVBA202Issue448RequiresReachingFileSystemObjectFactory(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal replaceIt As Boolean)
  Dim fs As Object
  Dim textFile As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  If replaceIt Then Set fs = CreateObject("Other.Component")
  Set textFile = fs.OpenTextFile("sample.txt", 1)
  Debug.Print textFile.ReadAll
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 8 {
		t.Fatalf("a late-bound factory after a possible receiver replacement must remain nullable: %+v", got)
	}
}

func TestVBA202Issue448DoesNotTreatResumeNextFileFactoryAsAssigned(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim fs As Object
  Dim textFile As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  On Error Resume Next
  Set textFile = fs.OpenTextFile("missing.txt", 1)
  On Error GoTo 0
  Debug.Print textFile.ReadAll
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 9 {
		t.Fatalf("a FileSystemObject factory under Resume Next must remain nullable: %+v", got)
	}
}

func TestVBA202Issue448InvalidatesUnresolvedExpressionByRefObject(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim target As Object
  Dim result As Object
  Set target = CreateObject("Scripting.Dictionary")
  Set result = ExternalObject(target)
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 7 {
		t.Fatalf("an unresolved function expression may mutate a ByRef object argument: %+v", got)
	}
}

func TestVBA202Issue448LimitsCopyToPreservationToADODBStream(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim source As Object
  Dim target As Object
  Set source = CreateObject("Other.Component")
  Set target = CreateObject("Scripting.Dictionary")
  source.CopyTo target
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 8 {
		t.Fatalf("an unrelated CopyTo method must retain ByRef mutation effects: %+v", got)
	}
}

func TestVBA202Issue448PreservesCopyToForConfirmedProjectStreamFactory(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Const StreamProgID As String = "ADODB.Stream"

Private Function CreateStreamObject(ByVal streamType As Long) As Object
  Set CreateStreamObject = CreateObject(StreamProgID)
  CreateStreamObject.Type = streamType
  CreateStreamObject.Open
End Function

Public Sub Run()
  Dim source As Object
  Dim target As Object
  Set source = CreateStreamObject(2)
  Set target = CreateStreamObject(1)
  With source
    .CopyTo target
  End With
  Debug.Print target.Position
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a confirmed project ADODB.Stream factory should preserve CopyTo destination state: %+v", got)
	}
}

func TestVBA202Issue448RequiresPredicateNonNothingContract(t *testing.T) {
	t.Parallel()
	for _, test := range []struct {
		name         string
		predicate    string
		wantFindings int
	}{
		{
			name: "guarded",
			predicate: `Private Function IsExcelTable(ByVal candidate As Object) As Boolean
  If candidate Is Nothing Then Exit Function
  IsExcelTable = (TypeName(candidate) = "ListObject")
End Function
`,
			wantFindings: 0,
		},
		{
			name: "unguarded",
			predicate: `Private Function IsExcelTable(ByVal candidate As Object) As Boolean
  IsExcelTable = True
End Function
`,
			wantFindings: 1,
		},
		{
			name: "true-before-nothing-guard",
			predicate: `Private Function IsExcelTable(ByVal candidate As Object) As Boolean
  IsExcelTable = True
  If candidate Is Nothing Then Exit Function
End Function
`,
			wantFindings: 1,
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			dir := t.TempDir()
			writeModule(t, dir, "Main.bas", `Option Explicit
`+test.predicate+`
Public Sub Run(ByVal candidate As Object)
  If IsExcelTable(candidate) Then
    Debug.Print candidate.Name
  End If
End Sub
`)

			findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
			if err != nil {
				t.Fatal(err)
			}
			if got := findingsByCode(findings, "VBA202"); len(got) != test.wantFindings {
				t.Fatalf("predicate contract %s produced %+v, want %d finding(s)", test.name, got, test.wantFindings)
			}
		})
	}
}

func TestVBA202Issue448DoesNotRefineNegatedPredicateTrueBranch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function IsExcelTable(ByVal candidate As Object) As Boolean
  If candidate Is Nothing Then Exit Function
  IsExcelTable = True
End Function

Public Sub Run(ByVal candidate As Object)
  If Not IsExcelTable(candidate) Then
    Debug.Print candidate.Name
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 9 {
		t.Fatalf("the true branch of a negated predicate remains nullable: %+v", got)
	}
}

func TestVBA202Issue448PropagatesPrivateByValObjectEntryGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub RaiseContractError()
  Err.Raise 5
End Sub

Private Sub ReadParseError(ByVal dom As Object)
  Debug.Print dom.parseError.reason
  RaiseContractError
End Sub

Public Sub Run()
  Dim dom As Object
  Set dom = CreateObject("MSXML2.DOMDocument")
  ReadParseError dom
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a private ByVal object helper called with a constructed object should have a non-Nothing entry: %+v", got)
	}
}

func TestVBA202Issue448PropagatesCallerGuardIntoPrivateByValObjectHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub ReadObject(ByVal value As Object)
  Debug.Print value.Name
End Sub

Public Sub Run(ByVal value As Object)
  If value Is Nothing Then Err.Raise 5
  ReadObject value
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a caller guard must establish a non-Nothing entry for a private ByVal object helper: %+v", got)
	}
}

func TestVBA202Issue448PropagatesGuardThroughTypeNameDispatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub WriteObjectValue(ByVal value As Object)
  If value Is Nothing Then Exit Sub
  Select Case TypeName(value)
    Case "Collection"
      WriteCollection value
    Case "Dictionary"
      WriteDictionary value
  End Select
End Sub

Private Function StringifyCollection(ByVal value As Object) As String
  WriteCollection value
End Function

Private Sub WriteCollection(ByVal value As Object)
  If value.Count = 0 Then Exit Sub
End Sub

Private Function StringifyDictionary(ByVal value As Object) As String
  WriteDictionary value
End Function

Private Sub WriteDictionary(ByVal value As Object)
  If value.Count = 0 Then Exit Sub
End Sub

Public Sub Run(ByVal value As Object)
  WriteObjectValue value
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a caller-side Nothing guard must propagate through TypeName dispatch to private ByVal object helpers: %+v", got)
	}
}

func TestVBA202Issue448PropagatesClassFunctionResultToPrivateCollectionHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "ROneCOne.cls", `Option Explicit
Public Property Get NewEnum() As IUnknown
  Set NewEnum = MaterializeValues
End Property

Private Function MaterializeValues() As Collection
  Dim values As Collection
  Set values = New Collection
  AddUnwrappedValue values
  Set MaterializeValues = values
End Function

Private Sub AddUnwrappedValue(ByVal values As Collection)
  values.Add "item"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a collection created by a reachable class helper must reach its private consumer as non-Nothing: %+v", got)
	}
}

func TestVBA202Issue448KeepsNullablePublicCollectionReceiver(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Append(ByVal values As Collection)
  values.Add "item"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Append" {
		t.Fatalf("a public Collection parameter may be Nothing: %+v", got)
	}
}

func TestVBA202Issue448KeepsNullableCollectionReceiverThroughErrorHandler(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function RecursiveDir(ByVal values As Collection) As Variant
  On Error GoTo ErrTrap
  Dim folders As New Collection
  values.Add "item"
  For Each folder In folders
    Call RecursiveDir(values)
  Next folder
ExitProcedure:
  On Error Resume Next
ErrTrap:
  Select Case Err.Number
    Case Is <> 0
      Resume ExitProcedure
    Case Else
      Resume ExitProcedure
  End Select
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "RecursiveDir" {
		t.Fatalf("an error handler does not initialize a nullable Collection parameter: %+v", got)
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

func TestVBA202Issue448DoesNotUseModuleObjectForVariantShadow(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedSheet As Worksheet

Public Sub LocalShadow()
  Dim sharedSheet As Variant
  Debug.Print sharedSheet.Cells(1, 1)
End Sub

Public Sub ParameterShadow(ByVal sharedSheet As Variant)
  Debug.Print sharedSheet.Cells(1, 1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Variant shadows must not resolve to the module object: %+v", got)
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

func TestVBA202Issue448DoesNotUseByValParameterPostconditionForCaller(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub InitializeCopy(ByVal target As Object)
  Set target = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Run()
  Dim target As Object
  InitializeCopy target
  Debug.Print target.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Run" {
		t.Fatalf("a ByVal parameter postcondition must not initialize the caller's object: %+v", got)
	}
}

func TestVBA202Issue448RecognizesIntrinsicObjectByValArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub ObserveWorkbook(ByVal target As Workbook)
  Debug.Print target.Name
End Sub

Private Sub ObserveApplication(ByVal target As Object)
  Debug.Print target.Name
End Sub

Public Sub Run()
  ObserveWorkbook ThisWorkbook
  ObserveApplication Application
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("intrinsic object arguments must preserve their non-Nothing state: %+v", got)
	}
}

func TestVBA202Issue448DoesNotInvalidateObjectStoredInDictionary(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function BuildNode(ByVal textValue As String) As Object
  Dim node As Object
  Set node = CreateObject("Scripting.Dictionary")

  Dim children As Collection
  Set children = New Collection
  node.Add "children", children
  If Len(textValue) > 0 Then children.Add textValue

  Set BuildNode = node
End Function

Public Sub Run()
  Dim node As Object
  Set node = BuildNode("value")
  Debug.Print node("children").Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a Dictionary Add must not invalidate an object argument: %+v", got)
	}
}

func TestVBA202Issue448PreservesObjectArgumentToLateBoundAdd(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub SetJsonAttr(ByVal node As Object, ByVal attrName As String, ByVal attrValue As String)
  Dim attributes As Object
  Set attributes = CreateObject("Scripting.Dictionary")
  node.Add "attr", attributes
  attributes(attrName) = attrValue
End Sub

Public Sub Run()
  Dim node As Object
  Set node = CreateObject("Scripting.Dictionary")
  SetJsonAttr node, "key", "value"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a late-bound Dictionary Add must not invalidate its object argument: %+v", got)
	}
}

func TestVBA202Issue448RequiresContainerReceiverForLateBoundAdd(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim custom As Object
  Dim value As Object
  Set custom = CreateObject("Other.Component")
  Set value = CreateObject("Scripting.Dictionary")
  custom.Add value
  Debug.Print value.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 8 {
		t.Fatalf("an unrelated late-bound Add receiver must not preserve an object argument: %+v", got)
	}
}

func TestVBA202Issue448RecognizesDictionaryFactoryBranchesForLateBoundAdd(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function CreateDictionary(ByVal useNative As Boolean) As Object
  If useNative Then
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
  Else
    Set CreateDictionary = New Dictionary
  End If
End Function

Public Sub Run()
  Dim dict As Object
  Set dict = CreateDictionary(False)
  dict.Add "item", dict
  Debug.Print dict.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a Dictionary factory with constructor branches must preserve its receiver: %+v", got)
	}
}

func TestVBA202Issue448AllowsSameFactoryReassignmentOnReachablePath(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function CreateDictionary() As Object
  Set CreateDictionary = CreateObject("Scripting.Dictionary")
End Function

Public Sub Run(ByVal resetIt As Boolean)
  Dim dict As Object
  Set dict = CreateDictionary()
  If resetIt Then Set dict = CreateDictionary()
  dict.Add "item", dict
  Debug.Print dict.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("same-progid receiver reassignments must preserve a Dictionary Add argument: %+v", got)
	}
}

func TestVBA202Issue448PropagatesNestedObjectFactoryReturn(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function NewDictionary() As Object
  Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Private Function BuildEnvelope() As Object
  Dim envelope As Object
  Set envelope = NewDictionary()
  Set BuildEnvelope = envelope
End Function

Private Sub ConsumeEnvelope(ByVal envelope As Object)
  Debug.Print envelope.Exists("respond")
End Sub

Public Sub Run()
  Dim envelope As Object
  Set envelope = BuildEnvelope()
  ConsumeEnvelope envelope
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("nested object factory returns must establish a non-Nothing ByVal argument: %+v", got)
	}
}

func TestVBA202Issue448PreservesObjectThroughArrayIntrinsicArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Invoke(ByVal callback As Object, ByVal envelope As Object)
  Call callback.RunEx(Array(envelope))
  Debug.Print envelope.Exists("respond")
End Sub

Public Sub Run()
  Dim callback As Object
  Dim envelope As Object
  Set callback = CreateObject("Scripting.Dictionary")
  Set envelope = CreateObject("Scripting.Dictionary")
  Invoke callback, envelope
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Array must not invalidate an object used as a Variant array element: %+v", got)
	}
}

func TestVBA202Issue448PreservesObjectThroughKnownByValDeclare(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pUnk As IUnknown, ByVal pHwnd As LongPtr) As Long

Private Sub ReadFrame(ByVal frm As Object)
  Dim hwnd As LongPtr
  Call IUnknown_GetWindow(frm, VarPtr(hwnd))
  Debug.Print frm.Parent.Name
End Sub

Public Sub Run(ByVal frm As Object)
  If frm Is Nothing Then Err.Raise 5
  ReadFrame frm
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a known ByVal COM declare must not invalidate its object argument: %+v", got)
	}
}

func TestVBA202Issue448PreservesNestedPrivateByValObjectArguments(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub AppendRow(ByVal rows As Collection)
  AddRow rows
End Sub

Private Sub AddRow(ByVal rows As Collection)
  rows.Add "value"
End Sub

Public Sub Run()
  Dim rows As Collection
  Set rows = New Collection
  AppendRow rows
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("nested private ByVal object arguments must preserve their state: %+v", got)
	}
}

func TestVBA202Issue448PreservesRecursivePrivateByValObjectArguments(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub AppendRows(ByVal rows As Collection, ByVal depth As Long)
  If depth > 0 Then AppendRows rows, depth - 1
  rows.Add "value"
End Sub

Public Sub Run()
  Dim rows As Collection
  Set rows = New Collection
  AppendRows rows, 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("recursive private ByVal object arguments must preserve their state: %+v", got)
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

func TestVBA202Issue448UsesClassInitializeFieldSummary(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Widget.cls", `Attribute VB_Name = "Widget"
Option Explicit
Private cached As Object

Private Sub Class_Initialize()
  Set cached = CreateObject("Scripting.Dictionary")
End Sub

Public Sub UseCached()
  Debug.Print cached.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Class_Initialize assignment should establish the class field: %+v", got)
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

func TestVBA202Issue448TracksTypeNameExcelTableBranchState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ResolveTableRange(ByVal source As Object, ByVal wantHeaders As Boolean) As Object
  Dim owner As Object
  Dim resolved As Object
  If source Is Nothing Then Err.Raise 5
  Select Case TypeName(source)
    Case "ListObject"
      Set owner = source
    Case "ListColumn"
      Set owner = source.Parent
    Case Else
      Set ResolveTableRange = source
      Exit Function
  End Select
  If wantHeaders Then
    Set resolved = source.Range.Resize(1 + owner.ListRows.Count)
  Else
    Set resolved = source.DataBodyRange
  End If
  If resolved Is Nothing Then Err.Raise 5
  Set ResolveTableRange = resolved
End Function

Public Sub Run(ByVal source As Object)
  Dim result As Object
  Set result = ResolveTableRange(source, True)
  Debug.Print result.Address
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("TypeName Excel table branches should preserve source and owner state: %+v", got)
	}
}

func TestVBA202Issue448DoesNotAssumeTypeNameExcelMemberUnderResumeNext(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal source As Object)
  Dim owner As Object
  If source Is Nothing Then Exit Sub
  Select Case TypeName(source)
    Case "ListColumn"
      On Error Resume Next
      Set owner = source.Parent
      On Error GoTo 0
      Debug.Print owner.Name
  End Select
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 10 {
		t.Fatalf("Resume Next may leave a TypeName Excel member result Nothing: %+v", got)
	}
}

func TestVBA202Issue448PropagatesIsExcelTableTypeIntoPrivateHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function IsExcelTable(ByVal candidate As Object) As Boolean
  If candidate Is Nothing Then Exit Function
  IsExcelTable = (TypeName(candidate) = "ListObject")
End Function

Private Function WriteToListObject(ByVal listObject As Object) As Long
  Dim sheet As Object
  Set sheet = listObject.Parent
  sheet.Cells(1, 1).Value = 1
  WriteToListObject = 1
End Function

Public Function ToRange(ByVal target As Object) As Long
  If Not IsExcelTable(target) Then Err.Raise 5
  ToRange = WriteToListObject(target)
End Function

Public Sub Run(ByVal target As Object)
  ToRange target
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("IsExcelTable should propagate ListObject state into the private helper: %+v", got)
	}
}

func TestVBA202Issue448TracksMsxmlSelectNodesResult(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function RunXPath(ByVal provider As Object) As Object
  If provider Is Nothing Then Err.Raise 5
  Set RunXPath = provider.SelectNodes("//item")
End Function

Public Function FirstMatch(ByVal provider As Object) As Object
  Dim matched As Object
  Set matched = RunXPath(provider)
  If matched.Length > 0 Then Set FirstMatch = matched.Item(0)
End Function

Public Sub DirectSelectNodes(ByVal provider As Object)
  Dim nodes As Object
  If provider Is Nothing Then Exit Sub
  Set nodes = provider.SelectNodes("//item")
  If nodes.Length = 0 Then Exit Sub
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("successful MSXML SelectNodes should establish a NodeList result: %+v", got)
	}
}

func TestVBA202Issue448DoesNotTreatMsxmlSelectNodesUnderResumeNextAsAssigned(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim provider As Object
  Dim nodes As Object
  Set provider = CreateObject("MSXML2.DOMDocument.6.0")
  On Error Resume Next
  Set nodes = provider.SelectNodes("//item")
  On Error GoTo 0
  Debug.Print nodes.Length
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 9 {
		t.Fatalf("SelectNodes under Resume Next must remain nullable: %+v", got)
	}
}

func TestVBA202Issue448AcceptsCheckedMsxmlSelectNodesResumeNext(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub RaiseXmlError(ByVal message As String)
  Err.Raise 5, "XML", message
End Sub

Public Sub Run()
  Dim provider As Object
  Dim nodes As Object
  Set provider = CreateObject("MSXML2.DOMDocument.6.0")
  On Error Resume Next
  Set nodes = provider.SelectNodes("//item")
  If Err.Number <> 0 Then
    On Error GoTo 0
    RaiseXmlError "XPath failed"
  End If
  On Error GoTo 0
  Debug.Print nodes.Length
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a checked SelectNodes failure must not leave a nullable result on normal continuation: %+v", got)
	}
}

func TestVBA202Issue448TracksRegExpExecuteResultThroughModuleInitializer(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.cls", `Option Explicit
Private Const REGEX_PROG_ID As String = "VBScript.RegExp"
Private mProviderObject As Object

Private Sub ConfigureRegex()
  Set mProviderObject = CreateObject(REGEX_PROG_ID)
  On Error GoTo BadPattern
  mProviderObject.Test vbNullString
  On Error GoTo 0
  Exit Sub
BadPattern:
  Set mProviderObject = Nothing
  Err.Raise 5
End Sub

Public Function MatchCount(ByVal inputText As String) As Long
  Dim matches As Object
  ConfigureRegex
  Set matches = mProviderObject.Execute(inputText)
  MatchCount = matches.Count
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("RegExp.Execute should establish a MatchCollection after the module initializer: %+v", got)
	}
}

func TestVBA202Issue448PropagatesForEachObjectIntoHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub ConsumeMatch(ByVal rawMatch As Object)
  Debug.Print rawMatch.Value
End Sub

Public Sub Run(ByVal inputText As String)
  Dim expression As Object
  Dim rawMatch As Object
  Set expression = CreateObject("VBScript.RegExp")
  For Each rawMatch In expression.Execute(inputText)
    ConsumeMatch rawMatch
  Next rawMatch
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a For Each object should remain non-Nothing when passed to a helper: %+v", got)
	}
}

func TestVBA202Issue448PropagatesForEachObjectIntoExpressionHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ReadMatch(ByVal rawMatch As Object) As String
  ReadMatch = CStr(rawMatch.Value)
End Function

Public Sub Run(ByVal inputText As String)
  Dim expression As Object
  Dim rawMatch As Object
  Dim results As Collection
  Set expression = CreateObject("VBScript.RegExp")
  Set results = New Collection
  For Each rawMatch In expression.Execute(inputText)
    results.Add ReadMatch(rawMatch)
  Next rawMatch
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a For Each object should remain non-Nothing in an expression helper: %+v", got)
	}
}

func TestVBA202Issue448PropagatesRegexMatchIntoBuilder(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Regex.cls", `Attribute VB_Name = "Regex"
Option Explicit

Private Function BuildRegexMatch(ByVal rawMatch As Object) As String
  BuildRegexMatch = CStr(rawMatch.Value)
End Function

Public Function Matches(ByVal inputText As String) As String
  Dim expression As Object
  Dim rawMatch As Object
  Set expression = CreateObject("VBScript.RegExp")
  For Each rawMatch In expression.Execute(inputText)
    Matches = BuildRegexMatch(rawMatch)
  Next rawMatch
End Function

Public Function Match(ByVal inputText As String) As String
  Dim expression As Object
  Dim firstMatches As Object
  Set expression = CreateObject("VBScript.RegExp")
  Set firstMatches = expression.Execute(inputText)
  If firstMatches.Count > 0 Then
    Match = BuildRegexMatch(firstMatches.Item(0))
  End If
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a Match object from Execute should remain non-Nothing in the builder: %+v", got)
	}
}

func TestVBA202Issue448DoesNotTreatRegExpExecuteUnderResumeNextAsAssigned(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function MatchCount(ByVal inputText As String) As Long
  Dim expression As Object
  Dim matches As Object
  Set expression = CreateObject("VBScript.RegExp")
  On Error Resume Next
  Set matches = expression.Execute(inputText)
  On Error GoTo 0
  MatchCount = matches.Count
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Line != 9 {
		t.Fatalf("RegExp.Execute under Resume Next must remain nullable: %+v", got)
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

func TestVBA202Issue448RecognizesExcelFactoryFunctionResult(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private remoteWorkbook As Workbook

Private Function createRemoteWorkbook() As Workbook
  Dim app As Application: Set app = CreateObject("Excel.Application")
  Set createRemoteWorkbook = app.Workbooks.Add
End Function

Public Sub Run()
  Set remoteWorkbook = createRemoteWorkbook()
  Call remoteWorkbook.Application.Run("TimerMain.StartTimer")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Excel factory function result should establish the object before use: %+v", got)
	}
}

func TestVBA202Issue448TracksInitializedFunctionReturnThroughByRefCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Append(ByRef output As Collection, ByVal input As Collection)
  Call output.Add(input)
End Sub

Private Function Values() As Collection
  Dim input As Collection
  Set input = New Collection
  Set Values = New Collection
  Call Append(Values, input)
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an initialized function return passed to a private ByRef helper should be safe: %+v", got)
	}
}

func TestObjectDirectCallSummaryUsesCallFileForDuplicateModules(t *testing.T) {
	t.Parallel()
	call := procedureir.CallSite{
		File:       "../../src/stdTimer.cls",
		Module:     "stdTimer",
		Resolution: procedureir.CallResolution{Status: procedureir.ResolutionAmbiguous},
		Callee:     procedureir.Callee{BaseName: "createRemoteWorkbook"},
	}
	summary, ok := objectDirectCallSummary(sourceProcedure{Module: "stdTimer"}, call, map[string]objectProcedureSummary{
		"primary": {
			File:           "src/stdTimer.cls",
			Module:         "stdTimer",
			QualifiedName:  "stdTimer.createRemoteWorkbook",
			ReturnAssigned: true,
		},
		"wip": {
			File:           "src/WIP/stdTimer.cls",
			Module:         "stdTimer",
			QualifiedName:  "stdTimer.createRemoteWorkbook",
			ReturnAssigned: false,
		},
	})
	if !ok || !summary.ReturnAssigned {
		t.Fatalf("same-file function summary = %+v, ok=%v; want assigned primary summary", summary, ok)
	}
}

func TestObjectCalleeKeyRejectsDuplicateSameModuleBareNames(t *testing.T) {
	t.Parallel()
	analysis := &objectAnalysisContext{plans: map[string]*objectProcedurePlan{
		"first":  {proc: sourceProcedure{Module: "Main", Name: "Helper"}},
		"second": {proc: sourceProcedure{Module: "Main", Name: "Helper"}},
	}}
	call := procedureir.CallSite{
		Module: "Main",
		Callee: procedureir.Callee{BaseName: "Helper"},
	}
	if key, ok := analysis.objectCalleeKey(call); ok || key != "" {
		t.Fatalf("duplicate same-module bare call = (%q, %v), want ambiguous", key, ok)
	}
}

func TestObjectCalleeKeyUsesUniqueSameModuleForAmbiguousCall(t *testing.T) {
	t.Parallel()
	analysis := &objectAnalysisContext{plans: map[string]*objectProcedurePlan{
		"only": {proc: sourceProcedure{Module: "Main", Name: "Helper"}},
	}}
	call := procedureir.CallSite{
		Module:     "Main",
		Resolution: procedureir.CallResolution{Status: procedureir.ResolutionAmbiguous},
		Callee:     procedureir.Callee{BaseName: "Helper"},
	}
	if key, ok := analysis.objectCalleeKey(call); !ok || key != "only" {
		t.Fatalf("unique same-module ambiguous call = (%q, %v), want (only, true)", key, ok)
	}
}

func TestObjectCallEffectsSkipsAmbiguousDirectSummaries(t *testing.T) {
	t.Parallel()
	value := objectVariable{Scope: procedureir.ScopeParameter, Name: "value"}
	call := procedureir.CallSite{
		File: "src/Main.bas", Module: "Main",
		Caller:    procedureir.ProcedureRef{QualifiedName: "Main.Run"},
		Callee:    procedureir.Callee{BaseName: "Touch"},
		Arguments: procedureir.Arguments{Count: 1, ExpressionIDs: []int{1}},
	}
	declarations := declarationScope{parameters: map[string]sourceDeclaration{
		"value": {Name: "value", Type: "Collection", Object: true, Parameter: true},
	}}
	vars := map[string]objectVariable{value.key(): value}
	expressions := map[int]procedureir.Expression{
		1: {ID: 1, Kind: procedureir.ExpressionIdentifier, Text: "value"},
	}
	facts := newProcedureAnalysisFacts(nil, []procedureir.Expression{expressions[1]}, nil, nil)
	summary := func() objectProcedureSummary {
		return objectProcedureSummary{
			File: "src/Main.bas", Module: "Main", QualifiedName: "Main.Touch",
			Params: []objectParameterSummary{{Name: "value", Object: true, ByRef: true}},
		}
	}
	summaries := map[string]objectProcedureSummary{"first": summary(), "second": summary()}
	state := map[string]bool{value.key(): true}
	applyObjectCallEffects(sourceProcedure{}, call, state, vars, declarations, facts, summaries)
	if state[value.key()] {
		t.Fatalf("ambiguous direct summaries must not preserve a nullable ByRef object state: %+v", state)
	}
}

func TestObjectFlowUsesLexicalBindingForVariantShadows(t *testing.T) {
	t.Parallel()
	module := map[string]sourceDeclaration{
		"sharedsheet": {Name: "sharedSheet", Type: "Worksheet", Object: true},
	}
	moduleVariable := objectVariable{Scope: procedureir.ScopeModule, Name: "sharedSheet"}
	moduleSummary := objectProcedureSummary{
		File: "src/Main.bas", Module: "Main", QualifiedName: "Main.Initialize",
		ModuleAssigned: map[string]bool{"sharedsheet": false},
		ModuleWritten:  map[string]bool{"sharedsheet": true},
	}
	actualSummary := objectProcedureSummary{
		File: "src/Main.bas", Module: "Main", QualifiedName: "Main.Touch",
		Params:        []objectParameterSummary{{Name: "value", Object: true, ByRef: true}},
		ByRefAssigned: map[int]bool{0: false},
		ByRefWritten:  map[int]bool{0: true},
	}
	expressions := map[int]procedureir.Expression{
		1: {ID: 1, Kind: procedureir.ExpressionIdentifier, Text: "sharedSheet"},
	}
	facts := newProcedureAnalysisFacts(nil, []procedureir.Expression{expressions[1]}, nil, nil)
	call := procedureir.CallSite{
		File: "src/Main.bas", Module: "Main", Caller: procedureir.ProcedureRef{QualifiedName: "Main.Run"},
		Callee: procedureir.Callee{BaseName: "Initialize"},
	}
	actualCall := procedureir.CallSite{
		File: "src/Main.bas", Module: "Main", Caller: procedureir.ProcedureRef{QualifiedName: "Main.Run"},
		Callee:    procedureir.Callee{BaseName: "Touch"},
		Arguments: procedureir.Arguments{Count: 1, ExpressionIDs: []int{1}},
	}
	for _, test := range []struct {
		name       string
		local      map[string]sourceDeclaration
		parameters map[string]sourceDeclaration
	}{
		{name: "local", local: map[string]sourceDeclaration{"sharedsheet": {Name: "sharedSheet", Type: "Variant"}}},
		{name: "parameter", parameters: map[string]sourceDeclaration{"sharedsheet": {Name: "sharedSheet", Type: "Variant", Parameter: true}}},
	} {
		t.Run(test.name, func(t *testing.T) {
			declarations := declarationScope{module: module, local: test.local, parameters: test.parameters}
			state := map[string]bool{moduleVariable.key(): true}
			vars := map[string]objectVariable{moduleVariable.key(): moduleVariable}
			applyObjectCallEffects(sourceProcedure{}, call, state, vars, declarations, nil, map[string]objectProcedureSummary{"initialize": moduleSummary})
			if !state[moduleVariable.key()] {
				t.Fatalf("initialized module object was changed through a shadowed Variant: %+v", state)
			}
			applyObjectCallEffects(sourceProcedure{}, actualCall, state, vars, declarations, facts, map[string]objectProcedureSummary{"touch": actualSummary})
			if !state[moduleVariable.key()] {
				t.Fatalf("ByRef actual resolution selected the module object through a shadowed Variant: %+v", state)
			}
			assigned, present := objectCallParameterAssigned(sourceProcedure{Name: "Run"}, declarations, actualCall, actualSummary, 0, objectCallActuals(actualCall, facts), state, vars, objectFlowContext{facts: facts}, map[string]objectProcedureSummary{})
			if assigned || present {
				t.Fatalf("entry-call actual resolution selected a shadowed Variant: assigned=%v present=%v", assigned, present)
			}
		})
	}
}

func TestVBA202Issue448PreservesObjectStateAcrossUnresolvedPrivateByValCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Touch(ByVal values As Collection)
  Call values.Add("callee")
End Sub

Public Sub Run()
  Dim values As Collection
  Set values = New Collection
  Call Touch(values)
  Call values.Add("caller")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a private ByVal object call must preserve the caller's reference: %+v", got)
	}
}

func TestVBA202Issue448TracksModuleInitializationThroughMeCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Main.cls", `Attribute VB_Name = "Main"
Option Explicit
Private target As Worksheet

Private Sub InitializeTarget()
  Set target = ThisWorkbook.Worksheets(1)
End Sub

Private Sub UseTarget()
  Debug.Print target.Name
End Sub

Public Sub Run()
  If target Is Nothing Then Me.InitializeTarget
  Call UseTarget
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("Me initializer should establish the module object before the private helper: %+v", got)
	}
}

func TestVBA202Issue448KeepsByRefEffectsOnTerminalErrorPaths(t *testing.T) {
	t.Parallel()
	for _, test := range []struct {
		name string
		body string
	}{
		{
			name: "resume-next",
			body: `Public Sub Run()
  Dim value As Object
  Set value = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  ClearAndRaise value
  On Error GoTo 0
  Debug.Print value.Exists("key")
End Sub
`,
		},
		{
			name: "handler",
			body: `Public Sub Run()
  Dim value As Object
  Set value = CreateObject("Scripting.Dictionary")
  On Error GoTo Handler
  ClearAndRaise value
  Exit Sub
Handler:
  Debug.Print value.Exists("key")
End Sub
`,
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			dir := t.TempDir()
			writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub ClearAndRaise(ByRef value As Object)
  Set value = Nothing
  Err.Raise 5
End Sub

`+test.body)

			findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
			if err != nil {
				t.Fatal(err)
			}
			if got := findingsByCode(findings, "VBA202"); len(got) != 1 {
				t.Fatalf("ByRef effects before a terminal error must reach %s: %+v", test.name, got)
			}
		})
	}
}

func TestVBA202Issue448RequiresCollectionReturnOnEveryNormalPath(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function MaybeItem(ByVal rows As Collection, ByVal takeIt As Boolean) As Object
  If takeIt Then Set MaybeItem = rows(1)
End Function

Public Sub Run()
  Dim rows As Collection
  Dim item As Object
  Set rows = New Collection
  Set item = MaybeItem(rows, False)
  Debug.Print item.Name
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 1 {
		t.Fatalf("a conditional collection-item return must remain nullable: %+v", got)
	}
}

func TestVBA202Issue448PropagatesObjectFactoryProgIDThroughCallExpression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function NewDictionary() As Object
  Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Private Function ForwardDictionary() As Object
  Set ForwardDictionary = NewDictionary()
End Function

Private Sub AddValue(ByVal dictionary As Object, ByVal item As Object)
  dictionary.Add "key", item
  Debug.Print item.Name
End Sub

Public Sub Run()
  Dim dictionary As Object
  Dim item As Object
  Set dictionary = ForwardDictionary()
  Set item = CreateObject("Scripting.Dictionary")
  AddValue dictionary, item
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a project-local factory call expression must preserve Dictionary Add arguments: %+v", got)
	}
}

func TestVBA202Issue448RejectsFSOAfterNonFactoryReceiverReplacement(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim fs As Object
  Dim textFile As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set fs = New Collection
  Set textFile = fs.OpenTextFile("sample.txt", 1)
  Debug.Print textFile.ReadAll
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 1 {
		t.Fatalf("a non-FSO receiver replacement must not preserve OpenTextFile: %+v", got)
	}
}

func TestVBA202Issue448DoesNotTreatShadowedNonCallableAsTerminal(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Helper(ByRef value As Object)
  Err.Raise 5
End Sub

Public Sub Run()
  Dim Helper As Object
  Dim value As Object
  Set value = CreateObject("Scripting.Dictionary")
  Call Helper(value)
  Debug.Print value.Exists("key")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 1 {
		t.Fatalf("a non-callable shadow must not remove the reachable continuation: %+v", got)
	}
}

func TestVBA202Issue448TracksTerminalCollectionFactoryBranch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub RaiseContractError()
  Err.Raise 5
End Sub

Private Function BuildDynamicArguments(ByVal takeIt As Boolean) As Collection
  Dim result As Collection
  Set result = New Collection
  If takeIt Then
    Set BuildDynamicArguments = result
    Exit Function
  End If
  RaiseContractError
End Function

Public Sub Run()
  Dim values As Collection
  Set values = BuildDynamicArguments(False)
  Debug.Print values.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a terminal invalid branch must not make a Collection factory nullable: %+v", got)
	}
}

func TestVBA202Issue448TracksImplicitCollectionFunctionResults(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function NewValues() As Collection
  Set NewValues = New Collection
End Function

Private Function ForwardValues() As Collection
  Set ForwardValues = NewValues
End Function

Public Sub Run()
  Dim values As Collection
  Set values = ForwardValues
  Debug.Print values.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("bare object-returning function results should preserve Collection state: %+v", got)
	}
}

func TestVBA202Issue448TracksImplicitObjectFunctionResults(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function NewDictionary() As Object
  Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Private Function ForwardDictionary() As Object
  Set ForwardDictionary = NewDictionary
End Function

Public Sub Run()
  Dim dictionary As Object
  Set dictionary = ForwardDictionary
  Debug.Print dictionary.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("bare object-returning function results should preserve Object state: %+v", got)
	}
}

func TestVBA202Issue448TracksImplicitObjectFunctionWithOptionalArgument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function BuildCommand(Optional ByVal dataRow As Object) As Object
  Dim commandObject As Object
  Dim parametersObject As Object
  Set commandObject = CreateObject("Scripting.Dictionary")
  CallByName commandObject, "CompareMode", VbLet, 0
  Set parametersObject = CallByName(commandObject, "Item", VbGet, "parameters")
  Set BuildCommand = commandObject
End Function

Public Sub Run()
  Dim commandObject As Object
  Set commandObject = BuildCommand
  Debug.Print commandObject.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an optional-argument object-returning function should preserve Object state: %+v", got)
	}
}

func TestVBA202Issue448TracksObjectFactoryMemberResults(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function CreateFileSystemObject() As Object
  Dim fileSystem As Object
  Set fileSystem = CreateObject("Scripting.FileSystemObject")
  Set CreateFileSystemObject = fileSystem
End Function

Private Sub WalkFolder(ByVal folderPath As String)
  Dim folder As Object
  Dim item As Object
  Set folder = CreateFileSystemObject().GetFolder(folderPath)
  For Each item In folder.Files
    Debug.Print item.Path
  Next item
End Sub

Public Sub Run(ByVal folderPath As String)
  WalkFolder folderPath
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an object factory member result should preserve Object state: %+v", got)
	}
}

func TestVBA202Issue448TracksSplitObjectFactoryMemberResults(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function CreateFileSystemObject() As Object
  Dim fileSystem As Object
  Set fileSystem = CreateObject("Scripting.FileSystemObject")
  Set CreateFileSystemObject = fileSystem
End Function

Private Sub CollectFolderEntries(ByVal folderObject As Object)
  Dim entry As Object
  For Each entry In folderObject.Files
    Debug.Print entry.Name
  Next entry
End Sub

Private Sub DeleteFolder(ByVal directoryPath As String, ByVal recursive As Boolean)
  Dim folderObject As Object
  Dim fso As Object
  Set fso = CreateFileSystemObject()
  If Not recursive Then
    Set folderObject = fso.GetFolder(directoryPath)
    If folderObject.Files.Count > 0 Or folderObject.SubFolders.Count > 0 Then
      Err.Raise 5
    End If
  End If
  fso.DeleteFolder directoryPath, True
End Sub

Public Sub Run(ByVal directoryPath As String)
  Dim fso As Object
  Set fso = CreateFileSystemObject()
  CollectFolderEntries fso.GetFolder(directoryPath)
  DeleteFolder directoryPath, False
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a helper-returned FileSystemObject member result should preserve Object state: %+v", got)
	}
}

func TestVBA202Issue448RejectsUninitializedCollectionReturnCycle(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function FirstValues() As Collection
  Set FirstValues = SecondValues
End Function

Private Function SecondValues() As Collection
  Set SecondValues = FirstValues
End Function

Public Sub Run()
  Dim values As Collection
  Set values = FirstValues
  Debug.Print values.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 1 {
		t.Fatalf("a collection return cycle without an allocating base must remain nullable: %+v", got)
	}
}

func TestVBA202Issue448TracksTypedCollectionMemberResults(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Factory.cls", `Attribute VB_Name = "Factory"
Option Explicit
Public Function BuildValues() As Collection
  Set BuildValues = New Collection
End Function
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim factory As Factory
  Dim values As Collection
  Set factory = New Factory
  Set values = factory.BuildValues
  values.Add "item"
  Debug.Print values.Count
  Debug.Print values.Item(1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a typed Collection member result should preserve Collection state: %+v", got)
	}
}

func TestVBA202Issue448TracksTypedObjectMemberResults(t *testing.T) {
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
  Dim dictionary As Object
  Set factory = New Factory
  Set dictionary = factory.BuildDictionary
  Debug.Print dictionary.Count
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a typed Object member result should preserve object state: %+v", got)
	}
}

func TestVBA202Issue448PropagatesInitializedCollectionIntoFriendHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Worker.cls", `Attribute VB_Name = "Worker"
Option Explicit
Friend Sub Consume(ByVal values As Collection)
  Debug.Print values.Count
End Sub

Public Sub Run()
  Dim values As Collection
  Set values = New Collection
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an initialized Collection should reach a Friend helper: %+v", got)
	}
}

func TestVBA202Issue448PropagatesCollectionThroughExpressionHelpers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Query.cls", `Attribute VB_Name = "Query"
Option Explicit

Private mBodies As Collection
Private mRole As Long

Private Sub RequireQueryable(ByVal memberName As String)
  If mRole <> 1 Then
    RaiseContractError memberName
  End If
End Sub

Private Sub RaiseContractError(ByVal memberName As String)
  Err.Raise 5, "Query", memberName
End Sub

Private Function RenderNode(ByVal node As Query, ByVal values As Collection) As String
  Select Case 1
    Case 1
      RenderNode = RenderValue(node, values)
    Case 2
      RenderNode = RenderOperator(node, values)
  End Select
End Function

Private Function RenderOperator(ByVal node As Query, ByVal values As Collection) As String
  RenderOperator = RenderNode(node, values)
End Function

Private Function RenderValue(ByVal node As Query, ByVal values As Collection) As String
  If IsNullValueNode(node) Then Exit Function
  values.Add "value"
  RenderValue = "?"
End Function

Private Function IsNullValueNode(ByVal node As Query) As Boolean
  If node Is Nothing Then Exit Function
  IsNullValueNode = False
End Function

Private Function BuildSql() As String
  Dim body As Query
  Dim values As Collection
  Set values = New Collection
  For Each body In mBodies
    BuildSql = BuildSql & RenderNode(body, values)
  Next body
End Function

Public Function ToSqlString() As String
  RequireQueryable "ToSqlString"
  ToSqlString = BuildSql()
End Function

Public Function RunSql(ByVal node As Query) As String
  mRole = 1
  Set mBodies = New Collection
  mBodies.Add node
  RunSql = ToSqlString()
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an initialized Collection should reach expression helper parameters: %+v", got)
	}
}

func TestVBA202Issue448KeepsNullablePublicCollectionThroughPrivateByValHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Consume(ByVal values As Collection)
  values.Add "item"
End Sub

Public Sub Run(ByVal values As Collection)
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA202")
	if len(got) != 1 || got[0].Procedure != "Consume" {
		t.Fatalf("a nullable public Collection must remain nullable through a private ByVal helper: %+v", got)
	}
}

func TestVBA202Issue448PropagatesCollectionReturnThroughValidationHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function Normalize(ByVal candidate As Variant) As Collection
  Dim result As Collection
  Set result = New Collection
  If Not IsObject(candidate) Then RaiseContractError
  Set Normalize = result
End Function

Private Sub Configure(ByVal values As Collection)
  Dim item As Object
  Set item = values.Item(1)
End Sub

Private Sub RaiseContractError()
  Err.Raise 5
End Sub

Public Sub Run(ByVal candidate As Variant)
  Dim values As Collection
  Set values = Normalize(candidate)
  Configure values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a validated Collection return should reach the ByVal helper: %+v", got)
	}
}

func TestVBA202Issue448PropagatesCollectionReturnIntoClassFriendHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Factory.cls", `Attribute VB_Name = "Factory"
Option Explicit

Private Function Normalize(ByVal candidate As Variant) As Collection
  Dim result As Collection
  Set result = New Collection
  If Not IsObject(candidate) Then RaiseContractError
  Set Normalize = result
End Function

Friend Sub Configure(ByVal values As Collection)
  Dim item As Object
  Set item = values.Item(1)
End Sub

Private Sub RaiseContractError()
  Err.Raise 5
End Sub

Public Sub Run(ByVal candidate As Variant)
  Dim values As Collection
  Dim target As Factory
  Set values = Normalize(candidate)
  Set target = New Factory
  target.Configure values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a validated Collection return should reach a class Friend helper: %+v", got)
	}
}

func TestVBA202Issue448PreservesCollectionAfterLoopForReceiverCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Worker.cls", `Attribute VB_Name = "Worker"
Option Explicit

Friend Sub Consume(ByVal values As Collection)
  Debug.Print values.Count
End Sub

Public Sub Run(ParamArray arguments() As Variant)
  Dim raw As Variant
  Dim target As Worker
  Dim values As Collection
  Set values = New Collection
  For Each raw In arguments
    values.Add raw
  Next raw
  Set target = New Worker
  target.Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("a Collection assigned before a loop should reach a receiver helper: %+v", got)
	}
}

func TestVBA202Issue448DoesNotUseLaterConstructorAssignment(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Consume(ByVal values As Collection)
  Dim item As Object
  Set item = values.Item(1)
End Sub

Public Sub Run()
  Dim values As Collection
  Consume values
  Set values = New Collection
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 1 {
		t.Fatalf("a constructor assigned after the call must not establish entry state: %+v", got)
	}
}
