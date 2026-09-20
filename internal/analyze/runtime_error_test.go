package analyze

import (
	"encoding/json"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
)

func TestRuntimeConstantEnvironmentUsesImmutableOverlay(t *testing.T) {
	base := constexpr.NewValues(map[string]constexpr.Value{
		"ModuleLimit": {Kind: constexpr.ValueLong, Integer: 7},
		"SharedText":  {Kind: constexpr.ValueString, String: "base"},
	})
	state := runtimeConstantState{
		"modulelimit": {Kind: constexpr.ValueLong, Integer: 11},
		"localvalue":  {Kind: constexpr.ValueLong, Integer: 13},
	}
	env := runtimeConstantEnvironment(base, state)
	if got, ok := env.Resolve("MODULELIMIT"); !ok || got.Integer != 11 {
		t.Fatalf("overlay state = %#v, want local shadow value", got)
	}
	if got, ok := env.Resolve("sharedtext"); !ok || got.String != "base" {
		t.Fatalf("overlay base fallback = %#v, want base value", got)
	}
	if got, ok := env.Resolve("LocalValue"); !ok || got.Integer != 13 {
		t.Fatalf("overlay local value = %#v, want state value", got)
	}
	if _, ok := env.Resolve("missing"); ok {
		t.Fatal("overlay resolved an unknown constant")
	}
}

func TestRuntimeConstantScopeHidesProcedureLocalShadows(t *testing.T) {
	base := constexpr.NewValues(map[string]constexpr.Value{
		"ProjectZero": {Kind: constexpr.ValueLong, Integer: 0},
		"SharedText":  {Kind: constexpr.ValueString, String: "base"},
	})
	scope := runtimeConstantScope{
		base:   base,
		hidden: map[string]bool{"projectzero": true},
	}

	if _, ok := scope.Resolve("PROJECTZERO"); ok {
		t.Fatal("procedure-local declaration did not hide the project constant")
	}
	if got, ok := scope.Resolve("SharedText"); !ok || got.String != "base" {
		t.Fatalf("unshadowed project constant = %#v, %v; want base value", got, ok)
	}
}

func TestVBA249DetectsLiteralAndProjectConstantRuntimeFailures(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Const ZeroValue As Long = 0
Public Const BadText As String = "not numeric"

Public Sub Run()
  Dim result As Double
  result = 10 / 0
  result = 10 / ZeroValue
  result = 10 \ 0
  result = 10 Mod ZeroValue
  result = BadText + 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 5 {
		t.Fatalf("VBA249 findings = %+v, want deterministic division and numeric failures", got)
	}
	if got[0].Severity != "error" || !strings.Contains(got[0].Message, "guaranteed") {
		t.Fatalf("unexpected VBA249 finding = %+v", got[0])
	}
	if got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "division_by_zero" {
		t.Fatalf("unexpected runtime error context = %+v", got[0].RuntimeError)
	}
	encoded, err := json.Marshal(got[0])
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"runtime_error":{"kind":"division_by_zero"}`) {
		t.Fatalf("runtime error context missing from JSON: %s", encoded)
	}
}

func TestVBA249LeavesUnknownAndNumericStringOperandsSilent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim result As Double
  Dim denominator As Double
  result = 10 / denominator
  result = "123" + 1
  result = "1,2" + 1
  result = MissingText + 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("unknown or numeric-string operands should remain silent: %+v", got)
	}
}

func TestVBA249CanBeDisabled(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Public Sub Run()
  Debug.Print 10 / 0
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDeterministicRuntimeErrors = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("disabled VBA249 produced findings: %+v", got)
	}
}

func TestVBA249PropagatesProcedureConstantsAcrossAllBranches(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
  Dim result As Double
  Dim denominator As Double
  If chooseFirst Then
    denominator = 0
  Else
    denominator = 0
  End If
  result = 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 1 {
		t.Fatalf("zero denominator on every branch should be reported once: %+v", got)
	}
}

func TestVBA249LeavesConflictingProcedureConstantsSilent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
  Dim result As Double
  Dim denominator As Double
  If chooseFirst Then
    denominator = 0
  Else
    denominator = 1
  End If
  result = 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("conflicting branch values should remain silent: %+v", got)
	}
}

func TestVBA249ResolvesProcedureConstAndConstantBranchReachability(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Const ZeroValue As Long = 0
  Dim denominator As Long
  If False Then
    denominator = 1
  Else
    denominator = ZeroValue
  End If
  Debug.Print 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 1 {
		t.Fatalf("constant unreachable branch and local Const should prove zero divisor: %+v", got)
	}
}

func TestVBA249DoesNotReportReDimBeforeIndexedUseInSelectCase(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function AlignmentCenters(ByVal version As Long) As Variant
  Dim centers() As Long
  Select Case version
    Case 1
      AlignmentCenters = Empty
    Case 2
      ReDim centers(0 To 1)
      centers(0) = 6
      centers(1) = 18
      AlignmentCenters = centers
	    Case 3
	      ReDim centers(0 To 1)
	      centers(0) = 6
	      centers(1) = 22
	      AlignmentCenters = centers
	    Case 4
	      ReDim centers(0 To 1)
	      centers(0) = 6
	      centers(1) = 26
	      AlignmentCenters = centers
	    Case Else
      Err.Raise vbObjectError + 1004
  End Select
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("indexed uses after branch-local ReDim are allocated: %+v", got)
	}
}

func TestVBA249RetainsSelectCaseUseBeforeReDim(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ReadValue(ByVal mode As Long) As Long
  Dim values() As Long
  Select Case mode
    Case 1
      ReadValue = values(0)
    Case 2
      ReDim values(0 To 1)
      values(0) = 1
      ReadValue = values(0)
  End Select
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("unallocated Select Case access must remain deterministic: %+v", got)
	}
}

func TestVBA249RetainsSelectCaseHeaderUseBeforeBranchReDim(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Long
  Select Case values(0)
    Case 1
      ReDim values(0 To 1)
      values(0) = 1
  End Select
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("Select Case header access must remain deterministic: %+v", got)
	}
}

func TestVBA249DoesNotTreatReDimPreserveAsAllocationOfUnallocatedArray(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Long
  On Error Resume Next
  Select Case 1
    Case 1
      ReDim Preserve values(0 To 1)
      values(0) = 1
  End Select
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("ReDim Preserve on an unallocated array must not establish allocation: %+v", got)
	}
}

func TestVBA249DetectsConditionalReDimPreserveFailure(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal shouldResize As Boolean)
  Dim values() As Long
  If shouldResize Then
    ReDim Preserve values(0 To 1)
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("conditional ReDim Preserve must retain its reachable unallocated-array failure: %+v", got)
	}
}

func TestVBA249ReportsResumeNextContinuationAfterDeterministicPreserveFailure(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private failed() As String

Private Function UpdateLinks(strType As String) As Boolean
  On Error GoTo ErrTrap
  Select Case strType
    Case "Table"
      ReDim Preserve failed(1)
      failed(1) = "x"
  End Select
  UpdateLinks = True
  Exit Function
ErrTrap:
  If Err.Number = 52 Then
    Resume Next
  End If
  Resume Next
End Function

Public Sub Run()
  UpdateLinks "Table"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	lines := map[int]bool{}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Procedure != "UpdateLinks" || finding.RuntimeError == nil || finding.RuntimeError.Kind != "array_unallocated" {
			continue
		}
		lines[finding.Line] = true
	}
	if !lines[8] || !lines[9] {
		t.Fatalf("deterministic ReDim Preserve failure must cover the failed expression and Resume Next continuation: %+v", findingsByCode(findings, "VBA249"))
	}
}

func TestVBA249DoesNotReportPreserveAfterItsBoundQuerySucceeds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Expand(ByRef values() As Long)
  Dim newUpper As Long
  newUpper = UBound(values) + 1
  ReDim Preserve values(0 To UBound(values) + 1)
End Sub

Public Sub Run()
  Dim values() As Long
  Expand values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Procedure == "Expand" && finding.Line == 5 {
			t.Fatalf("the Preserve statement is unreachable after a failed same-array UBound query: %+v", finding)
		}
	}
}

func TestVBA249DoesNotReportUseAfterFailingPreserve(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Expand(ByRef values() As Long)
  ReDim Preserve values(0 To UBound(values) + 1)
  values(0) = 1
End Sub

Public Sub Run()
  Dim values() As Long
  Expand values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Procedure == "Expand" && finding.Line == 4 {
			t.Fatalf("an execution-failing Preserve statement must not make its following use reachable: %+v", finding)
		}
	}
}

func TestVBA249StopsAfterFailingPreserveWithTerminatingErrorHandler(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Long
  On Error GoTo Handler
  ReDim Preserve values(0 To 1)
  values(0) = 1
  Exit Sub
Handler:
  Exit Sub
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	preserveFound := false
	for _, finding := range findingsByCode(findings, "VBA249") {
		switch finding.Line {
		case 5:
			preserveFound = true
		case 6:
			t.Fatalf("a terminating error handler must not make the post-Preserve use reachable: %+v", finding)
		}
	}
	if !preserveFound {
		t.Fatalf("the unallocated Preserve must remain a true positive: %+v", findingsByCode(findings, "VBA249"))
	}
}

func TestVBA249StopsAfterFailingPreserveInNestedBlockWithTerminatingErrorHandler(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ParseValue(ByVal shouldParse As Boolean) As Boolean
  On Error GoTo ErrorHandler
  Dim parents() As Long
  Dim depth As Long
  Dim ubParents As Long
  If shouldParse Then
    depth = depth + 1
    If depth > ubParents Then
      ReDim Preserve parents(0 To depth)
      ubParents = depth
    End If
    parents(depth) = 1
  End If
  ParseValue = True
  Exit Function
ErrorHandler:
  ParseValue = False
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	preserveFound := false
	for _, finding := range findingsByCode(findings, "VBA249") {
		switch finding.Line {
		case 10:
			preserveFound = true
		case 13:
			t.Fatalf("a terminating handler must not make the nested post-Preserve use reachable: %+v", finding)
		}
	}
	if !preserveFound {
		t.Fatalf("the nested unallocated Preserve must remain a true positive: %+v", findingsByCode(findings, "VBA249"))
	}
}

func TestVBA249StopsAfterFailingPreserveInsideLoopWithTerminatingErrorHandler(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ParseValue() As Boolean
  On Error GoTo ErrorHandler
  Dim parents() As Long
  Dim depth As Long
  Dim ubParents As Long
  Dim i As Long
  Do While i <= 0
    depth = depth + 1
    If depth > ubParents Then
      ReDim Preserve parents(0 To depth)
      ubParents = depth
    End If
    parents(depth) = 1
    i = i + 1
  Loop
  ParseValue = True
  Exit Function
ErrorHandler:
  ParseValue = False
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	preserveFound := false
	for _, finding := range findingsByCode(findings, "VBA249") {
		switch finding.Line {
		case 11:
			preserveFound = true
		case 14:
			t.Fatalf("a terminating handler must stop the loop-body normal continuation after Preserve: %+v", finding)
		}
	}
	if !preserveFound {
		t.Fatalf("the loop-body unallocated Preserve must remain a true positive: %+v", findingsByCode(findings, "VBA249"))
	}
}

func TestVBA249StopsAfterFailingPreserveBeforeAnotherArrayOperation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim first() As Long
  Dim second() As Long
  ReDim Preserve first(0 To 1)
  ReDim Preserve second(0 To 1)
  second(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	var firstPreserveFound bool
	for _, finding := range got {
		switch finding.Line {
		case 5:
			firstPreserveFound = true
		case 6, 7:
			t.Fatalf("operation after a fatal Preserve must be unreachable: %+v", finding)
		}
	}
	if !firstPreserveFound {
		t.Fatalf("expected the first unallocated Preserve diagnostic: %+v", got)
	}
}

func TestVBA249StopsAfterFailingIndexedUseBeforeAnotherArrayOperation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal shouldRun As Boolean)
  Dim values() As Long
  If shouldRun Then
    values(0) = 1
    values(1) = 2
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	var firstUseFound bool
	for _, finding := range got {
		switch finding.Line {
		case 5:
			firstUseFound = true
		case 6:
			t.Fatalf("operation after a fatal indexed access must be unreachable: %+v", finding)
		}
	}
	if !firstUseFound {
		t.Fatalf("expected the first unallocated indexed access diagnostic: %+v", got)
	}
}

func TestVBA249HonorsErrorHandlerResetBeforePreserve(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Long
  Dim i As Long
  On Error Resume Next: i = UBound(values): On Error GoTo 0
  ReDim Preserve values(0 To 1)
  values(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Procedure == "Run" && finding.Line == 7 {
			t.Fatalf("the indexed use is unreachable after Preserve with a reset handler: %+v", finding)
		}
	}
	var preserveFound bool
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Procedure == "Run" && finding.Line == 6 {
			preserveFound = true
			break
		}
	}
	if !preserveFound {
		t.Fatalf("expected the unallocated Preserve diagnostic after On Error GoTo 0: %+v", findingsByCode(findings, "VBA249"))
	}
}

func TestVBA249RespectsProcedureLocalShadowingOfEnumMember(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Enum RegistryType
  iItem = 2
End Enum

Public Sub Run()
  Dim values() As String
  Dim iItem As Long
  iItem = 0
  ReDim values(0 To 0)
  values(iItem) = "ok"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a procedure-local variable must shadow the same-named Enum member: %+v", got)
	}
}

func TestVBA249PropagatesSafeUBoundIntoForBody(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function SafeUBoundValues(ByRef values() As Long) As Long
  On Error GoTo noValues
  SafeUBoundValues = UBound(values)
  Exit Function
noValues:
  SafeUBoundValues = -1
End Function

Public Sub Run(ByRef values() As Long)
  Dim ub As Long
  ub = SafeUBoundValues(values)
  Dim i As Long
  For i = 0 To ub
    values(i) = 1
  Next
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a safe-bound loop body must not retain a deterministic unallocated-array finding: %+v", got)
	}
}

func TestVBA249DoesNotReportDirectBoundsLoopBodyAfterSuccessfulBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByRef values() As Long)
  Dim i As Long
  For i = LBound(values) To UBound(values)
    values(i) = 1
  Next
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Line == 5 {
			t.Fatalf("a direct bounds loop body must not retain an unallocated-array finding: %+v", findingsByCode(findings, "VBA249"))
		}
	}
}

func TestVBA249DoesNotReportAfterFailingPreserveInGuaranteedLoop(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal childHandle As Long)
  Dim values() As Long
  Dim i As Long
  If childHandle <> 0 Then
    Do While (childHandle <> 0)
      ReDim Preserve values(0 To 0)
      childHandle = 0
    Loop
    For i = LBound(values) To UBound(values)
      values(i) = 1
    Next
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	preserveFound := false
	for _, finding := range findingsByCode(findings, "VBA249") {
		switch finding.Line {
		case 7:
			preserveFound = true
		case 10, 11:
			t.Fatalf("a bounds loop after a guaranteed failing Preserve must be unreachable: %+v", findingsByCode(findings, "VBA249"))
		}
	}
	if !preserveFound {
		t.Fatalf("the failing Preserve must remain a true positive: %+v", findingsByCode(findings, "VBA249"))
	}
}

func TestVBA249DoesNotReportZeroCountCleanupLoopForUnallocatedArray(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal shouldParse As Boolean)
  Dim values() As Long
  Dim valueCount As Long: valueCount = 0
  If shouldParse Then
    valueCount = valueCount + 1
    ReDim Preserve values(1 To valueCount)
  End If
  Dim startCount As Long: startCount = valueCount
  Dim i As Long
  For i = 1 To startCount
    Debug.Print values(i)
  Next
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	preserveFound := false
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Line == 7 {
			preserveFound = true
		}
		if finding.Line == 12 {
			t.Fatalf("a zero-count cleanup loop must not retain an unallocated-array finding: %+v", findingsByCode(findings, "VBA249"))
		}
	}
	if !preserveFound {
		t.Fatalf("the conditional first Preserve must remain a true positive: %+v", findingsByCode(findings, "VBA249"))
	}
}

func TestVBA249RetainsIndexedUseAfterInvalidBranchReDim(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Long
  On Error Resume Next
  Select Case 1
    Case 1
      ReDim values(1 To 0)
      values(0) = 1
  End Select
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("invalid ReDim bounds must not establish allocation: %+v", got)
	}
}

func TestVBA249RetainsOuterUseAfterNestedSelectCaseMayNotMatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal mode As Long)
  Dim values() As Long
  Select Case mode
    Case 1
      Select Case mode
        Case 2
          ReDim values(0 To 1)
          values(0) = 1
      End Select
      values(0) = 1
  End Select
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("outer use after a non-matching nested Select Case must remain deterministic: %+v", got)
	}
}

func TestVBA249RetainsUseAfterByRefArrayCallMayErase(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub ClearArray(ByRef values() As Long)
  Erase values
End Sub

Public Sub Run()
  Dim values() As Long
  Select Case 1
    Case 1
      ReDim values(0 To 1)
      Call ClearArray(values)
      values(0) = 1
  End Select
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("a ByRef array call may erase allocation before the indexed use: %+v", got)
	}
}

func TestVBA249RetainsBlockConditionalByRefOutputState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function Receive(ByRef output() As Byte, ByVal hasData As Boolean) As Boolean
  Dim queued() As Byte
  If Not hasData Then Exit Function
  output = queued
  Receive = True
End Function

Public Sub Run()
  Dim output() As Byte
  If Receive(output, True) Then
    Debug.Print UBound(output)
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) == 0 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("a block conditional ByRef output must retain its possible unallocated state: %+v", got)
	}
}

func TestVBA227RetainsPossibleUnallocatedConditionalByRefMemberOutput(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Type BinaryMessage
  data() As Byte
End Type

Private queue() As BinaryMessage
Private count As Long

Private Function Receive(ByRef output() As Byte) As Boolean
  If count > 0 Then
    output = queue(0).data
    count = count - 1
    Receive = True
  End If
End Function

Public Sub Run()
  Dim output() As Byte
  ReDim queue(0 To 1)
  count = 1
  If Receive(output) Then
    Debug.Print UBound(output)
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Procedure == "Run" {
			t.Fatalf("a conditional member ByRef output must not be promoted to deterministic VBA249: %+v", finding)
		}
	}
	for _, finding := range findingsByCode(findings, "VBA227") {
		if finding.Procedure == "Run" && finding.Line == 22 && finding.arrayOperationKey == "bound:ubound:output:unallocated" {
			return
		}
	}
	t.Fatalf("a conditional member ByRef output must retain a possible unallocated-array warning: %+v", findingsByCode(findings, "VBA227"))
}

func TestVBA249DoesNotReportUnknownByRefArrayOutput(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function Fill(ByRef output() As Byte) As Boolean
  ReDim Preserve output(0 To UBound(output) + 1)
  Fill = True
End Function

Public Sub Run()
  Dim output() As Byte
  Call Fill(output)
  Debug.Print UBound(output)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Procedure == "Run" {
			t.Fatalf("a caller must not inherit a deterministic array failure past a callee failure boundary: %+v", finding)
		}
	}
}

func TestVBA249RetainsPossibleUnallocatedArrayFieldByRefOutput(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Type BinaryMessage
  data() As Byte
End Type

Private queue() As BinaryMessage
Private count As Long

Private Sub Initialize()
  ReDim queue(0 To 1)
End Sub

Private Sub Enqueue(ByRef payload() As Byte)
  queue(0).data = payload
  count = 1
End Sub

Private Function Receive(ByRef output() As Byte) As Boolean
  If count > 0 Then
    output = queue(0).data
    Erase queue(0).data
    count = count - 1
    Receive = True
  End If
End Function

Public Sub Run()
  Dim payload() As Byte
  Initialize
  Enqueue payload
  If Receive(payload) Then
    Debug.Print UBound(payload)
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	for _, finding := range got {
		if finding.Procedure == "Run" && finding.RuntimeError != nil && finding.RuntimeError.Kind == "array_unallocated" {
			return
		}
	}
	t.Fatalf("a dynamic array field may carry an unallocated array through a successful ByRef output: %+v", got)
}

func TestVBA227RetainsPossibleUnallocatedQueuedPayloadByRefOutput(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Type BinaryMessage
  data() As Byte
End Type

Private queue() As BinaryMessage
Private count As Long

Private Sub Initialize()
  ReDim queue(0 To 1)
End Sub

Private Sub Process(ByRef payload() As Byte, ByVal payloadLen As Long)
  Dim binaryData() As Byte
  If payloadLen > 0 Then
    binaryData = payload
  Else
    Erase payload
    binaryData = payload
  End If
  queue(0).data = binaryData
  count = count + 1
End Sub

Private Function Receive(ByRef output() As Byte) As Boolean
  If count > 0 Then
    output = queue(0).data
    Erase queue(0).data
    count = count - 1
    Receive = True
  End If
End Function

Public Sub Run()
  Dim payload() As Byte
  Initialize
  Do
    Process payload, 0
    If Receive(payload) Then
      Debug.Print UBound(payload)
    End If
    Exit Do
  Loop
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a possible unallocated queued payload must not be promoted to deterministic VBA249: %+v", got)
	}
	got := findingsByCode(findings, "VBA227")
	for _, finding := range got {
		if finding.Procedure == "Run" && finding.Line == 40 && finding.arrayOperationKey == "bound:ubound:payload:unallocated" {
			return
		}
	}
	t.Fatalf("a zero-length queued payload must retain a possible unallocated-array warning at its UBound use: %+v", got)
}

func TestVBA249DoesNotReportBranchLocalRedimOrUnknownArrayAssignments(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal mode As Long, ByVal input As Variant)
  Dim buffer() As Byte
  Dim unknownBounds() As Byte
  Dim unknownIndex() As Byte
  Select Case mode
    Case 1
      Dim localBuffer() As Byte: ReDim localBuffer(0 To 1)
      Dim i As Long
      For i = 0 To UBound(localBuffer)
        localBuffer(i) = 1
      Next i
    Case 2
      unknownBounds = input
      Debug.Print UBound(unknownBounds)
    Case 3
      unknownIndex = input
      Debug.Print unknownIndex(0)
  End Select
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("branch-local ReDim and unknown Variant assignments must remain silent for deterministic VBA249: %+v", got)
	}
}

func TestVBA249DoesNotReportSplitAssignedArrayInsideWithBlock(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function ReadParts(ByVal source As String) As Long
  Dim parts() As String
  With Application
    parts = Split(source, ",")
    ReadParts = UBound(parts)
  End With
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a Split-assigned array must remain allocated inside a With block: %+v", got)
	}
}

func TestVBA249ReportsPrecisePreserveInsideWithBlock(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub FixLostState(ByVal hasTerminated As Boolean)
  Dim terminated() As Long
  Dim terminatedCount As Long
  Dim terminatedUB As Long
  With Application
    If hasTerminated Then
      terminatedCount = terminatedCount + 1
      If terminatedCount >= terminatedUB Then
        terminatedUB = terminatedCount * 2
        ReDim Preserve terminated(0 To terminatedUB)
      End If
      terminated(terminatedCount) = 1
    End If
  End With
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	var preserveFound bool
	for _, finding := range got {
		switch finding.Line {
		case 6:
			t.Fatalf("a With block entry must not receive the inner array failure: %+v", finding)
		case 11:
			preserveFound = true
		}
	}
	if !preserveFound {
		t.Fatalf("the first Preserve inside the reachable branch must remain diagnosed precisely: %+v", got)
	}
}

func TestVBA249StopsAfterFailingPreserveInsideNestedLoopAndWith(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub FixLostState(ByVal ptr As Long)
  Dim terminated() As Long
  Dim terminatedCount As Long
  Dim terminatedUB As Long
  With Application
    Do While ptr <> 0
      If ptr > 0 Then
        terminatedCount = terminatedCount + 1
        If terminatedCount >= terminatedUB Then
          terminatedUB = terminatedCount * 2
          ReDim Preserve terminated(0 To terminatedUB)
        End If
        terminated(terminatedCount) = ptr
      End If
      ptr = ptr - 1
    Loop
    If terminatedCount = 0 Then Exit Sub
    Debug.Print terminated(terminatedCount)
  End With
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA249") {
		if finding.Line != 12 {
			t.Fatalf("a fatal Preserve inside nested Do/With must stop later normal-flow diagnostics: %+v", finding)
		}
	}
}

func TestAnalyzerVBA227PropagatesSuccessfulVariantArrayByRefAssignment(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Fill(ByRef output() As Variant, ByVal source As Variant)
  output = source
End Sub

Public Sub Run()
  Dim output() As Variant
  Dim source As Variant
  source = Array(1, 2)
  Fill output, source
  Debug.Print output(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a successful Variant-to-ByRef-array assignment must establish the output allocation: %+v", got)
	}
}

func TestVBA249PropagatesPrivateModuleAllocationAfterCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long

Private Sub SetupValues()
  ReDim values(0 To 1)
End Sub

Public Sub Run()
  SetupValues
  values(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a private module allocation before indexed use must be visible to VBA249: %+v", got)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a private module allocation before indexed use must not retain a lifecycle warning: %+v", got)
	}
}

func TestVBA249PropagatesModuleAllocationIntoRegisteredCallback(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long
Private valueCount As Long

Private Sub InitValues()
    If valueCount > 0 Then Exit Sub
    ReDim values(0 To 1)
    valueCount = 2
End Sub

Private Function ValidIndex(ByVal index As Long) As Boolean
    If index < 0 Then Exit Function
    InitValues
    ValidIndex = True
End Function

Private Sub InitCallbackThunk()
    Dim opcodes() As Byte
    ReDim opcodes(0 To 1)
    Dim callbackAddress As Long
    callbackAddress = GetAddressOf(AddressOf Callback)
End Sub

Private Sub InitCallbackWindow()
    InitCallbackThunk
End Sub

Public Function Callback(ByVal value As Long) As Long
    Callback = values(value)
End Function

Public Sub Run()
    If Not ValidIndex(0) Then Exit Sub
    InitCallbackWindow
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a callback registered after validated module initialization must retain the allocation proof: %+v", got)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a callback registered after validated module initialization must not retain a lifecycle warning: %+v", got)
	}
}

func TestVBA249PropagatesAllocationThroughSuccessfulModuleFunctionGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long

Private Sub SetupValues()
    ReDim values(0 To 1)
End Sub

Private Function ValidIndex(ByVal index As Long) As Boolean
    If index < 0 Then Exit Function
    SetupValues
    ValidIndex = True
End Function

Public Sub Run()
    If Not ValidIndex(0) Then Exit Sub
    values(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a successful module function guard must retain the setup allocation: %+v", got)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a successful module function guard must not retain a lifecycle warning: %+v", got)
	}
}

func TestVBA249PropagatesCountGuardedModuleAllocation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long
Private valueCount As Long

Private Sub InitValues()
    If valueCount > 0 Then Exit Sub
    ReDim values(0 To 1)
    valueCount = 2
End Sub

Public Sub Run()
    InitValues
    values(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a positive module count after ReDim must preserve the allocation proof: %+v", got)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a positive module count after ReDim must not retain a lifecycle warning: %+v", got)
	}
}

func TestVBA249RecognizesObjectSetupGuardForModuleArrays(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Renderer.bas", `Option Explicit
Private mForm As Object
Private mControls() As Object
Private mReady As Boolean

Public Sub AttachForm(ByVal formInstance As Object)
  Set mForm = formInstance
  ReDim mControls(1 To 2)
End Sub

Public Sub BuildScene()
  If mForm Is Nothing Then
    Exit Sub
  End If
  Set mControls(1) = CreateObject("Scripting.Dictionary")
  mReady = True
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("an object setup guard should prove the module array allocation: %+v", got)
	}
}

func TestVBA249RejectsObjectSetupUseBeforeGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Renderer.bas", `Option Explicit
Private mForm As Object
Private mControls() As Object

Public Sub AttachForm(ByVal formInstance As Object)
  Set mForm = formInstance
  ReDim mControls(1 To 2)
End Sub

Public Sub BuildScene()
  Set mControls(1) = CreateObject("Scripting.Dictionary")
  If mForm Is Nothing Then
    Exit Sub
  End If
  Set mControls(1) = CreateObject("Scripting.Dictionary")
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	seenBeforeGuard := false
	seenAfterGuard := false
	for _, finding := range got {
		switch finding.Line {
		case 11:
			seenBeforeGuard = true
		case 15:
			seenAfterGuard = true
		}
	}
	if !seenBeforeGuard || seenAfterGuard {
		t.Fatalf("only the indexed use before the object guard should remain unsafe: %+v", got)
	}
}

func TestVBA249RejectsUnsafeObjectSetupArrayProof(t *testing.T) {
	t.Parallel()
	cases := []struct {
		name  string
		setup string
		extra string
	}{
		{
			name: "conditional redim",
			setup: `Public Sub AttachForm(ByVal formInstance As Object, ByVal allocate As Boolean)
  Set mForm = formInstance
  If allocate Then
    ReDim mControls(1 To 2)
  End If
End Sub`,
		},
		{
			name: "erase after redim",
			setup: `Public Sub AttachForm(ByVal formInstance As Object)
  Set mForm = formInstance
  ReDim mControls(1 To 2)
  Erase mControls
End Sub`,
		},
		{
			name: "erase in separate reset procedure",
			setup: `Public Sub AttachForm(ByVal formInstance As Object)
  Set mForm = formInstance
  ReDim mControls(1 To 2)
End Sub`,
			extra: `Public Sub ResetControls()
  Erase mControls
End Sub`,
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			writeModule(t, dir, "Renderer.bas", "Option Explicit\n"+
				"Private mForm As Object\n"+
				"Private mControls() As Object\n\n"+
				tc.setup+"\n\n"+
				tc.extra+"\n\n"+
				`Public Sub BuildScene()
  If mForm Is Nothing Then
    Exit Sub
  End If
  Set mControls(1) = CreateObject("Scripting.Dictionary")
End Sub
`)

			findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
			if err != nil {
				t.Fatal(err)
			}
			if got := findingsByCode(findings, "VBA249"); len(got) == 0 {
				t.Fatalf("unsafe object setup must not prove persistent allocation: %+v", findings)
			}
		})
	}
}

func TestVBA249TracksArrayStateInsideWithBlock(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function ReadBuffers() As Long
    Dim sendBuf() As Byte
    Dim recvBuf() As Byte
    With Application
        sendBuf = StrConv("CONNECT", vbFromUnicode)
        Debug.Print sendBuf(0)
        Debug.Print UBound(sendBuf)
        ReDim recvBuf(0 To 1)
        Debug.Print recvBuf(0)
        Debug.Print UBound(recvBuf)
    End With
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("array state inside a With block must follow earlier StrConv/ReDim operations: %+v", got)
	}
}

func TestVBA249TracksArrayAllocationInsideNestedIfBlock(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
    Dim data() As Byte
    Dim localBytes() As Byte
    Dim i As Long
    If True Then
        For i = LBound(data) To UBound(data)
            data(i) = 1
            ReDim localBytes(0 To 1)
            localBytes(0) = data(i)
        Next i
    End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	seen := map[string]bool{}
	for _, finding := range got {
		seen[finding.arrayOperationKey] = true
		if finding.arrayOperationKey == "index:localbytes:unallocated" {
			t.Fatalf("a local array indexed after a dominating ReDim must not be reported: %+v", got)
		}
	}
	for _, operation := range []string{
		"bound:lbound:data:unallocated",
		"bound:ubound:data:unallocated",
		"index:data:unallocated",
	} {
		if !seen[operation] {
			t.Fatalf("missing deterministic unallocated data evidence %q: %+v", operation, got)
		}
	}
}

func TestAnalyzerVBA227PropagatesClassAllocationThroughRecursiveParser(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "JSON.cls", `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Option Explicit
Private m_Chars() As Integer
Private m_Tokens() As Long
Private m_Length As Long
Private m_Index As Long

Private Sub LoadText(ByRef text As String)
    Erase m_Chars
    m_Length = Len(text)
    m_Index = 0
    If m_Length = 0 Then Exit Sub
    ReDim m_Chars(0 To m_Length - 1)
    ReDim m_Tokens(0 To 1)
    m_Index = 0
    ParseValue 0, 0, 0
End Sub

Private Function ParseValue(ByVal parentID As Long, ByVal keyStart As Long, ByVal keyLen As Long) As Long
    If m_Index >= m_Length Then Exit Function
    ParseValue = AddToken()
    Select Case m_Chars(m_Index)
        Case 123
            ParseObject parentID
        Case 91
            ParseArray parentID
    End Select
End Function

Private Function AddToken() As Long
    AddToken = 1
    ReDim Preserve m_Tokens(0 To 1)
End Function

Private Sub ParseObject(ByVal parentID As Long)
    m_Index = m_Index + 1
    Do While m_Index < m_Length
        If m_Index >= m_Length Then Exit Do
        If m_Chars(m_Index) = 125 Then Exit Sub
        ParseValue parentID, 0, 0
        If m_Index >= m_Length Then Exit Do
        If m_Chars(m_Index) = 44 Then m_Index = m_Index + 1
    Loop
End Sub

Private Sub ParseArray(ByVal parentID As Long)
    m_Index = m_Index + 1
    Do While m_Index < m_Length
        If m_Index >= m_Length Then Exit Do
        If m_Chars(m_Index) = 93 Then Exit Sub
        ParseValue parentID, 0, 0
        If m_Index >= m_Length Then Exit Do
        If m_Chars(m_Index) = 44 Then m_Index = m_Index + 1
    Loop
End Sub

Public Sub Run()
    Dim textValue As String
    textValue = "x"
    LoadText textValue
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a class array allocated before recursive parsing should remain allocated: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesQualifiedClassAllocationToGuardedReader(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "JSON.cls", `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Option Explicit
Private chars() As Long
Private tokens() As Long
Private tokenCount As Long
Private textLength As Long

Friend Sub LoadText(ByVal text As String)
    Erase chars
    Erase tokens
    tokenCount = 0
    textLength = Len(text)
    If textLength = 0 Then Exit Sub
    ReDim chars(0 To textLength - 1)
    ReDim tokens(1 To 2)
    tokenCount = 1
End Sub

Friend Function ReadToken(ByVal tokenID As Long) As Long
    If tokenID < 1 Or tokenID > tokenCount Then Exit Function
    ReadToken = tokens(tokenID)
End Function

Public Function Parse(ByVal text As String) As JSON
    Dim doc As JSON
    Set doc = New JSON
    doc.LoadText text
    Set Parse = doc
End Function

Public Sub Run()
    Dim doc As JSON
    Set doc = Parse("x")
    Debug.Print doc.ReadToken(1)
End Sub
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub RunExternal()
    Dim parser As JSON
    Dim doc As JSON
    Set parser = New JSON
    Set doc = parser.Parse("x")
    Debug.Print doc.ReadToken(1)
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a guarded reader reached after qualified class initialization must retain the class array allocation: %+v", got)
	}
}

func TestAnalyzerVBA227UsesPositiveModuleCountAsArrayGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "CountGuard.cls", `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountGuard"
Option Explicit
Private values() As Long
Private valueCount As Long

Private Sub LoadValues()
    Erase values
    valueCount = 0
    ReDim values(0 To 1)
    valueCount = 1
End Sub

Public Function ValueAt(ByVal index As Long) As Long
    If index < 1 Or index > valueCount Then Exit Function
    ValueAt = values(index - 1)
End Function
`)

	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a positive module count after a completed allocation must guard the array access: %+v", got)
	}
}

func TestAnalyzerVBA227RejectsStalePositiveModuleCountAfterErase(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "StaleCount.cls", `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaleCount"
Option Explicit
Private values() As Long
Private valueCount As Long

Private Sub LoadValues()
    Erase values
    valueCount = 0
    ReDim values(0 To 1)
    valueCount = 1
End Sub

Private Sub ClearValues()
    Erase values
End Sub

Public Function ValueAt(ByVal index As Long) As Long
    If index < 1 Or index > valueCount Then Exit Function
    ValueAt = values(index - 1)
End Function

Public Sub Run()
    LoadValues
    ClearValues
    Debug.Print ValueAt(1)
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = true
	cfg.Analyze.DetectDeterministicRuntimeErrors = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) == 0 {
		t.Fatalf("an erase without resetting the count must not turn the count into an allocation proof: %+v", findings)
	}
}

func TestAnalyzerVBA227PreservesRecursiveModuleInvalidation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long

Private Sub ClearRecursively(ByVal depth As Long)
    Erase values
    If depth > 0 Then ClearRecursively depth - 1
End Sub

Public Sub Run()
    ReDim values(0 To 1)
    ClearRecursively 1
    Debug.Print values(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Procedure != "Run" {
		t.Fatalf("a recursive Erase must invalidate the caller module array: %+v", got)
	}
}

func TestAnalyzerVBA227DoesNotInvalidateAllocatedModuleArrayThroughReadOnlyRecursiveHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long

Private Sub Visit(ByVal depth As Long)
    If depth > 0 Then Visit depth - 1
    Debug.Print values(0)
End Sub

Public Sub Run()
    ReDim values(0 To 1)
    Visit 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a read-only recursive helper must preserve the caller module array allocation: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesHashStorageCapacityGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "HashStore.cls", `Attribute VB_Name = "HashStore"
Option Explicit
Private slots() As Long
Private codes() As Long
Private keys() As HashStore
Private values() As HashStore
Private capacity As Long
Private dirty As Boolean
Private supported As Boolean
Private unsafeCapacity As Long

Private Function HashIndexSupported() As Boolean
    HashIndexSupported = supported
End Function

Private Sub ResetHashIndex()
    Erase slots
    Erase codes
    Erase keys
    Erase values
    capacity = 0
End Sub

Private Sub RebuildHashIndex()
    If Not HashIndexSupported Then
        ResetHashIndex
        Exit Sub
    End If
    capacity = 2
    ReDim slots(0 To capacity - 1)
    ReDim codes(0 To capacity - 1)
    ReDim keys(0 To capacity - 1)
    ReDim values(0 To capacity - 1)
    InsertHashSlot 1
End Sub

Private Sub EnsureHashIndexCurrent()
    If Not dirty Then Exit Sub
    dirty = False
    RebuildHashIndex
End Sub

Private Function FindSlot(ByVal key As Long) As Long
    FindSlot = -1
    EnsureHashIndexCurrent
    If capacity = 0 Then Exit Function
    If slots(0) = key Then FindSlot = 0
    Set keys(0) = values(0)
End Function

Private Sub InsertHashSlot(ByVal key As Long)
    slots(0) = key
    Set keys(0) = values(0)
End Sub

Private Function UnsafeSlot(ByVal allocate As Boolean) As Long
    Dim unsafeValues() As Long
    UnsafeSlot = -1
    unsafeCapacity = 1
    If allocate Then ReDim unsafeValues(0 To 1)
    Debug.Print unsafeValues(0)
End Function

Public Sub Run()
    supported = True
    RebuildHashIndex
    Dim slot As Long
    slot = FindSlot(1)
    If slot < 0 Then Exit Sub
    Debug.Print slots(slot)
    Set keys(slot) = values(slot)
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 1 || got[0].Procedure != "UnsafeSlot" {
		t.Fatalf("a storage guard must not hide an unrelated unallocated array: %+v", got)
	}
}

func TestVBA249DoesNotReportPreserveAfterProvenModuleCapacityInitialization(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Cache.cls", `Attribute VB_Name = "Cache"
Option Explicit
Private tokens() As Long
Private tokenCount As Long

Private Sub Reset()
  Erase tokens
  tokenCount = 0
End Sub

Private Sub Ensure(ByVal additional As Long)
  Dim required As Long
  Dim newCapacity As Long
  required = additional
  If required <= tokenCount Then Exit Sub

  newCapacity = tokenCount
  If newCapacity = 0 Then
    newCapacity = 32
    Do While newCapacity < required
      newCapacity = newCapacity * 2
    Loop
    ReDim tokens(1 To newCapacity)
    tokenCount = newCapacity
    Exit Sub
  End If

  Do While newCapacity < required
    newCapacity = newCapacity * 2
  Loop
  ReDim Preserve tokens(1 To newCapacity)
  tokenCount = newCapacity
End Sub

Public Sub Run()
  Ensure 100
  Ensure 100
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("a module array proven allocated by its nonzero capacity must not report Preserve as unallocated: %+v", got)
	}
}

func TestVBA249DoesNotReportConfiguredModuleStorageConsumers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Storage.cls", `Attribute VB_Name = "Storage"
Option Explicit
Private tableItems() As Long
Private collectionItems() As Long

Private Sub ConfigureDataTable()
  ReDim tableItems(1 To 2)
End Sub

Private Sub ConfigureGenericCollection()
  ReDim collectionItems(1 To 2)
End Sub

Private Sub RebuildPrimaryKeyIndex()
  Dim index As Long
  For index = 1 To 2
    Debug.Print tableItems(index)
  Next index
End Sub

Private Sub IndexDataRow()
  Dim index As Long
  For index = 1 To 2
    Debug.Print tableItems(index)
  Next index
End Sub

Private Function CreateCollectionNode(ByVal index As Long) As Long
  CreateCollectionNode = collectionItems(index)
End Function

Public Sub Run()
  ConfigureDataTable
  ConfigureGenericCollection
  RebuildPrimaryKeyIndex
  IndexDataRow
  Debug.Print CreateCollectionNode(1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("configured module storage consumers must retain their allocation proof: %+v", got)
	}
}

func TestVBA249InvalidatesKnownValueAfterByRefMutation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Mutate(ByRef value As Long)
  value = 1
End Sub

Public Sub Run()
  Dim denominator As Long
  denominator = 0
  Mutate denominator
  Debug.Print 10 / denominator
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("ByRef mutation must invalidate the zero fact: %+v", got)
	}
}

func TestVBA249BatchAndRealtimeResultsMatchAndRemainNonBlocking(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Debug.Print 10 / 0
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	cfg := config.Default()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	realtime, err := SourceRealtimeFindings(dir, path, cfg, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got, want := findingsByCode(realtime, "VBA249"), findingsByCode(batch, "VBA249"); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime VBA249 findings differ:\nbatch=%+v\nrealtime=%+v", want, got)
	}
	if blocking := BlockingFindings(findingsByCode(batch, "VBA249")); len(blocking) != 0 {
		t.Fatalf("runtime-error findings must not block preflight: %+v", blocking)
	}
}

func TestVBA249DetectsStrictNumericConversionFailures(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim value As Long
  value = CInt("not numeric")
  value = CInt(40000)
  value = CInt(32767.4)
  value = CInt("40000")
  value = CInt(12)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 2 {
		t.Fatalf("expected only known conversion type/range failures, got %+v", got)
	}
	if got[0].RuntimeError == nil || got[1].RuntimeError == nil {
		t.Fatalf("conversion findings must carry runtime context: %+v", got)
	}
	for _, finding := range got {
		if finding.Line == 6 {
			t.Fatalf("banker's-rounded CInt(32767.4) must remain valid: %+v", got)
		}
	}
	kinds := []string{got[0].RuntimeError.Kind, got[1].RuntimeError.Kind}
	if !strings.Contains(strings.Join(kinds, ","), "conversion_type_mismatch") || !strings.Contains(strings.Join(kinds, ","), "conversion_overflow") {
		t.Fatalf("unexpected conversion kinds: %v", kinds)
	}
}

func TestVBA249DetectsNestedNumericConversionFailure(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim value As Long
  value = 1 + CInt(40000)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 1 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "conversion_overflow" {
		t.Fatalf("nested conversion overflow should be reported once: %+v", got)
	}
}

func TestVBA249KeepsUnrelatedArrayLifecycleWarningOnSameLine(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values(0 To 1) As Long
  Dim scalar As Long
  Debug.Print LBound(scalar): values(3) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got227 := findingsByCode(findings, "VBA227")
	got249 := findingsByCode(findings, "VBA249")
	if len(got227) != 1 || !strings.Contains(got227[0].Message, "LBound") {
		t.Fatalf("unrelated scalar LBound warning must remain on a line with a deterministic array failure: %+v", got227)
	}
	if len(got249) != 1 || got249[0].RuntimeError == nil || got249[0].RuntimeError.Kind != "array_subscript_out_of_bounds" {
		t.Fatalf("deterministic array bounds failure should remain on the same line: %+v", got249)
	}
}

func TestVBA249DetectsKnownArrayBoundsAndKeepsUnknownShapesSilent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
  Dim fixed(0 To 1) As Long
  Dim values() As Long
  fixed(2) = 1
  If chooseFirst Then
    ReDim values(0 To 1)
  Else
    values = ExternalArray()
  End If
  values(2) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 1 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_subscript_out_of_bounds" {
		t.Fatalf("only the fixed-array out-of-bounds access should be deterministic: %+v", got)
	}
}

func TestVBA227DoesNotPromoteConditionalByRefAllocationToDeterministicFailure(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal arguments As Collection)
	Dim values() As Variant
	PrepareInvocationValues arguments, values
	Select Case arguments.Count
	  Case 1
	  Debug.Print values(0)
	End Select
End Sub

Private Sub PrepareInvocationValues(ByVal arguments As Collection, ByRef values() As Variant)
	If arguments.Count = 0 Then Exit Sub
	ReDim values(0 To arguments.Count - 1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("the caller's positive guard must exclude deterministic unallocated-array failure: %+v", got)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("the caller's positive count case proves the conditional ByRef output allocated: %+v", got)
	}
}

func TestVBA249PreservesModuleArrayAllocationAcrossPlainReDim(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "ModuleArray.cls", `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleArray"
Option Explicit
Private values() As Variant

Public Function LoadValues() As Boolean
  On Error GoTo ErrTrap
  Dim maxIndex As Integer
  maxIndex = 37
  ReDim values(maxIndex, 2)
  values(0, 0) = "first"
  values(37, 2) = "last"
  LoadValues = True
  Exit Function
ErrTrap:
  Resume ExitProcedure
ExitProcedure:
End Function
`)

	for _, strategy := range []arrayCFGStrategy{arrayCFGStrategyLegacy, arrayCFGStrategyCompact} {
		findings, err := (Analyzer{RootDir: dir, Config: config.Default(), arrayStrategy: strategy}).Run()
		if err != nil {
			t.Fatal(err)
		}
		if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
			t.Fatalf("strategy %d: a plain ReDim must establish a module array before indexed writes: %+v", strategy, got)
		}
	}
}

func TestVBA249PreservesModuleArrayAllocationAcrossLoopReDimPreserve(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "LoopArray.cls", `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoopArray"
Option Explicit
Private values() As String

Public Sub LoadValues()
  On Error GoTo ErrTrap
  Dim index As Integer
  ReDim values(2, 0)
  For index = 0 To 1
    ReDim Preserve values(2, index)
    values(0, index) = "value"
  Next index
ExitProcedure:
  On Error Resume Next
  Exit Sub
ErrTrap:
  Resume ExitProcedure
End Sub
`)

	for _, strategy := range []arrayCFGStrategy{arrayCFGStrategyLegacy, arrayCFGStrategyCompact} {
		findings, err := (Analyzer{RootDir: dir, Config: config.Default(), arrayStrategy: strategy}).Run()
		if err != nil {
			t.Fatal(err)
		}
		if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
			t.Fatalf("strategy %d: an allocated module array must remain allocated through loop Preserve: %+v", strategy, got)
		}
	}
}
