package analyze

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestSourceRealtimeFindingsParsedMatchesSource(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	source := []byte("Option Explicit\nPublic Sub Run()\n  Dim found As Range\n  Set found = Range(\"A1\").Find(What:=\"x\")\n  Debug.Print found.Value\nEnd Sub\n")
	cfg := config.Default()
	want, err := SourceRealtimeFindings(dir, path, cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	got, err := SourceRealtimeFindingsParsed(dir, cfg, doc)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("SourceRealtimeFindingsParsed = %+v, want %+v", got, want)
	}
}

func TestAnalyzerFindsMissingSetForObjectVariable(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  ws = ThisWorkbook.Worksheets("Sheet1")
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA101", 4)
}

func TestAnalyzerFindsMissingSetForModuleLevelObjectVariable(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private ws As Worksheet
Public Sub Run()
  ws = ThisWorkbook.Worksheets("Sheet1")
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA101", 4)
}

func TestAnalyzerFindsMissingSetForObjectReturningFunction(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function FindRange() As Range
  Set FindRange = Sheet1.Range("A1")
End Function
Public Sub Run()
  Dim result As Range
  result = FindRange()
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA102", 7)
}

func TestAnalyzerFindsMissingSetInObjectReturningFunction(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function GetSheet() As Worksheet
  GetSheet = ThisWorkbook.Worksheets(1)
End Function
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA103", 3)
	finding := findFinding(t, findings, "VBA103", 3)
	if !containsAll(finding.Suggestion, "Set GetSheet = ...", "Worksheet") {
		t.Fatalf("unexpected VBA103 suggestion: %q", finding.Suggestion)
	}
}

func TestAnalyzerIgnoresScalarAndSetAssignments(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function FindRange() As Range
  Set FindRange = Sheet1.Range("A1")
End Function
Public Sub Run()
  Dim n As Long
  n = 1
  Dim result As Range
  Set result = FindRange()
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(findings) != 0 {
		t.Fatalf("expected no findings, got %+v", findings)
	}
}

func TestAnalyzerDoesNotReportObjectFunctionAssignmentToScalar(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function FindRange() As Range
  Set FindRange = Sheet1.Range("A1")
End Function
Public Sub Run()
  Dim counter As Long
  counter = FindRange()
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findings {
		if finding.Code == "VBA102" {
			t.Fatalf("VBA102 should require an object-typed target variable: %+v", findings)
		}
	}
}

func TestAnalyzerFailsOnParserRecovery(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Broken(ByVal value As String
End Function
`)

	_, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err == nil {
		t.Fatal("expected parser recovery error")
	}
	if !strings.Contains(err.Error(), "VBA parser reported errors or missing nodes") {
		t.Fatalf("unexpected parse error: %v", err)
	}
}

func TestAnalyzerFindsWorksheetMemberAssignedOnVariable(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws.DisplayGridlines = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA104", 5)
}

func TestAnalyzerFindsWorksheetMemberOnModuleLevelVariable(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private ws As Worksheet
Public Sub Run()
  Set ws = ThisWorkbook.Worksheets(1)
  ws.DisplayGridlines = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA104", 5)
}

func TestAnalyzerFindsWorksheetMemberAssignedInsideWithBlock(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  With ws
    .DisplayGridlines = False
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA104", 6)
}

func TestAnalyzerFindsMissingXlflowLogHelperSource(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Call XlflowLog("start")
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA105", 3)
	finding := findFinding(t, findings, "VBA105", 3)
	if !containsAll(finding.Suggestion, "XlflowDebug.Log", "xlflow run --json") {
		t.Fatalf("unexpected VBA105 suggestion: %q", finding.Suggestion)
	}
}

func TestAnalyzerFindsMissingXlflowSetTraceFileHelperSource(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  XlflowTrace.XlflowSetTraceFile "C:\Temp\xlflow\trace.log"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA106", 3)
	finding := findFinding(t, findings, "VBA106", 3)
	if !containsAll(finding.Suggestion, "XlflowDebug.Log", "xlflow run --json") {
		t.Fatalf("unexpected VBA106 suggestion: %q", finding.Suggestion)
	}
}

func TestAnalyzerStillFlagsLegacyTraceHelpersWhenHelperSourceExists(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Call XlflowLog("start")
  XlflowTrace.XlflowSetTraceFile "C:\Temp\xlflow\trace.log"
End Sub
`)
	writeModule(t, dir, "XlflowTrace.bas", `Option Explicit
Public Sub XlflowLog(ByVal message As String)
End Sub
Public Sub XlflowSetTraceFile(ByVal path As String)
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA105", 3)
	assertFinding(t, findings, "VBA106", 4)
}

func TestAnalyzerSidecarModeSkipsGeneratedFRMCodeDiagnostics(t *testing.T) {
	dir := t.TempDir()
	formsDir := filepath.Join(dir, "src", "forms")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	frmBody := "VERSION 5.00\nBegin {GUID} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n\nPublic Sub BreakAnalyzer()\n  Dim ws As Worksheet\n  Set ws = ThisWorkbook.Worksheets(1)\n  ws.DisplayGridlines = True\nEnd Sub\n"
	sidecarBody := "Option Explicit\n\nPublic Sub BreakAnalyzer()\n  Dim ws As Worksheet\n  Set ws = ThisWorkbook.Worksheets(1)\n  ws.DisplayGridlines = True\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "UserForm1.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	var vba104 []Finding
	for _, finding := range findings {
		if finding.Code == "VBA104" {
			vba104 = append(vba104, finding)
		}
	}
	if len(vba104) != 1 {
		t.Fatalf("expected one VBA104 finding from sidecar mode, got %+v", vba104)
	}
	if vba104[0].File != "src/forms/code/UserForm1.bas" {
		t.Fatalf("expected sidecar file to be authoritative, got %+v", vba104[0])
	}
}

func TestAnalyzerFindsDefaultRuntimeRiskRules(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim found As Range
  Dim ws As Worksheet
  Set found = Range("A:A").Find("x")
  Debug.Print found.Value
  ws.Range("A1").Value = 1
  Application.EnableEvents = False
  On Error GoTo ErrHandler
  Debug.Print "work"
ErrHandler:
  Debug.Print Err.Description
  Dim values() As Variant
  ReDim Preserve values(1 To 2, 1 To 3)
  If ws = Nothing Then Debug.Print "missing"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for code, line := range map[string]int{
		"VBA201": 6,
		"VBA202": 7,
		"VBA203": 8,
		"VBA204": 11,
		"VBA205": 5,
		"VBA208": 14,
		"VBA209": 15,
	} {
		assertFinding(t, findings, code, line)
	}
}

func TestAnalyzerHonorsDisabledRuleIDs(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim found As Range
  Set found = Range("A:A").Find("x")
  found.Value = 1
End Sub
`)
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
disabled_rules = ["VBA205"]
`)
	if err := os.WriteFile(filepath.Join(dir, config.FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA205"); len(got) != 0 {
		t.Fatalf("VBA205 should be disabled: %+v", got)
	}
	if got := findingsByCode(findings, "VBA201"); len(got) == 0 {
		t.Fatalf("VBA201 should remain enabled: %+v", findings)
	}
}

func TestAnalyzerSupportsInlineSuppressions(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ' xlflow:disable-next-line VBA205
  Range("A1").Value = 1
  Cells(1, 1).Value = 2 ' xlflow:disable-line VBA205
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA205"); len(got) != 0 {
		t.Fatalf("VBA205 should be suppressed: %+v", got)
	}
}

func TestAnalyzerReportsUnknownAndUnusedInlineSuppressionsAsWarnings(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ' xlflow:disable-next-line VBA999
  Debug.Print "ok"
  ' xlflow:disable-next-line VBA205
  Debug.Print "still ok"
End Sub
`)

	result, err := Analyzer{RootDir: dir, Config: config.Default()}.RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Findings) != 0 {
		t.Fatalf("findings = %+v, want none", result.Findings)
	}
	if !hasWarning(result.Warnings, "unknown_inline_suppression_rule", "VBA999") {
		t.Fatalf("expected unknown suppression warning, got %+v", result.Warnings)
	}
	if !hasWarning(result.Warnings, "unused_inline_suppression", "VBA205") {
		t.Fatalf("expected unused suppression warning, got %+v", result.Warnings)
	}
}

func TestAnalyzerDoesNotSuppressPreflightBlockingDiagnostics(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ' xlflow:disable-next-line VBA104
  ws.DisplayGridlines = False
End Sub
`)

	result, err := Analyzer{RootDir: dir, Config: config.Default()}.RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(result.Findings, "VBA104"); len(got) != 1 {
		t.Fatalf("VBA104 should remain unsuppressed: findings=%+v warnings=%+v", result.Findings, result.Warnings)
	}
	if !hasWarning(result.Warnings, "unsupported_inline_suppression_rule", "VBA104") {
		t.Fatalf("expected unsupported suppression warning, got %+v", result.Warnings)
	}
}

func TestAnalyzerRuntimeRiskRulesAllowGuardedPatterns(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Build() As Range
  Set Build = Sheet1.Range("A1")
End Function
Public Sub Run()
  Dim found As Range
  Dim ws As Worksheet
  Dim oldEvents As Boolean
  oldEvents = Application.EnableEvents
  Set ws = ThisWorkbook.Worksheets(1)
  Set found = ws.Range("A:A").Find("x")
  If found Is Nothing Then Exit Sub
  Debug.Print found.Value
  Application.EnableEvents = False
Cleanup:
  Application.EnableEvents = oldEvents
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VBA201", "VBA202", "VBA203", "VBA204", "VBA205", "VBA210"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("%s should not trigger for guarded pattern: %+v", code, got)
		}
	}
}

func TestAnalyzerApplicationStateAllowsPushPopRestorePattern(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private fastModeDepth As Long
Private savedCalculation As XlCalculation
Private savedEnableEvents As Boolean
Private savedScreenUpdating As Boolean

Private Sub PushFastMode()
  If fastModeDepth = 0 Then
    savedCalculation = Application.Calculation
    savedEnableEvents = Application.EnableEvents
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
  End If
  fastModeDepth = fastModeDepth + 1
End Sub

Private Sub PopFastMode()
  If fastModeDepth <= 0 Then Exit Sub
  fastModeDepth = fastModeDepth - 1
  If fastModeDepth = 0 Then
    Application.Calculation = savedCalculation
    Application.EnableEvents = savedEnableEvents
    Application.ScreenUpdating = savedScreenUpdating
  End If
End Sub

Public Sub Run()
  Call PushFastMode
  On Error GoTo Cleanup
  Debug.Print "work"
Cleanup:
  Call PopFastMode
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA203"); len(got) != 0 {
		t.Fatalf("VBA203 should allow paired Push/Pop Application state restore: %+v", got)
	}
}

func TestAnalyzerApplicationStateStillFlagsUnpairedPushPattern(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.ScreenUpdating = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA203", 3)
}

func TestAnalyzerApplicationStateAllowsPropagatedRestoreEffect(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
End Sub

Private Sub RestoreEvents()
  Application.EnableEvents = True
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub PopFastMode()
  RestoreEvents
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA203"); len(got) != 0 {
		t.Fatalf("VBA203 should accept a uniquely propagated restore effect: %+v", got)
	}
}

func TestAnalyzerApplicationStateAllowsUniqueProjectVisibleRestorePair(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
End Sub
`)
	writeModule(t, dir, "StateHelpers.bas", `Option Explicit
Public Sub PopFastMode()
  RestoreEvents
End Sub

Private Sub RestoreEvents()
  Application.EnableEvents = True
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA203"); len(got) != 0 {
		t.Fatalf("VBA203 should accept one project-visible paired restore: %+v", got)
	}
}

func TestAnalyzerApplicationStateRejectsAmbiguousProjectRestorePair(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
End Sub
`)
	for _, name := range []string{"StateHelpersA.bas", "StateHelpersB.bas"} {
		writeModule(t, dir, name, `Option Explicit
Public Sub PopFastMode()
  Application.EnableEvents = True
End Sub
`)
	}

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA203", 3)
}

func TestAnalyzerApplicationStateRejectsCrossModuleClassMethodPair(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
End Sub
`)
	writeClass(t, dir, "StateHelper.cls", `Option Explicit
Public Sub PopFastMode()
  Application.EnableEvents = True
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA203", 3)
}

func TestAnalyzerApplicationStateRejectsRestorePropagatedFromBareClassMethod(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
End Sub
`)
	writeModule(t, dir, "Helpers.bas", `Option Explicit
Public Sub PopFastMode()
  RestoreEvents
End Sub
`)
	writeClass(t, dir, "StateClass.cls", `Option Explicit
Public Sub RestoreEvents()
  Application.EnableEvents = True
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA203", 3)
}

func TestAnalyzerApplicationStateRejectsRestorePropagatedFromBareUserFormMethod(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
End Sub
`)
	writeModule(t, dir, "Helpers.bas", `Option Explicit
Public Sub PopFastMode()
  RestoreEvents
End Sub
`)
	writeFormSidecar(t, dir, "UserForm1.bas", `Option Explicit
Public Sub RestoreEvents()
  Application.EnableEvents = True
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA203", 3)
}

func TestAnalyzerApplicationStatePreservesInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  ' xlflow:disable-next-line VBA203
  Application.EnableEvents = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA203"); len(got) != 0 {
		t.Fatalf("VBA203 inline suppression should remain effective: %+v", got)
	}
}

func TestAnalyzerErrorHandlerFallthroughSuggestsConcreteExitStatement(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error GoTo ExitSub
  Debug.Print "work"
ExitSub:
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	finding := findFinding(t, findings, "VBA204", 5)
	if !containsAll(finding.Suggestion, "`Exit Sub`", "`ExitSub:`") {
		t.Fatalf("unexpected VBA204 suggestion: %q", finding.Suggestion)
	}
}

func TestAnalyzerIRVBA204PreservesPropertyExitAndCleanupSemantics(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Property Get SafeValue() As Long
  On Error GoTo Handler
  SafeValue = 1
  Exit Property
Handler:
  SafeValue = 0
End Property

Public Property Get UnsafeValue() As Long
  On Error GoTo Handler
  UnsafeValue = 1
Handler:
  UnsafeValue = 0
End Property

Public Sub CleanupAllowed()
  On Error GoTo Cleanup
  Debug.Print "work"
Cleanup:
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA204")
	if len(got) != 1 || got[0].Procedure != "UnsafeValue" || got[0].Line != 13 {
		t.Fatalf("VBA204 findings = %+v, want only UnsafeValue handler", got)
	}
	if !containsAll(got[0].Suggestion, "`Exit Property`", "`Handler:`") {
		t.Fatalf("unexpected property suggestion: %+v", got[0])
	}
}

func TestAnalyzerIRVBA204PreservesInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error GoTo Handler
  Debug.Print "work"
  ' xlflow:disable-next-line VBA204
Handler:
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA204"); len(got) != 0 {
		t.Fatalf("VBA204 should remain suppressible: %+v", got)
	}
}

func TestAnalyzerIRVBA204DoesNotTreatNestedExitAsUnconditional(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal stopEarly As Boolean)
  On Error GoTo Handler
  If stopEarly Then
    Exit Sub
  End If
Handler:
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA204")
	if len(got) != 1 || got[0].Procedure != "Run" || got[0].Line != 7 {
		t.Fatalf("VBA204 findings = %+v, want conditional Exit Sub fallthrough", got)
	}
}

func TestAnalyzerIRVBA204DoesNotNestHandlerAfterSingleLineIf(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal stopEarly As Boolean)
  On Error GoTo Handler
  If stopEarly Then Exit Sub
Handler:
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA204")
	if len(got) != 1 || got[0].Procedure != "Run" || got[0].Line != 5 {
		t.Fatalf("VBA204 findings = %+v, want single-line If fallthrough", got)
	}
}

func TestAnalyzerCFGVBA204DoesNotReportHandlerSkippedByGoto(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error GoTo Handler
  Debug.Print "work"
  GoTo Done
Handler:
  Debug.Print "failed"
Done:
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA204"); len(got) != 0 {
		t.Fatalf("VBA204 should not report a handler skipped by normal GoTo: %+v", got)
	}
}

func TestAnalyzerCFGVBA204DoesNotTreatGotoHandlerAsFallthrough(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error GoTo Handler
  GoTo Handler
  Exit Sub
Handler:
  Debug.Print "handled"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA204"); len(got) != 0 {
		t.Fatalf("VBA204 should not treat explicit GoTo Handler as fallthrough: %+v", got)
	}
}

func TestAnalyzerChecksObjectUseOnSetAssignmentRHS(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Dim rng As Range
  Set rng = ws.Range("A1")
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA202", 5)
}

func TestAnalyzerDoesNotTreatAnyObjectMentionAsInitialization(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  If ws Is Nothing Then Debug.Print "missing"
  ws.Range("A1").Value = 1
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA202", 5)
}

func TestAnalyzerAllowsKnownByRefObjectInitializer(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub InitSheet(ByRef target As Worksheet)
  Set target = ThisWorkbook.Worksheets(1)
End Sub
Public Sub Run()
  Dim ws As Worksheet
  InitSheet ws
  ws.Range("A1").Value = 1
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("known ByRef object initializer should suppress VBA202: %+v", got)
	}
}

func TestAnalyzerOptInRuntimeRiskRules(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub NeedsLong(ByRef value As Long)
End Sub
Public Function MissingReturn() As Range
End Function
Public Sub Run()
  Dim dict As Dictionary
  Set dict = CreateObject("Scripting.Dictionary")
  NeedsLong "abc"
  Debug.Print dict("missing")
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectByRefArgumentMismatch = true
	cfg.Analyze.DetectDictionaryCollectionGuard = true
	cfg.Analyze.DetectFunctionReturnPath = true

	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA206", 9)
	assertFinding(t, findings, "VBA207", 10)
	assertFinding(t, findings, "VBA210", 4)
}

func TestAnalyzerByRefMismatchHandlesLowercaseCallKeyword(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub NeedsLong(ByRef value As Long)
End Sub
Public Sub Run()
  call NeedsLong("abc")
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectByRefArgumentMismatch = true

	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA206", 5)
}

func TestAnalyzerArrayComparisonUsesIdentifierBoundaries(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim a() As Variant
  Dim total As Long
  Dim amount As Long
  If total = amount Then Debug.Print "ok"
  If a = amount Then Debug.Print "bad"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA209")
	if len(got) != 1 || got[0].Line != 7 {
		t.Fatalf("expected only array comparison on line 7, got %+v", got)
	}
}

func TestAnalyzerArrayComparisonIgnoresFunctionReturnAssignment(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function CopyValues() As Variant
  Dim values() As String
  ReDim values(0 To 0)
  values(0) = "value"
	Let CopyValues = values
End Function

Public Sub Run()
  Dim values() As String
  If values = "unexpected" Then Debug.Print "bad"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA209")
	if len(got) != 1 || got[0].Line != 11 {
		t.Fatalf("expected only the array comparison on line 11, got %+v", got)
	}
}

func TestAnalyzerArrayComparisonIgnoresElementsAndProcedureHeaders(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private fastModeDepth As Long
Private Type DownloadItem
  Filename As String
End Type
Private C_Down() As DownloadItem
Private TestArray() As Variant

Private Sub PopFastMode()
  If fastModeDepth <= 0 Then Exit Sub
  fastModeDepth = fastModeDepth - 1
End Sub

Public Sub Run()
  Dim values() As Variant
  Dim dataMatrix() As Variant
  Dim stepData() As Variant
  Dim columnIndex As Long
  Dim stepIndex As Long
  Dim outputColumn As Long
  Dim rowIndex As Long
  Dim i As Long
  values(1, columnIndex) = i
  dataMatrix(stepIndex, outputColumn) = stepData(rowIndex, 1)
  C_Down(i).Filename = TestArray(i, 0)
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA209"); len(got) != 0 {
		t.Fatalf("VBA209 should ignore array element and UDT-array member assignments: %+v", got)
	}
}

func TestAnalyzerExpandedExcelMemberMismatch(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws.ScreenUpdating = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA211", 5)
}

func TestAnalyzerDetectsNonShortCircuitObjectGuards(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim deck As Collection
  Dim hand As Collection
  Dim bag As Collection
  Dim cards As Collection
  If deck Is Nothing Or deck.Count = 0 Then Exit Sub
  If hand Is Nothing Or hand.Item(1) Is Nothing Then Exit Sub
  If Not bag Is Nothing And bag.Count > 0 Then Debug.Print bag.Count
  If Not cards Is Nothing And cards.Item(1) Then Debug.Print cards.Count
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, line := range []int{7, 8, 9, 10} {
		assertFinding(t, findings, "VBA212", line)
	}
	finding := findFinding(t, findings, "VBA212", 7)
	if !containsAll(finding.Message, "deck", "non-short-circuit") ||
		!containsAll(finding.Reason, "And/Or", "runtime error 91") ||
		!containsAll(finding.Suggestion, "separate If statements") {
		t.Fatalf("unexpected VBA212 text: %+v", finding)
	}
}

func TestAnalyzerNonShortCircuitObjectGuardAllowsSeparateAndDifferentObjects(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim deck As Collection
  Dim other As Collection
  If deck Is Nothing Then Exit Sub
  If deck.Count = 0 Then Exit Sub
  If other Is Nothing Or deck.Count = 0 Then Exit Sub
  Debug.Print "If deck Is Nothing Or deck.Count = 0 Then"
  ' If deck Is Nothing Or deck.Count = 0 Then
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA212"); len(got) != 0 {
		t.Fatalf("VBA212 should allow safe or unrelated patterns, got %+v", got)
	}
}

func TestAnalyzerNonShortCircuitObjectGuardCanBeDisabled(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim deck As Collection
  If deck Is Nothing Or deck.Count = 0 Then Exit Sub
End Sub
`)
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
disabled_rules = ["VBA212"]
`)
	if err := os.WriteFile(filepath.Join(dir, config.FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA212"); len(got) != 0 {
		t.Fatalf("VBA212 should be disabled: %+v", got)
	}
}

func TestAnalyzerNonShortCircuitObjectGuardDedupesMultilineExpression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim deck As Collection
  Dim hand As Collection
  If deck Is Nothing Or deck.Count = 0 Or _
     hand Is Nothing Or hand.Count = 0 Then Exit Sub
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA212")
	if len(got) != 2 {
		t.Fatalf("VBA212 findings = %+v, want one finding per guarded object", got)
	}
	counts := map[string]int{}
	for _, finding := range got {
		counts[finding.Message]++
	}
	for message, count := range counts {
		if count != 1 {
			t.Fatalf("VBA212 duplicate finding for %q: %+v", message, got)
		}
	}
}

func TestAnalyzerNonShortCircuitObjectGuardSupportsInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim deck As Collection
  ' xlflow:disable-next-line VBA212
  If deck Is Nothing Or deck.Count = 0 Then Exit Sub
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA212"); len(got) != 0 {
		t.Fatalf("VBA212 should be suppressed: %+v", got)
	}
}

func TestAnalyzerDictionaryIterationValueUsageFindsKnownAndInferredDictionaries(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim earlyBound As Scripting.Dictionary
  Dim lateBound As Object
  Dim replacement As Object
  Dim item As Variant
  Set lateBound = CreateObject("Scripting.Dictionary")
  Set replacement = New Scripting.Dictionary
  For Each item In earlyBound
    Debug.Print item.Name
  Next item
  For Each item In lateBound
    Debug.Print item.Caption
  Next item
  For Each item In replacement
    Debug.Print item.Value
  Next item
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryIterationValueUsage = true
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA213")
	if len(got) != 3 {
		t.Fatalf("VBA213 findings = %+v, want three direct dictionary iteration findings", got)
	}
	for _, finding := range got {
		if !containsAll(finding.Reason, "Dictionary iteration yields keys") ||
			!containsAll(finding.Suggestion, ".Items", "(") {
			t.Fatalf("unexpected VBA213 finding: %+v", finding)
		}
	}
}

func TestAnalyzerDictionaryIterationValueUsageFindsObjectAssignmentAndIgnoresKeyUsage(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim dict As Dictionary
  Dim key As Variant
  Dim value As Object
  For Each key In dict
    Debug.Print key, dict(key)
  Next key
  For Each key In dict.Items
    Debug.Print key.Name
  Next key
  For Each key In dict
    For Each value In dict.Items
      Debug.Print value.Name
    Next value
    Set value = key
  Next key
  Debug.Print key.Name
  ' Debug.Print key.Name
  Debug.Print "key.Name"
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryIterationValueUsage = true
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA213")
	if len(got) != 1 || got[0].Line != 16 || !strings.Contains(got[0].Message, "key") {
		t.Fatalf("VBA213 findings = %+v, want only Set value = key on line 16", got)
	}
}

func TestAnalyzerDictionaryIterationValueUsageIsOptInAndInvalidatesInference(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim dict As Object
  Dim item As Variant
  Set dict = CreateObject("Scripting.Dictionary")
  Set dict = CreateObject("Scripting.FileSystemObject")
  For Each item In dict
    Debug.Print item.Name
  Next item
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA213"); len(got) != 0 {
		t.Fatalf("VBA213 should be opt-in and ignore invalidated inference: %+v", got)
	}
}

func TestAnalyzerDictionaryIterationValueUsageRecognizesWithAndAllowsReboundValues(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim dict As Dictionary
  Dim item As Variant
  For Each item In dict
    With item
      Debug.Print .Name
    End With
  Next item
  For Each item In dict
    Set item = dict(item)
    Debug.Print item.Name
  Next item
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryIterationValueUsage = true
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA213")
	if len(got) != 1 || got[0].Line != 6 || !strings.Contains(got[0].Message, "item") {
		t.Fatalf("VBA213 findings = %+v, want only With item on line 6", got)
	}
}

func TestAnalyzerRuntimeRiskRulesIgnoreCommentsAndStrings(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Debug.Print "Range(""A1"") On Error GoTo ErrHandler Application.EnableEvents = False"
  ' Set found = Range("A:A").Find("x")
  ' Debug.Print found.Value
  ' ErrHandler:
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VBA201", "VBA203", "VBA204", "VBA205"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("%s should ignore comments and strings: %+v", code, got)
		}
	}
}

func writeModule(t *testing.T, dir, name, body string) {
	t.Helper()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(src, name), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func writeClass(t *testing.T, dir, name, body string) {
	t.Helper()
	src := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(src, name), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func writeFormSidecar(t *testing.T, dir, name, body string) {
	t.Helper()
	src := filepath.Join(dir, "src", "forms", "code")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(src, name), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	formName := strings.TrimSuffix(name, filepath.Ext(name)) + ".frm"
	if err := os.WriteFile(filepath.Join(dir, "src", "forms", formName), []byte("VERSION 5.00\nBegin VB.UserForm "+strings.TrimSuffix(name, filepath.Ext(name))+"\nEnd\n"), 0o644); err != nil {
		t.Fatal(err)
	}
}

func findingsByCode(findings []Finding, code string) []Finding {
	var matches []Finding
	for _, finding := range findings {
		if finding.Code == code {
			matches = append(matches, finding)
		}
	}
	return matches
}

func hasWarning(warnings []map[string]any, code string, rule string) bool {
	for _, warning := range warnings {
		if warning["code"] == code && warning["rule"] == rule {
			return true
		}
	}
	return false
}

func assertFinding(t *testing.T, findings []Finding, code string, line int) {
	t.Helper()
	for _, finding := range findings {
		if finding.Code == code && finding.Line == line && len(finding.NearbyCode) > 0 && finding.File == "src/modules/Main.bas" {
			return
		}
	}
	t.Fatalf("missing %s line %d in %+v", code, line, findings)
}

func findFinding(t *testing.T, findings []Finding, code string, line int) Finding {
	t.Helper()
	for _, finding := range findings {
		if finding.Code == code && finding.Line == line {
			return finding
		}
	}
	t.Fatalf("missing %s line %d in %+v", code, line, findings)
	return Finding{}
}

func containsAll(text string, parts ...string) bool {
	for _, part := range parts {
		if !strings.Contains(text, part) {
			return false
		}
	}
	return true
}
