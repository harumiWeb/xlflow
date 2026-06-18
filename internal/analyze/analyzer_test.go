package analyze

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

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

func findingsByCode(findings []Finding, code string) []Finding {
	var matches []Finding
	for _, finding := range findings {
		if finding.Code == code {
			matches = append(matches, finding)
		}
	}
	return matches
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
