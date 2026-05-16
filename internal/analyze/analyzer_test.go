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
	if !containsAll(finding.Suggestion, "xlflow trace enable", "xlflow run --trace") {
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
	if !containsAll(finding.Suggestion, "xlflow run --trace", "xlflow trace enable") {
		t.Fatalf("unexpected VBA106 suggestion: %q", finding.Suggestion)
	}
}

func TestAnalyzerDoesNotFlagTraceHelperCallsWhenHelperSourceExists(t *testing.T) {
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
	for _, finding := range findings {
		if finding.Code == "VBA105" || finding.Code == "VBA106" {
			t.Fatalf("unexpected missing trace helper finding: %+v", findings)
		}
	}
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
