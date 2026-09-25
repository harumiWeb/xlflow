package analyze

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func runUdfCellReferenceAnalysis(t *testing.T, modules map[string]string) []Finding {
	t.Helper()
	dir := t.TempDir()
	for name, source := range modules {
		writeModule(t, dir, name, source)
	}
	cfg := config.Default()
	cfg.Analyze.DetectUdfCellReferenceNames = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	return findings
}

func runUdfCellReferenceProjectAnalysis(t *testing.T, files []sourceproject.SourceFile) []Finding {
	t.Helper()
	cfg := config.Default()
	cfg.Analyze.DetectUdfCellReferenceNames = true
	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), sourceproject.SourceProject{Files: files})
	if err != nil {
		t.Fatal(err)
	}
	return result.Findings
}

func TestUdfCellReferenceReportsPublicFunctionNamedAfterA1Reference(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function ABC123() As Double
  ABC123 = 42
End Function
`})
	got := findingsByCode(findings, "VBA270")
	if len(got) != 1 || !strings.Contains(got[0].Message, "ABC123") || !strings.Contains(got[0].Message, "A1-style") {
		t.Fatalf("VBA270 findings = %+v, want one finding naming ABC123 as A1-style", got)
	}
	if got[0].Line != 2 {
		t.Fatalf("VBA270 line = %d, want the Function declaration line", got[0].Line)
	}
}

func TestUdfCellReferenceHighlightsOnlyTheNameToken(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function ABC123() As Double
  Dim counter As Long
  counter = 1
  ABC123 = counter
End Function
`})
	got := findingsByCode(findings, "VBA270")
	if len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want one finding", got)
	}
	finding := got[0]
	if finding.Line != 2 || finding.Column != 17 || finding.EndLine != 2 || finding.EndColumn != 23 {
		t.Fatalf("VBA270 range = %d:%d-%d:%d, want the ABC123 name token span 2:17-2:23", finding.Line, finding.Column, finding.EndLine, finding.EndColumn)
	}
}

func TestUdfCellReferenceReportsZeroPaddedRowNames(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function A00000000000000000001() As Double
  A00000000000000000001 = 1
End Function
Public Function A0001048576() As Double
  A0001048576 = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 2 {
		t.Fatalf("VBA270 findings = %+v, want two findings for zero-padded row names", got)
	}
}

func TestUdfCellReferenceScansIntactFunctionsBesideParseErrors(t *testing.T) {
	// An Optional default using a radix literal is an accepted parser-recovery
	// shape: the file still carries parse errors, but analysis continues and
	// unrelated intact declarations must still be scanned.
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function TakesDefault(Optional p As Long = &HFF) As Double
  TakesDefault = p
End Function
Public Function A1() As Double
  A1 = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want the intact Function flagged despite the unrelated parse error", got)
	}
}

func TestUdfCellReferenceReportsImplicitPublicFunction(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Function A1() As Double
  A1 = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want one finding for implicit Public Function", got)
	}
}

func TestUdfCellReferenceReportsCaseInsensitiveAndBoundaryNames(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function xfd1048576() As Double
  xfd1048576 = 1
End Function
Public Function A1048576() As Double
  A1048576 = 1
End Function
Public Function XFD1048576() As Double
  XFD1048576 = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 3 {
		t.Fatalf("VBA270 findings = %+v, want three findings at the XFD/1048576 boundary", got)
	}
}

func TestUdfCellReferenceSkipsOutOfRangeAndNonReferenceNames(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function XFE1() As Double
  XFE1 = 1
End Function
Public Function XFE1048576() As Double
  XFE1048576 = 1
End Function
Public Function A1048577() As Double
  A1048577 = 1
End Function
Public Function A0() As Double
  A0 = 1
End Function
Public Function AAAA1() As Double
  AAAA1 = 1
End Function
Public Function A1B() As Double
  A1B = 1
End Function
Public Function AB12C() As Double
  AB12C = 1
End Function
Public Function Helper() As Double
  Helper = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for out-of-range or non-reference names", got)
	}
}

func TestUdfCellReferenceReportsR1C1StyleNames(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function R1C1() As Double
  R1C1 = 1
End Function
Public Function R1048576C16384() As Double
  R1048576C16384 = 1
End Function
`})
	got := findingsByCode(findings, "VBA270")
	if len(got) != 2 {
		t.Fatalf("VBA270 findings = %+v, want two R1C1-style findings", got)
	}
	for _, f := range got {
		if !strings.Contains(f.Message, "R1C1-style") {
			t.Fatalf("VBA270 message = %q, want the R1C1-style label", f.Message)
		}
	}
}

func TestUdfCellReferenceReportsSingleColumnA1Names(t *testing.T) {
	// R5 and C3 are plain A1-style references (column R or C, row 5 or 3),
	// independent of their R1C1 row/column interpretation.
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function R5() As Double
  R5 = 1
End Function
Public Function C3() As Double
  C3 = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 2 {
		t.Fatalf("VBA270 findings = %+v, want two A1-style findings", got)
	}
}

func TestUdfCellReferenceSkipsR1C1OutOfRangeAndNonCellForms(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Public Function R1048577C1() As Double
  R1048577C1 = 1
End Function
Public Function R1C16385() As Double
  R1C16385 = 1
End Function
Public Function R0C1() As Double
  R0C1 = 1
End Function
Public Function R1C0() As Double
  R1C0 = 1
End Function
Public Function RC() As Double
  RC = 1
End Function
Public Function R1C1X() As Double
  R1C1X = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for out-of-range R1C1 or non-cell forms", got)
	}
}

func TestUdfCellReferenceSkipsPrivateFriendAndNonFunctionProcedures(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Explicit
Private Function A1() As Double
  A1 = 1
End Function
Public Sub A2()
End Sub
Public Property Get A3() As Double
  A3 = 1
End Property
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for Private members or non-Function procedures", got)
	}
	projectFindings := runUdfCellReferenceProjectAnalysis(t, []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source:     []byte("Friend Function A1() As Double\n  A1 = 1\nEnd Function\n"),
	}})
	if got := findingsByCode(projectFindings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for Friend members", got)
	}
}

func TestUdfCellReferenceSkipsOptionPrivateModule(t *testing.T) {
	findings := runUdfCellReferenceAnalysis(t, map[string]string{"Main.bas": `Option Private Module
Public Function A1() As Double
  A1 = 1
End Function
`})
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none inside Option Private Module", got)
	}
}

func TestUdfCellReferenceReportsInMemoryStandardModule(t *testing.T) {
	findings := runUdfCellReferenceProjectAnalysis(t, []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source:     []byte("Option Explicit\nPublic Function ABC123() As Double\n  ABC123 = 42\nEnd Function\n"),
	}})
	if got := findingsByCode(findings, "VBA270"); len(got) != 1 {
		t.Fatalf("VBA270 findings = %+v, want one finding through AnalyzeProject", got)
	}
}

func TestUdfCellReferenceSkipsNonStandardModuleKinds(t *testing.T) {
	source := []byte("Public Function A1() As Double\n  A1 = 1\nEnd Function\n")
	findings := runUdfCellReferenceProjectAnalysis(t, []sourceproject.SourceFile{
		{Path: "virtual/Worker.cls", ModuleKind: sourceproject.ModuleKindClass, Source: source},
		{Path: "virtual/Sheet1.cls", ModuleKind: sourceproject.ModuleKindDocument, Source: source},
		{Path: "virtual/UserForm1.frm", ModuleKind: sourceproject.ModuleKindForm, Source: source},
	})
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none for class/document/form modules", got)
	}
}

func TestUdfCellReferenceDisabledByDefault(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function A1() As Double
  A1 = 1
End Function
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA270"); len(got) != 0 {
		t.Fatalf("VBA270 findings = %+v, want none while the opt-in rule is disabled", got)
	}
}
