package analyze

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func vba242Config() config.Config {
	cfg := config.Default()
	cfg.Analyze.DetectExpensiveFullRangeOperations = true
	return cfg
}

func TestVBA242DetectsFullRangeSinksAndLoopSeverity(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet, rng As Range)
    Dim i As Long
    Columns("A:A").Formula = "=1"
    Rows("1:1").Calculate
    Cells.Interior.Color = 1
    ws.UsedRange.Replace What:="x", Replacement:="y"
    rng.EntireRow.ClearFormats
    rng.EntireColumn.ClearFormats
    ws.Columns(1).Sort
    ws.Sort.SetRange ws.Columns("A:A")
    Range("A:C").Find What:="x"
    For i = 1 To 3
        ws.Columns("A:A").Formula = "=2"
    Next i
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: vba242Config()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA242")
	if len(got) != 10 {
		t.Fatalf("VBA242 findings = %+v, want 10 sinks", got)
	}
	for _, finding := range got {
		if finding.Line == 16 {
			if finding.Severity != "warning" {
				t.Fatalf("loop finding = %+v, want warning", finding)
			}
			continue
		}
		if finding.Severity != "information" {
			t.Fatalf("non-loop finding = %+v, want information", finding)
		}
	}
}

func TestVBA242AcceptsBoundedRangesAndIgnoresCommentsAndStrings(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet, rng As Range, n As Long)
    Range("A1:B10").Formula = "=1"
    ws.Range(Cells(1, 1), Cells(n, 2)).Formula = "=2"
    ws.Columns(1).Resize(n).Formula = "=3"
    Intersect(ws.Columns(1), ws.Range("A1:A10")).Formula = "=4"
    Dim text As String
	text = "Rows(""1:1"").Formula = ""=x"""
    ' ws.Columns("A:A").Formula = "=x"
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: vba242Config()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA242"); len(got) != 0 {
		t.Fatalf("bounded/comment/string operations produced VBA242: %+v", got)
	}
}

func TestVBA242RealtimeMatchesBatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := []byte(`Option Explicit
Public Sub Run(ws As Worksheet)
    Dim i As Long
    For i = 1 To 100
        ws.Rows("1:1").Find What:="x"
    Next i
End Sub
`)
	writeModule(t, dir, "Main.bas", string(source))
	batch, err := (Analyzer{RootDir: dir, Config: vba242Config()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	realtime, err := SourceRealtimeFindings(dir, path, vba242Config(), source)
	if err != nil {
		t.Fatal(err)
	}
	batchFindings := findingsByCode(batch, "VBA242")
	realtimeFindings := findingsByCode(realtime, "VBA242")
	if len(batchFindings) != 1 || len(realtimeFindings) != 1 {
		t.Fatalf("batch/realtime VBA242 = %+v / %+v, want one each", batchFindings, realtimeFindings)
	}
	if batchFindings[0].Severity != "warning" || realtimeFindings[0].Severity != "warning" {
		t.Fatalf("batch/realtime severity = %q / %q, want warning", batchFindings[0].Severity, realtimeFindings[0].Severity)
	}
}

func TestVBA242CanBeDisabledAndSuppressed(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    ' xlflow:disable-next-line VBA242
    Columns("A:A").Formula = "=1"
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	cfg := vba242Config()
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA242"); len(got) != 0 {
		t.Fatalf("suppressed VBA242 finding = %+v", got)
	}
	cfg.Analyze.DetectExpensiveFullRangeOperations = false
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA242"); len(got) != 0 {
		t.Fatalf("disabled VBA242 finding = %+v", got)
	}
}

func TestVBA242ShapeHelpers(t *testing.T) {
	t.Parallel()
	for _, test := range []struct {
		arg  string
		kind string
	}{
		{`"A:A"`, "column"}, {`"A:C"`, "column"}, {`"1:1"`, "row"}, {`"1:3"`, "row"},
	} {
		if got := vba242RangeKind(test.arg); got != test.kind {
			t.Fatalf("vba242RangeKind(%q) = %q, want %q", test.arg, got, test.kind)
		}
	}
	if !vba242FullColumnArgument(`"A:A"`) || !vba242FullColumnArgument("1") {
		t.Fatal("full-column argument helper rejected literal forms")
	}
	if !vba242FullRowArgument(`"1:1"`) || !vba242FullRowArgument("1") {
		t.Fatal("full-row argument helper rejected literal forms")
	}
	if vba242FullRangeArgument(`"A1:B10"`) {
		t.Fatal("bounded A1 range classified as full range")
	}
	if !strings.Contains(vba242Code(`Rows("1:1").Calculate ' comment`), "Calculate") {
		t.Fatal("comment sanitizer removed executable code")
	}
}

func TestVBA242DetectsParenthesizedSinksAndLoopHeader(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet, sortObject As Sort)
    Dim n As Long
    ws.Columns("A:A").Find("x")
    ws.Columns("A:A").Replace("x", "y")
    sortObject.SetRange(ws.Columns("A:A"))
    Do While ws.Rows("1:1").Find("x") Is Nothing
        Exit Do
    Loop
    While ws.Columns("A:A").Find("x") Is Nothing
        n = n + 1
    Wend
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: vba242Config()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA242")
	if len(got) != 5 {
		t.Fatalf("parenthesized/header VBA242 findings = %+v, want 5", got)
	}
	warnings := 0
	for _, finding := range got {
		if finding.Severity == "warning" {
			warnings++
			if !strings.Contains(finding.Message, "reachable loop") {
				t.Fatalf("loop header finding = %+v, want loop context", finding)
			}
		} else if finding.Severity != "information" {
			t.Fatalf("sink finding = %+v, want information", finding)
		}
	}
	if warnings != 2 {
		t.Fatalf("parenthesized/header warnings = %d, want Do While and While header warnings: %+v", warnings, got)
	}
}

func TestVBA242BoundsAreShapeLocalAndMaxCellsAndShadowedBuiltins(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet)
    ws.Columns("A:A").Formula = "=1": Set ws = ws.Resize(2, 2)
    Intersect(ws.Columns("A:A"), ws.Range("A1:A10")).Formula = "=2"
    Range(Cells(1, 1), Cells(1048576, 16384)).Formula = "=3"
    Dim Columns As Variant
    Columns("A:A") = 1
End Sub

Public Function Rows(value As Variant) As Variant
    Rows = value
End Function
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: vba242Config()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA242")
	if len(got) != 2 {
		t.Fatalf("shape-local/max-cells/shadow VBA242 findings = %+v, want 2", got)
	}
	for _, finding := range got {
		if finding.Line != 5 && finding.Line != 7 {
			t.Fatalf("unexpected shape-local/max-cells finding = %+v", finding)
		}
	}
}

func TestVBA242DoesNotAttributeUnrelatedSinks(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet)
    Columns("A:A").Select: ws.Calculate
    ws.Find("x", Rows("1:1"))
    ws.UsedRange.Formula = "=1"
    UsedRange.Calculate
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: vba242Config()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA242")
	if len(got) != 2 {
		t.Fatalf("unrelated sink VBA242 findings = %+v, want explicit and unqualified UsedRange only", got)
	}
	for _, finding := range got {
		if finding.Line != 7 && finding.Line != 8 {
			t.Fatalf("unexpected UsedRange finding = %+v", finding)
		}
	}
}

func TestVBA242HandlesNoArgumentRowsColumnsAndContinuations(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(ws As Worksheet)
    Columns.Formula = "=1"
    Rows.Calculate
    ws.Columns("A:A") _
        .Find("x")
    Columns("A:A").Select: Rem Rows("1:1").Formula = "=ignored"
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: vba242Config()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA242")
	if len(got) != 3 {
		t.Fatalf("no-argument/continuation VBA242 findings = %+v, want three", got)
	}
	for _, finding := range got {
		if finding.Line == 9 {
			t.Fatalf("inline Rem comment produced VBA242 finding = %+v", finding)
		}
	}
}
