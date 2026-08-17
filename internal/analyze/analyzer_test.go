package analyze

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/typedb"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestSourceRealtimeRuleIDsMatchRegistry(t *testing.T) {
	t.Parallel()
	var registryIDs []string
	for _, rule := range staticrules.ByFamily(staticrules.FamilyAnalyze) {
		if rule.Realtime {
			registryIDs = append(registryIDs, rule.ID)
		}
	}
	if !reflect.DeepEqual(sourceRealtimeRuleIDs, registryIDs) {
		t.Fatalf("source realtime implementations = %v, registry = %v", sourceRealtimeRuleIDs, registryIDs)
	}
	for _, id := range sourceRealtimeRuleIDs {
		metadata, _ := staticrules.Lookup(id)
		if metadata.Configurable {
			if _, ok := config.AnalyzeRuleEnabled(config.Default().Analyze, id); !ok {
				t.Fatalf("source realtime rule %s has no config adapter", id)
			}
		}
	}
}

func TestProcedureEffectIdentityCanonicalizesPath(t *testing.T) {
	t.Parallel()
	document := procedureir.DocumentIR{
		Path:       filepath.Join("modules", "..", "modules", "Main.bas"),
		ModuleName: "Main",
		ModuleKind: "standard",
	}
	symbol := procedureir.ProcedureSymbol{
		Name:             "Run",
		QualifiedName:    "Main.Run",
		Kind:             procedureir.ProcedureSub,
		Visibility:       "public",
		DeclarationRange: vbaast.Range{StartLine: 3},
	}

	got := procedureEffectIdentity(document, symbol)
	want := filepath.ToSlash(filepath.Join("modules", "Main.bas"))
	if got.File != want {
		t.Fatalf("canonical file = %q, want %q", got.File, want)
	}
}

func TestVBA225DetectsIndexedCellReadsWritesAndFormatting(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim rng As Range
  Dim i As Long
  Set rng = Range("A1:C100")
  For i = 1 To 100
    Debug.Print rng.Cells(i, 1).Value2
    rng.Cells(i, 2).Value = i
    rng.Cells(i, 3).Formula = "=1"
    rng.Cells(i, 1).Font.Bold = True
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA225")
	if len(got) != 1 {
		t.Fatalf("VBA225 findings = %+v, want one aggregated loop finding", got)
	}
	if got[0].Severity != "warning" || !strings.Contains(got[0].Message, "read") || !strings.Contains(got[0].Message, "write") || !strings.Contains(got[0].Message, "formatting") || !strings.Contains(got[0].Reason, "COM") {
		t.Fatalf("unexpected VBA225 finding: %+v", got[0])
	}
}

func TestVBA225NestedSeverityAndSmallLoopExemption(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  Dim j As Long
  For i = 1 To 100
    For j = 1 To 100
      Cells(i, j).Value2 = j
    Next j
  Next i
  For i = 1 To 1 + 2
    Cells(i, 1).Value2 = i
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA225")
	if len(got) != 1 || got[0].Severity != "warning" || !strings.Contains(got[0].Message, "Nested loop depth: 2") {
		t.Fatalf("nested VBA225 findings = %+v, want one depth-2 warning", got)
	}
}

func TestVBA225UsesNearestNonSmallLoop(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  Dim j As Long
  For i = 1 To 100
    For j = 1 To 2
      Cells(i, j).Value2 = j
    Next j
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA225")
	if len(got) != 1 || got[0].Line != 5 || got[0].Severity != "warning" {
		t.Fatalf("small inner loop should roll up to outer loop: %+v", got)
	}
}

func TestVBA225IgnoresStringAndCommentText(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  For i = 1 To 100
    Debug.Print "Cells(i, 1).Value2"
    ' Range("A1").Value2 is only a comment
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA225"); len(got) != 0 {
		t.Fatalf("string/comment Excel text should not produce VBA225: %+v", got)
	}
}

func TestVBA225IgnoresArrayNamedCellsAndHelperPropagation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Helpers.bas", `Option Explicit
Private Function JoinCells(ByVal inputText As String) As String
  Dim cells As Variant
  Dim index As Long
  Dim result As String
  cells = Split(inputText, "|")
  For index = LBound(cells) To UBound(cells)
    result = result & CStr(cells(index))
  Next index
  JoinCells = result
End Function
`)
	writeModule(t, dir, "Bindings.bas", `Option Explicit
Public cells As Range
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim index As Long
  Dim result As String
  For index = 1 To 100
    result = JoinCells(CStr(index))
  Next index
End Sub

Public Sub ReadProjectRangeNamedCells()
  Dim index As Long
  Dim result As Variant
  For index = 1 To 100
    result = cells(index, 1).Value2
  Next index
End Sub

Public Sub ReadRangeNamedCells()
  Dim cells As Range
  Dim index As Long
  Dim result As Variant
  Set cells = Range("A1:A100")
  For index = 1 To 100
    result = cells(index, 1).Value2
  Next index
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA225")
	if len(got) != 2 {
		t.Fatalf("array named cells should not produce direct or propagated VBA225, while Range bindings remain eligible: %+v", got)
	}
	procedures := map[string]bool{}
	for _, finding := range got {
		procedures[finding.Procedure] = true
	}
	if !procedures["ReadRangeNamedCells"] || !procedures["ReadProjectRangeNamedCells"] {
		t.Fatalf("Range bindings should remain eligible across the current and another module: %+v", got)
	}

	variantDir := t.TempDir()
	writeModule(t, variantDir, "Bindings.bas", `Option Explicit
Public cells As Variant
`)
	writeModule(t, variantDir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim index As Long
  Dim result As Variant
  For index = 1 To 100
    result = cells(index, 1).Value2
  Next index
End Sub
`)
	variantFindings, err := (Analyzer{RootDir: variantDir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(variantFindings, "VBA225"); len(got) != 0 {
		t.Fatalf("project Variant cells should not produce VBA225: %+v", got)
	}

	realtimePath := filepath.Join(dir, "src", "modules", "Realtime.bas")
	realtimeSource := []byte(`Option Explicit
Public Sub Run()
  Dim cells As Variant
  Dim index As Long
  cells = Split("a|b", "|")
  For index = LBound(cells) To UBound(cells)
    Debug.Print CStr(cells(index))
  Next index
End Sub
`)
	realtime, err := SourceRealtimeFindings(dir, realtimePath, config.Default(), realtimeSource)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA225"); len(got) != 0 {
		t.Fatalf("realtime array named cells should not produce VBA225: %+v", got)
	}

	projectRealtimePath := filepath.Join(dir, "src", "modules", "RealtimeProject.bas")
	projectRealtimeSource := []byte(`Option Explicit
Public Sub Run()
  Dim index As Long
  Dim result As Variant
  For index = 1 To 100
    result = cells(index, 1).Value2
  Next index
End Sub
`)
	projectRealtime, err := SourceRealtimeFindings(dir, projectRealtimePath, config.Default(), projectRealtimeSource)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(projectRealtime, "VBA225"); len(got) != 1 {
		t.Fatalf("realtime project Range binding should produce VBA225: %+v", got)
	}

	variantRealtimePath := filepath.Join(variantDir, "src", "modules", "Realtime.bas")
	variantRealtimeSource := []byte(`Option Explicit
Public Sub Run()
  Dim index As Long
  Dim result As Variant
  For index = 1 To 100
    result = cells(index, 1).Value2
  Next index
End Sub
`)
	variantRealtime, err := SourceRealtimeFindings(variantDir, variantRealtimePath, config.Default(), variantRealtimeSource)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(variantRealtime, "VBA225"); len(got) != 0 {
		t.Fatalf("realtime project Variant cells should not produce VBA225: %+v", got)
	}
}

func TestVBA225SupportsForEachOffsetWorksheetFunctionsAndBorders(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim rng As Range
  Dim cell As Range
  Set rng = Range("A1:Z100")
  For Each cell In rng.Cells
    cell.Value2 = cell.Offset(0, 1).Value2
    cell.NumberFormat = "0"
    cell.Borders.LineStyle = 1
    cell.Interior.Color = 16777215
    Debug.Print Application.WorksheetFunction.Sum(cell)
  Next cell
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA225")
	if len(got) != 1 || !strings.Contains(got[0].Message, "read") || !strings.Contains(got[0].Message, "write") || !strings.Contains(got[0].Message, "formatting") || !strings.Contains(got[0].Message, "worksheet function") {
		t.Fatalf("For Each VBA225 findings = %+v, want all access categories", got)
	}
	if !strings.Contains(got[0].Message, "color") {
		t.Fatalf("For Each VBA225 finding = %+v, want Interior.Color member", got[0])
	}
}

func TestVBA225SupportsDoAndWhileLoops(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  i = 1
  Do While i <= 100
    Cells(i, 1).Value2 = i
    i = i + 1
  Loop
  While i <= 200
    Cells(i, 1).Formula = "=1"
    i = i + 1
  Wend
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA225"); len(got) != 2 {
		t.Fatalf("Do/While VBA225 findings = %+v, want two loop findings", got)
	}
}

func TestVBA225BulkOperationsAndSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim rng As Range
  Dim values As Variant
  Dim i As Long
  Set rng = Range("A1:A100")
  values = rng.Value2
  rng.Value2 = values
  rng.Font.Bold = True
  For i = 1 To 100
    values = Range("A1:A100").Value2
    Range("B1:B100").Value2 = values
    Range("A1:A100").NumberFormat = "0"
  Next i
  ' xlflow:disable-next-line VBA225
  For i = 1 To 100
    Debug.Print rng.Cells(i, 1).Value2
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA225"); len(got) != 0 {
		t.Fatalf("suppressed VBA225 findings = %+v", got)
	}
	realtime, err := SourceRealtimeFindings(dir, path, config.Default(), []byte(`Option Explicit
Public Sub Run()
  Dim rng As Range
  Dim i As Long
  Set rng = Range("A1:A100")
  For i = 1 To 100
    Debug.Print rng.Cells(i, 1).Value2
  Next i
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA225"); len(got) != 1 {
		t.Fatalf("realtime VBA225 findings = %+v, want one", got)
	}
}

func TestVBA225TracksUniqueAndTransitiveHelpers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Helpers.bas", `Option Explicit
Public Sub ReadCell(ByVal rng As Range, ByVal i As Long)
  Debug.Print rng.Cells(i, 1).Value2
End Sub

Public Sub ReadCellThroughHelper(ByVal rng As Range, ByVal i As Long)
  ReadCell rng, i
End Sub
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim rng As Range
  Dim i As Long
  Set rng = Range("A1:A100")
  For i = 1 To 100
    ReadCellThroughHelper rng, i
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA225")
	if len(got) != 1 || !strings.Contains(got[0].Message, "ReadCellThroughHelper") || !strings.Contains(got[0].Message, "read") || !strings.Contains(got[0].Message, "value2") {
		t.Fatalf("helper VBA225 findings = %+v, want one transitive helper finding", got)
	}

	realtimePath := filepath.Join(dir, "src", "modules", "Realtime.bas")
	realtimeSource := []byte(`Option Explicit
Private Sub ReadLocal(ByVal rng As Range, ByVal i As Long)
  Debug.Print rng.Cells(i, 1).Value2
End Sub

Public Sub Run()
  Dim rng As Range
  Dim i As Long
  Set rng = Range("A1:A100")
  For i = 1 To 100
    ReadLocal rng, i
  Next i
End Sub
`)
	if err := os.MkdirAll(filepath.Dir(realtimePath), 0o755); err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, realtimePath, config.Default(), realtimeSource)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA225"); len(got) != 1 || !strings.Contains(got[0].Message, "ReadLocal") {
		t.Fatalf("realtime helper VBA225 findings = %+v, want one local helper finding", got)
	}
}

func TestVBA225SkipsAmbiguousHelpers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "First.bas", `Option Explicit
Public Sub ReadCell(ByVal rng As Range, ByVal i As Long)
  Debug.Print rng.Cells(i, 1).Value2
End Sub
`)
	writeModule(t, dir, "Second.bas", `Option Explicit
Public Sub ReadCell(ByVal rng As Range, ByVal i As Long)
  Debug.Print rng.Cells(i, 1).Value2
End Sub
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim rng As Range
  Dim i As Long
  Set rng = Range("A1:A100")
  For i = 1 To 100
    ReadCell rng, i
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA225"); len(got) != 0 {
		t.Fatalf("ambiguous helper VBA225 findings = %+v, want none", got)
	}
}

func TestVBA238DetectsLoopInvariantExcelObjectResolution(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  For i = 1 To 100
    ThisWorkbook.Worksheets("Data").Cells(i, 1).Value2 = i
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA238")
	if len(got) != 1 {
		t.Fatalf("VBA238 findings = %+v, want one", got)
	}
	if got[0].Line != 5 || !strings.Contains(got[0].Message, `ThisWorkbook.Worksheets("Data")`) || !strings.Contains(got[0].Suggestion, "Dim cachedWorksheet As Worksheet") || !strings.Contains(got[0].Suggestion, "Set cachedWorksheet") {
		t.Fatalf("unexpected VBA238 finding: %+v", got[0])
	}
}

func TestVBA238PreservesCRLFMultilineRange(t *testing.T) {
	t.Parallel()
	expression := procedureir.Expression{Range: vbaast.Range{StartLine: 4, StartColumn: 3}}
	loopInvariantExpandRange(&expression, "Worksheets( _\r\n  \"Data\"\r\n)")
	if expression.Range.EndLine != 6 || expression.Range.EndColumn != 2 {
		t.Fatalf("VBA238 CRLF range = %+v, want end line 6 column 2", expression.Range)
	}
}

func TestVBA238SkipsLoopDependentAndTrivialLocalAccess(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("Data")
  For i = 1 To 100
    ws.ListObjects("Sales").ListRows(i).Range.Value2 = i
    ThisWorkbook.Worksheets(CStr(i)).Cells(1, 1).Value2 = i
    ThisWorkbook.Worksheets(OtherSheet).Cells(1, 1).Value2 = i
    Debug.Print ws
  Next i
End Sub
Public Sub Other()
  Const OtherSheet As String = "Data"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA238")
	if len(got) != 1 || !strings.Contains(got[0].Message, `ws.ListObjects("Sales")`) {
		t.Fatalf("unexpected dependent/local VBA238 findings: %+v", got)
	}
}

func TestVBA238DetectsNamedRangePivotAndChartSelectors(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("Data")
  For i = 1 To 100
    ThisWorkbook.Names("SalesTotal").RefersToRange.Value2 = i
    ThisWorkbook.Worksheets("Data").Range("A1", "B2").Value2 = i
    Workbooks("Book.xlsx").Worksheets("Data").Range("A1").Value2 = i
    ws.PivotTables("Summary").PivotFields("Amount").DataRange.Value2 = i
    ws.ChartObjects("Trend").Chart.ChartTitle.Text = CStr(i)
    ws.Charts("Trend2").ChartTitle.Text = CStr(i)
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA238")
	if len(got) != 6 {
		t.Fatalf("VBA238 selector findings = %+v, want two range, named range, pivot field, and chart selectors", got)
	}
	joined := make([]string, 0, len(got))
	for _, finding := range got {
		joined = append(joined, finding.Message)
	}
	rangeSuggestions := make([]string, 0, 2)
	for _, finding := range got {
		if strings.Contains(finding.Message, "Range(") {
			rangeSuggestions = append(rangeSuggestions, finding.Suggestion)
		}
	}
	if len(rangeSuggestions) != 2 || rangeSuggestions[0] == rangeSuggestions[1] {
		t.Fatalf("VBA238 cache names should be unique: %+v", rangeSuggestions)
	}
	for _, want := range []string{"Workbooks(\"Book.xlsx\")", "Names(\"SalesTotal\")", "Range(\"A1\", \"B2\")", "PivotFields(\"Amount\")", "ChartObjects(\"Trend\")", "Charts(\"Trend2\")"} {
		found := false
		for _, message := range joined {
			if strings.Contains(message, want) {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("VBA238 selector findings missing %q: %+v", want, got)
		}
	}
}

func TestVBA238NormalizesMultilineAndSupportsWithNestedLoops(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  Dim j As Long
  For i = 1 To 100
    With ThisWorkbook.Worksheets( _
        "Data" _
      )
      For j = 1 To 100
        .ListObjects("Sales").ListRows(j).Range.Value2 = i
      Next j
    End With
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA238")
	if len(got) != 1 {
		t.Fatalf("VBA238 nested/With findings = %+v, want the maximal table resolution", got)
	}
	sawListObject := false
	for _, finding := range got {
		if finding.Severity != "warning" {
			t.Fatalf("unexpected nested VBA238 finding: %+v", finding)
		}
		sawListObject = sawListObject || strings.Contains(finding.Message, "ListObjects")
	}
	if !sawListObject {
		t.Fatalf("nested/With VBA238 finding missing table chain: %+v", got)
	}
}

func TestVBA238DeduplicatesEquivalentMultilineChains(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim i As Long
  For i = 1 To 100
    Debug.Print ThisWorkbook.Worksheets("Data").Range("A1").Value2
    Debug.Print ThisWorkbook.Worksheets( _
      "Data" _
    ).Range("A1").Value2
  Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA238")
	if len(got) != 1 {
		t.Fatalf("VBA238 normalized findings = %+v, want one maximal range chain", got)
	}
	if !strings.Contains(got[0].Message, `Range("A1")`) {
		t.Fatalf("VBA238 normalized chain missing maximal range: %+v", got)
	}
}

func TestVBA238HandlesForEachDoAndWhileBoundaries(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ForEachHeader()
  Dim row As ListRow
  For Each row In ThisWorkbook.Worksheets("Data").ListObjects("Sales").ListRows
    Debug.Print row.Index
  Next row
End Sub

Public Sub ForEachNested()
  Dim row As ListRow
  Dim i As Long
  For i = 1 To 100
    For Each row In ThisWorkbook.Worksheets("Data").ListObjects("Sales").ListRows
      Debug.Print row.Index
    Next row
  Next i
End Sub

Public Sub DoAndWhile()
  Dim i As Long
  Do While i < 100
    ThisWorkbook.Worksheets(CStr(i)).Cells(1, 1).Value2 = i
    ThisWorkbook.Worksheets("Data").Cells(i, 1).Value2 = i
    i = i + 1
  Loop
  i = 0
  While i < 100
    ThisWorkbook.Worksheets(CStr(i)).Cells(1, 1).Value2 = i
    ThisWorkbook.Worksheets("Data").Cells(i, 1).Value2 = i
    i = i + 1
  Wend
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA238")
	if len(got) != 3 {
		t.Fatalf("VBA238 loop boundary findings = %+v, want nested For Each and two invariant Do/While chains", got)
	}
	for _, finding := range got {
		if strings.Contains(finding.Message, `Worksheets(CStr(i))`) {
			t.Fatalf("loop-dependent selector was reported: %+v", finding)
		}
	}
}

func TestVBA238CanBeDisabledAndSuppressed(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte(`Option Explicit
Public Sub Run()
  Dim i As Long
  For i = 1 To 100
    ' xlflow:disable-next-line VBA238
    ThisWorkbook.Worksheets("Data").Cells(i, 1).Value2 = i
  Next i
End Sub
`)
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, source, 0o644); err != nil {
		t.Fatal(err)
	}
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA238"); len(got) != 0 {
		t.Fatalf("suppressed VBA238 findings = %+v", got)
	}
	realtimeSource := []byte(`Option Explicit
Public Sub Run()
  Dim i As Long
  For i = 1 To 100
    ThisWorkbook.Worksheets("Data").Cells(i, 1).Value2 = i
  Next i
End Sub
`)
	realtime, err := SourceRealtimeFindings(dir, path, config.Default(), realtimeSource)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA238"); len(got) != 1 {
		t.Fatalf("realtime VBA238 findings = %+v, want one", got)
	}
	cfg := config.Default()
	cfg.Analyze.DetectLoopInvariantExcelObjectResolution = false
	if err := os.WriteFile(path, realtimeSource, 0o644); err != nil {
		t.Fatal(err)
	}
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA238"); len(got) != 0 {
		t.Fatalf("disabled VBA238 findings = %+v", got)
	}
}

func TestVBA238RealtimeCoversDocumentedSelectors(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Dim i As Long
  Dim j As Long
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("Data")
  For i = 1 To 100
    Workbooks("Book.xlsx").Worksheets("Data").Range("A1").Value2 = i
    ThisWorkbook.Names("SalesTotal").RefersToRange.Value2 = i
    ws.ListObjects("Sales").ListRows(i).Range.Value2 = i
    ws.PivotTables("Summary").PivotFields("Amount").DataRange.Value2 = i
    ws.ChartObjects("Trend").Chart.ChartTitle.Text = CStr(i)
    ws.Charts("Trend2").ChartTitle.Text = CStr(i)
    ThisWorkbook.Worksheets(CStr(i)).Cells(1, 1).Value2 = i
    With ThisWorkbook.Worksheets( _
        "Data" _
      )
      For j = 1 To 100
        .ListObjects("Sales").ListRows(j).Range.Value2 = i
      Next j
    End With
  Next i
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	findings, err := SourceRealtimeFindings(dir, path, config.Default(), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA238")
	if len(got) != 7 {
		t.Fatalf("realtime VBA238 documented selector findings = %+v, want seven", got)
	}
	for _, want := range []string{"Workbooks(\"Book.xlsx\")", "Names(\"SalesTotal\")", "ListObjects(\"Sales\")", "PivotFields(\"Amount\")", "ChartObjects(\"Trend\")", "Charts(\"Trend2\")"} {
		found := false
		for _, finding := range got {
			if strings.Contains(finding.Message, want) {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("realtime VBA238 findings missing %q: %+v", want, got)
		}
	}
	for _, finding := range got {
		if strings.Contains(finding.Message, `Worksheets(CStr(i))`) {
			t.Fatalf("realtime loop-dependent selector was reported: %+v", finding)
		}
	}
}

func TestAnalyzerDetectsProcedureLocalResourceLeaks(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Resources.bas", `Option Explicit

Public Sub SafeWorkbook(ByVal path As String)
  Dim wb As Workbook
  On Error GoTo Cleanup
  Set wb = Workbooks.Open(path)
  If wb.ReadOnly Then GoTo Cleanup
Cleanup:
  wb.Close SaveChanges:=False
End Sub

Public Sub WorkbookLeak(ByVal path As String)
  Dim wb As Workbook
  Set wb = Application.Workbooks.Open(path)
  Exit Sub
End Sub

Public Function TransferWorkbook(ByVal path As String) As Workbook
  Dim wb As Workbook
  On Error GoTo Cleanup
  Set wb = Workbooks.Open(path)
  Set TransferWorkbook = wb
  Exit Function
Cleanup:
  wb.Close SaveChanges:=False
End Function

Public Sub BorrowedWorkbook(ByVal wb As Workbook)
  Debug.Print wb.Name
End Sub

Public Sub SafeFile(ByVal path As String)
  Dim handle As Integer
  Dim aliasHandle As Integer
  On Error GoTo Cleanup
  handle = FreeFile
  Open path For Output As #handle
  aliasHandle = handle
  Close #aliasHandle
  Exit Sub
Cleanup:
  Close #handle
End Sub

Public Sub SafeCloseAll(ByVal path As String)
  Dim handle As Integer
  handle = FreeFile
  Open path For Output As #handle
  Close
End Sub

Public Sub FileLeak(ByVal path As String)
  Dim handle As Integer
  handle = FreeFile
  Open path For Output As #handle
  Exit Sub
End Sub

Public Sub NestedBranchLeak(ByVal path As String, ByVal enabled As Boolean)
  Dim wb As Workbook
  If enabled Then
    Set wb = Workbooks.Open(path)
    If wb.ReadOnly Then Exit Sub
  End If
End Sub

Public Sub ReassignedWorkbookLeak(ByVal firstPath As String, ByVal secondPath As String)
  Dim wb As Workbook
  Set wb = Workbooks.Open(firstPath)
  Set wb = Workbooks.Open(secondPath)
  wb.Close SaveChanges:=False
End Sub

Public Sub SuppressedWorkbookLeak(ByVal path As String)
  Dim wb As Workbook
  ' xlflow:disable-next-line VBA219
  Set wb = Workbooks.Open(path)
End Sub

Public Sub UnreachableWorkbookOpen(ByVal path As String)
  Exit Sub
  Dim wb As Workbook
  Set wb = Workbooks.Open(path)
End Sub

Public Sub PreOpenFileAlias(ByVal path As String)
  Dim handle As Integer
  Dim aliasHandle As Integer
  On Error GoTo Cleanup
  handle = FreeFile
  aliasHandle = handle
  Open path For Output As #handle
  Close #aliasHandle
  Exit Sub
Cleanup:
  Close #handle
End Sub

Public Function FailedTransfer(ByVal path As String) As Workbook
  Dim wb As Workbook
  Set wb = Workbooks.Open(path)
  Set FailedTransfer = wb
  Err.Raise 5
End Function

Public Function OverwrittenTransfer(ByVal path As String) As Workbook
  Dim wb As Workbook
  On Error GoTo Cleanup
  Set wb = Workbooks.Open(path)
  Set OverwrittenTransfer = wb
  Set OverwrittenTransfer = Nothing
  Exit Function
Cleanup:
  wb.Close SaveChanges:=False
End Function

Public Sub ExceptionalAssignmentKeepsOwner(ByVal path As String)
  Dim wb As Workbook
  Dim aliasWb As Workbook
  On Error GoTo Cleanup
  Set wb = Workbooks.Open(path)
  Set aliasWb = wb
  Set wb = GetOtherWorkbook()
  aliasWb.Close SaveChanges:=False
  Exit Sub
Cleanup:
  wb.Close SaveChanges:=False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA219")
	if len(got) != 6 {
		t.Fatalf("VBA219 findings = %+v, want workbook, file, nested-branch, reassignment, failed-transfer, and overwritten-transfer leaks", got)
	}
	want := map[string]int{"WorkbookLeak": 14, "FileLeak": 55, "NestedBranchLeak": 62, "ReassignedWorkbookLeak": 69, "FailedTransfer": 101, "OverwrittenTransfer": 109}
	for _, finding := range got {
		line, ok := want[finding.Procedure]
		if !ok || finding.Line != line || !strings.Contains(finding.Reason, "without a matching Close") || !strings.Contains(finding.Suggestion, "cleanup path") {
			t.Fatalf("unexpected VBA219 finding: %+v", finding)
		}
	}

	source, err := os.ReadFile(filepath.Join(dir, "src", "modules", "Resources.bas"))
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Resources.bas"), config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA219"); len(got) != 6 {
		t.Fatalf("realtime VBA219 findings = %+v, want six", got)
	}

	cfg := config.Default()
	cfg.Analyze.DetectResourceLeaks = false
	findings, err = Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA219"); len(got) != 0 {
		t.Fatalf("disabled VBA219 should not report in batch: %+v", got)
	}
	realtime, err = SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Resources.bas"), cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA219"); len(got) != 0 {
		t.Fatalf("disabled VBA219 should not report in realtime: %+v", got)
	}
}

func TestVBA219RequiresWorkbookTypedAcquisitionOwner(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "ResourceFactory.cls", `Option Explicit

Public Function CustomReturn(ByVal path As String) As ResourceFactory
  Set CustomReturn = Workbooks.Open(path)
End Function

Public Function WorkbookReturn(ByVal path As String) As Workbook
  Set WorkbookReturn = Workbooks.Open(path)
End Function

Public Function QualifiedWorkbookReturn(ByVal path As String) As Excel.Workbook
  Set QualifiedWorkbookReturn = Workbooks.Open(path)
End Function

Public Sub WorkbookLeak(ByVal path As String)
  Dim wb As Workbook
  Set wb = Workbooks.Open(path)
End Sub

Public Sub QualifiedWorkbookLeak(ByVal path As String)
  Dim wb As Excel.Workbook
  Set wb = Workbooks.Open(path)
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA219")
	if len(got) != 2 || got[0].Procedure != "WorkbookLeak" || got[0].Line != 17 || got[1].Procedure != "QualifiedWorkbookLeak" || got[1].Line != 22 {
		t.Fatalf("VBA219 findings = %+v, want only Workbook-typed local leaks", got)
	}

	source, err := os.ReadFile(filepath.Join(dir, "src", "classes", "ResourceFactory.cls"))
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "classes", "ResourceFactory.cls"), config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA219"); len(got) != 2 || got[0].Procedure != "WorkbookLeak" || got[1].Procedure != "QualifiedWorkbookLeak" {
		t.Fatalf("realtime VBA219 findings = %+v, want the same Workbook-typed local leaks", got)
	}
}

func TestResourceLeakDoesNotTrustRecoveredRelease(t *testing.T) {
	t.Parallel()
	acquisition := procedureir.Statement{ID: 1, Kind: procedureir.StatementSet, Text: `Set wb = Workbooks.Open(path)`}
	recoveredClose := procedureir.Statement{ID: 2, Kind: procedureir.StatementCall, Text: `wb.Close`, Recovered: true}
	graph := vbacfg.Graph{
		Blocks: []vbacfg.Block{
			{ID: 1, Kind: vbacfg.BlockEntry},
			{ID: 2, Kind: vbacfg.BlockStatement, StatementID: 1, Statement: &acquisition},
			{ID: 3, Kind: vbacfg.BlockStatement, StatementID: 2, Statement: &recoveredClose},
			{ID: 4, Kind: vbacfg.BlockNormalExit},
		},
		Edges: []vbacfg.Edge{
			{From: 1, To: 2, Class: vbacfg.EdgeNormal},
			{From: 2, To: 3, Class: vbacfg.EdgeNormal},
			{From: 3, To: 4, Class: vbacfg.EdgeNormal},
		},
		Entry: 1, NormalExit: 4,
	}
	witness, leaked := resourceLeakWitness(sourceProcedure{Graph: &graph}, resourceAcquisition{StatementID: 1, Kind: resourceWorkbook, Owner: "wb"})
	if !leaked || witness.Kind != "normal exit" {
		t.Fatalf("recovered Close must not prove release: witness=%+v leaked=%v", witness, leaked)
	}
}

func TestAnalyzerVBA201IgnoresProjectFindMethods(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "WorksheetRelationshipMapper.cls", `Attribute VB_Name = "WorksheetRelationshipMapper"
Option Explicit
Public Function Find(sheetName As String, key) As Object
  Set Find = New JsonValue
End Function
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Private m_wrm As WorksheetRelationshipMapper
Public Sub Run()
  Dim result As Object
  Set result = m_wrm.Find("Sheet", "key")
  Debug.Print result.ToJson()
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA201"); len(got) != 0 {
		t.Fatalf("project Find method should not trigger VBA201: %+v", got)
	}
	source, err := os.ReadFile(filepath.Join(dir, "src", "modules", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA201"); len(got) != 0 {
		t.Fatalf("realtime project Find method should not trigger VBA201: %+v", got)
	}
}

func TestAnalyzerVBA201FindsImplicitWithRangeReceiver(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim found As Range
  With Range("A1:A10")
    Set found = .Find(What:="x")
    Debug.Print found.Value
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA201"); len(got) != 1 || got[0].Line != 6 {
		t.Fatalf("implicit With Range.Find should trigger VBA201 on the dereference line: %+v", got)
	}

	source, err := os.ReadFile(filepath.Join(dir, "src", "modules", "Main.bas"))
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA201"); len(got) != 1 || got[0].Line != 6 {
		t.Fatalf("realtime implicit With Range.Find should trigger VBA201 on the dereference line: %+v", got)
	}
}

func TestSourceRealtimeFindingsParsedMatchesSource(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerDoesNotInferAmbiguousObjectFunctionReturn(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "First.bas", `Option Explicit
Public Function FindRange() As Range
  Set FindRange = Sheet1.Range("A1")
End Function
`)
	writeModule(t, dir, "Second.bas", `Option Explicit
Public Function FindRange() As Range
  Set FindRange = Sheet1.Range("B1")
End Function
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim result As Range
  result = FindRange()
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findings {
		if finding.Code == "VBA102" && finding.Line == 4 {
			t.Fatalf("ambiguous object return must not produce VBA102: %+v", findings)
		}
	}
}

func TestAnalyzerFindsMissingSetInObjectReturningFunction(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	var parseErr *ParseError
	if !errors.As(err, &parseErr) {
		t.Fatalf("expected ParseError, got %T: %v", err, err)
	}
	if parseErr.Path != filepath.Join(dir, "src", "modules", "Main.bas") || !parseErr.HasError && !parseErr.HasMissing {
		t.Fatalf("unexpected ParseError: %+v", parseErr)
	}
}

func TestAnalyzerContinuesAfterDeclarationKeywordRecovery(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
	  Dim b() As Byte, ReDim b(10)
	  Debug.Print b(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatalf("declaration keyword recovery should not abort analyzer: %v", err)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 1 || got[0].Line != 4 || got[0].RuntimeError == nil || got[0].RuntimeError.Kind != "array_unallocated" {
		t.Fatalf("declaration recovery should retain only the actual deterministic array access finding, got %+v", got)
	}
}

func TestAnalyzerContinuesAfterIdentifierTypeCharacterRecovery(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim text$, whole%, longValue&, singleValue!, doubleValue#, money@, longLong^
End Sub
`)

	if _, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run(); err != nil {
		t.Fatalf("legal identifier type-character recovery should not abort analyzer: %v", err)
	}
}

func TestAnalyzerRejectsUnrelatedSameLineDeclarationRecovery(t *testing.T) {
	t.Parallel()

	for _, test := range []struct {
		name string
		line string
	}{
		{name: "declaration keyword", line: "Dim x As Double, Dim i As Long: value ="},
		{name: "type character", line: "Dim value!: other ="},
	} {
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			dir := t.TempDir()
			writeModule(t, dir, "Main.bas", "Option Explicit\nPublic Sub Run()\n  "+test.line+"\nEnd Sub\n")

			_, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
			var parseErr *ParseError
			if !errors.As(err, &parseErr) {
				t.Fatalf("unrelated same-line recovery should remain a ParseError, got %T: %v", err, err)
			}
		})
	}
}

func TestAnalyzerFindsWorksheetMemberAssignedOnVariable(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerDetectsAmbiguousExcelScopeRoots(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub InteractiveEntry()
  ActiveWorkbook.Save
  Application.ActiveSheet.Range("A1").Value = 1
  With ActiveWorkbook
    .Save
  End With
  ActiveCell.Value = "x"
  Selection.Clear
  Range("A1").Value = 1
  Cells(1, 1).Value = 2
  Rows(1).Delete
  Columns(1).Hidden = True
  Worksheets("Data").Range("A1").Value = 1
  Sheets(1).Name = "Data"
  Workbooks(1).Close
  Application.Windows(1).Visible = False
  Workbooks.Open "C:\\temp\\Book.xlsm"
  Application.Workbooks.Open _
    Filename:="C:\\temp\\Other.xlsm"
  Set wb = Workbooks.Open( _
    Filename:="C:\\temp\\Captured.xlsm" _
  )
  Set wb = Application.Workbooks.Open("C:\\temp\\CapturedToo.xlsm")
  Dim another As Workbook: Set another = Workbooks.Open("C:\\temp\\CapturedAfterDeclaration.xlsm")
  Set first = Workbooks.Open("C:\\temp\\First.xlsm"): Workbooks.Open "C:\\temp\\Second.xlsm"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	if len(got) != 16 {
		t.Fatalf("VBA205 findings = %+v, want 16", got)
	}
	for line, expected := range map[int]struct{ root, suggestion string }{
		3: {"ActiveWorkbook", "explicit Workbook"}, 4: {"ActiveSheet", "explicit Workbook"}, 5: {"ActiveWorkbook", "explicit Workbook"}, 8: {"ActiveCell", "explicit Workbook"}, 9: {"Selection", "explicit Workbook"},
		10: {"Range", "explicit Worksheet or Range"}, 11: {"Cells", "explicit Worksheet or Range"}, 12: {"Rows", "explicit Worksheet or Range"}, 13: {"Columns", "explicit Worksheet or Range"}, 14: {"Worksheets", "ThisWorkbook.Worksheets"},
		15: {"Sheets", "ThisWorkbook.Sheets"}, 16: {"Workbooks(1)", "by name"}, 17: {"Windows(1)", "by name"}, 18: {"Workbooks.Open", "Set wb = Workbooks.Open"}, 19: {"Workbooks.Open", "Set wb = Workbooks.Open"}, 26: {"Workbooks.Open", "Set wb = Workbooks.Open"},
	} {
		matching := false
		for _, finding := range got {
			if finding.Line == line && strings.Contains(finding.Message, expected.root) && finding.Reason != "" && strings.Contains(finding.Suggestion, expected.suggestion) {
				matching = true
				break
			}
		}
		if !matching {
			t.Fatalf("expected VBA205 mentioning %q with suggestion %q on line %d: %+v", expected.root, expected.suggestion, line, got)
		}
	}
	for _, finding := range got {
		if finding.Line == 21 || finding.Line == 24 || finding.Line == 25 {
			t.Fatalf("captured Workbooks.Open should not be reported: %+v", finding)
		}
	}
}

func TestAnalyzerVBA205RespectsProcedureIRSymbolResolution(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Range(ByVal start As Long, ByVal count As Long) As Long
  Range = start + count
End Function

Public Function Rows() As Long
  Rows = 1
End Function

Public Sub UseLocalNames(ByVal cells As Collection, ByVal rows As Collection)
  Dim columns As Collection
  Set columns = New Collection
  cells.Add 1
  rows.Add 1
  columns.Add 1
End Sub
`)
	writeClass(t, dir, "Table.cls", `Attribute VB_Name = "Table"
Public Property Get Columns() As Collection
Attribute Columns.VB_Description = "Returns local columns."
  Set Columns = New Collection
End Property

Public Property Get Rows() As Collection
  Set Rows = New Collection
End Property

Public Sub AddRow(ByVal cells As Collection)
  Rows.Add cells
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA205"); len(got) != 0 {
		t.Fatalf("resolved project and local symbols should not report VBA205: %+v", got)
	}
}

func TestAnalyzerVBA205RespectsPrivateProcedureVisibility(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function Rows(ByVal index As Long) As Object
  Set Rows = Nothing
End Function

Public Sub Run()
  Rows(1).Value = 1
  Range("A1").Value = 2
End Sub
`)
	writeModule(t, dir, "Helpers.bas", `Option Explicit
Private Function Range(ByVal address As String) As String
  Range = address
End Function
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	if len(got) != 1 || got[0].Line != 8 || !strings.Contains(got[0].Message, "Range") {
		t.Fatalf("private procedure visibility = %+v, want only cross-module Range finding on line 8", got)
	}
}

func TestAnalyzerAcceptsExplicitExcelScopeReferences(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal wb As Workbook, ByVal ws As Worksheet)
  ThisWorkbook.Worksheets("Data").Range("A1").Value = 1
  wb.Worksheets("Data").Range("A1").Value = 2
  ws.Range("A1").Value = 3
  Workbooks("Book.xlsm").Worksheets("Data").Range("A1").Value = 4
  Windows("Book.xlsm").Visible = True
  Set wb = Workbooks.Open("C:\\temp\\Book.xlsm")
  With wb.Worksheets("Data")
    .Range("A2").Value = 5
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA205"); len(got) != 0 {
		t.Fatalf("explicit Excel scopes should not be reported: %+v", got)
	}
}

func TestAnalyzerDetectsAmbiguousExcelScopeRootsInWithHeaders(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  With ActiveSheet
    .Range("A1").Value = 1
  End With
  With Worksheets("Data")
    .Range("A1").Value = 2
  End With
  With Workbooks.Open("C:\\temp\\Book.xlsm")
    .Worksheets(1).Range("A1").Value = 3
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	for line, root := range map[int]string{3: "ActiveSheet", 6: "Worksheets", 9: "Workbooks.Open"} {
		found := false
		for _, finding := range got {
			if finding.Line == line && strings.Contains(finding.Message, root) {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("expected VBA205 mentioning %q on With header line %d: %+v", root, line, got)
		}
	}
}

func TestAnalyzerDetectsApplicationScopedAmbiguousRoots(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Application.Range("A1").Value = 1
  Application.Cells(1, 1).Value = 2
  Application.Rows(1).Delete
  Application.Columns(1).Hidden = True
  Application.Worksheets("Data").Range("A1").Value = 3
  Application.Sheets("Data").Range("A1").Value = 4
  If True Then Set wb = Workbooks.Open("C:\\temp\\Captured.xlsm")
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	for line, root := range map[int]string{3: "Range", 4: "Cells", 5: "Rows", 6: "Columns", 7: "Worksheets", 8: "Sheets"} {
		found := false
		for _, finding := range got {
			if finding.Line == line && strings.Contains(finding.Message, root) {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("expected VBA205 mentioning %q on Application root line %d: %+v", root, line, got)
		}
	}
	if len(got) != 6 {
		t.Fatalf("VBA205 findings = %+v, want six Application-root findings", got)
	}
}

func TestAnalyzerDetectsAllNumericPositionalWorkbookAndWindowIndices(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Workbooks(2).Close
  Application.Windows(3).Visible = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	for line, root := range map[int]string{3: "Workbooks(2)", 4: "Windows(3)"} {
		found := false
		for _, finding := range got {
			if finding.Line == line && strings.Contains(finding.Message, root) {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("expected VBA205 mentioning %q on line %d: %+v", root, line, got)
		}
	}
}

func TestAnalyzerClassifiesEachWorkbooksOpenCaptureIndependently(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal useA As Boolean)
  If useA Then Set wb = Workbooks.Open("A.xlsm") Else Workbooks.Open "B.xlsm"
  Set plain=Workbooks.Open("NoSpace.xlsm")
  Set app=Application.Workbooks.Open("Application.xlsm")
  Set books(1) = Workbooks.Open("Array.xlsm")
  Set holder.Book = Workbooks.Open("Property.xlsm")
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	if len(got) != 1 || got[0].Line != 3 || !strings.Contains(got[0].Message, "Workbooks.Open") {
		t.Fatalf("Workbooks.Open findings = %+v, want only uncaptured Else branch", got)
	}
}

func TestAnalyzerAddInSheetCollectionSuggestsCallerWorkbook(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "AddInEntry.bas", `Option Explicit
Public Sub Run()
  Worksheets("Data").Range("A1").Value = 1
  Sheets("Data").Range("A2").Value = 2
End Sub
`)
	cfg := config.Default()
	cfg.Excel.Path = "build/AddIn.xlam"
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	if len(got) != 2 {
		t.Fatalf("add-in sheet collection findings = %+v, want Worksheets and Sheets guidance", got)
	}
	for _, finding := range got {
		if !strings.Contains(finding.Suggestion, "caller Workbook") || strings.Contains(finding.Suggestion, "ThisWorkbook") {
			t.Fatalf("add-in sheet collection finding = %+v, want caller-workbook guidance", finding)
		}
	}
}

func TestAnalyzerReportsThisWorkbookOnlyForAddInStandardModules(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "AddInEntry.bas", `Option Explicit
Public Sub Run()
  ThisWorkbook.Worksheets("Data").Range("A1").Value = 1
End Sub
`)
	workbookDir := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbookDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbookDir, "ThisWorkbook.cls"), []byte(`Attribute VB_Name = "ThisWorkbook"
Option Explicit
Private Sub Workbook_Open()
  ThisWorkbook.Worksheets("Data").Range("A1").Value = 1
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Excel.Path = "build/AddIn.xlam"
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	if len(got) != 1 || got[0].File != "src/modules/AddInEntry.bas" || got[0].Line != 3 || !strings.Contains(got[0].Message, "ThisWorkbook") || !strings.Contains(got[0].Suggestion, "caller workbook") {
		t.Fatalf("add-in ThisWorkbook findings = %+v, want one standard-module finding", got)
	}
}

func TestAnalyzerReportsActiveSheetDependenciesInPrivateHelpers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  WriteStatus
End Sub

Private Sub WriteStatus()
  Range("A1").Value = "done"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA205")
	if len(got) != 1 || got[0].Procedure != "WriteStatus" || got[0].Line != 7 {
		t.Fatalf("helper scope dependency = %+v, want WriteStatus line 7", got)
	}
	if !strings.Contains(got[0].Suggestion, "Worksheet") {
		t.Fatalf("helper suggestion = %q, want explicit Worksheet guidance", got[0].Suggestion)
	}
}

func TestAnalyzerVBA205IgnoresCommentsAndStringLiterals(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Debug.Print "ActiveWorkbook Worksheets(1) Workbooks.Open"
  ' Selection.Cells(1, 1).Value = Workbooks(1).Name
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA205"); len(got) != 0 {
		t.Fatalf("VBA205 should ignore comments and string literals: %+v", got)
	}
}

func TestAnalyzerHonorsDisabledRuleIDs(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerDoesNotSuppressVBA216(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeWorkbookModule(t, dir, "Sheet1.bas")
	writeWorkbookModule(t, dir, "Sheet2.bas")
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim lastRow As Long
  ' xlflow:disable-next-line VBA216
  lastRow = Sheet1.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row
End Sub
`)

	result, err := Analyzer{RootDir: dir, Config: config.Default()}.RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(result.Findings, "VBA216"); len(got) != 1 {
		t.Fatalf("VBA216 should remain unsuppressed: findings=%+v warnings=%+v", result.Findings, result.Warnings)
	}
	if !hasWarning(result.Warnings, "unsupported_inline_suppression_rule", "VBA216") {
		t.Fatalf("expected unsupported suppression warning, got %+v", result.Warnings)
	}
}

func TestAnalyzerRuntimeRiskRulesAllowGuardedPatterns(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	if got := findingsByCode(findings, "VBA221"); len(got) != 0 {
		t.Fatalf("VBA221 should not report a paired Push/Pop helper: %+v", got)
	}
}

func TestAnalyzerApplicationStateAllowsEitherSameModuleRestoreAlias(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
End Sub

Private Sub PopFastMode()
  Debug.Print "cleanup"
End Sub

Private Sub RestoreFastMode()
  Application.EnableEvents = True
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA203"); len(got) != 0 {
		t.Fatalf("VBA203 should accept either same-module restore alias: %+v", got)
	}
}

func TestAnalyzerApplicationStateRejectsPopThatDisablesEvents(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PushFastMode()
  Application.EnableEvents = False
End Sub

Private Sub PopFastMode()
  Application.EnableEvents = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA203", 3)
}

func TestAnalyzerApplicationStateStillFlagsUnpairedPushPattern(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerApplicationStateChecksEveryConfiguredProperty(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub UnsafeAllProperties()
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  Application.Calculation = xlCalculationManual
  Application.StatusBar = "working"
  Application.Cursor = xlWait
  Application.Interactive = False
  Application.AskToUpdateLinks = False
  Application.AutomationSecurity = msoAutomationSecurityForceDisable
  Application.CutCopyMode = xlCopy
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	if len(got) != 10 {
		t.Fatalf("VBA203 findings = %+v, want one per Application property", got)
	}
	for _, property := range []string{"ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation", "StatusBar", "Cursor", "Interactive", "AskToUpdateLinks", "AutomationSecurity", "CutCopyMode"} {
		found := false
		for _, finding := range got {
			if strings.Contains(finding.Message, "Application."+property) && strings.Contains(finding.Reason, "exit") {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("missing %s all-path finding: %+v", property, got)
		}
	}
}

func TestAnalyzerApplicationStateRecognizesWithApplicationSharedCleanupAndCopies(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub SafeCleanup(ByVal invalidInput As Boolean)
  Dim savedEvents As Boolean
  Dim copiedEvents As Boolean
  Dim savedStatus As Variant
  On Error GoTo Cleanup
  With Application
    savedEvents = .EnableEvents
    copiedEvents = savedEvents
    savedStatus = .StatusBar
    .EnableEvents = False
    .StatusBar = "working"
  End With
  If invalidInput Then GoTo Cleanup
  Debug.Print "work"
Cleanup:
  With Application
    .StatusBar = savedStatus
    .EnableEvents = copiedEvents
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA203"); len(got) != 0 {
		t.Fatalf("shared cleanup and copied saved values should be safe: %+v", got)
	}
}

func TestAnalyzerApplicationStateReportsEarlyExitAndErrorHandlerPaths(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub UnsafeExitSub(ByVal invalidInput As Boolean)
  Dim savedEvents As Boolean
  savedEvents = Application.EnableEvents
  On Error GoTo Handler
  Application.EnableEvents = False
  If invalidInput Then Exit Sub
  Err.Raise 5
Cleanup:
  Application.EnableEvents = savedEvents
  Exit Sub
Handler:
  Exit Sub
End Sub

Public Sub UnsafeErrorHandler()
  Dim savedCursor As Long
  savedCursor = Application.Cursor
  On Error GoTo Handler
  Application.Cursor = xlWait
  Err.Raise 5
  Exit Sub
Handler:
  Exit Sub
End Sub

Public Sub UnsafeNestedBranches(ByVal outer As Boolean, ByVal inner As Boolean)
  Dim savedLinks As Boolean
  savedLinks = Application.AskToUpdateLinks
  If outer Then
    Application.AskToUpdateLinks = False
    If inner Then Exit Sub
  End If
  Application.AskToUpdateLinks = savedLinks
End Sub

Public Function UnsafeExitFunction(ByVal done As Boolean) As Long
  Dim savedAlerts As Boolean
  savedAlerts = Application.DisplayAlerts
  Application.DisplayAlerts = False
  If done Then Exit Function
  Application.DisplayAlerts = savedAlerts
End Function

Public Property Get UnsafeExitProperty(ByVal done As Boolean) As Long
  Dim savedInteractive As Boolean
  savedInteractive = Application.Interactive
  Application.Interactive = False
  If done Then Exit Property
  Application.Interactive = savedInteractive
End Property
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	for _, procedure := range []string{"UnsafeExitSub", "UnsafeErrorHandler", "UnsafeNestedBranches", "UnsafeExitFunction", "UnsafeExitProperty"} {
		found := false
		for _, finding := range got {
			if finding.Procedure == procedure {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("missing %s early-exit finding: %+v", procedure, got)
		}
	}
	var handlerPath bool
	for _, finding := range got {
		if finding.Procedure == "UnsafeErrorHandler" && strings.Contains(finding.Reason, "error-handler path") {
			handlerPath = true
		}
	}
	if !handlerPath {
		t.Fatalf("missing error-handler exit witness: %+v", got)
	}
}

func TestAnalyzerApplicationStateRejectsInvalidOrConditionalSavedValue(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ReassignedSavedValue()
  Dim savedEvents As Boolean
  savedEvents = Application.EnableEvents
  Application.EnableEvents = False
  savedEvents = True
  Application.EnableEvents = savedEvents
End Sub

Public Sub ConditionalSavedValue(ByVal changeState As Boolean)
  Dim savedAlerts As Boolean
  If changeState Then
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.DisplayAlerts = savedAlerts
  End If
  Application.DisplayAlerts = savedAlerts
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	for _, procedure := range []string{"ReassignedSavedValue", "ConditionalSavedValue"} {
		found := false
		for _, finding := range got {
			if finding.Procedure == procedure {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("%s should not prove a restore from an invalid saved value: %+v", procedure, got)
		}
	}
}

func TestAnalyzerApplicationStatePreservesRestoreCoverageAcrossHandlerMerge(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub HandlerMerge()
  On Error GoTo Handler
  Dim savedAlerts As Boolean
  savedAlerts = Application.DisplayAlerts
  Application.DisplayAlerts = False
  Call Work
Cleanup:
  Application.DisplayAlerts = savedAlerts
  Exit Sub
Handler:
  Resume Cleanup
End Sub

Private Sub Work()
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	if len(got) != 1 || got[0].Line != 9 {
		t.Fatalf("handler merge findings = %+v, want only the potentially uninitialized cleanup restore on line 9", got)
	}
}

func TestAnalyzerApplicationStateDoesNotFollowErrRaiseIntoLexicalCleanup(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub ReraiseBeforeCleanup()
  On Error GoTo Handler
  Dim savedAlerts As Boolean
  savedAlerts = Application.DisplayAlerts
  Application.DisplayAlerts = False
  Call Work
  GoTo Cleanup
Handler:
  Application.DisplayAlerts = savedAlerts
  Err.Raise Err.Number
Cleanup:
  Application.DisplayAlerts = savedAlerts
End Sub

Private Sub Work()
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	if len(got) != 1 || got[0].Line != 10 {
		t.Fatalf("reraise cleanup findings = %+v, want only the potentially uninitialized handler restore on line 10", got)
	}
}

func TestAnalyzerApplicationStateRecognizesRepeatedGuardedRestore(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub GuardedRestore(ByVal itemCount As Long)
  Dim savedUpdating As Boolean
  If itemCount > 0 Then
    savedUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
  End If
  Call Work
  If itemCount > 0 Then
    Application.ScreenUpdating = savedUpdating
  End If
End Sub

Private Sub Work()
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	if len(got) != 1 || got[0].Line != 6 {
		t.Fatalf("repeated guard findings = %+v, want only the state change that can fail before restoration on line 6", got)
	}
}

func TestAnalyzerApplicationStateRejectsChangedGuardBindings(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private moduleEnabled As Boolean

Public Sub ReassignedLocalGuard(ByVal itemCount As Long)
  Dim savedUpdating As Boolean
  If itemCount > 0 Then
    savedUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
  End If
  itemCount = 0
  If itemCount > 0 Then
    Application.ScreenUpdating = savedUpdating
  End If
End Sub

Public Sub StableModuleGuard()
  Dim savedUpdating As Boolean
  If moduleEnabled Then
    savedUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
  End If
  If moduleEnabled Then
    Application.ScreenUpdating = savedUpdating
  End If
End Sub

Public Sub CalledModuleGuard()
  Dim savedUpdating As Boolean
  If moduleEnabled Then
    savedUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
  End If
  MutateModuleGuard
  If moduleEnabled Then
    Application.ScreenUpdating = savedUpdating
  End If
End Sub

Private Sub MutateModuleGuard()
  moduleEnabled = False
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	restoreFindings := map[string]bool{}
	for _, finding := range got {
		if finding.Line == 12 || finding.Line == 35 {
			restoreFindings[finding.Procedure] = true
		}
		if finding.Procedure == "StableModuleGuard" && finding.Line == 23 {
			t.Fatalf("stable module guard restore should remain recognized: %+v", got)
		}
	}
	for _, procedure := range []string{"ReassignedLocalGuard", "CalledModuleGuard"} {
		if !restoreFindings[procedure] {
			t.Fatalf("%s should retain its guarded-restore finding: %+v", procedure, got)
		}
	}
}

func TestAnalyzerApplicationStateGuardComparisonPreservesStringLiterals(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub DifferentStringGuards(ByVal value As String)
  Dim savedUpdating As Boolean
  If value = "a b" Then
    savedUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
  End If
  If value = "ab" Then
    Application.ScreenUpdating = savedUpdating
  End If
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA203")
	for _, finding := range got {
		if finding.Line == 9 {
			return
		}
	}
	t.Fatalf("different string-literal guards should retain the restore finding: %+v", got)
}

func TestVBA221ReportsImmediateCallerAndUncertainCalleeOncePerProperty(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Main()
  PrepareExcel
End Sub

Private Sub PrepareExcel()
  Application.EnableEvents = False
  Application.EnableEvents = False
  MaybeRestoreEvents
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA203"); len(got) != 2 {
		t.Fatalf("VBA203 root causes = %+v, want both PrepareExcel assignments", got)
	} else {
		lines := map[int]bool{}
		for _, finding := range got {
			if finding.Procedure != "PrepareExcel" {
				t.Fatalf("VBA203 root cause procedure = %+v, want PrepareExcel", finding)
			}
			lines[finding.Line] = true
		}
		if !lines[7] || !lines[8] {
			t.Fatalf("VBA203 root cause lines = %+v, want 7 and 8", got)
		}
	}
	got := findingsByCode(findings, "VBA221")
	if len(got) != 1 || got[0].Procedure != "Main" || got[0].Line != 3 || !strings.Contains(got[0].Message, "Application.EnableEvents") || !strings.Contains(got[0].Reason, "Main.PrepareExcel") || !strings.Contains(got[0].Reason, "unresolved") {
		t.Fatalf("VBA221 caller context = %+v, want one uncertain Main call finding", got)
	}
}

func TestVBA221DoesNotRepeatTransitiveLeakAtAncestorOrRealtime(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Main()
  Wrapper
End Sub

Private Sub Wrapper()
  PrepareExcel
End Sub

Private Sub PrepareExcel()
  Application.ScreenUpdating = False
End Sub
`
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	writeModule(t, dir, "Main.bas", source)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA221")
	if len(got) != 1 || got[0].Procedure != "Wrapper" || got[0].Line != 7 || !strings.Contains(got[0].Reason, "Main.PrepareExcel") {
		t.Fatalf("VBA221 should report only the immediate caller: %+v", got)
	}
	if realtime, err := SourceRealtimeFindings(dir, path, config.Default(), []byte(source)); err != nil {
		t.Fatal(err)
	} else if got := findingsByCode(realtime, "VBA221"); len(got) != 0 {
		t.Fatalf("batch-only VBA221 must not be returned in realtime: %+v", got)
	}
}

func TestVBA221IgnoresUnreachableCallerPaths(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Main()
  GoTo Done
  PrepareExcel
Done:
End Sub

Private Sub PrepareExcel()
  Application.EnableEvents = False
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA221"); len(got) != 0 {
		t.Fatalf("unreachable call must not report VBA221: %+v", got)
	}
}

func TestVBA221HonorsConfigAndInlineSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Main()
	  PrepareExcel
End Sub

Private Sub PrepareExcel()
  Application.DisplayAlerts = False
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA221"); len(got) != 1 {
		t.Fatalf("default VBA221 should report the direct caller: %+v", got)
	}
	cfg := config.Default()
	cfg.Analyze.DetectApplicationStateCallEffects = false
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA221"); len(got) != 0 {
		t.Fatalf("disabled VBA221 should not report: %+v", got)
	}
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Main()
  ' xlflow:disable-next-line VBA221
  PrepareExcel
End Sub

Private Sub PrepareExcel()
  Application.DisplayAlerts = False
End Sub
`)
	findings, err = (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA221"); len(got) != 0 {
		t.Fatalf("inline suppression should remove VBA221: %+v", got)
	}
}

func TestAnalyzerErrorHandlerFallthroughSuggestsConcreteExitStatement(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerCFGVBA204AllowsQualifiedCleanupLabels(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub AuthCleanup()
  On Error GoTo auth_Cleanup
  Debug.Print "work"
auth_Cleanup:
  Debug.Print "cleanup"
End Sub

Public Sub AutoProxyCleanup()
  On Error GoTo AutoProxy_Cleanup
  Debug.Print "work"
AutoProxy_Cleanup:
  Debug.Print "cleanup"
End Sub

Public Sub CleanupNamedHandler()
  On Error GoTo CleanupErrorHandler
  Debug.Print "work"
CleanupErrorHandler:
  Debug.Print "handled"
End Sub

Public Sub LegacyCleanUp()
  On Error GoTo worker_clean_up
  Debug.Print "work"
worker_clean_up:
  Debug.Print "cleanup"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA204")
	if len(got) != 1 || got[0].Procedure != "CleanupNamedHandler" || got[0].Line != 19 {
		t.Fatalf("VBA204 findings = %+v, want only the non-cleanup suffix control", got)
	}
	modulePath := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(modulePath)
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, modulePath, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	got = findingsByCode(realtime, "VBA204")
	if len(got) != 1 || got[0].Procedure != "CleanupNamedHandler" || got[0].Line != 19 {
		t.Fatalf("realtime VBA204 findings = %+v, want only the non-cleanup suffix control", got)
	}
}

func TestAnalyzerCFGVBA204AllowsSemanticSharedCleanupHandlers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function CloseSharedHandle(ByVal handle As Long) As Boolean
  On Error GoTo ErrorHandling
  If handle <> 0 Then Debug.Print handle
ErrorHandling:
  Close #1
End Function

Public Function FinalizeDimensions(ByRef values As Variant) As Long
  Dim dimension As Long
  Dim tempBound As Long
  On Error GoTo FinalDimension
  For dimension = 1 To 60
    tempBound = LBound(values, dimension)
  Next dimension
FinalDimension:
  FinalizeDimensions = dimension - 1
End Function

Public Function ClosePipe(ByVal fileHandle As Long) As Long
  On Error GoTo utc_ErrorHandling
  Do While fileHandle > 0
    fileHandle = fileHandle - 1
  Loop
utc_ErrorHandling:
  ClosePipe = CLng(pclose(fileHandle))
End Function

Public Function LoopFailure() As Long
  Dim i As Long
  On Error GoTo Handler
  For i = 1 To 2
    Debug.Print i
  Next i
Handler:
  LoopFailure = 0
End Function
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA204")
	if len(got) != 1 || got[0].Procedure != "LoopFailure" || got[0].Line != 35 {
		t.Fatalf("semantic shared cleanup handlers should only retain LoopFailure: %+v", got)
	}

	modulePath := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(modulePath)
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, modulePath, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	got = findingsByCode(realtime, "VBA204")
	if len(got) != 1 || got[0].Procedure != "LoopFailure" || got[0].Line != 35 {
		t.Fatalf("realtime semantic shared cleanup handlers should only retain LoopFailure: %+v", got)
	}
}

func TestAnalyzerCFGVBA204RejectsTextualCleanupLookalikes(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function StringLookalike() As Long
  Dim message As String
  On Error GoTo Handler
  StringLookalike = 1
Handler:
  message = "reset = nothing"
End Function

Public Function CallLookalike() As Long
  On Error GoTo Handler
  CallLookalike = 1
Handler:
  CloseAndPurgeAll
End Function

Public Function AssignmentLookalike() As Long
  Dim handle As Object
  On Error GoTo Handler
  AssignmentLookalike = 1
Handler:
  Set handle = NothingCached
End Function

Public Function ExactCleanup() As Long
  On Error GoTo Handler
  ExactCleanup = 1
Handler:
  Close #1
End Function

Public Function AliasedCleanup(ByVal fileHandle As Long) As Long
  On Error GoTo Handler
  AliasedCleanup = 1
Handler:
  AliasedCleanup = CLng(utc_pclose(fileHandle))
End Function
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA204")
	if len(got) != 3 || got[0].Procedure != "StringLookalike" || got[1].Procedure != "CallLookalike" || got[2].Procedure != "AssignmentLookalike" {
		t.Fatalf("textual cleanup lookalikes should retain VBA204 while exact Close remains safe: %+v", got)
	}

	modulePath := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(modulePath)
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, modulePath, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	got = findingsByCode(realtime, "VBA204")
	if len(got) != 3 || got[0].Procedure != "StringLookalike" || got[1].Procedure != "CallLookalike" || got[2].Procedure != "AssignmentLookalike" {
		t.Fatalf("realtime textual cleanup lookalikes should retain VBA204: %+v", got)
	}
}

func TestAnalyzerCFGVBA204DoesNotTreatDirectErrRaiseAsFallthrough(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub UnconditionalRaise()
  On Error GoTo Handler
  Err.Raise 5
Handler:
  Debug.Print "handled"
End Sub

Public Sub ConditionalRaise(ByVal fail As Boolean)
  On Error GoTo Handler
  If fail Then Err.Raise 5
Handler:
  Debug.Print "handled"
End Sub

Public Sub ConditionalCompilationPath()
  On Error GoTo Handler
#If Mac Then
  Err.Raise 5
#Else
  Debug.Print "work"
#End If
Handler:
  Debug.Print "handled"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA204")
	if len(got) != 2 || got[0].Procedure != "ConditionalRaise" || got[0].Line != 12 ||
		got[1].Procedure != "ConditionalCompilationPath" || got[1].Line != 23 {
		t.Fatalf("VBA204 findings = %+v, want the conditional and alternate-compilation fallthroughs", got)
	}
	modulePath := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(modulePath)
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, modulePath, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	got = findingsByCode(realtime, "VBA204")
	if len(got) != 2 || got[0].Procedure != "ConditionalRaise" || got[0].Line != 12 ||
		got[1].Procedure != "ConditionalCompilationPath" || got[1].Line != 23 {
		t.Fatalf("realtime VBA204 findings = %+v, want the conditional and alternate-compilation fallthroughs", got)
	}
}

func TestAnalyzerVBA214AllowsNarrowCompatibilityProbes(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub DirectProbe()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("Data")
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub
End Sub

Public Sub CheckedProbe()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("Data")
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Public Sub ReplacedByHandler()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("Data")
  On Error GoTo Handler
  Exit Sub
Handler:
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA214"); len(got) != 0 {
		t.Fatalf("narrow probes should not report VBA214: %+v", got)
	}
}

func TestAnalyzerVBA214AllowsErrDescriptionCompatibilityProbe(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub DescriptionProbe()
  Dim ws As Worksheet
  Dim probeError As String
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("Data")
  If Err.Number <> 0 Then
    probeError = Err.Description
    Err.Clear
  End If
  On Error GoTo 0
  If Len(probeError) > 0 Then Exit Sub
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA214"); len(got) != 0 {
		t.Fatalf("Err.Description probe should not report VBA214: %+v", got)
	}
}

func TestAnalyzerVBA214ReportsScopeBoundsAndEarlyExits(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub BroadScope()
  On Error Resume Next
  Debug.Print "one"
  Debug.Print "two"
  On Error GoTo 0
End Sub

Public Sub EarlyExit()
  On Error Resume Next
  Debug.Print "one"
  Exit Sub
End Sub

Public Sub NaturalExit()
  On Error Resume Next
  Debug.Print "one"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 3 {
		t.Fatalf("VBA214 findings = %+v, want three", got)
	}
	wantEnds := map[int]int{3: 6, 10: 12, 16: 18}
	for _, finding := range got {
		if end, ok := wantEnds[finding.Line]; !ok || finding.ScopeEndLine != end {
			t.Fatalf("unexpected VBA214 scope boundary: %+v, want starts/ends %+v", finding, wantEnds)
		}
		if finding.Severity != "warning" || !containsAll(finding.Message, "line "+strconvItoa(finding.Line), "line "+strconvItoa(finding.ScopeEndLine)) {
			t.Fatalf("unexpected VBA214 finding: %+v", finding)
		}
	}
}

func TestAnalyzerVBA214ReportsContinuationAfterObjectProbeBeforeRestore(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub UsesFailedProbe()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("Data")
  Debug.Print ws.Name
  On Error GoTo 0
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 1 || got[0].Line != 4 || got[0].ScopeEndLine != 7 || got[0].Severity != "warning" {
		t.Fatalf("VBA214 object-probe continuation = %+v", got)
	}
}

func TestAnalyzerVBA214DoesNotTreatStringLiteralsAsErrProbes(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error Resume Next
  Debug.Print "Err.Number"
  Debug.Print "two"
  On Error GoTo 0
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 1 || got[0].Line != 3 || got[0].ScopeEndLine != 6 {
		t.Fatalf("VBA214 must not treat string literals as Err probes: %+v", got)
	}
}

func TestAnalyzerVBA214ElevatesResolvedProjectCallsAndWarnsUnresolvedCalls(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Helper()
End Sub

Public Function ValueHelper() As Long
  ValueHelper = 1
End Function

Public Sub LocalCall()
  On Error Resume Next
  Helper
  On Error GoTo 0
End Sub

Public Sub LocalFunctionCall()
  Dim value As Long
  On Error Resume Next
  value = ValueHelper()
  On Error GoTo 0
End Sub

Public Sub UnknownCall()
  On Error Resume Next
  MissingHelper
  On Error GoTo 0
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 3 {
		t.Fatalf("VBA214 findings = %+v, want three", got)
	}
	severityByProcedure := map[string]string{}
	for _, finding := range got {
		severityByProcedure[finding.Procedure] = finding.Severity
	}
	if severityByProcedure["LocalCall"] != "warning" || severityByProcedure["LocalFunctionCall"] != "warning" || severityByProcedure["UnknownCall"] != "warning" {
		t.Fatalf("VBA214 call severities = %+v, want warning-only inference", severityByProcedure)
	}
	if blocking := findingsByCode(BlockingFindings(findings), "VBA214"); len(blocking) != 0 {
		t.Fatalf("VBA214 must not block source preflight: %+v", blocking)
	}
}

func TestAnalyzerVBA214TracksNestedBranchAndHandlerScopes(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub BranchLeak(ByVal stopEarly As Boolean)
  On Error Resume Next
  If stopEarly Then
    Exit Sub
  End If
  Debug.Print "work"
  On Error GoTo 0
End Sub

Public Sub HandlerLeak()
  On Error GoTo Handler
  Debug.Print "work"
  Exit Sub
Handler:
  On Error Resume Next
  Debug.Print "work"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 3 {
		t.Fatalf("VBA214 findings = %+v, want branch restore/exit and handler exit", got)
	}
	boundaries := map[string]map[int]bool{}
	for _, finding := range got {
		if boundaries[finding.Procedure] == nil {
			boundaries[finding.Procedure] = map[int]bool{}
		}
		boundaries[finding.Procedure][finding.ScopeEndLine] = true
	}
	if !boundaries["BranchLeak"][5] || !boundaries["BranchLeak"][8] || !boundaries["HandlerLeak"][18] {
		t.Fatalf("VBA214 scope ends = %+v", boundaries)
	}
}

func TestAnalyzerVBA214DoesNotFollowMergedDisabledErrorMode(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error Resume Next
  If Err.Number <> 0 Then
    On Error GoTo 0
  End If
  Debug.Print "probe"
  On Error GoTo 0
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA214"); len(got) != 0 {
		t.Fatalf("merged disabled-mode error edge must not leak Resume Next scope: %+v", got)
	}
}

func TestAnalyzerVBA214ReportsAllProcedureExitKindsAndHandlerReplacement(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function FunctionExit() As Long
  On Error Resume Next
  Exit Function
End Function

Public Property Get PropertyExit() As Long
  On Error Resume Next
  Exit Property
End Property

Public Sub TerminatesProcess()
  On Error Resume Next
  End
End Sub

Public Sub UnsafeHandlerReplacement()
  On Error Resume Next
  Debug.Print "one"
  Debug.Print "two"
  On Error GoTo Handler
Handler:
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	procedures := map[string]bool{}
	for _, finding := range got {
		procedures[finding.Procedure] = true
	}
	for _, procedure := range []string{"FunctionExit", "PropertyExit", "TerminatesProcess", "UnsafeHandlerReplacement"} {
		if !procedures[procedure] {
			t.Fatalf("missing VBA214 for %s: %+v", procedure, got)
		}
	}
}

func TestAnalyzerVBA214ReportsUnknownGotoExit(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  On Error Resume Next
  GoTo MissingLabel
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 1 || got[0].Line != 3 || got[0].ScopeEndLine != 4 || !strings.Contains(got[0].Reason, "exit before") {
		t.Fatalf("VBA214 unknown goto exit = %+v", got)
	}
}

func TestAnalyzerVBA214ElevatesProjectCallsInControlConditions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function ProjectPredicate() As Boolean
  ProjectPredicate = True
End Function

Public Sub Run()
  On Error Resume Next
  If ProjectPredicate() Then
    Debug.Print "work"
  End If
  On Error GoTo 0
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 1 || got[0].Severity != "warning" || got[0].Procedure != "Run" {
		t.Fatalf("project call in control condition should remain a VBA214 warning: %+v", got)
	}
}

func TestAnalyzerVBA214HonorsInlineAndConfigSuppressionIndependentlyOfVB004(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub InlineSuppressed()
  ' xlflow:disable-next-line VBA214
  On Error Resume Next
  Debug.Print "one"
  Exit Sub
End Sub

Public Sub Reported()
  On Error Resume Next
  Debug.Print "one"
  Exit Sub
End Sub
`)
	cfg := config.Default()
	cfg.Lint.ForbidOnErrorResumeNext = false
	issues, err := (lint.Linter{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB004" {
			t.Fatalf("VB004 should be disabled independently: %+v", issues)
		}
	}
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA214")
	if len(got) != 1 || got[0].Procedure != "Reported" {
		t.Fatalf("VBA214 should remain independent from VB004 and honor inline suppression: %+v", got)
	}

	cfg.Analyze.DetectLeakedOnErrorResumeNextScopes = false
	findings, err = Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA214"); len(got) != 0 {
		t.Fatalf("disabled VBA214 should not report: %+v", got)
	}
}

func TestSourceRealtimeAnalysisExcludesBatchOnlyVBA214(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	source := []byte(`Option Explicit
Public Sub Run()
  On Error Resume Next
  Debug.Print "one"
  Debug.Print "two"
  On Error GoTo 0
End Sub
`)
	findings, err := SourceRealtimeFindings(dir, path, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA214"); len(got) != 0 {
		t.Fatalf("batch-only VBA214 should not appear in realtime findings: %+v", got)
	}
}

func TestVBA215MatchesBatchAndRealtimeAnalysisAndHonorsSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
    Dim rng As Range
    rng.Find "missing"
    rng.Replace What:="old", Replacement:="new", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, MatchByte:=False
    ' xlflow:disable-next-line VBA215
    rng.Replace "old", "new"
End Sub
`)
	cfg := config.Default()
	batch, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "Main.bas"), cfg, []byte(`Option Explicit
Public Sub Run()
    Dim rng As Range
    rng.Find "missing"
    rng.Replace What:="old", Replacement:="new", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, MatchByte:=False
    ' xlflow:disable-next-line VBA215
    rng.Replace "old", "new"
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	for name, findings := range map[string][]Finding{"batch": batch, "realtime": realtime} {
		got := findingsByCode(findings, "VBA215")
		if len(got) != 1 || got[0].Line != 4 || !strings.Contains(got[0].Message, "LookIn, LookAt, SearchOrder, MatchByte") {
			t.Fatalf("%s VBA215 findings = %+v", name, got)
		}
	}

	cfg.Analyze.DetectStatefulExcelCallArguments = false
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA215"); len(got) != 0 {
		t.Fatalf("disabled VBA215 should not report: %+v", got)
	}
}

func TestBatchTypedExcelRulesReusePreparedParsedDocument(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	source := []byte(`Option Explicit
Public Sub Run()
    Dim rng As Range
    Dim result As Variant
    rng.Find "missing"
    rng.SpecialCells xlCellTypeVisible
    result = Application.Match("missing", rng, 0)
    Debug.Print result
End Sub
`)
	parsed, err := vbaast.ParseDocument(path, source)
	if err != nil {
		t.Fatal(err)
	}
	document := intel.Document{Path: path, Source: string(source)}
	document.Snapshot = intel.NewAnalysisSnapshotWithParsedDocument(document, parsed)
	defer document.Snapshot.Retire()
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}
	file := parsedFile{
		Path:          path,
		Lines:         normalizedSourceLines(string(source)),
		Source:        source,
		IntelDocument: document,
	}
	analysis := Analyzer{RootDir: dir, Config: config.Default(), typeDB: db}
	if got := analysis.statefulExcelCallArgumentFindings(file); len(got) == 0 {
		t.Fatal("VBA215 fixture should exercise the shared typed document index")
	}
	if got := analysis.excelAPIFailureContractFindings(file); len(got) == 0 {
		t.Fatal("VBA218 fixture should exercise the shared typed document index")
	}
	if got := document.Snapshot.ParseCount(); got != 1 {
		t.Fatalf("shared VBA215/VBA218 snapshot parse count = %d, want 1", got)
	}
}

func TestVBA218MatchesBatchAndRealtimeAnalysisAndHonorsFailureContracts(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	source := `Option Explicit
Public Sub Run()
    Dim rng As Range
    Dim result As Variant
    rng.SpecialCells xlCellTypeVisible
    On Error GoTo Handler
    WorksheetFunction.Match "key", rng, 0
    On Error GoTo 0
    On Error Resume Next
    WorksheetFunction.VLookup "key", rng, 2, False
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    result = Application.Match("key", rng, 0)
    If IsError(result) Then Exit Sub
    Debug.Print result
    Debug.Print Application.VLookup("key", rng, 2, False)
Handler:
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	cfg := config.Default()
	batch, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for name, findings := range map[string][]Finding{"batch": batch, "realtime": realtime} {
		got := findingsByCode(findings, "VBA218")
		if len(got) != 2 || got[0].Line != 5 || got[1].Line != 16 {
			t.Fatalf("%s VBA218 findings = %+v", name, got)
		}
		if !strings.Contains(got[0].Message, "may raise") || !strings.Contains(got[1].Message, "Variant/Error") {
			t.Fatalf("%s VBA218 contract messages = %+v", name, got)
		}
	}

	cfg.Analyze.DetectExcelAPIFailureContracts = false
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA218"); len(got) != 0 {
		t.Fatalf("disabled VBA218 should not report: %+v", got)
	}
}

func TestVBA218RecognizesLocalIsErrorGuardAndCVErrWrapper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function IsLookupError(ByVal value As Variant) As Boolean
    IsLookupError = IsError(value)
End Function

Private Function TryVisible(ByVal rng As Range) As Variant
    On Error GoTo Missing
    TryVisible = rng.SpecialCells(xlCellTypeVisible)
    Exit Function
Missing:
    TryVisible = CVErr(xlErrNA)
End Function

Public Sub Run()
    Dim rng As Range
    Dim result As Variant
    result = Application.XLookup("key", rng, rng)
    If IsLookupError(result) Then Exit Sub
    Debug.Print result
    result = TryVisible(rng)
    If IsError(result) Then Exit Sub
    Debug.Print result
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA218"); len(got) != 0 {
		t.Fatalf("guarded local wrapper calls should not report VBA218: %+v", got)
	}
}

func TestVBA218UsesUniqueCrossModuleIsErrorGuardOnlyInBatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Guards.bas", `Option Explicit
Public Function IsLookupFailure(ByVal value As Variant) As Boolean
    IsLookupFailure = IsError(value)
End Function
`)
	source := `Option Explicit
Public Sub Run()
    Dim rng As Range
    Dim result As Variant
    result = Application.Match("key", rng, 0)
    If IsLookupFailure(result) Then Exit Sub
    Debug.Print result
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	batch, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(batch, "VBA218"); len(got) != 0 {
		t.Fatalf("uniquely resolved cross-module guard should suppress batch VBA218: %+v", got)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "Main.bas"), config.Default(), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA218"); len(got) != 1 || got[0].Line != 5 {
		t.Fatalf("realtime analysis must remain document-local: %+v", got)
	}
}

func TestVBA218RejectsNonDominatingAndUninspectedGuards(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function IsLookupError(ByVal value As Variant) As Boolean
    IsLookupError = IsError(value)
End Function

Public Sub Run(ByVal retry As Boolean)
    Dim rng As Range
    Dim result As Variant
    On Error Resume Next
    rng.SpecialCells xlCellTypeVisible
    Err.Clear
    On Error GoTo 0
    result = Application.Match("key", rng, 0)
    If IsError(result) And retry Then Exit Sub
    Debug.Print result
    result = Application.XLookup("key", rng, rng)
    If checker.IsLookupError(result) Then Exit Sub
    Debug.Print result
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA218")
	if len(got) != 3 || got[0].Line != 10 || got[1].Line != 13 || got[2].Line != 16 {
		t.Fatalf("non-dominating/unchecked VBA218 findings = %+v", got)
	}
}

func TestVBA218TracksCVErrWrapperResultAtCaller(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function TryVisible(ByVal rng As Range) As Variant
    On Error GoTo Missing
    TryVisible = rng.SpecialCells(xlCellTypeVisible)
    Exit Function
Missing:
    TryVisible = CVErr(xlErrNA)
End Function

Public Sub Run()
    Dim rng As Range
    TryVisible(rng)
    Debug.Print TryVisible(rng)
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA218")
	if len(got) != 1 || got[0].Line != 13 || got[0].Column != 17 || got[0].EndLine != 13 || !strings.Contains(got[0].Message, "TryVisible may return") {
		t.Fatalf("unchecked CVErr wrapper result should report VBA218: %+v", got)
	}
}

func TestVBA218SuppressesDisabledCVErrWrapperFindingsInBatchAndRealtime(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	source := `Option Explicit
Private Function TryVisible(ByVal rng As Range) As Variant
  On Error GoTo Missing
  TryVisible = rng.SpecialCells(xlCellTypeVisible)
  Exit Function
Missing:
  TryVisible = CVErr(xlErrNA)
End Function
Public Sub Run(ByVal rng As Range)
  Debug.Print TryVisible(rng)
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	cfg := config.Default()
	cfg.Analyze.DetectExcelAPIFailureContracts = false
	batch, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, path, cfg, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	for name, findings := range map[string][]Finding{"batch": batch, "realtime": realtime} {
		if got := findingsByCode(findings, "VBA218"); len(got) != 0 {
			t.Fatalf("disabled VBA218 should suppress %s CVErr wrapper findings: %+v", name, got)
		}
	}
}

func TestVBA218RecognizesLocalGuardAliasInRealtimeWrapperChecks(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	source := `Option Explicit
Private Function IsLookupError(ByVal value As Variant) As Boolean
  IsLookupError = IsError(value)
End Function
Private Function TryVisible(ByVal rng As Range) As Variant
  On Error GoTo Missing
  TryVisible = rng.SpecialCells(xlCellTypeVisible)
  Exit Function
Missing:
  TryVisible = CVErr(xlErrNA)
End Function
Public Sub Run(ByVal rng As Range)
  Dim result As Variant
  result = TryVisible(rng)
  If IsLookupError(result) Then Exit Sub
  Debug.Print result
End Sub
`
	findings, err := SourceRealtimeFindings(dir, path, config.Default(), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA218"); len(got) != 0 {
		t.Fatalf("local IsError guard alias should suppress realtime wrapper finding: %+v", got)
	}
}

func TestVBA218StopsVariantTrackingAfterImmediateReassignment(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal values As Range)
  Dim result As Variant
  result = Application.Match("needle", values, 0)
  result = 1
  Debug.Print result
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA218"); len(got) != 0 {
		t.Fatalf("reassigned Variant/Error result should be dead: %+v", got)
	}
}

func TestVBA218RecognizesMultilineIsErrorExitGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal values As Range)
  Dim result As Variant
  result = Application.Match("needle", values, 0)
  If IsError(result) Then
    Exit Sub
  End If
  Debug.Print result
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA218"); len(got) != 0 {
		t.Fatalf("multiline IsError exit guard should suppress VBA218: %+v", got)
	}
}

func TestVBA218TracksUniqueCrossModuleCVErrWrapperOnlyInBatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Guards.bas", `Option Explicit
Public Function TryVisible(ByVal rng As Range) As Variant
    On Error GoTo Missing
    TryVisible = rng.SpecialCells(xlCellTypeVisible)
    Exit Function
Missing:
    TryVisible = CVErr(xlErrNA)
End Function
`)
	source := `Option Explicit
Public Sub Run()
    Dim rng As Range
    Debug.Print TryVisible(rng)
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	batch, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(batch, "VBA218"); len(got) != 1 || got[0].Line != 4 {
		t.Fatalf("batch analysis should resolve cross-module CVErr wrapper: %+v", got)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "Main.bas"), config.Default(), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA218"); len(got) != 0 {
		t.Fatalf("realtime analysis must not resolve cross-module CVErr wrappers: %+v", got)
	}
}

func TestWorksheetRootFindingsAppearInRealtimeAnalysis(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte(`Option Explicit
Public Sub Run()
  Dim inputSheet As Worksheet
  Dim outputSheet As Worksheet
  Dim lastRow As Long
  Set inputSheet = ThisWorkbook.Worksheets("Input")
  Set outputSheet = ThisWorkbook.Worksheets("Output")
  lastRow = inputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row
  lastRow = Cells(Rows.Count, 1).End(xlDown).Row
End Sub
`)

	findings, err := SourceRealtimeFindings(dir, path, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA216"); len(got) != 1 || got[0].Line != 8 || got[0].Severity != "error" {
		t.Fatalf("realtime VBA216 findings = %+v", got)
	}
	if got := findingsByCode(findings, "VBA217"); len(got) != 2 || got[0].Line != 9 || got[1].Line != 9 {
		t.Fatalf("realtime VBA217 findings = %+v", got)
	}
}

func TestWorksheetRootRealtimeAnalysisUsesWorkbookCodenames(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeWorkbookModule(t, dir, "InputSheet.bas")
	writeWorkbookModule(t, dir, "OutputSheet.bas")
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte(`Option Explicit
Public Sub Run()
  Dim lastRow As Long
  lastRow = InputSheet.Cells(OutputSheet.Rows.Count, 1).End(xlUp).Row
End Sub
`)

	findings, err := SourceRealtimeFindings(dir, path, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA216"); len(got) != 1 || got[0].Line != 4 {
		t.Fatalf("realtime workbook-codename VBA216 findings = %+v", got)
	}
}

func TestWorksheetRootRealtimeAnalysisHandlesContinuationsAndWithHeaders(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeWorkbookModule(t, dir, "InputSheet.bas")
	writeWorkbookModule(t, dir, "OutputSheet.bas")
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte(`Option Explicit
Public Sub Run()
  Dim lastRow As Long
  lastRow = InputSheet.Cells( _
      OutputSheet.Rows.Count, 1).End(xlUp).Row
  With InputSheet.Range( _
      OutputSheet.Cells(1, 1), OutputSheet.Cells(2, 1))
  End With
End Sub
`)

	findings, err := SourceRealtimeFindings(dir, path, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA216")
	if len(got) != 2 || got[0].Line != 4 || got[1].Line != 6 {
		t.Fatalf("realtime VBA216 continuation and With-header findings = %+v", got)
	}
}

func TestAnalyzerChecksObjectUseOnSetAssignmentRHS(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerVBA202CarriesAllocatedObjectThroughLineContinuationCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Append(ByVal target As Collection, ByVal value As String)
  target.Add value
End Sub

Public Sub Run()
  Dim values As Collection
  Set values = New Collection
  Append values, _
    "value"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA202"); len(got) != 0 {
		t.Fatalf("an allocated object should remain non-Nothing through a line-continuation call: %+v", got)
	}
}

func TestAnalyzerRuntimeRiskRules(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub NeedsLong(ByRef value As Long)
End Sub
Public Function MissingReturn() As Range
End Function
Public Sub Run()
  Dim dict As Dictionary
  Dim text As String
  Set dict = CreateObject("Scripting.Dictionary")
  NeedsLong text
  Debug.Print dict("missing")
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDictionaryCollectionGuard = true
	cfg.Analyze.DetectFunctionReturnPath = true

	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA228", 10)
	assertFinding(t, findings, "VBA207", 11)
	assertFinding(t, findings, "VBA210", 4)
}

func TestAnalyzerVBA210UsesCFGReturnPaths(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function MissingReturn() As Long
End Function

Public Property Get MissingProperty() As Long
End Property

Public Property Get EarlyProperty() As Long
  Exit Property
  EarlyProperty = 1
End Property

Public Property Get LetValue() As Long
  Let LetValue = 1
End Property

Public Function EarlyExit() As Long
  Exit Function
  EarlyExit = 1
End Function

Public Function BranchMissing(ByVal flag As Boolean) As Long
  If flag Then
    BranchMissing = 1
  Else
    Debug.Print "not assigned"
  End If
End Function

Public Function BranchSafe(ByVal flag As Boolean) As Long
  If flag Then
    BranchSafe = 1
  Else
    BranchSafe = 0
  End If
End Function

Public Function DominatingAssignment(ByVal flag As Boolean) As Long
  DominatingAssignment = 0
  If flag Then Exit Function
  DominatingAssignment = 1
End Function

Public Function SharedCleanup(ByVal flag As Boolean) As Long
  On Error GoTo Handler
  If flag Then GoTo Cleanup
  Debug.Print "work"
Cleanup:
  SharedCleanup = 0
  Exit Function
Handler:
  GoTo Cleanup
End Function

Public Function UnsafeHandler() As Long
  On Error GoTo Handler
  Err.Raise 5
  UnsafeHandler = 1
  Exit Function
Handler:
  Exit Function
End Function

Public Function SafeHandler() As Long
  On Error GoTo Handler
  Err.Raise 5
  SafeHandler = 1
  Exit Function
Handler:
  SafeHandler = 0
End Function

Public Function ReraisedHandler() As Long
  On Error GoTo Handler
  ReraisedHandler = 1
  Exit Function
Handler:
  Err.Raise 5
End Function

Public Function RaiseOnly() As Long
  Err.Raise 5
End Function

Public Function CallRaiseOnly() As Long
  Call Err.Raise(5)
End Function

Public Function ObjectValueAssignment() As Range
  ObjectValueAssignment = Sheet1.Range("A1")
End Function

Public Function ObjectSetAssignment() As Range
  Set ObjectSetAssignment = Sheet1.Range("A1")
End Function

Public Function ValueSetAssignment() As Long
  Set ValueSetAssignment = CreateObject("Scripting.Dictionary")
End Function

Public Function CommentOnly() As Long
  ' CommentOnly = 1
  Debug.Print "CommentOnly = 1"
End Function

Public Property Let Writable(ByVal value As Long)
End Property

Public Property Set ObjectWritable(ByVal value As Object)
End Property

Public Sub NoReturn()
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectFunctionReturnPath = true

	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA210")
	want := map[string]bool{
		"MissingReturn":      true,
		"MissingProperty":    true,
		"EarlyProperty":      true,
		"EarlyExit":          true,
		"BranchMissing":      true,
		"UnsafeHandler":      true,
		"ValueSetAssignment": true,
		"CommentOnly":        true,
	}
	for _, finding := range got {
		delete(want, finding.Procedure)
	}
	if len(want) != 0 {
		t.Fatalf("missing VBA210 procedures: %v; findings=%+v", want, got)
	}
	for _, procedure := range []string{"BranchSafe", "DominatingAssignment", "SharedCleanup", "SafeHandler", "ReraisedHandler", "RaiseOnly", "CallRaiseOnly", "ObjectSetAssignment", "LetValue", "Writable", "ObjectWritable", "NoReturn"} {
		for _, finding := range got {
			if finding.Procedure == procedure {
				t.Fatalf("unexpected VBA210 for %s: %+v", procedure, finding)
			}
		}
	}
	var unsafeHandler Finding
	for _, finding := range got {
		if finding.Procedure == "UnsafeHandler" {
			unsafeHandler = finding
			break
		}
	}
	if unsafeHandler.Procedure != "UnsafeHandler" || !strings.Contains(unsafeHandler.Reason, "error-handler path") {
		t.Fatalf("missing error-handler witness: %+v", got)
	}
	if !strings.Contains(unsafeHandler.Reason, "line ") {
		t.Fatalf("VBA210 reason should identify the representative exit: %+v", unsafeHandler)
	}
	for _, finding := range got {
		switch finding.Procedure {
		case "EarlyExit":
			if !strings.Contains(finding.Message, "Exit Function") {
				t.Fatalf("missing Exit Function witness: %+v", finding)
			}
		case "EarlyProperty":
			if !strings.Contains(finding.Message, "Exit Property") {
				t.Fatalf("missing Exit Property witness: %+v", finding)
			}
		}
	}
	if gotObject := findingsByCode(findings, "VBA103"); len(gotObject) != 1 || gotObject[0].Procedure != "ObjectValueAssignment" {
		t.Fatalf("expected only the object syntax diagnostic for ObjectValueAssignment: %+v", gotObject)
	}
}

func TestAnalyzerByRefMismatchHandlesLowercaseCallKeyword(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub NeedsLong(ByRef value As Long)
End Sub
Public Sub Run()
  Dim text As String
  call NeedsLong(text)
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertFinding(t, findings, "VBA228", 6)
}

func TestAnalyzerCompileEquivalentByRefMismatchIgnoresVBA206Disable(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub NeedsLong(ByRef value As Long)
End Sub
Public Sub Run()
  Dim text As String
  NeedsLong text
  NeedsLong (text)
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectByRefArgumentMismatch = false
	findings, err := Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA228"); len(got) != 1 || got[0].Line != 6 || got[0].Severity != "error" {
		t.Fatalf("compile-equivalent ByRef finding = %+v", got)
	}
	if got := findingsByCode(findings, "VBA206"); len(got) != 0 {
		t.Fatalf("VBA206 should remain disabled while VBA228 stays active: %+v", got)
	}
}

func TestAnalyzerDoesNotProjectLintOnlyCompileEquivalentFindings(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim value As Long
  Set value = 1
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VB037"); len(got) != 0 {
		t.Fatalf("VB037 is lint/LSP-only and must not be projected by analyze: %+v", got)
	}
}

func TestAnalyzerByRefUsesProjectLocalNamedSignatures(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Receiver.bas", `Option Explicit
Public Sub ReplaceText(ByRef target As String, Optional ByVal suffix As String = "")
End Sub
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim count As Long
  Receiver.ReplaceText target:=count
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA228")
	if len(got) != 1 || got[0].Line != 4 || !strings.Contains(got[0].Message, "requires String") {
		t.Fatalf("named project-local ByRef finding = %+v", got)
	}
}

func TestAnalyzerByRefSkipsAmbiguousCallsAndHonorsSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "First.bas", `Option Explicit
Public Sub TakeText(ByRef target As String)
End Sub
`)
	writeModule(t, dir, "Second.bas", `Option Explicit
Public Sub TakeText(ByRef target As String)
End Sub
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub TakeLong(ByRef target As Long)
End Sub
Public Sub Run()
  Dim value As Long
  Dim text As String
  TakeText value
	  TakeLong text ' xlflow:disable-line VBA206
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA228"); len(got) != 1 {
		t.Fatalf("ambiguous ByRef calls should be skipped but compile-equivalent mismatches remain unsuppressible: %+v", got)
	}
}

func TestAnalyzerArrayComparisonUsesIdentifierBoundaries(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerArrayComparisonFindingsHaveStableOrder(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim leftValues() As Variant
  Dim rightValues() As Variant
  If leftValues = rightValues Then Debug.Print "bad"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA209")
	if len(got) != 2 || got[0].Line != 5 || got[1].Line != 5 {
		t.Fatalf("expected two stable same-line array comparison findings, got %+v", got)
	}
	if !strings.Contains(got[0].Message, "leftValues") || !strings.Contains(got[1].Message, "rightValues") {
		t.Fatalf("array comparison findings are not emitted in stable variable order: %+v", got)
	}
}

func TestAnalyzerArrayComparisonUsesDirectExpressionOperands(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Variant
  Dim other() As Variant
  Dim scalar As Long
  Dim i As Long
  Dim flag As Boolean
  values = Split("a", ",")
  values = scalar
  other = values
  If UBound(values) > 0 Then Debug.Print "bounds"
  If CountValues(values) = 0 Then Debug.Print "count"
  For i = LBound(values) To UBound(values)
    If values = scalar Then Debug.Print "bad"
  Next i
  If (values) <> other Then Debug.Print "bad"
  flag = (values = scalar)
  values(i) = scalar
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA209")
	if len(got) != 4 {
		t.Fatalf("expected four direct array-operand findings, got %+v", got)
	}
	wantLines := []int{14, 16, 16, 17}
	for i, finding := range got {
		if finding.Line != wantLines[i] {
			t.Fatalf("VBA209[%d] line = %d, want %d: %+v", i, finding.Line, wantLines[i], got)
		}
	}
	if strings.Contains(got[0].Message, "other") {
		t.Fatalf("unexpected array assignment or argument finding: %+v", got)
	}
	if !strings.Contains(got[1].Message, "other") || !strings.Contains(got[2].Message, "values") || !strings.Contains(got[3].Message, "values") {
		t.Fatalf("unexpected direct operand findings: %+v", got)
	}
}

func TestAnalyzerArrayComparisonDoesNotTreatColonDeclarationsAsArrays(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim objectValue As Object: Set objectValue = CreateObject("Scripting.Dictionary")
  If objectValue Is Nothing Then Debug.Print "missing"
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA209"); len(got) != 0 {
		t.Fatalf("colon-separated object declaration must not become an array comparison: %+v", got)
	}
}

func TestAnalyzerObjectNothingAssignmentIsNotComparison(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal shouldClear As Boolean)
  Dim assignedObject As Object
  Dim comparedObject As Object
  If shouldClear Then Set assignedObject = Nothing: If comparedObject = Nothing Then Debug.Print "bad"
  If comparedObject Is Nothing Then Debug.Print "safe"
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA209")
	if len(got) != 1 || got[0].Line != 5 ||
		!strings.Contains(got[0].Message, "comparedObject") ||
		strings.Contains(got[0].Message, "assignedObject") {
		t.Fatalf("expected only the scalar object comparison for comparedObject, got %+v", got)
	}
}

func TestVBA209BatchAndRealtimeResultsMatch(t *testing.T) {
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Dim values() As Variant
  Dim scalar As Long
  values = Split("a", ",")
  If values = scalar Then Debug.Print "bad"
  If UBound(values) > 0 Then Debug.Print "safe"
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
	if got, want := findingsByCode(realtime, "VBA209"), findingsByCode(batch, "VBA209"); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime VBA209 findings differ:\nbatch=%+v\nrealtime=%+v", want, got)
	}
}

func TestAnalyzerArrayLifecycleAndDimensionSafety(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function BuildValues() As Variant
  Dim result() As Long
  ReDim result(1 To 2, 1 To 2)
  result(1, 1) = 1
  BuildValues = result
End Function

Public Sub Run()
  Dim fixed(1 To 2) As Long
  Dim values() As Long
  Dim unknown As Variant
  Dim objects() As Worksheet
  fixed(1) = 1
  values(0) = 1
  ReDim values(0 To 1, 0 To 1)
  values(0, 0) = 1
  ReDim Preserve values(1 To 3, 0 To 1)
  Erase values
  If LBound(values, 2) > 0 Then Debug.Print "bad"
  unknown = ExternalArray()
  unknown(0) = 1
  values = BuildValues()
  values(1, 1) = 2
  ReDim fixed(1 To 3)
  ReDim objects(0 To 0)
  objects(0) = Worksheets(1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got227 := findingsByCode(findings, "VBA227")
	got249 := findingsByCode(findings, "VBA249")
	allArrayFindings := append(append([]Finding(nil), got227...), got249...)
	for _, line := range []int{15, 20, 25} {
		found := false
		for _, finding := range allArrayFindings {
			if finding.Line == line {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("expected deterministic array finding on line %d, got VBA227=%+v VBA249=%+v", line, got227, got249)
		}
	}
	if got208 := findingsByCode(findings, "VBA208"); len(got208) != 1 || got208[0].Line != 18 {
		t.Fatalf("expected one non-final ReDim Preserve warning, got %+v", got208)
	}
	if got101 := findingsByCode(findings, "VBA101"); len(got101) != 1 || got101[0].Line != 27 {
		t.Fatalf("expected missing Set for object array element, got %+v", got101)
	}
	if len(got227) != 1 || len(got249) != 2 {
		t.Fatalf("unexpected array ownership projection: VBA227=%+v VBA249=%+v", got227, got249)
	}
}

func TestAnalyzerVBA227WholeArrayAssignmentDoesNotLookLikeElementAccess(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Variant
  values() = Split("a|b", "|")
  If UBound(values) <> 1 Then Debug.Print "bad"
  values(0) = "ok"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("whole-array Split assignment should establish allocation without an indexed-use finding: %+v", got)
	}
	if got := findingsByCode(findings, "VBA249"); len(got) != 0 {
		t.Fatalf("whole-array Split assignment should not leave a deterministic duplicate runtime finding: %+v", got)
	}
}

func TestAnalyzerVBA227RecognizesStringAssignmentToByteArray(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub WriteEscaped(ByVal text As String)
  Dim bytes() As Byte
  Dim index As Long
  Dim length As Long

  length = Len(text)
  If length = 0 Then Exit Sub
  bytes = text
  For index = 1 To length
    If bytes((index - 1) * 2 + 1) = 0 Then
      Debug.Print bytes((index - 1) * 2)
    End If
  Next index
End Sub

Private Sub Unsafe(ByVal text As String)
  Dim bytes() As Byte
  bytes = text
  Debug.Print bytes(0)
End Sub

Public Sub Run()
  WriteEscaped "text"
  Unsafe "text"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 1 || got[0].Procedure != "Unsafe" {
		t.Fatalf("only an unguarded String-to-Byte-array access should remain diagnosed: %+v", got)
	}
}

func TestAnalyzerVBA227RecognizesQualifiedAndNestedArrayFactoryCalls(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim source As Object
  Dim parts() As String
  parts = VBA.Split(source.Path, ".")
  If UBound(parts) > 0 Then Debug.Print parts(0)
  parts = VBA.Split(VBA.Right$(source.Path, 3), ".")
  If UBound(parts) > 0 Then Debug.Print parts(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("qualified and nested Split calls should establish array allocation: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesIsArrayGuardToVariantAssignment(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal source As Variant)
  Dim values() As String
  If IsArray(source) Then
    values = source
  Else
    values = Split(CStr(source), ",")
  End If
  If UBound(values) < 0 Then Debug.Print "bad"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("an IsArray guard should establish allocation for a Variant whole-array assignment: %+v", got)
	}
}

func TestAnalyzerVBA227RecognizesVariantByteArrayTransfer(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit

Private Function EmptyByteArray() As Variant
  Dim bytes() As Byte
  bytes = vbNullString
  EmptyByteArray = bytes
End Function

Private Function ReadByteArray() As Variant
  Dim stream As Object
  Dim encoded As Variant
  Set stream = CreateObject("ADODB.Stream")
  encoded = stream.Read(-1)
  ReadByteArray = encoded
End Function

Private Function EncodeText(ByVal text As String) As Variant
  If Len(text) = 0 Then
    EncodeText = EmptyByteArray()
  Else
    EncodeText = ReadByteArray()
  End If
End Function

Private Sub RaiseContractError()
  Err.Raise 5
End Sub

Private Sub Coerce(ByVal value As Variant, ByRef bytes() As Byte)
  If VarTypeOf(value) = vbString Then
    bytes = EncodeText(CStr(value))
  ElseIf VarTypeOf(value) = (vbArray Or vbByte) Then
    bytes = value
  Else
    RaiseContractError
  End If
  If UBound(bytes) < LBound(bytes) Then Debug.Print "empty"
End Sub

Public Sub Run(ByVal value As Variant)
  Dim bytes() As Byte
  Coerce value, bytes
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a Variant Byte-array transfer and zero-length array should make the bounds query safe: %+v", got)
	}
}

func TestAnalyzerVBA227KeepsEmptyByteArrayElementAccessUnsafe(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit

Public Sub Run()
  Dim bytes() As Byte
  bytes = vbNullString
  If UBound(bytes) < LBound(bytes) Then Debug.Print "empty"
  bytes(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 7 {
		t.Fatalf("expected only the empty-array element access to be reported, got %+v", got)
	}
}

func TestAnalyzerVBA227KeepsPotentiallyEmptyVariantByteArrayElementAccessUnsafe(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit

Private Function EmptyByteArray() As Variant
  Dim bytes() As Byte
  bytes = vbNullString
  EmptyByteArray = bytes
End Function

Private Function ReadByteArray() As Variant
  Dim stream As Object
  ReadByteArray = stream.Read(-1)
End Function

Private Function EncodeText(ByVal text As String) As Variant
  If Len(text) = 0 Then
    EncodeText = EmptyByteArray()
  Else
    EncodeText = ReadByteArray()
  End If
End Function

Public Sub Run(ByVal text As String)
  Dim bytes() As Byte
  bytes = EncodeText(text)
  bytes(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 25 {
		t.Fatalf("expected the possibly empty Variant array element access to be reported, got %+v", got)
	}
}

func TestAnalyzerVBA227CarriesAllocatedArrayThroughPrivateByRefCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Consume(ByRef values() As String)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Public Sub Run()
  Dim values() As String
  values = Split("a|b", "|")
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("an allocated array passed to a unique private ByRef helper should remain allocated: %+v", got)
	}
}

func TestAnalyzerVBA227CarriesAllocatedArrayThroughLineContinuationByRefCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Consume(ByRef values() As Byte, ByVal count As Long)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Public Sub Run()
  Dim values() As Byte
  ReDim values(0 To 1)
  Consume values, _
    2
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("an allocated array should remain allocated through a line-continuation ByRef call: %+v", got)
	}
}

func TestAnalyzerVBA227CarriesAllocatedArrayThroughRecursivePrivateByRefCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub SortValues(ByRef values() As String, ByVal low As Long, ByVal high As Long)
  If low >= high Then Exit Sub
  Debug.Print values(low)
  SortValues values, low, high - 1
End Sub

Public Sub Run()
  Dim values() As String
  ReDim values(1 To 2)
  values(1) = "a"
  values(2) = "b"
  SortValues values, 1, 2
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("an allocated array should remain allocated through a recursive private ByRef helper: %+v", got)
	}
}

func TestAnalyzerVBA227CarriesAllocatedArrayThroughPrivateByRefHelperChain(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Consume(ByRef values() As String)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Private Sub Forward(ByRef values() As String)
  Consume values
End Sub

Public Sub Run()
  Dim values() As String
  values = Split("a|b", "|")
  Forward values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("an allocated array should remain allocated through a private ByRef helper chain: %+v", got)
	}
}

func TestAnalyzerVBA227CarriesAllocationFromByRefOutputWithCollectionCount(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub PrepareInvocationValues(ByVal arguments As Collection, ByRef values() As Variant)
  If arguments.Count = 0 Then Exit Sub
  ReDim values(0 To arguments.Count - 1)
End Sub

Private Function InvokeMethod(ByVal arguments As Collection) As Variant
  Dim values() As Variant
  PrepareInvocationValues arguments, values
  If arguments.Count > 0 Then Debug.Print values(0)
  Debug.Print RunValues(values, arguments.Count)
  Select Case arguments.Count
    Case 0
      InvokeMethod = CallByName(arguments, "Value", VbMethod)
    Case 1
      InvokeMethod = CallByName( _
        arguments, "Value", VbMethod, values(0))
  End Select
End Function

Private Function RunValues(ByRef values() As Variant, ByVal argumentCount As Long) As Variant
  Select Case argumentCount
    Case 0
      RunValues = Empty
    Case 1
      RunValues = Application.Run("Value", values(0))
  End Select
End Function

Public Sub Run()
  Dim arguments As Collection
  Set arguments = New Collection
  arguments.Add "value"
  Debug.Print InvokeMethod(arguments)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a positive-count dispatch should not report its case-local array access: %+v", got)
	}
}

func TestAnalyzerVBA227CarriesArrayLengthFromPairedByRefOutput(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub CoerceHashInput(ByVal value As Variant, ByRef bytes() As Byte, ByRef byteLength As Long)
  bytes = value
  If UBound(bytes) < LBound(bytes) Then
    byteLength = 0
  Else
    byteLength = UBound(bytes) - LBound(bytes) + 1
  End If
End Sub

Private Sub BcryptHashBytes(ByVal useHmac As Boolean, ByRef keyBytes() As Byte, ByVal keyLength As Long, ByRef dataBytes() As Byte, ByVal dataLength As Long)
  If useHmac And keyLength > 0 Then
    Debug.Print keyBytes(LBound(keyBytes))
  End If
  If dataLength > 0 Then
    Debug.Print dataBytes(LBound(dataBytes))
  End If
End Sub

Private Sub HmacHash(ByVal value As Variant)
  Dim keyBytes() As Byte
  Dim keyLength As Long
  Dim dataBytes() As Byte
  Dim dataLength As Long
  CoerceHashInput value, keyBytes, keyLength
  CoerceHashInput value, dataBytes, dataLength
  BcryptHashBytes True, keyBytes, keyLength, dataBytes, dataLength
End Sub

Private Sub PlainHash(ByVal value As Variant)
  Dim dataBytes() As Byte
  Dim dataLength As Long
  Dim keyBytes() As Byte
  CoerceHashInput value, dataBytes, dataLength
  BcryptHashBytes False, keyBytes, 0, dataBytes, dataLength
End Sub

Public Sub Run()
  Dim source() As Byte
  ReDim source(0 To 1)
  HmacHash source
  PlainHash source
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	seenCoerce := false
	for _, finding := range findingsByCode(findings, "VBA227") {
		if finding.Procedure == "CoerceHashInput" {
			seenCoerce = true
		}
		if finding.Procedure == "BcryptHashBytes" {
			t.Fatalf("paired ByRef array/length outputs should satisfy guarded helper accesses: %+v", finding)
		}
	}
	if !seenCoerce {
		t.Fatalf("the fixture should still analyze the helper's direct bound probes")
	}
}

func TestAnalyzerVBA227ProcessesArrayAllocationsInMultilineCFGBlocks(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Consume(ByRef lines() As String)
  Dim fields() As String
  Dim lineIndex As Long
  For lineIndex = LBound(lines) To UBound(lines)
    fields = Split(lines(lineIndex), ",")
    If UBound(fields) > 0 Then Debug.Print fields(0)
  Next lineIndex
End Sub

Public Sub Run()
  Dim lines() As String
  lines = Split("a,b", ",")
  Consume lines
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("an allocation before a bounds query in a multiline CFG block should be honored: %+v", got)
	}
}

func TestAnalyzerVBA227KeepsDeterministicAllocationOnResumeNextEdges(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal shouldAllocate As Boolean)
  On Error Resume Next
  Dim values() As Long
  If shouldAllocate Then
    ReDim values(0 To 1)
    values(0) = 1
  End If
End Sub

Public Sub Copy()
  On Error Resume Next
  Dim values() As String
  values = Split("a,b", ",")
  values(0) = "a"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a valid plain ReDim must establish allocation on the same Resume Next branch: %+v", got)
	}
}

func TestAnalyzerVBA227DoesNotTreatWholeArrayCallAsIndexedAccess(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Consume(ByVal values As Variant)
End Sub

Public Sub Run()
  On Error Resume Next
  Dim values() As String
  ReDim values(0 To 0, 0 To 0)
  ReDim Preserve values(0 To 0, 0 To 1)
  Consume values()
End Sub`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a whole-array call argument must not be treated as an indexed access: %+v", got)
	}
}

func TestAnalyzerVBA227AllowsRepeatedPrivateByRefCallsOnOneLine(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub Consume(ByRef values() As String)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Public Sub Run()
  Dim values(0 To 1) As String
  Consume values: Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("repeated calls to the same private ByRef helper should preserve the fixed-array proof: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesModuleArrayAllocationThroughPrivateSetupCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long

Private Sub SetupValues()
  ReDim values(0 To 1)
End Sub

Private Sub Consume(ByRef items() As Long)
  If UBound(items) > 0 Then Debug.Print items(0)
End Sub

Private Sub RunInternal()
  SetupValues
  Consume values
End Sub

Public Sub Run()
  RunInternal
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a module-level array allocated by a private setup call should remain allocated for a private ByRef helper: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesFormInitializationArrayState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeFormSidecar(t, dir, "Dialog.bas", `Option Explicit
Private values() As Long

Private Sub UserForm_Initialize()
  SetupValues
End Sub

Private Sub SetupValues()
  ReDim values(0 To 1)
End Sub

Private Sub Consume(ByRef items() As Long)
  If UBound(items) > 0 Then Debug.Print items(0)
End Sub

Private Sub cmdStart_Click()
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a form-initialized module array should remain allocated for later private ByRef helpers: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesClassInitializationArrayState(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Dialog.cls", `Option Explicit
Private values() As Long

Private Sub Class_Initialize()
  SetupValues
End Sub

Private Sub SetupValues()
  ReDim values(0 To 1)
End Sub

Private Sub Consume(ByRef items() As Long)
  If UBound(items) > 0 Then Debug.Print items(0)
End Sub

Public Sub Run()
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a class-initialized module array should remain allocated for later private ByRef helpers: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesConfiguredClassArrayStateThroughGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Configured.cls", `Option Explicit
Private items() As Long
Private configured As Boolean

Private Sub ConfigureItems()
  configured = True
  ResetArray items
End Sub

Private Sub ResetArray(ByRef target() As Long)
  ReDim target(1 To 2)
End Sub

Private Sub Consume(ByRef values() As Long)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Private Sub RequireItems(ByVal memberName As String)
  If Not configured Then Err.Raise 5
End Sub

Public Sub Run()
  ConfigureItems
  RequireItems "Run"
  Consume items
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a configured class array should remain allocated through a ByRef reset and guard: %+v", got)
	}
}

func TestAnalyzerVBA227UsesConfiguredClassArrayGuards(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Configured.cls", `Option Explicit
Private items() As Long
Private keys() As Long
Private priorities() As Long
Private roleValue As Long
Private collectionKind As Long

Friend Sub ConfigureDataTable()
  roleValue = 1
  ResetArray items
  ResetArray keys
End Sub

Friend Sub ConfigureGenericCollection()
  roleValue = 2
  collectionKind = 3
  ResetArray items
  ResetArray keys
  ResetArray priorities
End Sub

Private Sub ResetArray(ByRef target() As Long)
  ReDim target(1 To 2)
End Sub

Private Sub Consume(ByRef values() As Long)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Private Sub RequireDataTable(ByVal memberName As String)
  If roleValue <> 1 Then Err.Raise 5
End Sub

Private Function IsPriorityQueueKind(ByVal value As Long) As Boolean
  IsPriorityQueueKind = (value = 3)
End Function

Private Sub RequirePriorityQueue(ByVal memberName As String)
  If Not IsPriorityQueueKind(collectionKind) Then Err.Raise 5
End Sub

Private Function IsGenericCollectionRole(ByVal value As Long) As Boolean
  IsGenericCollectionRole = (value = 2)
End Function

Public Sub DataTableRead()
  RequireDataTable "DataTableRead"
  Consume items
  Consume keys
End Sub

Public Sub PriorityRead()
  RequirePriorityQueue "PriorityRead"
  Consume items
  Consume priorities
End Sub

Public Sub GenericRead()
  If IsGenericCollectionRole(roleValue) Then
    Consume items
    Consume keys
    Consume priorities
  End If
End Sub

Public Sub UnsafeRead()
  Consume items
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, procedure := range []string{"DataTableRead", "PriorityRead", "GenericRead"} {
		for _, finding := range findingsByCode(findings, "VBA227") {
			if finding.Procedure == procedure {
				t.Fatalf("configured %s array guard should establish allocation: %+v", procedure, finding)
			}
		}
	}
	unsafe := 0
	for _, finding := range findings {
		if finding.Procedure == "Consume" && finding.Code == "VBA227" {
			unsafe++
		}
	}
	if unsafe == 0 {
		t.Fatalf("an unguarded configured-class array access must remain reportable: %+v", findings)
	}
}

func TestAnalyzerVBA227KeepsConditionalModuleSetupConservative(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long

Private Sub SetupValues()
  ReDim values(0 To 1)
End Sub

Private Sub Consume(ByRef items() As Long)
  If UBound(items) > 0 Then Debug.Print items(0)
End Sub

Public Sub Run(ByVal shouldSetup As Boolean)
  If shouldSetup Then SetupValues
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 2 {
		t.Fatalf("a conditional setup call must not prove a module-level array allocation: %+v", got)
	}
}

func TestAnalyzerVBA227DoesNotPropagateShadowedSetupArray(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private values() As Long

Private Sub SetupValues()
  Dim values() As Long
  ReDim values(0 To 1)
End Sub

Private Sub Consume(ByRef items() As Long)
  If UBound(items) > 0 Then Debug.Print items(0)
End Sub

Public Sub Run()
  SetupValues
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 2 {
		t.Fatalf("a local array shadowing the module target must not prove the module allocation: %+v", got)
	}
}

func TestAnalyzerVBA227KeepsPublicByRefArrayCallsConservative(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Consume(ByRef values() As String)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Public Sub Run()
  Dim values() As String
  values = Split("a|b", "|")
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 2 || got[0].Procedure != "Consume" || got[1].Procedure != "Consume" {
		t.Fatalf("public ByRef helper must remain conservative: %+v", got)
	}
}

func TestAnalyzerVBA227TreatsOptionPrivateModuleAsProjectPrivate(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Private Module
Option Explicit
Public Sub Consume(ByRef values() As String)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Public Sub Run()
  Dim values() As String
  values = Split("a|b", "|")
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("Option Private Module should allow proven project-local ByRef propagation: %+v", got)
	}
}

func TestAnalyzerVBA227ValidatesScalarArrayOperationsAndIterableSources(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim scalar As Long
  Dim values() As Long
  Dim unknown As Variant
  Dim implicit
  Dim userValue As UserDefinedType
  Dim item As Variant
  Erase scalar
  If LBound(scalar) > 0 Then Debug.Print "bad"
  If UBound(scalar) > 0 Then Debug.Print "bad"
  For Each item In scalar
  Next item
  For Each item In True
  Next item
  For Each item In 1 + 2
  Next item
  For Each item In unknown
  Next item
  For Each item In implicit
  Next item
  For Each item In userValue
  Next item
  Erase implicit
  If LBound(implicit) > 0 Then Debug.Print "bad"
  ReDim implicit(0 To 1)
  ReDim values(3 To 1)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 7 {
		t.Fatalf("scalar array operations and impossible ReDim bounds = %+v, want seven findings", got)
	}
	for _, finding := range got {
		if finding.Line == 18 {
			t.Fatalf("unknown Variant For Each source must remain conservative: %+v", finding)
		}
	}
	for _, line := range []int{9, 10, 11, 12, 14, 16, 27} {
		found := false
		for _, finding := range got {
			if finding.Line == line {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("expected VBA227 on line %d, got %+v", line, got)
		}
	}
}

func TestAnalyzerVBA227KeepsUnknownVariantReDimConservative(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim unknown As Variant
  Dim proven As Variant
  ReDim unknown(0 To 1)
  proven = Array(1, 2)
  ReDim Preserve proven(0 To 2)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA227") {
		if finding.Line == 5 {
			t.Fatalf("unknown Variant ReDim must remain silent: %+v", finding)
		}
	}
}

func TestAnalyzerVBA227AppliesOptionBaseToReDimBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Option Base 1
Public Sub Run()
  Dim values() As Long
  ReDim values(0)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 5 {
		t.Fatalf("Option Base 1 should make ReDim values(0) impossible: %#v", got)
	}
}

func TestAnalyzerVBA227EvaluatesEnumReDimBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Enum Bounds
  LowerBound = 3
End Enum
Public Sub Run()
  Dim values() As Long
  ReDim values(LowerBound To 1)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 7 {
		t.Fatalf("Enum bound should make ReDim impossible: %#v", got)
	}
}

func TestAnalyzerVBA227EvaluatesConstReDimBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Const LowerBound As Long = 3
Public Sub Run()
  Dim values() As Long
  ReDim values(LowerBound To 1)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 5 {
		t.Fatalf("Const bound should make ReDim impossible: %#v", got)
	}
}

func TestAnalyzerVBA227ScopesLocalConstReDimBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Sub First()
  Const LowerBound As Long = 3
  Dim values() As Long
  ReDim values(LowerBound To 1)
End Sub

Private Sub Second()
  Const LowerBound As Long = 0
  Dim values() As Long
  ReDim values(LowerBound To 1)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 5 {
		t.Fatalf("local Const bounds must be scoped to each procedure: %#v", got)
	}
}

func TestAnalyzerVBA227AcceptsParamArrayForEachSource(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ParamArray values() As Variant)
  Dim item As Variant
  For Each item In values
  Next item
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("ParamArray is always iterable and should not produce VBA227: %#v", got)
	}
}

func TestAnalyzerVBA227TreatsParamArrayAsAllocated(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ParamArray values() As Variant)
  If UBound(values) >= 0 Then Debug.Print values(0)
End Sub

Public Sub EmptyCall()
  Run
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("ParamArray bounds and element access should use its guaranteed array shape: %+v", got)
	}
}

func TestAnalyzerVBA227RejectsScalarArrayElementAsForEachSource(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Long
  Dim item As Variant
  ReDim values(0 To 1)
  For Each item In values(0)
  Next item
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 1 || got[0].Line != 6 {
		t.Fatalf("For Each over a scalar array element should produce VBA227: %#v", got)
	}
}

func TestAnalyzerVBA227RejectsEraseAndBoundsOnObjectScalar(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim value As Object
  Erase value
  If LBound(value) > 0 Then Debug.Print "bad"
  If UBound(value) > 0 Then Debug.Print "bad"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 3 {
		t.Fatalf("object scalar array operations should be diagnosed: %#v", got)
	}
}

func TestAnalyzerVBA227RejectsBoundsOnScalarExpressions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  If LBound(1) > 0 Then Debug.Print "bad"
  If UBound("value") > 0 Then Debug.Print "bad"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 2 {
		t.Fatalf("literal scalar bound targets should be diagnosed: %#v", got)
	}
}

func TestAnalyzerVBA227RecognizesFixedArrayShapeWithValidByRefParameter(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Resize(ByRef values() As Long)
  Dim fixedValues(0 To 1) As Long
  ReDim fixedValues(0 To 2)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 1 || got[0].Line != 4 {
		t.Fatalf("fixed local array ReDim should be diagnosed with a valid ByRef parameter: %#v", got)
	}
}

func TestAnalyzerVBA227TreatsByRefArrayAsAllocatedAfterReDim(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Probe(ByRef values() As Long)
  ReDim values(0 To 1)
  If LBound(values) <> 0 Then Debug.Print "bad"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("ByRef array is allocated after ReDim: %#v", got)
	}
}

func TestAnalyzerVBA227UsesSuccessfulArrayLengthGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Private Function CountBytes(ByRef values() As Byte) As Long
  On Error GoTo EmptyValues
  CountBytes = UBound(values) - LBound(values) + 1
  Exit Function
EmptyValues:
  CountBytes = 0
End Function

Public Sub Guarded(ByRef values() As Byte)
  If CountBytes(values) = 0 Then
    Debug.Print "empty"
  Else
    If LBound(values) > 0 Then Debug.Print "safe"
    values(LBound(values)) = 1
  End If
End Sub

Public Sub PositiveGuard(ByRef values() As Byte)
  If CountBytes(values) > 0 Then
    If UBound(values) > 0 Then Debug.Print "safe"
  End If
End Sub

Public Sub Unguarded(ByRef values() As Byte)
  If LBound(values) > 0 Then Debug.Print "unsafe"
End Sub
`
	writeModule(t, dir, "Main.bas", source)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	assertGuardResults := func(surface string, got []Finding) {
		for _, finding := range findingsByCode(got, "VBA227") {
			if finding.Procedure == "CountBytes" {
				t.Fatalf("%s: the recognized allocation probe handles its own bounds failure: %+v", surface, finding)
			}
			if finding.Procedure == "Guarded" || finding.Procedure == "PositiveGuard" {
				t.Fatalf("%s: successful array-length guard should establish allocation on its safe branch: %+v", surface, finding)
			}
		}
		unsafeCount := 0
		for _, finding := range findingsByCode(got, "VBA227") {
			if finding.Procedure == "Unguarded" {
				unsafeCount++
			}
		}
		if unsafeCount == 0 {
			t.Fatalf("%s: unguarded array bounds access should remain diagnosed: %+v", surface, findingsByCode(got, "VBA227"))
		}
	}
	assertGuardResults("batch", findings)
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	assertGuardResults("realtime", realtime)
}

func TestAnalyzerVBA227UsesSuccessfulArrayLengthGuardForModuleArray(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Private moduleBytes() As Byte

Private Function CountBytes(ByRef values() As Byte) As Long
  On Error GoTo EmptyValues
  CountBytes = UBound(values) - LBound(values) + 1
  Exit Function
EmptyValues:
  CountBytes = 0
End Function

Private Sub WriteFileOverlapped(ByVal handle As Long, ByVal address As Long, ByVal count As Long, ByVal flags As Long, ByVal extra As Long)
End Sub

Private Sub Guarded(ByRef payload() As Byte)
  moduleBytes = payload
  If CountBytes(moduleBytes) = 0 Then Exit Sub
  WriteFileOverlapped 0, VarPtr(moduleBytes(0)), CountBytes(moduleBytes), 0, 0
End Sub

Public Sub Run()
  Dim payload() As Byte
  ReDim payload(0 To 0)
  Guarded payload
End Sub
`
	writeModule(t, dir, "Main.bas", source)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a successful array-length guard should establish allocation for a module array: %+v", got)
	}
}

func TestAnalyzerVBA227RecognizesImplicitZeroArrayLengthGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private moduleBytes() As Byte

Private Function ByteArrayLength(ByRef values() As Byte) As Long
  On Error GoTo Unallocated
  ByteArrayLength = UBound(values) - LBound(values) + 1
  Exit Function
Unallocated:
End Function

Private Sub Guarded(ByRef payload() As Byte)
  moduleBytes = payload
  If ByteArrayLength(moduleBytes) = 0 Then Exit Sub
  Debug.Print moduleBytes(0)
End Sub

Public Sub Run()
  Dim payload() As Byte
  ReDim payload(0 To 0)
  Guarded payload
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("an allocation probe with an implicit zero recovery return should establish allocation: %+v", got)
	}
}

func TestAnalyzerVBA227RecognizesResumeNextArrayCapacityProbe(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit

Private Sub AppendBytes(ByRef target() As Byte, ByRef targetCount As Long, ByRef source() As Byte, ByVal sourceCount As Long)
  Dim capacity As Long
  Dim index As Long
  On Error Resume Next
  capacity = UBound(target) + 1
  If Err.Number <> 0 Then
    Err.Clear
    capacity = 0
  End If
  On Error GoTo 0
  If targetCount + sourceCount > capacity Then
    capacity = targetCount + sourceCount
    ReDim Preserve target(0 To capacity - 1)
  End If
  For index = 0 To sourceCount - 1
    target(targetCount + index) = source(index)
  Next index
End Sub

Private Sub Unsafe(ByRef values() As Byte)
  Dim capacity As Long
  On Error Resume Next
  capacity = UBound(values) + 1
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error GoTo 0
  Debug.Print values(0)
End Sub

Public Sub Run()
  Dim target() As Byte
  Dim source() As Byte
  ReDim source(0 To 0)
  AppendBytes target, 0, source, 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 2 || got[0].Procedure != "Unsafe" || got[1].Procedure != "Unsafe" {
		t.Fatalf("only the unguarded Resume Next probe should remain unsafe: %+v", got)
	}
}

func TestAnalyzerVBA227UsesArrayFunctionReturnShape(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Values() As Long()
  ReDim Values(0 To 1)
  Values(0) = 1
End Function
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("dynamic array function return should be resizable: %#v", got)
	}
}

func TestAnalyzerVBA227CarriesArrayLengthGuardThroughPrivateByRefCall(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function CountBytes(ByRef values() As Byte) As Long
  On Error GoTo EmptyValues
  CountBytes = UBound(values) - LBound(values) + 1
  Exit Function
EmptyValues:
  CountBytes = 0
End Function

Private Sub Consume(ByRef sourceBytes() As Byte, ByRef resultBytes() As Byte)
  sourceBytes(0) = resultBytes(0)
End Sub

Public Sub Run(ByRef sourceBytes() As Byte)
  Dim resultBytes() As Byte
  ReDim resultBytes(0)
  If CountBytes(sourceBytes) = 0 Then Exit Sub
  Consume sourceBytes, resultBytes
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a successful allocation-probe guard should carry both arrays through a private ByRef call: %+v", got)
	}
}

func TestAnalyzerVBA227RecognizesVariantArrayLengthGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ArrayLength(ByVal value As Variant) As Long
  On Error GoTo EmptyValue
  ArrayLength = UBound(value) + 1
  Exit Function
EmptyValue:
  ArrayLength = 0
End Function

Private Sub Consume(ByRef values() As Byte)
  values(0) = 1
End Sub

Public Sub Run(ByRef values() As Byte)
  If ArrayLength(values) = 0 Then Exit Sub
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a positive UBound-based Variant array-length guard should carry allocation through a private ByRef call: %+v", got)
	}
}

func TestAnalyzerVBA227CarriesArrayLengthGuardThroughScalarAssignment(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ArrayLength(ByVal value As Variant) As Long
  On Error GoTo EmptyValue
  ArrayLength = UBound(value) + 1
  Exit Function
EmptyValue:
  ArrayLength = 0
End Function

Public Sub Run(ByRef values() As Byte)
  Dim length As Long
  length = ArrayLength(values)
  If length > 0 Then
    values(0) = 1
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a positive allocation-probe result stored in a scalar should establish allocation on the guarded branch: %+v", got)
	}
}

func TestAnalyzerVBA227CarriesArrayLengthGuardThroughPropertyLet(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Utils.bas", `Option Explicit
Public Function ArrayLength(ByVal value As Variant) As Long
  On Error GoTo EmptyValue
  ArrayLength = UBound(value) + 1
  Exit Function
EmptyValue:
  ArrayLength = 0
End Function
`)
	writeClass(t, dir, "TextWindow.cls", `Option Explicit
#If Mac Then
Public Property Let Text(ByVal value As String)
  Dim bytes() As Byte
  Dim length As Long
  bytes = value
  length = ArrayLength(bytes)
  If length > 0 Then
    bytes(0) = 1
  End If
End Property
#End If
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a positive allocation-probe result should also establish allocation in a Property Let: %+v", got)
	}
}

func TestAnalyzerVBA227RecognizesUBoundOnlyArrayLengthGuard(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function GetbSize(ByRef values() As Byte) As Long
  On Error GoTo EmptyValues
  GetbSize = UBound(values) + 1
  Exit Function
EmptyValues:
  GetbSize = 0
End Function

Private Sub Consume(ByRef values() As Byte)
  If GetbSize(values) = 0 Then Exit Sub
  Debug.Print UBound(values)
End Sub

Public Sub Run(ByRef values() As Byte)
  Consume values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a positive UBound-only array-length guard should establish allocation: %+v", got)
	}
}

func TestAnalyzerVBA227ResolvesArrayReturnHelperChainsToAFixedPoint(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values() As Long
  values = Wrapper()
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub

Private Function Wrapper() As Long()
  Wrapper = Leaf()
End Function

Private Function Leaf() As Long()
  Dim result() As Long
  ReDim result(0 To 1)
  result(0) = 1
  Leaf = result
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("array-return helper chain should converge to an allocated summary: %#v", got)
	}
}

func TestAnalyzerVBA227ExcludesDefinitelyFailingArrayReturnPaths(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function MaybeValues(ByVal emptyInput As Boolean) As Long()
  Dim values() As Long
  If emptyInput Then
    ReDim values(0 To -1)
  Else
    ReDim values(0 To 1)
  End If
  MaybeValues = values
End Function

Public Sub Run()
  Dim values() As Long
  values = MaybeValues(False)
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA227") {
		if finding.Procedure == "Run" {
			t.Fatalf("a definitely failing return path should not make the successful array return unknown: %+v", finding)
		}
	}
}

func TestAnalyzerVBA227UsesAllocationGuardInArrayReturnSummary(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function CountValues(ByRef values() As Long) As Long
  On Error GoTo EmptyValues
  CountValues = UBound(values) - LBound(values) + 1
  Exit Function
EmptyValues:
  CountValues = 0
End Function

Private Function MakeValues() As Long()
  Dim values() As Long
  Dim unknownSource As Variant
  values = unknownSource
  If CountValues(values) = 0 Then
    ReDim values(0 To -1)
  Else
    ReDim values(0 To 1)
  End If
  MakeValues = values
End Function

Public Sub Run()
  Dim values() As Long
  values = MakeValues()
  If UBound(values) > 0 Then Debug.Print values(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA227") {
		if finding.Procedure == "Run" {
			t.Fatalf("the successful allocation-guard branch should make the array return safe: %+v", finding)
		}
	}
}

func TestAnalyzerVBA227RejectsScalarFunctionForEachSource(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function ScalarValue() As Long
  ScalarValue = 1
End Function

Public Sub Run()
  Dim item As Variant
  For Each item In ScalarValue()
  Next item
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 8 {
		t.Fatalf("scalar function For Each source should produce VBA227: %#v", got)
	}
}

func TestAnalyzerVBA227AllowsCollectionFunctionForEachSource(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Items() As Collection
  Set Items = New Collection
End Function

Public Sub Run()
  Dim item As Variant
  For Each item In Items()
  Next item
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("Collection function For Each source should remain valid: %#v", got)
	}
}

func TestAnalyzerProjectsVB060AndVB061IntoAnalyzeAndPreflight(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Const Limit As Long = 2
Public Sub Run()
  Dim bad(3 To 1) As Long
  Limit = 3
End Sub
`)
	result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(result.Findings, "VB060"); len(got) != 1 || got[0].Severity != "error" {
		t.Fatalf("VB060 analyze findings = %#v", got)
	}
	if got := findingsByCode(result.Findings, "VB061"); len(got) != 1 || got[0].Severity != "error" {
		t.Fatalf("VB061 analyze findings = %#v", got)
	}
	if got := findingsByCode(result.PreflightFindings, "VB060"); len(got) != 1 {
		t.Fatalf("VB060 preflight findings = %#v", got)
	}
	if got := findingsByCode(result.PreflightFindings, "VB061"); len(got) != 1 {
		t.Fatalf("VB061 preflight findings = %#v", got)
	}
}

func TestVBA208AllowsOneDimensionalAndStableNonFinalDimensions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal rowCount As Long)
  Dim values() As Long
  ReDim Preserve values(0 To rowCount)

  Dim matrix() As Variant
  ReDim matrix(1 To rowCount, 1 To 1)
  Dim columnCount As Long
  For columnCount = 2 To 3
    ReDim Preserve matrix(1 To ROWCOUNT, 1 To columnCount)
  Next columnCount
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA208"); len(got) != 0 {
		t.Fatalf("one-dimensional Preserve and stable non-final dimensions should be safe: %+v", got)
	}
}

func TestVBA208RejectsChangedSymbolicNonFinalDimension(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal rowCount As Long, ByVal actualRows As Long)
  Dim matrix() As Variant
  ReDim matrix(1 To rowCount, 1 To 1)
  ReDim Preserve matrix(1 To actualRows, 1 To 2)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA208")
	if len(got) != 1 || got[0].Line != 5 {
		t.Fatalf("a changed symbolic non-final dimension should report once: %+v", got)
	}
}

func TestVBA208DoesNotTreatArrayMemberAsDirectReDimTarget(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Type Bucket
  Elements() As Long
End Type
Private buckets() As Bucket

Public Sub Grow(ByVal index As Long)
  ReDim buckets(0 To 1, 0 To 1)
  ReDim Preserve buckets(index).Elements(0 To 1)
  ReDim Preserve buckets(0 To 1, 0 To 2)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA208"); len(got) != 0 {
		t.Fatalf("a member array must not be attributed to its receiver array: %+v", got)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("nested member ReDim must preserve the receiver lifecycle result: %+v", got)
	}
}

func TestAnalyzerArrayLifecycleIsDeterministicAcrossRuns(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal chooseFirst As Boolean)
  Dim values() As Long
  If chooseFirst Then
    ReDim values(0 To 1)
  Else
    Erase values
  End If
  If chooseFirst Then
    values(0) = 1
  Else
    values(1) = 2
  End If
End Sub
`)
	var want []Finding
	for i := 0; i < 10; i++ {
		findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
		if err != nil {
			t.Fatal(err)
		}
		if i == 0 {
			want = findings
			if len(want) == 0 {
				t.Fatal("expected array lifecycle findings, got none")
			}
			continue
		}
		if !reflect.DeepEqual(findings, want) {
			t.Fatalf("array lifecycle findings changed on run %d:\n got=%+v\nwant=%+v", i, findings, want)
		}
	}
}

func TestAnalyzerArrayLifecycleSuppressesRangeValueDuplicates(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1:B2").Value2
  Debug.Print values(1, 1)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("Range.Value shape diagnostics must remain owned by VBA226: %+v", got)
	}
}

func TestAnalyzerArrayLifecycleUsesConservativeBranchAndVariantStates(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal shouldAllocate As Boolean)
  Dim values() As Long
  Dim value As Variant
  If shouldAllocate Then ReDim values(0 To 0)
  values(0) = 1
  value = Array(1, 2)
  value(0) = 1
  value = 42
  value(0) = 1
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 6 {
		t.Fatalf("expected only the conservative dynamic-array diagnostic; unknown Variant operations stay quiet, got %+v", got)
	}
}

func TestAnalyzerArrayLifecyclePreservesMatchingShapesAndMultiTargetTransitions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run(ByVal shouldAllocate As Boolean)
  Dim values() As Long
  Dim leftValues() As Long
  Dim rightValues() As Long
  If shouldAllocate Then
    ReDim values(0 To 1)
  Else
    ReDim values(0 To 1)
  End If
  If UBound(values) <> 1 Then Debug.Print "bad"
  values(1) = 1
  Dim variantValues() As Variant
  If shouldAllocate Then
    ReDim variantValues(0 To 1)
  Else
    variantValues = Array(1, 2)
  End If
  If UBound(variantValues) <> 1 Then Debug.Print "bad"
  variantValues(0) = 1
  ReDim leftValues(1 To 2), rightValues(1 To 3)
  leftValues(2) = 1
  rightValues(3) = 1
  Erase leftValues, rightValues
  leftValues(1) = 1
  rightValues(1) = 1
End Sub
`
	writeModule(t, dir, "Main.bas", source)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA249")
	if len(got) != 2 {
		t.Fatalf("expected only the two accesses after Erase to report as deterministic runtime errors, got %+v", got)
	}
	expected := map[int]string{25: "leftValues", 26: "rightValues"}
	for _, finding := range got {
		name, ok := expected[finding.Line]
		lineText := ""
		if finding.Line > 0 {
			lines := strings.Split(source, "\n")
			if finding.Line <= len(lines) {
				lineText = lines[finding.Line-1]
			}
		}
		if !ok || !strings.Contains(lineText, name) {
			t.Fatalf("unexpected VBA249 finding after multi-target Erase: %+v", finding)
		}
		if finding.RuntimeError == nil || finding.RuntimeError.Kind != "array_unallocated" {
			t.Fatalf("unexpected VBA249 context after multi-target Erase: %+v", finding)
		}
	}
}

func TestAnalyzerArrayLifecycleAcceptsReDimTypeSuffixAndUnknownDimension(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal dimension As Long)
	Dim values() As Long
	ReDim values(0 To 1) As Long
	values(1) = 1
	If UBound(values, dimension) > 0 Then Debug.Print "ok"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("typed ReDim and an unknown dimension expression should not produce VBA227: %+v", got)
	}
}

func TestAnalyzerArrayLifecycleQualifiedReDimAndUnknownArrayAssignment(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim qualified() As Variant
  ReDim qualified(0 To 0) As Scripting.Dictionary
  qualified(0) = Empty
  Dim values() As Long
  ReDim values(0 To 0)
  values = ExternalValues()
  values(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 1 || got[0].Line != 9 {
		t.Fatalf("qualified ReDim should establish allocation while unknown array assignment should invalidate it: %+v", got)
	}
}

func TestAnalyzerArrayLifecycleKeepsAlwaysEnabledObjectArrayFindings(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim objects() As Worksheet
  ReDim objects(0 To 0)
  objects(0) = Worksheets(1)
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = false
	cfg.Analyze.DetectRedimPreserveDimension = false
	cfg.Analyze.DetectObjectArrayComparison = false

	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA101"); len(got) != 1 {
		t.Fatalf("always-enabled object-array missing-Set finding was suppressed: %+v", got)
	}
}

func TestAnalyzerArrayReturnSummariesRequireDefiniteAllocation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function MaybeValues(ByVal shouldAllocate As Boolean) As Variant
  Dim result() As Long
  If shouldAllocate Then
    ReDim result(0 To 0)
  End If
  MaybeValues = result
End Function

Private Function AlwaysValues(ByVal shouldAllocate As Boolean) As Variant
  Dim result() As Long
  If shouldAllocate Then
    ReDim result(0 To 0)
  Else
    ReDim result(0 To 0)
  End If
  AlwaysValues = result
End Function

Public Sub Run()
  Dim unsafe As Variant
  Dim safe As Variant
  unsafe = MaybeValues(True)
  unsafe(0) = 1
  safe = AlwaysValues(True)
  safe(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 0 {
		t.Fatalf("Variant array-return assignments remain unknown unless array nature is proven: %+v", got)
	}
}

func TestAnalyzerArrayReturnSummaryDeduplicatesLoopVisits(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function LoopValues() As Variant
  Dim result() As Long
  ReDim result(0 To 0)
  Dim i As Long
  For i = 1 To 2
    Debug.Print i
  Next i
  LoopValues = result
End Function

Public Sub Run()
  Dim values As Variant
  values = LoopValues()
  values(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("a loop revisiting one allocated return assignment must not invalidate its summary: %+v", got)
	}
}

func TestAnalyzerArrayReturnSummaryExcludesScalarRangeValue(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Function ReadScalar() As Variant
  ReadScalar = Range("A1").Value
End Function

Public Sub Run()
  Dim value As Variant
  value = ReadScalar()
  value(0) = 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA227")
	if len(got) != 0 {
		t.Fatalf("scalar Range.Value return must remain unknown without an array proof: %+v", got)
	}
}

func TestVBA227RealtimeArrayReturnSummariesAreDocumentLocal(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Helper.bas", `Option Explicit
Public Function BuildValues() As Variant
  Dim result() As Long
  ReDim result(0 To 0)
  BuildValues = result
End Function
`)
	source := `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = BuildValues()
  values(0) = 1
End Sub
`
	writeModule(t, dir, "Main.bas", source)

	findings, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("realtime analysis must keep external Variant array returns unknown: %+v", got)
	}
}

func TestVBA227BatchAndRealtimeResultsMatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Dim values() As Long
	values(0) = 1 ' xlflow:disable-line VBA227 VBA249
  ReDim values(0 To 1)
  values(1) = 2
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
	if got, want := findingsByCode(realtime, "VBA227"), findingsByCode(batch, "VBA227"); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime VBA227 findings differ:\nbatch=%+v\nrealtime=%+v", want, got)
	}
	if len(batch) != 0 {
		t.Fatalf("inline suppression should hide the only VBA227 finding: %+v", batch)
	}
}

func TestAnalyzerArrayComparisonIgnoresElementsAndProcedureHeaders(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	if !containsAll(finding.Message, "deck", "deck.Count", "non-short-circuit") ||
		!containsAll(finding.Reason, "And/Or", "eager", "runtime error 91") ||
		!containsAll(finding.Suggestion, "If/ElseIf") {
		t.Fatalf("unexpected VBA212 text: %+v", finding)
	}
}

func TestAnalyzerNonShortCircuitObjectGuardAllowsSeparateAndDifferentObjects(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerVBA212ExpandedEagerContainersAndAccessPaths(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim parent As Object
  Dim arr As Variant
  Dim arr2 As Variant
  Dim dict As Object
  If (parent Is Nothing Or parent.Child.Count = 0) Then Exit Sub
  If IsArray(arr) And arr(0) = 1 Then Exit Sub
  If Not IsArray(arr2) And arr2(0) = 1 Then Exit Sub
  If dict Is Nothing Or dict("k").Count = 0 Then Exit Sub
  If IIf(parent Is Nothing, True, parent.Child.Count = 0) Then Exit Sub
  If Choose(1, parent Is Nothing, parent.Child.Count = 0) Then Exit Sub
  If Switch(parent Is Nothing, True, parent.Child.Count = 0) Then Exit Sub
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA212")
	if len(got) < 6 {
		t.Fatalf("expanded VBA212 findings = %+v, want nested/member/index/eager hazards", got)
	}
	var sawNested, sawArray, sawDict, sawIIf, sawChoose, sawSwitch bool
	for _, finding := range got {
		sawNested = sawNested || containsAll(finding.Message, "parent", "parent.Child.Count")
		sawArray = sawArray || containsAll(finding.Message, "arr", "arr(0)")
		sawDict = sawDict || containsAll(finding.Message, "dict", `dict("k").Count`)
		sawIIf = sawIIf || containsAll(finding.Suggestion, "IIf", "If/Else")
		sawChoose = sawChoose || containsAll(finding.Suggestion, "Choose", "Select Case")
		sawSwitch = sawSwitch || containsAll(finding.Suggestion, "Switch", "If/ElseIf")
		if finding.Column == 0 || finding.EndColumn <= finding.Column {
			t.Fatalf("VBA212 should identify unsafe operand range: %+v", finding)
		}
	}
	if !sawNested || !sawArray || !sawDict || !sawIIf || !sawChoose || !sawSwitch {
		t.Fatalf("missing expanded VBA212 categories: nested=%v array=%v dict=%v iif=%v choose=%v switch=%v findings=%+v", sawNested, sawArray, sawDict, sawIIf, sawChoose, sawSwitch, got)
	}
}

func TestAnalyzerVBA212ResolvedSideEffectGetterOnlyInEagerExpressions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Widget.cls", `Option Explicit
Public Property Get Value() As Long
  Application.Calculate
  Value = 1
End Property
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim widget As Widget
  If widget.Value > 0 Then Exit Sub
  If widget Is Nothing Or widget.Value > 0 Then Exit Sub
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA212")
	if len(got) != 2 {
		t.Fatalf("resolved getter findings = %+v, want object guard plus getter hazard", got)
	}
	sawGuard, sawGetter := false, false
	for _, finding := range got {
		if finding.Line != 5 || !strings.Contains(finding.Message, "widget.Value") {
			t.Fatalf("unexpected resolved getter finding: %+v", finding)
		}
		sawGuard = sawGuard || strings.Contains(finding.Message, "dereferences widget")
		sawGetter = sawGetter || strings.Contains(finding.Message, "getter")
	}
	if !sawGuard || !sawGetter {
		t.Fatalf("resolved getter findings should retain guard and getter diagnostics: %+v", got)
	}
}

func TestAnalyzerVBA212IgnoresSafeValuesOtherOperatorsAndUserFunctions(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function IIf(ByVal condition As Boolean, ByVal yesValue As Variant, ByVal noValue As Variant) As Variant
  IIf = yesValue
End Function
Public Sub Run()
  Dim obj As Object
  Dim other As Object
  Dim arr As Variant
  If obj Is Nothing Or 1 = 0 Then Exit Sub
  If obj Is Nothing Or other Is Nothing Then Exit Sub
  If obj Is Nothing Xor obj.Count = 0 Then Exit Sub
  If obj Is Nothing Eqv obj.Count = 0 Then Exit Sub
  If obj Is Nothing Imp obj.Count = 0 Then Exit Sub
  If IsArray(arr) Or 1 = 0 Then Exit Sub
  If IIf(obj Is Nothing, True, obj.Count) Then Exit Sub
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA212"); len(got) != 0 {
		t.Fatalf("safe/value-only/operator/user-function expressions should be ignored: %+v", got)
	}
}

func TestAnalyzerVBA212BuiltinNamesIgnoreCommentsStringsAndLongerDeclarations(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
' Function IIfSafe should not shadow the builtin IIf.
Public Function IIfSafe(ByVal value As Variant) As Variant
  IIfSafe = value
End Function
Public Sub Run()
  Dim obj As Object
  Dim text As String
  text = "Function IIf is only text"
  If IIf(obj Is Nothing, True, obj.Count) Then Exit Sub
End Sub
`)
	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA212")
	if len(got) != 1 || !containsAll(got[0].Message, "obj", "obj.Count") || !containsAll(got[0].Suggestion, "IIf", "If/Else") {
		t.Fatalf("builtin IIf should not be shadowed by comments, strings, or IIfSafe: %+v", got)
	}
}

func TestRealtimeVBA212ResolvesOnlySameDocumentGetterEffects(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	path := filepath.Join(dir, "Widget.cls")
	source := []byte(`Option Explicit
Public Property Get Value() As Boolean
  Application.Calculate
  Value = True
End Property
Public Sub Run()
  If Me.Value And True Then Exit Sub
End Sub
`)
	findings, err := SourceRealtimeFindingsContext(context.Background(), dir, path, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA212")
	if len(got) != 1 || !containsAll(got[0].Message, "Me.Value", "getter") {
		t.Fatalf("same-document realtime getter finding = %+v", got)
	}
}

func TestAnalyzerDictionaryIterationValueUsageFindsKnownAndInferredDictionaries(t *testing.T) {
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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
	t.Parallel()
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

func TestAnalyzerVBA216DetectsDistinctWorksheetRoots(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	tracker := newWorksheetRootTracker(nil)
	tracker.observeSetAssignment(`Set inputSheet = ThisWorkbook.Worksheets("Input")`)
	tracker.observeSetAssignment(`Set outputSheet = ThisWorkbook.Worksheets("Output")`)
	if input, output := tracker.variables["inputsheet"], tracker.variables["outputsheet"]; input.kind != worksheetRootExplicit || output.kind != worksheetRootExplicit || input.identity == output.identity {
		t.Fatalf("worksheet selector identities = %+v / %+v", input, output)
	}
	writeWorkbookModule(t, dir, "InputSheet.bas")
	writeWorkbookModule(t, dir, "OutputSheet.bas")
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub VariableRoots()
  Dim inputSheet As Worksheet
  Dim outputSheet As Worksheet
  Dim lastRow As Long
  Set inputSheet = InputSheet
  Set outputSheet = OutputSheet
  lastRow = inputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row
End Sub

Public Sub WorkbookSelectorRoots()
  Dim inputSheet As Worksheet
  Dim outputSheet As Worksheet
  Dim lastRow As Long
  Set inputSheet = ThisWorkbook.Worksheets("Input")
  Set outputSheet = ThisWorkbook.Worksheets("Output")
  lastRow = inputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row
End Sub

Public Sub CodenameRangeRoots()
  Dim result As Range
  Set result = InputSheet.Range(OutputSheet.Cells(1, 1), OutputSheet.Cells(2, 1))
End Sub

Public Sub WithRoots()
  Dim lastRow As Long
  With InputSheet
    With .Range("A1")
      lastRow = .Cells(OutputSheet.Rows.Count, 1).End(xlUp).Row
    End With
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA216")
	if len(got) != 4 {
		t.Fatalf("VBA216 findings = %+v, want four", got)
	}
	for _, finding := range got {
		if finding.Severity != "error" || !strings.Contains(finding.Suggestion, "Cells(") {
			t.Fatalf("unexpected VBA216 finding: %+v", finding)
		}
	}
	if blocking := findingsByCode(BlockingFindings(findings), "VBA216"); len(blocking) != 4 {
		t.Fatalf("VBA216 must block preflight: %+v", blocking)
	}
}

func TestAnalyzerVBA216OnlyComparesProvableRootIdentities(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeWorkbookModule(t, dir, "InputSheet.bas")
	writeWorkbookModule(t, dir, "OutputSheet.bas")
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Selectors(ByVal Sheet1 As Worksheet, ByVal Sheet2 As Worksheet, ByVal position As Long)
  Dim lastRow As Long
  Dim target As Worksheet
  lastRow = ThisWorkbook.Worksheets("Data").Cells(ThisWorkbook.Sheets("Data").Rows.Count, 1).End(xlUp).Row
  lastRow = ThisWorkbook.Worksheets(position).Cells(ThisWorkbook.Worksheets("Data").Rows.Count, 1).End(xlUp).Row
  lastRow = InputSheet.Cells(ThisWorkbook.Worksheets("Input").Rows.Count, 1).End(xlUp).Row
  lastRow = Sheet1.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row
  Set target = InputSheet
  Set target = OutputSheet
  lastRow = target.Cells(OutputSheet.Rows.Count, 1).End(xlUp).Row
  lastRow = InputSheet.Cells(OutputSheet.Rows.Count, 1).End(xlUp).Row
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA216")
	if len(got) != 1 || got[0].Line != 12 {
		t.Fatalf("VBA216 must report only literal-name or codename mismatches, got %+v", got)
	}
}

func TestAnalyzerVBA216AnalyzesContinuationsAndWithHeaders(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeWorkbookModule(t, dir, "InputSheet.bas")
	writeWorkbookModule(t, dir, "OutputSheet.bas")
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim lastRow As Long
  lastRow = InputSheet.Cells( _
      OutputSheet.Rows.Count, 1).End(xlUp).Row
  With InputSheet.Range( _
      OutputSheet.Cells(1, 1), OutputSheet.Cells(2, 1))
  End With
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA216")
	if len(got) != 2 || got[0].Line != 4 || got[1].Line != 6 {
		t.Fatalf("VBA216 continuation and With-header findings = %+v", got)
	}
}

func TestAnalyzerVBA217AnalyzesContinuationStatements(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim lastRow As Long
  lastRow = Cells( _
      Rows.Count, 1).End(xlDown).Row
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA217")
	if len(got) != 2 || got[0].Line != 4 || got[1].Line != 4 {
		t.Fatalf("VBA217 continuation findings = %+v", got)
	}
}

func TestWorksheetRootMemberOffsetsPreserveUTF8(t *testing.T) {
	t.Parallel()
	tracker := newWorksheetRootTracker(map[string]string{"inputsheet": "InputSheet", "outputsheet": "OutputSheet"})
	accesses := worksheetMemberAccesses(`lastRow = Len("İ") + InputSheet.Cells(OutputSheet.Rows.Count, 1).End(xlUp).Row`, tracker)
	if len(accesses) != 2 || accesses[0].root.identity != "codename:inputsheet" || accesses[1].root.identity != "codename:outputsheet" {
		t.Fatalf("worksheet accesses after UTF-8 text = %+v", accesses)
	}
}

func TestAnalyzerVBA216AcceptsSameWorksheetRootsAndUnknowns(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub SameRoot()
  Dim ws As Worksheet
  Dim lastRow As Long
  Set ws = ThisWorkbook.Worksheets("Data")
  lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  With ws
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
  End With
  Set ws = GetWorksheetAtRuntime()
  lastRow = ws.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA216"); len(got) != 0 {
		t.Fatalf("same and unknown roots must not report VBA216: %+v", got)
	}
}

func TestAnalyzerVBA217ReportsOnlyUnstableLastRowPatterns(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim lastRow As Long
  Dim count As Long
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row
  lastRow = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row
  lastRow = Sheet1.Cells(1, 1).End(xlDown).Row
  lastRow = Sheet1.UsedRange.Rows.Count
  lastRow = Sheet1.Range("A1").CurrentRegion.Rows.Count
  lastRow = Sheet1.UsedRange.Row + Sheet1.UsedRange.Rows.Count - 1
  count = Sheet1.UsedRange.Rows.Count
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA217")
	if len(got) != 5 {
		t.Fatalf("VBA217 findings = %+v, want five", got)
	}
	for _, finding := range got {
		if finding.Severity != "warning" {
			t.Fatalf("VBA217 must be a warning: %+v", finding)
		}
	}
	if blocking := findingsByCode(BlockingFindings(findings), "VBA217"); len(blocking) != 0 {
		t.Fatalf("VBA217 must not block preflight: %+v", blocking)
	}
}

func TestAnalyzerVBA217HonorsDisableAndInlineSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub InlineSuppressed()
  Dim lastRow As Long
  ' xlflow:disable-next-line VBA217
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA217"); len(got) != 0 {
		t.Fatalf("VBA217 inline suppression should apply: %+v", got)
	}

	cfg := config.Default()
	cfg.Analyze.DetectUnstableLastRowPatterns = false
	findings, err = Analyzer{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA217"); len(got) != 0 {
		t.Fatalf("disabled VBA217 should not report: %+v", got)
	}
}

func TestAnalyzerWorksheetRootRulesIgnoreStringLiterals(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim lastRow As Long
  lastRow = Len("Cells(Rows.Count, 1).End(xlDown).Row")
End Sub
`)

	findings, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VBA216", "VBA217"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("%s should ignore string literals: %+v", code, got)
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

func TestVBA220DetectsEventReentryAndHonorsSafeEventCleanup(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	workbook := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbook, 0o755); err != nil {
		t.Fatal(err)
	}
	source := `Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
  Target.Value = 1
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  SelectElsewhere
End Sub

Private Sub SelectElsewhere()
  Application.Goto Range("A1")
End Sub

Private Sub Worksheet_Calculate()
  MysteryOperation
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  Dim oldEvents As Boolean
  oldEvents = Application.EnableEvents
  On Error GoTo Cleanup
  Application.EnableEvents = False
  Range("C1").Value = 1
Cleanup:
  Application.EnableEvents = oldEvents
End Sub

Private Sub Workbook_Open()
  Application.EnableEvents = False
  Range("D1").Value = 1
End Sub
`
	path := filepath.Join(workbook, "Sheet1.cls")
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA220")
	if len(got) < 4 {
		t.Fatalf("VBA220 findings = %+v, want direct, helper, uncertainty, and un-restored guard hazards", got)
	}
	if !strings.Contains(got[0].Message, "same-event") {
		t.Fatalf("first VBA220 should be direct recursion: %+v", got[0])
	}
	for _, finding := range got {
		if finding.Procedure == "Workbook_SheetChange" {
			t.Fatalf("safe cleanup should suppress VBA220: %+v", finding)
		}
	}
	if got := findingsByCode(findings, "VBA203"); len(got) == 0 {
		t.Fatalf("missing EnableEvents restoration must remain reportable: %+v", findings)
	}
	if realtime, err := SourceRealtimeFindings(dir, path, config.Default(), []byte(source)); err != nil {
		t.Fatal(err)
	} else if got := findingsByCode(realtime, "VBA220"); len(got) != 0 {
		t.Fatalf("batch-only VBA220 must not be returned in realtime: %+v", got)
	}
}

func TestVBA220ReportsUserFormControlReentryWithoutEnableEventsExemption(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeFormSidecar(t, dir, "Dialog.bas", `Option Explicit
Private Sub TextBox1_Change()
  Application.EnableEvents = False
  Me.TextBox1.Value = "updated"
  Application.EnableEvents = True
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA220")
	if len(got) != 1 || !strings.Contains(got[0].Message, "same-event") {
		t.Fatalf("UserForm control change should remain a same-event VBA220 hazard: %+v", got)
	}
}

func TestVBA220AcceptsDelegatedWorkCoveredBySafeEventCleanup(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	workbook := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbook, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "Sheet1.cls"), []byte(`Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
  SafelyWrite
End Sub

Private Sub SafelyWrite()
  Dim oldEvents As Boolean
  oldEvents = Application.EnableEvents
  On Error GoTo Cleanup
  Application.EnableEvents = False
  WriteCell
Cleanup:
  Application.EnableEvents = oldEvents
End Sub

Private Sub WriteCell()
  Range("A1").Value = 1
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, finding := range findingsByCode(findings, "VBA220") {
		if strings.Contains(finding.Message, "same-event") || strings.Contains(finding.Message, "broader event-chain") {
			t.Fatalf("delegated work covered by cleanup should not report a confirmed VBA220 hazard: %+v", finding)
		}
	}
}

func TestVBA220ReportsAmbiguousCallsAsUncertainty(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	workbook := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbook, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "Sheet1.cls"), []byte(`Option Explicit
Private Sub Worksheet_Calculate()
  RefreshData
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	writeModule(t, dir, "First.bas", "Public Sub RefreshData()\nEnd Sub\n")
	writeModule(t, dir, "Second.bas", "Public Sub RefreshData()\nEnd Sub\n")
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA220")
	if len(got) != 1 || !strings.Contains(got[0].Message, "ambiguous") || strings.Contains(got[0].Message, "same-event") {
		t.Fatalf("ambiguous event call must be an uncertainty, not a confirmed recursion: %+v", got)
	}
}

func TestVBA220ReportsOneFindingPerResolvedCallSite(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	workbook := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbook, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "Sheet1.cls"), []byte(`Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
  TriggerSelectionChange
End Sub

Private Sub TriggerSelectionChange()
  Application.Goto Range("A1")
  MysteryOperation
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA220")
	if len(got) != 1 {
		t.Fatalf("VBA220 findings = %+v, want one finding at the resolved call site", got)
	}
	if got[0].Line != 3 || !strings.Contains(got[0].Message, "broader event-chain") {
		t.Fatalf("VBA220 finding = %+v, want the confirmed event risk on line 3", got[0])
	}
}

func TestVBA220KeepsDistinctResolvedCallsInOneStatement(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	workbook := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbook, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "Sheet1.cls"), []byte(`Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
  If WriteCell(TriggerSelection()) Then
  End If
End Sub

Private Function WriteCell(ByVal triggered As Boolean) As Boolean
  Range("A1").Value = 1
  WriteCell = triggered
End Function

Private Function TriggerSelection() As Boolean
  Application.Goto Range("A1")
  TriggerSelection = True
End Function
`), 0o644); err != nil {
		t.Fatal(err)
	}

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA220")
	if len(got) != 2 {
		t.Fatalf("VBA220 findings = %+v, want one finding for each resolved call boundary", got)
	}
	if !strings.Contains(got[0].Message, "same-event") || !strings.Contains(got[1].Message, "broader event-chain") {
		t.Fatalf("VBA220 findings = %+v, want same-event and broader event-chain risks", got)
	}
}

func writeWorkbookModule(t *testing.T, dir, name string) {
	t.Helper()
	src := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	module := strings.TrimSuffix(name, filepath.Ext(name))
	body := "Attribute VB_Name = \"" + module + "\"\nOption Explicit\n"
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

func TestAnalyzerVBA229IsBlockingAndUnsuppressible(t *testing.T) {
	useCompleteTestTypeDB(t)
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Probe()
    ' xlflow:disable-next-line VBA229
    Dim value As DefinitelyNotARealType
End Sub
`)
	result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA229")
	if len(findings) != 1 || findings[0].Severity != "error" || findings[0].Line != 4 {
		t.Fatalf("VBA229 findings = %+v, want one blocking error on line 4", findings)
	}
	if blocking := findingsByCode(BlockingFindings(result.Findings), "VBA229"); len(blocking) != 1 {
		t.Fatalf("VBA229 blocking findings = %+v, want one", blocking)
	}
}

func TestAnalyzerVBA229AcceptsQualifiedProjectType(t *testing.T) {
	useCompleteTestTypeDB(t)
	dir := t.TempDir()
	writeModule(t, dir, "Types.bas", `Option Explicit
Public Type Payload
    Value As Long
End Type
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Probe()
    Dim value As Types.Payload
End Sub
`)
	result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA229"); len(findings) != 0 {
		t.Fatalf("qualified project type should resolve: %+v", findings)
	}
}

func TestAnalyzerVBA229FailsClosedWithoutTypeDBManifest(t *testing.T) {
	typeDBDir := t.TempDir()
	t.Setenv(typedb.EnvDir, typeDBDir)
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Probe()
    Dim value As ReferencedLibraryType
End Sub
`)
	result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA229"); len(findings) != 0 {
		t.Fatalf("missing TypeDB manifest must not produce VBA229: %+v", findings)
	}
}

func TestAnalyzerVBA229AcceptsUserFormAndProjectQualifiedClass(t *testing.T) {
	useCompleteTestTypeDB(t)
	dir := t.TempDir()
	writeFormSidecar(t, dir, "CalendarPicker.bas", `Option Explicit
Public Sub ShowPicker()
End Sub
`)
	writeClass(t, dir, "WebDriver.cls", `Attribute VB_Name = "WebDriver"
Option Explicit
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Probe()
    Dim picker As CalendarPicker
    Dim driver As SeleniumVBA.WebDriver
End Sub
`)
	cfg := config.Default()
	cfg.Project.Name = "SeleniumVBA"
	cfg.UserForm.CodeSource = "sidecar"
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA229"); len(findings) != 0 {
		t.Fatalf("project object types should resolve: %+v", findings)
	}
}

func TestAnalyzerVBA229AcceptsEmbeddedEnumGroups(t *testing.T) {
	useCompleteTestTypeDB(t)
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Probe()
    Dim valueKind As VbVarType
    Dim calculation As XlCalculation
End Sub
`)
	result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA229"); len(findings) != 0 {
		t.Fatalf("embedded enum groups should resolve: %+v", findings)
	}
}

func TestAnalyzerVBA227PropagatesModuleArrayThroughPrivateHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Archive.cls", `Attribute VB_Name = "Archive"
Option Explicit

Private mBytes() As Byte

Private Function ReadBytes() As Variant
    Dim values() As Byte
    ReDim values(0 To 0)
    ReadBytes = values
End Function

Public Sub Run()
    mBytes = ReadBytes()
    ConsumeBytes
End Sub

Private Sub ConsumeBytes()
    If mBytes(0) = 0 Then Debug.Print "ok"
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("allocated module array passed through a private helper should not report VBA227: %+v", got)
	}
}

func TestAnalyzerVBA227PropagatesClassInitializerArrayThroughPrivateHelper(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeClass(t, dir, "Archive.cls", `Attribute VB_Name = "Archive"
Option Explicit

Private mBytes() As Byte

Private Sub Class_Initialize()
    ReDim mBytes(0 To 0)
End Sub

Public Sub Run()
    ConsumeBytes
End Sub

Private Sub ConsumeBytes()
    If mBytes(0) = 0 Then Debug.Print "ok"
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectArrayLifecycleSafety = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA227"); len(got) != 0 {
		t.Fatalf("class-initialized module array passed through a private helper should not report VBA227: %+v", got)
	}
}

func useCompleteTestTypeDB(t *testing.T) {
	t.Helper()
	typeDBDir := t.TempDir()
	output := "test.generated.json"
	if err := os.WriteFile(filepath.Join(typeDBDir, output), []byte(`{"types":[]}`), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := typedb.WriteManifest(typeDBDir, typedb.Manifest{
		GeneratorVersion: "test",
		Libraries:        []typedb.ManifestLibrary{{Name: "Test", Output: output}},
	}); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, typeDBDir)
}
