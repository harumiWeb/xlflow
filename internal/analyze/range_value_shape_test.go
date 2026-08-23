package analyze

import (
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestVBA226DetectsOneDimensionalRangeValueUse(t *testing.T) {
	t.Parallel()
	for _, member := range []string{"Value", "Value2"} {
		t.Run(member, func(t *testing.T) {
			dir := t.TempDir()
			source := `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1:A10").MEMBER
  Debug.Print values(1)
  Debug.Print UBound(values)
  Debug.Print values(1, 1)
End Sub
`
			source = strings.Replace(source, "MEMBER", member, 1)
			writeModule(t, dir, "Main.bas", source)

			findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
			if err != nil {
				t.Fatal(err)
			}
			got := findingsByCode(findings, "VBA226")
			if len(got) != 2 {
				t.Fatalf("VBA226 findings = %+v, want one-dimensional access and omitted-bound dimension", got)
			}
			if !strings.Contains(got[0].Message, "one array index") && !strings.Contains(got[1].Message, "one array index") {
				t.Fatalf("missing one-dimensional access diagnostic: %+v", got)
			}
		})
	}
}

func TestVBA226AcceptsSafeTwoDimensionalAccessAndPassThrough(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1:B10").Value2
  Debug.Print values(1, 1)
  Range("D1:E10").Value2 = values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 0 {
		t.Fatalf("safe two-dimensional access should not report: %+v", got)
	}
}

func TestVBA226DetectsSingleCellAndDynamicShapesButAcceptsGuardedArrayAccess(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal lastCell As String)
  Dim singleValue As Variant
  Dim dynamicValues As Variant
  singleValue = Range("A1").Value2
  Debug.Print singleValue
  Debug.Print singleValue(1)
  Debug.Print UBound(singleValue)
  dynamicValues = Range("A1:" & lastCell).Value2
  Debug.Print dynamicValues(1, 1)
  If IsArray(dynamicValues) Then
    Debug.Print dynamicValues(1, 1)
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA226")
	if len(got) != 3 {
		t.Fatalf("VBA226 findings = %+v, want scalar indexing, scalar bounds, and unguarded dynamic shape use", got)
	}
	for _, finding := range got {
		if finding.Line != 7 && finding.Line != 8 && finding.Line != 10 {
			t.Fatalf("unexpected VBA226 location: %+v", finding)
		}
	}
}

func TestVBA226DetectsIncompatibleDestinationShape(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1:B2").Value2
  Range("D1:D4").Value2 = values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA226")
	if len(got) != 1 || !strings.Contains(got[0].Message, "incompatible") {
		t.Fatalf("incompatible destination findings = %+v", got)
	}
}

func TestVBA226DetectsHorizontalAndRectangularIndexOrder(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim horizontal As Variant
  Dim rectangular As Variant
  Dim rowIndex As Long
  Dim columnIndex As Long
  horizontal = Range("A1:J1").Value2
  For rowIndex = 1 To 10
    Debug.Print horizontal(rowIndex, 1)
    Debug.Print horizontal(1, rowIndex)
  Next rowIndex
  rectangular = Range("A1:B10").Value2
  For rowIndex = 1 To 10
    For columnIndex = 1 To 2
      Debug.Print rectangular(columnIndex, rowIndex)
      Debug.Print rectangular(rowIndex, columnIndex)
    Next columnIndex
  Next rowIndex
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA226")
	if len(got) != 2 {
		t.Fatalf("horizontal/rectangular findings = %+v, want two provable dimension-order violations", got)
	}
	if got[0].Line >= got[1].Line {
		t.Fatalf("VBA226 findings are not source ordered: %+v", got)
	}
}

func TestVBA226RecognizesArrayGuardAndConservativeBranchMerge(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal useArray As Boolean)
  Dim values As Variant
  If useArray Then
    values = Range("A1:B2").Value2
  Else
    values = Range("A1").Value2
  End If
  If Not IsArray(values) Then Exit Sub
  Debug.Print values(1, 1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 0 {
		t.Fatalf("dominant IsArray guard should accept the two-dimensional access: %+v", got)
	}

	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal useArray As Boolean)
  Dim values As Variant
  If useArray Then
    values = Range("A1:B2").Value2
  Else
    values = Range("A1").Value2
  End If
  Debug.Print values(1, 1)
End Sub
`)
	findings, err = (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 1 {
		t.Fatalf("branch-merged scalar/array shape should be uncertain: %+v", got)
	}
}

func TestVBA226ClearsArrayGuardAfterReassignment(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal lastCell As String)
  Dim values As Variant
  values = Range("A1:B2").Value2
  If IsArray(values) Then
    values = Range("A1:" & lastCell).Value2
    Debug.Print values(1, 1)
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 1 {
		t.Fatalf("reassigned uncertain value should not retain the old IsArray guard: %+v", got)
	}
}

func TestVBA226HonorsConfigurationAndInlineSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1:A10").Value2
  ' xlflow:disable-next-line VBA226
  Debug.Print values(1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 0 {
		t.Fatalf("inline suppression should remove VBA226: %+v", got)
	}

	cfg := config.Default()
	cfg.Analyze.DetectRangeValueArrayShape = false
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 0 {
		t.Fatalf("disabled VBA226 should not report: %+v", got)
	}
}

func TestVBA226BatchAndRealtimeResultsMatch(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1:A10").Value2
  Debug.Print values(1)
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
	if got, want := findingsByCode(realtime, "VBA226"), findingsByCode(batch, "VBA226"); !reflect.DeepEqual(got, want) {
		t.Fatalf("batch/realtime VBA226 findings differ:\nbatch=%+v\nrealtime=%+v", want, got)
	}
}

func TestVBA226FallsBackToSourceLinesForEmptyProcedureIR(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1").Value2
  Debug.Print values(1)
End Sub
`
	file := parsedFile{
		Path:   "Main.bas",
		Lines:  normalizedSourceLines(source),
		Source: []byte(source),
	}
	proc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 6}
	if !rangeValueShapeApplicable(file, proc) {
		t.Fatal("empty procedure projection should use source-line applicability")
	}
	findings := (Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(file, proc)
	got := findingsByCode(findings, "VBA226")
	if len(got) != 1 || got[0].Line != 5 || !strings.Contains(got[0].Message, "Single-cell") {
		t.Fatalf("source-line fallback findings = %+v, want one scalar-index finding on line 5", got)
	}

	guardedSource := `Option Explicit
Public Sub Run(ByVal lastCell As String)
  Dim values As Variant
	values = Range("A1:" & lastCell).Value2
	If IsArray(values) Then
	    If lastCell <> "" Then
	      Debug.Print values(1, 1)
	    End If
	    Debug.Print values(1, 1)
	End If
End Sub
`
	guardedFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(guardedSource), Source: []byte(guardedSource)}
	guardedProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 10}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(guardedFile, guardedProc), "VBA226"); len(got) != 0 {
		t.Fatalf("source-line IsArray guard should suppress uncertain two-dimensional access: %+v", got)
	}

	negatedSource := `Option Explicit
Public Sub Run(ByVal lastCell As String)
  Dim values As Variant
  values = Range("A1:" & lastCell).Value2
  If Not IsArray(values) Then
    Debug.Print values
  Else
    Debug.Print values(1, 1)
  End If
End Sub
`
	negatedFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(negatedSource), Source: []byte(negatedSource)}
	negatedProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 9}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(negatedFile, negatedProc), "VBA226"); len(got) != 0 {
		t.Fatalf("negative IsArray guard Else branch should suppress uncertain two-dimensional access: %+v", got)
	}

	inlineSource := `Option Explicit
Public Sub Run(ByVal lastCell As String)
  Dim values As Variant
  values = Range("A1:" & lastCell).Value2
  If IsArray(values) Then Debug.Print values(1, 1)
End Sub
`
	inlineFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(inlineSource), Source: []byte(inlineSource)}
	inlineProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 6}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(inlineFile, inlineProc), "VBA226"); len(got) != 0 {
		t.Fatalf("single-line IsArray guard should suppress uncertain two-dimensional access: %+v", got)
	}

	branchSource := `Option Explicit
Public Sub Run(ByVal useArray As Boolean)
  Dim values As Variant
  values = Range("A1").Value2
  If useArray Then
    values = Range("A1:B2").Value2
  End If
  Debug.Print values(1, 1)
End Sub
`
	branchFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(branchSource), Source: []byte(branchSource)}
	branchProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 9}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(branchFile, branchProc), "VBA226"); len(got) != 1 {
		t.Fatalf("empty-IR branch merge should retain the scalar path: %+v", got)
	}

	partialSource := `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1").Value2
  Debug.Print values(1)
End Sub
`
	partialFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(partialSource), Source: []byte(partialSource)}
	partialProc := sourceProcedure{
		Name:       "Run",
		StartLine:  2,
		EndLine:    6,
		Statements: []procedureir.Statement{{Text: "Debug.Print values(1)"}},
		Features:   procedureFeatureSet{unknown: featureRangeArray},
	}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(partialFile, partialProc), "VBA226"); len(got) != 1 {
		t.Fatalf("unknown partial projection should use source fallback: %+v", got)
	}

	commentSource := `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1").Value2
  Rem values(1)
End Sub
`
	commentFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(commentSource), Source: []byte(commentSource)}
	commentProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 6}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(commentFile, commentProc), "VBA226"); len(got) != 0 {
		t.Fatalf("Rem comments must be ignored by source fallback: %+v", got)
	}
	colonCommentSource := `Option Explicit
Public Sub Run()
  Dim values As Variant
  values = Range("A1").Value2
  values = values: Rem values(1): values(1)
End Sub
`
	colonCommentFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(colonCommentSource), Source: []byte(colonCommentSource)}
	colonCommentProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 6}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(colonCommentFile, colonCommentProc), "VBA226"); len(got) != 0 {
		t.Fatalf("colon-separated Rem comments must ignore the rest of the line: %+v", got)
	}

	earlyExitSource := `Option Explicit
Public Sub Run(ByVal lastCell As String)
  Dim values As Variant
  values = Range("A1:" & lastCell).Value2
  If Not IsArray(values) Then Exit Sub
  Debug.Print values(1, 1)
End Sub
`
	earlyExitFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(earlyExitSource), Source: []byte(earlyExitSource)}
	earlyExitProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 7}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(earlyExitFile, earlyExitProc), "VBA226"); len(got) != 0 {
		t.Fatalf("early-exit negative IsArray guard should preserve the surviving array path: %+v", got)
	}

	blockEarlyExitSource := `Option Explicit
Public Sub Run(ByVal lastCell As String)
  Dim values As Variant
  values = Range("A1:" & lastCell).Value2
  If Not IsArray(values) Then
    Exit Sub
  End If
  Debug.Print values(1, 1)
End Sub
`
	blockEarlyExitFile := parsedFile{Path: "Main.bas", Lines: normalizedSourceLines(blockEarlyExitSource), Source: []byte(blockEarlyExitSource)}
	blockEarlyExitProc := sourceProcedure{Name: "Run", StartLine: 2, EndLine: 8}
	if got := findingsByCode((Analyzer{RootDir: ".", Config: config.Default()}).rangeValueShapeFindings(blockEarlyExitFile, blockEarlyExitProc), "VBA226"); len(got) != 0 {
		t.Fatalf("block early-exit negative IsArray guard should preserve the surviving array path: %+v", got)
	}
}

func TestVBA226TracksOnlyRangeValueOrigins(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByRef source() As Byte, ByVal text As String)
  Dim bytes() As Byte
  Dim ansiBytes() As Byte
  Dim index As Long
  bytes = source
  ReDim ansiBytes(UBound(bytes))
  For index = LBound(bytes) To UBound(bytes)
    ansiBytes(index) = bytes(index)
  Next index
  bytes = text
  Debug.Print LBound(bytes), UBound(bytes), bytes(0)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 0 {
		t.Fatalf("ordinary arrays and string-to-byte assignments are not Range.Value origins: %+v", got)
	}
}

func TestVBA226UsesDirectNestedCellsReceiverShape(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal ws As Worksheet, ByVal lastRow As Long)
  Dim values As Variant
  Dim nestedValues As Variant
  Dim rowIndex As Long
  values = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 14)).Value2
  If IsArray(values) Then
    For rowIndex = 1 To UBound(values, 1)
      Debug.Print values(rowIndex, 1)
    Next rowIndex
  End If
  nestedValues = ws.Range(ws.Range("A1"), ws.Range("B2")).Value2
  If IsArray(nestedValues) Then
    Debug.Print nestedValues(1, 1)
  End If
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 0 {
		t.Fatalf("outer Range receiver must not be classified as its first nested Cells call: %+v", got)
	}
}

func TestVBA226PreservesStructuralRangeAliases(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run(ByVal ws As Worksheet)
  Dim sourceRange As Range
  Dim values As Variant
  Set sourceRange = ws.Range("A1:B2")
  values = sourceRange.Value2
  Debug.Print values(1, 1)
  ws.Range("D1:E2").Value2 = values
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA226"); len(got) != 0 {
		t.Fatalf("structural Range aliases should preserve their exact shape: %+v", got)
	}
}

func TestVBA226RecognizesProvablyMultiCellDynamicRange(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Const COLUMN_COUNT As Long = 14

Public Sub SafeRun(ByVal ws As Worksheet, ByVal lastRow As Long)
  Dim values As Variant
  Dim rowIndex As Long
  values = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, COLUMN_COUNT)).Value2
  For rowIndex = 1 To UBound(values, 1)
    Debug.Print values(rowIndex, 1)
  Next rowIndex
End Sub

Public Sub UncertainRun(ByVal ws As Worksheet, ByVal lastRow As Long)
  Dim values As Variant
  values = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).Value2
  Debug.Print values(1, 1)
End Sub

Public Sub ScalarRun(ByVal ws As Worksheet)
  Dim value As Variant
  value = ws.Range(ws.Cells(2, 1), ws.Cells(2, 1)).Value2
  Debug.Print value(1, 1)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA226")
	counts := map[string]int{}
	for _, finding := range got {
		counts[finding.Procedure]++
	}
	if counts["UncertainRun"] != 1 || counts["ScalarRun"] != 1 || len(counts) != 2 {
		t.Fatalf("VBA226 findings per procedure = %v, want one UncertainRun and one ScalarRun: %+v", counts, got)
	}
}

func TestVBA226PartialShapeSuggestionOmitsUnknownExtent(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Const COLUMN_COUNT As Long = 14

Public Sub Run(ByVal ws As Worksheet, ByVal lastRow As Long)
  Dim values As Variant
  values = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, COLUMN_COUNT)).Value2
  Debug.Print values(1, 15)
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA226")
	if len(got) != 1 {
		t.Fatalf("partial-shape findings = %+v, want one known-column violation", got)
	}
	if got[0].Suggestion != "Keep the second index within 1..14." || strings.Contains(got[0].Suggestion, "1..0") {
		t.Fatalf("partial-shape suggestion = %q", got[0].Suggestion)
	}
}
