package analyze

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA250DetectsUnsafeWorksheetAndRangeSelect(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub UnsafeWorksheet()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws.Select
End Sub

Public Sub UnsafeRange()
  Dim ws As Worksheet
  Dim rng As Range
  Set ws = ThisWorkbook.Worksheets(1)
  Set rng = ws.Range("A1")
  rng.Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 5, 13)
}

func TestVBA250TracksActivationAcrossBranches(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub UnknownActivation(ByVal choose As Boolean)
  Dim wb As Workbook
  Dim ws As Worksheet
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(1)
  If choose Then wb.Activate
  ws.Select
End Sub

Public Sub SameActivation(ByVal choose As Boolean)
  Dim wb As Workbook
  Dim ws As Worksheet
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(1)
  If choose Then
    wb.Activate
  Else
    wb.Activate
  End If
  ws.Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 8)
}

func TestVBA250HandlesWithBlocksAndAliases(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub UnsafeWith()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  With ws
    .Select
  End With
End Sub

Public Sub SafeWith()
  Dim wb As Workbook
  Dim ws As Worksheet
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(1)
  wb.Activate
  With ws
    .Select
    .Range("A1").Select
  End With
End Sub

Public Sub UnsafeAlias()
  Dim ws As Worksheet
  Dim aliasWs As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  Set aliasWs = ws
  aliasWs.Select
End Sub

Public Sub SafeAlias()
  Dim wb As Workbook
  Dim ws As Worksheet
  Dim aliasWs As Worksheet
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(1)
  Set aliasWs = ws
  wb.Activate
  aliasWs.Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 6, 27)
}

func TestVBA250InvalidatesActiveStateAfterUnknownCall(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim wb As Workbook
  Dim ws As Worksheet
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(1)
  wb.Activate
  ws.Select
  UnknownOperation
  ws.Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 10)
}

func TestVBA250DoesNotAssumeAdditiveWorksheetSelectReplacesSelection(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim wb As Workbook
  Dim ws As Worksheet
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(1)
  wb.Activate
  ws.Select False
  ws.Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 9)
}

func TestVBA250SelectReplacementParsing(t *testing.T) {
	t.Parallel()
	for _, test := range []struct {
		text     string
		replaces bool
	}{
		{text: "ws.Select", replaces: true},
		{text: "ws.Select True", replaces: true},
		{text: "ws.Select vbTrue", replaces: true},
		{text: "ws.Select -1", replaces: true},
		{text: "ws.Select (True)", replaces: true},
		{text: "ws.Select Replace:=True", replaces: true},
		{text: "ws.Select False", replaces: false},
		{text: "ws.Select (False)", replaces: false},
		{text: "ws.Select Replace:=False", replaces: false},
	} {
		operation, ok := parseActiveUIOperation(test.text)
		if !ok || operation.selectReplaces != test.replaces {
			t.Errorf("parseActiveUIOperation(%q) = %+v, %v; want selectReplaces=%v", test.text, operation, ok, test.replaces)
		}
	}
	if operation, ok := parseActiveUIOperation("With ws\n  .Select\nEnd With"); ok {
		t.Fatalf("parseActiveUIOperation(With block) = %+v, want no operation", operation)
	}
	for _, text := range []string{`Debug.Print "text.Select"`, `Rem ws.Activate`, `ws.Select ' text.Activate`} {
		if operation, ok := parseActiveUIOperation(text); ok && strings.EqualFold(operation.method, "activate") {
			t.Fatalf("parseActiveUIOperation(%q) = %+v, want no Activate operation", text, operation)
		}
	}
}

func TestVBA250BatchAndRealtimeParity(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Option Explicit
Public Sub Run()
  Dim wb As Workbook
  Dim ws As Worksheet
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(1)
  ws.Select
  wb.Activate
  ws.Select
  UnknownOperation
  ws.Select
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
	batchVBA250 := findingsByCode(batch, "VBA250")
	realtimeVBA250 := findingsByCode(realtime, "VBA250")
	assertVBA250Lines(t, batchVBA250, 7, 11)
	assertVBA250Lines(t, realtimeVBA250, 7, 11)
	if len(batchVBA250) != len(realtimeVBA250) {
		t.Fatalf("batch/realtime VBA250 findings = %+v / %+v, want equal", batchVBA250, realtimeVBA250)
	}
	for i := range batchVBA250 {
		if batchVBA250[i].Line != realtimeVBA250[i].Line || batchVBA250[i].Procedure != realtimeVBA250[i].Procedure {
			t.Fatalf("batch/realtime VBA250 findings differ: %+v / %+v", batchVBA250, realtimeVBA250)
		}
	}
}

func TestVBA250TracksStableRangePropertyParent(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  ThisWorkbook.Worksheets(1).Range("A1").Parent.Activate
  ThisWorkbook.Worksheets(1).Range("A1").Select
End Sub
`
	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings)
}

func TestVBA250ResolvesChainedWorksheetRangeFromInnermostMember(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  ThisWorkbook.Activate
  ThisWorkbook.Worksheets(1).Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 4)
}

func TestVBA250DoesNotReuseUnqualifiedRangeAfterWorksheetChanges(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim rng As Range
  Dim ws As Worksheet
  Set rng = Range("A1")
  Set ws = ThisWorkbook.Worksheets(2)
  ws.Activate
  rng.Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 8)
}

func TestVBA250LeavesNonExcelObjectsUnknown(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim obj As OtherProject.Worksheet
  obj.Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings)
}

func TestVBA250RecognizesApplicationRoots(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Application.ActiveWorkbook.Activate
  Application.Worksheets(1).Select
  Application.Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings)
}

func TestVBA250RecognizesMemberWhitespace(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws  .  Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 5)
}

func TestVBA250PreservesWhitespaceAndStableCollectionSelectors(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  ThisWorkbook . Activate
  ThisWorkbook . Worksheets(1) . Activate
  ThisWorkbook . Worksheets(1) . Range("A 1") . Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings)
}

func TestVBA250InvalidatesStateAfterUnsupportedActivate(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim win As Window
  Set win = Application.ActiveWindow
  ThisWorkbook.Activate
  ThisWorkbook.Worksheets(1).Activate
  win.Activate
  ThisWorkbook.Worksheets(1).Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 8)
}

func TestVBA250DistinguishesLiteralCollectionSelectors(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim ws1 As Worksheet
  Dim ws2 As Worksheet
  Set ws1 = ThisWorkbook.Worksheets("Sheet 1")
  Set ws2 = ThisWorkbook.Worksheets("Sheet1")
  ThisWorkbook.Activate
  ws1.Activate
  ws2.Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 9)
}

func TestVBA250ClearsAliasesAfterUnknownCall(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ThisWorkbook.Activate
  ws.Activate
  UnknownOperation ws
  ThisWorkbook.Activate
  ws.Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 9)
}

func TestVBA250DoesNotInvalidateStateForDebugPrint(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  ThisWorkbook.Activate
  ThisWorkbook.Worksheets(1).Activate
  Debug.Print "ok"
  ThisWorkbook.Worksheets(1).Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings)
}

func TestVBA250RefinesReverseActiveSheetComparison(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  If ActiveSheet Is ws Then
    ws.Range("A1").Select
  End If
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings)
}

func TestVBA250LeavesShadowedExcelRootsUnknown(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Private Function Range(ByVal name As String) As Object
  Set Range = Nothing
End Function

Public Sub Run()
  ThisWorkbook.Activate
  Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings)
}

func TestVBA250InvalidatesStateForUnknownExcelMemberReceiver(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  Dim obj As Object
  Set obj = CreateObject("Scripting.Dictionary")
  ThisWorkbook.Activate
  ThisWorkbook.Worksheets(1).Activate
  obj.Value = 1
  ThisWorkbook.Worksheets(1).Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 8)
}

func TestVBA250DoesNotTrustActivateUnderResumeNext(t *testing.T) {
	t.Parallel()
	source := `Option Explicit
Public Sub Run()
  On Error Resume Next
  ThisWorkbook.Activate
  ThisWorkbook.Worksheets(1).Activate
  On Error GoTo 0
  ThisWorkbook.Worksheets(1).Range("A1").Select
End Sub
`

	findings := runVBA250Batch(t, source)
	assertVBA250Lines(t, findings, 7)
}

func runVBA250Batch(t *testing.T, source string) []Finding {
	t.Helper()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	return findings
}

func assertVBA250Lines(t *testing.T, findings []Finding, want ...int) {
	t.Helper()
	got := findingsByCode(findings, "VBA250")
	if len(got) != len(want) {
		t.Fatalf("VBA250 findings = %+v, want lines %v", got, want)
	}
	seen := make(map[int]bool, len(got))
	for _, finding := range got {
		seen[finding.Line] = true
		if finding.Severity != "warning" {
			t.Errorf("VBA250 severity = %q, want warning: %+v", finding.Severity, finding)
		}
	}
	for _, line := range want {
		if !seen[line] {
			t.Errorf("VBA250 missing line %d: %+v", line, got)
		}
	}
}
