package intel

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestDiagnosticsHandlesMalformedSourceAndJapaneseText(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		URI:    "file:///C:/work/%E6%97%A5%E6%9C%AC/Main.bas",
		Path:   filepath.Join(t.TempDir(), "日本", "Main.bas"),
		Source: "Sub Broken()\n    Debug.Print \"日本語\"\n",
	}

	diagnostics := analyzer.Diagnostics(doc)
	if len(diagnostics) == 0 {
		t.Fatal("expected diagnostics for malformed source")
	}
	for _, diag := range diagnostics {
		if diag.Source != "xlflow" {
			t.Fatalf("diagnostic source = %q, want xlflow", diag.Source)
		}
		if diag.Range.Start.Character < 0 || diag.Range.End.Character < 0 {
			t.Fatalf("negative range in diagnostic: %+v", diag)
		}
	}
}

func TestDiagnosticsUseSharedLintRulesForUnsavedContent(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Range("A1").Select
End Sub
`,
	}

	diagnostics := analyzer.Diagnostics(doc)
	if !hasDiagnostic(diagnostics, "VB002") {
		t.Fatalf("VB002 diagnostic missing: %+v", diagnostics)
	}

	doc.Source = `Option Explicit
Public Sub Run()
    missingValue = 1
End Sub
`
	diagnostics = analyzer.Diagnostics(doc)
	if !hasDiagnostic(diagnostics, "VB029") {
		t.Fatalf("VB029 diagnostic missing: %+v", diagnostics)
	}

	doc.Source = `Option Explicit
Public Sub Run()
    ?? "hoge"
End Sub
`
	diagnostics = analyzer.Diagnostics(doc)
	if !hasDiagnostic(diagnostics, "VB032") {
		t.Fatalf("VB032 diagnostic missing: %+v", diagnostics)
	}

	doc.Source = `Option Explicit
Public Sub Run()
    Dim missingValue As Long
    missingValue = 1
End Sub
`
	doc.Source = `Attribute VB_Name = "Main"
` + doc.Source
	if diagnostics := analyzer.Diagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("expected undeclared diagnostic to clear, got %+v", diagnostics)
	}

	doc.Source = `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Range("A1").Value = 1
End Sub
`
	if diagnostics := analyzer.Diagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("expected diagnostics to clear, got %+v", diagnostics)
	}
}

func TestDiagnosticsConvertLintByteColumnsToUTF16AfterJapaneseText(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	line := `    Debug.Print "😀日本": Range("A1").Select`
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: "Option Explicit\nPublic Sub Run()\n" +
			line + "\n" +
			"End Sub\n",
	}

	diagnostics := analyzer.Diagnostics(doc)
	var selectDiag *Diagnostic
	for i := range diagnostics {
		if diagnostics[i].Code == "VB002" {
			selectDiag = &diagnostics[i]
			break
		}
	}
	if selectDiag == nil {
		t.Fatalf("VB002 diagnostic missing: %+v", diagnostics)
	}
	want := utf16Len(line[:strings.Index(line, "Select")])
	if selectDiag.Range.Start.Line != 2 || selectDiag.Range.Start.Character != want {
		t.Fatalf("VB002 range start = %+v, want line 2 character %d", selectDiag.Range.Start, want)
	}
}

func TestSignatureHelpResolvesBuiltinMembersAndActiveParameter(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim ws As Worksheet
    ws.Range("A1",
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := `    ws.Range("A1",`

	help, err := analyzer.SignatureHelp(doc, Position{Line: 3, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if got := help.Signatures[0].Label; got != "Excel.Worksheet.Range(Cell1 As Variant, Optional Cell2 As Variant) As Excel.Range" {
		t.Fatalf("signature label = %q", got)
	}
	if help.ActiveParameter != 1 {
		t.Fatalf("active parameter = %d, want 1", help.ActiveParameter)
	}
}

func TestSignatureHelpResolvesParenlessMemberCall(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim dict As Scripting.Dictionary
    dict.Add "A",
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := `    dict.Add "A",`

	help, err := analyzer.SignatureHelp(doc, Position{Line: 3, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if got := help.Signatures[0].Label; got != "Scripting.Dictionary.Add(Key As Variant, Item As Variant) As void" {
		t.Fatalf("signature label = %q", got)
	}
	if help.ActiveParameter != 1 {
		t.Fatalf("active parameter = %d, want 1", help.ActiveParameter)
	}
}

func TestSignatureHelpResolvesParenlessMemberCallAfterSpace(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := "Option Explicit\nSub Test()\n    Dim dict As Scripting.Dictionary\n    dict.Add \nEnd Sub\n"
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := "    dict.Add "

	help, err := analyzer.SignatureHelp(doc, Position{Line: 3, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if got := help.Signatures[0].Label; got != "Scripting.Dictionary.Add(Key As Variant, Item As Variant) As void" {
		t.Fatalf("signature label = %q", got)
	}
	if help.ActiveParameter != 0 {
		t.Fatalf("active parameter = %d, want 0", help.ActiveParameter)
	}
}

func TestSignatureHelpResolvesProjectFunctionAndNamedArgument(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Function Foo(name As String, count As Long) As Boolean
End Function
Sub Test()
    Foo(name:=
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := `    Foo(name:=`

	help, err := analyzer.SignatureHelp(doc, Position{Line: 4, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if !strings.Contains(help.Signatures[0].Label, "Foo(name As String, count As Long) As Boolean") {
		t.Fatalf("signature label = %q", help.Signatures[0].Label)
	}
	if help.ActiveParameter != 0 {
		t.Fatalf("active parameter = %d, want named argument index 0", help.ActiveParameter)
	}
}

func TestSignatureHelpResolvesVBABuiltinFunction(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    MsgBox "Hello",
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := `    MsgBox "Hello",`

	help, err := analyzer.SignatureHelp(doc, Position{Line: 2, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if got := help.Signatures[0].Label; got != "VBA.Global.MsgBox(Prompt As Variant, Optional Buttons As VbMsgBoxStyle, Optional Title As Variant, Optional HelpFile As Variant, Optional Context As Variant) As VbMsgBoxResult" {
		t.Fatalf("signature label = %q", got)
	}
	if help.ActiveParameter != 1 {
		t.Fatalf("active parameter = %d, want 1", help.ActiveParameter)
	}
}

func TestSignatureHelpResolvesFormattedNestedCallTarget(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Debug.Print IIf(True, MsgBox(
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := `    Debug.Print IIf(True, MsgBox(`

	help, err := analyzer.SignatureHelp(doc, Position{Line: 2, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one MsgBox signature", help)
	}
	if got := help.Signatures[0].Label; !strings.HasPrefix(got, "VBA.Global.MsgBox(") {
		t.Fatalf("signature label = %q, want MsgBox signature", got)
	}
}

func TestDiagnosticsTreatErrAsBuiltinGlobal(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Sub Test(expected As Variant, actual As Variant)
    If IsObject(expected) Or IsObject(actual) Then
        Err.Raise vbObjectError + 514, "XlflowAssert.AssertEquals", "AssertEquals supports scalar values only. Compare object properties such as Range.Value2."
    End If
End Sub
`,
	}

	diagnostics := analyzer.Diagnostics(doc)
	if hasDiagnosticMessage(diagnostics, `Undeclared identifier "Err"`) {
		t.Fatalf("Err should be treated as a built-in global, got %+v", diagnostics)
	}

	line := `        Err.Raise vbObjectError + 514, "XlflowAssert.AssertEquals",`
	help, err := analyzer.SignatureHelp(doc, Position{Line: 3, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one Err.Raise signature", help)
	}
	if got := help.Signatures[0].Label; got != "VBA.ErrObject.Raise(Number As Long, Optional Source As String, Optional Description As String, Optional HelpFile As String, Optional HelpContext As Long) As void" {
		t.Fatalf("signature label = %q", got)
	}
}

func TestSignatureHelpResolvesRangeFindNamedArgument(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim rng As Range
    rng.Find(What:="A", LookAt:=
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := `    rng.Find(What:="A", LookAt:=`

	help, err := analyzer.SignatureHelp(doc, Position{Line: 3, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if !strings.Contains(help.Signatures[0].Label, "Excel.Range.Find(What As Variant, Optional After As Variant, Optional LookIn As Variant, Optional LookAt As Variant") {
		t.Fatalf("signature label = %q", help.Signatures[0].Label)
	}
	if help.ActiveParameter != 3 {
		t.Fatalf("active parameter = %d, want named argument index 3", help.ActiveParameter)
	}
}

func TestSignatureHelpResolvesExcelNamedArgument(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Application.Workbooks.Open(Filename:=path, ReadOnly:=
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}
	line := `    Application.Workbooks.Open(Filename:=path, ReadOnly:=`

	help, err := analyzer.SignatureHelp(doc, Position{Line: 2, Character: utf16Len(line)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if !strings.Contains(help.Signatures[0].Label, "Excel.Workbooks.Open(Filename As String, Optional UpdateLinks As Variant, Optional ReadOnly As Variant") {
		t.Fatalf("signature label = %q", help.Signatures[0].Label)
	}
	if help.ActiveParameter != 2 {
		t.Fatalf("active parameter = %d, want named argument index 2", help.ActiveParameter)
	}
}

func TestDiagnosticsIncludeArgumentCountAndNamedArgumentWarnings(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Sub Test()
    Dim dict As Scripting.Dictionary
    dict.Add "A"
    Range()
    dict.Add Key:="A", Value:=1
End Sub
`,
	}

	diagnostics := analyzer.Diagnostics(doc)
	vb030 := diagnosticsByCode(diagnostics, "VB030")
	if len(vb030) < 3 {
		t.Fatalf("expected argument diagnostics, got %+v", diagnostics)
	}
	if !hasDiagnosticMessage(vb030, "expects at least 2 argument") {
		t.Fatalf("missing Dictionary.Add argument count diagnostic: %+v", vb030)
	}
	if !hasDiagnosticMessage(vb030, "Range expects at least 1 argument") {
		t.Fatalf("missing Range argument count diagnostic: %+v", vb030)
	}
	if !hasDiagnosticMessage(vb030, "Unknown named argument: Value") {
		t.Fatalf("missing unknown named argument diagnostic: %+v", vb030)
	}
}

func TestDiagnosticsIncludeParenlessCallAfterSpace(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:   filepath.Join(t.TempDir(), "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    Dim dict As Scripting.Dictionary\n    dict.Add \nEnd Sub\n",
	}

	diagnostics := analyzer.Diagnostics(doc)
	vb030 := diagnosticsByCode(diagnostics, "VB030")
	if !hasDiagnosticMessage(vb030, "expects at least 2 argument") {
		t.Fatalf("missing parenless empty argument diagnostic: %+v", diagnostics)
	}
}

func TestArgumentDiagnosticsIgnoreDeclarationsAndKeepControlFlowCalls(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Private Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Property Get Title() As String
    Title = "ok"
End Property

Sub Run()
    Dim rng As Range
    If Len("abc") > 0 Then
        rng.Find()
    End If
End Sub
`,
	}

	diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB030")
	if len(diagnostics) != 1 {
		t.Fatalf("VB030 diagnostics = %+v, want only Range.Find argument warning", diagnostics)
	}
	if !strings.Contains(diagnostics[0].Message, "Find expects at least 1 argument") {
		t.Fatalf("VB030 diagnostic = %+v, want Range.Find argument warning", diagnostics[0])
	}
	if diagnostics[0].Range.Start.Line != 10 {
		t.Fatalf("VB030 range = %+v, want Range.Find line", diagnostics[0].Range)
	}
}

func TestIDESmokeCoversCompletionHoverSignatureAndDiagnostics(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Smoke()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "A"
    Dim rng As Range
    rng.Find()
    rng.Find(What:="A", LookAt:=
    rng.Font.Co
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Smoke.bas"), Source: source}

	hoverLine := `    Set dict = CreateObject("Scripting.Dictionary")`
	hover, err := analyzer.Hover(doc, Position{Line: 3, Character: utf16Len(hoverLine[:strings.Index(hoverLine, "dict")+2])}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "dict As Scripting.Dictionary") {
		t.Fatalf("unexpected dict hover: %+v", hover)
	}

	signatureLine := `    rng.Find(What:="A", LookAt:=`
	help, err := analyzer.SignatureHelp(doc, Position{Line: 7, Character: utf16Len(signatureLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 || help.ActiveParameter != 3 {
		t.Fatalf("signature help = %+v, want Range.Find LookAt parameter", help)
	}

	completionLine := `    rng.Font.Co`
	items, err := analyzer.Completions(doc, Position{Line: 8, Character: utf16Len(completionLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Color") {
		t.Fatalf("Font.Color completion missing: %+v", items)
	}

	diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB030")
	if !hasDiagnosticMessage(diagnostics, "Add expects at least 2 argument") || !hasDiagnosticMessage(diagnostics, "Find expects at least 1 argument") {
		t.Fatalf("expected Dictionary.Add and Range.Find argument diagnostics, got %+v", diagnostics)
	}
}

func TestE2ESmokeMemberCompletionsAfterInferredTypes(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit

Private Sub Sample()
    Dim wb As Excel.Workbook
    Set wb = Application.Workbooks.Open("sample.txt")

    Dim ws As Excel.Worksheet
    Set ws = wb.Worksheets("Input")

    Dim rng As Excel.Range
    Set rng = ws.Range("A1").Resize(10,3)

    Dim lo As Excel.ListObject
    Set lo = ws.ListObjects("SalesTable")

    Dim amountRange As Excel.Range
    Set amountRange = lo.ListColumns("Amount").DataBodyRange

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    dict.
    amountRange.
    rs.
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Smoke.bas"), Source: source}
	cases := []struct {
		lineNo int
		line   string
		want   string
	}{
		{24, `    dict.`, "Exists"},
		{25, `    amountRange.`, "Cells"},
		{26, `    rs.`, "Fields"},
	}
	for _, tc := range cases {
		items, err := analyzer.Completions(doc, Position{Line: tc.lineNo, Character: utf16Len(tc.line)}, []Document{doc})
		if err != nil {
			t.Fatal(err)
		}
		if !hasCompletion(items, tc.want) {
			t.Fatalf("%s completion missing %s: %+v", strings.TrimSpace(tc.line), tc.want, items)
		}
	}
}

func TestDiagnosticsReportOutOfScopeMemberReceiver(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Smoke.bas"),
		Source: `Option Explicit

Private Sub Sample()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Workbook", "Book1"
End Sub

Private Sub DiagnosticSmoke()
    dict.Add "OnlyKey"
End Sub
`,
	}

	diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB029")
	if !hasDiagnosticMessage(diagnostics, `Undeclared identifier "dict"`) {
		t.Fatalf("missing out-of-scope dict diagnostic: %+v", diagnostics)
	}
}

func TestDiagnosticsDoNotReportExcelRangeCallChainArgumentsAsUndeclared(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Sub TestLastRow()
    Dim ws As Worksheet
    Dim lastRowA As Long

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowA = Cells(ws.Rows.Count, 1).End(xlUp).Row
    Debug.Print ws.Range("A1").Value
    Debug.Print ws.Range("A1:B10").Rows.Count
    Debug.Print ThisWorkbook.Worksheets("Sheet1").Cells(1, "A").Value
End Sub
`,
	}

	diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB029")
	if len(diagnostics) != 0 {
		t.Fatalf("expected no undeclared diagnostics for Excel member chains, got %+v", diagnostics)
	}
}

func TestE2ESmokeNamespaceBuiltinAndWithSignatureCompletions(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit

Private Sub Sample()
    Dim wb As Excel.
    Dim rng As Excel.Range
    With rng
        .Offset(
        .Offset(1,0).
        .Font.Bold = Tr
        .Value = No
    End With
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    lblStatus.Caption = CS
    lblStatus.Caption = CStr(dict.
    ts.WriteLine CStr(dict.
    Deb
End Sub

Private Function BuildMessage(ByVal title As String, By
End Function
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Smoke.bas"), Source: source}

	cases := []struct {
		lineNo int
		line   string
		want   string
	}{
		{3, `    Dim wb As Excel.`, "Workbook"},
		{7, `        .Offset(1,0).`, "Value"},
		{8, `        .Font.Bold = Tr`, "True"},
		{9, `        .Value = No`, "Now"},
		{13, `    lblStatus.Caption = CS`, "CStr"},
		{14, `    lblStatus.Caption = CStr(dict.`, "Item"},
		{15, `    ts.WriteLine CStr(dict.`, "Count"},
		{16, `    Deb`, "Debug.Print"},
		{19, `Private Function BuildMessage(ByVal title As String, By`, "ByVal"},
	}
	for _, tc := range cases {
		items, err := analyzer.Completions(doc, Position{Line: tc.lineNo, Character: utf16Len(tc.line)}, []Document{doc})
		if err != nil {
			t.Fatal(err)
		}
		if !hasCompletion(items, tc.want) {
			t.Fatalf("%s completion missing %s: %+v", strings.TrimSpace(tc.line), tc.want, items)
		}
	}

	offsetLine := `        .Offset(`
	help, err := analyzer.SignatureHelp(doc, Position{Line: 6, Character: utf16Len(offsetLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 || !strings.Contains(help.Signatures[0].Label, "Excel.Range.Offset(Optional RowOffset As Variant, Optional ColumnOffset As Variant)") {
		t.Fatalf("Offset signature help = %+v", help)
	}
}

func TestExcelIdiomsSignatureHelpAndNamedArgumentCompletions(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit

Public Sub TestWithBlockInference()
    Dim ws As Excel.Worksheet
    Set ws = ThisWorkbook.Worksheets(1)

    With ws.Range("A1:D10")
        .Offset(1,0).Resize(
    End With

    With ThisWorkbook.Worksheets(1).ListObjects(
        .ListColumns(1).DataBodyRange.Font.Bold = True
    End With
End Sub

Public Sub TestSignatureHelp()
    Dim wb As Excel.Workbook
    Set wb = Workbooks.Open( _
    Filename:="sample.xlsx", _
    ReadOnly:=True, _
    AddToMru:=False _
    )
End Sub

Public Sub TestCommonExcelIdioms()
    Dim ws As Worksheet
    Dim r As Range
    Set ws = ThisWorkbook.Worksheets("Data")
    Set r = ws.Range(ws.Cells(1,1), ws.Cells(10, 5))

    r.AutoFilter Field:=1, Criteria1:="<>"
    r.Sort O

    ws.ListObjects("Table1").DataBodyRange.Copy D

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Idioms.bas"), Source: source}

	resizeLine := `        .Offset(1,0).Resize(`
	help, err := analyzer.SignatureHelp(doc, Position{Line: 7, Character: utf16Len(resizeLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 || !strings.Contains(help.Signatures[0].Label, "Excel.Range.Resize(Optional RowSize As Variant, Optional ColumnSize As Variant)") {
		t.Fatalf("Resize signature help = %+v", help)
	}

	listObjectsLine := `    With ThisWorkbook.Worksheets(1).ListObjects(`
	help, err = analyzer.SignatureHelp(doc, Position{Line: 10, Character: utf16Len(listObjectsLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 || !strings.Contains(help.Signatures[0].Label, "Excel.ListObjects.Item(Index As Variant) As Excel.ListObject") {
		t.Fatalf("ListObjects default member signature help = %+v", help)
	}

	readOnlyLine := `    ReadOnly:=True, _`
	help, err = analyzer.SignatureHelp(doc, Position{Line: 19, Character: utf16Len(readOnlyLine[:strings.Index(readOnlyLine, "True")])}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 || !strings.Contains(help.Signatures[0].Label, "Excel.Workbooks.Open(Filename As String, Optional UpdateLinks As Variant, Optional ReadOnly As Variant") || help.ActiveParameter != 2 {
		t.Fatalf("multi-line Workbooks.Open signature help = %+v", help)
	}

	sortLine := `    r.Sort O`
	items, err := analyzer.Completions(doc, Position{Line: 31, Character: utf16Len(sortLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Order1:=") {
		t.Fatalf("Sort named argument completion missing Order1:= %+v", items)
	}

	copyLine := `    ws.ListObjects("Table1").DataBodyRange.Copy D`
	items, err = analyzer.Completions(doc, Position{Line: 33, Character: utf16Len(copyLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Destination:=") {
		t.Fatalf("Copy named argument completion missing Destination:= %+v", items)
	}

	shellLine := `    shell.`
	items, err = analyzer.Completions(doc, Position{Line: 37, Character: utf16Len(shellLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Run") {
		t.Fatalf("WScript.Shell completion missing Run: %+v", items)
	}

	reLine := `    re.`
	items, err = analyzer.Completions(doc, Position{Line: 41, Character: utf16Len(reLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Pattern") || !hasCompletion(items, "Test") {
		t.Fatalf("VBScript.RegExp completion missing Pattern/Test: %+v", items)
	}
}

func TestDocumentSymbolsUseUnsavedDocumentContent(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub RunReport()
End Sub
`,
	}

	symbols, err := analyzer.DocumentSymbols(doc)
	if err != nil {
		t.Fatal(err)
	}
	if !hasSymbol(symbols, "RunReport") {
		t.Fatalf("RunReport not found in symbols: %+v", symbols)
	}
}

func TestRunnableProceduresFiltersCodeLensTargets(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		URI:        "file:///C:/work/src/modules/Main.bas",
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub RunReport()
End Sub
Private Sub HiddenRunner()
End Sub
Public Sub WithArg(ByVal value As Long)
End Sub
Public Function Build() As String
End Function
Public Property Get Title() As String
End Property
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Public Sub Test_RunReport()
End Sub
Public Sub Totals_Test()
End Sub
Public Sub totals_test()
End Sub
Public Sub TEST_Total()
End Sub
`,
	}

	procedures, err := analyzer.RunnableProcedures(doc, DefaultCodeLensConfig())
	if err != nil {
		t.Fatal(err)
	}
	if got := runnableProcedureNames(procedures); !reflect.DeepEqual(got, []string{"RunReport:sub", "HiddenRunner:sub", "Test_RunReport:test", "Totals_Test:test", "totals_test:test", "TEST_Total:test"}) {
		t.Fatalf("runnable procedures = %#v", got)
	}
	for _, procedure := range procedures {
		if procedure.URI != doc.URI || procedure.ModuleName != "Main" || !strings.HasPrefix(procedure.QualifiedName, "Main.") {
			t.Fatalf("unexpected runnable procedure metadata: %+v", procedure)
		}
	}

	noTestsCfg := DefaultCodeLensConfig()
	noTestsCfg.RunTests = false
	procedures, err = analyzer.RunnableProcedures(doc, noTestsCfg)
	if err != nil {
		t.Fatal(err)
	}
	if got := runnableProcedureNames(procedures); !reflect.DeepEqual(got, []string{"RunReport:sub", "HiddenRunner:sub"}) {
		t.Fatalf("runnable procedures without tests = %#v", got)
	}

	disabledCfg := DefaultCodeLensConfig()
	disabledCfg.Enabled = false
	procedures, err = analyzer.RunnableProcedures(doc, disabledCfg)
	if err != nil {
		t.Fatal(err)
	}
	if len(procedures) != 0 {
		t.Fatalf("disabled runnable procedures = %+v, want none", procedures)
	}
}

func TestRunnableProceduresUserFormEventsAreConfigurable(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "CustomerForm.frm"),
		ModuleKind: "form",
		Source: `VERSION 5.00
Begin VB.UserForm CustomerForm
End
Attribute VB_Name = "CustomerForm"
Option Explicit
Private Sub UserForm_Initialize()
End Sub
Private Sub cmdOK_Click()
End Sub
Private Sub cmd_OK_Click()
End Sub
Public Sub ShowForTest()
End Sub
Public Sub Test_Form()
End Sub
`,
	}

	procedures, err := analyzer.RunnableProcedures(doc, DefaultCodeLensConfig())
	if err != nil {
		t.Fatal(err)
	}
	if got := runnableProcedureNames(procedures); !reflect.DeepEqual(got, []string{"ShowForTest:sub", "Test_Form:test"}) {
		t.Fatalf("default form runnable procedures = %#v", got)
	}

	cfg := DefaultCodeLensConfig()
	cfg.UserFormEvents = true
	procedures, err = analyzer.RunnableProcedures(doc, cfg)
	if err != nil {
		t.Fatal(err)
	}
	if got := runnableProcedureNames(procedures); !reflect.DeepEqual(got, []string{"UserForm_Initialize:sub", "cmdOK_Click:sub", "cmd_OK_Click:sub", "ShowForTest:sub", "Test_Form:test"}) {
		t.Fatalf("form runnable procedures with events = %#v", got)
	}
}

func TestDocumentSymbolsSkipEmptyNamesFromIncompleteSource(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     "Option Explicit\nPublic Sub \n    Dim \nEnd Sub\n",
	}

	symbols, err := analyzer.DocumentSymbols(doc)
	if err != nil {
		t.Fatal(err)
	}
	for _, sym := range symbols {
		if strings.TrimSpace(sym.Name) == "" {
			t.Fatalf("empty symbol name should not be returned: %+v", symbols)
		}
		if !rangeContains(sym.Range, sym.Selection) {
			t.Fatalf("symbol selection range must be contained in full range: %+v", sym)
		}
	}
}

func TestDocumentSymbolsKeepSelectionInsideRangeForBareIncompleteSource(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     "P",
	}

	symbols, err := analyzer.DocumentSymbols(doc)
	if err != nil {
		t.Fatal(err)
	}
	for _, sym := range symbols {
		if !rangeContains(sym.Range, sym.Selection) {
			t.Fatalf("symbol selection range must be contained in full range: %+v", sym)
		}
	}
}

func TestWorkspaceSymbolsPreferOpenDocumentOverFilesystemContent(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Option Explicit\nPublic Sub OldName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	open := Document{
		URI:    "file:///C:/work/src/modules/Main.bas",
		Path:   path,
		Source: "Option Explicit\nPublic Sub NewName()\nEnd Sub\n",
	}

	symbols, err := analyzer.WorkspaceSymbols([]Document{open}, "")
	if err != nil {
		t.Fatal(err)
	}
	if hasSymbol(symbols, "OldName") {
		t.Fatalf("stale filesystem symbol should be hidden while document is open: %+v", symbols)
	}
	if !hasSymbol(symbols, "NewName") {
		t.Fatalf("open document symbol missing: %+v", symbols)
	}
}

func TestResolveExpressionTypeHandlesExcelCollectionsAndCreateObject(t *testing.T) {
	analyzer := newTestAnalyzer(t)

	if got, ok := analyzer.ResolveExpressionType(`Worksheets("Input").Range("A1")`); !ok || got != "Excel.Range" {
		t.Fatalf("Worksheets(...).Range(...) = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Workbooks.Open("book.xlsx")`); !ok || got != "Excel.Workbook" {
		t.Fatalf("Workbooks.Open(...) = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Range("A1").Font`); !ok || got != "Excel.Font" {
		t.Fatalf("Range(...).Font = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Cells(1, "A").End(xlUp)`); !ok || got != "Excel.Range" {
		t.Fatalf("Cells(...).End(...) = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Application.WorksheetFunction`); !ok || got != "Excel.WorksheetFunction" {
		t.Fatalf("Application.WorksheetFunction = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Worksheets("入力").ListObjects("売上一覧").ListColumns("金額").DataBodyRange`); !ok || got != "Excel.Range" {
		t.Fatalf("ListObjects(...).ListColumns(...).DataBodyRange = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Worksheets("集計").PivotTables("Pivot1").PivotFields("Name").DataRange`); !ok || got != "Excel.Range" {
		t.Fatalf("PivotTables(...).PivotFields(...).DataRange = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Worksheets("入力").Shapes("Button1").TextFrame.Characters.Text`); !ok || got != "String" {
		t.Fatalf("Shapes(...).TextFrame.Characters.Text = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Sheets("入力").Range("A1")`); !ok || got != "Excel.Range" {
		t.Fatalf("Sheets(...).Range(...) = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`ThisWorkbook.Sheets("入力").ListObjects("売上一覧").DataBodyRange`); !ok || got != "Excel.Range" {
		t.Fatalf("ThisWorkbook.Sheets(...).ListObjects(...).DataBodyRange = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Sheets("入力").Shapes("Button1").TextFrame.Characters.Text`); !ok || got != "String" {
		t.Fatalf("Sheets(...).Shapes(...).TextFrame.Characters.Text = %q, %v", got, ok)
	}

	doc := Document{
		Source: `Option Explicit
Sub Test()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "a", 1
End Sub
`,
	}
	if got, ok := analyzer.inferWordType(doc, "dict"); !ok || got != "Scripting.Dictionary" {
		t.Fatalf("dict type = %q, %v", got, ok)
	}
}

func TestListObjectChainCompletionAndHover(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim ws As Worksheet
    ws.ListObjects("売上一覧").ListColumns("金額").DataBodyRange
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	line := `    ws.ListObjects("売上一覧").ListColumns("金額").DataBodyRange`
	completionPrefix := line[:strings.Index(line, "DataBody")+len("DataBody")]
	items, err := analyzer.Completions(doc, Position{Line: 3, Character: utf16Len(completionPrefix)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "DataBodyRange") {
		t.Fatalf("ListColumn.DataBodyRange completion missing: %+v", items)
	}

	hover, err := analyzer.Hover(doc, Position{Line: 3, Character: utf16Len(line[:strings.Index(line, "DataBodyRange")+len("Data")])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.ListColumn.DataBodyRange As Excel.Range") {
		t.Fatalf("unexpected ListColumn.DataBodyRange hover: %+v", hover)
	}
}

func TestSheetsDefaultMemberAssumesWorksheetForWorksheetMembers(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Sheets("入力").Ra
    ThisWorkbook.Sheets("入力").ListObjects("売上一覧").DataBodyRange
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	rangeLine := `    Sheets("入力").Ra`
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(rangeLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Range") {
		t.Fatalf("Sheets default member should complete Worksheet.Range: %+v", items)
	}

	dataLine := `    ThisWorkbook.Sheets("入力").ListObjects("売上一覧").DataBodyRange`
	hover, err := analyzer.Hover(doc, Position{Line: 3, Character: utf16Len(dataLine[:strings.Index(dataLine, "ListObjects")+len("List")])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.Worksheet.ListObjects As Excel.ListObjects") {
		t.Fatalf("unexpected Sheets worksheet-assumption hover: %+v", hover)
	}
}

func TestPivotAndShapeChainCompletionAndHover(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim ws As Worksheet
    ws.PivotTables("Pivot1").PivotFields("Name").Data
    ws.Shapes("Button1").TextFrame.Characters.Te
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	pivotLine := `    ws.PivotTables("Pivot1").PivotFields("Name").Data`
	items, err := analyzer.Completions(doc, Position{Line: 3, Character: utf16Len(pivotLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "DataRange") {
		t.Fatalf("PivotField.DataRange completion missing: %+v", items)
	}

	shapeLine := `    ws.Shapes("Button1").TextFrame.Characters.Te`
	items, err = analyzer.Completions(doc, Position{Line: 4, Character: utf16Len(shapeLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Text") {
		t.Fatalf("Characters.Text completion missing: %+v", items)
	}

	hover, err := analyzer.Hover(doc, Position{Line: 4, Character: utf16Len(shapeLine[:strings.Index(shapeLine, "TextFrame")+len("Text")])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.Shape.TextFrame As Excel.TextFrame") {
		t.Fatalf("unexpected Shape.TextFrame hover: %+v", hover)
	}
}

func TestSetAssignmentInferencePropagatesRightHandExpressionTypes(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim ws As Worksheet
    Dim target As Object
    Set target = ws.Range("A1")
    target.Va

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(path)
    ts.Read
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	if got, ok := analyzer.inferWordType(doc, "target"); !ok || got != "Excel.Range" {
		t.Fatalf("target type = %q, %v, want Excel.Range", got, ok)
	}
	targetLine := `    target.Va`
	items, err := analyzer.Completions(doc, Position{Line: 5, Character: utf16Len(targetLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Value") {
		t.Fatalf("target should complete Range.Value after Set assignment: %+v", items)
	}
	hoverLine := `    Set target = ws.Range("A1")`
	hover, err := analyzer.Hover(doc, Position{Line: 4, Character: utf16Len(hoverLine[:strings.Index(hoverLine, "target")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "target As Excel.Range") || !strings.Contains(hover.Contents, "Source: inferred from Set assignment") {
		t.Fatalf("unexpected target hover: %+v", hover)
	}

	if got, ok := analyzer.inferWordType(doc, "ts"); !ok || got != "Scripting.TextStream" {
		t.Fatalf("ts type = %q, %v, want Scripting.TextStream", got, ok)
	}
	textLine := `    ts.Read`
	items, err = analyzer.Completions(doc, Position{Line: 11, Character: utf16Len(textLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "ReadLine") {
		t.Fatalf("ts should complete TextStream.ReadLine after method-return inference: %+v", items)
	}
}

func TestPracticalWorkbookFileSystemAndWorksheetChains(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Public Sub hoge()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.OpenTextFile("hoge.txt")

    ts.ReadLine
End Sub

Public Sub fuga()
    Dim wb As Workbook
    Set wb = Application.Workbooks.Open("Sample.xlsx")

    Dim ws As Worksheet
    Set ws = wb.Worksheets("Input")

    wb.Worksheets(1).Range("A1")
    wb.Sheets("Sheet1")
    ws.ListObjects("Table1").DataBodyRange
    ws.PivotTables("Pivot1").PivotFields("Name")
    ws.Shapes("Button1").TextFrame
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	assertType := func(expr, want string) {
		t.Helper()
		if got, ok := analyzer.resolveDocumentExpressionTypeAt(doc, expr, len(source)); !ok || got != want {
			t.Fatalf("%s = %q, %v, want %s", expr, got, ok, want)
		}
	}

	assertType(`fso.OpenTextFile("hoge.txt")`, "Scripting.TextStream")
	assertType(`ts.ReadLine`, "String")
	assertType(`Application.Workbooks.Open("Sample.xlsx")`, "Excel.Workbook")
	assertType(`wb.Worksheets(1).Range("A1")`, "Excel.Range")
	assertType(`wb.Sheets("Sheet1")`, "Excel.Worksheet")
	assertType(`ws.ListObjects("Table1").DataBodyRange`, "Excel.Range")
	assertType(`ws.PivotTables("Pivot1").PivotFields("Name")`, "Excel.PivotField")
	assertType(`ws.Shapes("Button1").TextFrame`, "Excel.TextFrame")

	line := `    ts.Read`
	items, err := analyzer.Completions(doc, Position{Line: 8, Character: utf16Len(line)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "ReadLine") {
		t.Fatalf("ts.Read should complete TextStream.ReadLine: %+v", items)
	}
}

func TestHoverUsesUTF16PositionsAfterJapaneseText(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := "Option Explicit\nSub Test()\n    Debug.Print \"日本語\" & Range(\"A1\").Value\nEnd Sub\n"
	line := `    Debug.Print "日本語" & Range("A1").Value`
	character := utf16Len(line[:strings.Index(line, "Range")+len("Ra")])
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: character}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.Range") {
		t.Fatalf("unexpected hover: %+v", hover)
	}
}

func TestHoverAndCompletionsResolveBuiltInCollection(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Private Sub Earlier(ByVal result As Boolean)
End Sub

Sub Test()
    Dim result As Collection
    Set result = New Collection
    result.Co
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	hoverLine := `    Set result = New Collection`
	hover, err := analyzer.Hover(doc, Position{Line: 6, Character: utf16Len(hoverLine[:strings.Index(hoverLine, "result")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "VBA.Collection") {
		t.Fatalf("unexpected Collection hover: %+v", hover)
	}
	if strings.Contains(hover.Contents, "Boolean") {
		t.Fatalf("Collection hover should not use earlier Boolean declaration: %+v", hover)
	}

	completionLine := `    result.Co`
	items, err := analyzer.Completions(doc, Position{Line: 7, Character: utf16Len(completionLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Count") {
		t.Fatalf("Collection.Count completion missing: %+v", items)
	}
}

func TestHoverFormatsMemberSignaturesAndSources(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "a", 1
    Dim ws As Worksheet
    ws.Range("A1").Value = 1
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	dictLine := `    Set dict = CreateObject("Scripting.Dictionary")`
	hover, err := analyzer.Hover(doc, Position{Line: 3, Character: utf16Len(dictLine[:strings.Index(dictLine, "dict")+2])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "dict As Scripting.Dictionary") || !strings.Contains(hover.Contents, "Source: inferred from CreateObject") {
		t.Fatalf("unexpected dict hover: %+v", hover)
	}

	addLine := `    dict.Add "a", 1`
	hover, err = analyzer.Hover(doc, Position{Line: 4, Character: utf16Len(addLine[:strings.Index(addLine, "Add")+2])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Scripting.Dictionary.Add(Key As Variant, Item As Variant) As void") || !strings.Contains(hover.Contents, "Source: built-in Scripting object model DB") {
		t.Fatalf("unexpected Dictionary.Add hover: %+v", hover)
	}

	rangeLine := `    ws.Range("A1").Value = 1`
	hover, err = analyzer.Hover(doc, Position{Line: 6, Character: utf16Len(rangeLine[:strings.Index(rangeLine, "Range")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.Worksheet.Range(Cell1 As Variant, Optional Cell2 As Variant) As Excel.Range") {
		t.Fatalf("unexpected Worksheet.Range hover: %+v", hover)
	}
}

func TestCompletionsReturnMemberAndGlobalCandidates(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := "Option Explicit\nSub Test()\n    Worksheets(\"Input\").Ra\nEnd Sub\n"
	memberLine := `    Worksheets("Input").Ra`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(memberLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Range") {
		t.Fatalf("Range completion missing: %+v", items)
	}
	rangeCompletion, _ := findCompletion(items, "Range")
	if rangeCompletion.Detail != "Excel.Worksheet.Range(Cell1 As Variant, Optional Cell2 As Variant) As Excel.Range" {
		t.Fatalf("Range completion detail = %q", rangeCompletion.Detail)
	}

	fontSource := "Option Explicit\nSub Test()\n    Range(\"A1\").Font.Co\nEnd Sub\n"
	fontLine := `    Range("A1").Font.Co`
	fontDoc := Document{Path: filepath.Join(t.TempDir(), "Formatting.bas"), Source: fontSource}
	items, err = analyzer.Completions(fontDoc, Position{Line: 2, Character: utf16Len(fontLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Color") {
		t.Fatalf("Font.Color completion missing: %+v", items)
	}

	globalDoc := Document{Path: filepath.Join(t.TempDir(), "Globals.bas"), Source: "Option Explicit\nSub Test()\n    xlU\nEnd Sub\n"}
	items, err = analyzer.Completions(globalDoc, Position{Line: 2, Character: utf16Len("    xlU")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "xlUp") {
		t.Fatalf("xlUp completion missing: %+v", items)
	}
	layoutDoc := Document{Path: filepath.Join(t.TempDir(), "Layout.bas"), Source: "Option Explicit\nSub Test()\n    xlLand\nEnd Sub\n"}
	items, err = analyzer.Completions(layoutDoc, Position{Line: 2, Character: utf16Len("    xlLand")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "xlLandscape") {
		t.Fatalf("xlLandscape completion missing: %+v", items)
	}
}

func TestWithBlockHoverAndCompletionsUseActiveReceiverType(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub Test()
    Dim ws As Worksheet
    With ws.Range("A1")
        .Va
        With .Font
            .Bo
        End With
    End With
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	withStartLine := `    With ws.Ra`
	items, err := analyzer.Completions(doc, Position{Line: 3, Character: utf16Len(withStartLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Range") {
		t.Fatalf("With receiver expression should complete Worksheet.Range: %+v", items)
	}

	valueLine := `        .Va`
	items, err = analyzer.Completions(doc, Position{Line: 4, Character: utf16Len(valueLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Value") {
		t.Fatalf("With Range should complete Range.Value: %+v", items)
	}
	if hasCompletion(items, "Bold") {
		t.Fatalf("outer With Range should not complete Font.Bold directly: %+v", items)
	}

	hoverLine := `        With .Font`
	hover, err := analyzer.Hover(doc, Position{Line: 5, Character: utf16Len(hoverLine[:strings.Index(hoverLine, "Font")+2])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.Range.Font As Excel.Font") {
		t.Fatalf("unexpected With .Font hover: %+v", hover)
	}

	boldLine := `            .Bo`
	items, err = analyzer.Completions(doc, Position{Line: 6, Character: utf16Len(boldLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Bold") {
		t.Fatalf("nested With Font should complete Font.Bold: %+v", items)
	}
}

func TestCompletionsReturnTypeCandidatesInDeclarationContexts(t *testing.T) {
	root := t.TempDir()
	classDir := filepath.Join(root, "src", "classes")
	if err := os.MkdirAll(classDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classDir, "MyService.cls"), []byte(`VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "MyService"
Option Explicit
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:   filepath.Join(root, "src", "modules", "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    Dim ws As Wo\nEnd Sub\n",
	}
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Dim ws As Wo")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Workbook") {
		t.Fatalf("Workbook type alias completion missing: %+v", items)
	}
	if hasCompletion(items, "xlWorksheet") || hasCompletion(items, "ThisWorkbook") {
		t.Fatalf("type context should not include constants or globals: %+v", items)
	}
	workbook, ok := findCompletion(items, "Workbook")
	if !ok || workbook.ReplaceRange == nil || workbook.ReplaceRange.Start.Character != utf16Len("    Dim ws As ") {
		t.Fatalf("Workbook should replace only typed type prefix: %+v", workbook)
	}

	doc.Source = "Option Explicit\nSub Test()\n    Dim value As Str\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Dim value As Str")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "String") {
		t.Fatalf("String built-in type completion missing: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Test()\n    Set dict = New Di\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Set dict = New Di")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Dictionary") {
		t.Fatalf("Dictionary alias completion missing after New: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Test()\n    Dim service As My\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Dim service As My")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "MyService") {
		t.Fatalf("class module type completion missing: %+v", items)
	}
}

func TestCompletionsReturnProgIDsInsideCreateObjectString(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	line := `    Set dict = CreateObject("`
	doc := Document{
		Path:   filepath.Join(t.TempDir(), "Main.bas"),
		Source: "Option Explicit\nSub Test()\n" + line + "\nEnd Sub\n",
	}

	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(line)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	item, ok := findCompletion(items, "Scripting.Dictionary")
	if !ok {
		t.Fatalf("Scripting.Dictionary ProgID completion missing: %+v", items)
	}
	if item.Detail != "Scripting.Dictionary" {
		t.Fatalf("ProgID detail = %q, want resolved type", item.Detail)
	}
	for _, want := range []string{"ADODB.Connection", "ADODB.Recordset", "Excel.Application"} {
		item, ok := findCompletion(items, want)
		if !ok {
			t.Fatalf("%s ProgID completion missing: %+v", want, items)
		}
		if item.Detail != want {
			t.Fatalf("%s ProgID detail = %q, want resolved type", want, item.Detail)
		}
	}
	if item.ReplaceRange == nil || item.ReplaceRange.Start.Character != utf16Len(`    Set dict = CreateObject("`) || item.ReplaceRange.End.Character != utf16Len(line) {
		t.Fatalf("ProgID replace range = %+v, want string literal content range", item.ReplaceRange)
	}

	otherString := `    Debug.Print "xlU`
	doc.Source = "Option Explicit\nSub Test()\n" + otherString + "\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(otherString)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 0 {
		t.Fatalf("non-CreateObject strings should not return completions: %+v", items)
	}
}

func TestCompletionsReturnSetObjectInitializersAndVBAIsFunctions(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Sub Test()
    Dim target As Object
    Set target = N
    Set target = No
    Set target = Get
    If IsObject(
    If Is
End Sub
`,
	}

	newLine := `    Set target = N`
	items, err := analyzer.Completions(doc, Position{Line: 3, Character: utf16Len(newLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	newItem, ok := findCompletion(items, "New")
	if !ok || !newItem.Snippet || !strings.Contains(newItem.InsertText, "New ${1:Collection}") {
		t.Fatalf("New snippet completion missing after Set RHS: %+v", items)
	}
	if !hasCompletion(items, "Nothing") {
		t.Fatalf("Nothing completion missing after Set RHS: %+v", items)
	}

	nothingLine := `    Set target = No`
	items, err = analyzer.Completions(doc, Position{Line: 4, Character: utf16Len(nothingLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Nothing") {
		t.Fatalf("Nothing completion missing after Set RHS prefix No: %+v", items)
	}

	getObjectLine := `    Set target = Get`
	items, err = analyzer.Completions(doc, Position{Line: 5, Character: utf16Len(getObjectLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "GetObject") {
		t.Fatalf("GetObject completion missing after Set RHS: %+v", items)
	}

	isLine := `    If Is`
	items, err = analyzer.Completions(doc, Position{Line: 7, Character: utf16Len(isLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"IsObject", "IsNull", "IsNumeric", "IsDate"} {
		if !hasCompletion(items, want) {
			t.Fatalf("%s completion missing in condition context: %+v", want, items)
		}
	}

	helpLine := `    If IsObject(`
	help, err := analyzer.SignatureHelp(doc, Position{Line: 6, Character: utf16Len(helpLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 || !strings.Contains(help.Signatures[0].Label, "VBA.Global.IsObject(Identifier As Variant) As Boolean") {
		t.Fatalf("IsObject signature help = %+v", help)
	}
}

func TestCompletionsReturnModuleProcedureCandidates(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	utilsPath := filepath.Join(moduleDir, "Utils.bas")
	if err := os.WriteFile(utilsPath, []byte(`Attribute VB_Name = "Utils"
Option Explicit
Public Function BuildName() As String
End Function
Private Function HiddenName() As String
End Function
Sub RunDefault()
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:   filepath.Join(moduleDir, "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    Utils.Bu\nEnd Sub\n",
	}
	line := `    Utils.Bu`
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(line)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "BuildName") {
		t.Fatalf("Utils.BuildName completion missing: %+v", items)
	}
	if hasCompletion(items, "HiddenName") {
		t.Fatalf("private module member should not be completed: %+v", items)
	}

	openUtils := Document{
		URI:    "file:///C:/work/src/modules/Utils.bas",
		Path:   utilsPath,
		Source: "Attribute VB_Name = \"Utils\"\nOption Explicit\nPublic Function UnsavedName() As String\nEnd Function\n",
	}
	doc.Source = "Option Explicit\nSub Test()\n    Utils.Un\nEnd Sub\n"
	line = `    Utils.Un`
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(line)}, []Document{openUtils})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "UnsavedName") {
		t.Fatalf("open document module member completion missing: %+v", items)
	}
	if hasCompletion(items, "BuildName") {
		t.Fatalf("stale filesystem module member should be hidden while open document exists: %+v", items)
	}
}

func TestCompletionsRespectCallAndSetRHSContexts(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	classDir := filepath.Join(root, "src", "classes")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(classDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Helpers.bas"), []byte(`Attribute VB_Name = "Helpers"
Option Explicit
Public Sub RunReport()
End Sub
Private Sub HiddenReport()
End Sub
Public Function BuildWorkbook() As Workbook
End Function
Public Function BuildName() As String
End Function
Public Const PUBLIC_VALUE As Long = 1
`), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classDir, "ReportService.cls"), []byte(`VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ReportService"
Option Explicit
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:   filepath.Join(moduleDir, "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    Call Ru\n    Set ws = Th\n    Set service = Re\nEnd Sub\n",
	}

	callLine := `    Call Ru`
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(callLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "RunReport") {
		t.Fatalf("Call context should include public Sub: %+v", items)
	}
	if hasCompletion(items, "HiddenReport") || hasCompletion(items, "PUBLIC_VALUE") || hasCompletion(items, "ThisWorkbook") {
		t.Fatalf("Call context should exclude private/non-callable/global candidates: %+v", items)
	}

	setWorkbookLine := `    Set ws = Th`
	items, err = analyzer.Completions(doc, Position{Line: 3, Character: utf16Len(setWorkbookLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "ThisWorkbook") {
		t.Fatalf("Set RHS context should include object globals: %+v", items)
	}
	if hasCompletion(items, "xlUp") || hasCompletion(items, "String") || hasCompletion(items, "PUBLIC_VALUE") || hasCompletion(items, "BuildName") {
		t.Fatalf("Set RHS context should exclude constants, primitive types, and scalar functions: %+v", items)
	}

	setServiceLine := `    Set service = Re`
	items, err = analyzer.Completions(doc, Position{Line: 4, Character: utf16Len(setServiceLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "ReportService") {
		t.Fatalf("Set RHS context should include class modules: %+v", items)
	}
}

func TestCompletionsRespectValueRHSContext(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Helpers.bas"), []byte(`Attribute VB_Name = "Helpers"
Option Explicit
Public Sub RunReport()
End Sub
Public Function BuildName() As String
End Function
Public Function BuildWorkbook() As Workbook
End Function
Public Const PUBLIC_VALUE As Long = 1
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path: filepath.Join(moduleDir, "Main.bas"),
		Source: `Option Explicit
Private Const MODULE_VALUE As Long = 2
Sub Test()
    Dim localValue As Long
    Dim value As Variant
    value = xl
    value = Bu
    value = PU
    value = local
    Range("A1").Value = xl
End Sub
`,
	}

	constLine := `    value = xl`
	items, err := analyzer.Completions(doc, Position{Line: 5, Character: utf16Len(constLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "xlUp") {
		t.Fatalf("value RHS context should include constants: %+v", items)
	}
	if hasCompletion(items, "Dim") || hasCompletion(items, "Workbook") {
		t.Fatalf("value RHS context should exclude snippets and type-only candidates: %+v", items)
	}

	funcLine := `    value = Bu`
	items, err = analyzer.Completions(doc, Position{Line: 6, Character: utf16Len(funcLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "BuildName") || !hasCompletion(items, "BuildWorkbook") {
		t.Fatalf("value RHS context should include value-returning functions: %+v", items)
	}
	if hasCompletion(items, "RunReport") {
		t.Fatalf("value RHS context should exclude Sub procedures: %+v", items)
	}

	publicConstLine := `    value = PU`
	items, err = analyzer.Completions(doc, Position{Line: 7, Character: utf16Len(publicConstLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "PUBLIC_VALUE") {
		t.Fatalf("value RHS context should include visible constants: %+v", items)
	}

	localLine := `    value = local`
	items, err = analyzer.Completions(doc, Position{Line: 8, Character: utf16Len(localLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "localValue") {
		t.Fatalf("value RHS context should include current procedure locals: %+v", items)
	}

	propertyLine := `    Range("A1").Value = xl`
	items, err = analyzer.Completions(doc, Position{Line: 9, Character: utf16Len(propertyLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "xlUp") {
		t.Fatalf("value RHS context should support property assignment targets: %+v", items)
	}
}

func TestCompletionsRespectConditionContexts(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Helpers.bas"), []byte(`Attribute VB_Name = "Helpers"
Option Explicit
Public Function BuildReady() As Boolean
End Function
Public Const PUBLIC_FLAG As Boolean = True
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path: filepath.Join(moduleDir, "Main.bas"),
		Source: `Option Explicit
Sub Test()
    Dim localFlag As Boolean
    If PU
    If local
    Do While Bu
    If localFlag = PU
End Sub
`,
	}

	publicFlagLine := `    If PU`
	items, err := analyzer.Completions(doc, Position{Line: 3, Character: utf16Len(publicFlagLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "PUBLIC_FLAG") {
		t.Fatalf("condition context should include constants: %+v", items)
	}
	if hasCompletion(items, "Dim") || hasCompletion(items, "String") {
		t.Fatalf("condition context should exclude snippets and type-only candidates: %+v", items)
	}

	localLine := `    If local`
	items, err = analyzer.Completions(doc, Position{Line: 4, Character: utf16Len(localLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "localFlag") {
		t.Fatalf("condition context should include current procedure locals: %+v", items)
	}

	functionLine := `    Do While Bu`
	items, err = analyzer.Completions(doc, Position{Line: 5, Character: utf16Len(functionLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "BuildReady") {
		t.Fatalf("condition context should include value-returning functions: %+v", items)
	}

	comparisonLine := `    If localFlag = PU`
	items, err = analyzer.Completions(doc, Position{Line: 6, Character: utf16Len(comparisonLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "PUBLIC_FLAG") {
		t.Fatalf("condition context should complete after comparison operators: %+v", items)
	}
}

func TestCompletionsReturnControlFlowSnippetsAndForEachValues(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Helpers.bas"), []byte(`Attribute VB_Name = "Helpers"
Option Explicit
Public Function BuildItems() As Collection
End Function
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path: filepath.Join(moduleDir, "Main.bas"),
		Source: `Option Explicit
Sub Test()
    Dim items As Collection
    I
    For Each item In Bu
    For Each item In items
End Sub
`,
	}

	snippetLine := `    I`
	items, err := analyzer.Completions(doc, Position{Line: 3, Character: utf16Len(snippetLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "If Then") || !hasCompletion(items, "If Else") {
		t.Fatalf("procedure statement context should include If snippets: %+v", items)
	}

	functionLine := `    For Each item In Bu`
	items, err = analyzer.Completions(doc, Position{Line: 4, Character: utf16Len(functionLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "BuildItems") {
		t.Fatalf("For Each In context should include value-returning functions: %+v", items)
	}
	if hasCompletion(items, "For Each") || hasCompletion(items, "Collection") {
		t.Fatalf("For Each In context should exclude snippets and type-only candidates: %+v", items)
	}

	localLine := `    For Each item In items`
	items, err = analyzer.Completions(doc, Position{Line: 5, Character: utf16Len(localLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "items") {
		t.Fatalf("For Each In context should include local collection variables: %+v", items)
	}
}

func TestCompletionsHidePrivateSymbolsFromOtherModules(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Utils.bas"), []byte(`Attribute VB_Name = "Utils"
Option Explicit
Private Const SECRET_VALUE As Long = 1
Public Const PUBLIC_VALUE As Long = 2
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:   filepath.Join(moduleDir, "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    SE\nEnd Sub\n",
	}
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    SE")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hasCompletion(items, "SECRET_VALUE") {
		t.Fatalf("private const from another module should not be completed: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Test()\n    PU\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    PU")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "PUBLIC_VALUE") {
		t.Fatalf("public const from another module should be completed: %+v", items)
	}

	doc.Source = "Option Explicit\nPrivate Const SECRET_VALUE As Long = 1\nSub Test()\n    SE\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 3, Character: utf16Len("    SE")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "SECRET_VALUE") {
		t.Fatalf("private const from the same document should remain available: %+v", items)
	}
}

func TestCompletionsHideLocalsFromOtherProcedures(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Private Const MODULE_SECRET As Long = 1

Sub First()
    Dim firstLocal As Long
    fi
End Sub

Sub Second()
    Dim secondLocal As Long
    se
End Sub
`,
	}

	items, err := analyzer.Completions(doc, Position{Line: 5, Character: utf16Len("    fi")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "firstLocal") {
		t.Fatalf("current procedure local should be completed: %+v", items)
	}
	if hasCompletion(items, "secondLocal") {
		t.Fatalf("other procedure local should not be completed: %+v", items)
	}

	items, err = analyzer.Completions(doc, Position{Line: 10, Character: utf16Len("    se")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "secondLocal") {
		t.Fatalf("current procedure local should be completed: %+v", items)
	}
	if hasCompletion(items, "firstLocal") {
		t.Fatalf("other procedure local should not be completed: %+v", items)
	}

	doc.Source = strings.Replace(doc.Source, "    fi", "    MO", 1)
	items, err = analyzer.Completions(doc, Position{Line: 5, Character: utf16Len("    MO")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "MODULE_SECRET") {
		t.Fatalf("same module private const should remain available: %+v", items)
	}
}

func TestCompletionsReturnModuleDeclarationSnippets(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit

Pu
Public Sub Existing()
End Sub
`,
	}

	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("Pu")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	publicSub, ok := findCompletion(items, "Public Sub")
	if !ok {
		t.Fatalf("Public Sub completion missing: %+v", items)
	}
	if !publicSub.Snippet || !strings.Contains(publicSub.InsertText, "End Sub") {
		t.Fatalf("Public Sub should be a procedure snippet: %+v", publicSub)
	}
	if !hasCompletion(items, "Public Function") {
		t.Fatalf("Public Function completion missing: %+v", items)
	}
	doc.Source = "Option Explicit\n\nPublic S\nPublic Sub Existing()\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("Public S")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	publicSub, ok = findCompletion(items, "Public Sub")
	if !ok {
		t.Fatalf("Public Sub completion missing for multi-word prefix: %+v", items)
	}
	if publicSub.ReplaceRange == nil || publicSub.ReplaceRange.Start.Character != 0 || publicSub.ReplaceRange.End.Character != utf16Len("Public S") {
		t.Fatalf("Public Sub should replace the typed declaration prefix: %+v", publicSub.ReplaceRange)
	}

	doc.Source = ""
	items, err = analyzer.Completions(doc, Position{Line: 0, Character: 0}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Option Explicit") {
		t.Fatalf("empty module should offer Option Explicit: %+v", items)
	}

	doc.Source = "Option Explicit"
	items, err = analyzer.Completions(doc, Position{Line: 0, Character: utf16Len("Option Explicit")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hasCompletion(items, "Option Explicit") {
		t.Fatalf("completed Option Explicit should not be offered again: %+v", items)
	}

	doc.Source = "Option Explicit\n"
	items, err = analyzer.Completions(doc, Position{Line: 1, Character: 0}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 0 {
		t.Fatalf("blank line after Option Explicit should not offer module snippets: %+v", items)
	}
	if hasCompletion(items, "Option Explicit") {
		t.Fatalf("line after Option Explicit should not offer Option Explicit again: %+v", items)
	}

	doc.Source = "Option Explicit\n\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: 0}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(items) != 0 {
		t.Fatalf("blank module line after existing content should not offer module snippets: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Existing()\n    Pu\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Pu")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hasCompletion(items, "Public Sub") {
		t.Fatalf("module-level declaration snippets should not appear inside procedures: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Existing()\n    Di\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Di")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Dim") {
		t.Fatalf("procedure Dim snippet should appear inside procedures: %+v", items)
	}
	dimSnippet, ok := findCompletion(items, "Dim")
	if !ok || !dimSnippet.Snippet || !strings.Contains(dimSnippet.InsertText, "As ${2:Variant}") {
		t.Fatalf("Dim should be a local declaration snippet: %+v", dimSnippet)
	}

	doc.Source = "Option Explicit\nSub Existing()\n    Se\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Se")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	setSnippet, ok := findCompletion(items, "Set")
	if !ok || !setSnippet.Snippet || !strings.Contains(setSnippet.InsertText, "Set ${1:target} = ${2:expression}") {
		t.Fatalf("Set should be an object assignment snippet: %+v", setSnippet)
	}
}

func TestHoverUsesConstantMetadata(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	line := "    Debug.Print xlLandscape"
	doc := Document{Path: filepath.Join(t.TempDir(), "Constants.bas"), Source: "Option Explicit\nSub Test()\n" + line + "\nEnd Sub\n"}

	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: utf16Len(line[:strings.Index(line, "xlLandscape")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "XlPageOrientation") {
		t.Fatalf("unexpected constant hover: %+v", hover)
	}
}

func TestUserFormControlsResolveForHoverCompletionAndDefinition(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
   Begin MSForms.TextBox txtName
   End
End
Attribute VB_Name = "CustomerForm"
Option Explicit
Private Sub UserForm_Initialize()
    Me.txtName.Text = "Alice"
    Me.Controls("txtName").Text = "Bob"
End Sub
`
	doc := Document{
		URI:        "file:///C:/work/src/forms/CustomerForm.frm",
		Path:       filepath.Join(t.TempDir(), "CustomerForm.frm"),
		ModuleKind: "form",
		Source:     source,
	}

	hoverLine := `    Me.txtName.Text = "Alice"`
	hover, err := analyzer.Hover(doc, Position{Line: 8, Character: utf16Len(hoverLine[:strings.Index(hoverLine, "txtName")+3])}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "MSForms.TextBox") || !strings.Contains(hover.Contents, "Source: UserForm control") {
		t.Fatalf("unexpected hover: %+v", hover)
	}

	controlLine := `    Me.Controls("txtName").Te`
	items, err := analyzer.Completions(doc, Position{Line: 9, Character: utf16Len(controlLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Text") {
		t.Fatalf("Text completion missing: %+v", items)
	}

	defs, err := analyzer.Definition(doc, Position{Line: 8, Character: utf16Len(`    Me.txtN`)}, []Document{doc}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(defs) == 0 || defs[0].Range.Start.Line != 2 {
		t.Fatalf("control definition = %+v, want declaration on .frm Begin line", defs)
	}
}

func TestUserFormControlsResolveFromSidecarCodeBehind(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	frm := `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
   Begin MSForms.TextBox txtName
   End
End
Attribute VB_Name = "CustomerForm"
`
	if err := os.WriteFile(filepath.Join(formsDir, "CustomerForm.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}
	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:       filepath.Join(codeDir, "CustomerForm.bas"),
		ModuleKind: "form",
		Source:     "Option Explicit\nPrivate Sub UserForm_Initialize()\n    Me.txtName.Text = \"Alice\"\nEnd Sub\n",
	}
	line := `    Me.txtName.Text = "Alice"`
	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: utf16Len(line[:strings.Index(line, "txtName")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "MSForms.TextBox") {
		t.Fatalf("unexpected sidecar hover: %+v", hover)
	}
}

func TestReferencesFindOpenDocumentOccurrencesAndCanSkipDeclaration(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Public Sub RunReport()
    RunReport
End Sub
`
	doc := Document{
		URI:        "file:///C:/work/src/modules/Main.bas",
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	pos := Position{Line: 1, Character: utf16Len("Public Sub Run")}

	withDecl, err := analyzer.References(doc, pos, []Document{doc}, true, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withDecl) != 2 {
		t.Fatalf("references with declaration = %d, want 2: %+v", len(withDecl), withDecl)
	}
	withoutDecl, err := analyzer.References(doc, pos, []Document{doc}, false, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withoutDecl) != 1 || withoutDecl[0].Range.Start.Line != 2 {
		t.Fatalf("references without declaration = %+v, want only call line", withoutDecl)
	}
}

func TestDefinitionAndReferencesPreferCurrentProcedureLocal(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub First()
    Dim value As Long
    value = 1
End Sub

Sub Second()
    Dim value As Long
    value = 2
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	pos := Position{Line: 3, Character: utf16Len("    val")}

	defs, err := analyzer.Definition(doc, pos, []Document{doc}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(defs) != 1 || defs[0].Range.Start.Line != 2 {
		t.Fatalf("definition = %+v, want First.value declaration only", defs)
	}

	withDecl, err := analyzer.References(doc, pos, []Document{doc}, true, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withDecl) != 2 || withDecl[0].Range.Start.Line != 2 || withDecl[1].Range.Start.Line != 3 {
		t.Fatalf("local references with declaration = %+v, want First declaration and assignment only", withDecl)
	}

	withoutDecl, err := analyzer.References(doc, pos, []Document{doc}, false, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withoutDecl) != 1 || withoutDecl[0].Range.Start.Line != 3 {
		t.Fatalf("local references without declaration = %+v, want First assignment only", withoutDecl)
	}
}

func TestDefinitionAndReferencesResolveCurrentProcedureParameter(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub First(ByVal value As Long)
    value = value + 1
End Sub

Sub Second(ByVal value As Long)
    value = value + 2
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	pos := Position{Line: 2, Character: utf16Len("    val")}

	defs, err := analyzer.Definition(doc, pos, []Document{doc}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(defs) != 1 || defs[0].Range.Start.Line != 1 {
		t.Fatalf("parameter definition = %+v, want First.value parameter only", defs)
	}

	withDecl, err := analyzer.References(doc, pos, []Document{doc}, true, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withDecl) != 3 || withDecl[0].Range.Start.Line != 1 || withDecl[1].Range.Start.Line != 2 || withDecl[2].Range.Start.Line != 2 {
		t.Fatalf("parameter references with declaration = %+v, want First parameter and body references only", withDecl)
	}

	withoutDecl, err := analyzer.References(doc, pos, []Document{doc}, false, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withoutDecl) != 2 || withoutDecl[0].Range.Start.Line != 2 || withoutDecl[1].Range.Start.Line != 2 {
		t.Fatalf("parameter references without declaration = %+v, want First body references only", withoutDecl)
	}
}

func TestHoverUsesCurrentProcedureParameterSymbol(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub First(value)
    Debug.Print value
End Sub

Sub Second(value As String)
    Debug.Print value
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	line := `    Debug.Print value`
	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: utf16Len(line[:strings.Index(line, "value")+3])}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "value") {
		t.Fatalf("parameter hover missing: %+v", hover)
	}
	if strings.Contains(hover.Contents, "String") {
		t.Fatalf("parameter hover should not use another procedure parameter: %+v", hover)
	}
}

func TestSemanticTokensCoverDeclarationsBuiltinsMembersAndUTF16(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
' 日本語コメント
Public Type Customer
    Name As String
End Type
Public Enum Status
    StatusReady = 1
End Enum
Public Sub Run(ByVal value As String)
    Dim ws As Worksheet
    Set ws = Worksheets("Input")
    Debug.Print value, xlLandscape
    ws.Range("A1").Value = 1
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	tokens, err := analyzer.SemanticTokens(doc, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}

	if !hasSemanticToken(tokens, SemanticTokenFunction, "Run", source) {
		t.Fatalf("function token for Run missing: %+v", tokens)
	}
	if !hasSemanticToken(tokens, SemanticTokenParameter, "value", source) {
		t.Fatalf("parameter token for value missing: %+v", tokens)
	}
	if !hasSemanticToken(tokens, SemanticTokenClass, "Worksheet", source) {
		t.Fatalf("default-library class token for Worksheet missing: %+v", tokens)
	}
	if !hasSemanticToken(tokens, SemanticTokenFunction, "Debug", source) && !hasSemanticToken(tokens, SemanticTokenMethod, "Print", source) {
		t.Fatalf("VBA global function/member token missing: %+v", tokens)
	}
	if !hasSemanticToken(tokens, SemanticTokenEnumMember, "xlLandscape", source) {
		t.Fatalf("enum member token for xlLandscape missing: %+v", tokens)
	}
	if !hasSemanticToken(tokens, SemanticTokenProperty, "Value", source) {
		t.Fatalf("member property token for Value missing: %+v", tokens)
	}
	commentLine := lineIndex(source, "' 日本語コメント")
	if !hasSemanticTokenAtLine(tokens, SemanticTokenComment, commentLine) {
		t.Fatalf("comment token after Japanese text missing: %+v", tokens)
	}
}

func TestSemanticTokensCoverUserFormControlsAndMalformedSource(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `VERSION 5.00
Begin VB.Form UserForm1
   Caption = "UserForm1"
   Begin MSForms.TextBox txtName
      Left = 12
   End
End
Attribute VB_Name = "UserForm1"
Private Sub UserForm_Initialize()
    Me.txtName.Text = "ready"
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "UserForm1.frm"),
		ModuleKind: "form",
		Source:     source,
	}
	tokens, err := analyzer.SemanticTokens(doc, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasSemanticToken(tokens, SemanticTokenVariable, "txtName", source) {
		t.Fatalf("UserForm control token missing: %+v", tokens)
	}
	if !hasSemanticToken(tokens, SemanticTokenEvent, "Initialize", source) && !hasSemanticToken(tokens, SemanticTokenFunction, "UserForm_Initialize", source) {
		t.Fatalf("UserForm event/function token missing: %+v", tokens)
	}

	malformed := doc
	malformed.Source = "Public Sub Broken(\n    Debug.Print \"unterminated\n"
	if _, err := analyzer.SemanticTokens(malformed, []Document{malformed}); err != nil {
		t.Fatalf("semantic tokens should be best-effort on malformed source: %v", err)
	}
}

func TestSemanticTokensTolerateNilDatabaseForMemberTokens(t *testing.T) {
	analyzer := Analyzer{}
	source := `Option Explicit
Public Sub Run()
    Range("A1").Font.Color = 1
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}

	if _, err := analyzer.SemanticTokens(doc, []Document{doc}); err != nil {
		t.Fatalf("semantic tokens should tolerate nil database: %v", err)
	}
}

func newTestAnalyzer(t *testing.T) Analyzer {
	t.Helper()
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}
	return Analyzer{RootDir: t.TempDir(), Config: config.Default(), DB: db}
}

func hasCompletion(items []Completion, label string) bool {
	_, ok := findCompletion(items, label)
	return ok
}

func findCompletion(items []Completion, label string) (Completion, bool) {
	for _, item := range items {
		if item.Label == label {
			return item, true
		}
	}
	return Completion{}, false
}

func hasSymbol(symbols []Symbol, name string) bool {
	for _, symbol := range symbols {
		if symbol.Name == name {
			return true
		}
	}
	return false
}

func runnableProcedureNames(procedures []RunnableProcedure) []string {
	out := make([]string, 0, len(procedures))
	for _, procedure := range procedures {
		out = append(out, procedure.Name+":"+procedure.Kind)
	}
	return out
}

func hasDiagnostic(diagnostics []Diagnostic, code string) bool {
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}

func diagnosticsByCode(diagnostics []Diagnostic, code string) []Diagnostic {
	var out []Diagnostic
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			out = append(out, diagnostic)
		}
	}
	return out
}

func hasDiagnosticMessage(diagnostics []Diagnostic, text string) bool {
	for _, diagnostic := range diagnostics {
		if strings.Contains(diagnostic.Message, text) {
			return true
		}
	}
	return false
}

func hasSemanticToken(tokens []SemanticToken, tokenType, text, source string) bool {
	for _, token := range tokens {
		if token.Type == tokenType && rangeText(source, token.Range) == text {
			return true
		}
	}
	return false
}

func hasSemanticTokenAtLine(tokens []SemanticToken, tokenType string, line int) bool {
	for _, token := range tokens {
		if token.Type == tokenType && token.Range.Start.Line == line {
			return true
		}
	}
	return false
}

func rangeText(source string, r Range) string {
	line := lineAt(source, r.Start.Line)
	start := byteIndexForUTF16(line, r.Start.Character)
	end := byteIndexForUTF16(line, r.End.Character)
	if start < 0 || end < start || end > len(line) {
		return ""
	}
	return line[start:end]
}

func lineIndex(source, contains string) int {
	for i, line := range normalizedLines(source) {
		if strings.Contains(line, contains) {
			return i
		}
	}
	return -1
}
