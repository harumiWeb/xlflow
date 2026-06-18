package lint

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestLinterFindsMVPRules(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Sub Main()
Dim value
Range("A1").Select
ActiveCell.Activate
On Error Resume Next
End Sub
Public SharedState As String
Sub Prompt()
Application.GetOpenFilename
Application.GetSaveAsFilename
Application.FileDialog(msoFileDialogFilePicker).Show
InputBox "Path?"
MsgBox "Done"
UserForm1.Show
DoEvents
Shell "notepad.exe"
CreateObject("WScript.Shell").Popup "Done"
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	wantCodes := map[string]bool{"VB001": false, "VB002": false, "VB003": false, "VB004": false, "VB005": false, "VB006": false, "VB007": false}
	for _, issue := range issues {
		if _, ok := wantCodes[issue.Code]; ok {
			wantCodes[issue.Code] = true
		}
	}
	for code, found := range wantCodes {
		if !found {
			t.Fatalf("missing lint issue %s in %+v", code, issues)
		}
	}
	foundBoundaryMetadata := false
	foundDisableHint := false
	foundMsgBoxWrapperHint := false
	foundInputBoxWrapperHint := false
	foundGetOpenWrapperHint := false
	foundSaveAsWrapperHint := false
	foundFileDialogWrapperHint := false
	for _, issue := range issues {
		if issue.Code == "VB007" && issue.Kind != "" && issue.Symbol != "" && issue.Suggestion != "" {
			foundBoundaryMetadata = true
		}
		if issue.Code == "VB007" && strings.Contains(issue.Message, "[lint].forbid_interactive_input = false") {
			foundDisableHint = true
		}
		if issue.Code == "VB007" && issue.Symbol == "MsgBox" && strings.Contains(issue.Suggestion, "XlflowUI") && strings.Contains(issue.Message, "XlflowUI") {
			foundMsgBoxWrapperHint = true
		}
		if issue.Code == "VB007" && issue.Symbol == "InputBox" && strings.Contains(issue.Suggestion, "XlflowUI") && strings.Contains(issue.Message, "XlflowUI") {
			foundInputBoxWrapperHint = true
		}
		if issue.Code == "VB007" && issue.Symbol == "Application.GetOpenFilename" && strings.Contains(issue.Suggestion, "XlflowUI.GetOpenFilename") && strings.Contains(issue.Message, "XlflowUI") {
			foundGetOpenWrapperHint = true
		}
		if issue.Code == "VB007" && issue.Symbol == "Application.GetSaveAsFilename" && strings.Contains(issue.Suggestion, "XlflowUI.GetSaveAsFilename") && strings.Contains(issue.Message, "XlflowUI") {
			foundSaveAsWrapperHint = true
		}
		if issue.Code == "VB007" && issue.Symbol == "Application.FileDialog" && strings.Contains(issue.Suggestion, "XlflowUI.FileDialogOpen") && strings.Contains(issue.Message, "XlflowUI") {
			foundFileDialogWrapperHint = true
		}
	}
	if !foundBoundaryMetadata {
		t.Fatalf("expected VB007 to include GUI boundary metadata: %+v", issues)
	}
	if !foundDisableHint {
		t.Fatalf("expected VB007 to explain how to disable interactive-input lint: %+v", issues)
	}
	if !foundMsgBoxWrapperHint {
		t.Fatalf("expected VB007 to recommend XlflowUI for raw MsgBox usage: %+v", issues)
	}
	if !foundInputBoxWrapperHint {
		t.Fatalf("expected VB007 to recommend XlflowUI for raw InputBox usage: %+v", issues)
	}
	if !foundGetOpenWrapperHint {
		t.Fatalf("expected VB007 to recommend XlflowUI for raw GetOpenFilename usage: %+v", issues)
	}
	if !foundSaveAsWrapperHint {
		t.Fatalf("expected VB007 to recommend XlflowUI for raw GetSaveAsFilename usage: %+v", issues)
	}
	if !foundFileDialogWrapperHint {
		t.Fatalf("expected VB007 to recommend XlflowUI for raw FileDialog usage: %+v", issues)
	}
}

func TestLinterAllowsSelectCase(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Main()
Select Case 1
Case 1
End Select
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB002" {
			t.Fatalf("Select Case should not trigger VB002: %+v", issues)
		}
	}
}

func TestLinterUsesASTForDeclaratorsAndColumns(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public SharedState As String
Private c As String, d

Sub Main()
  Dim a, b As Long
  Dim localValue
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb005 := issuesByCode(issues, "VB005")
	if len(vb005) != 3 {
		t.Fatalf("expected three VB005 findings, got %+v", vb005)
	}
	assertIssue(t, vb005, "VB005", 3)
	assertIssue(t, vb005, "VB005", 6)
	assertIssue(t, vb005, "VB005", 7)
	a := findIssue(t, vb005, "VB005", 6)
	if a.Column != 7 {
		t.Fatalf("expected Dim a column 7, got %+v", a)
	}
	for _, issue := range vb005 {
		if issue.Line == 6 && issue.Column != 7 {
			t.Fatalf("Dim a should be the only line 6 implicit Variant, got %+v", vb005)
		}
	}
	vb006 := issuesByCode(issues, "VB006")
	if len(vb006) != 1 || vb006[0].Line != 2 {
		t.Fatalf("expected only module-level Public variable to trigger VB006, got %+v", vb006)
	}
}

func TestLinterASTIgnoresCommentsAndStringsForKeywordRules(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Main()
  Debug.Print ".Select .Activate On Error Resume Next"
  ' Range("A1").Select
  ' ActiveCell.Activate
  ' On Error Resume Next
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VB002", "VB003", "VB004"} {
		if got := issuesByCode(issues, code); len(got) != 0 {
			t.Fatalf("%s should ignore comments and strings, got %+v", code, got)
		}
	}
}

func TestLinterASTDetectsMemberAccessAndOnError(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Main()
  Range("A1").Select
  ActiveCell.Activate
  With Worksheets(1)
    .Range("A1").Select
  End With
  On Error Resume Next
  On Error GoTo ErrHandler
ErrHandler:
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb002 := issuesByCode(issues, "VB002")
	if len(vb002) != 2 {
		t.Fatalf("expected two Select findings, got %+v", vb002)
	}
	assertIssue(t, vb002, "VB002", 3)
	assertIssue(t, vb002, "VB002", 6)
	vb003 := issuesByCode(issues, "VB003")
	if len(vb003) != 1 || vb003[0].Line != 4 {
		t.Fatalf("expected one Activate finding, got %+v", vb003)
	}
	vb004 := issuesByCode(issues, "VB004")
	if len(vb004) != 1 || vb004[0].Line != 8 {
		t.Fatalf("expected only On Error Resume Next to trigger VB004, got %+v", vb004)
	}
}

func TestLinterReportsParserRecovery(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Main(
  Range("A1").Value = 1
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb014 := issuesByCode(issues, "VB014")
	if len(vb014) == 0 {
		t.Fatalf("expected parser recovery issue, got %+v", issues)
	}
	if vb014[0].Line == 0 || vb014[0].Column == 0 {
		t.Fatalf("expected parser recovery issue to include line and column, got %+v", vb014[0])
	}
	assertIssue(t, PushBlockingIssues(issues), "VB014", vb014[0].Line)
}

func TestLinterHandlesImplicitVariantsInsideUDTs(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Private Type TypedConfig
  Version As Long
  Label As String
End Type

Private Type UntypedConfig
  MissingField
End Type

Sub Main()
  Dim outsideValue
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Types.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb005Lines := make(map[int]bool)
	for _, issue := range issues {
		if issue.Code == "VB005" {
			vb005Lines[issue.Line] = true
		}
	}
	if vb005Lines[2] {
		t.Fatalf("typed UDT declaration should not trigger VB005: %+v", issues)
	}
	if vb005Lines[3] || vb005Lines[4] {
		t.Fatalf("typed UDT fields should not trigger VB005: %+v", issues)
	}
	if !vb005Lines[8] {
		t.Fatalf("untyped UDT field should trigger VB005: %+v", issues)
	}
	if !vb005Lines[12] {
		t.Fatalf("normal implicit variant outside UDT should still trigger VB005: %+v", issues)
	}
	if len(vb005Lines) != 2 {
		t.Fatalf("expected exactly two VB005 findings, got %+v", issues)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 0 {
		t.Fatalf("legal UDT implicit Variant fallback should not also trigger VB014: %+v", got)
	}
}

func TestLinterIgnoresConditionalCompilationDirectivesInsideUDTs(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Private Type NativeConfig
#If VBA7 Then
  Handle As LongPtr
#Else
  Handle As Long
#End If
  MissingField
End Type
`
	if err := os.WriteFile(filepath.Join(src, "Types.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb005Lines := make(map[int]bool)
	for _, issue := range issues {
		if issue.Code == "VB005" {
			vb005Lines[issue.Line] = true
		}
	}
	if vb005Lines[3] || vb005Lines[5] || vb005Lines[7] {
		t.Fatalf("conditional compilation directives inside UDT should not trigger VB005: %+v", issues)
	}
	if !vb005Lines[8] {
		t.Fatalf("untyped UDT field should still trigger VB005 after conditional directives: %+v", issues)
	}
	if len(vb005Lines) != 1 {
		t.Fatalf("expected exactly one VB005 finding, got %+v", issues)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 0 {
		t.Fatalf("legal UDT implicit Variant fallback should not also trigger VB014: %+v", got)
	}
}

func TestLinterAllowsInteractiveInputWhenDisabled(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Main()
Application.GetOpenFilename
InputBox "Path?"
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Lint.ForbidInteractiveInput = false
	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB007" {
			t.Fatalf("VB007 should be disabled: %+v", issues)
		}
	}
}

func TestLinterIgnoresXlflowUIWrappers(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String) As VbMsgBoxResult
  MsgBox = VBA.Interaction.MsgBox(Prompt)
End Function

Public Function InputBox(ByVal Id As String, ByVal Prompt As String) As String
  InputBox = VBA.Interaction.InputBox(Prompt)
End Function

Sub Main()
  Dim result As VbMsgBoxResult
  result = XlflowUI.MsgBox("confirm-save", "Done")
  Debug.Print XlflowUI.InputBox("customer-name", "Name")
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "XlflowUI.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB007" {
			t.Fatalf("wrapper helper should not trigger VB007: %+v", issues)
		}
	}
}

func TestLinterFindsTypographicQuotesThatTriggerVBECompileDialogs(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  If Mid$(text, index, 1) <> “\"\" Then\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	blocking := PushBlockingIssues(issues)
	if len(blocking) != 1 {
		t.Fatalf("expected one push-blocking typographic quote issue, got %+v", blocking)
	}
	if blocking[0].Code != "VB008" || blocking[0].Severity != "error" || blocking[0].Line != 3 {
		t.Fatalf("unexpected typographic quote issue: %+v", blocking[0])
	}
}

func TestLinterFindsLikelyCStyleQuoteEscapesThatTriggerVBECompileDialogs(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  If Mid$(text, index, 1) <> \"\\\"\" Then\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	blocking := PushBlockingIssues(issues)
	if len(blocking) != 1 {
		t.Fatalf("expected one push-blocking C-style escape issue, got %+v", blocking)
	}
	if blocking[0].Code != "VB009" || blocking[0].Severity != "error" || blocking[0].Line != 3 {
		t.Fatalf("unexpected C-style escape issue: %+v", blocking[0])
	}
}

func TestLinterAllowsValidProcedureBoundaries(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Sub Foo()
End Sub

Private Function Bar() As String
End Function

Friend Property Get Name() As String
End Property

Public Property Let Name(ByVal value As String)
End Property

Public Property Set Item(ByVal value As Object)
End Property

Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if strings.HasPrefix(issue.Code, "VB01") {
			t.Fatalf("valid procedure source should not trigger syntax lint: %+v", issues)
		}
	}
}

func TestLinterFindsProcedureBoundarySyntaxErrors(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Sub Unterminated()

End Function

End Property
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertIssue(t, issues, "VB012", 4)
	assertIssue(t, issues, "VB011", 6)

	blocking := PushBlockingIssues(issues)
	assertIssue(t, blocking, "VB012", 4)
	assertIssue(t, blocking, "VB011", 6)
}

func TestLinterFindsUnterminatedProcedureAtStartLine(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Function MissingClose() As String
    MissingClose = "x"
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	issue := findIssue(t, issues, "VB010", 2)
	if issue.Symbol == "" {
		t.Fatalf("expected VB010 to include procedure symbol: %+v", issue)
	}
	assertIssue(t, PushBlockingIssues(issues), "VB010", 2)
}

func TestLinterProcedureScannerIgnoresCommentsStringsAndDesignerEnd(t *testing.T) {
	dir := t.TempDir()
	formsDir := filepath.Join(dir, "src", "forms")
	if err := os.MkdirAll(formsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1
End
Attribute VB_Name = "UserForm1"
Option Explicit
' Sub Fake()
Public Sub Run()
    Debug.Print "End Sub"
End Sub
`
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.UserForm.CodeSource = "frm"
	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB010" || issue.Code == "VB011" || issue.Code == "VB012" {
			t.Fatalf("comments, strings, and designer End should not trigger procedure lint: %+v", issues)
		}
	}
}

func TestLinterHandlesContinuedProcedureDeclaration(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Sub Run( _
    ByVal value As String)
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB010" || issue.Code == "VB011" || issue.Code == "VB012" || issue.Code == "VB013" {
			t.Fatalf("valid continued declaration should not trigger syntax lint: %+v", issues)
		}
	}
}

func TestLinterFindsMissingWhitespaceBeforeLineContinuation(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Sub Run()
    Debug.Print "hello"_
    Debug.Print "hello" _
    Debug.Print "abc_"
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertIssue(t, issues, "VB013", 3)
	assertIssue(t, PushBlockingIssues(issues), "VB013", 3)
	var count int
	for _, issue := range issues {
		if issue.Code == "VB013" {
			count++
		}
	}
	if count != 1 {
		t.Fatalf("expected one VB013 issue, got %+v", issues)
	}
}

func TestLinterAllowsIdentifiersEndingWithUnderscore(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Sub Run()
    Dim total_ As Long
    total_ = 1
    Debug.Print total_
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB013" {
			t.Fatalf("identifier ending with underscore should not trigger VB013: %+v", issues)
		}
	}
}

func TestLinterHandlesOneLineProcedureStatements(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Foo(): End Sub
Function Bar() As String: Bar = "x": End Function
Property Get Name() As String: Name = "x": End Property
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB010" || issue.Code == "VB011" || issue.Code == "VB012" {
			t.Fatalf("one-line procedures should not trigger structure lint: %+v", issues)
		}
	}
}

func TestLinterSidecarModeSkipsGeneratedFRMCodeDiagnostics(t *testing.T) {
	dir := t.TempDir()
	formsDir := filepath.Join(dir, "src", "forms")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	body := "Option Explicit\nPublic Sub Run()\n  If Mid$(text, index, 1) <> \"\\\"\" Then\nEnd Sub\n"
	frmBody := "VERSION 5.00\nBegin {GUID} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\n" + body
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "UserForm1.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	var vb009 []Issue
	for _, issue := range issues {
		if issue.Code == "VB009" {
			vb009 = append(vb009, issue)
		}
	}
	if len(vb009) != 1 {
		t.Fatalf("expected one VB009 issue from sidecar mode, got %+v", vb009)
	}
	if vb009[0].File != "src/forms/code/UserForm1.bas" {
		t.Fatalf("expected sidecar file to be authoritative, got %+v", vb009[0])
	}
}

func TestLinterFindsDefaultASTBackedRules(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim a, b As Long
  Range("A1").Value = 1
  Worksheets(1).Range("A1").Value = 2
  On Error GoTo ErrHandler
  Debug.Print "work"
ErrHandler:
  Debug.Print Err.Description
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertIssue(t, issues, "VB015", 4)
	assertIssue(t, issues, "VB016", 8)
	assertIssue(t, issues, "VB019", 3)
}

func TestLinterAllowsQualifiedExcelAccessAndNarrowResumeNext(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(1)
  ws.Range("A1").Value = 1
  With ws
    .Cells(1, 1).Value = 1
  End With
  On Error Resume Next
  Err.Clear
  On Error GoTo 0
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VB004", "VB015"} {
		if got := issuesByCode(issues, code); len(got) != 0 {
			t.Fatalf("%s should not trigger for qualified/narrow pattern: %+v", code, got)
		}
	}
}

func TestLinterOptInProcedureLocalRules(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Private moduleValue As Long
Private Sub Used()
End Sub
Private Sub Unused()
End Sub
Public Sub Run()
  Dim moduleValue As Long
  Dim staleValue As Long
  Dim item As Long
  Application.EnableEvents = False
  ActiveSheet.Range("A1").Value = 1
  Set found = Range("A:A").Find("x")
  Debug.Print found.Value
  Foo (bar)
  For Each item In Range("A1:A2")
  Next item
  Resume Next
  Used
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectScopeShadowing = true
	cfg.Lint.DetectUnusedLocalVariables = true
	cfg.Lint.DetectUnusedPrivateProcedures = true
	cfg.Lint.DetectRangeFindNothingCheck = true

	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for code, line := range map[string]int{
		"VB017": 11,
		"VB018": 8,
		"VB020": 9,
		"VB021": 5,
		"VB022": 15,
		"VB023": 16,
		"VB024": 12,
		"VB025": 14,
		"VB026": 18,
	} {
		assertIssue(t, issues, code, line)
	}
}

func TestLinterDetectsNestedWithAmbiguityWhenEnabled(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  With ThisWorkbook
    With .Worksheets(1)
      .Range("A1").Value = 1
    End With
  End With
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectNestedWithAmbiguity = true
	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	assertIssue(t, issues, "VB027", 5)
}

func TestLinterNewASTRulesIgnoreCommentsAndStrings(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Debug.Print "Range(""A1"") On Error GoTo ErrHandler"
  ' Range("A1").Value = 1
  ' ErrHandler:
End Sub
`)
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VB015", "VB016"} {
		if got := issuesByCode(issues, code); len(got) != 0 {
			t.Fatalf("%s should ignore comments and strings: %+v", code, got)
		}
	}
}

func TestLinterSortsIssuesStably(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "B.bas", "Sub B()\nRange(\"A1\").Value = 1\nEnd Sub\n")
	writeLintModule(t, dir, "A.bas", "Sub A()\nRange(\"A1\").Value = 1\nEnd Sub\n")
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) < 2 {
		t.Fatalf("expected lint issues, got %+v", issues)
	}
	last := Issue{}
	for i, issue := range issues {
		if i > 0 && (last.File > issue.File ||
			(last.File == issue.File && last.Line > issue.Line) ||
			(last.File == issue.File && last.Line == issue.Line && last.Column > issue.Column) ||
			(last.File == issue.File && last.Line == issue.Line && last.Column == issue.Column && last.Code > issue.Code)) {
			t.Fatalf("issues not sorted: %+v", issues)
		}
		last = issue
	}
}

func assertIssue(t *testing.T, issues []Issue, code string, line int) {
	t.Helper()
	findIssue(t, issues, code, line)
}

func findIssue(t *testing.T, issues []Issue, code string, line int) Issue {
	t.Helper()
	for _, issue := range issues {
		if issue.Code == code && issue.Line == line {
			return issue
		}
	}
	t.Fatalf("missing issue %s at line %d in %+v", code, line, issues)
	return Issue{}
}

func issuesByCode(issues []Issue, code string) []Issue {
	var filtered []Issue
	for _, issue := range issues {
		if issue.Code == code {
			filtered = append(filtered, issue)
		}
	}
	return filtered
}

func writeLintModule(t *testing.T, dir, name, body string) {
	t.Helper()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(src, name), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}
