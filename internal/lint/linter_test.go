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
	foundDialogWrapperHint := false
	for _, issue := range issues {
		if issue.Code == "VB007" && issue.Kind != "" && issue.Symbol != "" && issue.Suggestion != "" {
			foundBoundaryMetadata = true
		}
		if issue.Code == "VB007" && strings.Contains(issue.Message, "[lint].forbid_interactive_input = false") {
			foundDisableHint = true
		}
		if issue.Code == "VB007" && (issue.Symbol == "MsgBox" || issue.Symbol == "InputBox") && strings.Contains(issue.Suggestion, "XlflowUI") && strings.Contains(issue.Message, "XlflowUI") {
			foundDialogWrapperHint = true
		}
	}
	if !foundBoundaryMetadata {
		t.Fatalf("expected VB007 to include GUI boundary metadata: %+v", issues)
	}
	if !foundDisableHint {
		t.Fatalf("expected VB007 to explain how to disable interactive-input lint: %+v", issues)
	}
	if !foundDialogWrapperHint {
		t.Fatalf("expected VB007 to recommend XlflowUI for raw MsgBox/InputBox usage: %+v", issues)
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
