package lint

import (
	"os"
	"path/filepath"
	"strconv"
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

func TestLinterHonorsDisabledRuleIDs(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Range("A1").Select
  ActiveCell.Activate
End Sub
`)
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
disabled_rules = ["VB002"]
`)
	if err := os.WriteFile(filepath.Join(dir, config.FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB002"); len(got) != 0 {
		t.Fatalf("VB002 should be disabled: %+v", got)
	}
	if got := issuesByCode(issues, "VB003"); len(got) == 0 {
		t.Fatalf("VB003 should remain enabled: %+v", issues)
	}
}

func TestLinterSupportsInlineSuppressions(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ' xlflow:disable-next-line VB002
  Range("A1").Select
  Range("A2").Select ' xlflow:disable-line VB002
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB002"); len(got) != 0 {
		t.Fatalf("VB002 should be suppressed: %+v", got)
	}
}

func TestLinterSupportsMultipleInlineSuppressionIDs(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ' xlflow:disable-next-line VB002 VB003
  Range("A1").Select: ActiveCell.Activate
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB002"); len(got) != 0 {
		t.Fatalf("VB002 should be suppressed: %+v", got)
	}
	if got := issuesByCode(issues, "VB003"); len(got) != 0 {
		t.Fatalf("VB003 should be suppressed: %+v", got)
	}
}

func TestLinterInlineSuppressionKeepsUnrelatedSameLineDiagnostic(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ' xlflow:disable-next-line VB002
  Range("A1").Select: ActiveCell.Activate
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB002"); len(got) != 0 {
		t.Fatalf("VB002 should be suppressed: %+v", got)
	}
	if got := issuesByCode(issues, "VB003"); len(got) != 1 {
		t.Fatalf("VB003 should remain: %+v", issues)
	}
}

func TestLinterReportsUnknownAndUnusedInlineSuppressionsAsWarnings(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
  ' xlflow:disable-next-line VB999
  Debug.Print "ok"
  ' xlflow:disable-next-line VB002
  Debug.Print "still ok"
End Sub
`)

	result, err := Linter{RootDir: dir, Config: config.Default()}.RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Issues) != 0 {
		t.Fatalf("issues = %+v, want none", result.Issues)
	}
	if !hasWarning(result.Warnings, "unknown_inline_suppression_rule", "VB999") {
		t.Fatalf("expected unknown suppression warning, got %+v", result.Warnings)
	}
	if !hasWarning(result.Warnings, "unused_inline_suppression", "VB002") {
		t.Fatalf("expected unused suppression warning, got %+v", result.Warnings)
	}
}

func TestLinterDoesNotSuppressPreflightBlockingDiagnostics(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", "Option Explicit\nPublic Sub Run()\n  ' xlflow:disable-next-line VB008\n  Debug.Print “bad quote”\nEnd Sub\n")

	result, err := Linter{RootDir: dir, Config: config.Default()}.RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(result.Issues, "VB008"); len(got) != 1 {
		t.Fatalf("VB008 should remain unsuppressed: issues=%+v warnings=%+v", result.Issues, result.Warnings)
	}
	if !hasWarning(result.Warnings, "unsupported_inline_suppression_rule", "VB008") {
		t.Fatalf("expected unsupported suppression warning, got %+v", result.Warnings)
	}
}

func TestLinterConfigDisabledRulesComposeWithInlineSuppressions(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  ' xlflow:disable-next-line VB002
  Range("A1").Select
  ActiveCell.Activate
End Sub
`)
	cfg := config.Default()
	cfg.Lint.ForbidSelect = false
	cfg.Lint.DisabledRules = []string{"VB002"}

	result, err := Linter{RootDir: dir, Config: cfg}.RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(result.Issues, "VB002"); len(got) != 0 {
		t.Fatalf("VB002 should be globally disabled: %+v", got)
	}
	if got := issuesByCode(result.Issues, "VB003"); len(got) == 0 {
		t.Fatalf("VB003 should remain enabled: %+v", result.Issues)
	}
	if !hasWarning(result.Warnings, "unused_inline_suppression", "VB002") {
		t.Fatalf("expected unused inline warning for globally disabled VB002, got %+v", result.Warnings)
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

func TestLinterBlocksBareDialogsWhenXlflowUIIsPresent(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	uiBody := `Attribute VB_Name = "XlflowUI"
Option Explicit
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String) As VbMsgBoxResult
End Function
Public Function InputBox(ByVal Id As String, ByVal Prompt As String) As String
End Function
`
	if err := os.WriteFile(filepath.Join(src, "XlflowUI.bas"), []byte(uiBody), 0o644); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
  MsgBox "Done"
  Call MsgBox("Done")
  Dim answer As String
  answer = InputBox("Name?")
  XlflowUI.MsgBox "confirm", "OK"
  VBA.Interaction.MsgBox "Native"
  Debug.Print "MsgBox in a string"
  ' InputBox "comment"
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
	vb028 := issuesByCode(issues, "VB028")
	if len(vb028) != 3 {
		t.Fatalf("expected three XlflowUI name-collision issues, got %+v", vb028)
	}
	assertIssue(t, vb028, "VB028", 5)
	assertIssue(t, vb028, "VB028", 6)
	assertIssue(t, vb028, "VB028", 8)
	for _, issue := range vb028 {
		if issue.Severity != "error" || issue.Kind != "xlflowui_name_collision" {
			t.Fatalf("unexpected VB028 metadata: %+v", issue)
		}
		if !strings.Contains(issue.Message, "VBA.Interaction."+issue.Symbol) || !strings.Contains(issue.Suggestion, "XlflowUI."+issue.Symbol) {
			t.Fatalf("VB028 should explain both supported remedies: %+v", issue)
		}
	}
	blocking := issuesByCode(PushBlockingIssues(issues), "VB028")
	assertIssue(t, blocking, "VB028", 5)
	assertIssue(t, blocking, "VB028", 6)
	assertIssue(t, blocking, "VB028", 8)
}

func TestLinterVB028UsesStatementCallContext(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	uiBody := `Attribute VB_Name = "XlflowUI"
Option Explicit
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String) As VbMsgBoxResult
End Function
`
	if err := os.WriteFile(filepath.Join(src, "XlflowUI.bas"), []byte(uiBody), 0o644); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(): MsgBox "Done": End Sub

Private Sub Assignments()
  Dim MsgBox As String
  MsgBox = "not a dialog call"
  Debug.Print MsgBox
End Sub

Public Function NativeName() As String
  NativeName = "ok"
End Function
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
	vb028 := issuesByCode(issues, "VB028")
	if len(vb028) != 1 {
		t.Fatalf("expected only the one-line procedure dialog call to trigger VB028, got %+v", vb028)
	}
	assertIssue(t, vb028, "VB028", 4)
}

func TestLinterAllowsBareDialogsWithoutXlflowUI(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  MsgBox "Done"
  Debug.Print InputBox("Name?")
End Sub
`)
	cfg := config.Default()
	cfg.Lint.ForbidInteractiveInput = false
	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB028"); len(got) != 0 {
		t.Fatalf("VB028 should require XlflowUI.bas to be present, got %+v", got)
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
	blocking := issuesByCode(PushBlockingIssues(issues), "VB008")
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
	blocking := issuesByCode(PushBlockingIssues(issues), "VB009")
	if len(blocking) != 1 {
		t.Fatalf("expected one push-blocking C-style escape issue, got %+v", blocking)
	}
	if blocking[0].Code != "VB009" || blocking[0].Severity != "error" || blocking[0].Line != 3 {
		t.Fatalf("unexpected C-style escape issue: %+v", blocking[0])
	}
}

func TestLinterKeepsEarlierCStyleQuoteEscapeWhenLaterQuoteExists(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", "Option Explicit\nPublic Sub Run()\n  s = \"\\\"\": Debug.Print \"x\"\nEnd Sub\n")
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	blocking := issuesByCode(PushBlockingIssues(issues), "VB009")
	if len(blocking) != 1 {
		t.Fatalf("expected one push-blocking C-style escape issue, got %+v", blocking)
	}
	if blocking[0].Line != 3 {
		t.Fatalf("unexpected C-style escape issue: %+v", blocking[0])
	}
}

func TestLinterAllowsVBAJSONEscapedQuoteStrings(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "JsonConverter.bas", `Attribute VB_Name = "JsonConverter"
Option Explicit
Private Function JsonEscape(ByVal json_Char As String) As String
  Select Case AscW(json_Char)
  Case 34
    json_Char = "\"""
  Case 92
    json_Char = "\\"
  End Select
  JsonEscape = json_Char
End Function
`)
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB009" {
			t.Fatalf("valid VBA-JSON escaped quote strings should not trigger VB009: %+v", issues)
		}
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

func TestLinterFindsVBAContinuationLineOverflow(t *testing.T) {
	tests := []struct {
		name          string
		body          string
		line          int
		kind          string
		symbol        string
		continuations int
	}{
		{
			name: "procedure declaration",
			body: "Attribute VB_Name = \"Main\"\nOption Explicit\n" + continuedLogicalLine(
				"Public Sub TooMany(ByVal a0 As Long,", "    ByVal arg As Long,", "    ByVal finalArg As Long)\nEnd Sub\n", 25),
			line:          3,
			kind:          "procedure_declaration",
			symbol:        "TooMany",
			continuations: 25,
		},
		{
			name: "procedure call",
			body: "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n" + continuedLogicalLine(
				"    Call TooMany(a0,", "        arg,", "        finalArg)\nEnd Sub\n", 25),
			line:          4,
			kind:          "procedure_call",
			symbol:        "TooMany",
			continuations: 25,
		},
		{
			name: "generic logical line",
			body: "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n" + continuedLogicalLine(
				"    result = a0 +", "        arg +", "        finalArg\nEnd Sub\n", 27),
			line:          4,
			kind:          "logical_line",
			continuations: 27,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			writeLintModule(t, dir, "Main.bas", tc.body)
			issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
			if err != nil {
				t.Fatal(err)
			}
			vb015 := issuesByCode(issues, "VB015")
			if len(vb015) != 1 {
				t.Fatalf("expected one VB015 issue, got %+v", vb015)
			}
			issue := vb015[0]
			if issue.Line != tc.line || issue.Kind != tc.kind || issue.Symbol != tc.symbol {
				t.Fatalf("unexpected VB015 metadata: %+v", issue)
			}
			if !strings.Contains(issue.Message, "uses "+strconv.Itoa(tc.continuations)+" continuation lines") || !strings.Contains(issue.Message, "at most 24") || issue.Suggestion == "" {
				t.Fatalf("unexpected VB015 diagnostic: %+v", issue)
			}
			assertIssue(t, PushBlockingIssues(issues), "VB015", tc.line)
		})
	}
}

func TestLinterAllowsVBAContinuationLineLimitAndIgnoresStringsAndComments(t *testing.T) {
	dir := t.TempDir()
	body := "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n" +
		continuedLogicalLine("    result = a0 +", "        arg +", "        finalArg\n", 24) +
		continuedLogicalLine("    result = b0 +", "        arg +", "        finalArg\n", 24) +
		strings.Repeat("    Debug.Print \"_\"\n", 30) +
		strings.Repeat("    ' _\n", 30) +
		"End Sub\n"
	writeLintModule(t, dir, "Main.bas", body)
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB015"); len(got) != 0 {
		t.Fatalf("valid continuation limits, strings, and comments should not trigger VB015: %+v", got)
	}
}

func TestLinterFindsRepeatedQuestionShorthand(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Public Sub Run()
    ?? "hoge"
    ??? "fuga"
    ? "ok"
    Debug.Print "??"
    ' ?? "comment"
    Debug.Print "ok": ?? "after colon"
StartLabel: ?? "after label"
10  ?? "after line number"
    Debug.Print _
        ?? "continued expression"
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb032 := issuesByCode(issues, "VB032")
	if len(vb032) != 5 {
		t.Fatalf("expected five VB032 findings, got %+v", vb032)
	}
	assertIssue(t, vb032, "VB032", 3)
	assertIssue(t, vb032, "VB032", 4)
	colonIssue := findIssue(t, vb032, "VB032", 8)
	if colonIssue.Column == 0 {
		t.Fatalf("expected VB032 to include a column, got %+v", colonIssue)
	}
	assertIssue(t, vb032, "VB032", 9)
	lineNumberIssue := findIssue(t, vb032, "VB032", 10)
	if lineNumberIssue.Column != 5 {
		t.Fatalf("expected line-number-prefixed VB032 at column 5, got %+v", lineNumberIssue)
	}
	blocking := issuesByCode(PushBlockingIssues(issues), "VB032")
	if len(blocking) != 5 {
		t.Fatalf("VB032 should be push-blocking, got %+v", blocking)
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
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
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
	for _, code := range []string{"VB004"} {
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
  Dim Item As Long
  Application.EnableEvents = False
  ActiveSheet.Range("A1").Value = 1
  Set found = Range("A:A").Find("x")
  Debug.Print found.Value
  Foo (bar)
  For Each Item In Range("A1:A2")
  Next Item
  Resume Next
  Used
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectScopeShadowing = true
	cfg.Lint.DetectUnusedLocalVariables = true
	cfg.Lint.DetectUnusedPrivateProcedures = true

	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	for code, line := range map[string]int{
		"VB018": 8,
		"VB020": 9,
		"VB021": 5,
		"VB022": 15,
		"VB023": 16,
		"VB026": 18,
	} {
		assertIssue(t, issues, code, line)
	}
	if issue := findIssue(t, issues, "VB023", 16); issue.Symbol != "Item" {
		t.Fatalf("expected VB023 to preserve declaration casing, got %+v", issue)
	}
}

func TestLinterUnusedLocalVariableUsesProcedureBounds(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim total As Long
  total = 1
  Debug.Print total
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectUnusedLocalVariables = true

	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB020"); len(got) != 0 {
		t.Fatalf("VB020 should scan to the enclosing procedure end: %+v", got)
	}
}

func TestLinterUnusedLocalVariableIgnoresWriteOnlyAssignments(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim writtenOnly As Long
  Dim assignedObject As Object
  Dim updated As Long
  writtenOnly = 1
  Set assignedObject = Nothing
  updated = updated + 1
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectUnusedLocalVariables = true

	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb020 := issuesByCode(issues, "VB020")
	if len(vb020) != 2 || vb020[0].Symbol != "writtenOnly" || vb020[1].Symbol != "assignedObject" {
		t.Fatalf("expected only write-only assignments to trigger VB020, got %+v", vb020)
	}
}

func TestLinterUnusedLocalVariableTreatsOneLineConditionalAssignmentsAsWrites(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim scalarValue As Long
  Dim objectValue As Object
  If True Then scalarValue = 1
  If True Then Set objectValue = Nothing
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectUnusedLocalVariables = true

	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb020 := issuesByCode(issues, "VB020")
	if len(vb020) != 2 || vb020[0].Symbol != "scalarValue" || vb020[1].Symbol != "objectValue" {
		t.Fatalf("expected one-line conditional write-only assignments to trigger VB020, got %+v", vb020)
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
	writeLintModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
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
	if len(issues) != 0 {
		t.Fatalf("comments and strings should not trigger lint issues: %+v", issues)
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

func TestLinterLintSourceUsesUnsavedContent(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte("Sub Run()\n    Range(\"A1\").Select\n    Dim value\nEnd Sub\n")

	issues, err := Linter{RootDir: dir, Config: config.Default()}.LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	assertIssue(t, issues, "VB001", 1)
	assertIssue(t, issues, "VB002", 2)
	assertIssue(t, issues, "VB005", 3)
	assertIssue(t, issues, "VB020", 3)
}

func TestLinterLintSourceAppliesInlineSuppressions(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte(`Option Explicit
Public Sub Run()
    Dim x As Integer ' xlflow:disable-line VB020
    x = 2
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB020"); len(got) != 0 {
		t.Fatalf("VB020 should be suppressed for in-memory lint source, got %+v", got)
	}
}

func TestLinterReportsUndeclaredAssignmentsWithOptionExplicit(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Private moduleValue As Long

Public Function Build(ByVal inputValue As Long) As Long
  Dim localValue As Long
  localValue = inputValue
  moduleValue = localValue
  Build = moduleValue
  Range("A1") = moduleValue
  missingValue = 1
  For index = 1 To 3
  Next index
End Function
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb029 := issuesByCode(issues, "VB029")
	if len(vb029) != 2 {
		t.Fatalf("expected two undeclared variable issues, got %+v", vb029)
	}
	missing := findIssue(t, vb029, "VB029", 10)
	if missing.Symbol != "missingValue" || missing.Kind != "undeclared_variable" || missing.Column != 3 {
		t.Fatalf("unexpected missingValue issue: %+v", missing)
	}
	index := findIssue(t, vb029, "VB029", 11)
	if index.Symbol != "index" || index.Column != 7 {
		t.Fatalf("unexpected For index issue: %+v", index)
	}
	assertIssue(t, PushBlockingIssues(issues), "VB029", 10)
}

func TestLinterAllowsModuleVariablesDeclaredInsideConditionalCompilationBlocks(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Private Const TimerIntervalMs As Long = 180

#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal uIDEvent As LongPtr) As Long
Private mTimerId As LongPtr
#Else
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal uIDEvent As Long) As Long
Private mTimerId As Long
#End If

Public Sub StartLoop()
    If mTimerId <> 0 Then
        Exit Sub
    End If

    #If VBA7 Then
    mTimerId = SetTimer(0, 0, TimerIntervalMs, AddressOf MazeChaseTimerProc)
    #Else
    mTimerId = SetTimer(0, 0, TimerIntervalMs, AddressOf MazeChaseTimerProc)
    #End If
End Sub

Public Sub StopLoop()
    If mTimerId = 0 Then
        Exit Sub
    End If

    KillTimer 0, mTimerId
    mTimerId = 0
    stillMissing = 1
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb029 := issuesByCode(issues, "VB029")
	if len(vb029) != 1 {
		t.Fatalf("expected only the truly undeclared assignment to trigger VB029, got %+v", vb029)
	}
	if vb029[0].Symbol != "stillMissing" {
		t.Fatalf("mTimerId declared inside conditional compilation should not trigger VB029: %+v", vb029)
	}
}

func TestLinterUndeclaredAssignmentsRequireOptionExplicit(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Public Sub Run()
  missingValue = 1
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB029"); len(got) != 0 {
		t.Fatalf("VB029 should require Option Explicit in the source, got %+v", got)
	}
}

func TestLinterRequiresVBNameAttributeForStandardModules(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	vb031 := issuesByCode(issues, "VB031")
	if len(vb031) != 1 {
		t.Fatalf("expected one missing VB_Name issue, got %+v", vb031)
	}
	issue := vb031[0]
	if issue.Line != 1 || issue.Kind != "missing_module_attribute" || issue.Symbol != "VB_Name" || !strings.Contains(issue.Suggestion, `Attribute VB_Name = "Main"`) {
		t.Fatalf("unexpected VB031 issue: %+v", issue)
	}
	assertIssue(t, PushBlockingIssues(issues), "VB031", 1)
}

func TestLinterAcceptsVBNameAttributeForStandardModules(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB031"); len(got) != 0 {
		t.Fatalf("VB031 should not be reported when Attribute VB_Name is present: %+v", got)
	}
}

func TestLinterRejectsEmptyVBNameAttributeForStandardModules(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Attribute VB_Name = ""
Option Explicit
Public Sub Run()
End Sub
`)

	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB031"); len(got) != 1 {
		t.Fatalf("VB031 should be reported for empty Attribute VB_Name: %+v", got)
	}
}

func TestLinterVBNameAttributeOnlyAppliesToStandardModules(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	classes := filepath.Join(dir, "src", "classes")
	forms := filepath.Join(dir, "src", "forms")
	sidecars := filepath.Join(forms, "code")
	workbook := filepath.Join(dir, "src", "workbook")
	for _, path := range []string{classes, forms, sidecars, workbook} {
		if err := os.MkdirAll(path, 0o755); err != nil {
			t.Fatal(err)
		}
	}
	if err := os.WriteFile(filepath.Join(classes, "Class1.cls"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(forms, "UserForm1.frm"), []byte("VERSION 5.00\nBegin VB.Form UserForm1\nEnd\nOption Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(sidecars, "UserForm1.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "ThisWorkbook.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "Sheet1.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB031"); len(got) != 0 {
		t.Fatalf("VB031 should not apply to class modules, forms, UserForm sidecars, or document modules: %+v", got)
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

func hasWarning(warnings []map[string]any, code string, rule string) bool {
	for _, warning := range warnings {
		if warning["code"] == code && warning["rule"] == rule {
			return true
		}
	}
	return false
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

func continuedLogicalLine(first, middle, last string, continuations int) string {
	var b strings.Builder
	for i := 0; i < continuations; i++ {
		if i == 0 {
			b.WriteString(first)
		} else {
			b.WriteString(middle)
		}
		b.WriteString(" _\n")
	}
	b.WriteString(last)
	return b.String()
}
