package lint

import (
	"context"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"testing"
	"time"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestLintParsedContextReturnsCancellationWithoutPartialIssues(t *testing.T) {
	doc, err := vbaast.ParseDocument("Main.bas", []byte("Sub Main()\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	issues, err := (Linter{}).LintParsedContext(ctx, doc)
	if !errors.Is(err, context.Canceled) || issues != nil {
		t.Fatalf("canceled result = (%+v, %v)", issues, err)
	}
}

func TestLintParsedContextCancelsInFlightAndReleasesParsedDocument(t *testing.T) {
	var source strings.Builder
	for i := 0; i < 1200; i++ {
		source.WriteString("Private Sub Work" + strconv.Itoa(i) + "()\n")
		source.WriteString("  Dim value As Long\n")
		source.WriteString("  value = " + strconv.Itoa(i) + "\n")
		source.WriteString("End Sub\n")
	}
	doc, err := vbaast.ParseDocument("Large.bas", []byte(source.String()))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	ctx := &lintCheckpointContext{cancelAt: 16}
	issues, err := (Linter{}).LintParsedContext(ctx, doc)
	if !errors.Is(err, context.Canceled) || issues != nil {
		t.Fatalf("in-flight canceled result = (%+v, %v)", issues, err)
	}
	if ctx.checks < ctx.cancelAt {
		t.Fatalf("cancellation checks = %d, want at least %d", ctx.checks, ctx.cancelAt)
	}

	readDone := make(chan error, 1)
	go func() { readDone <- doc.Read(func(vbaast.ParsedView) error { return nil }) }()
	select {
	case err := <-readDone:
		if err != nil {
			t.Fatalf("read after cancellation: %v", err)
		}
	case <-time.After(time.Second):
		t.Fatal("ParsedDocument.Read remained locked after cancellation")
	}
}

type lintCheckpointContext struct {
	checks   int
	cancelAt int
}

func (c *lintCheckpointContext) Deadline() (time.Time, bool) { return time.Time{}, false }
func (c *lintCheckpointContext) Done() <-chan struct{}       { return nil }
func (c *lintCheckpointContext) Value(any) any               { return nil }
func (c *lintCheckpointContext) Err() error {
	c.checks++
	if c.checks >= c.cancelAt {
		return context.Canceled
	}
	return nil
}

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

func TestLinterAcceptsBoundedErrNumberProbes(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit
Sub Main()
  Dim capacity As Long
  On Error Resume Next
  capacity = UBound(Array()) + 1
  If Err.Number <> 0 Then
    Err.Clear
    capacity = 0
  End If
  On Error GoTo 0
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB004"); len(got) != 0 {
		t.Fatalf("bounded Err.Number probe should not trigger VB004: %+v", got)
	}
}

func TestLinterReportsUnobservedOnErrorCleanupScopes(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit
Sub Main()
  On Error Resume Next
  Debug.Print 1
  Debug.Print 2
  Debug.Print 3
  Debug.Print 4
  Debug.Print 5
  Debug.Print 6
  On Error GoTo 0
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB004"); len(got) != 1 {
		t.Fatalf("unobserved cleanup scope should retain VB004: %+v", got)
	}
}

func TestConfusingParenthesizedCallIgnoresFunctionArguments(t *testing.T) {
	if name, ok := confusingParenthesizedCall("Mid$(alphabet, (block \\ 64) + 1, 1)"); ok || name != "" {
		t.Fatalf("Mid$ argument expression = (%q, %t), want no VB022", name, ok)
	}
	if name, ok := confusingParenthesizedCall("Run (value)"); !ok || name != "Run" {
		t.Fatalf("ambiguous call = (%q, %t), want Run", name, ok)
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
	if len(vb014) != 1 {
		t.Fatalf("expected one parser recovery issue, got %+v", issues)
	}
	issue := vb014[0]
	if issue.Line == 0 || issue.Column == 0 {
		t.Fatalf("expected parser recovery issue to include line and column, got %+v", issue)
	}
	if issue.Kind != "parser_recovery" || issue.ParserNode != "ERROR" || issue.ParserToken == "" || issue.Context == "" {
		t.Fatalf("expected ERROR recovery context, got %+v", issue)
	}
	if issue.Message != parserRecoveryMessage || issue.Suggestion != parserRecoverySuggestion || strings.Contains(strings.ToLower(issue.Message), "syntax error") {
		t.Fatalf("expected neutral parser recovery guidance, got %+v", issue)
	}
	encoded, err := json.Marshal(issue)
	if err != nil {
		t.Fatal(err)
	}
	var payload map[string]any
	if err := json.Unmarshal(encoded, &payload); err != nil {
		t.Fatal(err)
	}
	for field, want := range map[string]string{
		"parser_node":  issue.ParserNode,
		"parser_token": issue.ParserToken,
		"context":      issue.Context,
	} {
		if got, ok := payload[field].(string); !ok || got != want {
			t.Fatalf("JSON %s = %#v, want %q", field, payload[field], want)
		}
	}
	assertIssue(t, PushBlockingIssues(issues), "VB014", issue.Line)
}

func TestLinterReportsMissingParserRecoveryContext(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte(`Option Explicit
Sub Main()
  Debug.Print (1
End Sub
`)

	issues, err := (Linter{RootDir: dir, Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	vb014 := issuesByCode(issues, "VB014")
	if len(vb014) != 1 {
		t.Fatalf("expected one parser recovery issue, got %+v", issues)
	}
	issue := vb014[0]
	if issue.Kind != "parser_recovery" || issue.ParserNode != "MISSING" || issue.ParserToken == "" {
		t.Fatalf("expected MISSING recovery detail, got %+v", issue)
	}
	if issue.Context != "Debug.Print (1" || issue.Line != 3 || issue.Column == 0 {
		t.Fatalf("expected missing recovery location and source context, got %+v", issue)
	}
	assertIssue(t, PushBlockingIssues(issues), "VB014", issue.Line)
}

func TestLinterReportsLikelyUnclosedBlocks(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	tests := []struct {
		name         string
		body         string
		kind         string
		closer       string
		openingLine  int
		expectedLine int
	}{
		{
			name: "multiline If",
			body: "Option Explicit\nSub Main()\n  If enabled Then\n    Debug.Print \"x\"\nEnd Sub\n",
			kind: "if", closer: "End If", openingLine: 3, expectedLine: 5,
		},
		{
			name: "For Each",
			body: "Option Explicit\nSub Main()\n  For Each item In items\n    Debug.Print item\nEnd Sub\n",
			kind: "for", closer: "Next", openingLine: 3, expectedLine: 5,
		},
		{
			name: "Do Loop",
			body: "Option Explicit\nSub Main()\n  Do While ready\n    Debug.Print \"x\"\nEnd Sub\n",
			kind: "do", closer: "Loop", openingLine: 3, expectedLine: 5,
		},
		{
			name: "While Wend",
			body: "Option Explicit\nSub Main()\n  While ready\n    Debug.Print \"x\"\nEnd Sub\n",
			kind: "while", closer: "Wend", openingLine: 3, expectedLine: 5,
		},
		{
			name: "With End With",
			body: "Option Explicit\nSub Main()\n  With Application\n    .ScreenUpdating = False\nEnd Sub\n",
			kind: "with", closer: "End With", openingLine: 3, expectedLine: 5,
		},
		{
			name: "Select Case",
			body: "Option Explicit\nSub Main()\n  Select Case value\n  Case 1\n    Debug.Print \"x\"\nEnd Sub\n",
			kind: "select", closer: "End Select", openingLine: 3, expectedLine: 6,
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, []byte(tc.body))
			if err != nil {
				t.Fatal(err)
			}
			vb014 := issuesByCode(issues, "VB014")
			if len(vb014) != 1 {
				t.Fatalf("VB014 issues = %+v, want one targeted issue", vb014)
			}
			issue := vb014[0]
			if issue.Kind != "parser_recovery" || issue.BlockKind != tc.kind || issue.ExpectedCloser != tc.closer || issue.OpeningLine != tc.openingLine || issue.Line != tc.expectedLine {
				t.Fatalf("unexpected targeted recovery issue: %+v", issue)
			}
			if !strings.Contains(issue.Message, "Possible missing '"+tc.closer+"'") || !strings.Contains(issue.Message, "opened at line "+strconv.Itoa(tc.openingLine)) {
				t.Fatalf("targeted message = %q", issue.Message)
			}
			assertIssue(t, PushBlockingIssues(issues), "VB014", tc.expectedLine)
		})
	}
}

func TestLinterLocatesNestedUnclosedBlockAtParentCloser(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	source := []byte("Option Explicit\nSub Main()\n  If outerReady Then\n    If innerReady Then\n      Debug.Print \"x\"\n  End If\nEnd Sub\n")
	issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	vb014 := issuesByCode(issues, "VB014")
	if len(vb014) != 1 {
		t.Fatalf("VB014 issues = %+v, want one nested-block diagnostic", vb014)
	}
	issue := vb014[0]
	if issue.BlockKind != "if" || issue.ExpectedCloser != "End If" || issue.OpeningLine != 4 || issue.OpeningColumn != 5 || issue.Line != 6 || issue.Column != 3 {
		t.Fatalf("nested-block diagnostic = %+v, want inner If at its parent End If", issue)
	}
	if !strings.Contains(issue.Message, "opened at line 4") {
		t.Fatalf("nested-block message = %q", issue.Message)
	}
}

func TestLinterPreservesContinuationTailLocationForUnclosedBlock(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	source := []byte("Sub Main()\n  result = _\n    1: If ready Then\n      Debug.Print \"x\"\nEnd Sub\n")
	issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	vb014 := issuesByCode(issues, "VB014")
	if len(vb014) != 1 {
		t.Fatalf("VB014 issues = %+v, want one targeted issue", vb014)
	}
	issue := vb014[0]
	if issue.BlockKind != "if" || issue.OpeningLine != 3 || issue.OpeningColumn != 8 || issue.Line != 5 || issue.Column != 1 {
		t.Fatalf("continuation-tail diagnostic location = %+v", issue)
	}
}

func TestUnmatchedBlockCandidatesStayConservative(t *testing.T) {
	valid := "Sub Main()\n  If ready Then\n    For Each item In items\n      Debug.Print item\n    Next item\n  End If\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(valid); !reliable || len(candidates) != 0 {
		t.Fatalf("valid source candidates = %+v, reliable=%t", candidates, reliable)
	}

	singleLineIf := "Sub Main()\n  If ready Then Debug.Print \"End If\"\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(singleLineIf); !reliable || len(candidates) != 0 {
		t.Fatalf("single-line If candidates = %+v, reliable=%t", candidates, reliable)
	}

	colonClosed := "Sub Main()\n  For Each item In items: Debug.Print item: Next item\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(colonClosed); !reliable || len(candidates) != 0 {
		t.Fatalf("colon-closed block candidates = %+v, reliable=%t", candidates, reliable)
	}

	sharedNext := "Sub Main()\n  For Each outerItem In items\n    For Each innerItem In items\n      Debug.Print innerItem\n    Next innerItem, outerItem\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(sharedNext); !reliable || len(candidates) != 0 {
		t.Fatalf("shared Next candidates = %+v, reliable=%t", candidates, reliable)
	}

	loopUntil := "Sub Main()\n  Do While ready\n    Debug.Print \"x\"\n  Loop Until finished\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(loopUntil); !reliable || len(candidates) != 0 {
		t.Fatalf("Loop Until candidates = %+v, reliable=%t", candidates, reliable)
	}

	nextPrefixedIdentifiers := "Sub Main()\n  nextChar = Mid$(text, position, 1)\n  NextHashCapacity = result\n  nextSlot = (slot + 1) And mask\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(nextPrefixedIdentifiers); !reliable || len(candidates) != 0 {
		t.Fatalf("Next-prefixed identifier candidates = %+v, reliable=%t", candidates, reliable)
	}

	continuation := "Sub Main()\n  If ready _\n      And enabled Then\n    Debug.Print \"x\"\nEnd Sub\n"
	candidates, reliable := unmatchedBlockCandidates(continuation)
	if !reliable || len(candidates) != 1 || candidates[0].kind != "if" || candidates[0].openingLine != 2 || candidates[0].expectedLine != 5 {
		t.Fatalf("continued If candidates = %+v, reliable=%t", candidates, reliable)
	}

	nestedParentCloser := "Sub Main()\n  If outerReady Then\n    If innerReady Then\n      Debug.Print \"x\"\n  End If\nEnd Sub\n"
	candidates, reliable = unmatchedBlockCandidates(nestedParentCloser)
	if !reliable || len(candidates) != 1 || candidates[0].kind != "if" || candidates[0].openingLine != 3 || candidates[0].expectedLine != 5 || candidates[0].expectedColumn != 3 {
		t.Fatalf("nested parent-closer candidates = %+v, reliable=%t", candidates, reliable)
	}

	continuationTail := "Sub Main()\n  result = _\n    1: If ready Then\n      Debug.Print \"x\"\nEnd Sub\n"
	candidates, reliable = unmatchedBlockCandidates(continuationTail)
	if !reliable || len(candidates) != 1 || candidates[0].openingLine != 3 || candidates[0].openingColumn != 8 || candidates[0].expectedLine != 5 || candidates[0].expectedColumn != 1 {
		t.Fatalf("continuation-tail candidates = %+v, reliable=%t", candidates, reliable)
	}

	continuedRem := "Sub Main()\n  Rem note _\n    more: If ready Then\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(continuedRem); !reliable || len(candidates) != 0 {
		t.Fatalf("continued Rem candidates = %+v, reliable=%t", candidates, reliable)
	}

	continuedSingleLineIf := "Sub Main()\n  If ready Then _\n    : Debug.Print \"x\"\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(continuedSingleLineIf); !reliable || len(candidates) != 0 {
		t.Fatalf("continued single-line If candidates = %+v, reliable=%t", candidates, reliable)
	}

	conditional := "Sub Main()\n#If VBA7 Then\n  If ready Then\n#End If\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(conditional); reliable || len(candidates) != 0 {
		t.Fatalf("conditional source candidates = %+v, reliable=%t", candidates, reliable)
	}

	comment := "Sub Main()\n  ' If ready Then\n  Rem For Each item In items\nEnd Sub\n"
	if candidates, reliable := unmatchedBlockCandidates(comment); !reliable || len(candidates) != 0 {
		t.Fatalf("comment source candidates = %+v, reliable=%t", candidates, reliable)
	}
}

func TestLinterAcceptsNextPrefixedIdentifiersWithoutParserRecovery(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "classes", "Example.cls")
	source := []byte("Option Explicit\nPrivate Sub ReadNext()\n  Dim nextChar As String\n  Dim nextSlot As Long\n  nextChar = Mid$(text, position, 1)\n  nextSlot = (slot + 1) And mask\nEnd Sub\nPrivate Function NextHashCapacity() As Long\n  NextHashCapacity = 1\nEnd Function\n")
	issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 0 {
		t.Fatalf("Next-prefixed identifiers should not trigger VB014: %+v", got)
	}
}

func TestLinterFallsBackToGenericRecoveryForConditionalCompilation(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	source := []byte("Option Explicit\nSub Main()\n#If VBA7 Then\n  If ready Then\n    Debug.Print \"x\"\n#End If\nEnd Sub\n")
	issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	vb014 := issuesByCode(issues, "VB014")
	if len(vb014) != 1 {
		t.Fatalf("VB014 issues = %+v, want one generic recovery issue", vb014)
	}
	issue := vb014[0]
	if issue.BlockKind != "" || issue.ExpectedCloser != "" || issue.Message != parserRecoveryMessage {
		t.Fatalf("conditional compilation should retain generic recovery guidance: %+v", issue)
	}
}

func TestConditionalIfBalanceAcrossCompilationBranches(t *testing.T) {
	balanced := "Sub Main()\n#If Win64 Then\n  If a Then\n#Else\n  If b Then\n#End If\n    Debug.Print \"x\"\n  Else\n    Debug.Print \"y\"\n  End If\nEnd Sub\n"
	if conditionalIfBalanceInvalid(balanced) {
		t.Fatal("equivalent conditional If headers should merge into one balanced shared block")
	}

	imbalanced := "Sub Main()\n#If Win64 Then\n  If a Then\n#End If\n    Debug.Print \"x\"\nEnd Sub\n"
	if !conditionalIfBalanceInvalid(imbalanced) {
		t.Fatal("an implicit empty compilation branch must expose the unmatched If")
	}
}

func TestLinterAcceptsConditionalCompilationSplitIfWithFlatCST(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	source := []byte("Option Explicit\nSub Main()\n#If Win64 Then\n  If a Then\n#Else\n  If b Then\n#End If\n    Debug.Print \"x\"\n  Else\n    Debug.Print \"y\"\n  End If\nEnd Sub\n")

	issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 0 {
		t.Fatalf("balanced conditional split If should not trigger VB014 with the flat CST: %+v", got)
	}
}

func TestLinterFlagsUnbalancedConditionalCompilationSplitIfWithFlatCST(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	source := []byte("Option Explicit\nSub Main()\n#If Win64 Then\n  If a Then\n#End If\n    Debug.Print \"x\"\nEnd Sub\n")

	issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB014"); len(got) == 0 {
		t.Fatal("unbalanced conditional split If should still trigger VB014 with the flat CST")
	}
}

func TestLinterAcceptsContinuedIfWithConditionalCompilation(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	source := []byte("Option Explicit\nSub Main()\n#If Win64 Then\n  Debug.Print \"64\"\n#Else\n  Debug.Print \"32\"\n#End If\n  If ready _\n    And enabled Then\n    Debug.Print \"x\"\n  End If\nEnd Sub\n")

	issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB014"); len(got) != 0 {
		t.Fatalf("continued If with unrelated conditional compilation should not trigger VB014: %+v", got)
	}
}

func TestLinterRejectsInvalidFlatIfBranchOwnership(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	tests := map[string]string{
		"orphan Else":   "Option Explicit\nSub Main()\n  Else\n    Debug.Print \"x\"\nEnd Sub\n",
		"orphan ElseIf": "Option Explicit\nSub Main()\n  ElseIf ready Then\n    Debug.Print \"x\"\nEnd Sub\n",
		"orphan Else with conditional compilation": "Option Explicit\nSub Main()\n#If Win64 Then\n  Debug.Print \"64\"\n#Else\n  Debug.Print \"32\"\n#End If\n  Else\n    Debug.Print \"x\"\nEnd Sub\n",
		"duplicate Else":    "Option Explicit\nSub Main()\n  If ready Then\n    Debug.Print \"x\"\n  Else\n    Debug.Print \"y\"\n  Else\n    Debug.Print \"z\"\n  End If\nEnd Sub\n",
		"ElseIf after Else": "Option Explicit\nSub Main()\n  If ready Then\n    Debug.Print \"x\"\n  Else\n    Debug.Print \"y\"\n  ElseIf fallback Then\n    Debug.Print \"z\"\n  End If\nEnd Sub\n",
	}

	for name, source := range tests {
		t.Run(name, func(t *testing.T) {
			issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, []byte(source))
			if err != nil {
				t.Fatal(err)
			}
			if got := issuesByCode(issues, "VB014"); len(got) == 0 {
				t.Fatalf("invalid flat If branch ownership should trigger VB014: %s", source)
			}
		})
	}
}

func TestLinterFallsBackToGenericRecoveryForAmbiguousBlocks(t *testing.T) {
	path := filepath.Join(t.TempDir(), "src", "modules", "Main.bas")
	tests := []struct {
		name   string
		source string
	}{
		{
			name:   "malformed For opener",
			source: "Option Explicit\nSub Main()\n  For item\n    Debug.Print item\nEnd Sub\n",
		},
		{
			name:   "malformed Next counter",
			source: "Option Explicit\nSub Main()\n  For item = 1 To 2\n    Debug.Print item\n  Next item,\nEnd Sub\n",
		},
		{
			name:   "mismatched closer",
			source: "Option Explicit\nSub Main()\n  If ready Then\n    Debug.Print \"x\"\n  Next\nEnd Sub\n",
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			issues, err := (Linter{RootDir: filepath.Dir(filepath.Dir(filepath.Dir(path))), Config: config.Default()}).LintSource(path, []byte(tc.source))
			if err != nil {
				t.Fatal(err)
			}
			vb014 := issuesByCode(issues, "VB014")
			if len(vb014) != 1 {
				t.Fatalf("VB014 issues = %+v, want one generic recovery issue", vb014)
			}
			issue := vb014[0]
			if issue.BlockKind != "" || issue.ExpectedCloser != "" || issue.OpeningLine != 0 || issue.Message != parserRecoveryMessage {
				t.Fatalf("ambiguous source should retain generic recovery guidance: %+v", issue)
			}
		})
	}
}

func TestParserRecoveryDetailsForBlocksAssociateOnlyMatchingRecovery(t *testing.T) {
	blocks := []unclosedBlockCandidate{
		{openingLine: 2, expectedLine: 9},
		{openingLine: 4, expectedLine: 6},
		{openingLine: 11, expectedLine: 14},
	}
	details := []parserRecoveryDetail{
		{range_: vbaast.Range{StartLine: 5}, node: "ERROR", token: "inner", context: "inner context"},
		{range_: vbaast.Range{StartLine: 12}, node: "MISSING", token: "later", context: "later context"},
	}
	associated := parserRecoveryDetailsForBlocks(details, blocks)
	if associated[0].node != "" {
		t.Fatalf("outer block should not inherit nested recovery detail: %+v", associated)
	}
	if associated[1].token != "inner" || associated[2].token != "later" {
		t.Fatalf("associated recovery details = %+v", associated)
	}
}

func TestParserRecoveryDetailBoundsUnicodeText(t *testing.T) {
	text := strings.Repeat("あ", maxParserRecoveryContextRunes+1)
	got := truncateParserRecoveryText(text, maxParserRecoveryContextRunes)
	if utf8.RuneCountInString(got) != maxParserRecoveryContextRunes || !strings.HasSuffix(got, "…") {
		t.Fatalf("unexpected bounded text %q", got)
	}
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
			name: "procedure call with Unicode name",
			body: "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n" + continuedLogicalLine(
				"    Call 実行(a0,", "        arg,", "        finalArg)\nEnd Sub\n", 25),
			line:          4,
			kind:          "procedure_call",
			symbol:        "実行",
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

func TestLinterVB021UsesRootedReachabilityAndClusters(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Private Sub Used()
End Sub

Private Sub UnusedRoot()
  UnusedA
End Sub

Private Sub UnusedA()
  UnusedB
End Sub

Private Sub UnusedB()
End Sub

Private Sub Foo_Bar()
End Sub

Private Sub Isolated()
End Sub

Public Sub Run()
  Used
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectUnusedPrivateProcedures = true

	issues, err := (Linter{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB021")
	if len(got) != 5 {
		t.Fatalf("VB021 issues = %d, want five unreachable private procedures: %+v", len(got), got)
	}
	for _, name := range []string{"UnusedRoot", "UnusedA", "UnusedB", "Foo_Bar", "Isolated"} {
		found := false
		for _, issue := range got {
			if issue.Symbol == name {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("missing VB021 for %s: %+v", name, got)
		}
	}
	for _, issue := range got {
		if issue.Symbol == "Used" {
			t.Fatalf("called private procedure was reported: %+v", issue)
		}
	}
	clusterContext := ""
	for _, issue := range got {
		if strings.Contains(issue.Context, "Unreachable private call cluster:") {
			clusterContext = issue.Context
			break
		}
	}
	for _, name := range []string{"Main.UnusedRoot", "Main.UnusedA", "Main.UnusedB"} {
		if !strings.Contains(clusterContext, name) {
			t.Fatalf("cluster context %q does not include %s", clusterContext, name)
		}
	}
	encoded, err := json.Marshal(got)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"context":"Unreachable private call cluster:`) {
		t.Fatalf("VB021 cluster context missing from JSON: %s", encoded)
	}
}

func TestLinterVB021KeepsInlineSuppressionLineBased(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
' xlflow:disable-next-line VB021
Private Sub Suppressed()
End Sub

Private Sub Reported()
End Sub

Public Sub Run()
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectUnusedPrivateProcedures = true

	result, err := (Linter{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(result.Issues, "VB021")
	if len(got) != 1 || got[0].Symbol != "Reported" {
		t.Fatalf("VB021 inline suppression = %+v", got)
	}
	if got[0].Line != 6 {
		t.Fatalf("VB021 declaration line = %d, want 6", got[0].Line)
	}
}

func TestLinterVB021RecognizesTestProceduresAsRoots(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Tests.bas", `Option Explicit
Public Sub TestWorkflow(ByVal input As Long)
  TestHelper
End Sub

Private Sub TestHelper()
End Sub

Private Sub TestOrphan()
End Sub
`)
	cfg := config.Default()
	cfg.Project.Entry = "Missing.Run"
	cfg.Lint.DetectUnusedPrivateProcedures = true

	issues, err := (Linter{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB021")
	if len(got) != 1 || got[0].Symbol != "TestOrphan" {
		t.Fatalf("test root reachability = %+v", got)
	}
}

func TestLinterVB021HandlesDynamicReachabilityConservatively(t *testing.T) {
	t.Run("known target from reachable caller", func(t *testing.T) {
		dir := t.TempDir()
		writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Application.Run "Main.DynamicTarget"
End Sub

Private Sub DynamicTarget()
End Sub

Private Sub Orphan()
  Application.Run "Main.OtherTarget"
End Sub

Private Sub OtherTarget()
End Sub
`)
		cfg := config.Default()
		cfg.Lint.DetectUnusedPrivateProcedures = true
		issues, err := (Linter{RootDir: dir, Config: cfg}).Run()
		if err != nil {
			t.Fatal(err)
		}
		got := issuesByCode(issues, "VB021")
		if len(got) != 2 {
			t.Fatalf("VB021 issues = %d, want unreachable caller and target: %+v", len(got), got)
		}
		for _, issue := range got {
			if issue.Symbol == "DynamicTarget" {
				t.Fatalf("known dynamic target should be possibly reachable: %+v", got)
			}
		}
	})

	t.Run("unknown target from reachable caller", func(t *testing.T) {
		dir := t.TempDir()
		writeLintModule(t, dir, "Main.bas", `Option Explicit
Private callbackName As String

Public Sub Run()
  Application.OnKey "{F1}", callbackName
End Sub

Private Sub CandidateA()
End Sub

Private Sub CandidateB()
End Sub
`)
		cfg := config.Default()
		cfg.Lint.DetectUnusedPrivateProcedures = true
		issues, err := (Linter{RootDir: dir, Config: cfg}).Run()
		if err != nil {
			t.Fatal(err)
		}
		if got := issuesByCode(issues, "VB021"); len(got) != 0 {
			t.Fatalf("unknown dynamic call from reachable root should suppress VB021: %+v", got)
		}
	})
}

func TestLinterVB021TreatsPublicStandardModuleAPIsAsPossibleRoots(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Api.bas", `Option Explicit
Public Function PublicApi(ByVal value As String) As String
  PublicApi = PrivateHelper(value)
End Function

Private Function PrivateHelper(ByVal value As String) As String
  PrivateHelper = value
End Function

Private Sub Orphan()
End Sub
`)
	cfg := config.Default()
	cfg.Project.Entry = "Missing.Run"
	cfg.Lint.DetectUnusedPrivateProcedures = true
	issues, err := (Linter{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB021")
	if len(got) != 1 || got[0].Symbol != "Orphan" {
		t.Fatalf("public standard-module API reachability = %+v", got)
	}
}

func TestLinterVB021RecognizesEventsAndWithEventsHandlers(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	cfg.Project.Entry = "Missing.Run"
	cfg.Lint.DetectUnusedPrivateProcedures = true
	files := map[string]string{
		"src/workbook/ThisWorkbook.cls": `VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Private Sub Workbook_Open()
  WorkbookHelper
End Sub
Private Sub WorkbookHelper()
End Sub
Private Sub Workbook_Helper()
End Sub
Private Sub DocOrphan()
End Sub
`,
		"src/workbook/Sheet1.cls": `VERSION 1.0 CLASS
Attribute VB_Name = "Sheet1"
Private Sub Worksheet_Change(ByVal Target As Range)
  WorksheetHelper
End Sub
Private Sub WorksheetHelper()
End Sub
Private Sub Worksheet_Helper()
End Sub
Private Sub SheetOrphan()
End Sub
`,
		"src/classes/Watcher.cls": `VERSION 1.0 CLASS
Attribute VB_Name = "Watcher"
Public WithEvents App As Excel.Application
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  EventHelper
End Sub
Private Sub EventHelper()
End Sub
Private Sub EventOrphan()
End Sub
`,
		"src/forms/UserForm1.frm": `VERSION 5.00
Begin VB.UserForm UserForm1
   Begin VB.CommandButton cmdOK
   End
End
Attribute VB_Name = "UserForm1"
`,
		"src/forms/code/UserForm1.bas": `Attribute VB_Name = "UserForm1"
Private Sub UserForm_Initialize()
  FormHelper
End Sub
Private Sub cmdOK_Click()
  FormClickHelper
End Sub
Private Sub FormHelper()
End Sub
Private Sub FormClickHelper()
End Sub
Private Sub FormOrphan()
End Sub
`,
		"src/modules/Main.bas": `Attribute VB_Name = "Main"
Private Sub Auto_Open()
  AutoHelper
End Sub
Private Sub AutoHelper()
End Sub
Private Sub AutoOrphan()
End Sub
Private Sub Foo_Bar()
End Sub
`,
	}
	for name, body := range files {
		path := filepath.Join(dir, filepath.FromSlash(name))
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
			t.Fatal(err)
		}
	}

	issues, err := (Linter{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, "VB021")
	for _, name := range []string{"DocOrphan", "SheetOrphan", "Workbook_Helper", "Worksheet_Helper", "EventOrphan", "FormOrphan", "AutoOrphan", "Foo_Bar"} {
		found := false
		for _, issue := range got {
			if issue.Symbol == name {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("missing event/root VB021 for %s: %+v", name, got)
		}
	}
	for _, name := range []string{"WorkbookHelper", "WorksheetHelper", "EventHelper", "FormHelper", "FormClickHelper", "AutoHelper", "App_SheetChange", "Workbook_Open", "Worksheet_Change", "UserForm_Initialize", "cmdOK_Click", "Auto_Open"} {
		for _, issue := range got {
			if issue.Symbol == name {
				t.Fatalf("rooted event procedure was reported: %s (%+v)", name, got)
			}
		}
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

func TestLinterUnusedLocalVariableCountsEarlierConstOnSameLine(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Const firstValue As Long = 1: Const secondValue As Long = firstValue
  Debug.Print secondValue
End Sub
`)
	cfg := config.Default()
	cfg.Lint.DetectUnusedLocalVariables = true

	issues, err := Linter{RootDir: dir, Config: cfg}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB020"); len(got) != 0 {
		t.Fatalf("same-line constant references should count as reads: %+v", got)
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

func TestLinterLintParsedMatchesLintSource(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	source := []byte("Option Explicit\nPublic Sub Run()\n    Dim value\n    Range(\"A1\").Select\nEnd Sub\n")
	linter := Linter{RootDir: dir, Config: config.Default()}
	want, err := linter.LintSource(path, source)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	got, err := linter.LintParsed(doc)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("LintParsed issues = %+v, want %+v", got, want)
	}
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

func TestLinterDoesNotTreatMultilineComparisonArgumentsAsAssignments(t *testing.T) {
	dir := t.TempDir()
	writeLintModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Private Function EqualsText(ByVal leftValue As String, ByVal rightValue As String) As Boolean
  EqualsText = (StrComp( _
    leftValue, _
    rightValue, _
    vbTextCompare) = 0)
End Function
`)
	issues, err := Linter{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := issuesByCode(issues, "VB029"); len(got) != 0 {
		t.Fatalf("multiline comparison argument should not trigger VB029: %+v", got)
	}
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

func TestLinterProcedureNameConstantRule(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Lint.ProcedureNameConstant = config.ProcedureNameConstantConfig{Enabled: true, ConstantName: "procedure_name"}
	source := `Option Explicit
Const PROCEDURE_NAME As String = "ModuleLevel"

Public Sub Matching()
    Const PROCEDURE_NAME As String = "Matching"
End Sub

Public Function CaseOnly() As String
    Const PROCEDURE_NAME As String = "caseonly"
End Function

Public Sub Multiple()
    Const PROCEDURE_NAME As String = "OldMultiple", OTHER_NAME As String = "ignored"
End Sub

Public Sub DynamicValue()
    Const PROCEDURE_NAME As String = "Dynamic" & "Value"
End Sub

Public Property Get Caption() As String
    Const PROCEDURE_NAME As String = "OldGet"
End Property

Public Property Let Caption(ByVal value As String)
    Const PROCEDURE_NAME As String = "OldLet"
End Property

Public Property Set Caption(ByVal value As Object)
    Const PROCEDURE_NAME As String = "OldSet"
End Property
`
	issues, err := (Linter{RootDir: root, Config: cfg}).LintSource(filepath.Join(root, "Main.bas"), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, procedureNameConstantRuleID)
	if len(got) != 5 {
		t.Fatalf("VB044 issues = %d, want 5: %+v", len(got), got)
	}
	want := map[string]string{
		"caseonly":    "CaseOnly",
		"OldMultiple": "Multiple",
		"OldGet":      "Caption",
		"OldLet":      "Caption",
		"OldSet":      "Caption",
	}
	for _, issue := range got {
		matched := false
		for current, expected := range want {
			if strings.Contains(issue.Message, current) && strings.Contains(issue.Message, expected) {
				matched = issue.Kind == "procedure_name_constant" && issue.Symbol == "PROCEDURE_NAME" && issue.Column > 0
				break
			}
		}
		if !matched {
			t.Fatalf("unexpected VB044 issue: %+v", issue)
		}
	}
}

func TestLinterProcedureNameConstantRuleScansModuleKinds(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.UserForm.CodeSource = "frm"
	cfg.Lint.ProcedureNameConstant = config.ProcedureNameConstantConfig{Enabled: true, ConstantName: "PROCEDURE_NAME"}
	files := map[string]string{
		"src/modules/Main.bas": `Attribute VB_Name = "Main"
Option Explicit
Public Sub StandardProcedure()
    Const PROCEDURE_NAME As String = "OldStandard"
End Sub
`,
		"src/classes/Worker.cls": `Attribute VB_Name = "Worker"
Option Explicit
Public Sub ClassProcedure()
    Const PROCEDURE_NAME As String = "OldClass"
End Sub
`,
		"src/workbook/ThisWorkbook.cls": `Attribute VB_Name = "ThisWorkbook"
Option Explicit
Private Sub Workbook_Open()
    Const PROCEDURE_NAME As String = "OldWorkbook"
End Sub
`,
		"src/forms/Dialog.frm": `VERSION 5.00
Begin VB.UserForm Dialog
End
Attribute VB_Name = "Dialog"
Option Explicit
Private Sub UserForm_Initialize()
    Const PROCEDURE_NAME As String = "OldForm"
End Sub
`,
	}
	for name, body := range files {
		path := filepath.Join(root, filepath.FromSlash(name))
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	issues, err := (Linter{RootDir: root, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := issuesByCode(issues, procedureNameConstantRuleID)
	if len(got) != 4 {
		t.Fatalf("VB044 issues = %d, want 4: %+v", len(got), got)
	}
	wantFiles := map[string]bool{
		"src/modules/Main.bas":          true,
		"src/classes/Worker.cls":        true,
		"src/workbook/ThisWorkbook.cls": true,
		"src/forms/Dialog.frm":          true,
	}
	for _, issue := range got {
		if !wantFiles[issue.File] {
			t.Fatalf("unexpected VB044 file: %+v", issue)
		}
	}
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
