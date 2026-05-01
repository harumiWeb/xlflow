package lint

import (
	"os"
	"path/filepath"
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
	for _, issue := range issues {
		if issue.Code == "VB007" && issue.Kind != "" && issue.Symbol != "" && issue.Suggestion != "" {
			foundBoundaryMetadata = true
			break
		}
	}
	if !foundBoundaryMetadata {
		t.Fatalf("expected VB007 to include GUI boundary metadata: %+v", issues)
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
