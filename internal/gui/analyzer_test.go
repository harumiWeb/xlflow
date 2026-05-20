package gui

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestAnalyzerDetectsGUIBoundaries(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Main()
' MsgBox "commented out"
Debug.Print "MsgBox"
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

	boundaries, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	want := map[string]string{
		"Application.GetOpenFilename":         "file_picker",
		"Application.GetSaveAsFilename":       "file_picker",
		"Application.FileDialog":              "file_picker",
		"InputBox":                            "modal_dialog",
		"MsgBox":                              "modal_dialog",
		"UserForm.Show":                       "user_form",
		"DoEvents":                            "message_pump",
		"Shell":                               "external_process",
		`CreateObject("WScript.Shell").Popup`: "modal_dialog",
	}
	found := map[string]string{}
	foundSuggestions := map[string]string{}
	foundMessages := map[string]string{}
	for _, boundary := range boundaries {
		found[boundary.Symbol] = boundary.Kind
		foundSuggestions[boundary.Symbol] = boundary.Suggestion
		foundMessages[boundary.Symbol] = boundary.Message
		if boundary.File != "src/modules/Main.bas" {
			t.Fatalf("file = %q", boundary.File)
		}
		if boundary.Severity != "interactive-only" {
			t.Fatalf("severity = %q", boundary.Severity)
		}
	}
	if len(found) != len(want) {
		t.Fatalf("found %d boundaries, want %d: %+v", len(found), len(want), boundaries)
	}
	for symbol, kind := range want {
		if found[symbol] != kind {
			t.Fatalf("%s kind = %q, want %q", symbol, found[symbol], kind)
		}
	}
	if !strings.Contains(foundMessages["MsgBox"], "XlflowUI") || !strings.Contains(foundSuggestions["MsgBox"], "XlflowUI.MsgBox") {
		t.Fatalf("expected MsgBox boundary to recommend XlflowUI, got message=%q suggestion=%q", foundMessages["MsgBox"], foundSuggestions["MsgBox"])
	}
	if !strings.Contains(foundMessages["InputBox"], "XlflowUI") || !strings.Contains(foundSuggestions["InputBox"], "XlflowUI.InputBox") {
		t.Fatalf("expected InputBox boundary to recommend XlflowUI, got message=%q suggestion=%q", foundMessages["InputBox"], foundSuggestions["InputBox"])
	}
}

func TestStripCommentKeepsApostropheInsideStrings(t *testing.T) {
	got := StripComment(`MsgBox "it''s ""done""" ' trailing`)
	if got != `MsgBox "it''s ""done""" ` {
		t.Fatalf("StripComment = %q", got)
	}
}

func TestAnalyzerIgnoresXlflowUIWrappers(t *testing.T) {
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

	boundaries, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 0 {
		t.Fatalf("expected wrapper helper to be ignored, got %+v", boundaries)
	}
}

func TestAnalyzerDetectsFullyQualifiedRawDialogsOutsideXlflowUI(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Sub Main()
  Dim decision As VbMsgBoxResult
  decision = VBA.Interaction.MsgBox("Done")
  Debug.Print VBA.Interaction.InputBox("Name")
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	boundaries, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	found := map[string]bool{}
	for _, boundary := range boundaries {
		if boundary.Symbol == "MsgBox" || boundary.Symbol == "InputBox" {
			found[boundary.Symbol] = true
		}
	}
	if !found["MsgBox"] || !found["InputBox"] {
		t.Fatalf("expected fully qualified raw dialogs to be detected, got %+v", boundaries)
	}
}
