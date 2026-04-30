package gui

import (
	"os"
	"path/filepath"
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
	for _, boundary := range boundaries {
		found[boundary.Symbol] = boundary.Kind
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
}

func TestStripCommentKeepsApostropheInsideStrings(t *testing.T) {
	got := StripComment(`MsgBox "it''s ""done""" ' trailing`)
	if got != `MsgBox "it''s ""done""" ` {
		t.Fatalf("StripComment = %q", got)
	}
}
