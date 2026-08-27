package gui

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func writeHeadlessFixture(t *testing.T, files map[string]string) string {
	t.Helper()
	dir := t.TempDir()
	for name, body := range files {
		path := filepath.Join(dir, "src", "modules", name)
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	return dir
}

func TestAnalyzerRunHeadlessFiltersUnreachableGUIBoundaries(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\n  Debug.Print \"safe\"\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 0 {
		t.Fatalf("expected unreachable GUI boundary to be filtered, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessRetainsReachableGUIBoundaries(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\n  LegacyGui.PickFile\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 1 || boundaries[0].Symbol != "Application.GetOpenFilename" {
		t.Fatalf("expected reachable GUI boundary, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessDoesNotTreatExternalCallsAsProjectUncertainty(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\n  Debug.Print \"safe\"\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 0 {
		t.Fatalf("external Debug.Print should not force project-wide fallback, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessIgnoresUnknownDynamicCallsFromUnreachableProcedures(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\n  Debug.Print \"safe\"\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n\nPrivate Sub Dead()\n  Application.Run macroName\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 0 {
		t.Fatalf("unknown dynamic call in an unreachable procedure should not force fallback, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessTracksStaticDynamicTargets(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\n  Application.Run \"LegacyGui.PickFile\"\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 1 || boundaries[0].Symbol != "Application.GetOpenFilename" {
		t.Fatalf("expected static dynamic target to be retained, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessFallsBackForUnknownDynamicTarget(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\n  Application.Run macroName\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 1 || boundaries[0].Symbol != "Application.GetOpenFilename" {
		t.Fatalf("expected unknown dynamic target to preserve project-wide safety, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessRetainsObjectModuleBoundariesAsPossibleRoots(t *testing.T) {
	dir := t.TempDir()
	classes := filepath.Join(dir, "src", "classes")
	modules := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(classes, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(modules, "SafeEntry.bas"), []byte("Option Explicit\nPublic Sub Run()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classes, "LegacyGui.cls"), []byte("VERSION 1.0 CLASS\nAttribute VB_Name = \"LegacyGui\"\nOption Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 1 || boundaries[0].Symbol != "Application.GetOpenFilename" {
		t.Fatalf("object-module boundaries must remain conservatively reachable, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessRetainsAutoMacroBoundariesAsPossibleRoots(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas":  "Option Explicit\nPublic Sub Run()\nEnd Sub\n",
		"LegacyAuto.bas": "Option Explicit\nPublic Sub Auto_Open()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 1 || boundaries[0].Symbol != "Application.GetOpenFilename" {
		t.Fatalf("automatic macro boundaries must remain conservatively reachable, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessRetainsImplicitPropertyBoundariesAsPossibleRoots(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\n  Debug.Print \"safe\"\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPrivate Property Get GuiValue() As String\n  GuiValue = InputBox(\"Name?\")\nEnd Property\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("SafeEntry.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 1 || boundaries[0].Symbol != "InputBox" {
		t.Fatalf("implicit property boundaries must remain conservatively reachable, got %+v", boundaries)
	}
}

func TestAnalyzerRunHeadlessRetainsBoundariesForUnknownTarget(t *testing.T) {
	dir := writeHeadlessFixture(t, map[string]string{
		"SafeEntry.bas": "Option Explicit\nPublic Sub Run()\nEnd Sub\n",
		"LegacyGui.bas": "Option Explicit\nPublic Sub PickFile()\n  Application.GetOpenFilename \"CSV files (*.csv),*.csv\"\nEnd Sub\n",
	})

	boundaries, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunHeadless("Missing.Run")
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 1 {
		t.Fatalf("unknown target must retain project-wide GUI boundaries, got %+v", boundaries)
	}
}

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
	if !strings.Contains(foundMessages["Application.GetOpenFilename"], "XlflowUI") || !strings.Contains(foundSuggestions["Application.GetOpenFilename"], "XlflowUI.GetOpenFilename") {
		t.Fatalf("expected GetOpenFilename boundary to recommend XlflowUI, got message=%q suggestion=%q", foundMessages["Application.GetOpenFilename"], foundSuggestions["Application.GetOpenFilename"])
	}
	if !strings.Contains(foundMessages["Application.GetSaveAsFilename"], "XlflowUI") || !strings.Contains(foundSuggestions["Application.GetSaveAsFilename"], "XlflowUI.GetSaveAsFilename") {
		t.Fatalf("expected GetSaveAsFilename boundary to recommend XlflowUI, got message=%q suggestion=%q", foundMessages["Application.GetSaveAsFilename"], foundSuggestions["Application.GetSaveAsFilename"])
	}
	if !strings.Contains(foundMessages["Application.FileDialog"], "XlflowUI") || !strings.Contains(foundSuggestions["Application.FileDialog"], "XlflowUI.FileDialogOpen") {
		t.Fatalf("expected FileDialog boundary to recommend XlflowUI, got message=%q suggestion=%q", foundMessages["Application.FileDialog"], foundSuggestions["Application.FileDialog"])
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

func TestAnalyzerIgnoresModifierlessWrapperDeclarations(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Option Explicit
Function MsgBox(ByVal Id As String, ByVal Prompt As String) As VbMsgBoxResult
End Function

Function InputBox(ByVal Id As String, ByVal Prompt As String) As String
End Function
`
	if err := os.WriteFile(filepath.Join(src, "DialogWrappers.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	boundaries, err := Analyzer{RootDir: dir, Config: config.Default()}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(boundaries) != 0 {
		t.Fatalf("expected modifierless wrapper declarations to be ignored, got %+v", boundaries)
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
