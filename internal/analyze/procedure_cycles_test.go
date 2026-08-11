package analyze

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA244ReportsIndependentCyclesWithClosedDeterministicPaths(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Alpha()
  Beta
End Sub
Public Sub Beta()
  Alpha
End Sub
Public Sub Gamma()
  Delta
End Sub
Public Sub Delta()
  Gamma
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA244")
	if len(got) != 2 {
		t.Fatalf("VBA244 findings = %+v, want two independent cycles", got)
	}
	for _, finding := range got {
		if finding.Severity != "information" || finding.CallCycle == nil || len(finding.CallCycle.Path) != 3 {
			t.Fatalf("ordinary cycle finding = %+v", finding)
		}
		if finding.CallCycle.Path[0].QualifiedName != finding.CallCycle.Path[2].QualifiedName {
			t.Fatalf("cycle path is not closed: %+v", finding.CallCycle.Path)
		}
		if len(finding.CallCycle.Edges) != 2 || finding.CallCycle.CrossModule {
			t.Fatalf("cycle edges/context = %+v", finding.CallCycle)
		}
	}
	again, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	firstCycles, secondCycles := findingsByCode(got, "VBA244"), findingsByCode(again, "VBA244")
	if len(firstCycles) != len(secondCycles) {
		t.Fatalf("repeat cycle count = %d/%d", len(firstCycles), len(secondCycles))
	}
	for i := range firstCycles {
		if firstCycles[i].Message != secondCycles[i].Message || firstCycles[i].File != secondCycles[i].File || firstCycles[i].Line != secondCycles[i].Line {
			t.Fatalf("cycle output changed: first=%+v second=%+v", firstCycles[i], secondCycles[i])
		}
	}
}

func TestVBA244IdentifiesCrossModuleAndEventCycles(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Public Sub Alpha()
  Beta
End Sub
`)
	writeModule(t, dir, "Helpers.bas", `Public Sub Beta()
  Alpha
End Sub
`)
	workbook := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbook, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbook, "Sheet1.cls"), []byte(`Private Sub Worksheet_Change(ByVal Target As Range)
  Worksheet_Change Target
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA244")
	if len(got) != 2 {
		t.Fatalf("VBA244 findings = %+v, want cross-module and event cycles", got)
	}
	var cross, event bool
	for _, finding := range got {
		if finding.CallCycle == nil {
			t.Fatalf("missing cycle context: %+v", finding)
		}
		cross = cross || finding.CallCycle.CrossModule
		if len(finding.CallCycle.EventHandlers) > 0 {
			event = true
			if finding.Severity != "warning" {
				t.Fatalf("event cycle severity = %s, want warning", finding.Severity)
			}
		}
	}
	if !cross || !event {
		t.Fatalf("cycle classifications cross=%v event=%v findings=%+v", cross, event, got)
	}
}

func TestVBA244ElevatesDangerousEffectsAndRetainsUncertainty(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Entry()
  Worker
End Sub
Public Sub Worker()
  On Error Resume Next
  Application.ScreenUpdating = False
  Open "output.txt" For Output As #1
  MissingDispatch
  Leaf
  Entry
End Sub
Public Sub Leaf()
  Application.DisplayAlerts = False
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA244")
	if len(got) != 1 {
		t.Fatalf("VBA244 findings = %+v, want one cycle", got)
	}
	finding := got[0]
	if finding.Severity != "warning" || finding.CallCycle == nil {
		t.Fatalf("dangerous cycle finding = %+v", finding)
	}
	if len(finding.CallCycle.DangerousEffects) < 3 {
		t.Fatalf("dangerous effects = %+v", finding.CallCycle.DangerousEffects)
	}
	var propagated bool
	for _, effect := range finding.CallCycle.DangerousEffects {
		if effect.Origin == "Main.Leaf" {
			propagated = true
		}
	}
	if !propagated {
		t.Fatalf("propagated dangerous effect was not retained: %+v", finding.CallCycle.DangerousEffects)
	}
	if len(finding.CallCycle.Uncertainty) == 0 {
		t.Fatalf("uncertainty was not retained: %+v", finding.CallCycle)
	}
	if !strings.Contains(finding.Message, "Main.Entry") || !strings.Contains(finding.Message, "Main.Worker") {
		t.Fatalf("complete cycle path missing: %s", finding.Message)
	}
}

func TestVBA244UncertaintyAloneDoesNotElevate(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Public Sub A()
  MissingDispatch
  B
End Sub
Public Sub B()
  A
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA244")
	if len(got) != 1 || got[0].Severity != "information" || got[0].CallCycle == nil || len(got[0].CallCycle.Uncertainty) == 0 {
		t.Fatalf("uncertain ordinary cycle = %+v, want information with uncertainty", got)
	}
}

func TestVBA244CanBeDisabled(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", "Public Sub A()\n  A\nEnd Sub\n")
	cfg := config.Default()
	cfg.Analyze.DisabledRules = []string{"VBA244"}
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	loaded, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	findings, err := (Analyzer{RootDir: dir, Config: loaded}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA244"); len(got) != 0 {
		t.Fatalf("disabled VBA244 findings = %+v", got)
	}
}

func TestVBA244HonorsInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Public Sub A()
  ' xlflow:disable-next-line VBA244
  A
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA244"); len(got) != 0 {
		t.Fatalf("inline-suppressed VBA244 findings = %+v", got)
	}
}
