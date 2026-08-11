package analyze

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA241DetectsLoopFormsAndClassifiesBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(items As Variant)
    Dim values() As Long
    Dim fixed(1 To 3) As Long
    Dim i As Long
    Dim item As Variant
    For i = 1 To 3
        ReDim Preserve values(1 To 10)
    Next i
    For Each item In items
        ReDim Preserve values(1 To Grow(item))
    Next item
    While i < 4
        ReDim Preserve values(1 To 10)
        i = i + 1
    Wend
    Do
        Do While i < 5
            ReDim Preserve values(1 To i)
            i = i + 1
        Loop
        Exit Do
    Loop
    Do Until i > 6
        ReDim Preserve values(1 To i)
        i = i + 1
    Loop
    Do
        ReDim Preserve values(1 To i)
        i = i + 1
    Loop While i < 7
    Do
        ReDim Preserve values(1 To i)
        i = i + 1
    Loop Until i > 5
End Sub

Private Function Grow(value As Variant) As Long
    Grow = value + 1
End Function
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA241")
	expected := map[int]struct {
		severity string
		markers  []string
	}{
		10: {severity: "information", markers: []string{"loop-invariant"}},
		13: {severity: "warning", markers: []string{"loop-variable-dependent"}},
		16: {severity: "information", markers: []string{"loop-invariant"}},
		21: {severity: "warning", markers: []string{"loop-variable-dependent", "Nested loop depth"}},
		27: {severity: "warning", markers: []string{"loop-variable-dependent"}},
		31: {severity: "warning", markers: []string{"loop-variable-dependent"}},
		35: {severity: "warning", markers: []string{"loop-variable-dependent"}},
	}
	if len(got) != len(expected) {
		t.Fatalf("VBA241 findings = %+v, want one finding at each expected source line", got)
	}
	byLine := make(map[int]Finding, len(got))
	for _, finding := range got {
		if _, duplicate := byLine[finding.Line]; duplicate {
			t.Fatalf("duplicate VBA241 finding at source line %d: %+v", finding.Line, got)
		}
		byLine[finding.Line] = finding
	}
	for line, want := range expected {
		finding, ok := byLine[line]
		if !ok {
			t.Fatalf("missing VBA241 finding at source line %d: %+v", line, got)
		}
		if finding.Severity != want.severity {
			t.Fatalf("VBA241 finding at line %d = %+v, want severity %q", line, finding, want.severity)
		}
		for _, marker := range want.markers {
			if !strings.Contains(finding.Message, marker) {
				t.Fatalf("VBA241 finding at line %d = %+v, want %q classification", line, finding, marker)
			}
		}
	}
}

func TestVBA241KeepsUnchangedDoConditionInvariant(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim values() As Long
    Dim limit As Long
    Dim ticks As Long
    limit = 3
    Do While limit > ticks
        ReDim Preserve values(1 To limit)
        ticks = ticks + 1
    Loop
    Do
        ReDim Preserve values(1 To limit)
        ticks = ticks + 1
    Loop Until limit <= ticks
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA241")
	expectedLines := map[int]bool{10: true, 14: true}
	if len(got) != len(expectedLines) {
		t.Fatalf("unchanged Do condition findings = %+v, want one finding at each expected source line", got)
	}
	for _, finding := range got {
		if !expectedLines[finding.Line] || finding.Severity != "information" || !strings.Contains(finding.Message, "loop-invariant") {
			t.Fatalf("unchanged Do condition finding = %+v, want information loop-invariant classification at lines 10 and 14", finding)
		}
		delete(expectedLines, finding.Line)
	}
	if len(expectedLines) != 0 {
		t.Fatalf("missing unchanged Do condition findings at lines: %+v", expectedLines)
	}
}

func TestVBA241ExcludesPreallocatedAndFixedArrays(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim values() As Long
    Dim fixed(1 To 3) As Long
    Dim i As Long
    ReDim values(1 To 10)
    For i = 1 To 3
        values(i) = i
    Next i
    For i = 1 To 3
        ReDim Preserve fixed(1 To i)
    Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA241"); len(got) != 0 {
		t.Fatalf("preallocated/fixed arrays produced VBA241: %+v", got)
	}
}

func TestVBA241DetectsVariantArrayTarget(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim values As Variant
    Dim i As Long
    For i = 1 To 3
        ReDim Preserve values(1 To i)
    Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA241"); len(got) != 1 {
		t.Fatalf("Variant array target findings = %+v, want one", got)
	}
}

func TestVBA241AggregatesMultipleTargetsAndMultidimensionalBounds(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim first() As Long
    Dim second() As Long
    Dim i As Long
    For i = 1 To 3
        ReDim Preserve first(1 To 2, 1 To i), second(1 To 10)
    Next i
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA241")
	if len(got) != 1 {
		t.Fatalf("multiple target findings = %+v, want one aggregate", got)
	}
	if got[0].Severity != "warning" || !strings.Contains(got[0].Message, "first") || !strings.Contains(got[0].Message, "second") {
		t.Fatalf("unexpected aggregate finding: %+v", got[0])
	}
}

func TestVBA241CanBeDisabled(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim values() As Long
    Dim i As Long
    For i = 1 To 3
        ReDim Preserve values(1 To i)
    Next i
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectRedimPreserveInLoops = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA241"); len(got) != 0 {
		t.Fatalf("disabled VBA241 produced findings: %+v", got)
	}
}

func TestVBA241RealtimeMatchesBatchClassification(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := []byte(`Option Explicit
Public Sub Run()
    Dim values() As Long
    Dim i As Long
    For i = 1 To 100
        ReDim Preserve values(1 To i)
    Next i
End Sub
`)
	writeModule(t, dir, "Main.bas", string(source))
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	findings, err := SourceRealtimeFindings(dir, path, config.Default(), source)
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA241")
	if len(got) != 1 || got[0].Severity != "warning" || !strings.Contains(got[0].Message, "loop-variable-dependent") {
		t.Fatalf("realtime VBA241 findings = %+v, want one growth warning", got)
	}
}
