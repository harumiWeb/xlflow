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
    For i = 1 To 3
        ReDim Preserve values(1 To 10)
    Next i
    For Each i In items
        ReDim Preserve values(1 To Grow(i))
    Next i
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
    Do
        ReDim Preserve values(1 To i)
        i = i + 1
    Loop While i < 7
    Do
        ReDim Preserve values(1 To i)
        Exit Do
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
	if len(got) != 6 {
		t.Fatalf("VBA241 findings = %+v, want six loop-body findings", got)
	}
	if got[0].Severity != "information" {
		t.Fatalf("constant single-loop finding severity = %q, want information: %+v", got[0].Severity, got[0])
	}
	var dependent, nested bool
	for _, finding := range got {
		if strings.Contains(finding.Message, "loop-variable-dependent") {
			dependent = finding.Severity == "warning"
		}
		if strings.Contains(finding.Message, "Nested loop depth") {
			nested = finding.Severity == "warning"
		}
	}
	if !dependent || !nested {
		t.Fatalf("missing dependent/nested classifications: %+v", got)
	}
	postTestDependent := false
	for _, finding := range got {
		if finding.Line == 26 && finding.Severity == "warning" && strings.Contains(finding.Message, "loop-variable-dependent") {
			postTestDependent = true
		}
	}
	if !postTestDependent {
		t.Fatalf("post-test Do Loop While dependency was not classified: %+v", got)
	}
	preTestDependent := false
	for _, finding := range got {
		if finding.Line == 20 && finding.Severity == "warning" && strings.Contains(finding.Message, "loop-variable-dependent") {
			preTestDependent = true
		}
	}
	if !preTestDependent {
		t.Fatalf("pre-test Do While dependency was not classified: %+v", got)
	}
	postUntilDependent := false
	for _, finding := range got {
		if finding.Line == 30 && finding.Severity == "warning" && strings.Contains(finding.Message, "loop-variable-dependent") {
			postUntilDependent = true
		}
	}
	if !postUntilDependent {
		t.Fatalf("post-test Do Until dependency was not classified: %+v", got)
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
