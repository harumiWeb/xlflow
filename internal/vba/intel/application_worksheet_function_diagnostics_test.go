package intel

import (
	"context"
	"errors"
	"fmt"
	"path/filepath"
	"strings"
	"testing"
)

func TestApplicationWorksheetFunctionDispatchDiagnosticsResolveApplicationReceivers(t *testing.T) {
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim result As Variant
    Dim app As Excel.Application
    result = Application.Sum(1)
    result = app.Sum(1)
    With Application
        .Sum 1
    End With
    result = Application.WorksheetFunction.Sum(1)
    result = WorksheetFunction.Sum(1)
    result = app.WorksheetFunction.Sum(1)
    result = Application.Calculate
    result = Sum(1)
End Sub
`,
	}

	diagnostics := analyzer.ApplicationWorksheetFunctionDispatchDiagnostics(doc)
	if len(diagnostics) != 3 {
		t.Fatalf("VBA261 diagnostics = %+v, want three", diagnostics)
	}
	wantLines := map[int]string{4: "Sum", 5: "Sum", 7: "Sum"}
	for _, diagnostic := range diagnostics {
		want, ok := wantLines[diagnostic.Range.Start.Line]
		if !ok {
			t.Fatalf("VBA261 diagnostic on unexpected line: %+v", diagnostic)
		}
		if !strings.Contains(diagnostic.Message, fmt.Sprintf("Application.WorksheetFunction.%s", want)) ||
			!strings.Contains(diagnostic.Message, "Double") ||
			!strings.Contains(diagnostic.Message, "worksheet error values") ||
			!strings.Contains(diagnostic.Message, "runtime error") {
			t.Fatalf("VBA261 diagnostic = %+v, want binding, type, and error-behavior explanation", diagnostic)
		}
		if diagnostic.Code != "VBA261" || diagnostic.Rule != "VBA261" || diagnostic.Severity != "information" || diagnostic.Confidence != "high" {
			t.Fatalf("VBA261 diagnostic metadata = %+v", diagnostic)
		}
		if diagnostic.Range.End.Character-diagnostic.Range.Start.Character != len(want) {
			t.Fatalf("VBA261 range = %+v, want %q member range", diagnostic.Range, want)
		}
		delete(wantLines, diagnostic.Range.Start.Line)
	}
	if len(wantLines) != 0 {
		t.Fatalf("VBA261 missed lines: %+v", wantLines)
	}
}

func TestApplicationWorksheetFunctionDispatchDiagnosticsFailOpenForUnknownReceivers(t *testing.T) {
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim result As Variant
    Dim Application As Object
    Dim app As Object
    Dim lateBound As Variant
    result = Application.Sum(1)
    result = app.Sum(1)
    result = lateBound.Sum(1)
    result = CreateObject("Excel.Application").Sum(1)
End Sub
`,
	}
	if diagnostics := analyzer.ApplicationWorksheetFunctionDispatchDiagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("shadowed/late-bound VBA261 diagnostics = %+v, want none", diagnostics)
	}
}

func TestApplicationWorksheetFunctionDispatchDiagnosticsFailOpenForIncompleteTypeLib(t *testing.T) {
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Application.Sum 1
End Sub
`,
	}
	for _, test := range []struct {
		name       string
		source     string
		confidence string
	}{
		{name: "curated", source: "xlflow", confidence: "curated"},
		{name: "incomplete", source: "typelib", confidence: "incomplete"},
	} {
		t.Run(test.name, func(t *testing.T) {
			analyzer := newWorksheetFunctionAnalyzer(t, test.source, test.confidence)
			if diagnostics := analyzer.ApplicationWorksheetFunctionDispatchDiagnostics(doc); len(diagnostics) != 0 {
				t.Fatalf("%s VBA261 diagnostics = %+v, want none", test.name, diagnostics)
			}
		})
	}
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	analyzer.TypeDBResolutionIncomplete = true
	if diagnostics := analyzer.ApplicationWorksheetFunctionDispatchDiagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("incomplete TypeDB VBA261 diagnostics = %+v, want none", diagnostics)
	}
}

func TestApplicationWorksheetFunctionDispatchDiagnosticsContextHonorsCancellation(t *testing.T) {
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err := analyzer.ApplicationWorksheetFunctionDispatchDiagnosticsContext(ctx, Document{Source: "Application.Sum(1)"})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled VBA261 diagnostics error = %v, want context.Canceled", err)
	}
}
