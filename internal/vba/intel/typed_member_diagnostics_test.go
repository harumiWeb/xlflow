package intel

import (
	"context"
	"errors"
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestUnavailableWorksheetFunctionMemberDiagnosticsResolveTypedReceivers(t *testing.T) {
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim result As Variant
    Dim wf As Excel.WorksheetFunction
    result = wf.Abs(1)
    result = WorksheetFunction.Abs(1)
    result = Application.WorksheetFunction.Concatenate("a", "b")
    With Application.WorksheetFunction
        .Abs(1)
    End With
    result = WorksheetFunction.Abs( _
        1)
    result = WorksheetFunction.Sum(1)
End Sub
`,
	}

	diagnostics := analyzer.UnavailableWorksheetFunctionMemberDiagnostics(doc)
	if len(diagnostics) != 5 {
		t.Fatalf("VBA252 diagnostics = %+v, want five", diagnostics)
	}
	wantLines := map[int]string{
		4:  "Abs",
		5:  "Abs",
		6:  "Concatenate",
		8:  "Abs",
		10: "Abs",
	}
	for _, diagnostic := range diagnostics {
		want, ok := wantLines[diagnostic.Range.Start.Line]
		if !ok {
			t.Fatalf("VBA252 diagnostic on unexpected line: %+v", diagnostic)
		}
		if !strings.Contains(diagnostic.Message, fmt.Sprintf("%q", want)) {
			t.Fatalf("VBA252 diagnostic = %+v, want member %q", diagnostic, want)
		}
		if diagnostic.Code != "VBA252" || diagnostic.Rule != "VBA252" || diagnostic.Severity != "warning" || diagnostic.Confidence != "high" {
			t.Fatalf("VBA252 diagnostic metadata = %+v", diagnostic)
		}
		if diagnostic.Range.End.Character-diagnostic.Range.Start.Character != len(want) {
			t.Fatalf("VBA252 range = %+v, want %q member range", diagnostic.Range, want)
		}
		wantLines[diagnostic.Range.Start.Line] = ""
	}
	for line, member := range wantLines {
		if member != "" {
			t.Fatalf("VBA252 missed line %d member %q: %+v", line, member, diagnostics)
		}
	}
}

func TestUnavailableWorksheetFunctionMemberDiagnosticsCanonicalizesWorksheetFunctionAliases(t *testing.T) {
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Local()
    Dim result As Variant
    Dim wf As WorksheetFunction
    result = wf.Abs(1)
End Sub

Public Sub Parameter(wf As WorksheetFunction)
    Dim result As Variant
    result = wf.Concatenate("a", "b")
End Sub
`,
	}

	diagnostics := analyzer.UnavailableWorksheetFunctionMemberDiagnostics(doc)
	if len(diagnostics) != 2 {
		t.Fatalf("VBA252 diagnostics = %+v, want local and parameter findings", diagnostics)
	}
	want := map[int]string{4: "Abs", 9: "Concatenate"}
	for _, diagnostic := range diagnostics {
		member, ok := want[diagnostic.Range.Start.Line]
		if !ok {
			t.Fatalf("VBA252 diagnostic on unexpected line: %+v", diagnostic)
		}
		if !strings.Contains(diagnostic.Message, fmt.Sprintf("%q", member)) {
			t.Fatalf("VBA252 diagnostic = %+v, want member %q", diagnostic, member)
		}
		delete(want, diagnostic.Range.Start.Line)
	}
	if len(want) != 0 {
		t.Fatalf("VBA252 missed alias findings: %+v", want)
	}
}

func TestUnavailableWorksheetFunctionMemberDiagnosticsFailOpenForShadowedAndUnknownTypes(t *testing.T) {
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim result As Variant
    Dim WorksheetFunction As Object
    result = WorksheetFunction.Abs(1)
End Sub

Public Sub ApplicationShadow()
    Dim result As Variant
    Dim Application As Object
    result = Application.WorksheetFunction.Abs(1)
End Sub

Public Sub VariantShadow()
    Dim result As Variant
    Dim WorksheetFunction As Variant
    result = WorksheetFunction.Abs(1)
End Sub

Public Sub LateBound()
    Dim result As Variant
    Dim lateBound As Object
    result = lateBound.WorksheetFunction.Abs(1)
End Sub
`,
	}
	if diagnostics := analyzer.UnavailableWorksheetFunctionMemberDiagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("shadowed/late-bound VBA252 diagnostics = %+v, want none", diagnostics)
	}
}

func TestUnavailableWorksheetFunctionMemberDiagnosticsFailOpenForCuratedAndIncompleteTypes(t *testing.T) {
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim result As Variant
    result = WorksheetFunction.Abs(1)
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
			if diagnostics := analyzer.UnavailableWorksheetFunctionMemberDiagnostics(doc); len(diagnostics) != 0 {
				t.Fatalf("%s VBA252 diagnostics = %+v, want none", test.name, diagnostics)
			}
		})
	}
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	analyzer.TypeDBResolutionIncomplete = true
	if diagnostics := analyzer.UnavailableWorksheetFunctionMemberDiagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("incomplete TypeDB VBA252 diagnostics = %+v, want none", diagnostics)
	}
}

func TestTypedMemberDiagnosticsContextHonorsCancellation(t *testing.T) {
	analyzer := newGeneratedWorksheetFunctionAnalyzer(t)
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err := analyzer.TypedMemberDiagnosticsContext(ctx, Document{Source: "WorksheetFunction.Abs(1)"}, TypedMemberDiagnosticSpec{
		Receiver: "Excel.WorksheetFunction",
	})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled typed member diagnostics error = %v, want context.Canceled", err)
	}
}

func TestPreferUnavailableWorksheetFunctionDiagnosticsOwnsOnlyExactUnknownMemberRange(t *testing.T) {
	ownedRange := Range{Start: Position{Line: 2, Character: 20}, End: Position{Line: 2, Character: 23}}
	otherRange := Range{Start: Position{Line: 7, Character: 4}, End: Position{Line: 7, Character: 10}}
	diagnostics := []Diagnostic{
		{Code: "VBA252", Range: ownedRange},
		{Code: "VB033", Range: ownedRange},
		{Code: "VB033", Range: otherRange},
		{Code: "VBA211", Range: ownedRange},
	}
	got := preferUnavailableWorksheetFunctionDiagnostics(diagnostics)
	if len(got) != 3 {
		t.Fatalf("deduplicated diagnostics = %+v, want VBA252, unrelated VB033, and VBA211", got)
	}
	for _, diagnostic := range got {
		if diagnostic.Code == "VB033" && diagnostic.Range == ownedRange {
			t.Fatalf("same-range VB033 was not removed: %+v", got)
		}
	}
}

func newGeneratedWorksheetFunctionAnalyzer(t *testing.T) Analyzer {
	t.Helper()
	return newWorksheetFunctionAnalyzer(t, "typelib", "generated")
}

func newWorksheetFunctionAnalyzer(t *testing.T, source, confidence string) Analyzer {
	t.Helper()
	db := vbadb.New()
	body := fmt.Sprintf(`{
  "types": [
    {
      "name": "Excel.Application",
      "library": "Excel",
      "kind": "class",
      "source": %q,
      "confidence": %q,
      "properties": [{"name": "WorksheetFunction", "return_type": "Excel.WorksheetFunction"}]
    },
    {
      "name": "Excel.WorksheetFunction",
      "library": "Excel",
      "kind": "class",
      "aliases": ["WorksheetFunction"],
      "source": %q,
      "confidence": %q,
      "methods": [{"name": "Sum", "return_type": "Double"}]
    }
  ],
  "global_values": {"Application": "Excel.Application"}
}`, source, confidence, source, confidence)
	if err := db.MergeJSON([]byte(body)); err != nil {
		t.Fatal(err)
	}
	return Analyzer{RootDir: t.TempDir(), Config: config.Default(), DB: db}
}
