package intel

import (
	"path/filepath"
	"testing"
)

func TestIssue772ApplicationMinMaxResolveWithoutUnknownMemberDiagnostics(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	if err := analyzer.DB.MergeJSON([]byte(`{
  "types": [{
    "name": "Excel.Application",
    "confidence": "generated",
    "source": "typelib"
  }]
}`)); err != nil {
		t.Fatal(err)
	}
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim sourceRange As Range
    Dim dataMinimum As Double
    Dim dataMaximum As Double
    dataMinimum = Application.Min(sourceRange)
    dataMaximum = Application.Max(sourceRange)
End Sub
`,
	}

	for _, name := range []string{"Min", "Max"} {
		if _, ok := analyzer.DB.ResolveMember("Excel.Application", name); !ok {
			t.Fatalf("Excel.Application.%s should be present in the merged object model", name)
		}
	}
	diagnostics := diagnosticsByCode(analyzer.Diagnostics(doc), "VB033")
	if len(diagnostics) != 0 {
		t.Fatalf("Application.Min/Max should resolve without VB033: %+v", diagnostics)
	}
}

func TestIssue772DecimalLiteralsInMultilineWorksheetFunctionCallsAreNotUndeclared(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim dataMinimum As Double
    Dim dataMaximum As Double
    Dim minimumValue As Double
    Dim maximumValue As Double
    minimumValue = _
        WorksheetFunction.Floor_Math( _
            dataMinimum, _
            0.5) - 0.5
    maximumValue = _
        WorksheetFunction.Ceiling_Math( _
            dataMaximum, _
            0.5) + 0.5
    Debug.Print minimumValue, maximumValue
End Sub
`,
	}

	diagnostics := analyzer.Diagnostics(doc)
	if len(diagnostics) != 0 {
		t.Fatalf("valid multiline WorksheetFunction calls should be diagnostic-free: %+v", diagnostics)
	}
}

func TestParenlessCallOnLineValidatesCompleteTargets(t *testing.T) {
	tests := []struct {
		name       string
		line       string
		want       bool
		wantTarget string
	}{
		{name: "identifier", line: "Foo value", want: true, wantTarget: "Foo"},
		{name: "member", line: "Foo.Bar value", want: true, wantTarget: "Foo.Bar"},
		{name: "with member", line: ".Bar value", want: true, wantTarget: ".Bar"},
		{name: "parenthesized member", line: "Foo(1).Bar value", want: true, wantTarget: "Foo(1).Bar"},
		{name: "trailing dot", line: "Foo. value"},
		{name: "trailing text after parenthesized target", line: "Foo(1)Bar value"},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			call, ok := parenlessCallOnLine(tc.line)
			if ok != tc.want {
				t.Fatalf("parenlessCallOnLine(%q) = (%+v, %t), want ok=%t", tc.line, call, ok, tc.want)
			}
			if tc.want && call.Target != tc.wantTarget {
				t.Fatalf("parenlessCallOnLine(%q) target = %q, want %q", tc.line, call.Target, tc.wantTarget)
			}
		})
	}
}
