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
