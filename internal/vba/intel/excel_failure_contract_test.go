package intel

import (
	"path/filepath"
	"testing"
)

func TestExcelAPIFailureContractDiagnosticsUseResolvedMembers(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	addFailureContractMembers(t, analyzer)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim rng As Range
    Dim lateBound As Object
    rng.SpecialCells xlCellTypeVisible
    WorksheetFunction.Match "key", rng, 0
    Debug.Print Application.VLookup("key", rng, 2, False)
    Application.Match "discarded", rng, 0
    Consume(Application.XLookup("key", rng, rng))
    If IsError(Application.Match("key", rng, 0)) Then Exit Sub
    lateBound.Match "ignored", rng, 0
End Sub
`,
	}

	findings := analyzer.ResolvedExcelAPIFailureCalls(doc)
	if len(findings) != 5 {
		t.Fatalf("resolved failure-contract calls = %+v, want five", findings)
	}
	byAPI := map[string]ExcelAPIFailureCall{}
	for _, finding := range findings {
		byAPI[finding.API] = finding
	}
	if got := byAPI["Range.SpecialCells"]; got.Contract != ExcelAPIFailureRaisesError {
		t.Fatalf("Range.SpecialCells contract = %+v, want raises_error", got)
	}
	if got := byAPI["WorksheetFunction.Match"]; got.Contract != ExcelAPIFailureRaisesError {
		t.Fatalf("WorksheetFunction.Match contract = %+v, want raises_error", got)
	}
	if got := byAPI["Application.VLookup"]; got.Contract != ExcelAPIFailureReturnsErrorValue {
		t.Fatalf("Application.VLookup contract = %+v, want returns_error_value", got)
	}
	if got := byAPI["Application.XLookup"]; got.Contract != ExcelAPIFailureReturnsErrorValue {
		t.Fatalf("nested Application.XLookup contract = %+v, want returns_error_value", got)
	}
	if got := byAPI["Application.Match"]; !got.DirectIsError {
		t.Fatalf("Application.Match direct IsError guard = %+v, want true", got)
	}

	diagnostics := analyzer.ExcelAPIFailureContractDiagnostics(doc)
	if len(diagnostics) != 4 {
		t.Fatalf("VBA218 diagnostics = %+v, want four (direct IsError suppressed)", diagnostics)
	}
	for _, diagnostic := range diagnostics {
		if diagnostic.Code != "VBA218" || diagnostic.Severity != "warning" || diagnostic.Rule != "VBA218" {
			t.Fatalf("unexpected VBA218 diagnostic: %+v", diagnostic)
		}
	}
	if diagnostics[0].Message != "Range.SpecialCells may raise a runtime error when no result is available; handle that error path locally." {
		t.Fatalf("Range.SpecialCells message = %q", diagnostics[0].Message)
	}
	wantXLookup := "Application.XLookup may return a Variant/Error when no result is available; check it with IsError before consuming the value."
	seenXLookup := false
	for _, diagnostic := range diagnostics {
		if diagnostic.Message == wantXLookup {
			seenXLookup = true
		}
	}
	if !seenXLookup {
		t.Fatalf("VBA218 diagnostics = %+v, missing %q", diagnostics, wantXLookup)
	}
}

func TestExcelAPIFailureContractDiagnosticsIgnoreDiscardedApplicationResult(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	addFailureContractMembers(t, analyzer)
	diagnostics := analyzer.ExcelAPIFailureContractDiagnostics(Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim rng As Range
    Application.Match "discarded", rng, 0
End Sub
`,
	})
	if len(diagnostics) != 0 {
		t.Fatalf("discarded Application.Match result should not report VBA218: %+v", diagnostics)
	}
}

func TestExcelAPIFailureContractDiagnosticsCoverWorksheetFunctionMembersAndWith(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	addFailureContractMembers(t, analyzer)
	diagnostics := analyzer.ExcelAPIFailureContractDiagnostics(Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Dim rng As Range
    With rng
        .SpecialCells(xlCellTypeVisible)
    End With
    WorksheetFunction.VLookup "key", rng, 2, False
    Application.WorksheetFunction.Match "key", rng, 0
    WorksheetFunction.XLookup "key", rng, rng
    WorksheetFunction.Index rng, 1, 1
End Sub
`,
	})
	if len(diagnostics) != 5 {
		t.Fatalf("VBA218 diagnostics = %+v, want five", diagnostics)
	}
	want := []string{
		"Range.SpecialCells may raise a runtime error when no result is available; handle that error path locally.",
		"WorksheetFunction.VLookup may raise a runtime error when no result is available; handle that error path locally.",
		"WorksheetFunction.XLookup may raise a runtime error when no result is available; handle that error path locally.",
		"WorksheetFunction.Index may raise a runtime error when no result is available; handle that error path locally.",
	}
	seen := map[string]bool{}
	for _, diagnostic := range diagnostics {
		seen[diagnostic.Message] = true
	}
	for _, message := range want {
		if !seen[message] {
			t.Fatalf("VBA218 messages = %+v, missing %q", diagnostics, message)
		}
	}
}

func addFailureContractMembers(t *testing.T, analyzer Analyzer) {
	t.Helper()
	if err := analyzer.DB.MergeJSON([]byte(`{
  "types": [
    { "name": "Excel.Range", "methods": [
      { "name": "SpecialCells", "return_type": "Excel.Range", "parameters": [{ "name": "Type", "type": "Variant" }] }
    ] },
    { "name": "Excel.Application", "methods": [
      { "name": "Match", "return_type": "Variant", "parameters": [{ "name": "Arg1", "type": "Variant" }] },
      { "name": "VLookup", "return_type": "Variant", "parameters": [{ "name": "Arg1", "type": "Variant" }] },
      { "name": "XLookup", "return_type": "Variant", "parameters": [{ "name": "Arg1", "type": "Variant" }] }
    ] },
    { "name": "Excel.WorksheetFunction", "methods": [
      { "name": "Match", "return_type": "Variant", "parameters": [{ "name": "Arg1", "type": "Variant" }] },
      { "name": "VLookup", "return_type": "Variant", "parameters": [{ "name": "Arg1", "type": "Variant" }] },
      { "name": "XLookup", "return_type": "Variant", "parameters": [{ "name": "Arg1", "type": "Variant" }] },
      { "name": "Index", "return_type": "Variant", "parameters": [{ "name": "Arg1", "type": "Variant" }] }
    ] }
  ]
}`)); err != nil {
		t.Fatal(err)
	}
}
