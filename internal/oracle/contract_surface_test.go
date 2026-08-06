package oracle

import "testing"

func TestDiagnosticSurfaceExpectationRequiresEveryNamedProjection(t *testing.T) {
	expect := AnalysisExpectation{ExpectedDiagnostics: []DiagnosticExpectation{{Code: "VBA001", Severity: "warning", Surfaces: []string{"lint", "analyze"}}}}
	projections := map[string][]Diagnostic{
		"lint":    {{Code: "VBA001", Severity: "warning", Surface: "lint"}},
		"analyze": {{Code: "VBA001", Severity: "warning", Surface: "analyze"}},
	}
	if err := CheckDiagnosticProjections(expect, projections); err != nil {
		t.Fatal(err)
	}
	delete(projections, "analyze")
	if err := CheckDiagnosticProjections(expect, projections); err == nil {
		t.Fatal("expected missing explicitly scoped projection to fail")
	}
}

func TestDiagnosticContractRejectsDuplicateNamedSurface(t *testing.T) {
	expect := AnalysisExpectation{ExpectedDiagnostics: []DiagnosticExpectation{{Code: "VBA001", Surfaces: []string{"lint", "lint"}}}}
	if err := validateDiagnosticExpectations(expect.ExpectedDiagnostics, "expected", "sample"); err == nil {
		t.Fatal("expected duplicate surface validation error")
	}
}

func TestDiagnosticProjectionRejectsMismatchedSurface(t *testing.T) {
	expect := AnalysisExpectation{ExpectedDiagnostics: []DiagnosticExpectation{{Code: "VBA001", Surfaces: []string{"analyze"}}}}
	projections := map[string][]Diagnostic{
		"lint": {{Code: "VBA001", Surface: "analyze"}},
	}
	if err := CheckDiagnosticProjections(expect, projections); err == nil {
		t.Fatal("expected mismatched projection surface to fail")
	}
}
