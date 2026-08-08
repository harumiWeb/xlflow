package oracle

import staticcontract "github.com/harumiWeb/xlflow/internal/staticanalysis/contract"

// Diagnostic and its nested range remain aliases so existing oracle callers
// and fixture tests retain their source and JSON compatibility.
type Diagnostic = staticcontract.Diagnostic

// CheckDiagnostics applies the case's deterministic analyzer expectation.
func (c Case) CheckDiagnostics(actual []Diagnostic) error {
	return CheckDiagnosticContract(c.Analysis, actual)
}

// CheckDiagnosticSurfaces applies the case expectation to lint/analyze/LSP
// projections represented by surface name.
func (c Case) CheckDiagnosticSurfaces(projections map[string][]Diagnostic) error {
	return CheckDiagnosticProjections(c.Analysis, projections)
}

// CheckDiagnosticContract preserves the oracle API while delegating generic
// diagnostic matching to the static-analysis contract package.
func CheckDiagnosticContract(expect AnalysisExpectation, actual []Diagnostic) error {
	return staticcontract.Check(expectationSet(expect), actual)
}

// CheckDiagnosticProjections preserves the oracle API while delegating shared
// projection checks to the static-analysis contract package.
func CheckDiagnosticProjections(expect AnalysisExpectation, projections map[string][]Diagnostic) error {
	return staticcontract.CheckProjections(expectationSet(expect), projections)
}

// NormalizeDiagnostics preserves the oracle API for existing fixture callers.
func NormalizeDiagnostics(diagnostics []Diagnostic) []Diagnostic {
	return staticcontract.Normalize(diagnostics)
}

func expectationSet(expect AnalysisExpectation) staticcontract.ExpectationSet {
	return staticcontract.ExpectationSet{
		ExpectedDiagnostics:  expect.ExpectedDiagnostics,
		ForbiddenDiagnostics: expect.ForbiddenDiagnostics,
	}
}

func containsString(values []string, value string) bool {
	for _, candidate := range values {
		if candidate == value {
			return true
		}
	}
	return false
}
