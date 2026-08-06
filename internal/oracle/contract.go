package oracle

import (
	"fmt"
	"sort"
)

// Diagnostic is the normalized shape consumed by fixture contract tests. The
// analyzer and LSP adapters can project their native diagnostics into this
// shape without importing Excel or COM code.
type Diagnostic struct {
	Code     string `json:"code"`
	Severity string `json:"severity,omitempty"`
	Range    *Range `json:"range,omitempty"`
	Surface  string `json:"surface,omitempty"`
}

// CheckDiagnostics applies the case's deterministic analyzer expectation.
func (c Case) CheckDiagnostics(actual []Diagnostic) error {
	return CheckDiagnosticContract(c.Analysis, actual)
}

// CheckDiagnosticSurfaces applies the case expectation to lint/analyze/LSP
// projections represented by surface name.
func (c Case) CheckDiagnosticSurfaces(projections map[string][]Diagnostic) error {
	return CheckDiagnosticProjections(c.Analysis, projections)
}

func CheckDiagnosticContract(expect AnalysisExpectation, actual []Diagnostic) error {
	for _, forbidden := range expect.ForbiddenDiagnostics {
		if countDiagnostics(forbidden, actual) > 0 {
			return fmt.Errorf("forbidden diagnostic %s was emitted", forbidden.Code)
		}
	}
	used := make([]bool, len(actual))
	for _, required := range expect.ExpectedDiagnostics {
		if len(required.Surfaces) > 0 {
			// An explicitly scoped expectation applies once on every named
			// projection surface. This prevents a fixture from claiming LSP
			// parity while only checking one surface in the union.
			for _, surface := range required.Surfaces {
				matched := false
				for i, diagnostic := range actual {
					if used[i] || diagnostic.Surface != surface || !diagnosticMatches(required, diagnostic) {
						continue
					}
					used[i] = true
					matched = true
					break
				}
				if !matched {
					return fmt.Errorf("expected diagnostic %s was not emitted on %s", required.Code, surface)
				}
			}
			continue
		}
		matched := false
		for i, diagnostic := range actual {
			if used[i] || !diagnosticMatches(required, diagnostic) {
				continue
			}
			used[i] = true
			matched = true
			break
		}
		if !matched {
			return fmt.Errorf("expected diagnostic %s was not emitted", required.Code)
		}
	}
	return nil
}

// CheckDiagnosticProjections applies one fixture contract to diagnostics
// projected by lint, analyze, and LSP adapters. It is intentionally Excel-free
// and can be called from ordinary Go tests.
func CheckDiagnosticProjections(expect AnalysisExpectation, projections map[string][]Diagnostic) error {
	all := make([]Diagnostic, 0)
	for surface, diagnostics := range projections {
		for _, diagnostic := range diagnostics {
			if diagnostic.Surface != "" && diagnostic.Surface != surface {
				return fmt.Errorf("diagnostic %s has surface %q in %q projection", diagnostic.Code, diagnostic.Surface, surface)
			}
			diagnostic.Surface = surface
			all = append(all, diagnostic)
		}
	}
	if err := CheckDiagnosticContract(expect, all); err != nil {
		return err
	}
	// Duplicate diagnostics are a contract failure when the fixture explicitly
	// claims that a diagnostic is present. Unrelated diagnostics are left to
	// the owning analyzer surface because some projections intentionally use
	// different ranges or severity policies for advisory findings.
	for _, required := range expect.ExpectedDiagnostics {
		for surface, diagnostics := range projections {
			if len(required.Surfaces) > 0 && !containsString(required.Surfaces, surface) {
				continue
			}
			count := 0
			for _, diagnostic := range diagnostics {
				if diagnostic.Code == required.Code &&
					(required.Severity == "" || diagnostic.Severity == required.Severity) &&
					(required.Range == nil || rangesEqual(required.Range, diagnostic.Range)) {
					count++
				}
			}
			if count > 1 {
				return fmt.Errorf("diagnostic %s is duplicated on %s", required.Code, surface)
			}
		}
	}
	// A shared rule must not silently change severity/range when projected to a
	// different protocol surface. Duplicate code-only observations are allowed
	// because one source can legitimately produce multiple diagnostics.
	expectedCodes := map[string]bool{}
	for _, required := range expect.ExpectedDiagnostics {
		expectedCodes[required.Code] = true
	}
	seen := map[string]map[string]Diagnostic{}
	for _, diagnostic := range all {
		if !expectedCodes[diagnostic.Code] {
			continue
		}
		if required, ok := expectedForCode(expect.ExpectedDiagnostics, diagnostic.Code); ok && len(required.Surfaces) > 0 && !containsString(required.Surfaces, diagnostic.Surface) {
			continue
		}
		bySurface := seen[diagnostic.Code]
		if bySurface == nil {
			bySurface = map[string]Diagnostic{}
			seen[diagnostic.Code] = bySurface
		}
		if _, ok := bySurface[diagnostic.Surface]; ok {
			continue
		}
		for _, prior := range bySurface {
			if prior.Severity != diagnostic.Severity || !rangesEqual(prior.Range, diagnostic.Range) {
				return fmt.Errorf("diagnostic %s differs across projections", diagnostic.Code)
			}
		}
		bySurface[diagnostic.Surface] = diagnostic
	}
	return nil
}

func expectedForCode(expectations []DiagnosticExpectation, code string) (DiagnosticExpectation, bool) {
	for _, expectation := range expectations {
		if expectation.Code == code {
			return expectation, true
		}
	}
	return DiagnosticExpectation{}, false
}

func countDiagnostics(expect DiagnosticExpectation, actual []Diagnostic) int {
	count := 0
	for _, diagnostic := range actual {
		if diagnosticMatches(expect, diagnostic) {
			count++
		}
	}
	return count
}

func diagnosticMatches(expect DiagnosticExpectation, diagnostic Diagnostic) bool {
	if diagnostic.Code != expect.Code {
		return false
	}
	if expect.Severity != "" && diagnostic.Severity != expect.Severity {
		return false
	}
	if expect.Range != nil && !rangesEqual(expect.Range, diagnostic.Range) {
		return false
	}
	if len(expect.Surfaces) > 0 && !containsString(expect.Surfaces, diagnostic.Surface) {
		return false
	}
	return true
}

func rangesEqual(left, right *Range) bool {
	if left == nil || right == nil {
		return left == right
	}
	return *left == *right
}

func containsString(values []string, value string) bool {
	for _, candidate := range values {
		if candidate == value {
			return true
		}
	}
	return false
}

// NormalizeDiagnostics gives tests deterministic ordering while preserving
// duplicates (duplicate diagnostics are meaningful contract failures).
func NormalizeDiagnostics(diagnostics []Diagnostic) []Diagnostic {
	result := append([]Diagnostic(nil), diagnostics...)
	sort.SliceStable(result, func(i, j int) bool {
		if result[i].Code != result[j].Code {
			return result[i].Code < result[j].Code
		}
		if result[i].Severity != result[j].Severity {
			return result[i].Severity < result[j].Severity
		}
		if result[i].Surface != result[j].Surface {
			return result[i].Surface < result[j].Surface
		}
		return rangeKey(result[i].Range) < rangeKey(result[j].Range)
	})
	return result
}

func rangeKey(value *Range) string {
	if value == nil {
		return ""
	}
	return fmt.Sprintf("%08d:%08d:%08d:%08d", value.StartLine, value.StartColumn, value.EndLine, value.EndColumn)
}
