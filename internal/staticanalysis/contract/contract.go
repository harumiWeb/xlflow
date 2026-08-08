// Package contract defines protocol-neutral diagnostic expectations shared by
// static-analysis evidence consumers.
package contract

import (
	"errors"
	"fmt"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

// Diagnostic is the normalized shape consumed by diagnostic contract checks.
type Diagnostic struct {
	Code     string `json:"code"`
	Severity string `json:"severity,omitempty"`
	Range    *Range `json:"range,omitempty"`
	Surface  string `json:"surface,omitempty"`
}

// Range identifies an exact source range.
type Range struct {
	StartLine   int `json:"start_line,omitempty"`
	StartColumn int `json:"start_column,omitempty"`
	EndLine     int `json:"end_line,omitempty"`
	EndColumn   int `json:"end_column,omitempty"`
}

// DiagnosticExpectation describes a required or forbidden diagnostic.
// Omitted severity, range, and surfaces act as wildcards.
type DiagnosticExpectation struct {
	Code     string   `json:"code"`
	Severity string   `json:"severity,omitempty"`
	Range    *Range   `json:"range,omitempty"`
	Surfaces []string `json:"surfaces,omitempty"`
}

// ExpectationSet groups required and forbidden diagnostic expectations.
type ExpectationSet struct {
	ExpectedDiagnostics  []DiagnosticExpectation `json:"expected_diagnostics,omitempty"`
	ForbiddenDiagnostics []DiagnosticExpectation `json:"forbidden_diagnostics,omitempty"`
}

// Check verifies required and forbidden diagnostics. Each actual diagnostic
// can satisfy at most one required expectation. Matching may reassign a broad
// expectation so a more constrained expectation can consume its only match.
func Check(expect ExpectationSet, actual []Diagnostic) error {
	for _, forbidden := range expect.ForbiddenDiagnostics {
		if countDiagnostics(forbidden, actual) > 0 {
			return fmt.Errorf("forbidden diagnostic %s was emitted", forbidden.Code)
		}
	}
	type requiredSlot struct {
		expectation DiagnosticExpectation
		surface     string
	}
	slots := make([]requiredSlot, 0, len(expect.ExpectedDiagnostics))
	for _, required := range expect.ExpectedDiagnostics {
		if len(required.Surfaces) > 0 {
			for _, surface := range required.Surfaces {
				slots = append(slots, requiredSlot{expectation: required, surface: surface})
			}
			continue
		}
		slots = append(slots, requiredSlot{expectation: required})
	}
	actualToSlot := make([]int, len(actual))
	for index := range actualToSlot {
		actualToSlot[index] = -1
	}
	var assign func(int, []bool) bool
	assign = func(slotIndex int, seen []bool) bool {
		slot := slots[slotIndex]
		for actualIndex, diagnostic := range actual {
			if seen[actualIndex] || slot.surface != "" && diagnostic.Surface != slot.surface || !Matches(slot.expectation, diagnostic) {
				continue
			}
			seen[actualIndex] = true
			previous := actualToSlot[actualIndex]
			if previous < 0 || assign(previous, seen) {
				actualToSlot[actualIndex] = slotIndex
				return true
			}
		}
		return false
	}
	for slotIndex, slot := range slots {
		if assign(slotIndex, make([]bool, len(actual))) {
			continue
		}
		if slot.surface != "" {
			return fmt.Errorf("expected diagnostic %s was not emitted on %s", slot.expectation.Code, slot.surface)
		}
		return fmt.Errorf("expected diagnostic %s was not emitted", slot.expectation.Code)
	}
	return nil
}

// CheckProjections verifies diagnostics grouped by public analysis surface and
// rejects projection-specific duplicates or changes in severity/range.
func CheckProjections(expect ExpectationSet, projections map[string][]Diagnostic) error {
	all := make([]Diagnostic, 0)
	surfaces := make([]string, 0, len(projections))
	for surface := range projections {
		surfaces = append(surfaces, surface)
	}
	sort.Strings(surfaces)
	for _, surface := range surfaces {
		diagnostics := projections[surface]
		for _, diagnostic := range diagnostics {
			if diagnostic.Surface != "" && diagnostic.Surface != surface {
				return fmt.Errorf("diagnostic %s has surface %q in %q projection", diagnostic.Code, diagnostic.Surface, surface)
			}
			diagnostic.Surface = surface
			all = append(all, diagnostic)
		}
	}
	if err := Check(expect, all); err != nil {
		return err
	}
	for _, required := range expect.ExpectedDiagnostics {
		for _, surface := range surfaces {
			diagnostics := projections[surface]
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

// ValidateExpectations checks expectations against the static-analysis rule
// registry, including canonical IDs and supported severities/surfaces.
func ValidateExpectations(expectations []DiagnosticExpectation) error {
	for _, expectation := range expectations {
		rule, err := CanonicalRuleMetadata(expectation.Code)
		if err != nil {
			return fmt.Errorf("invalid diagnostic code %q: %w", expectation.Code, err)
		}
		if severity := expectation.Severity; severity != "" && !ruleSupportsSeverity(rule, severity) {
			return fmt.Errorf("diagnostic %q has unsupported severity %q", expectation.Code, severity)
		}
		seenSurfaces := map[string]struct{}{}
		for _, surface := range expectation.Surfaces {
			if !ruleSupportsSurface(rule, surface) {
				return fmt.Errorf("diagnostic %q has unsupported surface %q", expectation.Code, surface)
			}
			if _, duplicate := seenSurfaces[surface]; duplicate {
				return fmt.Errorf("diagnostic %q repeats surface %q", expectation.Code, surface)
			}
			seenSurfaces[surface] = struct{}{}
		}
		if value := expectation.Range; value != nil {
			if value.StartLine < 0 || value.StartColumn < 0 || value.EndLine < 0 || value.EndColumn < 0 {
				return fmt.Errorf("diagnostic %q has a negative source range", expectation.Code)
			}
			if value.EndLine < value.StartLine || value.EndLine == value.StartLine && value.EndColumn < value.StartColumn {
				return fmt.Errorf("diagnostic %q has an incoherent source range", expectation.Code)
			}
		}
	}
	return nil
}

// Validate checks both halves of an expectation set against the registry.
func Validate(expect ExpectationSet) error {
	if err := ValidateExpectations(expect.ExpectedDiagnostics); err != nil {
		return fmt.Errorf("expected diagnostics: %w", err)
	}
	if err := ValidateExpectations(expect.ForbiddenDiagnostics); err != nil {
		return fmt.Errorf("forbidden diagnostics: %w", err)
	}
	return nil
}

// Normalize gives callers deterministic ordering while preserving duplicates.
func Normalize(diagnostics []Diagnostic) []Diagnostic {
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
		if Matches(expect, diagnostic) {
			count++
		}
	}
	return count
}

// Matches reports whether a diagnostic satisfies an expectation's optional
// severity, exact range, and surface-set constraints.
func Matches(expect DiagnosticExpectation, diagnostic Diagnostic) bool {
	return diagnostic.Code == expect.Code &&
		(expect.Severity == "" || diagnostic.Severity == expect.Severity) &&
		(expect.Range == nil || rangesEqual(expect.Range, diagnostic.Range)) &&
		(len(expect.Surfaces) == 0 || containsString(expect.Surfaces, diagnostic.Surface))
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

// CanonicalRuleMetadata resolves a canonical diagnostic ID through the shared
// static-analysis rule registry.
func CanonicalRuleMetadata(code string) (rules.RuleMetadata, error) {
	if strings.TrimSpace(code) == "" {
		return rules.RuleMetadata{}, errors.New("code is empty")
	}
	rule, ok := rules.Lookup(code)
	if !ok {
		return rules.RuleMetadata{}, errors.New("code is not in the static-analysis rule registry")
	}
	if code != rule.ID {
		return rules.RuleMetadata{}, fmt.Errorf("code must use canonical registry ID %q", rule.ID)
	}
	return rule, nil
}

func ruleSupportsSurface(rule rules.RuleMetadata, surface string) bool {
	for _, supported := range rule.Surfaces {
		if surface == string(supported) {
			return true
		}
	}
	return false
}

func ruleSupportsSeverity(rule rules.RuleMetadata, severity string) bool {
	for _, supported := range rule.SupportedSeverities {
		if severity == string(supported) {
			return true
		}
	}
	return false
}

func rangeKey(value *Range) string {
	if value == nil {
		return ""
	}
	return fmt.Sprintf("%08d:%08d:%08d:%08d", value.StartLine, value.StartColumn, value.EndLine, value.EndColumn)
}
