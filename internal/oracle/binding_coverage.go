package oracle

import (
	"fmt"
	"sort"
	"strings"

	staticcontract "github.com/harumiWeb/xlflow/internal/staticanalysis/contract"
	"github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

// BindingCoverage is the deterministic, Excel-free summary of the oracle
// corpus. Rules and fixture IDs are sorted before they are stored in the
// report so callers can safely render or compare it.
type BindingCoverage struct {
	AssertedFixtures int
	BoundFixtures    int
	PartialFixtures  int
	UnboundFixtures  int
	NotApplicable    int

	CompleteRules        int
	MissingNegativeRules int
	MissingPositiveRules int

	BoundIDs         []string
	PartialIDs       []string
	UnboundIDs       []string
	NotApplicableIDs []string
	Rules            []BindingRuleCoverage
}

// BindingRuleCoverage describes positive and negative evidence for one
// diagnostic rule. Missing fields explain why a bound rule is incomplete.
type BindingRuleCoverage struct {
	Code              string
	RejectedCases     []string
	AcceptedControls  []string
	CoveredSurfaces   []string
	MissingSurfaces   []string
	MissingPositive   bool
	MissingNegative   bool
	InformationalOnly bool
}

// ValidateBindingCoverage validates relationships that require the complete
// fixture corpus. It returns a report even when validation fails, allowing a
// caller to print useful deterministic diagnostics before failing CI.
func ValidateBindingCoverage(cases []Case) (BindingCoverage, error) {
	report := BindingCoverage{}
	byID := make(map[string]Case, len(cases))
	var validationErrors []string
	for _, c := range cases {
		if _, exists := byID[c.ID]; exists {
			validationErrors = append(validationErrors, fmt.Sprintf("duplicate oracle case %q", c.ID))
			continue
		}
		byID[c.ID] = c
		switch c.VBE.Expected {
		case ExpectedAccepted, ExpectedRejected:
			report.AssertedFixtures++
		}
		switch c.Analysis.BindingStatus {
		case BindingBound:
			report.BoundFixtures++
			report.BoundIDs = append(report.BoundIDs, c.ID)
		case BindingPartiallyBound:
			report.PartialFixtures++
			report.PartialIDs = append(report.PartialIDs, c.ID)
		case BindingUnbound:
			report.UnboundFixtures++
			report.UnboundIDs = append(report.UnboundIDs, c.ID)
		case BindingNotApplicable:
			report.NotApplicable++
			report.NotApplicableIDs = append(report.NotApplicableIDs, c.ID)
		}
		for _, code := range c.Analysis.RuleCodes {
			if _, err := staticcontract.CanonicalRuleMetadata(code); err != nil {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: %v", c.ID, err))
			}
		}
		for _, expectation := range append(append([]DiagnosticExpectation(nil), c.Analysis.ExpectedDiagnostics...), c.Analysis.ForbiddenDiagnostics...) {
			if _, err := staticcontract.CanonicalRuleMetadata(expectation.Code); err != nil {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: %v", c.ID, err))
			}
		}
	}

	for _, c := range cases {
		controls := c.Analysis.NegativeControls
		if len(controls) == 0 {
			continue
		}
		if c.VBE.Expected != ExpectedRejected {
			validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: negative_controls are allowed only on rejected fixtures", c.ID))
		}
		if c.Analysis.BindingStatus != BindingPartiallyBound && c.Analysis.BindingStatus != BindingBound {
			validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: negative_controls require partially-bound or bound status", c.ID))
		}
		seenControls := make(map[string]struct{}, len(controls))
		for _, controlID := range controls {
			trimmed := strings.TrimSpace(controlID)
			if trimmed == "" || trimmed != controlID || !caseIDPattern.MatchString(trimmed) {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: invalid negative control ID %q", c.ID, controlID))
				continue
			}
			if _, duplicate := seenControls[trimmed]; duplicate {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: duplicate negative control %q", c.ID, trimmed))
				continue
			}
			seenControls[trimmed] = struct{}{}
			if trimmed == c.ID {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: negative control cannot reference itself", c.ID))
				continue
			}
			control, ok := byID[trimmed]
			if !ok {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: negative control %q does not exist", c.ID, trimmed))
				continue
			}
			if control.VBE.Expected != ExpectedAccepted {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: negative control %q is not VBE accepted", c.ID, trimmed))
			}
			if control.Analysis.BindingStatus != BindingBound {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: negative control %q must be bound", c.ID, trimmed))
			}
			matchingRule := false
			for _, forbidden := range control.Analysis.ForbiddenDiagnostics {
				if containsString(c.Analysis.RuleCodes, forbidden.Code) {
					matchingRule = true
					break
				}
			}
			if !matchingRule {
				validationErrors = append(validationErrors, fmt.Sprintf("oracle case %q: negative control %q forbids none of the parent rule codes", c.ID, trimmed))
			}
		}
	}

	// Keep the graph check explicit even though the accepted-only target rule
	// makes cycles impossible in valid data. This protects the contract if the
	// relationship rules are extended later.
	validationErrors = append(validationErrors, validateNegativeControlCycles(cases, byID)...)

	ruleMap := make(map[string]*BindingRuleCoverage)
	for _, c := range cases {
		if c.Analysis.BindingStatus != BindingBound {
			continue
		}
		for _, code := range c.Analysis.RuleCodes {
			if supplementalAcceptedStyleBinding(c, code) {
				continue
			}
			coverage := ruleMap[code]
			if coverage == nil {
				coverage = &BindingRuleCoverage{Code: code}
				ruleMap[code] = coverage
			}
			if c.VBE.Expected == ExpectedRejected {
				coverage.RejectedCases = appendUnique(coverage.RejectedCases, c.ID)
			}
		}
	}

	for _, c := range cases {
		if c.Analysis.BindingStatus != BindingBound || c.VBE.Expected != ExpectedRejected {
			continue
		}
		for _, code := range c.Analysis.RuleCodes {
			coverage := ruleMap[code]
			positiveSurfaces := expectedSurfaces(c.Analysis.ExpectedDiagnostics, code)
			controlIDs := c.Analysis.NegativeControls
			matchedControl := false
			localCoveredSurfaces := []string{}
			for _, controlID := range controlIDs {
				control, ok := byID[controlID]
				if !ok || control.Analysis.BindingStatus != BindingBound || control.VBE.Expected != ExpectedAccepted {
					continue
				}
				for _, forbidden := range control.Analysis.ForbiddenDiagnostics {
					if forbidden.Code != code {
						continue
					}
					matchedControl = true
					coverage.AcceptedControls = appendUnique(coverage.AcceptedControls, control.ID)
					for _, surface := range effectiveSurfaces(code, forbidden.Surfaces) {
						coverage.CoveredSurfaces = appendUnique(coverage.CoveredSurfaces, surface)
						localCoveredSurfaces = appendUnique(localCoveredSurfaces, surface)
					}
				}
			}
			if !matchedControl {
				coverage.MissingNegative = true
			}
			localMissingSurfaces := []string{}
			for _, surface := range positiveSurfaces {
				if !containsString(localCoveredSurfaces, surface) {
					localMissingSurfaces = appendUnique(localMissingSurfaces, surface)
					coverage.MissingSurfaces = appendUnique(coverage.MissingSurfaces, surface)
				}
			}
			if len(localMissingSurfaces) > 0 {
				validationErrors = append(validationErrors, fmt.Sprintf("fixture %s rule %s: missing accepted coverage on surfaces %s", c.ID, code, strings.Join(sortedUnique(localMissingSurfaces), ", ")))
			}
		}
	}

	for _, c := range cases {
		if c.Analysis.BindingStatus != BindingBound || c.VBE.Expected != ExpectedAccepted {
			continue
		}
		for _, code := range c.Analysis.RuleCodes {
			if supplementalAcceptedStyleBinding(c, code) {
				continue
			}
			coverage := ruleMap[code]
			boundReference := acceptedReferencedByRejected(c.ID, code, cases, BindingBound)
			partialReference := acceptedReferencedByRejected(c.ID, code, cases, BindingPartiallyBound)
			if coverage == nil {
				coverage = &BindingRuleCoverage{Code: code}
				ruleMap[code] = coverage
			}
			if !boundReference {
				if partialReference {
					coverage.InformationalOnly = true
				} else {
					coverage.MissingPositive = true
				}
			}
		}
	}

	for _, coverage := range ruleMap {
		coverage.RejectedCases = sortedUnique(coverage.RejectedCases)
		coverage.AcceptedControls = sortedUnique(coverage.AcceptedControls)
		coverage.CoveredSurfaces = sortedUnique(coverage.CoveredSurfaces)
		coverage.MissingSurfaces = sortedUnique(coverage.MissingSurfaces)
		if !coverage.InformationalOnly && (coverage.MissingNegative || len(coverage.MissingSurfaces) > 0 || len(coverage.AcceptedControls) == 0) {
			report.MissingNegativeRules++
		}
		if !coverage.InformationalOnly && (coverage.MissingPositive || len(coverage.RejectedCases) == 0) {
			report.MissingPositiveRules++
		}
		if !coverage.InformationalOnly && !coverage.MissingPositive && !coverage.MissingNegative && len(coverage.RejectedCases) > 0 && len(coverage.AcceptedControls) > 0 && len(coverage.MissingSurfaces) == 0 {
			report.CompleteRules++
		}
		report.Rules = append(report.Rules, *coverage)
	}
	sort.Slice(report.Rules, func(i, j int) bool { return report.Rules[i].Code < report.Rules[j].Code })
	sort.Strings(report.BoundIDs)
	sort.Strings(report.PartialIDs)
	sort.Strings(report.UnboundIDs)
	sort.Strings(report.NotApplicableIDs)

	validationErrors = append(validationErrors, coverageValidationErrors(report.Rules)...)
	if len(validationErrors) > 0 {
		sort.Strings(validationErrors)
		return report, fmt.Errorf("oracle binding coverage validation failed:\n- %s", strings.Join(uniqueStrings(validationErrors), "\n- "))
	}
	return report, nil
}

func supplementalAcceptedStyleBinding(c Case, code string) bool {
	if c.VBE.Expected != ExpectedAccepted || c.Analysis.BindingStatus != BindingBound {
		return false
	}
	rule, err := staticcontract.CanonicalRuleMetadata(code)
	return err == nil && rule.Family == rules.FamilyLint && rule.Category == rules.CategoryMaintainability && rule.EvidenceClass == rules.EvidenceMaintainability && !rule.CompileEquivalent && rule.DefaultSeverity == rules.SeverityWarning && !rule.PreflightBlocking
}

func coverageValidationErrors(rulesCoverage []BindingRuleCoverage) []string {
	var result []string
	for _, coverage := range rulesCoverage {
		if !coverage.InformationalOnly && (coverage.MissingPositive || len(coverage.RejectedCases) == 0) {
			result = append(result, fmt.Sprintf("rule %s: missing rejected positive evidence", coverage.Code))
		}
		if !coverage.InformationalOnly && (coverage.MissingNegative || len(coverage.MissingSurfaces) > 0 || len(coverage.AcceptedControls) == 0) {
			message := fmt.Sprintf("rule %s: missing accepted negative coverage", coverage.Code)
			if len(coverage.MissingSurfaces) > 0 {
				message += fmt.Sprintf(" on surfaces %s", strings.Join(coverage.MissingSurfaces, ", "))
			}
			result = append(result, message)
		}
	}
	return result
}

func validateNegativeControlCycles(cases []Case, byID map[string]Case) []string {
	graph := make(map[string][]string)
	for _, c := range cases {
		graph[c.ID] = append([]string(nil), c.Analysis.NegativeControls...)
	}
	state := make(map[string]uint8)
	var stack []string
	var result []string
	var visit func(string)
	visit = func(id string) {
		switch state[id] {
		case 1:
			start := 0
			for i, value := range stack {
				if value == id {
					start = i
					break
				}
			}
			cycle := append(append([]string(nil), stack[start:]...), id)
			result = append(result, fmt.Sprintf("negative control cycle: %s", strings.Join(cycle, " -> ")))
			return
		case 2:
			return
		}
		state[id] = 1
		stack = append(stack, id)
		for _, target := range graph[id] {
			if _, ok := byID[target]; ok {
				visit(target)
			}
		}
		stack = stack[:len(stack)-1]
		state[id] = 2
	}
	for _, c := range cases {
		visit(c.ID)
	}
	return uniqueStrings(result)
}

func acceptedReferencedByRejected(controlID, code string, cases []Case, status string) bool {
	for _, c := range cases {
		if c.VBE.Expected != ExpectedRejected || c.Analysis.BindingStatus != status || !containsString(c.Analysis.RuleCodes, code) {
			continue
		}
		if containsString(c.Analysis.NegativeControls, controlID) {
			return true
		}
	}
	return false
}

func expectedSurfaces(expectations []DiagnosticExpectation, code string) []string {
	var surfaces []string
	for _, expectation := range expectations {
		if expectation.Code == code {
			surfaces = append(surfaces, effectiveSurfaces(code, expectation.Surfaces)...)
		}
	}
	if len(surfaces) == 0 {
		return effectiveSurfaces(code, nil)
	}
	return sortedUnique(surfaces)
}

func effectiveSurfaces(code string, declared []string) []string {
	if len(declared) > 0 {
		return sortedUnique(append([]string(nil), declared...))
	}
	rule, ok := rules.Lookup(code)
	if !ok {
		return nil
	}
	surfaces := make([]string, 0, len(rule.Surfaces))
	for _, surface := range rule.Surfaces {
		surfaces = append(surfaces, string(surface))
	}
	return sortedUnique(surfaces)
}

func appendUnique(values []string, value string) []string {
	if !containsString(values, value) {
		return append(values, value)
	}
	return values
}

func sortedUnique(values []string) []string {
	result := uniqueStrings(values)
	sort.Strings(result)
	return result
}

func uniqueStrings(values []string) []string {
	seen := make(map[string]struct{}, len(values))
	result := make([]string, 0, len(values))
	for _, value := range values {
		if _, ok := seen[value]; ok {
			continue
		}
		seen[value] = struct{}{}
		result = append(result, value)
	}
	return result
}

// String renders a stable human-readable report suitable for go test -v.
func (r BindingCoverage) String() string {
	var b strings.Builder
	fmt.Fprintln(&b, "VBE oracle binding coverage")
	fmt.Fprintln(&b)
	fmt.Fprintf(&b, "Asserted fixtures: %d\n", r.AssertedFixtures)
	fmt.Fprintf(&b, "Bound: %d\n", r.BoundFixtures)
	fmt.Fprintf(&b, "Partially bound: %d\n", r.PartialFixtures)
	fmt.Fprintf(&b, "Unbound: %d\n", r.UnboundFixtures)
	fmt.Fprintf(&b, "Not applicable: %d\n", r.NotApplicable)
	fmt.Fprintln(&b)
	fmt.Fprintf(&b, "Rules with complete positive/negative coverage: %d\n", r.CompleteRules)
	fmt.Fprintf(&b, "Rules missing accepted controls: %d\n", r.MissingNegativeRules)
	fmt.Fprintf(&b, "Rules missing rejected evidence: %d\n", r.MissingPositiveRules)
	writeFixtureIDs := func(label string, ids []string) {
		fmt.Fprintf(&b, "\n%s\n", label)
		for _, id := range ids {
			fmt.Fprintf(&b, "- %s\n", id)
		}
	}
	writeFixtureIDs("Bound fixtures:", r.BoundIDs)
	writeFixtureIDs("Partially-bound fixtures:", r.PartialIDs)
	writeFixtureIDs("Unbound fixtures:", r.UnboundIDs)
	writeFixtureIDs("Not-applicable fixtures:", r.NotApplicableIDs)
	if len(r.Rules) > 0 {
		fmt.Fprintln(&b, "\nRules:")
		for _, coverage := range r.Rules {
			fmt.Fprintf(&b, "- %s: rejected=[%s] accepted=[%s] covered_surfaces=[%s] missing_surfaces=[%s] informational_only=%t\n", coverage.Code, strings.Join(coverage.RejectedCases, ", "), strings.Join(coverage.AcceptedControls, ", "), strings.Join(coverage.CoveredSurfaces, ", "), strings.Join(coverage.MissingSurfaces, ", "), coverage.InformationalOnly)
		}
	}
	return strings.TrimRight(b.String(), "\n")
}
