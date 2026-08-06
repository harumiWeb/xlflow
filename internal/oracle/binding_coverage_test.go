package oracle

import (
	"strings"
	"testing"
)

func coverageRejected(id, code string, controls []string, surfaces []string) Case {
	return Case{
		ID:  id,
		VBE: VBEExpectation{Expected: ExpectedRejected},
		Analysis: AnalysisExpectation{
			BindingStatus:    BindingBound,
			RuleCodes:        []string{code},
			NegativeControls: controls,
			ExpectedDiagnostics: []DiagnosticExpectation{{
				Code: code, Surfaces: surfaces,
			}},
		},
	}
}

func coverageAccepted(id, code string, surfaces []string) Case {
	return Case{
		ID:  id,
		VBE: VBEExpectation{Expected: ExpectedAccepted},
		Analysis: AnalysisExpectation{
			BindingStatus: BindingBound,
			RuleCodes:     []string{code},
			ForbiddenDiagnostics: []DiagnosticExpectation{{
				Code: code, Surfaces: surfaces,
			}},
		},
	}
}

func TestValidateBindingCoverageAcceptsCompletePair(t *testing.T) {
	cases := []Case{
		coverageRejected("rejected", "VBA101", []string{"accepted"}, nil),
		coverageAccepted("accepted", "VBA101", nil),
	}
	report, err := ValidateBindingCoverage(cases)
	if err != nil {
		t.Fatalf("ValidateBindingCoverage() error = %v", err)
	}
	if report.CompleteRules != 1 || report.MissingNegativeRules != 0 || report.MissingPositiveRules != 0 {
		t.Fatalf("unexpected report: %+v", report)
	}
	if got := report.Rules[0].CoveredSurfaces; len(got) != 1 || got[0] != "analyze" {
		t.Fatalf("covered surfaces = %v, want [analyze]", got)
	}
}

func TestValidateBindingCoverageRejectsInvalidControls(t *testing.T) {
	tests := []struct {
		name       string
		controls   []string
		extraCases []Case
		errSubstr  string
	}{
		{name: "unknown", controls: []string{"missing"}, errSubstr: "does not exist"},
		{name: "self", controls: []string{"rejected"}, errSubstr: "cannot reference itself"},
		{name: "duplicate", controls: []string{"accepted", "accepted"}, extraCases: []Case{coverageAccepted("accepted", "VBA101", nil)}, errSubstr: "duplicate negative control"},
		{name: "rejected control", controls: []string{"other"}, extraCases: []Case{coverageRejected("other", "VBA101", nil, nil)}, errSubstr: "is not VBE accepted"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cases := []Case{coverageRejected("rejected", "VBA101", tt.controls, nil)}
			cases = append(cases, tt.extraCases...)
			if _, err := ValidateBindingCoverage(cases); err == nil || !strings.Contains(err.Error(), tt.errSubstr) {
				t.Fatalf("ValidateBindingCoverage() error = %v, want substring %q", err, tt.errSubstr)
			}
		})
	}
}

func TestValidateBindingCoverageRequiresCompletePositiveNegativeCoverage(t *testing.T) {
	tests := []struct {
		name      string
		cases     []Case
		errSubstr string
	}{
		{
			name: "missing control",
			cases: []Case{
				coverageRejected("rejected", "VBA101", nil, []string{"analyze"}),
				coverageAccepted("accepted", "VBA101", []string{"analyze"}),
			},
			errSubstr: "missing accepted negative coverage",
		},
		{
			name: "control does not forbid code",
			cases: []Case{
				coverageRejected("rejected", "VBA101", []string{"accepted"}, []string{"analyze"}),
				coverageAccepted("accepted", "VBA102", []string{"analyze"}),
			},
			errSubstr: "missing accepted negative coverage",
		},
		{
			name: "missing positive",
			cases: []Case{
				coverageAccepted("accepted", "VBA101", []string{"analyze"}),
			},
			errSubstr: "missing rejected positive evidence",
		},
		{
			name: "surface gap",
			cases: []Case{
				coverageRejected("rejected", "VBA206", []string{"accepted"}, []string{"analyze", "lsp"}),
				coverageAccepted("accepted", "VBA206", []string{"analyze"}),
			},
			errSubstr: "on surfaces lsp",
		},
		{
			name: "surfaces split across rejected fixtures",
			cases: []Case{
				coverageRejected("rejected-analyze", "VBA206", []string{"analyze-control"}, []string{"analyze", "lsp"}),
				coverageRejected("rejected-lsp", "VBA206", []string{"lsp-control"}, []string{"analyze", "lsp"}),
				coverageAccepted("analyze-control", "VBA206", []string{"analyze"}),
				coverageAccepted("lsp-control", "VBA206", []string{"lsp"}),
			},
			errSubstr: "fixture rejected-analyze rule VBA206",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if _, err := ValidateBindingCoverage(tt.cases); err == nil || !strings.Contains(err.Error(), tt.errSubstr) {
				t.Fatalf("ValidateBindingCoverage() error = %v, want substring %q", err, tt.errSubstr)
			}
		})
	}
}

func TestValidateBindingCoverageRejectsNegativeControlCycle(t *testing.T) {
	a := coverageRejected("a", "VBA101", []string{"b"}, nil)
	b := coverageRejected("b", "VBA101", []string{"a"}, nil)
	if _, err := ValidateBindingCoverage([]Case{a, b}); err == nil || !strings.Contains(err.Error(), "negative control cycle") {
		t.Fatalf("ValidateBindingCoverage() error = %v, want cycle error", err)
	}
}

func TestValidateBindingCoverageRejectsControlsOnWrongFixtureState(t *testing.T) {
	accepted := coverageAccepted("accepted", "VBA101", nil)
	accepted.Analysis.NegativeControls = []string{"rejected"}
	if _, err := ValidateBindingCoverage([]Case{
		coverageRejected("rejected", "VBA101", nil, nil),
		accepted,
	}); err == nil || !strings.Contains(err.Error(), "allowed only on rejected fixtures") {
		t.Fatalf("ValidateBindingCoverage() error = %v, want wrong-state error", err)
	}

	partial := coverageRejected("partial", "VBA101", []string{"accepted"}, nil)
	partial.Analysis.BindingStatus = BindingUnbound
	if _, err := ValidateBindingCoverage([]Case{partial, coverageAccepted("accepted", "VBA101", nil)}); err == nil || !strings.Contains(err.Error(), "require partially-bound or bound status") {
		t.Fatalf("ValidateBindingCoverage() error = %v, want status error", err)
	}

	partialControl := coverageAccepted("partial-control", "VBA101", nil)
	partialControl.Analysis.BindingStatus = BindingPartiallyBound
	if _, err := ValidateBindingCoverage([]Case{
		coverageRejected("partial", "VBA101", []string{"partial-control"}, nil),
		partialControl,
	}); err == nil || !strings.Contains(err.Error(), "must be bound") {
		t.Fatalf("ValidateBindingCoverage() error = %v, want control-bound error", err)
	}
}

func TestBindingCoverageStringIsDeterministic(t *testing.T) {
	report, err := ValidateBindingCoverage([]Case{
		coverageRejected("z-rejected", "VBA101", []string{"a-accepted"}, nil),
		coverageAccepted("a-accepted", "VBA101", nil),
		{ID: "unbound-z", VBE: VBEExpectation{Expected: ExpectedAccepted}, Analysis: AnalysisExpectation{BindingStatus: BindingUnbound}},
		{ID: "unbound-a", VBE: VBEExpectation{Expected: ExpectedRejected}, Analysis: AnalysisExpectation{BindingStatus: BindingUnbound}},
	})
	if err != nil {
		t.Fatal(err)
	}
	output := report.String()
	if strings.Index(output, "- unbound-a") > strings.Index(output, "- unbound-z") {
		t.Fatalf("unbound IDs are not sorted:\n%s", output)
	}
	if !strings.Contains(output, "Rules with complete positive/negative coverage: 1") {
		t.Fatalf("report missing complete rule count:\n%s", output)
	}
}
