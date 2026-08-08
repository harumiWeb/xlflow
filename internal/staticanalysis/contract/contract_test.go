package contract

import (
	"reflect"
	"strings"
	"testing"
)

func TestCheckExpectedForbiddenAndExactRange(t *testing.T) {
	wantRange := &Range{StartLine: 4, StartColumn: 2, EndLine: 4, EndColumn: 8}
	expect := ExpectationSet{
		ExpectedDiagnostics:  []DiagnosticExpectation{{Code: "VBA001", Severity: "warning", Range: wantRange}},
		ForbiddenDiagnostics: []DiagnosticExpectation{{Code: "VBA002"}},
	}
	actual := []Diagnostic{{Code: "VBA001", Severity: "warning", Range: &Range{StartLine: 4, StartColumn: 2, EndLine: 4, EndColumn: 8}}}
	if err := Check(expect, actual); err != nil {
		t.Fatal(err)
	}

	actual[0].Range.EndColumn = 9
	if err := Check(expect, actual); err == nil || !strings.Contains(err.Error(), "was not emitted") {
		t.Fatalf("err=%v, want exact-range mismatch", err)
	}
	actual = append(actual, Diagnostic{Code: "VBA002"})
	if err := Check(expect, actual); err == nil || !strings.Contains(err.Error(), "forbidden") {
		t.Fatalf("err=%v, want forbidden diagnostic failure", err)
	}
}

func TestCheckRequiresDistinctDiagnosticsForDuplicateExpectations(t *testing.T) {
	expect := ExpectationSet{ExpectedDiagnostics: []DiagnosticExpectation{{Code: "VBA001"}, {Code: "VBA001"}}}
	if err := Check(expect, []Diagnostic{{Code: "VBA001"}}); err == nil {
		t.Fatal("expected one actual diagnostic not to satisfy two expectations")
	}
}

func TestCheckReassignsBroadMatchToSatisfySpecificExpectation(t *testing.T) {
	exactRange := &Range{StartLine: 4, StartColumn: 1, EndLine: 4, EndColumn: 8}
	expect := ExpectationSet{ExpectedDiagnostics: []DiagnosticExpectation{
		{Code: "VBA101"},
		{Code: "VBA101", Severity: "warning", Range: exactRange, Surfaces: []string{"analyze"}},
	}}
	actual := []Diagnostic{
		{Code: "VBA101", Severity: "warning", Range: exactRange, Surface: "analyze"},
		{Code: "VBA101", Severity: "warning", Range: &Range{StartLine: 9}, Surface: "analyze"},
	}
	if err := Check(expect, actual); err != nil {
		t.Fatal(err)
	}
}

func TestCheckProjections(t *testing.T) {
	expect := ExpectationSet{ExpectedDiagnostics: []DiagnosticExpectation{{Code: "VBA001", Severity: "warning", Surfaces: []string{"analyze", "lsp"}}}}
	projections := map[string][]Diagnostic{
		"analyze": {{Code: "VBA001", Severity: "warning"}},
		"lsp":     {{Code: "VBA001", Severity: "warning"}},
	}
	if err := CheckProjections(expect, projections); err != nil {
		t.Fatal(err)
	}
	projections["lsp"][0].Severity = "error"
	if err := CheckProjections(expect, projections); err == nil {
		t.Fatal("expected projection severity mismatch")
	}
	projections["lsp"] = []Diagnostic{{Code: "VBA001", Severity: "warning"}, {Code: "VBA001", Severity: "warning"}}
	if err := CheckProjections(expect, projections); err == nil {
		t.Fatal("expected duplicate projection diagnostic")
	}

	surfaceMismatch := map[string][]Diagnostic{
		"lsp":     {{Code: "VBA101", Severity: "warning", Surface: "analyze"}},
		"analyze": {{Code: "VBA101", Severity: "warning", Surface: "lsp"}},
	}
	if err := CheckProjections(expect, surfaceMismatch); err == nil || !strings.Contains(err.Error(), `in "analyze" projection`) {
		t.Fatalf("surface mismatch error = %v", err)
	}

	projections = map[string][]Diagnostic{
		"analyze": {{Code: "VBA001", Severity: "warning", Range: &Range{StartLine: 1, EndLine: 1}}},
		"lsp":     {{Code: "VBA001", Severity: "warning", Range: &Range{StartLine: 2, EndLine: 2}}},
	}
	if err := CheckProjections(expect, projections); err == nil || !strings.Contains(err.Error(), "differs across projections") {
		t.Fatalf("range mismatch error = %v", err)
	}
}

func TestNormalizeIsDeterministicAndNonMutating(t *testing.T) {
	input := []Diagnostic{
		{Code: "VBA002", Severity: "warning"},
		{Code: "VBA001", Severity: "warning", Surface: "lsp", Range: &Range{StartLine: 2}},
		{Code: "VBA001", Severity: "warning", Surface: "analyze", Range: &Range{StartLine: 1}},
	}
	got := Normalize(input)
	want := []Diagnostic{input[2], input[1], input[0]}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("Normalize()=%+v, want %+v", got, want)
	}
	if input[0].Code != "VBA002" {
		t.Fatalf("Normalize mutated input: %+v", input)
	}
}

func TestValidateUsesRuleRegistry(t *testing.T) {
	tests := []struct {
		name string
		item DiagnosticExpectation
		want string
	}{
		{name: "valid", item: DiagnosticExpectation{Code: "VBA101", Severity: "warning", Surfaces: []string{"analyze"}}},
		{name: "canonical", item: DiagnosticExpectation{Code: "vba101"}, want: "canonical registry ID"},
		{name: "unknown", item: DiagnosticExpectation{Code: "VBA999"}, want: "not in the static-analysis rule registry"},
		{name: "severity", item: DiagnosticExpectation{Code: "VBA101", Severity: "error"}, want: "unsupported severity"},
		{name: "surface", item: DiagnosticExpectation{Code: "VBA101", Surfaces: []string{"lint"}}, want: "unsupported surface"},
		{name: "duplicate surface", item: DiagnosticExpectation{Code: "VBA101", Surfaces: []string{"analyze", "analyze"}}, want: "repeats surface"},
		{name: "negative range", item: DiagnosticExpectation{Code: "VBA101", Range: &Range{StartLine: -1}}, want: "negative source range"},
		{name: "reversed lines", item: DiagnosticExpectation{Code: "VBA101", Range: &Range{StartLine: 4, EndLine: 3}}, want: "incoherent source range"},
		{name: "reversed columns", item: DiagnosticExpectation{Code: "VBA101", Range: &Range{StartLine: 4, StartColumn: 8, EndLine: 4, EndColumn: 2}}, want: "incoherent source range"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := ValidateExpectations([]DiagnosticExpectation{tt.item})
			if tt.want == "" && err != nil {
				t.Fatal(err)
			}
			if tt.want != "" && (err == nil || !strings.Contains(err.Error(), tt.want)) {
				t.Fatalf("err=%v, want substring %q", err, tt.want)
			}
		})
	}
}

func TestValidateChecksBothExpectationSets(t *testing.T) {
	expect := ExpectationSet{ForbiddenDiagnostics: []DiagnosticExpectation{{Code: "vba101"}}}
	if err := Validate(expect); err == nil || !strings.Contains(err.Error(), "forbidden diagnostics") {
		t.Fatalf("Validate error = %v", err)
	}
}

func TestMatchesUsesSurfaceSetMembership(t *testing.T) {
	expect := DiagnosticExpectation{Code: "VBA101", Surfaces: []string{"analyze", "lsp"}}
	if !Matches(expect, Diagnostic{Code: "VBA101", Surface: "lsp"}) {
		t.Fatal("expected lsp diagnostic to match surface set")
	}
	if Matches(expect, Diagnostic{Code: "VBA101", Surface: "lint"}) {
		t.Fatal("unexpected lint diagnostic match")
	}
}
