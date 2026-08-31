package corpus

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/lint"
)

func reviewAt(classification ReviewClassification, code string, line int) DiagnosticReview {
	return DiagnosticReview{
		SchemaVersion: ReviewSchemaVersion, Project: "self/sample", File: "src/Main.bas", Classification: classification,
		Diagnostic: ReviewedDiagnostic{
			Code: code, Severity: "warning", Surface: "analyze", Count: 1,
			Range: &ReviewedRange{StartLine: line, EndLine: line},
		},
		Rationale: "Reviewed against the source and rule contract.",
	}
}

func diagnosticAt(code string, line int) Diagnostic {
	return Diagnostic{
		Project: "self/sample", File: "src/Main.bas", Surface: SurfaceAnalyze,
		Code: code, Severity: "warning", Line: line, EndLine: line,
	}
}

func TestLoadDiagnosticReviewsStrictJSONLines(t *testing.T) {
	path := filepath.Join(t.TempDir(), "diagnostics.jsonl")
	data := `{"schema_version":1,"project":"self/sample","file":"src/Main.bas","classification":"true-positive","diagnostic":{"code":"VBA225","severity":"warning","surface":"analyze","range":{"start_line":3,"start_column":0,"end_line":3,"end_column":0},"count":1},"rationale":"Reviewed against the source and rule contract."}` + "\n"
	if err := os.WriteFile(path, []byte(data), 0o644); err != nil {
		t.Fatal(err)
	}
	reviews, err := LoadDiagnosticReviews(path)
	if err != nil {
		t.Fatal(err)
	}
	if len(reviews) != 1 || reviews[0].Diagnostic.Code != "VBA225" {
		t.Fatalf("reviews = %#v", reviews)
	}
	if err := os.WriteFile(path, []byte(strings.Replace(data, `,"rationale"`, `,"extra":true,"rationale"`, 1)), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadDiagnosticReviews(path); err == nil || !strings.Contains(err.Error(), "unknown field") {
		t.Fatalf("unknown field error = %v", err)
	}
	if err := os.WriteFile(path, []byte(" "+data), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadDiagnosticReviews(path); err == nil || !strings.Contains(err.Error(), "canonical JSON") {
		t.Fatalf("leading whitespace error = %v", err)
	}
	if err := os.WriteFile(path, []byte(data+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadDiagnosticReviews(path); err == nil || !strings.Contains(err.Error(), "is blank") {
		t.Fatalf("blank line error = %v", err)
	}
	trailing := strings.TrimSuffix(data, "\n") + `{}` + "\n"
	if err := os.WriteFile(path, []byte(trailing), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadDiagnosticReviews(path); err == nil || !strings.Contains(err.Error(), "trailing JSON") {
		t.Fatalf("trailing JSON error = %v", err)
	}
	unsupported := strings.Replace(data, `"schema_version":1`, `"schema_version":2`, 1)
	if err := os.WriteFile(path, []byte(unsupported), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadDiagnosticReviews(path); err == nil || !strings.Contains(err.Error(), "unsupported schema_version") {
		t.Fatalf("schema version error = %v", err)
	}
}

func TestLoadDiagnosticReviewsRejectsConflictingIdentityAndOrder(t *testing.T) {
	path := filepath.Join(t.TempDir(), "diagnostics.jsonl")
	first := `{"schema_version":1,"project":"self/sample","file":"src/Main.bas","classification":"true-positive","diagnostic":{"code":"VBA225","severity":"warning","surface":"analyze","range":{"start_line":3,"start_column":0,"end_line":3,"end_column":0},"count":1},"rationale":"first"}`
	conflict := `{"schema_version":1,"project":"self/sample","file":"src/Main.bas","classification":"false-positive","diagnostic":{"code":"VBA225","severity":"warning","surface":"analyze","range":{"start_line":3,"start_column":0,"end_line":3,"end_column":0},"count":1},"rationale":"conflict","regression_exception":"not yet reducible"}`
	if err := os.WriteFile(path, []byte(first+"\n"+conflict+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadDiagnosticReviews(path); err == nil || !strings.Contains(err.Error(), "repeats an earlier diagnostic identity") {
		t.Fatalf("conflicting identity error = %v", err)
	}

	before := strings.Replace(first, `"project":"self/sample"`, `"project":"self/aaa"`, 1)
	if err := os.WriteFile(path, []byte(first+"\n"+before+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadDiagnosticReviews(path); err == nil || !strings.Contains(err.Error(), "canonical order") {
		t.Fatalf("order error = %v", err)
	}
}

func TestCommittedDiagnosticReviews(t *testing.T) {
	repoRoot, err := filepath.Abs(filepath.Join("..", "..", ".."))
	if err != nil {
		t.Fatal(err)
	}
	corpusRoot := filepath.Join(repoRoot, "testdata", "static-analysis-corpus")
	reviews, err := LoadDiagnosticReviews(filepath.Join(corpusRoot, "reviews", "diagnostics.jsonl"))
	if err != nil {
		t.Fatal(err)
	}
	if len(reviews) != 1515 {
		t.Fatalf("committed reviews = %d, want 1515", len(reviews))
	}
	if err := ValidateReviewSources(repoRoot, corpusRoot, reviews); err != nil {
		t.Fatal(err)
	}
	invalid := append([]DiagnosticReview(nil), reviews...)
	target := -1
	for index := range invalid {
		if invalid[index].RegressionTest != "" {
			target = index
			break
		}
	}
	if target < 0 {
		t.Fatal("committed reviews contain no regression_test")
	}
	invalid[target].RegressionTest = "internal/vba/dataflow/analysis_test.go::TestDoesNotExist"
	if err := ValidateReviewSources(repoRoot, corpusRoot, invalid); err == nil || !strings.Contains(err.Error(), "does not name a function") {
		t.Fatalf("missing regression function error = %v", err)
	}
}

func TestCommittedCorpusReviewMetrics(t *testing.T) {
	repoRoot, err := filepath.Abs(filepath.Join("..", "..", ".."))
	if err != nil {
		t.Fatal(err)
	}
	corpusRoot := filepath.Join(repoRoot, "testdata", "static-analysis-corpus")
	reviews, err := LoadDiagnosticReviews(filepath.Join(corpusRoot, "reviews", "diagnostics.jsonl"))
	if err != nil {
		t.Fatal(err)
	}
	if err := ValidateReviewSources(repoRoot, corpusRoot, reviews); err != nil {
		t.Fatal(err)
	}
	snapshotRoot := filepath.Join(corpusRoot, "snapshots")
	ids, err := DiscoverSnapshotIDs(snapshotRoot)
	if err != nil {
		t.Fatal(err)
	}
	snapshots, err := LoadSnapshotSet(snapshotRoot, ids)
	if err != nil {
		t.Fatal(err)
	}
	metrics, err := EvaluateSnapshotReviews(reviews, snapshots)
	if err != nil {
		t.Fatal(err)
	}
	if metrics.Reviewed != 1697 || metrics.TP != 729 || metrics.FP != 968 {
		t.Fatalf("committed review metrics = %#v, want Reviewed=1697 TP=729 FP=968", metrics)
	}
	// Guard the direction of the corpus review ledger, not only its exact size.
	if precision := float64(metrics.TP) / float64(metrics.TP+metrics.FP); precision < 0.429 {
		t.Fatalf("committed review precision = %.3f, want >= 0.429", precision)
	}
	t.Log(FormatReviewMetrics(metrics))
}

func TestEvaluateDiagnosticReviewsEnforcesTrueAndFalsePositiveContracts(t *testing.T) {
	tp := reviewAt(ReviewTruePositive, "VBA225", 3)
	fp := reviewAt(ReviewFalsePositive, "VBA224", 9)
	fp.RegressionTest = "internal/vba/dataflow/analysis_test.go::TestAnalyzeProcedureFindingsDoNotDependOnWorklistRank"
	report := Report{Diagnostics: []Diagnostic{diagnosticAt("VBA225", 3), diagnosticAt("VBA201", 12)}}

	metrics, err := EvaluateDiagnosticReviews([]DiagnosticReview{tp, fp}, report)
	if err != nil {
		t.Fatal(err)
	}
	if metrics.Reviewed != 2 || metrics.TP != 1 || metrics.FP != 1 || metrics.Unreviewed != 1 {
		t.Fatalf("metrics = %#v", metrics)
	}

	if _, err := EvaluateDiagnosticReviews([]DiagnosticReview{tp}, Report{}); err == nil || !strings.Contains(err.Error(), "expected diagnostic VBA225") {
		t.Fatalf("missing TP error = %v", err)
	}
	if _, err := EvaluateDiagnosticReviews([]DiagnosticReview{fp}, Report{Diagnostics: []Diagnostic{diagnosticAt("VBA224", 9)}}); err == nil || !strings.Contains(err.Error(), "forbidden diagnostic VBA224") {
		t.Fatalf("reappearing FP error = %v", err)
	}
}

func TestEvaluateDiagnosticReviewsAllowsCollidingBaselineOccurrences(t *testing.T) {
	fp := reviewAt(ReviewFalsePositive, "VBA224", 9)
	fp.Diagnostic.AllowedOccurrences = 2
	fp.RegressionTest = "internal/vba/dataflow/analysis_test.go::TestAnalyzeProcedureFindingsDoNotDependOnWorklistRank"
	baseline := Report{Diagnostics: []Diagnostic{diagnosticAt("VBA224", 9), diagnosticAt("VBA224", 9)}}
	metrics, err := EvaluateDiagnosticReviews([]DiagnosticReview{fp}, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if metrics.FP != 1 || metrics.Unreviewed != 2 {
		t.Fatalf("metrics = %#v", metrics)
	}
	baseline.Diagnostics = append(baseline.Diagnostics, diagnosticAt("VBA224", 9))
	if _, err := EvaluateDiagnosticReviews([]DiagnosticReview{fp}, baseline); err == nil || !strings.Contains(err.Error(), "exceeded allowed occurrences (3 > 2)") {
		t.Fatalf("collision threshold error = %v", err)
	}
}

func TestEvaluateSnapshotReviewsUsesCommittedStartIdentity(t *testing.T) {
	tp := reviewAt(ReviewTruePositive, "VBA225", 3)
	tp.Diagnostic.Range.EndLine = 4
	fp := reviewAt(ReviewFalsePositive, "VBA224", 9)
	fp.RegressionTest = "internal/vba/dataflow/analysis_test.go::TestAnalyzeProcedureFindingsDoNotDependOnWorklistRank"
	id := SnapshotID{Project: "self/sample", Surface: SurfaceAnalyze}
	set := SnapshotSet{id: {
		{Project: "self/sample", File: "src/Main.bas", Surface: SurfaceAnalyze, Code: "VBA225", Severity: "warning", Line: 3},
		{Project: "self/sample", File: "src/Main.bas", Surface: SurfaceAnalyze, Code: "VBA201", Severity: "warning", Line: 12},
	}}
	metrics, err := EvaluateSnapshotReviews([]DiagnosticReview{tp, fp}, set)
	if err != nil {
		t.Fatal(err)
	}
	if metrics.Reviewed != 2 || metrics.Unreviewed != 1 || len(metrics.Rules) != 3 {
		t.Fatalf("metrics = %#v", metrics)
	}
	set[id] = append(set[id], SnapshotDiagnostic{Project: "self/sample", File: "src/Main.bas", Surface: SurfaceAnalyze, Code: "VBA224", Severity: "warning", Line: 9})
	metrics, err = EvaluateSnapshotReviews([]DiagnosticReview{tp, fp}, set)
	if err != nil || metrics.Unreviewed != 2 {
		t.Fatalf("start-only forbidden snapshot must remain unreviewed: metrics=%#v err=%v", metrics, err)
	}
	set[id] = append(set[id], SnapshotDiagnostic{Project: "self/sample", File: "src/Main.bas", Surface: SurfaceAnalyze, Code: "VBA225", Severity: "warning", Line: 3})
	metrics, err = EvaluateSnapshotReviews([]DiagnosticReview{tp, fp}, set)
	if err != nil || metrics.TP != 1 || metrics.Unreviewed != 3 {
		t.Fatalf("extra TP occurrence must remain unreviewed: metrics=%#v err=%v", metrics, err)
	}

	invalid := tp
	invalid.Diagnostic.Count = 0
	if _, err := EvaluateSnapshotReviews([]DiagnosticReview{invalid}, set); err == nil || !strings.Contains(err.Error(), "review 1") {
		t.Fatalf("invalid direct review error = %v", err)
	}
}

func TestFormatReviewMetricsIsDeterministicAndReviewedOnly(t *testing.T) {
	metrics := ReviewMetrics{
		Reviewed: 3, TP: 2, FP: 1, Unreviewed: 7,
		Rules: []RuleReviewMetrics{
			{Rule: "VBA225", Reviewed: 2, TP: 2, Precision: 1, HasPrecision: true},
			{Rule: "VBA224", Reviewed: 1, FP: 1, HasPrecision: true},
		},
	}
	want := "corpus reviews: reviewed=3 tp=2 fp=1 unreviewed=7\n" +
		"rule reviewed unreviewed tp fp precision\n" +
		"VBA224 1 0 0 1 0.0%\n" +
		"VBA225 2 0 2 0 100.0%\n"
	if got := FormatReviewMetrics(metrics); got != want {
		t.Fatalf("metrics format = %q, want %q", got, want)
	}
}

func TestDiagnosticRangeNormalization(t *testing.T) {
	got := normalizeAnalyzeDiagnostics("self/sample", []analyze.Finding{{
		Code: "VBA225", Severity: "warning", File: "src/Main.bas", Line: 4, Column: 2, EndLine: 6, EndColumn: 8,
	}})
	if len(got) != 1 || got[0].EndLine != 6 || got[0].EndColumn != 8 {
		t.Fatalf("analyze diagnostic range = %#v", got)
	}
	lintGot := normalizeLintDiagnostics("self/sample", []lint.Issue{{Code: "VB001", Severity: "warning", File: "src/Main.bas", Line: 7, Column: 3}})
	if len(lintGot) != 1 || lintGot[0].EndLine != 7 || lintGot[0].EndColumn != 3 {
		t.Fatalf("lint diagnostic range = %#v", lintGot)
	}
}

func TestDiagnosticReviewRejectsReversedSingleLineRange(t *testing.T) {
	review := reviewAt(ReviewTruePositive, "VBA225", 3)
	review.Diagnostic.Range.StartColumn = 8
	review.Diagnostic.Range.EndColumn = 2
	if err := validateDiagnosticReview(review); err == nil || !strings.Contains(err.Error(), "incoherent source range") {
		t.Fatalf("reversed range error = %v", err)
	}
}
