package corpus

import (
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

func snapshotAt(project, file string, surface Surface, code string, line int) SnapshotDiagnostic {
	return SnapshotDiagnostic{
		Project: project, File: file, Surface: surface,
		Code: code, Severity: "warning", Line: line,
	}
}

func TestFindSnapshotReviewCandidatesReturnsOnlyUnreviewedRuleRows(t *testing.T) {
	row := snapshotAt("self/sample", "src/Main.bas", SurfaceAnalyze, "VBA225", 3)
	otherProject := snapshotAt("self/alpha", "src/Alpha.bas", SurfaceLint, "VBA225", 9)
	otherRule := snapshotAt("self/sample", "src/Main.bas", SurfaceAnalyze, "VBA224", 12)
	snapshots := SnapshotSet{
		{Project: "self/sample", Surface: SurfaceAnalyze}: {row, row, otherRule},
		{Project: "self/alpha", Surface: SurfaceLint}:     {otherProject},
	}
	tp := reviewAt(ReviewTruePositive, "VBA225", 3)

	report, err := FindSnapshotReviewCandidates([]DiagnosticReview{tp}, snapshots, "vba225", 20)
	if err != nil {
		t.Fatal(err)
	}
	if report.Rule != "VBA225" || report.Total != 2 || len(report.Rows) != 2 {
		t.Fatalf("report = %#v", report)
	}
	if report.Rows[0] != otherProject || report.Rows[1] != row {
		t.Fatalf("rows = %#v", report.Rows)
	}
}

func TestFindSnapshotReviewCandidatesPreservesFalsePositiveCollisionsAndLimit(t *testing.T) {
	first := snapshotAt("self/sample", "src/Main.bas", SurfaceAnalyze, "VBA224", 3)
	second := snapshotAt("self/sample", "src/Main.bas", SurfaceAnalyze, "VBA224", 7)
	fp := reviewAt(ReviewFalsePositive, "VBA224", 3)
	fp.Diagnostic.AllowedOccurrences = 1
	snapshots := SnapshotSet{
		{Project: "self/sample", Surface: SurfaceAnalyze}: {second, first},
	}

	report, err := FindSnapshotReviewCandidates([]DiagnosticReview{fp}, snapshots, "VBA224", 1)
	if err != nil {
		t.Fatal(err)
	}
	if report.Total != 2 || len(report.Rows) != 1 || report.Rows[0] != first {
		t.Fatalf("report = %#v", report)
	}
}

func TestFindSnapshotReviewCandidatesRejectsInvalidInputAndBrokenEvidence(t *testing.T) {
	if _, err := FindSnapshotReviewCandidates(nil, nil, "VBA999", 20); err == nil || !strings.Contains(err.Error(), "unknown diagnostic rule") {
		t.Fatalf("unknown rule error = %v", err)
	}
	if _, err := FindSnapshotReviewCandidates(nil, nil, "VBA225", 0); err == nil || !strings.Contains(err.Error(), "positive") {
		t.Fatalf("invalid limit error = %v", err)
	}
	tp := reviewAt(ReviewTruePositive, "VBA225", 3)
	if _, err := FindSnapshotReviewCandidates([]DiagnosticReview{tp}, SnapshotSet{}, "VBA225", 20); err == nil || !strings.Contains(err.Error(), "committed review contract violations") {
		t.Fatalf("broken evidence error = %v", err)
	}
}

func TestFormatSnapshotReviewCandidates(t *testing.T) {
	report := SnapshotReviewCandidateReport{
		Rule:  "VBA225",
		Total: 2,
		Rows: []SnapshotDiagnostic{
			snapshotAt("self/alpha", "src/Alpha.bas", SurfaceLint, "VBA225", 9),
		},
	}
	want := "corpus review candidates: rule=VBA225 unreviewed=2 showing=1\n" +
		"project\tfile\tsurface\tseverity\tstart\n" +
		"self/alpha\tsrc/Alpha.bas\tlint\twarning\t9:0\n" +
		"note: snapshot positions are start-only; obtain the full normalized range before editing the review ledger, and use corpus:test for exact TP/FP contracts.\n"
	if got := FormatSnapshotReviewCandidates(report); got != want {
		t.Fatalf("format = %q, want %q", got, want)
	}
}

func TestCommittedCorpusReviewCandidates(t *testing.T) {
	if os.Getenv("XLFLOW_PRINT_CORPUS_REVIEW_CANDIDATES") != "1" {
		t.Skip("set XLFLOW_PRINT_CORPUS_REVIEW_CANDIDATES=1 through task corpus:review-candidates")
	}
	args := strings.Fields(os.Getenv("XLFLOW_CORPUS_REVIEW_CANDIDATES_ARGS"))
	if len(args) < 1 || len(args) > 2 {
		t.Fatal("usage: task corpus:review-candidates -- <RULE> [LIMIT]")
	}
	limit := 20
	if len(args) == 2 {
		parsed, err := strconv.Atoi(args[1])
		if err != nil {
			t.Fatalf("candidate limit %q must be a positive integer", args[1])
		}
		limit = parsed
	}

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
	ids, err := discoverCommittedSnapshotIDs(snapshotRoot)
	if err != nil {
		t.Fatal(err)
	}
	snapshots, err := LoadSnapshotSet(snapshotRoot, ids)
	if err != nil {
		t.Fatal(err)
	}
	report, err := FindSnapshotReviewCandidates(reviews, snapshots, args[0], limit)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("\n" + FormatSnapshotReviewCandidates(report))
}
