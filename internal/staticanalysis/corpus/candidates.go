package corpus

import (
	"fmt"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

// SnapshotReviewCandidateReport contains a bounded, deterministic view of
// current snapshot rows that have not been consumed by true-positive review
// evidence. Rows retain snapshot multiplicity and start-only positions.
type SnapshotReviewCandidateReport struct {
	Rule  string
	Total int
	Rows  []SnapshotDiagnostic
}

// FindSnapshotReviewCandidates finds unreviewed committed snapshot rows for a
// rule. Committed true-positive start-position and multiplicity contracts are
// verified first so a stale ledger cannot silently skew the result. Exact
// end-range and false-positive contracts remain the full corpus run's job.
func FindSnapshotReviewCandidates(reviews []DiagnosticReview, snapshots SnapshotSet, rule string, limit int) (SnapshotReviewCandidateReport, error) {
	metadata, ok := rules.Lookup(rule)
	if !ok {
		return SnapshotReviewCandidateReport{}, fmt.Errorf("unknown diagnostic rule %q", rule)
	}
	if limit < 1 {
		return SnapshotReviewCandidateReport{}, fmt.Errorf("review candidate limit must be positive, got %d", limit)
	}
	rows := make([]SnapshotDiagnostic, 0)
	for _, id := range snapshots.IDs() {
		rows = append(rows, snapshots[id]...)
	}
	sort.SliceStable(rows, func(i, j int) bool { return snapshotReviewCandidateLess(rows[i], rows[j]) })
	used, err := consumeSnapshotTruePositiveReviews(reviews, rows)
	if err != nil {
		return SnapshotReviewCandidateReport{}, err
	}

	report := SnapshotReviewCandidateReport{Rule: metadata.ID}
	for index, row := range rows {
		if used[index] || row.Code != metadata.ID {
			continue
		}
		report.Total++
		if len(report.Rows) < limit {
			report.Rows = append(report.Rows, row)
		}
	}
	return report, nil
}

func snapshotReviewCandidateLess(a, b SnapshotDiagnostic) bool {
	if a.Project != b.Project {
		return a.Project < b.Project
	}
	if a.File != b.File {
		return a.File < b.File
	}
	if surfaceRank(a.Surface) != surfaceRank(b.Surface) {
		return surfaceRank(a.Surface) < surfaceRank(b.Surface)
	}
	if a.Line != b.Line {
		return a.Line < b.Line
	}
	if a.Column != b.Column {
		return a.Column < b.Column
	}
	return a.Severity < b.Severity
}

// FormatSnapshotReviewCandidates renders a stable, grep-friendly candidate
// list and calls out that committed snapshots do not contain end positions.
func FormatSnapshotReviewCandidates(report SnapshotReviewCandidateReport) string {
	var builder strings.Builder
	fmt.Fprintf(&builder, "corpus review candidates: rule=%s unreviewed=%d showing=%d\n", report.Rule, report.Total, len(report.Rows))
	builder.WriteString("project\tfile\tsurface\tseverity\tstart\n")
	for _, row := range report.Rows {
		fmt.Fprintf(&builder, "%s\t%s\t%s\t%s\t%d:%d\n", row.Project, row.File, row.Surface, row.Severity, row.Line, row.Column)
	}
	builder.WriteString("note: snapshot positions are start-only; obtain the full normalized range before editing the review ledger, and use corpus:test for exact TP/FP contracts.\n")
	return builder.String()
}
