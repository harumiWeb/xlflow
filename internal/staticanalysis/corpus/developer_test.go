package corpus

import (
	"path/filepath"
	"strings"
	"testing"
)

func TestLoadSnapshotSubsetIgnoresUnselectedCommittedFiles(t *testing.T) {
	root := filepath.Join(t.TempDir(), "snapshots")
	first := SnapshotID{Project: "self/first", Surface: SurfaceAnalyze}
	second := SnapshotID{Project: "self/second", Surface: SurfaceLint}
	set := SnapshotSet{
		first:  {{Project: first.Project, Surface: first.Surface, File: "src/Main.bas", Code: "VBA225", Severity: "warning", Line: 3}},
		second: {},
	}
	if err := WriteSnapshotSet(root, set); err != nil {
		t.Fatal(err)
	}
	got, err := LoadSnapshotSubset(root, []SnapshotID{first})
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 1 || len(got[first]) != 1 {
		t.Fatalf("subset = %#v", got)
	}
}

func TestResolveReviewDetailsRetainsRangesAndMultiplicity(t *testing.T) {
	row := SnapshotDiagnostic{Project: "self/sample", Surface: SurfaceAnalyze, File: "src/Main.bas", Code: "VBA225", Severity: "warning", Line: 3, Column: 2}
	candidates := SnapshotReviewCandidateReport{Rule: "VBA225", Rows: []SnapshotDiagnostic{row, row}}
	report := Report{Diagnostics: []Diagnostic{
		{Project: row.Project, Surface: row.Surface, File: row.File, Code: row.Code, Severity: row.Severity, Line: 3, Column: 2, EndLine: 3, EndColumn: 8},
		{Project: row.Project, Surface: row.Surface, File: row.File, Code: row.Code, Severity: row.Severity, Line: 3, Column: 2, EndLine: 4, EndColumn: 1},
	}}
	details, err := ResolveReviewDetails(candidates, report)
	if err != nil {
		t.Fatal(err)
	}
	if len(details) != 2 || details[0].EndColumn != 8 || details[1].EndLine != 4 {
		t.Fatalf("details = %#v", details)
	}
}

func TestBuildReviewDraftsIsCanonicalAndRequiresHumanClassificationFields(t *testing.T) {
	detail := ReviewDetail{Project: "self/sample", File: "src/Main.bas", Surface: SurfaceAnalyze, Code: "VBA225", Severity: "warning", Line: 3, EndLine: 3}
	drafts, err := BuildReviewDrafts([]ReviewDetail{detail, detail}, ReviewFalsePositive, "Receiver resolution was incorrect.", "internal/analyze/analyzer_test.go::TestFocused", "")
	if err != nil {
		t.Fatal(err)
	}
	if len(drafts) != 1 || drafts[0].Diagnostic.Count != 2 {
		t.Fatalf("drafts = %#v", drafts)
	}
	encoded, err := EncodeReviewDrafts(drafts)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"count":2`) || !strings.HasSuffix(string(encoded), "\n") {
		t.Fatalf("encoded = %q", encoded)
	}
	if _, err := BuildReviewDrafts([]ReviewDetail{detail}, ReviewFalsePositive, "reason", "", ""); err == nil {
		t.Fatal("false-positive draft without regression evidence succeeded")
	}
	distinctRange := detail
	distinctRange.EndLine = 4
	drafts, err = BuildReviewDrafts([]ReviewDetail{detail, distinctRange}, ReviewTruePositive, "Reviewed against the rule contract.", "", "")
	if err != nil {
		t.Fatal(err)
	}
	if len(drafts) != 2 {
		t.Fatalf("distinct ranges were grouped together: %#v", drafts)
	}
}

func TestFilterReportAndSnapshotsKeepEmptySurfaces(t *testing.T) {
	id := SnapshotID{Project: "self/sample", Surface: SurfaceAnalyze}
	report := Report{Surfaces: []SurfaceResult{{Project: id.Project, Surface: id.Surface, Diagnostics: []Diagnostic{{Project: id.Project, Surface: id.Surface, Code: "VBA224"}}}}}
	filteredReport := FilterReport(report, "VBA225")
	if len(filteredReport.Surfaces) != 1 || len(filteredReport.Surfaces[0].Diagnostics) != 0 {
		t.Fatalf("report = %#v", filteredReport)
	}
	filteredSnapshots := FilterSnapshots(SnapshotSet{id: {{Project: id.Project, Surface: id.Surface, File: "src/Main.bas", Code: "VBA224", Severity: "warning", Line: 1}}}, "VBA225")
	if rows, ok := filteredSnapshots[id]; !ok || len(rows) != 0 {
		t.Fatalf("snapshots = %#v", filteredSnapshots)
	}
}
