package corpus

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/typedb"
)

func snapshotTestRow(project string, surface Surface, file string, line, column int, code string) SnapshotDiagnostic {
	return SnapshotDiagnostic{
		Project: project, Surface: surface, File: file, Line: line, Column: column,
		Code: code, Severity: "warning",
	}
}

func TestSnapshotSetRoundTripAndEmptySurface(t *testing.T) {
	root := filepath.Join(t.TempDir(), "snapshots")
	id := SnapshotID{Project: "self/example", Surface: SurfaceLint}
	set := SnapshotSet{
		id: {
			snapshotTestRow(id.Project, id.Surface, "src/Main.bas", 4, 2, "VB001"),
			snapshotTestRow(id.Project, id.Surface, "src/Main.bas", 4, 2, "VB001"),
		},
		{Project: "self/example", Surface: SurfaceAnalyze}: {},
	}
	if err := WriteSnapshotSet(root, set); err != nil {
		t.Fatal(err)
	}
	got, err := LoadSnapshotSet(root, set.IDs())
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, set) {
		t.Fatalf("round-trip set = %#v, want %#v", got, set)
	}
	if info, err := os.Stat(filepath.Join(root, "analyze", "self", "example.jsonl")); err != nil {
		t.Fatal(err)
	} else if info.Size() != 0 {
		t.Fatalf("empty snapshot size = %d, want 0", info.Size())
	}
}

func TestSnapshotSetRejectsMissingUnexpectedAndNonCanonicalRows(t *testing.T) {
	root := filepath.Join(t.TempDir(), "snapshots")
	id := SnapshotID{Project: "self/example", Surface: SurfaceLint}
	set := SnapshotSet{id: {snapshotTestRow(id.Project, id.Surface, "Main.bas", 1, 0, "VB001")}}
	if err := WriteSnapshotSet(root, set); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadSnapshotSet(root, append(set.IDs(), SnapshotID{Project: "self/missing", Surface: SurfaceAnalyze})); err == nil {
		t.Fatal("missing expected snapshot was accepted")
	}
	if err := os.WriteFile(filepath.Join(root, "lint", "stale.jsonl"), []byte{}, 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadSnapshotSet(root, set.IDs()); err == nil || !strings.Contains(err.Error(), "unexpected file") {
		t.Fatalf("stale snapshot error = %v", err)
	}
	if err := os.Remove(filepath.Join(root, "lint", "stale.jsonl")); err != nil {
		t.Fatal(err)
	}
	rowPath := filepath.Join(root, "lint", "self", "example.jsonl")
	if err := os.WriteFile(rowPath, []byte(`{"project":"self/other","file":"Main.bas","surface":"lint","code":"VB001","severity":"warning","line":1,"column":0}`+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadSnapshotSet(root, set.IDs()); err == nil || !strings.Contains(err.Error(), "does not match") {
		t.Fatalf("project mismatch error = %v", err)
	}
	if err := os.WriteFile(rowPath, []byte(`{"surface":"lint","project":"self/example","file":"Main.bas","code":"VB001","severity":"warning","line":1,"column":0}`+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadSnapshotSet(root, set.IDs()); err == nil || !strings.Contains(err.Error(), "canonical") {
		t.Fatalf("non-canonical row error = %v", err)
	}
	if err := os.WriteFile(rowPath, []byte(`{"project":"self/example","file":"Main.bas","surface":"lint","code":"VB001","severity":"warning","line":1,"column":0,"extra":true}`+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadSnapshotSet(root, set.IDs()); err == nil || !strings.Contains(err.Error(), "unknown field") {
		t.Fatalf("unknown field error = %v", err)
	}
}

func TestCompareSnapshotSetsPreservesMultiplicity(t *testing.T) {
	id := SnapshotID{Project: "self/example", Surface: SurfaceAnalyze}
	a := snapshotTestRow(id.Project, id.Surface, "Main.bas", 2, 1, "VBA201")
	b := snapshotTestRow(id.Project, id.Surface, "Main.bas", 3, 1, "VBA201")
	want := SnapshotSet{id: {a, a, b}}
	got := SnapshotSet{id: {a, b, b}}
	diff := CompareSnapshotSets(want, got)
	if !reflect.DeepEqual(diff.Added, []SnapshotDiagnostic{b}) || !reflect.DeepEqual(diff.Removed, []SnapshotDiagnostic{a}) {
		t.Fatalf("diff = %#v, want added [%#v], removed [%#v]", diff, b, a)
	}
}

func TestSnapshotDiffByRuleAndFormatIsDeterministic(t *testing.T) {
	diff := SnapshotDiff{
		Added: []SnapshotDiagnostic{
			{Project: "self/z", File: "src/Z.bas", Surface: SurfaceAnalyze, Code: "VBA209", Severity: "warning", Line: 3},
			{Project: "self/a", File: `src\A.bas`, Surface: SurfaceLint, Code: "VBA209", Severity: "error", Line: 2, Column: 4},
			{Project: "self/a", File: `src\A.bas`, Surface: SurfaceLint, Code: "VBA209", Severity: "error", Line: 2, Column: 4},
		},
		Removed: []SnapshotDiagnostic{
			{Project: "third_party/vba", File: "WebClient.cls", Surface: SurfaceAnalyze, Code: "VBA209", Severity: "warning", Line: 1},
			{Project: "self/x", File: "src/X.bas", Surface: SurfaceLint, Code: "VBA206", Severity: "warning", Line: 5, Column: 1},
		},
	}

	groups := diff.ByRule()
	if len(groups) != 2 || groups[0].Code != "VBA206" || groups[1].Code != "VBA209" {
		t.Fatalf("rule groups = %#v, want VBA206 then VBA209", groups)
	}
	if len(groups[1].Added) != 3 || len(groups[1].Removed) != 1 {
		t.Fatalf("VBA209 group = %#v, want three additions and one removal", groups[1])
	}

	want := "Real-world static-analysis corpus changed\n\n" +
		"VBA206 +0 -1\n" +
		"  - self/x/src/X.bas:5:1 [lint warning]\n\n" +
		"VBA209 +3 -1\n" +
		"  + self/a/src/A.bas:2:4 [lint error]\n" +
		"  + self/a/src/A.bas:2:4 [lint error]\n" +
		"  + self/z/src/Z.bas:3 [analyze warning]\n" +
		"  - third_party/vba/WebClient.cls:1 [analyze warning]\n"
	if got := FormatSnapshotDiff(diff); got != want {
		t.Fatalf("formatted diff = %q, want %q", got, want)
	}

	shuffled := SnapshotDiff{
		Added:   []SnapshotDiagnostic{diff.Added[2], diff.Added[0], diff.Added[1]},
		Removed: []SnapshotDiagnostic{diff.Removed[0], diff.Removed[1]},
	}
	if got := FormatSnapshotDiff(shuffled); got != want {
		t.Fatalf("shuffled formatted diff = %q, want %q", got, want)
	}
}

func TestFormatSnapshotDiffNormalizesPathsBeforeSorting(t *testing.T) {
	diff := SnapshotDiff{
		Added: []SnapshotDiagnostic{
			{Project: `self\z`, File: `src\Z.bas`, Surface: SurfaceAnalyze, Code: "VBA209", Severity: "warning", Line: 2},
			{Project: "self/a", File: "src/A.bas", Surface: SurfaceAnalyze, Code: "VBA209", Severity: "warning", Line: 1},
		},
	}

	want := "Real-world static-analysis corpus changed\n\n" +
		"VBA209 +2 -0\n" +
		"  + self/a/src/A.bas:1 [analyze warning]\n" +
		"  + self/z/src/Z.bas:2 [analyze warning]\n"
	if got := FormatSnapshotDiff(diff); got != want {
		t.Fatalf("formatted diff = %q, want %q", got, want)
	}
}

func TestFormatSnapshotDiffEmpty(t *testing.T) {
	if got := FormatSnapshotDiff(SnapshotDiff{}); got != "" {
		t.Fatalf("empty formatted diff = %q, want empty", got)
	}
}

func TestFormatSnapshotDiffRetainsLargeRuleDelta(t *testing.T) {
	const count = 73
	added := make([]SnapshotDiagnostic, count)
	for i := range added {
		added[i] = SnapshotDiagnostic{
			Project:  fmt.Sprintf("self/project-%02d", i),
			File:     "src/Main.bas",
			Surface:  SurfaceAnalyze,
			Code:     "VBA209",
			Severity: "warning",
			Line:     i + 1,
		}
	}

	report := FormatSnapshotDiff(SnapshotDiff{Added: added})
	if !strings.Contains(report, "VBA209 +73 -0\n") {
		excerpt := report
		if len(excerpt) > 100 {
			excerpt = excerpt[:100]
		}
		t.Fatalf("large report missing aggregate header: %q", excerpt)
	}
	if got := strings.Count(report, "\n  + "); got != count {
		t.Fatalf("large report detail count = %d, want %d", got, count)
	}
}

func TestSnapshotNormalizationDropsDiagnosticProse(t *testing.T) {
	first := normalizeLintDiagnostics("self/example", []lint.Issue{{Code: "VB001", Severity: "warning", File: "Main.bas", Line: 2, Message: "old explanation", Suggestion: "old suggestion"}})
	second := normalizeLintDiagnostics("self/example", []lint.Issue{{Code: "VB001", Severity: "warning", File: "Main.bas", Line: 2, Message: "new explanation", Suggestion: "new suggestion"}})
	firstSet, err := SnapshotSetFromReport(Report{Surfaces: []SurfaceResult{{Project: "self/example", Surface: SurfaceLint, Diagnostics: first}}})
	if err != nil {
		t.Fatal(err)
	}
	secondSet, err := SnapshotSetFromReport(Report{Surfaces: []SurfaceResult{{Project: "self/example", Surface: SurfaceLint, Diagnostics: second}}})
	if err != nil {
		t.Fatal(err)
	}
	if diff := CompareSnapshotSets(firstSet, secondSet); !diff.Empty() {
		t.Fatalf("prose-only change produced snapshot diff: %#v", diff)
	}
}

func TestSnapshotNormalizationRejectsTemporaryAndAbsolutePaths(t *testing.T) {
	if _, err := normalizeSnapshotDiagnostic(Diagnostic{Project: "self/example", Surface: SurfaceAnalyze, File: filepath.Join("src", "Main.bas"), Code: "VBA001", Severity: "warning", Line: 1}); err != nil {
		t.Fatalf("platform-relative path rejected: %v", err)
	}
	if _, err := normalizeSnapshotDiagnostic(Diagnostic{Project: "self/example", Surface: SurfaceAnalyze, File: filepath.Join(t.TempDir(), "Main.bas"), Code: "VBA001", Severity: "warning", Line: 1}); err == nil {
		t.Fatal("absolute temporary path was accepted")
	}
}

func TestWriteSnapshotSetDoesNotPublishInvalidUpdate(t *testing.T) {
	root := filepath.Join(t.TempDir(), "snapshots")
	id := SnapshotID{Project: "self/example", Surface: SurfaceLint}
	staleID := SnapshotID{Project: "self/stale", Surface: SurfaceLint}
	original := SnapshotSet{id: {snapshotTestRow(id.Project, id.Surface, "Main.bas", 1, 0, "VB001")}}
	if err := WriteSnapshotSet(root, SnapshotSet{id: original[id], staleID: {}}); err != nil {
		t.Fatal(err)
	}
	if err := WriteSnapshotSet(root, original); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(root, "lint", "self", "stale.jsonl")); !os.IsNotExist(err) {
		t.Fatalf("stale snapshot still exists, stat error = %v", err)
	}
	invalid := SnapshotSet{id: {{Project: id.Project, Surface: id.Surface, File: "../escape.bas", Code: "VB001", Severity: "warning", Line: 1}}}
	if err := WriteSnapshotSet(root, invalid); err == nil {
		t.Fatal("invalid snapshot update was accepted")
	}
	got, err := LoadSnapshotSet(root, original.IDs())
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, original) {
		t.Fatalf("invalid update changed existing snapshots: got %#v, want %#v", got, original)
	}
}

func runRealWorldCorpus(repoRoot string, logf ...func(string, ...any)) (Report, error) {
	var timingLogf func(string, ...any)
	if len(logf) > 0 {
		timingLogf = logf[0]
	}
	started := time.Now()
	defer func() {
		if timingLogf != nil {
			timingLogf("corpus timing: total=%s", time.Since(started))
		}
	}()

	nativeProjects, err := DiscoverExampleProjects(repoRoot)
	if err != nil {
		return Report{}, err
	}
	var nativeReport Report
	nativeFailed := false
	for _, project := range nativeProjects {
		projectStarted := time.Now()
		report, err := RunProjects([]NativeProject{project})
		nativeReport.Surfaces = append(nativeReport.Surfaces, report.Surfaces...)
		nativeReport.Diagnostics = append(nativeReport.Diagnostics, report.Diagnostics...)
		nativeReport.Failures = append(nativeReport.Failures, report.Failures...)
		if err != nil {
			nativeFailed = true
		}
		if timingLogf != nil {
			timingLogf("corpus timing: project=%s elapsed=%s", project.ID, time.Since(projectStarted))
		}
	}
	var nativeErr error
	if nativeFailed {
		nativeErr = &RunError{Failures: append([]Failure(nil), nativeReport.Failures...)}
	}
	manifestPath := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "manifest.json")
	manifest, corpusRoot, manifestErr := LoadManifest(manifestPath)
	if manifestErr != nil {
		return nativeReport, errors.Join(nativeErr, manifestErr)
	}
	thirdProjects, selectErr := SelectThirdPartyProjects(manifest.Projects, nil)
	if selectErr != nil {
		return nativeReport, errors.Join(nativeErr, selectErr)
	}
	var thirdReport Report
	thirdFailed := false
	for _, project := range thirdProjects {
		projectStarted := time.Now()
		report, err := RunThirdPartyProjects(corpusRoot, []Project{project}, MaterializeOptions{})
		thirdReport.Surfaces = append(thirdReport.Surfaces, report.Surfaces...)
		thirdReport.Diagnostics = append(thirdReport.Diagnostics, report.Diagnostics...)
		thirdReport.Failures = append(thirdReport.Failures, report.Failures...)
		thirdReport.Skipped = append(thirdReport.Skipped, report.Skipped...)
		if err != nil {
			thirdFailed = true
		}
		if timingLogf != nil {
			timingLogf("corpus timing: project=third_party/%s elapsed=%s", project.ID, time.Since(projectStarted))
		}
	}
	var thirdErr error
	if thirdFailed {
		sortThirdPartyFailuresAndSkips(&thirdReport)
		thirdErr = &RunError{Failures: append([]Failure(nil), thirdReport.Failures...)}
	}
	combined := Report{
		Surfaces:    append(append([]SurfaceResult(nil), nativeReport.Surfaces...), thirdReport.Surfaces...),
		Diagnostics: append(append([]Diagnostic(nil), nativeReport.Diagnostics...), thirdReport.Diagnostics...),
		Failures:    append(append([]Failure(nil), nativeReport.Failures...), thirdReport.Failures...),
		Skipped:     append(append([]Skip(nil), nativeReport.Skipped...), thirdReport.Skipped...),
	}
	sortDiagnostics(combined.Diagnostics)
	sortThirdPartyFailuresAndSkips(&combined)
	return combined, errors.Join(nativeErr, thirdErr)
}

func TestRealWorldCorpusSnapshots(t *testing.T) {
	repoRoot, err := filepath.Abs(filepath.Join("..", "..", ".."))
	if err != nil {
		t.Fatal(err)
	}
	// Snapshot identities must not depend on a developer's generated TypeLib
	// database. Use only the embedded built-in database for every platform.
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	report, runErr := runRealWorldCorpus(repoRoot, t.Logf)
	if runErr != nil {
		t.Fatalf("real-world corpus execution failed: %v", runErr)
	}
	actual, err := SnapshotSetFromReport(report)
	if err != nil {
		t.Fatal(err)
	}
	snapshotRoot := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "snapshots")
	if os.Getenv("XLFLOW_UPDATE_CORPUS_SNAPSHOTS") == "1" {
		if err := WriteSnapshotSet(snapshotRoot, actual); err != nil {
			t.Fatal(err)
		}
		return
	}
	expected, err := LoadSnapshotSet(snapshotRoot, actual.IDs())
	if err != nil {
		t.Fatal(err)
	}
	if diff := CompareSnapshotSets(expected, actual); !diff.Empty() {
		t.Fatal(FormatSnapshotDiff(diff))
	}
}
