package corpus

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"sync/atomic"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/typedb"
)

const (
	runRealWorldCorpusEnv     = "XLFLOW_RUN_REALWORLD_CORPUS"
	updateCorpusSnapshotsEnv  = "XLFLOW_UPDATE_CORPUS_SNAPSHOTS"
	maxParallelCorpusProjects = 2
)

func realWorldCorpusMode() (run, update bool) {
	run = os.Getenv(runRealWorldCorpusEnv) == "1"
	update = run && os.Getenv(updateCorpusSnapshotsEnv) == "1"
	return run, update
}

func TestRealWorldCorpusSnapshotMode(t *testing.T) {
	tests := []struct {
		name       string
		runEnv     string
		updateEnv  string
		wantRun    bool
		wantUpdate bool
	}{
		{name: "disabled", wantRun: false, wantUpdate: false},
		{name: "update requires explicit run", updateEnv: "1", wantRun: false, wantUpdate: false},
		{name: "run only", runEnv: "1", wantRun: true, wantUpdate: false},
		{name: "run and update", runEnv: "1", updateEnv: "1", wantRun: true, wantUpdate: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Setenv(runRealWorldCorpusEnv, tt.runEnv)
			t.Setenv(updateCorpusSnapshotsEnv, tt.updateEnv)
			gotRun, gotUpdate := realWorldCorpusMode()
			if gotRun != tt.wantRun || gotUpdate != tt.wantUpdate {
				t.Fatalf("real-world corpus mode = run:%t update:%t, want run:%t update:%t", gotRun, gotUpdate, tt.wantRun, tt.wantUpdate)
			}
		})
	}
}

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

type corpusProjectResult struct {
	report  Report
	err     error
	elapsed time.Duration
}

func runParallelCorpusProjects(t *testing.T, group string, names []string, run func(int) (Report, error)) []corpusProjectResult {
	t.Helper()
	results := make([]corpusProjectResult, len(names))
	limiter := make(chan struct{}, maxParallelCorpusProjects)
	t.Run(group, func(t *testing.T) {
		for i, name := range names {
			i, name := i, name
			t.Run(name, func(t *testing.T) {
				t.Parallel()
				limiter <- struct{}{}
				defer func() { <-limiter }()
				started := time.Now()
				report, err := run(i)
				results[i] = corpusProjectResult{
					report:  report,
					err:     err,
					elapsed: time.Since(started),
				}
			})
		}
	})
	return results
}

func appendCorpusReport(dst *Report, src Report) {
	dst.Surfaces = append(dst.Surfaces, src.Surfaces...)
	dst.Diagnostics = append(dst.Diagnostics, src.Diagnostics...)
	dst.Failures = append(dst.Failures, src.Failures...)
	dst.Skipped = append(dst.Skipped, src.Skipped...)
}

func runRealWorldCorpus(t *testing.T, repoRoot string, logf ...func(string, ...any)) (Report, error) {
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
	nativeNames := make([]string, len(nativeProjects))
	for i, project := range nativeProjects {
		nativeNames[i] = project.ID
	}
	nativeResults := runParallelCorpusProjects(t, "native", nativeNames, func(i int) (Report, error) {
		return RunProjects([]NativeProject{nativeProjects[i]})
	})
	for i, result := range nativeResults {
		appendCorpusReport(&nativeReport, result.report)
		if result.err != nil {
			nativeFailed = true
		}
		if timingLogf != nil {
			timingLogf("corpus timing: project=%s elapsed=%s", nativeProjects[i].ID, result.elapsed)
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
	inventory, inventoryErr := ValidateInventory(corpusRoot, manifest)
	if timingLogf != nil {
		if inventoryErr == nil {
			timingLogf("%s", FormatInventorySummary(inventory))
		} else {
			timingLogf("corpus inventory validation failed: %v", inventoryErr)
		}
	}
	if inventoryErr != nil {
		return nativeReport, errors.Join(nativeErr, inventoryErr)
	}
	thirdProjects, selectErr := SelectThirdPartyProjects(manifest.Projects, nil)
	if selectErr != nil {
		return nativeReport, errors.Join(nativeErr, selectErr)
	}
	var thirdReport Report
	thirdFailed := false
	thirdNames := make([]string, len(thirdProjects))
	for i, project := range thirdProjects {
		thirdNames[i] = project.ID
	}
	thirdResults := runParallelCorpusProjects(t, "third_party", thirdNames, func(i int) (Report, error) {
		return RunThirdPartyProjects(corpusRoot, []Project{thirdProjects[i]}, MaterializeOptions{})
	})
	for i, result := range thirdResults {
		appendCorpusReport(&thirdReport, result.report)
		if result.err != nil {
			thirdFailed = true
		}
		if timingLogf != nil {
			timingLogf("corpus timing: project=third_party/%s elapsed=%s", thirdProjects[i].ID, result.elapsed)
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
	if timingLogf != nil {
		timingLogf("corpus inventory: skipped=%d", len(combined.Skipped))
		timingLogf("%s", FormatDiagnosticSummary(SummarizeDiagnostics(combined)))
	}
	return combined, errors.Join(nativeErr, thirdErr)
}

func TestParallelCorpusProjectsPreserveOrderAndFailures(t *testing.T) {
	names := []string{"first", "second", "third"}
	wantPeak := maxParallelCorpusProjects
	if parallelFlag := flag.Lookup("test.parallel"); parallelFlag != nil {
		if parallel, err := strconv.Atoi(parallelFlag.Value.String()); err == nil && parallel < wantPeak {
			wantPeak = parallel
		}
	}
	admitted := make(chan struct{}, len(names))
	completed := make(chan int, len(names))
	release := make(chan struct{})
	var active atomic.Int32
	var peak atomic.Int32
	go func() {
		for i := 0; i < wantPeak; i++ {
			<-admitted
		}
		close(release)
	}()
	results := runParallelCorpusProjects(t, "synthetic", names, func(i int) (Report, error) {
		current := active.Add(1)
		for {
			previous := peak.Load()
			if current <= previous || peak.CompareAndSwap(previous, current) {
				break
			}
		}
		admitted <- struct{}{}
		<-release
		active.Add(-1)
		completed <- i
		report := Report{
			Diagnostics: []Diagnostic{{
				Project: names[i], Surface: SurfaceAnalyze, Code: fmt.Sprintf("VBA%d", i),
			}},
		}
		if i != 1 {
			return report, nil
		}
		failure := Failure{Project: names[i], Surface: SurfaceAnalyze, Kind: FailureExecution, Err: errors.New("synthetic worker failure")}
		report.Failures = []Failure{failure}
		return report, &RunError{Failures: []Failure{failure}}
	})

	if len(results) != len(names) {
		t.Fatalf("result count = %d, want %d", len(results), len(names))
	}
	if got := peak.Load(); got != int32(wantPeak) {
		t.Fatalf("peak corpus project concurrency = %d, want %d", got, wantPeak)
	}
	seen := make(map[int]bool, len(names))
	var combined Report
	var failures []Failure
	for i, result := range results {
		if len(result.report.Diagnostics) != 1 || result.report.Diagnostics[0].Project != names[i] {
			t.Fatalf("result[%d] project = %#v, want %q", i, result.report.Diagnostics, names[i])
		}
		appendCorpusReport(&combined, result.report)
		if result.err == nil {
			continue
		}
		var runErr *RunError
		if !errors.As(result.err, &runErr) {
			t.Fatalf("result[%d] error = %T, want *RunError", i, result.err)
		}
		failures = append(failures, runErr.Failures...)
	}
	for len(seen) < len(names) {
		seen[<-completed] = true
	}
	if len(combined.Diagnostics) != len(names) {
		t.Fatalf("combined diagnostics = %d, want %d", len(combined.Diagnostics), len(names))
	}
	if len(combined.Failures) != 1 || len(failures) != 1 || failures[0].Project != names[1] {
		t.Fatalf("combined failures = %#v, worker failures = %#v", combined.Failures, failures)
	}
	if got := (&RunError{Failures: failures}).Error(); got != "static-analysis corpus execution failed for 1 project surface(s)" {
		t.Fatalf("RunError summary = %q", got)
	}
}

func TestRealWorldCorpusSnapshots(t *testing.T) {
	shouldRun, update := realWorldCorpusMode()
	if !shouldRun {
		t.Skipf("set %s=1 to run the real-world corpus snapshot test", runRealWorldCorpusEnv)
	}
	repoRoot, err := filepath.Abs(filepath.Join("..", "..", ".."))
	if err != nil {
		t.Fatal(err)
	}
	// Snapshot identities must not depend on a developer's generated TypeLib
	// database. Use only the embedded built-in database for every platform.
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	report, runErr := runRealWorldCorpus(t, repoRoot, t.Logf)
	if runErr != nil {
		t.Fatalf("real-world corpus execution failed: %v", runErr)
	}
	actual, err := SnapshotSetFromReport(report)
	if err != nil {
		t.Fatal(err)
	}
	reviewPath := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "reviews", "diagnostics.jsonl")
	reviews, err := LoadDiagnosticReviews(reviewPath)
	if err != nil {
		t.Fatal(err)
	}
	if err := ValidateReviewSources(repoRoot, filepath.Join(repoRoot, "testdata", "static-analysis-corpus"), reviews); err != nil {
		t.Fatal(err)
	}
	metrics, err := EvaluateDiagnosticReviews(reviews, report)
	if err != nil {
		t.Fatal(err)
	}
	t.Log(FormatReviewMetrics(metrics))
	snapshotRoot := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "snapshots")
	if update {
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
