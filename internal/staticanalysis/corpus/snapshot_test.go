package corpus

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/lint"
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

func runRealWorldCorpus(repoRoot string) (Report, error) {
	nativeProjects, err := DiscoverExampleProjects(repoRoot)
	if err != nil {
		return Report{}, err
	}
	nativeReport, nativeErr := RunProjects(nativeProjects)
	manifestPath := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "manifest.json")
	manifest, corpusRoot, manifestErr := LoadManifest(manifestPath)
	if manifestErr != nil {
		return nativeReport, errors.Join(nativeErr, manifestErr)
	}
	thirdProjects, selectErr := SelectThirdPartyProjects(manifest.Projects, nil)
	if selectErr != nil {
		return nativeReport, errors.Join(nativeErr, selectErr)
	}
	thirdReport, thirdErr := RunThirdPartyProjects(corpusRoot, thirdProjects, MaterializeOptions{})
	combined := Report{
		Surfaces:    append(append([]SurfaceResult(nil), nativeReport.Surfaces...), thirdReport.Surfaces...),
		Diagnostics: append(append([]Diagnostic(nil), nativeReport.Diagnostics...), thirdReport.Diagnostics...),
		Failures:    append(append([]Failure(nil), nativeReport.Failures...), thirdReport.Failures...),
		Skipped:     append(append([]Skip(nil), nativeReport.Skipped...), thirdReport.Skipped...),
	}
	return combined, errors.Join(nativeErr, thirdErr)
}

func TestRealWorldCorpusSnapshots(t *testing.T) {
	repoRoot, err := filepath.Abs(filepath.Join("..", "..", ".."))
	if err != nil {
		t.Fatal(err)
	}
	report, runErr := runRealWorldCorpus(repoRoot)
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
		t.Fatalf("real-world static-analysis corpus changed: added=%v removed=%v", diff.Added, diff.Removed)
	}
}
