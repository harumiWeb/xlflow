package corpus

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

var realWorldCorpusBenchmarkProjectIDs = []string{"std-vba", "ronecone"}

// BenchmarkRealWorldCorpus provides an opt-in, developer-only profiling path
// for the slowest checked-in third-party projects. The benchmark deliberately
// bypasses snapshot and review evaluation; invoke it with -bench and use a
// project/stage sub-benchmark name to keep profiles focused.
func BenchmarkRealWorldCorpus(b *testing.B) {
	repoRoot, err := filepath.Abs(filepath.Join("..", "..", ".."))
	if err != nil {
		b.Fatal(err)
	}
	manifestPath := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "manifest.json")
	manifest, corpusRoot, err := LoadManifest(manifestPath)
	if err != nil {
		b.Fatal(err)
	}
	projects, err := SelectThirdPartyProjects(manifest.Projects, realWorldCorpusBenchmarkProjectIDs)
	if err != nil {
		b.Fatal(err)
	}
	if len(projects) != len(realWorldCorpusBenchmarkProjectIDs) {
		b.Fatalf("selected %d benchmark projects, want %d", len(projects), len(realWorldCorpusBenchmarkProjectIDs))
	}

	// Results must not depend on a developer-generated TypeLib database.
	b.Setenv(typedb.EnvDir, filepath.Join(b.TempDir(), "typelib"))
	for _, project := range projects {
		project := project
		if !project.Enabled {
			b.Fatalf("benchmark project %q is disabled", project.ID)
		}
		b.Run(project.ID, func(b *testing.B) {
			b.Run("full-pipeline", func(b *testing.B) {
				benchmarkRealWorldCorpusFullPipeline(b, corpusRoot, project)
			})
			b.Run("analyze-only/cold", func(b *testing.B) {
				benchmarkRealWorldCorpusAnalyzeOnlyMode(b, corpusRoot, project, corpusBenchmarkCold)
			})
			b.Run("analyze-only/warm", func(b *testing.B) {
				benchmarkRealWorldCorpusAnalyzeOnlyMode(b, corpusRoot, project, corpusBenchmarkWarm)
			})
			b.Run("analyze-only/local-edit", func(b *testing.B) {
				benchmarkRealWorldCorpusEditMode(b, corpusRoot, project, corpusBenchmarkLocalEdit)
			})
			b.Run("analyze-only/dependency-edit", func(b *testing.B) {
				benchmarkRealWorldCorpusEditMode(b, corpusRoot, project, corpusBenchmarkDependencyEdit)
			})
		})
	}
}

func benchmarkRealWorldCorpusFullPipeline(b *testing.B, corpusRoot string, project Project) {
	tempRoot := b.TempDir()
	b.ReportAllocs()
	b.ResetTimer()

	var diagnostics, surfaces int
	for i := 0; i < b.N; i++ {
		report, err := RunThirdPartyProjects(corpusRoot, []Project{project}, MaterializeOptions{TempRoot: tempRoot})
		if err != nil {
			b.Fatalf("full pipeline for %q failed: %v", project.ID, err)
		}
		if len(report.Failures) != 0 {
			b.Fatalf("full pipeline for %q reported %d failure(s)", project.ID, len(report.Failures))
		}
		if len(report.Skipped) != 0 {
			b.Fatalf("full pipeline for %q unexpectedly skipped", project.ID)
		}
		diagnostics += len(report.Diagnostics)
		surfaces += len(report.Surfaces)
	}
	b.ReportMetric(float64(diagnostics)/float64(b.N), "diagnostics/op")
	b.ReportMetric(float64(project.SourceCounts.Total()), "source-files/op")
	b.ReportMetric(float64(surfaces)/float64(b.N), "surfaces/op")
}

type corpusBenchmarkMode uint8

const (
	corpusBenchmarkCold corpusBenchmarkMode = iota
	corpusBenchmarkWarm
)

type corpusBenchmarkEditMode uint8

const (
	corpusBenchmarkLocalEdit corpusBenchmarkEditMode = iota
	corpusBenchmarkDependencyEdit
)

// benchmarkRealWorldCorpusAnalyzeOnlyMode keeps cold and warm measurements
// independent. Cold runs deliberately use a new process-local query store for
// each operation, while warm runs prime one store outside the timed region and
// reuse it for every measured operation. The source workspace is materialized
// once in both cases so filesystem setup is not confused with query-store
// overhead.
func benchmarkRealWorldCorpusAnalyzeOnlyMode(b *testing.B, corpusRoot string, project Project, mode corpusBenchmarkMode) {
	workspace, err := MaterializeThirdPartyProject(corpusRoot, project, MaterializeOptions{TempRoot: b.TempDir()})
	if err != nil {
		b.Fatalf("materialize %q: %v", project.ID, err)
	}
	b.Cleanup(func() {
		if err := workspace.Close(); err != nil {
			b.Errorf("cleanup %q: %v", project.ID, err)
		}
	})
	exec := productionExecutor()
	cfg, err := exec.loadConfig(workspace.Root)
	if err != nil {
		b.Fatalf("load config for %q: %v", project.ID, err)
	}

	recorder := analysisstats.NewRecorder()
	var store *semanticquery.Store
	if mode == corpusBenchmarkWarm {
		store = semanticquery.New(semanticquery.Options{})
		// Prime the same revision before starting the timer. The prime context
		// intentionally has no recorder so its misses/recomputations are not
		// reported as part of the warm operation metrics.
		if _, _, err := runCorpusAnalyzeOnly(contextWithSemanticQueryStore(context.Background(), store, nil), workspace.Root, cfg); err != nil {
			b.Fatalf("prime analysis %q: %v", project.ID, err)
		}
	}
	b.ReportAllocs()
	b.ResetTimer()
	var findings, warnings int
	for i := 0; i < b.N; i++ {
		if mode == corpusBenchmarkCold {
			store = semanticquery.New(semanticquery.Options{})
		}
		result, resultWarnings, err := runCorpusAnalyzeOnly(contextWithSemanticQueryStore(context.Background(), store, recorder), workspace.Root, cfg)
		if err != nil {
			b.Fatalf("analyze %q: %v", project.ID, err)
		}
		findings += result
		warnings += resultWarnings
	}
	b.ReportMetric(float64(findings)/float64(b.N), "findings/op")
	b.ReportMetric(float64(project.SourceCounts.Total()), "source-files/op")
	b.ReportMetric(float64(warnings)/float64(b.N), "warnings/op")
	reportCorpusAnalysisRecorderMetrics(b, recorder, b.N)
}

func contextWithSemanticQueryStore(ctx context.Context, store *semanticquery.Store, recorder *analysisstats.Recorder) context.Context {
	if ctx == nil {
		ctx = context.Background()
	}
	if recorder != nil {
		ctx = analysisstats.WithRecorder(ctx, recorder)
	}
	return semanticquery.WithContext(ctx, semanticquery.Context{Store: store, Metrics: recorder})
}

func runCorpusAnalyzeOnly(ctx context.Context, root string, cfg config.Config) (findings, warnings int, err error) {
	result, err := (analyze.Analyzer{RootDir: root, Config: cfg}).RunResultContext(ctx)
	if err != nil {
		return 0, 0, err
	}
	return len(result.Findings), len(result.Warnings), nil
}

// benchmarkRealWorldCorpusEditMode exercises revision reuse with benchmark-
// owned modules. The dependency case cycles callee body/signature/effect edits
// and a caller redirect. The files are deliberately independent from the
// checked-in corpus so the benchmark never mutates a real source fixture. File
// writes are outside the timed region; the timed operation includes parsing,
// revision fingerprinting, invalidation, and analysis of the edited revision.
func benchmarkRealWorldCorpusEditMode(b *testing.B, corpusRoot string, project Project, mode corpusBenchmarkEditMode) {
	workspace, err := MaterializeThirdPartyProject(corpusRoot, project, MaterializeOptions{TempRoot: b.TempDir()})
	if err != nil {
		b.Fatalf("materialize %q: %v", project.ID, err)
	}
	b.Cleanup(func() {
		if err := workspace.Close(); err != nil {
			b.Errorf("cleanup %q: %v", project.ID, err)
		}
	})
	exec := productionExecutor()
	cfg, err := exec.loadConfig(workspace.Root)
	if err != nil {
		b.Fatalf("load config for %q: %v", project.ID, err)
	}
	moduleRoot := filepath.Join(workspace.Root, "src", "modules")
	localPath := filepath.Join(moduleRoot, "Benchmark724Local.bas")
	calleePath := filepath.Join(moduleRoot, "Benchmark724DependencyCallee.bas")
	callerPath := filepath.Join(moduleRoot, "Benchmark724DependencyCaller.bas")
	redirectPath := filepath.Join(moduleRoot, "Benchmark724DependencyRedirect.bas")
	// Prime with a state that differs from the first timed operation so a
	// benchtime=1x run still exercises one real revision edit.
	localSource := []byte("Attribute VB_Name = \"Benchmark724Local\"\nOption Explicit\nPublic Sub BenchmarkLocal()\n    Dim values() As Long\n    ReDim values(1 To 2)\n    values(1) = 0\nEnd Sub\n")
	calleeSource := []byte("Attribute VB_Name = \"Benchmark724DependencyCallee\"\nOption Explicit\nPublic Function Benchmark724DependencyCallee(value As Long) As Long\n    Dim values() As Long\n    ReDim values(1 To 2)\n    values(1) = value\n    Range(\"A1\").Value = value\n    Benchmark724DependencyCallee = value\nEnd Function\n")
	callerSource := []byte("Attribute VB_Name = \"Benchmark724DependencyCaller\"\nOption Explicit\nPublic Sub Benchmark724DependencyCaller()\n    Dim value As Long\n    value = Benchmark724DependencyCallee(1)\nEnd Sub\n")
	redirectSource := []byte("Attribute VB_Name = \"Benchmark724DependencyRedirect\"\nOption Explicit\nPublic Function Benchmark724DependencyRedirect(value As Long) As Long\n    Benchmark724DependencyRedirect = value\nEnd Function\n")
	for path, source := range map[string][]byte{localPath: localSource, calleePath: calleeSource, callerPath: callerSource, redirectPath: redirectSource} {
		if err := os.WriteFile(path, source, 0o644); err != nil {
			b.Fatalf("write benchmark module %q: %v", path, err)
		}
	}
	store := semanticquery.New(semanticquery.Options{})
	if _, _, err := runCorpusAnalyzeOnly(contextWithSemanticQueryStore(context.Background(), store, nil), workspace.Root, cfg); err != nil {
		b.Fatalf("prime edit analysis %q: %v", project.ID, err)
	}
	recorder := analysisstats.NewRecorder()
	b.ReportAllocs()
	b.ResetTimer()
	var findings, warnings int
	for i := 0; i < b.N; i++ {
		b.StopTimer()
		if mode == corpusBenchmarkLocalEdit {
			body := "    values(1) = 1\n"
			if i&1 == 1 {
				body = "    values(1) = 2\n"
			}
			updated := []byte("Attribute VB_Name = \"Benchmark724Local\"\nOption Explicit\nPublic Sub BenchmarkLocal()\n    Dim values() As Long\n    ReDim values(1 To 2)\n" + body + "End Sub\n")
			if err := os.WriteFile(localPath, updated, 0o644); err != nil {
				b.Fatalf("write local edit: %v", err)
			}
		} else {
			variant := i & 3
			parameterType := "Long"
			if variant == 2 {
				parameterType = "Variant"
			}
			calleeBody := "    Dim values() As Long\n    ReDim values(1 To 2)\n    values(1) = value\n    Benchmark724DependencyCallee = value\n"
			if variant == 1 {
				calleeBody = "    Dim values() As Long\n    ReDim values(1 To 2)\n    values(1) = value\n    Range(\"A1\").Value = value\n    Benchmark724DependencyCallee = value\n"
			}
			updated := []byte("Attribute VB_Name = \"Benchmark724DependencyCallee\"\nOption Explicit\nPublic Function Benchmark724DependencyCallee(value As " + parameterType + ") As Long\n" + calleeBody + "End Function\n")
			if err := os.WriteFile(calleePath, updated, 0o644); err != nil {
				b.Fatalf("write dependency edit: %v", err)
			}
			if variant == 3 {
				callerSource = []byte("Attribute VB_Name = \"Benchmark724DependencyCaller\"\nOption Explicit\nPublic Sub Benchmark724DependencyCaller()\n    Dim value As Long\n    value = Benchmark724DependencyRedirect(1)\nEnd Sub\n")
			} else {
				callerSource = []byte("Attribute VB_Name = \"Benchmark724DependencyCaller\"\nOption Explicit\nPublic Sub Benchmark724DependencyCaller()\n    Dim value As Long\n    value = Benchmark724DependencyCallee(1)\nEnd Sub\n")
			}
			if err := os.WriteFile(callerPath, callerSource, 0o644); err != nil {
				b.Fatalf("write dependency caller edit: %v", err)
			}
		}
		b.StartTimer()
		result, resultWarnings, err := runCorpusAnalyzeOnly(contextWithSemanticQueryStore(context.Background(), store, recorder), workspace.Root, cfg)
		if err != nil {
			b.Fatalf("analyze edit %q: %v", project.ID, err)
		}
		findings += result
		warnings += resultWarnings
	}
	b.ReportMetric(float64(findings)/float64(b.N), "findings/op")
	b.ReportMetric(float64(project.SourceCounts.Total()+4), "source-files/op")
	b.ReportMetric(float64(warnings)/float64(b.N), "warnings/op")
	reportCorpusAnalysisRecorderMetrics(b, recorder, b.N)
}

func reportCorpusAnalysisRecorderMetrics(b *testing.B, recorder *analysisstats.Recorder, iterations int) {
	b.Helper()
	if recorder == nil || iterations <= 0 {
		return
	}
	stages, counters := recorder.Totals()
	for _, stage := range stages {
		// Domain stages use slash-qualified names. Normalize the slash before
		// emitting benchmark metrics so each stage is reported as one stable
		// metric name (the /op suffix below remains the standard unit suffix).
		metricName := "stage_" + normalizeCorpusBenchmarkMetricName(stage.Name)
		b.ReportMetric(float64(stage.Elapsed.Nanoseconds())/float64(iterations), metricName+"-ns/op")
		b.ReportMetric(float64(stage.Calls)/float64(iterations), metricName+"-calls/op")
		b.ReportMetric(float64(stage.ResultCount)/float64(iterations), metricName+"-results/op")
	}
	for _, counter := range counters {
		metricName := "counter_" + normalizeCorpusBenchmarkMetricName(counter.Name)
		if strings.HasPrefix(counter.Name, "max_") {
			b.ReportMetric(float64(counter.Value), metricName)
			continue
		}
		b.ReportMetric(float64(counter.Value)/float64(iterations), metricName+"/op")
	}
}

func normalizeCorpusBenchmarkMetricName(name string) string {
	return strings.NewReplacer("-", "_", "/", "_").Replace(name)
}

func TestCorpusBenchmarkMetricNameNormalizesDomainSlash(t *testing.T) {
	if got := normalizeCorpusBenchmarkMetricName("procedure_local/source_scan"); got != "procedure_local_source_scan" {
		t.Fatalf("normalized domain metric = %q, want %q", got, "procedure_local_source_scan")
	}
}
