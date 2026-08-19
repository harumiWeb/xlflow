package corpus

import (
	"context"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
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
			b.Run("analyze-only", func(b *testing.B) {
				benchmarkRealWorldCorpusAnalyzeOnly(b, corpusRoot, project)
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

func benchmarkRealWorldCorpusAnalyzeOnly(b *testing.B, corpusRoot string, project Project) {
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
	ctx := analysisstats.WithRecorder(context.Background(), recorder)
	b.ReportAllocs()
	b.ResetTimer()
	var findings, warnings int
	for i := 0; i < b.N; i++ {
		result, err := (analyze.Analyzer{RootDir: workspace.Root, Config: cfg}).RunResultContext(ctx)
		if err != nil {
			b.Fatalf("analyze %q: %v", project.ID, err)
		}
		findings += len(result.Findings)
		warnings += len(result.Warnings)
	}
	b.ReportMetric(float64(findings)/float64(b.N), "findings/op")
	b.ReportMetric(float64(project.SourceCounts.Total()), "source-files/op")
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
		metricName := "stage_" + strings.ReplaceAll(stage.Name, "-", "_")
		b.ReportMetric(float64(stage.Elapsed.Nanoseconds())/float64(iterations), metricName+"-ns/op")
		b.ReportMetric(float64(stage.Calls)/float64(iterations), metricName+"-calls/op")
		b.ReportMetric(float64(stage.ResultCount)/float64(iterations), metricName+"-results/op")
	}
	for _, counter := range counters {
		metricName := "counter_" + strings.ReplaceAll(counter.Name, "-", "_")
		if strings.HasPrefix(counter.Name, "max_") {
			b.ReportMetric(float64(counter.Value), metricName)
			continue
		}
		b.ReportMetric(float64(counter.Value)/float64(iterations), metricName+"/op")
	}
}
