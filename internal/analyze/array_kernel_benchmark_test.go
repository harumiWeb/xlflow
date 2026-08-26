package analyze

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

// arrayKernelBenchmarkWorkload describes the source shapes used to compare
// array-heavy procedure analysis. Source construction is deliberately outside
// the timed benchmark loop; the benchmark therefore measures analyzer work,
// including any procedure-local array preparation and CFG traversal.
type arrayKernelBenchmarkWorkload struct {
	name          string
	source        string
	lines         int
	procedures    int
	arrayMarkers  int
	statementHint int
}

// BenchmarkArrayAnalysisKernel covers the workload classes called out by the
// array semantic-kernel design. The single/all projection modes make it
// possible to compare a representative array rule with the simultaneously
// enabled array rule set. Recorder metrics are emitted through the shared
// benchmark helper, so newly added array_kernel_runs, array_cfg_walks, and
// array_projection_runs counters are reported without coupling this fixture to
// their eventual Go identifiers.
func BenchmarkArrayAnalysisKernel(b *testing.B) {
	for _, workload := range arrayKernelBenchmarkWorkloads() {
		workload := workload
		b.Run(workload.name, func(b *testing.B) {
			for _, mode := range []struct {
				name string
				cfg  func(*config.Config)
			}{
				{name: "no-array-rules", cfg: configureArrayBenchmarkRulesNone},
				{name: "VBA227", cfg: configureArrayBenchmarkRulesVBA227},
				{name: "all-array-rules", cfg: configureArrayBenchmarkRulesAll},
			} {
				mode := mode
				b.Run(mode.name, func(b *testing.B) {
					root := b.TempDir()
					fixture := writeArrayKernelBenchmarkProject(b, root, workload)
					b.Setenv(typedb.EnvDir, filepath.Join(b.TempDir(), "typelib"))
					cfg := config.Default()
					mode.cfg(&cfg)
					recorder := analysisstats.NewRecorder()
					ctx := analysisstats.WithRecorder(context.Background(), recorder)

					b.ReportAllocs()
					b.ResetTimer()
					findings := 0
					warnings := 0
					for i := 0; i < b.N; i++ {
						result, err := (Analyzer{RootDir: root, Config: cfg}).RunResultContext(ctx)
						if err != nil {
							b.Fatalf("analyze %s array fixture (%s): %v", workload.name, mode.name, err)
						}
						findings += len(result.Findings)
						warnings += len(result.Warnings)
					}
					b.StopTimer()

					b.ReportMetric(float64(fixture.lines), "lines/op")
					b.ReportMetric(float64(fixture.procedures), "procedures/op")
					b.ReportMetric(float64(fixture.arrayMarkers), "array-markers/op")
					b.ReportMetric(float64(fixture.statementHint), "statement-hint/op")
					b.ReportMetric(float64(findings)/float64(b.N), "findings/op")
					b.ReportMetric(float64(warnings)/float64(b.N), "warnings/op")
					reportAnalysisRecorderMetrics(b, recorder, b.N)
				})
			}
		})
	}
}

// BenchmarkArrayAdvancedCFGStrategies compares the migrated source-line,
// edge-refined, combined, and ByRef workloads against the retained legacy
// oracle. The strategy selector is private and benchmark-only; production
// analysis uses auto.
func BenchmarkArrayAdvancedCFGStrategies(b *testing.B) {
	workloads := []arrayKernelBenchmarkWorkload{
		arrayKernelBenchmarkWorkloads()[2], // wide ReDim/P preserve CFG
		arrayKernelBenchmarkWorkloads()[3], // edge-heavy branches
		arrayKernelBenchmarkWorkloads()[5], // ByRef/module summary path
	}
	strategies := []struct {
		name     string
		strategy arrayCFGStrategy
	}{
		{name: "auto", strategy: arrayCFGStrategyAuto},
		{name: "compact", strategy: arrayCFGStrategyCompact},
		{name: "legacy", strategy: arrayCFGStrategyLegacy},
	}
	for _, workload := range workloads {
		workload := workload
		for _, strategy := range strategies {
			strategy := strategy
			b.Run(workload.name+"/"+strategy.name, func(b *testing.B) {
				root := b.TempDir()
				writeArrayKernelBenchmarkProject(b, root, workload)
				b.Setenv(typedb.EnvDir, filepath.Join(b.TempDir(), "typelib"))
				cfg := config.Default()
				configureArrayBenchmarkRulesAll(&cfg)
				recorder := analysisstats.NewRecorder()
				ctx := analysisstats.WithRecorder(context.Background(), recorder)
				b.ReportAllocs()
				b.ResetTimer()
				for i := 0; i < b.N; i++ {
					if _, err := (Analyzer{RootDir: root, Config: cfg, arrayStrategy: strategy.strategy}).RunResultContext(ctx); err != nil {
						b.Fatal(err)
					}
				}
				b.StopTimer()
				_, counters := recorder.Totals()
				for _, name := range []string{"array_compact_cfg_walks", "array_legacy_cfg_walks", "array_cfg_fallbacks"} {
					for _, counter := range counters {
						if counter.Name == name {
							b.ReportMetric(float64(counter.Value)/float64(b.N), name+"/op")
						}
					}
				}
			})
		}
	}
}

// TestArrayKernelBenchmarkFixtureScale keeps the benchmark matrix visible to
// ordinary test runs without executing the analyzer for every workload.
func TestArrayKernelBenchmarkFixtureScale(t *testing.T) {
	for _, workload := range arrayKernelBenchmarkWorkloads() {
		workload := workload
		t.Run(workload.name, func(t *testing.T) {
			fixture := writeArrayKernelBenchmarkProject(t, t.TempDir(), workload)
			if fixture.lines != workload.lines || fixture.procedures != workload.procedures {
				t.Fatalf("fixture dimensions = lines %d/procedures %d, want %d/%d", fixture.lines, fixture.procedures, workload.lines, workload.procedures)
			}
			if fixture.arrayMarkers != workload.arrayMarkers {
				t.Fatalf("fixture array markers = %d, want %d", fixture.arrayMarkers, workload.arrayMarkers)
			}
			if fixture.source == "" {
				t.Fatal("fixture source path is empty")
			}
		})
	}
}

// TestArrayKernelProjectionWorkCounters makes the shared-kernel telemetry an
// executable contract. The counters are part of the current recorder API, so
// a missing counter is a test failure rather than a reason to skip coverage.
func TestArrayKernelProjectionWorkCounters(t *testing.T) {
	workload := arrayKernelBenchmarkWorkloads()[2] // redim-heavy
	single := runArrayKernelBenchmarkCounterFixture(t, workload, configureArrayBenchmarkRulesVBA227)
	all := runArrayKernelBenchmarkCounterFixture(t, workload, configureArrayBenchmarkRulesAll)
	for mode, counters := range map[string]map[string]uint64{"single": single, "all": all} {
		for _, name := range []string{"array_kernel_runs", "array_cfg_walks", "array_projection_runs"} {
			if _, ok := counters[name]; !ok {
				t.Fatalf("%s telemetry is missing %q: %+v", mode, name, counters)
			}
		}
	}
	if single["array_kernel_runs"] != 1 || all["array_kernel_runs"] != 1 {
		t.Fatalf("array kernel runs = single %d/all %d, want one per procedure revision", single["array_kernel_runs"], all["array_kernel_runs"])
	}
	if single["array_projection_runs"] != 1 || all["array_projection_runs"] != 5 {
		t.Fatalf("array projection runs = single %d/all %d, want exact single=1/all=5", single["array_projection_runs"], all["array_projection_runs"])
	}
	// VBA226 is the documented exceptional secondary pass. Any increase in
	// CFG walks beyond that one pass would indicate that projections are still
	// rebuilding the main array fixed point.
	if all["array_cfg_walks"] > single["array_cfg_walks"]+1 {
		t.Fatalf("array CFG walks = single %d/all %d, want at most one secondary pass", single["array_cfg_walks"], all["array_cfg_walks"])
	}
}

func TestArrayCFGStrategyTelemetry(t *testing.T) {
	workload := arrayKernelBenchmarkWorkloads()[5]
	for _, strategy := range []struct {
		name     string
		strategy arrayCFGStrategy
		wantKey  string
	}{
		{name: "compact", strategy: arrayCFGStrategyCompact, wantKey: "array_compact_cfg_walks"},
		{name: "legacy", strategy: arrayCFGStrategyLegacy, wantKey: "array_legacy_cfg_walks"},
	} {
		strategy := strategy
		t.Run(strategy.name, func(t *testing.T) {
			root := t.TempDir()
			writeArrayKernelBenchmarkProject(t, root, workload)
			t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
			cfg := config.Default()
			configureArrayBenchmarkRulesAll(&cfg)
			recorder := analysisstats.NewRecorder()
			if _, err := (Analyzer{RootDir: root, Config: cfg, arrayStrategy: strategy.strategy}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder)); err != nil {
				t.Fatal(err)
			}
			_, counters := recorder.Totals()
			found := false
			for _, counter := range counters {
				if counter.Name == strategy.wantKey && counter.Value > 0 {
					found = true
				}
			}
			if !found {
				t.Fatalf("strategy %s did not report %s: %+v", strategy.name, strategy.wantKey, counters)
			}
		})
	}
}

func TestArrayKernelComparisonOnlyDoesNotStartCFGWalk(t *testing.T) {
	workload := arrayKernelBenchmarkWorkload{
		name:          "comparison-only",
		source:        arrayBenchmarkComparisonOnlySource(),
		lines:         6,
		procedures:    1,
		arrayMarkers:  2,
		statementHint: 4,
	}
	counters := runArrayKernelBenchmarkCounterFixture(t, workload, configureArrayBenchmarkRulesVBA209)
	if counters["array_kernel_runs"] != 1 || counters["array_cfg_walks"] != 0 || counters["array_projection_runs"] != 1 {
		t.Fatalf("VBA209-only counters = %+v, want kernel=1/cfg=0/projection=1", counters)
	}
}

func TestArrayKernelAlwaysOnObjectArrayRunsWithRulesDisabled(t *testing.T) {
	workload := arrayKernelBenchmarkWorkload{
		name:          "object-array-only",
		source:        arrayBenchmarkObjectArrayOnlySource(),
		lines:         7,
		procedures:    1,
		arrayMarkers:  3,
		statementHint: 6,
	}
	counters := runArrayKernelBenchmarkCounterFixture(t, workload, configureArrayBenchmarkRulesNone)
	if counters["array_kernel_runs"] != 1 || counters["array_projection_runs"] != 1 {
		t.Fatalf("object-array-only counters = %+v, want kernel=1/projection=1", counters)
	}
}

func runArrayKernelBenchmarkCounterFixture(t *testing.T, workload arrayKernelBenchmarkWorkload, configure func(*config.Config)) map[string]uint64 {
	t.Helper()
	root := t.TempDir()
	writeArrayKernelBenchmarkProject(t, root, workload)
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	cfg := config.Default()
	configure(&cfg)
	recorder := analysisstats.NewRecorder()
	if _, err := (Analyzer{RootDir: root, Config: cfg}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder)); err != nil {
		t.Fatalf("analyze %s array fixture: %v", workload.name, err)
	}
	_, counters := recorder.Totals()
	values := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		values[counter.Name] = counter.Value
	}
	return values
}

func arrayKernelBenchmarkWorkloads() []arrayKernelBenchmarkWorkload {
	workloads := []arrayKernelBenchmarkWorkload{
		arrayKernelBenchmarkWorkload{
			name:          "no-arrays",
			source:        arrayBenchmarkNoArraysSource(),
			procedures:    1,
			arrayMarkers:  0,
			statementHint: 8,
		},
		arrayKernelBenchmarkWorkload{
			name:          "simple-dynamic",
			source:        arrayBenchmarkSimpleDynamicSource(),
			procedures:    1,
			arrayMarkers:  4,
			statementHint: 8,
		},
		arrayKernelBenchmarkWorkload{
			name:          "redim-heavy",
			source:        arrayBenchmarkRedimHeavySource(32),
			procedures:    1,
			arrayMarkers:  35,
			statementHint: 40,
		},
		arrayKernelBenchmarkWorkload{
			name:          "branch-heavy",
			source:        arrayBenchmarkBranchHeavySource(16),
			procedures:    1,
			arrayMarkers:  33,
			statementHint: 64,
		},
		arrayKernelBenchmarkWorkload{
			name:          "multidimensional",
			source:        arrayBenchmarkMultidimensionalSource(),
			procedures:    1,
			arrayMarkers:  8,
			statementHint: 12,
		},
		arrayKernelBenchmarkWorkload{
			name:          "byref-array-flow",
			source:        arrayBenchmarkByRefSource(),
			procedures:    2,
			arrayMarkers:  7,
			statementHint: 14,
		},
		arrayKernelBenchmarkWorkload{
			name:          "large-cfg",
			source:        arrayBenchmarkBranchHeavySource(96),
			procedures:    1,
			arrayMarkers:  193,
			statementHint: 384,
		},
	}
	for i := range workloads {
		workloads[i].lines = len(normalizedSourceLines(workloads[i].source)) - 1
	}
	return workloads
}

func configureArrayBenchmarkRulesNone(cfg *config.Config) {
	configureArrayBenchmarkRules(cfg, false)
}

func configureArrayBenchmarkRulesVBA227(cfg *config.Config) {
	configureArrayBenchmarkRules(cfg, false)
	cfg.Analyze.DetectArrayLifecycleSafety = true
}

func configureArrayBenchmarkRulesVBA209(cfg *config.Config) {
	configureArrayBenchmarkRules(cfg, false)
	cfg.Analyze.DetectObjectArrayComparison = true
}

func configureArrayBenchmarkRulesAll(cfg *config.Config) {
	configureArrayBenchmarkRules(cfg, true)
}

func configureArrayBenchmarkRules(cfg *config.Config, enabled bool) {
	cfg.Analyze.DetectRedimPreserveDimension = enabled
	cfg.Analyze.DetectObjectArrayComparison = enabled
	cfg.Analyze.DetectRangeValueArrayShape = enabled
	cfg.Analyze.DetectArrayLifecycleSafety = enabled
	cfg.Analyze.DetectRedimPreserveInLoops = enabled
	cfg.Analyze.DetectDeterministicRuntimeErrors = enabled
}

type arrayKernelBenchmarkFixture struct {
	source        string
	lines         int
	procedures    int
	arrayMarkers  int
	statementHint int
}

func writeArrayKernelBenchmarkProject(tb testing.TB, root string, workload arrayKernelBenchmarkWorkload) arrayKernelBenchmarkFixture {
	tb.Helper()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		tb.Fatalf("create array benchmark source directory: %v", err)
	}
	path := filepath.Join(modules, "Arrays.bas")
	if err := os.WriteFile(path, []byte(workload.source), 0o644); err != nil {
		tb.Fatalf("write %s array benchmark fixture: %v", workload.name, err)
	}
	return arrayKernelBenchmarkFixture{
		source:        path,
		lines:         len(normalizedSourceLines(workload.source)) - 1,
		procedures:    workload.procedures,
		arrayMarkers:  workload.arrayMarkers,
		statementHint: workload.statementHint,
	}
}

func arrayBenchmarkNoArraysSource() string {
	return `Option Explicit

Public Sub Run()
    Dim value As Long
    Dim branch As Long
    value = 1
    For branch = 1 To 8
        If value Mod 2 = 0 Then
            value = value + branch
        Else
            value = value - branch
        End If
    Next branch
    Debug.Print value
End Sub
`
}

func arrayBenchmarkSimpleDynamicSource() string {
	return `Option Explicit

Public Sub Run()
    Dim values() As Long
    ReDim values(0 To 31)
    values(1) = values(2) + 1
    If values(1) > 0 Then
        Debug.Print UBound(values)
    End If
End Sub
`
}

func arrayBenchmarkComparisonOnlySource() string {
	return `Option Explicit

Public Sub Run()
    Dim values() As Long
    If values = 1 Then
        Debug.Print 1
    End If
End Sub
`
}

func arrayBenchmarkObjectArrayOnlySource() string {
	return `Option Explicit

Public Sub Run()
    Dim values() As Object
    ReDim values(0 To 0)
    values(0) = Nothing
End Sub
`
}

func arrayBenchmarkRedimHeavySource(count int) string {
	var source strings.Builder
	source.WriteString("Option Explicit\n\nPublic Sub Run()\n    Dim values() As Long\n    Dim index As Long\n    ReDim values(0 To 1)\n    For index = 0 To 1\n        ReDim Preserve values(0 To index + 2)\n    Next index\n")
	for i := 0; i < count; i++ {
		fmt.Fprintf(&source, "    ReDim Preserve values(0 To %d)\n", i+2)
	}
	source.WriteString("    values(1) = UBound(values)\nEnd Sub\n")
	return source.String()
}

func arrayBenchmarkBranchHeavySource(branches int) string {
	var source strings.Builder
	source.WriteString("Option Explicit\n\nPublic Sub Run()\n    Dim values() As Long\n    Dim selector As Long\n    selector = 1\n")
	for i := 0; i < branches; i++ {
		if i == 0 {
			source.WriteString("    If selector = 0 Then\n")
		} else {
			source.WriteString("    ElseIf selector = 0 Then\n")
		}
		fmt.Fprintf(&source, "        ReDim values(0 To %d)\n", i+1)
	}
	source.WriteString("    Else\n        ReDim values(0 To 1)\n    End If\n    Debug.Print UBound(values)\nEnd Sub\n")
	return source.String()
}

func arrayBenchmarkMultidimensionalSource() string {
	return `Option Explicit

Public Sub Run()
    Dim matrix() As Double
    ReDim matrix(0 To 31, 0 To 15, 0 To 3)
    matrix(1, 2, 3) = matrix(2, 3, 1)
    ReDim Preserve matrix(0 To 31, 0 To 15, 0 To 7)
    Debug.Print LBound(matrix, 1), UBound(matrix, 3)
End Sub
`
}

func arrayBenchmarkByRefSource() string {
	return `Option Explicit

Private Sub Fill(ByRef values() As Long)
    ReDim values(0 To 31)
    values(0) = 1
End Sub

Public Sub Run()
    Dim values() As Long
    Fill values
    values(1) = values(0) + 1
    Debug.Print UBound(values)
End Sub
`
}
