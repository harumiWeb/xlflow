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
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// BenchmarkSyntheticProject measures batch analysis as the number of project
// procedures grows. The source is generated before the timer starts so these
// measurements describe analysis work rather than fixture construction.
func BenchmarkSyntheticProject(b *testing.B) {
	for _, procedureCount := range []int{100, 500, 1000} {
		procedureCount := procedureCount
		b.Run(fmt.Sprintf("%d-procedures", procedureCount), func(b *testing.B) {
			root := b.TempDir()
			writeSyntheticBenchmarkProject(b, root, procedureCount)

			// Keep the benchmark independent of any developer-generated TypeLib
			// database. A missing database is a valid, deterministic analyzer
			// configuration for this fixture.
			b.Setenv(typedb.EnvDir, filepath.Join(b.TempDir(), "typelib"))
			cfg := config.Default()
			recorder := analysisstats.NewRecorder()
			ctx := analysisstats.WithRecorder(context.Background(), recorder)

			b.ReportAllocs()
			b.ResetTimer()
			findings := 0
			warnings := 0
			for i := 0; i < b.N; i++ {
				result, err := (Analyzer{RootDir: root, Config: cfg}).RunResultContext(ctx)
				if err != nil {
					b.Fatalf("analyze synthetic %d-procedure project: %v", procedureCount, err)
				}
				findings += len(result.Findings)
				warnings += len(result.Warnings)
			}
			b.StopTimer()

			b.ReportMetric(float64(procedureCount), "procedures/op")
			b.ReportMetric(float64(findings)/float64(b.N), "findings/op")
			b.ReportMetric(float64(warnings)/float64(b.N), "warnings/op")
			reportAnalysisRecorderMetrics(b, recorder, b.N)
		})
	}
}

// BenchmarkSingleModuleSynthetic measures the pathological single-file shapes
// that do not get represented by project-scale fixtures with many modules.
// Every source file and its workload metadata are prepared before the timer
// starts so the timed region contains only analyzer work.
func BenchmarkSingleModuleSynthetic(b *testing.B) {
	for _, workload := range singleModuleBenchmarkWorkloads() {
		workload := workload
		b.Run(workload.benchmarkName(), func(b *testing.B) {
			root := b.TempDir()
			fixture := writeSingleModuleBenchmarkProject(b, root, workload)

			// Keep the benchmark independent of any developer-generated TypeLib
			// database. A missing database is a valid, deterministic analyzer
			// configuration for this fixture.
			b.Setenv(typedb.EnvDir, filepath.Join(b.TempDir(), "typelib"))
			cfg := config.Default()
			recorder := analysisstats.NewRecorder()
			recordSingleModuleBenchmarkDimensions(recorder, fixture)
			ctx := analysisstats.WithRecorder(context.Background(), recorder)

			b.ReportAllocs()
			b.ResetTimer()
			findings := 0
			warnings := 0
			for i := 0; i < b.N; i++ {
				result, err := (Analyzer{RootDir: root, Config: cfg}).RunResultContext(ctx)
				if err != nil {
					b.Fatalf("analyze single-module %s fixture: %v", workload.benchmarkName(), err)
				}
				findings += len(result.Findings)
				warnings += len(result.Warnings)
			}
			b.StopTimer()

			b.ReportMetric(float64(fixture.lines), "lines/op")
			b.ReportMetric(float64(fixture.procedures), "procedures/op")
			b.ReportMetric(float64(fixture.moduleDeclarations), "module-declarations/op")
			b.ReportMetric(float64(fixture.calls), "calls/op")
			b.ReportMetric(float64(findings)/float64(b.N), "findings/op")
			b.ReportMetric(float64(warnings)/float64(b.N), "warnings/op")
			reportAnalysisRecorderMetrics(b, recorder, b.N)
		})
	}
}

// BenchmarkObjectAnalysisWorklist isolates the object-summary and entry-state
// stages on a deliberately reverse-ordered object-return chain. The source is
// parsed and its IR/CFG are constructed before timing starts so the benchmark
// measures propagation rather than fixture construction or parsing.
func BenchmarkObjectAnalysisWorklist(b *testing.B) {
	for _, procedureCount := range []int{100, 500, 1000} {
		procedureCount := procedureCount
		b.Run(fmt.Sprintf("%d-procedures", procedureCount), func(b *testing.B) {
			root := b.TempDir()
			writeObjectWorklistBenchmarkProject(b, root, procedureCount)
			files := loadObjectWorklistBenchmarkFiles(b, root)
			recorder := analysisstats.NewRecorder()

			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				analysis := buildObjectAnalysisPlans(files)
				analysis.buildSummaries()
				analysis.buildEntryStates()
				recorder.Add("object_summary_evaluations", uint64(analysis.summaryEvaluations))
				recorder.Add("object_entry_flow_evaluations", uint64(analysis.entryFlowEvaluations))
			}
			b.StopTimer()
			b.ReportMetric(float64(procedureCount), "procedures/op")
			reportAnalysisRecorderMetrics(b, recorder, b.N)
		})
	}
}

const objectWorklistBenchmarkModuleFile = "Chain.bas"

func loadObjectWorklistBenchmarkFiles(tb testing.TB, root string) []parsedFile {
	tb.Helper()
	path := filepath.Join(root, "src", "modules", objectWorklistBenchmarkModuleFile)
	source, err := os.ReadFile(path)
	if err != nil {
		tb.Fatalf("read object benchmark source: %v", err)
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		tb.Fatalf("parse object benchmark source: %v", err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{RootDir: root, Path: path, ModuleKind: "standard"}, doc)
	if err != nil {
		tb.Fatalf("build object benchmark IR: %v", err)
	}
	module := strings.TrimSpace(ir.ModuleName)
	if module == "" {
		module = "Chain"
	}
	resolverSymbols := make([]procedureir.ResolverSymbol, 0, len(ir.Procedures))
	for _, procedure := range ir.Procedures {
		resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
			Name: procedure.Symbol.Name, Type: procedure.Symbol.ReturnType,
			Module: module, ModuleKind: ir.ModuleKind, Kind: string(procedure.Symbol.Kind),
			Visibility: procedure.Symbol.Visibility, File: path,
			Line: procedure.Symbol.DeclarationRange.StartLine,
		})
	}
	ir = procedureir.Resolve(ir, procedureir.NewResolver(resolverSymbols))
	controlFlow := vbacfg.BuildDocument(ir)
	file := parsedFile{
		Path:       path,
		Lines:      normalizedSourceLines(string(source)),
		Module:     module,
		ModuleKind: "standard",
		Source:     source,
		IR:         ir,
		CFG:        controlFlow,
	}
	return []parsedFile{file}
}

func writeObjectWorklistBenchmarkProject(tb testing.TB, root string, procedureCount int) {
	tb.Helper()
	if procedureCount < 2 {
		tb.Fatalf("object worklist benchmark requires at least two procedures, got %d", procedureCount)
	}
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		tb.Fatalf("create object benchmark source directory: %v", err)
	}
	var source strings.Builder
	source.WriteString("Option Explicit\n\n")
	for index := 0; index < procedureCount; index++ {
		fmt.Fprintf(&source, "Private Function Proc%03d() As Worksheet\n", index)
		if index+1 < procedureCount {
			fmt.Fprintf(&source, "  Set Proc%03d = Proc%03d()\n", index, index+1)
		} else {
			fmt.Fprintf(&source, "  Set Proc%03d = ThisWorkbook.Worksheets(1)\n", index)
		}
		source.WriteString("End Function\n\n")
	}
	source.WriteString("Public Sub Run()\n  Dim target As Worksheet\n  Set target = Proc000()\n  Debug.Print target.Name\nEnd Sub\n")
	if err := os.WriteFile(filepath.Join(modules, objectWorklistBenchmarkModuleFile), []byte(source.String()), 0o644); err != nil {
		tb.Fatalf("write object benchmark source: %v", err)
	}
}

// TestSyntheticBenchmarkFixtureScale keeps the benchmark's workload contract
// visible to ordinary tests. The benchmark itself remains free of timing
// assertions so it can run on different developer machines and CI runners.
func TestSyntheticBenchmarkFixtureScale(t *testing.T) {
	for _, procedureCount := range []int{100, 500, 1000} {
		root := t.TempDir()
		modules := writeSyntheticBenchmarkProject(t, root, procedureCount)
		if got := syntheticProcedureCount(modules); got != procedureCount {
			t.Fatalf("synthetic %d-procedure fixture contains %d procedures", procedureCount, got)
		}
		if len(modules) < 2 {
			t.Fatalf("synthetic %d-procedure fixture uses %d modules, want multiple modules", procedureCount, len(modules))
		}
	}
}

// TestSingleModuleBenchmarkFixtureScale keeps the single-file workload matrix
// visible to ordinary tests without running the expensive analyzer benchmark.
func TestSingleModuleBenchmarkFixtureScale(t *testing.T) {
	for _, workload := range singleModuleBenchmarkWorkloads() {
		workload := workload
		t.Run(workload.benchmarkName(), func(t *testing.T) {
			fixture := writeSingleModuleBenchmarkProject(t, t.TempDir(), workload)
			if fixture.sourcePath == "" || fixture.lines == 0 {
				t.Fatalf("single-module %s fixture has no source metadata: %+v", workload.benchmarkName(), fixture)
			}
			if fixture.procedures == 0 {
				t.Fatalf("single-module %s fixture has no procedures", workload.benchmarkName())
			}
			if fixture.moduleCount != 1 {
				t.Fatalf("single-module %s fixture contains %d modules, want 1", workload.benchmarkName(), fixture.moduleCount)
			}
			if fixture.procedures != workload.expectedProcedures() {
				t.Fatalf("single-module %s procedures = %d, want %d", workload.benchmarkName(), fixture.procedures, workload.expectedProcedures())
			}
			if fixture.moduleDeclarations != workload.expectedModuleDeclarations() {
				t.Fatalf("single-module %s declarations = %d, want %d", workload.benchmarkName(), fixture.moduleDeclarations, workload.expectedModuleDeclarations())
			}
			if fixture.calls != workload.expectedCalls() {
				t.Fatalf("single-module %s calls = %d, want %d", workload.benchmarkName(), fixture.calls, workload.expectedCalls())
			}
		})
	}
}

type singleModuleBenchmarkWorkload struct {
	shape string
	size  int
}

func singleModuleBenchmarkWorkloads() []singleModuleBenchmarkWorkload {
	workloads := make([]singleModuleBenchmarkWorkload, 0, 17)
	for _, shape := range []string{"independent", "chain", "declarations"} {
		for _, size := range []int{500, 1000, 2000} {
			workloads = append(workloads, singleModuleBenchmarkWorkload{shape: shape, size: size})
		}
	}
	for _, shape := range []string{"dense", "cyclic"} {
		for _, size := range []int{500, 1000} {
			workloads = append(workloads, singleModuleBenchmarkWorkload{shape: shape, size: size})
		}
	}
	for _, size := range []int{100, 500, 1000, 2000} {
		workloads = append(workloads, singleModuleBenchmarkWorkload{shape: "cfg-independent", size: size})
	}
	return workloads
}

func (w singleModuleBenchmarkWorkload) benchmarkName() string {
	unit := "procedures"
	switch w.shape {
	case "declarations":
		unit = "declarations"
	case "cfg-independent":
		unit = "branches"
	}
	return fmt.Sprintf("%s/%d-%s", w.shape, w.size, unit)
}

func (w singleModuleBenchmarkWorkload) expectedProcedures() int {
	switch w.shape {
	case "declarations", "cfg-independent":
		return 1
	default:
		return w.size
	}
}

func (w singleModuleBenchmarkWorkload) expectedModuleDeclarations() int {
	switch w.shape {
	case "declarations":
		return w.size
	default:
		return 0
	}
}

func (w singleModuleBenchmarkWorkload) expectedCalls() int {
	switch w.shape {
	case "chain":
		return w.size - 1
	case "dense":
		const fanout = 8
		calls := 0
		for index := 0; index < w.size; index++ {
			remaining := w.size - index - 1
			if remaining > fanout {
				remaining = fanout
			}
			calls += remaining
		}
		return calls
	case "cyclic":
		return w.size
	default:
		return 0
	}
}

type singleModuleBenchmarkFixture struct {
	sourcePath         string
	moduleCount        int
	lines              int
	procedures         int
	moduleDeclarations int
	calls              int
	maxStatements      int
	maxCFGBlocks       int
	maxCFGEdges        int
}

func writeSingleModuleBenchmarkProject(tb testing.TB, root string, workload singleModuleBenchmarkWorkload) singleModuleBenchmarkFixture {
	tb.Helper()
	srcDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(srcDir, 0o755); err != nil {
		tb.Fatalf("create single-module benchmark source directory: %v", err)
	}
	path := filepath.Join(srcDir, "Monster.bas")
	source := singleModuleBenchmarkSource(workload)
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		tb.Fatalf("write single-module %s fixture: %v", workload.benchmarkName(), err)
	}
	entries, err := os.ReadDir(srcDir)
	if err != nil {
		tb.Fatalf("read single-module benchmark source directory: %v", err)
	}

	fixture := singleModuleBenchmarkFixture{sourcePath: path, moduleCount: len(entries)}
	inspectSingleModuleBenchmarkFixture(tb, root, path, source, &fixture)
	return fixture
}

func singleModuleBenchmarkSource(workload singleModuleBenchmarkWorkload) string {
	var source strings.Builder
	source.WriteString("Option Explicit\n\n")
	switch workload.shape {
	case "independent":
		for index := 0; index < workload.size; index++ {
			fmt.Fprintf(&source, "Private Sub Independent%04d()\n    Dim value As Long\n    value = %d\nEnd Sub\n\n", index, index)
		}
	case "chain":
		for index := 0; index < workload.size; index++ {
			fmt.Fprintf(&source, "Private Sub Chain%04d()\n", index)
			if index+1 < workload.size {
				fmt.Fprintf(&source, "    Call Chain%04d\n", index+1)
			} else {
				source.WriteString("    Dim value As Long\n    value = 1\n")
			}
			source.WriteString("End Sub\n\n")
		}
	case "dense":
		const fanout = 8
		for index := 0; index < workload.size; index++ {
			fmt.Fprintf(&source, "Private Sub Dense%04d()\n", index)
			for offset := 1; offset <= fanout && index+offset < workload.size; offset++ {
				fmt.Fprintf(&source, "    Call Dense%04d\n", index+offset)
			}
			if index+1 >= workload.size {
				source.WriteString("    Dim value As Long\n    value = 1\n")
			}
			source.WriteString("End Sub\n\n")
		}
	case "cyclic":
		for index := 0; index < workload.size; index++ {
			fmt.Fprintf(&source, "Private Sub Cyclic%04d()\n    Call Cyclic%04d\nEnd Sub\n\n", index, (index+1)%workload.size)
		}
	case "declarations":
		for index := 0; index < workload.size; index++ {
			fmt.Fprintf(&source, "Private Const ModuleConst%04d As Long = %d\n", index, index)
		}
		source.WriteString("\nPrivate Sub ReadModuleDeclaration()\n    Dim value As Long\n    value = ModuleConst0000\nEnd Sub\n")
	case "cfg-independent":
		source.WriteString("Private Sub IndependentLargeCFG(ByRef value As Long)\n")
		for index := 0; index < workload.size; index++ {
			fmt.Fprintf(&source, "    If value = %d Then\n        value = value + 1\n    Else\n        value = value - 1\n    End If\n", index)
		}
		source.WriteString("End Sub\n")
	}
	return source.String()
}

func inspectSingleModuleBenchmarkFixture(tb testing.TB, root, path, source string, fixture *singleModuleBenchmarkFixture) {
	tb.Helper()
	doc, err := vbaast.ParseDocument(path, []byte(source))
	if err != nil {
		tb.Fatalf("parse single-module benchmark source: %v", err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{RootDir: root, Path: path, ModuleKind: "standard"}, doc)
	if err != nil {
		tb.Fatalf("build single-module benchmark IR: %v", err)
	}
	fixture.lines = physicalSourceLineCount(normalizedSourceLines(source))
	fixture.procedures = len(ir.Procedures)
	for _, declaration := range ir.Declarations {
		if declaration.Scope == procedureir.ScopeModule || declaration.Scope == procedureir.ScopeProject {
			fixture.moduleDeclarations++
		}
	}
	for _, procedure := range ir.Procedures {
		if len(procedure.Statements) > fixture.maxStatements {
			fixture.maxStatements = len(procedure.Statements)
		}
		fixture.calls += len(procedure.Calls)
	}
	controlFlow := vbacfg.BuildDocument(ir)
	for _, graph := range controlFlow.Graphs {
		if len(graph.Blocks) > fixture.maxCFGBlocks {
			fixture.maxCFGBlocks = len(graph.Blocks)
		}
		if len(graph.Edges) > fixture.maxCFGEdges {
			fixture.maxCFGEdges = len(graph.Edges)
		}
	}
}

func recordSingleModuleBenchmarkDimensions(recorder *analysisstats.Recorder, fixture singleModuleBenchmarkFixture) {
	if recorder == nil {
		return
	}
	recorder.AddMax("max_lines_per_file", uint64(fixture.lines))
	recorder.AddMax("max_procedures_per_file", uint64(fixture.procedures))
	recorder.AddMax("max_calls_per_file", uint64(fixture.calls))
	recorder.AddMax("max_statements_per_procedure", uint64(fixture.maxStatements))
	recorder.AddMax("max_cfg_blocks_per_procedure", uint64(fixture.maxCFGBlocks))
	recorder.AddMax("max_cfg_edges_per_procedure", uint64(fixture.maxCFGEdges))
}

type syntheticBenchmarkModule struct {
	name       string
	source     string
	procedures int
}

func writeSyntheticBenchmarkProject(tb testing.TB, root string, procedureCount int) []syntheticBenchmarkModule {
	tb.Helper()
	if procedureCount < 2 {
		tb.Fatalf("synthetic benchmark requires at least two procedures, got %d", procedureCount)
	}

	modules := syntheticBenchmarkModules(procedureCount)
	srcDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(srcDir, 0o755); err != nil {
		tb.Fatalf("create synthetic source directory: %v", err)
	}
	for _, module := range modules {
		if err := os.WriteFile(filepath.Join(srcDir, module.name), []byte(module.source), 0o644); err != nil {
			tb.Fatalf("write synthetic module %q: %v", module.name, err)
		}
	}
	return modules
}

func syntheticBenchmarkModules(procedureCount int) []syntheticBenchmarkModule {
	modules := []syntheticBenchmarkModule{
		{
			name:       "Shared.bas",
			procedures: 1,
			source: `Option Explicit

Public Sub SharedTransform(ByRef value As Long, ByRef worker As Object, ByRef text As String)
    Dim branch As Long
    If worker Is Nothing Then
        Set worker = CreateObject("Scripting.Dictionary")
    End If
    For branch = 1 To 3
        If (value + branch) Mod 2 = 0 Then
            value = value + branch
        Else
            value = value - branch
        End If
    Next branch
    worker.Item(text) = value
End Sub
`,
		},
	}

	remaining := procedureCount - 1
	moduleNumber := 1
	for remaining > 0 {
		count := remaining
		if count > 50 {
			count = 50
		}
		var source strings.Builder
		source.WriteString("Option Explicit\n\n")
		for i := 0; i < count; i++ {
			procedureNumber := procedureCount - remaining + i + 1
			fmt.Fprintf(&source, `Public Sub Proc%04d(ByRef value As Long, ByRef worker As Object, ByRef text As String)
    Dim local As Long
    Dim branch As Long
    local = value + %d
    For branch = 1 To 4
        If local Mod 2 = 0 Then
            local = local + branch
        ElseIf local > branch Then
            local = local - branch
        Else
            local = local + 1
        End If
    Next branch
    Call SharedTransform(local, worker, text)
    value = local
End Sub

`, procedureNumber, procedureNumber)
		}
		modules = append(modules, syntheticBenchmarkModule{
			name:       fmt.Sprintf("Module%02d.bas", moduleNumber),
			source:     source.String(),
			procedures: count,
		})
		remaining -= count
		moduleNumber++
	}
	return modules
}

func syntheticProcedureCount(modules []syntheticBenchmarkModule) int {
	count := 0
	for _, module := range modules {
		count += module.procedures
	}
	return count
}

func reportAnalysisRecorderMetrics(b *testing.B, recorder *analysisstats.Recorder, iterations int) {
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
