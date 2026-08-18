package analyze

import (
	"context"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

func TestBatchAnalysisProfilingPreservesResultsAndReportsWorkload(t *testing.T) {
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	source := "Option Explicit\nPublic Sub Run(ByRef target As Long)\n  Dim value As Long\n  value = 1\n  target = value\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(modules, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	cfg := config.Default()
	analyzer := Analyzer{RootDir: root, Config: cfg}

	want, err := analyzer.RunResultContext(context.Background())
	if err != nil {
		t.Fatal(err)
	}
	recorder := analysisstats.NewRecorder()
	got, err := analyzer.RunResultContext(analysisstats.WithRecorder(context.Background(), recorder))
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("profiled result changed:\n got: %#v\nwant: %#v", got, want)
	}

	stages, counters := recorder.Totals()
	stageByName := make(map[string]analysisstats.Stage, len(stages))
	for _, stage := range stages {
		stageByName[stage.Name] = stage
	}
	for _, name := range []string{
		"source_discovery", "file_read", "parse", "procedure_ir", "cfg",
		"effect_summaries", "object_procedure_summaries", "object_entry_states",
		"project_context", "typedb_load", "project_symbols", "project_wide_diagnostics",
		"file_procedure_diagnostics", "byref_diagnostics", "compile_equivalent_diagnostics",
		"suppression_finalization", "analyze_total",
	} {
		stage, ok := stageByName[name]
		if !ok {
			t.Fatalf("missing stage %q: %+v", name, stages)
		}
		if stage.Outcome != "ok" || stage.Calls < 1 {
			t.Fatalf("stage %q = %+v", name, stage)
		}
	}
	counterByName := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		counterByName[counter.Name] = counter.Value
	}
	for _, name := range []string{
		"file_count", "procedure_count", "statement_count", "expression_count",
		"call_site_count", "cfg_block_count", "cfg_edge_count", "project_symbol_count",
	} {
		if _, ok := counterByName[name]; !ok {
			t.Fatalf("missing counter %q: %+v", name, counters)
		}
	}
	if counterByName["file_count"] != 1 || counterByName["procedure_count"] != 1 {
		t.Fatalf("workload counters = %+v", counters)
	}
}
func TestBatchAnalysisSkipsVBA202ContextWhenDisabled(t *testing.T) {
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	source := "Option Explicit\nPublic Sub Run()\n  Dim target As Worksheet\n  Debug.Print target.Name\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(modules, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	cfg := config.Default()
	cfg.Analyze.DetectObjectUseBeforeSet = false
	recorder := analysisstats.NewRecorder()
	result, err := (Analyzer{RootDir: root, Config: cfg}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(result.Findings, "VBA202"); len(got) != 0 {
		t.Fatalf("disabled VBA202 produced findings: %+v", got)
	}
	stages, counters := recorder.Totals()
	byName := make(map[string]analysisstats.Stage, len(stages))
	for _, stage := range stages {
		byName[stage.Name] = stage
	}
	for _, name := range []string{"object_procedure_summaries", "object_entry_states"} {
		stage, ok := byName[name]
		if !ok {
			t.Fatalf("missing disabled VBA202 stage %q: %+v", name, stages)
		}
		if stage.ResultCount != 0 {
			t.Fatalf("disabled VBA202 stage %q performed object analysis: %+v", name, stage)
		}
	}
	counterByName := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		counterByName[counter.Name] = counter.Value
	}
	for _, name := range []string{"object_summary_evaluations", "object_entry_flow_evaluations"} {
		value, ok := counterByName[name]
		if !ok || value != 0 {
			t.Fatalf("disabled VBA202 counter %q = %d (present=%t), want zero", name, value, ok)
		}
	}
}

func TestVBA202WorklistEvaluationCountScalesWithDependencyChain(t *testing.T) {
	const procedureCount = 40
	root := t.TempDir()
	writeObjectWorklistBenchmarkProject(t, root, procedureCount)
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	recorder := analysisstats.NewRecorder()
	result, err := (Analyzer{RootDir: root, Config: config.Default()}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(result.Findings, "VBA202"); len(got) != 0 {
		t.Fatalf("resolved object-return chain produced VBA202 findings: %+v", got)
	}
	repeated, err := (Analyzer{RootDir: root, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(repeated, result.Findings) {
		t.Fatalf("worklist result is not deterministic:\nfirst: %#v\nsecond: %#v", result.Findings, repeated)
	}
	_, counters := recorder.Totals()
	values := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		values[counter.Name] = counter.Value
	}
	for _, name := range []string{"object_summary_evaluations", "object_entry_flow_evaluations"} {
		if _, ok := values[name]; !ok {
			t.Fatalf("missing worklist evaluation counter %q: %+v", name, counters)
		}
		if values[name] < procedureCount {
			t.Fatalf("%s = %d, want at least one evaluation per procedure", name, values[name])
		}
	}
	limit := uint64((procedureCount + 1) * 5)
	if values["object_summary_evaluations"] > limit {
		t.Fatalf("summary evaluations = %d, want no more than %d for %d procedures", values["object_summary_evaluations"], limit, procedureCount+1)
	}
	if values["object_entry_flow_evaluations"] > limit {
		t.Fatalf("entry flow evaluations = %d, want no more than %d for %d procedures", values["object_entry_flow_evaluations"], limit, procedureCount+1)
	}
}
