package analyze

import (
	"context"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

// procedureLocalProfileDomains is the stable, aggregate-only profiling
// surface for the procedure-local analyzer. Keep this list in the test so a
// newly added domain cannot silently disappear from the developer profile.
var procedureLocalProfileDomains = []string{
	analysisstats.ProcedureLocalSourceScan,
	analysisstats.ProcedureLocalRuntime,
	analysisstats.ProcedureLocalArray,
	analysisstats.ProcedureLocalObject,
	analysisstats.ProcedureLocalDictionary,
	analysisstats.ProcedureLocalError,
	analysisstats.ProcedureLocalDataflow,
	analysisstats.ProcedureLocalResource,
	analysisstats.ProcedureLocalExcel,
	analysisstats.ProcedureLocalApplicationState,
	analysisstats.ProcedureLocalOther,
}

var procedureLocalProfileCounters = []string{
	analysisstats.RuntimeCandidateProceduresCounter,
	analysisstats.ArrayCandidateProceduresCounter,
	analysisstats.ObjectCandidateProceduresCounter,
	analysisstats.DictionaryCandidateProceduresCounter,
	analysisstats.ErrorCandidateProceduresCounter,
	analysisstats.DataflowCandidateProceduresCounter,
	analysisstats.ResourceCandidateProceduresCounter,
	analysisstats.ExcelCandidateProceduresCounter,
	analysisstats.ApplicationStateCandidateProceduresCounter,
	analysisstats.ArrayCFGWalksCounter,
	analysisstats.DataflowCFGWalksCounter,
	analysisstats.DictionaryCFGWalksCounter,
	analysisstats.ErrorCFGWalksCounter,
	analysisstats.ResourceCFGWalksCounter,
	analysisstats.ExcelCFGWalksCounter,
	analysisstats.RuntimeCFGWalksCounter,
	analysisstats.SourceLineScansCounter,
	analysisstats.SemanticKernelRunsCounter,
}

func TestPhysicalSourceLineCount(t *testing.T) {
	for _, test := range []struct {
		name   string
		source string
		want   int
	}{
		{name: "empty", source: "", want: 0},
		{name: "without terminal newline", source: "Option Explicit", want: 1},
		{name: "with terminal newline", source: "Option Explicit\n", want: 1},
		{name: "with trailing blank line", source: "Option Explicit\n\n", want: 2},
	} {
		t.Run(test.name, func(t *testing.T) {
			if got := physicalSourceLineCount(normalizedSourceLines(test.source)); got != test.want {
				t.Fatalf("physicalSourceLineCount(%q) = %d, want %d", test.source, got, test.want)
			}
		})
	}
}

func TestBatchAnalysisProfilingPreservesResultsAndReportsWorkload(t *testing.T) {
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	source := "Option Explicit\nPrivate moduleValue As Long\nPublic Sub Run(ByRef target As Long)\n  Dim value As Long\n  value = 1\n  target = value\nEnd Sub\n"
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
		"project_context", "project_context_indexes", "typedb_load", "project_symbols", "project_wide_diagnostics",
		"procedure_local_diagnostics", "typed_excel_diagnostics", "byref_diagnostics", "compile_equivalent_diagnostics",
		"suppression_and_finalize", "analyze_total",
	} {
		stage, ok := stageByName[name]
		if !ok {
			t.Fatalf("missing stage %q: %+v", name, stages)
		}
		if stage.Outcome != "ok" || stage.Calls < 1 {
			t.Fatalf("stage %q = %+v", name, stage)
		}
	}
	validDomains := make(map[string]bool, len(procedureLocalProfileDomains))
	for _, name := range procedureLocalProfileDomains {
		validDomains[name] = true
	}
	for _, stage := range stages {
		if !strings.HasPrefix(stage.Name, "procedure_local/") {
			continue
		}
		if !validDomains[stage.Name] {
			t.Fatalf("unknown procedure-local domain stage %q", stage.Name)
		}
		if stage.Calls < 1 || stage.Outcome != "ok" {
			t.Fatalf("procedure-local domain stage %q = %+v", stage.Name, stage)
		}
	}
	for _, name := range []string{
		analysisstats.ProcedureLocalSourceScan,
		analysisstats.ProcedureLocalRuntime,
		analysisstats.ProcedureLocalDictionary,
		analysisstats.ProcedureLocalOther,
	} {
		if _, ok := stageByName[name]; !ok {
			t.Fatalf("missing representative procedure-local domain stage %q: %+v", name, stages)
		}
	}
	counterByName := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		counterByName[counter.Name] = counter.Value
	}
	for _, name := range []string{
		"file_count", "procedure_count", "statement_count", "expression_count",
		"call_site_count", "cfg_block_count", "cfg_edge_count", "project_symbol_count",
		"byref_diagnostic_passes", "line_count", "module_declaration_count",
		analysisstats.ModuleFactBuildsCounter, analysisstats.ProcedureFactBuildsCounter,
		"max_lines_per_file", "max_procedures_per_file", "max_calls_per_file",
		"max_statements_per_procedure", "max_cfg_blocks_per_procedure", "max_cfg_edges_per_procedure",
	} {
		if _, ok := counterByName[name]; !ok {
			t.Fatalf("missing counter %q: %+v", name, counters)
		}
	}
	if counterByName[analysisstats.CapabilityResolutionBuildsCounter] != 1 {
		t.Fatalf("resolution capability builds = %d, want one", counterByName[analysisstats.CapabilityResolutionBuildsCounter])
	}
	validCounters := make(map[string]bool, len(procedureLocalProfileCounters))
	for _, name := range procedureLocalProfileCounters {
		validCounters[name] = true
	}
	for name := range counterByName {
		if validCounters[name] {
			continue
		}
		looksLikeProcedureLocalCounter := strings.HasSuffix(name, "_candidate_procedures") ||
			strings.HasSuffix(name, "_cfg_walks") ||
			name == analysisstats.SourceLineScansCounter ||
			name == analysisstats.SemanticKernelRunsCounter
		if looksLikeProcedureLocalCounter {
			t.Fatalf("unknown procedure-local work counter %q", name)
		}
	}
	if counterByName[analysisstats.SourceLineScansCounter] == 0 || counterByName[analysisstats.SemanticKernelRunsCounter] == 0 {
		t.Fatalf("procedure-local traversal counters = %+v", counters)
	}
	if counterByName["file_count"] != 1 || counterByName["procedure_count"] != 1 {
		t.Fatalf("workload counters = %+v", counters)
	}
	if counterByName[analysisstats.ModuleFactBuildsCounter] != 1 || counterByName[analysisstats.ProcedureFactBuildsCounter] != 1 {
		t.Fatalf("shared fact builds = module %d, procedure %d; want one each for one file/procedure revision", counterByName[analysisstats.ModuleFactBuildsCounter], counterByName[analysisstats.ProcedureFactBuildsCounter])
	}
	wantLineCount := uint64(len(normalizedSourceLines(source)) - 1)
	if counterByName["line_count"] != wantLineCount ||
		counterByName["max_lines_per_file"] != wantLineCount ||
		counterByName["module_declaration_count"] != 1 ||
		counterByName["max_procedures_per_file"] != 1 ||
		counterByName["max_calls_per_file"] != 0 {
		t.Fatalf("line/module maximum counters = %+v", counters)
	}
	if counterByName["max_statements_per_procedure"] == 0 ||
		counterByName["max_cfg_blocks_per_procedure"] == 0 ||
		counterByName["max_cfg_edges_per_procedure"] == 0 {
		t.Fatalf("procedure maximum counters = %+v", counters)
	}
	if counterByName["byref_diagnostic_passes"] != 1 {
		t.Fatalf("ByRef analysis passes = %d, want one per file revision", counterByName["byref_diagnostic_passes"])
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
	for _, name := range []string{
		"object_summary_evaluations", "object_entry_flow_evaluations",
	} {
		value, ok := counterByName[name]
		if ok && value != 0 {
			t.Fatalf("disabled VBA202 counter %q = %d (present=%t), want zero", name, value, ok)
		}
	}
}

func TestBatchAnalysisDoesNotCountDisabledRuntimeCandidates(t *testing.T) {
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(modules, "Main.bas"), []byte("Option Explicit\nPublic Sub Run()\n  Debug.Print 1 / 0\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	cfg := config.Default()
	cfg.Analyze.DetectDeterministicRuntimeErrors = false
	recorder := analysisstats.NewRecorder()
	if _, err := (Analyzer{RootDir: root, Config: cfg}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder)); err != nil {
		t.Fatal(err)
	}
	_, counters := recorder.Totals()
	for _, counter := range counters {
		if counter.Name == analysisstats.RuntimeCandidateProceduresCounter || counter.Name == analysisstats.RuntimeCFGWalksCounter {
			t.Fatalf("disabled runtime counter %q = %d, want absent", counter.Name, counter.Value)
		}
	}
}

func TestBatchAnalysisProfilesHTTPOnlyDataflowWork(t *testing.T) {
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	source := "Option Explicit\nPublic Sub Run()\n  Dim request As Object\n  Set request = CreateObject(\"MSXML2.XMLHTTP\")\n  request.Open \"GET\", \"http://example.test\", False\n  request.Send\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(modules, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, filepath.Join(t.TempDir(), "typelib"))
	cfg := config.Default()
	cfg.Analyze.DetectUntrustedDataFlow = false
	cfg.Analyze.DetectUnsafeCommandConstruction = false
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	cfg.Analyze.DetectUnsafeHTTPConfiguration = true
	cfg.Analyze.DetectMissingHTTPTimeout = true
	recorder := analysisstats.NewRecorder()
	if _, err := (Analyzer{RootDir: root, Config: cfg}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder)); err != nil {
		t.Fatal(err)
	}
	stages, counters := recorder.Totals()
	stageByName := make(map[string]analysisstats.Stage, len(stages))
	for _, stage := range stages {
		stageByName[stage.Name] = stage
	}
	if stageByName[analysisstats.ProcedureLocalDataflow].Calls == 0 {
		t.Fatalf("HTTP-only dataflow stage missing: %+v", stages)
	}
	counterByName := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		counterByName[counter.Name] = counter.Value
	}
	for _, name := range []string{
		analysisstats.DataflowCandidateProceduresCounter,
		analysisstats.DataflowCFGWalksCounter,
		analysisstats.SemanticKernelRunsCounter,
	} {
		if counterByName[name] == 0 {
			t.Fatalf("HTTP-only dataflow counter %q = %d; counters = %+v", name, counterByName[name], counters)
		}
	}
	if counterByName[analysisstats.CapabilityDataflowBuildsCounter] != 1 {
		t.Fatalf("dataflow capability builds = %d, want one", counterByName[analysisstats.CapabilityDataflowBuildsCounter])
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
	if values[analysisstats.ModuleFactBuildsCounter] != 1 {
		t.Fatalf("module fact builds = %d, want one for the file revision", values[analysisstats.ModuleFactBuildsCounter])
	}
	// The fixture contains the dependency chain plus its Run entry point.
	wantProcedureFacts := uint64(procedureCount + 1)
	if values[analysisstats.ProcedureFactBuildsCounter] != wantProcedureFacts {
		t.Fatalf("procedure fact builds = %d, want one per procedure revision (%d procedures)", values[analysisstats.ProcedureFactBuildsCounter], wantProcedureFacts)
	}
	limit := uint64((procedureCount + 1) * 5)
	if values["object_summary_evaluations"] > limit {
		t.Fatalf("summary evaluations = %d, want no more than %d for %d procedures", values["object_summary_evaluations"], limit, procedureCount+1)
	}
	if values["object_entry_flow_evaluations"] > limit {
		t.Fatalf("entry flow evaluations = %d, want no more than %d for %d procedures", values["object_entry_flow_evaluations"], limit, procedureCount+1)
	}
}
