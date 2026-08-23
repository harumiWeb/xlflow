package analysisstats

import (
	"context"
	"errors"
	"reflect"
	"sync"
	"testing"
	"time"
)

func TestMeasureWaitPopulatesWaitDuration(t *testing.T) {
	recorder := NewRecorder()
	ctx := WithRecorder(context.Background(), recorder)
	finish := MeasureWait(ctx, "procedure_ir_singleflight")
	time.Sleep(time.Millisecond)
	finish(nil)

	stages, _ := recorder.Snapshot()
	if len(stages) != 1 || stages[0].Name != "procedure_ir_singleflight" || stages[0].Wait <= 0 || stages[0].Elapsed != 0 || stages[0].Outcome != "ok" {
		t.Fatalf("wait stage = %+v", stages)
	}
}

func TestMeasureWithoutRecorderIsNoOp(t *testing.T) {
	finish := Measure(context.Background(), "parse")
	finish(1, nil)
	finishWait := MeasureWait(context.Background(), "procedure_ir_singleflight")
	finishWait(nil)
}

func TestTotalsAggregateStagesAndSortCounters(t *testing.T) {
	recorder := NewRecorder()
	recorder.Record(Stage{Name: "parse", Elapsed: time.Millisecond, Outcome: "ok", ResultCount: 1})
	recorder.Record(Stage{Name: "cfg", Elapsed: 2 * time.Millisecond, Outcome: "ok", ResultCount: 2})
	recorder.Record(Stage{Name: "parse", Elapsed: 3 * time.Millisecond, Outcome: "error", ResultCount: 4})
	recorder.Add("procedure_count", 3)
	recorder.AddSum("file_count", 1)
	recorder.AddSum("procedure_count", 2)
	recorder.AddMax("max_lines_per_file", 11)
	recorder.AddMax("max_lines_per_file", 7)
	recorder.AddMax("max_procedures_per_file", 2)
	recorder.AddMax("max_procedures_per_file", 5)

	stages, counters := recorder.Totals()
	if len(stages) != 2 {
		t.Fatalf("stage totals = %+v", stages)
	}
	if got := stages[0]; got.Name != "parse" || got.Elapsed != 4*time.Millisecond || got.ResultCount != 5 || got.Calls != 2 || got.Outcome != "error" {
		t.Fatalf("parse total = %+v", got)
	}
	if got := stages[1]; got.Name != "cfg" || got.Calls != 1 || got.Outcome != "ok" {
		t.Fatalf("cfg total = %+v", got)
	}
	wantCounters := []Counter{
		{Name: "file_count", Value: 1},
		{Name: "max_lines_per_file", Value: 11},
		{Name: "max_procedures_per_file", Value: 5},
		{Name: "procedure_count", Value: 5},
	}
	if !reflect.DeepEqual(counters, wantCounters) {
		t.Fatalf("counters = %+v, want %+v", counters, wantCounters)
	}
}

func TestRecordFactBuildCounters(t *testing.T) {
	recorder := NewRecorder()
	recorder.RecordModuleFactBuild()
	recorder.RecordProcedureFactBuild()
	recorder.RecordProcedureFactBuilds(1)

	_, counters := recorder.Totals()
	want := []Counter{
		{Name: ModuleFactBuildsCounter, Value: 1},
		{Name: ProcedureFactBuildsCounter, Value: 2},
	}
	if !reflect.DeepEqual(counters, want) {
		t.Fatalf("fact build counters = %+v, want %+v", counters, want)
	}
}

func TestRecordFactBuildCountersNilSafe(t *testing.T) {
	var recorder *Recorder
	recorder.RecordModuleFactBuild()
	recorder.RecordProcedureFactBuild()
}

func TestCapabilityBuildTelemetry(t *testing.T) {
	recorder := NewRecorder()

	// Exercise the typed helpers so each capability has one canonical counter
	// name, while the generic helper remains available to capability planners
	// that represent capabilities as strings.
	recorder.RecordCapabilityResolutionBuild()
	recorder.RecordCapabilityEffectsBuild()
	recorder.RecordCapabilityArrayBuild()
	recorder.RecordCapabilityObjectBuild()
	recorder.RecordCapabilityDataflowBuild()
	recorder.RecordCapabilityDictionaryBuild()
	recorder.RecordCapabilityApplicationStateBuild()
	recorder.RecordCapabilityEventReentryBuild()
	recorder.RecordCapabilityPublicAPITypeIndexBuild()
	recorder.RecordCapabilityExcelLoopSymbolsBuild()

	finish := MeasureCapabilityBuild(WithRecorder(context.Background(), recorder), CapabilityEffectsBuildsCounter)
	finish(nil)

	stages, counters := recorder.Totals()
	counterByName := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		counterByName[counter.Name] = counter.Value
	}
	for _, name := range CapabilityBuildCounters {
		want := uint64(1)
		if name == CapabilityEffectsBuildsCounter {
			want = 2
		}
		if got := counterByName[name]; got != want {
			t.Errorf("capability counter %q = %d, want %d", name, got, want)
		}
	}
	if len(stages) != 1 || stages[0].Name != CapabilityEffectsBuildsCounter || stages[0].Outcome != "ok" || stages[0].Elapsed < 0 {
		t.Fatalf("capability stages = %+v", stages)
	}
}

func TestCapabilityBuildTelemetryNilSafe(t *testing.T) {
	var recorder *Recorder
	recorder.RecordCapabilityBuild(CapabilityResolutionBuildsCounter)
	recorder.RecordCapabilityBuilds(CapabilityEffectsBuildsCounter, 1)
	recorder.RecordCapabilityBuildWithElapsed(CapabilityArrayBuildsCounter, time.Second, "ok")
	MeasureCapabilityBuild(context.Background(), CapabilityObjectBuildsCounter)(nil)
}

func TestMeasureRecordsCanceledOutcome(t *testing.T) {
	recorder := NewRecorder()
	ctx, cancel := context.WithCancel(WithRecorder(context.Background(), recorder))
	finish := Measure(ctx, "parse")
	cancel()
	finish(0, errors.New("stopped"))

	stages, _ := recorder.Totals()
	if len(stages) != 1 || stages[0].Outcome != "canceled" {
		t.Fatalf("stages = %+v", stages)
	}
}

func TestDomainAggregateMergesFixedStagesAndCounters(t *testing.T) {
	recorder := NewRecorder()
	aggregate := NewAggregate(recorder)
	aggregate.Record(DomainDataflow, 3*time.Millisecond, 2, "ok")
	aggregate.Record(DomainSourceScan, time.Millisecond, 0, "ok")
	aggregate.Record(DomainDataflow, 5*time.Millisecond, 1, "error")
	aggregate.AddCounter(CounterDataflowCandidateProcedures, 3)
	aggregate.AddCounter(CounterDataflowCFGWalks, 7)
	aggregate.AddCounter(CounterSourceLineScans, 11)
	aggregate.Merge()

	stages, counters := recorder.Totals()
	wantStages := []Stage{
		{Name: ProcedureLocalSourceScan, Elapsed: time.Millisecond, Outcome: "ok", Calls: 1},
		{Name: ProcedureLocalDataflow, Elapsed: 8 * time.Millisecond, Outcome: "error", ResultCount: 3, Calls: 2},
	}
	if !reflect.DeepEqual(stages, wantStages) {
		t.Fatalf("stages = %+v, want %+v", stages, wantStages)
	}
	wantCounters := []Counter{
		{Name: DataflowCandidateProceduresCounter, Value: 3},
		{Name: DataflowCFGWalksCounter, Value: 7},
		{Name: SourceLineScansCounter, Value: 11},
	}
	if !reflect.DeepEqual(counters, wantCounters) {
		t.Fatalf("counters = %+v, want %+v", counters, wantCounters)
	}
}

func TestDomainAggregateMergesInParallel(t *testing.T) {
	const workers = 16
	recorder := NewRecorder()
	var wait sync.WaitGroup
	wait.Add(workers)
	for index := 0; index < workers; index++ {
		go func() {
			defer wait.Done()
			aggregate := NewAggregate(recorder)
			aggregate.Record(DomainArray, time.Nanosecond, 1, "ok")
			aggregate.Record(DomainDataflow, 2*time.Nanosecond, 2, "ok")
			aggregate.AddCounter(CounterArrayCandidateProcedures, 1)
			aggregate.AddCounter(CounterDataflowCFGWalks, 2)
			aggregate.Merge()
		}()
	}
	wait.Wait()

	stages, counters := recorder.Totals()
	if len(stages) != 2 || stages[0].Name != ProcedureLocalArray || stages[1].Name != ProcedureLocalDataflow {
		t.Fatalf("stage order = %+v", stages)
	}
	if stages[0].Calls != workers || stages[0].ResultCount != workers || stages[1].Calls != workers || stages[1].ResultCount != 2*workers {
		t.Fatalf("stage totals = %+v", stages)
	}
	wantCounters := []Counter{
		{Name: ArrayCandidateProceduresCounter, Value: workers},
		{Name: DataflowCFGWalksCounter, Value: 2 * workers},
	}
	if !reflect.DeepEqual(counters, wantCounters) {
		t.Fatalf("counters = %+v, want %+v", counters, wantCounters)
	}
}

func TestDomainAggregateMeasurement(t *testing.T) {
	recorder := NewRecorder()
	aggregate := NewAggregate(recorder)
	measurement := aggregate.Start(DomainRuntime)
	measurement.FinishOutcome(4, "canceled")
	aggregate.Merge()

	stages, _ := recorder.Totals()
	if len(stages) != 1 || stages[0].Name != ProcedureLocalRuntime || stages[0].Calls != 1 || stages[0].ResultCount != 4 || stages[0].Outcome != "canceled" {
		t.Fatalf("measurement stage = %+v", stages)
	}
}

func TestDomainAggregateNilFastPathHasNoAllocations(t *testing.T) {
	var aggregate *DomainAggregate
	allocs := testing.AllocsPerRun(1000, func() {
		measurement := aggregate.Start(DomainArray)
		measurement.Finish(0, nil)
		aggregate.Record(DomainArray, time.Second, 0, "ok")
		aggregate.AddCounter(CounterArrayCFGWalks, 1)
		aggregate.Merge()
	})
	if allocs != 0 {
		t.Fatalf("nil aggregate allocations = %v, want zero", allocs)
	}
}

func TestDomainAndCounterNamesAreStable(t *testing.T) {
	if DomainSourceScan.String() != ProcedureLocalSourceScan || DomainOther.String() != ProcedureLocalOther {
		t.Fatalf("domain names = %q, %q", DomainSourceScan, DomainOther)
	}
	if CounterRuntimeCandidateProcedures.String() != RuntimeCandidateProceduresCounter || CounterSemanticKernelRuns.String() != SemanticKernelRunsCounter {
		t.Fatalf("counter names = %q, %q", CounterRuntimeCandidateProcedures, CounterSemanticKernelRuns)
	}
	arrayCounterNames := []struct {
		counter WorkCounter
		want    string
	}{
		{CounterArrayKernelRuns, ArrayKernelRunsCounter},
		{CounterArrayProjectionRuns, ArrayProjectionRunsCounter},
	}
	for _, test := range arrayCounterNames {
		if got := test.counter.String(); got != test.want {
			t.Errorf("array counter %d name = %q, want %q", test.counter, got, test.want)
		}
	}
	plannerCounterNames := []struct {
		counter WorkCounter
		want    string
	}{
		{CounterRuntimePlannedRuns, "planned_runtime_runs"},
		{CounterRuntimeSkippedRuns, "skipped_runtime_runs"},
		{CounterArrayPlannedRuns, "planned_array_runs"},
		{CounterArraySkippedRuns, "skipped_array_runs"},
		{CounterObjectPlannedRuns, "planned_object_runs"},
		{CounterObjectSkippedRuns, "skipped_object_runs"},
		{CounterDictionaryPlannedRuns, "planned_dictionary_runs"},
		{CounterDictionarySkippedRuns, "skipped_dictionary_runs"},
		{CounterErrorPlannedRuns, "planned_error_runs"},
		{CounterErrorSkippedRuns, "skipped_error_runs"},
		{CounterDataflowPlannedRuns, "planned_dataflow_runs"},
		{CounterDataflowSkippedRuns, "skipped_dataflow_runs"},
		{CounterResourcePlannedRuns, "planned_resource_runs"},
		{CounterResourceSkippedRuns, "skipped_resource_runs"},
		{CounterExcelPlannedRuns, "planned_excel_runs"},
		{CounterExcelSkippedRuns, "skipped_excel_runs"},
		{CounterApplicationStatePlannedRuns, "planned_application_state_runs"},
		{CounterApplicationStateSkippedRuns, "skipped_application_state_runs"},
	}
	for _, test := range plannerCounterNames {
		if got := test.counter.String(); got != test.want {
			t.Errorf("counter %d name = %q, want %q", test.counter, got, test.want)
		}
	}
	if Domain(255).String() != ProcedureLocalOther || WorkCounter(255).String() != "" {
		t.Fatalf("invalid enum names = %q, %q", Domain(255), WorkCounter(255))
	}
}
