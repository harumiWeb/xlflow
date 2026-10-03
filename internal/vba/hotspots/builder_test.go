package hotspots

import (
	"context"
	"errors"
	"fmt"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
)

func TestBuildFromMetricsPreservesGraphSignalsAndOrdering(t *testing.T) {
	metrics := []proceduremetrics.ProcedureMetrics{
		hotspotMetric("src/A.bas", "A", "Root", "b.work"),
		hotspotMetric("src/A.bas", "A", "Isolated"),
		hotspotMetric("src/B.bas", "B", "Work", "c.work"),
		hotspotMetric("src/C.bas", "C", "Work", "b.work"),
	}

	report, stats, err := BuildFromMetrics(t.Context(), metrics)
	if err != nil {
		t.Fatalf("BuildFromMetrics() error = %v", err)
	}
	procedures := entitiesByModuleAndName(report.Procedures)
	modules := entitiesByName(report.Modules)
	assertRawSignal(t, procedures["A.Root"], SignalAffectedModules, 3)
	assertRawSignal(t, procedures["A.Isolated"], SignalAffectedModules, 1)
	assertRawSignal(t, procedures["B.Work"], SignalAffectedModules, 2)
	assertRawSignal(t, procedures["A.Root"], SignalCycleCount, 0)
	assertRawSignal(t, procedures["B.Work"], SignalCycleCount, 1)
	assertRawSignal(t, procedures["B.Work"], SignalCallFanIn, 2)
	assertRawSignal(t, modules["A"], SignalAffectedModules, 3)
	assertRawSignal(t, modules["B"], SignalCallFanIn, 2)
	assertRawSignal(t, modules["B"], SignalCycleCount, 1)

	if stats.ProcedureEdgeCount != 3 || stats.ModuleEdgeCount != 3 {
		t.Fatalf("edge stats = %+v", stats)
	}
	if stats.ProcedureSCCCount != 3 || stats.ModuleSCCCount != 2 {
		t.Fatalf("SCC stats = %+v", stats)
	}
	if stats.AffectedClosureSetUnions != 2 {
		t.Fatalf("closure union stats = %+v", stats)
	}

	reversed := []proceduremetrics.ProcedureMetrics{metrics[3], metrics[2], metrics[1], metrics[0]}
	reordered, _, err := BuildFromMetrics(t.Context(), reversed)
	if err != nil {
		t.Fatalf("BuildFromMetrics(reordered) error = %v", err)
	}
	if !reflect.DeepEqual(report, reordered) {
		t.Fatalf("report changes with input order\nfirst:  %#v\nsecond: %#v", report, reordered)
	}
}

func TestBuildFromMetricsSharesReachabilityAcrossSCCs(t *testing.T) {
	const (
		moduleCount         = 48
		proceduresPerModule = 4
	)
	metrics := make([]proceduremetrics.ProcedureMetrics, 0, moduleCount*proceduresPerModule)
	for moduleIndex := range moduleCount {
		module := fmt.Sprintf("M%02d", moduleIndex)
		file := fmt.Sprintf("src/%s.bas", module)
		for procedureIndex := range proceduresPerModule {
			name := fmt.Sprintf("P%02d", procedureIndex)
			nextIndex := (procedureIndex + 1) % proceduresPerModule
			calls := []string{strings.ToLower(module + "." + fmt.Sprintf("P%02d", nextIndex))}
			if procedureIndex == 0 && moduleIndex+1 < moduleCount {
				nextModule := fmt.Sprintf("M%02d", moduleIndex+1)
				calls = append(calls, strings.ToLower(nextModule+".P00"))
			}
			metrics = append(metrics, hotspotMetric(file, module, name, calls...))
		}
	}

	report, stats, err := BuildFromMetrics(t.Context(), metrics)
	if err != nil {
		t.Fatalf("BuildFromMetrics() error = %v", err)
	}
	if stats.ProcedureCount != moduleCount*proceduresPerModule || stats.ModuleCount != moduleCount {
		t.Fatalf("entity stats = %+v", stats)
	}
	if stats.ProcedureSCCCount != moduleCount || stats.ModuleSCCCount != moduleCount {
		t.Fatalf("expected each module's procedures to share one procedure SCC: %+v", stats)
	}
	if want := 2 * (moduleCount - 1); stats.AffectedClosureSetUnions != want {
		t.Fatalf("affected closure unions = %d, want %d; stats=%+v", stats.AffectedClosureSetUnions, want, stats)
	}
	if got := entitiesByModuleAndName(report.Procedures)["M00.P00"].RawSignals[string(SignalAffectedModules)]; got != moduleCount {
		t.Fatalf("first SCC affected module count = %d, want %d", got, moduleCount)
	}
	if got := entitiesByName(report.Modules)["M47"].RawSignals[string(SignalAffectedModules)]; got != 1 {
		t.Fatalf("last module affected module count = %d, want 1", got)
	}
	if stats.AffectedClosureDenseWordsStored == 0 || stats.AffectedClosureDenseWordScans == 0 {
		t.Fatalf("dense reachability path was not exercised: %+v", stats)
	}
	if stats.AffectedClosureComponentCardinalityReads != stats.ProcedureSCCCount+stats.ModuleSCCCount {
		t.Fatalf("component cardinalities were not cached per SCC: %+v", stats)
	}
}

func TestBuildFromMetricsKeepsSparseDisconnectedClosuresLinear(t *testing.T) {
	const entityCount = 2_048
	metrics := make([]proceduremetrics.ProcedureMetrics, entityCount)
	for index := range metrics {
		module := fmt.Sprintf("M%04d", index)
		metrics[index] = hotspotMetric("src/"+module+".bas", module, "Run")
	}

	report, stats, err := BuildFromMetrics(t.Context(), metrics)
	if err != nil {
		t.Fatalf("BuildFromMetrics() error = %v", err)
	}
	if stats.AffectedClosureSetCount != 2*entityCount {
		t.Fatalf("closure set count = %d, want %d; stats=%+v", stats.AffectedClosureSetCount, 2*entityCount, stats)
	}
	if stats.AffectedClosureSetUnions != 0 || stats.AffectedClosureSparseMergeSteps != 0 || stats.AffectedClosureDenseWordScans != 0 {
		t.Fatalf("disconnected graph performed closure merge work: %+v", stats)
	}
	if stats.AffectedClosureSparseEntriesStored != 2*entityCount || stats.AffectedClosureDenseWordsStored != 0 {
		t.Fatalf("disconnected closure storage is not sparse: %+v", stats)
	}
	if stats.AffectedClosureComponentCardinalityReads != 2*entityCount {
		t.Fatalf("component cardinality reads = %d, want %d", stats.AffectedClosureComponentCardinalityReads, 2*entityCount)
	}
	for _, entity := range report.Procedures {
		if got := entity.RawSignals[string(SignalAffectedModules)]; got != 1 {
			t.Fatalf("isolated procedure %s affected modules = %d, want 1", entity.ID, got)
		}
	}
}

func TestBuildFromMetricsRetainsBoundedCycleEnumeration(t *testing.T) {
	const procedureCount = 10
	metrics := make([]proceduremetrics.ProcedureMetrics, 0, procedureCount)
	for caller := range procedureCount {
		callees := make([]string, 0, procedureCount)
		for callee := range procedureCount {
			callees = append(callees, strings.ToLower(fmt.Sprintf("Cycles.P%02d", callee)))
		}
		metrics = append(metrics, hotspotMetric(
			"src/Cycles.bas", "Cycles", fmt.Sprintf("P%02d", caller), callees...,
		))
	}

	_, stats, err := BuildFromMetrics(t.Context(), metrics)
	if err != nil {
		t.Fatalf("BuildFromMetrics() error = %v", err)
	}
	if !stats.CycleEnumerationTruncated {
		t.Fatalf("dense cycle graph did not exhaust the fixed work budget: %+v", stats)
	}
	if stats.CycleEnumerationWork <= cycleEnumerationWorkBudget || stats.CycleEnumerationWork > cycleEnumerationWorkBudget+procedureCount {
		t.Fatalf("cycle work = %d, want budget exhaustion with at most one stack unwind per procedure", stats.CycleEnumerationWork)
	}
}

func TestBuildFromMetricsHonorsCancellationDuringGraphWork(t *testing.T) {
	metrics := make([]proceduremetrics.ProcedureMetrics, 256)
	for index := range metrics {
		module := fmt.Sprintf("M%03d", index)
		metrics[index] = hotspotMetric("src/"+module+".bas", module, "Run")
	}
	base, cancel := context.WithCancel(t.Context())
	defer cancel()
	ctx := &cancelAfterChecksContext{Context: base, cancel: cancel, remaining: 600}

	report, _, err := BuildFromMetrics(ctx, metrics)
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("BuildFromMetrics() error = %v, want context.Canceled", err)
	}
	if report.SchemaVersion != 0 {
		t.Fatalf("canceled build returned a partial report: %+v", report)
	}
}

type cancelAfterChecksContext struct {
	context.Context
	cancel    context.CancelFunc
	remaining int
}

func (c *cancelAfterChecksContext) Err() error {
	c.remaining--
	if c.remaining <= 0 {
		c.cancel()
	}
	return c.Context.Err()
}

func hotspotMetric(file, module, name string, callees ...string) proceduremetrics.ProcedureMetrics {
	return proceduremetrics.ProcedureMetrics{
		File: file, Module: module, ModuleKind: "standard", Name: name,
		Kind: procedureir.ProcedureSub, ResolvedCallees: callees,
	}
}

func entitiesByName(entities []Entity) map[string]Entity {
	result := make(map[string]Entity, len(entities))
	for _, entity := range entities {
		result[entity.Name] = entity
	}
	return result
}

func entitiesByModuleAndName(entities []Entity) map[string]Entity {
	result := make(map[string]Entity, len(entities))
	for _, entity := range entities {
		result[entity.Module+"."+entity.Name] = entity
	}
	return result
}

func assertRawSignal(t *testing.T, entity Entity, signal SignalName, want int) {
	t.Helper()
	if got := entity.RawSignals[string(signal)]; got != want {
		t.Errorf("%s %s = %d, want %d", entity.Name, signal, got, want)
	}
}
