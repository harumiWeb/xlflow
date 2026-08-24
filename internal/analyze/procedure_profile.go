package analyze

import (
	"context"

	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

type procedureDomain = analysisstats.Domain

const (
	procedureDomainSourceScan       = analysisstats.DomainSourceScan
	procedureDomainRuntime          = analysisstats.DomainRuntime
	procedureDomainArray            = analysisstats.DomainArray
	procedureDomainObject           = analysisstats.DomainObject
	procedureDomainDictionary       = analysisstats.DomainDictionary
	procedureDomainError            = analysisstats.DomainError
	procedureDomainDataflow         = analysisstats.DomainDataflow
	procedureDomainResource         = analysisstats.DomainResource
	procedureDomainExcel            = analysisstats.DomainExcel
	procedureDomainApplicationState = analysisstats.DomainApplicationState
	procedureDomainOther            = analysisstats.DomainOther
)

var gatedProcedureDomains = [...]procedureDomain{
	procedureDomainRuntime, procedureDomainArray, procedureDomainObject,
	procedureDomainDictionary, procedureDomainError, procedureDomainDataflow,
	procedureDomainResource, procedureDomainExcel, procedureDomainApplicationState,
}

// procedureDomainProfile is owned by one serial file analysis or one
// procedure batch. It never crosses worker boundaries. This keeps the hot
// procedure loop free from Recorder locks and makes the default (no recorder)
// path a nil fast path.
type procedureDomainProfile struct {
	aggregate *analysisstats.DomainAggregate
}

func newProcedureDomainProfile(ctx context.Context) *procedureDomainProfile {
	recorder := analysisstats.FromContext(ctx)
	if recorder == nil {
		return nil
	}
	return &procedureDomainProfile{aggregate: analysisstats.NewAggregate(recorder)}
}

type procedureDomainMeasurement struct {
	owner       *procedureDomainProfile
	measurement analysisstats.Measurement
}

func (p *procedureDomainProfile) begin(domain procedureDomain) procedureDomainMeasurement {
	if p == nil {
		return procedureDomainMeasurement{}
	}
	return procedureDomainMeasurement{owner: p, measurement: p.aggregate.Start(domain)}
}

func (m procedureDomainMeasurement) finish(resultCount int) {
	if m.owner == nil {
		return
	}
	m.measurement.FinishOutcome(resultCount, "ok")
}

func (m procedureDomainMeasurement) finishOutcome(ctx context.Context, resultCount int, err error) {
	if m.owner == nil {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	if ctx != nil && ctx.Err() != nil {
		outcome = "canceled"
	}
	m.measurement.FinishOutcome(resultCount, outcome)
}

func (p *procedureDomainProfile) add(counter analysisstats.WorkCounter, value uint64) {
	if p == nil {
		return
	}
	p.aggregate.AddCounter(counter, value)
}

func (p *procedureDomainProfile) candidate(seen *uint64, counter analysisstats.WorkCounter) {
	if p == nil || seen == nil {
		return
	}
	bit := uint(counter)
	if bit >= 64 || *seen&(uint64(1)<<bit) != 0 {
		return
	}
	*seen |= uint64(1) << bit
	p.aggregate.AddCounter(counter, 1)
}

// Keep the per-procedure candidate bitset wide enough for every current
// workload counter. If a future counter grows past the bitset limit, this
// package must widen the mask before candidate attribution can silently drop.
const procedureCandidateBitLimit = 64

var _ [procedureCandidateBitLimit - analysisstats.WorkCounterCount]struct{}

func (p *procedureDomainProfile) plannerDecision(plan procedureAnalysisPlan, domain procedureDomain) {
	if p == nil || !plan.enabledDomain(domain) {
		return
	}
	planned, skipped, ok := plannerCounters(domain)
	if !ok {
		return
	}
	if plan.runs(domain) {
		p.add(planned, 1)
		return
	}
	p.add(skipped, 1)
	// Keep the existing performance-log stage schema observable even when the
	// planner proves a domain irrelevant. This records an empty stage without
	// entering the domain or allocating its semantic state.
	measurement := p.begin(domain)
	measurement.finish(0)
}

func (p *procedureDomainProfile) plannerDecisions(plan procedureAnalysisPlan) {
	for _, domain := range gatedProcedureDomains {
		p.plannerDecision(plan, domain)
	}
}

// planSummary records the explicit kernel closure once per procedure. Domain
// planner counters remain available for compatibility, while these counters
// describe the plan actually handed to the executor. Iterating static enum
// order keeps this on the no-allocation hot path.
func (p *procedureDomainProfile) planSummary(plan procedureAnalysisPlan) {
	p.planSummaryForKernelMask(plan, ^uint16(0))
}

func (p *procedureDomainProfile) realtimePlanSummary(plan procedureAnalysisPlan) {
	mask := procedureKernelBit(procedureKernelRuntime) |
		procedureKernelBit(procedureKernelArray) |
		procedureKernelBit(procedureKernelDictionary) |
		procedureKernelBit(procedureKernelError) |
		procedureKernelBit(procedureKernelDataflow) |
		procedureKernelBit(procedureKernelResource) |
		procedureKernelBit(procedureKernelExcel)
	p.planSummaryForKernelMask(plan, mask)
}

func (p *procedureDomainProfile) planSummaryForKernelMask(plan procedureAnalysisPlan, mask uint16) {
	if p == nil {
		return
	}
	p.add(analysisstats.CounterAnalysisPlans, 1)
	for _, kernel := range canonicalProcedureKernelOrder {
		if mask&procedureKernelBit(kernel) == 0 || !plan.enabledKernel(kernel) {
			continue
		}
		if plan.runsKernel(kernel) {
			p.add(analysisstats.CounterPlannedKernelRuns, 1)
		} else {
			p.add(analysisstats.CounterSkippedKernelRuns, 1)
		}
	}
}

func (p *procedureDomainProfile) semanticResultReused() {
	if p == nil {
		return
	}
	p.add(analysisstats.CounterSemanticResultsReused, 1)
}

func plannerCounters(domain procedureDomain) (analysisstats.WorkCounter, analysisstats.WorkCounter, bool) {
	switch domain {
	case procedureDomainRuntime:
		return analysisstats.CounterRuntimePlannedRuns, analysisstats.CounterRuntimeSkippedRuns, true
	case procedureDomainArray:
		return analysisstats.CounterArrayPlannedRuns, analysisstats.CounterArraySkippedRuns, true
	case procedureDomainObject:
		return analysisstats.CounterObjectPlannedRuns, analysisstats.CounterObjectSkippedRuns, true
	case procedureDomainDictionary:
		return analysisstats.CounterDictionaryPlannedRuns, analysisstats.CounterDictionarySkippedRuns, true
	case procedureDomainError:
		return analysisstats.CounterErrorPlannedRuns, analysisstats.CounterErrorSkippedRuns, true
	case procedureDomainDataflow:
		return analysisstats.CounterDataflowPlannedRuns, analysisstats.CounterDataflowSkippedRuns, true
	case procedureDomainResource:
		return analysisstats.CounterResourcePlannedRuns, analysisstats.CounterResourceSkippedRuns, true
	case procedureDomainExcel:
		return analysisstats.CounterExcelPlannedRuns, analysisstats.CounterExcelSkippedRuns, true
	case procedureDomainApplicationState:
		return analysisstats.CounterApplicationStatePlannedRuns, analysisstats.CounterApplicationStateSkippedRuns, true
	default:
		return 0, 0, false
	}
}

func (p *procedureDomainProfile) kernel() {
	p.add(analysisstats.CounterSemanticKernelRuns, 1)
}

func (p *procedureDomainProfile) flush() {
	if p == nil || p.aggregate == nil {
		return
	}
	p.aggregate.Merge()
}
