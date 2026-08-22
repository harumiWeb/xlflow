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
	m.finishOutcome(resultCount, nil, nil)
}

func (m procedureDomainMeasurement) finishOutcome(resultCount int, ctx context.Context, err error) {
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

func (p *procedureDomainProfile) candidate(seen *uint32, counter analysisstats.WorkCounter) {
	if p == nil || seen == nil {
		return
	}
	bit := uint(counter)
	if bit >= 32 || *seen&(uint32(1)<<bit) != 0 {
		return
	}
	*seen |= uint32(1) << bit
	p.aggregate.AddCounter(counter, 1)
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
