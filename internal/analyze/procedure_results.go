package analyze

import (
	"context"

	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

// procedureSemanticResultStore owns semantic values for exactly one
// procedure analysis revision. It deliberately has no package-level cache:
// project summaries, configuration, and source revisions are all inputs to a
// result and must not leak between workers or requests.
//
// ArrayAnalysisResult is currently the shared result with more than one
// projection. Keeping the store explicit makes the lifetime and reuse
// contract visible and gives future kernels the same bounded extension point.
type procedureSemanticResultStore struct {
	array      *ArrayAnalysisResult
	arrayBuilt bool
	arrayReads uint8
}

func (s *procedureSemanticResultStore) materializeArray(cancelCtx context.Context, a Analyzer, file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, plan procedureAnalysisPlan) (*ArrayAnalysisResult, error) {
	if s == nil {
		return a.buildArrayAnalysisResultContext(cancelCtx, file, proc, ctx, moduleDecls, plan)
	}
	if !s.arrayBuilt {
		array, err := a.buildArrayAnalysisResultContext(cancelCtx, file, proc, ctx, moduleDecls, plan)
		if err != nil {
			return nil, err
		}
		s.array = array
		s.arrayBuilt = true
	}
	return s.array, nil
}

// arrayProjection returns the immutable result for a projection and records
// only reuse after the first consumer. Returning nil is safe for non-applicable
// array procedures and does not increment telemetry.
func (s *procedureSemanticResultStore) arrayProjection(profile *procedureDomainProfile) *ArrayAnalysisResult {
	if s == nil || s.array == nil {
		return nil
	}
	if s.arrayReads > 0 {
		if profile != nil {
			profile.semanticResultReused()
		}
	}
	s.arrayReads++
	return s.array
}

// Keep the store's telemetry dependency explicit at compile time. This also
// prevents an accidental replacement with an unscoped cross-procedure cache.
var _ analysisstats.WorkCounter = analysisstats.CounterSemanticResultsReused
