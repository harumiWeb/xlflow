package analyze

import (
	"context"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

// procedureSemanticResultStore owns semantic values for one procedure's
// projections. Its optional revision handle delegates cross-revision reuse to
// the process-local semantic query store; without a handle it retains the
// historical revision-local behavior used by focused callers.
//
// ArrayAnalysisResult and the two dataflow lane results are procedure-local
// values with bounded lifetimes. Keeping the store explicit makes the
// at-most-once materialization contract visible without introducing a
// cross-procedure finding builder.
type procedureSemanticResultStore struct {
	queryRevision        *semanticquery.Revision
	array                *ArrayAnalysisResult
	arrayBuilt           bool
	arrayHit             bool
	arrayReads           uint8
	genericDataflow      []Finding
	genericDataflowBuilt bool
	genericDataflowHit   bool
	httpDataflow         []Finding
	httpDataflowBuilt    bool
	httpDataflowHit      bool
}

func (s *procedureSemanticResultStore) materializeArray(cancelCtx context.Context, a Analyzer, file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, plan procedureAnalysisPlan) (*ArrayAnalysisResult, error) {
	if s == nil {
		return a.buildArrayAnalysisResultContext(cancelCtx, file, proc, ctx, moduleDecls, plan)
	}
	if !s.arrayBuilt {
		var array *ArrayAnalysisResult
		var err error
		if s.queryRevision != nil {
			key := semanticProcedureQueryKey(a, file, proc, "array", plan, semanticAnalysisCapability(ctx, file, proc, "array"))
			array, s.arrayHit, err = semanticquery.Evaluate[*ArrayAnalysisResult](cancelCtx, s.queryRevision, key, semanticQueryDependencies(key, file, proc), func(buildCtx context.Context) (*ArrayAnalysisResult, error) {
				return a.buildArrayAnalysisResultContext(buildCtx, file, proc, ctx, moduleDecls, plan)
			})
		} else {
			array, err = a.buildArrayAnalysisResultContext(cancelCtx, file, proc, ctx, moduleDecls, plan)
		}
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

func (s *procedureSemanticResultStore) materializeGenericDataflow(ctx context.Context, a Analyzer, file parsedFile, proc sourceProcedure, plan procedureAnalysisPlan) ([]Finding, error) {
	if s == nil {
		return a.dataFlowFindingsContext(ctx, file, proc)
	}
	if !s.genericDataflowBuilt {
		var findings []Finding
		var err error
		if s.queryRevision != nil {
			key := semanticProcedureQueryKey(a, file, proc, "dataflow", plan, semanticAnalysisCapability(analysisContext{}, file, proc, "dataflow"))
			findings, s.genericDataflowHit, err = semanticquery.Evaluate[[]Finding](ctx, s.queryRevision, key, semanticQueryDependencies(key, file, proc), func(buildCtx context.Context) ([]Finding, error) {
				return a.dataFlowFindingsContext(buildCtx, file, proc)
			})
		} else {
			findings, err = a.dataFlowFindingsContext(ctx, file, proc)
		}
		if err != nil {
			return nil, err
		}
		findings = cloneFindings(findings)
		s.genericDataflow = findings
		s.genericDataflowBuilt = true
	}
	return s.genericDataflow, nil
}

func (s *procedureSemanticResultStore) materializeHTTPDataflow(ctx context.Context, a Analyzer, file parsedFile, proc sourceProcedure, plan procedureAnalysisPlan) ([]Finding, error) {
	if s == nil {
		return a.httpTransportFindingsContext(ctx, file, proc)
	}
	if !s.httpDataflowBuilt {
		var findings []Finding
		var err error
		if s.queryRevision != nil {
			key := semanticProcedureQueryKey(a, file, proc, "http", plan, semanticAnalysisCapability(analysisContext{}, file, proc, "http"))
			findings, s.httpDataflowHit, err = semanticquery.Evaluate[[]Finding](ctx, s.queryRevision, key, semanticQueryDependencies(key, file, proc), func(buildCtx context.Context) ([]Finding, error) {
				return a.httpTransportFindingsContext(buildCtx, file, proc)
			})
		} else {
			findings, err = a.httpTransportFindingsContext(ctx, file, proc)
		}
		if err != nil {
			return nil, err
		}
		findings = cloneFindings(findings)
		s.httpDataflow = findings
		s.httpDataflowBuilt = true
	}
	return s.httpDataflow, nil
}

// Keep the store's telemetry dependency explicit at compile time. This also
// prevents an accidental replacement with an unscoped cross-procedure cache.
var _ analysisstats.WorkCounter = analysisstats.CounterSemanticResultsReused

func cloneFindings(findings []Finding) []Finding {
	if findings == nil {
		return nil
	}
	out := make([]Finding, len(findings))
	for index, finding := range findings {
		out[index] = cloneFinding(finding)
	}
	return out
}

func cloneFinding(in Finding) Finding {
	out := in
	out.NearbyCode = append([]string(nil), in.NearbyCode...)
	if in.CallCycle != nil {
		cycle := *in.CallCycle
		cycle.Path = append([]CallCycleNode(nil), in.CallCycle.Path...)
		cycle.Edges = append([]CallCycleEdge(nil), in.CallCycle.Edges...)
		cycle.EventHandlers = append([]string(nil), in.CallCycle.EventHandlers...)
		cycle.DangerousEffects = append([]CallCycleEffect(nil), in.CallCycle.DangerousEffects...)
		cycle.Uncertainty = append([]CallCycleUncertainty(nil), in.CallCycle.Uncertainty...)
		out.CallCycle = &cycle
	}
	if in.DataFlow != nil {
		dataFlow := *in.DataFlow
		dataFlow.Path = append([]DataFlowStep(nil), in.DataFlow.Path...)
		out.DataFlow = &dataFlow
	}
	if in.CommandExecution != nil {
		command := *in.CommandExecution
		out.CommandExecution = &command
	}
	if in.SQLExecution != nil {
		sql := *in.SQLExecution
		out.SQLExecution = &sql
	}
	if in.FileOperation != nil {
		fileOperation := *in.FileOperation
		if in.FileOperation.Overwrite != nil {
			overwrite := *in.FileOperation.Overwrite
			fileOperation.Overwrite = &overwrite
		}
		out.FileOperation = &fileOperation
	}
	if in.HTTPSecurity != nil {
		httpSecurity := *in.HTTPSecurity
		out.HTTPSecurity = &httpSecurity
	}
	if in.HTTPReliability != nil {
		httpReliability := *in.HTTPReliability
		out.HTTPReliability = &httpReliability
	}
	if in.OpaqueBoolean != nil {
		opaque := *in.OpaqueBoolean
		opaque.ParameterNames = append([]string(nil), in.OpaqueBoolean.ParameterNames...)
		if in.OpaqueBoolean.OptionalBooleanParameterCount != nil {
			count := *in.OpaqueBoolean.OptionalBooleanParameterCount
			opaque.OptionalBooleanParameterCount = &count
		}
		out.OpaqueBoolean = &opaque
	}
	if in.RuntimeError != nil {
		runtimeError := *in.RuntimeError
		out.RuntimeError = &runtimeError
	}
	if in.httpOwnedSinks != nil {
		out.httpOwnedSinks = make(map[int]bool, len(in.httpOwnedSinks))
		for key, value := range in.httpOwnedSinks {
			out.httpOwnedSinks[key] = value
		}
	}
	return out
}
