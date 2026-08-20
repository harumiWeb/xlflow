package effects

import (
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type procedureInput struct {
	id        ProcedureIdentity
	proc      procedureir.ProcedureIR
	graph     cfg.Graph
	reachable map[int]bool
}

type edge struct{ from, to string }

type membershipIndex[T any] struct {
	keys map[string]struct{}
	key  func(T) string
}

func newMembershipIndex[T any](capacity int, key func(T) string) membershipIndex[T] {
	return membershipIndex[T]{keys: make(map[string]struct{}, capacity), key: key}
}

// add reports whether value was absent. The key is computed exactly once so
// fixed-point propagation does not repeatedly rebuild keys for an expanding
// summary.
func (i *membershipIndex[T]) add(value T) bool {
	key := i.key(value)
	if _, exists := i.keys[key]; exists {
		return false
	}
	i.keys[key] = struct{}{}
	return true
}

// Build computes direct facts and then propagates bounded semantic state over
// uniquely resolved, reachable project-local calls until a fixed point.
func Build(documents []Document) ProjectSummary {
	project, _ := BuildWithStats(documents)
	return project
}

// BuildWithStats is Build with developer-facing fixed-point counters. Detailed
// source provenance remains at its origin and is reconstructed lazily from the
// project call graph by ProjectSummary.Lookup/All.
func BuildWithStats(documents []Document) (ProjectSummary, BuildStats) {
	inputs := collectInputs(documents)
	summaries := make(map[string]*ProcedureSummary, len(inputs))
	candidateKeys := candidateIndex(inputs)
	loggerTargets := loggerProcedureIndex(inputs, candidateKeys)
	rethrowTargets := rethrowProcedureIndex(inputs, candidateKeys)
	terminalTargets := terminalProcedureIndex(inputs, candidateKeys)
	var edges []edge
	for _, input := range inputs {
		summary := &ProcedureSummary{Identity: input.id}
		reachable := input.reachable
		statements := statementIndex(input.proc)
		extractStatements(summary, input.proc, reachable)
		extractErrorSummary(summary, input.proc, input.graph, reachable, candidateKeys, loggerTargets, rethrowTargets, terminalTargets)
		for _, call := range input.proc.Calls {
			statement := statements[call.StatementID]
			if !reachable[call.StatementID] || statement.Recovered {
				continue
			}
			extractCall(summary, call, statement)
			switch call.Resolution.Status {
			case procedureir.ResolutionMatched:
				if len(call.Resolution.Candidates) == 1 {
					if target, ok := candidateKeys[candidateKey(call.Resolution.Candidates[0])]; ok {
						edges = append(edges, edge{from: input.id.Key(), to: target})
					}
				}
			case procedureir.ResolutionAmbiguous, procedureir.ResolutionUnresolved,
				procedureir.ResolutionExternal, procedureir.ResolutionMemberCall,
				procedureir.ResolutionDynamic, procedureir.ResolutionIncomplete,
				procedureir.ResolutionNonCallable:
				summary.DirectUncertainty = append(summary.DirectUncertainty, uncertainty(input.id, call))
			}
		}
		dedupeDirect(summary)
		summary.semantic = semanticStateFromSummary(summary)
		summary.Error = errorSummaryWithState(summary.Error, summary.semantic)
		summaries[input.id.Key()] = summary
	}
	sort.Slice(edges, func(i, j int) bool {
		if edges[i].from != edges[j].from {
			return edges[i].from < edges[j].from
		}
		return edges[i].to < edges[j].to
	})
	stats := propagateBounded(summaries, edges)
	out := ProjectSummary{
		byKey:           map[string]int{},
		byCandidateLine: map[int][]int{},
		stats:           stats,
		provenance: &provenanceGraph{
			callers:   map[string][]string{},
			callees:   map[string][]string{},
			summaries: map[string]ProcedureSummary{},
		},
	}
	for _, item := range edges {
		out.provenance.callees[item.from] = appendUniqueKey(out.provenance.callees[item.from], item.to)
		out.provenance.callers[item.to] = appendUniqueKey(out.provenance.callers[item.to], item.from)
	}
	for _, summary := range summaries {
		out.procedures = append(out.procedures, *summary)
	}
	sortSummaries(out.procedures)
	for i := range out.procedures {
		out.byKey[out.procedures[i].Identity.Key()] = i
		out.provenance.summaries[out.procedures[i].Identity.Key()] = out.procedures[i]
		line := out.procedures[i].Identity.DeclarationLine
		out.byCandidateLine[line] = append(out.byCandidateLine[line], i)
	}
	for key := range out.provenance.callees {
		sort.Strings(out.provenance.callees[key])
	}
	for key := range out.provenance.callers {
		sort.Strings(out.provenance.callers[key])
	}
	return out, stats
}

func collectInputs(documents []Document) []procedureInput {
	var out []procedureInput
	for _, doc := range documents {
		graphs := map[string]cfg.Graph{}
		for _, graph := range doc.CFG.Graphs {
			graphs[graph.Procedure.QualifiedName+"\x00"+string(graph.Procedure.Kind)] = graph
		}
		for _, proc := range doc.IR.Procedures {
			graph := graphs[proc.Symbol.QualifiedName+"\x00"+string(proc.Symbol.Kind)]
			out = append(out, procedureInput{id: identity(doc.IR, proc), proc: proc, graph: graph, reachable: reachableStatements(proc, graph)})
		}
	}
	sort.Slice(out, func(i, j int) bool { return out[i].id.Key() < out[j].id.Key() })
	return out
}

func candidateIndex(inputs []procedureInput) map[string]string {
	out := map[string]string{}
	for _, input := range inputs {
		key := strings.Join([]string{input.id.File, strings.ToLower(input.id.QualifiedName), string(input.id.Kind), decimal(input.id.DeclarationLine)}, "\x00")
		out[key] = input.id.Key()
	}
	return out
}

func candidateKey(c procedureir.Candidate) string {
	return strings.Join([]string{canonicalPath(c.File), strings.ToLower(c.QualifiedName), strings.ToLower(c.Kind), decimal(c.Line)}, "\x00")
}

func appendUniqueKey(items []string, value string) []string {
	for _, item := range items {
		if item == value {
			return items
		}
	}
	return append(items, value)
}

func reachableStatements(proc procedureir.ProcedureIR, graph cfg.Graph) map[int]bool {
	out := map[int]bool{}
	if len(graph.Blocks) == 0 {
		// Absence of a graph is not proof of unreachability. This fallback also
		// keeps the layer conservative for independently constructed IR values.
		for _, statement := range proc.Statements {
			out[statement.ID] = true
		}
		return out
	}
	for _, id := range graph.Reachable(cfg.EdgeFilter{}) {
		block, ok := graph.BlockByID(id)
		if ok && block.Kind == cfg.BlockStatement {
			out[block.StatementID] = true
		}
	}
	return out
}

func statementIndex(proc procedureir.ProcedureIR) map[int]procedureir.Statement {
	out := map[int]procedureir.Statement{}
	for _, statement := range proc.Statements {
		out[statement.ID] = statement
	}
	return out
}

func uncertainty(origin ProcedureIdentity, call procedureir.CallSite) CallUncertainty {
	kind := UncertaintyUnresolved
	switch call.Resolution.Status {
	case procedureir.ResolutionAmbiguous:
		kind = UncertaintyAmbiguous
	case procedureir.ResolutionExternal:
		kind = UncertaintyExternal
	case procedureir.ResolutionMemberCall:
		kind = UncertaintyDynamic
	case procedureir.ResolutionDynamic:
		kind = UncertaintyDynamic
	}
	return CallUncertainty{Kind: kind, Origin: origin, Range: call.Range, StatementID: call.StatementID, CallID: call.ID, Callee: call.Callee.Text}
}

func dedupeDirect(summary *ProcedureSummary) {
	summary.Direct = uniqueEvidence(summary.Direct)
	summary.Error.Direct = uniqueErrorEvidence(summary.Error.Direct)
	summary.DirectUncertainty = uniqueUncertainty(summary.DirectUncertainty)
}

// propagateBounded propagates only finite semantic state.
func propagateBounded(summaries map[string]*ProcedureSummary, edges []edge) BuildStats {
	callers := map[string][]string{}
	for _, item := range edges {
		callers[item.to] = appendUniqueKey(callers[item.to], item.from)
	}
	for key := range callers {
		sort.Strings(callers[key])
	}
	queue := make([]string, 0, len(summaries))
	queued := map[string]bool{}
	for key := range summaries {
		queue = append(queue, key)
		queued[key] = true
	}
	sort.Strings(queue)
	var stats BuildStats
	for len(queue) > 0 {
		calleeKey := queue[0]
		queue = queue[1:]
		queued[calleeKey] = false
		stats.WorklistEvaluations++
		callee := summaries[calleeKey]
		if callee == nil || callee.semantic == nil {
			continue
		}
		for _, callerKey := range callers[calleeKey] {
			caller := summaries[callerKey]
			if caller == nil || caller.semantic == nil {
				continue
			}
			changed := mergeSemanticState(caller.semantic, callee.semantic, callerKey, calleeKey)
			caller.Error = errorSummaryWithState(caller.Error, caller.semantic)
			if changed && !queued[callerKey] {
				queue = append(queue, callerKey)
				queued[callerKey] = true
			}
		}
	}
	for _, summary := range summaries {
		if summary == nil || summary.semantic == nil {
			continue
		}
		direct := semanticStateFromSummary(summary).factCount()
		facts := summary.semantic.factCount()
		if facts > direct {
			facts -= direct
		} else {
			facts = 0
		}
		if facts > stats.MaxPropagatedFactsPerProcedure {
			stats.MaxPropagatedFactsPerProcedure = facts
		}
		stats.TotalPropagatedFacts += facts
	}
	return stats
}

func uniqueEvidence(in []Evidence) []Evidence {
	index := newMembershipIndex(len(in), evidenceKey)
	out := make([]Evidence, 0, len(in))
	for _, v := range in {
		if index.add(v) {
			out = append(out, v)
		}
	}
	return out
}
func uniqueUncertainty(in []CallUncertainty) []CallUncertainty {
	index := newMembershipIndex(len(in), uncertaintyKey)
	out := make([]CallUncertainty, 0, len(in))
	for _, v := range in {
		if index.add(v) {
			out = append(out, v)
		}
	}
	return out
}

func uniqueErrorEvidence(in []ErrorEvidence) []ErrorEvidence {
	index := newMembershipIndex(len(in), errorEvidenceKey)
	out := make([]ErrorEvidence, 0, len(in))
	for _, value := range in {
		if index.add(value) {
			out = append(out, value)
		}
	}
	return out
}
