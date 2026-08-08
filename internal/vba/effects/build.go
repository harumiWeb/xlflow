package effects

import (
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type procedureInput struct {
	id    ProcedureIdentity
	proc  procedureir.ProcedureIR
	graph cfg.Graph
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

type summaryMembership struct {
	identityKey string
	evidence    membershipIndex[Evidence]
	uncertainty membershipIndex[CallUncertainty]
}

func newSummaryMembership(summary *ProcedureSummary) *summaryMembership {
	index := &summaryMembership{
		identityKey: summary.Identity.Key(),
		evidence:    newMembershipIndex(len(summary.Direct)+len(summary.Propagated), evidenceKey),
		uncertainty: newMembershipIndex(len(summary.DirectUncertainty)+len(summary.PropagatedUncertainty), uncertaintyKey),
	}
	for _, fact := range summary.Direct {
		index.evidence.add(fact)
	}
	for _, fact := range summary.Propagated {
		index.evidence.add(fact)
	}
	for _, fact := range summary.DirectUncertainty {
		index.uncertainty.add(fact)
	}
	for _, fact := range summary.PropagatedUncertainty {
		index.uncertainty.add(fact)
	}
	return index
}

// Build computes direct facts and then propagates finite provenance sets over
// uniquely resolved, reachable project-local calls until a fixed point.
func Build(documents []Document) ProjectSummary {
	inputs := collectInputs(documents)
	summaries := make(map[string]*ProcedureSummary, len(inputs))
	candidateKeys := candidateIndex(inputs)
	var edges []edge
	for _, input := range inputs {
		summary := &ProcedureSummary{Identity: input.id}
		reachable := reachableStatements(input.proc, input.graph)
		statements := statementIndex(input.proc)
		extractStatements(summary, input.proc, reachable)
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
				procedureir.ResolutionExternal, procedureir.ResolutionMemberCall:
				summary.DirectUncertainty = append(summary.DirectUncertainty, uncertainty(input.id, call))
			}
		}
		dedupeDirect(summary)
		summaries[input.id.Key()] = summary
	}
	sort.Slice(edges, func(i, j int) bool {
		if edges[i].from != edges[j].from {
			return edges[i].from < edges[j].from
		}
		return edges[i].to < edges[j].to
	})
	propagate(summaries, edges)
	out := ProjectSummary{byKey: map[string]int{}}
	for _, summary := range summaries {
		out.procedures = append(out.procedures, *summary)
	}
	sortSummaries(out.procedures)
	for i := range out.procedures {
		out.byKey[out.procedures[i].Identity.Key()] = i
	}
	return out
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
			out = append(out, procedureInput{id: identity(doc.IR, proc), proc: proc, graph: graph})
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
		for _, block := range graph.Blocks {
			if block.ID == id && block.Kind == cfg.BlockStatement {
				out[block.StatementID] = true
				break
			}
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
	}
	return CallUncertainty{Kind: kind, Origin: origin, Range: call.Range, StatementID: call.StatementID, CallID: call.ID, Callee: call.Callee.Text}
}

func dedupeDirect(summary *ProcedureSummary) {
	summary.Direct = uniqueEvidence(summary.Direct)
	summary.DirectUncertainty = uniqueUncertainty(summary.DirectUncertainty)
}

func propagate(summaries map[string]*ProcedureSummary, edges []edge) {
	callers := map[string][]string{}
	for _, edge := range edges {
		callers[edge.to] = append(callers[edge.to], edge.from)
	}
	queue := make([]string, 0, len(summaries))
	queued := map[string]bool{}
	membership := make(map[string]*summaryMembership, len(summaries))
	for key := range summaries {
		queue = append(queue, key)
		queued[key] = true
		if summaries[key] != nil {
			membership[key] = newSummaryMembership(summaries[key])
		}
	}
	sort.Strings(queue)
	for len(queue) > 0 {
		calleeKey := queue[0]
		queue = queue[1:]
		queued[calleeKey] = false
		for _, callerKey := range callers[calleeKey] {
			caller, callee := summaries[callerKey], summaries[calleeKey]
			callerMembership := membership[callerKey]
			if caller == nil || callee == nil || callerMembership == nil {
				continue
			}
			changed := false
			propagateEvidence := func(facts []Evidence) {
				for _, fact := range facts {
					if fact.Origin.Key() == callerMembership.identityKey || !callerMembership.evidence.add(fact) {
						continue
					}
					caller.Propagated = append(caller.Propagated, fact)
					changed = true
				}
			}
			propagateEvidence(callee.Direct)
			propagateEvidence(callee.Propagated)

			propagateUncertainty := func(facts []CallUncertainty) {
				for _, fact := range facts {
					if fact.Origin.Key() == callerMembership.identityKey || !callerMembership.uncertainty.add(fact) {
						continue
					}
					caller.PropagatedUncertainty = append(caller.PropagatedUncertainty, fact)
					changed = true
				}
			}
			propagateUncertainty(callee.DirectUncertainty)
			propagateUncertainty(callee.PropagatedUncertainty)
			if changed && !queued[callerKey] {
				queue = append(queue, callerKey)
				queued[callerKey] = true
			}
		}
	}
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
