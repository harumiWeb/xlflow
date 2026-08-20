package effects

import (
	"sort"
	"strings"
)

func semanticStateFromSummary(summary *ProcedureSummary) *semanticState {
	state := newSemanticState()
	for _, evidence := range summary.Direct {
		state.effects[evidence.Effect] = struct{}{}
		switch evidence.Effect {
		case ChangesApplicationState:
			if target := semanticApplicationTarget(evidence.Target); target != "" {
				state.applicationChanges[target] = struct{}{}
			}
		case RestoresApplicationState:
			if target := semanticApplicationTarget(evidence.Target); target != "" {
				state.applicationRestores[target] = struct{}{}
			}
		}
	}
	for _, evidence := range summary.Error.Direct {
		state.errors[evidence.Behavior] = struct{}{}
	}
	for _, uncertainty := range summary.DirectUncertainty {
		state.uncertainty[uncertainty.Kind] = struct{}{}
	}
	return state
}

func semanticApplicationTarget(target string) string {
	target = strings.ToLower(strings.TrimSpace(target))
	if !strings.HasPrefix(target, "application.") {
		return ""
	}
	switch strings.TrimPrefix(target, "application.") {
	case "enableevents", "displayalerts", "screenupdating", "calculation", "statusbar", "cursor", "interactive", "asktoupdatelinks", "automationsecurity", "cutcopymode":
		return target
	default:
		// Keep the compact state bounded even if a future extractor accepts an
		// unrecognised Application member. Its direct evidence remains available
		// to the lazy compatibility projection.
		return ""
	}
}

func mergeSemanticState(dst, src *semanticState, owner, callee string) bool {
	if dst == nil || src == nil {
		return false
	}
	changed := false
	_, hadMayRaise := dst.errors[ErrorMayRaise]
	for kind := range src.effects {
		if _, ok := dst.effects[kind]; !ok {
			dst.effects[kind] = struct{}{}
			changed = true
		}
	}
	for target := range src.applicationChanges {
		if _, ok := dst.applicationChanges[target]; !ok {
			dst.applicationChanges[target] = struct{}{}
			changed = true
		}
	}
	for target := range src.applicationRestores {
		if _, ok := dst.applicationRestores[target]; !ok {
			dst.applicationRestores[target] = struct{}{}
			changed = true
		}
	}
	for behavior := range src.errors {
		if _, ok := dst.errors[behavior]; !ok {
			dst.errors[behavior] = struct{}{}
			changed = true
		}
	}
	for kind := range src.uncertainty {
		if _, ok := dst.uncertainty[kind]; !ok {
			dst.uncertainty[kind] = struct{}{}
			changed = true
		}
	}
	_, srcMayRaise := src.errors[ErrorMayRaise]
	if srcMayRaise && !hadMayRaise && dst.mayRaiseWitness == "" && owner != callee {
		dst.mayRaiseWitness = callee
		changed = true
	}
	return changed
}

func errorSummaryWithState(summary ErrorSummary, state *semanticState) ErrorSummary {
	if state == nil {
		return summary
	}
	summary.HasErrorHandler = hasErrorBehavior(state, ErrorHasHandler)
	summary.UsesResumeNext = hasErrorBehavior(state, ErrorUsesResumeNext)
	summary.SuppressesErrors = hasErrorBehavior(state, ErrorSuppresses)
	summary.RethrowsErrors = hasErrorBehavior(state, ErrorRethrows)
	summary.ReturnsSuccessFlag = hasErrorBehavior(state, ErrorReturnsSuccess)
	summary.MayRaise = hasErrorBehavior(state, ErrorMayRaise)
	summary.LogsAndContinues = hasErrorBehavior(state, ErrorLogsAndContinues)
	return summary
}

func hasErrorBehavior(state *semanticState, behavior ErrorBehaviorKind) bool {
	_, ok := state.errors[behavior]
	return ok
}

func (p ProjectSummary) reachableProcedureKeys(owner string) []string {
	if p.provenance == nil {
		return nil
	}
	seen := map[string]bool{owner: true}
	queue := append([]string(nil), p.provenance.callees[owner]...)
	for len(queue) > 0 {
		key := queue[0]
		queue = queue[1:]
		if seen[key] {
			continue
		}
		seen[key] = true
		queue = append(queue, p.provenance.callees[key]...)
	}
	out := make([]string, 0, len(seen)-1)
	for key := range seen {
		if key != owner {
			out = append(out, key)
		}
	}
	sort.Strings(out)
	return out
}

func (p ProjectSummary) propagatedEvidence(owner string, reachable []string) []Evidence {
	if p.provenance == nil {
		return nil
	}
	index := newMembershipIndex[Evidence](0, evidenceKey)
	var out []Evidence
	for _, key := range reachable {
		for _, evidence := range p.provenance.summaries[key].Direct {
			if evidence.Origin.Key() == owner || !index.add(evidence) {
				continue
			}
			out = append(out, evidence)
		}
	}
	sort.Slice(out, func(i, j int) bool { return evidenceKey(out[i]) < evidenceKey(out[j]) })
	return out
}

func (p ProjectSummary) propagatedUncertainty(owner string, reachable []string) []CallUncertainty {
	if p.provenance == nil {
		return nil
	}
	index := newMembershipIndex[CallUncertainty](0, uncertaintyKey)
	var out []CallUncertainty
	for _, key := range reachable {
		for _, uncertainty := range p.provenance.summaries[key].DirectUncertainty {
			if uncertainty.Origin.Key() == owner || !index.add(uncertainty) {
				continue
			}
			out = append(out, uncertainty)
		}
	}
	sort.Slice(out, func(i, j int) bool { return uncertaintyKey(out[i]) < uncertaintyKey(out[j]) })
	return out
}

func (p ProjectSummary) propagatedErrors(owner string, reachable []string, pathCache map[string]map[string][]string) []ErrorEvidence {
	if p.provenance == nil {
		return nil
	}
	index := newMembershipIndex[ErrorEvidence](0, errorEvidenceKey)
	var out []ErrorEvidence
	for _, key := range reachable {
		for _, evidence := range p.provenance.summaries[key].Error.Direct {
			if evidence.Behavior == ErrorMayRaise || evidence.Origin.Key() == owner {
				continue
			}
			origin := evidence.Origin.Key()
			paths, ok := pathCache[origin]
			if !ok {
				paths = p.errorPathsFromOrigin(origin)
				pathCache[origin] = paths
			}
			path, ok := paths[owner]
			if !ok {
				continue
			}
			copyEvidence := cloneErrorEvidence([]ErrorEvidence{evidence})[0]
			copyEvidence.CallChain = procedurePathIdentities(p.provenance.summaries, path)
			if len(copyEvidence.CallChain) == 0 || !index.add(copyEvidence) {
				continue
			}
			out = append(out, copyEvidence)
		}
	}
	if state, ok := p.provenance.summaries[owner]; ok && state.semantic != nil && state.semantic.mayRaiseWitness != "" {
		callee := state.semantic.mayRaiseWitness
		if calleeSummary, exists := p.provenance.summaries[callee]; exists {
			out = append(out, ErrorEvidence{
				Behavior: ErrorMayRaise, Origin: calleeSummary.Identity,
				CallChain: []ProcedureIdentity{state.Identity, calleeSummary.Identity},
				Target:    "callee", Value: calleeSummary.Identity.QualifiedName,
			})
		}
	}
	sort.Slice(out, func(i, j int) bool { return errorEvidenceKey(out[i]) < errorEvidenceKey(out[j]) })
	return out
}

func (p ProjectSummary) errorPathsFromOrigin(origin string) map[string][]string {
	if p.provenance == nil {
		return nil
	}
	keys := p.provenance.keys
	if len(keys) == 0 {
		keys = make([]string, 0, len(p.provenance.summaries))
		for key := range p.provenance.summaries {
			keys = append(keys, key)
		}
		sort.Strings(keys)
	}
	// Re-run the legacy first-arrival worklist for this origin. The eager
	// implementation selected a representative path when an ErrorEvidence key
	// first reached a caller; a shortest-path search can choose a different path
	// when a long branch is queued before a shorter branch becomes informative.
	// Keep every first-arrival path so all owners can reuse this one traversal.
	queue := append([]string(nil), keys...)
	queued := make(map[string]bool, len(keys))
	for _, key := range keys {
		queued[key] = true
	}
	paths := map[string][]string{origin: []string{origin}}
	for len(queue) > 0 {
		callee := queue[0]
		queue = queue[1:]
		queued[callee] = false
		path, hasPath := paths[callee]
		if !hasPath {
			continue
		}
		for _, caller := range p.provenance.callers[callee] {
			// Eager propagation excludes an origin from its own propagated
			// collection when a recursive edge closes the cycle.
			if caller == origin {
				continue
			}
			if _, exists := paths[caller]; exists {
				continue
			}
			nextPath := make([]string, 0, len(path)+1)
			nextPath = append(nextPath, caller)
			nextPath = append(nextPath, path...)
			paths[caller] = nextPath
			if !queued[caller] {
				queue = append(queue, caller)
				queued[caller] = true
			}
		}
	}
	return paths
}

func procedurePathIdentities(summaries map[string]ProcedureSummary, path []string) []ProcedureIdentity {
	out := make([]ProcedureIdentity, 0, len(path))
	for _, key := range path {
		if summary, ok := summaries[key]; ok {
			out = append(out, summary.Identity)
		}
	}
	return out
}
