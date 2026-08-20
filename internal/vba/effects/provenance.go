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

type errorWitnessReplay struct {
	propagated map[string][]ErrorEvidence
}

type replayEvidence struct {
	key    string
	origin string
}

type replayUncertainty struct {
	key    string
	origin string
}

type replaySummary struct {
	identity       ProcedureIdentity
	evidence       []replayEvidence
	evidenceSet    map[string]struct{}
	errors         []ErrorEvidence
	errorSet       map[string]struct{}
	uncertainty    []replayUncertainty
	uncertaintySet map[string]struct{}
	mayRaise       bool
}

// buildErrorWitnessReplay replays the legacy global procedure worklist for
// error witnesses. Keeping one queue for every origin preserves the original
// first-arrival representative path, including branch requeues caused by a
// different origin, without retaining transitive evidence in fixed-point state.
func buildErrorWitnessReplay(project ProjectSummary) *errorWitnessReplay {
	replay := &errorWitnessReplay{propagated: map[string][]ErrorEvidence{}}
	if project.provenance == nil {
		return replay
	}
	keys := project.provenance.keys
	if len(keys) == 0 {
		keys = make([]string, 0, len(project.provenance.summaries))
		for key := range project.provenance.summaries {
			keys = append(keys, key)
		}
		sort.Strings(keys)
	}
	states := make(map[string]*replaySummary, len(keys))
	for _, key := range keys {
		summary, ok := project.provenance.witness[key]
		if !ok {
			summary = project.provenance.summaries[key]
		}
		state := &replaySummary{
			identity:       summary.Identity,
			evidenceSet:    map[string]struct{}{},
			errorSet:       map[string]struct{}{},
			uncertaintySet: map[string]struct{}{},
		}
		for _, evidence := range summary.Direct {
			fact := replayEvidence{key: evidenceKey(evidence), origin: evidence.Origin.Key()}
			if _, exists := state.evidenceSet[fact.key]; !exists {
				state.evidenceSet[fact.key] = struct{}{}
				state.evidence = append(state.evidence, fact)
			}
		}
		for _, evidence := range summary.Error.Direct {
			key := errorEvidenceKey(evidence)
			if _, exists := state.errorSet[key]; !exists {
				state.errorSet[key] = struct{}{}
				state.errors = append(state.errors, cloneErrorEvidence([]ErrorEvidence{evidence})[0])
			}
			if evidence.Behavior == ErrorMayRaise {
				state.mayRaise = true
			}
		}
		for _, uncertainty := range summary.DirectUncertainty {
			fact := replayUncertainty{key: uncertaintyKey(uncertainty), origin: uncertainty.Origin.Key()}
			if _, exists := state.uncertaintySet[fact.key]; !exists {
				state.uncertaintySet[fact.key] = struct{}{}
				state.uncertainty = append(state.uncertainty, fact)
			}
		}
		states[key] = state
	}

	queue := append([]string(nil), keys...)
	queued := make(map[string]bool, len(keys))
	for _, key := range keys {
		queued[key] = true
	}
	for len(queue) > 0 {
		calleeKey := queue[0]
		queue = queue[1:]
		queued[calleeKey] = false
		callee := states[calleeKey]
		if callee == nil {
			continue
		}
		for _, callerKey := range project.provenance.callers[calleeKey] {
			caller := states[callerKey]
			if caller == nil {
				continue
			}
			changed := false
			for _, fact := range callee.evidence {
				if fact.origin == callerKey {
					continue
				}
				if _, exists := caller.evidenceSet[fact.key]; exists {
					continue
				}
				caller.evidenceSet[fact.key] = struct{}{}
				caller.evidence = append(caller.evidence, fact)
				changed = true
			}
			for _, evidence := range callee.errors {
				if evidence.Behavior == ErrorMayRaise || evidence.Origin.Key() == callerKey {
					continue
				}
				copyEvidence := cloneErrorEvidence([]ErrorEvidence{evidence})[0]
				copyEvidence.CallChain = prependErrorCaller(caller.identity, copyEvidence.CallChain, copyEvidence.Origin)
				key := errorEvidenceKey(copyEvidence)
				if _, exists := caller.errorSet[key]; exists {
					continue
				}
				caller.errorSet[key] = struct{}{}
				caller.errors = append(caller.errors, copyEvidence)
				changed = true
			}
			if callee.mayRaise && !caller.mayRaise {
				copyEvidence := ErrorEvidence{
					Behavior: ErrorMayRaise,
					Origin:   callee.identity,
					CallChain: []ProcedureIdentity{
						caller.identity,
						callee.identity,
					},
					Target: "callee", Value: callee.identity.QualifiedName,
				}
				key := errorEvidenceKey(copyEvidence)
				if _, exists := caller.errorSet[key]; !exists {
					caller.errorSet[key] = struct{}{}
					caller.errors = append(caller.errors, copyEvidence)
					changed = true
				}
				caller.mayRaise = true
			}
			for _, fact := range callee.uncertainty {
				if fact.origin == callerKey {
					continue
				}
				if _, exists := caller.uncertaintySet[fact.key]; exists {
					continue
				}
				caller.uncertaintySet[fact.key] = struct{}{}
				caller.uncertainty = append(caller.uncertainty, fact)
				changed = true
			}
			if changed && !queued[callerKey] {
				queue = append(queue, callerKey)
				queued[callerKey] = true
			}
		}
	}

	for _, key := range keys {
		state := states[key]
		if state == nil {
			continue
		}
		for _, evidence := range state.errors {
			if evidence.Behavior == ErrorMayRaise || evidence.Origin.Key() == key {
				continue
			}
			replay.propagated[key] = append(replay.propagated[key], cloneErrorEvidence([]ErrorEvidence{evidence})[0])
		}
		sort.Slice(replay.propagated[key], func(i, j int) bool {
			return errorEvidenceKey(replay.propagated[key][i]) < errorEvidenceKey(replay.propagated[key][j])
		})
	}
	return replay
}

func prependErrorCaller(caller ProcedureIdentity, chain []ProcedureIdentity, origin ProcedureIdentity) []ProcedureIdentity {
	for _, item := range chain {
		if item.Key() == caller.Key() {
			return append([]ProcedureIdentity(nil), chain...)
		}
	}
	if len(chain) == 0 {
		chain = []ProcedureIdentity{origin}
	}
	out := make([]ProcedureIdentity, 0, len(chain)+1)
	out = append(out, caller)
	out = append(out, chain...)
	return out
}

func (p ProjectSummary) propagatedErrors(owner string, replay *errorWitnessReplay) []ErrorEvidence {
	if p.provenance == nil {
		return nil
	}
	index := newMembershipIndex[ErrorEvidence](0, errorEvidenceKey)
	var out []ErrorEvidence
	if replay != nil {
		for _, evidence := range replay.propagated[owner] {
			if index.add(evidence) {
				out = append(out, cloneErrorEvidence([]ErrorEvidence{evidence})[0])
			}
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
