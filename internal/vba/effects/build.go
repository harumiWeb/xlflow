package effects

import (
	"crypto/sha256"
	"encoding/hex"
	"runtime"
	"sort"
	"strconv"
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
	return buildWithReuse(documents, nil, nil)
}

// BuildIncremental reuses immutable direct summaries from previous when the
// supplied changed-file set does not contain their owner. The reverse caller
// closure is recomputed from both the previous and current graphs so removed
// or redirected calls cannot leave stale propagated state behind.
//
// changedFiles must contain canonical or equivalent source paths. An empty
// set is a valid no-op revision and returns a structurally equivalent copy of
// previous without rebuilding its semantic kernels.
func BuildIncremental(documents []Document, previous ProjectSummary, changedFiles map[string]struct{}) ProjectSummary {
	project, _ := BuildIncrementalWithStats(documents, previous, changedFiles)
	return project
}

// BuildIncrementalWithStats is BuildIncremental with developer-facing
// counters for direct-summary reuse and fixed-point propagation.
func BuildIncrementalWithStats(documents []Document, previous ProjectSummary, changedFiles map[string]struct{}) (ProjectSummary, BuildStats) {
	if len(previous.procedures) == 0 {
		return BuildWithStats(documents)
	}
	return buildWithReuse(documents, &previous, changedFiles)
}

func buildWithReuse(documents []Document, previous *ProjectSummary, changedFiles map[string]struct{}) (ProjectSummary, BuildStats) {
	changedFiles = normalizeChangedFiles(changedFiles)
	inputs := collectInputs(documents)
	summaries := make(map[string]*ProcedureSummary, len(inputs))
	candidateKeys := candidateIndex(inputs)
	loggerTargets := loggerProcedureIndex(inputs, candidateKeys)
	rethrowTargets := rethrowProcedureIndex(inputs, candidateKeys)
	terminalTargets := terminalProcedureIndex(inputs, candidateKeys)
	errorContractsKey := errorContractsFingerprint(loggerTargets, rethrowTargets, terminalTargets)
	globalErrorContractInvalidation := previous != nil && previous.errorContractsFingerprint != errorContractsKey
	var edges []edge
	changedKeys := make(map[string]struct{})
	var directStats BuildStats
	for _, input := range inputs {
		key := input.id.Key()
		resolutionFingerprint := procedureResolutionFingerprint(input.proc)
		// The error wrapper indexes are project-wide dependencies of direct
		// extraction. If one changes, recompute every current direct summary so
		// an unchanged caller cannot retain stale suppression/rethrow evidence.
		changed := globalErrorContractInvalidation
		if changedPathMatches(changedFiles, input.id.File) {
			changed = true
		}
		if _, ok := changedFiles[key]; ok {
			changed = true
		}
		if previous != nil && !changed {
			if index, ok := previous.byKey[key]; ok && index >= 0 && index < len(previous.procedures) && previous.procedures[index].resolutionFingerprint != resolutionFingerprint {
				// A project declaration edit can change this procedure's resolved
				// calls/accesses even when its own source file is unchanged.
				changed = true
			}
		}
		if changed {
			changedKeys[key] = struct{}{}
		}
		if previous != nil && !changed {
			if index, ok := previous.byKey[key]; ok && index >= 0 && index < len(previous.procedures) {
				summary := cloneProcedureSummary(previous.procedures[index])
				rebindProcedureSummaryIdentity(&summary, input.id)
				summary.semantic = cloneSemanticState(summary.semantic)
				summaries[key] = &summary
			}
		}
		if _, reused := summaries[key]; reused {
			directStats.ReusedDirectProcedures++
			// The current graph is still collected below. Only the direct
			// semantic extraction is skipped for a matching immutable input.
			continue
		}
		directStats.RecomputedDirectProcedures++
		summary := &ProcedureSummary{Identity: input.id, resolutionFingerprint: resolutionFingerprint}
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
		summaries[key] = summary
	}
	if previous != nil {
		// A deleted procedure has no current input from which to seed the red
		// set. Preserve its old identity long enough to walk old reverse edges
		// and reset callers that depended on it.
		for _, summary := range previous.procedures {
			if changedPathMatches(changedFiles, summary.Identity.File) {
				changedKeys[summary.Identity.Key()] = struct{}{}
			}
		}
	}
	// Edges are cheap to collect from the current resolved inputs and must be
	// refreshed even when a procedure's direct kernel was reused.
	for _, input := range inputs {
		statements := statementIndex(input.proc)
		for _, call := range input.proc.Calls {
			statement := statements[call.StatementID]
			if !input.reachable[call.StatementID] || statement.Recovered || call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
				continue
			}
			if target, ok := candidateKeys[candidateKey(call.Resolution.Candidates[0])]; ok {
				edges = append(edges, edge{from: input.id.Key(), to: target})
			}
		}
	}
	sort.Slice(edges, func(i, j int) bool {
		if edges[i].from != edges[j].from {
			return edges[i].from < edges[j].from
		}
		return edges[i].to < edges[j].to
	})
	callers, callees := buildAdjacency(edges)
	if previous == nil {
		stats := propagateBounded(summaries, callers)
		stats.ReusedDirectProcedures = directStats.ReusedDirectProcedures
		stats.RecomputedDirectProcedures = directStats.RecomputedDirectProcedures
		return assembleProjectSummary(summaries, callers, callees, stats, errorContractsKey)
	}

	// Start the red set with directly changed procedures and traverse both
	// reverse graphs. Any caller in that closure must be reset to its direct
	// state before propagation, otherwise removed effects would remain sticky.
	red := make(map[string]struct{}, len(changedKeys))
	queue := make([]string, 0, len(changedKeys))
	for key := range changedKeys {
		red[key] = struct{}{}
		queue = append(queue, key)
	}
	for len(queue) > 0 {
		owner := queue[0]
		queue = queue[1:]
		neighborSet := make(map[string]struct{}, len(callers[owner]))
		for _, caller := range callers[owner] {
			neighborSet[caller] = struct{}{}
		}
		if previous.provenance != nil {
			for _, caller := range previous.provenance.callers[owner] {
				neighborSet[caller] = struct{}{}
			}
		}
		neighbors := make([]string, 0, len(neighborSet))
		for neighbor := range neighborSet {
			neighbors = append(neighbors, neighbor)
		}
		sort.Strings(neighbors)
		for _, neighbor := range neighbors {
			if _, exists := red[neighbor]; exists {
				continue
			}
			red[neighbor] = struct{}{}
			queue = append(queue, neighbor)
		}
	}
	// A reset caller must also be scheduled after its current and previous
	// callees are processed. These dependency seeds feed the caller's direct
	// state without widening the reset set to every downstream helper.
	seeds := make(map[string]struct{}, len(red))
	for key := range red {
		seeds[key] = struct{}{}
		for _, callee := range callees[key] {
			seeds[callee] = struct{}{}
		}
		if previous.provenance != nil {
			for _, callee := range previous.provenance.callees[key] {
				seeds[callee] = struct{}{}
			}
		}
	}
	for key := range red {
		if summary := summaries[key]; summary != nil {
			summary.semantic = semanticStateFromSummary(summary)
			summary.Error = errorSummaryWithState(summary.Error, summary.semantic)
		}
	}
	stats := propagateBoundedSeeded(summaries, callers, seeds)
	stats.ReusedDirectProcedures = directStats.ReusedDirectProcedures
	stats.RecomputedDirectProcedures = directStats.RecomputedDirectProcedures
	return assembleProjectSummary(summaries, callers, callees, stats, errorContractsKey)
}

// rebindProcedureSummaryIdentity updates display identities on a reused direct
// summary when a Windows path is reopened with different casing. The internal
// key intentionally folds that casing, but diagnostics must retain the current
// source spelling rather than leaking the previous revision's path.
func rebindProcedureSummaryIdentity(summary *ProcedureSummary, current ProcedureIdentity) {
	if summary == nil || summary.Identity.Key() != current.Key() {
		return
	}
	previous := summary.Identity
	rebind := func(identity *ProcedureIdentity) {
		if identity != nil && identity.Key() == previous.Key() {
			*identity = current
		}
	}
	rebind(&summary.Identity)
	for index := range summary.Direct {
		rebind(&summary.Direct[index].Origin)
	}
	for index := range summary.Propagated {
		rebind(&summary.Propagated[index].Origin)
	}
	for index := range summary.DirectUncertainty {
		rebind(&summary.DirectUncertainty[index].Origin)
	}
	for index := range summary.PropagatedUncertainty {
		rebind(&summary.PropagatedUncertainty[index].Origin)
	}
	for index := range summary.Error.Direct {
		rebind(&summary.Error.Direct[index].Origin)
		for chainIndex := range summary.Error.Direct[index].CallChain {
			rebind(&summary.Error.Direct[index].CallChain[chainIndex])
		}
	}
	for index := range summary.Error.Propagated {
		rebind(&summary.Error.Propagated[index].Origin)
		for chainIndex := range summary.Error.Propagated[index].CallChain {
			rebind(&summary.Error.Propagated[index].CallChain[chainIndex])
		}
	}
}

func normalizeChangedFiles(in map[string]struct{}) map[string]struct{} {
	if len(in) == 0 {
		return in
	}
	out := make(map[string]struct{}, len(in)*2)
	for key := range in {
		out[key] = struct{}{}
		canonical := canonicalComparisonPath(key)
		out[canonical] = struct{}{}
	}
	return out
}

func changedPathMatches(changed map[string]struct{}, path string) bool {
	canonical := canonicalComparisonPath(path)
	if _, ok := changed[canonical]; ok {
		return true
	}
	if runtime.GOOS != "windows" {
		return false
	}
	_, ok := changed[strings.ToLower(canonical)]
	return ok
}

func assembleProjectSummary(summaries map[string]*ProcedureSummary, callers map[string][]string, callees map[string][]string, stats BuildStats, errorContractsKey string) (ProjectSummary, BuildStats) {
	witnessSummaries := make(map[string]ProcedureSummary, len(summaries))
	for key, summary := range summaries {
		if summary != nil {
			// Keep extraction order for lazy compatibility replay. The public
			// summaries are sorted below, but the legacy propagation worklist saw
			// these direct slices before that presentation-only sort.
			witnessSummaries[key] = cloneProcedureSummary(*summary)
		}
	}
	out := ProjectSummary{
		byKey:                     map[string]int{},
		byCandidateLine:           map[int][]int{},
		errorContractsFingerprint: errorContractsKey,
		stats:                     stats,
		provenance: &provenanceGraph{
			callers:   callers,
			callees:   callees,
			summaries: map[string]ProcedureSummary{},
			witness:   witnessSummaries,
		},
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
	out.provenance.keys = make([]string, 0, len(out.provenance.summaries))
	for _, summary := range out.procedures {
		out.provenance.keys = append(out.provenance.keys, summary.Identity.Key())
	}
	out.materialization = newMaterializationCache(out)
	return out, stats
}

func collectInputs(documents []Document) []procedureInput {
	var out []procedureInput
	for _, doc := range documents {
		graphs := map[string]cfg.Graph{}
		for _, graph := range doc.CFG.Graphs {
			graphs[graph.Procedure.QualifiedName+"\x00"+string(graph.Procedure.Kind)] = graph
		}
		for procedureIndex, sourceProc := range doc.IR.Procedures {
			proc := sourceProc
			if resolved, ok := doc.Resolution.ResolvedProcedure(procedureIndex); ok {
				proc = resolved
			}
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
		key := strings.Join([]string{canonicalComparisonPath(input.id.File), strings.ToLower(input.id.QualifiedName), string(input.id.Kind), decimal(input.id.DeclarationLine)}, "\x00")
		out[key] = input.id.Key()
	}
	return out
}

func candidateKey(c procedureir.Candidate) string {
	return strings.Join([]string{canonicalComparisonPath(c.File), strings.ToLower(c.QualifiedName), strings.ToLower(c.Kind), decimal(c.Line)}, "\x00")
}

// procedureResolutionFingerprint captures the project-dependent facts that
// direct extraction consumes. The owning source file is still the primary
// invalidation input, but a declaration edit in another file can change an
// unchanged caller's call/access/event resolutions. Candidate paths go through
// canonicalPath so Windows case-only path changes do not defeat reuse.
func procedureResolutionFingerprint(proc procedureir.ProcedureIR) string {
	var builder strings.Builder
	write := func(values ...string) {
		for _, value := range values {
			builder.WriteString(value)
			builder.WriteByte('\x00')
		}
	}
	writeCandidate := func(candidate procedureir.Candidate) {
		write("candidate", canonicalComparisonPath(candidate.File), strings.ToLower(candidate.QualifiedName), strings.ToLower(candidate.Kind), strconv.Itoa(candidate.Line))
	}
	writeSymbolResolution := func(resolution procedureir.SymbolResolution) {
		write("symbol-resolution", string(resolution.Status), string(resolution.Scope), strconv.Itoa(len(resolution.Candidates)))
		for _, candidate := range resolution.Candidates {
			writeCandidate(candidate)
		}
	}
	write("procedure-resolution-v1", strconv.Itoa(len(proc.Calls)), strconv.Itoa(len(proc.Accesses)), strconv.Itoa(len(proc.RaiseEvents)))
	for _, call := range proc.Calls {
		write("call", strconv.Itoa(call.ID), strconv.Itoa(call.StatementID), string(call.Resolution.Status), strconv.FormatBool(call.Resolution.ProjectLocal), strconv.Itoa(len(call.Resolution.Candidates)))
		for _, candidate := range call.Resolution.Candidates {
			writeCandidate(candidate)
		}
	}
	for _, access := range proc.Accesses {
		write("access", strconv.Itoa(access.ID), access.Name, string(access.Mode), string(access.Scope), strconv.Itoa(access.StatementID), strconv.Itoa(access.ExpressionID))
		writeSymbolResolution(access.Resolution)
	}
	for _, event := range proc.RaiseEvents {
		write("raise-event", strconv.Itoa(event.ID), event.Name, event.Module, strconv.FormatBool(event.Recovered))
		writeSymbolResolution(event.Resolution)
	}
	hash := sha256.Sum256([]byte(builder.String()))
	return hex.EncodeToString(hash[:])
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

func buildAdjacency(edges []edge) (map[string][]string, map[string][]string) {
	callerSets := make(map[string]map[string]struct{})
	calleeSets := make(map[string]map[string]struct{})
	for _, item := range edges {
		callers := callerSets[item.to]
		if callers == nil {
			callers = map[string]struct{}{}
			callerSets[item.to] = callers
		}
		callers[item.from] = struct{}{}

		callees := calleeSets[item.from]
		if callees == nil {
			callees = map[string]struct{}{}
			calleeSets[item.from] = callees
		}
		callees[item.to] = struct{}{}
	}
	callers := make(map[string][]string, len(callerSets))
	for key, set := range callerSets {
		for caller := range set {
			callers[key] = append(callers[key], caller)
		}
		sort.Strings(callers[key])
	}
	callees := make(map[string][]string, len(calleeSets))
	for key, set := range calleeSets {
		for callee := range set {
			callees[key] = append(callees[key], callee)
		}
		sort.Strings(callees[key])
	}
	return callers, callees
}

// propagateBounded propagates only finite semantic state.
func propagateBounded(summaries map[string]*ProcedureSummary, callers map[string][]string) BuildStats {
	seeds := make(map[string]struct{}, len(summaries))
	for key := range summaries {
		seeds[key] = struct{}{}
	}
	return propagateBoundedSeeded(summaries, callers, seeds)
}

// propagateBoundedSeeded is the incremental counterpart of propagateBounded.
// Only the red dependency closure is scheduled; callers outside that closure
// retain their immutable fixed-point state.
func propagateBoundedSeeded(summaries map[string]*ProcedureSummary, callers map[string][]string, seeds map[string]struct{}) BuildStats {
	queue := make([]string, 0, len(summaries))
	queued := map[string]bool{}
	for key := range seeds {
		if _, exists := summaries[key]; !exists {
			continue
		}
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
