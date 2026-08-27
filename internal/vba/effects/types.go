// Package effects computes deterministic, protocol-neutral procedure effect
// summaries from resolved procedure IR and its control-flow graphs.
package effects

import (
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"sync"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type EffectKind string

const (
	WritesCells        EffectKind = "writes_cells"
	ChangesWorkbook    EffectKind = "changes_workbook"
	OpensWorkbook      EffectKind = "opens_workbook"
	OpensFile          EffectKind = "opens_file"
	ClosesWorkbook     EffectKind = "closes_workbook"
	DisablesEvents     EffectKind = "disables_events"
	RestoresEvents     EffectKind = "restores_events"
	ChangesCalculation EffectKind = "changes_calculation"
	Recalculates       EffectKind = "recalculates"
	ChangesSelection   EffectKind = "changes_selection"
	ChangesControls    EffectKind = "changes_controls"
	ShowsDialog        EffectKind = "shows_dialog"
	LaunchesProcess    EffectKind = "launches_process"
	SuppressesErrors   EffectKind = "suppresses_errors"
	RaisesError        EffectKind = "raises_error"
	// ChangesApplicationState and RestoresApplicationState preserve the
	// property-neutral state contract needed by consumers such as VBA203.
	// Target identifies the exact Application property.
	ChangesApplicationState  EffectKind = "changes_application_state"
	RestoresApplicationState EffectKind = "restores_application_state"
)

// ProcedureIdentity is stable across input ordering. DeclarationLine
// disambiguates procedures with otherwise identical source identities.
type ProcedureIdentity struct {
	File            string
	Module          string
	ModuleKind      string
	Name            string
	QualifiedName   string
	Kind            procedureir.ProcedureKind
	Visibility      string
	IsEventHandler  bool
	EventKind       string
	DeclarationLine int
}

func (id ProcedureIdentity) Key() string {
	return strings.Join([]string{
		canonicalComparisonPath(id.File), strings.ToLower(id.Module),
		strings.ToLower(id.QualifiedName), string(id.Kind),
		decimal(id.DeclarationLine),
	}, "\x00")
}

type Evidence struct {
	Effect      EffectKind
	Origin      ProcedureIdentity
	Range       vbaast.Range
	StatementID int
	CallID      int
	Target      string
	Value       string
}

// ErrorFailureOutput describes a ByRef parameter that carries an explicit
// failure sentinel from a handled-error path.
type ErrorFailureOutput struct {
	ParameterIndex int
	ParameterName  string
	Value          string
}

type UncertaintyKind string

const (
	UncertaintyAmbiguous  UncertaintyKind = "ambiguous"
	UncertaintyUnresolved UncertaintyKind = "unresolved"
	UncertaintyExternal   UncertaintyKind = "external"
	UncertaintyDynamic    UncertaintyKind = "dynamic"
)

type CallUncertainty struct {
	Kind        UncertaintyKind
	Origin      ProcedureIdentity
	Range       vbaast.Range
	StatementID int
	CallID      int
	Callee      string
}

// ErrorBehaviorKind is a procedure-level runtime error behavior. The string
// values intentionally match the analyzer-facing summary vocabulary.
type ErrorBehaviorKind string

const (
	ErrorHasHandler       ErrorBehaviorKind = "has_error_handler"
	ErrorUsesResumeNext   ErrorBehaviorKind = "uses_resume_next"
	ErrorSuppresses       ErrorBehaviorKind = "suppresses_errors"
	ErrorRethrows         ErrorBehaviorKind = "rethrows_errors"
	ErrorReturnsSuccess   ErrorBehaviorKind = "returns_success_flag"
	ErrorMayRaise         ErrorBehaviorKind = "may_raise"
	ErrorLogsAndContinues ErrorBehaviorKind = "logs_and_continues"
)

// ErrorEvidence preserves the procedure and source statement at which an
// error behavior originated. Propagation copies evidence unchanged, allowing
// consumers to diagnose the procedure where failure information was lost.
type ErrorEvidence struct {
	Behavior ErrorBehaviorKind
	Origin   ProcedureIdentity
	// CallChain is one stable representative route from the summary owner to
	// Origin. Direct evidence contains only Origin; propagated evidence is
	// prefixed at each uniquely resolved caller boundary.
	CallChain      []ProcedureIdentity
	Range          vbaast.Range
	StatementID    int
	CallID         int
	Target         string
	Value          string
	FailureOutputs []ErrorFailureOutput
}

// ErrorSummary is derived from CFG error paths and error statements. The
// booleans cover both Direct and Propagated evidence; callers that need the
// local contract should inspect Direct and evidence Origin.
type ErrorSummary struct {
	HasErrorHandler    bool
	UsesResumeNext     bool
	SuppressesErrors   bool
	RethrowsErrors     bool
	ReturnsSuccessFlag bool
	MayRaise           bool
	LogsAndContinues   bool
	Direct             []ErrorEvidence
	Propagated         []ErrorEvidence
}

type ProcedureSummary struct {
	Identity              ProcedureIdentity
	Error                 ErrorSummary
	Direct                []Evidence
	Propagated            []Evidence
	DirectUncertainty     []CallUncertainty
	PropagatedUncertainty []CallUncertainty
	// semantic is the bounded fixed-point state used while building a
	// project. Propagated evidence is materialized from the project call graph
	// only when a caller asks for a full ProcedureSummary.
	semantic              *semanticState
	resolutionFingerprint string
}

func (s ProcedureSummary) Has(kind EffectKind) bool {
	if s.semantic != nil && s.semantic.hasEffect(kind) {
		return true
	}
	for _, evidence := range s.Direct {
		if evidence.Effect == kind {
			return true
		}
	}
	for _, evidence := range s.Propagated {
		if evidence.Effect == kind {
			return true
		}
	}
	return false
}

type ProjectSummary struct {
	procedures      []ProcedureSummary
	byKey           map[string]int
	byCandidateLine map[int][]int
	provenance      *provenanceGraph
	materialization *materializationCache
	// errorContractsFingerprint records the project-wide wrapper indexes used
	// while extracting direct error summaries. An incremental build must
	// invalidate direct summaries when these indexes change, even if the
	// affected procedure is not in the source-file change set.
	errorContractsFingerprint string
	stats                     BuildStats
}

// BuildStats describes the bounded fixed-point computation. Propagated fact
// counts refer to semantic slots, not to lazily reconstructed source
// evidence.
type BuildStats struct {
	WorklistEvaluations            uint64
	MaxPropagatedFactsPerProcedure uint64
	TotalPropagatedFacts           uint64
	// ReusedDirectProcedures and RecomputedDirectProcedures describe the
	// incremental direct-summary frontier. They remain useful for a full build
	// as well: a full build recomputes every procedure and reuses none.
	ReusedDirectProcedures     uint64
	RecomputedDirectProcedures uint64
}

type semanticState struct {
	effects             map[EffectKind]struct{}
	applicationChanges  map[string]struct{}
	applicationRestores map[string]struct{}
	errors              map[ErrorBehaviorKind]struct{}
	uncertainty         map[UncertaintyKind]struct{}
	mayRaiseWitness     string
}

func newSemanticState() *semanticState {
	return &semanticState{
		effects:             map[EffectKind]struct{}{},
		applicationChanges:  map[string]struct{}{},
		applicationRestores: map[string]struct{}{},
		errors:              map[ErrorBehaviorKind]struct{}{},
		uncertainty:         map[UncertaintyKind]struct{}{},
	}
}

func cloneSemanticState(in *semanticState) *semanticState {
	if in == nil {
		return nil
	}
	out := newSemanticState()
	for key := range in.effects {
		out.effects[key] = struct{}{}
	}
	for key := range in.applicationChanges {
		out.applicationChanges[key] = struct{}{}
	}
	for key := range in.applicationRestores {
		out.applicationRestores[key] = struct{}{}
	}
	for key := range in.errors {
		out.errors[key] = struct{}{}
	}
	for key := range in.uncertainty {
		out.uncertainty[key] = struct{}{}
	}
	out.mayRaiseWitness = in.mayRaiseWitness
	return out
}

func (s *semanticState) hasEffect(kind EffectKind) bool {
	if s == nil {
		return false
	}
	_, ok := s.effects[kind]
	return ok
}

func (s *semanticState) factCount() uint64 {
	if s == nil {
		return 0
	}
	return uint64(len(s.effects) + len(s.applicationChanges) + len(s.applicationRestores) + len(s.errors) + len(s.uncertainty))
}

type provenanceGraph struct {
	callers   map[string][]string
	callees   map[string][]string
	summaries map[string]ProcedureSummary
	witness   map[string]ProcedureSummary
	keys      []string
}

type provenanceMaterialization struct {
	project     ProjectSummary
	replayOnce  sync.Once
	errorReplay *errorWitnessReplay
}

type materializationCache struct {
	mu           sync.RWMutex
	materializer *provenanceMaterialization
	byIndex      map[int]ProcedureSummary
}

func newProvenanceMaterialization(project ProjectSummary) *provenanceMaterialization {
	return &provenanceMaterialization{
		project: project,
	}
}

func newMaterializationCache(project ProjectSummary) *materializationCache {
	return &materializationCache{
		materializer: newProvenanceMaterialization(project),
		byIndex:      map[int]ProcedureSummary{},
	}
}

func (p ProjectSummary) materializeCached(index int) ProcedureSummary {
	if p.materialization == nil {
		return p.materialize(index)
	}
	cache := p.materialization
	cache.mu.RLock()
	summary, ok := cache.byIndex[index]
	cache.mu.RUnlock()
	if ok {
		return cloneProcedureSummary(summary)
	}

	// Materialization can traverse a large call graph. Do not hold the cache
	// lock while doing that work; concurrent lookups may compute different
	// indexes independently and publish one immutable result each.
	summary = cache.materializer.materialize(index)
	cache.mu.Lock()
	if cached, ok := cache.byIndex[index]; ok {
		summary = cached
	} else {
		cache.byIndex[index] = summary
	}
	cache.mu.Unlock()
	return cloneProcedureSummary(summary)
}

func (m *provenanceMaterialization) errorWitnessReplay() *errorWitnessReplay {
	m.replayOnce.Do(func() { m.errorReplay = buildErrorWitnessReplay(m.project) })
	return m.errorReplay
}

func (p ProjectSummary) Lookup(id ProcedureIdentity) (ProcedureSummary, bool) {
	i, ok := p.byKey[id.Key()]
	if !ok {
		return ProcedureSummary{}, false
	}
	return p.materializeCached(i), true
}

// LookupDirect returns one defensive direct-only summary while retaining its
// compact semantic state for bounded Has/error-flag queries.
func (p ProjectSummary) LookupDirect(id ProcedureIdentity) (ProcedureSummary, bool) {
	i, ok := p.byKey[id.Key()]
	if !ok {
		return ProcedureSummary{}, false
	}
	return cloneProcedureSummary(p.procedures[i]), true
}

// LookupCandidate returns the first summary in deterministic procedure order
// that matches the resolver candidate. The line index narrows the search while
// the final comparisons preserve the resolver's case-insensitive matching
// contract, including Unicode case folding.
func (p ProjectSummary) LookupCandidate(candidate procedureir.Candidate) (ProcedureSummary, bool) {
	candidateFile := filepath.ToSlash(candidate.File)
	for _, i := range p.byCandidateLine[candidate.Line] {
		id := p.procedures[i].Identity
		if strings.EqualFold(filepath.ToSlash(id.File), candidateFile) &&
			strings.EqualFold(id.QualifiedName, candidate.QualifiedName) &&
			strings.EqualFold(string(id.Kind), candidate.Kind) &&
			id.DeclarationLine == candidate.Line {
			return p.materializeCached(i), true
		}
	}
	return ProcedureSummary{}, false
}

// LookupCandidateDirect is the direct-only counterpart of LookupCandidate.
func (p ProjectSummary) LookupCandidateDirect(candidate procedureir.Candidate) (ProcedureSummary, bool) {
	candidateFile := filepath.ToSlash(candidate.File)
	for _, i := range p.byCandidateLine[candidate.Line] {
		id := p.procedures[i].Identity
		if strings.EqualFold(filepath.ToSlash(id.File), candidateFile) &&
			strings.EqualFold(id.QualifiedName, candidate.QualifiedName) &&
			strings.EqualFold(string(id.Kind), candidate.Kind) &&
			id.DeclarationLine == candidate.Line {
			return cloneProcedureSummary(p.procedures[i]), true
		}
	}
	return ProcedureSummary{}, false
}

// AllDirect returns defensive direct-only summaries in deterministic order.
// The compact state remains attached so Has and error flags include the fixed
// point result without materializing provenance collections.
func (p ProjectSummary) AllDirect() []ProcedureSummary {
	out := make([]ProcedureSummary, len(p.procedures))
	for i := range p.procedures {
		out[i] = cloneProcedureSummary(p.procedures[i])
	}
	return out
}

// Identities returns procedure identities in the same deterministic order as
// AllDirect and All, without cloning any evidence or provenance collections.
func (p ProjectSummary) Identities() []ProcedureIdentity {
	out := make([]ProcedureIdentity, len(p.procedures))
	for i := range p.procedures {
		out[i] = p.procedures[i].Identity
	}
	return out
}

// All returns a defensive copy in deterministic procedure order.
func (p ProjectSummary) All() []ProcedureSummary {
	out := make([]ProcedureSummary, len(p.procedures))
	if p.materialization == nil {
		materializer := newProvenanceMaterialization(p)
		for i := range p.procedures {
			out[i] = materializer.materialize(i)
		}
		return out
	}
	for i := range p.procedures {
		out[i] = p.materializeCached(i)
	}
	return out
}

// ProcedureCount returns the number of compact procedure records without
// materializing transitive provenance.
func (p ProjectSummary) ProcedureCount() int { return len(p.procedures) }

// Stats returns fixed-point counters captured during BuildWithStats.
func (p ProjectSummary) Stats() BuildStats { return p.stats }

// RebindPaths returns a defensive summary copy whose display identities use
// the current source paths. Internal identity keys deliberately fold Windows
// casing, so a historical cache product can be reused across a case-only
// reopen without leaking the previous revision's path into diagnostics.
func (p ProjectSummary) RebindPaths(paths map[string]string) ProjectSummary {
	if len(paths) == 0 || len(p.procedures) == 0 {
		return p
	}
	displayByKey := make(map[string]string, len(paths))
	for path, display := range paths {
		if strings.TrimSpace(display) == "" {
			display = path
		}
		displayByKey[canonicalComparisonPath(path)] = canonicalPath(display)
	}
	currentIdentities := make(map[string]ProcedureIdentity, len(p.procedures))
	out := p
	out.procedures = make([]ProcedureSummary, len(p.procedures))
	for index, summary := range p.procedures {
		cloned := cloneProcedureSummary(summary)
		if display, ok := displayByKey[canonicalComparisonPath(cloned.Identity.File)]; ok {
			current := cloned.Identity
			current.File = display
			rebindProcedureSummaryIdentity(&cloned, current)
		}
		out.procedures[index] = cloned
		currentIdentities[cloned.Identity.Key()] = cloned.Identity
	}
	if p.provenance == nil {
		out.materialization = nil
		return out
	}
	out.provenance = &provenanceGraph{
		callers:   p.provenance.callers,
		callees:   p.provenance.callees,
		summaries: make(map[string]ProcedureSummary, len(p.provenance.summaries)),
		witness:   make(map[string]ProcedureSummary, len(p.provenance.witness)),
		keys:      append([]string(nil), p.provenance.keys...),
	}
	for key, summary := range p.provenance.summaries {
		cloned := cloneProcedureSummary(summary)
		if current, ok := currentIdentities[key]; ok {
			rebindProcedureSummaryIdentity(&cloned, current)
		}
		out.provenance.summaries[key] = cloned
	}
	for key, summary := range p.provenance.witness {
		cloned := cloneProcedureSummary(summary)
		if current, ok := currentIdentities[key]; ok {
			rebindProcedureSummaryIdentity(&cloned, current)
		}
		out.provenance.witness[key] = cloned
	}
	out.materialization = newMaterializationCache(out)
	return out
}

func (p ProjectSummary) materialize(index int) ProcedureSummary {
	return newProvenanceMaterialization(p).materialize(index)
}

func (m *provenanceMaterialization) materialize(index int) ProcedureSummary {
	p := m.project
	if index < 0 || index >= len(p.procedures) {
		return ProcedureSummary{}
	}
	out := cloneProcedureSummary(p.procedures[index])
	if p.provenance == nil {
		return out
	}
	key := out.Identity.Key()
	reachable := p.reachableProcedureKeys(key)
	out.Propagated = p.propagatedEvidence(key, reachable)
	out.PropagatedUncertainty = p.propagatedUncertainty(key, reachable)
	out.Error.Propagated = p.propagatedErrors(key, m.errorWitnessReplay())
	if out.semantic != nil {
		out.Error = errorSummaryWithState(out.Error, out.semantic)
	}
	return out
}

// Document pairs canonical syntax-local IR and an optional resolution view
// with the CFG built from that exact IR.
type Document struct {
	// IR is the canonical syntax-local document. Resolution, when present,
	// supplies project-dependent call/access/event facts without replacing IR
	// with a deep-copied resolved document.
	IR         procedureir.DocumentIR
	Resolution procedureir.ResolvedDocumentView
	CFG        cfg.Document
}

func identity(doc procedureir.DocumentIR, proc procedureir.ProcedureIR) ProcedureIdentity {
	return ProcedureIdentity{
		File: canonicalPath(doc.Path), Module: doc.ModuleName, ModuleKind: doc.ModuleKind,
		Name: proc.Symbol.Name, QualifiedName: proc.Symbol.QualifiedName, Kind: proc.Symbol.Kind,
		Visibility:      proc.Symbol.Visibility,
		IsEventHandler:  proc.Symbol.IsEventHandler,
		EventKind:       proc.Symbol.EventKind,
		DeclarationLine: proc.Symbol.DeclarationRange.StartLine,
	}
}

func cloneProcedureSummary(in ProcedureSummary) ProcedureSummary {
	in.Direct = append([]Evidence(nil), in.Direct...)
	in.Propagated = append([]Evidence(nil), in.Propagated...)
	in.DirectUncertainty = append([]CallUncertainty(nil), in.DirectUncertainty...)
	in.PropagatedUncertainty = append([]CallUncertainty(nil), in.PropagatedUncertainty...)
	in.Error.Direct = cloneErrorEvidence(in.Error.Direct)
	in.Error.Propagated = cloneErrorEvidence(in.Error.Propagated)
	return in
}

func cloneErrorEvidence(items []ErrorEvidence) []ErrorEvidence {
	out := append([]ErrorEvidence(nil), items...)
	for i := range out {
		out[i].CallChain = append([]ProcedureIdentity(nil), out[i].CallChain...)
		out[i].FailureOutputs = append([]ErrorFailureOutput(nil), out[i].FailureOutputs...)
	}
	return out
}

func canonicalPath(path string) string {
	return filepath.ToSlash(filepath.Clean(path))
}

// canonicalComparisonPath is used only for stable internal keys. Keep the
// display path in ProcedureIdentity.File so diagnostics and JSON preserve the
// source spelling while Windows comparisons remain case-insensitive.
func canonicalComparisonPath(path string) string {
	path = canonicalPath(path)
	if runtime.GOOS == "windows" {
		return strings.ToLower(path)
	}
	return path
}

func decimal(value int) string {
	if value == 0 {
		return "0"
	}
	const digits = "0123456789"
	var buf [24]byte
	i := len(buf)
	for value > 0 {
		i--
		buf[i] = digits[value%10]
		value /= 10
	}
	return string(buf[i:])
}

func sortSummaries(items []ProcedureSummary) {
	sort.Slice(items, func(i, j int) bool { return items[i].Identity.Key() < items[j].Identity.Key() })
	for i := range items {
		sort.Slice(items[i].Direct, func(a, b int) bool { return evidenceKey(items[i].Direct[a]) < evidenceKey(items[i].Direct[b]) })
		sort.Slice(items[i].Propagated, func(a, b int) bool { return evidenceKey(items[i].Propagated[a]) < evidenceKey(items[i].Propagated[b]) })
		sort.Slice(items[i].DirectUncertainty, func(a, b int) bool {
			return uncertaintyKey(items[i].DirectUncertainty[a]) < uncertaintyKey(items[i].DirectUncertainty[b])
		})
		sort.Slice(items[i].PropagatedUncertainty, func(a, b int) bool {
			return uncertaintyKey(items[i].PropagatedUncertainty[a]) < uncertaintyKey(items[i].PropagatedUncertainty[b])
		})
		sort.Slice(items[i].Error.Direct, func(a, b int) bool {
			return errorEvidenceKey(items[i].Error.Direct[a]) < errorEvidenceKey(items[i].Error.Direct[b])
		})
		sort.Slice(items[i].Error.Propagated, func(a, b int) bool {
			return errorEvidenceKey(items[i].Error.Propagated[a]) < errorEvidenceKey(items[i].Error.Propagated[b])
		})
	}
}

func evidenceKey(e Evidence) string {
	return strings.Join([]string{string(e.Effect), e.Origin.Key(), decimal(e.Range.StartByte), decimal(e.Range.EndByte), decimal(e.StatementID), decimal(e.CallID), strings.ToLower(e.Target), strings.ToLower(e.Value)}, "\x00")
}

func uncertaintyKey(u CallUncertainty) string {
	return strings.Join([]string{string(u.Kind), u.Origin.Key(), decimal(u.Range.StartByte), decimal(u.Range.EndByte), decimal(u.StatementID), decimal(u.CallID), strings.ToLower(u.Callee)}, "\x00")
}

func errorEvidenceKey(e ErrorEvidence) string {
	parts := []string{string(e.Behavior), e.Origin.Key(), decimal(e.Range.StartByte), decimal(e.Range.EndByte), decimal(e.StatementID), decimal(e.CallID), strings.ToLower(e.Target), strings.ToLower(e.Value)}
	for _, output := range e.FailureOutputs {
		parts = append(parts, decimal(output.ParameterIndex), strings.ToLower(output.ParameterName), strings.ToLower(output.Value))
	}
	return strings.Join(parts, "\x00")
}
