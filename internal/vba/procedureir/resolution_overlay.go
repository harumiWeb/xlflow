package procedureir

import "strings"

// ProcedureIdentity identifies a procedure inside one immutable IR revision.
// The source range and ordinal disambiguate same-named conditional or
// overloaded declarations without using source text as a fact key.
type ProcedureIdentity struct {
	Path                 string
	Module               string
	ModuleKind           string
	QualifiedName        string
	Kind                 ProcedureKind
	DeclarationStartByte int
	Ordinal              int
}

// ResolutionFactKey is the stable key used by a resolution overlay. IDs are
// procedure-local and source ordered; the procedure identity prevents facts
// from one procedure being applied to another one after fragment reuse.
type ResolutionFactKey struct {
	Procedure ProcedureIdentity
	ID        int
}

// ResolvedCall is the project-dependent portion of a CallSite.
type ResolvedCall struct {
	Resolution CallResolution
}

// ResolvedAccess is the project-dependent portion of a VariableAccess. Scope
// is repeated because project resolution can refine an initially module or
// unresolved access scope.
type ResolvedAccess struct {
	Scope      SymbolScope
	Resolution SymbolResolution
}

// ResolvedEvent is the project-dependent portion of a RaiseEventReference.
type ResolvedEvent struct {
	Resolution SymbolResolution
}

// ResolutionOverlay contains only project-dependent facts. Its maps are kept
// private so callers can use the read-only view API without mutating the
// overlay accidentally.
type ResolutionOverlay struct {
	calls    map[ResolutionFactKey]ResolvedCall
	accesses map[ResolutionFactKey]ResolvedAccess
	events   map[ResolutionFactKey]ResolvedEvent
}

// ResolvedDocumentView combines immutable syntax-local IR with a read-only
// resolution overlay. The source IR is never cloned or mutated while the
// overlay is built.
type ResolvedDocumentView struct {
	document   DocumentIR
	overlay    ResolutionOverlay
	hasOverlay bool
	callIDs    [][]int
	accessIDs  [][]int
	eventIDs   [][]int
}

// HasOverlay reports whether this view contains project-dependent facts. It
// distinguishes a zero/compatibility view from a view built with a resolver.
func (v ResolvedDocumentView) HasOverlay() bool { return v.hasOverlay }

// ResolvedProcedure returns a procedure projection with project-dependent
// call, access, and RaiseEvent facts applied. Syntax-local collections such as
// declarations, statements, and expressions continue to alias the immutable
// revision. Only the small fact slices are copied, so this is not the full IR
// materialization performed by Materialize.
//
// The returned procedure and its fact slices are owned by the caller. The
// nested syntax-local values remain immutable revision data.
func (v ResolvedDocumentView) ResolvedProcedure(procedureIndex int) (ProcedureIR, bool) {
	if !v.hasOverlay {
		return ProcedureIR{}, false
	}
	if procedureIndex < 0 || procedureIndex >= len(v.document.Procedures) {
		return ProcedureIR{}, false
	}
	procedure := v.document.Procedures[procedureIndex]
	procedure.Calls = make([]CallSite, len(procedure.Calls))
	for callIndex, call := range v.document.Procedures[procedureIndex].Calls {
		call.Resolution = CallResolution{}
		if resolved, ok := v.resolvedCallAt(procedureIndex, callIndex); ok {
			call.Resolution = cloneCallResolution(resolved.Resolution)
		}
		if v.hasOverlay && !call.IsRaiseEvent && !isAssignmentTargetCall(call, procedure) {
			call.NonCallableNames = append([]string(nil), declarationNames(v.document.Declarations, procedure.Declarations)...)
		}
		procedure.Calls[callIndex] = call
	}
	procedure.Accesses = make([]VariableAccess, len(procedure.Accesses))
	for accessIndex, access := range v.document.Procedures[procedureIndex].Accesses {
		if resolved, ok := v.resolvedAccessAt(procedureIndex, accessIndex); ok {
			access.Scope = resolved.Scope
			access.Resolution = cloneSymbolResolution(resolved.Resolution)
		}
		procedure.Accesses[accessIndex] = access
	}
	procedure.RaiseEvents = make([]RaiseEventReference, len(procedure.RaiseEvents))
	for eventIndex, event := range v.document.Procedures[procedureIndex].RaiseEvents {
		if resolved, ok := v.resolvedEventAt(procedureIndex, eventIndex); ok {
			event.Resolution = cloneSymbolResolution(resolved.Resolution)
		}
		procedure.RaiseEvents[eventIndex] = event
	}
	return procedure, true
}

// ResolveView builds a read-only project-resolution view over in. The input
// document and all of its nested slices remain untouched.
func ResolveView(in DocumentIR, resolver Resolver) ResolvedDocumentView {
	return resolveViewWithFacts(in, resolver, documentFactIDs(in))
}

// ResolveViews builds two immutable resolution views over one document while
// sharing the revision-local fact IDs. The resolvers remain separate so
// procedure-only and full semantics cannot be conflated, while the document
// does not need to be scanned to assign IDs twice. Passing a nil resolver
// produces the same zero view as ResolveView.
func ResolveViews(in DocumentIR, procedureResolver, fullResolver Resolver) (ResolvedDocumentView, ResolvedDocumentView) {
	facts := documentFactIDs(in)
	return resolveViewWithFacts(in, procedureResolver, facts), resolveViewWithFacts(in, fullResolver, facts)
}

func resolveViewWithFacts(in DocumentIR, resolver Resolver, facts resolutionFactIDs) ResolvedDocumentView {
	view := ResolvedDocumentView{document: in, callIDs: facts.calls, accessIDs: facts.accesses, eventIDs: facts.events}
	if resolver == nil {
		return view
	}
	view.hasOverlay = true
	view.overlay = buildResolutionOverlay(in, resolver, view.callIDs, view.accessIDs, view.eventIDs)
	return view
}

func buildResolutionOverlay(
	in DocumentIR,
	resolver Resolver,
	callIDsByProcedure, accessIDsByProcedure, eventIDsByProcedure [][]int,
) ResolutionOverlay {
	overlay := ResolutionOverlay{
		calls:    make(map[ResolutionFactKey]ResolvedCall),
		accesses: make(map[ResolutionFactKey]ResolvedAccess),
		events:   make(map[ResolutionFactKey]ResolvedEvent),
	}
	eventResolver, hasEventResolver := resolver.(interface {
		ResolveEvent(SymbolReference) SymbolResolution
	})
	for procedureIndex, procedure := range in.Procedures {
		identity := procedureIdentity(in, procedureIndex)
		callIDs := callIDsByProcedure[procedureIndex]
		accessIDs := accessIDsByProcedure[procedureIndex]
		eventIDs := eventIDsByProcedure[procedureIndex]
		lexicalNonCallable := declarationNames(in.Declarations, procedure.Declarations)
		for callIndex, sourceCall := range procedure.Calls {
			if sourceCall.IsRaiseEvent {
				// RaiseEvent calls are filled from the syntax-local event facts
				// below. A generic Resolver cannot prove same-object visibility.
				continue
			}
			call := cloneCall(sourceCall)
			if !isAssignmentTargetCall(sourceCall, procedure) {
				call.NonCallableNames = append([]string(nil), lexicalNonCallable...)
			}
			overlay.calls[ResolutionFactKey{Procedure: identity, ID: callIDs[callIndex]}] = ResolvedCall{
				Resolution: cloneCallResolution(resolver.ResolveCall(call)),
			}
		}
		for eventIndex, event := range procedure.RaiseEvents {
			resolution := SymbolResolution{Scope: ScopeUnresolved, Status: ResolutionIncomplete}
			if hasEventResolver {
				resolution = eventResolver.ResolveEvent(SymbolReference{
					Name: event.Name, Module: in.ModuleName, Caller: event.Caller, Range: event.Range,
				})
			}
			overlay.events[ResolutionFactKey{Procedure: identity, ID: eventIDs[eventIndex]}] = ResolvedEvent{
				Resolution: cloneSymbolResolution(resolution),
			}
		}
		for callIndex, call := range procedure.Calls {
			if !call.IsRaiseEvent {
				continue
			}
			resolution := CallResolution{Status: ResolutionIncomplete}
			for eventIndex, event := range procedure.RaiseEvents {
				if !strings.EqualFold(event.Name, call.Callee.BaseName) {
					continue
				}
				resolved := overlay.events[ResolutionFactKey{Procedure: identity, ID: eventIDs[eventIndex]}]
				resolution = CallResolution{Status: resolved.Resolution.Status, Candidates: cloneCandidates(resolved.Resolution.Candidates)}
				break
			}
			overlay.calls[ResolutionFactKey{Procedure: identity, ID: callIDs[callIndex]}] = ResolvedCall{Resolution: resolution}
		}
		for accessIndex, access := range procedure.Accesses {
			scope := access.Scope
			resolution := SymbolResolution{Scope: scope}
			if access.Scope != ScopeUnresolved && access.Scope != ScopeProject {
				// Module-scope bindings historically bypass project resolution so
				// local variables continue to shadow imported symbols. Enum members
				// are the exception and are delegated to the canonical resolver.
				if access.Scope == ScopeModule {
					candidateResolution := resolver.ResolveSymbol(SymbolReference{
						Name: access.Name, Module: in.ModuleName,
						Caller: ProcedureRef{
							Name: procedure.Symbol.Name, Kind: procedure.Symbol.Kind,
							QualifiedName: procedure.Symbol.QualifiedName,
						}, Range: access.Range,
					})
					if allEnumCandidates(candidateResolution.Candidates) {
						resolution = candidateResolution
						scope = candidateResolution.Scope
					}
				}
			} else {
				resolution = cloneSymbolResolution(resolver.ResolveSymbol(SymbolReference{
					Name: access.Name, Module: in.ModuleName,
					Caller: ProcedureRef{
						Name: procedure.Symbol.Name, Kind: procedure.Symbol.Kind,
						QualifiedName: procedure.Symbol.QualifiedName,
					}, Range: access.Range,
				}))
				scope = resolution.Scope
				if scope == "" {
					scope = ScopeUnresolved
					resolution.Scope = ScopeUnresolved
				}
			}
			overlay.accesses[ResolutionFactKey{Procedure: identity, ID: accessIDs[accessIndex]}] = ResolvedAccess{
				Scope: scope, Resolution: cloneSymbolResolution(resolution),
			}
		}
	}
	return overlay
}

// ResolvedCall returns the overlay fact for a procedure-local call ID.
func (v ResolvedDocumentView) ResolvedCall(procedureIndex, callID int) (ResolvedCall, bool) {
	if procedureIndex < 0 || procedureIndex >= len(v.document.Procedures) {
		return ResolvedCall{}, false
	}
	key, ok := v.factKey(procedureIndex, callID, resolutionFactCall)
	if !ok {
		return ResolvedCall{}, false
	}
	resolved, ok := v.overlay.calls[key]
	if !ok {
		return ResolvedCall{}, false
	}
	resolved.Resolution = cloneCallResolution(resolved.Resolution)
	return resolved, true
}

// ResolvedAccess returns the overlay fact for a procedure-local access ID.
func (v ResolvedDocumentView) ResolvedAccess(procedureIndex, accessID int) (ResolvedAccess, bool) {
	if procedureIndex < 0 || procedureIndex >= len(v.document.Procedures) {
		return ResolvedAccess{}, false
	}
	key, ok := v.factKey(procedureIndex, accessID, resolutionFactAccess)
	if !ok {
		return ResolvedAccess{}, false
	}
	resolved, ok := v.overlay.accesses[key]
	if !ok {
		return ResolvedAccess{}, false
	}
	resolved.Resolution = cloneSymbolResolution(resolved.Resolution)
	return resolved, true
}

// ResolvedEvent returns the overlay fact for a procedure-local event ID.
func (v ResolvedDocumentView) ResolvedEvent(procedureIndex, eventID int) (ResolvedEvent, bool) {
	if procedureIndex < 0 || procedureIndex >= len(v.document.Procedures) {
		return ResolvedEvent{}, false
	}
	key, ok := v.factKey(procedureIndex, eventID, resolutionFactEvent)
	if !ok {
		return ResolvedEvent{}, false
	}
	resolved, ok := v.overlay.events[key]
	if !ok {
		return ResolvedEvent{}, false
	}
	resolved.Resolution = cloneSymbolResolution(resolved.Resolution)
	return resolved, true
}

// Materialize returns an independently owned resolved DocumentIR. This is the
// compatibility escape hatch for consumers that need to retain or mutate a
// resolved snapshot.
func (v ResolvedDocumentView) Materialize() DocumentIR {
	out := Clone(v.document)
	for procedureIndex := range out.Procedures {
		procedure := &out.Procedures[procedureIndex]
		for callIndex := range procedure.Calls {
			call := &procedure.Calls[callIndex]
			if resolved, ok := v.resolvedCallAt(procedureIndex, callIndex); ok {
				call.Resolution = cloneCallResolution(resolved.Resolution)
			}
			if !v.hasOverlay || call.IsRaiseEvent || isAssignmentTargetCall(*call, *procedure) {
				continue
			}
			call.NonCallableNames = append([]string(nil), declarationNames(out.Declarations, procedure.Declarations)...)
		}
		for eventIndex := range procedure.RaiseEvents {
			if resolved, ok := v.resolvedEventAt(procedureIndex, eventIndex); ok {
				procedure.RaiseEvents[eventIndex].Resolution = cloneSymbolResolution(resolved.Resolution)
			}
		}
		for accessIndex := range procedure.Accesses {
			if resolved, ok := v.resolvedAccessAt(procedureIndex, accessIndex); ok {
				procedure.Accesses[accessIndex].Scope = resolved.Scope
				procedure.Accesses[accessIndex].Resolution = cloneSymbolResolution(resolved.Resolution)
			}
		}
	}
	return out
}

type resolutionFactKind uint8

const (
	resolutionFactCall resolutionFactKind = iota
	resolutionFactAccess
	resolutionFactEvent
)

func (v ResolvedDocumentView) factKey(procedureIndex, id int, kind resolutionFactKind) (ResolutionFactKey, bool) {
	if procedureIndex < 0 || procedureIndex >= len(v.document.Procedures) {
		return ResolutionFactKey{}, false
	}
	procedure := v.document.Procedures[procedureIndex]
	if id < 0 {
		return ResolutionFactKey{}, false
	}
	ids := v.idsFor(procedureIndex, kind)
	if id > 0 {
		for _, candidate := range ids {
			if candidate == id {
				return v.factKeyAt(procedureIndex, candidate), true
			}
		}
		return ResolutionFactKey{}, false
	}
	// A zero ID occurs in hand-built IR. Resolve it only when there is one
	// zero-ID fact; otherwise the caller must use the deterministic ordinal.
	zero := -1
	for index := range ids {
		var raw int
		switch kind {
		case resolutionFactCall:
			raw = procedure.Calls[index].ID
		case resolutionFactAccess:
			raw = procedure.Accesses[index].ID
		case resolutionFactEvent:
			raw = procedure.RaiseEvents[index].ID
		}
		if raw == 0 {
			if zero >= 0 {
				return ResolutionFactKey{}, false
			}
			zero = index
		}
	}
	if zero < 0 {
		return ResolutionFactKey{}, false
	}
	return v.factKeyAt(procedureIndex, ids[zero]), true
}

func (v ResolvedDocumentView) factKeyAt(procedureIndex, id int) ResolutionFactKey {
	return ResolutionFactKey{Procedure: procedureIdentity(v.document, procedureIndex), ID: id}
}

func (v ResolvedDocumentView) resolvedCallAt(procedureIndex, factIndex int) (ResolvedCall, bool) {
	ids := v.idsFor(procedureIndex, resolutionFactCall)
	if factIndex < 0 || factIndex >= len(ids) {
		return ResolvedCall{}, false
	}
	resolved, ok := v.overlay.calls[ResolutionFactKey{Procedure: procedureIdentity(v.document, procedureIndex), ID: ids[factIndex]}]
	if !ok {
		return ResolvedCall{}, false
	}
	return resolved, true
}

func (v ResolvedDocumentView) resolvedAccessAt(procedureIndex, factIndex int) (ResolvedAccess, bool) {
	ids := v.idsFor(procedureIndex, resolutionFactAccess)
	if factIndex < 0 || factIndex >= len(ids) {
		return ResolvedAccess{}, false
	}
	resolved, ok := v.overlay.accesses[ResolutionFactKey{Procedure: procedureIdentity(v.document, procedureIndex), ID: ids[factIndex]}]
	return resolved, ok
}

func (v ResolvedDocumentView) resolvedEventAt(procedureIndex, factIndex int) (ResolvedEvent, bool) {
	ids := v.idsFor(procedureIndex, resolutionFactEvent)
	if factIndex < 0 || factIndex >= len(ids) {
		return ResolvedEvent{}, false
	}
	resolved, ok := v.overlay.events[ResolutionFactKey{Procedure: procedureIdentity(v.document, procedureIndex), ID: ids[factIndex]}]
	return resolved, ok
}

func (v ResolvedDocumentView) idsFor(procedureIndex int, kind resolutionFactKind) []int {
	switch kind {
	case resolutionFactCall:
		if procedureIndex >= len(v.callIDs) {
			return nil
		}
		return v.callIDs[procedureIndex]
	case resolutionFactAccess:
		if procedureIndex >= len(v.accessIDs) {
			return nil
		}
		return v.accessIDs[procedureIndex]
	case resolutionFactEvent:
		if procedureIndex >= len(v.eventIDs) {
			return nil
		}
		return v.eventIDs[procedureIndex]
	default:
		return nil
	}
}

type resolutionFactIDs struct {
	calls    [][]int
	accesses [][]int
	events   [][]int
}

func documentFactIDs(document DocumentIR) resolutionFactIDs {
	calls := make([][]int, len(document.Procedures))
	accesses := make([][]int, len(document.Procedures))
	events := make([][]int, len(document.Procedures))
	for procedureIndex, procedure := range document.Procedures {
		calls[procedureIndex] = factIDs(procedure, resolutionFactCall)
		accesses[procedureIndex] = factIDs(procedure, resolutionFactAccess)
		events[procedureIndex] = factIDs(procedure, resolutionFactEvent)
	}
	return resolutionFactIDs{calls: calls, accesses: accesses, events: events}
}

func procedureIdentity(document DocumentIR, procedureIndex int) ProcedureIdentity {
	procedure := document.Procedures[procedureIndex]
	return ProcedureIdentity{
		Path: document.Path, Module: document.ModuleName, ModuleKind: document.ModuleKind,
		QualifiedName: procedure.Symbol.QualifiedName, Kind: procedure.Symbol.Kind,
		DeclarationStartByte: procedure.Symbol.DeclarationRange.StartByte, Ordinal: procedureIndex,
	}
}

func factIDs(procedure ProcedureIR, kind resolutionFactKind) []int {
	var count int
	switch kind {
	case resolutionFactCall:
		count = len(procedure.Calls)
	case resolutionFactAccess:
		count = len(procedure.Accesses)
	case resolutionFactEvent:
		count = len(procedure.RaiseEvents)
	}
	ids := make([]int, count)
	used := make(map[int]struct{}, count)
	for index := 0; index < count; index++ {
		var raw int
		switch kind {
		case resolutionFactCall:
			raw = procedure.Calls[index].ID
		case resolutionFactAccess:
			raw = procedure.Accesses[index].ID
		case resolutionFactEvent:
			raw = procedure.RaiseEvents[index].ID
		}
		if raw <= 0 {
			raw = index + 1
		}
		if _, exists := used[raw]; exists {
			raw = index + 1
			for {
				if _, exists := used[raw]; !exists {
					break
				}
				raw++
			}
		}
		used[raw] = struct{}{}
		ids[index] = raw
	}
	return ids
}
