package procedureir

import (
	"path/filepath"
	"sort"
	"strings"
)

// Resolve returns a deep copy with project-dependent call and symbol
// resolutions applied. Syntax-local parameter/local/module scopes are kept.
// Consumers that only need read-only resolution should use ResolveView to
// avoid cloning the full procedure payload.
func Resolve(in DocumentIR, resolver Resolver) DocumentIR {
	return ResolveView(in, resolver).Materialize()
}

func allEnumCandidates(candidates []Candidate) bool {
	if len(candidates) == 0 {
		return false
	}
	for _, candidate := range candidates {
		if !isEnumMemberKind(candidate.Kind) {
			return false
		}
	}
	return true
}

func isAssignmentTargetCall(call CallSite, procedure ProcedureIR) bool {
	if call.StatementID <= 0 || call.StatementID > len(procedure.Statements) {
		return false
	}
	statement := procedure.Statements[call.StatementID-1]
	if statement.Kind == StatementReDim {
		return true
	}
	if statement.Kind != StatementAssignment && statement.Kind != StatementSet {
		return false
	}
	if statement.Target == nil {
		return false
	}
	return strings.EqualFold(cleanIdentifier(statement.Target.Text), cleanIdentifier(call.Callee.BaseName))
}

func declarationNames(module, procedure []Declaration) []string {
	result := make([]string, 0, len(module)+len(procedure))
	appendDeclaration := func(declaration Declaration) {
		if strings.TrimSpace(declaration.Name) == "" || strings.EqualFold(declaration.Kind, "return_slot") {
			return
		}
		if len(declaration.ConditionalBranches) > 0 || declaration.Recovered {
			return
		}
		// An indexed array access is represented by the grammar as a call
		// expression. It is not a procedure invocation and must not become a
		// VB052 non-callable target merely because the array is in lexical scope.
		if declaration.IsArray {
			return
		}
		typeName := strings.ToLower(strings.TrimSpace(declaration.Type))
		// Variant/Object values may be late-bound or default-member calls and
		// therefore cannot prove a non-callable target. Untyped declarations are
		// implicit Variant under VBA's rules as well.
		if declaration.IsObject || typeName == "" || typeName == "variant" || typeName == "object" ||
			(!isKnownScalarType(typeName) && !strings.EqualFold(declaration.Kind, "const") && !strings.EqualFold(declaration.Kind, "enum_member")) {
			return
		}
		// A parameter typed by an unavailable external library can expose a
		// default member or late-bound callable surface. Keep it conservative;
		// scalar project-local parameters still participate in the proof.
		if strings.EqualFold(declaration.Kind, "parameter") && strings.Contains(typeName, ".") {
			return
		}
		switch strings.ToLower(declaration.Kind) {
		case "variable", "const", "enum_member", "type", "enum", "parameter", "variable_declaration", "const_declaration":
			result = append(result, declaration.Name)
		}
	}
	for _, declaration := range module {
		appendDeclaration(declaration)
	}
	for _, declaration := range procedure {
		appendDeclaration(declaration)
	}
	return result
}

func isKnownScalarType(typeName string) bool {
	switch strings.ToLower(strings.TrimSpace(typeName)) {
	case "byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal", "date", "boolean", "string":
		return true
	default:
		return false
	}
}

func cloneCandidates(in []Candidate) []Candidate {
	return append([]Candidate(nil), in...)
}

func cloneCallResolution(in CallResolution) CallResolution {
	in.Candidates = append([]Candidate(nil), in.Candidates...)
	return in
}

func cloneSymbolResolution(in SymbolResolution) SymbolResolution {
	in.Candidates = append([]Candidate(nil), in.Candidates...)
	return in
}

type resolverEntry struct {
	Candidate
	typeName            string
	module              string
	moduleKind          string
	visibility          string
	parent              string
	recovered           bool
	conditionalBranches []ConditionalBranch
	isArray             bool
	isConst             bool
	valueShape          ValueShapeKind
}

// SymbolResolver is a deterministic protocol-neutral project resolver.
type SymbolResolver struct {
	byName      map[string][]resolverEntry
	modules     map[string]struct{}
	moduleKinds map[string]string
	complete    bool
	// procedureByName/modules/moduleKinds retain the procedure-only view that
	// the effects and procedure-local analyzers historically consumed. The
	// full resolver remains the canonical Resolution capability; callers that
	// need the legacy call graph semantics can request an immutable view over
	// these same indexes without rebuilding the project symbol table.
	procedureByName      map[string][]resolverEntry
	procedureModules     map[string]struct{}
	procedureModuleKinds map[string]string
}

func NewSymbolResolver(symbols []ResolverSymbol) SymbolResolver {
	return NewSymbolResolverWithCompleteness(symbols, true)
}

// NewSymbolResolverWithCompleteness creates a resolver over a project
// snapshot. A false completeness value is deliberately fail-open: positive
// matches remain usable by call graph consumers, while negative outcomes are
// returned as incomplete so diagnostics cannot mistake a partial index for a
// proven compile error.
func NewSymbolResolverWithCompleteness(symbols []ResolverSymbol, complete bool) SymbolResolver {
	out := SymbolResolver{
		byName:               map[string][]resolverEntry{},
		modules:              map[string]struct{}{},
		moduleKinds:          map[string]string{},
		complete:             complete,
		procedureByName:      map[string][]resolverEntry{},
		procedureModules:     map[string]struct{}{},
		procedureModuleKinds: map[string]string{},
	}
	for _, symbol := range symbols {
		if strings.TrimSpace(symbol.Name) == "" {
			continue
		}
		qualified := symbol.Name
		if symbol.Module != "" {
			qualified = symbol.Module + "." + symbol.Name
		}
		entry := resolverEntry{
			Candidate: Candidate{
				QualifiedName: qualified, Kind: symbol.Kind,
				File: normalizeCandidateFile(symbol.File), Line: symbol.Line,
			},
			typeName: symbol.Type,
			module:   symbol.Module, moduleKind: symbol.ModuleKind, visibility: symbol.Visibility,
			parent: symbol.Parent, recovered: symbol.Recovered,
			isArray:             symbol.IsArray,
			isConst:             symbol.IsConst,
			valueShape:          symbol.ValueShape,
			conditionalBranches: append([]ConditionalBranch(nil), symbol.ConditionalBranches...),
		}
		key := strings.ToLower(cleanIdentifier(symbol.Name))
		out.byName[key] = append(out.byName[key], entry)
		if isLegacyProjectProcedureKind(symbol.Kind) {
			// Match the legacy effects resolver, which indexed only the public
			// procedure identity fields and therefore never treated recovered or
			// conditional procedure metadata as uncertainty evidence.
			procedureEntry := entry
			procedureEntry.typeName = ""
			procedureEntry.parent = ""
			procedureEntry.recovered = false
			procedureEntry.conditionalBranches = nil
			procedureEntry.isArray = false
			procedureEntry.isConst = false
			procedureEntry.valueShape = ValueShapeUnknown
			out.procedureByName[key] = append(out.procedureByName[key], procedureEntry)
		}
		if strings.TrimSpace(symbol.Module) != "" {
			moduleKey := strings.ToLower(cleanIdentifier(symbol.Module))
			out.modules[moduleKey] = struct{}{}
			if _, exists := out.moduleKinds[moduleKey]; !exists && strings.TrimSpace(symbol.ModuleKind) != "" {
				out.moduleKinds[moduleKey] = strings.ToLower(strings.TrimSpace(symbol.ModuleKind))
			}
			if isLegacyProjectProcedureKind(symbol.Kind) {
				out.procedureModules[moduleKey] = struct{}{}
				if _, exists := out.procedureModuleKinds[moduleKey]; !exists && strings.TrimSpace(symbol.ModuleKind) != "" {
					out.procedureModuleKinds[moduleKey] = strings.ToLower(strings.TrimSpace(symbol.ModuleKind))
				}
			}
		}
	}
	for key := range out.byName {
		sortResolverEntries(out.byName[key])
	}
	for key := range out.procedureByName {
		sortResolverEntries(out.procedureByName[key])
	}
	return out
}

func sortResolverEntries(entries []resolverEntry) {
	sort.Slice(entries, func(i, j int) bool {
		a, b := entries[i], entries[j]
		if a.QualifiedName != b.QualifiedName {
			return a.QualifiedName < b.QualifiedName
		}
		if a.Kind != b.Kind {
			return a.Kind < b.Kind
		}
		if a.File != b.File {
			return a.File < b.File
		}
		return a.Line < b.Line
	})
}

// isLegacyProjectProcedureKind mirrors the procedure IR list used by the
// pre-planner effects resolver. Declare statements live in document
// declarations, but are not entries in DocumentIR.Procedures and therefore
// must not become project-local call-graph targets in the compatibility view.
func isLegacyProjectProcedureKind(kind string) bool {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "declare", "declare_sub", "declare_function":
		return false
	default:
		return isProcedureSymbolKind(kind)
	}
}

// ProcedureOnlyResolver returns an immutable resolver view backed by the
// procedure-only symbol indexes collected while constructing the full project
// resolver. It preserves the pre-capability-planner call/effect semantics
// without constructing a second project symbol index.
func ProcedureOnlyResolver(resolver Resolver) Resolver {
	r, ok := resolver.(SymbolResolver)
	if !ok {
		return resolver
	}
	r.byName = r.procedureByName
	r.modules = r.procedureModules
	r.moduleKinds = r.procedureModuleKinds
	// The historical effects resolver was built from the complete procedure
	// list and deliberately did not inherit TypeDB/path completeness. Preserve
	// that fail-open boundary for procedure-local/effect consumers.
	r.complete = true
	return r
}

// NewResolver creates the default resolver backed by project symbols.
func NewResolver(symbols []ResolverSymbol) SymbolResolver {
	return NewSymbolResolver(symbols)
}

// NewResolverWithCompleteness is the canonical constructor for workspace
// snapshots whose project model may be partial. NewResolver remains complete
// for backwards compatibility with existing callers and tests.
func NewResolverWithCompleteness(symbols []ResolverSymbol, complete bool) SymbolResolver {
	return NewSymbolResolverWithCompleteness(symbols, complete)
}

// WithCompleteness returns an immutable resolver view over the same symbol
// snapshot. It is useful when an index transitions from pending to complete
// without rebuilding syntax IR.
func (r SymbolResolver) WithCompleteness(complete bool) SymbolResolver {
	r.complete = complete
	return r
}

func (r SymbolResolver) Complete() bool { return r.complete }

func (r SymbolResolver) ResolveCall(site CallSite) CallResolution {
	if strings.HasPrefix(strings.TrimSpace(site.Callee.Text), ".") {
		return CallResolution{Status: ResolutionMemberCall}
	}
	if site.Callee.Receiver != nil && strings.TrimSpace(*site.Callee.Receiver) == "." {
		return CallResolution{Status: ResolutionMemberCall}
	}
	base := cleanIdentifier(strings.TrimPrefix(site.Callee.BaseName, "New "))
	caller := callerModule(site.Caller)
	if caller == "" {
		caller = strings.TrimSpace(site.Module)
	}
	callerProcedure := callerProcedureName(site.Caller)
	if site.Callee.Receiver == nil && containsFolded(site.NonCallableNames, base) {
		return r.negativeCallResolution(CallResolution{Status: ResolutionNonCallable, ProjectLocal: true})
	}
	entries := r.visibleForCaller(r.byName[strings.ToLower(base)], caller, callerProcedure)
	if isDynamicCall(site) {
		return r.negativeCallResolution(CallResolution{Status: ResolutionDynamic})
	}
	procedures := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if isProcedureSymbolKind(entry.Kind) && (site.Callee.Receiver != nil || isReceiverlessProcedureCandidate(entry, caller)) {
			procedures = append(procedures, entry)
		}
	}
	if site.Callee.Receiver != nil {
		receiver := cleanQualifiedName(*site.Callee.Receiver)
		if isExternalLikeReceiver(receiver) {
			return CallResolution{Status: ResolutionExternal}
		}
		if r.knownMeMember(receiver, base, caller) {
			return CallResolution{Status: ResolutionExternal}
		}
		matches := receiverMatches(procedures, receiver, base)
		if len(matches) == 0 && strings.EqualFold(receiver, "me") &&
			strings.EqualFold(r.moduleKinds[strings.ToLower(cleanIdentifier(caller))], "class") {
			matches = receiverMatchesModule(procedures, caller, base)
		}
		if len(matches) == 0 {
			if receiverType := r.projectReceiverType(receiver, caller, callerProcedure); receiverType != "" {
				matches = receiverMatchesModule(procedures, receiverType, base)
			}
		}
		switch len(matches) {
		case 1:
			return CallResolution{Status: ResolutionMatched, Candidates: entriesToCandidates(matches)}
		case 0:
			if r.hostObjectReceiver(receiver) {
				return CallResolution{Status: ResolutionExternal}
			}
			// A receiver that names a project module (or Me in a project
			// object module) is a proof of project-local intent. Unknown
			// receivers remain member calls: they may be Object/Variant,
			// late-bound, or supplied by an unavailable TypeLib.
			if r.projectLocalReceiver(receiver, caller, callerProcedure) {
				// A known dynamic/object member on a project object (for example
				// Me.UrlSegments where UrlSegments is a Dictionary) remains
				// late-bound.  The receiver itself is project-local, but the
				// member's callable boundary is not statically known.
				if receiverHasDynamicMember(entries, receiver, base, caller) {
					return CallResolution{Status: ResolutionMemberCall}
				}
				if nonCallable := receiverNonCallable(entries, receiver, base); len(nonCallable) > 0 {
					return r.negativeCallResolution(CallResolution{
						Status: ResolutionNonCallable, Candidates: entriesToCandidates(nonCallable), ProjectLocal: true,
					}, nonCallable)
				}
				return r.negativeCallResolution(CallResolution{Status: ResolutionUnresolved, ProjectLocal: true})
			}
			if nonCallable := receiverNonCallable(entries, receiver, base); len(nonCallable) > 0 {
				return r.negativeCallResolution(CallResolution{
					Status: ResolutionNonCallable, Candidates: entriesToCandidates(nonCallable),
				}, nonCallable)
			}
			if builtinLikeNames[strings.ToLower(base)] {
				return CallResolution{Status: ResolutionBuiltinLike}
			}
			return CallResolution{Status: ResolutionMemberCall}
		default:
			return r.negativeCallResolution(CallResolution{Status: ResolutionAmbiguous, Candidates: entriesToCandidates(matches)}, matches)
		}
	}
	switch len(procedures) {
	case 1:
		return CallResolution{Status: ResolutionMatched, Candidates: entriesToCandidates(procedures)}
	default:
		if len(procedures) > 1 {
			return r.negativeCallResolution(CallResolution{Status: ResolutionAmbiguous, Candidates: entriesToCandidates(procedures)})
		}
	}
	if nonCallable := receiverlessNonCallable(entries, caller, callerProcedure); len(nonCallable) > 0 {
		return r.negativeCallResolution(CallResolution{
			Status: ResolutionNonCallable, Candidates: entriesToCandidates(nonCallable), ProjectLocal: true,
		}, nonCallable)
	}
	if inaccessible := inaccessibleProcedures(r.byName[strings.ToLower(base)], caller); len(inaccessible) > 0 {
		return r.negativeCallResolution(CallResolution{
			Status: ResolutionNonCallable, Candidates: entriesToCandidates(inaccessible), ProjectLocal: true,
		}, inaccessible)
	}
	textKey := strings.ToLower(strings.TrimPrefix(site.Callee.Text, "New "))
	if builtinLikeNames[textKey] || builtinLikeNames[strings.ToLower(base)] {
		return CallResolution{Status: ResolutionBuiltinLike}
	}
	return CallResolution{Status: ResolutionUnresolved}
}

func inaccessibleProcedures(entries []resolverEntry, caller string) []resolverEntry {
	if strings.TrimSpace(caller) == "" {
		return nil
	}
	result := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if !isProcedureSymbolKind(entry.Kind) || !symbolIsPrivate(entry) || strings.EqualFold(entry.module, caller) {
			continue
		}
		result = append(result, entry)
	}
	return result
}

func containsFolded(values []string, needle string) bool {
	for _, value := range values {
		if strings.EqualFold(cleanIdentifier(value), needle) {
			return true
		}
	}
	return false
}

func (r SymbolResolver) negativeCallResolution(result CallResolution, evidence ...[]resolverEntry) CallResolution {
	if !r.complete || hasUncertainEntries(evidence...) {
		if result.Status == ResolutionNonCallable || result.Status == ResolutionAmbiguous || result.Status == ResolutionUnresolved {
			result.Status = ResolutionIncomplete
		}
	}
	return result
}

func hasUncertainEntries(groups ...[]resolverEntry) bool {
	for _, entries := range groups {
		for _, entry := range entries {
			if entry.recovered || len(entry.conditionalBranches) > 0 {
				return true
			}
		}
	}
	return false
}

func receiverlessNonCallable(entries []resolverEntry, caller, procedure string) []resolverEntry {
	result := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if entry.isArray || dynamicValueEntry(entry) || strings.EqualFold(entry.Kind, "parameter") || isProcedureSymbolKind(entry.Kind) || strings.EqualFold(entry.Kind, "module") {
			continue
		}
		if isVariableSymbolKind(entry.Kind) && !isKnownScalarType(entry.typeName) {
			continue
		}
		if entry.parent != "" && !isEnumMemberKind(entry.Kind) && !strings.EqualFold(entry.parent, procedure) {
			continue
		}
		if isReceiverlessProcedureCandidate(entry, caller) || strings.EqualFold(entry.module, caller) {
			result = append(result, entry)
		}
	}
	return result
}

func receiverNonCallable(entries []resolverEntry, receiver, base string) []resolverEntry {
	qualified := strings.ToLower(cleanQualifiedName(receiver) + "." + base)
	shortQualified := strings.ToLower(cleanIdentifier(lastNamePart(receiver)) + "." + base)
	result := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if entry.isArray || dynamicValueEntry(entry) || strings.EqualFold(entry.Kind, "parameter") || isProcedureSymbolKind(entry.Kind) {
			continue
		}
		if isVariableSymbolKind(entry.Kind) && !isKnownScalarType(entry.typeName) {
			continue
		}
		name := strings.ToLower(entry.QualifiedName)
		if name == qualified || name == shortQualified {
			result = append(result, entry)
		}
	}
	return result
}

func receiverHasDynamicMember(entries []resolverEntry, receiver, base, caller string) bool {
	receiver = cleanQualifiedName(receiver)
	for _, entry := range entries {
		if !dynamicValueEntry(entry) {
			continue
		}
		if strings.EqualFold(lastNamePart(entry.QualifiedName), base) && strings.EqualFold(entry.module, caller) && strings.EqualFold(receiver, "me") {
			return true
		}
		qualified := strings.ToLower(entry.QualifiedName)
		if qualified == strings.ToLower(receiver+"."+base) ||
			qualified == strings.ToLower(lastNamePart(receiver)+"."+base) {
			return true
		}
	}
	return false
}

func isVariableSymbolKind(kind string) bool {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "variable", "local_variable", "module_variable", "field", "withevents_field":
		return true
	default:
		return false
	}
}

func dynamicValueEntry(entry resolverEntry) bool {
	switch strings.ToLower(strings.TrimSpace(entry.Kind)) {
	case "variable", "module_variable", "local_variable", "field", "withevents_field", "parameter":
		// These declarations can be Variant/Object or otherwise late-bound.
	default:
		// Constants, enum members, events, and other declarations have a
		// statically known callable boundary and must remain evidence for the
		// non-callable diagnostic.
		return false
	}
	typeName := strings.ToLower(strings.TrimSpace(entry.typeName))
	if typeName == "" || typeName == "variant" || typeName == "object" ||
		typeName == "collection" || strings.Contains(typeName, ".") {
		return true
	}
	switch typeName {
	case "byte", "integer", "long", "longlong", "longptr", "single", "double",
		"currency", "decimal", "date", "boolean", "string":
		return false
	default:
		// An object or user-defined Type whose declaration is not available
		// locally may expose a default member or a late-bound callable surface.
		// Keep the negative proof conservative for external TypeLib classes.
		return true
	}
}

func (r SymbolResolver) projectLocalReceiver(receiver, caller, procedure string) bool {
	receiver = cleanQualifiedName(receiver)
	if strings.EqualFold(receiver, "me") {
		kind := r.moduleKinds[strings.ToLower(cleanIdentifier(caller))]
		// Form/document modules expose host-controlled members through Me;
		// without the Excel/Access object model available, unresolved members
		// must remain external rather than becoming project-negative evidence.
		return kind == "class"
	}
	if receiver == "" {
		return false
	}
	// A chained receiver can cross an unavailable TypeLib or a late-bound
	// default member (for example frmMain.lstBodyNav.List). The first segment
	// alone does not prove that the final member is project-local.
	if strings.Contains(receiver, ".") {
		return false
	}
	first := strings.ToLower(cleanIdentifier(strings.Split(receiver, ".")[0]))
	// A function result can itself be an object (`Specs.It(...)`). Do not
	// mistake the function's name for a standard-module qualifier when its
	// return type is unavailable or belongs to an external TypeLib.
	for _, entry := range r.visibleForCaller(r.byName[first], caller, procedure) {
		if !isProcedureSymbolKind(entry.Kind) {
			continue
		}
		if strings.TrimSpace(entry.typeName) == "" {
			return false
		}
		if receiverType := r.projectReceiverType(receiver, caller, procedure); receiverType != "" {
			return true
		}
		return false
	}
	if r.isProjectModule(first) {
		return true
	}
	return r.projectReceiverType(receiver, caller, procedure) != ""
}

func (r SymbolResolver) projectReceiverType(receiver, caller, procedure string) string {
	receiver = cleanQualifiedName(receiver)
	if receiver == "" {
		return ""
	}
	first := strings.ToLower(cleanIdentifier(strings.Split(receiver, ".")[0]))
	projectTypes := map[string]struct{}{}
	for _, entry := range r.visibleForCaller(r.byName[first], caller, procedure) {
		typeName := cleanQualifiedName(strings.TrimSpace(entry.typeName))
		if typeName == "" {
			// Multiple declarations of the same receiver name are common in
			// conditional branches.  Without a type for one branch, the call
			// target cannot be proven project-local.
			return ""
		}
		if knownExternalTypeName(typeName) {
			// A competing external/TypeLib declaration keeps the receiver
			// late-bound from the diagnostic policy's point of view.
			return ""
		}
		if r.isProjectModule(strings.ToLower(lastNamePart(typeName))) {
			projectTypes[strings.ToLower(lastNamePart(typeName))] = struct{}{}
			continue
		}
		// An unavailable user-defined type is not evidence of a project
		// class; it may come from a referenced TypeLib.
		return ""
	}
	if len(projectTypes) != 1 {
		return ""
	}
	for typeName := range projectTypes {
		return typeName
	}
	return ""
}

func (r SymbolResolver) isProjectModule(module string) bool {
	key := strings.ToLower(cleanIdentifier(lastNamePart(module)))
	if _, ok := r.modules[key]; !ok {
		return false
	}
	// TypeLib constants are indexed with their library as a module qualifier
	// so Enum lookup can retain every candidate. That qualifier is external
	// evidence, not proof that calls such as VBA.IsObject target this project.
	return !strings.EqualFold(strings.TrimSpace(r.moduleKinds[key]), "external")
}

func knownExternalTypeName(typeName string) bool {
	lower := strings.ToLower(lastNamePart(typeName))
	for _, prefix := range []string{"web", "spec", "dictionary", "collection", "filesystem", "adodb", "mshtml", "office", "msforms", "excel", "scripting", "xml"} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}

func receiverMatchesModule(entries []resolverEntry, module, base string) []resolverEntry {
	module = strings.ToLower(cleanIdentifier(lastNamePart(module)))
	base = strings.ToLower(cleanIdentifier(base))
	out := make([]resolverEntry, 0)
	for _, entry := range entries {
		if strings.ToLower(cleanIdentifier(lastNamePart(entry.module))) == module &&
			strings.EqualFold(lastNamePart(entry.QualifiedName), base) {
			out = append(out, entry)
		}
	}
	return out
}

func (r SymbolResolver) knownMeMember(receiver, base, caller string) bool {
	if !strings.EqualFold(cleanQualifiedName(receiver), "me") {
		return false
	}
	kind := strings.ToLower(strings.TrimSpace(r.moduleKinds[strings.ToLower(cleanIdentifier(caller))]))
	if kind != "form" && kind != "document" {
		return false
	}
	switch strings.ToLower(cleanIdentifier(base)) {
	case "show", "hide", "repaint", "refresh", "range", "cells", "rows", "columns",
		"usedrange", "intersect", "union", "activate", "select", "calculate",
		"protect", "unprotect", "visible", "name", "worksheets", "sheets",
		"save", "saveas", "close":
		return true
	default:
		return false
	}
}

func (r SymbolResolver) hostObjectReceiver(receiver string) bool {
	key := strings.ToLower(cleanIdentifier(lastNamePart(receiver)))
	kind := strings.ToLower(strings.TrimSpace(r.moduleKinds[key]))
	return kind == "form" || kind == "document"
}

func callerProcedureName(caller ProcedureRef) string {
	if strings.TrimSpace(caller.Name) != "" {
		return strings.TrimSpace(caller.Name)
	}
	return lastNamePart(caller.QualifiedName)
}

func isDynamicCall(site CallSite) bool {
	member := strings.ToLower(cleanIdentifier(site.Callee.Member))
	if member == "" {
		member = strings.ToLower(cleanIdentifier(site.Callee.BaseName))
	}
	if member == "callbyname" {
		return true
	}
	if site.Callee.Receiver == nil {
		return false
	}
	receiver := cleanQualifiedName(*site.Callee.Receiver)
	if !strings.EqualFold(lastNamePart(receiver), "application") {
		return false
	}
	switch member {
	case "run", "ontime", "onkey":
		return true
	default:
		return false
	}
}

func isReceiverlessProcedureCandidate(entry resolverEntry, callerModule string) bool {
	if strings.EqualFold(entry.module, callerModule) {
		return true
	}
	return entry.moduleKind == "" || strings.EqualFold(entry.moduleKind, "standard")
}

func (r SymbolResolver) ResolveSymbol(ref SymbolReference) SymbolResolution {
	caller := callerModule(ref.Caller)
	if caller == "" {
		caller = strings.TrimSpace(ref.Module)
	}
	entries := r.visibleForCaller(r.byName[strings.ToLower(cleanIdentifier(ref.Name))], caller, callerProcedureName(ref.Caller))
	if len(entries) > 0 {
		allEnumMembers := true
		for _, entry := range entries {
			if !isEnumMemberKind(entry.Kind) {
				allEnumMembers = false
				break
			}
		}
		if allEnumMembers {
			enumResult := r.ResolveEnumMember(EnumMemberReference{Name: ref.Name, Module: caller, Caller: ref.Caller, Range: ref.Range})
			scope := ScopeProject
			if enumResult.Status == ResolutionUnresolved || enumResult.Status == ResolutionIncomplete {
				scope = ScopeUnresolved
			}
			return SymbolResolution{Status: enumResult.Status, Scope: scope, Candidates: append([]Candidate(nil), enumResult.Candidates...)}
		}
	}
	var candidates []Candidate
	for _, entry := range entries {
		if isProcedureSymbolKind(entry.Kind) || strings.EqualFold(entry.Kind, "module") {
			continue
		}
		candidates = append(candidates, entry.Candidate)
	}
	if len(candidates) == 0 {
		result := SymbolResolution{Scope: ScopeUnresolved, Status: ResolutionUnresolved}
		if !r.complete {
			result.Status = ResolutionIncomplete
		}
		return result
	}
	status := ResolutionMatched
	if len(candidates) > 1 {
		status = ResolutionAmbiguous
	}
	return SymbolResolution{Scope: ScopeProject, Status: status, Candidates: candidates}
}

// ResolveEnumMember resolves a bare or qualified enum constant while
// preserving all candidates. Project enum declarations take precedence over
// TypeLib metadata because only project declarations provide lexical evidence
// that a bare name is ambiguous. External enum metadata can associate one
// globally exposed constant with multiple enum groups, so external-only
// multiplicity is reported as an external resolution rather than VB053.
func (r SymbolResolver) ResolveEnumMember(ref EnumMemberReference) EnumResolution {
	caller := callerModule(ref.Caller)
	if caller == "" {
		caller = strings.TrimSpace(ref.Module)
	}
	entries := r.visibleForCaller(r.byName[strings.ToLower(cleanIdentifier(ref.Name))], caller, callerProcedureName(ref.Caller))
	filtered := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if !isEnumMemberKind(entry.Kind) || strings.TrimSpace(entry.parent) == "" {
			continue
		}
		if strings.TrimSpace(ref.Enum) != "" {
			enum := cleanQualifiedName(ref.Enum)
			parent := cleanQualifiedName(entry.parent)
			qualifiedParent := cleanQualifiedName(entry.module + "." + entry.parent)
			if !strings.EqualFold(enum, parent) && !strings.EqualFold(enum, qualifiedParent) &&
				!strings.EqualFold(enum, lastNamePart(parent)) {
				continue
			}
		}
		filtered = append(filtered, entry)
	}
	// An unqualified member declared in the caller's module wins over public
	// members imported from other modules. If multiple same-module enums expose
	// that member, no lexical winner exists and the result remains ambiguous.
	if strings.TrimSpace(ref.Enum) == "" && caller != "" {
		local := filtered[:0]
		for _, entry := range filtered {
			if strings.EqualFold(entry.module, caller) {
				local = append(local, entry)
			}
		}
		if len(local) > 0 {
			filtered = local
		}
	}
	// Referenced-library constants are a fallback after project symbols. A
	// TypeLib may describe the same globally usable name under several enum
	// groups (or several generated library records); those records do not prove
	// a source-level ambiguity. Keep project candidates when present and leave
	// external-only collisions as an external, fail-open resolution.
	if project := nonExternalEnumMembers(filtered); len(project) > 0 {
		filtered = project
	}
	sort.SliceStable(filtered, func(i, j int) bool {
		return resolverCandidateLess(filtered[i].Candidate, filtered[j].Candidate)
	})
	result := EnumResolution{}
	switch len(filtered) {
	case 0:
		result.Status = ResolutionUnresolved
	case 1:
		result.Status = ResolutionMatched
		result.Candidates = entriesToCandidates(filtered)
	default:
		if allExternalEnumMembers(filtered) {
			result.Status = ResolutionExternal
		} else {
			result.Status = ResolutionAmbiguous
		}
		result.Candidates = entriesToCandidates(filtered)
	}
	if (!r.complete || hasUncertainEntries(filtered)) && result.Status != ResolutionMatched {
		result.Status = ResolutionIncomplete
	}
	return result
}

func nonExternalEnumMembers(entries []resolverEntry) []resolverEntry {
	result := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if !isExternalEnumMember(entry) {
			result = append(result, entry)
		}
	}
	return result
}

func allExternalEnumMembers(entries []resolverEntry) bool {
	if len(entries) == 0 {
		return false
	}
	for _, entry := range entries {
		if !isExternalEnumMember(entry) {
			return false
		}
	}
	return true
}

func isExternalEnumMember(entry resolverEntry) bool {
	return isEnumMemberKind(entry.Kind) && strings.EqualFold(strings.TrimSpace(entry.moduleKind), "external")
}

func resolverCandidateLess(a, b Candidate) bool {
	if a.QualifiedName != b.QualifiedName {
		return a.QualifiedName < b.QualifiedName
	}
	if a.Kind != b.Kind {
		return a.Kind < b.Kind
	}
	if a.File != b.File {
		return a.File < b.File
	}
	return a.Line < b.Line
}

// ResolveEvent resolves a RaiseEvent target in the declaring object module.
// Events from another module are deliberately excluded: VBA's RaiseEvent
// binding is not a general project procedure lookup.
func (r SymbolResolver) ResolveEvent(ref SymbolReference) SymbolResolution {
	module := strings.TrimSpace(ref.Module)
	if module == "" {
		module = callerModule(ref.Caller)
	}
	entries := r.visibleForCaller(r.byName[strings.ToLower(cleanIdentifier(ref.Name))], module, callerProcedureName(ref.Caller))
	candidates := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if !isEventSymbolKind(entry.Kind) || !strings.EqualFold(entry.module, module) {
			continue
		}
		candidates = append(candidates, entry)
	}
	result := SymbolResolution{Scope: ScopeProject}
	switch len(candidates) {
	case 0:
		result.Scope = ScopeUnresolved
		result.Status = ResolutionUnresolved
	case 1:
		result.Status = ResolutionMatched
		result.Candidates = entriesToCandidates(candidates)
	default:
		result.Status = ResolutionAmbiguous
		result.Candidates = entriesToCandidates(candidates)
	}
	if (!r.complete || hasUncertainEntries(candidates)) && result.Status != ResolutionMatched {
		result.Status = ResolutionIncomplete
	}
	return result
}

func isEventSymbolKind(kind string) bool {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "event", "event_declaration":
		return true
	default:
		return false
	}
}

func (r SymbolResolver) visibleForCaller(entries []resolverEntry, callerModule, callerProcedure string) []resolverEntry {
	out := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if symbolIsPrivate(entry) && !strings.EqualFold(entry.module, callerModule) {
			continue
		}
		if entry.parent != "" && !isEnumMemberKind(entry.Kind) && callerProcedure != "" && !strings.EqualFold(entry.parent, callerProcedure) {
			continue
		}
		out = append(out, entry)
	}
	return out
}

func symbolIsPrivate(entry resolverEntry) bool {
	if strings.EqualFold(entry.visibility, "private") {
		return true
	}
	if strings.TrimSpace(entry.visibility) != "" {
		return false
	}
	switch strings.ToLower(entry.Kind) {
	case "variable", "module_variable", "field", "withevents_field", "const", "constant":
		return true
	default:
		return false
	}
}

func callerModule(caller ProcedureRef) string {
	if index := strings.LastIndex(strings.TrimSpace(caller.QualifiedName), "."); index > 0 {
		return caller.QualifiedName[:index]
	}
	return ""
}

func receiverMatches(entries []resolverEntry, receiver, base string) []resolverEntry {
	qualified := strings.ToLower(receiver + "." + base)
	shortQualified := strings.ToLower(cleanIdentifier(lastNamePart(receiver)) + "." + base)
	out := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		name := strings.ToLower(entry.QualifiedName)
		if name == qualified || name == shortQualified {
			out = append(out, entry)
		}
	}
	return out
}

func entriesToCandidates(entries []resolverEntry) []Candidate {
	out := make([]Candidate, len(entries))
	for i := range entries {
		out[i] = entries[i].Candidate
	}
	return out
}

func isProcedureSymbolKind(kind string) bool {
	switch strings.ToLower(kind) {
	case "sub", "function", "property", "property_get", "property_let", "property_set",
		"declare", "declare_sub", "declare_function":
		return true
	default:
		return false
	}
}

func isEnumMemberKind(kind string) bool {
	switch strings.ToLower(kind) {
	case "enum_member", "enum_constant", "enumvalue":
		return true
	default:
		return false
	}
}

// IsConstKind reports whether a resolver symbol denotes an immutable constant
// value. Keep the accepted symbol kinds centralized so lint, LSP, and project
// resolution cannot drift in their VB060 handling.
func IsConstKind(kind string) bool {
	return strings.EqualFold(kind, "const") || strings.EqualFold(kind, "enum_member")
}

func cleanQualifiedName(text string) string {
	parts := strings.FieldsFunc(strings.TrimSpace(text), func(r rune) bool { return r == '.' || r == '!' })
	cleaned := make([]string, 0, len(parts))
	for _, part := range parts {
		if part = cleanIdentifier(part); part != "" {
			cleaned = append(cleaned, part)
		}
	}
	return strings.Join(cleaned, ".")
}

func isExternalLikeReceiver(receiver string) bool {
	for _, part := range strings.Split(receiver, ".") {
		switch strings.ToLower(cleanIdentifier(part)) {
		case "application", "debug", "excel", "worksheetfunction", "thisworkbook", "activesheet", "activeworkbook", "selection":
			return true
		}
	}
	return false
}

func normalizeCandidateFile(file string) string {
	if strings.TrimSpace(file) == "" {
		return ""
	}
	return filepath.ToSlash(filepath.Clean(file))
}

var builtinLikeNames = map[string]bool{
	"array": true, "asc": true, "cbool": true, "cbyte": true, "ccur": true,
	"cdate": true, "cdbl": true, "cdec": true, "choose": true, "chr": true,
	"cint": true, "clng": true, "clnglng": true, "clngptr": true, "cos": true,
	"createobject": true, "cstr": true, "date": true, "dateadd": true,
	"debug.print": true, "dir": true, "doevents": true, "environ": true, "error": true,
	"format": true, "getobject": true, "inputbox": true, "instr": true,
	"isarray": true, "isdate": true, "isempty": true, "iserror": true,
	"isnull": true, "isnumeric": true, "join": true, "lbound": true,
	"lcase": true, "left": true, "len": true, "mid": true, "msgbox": true,
	"replace": true, "right": true, "rnd": true, "shell": true, "split": true, "str": true,
	"trim": true, "typename": true, "ubound": true, "ucase": true, "val": true,
	// Office dialog collection members are late-bound through With receivers
	// in common projects and are not project-local procedure names.
	"selecteditems": true,
}
