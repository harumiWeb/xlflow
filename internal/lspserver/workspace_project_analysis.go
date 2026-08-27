package lspserver

import (
	"crypto/sha256"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// projectProcedureFingerprint is assembled from the canonical catalog owned by
// document analysis. The LSP dependency index never marshals a complete
// ProcedureIR to compare revisions.
type projectProcedureFingerprint struct {
	source      [sha256.Size]byte
	signature   [sha256.Size]byte
	module      [sha256.Size]byte
	conditional [sha256.Size]byte
	fallback    [sha256.Size]byte
}

type projectProcedureState struct {
	file        string
	fingerprint projectProcedureFingerprint
	callees     map[string]struct{}
	resolution  string
	uncertain   bool
}

type projectFileState struct {
	version       string
	procedureKeys []string
	moduleKey     string
	module        [sha256.Size]byte
	conditional   [sha256.Size]byte
	catalogSafe   bool
}

type projectDependencyView struct {
	procedures map[string]projectProcedureState
	reverse    map[string]map[string]struct{}
	files      map[string]projectFileState
	lookup     projectProcedureLookup
	revision   uint64
}

const projectDependencyRootKey = "\x00project-dependency-root"

func projectModuleDependencyKey(file string) string {
	return "module\x00" + symbolFileKey(file)
}

// projectImpactPaths retains the deterministic two-snapshot helper used by
// compatibility callers. The workspace index uses updateProjectDependencyView
// below so that unchanged files are not rebuilt.
func projectImpactPaths(before, after intel.ProjectAnalysisSnapshot) []string {
	return projectImpactPathsWithPerformance(before, after, nil)
}

func projectImpactPathsWithPerformance(before, after intel.ProjectAnalysisSnapshot, performance *performanceRecorder) []string {
	return projectImpactPathsWithPerformanceClass(before, after, performance, "background")
}

func projectImpactPathsWithPerformanceClass(before, after intel.ProjectAnalysisSnapshot, performance *performanceRecorder, class string) []string {
	if !before.Complete || !after.Complete {
		return nil
	}
	oldView := buildProjectDependencyViewWithPerformanceClass(before, performance, class)
	newView := buildProjectDependencyViewWithPerformanceClass(after, performance, class)
	changed := make(map[string]bool)
	for key, old := range oldView.procedures {
		if current, ok := newView.procedures[key]; !ok || projectProcedureStateChanged(old, current) {
			changed[key] = true
		}
	}
	for key, current := range newView.procedures {
		if old, ok := oldView.procedures[key]; !ok || projectProcedureStateChanged(old, current) {
			changed[key] = true
		}
	}
	if projectDependencyViewHasUncertainty(oldView) || projectDependencyViewHasUncertainty(newView) {
		changed[projectDependencyRootKey] = true
	}
	return projectImpactPathsFromViews(oldView, newView, changed)
}

func projectDependencyViewHasUncertainty(view projectDependencyView) bool {
	for key, state := range view.procedures {
		if key != projectDependencyRootKey && state.uncertain {
			return true
		}
	}
	return false
}

func projectStatesHaveUncertainty(states map[string]map[string]projectProcedureState) bool {
	for _, fileStates := range states {
		for _, state := range fileStates {
			if state.uncertain {
				return true
			}
		}
	}
	return false
}

func projectImpactPathsFromViews(oldView, newView projectDependencyView, changed map[string]bool) []string {
	queue := make([]string, 0, len(changed))
	for key := range changed {
		queue = append(queue, key)
	}
	seen := make(map[string]bool, len(changed))
	files := make(map[string]string)
	for key := range changed {
		if state, ok := newView.procedures[key]; ok {
			if state.file != "" {
				files[symbolFileKey(state.file)] = state.file
			}
		} else if state, ok := oldView.procedures[key]; ok {
			if state.file != "" {
				files[symbolFileKey(state.file)] = state.file
			}
		}
	}
	for len(queue) > 0 {
		callee := queue[0]
		queue = queue[1:]
		if seen[callee] {
			continue
		}
		seen[callee] = true
		callers := make(map[string]struct{}, len(oldView.reverse[callee])+len(newView.reverse[callee]))
		for caller := range oldView.reverse[callee] {
			callers[caller] = struct{}{}
		}
		for caller := range newView.reverse[callee] {
			callers[caller] = struct{}{}
		}
		for caller := range callers {
			if state, ok := newView.procedures[caller]; ok {
				if state.file != "" {
					files[symbolFileKey(state.file)] = state.file
				}
			} else if state, ok := oldView.procedures[caller]; ok {
				if state.file != "" {
					files[symbolFileKey(state.file)] = state.file
				}
			}
			if !seen[caller] {
				queue = append(queue, caller)
			}
		}
	}
	out := make([]string, 0, len(files))
	for _, file := range files {
		out = append(out, file)
	}
	sort.Slice(out, func(i, j int) bool { return symbolFileKey(out[i]) < symbolFileKey(out[j]) })
	return out
}

func buildProjectDependencyViewWithPerformanceClass(snapshot intel.ProjectAnalysisSnapshot, performance *performanceRecorder, class string) projectDependencyView {
	view := newProjectDependencyView()
	lookup := newProjectProcedureLookup(snapshot)
	view.lookup = lookup
	for _, document := range snapshot.Documents {
		states, keys := buildProjectProcedureStates(document, lookup, nil, false, performance, class)
		fileKey := symbolFileKey(document.IR.Path)
		moduleKey := projectModuleDependencyKey(document.IR.Path)
		view.procedures[moduleKey] = projectModuleState(document)
		view.files[fileKey] = projectFileState{
			version: document.Version, procedureKeys: keys,
			moduleKey: moduleKey,
			module:    document.ProcedureCatalog.ModuleContextHash, conditional: document.ProcedureCatalog.ConditionalHash,
			catalogSafe: projectProcedureCatalogUsable(document),
		}
		for key, state := range states {
			installProjectProcedureState(&view, key, state, nil, performance, class)
		}
	}
	return view
}

func projectModuleState(document intel.ProjectAnalysisDocument) projectProcedureState {
	fingerprint := projectProcedureFingerprint{
		module:      document.ProcedureCatalog.ModuleContextHash,
		conditional: document.ProcedureCatalog.ConditionalHash,
	}
	if !projectProcedureCatalogUsable(document) && len(document.IR.Declarations) > 0 {
		hasher := sha256.New()
		writeFingerprintText(hasher, document.Version, document.IR.Path, document.IR.ModuleName, document.IR.ModuleKind)
		for _, declaration := range document.IR.Declarations {
			writeFingerprintText(hasher, declaration.Name, declaration.Type, declaration.Visibility, declaration.Kind, declaration.Parent,
				decimalString(declaration.Range.StartLine), decimalString(declaration.Range.StartColumn), decimalString(declaration.Range.EndLine), decimalString(declaration.Range.EndColumn))
		}
		copy(fingerprint.fallback[:], hasher.Sum(nil))
	}
	return projectProcedureState{file: document.IR.Path, fingerprint: fingerprint}
}

func newProjectDependencyView() projectDependencyView {
	view := projectDependencyView{
		procedures: make(map[string]projectProcedureState),
		reverse:    make(map[string]map[string]struct{}),
		files:      make(map[string]projectFileState),
		lookup:     projectProcedureLookup{byQualified: make(map[string][]projectProcedureLookupEntry)},
	}
	view.procedures[projectDependencyRootKey] = projectProcedureState{}
	return view
}

type projectProcedureLookupEntry struct {
	key  string
	file string
	line int
}

type projectProcedureLookup struct {
	byQualified map[string][]projectProcedureLookupEntry
}

func newProjectProcedureLookup(snapshot intel.ProjectAnalysisSnapshot) projectProcedureLookup {
	lookup := projectProcedureLookup{byQualified: make(map[string][]projectProcedureLookupEntry)}
	for _, document := range snapshot.Documents {
		moduleKey := projectModuleDependencyKey(document.IR.Path)
		catalog := document.ProcedureCatalog
		for index, procedure := range document.IR.Procedures {
			identity := projectProcedureIdentity(catalog, index, procedure)
			key := projectProcedureIdentityKey(document.IR.Path, identity)
			line := procedure.Symbol.DeclarationRange.StartLine
			if index < len(catalog.Entries) {
				line = catalog.Entries[index].Range.Start.Line + 1
			}
			qualified := projectQualifiedKey(procedure.Symbol.QualifiedName, string(procedure.Symbol.Kind))
			lookup.byQualified[qualified] = append(lookup.byQualified[qualified], projectProcedureLookupEntry{key: key, file: document.IR.Path, line: line})
		}
		for _, declaration := range document.IR.Declarations {
			qualifiedName := declaration.Name
			if document.IR.ModuleName != "" {
				qualifiedName = document.IR.ModuleName + "." + declaration.Name
			}
			entry := projectProcedureLookupEntry{key: moduleKey, file: document.IR.Path, line: declaration.Range.StartLine}
			lookup.byQualified[projectQualifiedKey(qualifiedName, declaration.Kind)] = append(lookup.byQualified[projectQualifiedKey(qualifiedName, declaration.Kind)], entry)
			if declaration.Parent != "" && document.IR.ModuleName != "" {
				qualifiedWithParent := document.IR.ModuleName + "." + declaration.Parent + "." + declaration.Name
				lookup.byQualified[projectQualifiedKey(qualifiedWithParent, declaration.Kind)] = append(lookup.byQualified[projectQualifiedKey(qualifiedWithParent, declaration.Kind)], entry)
			}
		}
	}
	return lookup
}

func buildProjectProcedureStates(document intel.ProjectAnalysisDocument, lookup projectProcedureLookup, oldStates map[string]projectProcedureState, reuseUnchanged bool, performance *performanceRecorder, class string) (map[string]projectProcedureState, []string) {
	states := make(map[string]projectProcedureState, len(document.IR.Procedures))
	keys := make([]string, 0, len(document.IR.Procedures))
	catalog := document.ProcedureCatalog
	catalogSafe := projectProcedureCatalogUsable(document)
	for index, sourceProcedure := range document.IR.Procedures {
		identity := projectProcedureIdentity(catalog, index, sourceProcedure)
		key := projectProcedureIdentityKey(document.IR.Path, identity)
		fingerprint := projectProcedureFingerprintFor(catalog, index, sourceProcedure)
		if reuseUnchanged && catalogSafe {
			if previous, ok := oldStates[key]; ok && previous.fingerprint == fingerprint {
				states[key] = previous
				keys = append(keys, key)
				performance.addCounter(performanceCounterProcedureFingerprintReuses, 1, "workspace/project", performanceStageDependencyUpdate, class, document.IR.Path)
				continue
			}
		}
		procedure := sourceProcedure
		if resolved, ok := document.Resolution.ResolvedProcedure(index); ok {
			procedure = resolved
		}
		callees, resolution, uncertain := projectProcedureDependencies(procedure, lookup)
		if uncertain {
			callees[projectDependencyRootKey] = struct{}{}
		}
		states[key] = projectProcedureState{file: document.IR.Path, fingerprint: fingerprint, callees: callees, resolution: resolution, uncertain: uncertain}
		keys = append(keys, key)
		performance.addCounter(performanceCounterProcedureFingerprintBuilds, 1, "workspace/project", performanceStageDependencyUpdate, class, document.IR.Path)
		performance.addCounter(performanceCounterProceduresRevisited, 1, "workspace/project", performanceStageDependencyUpdate, class, document.IR.Path)
	}
	return states, keys
}

func projectProcedureIdentity(catalog intel.ProcedureCatalog, index int, procedure procedureir.ProcedureIR) intel.ProcedureIdentity {
	if index < len(catalog.Entries) && projectCatalogEntryMatches(catalog.Entries[index], procedure) {
		return catalog.Entries[index].Identity
	}
	name := strings.ToLower(strings.TrimSpace(procedure.Symbol.QualifiedName))
	if name == "" {
		name = strings.ToLower(strings.TrimSpace(procedure.Symbol.Name))
	}
	return intel.ProcedureIdentity{CanonicalName: name, Kind: strings.ToLower(strings.TrimSpace(string(procedure.Symbol.Kind))), Ordinal: index}
}

func projectProcedureIdentityKey(file string, identity intel.ProcedureIdentity) string {
	return strings.Join([]string{symbolFileKey(file), strings.ToLower(strings.TrimSpace(identity.CanonicalName)), strings.ToLower(strings.TrimSpace(identity.Kind)), decimalString(identity.Ordinal)}, "\x00")
}

func projectProcedureFingerprintFor(catalog intel.ProcedureCatalog, index int, procedure procedureir.ProcedureIR) projectProcedureFingerprint {
	if index < len(catalog.Entries) && projectCatalogEntryMatches(catalog.Entries[index], procedure) {
		entry := catalog.Entries[index]
		return projectProcedureFingerprint{source: entry.SourceHash, signature: entry.SignatureHash, module: catalog.ModuleContextHash, conditional: catalog.ConditionalHash}
	}
	// Compatibility callers may construct snapshots by hand without the
	// document-owned catalog. The normal LSP path always takes the branch above.
	hasher := sha256.New()
	writeFingerprintText(hasher, procedure.Symbol.QualifiedName, string(procedure.Symbol.Kind), procedure.Symbol.Name)
	for _, statement := range procedure.Statements {
		writeFingerprintText(hasher, string(statement.Kind), statement.Text, decimalString(statement.ID), decimalString(statement.Range.StartLine), decimalString(statement.Range.EndLine))
	}
	var fallback [sha256.Size]byte
	copy(fallback[:], hasher.Sum(nil))
	return projectProcedureFingerprint{fallback: fallback}
}

// projectProcedureCatalogUsable is the guard for catalog-based reuse.  A
// catalog is only authoritative when analysis marked it reusable and supplied
// one entry for every IR procedure with a matching identity.  Otherwise the
// dependency index rebuilds the affected project boundary conservatively.
func projectProcedureCatalogUsable(document intel.ProjectAnalysisDocument) bool {
	catalog := document.ProcedureCatalog
	if !catalog.ReuseSafe || len(catalog.Entries) != len(document.IR.Procedures) {
		return false
	}
	var zeroHash [sha256.Size]byte
	if catalog.ModuleContextHash == zeroHash || catalog.ConditionalHash == zeroHash {
		return false
	}
	identityCounts := make(map[string]int, len(catalog.Entries))
	for _, entry := range catalog.Entries {
		key := strings.ToLower(strings.TrimSpace(entry.Identity.CanonicalName)) + "\x00" + strings.ToLower(strings.TrimSpace(entry.Identity.Kind))
		identityCounts[key]++
	}
	identityOrdinals := make(map[string]int, len(identityCounts))
	for index, procedure := range document.IR.Procedures {
		entry := catalog.Entries[index]
		if entry.SourceHash == zeroHash || entry.SignatureHash == zeroHash || !projectCatalogEntryMatches(entry, procedure) {
			return false
		}
		key := strings.ToLower(strings.TrimSpace(entry.Identity.CanonicalName)) + "\x00" + strings.ToLower(strings.TrimSpace(entry.Identity.Kind))
		if identityCounts[key] > 1 && entry.Identity.Ordinal != identityOrdinals[key] {
			return false
		}
		identityOrdinals[key]++
	}
	return true
}

func projectCatalogEntryMatches(entry intel.ProcedureCatalogEntry, procedure procedureir.ProcedureIR) bool {
	canonical := strings.ToLower(strings.TrimSpace(entry.Identity.CanonicalName))
	if canonical == "" || entry.Identity.Ordinal < 0 {
		return false
	}
	name := strings.ToLower(strings.TrimSpace(procedure.Symbol.Name))
	qualified := strings.ToLower(strings.TrimSpace(procedure.Symbol.QualifiedName))
	if name == "" && qualified != "" {
		name = qualified
		if dot := strings.LastIndexByte(name, '.'); dot >= 0 {
			name = name[dot+1:]
		}
	}
	if canonical != name && canonical != qualified && !strings.HasSuffix(qualified, "."+canonical) {
		return false
	}
	return strings.EqualFold(strings.TrimSpace(entry.Identity.Kind), strings.TrimSpace(string(procedure.Symbol.Kind)))
}

func projectProcedureDependencies(procedure procedureir.ProcedureIR, lookup projectProcedureLookup) (map[string]struct{}, string, bool) {
	callees := make(map[string]struct{})
	var resolution strings.Builder
	uncertain := procedure.Symbol.Recovered || len(procedure.Symbol.ConditionalBranches) != 0
	for _, declaration := range procedure.Declarations {
		if declaration.Recovered || len(declaration.ConditionalBranches) != 0 {
			uncertain = true
			break
		}
	}
	for _, call := range procedure.Calls {
		resolution.WriteString(strings.ToLower(string(call.Resolution.Status)))
		resolution.WriteByte(':')
		for _, candidate := range call.Resolution.Candidates {
			key, ok := lookupCandidate(lookup, candidate)
			if ok {
				callees[key] = struct{}{}
				resolution.WriteString(key)
			} else {
				resolution.WriteString(strings.ToLower(candidate.QualifiedName))
				uncertain = true
			}
			resolution.WriteByte(';')
		}
		resolution.WriteByte('|')
		switch call.Resolution.Status {
		case procedureir.ResolutionAmbiguous, procedureir.ResolutionDynamic, procedureir.ResolutionIncomplete, procedureir.ResolutionNotAttempted:
			uncertain = true
		case procedureir.ResolutionUnresolved:
			uncertain = uncertain || call.Resolution.ProjectLocal
		case "":
			uncertain = true
		}
	}
	for _, access := range procedure.Accesses {
		resolution.WriteString("access:")
		uncertain = appendProjectSymbolResolution(&resolution, access.Resolution, lookup, callees, uncertain)
	}
	for _, event := range procedure.RaiseEvents {
		resolution.WriteString("event:")
		uncertain = appendProjectSymbolResolution(&resolution, event.Resolution, lookup, callees, uncertain)
	}
	return callees, resolution.String(), uncertain
}

func appendProjectSymbolResolution(resolution *strings.Builder, fact procedureir.SymbolResolution, lookup projectProcedureLookup, callees map[string]struct{}, uncertain bool) bool {
	resolution.WriteString(strings.ToLower(string(fact.Status)))
	resolution.WriteByte(':')
	for _, candidate := range fact.Candidates {
		key, ok := lookupCandidate(lookup, candidate)
		if ok {
			callees[key] = struct{}{}
			resolution.WriteString(key)
		} else {
			resolution.WriteString(strings.ToLower(candidate.QualifiedName))
			uncertain = true
		}
		resolution.WriteByte(';')
	}
	resolution.WriteByte('|')
	switch fact.Status {
	case procedureir.ResolutionAmbiguous, procedureir.ResolutionDynamic, procedureir.ResolutionIncomplete, procedureir.ResolutionNotAttempted:
		uncertain = true
	case procedureir.ResolutionUnresolved:
		uncertain = uncertain || fact.Scope == procedureir.ScopeProject || fact.Scope == procedureir.ScopeUnresolved
	case "":
		uncertain = true
	}
	return uncertain
}

func lookupCandidate(lookup projectProcedureLookup, candidate procedureir.Candidate) (string, bool) {
	entries := lookup.byQualified[projectQualifiedKey(candidate.QualifiedName, candidate.Kind)]
	if len(entries) == 1 {
		return entries[0].key, true
	}
	if len(entries) == 0 {
		return "", false
	}
	for _, entry := range entries {
		if candidate.Line == entry.line && candidateFileMatches(candidate.File, entry.file) {
			return entry.key, true
		}
	}
	return "", false
}

func projectQualifiedKey(name, kind string) string {
	return strings.ToLower(strings.TrimSpace(name)) + "\x00" + strings.ToLower(strings.TrimSpace(kind))
}

func candidateFileMatches(candidate, actual string) bool {
	candidate = filepath.ToSlash(filepath.Clean(strings.TrimSpace(candidate)))
	actual = filepath.ToSlash(filepath.Clean(strings.TrimSpace(actual)))
	if candidate == actual {
		return true
	}
	return strings.HasSuffix(actual, "/"+candidate) || strings.HasSuffix(candidate, "/"+actual)
}

func projectProcedureStateChanged(old, current projectProcedureState) bool {
	if old.fingerprint != current.fingerprint || old.resolution != current.resolution || old.uncertain != current.uncertain {
		return true
	}
	return !sameProjectStringSet(old.callees, current.callees)
}

func sameProjectStringSet(left, right map[string]struct{}) bool {
	if len(left) != len(right) {
		return false
	}
	for key := range left {
		if _, ok := right[key]; !ok {
			return false
		}
	}
	return true
}

func installProjectProcedureState(view *projectDependencyView, key string, state projectProcedureState, old *projectProcedureState, performance *performanceRecorder, class string) {
	if old != nil {
		for callee := range old.callees {
			if _, keep := state.callees[callee]; keep {
				continue
			}
			if callers := view.reverse[callee]; callers != nil {
				if _, present := callers[key]; present {
					removeProjectReverse(view.reverse, callee, key)
					performance.addCounter(performanceCounterDependencyEdgesUpdated, 1, "workspace/project", performanceStageDependencyUpdate, class, state.file)
				}
			}
		}
	}
	view.procedures[key] = state
	for callee := range state.callees {
		if old != nil {
			if _, existed := old.callees[callee]; existed {
				// Keep an unchanged reverse edge in place. If the view was
				// externally damaged, restore the missing posting below.
				if callers := view.reverse[callee]; callers != nil {
					if _, present := callers[key]; present {
						continue
					}
				}
			}
		}
		if view.reverse[callee] == nil {
			view.reverse[callee] = make(map[string]struct{})
		}
		if _, exists := view.reverse[callee][key]; !exists {
			view.reverse[callee][key] = struct{}{}
			performance.addCounter(performanceCounterDependencyEdgesUpdated, 1, "workspace/project", performanceStageDependencyUpdate, class, state.file)
		}
	}
}

func removeProjectReverse(reverse map[string]map[string]struct{}, callee, caller string) {
	callers := reverse[callee]
	if callers == nil {
		return
	}
	delete(callers, caller)
	if len(callers) == 0 {
		delete(reverse, callee)
	}
}

func removeProjectProcedureState(view *projectDependencyView, key string, performance *performanceRecorder, class string) {
	state, exists := view.procedures[key]
	if !exists {
		return
	}
	for callee := range state.callees {
		if callers := view.reverse[callee]; callers != nil {
			if _, present := callers[key]; present {
				performance.addCounter(performanceCounterDependencyEdgesUpdated, 1, "workspace/project", performanceStageDependencyUpdate, class, state.file)
			}
		}
		removeProjectReverse(view.reverse, callee, key)
	}
	delete(view.procedures, key)
	performance.addCounter(performanceCounterDependencyNodesUpdated, 1, "workspace/project", performanceStageDependencyUpdate, class, state.file)
}

// updateProjectDependencyView updates only changed files unless a declaration
// change can alter project resolution. All states may be revisited in that
// case, but only changed nodes and edges are published to the graph.
func updateProjectDependencyView(view *projectDependencyView, snapshot intel.ProjectAnalysisSnapshot, performance *performanceRecorder, class string) []string {
	lookup := view.lookup
	currentDocuments := make(map[string]int, len(snapshot.Documents))
	currentCatalogSafe := make(map[string]bool, len(snapshot.Documents))
	for index, document := range snapshot.Documents {
		fileKey := symbolFileKey(document.IR.Path)
		currentDocuments[fileKey] = index
		currentCatalogSafe[fileKey] = projectProcedureCatalogUsable(document)
	}
	changedFiles := make(map[string]bool)
	resolutionRefresh := false
	for fileKey, oldFile := range view.files {
		index, exists := currentDocuments[fileKey]
		if !exists {
			changedFiles[fileKey] = true
			resolutionRefresh = true
			continue
		}
		document := snapshot.Documents[index]
		if projectFileChanged(view, document, fileKey, oldFile, currentCatalogSafe[fileKey]) {
			changedFiles[fileKey] = true
		}
		if !oldFile.catalogSafe || !currentCatalogSafe[fileKey] {
			resolutionRefresh = true
		}
		if oldFile.module != document.ProcedureCatalog.ModuleContextHash || oldFile.conditional != document.ProcedureCatalog.ConditionalHash || (oldFile.version != document.Version && len(document.IR.Procedures) == 0) {
			resolutionRefresh = true
		}
	}
	for fileKey, index := range currentDocuments {
		if _, exists := view.files[fileKey]; !exists {
			changedFiles[fileKey] = true
			resolutionRefresh = true
			continue
		}
		if projectFileChanged(view, snapshot.Documents[index], fileKey, view.files[fileKey], currentCatalogSafe[fileKey]) {
			changedFiles[fileKey] = true
		}
		if !view.files[fileKey].catalogSafe || !currentCatalogSafe[fileKey] {
			resolutionRefresh = true
		}
		if view.files[fileKey].module != snapshot.Documents[index].ProcedureCatalog.ModuleContextHash || view.files[fileKey].conditional != snapshot.Documents[index].ProcedureCatalog.ConditionalHash || (view.files[fileKey].version != snapshot.Documents[index].Version && len(snapshot.Documents[index].IR.Procedures) == 0) {
			resolutionRefresh = true
		}
	}
	for fileKey := range changedFiles {
		index, exists := currentDocuments[fileKey]
		if exists && documentNeedsResolutionRefresh(view, snapshot.Documents[index], fileKey, currentCatalogSafe[fileKey]) {
			resolutionRefresh = true
		}
	}
	if resolutionRefresh {
		lookup = newProjectProcedureLookup(snapshot)
		view.lookup = lookup
		for index := range snapshot.Documents {
			changedFiles[symbolFileKey(snapshot.Documents[index].IR.Path)] = true
		}
	}
	for fileKey, oldFile := range view.files {
		if changedFiles[fileKey] {
			continue
		}
		performance.addCounter(performanceCounterProcedureFingerprintReuses, uint64(len(oldFile.procedureKeys)), "workspace/project", performanceStageDependencyUpdate, class, "")
	}

	newStates := make(map[string]map[string]projectProcedureState, len(changedFiles))
	newKeys := make(map[string][]string, len(changedFiles))
	for fileKey := range changedFiles {
		index, exists := currentDocuments[fileKey]
		if !exists {
			continue
		}
		oldStates := make(map[string]projectProcedureState, len(view.files[fileKey].procedureKeys))
		for _, key := range view.files[fileKey].procedureKeys {
			if state, ok := view.procedures[key]; ok {
				oldStates[key] = state
			}
		}
		states, keys := buildProjectProcedureStates(snapshot.Documents[index], lookup, oldStates, !resolutionRefresh, performance, class)
		newStates[fileKey], newKeys[fileKey] = states, keys
	}

	changed := make(map[string]bool)
	oldReverse := make(map[string]map[string]struct{})
	for fileKey := range changedFiles {
		oldFile := view.files[fileKey]
		oldModuleKey := oldFile.moduleKey
		if oldModuleKey == "" {
			oldModuleKey = projectModuleDependencyKey(fileKey)
		}
		if oldModule, exists := view.procedures[oldModuleKey]; exists {
			if index, currentExists := currentDocuments[fileKey]; currentExists {
				currentDocument := snapshot.Documents[index]
				currentModuleKey := projectModuleDependencyKey(currentDocument.IR.Path)
				currentModule := projectModuleState(currentDocument)
				if oldModuleKey != currentModuleKey {
					changed[oldModuleKey] = true
					oldReverse[oldModuleKey] = cloneProjectCallers(view.reverse[oldModuleKey])
					changed[currentModuleKey] = true
					oldReverse[currentModuleKey] = cloneProjectCallers(view.reverse[currentModuleKey])
				} else if projectProcedureStateChanged(oldModule, currentModule) {
					changed[oldModuleKey] = true
					oldReverse[oldModuleKey] = cloneProjectCallers(view.reverse[oldModuleKey])
				}
			} else {
				changed[oldModuleKey] = true
				oldReverse[oldModuleKey] = cloneProjectCallers(view.reverse[oldModuleKey])
			}
		} else if index, currentExists := currentDocuments[fileKey]; currentExists {
			currentModuleKey := projectModuleDependencyKey(snapshot.Documents[index].IR.Path)
			changed[currentModuleKey] = true
			oldReverse[currentModuleKey] = cloneProjectCallers(view.reverse[currentModuleKey])
		}
		for _, key := range oldFile.procedureKeys {
			if state, ok := view.procedures[key]; ok {
				if current, stillPresent := newStates[fileKey][key]; !stillPresent || projectProcedureStateChanged(state, current) {
					changed[key] = true
					oldReverse[key] = cloneProjectCallers(view.reverse[key])
				}
			}
		}
		for key := range newStates[fileKey] {
			if _, exists := view.procedures[key]; !exists {
				changed[key] = true
				oldReverse[key] = cloneProjectCallers(view.reverse[key])
			}
		}
	}
	if projectDependencyViewHasUncertainty(*view) || projectStatesHaveUncertainty(newStates) {
		changed[projectDependencyRootKey] = true
		oldReverse[projectDependencyRootKey] = cloneProjectCallers(view.reverse[projectDependencyRootKey])
	}
	impactFiles := make(map[string]string)
	for key := range changed {
		if state, ok := view.procedures[key]; ok {
			if state.file != "" {
				impactFiles[symbolFileKey(state.file)] = state.file
			}
		}
	}
	for _, callers := range oldReverse {
		for caller := range callers {
			if state, ok := view.procedures[caller]; ok {
				if state.file != "" {
					impactFiles[symbolFileKey(state.file)] = state.file
				}
			}
		}
	}
	for fileKey := range changedFiles {
		oldFile := view.files[fileKey]
		oldModuleKey := oldFile.moduleKey
		if oldModuleKey == "" {
			oldModuleKey = projectModuleDependencyKey(fileKey)
		}
		for _, key := range oldFile.procedureKeys {
			if _, exists := newStates[fileKey][key]; exists {
				continue
			}
			removeProjectProcedureState(view, key, performance, class)
		}
		if _, exists := currentDocuments[fileKey]; !exists {
			removeProjectProcedureState(view, oldModuleKey, performance, class)
		}
		if index, exists := currentDocuments[fileKey]; exists {
			currentDocument := snapshot.Documents[index]
			currentModuleKey := projectModuleDependencyKey(currentDocument.IR.Path)
			if oldModuleKey != currentModuleKey {
				removeProjectProcedureState(view, oldModuleKey, performance, class)
			}
			moduleState := projectModuleState(currentDocument)
			oldModule, hadOldModule := view.procedures[currentModuleKey]
			if !hadOldModule || projectProcedureStateChanged(oldModule, moduleState) {
				if hadOldModule {
					installProjectProcedureState(view, currentModuleKey, moduleState, &oldModule, performance, class)
				} else {
					installProjectProcedureState(view, currentModuleKey, moduleState, nil, performance, class)
				}
				performance.addCounter(performanceCounterDependencyNodesUpdated, 1, "workspace/project", performanceStageDependencyUpdate, class, currentDocument.IR.Path)
			}
			for key, state := range newStates[fileKey] {
				old, hadOld := view.procedures[key]
				if !hadOld || projectProcedureStateChanged(old, state) {
					if hadOld {
						installProjectProcedureState(view, key, state, &old, performance, class)
					} else {
						installProjectProcedureState(view, key, state, nil, performance, class)
					}
					performance.addCounter(performanceCounterDependencyNodesUpdated, 1, "workspace/project", performanceStageDependencyUpdate, class, snapshot.Documents[index].IR.Path)
				}
			}
			view.files[fileKey] = projectFileState{
				version: currentDocument.Version, procedureKeys: newKeys[fileKey], moduleKey: currentModuleKey,
				module: currentDocument.ProcedureCatalog.ModuleContextHash, conditional: currentDocument.ProcedureCatalog.ConditionalHash,
				catalogSafe: currentCatalogSafe[fileKey],
			}
		} else {
			delete(view.files, fileKey)
		}
	}

	return projectImpactPathsFromIncrementalView(view, changed, oldReverse, impactFiles)
}

func projectFileChanged(view *projectDependencyView, document intel.ProjectAnalysisDocument, fileKey string, oldFile projectFileState, catalogSafe bool) bool {
	if !oldFile.catalogSafe || !catalogSafe {
		return true
	}
	if oldFile.version != document.Version || oldFile.module != document.ProcedureCatalog.ModuleContextHash || oldFile.conditional != document.ProcedureCatalog.ConditionalHash {
		return true
	}
	// Hand-built compatibility snapshots may not carry a publication version.
	// Compare their canonical procedure fields so a changed body is still
	// detected without falling back to whole-IR serialization.
	if document.Version == "" {
		return documentHasProcedureChanges(view, document, fileKey)
	}
	return false
}

func documentHasProcedureChanges(view *projectDependencyView, document intel.ProjectAnalysisDocument, fileKey string) bool {
	oldByKey := make(map[string]projectProcedureState, len(view.files[fileKey].procedureKeys))
	for _, key := range view.files[fileKey].procedureKeys {
		oldByKey[key] = view.procedures[key]
	}
	for index, procedure := range document.IR.Procedures {
		identity := projectProcedureIdentity(document.ProcedureCatalog, index, procedure)
		key := projectProcedureIdentityKey(document.IR.Path, identity)
		old, ok := oldByKey[key]
		if !ok || old.fingerprint != projectProcedureFingerprintFor(document.ProcedureCatalog, index, procedure) {
			return true
		}
		delete(oldByKey, key)
	}
	return len(oldByKey) != 0
}

func projectImpactPathsFromIncrementalView(view *projectDependencyView, changed map[string]bool, oldReverse map[string]map[string]struct{}, impactFiles map[string]string) []string {
	queue := make([]string, 0, len(changed))
	for key := range changed {
		queue = append(queue, key)
	}
	seen := make(map[string]bool, len(changed))
	files := make(map[string]string, len(impactFiles))
	for key, file := range impactFiles {
		files[key] = file
	}
	for key := range changed {
		if state, ok := view.procedures[key]; ok {
			if state.file != "" {
				files[symbolFileKey(state.file)] = state.file
			}
		}
	}
	for len(queue) > 0 {
		callee := queue[0]
		queue = queue[1:]
		if seen[callee] {
			continue
		}
		seen[callee] = true
		callers := make(map[string]struct{}, len(oldReverse[callee])+len(view.reverse[callee]))
		for caller := range oldReverse[callee] {
			callers[caller] = struct{}{}
		}
		for caller := range view.reverse[callee] {
			callers[caller] = struct{}{}
		}
		for caller := range callers {
			if state, ok := view.procedures[caller]; ok {
				if state.file != "" {
					files[symbolFileKey(state.file)] = state.file
				}
			}
			if !seen[caller] {
				queue = append(queue, caller)
			}
		}
	}
	out := make([]string, 0, len(files))
	for _, file := range files {
		out = append(out, file)
	}
	sort.Slice(out, func(i, j int) bool { return symbolFileKey(out[i]) < symbolFileKey(out[j]) })
	return out
}

func documentNeedsResolutionRefresh(view *projectDependencyView, document intel.ProjectAnalysisDocument, fileKey string, catalogSafe bool) bool {
	if !view.files[fileKey].catalogSafe || !catalogSafe {
		return true
	}
	oldKeys := view.files[fileKey].procedureKeys
	oldByIdentity := make(map[string]projectProcedureState, len(oldKeys))
	for _, key := range oldKeys {
		oldByIdentity[key] = view.procedures[key]
	}
	for index, procedure := range document.IR.Procedures {
		identity := projectProcedureIdentity(document.ProcedureCatalog, index, procedure)
		key := projectProcedureIdentityKey(document.IR.Path, identity)
		old, exists := oldByIdentity[key]
		current := projectProcedureFingerprintFor(document.ProcedureCatalog, index, procedure)
		if !exists || old.fingerprint.signature != current.signature || old.fingerprint.module != current.module || old.fingerprint.conditional != current.conditional {
			return true
		}
		delete(oldByIdentity, key)
	}
	return len(oldByIdentity) != 0
}

func cloneProjectCallers(callers map[string]struct{}) map[string]struct{} {
	clone := make(map[string]struct{}, len(callers))
	for caller := range callers {
		clone[caller] = struct{}{}
	}
	return clone
}

func writeFingerprintText(hasher interface{ Write([]byte) (int, error) }, values ...string) {
	for _, value := range values {
		_, _ = hasher.Write([]byte(value))
		_, _ = hasher.Write([]byte{'\x00'})
	}
}

func decimalString(value int) string {
	if value == 0 {
		return "0"
	}
	var digits [24]byte
	i := len(digits)
	for value > 0 {
		i--
		digits[i] = byte('0' + value%10)
		value /= 10
	}
	return string(digits[i:])
}

func cloneProcedureCatalog(catalog intel.ProcedureCatalog) intel.ProcedureCatalog {
	catalog.Entries = append([]intel.ProcedureCatalogEntry(nil), catalog.Entries...)
	return catalog
}
