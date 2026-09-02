package analyze

import (
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func arrayPrivateProcedureTargets(files []parsedFile) map[string]sourceProcedure {
	targets := map[string]sourceProcedure{}
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		facts := file.ModuleFacts
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			visibility := strings.TrimSpace(proc.Visibility)
			private := strings.EqualFold(visibility, "Private") || strings.EqualFold(visibility, "Friend")
			modulePrivate := strings.EqualFold(visibility, "Public") && facts.privateModulePresent()
			if !private && !modulePrivate {
				continue
			}
			targets[arrayProcedureKey(proc)] = proc
		}
	}
	return targets
}

// buildArrayParticipantGraph classifies procedures and resolves all
// project-local call edges once. The two participant boundaries (the local
// fail-open plan and the narrower fixed-point plan) share this immutable graph
// so a revision does not scan every procedure and call site twice.
func buildArrayParticipantGraph(files []parsedFile, ctx analysisContext) *arrayParticipantGraph {
	all := make(map[string]sourceProcedure)
	fileByKey := make(map[string]parsedFile)
	byModule := make(map[string][]string)
	keyByIdentity := make(map[string]string)
	type entry struct {
		file parsedFile
		proc sourceProcedure
		base string
	}
	entries := make([]entry, 0)
	baseCounts := make(map[string]int)
	for _, file := range files {
		procedures := file.procedureView()
		for index := 0; index < procedures.Len(); index++ {
			proc := procedures.valueAt(index)
			base := arrayProcedureKey(proc)
			if base == "" {
				continue
			}
			entries = append(entries, entry{file: file, proc: proc, base: base})
			baseCounts[base]++
		}
	}
	for _, item := range entries {
		key := item.base
		if baseCounts[item.base] > 1 {
			key = arrayParticipantDisambiguatedKey(item.proc)
		}
		if _, exists := all[key]; exists {
			// Synthetic focused projections may omit Document/Index and still
			// produce identical identity fields. Keep their source-order ordinal
			// as a deterministic final discriminator.
			key = key + "|" + strconv.Itoa(len(all))
		}
		all[key] = item.proc
		fileByKey[key] = item.file
		identity := arrayParticipantProcedureIdentity(item.proc)
		if _, exists := keyByIdentity[identity]; exists {
			identity = identity + "\x00" + strconv.Itoa(len(all))
		}
		keyByIdentity[identity] = key
		module := strings.ToLower(strings.TrimSpace(item.proc.Module))
		byModule[module] = append(byModule[module], key)
	}
	if len(all) == 0 {
		return &arrayParticipantGraph{
			all:              all,
			fileByKey:        fileByKey,
			byModule:         byModule,
			keyByIdentity:    keyByIdentity,
			candidateIndex:   buildArrayCandidateIndex(all),
			adjacency:        map[string]map[string]bool{},
			reverse:          map[string]map[string]bool{},
			resolvedReverse:  map[string]map[string]bool{},
			callAdjacency:    map[string]map[string]bool{},
			knownSeeds:       map[string]bool{},
			intrinsicSeeds:   map[string]bool{},
			uncertainFacts:   map[bool]map[string]bool{false: {}, true: {}},
			uncertainCalls:   map[string]bool{},
			moduleArrayUsers: map[string][]string{},
		}
	}
	candidateIndex := buildArrayCandidateIndex(all)

	adjacency := make(map[string]map[string]bool, len(all))
	knownSeeds := make(map[string]bool, len(all))
	intrinsicSeeds := make(map[string]bool, len(all))
	uncertainFacts := map[bool]map[string]bool{false: {}, true: {}}
	uncertainCalls := make(map[string]bool)
	moduleArrayUsers := make(map[string][]string)
	callAdjacency := make(map[string]map[string]bool, len(all))
	type resolvedEdge struct {
		caller string
		target string
	}
	resolvedEdges := make([]resolvedEdge, 0)
	for key, proc := range all {
		file := fileByKey[key]
		moduleDecls := moduleDeclarationsForProcedure(files, proc)
		arraySeed := procedureArraySeed(proc)
		moduleArrayUse := procedureUsesModuleArray(file, proc, moduleDecls)
		shapeSeed := procedureHasArrayParameter(proc) || procedureReturnsArray(proc) || moduleArrayUse
		arraySeed = arraySeed || procedureHasArrayForEach(proc) || procedureHasObjectComparison(proc)
		if arraySeed || shapeSeed {
			knownSeeds[key] = true
			intrinsicSeeds[key] = true
		}
		if moduleArrayUse {
			module := strings.ToLower(strings.TrimSpace(proc.Module))
			moduleArrayUsers[module] = append(moduleArrayUsers[module], key)
		}
		for _, ignoreFeatureUnknown := range []bool{false, true} {
			if arrayParticipantFactsUncertain(proc, ignoreFeatureUnknown) {
				// Recovered statements, missing CFGs, and conditional IR facts are
				// bounded uncertainty. Keep the known module-array cluster fail-open
				// until resolution can prove a smaller dependency boundary.
				uncertainFacts[ignoreFeatureUnknown][key] = true
			}
		}
		for call := range proc.Calls.All() {
			resolution := arrayCallResolution(ctx, call)
			addEdge := func(target string) {
				if target == "" || target == key {
					if target == key {
						adjacency[key] = ensureArrayKeySet(adjacency[key])
						adjacency[key][target] = true
					}
					return
				}
				adjacency[key] = ensureArrayKeySet(adjacency[key])
				adjacency[key][target] = true
			}
			addCallEdge := func(target string) {
				if target == "" {
					return
				}
				callAdjacency[key] = ensureArrayKeySet(callAdjacency[key])
				callAdjacency[key][target] = true
			}
			if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
				candidate := resolution.Candidates[0]
				target := arrayCandidateKey(candidate, all, candidateIndex)
				addCallEdge(target)
				// Defer resolved-edge filtering until every procedure's
				// intrinsic seed has been classified. The source map is not
				// ordered, so checking the target while this loop runs would
				// make participant membership depend on map iteration order.
				resolvedEdges = append(resolvedEdges, resolvedEdge{caller: key, target: target})
				continue
			}
			if resolution.Status == procedureir.ResolutionAmbiguous || resolution.Status == procedureir.ResolutionUnresolved || resolution.Status == procedureir.ResolutionDynamic || resolution.Status == procedureir.ResolutionIncomplete {
				for _, candidate := range resolution.Candidates {
					target := arrayCandidateKey(candidate, all, candidateIndex)
					addCallEdge(target)
					addEdge(target)
				}
				if len(resolution.Candidates) == 0 {
					// A candidate-bearing ambiguous/dynamic/unresolved call is
					// bounded by those project-local candidates. Only a boundary
					// with no target identity expands to its owning module.
					uncertainCalls[key] = true
				}
			}
		}
	}
	for _, edge := range resolvedEdges {
		if intrinsicSeeds[edge.target] {
			adjacency[edge.caller] = ensureArrayKeySet(adjacency[edge.caller])
			adjacency[edge.caller][edge.target] = true
		}
	}
	// Keep reverse links for every project-local resolved edge separately from
	// the direct semantic graph. A bounded caller extension below lets a chain
	// such as Top -> Wrapper -> ArrayWorker reach the seed transitively when
	// facts are complete; recovered/incomplete targets use the module-array
	// fallback below instead of opening an unbounded caller hub.
	reverse := make(map[string]map[string]bool, len(adjacency))
	resolvedReverse := make(map[string]map[string]bool, len(adjacency))
	for _, edge := range resolvedEdges {
		if edge.target == "" {
			continue
		}
		resolvedReverse[edge.target] = ensureArrayKeySet(resolvedReverse[edge.target])
		resolvedReverse[edge.target][edge.caller] = true
	}
	for caller, callees := range adjacency {
		for callee := range callees {
			reverse[callee] = ensureArrayKeySet(reverse[callee])
			reverse[callee][caller] = true
		}
	}
	return &arrayParticipantGraph{
		all:              all,
		fileByKey:        fileByKey,
		byModule:         byModule,
		keyByIdentity:    keyByIdentity,
		candidateIndex:   candidateIndex,
		adjacency:        adjacency,
		reverse:          reverse,
		resolvedReverse:  resolvedReverse,
		callAdjacency:    callAdjacency,
		knownSeeds:       knownSeeds,
		intrinsicSeeds:   intrinsicSeeds,
		uncertainFacts:   uncertainFacts,
		uncertainCalls:   uncertainCalls,
		moduleArrayUsers: moduleArrayUsers,
	}
}

func (graph *arrayParticipantGraph) participantSet(ignoreFeatureUnknown bool) map[string]bool {
	participants := make(map[string]bool, len(graph.all))
	for key := range graph.knownSeeds {
		participants[key] = true
	}
	for key := range graph.uncertainFacts[ignoreFeatureUnknown] {
		participants[key] = true
	}
	closeArrayParticipantClosure(participants, graph.adjacency, graph.reverse)
	addResolvedCallerBoundary(participants, graph.resolvedReverse, graph.all, graph.byModule, ignoreFeatureUnknown)
	closeArrayParticipantClosure(participants, graph.adjacency, graph.reverse)
	globalFallback := false
	for key := range graph.uncertainFacts[ignoreFeatureUnknown] {
		if !participants[key] {
			continue
		}
		module := strings.ToLower(strings.TrimSpace(graph.all[key].Module))
		if module == "" {
			globalFallback = true
			continue
		}
		// Recovered or incomplete facts are expanded to the smallest known
		// module-array cluster. This preserves conservative state propagation
		// without turning every procedure in a giant module into a candidate.
		for _, user := range graph.moduleArrayUsers[module] {
			participants[user] = true
		}
	}
	for key := range graph.uncertainCalls {
		if !participants[key] {
			continue
		}
		// An unknown-only procedure remains a local participant for fail-open
		// diagnostics, but it must not open an entire giant module before a
		// semantic array seed reaches the uncertainty boundary.
		if !graph.intrinsicSeeds[key] && !ignoreFeatureUnknown {
			continue
		}
		module := strings.ToLower(strings.TrimSpace(graph.all[key].Module))
		if module == "" {
			globalFallback = true
			continue
		}
		if users := graph.moduleArrayUsers[module]; len(users) > 0 {
			for _, user := range users {
				participants[user] = true
			}
			continue
		}
		for _, procedure := range graph.byModule[module] {
			participants[procedure] = true
		}
	}
	if globalFallback {
		for key := range graph.all {
			participants[key] = true
		}
	}
	closeArrayParticipantClosure(participants, graph.adjacency, graph.reverse)
	return participants
}

// buildArrayParticipantSet derives the bounded interprocedural closure used
// by the array capability. Module declarations are deliberately not seeds:
// only a procedure that observes an array locally, exposes an array-shaped
// parameter/return, or reaches an array through a resolved call participates.
func buildArrayParticipantSet(files []parsedFile, ctx analysisContext) map[string]bool {
	return buildArrayParticipantGraph(files, ctx).participantSet(ctx.arrayIgnoreFeatureUnknown)
}

func buildArrayParticipantSets(files []parsedFile, ctx analysisContext) (map[string]bool, map[string]bool, map[string]string) {
	graph := buildArrayParticipantGraph(files, ctx)
	participants := graph.participantSet(ctx.arrayIgnoreFeatureUnknown)
	return participants, buildArrayInterproceduralParticipantSetFromGraph(graph, participants), graph.keyByIdentity
}

// buildArrayInterproceduralParticipantSet keeps the local fail-open plan
// separate from the fixed-point scope. A complete procedure that has no
// semantic array seed but carries an unknown array capability must retain its
// local array kernel/projection; it must not, by itself, make every array
// summary and module-entry solver walk the surrounding module.
func buildArrayInterproceduralParticipantSet(files []parsedFile, ctx analysisContext, participants map[string]bool) map[string]bool {
	graph := buildArrayParticipantGraph(files, ctx)
	return buildArrayInterproceduralParticipantSetFromGraph(graph, participants)
}

func buildArrayInterproceduralParticipantSetFromGraph(graph *arrayParticipantGraph, participants map[string]bool) map[string]bool {
	if len(participants) == 0 {
		return map[string]bool{}
	}
	// Derive the fixed-point boundary with feature-unknown bits ignored from the
	// shared graph. This preserves the proven semantic closure used before an
	// unknown-only local participant was added, while the outer participant plan
	// still retains that procedure for local fail-open diagnostics.
	legacy := graph.participantSet(true)
	legacyResult := make(map[string]bool, len(legacy))
	for key := range legacy {
		if participants[key] {
			legacyResult[key] = true
		}
	}
	all := graph.all
	fileByKey := graph.fileByKey
	moduleSizes := make(map[string]int)
	for _, proc := range all {
		moduleSizes[strings.ToLower(strings.TrimSpace(proc.Module))]++
	}
	connected := make(map[string]bool)
	for key := range graph.callAdjacency {
		if !participants[key] {
			continue
		}
		for target := range graph.callAdjacency[key] {
			if !participants[target] {
				continue
			}
			if legacy[key] {
				connected[target] = true
			}
			if legacy[target] {
				connected[key] = true
			}
		}
	}
	for key := range connected {
		if legacyResult[key] {
			continue
		}
		proc := graph.all[key]
		if !procedureArrayFactsUncertain(proc) || !procedureHasCompleteArrayFacts(proc) || proc.Features.unknown&arrayParticipantUnknownFeatures == 0 {
			continue
		}
		module := strings.ToLower(strings.TrimSpace(proc.Module))
		if moduleSizes[module] > arrayResolvedCallerModuleLimit {
			if !procedureHasDirectModuleArrayOperation(fileByKey[key], proc, fileByKey[key].moduleDecls()) {
				continue
			}
		}
		legacyResult[key] = true
	}
	return legacyResult
}

const arrayResolvedCallerModuleLimit = 512

func addResolvedCallerBoundary(participants map[string]bool, reverse map[string]map[string]bool, all map[string]sourceProcedure, byModule map[string][]string, ignoreFeatureUnknown bool) {
	keys := make([]string, 0, len(participants))
	for key := range participants {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, target := range keys {
		procedure, ok := all[target]
		if !ok || arrayParticipantFactsUncertain(procedure, ignoreFeatureUnknown) || len(byModule[strings.ToLower(strings.TrimSpace(procedure.Module))]) > arrayResolvedCallerModuleLimit {
			continue
		}
		for caller := range reverse[target] {
			participants[caller] = true
		}
	}
}

const arrayParticipantUnknownFeatures = featureArray | featureRangeArray | featureObject

func procedureArraySeed(proc sourceProcedure) bool {
	if proc.Features.present&(featureArray|featureRangeArray) != 0 {
		return true
	}
	// Resolved scalar calls without array-shaped evidence remain excluded. Any
	// unresolved/external call that leaves an array capability unknown is a seed
	// and is bounded by the uncertainty policy below.
	return false
}

func procedureArrayFactsUncertain(proc sourceProcedure) bool {
	if proc.Features.unknown&arrayParticipantUnknownFeatures != 0 {
		return true
	}
	if proc.IR == nil {
		return proc.Features.unknown != 0
	}
	return !procedureHasCompleteArrayFacts(proc)
}

func arrayParticipantFactsUncertain(proc sourceProcedure, ignoreFeatureUnknown bool) bool {
	if !ignoreFeatureUnknown {
		return procedureArrayFactsUncertain(proc)
	}
	if proc.IR == nil {
		return proc.Features.unknown != 0
	}
	return !procedureHasCompleteArrayFacts(proc)
}

// procedureHasCompleteArrayFacts reports whether the procedure's structural
// IR facts are complete. Feature unknowns are intentionally checked by the
// caller so complete IR with an unknown array capability can remain a local
// fail-open participant without widening the interprocedural scope.
func procedureHasCompleteArrayFacts(proc sourceProcedure) bool {
	if proc.IR == nil {
		return false
	}
	if proc.Document == nil || proc.Document.Parse.HasError || proc.Document.Parse.HasMissing || proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 || proc.Graph == nil {
		return false
	}
	if len(proc.Graph.UnknownFlowSources) > 0 {
		return false
	}
	for declaration := range proc.Declarations.All() {
		if declaration.Recovered || len(declaration.ConditionalBranches) > 0 {
			return false
		}
	}
	for statement := range proc.Statements.All() {
		if statement.Recovered || statement.Kind == procedureir.StatementUnknown || statement.Kind == procedureir.StatementRecovered || len(statement.ConditionalBranches) > 0 {
			return false
		}
	}
	for expression := range proc.Expressions.All() {
		if expression.Recovered || expression.Kind == procedureir.ExpressionUnknown && !isKnownNonValueExpressionSyntax(expression.SyntaxKind) {
			return false
		}
	}
	return true
}

func closeArrayParticipantClosure(participants map[string]bool, adjacency, reverse map[string]map[string]bool) {
	keys := make([]string, 0, len(participants))
	for key := range participants {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	queue := append([]string(nil), keys...)
	seen := make(map[string]bool, len(keys))
	for _, key := range keys {
		seen[key] = true
	}
	for head := 0; head < len(queue); head++ {
		current := queue[head]
		neighbors := make([]string, 0, len(adjacency[current])+len(reverse[current]))
		for caller := range reverse[current] {
			if !seen[caller] {
				neighbors = append(neighbors, caller)
			}
		}
		for callee := range adjacency[current] {
			if !seen[callee] {
				neighbors = append(neighbors, callee)
			}
		}
		sort.Strings(neighbors)
		for _, neighbor := range neighbors {
			if seen[neighbor] {
				continue
			}
			seen[neighbor] = true
			participants[neighbor] = true
			queue = append(queue, neighbor)
		}
	}
}

func moduleDeclarationsForProcedure(files []parsedFile, proc sourceProcedure) map[string]sourceDeclaration {
	for _, file := range files {
		if strings.EqualFold(strings.TrimSpace(file.Module), strings.TrimSpace(proc.Module)) || strings.EqualFold(strings.TrimSpace(file.IR.ModuleName), strings.TrimSpace(proc.Module)) {
			return file.moduleDecls()
		}
	}
	return nil
}

func procedureHasArrayParameter(proc sourceProcedure) bool {
	for parameter := range proc.Params.All() {
		if parameterIsArray(parameter) {
			return true
		}
	}
	return false
}

func procedureReturnsArray(proc sourceProcedure) bool {
	return proc.ReturnValueShape == procedureir.ValueShapeFixedArray ||
		proc.ReturnValueShape == procedureir.ValueShapeDynamicArray ||
		strings.Contains(proc.ReturnType, "()")
}

func procedureUsesModuleArray(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) bool {
	if len(moduleDecls) == 0 {
		return false
	}
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeModule {
			continue
		}
		for name, declaration := range moduleDecls {
			if strings.EqualFold(strings.TrimSpace(name), strings.TrimSpace(access.Name)) && declaration.Array && !declaration.Parameter {
				return true
			}
		}
	}
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct {
					if declaration, ok := moduleDecls[strings.ToLower(cleanIdentifier(redim.name))]; ok && declaration.Array && !declaration.Parameter {
						return true
					}
				}
			}
		}
		if lhs, _, indexed, ok := arrayAssignment(text); ok && !indexed {
			if declaration, declared := moduleDecls[strings.ToLower(cleanIdentifier(lhs))]; declared && declaration.Array && !declaration.Parameter {
				return true
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
			if declaration, declared := moduleDecls[strings.ToLower(cleanIdentifier(match[1]))]; declared && declaration.Array && !declaration.Parameter {
				return true
			}
		}
	}
	if proc.StartLine < 1 || proc.EndLine < proc.StartLine || len(file.Lines) == 0 {
		return false
	}
	if facts := file.moduleAnalysisFacts(); facts != nil {
		for name, declaration := range moduleDecls {
			if !declaration.Array || declaration.Parameter {
				continue
			}
			used := false
			facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
				if operation.Line >= proc.StartLine && operation.Line <= proc.EndLine {
					used = true
				}
			})
			if used {
				return true
			}
		}
	}
	return false
}

func procedureHasDirectModuleArrayOperation(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) bool {
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		for name, declaration := range moduleDecls {
			if declaration.Array && !declaration.Parameter && moduleArrayIndexedIdentifier(text, name) {
				return true
			}
		}
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct {
					if declaration, ok := moduleDecls[strings.ToLower(cleanIdentifier(redim.name))]; ok && declaration.Array && !declaration.Parameter {
						return true
					}
				}
			}
		}
	}
	return false
}

func moduleArrayIndexedIdentifier(text, name string) bool {
	text = strings.ToLower(text)
	name = strings.ToLower(strings.TrimSpace(name))
	if name == "" {
		return false
	}
	for start := 0; start <= len(text)-len(name); {
		relative := strings.Index(text[start:], name)
		if relative < 0 {
			return false
		}
		index := start + relative
		end := index + len(name)
		if (index == 0 || !isIdentifierPart(text[index-1])) && (end == len(text) || !isIdentifierPart(text[end])) {
			for end < len(text) && (text[end] == ' ' || text[end] == '\t') {
				end++
			}
			if end < len(text) && text[end] == '(' {
				return true
			}
		}
		start = index + len(name)
	}
	return false
}

func procedureHasArrayForEach(proc sourceProcedure) bool {
	for statement := range proc.Statements.All() {
		if statement.Kind == procedureir.StatementForEach {
			return true
		}
	}
	return false
}

func procedureHasObjectComparison(proc sourceProcedure) bool {
	if proc.Features.present&featureObject == 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		text := strings.ToLower(strings.TrimSpace(statement.Text))
		if strings.Contains(text, "nothing") && strings.Contains(text, "=") && !strings.Contains(text, " is ") {
			return true
		}
	}
	return false
}

func ensureArrayKeySet(set map[string]bool) map[string]bool {
	if set == nil {
		return map[string]bool{}
	}
	return set
}

func buildArrayCandidateIndex(all map[string]sourceProcedure) arrayCandidateIndex {
	index := arrayCandidateIndex{
		byName:          make(map[string]string, len(all)),
		byQualified:     make(map[string]string, len(all)),
		byQualifiedKind: make(map[arrayCandidateQualifiedKindKey]string, len(all)),
		byLineAndKind:   make(map[arrayCandidateLineKey]string),
	}
	keys := make([]string, 0, len(all))
	for key := range all {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, key := range keys {
		proc := all[key]
		name := strings.ToLower(strings.TrimSpace(proc.Name))
		if name == "" {
			continue
		}
		if _, exists := index.byName[name]; !exists {
			index.byName[name] = key
		}
		qualified := strings.ToLower(strings.TrimSpace(proc.Module + "." + proc.Name))
		if qualified != "." {
			if existing, exists := index.byQualified[qualified]; !exists {
				index.byQualified[qualified] = key
			} else if existing != key {
				index.byQualified[qualified] = ""
			}
			qualifiedKind := arrayCandidateQualifiedKindKey{qualified: qualified, kind: strings.ToLower(string(proc.ProcedureKind))}
			if existing, exists := index.byQualifiedKind[qualifiedKind]; !exists {
				index.byQualifiedKind[qualifiedKind] = key
			} else if existing != key {
				index.byQualifiedKind[qualifiedKind] = ""
			}
		}
		if proc.StartLine > 0 {
			lineKey := arrayCandidateLineKey{line: proc.StartLine, kind: strings.ToLower(string(proc.ProcedureKind))}
			if _, exists := index.byLineAndKind[lineKey]; !exists {
				index.byLineAndKind[lineKey] = key
			}
		}
	}
	return index
}

func arrayCandidateKey(candidate procedureir.Candidate, all map[string]sourceProcedure, index arrayCandidateIndex) string {
	qualified := strings.ToLower(strings.TrimSpace(candidate.QualifiedName))
	if proc, ok := all[qualified]; ok {
		return arrayProcedureKey(proc)
	}
	qualifiedKind := arrayCandidateQualifiedKindKey{qualified: qualified, kind: strings.ToLower(strings.TrimSpace(candidate.Kind))}
	if key := index.byQualifiedKind[qualifiedKind]; key != "" {
		return key
	}
	if key := index.byQualified[qualified]; key != "" {
		return key
	}
	if key, ok := index.byName[qualified]; ok {
		return key
	}
	if candidate.Line > 0 {
		lineKey := arrayCandidateLineKey{line: candidate.Line, kind: strings.ToLower(candidate.Kind)}
		if key, ok := index.byLineAndKind[lineKey]; ok {
			return key
		}
	}
	return ""
}

func arrayProcedureIsParticipant(ctx analysisContext, proc sourceProcedure) bool {
	participants := ctx.arrayInterproceduralParticipants
	if participants == nil {
		participants = ctx.arrayParticipants
	}
	if participants == nil {
		return true
	}
	return participants[arrayParticipantLookupKey(proc, ctx.arrayParticipantKeys)]
}
