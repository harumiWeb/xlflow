package analyze

import (
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type arrayByRefAllocationSummaries map[string]map[int]bool

// arrayByRefConditionalAllocations records a ByRef array output that is
// allocated only when a count-bearing input is positive. The outer key is the
// callee procedure; each entry maps the output array parameter index to the
// count-bearing input parameter index.
type arrayByRefConditionalAllocations map[string]map[int]int

// arrayByRefLengthAllocations records a ByRef array output whose paired
// ByRef scalar output is assigned a successful array length. A positive value
// of that scalar is therefore a conditional allocation proof for the array.
type arrayByRefLengthAllocations map[string]map[int]int

var (
	arrayByRefCountExitRe   = regexp.MustCompile(`(?i)^\s*if\s+([A-Za-z_]\w*)\s*\.\s*count\s*=\s*0\s+then\s+exit\s+(?:sub|function|property)\s*$`)
	arrayByRefCountRedimRe  = regexp.MustCompile(`(?i)^\s*redim\s+([A-Za-z_]\w*)\s*\(\s*0\s+to\s+([A-Za-z_]\w*)\s*\.\s*count\s*-\s*1\s*\)\s*$`)
	arrayByRefLengthFullRe  = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*-\s*lbound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
	arrayByRefLengthUpperRe = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*=\s*ubound\s*\(\s*([A-Za-z_]\w*)\s*\)\s*\+\s*1\s*$`)
)

type arrayModuleAllocationSummaries map[string]map[string]bool

// arrayModuleInvalidationSummaries records module arrays that may be
// unallocated or unknown when a project-local procedure returns normally.
// The summary starts from an allocated module-array state, so a fixed-size
// array and an Erase followed by a guaranteed ReDim remain allocated while a
// reachable conditional Erase is retained as an invalidation.
type arrayModuleInvalidationSummaries map[string]map[string]bool

func inferArrayModuleInvalidationSummaries(files []parsedFile, ctx analysisContext) arrayModuleInvalidationSummaries {
	summaries := arrayModuleInvalidationSummaries{}
	ctx.arrayModuleInvalidations = summaries
	ctx.arrayModuleInvalidationCacheWritable = true
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		moduleDecls := file.moduleDecls()
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			key := arrayProcedureKey(proc)
			if key == "" {
				continue
			}
			if _, cached := summaries[key]; cached {
				continue
			}
			summaries[key] = arrayPrivateModuleArrayInvalidationsWithVisiting(file, proc, moduleDecls, ctx, map[string]bool{})
		}
	}
	return summaries
}

type arrayProcedureDominators map[string]map[vbacfg.BlockID]bool

type arrayModuleConfigurationState struct {
	byProcedure       map[string]map[string]bool
	dataTable         map[string]bool
	genericCollection map[string]bool
}

// arrayModuleEntryStates records module-level arrays that are allocated at
// every known entry into a project-local helper. A private helper is analyzed
// independently from its callers, so without this summary an allocation made
// by a public entry procedure is lost as soon as the call crosses a procedure
// boundary.
type arrayModuleEntryStates map[string]map[string]bool

// arrayModuleReadyGuardStates records the stronger, source-owned invariant
// behind a module Boolean readiness guard. The implication is intentionally
// narrow: the guard has one source-owned True write, that write is reached
// only after the module array is allocated on every path, and direct array
// invalidation is paired with a dominating False write. This lets a public
// consumer prove its module array without trusting arbitrary caller state.
type arrayModuleReadyGuardStates map[string]map[string]map[string]bool

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

type arrayParticipantGraph struct {
	all              map[string]sourceProcedure
	fileByKey        map[string]parsedFile
	byModule         map[string][]string
	keyByIdentity    map[string]string
	candidateIndex   arrayCandidateIndex
	adjacency        map[string]map[string]bool
	reverse          map[string]map[string]bool
	resolvedReverse  map[string]map[string]bool
	callAdjacency    map[string]map[string]bool
	knownSeeds       map[string]bool
	intrinsicSeeds   map[string]bool
	uncertainFacts   map[bool]map[string]bool
	uncertainCalls   map[string]bool
	moduleArrayUsers map[string][]string
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

type arrayCandidateLineKey struct {
	line int
	kind string
}

type arrayCandidateQualifiedKindKey struct {
	qualified string
	kind      string
}

type arrayCandidateIndex struct {
	byName          map[string]string
	byQualified     map[string]string
	byQualifiedKind map[arrayCandidateQualifiedKindKey]string
	byLineAndKind   map[arrayCandidateLineKey]string
}

// buildArrayCandidateIndex preserves the old sorted-key tie breaking while
// avoiding a project-wide key collection and sort for every uncertain call.
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

func inferArrayByRefAllocationSummaries(files []parsedFile, ctx analysisContext, targets map[string]sourceProcedure) arrayByRefAllocationSummaries {
	summaries := arrayByRefAllocationSummaries{}
	procedures := make([]struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
	}, 0)
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !procedureHasByRefArrayParameter(proc) || !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			procedures = append(procedures, struct {
				file        parsedFile
				proc        sourceProcedure
				moduleDecls map[string]sourceDeclaration
			}{file: file, proc: proc, moduleDecls: moduleDecls})
		}
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		for call := range procedure.proc.Calls.All() {
			if targetKey, _, ok := arrayPrivateTargetForCall(ctx, targets, call); ok {
				dependents[targetKey] = append(dependents[targetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	contributions := make(arrayByRefAllocationSummaries, len(procedures))
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		procedure := procedures[index]
		key := arrayProcedureKey(procedure.proc)
		if !arrayProcedureIsParticipant(ctx, procedure.proc) {
			continue
		}
		if ctx.arrayStats != nil && procedure.proc.Graph != nil {
			ctx.arrayStats.addCFGWalk()
		}
		value := arrayByRefAllocationSummaryForProcedure(procedure.file, procedure.proc, summaries, ctx)
		old := arrayByRefAllocationSummaries{key: contributions[key]}
		fresh := arrayByRefAllocationSummaries{key: value}
		if arrayByRefAllocationSummariesEqual(old, fresh) {
			continue
		}
		if len(value) == 0 {
			delete(contributions, key)
			delete(summaries, key)
		} else {
			contributions[key] = value
			summaries[key] = value
		}
		for _, dependent := range dependents[key] {
			if !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
		}
	}
	return summaries
}

func arrayByRefAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, summaries arrayByRefAllocationSummaries, ctx analysisContext) map[int]bool {
	if proc.Graph == nil {
		return nil
	}
	parameters := map[string]int{}
	for index, parameter := range proc.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) {
			parameters[strings.ToLower(parameter.Name)] = index
		}
	}
	if len(parameters) == 0 {
		return nil
	}
	flowCtx := ctx
	flowCtx.arrayByRefAllocations = summaries
	moduleDecls := file.moduleDecls()
	return arrayByRefFlowAllocations(file, proc, flowCtx, moduleDecls)
}

// arrayByRefFlowAllocations proves ByRef array outputs at normal procedure
// exits. Unlike the direct-assignment summary above, this pass keeps the
// branch-local state established by an IsArray/type guard and meets it at the
// normal exit. It therefore recognizes helpers that fill the same output on
// multiple accepted input branches while excluding paths that terminate in a
// direct or project-local error raiser.
func arrayByRefFlowAllocations(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) map[int]bool {
	if proc.Graph == nil {
		return nil
	}
	variables := arrayVariables(file, proc, moduleDecls)
	initial := arrayInitialState(variables)
	graph := arrayVBA227Graph(proc, ctx)
	moduleArrays := arrayModuleNamesForProcedure(file, proc, moduleDecls)
	localGoSubAllocations := arrayLocalGoSubAllocationSummaries(proc, &graph, variables, ctx, arrayOptionBase(file), arrayIntegerConstants(file, proc, nil, nil), moduleArrays)
	parameterNames := make([]string, 0, proc.Params.Len())
	for _, parameter := range proc.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) {
			parameterNames = append(parameterNames, strings.ToLower(parameter.Name))
		}
	}
	var normalExit map[string]arrayValue
	hasNormalExit := false
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		out, _ := (Analyzer{}).arrayTransfer(file, proc, ctx, variables, in, text, line, nil, nil)
		out = applyArrayLocalGoSubStatementEffects(out, text, localGoSubAllocations)
		for _, call := range arrayCallsAtLine(proc.Calls, line) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
				out = applyArrayConditionalByRefCallEffects(out, proc, call, ctx)
			} else {
				out = applyArrayByRefCallEffects(out, proc, call, ctx)
			}
			out = applyArrayLocalGoSubEffects(out, proc, call, localGoSubAllocations)
		}
		return out
	}
	edgeState := func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
		out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
		if edge.To == graph.NormalExit() {
			if !hasNormalExit {
				normalExit = make(map[string]arrayValue, len(parameterNames))
				for _, name := range parameterNames {
					normalExit[name] = out[name]
				}
				hasNormalExit = true
			} else {
				for _, name := range parameterNames {
					normalExit[name] = meetArrayValue(normalExit[name], out[name])
				}
			}
		}
		return out
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, arrayAllocationTransferIsReliable, ctx.arrayStats)
	if !hasNormalExit {
		return nil
	}
	allocated := map[int]bool{}
	for index, parameter := range proc.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) {
			continue
		}
		value, ok := normalExit[strings.ToLower(parameter.Name)]
		if ok && value.kind == arrayAllocated && value.knownArray {
			allocated[index] = true
		}
	}
	return allocated
}

// inferArrayByRefConditionalAllocations recognizes the narrow, common output
// helper contract used by argument adapters:
//
//	If items.Count = 0 Then Exit Sub
//	ReDim values(0 To items.Count - 1)
//
// The guard must dominate the ReDim, and every normal path from the guard's
// non-exit branch must pass through that ReDim. This keeps the summary tied to
// the helper's actual control flow and avoids turning an arbitrary conditional
// ReDim into an allocation guarantee.
func inferArrayByRefConditionalAllocations(files []parsedFile) arrayByRefConditionalAllocations {
	summaries := arrayByRefConditionalAllocations{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if proc.Graph == nil {
				continue
			}
			parameters := map[string]int{}
			for index, parameter := range proc.Params.AllIndexed() {
				parameters[strings.ToLower(cleanIdentifier(parameter.Name))] = index
			}
			if len(parameters) == 0 {
				continue
			}
			guards := map[string]struct {
				statementID int
				line        int
				parameter   int
			}{}
			for statement := range proc.Statements.All() {
				match := arrayByRefCountExitRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
				if len(match) != 2 {
					continue
				}
				parameter, ok := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				if !ok {
					continue
				}
				guards[strings.ToLower(cleanIdentifier(match[1]))] = struct {
					statementID int
					line        int
					parameter   int
				}{statementID: statement.ID, line: statement.Range.StartLine, parameter: parameter}
			}
			if len(guards) == 0 {
				continue
			}
			for statement := range proc.Statements.All() {
				match := arrayByRefCountRedimRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
				if len(match) != 3 || arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
					continue
				}
				output, outputOK := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				guard, guardOK := guards[strings.ToLower(cleanIdentifier(match[2]))]
				if !outputOK || !guardOK || statement.Range.StartLine <= guard.line {
					continue
				}
				if output < 0 || output >= proc.Params.Len() || !parameterIsByRefArray(proc.Params.valueAt(output)) {
					continue
				}
				guardBlock, guardBlockOK := proc.Graph.BlockForStatement(guard.statementID)
				redimBlock, redimBlockOK := proc.Graph.BlockForStatement(statement.ID)
				if !guardBlockOK || !redimBlockOK {
					continue
				}
				guardDominatesRedim := false
				for _, candidate := range proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[redimBlock.ID] {
					if candidate == guardBlock.ID {
						guardDominatesRedim = true
						break
					}
				}
				if !guardDominatesRedim {
					continue
				}
				if !arrayFalseBranchRequiresBlock(*proc.Graph, guardBlock.ID, redimBlock.ID) {
					continue
				}
				key := arrayProcedureKey(proc)
				if summaries[key] == nil {
					summaries[key] = map[int]int{}
				}
				// Conflicting output contracts remain unknown rather than being
				// overwritten by declaration order.
				if previous, exists := summaries[key][output]; exists && previous != guard.parameter {
					delete(summaries[key], output)
					continue
				}
				summaries[key][output] = guard.parameter
			}
			if len(summaries[arrayProcedureKey(proc)]) == 0 {
				delete(summaries, arrayProcedureKey(proc))
			}
		}
	}
	return summaries
}

// inferArrayByRefLengthAllocations recognizes a helper that returns the
// successful length of a ByRef array through a paired ByRef scalar output:
//
// \tbyteLength = UBound(bytes) - LBound(bytes) + 1
//
// The assignment must dominate the procedure's normal exit, or be the normal
// branch of an explicit zero-length guard whose sibling assigns zero. A
// positive length at a caller then proves that the array-bound query completed
// successfully, without making the array unconditionally allocated on other
// paths.
func inferArrayByRefLengthAllocations(files []parsedFile) arrayByRefLengthAllocations {
	summaries := arrayByRefLengthAllocations{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if proc.Graph == nil {
				continue
			}
			parameters := map[string]int{}
			for index, parameter := range proc.Params.AllIndexed() {
				parameters[strings.ToLower(cleanIdentifier(parameter.Name))] = index
			}
			if len(parameters) == 0 {
				continue
			}
			dominators := arrayProcedureNormalExitDominators(proc)
			for statement := range proc.Statements.All() {
				text := strings.TrimSpace(normalizedCodeLine(statement.Text))
				match := arrayByRefLengthFullRe.FindStringSubmatch(text)
				if len(match) == 4 && !strings.EqualFold(cleanIdentifier(match[2]), cleanIdentifier(match[3])) {
					match = nil
				}
				if len(match) == 0 {
					upper := arrayByRefLengthUpperRe.FindStringSubmatch(text)
					if len(upper) == 3 {
						match = []string{upper[0], upper[1], upper[2], upper[2]}
					}
				}
				if len(match) != 4 {
					continue
				}
				lengthIndex, lengthOK := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				arrayIndex, arrayOK := parameters[strings.ToLower(cleanIdentifier(match[2]))]
				if !lengthOK || !arrayOK || lengthIndex == arrayIndex || !parameterIsByRefScalar(proc.Params.valueAt(lengthIndex)) || !parameterIsByRefArray(proc.Params.valueAt(arrayIndex)) {
					continue
				}
				dominatesExit := arrayProcedureBlockDominatesNormalExit(proc, statement.ID, dominators)
				if !dominatesExit {
					parent, parentOK := arrayByRefLengthGuard(proc, statement.ID)
					if !parentOK || !arrayProcedureBlockDominatesNormalExit(proc, parent.ID, dominators) || !arrayByRefLengthHasZeroBranch(proc, parent.ID, statement.ID, match[1]) {
						continue
					}
				}
				key := arrayProcedureKey(proc)
				if summaries[key] == nil {
					summaries[key] = map[int]int{}
				}
				if previous, exists := summaries[key][arrayIndex]; exists && previous != lengthIndex {
					delete(summaries[key], arrayIndex)
					continue
				}
				summaries[key][arrayIndex] = lengthIndex
			}
			if len(summaries[arrayProcedureKey(proc)]) == 0 {
				delete(summaries, arrayProcedureKey(proc))
			}
		}
	}
	return summaries
}

func arrayProcedureStatementByID(proc sourceProcedure, id int) (procedureir.Statement, bool) {
	for statement := range proc.Statements.All() {
		if statement.ID == id {
			return statement, true
		}
	}
	return procedureir.Statement{}, false
}

func arrayByRefLengthGuard(proc sourceProcedure, statementID int) (procedureir.Statement, bool) {
	seen := map[int]bool{}
	for statementID != 0 && !seen[statementID] {
		seen[statementID] = true
		statement, ok := arrayProcedureStatementByID(proc, statementID)
		if !ok {
			return procedureir.Statement{}, false
		}
		if statement.Kind == procedureir.StatementIf {
			return statement, true
		}
		statementID = statement.ParentID
	}
	return procedureir.Statement{}, false
}

func arrayByRefLengthHasZeroBranch(proc sourceProcedure, parentID, formulaID int, lengthName string) bool {
	for statement := range proc.Statements.All() {
		if statement.ID == formulaID || statement.ParentID != parentID {
			continue
		}
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(statement.Text))
		if ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(lengthName)) && strings.TrimSpace(rhs) == "0" {
			return true
		}
	}
	return false
}

func arrayFalseBranchRequiresBlock(graph vbacfg.Graph, guardBlock, requiredBlock vbacfg.BlockID) bool {
	queue := make([]vbacfg.BlockID, 0, 1)
	for _, edge := range graph.Edges {
		if edge.From == guardBlock && edge.Kind == vbacfg.EdgeBranchFalse && edge.Class == vbacfg.EdgeNormal {
			queue = append(queue, edge.To)
		}
	}
	if len(queue) == 0 {
		return false
	}
	visited := map[vbacfg.BlockID]bool{}
	reachedRequired := false
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if visited[current] {
			continue
		}
		visited[current] = true
		if current == requiredBlock {
			reachedRequired = true
			continue
		}
		if current == graph.NormalExit {
			return false
		}
		for _, edge := range graph.Edges {
			if edge.From != current || edge.Class != vbacfg.EdgeNormal {
				continue
			}
			queue = append(queue, edge.To)
		}
	}
	return reachedRequired
}

func arrayByRefAllocationSummariesEqual(left, right arrayByRefAllocationSummaries) bool {
	if len(left) != len(right) {
		return false
	}
	for procedure, parameters := range left {
		other, ok := right[procedure]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index := range parameters {
			if !other[index] {
				return false
			}
		}
	}
	return true
}

func inferArrayModuleConfigurationStates(files []parsedFile, summaries arrayModuleAllocationSummaries) map[string]arrayModuleConfigurationState {
	states := map[string]arrayModuleConfigurationState{}
	for _, file := range files {
		state := arrayModuleConfigurationState{byProcedure: map[string]map[string]bool{}}
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(strings.TrimSpace(proc.Name))
			if !strings.HasPrefix(name, "configure") {
				continue
			}
			arrays := summaries[arrayProcedureKey(proc)]
			if len(arrays) == 0 {
				continue
			}
			state.byProcedure[name] = cloneArrayNameSet(arrays)
			switch name {
			case "configuredatatable":
				state.dataTable = mergeArrayNameSets(state.dataTable, arrays)
			case "configuregenericcollection":
				state.genericCollection = mergeArrayNameSets(state.genericCollection, arrays)
			}
		}
		if len(state.byProcedure) > 0 {
			states[file.Path] = state
		}
	}
	return states
}

func cloneArrayNameSet(values map[string]bool) map[string]bool {
	if len(values) == 0 {
		return nil
	}
	clone := make(map[string]bool, len(values))
	for name, allocated := range values {
		clone[name] = allocated
	}
	return clone
}

func mergeArrayNameSets(left, right map[string]bool) map[string]bool {
	if len(right) == 0 {
		return left
	}
	if left == nil {
		left = map[string]bool{}
	}
	for name := range right {
		left[name] = true
	}
	return left
}

func applyArrayModuleCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	conditional := arrayProcedureLineHasInlineConditional(file, call.Range.StartLine)
	if conditional && arrayProcedureLineInlineConditionIsFalse(file, call.Range.StartLine) {
		return state
	}
	key, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	// Callers pass a block-local state to this transfer callback. The legacy
	// walkers keep predecessor input separate before invoking it, while the
	// compact cursor owns the map outright; update that private state in place
	// so module/ByRef effects do not reintroduce whole-state copies.
	updated := state
	if updated == nil {
		updated = arrayFlowState{}
	}
	markArgument := func(name string) {
		name = strings.ToLower(cleanIdentifier(name))
		variable, known := variables[name]
		if !known || !variable.isArray {
			return
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	markModule := func(name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			return
		}
		declaration, declared := moduleDecls[name]
		if !declared || !declaration.Array || declaration.Parameter {
			return
		}
		markArgument(name)
	}
	if !conditional {
		for name := range ctx.arrayModuleAllocations[key] {
			markModule(name)
		}
	}
	arguments, mapped := arrayCallFormalArguments(proc, target, call)
	if mapped && !conditional {
		for index := range ctx.arrayByRefAllocations[key] {
			if index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
				continue
			}
			markArgument(arguments[index])
		}
		for outputIndex, countIndex := range ctx.arrayByRefConditionalAllocations[key] {
			if outputIndex < 0 || outputIndex >= target.Params.Len() || countIndex < 0 || countIndex >= target.Params.Len() {
				continue
			}
			outputName := directArrayArgumentName(arguments[outputIndex])
			countSource := directArrayArgumentName(arguments[countIndex])
			if outputName == "" || countSource == "" {
				continue
			}
			name := strings.ToLower(outputName)
			variable, known := variables[name]
			if !known || !variable.isArray {
				continue
			}
			value := updated[name]
			if value.kind == arrayAllocated && value.knownArray {
				continue
			}
			if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, countSource) {
				continue
			}
			value.allocationCountSource = countSource
			updated[name] = value
		}
		for outputIndex, lengthIndex := range ctx.arrayByRefLengthAllocations[key] {
			if outputIndex >= len(arguments) || lengthIndex < 0 || lengthIndex >= len(arguments) {
				continue
			}
			outputName := directArrayArgumentName(arguments[outputIndex])
			lengthSource := directArrayArgumentName(arguments[lengthIndex])
			if outputName == "" || lengthSource == "" {
				continue
			}
			name := strings.ToLower(outputName)
			variable, known := variables[name]
			if !known || !variable.isArray {
				continue
			}
			value := updated[name]
			if value.kind == arrayAllocated && value.knownArray {
				continue
			}
			if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, lengthSource) {
				continue
			}
			value.allocationCountSource = lengthSource
			updated[name] = value
		}
		for name := range arrayConfigurationArraysForGuard(file, target, arguments, ctx.arrayModuleConfigurations[file.Path]) {
			markModule(name)
		}
	}
	if !ctx.arraySkipModuleInvalidationEffects {
		invalidated := arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)
		for name := range invalidated {
			if declarations.shadowsModule(name) {
				continue
			}
			if _, tracked := updated[name]; tracked {
				updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
			}
		}
	}
	return updated
}

// arrayPrivateModuleArrayInvalidations returns the normal-return invalidation
// summary for a project-local helper. A precomputed summary is used by the
// production context; focused compatibility callers compute the same summary
// on demand.
func arrayPrivateModuleArrayInvalidations(file parsedFile, target sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	key := arrayProcedureKey(target)
	if ctx.arrayModuleInvalidations != nil {
		if summary, ok := ctx.arrayModuleInvalidations[key]; ok {
			return summary
		}
	}
	return arrayPrivateModuleArrayInvalidationsWithVisiting(file, target, moduleDecls, ctx, map[string]bool{})
}

// arrayPrivateModuleArrayInvalidationsWithVisiting identifies module arrays
// that are not proven allocated at a target's normal exit. The summary starts
// from allocated arrays so it models the effect on a caller that already has
// a valid module-array allocation. Direct operations are evaluated through the
// normal CFG, and resolved local calls are summarized recursively.
func arrayPrivateModuleArrayInvalidationsWithVisiting(file parsedFile, target sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, visiting map[string]bool) map[string]bool {
	key := arrayProcedureKey(target)
	if !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
		return nil
	}
	names := arrayModuleNamesForProcedure(file, target, moduleDecls)
	if len(names) == 0 {
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = nil
		}
		return nil
	}
	if visiting[key] {
		// A recursive effect cycle has no finite normal-return proof. Keep all
		// visible module arrays conservative rather than assuming the cycle is
		// read-only.
		return names
	}
	if ctx.arrayModuleInvalidations != nil {
		if summary, ok := ctx.arrayModuleInvalidations[key]; ok {
			return summary
		}
	}
	visiting[key] = true
	defer delete(visiting, key)

	variables := arrayVariables(file, target, moduleDecls)
	initial := arrayInitialState(variables)
	for name := range names {
		value := initial[name]
		value.kind = arrayAllocated
		value.knownArray = true
		initial[name] = value
	}
	if target.Graph == nil {
		state := initial
		constants := arrayIntegerConstants(file, target, nil, nil)
		if target.Statements.Len() > 0 {
			for statement := range target.Statements.All() {
				line := statement.Range.StartLine
				if line < 1 {
					line = target.StartLine
				}
				text := strings.TrimSpace(normalizedCodeLine(statement.Text))
				if text == "" && line >= 1 && line <= len(file.Lines) {
					text = normalizedCodeLine(file.Lines[line-1])
				}
				if text == "" {
					continue
				}
				state = arrayModuleSummaryTransfer(file, target, ctx, variables, state, text, line, constants, moduleDecls, names, visiting)
			}
		} else {
			for line := target.StartLine; line <= target.EndLine && line <= len(file.Lines); line++ {
				state = arrayModuleSummaryTransfer(file, target, ctx, variables, state, normalizedCodeLine(file.Lines[line-1]), line, constants, moduleDecls, names, visiting)
			}
		}
		result := arrayModuleInvalidationsFromState(names, state)
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = result
		}
		return result
	}

	graph := target.Graph.View(vbacfg.EdgeFilter{NormalOnly: true, WithoutNormalErrRaiseContinuation: true})
	constants := arrayIntegerConstants(file, target, nil, nil)
	var normalExit arrayFlowState
	hasNormalExit := false
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		return arrayModuleSummaryTransfer(file, target, ctx, variables, in, text, line, constants, moduleDecls, names, visiting)
	}
	edgeState := func(_ vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		if edge.To != graph.NormalExit() {
			return out
		}
		if !hasNormalExit {
			normalExit = cloneArrayState(out)
			hasNormalExit = true
		} else {
			normalExit = meetArrayState(normalExit, out)
		}
		return out
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, nil, ctx.arrayStats)
	if !hasNormalExit {
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = nil
		}
		return nil
	}
	result := arrayModuleInvalidationsFromState(names, normalExit)
	if ctx.arrayModuleInvalidationCacheWritable && key != "" {
		ctx.arrayModuleInvalidations[key] = result
	}
	return result
}

func arrayModuleInvalidationsFromState(names map[string]bool, state arrayFlowState) map[string]bool {
	invalidated := map[string]bool{}
	for name := range names {
		value, known := state[name]
		if !known || value.kind != arrayAllocated || !value.knownArray {
			invalidated[name] = true
		}
	}
	if len(invalidated) == 0 {
		return nil
	}
	return invalidated
}

func arrayModuleSummaryTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	if condition, body, ok := arrayIfThenParts(text); ok && strings.TrimSpace(body) != "" {
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		condition = strings.TrimSpace(condition)
		lowerCondition := strings.ToLower(condition)
		switch {
		case strings.HasPrefix(lowerCondition, "if "):
			condition = strings.TrimSpace(condition[len("if "):])
		case strings.HasPrefix(lowerCondition, "elseif "):
			condition = strings.TrimSpace(condition[len("elseif "):])
		}
		thenState := arrayModuleSummaryTransferParts(file, proc, ctx, variables, cloneArrayState(state), thenBody, line, constants, moduleDecls, moduleArrays, visiting)
		elseState := cloneArrayState(state)
		if hasElse {
			elseState = arrayModuleSummaryTransferParts(file, proc, ctx, variables, elseState, elseBody, line, constants, moduleDecls, moduleArrays, visiting)
		}
		result := meetArrayState(thenState, elseState)
		applyCalls := func(value arrayFlowState, conditionValue, conditionKnown bool) arrayFlowState {
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				if conditionKnown && !arrayInlineConditionalCallIsReachable(file, call, conditionValue, hasElse) {
					continue
				}
				value = applyArrayModuleSummaryCallEffects(value, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
			}
			return value
		}
		if value, known := arraySourceOrderConstantBoolean(condition, constants); known {
			if value {
				result = applyCalls(thenState, true, true)
			} else if hasElse {
				result = applyCalls(elseState, false, true)
			} else {
				result = state
			}
		} else if !hasElse {
			result = applyCalls(meetArrayState(thenState, state), false, false)
		} else {
			result = applyCalls(result, false, false)
		}
		return result
	}
	state = arrayModuleSummaryTransferParts(file, proc, ctx, variables, state, text, line, constants, moduleDecls, moduleArrays, visiting)
	callsAtLine := arrayCallsAtLine(proc.Calls, line)
	for _, call := range callsAtLine {
		state = applyArrayModuleSummaryCallEffects(state, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
	}
	return state
}

func arrayModuleSummaryTransferParts(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	for _, part := range splitRangeValueSourceStatements(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		if condition, body, ok := arrayIfThenParts(part); ok && strings.TrimSpace(body) != "" {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			condition = strings.TrimSpace(condition)
			lowerCondition := strings.ToLower(condition)
			switch {
			case strings.HasPrefix(lowerCondition, "if "):
				condition = strings.TrimSpace(condition[len("if "):])
			case strings.HasPrefix(lowerCondition, "elseif "):
				condition = strings.TrimSpace(condition[len("elseif "):])
			}
			thenState := arrayModuleSummaryTransferParts(file, proc, ctx, variables, cloneArrayState(state), thenBody, line, constants, moduleDecls, moduleArrays, visiting)
			elseState := cloneArrayState(state)
			if hasElse {
				elseState = arrayModuleSummaryTransferParts(file, proc, ctx, variables, elseState, elseBody, line, constants, moduleDecls, moduleArrays, visiting)
			}
			if value, known := arraySourceOrderConstantBoolean(condition, constants); known {
				if value {
					state = thenState
				} else if hasElse {
					state = elseState
				}
			} else if hasElse {
				state = meetArrayState(thenState, elseState)
			} else {
				state = meetArrayState(thenState, state)
			}
			continue
		}
		state, _ = (Analyzer{}).arrayTransfer(file, proc, ctx, variables, state, part, line, constants, nil)
	}
	return state
}

func applyArrayModuleSummaryCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	// The procedure IR also represents an indexed array expression such as
	// `result(index)` or `mZipWork(offset)` as a CallSite. Those expressions
	// are handled by arrayTransfer; they are not procedure calls whose module
	// effects belong in this summary. Without this guard, the resolver can
	// reinterpret the expression as a source-local target and recursively
	// invalidate every module array while summarizing an otherwise read-only
	// indexed assignment.
	if call.Callee.Receiver == nil {
		name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
		if variable, ok := variables[name]; ok && (variable.isArray || variable.isVariant) {
			return state
		}
	}
	key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !resolved {
		target, resolved = arraySourceModuleTargetForCall(file, call, ctx)
		if !resolved {
			return state
		}
		key = arrayProcedureKey(target)
	}
	if !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
		return state
	}
	updated := cloneArrayState(state)
	if !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
		for name := range ctx.arrayModuleAllocations[key] {
			name = strings.ToLower(cleanIdentifier(name))
			if moduleArrays[name] {
				value := updated[name]
				value.kind = arrayAllocated
				value.knownArray = true
				updated[name] = value
			}
		}
	}
	for name := range arrayPrivateModuleArrayInvalidationsWithVisiting(file, target, moduleDecls, ctx, visiting) {
		if moduleArrays[name] {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	if !procedureHasByRefArrayParameter(target) {
		return updated
	}
	bindings, mapped := arrayCallArgumentBindings(proc, target, call)
	if !mapped {
		return updated
	}
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		name := strings.ToLower(directArrayArgumentName(binding.text))
		if !moduleArrays[name] {
			continue
		}
		if ctx.arrayByRefAllocations[key][binding.parameterIndex] {
			updated[name] = arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

func arrayPrivateCallMayInvalidateModuleArray(file parsedFile, caller sourceProcedure, target sourceProcedure, call procedureir.CallSite, moduleDecls map[string]sourceDeclaration, ctx analysisContext) bool {
	if len(arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)) > 0 {
		return true
	}
	if !procedureHasByRefArrayParameter(target) {
		return false
	}
	bindings, mapped := arrayCallArgumentBindings(caller, target, call)
	if !mapped {
		return true
	}
	for _, binding := range bindings {
		name := strings.ToLower(directArrayArgumentName(binding.text))
		declaration, declared := moduleDecls[name]
		if !declared || !declaration.Array || declaration.Parameter || binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			return true
		}
	}
	return false
}

// applyArrayUnknownModuleCallEffects is used by the recovered source-order
// path, where a public call cannot be matched to a private module effect
// summary. A source-local public target can still be inspected for direct
// module-array effects, including when it receives no explicit array
// argument. Calls that do not identify a source-local target are left alone:
// treating every unresolved/external call as a mutation of every visible
// module array turns unrelated object construction and host calls into false
// positives.
func applyArrayUnknownModuleCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	if ctx.arraySkipModuleInvalidationEffects {
		return state
	}
	if arrayProcedureLineInlineConditionIsFalse(file, call.Range.StartLine) {
		return state
	}
	if _, _, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call); resolved {
		return state
	}
	target, ok := arraySourceModuleTargetForCall(file, call, ctx)
	if !ok || !procedureUsesModuleArray(file, target, moduleDecls) {
		return state
	}
	invalidated := arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)
	if len(invalidated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range invalidated {
		if declarations.shadowsModule(name) {
			continue
		}
		variable, known := variables[name]
		if !known || !variable.isArray {
			continue
		}
		if _, tracked := updated[name]; tracked {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

// arraySourceModuleTargetForCall recovers a source-local target for the
// source-order fallback. The project resolver may report a public target, but
// it may also be incomplete for a recovered call. In the latter case a unique
// same-module source procedure is sufficient evidence for inspecting that
// procedure's module-array accesses; an absent target remains an external or
// late-bound call and must not invalidate all module arrays.
func arraySourceModuleTargetForCall(file parsedFile, call procedureir.CallSite, ctx analysisContext) (sourceProcedure, bool) {
	procedures := file.procedureView()
	if procedures.Len() == 0 {
		return sourceProcedure{}, false
	}
	resolution := call.Resolution
	if ctx.procedureResolver != nil {
		resolution = ctx.procedureResolver.ResolveCall(call)
	}
	if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
		qualifiedName := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
		for index := 0; index < procedures.Len(); index++ {
			target := procedures.valueAt(index)
			if strings.EqualFold(arrayProcedureKey(target), qualifiedName) {
				return target, true
			}
		}
	}
	// A receiver-bearing call that was not resolved to the exact source target
	// belongs to another object or to a late-bound member. Do not reinterpret a
	// same-module procedure with the same member name as its target.
	if call.Callee.Receiver != nil {
		return sourceProcedure{}, false
	}

	baseName := cleanIdentifier(strings.TrimPrefix(strings.TrimSpace(call.Callee.BaseName), "New "))
	if baseName == "" {
		baseName = cleanIdentifier(strings.TrimPrefix(strings.TrimSpace(call.Callee.Text), "New "))
	}
	callerModule := strings.TrimSpace(call.Caller.QualifiedName)
	if dot := strings.IndexByte(callerModule, '.'); dot >= 0 {
		callerModule = callerModule[:dot]
	}
	if callerModule == "" {
		callerModule = strings.TrimSpace(call.Module)
	}
	if callerModule == "" {
		callerModule = strings.TrimSpace(file.Module)
	}
	var match sourceProcedure
	matched := 0
	for index := 0; index < procedures.Len(); index++ {
		candidate := procedures.valueAt(index)
		if !strings.EqualFold(strings.TrimSpace(candidate.Name), baseName) ||
			!strings.EqualFold(strings.TrimSpace(candidate.Module), callerModule) {
			continue
		}
		visibility := strings.TrimSpace(candidate.Visibility)
		if strings.EqualFold(visibility, "Private") || strings.EqualFold(visibility, "Friend") {
			continue
		}
		match = candidate
		matched++
	}
	if matched == 1 {
		return match, true
	}
	return sourceProcedure{}, false
}

// applyArrayByRefCallEffects carries the invalidating side of a private
// ByRef-array contract. The ordinary CFG walk records proven allocation
// outputs, but the recovered source-order fallback also has to account for a
// preceding helper that can Erase or otherwise replace the caller's array.
// Read-only helpers preserve the caller's allocation state; helpers whose
// parameter effect cannot be proven make it unknown.
func applyArrayByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	key, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return applyArrayUnknownByRefCallEffects(state, proc, call, ctx)
	}
	if !procedureHasByRefArrayParameter(target) {
		return state
	}
	bindings, ok := arrayCallArgumentBindings(proc, target, call)
	if !ok {
		return state
	}
	// The compact CFG solver may revisit sibling branch blocks with a state map
	// that shares storage with their predecessor. Keep this call's mutation
	// isolated so an invalidating true branch cannot poison its false sibling.
	updated := cloneArrayState(state)
	apply := func(index int, argument string) {
		if index < 0 || index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
			return
		}
		name := directArrayArgumentName(argument)
		if name == "" {
			return
		}
		if ctx.arrayByRefAllocations[key][index] {
			value := updated[name]
			value.kind = arrayAllocated
			value.knownArray = true
			updated[name] = value
			return
		}
		// Conditional and paired-length output contracts carry a stronger
		// caller-side fact than the generic invalidation scan. Their output may
		// be unallocated for the zero-length case, but the existing count/length
		// refinement will establish the successful branch without losing that
		// relation to an unknown state here.
		if _, conditional := ctx.arrayByRefConditionalAllocations[key][index]; conditional {
			return
		}
		if _, pairedLength := ctx.arrayByRefLengthAllocations[key][index]; pairedLength {
			return
		}
		if arrayByRefParameterMayInvalidate(target, index, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	for _, binding := range bindings {
		apply(binding.parameterIndex, binding.text)
	}
	return updated
}

func applyArrayUnknownByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	if _, _, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call); resolved {
		return state
	}
	updated := cloneArrayState(state)
	markUnknown := func(argument string) {
		name := directArrayArgumentName(argument)
		if name == "" {
			return
		}
		if _, tracked := updated[name]; tracked {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	for _, argument := range call.Arguments.Named {
		markUnknown(argument.ValueText)
	}
	for _, argument := range arrayCallArgumentTexts(proc, call) {
		markUnknown(argument)
	}
	return updated
}

// applyArrayConditionalByRefCallEffects joins the two possible outcomes of
// an inline conditional call without treating the call's allocation summary
// as unconditional. A helper that may invalidate its ByRef array makes the
// joined state unknown; a helper that only allocates leaves the prior state,
// which is the conservative meet of the conditional paths.
func applyArrayConditionalByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return applyArrayUnknownByRefCallEffects(state, proc, call, ctx)
	}
	if !procedureHasByRefArrayParameter(target) {
		return state
	}
	bindings, ok := arrayCallArgumentBindings(proc, target, call)
	if !ok {
		return state
	}
	updated := cloneArrayState(state)
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		name := directArrayArgumentName(binding.text)
		if name == "" {
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

// arrayByRefParameterMayInvalidate reports whether a private ByRef-array
// parameter can lose its allocation while the procedure returns normally.
// Element writes and reads preserve an allocated input; Erase, unknown whole
// array replacement, and an unproven nested ByRef call do not. The recursion
// guard keeps recursive helper cycles from being treated as an additional
// mutation. Any direct mutation in the current cycle member is still found by
// its own scan.
func arrayByRefParameterMayInvalidate(proc sourceProcedure, parameterIndex int, ctx analysisContext, visiting map[string]bool) bool {
	if parameterIndex < 0 || parameterIndex >= proc.Params.Len() || !parameterIsByRefArray(proc.Params.valueAt(parameterIndex)) {
		return true
	}
	key := strings.ToLower(arrayProcedureKey(proc)) + "#" + strconv.Itoa(parameterIndex)
	if visiting[key] {
		return false
	}
	visiting[key] = true
	defer delete(visiting, key)
	name := strings.ToLower(cleanIdentifier(proc.Params.valueAt(parameterIndex).Name))
	names := map[string]bool{name: true}
	for statement := range proc.Statements.All() {
		if !arrayByRefStatementReachable(proc, statement) {
			continue
		}
		if statement.Recovered {
			return true
		}
		for _, part := range splitRangeValueSourceStatements(statement.Text) {
			part = strings.TrimSpace(part)
			if part == "" {
				continue
			}
			if arrayByRefParameterInlineMutation(part, names, ctx) {
				return true
			}
		}
		for _, nested := range arrayCallsAtLine(proc.Calls, statement.Range.StartLine) {
			if arrayByRefCallIsReadOnly(nested) {
				continue
			}
			if !arrayCallPassesDirectArrayArgument(proc, nested, name) {
				continue
			}
			nestedKey, nestedTarget, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, nested)
			if !resolved {
				return true
			}
			bindings, mapped := arrayCallArgumentBindings(proc, nestedTarget, nested)
			if !mapped {
				return true
			}
			for _, binding := range bindings {
				if directArrayArgumentName(binding.text) != name || binding.parameterIndex >= nestedTarget.Params.Len() || !parameterIsByRefArray(nestedTarget.Params.valueAt(binding.parameterIndex)) {
					continue
				}
				if ctx.arrayByRefAllocations[nestedKey][binding.parameterIndex] {
					continue
				}
				if arrayByRefParameterMayInvalidate(nestedTarget, binding.parameterIndex, ctx, visiting) {
					return true
				}
			}
		}
	}
	return false
}

func arrayByRefCallIsReadOnly(call procedureir.CallSite) bool {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return true
	}
	switch strings.ToLower(cleanIdentifier(call.Callee.BaseName)) {
	case "lbound", "ubound":
		return true
	default:
		return false
	}
}

func arrayByRefStatementReachable(proc sourceProcedure, statement procedureir.Statement) bool {
	if proc.Graph == nil {
		return true
	}
	block, ok := proc.Graph.BlockForStatement(statement.ID)
	if !ok {
		return true
	}
	return proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).IsReachable(block.ID)
}

func arrayByRefParameterInlineMutation(text string, names map[string]bool, ctx analysisContext) bool {
	if condition, body, ok := arrayIfThenParts(text); ok && strings.TrimSpace(body) != "" {
		condition = strings.TrimSpace(condition)
		lowerCondition := strings.ToLower(condition)
		switch {
		case strings.HasPrefix(lowerCondition, "if "):
			condition = strings.TrimSpace(condition[len("if "):])
		case strings.HasPrefix(lowerCondition, "elseif "):
			condition = strings.TrimSpace(condition[len("elseif "):])
		}
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		branches := []string{thenBody}
		if hasElse {
			branches = append(branches, elseBody)
		}
		if value, known := arraySourceOrderConstantBoolean(condition, nil); known {
			if value {
				branches = []string{thenBody}
			} else if hasElse {
				branches = []string{elseBody}
			} else {
				branches = nil
			}
		}
		for _, branch := range branches {
			for _, statement := range splitRangeValueSourceStatements(branch) {
				if arrayByRefParameterInlineMutation(strings.TrimSpace(statement), names, ctx) {
					return true
				}
			}
		}
		return false
	}
	text = strings.TrimSpace(text)
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
		for _, target := range splitArgs(match[1]) {
			if names[strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))] {
				return true
			}
		}
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			name := strings.ToLower(cleanIdentifier(redim.name))
			if !direct || !names[name] {
				continue
			}
			// ReDim Preserve retains an already allocated input array. Its shape
			// may change, but that is not an allocation invalidation at the
			// caller boundary.
			if strings.TrimSpace(match[1]) != "" {
				continue
			}
			if !arrayStatementAllocatesName(text, name, ctx) {
				return true
			}
		}
	}
	if lhs, _, indexed, ok := arrayAssignment(text); ok && !indexed {
		name := strings.ToLower(cleanIdentifier(lhs))
		return names[name] && !arrayStatementAllocatesName(text, name, ctx)
	}
	return false
}

func arrayConfigurationArraysForGuard(file parsedFile, target sourceProcedure, arguments []string, configurations arrayModuleConfigurationState) map[string]bool {
	name := strings.ToLower(strings.TrimSpace(target.Name))
	if !strings.HasPrefix(name, "require") || !arrayGuardProcedureRejectsInvalidState(file, target) || !arrayGuardTargetsCurrentObject(target, arguments) {
		return nil
	}
	if arrays := configurations.byProcedure["configure"+strings.TrimPrefix(name, "require")]; len(arrays) > 0 {
		return arrays
	}
	if name == "requireerror" {
		if arrays := configurations.byProcedure["configureaggregateerror"]; len(arrays) > 0 {
			return arrays
		}
	}
	body := strings.ToLower(strings.Join(file.Lines[max(0, target.StartLine-1):min(len(file.Lines), target.EndLine)], "\n"))
	if strings.Contains(body, "role_data_table") {
		return configurations.dataTable
	}
	if arrayGuardUsesGenericCollectionConfiguration(body) {
		return configurations.genericCollection
	}
	return nil
}

// applyArrayInternalStorageConfiguration carries the class-instance array
// contract into Friend/Private storage members that are called through a
// configured receiver. These members intentionally do not repeat a public
// role guard: their callers have already established the owning collection,
// data-row, or aggregate-error configuration on that receiver.
func applyArrayInternalStorageConfiguration(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, configurations arrayModuleConfigurationState) arrayFlowState {
	if !strings.EqualFold(strings.TrimSpace(proc.ModuleKind), "class") {
		return state
	}
	arrays := arrayInternalStorageConfigurationArrays(proc, configurations)
	if len(arrays) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayInternalStorageConfigurationArrays(proc sourceProcedure, configurations arrayModuleConfigurationState) map[string]bool {
	name := strings.ToLower(strings.TrimSpace(proc.Name))
	switch name {
	case "internalcollectionitems", "internalcollectionkeys", "internalcollectionpriorities",
		"internaladdlookupgroup", "internalappendcollectionitem", "internalappendcollectionkey",
		"internalappendcollectionpriority", "internalqueuevalue", "internalpushvalue":
		return configurations.genericCollection
	case "internaladdwrapped":
		return mergeArrayNameSets(cloneArrayNameSet(configurations.byProcedure["configurelist"]), configurations.genericCollection)
	case "internaldatacolumns":
		return configurations.dataTable
	case "internaldatarows", "internalappendrowcell", "acceptrowchanges", "rejectrowchanges":
		return configurations.byProcedure["configuredatarow"]
	case "internalinnerexceptions":
		return configurations.byProcedure["configureaggregateerror"]
	default:
		return nil
	}
}

func arrayGuardUsesGenericCollectionConfiguration(body string) bool {
	for _, marker := range []string{
		"isgenericcollectionrole",
		"ispriorityqueuekind",
		"isdictionarycollection",
		"issetcollection",
		"mcollectionkind",
		"ROLE_DICTIONARY",
		"ROLE_HASH_SET",
		"ROLE_COLLECTION",
		"ROLE_IMMUTABLE",
		"ROLE_CONCURRENT",
	} {
		if strings.Contains(body, strings.ToLower(marker)) {
			return true
		}
	}
	return false
}

func arrayGuardProcedureRejectsInvalidState(file parsedFile, target sourceProcedure) bool {
	start := max(0, target.StartLine-1)
	end := min(len(file.Lines), target.EndLine)
	if start >= end {
		return false
	}
	body := strings.ToLower(strings.Join(file.Lines[start:end], "\n"))
	return strings.Contains(body, "err.raise") || strings.Contains(body, "raisecontracterror")
}

func arrayGuardTargetsCurrentObject(target sourceProcedure, arguments []string) bool {
	if target.Params.Len() <= 1 {
		return true
	}
	return len(arguments) > 0 && strings.EqualFold(strings.TrimSpace(arguments[0]), "me")
}

func applyArrayModuleConfigurationBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, configurations arrayModuleConfigurationState, variables map[string]arrayVariable, file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue || statement.Condition == nil {
		return state
	}
	condition := strings.ToLower(strings.TrimSpace(statement.Condition.Text))
	var arrays map[string]bool
	switch {
	case arrayPositiveGenericCollectionKindBranch(condition):
		arrays = configurations.genericCollection
	case strings.Contains(condition, "role_immutable") && !strings.Contains(condition, "<> role_immutable"):
		arrays = configurations.genericCollection
	case strings.Contains(condition, "role_list") && !strings.Contains(condition, "<> role_list"):
		arrays = configurations.byProcedure["configurelist"]
	case strings.Contains(condition, "role_data_row") && !strings.Contains(condition, "<> role_data_row"):
		arrays = configurations.byProcedure["configuredatarow"]
	case strings.Contains(condition, "role_data_table") && !strings.Contains(condition, "<> role_data_table"):
		arrays = configurations.dataTable
	}
	if len(arrays) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayPositiveGenericCollectionKindBranch(condition string) bool {
	for _, marker := range []string{
		"isgenericcollectionrole",
		"ispriorityqueuekind",
		"issortedmapkind",
		"issortedsetkind",
	} {
		if strings.Contains(condition, marker) && !strings.Contains(condition, "not "+marker) {
			return true
		}
	}
	return false
}

func inferArrayModuleAllocationSummaries(files []parsedFile, ctx analysisContext, targets map[string]sourceProcedure, byRefSummaries arrayByRefAllocationSummaries) arrayModuleAllocationSummaries {
	summaries := arrayModuleAllocationSummaries{}
	procedures := make([]struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
	}, 0)
	// Package-local synthetic callers may omit ModuleFacts. Attach one local
	// immutable instance before expanding the procedure list so compatibility
	// paths do not rebuild and rescan module source once per procedure.
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			procedures = append(procedures, struct {
				file        parsedFile
				proc        sourceProcedure
				moduleDecls map[string]sourceDeclaration
			}{file: file, proc: proc, moduleDecls: moduleDecls})
		}
	}
	if len(procedures) == 0 {
		return summaries
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	dominators := arrayProcedureDominators{}
	for _, procedure := range procedures {
		dominators[arrayProcedureKey(procedure.proc)] = arrayProcedureNormalExitDominators(procedure.proc)
	}

	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		for call := range procedure.proc.Calls.All() {
			if targetKey, _, ok := arrayPrivateTargetForCall(ctx, targets, call); ok {
				dependents[targetKey] = append(dependents[targetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	contributions := make(arrayModuleAllocationSummaries, len(procedures))
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		procedure := procedures[index]
		if !arrayProcedureIsParticipant(ctx, procedure.proc) {
			continue
		}
		key := arrayProcedureKey(procedure.proc)
		value := arrayModuleAllocationSummaryForProcedure(procedure.file, procedure.proc, procedure.moduleDecls, targets, summaries, byRefSummaries, ctx, dominators[key])
		old := arrayModuleAllocationSummaries{key: contributions[key]}
		fresh := arrayModuleAllocationSummaries{key: value}
		if arrayModuleAllocationSummariesEqual(old, fresh) {
			continue
		}
		if len(value) == 0 {
			delete(contributions, key)
			delete(summaries, key)
		} else {
			contributions[key] = value
			summaries[key] = value
		}
		for _, dependent := range dependents[key] {
			if !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
		}
	}
	return summaries
}

func arrayModuleAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, targets map[string]sourceProcedure, summaries arrayModuleAllocationSummaries, byRefSummaries arrayByRefAllocationSummaries, ctx analysisContext, dominators map[vbacfg.BlockID]bool) map[string]bool {
	moduleArrays := map[string]bool{}
	for name, declaration := range moduleDecls {
		if declaration.Array && !declaration.Parameter {
			moduleArrays[strings.ToLower(name)] = true
		}
	}
	if len(moduleArrays) == 0 || proc.Graph == nil {
		return nil
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	for name := range moduleArrays {
		if declarations.shadowsModule(name) {
			delete(moduleArrays, name)
		}
	}
	idempotentSetupArrays := arrayModuleIdempotentSetupArrays(file, proc, moduleDecls, ctx)
	allocated := map[string]bool{}
	addDirectAllocation := func(statementID int, name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if !moduleArrays[name] || (!arrayProcedureBlockDominatesNormalExit(proc, statementID, dominators) && !idempotentSetupArrays[name]) {
			return
		}
		allocated[name] = true
	}
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct && !arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
					addDirectAllocation(statement.ID, redim.name)
				}
			}
		}
		if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed {
			name := strings.ToLower(cleanIdentifier(lhs))
			if moduleArrays[name] {
				if value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx); known && value.kind == arrayAllocated && value.knownArray {
					if !arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
						addDirectAllocation(statement.ID, name)
					}
				}
			}
		}
	}
	for call := range proc.Calls.All() {
		key, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
		if !ok {
			continue
		}
		calleeArrays := summaries[key]
		calleeByRefArrays := byRefSummaries[key]
		if len(calleeArrays) == 0 && len(calleeByRefArrays) == 0 {
			continue
		}
		guaranteed := !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) && arrayProcedureBlockDominatesNormalExit(proc, call.StatementID, dominators)
		if !guaranteed && arrayProcedureHasIdempotentSetupGuard(file, proc, call.Range.StartLine, moduleDecls) {
			guaranteed = true
		}
		if !guaranteed {
			continue
		}
		for name := range calleeArrays {
			if moduleArrays[name] {
				allocated[name] = true
			}
		}
		if len(calleeByRefArrays) > 0 {
			arguments, mapped := arrayCallFormalArguments(proc, target, call)
			if mapped {
				for index := range calleeByRefArrays {
					if index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
						continue
					}
					name := strings.ToLower(directArrayArgumentName(arguments[index]))
					if moduleArrays[name] {
						allocated[name] = true
					}
				}
			}
		}
	}
	return allocated
}

// arrayModuleIdempotentSetupArrays recognizes the narrow one-time module
// initialization idiom used by private helper routines:
//
//	If ready Then Exit Sub
//	ReDim values(...)
//	ready = True
//
// The direct ReDim does not dominate the procedure's normal exit because the
// already-initialized branch exits early.  The summary can nevertheless carry
// the allocation when the Boolean guard is module-scoped, is written only to
// True by this procedure, is not written elsewhere in the module, and is the
// final executable statement.  These constraints keep an arbitrary Boolean
// branch from becoming an allocation proof.
func arrayModuleIdempotentSetupArrays(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	if proc.StartLine < 1 || proc.EndLine < proc.StartLine || proc.StartLine > len(file.Lines) {
		return nil
	}
	end := min(len(file.Lines), proc.EndLine)
	start := max(0, proc.StartLine-1)
	facts := file.moduleAnalysisFacts()
	type setupGuard struct {
		name    string
		checkAt int
	}
	guards := make([]setupGuard, 0)
	for index := start; index < end; index++ {
		name, ok := facts.sourceLineSetupGuard(index)
		if !ok {
			continue
		}
		declaration, ok := moduleDecls[name]
		if !ok || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			continue
		}
		guards = append(guards, setupGuard{name: name, checkAt: index})
	}
	if len(guards) == 0 {
		return nil
	}
	constants := arrayIntegerConstants(file, proc, nil, nil)

	lastExecutable := -1
	for index := start; index < end; index++ {
		if facts.sourceLineIsExecutable(index) {
			lastExecutable = index
		}
	}
	if lastExecutable < start {
		return nil
	}

	guardWrites := map[string][]struct {
		line int
		rhs  string
	}{}
	for _, guard := range guards {
		facts.forEachArrayOperationFor(guard.name, func(operation moduleArrayOperationFact) {
			if operation.Kind != moduleArrayWholeAssignment {
				return
			}
			guardWrites[guard.name] = append(guardWrites[guard.name], struct {
				line int
				rhs  string
			}{line: operation.Line, rhs: operation.RHS})
		})
	}

	result := map[string]bool{}
	for _, guard := range guards {
		writes := guardWrites[guard.name]
		if len(writes) != 1 || writes[0].line != lastExecutable || !strings.EqualFold(writes[0].rhs, "true") {
			continue
		}
		setAt := writes[0].line
		for index := guard.checkAt + 1; index < setAt; index++ {
			facts.forEachArrayOperationAt(index, func(operation moduleArrayOperationFact) {
				if operation.Kind != moduleArrayDirectRedim || operation.Preserve {
					return
				}
				name := operation.Name
				declaration, declared := moduleDecls[name]
				if !declared || !declaration.Array || declaration.Parameter {
					return
				}
				if moduleArrayOperationHasOtherWrite(facts, name, index) {
					return
				}
				if !arrayModuleSetupReDimIsReliable(file, proc, guard.checkAt, index, setAt, guard.name, name, constants, ctx, moduleDecls) {
					return
				}
				result[name] = true
			})
		}
	}
	if len(result) == 0 {
		return nil
	}
	return result
}

// inferArrayModuleReadyGuardStates recognizes the module-level lifecycle
// invariant used by consumers such as CSV readers:
//
//	moduleArray = Split(...)
//	ready = True
//
// and later:
//
//	If Not ready Then Exit Function
//	use moduleArray(...)
//
// The existing module allocation summary is caller-oriented and therefore
// cannot establish the state at an independently callable public procedure.
// This summary is deliberately source-owned and fail-closed so an arbitrary
// Boolean assignment never becomes an array allocation proof.
func inferArrayModuleReadyGuardStates(files []parsedFile, ctx analysisContext) arrayModuleReadyGuardStates {
	states := arrayModuleReadyGuardStates{}
	for _, file := range files {
		facts := file.moduleAnalysisFacts()
		if facts == nil {
			continue
		}
		moduleDecls := file.moduleDecls()
		for guardName, guardDeclaration := range moduleDecls {
			guardName = strings.ToLower(cleanIdentifier(guardName))
			if guardName == "" || guardDeclaration.Array || guardDeclaration.Parameter || !strings.EqualFold(strings.TrimSpace(guardDeclaration.Type), "Boolean") || !arrayModuleReadyGuardSourceOwned(file, guardDeclaration) {
				continue
			}

			var writes []moduleArrayOperationFact
			valid := true
			trueWrites := make([]moduleArrayOperationFact, 0, 1)
			facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
				if !valid {
					return
				}
				owner, owned := arrayModuleProcedureAtLine(file, operation.Line+1)
				if !owned {
					valid = false
					return
				}
				scope := newDeclarationScope(file, owner)
				scope.module = moduleDecls
				if scope.shadowsModule(guardName) {
					return
				}
				if operation.Kind != moduleArrayWholeAssignment {
					valid = false
					return
				}
				rhs := strings.ToLower(strings.TrimSpace(operation.RHS))
				if rhs != "true" && rhs != "false" {
					valid = false
					return
				}
				writes = append(writes, operation)
				if rhs == "true" {
					trueWrites = append(trueWrites, operation)
				}
			})
			if !valid || len(trueWrites) != 1 {
				continue
			}
			writer, ok := arrayModuleProcedureAtLine(file, trueWrites[0].Line+1)
			if !ok {
				continue
			}
			allocated := arrayModuleReadyGuardAllocationProof(file, writer, trueWrites[0].Line+1, moduleDecls, ctx)
			safeAllocated := map[string]bool{}
			for name := range allocated {
				if arrayModuleReadyGuardLifecycleSafe(file, guardName, map[string]bool{name: true}, facts, moduleDecls, ctx) {
					safeAllocated[name] = true
				}
			}
			if len(safeAllocated) == 0 {
				continue
			}
			if states[file.Path] == nil {
				states[file.Path] = map[string]map[string]bool{}
			}
			states[file.Path][guardName] = safeAllocated
		}
	}
	return states
}

func arrayModuleReadyGuardSourceOwned(file parsedFile, declaration sourceDeclaration) bool {
	if declaration.Line < 1 || declaration.Line > len(file.Lines) {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(file.Lines[declaration.Line-1])))
	return !strings.HasPrefix(lower, "public ") && !strings.HasPrefix(lower, "global ")
}

func arrayModuleProcedureAtLine(file parsedFile, line int) (sourceProcedure, bool) {
	if line < 1 {
		return sourceProcedure{}, false
	}
	procedures := file.procedureView()
	for index := 0; index < procedures.Len(); index++ {
		procedure := procedures.valueAt(index)
		if line >= procedure.StartLine && line <= procedure.EndLine {
			return procedure, true
		}
	}
	return sourceProcedure{}, false
}

func arrayModuleReadyGuardAllocationProof(file parsedFile, proc sourceProcedure, readyLine int, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	if proc.Graph == nil || readyLine < 1 || readyLine > len(file.Lines) {
		return nil
	}
	variables := arrayVariables(file, proc, moduleDecls)
	candidates := map[string]bool{}
	for name, declaration := range moduleDecls {
		name = strings.ToLower(cleanIdentifier(name))
		if name != "" && declaration.Array && !declaration.Parameter && arrayModuleReadyGuardSourceOwned(file, declaration) {
			if variable, known := variables[name]; known && variable.isArray {
				candidates[name] = true
			}
		}
	}
	if len(candidates) == 0 {
		return nil
	}

	graph := arrayVBA227Graph(proc, ctx)
	initial := arrayInitialState(variables)
	seenReady := false
	failed := map[string]bool{}
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		if line == readyLine {
			seenReady = true
			for name := range candidates {
				value, known := in[name]
				if !known || value.kind != arrayAllocated || !value.knownArray {
					failed[name] = true
				}
			}
		}
		out, _ := (Analyzer{}).arrayVBA227Transfer(file, proc, ctx, variables, in, text, line, nil, nil, nil)
		for _, call := range arrayCallsAtLine(proc.Calls, line) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
		}
		return out
	}
	edgeState := func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
		out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
		return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, arrayAllocationTransferIsReliable, nil)
	if !seenReady {
		return nil
	}
	result := map[string]bool{}
	for name := range candidates {
		if !failed[name] {
			result[name] = true
		}
	}
	return result
}

func arrayModuleReadyGuardLifecycleSafe(file parsedFile, guardName string, arrays map[string]bool, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration, ctx analysisContext) bool {
	if facts == nil {
		return false
	}
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		safe := true
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if !safe {
				return
			}
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
			if !ok {
				safe = false
				return
			}
			scope := newDeclarationScope(file, owner)
			scope.module = moduleDecls
			if scope.shadowsModule(name) {
				return
			}
			line := operation.Line + 1
			if line < 1 || line > len(file.Lines) {
				safe = false
				return
			}
			switch operation.Kind {
			case moduleArrayWholeAssignment:
				lhs, rhs, indexed, assigned := arrayAssignment(arrayLogicalCodeLine(file.Lines, line))
				if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), name) {
					safe = false
					return
				}
				value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx)
				if !known || value.kind != arrayAllocated || !value.knownArray {
					safe = false
				}
			case moduleArrayDirectRedim:
				if arraySummaryStatementAlwaysFails(arrayLogicalCodeLine(file.Lines, line), arrayOptionBase(file), arrayIntegerConstants(file, owner, nil, nil)) {
					safe = false
				}
			case moduleArrayErase:
				if !arrayModuleReadyGuardFalseWriteDominates(file, owner, guardName, line, facts, moduleDecls) {
					safe = false
				}
			default:
				safe = false
			}
		})
		if !safe {
			return false
		}
	}
	return true
}

func arrayModuleReadyGuardFalseWriteDominates(file parsedFile, proc sourceProcedure, guardName string, eraseLine int, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration) bool {
	if proc.Graph == nil || facts == nil || eraseLine <= proc.StartLine {
		return false
	}
	falseLines := make([]int, 0, 1)
	facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
		if operation.Kind != moduleArrayWholeAssignment || !strings.EqualFold(strings.TrimSpace(operation.RHS), "false") {
			return
		}
		line := operation.Line + 1
		owner, ok := arrayModuleProcedureAtLine(file, line)
		if !ok || owner.StartByte != proc.StartByte || owner.StartLine != proc.StartLine || owner.EndLine != proc.EndLine {
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if scope.shadowsModule(guardName) {
			return
		}
		if line < eraseLine {
			falseLines = append(falseLines, line)
		}
	})
	if len(falseLines) == 0 {
		return false
	}
	eraseStatement, eraseOK := arrayModuleStatementAtLine(proc, eraseLine)
	if !eraseOK {
		return false
	}
	normalGraph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	eraseBlock, eraseBlockOK := normalGraph.BlockForStatement(eraseStatement.ID)
	if !eraseBlockOK || !normalGraph.IsReachable(eraseBlock.ID) {
		return false
	}
	for _, falseLine := range falseLines {
		if arrayModuleReadyGuardHasWriteBetween(file, proc, guardName, falseLine, eraseLine, facts, moduleDecls) {
			continue
		}
		falseStatement, falseOK := arrayModuleStatementAtLine(proc, falseLine)
		if !falseOK {
			continue
		}
		falseBlock, falseBlockOK := normalGraph.BlockForStatement(falseStatement.ID)
		if !falseBlockOK {
			continue
		}
		for _, dominator := range normalGraph.DominatorsOf(eraseBlock.ID) {
			if dominator == falseBlock.ID {
				return true
			}
		}
	}
	return false
}

func arrayModuleReadyGuardHasWriteBetween(file parsedFile, proc sourceProcedure, guardName string, startLine, endLine int, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration) bool {
	written := false
	facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
		if written || operation.Kind != moduleArrayWholeAssignment {
			return
		}
		line := operation.Line + 1
		if line <= startLine || line >= endLine {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, line)
		if !ok || owner.StartByte != proc.StartByte || owner.StartLine != proc.StartLine || owner.EndLine != proc.EndLine {
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if !scope.shadowsModule(guardName) {
			written = true
		}
	})
	return written
}

func arrayModuleStatementAtLine(proc sourceProcedure, line int) (procedureir.Statement, bool) {
	var found procedureir.Statement
	matched := false
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine != line {
			continue
		}
		if matched {
			return procedureir.Statement{}, false
		}
		found = statement
		matched = true
	}
	return found, matched
}

func arrayModuleSetupReDimIsReliable(file parsedFile, proc sourceProcedure, guardLine, redimLine, setLine int, guardName, name string, constants map[string]int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) bool {
	if redimLine <= guardLine || redimLine >= setLine || redimLine < 0 || setLine >= len(file.Lines) {
		return false
	}
	sourceLine := redimLine + 1
	if sourceLine < 1 || sourceLine > len(file.Lines) || arraySummaryStatementAlwaysFails(normalizedCodeLine(file.Lines[redimLine]), arrayOptionBase(file), constants) {
		return false
	}
	variables := arrayVariables(file, proc, moduleDecls)
	// A call between the allocation and the ready flag can erase or replace the
	// module array without leaving a direct operation fact in this procedure.
	// A resolved private helper is admitted only when its direct and ByRef array
	// effects are known not to invalidate this module's arrays; public,
	// unresolved, and otherwise unmodelled calls remain conservative.
	for call := range proc.Calls.All() {
		line := call.Range.StartLine
		if line >= sourceLine && line < setLine+1 {
			if arrayCallIsIndexedArrayAccess(proc, call, variables) {
				continue
			}
			if call.IsRaiseEvent {
				return false
			}
			if call.Resolution.Status == procedureir.ResolutionBuiltinLike {
				continue
			}
			_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			if !resolved || arrayPrivateCallMayInvalidateModuleArray(file, proc, target, call, moduleDecls, ctx) {
				return false
			}
		}
	}

	if proc.Graph == nil {
		// Compatibility projections may not carry CFGs. Accept only a straight-
		// line source interval in that case; any visible control construct makes
		// the allocation conditional or permits an unmodelled bypass.
		for line := guardLine + 1; line < setLine; line++ {
			if line == redimLine {
				continue
			}
			if arrayModuleSetupLineHasControlFlow(file.Lines[line]) {
				return false
			}
		}
		return true
	}

	findStatement := func(line int, match func(procedureir.Statement) bool) (procedureir.Statement, bool) {
		var found procedureir.Statement
		matched := false
		for statement := range proc.Statements.All() {
			if statement.Range.StartLine != line || !match(statement) {
				continue
			}
			if matched {
				return procedureir.Statement{}, false
			}
			found = statement
			matched = true
		}
		return found, matched
	}
	guardStatement, guardOK := findStatement(guardLine+1, func(statement procedureir.Statement) bool {
		return len(arraySetupGuardRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))) == 2
	})
	redimStatement, redimOK := findStatement(sourceLine, func(statement procedureir.Statement) bool {
		match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
		if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
			return false
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), cleanIdentifier(name)) {
				return true
			}
		}
		return false
	})
	readyStatement, readyOK := findStatement(setLine+1, func(statement procedureir.Statement) bool {
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(statement.Text))
		return ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(guardName)) && strings.EqualFold(strings.TrimSpace(rhs), "true")
	})
	if !guardOK || !redimOK || !readyOK {
		return false
	}
	guardBlock, guardBlockOK := proc.Graph.BlockForStatement(guardStatement.ID)
	redimBlock, redimBlockOK := proc.Graph.BlockForStatement(redimStatement.ID)
	readyBlock, readyBlockOK := proc.Graph.BlockForStatement(readyStatement.ID)
	if !guardBlockOK || !redimBlockOK || !readyBlockOK {
		return false
	}
	normalGraph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	if !normalGraph.IsReachable(redimBlock.ID) || !normalGraph.IsReachable(readyBlock.ID) {
		return false
	}
	dominatesReady := false
	for _, dominator := range normalGraph.DominatorsOf(readyBlock.ID) {
		if dominator == redimBlock.ID {
			dominatesReady = true
			break
		}
	}
	if !dominatesReady || !arrayFalseBranchRequiresBlock(*proc.Graph, guardBlock.ID, redimBlock.ID) {
		return false
	}
	return arrayModuleSetupReachesNormalExit(normalGraph, readyBlock.ID)
}

// arrayCallIsIndexedArrayAccess filters the procedure IR's expression-call
// projection from actual procedure invocations. The VBA tree-sitter grammar
// represents an indexed array expression such as mPow2(index) as a CallSite;
// an unresolved array expression must not make an otherwise straight-line
// setup helper look like it contains an unknown side effect.
func arrayCallIsIndexedArrayAccess(proc sourceProcedure, call procedureir.CallSite, variables map[string]arrayVariable) bool {
	name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
	if name == "" || strings.Contains(name, ".") || call.StatementID <= 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		if statement.ID != call.StatementID {
			continue
		}
		for _, use := range arrayIndexedUses(statement.Text, variables) {
			if strings.EqualFold(cleanIdentifier(use.name), name) {
				return true
			}
		}
		return false
	}
	return false
}

func arrayModuleSetupLineHasControlFlow(line string) bool {
	code := strings.TrimSpace(normalizedCodeLine(line))
	if code == "" {
		return false
	}
	lower := strings.ToLower(code)
	for _, prefix := range []string{
		"if ", "elseif ", "else", "end if", "for ", "for each ", "next", "do", "loop", "while ", "wend",
		"select ", "case ", "goto ", "on error ", "with ", "end with", "exit ",
	} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return strings.HasSuffix(code, ":")
}

func arrayModuleSetupReachesNormalExit(graph vbacfg.CFGView, from vbacfg.BlockID) bool {
	if from == graph.NormalExit() {
		return true
	}
	seen := map[vbacfg.BlockID]bool{from: true}
	queue := []vbacfg.BlockID{from}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		found := false
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if edge.To == graph.NormalExit() {
				found = true
				return false
			}
			if !seen[edge.To] {
				seen[edge.To] = true
				queue = append(queue, edge.To)
			}
			return true
		})
		if found {
			return true
		}
	}
	return false
}

func moduleArrayOperationHasOtherWrite(facts *moduleAnalysisFacts, name string, setupLine int) bool {
	name = strings.ToLower(cleanIdentifier(name))
	otherWrite := false
	facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
		if otherWrite {
			return
		}
		if operation.Kind == moduleArrayDirectRedim && operation.Line == setupLine {
			return
		}
		otherWrite = true
	})
	return otherWrite
}

func arrayProcedureLineHasInlineConditional(file parsedFile, line int) bool {
	if line < 1 || line > len(file.Lines) {
		return false
	}
	text := strings.ToLower(strings.TrimSpace(normalizedCodeLine(file.Lines[line-1])))
	return strings.HasPrefix(text, "if ") && strings.Contains(text, " then ")
}

func arrayProcedureLineInlineConditionIsFalse(file parsedFile, line int) bool {
	if !arrayProcedureLineHasInlineConditional(file, line) {
		return false
	}
	condition, body, ok := arrayIfThenParts(normalizedCodeLine(file.Lines[line-1]))
	if !ok || strings.TrimSpace(body) == "" {
		return false
	}
	condition = strings.TrimSpace(condition)
	lowerCondition := strings.ToLower(condition)
	if strings.HasPrefix(lowerCondition, "if ") {
		condition = strings.TrimSpace(condition[len("if "):])
	}
	value, known := arraySourceOrderConstantBoolean(condition, nil)
	return known && !value
}

func arrayInlineConditionalCallIsReachable(file parsedFile, call procedureir.CallSite, conditionValue, hasElse bool) bool {
	if conditionValue && !hasElse {
		return true
	}
	line := call.Range.StartLine
	if line < 1 || line > len(file.Lines) || call.Range.StartColumn <= 0 {
		return conditionValue || hasElse
	}
	raw := gui.StripComment(file.Lines[line-1])
	trimmed := strings.TrimSpace(raw)
	leading := len(raw) - len(strings.TrimLeft(raw, " \t"))
	_, body, ok := arrayIfThenParts(trimmed)
	if !ok || strings.TrimSpace(body) == "" {
		return true
	}
	prefixLength := 0
	switch {
	case strings.HasPrefix(strings.ToLower(trimmed), "if "):
		prefixLength = len("if ")
	case strings.HasPrefix(strings.ToLower(trimmed), "elseif "):
		prefixLength = len("elseif ")
	default:
		return true
	}
	rest := strings.TrimSpace(trimmed[prefixLength:])
	thenIndex := arrayTopLevelKeywordIndex(rest, "then")
	if thenIndex < 0 {
		return true
	}
	bodyStart := leading + prefixLength + (len(trimmed[prefixLength:]) - len(rest)) + thenIndex + len("then")
	elseIndex := arrayTopLevelKeywordIndex(body, "else")
	if elseIndex < 0 {
		return conditionValue
	}
	elseStart := bodyStart + elseIndex
	callColumn := call.Range.StartColumn - 1
	if callColumn < bodyStart {
		return true
	}
	inThen := callColumn < elseStart
	return inThen == conditionValue
}

func arrayProcedureNormalExitDominators(proc sourceProcedure) map[vbacfg.BlockID]bool {
	if proc.Graph == nil {
		return nil
	}
	dominators := proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[proc.Graph.NormalExit]
	result := make(map[vbacfg.BlockID]bool, len(dominators))
	for _, id := range dominators {
		result[id] = true
	}
	return result
}

func arrayProcedureBlockDominatesNormalExit(proc sourceProcedure, statementID int, dominators map[vbacfg.BlockID]bool) bool {
	if proc.Graph == nil || statementID <= 0 || len(dominators) == 0 {
		return false
	}
	block, ok := proc.Graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	return dominators[block.ID]
}

func arrayProcedureHasIdempotentSetupGuard(file parsedFile, proc sourceProcedure, candidateLine int, moduleDecls map[string]sourceDeclaration) bool {
	if candidateLine <= proc.StartLine {
		return false
	}
	guard := ""
	for line := proc.StartLine; line < candidateLine && line <= len(file.Lines); line++ {
		match := arraySetupGuardRe.FindStringSubmatch(normalizedCodeLine(file.Lines[line-1]))
		if len(match) != 2 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		declaration, ok := moduleDecls[name]
		if !ok || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			continue
		}
		guard = name
	}
	if guard == "" {
		return false
	}
	for line := candidateLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(file.Lines[line-1]))
		if ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), guard) && strings.EqualFold(strings.TrimSpace(rhs), "true") {
			return true
		}
	}
	return false
}

func arrayModuleAllocationSummariesEqual(left, right arrayModuleAllocationSummaries) bool {
	if len(left) != len(right) {
		return false
	}
	for procedure, arrays := range left {
		other, ok := right[procedure]
		if !ok || len(arrays) != len(other) {
			return false
		}
		for name := range arrays {
			if !other[name] {
				return false
			}
		}
	}
	return true
}

func arrayModuleInitializationStates(files []parsedFile, summaries arrayModuleAllocationSummaries) map[string]map[string]bool {
	states := map[string]map[string]bool{}
	for _, file := range files {
		moduleKind := strings.ToLower(strings.TrimSpace(file.ModuleKind))
		if moduleKind != "form" && moduleKind != "class" {
			continue
		}
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		initializer := arrayModuleInitializerName(moduleKind)
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !strings.EqualFold(strings.TrimSpace(proc.Name), initializer) {
				continue
			}
			for name := range summaries[arrayProcedureKey(proc)] {
				if declaration, ok := moduleDecls[name]; ok && declaration.Array && !declaration.Parameter {
					if states[file.Path] == nil {
						states[file.Path] = map[string]bool{}
					}
					states[file.Path][name] = true
				}
			}
		}
	}
	return states
}

func applyArrayModuleInitializationState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, initializationStates map[string]map[string]bool) arrayFlowState {
	if len(initializationStates[file.Path]) == 0 {
		return state
	}
	moduleKind := strings.ToLower(strings.TrimSpace(file.ModuleKind))
	if moduleKind != "form" && moduleKind != "class" {
		return state
	}
	initializer := arrayModuleInitializerName(moduleKind)
	if strings.EqualFold(strings.TrimSpace(proc.Name), initializer) {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range initializationStates[file.Path] {
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, ok := moduleDecls[name]
		variable, known := variables[name]
		if !ok || !known || !declaration.Array || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayModuleInitializerName(moduleKind string) string {
	if moduleKind == "form" {
		return "userform" + "_initialize"
	}
	return "class" + "_initialize"
}

// inferArrayModuleEntryStates propagates a module-array allocation from a
// known caller to a project-local helper. Private procedures are analyzed as
// standalone procedures, so their initial state cannot otherwise reflect an
// assignment performed by the caller. A fact is retained only when every
// resolved call from the same module reaches the helper with that array
// allocated. The fixed point also covers chains of private helpers.
func inferArrayModuleEntryStates(a Analyzer, files []parsedFile, ctx analysisContext) arrayModuleEntryStates {
	if len(ctx.arrayPrivateTargets) == 0 {
		return arrayModuleEntryStates{}
	}

	type procedureInfo struct {
		file        parsedFile
		proc        sourceProcedure
		key         string
		moduleDecls map[string]sourceDeclaration
		variables   map[string]arrayVariable
	}
	procedures := make([]procedureInfo, 0)
	moduleArrays := map[string]map[string]bool{}
	moduleFiles := map[string]string{}
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			key := arrayParticipantLookupKey(proc, ctx.arrayParticipantKeys)
			procedures = append(procedures, procedureInfo{
				file: file, proc: proc, moduleDecls: moduleDecls,
				variables: arrayVariables(file, proc, moduleDecls),
				key:       key,
			})
			moduleArrays[key] = arrayModuleNamesForProcedure(file, proc, moduleDecls)
			moduleFiles[key] = file.Path
		}
	}
	if len(procedures) == 0 {
		return arrayModuleEntryStates{}
	}

	initializationStates := arrayModuleInitializationStates(files, ctx.arrayModuleAllocations)
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	indexByKey := make(map[string]int, len(procedures))
	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		key := procedure.key
		indexByKey[key] = index
		for call := range procedure.proc.Calls.All() {
			_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			participantTargetKey := arrayParticipantLookupKey(target, ctx.arrayParticipantKeys)
			if ok && moduleFiles[participantTargetKey] == procedure.file.Path {
				dependents[participantTargetKey] = append(dependents[participantTargetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}

	evaluate := func(procedure procedureInfo, entries arrayModuleEntryStates) map[string]map[string]bool {
		variables := procedure.variables
		initial := arrayInitialState(variables)
		initial = applyArrayModuleInitializationState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, initializationStates)
		initial = applyArrayModuleReadyGuardState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, ctx.arrayModuleReadyGuards)
		initial = applyArrayModuleEntryState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, entries, ctx.arrayParticipantKeys)
		initial = applyArrayInternalStorageConfiguration(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, ctx.arrayModuleConfigurations[procedure.file.Path])
		candidates := map[string]map[string]bool{}
		recordCall := func(call procedureir.CallSite, state arrayFlowState) {
			_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			key := arrayParticipantLookupKey(target, ctx.arrayParticipantKeys)
			if !ok || moduleFiles[key] != procedure.file.Path {
				return
			}
			names := moduleArrays[key]
			if len(names) == 0 {
				return
			}
			candidate := candidates[key]
			if candidate == nil {
				candidate = cloneArrayNameSet(names)
				candidates[key] = candidate
			}
			for name := range names {
				value, known := state[name]
				if !known || value.kind != arrayAllocated || !value.knownArray {
					candidate[name] = false
				}
			}
		}
		visit := func(text string, line int, in arrayFlowState) arrayFlowState {
			for _, call := range arrayCallsAtLine(procedure.proc.Calls, line) {
				recordCall(call, in)
			}
			out, _ := a.arrayTransfer(procedure.file, procedure.proc, ctx, variables, in, text, line, nil, nil)
			for _, call := range arrayCallsAtLine(procedure.proc.Calls, line) {
				out = applyArrayModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
				out = applyArrayUnknownModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
			}
			return out
		}
		if procedure.proc.Graph == nil {
			state := initial
			for line := procedure.proc.StartLine; line <= procedure.proc.EndLine && line <= len(procedure.file.Lines); line++ {
				state = visit(normalizedCodeLine(procedure.file.Lines[line-1]), line, state)
			}
			return candidates
		}
		graph := arrayVBA227Graph(procedure.proc, ctx)
		if ctx.arrayStats != nil {
			ctx.arrayStats.addCFGWalk()
		}
		walkArrayCFGWithEdgesStats(&graph, procedure.file.Lines, initial, visit, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
			out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[procedure.file.Path], variables, procedure.file, procedure.proc, procedure.moduleDecls)
		}, ctx.arrayStats)
		return candidates
	}

	contributions := make(map[string]map[string]map[string]bool, len(procedures))
	entries := arrayModuleEntryStates{}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		key := procedures[index].key
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		contribution := evaluate(procedures[index], entries)
		if arrayModuleEntryContributionsEqual(contributions[key], contribution) {
			continue
		}
		contributions[key] = contribution
		next := arrayModuleEntryStates{}
		for _, caller := range procedures {
			for target, names := range contributions[caller.key] {
				if next[target] == nil {
					next[target] = cloneArrayNameSet(names)
					continue
				}
				for name := range next[target] {
					if !names[name] {
						delete(next[target], name)
					}
				}
			}
		}
		changedTargetSet := make(map[string]bool, len(entries)+len(next))
		for target := range entries {
			changedTargetSet[target] = true
		}
		for target := range next {
			changedTargetSet[target] = true
		}
		changedTargets := make([]string, 0, len(changedTargetSet))
		for target := range changedTargetSet {
			if !arrayModuleEntryTargetEqual(entries, next, target) {
				changedTargets = append(changedTargets, target)
			}
		}
		if len(changedTargets) == 0 {
			continue
		}
		entries = next
		sortArrayProcedureKeys(changedTargets, indexByKey)
		for _, target := range changedTargets {
			for _, dependent := range dependents[target] {
				if !queued[dependent] {
					queued[dependent] = true
					queue = append(queue, dependent)
				}
			}
			if index, ok := indexByKey[target]; ok && !queued[index] {
				queued[index] = true
				queue = append(queue, index)
			}
		}
	}
	return entries
}

func arrayModuleEntryContributionsEqual(left, right map[string]map[string]bool) bool {
	if len(left) != len(right) {
		return false
	}
	for key, names := range left {
		other, ok := right[key]
		if !ok || len(names) != len(other) {
			return false
		}
		for name, allocated := range names {
			if other[name] != allocated {
				return false
			}
		}
	}
	return true
}

func arrayModuleEntryTargetEqual(left, right arrayModuleEntryStates, target string) bool {
	leftNames, leftOK := left[target]
	rightNames, rightOK := right[target]
	if !leftOK || !rightOK {
		return !leftOK && !rightOK
	}
	if len(leftNames) != len(rightNames) {
		return false
	}
	for name, allocated := range leftNames {
		if rightNames[name] != allocated {
			return false
		}
	}
	return true
}

func arrayModuleNamesForProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]bool {
	moduleArrays := map[string]bool{}
	for name, declaration := range moduleDecls {
		if declaration.Array && !declaration.Parameter {
			moduleArrays[strings.ToLower(name)] = true
		}
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	for name := range moduleArrays {
		if declarations.shadowsModule(name) {
			delete(moduleArrays, name)
		}
	}
	return moduleArrays
}

func applyArrayModuleEntryState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, entries arrayModuleEntryStates, participantKeys ...map[string]string) arrayFlowState {
	var keyIndex map[string]string
	if len(participantKeys) > 0 {
		keyIndex = participantKeys[0]
	}
	allocated := entries[arrayParticipantLookupKey(proc, keyIndex)]
	if len(allocated) == 0 && len(keyIndex) > 0 {
		// Focused compatibility callers may still provide the legacy base key.
		allocated = entries[arrayProcedureKey(proc)]
	}
	if len(allocated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name, isAllocated := range allocated {
		if !isAllocated {
			continue
		}
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || declaration.Parameter || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func applyArrayModuleReadyGuardState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, guards arrayModuleReadyGuardStates) arrayFlowState {
	byGuard := guards[file.Path]
	if len(byGuard) == 0 {
		return state
	}
	guardName, ok := arrayModuleReadyGuardAtEntry(file, proc, moduleDecls)
	if !ok {
		return state
	}
	allocated := byGuard[guardName]
	if len(allocated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range allocated {
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || declaration.Parameter || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayModuleReadyGuardAtEntry(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) (string, bool) {
	if proc.StartLine < 1 || proc.EndLine <= proc.StartLine || proc.StartLine > len(file.Lines) {
		return "", false
	}
	end := min(len(file.Lines), proc.EndLine)
	for line := proc.StartLine + 1; line < end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" || strings.HasPrefix(text, "'") {
			continue
		}
		if declRe.MatchString(text) || strings.HasPrefix(strings.ToLower(text), "on error ") || strings.HasSuffix(text, ":") {
			continue
		}
		match := arrayModuleReadyGuardRe.FindStringSubmatch(text)
		if len(match) != 2 {
			return "", false
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		declaration, declared := moduleDecls[name]
		if !declared || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			return "", false
		}
		return name, true
	}
	return "", false
}

type arrayByRefEntryEvidence struct {
	seen                bool
	allocated           bool
	conditionCompatible bool
	condition           string
}

type arrayByRefCallCandidate struct {
	key    string
	target sourceProcedure
	call   procedureir.CallSite
}

type arrayLocalGoSubSummary struct {
	guaranteedAllocated map[string]bool
	unknown             map[string]bool
}

type arrayLocalGoSubAllocations map[string]arrayLocalGoSubSummary

func arrayLocalGoSubAllocationSummaries(proc sourceProcedure, graph *vbacfg.CFGView, variables map[string]arrayVariable, ctx analysisContext, base int, constants map[string]int, moduleArrays map[string]bool) arrayLocalGoSubAllocations {
	statements := make([]procedureir.Statement, 0, proc.Statements.Len())
	for statement := range proc.Statements.All() {
		statements = append(statements, statement)
	}
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartLine != statements[j].Range.StartLine {
			return statements[i].Range.StartLine < statements[j].Range.StartLine
		}
		return statements[i].ID < statements[j].ID
	})
	summaries := arrayLocalGoSubAllocations{}
	if graph == nil {
		return summaries
	}
	for index, label := range statements {
		if label.Kind != procedureir.StatementLabel {
			continue
		}
		labelName := arrayLocalGoSubLabelName(label)
		if labelName == "" {
			continue
		}
		end := len(statements)
		for cursor := index + 1; cursor < len(statements); cursor++ {
			if arrayLocalGoSubIsReturn(statements[cursor]) {
				end = cursor
				break
			}
		}
		if end == len(statements) {
			continue
		}
		summary := arrayLocalGoSubSummary{
			guaranteedAllocated: map[string]bool{},
			unknown:             map[string]bool{},
		}
		for name, variable := range variables {
			if variable.isArray && arrayLocalGoSubAllocationInvariant(proc, graph, statements, index, end, name, ctx, base, constants, moduleArrays) {
				summary.guaranteedAllocated[name] = true
			} else if variable.isArray && !variable.fixed && arrayLocalGoSubMayMutateName(proc, graph, statements, index, end, name, ctx, moduleArrays) {
				summary.unknown[name] = true
			}
		}
		summaries[labelName] = summary
	}
	return summaries
}

func arrayLocalGoSubAllocationInvariant(proc sourceProcedure, graph *vbacfg.CFGView, statements []procedureir.Statement, labelIndex, end int, name string, ctx analysisContext, base int, constants map[string]int, moduleArrays map[string]bool) bool {
	if graph == nil || labelIndex < 0 || labelIndex >= end || end > len(statements) {
		return false
	}
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	allowed := map[vbacfg.BlockID]procedureir.Statement{}
	var labelBlock vbacfg.Block
	labelFound := false
	for index := labelIndex; index < end; index++ {
		block, ok := graph.BlockForStatement(statements[index].ID)
		if !ok {
			return false
		}
		allowed[block.ID] = statements[index]
		if index == labelIndex {
			labelBlock = block
			labelFound = true
		}
	}
	if !labelFound {
		return false
	}
	returnBlock, returnOK := graph.BlockForStatement(statements[end].ID)
	if !returnOK {
		return false
	}

	type stateAtBlock struct {
		id        vbacfg.BlockID
		allocated bool
	}
	seenStates := map[vbacfg.BlockID]map[bool]bool{labelBlock.ID: {false: true}}
	queue := []stateAtBlock{{id: labelBlock.ID}}
	for len(queue) > 0 {
		currentState := queue[0]
		queue = queue[1:]
		currentID := currentState.id
		if _, ok := graph.BlockByID(currentID); !ok {
			return false
		}
		failed := false
		graph.ForEachOutgoing(currentID, func(edge vbacfg.Edge) bool {
			target, targetOK := graph.BlockByID(edge.To)
			if !targetOK {
				failed = true
				return true
			}
			// An unknown edge can leave the GoSub body through a dynamic or
			// recovered transfer. It is not evidence of a successful Return,
			// even when the predecessor happens to be allocated.
			if edge.Kind == vbacfg.EdgeUnknown || target.Kind == vbacfg.BlockUnknownExit {
				failed = true
				return true
			}
			statement, inside := allowed[target.ID]
			if !inside {
				// The first Return is deliberately outside the allowed body. Any
				// other statement target leaves the GoSub body without returning to
				// its caller, so it cannot be treated as a successful summary.
				if target.ID != returnBlock.ID && !arrayLocalGoSubIsTerminalBlock(target) {
					failed = true
					return true
				}
				if !currentState.allocated {
					failed = true
				}
				// Keep visiting sibling edges. A terminal edge is safe only when
				// every other edge from this block also preserves the summary.
				return true
			}
			nextAllocated := arrayLocalGoSubStateAfterStatement(statement.Text, name, currentState.allocated, ctx, base, constants)
			unknownCall, guaranteedCall := arrayLocalGoSubArrayCallEffect(proc, statement, name, ctx, moduleArrays)
			if unknownCall {
				nextAllocated = false
			} else if guaranteedCall {
				nextAllocated = true
			}
			if !seenStates[target.ID][nextAllocated] {
				if seenStates[target.ID] == nil {
					seenStates[target.ID] = map[bool]bool{}
				}
				seenStates[target.ID][nextAllocated] = true
				queue = append(queue, stateAtBlock{id: target.ID, allocated: nextAllocated})
			}
			return true
		})
		if failed {
			return false
		}
	}
	return true
}

func arrayLocalGoSubStateAfterStatement(text, name string, allocated bool, ctx analysisContext, base int, constants map[string]int) bool {
	names := map[string]bool{name: true}
	for _, statement := range splitRangeValueSourceStatements(text) {
		statement = strings.TrimSpace(statement)
		if statement == "" {
			continue
		}
		if arrayLocalGoSubAllocationIsReliable(statement, name, ctx, base, constants) {
			allocated = true
			continue
		}
		if arrayLocalGoSubPreserveKeepsAllocation(statement, name, allocated, base, constants) {
			continue
		}
		if arraySourceOrderInlineArrayMutation(statement, names, ctx) || arraySourceOrderMutatesArrayStatement(statement, names, ctx) {
			allocated = false
		}
	}
	return allocated
}

func arrayLocalGoSubPreserveKeepsAllocation(text, name string, allocated bool, base int, constants map[string]int) bool {
	if !allocated {
		return false
	}
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 || strings.TrimSpace(match[1]) == "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if direct && strings.EqualFold(cleanIdentifier(redim.name), name) && !arraySummaryStatementAlwaysFails(text, base, constants) {
			return true
		}
	}
	return false
}

func arrayLocalGoSubMayMutateName(proc sourceProcedure, graph *vbacfg.CFGView, statements []procedureir.Statement, labelIndex, end int, name string, ctx analysisContext, moduleArrays map[string]bool) bool {
	if graph == nil || labelIndex < 0 || labelIndex >= end || end > len(statements) {
		return true
	}
	for index := labelIndex; index < end; index++ {
		statement := statements[index]
		for _, part := range splitRangeValueSourceStatements(statement.Text) {
			part = strings.TrimSpace(part)
			if part == "" {
				continue
			}
			if arrayStatementAllocatesName(part, name, ctx) || arraySourceOrderInlineArrayMutation(part, map[string]bool{name: true}, ctx) || arraySourceOrderMutatesArrayStatement(part, map[string]bool{name: true}, ctx) {
				return true
			}
		}
		unknownCall, _ := arrayLocalGoSubArrayCallEffect(proc, statement, name, ctx, moduleArrays)
		if unknownCall {
			return true
		}
	}
	return false
}

// arrayLocalGoSubArrayCallEffect keeps a local GoSub summary fail-closed at
// calls that pass the tracked array directly. Only an already-proven private
// ByRef allocation contract is allowed to establish the post-call state;
// builtin-like calls are treated as non-mutating, while every other call may
// erase or otherwise invalidate the array behind the analyzer's view.
func arrayLocalGoSubArrayCallEffect(proc sourceProcedure, statement procedureir.Statement, name string, ctx analysisContext, moduleArrays map[string]bool) (unknown, guaranteedAllocated bool) {
	name = strings.ToLower(cleanIdentifier(name))
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
			continue
		}
		if call.StatementID != statement.ID && (call.StatementID != 0 || call.Range.StartLine != statement.Range.StartLine) {
			continue
		}
		arguments := arrayCallArgumentTexts(proc, call)
		relevantArgument := false
		key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if resolved {
			bindings, mapped := arrayCallArgumentBindings(proc, target, call)
			if mapped {
				for _, binding := range bindings {
					if directArrayArgumentName(binding.text) != name {
						continue
					}
					if binding.parameterIndex >= 0 && binding.parameterIndex < target.Params.Len() && target.Params.valueAt(binding.parameterIndex).ParamArray {
						continue
					}
					relevantArgument = true
					if binding.parameterIndex < target.Params.Len() && parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) && ctx.arrayByRefAllocations[key][binding.parameterIndex] {
						guaranteedAllocated = true
						continue
					}
					unknown = true
				}
			} else if call.Arguments.Count > 0 {
				// A resolved call with an incomplete argument projection cannot
				// prove which formal parameter receives the tracked array.
				unknown = true
			}
		} else {
			// Keep unresolved calls conservative by checking the raw actual
			// expressions. Named arguments are handled here only because there
			// is no resolved formal signature to bind them to.
			for _, argument := range call.Arguments.Named {
				if directArrayArgumentName(argument.ValueText) == name {
					relevantArgument = true
					unknown = true
				}
			}
			for _, argument := range arguments {
				if directArrayArgumentName(argument) == name {
					relevantArgument = true
					unknown = true
				}
			}
		}
		if call.Arguments.Count > 0 && len(arguments) != call.Arguments.Count {
			// An incomplete argument projection cannot prove that this call did
			// not receive the tracked array by reference.
			unknown = true
		}
		if !relevantArgument && moduleArrays[name] {
			if arrayLocalGoSubCallHasProvenModuleContract(ctx, call, name) {
				guaranteedAllocated = true
			} else {
				unknown = true
			}
		}
	}
	return unknown, guaranteedAllocated
}

func arrayLocalGoSubCallHasProvenModuleContract(ctx analysisContext, call procedureir.CallSite, name string) bool {
	key, _, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return false
	}
	return ctx.arrayModuleAllocations[key][strings.ToLower(cleanIdentifier(name))]
}

func arrayLocalGoSubIsTerminalBlock(block vbacfg.Block) bool {
	switch block.Kind {
	case vbacfg.BlockNormalExit, vbacfg.BlockExceptionalExit, vbacfg.BlockTerminationExit, vbacfg.BlockUnknownExit:
		return true
	default:
		return false
	}
}

func arrayLocalGoSubAllocationIsReliable(text, name string, ctx analysisContext, base int, constants map[string]int) bool {
	if !arrayStatementAllocatesName(text, name, ctx) {
		return false
	}
	return !arraySummaryStatementAlwaysFails(text, base, constants)
}

func arrayLocalGoSubLabelName(statement procedureir.Statement) string {
	label := statement.Label
	if label == "" {
		label = statement.Text
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(label, ":")))
}

func arrayLocalGoSubIsReturn(statement procedureir.Statement) bool {
	return strings.EqualFold(strings.TrimSpace(statement.Text), "return")
}

func arrayLocalGoSubTarget(proc sourceProcedure, call procedureir.CallSite) string {
	statement, ok := arrayProcedureStatementByID(proc, call.StatementID)
	if !ok || !strings.EqualFold(call.Callee.BaseName, "gosub") {
		return ""
	}
	text := strings.TrimSpace(statement.Text)
	if len(text) < len("gosub") || !strings.EqualFold(text[:len("gosub")], "gosub") {
		return ""
	}
	fields := strings.Fields(strings.TrimSpace(text[len("gosub"):]))
	if len(fields) == 0 {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(fields[0], ":")))
}

func applyArrayLocalGoSubEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, summaries arrayLocalGoSubAllocations) arrayFlowState {
	target := arrayLocalGoSubTarget(proc, call)
	if target == "" {
		return state
	}
	return applyArrayLocalGoSubSummary(state, target, summaries)
}

func applyArrayLocalGoSubSummary(state arrayFlowState, target string, summaries arrayLocalGoSubAllocations) arrayFlowState {
	summary, known := summaries[target]
	if !known {
		// A local GoSub can mutate caller/module arrays without carrying an
		// explicit array argument. An absent summary therefore means unknown
		// state, not "no effect"; retaining allocated here would suppress
		// VBA227 after cleanup or another unmodelled side effect.
		updated := cloneArrayState(state)
		for name := range updated {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
		return updated
	}
	updated := cloneArrayState(state)
	for name := range summary.unknown {
		updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
	}
	for name := range summary.guaranteedAllocated {
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func applyArrayLocalGoSubStatementEffects(state arrayFlowState, text string, summaries arrayLocalGoSubAllocations) arrayFlowState {
	for _, statement := range splitRangeValueSourceStatements(text) {
		target, ok := arrayLocalGoSubTargetFromStatementText(statement)
		if !ok {
			continue
		}
		state = applyArrayLocalGoSubSummary(state, target, summaries)
	}
	return state
}

func arrayLocalGoSubTargetFromStatementText(text string) (string, bool) {
	fields := strings.Fields(strings.TrimSpace(text))
	if len(fields) < 2 || !strings.EqualFold(fields[0], "gosub") {
		return "", false
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(fields[1], ":"))), true
}

func inferArrayByRefEntryStates(a Analyzer, files []parsedFile, ctx analysisContext) (map[string]map[int]bool, map[string]map[int]string) {
	targets := ctx.arrayPrivateTargets
	if len(targets) == 0 {
		return map[string]map[int]bool{}, map[string]map[int]string{}
	}
	moduleAllocationSummaries := ctx.arrayModuleAllocations
	moduleInitializationStates := arrayModuleInitializationStates(files, moduleAllocationSummaries)

	type callerInfo struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
		variables   map[string]arrayVariable
		constants   map[string]int
	}
	callers := make([]callerInfo, 0)
	for _, file := range files {
		moduleDecls := file.moduleDecls()
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			caller := procedures.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, caller) {
				continue
			}
			eligibleCaller := false
			for call := range caller.Calls.All() {
				_, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
				if ok && procedureHasByRefArrayParameter(target) && arrayProcedureIsParticipant(ctx, target) {
					eligibleCaller = true
					break
				}
			}
			if !eligibleCaller {
				continue
			}
			callers = append(callers, callerInfo{
				file: file, proc: caller, moduleDecls: moduleDecls,
				variables: arrayVariables(file, caller, moduleDecls),
				constants: arrayIntegerConstants(file, caller, a.visibleConstantValues, a.visibleConstants),
			})
		}
	}
	sort.SliceStable(callers, func(i, j int) bool {
		return arrayProcedureLess(callers[i].proc, callers[j].proc)
	})

	evaluateCaller := func(caller callerInfo, entries map[string]map[int]bool, conditions map[string]map[int]string) map[string]map[int]arrayByRefEntryEvidence {
		file := caller.file
		proc := caller.proc
		moduleDecls := caller.moduleDecls
		variables := caller.variables
		constants := caller.constants
		localCtx := ctx
		localCtx.arrayByRefEntryConditions = conditions
		evidence := map[string]map[int]arrayByRefEntryEvidence{}
		initial := arrayInitialState(variables)
		initial = applyArrayModuleInitializationState(initial, file, proc, variables, moduleDecls, moduleInitializationStates)
		initial = applyArrayModuleReadyGuardState(initial, file, proc, variables, moduleDecls, ctx.arrayModuleReadyGuards)
		initial = applyArrayByRefEntryStates(initial, proc, variables, entries, conditions)
		initial = applyArrayModuleEntryState(initial, file, proc, variables, moduleDecls, ctx.arrayModuleEntryStates, ctx.arrayParticipantKeys)
		initial = applyArrayInternalStorageConfiguration(initial, file, proc, variables, moduleDecls, ctx.arrayModuleConfigurations[file.Path])
		var baseView vbacfg.CFGView
		var summaryGraph *vbacfg.CFGView
		var worklistReachable map[vbacfg.BlockID]bool
		worklistReachableLines := map[int]bool{}
		if proc.Graph != nil {
			baseView = proc.Graph.View(vbacfg.EdgeFilter{})
			summaryGraph = &baseView
			worklistReachable = arrayCFGWorklistReachable(&baseView)
			for statement := range proc.Statements.All() {
				line := statement.Range.StartLine
				owner, ownerOK := baseView.BlockForStatement(statement.ID)
				if line > 0 && ownerOK && worklistReachable[owner.ID] {
					worklistReachableLines[line] = true
				}
			}
		}
		moduleArrays := arrayModuleNamesForProcedure(file, proc, moduleDecls)
		localGoSubAllocations := arrayLocalGoSubAllocationSummaries(proc, summaryGraph, variables, localCtx, arrayOptionBase(file), constants, moduleArrays)
		visitForBlock := func(text string, line int, in arrayFlowState, ownerStatementID int, filterNestedCalls, skipNestedState bool) arrayFlowState {
			if skipNestedState {
				return in
			}
			var eligible []arrayByRefCallCandidate
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				// A physical line can own more than one CFG statement, such as
				// `If Not SendFrameFor(...) Then Exit Function`.  Only the
				// statement block that owns the call site may contribute entry
				// evidence; otherwise the nested Exit block re-records the same
				// call with its pre-call state and poisons the caller contract.
				if ownerStatementID > 0 && call.StatementID != ownerStatementID {
					if !filterNestedCalls {
						continue
					}
					owner, ownerOK := baseView.BlockForStatement(call.StatementID)
					if ownerOK && worklistReachable[owner.ID] {
						continue
					}
				}
				key, target, ok := arrayPrivateTargetForCall(localCtx, targets, call)
				if !ok || !procedureHasByRefArrayParameter(target) || !arrayProcedureIsParticipant(localCtx, target) {
					continue
				}
				eligible = append(eligible, arrayByRefCallCandidate{key: key, target: target, call: call})
			}
			if len(eligible) > 0 {
				allSameTarget := true
				for _, entry := range eligible[1:] {
					if entry.key != eligible[0].key {
						allSameTarget = false
						break
					}
				}
				// Record and apply each call in source order. This matters when
				// one physical line invokes the same ByRef helper repeatedly: the
				// second call must see an Erase or other invalidation from the
				// first call rather than the original pre-line state.
				recordState := cloneArrayState(in)
				for _, entry := range eligible {
					record := allSameTarget
					if !record {
						// Nested calls on one source line are normally kept
						// conservative because the pre-line state cannot describe
						// mutations from an earlier, different helper. An outer
						// ByRef call whose array argument is a proven allocated
						// expression is independent of that ordering, however.
						allProven, hasExpression := arrayByRefCallHasProvenArrayArguments(file, entry.target, proc, entry.call, recordState, localCtx)
						record = allProven && (hasExpression || arrayByRefCallIsInnermostNested(entry.call, eligible))
					}
					if record {
						arrayRecordByRefCall(evidence, entry.key, entry.target, proc, entry.call, file, recordState, localCtx)
					}
					recordState = applyArrayModuleCallEffects(recordState, file, proc, entry.call, localCtx, variables, moduleDecls)
					if arrayProcedureLineHasInlineConditional(file, entry.call.Range.StartLine) {
						recordState = applyArrayConditionalByRefCallEffects(recordState, proc, entry.call, localCtx)
					} else {
						recordState = applyArrayByRefCallEffects(recordState, proc, entry.call, localCtx)
					}
					recordState = applyArrayLocalGoSubEffects(recordState, proc, entry.call, localGoSubAllocations)
				}
			}
			// ByRef entry proofs must use the same logical-line normalization as
			// VBA227 itself; otherwise a continued Split assignment can make a
			// later call-site argument look unallocated.
			out, _ := a.arrayVBA227Transfer(file, proc, localCtx, variables, in, text, line, constants, nil, nil)
			out = applyArrayLocalGoSubStatementEffects(out, text, localGoSubAllocations)
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				if ownerStatementID > 0 && call.StatementID != ownerStatementID {
					if !filterNestedCalls {
						continue
					}
					owner, ownerOK := baseView.BlockForStatement(call.StatementID)
					if ownerOK && worklistReachable[owner.ID] {
						// This callback is visiting the container statement's
						// source range. Nested calls are visited again by their own
						// CFG block, so applying their post-call effect here would
						// leak one branch into its siblings.
						continue
					}
				}
				out = applyArrayModuleCallEffects(out, file, proc, call, localCtx, variables, moduleDecls)
				out = applyArrayUnknownModuleCallEffects(out, file, proc, call, localCtx, variables, moduleDecls)
				if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
					out = applyArrayConditionalByRefCallEffects(out, proc, call, localCtx)
				} else {
					out = applyArrayByRefCallEffects(out, proc, call, localCtx)
				}
				out = applyArrayLocalGoSubEffects(out, proc, call, localGoSubAllocations)
			}
			return out
		}
		visit := func(text string, line int, in arrayFlowState) arrayFlowState {
			return visitForBlock(text, line, in, 0, false, false)
		}
		visitBlock := func(block vbacfg.Block, text string, line int, in arrayFlowState) arrayFlowState {
			filterNestedCalls := arrayCFGBlockOwnsNestedStatements(block)
			skipNestedState := false
			if filterNestedCalls && block.Statement != nil {
				start := block.Statement.Range.StartLine
				if start == 0 {
					start = block.Range.StartLine
				}
				skipNestedState = line > start && worklistReachableLines[line]
			}
			return visitForBlock(text, line, in, block.StatementID, filterNestedCalls, skipNestedState)
		}
		if proc.Graph == nil {
			state := initial
			for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
				state = visit(normalizedCodeLine(file.Lines[line-1]), line, state)
			}
			return evidence
		}
		fallbackFacts := buildArraySourceOrderFallbackFacts(file, proc, &baseView, variables, localCtx, constants)
		fallbackFacts.unknownFlow = len(proc.Graph.UnknownFlowSources) > 0
		if ctx.arrayStats != nil {
			ctx.arrayStats.addCFGWalk()
		}
		walkArrayCFGWithSourceLinesReliableStatsAndBlock(&baseView, file.Lines, initial, visit, visitBlock, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			out = applyArrayConditionalAllocationBranch(out, &baseView, block, edge)
			out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
		}, nil, ctx.arrayStats)
		// A recovered source construct can make a call block unreachable in the
		// CFG even though the call is valid VBA source.  The parser currently
		// represents some colon-separated single-line statements this way.  For
		// that narrow boundary, retain a call-site allocation proof from the
		// lexical source order; only direct array arguments with an allocation
		// invariant across all reachable branch alternatives are admitted.
		for call := range proc.Calls.All() {
			block, ok := proc.Graph.BlockForStatement(call.StatementID)
			if !ok || worklistReachable[block.ID] {
				continue
			}
			if !arrayByRefSourceOrderFallbackApplies(file, proc, &baseView, fallbackFacts, call) {
				continue
			}
			key, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
			if !ok || !procedureHasByRefArrayParameter(target) || !arrayProcedureIsParticipant(ctx, target) {
				continue
			}
			if state, proven := a.arrayByRefCallSourceOrderProof(file, fallbackFacts, localGoSubAllocations, proc, target, call, initial, ctx, variables, constants); proven {
				arrayRecordByRefCall(evidence, key, target, proc, call, file, state, ctx)
			}
		}
		return evidence
	}

	dependents := make(map[string][]int)
	indexByKey := make(map[string]int, len(callers))
	for index, caller := range callers {
		indexByKey[arrayProcedureKey(caller.proc)] = index
		for call := range caller.proc.Calls.All() {
			if key, target, ok := arrayPrivateTargetForCall(ctx, targets, call); ok && procedureHasByRefArrayParameter(target) && arrayProcedureIsParticipant(ctx, target) {
				dependents[key] = append(dependents[key], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	contributions := make(map[string]map[string]map[int]arrayByRefEntryEvidence, len(callers))
	entries := map[string]map[int]bool{}
	conditions := map[string]map[int]string{}
	queue := make([]int, len(callers))
	queued := make([]bool, len(callers))
	for index := range callers {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(callers) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		caller := callers[index]
		key := arrayProcedureKey(caller.proc)
		evidence := evaluateCaller(caller, entries, conditions)
		if arrayByRefEvidenceMapsEqual(contributions[key], evidence) {
			continue
		}
		contributions[key] = evidence
		merged := map[string]map[int]arrayByRefEntryEvidence{}
		for _, callerEvidence := range contributions {
			mergeArrayByRefEntryEvidence(merged, callerEvidence)
		}
		result := map[string]map[int]bool{}
		conditionalResult := map[string]map[int]string{}
		for targetKey, parameters := range merged {
			for parameterIndex, fact := range parameters {
				if !fact.seen {
					continue
				}
				if fact.allocated {
					if result[targetKey] == nil {
						result[targetKey] = map[int]bool{}
					}
					result[targetKey][parameterIndex] = true
				}
				if !fact.allocated && fact.conditionCompatible && fact.condition != "" {
					if conditionalResult[targetKey] == nil {
						conditionalResult[targetKey] = map[int]string{}
					}
					conditionalResult[targetKey][parameterIndex] = fact.condition
				}
			}
		}
		changedTargets := arrayByRefEntryChangedTargets(entries, conditions, result, conditionalResult)
		if len(changedTargets) == 0 {
			continue
		}
		sortArrayProcedureKeys(changedTargets, indexByKey)
		entries = result
		conditions = conditionalResult
		for _, target := range changedTargets {
			if dependent, ok := indexByKey[target]; ok && !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
			for _, dependent := range dependents[target] {
				if !queued[dependent] {
					queued[dependent] = true
					queue = append(queue, dependent)
				}
			}
		}
	}
	return entries, conditions
}

func arrayByRefEntryStatesEqual(left, right map[string]map[int]bool) bool {
	if len(left) != len(right) {
		return false
	}
	for target, parameters := range left {
		other, ok := right[target]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index := range parameters {
			if !other[index] {
				return false
			}
		}
	}
	return true
}

func arrayByRefEntryConditionsEqual(left, right map[string]map[int]string) bool {
	if len(left) != len(right) {
		return false
	}
	for target, parameters := range left {
		other, ok := right[target]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index, condition := range parameters {
			if other[index] != condition {
				return false
			}
		}
	}
	return true
}

func arrayByRefEvidenceMapsEqual(left, right map[string]map[int]arrayByRefEntryEvidence) bool {
	if len(left) != len(right) {
		return false
	}
	for target, leftParameters := range left {
		rightParameters, ok := right[target]
		if !ok || len(leftParameters) != len(rightParameters) {
			return false
		}
		for index, leftFact := range leftParameters {
			if rightParameters[index] != leftFact {
				return false
			}
		}
	}
	return true
}

func mergeArrayByRefEntryEvidence(dst, src map[string]map[int]arrayByRefEntryEvidence) {
	for target, parameters := range src {
		for index, incoming := range parameters {
			if dst[target] == nil {
				dst[target] = map[int]arrayByRefEntryEvidence{}
			}
			current, exists := dst[target][index]
			if !exists {
				dst[target][index] = incoming
				continue
			}
			current.allocated = current.allocated && incoming.allocated
			if !incoming.allocated && incoming.condition == "" {
				current.conditionCompatible = false
			}
			if incoming.condition != "" {
				if current.condition == "" {
					current.condition = incoming.condition
				} else if !strings.EqualFold(current.condition, incoming.condition) {
					current.conditionCompatible = false
				}
			}
			current.conditionCompatible = current.conditionCompatible && incoming.conditionCompatible
			current.seen = current.seen || incoming.seen
			dst[target][index] = current
		}
	}
}

func arrayByRefEntryChangedTargets(oldEntries map[string]map[int]bool, oldConditions map[string]map[int]string, newEntries map[string]map[int]bool, newConditions map[string]map[int]string) []string {
	keys := map[string]bool{}
	for target := range oldEntries {
		keys[target] = true
	}
	for target := range oldConditions {
		keys[target] = true
	}
	for target := range newEntries {
		keys[target] = true
	}
	for target := range newConditions {
		keys[target] = true
	}
	changed := make([]string, 0, len(keys))
	for target := range keys {
		oldEntry := map[string]map[int]bool{}
		newEntry := map[string]map[int]bool{}
		if names := oldEntries[target]; len(names) > 0 {
			oldEntry[target] = names
		}
		if names := newEntries[target]; len(names) > 0 {
			newEntry[target] = names
		}
		oldCondition := map[string]map[int]string{}
		newCondition := map[string]map[int]string{}
		if conditions := oldConditions[target]; len(conditions) > 0 {
			oldCondition[target] = conditions
		}
		if conditions := newConditions[target]; len(conditions) > 0 {
			newCondition[target] = conditions
		}
		if !arrayByRefEntryStatesEqual(oldEntry, newEntry) || !arrayByRefEntryConditionsEqual(oldCondition, newCondition) {
			changed = append(changed, target)
		}
	}
	return changed
}

func sortArrayProcedureKeys(keys []string, indexByKey map[string]int) {
	sort.SliceStable(keys, func(i, j int) bool {
		left, leftOK := indexByKey[keys[i]]
		right, rightOK := indexByKey[keys[j]]
		if leftOK && rightOK && left != right {
			return left < right
		}
		if leftOK != rightOK {
			return leftOK
		}
		return keys[i] < keys[j]
	})
}

func arrayOptionPrivateModule(lines []string) bool {
	for _, line := range lines {
		if strings.EqualFold(strings.TrimSpace(normalizedCodeLine(line)), "option private module") {
			return true
		}
	}
	return false
}

func arrayCallsAtLine(calls readOnlySpan[procedureir.CallSite], line int) []procedureir.CallSite {
	matched := make([]procedureir.CallSite, 0, 1)
	for call := range calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line {
			continue
		}
		matched = append(matched, call)
	}
	return matched
}

func arrayPrivateTargetForCall(ctx analysisContext, targets map[string]sourceProcedure, call procedureir.CallSite) (string, sourceProcedure, bool) {
	resolution := arrayCallResolution(ctx, call)
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return "", sourceProcedure{}, false
	}
	key := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
	target, ok := targets[key]
	return key, target, ok
}

// arrayCallResolution rechecks the one parser shape that currently loses an
// unparenthesized call's procedure name when its implicit member argument is
// written with a leading dot, for example `Consume .values`.  The ordinary
// resolver remains authoritative; the retry is deliberately limited to a
// single dotted argument so an unresolved object member is not reinterpreted
// as a project-local procedure call.
func arrayCallResolution(ctx analysisContext, call procedureir.CallSite) procedureir.CallResolution {
	resolution := call.Resolution
	if ctx.procedureResolver == nil {
		return resolution
	}
	resolution = ctx.procedureResolver.ResolveCall(call)
	if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
		return resolution
	}
	name := arrayImplicitMemberArgumentCallName(call.Callee.Text)
	if name == "" || strings.EqualFold(name, call.Callee.BaseName) {
		return resolution
	}
	retried := call
	retried.Callee.Text = name
	retried.Callee.BaseName = name
	retried.Callee.Member = name
	retried.Callee.Receiver = nil
	return ctx.procedureResolver.ResolveCall(retried)
}

func arrayImplicitMemberArgumentCallName(text string) string {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "call ") {
		text = strings.TrimSpace(text[len("call "):])
	}
	if text == "" || !isIdentifierStart(text[0]) {
		return ""
	}
	end := 1
	for end < len(text) && isIdentifierPart(text[end]) {
		end++
	}
	rest := strings.TrimSpace(text[end:])
	if len(rest) < 2 || rest[0] != '.' || !isIdentifierStart(rest[1]) {
		return ""
	}
	memberEnd := 2
	for memberEnd < len(rest) && isIdentifierPart(rest[memberEnd]) {
		memberEnd++
	}
	if strings.TrimSpace(rest[memberEnd:]) != "" {
		return ""
	}
	return cleanIdentifier(text[:end])
}

func procedureHasByRefArrayParameter(proc sourceProcedure) bool {
	for parameter := range proc.Params.All() {
		if parameterIsByRefArray(parameter) {
			return true
		}
	}
	return false
}

func arrayRecordByRefCall(evidence map[string]map[int]arrayByRefEntryEvidence, targetKey string, target, caller sourceProcedure, call procedureir.CallSite, file parsedFile, state arrayFlowState, ctx analysisContext) {
	// A self-recursive ByRef helper preserves the entry array state supplied by
	// its caller. Treating the recursive edge as an independent unknown entry
	// would poison the evidence from the allocated external call and keep the
	// callee conservative forever.
	if strings.EqualFold(targetKey, arrayProcedureKey(caller)) {
		return
	}
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return
	}
	arguments := make([]string, target.Params.Len())
	bound := make(map[int]bool, len(bindings))
	for _, binding := range bindings {
		arguments[binding.parameterIndex] = binding.text
		bound[binding.parameterIndex] = true
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) || !bound[index] {
			continue
		}
		name := ""
		name = directArrayArgumentName(arguments[index])
		value, known := state[name]
		allocated := known && value.kind == arrayAllocated && value.knownArray
		if !allocated && arrayQualifiedArgumentProvenAllocated(file, caller, call, arguments[index], ctx) {
			value = arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}
			known = true
			allocated = true
		}
		_, _, qualifiedMember := arrayQualifiedMemberParts(arguments[index])
		if !allocated && !qualifiedMember && index < len(arguments) {
			// A function returning a dynamic array is a valid ByRef array
			// argument in VBA.  The identifier-only path above cannot attach
			// that expression to a caller state entry, so consult the existing
			// array-return summaries before treating the callee entry as
			// unallocated.  Unknown or conditionally allocated returns remain
			// conservative because arrayExpressionState returns no allocated
			// proof for them.
			if returned, returnedKnown := arrayExpressionState(arguments[index], state, ctx); returnedKnown && returned.kind == arrayAllocated && returned.knownArray {
				value = returned
				known = true
				allocated = true
			}
		}
		condition := ""
		if known && !allocated {
			if source, positive := arrayVBA227PositiveLengthCondition(value.conditionalAllocationSource); positive {
				condition = arrayConditionalEntrySource(target, arguments, index, source)
			}
			if condition == "" && value.allocationCountSource != "" {
				condition = arrayConditionalEntrySource(target, arguments, index, value.allocationCountSource)
			}
		}
		if !allocated && condition == "" && arrayByRefCallArrayVacuouslyUnused(target, index, arguments) {
			// A call may intentionally pass an unallocated optional array (including
			// an array-return expression with no local state entry) to a helper whose
			// only uses are behind a false literal guard. It must not poison the
			// conditional-entry evidence collected from calls that do reach those
			// uses.
			continue
		}
		parameters := evidence[targetKey]
		if parameters == nil {
			parameters = map[int]arrayByRefEntryEvidence{}
		}
		fact := parameters[index]
		if !fact.seen {
			fact.allocated = allocated
			fact.conditionCompatible = allocated || condition != ""
			fact.condition = condition
		} else {
			fact.allocated = fact.allocated && allocated
			if !allocated && condition == "" {
				fact.conditionCompatible = false
			}
			if condition != "" {
				if fact.condition == "" {
					fact.condition = condition
				} else if !strings.EqualFold(fact.condition, condition) {
					fact.conditionCompatible = false
				}
			}
		}
		fact.seen = true
		parameters[index] = fact
		evidence[targetKey] = parameters
	}
}

type arraySourceOrderAllocation struct {
	line     int
	parentID int
}

type arraySourceOrderFallbackFacts struct {
	conditionalTransferLines   []int
	unconditionalTransferLines []int
	definiteExitLines          []int
	unknownFlow                bool
	parents                    map[int]procedureir.Statement
	allocations                map[string][]arraySourceOrderAllocation
	bypassTargetMin            map[int]int
	branchGroups               map[int]map[int]bool
	branchTransferBypass       map[int]map[string]int
	ambiguousTransferLines     map[int]bool
}
