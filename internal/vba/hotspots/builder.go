package hotspots

import (
	"context"
	"errors"
	"math/bits"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
)

const cycleEnumerationWorkBudget = 100_000

// BuildStats reports graph work performed while building a hotspot report.
// SCC and closure counters make the shared reachability work observable to
// architecture tests without adding fields to the stable report schema.
type BuildStats struct {
	ProcedureCount     int
	ModuleCount        int
	ProcedureEdgeCount int
	ModuleEdgeCount    int
	ProcedureSCCCount  int
	ModuleSCCCount     int
	// AffectedClosureSetCount is the number of final shared sets across both graphs.
	AffectedClosureSetCount  int
	AffectedClosureSetUnions int
	// Storage counters count retained module indices/words, excluding Go headers.
	AffectedClosureSparseEntriesStored int
	AffectedClosureDenseWordsStored    int
	// Scan counters count sparse merge steps and dense words visited during unions.
	AffectedClosureSparseMergeSteps int
	AffectedClosureDenseWordScans   int
	// Each SCC cardinality is read once and then reused for all entities in it.
	AffectedClosureComponentCardinalityReads int
	CycleEnumerationWork                     int
	CycleEnumerationTruncated                bool
}

// BuildFromMetrics projects existing procedure metrics into the shared
// procedure and module hotspot report. Reachable module sets are computed on
// SCC-condensed graphs and shared by every entity in the same component.
// Cancellation returns no partial report.
func BuildFromMetrics(ctx context.Context, procedures []proceduremetrics.ProcedureMetrics) (Report, BuildStats, error) {
	stats := BuildStats{ProcedureCount: len(procedures)}
	if ctx == nil {
		return Report{}, stats, errors.New("hotspot build requires a non-nil context")
	}
	if err := ctx.Err(); err != nil {
		return Report{}, stats, err
	}

	procedureIDs := make(map[string]string, len(procedures))
	moduleIDsByProcedure := make(map[string]string, len(procedures))
	procedureNodes := make([]string, 0, len(procedures))
	moduleNodeSet := make(map[string]struct{})
	procedureAdjacency := make(map[string]map[string]bool)
	moduleOut := make(map[string]map[string]bool)
	moduleIn := make(map[string]map[string]bool)
	callersByCallee := make(map[string]map[string]struct{})

	for _, procedure := range procedures {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		moduleID := procedure.File + "|" + procedure.Module
		id := procedure.File + "|" + procedure.Module + "|" + procedure.Name + "|" + string(procedure.Kind)
		procedureIDs[strings.ToLower(procedure.Module+"."+procedure.Name)] = id
		moduleIDsByProcedure[id] = moduleID
		procedureNodes = append(procedureNodes, id)
		moduleNodeSet[moduleID] = struct{}{}
		if moduleOut[moduleID] == nil {
			moduleOut[moduleID] = make(map[string]bool)
			moduleIn[moduleID] = make(map[string]bool)
		}
	}

	for _, procedure := range procedures {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		callerID := procedure.File + "|" + procedure.Module + "|" + procedure.Name + "|" + string(procedure.Kind)
		callerModuleID := procedure.File + "|" + procedure.Module
		for _, callee := range procedure.ResolvedCallees {
			if err := ctx.Err(); err != nil {
				return Report{}, stats, err
			}
			if callersByCallee[callee] == nil {
				callersByCallee[callee] = make(map[string]struct{})
			}
			callersByCallee[callee][callerID] = struct{}{}

			calleeID, ok := procedureIDs[callee]
			if !ok {
				continue
			}
			if procedureAdjacency[callerID] == nil {
				procedureAdjacency[callerID] = make(map[string]bool)
			}
			procedureAdjacency[callerID][calleeID] = true
			calleeModuleID := moduleIDsByProcedure[calleeID]
			if callerModuleID == calleeModuleID {
				continue
			}
			moduleOut[callerModuleID][calleeModuleID] = true
			moduleIn[calleeModuleID][callerModuleID] = true
		}
	}

	for _, targets := range procedureAdjacency {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		stats.ProcedureEdgeCount += len(targets)
	}
	for _, targets := range moduleOut {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		stats.ModuleEdgeCount += len(targets)
	}

	moduleNodes := make([]string, 0, len(moduleNodeSet))
	for moduleID := range moduleNodeSet {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		moduleNodes = append(moduleNodes, moduleID)
	}
	slices.Sort(moduleNodes)
	stats.ModuleCount = len(moduleNodes)

	procedureAffected, procedureSCCs, err := reachableModuleCounts(
		ctx, procedureNodes, procedureAdjacency, moduleIDsByProcedure, moduleNodes, &stats,
	)
	if err != nil {
		return Report{}, stats, err
	}
	stats.ProcedureSCCCount = procedureSCCs

	moduleByModule := make(map[string]string, len(moduleNodes))
	for _, moduleID := range moduleNodes {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		moduleByModule[moduleID] = moduleID
	}
	moduleAffected, moduleSCCs, err := reachableModuleCounts(
		ctx, moduleNodes, moduleOut, moduleByModule, moduleNodes, &stats,
	)
	if err != nil {
		return Report{}, stats, err
	}
	stats.ModuleSCCCount = moduleSCCs

	cycleParticipation, cycleWork, cycleTruncated, err := cycleParticipationCounts(ctx, procedureAdjacency)
	stats.CycleEnumerationWork = cycleWork
	stats.CycleEnumerationTruncated = cycleTruncated
	if err != nil {
		return Report{}, stats, err
	}

	procedureInputs := make([]Input, 0, len(procedures))
	moduleMap := make(map[string]Input, len(moduleNodes))
	for _, procedure := range procedures {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		id := procedure.File + "|" + procedure.Module + "|" + procedure.Name + "|" + string(procedure.Kind)
		moduleID := procedure.File + "|" + procedure.Module
		raw := map[string]int{
			string(SignalComplexity):       procedure.CyclomaticComplexity,
			string(SignalCallFanOut):       procedure.CallFanOut,
			string(SignalCallFanIn):        len(callersByCallee[strings.ToLower(procedure.Module+"."+procedure.Name)]),
			string(SignalAffectedModules):  procedureAffected[id],
			string(SignalExcelEffects):     procedure.ExcelEffectCount,
			string(SignalMutableReads):     procedure.MutableStateReads,
			string(SignalMutableWrites):    procedure.MutableStateWrites,
			string(SignalMutableMutations): procedure.MutableStateMutations,
			string(SignalCycleCount):       cycleParticipation[id],
			string(SignalExternalDeps):     procedure.ExternalDependencyCount,
			string(SignalErrorHandling):    procedure.ErrorHandlingCount,
			string(SignalResourceOwned):    procedure.ResourceOwnershipCount,
		}
		uncertainty := map[string]int{
			string(UncertaintyAmbiguous):  procedure.AmbiguousCallCount,
			string(UncertaintyUnresolved): procedure.UnresolvedCallCount,
			string(UncertaintyDynamic):    procedure.DynamicCallCount,
		}
		procedureInputs = append(procedureInputs, Input{
			ID: id, Kind: "procedure", File: procedure.File, Module: procedure.Module,
			ModuleKind: procedure.ModuleKind, Name: procedure.Name, ProcedureKind: string(procedure.Kind),
			DeclarationByte: procedure.DeclarationRange.StartByte, Line: procedure.DeclarationRange.StartLine,
			RawSignals: raw, Uncertainty: uncertainty,
		})

		module := moduleMap[moduleID]
		if module.ID == "" {
			module = Input{
				ID: moduleID, Kind: "module", File: procedure.File, Module: procedure.Module,
				ModuleKind: procedure.ModuleKind, Name: procedure.Module,
				Line: procedure.DeclarationRange.StartLine, RawSignals: (ModuleSignals{}).Map(),
				Uncertainty: make(map[string]int),
			}
		} else if line := procedure.DeclarationRange.StartLine; line > 0 && (module.Line == 0 || line < module.Line) {
			module.Line = line
		}
		module.RawSignals[string(SignalComplexity)] += procedure.CyclomaticComplexity
		if procedure.CyclomaticComplexity > module.RawSignals[string(SignalComplexityMax)] {
			module.RawSignals[string(SignalComplexityMax)] = procedure.CyclomaticComplexity
		}
		module.RawSignals[string(SignalCallFanOut)] += procedure.CallFanOut
		module.RawSignals[string(SignalExcelEffects)] += procedure.ExcelEffectCount
		module.RawSignals[string(SignalMutableReads)] += procedure.MutableStateReads
		module.RawSignals[string(SignalMutableWrites)] += procedure.MutableStateWrites
		module.RawSignals[string(SignalMutableMutations)] += procedure.MutableStateMutations
		module.RawSignals[string(SignalCycleCount)] += cycleParticipation[id]
		module.RawSignals[string(SignalExternalDeps)] += procedure.ExternalDependencyCount
		module.RawSignals[string(SignalErrorHandling)] += procedure.ErrorHandlingCount
		module.RawSignals[string(SignalResourceOwned)] += procedure.ResourceOwnershipCount
		for key, value := range uncertainty {
			module.Uncertainty[key] += value
		}
		if !strings.EqualFold(strings.TrimSpace(procedure.Visibility), "private") {
			module.RawSignals[string(SignalPublicProcedures)]++
		}
		moduleMap[moduleID] = module
	}

	moduleInputs := make([]Input, 0, len(moduleMap))
	for _, moduleID := range moduleNodes {
		if err := ctx.Err(); err != nil {
			return Report{}, stats, err
		}
		module := moduleMap[moduleID]
		module.RawSignals[string(SignalCallFanIn)] = len(moduleIn[moduleID])
		module.RawSignals[string(SignalCallFanOut)] = len(moduleOut[moduleID])
		module.RawSignals[string(SignalAffectedModules)] = moduleAffected[moduleID]
		moduleInputs = append(moduleInputs, module)
	}
	if err := ctx.Err(); err != nil {
		return Report{}, stats, err
	}
	report := BuildReport(procedureInputs, moduleInputs)
	if err := ctx.Err(); err != nil {
		return Report{}, stats, err
	}
	return report, stats, nil
}

func reachableModuleCounts(
	ctx context.Context,
	nodes []string,
	adjacency map[string]map[string]bool,
	modulesByNode map[string]string,
	moduleIDs []string,
	stats *BuildStats,
) (map[string]int, int, error) {
	components, componentOf, err := stronglyConnectedComponents(ctx, nodes, adjacency)
	if err != nil {
		return nil, 0, err
	}
	moduleIndexes := make(map[string]int, len(moduleIDs))
	for index, moduleID := range moduleIDs {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		moduleIndexes[moduleID] = index
	}
	ownedModules := make([][]int, len(components))
	for _, node := range nodes {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		moduleIndex, ok := moduleIndexes[modulesByNode[node]]
		if !ok {
			continue
		}
		component := componentOf[node]
		ownedModules[component] = append(ownedModules[component], moduleIndex)
	}
	sets := make([]moduleSet, len(components))
	for component, modules := range ownedModules {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		slices.Sort(modules)
		modules = slices.Compact(modules)
		set, err := newModuleSet(ctx, modules, len(moduleIDs))
		if err != nil {
			return nil, len(components), err
		}
		sets[component] = set
	}

	componentEdges := make([]map[int]struct{}, len(components))
	indegree := make([]int, len(components))
	for _, from := range nodes {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		fromComponent := componentOf[from]
		for to := range adjacency[from] {
			if err := ctx.Err(); err != nil {
				return nil, len(components), err
			}
			toComponent, exists := componentOf[to]
			if !exists || fromComponent == toComponent {
				continue
			}
			if componentEdges[fromComponent] == nil {
				componentEdges[fromComponent] = make(map[int]struct{})
			}
			if _, exists := componentEdges[fromComponent][toComponent]; exists {
				continue
			}
			componentEdges[fromComponent][toComponent] = struct{}{}
			indegree[toComponent]++
		}
	}

	ready := make([]int, 0, len(components))
	for component, degree := range indegree {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		if degree == 0 {
			ready = append(ready, component)
		}
	}
	topological := make([]int, 0, len(components))
	for head := 0; head < len(ready); head++ {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		component := ready[head]
		topological = append(topological, component)
		for next := range componentEdges[component] {
			if err := ctx.Err(); err != nil {
				return nil, len(components), err
			}
			indegree[next]--
			if indegree[next] == 0 {
				ready = append(ready, next)
			}
		}
	}
	if len(topological) != len(components) {
		return nil, len(components), errors.New("hotspot SCC condensation contains a cycle")
	}

	for index := len(topological) - 1; index >= 0; index-- {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		component := topological[index]
		for next := range componentEdges[component] {
			if err := ctx.Err(); err != nil {
				return nil, len(components), err
			}
			set, err := unionModuleSets(ctx, sets[component], sets[next], len(moduleIDs), stats)
			if err != nil {
				return nil, len(components), err
			}
			sets[component] = set
			stats.AffectedClosureSetUnions++
		}
	}

	componentCardinalities := make([]int, len(components))
	for component, set := range sets {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		componentCardinalities[component] = set.cardinality
		stats.AffectedClosureSetCount++
		stats.AffectedClosureComponentCardinalityReads++
		if len(set.dense) > 0 {
			stats.AffectedClosureDenseWordsStored += len(set.dense)
		} else {
			stats.AffectedClosureSparseEntriesStored += len(set.sparse)
		}
	}
	counts := make(map[string]int, len(nodes))
	for _, node := range nodes {
		if err := ctx.Err(); err != nil {
			return nil, len(components), err
		}
		counts[node] = componentCardinalities[componentOf[node]]
	}
	return counts, len(components), nil
}

// moduleSet is immutable after construction, allowing empty unions to reuse an
// existing set. Sparse sets keep isolated and low-fanout components
// proportional to the modules they actually reach.
type moduleSet struct {
	sparse      []uint32
	dense       []uint64
	cardinality int
}

func newModuleSet(ctx context.Context, moduleIndexes []int, universeSize int) (moduleSet, error) {
	if preferDenseModuleSet(len(moduleIndexes), universeSize) {
		words := make([]uint64, (universeSize+63)/64)
		for _, index := range moduleIndexes {
			if err := ctx.Err(); err != nil {
				return moduleSet{}, err
			}
			words[index/64] |= uint64(1) << uint(index%64)
		}
		return moduleSet{dense: words, cardinality: len(moduleIndexes)}, nil
	}
	sparse := make([]uint32, len(moduleIndexes))
	for index, moduleIndex := range moduleIndexes {
		if err := ctx.Err(); err != nil {
			return moduleSet{}, err
		}
		sparse[index] = uint32(moduleIndex)
	}
	return moduleSet{sparse: sparse, cardinality: len(sparse)}, nil
}

func unionModuleSets(ctx context.Context, left, right moduleSet, universeSize int, stats *BuildStats) (moduleSet, error) {
	if left.cardinality == 0 {
		return right, nil
	}
	if right.cardinality == 0 {
		return left, nil
	}
	if len(left.dense) > 0 && len(right.dense) > 0 {
		words := make([]uint64, len(left.dense))
		cardinality := 0
		for index := range words {
			if err := ctx.Err(); err != nil {
				return moduleSet{}, err
			}
			words[index] = left.dense[index] | right.dense[index]
			cardinality += bits.OnesCount64(words[index])
			stats.AffectedClosureDenseWordScans++
		}
		return moduleSet{dense: words, cardinality: cardinality}, nil
	}
	if len(left.dense) > 0 || len(right.dense) > 0 {
		denseSet, sparseSet := left, right
		if len(denseSet.dense) == 0 {
			denseSet, sparseSet = right, left
		}
		words := make([]uint64, len(denseSet.dense))
		for index, word := range denseSet.dense {
			if err := ctx.Err(); err != nil {
				return moduleSet{}, err
			}
			words[index] = word
			stats.AffectedClosureDenseWordScans++
		}
		cardinality := denseSet.cardinality
		for _, moduleIndex := range sparseSet.sparse {
			if err := ctx.Err(); err != nil {
				return moduleSet{}, err
			}
			stats.AffectedClosureSparseMergeSteps++
			wordIndex := int(moduleIndex) / 64
			mask := uint64(1) << (moduleIndex % 64)
			if words[wordIndex]&mask == 0 {
				words[wordIndex] |= mask
				cardinality++
			}
		}
		return moduleSet{dense: words, cardinality: cardinality}, nil
	}

	merged := make([]uint32, 0, min(left.cardinality+right.cardinality, universeSize))
	for leftIndex, rightIndex := 0, 0; leftIndex < len(left.sparse) || rightIndex < len(right.sparse); {
		if err := ctx.Err(); err != nil {
			return moduleSet{}, err
		}
		stats.AffectedClosureSparseMergeSteps++
		switch {
		case rightIndex == len(right.sparse) || (leftIndex < len(left.sparse) && left.sparse[leftIndex] < right.sparse[rightIndex]):
			merged = append(merged, left.sparse[leftIndex])
			leftIndex++
		case leftIndex == len(left.sparse) || right.sparse[rightIndex] < left.sparse[leftIndex]:
			merged = append(merged, right.sparse[rightIndex])
			rightIndex++
		default:
			merged = append(merged, left.sparse[leftIndex])
			leftIndex++
			rightIndex++
		}
	}
	if preferDenseModuleSet(len(merged), universeSize) {
		words := make([]uint64, (universeSize+63)/64)
		for _, moduleIndex := range merged {
			if err := ctx.Err(); err != nil {
				return moduleSet{}, err
			}
			words[int(moduleIndex)/64] |= uint64(1) << (moduleIndex % 64)
			stats.AffectedClosureSparseMergeSteps++
		}
		return moduleSet{dense: words, cardinality: len(merged)}, nil
	}
	return moduleSet{sparse: merged, cardinality: len(merged)}, nil
}

func preferDenseModuleSet(cardinality, universeSize int) bool {
	return universeSize > 0 && cardinality*32 >= universeSize
}

func stronglyConnectedComponents(
	ctx context.Context,
	nodes []string,
	adjacency map[string]map[string]bool,
) ([][]string, map[string]int, error) {
	nodeSet := make(map[string]struct{}, len(nodes))
	orderedNodes := make([]string, 0, len(nodes))
	for _, node := range nodes {
		if err := ctx.Err(); err != nil {
			return nil, nil, err
		}
		if _, exists := nodeSet[node]; exists {
			continue
		}
		nodeSet[node] = struct{}{}
		orderedNodes = append(orderedNodes, node)
	}
	slices.Sort(orderedNodes)

	neighbors := make(map[string][]string, len(orderedNodes))
	reverse := make(map[string][]string, len(orderedNodes))
	for _, node := range orderedNodes {
		if err := ctx.Err(); err != nil {
			return nil, nil, err
		}
		for next := range adjacency[node] {
			if err := ctx.Err(); err != nil {
				return nil, nil, err
			}
			if _, exists := nodeSet[next]; !exists {
				continue
			}
			neighbors[node] = append(neighbors[node], next)
			reverse[next] = append(reverse[next], node)
		}
		slices.Sort(neighbors[node])
	}
	for node := range reverse {
		if err := ctx.Err(); err != nil {
			return nil, nil, err
		}
		slices.Sort(reverse[node])
	}

	type frame struct {
		node string
		next int
	}
	visited := make(map[string]bool, len(orderedNodes))
	finishOrder := make([]string, 0, len(orderedNodes))
	for _, start := range orderedNodes {
		if visited[start] {
			continue
		}
		if err := ctx.Err(); err != nil {
			return nil, nil, err
		}
		visited[start] = true
		stack := []frame{{node: start}}
		for len(stack) > 0 {
			if err := ctx.Err(); err != nil {
				return nil, nil, err
			}
			top := &stack[len(stack)-1]
			if top.next < len(neighbors[top.node]) {
				next := neighbors[top.node][top.next]
				top.next++
				if !visited[next] {
					visited[next] = true
					stack = append(stack, frame{node: next})
				}
				continue
			}
			finishOrder = append(finishOrder, top.node)
			stack = stack[:len(stack)-1]
		}
	}

	clear(visited)
	components := make([][]string, 0)
	componentOf := make(map[string]int, len(orderedNodes))
	for index := len(finishOrder) - 1; index >= 0; index-- {
		start := finishOrder[index]
		if visited[start] {
			continue
		}
		if err := ctx.Err(); err != nil {
			return nil, nil, err
		}
		componentID := len(components)
		component := make([]string, 0, 1)
		visited[start] = true
		stack := []string{start}
		for len(stack) > 0 {
			if err := ctx.Err(); err != nil {
				return nil, nil, err
			}
			last := len(stack) - 1
			node := stack[last]
			stack = stack[:last]
			component = append(component, node)
			componentOf[node] = componentID
			for _, next := range reverse[node] {
				if err := ctx.Err(); err != nil {
					return nil, nil, err
				}
				if !visited[next] {
					visited[next] = true
					stack = append(stack, next)
				}
			}
		}
		components = append(components, component)
	}
	return components, componentOf, nil
}

// cycleParticipationCounts retains the canonical elementary-cycle traversal
// and fixed work budget used by metrics before hotspot aggregation was shared.
func cycleParticipationCounts(ctx context.Context, adjacency map[string]map[string]bool) (map[string]int, int, bool, error) {
	result := make(map[string]int)
	nodesSet := make(map[string]bool)
	for from, targets := range adjacency {
		if err := ctx.Err(); err != nil {
			return nil, 0, false, err
		}
		nodesSet[from] = true
		for to := range targets {
			if err := ctx.Err(); err != nil {
				return nil, 0, false, err
			}
			nodesSet[to] = true
		}
	}
	nodes := make([]string, 0, len(nodesSet))
	for node := range nodesSet {
		if err := ctx.Err(); err != nil {
			return nil, 0, false, err
		}
		nodes = append(nodes, node)
	}
	slices.Sort(nodes)
	work := 0
	exhausted := false
	for _, start := range nodes {
		if err := ctx.Err(); err != nil {
			return nil, work, exhausted, err
		}
		if exhausted {
			break
		}
		path := []string{start}
		visited := map[string]bool{start: true}
		var visit func(string) error
		visit = func(current string) error {
			if err := ctx.Err(); err != nil {
				return err
			}
			if exhausted {
				return nil
			}
			work++
			if work > cycleEnumerationWorkBudget {
				exhausted = true
				return nil
			}
			nexts := make([]string, 0, len(adjacency[current]))
			for next := range adjacency[current] {
				if err := ctx.Err(); err != nil {
					return err
				}
				nexts = append(nexts, next)
			}
			slices.Sort(nexts)
			for _, next := range nexts {
				if err := ctx.Err(); err != nil {
					return err
				}
				work++
				if work > cycleEnumerationWorkBudget {
					exhausted = true
					return nil
				}
				if next == start {
					for _, member := range path {
						if err := ctx.Err(); err != nil {
							return err
						}
						result[member]++
					}
					continue
				}
				if next < start || visited[next] {
					continue
				}
				visited[next] = true
				path = append(path, next)
				err := visit(next)
				path = path[:len(path)-1]
				delete(visited, next)
				if err != nil {
					return err
				}
			}
			return nil
		}
		if err := visit(start); err != nil {
			return nil, work, exhausted, err
		}
	}
	return result, work, exhausted, nil
}
