package analyze

import (
	"context"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// arrayCompactLattice is the adapter between the array allocation lattice and
// semanticstate's indexed storage.  The array domain deliberately keeps its
// existing meet semantics; only the representation and fixed-point scheduler
// are shared.
type arrayCompactLattice struct{}

func unknownArrayValue() arrayValue { return arrayValue{kind: arrayUnknown} }

func (arrayCompactLattice) Clone(value arrayValue) arrayValue {
	// Array shapes are immutable once published into the compact lattice.  Array
	// transfers replace a shape with a newly parsed slice instead of mutating a
	// slice held by an existing value, so a solver scratch copy only needs to
	// copy the value header.  Legacy map ownership still uses cloneArrayState.
	return value
}

func (l arrayCompactLattice) Join(dst *arrayValue, src arrayValue) bool {
	if dst == nil {
		return false
	}
	merged := meetArrayValue(*dst, src)
	if arrayValueEqual(*dst, merged) {
		return false
	}
	*dst = merged
	return true
}

type arrayCompactAdapter struct {
	environment semanticstate.Environment
	// names is the adapter-owned immutable participant order. Calling
	// Environment.Names repeatedly returns a defensive copy; retaining this
	// slice avoids a per-cursor/per-transfer allocation while preserving the
	// environment's external ownership boundary.
	names []string
}

func newArrayCompactAdapter(initials ...arrayFlowState) arrayCompactAdapter {
	nameSet := make(map[string]struct{})
	for _, initial := range initials {
		for name := range initial {
			nameSet[name] = struct{}{}
		}
	}
	names := make([]string, 0, len(nameSet))
	for name := range nameSet {
		names = append(names, name)
	}
	environment := semanticstate.NewEnvironment(names, names)
	return arrayCompactAdapter{environment: environment, names: environment.Names()}
}

// newArrayCompactAdapterForLines extends the initial semantic participants
// with assignment targets visible in the current CFG only. Callers often hold
// a module-wide source slice; scanning that entire slice would add unrelated
// procedures to every dense solver state. Array transfer deliberately tracks a
// small number of scalar witnesses (for example an ArrayLength result used as
// an allocation guard), so those witnesses are collected from the CFG's own
// statement ranges before indexing.
func newArrayCompactAdapterForLines(graph *vbacfg.CFGView, lines []string, initials ...arrayFlowState) arrayCompactAdapter {
	nameSet := make(map[string]struct{})
	for _, initial := range initials {
		for name := range initial {
			nameSet[name] = struct{}{}
		}
	}
	addAssignments := func(text string) {
		for _, line := range strings.Split(text, "\n") {
			if lhs, _, _, ok := arrayAssignment(line); ok {
				name := strings.ToLower(cleanIdentifier(lhs))
				if name != "" {
					nameSet[name] = struct{}{}
				}
			}
		}
	}
	if graph != nil {
		graph.ForEachBlock(func(block vbacfg.Block) bool {
			if block.Statement == nil {
				return true
			}
			addAssignments(block.Statement.Text)
			start := block.Statement.Range.StartLine
			end := block.Statement.Range.EndLine
			if start < 1 {
				start = block.Range.StartLine
			}
			if end < start {
				end = start
			}
			if start >= 1 && end <= len(lines) {
				for line := start; line <= end; line++ {
					addAssignments(lines[line-1])
				}
			}
			return true
		})
	}
	names := make([]string, 0, len(nameSet))
	for name := range nameSet {
		names = append(names, name)
	}
	environment := semanticstate.NewEnvironment(names, names)
	return arrayCompactAdapter{environment: environment, names: environment.Names()}
}

func (a arrayCompactAdapter) toFlow(view semanticstate.StateView[arrayValue]) arrayFlowState {
	flow := make(arrayFlowState, len(a.names))
	for _, name := range a.names {
		flow[name] = unknownArrayValue()
	}
	view.ForEach(func(id semanticstate.SymbolID, value arrayValue) bool {
		name, ok := a.environment.Name(id)
		if !ok {
			return true
		}
		flow[name] = arrayCompactLattice{}.Clone(value)
		return true
	})
	return flow
}

func (a arrayCompactAdapter) fromFlow(state *semanticstate.State[arrayValue], flow arrayFlowState) {
	state.Reset()
	for _, name := range a.names {
		value, ok := flow[name]
		if !ok {
			value = unknownArrayValue()
		}
		id, ok := a.environment.Symbol(name)
		if !ok {
			continue
		}
		state.Set(id, arrayCompactLattice{}.Clone(value))
	}
}

func (a arrayCompactAdapter) initialState(state *semanticstate.State[arrayValue], initial arrayFlowState) {
	a.fromFlow(state, initial)
}

// arrayCompactCursor is the compatibility boundary for the existing Array
// transfer callbacks.  Its map is allocated once per lane/scratch role and is
// reused for every transfer or edge.  The fixed-point state itself remains in
// semanticstate slots; the cursor prevents a fresh whole-state map and shape
// copy from being allocated at every callback boundary.
type arrayCompactCursor struct {
	adapter arrayCompactAdapter
	flow    arrayFlowState
}

func (a arrayCompactAdapter) newCursor() *arrayCompactCursor {
	flow := make(arrayFlowState, len(a.names))
	for _, name := range a.names {
		flow[name] = unknownArrayValue()
	}
	return &arrayCompactCursor{adapter: a, flow: flow}
}

func (c *arrayCompactCursor) load(view semanticstate.StateView[arrayValue]) arrayFlowState {
	if c == nil {
		return nil
	}
	for _, name := range c.adapter.names {
		c.flow[name] = unknownArrayValue()
	}
	view.ForEach(func(id semanticstate.SymbolID, value arrayValue) bool {
		name, ok := c.adapter.environment.Name(id)
		if ok {
			c.flow[name] = value
		}
		return true
	})
	return c.flow
}

func (c *arrayCompactCursor) store(state *semanticstate.State[arrayValue], flow arrayFlowState) {
	if c == nil || state == nil {
		return
	}
	state.Reset()
	for _, name := range c.adapter.names {
		value, ok := flow[name]
		if !ok {
			value = unknownArrayValue()
		}
		id, ok := c.adapter.environment.Symbol(name)
		if ok {
			state.Set(id, value)
		}
	}
}

// arrayCompactLane is the callback-compatible policy used by the indexed
// solver. Keeping the callbacks at this boundary lets the existing Array
// transfer code remain the source of diagnostic/evidence semantics while the
// fixed-point state itself stays in indexed slots.
type arrayCompactLane struct {
	initial             arrayFlowState
	stats               *arrayInterproceduralStats
	visit               func(text string, line int, in arrayFlowState) arrayFlowState
	edgeState           func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState
	stop                func(text string, line int) bool
	reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool
	sourceLines         bool
	// Conditional allocation facts are relational evidence: a positive branch
	// can prove allocation while the sibling branch retains the count source.
	// Preserve that source on the compact positive candidate so a later Select
	// Case can re-establish the same proof after the paths meet. The legacy map
	// walker obtains this behavior from its edge-local map ownership.
	preserveConditionalEvidence bool
}

// walkArrayCFGCompact runs one array policy on the shared semanticstate
// solver. Callback compatibility is intentionally retained at this boundary,
// so diagnostic/evidence code still receives arrayFlowState while joins happen
// directly in indexed slots without union-map allocation.
func walkArrayCFGCompact(ctx context.Context, graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, sourceLines bool) error {
	return walkArrayCFGCompactLanes(ctx, graph, lines, []arrayCompactLane{{
		initial:                     initial,
		visit:                       visit,
		edgeState:                   edgeState,
		sourceLines:                 sourceLines,
		preserveConditionalEvidence: edgeState != nil,
	}})
}

// walkArrayCFGCompactAdvanced is the Array-specific indexed solver boundary
// for source-line and edge-refined paths. It is intentionally kept in this
// package so the generic semanticstate solver does not acquire VBA policy.
func walkArrayCFGCompactAdvanced(ctx context.Context, graph *vbacfg.CFGView, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, stop func(text string, line int) bool, reliableExceptional func(statement *procedureir.Statement, in, out arrayFlowState) bool, sourceLines bool) error {
	return walkArrayCFGCompactLanes(ctx, graph, lines, []arrayCompactLane{{
		initial:                     initial,
		visit:                       visit,
		edgeState:                   edgeState,
		stop:                        stop,
		reliableExceptional:         reliableExceptional,
		sourceLines:                 sourceLines,
		preserveConditionalEvidence: edgeState != nil,
	}})
}

func walkArrayCFGCompactLanes(ctx context.Context, graph *vbacfg.CFGView, lines []string, policies []arrayCompactLane) error {
	if graph == nil {
		return nil
	}
	if len(policies) == 0 {
		return nil
	}
	initials := make([]arrayFlowState, 0, len(policies))
	for _, policy := range policies {
		initials = append(initials, policy.initial)
	}
	adapter := newArrayCompactAdapterForLines(graph, lines, initials...)
	index, err := semanticstate.NewIndexView(*graph)
	if err != nil {
		return err
	}
	lattice := arrayCompactLattice{}
	lanes := make([]semanticstate.Lane[arrayValue], len(policies))
	for laneIndex, policy := range policies {
		policy := policy
		stopped := make(map[semanticstate.BlockOrdinal]bool)
		transferCursor := adapter.newCursor()
		inputCursor := adapter.newCursor()
		outputCursor := adapter.newCursor()
		candidateCursor := adapter.newCursor()
		conditionalSources := map[string]string{}
		lanes[laneIndex] = semanticstate.Lane[arrayValue]{
			Initialize: func(_ context.Context, _ semanticstate.LaneOrdinal, state *semanticstate.State[arrayValue]) error {
				adapter.initialState(state, policy.initial)
				return nil
			},
			Transfer: func(_ context.Context, _ semanticstate.LaneOrdinal, ordinal semanticstate.BlockOrdinal, input semanticstate.StateView[arrayValue], output *semanticstate.State[arrayValue]) error {
				block, ok := graph.BlockAtOrdinal(vbacfg.BlockOrdinal(ordinal))
				if !ok {
					return nil
				}
				in := transferCursor.load(input)
				out, wasStopped := arrayCompactVisitBlockWithStop(lines, block, in, policy.visit, policy.stop, policy.sourceLines)
				if wasStopped {
					stopped[ordinal] = true
					output.Reset()
					return nil
				}
				transferCursor.store(output, out)
				return nil
			},
			EdgeDecision: func(_ context.Context, _ semanticstate.LaneOrdinal, edge semanticstate.Edge, input semanticstate.StateView[arrayValue], output semanticstate.StateView[arrayValue], candidate *semanticstate.State[arrayValue]) (semanticstate.EdgeDisposition, error) {
				// A stopped block must suppress every successor, including the
				// exceptional/uncertain paths that semanticstate initializes from input.
				if stopped[edge.From] {
					return semanticstate.EdgeSuppress, nil
				}
				if edge.Class == vbacfg.EdgeExceptional {
					if policy.sourceLines && policy.reliableExceptional != nil {
						from, fromOK := graph.BlockAtOrdinal(vbacfg.BlockOrdinal(edge.From))
						if fromOK {
							in := inputCursor.load(input)
							out := outputCursor.load(output)
							if policy.reliableExceptional(from.Statement, in, out) {
								candidateCursor.store(candidate, out)
							}
						}
					}
					return semanticstate.EdgePropagate, nil
				}
				// The legacy walker applies edgeState only to normal edges.
				if policy.edgeState == nil || edge.Uncertain {
					return semanticstate.EdgePropagate, nil
				}
				from, ok := graph.BlockAtOrdinal(vbacfg.BlockOrdinal(edge.From))
				if !ok {
					return semanticstate.EdgePropagate, nil
				}
				to, ok := graph.BlockAtOrdinal(vbacfg.BlockOrdinal(edge.To))
				if !ok {
					return semanticstate.EdgePropagate, nil
				}
				flow := candidateCursor.load(candidate.View())
				if policy.preserveConditionalEvidence {
					clear(conditionalSources)
					for name, value := range flow {
						if value.allocationCountSource != "" {
							conditionalSources[name] = value.allocationCountSource
						}
					}
				}
				refined := policy.edgeState(from, vbacfg.Edge{ID: edge.ID, From: from.ID, To: to.ID, Kind: edge.Kind, Class: edge.Class, Uncertain: edge.Uncertain}, flow)
				if policy.preserveConditionalEvidence {
					refined = preserveArrayConditionalEvidence(conditionalSources, refined)
				}
				candidateCursor.store(candidate, refined)
				return semanticstate.EdgePropagate, nil
			},
		}
	}
	solver, err := semanticstate.NewSolver(index, adapter.environment, lattice, lanes)
	if err != nil {
		return err
	}
	_, err = solver.SolveContext(ctx)
	return err
}

// walkArrayCFGCombinedCompact groups compatible Array policies by CFG view and
// advances each group with one indexed solver. Base/runtime lanes share their
// graph and therefore share the scheduler; the source-line VBA227 graph is a
// separate group because its edge set intentionally differs.
//
// The bool result distinguishes a preflight incompatibility (which callers
// may handle with the legacy walker) from a solve-time error (which must not be
// retried after callbacks have already produced findings/evidence).
func walkArrayCFGCombinedCompact(ctx context.Context, lines []string, lanes []arrayCFGWorklistLane) (bool, error) {
	type compactGroup struct {
		graph    *vbacfg.CFGView
		policies []arrayCompactLane
	}
	groups := make([]compactGroup, 0, len(lanes))
	groupByGraph := make(map[*vbacfg.CFGView]int, len(lanes))
	for _, lane := range lanes {
		if lane.Graph == nil || lane.Visit == nil {
			continue
		}
		// Declaration-only lanes have no semantic participants.  Their transfer
		// callbacks still emit scalar/object diagnostics from the variables side
		// table, so compacting an empty state would lose those findings.  Let the
		// caller use the legacy adapter for the whole combined walk instead.
		if len(lane.Initial) == 0 {
			return false, nil
		}
		if _, err := semanticstate.NewIndexView(*lane.Graph); err != nil {
			return false, nil
		}
		groupIndex, ok := groupByGraph[lane.Graph]
		if !ok {
			groupIndex = len(groups)
			groupByGraph[lane.Graph] = groupIndex
			groups = append(groups, compactGroup{graph: lane.Graph})
		}
		groups[groupIndex].policies = append(groups[groupIndex].policies, arrayCompactLane{
			initial:                     lane.Initial,
			stats:                       lane.Stats,
			visit:                       lane.Visit,
			edgeState:                   lane.EdgeState,
			stop:                        lane.Stop,
			reliableExceptional:         lane.ReliableExceptional,
			sourceLines:                 lane.SourceLines,
			preserveConditionalEvidence: lane.EdgeState != nil,
		})
	}
	for _, group := range groups {
		statsSeen := map[*arrayInterproceduralStats]bool{}
		for _, policy := range group.policies {
			if policy.stats == nil || statsSeen[policy.stats] {
				continue
			}
			statsSeen[policy.stats] = true
			policy.stats.addCompactWalk()
		}
		if err := walkArrayCFGCompactLanes(ctx, group.graph, lines, group.policies); err != nil {
			return true, err
		}
	}
	return true, nil
}

// preserveArrayConditionalEvidence keeps a count/length witness attached to a
// compact positive candidate after an edge policy turns it into an allocated
// value.  The witness is deliberately retained only when the policy cleared
// it; conflicting witnesses still meet to an empty string in meetArrayValue.
func preserveArrayConditionalEvidence(sources map[string]string, after arrayFlowState) arrayFlowState {
	for name, source := range sources {
		value, ok := after[name]
		if !ok || value.kind != arrayAllocated || value.allocationCountSource != "" {
			continue
		}
		value.allocationCountSource = source
		after[name] = value
	}
	return after
}

func arrayCompactVisitBlockWithStop(lines []string, block vbacfg.Block, visitIn arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, stop func(text string, line int) bool, sourceLines bool) (arrayFlowState, bool) {
	in := visitIn
	if visit == nil || block.Statement == nil {
		return in, false
	}
	if !sourceLines {
		line := block.Statement.Range.StartLine
		if line == 0 {
			line = block.Range.StartLine
		}
		text := block.Statement.Text
		if strings.TrimSpace(text) == "" && line >= 1 && line <= len(lines) {
			text = normalizedCodeLine(lines[line-1])
		}
		out := visit(text, line, in)
		return out, stop != nil && stop(text, line)
	}
	start := block.Statement.Range.StartLine
	if start == 0 {
		start = block.Range.StartLine
	}
	end := block.Statement.Range.EndLine
	if end < start {
		end = start
	}
	out := in
	if block.Statement.Kind == procedureir.StatementSelect && start >= 1 && start <= len(lines) {
		text := normalizedCodeLine(lines[start-1])
		out = visit(text, start, out)
		return out, stop != nil && stop(text, start)
	}
	if start == end && start >= 1 && start <= len(lines) {
		text := block.Statement.Text
		if strings.TrimSpace(text) == "" {
			text = normalizedCodeLine(lines[start-1])
		}
		out = visit(text, start, out)
		return out, stop != nil && stop(text, start)
	}
	if start >= 1 && end <= len(lines) {
		for line := start; line <= end; line++ {
			text := normalizedCodeLine(lines[line-1])
			if len(text) == 0 {
				continue
			}
			out = visit(text, line, out)
			if stop != nil && stop(text, line) {
				return out, true
			}
		}
		return out, false
	}
	text := block.Statement.Text
	if strings.TrimSpace(text) == "" && start >= 1 && start <= len(lines) {
		text = normalizedCodeLine(lines[start-1])
	}
	out = visit(text, start, out)
	return out, stop != nil && stop(text, start)
}
