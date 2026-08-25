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

func (arrayCompactLattice) Clone(value arrayValue) arrayValue {
	value.dimensions = append([]arrayDimension(nil), value.dimensions...)
	value.preserveShape = append([]arrayDimension(nil), value.preserveShape...)
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
}

func newArrayCompactAdapter(initial arrayFlowState) arrayCompactAdapter {
	names := make([]string, 0, len(initial))
	for name := range initial {
		names = append(names, name)
	}
	return arrayCompactAdapter{environment: semanticstate.NewEnvironment(names, names)}
}

func (a arrayCompactAdapter) toFlow(view semanticstate.StateView[arrayValue]) arrayFlowState {
	flow := make(arrayFlowState, view.Len())
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
	for name, value := range flow {
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

// walkArrayCFGCompact runs one array policy on the shared semanticstate
// solver. Callback compatibility is intentionally retained at this boundary,
// so diagnostic/evidence code still receives arrayFlowState while joins happen
// directly in indexed slots without union-map allocation.
func walkArrayCFGCompact(ctx context.Context, graph *vbacfg.Graph, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, edgeState func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState, sourceLines bool) error {
	if graph == nil {
		return nil
	}
	adapter := newArrayCompactAdapter(initial)
	index, err := semanticstate.NewIndex(*graph)
	if err != nil {
		return err
	}
	blocks := make(map[vbacfg.BlockID]vbacfg.Block, len(graph.Blocks))
	for _, block := range graph.Blocks {
		blocks[block.ID] = block
	}
	lattice := arrayCompactLattice{}
	lane := semanticstate.Lane[arrayValue]{
		Initialize: func(_ context.Context, _ semanticstate.LaneOrdinal, state *semanticstate.State[arrayValue]) error {
			adapter.initialState(state, initial)
			return nil
		},
		Transfer: func(_ context.Context, _ semanticstate.LaneOrdinal, ordinal semanticstate.BlockOrdinal, input semanticstate.StateView[arrayValue], output *semanticstate.State[arrayValue]) error {
			indexed, ok := index.Block(ordinal)
			if !ok {
				return nil
			}
			block, ok := blocks[indexed.ID]
			if !ok {
				return nil
			}
			in := adapter.toFlow(input)
			out := arrayCompactVisitBlock(lines, block, in, visit, sourceLines)
			adapter.fromFlow(output, out)
			return nil
		},
		Edge: func(_ context.Context, _ semanticstate.LaneOrdinal, edge semanticstate.Edge, _ semanticstate.StateView[arrayValue], _ semanticstate.StateView[arrayValue], candidate *semanticstate.State[arrayValue]) error {
			// The legacy walker applies edgeState only to normal edges. Exceptional
			// and uncertain edges carry the predecessor input state and must not
			// receive normal-edge refinements (guards/Select facts).
			if edgeState == nil || edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				return nil
			}
			fromIndexed, ok := index.Block(edge.From)
			if !ok {
				return nil
			}
			from, ok := blocks[fromIndexed.ID]
			if !ok {
				return nil
			}
			toIndexed, ok := index.Block(edge.To)
			if !ok {
				return nil
			}
			to, ok := blocks[toIndexed.ID]
			if !ok {
				return nil
			}
			flow := adapter.toFlow(candidate.View())
			refined := edgeState(from, vbacfg.Edge{ID: edge.ID, From: from.ID, To: to.ID, Kind: edge.Kind, Class: edge.Class, Uncertain: edge.Uncertain}, flow)
			adapter.fromFlow(candidate, refined)
			return nil
		},
	}
	solver, err := semanticstate.NewSolver(index, adapter.environment, lattice, []semanticstate.Lane[arrayValue]{lane})
	if err != nil {
		return err
	}
	_, err = solver.SolveContext(ctx)
	return err
}

func arrayCompactVisitBlock(lines []string, block vbacfg.Block, in arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState, sourceLines bool) arrayFlowState {
	if visit == nil || block.Statement == nil {
		return in
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
		return visit(text, line, in)
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
		return visit(normalizedCodeLine(lines[start-1]), start, out)
	}
	if start == end && start >= 1 && start <= len(lines) {
		text := block.Statement.Text
		if strings.TrimSpace(text) == "" {
			text = normalizedCodeLine(lines[start-1])
		}
		return visit(text, start, out)
	}
	if start >= 1 && end <= len(lines) {
		for line := start; line <= end; line++ {
			text := normalizedCodeLine(lines[line-1])
			if len(text) == 0 {
				continue
			}
			out = visit(text, line, out)
		}
		return out
	}
	text := block.Statement.Text
	if strings.TrimSpace(text) == "" && start >= 1 && start <= len(lines) {
		text = normalizedCodeLine(lines[start-1])
	}
	return visit(text, start, out)
}
