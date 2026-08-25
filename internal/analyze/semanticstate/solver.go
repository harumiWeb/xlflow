package semanticstate

import (
	"container/heap"
	"context"
	"fmt"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
)

// WorkItem is the deterministic scheduling key used by the solver.
type WorkItem struct {
	Block BlockOrdinal
	Lane  LaneOrdinal
}

// Stats describes one fixed-point run. The counters are intentionally
// procedure-local and do not change existing analyzer performance counters.
type Stats struct {
	Transfers int
	Joins     int
	Requeues  int
}

// Result owns an immutable snapshot of converged states. It remains valid
// after the Solver is reused for another procedure invocation.
type Result[T any] struct {
	layout Layout
	lanes  int
	blocks int
	states []State[T]
	order  []WorkItem
	stats  Stats
}

// State returns a read-only view of a converged block/lane state.
func (r Result[T]) State(block BlockOrdinal, lane LaneOrdinal) StateView[T] {
	if int(block) >= r.blocks || int(lane) >= r.lanes {
		return StateView[T]{}
	}
	return r.states[int(block)*r.lanes+int(lane)].View()
}

// Order returns work items in the order they were processed when Solver.RecordOrder
// was enabled. Repeated items are retained because they are useful for
// deterministic regression tests; production adapters normally leave the
// trace disabled.
func (r Result[T]) Order() []WorkItem { return append([]WorkItem(nil), r.order...) }

// Stats returns fixed-point counters for the run.
func (r Result[T]) Stats() Stats { return r.stats }

// Solver runs one finite-lattice analysis for one procedure revision. A
// Solver is intentionally single-use at a time; after SolveContext returns it
// may be reused sequentially and will retain its scratch backing arrays.
type Solver[T any] struct {
	index           *Index
	layout          Layout
	lattice         Lattice[T]
	lanes           []Lane[T]
	states          []State[T]
	transferScratch []State[T]
	edgeScratch     []State[T]
	inputScratch    []State[T]
	changed         []SymbolID
	queue           workQueue
	queued          []bool
	// RecordOrder retains the processed work-item trace in Result.Order for
	// deterministic diagnostics and solver tests. Production adapters leave it
	// disabled so loop-heavy procedures do not retain an otherwise unused
	// allocation proportional to transfer count.
	RecordOrder bool
}

// NewSolver constructs a procedure-local solver. The index and environment
// must describe the same revision; this package keeps no persistent cache.
func NewSolver[T any](index *Index, environment Environment, lattice Lattice[T], lanes []Lane[T]) (*Solver[T], error) {
	if index == nil {
		return nil, fmt.Errorf("semanticstate: nil CFG index")
	}
	if lattice == nil {
		return nil, fmt.Errorf("semanticstate: nil lattice")
	}
	if len(lanes) == 0 {
		return nil, fmt.Errorf("semanticstate: at least one lane is required")
	}
	if len(lanes) > int(^LaneOrdinal(0)) {
		return nil, fmt.Errorf("semanticstate: too many lanes: %d", len(lanes))
	}
	s := &Solver[T]{
		index:   index,
		layout:  environment.Layout(),
		lattice: lattice,
		lanes:   append([]Lane[T](nil), lanes...),
		queued:  make([]bool, index.BlockCount()*len(lanes)),
	}
	s.allocateStates()
	return s, nil
}

func (s *Solver[T]) allocateStates() {
	count := s.index.BlockCount() * len(s.lanes)
	if len(s.states) != count {
		s.states = make([]State[T], count)
	}
	if len(s.transferScratch) != len(s.lanes) {
		s.transferScratch = make([]State[T], len(s.lanes))
		s.edgeScratch = make([]State[T], len(s.lanes))
		s.inputScratch = make([]State[T], len(s.lanes))
	}
	for i := range s.states {
		if s.states[i].Layout() != s.layout {
			s.states[i] = newState[T](s.layout)
		} else {
			s.states[i].Reset()
		}
	}
	for i := range s.transferScratch {
		if s.transferScratch[i].Layout() != s.layout {
			s.transferScratch[i] = newState[T](s.layout)
			s.edgeScratch[i] = newState[T](s.layout)
			s.inputScratch[i] = newState[T](s.layout)
		} else {
			s.transferScratch[i].Reset()
			s.edgeScratch[i].Reset()
			s.inputScratch[i].Reset()
		}
	}
	clear(s.queued)
	s.queue = s.queue[:0]
	s.changed = s.changed[:0]
}

// Solve runs a context-independent fixed point.
func (s *Solver[T]) Solve() (Result[T], error) {
	return s.SolveContext(context.Background())
}

// SolveContext runs all lanes to a finite fixed point. It checks cancellation
// before queue operations, transfers, and edge propagation. Cancellation never
// returns a partial Result.
func (s *Solver[T]) SolveContext(ctx context.Context) (Result[T], error) {
	if s == nil {
		return Result[T]{}, fmt.Errorf("semanticstate: nil solver")
	}
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return Result[T]{}, err
	}
	s.allocateStates()
	entry := s.index.Entry()
	for laneIndex, lane := range s.lanes {
		if err := ctx.Err(); err != nil {
			return Result[T]{}, err
		}
		if lane.Initialize != nil {
			if err := lane.Initialize(ctx, LaneOrdinal(laneIndex), &s.states[s.stateIndex(entry, LaneOrdinal(laneIndex))]); err != nil {
				return Result[T]{}, err
			}
		}
		s.enqueue(WorkItem{Block: entry, Lane: LaneOrdinal(laneIndex)})
	}
	stats := Stats{}
	var order []WorkItem
	if s.RecordOrder {
		order = make([]WorkItem, 0, s.index.BlockCount()*len(s.lanes))
	}
	for len(s.queue) != 0 {
		if err := ctx.Err(); err != nil {
			return Result[T]{}, err
		}
		item := heap.Pop(&s.queue).(WorkItem)
		s.queued[s.stateIndex(item.Block, item.Lane)] = false
		if s.RecordOrder {
			order = append(order, item)
		}
		stats.Transfers++
		lane := s.lanes[item.Lane]
		input := &s.states[s.stateIndex(item.Block, item.Lane)]
		output := &s.transferScratch[item.Lane]
		output.Reset()
		if lane.Transfer != nil {
			if err := lane.Transfer(ctx, item.Lane, item.Block, input.View(), output); err != nil {
				return Result[T]{}, err
			}
		} else {
			output.CloneFrom(input.View(), s.lattice.Clone)
		}
		if err := ctx.Err(); err != nil {
			return Result[T]{}, err
		}
		inputView := input.View()
		for _, edge := range s.index.outgoing[item.Block] {
			if edge.To != item.Block {
				continue
			}
			snapshot := &s.inputScratch[item.Lane]
			snapshot.CloneFrom(inputView, s.lattice.Clone)
			inputView = snapshot.View()
			break
		}
		for edgeIndex, edge := range s.index.outgoing[item.Block] {
			if edgeIndex&0xff == 0 {
				if err := ctx.Err(); err != nil {
					return Result[T]{}, err
				}
			}
			candidate := &s.edgeScratch[item.Lane]
			base := output.View()
			if edge.Class == cfg.EdgeExceptional || edge.Uncertain {
				base = inputView
			}
			candidate.CloneFrom(base, s.lattice.Clone)
			if lane.Edge != nil {
				if err := lane.Edge(ctx, item.Lane, edge, inputView, output.View(), candidate); err != nil {
					return Result[T]{}, err
				}
			}
			stats.Joins++
			s.changed = s.changed[:0]
			destination := &s.states[s.stateIndex(edge.To, item.Lane)]
			if !destination.JoinFrom(candidate.View(), s.lattice, &s.changed) {
				continue
			}
			stats.Requeues++
			s.enqueue(WorkItem{Block: edge.To, Lane: item.Lane})
		}
	}
	result := Result[T]{layout: s.layout, lanes: len(s.lanes), blocks: s.index.BlockCount(), order: order, stats: stats}
	result.states = make([]State[T], len(s.states))
	for i := range s.states {
		result.states[i] = newState[T](s.layout)
		result.states[i].CloneFrom(s.states[i].View(), s.lattice.Clone)
	}
	return result, nil
}

func (s *Solver[T]) stateIndex(block BlockOrdinal, lane LaneOrdinal) int {
	return int(block)*len(s.lanes) + int(lane)
}

func (s *Solver[T]) enqueue(item WorkItem) {
	index := s.stateIndex(item.Block, item.Lane)
	if s.queued[index] {
		return
	}
	s.queued[index] = true
	heap.Push(&s.queue, item)
}

type workQueue []WorkItem

func (q workQueue) Len() int { return len(q) }

func (q workQueue) Less(i, j int) bool {
	if q[i].Block != q[j].Block {
		return q[i].Block < q[j].Block
	}
	return q[i].Lane < q[j].Lane
}

func (q workQueue) Swap(i, j int) { q[i], q[j] = q[j], q[i] }

func (q *workQueue) Push(value any) { *q = append(*q, value.(WorkItem)) }

func (q *workQueue) Pop() any {
	old := *q
	last := len(old) - 1
	value := old[last]
	*q = old[:last]
	return value
}
