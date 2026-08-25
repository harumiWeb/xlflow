package semanticstate

import (
	"context"
	"reflect"
	"strconv"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
)

type maxLattice struct{}

func (maxLattice) Clone(value int) int { return value }
func (maxLattice) Join(dst *int, src int) bool {
	if src <= *dst {
		return false
	}
	*dst = src
	return true
}

func testGraph() cfg.Graph {
	return cfg.Graph{
		Blocks: []cfg.Block{
			{ID: 0, Kind: cfg.BlockEntry},
			{ID: 1, Kind: cfg.BlockStatement},
			{ID: 2, Kind: cfg.BlockStatement},
			{ID: 3, Kind: cfg.BlockNormalExit},
		},
		Edges: []cfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: cfg.EdgeFallthrough, Class: cfg.EdgeNormal},
			{ID: 1, From: 1, To: 2, Kind: cfg.EdgeFallthrough, Class: cfg.EdgeNormal},
			{ID: 2, From: 1, To: 3, Kind: cfg.EdgeError, Class: cfg.EdgeExceptional},
			{ID: 3, From: 2, To: 3, Kind: cfg.EdgeProcedureExit, Class: cfg.EdgeNormal},
		},
		Entry: 0,
	}
}

func TestEnvironmentAndLayoutAreDeterministic(t *testing.T) {
	dense := NewEnvironment([]string{" B ", "a", "A"}, []string{"a", "b"})
	if got := dense.Names(); !reflect.DeepEqual(got, []string{"a", "b"}) {
		t.Fatalf("names = %#v", got)
	}
	if dense.Layout().Representation() != RepresentationDense {
		t.Fatalf("small environment selected %s", dense.Layout().Representation())
	}
	wideNames := make([]string, 100)
	for i := range wideNames {
		wideNames[i] = "name"
		if i > 0 {
			wideNames[i] += string(rune('a'+i%26)) + string(rune('0'+i/26))
		}
	}
	wide := NewEnvironment(wideNames, []string{"namea1"})
	if wide.Layout().Representation() != RepresentationSparse {
		t.Fatalf("wide sparse environment selected %s", wide.Layout().Representation())
	}
	first := NewEnvironment([]string{"z", "a", "m"})
	second := NewEnvironment([]string{"m", "z", "a"})
	if !reflect.DeepEqual(first.Names(), second.Names()) {
		t.Fatalf("name order changed: %#v vs %#v", first.Names(), second.Names())
	}
}

func TestStateJoinReportsOnlyChangedSlotsAndIsIdempotent(t *testing.T) {
	env := NewEnvironment([]string{"a", "b"})
	left := NewState[int](env.Layout())
	right := NewState[int](env.Layout())
	a, _ := env.Symbol("a")
	b, _ := env.Symbol("b")
	left.Set(a, 2)
	right.Set(a, 2)
	right.Set(b, 3)
	changed := []SymbolID{}
	if !left.JoinFrom(right.View(), maxLattice{}, &changed) {
		t.Fatalf("equal slot should not report a change")
	}
	if !reflect.DeepEqual(changed, []SymbolID{b}) {
		t.Fatalf("changed slots = %#v", changed)
	}
	changed = changed[:0]
	if left.JoinFrom(right.View(), maxLattice{}, &changed) {
		t.Fatalf("idempotent join changed state: %#v", changed)
	}
}

func TestSolverDeterministicAndExceptionalEdgesUseInput(t *testing.T) {
	env := NewEnvironment([]string{"value"})
	index, err := NewIndex(testGraph())
	if err != nil {
		t.Fatal(err)
	}
	value, _ := env.Symbol("value")
	makeSolver := func() *Solver[int] {
		solver, err := NewSolver[int](index, env, maxLattice{}, []Lane[int]{{
			Initialize: func(_ context.Context, _ LaneOrdinal, state *State[int]) error {
				state.Set(value, 1)
				return nil
			},
			Transfer: func(_ context.Context, _ LaneOrdinal, block BlockOrdinal, in StateView[int], out *State[int]) error {
				out.CloneFrom(in, nil)
				if block == 1 {
					out.Set(value, 4)
				}
				return nil
			},
		}})
		if err != nil {
			t.Fatal(err)
		}
		solver.RecordOrder = true
		return solver
	}
	one, err := makeSolver().Solve()
	if err != nil {
		t.Fatal(err)
	}
	two, err := makeSolver().Solve()
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(one.Order(), two.Order()) {
		t.Fatalf("worklist order changed: %#v vs %#v", one.Order(), two.Order())
	}
	if got, ok := one.State(2, 0).Value(value); !ok || got != 4 {
		t.Fatalf("normal edge lost transfer value: %d, %v", got, ok)
	}
	if got, ok := one.State(3, 0).Value(value); !ok || got != 4 {
		t.Fatalf("join result = %d, %v", got, ok)
	}
}

func TestSolverSnapshotsInputBeforeSelfEdgePropagation(t *testing.T) {
	env := NewEnvironment([]string{"value"})
	index, err := NewIndex(cfg.Graph{
		Blocks: []cfg.Block{
			{ID: 0, Kind: cfg.BlockEntry},
			{ID: 1, Kind: cfg.BlockStatement},
			{ID: 2, Kind: cfg.BlockNormalExit},
		},
		Edges: []cfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: cfg.EdgeFallthrough, Class: cfg.EdgeNormal},
			{ID: 1, From: 1, To: 1, Kind: cfg.EdgeBranchTrue, Class: cfg.EdgeNormal},
			{ID: 2, From: 1, To: 2, Kind: cfg.EdgeError, Class: cfg.EdgeExceptional},
		},
		Entry: 0,
	})
	if err != nil {
		t.Fatal(err)
	}
	value, _ := env.Symbol("value")
	var exceptionalInputs []int
	solver, err := NewSolver[int](index, env, maxLattice{}, []Lane[int]{{
		Initialize: func(_ context.Context, _ LaneOrdinal, state *State[int]) error {
			state.Set(value, 1)
			return nil
		},
		Transfer: func(_ context.Context, _ LaneOrdinal, block BlockOrdinal, in StateView[int], out *State[int]) error {
			out.CloneFrom(in, nil)
			if block == 1 {
				out.Set(value, 2)
			}
			return nil
		},
		Edge: func(_ context.Context, _ LaneOrdinal, edge Edge, input, _ StateView[int], _ *State[int]) error {
			if edge.From == 1 && edge.Class == cfg.EdgeExceptional {
				got, ok := input.Value(value)
				if !ok {
					t.Fatalf("exceptional edge input lost value")
				}
				exceptionalInputs = append(exceptionalInputs, got)
			}
			return nil
		},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := solver.Solve(); err != nil {
		t.Fatal(err)
	}
	if len(exceptionalInputs) < 2 || exceptionalInputs[0] != 1 {
		t.Fatalf("exceptional edge inputs = %#v, want first input 1 before self-edge update", exceptionalInputs)
	}
}

func TestSolverCancellationDoesNotPublishPartialResult(t *testing.T) {
	env := NewEnvironment([]string{"value"})
	index, err := NewIndex(testGraph())
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	solver, err := NewSolver[int](index, env, maxLattice{}, []Lane[int]{{}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := solver.SolveContext(ctx); err != context.Canceled {
		t.Fatalf("cancellation error = %v", err)
	}
}

func BenchmarkStateJoinRepresentations(b *testing.B) {
	cases := []struct {
		name    string
		symbols int
		touched int
	}{
		{name: "dense", symbols: 32, touched: 32},
		{name: "hybrid-dense", symbols: 128, touched: 64},
		{name: "sparse", symbols: 4096, touched: 16},
	}
	for _, tc := range cases {
		b.Run(tc.name, func(b *testing.B) {
			envNames := make([]string, tc.symbols)
			for i := range envNames {
				envNames[i] = "symbol" + strconv.Itoa(i)
			}
			touched := envNames[:tc.touched]
			env := NewEnvironment(envNames, touched)
			left := NewState[int](env.Layout())
			right := NewState[int](env.Layout())
			ids := make([]SymbolID, tc.touched)
			for i := 0; i < tc.touched; i++ {
				ids[i], _ = env.Symbol(envNames[i])
				id := ids[i]
				left.Set(id, i)
				right.Set(id, i+1)
			}
			lattice := maxLattice{}
			changed := make([]SymbolID, 0, tc.touched)
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				changed = changed[:0]
				left.JoinFrom(right.View(), lattice, &changed)
				b.StopTimer()
				left.Reset()
				for symbol, id := range ids {
					left.Set(id, symbol)
				}
				b.StartTimer()
			}
		})
	}
}
