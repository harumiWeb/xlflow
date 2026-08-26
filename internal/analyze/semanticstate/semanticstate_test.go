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

// taggedScalar models the shape used by compact semantic domains: absence is
// represented by a present slot carrying scalarAbsent, while an unknown fact
// is an explicit scalar value. Keeping both states in the test makes it
// impossible for a future State/JoinFrom change to conflate a missing slot
// with a propagated unknown.
type scalarTag uint8

const (
	scalarAbsent scalarTag = iota
	scalarUnknown
	scalarKnown
)

type taggedScalar struct {
	tag   scalarTag
	value string
}

type taggedScalarLattice struct{}

func (taggedScalarLattice) Clone(value taggedScalar) taggedScalar { return value }

func (taggedScalarLattice) Join(dst *taggedScalar, src taggedScalar) bool {
	if dst == nil {
		return false
	}
	merged := *dst
	switch {
	case dst.tag == scalarUnknown || src.tag == scalarUnknown:
		merged = taggedScalar{tag: scalarUnknown}
	case dst.tag == scalarAbsent:
		merged = src
	case src.tag == scalarAbsent:
		// Absence is the lattice bottom for this slot and must not erase a
		// fact learned on another path.
	case dst.tag == scalarKnown && src.tag == scalarKnown:
		if dst.value == src.value {
			return false
		}
		merged = taggedScalar{tag: scalarUnknown}
	default:
		merged = src
	}
	if merged == *dst {
		return false
	}
	*dst = merged
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

func TestStatePreservesExplicitAbsenceAndUnknownScalarFacts(t *testing.T) {
	env := NewEnvironment([]string{"object", "url"})
	object, _ := env.Symbol("object")
	url, _ := env.Symbol("url")
	state := NewState[taggedScalar](env.Layout())
	if !state.Set(object, taggedScalar{tag: scalarAbsent}) {
		t.Fatal("failed to store explicit absence")
	}
	if !state.Set(url, taggedScalar{tag: scalarUnknown}) {
		t.Fatal("failed to store explicit unknown")
	}
	if !state.View().Has(object) || !state.View().Has(url) {
		t.Fatal("explicit scalar facts must remain present slots")
	}
	if got, ok := state.View().Value(object); !ok || got.tag != scalarAbsent {
		t.Fatalf("absence fact = %#v, %v", got, ok)
	}
	if got, ok := state.View().Value(url); !ok || got.tag != scalarUnknown {
		t.Fatalf("unknown fact = %#v, %v", got, ok)
	}
	if got := state.View().IDs(); !reflect.DeepEqual(got, []SymbolID{object, url}) {
		t.Fatalf("scalar fact IDs = %#v, want object/url order", got)
	}

	clone := NewState[taggedScalar](env.Layout())
	clone.CloneFrom(state.View(), taggedScalarLattice{}.Clone)
	if got, ok := clone.View().Value(object); !ok || got.tag != scalarAbsent {
		t.Fatalf("cloned absence fact = %#v, %v", got, ok)
	}
	if got, ok := clone.View().Value(url); !ok || got.tag != scalarUnknown {
		t.Fatalf("cloned unknown fact = %#v, %v", got, ok)
	}
}

func TestTaggedScalarJoinReportsChangedSlotsAndIsIdempotent(t *testing.T) {
	env := NewEnvironment([]string{"object", "url", "credential"})
	object, _ := env.Symbol("object")
	url, _ := env.Symbol("url")
	credential, _ := env.Symbol("credential")
	lattice := taggedScalarLattice{}

	tests := []struct {
		name       string
		left       taggedScalar
		right      taggedScalar
		want       taggedScalar
		wantChange bool
	}{
		{name: "absence learns known", left: taggedScalar{tag: scalarAbsent}, right: taggedScalar{tag: scalarKnown, value: "url"}, want: taggedScalar{tag: scalarKnown, value: "url"}, wantChange: true},
		{name: "known ignores absence", left: taggedScalar{tag: scalarKnown, value: "url"}, right: taggedScalar{tag: scalarAbsent}, want: taggedScalar{tag: scalarKnown, value: "url"}},
		{name: "equal known is stable", left: taggedScalar{tag: scalarKnown, value: "url"}, right: taggedScalar{tag: scalarKnown, value: "url"}, want: taggedScalar{tag: scalarKnown, value: "url"}},
		{name: "conflicting known becomes unknown", left: taggedScalar{tag: scalarKnown, value: "url"}, right: taggedScalar{tag: scalarKnown, value: "other"}, want: taggedScalar{tag: scalarUnknown}, wantChange: true},
		{name: "unknown absorbs known", left: taggedScalar{tag: scalarKnown, value: "url"}, right: taggedScalar{tag: scalarUnknown}, want: taggedScalar{tag: scalarUnknown}, wantChange: true},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			left := NewState[taggedScalar](env.Layout())
			right := NewState[taggedScalar](env.Layout())
			left.Set(url, test.left)
			right.Set(url, test.right)
			changed := make([]SymbolID, 0, 1)
			if got := left.JoinFrom(right.View(), lattice, &changed); got != test.wantChange {
				t.Fatalf("JoinFrom changed = %v, want %v (IDs %#v)", got, test.wantChange, changed)
			}
			if got, ok := left.View().Value(url); !ok || got != test.want {
				t.Fatalf("joined scalar = %#v, %v; want %#v", got, ok, test.want)
			}
			if test.wantChange && !reflect.DeepEqual(changed, []SymbolID{url}) {
				t.Fatalf("changed IDs = %#v, want [%d]", changed, url)
			}
			changed = changed[:0]
			if left.JoinFrom(right.View(), lattice, &changed) {
				t.Fatalf("repeating join changed state: %#v", changed)
			}
			if len(changed) != 0 {
				t.Fatalf("repeating join reported IDs: %#v", changed)
			}
		})
	}

	// A join must report every changed slot in deterministic SymbolID order,
	// including a slot that was previously explicit absence.
	left := NewState[taggedScalar](env.Layout())
	right := NewState[taggedScalar](env.Layout())
	left.Set(object, taggedScalar{tag: scalarAbsent})
	right.Set(object, taggedScalar{tag: scalarKnown, value: "object"})
	right.Set(credential, taggedScalar{tag: scalarUnknown})
	changed := make([]SymbolID, 0, 2)
	if !left.JoinFrom(right.View(), lattice, &changed) {
		t.Fatal("multi-slot scalar join did not report changes")
	}
	if !reflect.DeepEqual(changed, []SymbolID{credential, object}) {
		t.Fatalf("multi-slot changed IDs = %#v, want [%d %d]", changed, credential, object)
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

func TestSolverUncertainEdgesUseInputWhileNormalEdgesUseOutput(t *testing.T) {
	index, err := NewIndex(cfg.Graph{
		Blocks: []cfg.Block{
			{ID: 0, Kind: cfg.BlockEntry},
			{ID: 1, Kind: cfg.BlockStatement},
			{ID: 2, Kind: cfg.BlockNormalExit},
			{ID: 3, Kind: cfg.BlockNormalExit},
		},
		Edges: []cfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: cfg.EdgeFallthrough, Class: cfg.EdgeNormal},
			{ID: 1, From: 1, To: 2, Kind: cfg.EdgeUnknown, Class: cfg.EdgeNormal, Uncertain: true},
			{ID: 2, From: 1, To: 3, Kind: cfg.EdgeFallthrough, Class: cfg.EdgeNormal},
		},
		Entry: 0,
	})
	if err != nil {
		t.Fatal(err)
	}
	env := NewEnvironment([]string{"value"})
	value, _ := env.Symbol("value")
	gotCandidates := map[cfg.EdgeID]int{}
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
		Edge: func(_ context.Context, _ LaneOrdinal, edge Edge, _ StateView[int], _ StateView[int], candidate *State[int]) error {
			got, ok := candidate.View().Value(value)
			if !ok {
				t.Fatalf("edge %d candidate lost value", edge.ID)
			}
			gotCandidates[edge.ID] = got
			return nil
		},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := solver.Solve(); err != nil {
		t.Fatal(err)
	}
	if gotCandidates[1] != 1 {
		t.Fatalf("uncertain edge candidate = %d, want predecessor input 1", gotCandidates[1])
	}
	if gotCandidates[2] != 2 {
		t.Fatalf("normal edge candidate = %d, want transfer output 2", gotCandidates[2])
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

func TestSolverEdgeDecisionCanSuppressPropagation(t *testing.T) {
	env := NewEnvironment([]string{"value"})
	index, err := NewIndex(cfg.Graph{
		Blocks: []cfg.Block{
			{ID: 0, Kind: cfg.BlockEntry},
			{ID: 1, Kind: cfg.BlockStatement},
			{ID: 2, Kind: cfg.BlockNormalExit},
		},
		Edges: []cfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: cfg.EdgeFallthrough, Class: cfg.EdgeNormal},
			{ID: 1, From: 1, To: 2, Kind: cfg.EdgeProcedureExit, Class: cfg.EdgeNormal},
		},
		Entry: 0,
	})
	if err != nil {
		t.Fatal(err)
	}
	value, _ := env.Symbol("value")
	solver, err := NewSolver[int](index, env, maxLattice{}, []Lane[int]{{
		Initialize: func(_ context.Context, _ LaneOrdinal, state *State[int]) error {
			state.Set(value, 1)
			return nil
		},
		Transfer: func(_ context.Context, _ LaneOrdinal, _ BlockOrdinal, in StateView[int], out *State[int]) error {
			out.CloneFrom(in, nil)
			return nil
		},
		EdgeDecision: func(_ context.Context, _ LaneOrdinal, edge Edge, _ StateView[int], _ StateView[int], _ *State[int]) (EdgeDisposition, error) {
			if edge.ID == 0 {
				return EdgeSuppress, nil
			}
			return EdgePropagate, nil
		},
	}})
	if err != nil {
		t.Fatal(err)
	}
	result, err := solver.Solve()
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := result.State(1, 0).Value(value); ok {
		t.Fatal("suppressed edge unexpectedly reached destination")
	}
	if _, ok := result.State(2, 0).Value(value); ok {
		t.Fatal("suppressed predecessor unexpectedly reached successor")
	}
}

func TestSolverEdgeDecisionUsesInputForExceptionalAndUncertainEdges(t *testing.T) {
	env := NewEnvironment([]string{"value"})
	index, err := NewIndex(cfg.Graph{
		Blocks: []cfg.Block{
			{ID: 0, Kind: cfg.BlockEntry},
			{ID: 1, Kind: cfg.BlockStatement},
			{ID: 2, Kind: cfg.BlockStatement},
			{ID: 3, Kind: cfg.BlockNormalExit},
			{ID: 4, Kind: cfg.BlockNormalExit},
		},
		Edges: []cfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: cfg.EdgeFallthrough, Class: cfg.EdgeNormal},
			{ID: 1, From: 1, To: 2, Kind: cfg.EdgeError, Class: cfg.EdgeExceptional},
			{ID: 2, From: 1, To: 3, Kind: cfg.EdgeBranchTrue, Class: cfg.EdgeNormal, Uncertain: true},
			{ID: 3, From: 1, To: 4, Kind: cfg.EdgeProcedureExit, Class: cfg.EdgeNormal},
		},
		Entry: 0,
	})
	if err != nil {
		t.Fatal(err)
	}
	value, _ := env.Symbol("value")
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
		EdgeDecision: func(_ context.Context, _ LaneOrdinal, edge Edge, input, output StateView[int], candidate *State[int]) (EdgeDisposition, error) {
			if edge.Class != cfg.EdgeExceptional && !edge.Uncertain {
				return EdgePropagate, nil
			}
			inputValue, inputOK := input.Value(value)
			outputValue, outputOK := output.Value(value)
			candidateValue, candidateOK := candidate.View().Value(value)
			if !inputOK || !outputOK || !candidateOK || inputValue != 1 || outputValue != 4 || candidateValue != 1 {
				t.Fatalf("edge %d states: input=(%d,%v) output=(%d,%v) candidate=(%d,%v)", edge.ID, inputValue, inputOK, outputValue, outputOK, candidateValue, candidateOK)
			}
			return EdgePropagate, nil
		},
	}})
	if err != nil {
		t.Fatal(err)
	}
	result, err := solver.Solve()
	if err != nil {
		t.Fatal(err)
	}
	for _, block := range []BlockOrdinal{2, 3} {
		if got, ok := result.State(block, 0).Value(value); !ok || got != 1 {
			t.Fatalf("edge to block %d joined output instead of input: %d, %v", block, got, ok)
		}
	}
	if got, ok := result.State(4, 0).Value(value); !ok || got != 4 {
		t.Fatalf("normal edge did not join output: %d, %v", got, ok)
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

func TestSolverCancellationDuringTransferDoesNotPublishPartialResult(t *testing.T) {
	env := NewEnvironment([]string{"value"})
	index, err := NewIndex(testGraph())
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	value, _ := env.Symbol("value")
	solver, err := NewSolver[int](index, env, maxLattice{}, []Lane[int]{{
		Initialize: func(_ context.Context, _ LaneOrdinal, state *State[int]) error {
			state.Set(value, 1)
			return nil
		},
		Transfer: func(_ context.Context, _ LaneOrdinal, _ BlockOrdinal, in StateView[int], out *State[int]) error {
			out.CloneFrom(in, nil)
			cancel()
			return nil
		},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := solver.SolveContext(ctx); err != context.Canceled {
		t.Fatalf("transfer cancellation error = %v", err)
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
