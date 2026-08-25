package semanticquery

import (
	"context"
	"errors"
	"fmt"
	"sync"
	"sync/atomic"
	"testing"
)

type testMetrics struct {
	hits, misses, invalidated, recomputed atomic.Uint64
}

func (m *testMetrics) RecordSemanticQueryHit()                           { m.hits.Add(1) }
func (m *testMetrics) RecordSemanticQueryMiss()                          { m.misses.Add(1) }
func (m *testMetrics) RecordSemanticQueryInvalidatedProcedures(n uint64) { m.invalidated.Add(n) }
func (m *testMetrics) RecordSemanticQueryRecomputedKernel()              { m.recomputed.Add(1) }

func TestHashAndKeyAreStable(t *testing.T) {
	if Hash("a", "bc") == Hash("ab", "c") {
		t.Fatal("length-prefixed hashes collided")
	}
	key := Key{Procedure: "module\x00proc", Fingerprint: Hash("body"), Kernel: "array", Config: "cfg", Capability: "cap"}
	if key.String() == "" || key.String() != (Key{Procedure: "module\x00proc", Fingerprint: Hash("body"), Kernel: "array", Config: "cfg", Capability: "cap"}).String() {
		t.Fatal("query key is not stable")
	}
}

func TestEvaluateSingleFlightAndReuse(t *testing.T) {
	metrics := &testMetrics{}
	store := New(Options{Metrics: metrics})
	revision := store.Begin("r1")
	defer revision.Close()
	key := Key{Procedure: "M.P", Fingerprint: Hash("body"), Kernel: "array"}
	var builds atomic.Int32
	var wg sync.WaitGroup
	values := make([]int, 8)
	for i := range values {
		wg.Add(1)
		go func(index int) {
			defer wg.Done()
			value, _, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) {
				builds.Add(1)
				return 42, nil
			})
			if err != nil {
				t.Errorf("evaluate: %v", err)
				return
			}
			values[index] = value
		}(i)
	}
	wg.Wait()
	if builds.Load() != 1 {
		t.Fatalf("builds = %d, want 1", builds.Load())
	}
	for _, value := range values {
		if value != 42 {
			t.Fatalf("value = %d, want 42", value)
		}
	}
	if _, hit, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) { return 0, errors.New("must not build") }); err != nil || !hit {
		t.Fatalf("reuse = hit %v, err %v", hit, err)
	}
	if metrics.recomputed.Load() != 1 || metrics.hits.Load() == 0 || metrics.misses.Load() == 0 {
		t.Fatalf("metrics = hits %d misses %d recomputed %d", metrics.hits.Load(), metrics.misses.Load(), metrics.recomputed.Load())
	}
}

func TestEvaluateRechecksDependencyFingerprintsOnHit(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("r1")
	defer revision.Close()
	key := Key{Procedure: "M.P", Fingerprint: Hash("body"), Kernel: "array"}
	firstDependency := Key{Procedure: "M.Callee", Fingerprint: Hash("v1"), Kernel: "effect"}
	secondDependency := Key{Procedure: "M.Callee", Fingerprint: Hash("v2"), Kernel: "effect"}
	if value, hit, err := Evaluate(context.Background(), revision, key, []Key{firstDependency}, func(context.Context) (int, error) { return 1, nil }); err != nil || hit || value != 1 {
		t.Fatalf("first evaluation = %d, hit %v, err %v", value, hit, err)
	}
	value, hit, err := Evaluate(context.Background(), revision, key, []Key{secondDependency}, func(context.Context) (int, error) { return 2, nil })
	if err != nil || hit || value != 2 {
		t.Fatalf("dependency change = %d, hit %v, err %v", value, hit, err)
	}
}

func TestCanceledBuildIsRetryable(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("r1")
	defer revision.Close()
	key := Key{Procedure: "M.P", Fingerprint: Hash("body"), Kernel: "http"}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, _, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) { return 0, ctx.Err() }); !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled error = %v", err)
	}
	value, hit, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) { return 7, nil })
	if err != nil || hit || value != 7 {
		t.Fatalf("retry = %d, hit %v, err %v", value, hit, err)
	}
}

func TestPanickingBuildDoesNotPoisonKey(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("r1")
	defer revision.Close()
	key := Key{Procedure: "M.P", Fingerprint: Hash("panic"), Kernel: "array"}
	func() {
		defer func() {
			if recover() == nil {
				t.Fatal("panic was swallowed")
			}
		}()
		_, _, _ = Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) {
			panic("boom")
		})
	}()
	value, hit, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) { return 9, nil })
	if err != nil || hit || value != 9 {
		t.Fatalf("retry after panic = %d, hit %v, err %v", value, hit, err)
	}
}

func TestInvalidateTraversesReverseDependenciesDeterministically(t *testing.T) {
	metrics := &testMetrics{}
	store := New(Options{Metrics: metrics})
	revision := store.Begin("r1")
	defer revision.Close()
	callee := Key{Procedure: "M.Callee", Fingerprint: Hash("callee"), Kernel: "effect"}
	caller := Key{Procedure: "M.Caller", Fingerprint: Hash("caller"), Kernel: "array"}
	project := Key{Procedure: "project", Fingerprint: Hash("project"), Kernel: "summary"}
	for _, item := range []struct {
		key   Key
		value int
	}{{callee, 1}, {caller, 2}, {project, 3}} {
		if _, _, err := Evaluate(context.Background(), revision, item.key, nil, func(context.Context) (int, error) { return item.value, nil }); err != nil {
			t.Fatal(err)
		}
	}
	store.RecordDependencies(caller, callee)
	store.RecordDependencies(project, caller)
	got := store.Invalidate(callee)
	want := []string{"M.Callee", "M.Caller", "project"}
	if len(got) != len(want) {
		t.Fatalf("invalidated = %#v, want %#v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("invalidated = %#v, want %#v", got, want)
		}
	}
	if metrics.invalidated.Load() != uint64(len(want)) {
		t.Fatalf("invalidated telemetry = %d", metrics.invalidated.Load())
	}
}

func TestInvalidateProceduresFindsPrefixAndDependents(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("r1")
	defer revision.Close()
	callee := Key{Procedure: "C:\\src\\Main.bas::Main::Callee::0", Fingerprint: Hash("callee"), Kernel: "array"}
	caller := Key{Procedure: "C:\\src\\Main.bas::Main::Caller::1", Fingerprint: Hash("caller"), Kernel: "array"}
	project := Key{Procedure: "project", Fingerprint: Hash("project"), Kernel: "summary"}
	for _, key := range []Key{callee, caller, project} {
		if _, _, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) { return 1, nil }); err != nil {
			t.Fatal(err)
		}
	}
	store.RecordDependencies(caller, callee)
	store.RecordDependencies(project, caller)
	got := store.InvalidateProcedures(`C:\src\Main.bas::Main::Caller`)
	want := []string{caller.Procedure, project.Procedure}
	if len(got) != len(want) {
		t.Fatalf("invalidated = %#v, want %#v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("invalidated = %#v, want %#v", got, want)
		}
	}
}

func TestClosedRevisionRejectsQueries(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("r1")
	revision.Close()
	_, _, err := Evaluate(context.Background(), revision, Key{Kernel: "x"}, nil, func(context.Context) (int, error) { return 1, nil })
	if !errors.Is(err, ErrRevisionClosed) {
		t.Fatalf("error = %v, want ErrRevisionClosed", err)
	}
}

func TestEvictionBoundsDependencyMetadata(t *testing.T) {
	store := New(Options{MaxEntries: 2})
	revision := store.Begin("r1")
	defer revision.Close()
	for index := 0; index < 32; index++ {
		key := Key{Procedure: "M.P", Fingerprint: Hash(fmt.Sprintf("%d", index)), Kernel: "array"}
		dependency := Key{Procedure: "M.Module", Fingerprint: Hash(fmt.Sprintf("dep-%d", index)), Kernel: "module"}
		if _, _, err := Evaluate(context.Background(), revision, key, []Key{dependency}, func(context.Context) (int, error) { return index, nil }); err != nil {
			t.Fatal(err)
		}
		store.Invalidate(key)
	}
	if len(store.entries) > 2 || len(store.keys) > 4 || len(store.reverse) > 4 || len(store.order) > 4 {
		t.Fatalf("dependency metadata exceeded bounds: entries=%d keys=%d reverse=%d order=%d", len(store.entries), len(store.keys), len(store.reverse), len(store.order))
	}
}

func TestRecordDependenciesBoundsMetadata(t *testing.T) {
	store := New(Options{MaxEntries: 2})
	for index := 0; index < 32; index++ {
		parent := Key{Procedure: "M.Parent", Fingerprint: Hash(fmt.Sprintf("parent-%d", index)), Kernel: "summary"}
		dependency := Key{Procedure: "M.Dependency", Fingerprint: Hash(fmt.Sprintf("dependency-%d", index)), Kernel: "module"}
		store.RecordDependencies(parent, dependency)
	}
	if len(store.entries) > 2 || len(store.keys) > 4 || len(store.reverse) > 2 || len(store.deps) > 2 || len(store.order) > 4 {
		t.Fatalf("recorded dependency metadata exceeded bounds: entries=%d keys=%d reverse=%d deps=%d order=%d", len(store.entries), len(store.keys), len(store.reverse), len(store.deps), len(store.order))
	}
}

func TestRecordDependenciesPlaceholderDoesNotHit(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("r1")
	defer revision.Close()
	parent := Key{Procedure: "M.Parent", Fingerprint: Hash("parent"), Kernel: "summary"}
	dependency := Key{Procedure: "M.Dependency", Fingerprint: Hash("dependency"), Kernel: "module"}
	store.RecordDependencies(parent, dependency)
	value, hit, err := Evaluate(context.Background(), revision, parent, []Key{dependency}, func(context.Context) (int, error) {
		return 42, nil
	})
	if err != nil || hit || value != 42 {
		t.Fatalf("metadata placeholder evaluation = %d, hit %v, err %v", value, hit, err)
	}
}
