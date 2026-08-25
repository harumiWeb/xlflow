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

func TestComparableKeyIdentityDoesNotCollapseStringSeparatorCollision(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("separator-collision")
	defer revision.Close()
	first := Key{Procedure: "a\x00b", Fingerprint: "c", Kernel: "array"}
	second := Key{Procedure: "a", Fingerprint: "b\x00c", Kernel: "array"}
	if first.String() != second.String() {
		t.Fatalf("test keys should collide under the old string identity: %q != %q", first.String(), second.String())
	}
	if value, _, err := Evaluate(context.Background(), revision, first, nil, func(context.Context) (int, error) { return 1, nil }); err != nil || value != 1 {
		t.Fatalf("first evaluation = %d, err %v", value, err)
	}
	if value, hit, err := Evaluate(context.Background(), revision, second, nil, func(context.Context) (int, error) { return 2, nil }); err != nil || hit || value != 2 {
		t.Fatalf("second evaluation = %d, hit %v, err %v", value, hit, err)
	}
}

func TestEvaluateDependencySetIsOrderIndependent(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("dependency-order")
	defer revision.Close()
	parent := Key{Procedure: "M.P", Fingerprint: Hash("parent"), Kernel: "array"}
	first := Key{Procedure: "M.First", Fingerprint: Hash("first"), Kernel: "effect"}
	second := Key{Procedure: "M.Second", Fingerprint: Hash("second"), Kernel: "effect"}
	if _, hit, err := Evaluate(context.Background(), revision, parent, []Key{first, second}, func(context.Context) (int, error) { return 1, nil }); err != nil || hit {
		t.Fatalf("first evaluation hit=%v err=%v", hit, err)
	}
	if value, hit, err := Evaluate(context.Background(), revision, parent, []Key{second, first}, func(context.Context) (int, error) { return 2, nil }); err != nil || !hit || value != 1 {
		t.Fatalf("reordered evaluation = %d, hit %v, err %v", value, hit, err)
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
	store.mu.Lock()
	if len(store.pending) != 0 || len(store.procedureEntries) != 0 || len(store.documentEntries) != 0 || len(store.epochs) != 0 {
		store.mu.Unlock()
		t.Fatalf("panic left query metadata: pending=%d procedures=%d documents=%d epochs=%d", len(store.pending), len(store.procedureEntries), len(store.documentEntries), len(store.epochs))
	}
	store.mu.Unlock()
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

func TestInvalidateUsesOutputEqualityForGreenRecovery(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("green-recovery")
	defer revision.Close()
	callee := Key{Procedure: "M.Callee", Fingerprint: Hash("callee"), Kernel: "effect"}
	caller := Key{Procedure: "M.Caller", Fingerprint: Hash("caller"), Kernel: "array"}
	project := Key{Procedure: "project", Fingerprint: Hash("project"), Kernel: "summary"}
	var callerBuilds, projectBuilds atomic.Int32
	if _, _, err := Evaluate(context.Background(), revision, callee, nil, func(context.Context) (int, error) { return 1, nil }); err != nil {
		t.Fatal(err)
	}
	if _, _, err := Evaluate(context.Background(), revision, caller, []Key{callee}, func(context.Context) (int, error) {
		callerBuilds.Add(1)
		return 7, nil
	}); err != nil {
		t.Fatal(err)
	}
	if _, _, err := Evaluate(context.Background(), revision, project, []Key{caller}, func(context.Context) (int, error) {
		projectBuilds.Add(1)
		return 9, nil
	}); err != nil {
		t.Fatal(err)
	}

	store.Invalidate(callee)
	if value, hit, err := Evaluate(context.Background(), revision, caller, []Key{callee}, func(context.Context) (int, error) {
		callerBuilds.Add(1)
		return 7, nil
	}); err != nil || hit || value != 7 {
		t.Fatalf("unchanged caller = %d, hit %v, err %v", value, hit, err)
	}
	if value, hit, err := Evaluate(context.Background(), revision, project, []Key{caller}, func(context.Context) (int, error) {
		projectBuilds.Add(1)
		return 9, nil
	}); err != nil || !hit || value != 9 {
		t.Fatalf("green project = %d, hit %v, err %v", value, hit, err)
	}
	if callerBuilds.Load() != 2 || projectBuilds.Load() != 1 {
		t.Fatalf("build counts after green recovery = caller %d project %d", callerBuilds.Load(), projectBuilds.Load())
	}

	store.Invalidate(callee)
	if value, hit, err := Evaluate(context.Background(), revision, caller, []Key{callee}, func(context.Context) (int, error) {
		callerBuilds.Add(1)
		return 8, nil
	}); err != nil || hit || value != 8 {
		t.Fatalf("changed caller = %d, hit %v, err %v", value, hit, err)
	}
	if value, hit, err := Evaluate(context.Background(), revision, project, []Key{caller}, func(context.Context) (int, error) {
		projectBuilds.Add(1)
		return 10, nil
	}); err != nil || hit || value != 10 {
		t.Fatalf("changed project = %d, hit %v, err %v", value, hit, err)
	}
}

func TestInvalidateMultipleRootsKeepsUnrecoveredRootRed(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("multi-root-red")
	defer revision.Close()
	first := Key{Procedure: "M.First", Fingerprint: Hash("first"), Kernel: "effect"}
	second := Key{Procedure: "M.Second", Fingerprint: Hash("second"), Kernel: "effect"}
	parent := Key{Procedure: "M.Parent", Fingerprint: Hash("parent"), Kernel: "array"}
	if _, _, err := Evaluate(context.Background(), revision, first, nil, func(context.Context) (int, error) { return 1, nil }); err != nil {
		t.Fatal(err)
	}
	if _, _, err := Evaluate(context.Background(), revision, second, nil, func(context.Context) (int, error) { return 2, nil }); err != nil {
		t.Fatal(err)
	}
	if _, _, err := Evaluate(context.Background(), revision, parent, []Key{first, second}, func(context.Context) (int, error) { return 3, nil }); err != nil {
		t.Fatal(err)
	}

	store.Invalidate(first, second)
	if value, hit, err := Evaluate(context.Background(), revision, first, nil, func(context.Context) (int, error) { return 1, nil }); err != nil || hit || value != 1 {
		t.Fatalf("unchanged first = %d, hit %v, err %v", value, hit, err)
	}
	var parentBuilds atomic.Int32
	if value, hit, err := Evaluate(context.Background(), revision, parent, []Key{first, second}, func(context.Context) (int, error) {
		parentBuilds.Add(1)
		return 3, nil
	}); err != nil || hit || value != 3 {
		t.Fatalf("parent with unrecovered second = %d, hit %v, err %v", value, hit, err)
	}
	if parentBuilds.Load() != 1 {
		t.Fatal("parent incorrectly recovered while the second root remained stale")
	}
}

func TestEvaluateRetriesAfterConcurrentInvalidation(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("epoch-retry")
	defer revision.Close()
	key := Key{Procedure: "M.P", Fingerprint: Hash("body"), Kernel: "array"}
	started := make(chan struct{})
	release := make(chan struct{})
	result := make(chan struct {
		value int
		hit   bool
		err   error
	}, 1)
	var builds atomic.Int32
	go func() {
		value, hit, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) {
			if builds.Add(1) == 1 {
				close(started)
				<-release
				return 1, nil
			}
			return 2, nil
		})
		result <- struct {
			value int
			hit   bool
			err   error
		}{value: value, hit: hit, err: err}
	}()
	<-started
	store.Invalidate(key)
	close(release)
	got := <-result
	if got.err != nil || got.hit || got.value != 2 {
		t.Fatalf("concurrent invalidation result = %#v, want value 2 and miss", got)
	}
	if builds.Load() != 2 {
		t.Fatalf("build count = %d, want 2", builds.Load())
	}
}

func TestInvalidateMarksTransitiveClosureBeforeIntermediateRebuild(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("transitive-red")
	defer revision.Close()
	leaf := Key{Procedure: "M.Leaf", Fingerprint: Hash("leaf"), Kernel: "effect"}
	middle := Key{Procedure: "M.Middle", Fingerprint: Hash("middle"), Kernel: "array"}
	top := Key{Procedure: "project", Fingerprint: Hash("top"), Kernel: "summary"}
	if _, _, err := Evaluate(context.Background(), revision, leaf, nil, func(context.Context) (int, error) { return 1, nil }); err != nil {
		t.Fatal(err)
	}
	if _, _, err := Evaluate(context.Background(), revision, middle, []Key{leaf}, func(context.Context) (int, error) { return 2, nil }); err != nil {
		t.Fatal(err)
	}
	if _, _, err := Evaluate(context.Background(), revision, top, []Key{middle}, func(context.Context) (int, error) { return 3, nil }); err != nil {
		t.Fatal(err)
	}
	store.Invalidate(leaf)
	var builds atomic.Int32
	value, hit, err := Evaluate(context.Background(), revision, top, []Key{middle}, func(context.Context) (int, error) {
		builds.Add(1)
		return 3, nil
	})
	if err != nil || hit || value != 3 {
		t.Fatalf("transitive top evaluation = %d, hit %v, err %v", value, hit, err)
	}
	if builds.Load() != 1 {
		t.Fatalf("top was not marked stale before middle rebuild: builds=%d", builds.Load())
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

func TestInvalidateProceduresUsesDocumentIndex(t *testing.T) {
	store := New(Options{})
	revision := store.Begin("document-index")
	defer revision.Close()
	mainA := Key{Procedure: `C:\src\Main.bas::Main::A::0`, Fingerprint: Hash("a"), Kernel: "array"}
	mainB := Key{Procedure: `C:\src\Main.bas::Main::B::1`, Fingerprint: Hash("b"), Kernel: "array"}
	other := Key{Procedure: `C:\src\Other.bas::Main::A::0`, Fingerprint: Hash("other"), Kernel: "array"}
	for _, key := range []Key{mainA, mainB, other} {
		if _, _, err := Evaluate(context.Background(), revision, key, nil, func(context.Context) (int, error) { return 1, nil }); err != nil {
			t.Fatal(err)
		}
	}
	got := store.InvalidateProcedures(`C:\src\Main.bas`)
	want := []string{mainA.Procedure, mainB.Procedure}
	if len(got) != len(want) {
		t.Fatalf("invalidated = %#v, want %#v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("invalidated = %#v, want %#v", got, want)
		}
	}
	if _, hit, err := Evaluate(context.Background(), revision, other, nil, func(context.Context) (int, error) { return 2, nil }); err != nil || !hit {
		t.Fatalf("unrelated document entry hit=%v err=%v", hit, err)
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
	if len(store.entries) > 2 || len(store.procedureEntries) > 4 || len(store.documentEntries) > 4 || len(store.reverse) > 4 || len(store.order) > 4 {
		t.Fatalf("dependency metadata exceeded bounds: entries=%d procedures=%d documents=%d reverse=%d order=%d", len(store.entries), len(store.procedureEntries), len(store.documentEntries), len(store.reverse), len(store.order))
	}
}

func TestRecordDependenciesBoundsMetadata(t *testing.T) {
	store := New(Options{MaxEntries: 2})
	for index := 0; index < 32; index++ {
		parent := Key{Procedure: "M.Parent", Fingerprint: Hash(fmt.Sprintf("parent-%d", index)), Kernel: "summary"}
		dependency := Key{Procedure: "M.Dependency", Fingerprint: Hash(fmt.Sprintf("dependency-%d", index)), Kernel: "module"}
		store.RecordDependencies(parent, dependency)
	}
	if len(store.entries) > 2 || len(store.procedureEntries) > 4 || len(store.documentEntries) > 4 || len(store.reverse) > 2 || len(store.deps) > 2 || len(store.order) > 4 {
		t.Fatalf("recorded dependency metadata exceeded bounds: entries=%d procedures=%d documents=%d reverse=%d deps=%d order=%d", len(store.entries), len(store.procedureEntries), len(store.documentEntries), len(store.reverse), len(store.deps), len(store.order))
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

