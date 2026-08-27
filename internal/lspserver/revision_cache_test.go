package lspserver

import (
	"context"
	"errors"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
)

func TestInitializeCapabilityTelemetryOmitsServerLifetimeTypeDB(t *testing.T) {
	recorder := analysisstats.NewRecorder()
	initializeCapabilityTelemetry(analysisstats.WithRecorder(context.Background(), recorder))
	_, counters := recorder.Totals()
	for _, counter := range counters {
		if counter.Name == analysisstats.CapabilityTypeDBBuildsCounter {
			t.Fatalf("LSP telemetry reported server-lifetime TypeDB counter: %+v", counter)
		}
	}
	for _, name := range analysisstats.CapabilityBuildCounters {
		if name == analysisstats.CapabilityTypeDBBuildsCounter {
			continue
		}
		found := false
		for _, counter := range counters {
			if counter.Name == name {
				found = true
				if counter.Value != 0 {
					t.Fatalf("LSP capability seed %q = %d, want zero", name, counter.Value)
				}
				break
			}
		}
		if !found {
			t.Fatalf("LSP capability seed %q missing: %+v", name, counters)
		}
	}
}

func TestRevisionCacheBuildsOnceForConcurrentRequests(t *testing.T) {
	var cache revisionCache[int]
	var builds atomic.Int32
	entered := make(chan struct{})
	release := make(chan struct{})
	first := make(chan int, 1)
	go func() {
		first <- cache.getOrBuild(11, true, func() int {
			builds.Add(1)
			close(entered)
			<-release
			return 42
		})
	}()
	<-entered

	const waiters = 32
	results := make(chan int, waiters)
	var wg sync.WaitGroup
	wg.Add(waiters)
	for i := 0; i < waiters; i++ {
		go func() {
			defer wg.Done()
			results <- cache.getOrBuild(11, true, func() int {
				builds.Add(1)
				return 99
			})
		}()
	}
	close(release)
	if got := <-first; got != 42 {
		t.Fatalf("first result = %d, want 42", got)
	}
	wg.Wait()
	close(results)
	for got := range results {
		if got != 42 {
			t.Fatalf("cached result = %d, want 42", got)
		}
	}
	if got := builds.Load(); got != 1 {
		t.Fatalf("build count = %d, want 1", got)
	}
}

func TestRevisionCacheInvalidatesOnRevisionOrCompletenessChange(t *testing.T) {
	var cache revisionCache[int]
	var builds int
	build := func(value int) func() int {
		return func() int {
			builds++
			return value
		}
	}

	if got := cache.getOrBuild(3, false, build(10)); got != 10 {
		t.Fatalf("initial value = %d, want 10", got)
	}
	if got := cache.getOrBuild(3, false, build(20)); got != 10 {
		t.Fatalf("same revision value = %d, want 10", got)
	}
	if got := cache.getOrBuild(3, true, build(30)); got != 30 {
		t.Fatalf("completeness transition value = %d, want 30", got)
	}
	if got := cache.getOrBuild(4, true, build(40)); got != 40 {
		t.Fatalf("revision transition value = %d, want 40", got)
	}
	if builds != 3 {
		t.Fatalf("build count = %d, want 3", builds)
	}
}

func TestRevisionCacheDoesNotReplaceNewerRevisionWithStaleResult(t *testing.T) {
	var cache revisionCache[int]
	var builds int
	build := func(value int) func() int {
		return func() int {
			builds++
			return value
		}
	}

	if got := cache.getOrBuild(8, true, build(80)); got != 80 {
		t.Fatalf("newer value = %d, want 80", got)
	}
	if got := cache.getOrBuild(7, true, build(70)); got != 70 {
		t.Fatalf("stale request value = %d, want 70", got)
	}
	if got := cache.getOrBuild(8, true, build(800)); got != 80 {
		t.Fatalf("newer cached value = %d, want 80", got)
	}
	if builds != 2 {
		t.Fatalf("build count = %d, want 2", builds)
	}
}

func TestRevisionCacheContextDoesNotPublishCanceledBuild(t *testing.T) {
	var cache revisionCache[int]
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, err := cache.getOrBuildContext(ctx, 12, true, func() (int, error) {
		t.Fatal("canceled cache build was invoked")
		return 42, nil
	}); !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled cache build error = %v, want context.Canceled", err)
	}

	var builds int
	got, err := cache.getOrBuildContext(context.Background(), 12, true, func() (int, error) {
		builds++
		return 7, nil
	})
	if err != nil || got != 7 || builds != 1 {
		t.Fatalf("retry after canceled build = (value=%d, err=%v, builds=%d)", got, err, builds)
	}
}

func TestRevisionCacheContextDoesNotPublishBuildCanceledAfterWork(t *testing.T) {
	var cache revisionCache[int]
	ctx, cancel := context.WithCancel(context.Background())
	var builds int
	if _, err := cache.getOrBuildContext(ctx, 13, true, func() (int, error) {
		builds++
		cancel()
		return 9, nil
	}); !errors.Is(err, context.Canceled) {
		t.Fatalf("post-build cancellation error = %v, want context.Canceled", err)
	}
	if _, err := cache.getOrBuildContext(context.Background(), 13, true, func() (int, error) {
		builds++
		return 11, nil
	}); err != nil || builds != 2 {
		t.Fatalf("post-build cancellation retry = (err=%v, builds=%d)", err, builds)
	}
}

func TestRevisionCacheContextWaiterCanCancelWhileBuildIsInFlight(t *testing.T) {
	var cache revisionCache[int]
	entered := make(chan struct{})
	release := make(chan struct{})
	firstDone := make(chan error, 1)
	go func() {
		_, err := cache.getOrBuildContext(context.Background(), 14, true, func() (int, error) {
			close(entered)
			<-release
			return 14, nil
		})
		firstDone <- err
	}()
	<-entered

	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	duplicateBuild := make(chan struct{}, 1)
	secondDone := make(chan error, 1)
	go func() {
		_, err := cache.getOrBuildContext(ctx, 14, true, func() (int, error) {
			duplicateBuild <- struct{}{}
			return 0, nil
		})
		secondDone <- err
	}()
	select {
	case err := <-secondDone:
		if !errors.Is(err, context.Canceled) {
			t.Fatalf("canceled waiter error = %v, want context.Canceled", err)
		}
	case <-time.After(time.Second):
		t.Fatal("canceled waiter remained blocked behind an in-flight build")
	}
	select {
	case <-duplicateBuild:
		t.Fatal("canceled waiter started a duplicate build")
	default:
	}
	close(release)
	if err := <-firstDone; err != nil {
		t.Fatalf("first build error = %v", err)
	}
}
