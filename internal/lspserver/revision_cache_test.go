package lspserver

import (
	"sync"
	"sync/atomic"
	"testing"
)

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
