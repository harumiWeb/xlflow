package lspserver

import (
	"context"
	"errors"
	"testing"
	"time"
)

func TestAnalysisSchedulerPrioritizesInteractiveThenFastThenBackground(t *testing.T) {
	scheduler := newAnalysisScheduler(2)
	backgroundHolder, _, _, err := scheduler.acquire(context.Background(), analysisWorkBackground)
	if err != nil {
		t.Fatal(err)
	}
	fastHolder, _, _, err := scheduler.acquire(context.Background(), analysisWorkFast)
	if err != nil {
		t.Fatal(err)
	}

	acquiredWork := make(chan acquiredSchedulerWork, 3)
	for _, class := range []analysisWorkClass{analysisWorkBackground, analysisWorkFast, analysisWorkInteractive} {
		go func() {
			release, _, _, acquireErr := scheduler.acquire(context.Background(), class)
			if acquireErr == nil {
				acquiredWork <- acquiredSchedulerWork{class: class, release: release}
			}
		}()
		waitForAnalysisWaiter(t, scheduler, class)
	}

	fastHolder()
	first := waitForScheduledClass(t, acquiredWork, analysisWorkInteractive)
	if state := scheduler.state(analysisWorkBackground); state.Current > scheduler.backgroundLimit {
		t.Fatalf("background workers = %d, limit %d", state.Current, scheduler.backgroundLimit)
	}
	first.release()
	second := waitForScheduledClass(t, acquiredWork, analysisWorkFast)
	backgroundHolder()
	second.release()
	third := waitForScheduledClass(t, acquiredWork, analysisWorkBackground)
	third.release()

	for _, class := range []analysisWorkClass{analysisWorkInteractive, analysisWorkFast, analysisWorkBackground} {
		state := scheduler.state(class)
		if state.Current != 0 || state.Maximum < 1 {
			t.Fatalf("%s state = %+v, want current 0 and maximum >= 1", class, state)
		}
	}
}

func TestAnalysisSchedulerCancellationRemovesWaitingWork(t *testing.T) {
	scheduler := newAnalysisScheduler(1)
	holder, _, _, err := scheduler.acquire(context.Background(), analysisWorkBackground)
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	done := make(chan error, 1)
	go func() {
		_, _, _, acquireErr := scheduler.acquire(ctx, analysisWorkBackground)
		done <- acquireErr
	}()
	waitForAnalysisWaiter(t, scheduler, analysisWorkBackground)
	cancel()
	select {
	case err := <-done:
		if !errors.Is(err, context.Canceled) {
			t.Fatalf("canceled waiter = %v, want context.Canceled", err)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("canceled waiter did not return")
	}
	if got := scheduler.waiterCount(analysisWorkBackground); got != 0 {
		t.Fatalf("background waiters = %d, want 0", got)
	}
	holder()
	release, wait, _, err := scheduler.acquire(context.Background(), analysisWorkBackground)
	if err != nil || wait != 0 {
		t.Fatalf("acquire after cancellation = (wait=%v, err=%v)", wait, err)
	}
	release()
}

func TestAnalysisSchedulerCancellationWakesOtherWaiters(t *testing.T) {
	scheduler := newAnalysisScheduler(1)
	scheduler.mu.Lock()
	changed := scheduler.changed
	scheduler.waiters[analysisWorkInteractive] = 1
	scheduler.waiters[analysisWorkFast] = 1
	scheduler.removeWaiterLocked(analysisWorkInteractive)
	scheduler.mu.Unlock()

	select {
	case <-changed:
	default:
		t.Fatal("removing a canceled waiter did not notify other waiters")
	}
	if got := scheduler.waiterCount(analysisWorkInteractive); got != 0 {
		t.Fatalf("interactive waiters = %d, want 0", got)
	}
	if got := scheduler.waiterCount(analysisWorkFast); got != 1 {
		t.Fatalf("fast waiters = %d, want 1", got)
	}
}

func waitForScheduledClass(t *testing.T, acquired <-chan acquiredSchedulerWork, want analysisWorkClass) acquiredSchedulerWork {
	t.Helper()
	select {
	case got := <-acquired:
		if got.class != want {
			got.release()
			t.Fatalf("scheduled class = %s, want %s", got.class, want)
		}
		return got
	case <-time.After(time.Second):
		t.Fatalf("timed out waiting for %s work", want)
		return acquiredSchedulerWork{}
	}
}

type acquiredSchedulerWork struct {
	class   analysisWorkClass
	release func()
}
