package analysisstats

import (
	"context"
	"sort"
	"sync"
	"time"
)

type Stage struct {
	Name        string
	Elapsed     time.Duration
	Wait        time.Duration
	Outcome     string
	ResultCount int
	Calls       int
}

type Recorder struct {
	mu       sync.Mutex
	stages   []Stage
	counters map[string]uint64
}

func NewRecorder() *Recorder { return &Recorder{counters: make(map[string]uint64)} }

func (r *Recorder) Record(stage Stage) {
	if r == nil {
		return
	}
	r.mu.Lock()
	if stage.Calls == 0 {
		stage.Calls = 1
	}
	r.stages = append(r.stages, stage)
	r.mu.Unlock()
}

// Totals returns stages aggregated by name and counters sorted by name. Stage
// order follows the first observation so callers can render the analysis
// pipeline in execution order while repeated per-file measurements collapse
// into one stable record.
func (r *Recorder) Totals() ([]Stage, []Counter) {
	stages, counters := r.Snapshot()
	if len(stages) == 0 && len(counters) == 0 {
		return nil, nil
	}
	indexes := make(map[string]int, len(stages))
	totals := make([]Stage, 0, len(stages))
	for _, stage := range stages {
		index, ok := indexes[stage.Name]
		if !ok {
			indexes[stage.Name] = len(totals)
			if stage.Calls == 0 {
				stage.Calls = 1
			}
			totals = append(totals, stage)
			continue
		}
		total := &totals[index]
		total.Elapsed += stage.Elapsed
		total.Wait += stage.Wait
		total.ResultCount += stage.ResultCount
		if stage.Calls == 0 {
			stage.Calls = 1
		}
		total.Calls += stage.Calls
		total.Outcome = combinedOutcome(total.Outcome, stage.Outcome)
	}
	names := make([]string, 0, len(counters))
	for name := range counters {
		names = append(names, name)
	}
	sort.Strings(names)
	orderedCounters := make([]Counter, 0, len(names))
	for _, name := range names {
		orderedCounters = append(orderedCounters, Counter{Name: name, Value: counters[name]})
	}
	return totals, orderedCounters
}

type Counter struct {
	Name  string
	Value uint64
}

func combinedOutcome(current, next string) string {
	rank := func(outcome string) int {
		switch outcome {
		case "canceled":
			return 3
		case "error":
			return 2
		case "ok":
			return 1
		default:
			return 0
		}
	}
	if rank(next) > rank(current) {
		return next
	}
	return current
}

func (r *Recorder) Add(name string, value uint64) {
	r.AddSum(name, value)
}

// AddSum adds value to an additive counter. Add is retained as a compatibility
// alias for existing instrumentation callers.
func (r *Recorder) AddSum(name string, value uint64) {
	if r == nil {
		return
	}
	r.mu.Lock()
	r.counters[name] += value
	r.mu.Unlock()
}

// AddMax records the greatest value observed for a maximum counter. Maximum
// workload dimensions are intentionally separate from additive counters so
// repeated observations do not turn a per-file or per-procedure maximum into
// a sum.
func (r *Recorder) AddMax(name string, value uint64) {
	if r == nil {
		return
	}
	r.mu.Lock()
	current, ok := r.counters[name]
	if !ok || value > current {
		r.counters[name] = value
	}
	r.mu.Unlock()
}

func (r *Recorder) Snapshot() ([]Stage, map[string]uint64) {
	if r == nil {
		return nil, nil
	}
	r.mu.Lock()
	defer r.mu.Unlock()
	stages := append([]Stage(nil), r.stages...)
	counters := make(map[string]uint64, len(r.counters))
	for key, value := range r.counters {
		counters[key] = value
	}
	return stages, counters
}

type recorderKey struct{}

func WithRecorder(ctx context.Context, recorder *Recorder) context.Context {
	if recorder == nil {
		return ctx
	}
	return context.WithValue(ctx, recorderKey{}, recorder)
}

func FromContext(ctx context.Context) *Recorder {
	if ctx == nil {
		return nil
	}
	recorder, _ := ctx.Value(recorderKey{}).(*Recorder)
	return recorder
}

func Measure(ctx context.Context, name string) func(int, error) {
	recorder := FromContext(ctx)
	if recorder == nil {
		return func(int, error) {}
	}
	started := time.Now()
	return func(count int, err error) {
		outcome := "ok"
		if err != nil {
			outcome = "error"
		}
		if ctx != nil && ctx.Err() != nil {
			outcome = "canceled"
		}
		recorder.Record(Stage{Name: name, Elapsed: time.Since(started), Outcome: outcome, ResultCount: count})
	}
}

// MeasureWait records time spent waiting for a shared artifact owned by
// another analysis request. It is separate from compute elapsed time so stage
// logs can distinguish contention from actual analysis work.
func MeasureWait(ctx context.Context, name string) func(error) {
	recorder := FromContext(ctx)
	if recorder == nil {
		return func(error) {}
	}
	started := time.Now()
	return func(err error) {
		outcome := "ok"
		if err != nil {
			outcome = "error"
		}
		if ctx != nil && ctx.Err() != nil {
			outcome = "canceled"
		}
		recorder.Record(Stage{Name: name, Wait: time.Since(started), Outcome: outcome})
	}
}
