package analysisstats

import (
	"context"
	"sync"
	"time"
)

type Stage struct {
	Name        string
	Elapsed     time.Duration
	Wait        time.Duration
	Outcome     string
	ResultCount int
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
	r.stages = append(r.stages, stage)
	r.mu.Unlock()
}

func (r *Recorder) Add(name string, value uint64) {
	if r == nil {
		return
	}
	r.mu.Lock()
	r.counters[name] += value
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
	started := time.Now()
	return func(count int, err error) {
		outcome := "ok"
		if err != nil {
			outcome = "error"
		}
		if ctx != nil && ctx.Err() != nil {
			outcome = "canceled"
		}
		if recorder := FromContext(ctx); recorder != nil {
			recorder.Record(Stage{Name: name, Elapsed: time.Since(started), Outcome: outcome, ResultCount: count})
		}
	}
}

// MeasureWait records time spent waiting for a shared artifact owned by
// another analysis request. It is separate from compute elapsed time so stage
// logs can distinguish contention from actual analysis work.
func MeasureWait(ctx context.Context, name string) func(error) {
	started := time.Now()
	return func(err error) {
		outcome := "ok"
		if err != nil {
			outcome = "error"
		}
		if ctx != nil && ctx.Err() != nil {
			outcome = "canceled"
		}
		if recorder := FromContext(ctx); recorder != nil {
			recorder.Record(Stage{Name: name, Wait: time.Since(started), Outcome: outcome})
		}
	}
}
