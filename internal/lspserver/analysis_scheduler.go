package lspserver

import (
	"context"
	"sync"
	"time"
)

type analysisWorkClass uint8

const (
	analysisWorkInteractive analysisWorkClass = iota
	analysisWorkFast
	analysisWorkBackground
	analysisWorkClassCount
)

func (c analysisWorkClass) String() string {
	switch c {
	case analysisWorkInteractive:
		return "interactive"
	case analysisWorkFast:
		return "fast"
	case analysisWorkBackground:
		return "background"
	default:
		return "unknown"
	}
}

type analysisWorkerState struct {
	Current int
	Maximum int
}

// analysisScheduler reserves practical execution capacity for latency-sensitive
// requests while keeping every class bounded. Work already running is not
// preempted; priority is applied whenever a permit is released.
type analysisScheduler struct {
	mu              sync.Mutex
	changed         chan struct{}
	totalLimit      int
	backgroundLimit int
	active          [analysisWorkClassCount]int
	maximum         [analysisWorkClassCount]int
	waiters         [analysisWorkClassCount]int
}

func newAnalysisScheduler(totalLimit int) *analysisScheduler {
	if totalLimit < 1 {
		totalLimit = 1
	}
	backgroundLimit := totalLimit
	if totalLimit > 1 {
		backgroundLimit--
	}
	return &analysisScheduler{
		changed:         make(chan struct{}),
		totalLimit:      totalLimit,
		backgroundLimit: backgroundLimit,
	}
}

func (s *analysisScheduler) acquire(ctx context.Context, class analysisWorkClass) (func(), time.Duration, analysisWorkerState, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, 0, analysisWorkerState{}, err
	}
	if s == nil {
		return func() {}, 0, analysisWorkerState{}, nil
	}

	started := time.Now()
	waiting := false
	for {
		s.mu.Lock()
		if err := ctx.Err(); err != nil {
			if waiting {
				s.waiters[class]--
			}
			s.mu.Unlock()
			return nil, time.Since(started), analysisWorkerState{}, err
		}
		if s.canAcquireLocked(class) {
			if waiting {
				s.waiters[class]--
			}
			s.active[class]++
			if s.active[class] > s.maximum[class] {
				s.maximum[class] = s.active[class]
			}
			state := analysisWorkerState{Current: s.active[class], Maximum: s.maximum[class]}
			s.mu.Unlock()
			wait := time.Duration(0)
			if waiting {
				wait = time.Since(started)
			}
			var once sync.Once
			return func() {
				once.Do(func() { s.release(class) })
			}, wait, state, nil
		}
		if !waiting {
			s.waiters[class]++
			waiting = true
		}
		changed := s.changed
		s.mu.Unlock()

		select {
		case <-changed:
		case <-ctx.Done():
			s.mu.Lock()
			if waiting {
				s.waiters[class]--
			}
			s.mu.Unlock()
			return nil, time.Since(started), analysisWorkerState{}, ctx.Err()
		}
	}
}

func (s *analysisScheduler) canAcquireLocked(class analysisWorkClass) bool {
	total := 0
	for _, active := range s.active {
		total += active
	}
	// A one-permit host cannot reserve a physical slot without starving
	// background analysis. Permit one interactive request to overcommit the
	// single background worker; the runtime scheduler still provides fairness.
	if class == analysisWorkInteractive && s.totalLimit == 1 && total == 1 && s.active[analysisWorkBackground] == 1 {
		return true
	}
	if total >= s.totalLimit {
		return false
	}
	switch class {
	case analysisWorkInteractive:
		return true
	case analysisWorkFast:
		return s.waiters[analysisWorkInteractive] == 0
	case analysisWorkBackground:
		return s.active[analysisWorkBackground] < s.backgroundLimit &&
			s.waiters[analysisWorkInteractive] == 0 && s.waiters[analysisWorkFast] == 0
	default:
		return false
	}
}

func (s *analysisScheduler) release(class analysisWorkClass) {
	s.mu.Lock()
	if s.active[class] > 0 {
		s.active[class]--
	}
	close(s.changed)
	s.changed = make(chan struct{})
	s.mu.Unlock()
}

func (s *analysisScheduler) limits() (total, background int) {
	if s == nil {
		return 1, 1
	}
	return s.totalLimit, s.backgroundLimit
}

func (s *analysisScheduler) state(class analysisWorkClass) analysisWorkerState {
	if s == nil {
		return analysisWorkerState{}
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	return analysisWorkerState{Current: s.active[class], Maximum: s.maximum[class]}
}

func (s *analysisScheduler) waiterCount(class analysisWorkClass) int {
	if s == nil {
		return 0
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.waiters[class]
}
