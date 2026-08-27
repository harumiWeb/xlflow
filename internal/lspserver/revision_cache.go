package lspserver

import (
	"context"
	"sync"
)

// revisionCache stores the result of a project-level computation for the
// currently observed analysis revision. One in-flight build is shared by
// concurrent requests, so they do not duplicate the computation. A new
// revision (or a change in the completeness dimension) invalidates the
// previous value.
//
// Values returned by the cache must be treated as immutable by callers. This
// is intentional: project analysis results are snapshots and sharing them is
// what makes the at-most-once guarantee useful.
type revisionCache[T any] struct {
	mu        sync.Mutex
	revision  uint64
	complete  bool
	valid     bool
	value     T
	building  bool
	buildDone chan struct{}
}

func (c *revisionCache[T]) getOrBuild(revision uint64, complete bool, build func() T) T {
	value, _ := c.getOrBuildContext(context.Background(), revision, complete, func() (T, error) {
		return build(), nil
	})
	return value
}

// getOrBuildContext is the cancellable variant used by project preparation.
// The cache tracks an in-flight build separately from the stored value. A
// canceled or failed build is never published, so a later request can retry
// the same revision while waiters can still observe cancellation.
func (c *revisionCache[T]) getOrBuildContext(ctx context.Context, revision uint64, complete bool, build func() (T, error)) (T, error) {
	var zero T
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return zero, err
	}

	for {
		c.mu.Lock()
		if c.valid && c.revision == revision && c.complete == complete {
			value := c.value
			c.mu.Unlock()
			return value, nil
		}
		if !c.building {
			done := make(chan struct{})
			c.building = true
			c.buildDone = done
			c.mu.Unlock()

			value, err := build()
			c.mu.Lock()
			c.building = false
			c.buildDone = nil
			close(done)
			if err != nil {
				c.mu.Unlock()
				return zero, err
			}
			if err := ctx.Err(); err != nil {
				c.mu.Unlock()
				return zero, err
			}
			// Do not let an older request replace a newer cached revision.
			if !c.valid || c.revision <= revision {
				c.revision = revision
				c.complete = complete
				c.value = value
				c.valid = true
			}
			c.mu.Unlock()
			return value, nil
		}
		done := c.buildDone
		c.mu.Unlock()
		select {
		case <-done:
			// A completed build and cancellation can become ready at the
			// same time. Recheck after the completion signal so select's
			// nondeterministic choice cannot let a canceled waiter proceed.
			if err := ctx.Err(); err != nil {
				return zero, err
			}
		case <-ctx.Done():
			return zero, ctx.Err()
		}
	}
}
