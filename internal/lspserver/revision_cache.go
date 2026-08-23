package lspserver

import "sync"

// revisionCache stores the result of a project-level computation for the
// currently observed analysis revision. The build function runs while the
// cache is locked, so concurrent requests for the same revision share one
// result instead of duplicating the computation. A new revision (or a change
// in the completeness dimension) invalidates the previous value.
//
// Values returned by the cache must be treated as immutable by callers. This
// is intentional: project analysis results are snapshots and sharing them is
// what makes the at-most-once guarantee useful.
type revisionCache[T any] struct {
	mu       sync.Mutex
	revision uint64
	complete bool
	valid    bool
	value    T
}

func (c *revisionCache[T]) getOrBuild(revision uint64, complete bool, build func() T) T {
	c.mu.Lock()
	defer c.mu.Unlock()

	if c.valid && c.revision == revision && c.complete == complete {
		return c.value
	}
	value := build()
	// A request for an older snapshot may still be finishing after a newer
	// snapshot was published. Keep the newer value as the cache entry so a
	// stale request cannot cause the next current-revision request to rebuild.
	if c.valid && c.revision > revision {
		return value
	}
	c.revision = revision
	c.complete = complete
	c.value = value
	c.valid = true
	return value
}
