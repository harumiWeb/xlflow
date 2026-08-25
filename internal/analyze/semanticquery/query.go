// Package semanticquery provides the process-local ownership boundary for
// revision-scoped semantic query results.  It deliberately stores opaque
// values: the analyzer owns the value's immutability and the query package
// owns only publication, single-flight, and dependency lifetime.
package semanticquery

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"fmt"
	"strings"
	"sync"
)

var ErrRevisionClosed = errors.New("semantic query revision is closed")

// Metrics is implemented by the analyzer performance recorder.  Keeping this
// interface small prevents the query package from depending on analyzer or
// protocol types while preserving the existing stderr-only telemetry path.
type Metrics interface {
	RecordSemanticQueryHit()
	RecordSemanticQueryMiss()
	RecordSemanticQueryInvalidatedProcedures(uint64)
	RecordSemanticQueryRecomputedKernel()
}

// Key identifies one semantic query.  Procedure and Fingerprint are kept
// separate so callers can report invalidated procedures without parsing an
// opaque cache key.  Config and Capability are content revisions, not a
// monotonically increasing workspace number; unchanged values can therefore
// be shared by adjacent revisions.
type Key struct {
	Procedure   string
	Fingerprint string
	Kernel      string
	Config      string
	Capability  string
}

func (k Key) String() string {
	return strings.Join([]string{k.Procedure, k.Fingerprint, k.Kernel, k.Config, k.Capability}, "\x00")
}

// Hash returns a deterministic lowercase SHA-256 digest of length-prefixed
// parts.  Length-prefixing keeps the digest stable even when a part contains
// the separator used by Key.String.
func Hash(parts ...string) string {
	h := sha256.New()
	for _, part := range parts {
		_, _ = fmt.Fprintf(h, "%d:", len(part))
		h.Write([]byte(part))
	}
	return hex.EncodeToString(h.Sum(nil))
}

type Options struct {
	MaxEntries   int
	MaxRevisions int
	Metrics      Metrics
}

var processStore = New(Options{})

// DefaultStore is the process-lifetime cache used by analyzer entrypoints
// that are not owned by an LSP server. Callers that own a workspace lifecycle
// may provide a bounded Store through Context instead.
func DefaultStore() *Store { return processStore }

type contextKey struct{}

// Context carries an optional process-local store through batch and LSP
// adapters without adding a protocol parameter to every analyzer entrypoint.
// The revision string is supplied by a coherent workspace snapshot when one
// exists; an empty string asks the batch entrypoint to derive one after parse.
type Context struct {
	Store    *Store
	Revision string
	Metrics  Metrics
}

func WithContext(ctx context.Context, value Context) context.Context {
	if ctx == nil {
		ctx = context.Background()
	}
	return context.WithValue(ctx, contextKey{}, value)
}

func FromContext(ctx context.Context) Context {
	if ctx == nil {
		return Context{}
	}
	value, _ := ctx.Value(contextKey{}).(Context)
	return value
}

type cacheValue struct {
	key   Key
	value any
}

// dependencyRecord is an eviction-tracked placeholder for a query whose
// value was published outside Evaluate. It keeps RecordDependencies metadata
// bounded without making a metadata-only record look like a cache hit.
type dependencyRecord struct{}

type pendingValue struct {
	done chan struct{}
}

type revisionState struct {
	id   string
	refs int
}

// Store owns values shared by revisions in one process.  Values are retained
// only as completed opaque objects.  The FIFO bound is intentionally an
// implementation detail; no disk or cross-process cache format is implied.
type Store struct {
	mu sync.Mutex

	entries map[string]cacheValue
	order   []string
	pending map[string]*pendingValue
	reverse map[string]map[string]Key
	keys    map[string]Key
	epochs  map[string]uint64
	deps    map[string][]string

	revisions    map[string]*revisionState
	sequence     uint64
	maxEntries   int
	maxRevisions int
	latest       string
	metrics      Metrics
}

func New(opts Options) *Store {
	maxEntries := opts.MaxEntries
	if maxEntries <= 0 {
		maxEntries = 8192
	}
	maxRevisions := opts.MaxRevisions
	if maxRevisions <= 0 {
		maxRevisions = 4
	}
	return &Store{
		entries:      make(map[string]cacheValue),
		pending:      make(map[string]*pendingValue),
		reverse:      make(map[string]map[string]Key),
		keys:         make(map[string]Key),
		epochs:       make(map[string]uint64),
		deps:         make(map[string][]string),
		revisions:    make(map[string]*revisionState),
		maxEntries:   maxEntries,
		maxRevisions: maxRevisions,
		metrics:      opts.Metrics,
	}
}

// Revision is a lightweight handle.  Several concurrent requests may hold
// the same revision; closing one handle does not retire values used by the
// others.
type Revision struct {
	store  *Store
	id     string
	mu     sync.Mutex
	closed bool
}

func (s *Store) Begin(id string) *Revision {
	if s == nil {
		return &Revision{}
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	if id == "" {
		s.sequence++
		id = fmt.Sprintf("anonymous-%d", s.sequence)
	}
	state := s.revisions[id]
	if state == nil {
		state = &revisionState{id: id}
		s.revisions[id] = state
	}
	state.refs++
	s.latest = id
	s.pruneRevisionsLocked()
	return &Revision{store: s, id: id}
}

func (r *Revision) ID() string {
	if r == nil {
		return ""
	}
	return r.id
}

func (r *Revision) Close() {
	if r == nil || r.store == nil {
		return
	}
	r.mu.Lock()
	if r.closed {
		r.mu.Unlock()
		return
	}
	r.closed = true
	r.mu.Unlock()
	r.store.mu.Lock()
	if state := r.store.revisions[r.id]; state != nil && state.refs > 0 {
		state.refs--
	}
	r.store.pruneRevisionsLocked()
	r.store.mu.Unlock()
}

func (r *Revision) isClosed() bool {
	if r == nil || r.store == nil {
		return true
	}
	r.mu.Lock()
	closed := r.closed
	r.mu.Unlock()
	return closed
}

// Evaluate returns a completed value, a hit flag, and an error.  A build is
// published only after it returns successfully and its context is still live.
// Waiters retry after a canceled or failed producer instead of receiving a
// partial value.
func (r *Revision) Evaluate(ctx context.Context, key Key, dependencies []Key, build func(context.Context) (any, error)) (any, bool, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if r == nil || r.store == nil || r.isClosed() {
		return nil, false, ErrRevisionClosed
	}
	s := r.store
	metrics := s.metricsFor(ctx)
	stable := key.String()
	for {
		if err := ctx.Err(); err != nil {
			return nil, false, err
		}
		s.mu.Lock()
		if cached, ok := s.entries[stable]; ok {
			if _, metadataOnly := cached.value.(dependencyRecord); metadataOnly {
				s.removeNodeLocked(stable)
				s.mu.Unlock()
				continue
			}
			if !s.dependenciesMatchLocked(stable, dependencies) {
				s.invalidateIDsLocked([]string{stable}, metrics)
				s.mu.Unlock()
				continue
			}
			s.mu.Unlock()
			if metrics != nil {
				metrics.RecordSemanticQueryHit()
			}
			return cached.value, true, nil
		}
		if pending := s.pending[stable]; pending != nil {
			done := pending.done
			s.mu.Unlock()
			select {
			case <-done:
				continue
			case <-ctx.Done():
				return nil, false, ctx.Err()
			}
		}
		pending := &pendingValue{done: make(chan struct{})}
		epoch := s.epochs[stable]
		s.pending[stable] = pending
		if metrics != nil {
			metrics.RecordSemanticQueryMiss()
		}
		s.mu.Unlock()

		var value any
		var err error
		func() {
			defer func() {
				if recovered := recover(); recovered != nil {
					s.mu.Lock()
					delete(s.pending, stable)
					close(pending.done)
					s.mu.Unlock()
					panic(recovered)
				}
			}()
			value, err = build(ctx)
		}()
		if err == nil {
			err = ctx.Err()
		}
		s.mu.Lock()
		delete(s.pending, stable)
		if err == nil && s.epochs[stable] == epoch {
			s.entries[stable] = cacheValue{key: key, value: value}
			s.keys[stable] = key
			s.order = append(s.order, stable)
			s.recordDependenciesLocked(stable, key, dependencies)
			s.pruneEntriesLocked()
			if metrics != nil {
				metrics.RecordSemanticQueryRecomputedKernel()
			}
		}
		close(pending.done)
		s.mu.Unlock()
		if err != nil {
			return nil, false, err
		}
		return value, false, nil
	}
}

func (s *Store) dependenciesMatchLocked(parent string, dependencies []Key) bool {
	stored := s.deps[parent]
	if len(stored) != len(dependencies) {
		return false
	}
	seen := make(map[string]struct{}, len(dependencies))
	for _, dependency := range dependencies {
		seen[dependency.String()] = struct{}{}
	}
	for _, dependencyID := range stored {
		if _, ok := seen[dependencyID]; !ok {
			return false
		}
	}
	return true
}

func (s *Store) metricsFor(ctx context.Context) Metrics {
	if value := FromContext(ctx).Metrics; value != nil {
		return value
	}
	if s == nil {
		return nil
	}
	return s.metrics
}

func Evaluate[T any](ctx context.Context, r *Revision, key Key, dependencies []Key, build func(context.Context) (T, error)) (T, bool, error) {
	var zero T
	value, hit, err := r.Evaluate(ctx, key, dependencies, func(ctx context.Context) (any, error) {
		return build(ctx)
	})
	if err != nil {
		return zero, hit, err
	}
	result, ok := value.(T)
	if !ok {
		return zero, hit, fmt.Errorf("semantic query %s returned %T, want %T", key.Kernel, value, zero)
	}
	return result, hit, nil
}

func (s *Store) recordDependenciesLocked(parent string, parentKey Key, dependencies []Key) {
	for _, dependencyID := range s.deps[parent] {
		if parents := s.reverse[dependencyID]; parents != nil {
			delete(parents, parent)
			if len(parents) == 0 {
				delete(s.reverse, dependencyID)
			}
		}
	}
	s.keys[parent] = parentKey
	ids := make([]string, 0, len(dependencies))
	seen := make(map[string]struct{}, len(dependencies))
	for _, dependency := range dependencies {
		dependencyID := dependency.String()
		if _, ok := seen[dependencyID]; ok {
			continue
		}
		seen[dependencyID] = struct{}{}
		ids = append(ids, dependencyID)
		s.keys[dependencyID] = dependency
		parents := s.reverse[dependencyID]
		if parents == nil {
			parents = make(map[string]Key)
			s.reverse[dependencyID] = parents
		}
		parents[parent] = parentKey
	}
	s.deps[parent] = ids
}

// RecordDependencies records a dependency for an already-computed query. It
// is useful for capability builders that publish an immutable value outside
// Evaluate but still need reverse invalidation edges. Metadata-only records
// are represented by an eviction-tracked placeholder so repeated external
// registrations cannot bypass MaxEntries.
func (s *Store) RecordDependencies(parent Key, dependencies ...Key) {
	if s == nil {
		return
	}
	s.mu.Lock()
	stable := parent.String()
	if _, ok := s.entries[stable]; !ok {
		s.entries[stable] = cacheValue{key: parent, value: dependencyRecord{}}
		s.order = append(s.order, stable)
	}
	s.recordDependenciesLocked(stable, parent, dependencies)
	s.pruneEntriesLocked()
	s.mu.Unlock()
}

// Invalidate removes changed queries and their transitive dependents. The
// returned procedure identities are sorted and unique for deterministic
// telemetry. Reverse traversal is completed before affected metadata is
// released, so callers can invalidate an old/new dependency union when a call
// is removed or redirected without retaining unbounded historical edges.
func (s *Store) Invalidate(changed ...Key) []string {
	return s.InvalidateContext(context.Background(), changed...)
}

// InvalidateContext is the request-aware invalidation form. The context may
// carry a recorder through Context.Metrics; callers without request telemetry
// can use Invalidate.
func (s *Store) InvalidateContext(ctx context.Context, changed ...Key) []string {
	if s == nil || len(changed) == 0 {
		return nil
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	queue := make([]string, 0, len(changed))
	for _, key := range changed {
		queue = append(queue, key.String())
	}
	return s.invalidateIDsLocked(queue, s.metricsFor(ctx))
}

// InvalidateProcedures removes all cached queries belonging to the supplied
// procedure identities (or identity prefixes) and their reverse dependents.
// LSP callers use path prefixes because a source edit can replace a procedure
// declaration and therefore cannot always reconstruct the old full key.
func (s *Store) InvalidateProcedures(procedures ...string) []string {
	return s.InvalidateProceduresContext(context.Background(), procedures...)
}

// InvalidateProceduresContext is the request-aware procedure invalidation
// form used when a workspace change is handled alongside a recorder.
func (s *Store) InvalidateProceduresContext(ctx context.Context, procedures ...string) []string {
	if s == nil || len(procedures) == 0 {
		return nil
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	queue := make([]string, 0)
	for id, key := range s.keys {
		for _, procedure := range procedures {
			if procedure != "" && (key.Procedure == procedure || strings.HasPrefix(key.Procedure, procedure)) {
				queue = append(queue, id)
				break
			}
		}
	}
	return s.invalidateIDsLocked(queue, s.metricsFor(ctx))
}

func (s *Store) invalidateIDsLocked(queue []string, metrics Metrics) []string {
	seen := make(map[string]bool)
	procedures := make(map[string]bool)
	affected := make([]string, 0, len(queue))
	for len(queue) > 0 {
		id := queue[0]
		queue = queue[1:]
		if seen[id] {
			continue
		}
		seen[id] = true
		affected = append(affected, id)
		s.epochs[id]++
		if key, ok := s.keys[id]; ok {
			if key.Procedure != "" {
				procedures[key.Procedure] = true
			}
		}
		for parent := range s.reverse[id] {
			queue = append(queue, parent)
		}
	}
	for _, id := range affected {
		s.removeNodeLocked(id)
	}
	if len(s.order) > s.maxEntries*2 {
		compact := make([]string, 0, len(s.entries))
		seenOrder := make(map[string]struct{}, len(s.entries))
		for _, id := range s.order {
			if _, ok := s.entries[id]; !ok {
				continue
			}
			if _, ok := seenOrder[id]; ok {
				continue
			}
			seenOrder[id] = struct{}{}
			compact = append(compact, id)
		}
		s.order = compact
	}
	result := make([]string, 0, len(procedures))
	for procedure := range procedures {
		result = append(result, procedure)
	}
	for i := 1; i < len(result); i++ {
		for j := i; j > 0 && result[j] < result[j-1]; j-- {
			result[j], result[j-1] = result[j-1], result[j]
		}
	}
	if metrics != nil {
		metrics.RecordSemanticQueryInvalidatedProcedures(uint64(len(result)))
	}
	return result
}

func (s *Store) removeNodeLocked(id string) {
	for _, dependencyID := range s.deps[id] {
		if parents := s.reverse[dependencyID]; parents != nil {
			delete(parents, id)
			if len(parents) == 0 {
				delete(s.reverse, dependencyID)
				s.maybeRemoveDependencyMetadataLocked(dependencyID)
			}
		}
	}
	delete(s.deps, id)
	delete(s.entries, id)
	if len(s.reverse[id]) == 0 && s.pending[id] == nil && len(s.deps[id]) == 0 {
		delete(s.reverse, id)
		delete(s.keys, id)
		delete(s.epochs, id)
	}
}

func (s *Store) maybeRemoveDependencyMetadataLocked(id string) {
	if _, ok := s.entries[id]; ok {
		return
	}
	if s.pending[id] != nil || len(s.reverse[id]) != 0 || len(s.deps[id]) != 0 {
		return
	}
	delete(s.keys, id)
	delete(s.epochs, id)
}

func (s *Store) pruneEntriesLocked() {
	for len(s.entries) > s.maxEntries && len(s.order) > 0 {
		oldest := s.order[0]
		s.order = s.order[1:]
		if _, ok := s.entries[oldest]; !ok {
			continue
		}
		s.removeNodeLocked(oldest)
	}
	if len(s.order) > s.maxEntries*2 {
		compact := make([]string, 0, len(s.entries))
		seen := make(map[string]struct{}, len(s.entries))
		for _, id := range s.order {
			if _, ok := s.entries[id]; !ok {
				continue
			}
			if _, ok := seen[id]; ok {
				continue
			}
			seen[id] = struct{}{}
			compact = append(compact, id)
		}
		s.order = compact
	}
}

func (s *Store) pruneRevisionsLocked() {
	if len(s.revisions) <= s.maxRevisions {
		return
	}
	for id, state := range s.revisions {
		if len(s.revisions) <= s.maxRevisions {
			break
		}
		if state.refs == 0 && id != s.latest {
			delete(s.revisions, id)
		}
	}
}
