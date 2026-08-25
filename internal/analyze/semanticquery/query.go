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
	"reflect"
	"sort"
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
	value any
}

// dependencyRecord is an eviction-tracked placeholder for a query whose
// value was published outside Evaluate. It keeps RecordDependencies metadata
// bounded without making a metadata-only record look like a cache hit.
type dependencyRecord struct{}

type pendingValue struct {
	done chan struct{}
}

// dependencySet keeps one parent's dependencies in canonical order and
// retains an order-independent digest for the common hit-path rejection. The
// exact membership check remains as a collision guard, but it does not build a
// temporary map or stringify keys.
type dependencySet struct {
	ids    []Key
	digest [32]byte
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

	// Key is comparable, so keep it as the cache identity directly.  The
	// previous string identity required building a NUL-delimited string for
	// every lookup and retained a second copy of every key in keys.
	entries map[Key]cacheValue
	order   []Key
	pending map[Key]*pendingValue
	reverse map[Key]map[Key]struct{}
	epochs  map[Key]uint64
	deps    map[Key]dependencySet
	stale   map[Key]map[Key]struct{}

	// procedureEntries contains exact procedure identities and their
	// structural prefixes (document, document::module, and so on).  Keeping
	// these indexes means document/procedure invalidation does not scan every
	// query key. documentEntries is kept separately so a normalized document
	// path has an explicit, exact lookup path.
	procedureEntries map[string]map[Key]struct{}
	documentEntries  map[string]map[Key]struct{}

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
		entries:          make(map[Key]cacheValue),
		pending:          make(map[Key]*pendingValue),
		reverse:          make(map[Key]map[Key]struct{}),
		epochs:           make(map[Key]uint64),
		deps:             make(map[Key]dependencySet),
		stale:            make(map[Key]map[Key]struct{}),
		procedureEntries: make(map[string]map[Key]struct{}),
		documentEntries:  make(map[string]map[Key]struct{}),
		revisions:        make(map[string]*revisionState),
		maxEntries:       maxEntries,
		maxRevisions:     maxRevisions,
		metrics:          opts.Metrics,
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
	stable := key
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
			if _, stale := s.stale[stable]; !stale {
				if !s.dependenciesMatchLocked(stable, dependencies) {
					s.invalidateIDsLocked([]Key{stable}, metrics)
					s.mu.Unlock()
					continue
				}
				s.mu.Unlock()
				if metrics != nil {
					metrics.RecordSemanticQueryHit()
				}
				return cached.value, true, nil
			}
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
		previous, hadPrevious := s.entries[stable]
		if hadPrevious {
			if _, metadataOnly := previous.value.(dependencyRecord); metadataOnly {
				hadPrevious = false
			}
		}
		staleReasons := append([]Key(nil), s.staleReasonsLocked(stable)...)
		s.pending[stable] = pending
		s.indexKeyLocked(stable)
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
					s.maybeRemoveKeyMetadataLocked(stable)
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
			parents := make([]Key, 0, len(s.reverse[stable]))
			for parent := range s.reverse[stable] {
				parents = append(parents, parent)
			}
			hadEntry := false
			if _, ok := s.entries[stable]; ok {
				hadEntry = true
			}
			s.entries[stable] = cacheValue{value: value}
			if !hadEntry {
				s.order = append(s.order, stable)
			}
			s.recordDependenciesLocked(stable, key, dependencies)
			unchanged := len(staleReasons) > 0 && hadPrevious && reflect.DeepEqual(previous.value, value)
			if len(staleReasons) > 0 {
				for _, parent := range parents {
					// Dependents are red because this immediate producer was
					// invalidated, not because of the producer's own upstream
					// reason. Clear or retain that edge-local reason only.
					s.clearStaleReasonLocked(parent, stable)
					if !unchanged {
						s.markStaleReasonLocked(parent, stable)
					}
				}
			}
			s.clearStaleLocked(stable)
			s.pruneEntriesLocked()
			if metrics != nil {
				metrics.RecordSemanticQueryRecomputedKernel()
			}
		}
		epochChanged := s.epochs[stable] != epoch
		if err != nil || epochChanged {
			s.maybeRemoveKeyMetadataLocked(stable)
		}
		close(pending.done)
		s.mu.Unlock()
		if err != nil {
			return nil, false, err
		}
		if epochChanged {
			// A newer revision invalidated this build while it was running. Do
			// not expose the uncommitted value; retry against the current epoch.
			continue
		}
		return value, false, nil
	}
}

func (s *Store) dependenciesMatchLocked(parent Key, dependencies []Key) bool {
	stored := s.deps[parent]
	if len(stored.ids) != len(dependencies) || stored.digest != dependencySetDigest(dependencies) {
		return false
	}
	// Stored dependencies are deduplicated by recordDependenciesLocked. A
	// membership comparison is deliberately allocation-free on the hit path;
	// dependency slices are normally tiny, and order is not semantically
	// significant. The digest above rejects almost all changed sets before this
	// collision-safe comparison runs.
	for _, dependencyID := range stored.ids {
		found := false
		for _, dependency := range dependencies {
			if dependency == dependencyID {
				found = true
				break
			}
		}
		if !found {
			return false
		}
	}
	return true
}

func dependencyKeyDigest(key Key) [32]byte {
	h := sha256.New()
	for _, part := range []string{key.Procedure, key.Fingerprint, key.Kernel, key.Config, key.Capability} {
		_, _ = fmt.Fprintf(h, "%d:", len(part))
		_, _ = h.Write([]byte(part))
	}
	var digest [32]byte
	copy(digest[:], h.Sum(nil))
	return digest
}

func dependencySetDigest(dependencies []Key) [32]byte {
	var digest [32]byte
	for _, dependency := range dependencies {
		item := dependencyKeyDigest(dependency)
		for index := range digest {
			digest[index] ^= item[index]
		}
	}
	return digest
}

func (s *Store) staleReasonsLocked(id Key) []Key {
	reasons := s.stale[id]
	if len(reasons) == 0 {
		return nil
	}
	out := make([]Key, 0, len(reasons))
	for reason := range reasons {
		out = append(out, reason)
	}
	return out
}

// markStaleReasonLocked marks a cached value red because the producer named
// by reason changed. Reasons are retained so an unchanged producer output can
// clear only the invalidation it caused while preserving unrelated red state.
func (s *Store) markStaleReasonLocked(id, reason Key) {
	reasons := s.stale[id]
	if reasons == nil {
		if _, hasEntry := s.entries[id]; !hasEntry {
			// A pending or not-yet-created node still needs an epoch bump to
			// reject a late publication, but it has no value to mark stale.
			s.epochs[id]++
			return
		}
		reasons = make(map[Key]struct{})
		s.stale[id] = reasons
	}
	reasons[reason] = struct{}{}
	// Repeated invalidations of the same producer still advance the epoch so
	// an in-flight build cannot publish across the newer source snapshot.
	s.epochs[id]++
}

func (s *Store) clearStaleReasonLocked(id, reason Key) {
	reasons := s.stale[id]
	if len(reasons) == 0 {
		return
	}
	delete(reasons, reason)
	if len(reasons) == 0 {
		delete(s.stale, id)
	}
}

func (s *Store) clearStaleLocked(id Key) {
	delete(s.stale, id)
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

func (s *Store) recordDependenciesLocked(parent Key, parentKey Key, dependencies []Key) {
	for _, dependencyID := range s.deps[parent].ids {
		if parents := s.reverse[dependencyID]; parents != nil {
			delete(parents, parent)
			if len(parents) == 0 {
				delete(s.reverse, dependencyID)
				s.maybeRemoveKeyMetadataLocked(dependencyID)
			}
		}
	}
	s.indexKeyLocked(parentKey)
	ids := make([]Key, 0, len(dependencies))
	for _, dependency := range dependencies {
		dependencyID := dependency
		duplicate := false
		for _, existing := range ids {
			if existing == dependencyID {
				duplicate = true
				break
			}
		}
		if duplicate {
			continue
		}
		ids = append(ids, dependencyID)
		s.indexKeyLocked(dependency)
		parents := s.reverse[dependencyID]
		if parents == nil {
			parents = make(map[Key]struct{})
			s.reverse[dependencyID] = parents
		}
		parents[parent] = struct{}{}
	}
	sort.Slice(ids, func(i, j int) bool { return keyLess(ids[i], ids[j]) })
	s.deps[parent] = dependencySet{ids: ids, digest: dependencySetDigest(ids)}
}

func keyLess(left, right Key) bool {
	if left.Procedure != right.Procedure {
		return left.Procedure < right.Procedure
	}
	if left.Fingerprint != right.Fingerprint {
		return left.Fingerprint < right.Fingerprint
	}
	if left.Kernel != right.Kernel {
		return left.Kernel < right.Kernel
	}
	if left.Config != right.Config {
		return left.Config < right.Config
	}
	return left.Capability < right.Capability
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
	stable := parent
	if _, ok := s.entries[stable]; !ok {
		s.entries[stable] = cacheValue{value: dependencyRecord{}}
		s.order = append(s.order, stable)
	}
	s.indexKeyLocked(parent)
	s.recordDependenciesLocked(stable, parent, dependencies)
	s.pruneEntriesLocked()
	s.mu.Unlock()
}

// Invalidate marks changed queries and their transitive dependents red. The
// returned procedure identities are sorted and unique for deterministic
// telemetry. Reverse traversal reports the full affected closure, while
// values are reevaluated lazily: an unchanged producer output clears its
// dependents back to green instead of forcing the whole closure to rebuild.
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
	queue := append([]Key(nil), changed...)
	return s.invalidateIDsLocked(queue, s.metricsFor(ctx))
}

// InvalidateProcedures marks all cached queries belonging to the supplied
// procedure identities (or identity prefixes), plus their reverse dependents,
// red. LSP callers use path prefixes because a source edit can replace a
// procedure declaration and therefore cannot always reconstruct the old full
// key.
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
	queue := make([]Key, 0)
	for _, procedure := range procedures {
		if procedure == "" {
			continue
		}
		for id := range s.procedureEntries[procedure] {
			queue = append(queue, id)
		}
		for id := range s.documentEntries[procedure] {
			queue = append(queue, id)
		}
	}
	return s.invalidateIDsLocked(queue, s.metricsFor(ctx))
}

func (s *Store) invalidateIDsLocked(queue []Key, metrics Metrics) []string {
	// Stale reasons are the immediate producer keys on each reverse edge. A
	// shared sentinel for a multi-root invalidation would let an unchanged
	// rebuild of root A clear the same reason from a dependent that is still
	// stale because of root B; edge-local reasons preserve recovery while this
	// remains one reverse traversal.
	roots := make([]Key, 0, len(queue))
	rootSeen := make(map[Key]struct{}, len(queue))
	for _, root := range queue {
		if _, ok := rootSeen[root]; ok {
			continue
		}
		rootSeen[root] = struct{}{}
		roots = append(roots, root)
	}
	seen := make(map[Key]bool)
	procedures := make(map[string]bool)
	affected := make([]Key, 0, len(roots))
	rootQueue := append([]Key(nil), roots...)
	for _, root := range roots {
		if seen[root] {
			continue
		}
		seen[root] = true
		affected = append(affected, root)
		if root.Procedure != "" {
			procedures[root.Procedure] = true
		}
		// A root has no invalidating edge in the supplied set, so use its own
		// identity as the reason that its cached value is red.
		s.markStaleReasonLocked(root, root)
	}
	for len(rootQueue) > 0 {
		id := rootQueue[0]
		rootQueue = rootQueue[1:]
		for parent := range s.reverse[id] {
			// The parent is red because this immediate producer is red. If
			// another changed root reaches the same parent, it contributes a
			// separate reason rather than sharing a batch sentinel.
			s.markStaleReasonLocked(parent, id)
			if seen[parent] {
				continue
			}
			seen[parent] = true
			affected = append(affected, parent)
			if parent.Procedure != "" {
				procedures[parent.Procedure] = true
			}
			rootQueue = append(rootQueue, parent)
		}
	}
	if len(affected) == 0 {
		if metrics != nil {
			metrics.RecordSemanticQueryInvalidatedProcedures(0)
		}
		return nil
	}
	for _, id := range affected {
		s.maybeRemoveKeyMetadataLocked(id)
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

func (s *Store) removeNodeLocked(id Key) {
	for _, dependencyID := range s.deps[id].ids {
		if parents := s.reverse[dependencyID]; parents != nil {
			delete(parents, id)
			if len(parents) == 0 {
				delete(s.reverse, dependencyID)
				s.maybeRemoveKeyMetadataLocked(dependencyID)
			}
		}
	}
	delete(s.deps, id)
	delete(s.entries, id)
	delete(s.stale, id)
	if len(s.reverse[id]) == 0 && s.pending[id] == nil && len(s.deps[id].ids) == 0 {
		delete(s.reverse, id)
		s.removeKeyIndexLocked(id)
		delete(s.epochs, id)
	}
}

func (s *Store) maybeRemoveKeyMetadataLocked(id Key) {
	if _, ok := s.entries[id]; ok {
		return
	}
	if s.pending[id] != nil || len(s.reverse[id]) != 0 || len(s.deps[id].ids) != 0 {
		return
	}
	s.removeKeyIndexLocked(id)
	delete(s.epochs, id)
	delete(s.stale, id)
}

func addKeyIndex(index map[string]map[Key]struct{}, name string, key Key) {
	if name == "" {
		return
	}
	keys := index[name]
	if keys == nil {
		keys = make(map[Key]struct{})
		index[name] = keys
	}
	keys[key] = struct{}{}
}

func removeKeyIndex(index map[string]map[Key]struct{}, name string, key Key) {
	if keys := index[name]; keys != nil {
		delete(keys, key)
		if len(keys) == 0 {
			delete(index, name)
		}
	}
}

// indexKeyLocked records all stable structural identities understood by
// InvalidateProcedures. Prefix strings are slices of the procedure identity,
// so indexing does not need another serialization of the Key itself.
func (s *Store) indexKeyLocked(key Key) {
	procedure := key.Procedure
	if procedure == "" {
		return
	}
	addKeyIndex(s.procedureEntries, procedure, key)
	first := strings.Index(procedure, "::")
	if first < 0 {
		addKeyIndex(s.documentEntries, procedure, key)
		return
	}
	addKeyIndex(s.documentEntries, procedure[:first], key)
	for end := first; end >= 0; {
		addKeyIndex(s.procedureEntries, procedure[:end], key)
		next := strings.Index(procedure[end+2:], "::")
		if next < 0 {
			break
		}
		end += 2 + next
	}
}

func (s *Store) removeKeyIndexLocked(key Key) {
	procedure := key.Procedure
	if procedure == "" {
		return
	}
	removeKeyIndex(s.procedureEntries, procedure, key)
	first := strings.Index(procedure, "::")
	if first < 0 {
		removeKeyIndex(s.documentEntries, procedure, key)
		return
	}
	removeKeyIndex(s.documentEntries, procedure[:first], key)
	for end := first; end >= 0; {
		removeKeyIndex(s.procedureEntries, procedure[:end], key)
		next := strings.Index(procedure[end+2:], "::")
		if next < 0 {
			break
		}
		end += 2 + next
	}
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
		compact := make([]Key, 0, len(s.entries))
		seen := make(map[Key]struct{}, len(s.entries))
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

