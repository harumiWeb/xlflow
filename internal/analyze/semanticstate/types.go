// Package semanticstate provides the procedure-local, indexed fixed-point
// state used by semantic analysis kernels.
//
// The package deliberately knows only about CFG identity and lattice state. It
// does not interpret VBA values or diagnostic evidence. A procedure creates an
// Environment and Index once, then reuses a Solver and its scratch states for
// the duration of that procedure revision.
package semanticstate

import (
	"context"
	"fmt"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
)

// SymbolID is an identity local to one procedure revision. It is intentionally
// not a source-level identifier and must not be persisted across revisions.
type SymbolID uint32

// BlockOrdinal is the deterministic index of a CFG block in one revision.
type BlockOrdinal uint32

// LaneOrdinal identifies an independent semantic policy in a solver run.
type LaneOrdinal uint16

// Representation is the physical representation selected for propagated
// slots. Hybrid selection means that the selector chooses Dense or Sparse per
// procedure; there is no third pointer-heavy representation on the hot path.
type Representation uint8

const (
	RepresentationDense Representation = iota
	RepresentationSparse
)

func (r Representation) String() string {
	if r == RepresentationSparse {
		return "sparse"
	}
	return "dense"
}

// Layout is immutable metadata shared by all states in one procedure.
type Layout struct {
	symbols        int
	touched        int
	representation Representation
}

// NewLayout applies the deterministic hybrid selector from ADR-0047. A
// domain with at most 64 symbols, or one touching at least 25% of its symbols,
// uses dense slots. touched is a count of distinct symbols known to be touched
// by the domain, not a number of transfer events.
func NewLayout(symbolCount, touched int) Layout {
	if symbolCount < 0 {
		symbolCount = 0
	}
	if touched < 0 {
		touched = 0
	}
	if touched > symbolCount {
		touched = symbolCount
	}
	representation := RepresentationSparse
	if symbolCount <= 64 || touched*4 >= symbolCount {
		representation = RepresentationDense
	}
	return Layout{symbols: symbolCount, touched: touched, representation: representation}
}

// SymbolCount returns the number of indexed symbols.
func (l Layout) SymbolCount() int { return l.symbols }

// TouchedCount returns the number used by the hybrid representation selector.
func (l Layout) TouchedCount() int { return l.touched }

// Representation reports the selected physical representation.
func (l Layout) Representation() Representation { return l.representation }

// Environment is immutable revision-local name/index data. Methods return
// copies where a slice is involved, so callers cannot mutate the index through
// an accessor.
type Environment struct {
	names  []string
	lookup map[string]SymbolID
	layout Layout
}

// NewEnvironment canonicalizes case-insensitive domain names, removes
// duplicates, and assigns IDs in canonical name order. An optional touched-name
// slice supplies the density estimate used by the hybrid selector.
func NewEnvironment(names []string, touchedNames ...[]string) Environment {
	canonical := make([]string, 0, len(names))
	seen := make(map[string]struct{}, len(names))
	for _, name := range names {
		name = canonicalName(name)
		if name == "" {
			continue
		}
		if _, ok := seen[name]; ok {
			continue
		}
		seen[name] = struct{}{}
		canonical = append(canonical, name)
	}
	sort.Strings(canonical)
	lookup := make(map[string]SymbolID, len(canonical))
	for id, name := range canonical {
		lookup[name] = SymbolID(id)
	}
	touched := 0
	if len(touchedNames) != 0 {
		touchedSeen := make(map[SymbolID]struct{}, len(touchedNames[0]))
		for _, name := range touchedNames[0] {
			if id, ok := lookup[canonicalName(name)]; ok {
				touchedSeen[id] = struct{}{}
			}
		}
		touched = len(touchedSeen)
	}
	return Environment{
		names:  canonical,
		lookup: lookup,
		layout: NewLayout(len(canonical), touched),
	}
}

// Layout returns immutable slot metadata for this revision.
func (e Environment) Layout() Layout { return e.layout }

// Symbol resolves a canonicalized name to its revision-local ID.
func (e Environment) Symbol(name string) (SymbolID, bool) {
	id, ok := e.lookup[canonicalName(name)]
	return id, ok
}

// Name returns the canonical name for an ID.
func (e Environment) Name(id SymbolID) (string, bool) {
	if int(id) >= len(e.names) {
		return "", false
	}
	return e.names[id], true
}

// Names returns canonical names in SymbolID order.
func (e Environment) Names() []string { return append([]string(nil), e.names...) }

func canonicalName(name string) string {
	name = strings.TrimSpace(name)
	name = strings.Trim(name, "[]")
	return strings.ToLower(name)
}

// Block is a stable index reference to a CFG block.
type Block struct {
	Ordinal BlockOrdinal
	ID      cfg.BlockID
}

// Edge is the protocol-neutral edge view consumed by a lane policy.
type Edge struct {
	ID        cfg.EdgeID
	From      BlockOrdinal
	To        BlockOrdinal
	Kind      cfg.EdgeKind
	Class     cfg.EdgeClass
	Uncertain bool
}

// Index is immutable CFG identity and outgoing-edge data. The input graph is
// copied while constructing the index; later graph mutations cannot affect a
// solver invocation.
type Index struct {
	blocks   []Block
	byID     map[cfg.BlockID]BlockOrdinal
	outgoing [][]Edge
	entry    BlockOrdinal
	entryID  cfg.BlockID
}

// NewIndex builds deterministic ordinals by sorting the graph's stable block
// IDs. Outgoing edges are ordered by stable edge ID (then destination and edge
// kind as tie-breakers) to keep adapter traversal reproducible while retaining
// CFG builder branch precedence.
func NewIndex(graph cfg.Graph) (*Index, error) {
	if len(graph.Blocks) == 0 {
		return nil, fmt.Errorf("semanticstate: CFG has no blocks")
	}
	ids := make([]cfg.BlockID, 0, len(graph.Blocks))
	seen := make(map[cfg.BlockID]struct{}, len(graph.Blocks))
	for _, block := range graph.Blocks {
		if _, ok := seen[block.ID]; ok {
			return nil, fmt.Errorf("semanticstate: duplicate CFG block %d", block.ID)
		}
		seen[block.ID] = struct{}{}
		ids = append(ids, block.ID)
	}
	sort.Slice(ids, func(i, j int) bool { return ids[i] < ids[j] })
	byID := make(map[cfg.BlockID]BlockOrdinal, len(ids))
	blocks := make([]Block, len(ids))
	for ordinal, id := range ids {
		byID[id] = BlockOrdinal(ordinal)
		blocks[ordinal] = Block{Ordinal: BlockOrdinal(ordinal), ID: id}
	}
	entry, ok := byID[graph.Entry]
	if !ok {
		return nil, fmt.Errorf("semanticstate: CFG entry %d is not a block", graph.Entry)
	}
	outgoing := make([][]Edge, len(blocks))
	for _, source := range graph.Edges {
		from, fromOK := byID[source.From]
		to, toOK := byID[source.To]
		if !fromOK || !toOK {
			return nil, fmt.Errorf("semanticstate: edge %d references unknown block", source.ID)
		}
		outgoing[from] = append(outgoing[from], Edge{
			ID: source.ID, From: from, To: to, Kind: source.Kind,
			Class: source.Class, Uncertain: source.Uncertain,
		})
	}
	for i := range outgoing {
		sort.SliceStable(outgoing[i], func(a, b int) bool {
			left, right := outgoing[i][a], outgoing[i][b]
			// CFG builders assign stable edge IDs in source order. Keeping that
			// order preserves domain-specific branch precedence while still
			// making the traversal independent of the input slice order.
			if left.ID != right.ID {
				return left.ID < right.ID
			}
			if left.To != right.To {
				return left.To < right.To
			}
			if left.Kind != right.Kind {
				return left.Kind < right.Kind
			}
			if left.Class != right.Class {
				return left.Class < right.Class
			}
			return !left.Uncertain && right.Uncertain
		})
	}
	return &Index{blocks: blocks, byID: byID, outgoing: outgoing, entry: entry, entryID: graph.Entry}, nil
}

// BlockCount returns the number of blocks in the indexed graph.
func (i *Index) BlockCount() int {
	if i == nil {
		return 0
	}
	return len(i.blocks)
}

// Entry returns the indexed entry block.
func (i *Index) Entry() BlockOrdinal { return i.entry }

// Block returns a stable block reference.
func (i *Index) Block(ordinal BlockOrdinal) (Block, bool) {
	if i == nil || int(ordinal) >= len(i.blocks) {
		return Block{}, false
	}
	return i.blocks[ordinal], true
}

// Ordinal resolves a source CFG block ID.
func (i *Index) Ordinal(id cfg.BlockID) (BlockOrdinal, bool) {
	if i == nil {
		return 0, false
	}
	ordinal, ok := i.byID[id]
	return ordinal, ok
}

// Blocks returns stable block references in ordinal order.
func (i *Index) Blocks() []Block {
	if i == nil {
		return nil
	}
	return append([]Block(nil), i.blocks...)
}

// Outgoing returns one block's immutable deterministic edge list. Index is
// immutable after construction; callers must treat the returned slice as
// read-only so hot solver iterations do not allocate a defensive copy.
func (i *Index) Outgoing(from BlockOrdinal) []Edge {
	if i == nil || int(from) >= len(i.outgoing) {
		return nil
	}
	return i.outgoing[from]
}

// Lattice is a finite-lattice adapter. Join must mutate *dst in place and
// report whether the destination changed. Clone must produce an independent
// value suitable for storing in another state; it is also the alias-safety
// boundary for slices, maps, and pointers held by a domain value.
type Lattice[T any] interface {
	Clone(T) T
	Join(dst *T, src T) bool
}

// StateView is a read-only view valid until the owning solver advances or
// reuses that state. Values are returned by value; the lattice Clone method is
// responsible for making pointer-bearing values safe before storing them.
type StateView[T any] struct{ state *State[T] }

// Has reports whether a slot is present (as opposed to lattice bottom).
func (v StateView[T]) Has(id SymbolID) bool {
	return v.state != nil && v.state.has(id)
}

// Value returns a present slot value.
func (v StateView[T]) Value(id SymbolID) (T, bool) {
	if v.state == nil {
		var zero T
		return zero, false
	}
	return v.state.value(id)
}

// IDs returns present IDs in ascending order.
func (v StateView[T]) IDs() []SymbolID {
	if v.state == nil {
		return nil
	}
	return v.state.ids()
}

// Len returns the number of present slots.
func (v StateView[T]) Len() int {
	if v.state == nil {
		return 0
	}
	return v.state.len()
}

// ForEach visits present slots in SymbolID order. Returning false stops the
// traversal.
func (v StateView[T]) ForEach(fn func(SymbolID, T) bool) {
	if v.state != nil {
		v.state.forEach(fn)
	}
}

// NewState creates an empty state using an immutable layout.
func NewState[T any](layout Layout) State[T] {
	return newState[T](layout)
}

// State is a compact indexed set of slots. Use StateView when passing state to
// transfer callbacks; mutation is reserved for destination and scratch states.
type State[T any] struct {
	layout  Layout
	dense   []T
	present []uint64
	sparse  []slot[T]
}

type slot[T any] struct {
	id    SymbolID
	value T
}

func newState[T any](layout Layout) State[T] {
	s := State[T]{layout: layout}
	if layout.representation == RepresentationDense {
		s.dense = make([]T, layout.symbols)
		s.present = make([]uint64, (layout.symbols+63)/64)
	}
	return s
}

// Layout reports the state's immutable layout.
func (s *State[T]) Layout() Layout {
	if s == nil {
		return Layout{}
	}
	return s.layout
}

// Representation reports the state's selected physical representation.
func (s *State[T]) Representation() Representation { return s.layout.representation }

// View returns a read-only view over the state.
func (s *State[T]) View() StateView[T] { return StateView[T]{state: s} }

// Reset clears all slots while retaining backing storage for reuse.
func (s *State[T]) Reset() {
	if s == nil {
		return
	}
	var zero T
	if s.layout.representation == RepresentationDense {
		for id := range s.dense {
			if s.bit(id) {
				s.dense[id] = zero
			}
		}
		clear(s.present)
		return
	}
	for i := range s.sparse {
		s.sparse[i].value = zero
	}
	s.sparse = s.sparse[:0]
}

// Set stores a slot value. It returns false when id is outside the layout.
// Callers that need an independent value should pass a lattice.Clone result.
func (s *State[T]) Set(id SymbolID, value T) bool {
	if s == nil || int(id) >= s.layout.symbols {
		return false
	}
	if s.layout.representation == RepresentationDense {
		s.dense[id] = value
		s.setBit(int(id))
		return true
	}
	index, found := s.sparseIndex(id)
	if found {
		s.sparse[index].value = value
		return true
	}
	s.sparse = append(s.sparse, slot[T]{})
	copy(s.sparse[index+1:], s.sparse[index:])
	s.sparse[index] = slot[T]{id: id, value: value}
	return true
}

// Delete removes one present slot.
func (s *State[T]) Delete(id SymbolID) bool {
	if s == nil || int(id) >= s.layout.symbols || !s.has(id) {
		return false
	}
	var zero T
	if s.layout.representation == RepresentationDense {
		s.dense[id] = zero
		s.clearBit(int(id))
		return true
	}
	index, _ := s.sparseIndex(id)
	s.sparse[index].value = zero
	copy(s.sparse[index:], s.sparse[index+1:])
	s.sparse = s.sparse[:len(s.sparse)-1]
	return true
}

// CloneFrom replaces the destination with an independent copy of src.
func (s *State[T]) CloneFrom(src StateView[T], clone func(T) T) {
	s.Reset()
	if clone == nil {
		clone = func(value T) T { return value }
	}
	src.ForEach(func(id SymbolID, value T) bool {
		s.Set(id, clone(value))
		return true
	})
}

// JoinFrom joins src into the destination in place. Each changed SymbolID is
// appended to changed, which is caller-owned and can be reused across edges.
// The bool result reports whether at least one slot changed.
func (s *State[T]) JoinFrom(src StateView[T], lattice Lattice[T], changed *[]SymbolID) bool {
	if s == nil || lattice == nil {
		return false
	}
	changedAny := false
	src.ForEach(func(id SymbolID, value T) bool {
		if !s.has(id) {
			s.Set(id, lattice.Clone(value))
			if changed != nil {
				*changed = append(*changed, id)
			}
			changedAny = true
			return true
		}
		current, _ := s.value(id)
		if lattice.Join(&current, value) {
			s.Set(id, current)
			if changed != nil {
				*changed = append(*changed, id)
			}
			changedAny = true
		}
		return true
	})
	return changedAny
}

func (s *State[T]) has(id SymbolID) bool {
	if s == nil || int(id) >= s.layout.symbols {
		return false
	}
	if s.layout.representation == RepresentationDense {
		return s.bit(int(id))
	}
	_, found := s.sparseIndex(id)
	return found
}

func (s *State[T]) value(id SymbolID) (T, bool) {
	if !s.has(id) {
		var zero T
		return zero, false
	}
	if s.layout.representation == RepresentationDense {
		return s.dense[id], true
	}
	index, _ := s.sparseIndex(id)
	return s.sparse[index].value, true
}

func (s *State[T]) len() int {
	if s.layout.representation == RepresentationDense {
		n := 0
		for _, bits := range s.present {
			n += bitsOnes(bits)
		}
		return n
	}
	return len(s.sparse)
}

func (s *State[T]) ids() []SymbolID {
	ids := make([]SymbolID, 0, s.len())
	s.forEach(func(id SymbolID, _ T) bool { ids = append(ids, id); return true })
	return ids
}

func (s *State[T]) forEach(fn func(SymbolID, T) bool) {
	if fn == nil {
		return
	}
	if s.layout.representation == RepresentationDense {
		for id := 0; id < len(s.dense); id++ {
			if s.bit(id) && !fn(SymbolID(id), s.dense[id]) {
				return
			}
		}
		return
	}
	for _, entry := range s.sparse {
		if !fn(entry.id, entry.value) {
			return
		}
	}
}

func (s *State[T]) sparseIndex(id SymbolID) (int, bool) {
	index := sort.Search(len(s.sparse), func(i int) bool { return s.sparse[i].id >= id })
	return index, index < len(s.sparse) && s.sparse[index].id == id
}

func (s *State[T]) bit(index int) bool {
	return s.present[index/64]&(uint64(1)<<uint(index%64)) != 0
}

func (s *State[T]) setBit(index int) {
	s.present[index/64] |= uint64(1) << uint(index%64)
}

func (s *State[T]) clearBit(index int) {
	s.present[index/64] &^= uint64(1) << uint(index%64)
}

func bitsOnes(value uint64) int {
	count := 0
	for value != 0 {
		value &= value - 1
		count++
	}
	return count
}

// TransferFunc advances one lane's block transfer. out is a reusable scratch
// state and is reset by the solver before the callback runs.
type TransferFunc[T any] func(context.Context, LaneOrdinal, BlockOrdinal, StateView[T], *State[T]) error

// EdgePolicy refines a propagated state for one edge. candidate is initialized
// from predecessor output on normal edges and predecessor input on exceptional
// or uncertain edges. A policy may explicitly copy from predecessorOutput when
// a domain has a narrow exception that remains valid across that edge.
type EdgePolicy[T any] func(context.Context, LaneOrdinal, Edge, StateView[T], StateView[T], *State[T]) error

// InitializeFunc seeds the entry state for one lane.
type InitializeFunc[T any] func(context.Context, LaneOrdinal, *State[T]) error

// Lane defines an independent transfer and edge policy. Lanes share indexed
// scheduling but never share mutable state.
type Lane[T any] struct {
	Transfer   TransferFunc[T]
	Edge       EdgePolicy[T]
	Initialize InitializeFunc[T]
}
