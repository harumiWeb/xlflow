package lspserver

import (
	"crypto/sha256"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

type documentSymbolSignature struct {
	version    int32
	sourceHash [sha256.Size]byte
	moduleKind string
	uri        string
}

type documentSymbolState struct {
	generation uint64
	signature  documentSymbolSignature
	valid      bool
}

type cachedDocumentSymbols struct {
	generation uint64
	signature  documentSymbolSignature
	symbols    []intel.Symbol
}

type documentSymbolCall struct {
	done    chan struct{}
	symbols []intel.Symbol
	err     error
}

type documentSymbolWorkKey struct {
	identity   string
	generation uint64
}

type documentSymbolCache struct {
	mu       sync.Mutex
	states   map[string]documentSymbolState
	entries  map[string]cachedDocumentSymbols
	inflight map[documentSymbolWorkKey]*documentSymbolCall
}

func newDocumentSymbolCache() *documentSymbolCache {
	return &documentSymbolCache{
		states:   make(map[string]documentSymbolState),
		entries:  make(map[string]cachedDocumentSymbols),
		inflight: make(map[documentSymbolWorkKey]*documentSymbolCall),
	}
}

func (c *documentSymbolCache) get(doc intel.Document, load func() ([]intel.Symbol, error)) ([]intel.Symbol, bool, error) {
	identity := documentSymbolKey(doc)
	if identity == "" {
		syms, err := load()
		return cloneSymbols(syms), false, err
	}
	signature := documentSymbolSignature{
		version:    doc.Version,
		sourceHash: sha256.Sum256([]byte(doc.Source)),
		moduleKind: doc.ModuleKind,
		uri:        doc.URI,
	}

	c.mu.Lock()
	state := c.states[identity]
	if !state.valid || state.signature != signature {
		state.generation++
		state.signature = signature
		state.valid = true
		c.states[identity] = state
		delete(c.entries, identity)
	}
	if entry, ok := c.entries[identity]; ok &&
		entry.generation == state.generation && entry.signature == signature {
		out := cloneSymbols(entry.symbols)
		c.mu.Unlock()
		return out, true, nil
	}

	workKey := documentSymbolWorkKey{identity: identity, generation: state.generation}
	if call, ok := c.inflight[workKey]; ok {
		c.mu.Unlock()
		<-call.done
		return cloneSymbols(call.symbols), call.err == nil, call.err
	}
	call := &documentSymbolCall{done: make(chan struct{})}
	c.inflight[workKey] = call
	c.mu.Unlock()

	syms, err := load()
	cloned := cloneSymbols(syms)

	c.mu.Lock()
	delete(c.inflight, workKey)
	current := c.states[identity]
	if err == nil && current.valid && current.generation == state.generation && current.signature == signature {
		c.entries[identity] = cachedDocumentSymbols{
			generation: state.generation,
			signature:  signature,
			symbols:    cloneSymbols(cloned),
		}
	}
	call.symbols = cloned
	call.err = err
	close(call.done)
	c.mu.Unlock()

	return cloneSymbols(cloned), false, err
}

func (c *documentSymbolCache) invalidate(doc intel.Document) {
	identity := documentSymbolKey(doc)
	if identity == "" {
		return
	}
	c.mu.Lock()
	state := c.states[identity]
	state.generation++
	state.valid = false
	c.states[identity] = state
	delete(c.entries, identity)
	c.mu.Unlock()
}
