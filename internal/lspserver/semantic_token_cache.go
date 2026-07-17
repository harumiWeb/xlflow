package lspserver

import (
	"crypto/sha256"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

type semanticTokenSignature struct {
	version    int32
	sourceHash [sha256.Size]byte
	moduleKind string
	uri        string
}

type cachedSemanticTokens struct {
	generation uint64
	signature  semanticTokenSignature
	data       []protocol.UInteger
}

type semanticTokenCall struct {
	done chan struct{}
	data []protocol.UInteger
	err  error
}

type semanticTokenWorkKey struct {
	identity   string
	generation uint64
	signature  semanticTokenSignature
}

type semanticTokenCache struct {
	mu         sync.Mutex
	generation uint64
	signatures map[string]semanticTokenSignature
	entries    map[string]cachedSemanticTokens
	inflight   map[semanticTokenWorkKey]*semanticTokenCall
}

func newSemanticTokenCache() *semanticTokenCache {
	return &semanticTokenCache{
		signatures: make(map[string]semanticTokenSignature),
		entries:    make(map[string]cachedSemanticTokens),
		inflight:   make(map[semanticTokenWorkKey]*semanticTokenCall),
	}
}

func (c *semanticTokenCache) begin() uint64 {
	c.mu.Lock()
	defer c.mu.Unlock()
	return c.generation
}

func (c *semanticTokenCache) get(
	doc intel.Document,
	generation uint64,
	load func() ([]protocol.UInteger, error),
) ([]protocol.UInteger, bool, error) {
	identity := documentSymbolKey(doc)
	if identity == "" {
		data, err := load()
		return cloneSemanticTokenData(data), false, err
	}
	signature := semanticTokenSignature{
		version:    doc.Version,
		sourceHash: sha256.Sum256([]byte(doc.Source)),
		moduleKind: doc.ModuleKind,
		uri:        doc.URI,
	}

	c.mu.Lock()
	if generation == c.generation {
		c.signatures[identity] = signature
		if entry, ok := c.entries[identity]; ok &&
			entry.generation == generation && entry.signature == signature {
			out := cloneSemanticTokenData(entry.data)
			c.mu.Unlock()
			return out, true, nil
		}
	}
	workKey := semanticTokenWorkKey{identity: identity, generation: generation, signature: signature}
	if call, ok := c.inflight[workKey]; ok {
		c.mu.Unlock()
		<-call.done
		return cloneSemanticTokenData(call.data), call.err == nil, call.err
	}
	call := &semanticTokenCall{done: make(chan struct{})}
	c.inflight[workKey] = call
	c.mu.Unlock()

	data, err := load()
	cloned := cloneSemanticTokenData(data)

	c.mu.Lock()
	delete(c.inflight, workKey)
	if err == nil && generation == c.generation && c.signatures[identity] == signature {
		c.entries[identity] = cachedSemanticTokens{
			generation: generation,
			signature:  signature,
			data:       cloneSemanticTokenData(cloned),
		}
	}
	call.data = cloned
	call.err = err
	close(call.done)
	c.mu.Unlock()

	return cloneSemanticTokenData(cloned), false, err
}

func (c *semanticTokenCache) invalidateAll() {
	c.mu.Lock()
	c.generation++
	c.signatures = make(map[string]semanticTokenSignature)
	c.entries = make(map[string]cachedSemanticTokens)
	c.mu.Unlock()
}

func cloneSemanticTokenData(data []protocol.UInteger) []protocol.UInteger {
	return append([]protocol.UInteger(nil), data...)
}
