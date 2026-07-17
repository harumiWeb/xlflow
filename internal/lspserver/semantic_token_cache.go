package lspserver

import (
	"crypto/sha256"
	"errors"
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
	done       chan struct{}
	generation uint64
	signature  semanticTokenSignature
	data       []protocol.UInteger
	err        error
	waiters    int
}

type semanticTokenCache struct {
	mu         sync.Mutex
	generation uint64
	signatures map[string]semanticTokenSignature
	entries    map[string]cachedSemanticTokens
	inflight   map[string]*semanticTokenCall
}

var errSemanticTokensSuperseded = errors.New("semantic token generation superseded")

func newSemanticTokenCache() *semanticTokenCache {
	return &semanticTokenCache{
		signatures: make(map[string]semanticTokenSignature),
		entries:    make(map[string]cachedSemanticTokens),
		inflight:   make(map[string]*semanticTokenCall),
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
	if generation != c.generation {
		c.mu.Unlock()
		return nil, false, errSemanticTokensSuperseded
	}
	c.signatures[identity] = signature
	if entry, ok := c.entries[identity]; ok &&
		entry.generation == generation && entry.signature == signature {
		out := cloneSemanticTokenData(entry.data)
		c.mu.Unlock()
		return out, true, nil
	}
	if call, ok := c.inflight[identity]; ok {
		call.waiters++
		c.mu.Unlock()
		<-call.done
		if call.generation == generation && call.signature == signature {
			return cloneSemanticTokenData(call.data), call.err == nil, call.err
		}
		return nil, false, errSemanticTokensSuperseded
	}
	call := &semanticTokenCall{done: make(chan struct{}), generation: generation, signature: signature}
	c.inflight[identity] = call
	c.mu.Unlock()

	data, err := load()
	cloned := cloneSemanticTokenData(data)

	c.mu.Lock()
	delete(c.inflight, identity)
	current := generation == c.generation && c.signatures[identity] == signature
	if err == nil && current {
		c.entries[identity] = cachedSemanticTokens{
			generation: generation,
			signature:  signature,
			data:       cloneSemanticTokenData(cloned),
		}
	}
	resultErr := err
	if resultErr == nil && !current {
		resultErr = errSemanticTokensSuperseded
	}
	call.data = cloned
	call.err = resultErr
	close(call.done)
	c.mu.Unlock()

	if resultErr != nil {
		return nil, false, resultErr
	}
	return cloneSemanticTokenData(cloned), false, nil
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
