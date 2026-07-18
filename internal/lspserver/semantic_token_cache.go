package lspserver

import (
	"crypto/sha256"
	"errors"
	"fmt"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const semanticTokenHistoryLimit = 4

type semanticTokenSignature struct {
	version    int32
	sourceHash [sha256.Size]byte
	moduleKind string
	uri        string
}

// cachedSemanticTokens is both the current result for a document revision and
// a retained delta base while that document remains open.
type cachedSemanticTokens struct {
	generation uint64
	revision   uint64
	signature  semanticTokenSignature
	resultID   string
	data       []protocol.UInteger
}

type semanticTokenCall struct {
	done       chan struct{}
	generation uint64
	revision   uint64
	signature  semanticTokenSignature
	result     cachedSemanticTokens
	err        error
	waiters    int
}

type semanticTokenCache struct {
	mu             sync.Mutex
	generation     uint64
	nextResultID   uint64
	revisions      map[string]uint64
	signatures     map[string]semanticTokenSignature
	entries        map[string]cachedSemanticTokens
	histories      map[string][]cachedSemanticTokens
	openIdentities map[string]struct{}
	inflight       map[string]*semanticTokenCall
}

var errSemanticTokensSuperseded = errors.New("semantic token generation superseded")

func newSemanticTokenCache() *semanticTokenCache {
	return &semanticTokenCache{
		signatures:     make(map[string]semanticTokenSignature),
		entries:        make(map[string]cachedSemanticTokens),
		histories:      make(map[string][]cachedSemanticTokens),
		openIdentities: make(map[string]struct{}),
		inflight:       make(map[string]*semanticTokenCall),
		revisions:      make(map[string]uint64),
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
) (cachedSemanticTokens, bool, error) {
	identity := documentSymbolKey(doc)
	if identity == "" {
		data, err := load()
		return cachedSemanticTokens{data: cloneSemanticTokenData(data)}, false, err
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
		return cachedSemanticTokens{}, false, errSemanticTokensSuperseded
	}
	revision := c.revisions[identity]
	c.signatures[identity] = signature
	if entry, ok := c.entries[identity]; ok &&
		entry.generation == generation && entry.revision == revision && entry.signature == signature {
		out := cloneCachedSemanticTokens(entry)
		c.mu.Unlock()
		return out, true, nil
	}
	if call, ok := c.inflight[identity]; ok {
		call.waiters++
		c.mu.Unlock()
		<-call.done
		if call.generation == generation && call.revision == revision && call.signature == signature {
			return cloneCachedSemanticTokens(call.result), call.err == nil, call.err
		}
		return cachedSemanticTokens{}, false, errSemanticTokensSuperseded
	}
	call := &semanticTokenCall{done: make(chan struct{}), generation: generation, revision: revision, signature: signature}
	c.inflight[identity] = call
	c.mu.Unlock()

	data, err := load()
	cloned := cloneSemanticTokenData(data)

	c.mu.Lock()
	delete(c.inflight, identity)
	current := generation == c.generation && c.revisions[identity] == revision && c.signatures[identity] == signature
	resultErr := err
	if err == nil && current {
		c.nextResultID++
		entry := cachedSemanticTokens{
			generation: generation,
			revision:   revision,
			signature:  signature,
			resultID:   fmt.Sprintf("xlflow-semantic-%d", c.nextResultID),
			data:       cloneSemanticTokenData(cloned),
		}
		c.entries[identity] = entry
		if _, open := c.openIdentities[identity]; open {
			c.appendHistoryLocked(identity, entry)
		}
		call.result = cloneCachedSemanticTokens(entry)
	} else if resultErr == nil {
		resultErr = errSemanticTokensSuperseded
	}
	call.err = resultErr
	close(call.done)
	c.mu.Unlock()

	if resultErr != nil {
		return cachedSemanticTokens{}, false, resultErr
	}
	return cloneCachedSemanticTokens(call.result), false, nil
}

func (c *semanticTokenCache) appendHistoryLocked(identity string, entry cachedSemanticTokens) {
	history := append(c.histories[identity], cloneCachedSemanticTokens(entry))
	if len(history) > semanticTokenHistoryLimit {
		history = history[len(history)-semanticTokenHistoryLimit:]
	}
	c.histories[identity] = history
}

// previous returns a delta base only for the same currently open document
// lifecycle. close and open discard its history before another source can use
// the same path identity.
func (c *semanticTokenCache) previous(doc intel.Document, resultID string) (cachedSemanticTokens, bool) {
	identity := documentSymbolKey(doc)
	if identity == "" || resultID == "" {
		return cachedSemanticTokens{}, false
	}
	c.mu.Lock()
	defer c.mu.Unlock()
	for _, entry := range c.histories[identity] {
		if entry.resultID == resultID {
			return cloneCachedSemanticTokens(entry), true
		}
	}
	return cachedSemanticTokens{}, false
}

// open starts a new lifecycle for an LSP-open document. A path can be reused
// after close, but its old result IDs must never become delta bases again.
func (c *semanticTokenCache) open(doc intel.Document) {
	identity := documentSymbolKey(doc)
	if identity == "" {
		return
	}
	c.mu.Lock()
	c.resetDocumentLocked(identity)
	c.openIdentities[identity] = struct{}{}
	c.mu.Unlock()
}

func (c *semanticTokenCache) close(doc intel.Document) {
	identity := documentSymbolKey(doc)
	if identity == "" {
		return
	}
	c.mu.Lock()
	c.resetDocumentLocked(identity)
	delete(c.openIdentities, identity)
	c.mu.Unlock()
}

func (c *semanticTokenCache) resetDocumentLocked(identity string) {
	c.revisions[identity]++
	delete(c.signatures, identity)
	delete(c.entries, identity)
	delete(c.histories, identity)
}

func (c *semanticTokenCache) invalidateAll() {
	c.mu.Lock()
	c.generation++
	c.signatures = make(map[string]semanticTokenSignature)
	c.entries = make(map[string]cachedSemanticTokens)
	c.histories = make(map[string][]cachedSemanticTokens)
	c.mu.Unlock()
}

// invalidateWorkspace supersedes current token results that may depend on the
// open workspace while retaining bounded per-document histories as delta bases.
func (c *semanticTokenCache) invalidateWorkspace() {
	c.mu.Lock()
	c.generation++
	c.signatures = make(map[string]semanticTokenSignature)
	c.entries = make(map[string]cachedSemanticTokens)
	c.mu.Unlock()
}

// invalidate retires semantic-token state for one document, including retained
// delta bases. It is used when that document is no longer trustworthy.
func (c *semanticTokenCache) invalidate(doc intel.Document) {
	c.close(doc)
}

func cloneCachedSemanticTokens(entry cachedSemanticTokens) cachedSemanticTokens {
	entry.data = cloneSemanticTokenData(entry.data)
	return entry
}

func cloneSemanticTokenData(data []protocol.UInteger) []protocol.UInteger {
	return append([]protocol.UInteger(nil), data...)
}
