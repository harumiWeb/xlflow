package lspserver

import (
	"context"
	"errors"
	"path/filepath"
	"runtime"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func TestSemanticTokenCacheUsesVersionContentAndFullIdentity(t *testing.T) {
	cache := newSemanticTokenCache()
	root := t.TempDir()
	doc := intel.Document{
		URI:        pathToFileURI(filepath.Join(root, "first", "Main.bas")),
		Path:       filepath.Join(root, "first", "Main.bas"),
		Source:     "Option Explicit\n",
		ModuleKind: "standard",
		Version:    7,
	}
	var loads atomic.Int32
	load := func() ([]protocol.UInteger, error) {
		loads.Add(1)
		return []protocol.UInteger{0, 0, 6, 12, 0}, nil
	}

	first, hit, err := cache.get(doc, cache.begin(), load)
	if err != nil || hit || len(first.data) != 5 || first.resultID == "" {
		t.Fatalf("first get = (%v, hit=%v, err=%v), want miss", first, hit, err)
	}
	first.data[0] = 99
	second, hit, err := cache.get(doc, cache.begin(), load)
	if err != nil || !hit || loads.Load() != 1 || second.data[0] != 0 || second.resultID != first.resultID {
		t.Fatalf("second get = (%v, hit=%v, err=%v, loads=%d), want isolated hit", second, hit, err, loads.Load())
	}

	doc.Source = "Option Explicit\nSub Changed()\nEnd Sub\n"
	if _, hit, err := cache.get(doc, cache.begin(), load); err != nil || hit || loads.Load() != 2 {
		t.Fatalf("same-version source change = (hit=%v, err=%v, loads=%d), want miss", hit, err, loads.Load())
	}
	doc.Version++
	if _, hit, err := cache.get(doc, cache.begin(), load); err != nil || hit || loads.Load() != 3 {
		t.Fatalf("version change = (hit=%v, err=%v, loads=%d), want miss", hit, err, loads.Load())
	}

	other := doc
	other.Path = filepath.Join(root, "second", "Main.bas")
	other.URI = pathToFileURI(other.Path)
	if _, hit, err := cache.get(other, cache.begin(), load); err != nil || hit || loads.Load() != 4 {
		t.Fatalf("same basename other directory = (hit=%v, err=%v, loads=%d), want miss", hit, err, loads.Load())
	}

	failing := errors.New("semantic generation failed")
	errorDoc := doc
	errorDoc.Source = "broken"
	errorLoad := func() ([]protocol.UInteger, error) { return nil, failing }
	if _, _, err := cache.get(errorDoc, cache.begin(), errorLoad); !errors.Is(err, failing) {
		t.Fatalf("first error = %v, want %v", err, failing)
	}
	if _, hit, err := cache.get(errorDoc, cache.begin(), errorLoad); !errors.Is(err, failing) || hit {
		t.Fatalf("second error = (hit=%v, err=%v), want uncached error", hit, err)
	}
}

func TestSemanticTokenCacheCoalescesMissesAndRejectsInvalidatedResult(t *testing.T) {
	cache := newSemanticTokenCache()
	doc := intel.Document{URI: "file:///C:/work/Main.bas", Path: `C:\work\Main.bas`, Source: "Option Explicit\n", Version: 1}
	generation := cache.begin()
	started := make(chan struct{})
	release := make(chan struct{})
	var once sync.Once
	var loads atomic.Int32
	load := func() ([]protocol.UInteger, error) {
		loads.Add(1)
		once.Do(func() { close(started) })
		<-release
		return []protocol.UInteger{0, 0, 6, 12, 0}, nil
	}

	const requests = 12
	start := make(chan struct{})
	results := make(chan error, requests)
	for range requests {
		go func() {
			<-start
			result, _, err := cache.get(doc, generation, load)
			if err == nil && len(result.data) != 5 {
				err = errors.New("unexpected token data")
			}
			results <- err
		}()
	}
	close(start)
	<-started
	close(release)
	for range requests {
		if err := <-results; err != nil {
			t.Fatal(err)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("concurrent loads = %d, want 1", loads.Load())
	}

	staleCache := newSemanticTokenCache()
	staleGeneration := staleCache.begin()
	staleStarted := make(chan struct{})
	staleRelease := make(chan struct{})
	staleDone := make(chan error, 1)
	go func() {
		_, _, err := staleCache.get(doc, staleGeneration, func() ([]protocol.UInteger, error) {
			close(staleStarted)
			<-staleRelease
			return []protocol.UInteger{0, 0, 1, 12, 0}, nil
		})
		staleDone <- err
	}()
	<-staleStarted
	staleCache.invalidate(doc)
	newerDoc := doc
	newerDoc.Version++
	newerDoc.Source = "Option Explicit\nSub Newer()\nEnd Sub\n"
	newerGeneration := staleCache.begin()
	var newerLoads atomic.Int32
	newerDone := make(chan error, 1)
	go func() {
		_, _, err := staleCache.get(newerDoc, newerGeneration, func() ([]protocol.UInteger, error) {
			newerLoads.Add(1)
			return []protocol.UInteger{0, 0, 2, 12, 0}, nil
		})
		newerDone <- err
	}()
	identity := documentSymbolKey(doc)
	deadline := time.Now().Add(time.Second)
	for {
		staleCache.mu.Lock()
		call := staleCache.inflight[identity]
		waiting := call != nil && call.waiters > 0
		staleCache.mu.Unlock()
		if waiting {
			break
		}
		if time.Now().After(deadline) {
			t.Fatal("newer generation did not wait for the active obsolete generation")
		}
		runtime.Gosched()
	}
	if newerLoads.Load() != 0 {
		t.Fatalf("newer loads started while obsolete generation was active: %d", newerLoads.Load())
	}
	close(staleRelease)
	if err := <-staleDone; !errors.Is(err, errSemanticTokensSuperseded) {
		t.Fatalf("stale generation error = %v, want superseded", err)
	}
	if err := <-newerDone; !errors.Is(err, errSemanticTokensSuperseded) {
		t.Fatalf("waiting newer generation error = %v, want retry notification", err)
	}
	fresh, hit, err := staleCache.get(newerDoc, staleCache.begin(), func() ([]protocol.UInteger, error) {
		newerLoads.Add(1)
		return []protocol.UInteger{0, 0, 2, 12, 0}, nil
	})
	if err != nil || hit || fresh.data[2] != 2 || newerLoads.Load() != 1 {
		t.Fatalf("fresh get = (%v, hit=%v, err=%v, loads=%d), want one uncached latest load", fresh, hit, err, newerLoads.Load())
	}
}

func TestSemanticTokenCacheRetainsHistoryOnlyForOpenDocuments(t *testing.T) {
	cache := newSemanticTokenCache()
	doc := intel.Document{URI: "file:///C:/work/Main.bas", Path: `C:\work\Main.bas`, Source: "Option Explicit\n", Version: 1}
	load := func() ([]protocol.UInteger, error) { return []protocol.UInteger{0, 0, 6, 12, 0}, nil }
	identity := documentSymbolKey(doc)

	if _, _, err := cache.get(doc, cache.begin(), load); err != nil {
		t.Fatal(err)
	}
	cache.mu.Lock()
	unopenedHistory := len(cache.histories[identity])
	cache.mu.Unlock()
	if unopenedHistory != 0 {
		t.Fatalf("unopened history count = %d, want 0", unopenedHistory)
	}

	cache.open(doc)
	if _, _, err := cache.get(doc, cache.begin(), load); err != nil {
		t.Fatal(err)
	}
	cache.mu.Lock()
	openedHistory := len(cache.histories[identity])
	cache.mu.Unlock()
	if openedHistory != 1 {
		t.Fatalf("open history count = %d, want 1", openedHistory)
	}

	cache.close(doc)
	cache.mu.Lock()
	_, stillOpen := cache.openIdentities[identity]
	closedHistory := len(cache.histories[identity])
	cache.mu.Unlock()
	if stillOpen || closedHistory != 0 {
		t.Fatalf("closed cache state = (open=%v, history=%d), want (false, 0)", stillOpen, closedHistory)
	}
}

func TestSemanticTokensFullCachesAndInvalidatesOnDocumentLifecycle(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	s.diagnostics = func(context.Context, intel.Document) []intel.Diagnostic { return nil }
	var generations atomic.Int32
	s.semanticTokenGenerator = func(doc intel.Document, open []intel.Document) ([]intel.SemanticToken, error) {
		generations.Add(1)
		return []intel.SemanticToken{{
			Range: intel.Range{Start: intel.Position{}, End: intel.Position{Character: len(open)}},
			Type:  intel.SemanticTokenKeyword,
		}}, nil
	}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 1, Text: "A",
	}}); err != nil {
		t.Fatal(err)
	}
	params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}
	first, err := s.semanticTokensFull(nil, params)
	if err != nil {
		t.Fatal(err)
	}
	first.Data[2] = 99
	second, err := s.semanticTokensFull(nil, params)
	if err != nil || generations.Load() != 1 || second.Data[2] != 1 {
		t.Fatalf("cached result = (%v, err=%v, generations=%d), want isolated hit", second.Data, err, generations.Load())
	}

	if err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Version: 2,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: "AB"}},
	}); err != nil {
		t.Fatal(err)
	}
	changed, err := s.semanticTokensFull(nil, params)
	if err != nil || generations.Load() != 2 || changed.Data[2] != 1 {
		t.Fatalf("changed result = (%v, err=%v, generations=%d), want regenerated data", changed.Data, err, generations.Load())
	}

	otherPath := filepath.Join(root, "src", "classes", "Other.cls")
	otherURI := pathToFileURI(otherPath)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(otherURI), LanguageID: "vba", Version: 1, Text: "Class Other",
	}}); err != nil {
		t.Fatal(err)
	}
	otherChanged, err := s.semanticTokensFull(nil, params)
	if err != nil || generations.Load() != 3 || otherChanged.Data[2] != 2 {
		t.Fatalf("cross-document invalidation = (%v, err=%v, generations=%d), want regenerated workspace-dependent tokens", otherChanged, err, generations.Load())
	}

	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}); err != nil {
		t.Fatal(err)
	}
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 2, Text: "AB",
	}}); err != nil {
		t.Fatal(err)
	}
	if _, err := s.semanticTokensFull(nil, params); err != nil || generations.Load() != 4 {
		t.Fatalf("close/reopen invalidation = (err=%v, generations=%d), want regeneration", err, generations.Load())
	}
}

func TestSemanticTokensFullSerializesObsoleteGenerationAndRetriesLatest(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	s.diagnostics = func(context.Context, intel.Document) []intel.Diagnostic { return nil }
	started := make(chan struct{})
	release := make(chan struct{})
	var generations atomic.Int32
	var active atomic.Int32
	var maximum atomic.Int32
	s.semanticTokenGenerator = func(doc intel.Document, _ []intel.Document) ([]intel.SemanticToken, error) {
		generation := generations.Add(1)
		current := active.Add(1)
		for {
			observed := maximum.Load()
			if current <= observed || maximum.CompareAndSwap(observed, current) {
				break
			}
		}
		if generation == 1 {
			close(started)
			<-release
		}
		active.Add(-1)
		return []intel.SemanticToken{{
			Range: intel.Range{Start: intel.Position{}, End: intel.Position{Character: len(doc.Source)}},
			Type:  intel.SemanticTokenKeyword,
		}}, nil
	}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 1, Text: "A",
	}}); err != nil {
		t.Fatal(err)
	}
	params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}
	type result struct {
		tokens *protocol.SemanticTokens
		err    error
	}
	results := make(chan result, 2)
	go func() {
		tokens, err := s.semanticTokensFull(nil, params)
		results <- result{tokens: tokens, err: err}
	}()
	<-started
	if err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Version: 2,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: "AB"}},
	}); err != nil {
		t.Fatal(err)
	}
	go func() {
		tokens, err := s.semanticTokensFull(nil, params)
		results <- result{tokens: tokens, err: err}
	}()
	identity := documentSymbolKey(intel.Document{Path: path, URI: uri})
	deadline := time.Now().Add(time.Second)
	for {
		s.semanticTokens.mu.Lock()
		call := s.semanticTokens.inflight[identity]
		waiting := call != nil && call.waiters > 0
		s.semanticTokens.mu.Unlock()
		if waiting {
			break
		}
		if time.Now().After(deadline) {
			t.Fatal("latest semantic request did not wait for obsolete generation")
		}
		runtime.Gosched()
	}
	close(release)
	for range 2 {
		select {
		case got := <-results:
			if got.err != nil {
				t.Fatal(got.err)
			}
			if got.tokens == nil || len(got.tokens.Data) != 5 || got.tokens.Data[2] != 2 {
				t.Fatalf("retried semantic tokens = %+v, want latest version length 2", got.tokens)
			}
		case <-time.After(2 * time.Second):
			t.Fatal("semantic token retry did not complete")
		}
	}
	if generations.Load() != 2 || maximum.Load() != 1 {
		t.Fatalf("generation stats = calls:%d max_active:%d, want 2 calls and one active", generations.Load(), maximum.Load())
	}
	s.semanticTokens.mu.Lock()
	history := append([]cachedSemanticTokens(nil), s.semanticTokens.histories[identity]...)
	s.semanticTokens.mu.Unlock()
	if len(history) != 1 || history[0].signature.version != 2 {
		t.Fatalf("cached history after stale generation = %+v, want only version 2", history)
	}
}
