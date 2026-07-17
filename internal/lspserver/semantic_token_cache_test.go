package lspserver

import (
	"context"
	"errors"
	"path/filepath"
	"sync"
	"sync/atomic"
	"testing"

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
	if err != nil || hit || len(first) != 5 {
		t.Fatalf("first get = (%v, hit=%v, err=%v), want miss", first, hit, err)
	}
	first[0] = 99
	second, hit, err := cache.get(doc, cache.begin(), load)
	if err != nil || !hit || loads.Load() != 1 || second[0] != 0 {
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
			data, _, err := cache.get(doc, generation, load)
			if err == nil && len(data) != 5 {
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
	staleCache.invalidateAll()
	close(staleRelease)
	if err := <-staleDone; err != nil {
		t.Fatal(err)
	}
	fresh, hit, err := staleCache.get(doc, staleCache.begin(), func() ([]protocol.UInteger, error) {
		return []protocol.UInteger{0, 0, 2, 12, 0}, nil
	})
	if err != nil || hit || fresh[2] != 2 {
		t.Fatalf("fresh get = (%v, hit=%v, err=%v), want uncached fresh data", fresh, hit, err)
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
	s.semanticTokenGenerator = func(doc intel.Document, _ []intel.Document) ([]intel.SemanticToken, error) {
		generations.Add(1)
		length := 1
		if len(doc.Source) > 1 {
			length = 2
		}
		return []intel.SemanticToken{{
			Range: intel.Range{Start: intel.Position{}, End: intel.Position{Character: length}},
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
	if err != nil || generations.Load() != 2 || changed.Data[2] != 2 {
		t.Fatalf("changed result = (%v, err=%v, generations=%d), want regenerated data", changed.Data, err, generations.Load())
	}

	otherPath := filepath.Join(root, "src", "classes", "Other.cls")
	otherURI := pathToFileURI(otherPath)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(otherURI), LanguageID: "vba", Version: 1, Text: "Class Other",
	}}); err != nil {
		t.Fatal(err)
	}
	if _, err := s.semanticTokensFull(nil, params); err != nil || generations.Load() != 3 {
		t.Fatalf("cross-document invalidation = (err=%v, generations=%d), want regeneration", err, generations.Load())
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
