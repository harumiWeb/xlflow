package lspserver

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"sync"
	"sync/atomic"
	"testing"

	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/doccomments"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func TestDocumentSymbolCacheUsesVersionHashAndFullIdentity(t *testing.T) {
	cache := newDocumentSymbolCache()
	root := t.TempDir()
	path := filepath.Join(root, "first", "Main.bas")
	doc := intel.Document{
		URI:        pathToFileURI(path),
		Path:       path,
		Source:     "Option Explicit\n",
		ModuleKind: "standard",
		Version:    7,
	}
	var loads atomic.Int32
	load := func() ([]intel.Symbol, error) {
		loads.Add(1)
		return []intel.Symbol{{Name: doc.Source}}, nil
	}

	first, hit, err := cache.get(doc, load)
	if err != nil || hit || len(first) != 1 {
		t.Fatalf("first get = (%+v, hit=%v, err=%v), want one-symbol miss", first, hit, err)
	}
	second, hit, err := cache.get(doc, load)
	if err != nil || !hit || len(second) != 1 || loads.Load() != 1 {
		t.Fatalf("second get = (%+v, hit=%v, err=%v, loads=%d), want cached hit", second, hit, err, loads.Load())
	}

	doc.Source = "Option Explicit\nSub Changed()\nEnd Sub\n"
	changed, hit, err := cache.get(doc, load)
	if err != nil || hit || loads.Load() != 2 || changed[0].Name != doc.Source {
		t.Fatalf("same-version source change = (%+v, hit=%v, err=%v, loads=%d)", changed, hit, err, loads.Load())
	}
	doc.Version++
	if _, hit, err := cache.get(doc, load); err != nil || hit || loads.Load() != 3 {
		t.Fatalf("version change = (hit=%v, err=%v, loads=%d), want miss", hit, err, loads.Load())
	}

	other := doc
	other.Path = filepath.Join(root, "second", "Main.bas")
	other.URI = pathToFileURI(other.Path)
	if _, hit, err := cache.get(other, load); err != nil || hit || loads.Load() != 4 {
		t.Fatalf("same basename in another directory = (hit=%v, err=%v, loads=%d), want miss", hit, err, loads.Load())
	}

	failing := errors.New("parse failed")
	errorDoc := doc
	errorDoc.Source = "broken"
	errorLoad := func() ([]intel.Symbol, error) {
		loads.Add(1)
		return nil, failing
	}
	if _, _, err := cache.get(errorDoc, errorLoad); !errors.Is(err, failing) {
		t.Fatalf("first error = %v, want %v", err, failing)
	}
	if _, hit, err := cache.get(errorDoc, errorLoad); !errors.Is(err, failing) || hit {
		t.Fatalf("second error = (hit=%v, err=%v), want uncached error", hit, err)
	}
}

func TestDocumentSymbolCacheCoalescesConcurrentMissesAndRejectsInvalidatedResult(t *testing.T) {
	cache := newDocumentSymbolCache()
	doc := intel.Document{URI: "file:///C:/work/Main.bas", Path: `C:\work\Main.bas`, Source: "Option Explicit\n", Version: 1}
	started := make(chan struct{})
	release := make(chan struct{})
	var once sync.Once
	var loads atomic.Int32
	load := func() ([]intel.Symbol, error) {
		loads.Add(1)
		once.Do(func() { close(started) })
		<-release
		return []intel.Symbol{{Name: "Main"}}, nil
	}

	const requests = 12
	start := make(chan struct{})
	results := make(chan error, requests)
	for range requests {
		go func() {
			<-start
			syms, _, err := cache.get(doc, load)
			if err == nil && (len(syms) != 1 || syms[0].Name != "Main") {
				err = errors.New("unexpected symbols")
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
	if got := loads.Load(); got != 1 {
		t.Fatalf("concurrent loads = %d, want 1", got)
	}

	cache.invalidate(doc)
	if _, hit, err := cache.get(doc, load); err != nil || hit || loads.Load() != 2 {
		t.Fatalf("post-invalidation get = (hit=%v, err=%v, loads=%d), want fresh miss", hit, err, loads.Load())
	}

	inflightCache := newDocumentSymbolCache()
	inflightStarted := make(chan struct{})
	inflightRelease := make(chan struct{})
	var inflightLoads atomic.Int32
	inflightLoad := func() ([]intel.Symbol, error) {
		inflightLoads.Add(1)
		close(inflightStarted)
		<-inflightRelease
		return []intel.Symbol{{Name: "Old"}}, nil
	}
	inflightResult := make(chan error, 1)
	go func() {
		_, _, err := inflightCache.get(doc, inflightLoad)
		inflightResult <- err
	}()
	<-inflightStarted
	inflightCache.invalidate(doc)
	close(inflightRelease)
	if err := <-inflightResult; err != nil {
		t.Fatal(err)
	}
	freshLoad := func() ([]intel.Symbol, error) {
		inflightLoads.Add(1)
		return []intel.Symbol{{Name: "Fresh"}}, nil
	}
	got, hit, err := inflightCache.get(doc, freshLoad)
	if err != nil || hit || inflightLoads.Load() != 2 || got[0].Name != "Fresh" {
		t.Fatalf("after in-flight invalidation = (%+v, hit=%v, err=%v, loads=%d), want fresh miss", got, hit, err, inflightLoads.Load())
	}
}

func TestDocumentSymbolCacheReturnsDeepCopies(t *testing.T) {
	cache := newDocumentSymbolCache()
	doc := intel.Document{URI: "file:///C:/work/Main.bas", Path: `C:\work\Main.bas`, Source: "Option Explicit\n", Version: 1}
	load := func() ([]intel.Symbol, error) {
		return []intel.Symbol{{
			Name:       "Build",
			Parameters: []intel.Parameter{{Name: "value", Type: "Long"}},
			Documentation: doccomments.SymbolDocumentation{
				Parameters:       map[string]string{"value": "original"},
				ParameterEntries: []doccomments.ParameterDoc{{Name: "value", Description: "original"}},
				UnknownSections:  map[string]string{"custom": "original"},
			},
		}}, nil
	}

	got, _, err := cache.get(doc, load)
	if err != nil {
		t.Fatal(err)
	}
	got[0].Name = "Mutated"
	got[0].Parameters[0].Name = "mutated"
	got[0].Documentation.Parameters["value"] = "mutated"
	got[0].Documentation.ParameterEntries[0].Description = "mutated"
	got[0].Documentation.UnknownSections["custom"] = "mutated"

	got, hit, err := cache.get(doc, load)
	if err != nil || !hit {
		t.Fatalf("cached get = (hit=%v, err=%v)", hit, err)
	}
	if got[0].Name != "Build" || got[0].Parameters[0].Name != "value" ||
		got[0].Documentation.Parameters["value"] != "original" ||
		got[0].Documentation.ParameterEntries[0].Description != "original" ||
		got[0].Documentation.UnknownSections["custom"] != "original" {
		t.Fatalf("cached symbols were mutated: %+v", got[0])
	}
}

func TestDocumentSymbolCacheTracksDiskOverlayCloseAndReopen(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	writeSource := func(name string) {
		t.Helper()
		source := "Option Explicit\nPublic Sub " + name + "()\nEnd Sub\n"
		if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	writeSource("SavedOne")
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	s.diagnostics = func(context.Context, intel.Document) []intel.Diagnostic { return nil }
	ctx := &glsp.Context{Notify: func(string, any) {}}
	uri := pathToFileURI(path)

	assertDocumentSymbolName(t, s, uri, "SavedOne", true)
	writeSource("SavedTwo")
	assertDocumentSymbolName(t, s, uri, "SavedTwo", true)
	assertDocumentSymbolName(t, s, uri, "SavedOne", false)

	openSource := "Option Explicit\nPublic Sub UnsavedName()\nEnd Sub\n"
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 11, Text: openSource,
	}}); err != nil {
		t.Fatal(err)
	}
	assertDocumentSymbolName(t, s, uri, "UnsavedName", true)
	assertDocumentSymbolName(t, s, uri, "SavedTwo", false)

	changedSource := "Option Explicit\nPublic Sub EditedName()\nEnd Sub\n"
	if err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Version: 12,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: changedSource}},
	}); err != nil {
		t.Fatal(err)
	}
	assertDocumentSymbolName(t, s, uri, "EditedName", true)
	assertDocumentSymbolName(t, s, uri, "UnsavedName", false)

	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}); err != nil {
		t.Fatal(err)
	}
	assertDocumentSymbolName(t, s, uri, "SavedTwo", true)
	assertDocumentSymbolName(t, s, uri, "EditedName", false)

	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 11, Text: openSource,
	}}); err != nil {
		t.Fatal(err)
	}
	assertDocumentSymbolName(t, s, uri, "UnsavedName", true)
}

func TestDocumentSymbolCacheRefreshesSidecarUserFormControls(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	codePath := filepath.Join(codeDir, "CustomerForm.bas")
	code := "Option Explicit\nPrivate Sub UserForm_Initialize()\nEnd Sub\n"
	if err := os.WriteFile(codePath, []byte(code), 0o644); err != nil {
		t.Fatal(err)
	}
	formPath := filepath.Join(formsDir, "CustomerForm.frm")
	writeForm := func(control string) {
		t.Helper()
		form := "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm\n" +
			"   Begin MSForms.TextBox " + control + "\n   End\nEnd\nAttribute VB_Name = \"CustomerForm\"\n"
		if err := os.WriteFile(formPath, []byte(form), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	writeForm("txtOld")
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	uri := pathToFileURI(codePath)

	assertDocumentSymbolName(t, s, uri, "txtOld", true)
	writeForm("txtNew")
	assertDocumentSymbolName(t, s, uri, "txtNew", true)
	assertDocumentSymbolName(t, s, uri, "txtOld", false)
}

func assertDocumentSymbolName(t *testing.T, s *Server, uri, name string, want bool) {
	t.Helper()
	result, err := s.documentSymbol(nil, &protocol.DocumentSymbolParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	})
	if err != nil {
		t.Fatal(err)
	}
	symbols, ok := result.([]protocol.DocumentSymbol)
	if !ok {
		t.Fatalf("documentSymbol result type = %T", result)
	}
	found := false
	for _, symbol := range symbols {
		if symbol.Name == name {
			found = true
			break
		}
	}
	if found != want {
		t.Fatalf("symbol %q found = %v, want %v; symbols=%+v", name, found, want, symbols)
	}
}
