package lspserver

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestWorkspaceSymbolIndexParsesOnceAndUpdatesOnlyChangedFile(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	a := filepath.Join(moduleDir, "A.bas")
	b := filepath.Join(moduleDir, "B.bas")
	if err := os.WriteFile(a, []byte("Sub Alpha()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(b, []byte("Sub Beta()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	counts := map[string]int{}
	parse := func(file symbols.SourceFile, source []byte) (indexedFileSymbols, error) {
		counts[file.Path]++
		name := strings.Fields(string(source))[1]
		name = strings.TrimSuffix(name, "()")
		return indexedFileSymbols{path: file.Path, moduleKind: file.ModuleKind, symbols: []intel.Symbol{{
			Name: name, Kind: "sub", Module: strings.TrimSuffix(filepath.Base(file.Path), filepath.Ext(file.Path)), File: filepath.ToSlash(file.Path),
		}}}, nil
	}
	index := newWorkspaceSymbolIndex(root, config.Default(), parse, nil)
	if err := index.waitReady(); err != nil {
		t.Fatal(err)
	}
	if counts[a] != 1 || counts[b] != 1 {
		t.Fatalf("initial parses = %#v, want one per file", counts)
	}
	if got, err := index.searchExact("Alpha"); err != nil || len(got) != 1 {
		t.Fatalf("exact = %#v, %v", got, err)
	}
	if got, err := index.searchPrefix("Be"); err != nil || len(got) != 1 || got[0].Name != "Beta" {
		t.Fatalf("prefix = %#v, %v", got, err)
	}
	if got, err := index.searchQualified("A.Alpha"); err != nil || len(got) != 1 {
		t.Fatalf("qualified = %#v, %v", got, err)
	}
	if got, err := index.searchModule("B"); err != nil || len(got) != 1 {
		t.Fatalf("module = %#v, %v", got, err)
	}
	if got, err := index.searchKind("sub"); err != nil || len(got) != 2 {
		t.Fatalf("kind = %#v, %v", got, err)
	}
	if got, err := index.searchContains("alp"); err != nil || len(got) != 1 {
		t.Fatalf("contains = %#v, %v", got, err)
	}
	if counts[a] != 1 || counts[b] != 1 {
		t.Fatalf("warm queries reparsed files: %#v", counts)
	}

	if err := os.WriteFile(a, []byte("Sub Updated()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := index.updatePath(a); err != nil {
		t.Fatal(err)
	}
	if got, _ := index.searchExact("Alpha"); len(got) != 0 {
		t.Fatalf("old symbol remains: %#v", got)
	}
	if got, _ := index.searchExact("Updated"); len(got) != 1 {
		t.Fatalf("updated symbol missing: %#v", got)
	}
	if counts[a] != 2 || counts[b] != 1 {
		t.Fatalf("changed file parses = %#v", counts)
	}
	if err := index.updatePath(a); err != nil {
		t.Fatal(err)
	}
	if counts[a] != 2 || counts[b] != 1 {
		t.Fatalf("duplicate watcher event reparsed: %#v", counts)
	}
}

func TestWorkspaceSymbolIndexWatcherAndOpenOverlay(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Sub SavedName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if s.handler.WorkspaceDidChangeWatchedFiles == nil {
		t.Fatal("watcher handler was not registered")
	}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	uri := pathToFileURI(path)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{URI: protocol.DocumentUri(uri), Version: 1, Text: "Sub OpenName()\nEnd Sub\n"}}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.symbols.searchExact("OpenName"); len(got) != 1 {
		t.Fatalf("open overlay missing: %#v", got)
	}
	if err := os.WriteFile(path, []byte("Sub DiskChanged()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := s.didChangeWatchedFiles(ctx, &protocol.DidChangeWatchedFilesParams{Changes: []protocol.FileEvent{{URI: protocol.DocumentUri(uri), Type: protocol.FileChangeTypeChanged}}}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.symbols.searchExact("OpenName"); len(got) != 1 {
		t.Fatalf("disk event replaced open overlay: %#v", got)
	}
	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.symbols.searchExact("DiskChanged"); len(got) != 1 {
		t.Fatalf("disk source not restored after close: %#v", got)
	}
	if err := os.Remove(path); err != nil {
		t.Fatal(err)
	}
	if err := s.didChangeWatchedFiles(ctx, &protocol.DidChangeWatchedFilesParams{Changes: []protocol.FileEvent{{URI: protocol.DocumentUri(uri), Type: protocol.FileChangeTypeDeleted}}}); err != nil {
		t.Fatal(err)
	}
	if got, _ := s.symbols.searchContains("Name"); len(got) != 0 {
		t.Fatalf("deleted source remains: %#v", got)
	}
}

func TestWorkspaceSymbolIndexIncludesConfiguredRootsAndDuplicateBasenames(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Src.Modules = "vba/modules"
	cfg.Src.Classes = "vba/classes"
	cfg.Src.Forms = "vba/forms"
	cfg.Src.Workbook = "vba/workbook"
	files := map[string]string{
		"vba/modules/Shared.bas": "Sub ModuleOnly()\nEnd Sub\n",
		"vba/classes/Shared.cls": "Public Sub ClassOnly()\nEnd Sub\n",
		"vba/workbook/This.bas":  "Sub WorkbookOnly()\nEnd Sub\n",
		"vba/forms/Screen.frm":   "Sub FormOnly()\nEnd Sub\n",
		"outside/Ignored.bas":    "Sub Ignored()\nEnd Sub\n",
	}
	for relative, source := range files {
		path := filepath.Join(root, filepath.FromSlash(relative))
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	s, cleanup, err := New(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.symbols.waitReady(); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"ModuleOnly", "ClassOnly", "WorkbookOnly", "FormOnly"} {
		if got, _ := s.symbols.searchExact(name); len(got) != 1 {
			t.Fatalf("configured-root symbol %q = %#v", name, got)
		}
	}
	if got, _ := s.symbols.searchExact("Ignored"); len(got) != 0 {
		t.Fatalf("outside-root symbol indexed: %#v", got)
	}
}

func TestWorkspaceSymbolIndexExcludesExternalOpenDocumentAndUsesAbsoluteDiscoveryPaths(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	inside := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(inside, []byte("Sub Inside()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	cwd, err := os.Getwd()
	if err != nil {
		t.Fatal(err)
	}
	relativeRoot, err := filepath.Rel(cwd, root)
	if err != nil {
		t.Fatal(err)
	}
	files, err := symbols.DiscoverSourceFiles(symbols.Options{RootDir: relativeRoot, Config: config.Default()})
	if err != nil || len(files) != 1 || !filepath.IsAbs(files[0].Path) {
		t.Fatalf("relative-root discovery = %#v, %v", files, err)
	}

	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.symbols.waitReady(); err != nil {
		t.Fatal(err)
	}
	external := filepath.Join(t.TempDir(), "External.bas")
	if err := os.WriteFile(external, []byte("Sub ExternalOnly()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	ctx := &glsp.Context{Notify: func(string, any) {}}
	uri := pathToFileURI(external)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{URI: protocol.DocumentUri(uri), Version: 1, Text: "Sub ExternalOnly()\nEnd Sub\n"}}); err != nil {
		t.Fatal(err)
	}
	if got, err := s.symbols.searchExact("ExternalOnly"); err != nil || len(got) != 0 {
		t.Fatalf("external open document leaked into index: %#v, %v", got, err)
	}
}
