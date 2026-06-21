package lspserver

import (
	"context"
	"net"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestCheckLoadsBuiltinDatabase(t *testing.T) {
	if err := Check(Options{RootDir: t.TempDir(), Config: config.Default()}); err != nil {
		t.Fatal(err)
	}
}

func TestFileURIPathRoundTripWithEscapedJapanesePath(t *testing.T) {
	path := filepath.Join(t.TempDir(), "日本 語#%dir", "Main.bas")
	uri := pathToFileURI(path)
	if !strings.HasPrefix(uri, "file:") {
		t.Fatalf("uri = %q, want file URI", uri)
	}

	got, err := fileURIToPath(uri)
	if err != nil {
		t.Fatal(err)
	}
	if normalizePathKey(got) != normalizePathKey(path) {
		t.Fatalf("roundtrip path = %q, want %q via %q", got, path, uri)
	}
}

func TestFileURIPathRoundTripWithWindowsUNCPath(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows UNC path handling is OS-specific")
	}
	path := `\\server\share\日本 語#%dir\Main.bas`
	uri := pathToFileURI(path)
	if !strings.HasPrefix(uri, "file://server/share/") {
		t.Fatalf("uri = %q, want UNC file URI with host and share", uri)
	}
	if strings.Contains(uri, "#") || strings.Contains(uri, "%dir") {
		t.Fatalf("uri = %q, want escaped special characters", uri)
	}

	got, err := fileURIToPath(uri)
	if err != nil {
		t.Fatal(err)
	}
	if normalizePathKey(got) != normalizePathKey(path) {
		t.Fatalf("roundtrip path = %q, want %q via %q", got, path, uri)
	}
}

func TestDocumentsOverlayNormalizesWindowsDriveLetterCase(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows drive-letter normalization is OS-specific")
	}
	root := `C:\work`
	docs := newDocuments(root)
	upper := "file:///C:/work/src/modules/Main.bas"
	lower := "file:///c:/work/src/modules/Main.bas"

	if _, err := docs.open(upper, "Option Explicit\nSub UpperName()\nEnd Sub\n"); err != nil {
		t.Fatal(err)
	}
	changed, err := docs.change(lower, "Option Explicit\nSub LowerName()\nEnd Sub\n")
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(changed.Source, "LowerName") {
		t.Fatalf("changed source was not stored through normalized key: %q", changed.Source)
	}
	if len(docs.openDocuments()) != 1 {
		t.Fatalf("open documents = %d, want 1", len(docs.openDocuments()))
	}
}

func TestDocumentsOverlayUsesUnsavedChangesAndClearsOnClose(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	docs := newDocuments(root)

	opened, err := docs.open(uri, "Option Explicit\nSub OldName()\nEnd Sub\n")
	if err != nil {
		t.Fatal(err)
	}
	if opened.ModuleKind != "standard" {
		t.Fatalf("module kind = %q, want standard", opened.ModuleKind)
	}

	changed, err := docs.change(uri, "Option Explicit\nSub NewName()\nEnd Sub\n")
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(changed.Source, "NewName") {
		t.Fatalf("changed source was not stored: %q", changed.Source)
	}
	if len(docs.openDocuments()) != 1 {
		t.Fatalf("open documents = %d, want 1", len(docs.openDocuments()))
	}

	docs.close(uri)
	if len(docs.openDocuments()) != 0 {
		t.Fatalf("open documents = %d, want 0 after close", len(docs.openDocuments()))
	}
}

func TestWorkspaceSymbolCacheUsesOpenDocumentOverlay(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Option Explicit\nSub SavedName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	uri := pathToFileURI(path)
	if _, err := s.docs.open(uri, "Option Explicit\nSub OpenName()\nEnd Sub\n"); err != nil {
		t.Fatal(err)
	}
	openSymbols, err := s.workspaceSymbol(nil, &protocol.WorkspaceSymbolParams{Query: "OpenName"})
	if err != nil {
		t.Fatal(err)
	}
	if !hasWorkspaceSymbol(openSymbols, "OpenName") {
		t.Fatalf("open document symbol missing: %+v", openSymbols)
	}
	savedSymbols, err := s.workspaceSymbol(nil, &protocol.WorkspaceSymbolParams{Query: "SavedName"})
	if err != nil {
		t.Fatal(err)
	}
	if hasWorkspaceSymbol(savedSymbols, "SavedName") {
		t.Fatalf("saved symbol should be hidden by open document overlay: %+v", savedSymbols)
	}

	if _, err := s.docs.change(uri, "Option Explicit\nSub ChangedName()\nEnd Sub\n"); err != nil {
		t.Fatal(err)
	}
	changedSymbols, err := s.workspaceSymbol(nil, &protocol.WorkspaceSymbolParams{Query: "ChangedName"})
	if err != nil {
		t.Fatal(err)
	}
	if !hasWorkspaceSymbol(changedSymbols, "ChangedName") {
		t.Fatalf("changed open document symbol missing: %+v", changedSymbols)
	}
	oldOpenSymbols, err := s.workspaceSymbol(nil, &protocol.WorkspaceSymbolParams{Query: "OpenName"})
	if err != nil {
		t.Fatal(err)
	}
	if hasWorkspaceSymbol(oldOpenSymbols, "OpenName") {
		t.Fatalf("stale open document symbol should be invalidated by source change: %+v", oldOpenSymbols)
	}
}

func TestNewCreatesLogFileWithoutStartingServer(t *testing.T) {
	root := t.TempDir()
	logPath := filepath.Join(".xlflow", "lsp.log")
	_, cleanup, err := New(Options{RootDir: root, Config: config.Default(), LogFile: logPath})
	if err != nil {
		t.Fatal(err)
	}
	cleanup()

	if _, err := os.Stat(filepath.Join(root, logPath)); err != nil {
		t.Fatal(err)
	}
}

func TestCompletionReturnsMemberCandidatesFromOpenDocument(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\nSub Test()\n    Worksheets(\"Input\").Ra\nEnd Sub\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}

	result, err := s.completion(nil, &protocol.CompletionParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 2, Character: 27},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	list, ok := result.(protocol.CompletionList)
	if !ok {
		t.Fatalf("completion result = %T, want CompletionList", result)
	}
	if !hasCompletionItem(list.Items, "Range") {
		t.Fatalf("Range completion missing: %+v", list.Items)
	}
}

func TestCompletionReturnsModuleProcedureCandidates(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Utils.bas"), []byte(`Attribute VB_Name = "Utils"
Option Explicit
Public Function BuildName() As String
End Function
`), 0o644); err != nil {
		t.Fatal(err)
	}

	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(moduleDir, "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\nSub Test()\n    Utils.Bu\nEnd Sub\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}

	result, err := s.completion(nil, &protocol.CompletionParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 2, Character: 12},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	list, ok := result.(protocol.CompletionList)
	if !ok {
		t.Fatalf("completion result = %T, want CompletionList", result)
	}
	if !hasCompletionItem(list.Items, "BuildName") {
		t.Fatalf("Utils.BuildName completion missing: %+v", list.Items)
	}
}

func TestCompletionReturnsModuleDeclarationSnippet(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\n\nPu\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}

	result, err := s.completion(nil, &protocol.CompletionParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 2, Character: 2},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	list, ok := result.(protocol.CompletionList)
	if !ok {
		t.Fatalf("completion result = %T, want CompletionList", result)
	}
	item, ok := findCompletionItem(list.Items, "Public Sub")
	if !ok {
		t.Fatalf("Public Sub completion missing: %+v", list.Items)
	}
	if item.InsertText == nil || !strings.Contains(*item.InsertText, "End Sub") {
		t.Fatalf("Public Sub insert text = %+v, want snippet with End Sub", item.InsertText)
	}
	if item.InsertTextFormat == nil || *item.InsertTextFormat != protocol.InsertTextFormatSnippet {
		t.Fatalf("Public Sub insert text format = %+v, want snippet", item.InsertTextFormat)
	}

	if _, err := s.docs.change(uri, "Option Explicit\n\nPublic S\n"); err != nil {
		t.Fatal(err)
	}
	result, err = s.completion(nil, &protocol.CompletionParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 2, Character: 8},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	list, ok = result.(protocol.CompletionList)
	if !ok {
		t.Fatalf("completion result = %T, want CompletionList", result)
	}
	item, ok = findCompletionItem(list.Items, "Public Sub")
	if !ok {
		t.Fatalf("Public Sub completion missing for multi-word prefix: %+v", list.Items)
	}
	edit, ok := item.TextEdit.(protocol.TextEdit)
	if !ok {
		t.Fatalf("Public Sub text edit = %T, want TextEdit", item.TextEdit)
	}
	if edit.Range.Start.Character != 0 || edit.Range.End.Character != 8 || !strings.Contains(edit.NewText, "End Sub") {
		t.Fatalf("Public Sub text edit = %+v, want replacement for typed prefix", edit)
	}
}

func TestCompletionReturnsTypeCandidates(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\nSub Test()\n    Dim ws As Wo\nEnd Sub\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}

	result, err := s.completion(nil, &protocol.CompletionParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 2, Character: 16},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	list, ok := result.(protocol.CompletionList)
	if !ok {
		t.Fatalf("completion result = %T, want CompletionList", result)
	}
	item, ok := findCompletionItem(list.Items, "Workbook")
	if !ok {
		t.Fatalf("Workbook type completion missing: %+v", list.Items)
	}
	if item.Kind == nil || *item.Kind != protocol.CompletionItemKindClass {
		t.Fatalf("Workbook kind = %+v, want class/type completion", item.Kind)
	}
	edit, ok := item.TextEdit.(protocol.TextEdit)
	if !ok {
		t.Fatalf("Workbook text edit = %T, want TextEdit", item.TextEdit)
	}
	if edit.Range.Start.Character != 14 || edit.Range.End.Character != 16 || edit.NewText != "Workbook" {
		t.Fatalf("Workbook text edit = %+v, want replacement for typed type prefix", edit)
	}
}

func TestJSONRPCIntegrationInitializeOpenCompletionAndExit(t *testing.T) {
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	serverSide, clientSide := net.Pipe()
	serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{handler: &s.handler})
	recorder := &rpcRecorder{}
	clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), recorder)
	defer func() { _ = clientConn.Close() }()

	var initResult protocol.InitializeResult
	if err := clientConn.Call(ctx, string(protocol.MethodInitialize), protocol.InitializeParams{}, &initResult); err != nil {
		t.Fatal(err)
	}
	if initResult.ServerInfo == nil || initResult.ServerInfo.Name != serverName {
		t.Fatalf("unexpected initialize result: %+v", initResult.ServerInfo)
	}
	if initResult.Capabilities.CompletionProvider == nil ||
		!containsString(initResult.Capabilities.CompletionProvider.TriggerCharacters, ".") ||
		!containsString(initResult.Capabilities.CompletionProvider.TriggerCharacters, "\"") ||
		containsString(initResult.Capabilities.CompletionProvider.TriggerCharacters, "P") {
		t.Fatalf("completion trigger characters = %+v, want member and string-literal LSP completions only", initResult.Capabilities.CompletionProvider)
	}
	if initResult.Capabilities.DocumentFormattingProvider == nil {
		t.Fatalf("documentFormattingProvider was not advertised: %+v", initResult.Capabilities)
	}

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	if err := clientConn.Notify(ctx, string(protocol.MethodTextDocumentDidOpen), protocol.DidOpenTextDocumentParams{
		TextDocument: protocol.TextDocumentItem{
			URI:        protocol.DocumentUri(uri),
			LanguageID: "vba",
			Version:    1,
			Text:       "Option Explicit\nSub Test()\n    Worksheets(\"Input\").Ra\nEnd Sub\n",
		},
	}); err != nil {
		t.Fatal(err)
	}

	var list protocol.CompletionList
	if err := clientConn.Call(ctx, string(protocol.MethodTextDocumentCompletion), protocol.CompletionParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 2, Character: 27},
		},
	}, &list); err != nil {
		t.Fatal(err)
	}
	if !hasCompletionItem(list.Items, "Range") {
		t.Fatalf("Range completion missing: %+v", list.Items)
	}

	if err := clientConn.Notify(ctx, string(protocol.MethodTextDocumentDidChange), protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Version:                2,
		},
		ContentChanges: []any{
			map[string]any{"text": "Option Explicit\nSub Test()\n    Range(\"A1\").Font.Co\nEnd Sub\n"},
		},
	}); err != nil {
		t.Fatal(err)
	}
	var fontList protocol.CompletionList
	if err := clientConn.Call(ctx, string(protocol.MethodTextDocumentCompletion), protocol.CompletionParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 2, Character: 23},
		},
	}, &fontList); err != nil {
		t.Fatal(err)
	}
	if !hasCompletionItem(fontList.Items, "Color") {
		t.Fatalf("Font.Color completion missing: %+v", fontList.Items)
	}

	var shutdown any
	if err := clientConn.Call(ctx, string(protocol.MethodShutdown), nil, &shutdown); err != nil {
		t.Fatal(err)
	}
	if err := clientConn.Notify(ctx, string(protocol.MethodExit), nil); err != nil {
		t.Fatal(err)
	}
	_ = serverConn.Close()
	if !recorder.seen(string(protocol.ServerTextDocumentPublishDiagnostics)) {
		t.Fatalf("expected publishDiagnostics notification, got %v", recorder.methods())
	}
}

func TestFormattingReturnsFullDocumentEditFromOpenDocument(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\nSub Test()\nIf True Then\nDebug.Print \"ok\"\nEnd If\nEnd Sub\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}

	edits, err := s.formatting(nil, &protocol.DocumentFormattingParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
		Options:      protocol.FormattingOptions{"tabSize": 4, "insertSpaces": true},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(edits) != 1 {
		t.Fatalf("formatting edits = %+v, want one full-document edit", edits)
	}
	if edits[0].Range.Start.Line != 0 || edits[0].Range.Start.Character != 0 {
		t.Fatalf("formatting edit start = %+v, want document start", edits[0].Range.Start)
	}
	if edits[0].Range.End.Line != 6 || edits[0].Range.End.Character != 0 {
		t.Fatalf("formatting edit end = %+v, want document end", edits[0].Range.End)
	}
	want := "Option Explicit\n\nSub Test()\n    If True Then\n        Debug.Print \"ok\"\n    End If\nEnd Sub\n"
	if edits[0].NewText != want {
		t.Fatalf("formatted text:\n%q\nwant:\n%q", edits[0].NewText, want)
	}
}

func TestFormattingSkipsFrmDocuments(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "forms", "UserForm1.frm")
	uri := pathToFileURI(path)
	if _, err := s.docs.open(uri, "VERSION 5.00\nBegin VB.Form UserForm1\nEnd\n"); err != nil {
		t.Fatal(err)
	}
	edits, err := s.formatting(nil, &protocol.DocumentFormattingParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
		Options:      protocol.FormattingOptions{"tabSize": 4, "insertSpaces": true},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(edits) != 0 {
		t.Fatalf("frm formatting edits = %+v, want none", edits)
	}
}

type rpcRecorder struct {
	mu          sync.Mutex
	methodsSeen []string
}

func (r *rpcRecorder) Handle(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	r.mu.Lock()
	r.methodsSeen = append(r.methodsSeen, req.Method)
	r.mu.Unlock()
	if !req.Notif {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeMethodNotFound, Message: "method not found"})
	}
}

func (r *rpcRecorder) seen(method string) bool {
	r.mu.Lock()
	defer r.mu.Unlock()
	for _, candidate := range r.methodsSeen {
		if candidate == method {
			return true
		}
	}
	return false
}

func (r *rpcRecorder) methods() []string {
	r.mu.Lock()
	defer r.mu.Unlock()
	return append([]string(nil), r.methodsSeen...)
}

func hasCompletionItem(items []protocol.CompletionItem, label string) bool {
	_, ok := findCompletionItem(items, label)
	return ok
}

func findCompletionItem(items []protocol.CompletionItem, label string) (protocol.CompletionItem, bool) {
	for _, item := range items {
		if item.Label == label {
			return item, true
		}
	}
	return protocol.CompletionItem{}, false
}

func hasWorkspaceSymbol(items []protocol.SymbolInformation, name string) bool {
	for _, item := range items {
		if item.Name == name {
			return true
		}
	}
	return false
}

func containsString(items []string, needle string) bool {
	for _, item := range items {
		if item == needle {
			return true
		}
	}
	return false
}
