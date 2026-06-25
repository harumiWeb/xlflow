package lspserver

import (
	"context"
	"encoding/json"
	"net"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
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
	if initResult.Capabilities.SignatureHelpProvider == nil ||
		!containsString(initResult.Capabilities.SignatureHelpProvider.TriggerCharacters, "(") ||
		!containsString(initResult.Capabilities.SignatureHelpProvider.TriggerCharacters, ",") ||
		!containsString(initResult.Capabilities.SignatureHelpProvider.TriggerCharacters, " ") {
		t.Fatalf("signatureHelpProvider = %+v, want trigger characters", initResult.Capabilities.SignatureHelpProvider)
	}
	if initResult.Capabilities.CodeLensProvider == nil ||
		initResult.Capabilities.CodeLensProvider.ResolveProvider == nil ||
		*initResult.Capabilities.CodeLensProvider.ResolveProvider {
		t.Fatalf("codeLensProvider = %+v, want resolveProvider=false", initResult.Capabilities.CodeLensProvider)
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

	if err := clientConn.Notify(ctx, string(protocol.MethodTextDocumentDidChange), protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Version:                3,
		},
		ContentChanges: []any{
			map[string]any{"text": "Option Explicit\nPublic Sub UnsavedRun()\nEnd Sub\nPublic Sub Test_UnsavedRun()\nEnd Sub\nPublic Sub WithArg(value As String)\nEnd Sub\n"},
		},
	}); err != nil {
		t.Fatal(err)
	}
	var lenses []protocol.CodeLens
	if err := clientConn.Call(ctx, string(protocol.MethodTextDocumentCodeLens), protocol.CodeLensParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	}, &lenses); err != nil {
		t.Fatal(err)
	}
	if len(lenses) != 2 {
		t.Fatalf("code lenses = %+v, want Run and Run Test", lenses)
	}
	if lenses[0].Command == nil || lenses[0].Command.Command != "xlflow.runProcedure" || lenses[0].Command.Title != "$(play) Run" {
		t.Fatalf("first code lens = %+v, want run procedure", lenses[0])
	}
	if lenses[1].Command == nil || lenses[1].Command.Command != "xlflow.runTestProcedure" || lenses[1].Command.Title != "$(beaker) Run Test" {
		t.Fatalf("second code lens = %+v, want run test procedure", lenses[1])
	}
	if lenses[0].Range.Start.Line != 1 || lenses[1].Range.Start.Line != 3 {
		t.Fatalf("code lens ranges = %+v, %+v", lenses[0].Range, lenses[1].Range)
	}
	args, ok := lenses[0].Command.Arguments[0].(map[string]any)
	if !ok || args["name"] != "UnsavedRun" || args["moduleName"] != "Main" || args["qualifiedName"] != "Main.UnsavedRun" || args["kind"] != "sub" {
		t.Fatalf("code lens args = %#v", lenses[0].Command.Arguments)
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

func TestSignatureHelpReturnsActiveParameter(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\nSub Test()\n    Dim ws As Worksheet\n    ws.Range(\"A1\",\nEnd Sub\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}
	line := `    ws.Range("A1",`
	help, err := s.signatureHelp(nil, &protocol.SignatureHelpParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 3, Character: protocol.UInteger(utf16Len(line))},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if help.ActiveParameter == nil || *help.ActiveParameter != 1 {
		t.Fatalf("active parameter = %+v, want 1", help.ActiveParameter)
	}
	if help.Signatures[0].Label != "Excel.Worksheet.Range(Cell1 As Variant, Optional Cell2 As Variant) As Excel.Range" {
		t.Fatalf("signature label = %q", help.Signatures[0].Label)
	}
	if len(help.Signatures[0].Parameters) != 2 {
		t.Fatalf("parameters = %+v", help.Signatures[0].Parameters)
	}
}

func TestSignatureHelpReturnsParenlessCallAfterSpace(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\nSub Test()\n    Dim dict As Scripting.Dictionary\n    dict.Add \nEnd Sub\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}
	line := "    dict.Add "
	help, err := s.signatureHelp(nil, &protocol.SignatureHelpParams{
		TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Position:     protocol.Position{Line: 3, Character: protocol.UInteger(utf16Len(line))},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if help == nil || len(help.Signatures) != 1 {
		t.Fatalf("signature help = %+v, want one signature", help)
	}
	if help.ActiveParameter == nil || *help.ActiveParameter != 0 {
		t.Fatalf("active parameter = %+v, want 0", help.ActiveParameter)
	}
	if help.Signatures[0].Label != "Scripting.Dictionary.Add(Key As Variant, Item As Variant) As void" {
		t.Fatalf("signature label = %q", help.Signatures[0].Label)
	}
}

func TestJSONRPCPublishesArgumentDiagnostics(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	serverSide, clientSide := net.Pipe()
	serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{handler: &s.handler})
	recorder := &rpcRecorder{}
	clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), recorder)
	defer func() { _ = clientConn.Close() }()

	var initResult protocol.InitializeResult
	if err := clientConn.Call(ctx, string(protocol.MethodInitialize), protocol.InitializeParams{}, &initResult); err != nil {
		t.Fatal(err)
	}

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	if err := clientConn.Notify(ctx, string(protocol.MethodTextDocumentDidOpen), protocol.DidOpenTextDocumentParams{
		TextDocument: protocol.TextDocumentItem{
			URI:        protocol.DocumentUri(uri),
			LanguageID: "vba",
			Version:    1,
			Text:       "Option Explicit\nSub Test()\n    Dim dict As Scripting.Dictionary\n    dict.Add \"A\"\nEnd Sub\n",
		},
	}); err != nil {
		t.Fatal(err)
	}

	deadline := time.Now().Add(time.Second)
	for time.Now().Before(deadline) {
		for _, params := range recorder.publishDiagnostics() {
			for _, diag := range params.Diagnostics {
				if strings.Contains(diag.Message, "Argument count mismatch") {
					_ = serverConn.Close()
					return
				}
			}
		}
		time.Sleep(10 * time.Millisecond)
	}
	_ = serverConn.Close()
	t.Fatalf("VB030 publishDiagnostics missing: %+v", recorder.publishDiagnostics())
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

func TestInitializeAdvertisesSemanticTokensProvider(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	result, err := s.initialize(nil, &protocol.InitializeParams{})
	if err != nil {
		t.Fatal(err)
	}
	init, ok := result.(protocol.InitializeResult)
	if !ok {
		t.Fatalf("initialize result = %T, want InitializeResult", result)
	}
	provider, ok := init.Capabilities.SemanticTokensProvider.(*protocol.SemanticTokensOptions)
	if !ok {
		t.Fatalf("semanticTokensProvider = %T, want *SemanticTokensOptions", init.Capabilities.SemanticTokensProvider)
	}
	if provider.Full != true {
		t.Fatalf("semantic full provider = %+v, want true", provider.Full)
	}
	if provider.Range != nil {
		t.Fatalf("semantic range provider = %+v, want nil", provider.Range)
	}
	if !containsString(provider.Legend.TokenTypes, "function") || !containsString(provider.Legend.TokenTypes, "property") {
		t.Fatalf("semantic token legend missing expected types: %+v", provider.Legend.TokenTypes)
	}
	if !containsString(provider.Legend.TokenModifiers, "defaultLibrary") {
		t.Fatalf("semantic token modifiers missing defaultLibrary: %+v", provider.Legend.TokenModifiers)
	}
}

func TestInitializeAdvertisesCodeLensProviderAndParsesConfig(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	result, err := s.initialize(nil, &protocol.InitializeParams{
		InitializationOptions: map[string]any{
			"codeLens": map[string]any{
				"enabled":        true,
				"runProcedure":   false,
				"runTests":       true,
				"userFormEvents": true,
			},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	init, ok := result.(protocol.InitializeResult)
	if !ok {
		t.Fatalf("initialize result = %T, want InitializeResult", result)
	}
	if init.Capabilities.CodeLensProvider == nil || init.Capabilities.CodeLensProvider.ResolveProvider == nil || *init.Capabilities.CodeLensProvider.ResolveProvider {
		t.Fatalf("codeLensProvider = %+v, want resolveProvider=false", init.Capabilities.CodeLensProvider)
	}
	if s.codeLensConfig.RunProcedure || !s.codeLensConfig.RunTests || !s.codeLensConfig.UserFormEvents {
		t.Fatalf("codeLensConfig = %+v", s.codeLensConfig)
	}
}

func TestCodeLensHonorsConfiguration(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	if _, err := s.docs.open(uri, "Option Explicit\nPublic Sub RunReport()\nEnd Sub\nPublic Sub Test_RunReport()\nEnd Sub\n"); err != nil {
		t.Fatal(err)
	}
	s.codeLensConfig = intel.DefaultCodeLensConfig()
	s.codeLensConfig.RunProcedure = false
	lenses, err := s.codeLens(nil, &protocol.CodeLensParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(lenses) != 1 || lenses[0].Command == nil || lenses[0].Command.Command != "xlflow.runTestProcedure" {
		t.Fatalf("lenses with runProcedure=false = %+v", lenses)
	}

	s.codeLensConfig = intel.DefaultCodeLensConfig()
	s.codeLensConfig.RunTests = false
	lenses, err = s.codeLens(nil, &protocol.CodeLensParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(lenses) != 1 || lenses[0].Command == nil || lenses[0].Command.Command != "xlflow.runProcedure" {
		t.Fatalf("lenses with runTests=false = %+v", lenses)
	}

	s.codeLensConfig = intel.DefaultCodeLensConfig()
	s.codeLensConfig.Enabled = false
	lenses, err = s.codeLens(nil, &protocol.CodeLensParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(lenses) != 0 {
		t.Fatalf("lenses with enabled=false = %+v", lenses)
	}
}

func TestSemanticTokensFullUsesOpenDocumentAndValidEncoding(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "modules", "Main.bas")
	uri := pathToFileURI(path)
	source := "Option Explicit\nSub OpenName()\n    Range(\"A1\").Value = xlLandscape\nEnd Sub\n"
	if _, err := s.docs.open(uri, source); err != nil {
		t.Fatal(err)
	}
	tokens, err := s.semanticTokensFull(nil, &protocol.SemanticTokensParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(tokens.Data) == 0 || len(tokens.Data)%5 != 0 {
		t.Fatalf("semantic token data length = %d, want non-empty multiple of 5: %+v", len(tokens.Data), tokens.Data)
	}
	decoded := decodeSemanticTokenTypes(tokens.Data)
	if !containsString(decoded, "function") {
		t.Fatalf("semantic token data missing function token: decoded=%+v data=%+v", decoded, tokens.Data)
	}
	if !containsString(decoded, "property") {
		t.Fatalf("semantic token data missing property token: decoded=%+v data=%+v", decoded, tokens.Data)
	}
	if !containsString(decoded, "enumMember") {
		t.Fatalf("semantic token data missing enum member token: decoded=%+v data=%+v", decoded, tokens.Data)
	}
	assertSemanticTokenDeltasValid(t, tokens.Data)
}

type rpcRecorder struct {
	mu          sync.Mutex
	methodsSeen []string
	paramsSeen  map[string][]json.RawMessage
}

func (r *rpcRecorder) Handle(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	r.mu.Lock()
	r.methodsSeen = append(r.methodsSeen, req.Method)
	if r.paramsSeen == nil {
		r.paramsSeen = map[string][]json.RawMessage{}
	}
	if req.Params != nil {
		copied := append(json.RawMessage(nil), (*req.Params)...)
		r.paramsSeen[req.Method] = append(r.paramsSeen[req.Method], copied)
	}
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

func (r *rpcRecorder) publishDiagnostics() []protocol.PublishDiagnosticsParams {
	r.mu.Lock()
	defer r.mu.Unlock()
	var out []protocol.PublishDiagnosticsParams
	for _, raw := range r.paramsSeen[string(protocol.ServerTextDocumentPublishDiagnostics)] {
		var params protocol.PublishDiagnosticsParams
		if err := json.Unmarshal(raw, &params); err == nil {
			out = append(out, params)
		}
	}
	return out
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

func decodeSemanticTokenTypes(data []protocol.UInteger) []string {
	var out []string
	for i := 0; i+4 < len(data); i += 5 {
		typeIndex := int(data[i+3])
		if typeIndex >= 0 && typeIndex < len(intel.SemanticTokenTypes) {
			out = append(out, intel.SemanticTokenTypes[typeIndex])
		}
	}
	return out
}

func assertSemanticTokenDeltasValid(t *testing.T, data []protocol.UInteger) {
	t.Helper()
	line, character := 0, 0
	for i := 0; i+4 < len(data); i += 5 {
		deltaLine := int(data[i])
		deltaStart := int(data[i+1])
		length := int(data[i+2])
		if length <= 0 {
			t.Fatalf("semantic token at %d has non-positive length: %+v", i, data[i:i+5])
		}
		if deltaLine == 0 {
			if deltaStart < 0 {
				t.Fatalf("semantic token at %d has negative delta start: %+v", i, data[i:i+5])
			}
			character += deltaStart
		} else {
			line += deltaLine
			character = deltaStart
		}
		if line < 0 || character < 0 {
			t.Fatalf("semantic token at %d decoded to negative position: line=%d character=%d", i, line, character)
		}
	}
}
