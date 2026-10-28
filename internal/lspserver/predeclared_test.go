package lspserver

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestWorkspaceVB029PredeclaredClassOverlay(t *testing.T) {
	root := t.TempDir()
	classPath := filepath.Join(root, "src", "classes", "Predeclared.cls")
	mainPath := filepath.Join(root, "src", "modules", "Main.bas")
	classSource := "Attribute VB_Name = \"Predeclared\"\nAttribute VB_PredeclaredId = False\nOption Explicit\n"
	mainSource := "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n  Set Predeclared = Nothing\nEnd Sub\n"
	for path, source := range map[string]string{classPath: classSource, mainPath: mainSource} {
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	server, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := server.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}
	classURI := protocol.DocumentUri(pathToFileURI(classPath))
	mainURI := protocol.DocumentUri(pathToFileURI(mainPath))
	published := make(chan protocol.PublishDiagnosticsParams, 32)
	ctx := &glsp.Context{Notify: func(method string, params any) {
		if method == string(protocol.ServerTextDocumentPublishDiagnostics) {
			select {
			case published <- params.(protocol.PublishDiagnosticsParams):
			case <-t.Context().Done():
			}
		}
	}}
	waitForVB029 := func(uri protocol.DocumentUri, want int) {
		t.Helper()
		waitCtx, cancel := context.WithTimeout(t.Context(), time.Minute)
		defer cancel()
		var latest []protocol.Diagnostic
		for {
			select {
			case params := <-published:
				if params.URI != uri {
					continue
				}
				latest = params.Diagnostics
				count := 0
				for _, diagnostic := range params.Diagnostics {
					if diagnostic.Code != nil && diagnostic.Code.Value == "VB029" {
						count++
					}
				}
				if count == want {
					return
				}
			case <-waitCtx.Done():
				t.Fatalf("waiting for VB029 count %d for %s: %v; latest=%+v", want, uri, waitCtx.Err(), latest)
			}
		}
	}
	// Keep document edits fast while using the production diagnostics pipeline.
	server.diagnosticsDebounce = time.Millisecond
	if err := server.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: classURI, LanguageID: "vba", Version: 1, Text: classSource,
	}}); err != nil {
		t.Fatal(err)
	}
	waitForVB029(classURI, 0)
	if err := server.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: mainURI, LanguageID: "vba", Version: 1, Text: mainSource,
	}}); err != nil {
		t.Fatal(err)
	}
	waitForVB029(mainURI, 1)
	for version, value := range []string{"False", "True", "False"} {
		for _, change := range []struct {
			uri  protocol.DocumentUri
			text string
		}{
			{classURI, strings.Replace(classSource, "False", value, 1)},
			// A normal edit requests fresh diagnostics for the dependent module.
			// Both documents remain unsaved, and the production snapshot callback
			// must discover the changed class through the server's document store.
			{mainURI, mainSource + strings.Repeat("\n", version+1)},
		} {
			if err := server.didChange(ctx, &protocol.DidChangeTextDocumentParams{
				TextDocument: protocol.VersionedTextDocumentIdentifier{
					TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: change.uri}, Version: int32(version + 2),
				},
				ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: change.text}},
			}); err != nil {
				t.Fatal(err)
			}
			if change.uri == classURI {
				waitForVB029(classURI, 0)
			}
		}
		want := 1
		if value == "True" {
			want = 0
		}
		waitForVB029(mainURI, want)
		if disk, err := os.ReadFile(classPath); err != nil || string(disk) != classSource {
			t.Fatalf("class must remain unsaved: disk=%q, err=%v", disk, err)
		}
	}
}
