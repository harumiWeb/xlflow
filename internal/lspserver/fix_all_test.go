package lspserver

import (
	"context"
	"encoding/json"
	"net"
	"os"
	"path/filepath"
	"slices"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const fixAllLetKind = "source.fixAll.xlflow.vb080"
const fixAllLetSource = "Option Explicit\r\nSub Work()\r\nDim count As Long\r\n    Let count = 1 ' 日本語😀\r\n    Let count = 2\r\nEnd Sub\r\n"

func TestDocumentFixAllRequiresOpenVersionedDocument(t *testing.T) {
	for _, versioned := range []bool{false, true} {
		for _, open := range []bool{false, true} {
			root := t.TempDir()
			cfg := config.Default()
			cfg.Lint.DetectLetAssignment = true
			s, cleanup, err := New(Options{RootDir: root, Config: cfg, TypeDBDir: t.TempDir()})
			if err != nil {
				t.Fatal(err)
				return
			}
			t.Cleanup(cleanup)
			s.codeActionDocumentChanges.Store(versioned)
			path := filepath.Join(root, "Main.bas")
			if err := os.WriteFile(path, []byte(fixAllLetSource), 0o644); err != nil {
				t.Fatal(err)
			}
			uri := protocol.DocumentUri(pathToFileURI(path))
			if open {
				if _, err := s.docs.open(string(uri), fixAllLetSource, 9); err != nil {
					t.Fatal(err)
				}
			}
			result, err := s.codeAction(nil, &protocol.CodeActionParams{
				TextDocument: protocol.TextDocumentIdentifier{URI: uri},
				Context:      protocol.CodeActionContext{Only: []protocol.CodeActionKind{fixAllLetKind}},
			})
			if err != nil {
				t.Fatal(err)
			}
			actions := result.([]protocol.CodeAction)
			if !versioned || !open {
				if len(actions) != 0 {
					t.Fatalf("versioned=%v open=%v: unexpected actions: %+v", versioned, open, actions)
				}
				doc, err := s.docs.getOrRead(string(uri))
				if err != nil {
					t.Fatal(err)
				}
				if doc.Snapshot.ParseCount() != 0 {
					t.Fatal("unsupported Fix All request parsed the document")
				}
				continue
			}
			if len(actions) != 1 || actions[0].Kind == nil || string(*actions[0].Kind) != fixAllLetKind || len(actions[0].Diagnostics) != 2 {
				t.Fatalf("Fix All action = %+v", actions)
				return
			}
			if actions[0].Edit == nil || len(actions[0].Edit.Changes) != 0 || len(actions[0].Edit.DocumentChanges) != 1 {
				t.Fatalf("expected one versioned document transaction: %+v", actions[0].Edit)
				return
			}
			change := actions[0].Edit.DocumentChanges[0].(protocol.TextDocumentEdit)
			if change.TextDocument.URI != uri || change.TextDocument.Version == nil || *change.TextDocument.Version != 9 || len(change.Edits) != 2 {
				t.Fatalf("versioned edits = %+v", change)
			}
			for i, raw := range change.Edits {
				edit := raw.(protocol.TextEdit)
				if edit.NewText != "" || edit.Range != (protocol.Range{Start: protocol.Position{Line: protocol.UInteger(3 + i), Character: 4}, End: protocol.Position{Line: protocol.UInteger(3 + i), Character: 8}}) {
					t.Fatalf("edit %d = %+v", i, edit)
				}
			}
		}
	}
}

func TestDocumentFixAllCapabilities(t *testing.T) {
	for _, versioned := range []bool{false, true} {
		s, cleanup, err := New(Options{RootDir: t.TempDir(), Config: config.Default(), TypeDBDir: t.TempDir()})
		if err != nil {
			t.Fatal(err)
			return
		}
		t.Cleanup(cleanup)
		// Decode through JSON, as the public capability negotiation does.
		params := fixAllInitializeParams(t, versioned)
		result, err := s.initialize(nil, params)
		if err != nil {
			t.Fatal(err)
		}
		options := result.(protocol.InitializeResult).Capabilities.CodeActionProvider.(protocol.CodeActionOptions)
		want := []protocol.CodeActionKind{protocol.CodeActionKindQuickFix, protocol.CodeActionKindRefactorRewrite}
		if versioned {
			for _, kind := range intel.FixAllCodeActionKinds() {
				want = append(want, protocol.CodeActionKind(kind))
			}
		}
		if !slices.Equal(options.CodeActionKinds, want) {
			t.Fatalf("versioned=%v kinds=%+v, want %+v", versioned, options.CodeActionKinds, want)
		}
	}
}

func TestJSONRPCDocumentFixAllSkipsAutomaticAggregation(t *testing.T) {
	ctx, cancel := context.WithTimeout(t.Context(), jsonrpcIntegrationTimeout)
	defer cancel()
	cfg := config.Default()
	cfg.Lint.DetectLetAssignment = true
	s, cleanup, err := New(Options{RootDir: t.TempDir(), Config: cfg, TypeDBDir: t.TempDir()})
	if err != nil {
		t.Fatal(err)
		return
	}
	defer cleanup()
	uri := pathToFileURI(filepath.Join(s.opts.RootDir, "Main.bas"))
	doc, err := s.docs.open(uri, fixAllLetSource, 3)
	if err != nil {
		t.Fatal(err)
	}
	serverSide, clientSide := net.Pipe()
	serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{handler: &s.handler, dispatch: s.dispatchCodeAction})
	defer func() { _ = serverConn.Close() }()
	clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), &rpcRecorder{})
	defer func() { _ = clientConn.Close() }()
	var initialized protocol.InitializeResult
	if err := clientConn.Call(ctx, "initialize", fixAllInitializeParams(t, true), &initialized); err != nil {
		t.Fatal(err)
	}
	var actions []protocol.CodeAction
	if err := clientConn.Call(ctx, "textDocument/codeAction", map[string]any{
		"textDocument": map[string]any{"uri": uri},
		"range":        protocol.Range{},
		"context":      map[string]any{"only": []string{fixAllLetKind}, "triggerKind": 2},
	}, &actions); err != nil {
		t.Fatal(err)
	}
	if len(actions) != 0 || doc.Snapshot.ParseCount() != 0 {
		t.Fatalf("automatic source request aggregated: actions=%+v parses=%d", actions, doc.Snapshot.ParseCount())
	}
	if err := clientConn.Call(ctx, "textDocument/codeAction", map[string]any{
		"textDocument": map[string]any{"uri": uri},
		"range":        protocol.Range{},
		"context":      map[string]any{"only": []string{fixAllLetKind}, "triggerKind": 1},
	}, &actions); err != nil {
		t.Fatal(err)
	}
	if len(actions) != 1 || actions[0].Kind == nil || string(*actions[0].Kind) != fixAllLetKind {
		t.Fatalf("explicit source request = %+v", actions)
	}
}

func fixAllInitializeParams(t *testing.T, versioned bool) *protocol.InitializeParams {
	t.Helper()
	body := `{"capabilities":{"workspace":{"workspaceEdit":{"documentChanges":false}}}}`
	if versioned {
		body = `{"capabilities":{"workspace":{"workspaceEdit":{"documentChanges":true}}}}`
	}
	var params protocol.InitializeParams
	if err := json.Unmarshal([]byte(body), &params); err != nil {
		t.Fatal(err)
	}
	return &params
}
