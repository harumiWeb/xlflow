package lspserver

import (
	"context"
	"fmt"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func TestSemanticTokensFullDeltaForLargeModuleEdits(t *testing.T) {
	base := largeModuleSource()
	cases := []struct {
		name string
		edit func(string) string
	}{
		{
			name: "beginning",
			edit: func(source string) string {
				return strings.Replace(source, "Option Explicit", "Option Private Module", 1)
			},
		},
		{
			name: "middle",
			edit: func(source string) string {
				return strings.Replace(source, "LargeProcedure012", "LargeProcedure012Renamed", 1)
			},
		},
		{
			name: "end",
			edit: func(source string) string {
				return source + "' End-of-module semantic token delta fixture.\n"
			},
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			s, ctx, uri := newSemanticTokenDeltaServer(t, base)
			params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}
			before, err := s.semanticTokensFull(nil, params)
			if err != nil || before.ResultID == nil || *before.ResultID == "" {
				t.Fatalf("initial full response = (%+v, %v), want resultId", before, err)
			}
			if err := s.didChange(ctx, wholeSemanticTokenChange(uri, 2, tc.edit(base))); err != nil {
				t.Fatal(err)
			}
			response, err := s.semanticTokensFullDelta(nil, &protocol.SemanticTokensDeltaParams{
				TextDocument:     protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
				PreviousResultID: *before.ResultID,
			})
			if err != nil {
				t.Fatal(err)
			}
			delta, ok := response.(*protocol.SemanticTokensDelta)
			if !ok {
				t.Fatalf("delta response = %T, want *SemanticTokensDelta", response)
			}
			after, err := s.semanticTokensFull(nil, params)
			if err != nil {
				t.Fatal(err)
			}
			if after.ResultID == nil || delta.ResultId == nil || *delta.ResultId != *after.ResultID {
				t.Fatalf("delta result id = %+v, full result id = %+v", delta.ResultId, after.ResultID)
			}
			applied := applySemanticTokenEdits(before.Data, delta.Edits)
			s.semanticTokens.invalidateWorkspace()
			fresh, err := s.semanticTokensFull(nil, params)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(applied, fresh.Data) {
				t.Fatalf("applied delta differs from fresh full response\ngot:  %v\nwant: %v", applied, fresh.Data)
			}
			if semanticTokenResponseSize(delta) >= semanticTokenResponseSize(after) {
				t.Fatalf("delta payload = %d, full payload = %d; want delta smaller", semanticTokenResponseSize(delta), semanticTokenResponseSize(after))
			}
		})
	}
}

func TestSemanticTokensFullDeltaReturnsEmptyDeltaForUnchangedDocument(t *testing.T) {
	s, _, uri := newSemanticTokenDeltaServer(t, largeModuleSource())
	params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}
	first, err := s.semanticTokensFull(nil, params)
	if err != nil || first.ResultID == nil {
		t.Fatalf("initial full response = (%+v, %v)", first, err)
	}
	second, err := s.semanticTokensFull(nil, params)
	if err != nil || second.ResultID == nil || *second.ResultID != *first.ResultID {
		t.Fatalf("repeated full response = (%+v, %v), want stable resultId %q", second, err, *first.ResultID)
	}
	response, err := s.semanticTokensFullDelta(nil, &protocol.SemanticTokensDeltaParams{
		TextDocument:     protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
		PreviousResultID: *first.ResultID,
	})
	if err != nil {
		t.Fatal(err)
	}
	delta, ok := response.(*protocol.SemanticTokensDelta)
	if !ok || len(delta.Edits) != 0 || delta.ResultId == nil || *delta.ResultId != *first.ResultID {
		t.Fatalf("unchanged delta = %+v, want empty delta with stable result id", response)
	}
}

func TestSemanticTokensFullDeltaFallsBackForUnknownExpiredAndUnrelatedResults(t *testing.T) {
	s, ctx, uri := newSemanticTokenDeltaServer(t, largeModuleSource())
	params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}
	initial, err := s.semanticTokensFull(nil, params)
	if err != nil || initial.ResultID == nil {
		t.Fatalf("initial full response = (%+v, %v)", initial, err)
	}
	assertSemanticTokenFullFallback(t, s, uri, "unknown-result")

	for version := 2; version <= semanticTokenHistoryLimit+2; version++ {
		source := largeModuleSource() + fmt.Sprintf("' Version %d\n", version)
		if err := s.didChange(ctx, wholeSemanticTokenChange(uri, int32(version), source)); err != nil {
			t.Fatal(err)
		}
		if _, err := s.semanticTokensFull(nil, params); err != nil {
			t.Fatal(err)
		}
	}
	assertSemanticTokenFullFallback(t, s, uri, *initial.ResultID)

	otherURI := pathToFileURI(filepath.Join(t.TempDir(), "src", "modules", "Other.bas"))
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(otherURI), LanguageID: "vba", Version: 1, Text: largeModuleSource(),
	}}); err != nil {
		t.Fatal(err)
	}
	other, err := s.semanticTokensFull(nil, &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(otherURI)}})
	if err != nil || other.ResultID == nil {
		t.Fatalf("other full response = (%+v, %v)", other, err)
	}
	assertSemanticTokenFullFallback(t, s, uri, *other.ResultID)

	beforeClose, err := s.semanticTokensFull(nil, params)
	if err != nil || beforeClose.ResultID == nil {
		t.Fatalf("pre-close full response = (%+v, %v)", beforeClose, err)
	}
	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}); err != nil {
		t.Fatal(err)
	}
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 99, Text: largeModuleSource(),
	}}); err != nil {
		t.Fatal(err)
	}
	assertSemanticTokenFullFallback(t, s, uri, *beforeClose.ResultID)
}

func TestSemanticTokenHistoryIsBoundedAndClearedOnClose(t *testing.T) {
	s, ctx, uri := newSemanticTokenDeltaServer(t, largeModuleSource())
	params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}
	for version := 1; version <= semanticTokenHistoryLimit+2; version++ {
		if version > 1 {
			source := largeModuleSource() + fmt.Sprintf("' Version %d\n", version)
			if err := s.didChange(ctx, wholeSemanticTokenChange(uri, int32(version), source)); err != nil {
				t.Fatal(err)
			}
		}
		if _, err := s.semanticTokensFull(nil, params); err != nil {
			t.Fatal(err)
		}
	}
	path, err := fileURIToPath(uri)
	if err != nil {
		t.Fatal(err)
	}
	identity := documentSymbolKey(intel.Document{URI: uri, Path: path})
	s.semanticTokens.mu.Lock()
	historyCount := len(s.semanticTokens.histories[identity])
	s.semanticTokens.mu.Unlock()
	if historyCount != semanticTokenHistoryLimit {
		t.Fatalf("history count = %d, want %d", historyCount, semanticTokenHistoryLimit)
	}
	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}); err != nil {
		t.Fatal(err)
	}
	s.semanticTokens.mu.Lock()
	historyCount = len(s.semanticTokens.histories[identity])
	s.semanticTokens.mu.Unlock()
	if historyCount != 0 {
		t.Fatalf("history count after close = %d, want 0", historyCount)
	}
}

func newSemanticTokenDeltaServer(t *testing.T, source string) (*Server, *glsp.Context, string) {
	t.Helper()
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(cleanup)
	s.diagnostics = func(context.Context, intel.Document) []intel.Diagnostic { return nil }
	uri := pathToFileURI(filepath.Join(root, "src", "modules", "LargeModule.bas"))
	ctx := &glsp.Context{Notify: func(string, any) {}}
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 1, Text: source,
	}}); err != nil {
		t.Fatal(err)
	}
	return s, ctx, uri
}

func wholeSemanticTokenChange(uri string, version int32, text string) *protocol.DidChangeTextDocumentParams {
	return &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Version: version,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: text}},
	}
}

func assertSemanticTokenFullFallback(t *testing.T, s *Server, uri, previousResultID string) {
	t.Helper()
	response, err := s.semanticTokensFullDelta(nil, &protocol.SemanticTokensDeltaParams{
		TextDocument:     protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
		PreviousResultID: previousResultID,
	})
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := response.(*protocol.SemanticTokens); !ok {
		t.Fatalf("delta response with base %q = %T, want *SemanticTokens fallback", previousResultID, response)
	}
}

func applySemanticTokenEdits(data []protocol.UInteger, edits []protocol.SemanticTokensEdit) []protocol.UInteger {
	out := append([]protocol.UInteger(nil), data...)
	for _, edit := range edits {
		start := int(edit.Start)
		end := start + int(edit.DeleteCount)
		out = append(append(append([]protocol.UInteger(nil), out[:start]...), edit.Data...), out[end:]...)
	}
	return out
}
