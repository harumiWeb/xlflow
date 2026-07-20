package lspserver

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestUserFormYAMLUsesRawDocumentAndCompletion(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	path := filepath.Join(root, "src", "forms", "specs", "Login.yaml")
	uri := pathToFileURI(path)
	doc, err := s.docs.open(uri, `schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: Login
controls: []
`)
	if err != nil {
		t.Fatal(err)
	}
	if got := s.documentKind(doc); got != DocumentKindUserFormYAML {
		t.Fatalf("document kind = %v, want UserForm YAML", got)
	}
	s.docs.mu.RLock()
	entry := s.docs.docs[normalizePathKey(path)]
	s.docs.mu.RUnlock()
	if entry.snapshot != nil {
		t.Fatal("UserForm YAML must not create a VBA analysis snapshot")
	}
	if got := s.documentDiagnostics(context.Background(), doc); len(got) != 0 {
		t.Fatalf("valid UserForm YAML diagnostics = %#v, want none", got)
	}

	completion, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Position: protocol.Position{Line: 2, Character: 4},
	}})
	if err != nil {
		t.Fatal(err)
	}
	list := completion.(protocol.CompletionList)
	if !hasCompletionLabel(list.Items, "basis") {
		t.Fatalf("UserForm YAML completion = %#v, want basis", list.Items)
	}
	hover, err := s.hover(nil, &protocol.HoverParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Position: protocol.Position{Line: 0, Character: 2},
	}})
	if err != nil || hover != nil {
		t.Fatalf("UserForm YAML hover = %#v, %v; want nil, nil", hover, err)
	}
}

func TestUserFormYAMLMalformedDiagnosticDoesNotBreakLaterRequests(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	uri := pathToFileURI(filepath.Join(root, "src", "forms", "specs", "Login.yml"))
	doc, err := s.docs.open(uri, "controls:\n  - type: TextBox\n    te\n")
	if err != nil {
		t.Fatal(err)
	}
	diagnostics := s.documentDiagnostics(context.Background(), doc)
	if len(diagnostics) != 1 || diagnostics[0].Code != "UFY001" {
		t.Fatalf("malformed YAML diagnostics = %#v, want UFY001", diagnostics)
	}
	completion, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Position: protocol.Position{Line: 2, Character: 6},
	}})
	if err != nil {
		t.Fatalf("completion after malformed YAML failed: %v", err)
	}
	if !hasCompletionLabel(completion.(protocol.CompletionList).Items, "text") {
		t.Fatalf("incomplete YAML completion = %#v, want TextBox text property", completion)
	}
}

func TestUserFormYAMLCompletionUsesProtocolEditsAndSnippets(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	valueURI := pathToFileURI(filepath.Join(root, "src", "forms", "specs", "Values.yaml"))
	if _, err := s.docs.open(valueURI, "controls:\n  - type: Label\n    enabled: \n"); err != nil {
		t.Fatal(err)
	}
	completion, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(valueURI)},
		Position:     protocol.Position{Line: 2, Character: protocol.UInteger(len("    enabled: "))},
	}})
	if err != nil {
		t.Fatal(err)
	}
	trueItem, ok := completionByProtocolLabel(completion.(protocol.CompletionList).Items, "true")
	edit, editOK := trueItem.TextEdit.(protocol.TextEdit)
	if !ok || !editOK || edit.NewText != "true" {
		t.Fatalf("boolean completion = %#v, want true text edit", trueItem)
	}

	snippetURI := pathToFileURI(filepath.Join(root, "src", "forms", "specs", "Snippet.yaml"))
	if _, err := s.docs.open(snippetURI, ""); err != nil {
		t.Fatal(err)
	}
	completion, err = s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(snippetURI)},
	}})
	if err != nil {
		t.Fatal(err)
	}
	snippet, ok := completionByProtocolLabel(completion.(protocol.CompletionList).Items, "UserForm document")
	if !ok || snippet.InsertTextFormat == nil || *snippet.InsertTextFormat != protocol.InsertTextFormatSnippet {
		t.Fatalf("document snippet = %#v, want snippet insert format", snippet)
	}
}

func TestUserFormYAMLCompletionSupportsIndentationlessControls(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	uri := pathToFileURI(filepath.Join(root, "src", "forms", "specs", "Indentationless.yaml"))
	if _, err := s.docs.open(uri, "controls:\n- type: Label\n  ca"); err != nil {
		t.Fatal(err)
	}
	completion, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
		Position:     protocol.Position{Line: 2, Character: 4},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletionLabel(completion.(protocol.CompletionList).Items, "caption") {
		t.Fatalf("indentationless completion = %#v, want caption", completion)
	}
}

func hasCompletionLabel(items []protocol.CompletionItem, want string) bool {
	_, ok := completionByProtocolLabel(items, want)
	return ok
}

func completionByProtocolLabel(items []protocol.CompletionItem, want string) (protocol.CompletionItem, bool) {
	for _, item := range items {
		if item.Label == want {
			return item, true
		}
	}
	return protocol.CompletionItem{}, false
}

func TestUnrelatedYAMLIsIgnored(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	uri := pathToFileURI(filepath.Join(root, "notes.yaml"))
	doc, err := s.docs.open(uri, "kind: xlflow.userform\n")
	if err != nil {
		t.Fatal(err)
	}
	if got := s.documentKind(doc); got != DocumentKindUnknown {
		t.Fatalf("document kind = %v, want unknown", got)
	}
	if diagnostics := s.documentDiagnostics(context.Background(), doc); len(diagnostics) != 0 {
		t.Fatalf("unrelated YAML diagnostics = %#v, want none", diagnostics)
	}
}
