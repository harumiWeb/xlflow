package lspserver

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestUserFormYAMLUsesRawDocumentAndEmptyFeatureResponses(t *testing.T) {
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
	if len(list.Items) != 0 {
		t.Fatalf("UserForm YAML completion = %#v, want empty foundation response", list.Items)
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
	doc, err := s.docs.open(uri, "controls:\n  - type: TextBox\n    ca\n")
	if err != nil {
		t.Fatal(err)
	}
	diagnostics := s.documentDiagnostics(context.Background(), doc)
	if len(diagnostics) != 1 || diagnostics[0].Code != "UFY001" {
		t.Fatalf("malformed YAML diagnostics = %#v, want UFY001", diagnostics)
	}
	if _, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}, Position: protocol.Position{Line: 2, Character: 6},
	}}); err != nil {
		t.Fatalf("completion after malformed YAML failed: %v", err)
	}
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
