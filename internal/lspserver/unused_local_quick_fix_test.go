package lspserver

import (
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestCodeActionRemovesUnusedLocalFromUnsavedBufferWithVersionedEdit(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	s, cleanup, err := New(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(cleanup)
	s.analyzer.DB = nil
	s.analyzer.SourceOnly = true
	s.codeActionDocumentChanges.Store(true)

	path := filepath.Join(root, "Main.bas")
	uri := protocol.DocumentUri(pathToFileURI(path))
	diskSource := "Option Explicit\r\nSub Work()\r\n    Dim scratch As Long\r\n    Debug.Print scratch\r\nEnd Sub\r\n"
	if err := os.WriteFile(path, []byte(diskSource), 0o600); err != nil {
		t.Fatal(err)
	}
	bufferSource := "Option Explicit\r\nSub Work()\r\n    Dim scratch As Long ' keep 日本語😀\r\nEnd Sub\r\n"
	doc, err := s.docs.open(string(uri), bufferSource, 37)
	if err != nil {
		t.Fatal(err)
	}

	result, err := s.codeAction(nil, &protocol.CodeActionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: uri},
		Range: protocol.Range{
			Start: protocol.Position{Line: 2, Character: 0},
			End:   protocol.Position{Line: 3, Character: 0},
		},
		Context: protocol.CodeActionContext{Only: []protocol.CodeActionKind{protocol.CodeActionKindQuickFix}},
	})
	if err != nil {
		t.Fatal(err)
	}
	actions, ok := result.([]protocol.CodeAction)
	if !ok || len(actions) != 1 {
		t.Fatalf("code actions = %#v, want one VB020 action", result)
	}
	action := actions[0]
	if action.Title != "Remove unused local variable" || action.Kind == nil || *action.Kind != protocol.CodeActionKindQuickFix {
		t.Fatalf("action = %+v, want the individual VB020 quick fix", action)
	}
	var published []intel.Diagnostic
	for _, diagnostic := range s.analyzer.Diagnostics(doc) {
		if diagnostic.Code == "VB020" {
			published = append(published, diagnostic)
		}
	}
	if len(published) != 1 {
		t.Fatalf("published VB020 diagnostics = %+v, want one for the unsaved buffer", published)
	}
	wantDiagnostic := toProtocolDiagnostic(published[0])
	if !reflect.DeepEqual(action.Diagnostics, []protocol.Diagnostic{wantDiagnostic}) {
		t.Fatalf("associated diagnostics = %+v, want published metadata %+v", action.Diagnostics, wantDiagnostic)
	}

	if action.Edit == nil || len(action.Edit.Changes) != 0 || len(action.Edit.DocumentChanges) != 1 {
		t.Fatalf("workspace edit = %+v, want one versioned document change", action.Edit)
	}
	change, ok := action.Edit.DocumentChanges[0].(protocol.TextDocumentEdit)
	if !ok {
		t.Fatalf("document change = %#v, want TextDocumentEdit", action.Edit.DocumentChanges[0])
	}
	if change.TextDocument.URI != uri || change.TextDocument.Version == nil || *change.TextDocument.Version != 37 {
		t.Fatalf("versioned target = %+v, want open buffer %s version 37", change.TextDocument, uri)
	}
	if len(change.Edits) != 1 {
		t.Fatalf("versioned edits = %+v, want one declaration deletion", change.Edits)
	}
	edit, ok := change.Edits[0].(protocol.TextEdit)
	if !ok {
		t.Fatalf("versioned edit = %#v, want TextEdit", change.Edits[0])
	}
	wantRange := protocol.Range{
		Start: protocol.Position{Line: 2, Character: 4},
		End:   protocol.Position{Line: 2, Character: 23},
	}
	if edit.Range != wantRange || edit.NewText != "" {
		t.Fatalf("versioned edit = %+v, want deletion of statement span %+v", edit, wantRange)
	}
}
