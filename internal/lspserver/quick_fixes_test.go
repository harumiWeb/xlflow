package lspserver

import (
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestCodeActionProjectsEveryEditAndDiagnostic(t *testing.T) {
	for _, versioned := range []bool{false, true} {
		for _, open := range []bool{false, true} {
			root := t.TempDir()
			s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
			if err != nil {
				t.Fatal(err)
			}
			t.Cleanup(cleanup)
			s.codeActionDocumentChanges.Store(versioned)
			uri := pathToFileURI(filepath.Join(root, "Main.bas"))
			doc := intel.Document{URI: uri, Path: filepath.Join(root, "Main.bas"), Version: 7}
			if open {
				if _, err := s.docs.open(uri, "", 7); err != nil {
					t.Fatal(err)
				}
			}
			in := intel.CodeAction{
				Title: "Correct two locations", Kind: "quickfix",
				Edits: []intel.CodeActionEdit{
					{Range: intel.Range{Start: intel.Position{Line: 1, Character: 2}, End: intel.Position{Line: 1, Character: 4}}, NewText: "First"},
					{Range: intel.Range{Start: intel.Position{Line: 3, Character: 6}, End: intel.Position{Line: 3, Character: 8}}, NewText: "Second"},
				},
				Diagnostics: []intel.Diagnostic{{Code: "VB084", Severity: "information", Source: "xlflow", Message: "Step 1 restates the default step.", Range: intel.Range{Start: intel.Position{Line: 1, Character: 9}, End: intel.Position{Line: 1, Character: 10}}}},
			}
			out := s.projectCodeActions(doc, protocol.DocumentUri(uri), []intel.CodeAction{in})
			if len(out) != 1 || out[0].Kind == nil || *out[0].Kind != protocol.CodeActionKindQuickFix || len(out[0].Diagnostics) != 1 {
				t.Fatalf("projected actions = %+v", out)
				return
			}
			if got := out[0].Diagnostics[0]; got.Code == nil || got.Code.Value != "VB084" || got.Range != toProtocolRange(in.Diagnostics[0].Range) || got.CodeDescription == nil {
				t.Fatalf("associated diagnostic = %+v", got)
			}
			var edits []protocol.TextEdit
			if versioned {
				if len(out[0].Edit.DocumentChanges) != 1 || len(out[0].Edit.Changes) != 0 {
					t.Fatalf("versioned edit = %+v", out[0].Edit)
				}
				change := out[0].Edit.DocumentChanges[0].(protocol.TextDocumentEdit)
				if change.TextDocument.URI != protocol.DocumentUri(uri) || (open && (change.TextDocument.Version == nil || *change.TextDocument.Version != 7)) || (!open && change.TextDocument.Version != nil) {
					t.Fatalf("versioned target = %+v", change.TextDocument)
				}
				for _, edit := range change.Edits {
					edits = append(edits, edit.(protocol.TextEdit))
				}
			} else {
				if len(out[0].Edit.DocumentChanges) != 0 {
					t.Fatal("legacy projection returned documentChanges")
				}
				edits = out[0].Edit.Changes[protocol.DocumentUri(uri)]
			}
			if len(edits) != len(in.Edits) {
				t.Fatalf("edits = %+v", edits)
			}
			for i, edit := range edits {
				if edit.Range != toProtocolRange(in.Edits[i].Range) || edit.NewText != in.Edits[i].NewText {
					t.Fatalf("edit %d = %+v", i, edit)
				}
			}
		}
	}
}

func TestCodeActionCorrectsRedundantStepInUnsavedBuffer(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Lint.DetectRedundantStepOne = true
	s, cleanup, err := New(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	uri := pathToFileURI(filepath.Join(root, "Main.bas"))
	if _, err := s.docs.open(uri, "Option Explicit\r\nSub Work()\r\nDim i As Long\r\nFor i = 1 To 3 Step 1 ' keep\r\nDebug.Print i\r\nNext i\r\nEnd Sub\r\n", 4); err != nil {
		t.Fatal(err)
	}
	result, err := s.codeAction(nil, &protocol.CodeActionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
		Range:        protocol.Range{Start: protocol.Position{Line: 3, Character: 19}, End: protocol.Position{Line: 3, Character: 20}},
		Context:      protocol.CodeActionContext{Only: []protocol.CodeActionKind{protocol.CodeActionKindQuickFix}},
	})
	if err != nil {
		t.Fatal(err)
	}
	actions := result.([]protocol.CodeAction)
	if len(actions) != 1 || len(actions[0].Diagnostics) != 1 || actions[0].Diagnostics[0].Code.Value != "VB084" {
		t.Fatalf("actions = %+v", actions)
		return
	}
	edits := actions[0].Edit.Changes[protocol.DocumentUri(uri)]
	if len(edits) != 1 || edits[0].NewText != "" || edits[0].Range != (protocol.Range{Start: protocol.Position{Line: 3, Character: 14}, End: protocol.Position{Line: 3, Character: 21}}) {
		t.Fatalf("VB084 edits = %+v", edits)
	}
}

func TestCodeActionCorrectsSetInUnsavedBuffer(t *testing.T) {
	for _, versioned := range []bool{false, true} {
		root := t.TempDir()
		s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
		if err != nil {
			t.Fatal(err)
		}
		t.Cleanup(cleanup)
		s.codeActionDocumentChanges.Store(versioned)
		uri := protocol.DocumentUri(pathToFileURI(filepath.Join(root, "Main.bas")))
		if _, err := s.docs.open(string(uri), "Option Explicit\r\nSub Work()\r\nDim count As Long\r\n    Set count = 1 ' keep 日本語😀\r\nEnd Sub\r\n", 9); err != nil {
			t.Fatal(err)
		}
		result, err := s.codeAction(nil, &protocol.CodeActionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: uri},
			Range:        protocol.Range{Start: protocol.Position{Line: 3, Character: 8}, End: protocol.Position{Line: 3, Character: 13}},
			Context:      protocol.CodeActionContext{Only: []protocol.CodeActionKind{protocol.CodeActionKindQuickFix}},
		})
		if err != nil {
			t.Fatal(err)
		}
		actions := result.([]protocol.CodeAction)
		if len(actions) != 1 || len(actions[0].Diagnostics) != 1 || actions[0].Diagnostics[0].Code.Value != "VB037" {
			t.Fatalf("actions = %+v", actions)
			return
		}
		var edits []protocol.TextEdit
		if versioned {
			change := actions[0].Edit.DocumentChanges[0].(protocol.TextDocumentEdit)
			if change.TextDocument.Version == nil || *change.TextDocument.Version != 9 {
				t.Fatalf("version = %+v", change.TextDocument)
			}
			for _, edit := range change.Edits {
				edits = append(edits, edit.(protocol.TextEdit))
			}
		} else {
			edits = actions[0].Edit.Changes[uri]
		}
		if len(edits) != 1 || edits[0].NewText != "" || edits[0].Range != (protocol.Range{Start: protocol.Position{Line: 3, Character: 4}, End: protocol.Position{Line: 3, Character: 8}}) {
			t.Fatalf("Set edits = %+v", edits)
		}
	}
}

func TestCodeActionWithholdsSetForDefaultMemberTypes(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(cleanup)
	for _, typeName := range []string{"Excel.Range", "ProjectValue"} {
		t.Run(typeName, func(t *testing.T) {
			if typeName == "ProjectValue" {
				classURI := pathToFileURI(filepath.Join(root, "ProjectValue.cls"))
				if _, err := s.docs.open(classURI, "Attribute VB_Name = \"ProjectValue\"\nPublic Property Get Value() As Long\nAttribute Value.VB_UserMemId = 0\nValue = 1\nEnd Property\n", 1); err != nil {
					t.Fatal(err)
				}
			}
			uri := protocol.DocumentUri(pathToFileURI(filepath.Join(root, "Main.bas")))
			source := "Option Explicit\nSub Work()\nDim target As " + typeName + "\nDim source As " + typeName + "\ntarget = source\nEnd Sub\n"
			if _, err := s.docs.open(string(uri), source, 1); err != nil {
				t.Fatal(err)
			}
			result, err := s.codeAction(nil, &protocol.CodeActionParams{
				TextDocument: protocol.TextDocumentIdentifier{URI: uri},
				Range:        protocol.Range{Start: protocol.Position{Line: 4}, End: protocol.Position{Line: 4, Character: 15}},
				Context:      protocol.CodeActionContext{Only: []protocol.CodeActionKind{protocol.CodeActionKindQuickFix}},
			})
			if err != nil {
				t.Fatal(err)
			}
			for _, action := range result.([]protocol.CodeAction) {
				for _, diagnostic := range action.Diagnostics {
					if diagnostic.Code != nil && diagnostic.Code.Value == "VB036" {
						t.Fatalf("unexpected reference-intent correction: %+v", action)
					}
				}
			}
		})
	}
}
