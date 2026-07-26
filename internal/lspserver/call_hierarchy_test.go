package lspserver

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestCallHierarchyPrepareIncomingAndOutgoing(t *testing.T) {
	root := t.TempDir()
	writeCallHierarchyModule(t, root, "Main", `Public Sub Run()
    Target
    Target
End Sub
`)
	writeCallHierarchyModule(t, root, "Target", `Public Sub Target()
End Sub
`)

	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}

	targetURI := protocol.DocumentUri(pathToFileURI(filepath.Join(root, "src", "modules", "Target.bas")))
	prepared, err := s.prepareCallHierarchy(nil, &protocol.CallHierarchyPrepareParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: targetURI},
		Position:     protocol.Position{Line: 1, Character: 12},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if len(prepared) != 1 || prepared[0].Name != "Target" || prepared[0].Kind != protocol.SymbolKindFunction || prepared[0].Data == nil {
		t.Fatalf("prepared target = %+v", prepared)
	}
	target := roundTripCallHierarchyItem(t, prepared[0])

	incoming, err := s.callHierarchyIncomingCalls(nil, &protocol.CallHierarchyIncomingCallsParams{Item: target})
	if err != nil {
		t.Fatal(err)
	}
	if len(incoming) != 1 || incoming[0].From.Name != "Run" || len(incoming[0].FromRanges) != 2 {
		t.Fatalf("incoming = %+v", incoming)
	}
	if incoming[0].FromRanges[0].Start.Line >= incoming[0].FromRanges[1].Start.Line {
		t.Fatalf("incoming ranges not sorted: %+v", incoming[0].FromRanges)
	}

	mainURI := protocol.DocumentUri(pathToFileURI(filepath.Join(root, "src", "modules", "Main.bas")))
	mainPrepared, err := s.prepareCallHierarchy(nil, &protocol.CallHierarchyPrepareParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: mainURI},
		Position:     protocol.Position{Line: 2, Character: 5},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if len(mainPrepared) != 1 || mainPrepared[0].Name != "Run" {
		t.Fatalf("prepared caller = %+v", mainPrepared)
	}
	outgoing, err := s.callHierarchyOutgoingCalls(nil, &protocol.CallHierarchyOutgoingCallsParams{Item: roundTripCallHierarchyItem(t, mainPrepared[0])})
	if err != nil {
		t.Fatal(err)
	}
	if len(outgoing) != 1 || outgoing[0].To.Name != "Target" || len(outgoing[0].FromRanges) != 2 {
		t.Fatalf("outgoing = %+v", outgoing)
	}
}

func TestCallHierarchyRejectsInvalidAndStaleItems(t *testing.T) {
	root := t.TempDir()
	writeCallHierarchyModule(t, root, "Main", `Public Sub Run()
End Sub
`)
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}

	uri := protocol.DocumentUri(pathToFileURI(filepath.Join(root, "src", "modules", "Main.bas")))
	outside, err := s.prepareCallHierarchy(nil, &protocol.CallHierarchyPrepareParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: uri},
		Position:     protocol.Position{Line: 0, Character: 0},
	}})
	if err != nil || len(outside) != 0 {
		t.Fatalf("prepare outside procedure = %+v, %v", outside, err)
	}

	invalid := protocol.CallHierarchyItem{Data: map[string]any{"name": "Run"}}
	incoming, err := s.callHierarchyIncomingCalls(nil, &protocol.CallHierarchyIncomingCallsParams{Item: invalid})
	if err != nil || len(incoming) != 0 {
		t.Fatalf("incoming invalid item = %+v, %v", incoming, err)
	}
	outgoing, err := s.callHierarchyOutgoingCalls(nil, &protocol.CallHierarchyOutgoingCallsParams{Item: invalid})
	if err != nil || len(outgoing) != 0 {
		t.Fatalf("outgoing invalid item = %+v, %v", outgoing, err)
	}

	prepared, err := s.prepareCallHierarchy(nil, &protocol.CallHierarchyPrepareParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: uri},
		Position:     protocol.Position{Line: 1, Character: 12},
	}})
	if err != nil || len(prepared) != 1 {
		t.Fatalf("prepare = %+v, %v", prepared, err)
	}
	stale := roundTripCallHierarchyItem(t, prepared[0])
	if err := os.WriteFile(filepath.Join(root, "src", "modules", "Main.bas"), []byte("Attribute VB_Name = \"Main\"\nPublic Sub Changed()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := s.didChangeWatchedFiles(&glsp.Context{Notify: func(string, any) {}}, &protocol.DidChangeWatchedFilesParams{Changes: []protocol.FileEvent{{URI: uri, Type: protocol.FileChangeTypeChanged}}}); err != nil {
		t.Fatal(err)
	}
	incoming, err = s.callHierarchyIncomingCalls(nil, &protocol.CallHierarchyIncomingCallsParams{Item: stale})
	if err != nil || len(incoming) != 0 {
		t.Fatalf("incoming stale item = %+v, %v", incoming, err)
	}
}

func TestCallHierarchyIncludesUnambiguousPropertyAccessor(t *testing.T) {
	root := t.TempDir()
	writeCallHierarchyModule(t, root, "Caller", `Public Function ReadValue() As Long
    ReadValue = Target.Value(1)
End Function
`)
	writeCallHierarchyModule(t, root, "Target", `Public Property Get Value(ByVal index As Long) As Long
	    Value = index
End Property
`)
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}

	callerURI := protocol.DocumentUri(pathToFileURI(filepath.Join(root, "src", "modules", "Caller.bas")))
	prepared, err := s.prepareCallHierarchy(nil, &protocol.CallHierarchyPrepareParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: callerURI},
		Position:     protocol.Position{Line: 2, Character: 8},
	}})
	if err != nil || len(prepared) != 1 {
		t.Fatalf("prepare caller = %+v, %v", prepared, err)
	}
	outgoing, err := s.callHierarchyOutgoingCalls(nil, &protocol.CallHierarchyOutgoingCallsParams{Item: roundTripCallHierarchyItem(t, prepared[0])})
	if err != nil {
		t.Fatal(err)
	}
	if len(outgoing) != 1 || outgoing[0].To.Name != "Value" {
		t.Fatalf("property outgoing = %+v", outgoing)
	}
}

func TestCallHierarchyOmitsAmbiguousPropertyAccessor(t *testing.T) {
	root := t.TempDir()
	writeCallHierarchyModule(t, root, "Caller", `Public Function ReadValue() As Long
    ReadValue = Target.Value(1)
End Function
`)
	writeCallHierarchyModule(t, root, "Target", `Public Property Get Value(ByVal index As Long) As Long
    Value = index
End Property

Public Property Let Value(ByVal index As Long, ByVal newValue As Long)
End Property
`)
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}

	uri := protocol.DocumentUri(pathToFileURI(filepath.Join(root, "src", "modules", "Caller.bas")))
	prepared, err := s.prepareCallHierarchy(nil, &protocol.CallHierarchyPrepareParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: uri},
		Position:     protocol.Position{Line: 2, Character: 8},
	}})
	if err != nil || len(prepared) != 1 {
		t.Fatalf("prepare caller = %+v, %v", prepared, err)
	}
	outgoing, err := s.callHierarchyOutgoingCalls(nil, &protocol.CallHierarchyOutgoingCallsParams{Item: roundTripCallHierarchyItem(t, prepared[0])})
	if err != nil {
		t.Fatal(err)
	}
	if len(outgoing) != 0 {
		t.Fatalf("ambiguous property outgoing = %+v", outgoing)
	}
}

func TestInitializeAdvertisesCallHierarchy(t *testing.T) {
	s, cleanup, err := New(Options{RootDir: t.TempDir(), Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	result, err := s.initialize(nil, &protocol.InitializeParams{})
	if err != nil {
		t.Fatal(err)
	}
	capabilities := result.(protocol.InitializeResult).Capabilities
	if enabled, ok := capabilities.CallHierarchyProvider.(bool); !ok || !enabled {
		t.Fatalf("callHierarchyProvider = %#v", capabilities.CallHierarchyProvider)
	}
	if s.handler.TextDocumentPrepareCallHierarchy == nil || s.handler.CallHierarchyIncomingCalls == nil || s.handler.CallHierarchyOutgoingCalls == nil {
		t.Fatal("call hierarchy handlers were not registered")
	}
}

func writeCallHierarchyModule(t *testing.T, root, name, body string) {
	t.Helper()
	path := filepath.Join(root, "src", "modules", name+".bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	source := "Attribute VB_Name = \"" + name + "\"\n" + body
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
}

func roundTripCallHierarchyItem(t *testing.T, item protocol.CallHierarchyItem) protocol.CallHierarchyItem {
	t.Helper()
	encoded, err := json.Marshal(item)
	if err != nil {
		t.Fatal(err)
	}
	var roundTripped protocol.CallHierarchyItem
	if err := json.Unmarshal(encoded, &roundTripped); err != nil {
		t.Fatal(err)
	}
	return roundTripped
}
