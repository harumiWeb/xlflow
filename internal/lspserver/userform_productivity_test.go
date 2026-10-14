package lspserver

import (
	"context"
	"encoding/json"
	"net"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestUserFormProductivityDecode(t *testing.T) {
	valid := []string{
		`{"type":"moveControls","moves":[{"controlId":"a","left":2,"top":3},{"controlId":"b","left":4,"top":3}]}`,
		`{"type":"alignControls","controlIds":["a","b"],"anchorId":"b","alignment":"left"}`,
		`{"type":"distributeControls","controlIds":["a","b","c"],"axis":"horizontal"}`,
		`{"type":"sameSizeControls","controlIds":["a","b"],"anchorId":"a","dimension":"size"}`,
		`{"type":"removeControls","controlIds":["a"],"cascade":false}`,
		`{"type":"reorderControls","controlIds":["a"],"command":"front"}`,
		`{"type":"setTabOrder","parentId":"","controlIds":[]}`,
		`{"type":"pasteControls","snapshot":{"format":"json","text":"{}","controlIds":["a"]},"allocations":[{"sourceId":"a","id":"new-a","name":"Label1"}],"parentId":"","offset":{"left":8,"top":8}}`,
	}
	for _, raw := range valid {
		t.Run(raw, func(t *testing.T) {
			op, e := parseUserFormEditOperation([]byte(raw))
			if e != nil {
				t.Fatal(e)
			}
			if e := validateUserFormEditTransaction([]formsedit.Operation{op}); e != nil {
				t.Fatal(e)
			}
			if e := validateUserFormEditTransaction([]formsedit.Operation{op, op}); e == nil {
				t.Fatal("arbitrary productivity batch accepted")
			}
		})
	}
	invalid := []string{
		`{"type":"removeControls","controlIds":["a"]}`,
		`{"type":"removeControls","controlIds":["a"],"cascade":null}`,
		`{"type":"reorderControls","controlIds":null,"command":"front"}`,
		`{"type":"moveControls","moves":[{"controlId":"a","top":0}]}`,
		`{"type":"moveControls","moves":[{"controlId":"a","left":0,"top":0,"width":4}]}`,
		`{"type":"moveControls","moves":[{"controlId":"a","left":0,"left":1,"top":0}]}`,
		`{"type":"moveControls","moves":[{"controlId":"a","left":1e999,"top":0}]}`,
		`{"type":"setTabOrder","controlIds":[]}`,
		`{"type":"pasteControls","snapshot":{"format":"json","text":"{}","controlIds":["a"]},"allocations":[{"sourceId":"a","id":"n"}],"parentId":"","offset":{"left":8,"top":8}}`,
		`{"type":"pasteControls","snapshot":{"format":"json","text":"{}","controlIds":["a"]},"allocations":[],"parentId":"","offset":{"left":8}}`,
	}
	for _, raw := range invalid {
		t.Run(raw, func(t *testing.T) {
			if _, e := parseUserFormEditOperation([]byte(raw)); e == nil {
				t.Fatal("malformed payload accepted")
			}
		})
	}
}

func TestUserFormProductivityRPC(t *testing.T) {
	ctx, cancel := context.WithTimeout(t.Context(), jsonrpcIntegrationTimeout)
	defer cancel()
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	serverSide, clientSide := net.Pipe()
	serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{handler: &s.handler, server: s, dispatch: s.dispatchRequest})
	defer func() { _ = serverConn.Close() }()
	clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), &rpcRecorder{})
	defer func() { _ = clientConn.Close() }()
	var initialized struct {
		Capabilities struct {
			Experimental map[string]any `json:"experimental"`
		} `json:"capabilities"`
	}
	if err := clientConn.Call(ctx, "initialize", protocol.InitializeParams{}, &initialized); err != nil {
		t.Fatal(err)
	}
	if initialized.Capabilities.Experimental["userFormProductivityEdit"] != true {
		t.Fatal("missing capability")
	}
	if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
		t.Fatal(err)
	}
	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.json"))
	// The file does not exist on disk: the unsaved request buffer is authoritative.
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main","width":240,"height":180},"controls":[{"id":"a","type":"Label","name":"Label1","left":10,"top":10,"width":30,"height":18,"caption":null,"visible":false,"tabIndex":0},{"id":"b","type":"Label","name":"Label2","left":80,"top":10,"width":30,"height":18}]}`
	var copied userFormCopyResult
	if err := clientConn.Call(ctx, userFormCopyMethod, userFormCopyParams{URI: uri, Version: 5, Text: source, ControlIDs: []string{"a", "b"}}, &copied); err != nil {
		t.Fatal(err)
	}
	if copied.Error != nil || copied.Version != 5 || copied.Snapshot == nil || copied.Document == nil || len(copied.Document.Controls) != 2 || copied.HasPicturePaths {
		t.Fatalf("copy response: %+v", copied)
	}
	if !strings.Contains(copied.Snapshot.Text, `"caption": null`) && !strings.Contains(copied.Snapshot.Text, `"caption":null`) {
		t.Fatalf("explicit null lost: %s", copied.Snapshot.Text)
	}
	raw, _ := json.Marshal(copied)
	if !strings.Contains(string(raw), `"hasPicturePaths":false`) {
		t.Fatal("false picture flag omitted")
	}
	var moved userFormEditResult
	params := map[string]any{"uri": uri, "version": 5, "text": source, "operations": []any{map[string]any{"type": "moveControls", "moves": []any{map[string]any{"controlId": "a", "left": 18, "top": 10}, map[string]any{"controlId": "b", "left": 88, "top": 10}}}}}
	if err := clientConn.Call(ctx, userFormEditMethod, params, &moved); err != nil {
		t.Fatal(err)
	}
	if moved.Error != nil || len(moved.Edits) == 0 {
		t.Fatalf("group move rejected: %+v", moved)
	}
	for _, ids := range [][]string{{"a", "a"}, {"missing"}, {"a", "missing"}} {
		r := s.userFormCopy(userFormCopyParams{URI: uri, Version: 5, Text: source, ControlIDs: ids})
		if r.Error == nil || r.Snapshot != nil {
			t.Fatalf("invalid selection copied: %+v", r)
		}
	}
	invalidCopy := []string{`{"uri":"x","version":0,"text":"","controlIds":null}`, `{"uri":"x","version":0,"text":"","controlIds":[],"extra":true}`, `{"uri":"x","version":-1,"text":"","controlIds":["a"]}`}
	for _, raw := range invalidCopy {
		if _, e := parseUserFormCopyParams([]byte(raw)); e == nil {
			t.Fatalf("bad copy payload accepted: %s", raw)
		}
	}
}

func TestUserFormProductivityScope(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.json"))
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main","width":400,"height":300},"controls":[{"id":"a","name":"Label1","type":"Label","left":10,"top":10,"width":20,"height":10},{"id":"frame","name":"Frame1","type":"Frame","left":50,"top":50,"width":200,"height":180},{"id":"child","name":"Label2","type":"Label","parentId":"frame","left":10,"top":10,"width":20,"height":10},{"id":"multi","name":"MultiPage1","type":"MultiPage","left":270,"top":20,"width":100,"height":100,"selectedIndex":0},{"id":"page","name":"Page1","type":"Page","parentId":"multi"}]}`
	for _, ids := range [][]string{{"a", "child"}, {"page"}, {"a", "a"}, {"missing"}} {
		op := formsedit.Operation{Type: formsedit.RemoveControls, ControlIDs: ids, Cascade: true}
		r := s.userFormEdit(userFormEditParams{URI: uri, Version: 7, Text: source, Operations: []formsedit.Operation{op}})
		if r.Error == nil || len(r.Edits) != 0 {
			t.Fatalf("unsupported deletion produced edits for %v: %+v", ids, r)
		}
		copy := s.userFormCopy(userFormCopyParams{URI: uri, Version: 7, Text: source, ControlIDs: ids})
		if copy.Error == nil || copy.Snapshot != nil {
			t.Fatalf("unsupported copy succeeded for %v: %+v", ids, copy)
		}
	}
	r := s.userFormEdit(userFormEditParams{URI: uri, Version: 7, Text: source, Operations: []formsedit.Operation{{Type: formsedit.SetTabOrder, ParentID: "multi", ControlIDs: []string{}}}})
	if r.Error == nil || len(r.Edits) != 0 {
		t.Fatalf("MultiPage tab-order parent accepted: %+v", r)
	}
	for _, id := range []string{"a", "frame", "multi", "child"} {
		copy := s.userFormCopy(userFormCopyParams{URI: uri, Version: 7, Text: source, ControlIDs: []string{id}})
		if copy.Error != nil || copy.Snapshot == nil {
			t.Fatalf("built-in subtree %s rejected: %+v", id, copy)
		}
	}
}
