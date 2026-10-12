package lspserver

import (
	"context"
	"encoding/json"
	"math"
	"net"
	"path/filepath"
	"slices"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const userFormParentJSON = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"frameA","name":"FrameA","type":"Frame","controls":[{"id":"before","name":"Before","type":"Label"},{"id":"moving","name":"Moving","type":"Frame","left":4,"top":8,"controls":[{"id":"inside","name":"Inside","type":"Label"},{"id":"relinked","parentId":"frameA","name":"Relinked","type":"Label"}]},{"id":"remains","name":"Remains","type":"Label"}]},{"id":"frameB","name":"FrameB","type":"Frame"},{"id":"plain","name":"Plain","type":"Label"},{"id":"multi","name":"Multi","type":"MultiPage","selectedIndex":0},{"id":"page","name":"Page","type":"Page","parentId":"multi"},{"id":"custom","name":"Custom","type":"VendorControl","progId":"Vendor.Control"}]}`

type userFormParentRPC struct {
	ctx    context.Context
	client *jsonrpc2.Conn
	server *Server
	root   string
}

func newUserFormParentRPC(t *testing.T) userFormParentRPC {
	t.Helper()
	ctx, cancel := context.WithTimeout(t.Context(), jsonrpcIntegrationTimeout)
	t.Cleanup(cancel)
	root := t.TempDir()
	server, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(cleanup)
	serverSide, clientSide := net.Pipe()
	serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{handler: &server.handler, server: server, dispatch: server.dispatchRequest})
	t.Cleanup(func() { _ = serverConn.Close() })
	clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), &rpcRecorder{})
	t.Cleanup(func() { _ = clientConn.Close() })
	var initialized struct {
		Capabilities struct {
			Experimental map[string]any `json:"experimental"`
		} `json:"capabilities"`
	}
	if err := clientConn.Call(ctx, "initialize", protocol.InitializeParams{}, &initialized); err != nil {
		t.Fatal(err)
	}
	if initialized.Capabilities.Experimental["userFormParentEdit"] != true {
		t.Fatalf("parent edit capability was not advertised: %+v", initialized.Capabilities.Experimental)
	}
	if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
		t.Fatal(err)
	}
	return userFormParentRPC{ctx: ctx, client: clientConn, server: server, root: root}
}

func (rpc userFormParentRPC) call(t *testing.T, format, source string, operation json.RawMessage) userFormEditResult {
	t.Helper()
	uri := pathToFileURI(filepath.Join(rpc.root, "src", "forms", "specs", "Main."+format))
	params, err := json.Marshal(map[string]any{
		"uri": uri, "version": 23, "text": source,
		"operations": []json.RawMessage{operation},
	})
	if err != nil {
		t.Fatal(err)
	}
	var result userFormEditResult
	if err := rpc.client.Call(rpc.ctx, userFormEditMethod, json.RawMessage(params), &result); err != nil {
		t.Fatal(err)
	}
	return result
}

func TestUserFormParentEditRawRPCPreservesLegacySubtree(t *testing.T) {
	rpc := newUserFormParentRPC(t)
	if _, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(userFormParentJSON)); err != nil {
		t.Fatalf("invalid parenting fixture: %v", err)
	}
	operation := json.RawMessage(`{"type":"setParent","controlId":"moving","parentId":"frameB","left":16,"top":18}`)
	result := rpc.call(t, "json", userFormParentJSON, operation)
	if result.Error != nil || result.Version != 23 || len(result.Edits) == 0 {
		t.Fatalf("raw SetParent RPC failed: %+v", result)
	}
	updated, err := applyLSPTextEdits(userFormParentJSON, result.Edits)
	if err != nil {
		t.Fatal(err)
	}
	document, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(updated))
	if err != nil {
		t.Fatalf("edited JSON is invalid: %v\n%s", err, updated)
	}
	controls := make(map[string]spec.FormSpecControl, len(document.Controls))
	for _, control := range document.Controls {
		controls[control.ID] = control
	}
	moving := controls["moving"]
	if moving.ParentID != "frameB" || moving.Left == nil || *moving.Left != 16 || moving.Top == nil || *moving.Top != 18 || moving.ZIndex == nil || *moving.ZIndex != 1 {
		t.Fatalf("reparented control lost its relationship, coordinates or effective zIndex: %+v", moving)
	}
	if controls["inside"].ParentID != "moving" || controls["relinked"].ParentID != "frameA" || controls["before"].ParentID != "frameA" || controls["remains"].ParentID != "frameA" {
		t.Fatalf("relocation changed nested or canonical-unrelated relationships: %+v", controls)
	}
	noOp := rpc.call(t, "json", updated, operation)
	if noOp.Error != nil || len(noOp.Edits) != 0 {
		t.Fatalf("same-parent SetParent was not a no-op: %+v", noOp)
	}
	noOpSource, err := applyLSPTextEdits(updated, noOp.Edits)
	if err != nil || noOpSource != updated {
		t.Fatalf("no-op changed source bytes: err=%v", err)
	}

	rootOperation := json.RawMessage(`{"type":"setParent","controlId":"moving","parentId":"","left":16,"top":18}`)
	rootResult := rpc.call(t, "json", updated, rootOperation)
	if rootResult.Error != nil || len(rootResult.Edits) == 0 {
		t.Fatalf("empty parentId did not reparent to root: %+v", rootResult)
	}
	rootSource, err := applyLSPTextEdits(updated, rootResult.Edits)
	if err != nil {
		t.Fatal(err)
	}
	rootDocument, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(rootSource))
	if err != nil {
		t.Fatal(err)
	}
	if controlByID(t, rootDocument, "moving").ParentID != "" {
		t.Fatalf("empty parentId did not produce a root control: %+v", rootDocument.Controls)
	}
	rootToFrame := rpc.call(t, "json", rootSource, json.RawMessage(`{"type":"setParent","controlId":"plain","parentId":"frameB","left":9,"top":10}`))
	if rootToFrame.Error != nil || len(rootToFrame.Edits) == 0 {
		t.Fatalf("root-to-Frame SetParent failed: %+v", rootToFrame)
	}
	rootToFrameSource, err := applyLSPTextEdits(rootSource, rootToFrame.Edits)
	if err != nil {
		t.Fatal(err)
	}
	rootToFrameDocument, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(rootToFrameSource))
	if err != nil || controlByID(t, rootToFrameDocument, "plain").ParentID != "frameB" {
		t.Fatalf("root-to-Frame relation was not retained: controls=%+v err=%v", rootToFrameDocument.Controls, err)
	}
	frameToRoot := rpc.call(t, "json", rootToFrameSource, json.RawMessage(`{"type":"setParent","controlId":"plain","parentId":"","left":9,"top":10}`))
	if frameToRoot.Error != nil || len(frameToRoot.Edits) == 0 {
		t.Fatalf("Frame-to-root SetParent failed: %+v", frameToRoot)
	}
	frameToRootSource, err := applyLSPTextEdits(rootToFrameSource, frameToRoot.Edits)
	if err != nil {
		t.Fatal(err)
	}
	frameToRootDocument, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(frameToRootSource))
	if err != nil || controlByID(t, frameToRootDocument, "plain").ParentID != "" {
		t.Fatalf("Frame-to-root relation was not retained: controls=%+v err=%v", frameToRootDocument.Controls, err)
	}

	addFrameChild := rpc.call(t, "json", rootSource, json.RawMessage(`{"type":"addControl","control":{"id":"frameChild","name":"FrameChild","type":"Label","parentId":"frameB"}}`))
	if addFrameChild.Error != nil || len(addFrameChild.Edits) == 0 {
		t.Fatalf("raw Frame-parented insertion failed: %+v", addFrameChild)
	}
	addedSource, err := applyLSPTextEdits(rootSource, addFrameChild.Edits)
	if err != nil {
		t.Fatal(err)
	}
	addedDocument, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(addedSource))
	if err != nil || controlByID(t, addedDocument, "frameChild").ParentID != "frameB" {
		t.Fatalf("Frame parent was not preserved after addControl: controls=%+v err=%v", addedDocument.Controls, err)
	}
}

func TestUserFormParentEditCanonicalDiagnosticsAndDesignerScope(t *testing.T) {
	rpc := newUserFormParentRPC(t)
	tests := []struct {
		name        string
		operation   string
		code        string
		unsupported bool
	}{
		{name: "unknown destination remains canonical", operation: `{"type":"setParent","controlId":"plain","parentId":"missing","left":1,"top":2}`, code: "UFV008"},
		{name: "self parent remains canonical", operation: `{"type":"setParent","controlId":"plain","parentId":"plain","left":1,"top":2}`, code: "UFV009"},
		{name: "cycle remains canonical", operation: `{"type":"setParent","controlId":"frameA","parentId":"moving","left":1,"top":2}`, code: "UFV010"},
		{name: "nonFrame destination is outside designer scope", operation: `{"type":"setParent","controlId":"plain","parentId":"multi","left":1,"top":2}`, unsupported: true},
		{name: "Page subject is outside designer scope", operation: `{"type":"setParent","controlId":"page","parentId":"frameB","left":1,"top":2}`, unsupported: true},
		{name: "custom subject is outside designer scope", operation: `{"type":"setParent","controlId":"custom","parentId":"frameB","left":1,"top":2}`, unsupported: true},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result := rpc.call(t, "json", userFormParentJSON, json.RawMessage(test.operation))
			if result.Error == nil || len(result.Edits) != 0 {
				t.Fatalf("invalid/out-of-scope SetParent was not atomic: %+v", result)
			}
			if test.unsupported {
				if result.Error.Code != compiler.Unsupported {
					t.Fatalf("error code = %q, want unsupported: %+v", result.Error.Code, result.Error)
				}
				return
			}
			if !slices.ContainsFunc(result.Error.Diagnostics, func(d formsedit.Diagnostic) bool { return d.Code == test.code }) {
				t.Fatalf("canonical diagnostic %s was lost: %+v", test.code, result.Error)
			}
		})
	}
}

func TestUserFormParentEditRequiresFiniteCoordinatesAtEveryBoundary(t *testing.T) {
	rpc := newUserFormParentRPC(t)
	uri := pathToFileURI(filepath.Join(rpc.root, "src", "forms", "specs", "Main.json"))
	for _, test := range []struct {
		name string
		op   formsedit.Operation
	}{
		{name: "missing left", op: formsedit.Operation{Type: formsedit.SetParent, ControlID: "plain", ParentID: "", Top: new(2.0)}},
		{name: "missing top", op: formsedit.Operation{Type: formsedit.SetParent, ControlID: "plain", ParentID: "", Left: new(1.0)}},
		{name: "NaN left", op: formsedit.Operation{Type: formsedit.SetParent, ControlID: "plain", Left: new(math.NaN()), Top: new(2.0)}},
		{name: "infinite top", op: formsedit.Operation{Type: formsedit.SetParent, ControlID: "plain", Left: new(1.0), Top: new(math.Inf(1))}},
	} {
		t.Run(test.name, func(t *testing.T) {
			result := rpc.server.userFormEdit(userFormEditParams{URI: uri, Version: 24, Text: userFormParentJSON, Operations: []formsedit.Operation{test.op}})
			if result.Error == nil || result.Error.Code != compiler.Invalid || len(result.Edits) != 0 {
				t.Fatalf("typed handler accepted invalid SetParent coordinates: %+v", result)
			}
		})
	}

	for _, operation := range []string{
		`{"type":"setParent","controlId":"plain","parentId":""}`,
		`{"type":"setParent","controlId":"plain","parentId":"","left":1}`,
		`{"type":"setParent","controlId":"plain","parentId":"","left":1,"top":1,"extra":true}`,
		`{"type":"setParent","controlId":"plain","parentId":"","left":1e999,"top":1}`,
	} {
		result := rpc.call(t, "json", userFormParentJSON, json.RawMessage(operation))
		if result.Error == nil || result.Error.Code != compiler.Invalid || len(result.Edits) != 0 {
			t.Errorf("raw RPC accepted invalid SetParent payload %s: %+v", operation, result)
		}
	}
	parsed, parseError := parseUserFormEditOperation([]byte(`{"type":"setParent","controlId":"plain","parentId":"","left":1,"top":2}`))
	if parseError != nil || parsed.ParentID != "" || parsed.Left == nil || parsed.Top == nil {
		t.Fatalf("explicit empty parentId or finite coordinates were not parsed: %+v, %+v", parsed, parseError)
	}
}

func controlByID(t *testing.T, document spec.FormSpec, id string) spec.FormSpecControl {
	t.Helper()
	for _, control := range document.Controls {
		if control.ID == id {
			return control
		}
	}
	t.Fatalf("control %q not found", id)
	return spec.FormSpecControl{}
}
