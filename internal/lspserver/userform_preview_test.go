package lspserver

import (
	"context"
	"net"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestUserFormPreviewRPC(t *testing.T) {
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
	// GLSP's ServerCapabilities.UnmarshalJSON drops experimental fields.
	// Check the actual wire contract without that lossy client-side decoder.
	var initialized struct {
		Capabilities struct {
			Experimental map[string]any `json:"experimental"`
		} `json:"capabilities"`
	}
	if err := clientConn.Call(ctx, "initialize", protocol.InitializeParams{}, &initialized); err != nil {
		t.Fatal(err)
	}
	if initialized.Capabilities.Experimental["userFormPreview"] != true {
		t.Fatalf("capability: %+v", initialized.Capabilities.Experimental)
	}
	if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
		t.Fatal(err)
	}
	var result userFormPreviewResult
	params := userFormPreviewParams{URI: pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml")), Version: 5, Text: previewSource}
	if err := clientConn.Call(ctx, "xlflow/userFormPreview", params, &result); err != nil {
		t.Fatal(err)
	}
	if result.Version != 5 || result.Document == nil || result.Error != nil {
		t.Fatalf("RPC preview: %+v", result)
	}
	if err := clientConn.Call(ctx, "xlflow/userFormPreview", map[string]any{}, &result); err == nil {
		t.Fatal("invalid params accepted")
	}
}

const previewSource = "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\n  width: 240\n  height: 180\ncontrols: []\n"

func TestUserFormPreviewUsesBufferAndConfiguredRoot(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Src.Forms = "custom/forms"
	s := &Server{opts: Options{RootDir: root, Config: cfg}}
	path := filepath.Join(root, "custom", "forms", "specs", "Main.yaml")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte("invalid disk source"), 0o644); err != nil {
		t.Fatal(err)
	}
	result := s.userFormPreview(userFormPreviewParams{URI: pathToFileURI(path), Version: 7, Text: previewSource})
	if result.Version != 7 || result.Error != nil || result.Document == nil || result.Document.Form.Name != "Main" {
		t.Fatalf("buffer preview: %+v", result)
	}
	body, err := os.ReadFile(path)
	if err != nil || string(body) != "invalid disk source" {
		t.Fatalf("source changed: %q, %v", body, err)
	}
	for _, relative := range []string{"Main.yaml", "src/forms/specs/Main.yaml", "custom/forms/specs/nested/Main.yaml", "custom/forms/specs/Main.txt"} {
		result := s.userFormPreview(userFormPreviewParams{URI: pathToFileURI(filepath.Join(root, relative)), Text: previewSource})
		if result.Error == nil || result.Error.Code != "notFormSpec" {
			t.Fatalf("accepted %s: %+v", relative, result)
		}
	}
}

func TestUserFormPreviewInvalidAndRecovery(t *testing.T) {
	root := t.TempDir()
	s := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml"))
	for _, source := range []string{"form: [", strings.Replace(previewSource, "xlflow.userform", "other", 1),
		strings.Replace(previewSource, "controls: []", "controls:\n  - id: a\n    type: Label\n    name: Label1\n    parentId: missing", 1)} {
		result := s.userFormPreview(userFormPreviewParams{URI: uri, Version: 2, Text: source})
		if result.Error == nil || result.Document != nil {
			t.Fatalf("invalid source accepted: %+v", result)
		}
	}
	result := s.userFormPreview(userFormPreviewParams{URI: uri, Version: 3, Text: previewSource})
	if result.Error != nil || result.Document == nil {
		t.Fatalf("recovery failed: %+v", result)
	}
}

func TestUserFormPreviewJSONAndWarnings(t *testing.T) {
	root := t.TempDir()
	s := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[]}`
	result := s.userFormPreview(userFormPreviewParams{URI: pathToFileURI(filepath.Join(root, "src/forms/specs/Main.json")), Version: 4, Text: source})
	if result.Error != nil || result.Document == nil {
		t.Fatalf("JSON: %+v", result)
	}
	source = strings.Replace(previewSource, "controls: []", "controls:\n  - id: custom\n    type: VendorWidget\n    name: Custom1\n    progId: Vendor.Widget.1", 1)
	result = s.userFormPreview(userFormPreviewParams{URI: pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yml")), Text: source})
	if result.Error != nil || result.Document == nil || len(result.Warnings) == 0 {
		t.Fatalf("warning preview: %+v", result)
	}
}
