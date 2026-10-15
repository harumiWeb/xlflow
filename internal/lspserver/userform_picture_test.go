package lspserver

import (
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"net"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestUserFormPictureRPC(t *testing.T) {
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
	params := userFormPictureParams{URI: pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml")), Data: "YQ=="}
	var result userFormPictureResult
	if err := clientConn.Call(ctx, userFormPictureMethod, params, &result); err == nil {
		t.Fatal("picture request accepted before initialize")
	}
	var initialized protocol.InitializeResult
	if err := clientConn.Call(ctx, "initialize", protocol.InitializeParams{}, &initialized); err != nil {
		t.Fatal(err)
	}
	if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
		t.Fatal(err)
	}
	if err := clientConn.Call(ctx, userFormPictureMethod, map[string]any{}, &result); err == nil {
		t.Fatal("empty picture request accepted")
	}
	if err := clientConn.Call(ctx, userFormPictureMethod, params, &result); err != nil || result.Error == nil || result.Error.Code != "invalidPicture" {
		t.Fatalf("invalid image RPC result: %+v, %v", result, err)
	}
	data, err := os.ReadFile("../vba/userforms/compiler/testdata/pictures-excel-authored/logo.bmp")
	if err != nil {
		t.Fatal(err)
	}
	params.Data = base64.StdEncoding.EncodeToString(data)
	result = userFormPictureResult{}
	if err := clientConn.Call(ctx, userFormPictureMethod, params, &result); err != nil || result.Error != nil || result.Format != "bmp" {
		t.Fatalf("valid image RPC result: %+v, %v", result, err)
	}
}

func TestUserFormPictureValidatesBytesAndProjectPathWithoutWrites(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Src.Forms = "custom/forms"
	s := &Server{opts: Options{RootDir: root, Config: cfg}}
	uri := pathToFileURI(filepath.Join(root, "custom/forms/specs/Main.yaml"))
	for _, format := range []string{"bmp", "jpg"} {
		data, err := os.ReadFile(filepath.Join("../vba/userforms/compiler/testdata/pictures-excel-authored", "logo."+format))
		if err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(root, "image."+format), data, 0o600); err != nil {
			t.Fatal(err)
		}
		fromBytes := s.userFormPicture(userFormPictureParams{URI: uri, Data: base64.StdEncoding.EncodeToString(data)})
		fromPath := s.userFormPicture(userFormPictureParams{URI: uri, Path: "image." + format})
		hash := sha256.Sum256(data)
		if fromBytes.Error != nil || fromPath.Error != nil || fromBytes.SHA256 != hex.EncodeToString(hash[:]) || fromPath.SHA256 != fromBytes.SHA256 || fromBytes.DataURI != fromPath.DataURI || fromBytes.Width <= 0 || fromBytes.Height <= 0 || fromBytes.Path != "" {
			t.Fatalf("validated picture mismatch: bytes=%+v path=%+v", fromBytes, fromPath)
		}
		if !strings.HasPrefix(fromBytes.DataURI, "data:image/"+fromBytes.Format+";base64,") {
			t.Fatal("incorrect preview MIME type")
		}
	}
	if _, err := os.Stat(filepath.Join(root, "custom")); !os.IsNotExist(err) {
		t.Fatalf("read-only validation created directories: %v", err)
	}
}

func TestUserFormPictureRejectsInvalidInputAndRootEscape(t *testing.T) {
	root := t.TempDir()
	s := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml"))
	for _, params := range []userFormPictureParams{
		{URI: uri},
		{URI: uri, Path: "image.bmp", Data: "YQ=="},
		{URI: uri, Path: "../image.bmp"},
		{URI: uri, Path: filepath.Join(root, "image.bmp")},
		{URI: uri, Data: "not base64!"},
		{URI: uri, Data: base64.StdEncoding.EncodeToString([]byte("not an image"))},
		{URI: pathToFileURI(filepath.Join(root, "other/Main.yaml")), Data: "YQ=="},
	} {
		if result := s.userFormPicture(params); result.Error == nil || result.DataURI != "" || result.SHA256 != "" {
			t.Fatalf("invalid request accepted: %+v", params)
		}
	}
	outside := t.TempDir()
	if err := os.WriteFile(filepath.Join(outside, "image.bmp"), []byte("secret"), 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(outside, filepath.Join(root, "linked")); err != nil {
		t.Logf("symlink unavailable: %v", err)
		return
	}
	if result := s.userFormPicture(userFormPictureParams{URI: uri, Path: "linked/image.bmp"}); result.Error == nil {
		t.Fatal("symlink escape accepted")
	}
}
