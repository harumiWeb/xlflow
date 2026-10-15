package lspserver

import (
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"path/filepath"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/picture"
	"github.com/sourcegraph/jsonrpc2"
)

const userFormPictureMethod = "xlflow/userFormPicture"

// The transport limit bounds decoding before the shared picture validator runs.
const userFormPictureMaxBytes = 16 << 20

type userFormPictureParams struct {
	URI  string `json:"uri"`
	Path string `json:"path,omitempty"`
	Data string `json:"data,omitempty"` // Base64 source bytes, never a filesystem path.
}

type userFormPictureResult struct {
	Path    string                `json:"path,omitempty"`
	Format  string                `json:"format,omitempty"`
	SHA256  string                `json:"sha256,omitempty"`
	DataURI string                `json:"dataURI,omitempty"`
	Width   int                   `json:"width,omitzero"`
	Height  int                   `json:"height,omitzero"`
	Error   *userFormPreviewError `json:"error,omitempty"`
}

// userFormPicture is read-only. External images arrive as bytes selected by the
// extension host; it never allows a webview to read an arbitrary absolute path.
func (s *Server) userFormPicture(params userFormPictureParams) userFormPictureResult {
	if _, err := s.userFormSpecInput(params.URI); err != nil {
		return userFormPictureResult{Error: err}
	}
	fail := func(err error) userFormPictureResult {
		return userFormPictureResult{Error: &userFormPreviewError{Code: "invalidPicture", Message: err.Error()}}
	}
	if (params.Path == "") == (params.Data == "") {
		return fail(fmt.Errorf("provide exactly one project-relative path or base64 image"))
	}
	var asset picture.Asset
	var err error
	if params.Path != "" {
		asset, err = picture.Load(s.opts.RootDir, params.Path)
	} else {
		if len(params.Data) > base64.StdEncoding.EncodedLen(userFormPictureMaxBytes) {
			return fail(fmt.Errorf("image exceeds the %d-byte limit", userFormPictureMaxBytes))
		}
		var data, raw []byte
		data, err = base64.StdEncoding.Strict().DecodeString(params.Data)
		if err == nil {
			raw, err = picture.Encode(data)
		}
		if err == nil {
			asset, err = picture.Decode(raw)
		}
	}
	if err != nil {
		return fail(err)
	}
	digest := sha256.Sum256(asset.Data)
	path := ""
	if params.Path != "" {
		path = filepath.ToSlash(filepath.Clean(params.Path))
	}
	return userFormPictureResult{
		Path: path, Format: asset.Format,
		SHA256: hex.EncodeToString(digest[:]), Width: asset.Width, Height: asset.Height,
		DataURI: "data:image/" + asset.Format + ";base64," + base64.StdEncoding.EncodeToString(asset.Data),
	}
}

// Wire this alongside the preview/edit dispatcher and advertise userFormPicture.
func (s *Server) dispatchUserFormPicture(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	if req.Notif {
		return
	}
	if !s.handler.IsInitialized() {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: -32002, Message: "server not initialized"})
		return
	}
	var params userFormPictureParams
	if req.Params == nil || json.Unmarshal(*req.Params, &params) != nil || params.URI == "" || (params.Path == "") == (params.Data == "") {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInvalidParams, Message: "invalid FormSpec picture parameters"})
		return
	}
	_ = conn.Reply(ctx, req.ID, s.userFormPicture(params))
}
