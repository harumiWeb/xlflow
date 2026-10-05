package lspserver

import (
	"context"
	"encoding/json"
	"errors"

	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"github.com/sourcegraph/jsonrpc2"
)

type userFormPreviewParams struct {
	URI     string `json:"uri"`
	Version int    `json:"version"`
	Text    string `json:"text"`
}

type userFormPreviewError struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Line    int    `json:"line,omitzero"`
	Column  int    `json:"column,omitzero"`
}

type userFormPreviewResult struct {
	Version  int                     `json:"version"`
	Document *forms.FormSpec         `json:"document,omitempty"`
	Warnings []forms.ValidationIssue `json:"warnings,omitempty"`
	Error    *userFormPreviewError   `json:"error,omitempty"`
}

func (s *Server) userFormPreview(params userFormPreviewParams) userFormPreviewResult {
	result := userFormPreviewResult{Version: params.Version}
	input, previewError := s.userFormSpecInput(params.URI)
	if previewError != nil {
		result.Error = previewError
		return result
	}
	doc, err := forms.ParseFormSpec(input, []byte(params.Text))
	if err != nil {
		result.Error = &userFormPreviewError{Code: "invalidDocument", Message: err.Error()}
		if detail, ok := errors.AsType[*forms.SpecError](err); ok {
			result.Error.Code, result.Error.Line, result.Error.Column = detail.Code, detail.Line, detail.Column
		}
		return result
	}
	result.Document, result.Warnings = &doc, doc.ValidationWarnings
	return result
}

func (s *Server) userFormSpecInput(uri string) (forms.SpecInput, *userFormPreviewError) {
	path, err := fileURIToPath(uri)
	if err != nil || !isUserFormSpecPath(s.opts.RootDir, s.opts.Config.Src.Forms, path) {
		return forms.SpecInput{}, &userFormPreviewError{Code: "notFormSpec", Message: "Open a FormSpec directly under the configured forms root's specs directory."}
	}
	kind := DetectDocumentKind(s.opts.RootDir, s.opts.Config.Src.Forms, path, "")
	format := "yaml"
	if kind == DocumentKindUserFormJSON {
		format = "json"
	} else if kind != DocumentKindUserFormYAML {
		return forms.SpecInput{}, &userFormPreviewError{Code: "notFormSpec", Message: "Unsupported FormSpec extension."}
	}
	return forms.SpecInput{DisplayPath: path, Format: format}, nil
}

func (s *Server) dispatchRequest(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) bool {
	if req.Method == userFormEditMethod {
		s.dispatchUserFormEdit(ctx, conn, req)
		return true
	}
	if req.Method != "xlflow/userFormPreview" {
		return s.dispatchCodeAction(ctx, conn, req)
	}
	if req.Notif {
		return true
	}
	if !s.handler.IsInitialized() {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: -32002, Message: "server not initialized"})
		return true
	}
	var params userFormPreviewParams
	if req.Params == nil || json.Unmarshal(*req.Params, &params) != nil || params.URI == "" || params.Version < 0 {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInvalidParams, Message: "invalid FormSpec preview parameters"})
		return true
	}
	_ = conn.Reply(ctx, req.ID, s.userFormPreview(params))
	return true
}
