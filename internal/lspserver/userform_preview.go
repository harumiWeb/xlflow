package lspserver

import (
	"context"
	"encoding/json"
	"errors"
	"strings"

	vbaintel "github.com/harumiWeb/xlflow/internal/vba/intel"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
	formsintel "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/intel"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

type userFormPreviewParams struct {
	PropertyEditorVersion int    `json:"propertyEditorVersion,omitzero"`
	URI                   string `json:"uri"`
	Version               int    `json:"version"`
	Text                  string `json:"text"`
}

type userFormPreviewError struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Line    int    `json:"line,omitzero"`
	Column  int    `json:"column,omitzero"`
}

type userFormPreviewResult struct {
	Version       int                         `json:"version"`
	Document      *forms.FormSpec             `json:"document,omitempty"`
	Warnings      []userFormPreviewWarning    `json:"warnings,omitempty"`
	Error         *userFormPreviewError       `json:"error,omitempty"`
	PropertyGrid  *formsedit.PropertyGridData `json:"propertyGrid,omitempty"`
	PropertyError *userFormPreviewError       `json:"propertyError,omitempty"`
}

type userFormPreviewWarning struct {
	forms.ValidationIssue
	Range *protocol.Range `json:"range,omitempty"`
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
	result.Document = &doc
	var syntax *formsintel.Document
	var jsonSyntax *formsintel.JSONDocument
	switch input.Format {
	case "yaml":
		syntax = formsintel.ParseYAML(params.Text)
	case "json":
		jsonSyntax = formsintel.ParseJSON(params.Text)
	}
	for _, issue := range doc.ValidationWarnings {
		warning := userFormPreviewWarning{ValidationIssue: issue}
		warning.Range = userFormPreviewIssueRange(syntax, jsonSyntax, issue)
		result.Warnings = append(result.Warnings, warning)
	}
	grid, err := formsedit.PropertyGrid(input, []byte(params.Text), params.PropertyEditorVersion)
	if err != nil {
		// A preview-valid source can be ambiguous for source-preserving edits.
		// Retain rendering, but do not publish guessed authoring values.
		result.PropertyError = &userFormPreviewError{Code: "propertyUnavailable", Message: err.Error()}
	} else {
		result.PropertyGrid = &grid
	}
	return result
}

func userFormPreviewIssueRange(yamlSyntax *formsintel.Document, jsonSyntax *formsintel.JSONDocument, issue forms.ValidationIssue) *protocol.Range {
	var r vbaintel.Range
	switch {
	case yamlSyntax != nil && yamlSyntax.ParseError == nil:
		r = userFormValidationIssueRange(yamlSyntax, issue)
	case jsonSyntax != nil && jsonSyntax.ParseError == nil:
		r = userFormJSONValidationIssueRange(jsonSyntax, issue)
	default:
		return nil
	}
	return &protocol.Range{
		Start: protocol.Position{Line: uint32(r.Start.Line), Character: uint32(r.Start.Character)},
		End:   protocol.Position{Line: uint32(r.End.Line), Character: uint32(r.End.Character)},
	}
}
func userFormJSONValidationIssueRange(syntax *formsintel.JSONDocument, issue forms.ValidationIssue) vbaintel.Range {
	if syntax == nil {
		return vbaintel.Range{}
	}
	if field, ok := syntax.Field(issue.Field); ok {
		if userFormIssueUsesKeyRange(issue.Code) || isEmptyRange(field.ValueRange) {
			return userFormPreviewIntelRange(field.KeyRange)
		}
		return userFormPreviewIntelRange(field.ValueRange)
	}
	return userFormJSONNearestFieldRange(syntax, issue.Field)
}

func userFormJSONNearestFieldRange(syntax *formsintel.JSONDocument, path string) vbaintel.Range {
	for current := path; ; current = userFormParentPath(current) {
		if field, ok := syntax.Field(current); ok {
			return userFormPreviewIntelRange(field.KeyRange)
		}
		if field, ok := userFormJSONFirstChildField(syntax, current); ok {
			return userFormPreviewIntelRange(field.KeyRange)
		}
		if current == "" {
			break
		}
	}
	if field, ok := userFormJSONFirstChildField(syntax, ""); ok {
		return userFormPreviewIntelRange(field.KeyRange)
	}
	return vbaintel.Range{}
}

func userFormPreviewIntelRange(r formsintel.Range) vbaintel.Range {
	return vbaintel.Range{
		Start: vbaintel.Position{Line: r.Start.Line, Character: r.Start.Character},
		End:   vbaintel.Position{Line: r.End.Line, Character: r.End.Character},
	}
}

func userFormJSONFirstChildField(syntax *formsintel.JSONDocument, parent string) (formsintel.JSONFieldNodes, bool) {
	prefix := parent
	if prefix != "" {
		prefix += "."
	}
	var (
		first formsintel.JSONFieldNodes
		found bool
	)
	for path, field := range syntax.Fields {
		rest := strings.TrimPrefix(path, prefix)
		if path == prefix || !strings.HasPrefix(path, prefix) || strings.Contains(rest, ".") || strings.Contains(rest, "[") {
			continue
		}
		if !found || positionBefore(field.KeyRange.Start, first.KeyRange.Start) {
			first, found = field, true
		}
	}
	return first, found
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
	if req.Method == userFormEventsMethod {
		s.dispatchUserFormEvents(ctx, conn, req)
		return true
	}
	if req.Method == userFormEventEditMethod {
		s.dispatchUserFormEventEdit(ctx, conn, req)
		return true
	}
	if req.Method == userFormPictureMethod {
		s.dispatchUserFormPicture(ctx, conn, req)
		return true
	}
	if req.Method == userFormCopyMethod {
		s.dispatchUserFormCopy(ctx, conn, req)
		return true
	}
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
