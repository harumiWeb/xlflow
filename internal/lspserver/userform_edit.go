package lspserver

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"math"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const userFormEditMethod = "xlflow/userFormEdit"

type userFormEditParams struct {
	URI        string
	Version    int
	Text       string
	Operations []formsedit.Operation
}

type userFormEditError struct {
	Code        string                 `json:"code"`
	Message     string                 `json:"message"`
	Diagnostics []formsedit.Diagnostic `json:"diagnostics,omitempty"`
}

type userFormEditResult struct {
	Version  int                    `json:"version"`
	Edits    []protocol.TextEdit    `json:"edits"`
	Warnings []spec.ValidationIssue `json:"warnings,omitempty"`
	Error    *userFormEditError     `json:"error,omitempty"`
}

func (s *Server) userFormEdit(params userFormEditParams) userFormEditResult {
	result := userFormEditResult{Version: params.Version, Edits: []protocol.TextEdit{}}
	input, pathError := s.userFormSpecInput(params.URI)
	if pathError != nil {
		result.Error = &userFormEditError{Code: compiler.Unsupported, Message: pathError.Message}
		return result
	}
	for index, operation := range params.Operations {
		if operation.Type != formsedit.MoveControl && operation.Type != formsedit.ResizeControl {
			result.Error = &userFormEditError{
				Code:    compiler.Unsupported,
				Message: fmt.Sprintf("operation %q is not exposed by the LSP edit method", operation.Type),
				Diagnostics: []formsedit.Diagnostic{{
					OperationIndex: index,
					Operation:      operation.Type,
					Code:           "UFE007",
					Message:        "The LSP edit method supports only moveControl and resizeControl.",
				}},
			}
			return result
		}
	}
	if requestError := validateUserFormEditTransaction(params.Operations); requestError != nil {
		result.Error = requestError
		return result
	}
	changed, err := formsedit.Apply(input, []byte(params.Text), params.Operations)
	if err != nil {
		result.Error = userFormEditErrorFrom(err)
		return result
	}
	result.Warnings = changed.Warnings
	for _, edit := range changed.Edits {
		result.Edits = append(result.Edits, protocol.TextEdit{
			Range: protocol.Range{
				Start: sourcePositionAtUTF16([]byte(params.Text), edit.Start),
				End:   sourcePositionAtUTF16([]byte(params.Text), edit.End),
			},
			NewText: edit.Text,
		})
	}
	return result
}

func userFormEditErrorFrom(err error) *userFormEditError {
	result := &userFormEditError{Code: compiler.Invalid, Message: err.Error()}
	if detail, ok := errors.AsType[*formsedit.Error](err); ok {
		result.Diagnostics = detail.Diagnostics
		if len(detail.Diagnostics) > 0 {
			switch detail.Diagnostics[0].Code {
			case "UFE001", "UFE005", "UFE006", "UFE007":
				result.Code = compiler.Unsupported
			case "UFE004", "UFE008":
				result.Code = compiler.Conflict
			}
		}
	}
	return result
}

func parseUserFormEditParams(raw []byte) (userFormEditParams, *userFormEditError) {
	var params userFormEditParams
	fields, err := decodeJSONObject(raw)
	if err != nil {
		return params, invalidUserFormEditParams(err)
	}
	for key := range fields {
		if key != "uri" && key != "version" && key != "text" && key != "operations" {
			return params, invalidUserFormEditParams(fmt.Errorf("unknown request field %q", key))
		}
	}
	uri, ok := fields["uri"]
	if !ok || json.Unmarshal(uri, &params.URI) != nil || bytes.Equal(bytes.TrimSpace(uri), []byte("null")) || params.URI == "" {
		return params, invalidUserFormEditParams(fmt.Errorf("uri must be a non-empty string"))
	}
	version, ok := fields["version"]
	if !ok || bytes.Equal(bytes.TrimSpace(version), []byte("null")) || json.Unmarshal(version, &params.Version) != nil || params.Version < 0 {
		return params, invalidUserFormEditParams(fmt.Errorf("version must be a non-negative integer"))
	}
	text, ok := fields["text"]
	if !ok || bytes.Equal(bytes.TrimSpace(text), []byte("null")) || json.Unmarshal(text, &params.Text) != nil {
		return params, invalidUserFormEditParams(fmt.Errorf("text must be a string"))
	}
	operations, ok := fields["operations"]
	if !ok || len(bytes.TrimSpace(operations)) == 0 || bytes.TrimSpace(operations)[0] != '[' {
		return params, invalidUserFormEditParams(fmt.Errorf("operations must be an array"))
	}
	var rawOperations []json.RawMessage
	if err := json.Unmarshal(operations, &rawOperations); err != nil {
		return params, invalidUserFormEditParams(err)
	}
	params.Operations = make([]formsedit.Operation, 0, len(rawOperations))
	for index, rawOperation := range rawOperations {
		operation, requestError := parseUserFormEditOperation(rawOperation)
		if requestError != nil {
			requestError.Message = fmt.Sprintf("operation %d: %s", index, requestError.Message)
			return params, requestError
		}
		params.Operations = append(params.Operations, operation)
	}
	if requestError := validateUserFormEditTransaction(params.Operations); requestError != nil {
		return params, requestError
	}
	return params, nil
}

func validateUserFormEditTransaction(operations []formsedit.Operation) *userFormEditError {
	if len(operations) < 1 || len(operations) > 2 {
		return invalidUserFormEditParams(fmt.Errorf("a transaction requires one geometry operation or one move/resize pair"))
	}
	if len(operations) == 2 && (operations[0].ControlID != operations[1].ControlID || operations[0].Type == operations[1].Type) {
		return invalidUserFormEditParams(fmt.Errorf("a move/resize pair must target one control with distinct operation types"))
	}
	return nil
}

func parseUserFormEditOperation(raw []byte) (formsedit.Operation, *userFormEditError) {
	fields, err := decodeJSONObject(raw)
	if err != nil {
		return formsedit.Operation{}, invalidUserFormEditParams(err)
	}
	typeJSON, ok := fields["type"]
	var operationType formsedit.OperationType
	if !ok || bytes.Equal(bytes.TrimSpace(typeJSON), []byte("null")) || json.Unmarshal(typeJSON, &operationType) != nil || operationType == "" {
		return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("type must be a non-empty string"))
	}
	var allowed []string
	switch operationType {
	case formsedit.MoveControl:
		allowed = []string{"type", "controlId", "left", "top"}
	case formsedit.ResizeControl:
		allowed = []string{"type", "controlId", "width", "height"}
	default:
		return formsedit.Operation{}, &userFormEditError{Code: compiler.Unsupported, Message: fmt.Sprintf("operation %q is not exposed by the LSP edit method", operationType)}
	}
	if err := validateJSONFields(fields, allowed); err != nil {
		return formsedit.Operation{}, invalidUserFormEditParams(err)
	}
	controlIDJSON, ok := fields["controlId"]
	var controlID string
	if !ok || bytes.Equal(bytes.TrimSpace(controlIDJSON), []byte("null")) || json.Unmarshal(controlIDJSON, &controlID) != nil || controlID == "" {
		return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("controlId must be a non-empty string"))
	}
	op := formsedit.Operation{Type: operationType, ControlID: controlID}
	if operationType == formsedit.MoveControl {
		left, err := decodeFiniteJSONNumber(fields["left"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("left: %w", err))
		}
		top, err := decodeFiniteJSONNumber(fields["top"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("top: %w", err))
		}
		op.Left, op.Top = new(left), new(top)
	} else {
		width, err := decodeFiniteJSONNumber(fields["width"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("width: %w", err))
		}
		height, err := decodeFiniteJSONNumber(fields["height"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("height: %w", err))
		}
		op.Width, op.Height = new(width), new(height)
	}
	return op, nil
}

func decodeJSONObject(raw []byte) (map[string]json.RawMessage, error) {
	decoder := json.NewDecoder(bytes.NewReader(raw))
	token, err := decoder.Token()
	if err != nil {
		return nil, err
	}
	delimiter, ok := token.(json.Delim)
	if !ok || delimiter != '{' {
		return nil, fmt.Errorf("expected an object")
	}
	fields := map[string]json.RawMessage{}
	for decoder.More() {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}
		key, ok := token.(string)
		if !ok {
			return nil, fmt.Errorf("object key must be a string")
		}
		if _, duplicate := fields[key]; duplicate {
			return nil, fmt.Errorf("duplicate field %q", key)
		}
		var value json.RawMessage
		if err := decoder.Decode(&value); err != nil {
			return nil, err
		}
		fields[key] = value
	}
	if _, err := decoder.Token(); err != nil {
		return nil, err
	}
	if err := decoder.Decode(new(any)); err != io.EOF {
		if err == nil {
			return nil, fmt.Errorf("unexpected trailing JSON value")
		}
		return nil, err
	}
	return fields, nil
}

func validateJSONFields(fields map[string]json.RawMessage, allowed []string) error {
	allowedSet := make(map[string]struct{}, len(allowed))
	for _, key := range allowed {
		allowedSet[key] = struct{}{}
	}
	for key := range fields {
		if _, ok := allowedSet[key]; !ok {
			return fmt.Errorf("unexpected payload field %q", key)
		}
	}
	for _, key := range allowed {
		if _, ok := fields[key]; !ok {
			return fmt.Errorf("missing payload field %q", key)
		}
	}
	return nil
}

func decodeFiniteJSONNumber(raw json.RawMessage) (float64, error) {
	if len(raw) == 0 || bytes.Equal(bytes.TrimSpace(raw), []byte("null")) {
		return 0, fmt.Errorf("a finite number is required")
	}
	var value float64
	if err := json.Unmarshal(raw, &value); err != nil {
		return 0, fmt.Errorf("a finite number is required")
	}
	if math.IsNaN(value) || math.IsInf(value, 0) {
		return 0, fmt.Errorf("a finite number is required")
	}
	return value, nil
}

func invalidUserFormEditParams(err error) *userFormEditError {
	return &userFormEditError{Code: compiler.Invalid, Message: "invalid UserForm edit request: " + err.Error()}
}

func sourcePositionAtUTF16(source []byte, offset int) protocol.Position {
	offset = max(0, min(offset, len(source)))
	line, character := 0, 0
	for index := 0; index < offset; {
		switch source[index] {
		case '\r':
			line++
			character = 0
			index++
			if index < offset && source[index] == '\n' {
				index++
			}
		case '\n':
			line++
			character = 0
			index++
		default:
			r, size := utf8.DecodeRune(source[index:])
			if size == 0 || index+size > offset {
				size = 1
			}
			character += len(utf16.Encode([]rune{r}))
			index += size
		}
	}
	return protocol.Position{Line: protocol.UInteger(line), Character: protocol.UInteger(character)}
}

func (s *Server) dispatchUserFormEdit(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	if req.Notif {
		return
	}
	if !s.handler.IsInitialized() {
		_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: -32002, Message: "server not initialized"})
		return
	}
	var raw []byte
	if req.Params != nil {
		raw = *req.Params
	}
	params, requestError := parseUserFormEditParams(raw)
	if requestError != nil {
		_ = conn.Reply(ctx, req.ID, userFormEditResult{Version: params.Version, Edits: []protocol.TextEdit{}, Error: requestError})
		return
	}
	_ = conn.Reply(ctx, req.ID, s.userFormEdit(params))
}
