package lspserver

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"reflect"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
	"github.com/sourcegraph/jsonrpc2"
)

const userFormCopyMethod = "xlflow/userFormCopy"

func isUserFormProductivityOperation(t formsedit.OperationType) bool {
	switch t {
	case formsedit.MoveControls, formsedit.AlignControls, formsedit.DistributeControls, formsedit.SameSizeControls, formsedit.RemoveControls, formsedit.ReorderControls, formsedit.SetTabOrder, formsedit.PasteControls:
		return true
	}
	return false
}

func parseUserFormProductivityOperation(raw []byte, fields map[string]json.RawMessage, t formsedit.OperationType) (formsedit.Operation, *userFormEditError) {
	required := []string{"type"}
	switch t {
	case formsedit.MoveControls:
		required = append(required, "moves")
	case formsedit.AlignControls:
		required = append(required, "controlIds", "anchorId", "alignment")
	case formsedit.DistributeControls:
		required = append(required, "controlIds", "axis")
	case formsedit.SameSizeControls:
		required = append(required, "controlIds", "anchorId", "dimension")
	case formsedit.RemoveControls:
		required = append(required, "controlIds", "cascade")
	case formsedit.ReorderControls:
		required = append(required, "controlIds", "command")
	case formsedit.SetTabOrder:
		required = append(required, "controlIds", "parentId")
	case formsedit.PasteControls:
		required = append(required, "snapshot", "allocations", "parentId", "offset")
	}
	if err := validateJSONFields(fields, required); err != nil {
		return formsedit.Operation{}, invalidUserFormEditParams(err)
	}
	if err := rejectDuplicateJSONFields(raw); err != nil {
		return formsedit.Operation{}, invalidUserFormEditParams(err)
	}
	if err := validateJSONStructPayload(raw, reflect.TypeFor[formsedit.Operation]()); err != nil {
		return formsedit.Operation{}, invalidUserFormEditParams(err)
	}
	var op formsedit.Operation
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.DisallowUnknownFields()
	if err := decoder.Decode(&op); err != nil {
		return formsedit.Operation{}, invalidUserFormEditParams(err)
	}
	// Required nested fields cannot silently become a Go zero value.
	check := func(raw json.RawMessage, keys []string) error {
		f, e := decodeJSONObject(raw)
		if e != nil {
			return e
		}
		return validateJSONFields(f, keys)
	}
	if t == formsedit.MoveControls {
		var moves []json.RawMessage
		_ = json.Unmarshal(fields["moves"], &moves)
		for _, m := range moves {
			if e := check(m, []string{"controlId", "left", "top"}); e != nil {
				return formsedit.Operation{}, invalidUserFormEditParams(e)
			}
		}
	}
	if t == formsedit.PasteControls {
		for key, keys := range map[string][]string{"snapshot": {"format", "text", "controlIds"}, "offset": {"left", "top"}} {
			if e := check(fields[key], keys); e != nil {
				return formsedit.Operation{}, invalidUserFormEditParams(e)
			}
		}
		var allocations []json.RawMessage
		_ = json.Unmarshal(fields["allocations"], &allocations)
		for _, a := range allocations {
			if e := check(a, []string{"sourceId", "id", "name"}); e != nil {
				return formsedit.Operation{}, invalidUserFormEditParams(e)
			}
		}
	}
	return op, nil
}

func validateDesignerSelection(document spec.FormSpec, ids []string) *userFormEditError {
	if len(ids) == 0 {
		return invalidUserFormEditParams(fmt.Errorf("controlIds must not be empty"))
	}
	byID := make(map[string]spec.FormSpecControl, len(document.Controls))
	for _, c := range document.Controls {
		byID[c.ID] = c
	}
	seen := map[string]bool{}
	parent := ""
	for i, id := range ids {
		c, ok := byID[id]
		if !ok || id == "" || seen[id] {
			return invalidUserFormEditParams(fmt.Errorf("controlIds must identify distinct existing controls"))
		}
		seen[id] = true
		if !isDesignerBuiltInControl(c) || isDesignerPage(c) {
			return &userFormEditError{Code: compiler.Unsupported, Message: "Designer selection does not support Page or custom controls"}
		}
		if i == 0 {
			parent = c.ParentID
		} else if parent != c.ParentID {
			return invalidUserFormEditParams(fmt.Errorf("selected controls must share one parent"))
		}
	}
	if parent != "" {
		p, ok := byID[parent]
		if !ok || (!isDesignerFrame(p) && !isDesignerPage(p)) {
			return &userFormEditError{Code: compiler.Unsupported, Message: "Designer selection requires a Form, Frame or Page parent"}
		}
	}
	return nil
}

func validateDesignerProductivity(input spec.SpecInput, source []byte, op formsedit.Operation) *userFormEditError {
	doc, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return nil
	}
	if op.Type == formsedit.SetTabOrder {
		return validateDesignerContainerParent(input, source, op.ParentID, "")
	}
	if op.Type == formsedit.PasteControls {
		if e := validateDesignerContainerParent(input, source, op.ParentID, ""); e != nil {
			return e
		}
		if op.Snapshot == nil {
			return invalidUserFormEditParams(fmt.Errorf("snapshot is required"))
		}
		snapshotInput := spec.SpecInput{Format: op.Snapshot.Format}
		_, copied, err := formsedit.CopyControls(snapshotInput, []byte(op.Snapshot.Text), op.Snapshot.ControlIDs)
		if err != nil {
			return userFormEditErrorFrom(err)
		}
		return validateDesignerCopiedTree(copied)
	}
	ids := op.ControlIDs
	if op.Type == formsedit.MoveControls {
		ids = make([]string, len(op.Moves))
		for i, m := range op.Moves {
			ids[i] = m.ControlID
		}
	}
	return validateDesignerSelection(doc, ids)
}

func validateDesignerCopiedTree(doc spec.FormSpec) *userFormEditError {
	for _, c := range doc.Controls {
		if !isDesignerBuiltInControl(c) {
			return &userFormEditError{Code: compiler.Unsupported, Message: "Designer clipboard does not support custom control subtrees"}
		}
	}
	return nil
}

type userFormCopyParams struct {
	URI        string   `json:"uri"`
	Version    int      `json:"version"`
	Text       string   `json:"text"`
	ControlIDs []string `json:"controlIds"`
}
type userFormCopyResult struct {
	Version         int                          `json:"version"`
	Snapshot        *formsedit.ClipboardSnapshot `json:"snapshot,omitempty"`
	Document        *spec.FormSpec               `json:"document,omitempty"`
	HasPicturePaths bool                         `json:"hasPicturePaths"`
	Error           *userFormEditError           `json:"error,omitempty"`
}

func parseUserFormCopyParams(raw []byte) (userFormCopyParams, *userFormEditError) {
	var p userFormCopyParams
	fields, err := decodeJSONObject(raw)
	if err != nil {
		return p, invalidUserFormEditParams(err)
	}
	if err = validateJSONFields(fields, []string{"uri", "version", "text", "controlIds"}); err != nil {
		return p, invalidUserFormEditParams(err)
	}
	if err = rejectDuplicateJSONFields(raw); err != nil {
		return p, invalidUserFormEditParams(err)
	}
	if err = validateJSONStructPayload(raw, reflect.TypeFor[userFormCopyParams]()); err != nil {
		return p, invalidUserFormEditParams(err)
	}
	if err = json.Unmarshal(raw, &p); err != nil {
		return p, invalidUserFormEditParams(err)
	}
	if p.URI == "" || p.Version < 0 || len(p.ControlIDs) == 0 {
		return p, invalidUserFormEditParams(fmt.Errorf("uri, nonnegative version and controlIds are required"))
	}
	return p, nil
}
func (s *Server) userFormCopy(p userFormCopyParams) userFormCopyResult {
	r := userFormCopyResult{Version: p.Version}
	input, pathError := s.userFormSpecInput(p.URI)
	if pathError != nil {
		r.Error = &userFormEditError{Code: compiler.Unsupported, Message: pathError.Message}
		return r
	}
	doc, err := spec.ParseFormSpec(input, []byte(p.Text))
	if err != nil {
		r.Error = userFormEditErrorFrom(err)
		return r
	}
	if e := validateDesignerSelection(doc, p.ControlIDs); e != nil {
		r.Error = e
		return r
	}
	snapshot, copied, err := formsedit.CopyControls(input, []byte(p.Text), p.ControlIDs)
	if err != nil {
		r.Error = userFormEditErrorFrom(err)
		return r
	}
	if e := validateDesignerCopiedTree(copied); e != nil {
		r.Error = e
		return r
	}
	r.Snapshot = &snapshot
	r.Document = &copied
	for _, c := range copied.Controls {
		if c.Picture != nil && c.Picture.Path != "" {
			r.HasPicturePaths = true
		}
	}
	return r
}
func (s *Server) dispatchUserFormCopy(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
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
	p, e := parseUserFormCopyParams(raw)
	if e != nil {
		_ = conn.Reply(ctx, req.ID, userFormCopyResult{Version: p.Version, Error: e})
		return
	}
	_ = conn.Reply(ctx, req.ID, s.userFormCopy(p))
}
