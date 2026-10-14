package lspserver

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"math"
	"reflect"
	"strings"
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
		if !isUserFormEditOperation(operation.Type) {
			result.Error = &userFormEditError{
				Code:    compiler.Unsupported,
				Message: fmt.Sprintf("operation %q is not exposed by the LSP edit method", operation.Type),
				Diagnostics: []formsedit.Diagnostic{{
					OperationIndex: index,
					Operation:      operation.Type,
					Code:           "UFE007",
					Message:        "The LSP edit method supports geometry, scalar properties, control and Page structure, and tab collection operations.",
				}},
			}
			return result
		}
	}
	if requestError := validateUserFormEditTransaction(params.Operations); requestError != nil {
		result.Error = requestError
		return result
	}
	newMultiPageInsertion := isDesignerMultiPageInitialInsertionCandidate(params.Operations)
	for _, operation := range params.Operations {
		if isUserFormProductivityOperation(operation.Type) {
			if requestError := validateDesignerProductivity(input, []byte(params.Text), operation); requestError != nil {
				result.Error = requestError
				return result
			}
		}
		if operation.Type == formsedit.AddControl {
			if !newMultiPageInsertion {
				if requestError := validateDesignerControlAddition(operation.Control); requestError != nil {
					result.Error = requestError
					return result
				}
			}
			if requestError := validateDesignerContainerParent(input, []byte(params.Text), operation.Control.ParentID, operation.Control.ID); requestError != nil {
				result.Error = requestError
				return result
			}
		}
		if operation.Type == formsedit.SetParent {
			if requestError := validateDesignerSetParent(input, []byte(params.Text), operation); requestError != nil {
				result.Error = requestError
				return result
			}
		}
		if operation.Type == formsedit.AddPage {
			if !newMultiPageInsertion {
				if requestError := validateDesignerPageAddition(input, []byte(params.Text), operation.Control); requestError != nil {
					result.Error = requestError
					return result
				}
			}
		}
		if operation.Type == formsedit.RemovePage || operation.Type == formsedit.ReorderPage {
			if requestError := validateDesignerPageTarget(input, []byte(params.Text), operation.ControlID); requestError != nil {
				result.Error = requestError
				return result
			}
		}
		if operation.Type == formsedit.AddTab || operation.Type == formsedit.RemoveTab || operation.Type == formsedit.ReorderTab || operation.Type == formsedit.SetTabProperty {
			if requestError := validateDesignerTabTarget(input, []byte(params.Text), operation.ControlID); requestError != nil {
				result.Error = requestError
				return result
			}
		}
		if operation.Type == formsedit.RemoveControl {
			// Resolve legacy nested Pages through the canonical flattened model.
			// Invalid source remains Apply's responsibility so its structured
			// validation diagnostics and ambiguity checks are preserved.
			document, err := spec.ParseFormSpec(input, []byte(params.Text))
			if err == nil {
				for _, control := range document.Controls {
					if control.ID == operation.ControlID && strings.EqualFold(strings.TrimSpace(control.Type), "Page") {
						result.Error = &userFormEditError{Code: compiler.Unsupported, Message: "direct Page deletion is not exposed by the Designer edit method"}
						return result
					}
				}
			}
		}
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

func isUserFormEditOperation(operationType formsedit.OperationType) bool {
	if isUserFormProductivityOperation(operationType) {
		return true
	}
	switch operationType {
	case formsedit.MoveControl, formsedit.ResizeControl,
		formsedit.SetFormProperty, formsedit.SetControlProperty,
		formsedit.AddControl, formsedit.RemoveControl, formsedit.SetParent,
		formsedit.AddPage, formsedit.RemovePage, formsedit.ReorderPage,
		formsedit.AddTab, formsedit.RemoveTab, formsedit.ReorderTab, formsedit.SetTabProperty:
		return true
	default:
		return false
	}
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
	if len(operations) == 1 && isUserFormProductivityOperation(operations[0].Type) {
		return nil
	}
	if isDesignerMultiPageInitialInsertionCandidate(operations) {
		return validateDesignerMultiPageInitialInsertion(operations)
	}
	if len(operations) == 1 {
		switch operations[0].Type {
		case formsedit.SetFormProperty, formsedit.SetControlProperty,
			formsedit.MoveControl, formsedit.ResizeControl,
			formsedit.AddControl, formsedit.RemoveControl, formsedit.SetParent,
			formsedit.AddPage, formsedit.RemovePage, formsedit.ReorderPage,
			formsedit.AddTab, formsedit.RemoveTab, formsedit.ReorderTab, formsedit.SetTabProperty:
			return nil
		}
	}
	if len(operations) < 1 || len(operations) > 2 {
		return invalidUserFormEditParams(fmt.Errorf("a transaction requires one property, geometry or structural operation, one move/resize pair, or one root width/height resize pair"))
	}
	if len(operations) == 2 {
		if isUserFormResizePropertyPair(operations) {
			return nil
		}
		first, second := operations[0], operations[1]
		moveResize := first.Type == formsedit.MoveControl && second.Type == formsedit.ResizeControl
		resizeMove := first.Type == formsedit.ResizeControl && second.Type == formsedit.MoveControl
		if first.ControlID != second.ControlID || (!moveResize && !resizeMove) {
			return invalidUserFormEditParams(fmt.Errorf("a two-operation transaction must be a same-control move/resize pair or a root width/height property pair"))
		}
	}
	return nil
}

func isDesignerMultiPageInitialInsertionCandidate(operations []formsedit.Operation) bool {
	return len(operations) == 3 && operations[0].Type == formsedit.AddControl &&
		operations[1].Type == formsedit.AddPage && operations[2].Type == formsedit.AddPage
}

func validateDesignerMultiPageInitialInsertion(operations []formsedit.Operation) *userFormEditError {
	owner := operations[0].Control
	if owner == nil {
		return invalidUserFormEditParams(fmt.Errorf("initial MultiPage insertion requires an owner control"))
	}
	if requestError := validateDesignerControlAddition(owner); requestError != nil {
		return requestError
	}
	if !isDesignerMultiPage(*owner) {
		return &userFormEditError{Code: compiler.Unsupported, Message: "initial Page insertion requires a built-in MultiPage owner"}
	}
	if owner.ID == "" || strings.TrimSpace(owner.Name) == "" || owner.Tabs != nil ||
		owner.SelectedIndex == nil || *owner.SelectedIndex != -1 {
		return invalidUserFormEditParams(fmt.Errorf("initial MultiPage requires an id, name, selectedIndex -1 and no tabs"))
	}

	ids := map[string]struct{}{owner.ID: {}}
	names := []string{strings.TrimSpace(owner.Name)}
	for _, operation := range operations[1:] {
		page := operation.Control
		if requestError := validateDesignerPageAdditionToNewMultiPage(page, owner.ID); requestError != nil {
			return requestError
		}
		if _, exists := ids[page.ID]; exists {
			return invalidUserFormEditParams(fmt.Errorf("initial MultiPage controls must have distinct IDs"))
		}
		ids[page.ID] = struct{}{}
		name := strings.TrimSpace(page.Name)
		for _, existing := range names {
			if strings.EqualFold(existing, name) {
				return invalidUserFormEditParams(fmt.Errorf("initial MultiPage control names must be unique ignoring case"))
			}
		}
		names = append(names, name)
	}
	return nil
}

func isUserFormResizePropertyPair(operations []formsedit.Operation) bool {
	if len(operations) != 2 {
		return false
	}
	fields := map[string]bool{}
	for _, operation := range operations {
		if operation.Type != formsedit.SetFormProperty ||
			(operation.Field != "width" && operation.Field != "height") ||
			!operation.ValuePresent || !isFiniteUserFormEditNumber(operation.Value) || fields[operation.Field] {
			return false
		}
		fields[operation.Field] = true
	}
	return fields["width"] && fields["height"]
}

func isFiniteUserFormEditNumber(value any) bool {
	if number, ok := value.(json.Number); ok {
		_, err := decodeFiniteJSONNumber(json.RawMessage(number))
		return err == nil
	}
	number := reflect.ValueOf(value)
	if !number.IsValid() {
		return false
	}
	switch number.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		return true
	case reflect.Float32, reflect.Float64:
		return !math.IsNaN(number.Float()) && !math.IsInf(number.Float(), 0)
	default:
		return false
	}
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
	if isUserFormProductivityOperation(operationType) {
		return parseUserFormProductivityOperation(raw, fields, operationType)
	}
	switch operationType {
	case formsedit.MoveControl:
		allowed = []string{"type", "controlId", "left", "top"}
	case formsedit.ResizeControl:
		allowed = []string{"type", "controlId", "width", "height"}
	case formsedit.SetFormProperty:
		allowed = []string{"type", "field", "value"}
	case formsedit.SetControlProperty:
		allowed = []string{"type", "controlId", "field", "value"}
	case formsedit.AddControl:
		allowed = []string{"type", "control"}
	case formsedit.RemoveControl:
		allowed = []string{"type", "controlId", "cascade"}
	case formsedit.SetParent:
		allowed = []string{"type", "controlId", "parentId", "left", "top"}
	case formsedit.AddPage:
		allowed = []string{"type", "control"}
	case formsedit.RemovePage:
		allowed = []string{"type", "controlId", "cascade"}
	case formsedit.ReorderPage:
		allowed = []string{"type", "controlId", "index"}
	case formsedit.AddTab:
		allowed = []string{"type", "controlId", "tab"}
	case formsedit.RemoveTab:
		allowed = []string{"type", "controlId", "index"}
	case formsedit.ReorderTab:
		allowed = []string{"type", "controlId", "index", "toIndex"}
	case formsedit.SetTabProperty:
		allowed = []string{"type", "controlId", "index", "field", "value"}
	default:
		return formsedit.Operation{}, &userFormEditError{Code: compiler.Unsupported, Message: fmt.Sprintf("operation %q is not exposed by the LSP edit method", operationType)}
	}
	if operationType == formsedit.RemoveControl || operationType == formsedit.RemovePage {
		if err := validateJSONFieldSet(fields, []string{"type", "controlId"}, []string{"cascade"}); err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
	} else if err := validateJSONFields(fields, allowed); err != nil {
		return formsedit.Operation{}, invalidUserFormEditParams(err)
	}
	if operationType == formsedit.AddControl {
		control, err := decodeUserFormControl(fields["control"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("control: %w", err))
		}
		if requestError := validateDesignerControlAddition(control); requestError != nil {
			return formsedit.Operation{}, requestError
		}
		if isDesignerTabStrip(*control) && len(control.Tabs) != 0 {
			if requestError := validateDesignerInitialTabStripPayload(fields["control"]); requestError != nil {
				return formsedit.Operation{}, requestError
			}
		}
		return formsedit.Operation{Type: operationType, Control: control}, nil
	}
	if operationType == formsedit.AddPage {
		control, err := decodeUserFormPage(fields["control"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("control: %w", err))
		}
		return formsedit.Operation{Type: operationType, Control: control}, nil
	}
	if operationType == formsedit.AddTab {
		controlID, err := decodeRequiredJSONString(fields, "controlId")
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
		tab, err := decodeUserFormTab(fields["tab"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("tab: %w", err))
		}
		return formsedit.Operation{Type: operationType, ControlID: controlID, Tab: tab}, nil
	}
	if operationType == formsedit.RemoveControl || operationType == formsedit.RemovePage {
		controlID, err := decodeRequiredJSONString(fields, "controlId")
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
		operation := formsedit.Operation{Type: operationType, ControlID: controlID}
		if cascade, ok := fields["cascade"]; ok {
			if bytes.Equal(bytes.TrimSpace(cascade), []byte("null")) || json.Unmarshal(cascade, &operation.Cascade) != nil {
				return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("cascade must be a boolean"))
			}
		}
		return operation, nil
	}
	if operationType == formsedit.RemoveTab || operationType == formsedit.ReorderTab || operationType == formsedit.ReorderPage || operationType == formsedit.SetTabProperty {
		controlID, err := decodeRequiredJSONString(fields, "controlId")
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
		index, err := decodeRequiredJSONInt(fields, "index")
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
		op := formsedit.Operation{Type: operationType, ControlID: controlID, Index: new(index)}
		if operationType == formsedit.ReorderTab {
			toIndex, err := decodeRequiredJSONInt(fields, "toIndex")
			if err != nil {
				return formsedit.Operation{}, invalidUserFormEditParams(err)
			}
			op.ToIndex = new(toIndex)
		}
		if operationType == formsedit.SetTabProperty {
			field, err := decodeRequiredJSONString(fields, "field")
			if err != nil {
				return formsedit.Operation{}, invalidUserFormEditParams(err)
			}
			value, err := decodeJSONScalar(fields, "value")
			if err != nil {
				return formsedit.Operation{}, invalidUserFormEditParams(err)
			}
			op.Field, op.Value, op.ValuePresent = field, value, true
		}
		return op, nil
	}
	if operationType == formsedit.SetParent {
		controlID, err := decodeRequiredJSONString(fields, "controlId")
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
		parentIDJSON, ok := fields["parentId"]
		var parentID string
		if !ok || bytes.Equal(bytes.TrimSpace(parentIDJSON), []byte("null")) || json.Unmarshal(parentIDJSON, &parentID) != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("parentId must be a string"))
		}
		left, err := decodeFiniteJSONNumber(fields["left"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("left: %w", err))
		}
		top, err := decodeFiniteJSONNumber(fields["top"])
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("top: %w", err))
		}
		return formsedit.Operation{Type: operationType, ControlID: controlID, ParentID: parentID, Left: new(left), Top: new(top)}, nil
	}
	if operationType == formsedit.SetFormProperty || operationType == formsedit.SetControlProperty {
		field, err := decodeRequiredJSONString(fields, "field")
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
		value, err := decodeJSONScalar(fields, "value")
		if err != nil {
			return formsedit.Operation{}, invalidUserFormEditParams(err)
		}
		op := formsedit.Operation{Type: operationType, Field: field, Value: value, ValuePresent: true}
		if operationType == formsedit.SetControlProperty {
			if bytes.Equal(bytes.TrimSpace(fields["controlId"]), []byte("null")) || json.Unmarshal(fields["controlId"], &op.ControlID) != nil || op.ControlID == "" {
				return formsedit.Operation{}, invalidUserFormEditParams(fmt.Errorf("controlId must be a non-empty string"))
			}
		}
		return op, nil
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

func validateDesignerControlAddition(control *spec.FormSpecControl) *userFormEditError {
	if control == nil {
		return invalidUserFormEditParams(fmt.Errorf("control must be an object"))
	}
	contract, known := spec.LookupControlContract(control.Type)
	if len(control.Controls) != 0 || !known || contract.Type == "Page" {
		return &userFormEditError{Code: compiler.Unsupported, Message: "Designer insertion supports only flat built-in controls other than Page"}
	}
	if control.ProgID != "" && !strings.EqualFold(control.ProgID, contract.ProgID) {
		return &userFormEditError{Code: compiler.Unsupported, Message: "Designer insertion does not support custom ProgIDs"}
	}
	if strings.EqualFold(contract.Type, "TabStrip") && len(control.Tabs) != 0 {
		if requestError := validateDesignerInitialTabStripAddition(control); requestError != nil {
			return requestError
		}
	}
	return nil
}

func validateDesignerInitialTabStripAddition(control *spec.FormSpecControl) *userFormEditError {
	if control == nil || !isDesignerTabStrip(*control) || len(control.Tabs) != 2 ||
		control.SelectedIndex == nil || *control.SelectedIndex != 0 {
		return &userFormEditError{Code: compiler.Unsupported, Message: "Designer TabStrip insertion supports only two initial tabs with selectedIndex 0"}
	}
	seenNames := make([]string, 0, len(control.Tabs))
	for _, tab := range control.Tabs {
		name := strings.TrimSpace(tab.Name)
		if name == "" || tab.Caption == nil || tab.ControlTipText != nil || tab.Tag != nil ||
			tab.Accelerator != nil || tab.Enabled != nil || tab.Visible != nil {
			return invalidUserFormEditParams(fmt.Errorf("initial TabStrip tabs require only non-empty name and caption fields"))
		}
		for _, seen := range seenNames {
			if strings.EqualFold(seen, name) {
				return invalidUserFormEditParams(fmt.Errorf("initial TabStrip tab names must be unique ignoring case"))
			}
		}
		seenNames = append(seenNames, name)
	}
	return nil
}

func validateDesignerInitialTabStripPayload(raw json.RawMessage) *userFormEditError {
	fields, err := decodeJSONObject(raw)
	if err != nil {
		return invalidUserFormEditParams(err)
	}
	tabsJSON, ok := fields["tabs"]
	if !ok {
		return invalidUserFormEditParams(fmt.Errorf("initial TabStrip tabs must be explicitly provided"))
	}
	var tabs []json.RawMessage
	if err := json.Unmarshal(tabsJSON, &tabs); err != nil {
		return invalidUserFormEditParams(fmt.Errorf("tabs: %w", err))
	}
	for _, rawTab := range tabs {
		tabFields, err := decodeJSONObject(rawTab)
		if err != nil {
			return invalidUserFormEditParams(fmt.Errorf("tab: %w", err))
		}
		if err := validateJSONFieldSet(tabFields, []string{"name", "caption"}, nil); err != nil {
			return invalidUserFormEditParams(fmt.Errorf("tab: %w", err))
		}
		captionJSON, ok := tabFields["caption"]
		var caption string
		if !ok || bytes.Equal(bytes.TrimSpace(captionJSON), []byte("null")) || json.Unmarshal(captionJSON, &caption) != nil {
			return invalidUserFormEditParams(fmt.Errorf("tab.caption must be a string"))
		}
	}
	return nil
}

func validateDesignerSetParent(input spec.SpecInput, source []byte, operation formsedit.Operation) *userFormEditError {
	if operation.Left == nil || operation.Top == nil || math.IsNaN(*operation.Left) || math.IsInf(*operation.Left, 0) || math.IsNaN(*operation.Top) || math.IsInf(*operation.Top, 0) {
		return invalidUserFormEditParams(fmt.Errorf("setParent requires finite left and top coordinates"))
	}
	document, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return nil
	}
	var subject *spec.FormSpecControl
	controlsByID := make(map[string]spec.FormSpecControl, len(document.Controls))
	for index := range document.Controls {
		control := &document.Controls[index]
		controlsByID[control.ID] = *control
		if control.ID == operation.ControlID {
			subject = control
		}
	}
	if subject == nil {
		return nil
	}
	if strings.EqualFold(strings.TrimSpace(subject.Type), "Page") || !isDesignerBuiltInControl(*subject) {
		return &userFormEditError{Code: compiler.Unsupported, Message: "Designer reparenting does not support Page or custom controls"}
	}
	if subject.ParentID != "" {
		parent, ok := controlsByID[subject.ParentID]
		if ok && !isDesignerFrame(parent) && !isDesignerPage(parent) {
			return &userFormEditError{Code: compiler.Unsupported, Message: "Designer reparenting supports controls whose current parent is root, Frame or Page"}
		}
	}
	return validateDesignerContainerParent(input, source, operation.ParentID, operation.ControlID)
}

func validateDesignerContainerParent(input spec.SpecInput, source []byte, parentID, subjectID string) *userFormEditError {
	if parentID == "" || parentID == subjectID {
		return nil
	}
	document, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return nil
	}
	for _, control := range document.Controls {
		if control.ID == parentID && !isDesignerFrame(control) && !isDesignerPage(control) {
			return &userFormEditError{Code: compiler.Unsupported, Message: "Designer reparenting supports only Frame or Page destinations"}
		}
	}
	return nil
}

func validateDesignerPageAddition(input spec.SpecInput, source []byte, control *spec.FormSpecControl) *userFormEditError {
	if requestError := validateDesignerPageAdditionShape(control); requestError != nil {
		return requestError
	}
	return validateDesignerMultiPageParent(input, source, control.ParentID)
}

func validateDesignerPageAdditionToNewMultiPage(control *spec.FormSpecControl, parentID string) *userFormEditError {
	if requestError := validateDesignerPageAdditionShape(control); requestError != nil {
		return requestError
	}
	if control.ParentID != parentID {
		return invalidUserFormEditParams(fmt.Errorf("addPage parentId must reference the MultiPage created in this transaction"))
	}
	return nil
}

func validateDesignerPageAdditionShape(control *spec.FormSpecControl) *userFormEditError {
	if control == nil {
		return invalidUserFormEditParams(fmt.Errorf("control must be an object"))
	}
	if !isDesignerPage(*control) {
		return &userFormEditError{Code: compiler.Unsupported, Message: "Designer page insertion supports only the built-in Page control"}
	}
	if control.ID == "" || control.Name == "" || control.ParentID == "" || len(control.Controls) != 0 || len(control.Tabs) != 0 || control.Left != nil || control.Top != nil || control.Width != nil || control.Height != nil {
		return invalidUserFormEditParams(fmt.Errorf("addPage requires a flat Page with id, name and parentId and no geometry"))
	}
	return nil
}

func validateDesignerPageTarget(input spec.SpecInput, source []byte, pageID string) *userFormEditError {
	document, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return nil
	}
	controlsByID := make(map[string]spec.FormSpecControl, len(document.Controls))
	for _, control := range document.Controls {
		controlsByID[control.ID] = control
	}
	page, ok := controlsByID[pageID]
	if !ok {
		return nil
	}
	if !isDesignerPage(page) {
		return &userFormEditError{Code: compiler.Unsupported, Message: "Designer page operations support only built-in Page controls"}
	}
	parent, ok := controlsByID[page.ParentID]
	if !ok {
		return nil
	}
	if !isDesignerMultiPage(parent) {
		return &userFormEditError{Code: compiler.Unsupported, Message: "Designer page operations require a built-in MultiPage parent"}
	}
	return nil
}

func validateDesignerMultiPageParent(input spec.SpecInput, source []byte, parentID string) *userFormEditError {
	document, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return nil
	}
	for _, control := range document.Controls {
		if control.ID == parentID && !isDesignerMultiPage(control) {
			return &userFormEditError{Code: compiler.Unsupported, Message: "Designer pages require a built-in MultiPage parent"}
		}
	}
	return nil
}

func validateDesignerTabTarget(input spec.SpecInput, source []byte, controlID string) *userFormEditError {
	document, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return nil
	}
	for _, control := range document.Controls {
		if control.ID == controlID && !isDesignerTabStrip(control) {
			return &userFormEditError{Code: compiler.Unsupported, Message: "Designer tab operations support only the built-in TabStrip control"}
		}
	}
	return nil
}

func isDesignerBuiltInControl(control spec.FormSpecControl) bool {
	contract, ok := spec.LookupControlContract(control.Type)
	if !ok {
		return false
	}
	return control.ProgID == "" || strings.EqualFold(control.ProgID, contract.ProgID)
}

func isDesignerFrame(control spec.FormSpecControl) bool {
	contract, ok := spec.LookupControlContract(control.Type)
	return ok && strings.EqualFold(contract.Type, "Frame") && isDesignerBuiltInControl(control)
}

func isDesignerPage(control spec.FormSpecControl) bool {
	contract, ok := spec.LookupControlContract(control.Type)
	return ok && strings.EqualFold(contract.Type, "Page") && isDesignerBuiltInControl(control)
}

func isDesignerMultiPage(control spec.FormSpecControl) bool {
	contract, ok := spec.LookupControlContract(control.Type)
	return ok && strings.EqualFold(contract.Type, "MultiPage") && isDesignerBuiltInControl(control)
}

func isDesignerTabStrip(control spec.FormSpecControl) bool {
	contract, ok := spec.LookupControlContract(control.Type)
	return ok && strings.EqualFold(contract.Type, "TabStrip") && isDesignerBuiltInControl(control)
}

func decodeUserFormPage(raw json.RawMessage) (*spec.FormSpecControl, error) {
	fields, err := decodeJSONObject(raw)
	if err != nil {
		return nil, err
	}
	if err := rejectDuplicateJSONFields(raw); err != nil {
		return nil, err
	}
	if err := validateJSONFieldSet(fields, []string{"id", "name", "type", "parentId"}, []string{"caption"}); err != nil {
		return nil, err
	}
	for _, key := range []string{"id", "name", "parentId"} {
		if _, err := decodeRequiredJSONString(fields, key); err != nil {
			return nil, err
		}
	}
	typeName, err := decodeRequiredJSONString(fields, "type")
	if err != nil {
		return nil, err
	}
	if !strings.EqualFold(typeName, "Page") {
		return nil, fmt.Errorf("type must be Page")
	}
	if err := validateJSONStructPayload(raw, reflect.TypeFor[spec.FormSpecControl]()); err != nil {
		return nil, err
	}
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.UseNumber()
	decoder.DisallowUnknownFields()
	var control spec.FormSpecControl
	if err := decoder.Decode(&control); err != nil {
		return nil, err
	}
	return &control, nil
}

func decodeUserFormTab(raw json.RawMessage) (*spec.FormSpecTab, error) {
	fields, err := decodeJSONObject(raw)
	if err != nil {
		return nil, err
	}
	if err := rejectDuplicateJSONFields(raw); err != nil {
		return nil, err
	}
	if err := validateJSONFieldSet(fields, []string{"name"}, []string{"caption", "controlTipText", "tag", "accelerator", "enabled", "visible"}); err != nil {
		return nil, err
	}
	if _, err := decodeRequiredJSONString(fields, "name"); err != nil {
		return nil, err
	}
	if err := validateJSONStructPayload(raw, reflect.TypeFor[spec.FormSpecTab]()); err != nil {
		return nil, err
	}
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.DisallowUnknownFields()
	var tab spec.FormSpecTab
	if err := decoder.Decode(&tab); err != nil {
		return nil, err
	}
	return &tab, nil
}

func decodeUserFormControl(raw json.RawMessage) (*spec.FormSpecControl, error) {
	fields, err := decodeJSONObject(raw)
	if err != nil {
		return nil, err
	}
	if err := rejectDuplicateJSONFields(raw); err != nil {
		return nil, err
	}
	if err := validateJSONStructPayload(raw, reflect.TypeFor[spec.FormSpecControl]()); err != nil {
		return nil, err
	}
	for _, key := range []string{"id", "name", "type"} {
		if _, err := decodeRequiredJSONString(fields, key); err != nil {
			return nil, err
		}
	}
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.UseNumber()
	decoder.DisallowUnknownFields()
	var control spec.FormSpecControl
	if err := decoder.Decode(&control); err != nil {
		return nil, err
	}
	if len(control.Controls) > 0 {
		return nil, fmt.Errorf("must contain exactly one flat control")
	}
	return &control, nil
}

func validateJSONStructPayload(raw json.RawMessage, valueType reflect.Type) error {
	if bytes.Equal(bytes.TrimSpace(raw), []byte("null")) {
		return fmt.Errorf("null cannot be represented in the control payload")
	}
	for valueType.Kind() == reflect.Pointer {
		valueType = valueType.Elem()
	}
	switch valueType.Kind() {
	case reflect.Struct:
		fields, err := decodeJSONObject(raw)
		if err != nil {
			return err
		}
		known := make(map[string]reflect.Type, valueType.NumField())
		for index := range valueType.NumField() {
			field := valueType.Field(index)
			if !field.IsExported() {
				continue
			}
			name, _, _ := strings.Cut(field.Tag.Get("json"), ",")
			if name == "-" {
				continue
			}
			if name == "" {
				name = field.Name
			}
			known[name] = field.Type
		}
		for name, fieldValue := range fields {
			fieldType, ok := known[name]
			if !ok {
				return fmt.Errorf("unknown or noncanonical field %q", name)
			}
			if err := validateJSONStructPayload(fieldValue, fieldType); err != nil {
				return fmt.Errorf("%s: %w", name, err)
			}
		}
	case reflect.Slice, reflect.Array:
		var values []json.RawMessage
		if err := json.Unmarshal(raw, &values); err != nil {
			return err
		}
		for index, item := range values {
			if err := validateJSONStructPayload(item, valueType.Elem()); err != nil {
				return fmt.Errorf("[%d]: %w", index, err)
			}
		}
	}
	return nil
}

func decodeRequiredJSONString(fields map[string]json.RawMessage, key string) (string, error) {
	raw, ok := fields[key]
	var value string
	if !ok || bytes.Equal(bytes.TrimSpace(raw), []byte("null")) || json.Unmarshal(raw, &value) != nil || value == "" {
		return "", fmt.Errorf("%s must be a non-empty string", key)
	}
	return value, nil
}

func decodeRequiredJSONInt(fields map[string]json.RawMessage, key string) (int, error) {
	raw, ok := fields[key]
	var value int
	if !ok || bytes.Equal(bytes.TrimSpace(raw), []byte("null")) || json.Unmarshal(raw, &value) != nil {
		return 0, fmt.Errorf("%s must be an integer", key)
	}
	return value, nil
}

func decodeJSONScalar(fields map[string]json.RawMessage, key string) (any, error) {
	raw, ok := fields[key]
	if !ok {
		return nil, fmt.Errorf("missing payload field %q", key)
	}
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.UseNumber()
	var value any
	if err := decoder.Decode(&value); err != nil {
		return nil, err
	}
	switch scalar := value.(type) {
	case nil, string, bool:
	case json.Number:
		if _, err := decodeFiniteJSONNumber([]byte(scalar)); err != nil {
			return nil, err
		}
	default:
		return nil, fmt.Errorf("%s must be a scalar or null", key)
	}
	return value, nil
}

func rejectDuplicateJSONFields(raw []byte) error {
	decoder := json.NewDecoder(bytes.NewReader(raw))
	if err := consumeJSONValue(decoder); err != nil {
		return err
	}
	if _, err := decoder.Token(); err != io.EOF {
		if err == nil {
			return fmt.Errorf("unexpected trailing JSON value")
		}
		return err
	}
	return nil
}

func consumeJSONValue(decoder *json.Decoder) error {
	token, err := decoder.Token()
	if err != nil {
		return err
	}
	delimiter, ok := token.(json.Delim)
	if !ok {
		return nil
	}
	switch delimiter {
	case '{':
		seen := map[string]struct{}{}
		for decoder.More() {
			keyToken, err := decoder.Token()
			if err != nil {
				return err
			}
			key, ok := keyToken.(string)
			if !ok {
				return fmt.Errorf("object key must be a string")
			}
			if _, duplicate := seen[key]; duplicate {
				return fmt.Errorf("duplicate field %q", key)
			}
			seen[key] = struct{}{}
			if err := consumeJSONValue(decoder); err != nil {
				return err
			}
		}
		_, err = decoder.Token()
		return err
	case '[':
		for decoder.More() {
			if err := consumeJSONValue(decoder); err != nil {
				return err
			}
		}
		_, err = decoder.Token()
		return err
	default:
		return fmt.Errorf("unexpected JSON delimiter %q", delimiter)
	}
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
	return validateJSONFieldSet(fields, allowed, nil)
}

func validateJSONFieldSet(fields map[string]json.RawMessage, required, optional []string) error {
	allowedSet := make(map[string]struct{}, len(required)+len(optional))
	for _, key := range required {
		allowedSet[key] = struct{}{}
	}
	for _, key := range optional {
		allowedSet[key] = struct{}{}
	}
	for key := range fields {
		if _, ok := allowedSet[key]; !ok {
			return fmt.Errorf("unexpected payload field %q", key)
		}
	}
	for _, key := range required {
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
