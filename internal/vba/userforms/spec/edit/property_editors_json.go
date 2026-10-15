package edit

import (
	"bytes"
	"encoding/json"
	"maps"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func jsonEditorOperation(source []byte, root *jsonSourceNode, nodes map[string][]*jsonSourceNode, model spec.FormSpec, op Operation) ([]SourceEdit, error) {
	target := jsonSourceMemberValue(root, "form")
	properties := EditableFormEditorProperties()
	if op.ControlID != "" {
		control, ok := controlByID(model, op.ControlID)
		if !ok || len(nodes[op.ControlID]) != 1 {
			return nil, failure(-1, op, "UFE004", "Control ID did not resolve to exactly one JSON control.", "Use a unique explicit control ID.")
		}
		if op.Type == SetGeometry && strings.EqualFold(control.Type, "Page") {
			return nil, unsupportedProperty(op)
		}
		target = nodes[op.ControlID][0]
		properties = editablePropertiesForControl(control, true)
	}
	if op.Type == SetGeometry {
		values, err := geometryValues(op)
		if err != nil {
			return nil, err
		}
		return jsonScalarEdits(source, target, values)
	}
	contract, ok := properties[op.Field]
	if !ok {
		return nil, unsupportedProperty(op)
	}
	if err := checkPropertyValue(op, contract); err != nil {
		return nil, err
	}
	if bag := jsonSourceMemberValue(target, "properties"); bag != nil && legacyPropertyName(op.Field) != "" {
		count := 0
		for _, member := range bag.members {
			if strings.EqualFold(member.key, legacyPropertyName(op.Field)) {
				count++
			}
		}
		if count > 1 {
			return nil, failure(-1, op, "UFE008", "Ambiguous legacy property aliases.", "Keep one canonical legacy property key before editing.")
		}
	}
	if op.Field == "font" && jsonSourceMemberValue(target, "font") == nil {
		if bag := jsonSourceMemberValue(target, "properties"); bag != nil {
			for _, member := range bag.members {
				if strings.EqualFold(member.key, "font") {
					var legacy any
					decoder := json.NewDecoder(bytes.NewReader(source[member.value.start:member.value.end]))
					decoder.UseNumber()
					if err := decoder.Decode(&legacy); err != nil {
						return nil, err
					}
					value, err := mergeLegacyFont(op, contract, legacy)
					if err != nil {
						return nil, err
					}
					op.Value = value
				}
			}
		}
	}
	encoded, err := marshalEditorValue(op.Value)
	if err != nil {
		return nil, err
	}
	var edits []SourceEdit
	old := jsonSourceMemberValue(target, op.Field)
	if object, ok := op.Value.(map[string]any); ok && old != nil && old.kind == jsonSourceObject && op.Field != "picture" {
		values := map[string]json.RawMessage{}
		for key, value := range object {
			encoded, err := marshalJSONScalar(value)
			if err != nil {
				return nil, err
			}
			values[key] = encoded
		}
		edits, err = jsonPropertyEdits(source, old, slices.Sorted(maps.Keys(values)), values)
	} else if old != nil {
		equal, compareErr := equalJSONScalar(source[old.start:old.end], encoded)
		if compareErr != nil {
			return nil, compareErr
		}
		if !equal {
			edits = []SourceEdit{{Start: old.start, End: old.end, Text: string(encoded)}}
		}
	} else {
		edits, err = jsonPropertyEdits(source, target, []string{op.Field}, map[string]json.RawMessage{op.Field: encoded})
	}
	if err != nil {
		return nil, err
	}
	if op.Type == SetControlProperty && legacyPropertyName(op.Field) != "" {
		bag := jsonSourceMemberValue(target, "properties")
		if bag != nil {
			for index, member := range bag.members {
				if strings.EqualFold(member.key, legacyPropertyName(op.Field)) {
					edits = append(edits, jsonRemovePropertyEdit(bag, index))
				}
			}
		}
	}
	return edits, nil
}

func jsonRemovePropertyEdit(object *jsonSourceNode, index int) SourceEdit {
	member := object.members[index]
	start, end := member.keyStart, member.value.end
	if index+1 < len(object.members) {
		end = object.members[index+1].keyStart
	} else if index > 0 {
		start = object.members[index-1].value.end
	}
	return SourceEdit{Start: start, End: end}
}
