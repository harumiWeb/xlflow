package edit

import (
	"cmp"
	"encoding/json"
	"fmt"
	"math"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

func controlByID(model spec.FormSpec, id string) (spec.FormSpecControl, bool) {
	index := slices.IndexFunc(model.Controls, func(control spec.FormSpecControl) bool { return control.ID == id })
	if index < 0 {
		return spec.FormSpecControl{}, false
	}
	return model.Controls[index], true
}

func jsonScalarEdits(source []byte, object *jsonSourceNode, values map[string]any) ([]SourceEdit, error) {
	encoded := make(map[string]json.RawMessage, len(values))
	for key, value := range values {
		data, err := marshalJSONScalar(value)
		if err != nil {
			return nil, err
		}
		encoded[key] = data
	}
	return jsonRawPropertyEdits(source, object, encoded)
}

func jsonRawPropertyEdits(source []byte, object *jsonSourceNode, values map[string]json.RawMessage) ([]SourceEdit, error) {
	keys := make([]string, 0, len(values))
	for key := range values {
		keys = append(keys, key)
	}
	slices.Sort(keys)
	return jsonPropertyEdits(source, object, keys, values)
}

func jsonTabOperation(source []byte, control *jsonSourceNode, model spec.FormSpecControl, op Operation) ([]SourceEdit, error) {
	if control == nil || control.kind != jsonSourceObject || !strings.EqualFold(model.Type, "TabStrip") {
		return nil, failure(-1, op, "UFE005", "Tab operations require a TabStrip control.", "Use the ID of a TabStrip control.")
	}
	tabs := jsonSourceMemberValue(control, "tabs")
	if tabs != nil && tabs.kind != jsonSourceArray && (op.Type != AddTab || tabs.kind != jsonSourceOther || tabs.text != "null") {
		return nil, failure(-1, op, "UFE005", "Tab collection must be an array.", "Use a valid tabs array and retry.")
	}
	count := 0
	if tabs != nil && tabs.kind == jsonSourceArray {
		count = len(tabs.elements)
	}
	values := map[string]any{}
	if op.Type == AddTab {
		if op.Tab == nil {
			return nil, failure(-1, op, "UFE003", "AddTab requires a tab payload.", "Supply tab.name and optional tab properties.")
		}
		encoded, err := json.Marshal(op.Tab)
		if err != nil {
			return nil, failure(-1, op, "UFE003", "Tab payload is not serializable: "+err.Error(), "Supply serializable FormSpec tab fields.")
		}
		if tabs == nil || tabs.kind != jsonSourceArray {
			raw := map[string]json.RawMessage{"tabs": json.RawMessage("[" + string(encoded) + "]")}
			if selected, changed := selectedAfterAdd(model.SelectedIndex, count); changed {
				selection, err := marshalJSONScalar(selected)
				if err != nil {
					return nil, err
				}
				raw["selectedIndex"] = selection
			}
			return jsonRawPropertyEdits(source, control, raw)
		} else {
			edits := []SourceEdit{jsonArrayAppendEdit(source, tabs, control, encoded)}
			if selected, changed := selectedAfterAdd(model.SelectedIndex, count); changed {
				values["selectedIndex"] = selected
			}
			if len(values) > 0 {
				selection, err := jsonScalarEdits(source, control, values)
				if err != nil {
					return nil, err
				}
				edits = append(edits, selection...)
			}
			return edits, nil
		}
	}
	if tabs == nil || tabs.kind != jsonSourceArray {
		return nil, failure(-1, op, "UFE004", "Tab collection was not found.", "Use an explicitly authored tabs array.")
	}
	if op.Type == SetTabProperty || op.Type == RemoveTab {
		if op.Index == nil || *op.Index < 0 || *op.Index >= count {
			return nil, failure(-1, op, "UFE003", "Tab index is out of range.", "Use a valid zero-based tab index.")
		}
	}
	switch op.Type {
	case RemoveTab:
		if selected, changed := selectedAfterRemove(model.SelectedIndex, *op.Index, count); changed {
			values["selectedIndex"] = selected
		}
		edits, err := jsonScalarEdits(source, control, values)
		if err != nil {
			return nil, err
		}
		return append(edits, jsonArrayItemRemovalEdit(tabs, *op.Index)), nil
	case ReorderTab:
		if op.Index == nil || op.ToIndex == nil || *op.Index < 0 || *op.Index >= count || *op.ToIndex < 0 || *op.ToIndex >= count {
			return nil, failure(-1, op, "UFE003", "ReorderTab requires valid from and to indices.", "Supply zero-based indices within the tab collection.")
		}
		if *op.Index == *op.ToIndex {
			return nil, nil
		}
		if selected, changed := selectedAfterReorder(model.SelectedIndex, *op.Index, *op.ToIndex); changed {
			values["selectedIndex"] = selected
		}
		edits, err := jsonScalarEdits(source, control, values)
		if err != nil {
			return nil, err
		}
		return append(edits, jsonArrayReorderEdit(source, tabs, control, *op.Index, *op.ToIndex)), nil
	case SetTabProperty:
		contract, ok := EditableTabProperties()[op.Field]
		if !ok {
			return nil, unsupportedProperty(op)
		}
		if err := checkPropertyValue(op, contract); err != nil {
			return nil, err
		}
		return jsonScalarEdits(source, tabs.elements[*op.Index], map[string]any{op.Field: op.Value})
	default:
		return nil, failure(-1, op, "UFE003", "Unknown tab operation.", "Use a supported tab operation type.")
	}
}

func jsonArrayItemRemovalEdit(array *jsonSourceNode, index int) SourceEdit {
	item := array.elements[index]
	if len(array.elements) == 1 {
		return SourceEdit{Start: item.start, End: item.end}
	}
	if index == 0 {
		return SourceEdit{Start: item.start, End: array.elements[1].start}
	}
	return SourceEdit{Start: array.elements[index-1].end, End: item.end}
}

func jsonArrayReorderEdit(source []byte, array, parent *jsonSourceNode, from, to int) SourceEdit {
	ordered := slices.Clone(array.elements)
	item := ordered[from]
	ordered = slices.Delete(ordered, from, from+1)
	ordered = slices.Insert(ordered, to, item)
	style := jsonArrayStyle(source, array, parent)
	var body string
	if style.multiline {
		var builder strings.Builder
		for index, element := range ordered {
			if index == 0 {
				builder.WriteString(style.newline)
			} else {
				builder.WriteByte(',')
				builder.WriteString(style.newline)
			}
			builder.WriteString(style.itemIndent)
			builder.Write(source[element.start:element.end])
		}
		builder.WriteString(style.newline)
		builder.WriteString(style.closeIndent)
		body = builder.String()
	} else {
		var builder strings.Builder
		if len(ordered) > 0 {
			builder.Write(source[array.start+1 : array.elements[0].start])
		}
		for index, element := range ordered {
			if index > 0 {
				builder.WriteByte(',')
				builder.WriteString(style.afterComma)
			}
			builder.Write(source[element.start:element.end])
		}
		if len(ordered) > 0 {
			builder.Write(source[array.elements[len(array.elements)-1].end : array.end-1])
		}
		body = builder.String()
	}
	return SourceEdit{Start: array.start + 1, End: array.end - 1, Text: body}
}

func orderedChildren(controls []spec.FormSpecControl, parentID, typeName string) []spec.FormSpecControl {
	children := make([]spec.FormSpecControl, 0)
	for _, control := range controls {
		if control.ParentID == parentID && (typeName == "" || strings.EqualFold(control.Type, typeName)) {
			children = append(children, control)
		}
	}
	slices.SortStableFunc(children, func(left, right spec.FormSpecControl) int {
		return cmp.Compare(*left.ZIndex, *right.ZIndex)
	})
	return children
}

func selectedAfterAdd(selected *int, oldCount int) (int, bool) {
	if selected != nil && oldCount == 0 && *selected == -1 {
		return 0, true
	}
	return 0, false
}

func selectedAfterRemove(selected *int, removed, oldCount int) (int, bool) {
	if selected == nil {
		return 0, false
	}
	switch {
	case *selected == removed:
		if oldCount == 1 {
			return -1, true
		}
		return 0, true
	case removed < *selected:
		return *selected - 1, true
	default:
		return 0, false
	}
}

func selectedAfterReorder(selected *int, from, to int) (int, bool) {
	if selected == nil || from == to {
		return 0, false
	}
	switch {
	case *selected == from:
		return to, true
	case from < *selected && to >= *selected:
		return *selected - 1, true
	case from > *selected && to <= *selected:
		return *selected + 1, true
	default:
		return 0, false
	}
}

func (e *engine) addPage(op Operation) error {
	if err := validateAddPage(op); err != nil {
		return err
	}
	parentID := op.Control.ParentID
	parent, ok := e.controls[parentID]
	if !ok {
		return failure(-1, op, "UFE004", "MultiPage control ID was not found.", "Set control.parentId to a MultiPage control ID from the current FormSpec.")
	}
	if !strings.EqualFold(parent.model.Type, "MultiPage") || e.unsafeTarget(parent.node) || field(parent.node, "id") == nil {
		return failure(-1, op, "UFE005", "AddPage requires an explicit MultiPage target.", "Use the ID of a MultiPage control.")
	}
	pages := orderedChildren(e.model.Controls, parentID, "Page")
	index, err := nextPageZIndex(pages, op)
	if err != nil {
		return err
	}
	page := *op.Control
	page.ZIndex = new(index)
	var node yaml.Node
	if err := node.Encode(page); err != nil {
		return err
	}
	if err := e.appendControl(&node); err != nil {
		return err
	}
	if selected, changed := selectedAfterAdd(parent.model.SelectedIndex, len(pages)); changed {
		return e.setControl(parentID, "selectedIndex", selected)
	}
	return nil
}

func validateAddPage(op Operation) error {
	if op.Control == nil || op.Control.ID == "" || op.Control.Name == "" || !strings.EqualFold(op.Control.Type, "Page") || op.Control.ParentID == "" || len(op.Control.Controls) > 0 || op.Control.Left != nil || op.Control.Top != nil || op.Control.Width != nil || op.Control.Height != nil {
		return failure(-1, op, "UFE003", "AddPage requires a flat Page with parentId and without child controls or geometry.", "Supply control.id, control.name, control.type=Page and control.parentId.")
	}
	if _, err := json.Marshal(op.Control); err != nil {
		return failure(-1, op, "UFE003", "Page payload is not serializable: "+err.Error(), "Supply serializable FormSpec Page fields.")
	}
	return nil
}

func nextPageZIndex(pages []spec.FormSpecControl, op Operation) (int, error) {
	if len(pages) == 0 {
		return 0, nil
	}
	last := *pages[len(pages)-1].ZIndex
	if last == math.MaxInt {
		return 0, failure(-1, op, "UFE003", "Page order has no available zIndex after the final page.", "Reorder the existing pages before adding another page.")
	}
	return last + 1, nil
}

func (e *engine) pageTarget(op Operation) (controlRef, controlRef, []spec.FormSpecControl, int, error) {
	page, ok := e.controls[op.ControlID]
	if !ok || !strings.EqualFold(page.model.Type, "Page") || page.model.ParentID == "" {
		return controlRef{}, controlRef{}, nil, -1, failure(-1, op, "UFE004", "Page control ID was not found.", "Use a Page control ID from the current FormSpec.")
	}
	parent, ok := e.controls[page.model.ParentID]
	if !ok || !strings.EqualFold(parent.model.Type, "MultiPage") {
		return controlRef{}, controlRef{}, nil, -1, failure(-1, op, "UFE005", "Page does not have a canonical MultiPage parent.", "Use a Page under a MultiPage control.")
	}
	pages := orderedChildren(e.model.Controls, parent.model.ID, "Page")
	index := slices.IndexFunc(pages, func(candidate spec.FormSpecControl) bool { return candidate.ID == op.ControlID })
	if index < 0 {
		return controlRef{}, controlRef{}, nil, -1, failure(-1, op, "UFE004", "Page control ID was not found among its MultiPage children.", "Use a Page control ID from the current FormSpec.")
	}
	if e.unsafeTarget(page.node) || e.unsafeTarget(parent.node) || field(page.node, "id") == nil || field(parent.node, "id") == nil {
		return controlRef{}, controlRef{}, nil, -1, unsupported()
	}
	return page, parent, pages, index, nil
}

func (e *engine) removePage(op Operation) error {
	_, parent, pages, index, err := e.pageTarget(op)
	if err != nil {
		return err
	}
	if selected, changed := selectedAfterRemove(parent.model.SelectedIndex, index, len(pages)); changed {
		if err := e.setControl(parent.model.ID, "selectedIndex", selected); err != nil {
			return err
		}
	}
	return e.removeControl(op)
}

func (e *engine) reorderPage(op Operation) error {
	if op.Index == nil {
		return failure(-1, op, "UFE003", "ReorderPage requires an index.", "Supply a zero-based page index.")
	}
	_, parent, pages, oldIndex, err := e.pageTarget(op)
	if err != nil {
		return err
	}
	if *op.Index < 0 || *op.Index >= len(pages) {
		return failure(-1, op, "UFE003", "Page index is out of range.", "Use an index within the MultiPage pages.")
	}
	if selected, changed := selectedAfterReorder(parent.model.SelectedIndex, oldIndex, *op.Index); changed {
		if err := e.setControl(parent.model.ID, "selectedIndex", selected); err != nil {
			return err
		}
	}
	return e.reorder(op)
}

func (e *engine) applyTab(op Operation) error {
	control, ok := e.controls[op.ControlID]
	if !ok {
		return failure(-1, op, "UFE004", "TabStrip control ID was not found.", "Use a TabStrip control ID from the current FormSpec.")
	}
	if !strings.EqualFold(control.model.Type, "TabStrip") || e.unsafeTarget(control.node) || field(control.node, "id") == nil {
		return failure(-1, op, "UFE005", "Tab operations require an explicit TabStrip target.", "Use the ID of a TabStrip control.")
	}
	sequence := field(control.node, "tabs")
	if sequence != nil && (sequence.Kind != yaml.SequenceNode && (op.Type != AddTab || sequence.Kind != yaml.ScalarNode || sequence.Tag != "!!null") || unsafe(sequence)) {
		return unsupported()
	}
	count := 0
	if sequence != nil {
		count = len(sequence.Content)
	}
	switch op.Type {
	case AddTab:
		if op.Tab == nil {
			return failure(-1, op, "UFE003", "AddTab requires a tab payload.", "Supply tab.name and optional tab properties.")
		}
		if _, err := json.Marshal(op.Tab); err != nil {
			return failure(-1, op, "UFE003", "Tab payload is not serializable: "+err.Error(), "Supply serializable FormSpec tab fields.")
		}
		var node yaml.Node
		if err := node.Encode(*op.Tab); err != nil {
			return err
		}
		if sequence == nil {
			sequence = &yaml.Node{Kind: yaml.SequenceNode, Tag: "!!seq", Content: []*yaml.Node{&node}}
			if err := e.set(control.node, "tabs", sequence); err != nil {
				return err
			}
		} else if sequence.Kind == yaml.ScalarNode && sequence.Tag == "!!null" {
			node.Style |= yaml.FlowStyle
			value := &yaml.Node{Kind: yaml.SequenceNode, Tag: "!!seq", Style: yaml.FlowStyle, Content: []*yaml.Node{&node}}
			text, err := encode(value, 2, e.newline())
			if err != nil {
				return err
			}
			start, end := e.offset(sequence), e.end(sequence)
			*sequence = *value
			if err := e.publish(SourceEdit{Start: start, End: end, Text: text}); err != nil {
				return err
			}
		} else if err := e.appendSequenceItem(sequence, &node); err != nil {
			return err
		}
		if selected, changed := selectedAfterAdd(control.model.SelectedIndex, count); changed {
			return e.setControl(op.ControlID, "selectedIndex", selected)
		}
		return nil
	case RemoveTab:
		if op.Index == nil || *op.Index < 0 || *op.Index >= count {
			return failure(-1, op, "UFE003", "RemoveTab requires an index within the tab collection.", "Supply a valid zero-based tab index.")
		}
		if selected, changed := selectedAfterRemove(control.model.SelectedIndex, *op.Index, count); changed {
			if err := e.setControl(op.ControlID, "selectedIndex", selected); err != nil {
				return err
			}
			var err error
			sequence, err = e.tabSequenceAfterReload(op)
			if err != nil {
				return err
			}
		}
		return e.removeItem(controlRef{node: sequence.Content[*op.Index], sequence: sequence})
	case ReorderTab:
		if op.Index == nil || op.ToIndex == nil || *op.Index < 0 || *op.Index >= count || *op.ToIndex < 0 || *op.ToIndex >= count {
			return failure(-1, op, "UFE003", "ReorderTab requires valid from and to indices.", "Supply zero-based indices within the tab collection.")
		}
		if *op.Index == *op.ToIndex {
			return nil
		}
		if selected, changed := selectedAfterReorder(control.model.SelectedIndex, *op.Index, *op.ToIndex); changed {
			if err := e.setControl(op.ControlID, "selectedIndex", selected); err != nil {
				return err
			}
			var err error
			sequence, err = e.tabSequenceAfterReload(op)
			if err != nil {
				return err
			}
		}
		return e.reorderSequence(sequence, *op.Index, *op.ToIndex)
	case SetTabProperty:
		if op.Index == nil || *op.Index < 0 || *op.Index >= count {
			return failure(-1, op, "UFE003", "SetTabProperty requires an index within the tab collection.", "Supply a valid zero-based tab index.")
		}
		contract, ok := EditableTabProperties()[op.Field]
		if !ok {
			return unsupportedProperty(op)
		}
		if err := checkPropertyValue(op, contract); err != nil {
			return err
		}
		node, err := scalar(op.Value)
		if err != nil {
			return err
		}
		return e.set(sequence.Content[*op.Index], op.Field, node)
	default:
		return failure(-1, op, "UFE003", "Unknown tab operation.", "Use a supported tab operation type.")
	}
}

func (e *engine) tabSequenceAfterReload(op Operation) (*yaml.Node, error) {
	control, ok := e.controls[op.ControlID]
	if !ok {
		return nil, failure(-1, op, "UFE004", "TabStrip control ID was not found after updating its selection.", "Use a TabStrip control ID from the current FormSpec.")
	}
	sequence := field(control.node, "tabs")
	if sequence == nil || sequence.Kind != yaml.SequenceNode || unsafe(sequence) {
		return nil, unsupported()
	}
	return sequence, nil
}

func (e *engine) appendSequenceItem(sequence, node *yaml.Node) error {
	if sequence == nil || sequence.Kind != yaml.SequenceNode {
		return fmt.Errorf("target is not a sequence")
	}
	if sequence.Anchor != "" {
		return unsupported()
	}
	if sequence.Style&yaml.FlowStyle != 0 {
		if unsafe(sequence) {
			return unsupported()
		}
		start, end := e.offset(sequence), e.end(sequence)
		sequence.Content = append(sequence.Content, node)
		text, err := encodeContainer(sequence, e.newline())
		if err != nil {
			return err
		}
		return e.publish(SourceEdit{Start: start, End: end, Text: text})
	}
	indent := max(0, sequence.Column-1)
	text, err := encode(&yaml.Node{Kind: yaml.SequenceNode, Content: []*yaml.Node{node}}, 2, e.newline())
	if err != nil {
		return err
	}
	padding := strings.Repeat(" ", indent)
	text = padding + strings.ReplaceAll(text, e.newline(), e.newline()+padding)
	position := e.lineEnd(e.end(sequence))
	if position > 0 && e.source[position-1] != '\n' {
		text = e.newline() + text
	} else {
		text += e.newline()
	}
	sequence.Content = append(sequence.Content, node)
	return e.publish(SourceEdit{Start: position, End: position, Text: text})
}

func (e *engine) reorderSequence(sequence *yaml.Node, from, to int) error {
	if sequence == nil || sequence.Kind != yaml.SequenceNode || unsafe(sequence) {
		return unsupported()
	}
	start, end := e.offset(sequence), e.end(sequence)
	item := sequence.Content[from]
	items := slices.Delete(sequence.Content, from, from+1)
	sequence.Content = slices.Insert(items, to, item)
	text, err := encodeContainer(sequence, e.newline())
	if err != nil {
		return err
	}
	if sequence.Style&yaml.FlowStyle == 0 {
		padding := strings.Repeat(" ", max(0, sequence.Column-1))
		text = strings.ReplaceAll(text, e.newline(), e.newline()+padding)
	}
	return e.publish(SourceEdit{Start: start, End: end, Text: text})
}
