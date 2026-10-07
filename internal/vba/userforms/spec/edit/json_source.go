package edit

import (
	"bytes"
	"cmp"
	"encoding/json"
	"errors"
	"fmt"
	"math"
	"reflect"
	"slices"
	"strings"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

type jsonSourceKind uint8

const (
	jsonSourceInvalid jsonSourceKind = iota
	jsonSourceObject
	jsonSourceArray
	jsonSourceString
	jsonSourceNumber
	jsonSourceOther
)

type jsonSourceMember struct {
	key      string
	keyStart int
	keyEnd   int
	value    *jsonSourceNode
}

type jsonSourceNode struct {
	kind     jsonSourceKind
	start    int
	end      int
	text     string
	members  []jsonSourceMember
	elements []*jsonSourceNode
}

type jsonSourceParser struct {
	source []byte
	index  int
}

type duplicateJSONKeyError struct {
	key    string
	offset int
}

func (e *duplicateJSONKeyError) Error() string {
	return fmt.Sprintf("duplicate object key %q at byte %d", e.key, e.offset)
}

func applyJSON(input spec.SpecInput, source []byte, operations []Operation) (Result, error) {
	originalRoot, err := parseJSONSource(source)
	if err != nil {
		if _, duplicate := errors.AsType[*duplicateJSONKeyError](err); duplicate {
			return Result{}, failure(-1, Operation{}, "UFE008", err.Error(), "Remove duplicate object keys.")
		}
		return Result{}, validationError(-1, Operation{}, err)
	}
	model, err := spec.ParseFormSpec(input, source)
	if err != nil {
		return Result{}, validationError(-1, Operation{}, err)
	}
	updated := bytes.Clone(source)
	pieces := pieceTable{{start: 0, end: len(source)}}
	types := jsonControlTypes(model)
	originalControls := map[string][]*jsonSourceNode{}
	indexJSONControls(jsonSourceMemberValue(originalRoot, "controls"), originalControls)
	newControls := map[string]bool{}
	applyEdits := func(edits []SourceEdit) {
		for i := len(edits) - 1; i >= 0; i-- {
			edit := edits[i]
			pieces = pieces.replace(edit)
			updated = applySourceEdits(updated, []SourceEdit{edit})
		}
	}
	for index, op := range operations {
		if err := checkPayload(op); err != nil {
			return Result{}, annotate(index, op, err)
		}
		root, err := parseJSONSource(updated)
		if err != nil {
			return Result{}, annotate(index, op, err)
		}
		if op.Type == AddControl {
			if err := validateAddControl(op); err != nil {
				return Result{}, annotate(index, op, err)
			}
			encoded, err := json.Marshal(op.Control)
			if err != nil {
				return Result{}, annotate(index, op, failure(-1, op, "UFE003", "Control payload is not serializable: "+err.Error(), "Supply serializable FormSpec control fields."))
			}
			edit, err := jsonAddControlEdit(updated, root, encoded)
			if err != nil {
				return Result{}, annotate(index, op, err)
			}
			applyEdits([]SourceEdit{edit})
			newControls[op.Control.ID] = true
			model, err = normalizedJSONModel(updated)
			if err != nil {
				return Result{}, annotate(index, op, err)
			}
			types = jsonControlTypes(model)
			continue
		}
		nodes := map[string][]*jsonSourceNode{}
		indexJSONControls(jsonSourceMemberValue(root, "controls"), nodes)
		if op.Type == RemoveControl {
			if len(nodes[op.ControlID]) != 1 {
				return Result{}, failure(index, op, "UFE004", "Control ID did not resolve to exactly one JSON control.", "Use a unique explicit control ID.")
			}
			removeIDs, err := jsonControlRemovalOrder(op, model, root, nodes)
			if err != nil {
				return Result{}, annotate(index, op, err)
			}
			removeSet := make(map[string]bool, len(removeIDs))
			for _, id := range removeIDs {
				removeSet[id] = true
			}
			for _, id := range removeIDs {
				currentRoot, err := parseJSONSource(updated)
				if err != nil {
					return Result{}, annotate(index, op, err)
				}
				currentNodes := map[string][]*jsonSourceNode{}
				indexJSONControls(jsonSourceMemberValue(currentRoot, "controls"), currentNodes)
				if len(currentNodes[id]) != 1 {
					return Result{}, failure(index, op, "UFE004", "Control ID did not resolve to exactly one JSON control.", "Use a unique explicit control ID.")
				}
				edit, err := jsonRemoveControlEdit(updated, currentRoot, currentNodes[id][0], removeSet)
				if err != nil {
					return Result{}, annotate(index, op, err)
				}
				applyEdits([]SourceEdit{edit})
				delete(newControls, id)
			}
			model, err = normalizedJSONModel(updated)
			if err != nil {
				return Result{}, annotate(index, op, err)
			}
			types = jsonControlTypes(model)
			continue
		}
		var target *jsonSourceNode
		var originalTarget *jsonSourceNode
		values := map[string]json.RawMessage{}
		keys := []string{}
		put := func(key string, value any) error {
			encoded, err := marshalJSONScalar(value)
			if err != nil {
				return err
			}
			if old := jsonSourceMemberValue(originalTarget, key); old != nil && old.kind != jsonSourceObject && old.kind != jsonSourceArray {
				original := source[old.start:old.end]
				equal, err := equalJSONScalar(original, encoded)
				if err != nil {
					return err
				}
				if equal {
					encoded = bytes.Clone(original)
				}
			}
			keys = append(keys, key)
			values[key] = encoded
			return nil
		}
		if op.Type == SetFormProperty {
			contract, ok := EditableFormProperties()[op.Field]
			if !ok {
				return Result{}, annotate(index, op, unsupportedProperty(op))
			}
			if err := checkPropertyValue(op, contract); err != nil {
				return Result{}, annotate(index, op, err)
			}
			target = jsonSourceMemberValue(root, "form")
			originalTarget = jsonSourceMemberValue(originalRoot, "form")
			key := op.Field
			if nested, ok := strings.CutPrefix(key, "build."); ok {
				build := jsonSourceMemberValue(target, "build")
				if build == nil {
					encodedValue, err := marshalJSONScalar(op.Value)
					if err != nil {
						return Result{}, annotate(index, op, err)
					}
					encoded, err := json.Marshal(map[string]json.RawMessage{nested: encodedValue})
					if err != nil {
						return Result{}, annotate(index, op, err)
					}
					keys, values["build"] = []string{"build"}, encoded
				} else {
					target = build
					originalTarget = jsonSourceMemberValue(originalTarget, "build")
					if err := put(nested, op.Value); err != nil {
						return Result{}, annotate(index, op, err)
					}
				}
			} else if err := put(key, op.Value); err != nil {
				return Result{}, annotate(index, op, err)
			}
		} else {
			if op.Type != MoveControl && op.Type != ResizeControl && op.Type != SetControlProperty {
				return Result{}, failure(index, op, "UFE007", "Unsupported JSON semantic operation.", "Use a scalar property or geometry edit.")
			}
			if len(nodes[op.ControlID]) != 1 {
				return Result{}, failure(index, op, "UFE004", "Control ID did not resolve to exactly one JSON control.", "Use a unique explicit control ID.")
			}
			target = nodes[op.ControlID][0]
			if !newControls[op.ControlID] && len(originalControls[op.ControlID]) == 1 {
				originalTarget = originalControls[op.ControlID][0]
			}
			switch op.Type {
			case SetControlProperty:
				contract, ok := EditableControlProperties(types[op.ControlID])[op.Field]
				if !ok {
					return Result{}, annotate(index, op, unsupportedProperty(op))
				}
				if err := checkPropertyValue(op, contract); err != nil {
					return Result{}, annotate(index, op, err)
				}
				if err := put(op.Field, op.Value); err != nil {
					return Result{}, annotate(index, op, err)
				}
			case MoveControl, ResizeControl:
				var axes []struct {
					key   string
					value *float64
				}
				if op.Type == MoveControl {
					axes = []struct {
						key   string
						value *float64
					}{{"left", op.Left}, {"top", op.Top}}
				} else {
					axes = []struct {
						key   string
						value *float64
					}{{"width", op.Width}, {"height", op.Height}}
				}
				for _, axis := range axes {
					if axis.value == nil {
						return Result{}, annotate(index, op, missingGeometry(op))
					}
					if math.IsNaN(*axis.value) || math.IsInf(*axis.value, 0) {
						return Result{}, failure(index, op, "UFE003", "Geometry values must be finite numbers.", "Supply finite numbers for both geometry axes.")
					}
					if err := put(axis.key, *axis.value); err != nil {
						return Result{}, annotate(index, op, err)
					}
				}
			}
		}
		edits, err := jsonPropertyEdits(updated, target, keys, values)
		if err != nil {
			return Result{}, annotate(index, op, err)
		}
		applyEdits(edits)
	}
	after, err := spec.ParseFormSpec(input, updated)
	if err != nil {
		index, op := -1, Operation{}
		if len(operations) > 0 {
			index, op = len(operations)-1, operations[len(operations)-1]
		}
		return Result{}, validationError(index, op, err)
	}
	return Result{Source: updated, Document: after, Edits: pieces.edits(source), Warnings: after.ValidationWarnings}, nil
}

func jsonControlTypes(model spec.FormSpec) map[string]string {
	types := make(map[string]string, len(model.Controls))
	for _, control := range model.Controls {
		types[control.ID] = control.Type
	}
	return types
}

func normalizedJSONModel(source []byte) (spec.FormSpec, error) {
	var model spec.FormSpec
	if err := json.Unmarshal(source, &model); err != nil {
		return spec.FormSpec{}, err
	}
	return spec.NormalizeFormSpec(model), nil
}

func marshalJSONScalar(value any) ([]byte, error) {
	node, err := scalar(value)
	if err != nil {
		return nil, err
	}
	// Encode the scalar, not arbitrary MarshalJSON methods on named Go values.
	var plain any
	if err := node.Decode(&plain); err != nil {
		return nil, err
	}
	if number, ok := value.(json.Number); ok {
		plain = number
	}
	return json.Marshal(plain)
}

func jsonPropertyEdits(source []byte, object *jsonSourceNode, keys []string, values map[string]json.RawMessage) ([]SourceEdit, error) {
	if object == nil || object.kind != jsonSourceObject {
		return nil, failure(-1, Operation{}, "UFE005", "Property target must be an object.", "Supply an explicit authoring object.")
	}
	var edits []SourceEdit
	var missing []string
	for _, key := range keys {
		old := jsonSourceMemberValue(object, key)
		if old == nil {
			missing = append(missing, key)
			continue
		}
		if old.kind == jsonSourceObject || old.kind == jsonSourceArray {
			return nil, failure(-1, Operation{}, "UFE005", "Cannot replace a non-scalar authoring field.", "Use a scalar field.")
		}
		equal, err := equalJSONScalar(source[old.start:old.end], values[key])
		if err != nil {
			return nil, err
		}
		if equal {
			continue
		}
		edits = append(edits, SourceEdit{Start: old.start, End: old.end, Text: string(values[key])})
	}
	if len(missing) > 0 {
		edits = append(edits, jsonObjectInsertion(source, object, missing, values))
	}
	slices.SortFunc(edits, func(a, b SourceEdit) int { return cmp.Compare(a.Start, b.Start) })
	return edits, nil
}

func equalJSONScalar(left, right []byte) (bool, error) {
	decode := func(data []byte) (any, error) {
		var value any
		decoder := json.NewDecoder(bytes.NewReader(data))
		decoder.UseNumber()
		err := decoder.Decode(&value)
		return value, err
	}
	a, err := decode(left)
	if err != nil {
		return false, err
	}
	b, err := decode(right)
	if err != nil {
		return false, err
	}
	if reflect.DeepEqual(a, b) {
		return true, nil
	}
	an, bn := rational(a), rational(b)
	return an != nil && bn != nil && an.Cmp(bn) == 0, nil
}

func parseJSONSource(source []byte) (*jsonSourceNode, error) {
	if !utf8.Valid(source) {
		return nil, fmt.Errorf("source is not valid UTF-8")
	}
	if !json.Valid(source) {
		return nil, fmt.Errorf("invalid JSON syntax")
	}
	p := jsonSourceParser{source: source}
	p.skipSpace()
	root, err := p.value()
	if err != nil {
		return nil, err
	}
	p.skipSpace()
	if p.index != len(source) {
		return nil, fmt.Errorf("unexpected data at byte %d", p.index)
	}
	return root, nil
}

func (p *jsonSourceParser) value() (*jsonSourceNode, error) {
	p.skipSpace()
	if p.index >= len(p.source) {
		return nil, fmt.Errorf("expected a JSON value at byte %d", p.index)
	}
	switch p.source[p.index] {
	case '{':
		return p.object()
	case '[':
		return p.array()
	case '"':
		return p.string()
	default:
		return p.primitive()
	}
}

func (p *jsonSourceParser) object() (*jsonSourceNode, error) {
	node := &jsonSourceNode{kind: jsonSourceObject, start: p.index}
	p.index++
	keys := map[string]struct{}{}
	p.skipSpace()
	if p.consume('}') {
		node.end = p.index
		return node, nil
	}
	for {
		p.skipSpace()
		keyNode, err := p.string()
		if err != nil {
			return nil, err
		}
		if _, duplicate := keys[keyNode.text]; duplicate {
			return nil, &duplicateJSONKeyError{key: keyNode.text, offset: keyNode.start}
		}
		keys[keyNode.text] = struct{}{}
		p.skipSpace()
		if !p.consume(':') {
			return nil, fmt.Errorf("expected ':' after object key at byte %d", p.index)
		}
		value, err := p.value()
		if err != nil {
			return nil, err
		}
		node.members = append(node.members, jsonSourceMember{key: keyNode.text, keyStart: keyNode.start, keyEnd: keyNode.end, value: value})
		p.skipSpace()
		if p.consume('}') {
			node.end = p.index
			return node, nil
		}
		if !p.consume(',') {
			return nil, fmt.Errorf("expected ',' or '}' at byte %d", p.index)
		}
	}
}

func (p *jsonSourceParser) array() (*jsonSourceNode, error) {
	node := &jsonSourceNode{kind: jsonSourceArray, start: p.index}
	p.index++
	p.skipSpace()
	if p.consume(']') {
		node.end = p.index
		return node, nil
	}
	for {
		value, err := p.value()
		if err != nil {
			return nil, err
		}
		node.elements = append(node.elements, value)
		p.skipSpace()
		if p.consume(']') {
			node.end = p.index
			return node, nil
		}
		if !p.consume(',') {
			return nil, fmt.Errorf("expected ',' or ']' at byte %d", p.index)
		}
	}
}

func (p *jsonSourceParser) string() (*jsonSourceNode, error) {
	if !p.consume('"') {
		return nil, fmt.Errorf("expected a JSON string at byte %d", p.index)
	}
	start := p.index - 1
	for p.index < len(p.source) {
		switch p.source[p.index] {
		case '\\':
			p.index += 2
		case '"':
			p.index++
			var value string
			if err := json.Unmarshal(p.source[start:p.index], &value); err != nil {
				return nil, fmt.Errorf("invalid JSON string at byte %d: %w", start, err)
			}
			return &jsonSourceNode{kind: jsonSourceString, start: start, end: p.index, text: value}, nil
		default:
			if p.source[p.index] < 0x20 {
				return nil, fmt.Errorf("unescaped control character at byte %d", p.index)
			}
			p.index++
		}
	}
	return nil, fmt.Errorf("unterminated JSON string at byte %d", start)
}

func (p *jsonSourceParser) primitive() (*jsonSourceNode, error) {
	start := p.index
	for p.index < len(p.source) && !isJSONDelimiter(p.source[p.index]) {
		p.index++
	}
	if start == p.index {
		return nil, fmt.Errorf("expected a JSON value at byte %d", p.index)
	}
	token := p.source[start:p.index]
	if bytes.Equal(token, []byte("true")) || bytes.Equal(token, []byte("false")) || bytes.Equal(token, []byte("null")) {
		return &jsonSourceNode{kind: jsonSourceOther, start: start, end: p.index, text: string(token)}, nil
	}
	var number json.Number
	if err := json.Unmarshal(token, &number); err != nil {
		return nil, fmt.Errorf("invalid JSON value at byte %d", start)
	}
	return &jsonSourceNode{kind: jsonSourceNumber, start: start, end: p.index, text: string(token)}, nil
}

func (p *jsonSourceParser) skipSpace() {
	for p.index < len(p.source) {
		switch p.source[p.index] {
		case ' ', '\t', '\r', '\n':
			p.index++
		default:
			return
		}
	}
}

func (p *jsonSourceParser) consume(value byte) bool {
	if p.index < len(p.source) && p.source[p.index] == value {
		p.index++
		return true
	}
	return false
}

func isJSONDelimiter(value byte) bool {
	switch value {
	case ' ', '\t', '\r', '\n', ',', ']', '}':
		return true
	default:
		return false
	}
}

func jsonSourceMemberValue(object *jsonSourceNode, key string) *jsonSourceNode {
	if object == nil || object.kind != jsonSourceObject {
		return nil
	}
	for _, member := range object.members {
		if member.key == key {
			return member.value
		}
	}
	return nil
}

func indexJSONControls(sequence *jsonSourceNode, nodesByID map[string][]*jsonSourceNode) {
	if sequence == nil || sequence.kind != jsonSourceArray {
		return
	}
	for _, control := range sequence.elements {
		if control.kind != jsonSourceObject {
			continue
		}
		if id := jsonSourceMemberValue(control, "id"); id != nil && id.kind == jsonSourceString {
			nodesByID[id.text] = append(nodesByID[id.text], control)
		}
		indexJSONControls(jsonSourceMemberValue(control, "controls"), nodesByID)
	}
}

func jsonControlRemovalOrder(op Operation, model spec.FormSpec, root *jsonSourceNode, nodes map[string][]*jsonSourceNode) ([]string, error) {
	remove := map[string]bool{op.ControlID: true}
	for changed := true; changed; {
		changed = false
		for _, control := range model.Controls {
			if remove[control.ParentID] && !remove[control.ID] {
				remove[control.ID] = true
				changed = true
			}
		}
	}
	if len(remove) > 1 && !op.Cascade {
		return nil, failure(-1, op, "UFE007", "The control has descendants.", "Use cascade=true or remove/reparent descendants first.")
	}

	depths := map[string]int{}
	var indexDepth func(*jsonSourceNode, int)
	indexDepth = func(sequence *jsonSourceNode, depth int) {
		if sequence == nil || sequence.kind != jsonSourceArray {
			return
		}
		for _, control := range sequence.elements {
			if control.kind != jsonSourceObject {
				continue
			}
			if id := jsonSourceMemberValue(control, "id"); id != nil && id.kind == jsonSourceString {
				depths[id.text] = depth
			}
			indexDepth(jsonSourceMemberValue(control, "controls"), depth+1)
		}
	}
	indexDepth(jsonSourceMemberValue(root, "controls"), 0)

	order := make([]string, 0, len(remove))
	positions := make(map[string]int, len(remove))
	for position, control := range model.Controls {
		if remove[control.ID] {
			if _, exists := positions[control.ID]; !exists {
				positions[control.ID] = position
				order = append(order, control.ID)
			}
		}
	}
	if len(order) != len(remove) {
		return nil, failure(-1, op, "UFE004", "A canonical control did not resolve to exactly one JSON control.", "Use a unique explicit control ID.")
	}
	for _, id := range order {
		if len(nodes[id]) != 1 {
			return nil, failure(-1, op, "UFE004", "Control ID did not resolve to exactly one JSON control.", "Use a unique explicit control ID.")
		}
	}
	slices.SortFunc(order, func(left, right string) int {
		if depthOrder := cmp.Compare(depths[right], depths[left]); depthOrder != 0 {
			return depthOrder
		}
		return cmp.Compare(positions[right], positions[left])
	})
	return order, nil
}

func jsonAddControlEdit(source []byte, root *jsonSourceNode, encoded []byte) (SourceEdit, error) {
	controls := jsonSourceMemberValue(root, "controls")
	if controls.kind == jsonSourceArray {
		return jsonArrayAppendEdit(source, controls, root, encoded), nil
	}
	if controls.kind == jsonSourceOther && controls.text == "null" {
		return SourceEdit{Start: controls.start, End: controls.end, Text: string(jsonArrayValueForParent(source, root, encoded))}, nil
	}
	return SourceEdit{}, failure(-1, Operation{}, "UFE005", "The controls field must be an array or null.", "Use a valid controls array and retry.")
}

func jsonArrayValueForParent(source []byte, parent *jsonSourceNode, encoded []byte) []byte {
	parentStyle := jsonObjectStyle(source, parent)
	if !parentStyle.multiline {
		return []byte("[" + string(encoded) + "]")
	}
	indentUnit := jsonIndentUnit(source, parent)
	itemIndent := parentStyle.indent + indentUnit
	formatted := jsonIndentValue(encoded, itemIndent, indentUnit, parentStyle.newline)
	return []byte("[" + parentStyle.newline + itemIndent + string(formatted) + parentStyle.newline + parentStyle.indent + "]")
}

func jsonArrayAppendEdit(source []byte, array, parent *jsonSourceNode, encoded []byte) SourceEdit {
	style := jsonArrayStyle(source, array, parent)
	if len(array.elements) == 0 {
		var contents string
		if style.multiline {
			formatted := jsonIndentValue(encoded, style.itemIndent, style.indentUnit, style.newline)
			contents = style.newline + style.itemIndent + string(formatted) + style.newline + style.closeIndent
		} else {
			contents = string(encoded)
		}
		return SourceEdit{Start: array.start + 1, End: array.end - 1, Text: contents}
	}
	last := array.elements[len(array.elements)-1]
	if style.multiline {
		formatted := jsonIndentValue(encoded, style.itemIndent, style.indentUnit, style.newline)
		return SourceEdit{Start: last.end, End: last.end, Text: "," + style.newline + style.itemIndent + string(formatted)}
	}
	return SourceEdit{Start: last.end, End: last.end, Text: "," + style.afterComma + string(encoded)}
}

type jsonArrayFormatting struct {
	afterComma  string
	newline     string
	itemIndent  string
	closeIndent string
	indentUnit  string
	multiline   bool
}

func jsonArrayStyle(source []byte, array, parent *jsonSourceNode) jsonArrayFormatting {
	style := jsonArrayFormatting{newline: "\n", indentUnit: jsonIndentUnit(source, parent)}
	keyIndent := jsonArrayKeyIndent(source, parent, array)
	parentStyle := jsonObjectStyle(source, parent)
	if len(array.elements) == 0 {
		body := source[array.start+1 : array.end-1]
		if bytes.ContainsAny(body, "\r\n") || parentStyle.multiline {
			style.multiline = true
			style.newline = parentStyle.newline
			if bytes.ContainsAny(body, "\r\n") {
				style.newline = jsonNewlineIn(body)
			}
			style.closeIndent = keyIndent
			style.itemIndent = keyIndent + style.indentUnit
		}
		return style
	}

	for i := 1; i < len(array.elements); i++ {
		spacing := source[array.elements[i-1].end:array.elements[i].start]
		if bytes.ContainsAny(spacing, "\r\n") {
			style.multiline = true
			style.newline = jsonNewlineIn(spacing)
			style.itemIndent = jsonIndentForKey(source, array.elements[i].start)
		} else if i == len(array.elements)-1 && len(spacing) > 1 {
			style.afterComma = string(spacing[1:])
		}
	}
	first := array.elements[0]
	last := array.elements[len(array.elements)-1]
	leading := source[array.start+1 : first.start]
	trailing := source[last.end : array.end-1]
	if bytes.ContainsAny(leading, "\r\n") || bytes.ContainsAny(trailing, "\r\n") {
		style.multiline = true
		if bytes.ContainsAny(leading, "\r\n") {
			style.newline = jsonNewlineIn(leading)
		} else {
			style.newline = jsonNewlineIn(trailing)
		}
	}
	if style.multiline {
		if style.itemIndent == "" {
			style.itemIndent = jsonIndentForKey(source, last.start)
		}
		if style.itemIndent == "" {
			style.itemIndent = keyIndent + style.indentUnit
		}
		style.closeIndent = jsonIndentForKey(source, array.end-1)
		if style.closeIndent == "" {
			style.closeIndent = keyIndent
		}
		if last.kind == jsonSourceObject && len(last.members) > 0 {
			memberIndent := jsonIndentForKey(source, last.members[0].keyStart)
			if strings.HasPrefix(memberIndent, style.itemIndent) && len(memberIndent) > len(style.itemIndent) {
				style.indentUnit = memberIndent[len(style.itemIndent):]
			}
		}
	}
	return style
}

func jsonArrayKeyIndent(source []byte, parent, array *jsonSourceNode) string {
	if parent == nil {
		return ""
	}
	for _, member := range parent.members {
		if member.value == array {
			return jsonIndentForKey(source, member.keyStart)
		}
	}
	return ""
}

func jsonIndentUnit(source []byte, object *jsonSourceNode) string {
	if object != nil {
		for _, member := range object.members {
			baseIndent := jsonIndentForKey(source, member.keyStart)
			switch member.value.kind {
			case jsonSourceObject:
				if len(member.value.members) > 0 {
					childIndent := jsonIndentForKey(source, member.value.members[0].keyStart)
					if strings.HasPrefix(childIndent, baseIndent) && len(childIndent) > len(baseIndent) {
						return childIndent[len(baseIndent):]
					}
				}
			case jsonSourceArray:
				if len(member.value.elements) > 0 {
					childIndent := jsonIndentForKey(source, member.value.elements[0].start)
					if strings.HasPrefix(childIndent, baseIndent) && len(childIndent) > len(baseIndent) {
						return childIndent[len(baseIndent):]
					}
				}
			}
		}
	}
	return "  "
}

func jsonIndentValue(encoded []byte, prefix, indent, newline string) []byte {
	var formatted bytes.Buffer
	if err := json.Indent(&formatted, encoded, prefix, indent); err != nil {
		return bytes.Clone(encoded)
	}
	return []byte(strings.ReplaceAll(formatted.String(), "\n", newline))
}

func jsonControlArray(sequence, target *jsonSourceNode) *jsonSourceNode {
	if sequence == nil || sequence.kind != jsonSourceArray {
		return nil
	}
	for _, control := range sequence.elements {
		if control == target {
			return sequence
		}
		if nested := jsonControlArray(jsonSourceMemberValue(control, "controls"), target); nested != nil {
			return nested
		}
	}
	return nil
}

func jsonRemoveControlEdit(source []byte, root, target *jsonSourceNode, removeSet map[string]bool) (SourceEdit, error) {
	if jsonHasUnselectedNestedControl(target, removeSet) {
		return SourceEdit{}, failure(-1, Operation{}, "UFE006", "Removing this JSON control would also remove a nested control outside the canonical removal set.", "Resolve the nested control relationship before retrying.")
	}
	array := jsonControlArray(jsonSourceMemberValue(root, "controls"), target)
	if array == nil {
		return SourceEdit{}, fmt.Errorf("control array for JSON source node was not found")
	}
	index := slices.Index(array.elements, target)
	if index < 0 {
		return SourceEdit{}, fmt.Errorf("control item in JSON source array was not found")
	}
	if len(array.elements) == 1 {
		return SourceEdit{Start: target.start, End: target.end}, nil
	}
	if index == 0 {
		return SourceEdit{Start: target.start, End: array.elements[1].start}, nil
	}
	return SourceEdit{Start: array.elements[index-1].end, End: target.end}, nil
}

func jsonHasUnselectedNestedControl(control *jsonSourceNode, removeSet map[string]bool) bool {
	children := jsonSourceMemberValue(control, "controls")
	if children == nil || children.kind != jsonSourceArray {
		return false
	}
	for _, child := range children.elements {
		if child.kind != jsonSourceObject {
			return true
		}
		id := jsonSourceMemberValue(child, "id")
		if id == nil || id.kind != jsonSourceString || !removeSet[id.text] || jsonHasUnselectedNestedControl(child, removeSet) {
			return true
		}
	}
	return false
}

func jsonObjectInsertion(source []byte, object *jsonSourceNode, keys []string, values map[string]json.RawMessage) SourceEdit {
	position := object.start + 1
	style := jsonObjectStyle(source, object)
	var text strings.Builder
	if len(object.members) > 0 {
		position = object.members[len(object.members)-1].value.end
		text.WriteByte(',')
		text.WriteString(style.afterComma)
	} else if style.multiline {
		text.WriteString(style.newline)
		text.WriteString(style.indent)
	}
	for i, key := range keys {
		if i > 0 {
			text.WriteByte(',')
			text.WriteString(style.afterComma)
		}
		encodedKey, _ := json.Marshal(key)
		text.Write(encodedKey)
		text.WriteString(style.colon)
		text.Write(values[key])
	}
	return SourceEdit{Start: position, End: position, Text: text.String()}
}

type jsonObjectFormatting struct {
	afterComma string
	colon      string
	newline    string
	indent     string
	multiline  bool
}

func jsonObjectStyle(source []byte, object *jsonSourceNode) jsonObjectFormatting {
	style := jsonObjectFormatting{colon: ":", newline: "\n"}
	if len(object.members) == 0 && bytes.ContainsAny(source[object.start+1:object.end-1], "\r\n") {
		style.multiline = true
		style.newline = jsonNewlineIn(source[object.start+1 : object.end-1])
		style.indent = jsonIndentForKey(source, object.end-1) + "  "
	}
	for _, member := range object.members {
		if style.colon == ":" {
			style.colon = string(source[member.keyEnd:member.value.start])
		}
		if style.indent == "" {
			lineStart := bytes.LastIndexByte(source[:member.keyStart], '\n') + 1
			prefix := source[lineStart:member.keyStart]
			if len(bytes.Trim(prefix, " \t")) == 0 {
				style.indent = string(prefix)
			}
		}
	}
	if len(object.members) > 1 {
		for i := 1; i < len(object.members); i++ {
			previous := object.members[i-1].value.end
			current := object.members[i].keyStart
			spacing := source[previous:current]
			if bytes.ContainsAny(spacing, "\r\n") {
				style.multiline = true
				style.afterComma = jsonNewlineIn(spacing) + jsonIndentForKey(source, current)
				break
			}
			if len(spacing) > 1 {
				style.afterComma = string(spacing[1:])
			}
		}
	}
	if !style.multiline && len(object.members) > 0 {
		lastEnd := object.members[len(object.members)-1].value.end
		if bytes.ContainsAny(source[lastEnd:object.end-1], "\r\n") {
			style.multiline = true
			style.newline = jsonNewlineIn(source[lastEnd : object.end-1])
			if style.indent == "" {
				style.indent = jsonIndentForKey(source, object.members[0].keyStart)
			}
			style.afterComma = style.newline + style.indent
		}
	}
	if style.multiline {
		if style.indent == "" {
			closeIndent := jsonIndentForKey(source, object.end-1)
			style.indent = closeIndent + "  "
		}
		if style.afterComma == "" {
			style.afterComma = style.newline + style.indent
		}
	}
	return style
}

func jsonNewlineIn(source []byte) string {
	if i := bytes.IndexByte(source, '\n'); i >= 0 && i > 0 && source[i-1] == '\r' {
		return "\r\n"
	}
	if bytes.ContainsRune(source, '\r') && !bytes.ContainsRune(source, '\n') {
		return "\r"
	}
	return "\n"
}

func jsonIndentForKey(source []byte, offset int) string {
	lineStart := bytes.LastIndexByte(source[:offset], '\n') + 1
	indent := source[lineStart:offset]
	if len(bytes.Trim(indent, " \t")) == 0 {
		return string(indent)
	}
	return ""
}

func applySourceEdits(source []byte, edits []SourceEdit) []byte {
	var result bytes.Buffer
	result.Grow(len(source))
	cursor := 0
	for _, edit := range edits {
		result.Write(source[cursor:edit.Start])
		result.WriteString(edit.Text)
		cursor = edit.End
	}
	result.Write(source[cursor:])
	return result.Bytes()
}
