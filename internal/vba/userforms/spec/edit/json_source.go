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
	types := map[string]string{}
	originalControls := map[string][]*jsonSourceNode{}
	indexJSONControls(jsonSourceMemberValue(originalRoot, "controls"), originalControls)
	for _, control := range model.Controls {
		types[control.ID] = control.Type
	}
	for index, op := range operations {
		if err := checkPayload(op); err != nil {
			return Result{}, annotate(index, op, err)
		}
		root, err := parseJSONSource(updated)
		if err != nil {
			return Result{}, annotate(index, op, err)
		}
		nodes := map[string][]*jsonSourceNode{}
		indexJSONControls(jsonSourceMemberValue(root, "controls"), nodes)
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
			if len(originalControls[op.ControlID]) == 1 {
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
		// Apply backwards so all ranges remain relative to this source snapshot.
		for i := len(edits) - 1; i >= 0; i-- {
			edit := edits[i]
			pieces = pieces.replace(edit)
			updated = applySourceEdits(updated, []SourceEdit{edit})
		}
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
