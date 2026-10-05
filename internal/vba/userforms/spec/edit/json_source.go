package edit

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"math"
	"slices"
	"strconv"
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
	if !json.Valid(source) {
		return Result{}, failure(-1, Operation{}, "UFE002", "Invalid JSON FormSpec syntax.", "Correct the JSON syntax and retry.")
	}
	if _, err := spec.ParseFormSpec(input, source); err != nil {
		return Result{}, validationError(-1, Operation{}, err)
	}
	root, err := parseJSONSource(source)
	if err != nil {
		if _, duplicate := errors.AsType[*duplicateJSONKeyError](err); duplicate {
			return Result{}, failure(-1, Operation{}, "UFE008", err.Error(), "Remove duplicate keys so the target and geometry fields are unambiguous.")
		}
		return Result{}, failure(-1, Operation{}, "UFE002", "Invalid JSON FormSpec: "+err.Error(), "Correct the JSON syntax and remove duplicate object keys.")
	}
	if root.kind != jsonSourceObject {
		return Result{}, failure(-1, Operation{}, "UFE002", "The JSON FormSpec root must be an object.", "Supply a FormSpec object.")
	}

	nodesByID := map[string][]*jsonSourceNode{}
	controls := jsonSourceMemberValue(root, "controls")
	indexJSONControls(controls, nodesByID)

	changes := map[*jsonSourceNode]map[string]float64{}
	changedObjects := make([]*jsonSourceNode, 0)
	for index, op := range operations {
		if op.Type != MoveControl && op.Type != ResizeControl {
			return Result{}, failure(index, op, "UFE007", "JSON FormSpecs support only moveControl and resizeControl edits.", "Use moveControl or resizeControl for JSON geometry edits.")
		}
		if err := checkPayload(op); err != nil {
			return Result{}, annotate(index, op, err)
		}
		var fields []struct {
			name  string
			value *float64
		}
		switch op.Type {
		case MoveControl:
			if op.Left == nil || op.Top == nil {
				return Result{}, annotate(index, op, missingGeometry(op))
			}
			fields = []struct {
				name  string
				value *float64
			}{{"left", op.Left}, {"top", op.Top}}
		case ResizeControl:
			if op.Width == nil || op.Height == nil {
				return Result{}, annotate(index, op, missingGeometry(op))
			}
			fields = []struct {
				name  string
				value *float64
			}{{"width", op.Width}, {"height", op.Height}}
		}

		targets := nodesByID[op.ControlID]
		if len(targets) != 1 {
			return Result{}, failure(index, op, "UFE004", "Control ID did not resolve to exactly one JSON control.", "Use a unique explicit control ID from the current FormSpec.")
		}
		target := targets[0]
		for _, geometry := range fields {
			if math.IsNaN(*geometry.value) || math.IsInf(*geometry.value, 0) {
				return Result{}, failure(index, op, "UFE003", "Geometry values must be finite numbers.", "Supply finite JSON numbers for both geometry axes.")
			}
			if _, exists := changes[target]; !exists {
				changes[target] = map[string]float64{}
				changedObjects = append(changedObjects, target)
			}
			changes[target][geometry.name] = *geometry.value
		}
	}

	edits := make([]SourceEdit, 0, len(changes)*2)
	for _, object := range changedObjects {
		values := changes[object]
		missing := make([]string, 0, len(values))
		for _, key := range []string{"left", "top", "width", "height"} {
			value, requested := values[key]
			if !requested {
				continue
			}
			old := jsonSourceMemberValue(object, key)
			if old == nil {
				missing = append(missing, key)
				continue
			}
			if old.kind != jsonSourceNumber {
				return Result{}, failure(-1, Operation{}, "UFE002", "Existing geometry must be a JSON number.", "Correct the targeted geometry value and retry.")
			}
			oldValue, err := strconv.ParseFloat(old.text, 64)
			if err == nil && oldValue == value {
				continue
			}
			edits = append(edits, SourceEdit{Start: old.start, End: old.end, Text: strconv.FormatFloat(value, 'g', -1, 64)})
		}
		if len(missing) > 0 {
			edits = append(edits, jsonObjectInsertion(source, object, missing, values))
		}
	}
	slices.SortFunc(edits, func(a, b SourceEdit) int {
		if a.Start != b.Start {
			return a.Start - b.Start
		}
		if a.End != b.End {
			return a.End - b.End
		}
		return strings.Compare(a.Text, b.Text)
	})
	for i := 1; i < len(edits); i++ {
		if edits[i-1].End > edits[i].Start {
			return Result{}, failure(-1, Operation{}, "UFE009", "Generated JSON source edits overlap.", "Keep the source unchanged and report this edit.")
		}
	}

	updated := applySourceEdits(source, edits)
	after, err := spec.ParseFormSpec(input, updated)
	if err != nil {
		lastIndex, last := -1, Operation{}
		if len(operations) > 0 {
			lastIndex, last = len(operations)-1, operations[len(operations)-1]
		}
		return Result{}, validationError(lastIndex, last, err)
	}
	return Result{Source: updated, Document: after, Edits: edits, Warnings: after.ValidationWarnings}, nil
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

func jsonObjectInsertion(source []byte, object *jsonSourceNode, keys []string, values map[string]float64) SourceEdit {
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
		text.WriteString(strconv.FormatFloat(values[key], 'g', -1, 64))
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
