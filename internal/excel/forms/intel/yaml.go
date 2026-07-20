// Package intel provides syntax-aware, editor-oriented helpers for UserForm
// specification documents. It deliberately does not depend on LSP types.
package intel

import (
	"fmt"
	"sort"
	"strings"
	"unicode/utf16"

	"gopkg.in/yaml.v3"
)

type Position struct {
	Line      int
	Character int
}

type Range struct {
	Start Position
	End   Position
}

type FieldNodes struct {
	Key        *yaml.Node
	Value      *yaml.Node
	KeyRange   Range
	ValueRange Range
}

// Document retains the YAML syntax tree and a field-path index for a source
// buffer. ParseError is non-nil only when the document is not syntactically
// complete; callers can still use CursorContextAt in that state.
type Document struct {
	Source     string
	Root       *yaml.Node
	Fields     map[string]FieldNodes
	ParseError error
	lines      []string
}

func ParseYAML(source string) *Document {
	doc := &Document{Source: source, Fields: make(map[string]FieldNodes), lines: splitLines(source)}
	var root yaml.Node
	if err := yaml.Unmarshal([]byte(source), &root); err != nil {
		doc.ParseError = err
		return doc
	}
	doc.Root = &root
	if len(root.Content) != 0 {
		doc.indexNode(root.Content[0], "")
	}
	return doc
}

func (d *Document) Field(path string) (FieldNodes, bool) {
	field, ok := d.Fields[path]
	return field, ok
}

func (d *Document) FieldPathForNode(node *yaml.Node) (string, bool) {
	for path, field := range d.Fields {
		if field.Key == node || field.Value == node {
			return path, true
		}
	}
	return "", false
}

func (d *Document) FieldPathAt(pos Position) (string, bool) {
	for path, field := range d.Fields {
		if contains(field.KeyRange, pos) || contains(field.ValueRange, pos) {
			return path, true
		}
	}
	return "", false
}

func (d *Document) CursorContextAt(pos Position) CursorContext {
	if path, ok := d.FieldPathAt(pos); ok {
		return CursorContext{ParentPath: parentPath(path), FieldPath: path, LinePrefix: linePrefix(d.line(pos.Line), pos.Character), ExistingKeys: d.siblingKeys(parentPath(path))}
	}
	return fallbackCursorContext(d.lines, pos)
}

func (d *Document) indexNode(node *yaml.Node, path string) {
	if node == nil {
		return
	}
	switch node.Kind {
	case yaml.MappingNode:
		for i := 0; i+1 < len(node.Content); i += 2 {
			key, value := node.Content[i], node.Content[i+1]
			fieldPath := joinPath(path, key.Value)
			d.Fields[fieldPath] = FieldNodes{
				Key: key, Value: value, KeyRange: d.rangeForKey(key), ValueRange: d.rangeForValue(value),
			}
			d.indexNode(value, fieldPath)
		}
	case yaml.SequenceNode:
		for index, value := range node.Content {
			d.indexNode(value, fmt.Sprintf("%s[%d]", path, index))
		}
	}
}

func (d *Document) siblingKeys(path string) []string {
	prefix := path
	if prefix != "" {
		prefix += "."
	}
	keys := make([]string, 0)
	for candidate := range d.Fields {
		if !strings.HasPrefix(candidate, prefix) {
			continue
		}
		rest := strings.TrimPrefix(candidate, prefix)
		if strings.Contains(rest, ".") || strings.Contains(rest, "[") {
			continue
		}
		keys = append(keys, rest)
	}
	sort.Strings(keys)
	return keys
}

func (d *Document) rangeForKey(node *yaml.Node) Range {
	start := d.nodeStart(node)
	line := d.line(start.Line)
	startByte := byteOffsetForUTF16(line, start.Character)
	endByte := mappingKeyEnd(line, startByte)
	return Range{Start: start, End: positionForByte(line, start.Line, endByte)}
}

func (d *Document) rangeForValue(node *yaml.Node) Range {
	start := d.nodeStart(node)
	if node == nil || node.Kind != yaml.ScalarNode {
		return Range{Start: start, End: start}
	}
	line := d.line(start.Line)
	startByte := byteOffsetForUTF16(line, start.Character)
	endByte := scalarEnd(line, startByte)
	return Range{Start: start, End: positionForByte(line, start.Line, endByte)}
}

func (d *Document) nodeStart(node *yaml.Node) Position {
	if node == nil || node.Line <= 0 {
		return Position{}
	}
	lineIndex := node.Line - 1
	line := d.line(lineIndex)
	column := max(0, node.Column-1)
	if column > len(line) {
		column = len(line)
	}
	return positionForByte(line, lineIndex, column)
}

func (d *Document) line(index int) string {
	if index < 0 || index >= len(d.lines) {
		return ""
	}
	return d.lines[index]
}

type CursorContext struct {
	ParentPath    string
	FieldPath     string
	SequenceIndex int
	ExistingKeys  []string
	LinePrefix    string
}

func fallbackCursorContext(lines []string, pos Position) CursorContext {
	lineIndex := min(max(0, pos.Line), len(lines)-1)
	if len(lines) == 0 {
		return CursorContext{}
	}
	type mapping struct {
		indent int
		path   string
		keys   map[string]struct{}
	}
	type sequence struct {
		indent int
		path   string
		next   int
		index  int
	}
	mappings := []mapping{{indent: -1, path: "", keys: map[string]struct{}{}}}
	sequences := []sequence{}
	for i := 0; i <= lineIndex; i++ {
		line := lines[i]
		indent := len(line) - len(strings.TrimLeft(line, " "))
		trimmed := strings.TrimSpace(line)
		if trimmed == "" || strings.HasPrefix(trimmed, "#") {
			continue
		}
		isSequenceItem := strings.HasPrefix(trimmed, "-") && (len(trimmed) == 1 || trimmed[1] == ' ')
		// YAML permits a sequence entry at the same indentation as the key whose
		// value is that sequence (`controls:\n- ...`). Preserve that mapping as
		// the sequence owner before normal same-indent mapping unwinding.
		sameIndentSequenceOwner := ""
		if isSequenceItem && len(mappings) > 1 && mappings[len(mappings)-1].indent == indent {
			sameIndentSequenceOwner = mappings[len(mappings)-1].path
		}
		for len(mappings) > 1 && mappings[len(mappings)-1].indent >= indent {
			mappings = mappings[:len(mappings)-1]
		}
		for len(sequences) > 0 && sequences[len(sequences)-1].indent > indent {
			sequences = sequences[:len(sequences)-1]
		}
		if isSequenceItem {
			seqIndex := len(sequences) - 1
			if seqIndex < 0 || sequences[seqIndex].indent != indent {
				parent := mappings[len(mappings)-1].path
				if sameIndentSequenceOwner != "" {
					parent = sameIndentSequenceOwner
				}
				sequences = append(sequences, sequence{indent: indent, path: parent, index: -1})
				seqIndex = len(sequences) - 1
			}
			sequences[seqIndex].index = sequences[seqIndex].next
			sequences[seqIndex].next++
			itemPath := fmt.Sprintf("%s[%d]", sequences[seqIndex].path, sequences[seqIndex].index)
			mappings = append(mappings, mapping{indent: indent, path: itemPath, keys: map[string]struct{}{}})
			trimmed = strings.TrimSpace(strings.TrimPrefix(trimmed, "-"))
			if trimmed == "" {
				continue
			}
		}
		key, hasColon := yamlKey(trimmed)
		if !hasColon {
			continue
		}
		current := &mappings[len(mappings)-1]
		current.keys[key] = struct{}{}
		value := strings.TrimSpace(strings.TrimPrefix(trimmed[len(key):], ":"))
		if value == "" {
			childPath := joinPath(current.path, key)
			mappings = append(mappings, mapping{indent: indent, path: childPath, keys: map[string]struct{}{}})
		}
	}
	current := mappings[len(mappings)-1]
	keys := make([]string, 0, len(current.keys))
	for key := range current.keys {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	sequenceIndex := -1
	if len(sequences) > 0 {
		sequenceIndex = sequences[len(sequences)-1].index
	}
	return CursorContext{ParentPath: current.path, SequenceIndex: sequenceIndex, ExistingKeys: keys, LinePrefix: linePrefix(lines[lineIndex], pos.Character)}
}

func yamlKey(line string) (string, bool) {
	colon := strings.IndexByte(line, ':')
	if colon <= 0 {
		return "", false
	}
	key := strings.TrimSpace(line[:colon])
	return key, key != ""
}

func joinPath(parent, key string) string {
	if parent == "" {
		return key
	}
	return parent + "." + key
}

func parentPath(path string) string {
	if index := strings.LastIndex(path, "."); index >= 0 {
		return path[:index]
	}
	return ""
}

func contains(r Range, pos Position) bool {
	return (pos.Line > r.Start.Line || pos.Line == r.Start.Line && pos.Character >= r.Start.Character) &&
		(pos.Line < r.End.Line || pos.Line == r.End.Line && pos.Character <= r.End.Character)
}

func splitLines(source string) []string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	return strings.Split(source, "\n")
}

func positionForByte(line string, lineIndex, byteOffset int) Position {
	byteOffset = min(max(0, byteOffset), len(line))
	return Position{Line: lineIndex, Character: len(utf16.Encode([]rune(line[:byteOffset])))}
}

func byteOffsetForUTF16(line string, character int) int {
	if character <= 0 {
		return 0
	}
	units := 0
	for offset, r := range line {
		next := units + len(utf16.Encode([]rune{r}))
		if next > character {
			return offset
		}
		units = next
	}
	return len(line)
}

func mappingKeyEnd(line string, start int) int {
	quoted := start < len(line) && (line[start] == '\'' || line[start] == '"')
	if quoted {
		return quotedEnd(line, start)
	}
	if colon := strings.IndexByte(line[start:], ':'); colon >= 0 {
		return start + colon
	}
	return len(strings.TrimRight(line, " \t"))
}

func scalarEnd(line string, start int) int {
	if start >= len(line) {
		return start
	}
	if line[start] == '\'' || line[start] == '"' {
		return quotedEnd(line, start)
	}
	end := len(line)
	if comment := strings.IndexByte(line[start:], '#'); comment >= 0 {
		end = start + comment
	}
	for end > start && (line[end-1] == ' ' || line[end-1] == '\t') {
		end--
	}
	return end
}

func quotedEnd(line string, start int) int {
	quote := line[start]
	for i := start + 1; i < len(line); i++ {
		if line[i] == quote && (quote == '\'' || i == start+1 || line[i-1] != '\\') {
			return i + 1
		}
	}
	return len(line)
}

func linePrefix(line string, character int) string {
	return line[:byteOffsetForUTF16(line, character)]
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}
