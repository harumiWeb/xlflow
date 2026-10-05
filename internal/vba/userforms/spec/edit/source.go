package edit

import (
	"bytes"
	"fmt"
	"slices"
	"strings"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

// A piece table retains provenance through sequential operations. This makes
// the returned edits relative to the original snapshot even after insertions.
type piece struct {
	start, end int
	text       string
	inserted   bool
}
type pieceTable []piece

func (p piece) length() int {
	if p.inserted {
		return len(p.text)
	}
	return p.end - p.start
}
func (p piece) cut(a, b int) piece {
	if p.inserted {
		p.text = p.text[a:b]
	} else {
		p.end = p.start + b
		p.start += a
	}
	return p
}
func (p pieceTable) replace(edit SourceEdit) pieceTable {
	var before, after pieceTable
	pos := 0
	for _, part := range p {
		end := pos + part.length()
		if pos < edit.Start {
			before = append(before, part.cut(0, min(part.length(), edit.Start-pos)))
		}
		if end > edit.End {
			after = append(after, part.cut(max(0, edit.End-pos), part.length()))
		}
		pos = end
	}
	if edit.Text != "" {
		before = append(before, piece{text: edit.Text, inserted: true})
	}
	return append(before, after...)
}
func (p pieceTable) edits(original []byte) []SourceEdit {
	var result []SourceEdit
	cursor := 0
	var text strings.Builder
	emit := func(end int) {
		if string(original[cursor:end]) != text.String() {
			result = append(result, SourceEdit{Start: cursor, End: end, Text: text.String()})
		}
		text.Reset()
	}
	for _, part := range p {
		if part.inserted {
			text.WriteString(part.text)
			continue
		}
		if part.start > cursor || text.Len() > 0 {
			emit(part.start)
		}
		cursor = part.end
	}
	if cursor < len(original) || text.Len() > 0 {
		emit(len(original))
	}
	return result
}

func (e *engine) patch(edit SourceEdit) error {
	if edit.Start < 0 || edit.End < edit.Start || edit.End > len(e.source) {
		return fmt.Errorf("invalid generated source range")
	}
	e.pieces = e.pieces.replace(edit)
	next := make([]byte, 0, len(e.source)+len(edit.Text)-(edit.End-edit.Start))
	next = append(next, e.source[:edit.Start]...)
	next = append(next, edit.Text...)
	next = append(next, e.source[edit.End:]...)
	e.source = next
	return e.reload()
}

func (e *engine) publish(edit SourceEdit) error {
	expected, err := modelFromNode(e.root)
	if err != nil {
		// Source validation supplies the canonical UFV code before typed
		// decoding can turn a bad field type into a generic YAML error.
		body, marshalErr := yaml.Marshal(e.root)
		if marshalErr == nil {
			if _, validationErr := spec.ParseFormSpec(spec.SpecInput{Format: "yaml"}, body); validationErr != nil {
				return validationError(-1, Operation{}, validationErr)
			}
		}
		return err
	}
	if err := e.patch(edit); err != nil {
		return err
	}
	if !sameModel(expected, e.model) {
		return failure(-1, Operation{}, "UFE009", "Generated source does not match the intended FormSpec.", "Keep the source unchanged and report this edit.")
	}
	return nil
}

func (e *engine) offset(n *yaml.Node) int {
	pos := 0
	for line := 1; line < n.Line; line++ {
		i := bytes.IndexByte(e.source[pos:], '\n')
		if i < 0 {
			return len(e.source)
		}
		pos += i + 1
	}
	for column := 1; column < n.Column && pos < len(e.source); column++ {
		_, size := utf8.DecodeRune(e.source[pos:])
		pos += size
	}
	return pos
}
func (e *engine) lineStart(pos int) int { return bytes.LastIndexByte(e.source[:pos], '\n') + 1 }
func (e *engine) lineEnd(pos int) int {
	if i := bytes.IndexByte(e.source[pos:], '\n'); i >= 0 {
		return pos + i + 1
	}
	return len(e.source)
}
func (e *engine) newline() string {
	if bytes.Contains(e.source, []byte("\r\n")) {
		return "\r\n"
	}
	return "\n"
}

func (e *engine) parent(n *yaml.Node) *yaml.Node {
	var found *yaml.Node
	var walk func(*yaml.Node)
	walk = func(current *yaml.Node) {
		for _, child := range current.Content {
			if child == n {
				found = current
				return
			}
			walk(child)
			if found != nil {
				return
			}
		}
	}
	walk(e.root)
	return found
}
func (e *engine) inFlow(n *yaml.Node) bool {
	for parent := e.parent(n); parent != nil; parent = e.parent(parent) {
		if parent.Style&yaml.FlowStyle != 0 {
			return true
		}
	}
	return false
}

// end locates lexical token ends; yaml.Node only supplies starts. Columns are
// rune-based, whereas public edit offsets are bytes.
func (e *engine) end(n *yaml.Node) int {
	start := e.offset(n)
	if n.Style&yaml.FlowStyle != 0 && (n.Kind == yaml.MappingNode || n.Kind == yaml.SequenceNode) {
		return e.flowEnd(start)
	}
	if n.Kind != yaml.ScalarNode {
		end := start
		for _, child := range n.Content {
			end = max(end, e.end(child))
		}
		return end
	}
	if start >= len(e.source) {
		return start
	}
	// yaml.Node starts at an explicit tag, rather than at the scalar body.
	// Skip that property before choosing quoted/plain scanning; otherwise a
	// quoted '#' or flow delimiter would incorrectly terminate the value.
	if n.Style&(yaml.SingleQuotedStyle|yaml.DoubleQuotedStyle) != 0 && e.source[start] == '!' {
		if start+1 < len(e.source) && e.source[start+1] == '<' {
			if close := bytes.IndexByte(e.source[start+2:], '>'); close >= 0 {
				start += close + 3
			}
		} else {
			for start < len(e.source) && !strings.ContainsRune(" \t\r\n,]}", rune(e.source[start])) {
				start++
			}
		}
		for start < len(e.source) && strings.ContainsRune(" \t\r\n", rune(e.source[start])) {
			start++
		}
		if start >= len(e.source) {
			return start
		}
	}
	quote := e.source[start]
	if quote == '\'' || quote == '"' {
		for i := start + 1; i < len(e.source); i++ {
			if quote == '"' && e.source[i] == '\\' {
				i++
				continue
			}
			if e.source[i] == quote {
				if quote == '\'' && i+1 < len(e.source) && e.source[i+1] == '\'' {
					i++
					continue
				}
				return i + 1
			}
		}
		return len(e.source)
	}
	if n.Style&(yaml.LiteralStyle|yaml.FoldedStyle) != 0 {
		parent := e.parent(n)
		indent := 0
		if parent != nil {
			indent = max(0, parent.Column-1)
		}
		pos := e.lineEnd(start)
		end := pos
		for pos < len(e.source) {
			next := e.lineEnd(pos)
			line := e.source[pos:next]
			trim := bytes.TrimSpace(line)
			spaces := len(line) - len(bytes.TrimLeft(line, " "))
			if len(trim) > 0 && spaces <= indent {
				break
			}
			end = next
			pos = next
		}
		// Keep the final line separator outside the replacement.
		if end > start && e.source[end-1] == '\n' {
			end--
			if end > start && e.source[end-1] == '\r' {
				end--
			}
		}
		return end
	}
	end := e.lineEnd(start)
	flow := e.inFlow(n)
	for i := start; i < end; i++ {
		b := e.source[i]
		if b == '\n' || b == '\r' || (b == '#' && (i == start || e.source[i-1] == ' ' || e.source[i-1] == '\t')) || (flow && (b == ',' || b == ']' || b == '}')) {
			end = i
			break
		}
	}
	// Plain scalars can continue on indented subsequent lines.
	if !flow && strings.Contains(n.Value, " ") {
		parent := e.parent(n)
		indent := 0
		if parent != nil {
			indent = parent.Column - 1
		}
		pos := e.lineEnd(start)
		for pos < len(e.source) {
			next := e.lineEnd(pos)
			line := e.source[pos:next]
			trim := bytes.TrimSpace(line)
			spaces := len(line) - len(bytes.TrimLeft(line, " "))
			if len(trim) == 0 || spaces <= indent || trim[0] == '#' {
				break
			}
			end = next
			pos = next
		}
	}
	for end > start && (e.source[end-1] == ' ' || e.source[end-1] == '\t' || e.source[end-1] == '\r' || e.source[end-1] == '\n') {
		end--
	}
	return end
}

func (e *engine) flowEnd(start int) int {
	depth := 0
	quote := byte(0)
	for i := start; i < len(e.source); i++ {
		b := e.source[i]
		if quote != 0 {
			if quote == '"' && b == '\\' {
				i++
				continue
			}
			if b == quote {
				if quote == '\'' && i+1 < len(e.source) && e.source[i+1] == '\'' {
					i++
					continue
				}
				quote = 0
			}
			continue
		}
		if b == '\'' || b == '"' {
			quote = b
			continue
		}
		if b == '#' && (i == start || e.source[i-1] == ' ' || e.source[i-1] == '\n') {
			i = e.lineEnd(i) - 1
			continue
		}
		if b == '[' || b == '{' {
			depth++
		}
		if b == ']' || b == '}' {
			depth--
			if depth == 0 {
				return i + 1
			}
		}
	}
	return len(e.source)
}

func encode(n *yaml.Node, indent int, newline string) (string, error) {
	var out bytes.Buffer
	encoder := yaml.NewEncoder(&out)
	encoder.SetIndent(max(2, indent))
	if err := encoder.Encode(n); err != nil {
		return "", err
	}
	if err := encoder.Close(); err != nil {
		return "", err
	}
	return strings.ReplaceAll(strings.TrimSuffix(out.String(), "\n"), "\n", newline), nil
}

func encodeContainer(n *yaml.Node, newline string) (string, error) {
	// These comments lie outside the container's lexical range and remain in
	// the original source. Retaining them in the encoder would duplicate them.
	local := *n
	local.HeadComment, local.LineComment, local.FootComment = "", "", ""
	return encode(&local, 2, newline)
}

func (e *engine) set(mapping *yaml.Node, key string, value *yaml.Node) error {
	if mapping == nil || mapping.Kind != yaml.MappingNode {
		return fmt.Errorf("target is not a mapping")
	}
	if e.unsafeTarget(mapping) {
		return unsupported()
	}
	old := field(mapping, key)
	if unsafe(old) {
		return unsupported()
	}
	if old != nil && equalNode(old, value) {
		return nil
	}
	// Retain legal quoting style; let yaml.v3 add quotes when plain text would
	// otherwise be interpreted as a boolean, number, null or YAML syntax.
	if old != nil && value.Kind == yaml.ScalarNode && value.Tag == "!!str" {
		value.Style = old.Style & (yaml.SingleQuotedStyle | yaml.DoubleQuotedStyle)
		if strings.ContainsAny(value.Value, "\r\n\u0085\u2028\u2029") {
			value.Style = yaml.DoubleQuotedStyle
		}
	}
	if old != nil && old.Kind == yaml.ScalarNode && value.Kind == yaml.ScalarNode {
		text, err := encode(value, 2, e.newline())
		if err != nil {
			return err
		}
		edit := SourceEdit{Start: e.offset(old), End: e.end(old), Text: text}
		if old.LineComment != "" && edit.End > e.lineEnd(edit.Start) {
			// A block/plain multiline scalar's header comment is inside its
			// replacement range; relocate it onto the new single-line scalar.
			edit.Text += " " + old.LineComment
		}
		*old = *value
		return e.publish(edit)
	}
	if mapping.Style&yaml.FlowStyle != 0 {
		if unsafe(mapping) {
			return unsupported()
		}
		start, end := e.offset(mapping), e.end(mapping)
		if old != nil {
			*old = *value
		} else {
			mapping.Content = append(mapping.Content, &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: key}, value)
		}
		text, err := encodeContainer(mapping, e.newline())
		if err != nil {
			return err
		}
		return e.publish(SourceEdit{Start: start, End: end, Text: text})
	}
	if old != nil {
		return fmt.Errorf("cannot replace a non-scalar authoring field")
	}
	indent := max(0, mapping.Column-1)
	text, err := encode(&yaml.Node{Kind: yaml.MappingNode, Content: []*yaml.Node{{Kind: yaml.ScalarNode, Tag: "!!str", Value: key}, value}}, 2, e.newline())
	if err != nil {
		return err
	}
	padding := strings.Repeat(" ", indent)
	text = padding + strings.ReplaceAll(text, e.newline(), e.newline()+padding)
	pos := e.lineEnd(e.end(mapping))
	if pos > 0 && e.source[pos-1] != '\n' {
		text = e.newline() + text
	} else {
		text += e.newline()
	}
	mapping.Content = append(mapping.Content, &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: key}, value)
	return e.publish(SourceEdit{Start: pos, End: pos, Text: text})
}

func (e *engine) appendControl(node *yaml.Node) error {
	seq := field(e.root, "controls")
	if seq == nil || seq.Kind != yaml.SequenceNode {
		return fmt.Errorf("controls must be a sequence")
	}
	if seq.Anchor != "" {
		return unsupported()
	}
	if seq.Style&yaml.FlowStyle != 0 {
		start, end := e.offset(seq), e.end(seq)
		seq.Content = append(seq.Content, node)
		if unsafe(seq) {
			return unsupported()
		}
		text, err := encodeContainer(seq, e.newline())
		if err != nil {
			return err
		}
		return e.publish(SourceEdit{Start: start, End: end, Text: text})
	}
	indent := max(0, seq.Column-1)
	text, err := encode(&yaml.Node{Kind: yaml.SequenceNode, Content: []*yaml.Node{node}}, 2, e.newline())
	if err != nil {
		return err
	}
	padding := strings.Repeat(" ", indent)
	text = padding + strings.ReplaceAll(text, e.newline(), e.newline()+padding)
	pos := e.lineEnd(e.end(seq))
	if pos > 0 && e.source[pos-1] != '\n' {
		text = e.newline() + text
	} else {
		text += e.newline()
	}
	seq.Content = append(seq.Content, node)
	return e.publish(SourceEdit{Start: pos, End: pos, Text: text})
}

func (e *engine) removeItem(c controlRef) error {
	seq := c.sequence
	if seq.Anchor != "" {
		return unsupported()
	}
	index := slices.Index(seq.Content, c.node)
	if index < 0 {
		return fmt.Errorf("control sequence item not found")
	}
	if seq.Style&yaml.FlowStyle != 0 {
		start, end := e.offset(seq), e.end(seq)
		seq.Content = slices.Delete(seq.Content, index, index+1)
		if unsafe(seq) {
			return unsupported()
		}
		text, err := encodeContainer(seq, e.newline())
		if err != nil {
			return err
		}
		return e.publish(SourceEdit{Start: start, End: end, Text: text})
	}
	start := e.lineStart(e.offset(c.node))
	end := e.lineEnd(e.end(c.node))
	if c.node.HeadComment != "" {
		// HeadComment belongs to this item according to the YAML parser. Never
		// consume comments outside the item's own indentation scope.
		count := strings.Count(c.node.HeadComment, "\n") + 1
		for count > 0 && start > 0 {
			previous := e.lineStart(start - 1)
			line := bytes.TrimSpace(e.source[previous:start])
			if len(line) == 0 {
				start = previous
				continue
			}
			if line[0] != '#' {
				break
			}
			start = previous
			count--
		}
	}
	if len(seq.Content) == 1 {
		// Leave the empty sequence on the original sequence indentation line.
		text := strings.Repeat(" ", max(0, seq.Column-1)) + "[]"
		if end > 0 && e.source[end-1] == '\n' {
			text += e.newline()
		}
		seq.Content = nil
		return e.publish(SourceEdit{Start: start, End: end, Text: text})
	}
	seq.Content = slices.Delete(seq.Content, index, index+1)
	return e.publish(SourceEdit{Start: start, End: end})
}

func (e *engine) removeControl(op Operation) error {
	remove := map[string]bool{op.ControlID: true}
	for changed := true; changed; {
		changed = false
		for _, c := range e.model.Controls {
			if remove[c.ParentID] && !remove[c.ID] {
				remove[c.ID] = true
				changed = true
			}
		}
	}
	if len(remove) > 1 && !op.Cascade {
		return failure(-1, op, "UFE007", "The control has descendants.", "Use cascade=true or remove/reparent descendants first.")
	}
	// Children first, so removing a legacy nested parent cannot invalidate
	// source references still needed for separate flat descendants.
	controls := slices.Clone(e.model.Controls)
	for i := len(controls) - 1; i >= 0; i-- {
		id := controls[i].ID
		if !remove[id] {
			continue
		}
		c, ok := e.controls[id]
		if !ok {
			continue
		}
		if unsafe(c.node) {
			return unsupported()
		}
		if err := e.removeItem(c); err != nil {
			return err
		}
	}
	return nil
}

func (e *engine) setParent(op Operation) error {
	c := e.controls[op.ControlID]
	if c.model.ParentID != op.ParentID {
		if c.parent != "" {
			if unsafe(c.node) {
				return unsupported()
			}
			// Move the legacy nested item to the flat canonical collection. The
			// item's own subtree is serialized locally; other items remain raw.
			var node yaml.Node
			body, err := yaml.Marshal(c.node)
			if err != nil {
				return err
			}
			if err := yaml.Unmarshal(body, &node); err != nil {
				return err
			}
			moved := node.Content[0]
			setNode(moved, "parentId", op.ParentID)
			if field(moved, "zIndex") == nil && c.model.ZIndex != nil {
				z, _ := scalar(*c.model.ZIndex)
				moved.Content = append(moved.Content, &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: "zIndex"}, z)
			}
			if err := e.removeItem(c); err != nil {
				return err
			}
			if err := e.appendControl(moved); err != nil {
				return err
			}
		} else if err := e.setControl(op.ControlID, "parentId", op.ParentID); err != nil {
			return err
		}
	}
	if op.Left != nil {
		if err := e.setControl(op.ControlID, "left", *op.Left); err != nil {
			return err
		}
	}
	if op.Top != nil {
		return e.setControl(op.ControlID, "top", *op.Top)
	}
	return nil
}

func setNode(mapping *yaml.Node, key, value string) {
	n, _ := scalar(value)
	if old := field(mapping, key); old != nil {
		*old = *n
	} else {
		mapping.Content = append(mapping.Content, &yaml.Node{Kind: yaml.ScalarNode, Tag: "!!str", Value: key}, n)
	}
}

// modelFromNode is used by tests and validation to inspect intended YAML
// without importing editor protocols or loading assets from the filesystem.
func modelFromNode(root *yaml.Node) (spec.FormSpec, error) {
	var model spec.FormSpec
	if err := root.Decode(&model); err != nil {
		return model, err
	}
	return spec.NormalizeFormSpec(model), nil
}
