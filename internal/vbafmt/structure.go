package vbafmt

import (
	"errors"
	"fmt"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type formatParseError struct {
	hasError   bool
	hasMissing bool
}

func (e formatParseError) Error() string {
	return fmt.Sprintf("VBA parser reported errors or missing nodes (error=%t, missing=%t)", e.hasError, e.hasMissing)
}

func isFormatParseError(err error) bool {
	return IsFormatParseError(err)
}

// IsFormatParseError reports whether err means formatting was skipped because
// the VBA parser found an incomplete or invalid syntax tree.
func IsFormatParseError(err error) bool {
	var target formatParseError
	return errors.As(err, &target)
}

type lineIndentModel struct {
	levels []int
}

func parseFormattingModel(text string) (*lineIndentModel, error) {
	lines := splitLines(text)
	model := &lineIndentModel{
		levels: make([]int, len(lines)+1),
	}
	if len(lines) == 0 {
		return model, nil
	}

	parser, err := vbaast.NewParser()
	if err != nil {
		return nil, err
	}
	defer parser.Close()

	parsed := parser.Parse("<fmt>", []byte(text))
	defer parsed.Close()

	if parsed.HasError || parsed.HasMissing {
		return nil, formatParseError{hasError: parsed.HasError, hasMissing: parsed.HasMissing}
	}

	ifEvents := make([]flatIfEvent, 0)
	vbaast.Walk(parsed.Root, func(node *tree_sitter.Node) bool {
		applyNodeIndent(model, node)
		switch node.Kind() {
		case "if_statement", "elseif_fragment", "else_fragment", "end_if_fragment":
			ifEvents = append(ifEvents, flatIfEvent{
				kind:  node.Kind(),
				start: startLine(node),
				end:   int(node.EndPosition().Row) + 1,
			})
		}
		return true
	})
	applyFlatIfIndent(model, lines, ifEvents)
	return model, nil
}

func (m *lineIndentModel) level(line int) int {
	if m == nil || line < 1 || line >= len(m.levels) {
		return 0
	}
	if m.levels[line] < 0 {
		return 0
	}
	return m.levels[line]
}

func (m *lineIndentModel) formatLine(line int, content string) string {
	return strings.Repeat(" ", m.level(line)*indentWidth) + content
}

func applyNodeIndent(model *lineIndentModel, node *tree_sitter.Node) {
	switch node.Kind() {
	case "sub_declaration",
		"function_declaration",
		"property_get_declaration",
		"property_let_declaration",
		"property_set_declaration",
		"property_declaration",
		"conditional_sub_declaration",
		"conditional_function_declaration",
		"conditional_property_declaration",
		"type_declaration",
		"enum_declaration",
		"select_statement",
		"for_statement",
		"for_each_statement",
		"do_statement",
		"while_statement",
		"with_statement":
		addNodeInteriorIndent(model, node)
	case "elseif_clause", "else_clause", "case_clause":
		addLineIndent(model, startLine(node), -1)
	}
}

type flatIfEvent struct {
	kind  string
	start int
	end   int
}

type flatIfConditionalFrame struct {
	baseline int
	branches []int
	sawElse  bool
}

func applyFlatIfIndent(model *lineIndentModel, lines []string, events []flatIfEvent) {
	eventsByLine := make(map[int]flatIfEvent, len(events))
	for _, event := range events {
		eventsByLine[event.start] = event
	}

	depth := 0
	frames := make([]flatIfConditionalFrame, 0)
	pendingOpenEnds := make(map[int]int)
	for index, line := range lines {
		lineNumber := index + 1
		lower := strings.ToLower(strings.TrimSpace(stripTrailingComment(line)))
		switch {
		case strings.HasPrefix(lower, "#if ") && strings.HasSuffix(lower, " then"):
			addLineIndent(model, lineNumber, depth)
			frames = append(frames, flatIfConditionalFrame{baseline: depth})
			continue
		case strings.HasPrefix(lower, "#elseif ") && strings.HasSuffix(lower, " then"):
			if len(frames) == 0 {
				continue
			}
			frame := &frames[len(frames)-1]
			frame.branches = append(frame.branches, depth)
			depth = frame.baseline
			addLineIndent(model, lineNumber, depth)
			continue
		case lower == "#else":
			if len(frames) == 0 {
				continue
			}
			frame := &frames[len(frames)-1]
			frame.branches = append(frame.branches, depth)
			frame.sawElse = true
			depth = frame.baseline
			addLineIndent(model, lineNumber, depth)
			continue
		case lower == "#end if":
			if len(frames) == 0 {
				continue
			}
			frame := frames[len(frames)-1]
			frames = frames[:len(frames)-1]
			frame.branches = append(frame.branches, depth)
			if !frame.sawElse {
				frame.branches = append(frame.branches, frame.baseline)
			}
			addLineIndent(model, lineNumber, frame.baseline)
			depth = mergedFlatIfDepth(frame.branches)
			continue
		}

		event := eventsByLine[lineNumber]
		switch event.kind {
		case "if_statement":
			addLineIndent(model, lineNumber, depth)
			if event.end <= lineNumber {
				depth++
			} else {
				pendingOpenEnds[event.end]++
			}
		case "elseif_fragment", "else_fragment":
			addLineIndent(model, lineNumber, max(depth-1, 0))
		case "end_if_fragment":
			depth = max(depth-1, 0)
			addLineIndent(model, lineNumber, depth)
		default:
			addLineIndent(model, lineNumber, depth)
		}
		depth += pendingOpenEnds[lineNumber]
	}
}

func mergedFlatIfDepth(branches []int) int {
	if len(branches) == 0 {
		return 0
	}
	depth := branches[0]
	for _, branch := range branches[1:] {
		if branch < depth {
			depth = branch
		}
	}
	return depth
}

func addNodeInteriorIndent(model *lineIndentModel, node *tree_sitter.Node) {
	r := vbaast.NodeRange(node)
	addLineRangeIndent(model, r.StartLine+1, r.EndLine-1, 1)
}

func startLine(node *tree_sitter.Node) int {
	return int(node.StartPosition().Row) + 1
}

func addLineIndent(model *lineIndentModel, line int, delta int) {
	addLineRangeIndent(model, line, line, delta)
}

func addLineRangeIndent(model *lineIndentModel, start, end, delta int) {
	if model == nil || len(model.levels) == 0 {
		return
	}
	if start < 1 {
		start = 1
	}
	last := len(model.levels) - 1
	if end > last {
		end = last
	}
	if start > end {
		return
	}
	for line := start; line <= end; line++ {
		model.levels[line] += delta
		if model.levels[line] < 0 {
			model.levels[line] = 0
		}
	}
}
