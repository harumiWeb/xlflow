package vbafmt

import (
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
	_, ok := err.(formatParseError)
	return ok
}

type lineIndentModel struct {
	levels []int
}

func parseFormattingModel(text string, strict bool) (*lineIndentModel, error) {
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

	parsed := parser.Parse("<fmt>", []byte(sourceForFormatterParse(lines)))
	defer parsed.Close()

	if parsed.HasError || parsed.HasMissing {
		if !strict {
			return heuristicFormattingModel(lines), nil
		}
		return nil, formatParseError{hasError: parsed.HasError, hasMissing: parsed.HasMissing}
	}

	vbaast.Walk(parsed.Root, func(node *tree_sitter.Node) bool {
		applyNodeIndent(model, node)
		return true
	})
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

func sourceForFormatterParse(lines []string) string {
	normalized := make([]string, 0, len(lines))
	for _, line := range lines {
		directive := parseLineNumberDirective(line)
		if directive.Has {
			normalized = append(normalized, directive.Content)
		} else {
			normalized = append(normalized, line)
		}
	}
	return strings.Join(normalized, "\n")
}

func heuristicFormattingModel(lines []string) *lineIndentModel {
	model := &lineIndentModel{
		levels: make([]int, len(lines)+1),
	}
	level := 0
	headerEnded := false
	inBeginBlock := false

	for i, line := range lines {
		lineNumber := i + 1
		if !headerEnded {
			trimmedForHeader := strings.TrimLeft(line, " \t")
			upperForHeader := strings.ToUpper(trimmedForHeader)
			if inBeginBlock {
				if upperForHeader == "END" {
					inBeginBlock = false
				}
				continue
			}
			if isClassHeaderLine(line) || isBlankLine(line) {
				if upperForHeader == "BEGIN" {
					inBeginBlock = true
				}
				continue
			}
			headerEnded = true
		}

		directive := parseLineNumberDirective(line)
		trimmed := strings.TrimRight(directive.Content, " \t")
		content := strings.TrimLeft(trimmed, " \t")
		if isBlankLine(trimmed) || isVBACommentLine(trimmed) || isLabelLine(content) {
			model.levels[lineNumber] = level
			continue
		}

		keyword, isStructural := classifyLine(content)
		if isStructural && isDedentKeyword(keyword) {
			level--
			if level < 0 {
				level = 0
			}
		}
		model.levels[lineNumber] = level
		if isStructural && isIndentKeyword(keyword) {
			level++
		}
	}

	return model
}

func applyNodeIndent(model *lineIndentModel, node *tree_sitter.Node) {
	switch node.Kind() {
	case "sub_declaration",
		"function_declaration",
		"property_get_declaration",
		"property_let_declaration",
		"property_set_declaration",
		"conditional_sub_declaration",
		"conditional_function_declaration",
		"conditional_property_declaration",
		"type_declaration",
		"enum_declaration",
		"if_statement",
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
