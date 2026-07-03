package vbafmt

import (
	"bytes"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type textEdit struct {
	start int
	end   int
	text  string
}

var (
	symbolBinaryOperators  = map[string]bool{"=": true, "<>": true, "<": true, "<=": true, ">": true, ">=": true, "+": true, "-": true, "*": true, "/": true, "\\": true, "^": true, "&": true}
	keywordBinaryOperators = map[string]bool{"AND": true, "OR": true, "XOR": true, "EQV": true, "IMP": true, "IS": true, "LIKE": true, "MOD": true}
)

func formatOperatorSpacing(text string, isClass bool) (string, error) {
	parser, err := vbaast.NewParser()
	if err != nil {
		return "", err
	}
	defer parser.Close()

	source := []byte(text)
	parsed := parser.Parse("<fmt>", source)
	defer parsed.Close()

	if parsed.HasError || parsed.HasMissing {
		return "", formatParseError{hasError: parsed.HasError, hasMissing: parsed.HasMissing}
	}

	lines := splitLines(text)
	lineStarts := lineStartOffsets(text, lines)
	unsafe := unsafeOperatorSpacingLines(lines, isClass)
	editByRange := map[[2]int]textEdit{}

	vbaast.Walk(parsed.Root, func(node *tree_sitter.Node) bool {
		for _, edit := range operatorSpacingEditsForNode(node, source, lineStarts, unsafe) {
			editByRange[[2]int{edit.start, edit.end}] = edit
		}
		return true
	})
	for _, edit := range lexicalOperatorSpacingEdits(source, lines, lineStarts, unsafe) {
		editByRange[[2]int{edit.start, edit.end}] = edit
	}

	if len(editByRange) == 0 {
		return text, nil
	}
	edits := make([]textEdit, 0, len(editByRange))
	for _, edit := range editByRange {
		edits = append(edits, edit)
	}
	sort.SliceStable(edits, func(i, j int) bool {
		return edits[i].start > edits[j].start
	})

	out := []byte(text)
	for _, edit := range edits {
		out = append(out[:edit.start], append([]byte(edit.text), out[edit.end:]...)...)
	}
	return string(out), nil
}

func operatorSpacingEditsForNode(node *tree_sitter.Node, source []byte, lineStarts []int, unsafe []bool) []textEdit {
	if node == nil {
		return nil
	}
	var opStart, opEnd int
	var ok bool

	switch node.Kind() {
	case "assignment_statement", "set_statement":
		opStart, opEnd, ok = operatorBetweenFields(node, source, "left", "right", map[string]bool{"=": true})
	case "comparison_expression":
		op := node.ChildByFieldName("operator")
		if op == nil {
			return nil
		}
		text := strings.TrimSpace(op.Utf8Text(source))
		ok = safeOperatorText(text, symbolBinaryOperators[text], keywordAllowed(text, keywordBinaryOperators))
		opStart = int(op.StartByte())
		opEnd = int(op.EndByte())
	case "binary_expression":
		opStart, opEnd, ok = operatorBetweenFirstNamedChildren(node, source, symbolBinaryOperators, keywordBinaryOperators)
	case "condition_binary_expression", "logical_value_expression":
		opStart, opEnd, ok = operatorBetweenFirstNamedChildren(node, source, nil, map[string]bool{"AND": true, "OR": true})
	default:
		return nil
	}
	if !ok || opStart >= opEnd {
		return nil
	}
	if !operatorRangeIsSafe(source, opStart, opEnd, lineStarts, unsafe) {
		return nil
	}
	return spacingEditsAroundOperator(source, opStart, opEnd)
}

func operatorBetweenFields(node *tree_sitter.Node, source []byte, leftField, rightField string, symbols map[string]bool) (int, int, bool) {
	left := node.ChildByFieldName(leftField)
	right := node.ChildByFieldName(rightField)
	if left == nil || right == nil {
		return 0, 0, false
	}
	return operatorBetween(source, int(left.EndByte()), int(right.StartByte()), symbols, nil)
}

func operatorBetweenFirstNamedChildren(node *tree_sitter.Node, source []byte, symbols, keywords map[string]bool) (int, int, bool) {
	if node.NamedChildCount() < 2 {
		return 0, 0, false
	}
	left := node.NamedChild(0)
	right := node.NamedChild(1)
	if left == nil || right == nil {
		return 0, 0, false
	}
	return operatorBetween(source, int(left.EndByte()), int(right.StartByte()), symbols, keywords)
}

func operatorBetween(source []byte, start, end int, symbols, keywords map[string]bool) (int, int, bool) {
	if start < 0 || end > len(source) || start >= end {
		return 0, 0, false
	}
	left := start
	for left < end && isHorizontalWhitespace(source[left]) {
		left++
	}
	right := end
	for right > left && isHorizontalWhitespace(source[right-1]) {
		right--
	}
	if left >= right || bytes.ContainsAny(source[left:right], "\r\n") {
		return 0, 0, false
	}
	op := string(source[left:right])
	if safeOperatorText(op, symbols != nil && symbols[op], keywordAllowed(op, keywords)) {
		return left, right, true
	}
	return 0, 0, false
}

func safeOperatorText(op string, symbolAllowed, keywordAllowed bool) bool {
	if symbolAllowed {
		return true
	}
	return keywordAllowed
}

func keywordAllowed(op string, keywords map[string]bool) bool {
	if len(keywords) == 0 {
		return false
	}
	return keywords[strings.ToUpper(op)]
}

func operatorRangeIsSafe(source []byte, start, end int, lineStarts []int, unsafe []bool) bool {
	line := lineForOffset(lineStarts, start)
	if line < 0 || line >= len(unsafe) || unsafe[line] {
		return false
	}
	if lineForOffset(lineStarts, end) != line {
		return false
	}
	lineStart := lineStarts[line]
	lineEnd := len(source)
	if line+1 < len(lineStarts) {
		lineEnd = lineStarts[line+1]
		for lineEnd > lineStart && (source[lineEnd-1] == '\n' || source[lineEnd-1] == '\r') {
			lineEnd--
		}
	}
	codeEnd := trailingCommentOffset(source[lineStart:lineEnd])
	if codeEnd >= 0 && start >= lineStart+codeEnd {
		return false
	}
	return true
}

func spacingEditsAroundOperator(source []byte, opStart, opEnd int) []textEdit {
	left := opStart
	for left > 0 && isHorizontalWhitespace(source[left-1]) {
		left--
	}
	right := opEnd
	for right < len(source) && isHorizontalWhitespace(source[right]) {
		right++
	}
	edits := make([]textEdit, 0, 2)
	if left != opStart || opStart == 0 || source[opStart-1] != ' ' {
		edits = append(edits, textEdit{start: left, end: opStart, text: " "})
	}
	if right != opEnd || opEnd >= len(source) || source[opEnd] != ' ' {
		edits = append(edits, textEdit{start: opEnd, end: right, text: " "})
	}
	return edits
}

func lexicalOperatorSpacingEdits(source []byte, lines []string, lineStarts []int, unsafe []bool) []textEdit {
	var edits []textEdit
	for lineNo, line := range lines {
		if lineNo >= len(unsafe) || unsafe[lineNo] {
			continue
		}
		lineStart := lineStarts[lineNo]
		comment := trailingCommentOffset([]byte(line))
		lineEnd := len(line)
		if comment >= 0 {
			lineEnd = comment
		}
		inString := false
		for i := 0; i < lineEnd; i++ {
			ch := line[i]
			if ch == '"' {
				if inString && i+1 < lineEnd && line[i+1] == '"' {
					i++
					continue
				}
				inString = !inString
				continue
			}
			if inString || isHorizontalWhitespace(ch) {
				continue
			}
			if opLen := lexicalSymbolOperatorLength(line, i, lineEnd); opLen > 0 {
				if lexicalSymbolOperatorIsBinary(line, i, i+opLen, lineEnd) {
					edits = append(edits, spacingEditsAroundOperator(source, lineStart+i, lineStart+i+opLen)...)
					i += opLen - 1
				}
				continue
			}
			if isIdentifierStartByte(ch) {
				j := i + 1
				for j < lineEnd && isIdentifierPartByte(line[j]) {
					j++
				}
				word := strings.ToUpper(line[i:j])
				if keywordBinaryOperators[word] && lexicalKeywordOperatorIsBinary(line, i, j, lineEnd) {
					edits = append(edits, spacingEditsAroundOperator(source, lineStart+i, lineStart+j)...)
				}
				i = j - 1
			}
		}
	}
	return edits
}

func lexicalSymbolOperatorLength(line string, i, end int) int {
	if i >= end {
		return 0
	}
	if i+2 <= end {
		op := line[i : i+2]
		if op == "<>" || op == "<=" || op == ">=" {
			return 2
		}
	}
	switch line[i] {
	case '=', '<', '>', '+', '-', '*', '/', '\\', '^', '&':
		return 1
	default:
		return 0
	}
}

func lexicalSymbolOperatorIsBinary(line string, start, end, lineEnd int) bool {
	if start > 0 && line[start-1] == ':' && end <= lineEnd && line[start:end] == "=" {
		return false
	}
	prev := previousSignificantByte(line, start)
	next := nextSignificantByte(line, end, lineEnd)
	if !isOperandEndByte(prev) {
		return false
	}
	if isOperandStartByte(next) {
		return true
	}
	if (next == '+' || next == '-') && isOperandStartByte(nextSignificantByte(line, nextSignificantIndex(line, end, lineEnd)+1, lineEnd)) {
		return true
	}
	return false
}

func lexicalKeywordOperatorIsBinary(line string, start, end, lineEnd int) bool {
	if start > 0 && isIdentifierPartByte(line[start-1]) {
		return false
	}
	if end < lineEnd && isIdentifierPartByte(line[end]) {
		return false
	}
	return isOperandEndByte(previousSignificantByte(line, start)) && isOperandStartByte(nextSignificantByte(line, end, lineEnd))
}

func previousSignificantByte(line string, before int) byte {
	for i := before - 1; i >= 0; i-- {
		if !isHorizontalWhitespace(line[i]) {
			return line[i]
		}
	}
	return 0
}

func nextSignificantByte(line string, after, lineEnd int) byte {
	idx := nextSignificantIndex(line, after, lineEnd)
	if idx < 0 {
		return 0
	}
	return line[idx]
}

func nextSignificantIndex(line string, after, lineEnd int) int {
	for i := after; i < lineEnd; i++ {
		if !isHorizontalWhitespace(line[i]) {
			return i
		}
	}
	return -1
}

func isOperandEndByte(ch byte) bool {
	return isIdentifierPartByte(ch) || (ch >= '0' && ch <= '9') || ch == '"' || ch == ')' || ch == '!'
}

func isOperandStartByte(ch byte) bool {
	return isIdentifierStartByte(ch) || (ch >= '0' && ch <= '9') || ch == '"' || ch == '(' || ch == '.'
}

func isIdentifierStartByte(ch byte) bool {
	return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch == '_'
}

func isIdentifierPartByte(ch byte) bool {
	return isIdentifierStartByte(ch) || (ch >= '0' && ch <= '9')
}

func unsafeOperatorSpacingLines(lines []string, isClass bool) []bool {
	unsafe := make([]bool, len(lines))
	inClassBeginBlock := false
	headerEnded := !isClass
	inContinuation := false
	for i, line := range lines {
		trimmed := strings.TrimLeft(line, " \t")
		upper := strings.ToUpper(trimmed)
		directive := parseLineNumberDirective(line)
		content := strings.TrimLeft(directive.Content, " \t")
		code := stripStringLiterals(stripTrailingComment(content))
		codeTrim := strings.TrimSpace(code)

		if inContinuation {
			unsafe[i] = true
		}
		if !headerEnded {
			if inClassBeginBlock || isClassHeaderLine(line) || isBlankLine(line) {
				unsafe[i] = true
				switch upper {
				case "BEGIN":
					inClassBeginBlock = true
				case "END":
					inClassBeginBlock = false
				}
			} else {
				headerEnded = true
			}
		}
		if isBlankLine(line) || isVBACommentLine(line) || strings.HasPrefix(codeTrim, "#") || strings.ContainsAny(code, "[]") {
			unsafe[i] = true
		}
		if hasExplicitLineContinuation(content) {
			unsafe[i] = true
			inContinuation = true
		} else {
			inContinuation = false
		}
	}
	return unsafe
}

func lineStartOffsets(text string, lines []string) []int {
	starts := make([]int, len(lines))
	offset := 0
	for i, line := range lines {
		starts[i] = offset
		offset += len(line)
		if offset < len(text) {
			switch text[offset] {
			case '\r':
				offset++
				if offset < len(text) && text[offset] == '\n' {
					offset++
				}
			case '\n':
				offset++
			}
		}
	}
	return starts
}

func lineForOffset(starts []int, offset int) int {
	idx := sort.Search(len(starts), func(i int) bool {
		return starts[i] > offset
	}) - 1
	return idx
}

func trailingCommentOffset(line []byte) int {
	inString := false
	for i := 0; i < len(line); i++ {
		ch := line[i]
		if ch == '"' {
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			continue
		}
		if inString {
			continue
		}
		if ch == '\'' {
			return i
		}
		if i+3 <= len(line) && strings.EqualFold(string(line[i:i+3]), "REM") {
			if i == 0 || isHorizontalWhitespace(line[i-1]) {
				if i+3 == len(line) || isHorizontalWhitespace(line[i+3]) {
					return i
				}
			}
		}
	}
	return -1
}

func isHorizontalWhitespace(ch byte) bool {
	return ch == ' ' || ch == '\t'
}
