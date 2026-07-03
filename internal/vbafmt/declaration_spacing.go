package vbafmt

import (
	"regexp"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

var fixedStringDeclarationRe = regexp.MustCompile(`(?i)\bAs\s+String\s*\*`)

type declarationToken struct {
	text string
}

func formatDeclarationSpacing(text string, isClass bool) (string, error) {
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
	lineSet := map[int]bool{}

	vbaast.Walk(parsed.Root, func(node *tree_sitter.Node) bool {
		if declarationSpacingNodeKind(node.Kind()) {
			line := lineForOffset(lineStarts, int(node.StartByte()))
			if line >= 0 && line < len(lines) && !unsafe[line] {
				lineSet[line] = true
			}
		}
		return true
	})

	if len(lineSet) == 0 {
		return text, nil
	}

	edits := make([]textEdit, 0, len(lineSet))
	for lineNo := range lineSet {
		edit, ok := declarationSpacingEditForLine(lines[lineNo], lineStarts[lineNo])
		if ok {
			edits = append(edits, edit)
		}
	}
	if len(edits) == 0 {
		return text, nil
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

func declarationSpacingNodeKind(kind string) bool {
	switch kind {
	case "variable_declaration",
		"const_declaration",
		"sub_declaration",
		"function_declaration",
		"property_declaration",
		"property_get_declaration",
		"property_let_declaration",
		"property_set_declaration":
		return true
	default:
		return false
	}
}

func declarationSpacingEditForLine(line string, lineStart int) (textEdit, bool) {
	directive := parseLineNumberDirective(line)
	prefixLen := len(line) - len(directive.Content)
	content := directive.Content
	indent := leadingWhitespace(content)
	codeStart := prefixLen + len(indent)
	body := content[len(indent):]
	commentStart := trailingCommentOffset([]byte(body))
	if commentStart >= 0 {
		body = body[:commentStart]
	}
	body = strings.TrimRight(body, " \t")
	if !declarationSpacingLineIsSafe(body) {
		return textEdit{}, false
	}
	formatted, ok := formatDeclarationCode(body)
	if !ok || formatted == body {
		return textEdit{}, false
	}
	return textEdit{
		start: lineStart + codeStart,
		end:   lineStart + codeStart + len(body),
		text:  formatted,
	}, true
}

func declarationSpacingLineIsSafe(code string) bool {
	trimmed := strings.TrimSpace(code)
	if trimmed == "" {
		return false
	}
	upper := strings.ToUpper(stripStringLiterals(trimmed))
	if strings.HasPrefix(upper, "#") || strings.HasPrefix(upper, "ATTRIBUTE ") || strings.Contains(upper, "DECLARE ") {
		return false
	}
	if strings.ContainsAny(upper, "[]:") {
		return false
	}
	return !fixedStringDeclarationRe.MatchString(trimmed)
}

func formatDeclarationCode(code string) (string, bool) {
	tokens, ok := scanDeclarationTokens(code)
	if !ok || len(tokens) == 0 {
		return "", false
	}
	var b strings.Builder
	for i, token := range tokens {
		prev := declarationToken{}
		if i > 0 {
			prev = tokens[i-1]
		}
		appendDeclarationToken(&b, prev, token, i == 0)
	}
	return b.String(), true
}

func scanDeclarationTokens(code string) ([]declarationToken, bool) {
	tokens := make([]declarationToken, 0)
	for i := 0; i < len(code); {
		ch := code[i]
		if isHorizontalWhitespace(ch) {
			i++
			continue
		}
		if ch == '"' {
			start := i
			i++
			for i < len(code) {
				if code[i] == '"' {
					if i+1 < len(code) && code[i+1] == '"' {
						i += 2
						continue
					}
					i++
					break
				}
				i++
			}
			if i > len(code) {
				return nil, false
			}
			tokens = append(tokens, declarationToken{text: code[start:i]})
			continue
		}
		if isDeclarationIdentifierStart(ch) {
			start := i
			i++
			for i < len(code) && isDeclarationIdentifierPart(code[i]) {
				i++
			}
			tokens = append(tokens, declarationToken{text: code[start:i]})
			continue
		}
		if ch >= '0' && ch <= '9' {
			start := i
			i++
			for i < len(code) && isDeclarationNumberPart(code[i]) {
				i++
			}
			tokens = append(tokens, declarationToken{text: code[start:i]})
			continue
		}
		switch ch {
		case '(', ')', ',', '.', '=':
			tokens = append(tokens, declarationToken{text: code[i : i+1]})
			i++
		default:
			return nil, false
		}
	}
	return tokens, true
}

func appendDeclarationToken(b *strings.Builder, prev, cur declarationToken, first bool) {
	text := cur.text
	if first {
		b.WriteString(text)
		return
	}
	switch text {
	case ",":
		trimBuilderRightSpace(b)
		b.WriteString(",")
	case ")":
		trimBuilderRightSpace(b)
		b.WriteString(")")
	case "(", ".":
		trimBuilderRightSpace(b)
		b.WriteString(text)
	case "=":
		trimBuilderRightSpace(b)
		writeBuilderSpace(b)
		b.WriteString("=")
	default:
		switch prev.text {
		case "(", ".":
			b.WriteString(text)
		case ",":
			writeBuilderSpace(b)
			b.WriteString(text)
		case "=":
			writeBuilderSpace(b)
			b.WriteString(text)
		default:
			writeBuilderSpace(b)
			b.WriteString(text)
		}
	}
}

func writeBuilderSpace(b *strings.Builder) {
	if b.Len() == 0 {
		return
	}
	text := b.String()
	if text[len(text)-1] != ' ' {
		b.WriteByte(' ')
	}
}

func trimBuilderRightSpace(b *strings.Builder) {
	text := b.String()
	trimmed := strings.TrimRight(text, " ")
	if len(trimmed) == len(text) {
		return
	}
	b.Reset()
	b.WriteString(trimmed)
}

func isDeclarationIdentifierStart(ch byte) bool {
	return isIdentifierStartByte(ch)
}

func isDeclarationIdentifierPart(ch byte) bool {
	return isIdentifierPartByte(ch) || strings.ContainsRune("$%&!#@^", rune(ch))
}

func isDeclarationNumberPart(ch byte) bool {
	return (ch >= '0' && ch <= '9') || ch == '.' || strings.ContainsRune("$%&!#@^", rune(ch))
}
