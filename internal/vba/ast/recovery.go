package ast

import (
	"bytes"
	"regexp"
	"strings"

	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

var declarationLinePrefix = regexp.MustCompile(`(?i)^[\t ]*(?:Dim|ReDim)\b`)

var declarationKeywordAfterComma = regexp.MustCompile(`(?i),[\t ]*((?:Dim|ReDim)\b)`)

var declarationWithTypeCharacter = regexp.MustCompile(`(?i)^\s*(?:Dim|ReDim|Static|Public|Private|Friend|Const)\b`)

var optionalNumericDefault = regexp.MustCompile(`(?i)\bOptional\b[^\r\n=]*=[\t ]*([+-]?[\t ]*(?:(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|&H[0-9A-F]+|&O[0-7]+)[%&!#@^]?)`)

// IsDeclarationKeywordRecovery reports the narrow recovery shape introduced
// when tree-sitter-vba rejects a second Dim/ReDim keyword after a declaration
// comma. It deliberately requires every recovery node to be on one of those
// source lines, so ordinary parser failures keep their existing behavior.
func IsDeclarationKeywordRecovery(root *tree_sitter.Node, source []byte) bool {
	if root == nil || (!root.HasError() && !HasMissing(root)) {
		return false
	}
	spans := declarationKeywordRecoverySpans(source)
	if len(spans) == 0 {
		return false
	}
	return recoveryNodesOverlapSpans(root, spans)
}

// IsIdentifierTypeCharacterRecovery recognizes the v0.12.0 CST shape for a
// legal declaration whose Single (!) suffix is recovered as an error node.
// The source-level suffix check is deliberately tied to a variable declarator
// so an unrelated bang expression cannot suppress parser diagnostics.
func IsIdentifierTypeCharacterRecovery(root *tree_sitter.Node, source []byte) bool {
	if root == nil || (!root.HasError() && !HasMissing(root)) {
		return false
	}
	spans := make([]byteSpan, 0)
	sourceLines := strings.Split(string(source), "\n")
	Walk(root, func(node *tree_sitter.Node) bool {
		if node.Kind() != "variable_declarator" {
			return true
		}
		end := int(node.EndByte())
		if end >= 0 && end < len(source) && source[end] == '!' {
			lineNo := NodeRange(node).StartLine
			if lineNo < 1 || lineNo > len(sourceLines) || !declarationWithTypeCharacter.MatchString(strings.TrimSuffix(sourceLines[lineNo-1], "\r")) {
				return true
			}
			spanEnd := end + 1
			// In a multi-declarator line, v0.12.0 attaches the recovery ERROR
			// to the comma immediately after ! instead of to the suffix itself.
			// Include only that declaration delimiter, not arbitrary trailing
			// syntax on the same line.
			for spanEnd < len(source) && (source[spanEnd] == ' ' || source[spanEnd] == '\t') {
				spanEnd++
			}
			if spanEnd < len(source) && source[spanEnd] == ',' {
				spanEnd++
			}
			spans = append(spans, byteSpan{start: end, end: spanEnd})
		}
		return true
	})
	if len(spans) == 0 {
		return false
	}
	return recoveryNodesOverlapSpans(root, spans)
}

// IsNumericLiteralRecovery recognizes the narrow CST recovery shape produced
// when a legal VBA numeric literal in an Optional default uses a radix prefix
// or numeric type-declaration suffix that the parser represents as an ERROR
// node. It is deliberately limited to the literal span so unrelated parser
// failures remain fail-closed.
func IsNumericLiteralRecovery(root *tree_sitter.Node, source []byte) bool {
	if root == nil || (!root.HasError() && !HasMissing(root)) {
		return false
	}
	spans := make([]byteSpan, 0)
	for _, match := range optionalNumericDefault.FindAllSubmatchIndex(source, -1) {
		if len(match) >= 4 && match[2] >= 0 && match[3] > match[2] {
			spans = append(spans, byteSpan{start: match[2], end: match[3]})
		}
	}
	if len(spans) == 0 {
		return false
	}
	return recoveryNodesOverlapSpans(root, spans)
}

type byteSpan struct {
	start int
	end   int
}

func declarationKeywordRecoverySpans(source []byte) []byteSpan {
	spans := make([]byteSpan, 0)
	for lineStart := 0; lineStart <= len(source); {
		lineEnd := len(source)
		if offset := bytes.IndexByte(source[lineStart:], '\n'); offset >= 0 {
			lineEnd = lineStart + offset
		}
		line := source[lineStart:lineEnd]
		if declarationLinePrefix.Match(line) {
			for _, match := range declarationKeywordAfterComma.FindAllSubmatchIndex(line, -1) {
				if len(match) >= 4 && match[2] >= 0 && match[3] > match[2] {
					spans = append(spans, byteSpan{start: lineStart + match[2], end: lineStart + match[3]})
				}
			}
		}
		if lineEnd == len(source) {
			break
		}
		lineStart = lineEnd + 1
	}
	return spans
}

func recoveryNodesOverlapSpans(root *tree_sitter.Node, spans []byteSpan) bool {
	recoveryNodes := 0
	valid := true
	Walk(root, func(node *tree_sitter.Node) bool {
		if node.IsError() || node.IsMissing() {
			recoveryNodes++
			if !nodeOverlapsAnySpan(node, spans) {
				valid = false
			}
		}
		return true
	})
	return valid && recoveryNodes > 0
}

func nodeOverlapsAnySpan(node *tree_sitter.Node, spans []byteSpan) bool {
	start, end := int(node.StartByte()), int(node.EndByte())
	for _, span := range spans {
		if start == end {
			if start >= span.start && start <= span.end {
				return true
			}
			continue
		}
		if start < span.end && end > span.start {
			return true
		}
	}
	return false
}
