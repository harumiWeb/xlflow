package ast

import (
	"regexp"
	"strings"

	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

var declarationKeywordAfterComma = regexp.MustCompile(`(?i)^\s*(?:Dim|ReDim)\b[^\r\n]*,\s*(?:Dim|ReDim)\b`)

var declarationWithTypeCharacter = regexp.MustCompile(`(?i)^\s*(?:Dim|ReDim|Static|Public|Private|Friend|Const)\b`)

// IsDeclarationKeywordRecovery reports the narrow recovery shape introduced
// when tree-sitter-vba rejects a second Dim/ReDim keyword after a declaration
// comma. It deliberately requires every recovery node to be on one of those
// source lines, so ordinary parser failures keep their existing behavior.
func IsDeclarationKeywordRecovery(root *tree_sitter.Node, source []byte) bool {
	if root == nil || (!root.HasError() && !HasMissing(root)) {
		return false
	}
	lines := make(map[int]struct{})
	for lineNo, line := range strings.Split(string(source), "\n") {
		if declarationKeywordAfterComma.MatchString(strings.TrimSuffix(line, "\r")) {
			lines[lineNo+1] = struct{}{}
		}
	}
	if len(lines) == 0 {
		return false
	}
	recoveryNodes := 0
	Walk(root, func(node *tree_sitter.Node) bool {
		if node.IsError() || node.IsMissing() {
			recoveryNodes++
			if _, ok := lines[NodeRange(node).StartLine]; !ok {
				lines = nil
			}
		}
		return true
	})
	return lines != nil && recoveryNodes > 0
}

// IsIdentifierTypeCharacterRecovery recognizes the v0.12.0 CST shape for a
// legal declaration whose Single (!) suffix is recovered as an error node.
// The source-level suffix check is deliberately tied to a variable declarator
// so an unrelated bang expression cannot suppress parser diagnostics.
func IsIdentifierTypeCharacterRecovery(root *tree_sitter.Node, source []byte) bool {
	if root == nil || (!root.HasError() && !HasMissing(root)) {
		return false
	}
	declarationLines := make(map[int]struct{})
	candidates := 0
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
			declarationLines[lineNo] = struct{}{}
			candidates++
		}
		return true
	})
	if candidates == 0 {
		return false
	}
	recoveryNodes := 0
	valid := true
	Walk(root, func(node *tree_sitter.Node) bool {
		if node.IsError() || node.IsMissing() {
			recoveryNodes++
			if _, ok := declarationLines[NodeRange(node).StartLine]; !ok {
				valid = false
			}
		}
		return true
	})
	return valid && recoveryNodes > 0
}
