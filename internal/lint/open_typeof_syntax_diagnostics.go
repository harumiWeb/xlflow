package lint

import (
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

const (
	openSyntaxDiagnosticCode   = "VB064"
	typeOfSyntaxDiagnosticCode = "VB065"
)

// openTypeOfSyntaxIssues recognizes only recovery shapes that tree-sitter
// identifies unambiguously.  Generic ERROR-only forms remain VB014; this
// helper must not turn broad source-text guesses into compile-equivalent
// diagnostics.
func (l Linter) openTypeOfSyntaxIssues(path, source string, root *tree_sitter.Node) []Issue {
	if root == nil {
		return nil
	}
	issues := make([]Issue, 0, 2)
	vbaast.Walk(root, func(node *tree_sitter.Node) bool {
		switch node.Kind() {
		case "open_statement":
			mode := node.ChildByFieldName("mode")
			if mode != nil && syntaxNodeHasMissing(mode) {
				issues = append(issues, l.openTypeOfIssue(path, vbaast.NodeRange(mode), openSyntaxDiagnosticCode, "missing_mode", "Open must specify a file mode after For.", "Add Input, Output, Append, Random, or Binary after For."))
			}
		case "block":
			issues = append(issues, l.typeOfTrailingTokenIssues(path, source, node)...)
		}
		return true
	})
	return issues
}

// tree-sitter wraps an omitted required token in a named grammar node (for
// example file_mode) whose missing token is not exposed as a child. Its
// S-expression still carries the explicit MISSING marker, which is safer than
// guessing from source text.
func syntaxNodeHasMissing(node *tree_sitter.Node) bool {
	return node != nil && (node.IsMissing() || strings.Contains(node.ToSexp(), "MISSING"))
}

func (l Linter) typeOfTrailingTokenIssues(path, source string, block *tree_sitter.Node) []Issue {
	var issues []Issue
	for i := uint(0); i+1 < block.NamedChildCount(); i++ {
		statement := block.NamedChild(i)
		if statement == nil || statement.Kind() != "expression_statement" || statement.NamedChildCount() == 0 {
			continue
		}
		expression := statement.NamedChild(0)
		if expression == nil || expression.Kind() != "type_of_expression" {
			continue
		}
		next := block.NamedChild(i + 1)
		if next == nil || next.Kind() != "call_statement" || next.StartPosition().Row != statement.StartPosition().Row || next.StartByte() <= statement.EndByte() {
			continue
		}
		between := source[statement.EndByte():next.StartByte()]
		trailing := strings.TrimSpace(source[expression.EndByte():next.EndByte()])
		if strings.ContainsAny(between, ":\r\n") || trailing == "" {
			continue
		}
		range_ := vbaast.NodeRange(next)
		issues = append(issues, l.openTypeOfIssue(path, range_, typeOfSyntaxDiagnosticCode, "trailing_token", "A TypeOf expression cannot be followed by another token on the same statement.", "Remove the trailing token or separate it with a valid statement boundary."))
	}
	return issues
}

func (l Linter) openTypeOfIssue(path string, r vbaast.Range, code, kind, message, suggestion string) Issue {
	issue := l.issueAt(path, r, code, "error", message)
	issue.EndLine = r.EndLine
	issue.EndColumn = r.EndColumn
	issue.Kind = kind
	issue.Suggestion = suggestion
	return issue
}
