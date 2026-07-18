package lint

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

const procedureNameConstantRuleID = "VB044"

// ProcedureNameConstantFix identifies the string literal that should be
// replaced when a configured procedure-name constant has drifted.
// Byte offsets are relative to the UTF-8 source buffer supplied to the linter.
type ProcedureNameConstantFix struct {
	StartByte    int
	EndByte      int
	Line         int
	Column       int
	ConstantName string
	CurrentValue string
	ExpectedName string
}

// ProcedureNameConstantFixesParsed reports source edits for configured local
// procedure-name constants. It never modifies the source document.
func (l Linter) ProcedureNameConstantFixesParsed(doc *vbaast.ParsedDocument) ([]ProcedureNameConstantFix, error) {
	if !l.Config.Lint.ProcedureNameConstant.Enabled {
		return nil, nil
	}
	var fixes []ProcedureNameConstantFix
	err := doc.Read(func(view vbaast.ParsedView) error {
		fixes = procedureNameConstantFixes(view.Root, view.Source, l.Config.Lint.ProcedureNameConstant)
		return nil
	})
	return fixes, err
}

func (l Linter) procedureNameConstantIssues(path string, root *tree_sitter.Node, source []byte) []Issue {
	fixes := procedureNameConstantFixes(root, source, l.Config.Lint.ProcedureNameConstant)
	issues := make([]Issue, 0, len(fixes))
	for _, fix := range fixes {
		issue := l.issueAt(path, sourceByteRange(source, fix.StartByte, fix.EndByte), procedureNameConstantRuleID, "warning", fmt.Sprintf("Local constant %q is %q but its enclosing procedure is %q.", fix.ConstantName, fix.CurrentValue, fix.ExpectedName))
		issue.Kind = "procedure_name_constant"
		issue.Symbol = fix.ConstantName
		issue.Suggestion = fmt.Sprintf("Update the string literal to %q.", fix.ExpectedName)
		issues = append(issues, issue)
	}
	return issues
}

func procedureNameConstantFixes(root *tree_sitter.Node, source []byte, cfg config.ProcedureNameConstantConfig) []ProcedureNameConstantFix {
	if root == nil || !cfg.Enabled || strings.TrimSpace(cfg.ConstantName) == "" {
		return nil
	}
	var fixes []ProcedureNameConstantFix
	collectProcedureNameConstantFixes(root, source, cfg.ConstantName, &fixes)
	return fixes
}

func collectProcedureNameConstantFixes(node *tree_sitter.Node, source []byte, constantName string, fixes *[]ProcedureNameConstantFix) {
	if node == nil {
		return
	}
	if procedureNameConstantProcedure(node.Kind()) {
		procedureName := procedureNameConstantProcedureName(node, source)
		if procedureName != "" {
			for i := uint(0); i < node.NamedChildCount(); i++ {
				collectProcedureLocalConstantFixes(node.NamedChild(i), source, constantName, procedureName, fixes)
			}
		}
		return
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		collectProcedureNameConstantFixes(node.NamedChild(i), source, constantName, fixes)
	}
}

func procedureNameConstantProcedure(kind string) bool {
	switch kind {
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration":
		return true
	default:
		return false
	}
}

func procedureNameConstantProcedureName(node *tree_sitter.Node, source []byte) string {
	if name := childByFieldNameAny(node, "name"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	if name := firstNamedChildKind(node, "identifier"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	return ""
}

func collectProcedureLocalConstantFixes(node *tree_sitter.Node, source []byte, constantName, procedureName string, fixes *[]ProcedureNameConstantFix) {
	if node == nil {
		return
	}
	if procedureNameConstantProcedure(node.Kind()) {
		return
	}
	if node.Kind() == "const_declaration" {
		for i := uint(0); i < node.NamedChildCount(); i++ {
			declarator := node.NamedChild(i)
			if declarator == nil || declarator.Kind() != "const_declarator" {
				continue
			}
			name := procedureNameConstantDeclaratorName(declarator, source)
			if !strings.EqualFold(name, constantName) {
				continue
			}
			start, end, current, ok := procedureNameConstantLiteral(declarator, source)
			if !ok || current == procedureName {
				continue
			}
			r := sourceByteRange(source, start, end)
			*fixes = append(*fixes, ProcedureNameConstantFix{
				StartByte:    start,
				EndByte:      end,
				Line:         r.StartLine,
				Column:       r.StartColumn,
				ConstantName: name,
				CurrentValue: current,
				ExpectedName: procedureName,
			})
		}
		return
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child != nil && !procedureNameConstantProcedure(child.Kind()) {
			collectProcedureLocalConstantFixes(child, source, constantName, procedureName, fixes)
		}
	}
}

func procedureNameConstantDeclaratorName(node *tree_sitter.Node, source []byte) string {
	if name := childByFieldNameAny(node, "name"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	if name := firstNamedChildKind(node, "identifier"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	return ""
}

func procedureNameConstantLiteral(node *tree_sitter.Node, source []byte) (int, int, string, bool) {
	if value := childByFieldNameAny(node, "value", "initializer", "default_value"); value != nil {
		if start, end, text, ok := directVBAStringLiteral(value, source); ok {
			return start, end, text, true
		}
	}
	return directVBAStringLiteral(node, source)
}

func directVBAStringLiteral(node *tree_sitter.Node, source []byte) (int, int, string, bool) {
	if node == nil {
		return 0, 0, "", false
	}
	start, end := int(node.StartByte()), int(node.EndByte())
	if start < 0 || end < start || end > len(source) {
		return 0, 0, "", false
	}
	raw := string(source[start:end])
	leading := len(raw) - len(strings.TrimLeft(raw, " \t\r\n"))
	text := strings.TrimSpace(raw)
	if value, ok := decodeVBAStringLiteral(text); ok {
		return start + leading, start + leading + len(text), value, true
	}
	equals := strings.Index(raw, "=")
	if equals < 0 {
		return 0, 0, "", false
	}
	rawValue := raw[equals+1:]
	leading = equals + 1 + len(rawValue) - len(strings.TrimLeft(rawValue, " \t\r\n"))
	text = strings.TrimSpace(rawValue)
	value, ok := decodeVBAStringLiteral(text)
	if !ok {
		return 0, 0, "", false
	}
	return start + leading, start + leading + len(text), value, true
}

func decodeVBAStringLiteral(text string) (string, bool) {
	if len(text) < 2 || text[0] != '"' || text[len(text)-1] != '"' {
		return "", false
	}
	for i := 1; i < len(text)-1; i++ {
		if text[i] != '"' {
			continue
		}
		if i+1 >= len(text)-1 || text[i+1] != '"' {
			return "", false
		}
		i++
	}
	return strings.ReplaceAll(text[1:len(text)-1], `""`, `"`), true
}

func sourceByteRange(source []byte, start, end int) vbaast.Range {
	start = max(0, min(start, len(source)))
	end = max(start, min(end, len(source)))
	line, column := 1, 1
	startLine, startColumn := line, column
	for i := 0; i < end; i++ {
		if i == start {
			startLine, startColumn = line, column
		}
		if source[i] == '\n' {
			line++
			column = 1
		} else {
			column++
		}
	}
	if start == end {
		startLine, startColumn = line, column
	}
	return vbaast.Range{StartLine: startLine, StartColumn: startColumn, EndLine: line, EndColumn: column, StartByte: start, EndByte: end}
}
