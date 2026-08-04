package lint

import (
	"context"
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
	return l.ProcedureNameConstantFixesParsedContext(context.Background(), doc)
}

// ProcedureNameConstantFixesParsedContext is the cancellable variant of
// ProcedureNameConstantFixesParsed.
func (l Linter) ProcedureNameConstantFixesParsedContext(ctx context.Context, doc *vbaast.ParsedDocument) ([]ProcedureNameConstantFix, error) {
	if !l.Config.Lint.ProcedureNameConstant.Enabled {
		return nil, nil
	}
	var fixes []ProcedureNameConstantFix
	err := doc.Read(func(view vbaast.ParsedView) error {
		var scanErr error
		fixes, scanErr = procedureNameConstantFixesContext(ctx, view.Root, view.Source, l.Config.Lint.ProcedureNameConstant)
		return scanErr
	})
	return fixes, err
}

func (l Linter) procedureNameConstantIssues(path string, root *tree_sitter.Node, source []byte) []Issue {
	issues, _ := l.procedureNameConstantIssuesContext(context.Background(), path, root, source)
	return issues
}

func (l Linter) procedureNameConstantIssuesContext(ctx context.Context, path string, root *tree_sitter.Node, source []byte) ([]Issue, error) {
	fixes, err := procedureNameConstantFixesContext(ctx, root, source, l.Config.Lint.ProcedureNameConstant)
	if err != nil {
		return nil, err
	}
	issues := make([]Issue, 0, len(fixes))
	for i, fix := range fixes {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		issue := l.issueAt(path, sourceByteRange(source, fix.StartByte, fix.EndByte), procedureNameConstantRuleID, "warning", fmt.Sprintf("Local constant %q is %q but its enclosing procedure is %q.", fix.ConstantName, fix.CurrentValue, fix.ExpectedName))
		issue.Kind = "procedure_name_constant"
		issue.Symbol = fix.ConstantName
		issue.Suggestion = fmt.Sprintf("Update the string literal to %q.", fix.ExpectedName)
		issues = append(issues, issue)
	}
	return issues, ctx.Err()
}

func procedureNameConstantFixes(root *tree_sitter.Node, source []byte, cfg config.ProcedureNameConstantConfig) []ProcedureNameConstantFix {
	fixes, _ := procedureNameConstantFixesContext(context.Background(), root, source, cfg)
	return fixes
}

func procedureNameConstantFixesContext(ctx context.Context, root *tree_sitter.Node, source []byte, cfg config.ProcedureNameConstantConfig) ([]ProcedureNameConstantFix, error) {
	if root == nil || !cfg.Enabled || strings.TrimSpace(cfg.ConstantName) == "" {
		return nil, nil
	}
	var fixes []ProcedureNameConstantFix
	visited := uint64(0)
	if err := collectProcedureNameConstantFixesContext(ctx, root, source, cfg.ConstantName, &fixes, &visited); err != nil {
		return nil, err
	}
	return fixes, ctx.Err()
}

func collectProcedureNameConstantFixes(node *tree_sitter.Node, source []byte, constantName string, fixes *[]ProcedureNameConstantFix) {
	visited := uint64(0)
	_ = collectProcedureNameConstantFixesContext(context.Background(), node, source, constantName, fixes, &visited)
}

func collectProcedureNameConstantFixesContext(ctx context.Context, node *tree_sitter.Node, source []byte, constantName string, fixes *[]ProcedureNameConstantFix, visited *uint64) error {
	if node == nil {
		return nil
	}
	*visited++
	if *visited&0xff == 0 {
		if err := ctx.Err(); err != nil {
			return err
		}
	}
	if procedureNameConstantProcedure(node.Kind()) {
		procedureName := procedureNameConstantProcedureName(node, source)
		if procedureName != "" {
			for i := uint(0); i < node.NamedChildCount(); i++ {
				if err := collectProcedureLocalConstantFixesContext(ctx, node.NamedChild(i), source, constantName, procedureName, fixes, visited); err != nil {
					return err
				}
			}
		}
		return nil
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		if err := collectProcedureNameConstantFixesContext(ctx, node.NamedChild(i), source, constantName, fixes, visited); err != nil {
			return err
		}
	}
	return nil
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
	visited := uint64(0)
	_ = collectProcedureLocalConstantFixesContext(context.Background(), node, source, constantName, procedureName, fixes, &visited)
}

func collectProcedureLocalConstantFixesContext(ctx context.Context, node *tree_sitter.Node, source []byte, constantName, procedureName string, fixes *[]ProcedureNameConstantFix, visited *uint64) error {
	if node == nil {
		return nil
	}
	*visited++
	if *visited&0xff == 0 {
		if err := ctx.Err(); err != nil {
			return err
		}
	}
	if procedureNameConstantProcedure(node.Kind()) {
		return nil
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
		return nil
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child != nil && !procedureNameConstantProcedure(child.Kind()) {
			if err := collectProcedureLocalConstantFixesContext(ctx, child, source, constantName, procedureName, fixes, visited); err != nil {
				return err
			}
		}
	}
	return nil
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
