package intel

import (
	"context"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"

	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// A mutation finding is necessary but not sufficient: a read-only ByRef can
// observe writes through another alias. The first provider accepts only one
// scalar parameter and a straight-line body with no other bindings or calls.
func byValParameterQuickFixesContext(ctx context.Context, doc Document, parsed *ast.ParsedDocument, selection Range, diagnostics []Diagnostic) ([]CodeAction, error) {
	if len(diagnostics) == 0 {
		return nil, ctx.Err()
	}
	byRange := make(map[Range]Diagnostic)
	for _, diagnostic := range diagnostics {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		if diagnostic.Code == "VBA275" {
			byRange[diagnostic.Range] = diagnostic
		}
	}
	var out []CodeAction
	err := parsed.ReadContext(ctx, func(view ast.ParsedView) error {
		if view.Root == nil || view.HasError || view.HasMissing {
			return nil
		}
		// Only direct module children qualify, excluding conditional branches.
		// Count names before selecting so duplicate declarations fail closed.
		names := make(map[string]int)
		for i := range view.Root.NamedChildCount() {
			if err := ctx.Err(); err != nil {
				return err
			}
			node := view.Root.NamedChild(i)
			if name := node.ChildByFieldName("name"); name != nil {
				names[strings.ToLower(name.Utf8Text(view.Source))]++
			}
		}
		for i := range view.Root.NamedChildCount() {
			if err := ctx.Err(); err != nil {
				return err
			}
			node := view.Root.NamedChild(i)
			if int(node.StartPosition().Row) > selection.End.Line || int(node.EndPosition().Row) < selection.Start.Line {
				continue
			}
			start, end, ok := byValModifierEdit(ctx, node, view.Source, names)
			if !ok {
				continue
			}
			rng := Range{Start: positionForDocumentByteOffset(doc, start), End: positionForDocumentByteOffset(doc, end)}
			if !rangeIntersects(selection, rng) {
				continue
			}
			if diagnostic, ok := byRange[rng]; ok {
				out = append(out, CodeAction{Title: "Pass parameter ByVal", Kind: "quickfix", Edits: []CodeActionEdit{{Range: rng, NewText: "ByVal"}}, Diagnostics: []Diagnostic{diagnostic}})
			}
		}
		return ctx.Err()
	})
	return out, err
}

func byValModifierEdit(ctx context.Context, node *tree_sitter.Node, source []byte, names map[string]int) (int, int, bool) {
	if node.Kind() != "function_declaration" {
		return 0, 0, false
	}
	visibility, name := node.ChildByFieldName("visibility"), node.ChildByFieldName("name")
	if visibility == nil || !strings.EqualFold(visibility.Utf8Text(source), "Private") || name == nil || name.Kind() != "identifier" || !isIdentifier(name.Utf8Text(source)) || names[strings.ToLower(name.Utf8Text(source))] != 1 || node.ChildByFieldName("modifiers") != nil {
		return 0, 0, false
	}
	parameters := procedureir.ParametersFromNode(node, source)
	if len(parameters) != 1 {
		return 0, 0, false
	}
	p := parameters[0]
	if p.Recovered || p.Optional || p.ParamArray || p.IsArray || p.HasDefault || !p.PassingExplicit || p.PassingRange == nil || !strings.EqualFold(p.Passing, "ByRef") || !byValPrimitiveType(p.Type) || !isIdentifier(p.Name) || strings.EqualFold(p.Name, name.Utf8Text(source)) {
		return 0, 0, false
	}
	// Positively recognize the complete parameter spelling. This excludes type
	// characters, brackets, fixed strings and parser-normalized extra syntax.
	words := strings.Fields(string(source[p.Range.StartByte:p.Range.EndByte]))
	if len(words) != 4 || !strings.EqualFold(words[0], "ByRef") || words[1] != p.Name || !strings.EqualFold(words[2], "As") || !strings.EqualFold(words[3], p.Type) {
		return 0, 0, false
	}
	typ := node.ChildByFieldName("type")
	if typ == nil || !byValPrimitiveType(strings.TrimSpace(strings.TrimPrefix(strings.ToLower(typ.Utf8Text(source)), "as"))) {
		return 0, 0, false
	}
	// The header, body and terminator may contain comments/newlines only in
	// addition to these known fields; procedure attributes are not harmless.
	body := node.ChildByFieldName("body")
	if body == nil {
		return 0, 0, false
	}
	for i := range node.NamedChildCount() {
		child := node.NamedChild(i)
		switch child.Kind() {
		case "visibility", "identifier", "parameter_list", "as_type_clause", "end_function_statement", "newline", "comment":
		case "block":
			if !byValScalarBody(ctx, child, source, p.Name, name.Utf8Text(source)) {
				return 0, 0, false
			}
		default:
			return 0, 0, false
		}
	}
	headerEnd := int(body.StartByte())
	header := string(source[node.StartByte():headerEnd])
	first, rest, _ := strings.Cut(header, "\n")
	// Other header-adjacent comments may occur before the body, but a physical
	// continuation or colon in the declaration is outside this increment.
	if strings.HasSuffix(strings.TrimSpace(stripLineComment(first)), "_") || strings.Contains(stripLineComment(first), ":") || strings.TrimSpace(stripLineComment(first)) == "" {
		return 0, 0, false
	}
	for line := range strings.SplitSeq(rest, "\n") {
		if strings.TrimSpace(stripLineComment(line)) != "" {
			return 0, 0, false
		}
	}
	return p.PassingRange.StartByte, p.PassingRange.EndByte, true
}

func byValPrimitiveType(name string) bool {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "boolean", "byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "date", "string":
		return true
	}
	return false
}

func byValScalarBody(ctx context.Context, node *tree_sitter.Node, source []byte, parameter, procedure string) bool {
	assignments := 0
	for i := range node.NamedChildCount() {
		if ctx.Err() != nil {
			return false
		}
		child := node.NamedChild(i)
		switch child.Kind() {
		case "newline", "comment":
		case "assignment_statement", "let_statement":
			left, right := child.ChildByFieldName("left"), child.ChildByFieldName("right")
			if left == nil || left.Kind() != "identifier" || !strings.EqualFold(left.Utf8Text(source), procedure) || !byValScalarExpression(ctx, right, source, parameter) {
				return false
			}
			assignments++
		default:
			return false
		}
		if child.Kind() != "newline" && child.Kind() != "comment" && child.StartPosition().Row != child.EndPosition().Row {
			return false
		}
		if child.Kind() != "newline" && child.Kind() != "comment" {
			start, end := int(child.StartByte()), int(child.EndByte())
			lineStart := start - int(child.StartPosition().Column)
			lineEnd, _ := rawLineBounds(string(source), end)
			if strings.Trim(string(source[lineStart:start]), " \t") != "" || strings.TrimSpace(stripLineComment(string(source[end:lineEnd]))) != "" {
				return false
			}
		}
	}
	return assignments > 0
}

func byValScalarExpression(ctx context.Context, node *tree_sitter.Node, source []byte, parameter string) bool {
	if node == nil || ctx.Err() != nil {
		return false
	}
	switch node.Kind() {
	case "identifier":
		return strings.EqualFold(node.Utf8Text(source), parameter)
	case "number_literal", "string_literal", "boolean_literal", "date_literal":
		return true
	case "binary_expression", "unary_expression", "parenthesized_expression":
		for i := range node.NamedChildCount() {
			if !byValScalarExpression(ctx, node.NamedChild(i), source, parameter) {
				return false
			}
		}
		return true
	}
	return false
}
