package intel

import (
	"context"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type setFixCandidate struct {
	line, start, leftStart int
	lhs, rhs, kind         string
}

// setAssignmentQuickFixesContext borrows syntax first and resolves types only
// after releasing the tree lock. It never invokes whole-document diagnostics.
func (a Analyzer) setAssignmentQuickFixesContext(ctx context.Context, doc Document, parsed *ast.ParsedDocument, selection Range) ([]CodeAction, error) {
	if a.DB == nil || a.TypeDBResolutionIncomplete {
		return nil, nil
	}
	var candidates []setFixCandidate
	lines := documentLines(doc)
	err := parsed.ReadContext(ctx, func(view ast.ParsedView) error {
		if view.Root == nil || view.Root.HasError() || ast.HasMissing(view.Root) {
			return nil
		}
		var walk func(*tree_sitter.Node) error
		walk = func(node *tree_sitter.Node) error {
			if err := ctx.Err(); err != nil {
				return err
			}
			if int(node.EndPosition().Row) < selection.Start.Line || int(node.StartPosition().Row) > selection.End.Line {
				return nil
			}
			if strings.Contains(node.Kind(), "conditional") || strings.Contains(node.Kind(), "preprocessor") {
				return nil
			}
			switch node.Kind() {
			case "assignment_statement", "let_statement", "set_statement":
				left, right := node.ChildByFieldName("left"), node.ChildByFieldName("right")
				if left == nil || right == nil || left.Kind() != "identifier" || node.StartPosition().Row != node.EndPosition().Row {
					return nil
				}
				start, leftStart := int(node.StartByte()), int(left.StartByte())
				line := int(node.StartPosition().Row)
				// The diagnostic predicate currently handles one physical assignment
				// per line. Reject labels, inline If/colon forms and continuations.
				lineStart := start - int(node.StartPosition().Column)
				if strings.Trim(string(view.Source[lineStart:start]), " \t") != "" {
					return nil
				}
				if line > 0 && strings.HasSuffix(strings.TrimSpace(lines[line-1]), "_") {
					return nil
				}
				if suffix := stripLineComment(lines[line][int(node.EndPosition().Column):]); strings.TrimSpace(suffix) != "" {
					return nil
				}
				lhs, rhs := left.Utf8Text(view.Source), right.Utf8Text(view.Source)
				switch right.Kind() {
				case "identifier", "string_literal", "number_literal", "boolean_literal":
				default:
					return nil
				}
				if !isIdentifier(lhs) {
					return nil
				}
				prefix := strings.Trim(string(view.Source[start:leftStart]), " \t")
				if node.Kind() == "assignment_statement" && prefix != "" || node.Kind() == "let_statement" && !strings.EqualFold(prefix, "Let") || node.Kind() == "set_statement" && !strings.EqualFold(prefix, "Set") {
					return nil
				}
				candidates = append(candidates, setFixCandidate{line: line, start: start, leftStart: leftStart, lhs: lhs, rhs: rhs, kind: node.Kind()})
				return nil
			}
			for i := range node.NamedChildCount() {
				if err := walk(node.NamedChild(i)); err != nil {
					return err
				}
			}
			return nil
		}
		return walk(view.Root)
	})
	if err != nil || len(candidates) == 0 {
		return nil, err
	}
	index, ok := a.documentIndexFor(doc)
	if !ok {
		return nil, ctx.Err()
	}
	var out []CodeAction
	var associated []Diagnostic
	for _, candidate := range candidates {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		lhsType, ok := a.setFixDeclaredType(doc, index, candidate.lhs, candidate.leftStart)
		if !ok {
			continue
		}
		rhsType, literal := byRefLiteralType(candidate.rhs)
		if !literal {
			if !isIdentifier(candidate.rhs) {
				continue
			}
			rhsType, ok = a.setFixDeclaredType(doc, index, candidate.rhs, candidate.leftStart)
			if !ok {
				continue
			}
		}
		code, title, text := "VB036", "Insert Set for object assignment", "Set "
		end := candidate.start
		if candidate.kind == "set_statement" {
			if !valueDiagnosticType(lhsType) || !valueDiagnosticType(rhsType) {
				continue
			}
			code, title, text, end = "VB037", "Remove Set from value assignment", "", candidate.leftStart
		} else {
			if !concreteObjectDiagnosticType(lhsType) || !concreteObjectDiagnosticType(rhsType) {
				continue
			}
			if assignable, known := a.DB.IsAssignable(lhsType, rhsType); !known || !assignable {
				continue
			}
			if candidate.kind == "let_statement" {
				end = candidate.leftStart
			}
		}
		diagnostics := a.assignmentDiagnosticsInLinesContext(ctx, doc, candidate.line, candidate.line+1)
		for _, diagnostic := range diagnostics {
			if diagnostic.Code != code {
				continue
			}
			edit := CodeActionEdit{Range: Range{Start: positionForDocumentByteOffset(doc, candidate.start), End: positionForDocumentByteOffset(doc, end)}, NewText: text}
			if rangeIntersects(selection, diagnostic.Range) || rangeIntersects(selection, edit.Range) {
				out = append(out, CodeAction{Title: title, Kind: "quickfix", Edits: []CodeActionEdit{edit}, Diagnostics: []Diagnostic{diagnostic}})
				associated = append(associated, diagnostic)
			}
		}
	}
	// Parse source suppression directives once, including for a multiline
	// selection. Filtering preserves the original diagnostic order.
	visible := applyDocumentInlineSuppressions(a.RootDir, doc, associated)
	kept := 0
	for _, action := range out {
		if kept < len(visible) && action.Diagnostics[0] == visible[kept] {
			out[kept] = action
			kept++
		}
	}
	return out[:kept], ctx.Err()
}

// Only explicit, uniquely bound variables from this source revision qualify.
// Do not turn assignment inference, functions, properties or conditional symbols
// into authority to change object/value assignment semantics.
func (a Analyzer) setFixDeclaredType(doc Document, index *documentIndex, name string, offset int) (string, bool) {
	procedure := currentProcedureNameAt(doc, positionForDocumentByteOffset(doc, offset), nil)
	var locals, modules []Symbol
	for _, sym := range index.symbolsByName[indexName(name)] {
		if !a.sameDocumentSymbol(doc, sym) || !a.visibleDefinitionSymbol(doc, procedure, sym) {
			continue
		}
		if sym.Parent != "" {
			locals = append(locals, sym)
		} else {
			modules = append(modules, sym)
		}
	}
	bound := modules
	if len(locals) > 0 {
		bound = locals
	}
	if len(bound) != 1 {
		return "", false
	}
	sym := bound[0]
	if sym.Kind != "local_variable" && sym.Kind != "parameter" && sym.Kind != "module_variable" || sym.IsArray || len(sym.ConditionalBranches) > 0 || lowConfidenceDiagnosticType(sym.ReturnType) {
		return "", false
	}
	return canonicalDiagnosticType(a.DB, sym.ReturnType), true
}
