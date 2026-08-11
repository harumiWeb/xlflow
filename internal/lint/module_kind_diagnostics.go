package lint

import (
	"context"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// moduleKindIssuesContext contains compile-equivalent checks whose meaning is
// determined by the host module rather than by a declaration's spelling. The
// walker deliberately follows the same conservative conditional-compilation
// policy as declaration diagnostics: a statically discarded branch is not
// inspected, while an unknown branch is inspected independently.
func (l Linter) moduleKindIssuesContext(ctx context.Context, path string, source []byte, root *tree_sitter.Node) ([]Issue, error) {
	if root == nil {
		return nil, nil
	}
	if ctx == nil {
		ctx = context.Background()
	}
	walker := moduleKindDiagnosticWalker{ctx: ctx, linter: l, path: path, source: source, moduleKind: strings.ToLower(strings.TrimSpace(l.moduleKindForPath(path)))}
	walker.walk(root, moduleWalkState{module: true})
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	sort.SliceStable(walker.issues, func(i, j int) bool {
		if walker.issues[i].Line != walker.issues[j].Line {
			return walker.issues[i].Line < walker.issues[j].Line
		}
		if walker.issues[i].Column != walker.issues[j].Column {
			return walker.issues[i].Column < walker.issues[j].Column
		}
		return walker.issues[i].Kind < walker.issues[j].Kind
	})
	return walker.issues, nil
}

type moduleWalkState struct {
	module    bool
	procedure bool
}

type moduleKindDiagnosticWalker struct {
	ctx        context.Context
	linter     Linter
	path       string
	source     []byte
	moduleKind string
	issues     []Issue
}

func (w *moduleKindDiagnosticWalker) walk(node *tree_sitter.Node, state moduleWalkState) {
	if node == nil || w.ctx.Err() != nil {
		return
	}
	kind := node.Kind()
	if kind == "preprocessor_if" || kind == "type_preprocessor_if" {
		w.walkConditional(node, state)
		return
	}
	if kind == "preprocessor_else" || kind == "preprocessor_elseif" || kind == "type_preprocessor_else" || kind == "type_preprocessor_elseif" {
		if body := node.ChildByFieldName("body"); body != nil {
			w.walk(body, state)
		}
		return
	}

	if isModuleProcedureDeclarationKind(kind) {
		header := node
		if strings.HasPrefix(kind, "conditional_") {
			header = procedureHeaderNode(node)
		}
		if state.module && header != nil && strings.EqualFold(visibilityText(header, w.source), "Friend") &&
			(w.moduleKind == "standard" || w.moduleKind == "document") {
			w.add(header, declarationNameNode(header), "invalid_friend_module", declarationNodeName(header, w.source), "Friend procedures are not valid in standard or document modules.")
		}
		if strings.HasPrefix(kind, "conditional_") {
			// The conditional procedure helper owns branch traversal. Its body
			// is still a procedure body for Me and nested declaration checks.
			w.walkConditionalProcedure(node)
			return
		}
		state = moduleWalkState{procedure: true}
	}

	switch kind {
	case "event_declaration", "event_statement":
		if state.module && w.moduleKind == "standard" {
			w.add(node, declarationNameNode(node), "invalid_event_module", declarationNodeName(node, w.source), "Event declarations are not valid in standard modules.")
		}
	case "implements_statement":
		if state.module && w.moduleKind == "standard" {
			w.add(node, declarationNameNode(node), "invalid_implements_module", declarationNodeName(node, w.source), "Implements statements are not valid in standard modules.")
		}
	case "variable_declaration":
		w.variableIssues(node, state)
	case "const_declaration":
		if state.module && isObjectModuleKind(w.moduleKind) && strings.EqualFold(visibilityText(node, w.source), "Public") {
			for i := uint(0); i < node.NamedChildCount(); i++ {
				child := node.NamedChild(i)
				if child != nil && child.Kind() == "const_declarator" {
					w.add(node, declarationNameNode(child), "invalid_public_object_member", declarationNodeName(child, w.source), "Public Const declarations are not valid in object modules.")
				}
			}
		}
	case "type_declaration":
		if state.module && isObjectModuleKind(w.moduleKind) && strings.EqualFold(visibilityText(node, w.source), "Public") {
			w.add(node, declarationNameNode(node), "invalid_public_object_member", declarationNodeName(node, w.source), "Public Type declarations are not valid in object modules.")
		}
	case "declare_statement", "declare_sub_statement", "declare_function_statement":
		if state.module && isObjectModuleKind(w.moduleKind) && !strings.EqualFold(visibilityText(node, w.source), "Private") {
			w.add(node, declarationNameNode(node), "invalid_public_object_member", declarationNodeName(node, w.source), "Declare members in object modules must be Private.")
		}
	case "attribute_statement", "attribute_declaration":
		// Attributes are metadata, not VBA declarations. In particular, an
		// attribute value containing the text Me must not become VB051.
		return
	}

	if w.moduleKind == "standard" && kind == "identifier" && isMeKeywordNode(node, w.source) {
		w.add(node, node, "invalid_me_context", "Me", "The Me keyword is only valid in object modules.")
	}
	if kind == "type_declaration" || kind == "enum_declaration" {
		state = moduleWalkState{}
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		w.walk(node.NamedChild(i), state)
	}
}

func (w *moduleKindDiagnosticWalker) walkConditional(node *tree_sitter.Node, state moduleWalkState) {
	if known, value := conditionalConstant(node, w.source); known {
		if value {
			if body := node.ChildByFieldName("body"); body != nil {
				w.walk(body, state)
			}
			return
		}
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child == nil || (!strings.Contains(child.Kind(), "else") && !strings.Contains(child.Kind(), "elseif")) {
				continue
			}
			if body := child.ChildByFieldName("body"); body != nil {
				w.walk(body, state)
			}
		}
		return
	}
	if body := node.ChildByFieldName("body"); body != nil {
		w.walk(body, state)
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || (!strings.Contains(child.Kind(), "else") && !strings.Contains(child.Kind(), "elseif")) {
			continue
		}
		if body := child.ChildByFieldName("body"); body != nil {
			w.walk(body, state)
		}
	}
}

func (w *moduleKindDiagnosticWalker) walkConditionalProcedure(node *tree_sitter.Node) {
	for _, field := range []string{"consequence_body", "alternative_body", "body"} {
		if body := node.ChildByFieldName(field); body != nil {
			w.walk(body, moduleWalkState{procedure: true})
		}
	}
}

func (w *moduleKindDiagnosticWalker) variableIssues(node *tree_sitter.Node, state moduleWalkState) {
	withEvents := node.ChildByFieldName("with_events_modifier") != nil || hasWord(node.Utf8Text(w.source), "WithEvents")
	if withEvents {
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child == nil || child.Kind() != "variable_declarator" {
				continue
			}
			name := declarationNodeName(child, w.source)
			if !state.module || w.moduleKind == "standard" {
				w.add(node, declarationNameNode(child), "invalid_withevents_module", name, "WithEvents declarations are only valid as module-level members of class, document, or UserForm modules.")
				continue
			}
			if !w.validWithEventsShape(child) {
				w.add(node, declarationNameNode(child), "invalid_withevents_shape", name, "WithEvents declarations require a scalar event-capable object type and cannot use As New.")
			}
		}
	}
	if !state.module || !isObjectModuleKind(w.moduleKind) || !strings.EqualFold(visibilityText(node, w.source), "Public") {
		return
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || child.Kind() != "variable_declarator" {
			continue
		}
		// WithEvents shape diagnostics are more actionable than a second
		// public-array/fixed-string finding for the same declarator.
		if withEvents {
			continue
		}
		name := declarationNodeName(child, w.source)
		text := strings.ToLower(child.Utf8Text(w.source))
		if variableIsArray(child, w.source) {
			w.add(node, declarationNameNode(child), "invalid_public_object_member", name, "Public array fields are not valid in object modules.")
			continue
		}
		if variableIsFixedString(text) {
			w.add(node, declarationNameNode(child), "invalid_public_object_member", name, "Public fixed-length String fields are not valid in object modules.")
		}
	}
}

func (w *moduleKindDiagnosticWalker) validWithEventsShape(node *tree_sitter.Node) bool {
	text := strings.ToLower(node.Utf8Text(w.source))
	if variableIsArray(node, w.source) || strings.Contains(text, "as new ") {
		return false
	}
	typ := strings.ToLower(strings.TrimSpace(typeText(node, w.source)))
	typ = strings.Trim(typ, "[]")
	if typ == "" {
		return false
	}
	switch typ {
	case "object", "variant", "boolean", "byte", "currency", "date", "decimal", "double", "integer", "long", "longlong", "longptr", "single", "string":
		return false
	}
	// Qualified COM types and project classes are object candidates. The lint
	// boundary does not always carry a complete TypeLib/project snapshot;
	// unresolved or ambiguous candidates must remain fail-open rather than being
	// guessed as non-event-capable.
	if strings.Contains(typ, ".") {
		return true
	}
	// A project class is also accepted when the object snapshot confirms it;
	// absence from that snapshot is not proof that the type is non-event-capable.
	return true
}

func (w *moduleKindDiagnosticWalker) add(node, nameNode *tree_sitter.Node, kind, symbol, message string) {
	if w.ctx.Err() != nil {
		return
	}
	rng := vbaast.NodeRange(node)
	if nameNode != nil {
		rng = vbaast.NodeRange(nameNode)
	}
	issue := w.linter.issueAt(w.path, rng, "VB050", "error", message)
	if kind == "invalid_me_context" {
		issue.Code = "VB051"
	}
	issue.Kind = kind
	issue.Symbol = symbol
	w.issues = append(w.issues, issue)
}

func isObjectModuleKind(kind string) bool {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "class", "document", "form":
		return true
	default:
		return false
	}
}

func isModuleProcedureDeclarationKind(kind string) bool {
	switch kind {
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration", "conditional_sub_declaration", "conditional_function_declaration", "conditional_property_declaration":
		return true
	default:
		return false
	}
}

func isMeKeywordNode(node *tree_sitter.Node, source []byte) bool {
	if node == nil || !strings.EqualFold(strings.TrimSpace(node.Utf8Text(source)), "Me") {
		return false
	}
	parent := node.Parent()
	if parent == nil {
		return true
	}
	switch parent.Kind() {
	case "qualified_member_expression", "member_expression":
		// Only the receiver position is the VBA keyword. A member that happens
		// to be named Me is an ordinary identifier and must not become VB051.
		first := parent.NamedChild(0)
		return first != nil && first.StartByte() == node.StartByte() && first.EndByte() == node.EndByte()
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration", "variable_declarator", "const_declarator", "type_declaration", "enum_declaration", "parameter":
		// Declaration names are not uses of the implicit object keyword.
		return false
	default:
		return true
	}
}

func variableIsArray(node *tree_sitter.Node, source []byte) bool {
	name := declarationNameNode(node)
	if name == nil {
		return false
	}
	start, end := int(name.EndByte()), int(node.EndByte())
	if start < 0 || end < start || end > len(source) {
		return false
	}
	tail := strings.TrimSpace(string(source[start:end]))
	if strings.HasPrefix(tail, "(") {
		return true
	}
	// The grammar also permits the array suffix on the declared type (`As
	// Widget()`), in which case it is not adjacent to the variable name.
	lower := strings.ToLower(tail)
	if as := strings.Index(lower, " as "); as >= 0 {
		typeTail := strings.TrimSpace(tail[as+4:])
		return strings.Contains(typeTail, "(") && strings.HasSuffix(typeTail, ")")
	}
	return false
}

func variableIsFixedString(text string) bool {
	idx := strings.Index(text, "as ")
	if idx < 0 {
		return false
	}
	tail := text[idx:]
	return strings.Contains(tail, "string") && strings.Contains(tail, "*")
}
