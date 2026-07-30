package procedureir

import (
	"path/filepath"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

func BuildSource(opts BuildOptions, source []byte) (DocumentIR, error) {
	path := opts.Path
	if strings.TrimSpace(path) == "" {
		path = "Untitled.bas"
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		return DocumentIR{}, err
	}
	defer doc.Close()
	return BuildParsed(opts, doc)
}

func BuildParsed(opts BuildOptions, doc *vbaast.ParsedDocument) (DocumentIR, error) {
	var result DocumentIR
	err := doc.Read(func(view vbaast.ParsedView) error {
		path := firstNonEmpty(opts.Path, view.Path, "Untitled.bas")
		moduleName := firstNonEmpty(opts.ModuleName, moduleNameFromSource(path, view.Source))
		moduleKind := firstNonEmpty(opts.ModuleKind, moduleKindFromPath(path))
		display := displayPath(opts.RootDir, path)
		builder := documentBuilder{
			source:     view.Source,
			file:       display,
			moduleName: moduleName,
			moduleKind: moduleKind,
			parse: ParseSummary{
				HasError: view.HasError, HasMissing: view.HasMissing,
			},
			nextDeclarationID: 1,
		}
		result = builder.build(view.Root)
		return nil
	})
	return result, err
}

type documentBuilder struct {
	source            []byte
	file              string
	moduleName        string
	moduleKind        string
	parse             ParseSummary
	nextDeclarationID int
}

func (b *documentBuilder) build(root *tree_sitter.Node) DocumentIR {
	return b.buildSinglePass(root)
}

func (b *documentBuilder) procedureSymbol(node *tree_sitter.Node) ProcedureSymbol {
	name := nodeText(childByFieldOrKind(node, "name", "identifier"), b.source)
	kind := procedureKind(node, b.source)
	qualified := name
	if b.moduleName != "" && name != "" {
		qualified = b.moduleName + "." + name
	}
	bodyRange := vbaast.NodeRange(node)
	if body := childByKind(node, "block"); body != nil {
		bodyRange = vbaast.NodeRange(body)
	} else if end := procedureEndNode(node); end != nil {
		bodyRange = zeroRangeAtStart(vbaast.NodeRange(end))
	}
	event, eventKind := ClassifyEvent(b.moduleKind, name)
	return ProcedureSymbol{
		Name: name, QualifiedName: qualified, Kind: kind,
		Visibility: visibilityOfNode(node, b.source),
		Parameters: b.parameters(node), ReturnType: typeText(node, b.source),
		DeclarationRange: vbaast.NodeRange(node), BodyRange: bodyRange,
		IsEventHandler: event, EventKind: eventKind, Recovered: recovered(node),
	}
}

func (b *documentBuilder) parameters(node *tree_sitter.Node) []Parameter {
	list := node.ChildByFieldName("parameters")
	if list == nil {
		list = childByKind(node, "parameter_list")
	}
	if list == nil {
		return []Parameter{}
	}
	out := make([]Parameter, 0, list.NamedChildCount())
	for i := uint(0); i < list.NamedChildCount(); i++ {
		child := list.NamedChild(i)
		if child == nil || child.Kind() != "parameter" {
			continue
		}
		passing := "ByRef"
		if mode := child.ChildByFieldName("passing_mode"); mode != nil && mode.Kind() == "byval_modifier" {
			passing = "ByVal"
		}
		param := Parameter{
			Name: nodeText(childByFieldOrKind(child, "name", "identifier"), b.source),
			Type: typeText(child, b.source), Passing: passing, Range: vbaast.NodeRange(child),
			Optional:   child.ChildByFieldName("optional_modifier") != nil,
			ParamArray: child.ChildByFieldName("paramarray_modifier") != nil,
		}
		if value := child.ChildByFieldName("default_value"); value != nil {
			param.Default = nodeText(value, b.source)
		}
		out = append(out, param)
	}
	return out
}

func (b *documentBuilder) declarations(node *tree_sitter.Node, scope SymbolScope) []Declaration {
	declaratorKind := "variable_declarator"
	kind := "variable"
	if node.Kind() == "const_declaration" {
		declaratorKind, kind = "const_declarator", "const"
	}
	out := []Declaration{}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || child.Kind() != declaratorKind {
			continue
		}
		text := nodeText(child, b.source)
		declVisibility := visibilityOfNode(node, b.source)
		decl := Declaration{
			ID:   b.takeDeclarationID(),
			Name: nodeText(childByFieldOrKind(child, "name", "identifier"), b.source),
			Type: typeText(child, b.source), Scope: declarationScope(scope, declVisibility), Visibility: declVisibility,
			Kind: kind, IsArray: strings.Contains(text, "("), Range: vbaast.NodeRange(child),
			Recovered: recovered(child),
		}
		decl.IsObject = looksObjectType(decl.Type) || hasWord(text, "New")
		out = append(out, decl)
	}
	return out
}

func (b *documentBuilder) simpleDeclaration(node *tree_sitter.Node, scope SymbolScope) (Declaration, bool) {
	name := nodeText(childByFieldOrKind(node, "name", "identifier"), b.source)
	if name == "" {
		return Declaration{}, false
	}
	declVisibility := visibilityOfNode(node, b.source)
	return Declaration{
		ID: b.takeDeclarationID(), Name: name, Type: typeText(node, b.source),
		Scope: declarationScope(scope, declVisibility), Visibility: declVisibility,
		Kind: strings.TrimSuffix(node.Kind(), "_statement"), Range: vbaast.NodeRange(node),
		Recovered: recovered(node),
	}, true
}

func (b *documentBuilder) takeDeclarationID() int {
	id := b.nextDeclarationID
	b.nextDeclarationID++
	return id
}

func declarationScope(scope SymbolScope, visibility string) SymbolScope {
	if scope == ScopeModule && (strings.EqualFold(visibility, "Public") || strings.EqualFold(visibility, "Friend")) {
		return ScopeProject
	}
	return scope
}

func isProcedureNode(kind string) bool {
	switch kind {
	case "sub_declaration", "function_declaration", "property_declaration",
		"property_get_declaration", "property_let_declaration", "property_set_declaration",
		"conditional_sub_declaration", "conditional_function_declaration", "conditional_property_declaration":
		return true
	default:
		return false
	}
}

func procedureKind(node *tree_sitter.Node, source []byte) ProcedureKind {
	switch node.Kind() {
	case "sub_declaration":
		return ProcedureSub
	case "function_declaration":
		return ProcedureFunction
	case "conditional_sub_declaration":
		return ProcedureSub
	case "conditional_function_declaration":
		return ProcedureFunction
	case "conditional_property_declaration":
		text := strings.ToLower(nodeText(node, source))
		switch {
		case strings.Contains(text, "property get"):
			return ProcedurePropertyGet
		case strings.Contains(text, "property let"):
			return ProcedurePropertyLet
		case strings.Contains(text, "property set"):
			return ProcedurePropertySet
		default:
			return ProcedureProperty
		}
	case "property_get_declaration":
		return ProcedurePropertyGet
	case "property_let_declaration":
		return ProcedurePropertyLet
	case "property_set_declaration":
		return ProcedurePropertySet
	}
	text := strings.ToLower(nodeText(node, source))
	switch {
	case strings.Contains(text, "property get"):
		return ProcedurePropertyGet
	case strings.Contains(text, "property let"):
		return ProcedurePropertyLet
	case strings.Contains(text, "property set"):
		return ProcedurePropertySet
	default:
		return ProcedureProperty
	}
}

func statementKind(node *tree_sitter.Node) (StatementKind, bool) {
	kind := strings.ToLower(node.Kind())
	switch kind {
	case "variable_declaration", "const_declaration":
		return StatementDeclaration, true
	case "assignment_statement", "let_statement":
		return StatementAssignment, true
	case "set_statement":
		return StatementSet, true
	case "redim_statement":
		return StatementReDim, true
	case "call_statement":
		return StatementCall, true
	case "label_statement", "line_number_statement":
		return StatementLabel, true
	case "goto_statement":
		return StatementGoTo, true
	case "on_error_statement":
		return StatementOnError, true
	case "resume_statement":
		return StatementResume, true
	case "exit_statement":
		return StatementExit, true
	}
	switch {
	case strings.Contains(kind, "else_if") || strings.Contains(kind, "elseif"):
		return StatementElseIf, true
	case kind == "else_fragment":
		return StatementElse, true
	case strings.Contains(kind, "else") && strings.Contains(kind, "clause"):
		return StatementElse, true
	case strings.Contains(kind, "if") && (strings.Contains(kind, "statement") || strings.Contains(kind, "block")):
		return StatementIf, true
	case strings.Contains(kind, "select") && !strings.Contains(kind, "case"):
		return StatementSelect, true
	case strings.Contains(kind, "case"):
		return StatementCase, true
	case strings.Contains(kind, "for_each"):
		return StatementForEach, true
	case strings.HasPrefix(kind, "for_"):
		return StatementFor, true
	case strings.HasPrefix(kind, "do_"):
		return StatementDo, true
	case strings.HasPrefix(kind, "while_"):
		return StatementWhile, true
	case strings.HasPrefix(kind, "with_"):
		return StatementWith, true
	case kind == "error":
		return StatementRecovered, true
	case strings.HasSuffix(kind, "_statement") && !isNonExecutableStatement(kind):
		return StatementUnknown, true
	default:
		return "", false
	}
}

func isNonExecutableStatement(kind string) bool {
	switch kind {
	case "option_statement", "attribute_statement", "implements_statement", "declare_statement",
		"declare_sub_statement", "declare_function_statement", "event_statement":
		return true
	default:
		return false
	}
}

func expressionKind(kind string) ExpressionKind {
	lower := strings.ToLower(kind)
	switch {
	case lower == "identifier":
		return ExpressionIdentifier
	case strings.HasSuffix(lower, "_literal"):
		return ExpressionLiteral
	case lower == "qualified_member_expression" || lower == "implicit_member_expression" || lower == "member_expression":
		return ExpressionMember
	case lower == "call_expression":
		return ExpressionCall
	case lower == "new_expression":
		return ExpressionNew
	case strings.Contains(lower, "unary_expression"):
		return ExpressionUnary
	case strings.Contains(lower, "binary_expression") || lower == "comparison_expression":
		return ExpressionBinary
	case lower == "parenthesized_expression":
		return ExpressionParentheses
	default:
		return ExpressionUnknown
	}
}

func isExpressionNode(kind string) bool {
	lower := strings.ToLower(kind)
	return lower == "identifier" || strings.HasSuffix(lower, "_literal") ||
		strings.HasSuffix(lower, "_expression")
}

func typeText(node *tree_sitter.Node, source []byte) string {
	if node == nil {
		return ""
	}
	typ := node.ChildByFieldName("type")
	if typ == nil {
		if clause := childByKind(node, "as_type_clause"); clause != nil {
			typ = clause.ChildByFieldName("type")
			if typ == nil {
				typ = clause
			}
		}
	}
	text := nodeText(typ, source)
	if strings.HasPrefix(strings.ToLower(text), "as ") {
		text = strings.TrimSpace(text[3:])
	}
	return text
}

// ClassifyEvent centralizes the event-procedure naming contract used by IR
// consumers. Event declarations are not passed to this helper as procedures.
func ClassifyEvent(moduleKind, name string) (bool, string) {
	lower := strings.ToLower(strings.TrimSpace(name))
	if lower == "auto_open" || lower == "auto_close" {
		return true, "auto"
	}
	switch strings.ToLower(moduleKind) {
	case "document":
		if strings.HasPrefix(lower, "workbook_") {
			return true, "workbook"
		}
		if strings.HasPrefix(lower, "worksheet_") {
			return true, "worksheet"
		}
	case "form":
		if index := strings.LastIndex(lower, "_"); index > 0 && index < len(lower)-1 &&
			!strings.HasPrefix(lower, "test") && !strings.HasSuffix(lower, "_test") {
			return true, "userform"
		}
	}
	return false, ""
}

func procedureEndNode(node *tree_sitter.Node) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child != nil && strings.HasPrefix(child.Kind(), "end_") && strings.HasSuffix(child.Kind(), "_statement") {
			return child
		}
	}
	return nil
}

func zeroRangeAtStart(rng vbaast.Range) vbaast.Range {
	return vbaast.Range{
		StartLine: rng.StartLine, StartColumn: rng.StartColumn,
		EndLine: rng.StartLine, EndColumn: rng.StartColumn,
		StartByte: rng.StartByte, EndByte: rng.StartByte,
	}
}

func visibilityOfNode(node *tree_sitter.Node, source []byte) string {
	if node == nil {
		return ""
	}
	if modifier := node.ChildByFieldName("visibility"); modifier != nil {
		return normalizedVisibility(nodeText(modifier, source))
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		if child.Kind() == "visibility" {
			return normalizedVisibility(nodeText(child, source))
		}
		if child.Kind() != "procedure_modifier" {
			continue
		}
		for j := uint(0); j < child.NamedChildCount(); j++ {
			modifier := child.NamedChild(j)
			if modifier != nil && modifier.Kind() == "visibility" {
				return normalizedVisibility(nodeText(modifier, source))
			}
		}
	}
	header := nodeText(node, source)
	if line, _, ok := strings.Cut(header, "\n"); ok {
		header = strings.TrimSuffix(line, "\r")
	}
	fields := strings.Fields(header)
	if len(fields) == 0 {
		return ""
	}
	return normalizedVisibility(fields[0])
}

func normalizedVisibility(text string) string {
	for _, value := range []string{"Private", "Public", "Friend"} {
		if hasWord(text, value) {
			return value
		}
	}
	return ""
}

func recovered(node *tree_sitter.Node) bool {
	return node != nil && (node.IsError() || node.IsMissing() || node.HasError())
}

func childByFieldNameAny(node *tree_sitter.Node, names ...string) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	for _, name := range names {
		if child := node.ChildByFieldName(name); child != nil {
			return child
		}
	}
	return nil
}

func childByFieldOrKind(node *tree_sitter.Node, field, kind string) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	if child := node.ChildByFieldName(field); child != nil {
		return child
	}
	return childByKind(node, kind)
}

func childByKind(node *tree_sitter.Node, kind string) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child != nil && child.Kind() == kind {
			return child
		}
	}
	return nil
}

func firstNamedChild(node *tree_sitter.Node) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		if child := node.NamedChild(i); child != nil {
			return child
		}
	}
	return nil
}

func sameNode(a, b *tree_sitter.Node) bool {
	return a != nil && b != nil && a.Kind() == b.Kind() && a.StartByte() == b.StartByte() && a.EndByte() == b.EndByte()
}

func nodeText(node *tree_sitter.Node, source []byte) string {
	if node == nil {
		return ""
	}
	return strings.TrimSpace(node.Utf8Text(source))
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}

func moduleNameFromSource(path string, source []byte) string {
	for _, line := range strings.Split(string(source), "\n") {
		trimmed := strings.TrimSpace(line)
		if !strings.HasPrefix(strings.ToLower(trimmed), "attribute vb_name") {
			continue
		}
		if _, value, ok := strings.Cut(trimmed, "="); ok {
			if value = strings.Trim(strings.TrimSpace(value), `"`); value != "" {
				return value
			}
		}
	}
	return strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
}

func moduleKindFromPath(path string) string {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".cls":
		return "class"
	case ".frm":
		return "form"
	default:
		return "standard"
	}
}

func displayPath(root, path string) string {
	if strings.TrimSpace(root) == "" || !filepath.IsAbs(path) {
		return filepath.ToSlash(path)
	}
	rel, err := filepath.Rel(root, path)
	if err != nil {
		return filepath.ToSlash(path)
	}
	return filepath.ToSlash(rel)
}

func hasWord(text, word string) bool {
	for _, part := range strings.FieldsFunc(text, func(r rune) bool {
		return r != '_' && (r < '0' || r > '9') && (r < 'A' || r > 'Z') && (r < 'a' || r > 'z')
	}) {
		if strings.EqualFold(part, word) {
			return true
		}
	}
	return false
}

func cleanIdentifier(text string) string {
	text = strings.TrimSpace(strings.Trim(text, "[]"))
	return strings.TrimRight(text, "$%&#@^!")
}

func looksObjectType(typ string) bool {
	lower := strings.ToLower(cleanIdentifier(typ))
	switch lower {
	case "object", "application", "workbook", "worksheet", "range", "collection", "dictionary":
		return true
	default:
		return strings.Contains(lower, ".")
	}
}

func containsRange(outer, inner vbaast.Range) bool {
	return inner.StartByte >= outer.StartByte && inner.EndByte <= outer.EndByte
}

func containingExpressionID(expressions []Expression, rng vbaast.Range) int {
	id, width := 0, int(^uint(0)>>1)
	for _, expr := range expressions {
		if !containsRange(expr.Range, rng) {
			continue
		}
		if candidate := expr.Range.EndByte - expr.Range.StartByte; candidate < width {
			id, width = expr.ID, candidate
		}
	}
	return id
}

func statementOperand(text string, kind StatementKind) string {
	fields := strings.Fields(strings.TrimSpace(text))
	switch kind {
	case StatementLabel:
		return cleanIdentifier(strings.TrimSuffix(strings.TrimSpace(text), ":"))
	case StatementGoTo:
		if len(fields) >= 2 {
			return cleanIdentifier(fields[1])
		}
	case StatementOnError:
		if len(fields) >= 4 {
			rawTarget := strings.TrimSpace(fields[len(fields)-1])
			if strings.EqualFold(rawTarget, "Next") || rawTarget == "0" {
				return ""
			}
			return cleanIdentifier(rawTarget)
		}
	case StatementResume:
		if len(fields) >= 2 {
			rawTarget := strings.TrimSpace(fields[1])
			if strings.EqualFold(rawTarget, "Next") {
				return ""
			}
			return cleanIdentifier(rawTarget)
		}
	case StatementExit:
		if len(fields) >= 2 {
			return strings.ToLower(fields[1])
		}
	}
	return ""
}

func calleeFromNode(node *tree_sitter.Node, source []byte) Callee {
	text := nodeText(node, source)
	callee := Callee{Text: text}
	switch node.Kind() {
	case "qualified_member_expression":
		if receiver := childByFieldNameAny(node, "receiver", "object"); receiver != nil {
			value := nodeText(receiver, source)
			callee.Receiver = &value
		}
		callee.Member = nodeText(childByFieldNameAny(node, "member", "property"), source)
	case "implicit_member_expression":
		callee.Member = nodeText(childByFieldNameAny(node, "member", "property"), source)
	default:
		callee.Member = lastNamePart(text)
	}
	callee.Member = cleanIdentifier(callee.Member)
	callee.BaseName = callee.Member
	if callee.BaseName == "" {
		callee.BaseName = cleanIdentifier(lastNamePart(text))
		callee.Member = callee.BaseName
	}
	return callee
}

func argumentsFromCallNode(callNode, target *tree_sitter.Node, source []byte) Arguments {
	if callNode.Kind() == "call_expression" {
		return argumentsFromCallExpression(callNode, source)
	}
	if target != nil && target.Kind() == "call_expression" {
		return argumentsFromCallExpression(target, source)
	}
	if list := callNode.ChildByFieldName("arguments"); list != nil {
		return argumentsFromArgumentList(list, source)
	}
	return Arguments{Named: []NamedArgument{}}
}

func argumentsFromCallExpression(node *tree_sitter.Node, source []byte) Arguments {
	if list := node.ChildByFieldName("arguments"); list != nil {
		return argumentsFromArgumentList(list, source)
	}
	out := Arguments{Named: []NamedArgument{}}
	fn := node.ChildByFieldName("function")
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || sameNode(child, fn) {
			continue
		}
		out.Count++
		if child.Kind() == "named_argument" {
			out.Named = append(out.Named, namedArgument(child, source))
		}
	}
	return out
}

func argumentsFromArgumentList(node *tree_sitter.Node, source []byte) Arguments {
	out := Arguments{Named: []NamedArgument{}}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		out.Count++
		if child.Kind() == "named_argument" {
			out.Named = append(out.Named, namedArgument(child, source))
		}
	}
	return out
}

func namedArgument(node *tree_sitter.Node, source []byte) NamedArgument {
	return NamedArgument{
		Name:      cleanIdentifier(nodeText(node.ChildByFieldName("name"), source)),
		ValueText: nodeText(node.ChildByFieldName("value"), source),
	}
}

func lastNamePart(text string) string {
	text = strings.TrimSpace(strings.TrimPrefix(text, "New "))
	for _, separator := range []string{".", "!"} {
		if index := strings.LastIndex(text, separator); index >= 0 {
			text = text[index+1:]
		}
	}
	return text
}
