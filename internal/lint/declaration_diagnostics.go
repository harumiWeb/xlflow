package lint

import (
	"context"
	"fmt"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// declarationRecord is the small, Go-owned declaration view used by the
// compile-equivalent declaration rules. It deliberately lives at the lint
// boundary: symbols.Inspect remains the public symbol projection while this
// record retains scope and conditional-branch facts that the projection does
// not expose.
type declarationRecord struct {
	name      string
	key       string
	kind      string
	scope     string
	parent    string
	accessor  string
	rng       vbaast.Range
	startByte uint
	module    bool
	procedure bool
	path      []declarationBranch
}

type declarationBranch struct {
	group     string
	condition string
	branch    int
}

type declarationIndex struct {
	records []declarationRecord
}

func (l Linter) declarationIssuesContext(ctx context.Context, path string, source []byte, root *tree_sitter.Node) ([]Issue, error) {
	if root == nil {
		return nil, nil
	}
	index := collectDeclarationIndex(ctx, source, root)
	if ctx.Err() != nil {
		return nil, ctx.Err()
	}
	issues := make([]Issue, 0)
	seen := make(map[string][]declarationRecord)
	for _, record := range index.records {
		if ctx.Err() != nil {
			return nil, ctx.Err()
		}
		key := record.scope + "\x00" + record.key
		if record.kind == "option" {
			key = record.scope + "\x00option:" + record.key
		}
		duplicate := false
		for _, previous := range seen[key] {
			if declarationBranchesOverlap(previous.path, record.path) && !legalPropertyAccessorPair(previous, record) {
				duplicate = true
				break
			}
		}
		if duplicate {
			issue := l.issueAt(path, record.rng, "VB046", "error", fmt.Sprintf("Declaration %q conflicts with an earlier declaration in the same scope.", record.name))
			issue.Kind = "duplicate_declaration"
			issue.Symbol = record.name
			issues = append(issues, issue)
		}
		seen[key] = append(seen[key], record)
	}
	signatureIssues, signatureErr := l.procedureSignatureIssuesContext(ctx, path, source, root, index)
	if signatureErr != nil {
		return nil, signatureErr
	}
	issues = append(issues, signatureIssues...)
	issues = append(issues, l.declarationPlacementIssues(path, index.records)...)
	sort.SliceStable(issues, func(i, j int) bool {
		if issues[i].Line != issues[j].Line {
			return issues[i].Line < issues[j].Line
		}
		if issues[i].Column != issues[j].Column {
			return issues[i].Column < issues[j].Column
		}
		return issues[i].Code < issues[j].Code
	})
	return issues, nil
}

func legalPropertyAccessorPair(a, b declarationRecord) bool {
	if a.scope != b.scope || !strings.EqualFold(a.name, b.name) {
		return false
	}
	if a.kind != "procedure" || b.kind != "procedure" {
		return false
	}
	if a.accessor == "" || b.accessor == "" || a.accessor == b.accessor {
		return false
	}
	return true
}

func (l Linter) declarationPlacementIssues(path string, records []declarationRecord) []Issue {
	issues := make([]Issue, 0)
	prior := make([]declarationRecord, 0)
	moduleProcedures := make([]declarationRecord, 0)
	for _, record := range records {
		if record.kind == "option" {
			if record.scope != "module" {
				issue := l.issueAt(path, record.rng, "VB047", "error", "Option statements are only valid at module scope.")
				issue.Kind = "module_only_declaration_in_procedure"
				issue.Symbol = record.name
				issues = append(issues, issue)
				prior = append(prior, record)
				continue
			}
			declarationAfterOption := false
			for _, previous := range prior {
				// VBA permits an Implements clause to precede the Option
				// directives in class modules; it is not an Option-ordering
				// barrier even though it is a module-level declaration node.
				if previous.kind != "option" && previous.kind != "implements" && declarationBranchesOverlap(previous.path, record.path) {
					declarationAfterOption = true
					break
				}
			}
			if declarationAfterOption {
				issue := l.issueAt(path, record.rng, "VB047", "error", "Option statements must appear before declarations.")
				issue.Kind = "option_after_declaration"
				issue.Symbol = record.name
				issues = append(issues, issue)
			}
		} else if record.module && !record.procedure {
			procedureAfterModule := false
			for _, previous := range moduleProcedures {
				if declarationBranchesOverlap(previous.path, record.path) {
					procedureAfterModule = true
					break
				}
			}
			if procedureAfterModule {
				issue := l.issueAt(path, record.rng, "VB047", "error", "Module-level declarations must appear before the first procedure.")
				issue.Kind = "module_declaration_after_procedure"
				issue.Symbol = record.name
				issues = append(issues, issue)
			}
		} else if record.procedure && record.scope != "module" {
			issue := l.issueAt(path, record.rng, "VB047", "error", "Procedure declarations are only valid at module scope.")
			issue.Kind = "nested_procedure_declaration"
			issue.Symbol = record.name
			issues = append(issues, issue)
		} else if record.scope != "module" && isModuleOnlyDeclaration(record.kind) {
			issue := l.issueAt(path, record.rng, "VB047", "error", fmt.Sprintf("%s declarations are only valid at module scope.", declarationKindLabel(record.kind)))
			issue.Kind = "module_only_declaration_in_procedure"
			issue.Symbol = record.name
			issues = append(issues, issue)
		}
		prior = append(prior, record)
		if record.kind == "procedure" && record.module {
			moduleProcedures = append(moduleProcedures, record)
		}
	}
	return issues
}

func isModuleOnlyDeclaration(kind string) bool {
	switch kind {
	case "option", "type", "enum", "declare", "event", "implements":
		return true
	default:
		return false
	}
}

func declarationKindLabel(kind string) string {
	switch kind {
	case "declare":
		return "Declare"
	case "event":
		return "Event"
	case "implements":
		return "Implements"
	case "type":
		return "Type"
	case "enum":
		return "Enum"
	case "option":
		return "Option"
	default:
		return "Module-only"
	}
}

func declarationBranchesOverlap(a, b []declarationBranch) bool {
	sharedBranch := false
	for _, left := range a {
		for _, right := range b {
			if left.group == right.group && left.branch != right.branch {
				return false
			}
			if left.branch == right.branch && (left.group == right.group || (left.condition != "" && left.condition == right.condition)) {
				sharedBranch = true
			}
		}
	}
	// Two declarations guarded by unrelated conditional groups have no
	// provable compile-time overlap without evaluating the expressions.  Keep
	// the high-precision rule fail-open for that case; a shared branch (or an
	// unconditional declaration) is sufficient evidence of coexistence.
	if len(a) > 0 && len(b) > 0 && !sharedBranch {
		return false
	}
	return true
}

func collectDeclarationIndex(ctx context.Context, source []byte, root *tree_sitter.Node) declarationIndex {
	collector := declarationCollector{ctx: ctx, source: source}
	collector.walk(root, declarationWalkState{scope: "module", module: true})
	sort.SliceStable(collector.records, func(i, j int) bool {
		if collector.records[i].startByte != collector.records[j].startByte {
			return collector.records[i].startByte < collector.records[j].startByte
		}
		return collector.records[i].kind < collector.records[j].kind
	})
	return declarationIndex{records: collector.records}
}

type declarationWalkState struct {
	scope     string
	parent    string
	module    bool
	procedure bool
	path      []declarationBranch
}

type declarationCollector struct {
	ctx     context.Context
	source  []byte
	records []declarationRecord
}

func (c *declarationCollector) walk(node *tree_sitter.Node, state declarationWalkState) {
	if node == nil || c.ctx.Err() != nil {
		return
	}
	switch node.Kind() {
	case "preprocessor_if", "type_preprocessor_if":
		c.walkConditional(node, state)
		return
	case "preprocessor_else", "preprocessor_elseif", "type_preprocessor_else", "type_preprocessor_elseif":
		if body := node.ChildByFieldName("body"); body != nil {
			c.walk(body, state)
		}
		return
	}

	kind := node.Kind()
	switch {
	case kind == "option_statement":
		c.add(node, state, "option", optionDisplayName(node.Utf8Text(c.source)), "", false, false)
	case isProcedureDeclarationKind(kind):
		name := declarationProcedureName(node, c.source)
		if name != "" {
			accessor := procedureAccessorForNode(kind, node, c.source)
			c.add(node, state, "procedure", name, accessor, state.module, true)
			c.addProcedureParameters(node, name, state)
		}
		if kind == "conditional_sub_declaration" || kind == "conditional_function_declaration" || kind == "conditional_property_declaration" {
			c.walkConditionalProcedure(node, state)
			return
		}
		state = declarationWalkState{scope: "proc:" + fmt.Sprint(node.StartByte()), parent: name, procedure: true, path: state.path}
	case kind == "variable_declaration":
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child == nil || child.Kind() != "variable_declarator" {
				continue
			}
			name := declarationNodeName(child, c.source)
			if name != "" {
				c.add(child, state, "variable", name, "", state.module, false)
			}
		}
	case kind == "const_declaration":
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child == nil || child.Kind() != "const_declarator" {
				continue
			}
			name := declarationNodeName(child, c.source)
			if name != "" {
				c.add(child, state, "const", name, "", state.module, false)
			}
		}
	case kind == "type_declaration":
		name := declarationNodeName(node, c.source)
		if name != "" {
			c.add(node, state, "type", name, "", state.module, false)
		}
		state = declarationWalkState{scope: fmt.Sprintf("type:%s:%d", canonicalDeclarationKey(name), node.StartByte()), parent: name, module: false, path: state.path}
	case kind == "enum_declaration":
		name := declarationNodeName(node, c.source)
		if name != "" {
			c.add(node, state, "enum", name, "", state.module, false)
		}
		state = declarationWalkState{scope: fmt.Sprintf("enum:%s:%d", canonicalDeclarationKey(name), node.StartByte()), parent: name, module: false, path: state.path}
	case kind == "enum_member":
		name := declarationNodeName(node, c.source)
		if name != "" {
			c.add(node, state, "enum_member", name, "", false, false)
		}
	case kind == "type_member":
		name := declarationNodeName(node, c.source)
		if name != "" {
			c.add(node, state, "type_member", name, "", false, false)
		}
	case kind == "declare_statement" || kind == "declare_sub_statement" || kind == "declare_function_statement":
		c.add(node, state, "declare", declarationNodeName(node, c.source), "", state.module, false)
	case kind == "event_declaration" || kind == "event_statement":
		c.add(node, state, "event", declarationNodeName(node, c.source), "", state.module, false)
	case kind == "implements_statement":
		c.add(node, state, "implements", declarationNodeName(node, c.source), "", state.module, false)
	case kind == "call_statement" && state.scope != "module":
		// The grammar treats an Option directive in a procedure body as a
		// call-shaped recovery node (the Option grammar production is
		// module-only). Keep the AST identity and report the reserved keyword
		// rather than falling back to a text scan.
		if name := node.NamedChild(0); name != nil && name.Kind() == "identifier" && strings.EqualFold(cleanIdentifier(name.Utf8Text(c.source)), "option") {
			c.add(name, state, "option", "Option", "", false, false)
		}
	}

	for i := uint(0); i < node.NamedChildCount(); i++ {
		if c.ctx.Err() != nil {
			return
		}
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		c.walk(child, state)
	}
}

func (c *declarationCollector) addProcedureParameters(node *tree_sitter.Node, procedure string, state declarationWalkState) {
	list := node.ChildByFieldName("parameters")
	if list == nil && strings.HasPrefix(node.Kind(), "conditional_") {
		if header := node.ChildByFieldName("consequence"); header != nil {
			list = header.ChildByFieldName("parameters")
		}
	}
	if list == nil {
		return
	}
	for i := uint(0); i < list.NamedChildCount(); i++ {
		param := list.NamedChild(i)
		if param == nil || param.Kind() != "parameter" {
			continue
		}
		name := declarationNodeName(param, c.source)
		if name != "" {
			c.add(param, declarationWalkState{scope: "proc:" + fmt.Sprint(node.StartByte()), parent: procedure, path: state.path}, "parameter", name, "", false, false)
		}
	}
}

func (c *declarationCollector) walkConditional(node *tree_sitter.Node, state declarationWalkState) {
	group := fmt.Sprint(node.StartByte())
	condition := conditionalExpressionKey(node, c.source)
	if known, value := conditionalConstant(node, c.source); known {
		if value {
			if body := node.ChildByFieldName("body"); body != nil {
				c.walk(body, state)
			}
			return
		}
		// A statically false #If consequence is absent from every compiled
		// configuration.  Its #Else branch, when present, is unconditional
		// only relative to the discarded consequence; surviving alternatives
		// remain mutually exclusive with one another.
		branch := 1
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child == nil || (!strings.Contains(child.Kind(), "else") && !strings.Contains(child.Kind(), "elseif")) {
				continue
			}
			if body := child.ChildByFieldName("body"); body != nil {
				c.walk(body, declarationWalkState{
					scope: state.scope, parent: state.parent, module: state.module,
					procedure: state.procedure, path: appendBranch(state.path, group, condition, branch),
				})
			}
			branch++
		}
		return
	}
	branch := 0
	if body := node.ChildByFieldName("body"); body != nil {
		c.walk(body, declarationWalkState{scope: state.scope, parent: state.parent, module: state.module, procedure: state.procedure, path: appendBranch(state.path, group, condition, branch)})
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || child.Kind() == "newline" {
			continue
		}
		if strings.Contains(child.Kind(), "else") || strings.Contains(child.Kind(), "elseif") {
			branch++
			if body := child.ChildByFieldName("body"); body != nil {
				c.walk(body, declarationWalkState{scope: state.scope, parent: state.parent, module: state.module, procedure: state.procedure, path: appendBranch(state.path, group, condition, branch)})
			}
		}
	}
}

func (c *declarationCollector) walkConditionalProcedure(node *tree_sitter.Node, state declarationWalkState) {
	group := fmt.Sprint(node.StartByte())
	condition := conditionalExpressionKey(node, c.source)
	for branch, field := range []string{"consequence_body", "alternative_body"} {
		body := node.ChildByFieldName(field)
		if body == nil {
			continue
		}
		c.walk(body, declarationWalkState{scope: "proc:" + fmt.Sprint(node.StartByte()), parent: declarationNodeName(node, c.source), procedure: true, path: appendBranch(state.path, group, condition, branch)})
	}
	if body := node.ChildByFieldName("body"); body != nil {
		c.walk(body, declarationWalkState{scope: "proc:" + fmt.Sprint(node.StartByte()), parent: declarationNodeName(node, c.source), procedure: true, path: state.path})
	}
}

func appendBranch(path []declarationBranch, group, condition string, branch int) []declarationBranch {
	result := append([]declarationBranch(nil), path...)
	return append(result, declarationBranch{group: group, condition: condition, branch: branch})
}

func (c *declarationCollector) add(node *tree_sitter.Node, state declarationWalkState, kind, name, accessor string, module, procedure bool) {
	if node == nil || name == "" {
		return
	}
	rng := vbaast.NodeRange(node)
	if nameNode := declarationNameNode(node); nameNode != nil {
		rng = vbaast.NodeRange(nameNode)
	}
	key := canonicalDeclarationKey(name)
	if kind == "option" {
		key = optionKey(node.Utf8Text(c.source))
	}
	c.records = append(c.records, declarationRecord{
		name: name, key: key, kind: kind, scope: state.scope, parent: state.parent,
		accessor: accessor, rng: rng, startByte: node.StartByte(), module: module, procedure: procedure, path: append([]declarationBranch(nil), state.path...),
	})
}

func declarationProcedureName(node *tree_sitter.Node, source []byte) string {
	if node != nil && strings.HasPrefix(node.Kind(), "conditional_") {
		if header := node.ChildByFieldName("consequence"); header != nil {
			return declarationNodeName(header, source)
		}
	}
	return declarationNodeName(node, source)
}

func procedureAccessorForNode(kind string, node *tree_sitter.Node, source []byte) string {
	if node != nil && strings.HasPrefix(kind, "conditional_") {
		if header := node.ChildByFieldName("consequence"); header != nil {
			return procedureAccessor(header.Kind(), header.Utf8Text(source))
		}
	}
	return procedureAccessor(kind, node.Utf8Text(source))
}

func declarationNameNode(node *tree_sitter.Node) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	if strings.HasPrefix(node.Kind(), "conditional_") {
		if header := node.ChildByFieldName("consequence"); header != nil {
			return declarationNameNode(header)
		}
	}
	if name := node.ChildByFieldName("name"); name != nil {
		return name
	}
	switch node.Kind() {
	case "declare_statement", "declare_sub_statement", "declare_function_statement", "event_statement", "event_declaration", "enum_member", "type_member", "variable_declarator", "const_declarator":
		// These declaration forms have a single leading bare identifier when
		// the grammar does not expose a named field. Do not guess from later
		// identifiers such as library names, types, or bounds.
		if child := node.NamedChild(0); child != nil && child.Kind() == "identifier" {
			return child
		}
	case "implements_statement":
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child == nil || child.Kind() != "type_expression" {
				continue
			}
			for j := uint(0); j < child.NamedChildCount(); j++ {
				name := child.NamedChild(j)
				if name != nil && name.Kind() == "identifier" {
					return name
				}
			}
		}
	}
	return nil
}

func conditionalConstant(node *tree_sitter.Node, source []byte) (known, value bool) {
	if node == nil {
		return false, false
	}
	switch conditionalExpressionKey(node, source) {
	case "false", "0":
		return true, false
	case "true", "-1":
		return true, true
	default:
		return false, false
	}
}

func conditionalExpressionKey(node *tree_sitter.Node, source []byte) string {
	if node == nil {
		return ""
	}
	condition := node.ChildByFieldName("condition")
	if condition == nil {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(condition.Utf8Text(source)))
}

func isProcedureDeclarationKind(kind string) bool {
	switch kind {
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration", "conditional_sub_declaration", "conditional_function_declaration", "conditional_property_declaration":
		return true
	default:
		return false
	}
}

func procedureAccessor(kind, text string) string {
	switch kind {
	case "property_get_declaration":
		return "get"
	case "property_let_declaration":
		return "let"
	case "property_set_declaration":
		return "set"
	}
	lower := strings.ToLower(strings.TrimSpace(text))
	if idx := strings.IndexByte(lower, '\n'); idx >= 0 {
		lower = strings.TrimSpace(lower[:idx])
	}
	for _, accessor := range []string{"get", "let", "set"} {
		if strings.HasPrefix(lower, "property "+accessor+" ") {
			return accessor
		}
	}
	return ""
}

func declarationNodeName(node *tree_sitter.Node, source []byte) string {
	if node == nil {
		return ""
	}
	if name := declarationNameNode(node); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	return ""
}

func optionKey(text string) string {
	fields := strings.Fields(strings.TrimSpace(text))
	if len(fields) < 2 {
		return canonicalDeclarationKey(text)
	}
	return canonicalDeclarationKey(fields[1])
}

func optionDisplayName(text string) string {
	fields := strings.Fields(strings.TrimSpace(text))
	if len(fields) >= 2 {
		return strings.Join(fields[:2], " ")
	}
	return strings.TrimSpace(text)
}
