package lint

import (
	"slices"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// maintainabilityIssues scans one borrowed CST only when at least one optional
// inspection is selected. A recovered tree cannot prove that a block is empty.
func (l Linter) maintainabilityIssues(path string, source []byte, root *tree_sitter.Node) []Issue {
	cfg := l.Config.Lint
	enabled := [...]bool{
		cfg.DetectEmptyIf,
		cfg.DetectEmptyElse,
		cfg.DetectEmptyCase,
		cfg.DetectEmptyFor,
		cfg.DetectEmptyForEach,
		cfg.DetectEmptyDo,
		cfg.DetectEmptyWhile,
		cfg.DetectEmptyProcedure,
		cfg.DetectEmptyModule,
		cfg.DetectLegacyCall,
		cfg.DetectRemComment,
		cfg.DetectErrorStatement,
		cfg.DetectGlobalDeclaration,
		cfg.DetectLetAssignment,
		cfg.DetectIdentifierTypeSuffix,
		cfg.DetectWhileWend,
		cfg.DetectDefType,
		cfg.DetectRedundantStepOne,
		cfg.RequireExplicitStep,
		cfg.DetectRedundantOptionBaseZero,
		cfg.DetectModuleDim,
		cfg.DetectImplicitPublic,
		cfg.DetectMultipleDeclarations,
		cfg.DetectUnusedLabels,
		cfg.DetectStopStatement,
		cfg.DetectOnLocalError,
	}
	if !slices.Contains(enabled[:], true) {
		return nil
	}
	if root == nil || root.HasError() || vbaast.HasMissing(root) {
		return nil
	}
	var issues []Issue
	add := func(node *tree_sitter.Node, code, message string) {
		severity := "information"
		if code == "VB091" {
			severity = "warning"
		}
		issues = append(issues, l.issueAt(path, vbaast.NodeRange(node), code, severity, message))
	}
	if cfg.DetectEmptyModule && !hasExecutable(root, source) {
		add(root, "VB075", "Module has no executable statements.")
	}
	var walk func(*tree_sitter.Node, bool)
	walk = func(node *tree_sitter.Node, inProcedure bool) {
		if node == nil || node.HasError() || node.IsMissing() {
			return
		}
		kind := node.Kind()
		if maintainabilityProcedureKind(kind) {
			inProcedure = true
			if cfg.DetectEmptyProcedure && !hasExecutable(node.ChildByFieldName("body"), source) {
				add(node, "VB074", "Procedure has no executable statements.")
			}
			if cfg.DetectUnusedLabels {
				issues = append(issues, l.unusedLabelIssues(path, source, node)...)
			}
			if cfg.DetectImplicitPublic && (strings.EqualFold(l.moduleKindForPath(path), "standard") || strings.EqualFold(l.moduleKindForPath(path), "class")) && visibilityText(node, source) == "" {
				add(node, "VB088", "Public member has no explicit Public modifier.")
			}
		}
		if kind == "block" && (cfg.DetectEmptyIf || cfg.DetectEmptyElse) {
			l.emptyIfFragments(node, source, add)
		}
		switch kind {
		case "case_clause":
			if cfg.DetectEmptyCase && !hasExecutable(node.ChildByFieldName("body"), source) {
				add(node, "VB069", "Case branch has no executable statements.")
			}
		case "for_statement":
			if cfg.DetectEmptyFor && !hasExecutable(node.ChildByFieldName("body"), source) {
				add(node, "VB070", "For loop has no executable statements.")
			}
			step := node.ChildByFieldName("step")
			if step == nil && cfg.RequireExplicitStep {
				add(node, "VB085", "For loop omits Step.")
			} else if step != nil && cfg.DetectRedundantStepOne && strings.TrimSpace(step.Utf8Text(source)) == "1" {
				add(step, "VB084", "Step 1 restates the default step.")
			}
		case "for_each_statement":
			if cfg.DetectEmptyForEach && !hasExecutable(node.ChildByFieldName("body"), source) {
				add(node, "VB071", "For Each loop has no executable statements.")
			}
		case "do_statement":
			if cfg.DetectEmptyDo && !hasExecutable(node.ChildByFieldName("body"), source) {
				add(node, "VB072", "Do loop has no executable statements.")
			}
		case "while_statement":
			if cfg.DetectEmptyWhile && !hasExecutable(node.ChildByFieldName("body"), source) {
				add(node, "VB073", "While loop has no executable statements.")
			}
			if cfg.DetectWhileWend {
				add(node, "VB082", "Use Do...Loop instead of While...Wend.")
			}
		case "call_statement":
			if cfg.DetectErrorStatement && startsWithKeyword(node.Utf8Text(source), "Error") {
				add(node, "VB078", "Use Err.Raise instead of Error.")
			}
			if cfg.DetectLegacyCall && startsWithKeyword(node.Utf8Text(source), "Call") {
				add(node, "VB076", "Explicit Call keyword is optional.")
			}
		case "comment":
			if cfg.DetectRemComment && startsWithKeyword(node.Utf8Text(source), "Rem") {
				add(node, "VB077", "Use an apostrophe comment instead of Rem.")
			}
		case "on_error_statement":
			if cfg.DetectOnLocalError && startsWithKeyword(node.Utf8Text(source), "On Local Error") {
				add(node, "VB092", "On Local Error uses a legacy qualifier; use On Error.")
			}
		case "variable_declaration", "const_declaration", "enum_declaration":
			if cfg.DetectGlobalDeclaration && strings.EqualFold(visibilityText(node, source), "Global") {
				add(node, "VB079", "Use Public instead of Global.")
			}
			if cfg.DetectModuleDim && kind == "variable_declaration" && !inProcedure && startsWithKeyword(node.Utf8Text(source), "Dim") {
				add(node, "VB087", "Use Private for a module-scope Dim declaration.")
			}
			if cfg.DetectMultipleDeclarations && (kind == "variable_declaration" || kind == "const_declaration") {
				declarators, typed := 0, 0
				for i := range node.NamedChildCount() {
					child := node.NamedChild(i)
					if child.Kind() == "variable_declarator" || child.Kind() == "const_declarator" {
						declarators++
						if typeText(child, source) != "" {
							typed++
						}
					}
				}
				mixed := typed > 0 && typed < declarators
				if declarators > 1 && (kind != "variable_declaration" || !mixed || !cfg.DetectMultipleDeclaratorClarity) {
					add(node, "VB089", "Split multiple declarations into separate statements.")
				}
			}
		case "let_statement":
			if cfg.DetectLetAssignment {
				add(node, "VB080", "Explicit Let keyword is optional.")
			}
		case "def_type_statement":
			if cfg.DetectDefType {
				add(node, "VB083", "Declare types explicitly instead of using DefType.")
			}
		case "option_statement":
			if cfg.DetectRedundantOptionBaseZero && strings.EqualFold(strings.Join(strings.Fields(node.Utf8Text(source)), " "), "Option Base 0") {
				add(node, "VB086", "Option Base 0 restates the default.")
			}
		case "stop_statement":
			if cfg.DetectStopStatement {
				add(node, "VB091", "Remove Stop from distributable code.")
			}
		case "identifier", "bang_identifier":
			if cfg.DetectIdentifierTypeSuffix {
				raw := node.Utf8Text(source)
				if len(raw) > 1 && strings.ContainsAny(raw[len(raw)-1:], "$%&!#@") && !callCalleeIdentifier(node) {
					add(node, "VB081", "Use an As clause instead of an identifier type suffix.")
				}
			}
		}
		for i := range node.NamedChildCount() {
			walk(node.NamedChild(i), inProcedure)
		}
	}
	walk(root, false)
	return issues
}

func startsWithKeyword(text, keyword string) bool {
	text = strings.TrimSpace(text)
	return len(text) >= len(keyword) && strings.EqualFold(text[:len(keyword)], keyword) &&
		(len(text) == len(keyword) || text[len(keyword)] == ' ' || text[len(keyword)] == '\t')
}

func maintainabilityProcedureKind(kind string) bool {
	switch kind {
	case "sub_declaration", "function_declaration", "property_declaration",
		"property_get_declaration", "property_let_declaration", "property_set_declaration":
		return true
	}
	return false
}

func callCalleeIdentifier(node *tree_sitter.Node) bool {
	parent := node.Parent()
	if parent == nil {
		return false
	}
	switch parent.Kind() {
	case "call_expression", "call_statement", "member_expression":
		for _, field := range []string{"callee", "function", "member"} {
			if candidate := parent.ChildByFieldName(field); candidate != nil && candidate.StartByte() == node.StartByte() && candidate.EndByte() == node.EndByte() {
				return true
			}
		}
	}
	return false
}

func hasExecutable(node *tree_sitter.Node, source []byte) bool {
	if node == nil {
		return false
	}
	switch node.Kind() {
	case "comment", "newline", "option_statement", "attribute_statement", "attribute_declaration",
		"variable_declaration", "const_declaration", "type_declaration", "enum_declaration",
		"declare_statement", "declare_sub_statement", "declare_function_statement",
		"def_type_statement", "label_statement", "event_declaration", "event_statement":
		return false
	case "source_file", "program", "module", "block", "preprocessor_block", "single_line_block", "inline_statement_sequence":
		for i := range node.NamedChildCount() {
			if hasExecutable(node.NamedChild(i), source) {
				return true
			}
		}
		return false
	}
	if maintainabilityProcedureKind(node.Kind()) {
		return hasExecutable(node.ChildByFieldName("body"), source)
	}
	if strings.HasPrefix(node.Kind(), "preprocessor_") {
		for i := range node.NamedChildCount() {
			if hasExecutable(node.NamedChild(i), source) {
				return true
			}
		}
		return false
	}
	return true
}

func (l Linter) emptyIfFragments(block *tree_sitter.Node, source []byte, add func(*tree_sitter.Node, string, string)) {
	type frame struct {
		branch *tree_sitter.Node
		empty  bool
	}
	var stack []frame
	flush := func(item frame) {
		if item.branch == nil || !item.empty {
			return
		}
		switch item.branch.Kind() {
		case "if_statement", "elseif_fragment":
			if l.Config.Lint.DetectEmptyIf {
				add(item.branch, "VB067", "If branch has no executable statements.")
			}
		case "else_fragment":
			if l.Config.Lint.DetectEmptyElse {
				add(item.branch, "VB068", "Else branch has no executable statements.")
			}
		}
	}
	for i := range block.NamedChildCount() {
		child := block.NamedChild(i)
		switch child.Kind() {
		case "if_statement":
			if len(stack) > 0 {
				stack[len(stack)-1].empty = false
			}
			stack = append(stack, frame{branch: child, empty: true})
		case "elseif_fragment", "else_fragment":
			if len(stack) > 0 {
				flush(stack[len(stack)-1])
				stack[len(stack)-1] = frame{branch: child, empty: true}
			}
		case "end_if_fragment":
			if len(stack) > 0 {
				flush(stack[len(stack)-1])
				stack = stack[:len(stack)-1]
			}
		default:
			if len(stack) > 0 && hasExecutable(child, source) {
				stack[len(stack)-1].empty = false
			}
		}
	}
}

func (l Linter) unusedLabelIssues(path string, source []byte, procedure *tree_sitter.Node) []Issue {
	refs := make(map[string]bool)
	var labels []*tree_sitter.Node
	var walk func(*tree_sitter.Node)
	walk = func(node *tree_sitter.Node) {
		if node == nil {
			return
		}
		switch node.Kind() {
		case "preprocessor_if":
			// A branch-dependent target is not enough evidence for an unused label.
			labels = nil
			refs["*"] = true
			return
		case "label_statement":
			labels = append(labels, node)
		case "goto_statement", "gosub_statement", "resume_statement", "on_error_statement", "on_goto_statement":
			if node.Kind() == "on_goto_statement" {
				selector := node.ChildByFieldName("selector")
				if selector != nil {
					for i := range node.NamedChildCount() {
						child := node.NamedChild(i)
						if child.StartByte() >= selector.EndByte() {
							refs[strings.ToLower(strings.TrimSpace(child.Utf8Text(source)))] = true
						}
					}
				}
			} else if target := node.ChildByFieldName("target"); target != nil {
				name := strings.ToLower(strings.TrimSpace(target.Utf8Text(source)))
				if node.Kind() != "on_error_statement" || (name != "0" && name != "-1") {
					refs[name] = true
				}
			}
		}
		for i := range node.NamedChildCount() {
			walk(node.NamedChild(i))
		}
	}
	walk(procedure.ChildByFieldName("body"))
	if refs["*"] {
		return nil
	}
	var issues []Issue
	for _, label := range labels {
		if name := label.ChildByFieldName("name"); name != nil && !refs[strings.ToLower(strings.TrimSpace(name.Utf8Text(source)))] {
			issues = append(issues, l.issueAt(path, vbaast.NodeRange(name), "VB090", "information", "Label is never referenced by a branch."))
		}
	}
	return issues
}
