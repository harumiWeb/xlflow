package lint

import (
	"context"
	"strings"
	"unicode"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/gui"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

const callSyntaxDiagnosticCode = "VB059"

// callSyntaxIssuesFromAST reports only call forms whose syntax and context are
// unambiguous in the CST.  In particular, it deliberately does not turn a
// parser recovery node into a compile diagnostic by itself.  Recovery is used
// only as an additional signal for the explicit Call-target case, where the
// source statement and the surrounding AST identify the invalid target.
func (l Linter) callSyntaxIssuesFromAST(ctx context.Context, path, source string, root *tree_sitter.Node, ir *procedureir.DocumentIR) ([]Issue, error) {
	if root == nil {
		return nil, nil
	}
	functions := localFunctionNames(ir)
	state := callSyntaxWalker{
		linter:      l,
		ctx:         ctx,
		path:        path,
		source:      source,
		functions:   functions,
		seen:        make(map[string]bool),
		errorRanges: make([]vbaast.Range, 0),
	}

	// Collect parser-error ranges independently of the diagnostic walk.  The
	// ranges let the explicit-target check associate a recovery token with the
	// exact colon-separated source statement instead of with a whole line.
	state.collectErrors(root)
	if err := state.walkProcedureBodies(root); err != nil {
		return nil, err
	}
	state.invalidExplicitTargets()
	return state.issues, ctx.Err()
}

type callSyntaxWalker struct {
	linter      Linter
	ctx         context.Context
	path        string
	source      string
	functions   map[string]bool
	seen        map[string]bool
	errorRanges []vbaast.Range
	issues      []Issue
}

func localFunctionNames(ir *procedureir.DocumentIR) map[string]bool {
	functions := make(map[string]bool)
	if ir == nil {
		return functions
	}
	for _, procedure := range ir.Procedures {
		switch procedure.Symbol.Kind {
		case procedureir.ProcedureFunction, procedureir.ProcedurePropertyGet, procedureir.ProcedureProperty:
			name := strings.ToLower(cleanIdentifier(procedure.Symbol.Name))
			if name != "" && !procedure.Symbol.Recovered {
				functions[name] = true
			}
		}
	}
	return functions
}

func (w *callSyntaxWalker) collectErrors(node *tree_sitter.Node) {
	if node == nil {
		return
	}
	if node.IsError() || node.IsMissing() {
		w.errorRanges = append(w.errorRanges, vbaast.NodeRange(node))
	}
	for i := uint(0); i < node.ChildCount(); i++ {
		w.collectErrors(node.Child(i))
	}
}

func (w *callSyntaxWalker) walkProcedureBodies(root *tree_sitter.Node) error {
	var walk func(*tree_sitter.Node, bool) error
	walk = func(node *tree_sitter.Node, inBody bool) error {
		if node == nil {
			return nil
		}
		if err := w.ctx.Err(); err != nil {
			return err
		}
		if isProcedureDeclarationKind(node.Kind()) {
			body := node.ChildByFieldName("body")
			if body != nil {
				return walk(body, true)
			}
			return nil
		}
		if inBody {
			switch node.Kind() {
			case "call_statement":
				if err := w.callStatement(node); err != nil {
					return err
				}
			case "assignment_statement", "let_statement":
				w.assignmentFunctionExpression(node)
			case "if_statement", "elseif_fragment":
				w.ifFunctionExpression(node)
			case "identifier":
				w.bareFunctionIdentifier(node)
			}
		}
		for i := uint(0); i < node.NamedChildCount(); i++ {
			if err := walk(node.NamedChild(i), inBody); err != nil {
				return err
			}
		}
		return nil
	}
	return walk(root, false)
}

func (w *callSyntaxWalker) callStatement(node *tree_sitter.Node) error {
	if node == nil || node.IsError() || node.IsMissing() {
		return nil
	}
	statement := statementRangeAtByte(w.source, int(node.StartByte()))
	callee := node.ChildByFieldName("callee")
	if callee == nil {
		return nil
	}
	if callStatementHasExplicitCall(node, callee, w.source) {
		if statement.StartByte != int(node.StartByte()) {
			prefix := strings.TrimSpace(w.source[statement.StartByte:int(node.StartByte())])
			if n := leadingLineNumberLength(prefix); n > 0 {
				prefix = strings.TrimSpace(prefix[n:])
			}
			if prefix != "" && !strings.EqualFold(prefix, "Call") {
				// tree-sitter can expose the member continuation of an explicit
				// qualified call as a second call_statement. The first node owns
				// the statement.
				return nil
			}
		}
		rng := statement
		statementText := ""
		if rng.StartByte >= 0 && rng.EndByte <= len(w.source) {
			statementText = strings.TrimSpace(w.source[rng.StartByte:rng.EndByte])
			if n := leadingLineNumberLength(statementText); n > 0 {
				statementText = strings.TrimSpace(statementText[n:])
			}
		}
		word, rest := firstVBAWord(statementText)
		if !strings.EqualFold(word, "Call") {
			return nil
		}
		targetText := strings.TrimSpace(trimVBAStatementText(rest))
		target, remainder, targetOK := splitVBACallTarget(targetText)
		memberTarget := callee.Kind() == "qualified_member_expression" || callee.Kind() == "implicit_member_expression" || callee.Kind() == "member_expression"
		if memberTarget {
			target = strings.TrimSpace(callee.Utf8Text([]byte(w.source)))
			targetOK = target != ""
			calleeEnd := int(callee.EndByte())
			if calleeEnd >= 0 && calleeEnd <= rng.EndByte {
				remainder = strings.TrimSpace(trimVBAStatementText(w.source[calleeEnd:rng.EndByte]))
			}
		}
		if !targetOK || (!memberTarget && invalidVBAQualifiedTarget(target)) {
			w.add(node, "invalid_call_target", cleanIdentifier(targetText), "The target of an explicit Call statement must be a procedure name or member expression.", "Use a procedure or member name as the Call target.")
			return nil
		}
		if isVBAControlStatement(target) {
			w.add(node, "invalid_call_target", cleanIdentifier(targetText), "The target of an explicit Call statement must be a procedure name or member expression.", "Use a procedure or member name as the Call target.")
			return nil
		}
		arguments := node.ChildByFieldName("arguments")
		if arguments == nil && invalidExplicitCallRemainder(remainder) {
			w.add(node, "invalid_call_target", cleanIdentifier(targetText), "The target of an explicit Call statement must be a procedure name or member expression.", "Use a procedure or member name as the Call target.")
			return nil
		}
		// `Call Procedure` is the legal zero-argument form. An argument
		// suffix without an argument_list is the invalid form this rule owns.
		invalid := arguments == nil && strings.TrimSpace(remainder) != ""
		if arguments != nil {
			nodeEnd := int(node.EndByte())
			argumentsEnd := int(arguments.EndByte())
			if nodeEnd < argumentsEnd {
				invalid = true
			} else {
				remainder := strings.TrimSpace(trimVBAStatementText(w.source[argumentsEnd:nodeEnd]))
				if invalidExplicitCallRemainder(remainder) {
					w.add(node, "invalid_call_target", cleanIdentifier(targetText), "The target of an explicit Call statement must be a procedure name or member expression.", "Use a procedure or member name as the Call target.")
					return nil
				}
				// A qualified member call is represented by tree-sitter as
				// adjacent call_statement nodes (`Call rows(i).Message(...)`).
				// A dot continuation belongs to the same valid target; commas,
				// operators, and other trailing tokens do not.
				invalid = remainder != "" && !strings.HasPrefix(remainder, ".")
				if !invalid && rng.EndByte > nodeEnd {
					extra := strings.TrimSpace(trimVBAStatementText(w.source[nodeEnd:rng.EndByte]))
					if invalidExplicitCallRemainder(extra) {
						w.add(node, "invalid_call_target", cleanIdentifier(targetText), "The target of an explicit Call statement must be a procedure name or member expression.", "Use a procedure or member name as the Call target.")
						return nil
					}
					invalid = extra != "" && !strings.HasPrefix(extra, ".")
				}
			}
			if rng.EndByte < argumentsEnd && !hasVBAContinuationBetween(w.source, rng.EndByte, argumentsEnd) {
				invalid = true
			}
		}
		if invalid {
			w.add(node, "explicit_call_requires_parentheses", "", "Explicit Call statements require parentheses around their argument list.", "Put the argument list in parentheses after the Call target.")
		}
		return nil
	}
	if statement.StartByte != int(node.StartByte()) {
		prefix := strings.TrimSpace(w.source[statement.StartByte:int(node.StartByte())])
		if word, _ := firstVBAWord(prefix); strings.EqualFold(word, "Call") {
			return nil
		}
	}
	if callee.Kind() != "call_expression" {
		return nil
	}
	arguments := callee.ChildByFieldName("arguments")
	if arguments == nil || arguments.Kind() != "argument_list" {
		return nil
	}
	count := argumentSlotCount(arguments)
	switch {
	case count == 0:
		w.add(node, "standalone_empty_parentheses", cleanIdentifier(callTargetText(callee, w.source)), "A standalone call cannot use an empty parenthesized argument list; omit the parentheses.", "Remove the empty parentheses from the standalone call.")
	case count >= 2:
		w.add(node, "standalone_multi_parenthesized", cleanIdentifier(callTargetText(callee, w.source)), "A standalone multi-argument call cannot wrap the full argument list in parentheses.", "Use Call Name(arg1, arg2) or pass parenthesized arguments as Name (arg1), (arg2).")
	}
	return nil
}

func argumentSlotCount(arguments *tree_sitter.Node) int {
	if arguments == nil {
		return 0
	}
	count := int(arguments.NamedChildCount())
	commas := 0
	for i := uint(0); i < arguments.ChildCount(); i++ {
		child := arguments.Child(i)
		if child != nil && child.Kind() == "," {
			commas++
		}
	}
	if commas > 0 && commas+1 > count {
		count = commas + 1
	}
	return count
}

func hasVBAContinuationBetween(source string, start, end int) bool {
	if start < 0 {
		start = 0
	}
	if end > len(source) {
		end = len(source)
	}
	if start >= end {
		return false
	}
	lineStart := start
	if previous := strings.LastIndexByte(source[:start], '\n'); previous >= 0 {
		lineStart = previous + 1
	}
	for lineStart < end {
		nl := strings.IndexByte(source[lineStart:end], '\n')
		if nl < 0 {
			break
		}
		nl += lineStart
		line := strings.TrimSpace(strings.TrimSuffix(source[lineStart:nl], "\r"))
		if strings.HasSuffix(line, "_") {
			return true
		}
		lineStart = nl + 1
	}
	return false
}

func invalidVBAQualifiedTarget(target string) bool {
	target = strings.TrimSpace(target)
	if target == "" || strings.HasPrefix(target, ".") || strings.HasSuffix(target, ".") {
		return true
	}
	for _, part := range strings.Split(target, ".") {
		if part == "" {
			return true
		}
		first, _ := utf8.DecodeRuneInString(part)
		if first != '_' && !unicode.IsLetter(first) {
			return true
		}
		word, rest := firstVBAWord(part)
		if word == "" || strings.TrimSpace(rest) != "" || !strings.EqualFold(word, part) {
			return true
		}
	}
	return false
}

func invalidExplicitCallRemainder(remainder string) bool {
	remainder = strings.TrimSpace(remainder)
	if remainder == "" || strings.HasPrefix(remainder, "(") || strings.HasPrefix(remainder, ".") {
		return false
	}
	r, _ := utf8.DecodeRuneInString(remainder)
	switch r {
	case '+', '-', '*', '/', '\\', '=', '<', '>', '&':
		return true
	default:
		return false
	}
}

func (w *callSyntaxWalker) assignmentFunctionExpression(node *tree_sitter.Node) {
	right := node.ChildByFieldName("right")
	if right == nil {
		right = node.ChildByFieldName("value")
	}
	if right == nil || right.Kind() != "identifier" || !w.isFunction(right) {
		return
	}
	if !looksLikeBareFunctionArguments(sourceTailAfter(w.source, int(right.EndByte()))) {
		return
	}
	w.add(right, "function_expression_requires_parentheses", cleanIdentifier(right.Utf8Text([]byte(w.source))), "A Function call used in an expression requires parentheses around its argument list.", "Write the Function call as Name(arg1, arg2).")
}

func (w *callSyntaxWalker) ifFunctionExpression(node *tree_sitter.Node) {
	condition := node.ChildByFieldName("condition")
	if condition == nil || condition.Kind() != "identifier" || !w.isFunction(condition) {
		return
	}
	tail := truncateVBAKeyword(sourceTailAfter(w.source, int(condition.EndByte())), "Then")
	if !looksLikeBareFunctionArguments(tail) {
		return
	}
	w.add(condition, "function_expression_requires_parentheses", cleanIdentifier(condition.Utf8Text([]byte(w.source))), "A Function call used in an expression requires parentheses around its argument list.", "Write the Function call as Name(arg1, arg2).")
}

func (w *callSyntaxWalker) bareFunctionIdentifier(node *tree_sitter.Node) {
	if node == nil || !w.isFunction(node) || w.isCallCallee(node) {
		return
	}
	parent := node.Parent()
	if parent == nil || parent.Kind() != "unparenthesized_argument_list" {
		return
	}
	if !looksLikeBareFunctionArguments(sourceTailAfter(w.source, int(node.EndByte()))) {
		return
	}
	w.add(node, "function_expression_requires_parentheses", cleanIdentifier(node.Utf8Text([]byte(w.source))), "A Function call used in an expression requires parentheses around its argument list.", "Write the Function call as Name(arg1, arg2).")
}

func (w *callSyntaxWalker) isFunction(node *tree_sitter.Node) bool {
	return node != nil && w.functions[strings.ToLower(cleanIdentifier(node.Utf8Text([]byte(w.source))))]
}

func (w *callSyntaxWalker) isCallCallee(node *tree_sitter.Node) bool {
	for parent := node.Parent(); parent != nil; parent = parent.Parent() {
		switch parent.Kind() {
		case "call_expression":
			if sameTreeSitterNode(node, parent.ChildByFieldName("function")) {
				return true
			}
		case "call_statement":
			callee := parent.ChildByFieldName("callee")
			if callee != nil && containsTreeSitterNode(callee, node) {
				return true
			}
			return false
		}
	}
	return false
}

func (w *callSyntaxWalker) invalidExplicitTargets() {
	if len(w.errorRanges) == 0 {
		return
	}
	lines := strings.Split(w.source, "\n")
	for lineIndex, line := range lines {
		lineNumber := lineIndex + 1
		parts := splitStatementsWithColumns(gui.StripComment(line))
		for _, part := range parts {
			text := strings.TrimSpace(part.text)
			lineNumberLength := leadingLineNumberLength(text)
			if lineNumberLength > 0 {
				text = strings.TrimSpace(text[lineNumberLength:])
			}
			word, rest := firstVBAWord(text)
			if !strings.EqualFold(word, "Call") {
				continue
			}
			partStart := lineIndexByteOffset(w.source, lineIndex) + part.start
			partEnd := partStart + len(part.text)
			if !w.hasErrorInByteRange(partStart, partEnd) {
				continue
			}
			targetText := strings.TrimSpace(trimVBAStatementText(rest))
			if targetText != "" {
				if target, _, ok := splitVBACallTarget(targetText); ok && !invalidVBAQualifiedTarget(target) && !isVBAControlStatement(target) {
					continue
				}
			}
			rng := vbaast.Range{
				StartLine:   lineNumber,
				StartColumn: part.start + 1,
				EndLine:     lineNumber,
				EndColumn:   part.start + len(part.text) + 1,
				StartByte:   partStart,
				EndByte:     partEnd,
			}
			w.addRange(rng, "invalid_call_target", cleanIdentifier(targetText), "The target of an explicit Call statement must be a procedure name or member expression.", "Use a procedure or member name as the Call target.")
		}
	}
}

func (w *callSyntaxWalker) hasErrorInByteRange(start, end int) bool {
	for _, rng := range w.errorRanges {
		if rng.StartByte < end && rng.EndByte > start {
			return true
		}
	}
	return false
}

func (w *callSyntaxWalker) add(node *tree_sitter.Node, kind, symbol, message, suggestion string) {
	if node == nil {
		return
	}
	rng := statementRangeAtByte(w.source, int(node.StartByte()))
	if rng.StartLine == 0 {
		rng = vbaast.NodeRange(node)
	}
	w.addRange(rng, kind, symbol, message, suggestion)
}

func (w *callSyntaxWalker) addRange(rng vbaast.Range, kind, symbol, message, suggestion string) {
	key := kind + ":" + intString(rng.StartByte) + ":" + intString(rng.EndByte)
	if w.seen[key] {
		return
	}
	w.seen[key] = true
	issue := w.linter.issueAt(w.path, rng, callSyntaxDiagnosticCode, "error", message)
	issue.EndLine = rng.EndLine
	issue.EndColumn = rng.EndColumn
	issue.Kind = kind
	issue.Symbol = symbol
	issue.Suggestion = suggestion
	w.issues = append(w.issues, issue)
}

func callStatementHasExplicitCall(node, callee *tree_sitter.Node, source string) bool {
	if node == nil {
		return false
	}
	if node.StartByte() > 0 {
		lineStart := lineIndexByteOffset(source, strings.Count(source[:int(node.StartByte())], "\n"))
		prefix := strings.TrimSpace(source[lineStart:int(node.StartByte())])
		if strings.EqualFold(prefix, "Call") {
			return true
		}
	}
	if callee == nil || int(node.StartByte()) > int(callee.StartByte()) || int(callee.StartByte()) > len(source) {
		return false
	}
	word, _ := firstVBAWord(source[int(node.StartByte()):int(callee.StartByte())])
	return strings.EqualFold(word, "Call")
}

func callTargetText(node *tree_sitter.Node, source string) string {
	if node == nil {
		return ""
	}
	if fn := node.ChildByFieldName("function"); fn != nil {
		return fn.Utf8Text([]byte(source))
	}
	return node.Utf8Text([]byte(source))
}

func sourceTailAfter(source string, offset int) string {
	if offset < 0 || offset > len(source) {
		return ""
	}
	lineEnd := strings.IndexByte(source[offset:], '\n')
	if lineEnd < 0 {
		lineEnd = len(source)
	} else {
		lineEnd += offset
	}
	return trimVBAStatementText(source[offset:lineEnd])
}

func trimVBAStatementText(text string) string {
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '\'', ':':
			if !inString {
				return strings.TrimSpace(text[:i])
			}
		}
	}
	return strings.TrimSpace(text)
}

func looksLikeBareFunctionArguments(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" || strings.HasPrefix(text, "_") {
		return false
	}
	word, _ := firstVBAWord(text)
	switch strings.ToLower(word) {
	case "", "and", "or", "not", "is", "like", "mod", "eqv", "imp", "xor", "then", "else", "to", "as":
		return false
	}
	first, _ := utf8.DecodeRuneInString(text)
	switch first {
	case '+', '-', '*', '/', '\\', '=', '<', '>', '&', ',', ')', ']':
		return false
	}
	return true
}

func truncateVBAKeyword(text, keyword string) string {
	for offset := 0; offset < len(text); {
		if isVBAIdentifierRune(rune(text[offset])) {
			start := offset
			for offset < len(text) && isVBAIdentifierRune(rune(text[offset])) {
				offset++
			}
			if strings.EqualFold(text[start:offset], keyword) {
				return text[:start]
			}
			continue
		}
		offset++
	}
	return text
}

func firstVBAWord(text string) (string, string) {
	text = strings.TrimLeft(text, " \t\r\n")
	if text == "" {
		return "", ""
	}
	_, size := utf8.DecodeRuneInString(text)
	if !isVBAIdentifierRune([]rune(text[:size])[0]) {
		return "", text
	}
	end := size
	for end < len(text) {
		r, size := utf8.DecodeRuneInString(text[end:])
		if !isVBAIdentifierRune(r) {
			break
		}
		end += size
	}
	return text[:end], text[end:]
}

func statementRangeAtByte(source string, offset int) vbaast.Range {
	if offset < 0 || offset > len(source) {
		return vbaast.Range{}
	}
	lineIndex := strings.Count(source[:offset], "\n")
	lineStart := lineIndexByteOffset(source, lineIndex)
	lineEnd := strings.IndexByte(source[offset:], '\n')
	if lineEnd < 0 {
		lineEnd = len(source)
	} else {
		lineEnd += offset
	}
	line := source[lineStart:lineEnd]
	for _, part := range splitStatementsWithColumns(line) {
		start := lineStart + part.start
		end := start + len(part.text)
		if offset >= start && offset <= end {
			return vbaast.Range{
				StartLine: lineIndex + 1, StartColumn: part.start + 1,
				EndLine: lineIndex + 1, EndColumn: part.start + len(part.text) + 1,
				StartByte: start, EndByte: end,
			}
		}
	}
	return vbaast.Range{}
}

func lineIndexByteOffset(source string, lineIndex int) int {
	if lineIndex <= 0 {
		return 0
	}
	offset := 0
	for i := 0; i < lineIndex; i++ {
		next := strings.IndexByte(source[offset:], '\n')
		if next < 0 {
			return len(source)
		}
		offset += next + 1
	}
	return offset
}

func sameTreeSitterNode(a, b *tree_sitter.Node) bool {
	return a != nil && b != nil && a.Kind() == b.Kind() && a.StartByte() == b.StartByte() && a.EndByte() == b.EndByte()
}

func containsTreeSitterNode(parent, target *tree_sitter.Node) bool {
	if parent == nil || target == nil {
		return false
	}
	if sameTreeSitterNode(parent, target) {
		return true
	}
	for i := uint(0); i < parent.NamedChildCount(); i++ {
		if containsTreeSitterNode(parent.NamedChild(i), target) {
			return true
		}
	}
	return false
}
