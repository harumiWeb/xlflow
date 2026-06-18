package lint

import (
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/gui"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type Issue struct {
	Code             string `json:"code"`
	Severity         string `json:"severity"`
	File             string `json:"file"`
	Line             int    `json:"line"`
	Column           int    `json:"column,omitempty"`
	Message          string `json:"message"`
	Kind             string `json:"kind,omitempty"`
	Symbol           string `json:"symbol,omitempty"`
	Suggestion       string `json:"suggestion,omitempty"`
	parserRecoveryOK bool   `json:"-"`
}

type Linter struct {
	RootDir    string
	Config     config.Config
	PathFilter func(string) bool
}

type procedureFrame struct {
	Kind   string
	Name   string
	LineNo int
}

var (
	procedureKinds = map[string]bool{
		"sub_declaration":               true,
		"function_declaration":          true,
		"property_declaration":          true,
		"property_get_declaration":      true,
		"property_let_declaration":      true,
		"property_set_declaration":      true,
		"declare_statement":             true,
		"declare_sub_statement":         true,
		"declare_function_statement":    true,
		"event_declaration":             true,
		"event_statement":               true,
		"external_sub_declaration":      true,
		"external_function_declaration": true,
	}
)

const vb007DisableHint = "If this project intentionally uses dialogs or UserForms, set [lint].forbid_interactive_input = false in xlflow.toml to suppress VB007 for that project. Do this only for genuinely human-only workflows; for dialogs, prefer XlflowUI wrappers with stable dialog ids."

func (l Linter) Run() ([]Issue, error) {
	files, err := l.files()
	if err != nil {
		return nil, err
	}
	issues := make([]Issue, 0)
	for _, file := range files {
		fileIssues, err := l.lintFile(file)
		if err != nil {
			return nil, err
		}
		issues = append(issues, fileIssues...)
	}
	projectIssues, err := l.projectIssues()
	if err != nil {
		return nil, err
	}
	issues = append(issues, projectIssues...)
	sortIssues(issues)
	return issues, nil
}

func (l Linter) files() ([]string, error) {
	dirs := []string{
		l.Config.Src.Modules,
		l.Config.Src.Classes,
		l.Config.Src.Forms,
		l.Config.Src.Workbook,
		"tests",
	}
	var files []string
	for _, dir := range dirs {
		root := filepath.Join(l.RootDir, dir)
		if _, err := os.Stat(root); err != nil {
			if os.IsNotExist(err) {
				continue
			}
			return nil, err
		}
		err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if d.IsDir() {
				return nil
			}
			switch strings.ToLower(filepath.Ext(path)) {
			case ".bas", ".cls", ".frm":
				if !l.shouldIncludeFile(path) {
					return nil
				}
				files = append(files, path)
			}
			return nil
		})
		if err != nil {
			return nil, err
		}
	}
	return files, nil
}

func (l Linter) shouldIncludeFile(path string) bool {
	if l.PathFilter != nil && !l.PathFilter(path) {
		return false
	}
	if !strings.EqualFold(filepath.Ext(path), ".frm") {
		return true
	}
	if !strings.EqualFold(l.Config.UserForm.CodeSource, "sidecar") {
		return true
	}
	formsRoot := filepath.Clean(filepath.Join(l.RootDir, l.Config.Src.Forms))
	cleanPath := filepath.Clean(path)
	if !strings.HasPrefix(strings.ToLower(cleanPath), strings.ToLower(formsRoot)+strings.ToLower(string(os.PathSeparator))) &&
		!strings.EqualFold(cleanPath, formsRoot) {
		return true
	}
	sidecarPath := filepath.Join(formsRoot, "code", strings.TrimSuffix(filepath.Base(cleanPath), filepath.Ext(cleanPath))+".bas")
	if _, err := os.Stat(sidecarPath); err == nil {
		return false
	}
	return true
}

func (l Linter) lintFile(path string) ([]Issue, error) {
	source, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}

	issues, err := l.textSafetyIssues(path, string(source))
	if err != nil {
		return nil, err
	}

	parser, err := vbaast.NewParser()
	if err != nil {
		return nil, err
	}
	defer parser.Close()
	parsed := parser.Parse(path, source)
	defer parsed.Close()

	ctx := astLintContext{linter: l, path: path, source: source}
	ctx.lint(parsed.Root)
	issues = append(issues, ctx.issues...)
	if shouldReportParseIssue(parsed, issues) {
		issues = append(issues, ctx.parseIssue(parsed.Root))
	}
	issues = append(issues, l.flowIssues(path, string(source), parsed.Root)...)

	if l.Config.Lint.ForbidInteractiveInput {
		boundaries, err := gui.Analyzer{RootDir: l.RootDir, Config: l.Config}.AnalyzeFile(path)
		if err != nil {
			return nil, err
		}
		for _, boundary := range boundaries {
			issues = append(issues, Issue{
				Code:       "VB007",
				Severity:   "warning",
				File:       boundary.File,
				Line:       boundary.Line,
				Message:    boundary.Message + " " + boundary.Suggestion + " " + vb007DisableHint,
				Kind:       boundary.Kind,
				Symbol:     boundary.Symbol,
				Suggestion: boundary.Suggestion,
			})
		}
	}
	return issues, nil
}

func (l Linter) textSafetyIssues(path string, source string) ([]Issue, error) {
	var issues []Issue
	procedures := make([]procedureFrame, 0)
	inTypeBlock := false
	var logicalLine strings.Builder
	logicalStartLine := 0
	lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
	for i, line := range lines {
		lineNo := i + 1
		code := gui.StripComment(line)
		detectionCode := maskStringLiterals(code)
		trimmed := strings.TrimSpace(code)
		lower := strings.ToLower(trimmed)
		if isTypeEndLine(lower) {
			inTypeBlock = false
		}
		if l.Config.Lint.DetectImplicitVariant && inTypeBlock && looksImplicitTypeField(trimmed) {
			issue := l.issue(path, lineNo, "VB005", "warning", "Declare an explicit type with As <Type>.")
			issue.parserRecoveryOK = true
			issues = append(issues, issue)
		}
		if isTypeStartLine(lower) {
			inTypeBlock = true
		}
		if missingLineContinuationWhitespace(detectionCode) {
			issues = append(issues, l.issue(path, lineNo, "VB013", "error", "Line-continuation underscore must be preceded by whitespace."))
		}
		if containsTypographicQuote(code) {
			issues = append(issues, l.issue(path, lineNo, "VB008", "error", "Typographic quote found in VBA source. Use straight double quotes for string delimiters before pushing to Excel."))
		}
		if containsLikelyCStyleQuoteEscape(code) {
			issues = append(issues, l.issue(path, lineNo, "VB009", "error", "Likely C-style quote escape found in VBA source. Use doubled quotes, for example \"\"\"\", to represent a quote character."))
		}
		lineForProcedure := detectionCode
		if logicalStartLine == 0 {
			logicalStartLine = lineNo
		}
		if hasValidLineContinuation(detectionCode) {
			logicalLine.WriteString(strings.TrimRight(removeLineContinuationMarker(lineForProcedure), " \t"))
			logicalLine.WriteByte(' ')
			continue
		}
		if logicalLine.Len() > 0 {
			logicalLine.WriteString(lineForProcedure)
			lineForProcedure = logicalLine.String()
		}
		for _, statement := range splitStatements(lineForProcedure) {
			issues = append(issues, l.procedureBoundaryIssues(path, logicalStartLine, statement, &procedures)...)
		}
		logicalLine.Reset()
		logicalStartLine = 0
	}
	if logicalLine.Len() > 0 {
		for _, statement := range splitStatements(logicalLine.String()) {
			issues = append(issues, l.procedureBoundaryIssues(path, logicalStartLine, statement, &procedures)...)
		}
	}
	for _, procedure := range procedures {
		issue := l.issue(path, procedure.LineNo, "VB010", "error", "Unterminated "+procedure.Kind+" procedure.")
		issue.Symbol = procedure.Name
		issues = append(issues, issue)
	}
	return issues, nil
}

type astLintContext struct {
	linter            Linter
	path              string
	source            []byte
	issues            []Issue
	hasOptionExplicit bool
	withDepth         int
}

func (c *astLintContext) lint(root *tree_sitter.Node) {
	c.visit(root, false, false)
	if c.linter.Config.Lint.RequireOptionExplicit && !c.hasOptionExplicit {
		c.issues = append([]Issue{c.linter.issue(c.path, 1, "VB001", "error", "Missing Option Explicit.")}, c.issues...)
	}
}

func (c *astLintContext) visit(node *tree_sitter.Node, inProcedure bool, inType bool) {
	if node == nil {
		return
	}
	kind := node.Kind()
	switch kind {
	case "option_statement":
		if strings.EqualFold(normalizedNodeText(node, c.source), "Option Explicit") {
			c.hasOptionExplicit = true
		}
	case "type_declaration":
		inType = true
	case "qualified_member_expression", "implicit_member_expression":
		c.memberAccessIssue(node)
	case "on_error_statement":
		c.onErrorIssue(node)
	case "variable_declaration":
		c.variableDeclarationIssues(node, inProcedure, inType)
	}
	if procedureKinds[kind] {
		inProcedure = true
	}
	if kind == "with_statement" {
		c.withDepth++
		defer func() { c.withDepth-- }()
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		c.visit(node.NamedChild(i), inProcedure, inType)
	}
}

func (c *astLintContext) memberAccessIssue(node *tree_sitter.Node) {
	property := node.ChildByFieldName("property")
	if property == nil {
		return
	}
	name := cleanIdentifier(property.Utf8Text(c.source))
	switch {
	case c.linter.Config.Lint.ForbidSelect && strings.EqualFold(name, "Select"):
		c.issues = append(c.issues, c.linter.issueAt(c.path, vbaast.NodeRange(property), "VB002", "warning", "Avoid Select. Use direct object references instead."))
	case c.linter.Config.Lint.ForbidActivate && strings.EqualFold(name, "Activate"):
		c.issues = append(c.issues, c.linter.issueAt(c.path, vbaast.NodeRange(property), "VB003", "warning", "Avoid Activate. Use direct object references instead."))
	}
	if c.linter.Config.Lint.DetectNestedWithAmbiguity && node.Kind() == "implicit_member_expression" && c.withDepth >= 2 && isExcelObjectAccessName(name) {
		issue := c.linter.issueAt(c.path, vbaast.NodeRange(property), "VB027", "warning", "Avoid implicit Excel members inside nested With blocks; qualify the target explicitly.")
		issue.Symbol = "." + name
		c.issues = append(c.issues, issue)
	}
}

func (c *astLintContext) onErrorIssue(node *tree_sitter.Node) {
	if !c.linter.Config.Lint.ForbidOnErrorResumeNext {
		return
	}
	if strings.EqualFold(normalizedNodeText(node, c.source), "On Error Resume Next") {
		if c.hasNarrowOnErrorReset(vbaast.NodeRange(node).StartLine) {
			return
		}
		c.issues = append(c.issues, c.linter.issueAt(c.path, vbaast.NodeRange(node), "VB004", "warning", "Avoid On Error Resume Next without a narrow recovery block."))
	}
}

func (c *astLintContext) hasNarrowOnErrorReset(startLine int) bool {
	if startLine <= 0 {
		return false
	}
	lines := normalizedSourceLines(string(c.source))
	seen := 0
	for i := startLine; i < len(lines) && seen < 5; i++ {
		stmt := normalizedCodeLine(lines[i])
		if stmt == "" {
			continue
		}
		seen++
		lower := strings.ToLower(stmt)
		if lower == "on error goto 0" {
			return true
		}
	}
	return false
}

func (c *astLintContext) variableDeclarationIssues(node *tree_sitter.Node, inProcedure bool, inType bool) {
	if c.linter.Config.Lint.ForbidPublicModuleFields && !inProcedure && !inType && strings.EqualFold(visibilityText(node, c.source), "Public") {
		c.issues = append(c.issues, c.linter.issueAt(c.path, vbaast.NodeRange(node), "VB006", "warning", "Avoid Public module variables; pass state explicitly."))
	}
	if c.linter.Config.Lint.DetectMultipleDeclaratorClarity {
		declarators, typed := 0, 0
		for i := uint(0); i < node.NamedChildCount(); i++ {
			child := node.NamedChild(i)
			if child == nil || child.Kind() != "variable_declarator" {
				continue
			}
			declarators++
			if typeText(child, c.source) != "" {
				typed++
			}
		}
		if declarators > 1 && typed > 0 && typed < declarators {
			c.issues = append(c.issues, c.linter.issueAt(c.path, vbaast.NodeRange(node), "VB019", "warning", "In VBA, each declarator needs its own As <Type>; only explicitly typed names receive that type."))
		}
	}
	if !c.linter.Config.Lint.DetectImplicitVariant {
		return
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || child.Kind() != "variable_declarator" || typeText(child, c.source) != "" {
			continue
		}
		c.issues = append(c.issues, c.linter.issueAt(c.path, vbaast.NodeRange(child), "VB005", "warning", "Declare an explicit type with As <Type>."))
	}
}

func (c *astLintContext) parseIssue(root *tree_sitter.Node) Issue {
	r := vbaast.Range{StartLine: 1, StartColumn: 1}
	if node := firstParseProblem(root); node != nil {
		r = vbaast.NodeRange(node)
	}
	return c.linter.issueAt(c.path, r, "VB014", "error", "VBA parser recovered from syntax errors; inspect this source before pushing to Excel.")
}

func (l Linter) flowIssues(path string, source string, root *tree_sitter.Node) []Issue {
	cfg := l.Config.Lint
	if !cfg.DetectConfusingCallSyntax &&
		!cfg.DetectForEachControlType &&
		!cfg.DetectDangerousResume {
		return nil
	}
	lines := normalizedSourceLines(source)
	procedures := sourceProceduresFromAST(root, []byte(source))
	var issues []Issue
	for _, proc := range procedures {
		decls := procedureDeclarations(lines, proc)
		handlerLabels := onErrorHandlerLabels(lines, proc)
		if cfg.DetectConfusingCallSyntax {
			issues = append(issues, l.confusingCallIssues(path, lines, proc)...)
		}
		if cfg.DetectForEachControlType {
			issues = append(issues, l.forEachIssues(path, lines, proc, decls)...)
		}
		if cfg.DetectDangerousResume {
			issues = append(issues, l.dangerousResumeIssues(path, lines, proc, handlerLabels)...)
		}
	}
	return issues
}

type sourceProcedure struct {
	Kind      string
	Name      string
	StartLine int
	EndLine   int
}

type sourceDeclaration struct {
	Name string
	Type string
	Line int
}

func sourceProceduresFromAST(root *tree_sitter.Node, source []byte) []sourceProcedure {
	var procedures []sourceProcedure
	collectSourceProcedures(root, source, &procedures)
	sort.SliceStable(procedures, func(i, j int) bool {
		return procedures[i].StartLine < procedures[j].StartLine
	})
	return procedures
}

func collectSourceProcedures(node *tree_sitter.Node, source []byte, procedures *[]sourceProcedure) {
	if node == nil {
		return
	}
	switch node.Kind() {
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration":
		r := vbaast.NodeRange(node)
		kind := "Sub"
		if strings.Contains(node.Kind(), "function") {
			kind = "Function"
		} else if strings.Contains(node.Kind(), "property") {
			kind = "Property"
		}
		*procedures = append(*procedures, sourceProcedure{
			Kind:      kind,
			Name:      nodeProcedureName(node, source),
			StartLine: r.StartLine,
			EndLine:   r.EndLine,
		})
		return
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		collectSourceProcedures(node.NamedChild(i), source, procedures)
	}
}

func nodeProcedureName(node *tree_sitter.Node, source []byte) string {
	if name := node.ChildByFieldName("name"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	if name := firstNamedChildKind(node, "identifier"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	return ""
}

func procedureDeclarations(lines []string, proc sourceProcedure) map[string]sourceDeclaration {
	decls := map[string]sourceDeclaration{}
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(lines); i++ {
		lineNo := i + 1
		stmt := normalizedCodeLine(lines[i])
		lower := strings.ToLower(stmt)
		if strings.HasPrefix(lower, "dim ") || strings.HasPrefix(lower, "static ") {
			declText := strings.TrimSpace(stmt)
			if strings.HasPrefix(lower, "static ") {
				declText = strings.TrimSpace(declText[len("static "):])
			}
			if strings.HasPrefix(strings.ToLower(declText), "dim ") {
				declText = strings.TrimSpace(declText[len("dim "):])
			}
			for _, part := range strings.Split(declText, ",") {
				name, typ := declarationNameAndType(part)
				if name != "" {
					decls[strings.ToLower(name)] = sourceDeclaration{Name: name, Type: typ, Line: lineNo}
				}
			}
		}
	}
	return decls
}

func declarationNameAndType(text string) (string, string) {
	text = strings.TrimSpace(text)
	if text == "" {
		return "", ""
	}
	lower := strings.ToLower(text)
	namePart := text
	typ := ""
	if idx := strings.Index(lower, " as "); idx >= 0 {
		namePart = strings.TrimSpace(text[:idx])
		typ = strings.TrimSpace(text[idx+4:])
	}
	namePart = strings.TrimSpace(strings.TrimLeft(namePart, "()"))
	nameFields := strings.FieldsFunc(namePart, func(r rune) bool {
		return !isVBAIdentifierRune(r)
	})
	if len(nameFields) == 0 {
		return "", typ
	}
	return cleanIdentifier(nameFields[0]), typ
}

func onErrorHandlerLabels(lines []string, proc sourceProcedure) map[string]bool {
	labels := map[string]bool{}
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(lines); i++ {
		stmt := normalizedCodeLine(lines[i])
		lower := strings.ToLower(stmt)
		if strings.HasPrefix(lower, "on error goto ") && lower != "on error goto 0" {
			label := cleanIdentifier(strings.TrimSpace(stmt[len("on error goto "):]))
			if label != "" {
				labels[strings.ToLower(label)] = true
			}
		}
	}
	return labels
}

func (l Linter) confusingCallIssues(path string, lines []string, proc sourceProcedure) []Issue {
	var issues []Issue
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(lines); i++ {
		stmt := normalizedCodeLine(lines[i])
		name, ok := confusingParenthesizedCall(stmt)
		if !ok {
			continue
		}
		issue := l.issue(path, i+1, "VB022", "warning", "Avoid ambiguous parenthesized call syntax; use Call "+name+"(...) or pass arguments without parentheses.")
		issue.Symbol = name
		issues = append(issues, issue)
	}
	return issues
}

func confusingParenthesizedCall(stmt string) (string, bool) {
	fields := strings.Fields(stmt)
	if len(fields) < 2 || !strings.HasPrefix(fields[1], "(") {
		return "", false
	}
	name := cleanIdentifier(fields[0])
	if name == "" || isStatementKeyword(name) || strings.Contains(stmt, "=") || strings.Contains(name, ".") {
		return "", false
	}
	return name, true
}

func (l Linter) forEachIssues(path string, lines []string, proc sourceProcedure, decls map[string]sourceDeclaration) []Issue {
	var issues []Issue
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(lines); i++ {
		stmt := normalizedCodeLine(lines[i])
		lower := strings.ToLower(stmt)
		if !strings.HasPrefix(lower, "for each ") {
			continue
		}
		rest := strings.TrimSpace(stmt[len("for each "):])
		parts := strings.Fields(rest)
		if len(parts) == 0 {
			continue
		}
		name := cleanIdentifier(parts[0])
		decl, ok := decls[strings.ToLower(name)]
		if !ok {
			issue := l.issue(path, i+1, "VB023", "warning", "Declare the For Each control variable explicitly as Variant or an object type.")
			issue.Symbol = name
			issues = append(issues, issue)
			continue
		}
		if isObviouslyScalarType(decl.Type) {
			issue := l.issue(path, i+1, "VB023", "warning", "For Each control variables should be Variant or object-compatible, not "+decl.Type+".")
			issue.Symbol = name
			issues = append(issues, issue)
		}
	}
	return issues
}

func (l Linter) dangerousResumeIssues(path string, lines []string, proc sourceProcedure, handlerLabels map[string]bool) []Issue {
	var issues []Issue
	inHandler := false
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(lines); i++ {
		stmt := normalizedCodeLine(lines[i])
		if label, ok := labelName(stmt); ok && handlerLabels[strings.ToLower(label)] {
			inHandler = true
			continue
		}
		lower := strings.ToLower(stmt)
		if lower == "resume" || strings.HasPrefix(lower, "resume ") {
			if !inHandler {
				issues = append(issues, l.issue(path, i+1, "VB026", "warning", "Use Resume only inside a clear error-handler block."))
			}
		}
	}
	return issues
}

func labelName(stmt string) (string, bool) {
	stmt = strings.TrimSpace(stmt)
	if !strings.HasSuffix(stmt, ":") || strings.Contains(stmt, " ") {
		return "", false
	}
	name := cleanIdentifier(strings.TrimSuffix(stmt, ":"))
	return name, name != ""
}

func normalizedSourceLines(source string) []string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	return strings.Split(source, "\n")
}

func normalizedCodeLine(line string) string {
	return strings.Join(strings.Fields(maskStringLiterals(gui.StripComment(line))), " ")
}

func isStatementKeyword(name string) bool {
	switch strings.ToLower(name) {
	case "if", "for", "do", "loop", "while", "wend", "select", "case", "with", "set", "let", "call", "return", "resume", "on", "dim", "redim":
		return true
	default:
		return false
	}
}

func isObviouslyScalarType(typ string) bool {
	switch strings.ToLower(cleanIdentifier(typ)) {
	case "byte", "boolean", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal", "date", "string":
		return true
	default:
		return false
	}
}

func isExcelObjectAccessName(name string) bool {
	switch strings.ToLower(cleanIdentifier(name)) {
	case "range", "cells", "rows", "columns":
		return true
	default:
		return false
	}
}

func firstParseProblem(node *tree_sitter.Node) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	if node.IsError() || node.IsMissing() {
		return node
	}
	for i := uint(0); i < node.ChildCount(); i++ {
		if found := firstParseProblem(node.Child(i)); found != nil {
			return found
		}
	}
	return nil
}

func PushBlockingIssues(issues []Issue) []Issue {
	blocking := make([]Issue, 0)
	for _, issue := range issues {
		if issue.Code == "VB008" || issue.Code == "VB009" || issue.Code == "VB010" || issue.Code == "VB011" || issue.Code == "VB012" || issue.Code == "VB013" || issue.Code == "VB014" {
			blocking = append(blocking, issue)
		}
	}
	return blocking
}

func (l Linter) projectIssues() ([]Issue, error) {
	cfg := l.Config.Lint
	if !cfg.DetectScopeShadowing && !cfg.DetectUnusedLocalVariables && !cfg.DetectUnusedPrivateProcedures {
		return nil, nil
	}
	result, err := symbols.Inspect(symbols.Options{
		RootDir:        l.RootDir,
		Config:         l.Config,
		IncludePrivate: true,
		IncludeLabels:  false,
	})
	if err != nil {
		return nil, err
	}
	var issues []Issue
	if cfg.DetectScopeShadowing || cfg.DetectUnusedLocalVariables {
		issues = append(issues, l.symbolScopeIssues(result)...)
	}
	if cfg.DetectUnusedPrivateProcedures {
		callResult, err := calls.Inspect(calls.Options{RootDir: l.RootDir, Config: l.Config})
		if err != nil {
			return nil, err
		}
		issues = append(issues, l.unusedPrivateProcedureIssues(result, callResult)...)
	}
	return issues, nil
}

func (l Linter) symbolScopeIssues(result *symbols.Result) []Issue {
	if result == nil {
		return nil
	}
	var issues []Issue
	sourceCache := map[string]string{}
	for _, file := range result.Files {
		moduleNames := map[string]symbols.Symbol{}
		procedureNames := map[string]symbols.Symbol{}
		for _, sym := range file.Symbols {
			key := strings.ToLower(sym.Name)
			switch sym.Kind {
			case "module_variable", "field", "withevents_field", "const":
				if sym.Parent == "" && sym.Name != "" {
					moduleNames[key] = sym
				}
			case "sub", "function", "property", "property_get", "property_let", "property_set":
				if sym.Name != "" {
					procedureNames[key] = sym
				}
			}
		}
		sameScope := map[string]map[string]symbols.Symbol{}
		for _, sym := range file.Symbols {
			if sym.Kind != "local_variable" && (sym.Kind != "const" || sym.Parent == "") {
				continue
			}
			if l.Config.Lint.DetectScopeShadowing {
				key := strings.ToLower(sym.Name)
				if shadow, ok := moduleNames[key]; ok {
					issue := l.issueForSymbol(sym, "VB018", "warning", "Local declaration shadows module-level declaration "+shadow.Name+".")
					issue.Symbol = sym.Name
					issues = append(issues, issue)
				} else if shadow, ok := procedureNames[key]; ok {
					issue := l.issueForSymbol(sym, "VB018", "warning", "Local declaration shadows procedure "+shadow.Name+".")
					issue.Symbol = sym.Name
					issues = append(issues, issue)
				}
				scopeKey := sym.File + ":" + sym.Parent
				if sameScope[scopeKey] == nil {
					sameScope[scopeKey] = map[string]symbols.Symbol{}
				}
				if prior, ok := sameScope[scopeKey][key]; ok {
					issue := l.issueForSymbol(sym, "VB018", "warning", "Declaration shadows another declaration in the same procedure at line "+intString(prior.StartLine)+".")
					issue.Symbol = sym.Name
					issues = append(issues, issue)
				} else {
					sameScope[scopeKey][key] = sym
				}
			}
			if l.Config.Lint.DetectUnusedLocalVariables && !isIgnoredLocalName(sym.Name) {
				source, ok := sourceCache[sym.File]
				if !ok {
					body, err := os.ReadFile(filepath.Join(l.RootDir, filepath.FromSlash(sym.File)))
					if err != nil {
						continue
					}
					source = string(body)
					sourceCache[sym.File] = source
				}
				if !localNameReferenced(source, sym, procedureEndLineForSymbol(file.Symbols, sym)) {
					issue := l.issueForSymbol(sym, "VB020", "warning", "Procedure-local variable is declared but never referenced.")
					issue.Symbol = sym.Name
					issues = append(issues, issue)
				}
			}
		}
		if l.Config.Lint.DetectScopeShadowing {
			issues = append(issues, l.parameterShadowingIssues(file, moduleNames, procedureNames)...)
		}
	}
	return issues
}

func (l Linter) parameterShadowingIssues(file symbols.FileResult, moduleNames, procedureNames map[string]symbols.Symbol) []Issue {
	var issues []Issue
	for _, proc := range file.Symbols {
		if !procedureSymbolKind(proc.Kind) {
			continue
		}
		seen := map[string]symbols.Parameter{}
		for _, param := range proc.Parameters {
			key := strings.ToLower(param.Name)
			if key == "" {
				continue
			}
			if shadow, ok := moduleNames[key]; ok {
				issue := l.issueForRel(file.Path, proc.StartLine, proc.StartColumn, "VB018", "warning", "Parameter shadows module-level declaration "+shadow.Name+".")
				issue.Symbol = param.Name
				issues = append(issues, issue)
			} else if shadow, ok := procedureNames[key]; ok && !strings.EqualFold(shadow.Name, proc.Name) {
				issue := l.issueForRel(file.Path, proc.StartLine, proc.StartColumn, "VB018", "warning", "Parameter shadows procedure "+shadow.Name+".")
				issue.Symbol = param.Name
				issues = append(issues, issue)
			}
			if _, ok := seen[key]; ok {
				issue := l.issueForRel(file.Path, proc.StartLine, proc.StartColumn, "VB018", "warning", "Parameter shadows another parameter in the same procedure.")
				issue.Symbol = param.Name
				issues = append(issues, issue)
			}
			seen[key] = param
		}
	}
	return issues
}

func (l Linter) unusedPrivateProcedureIssues(symbolResult *symbols.Result, callResult *calls.Result) []Issue {
	if symbolResult == nil {
		return nil
	}
	called := map[string]bool{}
	if callResult != nil {
		for _, call := range callResult.Calls {
			called[strings.ToLower(call.Callee.BaseName)] = true
			called[strings.ToLower(call.Callee.Member)] = true
			called[strings.ToLower(call.Callee.Text)] = true
		}
	}
	var issues []Issue
	for _, file := range symbolResult.Files {
		for _, sym := range file.Symbols {
			if !procedureSymbolKind(sym.Kind) || !strings.EqualFold(sym.Visibility, "Private") || isKnownCallbackProcedure(sym.Name) {
				continue
			}
			if called[strings.ToLower(sym.Name)] || called[strings.ToLower(sym.Module+"."+sym.Name)] {
				continue
			}
			issue := l.issueForSymbol(sym, "VB021", "warning", "Private procedure is not called from parsed source.")
			issue.Symbol = sym.Name
			issues = append(issues, issue)
		}
	}
	return issues
}

func (l Linter) issueForSymbol(sym symbols.Symbol, code, severity, message string) Issue {
	return l.issueForRel(sym.File, sym.StartLine, sym.StartColumn, code, severity, message)
}

func (l Linter) issueForRel(file string, line, column int, code, severity, message string) Issue {
	if line == 0 {
		line = 1
	}
	return Issue{
		Code:     code,
		Severity: severity,
		File:     filepath.ToSlash(file),
		Line:     line,
		Column:   column,
		Message:  message,
	}
}

func sortIssues(issues []Issue) {
	sort.SliceStable(issues, func(i, j int) bool {
		a, b := issues[i], issues[j]
		if a.File != b.File {
			return a.File < b.File
		}
		if a.Line != b.Line {
			return a.Line < b.Line
		}
		if a.Column != b.Column {
			return a.Column < b.Column
		}
		return a.Code < b.Code
	})
}

func procedureSymbolKind(kind string) bool {
	switch kind {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return true
	default:
		return false
	}
}

func isIgnoredLocalName(name string) bool {
	name = strings.ToLower(strings.TrimSpace(name))
	return name == "" || name == "_" || strings.HasPrefix(name, "unused") || strings.HasPrefix(name, "ignore")
}

func procedureEndLineForSymbol(all []symbols.Symbol, sym symbols.Symbol) int {
	if sym.Parent == "" {
		return sym.EndLine
	}
	for _, candidate := range all {
		if !procedureSymbolKind(candidate.Kind) {
			continue
		}
		if !strings.EqualFold(candidate.Name, sym.Parent) {
			continue
		}
		if candidate.StartLine <= sym.StartLine && sym.StartLine <= candidate.EndLine {
			return candidate.EndLine
		}
	}
	return sym.EndLine
}

func localNameReferenced(source string, sym symbols.Symbol, endLine int) bool {
	lines := normalizedSourceLines(source)
	needle := sym.Name
	count := 0
	for i := sym.StartLine - 1; i < len(lines); i++ {
		lineNo := i + 1
		if sym.Parent != "" && endLine > 0 && lineNo > endLine {
			break
		}
		line := normalizedCodeLine(lines[i])
		if lineNo == sym.StartLine {
			line = stripDeclarationPrefix(line)
		}
		if hasWord(line, needle) {
			count++
		}
		if count > 0 {
			return true
		}
	}
	return false
}

func stripDeclarationPrefix(line string) string {
	lower := strings.ToLower(line)
	for _, prefix := range []string{"dim ", "static ", "private ", "public "} {
		if strings.HasPrefix(lower, prefix) {
			if idx := strings.Index(line, ","); idx >= 0 {
				return line[idx+1:]
			}
			return ""
		}
	}
	return line
}

func isKnownCallbackProcedure(name string) bool {
	lower := strings.ToLower(name)
	if strings.Contains(lower, "_") {
		return true
	}
	switch lower {
	case "auto_open", "auto_close", "workbook_open", "workbook_beforeclose":
		return true
	default:
		return false
	}
}

func intString(value int) string {
	if value == 0 {
		return "0"
	}
	digits := []byte{}
	for value > 0 {
		digits = append([]byte{byte('0' + value%10)}, digits...)
		value /= 10
	}
	return string(digits)
}

func hasSpecificSyntaxIssue(issues []Issue) bool {
	for _, issue := range issues {
		if issue.Code == "VB008" || issue.Code == "VB009" || issue.Code == "VB010" || issue.Code == "VB011" || issue.Code == "VB012" || issue.Code == "VB013" {
			return true
		}
	}
	return false
}

func shouldReportParseIssue(parsed *vbaast.ParseResult, issues []Issue) bool {
	if parsed == nil || (!parsed.HasError && !parsed.HasMissing) || hasSpecificSyntaxIssue(issues) {
		return false
	}
	problemLines := parseProblemLines(parsed.Root)
	if len(problemLines) == 0 {
		return true
	}
	for line := range problemLines {
		if !hasIssueAtLine(issues, "VB005", line) {
			return true
		}
	}
	return false
}

func parseProblemLines(root *tree_sitter.Node) map[int]bool {
	lines := make(map[int]bool)
	collectParseProblemLines(root, lines)
	return lines
}

func collectParseProblemLines(node *tree_sitter.Node, lines map[int]bool) {
	if node == nil {
		return
	}
	if node.IsError() || node.IsMissing() {
		r := vbaast.NodeRange(node)
		if r.StartLine > 0 {
			lines[r.StartLine] = true
		}
	}
	for i := uint(0); i < node.ChildCount(); i++ {
		collectParseProblemLines(node.Child(i), lines)
	}
}

func hasIssueAtLine(issues []Issue, code string, line int) bool {
	for _, issue := range issues {
		if issue.Code == code && issue.Line == line && issue.parserRecoveryOK {
			return true
		}
	}
	return false
}

func (l Linter) issue(path string, line int, code, severity, message string) Issue {
	return l.issueAt(path, vbaast.Range{StartLine: line}, code, severity, message)
}

func (l Linter) issueAt(path string, r vbaast.Range, code, severity, message string) Issue {
	file, err := filepath.Rel(l.RootDir, path)
	if err != nil {
		file = path
	}
	if r.StartLine == 0 {
		r.StartLine = 1
	}
	return Issue{
		Code:     code,
		Severity: severity,
		File:     filepath.ToSlash(file),
		Line:     r.StartLine,
		Column:   r.StartColumn,
		Message:  message,
	}
}

func containsTypographicQuote(line string) bool {
	return strings.ContainsAny(line, "“”‘’")
}

func containsLikelyCStyleQuoteEscape(line string) bool {
	for i := 0; i < len(line)-2; i++ {
		if line[i] != '\\' || line[i+1] != '"' {
			continue
		}
		quoteCount := 0
		for j := i + 1; j < len(line) && line[j] == '"'; j++ {
			quoteCount++
		}
		if quoteCount >= 2 {
			return true
		}
	}
	return false
}

func isTypeStartLine(lower string) bool {
	fields := strings.Fields(lower)
	if len(fields) == 0 {
		return false
	}
	if fields[0] == "type" {
		return true
	}
	return len(fields) > 1 && (fields[0] == "private" || fields[0] == "public") && fields[1] == "type"
}

func isTypeEndLine(lower string) bool {
	fields := strings.Fields(lower)
	return len(fields) == 2 && fields[0] == "end" && fields[1] == "type"
}

func looksImplicitTypeField(line string) bool {
	lower := strings.ToLower(strings.TrimSpace(line))
	if lower == "" || strings.Contains(lower, " as ") || isConditionalCompilationDirective(lower) {
		return false
	}
	return !isTypeStartLine(lower) && !isTypeEndLine(lower)
}

func isConditionalCompilationDirective(line string) bool {
	return strings.HasPrefix(line, "#if ") ||
		strings.HasPrefix(line, "#elseif ") ||
		line == "#else" ||
		line == "#end if"
}

func normalizedNodeText(node *tree_sitter.Node, source []byte) string {
	return strings.Join(strings.Fields(node.Utf8Text(source)), " ")
}

func visibilityText(node *tree_sitter.Node, source []byte) string {
	if visibility := node.ChildByFieldName("visibility"); visibility != nil {
		return normalizeKeyword(visibility.Utf8Text(source))
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child != nil && child.Kind() == "visibility" {
			return normalizeKeyword(child.Utf8Text(source))
		}
	}
	text := node.Utf8Text(source)
	for _, word := range []string{"Public", "Private", "Friend"} {
		if hasWord(text, word) {
			return word
		}
	}
	return ""
}

func typeText(node *tree_sitter.Node, source []byte) string {
	asType := node.ChildByFieldName("type")
	if asType == nil {
		asType = firstNamedChildKind(node, "as_type_clause")
	}
	if asType == nil {
		return ""
	}
	if typeExpr := asType.ChildByFieldName("type"); typeExpr != nil {
		return strings.TrimSpace(typeExpr.Utf8Text(source))
	}
	text := strings.TrimSpace(asType.Utf8Text(source))
	if strings.HasPrefix(strings.ToLower(text), "as ") {
		return strings.TrimSpace(text[3:])
	}
	return text
}

func firstNamedChildKind(node *tree_sitter.Node, kind string) *tree_sitter.Node {
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child != nil && child.Kind() == kind {
			return child
		}
	}
	return nil
}

func cleanIdentifier(text string) string {
	text = strings.TrimSpace(text)
	text = strings.Trim(text, "[]")
	text = strings.TrimRight(text, "$%&#@^!")
	return text
}

func normalizeKeyword(text string) string {
	text = strings.TrimSpace(text)
	if text == "" {
		return ""
	}
	lower := strings.ToLower(text)
	return strings.ToUpper(lower[:1]) + lower[1:]
}

func hasWord(text, word string) bool {
	fields := strings.FieldsFunc(text, func(r rune) bool {
		return !isVBAIdentifierRune(r)
	})
	for _, field := range fields {
		if strings.EqualFold(field, word) {
			return true
		}
	}
	return false
}

func isVBAIdentifierRune(r rune) bool {
	switch r {
	case '_', '$', '%', '&', '!', '#', '@', '^':
		return true
	}
	return r >= '0' && r <= '9' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z'
}

func (l Linter) procedureBoundaryIssues(path string, lineNo int, line string, procedures *[]procedureFrame) []Issue {
	var issues []Issue
	if endKind, ok := procedureEndKind(line); ok {
		if len(*procedures) == 0 {
			issues = append(issues, l.issue(path, lineNo, "VB011", "error", "Unexpected End "+endKind+" without a matching procedure."))
			return issues
		}
		top := (*procedures)[len(*procedures)-1]
		*procedures = (*procedures)[:len(*procedures)-1]
		if top.Kind != endKind {
			issue := l.issue(path, lineNo, "VB012", "error", "Mismatched End "+endKind+" for "+top.Kind+" procedure.")
			issue.Symbol = top.Name
			issues = append(issues, issue)
		}
		return issues
	}
	if start, ok := procedureStart(line, lineNo); ok {
		*procedures = append(*procedures, start)
	}
	return issues
}

func procedureEndKind(line string) (string, bool) {
	fields := lowerFields(line)
	if len(fields) < 2 || fields[0] != "end" {
		return "", false
	}
	switch fields[1] {
	case "sub":
		return "Sub", true
	case "function":
		return "Function", true
	case "property":
		return "Property", true
	default:
		return "", false
	}
}

func procedureStart(line string, lineNo int) (procedureFrame, bool) {
	fields := lowerFields(line)
	names := cleanedFields(line)
	if len(fields) == 0 || fields[0] == "rem" || strings.HasPrefix(fields[0], "#") {
		return procedureFrame{}, false
	}
	index := 0
	for index < len(fields) {
		switch fields[index] {
		case "public", "private", "friend", "static":
			index++
		default:
			goto afterModifiers
		}
	}
afterModifiers:
	if index >= len(fields) || fields[index] == "declare" {
		return procedureFrame{}, false
	}
	switch fields[index] {
	case "sub":
		return procedureFrame{Kind: "Sub", Name: procedureName(names, index+1), LineNo: lineNo}, true
	case "function":
		return procedureFrame{Kind: "Function", Name: procedureName(names, index+1), LineNo: lineNo}, true
	case "property":
		if index+1 >= len(fields) {
			return procedureFrame{}, false
		}
		switch fields[index+1] {
		case "get", "let", "set":
			return procedureFrame{Kind: "Property", Name: procedureName(names, index+2), LineNo: lineNo}, true
		default:
			return procedureFrame{}, false
		}
	default:
		return procedureFrame{}, false
	}
}

func procedureName(fields []string, index int) string {
	if index >= len(fields) {
		return ""
	}
	name, _, _ := strings.Cut(fields[index], "(")
	return name
}

func lowerFields(line string) []string {
	fields := cleanedFields(line)
	for i, field := range fields {
		fields[i] = strings.ToLower(field)
	}
	return fields
}

func cleanedFields(line string) []string {
	fields := strings.Fields(line)
	for i, field := range fields {
		fields[i] = strings.Trim(field, "(),")
	}
	return fields
}

func maskStringLiterals(line string) string {
	var b strings.Builder
	b.Grow(len(line))
	inString := false
	for i := 0; i < len(line); i++ {
		if line[i] != '"' {
			if inString {
				b.WriteByte(' ')
			} else {
				b.WriteByte(line[i])
			}
			continue
		}
		b.WriteByte('"')
		if inString && i+1 < len(line) && line[i+1] == '"' {
			b.WriteByte('"')
			i++
			continue
		}
		inString = !inString
	}
	return b.String()
}

func missingLineContinuationWhitespace(line string) bool {
	trimmed := strings.TrimRight(line, " \t")
	if !strings.HasSuffix(trimmed, "_") || len(trimmed) < 2 {
		return false
	}
	if endsWithIdentifierUnderscore(trimmed) {
		return false
	}
	return trimmed[len(trimmed)-2] != ' ' && trimmed[len(trimmed)-2] != '\t'
}

func hasValidLineContinuation(line string) bool {
	trimmed := strings.TrimRight(line, " \t")
	if !strings.HasSuffix(trimmed, "_") || len(trimmed) < 2 {
		return false
	}
	return trimmed[len(trimmed)-2] == ' ' || trimmed[len(trimmed)-2] == '\t'
}

func removeLineContinuationMarker(line string) string {
	trimmed := strings.TrimRight(line, " \t")
	if !strings.HasSuffix(trimmed, "_") {
		return line
	}
	return trimmed[:len(trimmed)-1]
}

func splitStatements(line string) []string {
	statements := make([]string, 0, 1)
	start := 0
	inString := false
	for i := 0; i < len(line); i++ {
		switch line[i] {
		case '"':
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case ':':
			if inString {
				continue
			}
			statement := strings.TrimSpace(line[start:i])
			if statement != "" {
				statements = append(statements, statement)
			}
			start = i + 1
		}
	}
	statement := strings.TrimSpace(line[start:])
	if statement != "" {
		statements = append(statements, statement)
	}
	return statements
}

func endsWithIdentifierUnderscore(line string) bool {
	end := len(line) - 1
	if end < 0 || line[end] != '_' {
		return false
	}
	start := end
	for start >= 0 && isIdentifierChar(line[start]) {
		start--
	}
	return start < end-1
}

func isIdentifierChar(b byte) bool {
	return b == '_' || (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') || (b >= '0' && b <= '9')
}
