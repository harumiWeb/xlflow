package lint

import (
	"bufio"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/gui"
)

type Issue struct {
	Code       string `json:"code"`
	Severity   string `json:"severity"`
	File       string `json:"file"`
	Line       int    `json:"line"`
	Message    string `json:"message"`
	Kind       string `json:"kind,omitempty"`
	Symbol     string `json:"symbol,omitempty"`
	Suggestion string `json:"suggestion,omitempty"`
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
	selectRe          = regexp.MustCompile(`(?i)\.\s*select\b`)
	activateRe        = regexp.MustCompile(`(?i)\.\s*activate\b`)
	onErrorResumeNext = regexp.MustCompile(`(?i)\bon\s+error\s+resume\s+next\b`)
	dimWithoutAs      = regexp.MustCompile(`(?i)^\s*(dim|private|public|static)\s+([^']+)$`)
	publicVarRe       = regexp.MustCompile(`(?i)^\s*public\s+\w+`)
	publicProcRe      = regexp.MustCompile(`(?i)^\s*public\s+(sub|function|property|type|enum|declare)\b`)
	typeStartRe       = regexp.MustCompile(`(?i)^\s*(private|public)?\s*type\b`)
	typeEndRe         = regexp.MustCompile(`(?i)^\s*end\s+type\b`)
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
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}

	var issues []Issue
	hasOptionExplicit := false
	inTypeBlock := false
	procedures := make([]procedureFrame, 0)
	var logicalLine strings.Builder
	logicalStartLine := 0
	scanner := bufio.NewScanner(f)
	lineNo := 0
	for scanner.Scan() {
		lineNo++
		line := scanner.Text()
		code := gui.StripComment(line)
		detectionCode := maskStringLiterals(code)
		trimmed := strings.TrimSpace(code)
		if strings.EqualFold(trimmed, "Option Explicit") {
			hasOptionExplicit = true
		}
		if missingLineContinuationWhitespace(detectionCode) {
			issues = append(issues, l.issue(path, lineNo, "VB013", "error", "Line-continuation underscore must be preceded by whitespace."))
		}
		if l.Config.Lint.ForbidSelect && selectRe.MatchString(code) {
			issues = append(issues, l.issue(path, lineNo, "VB002", "warning", "Avoid Select. Use direct object references instead."))
		}
		if l.Config.Lint.ForbidActivate && activateRe.MatchString(code) {
			issues = append(issues, l.issue(path, lineNo, "VB003", "warning", "Avoid Activate. Use direct object references instead."))
		}
		if l.Config.Lint.ForbidOnErrorResumeNext && onErrorResumeNext.MatchString(code) {
			issues = append(issues, l.issue(path, lineNo, "VB004", "warning", "Avoid On Error Resume Next without a narrow recovery block."))
		}
		if l.Config.Lint.DetectImplicitVariant && looksImplicitVariant(trimmed, inTypeBlock) {
			issues = append(issues, l.issue(path, lineNo, "VB005", "warning", "Declare an explicit type with As <Type>."))
		}
		if typeStartRe.MatchString(trimmed) {
			inTypeBlock = true
		} else if typeEndRe.MatchString(trimmed) {
			inTypeBlock = false
		}
		if l.Config.Lint.ForbidPublicModuleFields && looksPublicVariable(trimmed) {
			issues = append(issues, l.issue(path, lineNo, "VB006", "warning", "Avoid Public module variables; pass state explicitly."))
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
	if err := scanner.Err(); err != nil {
		if closeErr := f.Close(); closeErr != nil {
			return nil, closeErr
		}
		return nil, err
	}
	if err := f.Close(); err != nil {
		return nil, err
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
	if l.Config.Lint.RequireOptionExplicit && !hasOptionExplicit {
		issues = append([]Issue{l.issue(path, 1, "VB001", "error", "Missing Option Explicit.")}, issues...)
	}
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

func PushBlockingIssues(issues []Issue) []Issue {
	blocking := make([]Issue, 0)
	for _, issue := range issues {
		if issue.Code == "VB008" || issue.Code == "VB009" || issue.Code == "VB010" || issue.Code == "VB011" || issue.Code == "VB012" || issue.Code == "VB013" {
			blocking = append(blocking, issue)
		}
	}
	return blocking
}

func (l Linter) issue(path string, line int, code, severity, message string) Issue {
	file, err := filepath.Rel(l.RootDir, path)
	if err != nil {
		file = path
	}
	return Issue{
		Code:     code,
		Severity: severity,
		File:     filepath.ToSlash(file),
		Line:     line,
		Message:  message,
	}
}

func looksImplicitVariant(line string, inTypeBlock bool) bool {
	lower := strings.ToLower(strings.TrimSpace(line))
	if lower == "" || strings.Contains(lower, " as ") {
		return false
	}
	if inTypeBlock {
		return !typeEndRe.MatchString(line) && !isConditionalCompilationDirective(lower)
	}
	matches := dimWithoutAs.FindStringSubmatch(line)
	if len(matches) == 0 {
		return false
	}
	return !strings.HasPrefix(lower, "public sub ") &&
		!strings.HasPrefix(lower, "public function ") &&
		!strings.HasPrefix(lower, "public property ") &&
		!strings.HasPrefix(lower, "public type ") &&
		!strings.HasPrefix(lower, "public enum ") &&
		!strings.HasPrefix(lower, "public declare ") &&
		!strings.HasPrefix(lower, "private sub ") &&
		!strings.HasPrefix(lower, "private function ") &&
		!strings.HasPrefix(lower, "private property ") &&
		!strings.HasPrefix(lower, "private type ") &&
		!strings.HasPrefix(lower, "private enum ") &&
		!strings.HasPrefix(lower, "private declare ") &&
		!strings.HasPrefix(lower, "friend sub ") &&
		!strings.HasPrefix(lower, "friend function ") &&
		!strings.HasPrefix(lower, "friend property ")
}

func isConditionalCompilationDirective(line string) bool {
	return strings.HasPrefix(line, "#if ") ||
		strings.HasPrefix(line, "#elseif ") ||
		line == "#else" ||
		line == "#end if"
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

func looksPublicVariable(line string) bool {
	if !publicVarRe.MatchString(line) {
		return false
	}
	return !publicProcRe.MatchString(line)
}
