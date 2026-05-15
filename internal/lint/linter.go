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

var (
	selectRe          = regexp.MustCompile(`(?i)\.\s*select\b`)
	activateRe        = regexp.MustCompile(`(?i)\.\s*activate\b`)
	onErrorResumeNext = regexp.MustCompile(`(?i)\bon\s+error\s+resume\s+next\b`)
	dimWithoutAs      = regexp.MustCompile(`(?i)^\s*(dim|private|public|static)\s+([^']+)$`)
	publicVarRe       = regexp.MustCompile(`(?i)^\s*public\s+\w+`)
	publicProcRe      = regexp.MustCompile(`(?i)^\s*public\s+(sub|function|property|type|enum|declare)\b`)
)

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
	scanner := bufio.NewScanner(f)
	lineNo := 0
	for scanner.Scan() {
		lineNo++
		line := scanner.Text()
		code := gui.StripComment(line)
		trimmed := strings.TrimSpace(code)
		if strings.EqualFold(trimmed, "Option Explicit") {
			hasOptionExplicit = true
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
		if l.Config.Lint.DetectImplicitVariant && looksImplicitVariant(trimmed) {
			issues = append(issues, l.issue(path, lineNo, "VB005", "warning", "Declare an explicit type with As <Type>."))
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
				Message:    boundary.Message + " " + boundary.Suggestion,
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
		if issue.Code == "VB008" || issue.Code == "VB009" {
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

func looksImplicitVariant(line string) bool {
	if line == "" || strings.Contains(strings.ToLower(line), " as ") {
		return false
	}
	matches := dimWithoutAs.FindStringSubmatch(line)
	if len(matches) == 0 {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(line))
	return !strings.HasPrefix(lower, "public sub ") &&
		!strings.HasPrefix(lower, "public function ") &&
		!strings.HasPrefix(lower, "private sub ") &&
		!strings.HasPrefix(lower, "private function ")
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

func looksPublicVariable(line string) bool {
	if !publicVarRe.MatchString(line) {
		return false
	}
	return !publicProcRe.MatchString(line)
}
