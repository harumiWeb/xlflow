package suppression

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

type Family string

const (
	FamilyLint    Family = "lint"
	FamilyAnalyze Family = "analyze"
)

type Directive struct {
	Code       string
	File       string
	Line       int
	TargetLine int
	Mode       string
}

type Diagnostic struct {
	Code string
	File string
	Line int
}

func DirectivesForFiles(root string, paths []string) ([]Directive, []map[string]any, error) {
	var directives []Directive
	var warnings []map[string]any
	for _, path := range paths {
		fileDirectives, fileWarnings, err := directivesForFile(root, path)
		if err != nil {
			return nil, nil, err
		}
		directives = append(directives, fileDirectives...)
		warnings = append(warnings, fileWarnings...)
	}
	return directives, warnings, nil
}

func DirectivesForSource(root, path, source string) ([]Directive, []map[string]any) {
	file, err := filepath.Rel(root, path)
	if err != nil {
		file = path
	}
	if !filepath.IsAbs(path) {
		file = path
	}
	return directivesForSource(filepath.ToSlash(file), source)
}

func Apply(diagnostics []Diagnostic, directives []Directive, family Family) ([]bool, []map[string]any) {
	suppressed := make([]bool, len(diagnostics))
	used := make([]bool, len(directives))
	byTarget := map[string][]int{}
	for i, directive := range directives {
		if !appliesToFamily(directive.Code, family) {
			continue
		}
		key := suppressionKey(directive.File, directive.TargetLine, directive.Code)
		byTarget[key] = append(byTarget[key], i)
	}
	for i, diagnostic := range diagnostics {
		if diagnostic.Line <= 0 {
			continue
		}
		for _, directiveIndex := range byTarget[suppressionKey(diagnostic.File, diagnostic.Line, diagnostic.Code)] {
			suppressed[i] = true
			used[directiveIndex] = true
			break
		}
	}
	var warnings []map[string]any
	for i, directive := range directives {
		if !appliesToFamily(directive.Code, family) || used[i] {
			continue
		}
		warnings = append(warnings, map[string]any{
			"code":        "unused_inline_suppression",
			"message":     fmt.Sprintf("Inline suppression for %s at %s:%d did not suppress any %s diagnostic.", directive.Code, directive.File, directive.Line, family),
			"rule":        directive.Code,
			"file":        directive.File,
			"line":        directive.Line,
			"target_line": directive.TargetLine,
		})
	}
	return suppressed, warnings
}

func directivesForFile(root, path string) ([]Directive, []map[string]any, error) {
	body, err := os.ReadFile(path)
	if err != nil {
		return nil, nil, err
	}
	file, err := filepath.Rel(root, path)
	if err != nil {
		file = path
	}
	file = filepath.ToSlash(file)
	directives, warnings := directivesForSource(file, string(body))
	return directives, warnings, nil
}

func directivesForSource(file, source string) ([]Directive, []map[string]any) {
	lines := normalizedSourceLines(source)
	var directives []Directive
	var warnings []map[string]any
	for i, line := range lines {
		lineNo := i + 1
		comment, ok := apostropheComment(line)
		if !ok {
			continue
		}
		mode, ids, ok := parseDirective(comment)
		if !ok {
			continue
		}
		targetLine := lineNo
		if mode == "disable-next-line" {
			targetLine = lineNo + 1
		}
		for _, rawID := range ids {
			id := strings.ToUpper(strings.TrimSpace(rawID))
			if id == "" {
				continue
			}
			if !config.KnownDiagnosticID(id) {
				warnings = append(warnings, map[string]any{
					"code":    "unknown_inline_suppression_rule",
					"message": fmt.Sprintf("Unknown inline suppression diagnostic ID %s at %s:%d.", id, file, lineNo),
					"rule":    id,
					"file":    file,
					"line":    lineNo,
				})
				continue
			}
			if !config.InlineSuppressibleDiagnosticID(id) {
				warnings = append(warnings, map[string]any{
					"code":    "unsupported_inline_suppression_rule",
					"message": fmt.Sprintf("Diagnostic ID %s cannot be suppressed inline because it blocks source preflight.", id),
					"rule":    id,
					"file":    file,
					"line":    lineNo,
				})
				continue
			}
			directives = append(directives, Directive{
				Code:       id,
				File:       file,
				Line:       lineNo,
				TargetLine: targetLine,
				Mode:       mode,
			})
		}
	}
	return directives, warnings
}

func parseDirective(comment string) (string, []string, bool) {
	fields := strings.Fields(comment)
	for i, field := range fields {
		switch {
		case strings.EqualFold(field, "xlflow:disable-next-line"):
			return "disable-next-line", fields[i+1:], len(fields) > i+1
		case strings.EqualFold(field, "xlflow:disable-line"):
			return "disable-line", fields[i+1:], len(fields) > i+1
		}
	}
	return "", nil, false
}

func apostropheComment(line string) (string, bool) {
	inString := false
	for i := 0; i < len(line); i++ {
		switch line[i] {
		case '"':
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '\'':
			if !inString {
				return line[i+1:], true
			}
		}
	}
	return "", false
}

func normalizedSourceLines(source string) []string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	return strings.Split(source, "\n")
}

func suppressionKey(file string, line int, code string) string {
	return filepath.ToSlash(file) + ":" + fmt.Sprint(line) + ":" + strings.ToUpper(strings.TrimSpace(code))
}

func appliesToFamily(id string, family Family) bool {
	switch family {
	case FamilyLint:
		return config.LintDiagnosticID(id)
	case FamilyAnalyze:
		return config.AnalyzeDiagnosticID(id)
	default:
		return false
	}
}
