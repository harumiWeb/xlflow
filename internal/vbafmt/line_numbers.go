package vbafmt

import (
	"fmt"
	"regexp"
	"strings"
)

type LineNumberMode string

const (
	LineNumberModePreserve LineNumberMode = "preserve"
	LineNumberModeAdd      LineNumberMode = "add"
	LineNumberModeRemove   LineNumberMode = "remove"
	LineNumberModeRenumber LineNumberMode = "renumber"
)

type FormatConfig struct {
	LineNumbers LineNumberMode
}

type LineNumberWarning struct {
	Path    string `json:"path,omitempty"`
	Line    int    `json:"line,omitempty"`
	Message string `json:"message"`
}

type LineNumberSummary struct {
	Mode            LineNumberMode      `json:"mode"`
	FilesChanged    int                 `json:"files_changed,omitempty"`
	LinesAdded      int                 `json:"lines_added,omitempty"`
	LinesRemoved    int                 `json:"lines_removed,omitempty"`
	LinesRenumbered int                 `json:"lines_renumbered,omitempty"`
	Warnings        []LineNumberWarning `json:"warnings,omitempty"`
}

type lineNumberFileResult struct {
	Changed         bool
	LinesAdded      int
	LinesRemoved    int
	LinesRenumbered int
	Warnings        []LineNumberWarning
}

type formattedLine struct {
	Text          string
	Content       string
	HadLineNumber bool
	LineNumber    int
	Eligible      bool
	InputLine     int
}

type lineNumberDirective struct {
	Has     bool
	Number  int
	Content string
}

type numericLabelReference struct {
	Line   int
	Target int
}

var (
	lineNumberPrefixRe   = regexp.MustCompile(`^\s*(\d+)([ \t]+)(.+)$`)
	numericLabelTargetRe = regexp.MustCompile(`(?i)\b(?:GO\s*TO|GOSUB|RESUME)\s+(\d+)\b`)
)

func normalizeLineNumberMode(mode LineNumberMode) LineNumberMode {
	if mode == "" {
		return LineNumberModePreserve
	}
	return mode
}

func parseLineNumberDirective(line string) lineNumberDirective {
	match := lineNumberPrefixRe.FindStringSubmatch(line)
	if len(match) != 4 {
		return lineNumberDirective{Content: line}
	}
	return lineNumberDirective{
		Has:     true,
		Number:  mustParsePositiveInt(match[1]),
		Content: match[3],
	}
}

func collectNumericLabelReferences(lines []string) []numericLabelReference {
	var refs []numericLabelReference
	for i, line := range lines {
		directive := parseLineNumberDirective(line)
		content := directive.Content
		stripped := stripTrailingComment(content)
		matches := numericLabelTargetRe.FindAllStringSubmatch(stripped, -1)
		for _, match := range matches {
			if len(match) != 2 {
				continue
			}
			target := mustParsePositiveInt(match[1])
			if target == 0 {
				continue
			}
			refs = append(refs, numericLabelReference{
				Line:   i + 1,
				Target: target,
			})
		}
	}
	return refs
}

func normalizeFormattedLines(lines []formattedLine) []formattedLine {
	if len(lines) == 0 {
		return lines
	}

	normalized := make([]formattedLine, 0, len(lines))
	consecutiveBlanks := 0
	for i, line := range lines {
		if strings.TrimSpace(line.Text) == "" {
			consecutiveBlanks++
			if i == len(lines)-1 {
				continue
			}
			if consecutiveBlanks == 1 {
				normalized = append(normalized, formattedLine{})
			}
			continue
		}

		if consecutiveBlanks == 0 && len(normalized) > 0 {
			prev := lastNonBlankFormattedLine(normalized)
			if prev != nil && (isOptionExplicitGap(prev.Text) || isProcedureGap(prev.Text, line.Text)) {
				normalized = append(normalized, formattedLine{})
			}
		}

		consecutiveBlanks = 0
		normalized = append(normalized, line)
	}
	return normalized
}

func lastNonBlankFormattedLine(lines []formattedLine) *formattedLine {
	for i := len(lines) - 1; i >= 0; i-- {
		if strings.TrimSpace(lines[i].Text) != "" {
			return &lines[i]
		}
	}
	return nil
}

func formatTextDetailed(text string, isClass bool, cfg FormatConfig) (string, lineNumberFileResult, error) {
	lines := splitLines(text)
	if len(lines) == 0 {
		return "", lineNumberFileResult{}, nil
	}

	mode := normalizeLineNumberMode(cfg.LineNumbers)
	refs := collectNumericLabelReferences(lines)

	ind := &indenter{
		level: 0,
		fileCtx: &fileContext{
			lines:   lines,
			isClass: isClass,
		},
	}

	formatted := make([]formattedLine, 0, len(lines))
	headerEnded := false
	inBeginBlock := false
	inProcedure := false

	for i, line := range lines {
		if !headerEnded && isClass {
			trimmedForHeader := strings.TrimLeft(line, " \t")
			upperForHeader := strings.ToUpper(trimmedForHeader)
			if inBeginBlock {
				formatted = append(formatted, formattedLine{
					Text:      line,
					Content:   trimmedForHeader,
					InputLine: i + 1,
				})
				if upperForHeader == "END" {
					inBeginBlock = false
				}
				continue
			}
			if isClassHeaderLine(line) || isBlankLine(line) {
				formatted = append(formatted, formattedLine{
					Text:      line,
					Content:   trimmedForHeader,
					InputLine: i + 1,
				})
				if upperForHeader == "BEGIN" {
					inBeginBlock = true
				}
				continue
			}
			headerEnded = true
		}

		directive := parseLineNumberDirective(line)
		trimmed := strings.TrimRight(directive.Content, " \t")
		content := strings.TrimLeft(trimmed, " \t")
		isEmpty := isBlankLine(trimmed)
		isCommentLine := isVBACommentLine(trimmed)

		if isEmpty || isCommentLine {
			indent := strings.Repeat(" ", ind.level*indentWidth)
			outLine := indent
			if !isEmpty {
				outLine = indent + content
			}
			formatted = append(formatted, formattedLine{
				Text:          outLine,
				Content:       content,
				HadLineNumber: directive.Has,
				LineNumber:    directive.Number,
				InputLine:     i + 1,
			})
			continue
		}

		if isLabelLine(content) {
			formatted = append(formatted, formattedLine{
				Text:          content,
				Content:       content,
				HadLineNumber: directive.Has,
				LineNumber:    directive.Number,
				InputLine:     i + 1,
			})
			continue
		}

		keyword, isStructural := classifyLine(content)
		if isStructural && isDedentKeyword(keyword) {
			ind.level--
			if ind.level < 0 {
				ind.level = 0
			}
		}

		indent := strings.Repeat(" ", ind.level*indentWidth)
		outLine := indent + content
		formatted = append(formatted, formattedLine{
			Text:          outLine,
			Content:       content,
			HadLineNumber: directive.Has,
			LineNumber:    directive.Number,
			Eligible:      inProcedure && isLineNumberEligibleContent(content),
			InputLine:     i + 1,
		})

		if isProcedureStartKeyword(keyword) {
			inProcedure = true
		}
		if isStructural && isIndentKeyword(keyword) {
			ind.level++
		}
		if isProcedureEndKeyword(keyword) {
			inProcedure = false
		}
	}

	formatted = normalizeFormattedLines(formatted)
	rendered, lineResult := applyLineNumberMode(formatted, mode, refs)
	if rendered == "\n" {
		return "", lineResult, nil
	}
	if !strings.HasSuffix(rendered, "\n") {
		rendered += "\n"
	}
	return rendered, lineResult, nil
}

func applyLineNumberMode(lines []formattedLine, mode LineNumberMode, refs []numericLabelReference) (string, lineNumberFileResult) {
	mode = normalizeLineNumberMode(mode)
	hasAmbiguousRefs := len(refs) > 0
	result := lineNumberFileResult{}

	eligibleCount := 0
	numberedEligibleCount := 0
	for _, line := range lines {
		if !line.Eligible {
			continue
		}
		eligibleCount++
		if line.HadLineNumber {
			numberedEligibleCount++
		}
	}

	if hasAmbiguousRefs && mode != LineNumberModePreserve {
		ref := refs[0]
		result.Warnings = append(result.Warnings, LineNumberWarning{
			Line:    ref.Line,
			Message: fmt.Sprintf("Skipped line-number %s because numeric label reference %d was found; GoTo/Gosub/Resume targets are not rewritten automatically.", mode, ref.Target),
		})
		return renderFormattedLines(lines, nil), result
	}

	switch mode {
	case LineNumberModeAdd:
		if numberedEligibleCount > 0 && numberedEligibleCount < eligibleCount {
			result.Warnings = append(result.Warnings, LineNumberWarning{
				Message: "Skipped line-number add because the file mixes numbered and unnumbered executable statements; use renumber to normalize first.",
			})
			return renderFormattedLines(lines, nil), result
		}
		if eligibleCount == 0 || numberedEligibleCount == eligibleCount {
			return renderFormattedLines(lines, nil), result
		}
		assignments := make(map[int]int)
		next := 10
		for i, line := range lines {
			if !line.Eligible {
				continue
			}
			assignments[i] = next
			next += 10
			result.LinesAdded++
		}
		result.Changed = result.LinesAdded > 0
		return renderFormattedLines(lines, assignments), result
	case LineNumberModeRemove:
		for _, line := range lines {
			if line.HadLineNumber {
				result.LinesRemoved++
			}
		}
		result.Changed = result.LinesRemoved > 0
		return renderFormattedLinesWithoutLineNumbers(lines), result
	case LineNumberModeRenumber:
		if eligibleCount == 0 {
			return renderFormattedLines(lines, nil), result
		}
		assignments := make(map[int]int)
		next := 10
		for i, line := range lines {
			if !line.Eligible {
				continue
			}
			assignments[i] = next
			if !line.HadLineNumber || line.LineNumber != next {
				result.LinesRenumbered++
			}
			next += 10
		}
		result.Changed = result.LinesRenumbered > 0
		return renderFormattedLines(lines, assignments), result
	default:
		return renderFormattedLines(lines, nil), result
	}
}

func renderFormattedLines(lines []formattedLine, assignments map[int]int) string {
	var b strings.Builder
	for i, line := range lines {
		switch {
		case line.Text == "":
			b.WriteByte('\n')
		case assignments != nil:
			if n, ok := assignments[i]; ok {
				b.WriteString(formatLineNumberedLine(n, line.Text))
				b.WriteByte('\n')
				continue
			}
			if line.HadLineNumber && !line.Eligible {
				b.WriteString(formatLineNumberedLine(line.LineNumber, line.Text))
				b.WriteByte('\n')
				continue
			}
			b.WriteString(line.Text)
			b.WriteByte('\n')
		default:
			if line.HadLineNumber {
				b.WriteString(formatLineNumberedLine(line.LineNumber, line.Text))
				b.WriteByte('\n')
				continue
			}
			b.WriteString(line.Text)
			b.WriteByte('\n')
		}
	}
	return b.String()
}

func renderFormattedLinesWithoutLineNumbers(lines []formattedLine) string {
	var b strings.Builder
	for _, line := range lines {
		b.WriteString(line.Text)
		b.WriteByte('\n')
	}
	return b.String()
}

func formatLineNumberedLine(number int, text string) string {
	if text == "" {
		return ""
	}
	return fmt.Sprintf("%d  %s", number, text)
}

func isLineNumberEligibleContent(content string) bool {
	trimmed := strings.TrimSpace(content)
	if trimmed == "" || isVBACommentLine(trimmed) {
		return false
	}
	if strings.HasPrefix(trimmed, "#") {
		return false
	}
	if isLabelLine(trimmed) {
		return false
	}
	upper := strings.ToUpper(trimmed)
	if strings.HasPrefix(upper, "OPTION ") {
		return false
	}
	if strings.HasPrefix(upper, "ATTRIBUTE ") {
		return false
	}
	if keyword, ok := classifyLine(trimmed); ok && (isProcedureStartKeyword(keyword) || isProcedureEndKeyword(keyword)) {
		return false
	}
	return true
}

func isLabelLine(trimmed string) bool {
	if !strings.HasSuffix(trimmed, ":") {
		return false
	}
	name := strings.TrimSuffix(trimmed, ":")
	if name == "" || strings.ContainsAny(name, " \t") {
		return false
	}
	for _, r := range name {
		if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || (r >= '0' && r <= '9') || r == '_' {
			continue
		}
		return false
	}
	return true
}

func isProcedureStartKeyword(keyword string) bool {
	switch keyword {
	case "SUB", "FUNCTION", "PROPERTY GET", "PROPERTY LET", "PROPERTY SET":
		return true
	default:
		return false
	}
}

func isProcedureEndKeyword(keyword string) bool {
	switch keyword {
	case "END SUB", "END FUNCTION", "END PROPERTY":
		return true
	default:
		return false
	}
}

func mustParsePositiveInt(text string) int {
	n := 0
	for _, r := range text {
		n = n*10 + int(r-'0')
	}
	return n
}
