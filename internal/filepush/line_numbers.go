package filepush

import (
	"regexp"
	"strconv"
	"strings"
)

// Erl line-number instrumentation. This is a port of the .NET
// ErlLineNumberTransformer.TryAdd used by the Excel bridge: every input line
// is preserved verbatim and eligible procedure statements receive their
// physical source line number as a fixed-width, space-padded prefix followed
// by one space. The matching removal port lives in internal/filepull; the two
// share the same classification patterns as the bridge.
var (
	numberPrefixPattern   = regexp.MustCompile(`^\s*(\d+)\s+(.+)$`)
	numericTargetPattern  = regexp.MustCompile(`(?i)\b(?:GO\s*TO|GOSUB|RESUME)\s+(\d+)\b`)
	procedureStartPattern = regexp.MustCompile(`(?i)^\s*(?:(?:PUBLIC|PRIVATE|FRIEND|STATIC)\s+)*(?:SUB|FUNCTION)\b|^\s*(?:(?:PUBLIC|PRIVATE|FRIEND|STATIC)\s+)*PROPERTY\s+(?:GET|LET|SET)\b`)
	procedureEndPattern   = regexp.MustCompile(`(?i)^\s*END\s+(?:SUB|FUNCTION|PROPERTY)\b`)
	declarationPattern    = regexp.MustCompile(`(?i)^\s*(?:DIM|STATIC|CONST|PRIVATE|PUBLIC|FRIEND|GLOBAL|DECLARE|TYPE|ENUM|EVENT)\b`)
	labelPattern          = regexp.MustCompile(`(?i)^[A-Z_][A-Z0-9_]*:\s*$`)
	structuralPattern     = regexp.MustCompile(`(?i)^(?:SELECT\s+CASE|CASE(?:\s+ELSE)?|END\s+SELECT|ELSE|ELSEIF|END\s+(?:IF|WITH)|FOR(?:\s+EACH)?|NEXT|DO(?:\s+(?:WHILE|UNTIL))?|LOOP(?:\s+(?:WHILE|UNTIL))?|WHILE|WEND|WITH)\b`)
	blockIfPattern        = regexp.MustCompile(`(?i)^IF\b.*\bTHEN\s*$`)
)

type lineNumberIssue struct {
	Line    int
	Message string
}

// tryAddLineNumbers ports ErlLineNumberTransformer.TryAdd. It returns the
// transformed text, or the issue explaining why instrumentation would be
// unsafe (existing numeric labels or numeric branch targets).
func tryAddLineNumbers(text string) (string, *lineNumberIssue) {
	lines, newline := splitSourceLines(text)
	if issue := validateNoExistingNumbers(lines); issue != nil {
		return text, issue
	}

	width := len(strconv.Itoa(len(lines)))
	if width < 1 {
		width = 1
	}
	inProcedure := false
	continuationTail := false
	for index, content := range lines {
		if procedureStartPattern.MatchString(content) {
			inProcedure = true
			continuationTail = hasContinuation(content)
			continue
		}
		if procedureEndPattern.MatchString(content) {
			inProcedure = false
			continuationTail = false
			continue
		}
		if inProcedure && !continuationTail && lineNumberEligible(content) {
			number := strconv.Itoa(index + 1)
			if len(number) < width {
				number = strings.Repeat(" ", width-len(number)) + number
			}
			lines[index] = number + " " + content
		}
		continuationTail = hasContinuation(content)
	}
	return strings.Join(lines, newline), nil
}

func validateNoExistingNumbers(lines []string) *lineNumberIssue {
	for index, line := range lines {
		for _, match := range numericTargetPattern.FindAllStringSubmatch(stripStringsAndComment(line), -1) {
			if target, err := strconv.Atoi(match[1]); err == nil && target > 0 {
				return &lineNumberIssue{Line: index + 1, Message: "numeric GoTo, GoSub, or Resume target would be unsafe to transform"}
			}
		}
	}
	for index, line := range lines {
		if numberPrefixPattern.MatchString(line) {
			return &lineNumberIssue{Line: index + 1, Message: "existing numeric line label would be changed by Erl instrumentation"}
		}
	}
	return nil
}

func lineNumberEligible(line string) bool {
	trimmed := strings.TrimSpace(line)
	if trimmed == "" || strings.HasPrefix(trimmed, "'") || strings.HasPrefix(trimmed, "#") || declarationPattern.MatchString(trimmed) {
		return false
	}
	if labelPattern.MatchString(trimmed) || structuralPattern.MatchString(trimmed) {
		return false
	}
	return !blockIfPattern.MatchString(stripStringsAndComment(trimmed))
}

func hasContinuation(line string) bool {
	return strings.HasSuffix(strings.TrimRight(stripStringsAndComment(line), " \t"), " _")
}

// stripStringsAndComment removes string literal contents and trailing comments
// so keyword patterns never match inside quoted text.
func stripStringsAndComment(line string) string {
	var b strings.Builder
	inString := false
	for i := 0; i < len(line); i++ {
		ch := line[i]
		if ch == '"' {
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			continue
		}
		if !inString && ch == '\'' {
			break
		}
		if !inString {
			b.WriteByte(ch)
		}
	}
	return b.String()
}

// splitSourceLines ports the .NET SplitLines: the dominant newline style of
// the text wins (CRLF > LF > CR), so instrumentation never changes a file's
// line-ending convention.
func splitSourceLines(text string) (lines []string, newline string) {
	switch {
	case strings.Contains(text, "\r\n"):
		newline = "\r\n"
	case strings.Contains(text, "\n"):
		newline = "\n"
	case strings.Contains(text, "\r"):
		newline = "\r"
	default:
		newline = "\r\n"
	}
	if text == "" {
		return nil, newline
	}
	normalized := strings.ReplaceAll(text, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	return strings.Split(normalized, "\n"), newline
}
