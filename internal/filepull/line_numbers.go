package filepull

import (
	"errors"
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

var (
	ErrLineNumberSafety      = errors.New("file pull: VBA line number safety check failed")
	numberPrefixPattern      = regexp.MustCompile(`^\s*(\d+)\s+(.+)$`)
	exportedGeneratedPattern = regexp.MustCompile(`^\s*\d+  (.*)$`)
	numericTargetPattern     = regexp.MustCompile(`(?i)\b(?:GO\s*TO|GOSUB|RESUME)\s+(\d+)\b`)
	procedureStartPattern    = regexp.MustCompile(`(?i)^\s*(?:(?:PUBLIC|PRIVATE|FRIEND|STATIC)\s+)*(?:SUB|FUNCTION)\b|^\s*(?:(?:PUBLIC|PRIVATE|FRIEND|STATIC)\s+)*PROPERTY\s+(?:GET|LET|SET)\b`)
	procedureEndPattern      = regexp.MustCompile(`(?i)^\s*END\s+(?:SUB|FUNCTION|PROPERTY)\b`)
	declarationPattern       = regexp.MustCompile(`(?i)^\s*(?:DIM|STATIC|CONST|PRIVATE|PUBLIC|FRIEND|GLOBAL|DECLARE|TYPE|ENUM|EVENT)\b`)
	labelPattern             = regexp.MustCompile(`(?i)^[A-Z_][A-Z0-9_]*:\s*$`)
	structuralPattern        = regexp.MustCompile(`(?i)^(?:SELECT\s+CASE|CASE(?:\s+ELSE)?|END\s+SELECT|ELSE|ELSEIF|END\s+(?:IF|WITH)|FOR(?:\s+EACH)?|NEXT|DO(?:\s+(?:WHILE|UNTIL))?|LOOP(?:\s+(?:WHILE|UNTIL))?|WHILE|WEND|WITH)\b`)
	blockIfPattern           = regexp.MustCompile(`(?i)^IF\b.*\bTHEN\s*$`)
)

func removeGeneratedLineNumbers(source string) (string, error) {
	lines := strings.Split(source, "\n")
	for i, line := range lines {
		for _, match := range numericTargetPattern.FindAllStringSubmatch(stripStringsAndComment(line), -1) {
			target, _ := strconv.Atoi(match[1])
			if target > 0 {
				return "", lineNumberSafetyError(i+1, "numeric GoTo, GoSub, or Resume target would be unsafe to transform")
			}
		}
	}
	hasNumbers := false
	for _, line := range lines {
		if numberPrefixPattern.MatchString(line) {
			hasNumbers = true
			break
		}
	}
	if !hasNumbers {
		return source, nil
	}

	inProcedure := false
	continuationTail := false
	for i, original := range lines {
		directive := numberPrefixPattern.FindStringSubmatch(original)
		content := original
		if len(directive) == 3 {
			content = directive[2]
		}
		if procedureStartPattern.MatchString(content) {
			if len(directive) == 3 {
				return "", lineNumberSafetyError(i+1, "numeric line number on a procedure declaration is not xlflow-generated")
			}
			inProcedure = true
			continuationTail = hasContinuation(content)
			continue
		}
		if procedureEndPattern.MatchString(content) {
			if len(directive) == 3 {
				return "", lineNumberSafetyError(i+1, "numeric line number on a procedure boundary is not xlflow-generated")
			}
			inProcedure = false
			continuationTail = false
			continue
		}

		eligible := inProcedure && !continuationTail && lineNumberEligible(content)
		if len(directive) == 3 {
			number, err := strconv.Atoi(directive[1])
			generated := exportedGeneratedPattern.FindStringSubmatch(original)
			if !eligible || err != nil || number != i+1 || len(generated) != 2 {
				return "", lineNumberSafetyError(i+1, "numeric line label is not an xlflow-generated physical line number")
			}
			lines[i] = generated[1]
		} else if eligible {
			return "", lineNumberSafetyError(i+1, "missing xlflow-generated line number in an executable procedure statement")
		}
		continuationTail = hasContinuation(content)
	}
	return strings.Join(lines, "\n"), nil
}

func lineNumberSafetyError(line int, message string) error {
	return fmt.Errorf("%w at line %d: %s", ErrLineNumberSafety, line, message)
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
