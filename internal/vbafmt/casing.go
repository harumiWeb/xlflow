package vbafmt

import (
	"regexp"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

var (
	casingVarTypeRe  = regexp.MustCompile(`(?i)\b(?:Dim|Private|Public|Friend|Static)\s+([A-Za-z_][A-Za-z0-9_]*)\s+As\s+(?:New\s+)?([A-Za-z_][A-Za-z0-9_.]*)\b`)
	casingDeclNameRe = regexp.MustCompile(`(?i)\b(?:Dim|Private|Public|Friend|Static|Const)\s+([A-Za-z_][A-Za-z0-9_]*)\b`)
	casingAsNameRe   = regexp.MustCompile(`(?i)\b([A-Za-z_][A-Za-z0-9_]*)\s+As\s+(?:New\s+)?[A-Za-z_][A-Za-z0-9_.]*\b`)

	canonicalKeywords = canonicalMap([]string{
		"Option", "Explicit",
		"Private", "Public", "Friend", "Static",
		"Sub", "Function", "Property", "Get", "Let", "Set",
		"Dim", "Const", "As", "New",
		"If", "Then", "Else", "ElseIf", "End",
		"Select", "Case",
		"For", "Each", "To", "Step", "Next",
		"Do", "Loop", "While", "Until", "Wend",
		"With",
		"Exit",
		"On", "Error", "GoTo", "Resume",
		"Call",
		"And", "Or", "Xor", "Eqv", "Imp", "Is", "Like", "Mod",
		"True", "False", "Nothing", "Null", "Empty",
		"ByVal", "ByRef", "Optional",
	})

	canonicalBuiltins = canonicalMap([]string{
		"MsgBox", "InputBox", "Debug", "Print", "Application",
		"Workbook", "Worksheet", "Range",
		"Cells", "Rows", "Columns", "End", "Count", "Row", "Value",
		"Long", "String", "Boolean", "Integer", "Double", "Variant", "Object",
		"vbExclamation", "vbInformation", "vbCritical",
		"xlUp", "xlDown", "xlToLeft", "xlToRight",
	})

	casingGlobalTypes = map[string]string{
		"application":    "Application",
		"activeworkbook": "Workbook",
		"thisworkbook":   "Workbook",
		"activesheet":    "Worksheet",
		"cells":          "Range",
		"range":          "Range",
		"rows":           "Rows",
		"columns":        "Columns",
		"debug":          "Debug",
	}

	casingMemberReturnTypes = map[string]map[string]string{
		"Debug": {
			"print": "Variant",
		},
		"Application": {
			"cells": "Range", "range": "Range", "rows": "Rows", "columns": "Columns",
			"workbooks": "Workbooks", "worksheets": "Worksheets",
		},
		"Workbook": {
			"worksheets": "Worksheets", "sheets": "Worksheets",
		},
		"Worksheet": {
			"cells": "Range", "range": "Range", "rows": "Rows", "columns": "Columns",
		},
		"Range": {
			"cells": "Range", "rows": "Rows", "columns": "Columns", "end": "Range",
			"row": "Long", "value": "Variant", "count": "Long",
		},
		"Rows": {
			"count": "Long", "end": "Range",
		},
		"Columns": {
			"count": "Long", "end": "Range",
		},
		"Worksheets": {
			"item": "Worksheet", "count": "Long",
		},
		"Workbooks": {
			"item": "Workbook", "count": "Long",
		},
	}
)

func canonicalMap(names []string) map[string]string {
	out := make(map[string]string, len(names))
	for _, name := range names {
		out[strings.ToLower(name)] = name
	}
	return out
}

func formatCasing(text string, isClass bool, keywordCasing, builtinCasing bool) (string, error) {
	parser, err := vbaast.NewParser()
	if err != nil {
		return "", err
	}
	defer parser.Close()

	source := []byte(text)
	parsed := parser.Parse("<fmt>", source)
	defer parsed.Close()
	if parsed.HasError || parsed.HasMissing {
		return "", formatParseError{hasError: parsed.HasError, hasMissing: parsed.HasMissing}
	}

	lines := splitLines(text)
	lineStarts := lineStartOffsets(text, lines)
	unsafe := unsafeOperatorSpacingLines(lines, isClass)
	varTypes := casingVariableTypes(lines)
	userNames := casingUserNames(lines)

	var edits []textEdit
	for lineNo, line := range lines {
		if lineNo >= len(unsafe) || unsafe[lineNo] {
			continue
		}
		edits = append(edits, casingEditsForLine(line, lineStarts[lineNo], keywordCasing, builtinCasing, varTypes, userNames)...)
	}
	if len(edits) == 0 {
		return text, nil
	}
	sort.SliceStable(edits, func(i, j int) bool {
		return edits[i].start > edits[j].start
	})
	out := []byte(text)
	for _, edit := range edits {
		out = append(out[:edit.start], append([]byte(edit.text), out[edit.end:]...)...)
	}
	return string(out), nil
}

func casingVariableTypes(lines []string) map[string]string {
	out := map[string]string{}
	for _, line := range lines {
		code := stripStringLiterals(stripTrailingComment(parseLineNumberDirective(line).Content))
		if strings.Contains(strings.ToUpper(code), "DECLARE ") {
			continue
		}
		for _, match := range casingVarTypeRe.FindAllStringSubmatch(code, -1) {
			if len(match) != 3 {
				continue
			}
			if typ, ok := casingCanonicalType(match[2]); ok {
				out[strings.ToLower(match[1])] = typ
			}
		}
	}
	return out
}

func casingUserNames(lines []string) map[string]bool {
	out := map[string]bool{}
	inEnum := false
	for _, line := range lines {
		code := strings.TrimSpace(stripStringLiterals(stripTrailingComment(parseLineNumberDirective(line).Content)))
		if code == "" {
			continue
		}
		upper := strings.ToUpper(code)
		if strings.HasPrefix(upper, "END ENUM") {
			inEnum = false
			continue
		}
		if inEnum {
			if name := firstIdentifier(code); name != "" {
				out[strings.ToLower(name)] = true
			}
			continue
		}
		if strings.HasPrefix(upper, "ENUM ") || strings.Contains(upper, " ENUM ") {
			inEnum = true
		}
		if strings.Contains(upper, "DECLARE ") {
			continue
		}
		for _, match := range casingDeclNameRe.FindAllStringSubmatch(code, -1) {
			if len(match) == 2 {
				out[strings.ToLower(match[1])] = true
			}
		}
		for _, match := range casingAsNameRe.FindAllStringSubmatch(code, -1) {
			if len(match) == 2 {
				out[strings.ToLower(match[1])] = true
			}
		}
		fields := strings.Fields(code)
		for i := 0; i+1 < len(fields); i++ {
			word := strings.Trim(fields[i], " \t")
			next := fields[i+1]
			if idx := strings.IndexByte(next, '('); idx >= 0 {
				next = next[:idx]
			}
			next = strings.TrimSpace(next)
			switch strings.ToLower(word) {
			case "sub", "function", "get", "let", "set":
				if next != "" {
					out[strings.ToLower(next)] = true
				}
			}
		}
	}
	return out
}

func firstIdentifier(code string) string {
	for i := 0; i < len(code); i++ {
		if !isIdentifierStartByte(code[i]) {
			continue
		}
		j := i + 1
		for j < len(code) && isIdentifierPartByte(code[j]) {
			j++
		}
		return code[i:j]
	}
	return ""
}

func casingCanonicalType(raw string) (string, bool) {
	parts := strings.Split(raw, ".")
	name := parts[len(parts)-1]
	canon, ok := canonicalBuiltins[strings.ToLower(name)]
	if !ok {
		return "", false
	}
	switch canon {
	case "Application", "Workbook", "Worksheet", "Range", "Rows", "Columns", "Long", "String", "Boolean", "Integer", "Double", "Variant", "Object":
		return canon, true
	default:
		return "", false
	}
}

func casingEditsForLine(line string, lineStart int, keywordCasing, builtinCasing bool, varTypes map[string]string, userNames map[string]bool) []textEdit {
	directive := parseLineNumberDirective(line)
	prefixLen := len(line) - len(directive.Content)
	content := directive.Content
	codeEnd := len(content)
	if comment := trailingCommentOffset([]byte(content)); comment >= 0 {
		codeEnd = comment
	}
	code := content[:codeEnd]
	var edits []textEdit
	casingScanSegment(code, 0, len(code), lineStart+prefixLen, keywordCasing, builtinCasing, varTypes, userNames, &edits)
	return edits
}

func casingScanSegment(code string, start, end, baseOffset int, keywordCasing, builtinCasing bool, varTypes map[string]string, userNames map[string]bool, edits *[]textEdit) {
	currentType := ""
	for i := start; i < end; {
		ch := code[i]
		if ch == '"' {
			i = skipStringLiteral(code, i, end)
			currentType = ""
			continue
		}
		if ch == '(' {
			close := matchingParen(code, i, end)
			if close > i {
				casingScanSegment(code, i+1, close, baseOffset, keywordCasing, builtinCasing, varTypes, userNames, edits)
				i = close + 1
				continue
			}
			currentType = ""
			i++
			continue
		}
		if ch == '.' {
			j := i + 1
			for j < end && isHorizontalWhitespace(code[j]) {
				j++
			}
			if j < end && isIdentifierStartByte(code[j]) {
				nameStart := j
				j++
				for j < end && isIdentifierPartByte(code[j]) {
					j++
				}
				word := code[nameStart:j]
				if builtinCasing && currentType != "" {
					if nextType, ok := casingMemberType(currentType, word); ok {
						appendCasingEdit(edits, baseOffset+nameStart, baseOffset+j, word, canonicalBuiltins[strings.ToLower(word)])
						currentType = nextType
					} else {
						currentType = ""
					}
				} else {
					currentType = ""
				}
				i = j
				continue
			}
			currentType = ""
			i++
			continue
		}
		if isIdentifierStartByte(ch) {
			j := i + 1
			for j < end && isIdentifierPartByte(code[j]) {
				j++
			}
			word := code[i:j]
			lower := strings.ToLower(word)
			prevDot := previousSignificantByte(code, i) == '.'
			if !prevDot {
				if keywordCasing {
					appendCasingEdit(edits, baseOffset+i, baseOffset+j, word, canonicalKeywords[lower])
				}
				if builtinCasing && !userNames[lower] {
					if casingStandaloneBuiltinIsSafe(code, i, j, lower) {
						appendCasingEdit(edits, baseOffset+i, baseOffset+j, word, canonicalBuiltins[lower])
					}
				}
				currentType = casingExpressionType(lower, varTypes)
			}
			i = j
			continue
		}
		if !isHorizontalWhitespace(ch) {
			currentType = ""
		}
		i++
	}
}

func appendCasingEdit(edits *[]textEdit, start, end int, original, canonical string) {
	if canonical == "" || original == canonical {
		return
	}
	*edits = append(*edits, textEdit{start: start, end: end, text: canonical})
}

func casingExpressionType(lower string, varTypes map[string]string) string {
	if typ, ok := varTypes[lower]; ok {
		return typ
	}
	if typ, ok := casingGlobalTypes[lower]; ok {
		return typ
	}
	if typ, ok := casingCanonicalType(lower); ok {
		return typ
	}
	return ""
}

func casingMemberType(receiverType, member string) (string, bool) {
	members := casingMemberReturnTypes[receiverType]
	if members == nil {
		return "", false
	}
	next, ok := members[strings.ToLower(member)]
	return next, ok
}

func casingStandaloneBuiltinIsSafe(code string, start, end int, lower string) bool {
	canon := canonicalBuiltins[lower]
	if canon == "" {
		return false
	}
	if strings.HasPrefix(lower, "vb") || strings.HasPrefix(lower, "xl") {
		return true
	}
	if _, ok := casingGlobalTypes[lower]; ok {
		return true
	}
	if typeKeywordBefore(code, start) {
		_, ok := casingCanonicalType(canon)
		return ok
	}
	switch canon {
	case "MsgBox", "InputBox", "Debug":
		return statementHeadBefore(code, start)
	default:
		return false
	}
}

func typeKeywordBefore(code string, start int) bool {
	prefix := strings.TrimSpace(code[:start])
	fields := strings.Fields(prefix)
	if len(fields) == 0 {
		return false
	}
	switch strings.ToLower(fields[len(fields)-1]) {
	case "as", "new":
		return true
	default:
		return false
	}
}

func statementHeadBefore(code string, start int) bool {
	prefix := strings.TrimSpace(code[:start])
	if prefix == "" {
		return true
	}
	idx := strings.LastIndex(prefix, ":")
	if idx >= 0 && strings.TrimSpace(prefix[idx+1:]) == "" {
		return true
	}
	fields := strings.Fields(prefix)
	if len(fields) > 0 {
		switch strings.ToLower(fields[len(fields)-1]) {
		case "then", "else", "call":
			return true
		}
	}
	return false
}

func skipStringLiteral(code string, start, end int) int {
	i := start + 1
	for i < end {
		if code[i] == '"' {
			if i+1 < end && code[i+1] == '"' {
				i += 2
				continue
			}
			return i + 1
		}
		i++
	}
	return end
}

func matchingParen(code string, open, end int) int {
	depth := 0
	for i := open; i < end; i++ {
		switch code[i] {
		case '"':
			i = skipStringLiteral(code, i, end) - 1
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				return i
			}
		}
	}
	return -1
}
