package analyze

import (
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// worksheetRoot is intentionally conservative: only identities derived from a
// worksheet codename, a ThisWorkbook selector, or a prior Set assignment are
// explicit. Everything else remains implicit or unknown and cannot cause
// VBA216.
type worksheetRoot struct {
	kind     worksheetRootKind
	identity string
	label    string
}

type worksheetRootKind uint8

const (
	worksheetRootUnknown worksheetRootKind = iota
	worksheetRootImplicit
	worksheetRootExplicit
)

type worksheetRootTracker struct {
	codenames map[string]string
	variables map[string]worksheetRoot
	withStack []worksheetRoot
}

func newWorksheetRootTracker(codenames map[string]string) *worksheetRootTracker {
	copyCodenames := make(map[string]string, len(codenames))
	for key, value := range codenames {
		copyCodenames[key] = value
	}
	return &worksheetRootTracker{codenames: copyCodenames, variables: map[string]worksheetRoot{}}
}

// realtimeWorksheetCodenames reads only workbook component names. Names are
// stable even while an editor buffer is unsaved, and the active document is
// included explicitly so a newly opened workbook module is recognized too.
func realtimeWorksheetCodenames(rootDir, workbookSource, activePath string) map[string]string {
	codenames := map[string]string{}
	workbookRoot := filepath.Clean(filepath.Join(rootDir, workbookSource))
	_ = filepath.WalkDir(workbookRoot, func(path string, entry os.DirEntry, err error) error {
		if err != nil || entry == nil || entry.IsDir() {
			return nil
		}
		addRealtimeWorksheetCodename(codenames, path)
		return nil
	})
	if rel, err := filepath.Rel(workbookRoot, activePath); err == nil && rel != "" && !strings.HasPrefix(rel, "..") {
		addRealtimeWorksheetCodename(codenames, activePath)
	}
	return codenames
}

func addRealtimeWorksheetCodename(codenames map[string]string, path string) {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".bas", ".cls":
	default:
		return
	}
	name := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	if name == "" || strings.EqualFold(name, "ThisWorkbook") {
		return
	}
	codenames[strings.ToLower(name)] = name
}

func (t *worksheetRootTracker) pushWith(expression string) {
	t.withStack = append(t.withStack, t.resolve(expression))
}

func (t *worksheetRootTracker) popWith() {
	if len(t.withStack) > 0 {
		t.withStack = t.withStack[:len(t.withStack)-1]
	}
}

func (t *worksheetRootTracker) observeSetAssignment(statement string) {
	match := setWorksheetAssignmentRe.FindStringSubmatch(statement)
	if len(match) == 0 {
		return
	}
	name := strings.ToLower(match[1])
	root := t.resolve(match[2])
	if root.kind == worksheetRootExplicit {
		t.variables[name] = root
		return
	}
	delete(t.variables, name)
}

func (t *worksheetRootTracker) resolve(expression string) worksheetRoot {
	expression = trimOuterVBAParentheses(expression)
	if expression == "" {
		return worksheetRoot{kind: worksheetRootUnknown}
	}
	if strings.HasPrefix(expression, ".") {
		if len(t.withStack) == 0 {
			return worksheetRoot{kind: worksheetRootUnknown}
		}
		return t.withStack[len(t.withStack)-1]
	}
	lower := strings.ToLower(strings.ReplaceAll(expression, " ", ""))
	if strings.HasPrefix(lower, "activesheet") || strings.HasPrefix(lower, "activeworkbook") || strings.HasPrefix(lower, "selection") {
		return worksheetRoot{kind: worksheetRootImplicit, label: "the active worksheet"}
	}
	if strings.HasPrefix(lower, "thisworkbook.worksheets(") || strings.HasPrefix(lower, "thisworkbook.sheets(") {
		if end := balancedCallEnd(expression, strings.Index(expression, "(")); end > 0 {
			label := strings.TrimSpace(expression[:end+1])
			return worksheetRoot{kind: worksheetRootExplicit, identity: strings.ToLower(strings.ReplaceAll(label, " ", "")), label: label}
		}
	}
	name := firstVBAIdentifier(expression)
	if name == "" {
		return worksheetRoot{kind: worksheetRootUnknown}
	}
	if root, ok := t.variables[strings.ToLower(name)]; ok {
		return root
	}
	if codename, ok := t.codenames[strings.ToLower(name)]; ok {
		return worksheetRoot{kind: worksheetRootExplicit, identity: "codename:" + strings.ToLower(codename), label: codename}
	}
	if sheetCodenameRe.MatchString(name) {
		return worksheetRoot{kind: worksheetRootExplicit, identity: "codename:" + strings.ToLower(name), label: name}
	}
	return worksheetRoot{kind: worksheetRootUnknown}
}

func trimOuterVBAParentheses(expression string) string {
	expression = strings.TrimSpace(expression)
	for len(expression) > 1 && expression[0] == '(' {
		if end := balancedCallEnd(expression, 0); end == len(expression)-1 {
			expression = strings.TrimSpace(expression[1:end])
			continue
		}
		break
	}
	return expression
}

type worksheetMemberAccess struct {
	member string
	root   worksheetRoot
	start  int
	end    int
	args   string
}

var (
	setWorksheetAssignmentRe = regexp.MustCompile(`(?i)^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$`)
	sheetCodenameRe          = regexp.MustCompile(`(?i)^sheet[0-9]+$`)
	lastRowTargetRe          = regexp.MustCompile(`(?i)^\s*(?:let\s+)?(last_?row|end_?row)\s*=`)
	endDownRowRe             = regexp.MustCompile(`(?i)\.\s*end\s*\(\s*xlDown\s*\)\s*\.\s*row\b`)
	endRowRe                 = regexp.MustCompile(`(?i)\.\s*end\s*\([^)]*\)\s*\.\s*row\b`)
	usedRangeCountRe         = regexp.MustCompile(`(?i)\.\s*usedrange\s*\.\s*rows\s*\.\s*count\b`)
	usedRangeAdjustedRe      = regexp.MustCompile(`(?i)\.\s*usedrange\s*\.\s*row\s*\+\s*.*\.\s*usedrange\s*\.\s*rows\s*\.\s*count\s*-\s*1\b`)
	currentRegionCountRe     = regexp.MustCompile(`(?i)\.\s*currentregion\s*\.\s*rows\s*\.\s*count\b`)
)

func (a Analyzer) worksheetRootFindings(file parsedFile, proc sourceProcedure, lineNo int, statement string, roots *worksheetRootTracker) []Finding {
	accesses := worksheetMemberAccesses(statement, roots)
	code := maskVBAStringLiterals(statement)
	lastRowContext := lastRowTargetRe.MatchString(code) || endRowRe.MatchString(code)
	var findings []Finding
	if a.Config.Analyze.DetectWorksheetRootMismatch {
		for _, outer := range accesses {
			if outer.member != "cells" && outer.member != "range" || outer.root.kind != worksheetRootExplicit || outer.args == "" {
				continue
			}
			if nested, ok := distinctExplicitRootInArguments(accesses, outer); ok {
				findings = append(findings, a.simpleFinding(
					file, proc, lineNo, "VBA216", "error",
					outer.root.label+" "+displayWorksheetMember(outer.member)+" expression uses an argument rooted in "+nested.label+".",
					"Excel evaluates Cells and Range arguments against their own worksheet roots; mixing explicit worksheets can select the wrong cells or fail at runtime.",
					fullyQualifiedRootSuggestion(outer.root, outer.member),
				))
			}
		}
	}
	if !a.Config.Analyze.DetectUnstableLastRowPatterns || !lastRowContext {
		return findings
	}
	if hasImplicitWorksheetAccess(accesses) {
		findings = append(findings, a.simpleFinding(
			file, proc, lineNo, "VBA217", "warning",
			"Last-row calculation depends on an implicit worksheet root.",
			"Bare Range, Cells, or Rows members bind through the active worksheet, so the result can change with Excel's active context.",
			"Capture a Worksheet and use `ws.Cells(ws.Rows.Count, 1).End(xlUp).Row`.",
		))
	}
	if endDownRowRe.MatchString(code) {
		findings = append(findings, a.simpleFinding(
			file, proc, lineNo, "VBA217", "warning",
			"End(xlDown) is used as a last-row strategy.",
			"xlDown stops at the first blank boundary and can jump to the worksheet limit when the starting cell is blank.",
			"Use `ws.Cells(ws.Rows.Count, column).End(xlUp).Row` with a column that has reliable values.",
		))
	}
	if usedRangeCountRe.MatchString(code) && !usedRangeAdjustedRe.MatchString(code) {
		findings = append(findings, a.simpleFinding(
			file, proc, lineNo, "VBA217", "warning",
			"UsedRange.Rows.Count is used as a last row without its starting row.",
			"UsedRange can begin below row 1, so its row count alone is not a worksheet row number.",
			"Use `ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1` when UsedRange is the intended boundary.",
		))
	}
	if currentRegionCountRe.MatchString(code) {
		findings = append(findings, a.simpleFinding(
			file, proc, lineNo, "VBA217", "warning",
			"CurrentRegion.Rows.Count is used as a last-row boundary.",
			"CurrentRegion stops at blank rows and columns, so its boundary changes when worksheet layout changes.",
			"Use `ws.Cells(ws.Rows.Count, column).End(xlUp).Row` when a column defines the intended boundary.",
		))
	}
	return findings
}

func worksheetMemberAccesses(statement string, roots *worksheetRootTracker) []worksheetMemberAccess {
	lower := strings.ToLower(maskVBAStringLiterals(statement))
	var accesses []worksheetMemberAccess
	for index := 0; index < len(lower); {
		member, start, ok := nextWorksheetMember(lower, index)
		if !ok {
			break
		}
		end := start + len(member)
		rootText := receiverBeforeMember(statement, start)
		root := roots.resolve(rootText)
		if rootText == "" {
			root = worksheetRoot{kind: worksheetRootImplicit, label: "the active worksheet"}
		}
		access := worksheetMemberAccess{member: member, root: root, start: start, end: end}
		if member == "cells" || member == "range" {
			open := skipSpace(statement, end)
			if open < len(statement) && statement[open] == '(' {
				if close := balancedCallEnd(statement, open); close > open {
					access.args = statement[open+1 : close]
					access.end = close + 1
				}
			}
		}
		accesses = append(accesses, access)
		index = end
	}
	return accesses
}

func displayWorksheetMember(member string) string {
	if member == "" {
		return member
	}
	return strings.ToUpper(member[:1]) + member[1:]
}

func maskVBAStringLiterals(text string) string {
	masked := []byte(text)
	inString := false
	for index := 0; index < len(masked); index++ {
		if masked[index] != '"' {
			if inString {
				masked[index] = ' '
			}
			continue
		}
		if inString && index+1 < len(masked) && masked[index+1] == '"' {
			masked[index], masked[index+1] = ' ', ' '
			index++
			continue
		}
		masked[index] = ' '
		inString = !inString
	}
	return string(masked)
}

func nextWorksheetMember(lower string, offset int) (string, int, bool) {
	bestStart := len(lower)
	bestMember := ""
	for _, member := range []string{"cells", "range", "rows"} {
		for search := offset; search < len(lower); {
			index := strings.Index(lower[search:], member)
			if index < 0 {
				break
			}
			start := search + index
			if identifierBoundary(lower, start-1) && identifierBoundary(lower, start+len(member)) && start < bestStart {
				bestStart, bestMember = start, member
				break
			}
			search = start + len(member)
		}
	}
	return bestMember, bestStart, bestMember != ""
}

func identifierBoundary(text string, index int) bool {
	if index < 0 || index >= len(text) {
		return true
	}
	return !isVBAIdentifierPart(text[index])
}

func receiverBeforeMember(text string, memberStart int) string {
	index := memberStart - 1
	for index >= 0 && (text[index] == ' ' || text[index] == '\t') {
		index--
	}
	if index < 0 || text[index] != '.' {
		return ""
	}
	index--
	for index >= 0 && (text[index] == ' ' || text[index] == '\t') {
		index--
	}
	if index < 0 {
		return "."
	}
	if strings.ContainsRune("(,=:+-*/&", rune(text[index])) {
		return "."
	}
	end := index + 1
	depth := 0
	for index >= 0 {
		switch text[index] {
		case ')':
			depth++
		case '(':
			if depth == 0 {
				return strings.TrimSpace(text[index+1 : end])
			}
			depth--
		case ',', '=', '+', '-', '*', '/', '&', ':':
			if depth == 0 {
				return strings.TrimSpace(text[index+1 : end])
			}
		case ' ', '\t':
			if depth == 0 {
				return strings.TrimSpace(text[index+1 : end])
			}
		}
		index--
	}
	return strings.TrimSpace(text[:end])
}

func distinctExplicitRootInArguments(accesses []worksheetMemberAccess, outer worksheetMemberAccess) (worksheetRoot, bool) {
	for _, candidate := range accesses {
		if candidate.start <= outer.start || candidate.start >= outer.end || candidate.root.kind != worksheetRootExplicit {
			continue
		}
		if candidate.root.identity != outer.root.identity {
			return candidate.root, true
		}
	}
	return worksheetRoot{}, false
}

func hasImplicitWorksheetAccess(accesses []worksheetMemberAccess) bool {
	for _, access := range accesses {
		if access.root.kind == worksheetRootImplicit {
			return true
		}
	}
	return false
}

func fullyQualifiedRootSuggestion(root worksheetRoot, member string) string {
	if root.label == "" {
		return "Use one explicit Worksheet object for every Cells, Rows, and Range member in this expression."
	}
	if member == "cells" {
		return "Use `" + root.label + ".Cells(" + root.label + ".Rows.Count, 1).End(xlUp).Row` so both objects use " + root.label + "."
	}
	return "Use `" + root.label + ".Range(" + root.label + ".Cells(...), " + root.label + ".Cells(...))` so every argument uses " + root.label + "."
}

func firstVBAIdentifier(text string) string {
	for start := 0; start < len(text); start++ {
		if !isVBAIdentifierStart(text[start]) {
			continue
		}
		end := start + 1
		for end < len(text) && isVBAIdentifierPart(text[end]) {
			end++
		}
		return text[start:end]
	}
	return ""
}

func isVBAIdentifierStart(ch byte) bool {
	return ch == '_' || ch >= 'A' && ch <= 'Z' || ch >= 'a' && ch <= 'z'
}

func isVBAIdentifierPart(ch byte) bool {
	return isVBAIdentifierStart(ch) || ch >= '0' && ch <= '9'
}

func skipSpace(text string, index int) int {
	for index < len(text) && (text[index] == ' ' || text[index] == '\t') {
		index++
	}
	return index
}

func balancedCallEnd(text string, open int) int {
	depth := 0
	inString := false
	for index := open; index < len(text); index++ {
		if text[index] == '"' {
			inString = !inString
			continue
		}
		if inString {
			continue
		}
		switch text[index] {
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				return index
			}
		}
	}
	return -1
}
