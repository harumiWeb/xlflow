package analyze

import (
	"regexp"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var (
	loopInvariantForIteratorRe = regexp.MustCompile(`(?i)^\s*for\s+([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	loopInvariantForEachRe     = regexp.MustCompile(`(?i)^\s*for\s+each\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\b`)
	loopInvariantConstStringRe = regexp.MustCompile(`(?i)^\s*(?:(?:public|private|friend|static)\s+)?const\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:as\s+[A-Za-z_][A-Za-z0-9_]*)?\s*=\s*("(?:""|[^"])*")\s*$`)
)

var loopInvariantSelectorMembers = map[string]bool{
	"workbooks": true, "worksheets": true, "sheets": true,
	"range": true, "names": true, "listobjects": true, "listcolumns": true,
	"pivottables": true, "pivotfields": true,
	"chartobjects": true, "charts": true, "seriescollection": true,
}

type loopInvariantCandidate struct {
	Expression procedureir.Expression
	Statement  procedureir.Statement
	Chain      string
	Key        string
	Type       string
	Owner      excelLoopRegion
	Depth      int
}

type loopInvariantChainSegment struct {
	Text      string
	Start     int
	End       int
	Name      string
	Arguments string
	IsCall    bool
}

func (a Analyzer) excelLoopInvariantFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectLoopInvariantExcelObjectResolution || a.typeDB == nil {
		return nil
	}
	regions := excelLoopRegions(proc)
	if len(regions) == 0 {
		return nil
	}
	statements := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statements[statement.ID] = statement
	}
	facts := proc.Facts
	if facts == nil {
		// Standalone callers may omit the attached projection. Share one
		// compatibility projection across all statements in this procedure.
		facts = proc.analysisFacts()
	}
	constants := loopInvariantStringConstants(file, proc)
	expressionsByStatement := make(map[int][]procedureir.Expression)
	for _, statement := range proc.Statements {
		expressionsByStatement[statement.ID] = statementMemberExpressions(facts, statement)
	}
	seen := map[string]loopInvariantCandidate{}
	for _, statement := range proc.Statements {
		for _, expression := range expressionsByStatement[statement.ID] {
			candidate, ok := a.loopInvariantCandidate(file, proc, statement, expression, constants, statements)
			if !ok {
				continue
			}
			owners := containingExcelLoops(regions, statement.ID, expression.Range.StartLine)
			owner, ok := loopInvariantOwner(proc, owners, statement, expression, statements)
			if !ok {
				continue
			}
			candidate.Owner = owner
			candidate.Depth = owner.Depth
			key := strings.Join([]string{strconvItoa(owner.StatementID), candidate.Key}, "\x00")
			if previous, exists := seen[key]; !exists || candidate.Expression.Range.StartByte < previous.Expression.Range.StartByte {
				seen[key] = candidate
			}
		}
	}
	if len(seen) == 0 {
		return nil
	}
	candidates := make([]loopInvariantCandidate, 0, len(seen))
	for _, candidate := range seen {
		maximal := true
		for _, other := range seen {
			if other.Owner.StatementID != candidate.Owner.StatementID || other.Key == candidate.Key {
				continue
			}
			if strings.HasPrefix(other.Key, candidate.Key+".") {
				maximal = false
				break
			}
		}
		if !maximal {
			continue
		}
		candidates = append(candidates, candidate)
	}
	sort.SliceStable(candidates, func(i, j int) bool {
		if candidates[i].Expression.Range.StartLine != candidates[j].Expression.Range.StartLine {
			return candidates[i].Expression.Range.StartLine < candidates[j].Expression.Range.StartLine
		}
		if candidates[i].Expression.Range.StartColumn != candidates[j].Expression.Range.StartColumn {
			return candidates[i].Expression.Range.StartColumn < candidates[j].Expression.Range.StartColumn
		}
		return candidates[i].Key < candidates[j].Key
	})
	findings := make([]Finding, 0, len(candidates))
	issuedNames := map[string]bool{}
	for _, candidate := range candidates {
		variable := loopInvariantVariableName(proc, candidate.Type, issuedNames)
		message := "Loop repeatedly resolves invariant Excel object " + candidate.Chain + "."
		if candidate.Depth >= 2 {
			message += " Nested loop depth: " + strconvItoa(candidate.Depth) + "."
		}
		reason := "The expression does not depend on the enclosing loop variable, so Excel resolves the same object on every iteration."
		if candidate.Depth >= 2 {
			reason += " Nested loops multiply the repeated resolution cost."
		}
		suggestion := "Extract it before the loop: Dim " + variable + " As " + loopInvariantFriendlyType(candidate.Type) + "; Set " + variable + " = " + candidate.Chain + "; replace " + candidate.Chain + " with " + variable + " inside the loop."
		finding := a.simpleFinding(file, proc, candidate.Expression.Range.StartLine, "VBA238", "warning", message, reason, suggestion)
		finding.EndLine = candidate.Expression.Range.EndLine
		finding.EndColumn = candidate.Expression.Range.EndColumn
		findings = append(findings, finding)
	}
	return findings
}

func (a Analyzer) loopInvariantCandidate(file parsedFile, proc sourceProcedure, statement procedureir.Statement, expression procedureir.Expression, constants map[string]string, statements map[int]procedureir.Statement) (loopInvariantCandidate, bool) {
	expression.Text = loopInvariantExpressionSource(file.Source, expression)
	loopInvariantExpandRange(&expression, expression.Text)
	segments := splitLoopInvariantChain(expression.Text)
	if len(segments) == 0 {
		return loopInvariantCandidate{}, false
	}
	last := segments[len(segments)-1]
	if !last.IsCall || !loopInvariantSelectorMembers[strings.ToLower(last.Name)] {
		return loopInvariantCandidate{}, false
	}
	if !loopInvariantConstantSelector(normalizeLoopInvariantChain(last.Arguments), constants) {
		return loopInvariantCandidate{}, false
	}
	withPrefix, _ := loopInvariantWithPrefix(statement.ID, statements)
	chain := strings.TrimSpace(expression.Text)
	if withPrefix != "" && (strings.HasPrefix(strings.TrimSpace(chain), ".") || !strings.Contains(chain, ".")) {
		chain = strings.TrimSpace(withPrefix) + strings.TrimSpace(chain)
	}
	if chain == "" || !strings.Contains(chain, "(") {
		return loopInvariantCandidate{}, false
	}
	normalizedChain := normalizeLoopInvariantChain(chain)
	typ, resolved := resolveExcelExpressionType(file, a.typeDB, normalizedChain, expression.Range.StartLine-1, a.RootDir, a.Config)
	if !resolved && strings.EqualFold(last.Name, "charts") {
		// The Excel type database models ChartObjects but not the legacy
		// Worksheet/Workbook Charts collection. Resolve that selector from a
		// proven workbook/worksheet/application receiver instead of guessing
		// for late-bound objects.
		chainSegments := splitLoopInvariantChain(chain)
		if len(chainSegments) > 1 {
			receiver := strings.TrimSpace(chain[:chainSegments[len(chainSegments)-1].Start])
			receiver = strings.TrimSuffix(receiver, ".")
			receiverType, receiverResolved := resolveExcelExpressionType(file, a.typeDB, normalizeLoopInvariantChain(receiver), expression.Range.StartLine-1, a.RootDir, a.Config)
			if receiverResolved && isLoopInvariantChartCollectionReceiver(receiverType) {
				typ, resolved = "Excel.Chart", true
			}
		}
	}
	if !resolved || !isLoopInvariantExcelObjectType(typ) {
		return loopInvariantCandidate{}, false
	}
	if !loopInvariantHasSelector(chain) {
		return loopInvariantCandidate{}, false
	}
	return loopInvariantCandidate{
		Expression: expression,
		Statement:  statement,
		Chain:      strings.TrimSpace(chain),
		Key:        normalizeLoopInvariantChain(chain),
		Type:       typ,
	}, true
}

func loopInvariantExpandRange(expression *procedureir.Expression, text string) {
	if expression == nil {
		return
	}
	expression.Range.EndByte = expression.Range.StartByte + len(text)
	if newline := loopInvariantLastLineBreak(text); newline >= 0 {
		lineBreaks := strings.Count(text[:newline+1], "\n")
		if lineBreaks == 0 {
			lineBreaks = strings.Count(text[:newline+1], "\r")
		}
		expression.Range.EndLine = expression.Range.StartLine + lineBreaks
		expression.Range.EndColumn = len(text[newline+1:]) + 1
		return
	}
	expression.Range.EndLine = expression.Range.StartLine
	expression.Range.EndColumn = expression.Range.StartColumn + len(text)
}

func loopInvariantLastLineBreak(text string) int {
	if newline := strings.LastIndexByte(text, '\n'); newline >= 0 {
		return newline
	}
	return strings.LastIndexByte(text, '\r')
}

func loopInvariantExpressionSource(source []byte, expression procedureir.Expression) string {
	start, end := expression.Range.StartByte, expression.Range.EndByte
	if start < 0 || end < start || start >= len(source) {
		return expression.Text
	}
	if end > len(source) {
		end = len(source)
	}
	text := string(source[start:end])
	position := end
	for position < len(source) {
		for position < len(source) && (source[position] == ' ' || source[position] == '\t' || source[position] == '\r' || source[position] == '\n') {
			position++
		}
		if position < len(source) && source[position] == '_' && loopInvariantContinuation(string(source), position) {
			position++
			continue
		}
		break
	}
	if position >= len(source) || source[position] != '(' {
		return text
	}
	callEnd := loopInvariantCallEnd(source, position)
	if callEnd <= position {
		return text
	}
	return string(source[start:callEnd])
}

func loopInvariantCallEnd(source []byte, open int) int {
	depth := 0
	inString := false
	for i := open; i < len(source); i++ {
		switch source[i] {
		case '"':
			if inString && i+1 < len(source) && source[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString {
				depth--
				if depth == 0 {
					return i + 1
				}
			}
		}
	}
	return 0
}

func loopInvariantOwner(proc sourceProcedure, owners []excelLoopRegion, statement procedureir.Statement, expression procedureir.Expression, statements map[int]procedureir.Statement) (excelLoopRegion, bool) {
	withStatements := candidateWithStatements(statement.ID, statements)
	for _, owner := range owners {
		if owner.Small || owner.StatementID == statement.ID {
			continue
		}
		if loopInvariantDependsOnLoopVariable(proc, expression, withStatements, &owner) {
			continue
		}
		if loopInvariantRootWritten(proc, expression, withStatements, owner) {
			continue
		}
		return owner, true
	}
	return excelLoopRegion{}, false
}

func loopInvariantDependsOnLoopVariable(proc sourceProcedure, expression procedureir.Expression, withStatements []procedureir.Statement, owner *excelLoopRegion) bool {
	variables := map[string]bool{}
	if owner != nil {
		variables = loopInvariantLoopVariables(*owner, proc)
	}
	if len(variables) == 0 && len(withStatements) == 0 {
		return false
	}
	for _, access := range proc.Accesses {
		if !loopInvariantRangeContains(expression.Range, access.Range) {
			continue
		}
		if variables[strings.ToLower(access.Name)] {
			return true
		}
	}
	for _, statement := range withStatements {
		for _, access := range proc.Accesses {
			if access.StatementID == statement.ID && variables[strings.ToLower(access.Name)] {
				return true
			}
		}
	}
	return false
}

func loopInvariantRootWritten(proc sourceProcedure, expression procedureir.Expression, withStatements []procedureir.Statement, owner excelLoopRegion) bool {
	dependencies := map[string]bool{}
	for _, access := range proc.Accesses {
		if loopInvariantRangeContains(expression.Range, access.Range) {
			dependencies[strings.ToLower(access.Name)] = true
		}
	}
	for _, statement := range withStatements {
		for _, access := range proc.Accesses {
			if access.StatementID == statement.ID {
				dependencies[strings.ToLower(access.Name)] = true
			}
		}
	}
	if len(dependencies) == 0 {
		return false
	}
	for _, access := range proc.Accesses {
		if !owner.Body[access.StatementID] || !dependencies[strings.ToLower(access.Name)] {
			continue
		}
		if access.Mode == procedureir.AccessWrite || access.Mode == procedureir.AccessReadWrite {
			return true
		}
	}
	return false
}

func loopInvariantLoopVariables(owner excelLoopRegion, proc sourceProcedure) map[string]bool {
	variables := map[string]bool{}
	var header procedureir.Statement
	for _, statement := range proc.Statements {
		if statement.ID == owner.StatementID {
			header = statement
			break
		}
	}
	text := strings.TrimSpace(excelLoopHeaderText(header.Text))
	if match := loopInvariantForIteratorRe.FindStringSubmatch(text); len(match) == 2 {
		variables[strings.ToLower(match[1])] = true
	}
	if match := loopInvariantForEachRe.FindStringSubmatch(text); len(match) == 2 {
		variables[strings.ToLower(match[1])] = true
	}
	if len(variables) == 0 {
		// Do/While headers have no dedicated iterator syntax. Treat variables
		// read by the header condition as loop controls so dynamic selectors
		// such as Worksheets(names(i)) remain conservative.
		for _, access := range proc.Accesses {
			if access.StatementID == owner.StatementID && access.Mode != procedureir.AccessWrite {
				variables[strings.ToLower(access.Name)] = true
			}
		}
	}
	return variables
}

func loopInvariantWithPrefix(statementID int, statements map[int]procedureir.Statement) (string, []procedureir.Statement) {
	var chain []string
	var withStatements []procedureir.Statement
	current, ok := statements[statementID]
	for ok && current.ParentID != 0 {
		parent, found := statements[current.ParentID]
		if !found {
			break
		}
		if parent.Kind == procedureir.StatementWith {
			text := strings.TrimSpace(loopInvariantWithHeaderText(parent.Text))
			if len(text) >= len("With ") && strings.EqualFold(text[:len("With ")], "With ") {
				chain = append(chain, strings.TrimSpace(text[len("With "):]))
				withStatements = append(withStatements, parent)
			}
		}
		current = parent
	}
	for i, j := 0, len(chain)-1; i < j; i, j = i+1, j-1 {
		chain[i], chain[j] = chain[j], chain[i]
		withStatements[i], withStatements[j] = withStatements[j], withStatements[i]
	}
	if len(chain) == 0 {
		return "", nil
	}
	result := strings.TrimSpace(chain[0])
	for _, part := range chain[1:] {
		part = strings.TrimSpace(part)
		if strings.HasPrefix(part, ".") {
			result += part
		} else {
			result += "." + part
		}
	}
	return result, withStatements
}

func loopInvariantWithHeaderText(text string) string {
	lines := strings.Split(strings.ReplaceAll(text, "\r\n", "\n"), "\n")
	if len(lines) == 0 {
		return text
	}
	parts := make([]string, 0, len(lines))
	for _, line := range lines {
		line = rawWorksheetCodeLine(line)
		continues := vbaLineContinues(line)
		if continues {
			line = strings.TrimSpace(strings.TrimSuffix(strings.TrimSpace(line), "_"))
		}
		if line != "" {
			parts = append(parts, line)
		}
		if !continues {
			break
		}
	}
	return strings.Join(parts, " ")
}

func candidateWithStatements(statementID int, statements map[int]procedureir.Statement) []procedureir.Statement {
	_, withStatements := loopInvariantWithPrefix(statementID, statements)
	return withStatements
}

func loopInvariantRangeContains(outer, inner vbaast.Range) bool {
	return inner.StartByte >= outer.StartByte && inner.EndByte <= outer.EndByte
}

func loopInvariantHasSelector(chain string) bool {
	for _, segment := range splitLoopInvariantChain(chain) {
		if loopInvariantSelectorMembers[strings.ToLower(segment.Name)] && segment.IsCall {
			return true
		}
	}
	return false
}

func isLoopInvariantExcelObjectType(typ string) bool {
	lower := strings.ToLower(strings.TrimSpace(typ))
	for _, name := range []string{"workbook", "worksheet", "range", "name", "listobject", "listcolumn", "pivottable", "pivotfield", "chart", "chartobject", "series"} {
		if strings.HasSuffix(lower, name) || strings.Contains(lower, "."+name) {
			return true
		}
	}
	return false
}

func isLoopInvariantChartCollectionReceiver(typ string) bool {
	lower := strings.ToLower(strings.TrimSpace(typ))
	return strings.HasSuffix(lower, "workbook") || strings.HasSuffix(lower, "worksheet") || strings.HasSuffix(lower, "application")
}

func splitLoopInvariantChain(text string) []loopInvariantChainSegment {
	var segments []loopInvariantChainSegment
	start := 0
	depth := 0
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '.':
			if !inString && depth == 0 {
				segments = append(segments, loopInvariantChainSegment{Text: text[start:i], Start: start, End: i})
				start = i + 1
			}
		}
	}
	segments = append(segments, loopInvariantChainSegment{Text: text[start:], Start: start, End: len(text)})
	for i := range segments {
		segment := strings.TrimSpace(segments[i].Text)
		if strings.HasPrefix(segment, ".") {
			segment = strings.TrimSpace(strings.TrimPrefix(segment, "."))
		}
		segments[i].Text = segment
		open := strings.IndexByte(segment, '(')
		if open < 0 || !strings.HasSuffix(strings.TrimSpace(segment), ")") {
			segments[i].Name = strings.TrimSpace(segment)
			continue
		}
		segments[i].Name = strings.TrimSpace(segment[:open])
		segments[i].Arguments = strings.TrimSpace(segment[open+1 : len(strings.TrimSpace(segment))-1])
		segments[i].IsCall = loopInvariantIdentifier(segments[i].Name)
	}
	return segments
}

func loopInvariantConstantSelector(arguments string, constants map[string]string) bool {
	arguments = strings.TrimSpace(arguments)
	if arguments == "" {
		return false
	}
	parts := loopInvariantArgumentParts(arguments)
	if len(parts) == 0 {
		return false
	}
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if loopInvariantStringLiteral(part) {
			continue
		}
		if _, ok := constants[strings.ToLower(cleanIdentifier(part))]; !ok {
			return false
		}
	}
	return true
}

func loopInvariantArgumentParts(text string) []string {
	var parts []string
	start := 0
	depth := 0
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case ',':
			if !inString && depth == 0 {
				parts = append(parts, text[start:i])
				start = i + 1
			}
		}
	}
	if inString || depth != 0 {
		return nil
	}
	parts = append(parts, text[start:])
	return parts
}

func loopInvariantStringLiteral(text string) bool {
	text = strings.TrimSpace(text)
	return len(text) >= 2 && text[0] == '"' && text[len(text)-1] == '"'
}

func loopInvariantStringConstants(file parsedFile, proc sourceProcedure) map[string]string {
	constants := map[string]string{}
	for lineNumber, line := range file.Lines {
		// File-level declarations precede a procedure in normal VBA modules.
		// Do not let a procedure-local Const from a different procedure prove
		// a selector in this one.
		if proc.StartLine > 0 && lineNumber+1 > proc.StartLine {
			break
		}
		line = strings.TrimSpace(gui.StripComment(line))
		match := loopInvariantConstStringRe.FindStringSubmatch(line)
		if len(match) == 2 {
			constants[strings.ToLower(match[1])] = match[2]
		}
	}
	for _, statement := range proc.Statements {
		line := strings.TrimSpace(gui.StripComment(statement.Text))
		match := loopInvariantConstStringRe.FindStringSubmatch(line)
		if len(match) == 2 {
			constants[strings.ToLower(match[1])] = match[2]
		}
	}
	return constants
}

func normalizeLoopInvariantChain(text string) string {
	var out strings.Builder
	inString := false
	lineStart := true
	for i := 0; i < len(text); i++ {
		ch := text[i]
		if inString {
			out.WriteByte(ch)
			if ch == '"' {
				if i+1 < len(text) && text[i+1] == '"' {
					out.WriteByte(text[i+1])
					i++
				} else {
					inString = false
				}
			}
			lineStart = false
			continue
		}
		if ch == '"' {
			inString = true
			out.WriteByte(ch)
			lineStart = false
			continue
		}
		if ch == '\'' {
			for i < len(text) && text[i] != '\n' && text[i] != '\r' {
				i++
			}
			i--
			continue
		}
		if (ch == '\n' || ch == '\r') || ch == '_' && loopInvariantContinuation(text, i) || ch == ' ' || ch == '\t' {
			if ch == '\n' || ch == '\r' {
				lineStart = true
			}
			continue
		}
		if lineStart && (ch == 'R' || ch == 'r') && i+3 < len(text) && strings.EqualFold(text[i:i+3], "Rem") && isVBAIdentifierBoundary(text, i+3) {
			for i < len(text) && text[i] != '\n' && text[i] != '\r' {
				i++
			}
			i--
			continue
		}
		if isVBAIdentifierStart(ch) {
			start := i
			for i+1 < len(text) && isVBAIdentifierPart(text[i+1]) {
				i++
			}
			out.WriteString(strings.ToLower(text[start : i+1]))
			lineStart = false
			continue
		}
		out.WriteByte(ch)
		lineStart = false
	}
	return out.String()
}

func loopInvariantContinuation(text string, index int) bool {
	if index < 0 || index >= len(text) || text[index] != '_' {
		return false
	}
	for i := index + 1; i < len(text); i++ {
		if text[i] == ' ' || text[i] == '\t' {
			continue
		}
		return text[i] == '\n' || text[i] == '\r' || text[i] == ')'
	}
	return true
}

func isVBAIdentifierBoundary(text string, index int) bool {
	return index >= len(text) || !isVBAIdentifierPart(text[index])
}

func loopInvariantIdentifier(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" || !isVBAIdentifierStart(text[0]) {
		return false
	}
	for i := 1; i < len(text); i++ {
		if !isVBAIdentifierPart(text[i]) {
			return false
		}
	}
	return true
}

func loopInvariantFriendlyType(typ string) string {
	typ = strings.TrimSpace(typ)
	if dot := strings.LastIndexByte(typ, '.'); dot >= 0 {
		typ = typ[dot+1:]
	}
	if typ == "" {
		return "Object"
	}
	return typ
}

func loopInvariantVariableName(proc sourceProcedure, typ string, issued map[string]bool) string {
	base := "cached" + loopInvariantFriendlyType(typ)
	used := map[string]bool{}
	for name := range issued {
		used[name] = true
	}
	for _, declaration := range proc.Declarations {
		used[strings.ToLower(declaration.Name)] = true
	}
	for _, param := range proc.Params {
		used[strings.ToLower(param.Name)] = true
	}
	for _, access := range proc.Accesses {
		used[strings.ToLower(access.Name)] = true
	}
	name := base
	for index := 2; used[strings.ToLower(name)]; index++ {
		name = base + strconvItoa(index)
	}
	issued[strings.ToLower(name)] = true
	return name
}
