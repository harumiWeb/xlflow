package analyze

import (
	"regexp"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var redimWhileEqualityRe = regexp.MustCompile(`(?i)^\s*while\s+([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)\s*=\s*([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?|[-+]?[0-9]+)\s*$`)
var redimSelectCaseRe = regexp.MustCompile(`(?i)^\s*select\s+case\s+(.+?)\s*$`)

// redimPreserveLoopFinding reports a ReDim Preserve which executes on a
// repeated loop path. ReDim parsing deliberately stays shared with the array
// lifecycle rules; this rule only adds the loop/performance context.
func (a Analyzer) redimPreserveLoopFindings(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) []Finding {
	findings, _ := a.redimPreserveLoopFindingsPreparedWithApplicability(file, proc, moduleDecls, arrayVariables(file, proc, moduleDecls))
	return findings
}

// redimPreserveLoopFindingsPreparedWithApplicability reuses the immutable
// array catalog owned by ArrayAnalysisResult. The loop-performance projection
// has its own reachability/dependency policy, but it must not rescan
// declarations merely because another array projector already did so.
func (a Analyzer) redimPreserveLoopFindingsPreparedWithApplicability(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable) ([]Finding, bool) {
	regions := excelLoopRegions(proc)
	if len(regions) == 0 {
		return nil, false
	}
	if len(variables) == 0 {
		return nil, false
	}
	var findings []Finding
	applicable := false
	seen := map[string]bool{}
	statements := redimProcedureStatements(proc)
	constants := excelIntegerConstants(proc)
	for statement := range proc.Statements.All() {
		if statement.Recovered || !arrayRedimPreserveStatement(statement.Text) {
			continue
		}
		containing := containingExcelLoops(regions, statement.ID, statement.Range.StartLine)
		if len(containing) == 0 {
			continue
		}
		// A loop header itself is not a repeated body operation. The CFG
		// region's Body map is also what excludes unreachable statements.
		bodyLoops := containing[:0]
		for _, loop := range containing {
			if loop.Body[statement.ID] && !redimPreserveLoopBodyImpossible(proc, statements, constants, loop.StatementID) {
				bodyLoops = append(bodyLoops, loop)
			}
		}
		if len(bodyLoops) == 0 {
			continue
		}

		targets, spans := redimPreserveTargets(file, statement)
		if len(targets) == 0 {
			continue
		}
		var dynamic []string
		dynamicSeen := map[string]bool{}
		for _, target := range targets {
			// A Variant is not a known scalar in VBA: ReDim can turn it into
			// an array. Keep it eligible while fixed typed arrays remain excluded.
			if variable, ok := variables[strings.ToLower(target)]; ok && (variable.isVariant || variable.isArray && !variable.fixed) {
				name := strings.ToLower(variable.name)
				if !dynamicSeen[name] {
					dynamicSeen[name] = true
					dynamic = append(dynamic, variable.name)
				}
			}
		}
		if len(dynamic) == 0 {
			// VBA227 owns fixed arrays, scalar values, and unknown targets.
			continue
		}
		applicable = true

		dependent := false
		for _, loop := range bodyLoops {
			loopVars := redimLoopVariables(loop, proc)
			for access := range proc.Accesses.All() {
				if access.StatementID != statement.ID || !loopVars[strings.ToLower(access.Name)] {
					continue
				}
				if redimAccessInDimension(access, spans) {
					dependent = true
					break
				}
			}
			if dependent {
				break
			}
		}

		maxDepth := 1
		for _, loop := range bodyLoops {
			if loop.Depth > maxDepth {
				maxDepth = loop.Depth
			}
		}
		severity := "information"
		if dependent || maxDepth >= 2 {
			severity = "warning"
		}
		sort.Strings(dynamic)
		message := "ReDim Preserve repeatedly resizes " + strings.Join(dynamic, ", ") + " with a constant-size (loop-invariant) bound inside a loop."
		if dependent {
			message = "ReDim Preserve grows " + strings.Join(dynamic, ", ") + " from a loop-variable-dependent bound."
		}
		if maxDepth >= 2 {
			message += " Nested loop depth: " + strconvItoa(maxDepth) + "."
		}
		reason := "ReDim Preserve copies the existing array on each execution, so repeated resizing can become quadratic."
		if dependent {
			reason = "A dimension bound depends on an enclosing loop variable, causing the array to be copied as it grows on every iteration."
		}
		if maxDepth >= 2 {
			reason += " Nested loops multiply the repeated copy cost."
		}
		suggestion := "Preallocate the final array before the loop when its size is known. If the final size is unknown, grow a capacity variable geometrically and ReDim Preserve only when capacity is exhausted."
		finding := a.simpleFinding(file, proc, statement.Range.StartLine, "VBA241", severity, message, reason, suggestion)
		key := finding.Code + ":" + strconvItoa(finding.Line) + ":" + finding.Message
		if !seen[key] {
			seen[key] = true
			findings = append(findings, finding)
		}
	}
	sortFindings(findings)
	return findings, applicable
}

// redimPreserveLoopBodyImpossible recognizes a narrow, source-proven
// unreachable loop shape: a simple While equality under a Select Case clause
// compares the unchanged selector with a different enum member (or known
// integer). The CFG remains structurally conservative around Select Case, but
// VBA's case entry establishes the selector value before the loop condition is
// evaluated. Do not generalize this to arbitrary expressions or modified
// selectors; unknown paths must remain eligible for the advisory.
func redimPreserveLoopBodyImpossible(proc sourceProcedure, statements map[int]procedureir.Statement, constants map[string]int, loopID int) bool {
	loop, ok := statements[loopID]
	if !ok {
		return false
	}
	header := strings.TrimSpace(excelLoopHeaderText(loop.Text))
	match := redimWhileEqualityRe.FindStringSubmatch(header)
	if len(match) != 3 {
		return false
	}
	selectorName := strings.TrimSpace(match[1])
	loopValue := strings.TrimSpace(match[2])
	caseClause, caseValue, selectHeader, ok := redimEnclosingCase(proc, statements, loop)
	if !ok {
		return false
	}
	selectMatch := redimSelectCaseRe.FindStringSubmatch(strings.TrimSpace(excelLoopHeaderText(selectHeader.Text)))
	if len(selectMatch) != 2 || !strings.EqualFold(strings.TrimSpace(selectMatch[1]), selectorName) {
		return false
	}
	if redimCaseSelectorWrittenBeforeLoop(proc, statements, caseClause.ID, loop.ID, selectorName) {
		return false
	}
	return redimDistinctCaseValues(caseValue.Text, loopValue, constants)
}

func redimProcedureStatements(proc sourceProcedure) map[int]procedureir.Statement {
	statements := make(map[int]procedureir.Statement, proc.Statements.Len())
	for statement := range proc.Statements.All() {
		statements[statement.ID] = statement
	}
	return statements
}

func redimEnclosingCase(proc sourceProcedure, statements map[int]procedureir.Statement, loop procedureir.Statement) (caseClause, caseValue, selectHeader procedureir.Statement, ok bool) {
	for parentID := loop.ParentID; parentID != 0; {
		parent, exists := statements[parentID]
		if !exists {
			return procedureir.Statement{}, procedureir.Statement{}, procedureir.Statement{}, false
		}
		if parent.SyntaxKind == "case_clause" {
			selectStatement, exists := statements[parent.ParentID]
			if !exists || selectStatement.Kind != procedureir.StatementSelect {
				return procedureir.Statement{}, procedureir.Statement{}, procedureir.Statement{}, false
			}
			var expression procedureir.Statement
			count := 0
			for child := range proc.Statements.All() {
				if child.ParentID == parent.ID && child.SyntaxKind == "case_expression" {
					expression = child
					count++
				}
			}
			if count != 1 {
				return procedureir.Statement{}, procedureir.Statement{}, procedureir.Statement{}, false
			}
			return parent, expression, selectStatement, true
		}
		parentID = parent.ParentID
	}
	return procedureir.Statement{}, procedureir.Statement{}, procedureir.Statement{}, false
}

func redimCaseSelectorWrittenBeforeLoop(proc sourceProcedure, statements map[int]procedureir.Statement, caseClauseID, loopID int, selectorName string) bool {
	loop := statements[loopID]
	for access := range proc.Accesses.All() {
		if !strings.EqualFold(access.Name, selectorName) || (access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) {
			continue
		}
		if access.StatementID == loopID {
			continue
		}
		statement, ok := statements[access.StatementID]
		if !ok || statement.Range.StartLine >= loop.Range.StartLine || !redimStatementDescendsFrom(statements, access.StatementID, caseClauseID) {
			continue
		}
		return true
	}
	return false
}

func redimStatementDescendsFrom(statements map[int]procedureir.Statement, statementID, ancestorID int) bool {
	for statementID != 0 {
		if statementID == ancestorID {
			return true
		}
		statement, ok := statements[statementID]
		if !ok {
			return false
		}
		statementID = statement.ParentID
	}
	return false
}

func redimDistinctCaseValues(caseValue, loopValue string, constants map[string]int) bool {
	caseValue = strings.TrimSpace(caseValue)
	loopValue = strings.TrimSpace(loopValue)
	if caseValue == "" || loopValue == "" || strings.ContainsAny(caseValue, ",") || strings.Contains(strings.ToLower(caseValue), " to ") || strings.Contains(strings.ToLower(caseValue), " is ") {
		return false
	}
	if strings.EqualFold(caseValue, loopValue) {
		return false
	}
	if caseNumber, err := constantIntegerExpression(caseValue, constants); err == nil {
		if loopNumber, err := constantIntegerExpression(loopValue, constants); err == nil {
			return caseNumber != loopNumber
		}
	}
	caseDot := strings.LastIndex(caseValue, ".")
	loopDot := strings.LastIndex(loopValue, ".")
	if caseDot <= 0 || loopDot <= 0 {
		return false
	}
	return strings.EqualFold(caseValue[:caseDot], loopValue[:loopDot]) && !strings.EqualFold(caseValue[caseDot+1:], loopValue[loopDot+1:])
}

// redimLoopVariables extends the shared loop-variable extraction for Do loops.
// The IR represents pre/post `While/Until ...` conditions as a synthetic child
// StatementDo (`do_condition`), so their reads are attached to that child
// rather than the owning do_statement header.
func redimLoopVariables(owner excelLoopRegion, proc sourceProcedure) map[string]bool {
	variables := loopInvariantLoopVariables(owner, proc)
	header := procedureStatementByID(proc, owner.StatementID)
	if header.Kind != procedureir.StatementDo && header.Kind != procedureir.StatementWhile {
		return variables
	}
	conditionVariables := make(map[string]bool, len(variables))
	for name := range variables {
		conditionVariables[name] = true
	}
	for statement := range proc.Statements.All() {
		if statement.ParentID != owner.StatementID || statement.SyntaxKind != "do_condition" {
			continue
		}
		for access := range proc.Accesses.All() {
			if access.StatementID == statement.ID && access.Mode != procedureir.AccessWrite {
				conditionVariables[strings.ToLower(access.Name)] = true
			}
		}
	}
	for name := range conditionVariables {
		if redimLoopVariableModified(owner, proc, name) {
			variables[name] = true
		} else {
			delete(variables, name)
		}
	}
	return variables
}

func redimLoopVariableModified(owner excelLoopRegion, proc sourceProcedure, name string) bool {
	for access := range proc.Accesses.All() {
		if !owner.Body[access.StatementID] || !strings.EqualFold(access.Name, name) {
			continue
		}
		if access.Mode == procedureir.AccessWrite || access.Mode == procedureir.AccessReadWrite {
			return true
		}
	}
	return false
}

func procedureStatementByID(proc sourceProcedure, id int) procedureir.Statement {
	for statement := range proc.Statements.All() {
		if statement.ID == id {
			return statement
		}
	}
	return procedureir.Statement{}
}

func arrayRedimPreserveStatement(text string) bool {
	match := arrayRedimRe.FindStringSubmatch(text)
	return len(match) > 0 && strings.TrimSpace(match[1]) != ""
}

type redimDimensionSpan struct {
	start int
	end   int
}

// redimPreserveTargets returns direct ReDim targets and absolute byte spans
// for their dimension expressions. Spans prevent a loop variable used as the
// array target from being mistaken for a dynamic bound.
func redimPreserveTargets(file parsedFile, statement procedureir.Statement) ([]string, []redimDimensionSpan) {
	text := statement.Text
	base := statement.Range.StartByte
	if base >= 0 && statement.Range.EndByte > base && statement.Range.EndByte <= len(file.Source) {
		text = string(file.Source[base:statement.Range.EndByte])
	}
	match := arrayRedimRe.FindStringSubmatchIndex(text)
	if len(match) < 6 || match[2] < 0 {
		return nil, nil
	}
	restStart, restEnd := match[4], match[5]
	if restStart < 0 || restEnd < restStart {
		return nil, nil
	}
	rest := text[restStart:restEnd]
	var targets []string
	var spans []redimDimensionSpan
	for _, bounds := range splitRedimArgSpans(rest) {
		clause := strings.TrimSpace(rest[bounds[0]:bounds[1]])
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct {
			continue
		}
		// Keep dimension recognition owned by the shared array analysis. The
		// performance rule only classifies dependencies; it must not accept a
		// malformed or empty dimension clause through a second parser.
		if len(parseArrayDimensions(redim.dimensions, optionBase(file.Lines))) == 0 {
			continue
		}
		targets = append(targets, redim.name)
		open := strings.Index(clause, "(")
		close := strings.LastIndex(clause, ")")
		if open < 0 || close <= open {
			continue
		}
		leading := len(rest[bounds[0]:bounds[1]]) - len(strings.TrimLeft(rest[bounds[0]:bounds[1]], " \t\r\n"))
		spans = append(spans, redimDimensionSpan{start: base + restStart + bounds[0] + leading + open + 1, end: base + restStart + bounds[0] + leading + close})
	}
	return targets, spans
}

// splitRedimArgSpans is splitArgs with offsets retained for source ranges.
func splitRedimArgSpans(text string) [][2]int {
	var out [][2]int
	start, depth := 0, 0
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
				if strings.TrimSpace(text[start:i]) != "" {
					out = append(out, [2]int{start, i})
				}
				start = i + 1
			}
		}
	}
	if strings.TrimSpace(text[start:]) != "" {
		out = append(out, [2]int{start, len(text)})
	}
	return out
}

func redimAccessInDimension(access procedureir.VariableAccess, spans []redimDimensionSpan) bool {
	if access.Range.StartByte == 0 && access.Range.EndByte == 0 {
		// Some recovered/realtime IR records omit byte ranges. Without a range
		// we cannot prove that the access belongs to a dimension expression.
		return false
	}
	for _, span := range spans {
		if access.Range.StartByte >= span.start && access.Range.EndByte <= span.end {
			return true
		}
	}
	return false
}
