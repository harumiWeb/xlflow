package analyze

import (
	"fmt"
	"regexp"
	"sort"
	"strings"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

const value2PerformanceLargeCells = 100

var value2DateCurrencyCallRe = regexp.MustCompile(`(?i)\b(?:cdate|cvdate|ccur|formatdatetime|formatcurrency|isdate|dateadd|datediff|datepart|dateserial|datevalue|day|month|year|weekday|hour|minute|second)\s*\(`)
var value2DateCurrencyBuiltinRe = regexp.MustCompile(`(?i)\b(?:date|now|time)\b`)
var value2DateCurrencySubtypeCheckRe = regexp.MustCompile(`(?i)\b(?:vartype|typename)\s*\([^)]*\)\s*(?:=|<>|<=|>=)\s*\"?(?:vbdate|vbcurrency|date|currency)\"?\b`)
var value2DateCurrencyFormatRe = regexp.MustCompile(`(?i)\bformat\$?\s*\([^)]*(?:date|currency)`)
var value2DateCurrencyNamedCallRe = regexp.MustCompile(`(?i)\b[A-Za-z_][A-Za-z0-9_]*(?:date|currency)\s*(?:\(|\s)`)
var value2IdentifierPatterns sync.Map

func value2IdentifierPattern(name string) *regexp.Regexp {
	key := strings.ToLower(strings.TrimSpace(name))
	if cached, ok := value2IdentifierPatterns.Load(key); ok {
		return cached.(*regexp.Regexp)
	}
	pattern := regexp.MustCompile(`(?i)\b` + regexp.QuoteMeta(key) + `\b`)
	actual, _ := value2IdentifierPatterns.LoadOrStore(key, pattern)
	return actual.(*regexp.Regexp)
}

type value2PerformanceCandidate struct {
	Statement        procedureir.Statement
	Expression       procedureir.Expression
	MemberRange      ast.Range
	Receiver         procedureir.Expression
	Shape            rangeShape
	Dynamic          bool
	Write            bool
	VariantTransfer  bool
	ScalarProcessing bool
	Loops            []excelLoopRegion
}

// value2PerformanceFindings reports explicit Range.Value accesses only when
// a bulk, repeated, or Variant-array transfer signal makes the round-trip cost
// material. It deliberately does not make a global Value2 preference claim.
func (a Analyzer) value2PerformanceFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectValue2PerformanceOpportunities || a.typeDB == nil || len(proc.Statements) == 0 {
		return nil
	}
	facts := rangeValueFactsForProcedure(file, proc)
	states := value2PerformanceIncomingStates(proc, facts)
	regions := excelLoopRegions(proc)
	declarations := value2ProcedureTypes(proc)
	factsForMembers := proc.Facts
	if factsForMembers == nil {
		// Reuse one compatibility projection for all statement lookups when a
		// standalone caller has not attached procedure facts.
		factsForMembers = proc.analysisFacts()
	}
	var candidates []value2PerformanceCandidate
	seen := map[int]bool{}
	seenReceivers := map[string]int{}
	for _, statement := range proc.Statements {
		if statement.Recovered {
			continue
		}
		state, ok := states[statement.ID]
		if !ok {
			continue
		}
		for _, expression := range statementMemberExpressions(factsForMembers, statement) {
			if seen[expression.ID] {
				continue
			}
			receiver, ok := rangeValueMemberReceiver(&expression, facts, "value")
			if !ok || !value2ReceiverIsExcelRange(a, file, proc, receiver) {
				continue
			}
			memberRange, ok := value2MemberRange(expression, facts)
			if !ok {
				continue
			}
			shape, recognized := rangeValueRangeShapeExpression(&receiver, state, facts)
			dynamic := value2DynamicRangeReceiver(receiver, shape, state, facts, proc, statement.ID)
			if !recognized && !dynamic {
				continue
			}
			write := value2IsAssignmentTarget(statement, expression)
			owners := containingExcelLoops(regions, statement.ID, expression.Range.StartLine)
			owners = value2NonTrivialLoops(owners)
			large := shape.known && shape.rows > 0 && shape.cols > 0 && shape.rows*shape.cols >= value2PerformanceLargeCells
			variantTransfer := !write && value2ImmediateVariantTransfer(statement, expression, shape, dynamic, declarations)
			if !large && !dynamic && len(owners) == 0 && !variantTransfer {
				continue
			}
			candidate := value2PerformanceCandidate{
				Statement: statement, Expression: expression, MemberRange: memberRange,
				Receiver: receiver, Shape: shape, Dynamic: dynamic, Write: write,
				VariantTransfer: variantTransfer, ScalarProcessing: value2ScalarProcessing(proc, statement.ID, statement.Text), Loops: owners,
			}
			if value2IntentionalDateCurrency(proc, candidate, declarations) {
				continue
			}
			seen[expression.ID] = true
			key := strings.ToLower(strings.TrimSpace(receiver.Text))
			if existing, found := seenReceivers[key]; found {
				if len(candidates[existing].Loops) == 0 && len(candidate.Loops) > 0 {
					candidates[existing] = candidate
				}
				continue
			}
			seenReceivers[key] = len(candidates)
			candidates = append(candidates, candidate)
		}
	}
	if len(candidates) == 0 {
		return nil
	}
	sort.SliceStable(candidates, func(i, j int) bool {
		if candidates[i].MemberRange.StartLine != candidates[j].MemberRange.StartLine {
			return candidates[i].MemberRange.StartLine < candidates[j].MemberRange.StartLine
		}
		return candidates[i].MemberRange.StartColumn < candidates[j].MemberRange.StartColumn
	})
	findings := make([]Finding, 0, len(candidates))
	for _, candidate := range candidates {
		severity := "information"
		if len(candidate.Loops) > 0 {
			severity = "warning"
		}
		kind := "read"
		if candidate.Write {
			kind = "write"
		}
		message := fmt.Sprintf("Range.Value %s may benefit from Range.Value2 for this performance-sensitive access.", kind)
		reason := "Range.Value can coerce Excel dates and currency-formatted values into VBA Date or Currency subtypes; Range.Value2 returns the underlying value without those two coercions."
		if candidate.Dynamic {
			reason += " The range bounds are computed at runtime."
		} else if candidate.Shape.known {
			reason += fmt.Sprintf(" The access covers %d cells.", candidate.Shape.rows*candidate.Shape.cols)
		}
		if candidate.VariantTransfer {
			reason += " The result is transferred directly into a Variant array carrier."
		}
		if candidate.ScalarProcessing {
			reason += " The result is used for numeric or text processing."
		}
		if len(candidate.Loops) > 0 {
			reason += " The access repeats inside a reachable non-trivial loop."
		}
		finding := a.simpleFinding(file, proc, candidate.MemberRange.StartLine, "VBA243", severity, message, reason, "Use .Value2 only after confirming that Date or Currency subtypes are not required; keep .Value when those Excel coercions are intentional.")
		finding.Column = candidate.MemberRange.StartColumn + 1
		finding.EndLine = candidate.MemberRange.EndLine
		finding.EndColumn = candidate.MemberRange.EndColumn
		findings = append(findings, finding)
	}
	return findings
}

func value2ScalarProcessing(proc sourceProcedure, statementID int, statementText string) bool {
	left, _, assigned := rangeValueAssignment(statementText)
	name := rangeValueBareName(left)
	if !assigned || name == "" {
		return false
	}
	seen := false
	nameRe := value2IdentifierPattern(name)
	for _, statement := range proc.Statements {
		if statement.ID == statementID {
			seen = true
			continue
		}
		if !seen {
			continue
		}
		code := value2CodeWithoutComment(statement.Text)
		masked := strings.ToLower(value2MaskStrings(code))
		if nameRe.MatchString(masked) && value2KnownScalarProcessing(masked, name) {
			return true
		}
	}
	return false
}

func value2IsAssignmentTarget(statement procedureir.Statement, expression procedureir.Expression) bool {
	if statement.Target == nil || !sameExpression(statement.Target, expression) {
		return false
	}
	text := strings.TrimSpace(statement.Text)
	if len(text) >= len("Let ") && strings.EqualFold(text[:len("Let ")], "Let ") {
		text = strings.TrimSpace(text[len("Let "):])
	}
	target := strings.TrimSpace(expression.Text)
	return target != "" && strings.HasPrefix(strings.ToLower(text), strings.ToLower(target))
}

func value2NonTrivialLoops(owners []excelLoopRegion) []excelLoopRegion {
	out := make([]excelLoopRegion, 0, len(owners))
	for _, owner := range owners {
		if !owner.Small {
			out = append(out, owner)
		}
	}
	return out
}

func value2MemberRange(expression procedureir.Expression, facts rangeValueFacts) (ast.Range, bool) {
	current, ok := rangeValueUnwrapExpression(&expression, facts)
	if !ok || current.Kind != procedureir.ExpressionMember || len(current.Children) < 2 {
		return ast.Range{}, false
	}
	member, ok := facts.expressions[current.Children[len(current.Children)-1]]
	if !ok {
		return ast.Range{}, false
	}
	return member.Range, true
}

func value2ReceiverIsExcelRange(a Analyzer, file parsedFile, proc sourceProcedure, receiver procedureir.Expression) bool {
	lower := strings.ToLower(strings.TrimSpace(receiver.Text))
	if strings.HasPrefix(lower, "range(") && vba242ShadowedRoots(file, proc)["range"] {
		return false
	}
	typ, resolved := resolveExcelExpressionType(file, a.typeDB, receiver.Text, receiver.Range.StartLine-1, a.RootDir, a.Config)
	if resolved && isExcelRangeType(typ) {
		return true
	}
	if lower == "usedrange" && !vba242ShadowedRoots(file, proc)["usedrange"] {
		return true
	}
	for _, property := range []string{".currentregion", ".usedrange"} {
		if index := strings.Index(lower, property); index > 0 {
			base := strings.TrimSpace(receiver.Text[:index])
			baseType, baseResolved := resolveExcelExpressionType(file, a.typeDB, base, receiver.Range.StartLine-1, a.RootDir, a.Config)
			if baseResolved && (isExcelRangeType(baseType) || strings.Contains(strings.ToLower(baseType), "worksheet")) {
				return true
			}
		}
	}
	// Only recover an unqualified built-in Range call when the name is not
	// shadowed by a project declaration. Qualified members remain resolver-only
	// so user-defined objects with a Range-like member are not misclassified.
	return strings.HasPrefix(lower, "range(") && !vba242ShadowedRoots(file, proc)["range"]
}

func value2DynamicRangeReceiver(receiver procedureir.Expression, shape rangeShape, state rangeValueFlowState, facts rangeValueFacts, proc sourceProcedure, statementID int) bool {
	lower := strings.ToLower(strings.TrimSpace(receiver.Text))
	if lower == "usedrange" || strings.Contains(lower, ".currentregion") || strings.Contains(lower, ".usedrange") || strings.Contains(lower, ".resize(") || strings.Contains(lower, ".offset(") {
		return true
	}
	if name, arguments, ok := rangeValueDirectCall(receiver, facts); ok && strings.EqualFold(name, "range") && !shape.known {
		return value2DynamicRangeArguments(arguments, facts)
	}
	if current, ok := rangeValueUnwrapExpression(&receiver, facts); ok && current.Kind == procedureir.ExpressionIdentifier {
		name := strings.ToLower(strings.TrimSpace(current.Text))
		if tracked, found := state.ranges[name]; found && !tracked.known {
			return value2DynamicRangeAlias(proc, statementID, name)
		}
	}
	return false
}

func value2DynamicRangeArguments(arguments []procedureir.Expression, facts rangeValueFacts) bool {
	if len(arguments) == 2 {
		startName, _, startOK := rangeValueDirectCall(arguments[0], facts)
		endName, _, endOK := rangeValueDirectCall(arguments[1], facts)
		return startOK && endOK && strings.EqualFold(startName, "cells") && strings.EqualFold(endName, "cells")
	}
	if len(arguments) != 1 {
		return false
	}
	if _, literal := rangeValueStringLiteral(arguments[0], facts); literal {
		return false
	}
	text := strings.TrimSpace(arguments[0].Text)
	return value2RuntimeRangeArgument(text)
}

func value2DynamicRangeAlias(proc sourceProcedure, statementID int, name string) bool {
	return value2DynamicRangeAliasAt(proc, statementID, name, map[string]bool{})
}

func value2DynamicRangeAliasAt(proc sourceProcedure, statementID int, name string, visiting map[string]bool) bool {
	if visiting[name] {
		return false
	}
	visiting[name] = true
	limit := len(proc.Statements)
	for index, statement := range proc.Statements {
		if statement.ID == statementID {
			limit = index
			break
		}
	}
	for index := limit - 1; index >= 0; index-- {
		statement := proc.Statements[index]
		left, right, ok := rangeValueAssignment(statement.Text)
		left = value2TrimSetPrefix(left)
		if !ok || !strings.EqualFold(left, name) {
			continue
		}
		if value2DynamicRangeText(right) {
			return true
		}
		if alias := rangeValueBareName(right); alias != "" {
			return value2DynamicRangeAliasAt(proc, statement.ID, alias, visiting)
		}
		return false
	}
	return false
}

func value2TrimSetPrefix(text string) string {
	text = strings.TrimSpace(text)
	const prefix = "Set "
	if len(text) >= len(prefix) && strings.EqualFold(text[:len(prefix)], prefix) {
		return strings.TrimSpace(text[len(prefix):])
	}
	return text
}

func value2DynamicRangeText(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.Contains(lower, ".currentregion") || strings.Contains(lower, ".usedrange") || strings.Contains(lower, ".resize(") || strings.Contains(lower, ".offset(") {
		return true
	}
	if !strings.Contains(lower, "range(") {
		return false
	}
	if strings.Contains(lower, "cells(") && strings.Count(lower, "cells(") >= 2 {
		return true
	}
	return value2RuntimeRangeConcatenation(lower)
}

func value2RuntimeRangeConcatenation(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	open := strings.Index(lower, "range(")
	if open < 0 {
		return false
	}
	close := strings.LastIndex(lower, ")")
	if close <= open {
		return false
	}
	return value2RuntimeRangeArgument(lower[open+len("range(") : close])
}

func value2RuntimeRangeArgument(text string) bool {
	if !strings.Contains(text, "&") {
		return false
	}
	masked := strings.ReplaceAll(value2MaskStrings(strings.ToLower(text)), "&", "")
	return strings.TrimSpace(masked) != ""
}

func value2ImmediateVariantTransfer(statement procedureir.Statement, expression procedureir.Expression, shape rangeShape, dynamic bool, declarations map[string]string) bool {
	if statement.Target == nil || statement.Value == nil || !sameExpression(statement.Value, expression) {
		return false
	}
	if shape.known && shape.rows == 1 && shape.cols == 1 {
		return false
	}
	if !dynamic && !shape.array2D && !shape.known {
		return false
	}
	target := strings.ToLower(strings.TrimSpace(statement.Target.Text))
	return target != "" && strings.EqualFold(strings.TrimSpace(declarations[target]), "variant")
}

func sameExpression(left *procedureir.Expression, right procedureir.Expression) bool {
	return left != nil && left.ID == right.ID
}

func value2ProcedureTypes(proc sourceProcedure) map[string]string {
	types := map[string]string{}
	for _, declaration := range proc.Declarations {
		typ := strings.TrimSpace(declaration.Type)
		if typ == "" {
			typ = "Variant"
		}
		types[strings.ToLower(strings.TrimSpace(declaration.Name))] = typ
	}
	for _, parameter := range proc.Params {
		typ := strings.TrimSpace(parameter.Type)
		if typ == "" {
			typ = "Variant"
		}
		types[strings.ToLower(strings.TrimSpace(parameter.Name))] = typ
	}
	return types
}

func value2IntentionalDateCurrency(proc sourceProcedure, candidate value2PerformanceCandidate, declarations map[string]string) bool {
	code := value2CodeWithoutComment(candidate.Statement.Text)
	text := strings.ToLower(value2MaskStrings(code))
	if value2DateCurrencyCallRe.MatchString(text) || value2DateCurrencyBuiltinRe.MatchString(text) || value2DateCurrencyNamedCallRe.MatchString(text) || value2DateCurrencyFormat(code) || strings.Contains(text, "vbdate") || strings.Contains(text, "vbcurrency") || value2DateCurrencySubtypeCheck(code) || strings.Contains(text, "#") || strings.Contains(text, "@") {
		return true
	}
	left, right, assigned := rangeValueAssignment(candidate.Statement.Text)
	if candidate.Write && assigned && value2DateCurrencyExpression(right, declarations) {
		return true
	}
	if !candidate.Write && assigned {
		name := rangeValueBareName(left)
		if name != "" && value2DateCurrencyType(declarations[name]) {
			return true
		}
		if name != "" && value2LaterDateCurrencyUse(proc, candidate.Statement.ID, name, declarations) {
			return true
		}
	}
	return false
}

func value2DateCurrencyType(typ string) bool {
	lower := strings.ToLower(strings.TrimSpace(typ))
	return lower == "date" || lower == "currency"
}

func value2DateCurrencyExpression(text string, declarations map[string]string) bool {
	code := value2CodeWithoutComment(text)
	masked := strings.ToLower(value2MaskStrings(code))
	if value2DateCurrencyCallRe.MatchString(masked) || value2DateCurrencyBuiltinRe.MatchString(masked) || value2DateCurrencyNamedCallRe.MatchString(masked) || value2DateCurrencyFormat(code) || value2DateCurrencySubtypeCheck(code) || strings.Contains(masked, "#") || strings.Contains(masked, "@") {
		return true
	}
	for name, typ := range declarations {
		if value2DateCurrencyType(typ) && value2IdentifierPattern(name).MatchString(masked) {
			return true
		}
	}
	return false
}

func value2LaterDateCurrencyUse(proc sourceProcedure, statementID int, name string, declarations map[string]string) bool {
	seen := false
	nameRe := value2IdentifierPattern(name)
	for _, statement := range proc.Statements {
		if statement.ID == statementID {
			seen = true
			continue
		}
		code := value2CodeWithoutComment(statement.Text)
		masked := value2MaskStrings(code)
		if !seen || !nameRe.MatchString(masked) {
			continue
		}
		text := strings.ToLower(masked)
		if value2DateCurrencyCallRe.MatchString(text) || value2DateCurrencyBuiltinRe.MatchString(text) || value2DateCurrencyNamedCallRe.MatchString(text) || value2DateCurrencyFormat(code) || strings.Contains(text, "vbdate") || strings.Contains(text, "vbcurrency") || value2DateCurrencySubtypeCheck(code) || strings.Contains(text, "#") || strings.Contains(text, "@") {
			return true
		}
		left, _, assigned := rangeValueAssignment(code)
		if assigned && value2DateCurrencyType(declarations[strings.ToLower(strings.TrimSpace(left))]) {
			return true
		}
		if !value2KnownScalarProcessing(text, name) {
			return true
		}
	}
	return false
}

func value2DateCurrencySubtypeCheck(text string) bool {
	masked := value2MaskStringsExceptDateCurrency(text)
	return value2DateCurrencySubtypeCheckRe.MatchString(strings.ToLower(masked))
}

func value2DateCurrencyFormat(text string) bool {
	masked := value2MaskStringsExceptDateCurrency(text)
	return value2DateCurrencyFormatRe.MatchString(strings.ToLower(masked))
}

func value2CodeWithoutComment(text string) string {
	inString := false
	for i := 0; i < len(text); i++ {
		if !inString && (i == 0 || text[i-1] == ':' || text[i-1] == ' ' || text[i-1] == '\t') && i+3 <= len(text) && strings.EqualFold(text[i:i+3], "rem") && (i+3 == len(text) || text[i+3] == ' ' || text[i+3] == '\t') {
			return text[:i]
		}
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '\'':
			if !inString {
				return text[:i]
			}
		}
	}
	return text
}

func value2MaskStrings(text string) string {
	var builder strings.Builder
	builder.Grow(len(text))
	inString := false
	for i := 0; i < len(text); i++ {
		if text[i] == '"' {
			builder.WriteByte(' ')
			if inString && i+1 < len(text) && text[i+1] == '"' {
				builder.WriteByte(' ')
				i++
				continue
			}
			inString = !inString
			continue
		}
		if inString {
			if text[i] == '\r' || text[i] == '\n' {
				builder.WriteByte(text[i])
			} else {
				builder.WriteByte(' ')
			}
			continue
		}
		builder.WriteByte(text[i])
	}
	return builder.String()
}

func value2MaskStringsExceptDateCurrency(text string) string {
	var builder strings.Builder
	builder.Grow(len(text))
	for i := 0; i < len(text); i++ {
		if text[i] != '"' {
			builder.WriteByte(text[i])
			continue
		}
		start := i
		i++
		for i < len(text) {
			if text[i] == '"' {
				if i+1 < len(text) && text[i+1] == '"' {
					i += 2
					continue
				}
				i++
				break
			}
			i++
		}
		literal := text[start:i]
		lower := strings.ToLower(strings.Trim(literal, `"`))
		if lower == "date" || lower == "currency" {
			builder.WriteString(literal)
			i--
			continue
		}
		for range literal {
			builder.WriteByte(' ')
		}
		i--
	}
	return builder.String()
}

func value2KnownScalarProcessing(text, name string) bool {
	lower := strings.ToLower(text)
	if strings.Contains(lower, "debug.print") || strings.Contains(lower, "msgbox") || strings.Contains(lower, "print #") {
		return true
	}
	if strings.Contains(lower, "+") || strings.Contains(lower, "-") || strings.Contains(lower, "*") || strings.Contains(lower, "/") || strings.Contains(lower, "&") || strings.Contains(lower, "=") || strings.Contains(lower, ">") || strings.Contains(lower, "<") {
		return true
	}
	for _, fn := range []string{"len", "trim", "ltrim", "rtrim", "left", "right", "mid", "instr", "replace", "strcomp", "cstr", "cdbl", "clng", "cint", "isnumeric", "round"} {
		if strings.Contains(lower, fn+"(") {
			return true
		}
	}
	return false
}

func value2PerformanceIncomingStates(proc sourceProcedure, facts rangeValueFacts) map[int]rangeValueFlowState {
	states := map[int]rangeValueFlowState{}
	if proc.Graph == nil {
		state := newRangeValueFlowState()
		for _, statement := range proc.Statements {
			states[statement.ID] = cloneRangeValueFlowState(state)
			_, state = rangeValueStatement(state, statement, facts)
		}
		return states
	}
	reachable := map[cfg.BlockID]bool{}
	for _, id := range proc.Graph.Reachable(cfg.EdgeFilter{NormalOnly: true}) {
		reachable[id] = true
	}
	if !reachable[proc.Graph.Entry] {
		return states
	}
	blocks := map[cfg.BlockID]cfg.Block{}
	for _, block := range proc.Graph.Blocks {
		blocks[block.ID] = block
	}
	byBlock := map[cfg.BlockID]rangeValueFlowState{proc.Graph.Entry: newRangeValueFlowState()}
	queue := []cfg.BlockID{proc.Graph.Entry}
	queued := map[cfg.BlockID]bool{proc.Graph.Entry: true}
	for len(queue) > 0 {
		id := queue[0]
		queue = queue[1:]
		queued[id] = false
		in := byBlock[id]
		out := cloneRangeValueFlowState(in)
		if block, ok := blocks[id]; ok && block.Statement != nil {
			_, out = rangeValueStatement(in, *block.Statement, facts)
		}
		for _, edge := range proc.Graph.Edges {
			if edge.From != id || edge.Class != cfg.EdgeNormal || !reachable[edge.To] {
				continue
			}
			next, changed := mergeRangeValueFlowState(byBlock[edge.To], out, byBlock[edge.To].values != nil || byBlock[edge.To].ranges != nil || byBlock[edge.To].arrayGuards != nil)
			if changed {
				byBlock[edge.To] = next
				if !queued[edge.To] {
					queue = append(queue, edge.To)
					queued[edge.To] = true
				}
			}
		}
	}
	for id, block := range blocks {
		if block.Statement != nil {
			if state, ok := byBlock[id]; ok {
				states[block.Statement.ID] = state
			}
		}
	}
	return states
}
