package analyze

import (
	"cmp"
	"maps"
	"regexp"
	"slices"
	"strings"
	"unicode"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// VBA250 tracks only facts that are safe to carry across a procedure-local
// CFG. An empty identity means unknown; the identity strings are deliberately
// source-derived and are never presented to users.
type activeUIState struct {
	activeWorkbook  string
	activeWorksheet string
	resumeNext      bool
	bindings        map[string]activeUIValue
}

type activeUIValueKind uint8

const (
	activeUIUnknown activeUIValueKind = iota
	activeUIWorkbook
	activeUIWorksheet
	activeUIRange
)

type activeUIValue struct {
	kind     activeUIValueKind
	id       string
	parent   string
	workbook string
}

const (
	activeUIActiveWorkbook  = "workbook:active"
	activeUIActiveWorksheet = "worksheet:active"
)

type activeUIOperation struct {
	method              string
	receiver            string
	selectReplaces      bool
	selectModeSpecified bool
}

var (
	activeUIWithExpressionRe = regexp.MustCompile(`(?i)^\s*with\s+(.+)$`)
	activeUISetAssignmentRe  = regexp.MustCompile(`(?i)^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$`)
	activeUIIsComparisonRe   = regexp.MustCompile(`(?i)^(.+?)\s+is\s+(?:(not)\s+)?(.+?)\s*$`)
)

// activeUIStateFindings solves one procedure without consulting callers or
// callees. It intentionally fails open when a procedure has no CFG.
func (a Analyzer) activeUIStateFindings(file parsedFile, proc sourceProcedure) []Finding {
	if proc.Graph == nil {
		return nil
	}

	view := proc.Graph.View(cfg.EdgeFilter{})
	statements := make(map[int]procedureir.Statement, proc.Statements.Len())
	for statement := range proc.Statements.All() {
		statements[statement.ID] = statement
	}
	withExpressions := activeUIWithExpressions(statements)
	declarations := newDeclarationScope(file, proc)
	callsByStatement := activeUICallsByStatement(proc)
	unknownSources := make(map[cfg.BlockID]bool, len(proc.Graph.UnknownFlowSources))
	for _, source := range proc.Graph.UnknownFlowSources {
		unknownSources[source] = true
	}

	inStates := make(map[cfg.BlockID]activeUIState)
	initialized := make(map[cfg.BlockID]bool)
	entry := view.Entry()
	inStates[entry] = newActiveUIState()
	initialized[entry] = true
	queue := []cfg.BlockID{entry}
	inQueue := map[cfg.BlockID]bool{entry: true}

	for len(queue) > 0 {
		blockID := queue[0]
		queue = queue[1:]
		delete(inQueue, blockID)
		block, ok := view.BlockByID(blockID)
		if !ok {
			continue
		}
		state := inStates[blockID].clone()
		if unknownSources[blockID] {
			state.invalidateActive()
		}
		out := a.activeUITransfer(file, proc, declarations, withExpressions, callsByStatement, block.Statement, state)
		view.ForEachOutgoing(blockID, func(edge cfg.Edge) bool {
			edgeState := out.clone()
			if edge.Class == cfg.EdgeExceptional || edge.Uncertain {
				edgeState.invalidateActive()
			}
			if block.Statement != nil && !a.activeUIHasUnknownCall(file, proc, declarations, state, block.Statement.ID, callsByStatement, block.Statement, activeUIText(*block.Statement, withExpressions)) {
				edgeState = a.activeUIRefineBranch(file, proc, declarations, block.Statement, edge, edgeState)
			}
			if !initialized[edge.To] {
				inStates[edge.To] = edgeState
				initialized[edge.To] = true
				if !inQueue[edge.To] {
					queue = append(queue, edge.To)
					inQueue[edge.To] = true
				}
				return true
			}
			joined, changed := joinActiveUIState(inStates[edge.To], edgeState)
			if changed {
				inStates[edge.To] = joined
				if !inQueue[edge.To] {
					queue = append(queue, edge.To)
					inQueue[edge.To] = true
				}
			}
			return true
		})
	}

	findings := make([]Finding, 0)
	view.ForEachBlock(func(block cfg.Block) bool {
		if !initialized[block.ID] || block.Statement == nil {
			return true
		}
		statement := *block.Statement
		text := activeUIText(statement, withExpressions)
		operation, ok := parseActiveUIOperation(text)
		if !ok {
			return true
		}
		state := inStates[block.ID].clone()
		if unknownSources[block.ID] {
			state.invalidateActive()
		}
		if a.activeUIHasUnknownCall(file, proc, declarations, state, statement.ID, callsByStatement, &statement, text) {
			state.invalidateActive()
		}
		value := a.resolveActiveUIValue(file, proc, declarations, state, operation.receiver, statement.Range.StartLine)
		if operation.method == "select" {
			switch value.kind {
			case activeUIWorksheet:
				if !activeUIIdentityActive(value.parent, state.activeWorkbook) {
					findings = append(findings, a.activeUIFinding(file, proc, statement.Range.StartLine, "Worksheet.Select"))
				}
			case activeUIRange:
				if !activeUIIdentityActive(value.parent, state.activeWorksheet) {
					findings = append(findings, a.activeUIFinding(file, proc, statement.Range.StartLine, "Range.Select"))
				}
			}
		}
		return true
	})
	slices.SortStableFunc(findings, func(left, right Finding) int {
		if order := cmp.Compare(left.Line, right.Line); order != 0 {
			return order
		}
		return cmp.Compare(left.Message, right.Message)
	})
	return findings
}

func newActiveUIState() activeUIState {
	return activeUIState{bindings: make(map[string]activeUIValue)}
}

func (state activeUIState) clone() activeUIState {
	copyState := state
	copyState.bindings = maps.Clone(state.bindings)
	return copyState
}

func (state *activeUIState) invalidateActive() {
	if state == nil {
		return
	}
	state.activeWorkbook = ""
	state.activeWorksheet = ""
	clear(state.bindings)
}

func joinActiveUIState(left, right activeUIState) (activeUIState, bool) {
	joined := newActiveUIState()
	joined.activeWorkbook = commonActiveUIIdentity(left.activeWorkbook, right.activeWorkbook)
	joined.activeWorksheet = commonActiveUIIdentity(left.activeWorksheet, right.activeWorksheet)
	joined.resumeNext = left.resumeNext && right.resumeNext
	for name, value := range left.bindings {
		other, ok := right.bindings[name]
		if ok && value == other && value.kind != activeUIUnknown && value.id != "" {
			joined.bindings[name] = value
		}
	}
	changed := joined.activeWorkbook != left.activeWorkbook || joined.activeWorksheet != left.activeWorksheet || joined.resumeNext != left.resumeNext || len(joined.bindings) != len(left.bindings)
	if !changed {
		for name, value := range joined.bindings {
			if left.bindings[name] != value {
				changed = true
				break
			}
		}
	}
	return joined, changed
}

func commonActiveUIIdentity(left, right string) string {
	if left != "" && left == right {
		return left
	}
	return ""
}

func activeUICaptureValue(value activeUIValue, state activeUIState) activeUIValue {
	switch value.kind {
	case activeUIWorkbook:
		if value.id == activeUIActiveWorkbook {
			value.id = state.activeWorkbook
		}
	case activeUIWorksheet:
		if value.id == activeUIActiveWorksheet {
			value.id = state.activeWorksheet
		}
		if value.parent == activeUIActiveWorkbook {
			value.parent = state.activeWorkbook
		}
	case activeUIRange:
		if value.parent == activeUIActiveWorksheet {
			value.parent = state.activeWorksheet
		}
		if value.workbook == activeUIActiveWorkbook {
			value.workbook = state.activeWorkbook
		}
	}
	return value
}

func activeUIWithExpressions(statements map[int]procedureir.Statement) map[int]string {
	withByID := make(map[int]string)
	for id, statement := range statements {
		if statement.Kind != procedureir.StatementWith {
			continue
		}
		line := statement.Text
		if before, _, ok := strings.Cut(line, "\n"); ok {
			line = before
		}
		if match := activeUIWithExpressionRe.FindStringSubmatch(line); len(match) > 1 {
			withByID[id] = strings.TrimSpace(match[1])
		}
	}
	children := make(map[int][]procedureir.Statement)
	for _, statement := range statements {
		children[statement.ParentID] = append(children[statement.ParentID], statement)
	}
	for parent := range children {
		slices.SortStableFunc(children[parent], func(left, right procedureir.Statement) int {
			if left.Range.StartByte != right.Range.StartByte {
				return cmp.Compare(left.Range.StartByte, right.Range.StartByte)
			}
			return cmp.Compare(left.ID, right.ID)
		})
	}
	result := make(map[int]string)
	fullExpressions := make(map[int]string)
	nearestWith := func(parent int) string {
		for parent != 0 {
			if expression := withByID[parent]; expression != "" {
				return expression
			}
			statement, ok := statements[parent]
			if !ok {
				break
			}
			parent = statement.ParentID
		}
		return ""
	}
	for parent, siblings := range children {
		var previous procedureir.Statement
		for _, statement := range siblings {
			text := strings.TrimSpace(statement.Text)
			if !strings.HasPrefix(text, ".") {
				previous = statement
				continue
			}
			base := nearestWith(parent)
			if base == "" {
				if parentStatement, ok := statements[parent]; ok && activeUIChainText(parentStatement.Text) {
					base = strings.TrimSpace(parentStatement.Text)
					if fullExpressions[parentStatement.ID] != "" {
						base = fullExpressions[parentStatement.ID]
					}
				}
			}
			if previous.ID != 0 && activeUIChainText(previous.Text) {
				base = strings.TrimSpace(previous.Text)
				if fullExpressions[previous.ID] != "" {
					base = fullExpressions[previous.ID]
				}
			}
			if base != "" {
				result[statement.ID] = base
				fullExpressions[statement.ID] = base + text
			}
			previous = statement
		}
	}
	return result
}

func activeUIChainText(text string) bool {
	lower := strings.ToLower(strings.ReplaceAll(strings.TrimSpace(text), " ", ""))
	for _, member := range []string{".worksheets(", ".sheets(", ".range(", ".cells(", ".rows", ".columns", ".parent"} {
		if strings.Contains(lower, member) {
			return true
		}
	}
	return false
}

func activeUICallsByStatement(proc sourceProcedure) map[int][]procedureir.CallSite {
	result := make(map[int][]procedureir.CallSite)
	for call := range proc.Calls.All() {
		result[call.StatementID] = append(result[call.StatementID], call)
	}
	return result
}

func activeUIText(statement procedureir.Statement, withExpressions map[int]string) string {
	text := strings.TrimSpace(statement.Text)
	if expression := withExpressions[statement.ID]; expression != "" && strings.HasPrefix(text, ".") {
		text = expression + text
	}
	return strings.TrimSpace(text)
}

func parseActiveUIOperation(text string) (activeUIOperation, bool) {
	for _, method := range []string{"select", "activate"} {
		index, end, ok := activeUIFindLastMember(text, method)
		if !ok {
			continue
		}
		rawRemainder := text[end:]
		remainder := strings.TrimSpace(strings.ToLower(rawRemainder))
		if method == "activate" && remainder != "" && !strings.HasPrefix(remainder, "(") {
			continue
		}
		if method == "select" && strings.ContainsAny(rawRemainder, "\r\n") {
			continue
		}
		if method == "select" && strings.HasPrefix(remainder, "(") && !strings.HasSuffix(remainder, ")") {
			continue
		}
		receiver := strings.TrimSpace(text[:index])
		if strings.HasPrefix(strings.ToLower(receiver), "call ") {
			receiver = strings.TrimSpace(receiver[5:])
		}
		if receiver == "" || strings.Contains(strings.ToLower(receiver), " then ") {
			continue
		}
		selectReplaces := false
		selectModeSpecified := false
		if method == "select" {
			argumentText := strings.TrimSpace(text[end:])
			selectModeSpecified = argumentText != "" && argumentText != "()"
			selectReplaces = activeUISelectReplaces(argumentText)
		}
		return activeUIOperation{method: method, receiver: receiver, selectReplaces: selectReplaces, selectModeSpecified: selectModeSpecified}, true
	}
	return activeUIOperation{}, false
}

func activeUIFindLastMember(text, member string) (start, end int, ok bool) {
	bestStart := -1
	bestEnd := -1
	for offset := 0; offset < len(text); {
		candidateStart, candidateEnd, found := activeUIFindMemberFrom(text, member, offset)
		if !found {
			break
		}
		bestStart, bestEnd = candidateStart, candidateEnd
		offset = candidateEnd
	}
	if bestStart < 0 {
		return 0, 0, false
	}
	return bestStart, bestEnd, true
}

func activeUIFindMember(text, member string) (start, end int, ok bool) {
	return activeUIFindMemberFrom(text, member, 0)
}

func activeUIFindMemberFrom(text, member string, offset int) (start, end int, ok bool) {
	lowerMember := strings.ToLower(member)
	lowerText := strings.ToLower(text)
	trimmed := strings.TrimSpace(strings.ToLower(text))
	if strings.HasPrefix(trimmed, "rem") && (len(trimmed) == len("rem") || unicode.IsSpace(rune(trimmed[len("rem")]))) {
		return 0, 0, false
	}
	for index := offset; index < len(lowerText); index++ {
		if text[index] == '"' {
			index = activeUISkipStringLiteral(text, index) - 1
			continue
		}
		if text[index] == '\'' {
			break
		}
		if lowerText[index] != '.' {
			continue
		}
		memberStart := index + 1
		for memberStart < len(lowerText) && unicode.IsSpace(rune(lowerText[memberStart])) {
			memberStart++
		}
		if !strings.HasPrefix(lowerText[memberStart:], lowerMember) {
			continue
		}
		memberEnd := memberStart + len(lowerMember)
		if memberEnd < len(lowerText) {
			next := rune(lowerText[memberEnd])
			if unicode.IsLetter(next) || unicode.IsDigit(next) || next == '_' {
				continue
			}
		}
		return index, memberEnd, true
	}
	return 0, 0, false
}

func activeUICompact(text string) string {
	var compact strings.Builder
	compact.Grow(len(text))
	inString := false
	for index := 0; index < len(text); {
		r, size := utf8.DecodeRuneInString(text[index:])
		if r == '"' {
			compact.WriteRune(r)
			index += size
			if inString && index < len(text) && text[index] == '"' {
				compact.WriteByte('"')
				index++
				continue
			}
			inString = !inString
			continue
		}
		if !inString {
			if unicode.IsSpace(r) {
				index += size
				continue
			}
			r = unicode.ToLower(r)
		}
		compact.WriteRune(r)
		index += size
	}
	return compact.String()
}

func activeUISelectReplaces(argumentText string) bool {
	argumentText = strings.TrimSpace(argumentText)
	if argumentText == "" || argumentText == "()" {
		return true
	}
	if strings.HasPrefix(argumentText, "(") && strings.HasSuffix(argumentText, ")") {
		argumentText = strings.TrimSpace(argumentText[1 : len(argumentText)-1])
	}
	if index := strings.Index(argumentText, ","); index >= 0 {
		argumentText = strings.TrimSpace(argumentText[:index])
	}
	normalized := strings.ToLower(strings.ReplaceAll(argumentText, " ", ""))
	normalized = strings.TrimPrefix(normalized, "replace:=")
	return normalized == "true" || normalized == "vbtrue" || normalized == "1" || normalized == "-1"
}

func (a Analyzer) activeUIHasUnknownCall(file parsedFile, proc sourceProcedure, declarations declarationScope, state activeUIState, statementID int, calls map[int][]procedureir.CallSite, statement *procedureir.Statement, text string) bool {
	for _, call := range calls[statementID] {
		if !a.activeUICallModeled(file, proc, declarations, state, call) {
			return true
		}
	}
	if len(calls[statementID]) == 0 && a.activeUIHasUnmodeledMemberAccess(file, proc, declarations, state, statement, text) {
		return true
	}
	if statement != nil && statement.Kind == procedureir.StatementCall && len(calls[statementID]) == 0 {
		if _, ok := parseActiveUIOperation(text); !ok && !activeUIKnownStatementText(text) {
			return true
		}
	}
	return false
}

func (a Analyzer) activeUIHasUnmodeledMemberAccess(file parsedFile, proc sourceProcedure, declarations declarationScope, state activeUIState, statement *procedureir.Statement, text string) bool {
	if statement == nil || activeUIKnownStatementText(text) {
		return false
	}
	dot, memberStart, memberEnd, ok := activeUIFindAnyMember(text)
	if !ok {
		return false
	}
	member := strings.ToLower(strings.TrimSpace(text[memberStart:memberEnd]))
	if member == "select" || member == "activate" {
		return false
	}
	receiver := strings.TrimSpace(text[:dot])
	if receiver == "" {
		return false
	}
	if !activeUIIsModeledMember(member) {
		return true
	}
	receiverCopy := receiver
	call := procedureir.CallSite{
		Callee: procedureir.Callee{Receiver: &receiverCopy, Member: member},
		Range:  statement.Range,
	}
	return !a.activeUIKnownCallReceiver(file, proc, declarations, state, call)
}

func activeUIIsModeledMember(member string) bool {
	switch member {
	case "worksheets", "sheets", "range", "cells", "rows", "columns", "parent", "item", "count", "value", "value2", "address", "end", "resize":
		return true
	default:
		return false
	}
}

func activeUIFindAnyMember(text string) (dot, memberStart, memberEnd int, ok bool) {
	for index := 0; index < len(text); index++ {
		if text[index] == '"' {
			index = activeUISkipStringLiteral(text, index) - 1
			continue
		}
		if text[index] == '\'' {
			break
		}
		if text[index] != '.' {
			continue
		}
		start := index + 1
		for start < len(text) && unicode.IsSpace(rune(text[start])) {
			start++
		}
		if start >= len(text) || (!unicode.IsLetter(rune(text[start])) && text[start] != '_') {
			continue
		}
		end := start
		for end < len(text) {
			r := rune(text[end])
			if !unicode.IsLetter(r) && !unicode.IsDigit(r) && r != '_' {
				break
			}
			end++
		}
		if end > start {
			return index, start, end, true
		}
	}
	return 0, 0, 0, false
}

func (a Analyzer) activeUICallModeled(file parsedFile, proc sourceProcedure, declarations declarationScope, state activeUIState, call procedureir.CallSite) bool {
	member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
	if member != "" {
		switch member {
		case "activate", "select":
			return true
		case "worksheets", "sheets", "range", "cells", "rows", "columns", "parent", "item", "count", "value", "value2", "address", "end", "resize":
			return a.activeUIKnownCallReceiver(file, proc, declarations, state, call)
		case "print":
			return call.Callee.Receiver != nil && strings.EqualFold(strings.TrimSpace(*call.Callee.Receiver), "debug")
		default:
			return false
		}
	}
	name := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
	switch name {
	case "lcase", "ucase", "trim", "trim$", "cstr", "cint", "clng", "clnglng", "csng", "cdbl", "ccur", "cdec", "cdate", "val", "format", "typename", "isnumeric", "isdate", "len", "left", "left$", "right", "right$", "mid", "mid$", "replace", "split", "join", "array":
		return true
	default:
		return false
	}
}

func (a Analyzer) activeUIKnownCallReceiver(file parsedFile, proc sourceProcedure, declarations declarationScope, state activeUIState, call procedureir.CallSite) bool {
	if call.Callee.Receiver == nil || strings.TrimSpace(*call.Callee.Receiver) == "" {
		name := strings.ToLower(strings.TrimSpace(call.Callee.BaseName))
		switch name {
		case "worksheets", "sheets", "range", "cells", "rows", "columns":
			return !activeUIRootShadowed(file, declarations, name)
		default:
			return false
		}
	}
	receiver := strings.TrimSpace(*call.Callee.Receiver)
	if activeUIIsApplicationRoot(receiver) {
		return true
	}
	if value := a.resolveActiveUIValue(file, proc, declarations, state, receiver, call.Range.StartLine); value.kind != activeUIUnknown {
		return true
	}
	if declaration, ok := declarations.lookup(receiver); ok {
		return activeUIKnownExcelType(declaration.Type)
	}
	line := call.Range.StartLine
	if line <= 0 {
		line = 1
	}
	typ, resolved := resolveExcelExpressionType(file, a.typeDB, receiver, line-1, a.RootDir, a.Config)
	return resolved && activeUIKnownExcelType(typ)
}

func activeUIKnownStatementText(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return strings.HasPrefix(lower, "debug.print") || strings.HasPrefix(lower, "debug.")
}

func (a Analyzer) activeUITransfer(file parsedFile, proc sourceProcedure, declarations declarationScope, withExpressions map[int]string, calls map[int][]procedureir.CallSite, statement *procedureir.Statement, state activeUIState) activeUIState {
	if statement == nil {
		return state
	}
	text := activeUIText(*statement, withExpressions)
	if statement.Kind == procedureir.StatementOnError {
		if isOnErrorResumeNext(*statement) {
			state.resumeNext = true
		} else if restoresErrorHandling(*statement) {
			state.resumeNext = false
		}
		return state
	}
	if a.activeUIHasUnknownCall(file, proc, declarations, state, statement.ID, calls, statement, text) {
		state.invalidateActive()
	}
	if match := activeUISetAssignmentRe.FindStringSubmatch(text); len(match) > 2 {
		name := strings.ToLower(strings.TrimSpace(match[1]))
		value := a.resolveActiveUIValue(file, proc, declarations, state, match[2], statement.Range.StartLine)
		if value.kind == activeUIUnknown {
			if declaration, ok := declarations.lookup(name); ok {
				value.kind = activeUIValueKindForType(declaration.Type)
			}
		}
		value = activeUICaptureValue(value, state)
		if value.kind == activeUIUnknown || value.id == "" && value.parent == "" {
			delete(state.bindings, name)
		} else {
			state.bindings[name] = value
		}
		return state
	}
	operation, ok := parseActiveUIOperation(text)
	if !ok || a.activeUIHasUnknownCall(file, proc, declarations, state, statement.ID, calls, statement, text) {
		return state
	}
	value := a.resolveActiveUIValue(file, proc, declarations, state, operation.receiver, statement.Range.StartLine)
	switch operation.method {
	case "activate":
		if state.resumeNext {
			state.invalidateActive()
			return state
		}
		switch value.kind {
		case activeUIWorkbook:
			state.activeWorkbook = value.id
			state.activeWorksheet = ""
		case activeUIWorksheet:
			state.activeWorkbook = value.parent
			state.activeWorksheet = value.id
		default:
			state.invalidateActive()
		}
	case "select":
		if state.resumeNext {
			state.invalidateActive()
			return state
		}
		switch value.kind {
		case activeUIWorksheet:
			if activeUIIdentityActive(value.parent, state.activeWorkbook) {
				if !operation.selectModeSpecified || operation.selectReplaces {
					if value.parent != "" {
						state.activeWorkbook = value.parent
					}
					state.activeWorksheet = value.id
				}
			}
		case activeUIRange:
			if activeUIIdentityActive(value.parent, state.activeWorksheet) {
				state.activeWorksheet = value.parent
			}
		default:
			state.invalidateActive()
		}
	}
	return state
}

func (a Analyzer) activeUIRefineBranch(file parsedFile, proc sourceProcedure, declarations declarationScope, statement *procedureir.Statement, edge cfg.Edge, state activeUIState) activeUIState {
	if statement == nil || edge.Kind != cfg.EdgeBranchTrue {
		return state
	}
	condition := ""
	if statement.Condition != nil {
		condition = statement.Condition.Text
	}
	if condition == "" {
		condition = statement.Text
		if index := strings.Index(strings.ToLower(condition), " then "); index >= 0 {
			condition = condition[:index]
		}
		condition = strings.TrimSpace(strings.TrimPrefix(strings.TrimPrefix(condition, "If "), "if "))
	}
	match := activeUIIsComparisonRe.FindStringSubmatch(strings.TrimSpace(condition))
	if len(match) < 4 {
		return state
	}
	left := strings.TrimSpace(match[1])
	right := strings.TrimSpace(match[3])
	activeKind := activeUIComparisonKind(right)
	valueExpression := left
	if activeKind == "" {
		activeKind = activeUIComparisonKind(left)
		valueExpression = right
	}
	if activeKind == "" || match[2] != "" {
		return state
	}
	value := a.resolveActiveUIValue(file, proc, declarations, state, valueExpression, statement.Range.StartLine)
	if activeKind == "activeworkbook" && value.kind == activeUIWorkbook && value.id != "" {
		state.activeWorkbook = value.id
	}
	if activeKind == "activesheet" && value.kind == activeUIWorksheet && value.id != "" {
		state.activeWorksheet = value.id
		if value.parent != "" {
			state.activeWorkbook = value.parent
		}
	}
	return state
}

func activeUIComparisonKind(expression string) string {
	switch activeUICompact(expression) {
	case "activeworkbook", "application.activeworkbook":
		return "activeworkbook"
	case "activesheet", "application.activesheet":
		return "activesheet"
	default:
		return ""
	}
}

func activeUIIdentityActive(expected, actual string) bool {
	if expected == "" {
		return false
	}
	if expected == activeUIActiveWorkbook || expected == activeUIActiveWorksheet {
		return true
	}
	return actual != "" && expected == actual
}

func (a Analyzer) resolveActiveUIValue(file parsedFile, proc sourceProcedure, declarations declarationScope, state activeUIState, expression string, line int) activeUIValue {
	expression = trimOuterVBAParentheses(strings.TrimSpace(expression))
	if expression == "" {
		return activeUIValue{}
	}
	if strings.HasPrefix(expression, ".") {
		return activeUIValue{}
	}
	if value, ok := state.bindings[strings.ToLower(expression)]; ok {
		return value
	}
	lower := activeUICompact(expression)
	switch lower {
	case "thisworkbook":
		return activeUIValue{kind: activeUIWorkbook, id: "workbook:thisworkbook"}
	case "activeworkbook", "application.activeworkbook":
		return activeUIValue{kind: activeUIWorkbook, id: activeUIActiveWorkbook}
	case "activesheet", "application.activesheet":
		return activeUIValue{kind: activeUIWorksheet, id: activeUIActiveWorksheet, parent: activeUIActiveWorkbook}
	}
	if parentStart, parentEnd, ok := activeUIFindLastMember(expression, "parent"); ok && strings.TrimSpace(expression[parentEnd:]) == "" {
		base := a.resolveActiveUIValue(file, proc, declarations, state, expression[:parentStart], line)
		switch base.kind {
		case activeUIWorksheet:
			return activeUIValue{kind: activeUIWorkbook, id: base.parent}
		case activeUIRange:
			return activeUIValue{kind: activeUIWorksheet, id: base.parent, parent: base.workbook}
		}
	}
	if value, ok := a.resolveActiveUICollectionValue(file, proc, declarations, state, expression, line); ok {
		return value
	}
	if root := activeUICollectionRootName(lower); root != "" && activeUIRootShadowed(file, declarations, root) {
		return activeUIValue{}
	}
	if declaration, ok := declarations.lookup(expression); ok {
		return activeUIValue{kind: activeUIValueKindForType(declaration.Type)}
	}
	typ, resolved := resolveExcelExpressionType(file, a.typeDB, expression, line-1, a.RootDir, a.Config)
	if resolved {
		return activeUIValue{kind: activeUIValueKindForType(typ)}
	}
	return activeUIValue{}
}

func (a Analyzer) resolveActiveUICollectionValue(file parsedFile, proc sourceProcedure, declarations declarationScope, state activeUIState, expression string, line int) (activeUIValue, bool) {
	lower := activeUICompact(expression)
	for _, member := range []string{"worksheets", "sheets"} {
		if strings.HasPrefix(lower, member+"(") {
			if activeUIRootShadowed(file, declarations, member) {
				return activeUIValue{}, false
			}
			selector, stable := activeUICollectionIdentity(expression, len(member), 0)
			id := ""
			if stable {
				id = activeUIChildIdentity("worksheet", activeUIActiveWorkbook, selector)
			}
			return activeUIValue{kind: activeUIWorksheet, id: id, parent: activeUIActiveWorkbook}, true
		}
	}
	for _, member := range []string{"range", "cells", "rows", "columns"} {
		if strings.HasPrefix(lower, member+"(") {
			if activeUIRootShadowed(file, declarations, member) {
				return activeUIValue{}, false
			}
			selector, stable := activeUICollectionIdentity(expression, len(member), 0)
			id := ""
			if stable {
				id = activeUIChildIdentity("range", activeUIActiveWorksheet, selector)
			}
			return activeUIValue{kind: activeUIRange, id: id, parent: activeUIActiveWorksheet, workbook: activeUIActiveWorkbook}, true
		}
	}
	member, index, end, ok := activeUILastCollectionMember(expression)
	if !ok {
		return activeUIValue{}, false
	}
	ownerExpression := strings.TrimSpace(expression[:index])
	owner := a.resolveActiveUIValue(file, proc, declarations, state, ownerExpression, line)
	applicationRoot := activeUIIsApplicationRoot(ownerExpression)
	ownerKnownExcel := owner.kind != activeUIUnknown || applicationRoot
	if !ownerKnownExcel {
		ownerKnownExcel = a.activeUIOwnerCanBeExcel(file, proc, declarations, ownerExpression, line)
	}
	if !ownerKnownExcel {
		return activeUIValue{}, false
	}
	compactMember := activeUICompact(member)
	switch compactMember {
	case "worksheets", "sheets":
		if !applicationRoot && owner.kind != activeUIWorkbook {
			return activeUIValue{}, false
		}
		ownerID := owner.id
		if applicationRoot {
			ownerID = activeUIActiveWorkbook
		}
		if ownerID == "" {
			ownerID = activeUIExpressionIdentity(ownerExpression)
		}
		parent := owner.id
		if applicationRoot {
			parent = activeUIActiveWorkbook
		}
		selector, stable := activeUICollectionIdentity(expression, end, index)
		id := ""
		if stable {
			id = activeUIChildIdentity("worksheet", ownerID, selector)
		}
		return activeUIValue{kind: activeUIWorksheet, id: id, parent: parent}, true
	case "range", "cells", "rows", "columns":
		if !applicationRoot && owner.kind == activeUIWorkbook {
			return activeUIValue{}, false
		}
		ownerID := owner.id
		if applicationRoot {
			ownerID = activeUIActiveWorksheet
		}
		if ownerID == "" {
			ownerID = activeUIExpressionIdentity(ownerExpression)
		}
		parent := owner.id
		workbook := owner.workbook
		if applicationRoot {
			parent = activeUIActiveWorksheet
			workbook = activeUIActiveWorkbook
		} else if parent == "" {
			parent = "worksheet:owner:" + ownerID
		}
		if owner.kind == activeUIWorksheet {
			workbook = owner.parent
		}
		selector, stable := activeUICollectionIdentity(expression, end, index)
		id := ""
		if stable {
			id = activeUIChildIdentity("range", ownerID, selector)
		}
		return activeUIValue{kind: activeUIRange, id: id, parent: parent, workbook: workbook}, true
	}
	return activeUIValue{}, false
}

func activeUICollectionRootName(lowerExpression string) string {
	for _, member := range []string{"worksheets", "sheets", "range", "cells", "rows", "columns"} {
		if strings.HasPrefix(lowerExpression, member+"(") {
			return member
		}
	}
	return ""
}

func activeUICollectionIdentity(expression string, memberEnd, memberStart int) (string, bool) {
	if !activeUICollectionArgumentsStatic(expression, memberEnd) {
		return "", false
	}
	return activeUICompact(expression[memberStart:]), true
}

func activeUICollectionArgumentsStatic(expression string, offset int) bool {
	found := false
	for index := offset; index < len(expression); index++ {
		if expression[index] == '"' {
			index = activeUISkipStringLiteral(expression, index) - 1
			continue
		}
		if expression[index] != '(' {
			continue
		}
		end := activeUIMatchingParen(expression, index)
		if end < 0 || !activeUIStaticSelectorArguments(expression[index+1:end]) {
			return false
		}
		found = true
		index = end
	}
	return found
}

func activeUIMatchingParen(expression string, start int) int {
	depth := 0
	for index := start; index < len(expression); index++ {
		if expression[index] == '"' {
			index = activeUISkipStringLiteral(expression, index) - 1
			continue
		}
		switch expression[index] {
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

func activeUISkipStringLiteral(expression string, start int) int {
	for index := start + 1; index < len(expression); index++ {
		if expression[index] != '"' {
			continue
		}
		if index+1 < len(expression) && expression[index+1] == '"' {
			index++
			continue
		}
		return index + 1
	}
	return len(expression)
}

func activeUIStaticSelectorArguments(arguments string) bool {
	if strings.TrimSpace(arguments) == "" {
		return false
	}
	for index := 0; index < len(arguments); index++ {
		if arguments[index] == '"' {
			index = activeUISkipStringLiteral(arguments, index) - 1
			continue
		}
		r := rune(arguments[index])
		if unicode.IsLetter(r) || r == '_' {
			return false
		}
	}
	return true
}

func activeUIRootShadowed(file parsedFile, declarations declarationScope, name string) bool {
	if declarations.shadowsModule(name) {
		return true
	}
	for _, procedure := range file.procedureProjection() {
		if procedure.IR != nil && strings.EqualFold(procedure.IR.Symbol.Name, name) {
			return true
		}
	}
	return false
}

func activeUILastCollectionMember(expression string) (member string, start, end int, ok bool) {
	bestStart := -1
	bestEnd := -1
	for _, candidate := range []string{"worksheets", "sheets", "range", "cells", "rows", "columns"} {
		for offset := 0; offset < len(expression); {
			candidateStart, candidateEnd, found := activeUIFindMemberFrom(expression, candidate, offset)
			if !found {
				break
			}
			if candidateStart > bestStart {
				member = candidate
				bestStart, bestEnd = candidateStart, candidateEnd
			}
			offset = candidateEnd
		}
	}
	if bestStart < 0 {
		return "", 0, 0, false
	}
	return member, bestStart, bestEnd, true
}

func activeUIIsApplicationRoot(expression string) bool {
	return activeUICompact(expression) == "application"
}

func (a Analyzer) activeUIOwnerCanBeExcel(file parsedFile, proc sourceProcedure, declarations declarationScope, expression string, line int) bool {
	if activeUIIsApplicationRoot(expression) {
		return true
	}
	if declaration, ok := declarations.lookup(expression); ok {
		return activeUIKnownExcelType(declaration.Type)
	}
	typ, resolved := resolveExcelExpressionType(file, a.typeDB, expression, line-1, a.RootDir, a.Config)
	return resolved && activeUIKnownExcelType(typ)
}

func activeUIExpressionIdentity(expression string) string {
	return "expr:" + activeUICompact(expression)
}

func activeUIChildIdentity(kind, parent, selector string) string {
	if parent == "" {
		parent = "unknown"
	}
	return kind + ":" + parent + ":" + selector
}

func activeUIValueKindForType(typ string) activeUIValueKind {
	lower, ok := activeUIExcelTypeBase(typ)
	if !ok {
		return activeUIUnknown
	}
	switch lower {
	case "workbook", "_workbook":
		return activeUIWorkbook
	case "worksheet", "worksheets", "_worksheet", "_worksheets":
		return activeUIWorksheet
	case "range", "_range":
		return activeUIRange
	default:
		return activeUIUnknown
	}
}

func activeUIKnownExcelType(typ string) bool {
	base, ok := activeUIExcelTypeBase(typ)
	if !ok {
		return false
	}
	switch base {
	case "application", "_application", "workbook", "_workbook", "worksheet", "_worksheet", "worksheets", "_worksheets", "range", "_range", "listrow", "_listrow", "listcolumn", "_listcolumn", "listobject", "_listobject", "window", "_window", "chart", "_chart", "shape", "_shape":
		return true
	default:
		return false
	}
}

func activeUIExcelTypeBase(typ string) (string, bool) {
	compact := activeUICompact(typ)
	if compact == "" {
		return "", false
	}
	if dot := strings.LastIndexByte(compact, '.'); dot >= 0 {
		if compact[:dot] != "excel" {
			return "", false
		}
		compact = compact[dot+1:]
	}
	return compact, true
}

func (a Analyzer) activeUIFinding(file parsedFile, proc sourceProcedure, line int, operation string) Finding {
	if operation == "Worksheet.Select" {
		return a.simpleFinding(file, proc, line, "VBA250", "warning", operation+" may fail because its parent workbook is not known to be active.", "Worksheet.Select requires the target worksheet's parent workbook to be active on every reachable path.", "Activate the target workbook on every path before selecting the worksheet, or use an explicit non-UI operation.")
	}
	return a.simpleFinding(file, proc, line, "VBA250", "warning", operation+" may fail because its worksheet is not known to be active.", "Range.Select requires the target range's worksheet to be active on every reachable path.", "Activate the target worksheet on every path before selecting the range, or use an explicit non-UI operation.")
}
