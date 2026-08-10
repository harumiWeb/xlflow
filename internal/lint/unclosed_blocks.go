package lint

import (
	"sort"
	"strings"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/gui"
)

// unmatchedBlockCandidates performs a deliberately conservative lexical pass
// after tree-sitter has already reported recovery. It is not a second VBA
// parser: any conditional compilation or incompatible block structure makes
// the result unreliable and leaves VB014 as the generic recovery diagnostic.
func unmatchedBlockCandidates(source string) ([]unclosedBlockCandidate, bool) {
	statements, reliable := blockStatements(source)
	if !reliable {
		return nil, false
	}

	stack := make([]openBlock, 0)
	candidates := make([]unclosedBlockCandidate, 0)
	inTypeDeclaration := false
	for i := 0; i < len(statements); i++ {
		statement := statements[i]
		text := strings.TrimSpace(statement.text)
		if text == "" {
			continue
		}
		if isRemComment(text) {
			for i+1 < len(statements) && statements[i+1].group == statement.group {
				i++
			}
			continue
		}
		lower := strings.ToLower(text)
		anchor := blockAnchor{line: statement.line, column: statement.column}

		if isTypeDeclarationEnd(lower) {
			inTypeDeclaration = false
			continue
		}
		if isTypeDeclarationStart(lower) {
			inTypeDeclaration = true
			continue
		}
		if inTypeDeclaration {
			continue
		}

		if isProcedureEndStatement(lower) {
			candidates = appendUnclosedBlockCandidates(candidates, stack, anchor)
			stack = stack[:0]
			continue
		}
		if _, ok := procedureStart(text, statement.line); ok {
			if len(stack) > 0 {
				return nil, false
			}
			continue
		}
		if isIfBranchStatement(lower) {
			if len(stack) == 0 || stack[len(stack)-1].kind != "if" {
				return nil, false
			}
			if lower == "else" {
				if stack[len(stack)-1].elseSeen {
					return nil, false
				}
				stack[len(stack)-1].elseSeen = true
			} else if stack[len(stack)-1].elseSeen {
				return nil, false
			}
			continue
		}
		if isSelectCaseBranch(lower) {
			if len(stack) > 0 && stack[len(stack)-1].kind != "select" {
				return nil, false
			}
			continue
		}

		if closer, count, ok := blockCloser(lower); ok {
			var matched bool
			stack, candidates, matched = closeBlocks(stack, candidates, closer, count, anchor)
			if !matched {
				return nil, false
			}
			continue
		}
		if hasVBAKeywordPrefix(lower, "next") || strings.HasPrefix(lower, "loop while") || strings.HasPrefix(lower, "loop until") {
			return nil, false
		}

		if block, ok := blockOpener(lower, i+1 < len(statements) && statements[i+1].group == statement.group); ok {
			block.line = statement.line
			block.column = statement.column
			stack = append(stack, block)
		}
	}

	if len(stack) > 0 {
		lines := normalizedSourceLines(source)
		candidates = appendUnclosedBlockCandidates(candidates, stack, blockAnchor{line: len(lines), column: 1})
	}
	return candidates, true
}

// shouldReportStructuralParseIssue preserves VB014 when the CST intentionally
// accepts block fragments, such as conditional-compilation-split If blocks.
// The detailed scanner remains authoritative when it is reliable; ambiguous
// conditional compilation is checked across every possible branch depth.
func shouldReportStructuralParseIssue(source string) bool {
	blocks, reliable := unmatchedBlockCandidates(source)
	if reliable {
		return len(blocks) > 0
	}
	if hasConditionalCompilation(source) {
		return conditionalIfBalanceInvalid(source)
	}
	return hasBlockBoundarySyntax(source)
}

func hasConditionalCompilation(source string) bool {
	for _, line := range normalizedSourceLines(source) {
		if isConditionalCompilationDirective(strings.ToLower(strings.TrimSpace(gui.StripComment(line)))) {
			return true
		}
	}
	return false
}

func hasBlockBoundarySyntax(source string) bool {
	for _, line := range normalizedSourceLines(source) {
		lower := strings.ToLower(strings.TrimSpace(gui.StripComment(line)))
		if _, ok := blockOpener(lower, false); ok {
			return true
		}
		if _, _, ok := blockCloser(lower); ok {
			return true
		}
		if isIfBranchStatement(lower) || isSelectCaseBranch(lower) {
			return true
		}
	}
	return false
}

func isTypeDeclarationStart(lower string) bool {
	fields := strings.Fields(lower)
	if len(fields) < 2 || fields[len(fields)-2] != "type" {
		return len(fields) == 1 && fields[0] == "type"
	}
	for _, field := range fields[:len(fields)-2] {
		switch field {
		case "private", "public", "friend":
		default:
			return false
		}
	}
	return true
}

func isTypeDeclarationEnd(lower string) bool {
	return lower == "end type"
}

type conditionalIfFrame struct {
	baseline conditionalStateSet
	branches conditionalStateSet
	sawElse  bool
}

type conditionalBalanceState struct {
	runtime   string
	decisions map[string]bool
}

type conditionalStateSet map[string]conditionalBalanceState

func conditionalIfBalanceInvalid(source string) bool {
	statements, reliable := conditionalBlockStatements(source)
	if !reliable {
		return true
	}

	trackedConditions := repeatedConditionalConditions(statements)
	states := conditionalStateSet{}
	addConditionalState(states, conditionalBalanceState{decisions: map[string]bool{}})
	frames := make([]conditionalIfFrame, 0)
	for i, statement := range statements {
		lower := strings.ToLower(strings.TrimSpace(statement.text))
		switch {
		case strings.HasPrefix(lower, "#if ") && strings.HasSuffix(lower, " then"):
			condition := conditionalCompilationCondition(lower, "#if ")
			frame := conditionalIfFrame{branches: make(conditionalStateSet)}
			if trackedConditions[condition] {
				states, frame.baseline = splitConditionalIfStates(states, condition)
			} else {
				frame.baseline = cloneConditionalIfStates(states)
			}
			frames = append(frames, frame)
		case strings.HasPrefix(lower, "#elseif ") && strings.HasSuffix(lower, " then"):
			if len(frames) == 0 {
				return true
			}
			frame := &frames[len(frames)-1]
			if frame.sawElse {
				return true
			}
			mergeConditionalIfStates(frame.branches, states)
			condition := conditionalCompilationCondition(lower, "#elseif ")
			if trackedConditions[condition] {
				states, frame.baseline = splitConditionalIfStates(frame.baseline, condition)
			} else {
				states = cloneConditionalIfStates(frame.baseline)
			}
		case lower == "#else":
			if len(frames) == 0 {
				return true
			}
			frame := &frames[len(frames)-1]
			if frame.sawElse {
				return true
			}
			mergeConditionalIfStates(frame.branches, states)
			frame.sawElse = true
			states = cloneConditionalIfStates(frame.baseline)
		case lower == "#end if":
			if len(frames) == 0 {
				return true
			}
			frame := frames[len(frames)-1]
			frames = frames[:len(frames)-1]
			mergeConditionalIfStates(frame.branches, states)
			if !frame.sawElse {
				mergeConditionalIfStates(frame.branches, frame.baseline)
			}
			states = frame.branches
		case isIfBranchStatement(lower):
			var valid bool
			states, valid = applyConditionalIfBranch(states, lower == "else")
			if !valid {
				return true
			}
		case lower == "end if":
			var valid bool
			states, valid = closeConditionalIfStates(states)
			if !valid {
				return true
			}
		default:
			block, ok := blockOpener(lower, i+1 < len(statements) && statements[i+1].group == statement.group)
			if ok && block.kind == "if" {
				states = openConditionalIfStates(states)
			}
		}
	}
	if len(frames) != 0 {
		return true
	}
	for _, state := range states {
		if state.runtime != "" {
			return true
		}
	}
	return false
}

func cloneConditionalIfStates(source conditionalStateSet) conditionalStateSet {
	cloned := make(conditionalStateSet, len(source))
	for key, state := range source {
		cloned[key] = cloneConditionalBalanceState(state)
	}
	return cloned
}

func mergeConditionalIfStates(target, source conditionalStateSet) {
	for key, state := range source {
		target[key] = cloneConditionalBalanceState(state)
	}
}

func openConditionalIfStates(source conditionalStateSet) conditionalStateSet {
	opened := make(conditionalStateSet, len(source))
	for _, state := range source {
		state.runtime += "0"
		addConditionalState(opened, state)
	}
	return opened
}

func applyConditionalIfBranch(source conditionalStateSet, isElse bool) (conditionalStateSet, bool) {
	branched := make(conditionalStateSet, len(source))
	for _, state := range source {
		if state.runtime == "" || state.runtime[len(state.runtime)-1] == '1' {
			return nil, false
		}
		if isElse {
			state.runtime = state.runtime[:len(state.runtime)-1] + "1"
		}
		addConditionalState(branched, state)
	}
	return branched, true
}

func closeConditionalIfStates(source conditionalStateSet) (conditionalStateSet, bool) {
	closed := make(conditionalStateSet, len(source))
	for _, state := range source {
		if state.runtime == "" {
			return nil, false
		}
		state.runtime = state.runtime[:len(state.runtime)-1]
		addConditionalState(closed, state)
	}
	return closed, true
}

func repeatedConditionalConditions(statements []blockStatement) map[string]bool {
	counts := make(map[string]int)
	for _, statement := range statements {
		lower := strings.ToLower(strings.TrimSpace(statement.text))
		for _, prefix := range []string{"#if ", "#elseif "} {
			if strings.HasPrefix(lower, prefix) && strings.HasSuffix(lower, " then") {
				counts[conditionalCompilationCondition(lower, prefix)]++
				break
			}
		}
	}
	repeated := make(map[string]bool)
	for condition, count := range counts {
		if count > 1 {
			repeated[condition] = true
		}
	}
	return repeated
}

func conditionalCompilationCondition(lower, prefix string) string {
	condition := strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(lower, prefix), "then"))
	return strings.Join(strings.Fields(condition), " ")
}

func splitConditionalIfStates(source conditionalStateSet, condition string) (conditionalStateSet, conditionalStateSet) {
	active := make(conditionalStateSet)
	baseline := make(conditionalStateSet)
	for _, state := range source {
		if value, ok := state.decisions[condition]; ok {
			if value {
				addConditionalState(active, state)
			} else {
				addConditionalState(baseline, state)
			}
			continue
		}
		trueState := cloneConditionalBalanceState(state)
		trueState.decisions[condition] = true
		addConditionalState(active, trueState)
		falseState := cloneConditionalBalanceState(state)
		falseState.decisions[condition] = false
		addConditionalState(baseline, falseState)
	}
	return active, baseline
}

func cloneConditionalBalanceState(source conditionalBalanceState) conditionalBalanceState {
	decisions := make(map[string]bool, len(source.decisions))
	for condition, value := range source.decisions {
		decisions[condition] = value
	}
	source.decisions = decisions
	return source
}

func addConditionalState(target conditionalStateSet, state conditionalBalanceState) {
	target[conditionalStateKey(state)] = cloneConditionalBalanceState(state)
}

func conditionalStateKey(state conditionalBalanceState) string {
	conditions := make([]string, 0, len(state.decisions))
	for condition, value := range state.decisions {
		marker := "0"
		if value {
			marker = "1"
		}
		conditions = append(conditions, condition+"="+marker)
	}
	sort.Strings(conditions)
	return state.runtime + "\x1f" + strings.Join(conditions, ";")
}

type blockStatement struct {
	text   string
	line   int
	column int
	group  int
}

// blockStatements joins valid VBA continuation lines before splitting colon
// statements. Conditional compilation is intentionally excluded because the
// active branch cannot be known from exported source alone.
func blockStatements(source string) ([]blockStatement, bool) {
	return scanBlockStatements(source, false)
}

func conditionalBlockStatements(source string) ([]blockStatement, bool) {
	return scanBlockStatements(source, true)
}

func scanBlockStatements(source string, allowConditionalCompilation bool) ([]blockStatement, bool) {
	lines := normalizedSourceLines(source)
	statements := make([]blockStatement, 0)
	var logical strings.Builder
	positions := make([]blockAnchor, 0)
	appendPhysical := func(text string, line int) {
		logical.WriteString(text)
		for column := 1; column <= len(text); column++ {
			positions = append(positions, blockAnchor{line: line, column: column})
		}
	}
	appendSyntheticSpace := func(line, column int) {
		logical.WriteByte(' ')
		positions = append(positions, blockAnchor{line: line, column: column})
	}
	resetLogical := func() {
		logical.Reset()
		positions = positions[:0]
	}
	pending := false
	group := 0
	for i, line := range lines {
		code := gui.StripComment(line)
		trimmed := strings.TrimSpace(code)
		if isConditionalCompilationDirective(strings.ToLower(trimmed)) {
			if !allowConditionalCompilation || pending {
				return nil, false
			}
			group++
			statements = append(statements, blockStatement{
				text:   trimmed,
				line:   i + 1,
				column: len(code) - len(strings.TrimLeft(code, " \t")) + 1,
				group:  group,
			})
			continue
		}
		if !pending && trimmed == "" {
			continue
		}
		if hasValidLineContinuation(code) {
			continued := removeLineContinuationMarker(code)
			appendPhysical(continued, i+1)
			appendSyntheticSpace(i+1, len(continued)+1)
			pending = true
			continue
		}
		appendPhysical(code, i+1)
		parts := splitStatementsWithColumns(logical.String())
		group++
		for _, part := range parts {
			if part.start < 0 || part.start >= len(positions) {
				return nil, false
			}
			position := positions[part.start]
			statements = append(statements, blockStatement{
				text:   part.text,
				line:   position.line,
				column: position.column,
				group:  group,
			})
		}
		resetLogical()
		pending = false
	}
	if pending {
		return nil, false
	}
	return statements, true
}

type openBlock struct {
	kind     string
	label    string
	closer   string
	line     int
	column   int
	elseSeen bool
}

type blockAnchor struct {
	line   int
	column int
}

type unclosedBlockCandidate struct {
	kind           string
	label          string
	expectedCloser string
	openingLine    int
	openingColumn  int
	expectedLine   int
	expectedColumn int
}

func appendUnclosedBlockCandidates(candidates []unclosedBlockCandidate, blocks []openBlock, anchor blockAnchor) []unclosedBlockCandidate {
	for _, block := range blocks {
		candidates = append(candidates, unclosedBlockCandidate{
			kind:           block.kind,
			label:          block.label,
			expectedCloser: block.closer,
			openingLine:    block.line,
			openingColumn:  block.column,
			expectedLine:   anchor.line,
			expectedColumn: anchor.column,
		})
	}
	return candidates
}

func closeBlocks(stack []openBlock, candidates []unclosedBlockCandidate, kind string, count int, anchor blockAnchor) ([]openBlock, []unclosedBlockCandidate, bool) {
	if count > 1 {
		for range count {
			if len(stack) == 0 || stack[len(stack)-1].kind != kind {
				return stack, candidates, false
			}
			stack = stack[:len(stack)-1]
		}
		return stack, candidates, true
	}
	if nested := indentedParentCloser(stack, kind, anchor); nested >= 0 {
		candidates = appendUnclosedBlockCandidates(candidates, stack[nested+1:], anchor)
		stack = stack[:nested]
		return stack, candidates, true
	}
	for range count {
		if len(stack) == 0 || stack[len(stack)-1].kind != kind {
			return stack, candidates, false
		}
		stack = stack[:len(stack)-1]
	}
	return stack, candidates, true
}

// indentedParentCloser resolves the common recovery shape where a closer is
// aligned with an outer block while one or more nested blocks remain open.
// Indentation is advisory in VBA, so require exact parent alignment and every
// skipped nested opener to be indented further; otherwise retain generic
// parser-recovery guidance.
func indentedParentCloser(stack []openBlock, kind string, anchor blockAnchor) int {
	if len(stack) < 2 || stack[len(stack)-1].column <= anchor.column {
		return -1
	}
	for i := len(stack) - 2; i >= 0; i-- {
		if stack[i].kind != kind || stack[i].column != anchor.column {
			continue
		}
		for _, nested := range stack[i+1:] {
			if nested.column <= anchor.column {
				return -1
			}
		}
		return i
	}
	return -1
}

func blockOpener(lower string, hasFollowingStatementOnLine bool) (openBlock, bool) {
	switch {
	case strings.HasPrefix(lower, "if ") && strings.HasSuffix(lower, " then"):
		condition := strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(lower, "if "), " then"))
		if hasFollowingStatementOnLine || condition == "" {
			return openBlock{}, false
		}
		return openBlock{kind: "if", label: "multiline If", closer: "End If"}, true
	case validForEachOpener(lower) || validForOpener(lower):
		return openBlock{kind: "for", label: "For", closer: "Next"}, true
	case lower == "do" || validDoConditionOpener(lower, "do while ") || validDoConditionOpener(lower, "do until "):
		return openBlock{kind: "do", label: "Do", closer: "Loop"}, true
	case hasNonEmptySuffix(lower, "while "):
		return openBlock{kind: "while", label: "While", closer: "Wend"}, true
	case hasNonEmptySuffix(lower, "with "):
		return openBlock{kind: "with", label: "With", closer: "End With"}, true
	case hasNonEmptySuffix(lower, "select case "):
		return openBlock{kind: "select", label: "Select Case", closer: "End Select"}, true
	default:
		return openBlock{}, false
	}
}

func blockCloser(lower string) (kind string, count int, ok bool) {
	switch {
	case lower == "end if":
		return "if", 1, true
	case lower == "end with":
		return "with", 1, true
	case lower == "end select":
		return "select", 1, true
	case lower == "wend":
		return "while", 1, true
	case lower == "loop" || validDoConditionOpener(lower, "loop while ") || validDoConditionOpener(lower, "loop until "):
		return "do", 1, true
	case lower == "next":
		return "for", 1, true
	case strings.HasPrefix(lower, "next "):
		names := strings.Split(strings.TrimSpace(strings.TrimPrefix(lower, "next ")), ",")
		for _, name := range names {
			if !isVBAControlVariable(name) {
				return "", 0, false
			}
		}
		return "for", len(names), true
	default:
		return "", 0, false
	}
}

func validForEachOpener(lower string) bool {
	fields := strings.Fields(lower)
	return len(fields) >= 4 && fields[0] == "for" && fields[1] == "each" && isVBAControlVariable(fields[2]) && fields[3] == "in" && strings.TrimSpace(strings.Join(fields[4:], " ")) != ""
}

func validForOpener(lower string) bool {
	if !strings.HasPrefix(lower, "for ") || strings.HasPrefix(lower, "for each ") {
		return false
	}
	equals := strings.Index(lower, "=")
	if equals <= len("for ") || !isVBAControlVariable(lower[len("for "):equals]) {
		return false
	}
	rest := strings.Fields(lower[equals+1:])
	if len(rest) < 3 {
		return false
	}
	toIndex := -1
	for i := 1; i < len(rest)-1; i++ {
		if rest[i] == "to" {
			if toIndex >= 0 {
				return false
			}
			toIndex = i
		}
	}
	if toIndex < 1 || toIndex >= len(rest)-1 {
		return false
	}
	for i, field := range rest {
		if field == "step" && (i == len(rest)-1 || i < toIndex) {
			return false
		}
	}
	return true
}

func validDoConditionOpener(lower, prefix string) bool {
	return hasNonEmptySuffix(lower, prefix)
}

func hasNonEmptySuffix(text, prefix string) bool {
	return strings.HasPrefix(text, prefix) && strings.TrimSpace(strings.TrimPrefix(text, prefix)) != ""
}

// hasVBAKeywordPrefix distinguishes a statement keyword from identifiers that
// merely start with the same text, such as nextChar or NextHashCapacity.
func hasVBAKeywordPrefix(text, keyword string) bool {
	if !strings.HasPrefix(text, keyword) {
		return false
	}
	suffix := text[len(keyword):]
	if suffix == "" {
		return true
	}
	for _, r := range suffix {
		return r != '_' && !unicode.IsLetter(r) && !unicode.IsDigit(r) && !strings.ContainsRune("$%&!#@^", r)
	}
	return true
}

func isVBAControlVariable(text string) bool {
	name := strings.TrimSpace(text)
	if name == "" {
		return false
	}
	runes := []rune(name)
	if suffix := runes[len(runes)-1]; strings.ContainsRune("$%&!#@^", suffix) {
		runes = runes[:len(runes)-1]
	}
	if len(runes) == 0 || (runes[0] != '_' && !unicode.IsLetter(runes[0])) {
		return false
	}
	for _, r := range runes[1:] {
		if r != '_' && !unicode.IsLetter(r) && !unicode.IsDigit(r) {
			return false
		}
	}
	return true
}

func isProcedureEndStatement(lower string) bool {
	return lower == "end sub" || lower == "end function" || lower == "end property"
}

func isIfBranchStatement(lower string) bool {
	return lower == "else" || strings.HasPrefix(lower, "elseif ")
}

func isSelectCaseBranch(lower string) bool {
	return lower == "case else" || strings.HasPrefix(lower, "case ")
}

func isRemComment(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return lower == "rem" || strings.HasPrefix(lower, "rem ")
}
