package lint

import (
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
			if len(stack) > 0 && stack[len(stack)-1].kind != "if" {
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
		if strings.HasPrefix(lower, "next") || strings.HasPrefix(lower, "loop while") || strings.HasPrefix(lower, "loop until") {
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

type conditionalIfFrame struct {
	baseline map[int]struct{}
	branches map[int]struct{}
	sawElse  bool
}

func conditionalIfBalanceInvalid(source string) bool {
	depths := map[int]struct{}{0: {}}
	frames := make([]conditionalIfFrame, 0)
	for _, line := range normalizedSourceLines(source) {
		lower := strings.ToLower(strings.TrimSpace(gui.StripComment(line)))
		switch {
		case strings.HasPrefix(lower, "#if ") && strings.HasSuffix(lower, " then"):
			frames = append(frames, conditionalIfFrame{
				baseline: cloneDepthSet(depths),
				branches: make(map[int]struct{}),
			})
		case strings.HasPrefix(lower, "#elseif ") && strings.HasSuffix(lower, " then"):
			if len(frames) == 0 {
				return true
			}
			frame := &frames[len(frames)-1]
			mergeDepthSets(frame.branches, depths)
			depths = cloneDepthSet(frame.baseline)
		case lower == "#else":
			if len(frames) == 0 {
				return true
			}
			frame := &frames[len(frames)-1]
			mergeDepthSets(frame.branches, depths)
			frame.sawElse = true
			depths = cloneDepthSet(frame.baseline)
		case lower == "#end if":
			if len(frames) == 0 {
				return true
			}
			frame := frames[len(frames)-1]
			frames = frames[:len(frames)-1]
			mergeDepthSets(frame.branches, depths)
			if !frame.sawElse {
				mergeDepthSets(frame.branches, frame.baseline)
			}
			depths = frame.branches
		case isMultilineIfOpener(lower):
			depths = shiftDepths(depths, 1)
		case lower == "end if":
			var valid bool
			depths, valid = decrementDepths(depths)
			if !valid {
				return true
			}
		}
	}
	if len(frames) != 0 {
		return true
	}
	for depth := range depths {
		if depth != 0 {
			return true
		}
	}
	return false
}

func isMultilineIfOpener(lower string) bool {
	block, ok := blockOpener(lower, false)
	return ok && block.kind == "if"
}

func cloneDepthSet(source map[int]struct{}) map[int]struct{} {
	cloned := make(map[int]struct{}, len(source))
	mergeDepthSets(cloned, source)
	return cloned
}

func mergeDepthSets(target, source map[int]struct{}) {
	for depth := range source {
		target[depth] = struct{}{}
	}
}

func shiftDepths(source map[int]struct{}, delta int) map[int]struct{} {
	shifted := make(map[int]struct{}, len(source))
	for depth := range source {
		shifted[depth+delta] = struct{}{}
	}
	return shifted
}

func decrementDepths(source map[int]struct{}) (map[int]struct{}, bool) {
	decremented := make(map[int]struct{}, len(source))
	for depth := range source {
		if depth == 0 {
			return nil, false
		}
		decremented[depth-1] = struct{}{}
	}
	return decremented, true
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
			return nil, false
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
	kind   string
	label  string
	closer string
	line   int
	column int
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
