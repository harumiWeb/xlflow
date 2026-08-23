package analyze

import (
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// VBA's Range.Value and Range.Value2 contract is deliberately represented by
// a small shape lattice. A one-cell range returns a scalar; every known
// multi-cell range returns a two-dimensional, one-based Variant array. Any
// path or expression that prevents us from proving one of those facts is
// unknown and must not be treated as a one-dimensional array.
type rangeValueShapeKind uint8

const (
	rangeValueShapeUnknown rangeValueShapeKind = iota
	rangeValueShapeScalar
	rangeValueShapeArray2D
)

type rangeValueShape struct {
	kind rangeValueShapeKind
	rows int
	cols int
}

type rangeShape struct {
	known   bool
	array2D bool
	rows    int
	cols    int
}

type rangeValueFlowState struct {
	values      map[string]rangeValueShape
	ranges      map[string]rangeShape
	arrayGuards map[string]bool
}

type rangeValueFacts struct {
	rangeVariables map[string]bool
	intervals      map[string]rangeValueInterval
	constants      map[string]int
	expressions    map[int]procedureir.Expression
}

type rangeValueInterval struct {
	known bool
	min   int
	max   int
}

type rangeValueIssue struct {
	message    string
	reason     string
	suggestion string
}

type rangeValueSourceStatement struct {
	line    int
	endLine int
	text    string
}

type rangeValueSourceGuardFrame struct {
	name    string
	negated bool
	active  bool
}

type rangeValueSourceBranchFrame struct {
	before       rangeValueFlowState
	branches     rangeValueFlowState
	hasBranches  bool
	hasElse      bool
	exited       bool
	guardName    string
	guardNegated bool
}

type rangeValueSourceConditionalFrame struct {
	parentActive  bool
	active        bool
	branchTaken   bool
	branchUnknown bool
}

type rangeValueSourceLoopFrame struct {
	before rangeValueFlowState
}

var (
	rangeValueAddressRe     = regexp.MustCompile(`(?i)^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$`)
	rangeValueBoundRe       = regexp.MustCompile(`(?i)\b([UL])Bound\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*(?:,\s*([^)]*))?\)`)
	rangeValueGuardRe       = regexp.MustCompile(`(?i)^\s*If\s+(Not\s+)?IsArray\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)\s+Then\b`)
	rangeValueSetRe         = regexp.MustCompile(`(?i)^\s*Set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+?)\s*$`)
	rangeValueRedimRe       = regexp.MustCompile(`(?i)^\s*(?:ReDim|Erase)\s+(?:Preserve\s+)?([A-Za-z_][A-Za-z0-9_]*)\b`)
	rangeValueForEachRe     = regexp.MustCompile(`(?i)^\s*For\s+Each\s+([A-Za-z_][A-Za-z0-9_]*)\s+In\b`)
	rangeValueForVariableRe = regexp.MustCompile(`(?i)^\s*For\s+([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	rangeValueIdentifierRe  = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]*$`)
	rangeValueInlineGuardRe = regexp.MustCompile(`(?i)^\s*If\s+(Not\s+)?IsArray\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)\s+Then\s+(.+)$`)
)

func (a Analyzer) rangeValueShapeFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectRangeValueArrayShape {
		return nil
	}
	if len(proc.Statements) == 0 || rangeValueProjectionUnknown(proc) {
		return a.rangeValueShapeFindingsFromSource(file, proc)
	}
	facts := rangeValueFactsForProcedure(file, proc)
	if proc.Graph == nil {
		state := newRangeValueFlowState()
		var findings []Finding
		for _, statement := range proc.Statements {
			issues, next := rangeValueStatement(state, statement, facts)
			findings = appendRangeValueFindings(findings, file, proc, statement, issues, a)
			state = next
		}
		return findings
	}

	reachable := map[vbacfg.BlockID]bool{}
	for _, id := range proc.Graph.Reachable(vbacfg.EdgeFilter{NormalOnly: true}) {
		reachable[id] = true
	}
	if !reachable[proc.Graph.Entry] {
		return nil
	}

	states := map[vbacfg.BlockID]rangeValueFlowState{
		proc.Graph.Entry: newRangeValueFlowState(),
	}
	blocks := make(map[vbacfg.BlockID]vbacfg.Block, len(proc.Graph.Blocks))
	for _, block := range proc.Graph.Blocks {
		blocks[block.ID] = block
	}
	queue := []vbacfg.BlockID{proc.Graph.Entry}
	queued := map[vbacfg.BlockID]bool{proc.Graph.Entry: true}
	findings := map[string]Finding{}
	for len(queue) > 0 {
		id := queue[0]
		queue = queue[1:]
		queued[id] = false
		in := states[id]
		block, ok := blocks[id]
		out := cloneRangeValueFlowState(in)
		if ok && block.Statement != nil {
			issues, next := rangeValueStatement(in, *block.Statement, facts)
			out = next
			for _, issue := range issues {
				finding := a.simpleFinding(file, proc, block.Statement.Range.StartLine, "VBA226", "warning", issue.message, issue.reason, issue.suggestion)
				key := fmt.Sprintf("%d:%s", block.Statement.ID, issue.message)
				findings[key] = finding
			}
		}
		for _, edge := range proc.Graph.Edges {
			if edge.From != id || edge.Class != vbacfg.EdgeNormal || !reachable[edge.To] {
				continue
			}
			next := cloneRangeValueFlowState(out)
			applyRangeValueGuard(next, block.Statement, edge)
			merged, changed := mergeRangeValueFlowState(states[edge.To], next, states[edge.To].values != nil || states[edge.To].ranges != nil || states[edge.To].arrayGuards != nil)
			if !changed {
				continue
			}
			states[edge.To] = merged
			if !queued[edge.To] {
				queue = append(queue, edge.To)
				queued[edge.To] = true
			}
		}
	}

	out := make([]Finding, 0, len(findings))
	for _, finding := range findings {
		out = append(out, finding)
	}
	sort.SliceStable(out, func(i, j int) bool {
		if out[i].Line != out[j].Line {
			return out[i].Line < out[j].Line
		}
		return out[i].Message < out[j].Message
	})
	return out
}

// rangeValueShapeFindingsFromSource is a deliberately narrow recovery path.
// Procedure IR normally supplies canonical expressions and CFG edges; when a
// recovered procedure has no statement projection, source lines are the only
// available evidence. The textual transfer remains conservative and is used
// only for this unpopulated projection, so it cannot change complete-IR CFG
// semantics or add another fixed-point walk.
func (a Analyzer) rangeValueShapeFindingsFromSource(file parsedFile, proc sourceProcedure) []Finding {
	sourceStatements := rangeValueSourceStatements(file, proc)
	if len(sourceStatements) == 0 {
		return nil
	}
	facts := rangeValueFactsForProcedure(file, proc)
	state := newRangeValueFlowState()
	var guards []rangeValueSourceGuardFrame
	var branches []rangeValueSourceBranchFrame
	var loops []rangeValueSourceLoopFrame
	var findings []Finding
	for index, source := range sourceStatements {
		line := source.endLine
		if line <= 0 {
			line = source.line
		}
		statement := procedureir.Statement{
			ID:   index + 1,
			Text: source.text,
			Range: vbaast.Range{
				StartLine: line,
				EndLine:   line,
			},
		}
		if match := rangeValueInlineGuardRe.FindStringSubmatch(statement.Text); len(match) == 4 {
			name := strings.ToLower(match[2])
			negated := strings.TrimSpace(match[1]) != ""
			if rangeValueSourceEarlyExit(match[3]) {
				if negated {
					state.arrayGuards[name] = true
				} else {
					delete(state.arrayGuards, name)
				}
				continue
			}
			priorGuard := state.arrayGuards[name]
			if !negated {
				state.arrayGuards[name] = true
			} else {
				delete(state.arrayGuards, name)
			}
			body := statement
			body.Text = strings.TrimSpace(match[3])
			issues, next := rangeValueStatement(state, body, facts)
			if !next.arrayGuards[name] || !priorGuard {
				delete(next.arrayGuards, name)
			}
			findings = appendRangeValueFindings(findings, file, proc, body, issues, a)
			state = next
			continue
		}
		lower := strings.ToLower(strings.TrimSpace(statement.Text))
		if rangeValueSourceLoopStart(statement.Text) {
			loops = append(loops, rangeValueSourceLoopFrame{before: cloneRangeValueFlowState(state)})
		} else if rangeValueSourceLoopEnd(statement.Text) && len(loops) > 0 {
			frame := loops[len(loops)-1]
			loops = loops[:len(loops)-1]
			state = mergeRangeValueSourceStates(frame.before, state)
		} else if isRangeValueBlockIf(statement.Text) {
			frame := rangeValueSourceBranchFrame{before: cloneRangeValueFlowState(state)}
			if match := rangeValueGuardRe.FindStringSubmatch(statement.Text); len(match) == 3 {
				frame.guardName = strings.ToLower(match[2])
				frame.guardNegated = strings.TrimSpace(match[1]) != ""
			}
			branches = append(branches, frame)
		} else if strings.HasPrefix(lower, "elseif ") && len(branches) > 0 {
			frame := &branches[len(branches)-1]
			current := cloneRangeValueFlowState(state)
			if !frame.exited {
				if frame.hasBranches {
					frame.branches = mergeRangeValueSourceStates(frame.branches, current)
				} else {
					frame.branches = current
					frame.hasBranches = true
				}
			}
			frame.exited = false
			state = cloneRangeValueFlowState(frame.before)
		} else if rangeValueSourceElse(statement.Text) && len(branches) > 0 {
			frame := &branches[len(branches)-1]
			current := cloneRangeValueFlowState(state)
			frame.hasElse = true
			if !frame.exited {
				if frame.hasBranches {
					frame.branches = mergeRangeValueSourceStates(frame.branches, current)
				} else {
					frame.branches = current
					frame.hasBranches = true
				}
			}
			frame.exited = false
			state = cloneRangeValueFlowState(frame.before)
		} else if strings.HasPrefix(lower, "end if") && len(branches) > 0 {
			current := cloneRangeValueFlowState(state)
			frame := branches[len(branches)-1]
			branches = branches[:len(branches)-1]
			if frame.hasElse {
				if frame.exited {
					if frame.hasBranches {
						state = frame.branches
					} else {
						state = cloneRangeValueFlowState(frame.before)
					}
				} else if frame.hasBranches {
					state = mergeRangeValueSourceStates(frame.branches, current)
				} else {
					state = current
				}
			} else if frame.hasBranches {
				state = mergeRangeValueSourceStates(frame.before, frame.branches)
				if !frame.exited {
					state = mergeRangeValueSourceStates(state, current)
				}
			} else if frame.exited {
				state = cloneRangeValueFlowState(frame.before)
				if frame.guardName != "" {
					if frame.guardNegated {
						state.arrayGuards[frame.guardName] = true
					} else {
						delete(state.arrayGuards, frame.guardName)
					}
				}
			} else {
				state = mergeRangeValueSourceStates(frame.before, current)
			}
		}
		if rangeValueSourceEarlyExit(statement.Text) && len(branches) > 0 {
			branches[len(branches)-1].exited = true
		}
		updateRangeValueSourceGuard(&state, &guards, statement.Text)
		issues, next := rangeValueStatement(state, statement, facts)
		findings = appendRangeValueFindings(findings, file, proc, statement, issues, a)
		state = next
		if rangeValueSourceEarlyExit(statement.Text) && len(branches) == 0 {
			break
		}
	}
	return findings
}

func mergeRangeValueSourceStates(first, second rangeValueFlowState) rangeValueFlowState {
	merged, _ := mergeRangeValueFlowState(first, second, true)
	return merged
}

func updateRangeValueSourceGuard(state *rangeValueFlowState, guards *[]rangeValueSourceGuardFrame, text string) {
	if state == nil || guards == nil {
		return
	}
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.HasPrefix(lower, "end if") {
		if len(*guards) == 0 {
			return
		}
		frame := (*guards)[len(*guards)-1]
		*guards = (*guards)[:len(*guards)-1]
		if frame.name != "" && !state.arrayGuards[frame.name] {
			delete(state.arrayGuards, frame.name)
		}
		return
	}
	if rangeValueSourceElse(text) || strings.HasPrefix(lower, "elseif ") {
		if len(*guards) == 0 {
			return
		}
		if strings.HasPrefix(lower, "elseif ") {
			frame := &(*guards)[len(*guards)-1]
			if frame.name == "" {
				return
			}
			if frame.active {
				delete(state.arrayGuards, frame.name)
				frame.active = false
			}
			trimmed := strings.TrimSpace(text)
			match := rangeValueGuardRe.FindStringSubmatch("If " + strings.TrimSpace(trimmed[len("ElseIf "):]))
			if len(match) == 3 {
				frame.negated = strings.TrimSpace(match[1]) != ""
				frame.active = !frame.negated
				if frame.active {
					state.arrayGuards[frame.name] = true
				}
			}
			return
		}
		frame := &(*guards)[len(*guards)-1]
		if frame.name != "" {
			if frame.active {
				delete(state.arrayGuards, frame.name)
			}
			frame.active = frame.negated
			if frame.active {
				state.arrayGuards[frame.name] = true
			}
		}
		return
	}
	if !isRangeValueBlockIf(text) {
		return
	}
	match := rangeValueGuardRe.FindStringSubmatch(text)
	if len(match) != 3 {
		*guards = append(*guards, rangeValueSourceGuardFrame{})
		return
	}
	name := strings.ToLower(match[2])
	negated := strings.TrimSpace(match[1]) != ""
	frame := rangeValueSourceGuardFrame{name: name, negated: negated, active: !negated}
	*guards = append(*guards, frame)
	if frame.active {
		state.arrayGuards[name] = true
	} else {
		delete(state.arrayGuards, name)
	}
}

func isRangeValueBlockIf(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if !strings.HasPrefix(lower, "if ") {
		return false
	}
	then := strings.LastIndex(lower, " then")
	return then >= 0 && strings.TrimSpace(lower[then+len(" then"):]) == ""
}

func rangeValueSourceLoopStart(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return strings.HasPrefix(lower, "for ") || lower == "do" || strings.HasPrefix(lower, "do while ") || strings.HasPrefix(lower, "do until ") || strings.HasPrefix(lower, "while ")
}

func rangeValueSourceLoopEnd(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return rangeValueSourceKeywordLine(lower, "next") || rangeValueSourceKeywordLine(lower, "loop") || rangeValueSourceKeywordLine(lower, "wend")
}

func rangeValueSourceElse(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return rangeValueSourceKeywordLine(lower, "else")
}

func rangeValueSourceKeywordLine(lower, keyword string) bool {
	return lower == keyword || strings.HasPrefix(lower, keyword+" ")
}

func rangeValueSourceLinesApplicable(file parsedFile, proc sourceProcedure) bool {
	for _, source := range rangeValueSourceStatements(file, proc) {
		code := strings.ToLower(normalizedCodeLine(source.text))
		if strings.Contains(code, ".value") || strings.Contains(code, "range(") || strings.Contains(code, "resize(") || strings.Contains(code, "cells(") {
			return true
		}
	}
	return false
}

func rangeValueSourceEarlyExit(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{"exit sub", "exit function", "exit property", "return"} {
		if lower == prefix || strings.HasPrefix(lower, prefix+" ") {
			return true
		}
	}
	return false
}

func rangeValueSourceStatements(file parsedFile, proc sourceProcedure) []rangeValueSourceStatement {
	lines := file.Lines
	if len(lines) == 0 && len(file.Source) > 0 {
		lines = normalizedSourceLines(string(file.Source))
	}
	if len(lines) == 0 {
		return nil
	}
	start := proc.StartLine
	if start < 1 {
		start = 1
	}
	end := proc.EndLine
	if end <= 0 || end > len(lines) {
		end = len(lines)
	}
	if start > end {
		return nil
	}

	conditionalConstants := file.RangeValueModuleConstants
	if conditionalConstants == nil {
		conditionalConstants = rangeValueModuleIntegerConstants(lines, file.IR)
	}
	conditionalActive := true
	conditionals := []rangeValueSourceConditionalFrame{}
	var out []rangeValueSourceStatement
	var continued strings.Builder
	continuedLine := 0
	continuedEndLine := 0
	flush := func(line, endLine int) {
		text := strings.TrimSpace(continued.String())
		if text != "" {
			for _, part := range splitRangeValueSourceStatements(text) {
				if rangeValueIsRemComment(part) {
					break
				}
				part = rangeValueStripRemComment(part)
				if strings.TrimSpace(part) != "" {
					out = append(out, rangeValueSourceStatement{line: line, endLine: endLine, text: strings.TrimSpace(part)})
				}
			}
		}
		continued.Reset()
		continuedLine = 0
	}
	for line := 1; line <= end; line++ {
		text := rangeValueStripRemComment(rawWorksheetCodeLine(lines[line-1]))
		text = strings.TrimSpace(text)
		if handled, active := rangeValueSourceConditionalDirective(text, &conditionals, conditionalActive, conditionalConstants); handled {
			conditionalActive = active
			continue
		}
		if line < start || !conditionalActive {
			continue
		}
		if text == "" {
			continue
		}
		if continued.Len() == 0 {
			continuedLine = line
		}
		continuedEndLine = line
		continuedLineText := text
		if vbaLineContinues(text) {
			text = strings.TrimSpace(strings.TrimSuffix(text, "_"))
		}
		if continued.Len() > 0 && text != "" {
			continued.WriteByte(' ')
		}
		continued.WriteString(text)
		if !vbaLineContinues(continuedLineText) {
			flush(continuedLine, continuedEndLine)
		}
	}
	if continued.Len() > 0 {
		flush(continuedLine, continuedEndLine)
	}
	return out
}

func rangeValueSourceConditionalDirective(text string, stack *[]rangeValueSourceConditionalFrame, active bool, constants map[string]int) (bool, bool) {
	if stack == nil {
		return false, active
	}
	lower := strings.ToLower(strings.TrimSpace(text))
	switch {
	case strings.HasPrefix(lower, "#if "):
		condition := strings.TrimSpace(text[4:])
		condition = rangeValueSourceConditionalExpression(condition)
		parentActive := active
		frame := rangeValueSourceConditionalFrame{parentActive: parentActive}
		if parentActive {
			switch rangeValueSourceConditionalValue(condition, constants) {
			case 1:
				frame.active, frame.branchTaken = true, true
			case 0:
				frame.active = false
			default:
				frame.active, frame.branchUnknown = true, true
			}
		}
		*stack = append(*stack, frame)
		return true, frame.active
	case strings.HasPrefix(lower, "#elseif "):
		if len(*stack) == 0 {
			return true, active
		}
		frame := &(*stack)[len(*stack)-1]
		if !frame.parentActive || frame.branchTaken {
			frame.active = false
			return true, false
		}
		if frame.branchUnknown {
			frame.active = true
			return true, true
		}
		condition := strings.TrimSpace(text[len("#ElseIf "):])
		condition = rangeValueSourceConditionalExpression(condition)
		switch rangeValueSourceConditionalValue(condition, constants) {
		case 1:
			frame.active, frame.branchTaken = true, true
		case 0:
			frame.active = false
		default:
			frame.active, frame.branchUnknown = true, true
		}
		return true, frame.active
	case strings.HasPrefix(lower, "#else"):
		if len(*stack) == 0 {
			return true, active
		}
		frame := &(*stack)[len(*stack)-1]
		frame.active = frame.parentActive && !frame.branchTaken
		frame.branchTaken = frame.parentActive
		return true, frame.active
	case strings.HasPrefix(lower, "#end if"):
		if len(*stack) == 0 {
			return true, active
		}
		frame := (*stack)[len(*stack)-1]
		*stack = (*stack)[:len(*stack)-1]
		return true, frame.parentActive
	default:
		return false, active
	}
}

func rangeValueSourceConditionalValue(text string, constants map[string]int) int {
	switch strings.ToLower(strings.TrimSpace(text)) {
	case "true":
		return 1
	case "false":
		return 0
	}
	value, err := constantIntegerExpression(text, constants)
	if err != nil {
		return -1
	}
	if value == 0 {
		return 0
	}
	return 1
}

func rangeValueSourceConditionalExpression(text string) string {
	text = strings.TrimSpace(text)
	if strings.HasSuffix(strings.ToLower(text), " then") {
		return strings.TrimSpace(text[:len(text)-len(" then")])
	}
	return text
}

func rangeValueStripRemComment(text string) string {
	if rangeValueIsRemComment(text) {
		return ""
	}
	return text
}

func rangeValueIsRemComment(text string) bool {
	trimmed := strings.TrimSpace(text)
	if len(trimmed) == 3 && strings.EqualFold(trimmed, "rem") {
		return true
	}
	if len(trimmed) > 3 && strings.EqualFold(trimmed[:3], "rem") {
		next := trimmed[3]
		if next == ' ' || next == '\t' {
			return true
		}
	}
	return false
}

func splitRangeValueSourceStatements(text string) []string {
	var out []string
	start := 0
	inString := false
	depth := 0
	for index := 0; index < len(text); index++ {
		switch text[index] {
		case '"':
			if inString && index+1 < len(text) && text[index+1] == '"' {
				index++
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
		case ':':
			if !inString && depth == 0 {
				out = append(out, strings.TrimSpace(text[start:index]))
				start = index + 1
			}
		}
	}
	out = append(out, strings.TrimSpace(text[start:]))
	return out
}

func appendRangeValueFindings(findings []Finding, file parsedFile, proc sourceProcedure, statement procedureir.Statement, issues []rangeValueIssue, analyzer Analyzer) []Finding {
	for _, issue := range issues {
		findings = append(findings, analyzer.simpleFinding(file, proc, statement.Range.StartLine, "VBA226", "warning", issue.message, issue.reason, issue.suggestion))
	}
	return findings
}

func newRangeValueFlowState() rangeValueFlowState {
	return rangeValueFlowState{
		values:      map[string]rangeValueShape{},
		ranges:      map[string]rangeShape{},
		arrayGuards: map[string]bool{},
	}
}

func cloneRangeValueFlowState(in rangeValueFlowState) rangeValueFlowState {
	out := newRangeValueFlowState()
	for name, shape := range in.values {
		out.values[name] = shape
	}
	for name, shape := range in.ranges {
		out.ranges[name] = shape
	}
	for name, guarded := range in.arrayGuards {
		out.arrayGuards[name] = guarded
	}
	return out
}

func mergeRangeValueFlowState(existing, incoming rangeValueFlowState, hasExisting bool) (rangeValueFlowState, bool) {
	if !hasExisting {
		return cloneRangeValueFlowState(incoming), true
	}
	out := cloneRangeValueFlowState(existing)
	changed := false
	for name, shape := range incoming.values {
		prior, ok := out.values[name]
		if !ok {
			out.values[name] = rangeValueShape{kind: rangeValueShapeUnknown}
			changed = true
			continue
		}
		merged := mergeRangeValueShape(prior, shape)
		if merged != prior {
			out.values[name] = merged
			changed = true
		}
	}
	for name := range out.values {
		if _, ok := incoming.values[name]; !ok && out.values[name].kind != rangeValueShapeUnknown {
			out.values[name] = rangeValueShape{kind: rangeValueShapeUnknown}
			changed = true
		}
	}
	for name, shape := range incoming.ranges {
		prior, ok := out.ranges[name]
		if !ok {
			out.ranges[name] = rangeShape{}
			changed = true
			continue
		}
		merged := mergeRangeShape(prior, shape)
		if merged != prior {
			out.ranges[name] = merged
			changed = true
		}
	}
	for name := range out.ranges {
		if _, ok := incoming.ranges[name]; !ok && out.ranges[name].known {
			out.ranges[name] = rangeShape{}
			changed = true
		}
	}
	for name := range out.arrayGuards {
		if !incoming.arrayGuards[name] {
			delete(out.arrayGuards, name)
			changed = true
		}
	}
	return out, changed
}

func mergeRangeValueShape(a, b rangeValueShape) rangeValueShape {
	if a == b {
		return a
	}
	return rangeValueShape{kind: rangeValueShapeUnknown}
}

func mergeRangeShape(a, b rangeShape) rangeShape {
	if a == b {
		return a
	}
	return rangeShape{}
}

func rangeValueFactsForProcedure(file parsedFile, proc sourceProcedure) rangeValueFacts {
	facts := rangeValueFacts{
		rangeVariables: map[string]bool{},
		intervals:      map[string]rangeValueInterval{},
		constants:      map[string]int{},
		expressions:    map[int]procedureir.Expression{},
	}
	for _, declaration := range proc.Declarations {
		if isExcelRangeType(declaration.Type) {
			facts.rangeVariables[strings.ToLower(declaration.Name)] = true
		}
	}
	for _, expression := range proc.Expressions {
		facts.expressions[expression.ID] = expression
	}
	constants := rangeValueIntegerConstants(file.RangeValueModuleConstants, proc)
	facts.constants = constants
	for _, statement := range proc.Statements {
		text := strings.TrimSpace(excelLoopHeaderText(statement.Text))
		if match := forBoundsRe.FindStringSubmatch(text); len(match) > 0 {
			start, startErr := constantIntegerExpression(match[1], constants)
			end, endErr := constantIntegerExpression(match[2], constants)
			if startErr == nil && endErr == nil {
				min, max := start, end
				if min > max {
					min, max = max, min
				}
				name := ""
				if matchName := rangeValueForVariableRe.FindStringSubmatch(text); len(matchName) == 2 {
					name = strings.ToLower(matchName[1])
				}
				if name != "" {
					facts.intervals[name] = rangeValueInterval{known: true, min: min, max: max}
				}
			}
		}
	}
	return facts
}

func rangeValueModuleIntegerConstants(lines []string, ir procedureir.DocumentIR) map[string]int {
	constants := map[string]int{}
	insideProcedure := make([]bool, len(lines)+1)
	for _, candidate := range ir.Procedures {
		start := max(candidate.Symbol.DeclarationRange.StartLine, 1)
		end := min(candidate.Symbol.DeclarationRange.EndLine, len(lines))
		for line := start; line <= end; line++ {
			insideProcedure[line] = true
		}
	}
	conditionalDepth := 0
	for lineIndex, line := range lines {
		if insideProcedure[lineIndex+1] {
			continue
		}
		code := strings.TrimSpace(normalizedCodeLine(line))
		if strings.HasPrefix(strings.ToLower(code), "#if ") {
			conditionalDepth++
			continue
		}
		if strings.HasPrefix(strings.ToLower(code), "#end if") {
			if conditionalDepth > 0 {
				conditionalDepth--
			}
			continue
		}
		if conditionalDepth > 0 || strings.HasPrefix(strings.ToLower(code), "#elseif ") || strings.HasPrefix(strings.ToLower(code), "#else") {
			continue
		}
		match := constIntegerRe.FindStringSubmatch(code)
		if len(match) != 3 {
			continue
		}
		if value, err := constantIntegerExpression(match[2], constants); err == nil {
			constants[strings.ToLower(match[1])] = value
		}
	}
	return constants
}

func rangeValueIntegerConstants(moduleConstants map[string]int, proc sourceProcedure) map[string]int {
	constants := make(map[string]int, len(moduleConstants))
	for name, value := range moduleConstants {
		constants[name] = value
	}
	for _, statement := range proc.Statements {
		if len(statement.ConditionalBranches) > 0 {
			continue
		}
		match := constIntegerRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
		if len(match) != 3 {
			continue
		}
		if value, err := constantIntegerExpression(match[2], constants); err == nil {
			constants[strings.ToLower(match[1])] = value
		}
	}
	return constants
}

func rangeValueStatement(state rangeValueFlowState, statement procedureir.Statement, facts rangeValueFacts) ([]rangeValueIssue, rangeValueFlowState) {
	state = cloneRangeValueFlowState(state)
	var issues []rangeValueIssue
	raw := rangeValueStatementText(statement)
	code := normalizedCodeLine(raw)
	boundMatches := rangeValueBoundRe.FindAllStringSubmatch(code, -1)
	for _, name := range sortedRangeValueNames(state.values) {
		shape := state.values[name]
		for _, match := range rangeValueIndexedUses(code, name) {
			args := splitArgs(match)
			switch len(args) {
			case 1:
				if shape.kind == rangeValueShapeScalar {
					issues = append(issues, rangeValueIssue{
						message:    fmt.Sprintf("Single-cell Range.Value result %s is indexed as an array.", name),
						reason:     "A definite single-cell Range.Value or Range.Value2 result is a scalar Variant, not an array.",
						suggestion: fmt.Sprintf("Use %s directly instead of indexing the scalar result.", name),
					})
				} else {
					issues = append(issues, rangeValueIssue{
						message:    fmt.Sprintf("Range.Value result %s is used with one array index.", name),
						reason:     "A multi-cell Range.Value or Range.Value2 result is always a two-dimensional array, while a single-cell result is a scalar.",
						suggestion: fmt.Sprintf("Use %s(row, column) for a multi-cell range, or handle the single-cell case before indexing.", name),
					})
				}
			case 2:
				if shape.kind == rangeValueShapeScalar {
					issues = append(issues, rangeValueIssue{
						message:    fmt.Sprintf("Single-cell Range.Value result %s is indexed as an array.", name),
						reason:     "A definite single-cell Range.Value or Range.Value2 result is a scalar Variant, not an array.",
						suggestion: fmt.Sprintf("Use %s directly, or guard dynamic ranges with IsArray before using two-dimensional indexes.", name),
					})
				} else if shape.kind == rangeValueShapeUnknown && !state.arrayGuards[name] {
					issues = append(issues, rangeValueIssue{
						message:    fmt.Sprintf("Range.Value result %s may be a scalar before two-dimensional indexing.", name),
						reason:     "The range size is dynamic, so Range.Value can return either a scalar or a two-dimensional array on different executions.",
						suggestion: fmt.Sprintf("Guard %s with IsArray and use explicit row and column dimensions.", name),
					})
				} else if shape.kind == rangeValueShapeArray2D {
					if issue := rangeValueBoundsIssue(name, args, shape, state, facts); issue != nil {
						issues = append(issues, *issue)
					}
				}
			}
		}
		for _, match := range boundMatches {
			if !strings.EqualFold(match[2], name) {
				continue
			}
			dimension := strings.TrimSpace(match[3])
			if dimension == "" {
				issues = append(issues, rangeValueIssue{
					message:    fmt.Sprintf("UBound/LBound(%s) omits the array dimension.", name),
					reason:     "Range.Value and Range.Value2 return two-dimensional arrays for multi-cell ranges, so an omitted dimension hides whether rows or columns are being inspected.",
					suggestion: fmt.Sprintf("Specify the dimension explicitly, for example UBound(%s, 1) or UBound(%s, 2).", name, name),
				})
				continue
			}
			if shape.kind == rangeValueShapeScalar || (shape.kind == rangeValueShapeUnknown && !state.arrayGuards[name]) {
				issues = append(issues, rangeValueIssue{
					message:    fmt.Sprintf("%s may not be an array when its bounds are read.", name),
					reason:     "A single-cell or dynamic Range.Value result is not proven to be an array at this point.",
					suggestion: fmt.Sprintf("Guard %s with IsArray before calling LBound or UBound.", name),
				})
				continue
			}
			if dimensionValue, err := constantIntegerExpression(dimension, facts.constantValues()); err == nil && (dimensionValue < 1 || dimensionValue > 2) {
				issues = append(issues, rangeValueIssue{
					message:    fmt.Sprintf("%s uses unsupported array dimension %d for %s.", match[1]+"Bound", dimensionValue, name),
					reason:     "Range.Value and Range.Value2 arrays expose row and column dimensions only.",
					suggestion: fmt.Sprintf("Use dimension 1 for rows or dimension 2 for columns when reading %s bounds.", name),
				})
			}
		}
	}

	if left, right, ok := rangeValueAssignment(raw); ok {
		if sourceName := rangeValueBareName(right); sourceName != "" {
			if source, found := state.values[sourceName]; found {
				if target, foundTarget := rangeValueMemberReceiver(statement.Target, facts, "value", "value2"); foundTarget {
					if issue := rangeValueDestinationIssue(target, source, state, facts); issue != nil {
						issues = append(issues, *issue)
					}
				} else if receiver, foundTarget := rangeValueMemberReceiverText(left, "value", "value2"); foundTarget {
					if destination, recognized := rangeValueRangeShapeText(receiver, state, facts); recognized {
						if issue := rangeValueDestinationShapeIssue(destination, source); issue != nil {
							issues = append(issues, *issue)
						}
					}
				}
			}
		}
		if targetName := rangeValueBareName(left); targetName != "" {
			delete(state.arrayGuards, targetName)
			source, sourceKnown := rangeValueSourceShape(statement.Value, state, facts)
			if !sourceKnown {
				source, sourceKnown = rangeValueSourceShapeText(right, state, facts)
			}
			if sourceKnown {
				state.values[targetName] = source
				delete(state.ranges, targetName)
			} else if sourceName := rangeValueBareName(right); sourceName != "" {
				if source, found := state.values[sourceName]; found {
					state.values[targetName] = source
				} else if _, wasTracked := state.values[targetName]; wasTracked {
					state.values[targetName] = rangeValueShape{kind: rangeValueShapeUnknown}
				} else {
					delete(state.values, targetName)
				}
				delete(state.ranges, targetName)
			} else {
				if _, wasTracked := state.values[targetName]; wasTracked {
					state.values[targetName] = rangeValueShape{kind: rangeValueShapeUnknown}
				} else {
					delete(state.values, targetName)
				}
			}
		}
	}

	if match := rangeValueSetRe.FindStringSubmatch(raw); len(match) == 3 {
		name := strings.ToLower(match[1])
		shape, ok := rangeValueRangeShapeExpression(statement.Value, state, facts)
		if !ok {
			shape, ok = rangeValueRangeShapeText(match[2], state, facts)
		}
		if ok {
			state.ranges[name] = shape
		} else {
			state.ranges[name] = rangeShape{}
		}
		delete(state.values, name)
		delete(state.arrayGuards, name)
	}
	if match := rangeValueRedimRe.FindStringSubmatch(raw); len(match) == 2 {
		name := strings.ToLower(match[1])
		delete(state.values, name)
		delete(state.arrayGuards, name)
	}
	if match := rangeValueForEachRe.FindStringSubmatch(raw); len(match) == 2 {
		name := strings.ToLower(match[1])
		delete(state.values, name)
		delete(state.arrayGuards, name)
	}
	return deduplicateRangeValueIssues(issues), state
}

func (f rangeValueFacts) constantValues() map[string]int {
	return f.constants
}

func sortedRangeValueNames(values map[string]rangeValueShape) []string {
	names := make([]string, 0, len(values))
	for name := range values {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}

func rangeValueIndexedUses(code, name string) []string {
	re := regexp.MustCompile(`(?i)\b` + regexp.QuoteMeta(name) + `\s*\(([^()]*)\)`)
	matches := re.FindAllStringSubmatch(code, -1)
	out := make([]string, 0, len(matches))
	for _, match := range matches {
		if len(match) == 2 {
			out = append(out, match[1])
		}
	}
	return out
}

func rangeValueStatementText(statement procedureir.Statement) string {
	text := statement.Text
	switch statement.Kind {
	case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo,
		procedureir.StatementWhile, procedureir.StatementIf, procedureir.StatementElseIf,
		procedureir.StatementElse, procedureir.StatementSelect, procedureir.StatementCase,
		procedureir.StatementWith:
		text = excelLoopHeaderText(text)
	}
	return strings.TrimSpace(text)
}

func rangeValueBoundsIssue(name string, args []string, shape rangeValueShape, state rangeValueFlowState, facts rangeValueFacts) *rangeValueIssue {
	if len(args) != 2 {
		return nil
	}
	row := rangeValueIndexInterval(args[0], state, facts)
	column := rangeValueIndexInterval(args[1], state, facts)
	if shape.rows > 0 && row.known && (row.min < 1 || row.max > shape.rows) {
		return &rangeValueIssue{
			message:    fmt.Sprintf("%s indexes row dimension outside 1..%d.", name, shape.rows),
			reason:     "Range.Value arrays use row-first, column-second indexing and the inferred row range exceeds the source range.",
			suggestion: rangeValueBoundsSuggestion(shape),
		}
	}
	if shape.cols > 0 && column.known && (column.min < 1 || column.max > shape.cols) {
		return &rangeValueIssue{
			message:    fmt.Sprintf("%s indexes column dimension outside 1..%d.", name, shape.cols),
			reason:     "Range.Value arrays use row-first, column-second indexing and the inferred column range exceeds the source range.",
			suggestion: rangeValueBoundsSuggestion(shape),
		}
	}
	return nil
}

func rangeValueBoundsSuggestion(shape rangeValueShape) string {
	switch {
	case shape.rows > 0 && shape.cols > 0:
		return fmt.Sprintf("Keep the first index within 1..%d and the second within 1..%d.", shape.rows, shape.cols)
	case shape.rows > 0:
		return fmt.Sprintf("Keep the first index within 1..%d.", shape.rows)
	case shape.cols > 0:
		return fmt.Sprintf("Keep the second index within 1..%d.", shape.cols)
	default:
		return "Keep both indexes within the source range bounds."
	}
}

func rangeValueIndexInterval(text string, state rangeValueFlowState, facts rangeValueFacts) rangeValueInterval {
	text = strings.TrimSpace(text)
	if value, err := constantIntegerExpression(text, facts.constantValues()); err == nil {
		return rangeValueInterval{known: true, min: value, max: value}
	}
	if interval, ok := facts.intervals[strings.ToLower(cleanIdentifier(text))]; ok {
		return interval
	}
	match := rangeValueBoundRe.FindStringSubmatch(text)
	if len(match) == 4 {
		shape, ok := state.values[strings.ToLower(match[2])]
		if !ok || shape.kind != rangeValueShapeArray2D {
			return rangeValueInterval{}
		}
		dimension, err := constantIntegerExpression(strings.TrimSpace(match[3]), facts.constantValues())
		if err != nil {
			return rangeValueInterval{}
		}
		if dimension == 1 {
			if match[1] == "L" || match[1] == "l" {
				return rangeValueInterval{known: true, min: 1, max: 1}
			}
			return rangeValueInterval{known: true, min: 1, max: shape.rows}
		}
		if dimension == 2 {
			if match[1] == "L" || match[1] == "l" {
				return rangeValueInterval{known: true, min: 1, max: 1}
			}
			return rangeValueInterval{known: true, min: 1, max: shape.cols}
		}
	}
	return rangeValueInterval{}
}

func rangeValueAssignment(text string) (string, string, bool) {
	text = strings.TrimSpace(text)
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "if ") || strings.HasPrefix(lower, "elseif ") || strings.HasPrefix(lower, "for ") || strings.HasPrefix(lower, "while ") {
		return "", "", false
	}
	inString, depth := false, 0
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
		case '=':
			if inString || depth != 0 || (i > 0 && (text[i-1] == '<' || text[i-1] == '>')) || (i+1 < len(text) && text[i+1] == '>') {
				continue
			}
			left := strings.TrimSpace(strings.TrimPrefix(text[:i], "Let "))
			if left == "" {
				return "", "", false
			}
			return left, strings.TrimSpace(text[i+1:]), true
		}
	}
	return "", "", false
}

func rangeValueBareName(text string) string {
	text = strings.TrimSpace(text)
	text = strings.TrimPrefix(strings.TrimPrefix(text, "Let "), "let ")
	if !rangeValueIdentifierRe.MatchString(text) {
		return ""
	}
	return strings.ToLower(text)
}

func rangeValueSourceShape(value *procedureir.Expression, state rangeValueFlowState, facts rangeValueFacts) (rangeValueShape, bool) {
	receiver, ok := rangeValueMemberReceiver(value, facts, "value", "value2")
	if !ok {
		return rangeValueShape{}, false
	}
	shape, recognized := rangeValueRangeShapeExpression(&receiver, state, facts)
	if !recognized {
		return rangeValueShape{}, false
	}
	if shape.known && shape.rows == 1 && shape.cols == 1 {
		return rangeValueShape{kind: rangeValueShapeScalar}, true
	}
	if shape.known || shape.array2D {
		return rangeValueShape{kind: rangeValueShapeArray2D, rows: shape.rows, cols: shape.cols}, true
	}
	return rangeValueShape{kind: rangeValueShapeUnknown}, true
}

func rangeValueSourceShapeText(value string, state rangeValueFlowState, facts rangeValueFacts) (rangeValueShape, bool) {
	receiver, ok := rangeValueMemberReceiverText(value, "value", "value2")
	if !ok {
		return rangeValueShape{}, false
	}
	shape, recognized := rangeValueRangeShapeText(receiver, state, facts)
	if !recognized {
		return rangeValueShape{}, false
	}
	if shape.known && shape.rows == 1 && shape.cols == 1 {
		return rangeValueShape{kind: rangeValueShapeScalar}, true
	}
	if shape.known || shape.array2D {
		return rangeValueShape{kind: rangeValueShapeArray2D, rows: shape.rows, cols: shape.cols}, true
	}
	return rangeValueShape{kind: rangeValueShapeUnknown}, true
}

func rangeValueMemberReceiverText(expression string, members ...string) (string, bool) {
	text := strings.TrimSpace(expression)
	for _, member := range members {
		suffix := "." + strings.ToLower(member)
		lower := strings.ToLower(text)
		if strings.HasSuffix(lower, suffix) {
			receiver := strings.TrimSpace(text[:len(text)-len(suffix)])
			if receiver != "" {
				return receiver, true
			}
		}
	}
	return "", false
}

func rangeValueRangeShapeText(expression string, state rangeValueFlowState, facts rangeValueFacts) (rangeShape, bool) {
	text := strings.TrimSpace(expression)
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' && matchingParen(text, 0) == len(text)-1 {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	if name := rangeValueBareName(text); name != "" {
		if shape, found := state.ranges[name]; found {
			return shape, true
		}
		if facts.rangeVariables[name] {
			return rangeShape{}, true
		}
		return rangeShape{}, false
	}
	open := firstParenOutsideString(text)
	if open < 0 {
		return rangeShape{}, false
	}
	close := matchingParen(text, open)
	if close < 0 || strings.TrimSpace(text[close+1:]) != "" {
		return rangeShape{}, false
	}
	name := strings.ToLower(cleanIdentifier(lastName(strings.TrimSpace(text[:open]))))
	args := splitArgs(text[open+1 : close])
	switch name {
	case "cells":
		if len(args) == 2 {
			return rangeShape{known: true, rows: 1, cols: 1}, true
		}
		return rangeShape{}, true
	case "range":
		if shape, found := rangeValueTextCellsPairShape(args, facts); found {
			return shape, true
		}
		if shape, found := rangeValueTextLiteralRangeShape(args); found {
			return shape, true
		}
		return rangeShape{}, true
	case "offset", "resize":
		return rangeShape{}, true
	default:
		return rangeShape{}, false
	}
}

func rangeValueTextCellsPairShape(arguments []string, facts rangeValueFacts) (rangeShape, bool) {
	if len(arguments) != 2 {
		return rangeShape{}, false
	}
	startRow, startRowKnown, startColumn, startColumnKnown, startOK := rangeValueTextCellsCoordinates(arguments[0], facts)
	endRow, endRowKnown, endColumn, endColumnKnown, endOK := rangeValueTextCellsCoordinates(arguments[1], facts)
	if !startOK || !endOK {
		return rangeShape{}, false
	}
	rows, rowsKnown := rangeValueAxisLength(startRow, startRowKnown, endRow, endRowKnown)
	cols, colsKnown := rangeValueAxisLength(startColumn, startColumnKnown, endColumn, endColumnKnown)
	return rangeShape{
		known:   rowsKnown && colsKnown,
		array2D: (rowsKnown && rows > 1) || (colsKnown && cols > 1),
		rows:    rows,
		cols:    cols,
	}, true
}

func rangeValueTextCellsCoordinates(expression string, facts rangeValueFacts) (int, bool, int, bool, bool) {
	text := strings.TrimSpace(expression)
	open := firstParenOutsideString(text)
	if open < 0 {
		return 0, false, 0, false, false
	}
	close := matchingParen(text, open)
	if close < 0 || strings.TrimSpace(text[close+1:]) != "" {
		return 0, false, 0, false, false
	}
	name := strings.ToLower(cleanIdentifier(lastName(strings.TrimSpace(text[:open]))))
	if name != "cells" {
		return 0, false, 0, false, false
	}
	arguments := splitArgs(text[open+1 : close])
	if len(arguments) != 2 {
		return 0, false, 0, false, false
	}
	row, rowErr := constantIntegerExpression(arguments[0], facts.constantValues())
	column, columnErr := constantIntegerExpression(arguments[1], facts.constantValues())
	return row, rowErr == nil, column, columnErr == nil, true
}

func rangeValueTextLiteralRangeShape(arguments []string) (rangeShape, bool) {
	if len(arguments) != 1 && len(arguments) != 2 {
		return rangeShape{}, false
	}
	addresses := make([]string, len(arguments))
	for index, argument := range arguments {
		text := strings.TrimSpace(argument)
		if len(text) < 2 || text[0] != '"' || text[len(text)-1] != '"' {
			return rangeShape{}, false
		}
		addresses[index] = strings.ReplaceAll(text[1:len(text)-1], `""`, `"`)
	}
	if len(addresses) == 1 {
		match := rangeValueAddressRe.FindStringSubmatch(addresses[0])
		if len(match) != 5 {
			return rangeShape{}, false
		}
		startColumn := excelColumnNumber(match[1])
		startRow, _ := strconv.Atoi(match[2])
		endColumn, endRow := startColumn, startRow
		if match[3] != "" {
			endColumn = excelColumnNumber(match[3])
			endRow, _ = strconv.Atoi(match[4])
		}
		if startColumn <= 0 || endColumn < startColumn || endRow < startRow {
			return rangeShape{}, false
		}
		return rangeShape{known: true, array2D: endColumn > startColumn || endRow > startRow, rows: endRow - startRow + 1, cols: endColumn - startColumn + 1}, true
	}
	startMatch := rangeValueAddressRe.FindStringSubmatch(addresses[0])
	endMatch := rangeValueAddressRe.FindStringSubmatch(addresses[1])
	if len(startMatch) != 5 || len(endMatch) != 5 || startMatch[3] != "" || endMatch[3] != "" {
		return rangeShape{}, false
	}
	startColumn := excelColumnNumber(startMatch[1])
	startRow, _ := strconv.Atoi(startMatch[2])
	endColumn := excelColumnNumber(endMatch[1])
	endRow, _ := strconv.Atoi(endMatch[2])
	if startColumn <= 0 || endColumn < startColumn || endRow < startRow {
		return rangeShape{}, false
	}
	return rangeShape{known: true, array2D: endColumn > startColumn || endRow > startRow, rows: endRow - startRow + 1, cols: endColumn - startColumn + 1}, true
}

func rangeValueMemberReceiver(expression *procedureir.Expression, facts rangeValueFacts, members ...string) (procedureir.Expression, bool) {
	current, ok := rangeValueUnwrapExpression(expression, facts)
	if !ok || current.Kind != procedureir.ExpressionMember || len(current.Children) < 2 {
		return procedureir.Expression{}, false
	}
	member, ok := facts.expressions[current.Children[len(current.Children)-1]]
	if !ok {
		return procedureir.Expression{}, false
	}
	for _, candidate := range members {
		if strings.EqualFold(strings.TrimSpace(member.Text), candidate) {
			receiver, found := facts.expressions[current.Children[0]]
			return receiver, found
		}
	}
	return procedureir.Expression{}, false
}

func rangeValueUnwrapExpression(expression *procedureir.Expression, facts rangeValueFacts) (procedureir.Expression, bool) {
	if expression == nil {
		return procedureir.Expression{}, false
	}
	current := *expression
	for current.Kind == procedureir.ExpressionParentheses && len(current.Children) == 1 {
		next, ok := facts.expressions[current.Children[0]]
		if !ok {
			return procedureir.Expression{}, false
		}
		current = next
	}
	return current, true
}

func rangeValueRangeShapeExpression(expression *procedureir.Expression, state rangeValueFlowState, facts rangeValueFacts) (rangeShape, bool) {
	current, ok := rangeValueUnwrapExpression(expression, facts)
	if !ok {
		return rangeShape{}, false
	}
	if current.Kind == procedureir.ExpressionIdentifier {
		name := strings.ToLower(strings.TrimSpace(current.Text))
		if shape, found := state.ranges[name]; found {
			return shape, true
		}
		if facts.rangeVariables[name] {
			return rangeShape{}, true
		}
		return rangeShape{}, false
	}
	name, arguments, ok := rangeValueDirectCall(current, facts)
	if !ok {
		return rangeShape{}, false
	}
	switch strings.ToLower(name) {
	case "cells":
		return rangeShape{known: true, rows: 1, cols: 1}, true
	case "range":
		if shape, found := rangeValueLiteralRangeShape(arguments, facts); found {
			return shape, true
		}
		if len(arguments) == 2 {
			if shape, found := rangeValueCellsPairShape(arguments[0], arguments[1], facts); found {
				return shape, true
			}
		}
		return rangeShape{}, true
	case "offset", "resize":
		return rangeShape{}, true
	default:
		return rangeShape{}, false
	}
}

func rangeValueDirectCall(expression procedureir.Expression, facts rangeValueFacts) (string, []procedureir.Expression, bool) {
	if expression.Kind != procedureir.ExpressionCall || len(expression.Children) == 0 {
		return "", nil, false
	}
	callee, ok := facts.expressions[expression.Children[0]]
	if !ok {
		return "", nil, false
	}
	name, ok := rangeValueTerminalExpressionName(callee, facts)
	if !ok {
		return "", nil, false
	}
	arguments := make([]procedureir.Expression, 0, len(expression.Children)-1)
	for _, id := range expression.Children[1:] {
		argument, found := facts.expressions[id]
		if !found {
			return "", nil, false
		}
		arguments = append(arguments, argument)
	}
	return name, arguments, true
}

func rangeValueTerminalExpressionName(expression procedureir.Expression, facts rangeValueFacts) (string, bool) {
	current, ok := rangeValueUnwrapExpression(&expression, facts)
	if !ok {
		return "", false
	}
	if current.Kind == procedureir.ExpressionIdentifier {
		return strings.TrimSpace(current.Text), true
	}
	if current.Kind != procedureir.ExpressionMember || len(current.Children) < 2 {
		return "", false
	}
	member, ok := facts.expressions[current.Children[len(current.Children)-1]]
	if !ok || member.Kind != procedureir.ExpressionIdentifier {
		return "", false
	}
	return strings.TrimSpace(member.Text), true
}

func rangeValueLiteralRangeShape(arguments []procedureir.Expression, facts rangeValueFacts) (rangeShape, bool) {
	if len(arguments) == 1 {
		address, ok := rangeValueStringLiteral(arguments[0], facts)
		if !ok {
			return rangeShape{}, false
		}
		match := rangeValueAddressRe.FindStringSubmatch(address)
		if len(match) != 5 {
			return rangeShape{}, false
		}
		startCol := excelColumnNumber(match[1])
		startRow, _ := strconv.Atoi(match[2])
		endCol, endRow := startCol, startRow
		if match[3] != "" {
			endCol = excelColumnNumber(match[3])
			endRow, _ = strconv.Atoi(match[4])
		}
		if startCol > 0 && endCol >= startCol && endRow >= startRow {
			return rangeShape{known: true, array2D: endCol > startCol || endRow > startRow, rows: endRow - startRow + 1, cols: endCol - startCol + 1}, true
		}
		return rangeShape{}, false
	}
	if len(arguments) == 2 {
		startAddress, startOK := rangeValueStringLiteral(arguments[0], facts)
		endAddress, endOK := rangeValueStringLiteral(arguments[1], facts)
		if !startOK || !endOK {
			return rangeShape{}, false
		}
		startMatch := rangeValueAddressRe.FindStringSubmatch(startAddress)
		endMatch := rangeValueAddressRe.FindStringSubmatch(endAddress)
		if len(startMatch) != 5 || len(endMatch) != 5 || startMatch[3] != "" || endMatch[3] != "" {
			return rangeShape{}, false
		}
		startCol := excelColumnNumber(startMatch[1])
		startRow, _ := strconv.Atoi(startMatch[2])
		endCol := excelColumnNumber(endMatch[1])
		endRow, _ := strconv.Atoi(endMatch[2])
		if startCol > 0 && endCol >= startCol && endRow >= startRow {
			return rangeShape{known: true, array2D: endCol > startCol || endRow > startRow, rows: endRow - startRow + 1, cols: endCol - startCol + 1}, true
		}
	}
	return rangeShape{}, false
}

func rangeValueStringLiteral(expression procedureir.Expression, facts rangeValueFacts) (string, bool) {
	current, ok := rangeValueUnwrapExpression(&expression, facts)
	if !ok || current.Kind != procedureir.ExpressionLiteral {
		return "", false
	}
	text := strings.TrimSpace(current.Text)
	if len(text) < 2 || text[0] != '"' || text[len(text)-1] != '"' {
		return "", false
	}
	return strings.ReplaceAll(text[1:len(text)-1], `""`, `"`), true
}

func rangeValueCellsPairShape(start, end procedureir.Expression, facts rangeValueFacts) (rangeShape, bool) {
	startRow, startRowKnown, startCol, startColKnown, startOK := rangeValueCellsCoordinates(start, facts)
	endRow, endRowKnown, endCol, endColKnown, endOK := rangeValueCellsCoordinates(end, facts)
	if !startOK || !endOK {
		return rangeShape{}, false
	}
	rows, rowsKnown := rangeValueAxisLength(startRow, startRowKnown, endRow, endRowKnown)
	cols, colsKnown := rangeValueAxisLength(startCol, startColKnown, endCol, endColKnown)
	return rangeShape{
		known:   rowsKnown && colsKnown,
		array2D: (rowsKnown && rows > 1) || (colsKnown && cols > 1),
		rows:    rows,
		cols:    cols,
	}, true
}

func rangeValueCellsCoordinates(expression procedureir.Expression, facts rangeValueFacts) (int, bool, int, bool, bool) {
	current, ok := rangeValueUnwrapExpression(&expression, facts)
	if !ok {
		return 0, false, 0, false, false
	}
	name, arguments, ok := rangeValueDirectCall(current, facts)
	if !ok || !strings.EqualFold(name, "cells") || len(arguments) != 2 {
		return 0, false, 0, false, false
	}
	row, rowErr := constantIntegerExpression(arguments[0].Text, facts.constantValues())
	column, columnErr := constantIntegerExpression(arguments[1].Text, facts.constantValues())
	return row, rowErr == nil, column, columnErr == nil, true
}

func rangeValueAxisLength(start int, startKnown bool, end int, endKnown bool) (int, bool) {
	if !startKnown || !endKnown {
		return 0, false
	}
	if start > end {
		start, end = end, start
	}
	return end - start + 1, true
}

func rangeValueDestinationIssue(target procedureir.Expression, source rangeValueShape, state rangeValueFlowState, facts rangeValueFacts) *rangeValueIssue {
	destination, recognized := rangeValueRangeShapeExpression(&target, state, facts)
	if !recognized || !destination.known {
		return nil
	}
	return rangeValueDestinationShapeIssue(destination, source)
}

func rangeValueDestinationShapeIssue(destination rangeShape, source rangeValueShape) *rangeValueIssue {
	if source.kind != rangeValueShapeArray2D || source.rows <= 0 || source.cols <= 0 || !destination.known {
		return nil
	}
	if destination.rows == source.rows && destination.cols == source.cols {
		return nil
	}
	return &rangeValueIssue{
		message:    fmt.Sprintf("Two-dimensional Range.Value array (%d x %d) is assigned to an incompatible range shape (%d x %d).", source.rows, source.cols, destination.rows, destination.cols),
		reason:     "Excel Range.Value assignments require the destination shape to match the two-dimensional array when the values are transferred as a block.",
		suggestion: fmt.Sprintf("Resize the destination to %d rows by %d columns, or construct an array with the destination's shape.", source.rows, source.cols),
	}
}

func applyRangeValueGuard(state rangeValueFlowState, statement *procedureir.Statement, edge vbacfg.Edge) {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return
	}
	match := rangeValueGuardRe.FindStringSubmatch(rangeValueStatementText(*statement))
	if len(match) != 3 {
		return
	}
	negated := strings.TrimSpace(match[1]) != ""
	trueBranch := edge.Kind == vbacfg.EdgeBranchTrue
	arrayBranch := (trueBranch && !negated) || (!trueBranch && negated)
	name := strings.ToLower(match[2])
	if arrayBranch {
		state.arrayGuards[name] = true
	} else {
		delete(state.arrayGuards, name)
	}
}

func deduplicateRangeValueIssues(issues []rangeValueIssue) []rangeValueIssue {
	seen := map[string]bool{}
	out := make([]rangeValueIssue, 0, len(issues))
	for _, issue := range issues {
		if seen[issue.message] {
			continue
		}
		seen[issue.message] = true
		out = append(out, issue)
	}
	return out
}
