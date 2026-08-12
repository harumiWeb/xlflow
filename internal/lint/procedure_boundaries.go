package lint

import (
	"sort"
	"strconv"
	"strings"
)

// procedureBoundaryState is one possible source interpretation while a
// conditional-compilation group is being scanned.  VBA compiles exactly one
// branch, so each branch must get its own procedure stack; merging the raw
// source lines into one stack makes mutually-exclusive declarations appear to
// nest.
type procedureBoundaryState struct {
	procedures      []procedureFrame
	declarationKind string
}

type procedureBoundaryStates map[string]procedureBoundaryState

type conditionalProcedureFrame struct {
	baseline procedureBoundaryStates
	branches procedureBoundaryStates
	sawElse  bool
}

type procedureBoundaryTracker struct {
	linter       Linter
	path         string
	states       procedureBoundaryStates
	conditionals []conditionalProcedureFrame
	issues       []Issue
	issueKeys    map[string]struct{}
	ambiguous    bool
}

const maxProcedureBoundaryStates = 512

func newProcedureBoundaryTracker(linter Linter, path string) *procedureBoundaryTracker {
	tracker := &procedureBoundaryTracker{
		linter:    linter,
		path:      path,
		states:    make(procedureBoundaryStates),
		issueKeys: make(map[string]struct{}),
	}
	tracker.addState(tracker.states, procedureBoundaryState{})
	return tracker
}

func (t *procedureBoundaryTracker) processDirective(lineNo int, lower string) {
	switch {
	case strings.HasPrefix(lower, "#if ") && strings.HasSuffix(lower, " then"):
		baseline := cloneProcedureBoundaryStates(t.states)
		t.conditionals = append(t.conditionals, conditionalProcedureFrame{
			baseline: baseline,
			branches: make(procedureBoundaryStates),
		})
		t.states = cloneProcedureBoundaryStates(baseline)
	case strings.HasPrefix(lower, "#elseif ") && strings.HasSuffix(lower, " then"):
		if len(t.conditionals) == 0 {
			t.ambiguous = true
			return
		}
		frame := &t.conditionals[len(t.conditionals)-1]
		if frame.sawElse {
			t.ambiguous = true
			return
		}
		t.mergeStates(frame.branches, t.states)
		t.states = cloneProcedureBoundaryStates(frame.baseline)
	case lower == "#else":
		if len(t.conditionals) == 0 {
			t.ambiguous = true
			return
		}
		frame := &t.conditionals[len(t.conditionals)-1]
		if frame.sawElse {
			t.ambiguous = true
			return
		}
		t.mergeStates(frame.branches, t.states)
		frame.sawElse = true
		t.states = cloneProcedureBoundaryStates(frame.baseline)
	case lower == "#end if":
		if len(t.conditionals) == 0 {
			t.ambiguous = true
			return
		}
		frame := t.conditionals[len(t.conditionals)-1]
		t.conditionals = t.conditionals[:len(t.conditionals)-1]
		t.mergeStates(frame.branches, t.states)
		if !frame.sawElse {
			t.mergeStates(frame.branches, frame.baseline)
		}
		t.states = frame.branches
	default:
		_ = lineNo
	}
}

func (t *procedureBoundaryTracker) processStatement(lineNo int, line string) {
	if t.ambiguous {
		return
	}
	lower := strings.ToLower(strings.TrimSpace(line))
	if lower == "" {
		return
	}

	next := make(procedureBoundaryStates)
	for _, key := range sortedProcedureBoundaryStateKeys(t.states) {
		state := t.states[key]
		if state.declarationKind != "" {
			switch {
			case isTypeDeclarationEnd(lower) && state.declarationKind == "type":
				state.declarationKind = ""
			case isEnumDeclarationEnd(lower) && state.declarationKind == "enum":
				state.declarationKind = ""
			case isTypeDeclarationStart(lower):
				// A malformed nested declaration is left to the parser and the
				// generic syntax diagnostics; do not interpret its contents as a
				// procedure boundary.
			default:
				if endKind, ok := procedureEndKind(lower); ok {
					t.addBoundaryIssue(t.linter.issue(t.path, lineNo, "VB011", "error", "Unexpected End "+endKind+" without a matching procedure."))
					t.addState(next, state)
					continue
				}
				// A reserved procedure keyword used as a declaration member has
				// no procedure name. Keep the historical VB010 signal for that
				// invalid source instead of silently dropping the only finding.
				if start, ok := procedureStart(line, lineNo); ok && start.Name == "" {
					t.addBoundaryIssue(t.linter.issue(t.path, lineNo, "VB010", "error", "Unterminated "+start.Kind+" procedure."))
				}
				t.addState(next, state)
				continue
			}
			t.addState(next, state)
			continue
		}

		if isTypeDeclarationStart(lower) {
			state.declarationKind = "type"
			t.addState(next, state)
			continue
		}
		if isEnumDeclarationStart(lower) {
			state.declarationKind = "enum"
			t.addState(next, state)
			continue
		}
		if isTypeDeclarationEnd(lower) || isEnumDeclarationEnd(lower) {
			// An unmatched declaration terminator is handled by parser
			// recovery, not by procedure-boundary diagnostics.
			t.addState(next, state)
			continue
		}
		if endKind, ok := procedureEndKind(lower); ok {
			if len(state.procedures) == 0 {
				t.addBoundaryIssue(t.linter.issue(t.path, lineNo, "VB011", "error", "Unexpected End "+endKind+" without a matching procedure."))
				t.addState(next, state)
				continue
			}
			top := state.procedures[len(state.procedures)-1]
			state.procedures = state.procedures[:len(state.procedures)-1]
			if top.Kind != endKind {
				issue := t.linter.issue(t.path, lineNo, "VB012", "error", "Mismatched End "+endKind+" for "+top.Kind+" procedure.")
				issue.Symbol = top.Name
				t.addBoundaryIssue(issue)
			}
			t.addState(next, state)
			continue
		}
		if start, ok := procedureStart(line, lineNo); ok {
			state.procedures = append(state.procedures, start)
			t.addState(next, state)
			continue
		}
		{
			t.addState(next, state)
		}
	}
	t.states = next
}

func (t *procedureBoundaryTracker) finish() []Issue {
	if t.ambiguous || len(t.conditionals) != 0 {
		return t.issues
	}
	for _, key := range sortedProcedureBoundaryStateKeys(t.states) {
		state := t.states[key]
		for _, procedure := range state.procedures {
			issue := t.linter.issue(t.path, procedure.LineNo, "VB010", "error", "Unterminated "+procedure.Kind+" procedure.")
			issue.Symbol = procedure.Name
			t.addBoundaryIssue(issue)
		}
	}
	return t.issues
}

func (t *procedureBoundaryTracker) addBoundaryIssue(issue Issue) {
	key := strings.Join([]string{
		issue.Code,
		strconv.Itoa(issue.Line),
		strconv.Itoa(issue.Column),
		issue.Symbol,
		issue.Message,
	}, "\x1f")
	if _, exists := t.issueKeys[key]; exists {
		return
	}
	t.issueKeys[key] = struct{}{}
	t.issues = append(t.issues, issue)
}

func (t *procedureBoundaryTracker) addState(states procedureBoundaryStates, state procedureBoundaryState) {
	key := procedureBoundaryStateKey(state)
	if _, exists := states[key]; exists {
		return
	}
	if len(states) >= maxProcedureBoundaryStates {
		t.ambiguous = true
		return
	}
	state.procedures = append([]procedureFrame(nil), state.procedures...)
	states[key] = state
}

func (t *procedureBoundaryTracker) addStates(target, source procedureBoundaryStates) {
	for _, key := range sortedProcedureBoundaryStateKeys(source) {
		t.addState(target, source[key])
	}
}

func (t *procedureBoundaryTracker) mergeStates(target, source procedureBoundaryStates) {
	t.addStates(target, source)
}

func cloneProcedureBoundaryStates(source procedureBoundaryStates) procedureBoundaryStates {
	cloned := make(procedureBoundaryStates, len(source))
	for _, key := range sortedProcedureBoundaryStateKeys(source) {
		state := source[key]
		state.procedures = append([]procedureFrame(nil), state.procedures...)
		cloned[key] = state
	}
	return cloned
}

func sortedProcedureBoundaryStateKeys(states procedureBoundaryStates) []string {
	keys := make([]string, 0, len(states))
	for key := range states {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

func procedureBoundaryStateKey(state procedureBoundaryState) string {
	parts := []string{state.declarationKind}
	for _, procedure := range state.procedures {
		parts = append(parts, strings.Join([]string{
			procedure.Kind,
			procedure.Name,
			strconv.Itoa(procedure.LineNo),
		}, "\x1e"))
	}
	return strings.Join(parts, "\x1f")
}

func isEnumDeclarationStart(lower string) bool {
	fields := strings.Fields(lower)
	if len(fields) < 2 {
		return false
	}
	index := 0
	for index < len(fields) {
		switch fields[index] {
		case "public", "private", "friend":
			index++
		default:
			return fields[index] == "enum" && index+1 < len(fields)
		}
	}
	return false
}

func isEnumDeclarationEnd(lower string) bool {
	fields := strings.Fields(lower)
	return len(fields) == 2 && fields[0] == "end" && fields[1] == "enum"
}
