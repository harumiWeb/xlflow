package lint

import (
	"sort"
	"strconv"
	"strings"
)

// ProcedureBoundaryOutcome records the evidence-backed interpretation of a
// mismatched procedure terminator.  A zero-valued ProcedureBoundaryDecision
// is treated as compile-invalid so a nil classifier and a classifier that only
// fills the context both retain the historical VB012 behavior.
type ProcedureBoundaryOutcome string

const (
	// ProcedureBoundaryCompileInvalid means that the host/compiler rejects the
	// opener/terminator combination.  It is reported as VB012/error by default.
	ProcedureBoundaryCompileInvalid ProcedureBoundaryOutcome = "compile-invalid"
	// ProcedureBoundaryStyleMismatch means that the host accepts the source,
	// while the project may still want a separate style diagnostic.  A custom
	// diagnostic code can be supplied in ProcedureBoundaryDecision.  The
	// production classifier supplies the registry-backed VB066 style code.
	ProcedureBoundaryStyleMismatch ProcedureBoundaryOutcome = "style-mismatch"
	// ProcedureBoundaryAccepted suppresses a structural mismatch finding when
	// host evidence confirms that the source compiles.  Parser recovery remains
	// a separate concern and is not altered by this decision.
	ProcedureBoundaryAccepted ProcedureBoundaryOutcome = "accepted"
)

// ProcedureBoundaryContext is the normalized procedure opener and terminator
// pair presented to ProcedureBoundaryClassifier.  Property accessor identity
// is intentionally retained: VBA's Property Get/Let/Set forms share the
// Property terminator token but may have different host behavior.
type ProcedureBoundaryContext struct {
	ModuleKind     string
	OpenerKind     string
	Accessor       string
	TerminatorKind string
	ProcedureName  string
	OpeningLine    int
	ClosingLine    int
}

// ProcedureBoundaryDecision is the result of an evidence-backed
// classification. Code and Severity are only needed for style diagnostics;
// compile-invalid and accepted outcomes use VB012/error and no issue,
// respectively, when they are left empty. The production table supplies the
// registry-backed VB066 style code for the confirmed accepted mismatch.
type ProcedureBoundaryDecision struct {
	Outcome  ProcedureBoundaryOutcome
	Code     string
	Severity string
	Message  string
}

// ProcedureBoundaryClassifier lets tests and host-specific evidence tables
// classify a procedure terminator without coupling the parser's structural
// stack to a particular Office/VBE result.  The production tracker installs
// DefaultProcedureBoundaryClassifier when callers do not provide an override.
type ProcedureBoundaryClassifier func(ProcedureBoundaryContext) ProcedureBoundaryDecision

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
	classifier   ProcedureBoundaryClassifier
	path         string
	moduleKind   string
	states       procedureBoundaryStates
	conditionals []conditionalProcedureFrame
	issues       []Issue
	issueKeys    map[string]struct{}
	ambiguous    bool
}

const maxProcedureBoundaryStates = 512

func newProcedureBoundaryTracker(linter Linter, path string) *procedureBoundaryTracker {
	classifier := linter.ProcedureBoundaryClassifier
	if classifier == nil {
		classifier = DefaultProcedureBoundaryClassifier
	}
	tracker := &procedureBoundaryTracker{
		linter:     linter,
		classifier: classifier,
		path:       path,
		moduleKind: strings.ToLower(strings.TrimSpace(linter.moduleKindForPath(path))),
		states:     make(procedureBoundaryStates),
		issueKeys:  make(map[string]struct{}),
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
				if issue := t.mismatchedTerminatorIssue(top, endKind, lineNo); issue != nil {
					t.addBoundaryIssue(*issue)
				}
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

func (t *procedureBoundaryTracker) mismatchedTerminatorIssue(top procedureFrame, endKind string, lineNo int) *Issue {
	decision := ProcedureBoundaryDecision{Outcome: ProcedureBoundaryCompileInvalid}
	if t.classifier != nil {
		decision = t.classifier(ProcedureBoundaryContext{
			ModuleKind:     t.moduleKind,
			OpenerKind:     top.Kind,
			Accessor:       top.Accessor,
			TerminatorKind: endKind,
			ProcedureName:  top.Name,
			OpeningLine:    top.LineNo,
			ClosingLine:    lineNo,
		})
	}

	switch decision.Outcome {
	case ProcedureBoundaryAccepted:
		return nil
	case ProcedureBoundaryStyleMismatch:
		code := strings.TrimSpace(decision.Code)
		if code == "" {
			// A style decision without a registry code is intentionally
			// non-emitting.  The production classifier always supplies VB066;
			// this fallback keeps custom experiments from inventing an
			// unregistered diagnostic.
			return nil
		}
		severity := strings.TrimSpace(decision.Severity)
		if severity == "" {
			severity = "warning"
		}
		message := decision.Message
		if strings.TrimSpace(message) == "" {
			message = "Mismatched End " + endKind + " for " + top.Kind + " procedure."
		}
		issue := t.linter.issue(t.path, lineNo, code, severity, message)
		issue.Symbol = top.Name
		return &issue
	default:
		message := decision.Message
		if strings.TrimSpace(message) == "" {
			message = "Mismatched End " + endKind + " for " + top.Kind + " procedure."
		}
		issue := t.linter.issue(t.path, lineNo, "VB012", "error", message)
		issue.Symbol = top.Name
		return &issue
	}
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
			procedure.Accessor,
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
