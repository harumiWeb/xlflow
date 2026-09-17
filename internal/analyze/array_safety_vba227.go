package analyze

import (
	"regexp"
	"slices"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type arrayVBA227ResumeNextEdges map[vbacfg.BlockID]map[vbacfg.BlockID]bool

const arrayVBA227DirectErrNumberGuard = "\x00direct-err-number"

type arrayVBA227ResumeNextFailureGuard struct {
	target         string
	assignmentText string
	source         string
	successOnTrue  bool
}

type arrayVBA227ResumeNextFailureConditionInfo struct {
	source        string
	successOnTrue bool
	polarityKnown bool
	negated       bool
}

type arrayVBA227ResumeFacts struct {
	hasResumeTransfer                      bool
	boundsByName                           map[string][]procedureir.Statement
	resumeNextContinuations                map[vbacfg.BlockID][]vbacfg.BlockID
	resumeNextFailureGuardsByStatement     map[int]arrayVBA227ResumeNextFailureGuard
	resumeNextFailureGuardStatementsByLine map[int][]int
}

func buildArrayVBA227ResumeFacts(proc sourceProcedure) *arrayVBA227ResumeFacts {
	facts := &arrayVBA227ResumeFacts{
		boundsByName:            map[string][]procedureir.Statement{},
		resumeNextContinuations: map[vbacfg.BlockID][]vbacfg.BlockID{},
	}
	facts.resumeNextFailureGuardsByStatement, facts.resumeNextFailureGuardStatementsByLine = buildArrayVBA227ResumeNextFailureGuards(proc)
	for statement := range proc.Statements.All() {
		if statement.Kind == procedureir.StatementResume && statement.Control != nil {
			switch statement.Control.Transfer {
			case procedureir.TransferResumeNext, procedureir.TransferResumeLabel:
				facts.hasResumeTransfer = true
			}
		}
		for _, match := range arrayBoundCallRe.FindAllStringSubmatch(statement.Text, -1) {
			name := strings.ToLower(cleanIdentifier(match[2]))
			if name != "" {
				facts.boundsByName[name] = append(facts.boundsByName[name], statement)
			}
		}
	}
	if proc.Graph != nil && facts.hasResumeTransfer {
		facts.resumeNextContinuations = arrayVBA227ResumeNextContinuations(proc.Graph.View(vbacfg.EdgeFilter{}))
	}
	return facts
}

func buildArrayVBA227ResumeNextFailureGuards(proc sourceProcedure) (map[int]arrayVBA227ResumeNextFailureGuard, map[int][]int) {
	var guards map[int]arrayVBA227ResumeNextFailureGuard
	var statementIDsByLine map[int][]int
	if proc.Graph == nil {
		return nil, nil
	}
	// ProcedureIR statements are emitted in source order. Scan that immutable
	// order directly so procedures without a Range.Value/Value2 candidate do
	// not pay for a temporary slice or an O(S log S) sort.
	var pendingRange procedureir.Statement
	pendingRangeTarget := ""
	hasPendingRange := false
	var pendingFlag *arrayVBA227ResumeNextFailureGuard
	var pendingFlagStatement procedureir.Statement
	hasPendingFlag := false
	for statement := range proc.Statements.All() {
		if target := arrayVBA227RangeValueAssignment(statement.Text); target != "" {
			pendingRange = statement
			pendingRangeTarget = target
			hasPendingRange = true
			pendingFlag = nil
			hasPendingFlag = false
			continue
		}
		if !hasPendingRange {
			continue
		}
		if source, successOnTrue, ok := arrayVBA227ResumeNextFailureFlagAssignment(statement.Text); ok {
			if arrayVBA227StatementDominates(proc, pendingRange, statement) {
				pendingFlag = &arrayVBA227ResumeNextFailureGuard{source: source, successOnTrue: successOnTrue}
				pendingFlagStatement = statement
				hasPendingFlag = true
			} else {
				hasPendingRange = false
				pendingRange = procedureir.Statement{}
				pendingRangeTarget = ""
				pendingFlag = nil
				hasPendingFlag = false
			}
			continue
		}
		if !hasPendingFlag && arrayVBA227ResumeNextErrObservationAssignment(statement.Text) {
			// Reading Err.Number/Err.Description into a diagnostic string does
			// not replace the status that the following Boolean capture observes.
			// Keep this narrow: arbitrary assignments or calls still invalidate
			// the candidate before its failure flag is captured.
			continue
		}
		if hasPendingFlag && arrayVBA227ResumeNextCapturedStatusTransition(statement.Text) {
			// These conventional transitions do not replace the Boolean value
			// that was just captured from Err.Number. Keep the candidate through
			// them, but still require the next executable statement to be the
			// matching condition.
			continue
		}
		condition, ok := arrayVBA227ResumeNextFailureCondition(statement.Text)
		if !ok {
			// Err.Clear, another assignment, a call, or any other statement
			// can change the status captured by Err.Number. A guard is valid
			// only for the immediately following flag/condition sequence.
			hasPendingRange = false
			pendingRange = procedureir.Statement{}
			pendingRangeTarget = ""
			pendingFlag = nil
			hasPendingFlag = false
			continue
		}
		if hasPendingFlag && pendingFlag != nil {
			if condition.source == pendingFlag.source &&
				(!condition.polarityKnown || condition.successOnTrue == pendingFlag.successOnTrue) &&
				arrayVBA227StatementDominates(proc, pendingFlagStatement, statement) {
				if guards == nil {
					guards = make(map[int]arrayVBA227ResumeNextFailureGuard)
					statementIDsByLine = make(map[int][]int)
				}
				guard := *pendingFlag
				guard.target = pendingRangeTarget
				guard.assignmentText = strings.TrimSpace(pendingRange.Text)
				guards[pendingRange.ID] = guard
				statementIDsByLine[pendingRange.Range.StartLine] = append(statementIDsByLine[pendingRange.Range.StartLine], pendingRange.ID)
			}
		} else if condition.source == arrayVBA227DirectErrNumberGuard && arrayVBA227ResumeNextFailureConditionTerminates(statement.Text) && arrayVBA227StatementDominates(proc, pendingRange, statement) {
			if guards == nil {
				guards = make(map[int]arrayVBA227ResumeNextFailureGuard)
				statementIDsByLine = make(map[int][]int)
			}
			guards[pendingRange.ID] = arrayVBA227ResumeNextFailureGuard{
				target:         pendingRangeTarget,
				assignmentText: strings.TrimSpace(pendingRange.Text),
				source:         condition.source,
				successOnTrue:  condition.successOnTrue,
			}
			statementIDsByLine[pendingRange.Range.StartLine] = append(statementIDsByLine[pendingRange.Range.StartLine], pendingRange.ID)
		}
		hasPendingRange = false
		pendingRange = procedureir.Statement{}
		pendingRangeTarget = ""
		pendingFlag = nil
		hasPendingFlag = false
	}
	return guards, statementIDsByLine
}

func arrayVBA227ResumeNextCapturedStatusTransition(text string) bool {
	switch strings.ToLower(strings.TrimSpace(text)) {
	case "on error goto 0", "on error goto -1", "err.clear":
		return true
	default:
		return false
	}
}

func arrayVBA227ResumeNextErrObservationAssignment(text string) bool {
	lhs, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed || !errorSuccessIdentifierRE.MatchString(strings.TrimSpace(lhs)) {
		return false
	}
	hasErrObservation := false
	for _, term := range arrayVBA227TopLevelConcatenationTerms(rhs) {
		normalized := strings.ToLower(strings.Join(strings.Fields(strings.TrimSpace(term)), ""))
		switch {
		case resumeNextScopeExactErrNumber(term):
			hasErrObservation = true
		case normalized == "cstr(err.number)":
			hasErrObservation = true
		case resumeNextScopeErrDescriptionExpression(term):
			hasErrObservation = true
		case resumeNextScopeMaskedStringLiteral(strings.TrimSpace(term)):
		default:
			return false
		}
	}
	return hasErrObservation
}

func arrayVBA227TopLevelConcatenationTerms(text string) []string {
	terms := make([]string, 0, 2)
	start := 0
	depth := 0
	inString := false
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
		case '&':
			if !inString && depth == 0 {
				terms = append(terms, text[start:index])
				start = index + 1
			}
		}
	}
	return append(terms, text[start:])
}

func arrayVBA227ResumeNextFailureConditionTerminates(text string) bool {
	_, body, ok := arrayIfThenParts(text)
	if !ok {
		return false
	}
	body = strings.ToLower(strings.TrimSpace(body))
	for _, exitStatement := range []string{"exit sub", "exit function", "exit property"} {
		if body == exitStatement || strings.HasPrefix(body, exitStatement+" ") {
			return true
		}
	}
	return false
}

func arrayVBA227RangeValueAssignment(text string) string {
	lhs, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return ""
	}
	lower := strings.ToLower(strings.TrimSpace(rhs))
	if !strings.Contains(lower, ".value") && !strings.Contains(lower, ".value2") {
		return ""
	}
	return strings.ToLower(cleanIdentifier(lhs))
}

func arrayVBA227ResumeNextFailureFlagAssignment(text string) (string, bool, bool) {
	lhs, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return "", false, false
	}
	successOnTrue, ok := arrayVBA227ErrNumberGuardExpression(rhs)
	if !ok {
		return "", false, false
	}
	return strings.ToLower(cleanIdentifier(lhs)), successOnTrue, true
}

func arrayVBA227ErrNumberGuardExpression(text string) (bool, bool) {
	left, operator, right, ok := resumeNextScopeComparison(text)
	if !ok || strings.TrimSpace(right) == "" {
		return false, false
	}
	if resumeNextScopeExactErrNumber(left) && strings.TrimSpace(right) == "0" ||
		resumeNextScopeExactErrNumber(right) && strings.TrimSpace(left) == "0" {
		switch operator {
		case "=":
			return true, true
		case "<>":
			return false, true
		}
	}
	return false, false
}

func arrayVBA227ResumeNextFailureCondition(text string) (arrayVBA227ResumeNextFailureConditionInfo, bool) {
	condition := text
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	condition = strings.TrimSpace(condition)
	lower := strings.ToLower(condition)
	if strings.HasPrefix(lower, "if ") {
		condition = strings.TrimSpace(condition[len("if "):])
	} else if strings.HasPrefix(lower, "elseif ") {
		condition = strings.TrimSpace(condition[len("elseif "):])
	}
	if then := arrayTopLevelKeywordIndex(condition, "then"); then >= 0 {
		condition = strings.TrimSpace(condition[:then])
	}
	for len(condition) >= 2 && condition[0] == '(' && condition[len(condition)-1] == ')' {
		condition = strings.TrimSpace(condition[1 : len(condition)-1])
	}
	if successOnTrue, ok := arrayVBA227ErrNumberGuardExpression(condition); ok {
		return arrayVBA227ResumeNextFailureConditionInfo{
			source:        arrayVBA227DirectErrNumberGuard,
			successOnTrue: successOnTrue,
			polarityKnown: true,
		}, true
	}
	negated := false
	lower = strings.ToLower(condition)
	if strings.HasPrefix(lower, "not ") {
		condition = strings.TrimSpace(condition[len("not "):])
		negated = true
	}
	if !arrayEraseNameRe.MatchString(condition) {
		return arrayVBA227ResumeNextFailureConditionInfo{}, false
	}
	return arrayVBA227ResumeNextFailureConditionInfo{
		source:  strings.ToLower(cleanIdentifier(condition)),
		negated: negated,
	}, true
}

func arrayVBA227StatementDominates(proc sourceProcedure, source, target procedureir.Statement) bool {
	if proc.Graph == nil || source.ID <= 0 || target.ID <= 0 || source.ID == target.ID {
		return false
	}
	sourceBlock, sourceOK := proc.Graph.BlockForStatement(source.ID)
	targetBlock, targetOK := proc.Graph.BlockForStatement(target.ID)
	if !sourceOK || !targetOK {
		return false
	}
	if sourceBlock.ID == targetBlock.ID {
		if source.Range.StartLine != target.Range.StartLine {
			return source.Range.StartLine < target.Range.StartLine
		}
		return source.Range.StartByte < target.Range.StartByte
	}
	return proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).Dominates(sourceBlock.ID, targetBlock.ID)
}

func arrayVBA227ResumeFactsFor(proc sourceProcedure) *arrayVBA227ResumeFacts {
	if proc.arrayVBA227ResumeFacts != nil {
		return proc.arrayVBA227ResumeFacts
	}
	return buildArrayVBA227ResumeFacts(proc)
}

func (a Analyzer) arrayVBA227Transfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, resumeNextBefore []bool, vba227Graph *vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges) (arrayFlowState, []Finding) {
	state = arrayVBA227ClearLoopBodyBounds(state, line)
	state = arrayVBA227ClearConditionalAllocationGuards(state, proc, text, line, variables)
	state = arrayVBA227InvalidateNotNotMutationState(file, proc, line, state, variables, ctx)
	state = applyArrayModuleStorageSourceGuardState(state, file, proc, line, ctx, ctx.arrayModuleStorageGuards[file.Path])
	if strings.HasPrefix(strings.ToLower(strings.TrimSpace(text)), "case ") || arrayBoundCallRe.MatchString(text) {
		state = arrayVBA227RepeatedSelectCaseBoundsState(file, proc, line, state, variables)
	}
	transfer := func(input arrayFlowState, source string) (arrayFlowState, []Finding) {
		resumeNext := arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line)
		failureInput := input
		if resumeNext {
			// arrayTransfer mutates its state map in place. Keep the pre-assignment
			// value so a failed RHS under Resume Next can retain the existing LHS
			// value instead of observing the normal assignment result as input.
			failureInput = cloneArrayState(input)
		}
		output, findings := a.arrayTransfer(file, proc, ctx, variables, input, source, line, constants, capacityGuards)
		output = arrayVBA227AttachConditionalReDimState(output, proc, source, line, variables)
		output = arrayVBA227AttachReturnProvenance(output, source, proc, ctx, variables, constants)
		output = arrayVBA227AttachAllocationFlagState(file, proc, source, line, input, output, variables)
		if resumeNext {
			output = arrayVBA227PreserveResumeNextArrayFailure(failureInput, output, source, line, proc, ctx, variables)
		}
		if resumeNextBefore == nil || resumeNext {
			return output, findings
		}
		return output, arrayVBA227FilterNestedBoundIndexFindings(findings, source, variables)
	}
	if line >= 1 && line <= len(file.Lines) && vbaLineContinues(file.Lines[line-1]) && arrayVBA227HasArrayFactoryAssignment(text) {
		text = arrayLogicalCodeLine(file.Lines, line)
	}
	inlineText := text
	if line >= 1 && line <= len(file.Lines) {
		inlineText = normalizedCodeLine(file.Lines[line-1])
	}
	if redim, ok := inlineArrayRedimText(inlineText); ok {
		text = redim
	} else if assignment, ok := inlineArrayFactoryAssignmentText(inlineText); ok {
		text = assignment
	} else if assignment, ok := inlineArrayStrConvAssignmentText(inlineText); ok {
		text = assignment
	} else if assignment, ok := inlineArraySafeBoundAssignmentText(inlineText, ctx.arraySafeBoundGuards); ok {
		text = assignment
	} else if assignment, ok := inlineArrayDictionaryAssignmentText(inlineText); ok {
		text = assignment
	} else if assignment, ok := inlineArrayReturnAssignmentText(inlineText, ctx.arrayReturns); ok {
		text = assignment
	} else if assignment, ok := inlineArrayQualifiedReturnAssignmentText(file, proc, line, inlineText, ctx.arrayReturnsQualified); ok {
		text = assignment
	} else if assignment, ok := inlineArrayAssignmentText(inlineText); ok {
		text = assignment
	}
	if condition, body, ok := arrayIfThenParts(text); ok {
		if strings.TrimSpace(body) != "" {
			state = applyArrayModuleStorageCondition(state, condition, vbacfg.EdgeBranchTrue, line, file, proc, ctx, ctx.arrayModuleStorageGuards[file.Path])
		}
		if body != "" && !arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) && !arrayProcedureHasErrorHandling(proc) && arrayVBA227StatementAlwaysRaises(body) {
			if guardedState, safe := arrayNonEmptyGuardState(state, condition, variables); safe {
				_, findings := transfer(state, condition)
				return guardedState, findings
			}
		}
		if body != "" && !arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
			if guardedState, safe := arraySafeBoundBranchState(state, condition, vbacfg.EdgeBranchTrue, ctx.arraySafeBoundGuards, variables); safe {
				conditionState, findings := transfer(state, condition)
				thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
				thenState := guardedState
				if thenBody != "" {
					var thenFindings []Finding
					thenState, thenFindings = transfer(thenState, thenBody)
					findings = append(findings, thenFindings...)
				}
				elseState := conditionState
				if hasElse && elseBody != "" {
					var elseFindings []Finding
					elseState, elseFindings = transfer(elseState, elseBody)
					findings = append(findings, elseFindings...)
				}
				if hasElse {
					state = meetArrayState(thenState, elseState)
				} else {
					state = meetArrayState(thenState, conditionState)
				}
				return state, findings
			}
		}
		if argument, _, guard := arrayIsArrayGuardCondition(condition); guard && arrayElementBaseName(argument) != "" {
			state, findings := transfer(state, condition)
			guardedState := arrayElementGuardState(state, argument, variables)
			thenState := cloneArrayState(guardedState)
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			if thenBody != "" {
				var bodyFindings []Finding
				thenState, bodyFindings = transfer(thenState, thenBody)
				findings = append(findings, bodyFindings...)
			}
			if hasElse {
				elseState := cloneArrayState(guardedState)
				var elseFindings []Finding
				if elseBody != "" {
					elseState, elseFindings = transfer(elseState, elseBody)
				}
				findings = append(findings, elseFindings...)
				state = meetArrayState(thenState, elseState)
			} else {
				state = thenState
			}
			return state, findings
		}
		// A normal multi-line If evaluates its condition before either branch
		// can run. If a typed array's bounds query returns normally, the array
		// is allocated on both the true and false paths. Keep this refinement
		// narrow: ElseIf merging and inline bodies retain their existing CFG
		// handling, and Resume Next may continue after a failed query.
		if body == "" && strings.HasPrefix(strings.ToLower(strings.TrimSpace(condition)), "if ") && arrayVBA227HasBoundsCondition(condition) && !arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
			beforeBounds := cloneArrayState(state)
			state, findings := transfer(state, condition)
			state = arraySuccessfulBoundsState(state, condition, variables, arrayVBA227LoopBodyEndLine(proc, line))
			state = applyArrayModuleCoupledBoundsState(state, file, variables, file.moduleDecls(), condition)
			return arrayVBA227RetainBoundsFailureOnResume(state, beforeBounds, condition, variables, proc), findings
		}
	}
	state, findings := transfer(state, text)
	findings = a.arrayVBA227AddResumeBoundIndexFindings(findings, file, proc, line, text, state, variables, vba227Graph, resumeNextEdges)
	beforeBounds := cloneArrayState(state)
	findings = arrayVBA227FilterSuccessfulBoundsGuardBodyIndexFindings(findings, file, proc, line, variables, resumeNextBefore)
	findings = arrayVBA227FilterSuccessfulIndexedConditionBodyFindings(findings, file, proc, line, variables, resumeNextBefore, vba227Graph, resumeNextEdges)
	findings = arrayVBA227FilterConditionalBodyIndexFindings(findings, file, proc, line, state, variables, ctx, resumeNextBefore)
	findings = arrayVBA227FilterForBodyIndexFindings(findings, file, proc, line, state, variables, ctx, resumeNextBefore, vba227Graph, resumeNextEdges)
	if (arrayVBA227HasSuccessfulBoundsExpression(text) || arrayVBA227HasDictionaryBoundsExpression(text, state)) &&
		!arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) &&
		!strings.Contains(strings.ToLower(text), "on error resume next") {
		state = arraySuccessfulBoundsState(state, text, variables, arrayVBA227LoopBodyEndLine(proc, line))
		state = applyArrayModuleCoupledBoundsState(state, file, variables, file.moduleDecls(), text)
		state = arrayVBA227RetainBoundsFailureOnResume(state, beforeBounds, text, variables, proc)
	}
	// Source-line CFG blocks can contain an If condition and its body. Apply
	// the normal-path fact after the condition while the block is still being
	// processed so a nested element access in the body does not repeat the
	// condition's possible outer-array failure.
	if argument, _, ok := arrayIsArrayGuardCondition(text); ok {
		state = arrayElementGuardState(state, argument, variables)
	}
	if name, ok := arraySafeArrayPointerGuardTarget(file, proc, line, text, variables, ctx.arrayVBA227ExternalAPIs); ok {
		if value, known := state[name]; known {
			value.kind = arrayAllocated
			value.knownArray = true
			// A nonzero SAFEARRAY descriptor proves that bounds can be
			// queried, but it does not prove that the descriptor contains an
			// element. Retain a possible-empty state so indexed access remains
			// checked even when the incoming ByRef state had no shape facts.
			value.mayBeEmpty = true
			state[name] = value
		}
	}
	return state, findings
}

func arrayVBA227InvalidateNotNotMutationState(file parsedFile, proc sourceProcedure, line int, state arrayFlowState, variables map[string]arrayVariable, ctx analysisContext) arrayFlowState {
	access := procedureStatementAtLine(proc, line)
	if access.ID == 0 {
		return state
	}
	name, guardLine, ok := arrayVBA227EnclosingNotNotGuard(proc, access, variables)
	if !ok || arrayVBA227NoArrayMutationAfterGuard(file, proc, name, guardLine, line, ctx) {
		return state
	}
	value, known := state[name]
	if !known {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayUnknown
	value.knownArray = true
	value.mayBeEmpty = true
	value.mayBeUnallocated = true
	value.dimensions = nil
	value.preserveShape = nil
	updated[name] = value
	return updated
}

// arrayVBA227PreserveResumeNextArrayFailure keeps a possible failed call
// visible after `On Error Resume Next`. A documented or inferred array return
// normally establishes an allocated value, but a failed assignment leaves a
// Variant/array target uninitialized and VBA continues to the next statement.
// Only downgrade a target whose RHS was already proven to be an array; an
// arbitrary Variant call remains fail-open to avoid turning unknown values into
// diagnostics.
func arrayVBA227PreserveResumeNextArrayFailure(input, output arrayFlowState, text string, line int, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable) arrayFlowState {
	lhs, rhs, indexed, assigned := arrayAssignment(text)
	if !assigned || indexed {
		return output
	}
	name := strings.ToLower(cleanIdentifier(lhs))
	variable, known := variables[name]
	if !known || !variable.isArray && !variable.isVariant {
		return output
	}
	if !arrayVBA227MayFailArrayExpression(rhs) {
		return output
	}
	value, provenArray := arrayExpressionStateForProcedure(rhs, input, ctx, proc)
	if !provenArray || !value.knownArray {
		value, provenArray = arrayQualifiedReturnExpressionState(proc, line, rhs, variables, ctx)
	}
	if !provenArray || !value.knownArray {
		value, provenArray = arrayVBA227QualifiedArrayReturnValue(rhs, ctx, variables)
	}
	if !provenArray || !value.knownArray {
		return output
	}
	if _, assignedOutput := output[name]; !assignedOutput {
		return output
	}
	inputValue, assignedInput := input[name]
	if !assignedInput {
		inputValue = arrayValue{
			kind:       arrayUnknown,
			knownArray: variable.isArray,
			origin:     arrayOriginUnknown,
		}
	}
	if inputValue.kind != arrayAllocated || !inputValue.knownArray || inputValue.mayBeUnallocated {
		inputValue.mayBeUnallocated = true
	}
	updated := cloneArrayState(output)
	updatedValue := meetArrayValue(output[name], inputValue)
	if value.origin == arrayOriginRangeValue {
		if guard, ok := arrayVBA227ResumeNextFailureGuardForAssignment(proc, line, text, name); ok {
			updatedValue.resumeNextFailureFlagSource = guard.source
			updatedValue.resumeNextFailureFlagSuccessOnTrue = guard.successOnTrue
		} else {
			updatedValue.resumeNextFailureFlagSource = ""
			updatedValue.resumeNextFailureFlagSuccessOnTrue = false
		}
	}
	updated[name] = updatedValue
	return updated
}

func arrayVBA227ResumeNextFailureGuardForAssignment(proc sourceProcedure, line int, text, target string) (arrayVBA227ResumeNextFailureGuard, bool) {
	facts := arrayVBA227ResumeFactsFor(proc)
	statementIDs := facts.resumeNextFailureGuardStatementsByLine[line]
	if len(statementIDs) == 0 {
		return arrayVBA227ResumeNextFailureGuard{}, false
	}
	normalizedText := strings.TrimSpace(text)
	var match arrayVBA227ResumeNextFailureGuard
	matchedCount := 0
	var exactMatch arrayVBA227ResumeNextFailureGuard
	exactCount := 0
	for _, statementID := range statementIDs {
		guard, ok := facts.resumeNextFailureGuardsByStatement[statementID]
		if !ok || !strings.EqualFold(guard.target, target) {
			continue
		}
		matchedCount++
		if strings.EqualFold(guard.assignmentText, normalizedText) {
			exactMatch = guard
			exactCount++
		}
		match = guard
	}
	if matchedCount == 0 || matchedCount > 1 && exactCount != 1 {
		return arrayVBA227ResumeNextFailureGuard{}, false
	}
	if exactCount == 1 {
		return exactMatch, true
	}
	return match, true
}

func arrayVBA227MayFailArrayExpression(rhs string) bool {
	switch arrayCallName(rhs) {
	case "array", "filter", "split":
		// These VBA array factories establish an array result directly. The
		// existing transfer treats them as deterministic allocation facts,
		// including when the procedure has Resume Next enabled.
		return false
	default:
		return true
	}
}

func arrayVBA227QualifiedArrayReturnValue(rhs string, ctx analysisContext, variables map[string]arrayVariable) (arrayValue, bool) {
	receiver, member, ok := arrayMemberCallParts(rhs)
	if !ok {
		return arrayValue{}, false
	}
	variable, known := variables[strings.ToLower(cleanIdentifier(receiver))]
	if !known || variable.typ == "" {
		return arrayValue{}, false
	}
	typeName := strings.TrimSpace(variable.typ)
	if colon := strings.IndexByte(typeName, ':'); colon >= 0 {
		typeName = strings.TrimSpace(typeName[:colon])
	}
	return arrayQualifiedReturnValueForType(typeName, member, ctx)
}

// arraySafeArrayPointerGuardTarget recognizes the narrow low-level VBA idiom
// used to inspect a dynamic Byte-array descriptor without calling LBound or
// UBound first:
//
//	ptr = VarPtrArray(values)
//	If ptr = 0 Then Exit Function
//	CopyMemoryFromPtr pSA, ptr, LenB(pSA)
//	If pSA = 0 Then Exit Function
//
// The final guard's normal path has a nonzero SAFEARRAY descriptor, which is
// enough to make later bounds queries valid. Keep the contract contiguous and
// structural; a pointer-slot check alone, a different memory-copy shape, or a
// missing descriptor check must remain conservative.
func arraySafeArrayPointerGuardTarget(file parsedFile, proc sourceProcedure, line int, text string, variables map[string]arrayVariable, externalAPIs arrayVBA227ExternalAPISet) (string, bool) {
	if line <= proc.StartLine || line > proc.EndLine || line > len(file.Lines) {
		return "", false
	}
	guard := strings.TrimSpace(normalizedCodeLine(text))
	match := arraySafeArrayZeroExitGuardRe.FindStringSubmatch(guard)
	if len(match) != 2 {
		return "", false
	}
	descriptorName := strings.ToLower(cleanIdentifier(match[1]))
	previous := make([]string, 0, 3)
	for index := line - 2; index >= max(proc.StartLine-1, 0) && len(previous) < 3; index-- {
		candidate := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		if candidate == "" || strings.HasPrefix(candidate, "'") || strings.HasPrefix(candidate, "#") {
			continue
		}
		previous = append(previous, candidate)
	}
	if len(previous) != 3 {
		return "", false
	}
	// The scan above is backwards from the descriptor guard.
	copyText := previous[0]
	ptrGuard := previous[1]
	pointerAssignment := previous[2]
	copyMatch := arraySafeArrayPointerCopyRe.FindStringSubmatch(copyText)
	if len(copyMatch) != 4 || !externalAPIs.hasExternalDeclare(file, "CopyMemoryFromPtr") || !strings.EqualFold(copyMatch[1], descriptorName) || !strings.EqualFold(copyMatch[1], copyMatch[3]) {
		return "", false
	}
	ptrName := strings.ToLower(cleanIdentifier(copyMatch[2]))
	ptrGuardMatch := arraySafeArrayZeroExitGuardRe.FindStringSubmatch(ptrGuard)
	if len(ptrGuardMatch) != 2 || !strings.EqualFold(ptrGuardMatch[1], ptrName) {
		return "", false
	}
	lhs, rhs, indexed, assigned := arrayAssignment(pointerAssignment)
	if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), ptrName) || !strings.EqualFold(arrayCallName(rhs), "varptrarray") || !externalAPIs.hasExternalDeclare(file, arrayCallName(rhs)) {
		return "", false
	}
	open := firstParenOutsideString(rhs)
	if open < 0 {
		return "", false
	}
	close := matchingParen(rhs, open)
	if close < 0 || strings.TrimSpace(rhs[close+1:]) != "" {
		return "", false
	}
	arguments := splitArgs(rhs[open+1 : close])
	if len(arguments) != 1 {
		return "", false
	}
	arrayName := directArrayArgumentName(arguments[0])
	variable, known := variables[arrayName]
	if !known || !variable.isArray || !isByteArrayVariable(variable) {
		return "", false
	}
	return arrayName, true
}

func buildArrayVBA227ExternalAPISet(files []parsedFile) arrayVBA227ExternalAPISet {
	apis := arrayVBA227ExternalAPISet{
		declared:        map[string]bool{},
		privateByModule: map[string]map[string]bool{},
		procedures:      map[string]bool{},
	}
	for _, file := range files {
		for _, declaration := range file.IR.Declarations {
			kind := strings.ToLower(strings.TrimSpace(declaration.Kind))
			if strings.HasPrefix(kind, "declare") {
				name := strings.ToLower(cleanIdentifier(declaration.Name))
				if name == "" {
					continue
				}
				if strings.EqualFold(strings.TrimSpace(declaration.Visibility), "Private") {
					module := arrayVBA227ModuleKey(file)
					if apis.privateByModule[module] == nil {
						apis.privateByModule[module] = map[string]bool{}
					}
					apis.privateByModule[module][name] = true
				} else {
					apis.declared[name] = true
				}
			}
		}
		procedures := file.procedureView()
		for index := 0; index < procedures.Len(); index++ {
			name := strings.ToLower(cleanIdentifier(procedures.valueAt(index).Name))
			if name != "" {
				apis.procedures[name] = true
			}
		}
	}
	return apis
}

func arrayVBA227ModuleKey(file parsedFile) string {
	module := strings.TrimSpace(file.IR.ModuleName)
	if module == "" {
		module = strings.TrimSpace(file.Module)
	}
	if module == "" {
		module = strings.TrimSpace(file.Path)
	}
	return strings.ToLower(module)
}

func (apis arrayVBA227ExternalAPISet) hasExternalDeclare(file parsedFile, name string) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || apis.procedures[name] {
		return false
	}
	if apis.declared[name] {
		return true
	}
	return apis.privateByModule[arrayVBA227ModuleKey(file)][name]
}

// arrayVBA227DerivedZeroBasedLoopArray recognizes the narrow StrConv-to-Byte
// array protocol used by VBA networking code. UBound(array) + 1 can still
// raise, so the bound finding remains; once that assignment succeeds, the
// zero-based loop body cannot index an empty or unallocated result.
func arrayVBA227DerivedZeroBasedLoopArray(file parsedFile, proc sourceProcedure, loop procedureir.Statement, variables map[string]arrayVariable, ctx analysisContext) (string, bool) {
	header := strings.TrimSpace(loop.Text)
	if newline := strings.IndexAny(header, "\r\n"); newline >= 0 {
		header = strings.TrimSpace(header[:newline])
	}
	header = strings.TrimSpace(normalizedCodeLine(header))
	match := arrayForZeroBasedLengthRe.FindStringSubmatch(header)
	if len(match) != 2 {
		return "", false
	}
	lengthName := strings.ToLower(cleanIdentifier(match[1]))
	lengthLine := 0
	arrayName := ""
	lengthSource := ""
	lengthSourceLine := 0
	for line := loop.Range.StartLine - 1; line >= proc.StartLine && line <= len(file.Lines); line-- {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				continue
			}
			var ok bool
			arrayName, ok = arrayVBA227UBoundLengthSource(rhs)
			if ok {
				lengthSource = "ubound"
			} else {
				arrayName, ok = arrayVBA227SafeArrayLengthSource(rhs, ctx)
				if !ok {
					return "", false
				}
				lengthSource = "safe-array-length"
			}
			lengthLine = line
			lengthSourceLine = line
			break
		}
		if lengthLine != 0 {
			break
		}
	}
	if lengthLine == 0 || arrayName == "" {
		return "", false
	}
	variable, knownVariable := variables[arrayName]
	if !knownVariable || !variable.isArray || !isByteArrayVariable(variable) {
		return "", false
	}
	switch lengthSource {
	case "ubound":
		if !arrayVBA227StatementLineDominates(proc, lengthLine, loop) || !arrayVBA227HasZeroBasedStrConvAssignment(file, proc, arrayName, lengthLine) {
			return "", false
		}
	case "safe-array-length":
		sourceLine, ok := arrayVBA227ZeroBasedArraySourceLine(file, proc, arrayName, lengthLine)
		if !ok || sourceLine == 0 || lengthSourceLine <= sourceLine || !arrayVBA227NoArrayMutationBetween(file, proc, arrayName, sourceLine+1, lengthSourceLine) {
			return "", false
		}
	}
	for line := lengthLine + 1; line < loop.Range.StartLine && line <= len(file.Lines); line++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			if arrayVBA227MutatesArray(source, arrayName) {
				return "", false
			}
		}
	}
	return arrayName, true
}

func arrayVBA227NoArrayMutationAfterSafeArrayLength(file parsedFile, proc sourceProcedure, arrayName, indexName, lengthName string, line int, ctx analysisContext) bool {
	if line <= 0 || line > len(file.Lines) {
		return false
	}
	arrayName = strings.ToLower(cleanIdentifier(arrayName))
	lengthName = strings.ToLower(cleanIdentifier(lengthName))
	spans := splitRangeValueSourceStatementsWithOffsets(normalizedCodeLine(file.Lines[line-1]))
	found := false
	for index, span := range spans {
		lhs, rhs, indexed, assigned := arrayAssignment(span.text)
		if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
			continue
		}
		if _, ok := arrayVBA227SafeArrayLengthSource(rhs, ctx); !ok {
			continue
		}
		found = true
		for _, later := range spans[index+1:] {
			if arrayVBA227SourceMutatesArray(later.text, arrayName) {
				return false
			}
			lhs, _, indexed, assigned := arrayAssignment(later.text)
			if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				return false
			}
			if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), indexName) {
				_, rhs, _, _ := arrayAssignment(later.text)
				value, zero := integerLiteral(rhs)
				if !zero || value != 0 {
					return false
				}
			}
		}
		break
	}
	if !found {
		return false
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line || arrayVBA227SafeArrayLengthCall(proc, call, arrayName, ctx) {
			continue
		}
		if arrayVBA227CallMayInvalidateArray(proc, call, arrayName, ctx) || arrayVBA227CallMayMutateScalar(proc, call, arrayName, indexName, lengthName, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227NoArrayOrScalarMutationOnAccessLine(file parsedFile, proc sourceProcedure, arrayName, indexName, lengthName string, line int, ctx analysisContext) bool {
	if line <= 0 || line > len(file.Lines) {
		return false
	}
	arrayName = strings.ToLower(cleanIdentifier(arrayName))
	indexName = strings.ToLower(cleanIdentifier(indexName))
	lengthName = strings.ToLower(cleanIdentifier(lengthName))
	for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
		if arrayVBA227SourceMutatesArray(source, arrayName) {
			return false
		}
		lhs, _, indexed, assigned := arrayAssignment(source)
		if assigned && !indexed {
			lhs = strings.ToLower(cleanIdentifier(lhs))
			if lhs == indexName || lhs == lengthName {
				return false
			}
		}
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line {
			continue
		}
		arrayMutates := arrayVBA227CallMayInvalidateArray(proc, call, arrayName, ctx)
		scalarMutates := arrayVBA227CallMayMutateScalar(proc, call, arrayName, indexName, lengthName, ctx)
		if arrayMutates || scalarMutates {
			return false
		}
	}
	return true
}

func arrayVBA227NoUnsafeDoWhileBodyMutationAfterAccess(file parsedFile, proc sourceProcedure, loop procedureir.Statement, arrayName, indexName, lengthName string, accessLine, endLine int, ctx analysisContext) bool {
	if accessLine <= 0 || endLine <= accessLine {
		return false
	}
	stepDependencies := arrayVBA227LoopStepDependencies(file, proc, loop, indexName)
	if !arrayVBA227NoArrayMutationBetweenWithCalls(file, proc, arrayName, accessLine+1, endLine, ctx) {
		return false
	}
	if accessLine <= len(file.Lines) && arrayVBA227SourceAssignsStepDependency(normalizedCodeLine(file.Lines[accessLine-1]), stepDependencies) {
		return false
	}
	for line := max(accessLine+1, proc.StartLine); line < endLine && line <= len(file.Lines); line++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			if arrayVBA227UnsafeDoWhileScalarMutation(file, proc, loop, arrayName, line, source, indexName, lengthName, stepDependencies, ctx) {
				return false
			}
		}
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < accessLine || call.Range.StartLine >= endLine {
			continue
		}
		if (call.Range.StartLine > accessLine && arrayVBA227CallMayMutateScalar(proc, call, arrayName, indexName, lengthName, ctx)) || arrayVBA227CallMayMutateStepDependency(file, proc, call, stepDependencies, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227UnsafeDoWhileScalarMutation(file parsedFile, proc sourceProcedure, loop procedureir.Statement, arrayName string, line int, text, indexName, lengthName string, stepDependencies map[string]bool, ctx analysisContext) bool {
	for _, source := range splitRangeValueSourceStatements(text) {
		if arrayVBA227SourceAssignsStepDependency(source, stepDependencies) {
			return true
		}
		lhs, rhs, indexed, assigned := arrayAssignment(source)
		if assigned && !indexed {
			name := strings.ToLower(cleanIdentifier(lhs))
			if name == lengthName {
				return true
			}
			if name == indexName && !arrayVBA227IsProvenForwardLoopIncrement(file, proc, loop, arrayName, line, rhs, indexName, lengthName, ctx) {
				return true
			}
		}
		if _, body, ok := arrayIfThenParts(source); ok && strings.TrimSpace(body) != "" {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			branches := []string{thenBody}
			if hasElse {
				branches = append(branches, elseBody)
			}
			for _, branch := range branches {
				for _, nested := range splitRangeValueSourceStatements(branch) {
					lhs, rhs, indexed, assigned := arrayAssignment(nested)
					if assigned && !indexed {
						name := strings.ToLower(cleanIdentifier(lhs))
						if name == lengthName || name == indexName && !arrayVBA227IsProvenForwardLoopIncrement(file, proc, loop, arrayName, line, rhs, indexName, lengthName, ctx) {
							return true
						}
					}
				}
			}
		}
	}
	return false
}

func arrayVBA227LoopStepDependencies(file parsedFile, proc sourceProcedure, loop procedureir.Statement, indexName string) map[string]bool {
	dependencies := map[string]bool{}
	var collect func(string)
	collect = func(text string) {
		for _, source := range splitRangeValueSourceStatements(text) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), indexName) {
				arrayVBA227AddLoopStepDependency(file, rhs, indexName, dependencies)
			}
			if _, body, ok := arrayIfThenParts(source); ok && strings.TrimSpace(body) != "" {
				thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
				collect(thenBody)
				if hasElse {
					collect(elseBody)
				}
			}
		}
	}
	for line := loop.Range.StartLine + 1; line <= loop.Range.EndLine && line <= len(file.Lines); line++ {
		collect(normalizedCodeLine(file.Lines[line-1]))
	}
	queue := make([]string, 0, len(dependencies))
	for dependency := range dependencies {
		queue = append(queue, dependency)
	}
	visited := map[string]bool{}
	for len(queue) > 0 {
		name := queue[0]
		queue = queue[1:]
		if visited[name] {
			continue
		}
		visited[name] = true
		before := len(dependencies)
		for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
			for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
				lhs, rhs, indexed, assigned := arrayAssignment(source)
				if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
					arrayVBA227AddLoopStepDependency(file, rhs, name, dependencies)
				}
			}
		}
		for dependency := range dependencies {
			if !visited[dependency] && len(dependencies) > before {
				queue = append(queue, dependency)
			}
		}
	}
	return dependencies
}

func arrayVBA227AddLoopStepDependency(file parsedFile, rhs, indexName string, dependencies map[string]bool) {
	compact := strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(rhs))), "")
	prefix := strings.ToLower(cleanIdentifier(indexName)) + "+"
	if !strings.HasPrefix(compact, prefix) {
		return
	}
	step := strings.TrimSpace(strings.TrimPrefix(compact, prefix))
	if step == "" {
		return
	}
	if value, ok := integerLiteral(step); ok && value > 0 {
		return
	}
	if value, err := constantIntegerExpression(step, arrayIntegerModuleConstants(file)); err == nil && value > 0 {
		return
	}
	if name := arrayVBA227DirectScalarReference(step); name != "" && name != strings.ToLower(cleanIdentifier(indexName)) {
		dependencies[name] = true
	}
}

func arrayVBA227SourceAssignsStepDependency(text string, dependencies map[string]bool) bool {
	if len(dependencies) == 0 {
		return false
	}
	for _, source := range splitRangeValueSourceStatements(text) {
		lhs, _, indexed, assigned := arrayAssignment(source)
		if assigned && !indexed && dependencies[strings.ToLower(cleanIdentifier(lhs))] {
			return true
		}
		if _, body, ok := arrayIfThenParts(source); ok && strings.TrimSpace(body) != "" {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			if arrayVBA227SourceAssignsStepDependency(thenBody, dependencies) || hasElse && arrayVBA227SourceAssignsStepDependency(elseBody, dependencies) {
				return true
			}
		}
	}
	return false
}

func arrayVBA227CallMayMutateStepDependency(file parsedFile, proc sourceProcedure, call procedureir.CallSite, dependencies map[string]bool, ctx analysisContext) bool {
	for dependency := range dependencies {
		if !strings.Contains(dependency, ".") {
			if arrayVBA227CallMayMutateNamedScalar(file, proc, call, dependency, ctx) {
				return true
			}
			continue
		}
		receiver := dependency[:strings.IndexByte(dependency, '.')]
		for _, argument := range arrayCallArgumentTexts(proc, call) {
			if directArrayArgumentName(argument) == receiver {
				return true
			}
		}
		for _, argument := range call.Arguments.Named {
			if directArrayArgumentName(argument.ValueText) == receiver {
				return true
			}
		}
	}
	return false
}

func arrayVBA227IsForwardLoopIncrement(rhs, indexName string) bool {
	compact := strings.ReplaceAll(strings.ToLower(strings.TrimSpace(rhs)), " ", "")
	prefix := strings.ToLower(cleanIdentifier(indexName)) + "+"
	if !strings.HasPrefix(compact, prefix) {
		return false
	}
	step := strings.TrimSpace(strings.TrimPrefix(compact, prefix))
	if step == "" || strings.HasPrefix(step, "-") {
		return false
	}
	if value, ok := integerLiteral(step); ok {
		return value > 0
	}
	return false
}

func arrayVBA227IsProvenForwardLoopIncrement(file parsedFile, proc sourceProcedure, loop procedureir.Statement, arrayName string, line int, rhs, indexName, lengthName string, ctx analysisContext) bool {
	if arrayVBA227IsForwardLoopIncrement(rhs, indexName) {
		return true
	}
	compact := strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(rhs))), "")
	prefix := strings.ToLower(cleanIdentifier(indexName)) + "+"
	if !strings.HasPrefix(compact, prefix) {
		return false
	}
	stepName := arrayVBA227DirectScalarReference(strings.TrimPrefix(compact, prefix))
	if stepName == "" || stepName == strings.ToLower(cleanIdentifier(indexName)) || stepName == strings.ToLower(cleanIdentifier(lengthName)) {
		return false
	}
	return arrayVBA227PositiveLoopScalarAtLine(file, proc, loop, stepName, indexName, lengthName, line, ctx, map[string]bool{})
}

func arrayVBA227PositiveLoopScalarAtLine(file parsedFile, proc sourceProcedure, loop procedureir.Statement, name, indexName, lengthName string, line int, ctx analysisContext, visiting map[string]bool) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || visiting[name] {
		return false
	}
	if strings.Contains(name, ".") {
		return arrayVBA227PositiveMemberExpression(file, proc, name, line, ctx)
	}
	visiting[name] = true
	defer delete(visiting, name)
	target := procedureStatementAtLine(proc, line)
	if target.ID == 0 {
		return false
	}
	found := false
	dominatingAssignment := false
	for sourceLine := proc.StartLine; sourceLine <= line && sourceLine <= len(file.Lines); sourceLine++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[sourceLine-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), name) {
				continue
			}
			found = true
			positive := arrayVBA227PositiveLoopScalarExpression(file, proc, loop, name, rhs, indexName, lengthName, sourceLine, ctx, visiting)
			positiveAfterReDim := false
			if !positive {
				positiveAfterReDim = arrayVBA227PositiveScalarAfterSuccessfulReDim(file, proc, name, line, ctx)
				positive = positiveAfterReDim
			}
			if !positive {
				return false
			}
			if positiveAfterReDim || arrayVBA227StatementLineDominates(proc, sourceLine, target) {
				dominatingAssignment = true
			}
		}
	}
	if !found || !dominatingAssignment {
		return false
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < proc.StartLine || call.Range.StartLine > line {
			continue
		}
		if arrayVBA227CallMayMutateNamedScalar(file, proc, call, name, ctx) {
			return false
		}
	}
	return true
}

// A successful normal-path ReDim with an upper bound of step - 1 also proves
// that the step is positive at the following access. This is useful when the
// step is loaded from persistent member state and the access is after the
// scratch-array allocation: a zero or negative step would fail at ReDim and
// cannot reach the later array access. Keep the inference narrow and reject
// procedures with error handling, where a failed ReDim could resume elsewhere.
func arrayVBA227PositiveScalarAfterSuccessfulReDim(file parsedFile, proc sourceProcedure, name string, accessLine int, ctx analysisContext) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || accessLine <= proc.StartLine || arrayProcedureHasErrorHandling(proc) {
		return false
	}
	target := procedureStatementAtLine(proc, accessLine)
	if target.ID == 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementReDim || statement.Range.StartLine >= accessLine || !arrayVBA227StatementLineDominates(proc, statement.Range.StartLine, target) {
			continue
		}
		line := statement.Range.StartLine
		if line <= 0 || line > len(file.Lines) || len(splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1]))) != 1 {
			continue
		}
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
			continue
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct || !arrayVBA227RedimUsesPositiveScalar(redim.dimensions, name) {
				continue
			}
			if !arrayVBA227NoScalarAssignmentBeforeAccess(file, proc, name, line+1, accessLine, ctx) {
				continue
			}
			return true
		}
	}
	return false
}

func arrayVBA227RedimUsesPositiveScalar(dimensions, name string) bool {
	parts := splitArgs(dimensions)
	if len(parts) != 1 {
		return false
	}
	bound := canonicalArrayBoundExpression(parts[0])
	return bound == "0to"+name+"-1"
}

func arrayVBA227DirectScalarReference(text string) string {
	text = strings.TrimSpace(text)
	if text == "" || !isIdentifierStart(text[0]) {
		return ""
	}
	for index := 1; index < len(text); index++ {
		if text[index] == '.' {
			if index+1 >= len(text) || !isIdentifierStart(text[index+1]) {
				return ""
			}
			index++
			continue
		}
		if !isIdentifierPart(text[index]) {
			return ""
		}
	}
	return strings.ToLower(cleanIdentifier(text))
}

func arrayVBA227PositiveLoopScalarExpression(file parsedFile, proc sourceProcedure, loop procedureir.Statement, name, rhs, indexName, lengthName string, line int, ctx analysisContext, visiting map[string]bool) bool {
	expression := strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(rhs))), "")
	for len(expression) >= 2 && expression[0] == '(' && expression[len(expression)-1] == ')' {
		expression = strings.TrimSpace(expression[1 : len(expression)-1])
	}
	if value, ok := integerLiteral(expression); ok {
		return value > 0
	}
	if value, err := constantIntegerExpression(expression, arrayIntegerModuleConstants(file)); err == nil {
		return value > 0
	}
	if arrayVBA227PositiveMemberExpression(file, proc, expression, line, ctx) {
		return true
	}
	if expression == strings.ToLower(cleanIdentifier(lengthName))+"-"+strings.ToLower(cleanIdentifier(indexName)) {
		return true
	}
	if arrayVBA227PositiveScalarFallback(file, proc, loop, name, line) {
		return true
	}
	referenceName := directArrayArgumentName(expression)
	if referenceName == "" {
		return false
	}
	return arrayVBA227PositiveLoopScalarAtLine(file, proc, loop, referenceName, indexName, lengthName, line, ctx, visiting)
}

func arrayVBA227PositiveScalarFallback(file parsedFile, proc sourceProcedure, loop procedureir.Statement, name string, sourceLine int) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || sourceLine <= 0 || loop.Range.StartLine <= sourceLine {
		return false
	}
	target := procedureStatementAtLine(proc, loop.Range.StartLine)
	if target.ID == 0 {
		return false
	}
	constants := arrayIntegerModuleConstants(file)
	for guardLine := sourceLine + 1; guardLine < loop.Range.StartLine && guardLine <= len(file.Lines); guardLine++ {
		if !arrayVBA227NonPositiveScalarGuard(normalizedCodeLine(file.Lines[guardLine-1]), name) {
			continue
		}
		if !arrayVBA227StatementLineDominates(proc, guardLine, target) {
			continue
		}
		for candidateLine := guardLine + 1; candidateLine < loop.Range.StartLine && candidateLine <= len(file.Lines); candidateLine++ {
			candidate := strings.TrimSpace(normalizedCodeLine(file.Lines[candidateLine-1]))
			lower := strings.ToLower(candidate)
			if strings.HasPrefix(lower, "end if") {
				break
			}
			if strings.HasPrefix(lower, "else") {
				break
			}
			for _, source := range splitRangeValueSourceStatements(candidate) {
				lhs, rhs, indexed, assigned := arrayAssignment(source)
				if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), name) {
					continue
				}
				value, err := constantIntegerExpression(strings.TrimSpace(rhs), constants)
				if err == nil && value > 0 {
					return true
				}
				return false
			}
		}
	}
	return false
}

func arrayVBA227NonPositiveScalarGuard(text, name string) bool {
	trimmed := strings.TrimSpace(strings.ToLower(text))
	if !strings.HasPrefix(trimmed, "if ") || !strings.HasSuffix(trimmed, " then") {
		return false
	}
	condition := strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(trimmed, "if "), " then"))
	return condition == strings.ToLower(cleanIdentifier(name))+" <= 0"
}

func arrayVBA227PositiveMemberExpression(file parsedFile, proc sourceProcedure, expression string, line int, ctx analysisContext) bool {
	expression = strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(expression))), "")
	if expression == "" || !strings.Contains(expression, ".") {
		return false
	}
	for _, char := range expression {
		if char != '.' && !isVBAIdentifierRune(char) {
			return false
		}
	}
	found := false
	constants := arrayIntegerModuleConstants(file)
	target := procedureStatementAtLine(proc, line)
	for sourceLine := 1; sourceLine < line && sourceLine <= len(file.Lines); sourceLine++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[sourceLine-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !arrayVBA227MemberAssignmentMatches(lhs, expression) {
				continue
			}
			found = true
			owner, ok := arrayModuleProcedureAtLine(file, sourceLine)
			if !ok || !arrayVBA227MemberAssignmentIsUnconditional(owner, sourceLine) {
				return false
			}
			if arrayProcedureKey(owner) != arrayProcedureKey(proc) {
				participant := arrayProcedureIsParticipant(ctx, owner)
				available := arrayVBA227MemberAssignmentAvailableBefore(proc, owner, target, ctx)
				sameModule := strings.TrimSpace(owner.Module) != "" && strings.TrimSpace(proc.Module) != "" && strings.EqualFold(strings.TrimSpace(owner.Module), strings.TrimSpace(proc.Module))
				// A same-module Private Type member is source-owned by the caller's
				// module. A direct dominating call is enough to prove its value even
				// when the participant planner intentionally excludes the scalar-only
				// initializer from the array fixed-point cluster.
				if (!participant && !sameModule) || !available {
					return false
				}
			} else if target.ID == 0 || !arrayVBA227StatementLineDominates(owner, sourceLine, target) {
				return false
			}
			positive := arrayVBA227PositiveProcedureScalarExpression(file, owner, rhs, sourceLine, constants, ctx, map[string]bool{})
			if !positive {
				return false
			}
		}
		if arrayVBA227SourceAssignsMemberConditionally(normalizedCodeLine(file.Lines[sourceLine-1]), expression) {
			return false
		}
	}
	if line >= 1 && line <= len(file.Lines) && arrayVBA227SourceAssignsMember(normalizedCodeLine(file.Lines[line-1]), expression) {
		// A source line may contain both the loop access and a later colon-
		// separated member write. The IR range is line-based here, so reject
		// the whole line rather than carrying the earlier positive value past
		// an update whose exact segment cannot be proven.
		return false
	}
	return found
}

// arrayVBA227MemberAssignmentAvailableBefore proves the interprocedural part
// of a member-step contract. Participant membership only says that two
// procedures may exchange array facts; it does not say that a setter ran on
// the path reaching this access. Require a reachable call that dominates the
// access, and recursively require the same property for private helper calls.
func arrayVBA227MemberAssignmentAvailableBefore(caller, owner sourceProcedure, access procedureir.Statement, ctx analysisContext) bool {
	if access.ID == 0 || caller.Graph == nil {
		return false
	}
	return arrayVBA227MemberAssignmentCallBefore(caller, owner, access, ctx, map[string]bool{})
}

func arrayVBA227MemberAssignmentCallBefore(caller, owner sourceProcedure, access procedureir.Statement, ctx analysisContext, visiting map[string]bool) bool {
	for call := range caller.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine >= access.Range.StartLine || !applicationStateCallReachable(caller, call) {
			continue
		}
		if !arrayVBA227StatementLineDominates(caller, call.Range.StartLine, access) {
			continue
		}
		if arrayVBA227CallTargetsProcedure(ctx, call, owner) {
			return true
		}
		_, callee, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !ok || !arrayProcedureIsParticipant(ctx, callee) || arrayProcedureKey(callee) == arrayProcedureKey(caller) {
			continue
		}
		if arrayVBA227ProcedureCallsMemberBeforeNormalExit(callee, owner, ctx, visiting) {
			return true
		}
	}
	return false
}

func arrayVBA227ProcedureCallsMemberBeforeNormalExit(proc, owner sourceProcedure, ctx analysisContext, visiting map[string]bool) bool {
	key := arrayProcedureKey(proc)
	if key == "" || visiting[key] || proc.Graph == nil {
		return false
	}
	visiting[key] = true
	defer delete(visiting, key)
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || !applicationStateCallReachable(proc, call) || !arrayVBA227StatementLineDominatesNormalExit(proc, call.Range.StartLine) {
			continue
		}
		if arrayVBA227CallTargetsProcedure(ctx, call, owner) {
			return true
		}
		_, callee, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !ok || !arrayProcedureIsParticipant(ctx, callee) {
			continue
		}
		if arrayVBA227ProcedureCallsMemberBeforeNormalExit(callee, owner, ctx, visiting) {
			return true
		}
	}
	return false
}

func arrayVBA227CallTargetsProcedure(ctx analysisContext, call procedureir.CallSite, target sourceProcedure) bool {
	resolution := arrayCallResolution(ctx, call)
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return false
	}
	qualified := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
	if qualified == arrayProcedureKey(target) || qualified == strings.ToLower(strings.TrimSpace(target.Name)) {
		return true
	}
	_, resolved, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	return ok && arrayProcedureKey(resolved) == arrayProcedureKey(target)
}

func arrayVBA227MemberAssignmentIsUnconditional(proc sourceProcedure, line int) bool {
	statement := procedureStatementAtLine(proc, line)
	return statement.ID != 0 && arrayVBA227BranchOwner(proc, statement) == 0 && arrayVBA227StatementLineDominatesNormalExit(proc, line)
}

func arrayVBA227SourceAssignsMemberConditionally(text, member string) bool {
	for _, source := range splitRangeValueSourceStatements(text) {
		if _, _, indexed, assigned := arrayAssignment(source); assigned && !indexed {
			lhs, _, _, _ := arrayAssignment(source)
			if arrayVBA227MemberAssignmentMatches(lhs, member) {
				continue
			}
		}
		_, body, ok := arrayIfThenParts(source)
		if !ok || strings.TrimSpace(body) == "" {
			continue
		}
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		if arrayVBA227SourceAssignsMember(thenBody, member) || hasElse && arrayVBA227SourceAssignsMember(elseBody, member) {
			return true
		}
	}
	return false
}

func arrayVBA227SourceAssignsMember(text, member string) bool {
	for _, source := range splitRangeValueSourceStatements(text) {
		lhs, _, indexed, assigned := arrayAssignment(source)
		if assigned && !indexed && arrayVBA227MemberAssignmentMatches(lhs, member) {
			return true
		}
		_, body, ok := arrayIfThenParts(source)
		if !ok || strings.TrimSpace(body) == "" {
			continue
		}
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		if arrayVBA227SourceAssignsMember(thenBody, member) || hasElse && arrayVBA227SourceAssignsMember(elseBody, member) {
			return true
		}
	}
	return false
}

func arrayVBA227MemberAssignmentMatches(lhs, member string) bool {
	lhs = strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(lhs))), "")
	member = strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(member))), "")
	return lhs == member || strings.HasPrefix(member, ".") && strings.HasSuffix(lhs, member)
}

func arrayVBA227PositiveProcedureScalarExpression(file parsedFile, proc sourceProcedure, rhs string, line int, constants map[string]int, ctx analysisContext, visiting map[string]bool) bool {
	expression := strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(rhs))), "")
	for len(expression) >= 2 && expression[0] == '(' && expression[len(expression)-1] == ')' {
		expression = strings.TrimSpace(expression[1 : len(expression)-1])
	}
	if value, err := constantIntegerExpression(expression, constants); err == nil {
		return value > 0
	}
	name := directArrayArgumentName(expression)
	if name == "" {
		return false
	}
	return arrayVBA227PositiveProcedureScalarAtLine(file, proc, name, line, constants, ctx, visiting)
}

func arrayVBA227PositiveProcedureScalarAtLine(file parsedFile, proc sourceProcedure, name string, line int, constants map[string]int, ctx analysisContext, visiting map[string]bool) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || visiting[name] {
		return false
	}
	visiting[name] = true
	defer delete(visiting, name)
	found := false
	dominatingPositive := false
	target := procedureStatementAtLine(proc, line)
	if target.ID == 0 {
		return false
	}
	for sourceLine := proc.StartLine; sourceLine <= line && sourceLine <= len(file.Lines); sourceLine++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[sourceLine-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), name) {
				continue
			}
			found = true
			expressionPositive := arrayVBA227PositiveProcedureScalarExpression(file, proc, rhs, sourceLine, constants, ctx, visiting)
			floorPositive := arrayVBA227PositiveScalarFloor(file, proc, name, sourceLine, line, constants)
			if !expressionPositive && !floorPositive {
				return false
			}
			if arrayVBA227StatementLineDominates(proc, sourceLine, target) {
				dominatingPositive = true
			}
		}
	}
	if !found || !dominatingPositive {
		return false
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < proc.StartLine || call.Range.StartLine > line {
			continue
		}
		if arrayVBA227CallMayMutateNamedScalar(file, proc, call, name, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227PositiveScalarFloor(file parsedFile, proc sourceProcedure, name string, sourceLine, targetLine int, constants map[string]int) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || sourceLine <= 0 || targetLine <= sourceLine {
		return false
	}
	for guardLine := sourceLine + 1; guardLine < targetLine && guardLine <= len(file.Lines); guardLine++ {
		floor, ok := arrayVBA227PositiveScalarFloorGuard(normalizedCodeLine(file.Lines[guardLine-1]), name, constants)
		if !ok {
			continue
		}
		for candidateLine := guardLine + 1; candidateLine < targetLine && candidateLine <= len(file.Lines); candidateLine++ {
			candidate := strings.TrimSpace(normalizedCodeLine(file.Lines[candidateLine-1]))
			lower := strings.ToLower(candidate)
			if strings.HasPrefix(lower, "end if") || strings.HasPrefix(lower, "else") {
				break
			}
			for _, source := range splitRangeValueSourceStatements(candidate) {
				lhs, rhs, indexed, assigned := arrayAssignment(source)
				if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), name) {
					continue
				}
				value, err := constantIntegerExpression(strings.TrimSpace(rhs), constants)
				if err == nil && value >= floor {
					return true
				}
				return false
			}
		}
	}
	return false
}

func arrayVBA227PositiveScalarFloorGuard(text, name string, constants map[string]int) (int, bool) {
	trimmed := strings.Join(strings.Fields(strings.ToLower(strings.TrimSpace(text))), " ")
	if !strings.HasPrefix(trimmed, "if ") || !strings.HasSuffix(trimmed, " then") {
		return 0, false
	}
	condition := strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(trimmed, "if "), " then"))
	operator := strings.Index(condition, "<")
	if operator < 0 {
		return 0, false
	}
	lhs := directArrayArgumentName(condition[:operator])
	if lhs != strings.ToLower(cleanIdentifier(name)) {
		return 0, false
	}
	floorText := strings.TrimSpace(condition[operator+1:])
	if strings.HasPrefix(floorText, "=") {
		floorText = strings.TrimSpace(floorText[1:])
	}
	floor, err := constantIntegerExpression(floorText, constants)
	if err != nil || floor <= 0 {
		return 0, false
	}
	return floor, true
}

func arrayVBA227CallMayMutateNamedScalar(file parsedFile, proc sourceProcedure, call procedureir.CallSite, name string, ctx analysisContext) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || arrayByRefCallIsReadOnly(call) {
		return false
	}
	_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !resolved {
		if arrayVBA227ExternalCallPassesNamedScalarByVal(file, proc, call, name) {
			return false
		}
		for _, argument := range call.Arguments.Named {
			if directArrayArgumentName(argument.ValueText) == name {
				return true
			}
		}
		for _, argument := range arrayCallArgumentTexts(proc, call) {
			if directArrayArgumentName(argument) == name {
				return true
			}
		}
		return false
	}
	bindings, mapped := arrayCallArgumentBindings(proc, target, call)
	if !mapped {
		return arrayCallPassesDirectArrayArgument(proc, call, name)
	}
	for _, binding := range bindings {
		if directArrayArgumentName(binding.text) != name || binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() {
			continue
		}
		if parameterIsByRefScalar(target.Params.valueAt(binding.parameterIndex)) {
			return true
		}
	}
	return false
}

func arrayVBA227ExternalCallPassesNamedScalarByVal(file parsedFile, proc sourceProcedure, call procedureir.CallSite, name string) bool {
	callee := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
	if callee == "" {
		return false
	}
	matchedDeclaration := false
	for _, declaration := range file.IR.Declarations {
		kind := strings.ToLower(strings.TrimSpace(declaration.Kind))
		if !strings.HasPrefix(kind, "declare") || !strings.EqualFold(cleanIdentifier(declaration.Name), callee) {
			continue
		}
		matchedDeclaration = true
		target := sourceProcedure{Params: newReadOnlySpan(declaration.Parameters)}
		bindings, mapped := arrayCallArgumentBindings(proc, target, call)
		if !mapped {
			return false
		}
		argumentFound := false
		for _, binding := range bindings {
			if directArrayArgumentName(binding.text) != name {
				continue
			}
			argumentFound = true
			if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || parameterIsByRefScalar(target.Params.valueAt(binding.parameterIndex)) {
				return false
			}
		}
		if !argumentFound {
			return false
		}
	}
	return matchedDeclaration
}

func arrayVBA227CallMayMutateScalar(proc sourceProcedure, call procedureir.CallSite, arrayName, indexName, lengthName string, ctx analysisContext) bool {
	if arrayByRefCallIsReadOnly(call) {
		return false
	}
	if strings.EqualFold(cleanIdentifier(call.Callee.BaseName), cleanIdentifier(arrayName)) {
		return false
	}
	_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !resolved {
		for _, argument := range call.Arguments.Named {
			name := directArrayArgumentName(argument.ValueText)
			if name == indexName || name == lengthName {
				return true
			}
		}
		for _, argument := range arrayCallArgumentTexts(proc, call) {
			name := directArrayArgumentName(argument)
			if name == indexName || name == lengthName {
				return true
			}
		}
		return false
	}
	bindings, mapped := arrayCallArgumentBindings(proc, target, call)
	if !mapped {
		return arrayCallPassesDirectArrayArgument(proc, call, indexName) || arrayCallPassesDirectArrayArgument(proc, call, lengthName)
	}
	for _, binding := range bindings {
		name := directArrayArgumentName(binding.text)
		if name != indexName && name != lengthName || binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() {
			continue
		}
		if parameterIsByRefScalar(target.Params.valueAt(binding.parameterIndex)) {
			return true
		}
	}
	return false
}

func arrayVBA227SafeArrayLengthCall(proc sourceProcedure, call procedureir.CallSite, arrayName string, ctx analysisContext) bool {
	callee := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
	if callee == "" || !ctx.arraySafeArrayLengthGuards[callee] {
		return false
	}
	arguments := arrayCallArgumentTexts(proc, call)
	return len(arguments) == 1 && directArrayArgumentName(arguments[0]) == strings.ToLower(cleanIdentifier(arrayName))
}

// arrayVBA227DerivedDoWhileArray recognizes the lifecycle-safe part of a
// zero-based `Do While index < length` loop. A local index initialized to zero
// makes a reachable body imply a positive length; SafeArrayLen then proves
// that the source Byte array is allocated and non-empty. This deliberately
// proves only the allocation/emptiness contract. Variable-index bounds remain
// outside VBA227's range proof.
func arrayVBA227DerivedDoWhileArray(file parsedFile, proc sourceProcedure, loop procedureir.Statement, accessLine int, variables map[string]arrayVariable, ctx analysisContext) (string, bool) {
	header := strings.TrimSpace(loop.Text)
	if newline := strings.IndexAny(header, "\r\n"); newline >= 0 {
		header = strings.TrimSpace(header[:newline])
	}
	header = strings.TrimSpace(normalizedCodeLine(header))
	match := arrayDoWhileLengthRe.FindStringSubmatch(header)
	if len(match) != 3 {
		return "", false
	}
	indexName := strings.ToLower(cleanIdentifier(match[1]))
	lengthName := strings.ToLower(cleanIdentifier(match[2]))
	indexVariable, indexKnown := variables[indexName]
	lengthVariable, lengthKnown := variables[lengthName]
	if !indexKnown || indexVariable.isArray || indexVariable.isVariant || indexVariable.isObject || !lengthKnown || lengthVariable.isArray || lengthVariable.isVariant || lengthVariable.isObject {
		return "", false
	}
	indexZeroLine := 0
	for line := loop.Range.StartLine - 1; line >= proc.StartLine && line <= len(file.Lines); line-- {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), indexName) {
				continue
			}
			value, ok := integerLiteral(rhs)
			if !ok || value != 0 {
				return "", false
			}
			indexZeroLine = line
			break
		}
		if indexZeroLine != 0 {
			break
		}
	}
	if indexZeroLine == 0 {
		return "", false
	}
	if !arrayVBA227StatementLineDominatesOrUnconditional(proc, indexZeroLine, loop) {
		return "", false
	}
	if !arrayVBA227NoScalarAssignmentBeforeAccess(file, proc, indexName, indexZeroLine+1, loop.Range.StartLine, ctx) {
		return "", false
	}
	for line := proc.StartLine; line < loop.Range.StartLine && line <= len(file.Lines); line++ {
		if line == indexZeroLine {
			continue
		}
		if arrayVBA227SourceAssignsScalar(normalizedCodeLine(file.Lines[line-1]), indexName) {
			return "", false
		}
	}
	lengthLine := 0
	arrayName := ""
	for line := loop.Range.StartLine - 1; line >= proc.StartLine && line <= len(file.Lines); line-- {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				continue
			}
			var ok bool
			arrayName, ok = arrayVBA227SafeArrayLengthSource(rhs, ctx)
			if !ok {
				return "", false
			}
			lengthLine = line
			break
		}
		if lengthLine != 0 {
			break
		}
	}
	if lengthLine == 0 || arrayName == "" {
		return "", false
	}
	arrayVariable, arrayKnown := variables[arrayName]
	if !arrayKnown || !arrayVariable.isArray || !isByteArrayVariable(arrayVariable) {
		return "", false
	}
	if !arrayVBA227StatementLineDominates(proc, lengthLine, loop) {
		return "", false
	}
	for line := proc.StartLine; line < lengthLine && line <= len(file.Lines); line++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, _, indexed, assigned := arrayAssignment(source)
			if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				return "", false
			}
		}
	}
	sourceLine, sourceOK := arrayVBA227ZeroBasedArraySourceLine(file, proc, arrayName, lengthLine)
	if sourceOK {
		if sourceLine == 0 || lengthLine <= sourceLine || !arrayVBA227NoArrayMutationBetween(file, proc, arrayName, sourceLine+1, lengthLine) {
			return "", false
		}
	} else if !arrayVBA227AllIndexedUsesUseLowerBound(file, arrayName, accessLine, variables) {
		// A SafeArrayLen result may validate a ByRef array directly, without a
		// local zero-based factory assignment. That form is accepted only when
		// every indexed use on the normalized source line explicitly derives its
		// index from LBound; a plain values(offset) access still needs a proven
		// zero-based source.
		return "", false
	}
	if !arrayVBA227NoArrayMutationAfterSafeArrayLength(file, proc, arrayName, indexName, lengthName, lengthLine, ctx) {
		return "", false
	}
	noBetween := arrayVBA227NoArrayMutationBetweenWithCalls(file, proc, arrayName, lengthLine+1, accessLine, ctx)
	noAccess := arrayVBA227NoArrayOrScalarMutationOnAccessLine(file, proc, arrayName, indexName, lengthName, accessLine, ctx)
	if accessLine <= lengthLine || !noBetween || !noAccess {
		return "", false
	}
	noUnsafe := arrayVBA227NoUnsafeDoWhileBodyMutationAfterAccess(file, proc, loop, arrayName, indexName, lengthName, accessLine, loop.Range.EndLine, ctx)
	noIndex := arrayVBA227NoScalarAssignmentBetween(file, proc, indexName, loop.Range.StartLine+1, accessLine, ctx)
	noLength := arrayVBA227NoScalarAssignmentBetween(file, proc, lengthName, lengthLine+1, accessLine, ctx)
	if !noUnsafe {
		return "", false
	}
	if !noIndex {
		return "", false
	}
	if !noLength {
		return "", false
	}
	return arrayName, true
}

func arrayVBA227SafeArrayLengthSource(rhs string, ctx analysisContext) (string, bool) {
	name := arrayCallName(rhs)
	if name == "" || !ctx.arraySafeArrayLengthGuards[name] {
		return "", false
	}
	arguments, ok := arraySimpleCallArguments(rhs)
	if !ok || len(arguments) != 1 {
		return "", false
	}
	arrayName := directArrayArgumentName(arguments[0])
	return arrayName, arrayName != ""
}

func arrayVBA227ZeroBasedArraySourceLine(file parsedFile, proc sourceProcedure, arrayName string, beforeLine int) (int, bool) {
	arrayName = strings.ToLower(cleanIdentifier(arrayName))
	for line := beforeLine - 1; line >= proc.StartLine && line <= len(file.Lines); line-- {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), arrayName) {
				continue
			}
			callee := arrayCallName(rhs)
			if strings.EqualFold(callee, "strconv") || arrayVBA227ZeroBasedArrayFactoryFromLines(file, callee) || arrayVBA227ZeroBasedArrayFactory(file, callee) {
				return line, true
			}
			return 0, false
		}
	}
	return 0, false
}

func arrayVBA227ZeroBasedArraySourceBeforeOffset(file parsedFile, proc sourceProcedure, arrayName string, beforeLine, beforeOffset int) (int, int, bool) {
	arrayName = strings.ToLower(cleanIdentifier(arrayName))
	if arrayName == "" || beforeLine < proc.StartLine || beforeLine > len(file.Lines) {
		return 0, 0, false
	}
	for line := beforeLine; line >= proc.StartLine; line-- {
		spans := splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1]))
		for index := len(spans) - 1; index >= 0; index-- {
			span := spans[index]
			if line == beforeLine && (beforeOffset < 0 || span.start >= beforeOffset) {
				continue
			}
			lhs, rhs, indexed, assigned := arrayAssignment(span.text)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), arrayName) {
				continue
			}
			callee := arrayCallName(rhs)
			if strings.EqualFold(callee, "strconv") || arrayVBA227ZeroBasedArrayFactoryFromLines(file, callee) || arrayVBA227ZeroBasedArrayFactory(file, callee) {
				return line, span.start, true
			}
			return 0, 0, false
		}
	}
	return 0, 0, false
}

func arrayVBA227NoArrayMutationBetweenSourceAndProbe(file parsedFile, proc sourceProcedure, arrayName string, sourceLine, sourceOffset, probeLine, probeOffset int, ctx analysisContext) bool {
	if sourceLine <= 0 || probeLine < sourceLine || sourceOffset < 0 || probeOffset < 0 || probeLine == sourceLine && probeOffset <= sourceOffset {
		return false
	}
	if sourceLine == probeLine {
		return arrayVBA227NoArrayMutationWithinLine(file, proc, arrayName, sourceLine, sourceOffset, probeOffset, ctx)
	}
	if !arrayVBA227NoArrayMutationWithinLine(file, proc, arrayName, sourceLine, sourceOffset, len(arraySourceOrderStripComment(file.Lines[sourceLine-1])), ctx) {
		return false
	}
	if !arrayVBA227NoArrayMutationBetweenWithCalls(file, proc, arrayName, sourceLine+1, probeLine, ctx) {
		return false
	}
	return arrayVBA227NoArrayMutationWithinLine(file, proc, arrayName, probeLine, -1, probeOffset, ctx)
}

func arrayVBA227NoArrayMutationWithinLine(file parsedFile, proc sourceProcedure, arrayName string, line, startOffset, endOffset int, ctx analysisContext) bool {
	if line <= 0 || line > len(file.Lines) || startOffset < -1 || endOffset <= startOffset {
		return false
	}
	text := arraySourceOrderStripComment(file.Lines[line-1])
	for _, span := range splitRangeValueSourceStatementsWithOffsets(text) {
		if span.start <= startOffset || span.start >= endOffset {
			continue
		}
		if arrayVBA227SourceMutatesArray(span.text, arrayName) {
			return false
		}
	}
	lineStart := arraySourceOrderLineStartByteFrom(arraySourceOrderLineStarts(file.Source), line)
	if lineStart < 0 {
		return false
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line {
			continue
		}
		if call.Range.StartByte == 0 {
			if arrayVBA227CallMayInvalidateArray(proc, call, arrayName, ctx) {
				return false
			}
			continue
		}
		relative := call.Range.StartByte - lineStart
		if relative > startOffset && relative < endOffset && arrayVBA227CallMayInvalidateArray(proc, call, arrayName, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227NoArrayMutationBetween(file parsedFile, proc sourceProcedure, arrayName string, startLine, endLine int) bool {
	for line := max(startLine, proc.StartLine); line < endLine && line <= len(file.Lines); line++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			if arrayVBA227MutatesArray(source, arrayName) {
				return false
			}
		}
	}
	return true
}

func arrayVBA227NoArrayMutationBetweenWithCalls(file parsedFile, proc sourceProcedure, arrayName string, startLine, endLine int, ctx analysisContext) bool {
	if !arrayVBA227NoArrayMutationBetween(file, proc, arrayName, startLine, endLine) {
		return false
	}
	for line := max(startLine, proc.StartLine); line < endLine && line <= len(file.Lines); line++ {
		if arrayVBA227SourceMutatesArray(normalizedCodeLine(file.Lines[line-1]), arrayName) {
			return false
		}
	}
	arrayName = strings.ToLower(cleanIdentifier(arrayName))
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < startLine || call.Range.StartLine >= endLine {
			continue
		}
		if arrayVBA227CallMayInvalidateArray(proc, call, arrayName, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227NoArrayMutationAfterGuard(file parsedFile, proc sourceProcedure, arrayName string, guardLine, accessLine int, ctx analysisContext) bool {
	if guardLine <= 0 || accessLine < guardLine {
		return false
	}
	if guardLine != accessLine && !arrayVBA227NoArrayMutationBetweenWithCalls(file, proc, arrayName, guardLine+1, accessLine, ctx) {
		return false
	}
	if !arrayVBA227NoArrayMutationOnLine(file, proc, arrayName, accessLine, ctx) {
		return false
	}
	return arrayVBA227NoArrayMutationAfterLoopAccess(file, proc, arrayName, accessLine, ctx)
}

func arrayVBA227NoArrayMutationOnLine(file parsedFile, proc sourceProcedure, arrayName string, line int, ctx analysisContext) bool {
	if line <= 0 || line > len(file.Lines) {
		return false
	}
	if arrayVBA227SourceMutatesArray(normalizedCodeLine(file.Lines[line-1]), arrayName) {
		return false
	}
	arrayName = strings.ToLower(cleanIdentifier(arrayName))
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line {
			continue
		}
		if arrayVBA227CallMayInvalidateArray(proc, call, arrayName, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227NoArrayMutationAfterLoopAccess(file parsedFile, proc sourceProcedure, arrayName string, accessLine int, ctx analysisContext) bool {
	for loop := range proc.Statements.All() {
		switch loop.Kind {
		case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile:
		default:
			continue
		}
		if accessLine <= loop.Range.StartLine || accessLine >= loop.Range.EndLine {
			continue
		}
		if !arrayVBA227NoArrayMutationBetweenWithCalls(file, proc, arrayName, accessLine+1, loop.Range.EndLine, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227SourceMutatesArray(text, arrayName string) bool {
	for _, source := range splitRangeValueSourceStatements(text) {
		if arrayVBA227MutatesArray(source, arrayName) {
			return true
		}
		if _, body, ok := arrayIfThenParts(source); ok && strings.TrimSpace(body) != "" {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			branches := []string{thenBody}
			if hasElse {
				branches = append(branches, elseBody)
			}
			for _, branch := range branches {
				for _, nested := range splitRangeValueSourceStatements(branch) {
					if arrayVBA227MutatesArray(nested, arrayName) {
						return true
					}
				}
			}
		}
	}
	return false
}

func arrayVBA227CallMayInvalidateArray(proc sourceProcedure, call procedureir.CallSite, name string, ctx analysisContext) bool {
	if name == "" || arrayByRefCallIsReadOnly(call) {
		return false
	}
	if arrayVBA227SafeArrayLengthCall(proc, call, name, ctx) {
		return false
	}
	_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !resolved {
		for _, argument := range call.Arguments.Named {
			if directArrayArgumentName(argument.ValueText) == name {
				return true
			}
		}
		for _, argument := range arrayCallArgumentTexts(proc, call) {
			if directArrayArgumentName(argument) == name {
				return true
			}
		}
		return false
	}
	bindings, mapped := arrayCallArgumentBindings(proc, target, call)
	if !mapped {
		return arrayCallPassesDirectArrayArgument(proc, call, name)
	}
	for _, binding := range bindings {
		if directArrayArgumentName(binding.text) != name || binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			return true
		}
	}
	return false
}

func arrayVBA227NoScalarAssignmentBetween(file parsedFile, proc sourceProcedure, name string, startLine, endLine int, ctx analysisContext) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	for line := max(startLine, proc.StartLine); line < endLine && line <= len(file.Lines); line++ {
		if arrayVBA227SourceAssignsScalar(normalizedCodeLine(file.Lines[line-1]), name) {
			return false
		}
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < startLine || call.Range.StartLine >= endLine || arrayByRefCallIsReadOnly(call) {
			continue
		}
		_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !resolved {
			if arrayVBA227ExternalCallPassesNamedScalarByVal(file, proc, call, name) {
				continue
			}
			for _, argument := range call.Arguments.Named {
				if directArrayArgumentName(argument.ValueText) == name {
					return false
				}
			}
			for _, argument := range arrayCallArgumentTexts(proc, call) {
				if directArrayArgumentName(argument) == name {
					return false
				}
			}
			continue
		}
		bindings, mapped := arrayCallArgumentBindings(proc, target, call)
		if !mapped {
			if arrayCallPassesDirectArrayArgument(proc, call, name) {
				return false
			}
			continue
		}
		for _, binding := range bindings {
			if directArrayArgumentName(binding.text) != name || binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() {
				continue
			}
			if parameterIsByRefScalar(target.Params.valueAt(binding.parameterIndex)) {
				return false
			}
		}
	}
	return true
}

// arrayVBA227NoScalarAssignmentBeforeAccess extends the line-based scan with
// the physical access line. A colon-separated assignment on that line can
// occur after the ReDim and before the loop increment, even though the
// statement range exposes only the shared source line. Treating any matching
// assignment on the access line as invalid is deliberately conservative when
// the recovered statement column is unavailable.
func arrayVBA227NoScalarAssignmentBeforeAccess(file parsedFile, proc sourceProcedure, name string, startLine, accessLine int, ctx analysisContext) bool {
	if !arrayVBA227NoScalarAssignmentBetween(file, proc, name, startLine, accessLine, ctx) {
		return false
	}
	if accessLine < startLine || accessLine < 1 || accessLine > len(file.Lines) {
		return true
	}
	for _, source := range splitRangeValueSourceStatements(arraySourceOrderStripComment(file.Lines[accessLine-1])) {
		if arrayVBA227SourceAssignsScalar(source, name) {
			return false
		}
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != accessLine || arrayByRefCallIsReadOnly(call) {
			continue
		}
		if arrayVBA227CallMayMutateNamedScalar(file, proc, call, name, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227BranchRoot(proc sourceProcedure, statement procedureir.Statement) int {
	current := statement
	visited := map[int]bool{}
	for current.ParentID != 0 && !visited[current.ParentID] {
		visited[current.ParentID] = true
		parent := procedureStatementByID(proc, current.ParentID)
		if parent.ID == 0 {
			return 0
		}
		if parent.Kind == procedureir.StatementIf || parent.Kind == procedureir.StatementElseIf {
			return parent.ID
		}
		current = parent
	}
	return 0
}

func arrayVBA227IsAlternativeBranch(proc sourceProcedure, probe, candidate procedureir.Statement) bool {
	probeBranch := arrayVBA227BranchOwner(proc, probe)
	candidateBranch := arrayVBA227BranchOwner(proc, candidate)
	return probeBranch != 0 && candidateBranch != 0 && probeBranch != candidateBranch && arrayVBA227BranchRoot(proc, probe) == arrayVBA227BranchRoot(proc, candidate)
}

func arrayVBA227ScalarAssignmentsAllowedAfterSafeArrayProbe(text, name string, proc sourceProcedure, probe procedureir.Statement, line int) bool {
	for _, source := range splitRangeValueSourceStatements(text) {
		lhs, rhs, indexed, assigned := arrayAssignment(source)
		if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
			value, ok := integerLiteral(rhs)
			if !ok || value != 0 || !arrayVBA227IsAlternativeBranch(proc, probe, procedureStatementAtLine(proc, line)) {
				return false
			}
		}
		_, body, ok := arrayIfThenParts(source)
		if !ok || strings.TrimSpace(body) == "" {
			continue
		}
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		if !arrayVBA227ScalarAssignmentsAllowedAfterSafeArrayProbe(thenBody, name, proc, probe, line) || hasElse && !arrayVBA227ScalarAssignmentsAllowedAfterSafeArrayProbe(elseBody, name, proc, probe, line) {
			return false
		}
	}
	return true
}

func arrayVBA227NoScalarAssignmentAfterSafeArrayProbe(file parsedFile, proc sourceProcedure, name string, startLine, accessLine int, probe procedureir.Statement, ctx analysisContext) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" || probe.ID == 0 {
		return false
	}
	for line := max(startLine, proc.StartLine); line <= accessLine && line <= len(file.Lines); line++ {
		if !arrayVBA227ScalarAssignmentsAllowedAfterSafeArrayProbe(arraySourceOrderStripComment(file.Lines[line-1]), name, proc, probe, line) {
			return false
		}
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < startLine || call.Range.StartLine > accessLine || arrayByRefCallIsReadOnly(call) {
			continue
		}
		if arrayVBA227CallMayMutateNamedScalar(file, proc, call, name, ctx) {
			return false
		}
	}
	return true
}

func arrayVBA227SourceAssignsScalar(text, name string) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	for _, source := range splitRangeValueSourceStatements(text) {
		lhs, _, indexed, assigned := arrayAssignment(source)
		if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
			return true
		}
		_, body, ok := arrayIfThenParts(source)
		if !ok || strings.TrimSpace(body) == "" {
			continue
		}
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		if arrayVBA227SourceAssignsScalar(thenBody, name) || hasElse && arrayVBA227SourceAssignsScalar(elseBody, name) {
			return true
		}
	}
	return false
}

func arrayVBA227AccessUsesBound(file parsedFile, arrayName string, line int, variables map[string]arrayVariable, kind string) bool {
	if line <= 0 || line > len(file.Lines) {
		return false
	}
	arrayName = strings.ToLower(cleanIdentifier(arrayName))
	for _, use := range arrayIndexedUsesForSource(normalizedCodeLine(file.Lines[line-1]), variables) {
		if strings.EqualFold(cleanIdentifier(use.name), arrayName) && arrayUseHasSelfBoundQuery(use, kind) {
			return true
		}
	}
	return false
}

func arrayVBA227AllIndexedUsesUseLowerBound(file parsedFile, arrayName string, line int, variables map[string]arrayVariable) bool {
	if line <= 0 || line > len(file.Lines) {
		return false
	}
	want := strings.ToLower(cleanIdentifier(arrayName))
	found := false
	for _, use := range arrayIndexedUsesForSource(normalizedCodeLine(file.Lines[line-1]), variables) {
		if !strings.EqualFold(cleanIdentifier(use.name), want) {
			continue
		}
		found = true
		if !arrayUseHasSelfBoundQuery(use, "lbound") {
			return false
		}
	}
	return found
}

func arrayVBA227ZeroBasedArrayFactory(file parsedFile, name string) bool {
	if name == "" {
		return false
	}
	var target sourceProcedure
	found := 0
	for procedure := range file.procedureView().All() {
		if !strings.EqualFold(procedure.Name, name) || procedure.ProcedureKind != procedureir.ProcedureFunction && procedure.ProcedureKind != procedureir.ProcedurePropertyGet {
			continue
		}
		target = procedure
		found++
	}
	if found != 1 {
		return arrayVBA227ZeroBasedArrayFactoryFromLines(file, name)
	}
	source, ok := arrayProcedureReturnSource(file, target)
	if !ok {
		return arrayVBA227ZeroBasedArrayFactoryFromLines(file, name)
	}
	variables := arrayVariables(file, target, file.moduleDecls())
	variable, ok := variables[strings.ToLower(cleanIdentifier(source))]
	if !ok || !variable.isArray || !isByteArrayVariable(variable) {
		return arrayVBA227ZeroBasedArrayFactoryFromLines(file, name)
	}
	foundRedim := false
	for statement := range target.Statements.All() {
		if statement.Kind != procedureir.StatementReDim {
			continue
		}
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
			return arrayVBA227ZeroBasedArrayFactoryFromLines(file, name)
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct || !strings.EqualFold(cleanIdentifier(redim.name), source) {
				continue
			}
			if !arrayVBA227RedimStartsAtZero(redim.dimensions) {
				return arrayVBA227ZeroBasedArrayFactoryFromLines(file, name)
			}
			foundRedim = true
		}
	}
	if foundRedim {
		return true
	}
	return arrayVBA227ZeroBasedArrayFactoryFromLines(file, name)
}

func arrayVBA227ZeroBasedArrayFactoryFromLines(file parsedFile, name string) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	headerRe := regexp.MustCompile(`(?i)^\s*(?:(?:private|public|friend|static)\s+)*function\s+([A-Za-z_]\w*)\b`)
	for start := 0; start < len(file.Lines); start++ {
		header := strings.TrimSpace(normalizedCodeLine(file.Lines[start]))
		match := headerRe.FindStringSubmatch(header)
		if len(match) != 2 || !strings.EqualFold(match[1], name) {
			continue
		}
		end := len(file.Lines)
		for line := start + 1; line < len(file.Lines); line++ {
			if strings.EqualFold(strings.TrimSpace(normalizedCodeLine(file.Lines[line])), "end function") {
				end = line
				break
			}
		}
		source := ""
		for line := start + 1; line < end; line++ {
			text := strings.TrimSpace(normalizedCodeLine(file.Lines[line]))
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
				source = directArrayArgumentName(rhs)
				break
			}
		}
		if source == "" {
			return false
		}
		foundRedim := false
		for line := start + 1; line < end; line++ {
			text := strings.TrimSpace(normalizedCodeLine(file.Lines[line]))
			match := arrayRedimRe.FindStringSubmatch(text)
			if len(match) == 0 {
				continue
			}
			if strings.TrimSpace(match[1]) != "" {
				return false
			}
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct || !strings.EqualFold(cleanIdentifier(redim.name), source) {
					continue
				}
				if !arrayVBA227RedimStartsAtZero(redim.dimensions) {
					return false
				}
				foundRedim = true
			}
		}
		return foundRedim
	}
	return false
}

func arrayVBA227RedimStartsAtZero(dimensions string) bool {
	parts := splitArgs(dimensions)
	if len(parts) == 0 {
		return false
	}
	first := strings.ToLower(strings.TrimSpace(canonicalArrayBoundExpression(parts[0])))
	return strings.HasPrefix(first, "0to")
}

func arrayVBA227UBoundLengthSource(rhs string) (string, bool) {
	rhs = strings.TrimSpace(rhs)
	if strings.EqualFold(arrayCallName(rhs), "cbyte") {
		arguments, ok := arraySimpleCallArguments(rhs)
		if !ok || len(arguments) != 1 {
			return "", false
		}
		rhs = strings.TrimSpace(arguments[0])
	}
	match := arrayUBoundPlusOneRe.FindStringSubmatch(rhs)
	if len(match) != 2 {
		return "", false
	}
	return strings.ToLower(cleanIdentifier(match[1])), true
}

func arrayVBA227HasZeroBasedStrConvAssignment(file parsedFile, proc sourceProcedure, arrayName string, beforeLine int) bool {
	target := procedureStatementAtLine(proc, beforeLine)
	if target.ID <= 0 {
		return false
	}
	for line := beforeLine; line >= proc.StartLine && line <= len(file.Lines); line-- {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || !strings.EqualFold(cleanIdentifier(lhs), arrayName) {
				continue
			}
			if indexed {
				continue
			}
			return strings.EqualFold(arrayCallName(rhs), "strconv") && arrayVBA227StatementLineDominates(proc, line, target)
		}
	}
	return false
}

func arrayVBA227MutatesArray(text, arrayName string) bool {
	for _, source := range splitRangeValueSourceStatements(text) {
		if arrayVBA227MutatesArraySource(source, arrayName) {
			return true
		}
	}
	return false
}

func arrayVBA227MutatesArraySource(text, arrayName string) bool {
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct {
				legacy := arrayRedimClauseRe.FindStringSubmatch(clause)
				if len(legacy) == 0 {
					continue
				}
				redim = directArrayRedimClause{name: legacy[1]}
			}
			if strings.EqualFold(cleanIdentifier(redim.name), arrayName) {
				return true
			}
		}
	}
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) > 0 {
		for _, target := range splitArgs(match[1]) {
			if strings.EqualFold(cleanIdentifier(target), arrayName) {
				return true
			}
		}
	}
	lhs, _, indexed, assigned := arrayAssignment(text)
	if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), arrayName) {
		return true
	}
	if _, body, ok := arrayIfThenParts(text); ok && strings.TrimSpace(body) != "" {
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		if arrayVBA227MutatesArray(thenBody, arrayName) || hasElse && arrayVBA227MutatesArray(elseBody, arrayName) {
			return true
		}
	}
	return false
}

func arrayVBA227SafeArrayLengthHelperMayMutateArray(proc sourceProcedure, line int, text, arrayName string) bool {
	if arrayVBA227MutatesArray(text, arrayName) {
		return true
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line || arrayByRefCallIsReadOnly(call) {
			continue
		}
		if strings.EqualFold(cleanIdentifier(call.Callee.BaseName), "varptrarray") {
			continue
		}
		if arrayCallPassesDirectArrayArgument(proc, call, arrayName) {
			return true
		}
	}
	return false
}

func arrayVBA227StatementLineDominates(proc sourceProcedure, line int, target procedureir.Statement) bool {
	if proc.Graph == nil || line <= 0 || target.ID <= 0 {
		return false
	}
	source := procedureStatementAtLine(proc, line)
	if source.ID <= 0 || source.ID == target.ID {
		return false
	}
	sourceBlock, sourceOK := proc.Graph.BlockForStatement(source.ID)
	targetBlock, targetOK := proc.Graph.BlockForStatement(target.ID)
	if !sourceOK || !targetOK {
		return false
	}
	if sourceBlock.ID == targetBlock.ID {
		return source.Range.StartLine < target.Range.StartLine
	}
	return proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).Dominates(sourceBlock.ID, targetBlock.ID)
}

func arrayVBA227StatementLineDominatesNormalExit(proc sourceProcedure, line int) bool {
	if proc.Graph == nil || line <= 0 {
		return false
	}
	source := procedureStatementAtLine(proc, line)
	if source.ID == 0 {
		return false
	}
	sourceBlock, ok := proc.Graph.BlockForStatement(source.ID)
	if !ok {
		return false
	}
	return proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).Dominates(sourceBlock.ID, proc.Graph.NormalExit)
}

func arrayVBA227StatementLineDominatesOrUnconditional(proc sourceProcedure, line int, target procedureir.Statement) bool {
	if arrayVBA227StatementLineDominates(proc, line, target) {
		return true
	}
	if proc.Graph == nil || line <= 0 || target.ID <= 0 {
		return false
	}
	source := procedureStatementAtLine(proc, line)
	if source.ID <= 0 || source.ParentID != 0 || source.Range.StartLine >= target.Range.StartLine {
		return false
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementGoTo && statement.Kind != procedureir.StatementOnError && statement.Kind != procedureir.StatementResume && statement.Kind != procedureir.StatementLabel {
			continue
		}
		if statement.Range.StartLine >= source.Range.StartLine && statement.Range.StartLine < target.Range.StartLine {
			return false
		}
	}
	return true
}

// arrayVBA227FilterForBodyIndexFindings removes only the unallocated/empty
// index observations for a loop body whose For bound necessarily succeeded
// before the body could run. Source-line CFG blocks include a For header and
// its nested body in one scan, so the edge refinement is applied too late for
// the first body visit. A bound query embedded in a proven SafeArrayLen access
// is also removed when that query is itself known to be safe; unrelated bound
// findings and any known lower/upper-bound violation remain.
func arrayVBA227FilterForBodyIndexFindings(findings []Finding, file parsedFile, proc sourceProcedure, line int, state arrayFlowState, variables map[string]arrayVariable, ctx analysisContext, resumeNextBefore []bool, vba227Graph *vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges) []Finding {
	if line <= 0 {
		return findings
	}
	resumeNext := arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line)
	proven := map[string]bool{}
	provenNonEmpty := map[string]bool{}
	provenBounds := map[string]map[string]bool{}
	for statement := range proc.Statements.All() {
		if line <= statement.Range.StartLine || line >= statement.Range.EndLine {
			continue
		}
		switch statement.Kind {
		case procedureir.StatementFor:
			if !resumeNext {
				if name, ok := arrayVBA227DerivedZeroBasedLoopArray(file, proc, statement, variables, ctx); ok {
					proven[name] = true
					provenNonEmpty[name] = true
				}
			}
		case procedureir.StatementDo, procedureir.StatementWhile:
			if !resumeNext {
				name, ok := arrayVBA227DerivedDoWhileArray(file, proc, statement, line, variables, ctx)
				if ok {
					proven[name] = true
					provenNonEmpty[name] = true
					for _, kind := range []string{"lbound", "ubound"} {
						if !arrayVBA227AccessUsesBound(file, name, line, variables, kind) {
							continue
						}
						if provenBounds[name] == nil {
							provenBounds[name] = map[string]bool{}
						}
						provenBounds[name][kind] = true
					}
				}
			}
		}
		header := strings.TrimSpace(statement.Text)
		if newline := strings.IndexAny(header, "\r\n"); newline >= 0 {
			header = strings.TrimSpace(header[:newline])
		}
		header = strings.TrimSpace(normalizedCodeLine(header))
		argument := ""
		if _, _, countSource, _, ok := arrayForCountHeader(header); ok && !resumeNext {
			for name, value := range state {
				if value.allocationCountSource == "" || !arrayCountExpressionMatches(countSource, value.allocationCountSource) {
					continue
				}
				variable, known := variables[name]
				if known && (variable.isArray || variable.isVariant) {
					proven[name] = true
					provenNonEmpty[name] = true
				}
			}
		}
		if match := arrayForScalarBoundRe.FindStringSubmatch(header); len(match) == 2 && !resumeNext {
			bound, known := state[strings.ToLower(cleanIdentifier(match[1]))]
			if known {
				argument = bound.safeBoundProbe
			}
		}
		name := strings.ToLower(cleanIdentifier(argument))
		variable, known := variables[name]
		if !resumeNext && name != "" && known && (variable.isArray || variable.isVariant) {
			proven[name] = true
			provenNonEmpty[name] = true
		}
		if match := arrayForUBoundRe.FindStringSubmatch(header); len(match) == 3 && !resumeNext {
			name := strings.ToLower(cleanIdentifier(match[2]))
			variable, variableKnown := variables[name]
			_, valueKnown := state[name]
			start, startKnown := integerLiteral(match[1])
			if variableKnown && (variable.isArray || variable.isVariant) && valueKnown && startKnown && start >= 0 {
				// The For body is reachable only after UBound succeeded. When
				// the loop starts at a nonnegative index, an empty array cannot
				// reach the body (its upper bound is below the start). The known
				// lower bound, when available, additionally proves the index is
				// not below the array's lower bound.
				proven[name] = true
				provenNonEmpty[name] = true
			}
		}
		if match := arrayForBoundsRe.FindStringSubmatch(header); len(match) == 3 && strings.EqualFold(match[1], match[2]) {
			name := strings.ToLower(cleanIdentifier(match[1]))
			variable, variableKnown := variables[name]
			value, valueKnown := state[name]
			if variableKnown && (variable.isArray || variable.isVariant) && valueKnown {
				access := procedureStatementAtLine(proc, line)
				if access.ID != 0 {
					if arrayVBA227BoundsCanFail(value) && arrayVBA227ResumeCanReachForBody(proc, vba227Graph, resumeNextEdges, statement, access) {
						continue
					}
					if value.resumeBoundsFailurePossible && arrayVBA227ResumeCanReachPriorBoundsBody(proc, vba227Graph, resumeNextEdges, statement.Range.StartLine, name, access, variables) {
						continue
					}
				}
				// Reaching the body means both bounds queries completed and the
				// default positive step found at least one value between LBound
				// and UBound. This also proves a late-bound Variant snapshot is
				// non-empty on the body path, while the bound observations remain
				// findings on the header itself.
				proven[name] = true
				provenNonEmpty[name] = true
			}
		}
	}
	if len(proven) == 0 && len(provenNonEmpty) == 0 {
		return findings
	}
	filtered := findings[:0]
	for _, finding := range findings {
		remove := false
		if finding.Code == "VBA227" {
			for name := range proven {
				if finding.arrayOperationKey == arrayIndexOperationKey(name, "unallocated") ||
					provenNonEmpty[name] && finding.arrayOperationKey == arrayIndexOperationKey(name, "empty") {
					remove = true
					break
				}
				for kind := range provenBounds[name] {
					if finding.arrayOperationKey == arrayBoundOperationKey(kind, name, "unallocated") {
						remove = true
						break
					}
				}
			}
		}
		if !remove {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

// arrayVBA227FilterConditionalBodyIndexFindings removes an unallocated-array
// observation from a source-line CFG block when the block is inside the true
// body of a matching allocation guard. With blocks can keep the If header and
// its first body statement in one CFG block, so the edge refinement runs after
// the body has already been visited. Positive Not Not descriptor guards also
// make their bound queries safe, while their possible-empty fact remains
// conservative. Unrelated conditions and Else bodies remain conservative.
func arrayVBA227FilterConditionalBodyIndexFindings(findings []Finding, file parsedFile, proc sourceProcedure, line int, state arrayFlowState, variables map[string]arrayVariable, ctx analysisContext, resumeNextBefore []bool) []Finding {
	if line <= 0 || arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
		return findings
	}
	statement := procedureStatementAtLine(proc, line)
	if statement.ID == 0 {
		return findings
	}
	access := statement
	proven := map[string]bool{}
	provenNonEmpty := map[string]bool{}
	provenBounds := map[string]map[string]bool{}
	if name, guardLine, ok := arrayVBA227EnclosingNotNotGuard(proc, access, variables); ok && arrayVBA227NoArrayMutationAfterGuard(file, proc, name, guardLine, line, ctx) {
		proven[name] = true
		provenBounds[name] = map[string]bool{"lbound": true, "ubound": true}
	}
	inAlternative := false
	visited := map[int]bool{}
	for statement.ParentID != 0 && !visited[statement.ParentID] {
		visited[statement.ParentID] = true
		parent := procedureStatementByID(proc, statement.ParentID)
		if parent.ID == 0 {
			break
		}
		switch parent.Kind {
		case procedureir.StatementElse:
			inAlternative = true
		case procedureir.StatementElseIf:
			if !inAlternative {
				if name, positive := arrayNotNotByteArrayGuardTarget(parent.Text, variables); positive && arrayVBA227NoArrayMutationAfterGuard(file, proc, name, parent.Range.StartLine, line, ctx) {
					proven[name] = true
					provenBounds[name] = map[string]bool{"lbound": true, "ubound": true}
				}
				if lengthName, positive := arrayVBA227PositiveGuardConditionSource(parent, variables); positive {
					for name, value := range state {
						if value.allocationCountSource != "" && arrayCountExpressionMatches(lengthName, value.allocationCountSource) {
							proven[name] = true
						}
					}
					if name, ok := arrayVBA227UBoundLengthArray(proc, &parent, &access, lengthName, variables); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227AllocationProbeLengthArray(file, proc, access, state, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveSafeArrayLengthArray(file, proc, &access, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveConditionalReDimArray(file, proc, &access, &parent, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
				}
			}
			inAlternative = true
		case procedureir.StatementIf:
			if !inAlternative {
				if name, positive := arrayNotNotByteArrayGuardTarget(parent.Text, variables); positive && arrayVBA227NoArrayMutationAfterGuard(file, proc, name, parent.Range.StartLine, line, ctx) {
					proven[name] = true
					provenBounds[name] = map[string]bool{"lbound": true, "ubound": true}
				}
				if lengthName, positive := arrayVBA227PositiveGuardConditionSource(parent, variables); positive {
					for name, value := range state {
						if value.allocationCountSource != "" && arrayCountExpressionMatches(lengthName, value.allocationCountSource) {
							proven[name] = true
						}
					}
					if name, ok := arrayVBA227UBoundLengthArray(proc, &parent, &access, lengthName, variables); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227AllocationProbeLengthArray(file, proc, access, state, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveSafeArrayLengthArray(file, proc, &access, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveConditionalReDimArray(file, proc, &access, &parent, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
				}
			}
		}
		statement = parent
	}
	if len(proven) == 0 {
		text := strings.TrimSpace(access.Text)
		if text == "" && line >= 1 && line <= len(file.Lines) {
			text = strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		}
		if condition, body, ok := arrayIfThenParts(text); ok && body != "" {
			_, _, hasElse := arrayIfThenBodyParts(body)
			if !hasElse {
				if name, probe := arrayNotNotByteArrayGuardTarget(condition, variables); probe && arrayVBA227NoArrayMutationAfterGuard(file, proc, name, line, line, ctx) {
					proven[name] = true
					provenBounds[name] = map[string]bool{"lbound": true, "ubound": true}
				}
				if lengthName, positive := arrayVBA227PositiveLengthCondition(condition); positive {
					if name, probe := arrayVBA227AllocationProbeLengthArray(file, proc, access, state, lengthName, variables, ctx); probe {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, probe := arrayVBA227PositiveSafeArrayLengthArray(file, proc, &access, lengthName, variables, ctx); probe {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, probe := arrayVBA227PositiveConditionalReDimArray(file, proc, &access, &access, lengthName, variables, ctx); probe {
						proven[name] = true
						provenNonEmpty[name] = true
					}
				}
			}
		}
	}
	if len(proven) == 0 {
		return findings
	}
	filtered := findings[:0]
	for _, finding := range findings {
		remove := false
		if finding.Code == "VBA227" {
			for name := range proven {
				if finding.arrayOperationKey == arrayIndexOperationKey(name, "unallocated") ||
					provenNonEmpty[name] && finding.arrayOperationKey == arrayIndexOperationKey(name, "empty") {
					remove = true
					break
				}
				for kind := range provenBounds[name] {
					if finding.arrayOperationKey == arrayBoundOperationKey(kind, name, "unallocated") {
						remove = true
						break
					}
				}
			}
		}
		if !remove {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

func arrayVBA227EnclosingNotNotGuard(proc sourceProcedure, access procedureir.Statement, variables map[string]arrayVariable) (string, int, bool) {
	current := access
	skippedIfID := 0
	visited := map[int]bool{}
	for current.ParentID != 0 && !visited[current.ParentID] {
		visited[current.ParentID] = true
		parent := procedureStatementByID(proc, current.ParentID)
		if parent.ID == 0 {
			break
		}
		switch parent.Kind {
		case procedureir.StatementElse:
			// Skip only the If that owns this Else branch. An enclosing
			// positive guard may still dominate an unrelated nested Else.
			skippedIfID = parent.ParentID
		case procedureir.StatementElseIf:
			if parent.ID == skippedIfID {
				skippedIfID = 0
			} else if name, ok := arrayNotNotByteArrayGuardTarget(parent.Text, variables); ok {
				return name, parent.Range.StartLine, true
			} else {
				skippedIfID = parent.ParentID
			}
		case procedureir.StatementIf:
			if parent.ID == skippedIfID {
				skippedIfID = 0
			} else if name, ok := arrayNotNotByteArrayGuardTarget(parent.Text, variables); ok {
				return name, parent.Range.StartLine, true
			}
		}
		current = parent
	}
	return "", 0, false
}

// arrayVBA227FilterSuccessfulBoundsGuardBodyIndexFindings removes the
// redundant unallocated-array observation on the fallthrough line after a
// single-line If that evaluates UBound or LBound before terminating its true
// branch. Reaching the following line means the bound query completed
// normally, even when the CFG visits the body before applying the guard edge.
func arrayVBA227FilterSuccessfulBoundsGuardBodyIndexFindings(findings []Finding, file parsedFile, proc sourceProcedure, line int, variables map[string]arrayVariable, resumeNextBefore []bool) []Finding {
	if line <= 1 || line > len(file.Lines) || arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
		return findings
	}
	previous := normalizedCodeLine(file.Lines[line-2])
	condition, body, ok := arrayIfThenParts(previous)
	if !ok || body == "" || !arrayVBA227HasBoundsCondition(condition) {
		return findings
	}
	proven := map[string]bool{}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(condition, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		variable, known := variables[name]
		if known && variable.isArray && name != "" {
			proven[name] = true
		}
	}
	if len(proven) == 0 {
		return findings
	}
	filtered := findings[:0]
	for _, finding := range findings {
		remove := false
		if finding.Code == "VBA227" {
			for name := range proven {
				if finding.arrayOperationKey == arrayIndexOperationKey(name, "unallocated") {
					remove = true
					break
				}
			}
		}
		if !remove {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

// arrayVBA227FilterSuccessfulIndexedConditionBodyFindings removes the
// redundant unallocated-array observation in the true body after a
// multi-line If condition has already indexed the same array. Reaching that
// body means the condition's indexed access completed normally. Keep the
// proof tied to the immediately preceding If header and reject Else bodies;
// an unrelated earlier access must not establish allocation for a later one.
// A reachable error handler that can resume into the body also disables this
// normal-path proof because the condition may have failed before the re-entry.
func arrayVBA227FilterSuccessfulIndexedConditionBodyFindings(findings []Finding, file parsedFile, proc sourceProcedure, line int, variables map[string]arrayVariable, resumeNextBefore []bool, vba227Graph *vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges) []Finding {
	if line <= 1 || line > len(file.Lines) || arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
		return findings
	}
	condition, body, ok := arrayIfThenParts(normalizedCodeLine(file.Lines[line-2]))
	if !ok || body != "" {
		return findings
	}

	statement := procedureStatementAtLine(proc, line)
	if statement.ID == 0 {
		return findings
	}
	access := statement
	guardFound := false
	var guard procedureir.Statement
	visited := map[int]bool{}
	for statement.ParentID != 0 && !visited[statement.ParentID] {
		visited[statement.ParentID] = true
		parent := procedureStatementByID(proc, statement.ParentID)
		if parent.ID == 0 {
			return findings
		}
		if parent.Kind == procedureir.StatementElse {
			return findings
		}
		if (parent.Kind == procedureir.StatementIf || parent.Kind == procedureir.StatementElseIf) && parent.Range.StartLine == line-1 {
			guardFound = true
			guard = parent
			break
		}
		statement = parent
	}
	if !guardFound {
		return findings
	}

	proven := map[string]bool{}
	for _, use := range arrayIndexedUsesForSource(condition, variables) {
		if len(use.args) == 0 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(use.name))
		variable, known := variables[name]
		if known && (variable.isArray || variable.isVariant) && name != "" {
			proven[name] = true
		}
	}
	if len(proven) == 0 {
		return findings
	}
	if arrayVBA227ResumeCanReachIndexedConditionBody(proc, vba227Graph, resumeNextEdges, guard, access) {
		return findings
	}

	filtered := findings[:0]
	for _, finding := range findings {
		remove := false
		if finding.Code == "VBA227" {
			for name := range proven {
				if finding.arrayOperationKey == arrayIndexOperationKey(name, "unallocated") {
					remove = true
					break
				}
			}
		}
		if !remove {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

func arrayVBA227StatementAlwaysRaises(text string) bool {
	trimmed := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{"err.raise ", "err.raise(", "call err.raise ", "call err.raise("} {
		if strings.HasPrefix(trimmed, prefix) {
			return true
		}
	}
	return trimmed == "err.raise" || trimmed == "call err.raise"
}

func arrayNonEmptyGuardState(state arrayFlowState, condition string, variables map[string]arrayVariable) (arrayFlowState, bool) {
	condition = strings.TrimSpace(condition)
	if strings.HasPrefix(strings.ToLower(condition), "if ") {
		condition = strings.TrimSpace(condition[3:])
	}
	match := arrayEmptyGuardRe.FindStringSubmatch(condition)
	if len(match) != 2 {
		return state, false
	}
	name := strings.ToLower(cleanIdentifier(match[1]))
	variable, known := variables[name]
	if !known || !variable.isArray {
		return state, false
	}
	value, known := state[name]
	if !known {
		return state, false
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	value.mayBeEmpty = false
	updated[name] = value
	return updated, true
}

// arrayVBA227FilterNestedBoundIndexFindings removes only the redundant
// unallocated-index observation from an expression such as
// data(LBound(data)). The bound query itself remains a finding: if it fails,
// the indexed expression is never evaluated. Keep the proof conservative when
// the same array has another indexed use on the source line, because the
// normalized finding range cannot distinguish those uses.
func arrayVBA227FilterNestedBoundIndexFindings(findings []Finding, text string, variables map[string]arrayVariable) []Finding {
	indexedCounts := make(map[string]int)
	proofCandidates := make(map[string]bool)
	for _, use := range arrayIndexedUsesForSource(text, variables) {
		if len(use.args) == 0 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(use.name))
		if name == "" {
			continue
		}
		indexedCounts[name]++
		variable, known := variables[name]
		if known && variable.isArray && arrayUseHasSelfBoundsQuery(use) {
			proofCandidates[name] = true
		}
	}
	boundFailures := make(map[string]bool)
	for _, finding := range findings {
		if finding.Code != "VBA227" {
			continue
		}
		for _, kind := range []string{"lbound", "ubound"} {
			for name := range proofCandidates {
				if finding.arrayOperationKey == arrayBoundOperationKey(kind, name, "unallocated") {
					boundFailures[name] = true
				}
			}
		}
	}
	proofKeys := make(map[string]bool)
	for name := range proofCandidates {
		if indexedCounts[name] == 1 && boundFailures[name] {
			proofKeys[arrayIndexOperationKey(name, "unallocated")] = true
		}
	}
	if len(proofKeys) == 0 {
		return findings
	}
	filtered := findings[:0]
	for _, finding := range findings {
		if finding.Code == "VBA227" && proofKeys[finding.arrayOperationKey] {
			continue
		}
		filtered = append(filtered, finding)
	}
	return filtered
}

func arrayUseHasSelfBoundsQuery(use arrayUse) bool {
	return arrayUseHasSelfBoundQuery(use, "lbound") || arrayUseHasSelfBoundQuery(use, "ubound")
}

func arrayUseHasSelfBoundQuery(use arrayUse, kind string) bool {
	for _, argument := range use.args {
		for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(argument, -1) {
			if strings.EqualFold(bound[1], kind) && strings.EqualFold(cleanIdentifier(bound[2]), use.name) {
				return true
			}
		}
	}
	return false
}

func arrayVBA227HasBoundsCondition(text string) bool {
	return strings.HasPrefix(strings.ToLower(strings.TrimSpace(text)), "if ") && arrayBoundCallRe.MatchString(text)
}

func arrayVBA227HasSuccessfulBoundsExpression(text string) bool {
	trimmed := strings.ToLower(strings.TrimSpace(text))
	// Conditional expressions do not dominate the following source line:
	// an ElseIf condition can be skipped when an earlier branch is taken.
	// Branch-specific allocation facts belong to the CFG edge transfer, not
	// to this normal-path statement refinement.
	if strings.HasPrefix(trimmed, "if ") || strings.HasPrefix(trimmed, "elseif ") || strings.HasPrefix(trimmed, "else if ") {
		return false
	}
	seen := make(map[string]uint8)
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		if name == "" {
			continue
		}
		if strings.EqualFold(bound[1], "lbound") {
			seen[name] |= 1
		} else {
			seen[name] |= 2
		}
	}
	for _, bounds := range seen {
		if bounds != 0 {
			return true
		}
	}
	return false
}

func arrayVBA227LoopBodyEndLine(proc sourceProcedure, line int) int {
	if line <= 0 {
		return 0
	}
	endLine := 0
	for statement := range proc.Statements.All() {
		switch statement.Kind {
		case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile:
		default:
			continue
		}
		if line > statement.Range.StartLine && line < statement.Range.EndLine &&
			(endLine == 0 || statement.Range.EndLine < endLine) {
			endLine = statement.Range.EndLine
		}
	}
	return endLine
}

func arrayVBA227AttachConditionalReDimState(state arrayFlowState, proc sourceProcedure, text string, line int, variables map[string]arrayVariable) arrayFlowState {
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(state) == 0 || line <= 0 || len(match) == 0 {
		return state
	}
	guard, ok := arrayVBA227ConditionalReDimGuard(proc, line, variables)
	if !ok {
		return state
	}
	var updated arrayFlowState
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if !direct {
			legacy := arrayRedimClauseRe.FindStringSubmatch(clause)
			if len(legacy) == 0 {
				continue
			}
			redim = directArrayRedimClause{name: legacy[1], dimensions: legacy[2]}
		}
		name := strings.ToLower(cleanIdentifier(redim.name))
		variable, knownVariable := variables[name]
		value, knownValue := state[name]
		if !knownVariable || !variable.isArray || !knownValue || value.kind != arrayAllocated || !value.knownArray {
			continue
		}
		if value.conditionalAllocationSource != "" && value.conditionalAllocationSource != guard {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.conditionalAllocationSource = guard
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227ConditionalReDimGuard(proc sourceProcedure, line int, variables map[string]arrayVariable) (string, bool) {
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementReDim || statement.Range.StartLine != line {
			continue
		}
		parent := procedureStatementByID(proc, statement.ParentID)
		if parent.Kind != procedureir.StatementIf && parent.Kind != procedureir.StatementElseIf {
			return "", false
		}
		if guard, ok := arrayVBA227ScalarConditionSource(parent, variables); ok {
			return guard, true
		}
		return arrayVBA227PositiveScalarConditionSource(parent, variables)
	}
	return "", false
}

// arrayVBA227AttachAllocationFlagState records a local ready flag only when
// the flag assignment is a sibling of a plain ReDim and the incoming state
// proves that ReDim completed. Keeping the relation on the array value lets
// it survive joins where the flag itself is not modeled by arrayFlowState.
func arrayVBA227AttachAllocationFlagState(file parsedFile, proc sourceProcedure, text string, line int, input, output arrayFlowState, variables map[string]arrayVariable) arrayFlowState {
	if len(output) == 0 || line <= 0 || arrayProcedureHasErrorHandling(proc) {
		return output
	}
	lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(text))
	if !assigned || indexed {
		return output
	}
	flag := strings.ToLower(cleanIdentifier(lhs))
	variable, known := variables[flag]
	if !known || variable.isArray || variable.isVariant || variable.isObject || !strings.EqualFold(strings.TrimSpace(variable.typ), "Boolean") {
		return output
	}
	updated := output
	cloned := false
	for name, value := range output {
		if value.allocationFlagSource != flag {
			continue
		}
		if !cloned {
			updated = cloneArrayState(output)
			cloned = true
		}
		value.allocationFlagSource = ""
		updated[name] = value
	}
	if !strings.EqualFold(strings.TrimSpace(rhs), "true") {
		return updated
	}
	target, ok := arrayVBA227AllocationFlagTarget(file, proc, line, flag, variables)
	if !ok {
		return updated
	}
	value, known := input[target]
	if !known || value.kind != arrayAllocated || !value.knownArray || value.mayBeEmpty {
		return updated
	}
	targetValue, targetKnown := updated[target]
	if !targetKnown {
		return updated
	}
	if !cloned {
		updated = cloneArrayState(output)
		targetValue = updated[target]
	}
	targetValue.allocationFlagSource = flag
	updated[target] = targetValue
	return updated
}

func arrayVBA227AllocationFlagTarget(file parsedFile, proc sourceProcedure, line int, flag string, variables map[string]arrayVariable) (string, bool) {
	parentID := 0
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementAssignment || statement.Range.StartLine != line {
			continue
		}
		text := strings.TrimSpace(statement.Text)
		if text == "" && line >= 1 && line <= len(file.Lines) {
			text = strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if assigned && !indexed && strings.EqualFold(strings.TrimSpace(rhs), "true") && strings.EqualFold(cleanIdentifier(lhs), flag) {
			parentID = statement.ParentID
			break
		}
	}
	if parentID == 0 {
		return "", false
	}
	target := ""
	for statement := range proc.Statements.All() {
		if statement.ParentID != parentID || statement.Kind != procedureir.StatementReDim || statement.Range.StartLine >= line {
			continue
		}
		text := strings.TrimSpace(statement.Text)
		if text == "" && statement.Range.StartLine >= 1 && statement.Range.StartLine <= len(file.Lines) {
			text = strings.TrimSpace(normalizedCodeLine(file.Lines[statement.Range.StartLine-1]))
		}
		if text == "" || !strings.Contains(strings.ToLower(text), "redim") {
			return "", false
		}
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
			return "", false
		}
		clauses := splitArgs(match[2])
		if len(clauses) != 1 {
			return "", false
		}
		redim, direct := parseDirectArrayRedimClause(clauses[0])
		if !direct {
			return "", false
		}
		name := strings.ToLower(cleanIdentifier(redim.name))
		variable, known := variables[name]
		if !known || !variable.isArray || variable.fixed || name == "" {
			return "", false
		}
		if target != "" && target != name {
			return "", false
		}
		target = name
	}
	return target, target != ""
}

// arrayVBA227ClearConditionalAllocationGuards forgets a conditional ReDim fact
// when the scalar that controls it is assigned outside a matching guard. A
// later equality check is only useful if the value being checked is still the
// value that controlled the allocation; retaining the fact across an
// unrelated assignment would turn a path-sensitive proof into an unsound one.
func arrayVBA227ClearConditionalAllocationGuards(state arrayFlowState, proc sourceProcedure, text string, line int, variables map[string]arrayVariable) arrayFlowState {
	assigned, _, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return state
	}
	assignedName := strings.ToLower(cleanIdentifier(assigned))
	if assignedName == "" {
		return state
	}
	var updated arrayFlowState
	for name, value := range state {
		condition := value.conditionalAllocationSource
		if condition == "" || arrayVBA227ScalarConditionLHS(condition) != assignedName {
			continue
		}
		if arrayVBA227LineWithinMatchingGuard(proc, line, condition, variables) {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.conditionalAllocationSource = ""
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227LineWithinMatchingGuard(proc sourceProcedure, line int, condition string, variables map[string]arrayVariable) bool {
	statement := procedureStatementAtLine(proc, line)
	visited := map[int]bool{}
	for statement.ParentID != 0 && !visited[statement.ParentID] {
		visited[statement.ParentID] = true
		parent := procedureStatementByID(proc, statement.ParentID)
		if parent.ID == 0 {
			break
		}
		if parent.Kind == procedureir.StatementIf || parent.Kind == procedureir.StatementElseIf {
			if parentCondition, ok := arrayVBA227ScalarConditionSource(parent, variables); ok && parentCondition == condition {
				return true
			}
		}
		statement = parent
	}
	return false
}

func arrayVBA227ScalarConditionLHS(condition string) string {
	if match := arrayScalarConditionRe.FindStringSubmatch(condition); len(match) == 4 {
		return strings.ToLower(cleanIdentifier(match[1]))
	}
	if match := arrayScalarConditionReversedRe.FindStringSubmatch(condition); len(match) == 4 {
		return strings.ToLower(cleanIdentifier(match[3]))
	}
	return ""
}

func procedureStatementAtLine(proc sourceProcedure, line int) procedureir.Statement {
	var best procedureir.Statement
	for statement := range proc.Statements.All() {
		if line < statement.Range.StartLine || line > statement.Range.EndLine {
			continue
		}
		if best.ID == 0 || statement.Range.StartLine > best.Range.StartLine || statement.Range.EndLine < best.Range.EndLine {
			best = statement
		}
	}
	return best
}

func arrayVBA227ScalarConditionSource(statement procedureir.Statement, variables map[string]arrayVariable) (string, bool) {
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	condition = strings.TrimSpace(condition)
	lower := strings.ToLower(condition)
	if strings.HasPrefix(lower, "if ") {
		condition = strings.TrimSpace(condition[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		condition = strings.TrimSpace(condition[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(condition), " then"); then >= 0 && strings.TrimSpace(condition[then+5:]) == "" {
		condition = strings.TrimSpace(condition[:then])
	}
	for len(condition) >= 2 && condition[0] == '(' && condition[len(condition)-1] == ')' {
		condition = strings.TrimSpace(condition[1 : len(condition)-1])
	}
	if arrayConditionAndRe.MatchString(condition) || arrayConditionOrRe.MatchString(condition) {
		return "", false
	}
	if match := arrayScalarConditionRe.FindStringSubmatch(condition); len(match) == 4 {
		return arrayVBA227NormalizeScalarCondition(match[1], match[2], match[3], variables)
	}
	if match := arrayScalarConditionReversedRe.FindStringSubmatch(condition); len(match) == 4 {
		return arrayVBA227NormalizeScalarCondition(match[3], match[2], match[1], variables)
	}
	return "", false
}

func arrayVBA227PositiveScalarConditionSource(statement procedureir.Statement, variables map[string]arrayVariable) (string, bool) {
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	lhs, operator, literal, ok := arrayCountComparison(condition)
	if !ok {
		return "", false
	}
	if _, positive := arrayVBA227PositiveLengthCondition(condition); !positive {
		return "", false
	}
	return arrayVBA227NormalizeScalarCondition(lhs, operator, literal, variables)
}

func arrayVBA227PositiveGuardConditionSource(statement procedureir.Statement, variables map[string]arrayVariable) (string, bool) {
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	lhs, operator, literal, ok := arrayCountComparison(condition)
	if !ok {
		return "", false
	}
	value, err := strconv.Atoi(literal)
	if err != nil {
		return "", false
	}
	switch operator {
	case ">":
		if value < 0 {
			return "", false
		}
	case ">=":
		if value < 1 {
			return "", false
		}
	case "=":
		if value <= 0 {
			return "", false
		}
	default:
		return "", false
	}
	if _, ok := arrayVBA227NormalizeScalarCondition(lhs, operator, literal, variables); !ok {
		return "", false
	}
	name := strings.ToLower(cleanIdentifier(lhs))
	return name, name != ""
}

func arrayVBA227StatementStartsBefore(left, right procedureir.Statement) bool {
	if left.Range.StartByte != 0 && right.Range.StartByte != 0 {
		return left.Range.StartByte < right.Range.StartByte
	}
	if left.Range.StartLine != right.Range.StartLine {
		return left.Range.StartLine < right.Range.StartLine
	}
	if left.Range.StartColumn != right.Range.StartColumn {
		return left.Range.StartColumn < right.Range.StartColumn
	}
	return left.ID < right.ID
}

// arrayVBA227PositiveSafeArrayLengthArray recovers a positive SafeArrayLen
// proof when the length assignment is split across an If/Else branch. The
// positive guard excludes zero-literal fallback assignments, while every
// other assignment to the scalar must still be a recognized array-length
// probe. This keeps the recovery local to the guarded access.
func arrayVBA227PositiveSafeArrayLengthArray(file parsedFile, proc sourceProcedure, access *procedureir.Statement, lengthName string, variables map[string]arrayVariable, ctx analysisContext) (string, bool) {
	if access == nil || lengthName == "" {
		return "", false
	}
	lengthName = strings.ToLower(cleanIdentifier(lengthName))
	lengthVariable, known := variables[lengthName]
	if !known || lengthVariable.isArray || lengthVariable.isVariant || lengthVariable.isObject {
		return "", false
	}
	lengthLine := access.Range.StartLine
	if lengthLine <= proc.StartLine {
		return "", false
	}
	arrayName := ""
	probeLine := 0
	probeOffset := -1
	probeStatement := procedureir.Statement{}
	probeNonDominating := false
	hasZeroFallback := false
	for line := proc.StartLine; line < lengthLine && line <= len(file.Lines); line++ {
		for _, span := range splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(span.text)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				continue
			}
			if candidate, ok := arrayVBA227SafeArrayLengthSource(rhs, ctx); ok {
				statement := procedureStatementAtLine(proc, line)
				if statement.ID == 0 {
					return "", false
				}
				if !arrayVBA227StatementLineDominates(proc, line, *access) {
					probeNonDominating = true
				}
				if probeLine != 0 && !strings.EqualFold(candidate, arrayName) {
					return "", false
				}
				arrayName = candidate
				probeLine = line
				probeOffset = span.start
				probeStatement = statement
				continue
			}
			if value, ok := integerLiteral(rhs); ok && value == 0 {
				hasZeroFallback = true
				continue
			}
			return "", false
		}
	}
	if probeLine == 0 || arrayName == "" {
		return "", false
	}
	if probeNonDominating && !hasZeroFallback {
		return "", false
	}
	arrayValue, known := variables[arrayName]
	if !known || !arrayValue.isArray || !isByteArrayVariable(arrayValue) {
		return "", false
	}
	sourceLine, sourceOffset, ok := arrayVBA227ZeroBasedArraySourceBeforeOffset(file, proc, arrayName, probeLine, probeOffset)
	if !ok || sourceLine == 0 || probeLine < sourceLine || probeLine == sourceLine && sourceOffset >= probeOffset || !arrayVBA227NoArrayMutationBetweenSourceAndProbe(file, proc, arrayName, sourceLine, sourceOffset, probeLine, probeOffset, ctx) {
		return "", false
	}
	if !arrayVBA227NoArrayMutationWithinLine(file, proc, arrayName, probeLine, probeOffset, len(arraySourceOrderStripComment(file.Lines[probeLine-1])), ctx) {
		return "", false
	}
	if probeLine+1 < access.Range.StartLine && !arrayVBA227NoArrayMutationBetweenWithCalls(file, proc, arrayName, probeLine+1, access.Range.StartLine, ctx) {
		return "", false
	}
	if !arrayVBA227NoScalarAssignmentAfterSafeArrayProbe(file, proc, lengthName, probeLine+1, access.Range.StartLine, probeStatement, ctx) {
		return "", false
	}
	return arrayName, true
}

func arrayVBA227BranchOwner(proc sourceProcedure, statement procedureir.Statement) int {
	current := statement
	visited := map[int]bool{}
	for current.ParentID != 0 && !visited[current.ParentID] {
		visited[current.ParentID] = true
		parent := procedureStatementByID(proc, current.ParentID)
		if parent.ID == 0 {
			return 0
		}
		if parent.Kind == procedureir.StatementIf || parent.Kind == procedureir.StatementElseIf || parent.Kind == procedureir.StatementElse {
			return parent.ID
		}
		current = parent
	}
	return 0
}

// arrayVBA227PositiveConditionalReDimArray proves the companion form where a
// positive scalar is assigned alongside a conditional non-empty ReDim, with a
// zero fallback on the other branch. A later positive-length guard selects the
// ReDim arm, so the guarded element access cannot observe an unallocated or
// empty array.
func arrayVBA227PositiveConditionalReDimArray(file parsedFile, proc sourceProcedure, access, guard *procedureir.Statement, lengthName string, variables map[string]arrayVariable, ctx analysisContext) (string, bool) {
	if access == nil || lengthName == "" {
		return "", false
	}
	lengthName = strings.ToLower(cleanIdentifier(lengthName))
	lengthVariable, known := variables[lengthName]
	if !known || lengthVariable.isArray || lengthVariable.isVariant || lengthVariable.isObject {
		return "", false
	}
	positiveLine := 0
	positiveBranch := 0
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine >= access.Range.StartLine || statement.Kind != procedureir.StatementAssignment {
			continue
		}
		lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(normalizedCodeLine(statement.Text)))
		if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
			continue
		}
		value, ok := integerLiteral(rhs)
		if !ok {
			return "", false
		}
		if value <= 0 {
			continue
		}
		branch := arrayVBA227BranchOwner(proc, statement)
		if branch == 0 || positiveLine != 0 && positiveBranch != branch {
			return "", false
		}
		positiveLine = statement.Range.StartLine
		positiveBranch = branch
	}
	if positiveLine == 0 {
		return "", false
	}
	arrayName := ""
	redimLine := 0
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine >= access.Range.StartLine || statement.Kind != procedureir.StatementReDim {
			continue
		}
		if arrayVBA227BranchOwner(proc, statement) != positiveBranch {
			continue
		}
		text := strings.TrimSpace(normalizedCodeLine(statement.Text))
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
			return "", false
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct || !arrayVBA227RedimStartsAtZero(redim.dimensions) {
				continue
			}
			name := strings.ToLower(cleanIdentifier(redim.name))
			variable, known := variables[name]
			if !known || !variable.isArray || !isByteArrayVariable(variable) {
				continue
			}
			if arrayName != "" && arrayName != name {
				return "", false
			}
			arrayName = name
			redimLine = statement.Range.StartLine
		}
	}
	if arrayName == "" || redimLine == 0 || redimLine >= positiveLine {
		return "", false
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < proc.StartLine || call.Range.StartLine >= access.Range.StartLine {
			continue
		}
		if arrayVBA227CallMayMutateNamedScalar(file, proc, call, lengthName, ctx) {
			return "", false
		}
	}
	positivePathDominates := arrayVBA227StatementLineDominates(proc, positiveLine, *access) && arrayVBA227StatementLineDominates(proc, redimLine, *access)
	if !positivePathDominates {
		if guard == nil {
			return "", false
		}
		guardLength, positiveGuard := arrayVBA227PositiveGuardConditionSource(*guard, variables)
		if !positiveGuard || guardLength != lengthName || !arrayVBA227LocalScalarStartsAtZero(proc, lengthName) {
			return "", false
		}
	}
	if !arrayVBA227NoArrayMutationBetweenWithCalls(file, proc, arrayName, redimLine+1, access.Range.StartLine, ctx) {
		return "", false
	}
	return arrayName, true
}

func arrayVBA227LocalScalarStartsAtZero(proc sourceProcedure, name string) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	for declaration := range proc.Declarations.All() {
		if declaration.Scope == procedureir.ScopeLocal && !declaration.IsStatic && !declaration.IsArray && strings.EqualFold(cleanIdentifier(declaration.Name), name) {
			return true
		}
	}
	return false
}

func arrayVBA227AllocationProbeLengthArray(file parsedFile, proc sourceProcedure, access procedureir.Statement, state arrayFlowState, lengthName string, variables map[string]arrayVariable, ctx analysisContext) (string, bool) {
	value, known := state[strings.ToLower(cleanIdentifier(lengthName))]
	if !known || value.allocationProbe == "" {
		return "", false
	}
	arrayName := strings.ToLower(cleanIdentifier(value.allocationProbe))
	variable, known := variables[arrayName]
	if !known || !variable.isArray {
		return "", false
	}
	arrayValue, tracked := state[arrayName]
	if !tracked || !arrayValue.knownArray || arrayValue.kind == arrayUnallocated {
		return "", false
	}
	probeLine := 0
	probeStatement := procedureir.Statement{}
	probeNonDominating := false
	hasZeroFallback := false
	for line := proc.StartLine; line < access.Range.StartLine && line <= len(file.Lines); line++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				continue
			}
			candidate, ok := arrayVBA227SafeArrayLengthSource(rhs, ctx)
			if !ok {
				if value, literal := integerLiteral(rhs); literal && value == 0 {
					hasZeroFallback = true
				}
				continue
			}
			if !strings.EqualFold(candidate, arrayName) {
				continue
			}
			statement := procedureStatementAtLine(proc, line)
			if statement.ID == 0 {
				return "", false
			}
			if !arrayVBA227StatementLineDominates(proc, line, access) {
				probeNonDominating = true
			}
			probeLine = line
			probeStatement = statement
		}
	}
	if probeLine != 0 {
		if probeNonDominating && !hasZeroFallback || !arrayVBA227NoScalarAssignmentAfterSafeArrayProbe(file, proc, lengthName, probeLine+1, access.Range.StartLine, probeStatement, ctx) {
			return "", false
		}
	}
	return arrayName, true
}

func arrayVBA227UBoundLengthArray(proc sourceProcedure, guard, access *procedureir.Statement, lengthName string, variables map[string]arrayVariable) (string, bool) {
	if guard == nil || access == nil || lengthName == "" {
		return "", false
	}
	lengthName = strings.ToLower(cleanIdentifier(lengthName))
	lengthVariable, knownLength := variables[lengthName]
	if !knownLength || lengthVariable.isArray || lengthVariable.isVariant || lengthVariable.isObject || lengthVariable.static {
		return "", false
	}
	localScalar := false
	for declaration := range proc.Declarations.All() {
		if strings.EqualFold(declaration.Name, lengthName) && declaration.Scope == procedureir.ScopeLocal {
			localScalar = true
			break
		}
	}
	if !localScalar {
		return "", false
	}
	var selected procedureir.Statement
	arrayName := ""
	for candidate := range proc.Statements.All() {
		if candidate.Kind != procedureir.StatementAssignment || candidate.Range.StartLine >= guard.Range.StartLine {
			continue
		}
		text := strings.TrimSpace(candidate.Text)
		if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
			text = strings.TrimSpace(text[:newline])
		}
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
			continue
		}
		candidateArray, ok := arrayVBA227UBoundLengthSource(rhs)
		variable, known := variables[candidateArray]
		if !ok || !known || !variable.isArray {
			continue
		}
		if selected.ID == 0 || arrayVBA227StatementStartsBefore(selected, candidate) {
			selected = candidate
			arrayName = candidateArray
		}
	}
	if selected.ID == 0 || arrayName == "" {
		return "", false
	}
	for statement := range proc.Statements.All() {
		if statement.ID == selected.ID || !arrayVBA227StatementStartsBefore(statement, *access) {
			continue
		}
		if lhs, _, indexed, assigned := arrayAssignment(strings.TrimSpace(statement.Text)); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), lengthName) {
			return "", false
		}
		if arrayVBA227StatementStartsBefore(selected, statement) && arrayVBA227StatementMayMutateArray(proc, statement, arrayName) {
			return "", false
		}
	}
	return arrayName, true
}

func arrayVBA227StatementMayMutateArray(proc sourceProcedure, statement procedureir.Statement, arrayName string) bool {
	if arrayVBA227MutatesArray(statement.Text, arrayName) {
		return true
	}
	mutates := false
	forEachArrayCallAtLine(proc, statement.Range.StartLine, func(call procedureir.CallSite) {
		if mutates {
			return
		}
		if call.StatementID != 0 && statement.ID != 0 && call.StatementID != statement.ID {
			return
		}
		if arrayCallPassesDirectArrayArgument(proc, call, arrayName) {
			mutates = true
		}
	})
	if mutates {
		return true
	}
	// Recovered or unresolved call statements may not have a CallSite with
	// usable argument expression IDs. A whole-array mention is therefore
	// treated conservatively as a possible ByRef mutation; indexed mentions
	// remain available to the access itself and do not match this fallback.
	return arrayVBA227StatementMentionsWholeArray(statement.Text, arrayName)
}

func arrayVBA227StatementMentionsWholeArray(text, arrayName string) bool {
	want := strings.ToLower(cleanIdentifier(arrayName))
	if want == "" {
		return false
	}
	text = maskStringLiterals(gui.StripComment(text))
	for index := 0; index < len(text); index++ {
		if !isIdentifierStart(text[index]) || index > 0 && isIdentifierPart(text[index-1]) {
			continue
		}
		start := index
		index++
		for index < len(text) && isIdentifierPart(text[index]) {
			index++
		}
		if !strings.EqualFold(text[start:index], want) || start > 0 && (text[start-1] == '.' || text[start-1] == '!') {
			continue
		}
		for index < len(text) && (text[index] == ' ' || text[index] == '\t') {
			index++
		}
		if index >= len(text) || text[index] != '(' {
			return true
		}
	}
	return false
}

func arrayVBA227NormalizeScalarCondition(lhs, operator, rhs string, variables map[string]arrayVariable) (string, bool) {
	lhs = strings.ToLower(cleanIdentifier(lhs))
	rhs = strings.ToLower(strings.TrimSpace(rhs))
	variable, known := variables[lhs]
	if !known || variable.isArray || variable.isVariant || variable.isObject {
		return "", false
	}
	if rhsVariable, known := variables[rhs]; known && (rhsVariable.isArray || rhsVariable.isVariant || rhsVariable.isObject) {
		return "", false
	}
	return lhs + operator + rhs, true
}

func applyArrayVBA227ConditionalReDimBranch(state arrayFlowState, proc sourceProcedure, statement *procedureir.Statement, edge vbacfg.Edge, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf {
		return state
	}
	condition, ok := arrayVBA227ScalarConditionSource(*statement, variables)
	if !ok {
		return state
	}
	var updated arrayFlowState
	for name, value := range state {
		if !arrayVBA227ConditionalReDimBranchMatches(proc, statement, edge, value.conditionalAllocationSource, condition) {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227ConditionalReDimBranchMatches(proc sourceProcedure, statement *procedureir.Statement, edge vbacfg.Edge, allocationCondition, branchCondition string) bool {
	if edge.Kind == vbacfg.EdgeBranchTrue && allocationCondition == branchCondition {
		return true
	}
	if edge.Kind != vbacfg.EdgeBranchFalse {
		return false
	}
	positiveName, positive := arrayVBA227PositiveLengthCondition(allocationCondition)
	zeroName, zero := arrayVBA227ZeroLengthCondition(branchCondition)
	return positive && zero && positiveName == zeroName && arrayVBA227HasBoundsLengthAssignment(proc, statement, positiveName)
}

func arrayVBA227PositiveLengthCondition(condition string) (string, bool) {
	lhs, operator, literal, ok := arrayCountComparison(condition)
	if !ok {
		return "", false
	}
	value, err := strconv.Atoi(literal)
	if err != nil {
		return "", false
	}
	switch operator {
	case ">":
		if value != 0 {
			return "", false
		}
	case ">=":
		if value != 1 {
			return "", false
		}
	default:
		return "", false
	}
	name := strings.ToLower(cleanIdentifier(lhs))
	return name, name != ""
}

func arrayVBA227ZeroLengthCondition(condition string) (string, bool) {
	lhs, operator, literal, ok := arrayCountComparison(condition)
	if !ok || operator != "=" || literal != "0" {
		return "", false
	}
	name := strings.ToLower(cleanIdentifier(lhs))
	return name, name != ""
}

func arrayVBA227HasBoundsLengthAssignment(proc sourceProcedure, target *procedureir.Statement, name string) bool {
	if target == nil || name == "" {
		return false
	}
	for candidate := range proc.Statements.All() {
		if candidate.Kind != procedureir.StatementAssignment || candidate.Range.StartLine >= target.Range.StartLine {
			continue
		}
		text := strings.TrimSpace(candidate.Text)
		if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
			text = strings.TrimSpace(text[:newline])
		}
		match := arrayBoundsProbeRe.FindStringSubmatch(text)
		if len(match) != 4 || !strings.EqualFold(match[1], name) || !strings.EqualFold(match[2], match[3]) {
			continue
		}
		return arrayVBA227StatementLineDominates(proc, candidate.Range.StartLine, *target)
	}
	return false
}

func arrayVBA227HasDictionaryBoundsExpression(text string, state arrayFlowState) bool {
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		if !strings.EqualFold(bound[1], "ubound") {
			continue
		}
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		value, known := state[name]
		if known {
			if _, dictionary := arrayDictionaryCountSource(value.allocationCountSource); dictionary {
				return true
			}
		}
	}
	return false
}

func arraySuccessfulBoundsState(state arrayFlowState, text string, variables map[string]arrayVariable, loopEndLine int) arrayFlowState {
	// An explicitly declared Variant element array is still a known array, so a
	// successful bounds query establishes its allocation. An untyped Variant is
	// not marked isArray and remains conservative in the normal transfer path.
	var updated arrayFlowState
	dictionarySources := map[string]bool{}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		if strings.EqualFold(bound[1], "ubound") {
			if value, known := state[name]; known {
				if source, ok := arrayDictionaryCountSource(value.allocationCountSource); ok {
					dictionarySources[source] = true
				}
			}
		}
		variable, known := variables[name]
		value, knownValue := state[name]
		if !known || !knownValue || !variable.isArray && (!variable.isVariant || !value.knownArray) {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value = arrayVBA227RecordBoundsProof(value, loopEndLine)
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeUnallocated = false
		updated[name] = value
	}
	for name, value := range state {
		source, ok := arrayDictionaryCountSource(value.allocationCountSource)
		if !ok || !dictionarySources[source] {
			continue
		}
		variable, known := variables[name]
		if !known || !variable.isArray && !variable.isVariant {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value = updated[name]
		value = arrayVBA227RecordBoundsProof(value, loopEndLine)
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		value.mayBeUnallocated = false
		value.allocationCountSource = ""
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227RetainBoundsFailureOnResume(state, before arrayFlowState, text string, variables map[string]arrayVariable, proc sourceProcedure) arrayFlowState {
	if !arrayVBA227ResumeFactsFor(proc).hasResumeTransfer {
		return state
	}
	var updated arrayFlowState
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		variable, variableKnown := variables[name]
		prior, priorKnown := before[name]
		if name == "" || !variableKnown || !variable.isArray && !variable.isVariant || !priorKnown || !arrayVBA227BoundsCanFail(prior) {
			continue
		}
		value, valueKnown := state[name]
		if !valueKnown || value.mayBeUnallocated {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.resumeBoundsFailurePossible = true
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func (a Analyzer) arrayVBA227AddResumeBoundIndexFindings(findings []Finding, file parsedFile, proc sourceProcedure, line int, text string, state arrayFlowState, variables map[string]arrayVariable, vba227Graph *vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges) []Finding {
	if line <= 0 || !a.Config.Analyze.DetectArrayLifecycleSafety {
		return findings
	}
	access := procedureStatementAtLine(proc, line)
	if access.ID == 0 {
		return findings
	}
	seen := make(map[string]bool)
	for _, finding := range findings {
		if finding.Code == "VBA227" {
			seen[finding.arrayOperationKey] = true
		}
	}
	for _, use := range arrayIndexedUsesForSource(text, variables) {
		if len(use.args) == 0 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(use.name))
		value, known := state[name]
		if !known || !value.resumeBoundsFailurePossible || seen[arrayIndexOperationKey(name, "unallocated")] || !arrayVBA227ResumeCanReachPriorBoundsBody(proc, vba227Graph, resumeNextEdges, access.Range.StartLine, name, access, variables) {
			continue
		}
		finding := a.simpleFinding(file, proc, line, "VBA227", "warning", name+" is indexed before its array allocation is guaranteed.", "An array access can fail after an earlier bounds probe raises an error and an error handler resumes into this statement.", "Allocate the array on every path before indexing it, or guard the access with a proven allocation check.")
		finding.arrayLifecycleFinding = true
		finding.arrayOperationKey = arrayIndexOperationKey(name, "unallocated")
		findings = append(findings, finding)
		seen[finding.arrayOperationKey] = true
	}
	return findings
}

func arrayVBA227RecordBoundsProof(value arrayValue, loopEndLine int) arrayValue {
	if loopEndLine == 0 || value.boundsProof.loopEndLine != 0 || value.kind == arrayAllocated && value.knownArray && !value.mayBeUnallocated {
		return value
	}
	value.boundsProof = arrayBoundsProof{
		loopEndLine:                      loopEndLine,
		priorKind:                        value.kind,
		priorKnownArray:                  value.knownArray,
		priorMayBeEmpty:                  value.mayBeEmpty,
		priorMayBeUnallocated:            value.mayBeUnallocated,
		priorAllocationCount:             value.allocationCountSource,
		priorConditionalAllocationSource: value.conditionalAllocationSource,
	}
	return value
}

func arrayVBA227ClearLoopBodyBounds(state arrayFlowState, line int) arrayFlowState {
	if line <= 0 {
		return state
	}
	var updated arrayFlowState
	for name, value := range state {
		if value.boundsProof.loopEndLine != line {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.kind = value.boundsProof.priorKind
		value.knownArray = value.boundsProof.priorKnownArray
		value.mayBeEmpty = value.boundsProof.priorMayBeEmpty
		value.mayBeUnallocated = value.boundsProof.priorMayBeUnallocated
		value.allocationCountSource = value.boundsProof.priorAllocationCount
		value.conditionalAllocationSource = value.boundsProof.priorConditionalAllocationSource
		value.boundsProof = arrayBoundsProof{}
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayValueKnownLowerBound(value arrayValue) (int, bool) {
	for _, dimensions := range [][]arrayDimension{value.dimensions, value.preserveShape} {
		if len(dimensions) == 0 || !dimensions[0].lower.known {
			continue
		}
		return dimensions[0].lower.value, true
	}
	return 0, false
}

func arraySuccessfulConditionState(state arrayFlowState, statement *procedureir.Statement, variables map[string]arrayVariable, resumeNextBefore []bool, proc sourceProcedure) arrayFlowState {
	if statement == nil || (statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf) {
		return state
	}
	line := statement.Range.StartLine
	if arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
		return state
	}
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if strings.TrimSpace(condition) == "" {
		return state
	}
	updated := state
	cloned := false
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(condition, -1) {
		name := strings.ToLower(strings.TrimSpace(bound[2]))
		variable, known := variables[name]
		if !known || !variable.isArray {
			continue
		}
		value, known := updated[name]
		if !known {
			continue
		}
		if !cloned {
			updated = cloneArrayState(state)
			cloned = true
		}
		value = arrayVBA227RecordBoundsProof(value, arrayVBA227LoopBodyEndLine(proc, line))
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return arraySuccessfulBoundsState(updated, condition, variables, arrayVBA227LoopBodyEndLine(proc, line))
}

// arrayVBA227ResumeNextPrefixes computes the conservative "may have seen
// Resume Next" fact once per procedure. CFG joins retain the fact when any
// predecessor has enabled Resume Next, while explicit resets clear it only on
// paths that actually execute them. The source-order fallback is retained for
// focused or recovered procedures without a usable CFG.
func arrayVBA227ResumeNextPrefixes(file parsedFile, proc sourceProcedure) []bool {
	prefixes := make([]bool, len(file.Lines)+1)
	if proc.Graph != nil && len(proc.Graph.Blocks) > 0 {
		graph := proc.Graph.View(vbacfg.EdgeFilter{})
		resumeNextContinuations := arrayVBA227ResumeNextContinuations(graph)
		inStates := map[vbacfg.BlockID]bool{graph.Entry(): false}
		queued := map[vbacfg.BlockID]bool{graph.Entry(): true}
		for len(queued) > 0 {
			var id vbacfg.BlockID
			first := true
			for candidate := range queued {
				if first || candidate < id {
					id = candidate
					first = false
				}
			}
			delete(queued, id)
			active := inStates[id]
			block, ok := graph.BlockByID(id)
			if !ok {
				continue
			}
			out := active
			if block.Statement != nil {
				start := block.Statement.Range.StartLine
				if start == 0 {
					start = block.Range.StartLine
				}
				end := block.Statement.Range.EndLine
				if end < start {
					end = start
				}
				if start >= 1 && start <= len(file.Lines) {
					end = min(end, len(file.Lines))
					for line := start; line <= end; line++ {
						prefixes[line] = prefixes[line] || out
						out = arrayVBA227ResumeNextAfterStatement(out, normalizedCodeLine(file.Lines[line-1]))
					}
				} else {
					out = arrayVBA227ResumeNextAfterStatement(out, block.Statement.Text)
				}
			}
			graph.ForEachOutgoing(id, func(edge vbacfg.Edge) bool {
				next := out
				incoming, exists := inStates[edge.To]
				if !exists {
					inStates[edge.To] = next
					queued[edge.To] = true
					return true
				}
				if next && !incoming {
					inStates[edge.To] = true
					queued[edge.To] = true
				}
				return true
			})
			for _, target := range resumeNextContinuations[id] {
				if !inStates[target] {
					inStates[target] = true
					queued[target] = true
				}
			}
		}
		return prefixes
	}

	mayHaveResumeNext := false
	start := max(1, proc.StartLine)
	end := min(len(file.Lines), proc.EndLine)
	for line := start; line <= end; line++ {
		prefixes[line] = mayHaveResumeNext
		mayHaveResumeNext = arrayVBA227ResumeNextAfterStatement(mayHaveResumeNext, normalizedCodeLine(file.Lines[line-1]))
	}
	return prefixes
}

// arrayVBA227ResumeNextContinuations recovers the concrete continuation that
// VBA uses for Resume Next. The CFG models the instruction as an uncertain
// edge to UnknownExit because its target depends on the statement that raised
// the error. Error edges retain that missing association: when a Resume Next
// statement is reachable through a handler, its targets are the normal
// successors of the fault sites that enter that handler. Compound targets are
// precomputed from normal edges leaving each structured statement's region so
// nested constructs keep the same CFG join and do not fall through into a
// sibling branch.
func arrayVBA227ResumeNextContinuations(graph vbacfg.CFGView) map[vbacfg.BlockID][]vbacfg.BlockID {
	normalOutgoing := map[vbacfg.BlockID][]vbacfg.BlockID{}
	errorSources := map[vbacfg.BlockID]map[vbacfg.BlockID]bool{}
	blocksByID := map[vbacfg.BlockID]vbacfg.Block{}
	statementBlocks := map[int]vbacfg.BlockID{}
	statementKinds := map[int]procedureir.StatementKind{}
	parents := map[int]int{}
	children := map[int][]int{}
	graph.ForEachBlock(func(block vbacfg.Block) bool {
		blocksByID[block.ID] = block
		if block.Statement != nil {
			statementBlocks[block.Statement.ID] = block.ID
			statementKinds[block.Statement.ID] = block.Statement.Kind
			parents[block.Statement.ID] = block.Statement.ParentID
			children[block.Statement.ParentID] = append(children[block.Statement.ParentID], block.Statement.ID)
		}
		return true
	})
	for parent := range children {
		slices.Sort(children[parent])
	}

	compoundAncestors := map[vbacfg.BlockID][]vbacfg.BlockID{}
	compoundBlocks := map[vbacfg.BlockID]bool{}
	for blockID, block := range blocksByID {
		if block.Statement == nil {
			continue
		}
		for statementID := block.Statement.ID; statementID != 0; {
			ancestorID, ok := statementBlocks[statementID]
			if ok {
				ancestor := blocksByID[ancestorID]
				if ancestor.Statement != nil && arrayVBA227CompoundStatement(ancestor.Statement.Kind) {
					compoundAncestors[blockID] = append(compoundAncestors[blockID], ancestorID)
					compoundBlocks[ancestorID] = true
				}
			}
			next, ok := parents[statementID]
			if !ok {
				break
			}
			statementID = next
		}
	}

	compoundTargetSets := map[vbacfg.BlockID]map[vbacfg.BlockID]bool{}
	graph.ForEachEdge(func(edge vbacfg.Edge) bool {
		if edge.Class == vbacfg.EdgeNormal {
			normalOutgoing[edge.From] = append(normalOutgoing[edge.From], edge.To)
			if arrayVBA227NaturalContinuationEdge(edge.Kind) &&
				(edge.Kind != vbacfg.EdgeLoopExit || !arrayVBA227ExplicitLoopExit(blocksByID[edge.From])) {
				for _, compoundID := range compoundAncestors[edge.From] {
					compound := blocksByID[compoundID]
					target := blocksByID[edge.To]
					if target.Statement != nil && compound.Statement != nil &&
						arrayVBA227WithinStatement(target.Statement.ID, compound.Statement.ID, parents) {
						continue
					}
					if compoundTargetSets[compoundID] == nil {
						compoundTargetSets[compoundID] = map[vbacfg.BlockID]bool{}
					}
					compoundTargetSets[compoundID][edge.To] = true
				}
			}
		}
		if edge.Class == vbacfg.EdgeExceptional && edge.Kind == vbacfg.EdgeError {
			handler, ok := blocksByID[edge.To]
			if !ok || handler.Statement == nil || handler.Statement.Kind != procedureir.StatementLabel {
				return true
			}
			if errorSources[edge.To] == nil {
				errorSources[edge.To] = map[vbacfg.BlockID]bool{}
			}
			errorSources[edge.To][edge.From] = true
		}
		return true
	})
	if len(errorSources) == 0 {
		return nil
	}

	resumeNextBlocks := map[vbacfg.BlockID]bool{}
	graph.ForEachBlock(func(block vbacfg.Block) bool {
		if block.Statement != nil && block.Statement.Control != nil &&
			block.Statement.Control.Transfer == procedureir.TransferResumeNext {
			resumeNextBlocks[block.ID] = true
		}
		return true
	})
	if len(resumeNextBlocks) == 0 {
		return nil
	}

	handlerLabels := map[int]bool{}
	for handler := range errorSources {
		if block := blocksByID[handler]; block.Statement != nil {
			handlerLabels[block.Statement.ID] = true
		}
	}
	compoundTargets := make(map[vbacfg.BlockID][]vbacfg.BlockID, len(compoundBlocks))
	for compoundID := range compoundBlocks {
		targets := compoundTargetSets[compoundID]
		if len(targets) == 0 {
			compound := blocksByID[compoundID]
			var fallback vbacfg.BlockID
			if compound.Statement != nil {
				fallback = arrayVBA227SyntaxContinuation(compound.Statement.ID, graph.NormalExit(), statementBlocks, statementKinds, parents, children, handlerLabels)
			}
			if fallback != 0 {
				compoundTargets[compoundID] = []vbacfg.BlockID{fallback}
			} else if exit := graph.NormalExit(); exit != 0 {
				compoundTargets[compoundID] = []vbacfg.BlockID{exit}
			}
			continue
		}
		compoundTargets[compoundID] = make([]vbacfg.BlockID, 0, len(targets))
		for target := range targets {
			compoundTargets[compoundID] = append(compoundTargets[compoundID], target)
		}
		slices.Sort(compoundTargets[compoundID])
	}

	continuationSets := map[vbacfg.BlockID]map[vbacfg.BlockID]bool{}
	for handler, sources := range errorSources {
		reachable := map[vbacfg.BlockID]bool{handler: true}
		queue := []vbacfg.BlockID{handler}
		for len(queue) > 0 {
			current := queue[0]
			queue = queue[1:]
			for _, target := range normalOutgoing[current] {
				if reachable[target] {
					continue
				}
				reachable[target] = true
				queue = append(queue, target)
			}
		}
		for resumeBlock := range resumeNextBlocks {
			if !reachable[resumeBlock] {
				continue
			}
			if continuationSets[resumeBlock] == nil {
				continuationSets[resumeBlock] = map[vbacfg.BlockID]bool{}
			}
			for source := range sources {
				targets := normalOutgoing[source]
				if compound, ok := compoundTargets[source]; ok {
					targets = compound
				}
				for _, target := range targets {
					continuationSets[resumeBlock][target] = true
				}
			}
		}
	}
	if len(continuationSets) == 0 {
		return nil
	}

	continuations := make(map[vbacfg.BlockID][]vbacfg.BlockID, len(continuationSets))
	for resumeBlock, targets := range continuationSets {
		continuations[resumeBlock] = make([]vbacfg.BlockID, 0, len(targets))
		for target := range targets {
			continuations[resumeBlock] = append(continuations[resumeBlock], target)
		}
		slices.Sort(continuations[resumeBlock])
	}
	return continuations
}

func arrayVBA227NaturalContinuationEdge(kind vbacfg.EdgeKind) bool {
	switch kind {
	case vbacfg.EdgeGoto, vbacfg.EdgeProcedureExit, vbacfg.EdgeTermination,
		vbacfg.EdgeResume, vbacfg.EdgeError, vbacfg.EdgeUnknown:
		return false
	default:
		return true
	}
}

func arrayVBA227ExplicitLoopExit(block vbacfg.Block) bool {
	if block.Statement == nil || block.Statement.Control == nil {
		return false
	}
	switch block.Statement.Control.Transfer {
	case procedureir.TransferExitFor, procedureir.TransferExitDo:
		return true
	default:
		return false
	}
}

func arrayVBA227CompoundStatement(kind procedureir.StatementKind) bool {
	switch kind {
	case procedureir.StatementIf, procedureir.StatementElseIf, procedureir.StatementSelect,
		procedureir.StatementCase, procedureir.StatementFor, procedureir.StatementForEach,
		procedureir.StatementWhile, procedureir.StatementDo, procedureir.StatementWith:
		return true
	default:
		return false
	}
}

func arrayVBA227WithinStatement(statementID, ancestorID int, parents map[int]int) bool {
	if statementID == ancestorID {
		return true
	}
	seen := map[int]bool{}
	for current := parents[statementID]; current != 0 && !seen[current]; current = parents[current] {
		if current == ancestorID {
			return true
		}
		seen[current] = true
	}
	return false
}

func arrayVBA227SyntaxContinuation(statementID int, normalExit vbacfg.BlockID, statementBlocks map[int]vbacfg.BlockID, statementKinds map[int]procedureir.StatementKind, parents map[int]int, children map[int][]int, handlerLabels map[int]bool) vbacfg.BlockID {
	seen := map[int]bool{}
	current := statementID
	for current != 0 && !seen[current] {
		seen[current] = true
		parent := parents[current]
		for _, candidate := range children[parent] {
			if candidate <= current || arrayVBA227AlternativeChild(statementKinds[parent], statementKinds[candidate]) {
				continue
			}
			if handlerLabels[candidate] {
				break
			}
			if statementKinds[candidate] == procedureir.StatementLabel {
				continue
			}
			if block, ok := statementBlocks[candidate]; ok {
				return block
			}
		}
		if parent == 0 {
			break
		}
		if arrayVBA227LoopStatement(statementKinds[parent]) {
			if block, ok := statementBlocks[parent]; ok {
				return block
			}
		}
		current = parent
	}
	return normalExit
}

func arrayVBA227AlternativeChild(parent, child procedureir.StatementKind) bool {
	if parent == procedureir.StatementIf || parent == procedureir.StatementElseIf {
		return child == procedureir.StatementElse || child == procedureir.StatementElseIf
	}
	return parent == procedureir.StatementSelect && child == procedureir.StatementCase
}

func arrayVBA227LoopStatement(kind procedureir.StatementKind) bool {
	switch kind {
	case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile:
		return true
	default:
		return false
	}
}

func arrayVBA227ResumeNextAfterStatement(active bool, text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.Contains(lower, " else ") && strings.Contains(lower, "on error resume next") && strings.Contains(lower, "on error goto") {
		return true
	}
	next := active
	for _, statement := range splitRangeValueSourceStatements(text) {
		trimmed := strings.TrimSpace(statement)
		switch {
		case arrayOnErrorGotoZeroRe.MatchString(trimmed), arrayOnErrorGotoRe.MatchString(trimmed):
			next = false
		case arrayOnErrorResumeNextRe.MatchString(trimmed), arrayOnErrorResumeNextStatementRe.MatchString(trimmed):
			next = true
		}
	}
	return next
}

func arrayVBA227ResumeCanReachForBody(proc sourceProcedure, vba227Graph *vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges, loop, access procedureir.Statement) bool {
	return arrayVBA227ResumeCanReachIndexedConditionBody(proc, vba227Graph, resumeNextEdges, loop, access)
}

func arrayVBA227ResumeCanReachPriorBoundsBody(proc sourceProcedure, vba227Graph *vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges, beforeLine int, name string, access procedureir.Statement, variables map[string]arrayVariable) bool {
	variable, known := variables[name]
	if !known || variable.fixed {
		return false
	}
	facts := arrayVBA227ResumeFactsFor(proc)
	for _, statement := range facts.boundsByName[strings.ToLower(cleanIdentifier(name))] {
		if statement.Range.StartLine >= beforeLine {
			continue
		}
		if arrayVBA227ResumeCanReachIndexedConditionBody(proc, vba227Graph, resumeNextEdges, statement, access) {
			return true
		}
	}
	return false
}

func arrayVBA227BoundsCanFail(value arrayValue) bool {
	return value.kind != arrayAllocated || !value.knownArray || value.mayBeUnallocated
}

func arrayVBA227ResumeNextBeforeLine(prefixes []bool, line int) bool {
	return line >= 0 && line < len(prefixes) && prefixes[line]
}

func arrayVBA227ResumeCanReachIndexedConditionBody(proc sourceProcedure, vba227Graph *vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges, guard, access procedureir.Statement) bool {
	if guard.ID == 0 || access.ID == 0 {
		return false
	}
	if vba227Graph == nil && proc.Graph == nil {
		for statement := range proc.Statements.All() {
			if statement.Kind != procedureir.StatementResume || statement.Control == nil {
				continue
			}
			switch statement.Control.Transfer {
			case procedureir.TransferResumeNext, procedureir.TransferResumeLabel:
				return true
			}
		}
		return false
	}
	var graph vbacfg.CFGView
	if vba227Graph != nil {
		graph = *vba227Graph
	} else {
		graph = proc.Graph.View(vbacfg.EdgeFilter{})
	}
	guardBlock, ok := graph.BlockForStatement(guard.ID)
	if !ok {
		return false
	}
	bodyBlock, ok := graph.BlockForStatement(access.ID)
	if !ok {
		return false
	}
	handlers := map[vbacfg.BlockID]bool{}
	graph.ForEachOutgoing(guardBlock.ID, func(edge vbacfg.Edge) bool {
		if edge.Class == vbacfg.EdgeExceptional && edge.Kind == vbacfg.EdgeError {
			handlers[edge.To] = true
		}
		return true
	})
	for handler := range handlers {
		reachable := arrayVBA227NormalReachableBlocks(graph, handler)
		for blockID := range reachable {
			block, exists := graph.BlockByID(blockID)
			if !exists || block.Statement == nil || block.Statement.Control == nil {
				continue
			}
			switch block.Statement.Control.Transfer {
			case procedureir.TransferResumeNext:
				if arrayVBA227ResumeNextContinuationReachesBody(proc, graph, resumeNextEdges, block.ID, bodyBlock.ID, guardBlock.ID) ||
					arrayVBA227ResumeNextFaultPathReachesBody(graph, resumeNextEdges, guardBlock.ID, bodyBlock.ID) {
					return true
				}
			case procedureir.TransferResumeLabel:
				labelReachesBody := false
				graph.ForEachOutgoing(block.ID, func(edge vbacfg.Edge) bool {
					if edge.Class == vbacfg.EdgeExceptional && edge.Kind == vbacfg.EdgeResume && (arrayVBA227NormalPathReachesWithout(graph, edge.To, bodyBlock.ID, guardBlock.ID) || arrayVBA227ResumeNextPathReachesWithout(graph, edge.To, bodyBlock.ID, guardBlock.ID, resumeNextEdges) || arrayVBA227UnknownFlowCanReachBody(graph, edge.To, guardBlock.ID, proc.Graph.UnknownFlowSources, resumeNextEdges)) {
						labelReachesBody = true
					}
					return true
				})
				if labelReachesBody {
					return true
				}
			}
		}
	}
	return false
}

func arrayVBA227ResumeNextContinuationReachesBody(proc sourceProcedure, graph vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges, resumeBlock, bodyBlock, blocked vbacfg.BlockID) bool {
	for _, target := range arrayVBA227ResumeFactsFor(proc).resumeNextContinuations[resumeBlock] {
		if arrayVBA227ResumeNextPathReachesWithout(graph, target, bodyBlock, blocked, resumeNextEdges) {
			return true
		}
	}
	return false
}

func arrayVBA227ResumeNextFaultPathReachesBody(graph vbacfg.CFGView, resumeNextEdges arrayVBA227ResumeNextEdges, faultBlock, bodyBlock vbacfg.BlockID) bool {
	if faultBlock == bodyBlock {
		return true
	}
	reaches := false
	graph.ForEachOutgoing(faultBlock, func(edge vbacfg.Edge) bool {
		if (edge.Class == vbacfg.EdgeNormal || resumeNextEdges[faultBlock][edge.To]) && arrayVBA227NormalPathReachesWithout(graph, edge.To, bodyBlock, faultBlock) {
			reaches = true
		}
		return !reaches
	})
	return reaches
}

func arrayVBA227NormalReachableBlocks(graph vbacfg.CFGView, start vbacfg.BlockID) map[vbacfg.BlockID]bool {
	return arrayVBA227NormalReachableBlocksWithout(graph, start, 0)
}

func arrayVBA227NormalPathReachesWithout(graph vbacfg.CFGView, start, target, blocked vbacfg.BlockID) bool {
	if start == blocked {
		return false
	}
	if start == target {
		return true
	}
	reachable := arrayVBA227NormalReachableBlocksWithout(graph, start, blocked)
	return reachable[target]
}

func arrayVBA227ResumeNextPathReachesWithout(graph vbacfg.CFGView, start, target, blocked vbacfg.BlockID, resumeNextEdges arrayVBA227ResumeNextEdges) bool {
	if start == blocked {
		return false
	}
	if start == target {
		return true
	}
	reachable := arrayVBA227ResumeNextReachableBlocksWithout(graph, start, blocked, resumeNextEdges)
	return reachable[target]
}

func arrayVBA227UnknownFlowCanReachBody(graph vbacfg.CFGView, start, blocked vbacfg.BlockID, unknownFlowSources []vbacfg.BlockID, resumeNextEdges arrayVBA227ResumeNextEdges) bool {
	unknown := make(map[vbacfg.BlockID]struct{}, len(unknownFlowSources))
	for _, unknownFlowSource := range unknownFlowSources {
		unknown[unknownFlowSource] = struct{}{}
	}
	for blockID := range arrayVBA227ResumeNextReachableBlocksWithout(graph, start, blocked, resumeNextEdges) {
		if _, ok := unknown[blockID]; ok {
			return true
		}
	}
	return false
}

func arrayVBA227NormalReachableBlocksWithout(graph vbacfg.CFGView, start, blocked vbacfg.BlockID) map[vbacfg.BlockID]bool {
	return arrayVBA227ReachableBlocksWithout(graph, start, blocked, nil)
}

func arrayVBA227ResumeNextReachableBlocksWithout(graph vbacfg.CFGView, start, blocked vbacfg.BlockID, resumeNextEdges arrayVBA227ResumeNextEdges) map[vbacfg.BlockID]bool {
	return arrayVBA227ReachableBlocksWithout(graph, start, blocked, resumeNextEdges)
}

func arrayVBA227ReachableBlocksWithout(graph vbacfg.CFGView, start, blocked vbacfg.BlockID, resumeNextEdges arrayVBA227ResumeNextEdges) map[vbacfg.BlockID]bool {
	if start == blocked {
		return nil
	}
	reachable := map[vbacfg.BlockID]bool{start: true}
	queue := []vbacfg.BlockID{start}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if (edge.Class == vbacfg.EdgeNormal || resumeNextEdges[current][edge.To]) && edge.To != blocked && !reachable[edge.To] {
				reachable[edge.To] = true
				queue = append(queue, edge.To)
			}
			return true
		})
	}
	return reachable
}

// arrayVBA227ResumeNextContinuationEdges identifies exceptional edges that
// mirror a normal successor in the unfiltered CFG. The builder emits those
// pairs for On Error Resume Next, so the exceptional edge remains the valid
// continuation after arrayVBA227Graph removes an Err.Raise normal edge.
func arrayVBA227ResumeNextContinuationEdges(proc sourceProcedure) arrayVBA227ResumeNextEdges {
	if proc.Graph == nil {
		return nil
	}
	normal := map[vbacfg.BlockID]map[vbacfg.BlockID]bool{}
	continuations := arrayVBA227ResumeNextEdges{}
	graph := proc.Graph.View(vbacfg.EdgeFilter{})
	graph.ForEachEdge(func(edge vbacfg.Edge) bool {
		if edge.Class == vbacfg.EdgeNormal {
			if normal[edge.From] == nil {
				normal[edge.From] = map[vbacfg.BlockID]bool{}
			}
			normal[edge.From][edge.To] = true
		}
		return true
	})
	graph.ForEachEdge(func(edge vbacfg.Edge) bool {
		if edge.Class == vbacfg.EdgeExceptional && edge.Kind == vbacfg.EdgeError && normal[edge.From][edge.To] {
			if continuations[edge.From] == nil {
				continuations[edge.From] = map[vbacfg.BlockID]bool{}
			}
			continuations[edge.From][edge.To] = true
		}
		return true
	})
	return continuations
}

func arrayIfThenParts(text string) (condition, body string, ok bool) {
	text = strings.TrimSpace(text)
	lower := strings.ToLower(text)
	prefixLength := 0
	switch {
	case strings.HasPrefix(lower, "if "):
		prefixLength = len("if ")
	case strings.HasPrefix(lower, "elseif "):
		prefixLength = len("elseif ")
	default:
		return "", "", false
	}
	rest := strings.TrimSpace(text[prefixLength:])
	then := arrayTopLevelKeywordIndex(rest, "then")
	if then < 0 {
		return "", "", false
	}
	return strings.TrimSpace(text[:prefixLength] + rest[:then]), strings.TrimSpace(rest[then+len("then"):]), true
}

func arrayIfThenBodyParts(body string) (thenBody, elseBody string, hasElse bool) {
	elseIndex := arrayTopLevelKeywordIndex(body, "else")
	if elseIndex < 0 {
		return strings.TrimSpace(body), "", false
	}
	return strings.TrimSpace(body[:elseIndex]), strings.TrimSpace(body[elseIndex+len("else"):]), true
}

func arrayTopLevelKeywordIndex(text, keyword string) int {
	if keyword == "" {
		return -1
	}
	depth := 0
	inString := false
	for i := 0; i <= len(text)-len(keyword); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			continue
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		}
		if inString || depth != 0 || !strings.EqualFold(text[i:i+len(keyword)], keyword) {
			continue
		}
		if i > 0 && isIdentifierPart(text[i-1]) {
			continue
		}
		end := i + len(keyword)
		if end < len(text) && isIdentifierPart(text[end]) {
			continue
		}
		return i
	}
	return -1
}

func arrayVBA227HasArrayFactoryAssignment(text string) bool {
	_, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(rhs))
	return strings.HasPrefix(lower, "array(") || strings.HasPrefix(lower, "split(") || strings.HasPrefix(lower, "filter(")
}

// arrayVBA227RepeatedSelectCaseBoundsState carries a successful bounds query
// from a Case Else branch into a later Case Else branch that uses the same
// ByVal scalar selector. The CFG meets the other cases at End Select and
// therefore loses this correlation even though an unchanged selector makes
// the two Case Else paths equivalent.
func arrayVBA227RepeatedSelectCaseBoundsState(file parsedFile, proc sourceProcedure, line int, state arrayFlowState, variables map[string]arrayVariable) arrayFlowState {
	if state == nil || proc.Graph == nil {
		return state
	}
	current := procedureStatementAtLine(proc, line)
	currentCase, currentSelect := arrayVBA227EnclosingSelectCase(proc, current)
	if currentCase.ID == 0 || currentCase.Control == nil || !currentCase.Control.CaseElse || currentSelect.ID == 0 {
		return state
	}
	selectExpression := strings.TrimSpace(selectCaseExpression(currentSelect.Text))
	selector := strings.ToLower(cleanIdentifier(selectExpression))
	if selector == "" || !arrayVBA227StableSelectCaseSelector(proc, selectExpression) {
		return state
	}
	previousSelect, previousCase, ok := arrayVBA227PreviousSelectCaseElse(proc, currentSelect, selectExpression)
	if !ok {
		return state
	}
	proven := arrayVBA227TrailingSelectCaseBounds(file, proc, previousCase, variables)
	if len(proven) == 0 || arrayVBA227SelectCaseHasAssignment(file, previousCase.Range.StartLine+1, previousCase.Range.EndLine-1, selector) {
		return state
	}
	if !arrayVBA227SelectCaseRegionStable(file, proc, previousSelect.Range.EndLine+1, currentSelect.Range.StartLine-1, selector, proven) ||
		!arrayVBA227SelectCaseRegionStable(file, proc, currentCase.Range.StartLine+1, line-1, selector, proven) {
		return state
	}
	updated := state
	cloned := false
	for name := range proven {
		value, known := updated[name]
		variable, variableKnown := variables[name]
		if !known || !variableKnown || !variable.isArray {
			continue
		}
		if !cloned {
			updated = cloneArrayState(state)
			cloned = true
		}
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayVBA227EnclosingSelectCase(proc sourceProcedure, statement procedureir.Statement) (procedureir.Statement, procedureir.Statement) {
	for statement.ID != 0 {
		if statement.Kind == procedureir.StatementCase {
			parent := procedureStatementByID(proc, statement.ParentID)
			if parent.Kind == procedureir.StatementSelect {
				return statement, parent
			}
		}
		statement = procedureStatementByID(proc, statement.ParentID)
	}
	return procedureir.Statement{}, procedureir.Statement{}
}

func arrayVBA227PreviousSelectCaseElse(proc sourceProcedure, current procedureir.Statement, expression string) (procedureir.Statement, procedureir.Statement, bool) {
	want := canonicalArrayBoundExpression(expression)
	var previousSelect procedureir.Statement
	var previousCase procedureir.Statement
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSelect || statement.Range.EndLine >= current.Range.StartLine ||
			canonicalArrayBoundExpression(selectCaseExpression(statement.Text)) != want {
			continue
		}
		var candidate procedureir.Statement
		for child := range proc.Statements.All() {
			if child.ParentID == statement.ID && child.Kind == procedureir.StatementCase && child.Control != nil && child.Control.CaseElse {
				candidate = child
				break
			}
		}
		if candidate.ID == 0 || previousSelect.ID != 0 && statement.Range.EndLine <= previousSelect.Range.EndLine {
			continue
		}
		previousSelect = statement
		previousCase = candidate
	}
	return previousSelect, previousCase, previousSelect.ID != 0 && previousCase.ID != 0
}

func arrayVBA227StableSelectCaseSelector(proc sourceProcedure, expression string) bool {
	trimmed := strings.TrimSpace(expression)
	if trimmed == "" || !isIdentifierStart(trimmed[0]) {
		return false
	}
	for index := 1; index < len(trimmed); index++ {
		if !isIdentifierPart(trimmed[index]) {
			return false
		}
	}
	for parameter := range proc.Params.All() {
		if !strings.EqualFold(cleanIdentifier(parameter.Name), trimmed) || parameterIsArray(parameter) {
			continue
		}
		return strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
	}
	return false
}

func arrayVBA227TrailingSelectCaseBounds(file parsedFile, proc sourceProcedure, branch procedureir.Statement, variables map[string]arrayVariable) map[string]bool {
	proven := map[string]bool{}
	lastBoundLine := map[string]int{}
	lastExecutableLine := 0
	start := max(proc.StartLine, branch.Range.StartLine+1)
	end := min(proc.EndLine, min(branch.Range.EndLine, len(file.Lines)))
	for line := start; line <= end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		lower := strings.ToLower(text)
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") || lower == "end select" {
			continue
		}
		if !arrayVBA227SelectCaseStraightLine(text) {
			return nil
		}
		lastExecutableLine = line
		for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
			name := strings.ToLower(cleanIdentifier(bound[2]))
			variable, known := variables[name]
			if known && variable.isArray {
				proven[name] = true
				lastBoundLine[name] = line
			}
		}
	}
	if lastExecutableLine == 0 || len(proven) == 0 {
		return nil
	}
	for name := range proven {
		if lastBoundLine[name] != lastExecutableLine {
			return nil
		}
	}
	return proven
}

func arrayVBA227SelectCaseHasAssignment(file parsedFile, start, end int, target string) bool {
	if start > end {
		return false
	}
	start = max(1, start)
	end = min(end, len(file.Lines))
	for line := start; line <= end; line++ {
		text := normalizedCodeLine(file.Lines[line-1])
		lhs, _, indexed, assigned := arrayAssignment(text)
		if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), target) {
			return true
		}
	}
	return false
}

func arrayVBA227SelectCaseRegionStable(file parsedFile, proc sourceProcedure, start, end int, selector string, arrays map[string]bool) bool {
	if start > end {
		return true
	}
	start = max(1, start)
	end = min(end, len(file.Lines))
	for line := start; line <= end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") {
			continue
		}
		if !arrayVBA227SelectCaseStraightLine(text) || arrayHasCallsAtLine(proc, line) {
			return false
		}
		lhs, _, indexed, assigned := arrayAssignment(text)
		if assigned && !indexed {
			name := strings.ToLower(cleanIdentifier(lhs))
			if name == selector || arrays[name] {
				return false
			}
		}
		if arrayRedimRe.MatchString(text) || arrayEraseRe.MatchString(text) {
			return false
		}
	}
	return true
}

func arrayVBA227SelectCaseStraightLine(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{
		"if ", "elseif ", "else", "end if", "for ", "for each ", "next", "do", "loop", "while ", "wend",
		"select ", "case ", "goto ", "exit ", "on error ", "resume ", "with ", "end with",
	} {
		if strings.HasPrefix(lower, prefix) {
			return false
		}
	}
	return !strings.HasSuffix(lower, ":")
}
