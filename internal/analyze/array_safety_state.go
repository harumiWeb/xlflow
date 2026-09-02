package analyze

import (
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func arrayAllocationTransferIsReliable(statement *procedureir.Statement, in, out arrayFlowState) bool {
	if statement == nil || !arrayStateAddsAllocation(in, out) {
		return false
	}
	text := strings.TrimSpace(statement.Text)
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		// ReDim Preserve can fail when the prior allocation or shape is
		// unknown. A plain ReDim is the deterministic allocation boundary.
		return strings.TrimSpace(match[1]) == ""
	}
	_, rhs, indexed, ok := arrayAssignment(text)
	if !ok || indexed {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(rhs))
	return strings.HasPrefix(lower, "array(") || arrayCallName(rhs) == "split" || arrayCallName(rhs) == "filter"
}

func arrayStateAddsAllocation(in, out arrayFlowState) bool {
	for name, value := range out {
		if value.kind != arrayAllocated || !value.knownArray {
			continue
		}
		before, ok := in[name]
		if !ok || before.kind != arrayAllocated || !before.knownArray {
			return true
		}
	}
	return false
}

func arrayCountExpressionMatches(expression, source string) bool {
	if strings.TrimSpace(expression) == "" || strings.TrimSpace(source) == "" {
		return false
	}
	expression = canonicalArrayBoundExpression(expression)
	source = canonicalArrayBoundExpression(arrayCountSourceExpression(source))
	return expression == source || expression == source+".count"
}

func arrayCountSourceExpression(source string) string {
	source = strings.TrimSpace(source)
	if strings.HasPrefix(strings.ToLower(source), arrayDictionaryCountSourcePrefix) {
		return strings.TrimSpace(source[len(arrayDictionaryCountSourcePrefix):])
	}
	return source
}

func arrayDictionaryCountSource(source string) (string, bool) {
	source = strings.TrimSpace(source)
	if !strings.HasPrefix(strings.ToLower(source), arrayDictionaryCountSourcePrefix) {
		return "", false
	}
	expression := strings.TrimSpace(source[len(arrayDictionaryCountSourcePrefix):])
	if expression == "" {
		return "", false
	}
	return canonicalArrayBoundExpression(expression), true
}

func applyArrayConditionalAllocationBranch(state arrayFlowState, graph *vbacfg.CFGView, block vbacfg.Block, edge vbacfg.Edge) arrayFlowState {
	if block.Statement == nil {
		return state
	}
	if block.Statement.Kind == procedureir.StatementSelect && edge.Kind == vbacfg.EdgeCase && graph != nil {
		selectExpression := selectCaseExpression(block.Statement.Text)
		caseBlock, ok := graph.BlockByID(edge.To)
		if !ok {
			return state
		}
		caseOK := positiveSelectCaseValue(caseBlock.Statement)
		if !caseOK || strings.TrimSpace(selectExpression) == "" {
			return state
		}
		updated := state
		cloned := false
		for name, value := range state {
			if value.allocationCountSource == "" || !arrayCountExpressionMatches(selectExpression, value.allocationCountSource) {
				continue
			}
			if !cloned {
				updated = cloneArrayState(state)
				cloned = true
			}
			value.kind = arrayAllocated
			value.knownArray = true
			// Retain the count witness on the refined branch. A later join with
			// the sibling path can then preserve the conditional fact without
			// mutating the sibling's input state.
			updated[name] = value
		}
		return updated
	}
	if block.Statement.Kind != procedureir.StatementIf && block.Statement.Condition == nil {
		return state
	}
	condition := ""
	if block.Statement.Condition != nil {
		condition = block.Statement.Condition.Text
	} else {
		condition = block.Statement.Text
	}
	comparisons := []string{condition}
	if arrayConditionAndRe.MatchString(condition) {
		comparisons = arrayConditionAndRe.Split(condition, -1)
	}
	updated := state
	cloned := false
	for name, value := range state {
		for _, comparison := range comparisons {
			lhs, operator, literal, ok := arrayCountComparison(comparison)
			if !ok {
				continue
			}
			positiveBranch, ok := positiveArrayCountBranch(operator, literal)
			if !ok || edge.Kind != positiveBranch || value.allocationCountSource == "" || !arrayCountExpressionMatches(lhs, value.allocationCountSource) {
				continue
			}
			if !cloned {
				updated = cloneArrayState(state)
				cloned = true
			}
			value.kind = arrayAllocated
			value.knownArray = true
			// Keep the witness so joins remain conditional rather than turning
			// the positive branch into an unconditional fact for its siblings.
			updated[name] = value
			break
		}
	}
	return updated
}

func selectCaseExpression(text string) string {
	first := strings.TrimSpace(strings.SplitN(text, "\n", 2)[0])
	match := arraySelectCaseRe.FindStringSubmatch(first)
	if len(match) != 2 {
		return ""
	}
	return strings.TrimSpace(match[1])
}

func positiveSelectCaseValue(statement *procedureir.Statement) bool {
	if statement == nil {
		return false
	}
	first := strings.TrimSpace(strings.SplitN(statement.Text, "\n", 2)[0])
	match := arrayPositiveCaseRe.FindStringSubmatch(first)
	if len(match) != 2 {
		return false
	}
	value, err := strconv.Atoi(match[1])
	return err == nil && value > 0
}

func arrayCountComparison(text string) (lhs, operator, literal string, ok bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "if ") {
		text = strings.TrimSpace(text[3:])
	}
	if then := strings.Index(strings.ToLower(text), " then"); then >= 0 {
		text = strings.TrimSpace(text[:then])
	}
	match := arrayCountComparisonRe.FindStringSubmatch(text)
	if len(match) != 4 {
		return "", "", "", false
	}
	return strings.TrimSpace(match[1]), match[2], match[3], true
}

func positiveArrayCountBranch(operator, literal string) (vbacfg.EdgeKind, bool) {
	value, err := strconv.Atoi(literal)
	if err != nil {
		return "", false
	}
	switch operator {
	case "=":
		if value == 0 {
			return vbacfg.EdgeBranchFalse, true
		}
	case "<>":
		if value == 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">":
		if value >= 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">=":
		if value >= 1 {
			return vbacfg.EdgeBranchTrue, true
		}
	case "<":
		if value <= 1 {
			return vbacfg.EdgeBranchFalse, true
		}
	case "<=":
		if value <= 0 {
			return vbacfg.EdgeBranchFalse, true
		}
	}
	return "", false
}

// applyArrayAllocationGuard refines only the branch where a proven array
// length helper returns a positive value. The true branch of IsArray is also
// enough to establish array-ness for a Variant assignment. The opposite
// branch is left at its existing lattice value because an arbitrary caller
// may have additional paths or side effects that this rule cannot prove.
func applyArrayAllocationGuard(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, guards map[string]bool, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	if statement.Condition == nil {
		return state
	}
	if updated, ok := arrayStrPtrGuardState(state, statement.Condition.Text, edge.Kind, variables); ok {
		return arrayVBA227PropagateNonEmptyReturnInputs(updated, statement.Condition.Text, variables)
	}
	if argument, arrayBranch, ok := arrayIsArrayGuardCondition(statement.Condition.Text); ok {
		if name := directArrayArgumentName(argument); name != "" {
			if edge.Kind != arrayBranch {
				return state
			}
			variable, known := variables[name]
			if !known || !variable.isVariant {
				return state
			}
			value, known := state[name]
			if !known {
				return state
			}
			updated := cloneArrayState(state)
			value.kind = arrayAllocated
			value.knownArray = true
			updated[name] = value
			return updated
		}

		// Nested element guards are refined by arrayVBA227Transfer while the
		// source-line block is still being processed. Refining here would apply
		// the fact after later statements in the same block, potentially
		// overwriting an Erase or unknown assignment before the next block.
		return state
	}
	argument, allocatedBranch, ok := arrayAllocationGuardCondition(statement.Condition.Text, guards, state)
	if !ok || edge.Kind != allocatedBranch {
		return state
	}
	name := strings.ToLower(cleanIdentifier(argument))
	variable, known := variables[name]
	if !known || !variable.isArray && !variable.isVariant {
		return state
	}
	value, known := state[name]
	if !known {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	value.mayBeEmpty = false
	updated[name] = value
	return updated
}

// applyArrayNotEmptyGuardBranch recognizes the normal path of VBA's
// `(Not values) = -1` check. The true path is the unallocated/empty case; the
// false path proves that a dynamic array has at least one element.
func applyArrayNotEmptyGuardBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, proc sourceProcedure, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchFalse || statement.Condition == nil || arrayProcedureHasErrorHandling(proc) {
		return state
	}
	updated, ok := arrayNonEmptyGuardState(state, statement.Condition.Text, variables)
	if !ok {
		return state
	}
	return updated
}

// applyArrayAllocationFlagBranch restores the allocation established by a
// local ready flag on its true branch. The relation is attached to the array
// at the flag's proven assignment, so unrelated Boolean conditions remain
// conservative.
func applyArrayAllocationFlagBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue || statement.Condition == nil {
		return state
	}
	flags := arrayVBA227AllocationFlagNames(statement.Condition.Text, variables)
	if len(flags) == 0 {
		return state
	}
	flagSet := make(map[string]bool, len(flags))
	for _, flag := range flags {
		flagSet[flag] = true
	}
	var updated arrayFlowState
	for name, value := range state {
		if !flagSet[value.allocationFlagSource] {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227AllocationFlagNames(text string, variables map[string]arrayVariable) []string {
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
	terms := []string{condition}
	for index := 0; index < len(terms); index++ {
		term := terms[index]
		if split := arrayTopLevelKeywordIndex(term, "and"); split >= 0 {
			terms[index] = strings.TrimSpace(term[:split])
			terms = append(terms, strings.TrimSpace(term[split+len("and"):]))
		}
	}
	flags := make([]string, 0, len(terms))
	seen := map[string]bool{}
	for _, term := range terms {
		term = strings.TrimSpace(term)
		for len(term) >= 2 && term[0] == '(' && term[len(term)-1] == ')' {
			term = strings.TrimSpace(term[1 : len(term)-1])
		}
		lowerTerm := strings.ToLower(term)
		if strings.HasPrefix(lowerTerm, "not ") || !arrayEraseNameRe.MatchString(term) {
			continue
		}
		flag := strings.ToLower(cleanIdentifier(term))
		variable, known := variables[flag]
		if !known || variable.isArray || variable.isVariant || variable.isObject || !strings.EqualFold(strings.TrimSpace(variable.typ), "Boolean") || seen[flag] {
			continue
		}
		seen[flag] = true
		flags = append(flags, flag)
	}
	return flags
}

// applyArraySafeBoundGuard refines the branch where a helper that returns an
// upper bound (or -1 after a caught bounds failure) proves that its array is
// allocated and has a nonnegative upper bound. This is separate from the
// positive-length helper contract because zero is a successful upper-bound
// result, not its failure sentinel.
func applyArraySafeBoundGuard(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, guards map[string]bool, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	if statement.Condition == nil {
		return state
	}
	updated, ok := arraySafeBoundBranchState(state, statement.Condition.Text, edge.Kind, guards, variables)
	if !ok {
		return state
	}
	return updated
}

// applyArrayForBoundState refines the path that can enter a For body after a
// scalar safe-bound helper has returned a nonnegative result. Its -1 sentinel
// intentionally represents a zero-iteration loop, so no fact is propagated to
// the loop-exit path. A direct UBound expression is left to the existing
// bound diagnostic; for source-line CFG blocks, that finding is also the
// conservative evidence for body accesses when the bound can fail.
func applyArrayForBoundState(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, variables map[string]arrayVariable) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeLoopBody {
		return state
	}
	text := strings.TrimSpace(statement.Text)
	if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
		text = strings.TrimSpace(text[:newline])
	}
	text = strings.TrimSpace(normalizedCodeLine(text))
	if _, _, countSource, _, ok := arrayForCountHeader(text); ok {
		if updated, changed := arrayForCountArrayState(state, countSource, variables); changed {
			return updated
		}
	}
	match := arrayForScalarBoundRe.FindStringSubmatch(text)
	if len(match) != 2 {
		return state
	}
	bound, known := state[strings.ToLower(cleanIdentifier(match[1]))]
	if !known || bound.safeBoundProbe == "" {
		return state
	}
	return arrayForBoundArrayState(state, bound.safeBoundProbe, variables)
}

func arrayForCountHeader(text string) (loopVariable, start, countSource string, hasMinusOne bool, ok bool) {
	match := arrayForCountRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) != 5 {
		return "", "", "", false, false
	}
	if match[2] == "0" && strings.TrimSpace(match[4]) == "" {
		// `For i = 0 To items.Count` also enters when i == Count, so it
		// does not prove that an indexed access in the body is in range.
		return "", "", "", false, false
	}
	return match[1], match[2], match[3], strings.TrimSpace(match[4]) != "", true
}

func arrayForCountArrayState(state arrayFlowState, countSource string, variables map[string]arrayVariable) (arrayFlowState, bool) {
	var updated arrayFlowState
	for name, value := range state {
		if value.allocationCountSource == "" || !arrayCountExpressionMatches(countSource, value.allocationCountSource) {
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
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		value.allocationCountSource = ""
		updated[name] = value
	}
	if updated == nil {
		return state, false
	}
	return updated, true
}

func arrayForBoundArrayState(state arrayFlowState, argument string, variables map[string]arrayVariable) arrayFlowState {
	name := strings.ToLower(cleanIdentifier(argument))
	variable, known := variables[name]
	if !known || !variable.isArray && !variable.isVariant {
		return state
	}
	value, known := state[name]
	if !known {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	updated[name] = value
	return updated
}

func arraySafeBoundBranchState(state arrayFlowState, text string, branch vbacfg.EdgeKind, guards map[string]bool, variables map[string]arrayVariable) (arrayFlowState, bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "if ") {
		text = strings.TrimSpace(text[3:])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	functionName, argument, operator, literal, reversed, ok := parseArrayAllocationGuard(text)
	if ok {
		functionName = strings.ToLower(lastName(functionName))
		if !guards[functionName] {
			return state, false
		}
	} else {
		argument, operator, literal, reversed, ok = parseArraySafeBoundProbeVariable(text, state)
		if !ok {
			return state, false
		}
	}
	value, err := strconv.Atoi(literal)
	if err != nil {
		return state, false
	}
	allocatedBranch, ok := safeBoundNonnegativeBranch(operator, value, reversed)
	if !ok || branch != allocatedBranch {
		return state, false
	}
	name := strings.ToLower(cleanIdentifier(argument))
	variable, known := variables[name]
	if !known || !variable.isArray {
		return state, false
	}
	current, known := state[name]
	if !known {
		return state, false
	}
	updated := cloneArrayState(state)
	current.kind = arrayAllocated
	current.knownArray = true
	current.mayBeEmpty = false
	updated[name] = current
	return updated, true
}

func safeBoundNonnegativeBranch(operator string, value int, reversed bool) (vbacfg.EdgeKind, bool) {
	if reversed {
		switch operator {
		case ">":
			operator = "<"
		case ">=":
			operator = "<="
		case "<":
			operator = ">"
		case "<=":
			operator = ">="
		}
	}
	switch operator {
	case "=":
		if value >= 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">":
		if value >= -1 {
			return vbacfg.EdgeBranchTrue, true
		}
	case ">=":
		if value >= 0 {
			return vbacfg.EdgeBranchTrue, true
		}
	case "<":
		if value >= 0 {
			return vbacfg.EdgeBranchFalse, true
		}
	case "<=":
		if value >= -1 {
			return vbacfg.EdgeBranchFalse, true
		}
	}
	return "", false
}

func arrayIsArrayGuardCondition(text string) (string, vbacfg.EdgeKind, bool) {
	text = strings.TrimSpace(text)
	if condition, _, ok := arrayIfThenParts(text); ok {
		text = condition
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "if ") {
		text = strings.TrimSpace(text[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		text = strings.TrimSpace(text[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	negated := false
	if strings.HasPrefix(strings.ToLower(text), "not ") {
		negated = true
		text = strings.TrimSpace(text[4:])
	}
	match := arrayIsArrayGuardRe.FindStringSubmatch(text)
	if len(match) == 2 {
		branch := vbacfg.EdgeBranchTrue
		if negated {
			branch = vbacfg.EdgeBranchFalse
		}
		return match[1], branch, true
	}
	if negated {
		return "", "", false
	}
	if match := arrayByteArrayGuardRe.FindStringSubmatch(text); len(match) == 2 {
		return match[1], vbacfg.EdgeBranchTrue, true
	}
	return "", "", false
}

func arrayStrPtrGuardParts(text string, variables map[string]arrayVariable) ([]string, string, bool) {
	text = strings.TrimSpace(text)
	if condition, _, ok := arrayIfThenParts(text); ok {
		text = condition
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "if ") {
		text = strings.TrimSpace(text[3:])
	} else if strings.HasPrefix(lower, "elseif ") {
		text = strings.TrimSpace(text[len("elseif "):])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	if arrayConditionAndRe.MatchString(text) {
		return nil, "", false
	}
	parts := arrayConditionOrRe.Split(text, -1)
	if len(parts) == 0 {
		return nil, "", false
	}
	names := make([]string, 0, len(parts))
	singleOperator := ""
	for _, part := range parts {
		match := arrayStrPtrGuardRe.FindStringSubmatch(strings.TrimSpace(part))
		if len(match) != 3 {
			return nil, "", false
		}
		if len(parts) > 1 && match[2] != "=" {
			return nil, "", false
		}
		if len(parts) == 1 {
			singleOperator = match[2]
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		variable, knownVariable := variables[name]
		if !knownVariable || !isByteArrayVariable(variable) {
			return nil, "", false
		}
		names = append(names, name)
	}
	return names, singleOperator, true
}

// arrayStrPtrGuardState recognizes the established VBA Byte-array idiom
// `If StrPtr(values) = 0 Then ...`: the false branch is reached only when the
// array has a usable element. A compound zero-pointer guard joined by Or has
// the same property for every operand on its false branch. Keep this limited
// to declared arrays; StrPtr on arbitrary Variants or scalar expressions does
// not establish an array contract for this rule.
func arrayStrPtrGuardState(state arrayFlowState, text string, branch vbacfg.EdgeKind, variables map[string]arrayVariable) (arrayFlowState, bool) {
	names, singleOperator, ok := arrayStrPtrGuardParts(text, variables)
	if !ok {
		return state, false
	}
	for _, name := range names {
		if _, knownValue := state[name]; !knownValue {
			return state, false
		}
	}
	if len(names) > 1 {
		if branch != vbacfg.EdgeBranchFalse {
			return state, false
		}
	} else {
		requiredBranch := vbacfg.EdgeBranchFalse
		if singleOperator == "<>" {
			requiredBranch = vbacfg.EdgeBranchTrue
		}
		if branch != requiredBranch {
			return state, false
		}
	}
	updated := cloneArrayState(state)
	for _, name := range names {
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		updated[name] = value
	}
	return updated, true
}

// arrayVBA227PropagateNonEmptyReturnInputs transfers a caller-side input-array
// fact through a returned Byte-array value. The transfer is intentionally
// driven by the StrPtr branch that already proved the returned value non-empty;
// a bare return summary never establishes the input fact by itself.
func arrayVBA227PropagateNonEmptyReturnInputs(state arrayFlowState, text string, variables map[string]arrayVariable) arrayFlowState {
	names, _, ok := arrayStrPtrGuardParts(text, variables)
	if !ok {
		return state
	}
	updated := cloneArrayState(state)
	changed := false
	for _, name := range names {
		value, known := updated[name]
		if !known || value.nonEmptySource == "" {
			continue
		}
		source := strings.ToLower(cleanIdentifier(value.nonEmptySource))
		variable, knownVariable := variables[source]
		sourceValue, knownSource := updated[source]
		if !knownVariable || !variable.isArray || !knownSource {
			continue
		}
		sourceValue.kind = arrayAllocated
		sourceValue.knownArray = true
		sourceValue.mayBeEmpty = false
		updated[source] = sourceValue
		changed = true
	}
	if !changed {
		return state
	}
	return updated
}

func arrayElementBaseName(text string) string {
	text = strings.TrimSpace(text)
	open := strings.IndexByte(text, '(')
	if open <= 0 || matchingParen(text, open) != len(text)-1 {
		return ""
	}
	if strings.TrimSpace(text[open+1:len(text)-1]) == "" {
		return ""
	}
	return directArrayArgumentName(text[:open])
}

func arrayElementGuardState(state arrayFlowState, argument string, variables map[string]arrayVariable) arrayFlowState {
	name := arrayElementBaseName(argument)
	if name == "" {
		return state
	}
	variable, known := variables[name]
	if !known || !variable.isArray {
		return state
	}
	value, known := state[name]
	if !known {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	value.mayBeEmpty = false
	updated[name] = value
	return updated
}

// arrayVBA227Graph removes normal-flow edges after direct raises and after
// private helpers whose normal CFG has no path to the procedure exit.  The
// latter covers project-local error wrappers such as RaiseContractError:
// their call sites must not poison the normal allocation state with an
// impossible fall-through branch.
func arrayVBA227Graph(proc sourceProcedure, ctx analysisContext) vbacfg.CFGView {
	if proc.Graph == nil {
		return vbacfg.CFGView{}
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	removed := map[vbacfg.BlockID]bool{}
	for call := range proc.Calls.All() {
		_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !ok || !arrayProcedureAlwaysRaises(target) {
			continue
		}
		block, ok := graph.BlockForStatement(call.StatementID)
		if ok {
			removed[block.ID] = true
		}
	}
	if len(removed) == 0 {
		return graph
	}
	return graph.WithoutNormalContinuationsFrom(removed)
}

func arrayProcedureAlwaysRaises(proc sourceProcedure) bool {
	if proc.Graph == nil {
		return false
	}
	graph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true, WithoutNormalErrRaiseContinuation: true})
	return !graph.IsReachable(graph.NormalExit())
}

func arrayAllocationGuardCondition(text string, guards map[string]bool, state arrayFlowState) (string, vbacfg.EdgeKind, bool) {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "if ") {
		text = strings.TrimSpace(text[3:])
	}
	if then := strings.LastIndex(strings.ToLower(text), " then"); then >= 0 && strings.TrimSpace(text[then+5:]) == "" {
		text = strings.TrimSpace(text[:then])
	}
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}

	functionName, argument, operator, literal, reversed, ok := parseArrayAllocationGuard(text)
	if ok {
		functionName = strings.ToLower(lastName(functionName))
		if _, ok := guards[functionName]; !ok {
			return "", "", false
		}
	} else {
		argument, operator, literal, reversed, ok = parseArrayAllocationProbeVariable(text, state)
		if !ok {
			return "", "", false
		}
	}

	if operator == "" {
		return argument, vbacfg.EdgeBranchTrue, true
	}
	value, err := strconv.Atoi(literal)
	if err != nil {
		return "", "", false
	}
	if reversed {
		switch operator {
		case ">":
			operator = "<"
		case ">=":
			operator = "<="
		case "<":
			operator = ">"
		case "<=":
			operator = ">="
		}
	}

	// The helper's proven return domain is zero for the handled failure path
	// and a positive array length after successful bounds inspection.
	switch {
	case operator == "=" && value == 0:
		return argument, vbacfg.EdgeBranchFalse, true
	case operator == "<>" && value == 0:
		return argument, vbacfg.EdgeBranchTrue, true
	case operator == "<>" && value == 1:
		// A dimension-count probe returns zero for an unallocated value and
		// one for a one-dimensional array. Its `<> 1` rejection branch is
		// therefore safe to leave only on the false edge, just like the
		// ordinary positive-length probe's zero check.
		return argument, vbacfg.EdgeBranchFalse, true
	case operator == ">" && value >= 0:
		return argument, vbacfg.EdgeBranchTrue, true
	case operator == ">=" && value >= 1:
		return argument, vbacfg.EdgeBranchTrue, true
	case operator == "<" && value <= 0:
		return argument, vbacfg.EdgeBranchFalse, true
	case operator == "<=" && value < 0:
		return argument, vbacfg.EdgeBranchFalse, true
	default:
		return "", "", false
	}
}

func parseArrayAllocationGuard(text string) (functionName, argument, operator, literal string, reversed, ok bool) {
	if match := arrayGuardCallRe.FindStringSubmatch(text); len(match) == 5 {
		return match[1], match[2], match[3], match[4], false, true
	}
	if match := arrayGuardReversedRe.FindStringSubmatch(text); len(match) == 5 {
		return match[3], match[4], match[2], match[1], true, true
	}
	return "", "", "", "", false, false
}

func parseArrayAllocationProbeVariable(text string, state arrayFlowState) (argument, operator, literal string, reversed, ok bool) {
	if match := arrayGuardValueRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[1]))]
		if !exists || value.allocationProbe == "" {
			return "", "", "", false, false
		}
		return value.allocationProbe, match[2], match[3], false, true
	}
	if match := arrayGuardValueReversedRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[3]))]
		if !exists || value.allocationProbe == "" {
			return "", "", "", false, false
		}
		return value.allocationProbe, match[2], match[1], true, true
	}
	return "", "", "", false, false
}

func parseArraySafeBoundProbeVariable(text string, state arrayFlowState) (argument, operator, literal string, reversed, ok bool) {
	if match := arrayGuardValueRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[1]))]
		if !exists || value.safeBoundProbe == "" {
			return "", "", "", false, false
		}
		return value.safeBoundProbe, match[2], match[3], false, true
	}
	if match := arrayGuardValueReversedRe.FindStringSubmatch(text); len(match) == 4 {
		value, exists := state[strings.ToLower(cleanIdentifier(match[3]))]
		if !exists || value.safeBoundProbe == "" {
			return "", "", "", false, false
		}
		return value.safeBoundProbe, match[2], match[1], true, true
	}
	return "", "", "", false, false
}

func arrayAllocationProbeArgument(text string, guards map[string]bool) (string, bool) {
	functionName, argument, operator, literal, _, ok := parseArrayAllocationGuard(text)
	if !ok || operator != "" || literal != "" {
		return "", false
	}
	if _, ok := guards[strings.ToLower(lastName(functionName))]; !ok {
		return "", false
	}
	return strings.ToLower(cleanIdentifier(argument)), true
}

func arraySafeBoundProbeArgument(text string, guards map[string]bool) (string, bool) {
	functionName, argument, operator, literal, _, ok := parseArrayAllocationGuard(text)
	if !ok || operator != "" || literal != "" {
		return "", false
	}
	if _, ok := guards[strings.ToLower(lastName(functionName))]; !ok {
		return "", false
	}
	return strings.ToLower(cleanIdentifier(argument)), true
}

func mergeArrayState(states map[vbacfg.BlockID]arrayFlowState, id vbacfg.BlockID, incoming arrayFlowState) bool {
	current, exists := states[id]
	if !exists {
		states[id] = cloneArrayState(incoming)
		return true
	}
	merged := meetArrayState(current, incoming)
	if arrayStateEqual(current, merged) {
		return false
	}
	states[id] = merged
	return true
}

func meetArrayState(left, right arrayFlowState) arrayFlowState {
	out := arrayFlowState{}
	keys := map[string]bool{}
	for key := range left {
		keys[key] = true
	}
	for key := range right {
		keys[key] = true
	}
	for key := range keys {
		l, lok := left[key]
		r, rok := right[key]
		if !lok || !rok {
			out[key] = arrayValue{kind: arrayUnknown}
			continue
		}
		out[key] = meetArrayValue(l, r)
	}
	return out
}

func meetArrayValue(left, right arrayValue) arrayValue {
	out := arrayValue{
		kind:                            left.kind,
		knownArray:                      left.knownArray,
		mayBeEmpty:                      left.mayBeEmpty,
		origin:                          left.origin,
		dimensions:                      append([]arrayDimension(nil), left.dimensions...),
		preserveShape:                   append([]arrayDimension(nil), left.preserveShape...),
		allocationCountSource:           left.allocationCountSource,
		conditionalAllocationSource:     left.conditionalAllocationSource,
		allocationFlagSource:            left.allocationFlagSource,
		returnNonEmptyArrayParameter:    left.returnNonEmptyArrayParameter,
		returnPositiveScalarParameter:   left.returnPositiveScalarParameter,
		nonEmptySource:                  left.nonEmptySource,
		returnDescriptorSourceParameter: left.returnDescriptorSourceParameter,
		returnDescriptorStartParameter:  left.returnDescriptorStartParameter,
		returnDescriptorLengthParameter: left.returnDescriptorLengthParameter,
		returnDescriptorLowerParameter:  left.returnDescriptorLowerParameter,
		boundsProof:                     left.boundsProof,
	}
	if left.kind != right.kind {
		out.kind = arrayUnknown
	}
	if left.knownArray != right.knownArray {
		out.knownArray = false
	}
	if left.conditionalAllocationSource != right.conditionalAllocationSource {
		if left.conditionalAllocationSource == "" {
			out.conditionalAllocationSource = right.conditionalAllocationSource
		} else if right.conditionalAllocationSource == "" {
			out.conditionalAllocationSource = left.conditionalAllocationSource
		} else {
			out.conditionalAllocationSource = ""
		}
	}
	if left.allocationFlagSource != right.allocationFlagSource {
		if left.allocationFlagSource == "" {
			out.allocationFlagSource = right.allocationFlagSource
		} else if right.allocationFlagSource == "" {
			out.allocationFlagSource = left.allocationFlagSource
		} else {
			out.allocationFlagSource = ""
		}
	}
	if left.returnNonEmptyArrayParameter != right.returnNonEmptyArrayParameter {
		out.returnNonEmptyArrayParameter = ""
	}
	if left.returnPositiveScalarParameter != right.returnPositiveScalarParameter {
		out.returnPositiveScalarParameter = ""
	}
	if left.nonEmptySource != right.nonEmptySource {
		out.nonEmptySource = ""
	}
	if left.returnDescriptorSourceParameter != right.returnDescriptorSourceParameter ||
		left.returnDescriptorStartParameter != right.returnDescriptorStartParameter ||
		left.returnDescriptorLengthParameter != right.returnDescriptorLengthParameter ||
		left.returnDescriptorLowerParameter != right.returnDescriptorLowerParameter {
		out.returnDescriptorSourceParameter = ""
		out.returnDescriptorStartParameter = ""
		out.returnDescriptorLengthParameter = ""
		out.returnDescriptorLowerParameter = ""
	}
	out.mayBeEmpty = left.mayBeEmpty || right.mayBeEmpty
	if left.origin != right.origin {
		out.origin = arrayOriginUnknown
	}
	if !arrayDimensionsEqual(left.dimensions, right.dimensions) {
		out.dimensions = nil
	}
	out.preserveShape = meetArrayDimensions(left.preserveShape, right.preserveShape)
	if left.allocationProbe != right.allocationProbe {
		out.allocationProbe = ""
	} else {
		out.allocationProbe = left.allocationProbe
	}
	if left.safeBoundProbe != right.safeBoundProbe {
		out.safeBoundProbe = ""
	} else {
		out.safeBoundProbe = left.safeBoundProbe
	}
	if left.allocationCountSource != right.allocationCountSource {
		out.allocationCountSource = ""
	}
	if left.boundsProof != right.boundsProof {
		out.boundsProof = arrayBoundsProof{}
	}
	return out
}

func meetArrayDimensions(left, right []arrayDimension) []arrayDimension {
	if len(left) == 0 || len(left) != len(right) {
		return nil
	}
	out := make([]arrayDimension, len(left))
	for i := range left {
		out[i] = arrayDimension{
			lower: meetArrayBound(left[i].lower, right[i].lower),
			upper: meetArrayBound(left[i].upper, right[i].upper),
		}
	}
	return out
}

func meetArrayBound(left, right arrayBound) arrayBound {
	if arrayBoundsEquivalent(left, right) {
		return left
	}
	return arrayBound{}
}

func cloneArrayState(state arrayFlowState) arrayFlowState {
	out := arrayFlowState{}
	for name, value := range state {
		value.dimensions = append([]arrayDimension(nil), value.dimensions...)
		value.preserveShape = append([]arrayDimension(nil), value.preserveShape...)
		out[name] = value
	}
	return out
}

func arrayStateEqual(left, right arrayFlowState) bool {
	if len(left) != len(right) {
		return false
	}
	for key, l := range left {
		r, ok := right[key]
		if !ok || l.kind != r.kind || l.knownArray != r.knownArray || l.mayBeEmpty != r.mayBeEmpty || l.origin != r.origin || l.allocationProbe != r.allocationProbe || l.safeBoundProbe != r.safeBoundProbe || l.allocationCountSource != r.allocationCountSource || l.conditionalAllocationSource != r.conditionalAllocationSource || l.allocationFlagSource != r.allocationFlagSource || l.returnNonEmptyArrayParameter != r.returnNonEmptyArrayParameter || l.returnPositiveScalarParameter != r.returnPositiveScalarParameter || l.nonEmptySource != r.nonEmptySource || l.returnDescriptorSourceParameter != r.returnDescriptorSourceParameter || l.returnDescriptorStartParameter != r.returnDescriptorStartParameter || l.returnDescriptorLengthParameter != r.returnDescriptorLengthParameter || l.returnDescriptorLowerParameter != r.returnDescriptorLowerParameter || l.boundsProof != r.boundsProof || !arrayDimensionsEqual(l.dimensions, r.dimensions) || !arrayDimensionsEqual(l.preserveShape, r.preserveShape) {
			return false
		}
	}
	return true
}

func arrayDimensionsEqual(left, right []arrayDimension) bool {
	if len(left) != len(right) {
		return false
	}
	for i := range left {
		if left[i] != right[i] {
			return false
		}
	}
	return true
}

func arrayInitialState(variables map[string]arrayVariable) arrayFlowState {
	state := arrayFlowState{}
	for name, variable := range variables {
		// Only arrays and Variants can contribute to the allocation lattice.
		// Scalar/object entries remain available in `variables` for declaration
		// diagnostics, but carrying them through every CFG state is prohibitively
		// expensive for large modules.
		if !variable.isArray && !variable.isVariant {
			continue
		}
		if variable.isVariant && !variable.isArray {
			state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
			continue
		}
		if !variable.isArray {
			state[name] = arrayValue{kind: arrayUnknown, knownArray: false, origin: arrayOriginLocal}
			continue
		}
		value := arrayValue{
			knownArray:    variable.isArray,
			origin:        arrayOriginLocal,
			dimensions:    append([]arrayDimension(nil), variable.dimensions...),
			preserveShape: append([]arrayDimension(nil), variable.dimensions...),
		}
		if variable.fixed {
			value.kind = arrayAllocated
		} else if variable.parameter {
			if variable.paramArray {
				// ParamArray is materialized as an array even when the caller
				// supplies no arguments.  Its rank and bounds may be unknown,
				// but allocation itself is guaranteed by the procedure contract.
				value.kind = arrayAllocated
			} else {
				value.kind = arrayUnknown
			}
		} else {
			value.kind = arrayUnallocated
		}
		state[name] = value
	}
	return state
}

// applyArrayStaticInitializationState recognizes the narrow one-time setup
// idiom used by procedures that keep a reusable backing array in a Static
// local. The static readiness flag is passed to a resolved ByRef helper that
// sets its flag field on every normal return, and the first call then performs
// a successful direct ReDim before any array use. This lets the entry state
// carry the normal-call invariant without treating arbitrary Static arrays as
// allocated.
func applyArrayStaticInitializationState(state arrayFlowState, file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable) arrayFlowState {
	var updated arrayFlowState
	for name, variable := range variables {
		if !variable.static || !variable.isArray || variable.fixed || !arrayStaticArrayInitializationProven(file, proc, ctx, variables, name) {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeEmpty = false
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayStaticArrayInitializationProven(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, targetName string) bool {
	targetName = strings.ToLower(cleanIdentifier(targetName))
	target, ok := variables[targetName]
	if !ok || !target.static || !target.isArray || target.fixed || target.parameter {
		return false
	}
	if proc.StartLine < 1 || proc.EndLine < proc.StartLine || proc.StartLine > len(file.Lines) {
		return false
	}

	type staticReadyGuard struct {
		name  string
		index int
	}
	guards := make([]staticReadyGuard, 0, 1)
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), proc.EndLine)
	for index := start + 1; index < end; index++ {
		match := arrayStaticReadyGuardRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(file.Lines[index])))
		if len(match) != 2 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		ready, declared := variables[name]
		if !declared || !ready.static || ready.parameter || ready.isArray || ready.isVariant || ready.isObject || ready.knownScalar {
			// The readiness object must be a Static UDT-like local. A scalar
			// Boolean guard or a non-static value does not carry state across
			// calls and must remain on the ordinary CFG path.
			continue
		}
		guards = append(guards, staticReadyGuard{name: name, index: index})
	}
	if len(guards) != 1 {
		return false
	}
	guard := guards[0]

	redimIndex, ok := arrayStaticInitializationBlock(file, guard.index, end, targetName)
	if !ok {
		return false
	}
	// The target must not be used before the normal-path ReDim. Otherwise the
	// entry invariant would hide a genuine first-call access before setup.
	for index := start + 1; index < redimIndex; index++ {
		if arrayStaticSourceUsesTarget(file.Lines[index], targetName, variables) {
			return false
		}
	}
	for call := range proc.Calls.All() {
		if call.Range.StartLine-1 < redimIndex && arrayCallPassesDirectArrayArgument(proc, call, targetName) {
			return false
		}
	}

	// Admit exactly one direct ReDim for this target, and reject Erase or a
	// whole-array replacement anywhere in the procedure. Indexed writes after
	// the setup are harmless and are intentionally not filtered here.
	for index := start + 1; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
			usesTarget := false
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct && strings.EqualFold(cleanIdentifier(redim.name), targetName) {
					usesTarget = true
					if strings.TrimSpace(match[1]) != "" || index != redimIndex {
						return false
					}
				}
			}
			if usesTarget && index != redimIndex {
				return false
			}
		}
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 && strings.EqualFold(strings.TrimSpace(match[1]), targetName) {
			return false
		}
		if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), targetName) {
			return false
		}
	}

	initializerFound := false
	for call := range proc.Calls.All() {
		line := call.Range.StartLine - 1
		if line <= guard.index || line >= redimIndex {
			continue
		}
		if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
			return false
		}
		if initializerFound {
			return false
		}
		helper, parameter, resolved := arrayStaticReadyInitializer(file, proc, call, guard.name, ctx)
		if !resolved || helper.StartByte == proc.StartByte || !arrayStaticHelperSetsReadyFlag(file, helper, parameter) {
			return false
		}
		initializerFound = true
	}
	if !initializerFound {
		return false
	}

	// Keep the proof tied to the same straight-line pre-ReDim region. The
	// post-ReDim body may contain the implementation's indexed writes and
	// loops, but a conditional or loop before allocation would leave a bypass.
	for index := guard.index + 1; index < redimIndex; index++ {
		if arrayStaticPreRedimControlFlow(file.Lines[index]) {
			return false
		}
	}
	return true
}

func arrayStaticInitializationBlock(file parsedFile, guardIndex, end int, targetName string) (redimIndex int, ok bool) {
	if guardIndex < 0 || guardIndex >= end {
		return 0, false
	}
	ifDepth := 1
	redimIndex = -1
	for index := guardIndex + 1; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		lower := strings.ToLower(text)
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") {
			continue
		}
		if lower == "end if" {
			if ifDepth == 1 {
				return redimIndex, redimIndex >= 0
			}
			ifDepth--
			continue
		}
		if ifDepth == 1 && (lower == "else" || strings.HasPrefix(lower, "elseif ")) {
			return 0, false
		}
		if arrayStaticBlockIfStart(text) {
			if ifDepth == 1 && redimIndex < 0 {
				return 0, false
			}
			ifDepth++
			continue
		}
		if ifDepth == 1 && redimIndex < 0 && arrayStaticPreRedimControlFlow(file.Lines[index]) {
			return 0, false
		}
		if ifDepth != 1 {
			continue
		}
		match := arrayRedimRe.FindStringSubmatch(text)
		if len(match) == 0 {
			continue
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), targetName) {
				if strings.TrimSpace(match[1]) != "" || redimIndex >= 0 {
					return 0, false
				}
				redimIndex = index
			}
		}
	}
	return 0, false
}

func arrayStaticReadyInitializer(file parsedFile, caller sourceProcedure, call procedureir.CallSite, readyName string, ctx analysisContext) (sourceProcedure, string, bool) {
	resolution := arrayCallResolution(ctx, call)
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return sourceProcedure{}, "", false
	}
	want := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
	var helper sourceProcedure
	found := false
	for _, candidate := range file.Procedures {
		if arrayProcedureKey(candidate) != want {
			continue
		}
		if found {
			return sourceProcedure{}, "", false
		}
		helper = candidate
		found = true
	}
	if !found {
		return sourceProcedure{}, "", false
	}
	bindings, ok := arrayCallArgumentBindings(caller, helper, call)
	if !ok {
		return sourceProcedure{}, "", false
	}
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= helper.Params.Len() || directArrayArgumentName(binding.text) != strings.ToLower(cleanIdentifier(readyName)) {
			continue
		}
		parameter := helper.Params.valueAt(binding.parameterIndex)
		if parameterIsByRefScalar(parameter) && !arrayKnownScalarType(parameter.Type) && !isObjectType(parameter.Type) {
			return helper, helper.Params.valueAt(binding.parameterIndex).Name, true
		}
	}
	return sourceProcedure{}, "", false
}

func arrayStaticHelperSetsReadyFlag(file parsedFile, helper sourceProcedure, parameter string) bool {
	want := canonicalArrayBoundExpression(parameter + ".isSet")
	count := 0
	setLine := -1
	lastExecutable := -1
	depth := 0
	start := max(0, helper.StartLine-1)
	end := min(len(file.Lines), helper.EndLine)
	for index := start; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		lower := strings.ToLower(text)
		if arrayStaticExecutableSourceLine(text) {
			lastExecutable = index
		}
		if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") {
			continue
		}
		if arrayStaticBlockEnd(text) {
			if depth == 0 {
				return false
			}
			depth--
			continue
		}
		if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed && canonicalArrayBoundExpression(lhs) == want {
			if depth != 0 || !strings.EqualFold(strings.TrimSpace(rhs), "true") {
				return false
			}
			count++
			setLine = index
		}
		if arrayStaticBlockStart(text) {
			depth++
		}
		if strings.HasPrefix(lower, "on error resume next") {
			return false
		}
	}
	return depth == 0 && count == 1 && setLine == lastExecutable && setLine >= start
}

func arrayStaticSourceUsesTarget(text, targetName string, variables map[string]arrayVariable) bool {
	for _, use := range arrayIndexedUsesForSource(text, variables) {
		if strings.EqualFold(cleanIdentifier(use.name), targetName) && len(use.args) > 0 {
			return true
		}
	}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		if len(bound) > 2 && strings.EqualFold(cleanIdentifier(bound[2]), targetName) {
			return true
		}
	}
	if match := arrayForEachRe.FindStringSubmatch(strings.TrimSpace(text)); len(match) == 2 && strings.EqualFold(cleanIdentifier(match[1]), targetName) {
		return true
	}
	return false
}

func arrayStaticExecutableSourceLine(text string) bool {
	text = strings.TrimSpace(normalizedCodeLine(text))
	if text == "" || strings.HasPrefix(text, "'") || strings.HasPrefix(text, "#") || isProcedureHeaderLine(strings.ToLower(text)) {
		return false
	}
	switch strings.ToLower(text) {
	case "end sub", "end function", "end property", "else":
		return false
	default:
		return !arrayStaticBlockEnd(text)
	}
}

func arrayStaticBlockIfStart(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return (strings.HasPrefix(lower, "if ") || strings.HasPrefix(lower, "elseif ")) && strings.HasSuffix(lower, " then")
}

func arrayStaticBlockStart(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if arrayStaticBlockIfStart(lower) {
		return true
	}
	for _, prefix := range []string{"for ", "do", "while ", "select ", "with "} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}

func arrayStaticBlockEnd(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{"end if", "end with", "end select", "next", "loop", "wend"} {
		if lower == prefix || strings.HasPrefix(lower, prefix+" ") {
			return true
		}
	}
	return false
}

func arrayStaticPreRedimControlFlow(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(text)))
	if lower == "" || strings.HasPrefix(lower, "'") || strings.HasPrefix(lower, "#") {
		return false
	}
	if arrayStaticBlockIfStart(lower) {
		return true
	}
	for _, prefix := range []string{"for ", "do", "while ", "select ", "with ", "else", "goto ", "on error ", "exit "} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}

func applyArrayByRefEntryStates(state arrayFlowState, proc sourceProcedure, variables map[string]arrayVariable, entries map[string]map[int]bool, conditions map[string]map[int]string) arrayFlowState {
	parameters := entries[arrayProcedureKey(proc)]
	conditionalParameters := conditions[arrayProcedureKey(proc)]
	if len(parameters) == 0 && len(conditionalParameters) == 0 {
		return state
	}
	updated := cloneArrayState(state)
	for index, allocated := range parameters {
		if !allocated || index < 0 || index >= proc.Params.Len() {
			continue
		}
		name := strings.ToLower(proc.Params.valueAt(index).Name)
		variable, known := variables[name]
		value, exists := updated[name]
		if !known || !exists || !variable.isArray {
			continue
		}
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	for index, source := range conditionalParameters {
		if source == "" || index < 0 || index >= proc.Params.Len() {
			continue
		}
		name := strings.ToLower(proc.Params.valueAt(index).Name)
		variable, known := variables[name]
		value, exists := updated[name]
		if !known || !exists || !variable.isArray || value.kind == arrayAllocated && value.knownArray {
			continue
		}
		if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, source) {
			continue
		}
		value.allocationCountSource = source
		updated[name] = value
	}
	return updated
}

func arrayProcedureKey(proc sourceProcedure) string {
	module := strings.TrimSpace(proc.Module)
	name := strings.TrimSpace(proc.Name)
	if module == "" {
		return strings.ToLower(name)
	}
	if name == "" {
		return strings.ToLower(module)
	}
	return strings.ToLower(module + "." + name)
}

func arrayParticipantProcedureIdentity(proc sourceProcedure) string {
	path := arrayProcedureSourcePath(proc)
	if path == "" {
		path = strings.ToLower(strings.TrimSpace(proc.Module))
	}
	return strings.Join([]string{
		path,
		strconv.Itoa(proc.Index),
		strconv.Itoa(proc.StartByte),
		strconv.Itoa(proc.StartLine),
		strings.ToLower(strings.TrimSpace(proc.Name)),
		strings.ToLower(string(proc.ProcedureKind)),
	}, "\x00")
}

func arrayParticipantDisambiguatedKey(proc sourceProcedure) string {
	return arrayProcedureKey(proc) + "|" + arrayParticipantProcedureIdentity(proc)
}

func arrayParticipantLookupKey(proc sourceProcedure, participantKeys map[string]string) string {
	if len(participantKeys) > 0 {
		if key := participantKeys[arrayParticipantProcedureIdentity(proc)]; key != "" {
			return key
		}
	}
	return arrayProcedureKey(proc)
}

func arrayProcedureSourcePath(proc sourceProcedure) string {
	if proc.Document == nil {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(proc.Document.Path))
}

func arrayProcedureLess(left, right sourceProcedure) bool {
	leftPath, rightPath := arrayProcedureSourcePath(left), arrayProcedureSourcePath(right)
	if leftPath != rightPath {
		return leftPath < rightPath
	}
	if left.StartByte != right.StartByte {
		return left.StartByte < right.StartByte
	}
	if left.StartLine != right.StartLine {
		return left.StartLine < right.StartLine
	}
	if !strings.EqualFold(left.Module, right.Module) {
		return strings.ToLower(left.Module) < strings.ToLower(right.Module)
	}
	if !strings.EqualFold(left.Name, right.Name) {
		return strings.ToLower(left.Name) < strings.ToLower(right.Name)
	}
	if left.ProcedureKind != right.ProcedureKind {
		return string(left.ProcedureKind) < string(right.ProcedureKind)
	}
	return arrayProcedureKey(left) < arrayProcedureKey(right)
}
