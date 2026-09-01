package analyze

import (
	"regexp"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func (a Analyzer) arrayVBA227Transfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, resumeNextBefore []bool) (arrayFlowState, []Finding) {
	state = arrayVBA227ClearLoopBodyBounds(state, line)
	state = arrayVBA227ClearConditionalAllocationGuards(state, proc, text, line, variables)
	if strings.HasPrefix(strings.ToLower(strings.TrimSpace(text)), "case ") || arrayBoundCallRe.MatchString(text) {
		state = arrayVBA227RepeatedSelectCaseBoundsState(file, proc, line, state, variables)
	}
	transfer := func(input arrayFlowState, source string) (arrayFlowState, []Finding) {
		output, findings := a.arrayTransfer(file, proc, ctx, variables, input, source, line, constants, capacityGuards)
		output = arrayVBA227AttachConditionalReDimState(output, proc, source, line, variables)
		output = arrayVBA227AttachReturnProvenance(output, source, ctx, variables, constants)
		output = arrayVBA227AttachAllocationFlagState(file, proc, source, line, input, output, variables)
		if resumeNextBefore == nil || arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
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
			state, findings := transfer(state, condition)
			return arraySuccessfulBoundsState(state, condition, variables, arrayVBA227LoopBodyEndLine(proc, line)), findings
		}
	}
	state, findings := transfer(state, text)
	findings = arrayVBA227FilterSuccessfulBoundsGuardBodyIndexFindings(findings, file, proc, line, variables, resumeNextBefore)
	findings = arrayVBA227FilterConditionalBodyIndexFindings(findings, file, proc, line, state, variables, ctx, resumeNextBefore)
	findings = arrayVBA227FilterForBodyIndexFindings(findings, file, proc, line, state, variables, ctx, resumeNextBefore)
	if (arrayVBA227HasSuccessfulBoundsExpression(text) || arrayVBA227HasDictionaryBoundsExpression(text, state)) &&
		!arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) &&
		!strings.Contains(strings.ToLower(text), "on error resume next") {
		state = arraySuccessfulBoundsState(state, text, variables, arrayVBA227LoopBodyEndLine(proc, line))
	}
	// Source-line CFG blocks can contain an If condition and its body. Apply
	// the normal-path fact after the condition while the block is still being
	// processed so a nested element access in the body does not repeat the
	// condition's possible outer-array failure.
	if argument, _, ok := arrayIsArrayGuardCondition(text); ok {
		state = arrayElementGuardState(state, argument, variables)
	}
	if name, ok := arraySafeArrayPointerGuardTarget(file, proc, line, text, variables); ok {
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
func arraySafeArrayPointerGuardTarget(file parsedFile, proc sourceProcedure, line int, text string, variables map[string]arrayVariable) (string, bool) {
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
	if len(copyMatch) != 4 || !strings.EqualFold(copyMatch[1], descriptorName) || !strings.EqualFold(copyMatch[1], copyMatch[3]) {
		return "", false
	}
	ptrName := strings.ToLower(cleanIdentifier(copyMatch[2]))
	ptrGuardMatch := arraySafeArrayZeroExitGuardRe.FindStringSubmatch(ptrGuard)
	if len(ptrGuardMatch) != 2 || !strings.EqualFold(ptrGuardMatch[1], ptrName) {
		return "", false
	}
	lhs, rhs, indexed, assigned := arrayAssignment(pointerAssignment)
	if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), ptrName) || !strings.EqualFold(arrayCallName(rhs), "varptrarray") {
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
	for line := proc.StartLine; line < loop.Range.StartLine && line <= len(file.Lines); line++ {
		if line == indexZeroLine {
			continue
		}
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, _, indexed, assigned := arrayAssignment(source)
			if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), indexName) {
				return "", false
			}
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
	for line := proc.StartLine; line < lengthLine && line <= len(file.Lines); line++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, _, indexed, assigned := arrayAssignment(source)
			if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				return "", false
			}
		}
	}
	sourceLine, ok := arrayVBA227ZeroBasedArraySourceLine(file, proc, arrayName, lengthLine)
	if !ok || sourceLine == 0 || lengthLine <= sourceLine || !arrayVBA227NoArrayMutationBetween(file, proc, arrayName, sourceLine+1, lengthLine) {
		return "", false
	}
	if accessLine <= lengthLine || !arrayVBA227NoArrayMutationBetween(file, proc, arrayName, lengthLine+1, accessLine) {
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
	return assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), arrayName)
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
	for _, dominator := range proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[targetBlock.ID] {
		if dominator == sourceBlock.ID {
			return true
		}
	}
	return false
}

// arrayVBA227FilterForBodyIndexFindings removes only the unallocated/empty
// index observations for a loop body whose For bound necessarily succeeded
// before the body could run. Source-line CFG blocks include a For header and
// its nested body in one scan, so the edge refinement is applied too late for
// the first body visit; this narrow filter preserves the bound finding itself
// and any known lower/upper-bound violation.
func arrayVBA227FilterForBodyIndexFindings(findings []Finding, file parsedFile, proc sourceProcedure, line int, state arrayFlowState, variables map[string]arrayVariable, ctx analysisContext, resumeNextBefore []bool) []Finding {
	if line <= 0 || arrayVBA227ResumeNextBeforeLine(resumeNextBefore, line) {
		return findings
	}
	proven := map[string]bool{}
	provenNonEmpty := map[string]bool{}
	for statement := range proc.Statements.All() {
		if line <= statement.Range.StartLine || line >= statement.Range.EndLine {
			continue
		}
		switch statement.Kind {
		case procedureir.StatementFor:
			if name, ok := arrayVBA227DerivedZeroBasedLoopArray(file, proc, statement, variables, ctx); ok {
				proven[name] = true
				provenNonEmpty[name] = true
			}
		case procedureir.StatementDo, procedureir.StatementWhile:
			if name, ok := arrayVBA227DerivedDoWhileArray(file, proc, statement, line, variables, ctx); ok {
				proven[name] = true
				provenNonEmpty[name] = true
			}
		}
		header := strings.TrimSpace(statement.Text)
		if newline := strings.IndexAny(header, "\r\n"); newline >= 0 {
			header = strings.TrimSpace(header[:newline])
		}
		header = strings.TrimSpace(normalizedCodeLine(header))
		argument := ""
		if _, _, countSource, _, ok := arrayForCountHeader(header); ok {
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
		if match := arrayForScalarBoundRe.FindStringSubmatch(header); len(match) == 2 {
			bound, known := state[strings.ToLower(cleanIdentifier(match[1]))]
			if known {
				argument = bound.safeBoundProbe
			}
		}
		name := strings.ToLower(cleanIdentifier(argument))
		variable, known := variables[name]
		if name != "" && known && (variable.isArray || variable.isVariant) {
			proven[name] = true
			provenNonEmpty[name] = true
		}
		if match := arrayForUBoundRe.FindStringSubmatch(header); len(match) == 3 {
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
// body of a matching positive-length guard. With blocks can keep the If header
// and its first body statement in one CFG block, so the edge refinement runs
// after the body has already been visited. The filter is limited to the
// conditional ByRef allocation contract; unrelated conditions and Else bodies
// remain conservative.
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
					if name, ok := arrayVBA227AllocationProbeLengthArray(state, lengthName, variables); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveSafeArrayLengthArray(file, proc, &access, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveConditionalReDimArray(file, proc, &access, lengthName, variables); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
				}
			}
			inAlternative = true
		case procedureir.StatementIf:
			if !inAlternative {
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
					if name, ok := arrayVBA227AllocationProbeLengthArray(state, lengthName, variables); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveSafeArrayLengthArray(file, proc, &access, lengthName, variables, ctx); ok {
						proven[name] = true
						provenNonEmpty[name] = true
					}
					if name, ok := arrayVBA227PositiveConditionalReDimArray(file, proc, &access, lengthName, variables); ok {
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
			if lengthName, positive := arrayVBA227PositiveLengthCondition(condition); positive {
				if name, probe := arrayVBA227AllocationProbeLengthArray(state, lengthName, variables); probe {
					proven[name] = true
					provenNonEmpty[name] = true
				}
				if name, probe := arrayVBA227PositiveSafeArrayLengthArray(file, proc, &access, lengthName, variables, ctx); probe {
					proven[name] = true
					provenNonEmpty[name] = true
				}
				if name, probe := arrayVBA227PositiveConditionalReDimArray(file, proc, &access, lengthName, variables); probe {
					proven[name] = true
					provenNonEmpty[name] = true
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
			}
		}
		if !remove {
			filtered = append(filtered, finding)
		}
	}
	return filtered
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
	for _, argument := range use.args {
		for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(argument, -1) {
			if strings.EqualFold(cleanIdentifier(bound[2]), use.name) {
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
	for line := proc.StartLine; line < lengthLine && line <= len(file.Lines); line++ {
		for _, source := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			lhs, rhs, indexed, assigned := arrayAssignment(source)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), lengthName) {
				continue
			}
			if candidate, ok := arrayVBA227SafeArrayLengthSource(rhs, ctx); ok {
				if probeLine != 0 && !strings.EqualFold(candidate, arrayName) {
					return "", false
				}
				arrayName = candidate
				probeLine = line
				continue
			}
			if value, ok := integerLiteral(rhs); !ok || value != 0 {
				return "", false
			}
		}
	}
	if probeLine == 0 || arrayName == "" {
		return "", false
	}
	arrayValue, known := variables[arrayName]
	if !known || !arrayValue.isArray || !isByteArrayVariable(arrayValue) {
		return "", false
	}
	sourceLine, ok := arrayVBA227ZeroBasedArraySourceLine(file, proc, arrayName, probeLine)
	if !ok || sourceLine == 0 || probeLine <= sourceLine || !arrayVBA227NoArrayMutationBetween(file, proc, arrayName, sourceLine+1, access.Range.StartLine) {
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
func arrayVBA227PositiveConditionalReDimArray(file parsedFile, proc sourceProcedure, access *procedureir.Statement, lengthName string, variables map[string]arrayVariable) (string, bool) {
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
	if arrayName == "" || redimLine == 0 || redimLine >= positiveLine || !arrayVBA227NoArrayMutationBetween(file, proc, arrayName, redimLine+1, access.Range.StartLine) {
		return "", false
	}
	return arrayName, true
}

func arrayVBA227AllocationProbeLengthArray(state arrayFlowState, lengthName string, variables map[string]arrayVariable) (string, bool) {
	value, known := state[strings.ToLower(cleanIdentifier(lengthName))]
	if !known || value.allocationProbe == "" {
		return "", false
	}
	arrayName := strings.ToLower(cleanIdentifier(value.allocationProbe))
	variable, known := variables[arrayName]
	if !known || !variable.isArray {
		return "", false
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
	for _, call := range arrayCallsAtLine(proc.Calls, statement.Range.StartLine) {
		if call.StatementID != 0 && statement.ID != 0 && call.StatementID != statement.ID {
			continue
		}
		if arrayCallPassesDirectArrayArgument(proc, call, arrayName) {
			return true
		}
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
		if !known || !variable.isArray {
			continue
		}
		value, known := state[name]
		if !known {
			continue
		}
		if updated == nil {
			updated = cloneArrayState(state)
		}
		value = arrayVBA227RecordBoundsProof(value, loopEndLine)
		value.kind = arrayAllocated
		value.knownArray = true
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
		value.allocationCountSource = ""
		updated[name] = value
	}
	if updated == nil {
		return state
	}
	return updated
}

func arrayVBA227RecordBoundsProof(value arrayValue, loopEndLine int) arrayValue {
	if loopEndLine == 0 || value.boundsProof.loopEndLine != 0 || value.kind == arrayAllocated && value.knownArray {
		return value
	}
	value.boundsProof = arrayBoundsProof{
		loopEndLine:                      loopEndLine,
		priorKind:                        value.kind,
		priorKnownArray:                  value.knownArray,
		priorMayBeEmpty:                  value.mayBeEmpty,
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
// Resume Next" fact once per procedure. A reset is intentionally not modeled
// here because the array worklist does not carry VBA's procedure-level error
// mode and a reset may be reachable only on one branch.
func arrayVBA227ResumeNextPrefixes(file parsedFile, proc sourceProcedure) []bool {
	prefixes := make([]bool, len(file.Lines)+1)
	mayHaveResumeNext := false
	start := max(1, proc.StartLine)
	end := min(len(file.Lines), proc.EndLine)
	for line := start; line <= end; line++ {
		prefixes[line] = mayHaveResumeNext
		for _, statement := range splitRangeValueSourceStatements(normalizedCodeLine(file.Lines[line-1])) {
			if arrayOnErrorResumeNextStatementRe.MatchString(strings.TrimSpace(statement)) {
				mayHaveResumeNext = true
				break
			}
		}
	}
	return prefixes
}

func arrayVBA227ResumeNextBeforeLine(prefixes []bool, line int) bool {
	return line >= 0 && line < len(prefixes) && prefixes[line]
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
		if !arrayVBA227SelectCaseStraightLine(text) || len(arrayCallsAtLine(proc.Calls, line)) > 0 {
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
