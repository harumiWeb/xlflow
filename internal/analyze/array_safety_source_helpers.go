package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func arrayLogicalCodeLine(lines []string, line int) string {
	if line < 1 || line > len(lines) {
		return ""
	}
	logical := ""
	for index := line; index <= len(lines); index++ {
		raw := lines[index-1]
		part := strings.TrimSpace(normalizedCodeLine(raw))
		if strings.HasSuffix(part, "_") {
			part = strings.TrimSpace(strings.TrimSuffix(part, "_"))
		}
		if part != "" {
			if logical != "" {
				logical += " "
			}
			logical += part
		}
		if !vbaLineContinues(raw) {
			break
		}
	}
	return logical
}

func arrayLogicalSourceLine(lines []string, line int) string {
	if line < 1 || line > len(lines) {
		return ""
	}
	logical := ""
	for index := line; index <= len(lines); index++ {
		part := strings.TrimSpace(arraySourceOrderStripComment(lines[index-1]))
		if strings.HasSuffix(part, "_") {
			part = strings.TrimSpace(strings.TrimSuffix(part, "_"))
		}
		if part != "" {
			if logical != "" {
				logical += " "
			}
			logical += part
		}
		if !vbaLineContinues(lines[index-1]) {
			break
		}
	}
	return logical
}

func inlineArrayRedimText(text string) (string, bool) {
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return "", false
	}
	prefix := strings.TrimSpace(strings.ToLower(text[:colon]))
	if !strings.HasPrefix(prefix, "dim ") && !strings.HasPrefix(prefix, "static ") {
		return "", false
	}
	redim := strings.TrimSpace(text[colon+1:])
	if next := strings.IndexByte(redim, ':'); next >= 0 {
		redim = strings.TrimSpace(redim[:next])
	}
	if !strings.HasPrefix(strings.ToLower(redim), "redim ") {
		return "", false
	}
	return redim, true
}

func inlineArrayFactoryAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	switch arrayCallName(rhs) {
	case "array", "split", "filter":
		return remainder, true
	default:
		return "", false
	}
}

func inlineArrayDeclarationRemainder(text string) (string, bool) {
	colon := strings.IndexByte(text, ':')
	if colon < 0 {
		return "", false
	}
	prefix := strings.TrimSpace(strings.ToLower(text[:colon]))
	if !strings.HasPrefix(prefix, "dim ") && !strings.HasPrefix(prefix, "static ") {
		return "", false
	}
	remainder := strings.TrimSpace(text[colon+1:])
	if remainder == "" {
		return "", false
	}
	return remainder, true
}

func inlineArrayAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, _, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	return remainder, true
}

func inlineArrayStrConvAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed || !strings.EqualFold(arrayCallName(rhs), "strconv") {
		return "", false
	}
	return remainder, true
}

func inlineArraySafeBoundAssignmentText(text string, guards map[string]bool) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed || !guards[arrayCallName(rhs)] {
		return "", false
	}
	return remainder, true
}

func inlineArrayDictionaryAssignmentText(text string) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	_, _, ok = arrayDictionaryMemberParts(rhs)
	if !ok {
		return "", false
	}
	return remainder, true
}

func inlineArrayReturnAssignmentText(text string, returns map[string]arrayValue) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	_, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	value, known := returns[arrayCallName(rhs)]
	if !known || value.kind != arrayAllocated || !value.knownArray {
		return "", false
	}
	return remainder, true
}

func inlineArrayQualifiedReturnAssignmentText(file parsedFile, proc sourceProcedure, line int, text string, returns map[string]arrayValue) (string, bool) {
	remainder, ok := inlineArrayDeclarationRemainder(text)
	if !ok {
		return "", false
	}
	lhs, rhs, indexed, assigned := arrayAssignment(remainder)
	if !assigned || indexed {
		return "", false
	}
	receiver, member, ok := arrayMemberCallParts(rhs)
	if !ok {
		return "", false
	}
	typeName := arrayTypeNameCaseAtLine(file, proc, line, receiver)
	if typeName == "" {
		return "", false
	}
	value, known := returns[strings.ToLower(typeName+"."+member)]
	if !known || value.kind != arrayAllocated || !value.knownArray {
		return "", false
	}
	// The qualified summary proves only array allocation here. Replace the
	// member call with a recognized array factory so the ordinary transfer can
	// carry that fact without guessing the returned shape.
	return lhs + " = Array()", true
}

func arrayQualifiedReturnExpressionState(proc sourceProcedure, line int, rhs string, variables map[string]arrayVariable, ctx analysisContext) (arrayValue, bool) {
	member := arrayCallName(rhs)
	if member == "" || len(ctx.arrayReturnsQualified) == 0 {
		return arrayValue{}, false
	}
	var value arrayValue
	found := false
	matched := false
	for call := range proc.Calls.All() {
		if call.Range.StartLine != line || call.Callee.Receiver == nil || !strings.EqualFold(cleanIdentifier(call.Callee.Member), member) {
			continue
		}
		resolution := call.Resolution
		if ctx.procedureResolver != nil {
			resolution = ctx.procedureResolver.ResolveCall(call)
		}
		if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
			continue
		}
		matched = true
		key := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
		candidate, known := arrayQualifiedReturnValueForQualifiedName(key, ctx)
		if !known || candidate.kind != arrayAllocated || !candidate.knownArray {
			return arrayValue{}, false
		}
		if found {
			return arrayValue{}, false
		}
		value = candidate
		found = true
	}
	if found || matched {
		return value, found
	}
	return arrayQualifiedReturnValueFromType(rhs, variables, ctx)
}

func arrayQualifiedReturnValueFromType(rhs string, variables map[string]arrayVariable, ctx analysisContext) (arrayValue, bool) {
	receiver, member, ok := arrayQualifiedReturnMemberCallParts(rhs)
	if !ok {
		return arrayValue{}, false
	}
	typeName, ok := arrayQualifiedReturnObjectType(receiver, variables, ctx)
	if !ok {
		return arrayValue{}, false
	}
	value, known := arrayQualifiedReturnValueForType(typeName, member, ctx)
	if !known || value.kind != arrayAllocated || !value.knownArray {
		return arrayValue{}, false
	}
	return value, true
}

func arrayQualifiedReturnObjectType(expression string, variables map[string]arrayVariable, ctx analysisContext) (string, bool) {
	expression = strings.TrimSpace(expression)
	if expression == "" {
		return "", false
	}
	if variable, known := variables[strings.ToLower(cleanIdentifier(expression))]; known && strings.TrimSpace(variable.typ) != "" {
		return strings.TrimSpace(variable.typ), true
	}
	receiver, member, ok := arrayQualifiedReturnMemberCallParts(expression)
	if !ok {
		return "", false
	}
	receiverType, ok := arrayQualifiedReturnObjectType(receiver, variables, ctx)
	if !ok || len(ctx.functionReturnsQualified) == 0 {
		return "", false
	}
	returnType, known := arrayQualifiedReturnTypeForType(receiverType, member, ctx)
	if !known || strings.TrimSpace(returnType) == "" {
		return "", false
	}
	return strings.TrimSpace(returnType), true
}

func arrayQualifiedReturnLookupKeys(typeName string, projectObjectTypes map[string]bool) []string {
	typeName = strings.ToLower(cleanIdentifier(strings.TrimSpace(typeName)))
	if typeName == "" {
		return nil
	}
	keys := []string{typeName}
	short := strings.ToLower(cleanIdentifier(lastName(typeName)))
	if short != "" && short != typeName && projectObjectTypes[typeName] {
		keys = append(keys, short)
	}
	return keys
}

func arrayQualifiedReturnValueForType(typeName, member string, ctx analysisContext) (arrayValue, bool) {
	for _, typeKey := range arrayQualifiedReturnLookupKeys(typeName, ctx.projectObjectTypes) {
		key := typeKey + "." + strings.ToLower(cleanIdentifier(member))
		if value, known := ctx.arrayReturnsQualified[key]; known {
			return value, true
		}
	}
	return arrayValue{}, false
}

func arrayQualifiedReturnTypeForType(typeName, member string, ctx analysisContext) (string, bool) {
	for _, typeKey := range arrayQualifiedReturnLookupKeys(typeName, ctx.projectObjectTypes) {
		key := typeKey + "." + strings.ToLower(cleanIdentifier(member))
		if returnType, known := ctx.functionReturnsQualified[key]; known {
			return returnType, true
		}
	}
	return "", false
}

func arrayQualifiedReturnValueForQualifiedName(name string, ctx analysisContext) (arrayValue, bool) {
	name = strings.ToLower(strings.TrimSpace(name))
	if value, known := ctx.arrayReturnsQualified[name]; known {
		return value, true
	}
	dot := strings.LastIndexByte(name, '.')
	if dot <= 0 || dot >= len(name)-1 {
		return arrayValue{}, false
	}
	return arrayQualifiedReturnValueForType(name[:dot], name[dot+1:], ctx)
}

func arrayQualifiedReturnMemberCallParts(text string) (receiver, member string, ok bool) {
	trimmed := strings.TrimSpace(text)
	dot := arrayQualifiedReturnLastDot(trimmed)
	if dot <= 0 || dot >= len(trimmed)-1 {
		return "", "", false
	}
	receiver = strings.TrimSpace(trimmed[:dot])
	memberText := strings.TrimSpace(trimmed[dot+1:])
	if open := firstParenOutsideString(memberText); open >= 0 {
		close := matchingParen(memberText, open)
		if close < 0 || strings.TrimSpace(memberText[close+1:]) != "" {
			return "", "", false
		}
		memberText = strings.TrimSpace(memberText[:open])
	}
	member = cleanIdentifier(memberText)
	if receiver == "" || !arrayEraseNameRe.MatchString(member) {
		return "", "", false
	}
	return receiver, member, true
}

func arrayQualifiedReturnLastDot(text string) int {
	depth := 0
	last := -1
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
		case '.':
			if !inString && depth == 0 {
				last = index
			}
		}
	}
	return last
}

func arrayMemberCallParts(text string) (receiver, member string, ok bool) {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		close := matchingParen(trimmed, open)
		if close < 0 || strings.TrimSpace(trimmed[close+1:]) != "" {
			return "", "", false
		}
		trimmed = strings.TrimSpace(trimmed[:open])
	}
	dot := strings.LastIndexByte(trimmed, '.')
	if dot <= 0 || dot >= len(trimmed)-1 {
		return "", "", false
	}
	receiver = cleanIdentifier(strings.TrimSpace(trimmed[:dot]))
	member = cleanIdentifier(strings.TrimSpace(trimmed[dot+1:]))
	if !arrayEraseNameRe.MatchString(receiver) || !arrayEraseNameRe.MatchString(member) {
		return "", "", false
	}
	return receiver, member, true
}

func arrayDictionaryMemberParts(text string) (receiver, member string, ok bool) {
	trimmed := strings.TrimSpace(text)
	if open := firstParenOutsideString(trimmed); open >= 0 {
		close := matchingParen(trimmed, open)
		if close < 0 || strings.TrimSpace(trimmed[close+1:]) != "" {
			return "", "", false
		}
		trimmed = strings.TrimSpace(trimmed[:open])
	}
	dot := strings.LastIndexByte(trimmed, '.')
	if dot < 0 || dot >= len(trimmed)-1 {
		return "", "", false
	}
	receiver = strings.TrimSpace(trimmed[:dot])
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(trimmed[dot+1:])))
	if member != "keys" && member != "items" {
		return "", "", false
	}
	if receiver == "" {
		if !strings.HasPrefix(trimmed, ".") {
			return "", "", false
		}
		return "", member, true
	}
	for _, part := range strings.Split(receiver, ".") {
		if !arrayEraseNameRe.MatchString(strings.TrimSpace(part)) {
			return "", "", false
		}
	}
	return receiver, member, true
}

func arrayDictionaryMemberExpressionState(file parsedFile, proc sourceProcedure, line int, rhs string, variables map[string]arrayVariable, ctx analysisContext) (arrayValue, bool) {
	if receiver, key, ok := objectContainerDictionaryItemExpression(rhs); ok {
		if value, safe := arrayDictionaryItemContainerState(file, proc, line, receiver, key, ctx); safe {
			return value, true
		}
	}
	receiver, _, ok := arrayDictionaryMemberParts(rhs)
	knownNonEmpty := false
	if source := arrayLogicalSourceLine(file.Lines, line); source != "" {
		if _, sourceRHS, indexed, assigned := arrayAssignment(source); assigned && !indexed {
			knownNonEmpty = arrayDictionaryMemberKnownNonEmpty(file, line, sourceRHS)
		}
	}
	if !knownNonEmpty {
		knownNonEmpty = arrayDictionaryMemberKnownNonEmpty(file, line, rhs)
	}
	if !knownNonEmpty {
		knownNonEmpty = arrayDictionaryObjectReceiverKnownNonEmpty(file, proc, line, receiver, ctx)
	}
	if !ok {
		if !knownNonEmpty {
			return arrayValue{}, false
		}
		receiver, _, _, _ = arrayDictionarySnapshotParts(rhs)
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, proc, line)
	}
	if receiver == "" {
		return arrayValue{}, false
	}
	if !knownNonEmpty && !arrayDictionaryReceiverProven(file, proc, line, receiver, variables) &&
		!arrayDictionaryObjectReceiverProven(file, proc, line, receiver, ctx) {
		// A late-bound Keys/Items result is still an array-shaped value, but
		// its receiver may be Nothing or empty. Keep the unknown array visible
		// to the bound checker so UBound/LBound remain conservative until a
		// dictionary proof or a positive-count guard establishes safety.
		return arrayValue{
			kind:       arrayUnknown,
			knownArray: true,
			mayBeEmpty: true,
			origin:     arrayOriginUnknown,
		}, true
	}
	source := canonicalArrayBoundExpression(receiver)
	kind := arrayUnknown
	if knownNonEmpty {
		kind = arrayAllocated
	}
	return arrayValue{
		kind:                  kind,
		knownArray:            true,
		mayBeEmpty:            !knownNonEmpty,
		origin:                arrayOriginLocal,
		allocationCountSource: arrayDictionaryCountSourcePrefix + source,
	}, true
}

func arrayDictionaryItemContainerState(file parsedFile, proc sourceProcedure, line int, receiver, key string, ctx analysisContext) (arrayValue, bool) {
	if !strings.EqualFold(strings.TrimSpace(proc.Visibility), "private") {
		return arrayValue{}, false
	}
	index, owner := arrayObjectContainerIndex(ctx, file, proc)
	if index == nil {
		return arrayValue{}, false
	}
	statementID := arrayStatementIDAtLine(owner, line)
	if statementID <= 0 {
		return arrayValue{}, false
	}
	if !arrayDictionaryItemAllocationContract(index, owner, receiver, key, statementID, ctx) {
		return arrayValue{}, false
	}
	return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}, true
}

func arrayDictionaryItemAllocationContract(index *objectContainerIndex, owner sourceProcedure, receiver, key string, observationID int, ctx analysisContext) bool {
	if index == nil || receiver == "" || key == "" || observationID <= 0 {
		return false
	}
	declarations := objectFlowDeclarations(index.file, owner, index.moduleDecls)
	parameter, scope, ok := objectDeclarationBinding(receiver, declarations)
	if !ok || scope != procedureir.ScopeParameter || !parameter.Object || !strings.EqualFold(strings.TrimSpace(owner.Visibility), "private") {
		return false
	}
	parameterIndex := -1
	for candidateIndex, candidate := range owner.Params.AllIndexed() {
		if strings.EqualFold(cleanIdentifier(candidate.Name), receiver) {
			parameterIndex = candidateIndex
			break
		}
	}
	if parameterIndex < 0 {
		return false
	}
	foundCaller := false
	for _, caller := range index.procedures {
		for call := range caller.Calls.All() {
			if call.Callee.Receiver != nil || !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), owner.Name) ||
				!objectContainerCallResolutionUsable(call) || !objectContainerResolvedCandidateMatches(owner, call.Resolution.Candidates[0]) {
				continue
			}
			parameterNames := make([]string, 0, owner.Params.Len())
			for _, parameter := range owner.Params.AllIndexed() {
				parameterNames = append(parameterNames, parameter.Name)
			}
			arguments, argumentsOK := objectContainerCallArgumentsForParameters(caller, call, parameterNames)
			if !argumentsOK || parameterIndex >= len(arguments) {
				return false
			}
			actual := cleanIdentifier(arguments[parameterIndex])
			if actual == "" || strings.ContainsAny(actual, ".()") {
				return false
			}
			callerDeclarations := objectFlowDeclarations(index.file, caller, index.moduleDecls)
			actualDeclaration, actualScope, actualOK := objectDeclarationBinding(actual, callerDeclarations)
			if !actualOK || !actualDeclaration.Object || actualScope != procedureir.ScopeModule || !objectContainerModuleReceiverIsPrivate(index, actual) ||
				!objectContainerDictionaryReceiverKnown(index, caller, actual, actualScope, call.StatementID) {
				return false
			}
			existsGuarded := objectContainerExistsGuarded(index, caller, actual, key, call.StatementID)
			allowRemovals := !arrayDictionaryItemHasRemovalBeforeObservation(index, caller, actual, key, call.StatementID, !existsGuarded, ctx)
			if !arrayDictionaryItemAllWritesSafe(index, actual, key, ctx, allowRemovals) ||
				!arrayDictionaryItemCallerHasPresence(index, caller, actual, key, call.StatementID, ctx, allowRemovals) {
				return false
			}
			foundCaller = true
		}
	}
	return foundCaller
}

func arrayDictionaryItemAllWritesSafe(index *objectContainerIndex, receiver, key string, ctx analysisContext, allowRemovals bool) bool {
	foundArrayWrite := false
	for _, proc := range index.procedures {
		safeWrites, hasMutation, ok := arrayDictionaryItemSafeWrites(index, proc, receiver, key, ctx, allowRemovals)
		if !hasMutation {
			continue
		}
		if !ok {
			return false
		}
		if len(safeWrites) == 0 {
			continue
		}
		if !objectContainerNormalExitCoveredByWrites(proc, safeWrites) {
			return false
		}
		foundArrayWrite = true
	}
	return foundArrayWrite
}

func arrayDictionaryItemSafeWrites(index *objectContainerIndex, proc sourceProcedure, receiver, key string, ctx analysisContext, allowRemovals bool) (map[int]bool, bool, bool) {
	safeWrites := map[int]bool{}
	hasMutation := false
	declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
	declaration, scope, bound := objectDeclarationBinding(receiver, declarations)
	moduleDeclaration, moduleBound := index.moduleDecls[strings.ToLower(cleanIdentifier(receiver))]
	if !bound || !moduleBound || scope != procedureir.ScopeModule || declaration.Line != moduleDeclaration.Line {
		if objectContainerProcedureMayMutateReceiver(index, proc, receiver, key) {
			return safeWrites, true, false
		}
		return safeWrites, false, true
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		if !targetOK || !strings.EqualFold(target, receiver) {
			continue
		}
		if strings.EqualFold(proc.Name, "Class_Initialize") && objectConstructorExpression(statement.Value) {
			continue
		}
		return safeWrites, true, false
	}
	flowContext, _, contextOK := objectContainerGraphContext(index, proc)
	if !contextOK {
		return safeWrites, false, false
	}
	for call := range proc.Calls.All() {
		if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) {
			continue
		}
		member := strings.ToLower(cleanIdentifier(call.Callee.Member))
		switch member {
		case "exists", "count", "keys":
			continue
		case "add":
			callKey, keyOK := objectContainerDictionaryKey(proc, call)
			if !keyOK {
				return safeWrites, true, false
			}
			if !objectContainerKeysEqual(callKey, key) {
				continue
			}
			hasMutation = true
			if objectErrorResumeNextAt(proc, call.StatementID) {
				return safeWrites, hasMutation, false
			}
			arguments := objectContainerCallArguments(proc, call)
			if len(arguments) < 2 || !arrayDictionaryItemNonEmptyArrayExpression(index, proc, call.StatementID, arguments[1], receiver, key, flowContext, ctx) {
				return safeWrites, hasMutation, false
			}
			safeWrites[call.StatementID] = true
		case "item":
			statement, statementOK := objectContainerStatement(proc, call.StatementID)
			if !statementOK {
				return safeWrites, true, false
			}
			matched, assignment := objectContainerDictionaryItemTarget(statement, receiver, key)
			if !assignment {
				continue
			}
			hasMutation = true
			if !matched || !arrayDictionaryItemExactAssignment(statement, receiver, key) || objectErrorResumeNextAt(proc, call.StatementID) {
				return safeWrites, hasMutation, false
			}
			_, rhs, assigned := arrayAssignmentSides(statement.Text)
			if !assigned || !arrayDictionaryItemNonEmptyArrayExpression(index, proc, call.StatementID, rhs, receiver, key, flowContext, ctx) {
				return safeWrites, hasMutation, false
			}
			safeWrites[call.StatementID] = true
		case "remove":
			callKey, keyOK := objectContainerDictionaryKey(proc, call)
			if !keyOK {
				return safeWrites, true, false
			}
			if objectContainerKeysMayAlias(callKey, key) {
				hasMutation = true
				if !allowRemovals {
					return safeWrites, hasMutation, false
				}
			}
		case "removeall":
			hasMutation = true
			if !allowRemovals {
				return safeWrites, hasMutation, false
			}
		default:
			return safeWrites, true, false
		}
	}
	return safeWrites, hasMutation, true
}

func arrayDictionaryItemCallerHasPresence(index *objectContainerIndex, caller sourceProcedure, receiver, key string, observationID int, ctx analysisContext, allowRemovals bool) bool {
	if objectContainerExistsGuarded(index, caller, receiver, key, observationID) {
		return true
	}
	for _, proc := range index.procedures {
		safeWrites, hasMutation, ok := arrayDictionaryItemSafeWrites(index, proc, receiver, key, ctx, allowRemovals)
		if !hasMutation || !ok || len(safeWrites) == 0 || !objectContainerNormalExitCoveredByWrites(proc, safeWrites) {
			continue
		}
		if proc.StartLine == caller.StartLine {
			if objectContainerObservationCoveredByWrites(caller, observationID, safeWrites) {
				return true
			}
			continue
		}
		if objectContainerProcedureCalledBeforeObservation(index, caller, proc, observationID) {
			return true
		}
	}
	return false
}

func arrayDictionaryItemProcedureHasRemoval(index *objectContainerIndex, proc sourceProcedure, receiver, key string) bool {
	if index == nil {
		return false
	}
	for call := range proc.Calls.All() {
		if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) {
			continue
		}
		switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
		case "removeall":
			return true
		case "remove":
			callKey, keyOK := objectContainerDictionaryKey(proc, call)
			if !keyOK || objectContainerKeysMayAlias(callKey, key) {
				return true
			}
		}
	}
	return false
}

func arrayDictionaryItemProcedureHasRemovalBeforeObservation(index *objectContainerIndex, proc sourceProcedure, receiver, key string, observationID int, ctx analysisContext) bool {
	flowContext, _, contextOK := objectContainerGraphContext(index, proc)
	if !contextOK {
		return true
	}
	observationBlock, observationOK := flowContext.graph.BlockForStatement(observationID)
	if !observationOK {
		return true
	}
	safeWrites, _, safe := arrayDictionaryItemSafeWrites(index, proc, receiver, key, ctx, true)
	if !safe {
		return true
	}
	for call := range proc.Calls.All() {
		if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) {
			continue
		}
		member := strings.ToLower(cleanIdentifier(call.Callee.Member))
		if member != "remove" && member != "removeall" {
			continue
		}
		if member == "remove" {
			callKey, keyOK := objectContainerDictionaryKey(proc, call)
			if keyOK && !objectContainerKeysMayAlias(callKey, key) {
				continue
			}
		}
		if !objectContainerStatementBeforeObservation(flowContext.graph, call.StatementID, observationID) {
			continue
		}
		overwritten := objectContainerPathCoveredByWrites(proc, call.StatementID, observationBlock.ID, observationID, safeWrites)
		if !overwritten {
			return true
		}
	}
	return false
}

func arrayDictionaryItemProcedureRemovalInvalidatesNormalExit(index *objectContainerIndex, proc sourceProcedure, receiver, key string, ctx analysisContext) bool {
	flowContext, _, contextOK := objectContainerGraphContext(index, proc)
	if !contextOK {
		return true
	}
	safeWrites, _, safe := arrayDictionaryItemSafeWrites(index, proc, receiver, key, ctx, true)
	if !safe {
		return true
	}
	for call := range proc.Calls.All() {
		if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) {
			continue
		}
		member := strings.ToLower(cleanIdentifier(call.Callee.Member))
		if member != "remove" && member != "removeall" {
			continue
		}
		if member == "remove" {
			callKey, keyOK := objectContainerDictionaryKey(proc, call)
			if keyOK && !objectContainerKeysMayAlias(callKey, key) {
				continue
			}
		}
		removalBlock, removalOK := flowContext.graph.BlockForStatement(call.StatementID)
		if !removalOK || !objectContainerBlockCanReach(flowContext.graph, removalBlock.ID, flowContext.graph.NormalExit()) {
			continue
		}
		overwritten := objectContainerPathCoveredByWrites(proc, call.StatementID, flowContext.graph.NormalExit(), 0, safeWrites)
		if !overwritten {
			return true
		}
	}
	return false
}

func arrayDictionaryItemHasRemovalBeforeObservation(index *objectContainerIndex, caller sourceProcedure, receiver, key string, observationID int, includeExternal bool, ctx analysisContext) bool {
	if index == nil || observationID <= 0 {
		return true
	}
	for _, proc := range index.procedures {
		if !arrayDictionaryItemProcedureHasRemoval(index, proc, receiver, key) {
			continue
		}
		if proc.StartLine == caller.StartLine {
			if arrayDictionaryItemProcedureHasRemovalBeforeObservation(index, proc, receiver, key, observationID, ctx) {
				return true
			}
			continue
		}
		if includeExternal && strings.EqualFold(strings.TrimSpace(proc.Visibility), "public") {
			return true
		}
		if objectContainerProcedureMayReachObservation(index, caller, proc, observationID) &&
			arrayDictionaryItemProcedureRemovalInvalidatesNormalExit(index, proc, receiver, key, ctx) {
			return true
		}
	}
	return false
}

func arrayDictionaryItemNonEmptyArrayExpression(index *objectContainerIndex, proc sourceProcedure, statementID int, expression, receiver, key string, flowContext objectFlowContext, ctx analysisContext) bool {
	name := arrayCallName(expression)
	if name == "array" {
		open := firstParenOutsideString(strings.TrimSpace(expression))
		if open < 0 {
			return false
		}
		close := matchingParen(strings.TrimSpace(expression), open)
		return close == len(strings.TrimSpace(expression))-1 && len(splitArgs(strings.TrimSpace(expression)[open+1:close])) > 0
	}
	for call := range proc.Calls.All() {
		if call.StatementID != statementID || call.Callee.Receiver != nil || !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), name) {
			continue
		}
		appendInfo, appendOK := objectContainerAppendInfoForResolvedCall(index, proc, call)
		if !appendOK || appendInfo.arrayParameter < 0 {
			continue
		}
		arguments, argumentsOK := objectContainerCallArgumentsForProcedure(index, proc, call, call.Callee.BaseName)
		if !argumentsOK || appendInfo.arrayParameter >= len(arguments) {
			continue
		}
		arrayReceiver, arrayKey, sourceOK := objectContainerArraySource(proc, cleanIdentifier(arguments[appendInfo.arrayParameter]), statementID, flowContext)
		return sourceOK && strings.EqualFold(arrayReceiver, receiver) && objectContainerKeysEqual(arrayKey, key)
	}
	if value, known := ctx.arrayReturns[name]; known {
		return value.kind == arrayAllocated && value.knownArray && !value.mayBeEmpty
	}
	for _, candidate := range index.procedures {
		if !strings.EqualFold(cleanIdentifier(candidate.Name), name) || candidate.ProcedureKind != procedureir.ProcedureFunction && candidate.ProcedureKind != procedureir.ProcedurePropertyGet {
			continue
		}
		if candidate.ReturnValueShape == procedureir.ValueShapeDynamicArray || strings.Contains(strings.ReplaceAll(candidate.ReturnType, " ", ""), "()") {
			return arrayProcedureHasNonEmptyReturnAllocation(index.file, candidate)
		}
	}
	return false
}

func arrayDictionaryItemExactAssignment(statement procedureir.Statement, receiver, key string) bool {
	left, _, assigned := arrayAssignmentSides(statement.Text)
	if !assigned {
		return false
	}
	if leftReceiver, leftKey, literal := objectContainerDictionaryItemExpression(left); literal {
		return strings.EqualFold(leftReceiver, receiver) && objectContainerKeysEqual(leftKey, key)
	}
	if leftReceiver, leftKey, literal := objectContainerDefaultItemExpression(left); literal {
		return strings.EqualFold(leftReceiver, receiver) && objectContainerKeysEqual(leftKey, key)
	}
	return false
}

type arrayObjectContainerIndexCacheEntry struct {
	index      *objectContainerIndex
	procedures []sourceProcedure
}

func arrayObjectContainerIndex(ctx analysisContext, file parsedFile, proc sourceProcedure) (*objectContainerIndex, sourceProcedure) {
	if ctx.objectAnalysis != nil {
		key := objectSummaryKey(file.IR.Path, objectProcedureQualifiedName(proc), string(proc.ProcedureKind), proc.StartLine)
		if plan := ctx.objectAnalysis.plans[key]; plan != nil && plan.containerIndex != nil {
			return plan.containerIndex, plan.proc
		}
	}
	if ctx.procedureResolver != nil {
		cacheKey := file.Path
		if cacheKey == "" {
			cacheKey = file.IR.Path
		}
		if cacheKey != "" && ctx.arrayObjectContainerIndexCache != nil {
			if ctx.arrayObjectContainerIndexCacheMu != nil {
				ctx.arrayObjectContainerIndexCacheMu.RLock()
			}
			cached, cachedOK := ctx.arrayObjectContainerIndexCache[cacheKey]
			if ctx.arrayObjectContainerIndexCacheMu != nil {
				ctx.arrayObjectContainerIndexCacheMu.RUnlock()
			}
			if cachedOK {
				for _, candidate := range cached.procedures {
					if candidate.StartLine == proc.StartLine && strings.EqualFold(candidate.Name, proc.Name) {
						return cached.index, candidate
					}
				}
				return cached.index, proc
			}
		}
		resolvedIR := procedureir.Resolve(file.IR, ctx.procedureResolver)
		resolvedProcedures := sourceProceduresFromIRRef(&resolvedIR, file.CFG)
		resolvedFile := file
		resolvedFile.IR = resolvedIR
		resolvedFile.Procedures = resolvedProcedures
		index := buildObjectContainerIndex(resolvedFile)
		if cacheKey != "" && ctx.arrayObjectContainerIndexCache != nil {
			if ctx.arrayObjectContainerIndexCacheMu != nil {
				ctx.arrayObjectContainerIndexCacheMu.Lock()
			}
			ctx.arrayObjectContainerIndexCache[cacheKey] = arrayObjectContainerIndexCacheEntry{index: index, procedures: resolvedProcedures}
			if ctx.arrayObjectContainerIndexCacheMu != nil {
				ctx.arrayObjectContainerIndexCacheMu.Unlock()
			}
		}
		for _, candidate := range resolvedProcedures {
			if candidate.StartLine == proc.StartLine && strings.EqualFold(candidate.Name, proc.Name) {
				return index, candidate
			}
		}
		return index, proc
	}
	return buildObjectContainerIndex(file), proc
}

func arrayStatementIDAtLine(proc sourceProcedure, line int) int {
	for statement := range proc.Statements.All() {
		if statement.ID > 0 && statement.Range.StartLine == line {
			return statement.ID
		}
	}
	return 0
}

// arrayDictionaryMemberKnownNonEmpty recognizes the outer dictionary returned
// by a helper such as CreateLookupDict. The helper creates fixed members before
// it consumes its input, so Keys and Items on that outer dictionary always
// contain those members even when the input pair array is empty.
func arrayDictionaryMemberKnownNonEmpty(file parsedFile, line int, rhs string) bool {
	receiver, key, _, ok := arrayDictionarySnapshotParts(rhs)
	if !ok {
		return false
	}
	for procedure := range file.procedureView().All() {
		if !strings.EqualFold(procedure.Name, "CreateLookupDict") {
			continue
		}
		if arrayProcedureReturnsNonEmptyObjectMemberSet(file, procedure) &&
			arrayDictionaryMemberAssignmentUsesHelper(file, line, receiver, key) {
			return true
		}
	}
	return false
}

func arrayDictionarySnapshotParts(text string) (receiver, key, member string, ok bool) {
	canonical := canonicalArrayBoundExpression(text)
	lower := strings.ToLower(canonical)
	for _, candidate := range []string{"keys", "items"} {
		suffix := "." + candidate + "()"
		if !strings.HasSuffix(lower, suffix) {
			continue
		}
		prefix := canonical[:len(canonical)-len(suffix)]
		if len(prefix) == 0 || prefix[len(prefix)-1] != ')' {
			continue
		}
		open := -1
		for index := 0; index < len(prefix)-1; index++ {
			if prefix[index] == '(' && matchingParen(prefix, index) == len(prefix)-1 {
				open = index
				break
			}
		}
		if open <= 0 {
			continue
		}
		return strings.TrimSpace(prefix[:open]), strings.TrimSpace(prefix[open+1 : len(prefix)-1]), candidate, true
	}
	return "", "", "", false
}

func arrayProcedureReturnsNonEmptyObjectMemberSet(file parsedFile, procedure sourceProcedure) bool {
	returnedObject := ""
	start := max(0, procedure.StartLine-1)
	end := min(procedure.EndLine, len(file.Lines))
	for line := start; line < end; line++ {
		text := arrayLogicalSourceLine(file.Lines, line+1)
		if text == "" {
			continue
		}
		if lhs, rhs, assigned := arrayAssignmentSides(text); assigned && !strings.Contains(lhs, "(") && strings.EqualFold(cleanIdentifier(lhs), procedure.Name) {
			returnedObject = cleanIdentifier(rhs)
		}
	}
	if returnedObject == "" {
		return false
	}
	members := map[string]bool{}
	for line := start; line < end; line++ {
		text := arrayLogicalSourceLine(file.Lines, line+1)
		if text == "" {
			continue
		}
		base, memberKey, memberRHS, assigned := arrayMemberAssignmentParts(text)
		if !assigned || arrayCallName(memberRHS) != "createobject" {
			continue
		}
		if returnedObject != "" && !strings.EqualFold(cleanIdentifier(base), returnedObject) {
			continue
		}
		members[canonicalArrayBoundExpression(memberKey)] = true
	}
	return len(members) >= 2
}

func arrayDictionaryMemberAssignmentUsesHelper(file parsedFile, line int, receiver, key string) bool {
	wantReceiver := canonicalArrayBoundExpression(receiver)
	wantKey := canonicalArrayBoundExpression(key)
	assignedByHelper := false
	invalidBeforeSnapshot := false
	for sourceLine := 1; sourceLine <= len(file.Lines); sourceLine++ {
		text := arrayLogicalSourceLine(file.Lines, sourceLine)
		base, memberKey, memberRHS, assigned := arrayMemberAssignmentParts(text)
		if !assigned || canonicalArrayBoundExpression(base) != wantReceiver || canonicalArrayBoundExpression(memberKey) != wantKey {
			continue
		}
		if arrayCallName(memberRHS) == "createlookupdict" {
			// The initializing call may be in a different procedure that is
			// textually later in the module.  Keep the summary module-wide, but
			// let an earlier direct reassignment invalidate the fact for this
			// snapshot.
			assignedByHelper = true
		} else if sourceLine < line {
			invalidBeforeSnapshot = true
		}
	}
	return assignedByHelper && !invalidBeforeSnapshot
}

func arrayMemberAssignmentParts(text string) (receiver, key, rhs string, ok bool) {
	lhs, rhs, assigned := arrayAssignmentSides(text)
	if !assigned {
		return "", "", "", false
	}
	open := firstParenOutsideString(lhs)
	if open <= 0 {
		return "", "", "", false
	}
	close := matchingParen(lhs, open)
	if close != len(lhs)-1 {
		return "", "", "", false
	}
	return strings.TrimSpace(lhs[:open]), strings.TrimSpace(lhs[open+1 : close]), rhs, true
}

func arrayAssignmentSides(text string) (lhs, rhs string, ok bool) {
	trimmed := strings.TrimSpace(text)
	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "set ") || strings.HasPrefix(lower, "let ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	inString := false
	for index := 0; index < len(trimmed); index++ {
		switch trimmed[index] {
		case '"':
			if inString && index+1 < len(trimmed) && trimmed[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case '=':
			if inString || index > 0 && (trimmed[index-1] == '<' || trimmed[index-1] == '>' || trimmed[index-1] == '=') {
				continue
			}
			lhs = strings.TrimSpace(trimmed[:index])
			rhs = strings.TrimSpace(trimmed[index+1:])
			return lhs, rhs, lhs != "" && rhs != ""
		}
	}
	return "", "", false
}

func arrayWithReceiverAtLine(file parsedFile, proc sourceProcedure, line int) string {
	stack := make([]string, 0, 2)
	start := max(1, proc.StartLine)
	end := min(line-1, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[current-1]))
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if lower == "end with" {
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
			continue
		}
		if strings.HasPrefix(lower, "with ") {
			receiver := strings.TrimSpace(text[len("with "):])
			if receiver != "" {
				stack = append(stack, receiver)
			}
		}
	}
	if len(stack) == 0 {
		return ""
	}
	return stack[len(stack)-1]
}

func arrayDictionaryReceiverProven(file parsedFile, proc sourceProcedure, line int, receiver string, variables map[string]arrayVariable) bool {
	receiver = strings.TrimSpace(receiver)
	if receiver == "" {
		return false
	}
	if variable, known := variables[strings.ToLower(cleanIdentifier(receiver))]; known && isDictionaryType(variable.typ) {
		return true
	}
	if strings.EqualFold(arrayTypeNameCaseAtLine(file, proc, line, receiver), "Dictionary") {
		return true
	}
	if !strings.EqualFold(receiver, "This.children") || !strings.EqualFold(arraySelectCaseValueAtLine(file, proc, line, "This.iType"), "eJSONObject") {
		return false
	}
	for _, rawLine := range file.Lines {
		text := canonicalArrayBoundExpression(gui.StripComment(rawLine))
		if strings.Contains(text, "setthis.children=createdictionary") {
			return true
		}
	}
	return false
}

func arrayDictionaryObjectReceiverProven(file parsedFile, proc sourceProcedure, line int, receiver string, ctx analysisContext) bool {
	index, owner := arrayObjectContainerIndex(ctx, file, proc)
	if index == nil {
		return false
	}
	statementID := arrayStatementIDAtLine(owner, line)
	if statementID <= 0 {
		return false
	}
	declarations := objectFlowDeclarations(index.file, owner, index.moduleDecls)
	_, scope, ok := objectDeclarationBinding(receiver, declarations)
	if !ok {
		return false
	}
	if objectContainerDictionaryReceiverKnown(index, owner, receiver, scope, statementID) {
		return true
	}
	return arrayDictionaryFactoryReceiverProven(index, owner, receiver, scope, statementID, declarations)
}

func arrayDictionaryFactoryReceiverProven(index *objectContainerIndex, owner sourceProcedure, receiver string, scope procedureir.SymbolScope, statementID int, declarations declarationScope) bool {
	if index == nil || statementID <= 0 {
		return false
	}
	flowContext, _, ok := objectContainerGraphContext(index, owner)
	if !ok {
		return false
	}
	latest, found := objectContainerLastDominatingAssignment(owner, objectVariable{Scope: scope, Name: receiver}, statementID, declarations, flowContext)
	if !found || latest.Value == nil || objectErrorResumeNextAt(owner, latest.ID) {
		return false
	}
	factoryName := arrayCallName(latest.Value.Text)
	if factoryName == "" {
		return false
	}
	var factory sourceProcedure
	foundFactory := false
	for _, candidate := range index.procedures {
		if !strings.EqualFold(cleanIdentifier(candidate.Name), factoryName) {
			continue
		}
		if foundFactory {
			return false
		}
		factory = candidate
		foundFactory = true
	}
	return foundFactory && arrayDictionaryFactoryReturnsDictionary(index, factory)
}

func arrayDictionaryFactoryReturnsDictionary(index *objectContainerIndex, procedure sourceProcedure) bool {
	if index == nil || !isObjectType(procedure.ReturnType) {
		return false
	}
	_, _, ok := objectContainerGraphContext(index, procedure)
	if !ok {
		return false
	}
	returnIDs := map[int]bool{}
	for statement := range procedure.Statements.All() {
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		if !targetOK || !strings.EqualFold(target, procedure.Name) {
			continue
		}
		if statement.Value == nil || !objectContainerDictionaryConstructorExpression(statement.Value) || objectErrorResumeNextAt(procedure, statement.ID) {
			return false
		}
		returnIDs[statement.ID] = true
	}
	return len(returnIDs) > 0 && objectContainerNormalExitCoveredByWrites(procedure, returnIDs)
}

// arrayDictionaryObjectReceiverKnownNonEmpty proves that a dictionary snapshot
// contains an element after a fresh, dominating literal-key Add.  The object
// flow index supplies the constructor proof; this narrow local rule then
// rejects any later reassignment, Remove, RemoveAll, or unknown call that could
// invalidate the count before the snapshot is materialized.
func arrayDictionaryObjectReceiverKnownNonEmpty(file parsedFile, proc sourceProcedure, line int, receiver string, ctx analysisContext) bool {
	receiver = strings.TrimSpace(receiver)
	if receiver == "" || strings.ContainsAny(receiver, ".()") {
		return false
	}
	index, owner := arrayObjectContainerIndex(ctx, file, proc)
	if index == nil {
		return false
	}
	statementID := arrayStatementIDAtLine(owner, line)
	if statementID <= 0 {
		return false
	}
	if owner.Graph == nil {
		return false
	}
	graph := owner.Graph.WithoutNormalErrRaiseContinuationView()
	receiver = cleanIdentifier(receiver)
	flowContext, _, ok := objectContainerGraphContext(index, owner)
	if !ok {
		return false
	}
	declarations := objectFlowDeclarations(index.file, owner, index.moduleDecls)
	_, scope, ok := objectDeclarationBinding(receiver, declarations)
	if !ok {
		return false
	}
	latestAssignment, found := objectContainerLastDominatingAssignment(owner, objectVariable{Scope: scope, Name: receiver}, statementID, declarations, flowContext)
	if !found || latestAssignment.ID <= 0 {
		return false
	}
	addIDs := make([]int, 0, 1)
	for call := range owner.Calls.All() {
		if !strings.EqualFold(objectCallWithReceiverName(owner, call), receiver) ||
			!strings.EqualFold(cleanIdentifier(call.Callee.Member), "add") ||
			call.StatementID <= latestAssignment.ID ||
			!objectContainerStatementDominates(graph, call.StatementID, statementID) ||
			!objectContainerStatementBeforeObservation(graph, call.StatementID, statementID) ||
			objectErrorResumeNextAt(owner, call.StatementID) || objectContainerErrorHandlerActiveAt(owner, call.StatementID) {
			continue
		}
		arguments := objectContainerCallArguments(owner, call)
		if len(arguments) == 0 {
			continue
		}
		if _, literal := objectContainerLiteral(arguments[0]); !literal {
			continue
		}
		addIDs = append(addIDs, call.StatementID)
	}
	if len(addIDs) == 0 {
		return false
	}
	if !arrayDictionaryObjectReceiverProven(file, owner, line, receiver, ctx) {
		return false
	}

	for statement := range owner.Statements.All() {
		if statement.ID == statementID || !objectContainerStatementBeforeObservation(graph, statement.ID, statementID) {
			continue
		}
		afterAdd := false
		for _, addID := range addIDs {
			if statement.ID > addID {
				afterAdd = true
				break
			}
		}
		if !afterAdd || statement.ID >= statementID {
			continue
		}
		if target, targetOK := objectContainerSimpleTarget(statement); targetOK && strings.EqualFold(target, receiver) {
			return false
		}
	}
	for call := range owner.Calls.All() {
		if call.StatementID == statementID || !objectContainerStatementBeforeObservation(graph, call.StatementID, statementID) {
			continue
		}
		afterAdd := false
		for _, addID := range addIDs {
			if call.StatementID > addID {
				afterAdd = true
				break
			}
		}
		if !afterAdd || call.StatementID >= statementID {
			continue
		}
		if !strings.EqualFold(objectCallWithReceiverName(owner, call), receiver) {
			for _, argument := range objectContainerCallArguments(owner, call) {
				if strings.EqualFold(cleanIdentifier(argument), receiver) {
					return false
				}
			}
			continue
		}
		switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
		case "remove", "removeall":
			return false
		case "add", "exists", "count", "keys", "items", "item", "comparemode":
			continue
		default:
			return false
		}
	}
	return true
}

func arraySelectCaseValueAtLine(file parsedFile, proc sourceProcedure, line int, expression string) string {
	type frame struct {
		expression string
		caseValue  string
	}
	frames := make([]frame, 0, 2)
	want := canonicalArrayBoundExpression(expression)
	start := max(1, proc.StartLine)
	end := min(line, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.Join(strings.Fields(gui.StripComment(file.Lines[current-1])), " ")
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "select case ") {
			frames = append(frames, frame{expression: canonicalArrayBoundExpression(text[len("select case "):])})
			continue
		}
		if lower == "end select" {
			if len(frames) > 0 {
				frames = frames[:len(frames)-1]
			}
			continue
		}
		if !strings.HasPrefix(lower, "case ") || len(frames) == 0 {
			continue
		}
		caseText := strings.TrimSpace(text[len("case "):])
		if comma := strings.IndexByte(caseText, ','); comma >= 0 {
			caseText = strings.TrimSpace(caseText[:comma])
		}
		if strings.EqualFold(caseText, "else") {
			caseText = ""
		}
		frames[len(frames)-1].caseValue = caseText
	}
	for index := len(frames) - 1; index >= 0; index-- {
		if frames[index].expression == want {
			return frames[index].caseValue
		}
	}
	return ""
}

func arrayTypeNameCaseAtLine(file parsedFile, proc sourceProcedure, line int, receiver string) string {
	if receiver == "" || line <= 0 || len(file.Lines) == 0 {
		return ""
	}
	type frame struct {
		receiver string
		caseName string
	}
	frames := make([]frame, 0, 2)
	start := max(1, proc.StartLine)
	end := min(line, len(file.Lines))
	for current := start; current <= end; current++ {
		text := strings.Join(strings.Fields(gui.StripComment(file.Lines[current-1])), " ")
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "select case ") {
			expression := strings.TrimSpace(text[len("select case "):])
			selectedReceiver := ""
			if match := arrayTypeNameExpressionRe.FindStringSubmatch(expression); len(match) == 2 {
				selectedReceiver = cleanIdentifier(match[1])
			}
			frames = append(frames, frame{receiver: selectedReceiver})
			continue
		}
		if strings.HasPrefix(lower, "end select") {
			if len(frames) > 0 {
				frames = frames[:len(frames)-1]
			}
			continue
		}
		if !strings.HasPrefix(lower, "case ") || len(frames) == 0 {
			continue
		}
		caseText := strings.TrimSpace(text[len("case "):])
		if match := arrayQuotedCaseRe.FindStringSubmatch(caseText); len(match) == 2 {
			frames[len(frames)-1].caseName = match[1]
		} else {
			frames[len(frames)-1].caseName = ""
		}
	}
	for index := len(frames) - 1; index >= 0; index-- {
		if strings.EqualFold(frames[index].receiver, receiver) {
			return frames[index].caseName
		}
	}
	return ""
}
