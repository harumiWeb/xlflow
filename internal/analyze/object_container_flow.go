package analyze

import (
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// objectContainerIndex contains the immutable procedure view needed to carry
// a narrow container-shape proof across procedures. It intentionally models
// only shapes that are established by source-level writes: a Dictionary key
// receiving an Array of constructed objects, and a constructed object in that
// array receiving a constructed Collection under a literal key.
type objectContainerIndex struct {
	file        parsedFile
	procedures  []sourceProcedure
	moduleDecls map[string]sourceDeclaration
}

type objectContainerElement struct {
	procedureStart     int
	name               string
	storageStatementID int
}

type objectContainerAppendInfo struct {
	arrayParameter      int
	objectParameter     int
	arrayParameterName  string
	objectParameterName string
}

func buildObjectContainerIndex(file parsedFile) *objectContainerIndex {
	return &objectContainerIndex{
		file:        file,
		procedures:  file.procedures(),
		moduleDecls: file.moduleDecls(),
	}
}

func objectContainerCallArguments(proc sourceProcedure, call procedureir.CallSite) []string {
	actuals := objectCallActuals(call, proc.analysisFacts())
	arguments := make([]string, 0, len(actuals))
	for _, actual := range actuals {
		arguments = append(arguments, strings.TrimSpace(actual.text))
	}
	return arguments
}

func objectContainerCallArgumentsForParameters(proc sourceProcedure, call procedureir.CallSite, parameterNames []string) ([]string, bool) {
	actuals := objectContainerCallArguments(proc, call)
	if len(actuals) != len(call.Arguments.ExpressionIDs) {
		return nil, false
	}
	arguments := make([]string, len(parameterNames))
	assigned := make([]bool, len(parameterNames))
	namedIDs := map[int]string{}
	for _, named := range call.Arguments.Named {
		if named.ExpressionID == 0 {
			return nil, false
		}
		namedIDs[named.ExpressionID] = named.Name
	}
	nextPositional := 0
	for actualIndex, actual := range actuals {
		formalIndex := -1
		if name, named := namedIDs[call.Arguments.ExpressionIDs[actualIndex]]; named {
			for parameterIndex, parameterName := range parameterNames {
				if strings.EqualFold(cleanIdentifier(parameterName), cleanIdentifier(name)) {
					formalIndex = parameterIndex
					break
				}
			}
		} else {
			formalIndex = nextPositional
			nextPositional++
		}
		if formalIndex < 0 || formalIndex >= len(arguments) || assigned[formalIndex] {
			return nil, false
		}
		arguments[formalIndex] = actual
		assigned[formalIndex] = true
	}
	return arguments, true
}

func objectContainerCallArgumentsForProcedure(index *objectContainerIndex, proc sourceProcedure, call procedureir.CallSite, name string) ([]string, bool) {
	if index == nil {
		return nil, false
	}
	var callee sourceProcedure
	found := false
	for _, candidate := range index.procedures {
		if !strings.EqualFold(cleanIdentifier(candidate.Name), cleanIdentifier(name)) {
			continue
		}
		if found {
			return nil, false
		}
		callee = candidate
		found = true
	}
	if !found {
		return nil, false
	}
	parameterNames := make([]string, 0, callee.Params.Len())
	for _, parameter := range callee.Params.AllIndexed() {
		parameterNames = append(parameterNames, parameter.Name)
	}
	return objectContainerCallArgumentsForParameters(proc, call, parameterNames)
}

func objectContainerLiteral(text string) (string, bool) {
	text = strings.TrimSpace(text)
	if len(text) < 2 || text[0] != '"' || text[len(text)-1] != '"' {
		return "", false
	}
	value := strings.ReplaceAll(text[1:len(text)-1], `""`, `"`)
	if value == "" {
		return "", false
	}
	return value, true
}

func objectContainerKeysEqual(left, right string) bool {
	return left == right
}

// Scripting.Dictionary key comparison depends on CompareMode. Unless that
// mode is proven, a case-only difference can still refer to the observed key;
// use this predicate when deciding whether a later write invalidates a proof.
func objectContainerKeysMayAlias(left, right string) bool {
	return strings.EqualFold(left, right)
}

func objectContainerDictionaryKey(proc sourceProcedure, call procedureir.CallSite) (string, bool) {
	arguments := objectContainerCallArguments(proc, call)
	if len(arguments) == 0 {
		return "", false
	}
	return objectContainerLiteral(arguments[0])
}

func objectContainerReceiver(call procedureir.CallSite) string {
	if call.Callee.Receiver == nil {
		return ""
	}
	return cleanIdentifier(strings.TrimSpace(*call.Callee.Receiver))
}

func objectContainerStatement(proc sourceProcedure, statementID int) (procedureir.Statement, bool) {
	for statement := range proc.Statements.All() {
		if statement.ID == statementID {
			return statement, true
		}
	}
	return procedureir.Statement{}, false
}

func objectContainerAssignmentTarget(statement procedureir.Statement) string {
	if statement.Target != nil {
		return cleanIdentifier(strings.TrimSpace(statement.Target.Text))
	}
	left, _, ok := strings.Cut(statement.Text, "=")
	if !ok {
		return ""
	}
	left = strings.TrimSpace(left)
	for _, prefix := range []string{"Set ", "Let "} {
		if len(left) >= len(prefix) && strings.EqualFold(left[:len(prefix)], prefix) {
			left = strings.TrimSpace(left[len(prefix):])
			break
		}
	}
	if strings.ContainsAny(left, ".()") {
		return ""
	}
	return cleanIdentifier(left)
}

func objectContainerIndexedTarget(statement procedureir.Statement) (string, bool) {
	name, _, ok := objectContainerIndexedTargetExpression(statement)
	return name, ok
}

func objectContainerIndexedTargetExpression(statement procedureir.Statement) (string, string, bool) {
	text := strings.TrimSpace(statement.Text)
	if left, _, ok := strings.Cut(text, "="); ok {
		text = strings.TrimSpace(left)
	}
	for _, prefix := range []string{"Set ", "Let ", "ReDim Preserve ", "ReDim "} {
		if len(text) >= len(prefix) && strings.EqualFold(text[:len(prefix)], prefix) {
			text = strings.TrimSpace(text[len(prefix):])
			break
		}
	}
	open := strings.IndexByte(text, '(')
	if open <= 0 {
		return "", "", false
	}
	if _, ok := dcBalancedContentExact(text[open:]); !ok {
		return "", "", false
	}
	name := cleanIdentifier(strings.TrimSpace(text[:open]))
	inside, ok := dcBalancedContentExact(text[open:])
	return name, strings.TrimSpace(inside), name != "" && ok
}

func objectContainerSimpleTarget(statement procedureir.Statement) (string, bool) {
	text := strings.TrimSpace(statement.Text)
	if statement.Target != nil {
		text = strings.TrimSpace(statement.Target.Text)
	}
	for _, prefix := range []string{"Set ", "Let "} {
		if len(text) >= len(prefix) && strings.EqualFold(text[:len(prefix)], prefix) {
			text = strings.TrimSpace(text[len(prefix):])
			break
		}
	}
	if text == "" || strings.ContainsAny(text, ".()= ") {
		return "", false
	}
	return cleanIdentifier(text), true
}

func objectContainerArrayWrite(statement procedureir.Statement, arrayName string) bool {
	text := strings.TrimSpace(statement.Text)
	if len(text) >= len("Erase ") && strings.EqualFold(text[:len("Erase ")], "Erase ") {
		argument := strings.TrimSpace(strings.SplitN(text[len("Erase "):], ",", 2)[0])
		return strings.EqualFold(cleanIdentifier(argument), arrayName)
	}
	if statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet &&
		statement.Kind != procedureir.StatementReDim && statement.Kind != procedureir.StatementFor &&
		statement.Kind != procedureir.StatementForEach {
		return false
	}
	if name, ok := objectContainerIndexedTarget(statement); ok {
		return strings.EqualFold(name, arrayName)
	}
	if name, ok := objectContainerSimpleTarget(statement); ok {
		return strings.EqualFold(name, arrayName)
	}
	return false
}

func objectContainerArrayMutation(statement procedureir.Statement, arrayName string) bool {
	if objectContainerArrayWrite(statement, arrayName) {
		return true
	}
	text := strings.TrimSpace(statement.Text)
	if len(text) < len("Erase ") || !strings.EqualFold(text[:len("Erase ")], "Erase ") {
		return false
	}
	argument := strings.TrimSpace(strings.SplitN(text[len("Erase "):], ",", 2)[0])
	return strings.EqualFold(cleanIdentifier(argument), arrayName)
}

func objectContainerStatementCanReach(view cfg.CFGView, fromID, toID int) bool {
	from, fromOK := view.BlockForStatement(fromID)
	to, toOK := view.BlockForStatement(toID)
	if !fromOK || !toOK {
		return false
	}
	if from.ID == to.ID {
		return fromID <= toID
	}
	seen := map[cfg.BlockID]bool{from.ID: true}
	queue := []cfg.BlockID{from.ID}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		found := false
		view.ForEachOutgoing(current, func(edge cfg.Edge) bool {
			if seen[edge.To] {
				return true
			}
			if edge.To == to.ID {
				found = true
				return false
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
			return true
		})
		if found {
			return true
		}
	}
	return false
}

func objectContainerStatementDominates(view cfg.CFGView, statementID, targetID int) bool {
	statement, statementOK := view.BlockForStatement(statementID)
	target, targetOK := view.BlockForStatement(targetID)
	if !statementOK || !targetOK {
		return false
	}
	if statement.ID == target.ID {
		return statementID <= targetID
	}
	for _, dominator := range view.DominatorsOf(target.ID) {
		if dominator == statement.ID {
			return true
		}
	}
	return false
}

func objectContainerNormalExitDominates(proc sourceProcedure, statementID int) bool {
	if proc.Graph == nil {
		return false
	}
	return objectContainerNormalExitCoveredByWrites(proc, map[int]bool{statementID: true})
}

func objectContainerDictionaryItemExpression(text string) (string, string, bool) {
	text = strings.TrimSpace(text)
	for _, prefix := range []string{"Set ", "Let "} {
		if len(text) >= len(prefix) && strings.EqualFold(text[:len(prefix)], prefix) {
			text = strings.TrimSpace(text[len(prefix):])
			break
		}
	}
	lower := strings.ToLower(text)
	dot := strings.LastIndex(lower, ".item")
	if dot <= 0 {
		return "", "", false
	}
	receiver := cleanIdentifier(strings.TrimSpace(text[:dot]))
	if receiver == "" || strings.ContainsAny(receiver, ".()= ") {
		return "", "", false
	}
	rest := strings.TrimSpace(text[dot+len(".item"):])
	if !strings.HasPrefix(rest, "(") {
		return "", "", false
	}
	inside, ok := dcBalancedContentExact(rest)
	if !ok {
		return "", "", false
	}
	arguments := splitArgs(inside)
	if len(arguments) != 1 {
		return "", "", false
	}
	key, ok := objectContainerLiteral(arguments[0])
	return receiver, key, ok
}

func objectContainerDefaultItemExpression(text string) (string, string, bool) {
	receiver, key, literal, ok := objectContainerDefaultItemTarget(text)
	return receiver, key, literal && ok
}

func objectContainerDefaultItemTarget(text string) (string, string, bool, bool) {
	text = strings.TrimSpace(text)
	if left, _, ok := strings.Cut(text, "="); ok {
		text = strings.TrimSpace(left)
	}
	for _, prefix := range []string{"Set ", "Let "} {
		if len(text) >= len(prefix) && strings.EqualFold(text[:len(prefix)], prefix) {
			text = strings.TrimSpace(text[len(prefix):])
			break
		}
	}
	open := strings.IndexByte(text, '(')
	if open <= 0 {
		return "", "", false, false
	}
	inside, ok := dcBalancedContentExact(text[open:])
	if !ok {
		return "", "", false, false
	}
	receiver := cleanIdentifier(strings.TrimSpace(text[:open]))
	if receiver == "" || strings.ContainsAny(receiver, ".()= ") {
		return "", "", false, false
	}
	key, literal := objectContainerLiteral(inside)
	return receiver, key, literal, true
}

func objectContainerGraphContext(index *objectContainerIndex, proc sourceProcedure) (objectFlowContext, declarationScope, bool) {
	if index == nil || proc.Graph == nil {
		return objectFlowContext{}, declarationScope{}, false
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
	return newObjectFlowContext(proc, graph, index), declarations, true
}

func objectContainerVariableConstructed(index *objectContainerIndex, proc sourceProcedure, name string, statementID int, declarations declarationScope) bool {
	declaration, scope, ok := objectDeclarationBinding(name, declarations)
	if !ok || !declaration.Object {
		return false
	}
	flowContext, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return false
	}
	if declaration.NewExpression {
		// A Dim ... As New object starts initialized, but a later conditional
		// write or ByRef call can still make the reference unknown before use.
	} else {
		if !objectDominatingObjectAssignment(
			proc,
			objectVariable{Scope: scope, Name: name},
			statementID,
			declarations,
			flowContext,
		) {
			return false
		}
		latest, found := objectContainerLastDominatingAssignment(proc, objectVariable{Scope: scope, Name: name}, statementID, declarations, flowContext)
		if !found || latest.Kind == procedureir.StatementForEach || objectErrorResumeNextAt(proc, latest.ID) {
			return false
		}
	}
	return !objectContainerVariableHasUnsafeReachableMutation(index, proc, objectVariable{Scope: scope, Name: name}, statementID, declarations, flowContext)
}

func objectContainerVariableHasUnsafeReachableMutation(index *objectContainerIndex, proc sourceProcedure, variable objectVariable, observationStatementID int, declarations declarationScope, flowContext objectFlowContext) bool {
	if proc.Graph == nil {
		return true
	}
	for statement := range proc.Statements.All() {
		if !objectContainerStatementBeforeObservation(flowContext.graph, statement.ID, observationStatementID) {
			continue
		}
		if statement.Kind == procedureir.StatementForEach {
			target, targetOK := objectContainerLoopTarget(statement)
			if targetOK && strings.EqualFold(target, variable.Name) {
				return true
			}
			continue
		}
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementReDim {
			continue
		}
		target, targetOK := objectFlowTarget(proc, statement, declarations, flowContext)
		if !targetOK || target.key() != variable.key() {
			continue
		}
		block, blockOK := flowContext.graph.BlockForStatement(statement.ID)
		if !blockOK || !flowContext.graph.IsReachable(block.ID) ||
			!objectContainerStatementCanReach(flowContext.graph, statement.ID, observationStatementID) {
			continue
		}
		if statement.Value == nil || strings.EqualFold(strings.TrimSpace(statement.Value.Text), "nothing") {
			return true
		}
		if objectConstructorExpression(statement.Value) {
			if objectContainerErrorHandlerActiveAt(proc, statement.ID) {
				return true
			}
			continue
		}
		if objectContainerStatementDominates(flowContext.graph, statement.ID, observationStatementID) {
			continue
		}
		return true
	}
	for call := range proc.Calls.All() {
		if !objectContainerStatementBeforeObservation(flowContext.graph, call.StatementID, observationStatementID) ||
			!objectContainerStatementCanReach(flowContext.graph, call.StatementID, observationStatementID) {
			continue
		}
		name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
		if objectContainerCallPreservesArguments(call) || name == "ubound" || name == "lbound" ||
			objectContainerKnownPreservingObjectCall(index, proc, call, variable.Name) {
			continue
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if objectContainerArgumentIsIdentifier(argument, variable.Name) {
				return true
			}
		}
	}
	return false
}

func objectContainerKnownPreservingObjectCall(index *objectContainerIndex, proc sourceProcedure, call procedureir.CallSite, name string) bool {
	if call.Callee.Receiver != nil {
		receiver := objectContainerReceiver(call)
		if receiver == "" || !objectContainerKnownContainerReceiver(index, proc, receiver, call.StatementID) {
			return false
		}
		switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
		case "add", "remove", "removeall", "exists", "count", "keys", "item":
			return true
		default:
			return false
		}
	}
	if index == nil {
		return false
	}
	appendInfo, ok := objectContainerAppendInfoForResolvedCall(index, proc, call)
	if !ok || appendInfo.objectParameter < 0 {
		return false
	}
	arguments, ok := objectContainerCallArgumentsForProcedure(index, proc, call, call.Callee.BaseName)
	return ok && appendInfo.objectParameter < len(arguments) && objectContainerArgumentIsIdentifier(arguments[appendInfo.objectParameter], name)
}

func objectContainerKnownContainerReceiver(index *objectContainerIndex, proc sourceProcedure, receiver string, statementID int) bool {
	if index == nil {
		return false
	}
	declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
	declaration, scope, ok := objectDeclarationBinding(receiver, declarations)
	if !ok || !declaration.Object {
		return false
	}
	switch dcKindFromType(declaration.Type) {
	case dcDictionary, dcCollection:
		return true
	}
	if declaration.NewExpression && (dcKindFromType(declaration.Type) == dcDictionary || dcKindFromType(declaration.Type) == dcCollection) {
		return true
	}
	flowContext, _, contextOK := objectContainerGraphContext(index, proc)
	if !contextOK {
		return false
	}
	latest, found := objectContainerLastDominatingAssignment(proc, objectVariable{Scope: scope, Name: receiver}, statementID, declarations, flowContext)
	return found && !objectErrorResumeNextAt(proc, latest.ID) && latest.Value != nil && objectContainerKnownContainerConstructor(latest.Value)
}

func objectContainerKnownContainerConstructor(expression *procedureir.Expression) bool {
	if expression == nil {
		return false
	}
	if objectContainerDictionaryConstructorExpression(expression) {
		return true
	}
	value := strings.ToLower(strings.TrimSpace(expression.Text))
	return strings.HasPrefix(value, "new ") && dcKindFromType(strings.TrimSpace(value[len("new "):])) == dcCollection
}

func objectContainerLastDominatingAssignment(proc sourceProcedure, variable objectVariable, statementID int, declarations declarationScope, flowContext objectFlowContext) (procedureir.Statement, bool) {
	if proc.Graph == nil {
		return procedureir.Statement{}, false
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	callBlock, ok := graph.BlockForStatement(statementID)
	if !ok {
		return procedureir.Statement{}, false
	}
	dominators := graph.Dominators()
	var latest procedureir.Statement
	found := false
	latestID := -1
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementReDim && statement.Kind != procedureir.StatementForEach {
			continue
		}
		target, targetOK := objectFlowTarget(proc, statement, declarations, flowContext)
		if !targetOK || target.key() != variable.key() {
			continue
		}
		block, blockOK := graph.BlockForStatement(statement.ID)
		if !blockOK || !objectBlockSetContains(dominators[callBlock.ID], block.ID) || block.ID == callBlock.ID && statement.ID >= statementID {
			continue
		}
		if statement.ID > latestID {
			latest = statement
			found = true
			latestID = statement.ID
		}
	}
	return latest, found
}

func objectContainerDictionaryVariableConstructed(index *objectContainerIndex, proc sourceProcedure, name string, statementID int, declarations declarationScope) bool {
	declaration, scope, ok := objectDeclarationBinding(name, declarations)
	constructed := objectContainerVariableConstructed(index, proc, name, statementID, declarations)
	if !ok || !declaration.Object || !constructed {
		return false
	}
	flowContext, _, contextOK := objectContainerGraphContext(index, proc)
	if !contextOK {
		return false
	}
	baselineID := 0
	dictionaryProof := dcKindFromType(declaration.Type) == dcDictionary && declaration.NewExpression
	if !dictionaryProof {
		latest, found := objectContainerLastDominatingAssignment(proc, objectVariable{Scope: scope, Name: name}, statementID, declarations, flowContext)
		if !found || !objectContainerDictionaryConstructorExpression(latest.Value) {
			return false
		}
		baselineID = latest.ID
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		if objectContainerReceiverMemberAssignment(statement, name) {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		if !targetOK || !strings.EqualFold(target, name) || statement.ID == baselineID ||
			!objectContainerStatementBeforeObservation(flowContext.graph, statement.ID, statementID) {
			continue
		}
		if baselineID > 0 && !objectContainerStatementCanReach(flowContext.graph, baselineID, statement.ID) {
			continue
		}
		if !objectContainerDictionaryConstructorExpression(statement.Value) || objectErrorResumeNextAt(proc, statement.ID) {
			return false
		}
	}
	return true
}

func objectContainerReceiverMemberAssignment(statement procedureir.Statement, receiver string) bool {
	left, _, ok := strings.Cut(statement.Text, "=")
	if !ok {
		return false
	}
	if leftReceiver, _, _, target := objectContainerDefaultItemTarget(left); target && strings.EqualFold(leftReceiver, receiver) {
		return true
	}
	if leftReceiver, _, literal := objectContainerDictionaryItemExpression(left); literal && strings.EqualFold(leftReceiver, receiver) {
		return true
	}
	left = strings.ToLower(strings.TrimSpace(left))
	return strings.HasPrefix(left, strings.ToLower(cleanIdentifier(receiver))+".item(")
}

func objectContainerArrayArgumentNames(text string) ([]string, bool) {
	text = strings.TrimSpace(text)
	lower := strings.ToLower(text)
	if !strings.HasPrefix(lower, "array(") || !strings.HasSuffix(text, ")") {
		return nil, false
	}
	arguments := splitArgs(text[len("Array(") : len(text)-1])
	if len(arguments) == 0 {
		return nil, false
	}
	names := make([]string, 0, len(arguments))
	for _, argument := range arguments {
		name := cleanIdentifier(strings.TrimSpace(argument))
		if name == "" || name != strings.TrimSpace(argument) || strings.ContainsAny(name, ".()") {
			return nil, false
		}
		names = append(names, name)
	}
	return names, true
}

func objectContainerDirectArrayElements(index *objectContainerIndex, proc sourceProcedure, call procedureir.CallSite, declarations declarationScope) ([]objectContainerElement, bool) {
	if objectErrorResumeNextAt(proc, call.StatementID) {
		return nil, false
	}
	arguments := objectContainerCallArguments(proc, call)
	if len(arguments) < 2 {
		return nil, false
	}
	names, ok := objectContainerArrayArgumentNames(arguments[1])
	if !ok {
		return nil, false
	}
	elements := make([]objectContainerElement, 0, len(names))
	for _, name := range names {
		declaration, _, declared := objectDeclarationBinding(name, declarations)
		if !declared || !declaration.Object || !objectContainerDictionaryVariableConstructed(index, proc, name, call.StatementID, declarations) {
			return nil, false
		}
		elements = append(elements, objectContainerElement{procedureStart: proc.StartLine, name: cleanIdentifier(name), storageStatementID: call.StatementID})
	}
	return elements, true
}

func objectContainerAppendCallName(statement procedureir.Statement) (string, bool) {
	_, right, ok := strings.Cut(statement.Text, "=")
	if !ok {
		return "", false
	}
	right = strings.TrimSpace(right)
	open := strings.IndexByte(right, '(')
	if open <= 0 || !strings.HasSuffix(right, ")") {
		return "", false
	}
	name := cleanIdentifier(strings.TrimSpace(right[:open]))
	if name == "" || strings.ContainsAny(name, ". ") {
		return "", false
	}
	return name, true
}

func findObjectContainerAppendInfo(index *objectContainerIndex, name string) (objectContainerAppendInfo, bool) {
	if index == nil || name == "" {
		return objectContainerAppendInfo{}, false
	}
	var result objectContainerAppendInfo
	found := false
	for _, proc := range index.procedures {
		if !strings.EqualFold(cleanIdentifier(proc.Name), cleanIdentifier(name)) {
			continue
		}
		if found {
			return objectContainerAppendInfo{}, false
		}
		found = true
		result.arrayParameter = -1
		result.objectParameter = -1
		for index, parameter := range proc.Params.AllIndexed() {
			if parameter.IsArray && result.arrayParameter < 0 {
				result.arrayParameter = index
				result.arrayParameterName = parameter.Name
			}
			if result.objectParameter < 0 && isObjectType(parameter.Type) {
				result.objectParameter = index
				result.objectParameterName = parameter.Name
			}
		}
		for index, parameter := range proc.Params.AllIndexed() {
			if index == result.arrayParameter || index == result.objectParameter {
				continue
			}
			if parameter.IsArray || isObjectType(parameter.Type) || strings.EqualFold(cleanIdentifier(parameter.Type), "variant") {
				return objectContainerAppendInfo{}, false
			}
		}
		if result.arrayParameter < 0 || result.objectParameter < 0 || result.arrayParameter == result.objectParameter {
			return objectContainerAppendInfo{}, false
		}
		if !objectContainerAppendObjectParameterPreserved(proc, result.objectParameterName, result.arrayParameterName) {
			return objectContainerAppendInfo{}, false
		}
		if !objectContainerAppendArrayParameterPreserved(proc, result.arrayParameterName) {
			return objectContainerAppendInfo{}, false
		}
		if proc.IR == nil || !proc.IR.Symbol.IsArray && proc.ReturnValueShape != procedureir.ValueShapeFixedArray && proc.ReturnValueShape != procedureir.ValueShapeDynamicArray {
			return objectContainerAppendInfo{}, false
		}
		arrayName := cleanIdentifier(result.arrayParameterName)
		objectName := cleanIdentifier(result.objectParameterName)
		var redim, elementSet, returned procedureir.Statement
		redimFound := false
		elementSetFound := false
		returnFound := false
		for statement := range proc.Statements.All() {
			switch statement.Kind {
			case procedureir.StatementReDim:
				indexedName, indexed := objectContainerIndexedTarget(statement)
				if !indexed || !strings.EqualFold(indexedName, arrayName) || !strings.Contains(strings.ToLower(statement.Text), "redim preserve") {
					if objectContainerArrayWrite(statement, arrayName) {
						return objectContainerAppendInfo{}, false
					}
					continue
				}
				if !objectContainerAppendReDimShape(statement, arrayName) {
					return objectContainerAppendInfo{}, false
				}
				if redimFound {
					return objectContainerAppendInfo{}, false
				}
				redim = statement
				redimFound = true
			case procedureir.StatementSet:
				indexedName, indexed := objectContainerIndexedTarget(statement)
				if indexed && strings.EqualFold(indexedName, arrayName) && statement.Value != nil &&
					statement.Value.Kind == procedureir.ExpressionIdentifier && strings.EqualFold(cleanIdentifier(statement.Value.Text), objectName) {
					_, indexExpression, indexOK := objectContainerIndexedTargetExpression(statement)
					if !indexOK || !objectContainerAppendElementIndex(indexExpression, arrayName) {
						return objectContainerAppendInfo{}, false
					}
					if elementSetFound {
						return objectContainerAppendInfo{}, false
					}
					elementSet = statement
					elementSetFound = true
					continue
				}
				if objectContainerArrayWrite(statement, arrayName) {
					return objectContainerAppendInfo{}, false
				}
			case procedureir.StatementAssignment:
				if objectContainerArrayWrite(statement, arrayName) {
					return objectContainerAppendInfo{}, false
				}
				if objectContainerAssignmentTarget(statement) != "" && strings.EqualFold(objectContainerAssignmentTarget(statement), proc.Name) && statement.Value != nil &&
					statement.Value.Kind == procedureir.ExpressionIdentifier && strings.EqualFold(cleanIdentifier(statement.Value.Text), arrayName) {
					if returnFound {
						return objectContainerAppendInfo{}, false
					}
					returned = statement
					returnFound = true
				}
			default:
				if objectContainerArrayWrite(statement, arrayName) {
					return objectContainerAppendInfo{}, false
				}
			}
		}
		if !redimFound || !elementSetFound || !returnFound || objectErrorResumeNextAt(proc, redim.ID) ||
			objectErrorResumeNextAt(proc, elementSet.ID) || objectErrorResumeNextAt(proc, returned.ID) ||
			!objectContainerNormalExitDominates(proc, redim.ID) ||
			!objectContainerNormalExitDominates(proc, elementSet.ID) || !objectContainerNormalExitDominates(proc, returned.ID) ||
			!objectContainerStatementDominates(proc.Graph.WithoutNormalErrRaiseContinuationView(), redim.ID, elementSet.ID) ||
			!objectContainerStatementDominates(proc.Graph.WithoutNormalErrRaiseContinuationView(), elementSet.ID, returned.ID) {
			return objectContainerAppendInfo{}, false
		}
	}
	return result, found
}

func objectContainerAppendInfoForResolvedCall(index *objectContainerIndex, proc sourceProcedure, call procedureir.CallSite) (objectContainerAppendInfo, bool) {
	if index == nil || call.Callee.Receiver != nil || !objectContainerCallResolutionUsable(call) {
		return objectContainerAppendInfo{}, false
	}
	name := cleanIdentifier(call.Callee.BaseName)
	if name == "" {
		return objectContainerAppendInfo{}, false
	}
	resolved := call.Resolution.Candidates[0]
	found := false
	for _, candidate := range index.procedures {
		if !strings.EqualFold(cleanIdentifier(candidate.Name), name) || !objectContainerResolvedCandidateMatches(candidate, resolved) {
			continue
		}
		if found {
			return objectContainerAppendInfo{}, false
		}
		found = true
	}
	if !found {
		return objectContainerAppendInfo{}, false
	}
	return findObjectContainerAppendInfo(index, name)
}

func objectContainerAppendObjectParameterPreserved(proc sourceProcedure, name, arrayName string) bool {
	for statement := range proc.Statements.All() {
		if statement.Kind == procedureir.StatementForEach {
			if target, ok := objectContainerLoopTarget(statement); ok && strings.EqualFold(target, name) {
				return false
			}
			continue
		}
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		if target, ok := objectContainerSimpleTarget(statement); ok {
			if strings.EqualFold(target, name) {
				return false
			}
		}
		if statement.Value == nil || !objectContainerArgumentIsIdentifier(statement.Value.Text, name) {
			continue
		}
		indexedName, indexExpression, indexed := objectContainerIndexedTargetExpression(statement)
		if !indexed || !strings.EqualFold(indexedName, arrayName) || !objectContainerAppendElementIndex(indexExpression, arrayName) {
			return false
		}
	}
	for call := range proc.Calls.All() {
		if strings.EqualFold(objectCallWithReceiverName(proc, call), name) {
			return false
		}
	}
	for call := range proc.Calls.All() {
		for _, argument := range objectContainerCallArguments(proc, call) {
			argument = strings.TrimSpace(argument)
			for strings.HasPrefix(argument, "(") && strings.HasSuffix(argument, ")") {
				inside, ok := dcBalancedContentExact(argument)
				if !ok {
					break
				}
				argument = strings.TrimSpace(inside)
			}
			if strings.EqualFold(cleanIdentifier(argument), name) {
				return false
			}
		}
	}
	return true
}

func objectContainerAppendArrayParameterPreserved(proc sourceProcedure, name string) bool {
	for call := range proc.Calls.All() {
		callName := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
		if call.Resolution.Status == procedureir.ResolutionBuiltinLike || callName == "ubound" || callName == "lbound" {
			continue
		}
		if call.Callee.Receiver != nil && objectContainerArrayElementArgument(*call.Callee.Receiver, name) {
			return false
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if objectContainerArrayElementArgument(argument, name) {
				return false
			}
			if objectContainerArgumentIsIdentifier(argument, name) {
				return false
			}
		}
	}
	return true
}

func objectContainerArgumentIsIdentifier(text, name string) bool {
	text = strings.TrimSpace(text)
	for strings.HasPrefix(text, "(") && strings.HasSuffix(text, ")") {
		inside, ok := dcBalancedContentExact(text)
		if !ok {
			return false
		}
		text = strings.TrimSpace(inside)
	}
	return strings.EqualFold(cleanIdentifier(text), name)
}

func objectContainerAppendReDimShape(statement procedureir.Statement, arrayName string) bool {
	_, bounds, ok := objectContainerIndexedTargetExpression(statement)
	if !ok {
		return false
	}
	parts := strings.SplitN(strings.ToLower(bounds), " to ", 2)
	if len(parts) != 2 {
		return false
	}
	lower := strings.Join(strings.Fields(parts[0]), "")
	upper := strings.Join(strings.Fields(parts[1]), "")
	return lower == "0" && upper == "ubound("+strings.ToLower(cleanIdentifier(arrayName))+")+1"
}

func objectContainerAppendElementIndex(indexExpression, arrayName string) bool {
	normalizedIndex := strings.Join(strings.Fields(strings.ToLower(indexExpression)), "")
	normalizedUpper := "ubound(" + strings.ToLower(cleanIdentifier(arrayName)) + ")"
	return normalizedIndex == normalizedUpper
}

func objectContainerArrayElementArgument(text, arrayName string) bool {
	text = strings.TrimSpace(text)
	open := strings.IndexByte(text, '(')
	if open <= 0 || !strings.EqualFold(cleanIdentifier(strings.TrimSpace(text[:open])), arrayName) {
		return false
	}
	_, ok := dcBalancedContentExact(text[open:])
	return ok
}

func objectContainerReflectiveCall(call procedureir.CallSite) bool {
	return strings.EqualFold(cleanIdentifier(call.Callee.BaseName), "callbyname")
}

func objectContainerCallPreservesArguments(call procedureir.CallSite) bool {
	if call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return true
	}
	switch strings.ToLower(cleanIdentifier(call.Callee.BaseName)) {
	case "array", "isobject", "strcomp", "typename", "ubound", "lbound":
		return true
	default:
		return false
	}
}

func objectContainerDirectArrayElement(statement procedureir.Statement) (string, string, bool) {
	_, right, ok := strings.Cut(statement.Text, "=")
	if !ok {
		return "", "", false
	}
	right = strings.TrimSpace(right)
	open := strings.IndexByte(right, '(')
	if open <= 0 {
		return "", "", false
	}
	inside, ok := dcBalancedContentExact(right[open:])
	if !ok || strings.TrimSpace(inside) == "" {
		return "", "", false
	}
	return cleanIdentifier(strings.TrimSpace(right[:open])), strings.TrimSpace(inside), true
}

func objectContainerDirectArrayElementIndex(statement procedureir.Statement, arrayName string) (string, bool) {
	actualArray, indexExpression, ok := objectContainerDirectArrayElement(statement)
	if !ok || !strings.EqualFold(actualArray, arrayName) {
		return "", false
	}
	return indexExpression, true
}

func objectContainerArrayIndexLoopBound(proc sourceProcedure, arrayName, indexExpression string, statementID int, view cfg.CFGView) bool {
	indexExpression = strings.TrimSpace(indexExpression)
	if indexExpression == "" {
		return false
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementFor || !objectContainerStatementDominates(view, statement.ID, statementID) {
			continue
		}
		text := strings.TrimSpace(statement.Text)
		if len(text) < len("For ") || !strings.EqualFold(text[:len("For ")], "For ") {
			continue
		}
		text = strings.TrimSpace(text[len("For "):])
		if newline := strings.IndexByte(text, '\n'); newline >= 0 {
			text = strings.TrimSpace(text[:newline])
		}
		left, right, ok := strings.Cut(text, "=")
		if !ok || !strings.EqualFold(cleanIdentifier(strings.TrimSpace(left)), indexExpression) {
			continue
		}
		lowerRight := strings.ToLower(right)
		separator := strings.Index(lowerRight, " to ")
		if separator < 0 {
			continue
		}
		lower := strings.TrimSpace(right[:separator])
		upper := strings.TrimSpace(right[separator+len(" to "):])
		if step := strings.Index(strings.ToLower(upper), " step "); step >= 0 {
			if !strings.EqualFold(strings.TrimSpace(upper[step+len(" step "):]), "1") {
				continue
			}
			upper = strings.TrimSpace(upper[:step])
		}
		lowerOK := lower == "0" || strings.EqualFold(lower, "lbound("+arrayName+")")
		upperOK := strings.EqualFold(upper, "ubound("+arrayName+")")
		if lowerOK && upperOK {
			if !objectContainerStatementInLoopBody(view, statement.ID, statementID) {
				continue
			}
			if objectContainerLoopIndexReassigned(proc, indexExpression, statement.ID, statementID, view) {
				continue
			}
			return true
		}
	}
	return false
}

func objectContainerStatementInLoopBody(view cfg.CFGView, loopStatementID, observationStatementID int) bool {
	loopBlock, loopOK := view.BlockForStatement(loopStatementID)
	observationBlock, observationOK := view.BlockForStatement(observationStatementID)
	if !loopOK || !observationOK || loopBlock.ID == observationBlock.ID {
		return false
	}
	queue := make([]cfg.BlockID, 0, 1)
	view.ForEachOutgoing(loopBlock.ID, func(edge cfg.Edge) bool {
		if edge.Kind == cfg.EdgeLoopBody {
			queue = append(queue, edge.To)
		}
		return true
	})
	seen := map[cfg.BlockID]bool{}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		if current == observationBlock.ID {
			return true
		}
		view.ForEachOutgoing(current, func(edge cfg.Edge) bool {
			if edge.Kind == cfg.EdgeLoopExit || edge.Kind == cfg.EdgeLoopBack || edge.To == loopBlock.ID {
				return true
			}
			if !seen[edge.To] {
				queue = append(queue, edge.To)
			}
			return true
		})
	}
	return false
}

func objectContainerLoopIndexReassigned(proc sourceProcedure, indexName string, loopStatementID, observationStatementID int, view cfg.CFGView) bool {
	for statement := range proc.Statements.All() {
		if statement.ID == loopStatementID ||
			!objectContainerStatementCanReach(view, loopStatementID, statement.ID) ||
			!objectContainerStatementCanReach(view, statement.ID, observationStatementID) {
			continue
		}
		switch statement.Kind {
		case procedureir.StatementAssignment, procedureir.StatementSet:
			target, ok := objectContainerSimpleTarget(statement)
			if ok && strings.EqualFold(target, indexName) {
				return true
			}
		case procedureir.StatementFor, procedureir.StatementForEach:
			target, ok := objectContainerLoopTarget(statement)
			if ok && strings.EqualFold(target, indexName) {
				return true
			}
		}
	}
	for call := range proc.Calls.All() {
		if !objectContainerStatementCanReach(view, loopStatementID, call.StatementID) ||
			!objectContainerStatementCanReach(view, call.StatementID, observationStatementID) {
			continue
		}
		if objectContainerArrayIndexReadCall(proc, call, indexName) {
			continue
		}
		name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
		if objectContainerCallPreservesArguments(call) || name == "ubound" || name == "lbound" {
			continue
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if objectContainerArgumentIsIdentifier(argument, indexName) {
				return true
			}
		}
	}
	return false
}

func objectContainerArrayIndexReadCall(proc sourceProcedure, call procedureir.CallSite, indexName string) bool {
	if call.Callee.Receiver != nil {
		return false
	}
	statement, ok := objectContainerStatement(proc, call.StatementID)
	if !ok || statement.Value == nil {
		return false
	}
	text := strings.TrimSpace(statement.Value.Text)
	open := strings.IndexByte(text, '(')
	if open <= 0 || !strings.EqualFold(cleanIdentifier(strings.TrimSpace(text[:open])), cleanIdentifier(call.Callee.BaseName)) {
		return false
	}
	inside, balanced := dcBalancedContentExact(text[open:])
	return balanced && objectContainerArgumentIsIdentifier(inside, indexName)
}

func objectContainerLoopTarget(statement procedureir.Statement) (string, bool) {
	text := strings.TrimSpace(statement.Text)
	if newline := strings.IndexByte(text, '\n'); newline >= 0 {
		text = strings.TrimSpace(text[:newline])
	}
	lower := strings.ToLower(text)
	for _, prefix := range []string{"for each ", "for "} {
		if strings.HasPrefix(lower, prefix) {
			text = strings.TrimSpace(text[len(prefix):])
			break
		}
	}
	if text == statement.Text && !strings.HasPrefix(lower, "for ") {
		return "", false
	}
	if left, _, ok := strings.Cut(text, "="); ok {
		text = left
	} else if lowerIndex := strings.Index(strings.ToLower(text), " in "); lowerIndex >= 0 {
		text = text[:lowerIndex]
	} else {
		return "", false
	}
	name := cleanIdentifier(strings.TrimSpace(text))
	return name, name != "" && !strings.ContainsAny(name, ".()")
}

func objectContainerArrayIndexSafe(proc sourceProcedure, arrayName, indexExpression string, statementID int, view cfg.CFGView, elementCount int) bool {
	indexExpression = strings.TrimSpace(indexExpression)
	if value, err := strconv.Atoi(indexExpression); err == nil {
		// The merged contract records element identities, not path-specific array
		// lengths. Index zero is safe for every non-empty proof; higher literal
		// indexes require cardinality information that is not represented here.
		return value == 0 && elementCount > 0
	}
	return objectContainerArrayIndexLoopBound(proc, arrayName, indexExpression, statementID, view)
}

func objectContainerSafeAppendCallForElement(index *objectContainerIndex, proc sourceProcedure, call procedureir.CallSite, elementName string) bool {
	if objectContainerReflectiveCall(call) {
		return false
	}
	if call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return true
	}
	if call.Callee.Receiver != nil {
		return false
	}
	appendInfo, ok := objectContainerAppendInfoForResolvedCall(index, proc, call)
	if !ok {
		return false
	}
	arguments, argumentsOK := objectContainerCallArgumentsForProcedure(index, proc, call, call.Callee.BaseName)
	if !argumentsOK {
		return false
	}
	return appendInfo.objectParameter >= 0 && appendInfo.objectParameter < len(arguments) &&
		objectContainerArgumentIsIdentifier(arguments[appendInfo.objectParameter], elementName)
}

func objectContainerArraySource(proc sourceProcedure, arrayName string, statementID int, flowContext objectFlowContext) (string, string, bool) {
	if flowContext.graph.BlockCount() == 0 {
		return "", "", false
	}
	callBlock, ok := flowContext.graph.BlockForStatement(statementID)
	if !ok {
		return "", "", false
	}
	dominators := flowContext.graph.Dominators()
	var selected procedureir.Statement
	selectedID := -1
	for statement := range proc.Statements.All() {
		if statement.Value == nil || !strings.EqualFold(objectContainerAssignmentTarget(statement), arrayName) {
			continue
		}
		block, ok := flowContext.graph.BlockForStatement(statement.ID)
		if !ok || !objectBlockSetContains(dominators[callBlock.ID], block.ID) || block.ID == callBlock.ID && statement.ID >= statementID {
			continue
		}
		if statement.ID <= selectedID {
			continue
		}
		selected = statement
		selectedID = statement.ID
	}
	if selectedID < 0 {
		return "", "", false
	}
	if objectErrorResumeNextAt(proc, selectedID) {
		return "", "", false
	}
	receiver, key, sourceOK := objectContainerDictionaryItemExpression(selected.Value.Text)
	if !sourceOK {
		receiver, key, sourceOK = objectContainerDefaultItemExpression(selected.Value.Text)
	}
	if !sourceOK {
		return "", "", false
	}
	for statement := range proc.Statements.All() {
		if statement.ID == selectedID || !objectContainerArrayMutation(statement, arrayName) {
			continue
		}
		if objectContainerStatementCanReach(flowContext.graph, selectedID, statement.ID) &&
			objectContainerStatementCanReach(flowContext.graph, statement.ID, statementID) {
			return "", "", false
		}
	}
	for call := range proc.Calls.All() {
		if call.StatementID == selectedID || call.StatementID == statementID {
			continue
		}
		if !objectContainerStatementCanReach(flowContext.graph, selectedID, call.StatementID) ||
			!objectContainerStatementCanReach(flowContext.graph, call.StatementID, statementID) {
			continue
		}
		if objectContainerReflectiveCall(call) {
			return "", "", false
		}
		if call.Callee.Receiver != nil && objectContainerArrayElementArgument(*call.Callee.Receiver, arrayName) {
			return "", "", false
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if objectContainerArrayElementArgument(argument, arrayName) {
				return "", "", false
			}
			if !strings.EqualFold(cleanIdentifier(strings.TrimSpace(argument)), arrayName) {
				continue
			}
			name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
			if name != "ubound" && name != "lbound" {
				return "", "", false
			}
		}
	}
	return receiver, key, true
}

func objectContainerDictionaryItemTarget(statement procedureir.Statement, receiver, key string) (bool, bool) {
	left, _, hasAssignment := strings.Cut(statement.Text, "=")
	if !hasAssignment {
		return false, false
	}
	leftReceiver, leftKey, literal := objectContainerDictionaryItemExpression(left)
	if literal {
		return strings.EqualFold(leftReceiver, receiver) && objectContainerKeysMayAlias(leftKey, key), true
	}
	leftReceiver, leftKey, literal = objectContainerDefaultItemExpression(left)
	if literal {
		return strings.EqualFold(leftReceiver, receiver) && objectContainerKeysMayAlias(leftKey, key), true
	}
	leftReceiver, _, _, defaultTarget := objectContainerDefaultItemTarget(left)
	if defaultTarget && strings.EqualFold(leftReceiver, receiver) {
		// A non-literal default Item key may evaluate to the observed key.
		return true, true
	}
	left = strings.ToLower(strings.TrimSpace(left))
	for _, prefix := range []string{"set ", "let "} {
		if strings.HasPrefix(left, prefix) {
			left = strings.TrimSpace(left[len(prefix):])
			break
		}
	}
	needle := strings.ToLower(cleanIdentifier(receiver)) + ".item("
	if strings.HasPrefix(left, needle) {
		// A dynamic Item key may refer to the observed key. Keep the proof
		// conservative until the key is known to be different.
		return true, true
	}
	return false, false
}

func objectContainerStatementBeforeObservation(view cfg.CFGView, statementID, observationStatementID int) bool {
	if observationStatementID <= 0 {
		return true
	}
	return objectContainerStatementCanReach(view, statementID, observationStatementID)
}

func objectContainerBlockCanReach(view cfg.CFGView, fromID, toID cfg.BlockID) bool {
	if fromID == toID {
		return true
	}
	seen := map[cfg.BlockID]bool{fromID: true}
	queue := []cfg.BlockID{fromID}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		found := false
		view.ForEachOutgoing(current, func(edge cfg.Edge) bool {
			if edge.To == toID {
				found = true
				return false
			}
			if seen[edge.To] {
				return true
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
			return true
		})
		if found {
			return true
		}
	}
	return false
}

// objectContainerCallExceptionalPathCanReach reports whether a failed call can
// still reach the observation (or the caller's normal exit).  A callee's
// constructor/write proof is valid only on paths where the call completes;
// Resume Next and handlers otherwise let execution continue without that
// proof having taken effect.
func objectContainerCallExceptionalPathCanReach(view cfg.CFGView, statementID, observationStatementID int, requireDominance bool) bool {
	callBlock, ok := view.BlockForStatement(statementID)
	if !ok {
		return false
	}
	target := view.NormalExit()
	if !requireDominance {
		observation, observationOK := view.BlockForStatement(observationStatementID)
		if !observationOK {
			return false
		}
		target = observation.ID
	}
	mayBypass := false
	view.ForEachOutgoing(callBlock.ID, func(edge cfg.Edge) bool {
		if edge.Class == cfg.EdgeExceptional && objectContainerBlockCanReach(view, edge.To, target) {
			mayBypass = true
			return false
		}
		return true
	})
	return mayBypass
}

func objectContainerErrorHandlerActiveAt(proc sourceProcedure, statementID int) bool {
	active := false
	for statement := range proc.Statements.All() {
		if statement.ID == statementID {
			return active
		}
		if statement.Kind != procedureir.StatementOnError {
			continue
		}
		if statement.Control != nil {
			switch statement.Control.Transfer {
			case procedureir.TransferOnErrorGoto:
				active = true
			case procedureir.TransferOnErrorDisable, procedureir.TransferOnErrorResumeNext:
				active = false
			}
			continue
		}
		text := strings.ToLower(strings.TrimSpace(statement.Text))
		switch {
		case strings.Contains(text, "on error goto 0"), strings.Contains(text, "on error goto -1"), strings.Contains(text, "on error resume next"):
			active = false
		case strings.Contains(text, "on error goto"):
			active = true
		}
	}
	return active
}

func objectContainerExistsPresentBranchReaches(view cfg.CFGView, guardStatementID, observationStatementID int, negated bool) bool {
	guard, guardOK := view.BlockForStatement(guardStatementID)
	observation, observationOK := view.BlockForStatement(observationStatementID)
	if !guardOK || !observationOK {
		return false
	}
	var trueTarget, falseTarget cfg.BlockID
	trueFound, falseFound := false, false
	view.ForEachOutgoing(guard.ID, func(edge cfg.Edge) bool {
		switch edge.Kind {
		case cfg.EdgeBranchTrue:
			trueTarget, trueFound = edge.To, true
		case cfg.EdgeBranchFalse:
			falseTarget, falseFound = edge.To, true
		}
		return true
	})
	if !trueFound || !falseFound {
		return false
	}
	presentTarget, absentTarget := trueTarget, falseTarget
	if negated {
		presentTarget, absentTarget = falseTarget, trueTarget
	}
	return objectContainerBlockCanReach(view, presentTarget, observation.ID) && !objectContainerBlockCanReach(view, absentTarget, observation.ID)
}

func objectContainerExistsGuarded(index *objectContainerIndex, owner sourceProcedure, receiver, key string, observationStatementID int) bool {
	if index == nil || observationStatementID <= 0 {
		return false
	}
	flowContext, _, ok := objectContainerGraphContext(index, owner)
	if !ok {
		return false
	}
	for call := range owner.Calls.All() {
		if !strings.EqualFold(objectCallWithReceiverName(owner, call), receiver) || !strings.EqualFold(cleanIdentifier(call.Callee.Member), "exists") {
			continue
		}
		callKey, keyOK := objectContainerDictionaryKey(owner, call)
		if keyOK && objectContainerKeysEqual(callKey, key) && objectContainerStatementDominates(flowContext.graph, call.StatementID, observationStatementID) {
			statement, statementOK := objectContainerStatement(owner, call.StatementID)
			if !statementOK {
				continue
			}
			guardReceiver, guardKey, negated, guardOK := dcExistsCondition(statement.Text)
			guardLiteral, guardLiteralOK := objectContainerLiteral(guardKey)
			if !guardOK || !guardLiteralOK || !strings.EqualFold(guardReceiver, receiver) || !objectContainerKeysEqual(guardLiteral, callKey) || !objectContainerExistsPresentBranchReaches(flowContext.graph, call.StatementID, observationStatementID, negated) {
				continue
			}
			return true
		}
	}
	return false
}

func objectContainerReceiverAliasUnsafe(index *objectContainerIndex, owner, proc sourceProcedure, receiver string, observationStatementID int) bool {
	flowContext, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return true
	}
	for statement := range proc.Statements.All() {
		if !objectContainerStatementBeforeObservation(flowContext.graph, statement.ID, ownerObservationID(owner, proc, observationStatementID)) {
			continue
		}
		if statement.Kind == procedureir.StatementForEach {
			if target, targetOK := objectContainerLoopTarget(statement); targetOK && strings.EqualFold(target, receiver) {
				return true
			}
			continue
		}
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment || statement.Value == nil {
			continue
		}
		if !objectContainerArgumentIsIdentifier(statement.Value.Text, receiver) {
			continue
		}
		if target, targetOK := objectContainerSimpleTarget(statement); targetOK {
			if !strings.EqualFold(target, receiver) {
				return true
			}
			continue
		}
		return true
	}
	for call := range proc.Calls.All() {
		if ownerObservationID(owner, proc, observationStatementID) != 0 && call.StatementID == observationStatementID {
			continue
		}
		if !objectContainerStatementBeforeObservation(flowContext.graph, call.StatementID, ownerObservationID(owner, proc, observationStatementID)) {
			continue
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if strings.EqualFold(cleanIdentifier(strings.TrimSpace(argument)), receiver) {
				return true
			}
		}
	}
	return false
}

func objectContainerProcedureCalledBeforeObservation(index *objectContainerIndex, owner, target sourceProcedure, observationStatementID int) bool {
	return objectContainerProcedurePathBeforeObservation(index, owner, target, observationStatementID, true)
}

func objectContainerProcedureMayReachObservation(index *objectContainerIndex, owner, target sourceProcedure, observationStatementID int) bool {
	return objectContainerProcedurePathBeforeObservation(index, owner, target, observationStatementID, false)
}

// A fluent class method is an externally callable writer even when the local
// procedure graph does not contain a call to it.  Selenium-style action
// chains expose their mutators as `Public Function ... As <ClassName>` and
// return Me, so treating those methods as unreachable would erase the only
// evidence that initializes the module container.  Keep ordinary Public Subs
// and non-fluent functions call-graph constrained; they are covered by the
// conservative uncalled-writer rule below.
func objectContainerPublicFluentWriter(proc sourceProcedure) bool {
	if !strings.EqualFold(strings.TrimSpace(proc.ModuleKind), "class") ||
		!strings.EqualFold(strings.TrimSpace(proc.Visibility), "public") ||
		(proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet) {
		return false
	}
	module := cleanIdentifier(proc.Module)
	return module != "" && strings.EqualFold(cleanIdentifier(proc.ReturnType), module)
}

func objectContainerProcedurePathBeforeObservation(index *objectContainerIndex, owner, target sourceProcedure, observationStatementID int, requireDominance bool) bool {
	if index == nil {
		return false
	}
	if owner.StartLine == target.StartLine ||
		(strings.EqualFold(target.Name, "Class_Initialize") && strings.EqualFold(target.ModuleKind, "class")) {
		return true
	}
	reached := map[int]bool{owner.StartLine: true}
	queue := []sourceProcedure{owner}
	for len(queue) > 0 {
		caller := queue[0]
		queue = queue[1:]
		flowContext, _, ok := objectContainerGraphContext(index, caller)
		if !ok {
			continue
		}
		for call := range caller.Calls.All() {
			if call.Callee.Receiver != nil && !strings.EqualFold(cleanIdentifier(strings.TrimSpace(*call.Callee.Receiver)), "me") {
				continue
			}
			if !objectContainerCallResolutionUsable(call) {
				continue
			}
			if objectErrorResumeNextAt(caller, call.StatementID) || objectContainerErrorHandlerActiveAt(caller, call.StatementID) ||
				objectContainerCallExceptionalPathCanReach(flowContext.graph, call.StatementID, observationStatementID, caller.StartLine != owner.StartLine) {
				continue
			}
			if caller.StartLine == owner.StartLine {
				callBlock, blockOK := flowContext.graph.BlockForStatement(call.StatementID)
				if !blockOK || !flowContext.graph.IsReachable(callBlock.ID) ||
					requireDominance && !objectContainerStatementDominates(flowContext.graph, call.StatementID, observationStatementID) ||
					!requireDominance && !objectContainerStatementCanReach(flowContext.graph, call.StatementID, observationStatementID) {
					continue
				}
			} else {
				callBlock, blockOK := flowContext.graph.BlockForStatement(call.StatementID)
				if !blockOK || !flowContext.graph.IsReachable(callBlock.ID) ||
					requireDominance && !objectContainerBlockDominates(flowContext.graph, callBlock.ID, flowContext.graph.NormalExit()) ||
					!requireDominance && !objectContainerBlockCanReach(flowContext.graph, callBlock.ID, flowContext.graph.NormalExit()) {
					continue
				}
			}
			name := cleanIdentifier(call.Callee.BaseName)
			if name == "" {
				name = cleanIdentifier(call.Callee.Member)
			}
			for _, candidate := range index.procedures {
				if !strings.EqualFold(cleanIdentifier(candidate.Name), name) {
					continue
				}
				if !objectContainerResolvedCandidateMatches(candidate, call.Resolution.Candidates[0]) {
					continue
				}
				if candidate.StartLine == target.StartLine {
					return true
				}
				if reached[candidate.StartLine] {
					continue
				}
				reached[candidate.StartLine] = true
				queue = append(queue, candidate)
			}
		}
	}
	return false
}

func objectContainerCallResolutionUsable(call procedureir.CallSite) bool {
	if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
		return false
	}
	return true
}

func objectContainerResolvedCandidateMatches(proc sourceProcedure, candidate procedureir.Candidate) bool {
	qualified := strings.TrimSpace(candidate.QualifiedName)
	if qualified != "" && !strings.EqualFold(qualified, objectProcedureQualifiedName(proc)) {
		return false
	}
	return candidate.Line <= 0 || proc.StartLine <= 0 || candidate.Line == proc.StartLine
}

func objectContainerBlockDominates(view cfg.CFGView, blockID, targetID cfg.BlockID) bool {
	if !view.IsReachable(blockID) || !view.IsReachable(targetID) {
		return false
	}
	if blockID == targetID {
		return true
	}
	for _, dominator := range view.DominatorsOf(targetID) {
		if dominator == blockID {
			return true
		}
	}
	return false
}

func objectContainerProcedureRemovesKey(index *objectContainerIndex, proc sourceProcedure, receiver, key string, observationStatementID int) bool {
	flowContext, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return true
	}
	for call := range proc.Calls.All() {
		if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) || !objectContainerStatementBeforeObservation(flowContext.graph, call.StatementID, observationStatementID) {
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

func objectContainerProcedureMayMutateReceiver(index *objectContainerIndex, proc sourceProcedure, receiver, key string) bool {
	flowContext, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return true
	}
	for statement := range proc.Statements.All() {
		block, blockOK := flowContext.graph.BlockForStatement(statement.ID)
		if !blockOK || !flowContext.graph.IsReachable(block.ID) {
			continue
		}
		if statement.Kind == procedureir.StatementForEach {
			if target, targetOK := objectContainerLoopTarget(statement); targetOK && strings.EqualFold(target, receiver) {
				return true
			}
			continue
		}
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		if target, targetOK := objectContainerSimpleTarget(statement); targetOK && strings.EqualFold(target, receiver) {
			return true
		}
		if objectContainerReceiverMemberAssignment(statement, receiver) {
			return true
		}
	}
	for call := range proc.Calls.All() {
		block, blockOK := flowContext.graph.BlockForStatement(call.StatementID)
		if !blockOK || !flowContext.graph.IsReachable(block.ID) {
			continue
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if objectContainerArgumentIsIdentifier(argument, receiver) {
				return true
			}
		}
		if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) {
			continue
		}
		switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
		case "exists", "count", "keys":
			continue
		case "item":
			statement, statementOK := objectContainerStatement(proc, call.StatementID)
			if !statementOK {
				return true
			}
			_, assignment := objectContainerDictionaryItemTarget(statement, receiver, key)
			if assignment {
				return true
			}
		default:
			return true
		}
	}
	return false
}

func objectContainerModuleReceiverIsPrivate(index *objectContainerIndex, receiver string) bool {
	if index == nil {
		return false
	}
	declaration, ok := index.moduleDecls[strings.ToLower(cleanIdentifier(receiver))]
	if !ok || declaration.Line <= 0 || declaration.Line > len(index.file.Lines) {
		return false
	}
	line := strings.ToLower(strings.TrimSpace(normalizedCodeLine(index.file.Lines[declaration.Line-1])))
	return strings.HasPrefix(line, "private ")
}

func objectContainerDictionaryConstructorExpression(expression *procedureir.Expression) bool {
	if expression == nil {
		return false
	}
	text := strings.ToLower(strings.TrimSpace(expression.Text))
	if strings.HasPrefix(text, "createobject(") && strings.HasSuffix(text, ")") {
		argument := strings.TrimSpace(text[len("createobject(") : len(text)-1])
		if value, ok := objectContainerLiteral(argument); ok {
			return strings.EqualFold(value, "scripting.dictionary")
		}
	}
	return strings.HasPrefix(text, "new ") && dcKindFromType(strings.TrimSpace(text[len("new "):])) == dcDictionary
}

func objectContainerDictionaryReceiverKnown(index *objectContainerIndex, owner sourceProcedure, receiver string, scope procedureir.SymbolScope, observationStatementID int) bool {
	if index == nil {
		return false
	}
	declarations := objectFlowDeclarations(index.file, owner, index.moduleDecls)
	declaration, _, ok := objectDeclarationBinding(receiver, declarations)
	if !ok {
		return false
	}
	if dcKindFromType(declaration.Type) == dcDictionary && declaration.NewExpression {
		return true
	}
	if declaration.NewExpression {
		return false
	}
	if scope == procedureir.ScopeParameter {
		return false
	}
	if scope != procedureir.ScopeModule {
		return objectContainerDictionaryVariableConstructed(index, owner, receiver, observationStatementID, declarations)
	}
	for _, proc := range index.procedures {
		candidateDeclarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
		_, candidateScope, candidateOK := objectDeclarationBinding(receiver, candidateDeclarations)
		if !candidateOK || candidateScope != scope {
			continue
		}
		if proc.StartLine == owner.StartLine {
			if objectContainerDictionaryVariableConstructed(index, owner, receiver, observationStatementID, declarations) {
				return true
			}
			continue
		}
		if !objectContainerProcedureCalledBeforeObservation(index, owner, proc, observationStatementID) {
			continue
		}
		if objectContainerModuleDictionaryAssignmentKnown(index, proc, receiver) {
			return true
		}
	}
	return false
}

func objectContainerModuleDictionaryAssignmentKnown(index *objectContainerIndex, proc sourceProcedure, receiver string) bool {
	_, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return false
	}
	found := false
	dominatesExit := false
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		if !targetOK || !strings.EqualFold(target, receiver) {
			continue
		}
		found = true
		if !objectContainerDictionaryConstructorExpression(statement.Value) || objectErrorResumeNextAt(proc, statement.ID) {
			return false
		}
		if objectContainerNormalExitDominates(proc, statement.ID) {
			dominatesExit = true
		}
	}
	return found && dominatesExit
}

func objectContainerParameterReceiverMutationUnsafe(index *objectContainerIndex, proc sourceProcedure, receiver, key string, observationStatementID int) bool {
	flowContext, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return true
	}
	for statement := range proc.Statements.All() {
		if !objectContainerStatementBeforeObservation(flowContext.graph, statement.ID, observationStatementID) {
			continue
		}
		if statement.Kind == procedureir.StatementForEach {
			if target, targetOK := objectContainerLoopTarget(statement); targetOK && strings.EqualFold(target, receiver) {
				return true
			}
			continue
		}
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		if target, targetOK := objectContainerSimpleTarget(statement); targetOK && strings.EqualFold(target, receiver) {
			return true
		}
		if statement.Value != nil && objectContainerArgumentIsIdentifier(statement.Value.Text, receiver) {
			return true
		}
		if targetReceiver, targetKey, targetLiteral, targetOK := objectContainerDefaultItemTarget(statement.Text); targetOK &&
			strings.EqualFold(targetReceiver, receiver) && (!targetLiteral || objectContainerKeysMayAlias(targetKey, key)) {
			return true
		}
	}
	for call := range proc.Calls.All() {
		if call.StatementID == observationStatementID || !objectContainerStatementBeforeObservation(flowContext.graph, call.StatementID, observationStatementID) {
			continue
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if objectContainerArgumentIsIdentifier(argument, receiver) {
				return true
			}
		}
		if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) {
			continue
		}
		if objectContainerReflectiveCall(call) {
			return true
		}
		switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
		case "exists", "count", "keys":
			continue
		case "item":
			statement, statementOK := objectContainerStatement(proc, call.StatementID)
			if !statementOK {
				return true
			}
			matched, assignment := objectContainerDictionaryItemTarget(statement, receiver, key)
			if assignment && matched {
				return true
			}
		default:
			return true
		}
	}
	return false
}

func objectContainerNormalExitCoveredByWrites(proc sourceProcedure, statementIDs map[int]bool) bool {
	if proc.Graph == nil || len(statementIDs) == 0 {
		return false
	}
	view := proc.Graph.WithoutNormalErrRaiseContinuationView()
	if !view.IsReachable(view.NormalExit()) {
		return false
	}
	type state struct {
		block   cfg.BlockID
		covered bool
	}
	queue := []state{{block: view.Entry()}}
	seen := map[state]bool{}
	coveredByBlock := map[cfg.BlockID]bool{}
	for statementID := range statementIDs {
		if block, ok := view.BlockForStatement(statementID); ok {
			coveredByBlock[block.ID] = true
		}
	}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		covered := current.covered || coveredByBlock[current.block]
		if current.block == view.NormalExit() {
			if !covered {
				return false
			}
			continue
		}
		view.ForEachOutgoing(current.block, func(edge cfg.Edge) bool {
			nextCovered := covered
			if edge.Class == cfg.EdgeExceptional && !current.covered && coveredByBlock[current.block] {
				// An error raised by the write itself must not count as a
				// completed write on the handler/resume path.
				nextCovered = false
			}
			queue = append(queue, state{block: edge.To, covered: nextCovered})
			return true
		})
	}
	return true
}

func objectContainerObservationCoveredByWrites(proc sourceProcedure, observationStatementID int, statementIDs map[int]bool) bool {
	if proc.Graph == nil || observationStatementID <= 0 || len(statementIDs) == 0 {
		return false
	}
	view := proc.Graph.WithoutNormalErrRaiseContinuationView()
	target, ok := view.BlockForStatement(observationStatementID)
	if !ok || !view.IsReachable(target.ID) {
		return false
	}
	coveredByBlock := map[cfg.BlockID]bool{}
	for statementID := range statementIDs {
		if block, ok := view.BlockForStatement(statementID); ok {
			coveredByBlock[block.ID] = true
		}
	}
	type state struct {
		block   cfg.BlockID
		covered bool
	}
	queue := []state{{block: view.Entry()}}
	seen := map[state]bool{}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		covered := current.covered
		if current.block == target.ID {
			for statementID := range statementIDs {
				block, blockOK := view.BlockForStatement(statementID)
				if blockOK && block.ID == target.ID && statementID < observationStatementID {
					covered = true
				}
			}
			if !covered {
				return false
			}
			continue
		}
		covered = covered || coveredByBlock[current.block]
		view.ForEachOutgoing(current.block, func(edge cfg.Edge) bool {
			nextCovered := covered
			if edge.Class == cfg.EdgeExceptional && !current.covered && coveredByBlock[current.block] {
				nextCovered = false
			}
			queue = append(queue, state{block: edge.To, covered: nextCovered})
			return true
		})
	}
	return true
}

func objectContainerArrayContractForReceiverWithVisiting(index *objectContainerIndex, owner sourceProcedure, receiver, key string, observationStatementID int, visiting map[string]bool) ([]objectContainerElement, bool) {
	if index == nil || receiver == "" || key == "" {
		return nil, false
	}
	visitKey := "array:" + strconv.Itoa(owner.StartLine) + ":" + strings.ToLower(cleanIdentifier(receiver)) + ":" + key
	if visiting[visitKey] {
		return nil, false
	}
	visiting[visitKey] = true
	defer delete(visiting, visitKey)
	ownerDeclarations := objectFlowDeclarations(index.file, owner, index.moduleDecls)
	declaration, scope, ok := objectDeclarationBinding(receiver, ownerDeclarations)
	if !ok || !declaration.Object || dcKindFromType(declaration.Type) != dcDictionary && !strings.EqualFold(cleanIdentifier(declaration.Type), "object") {
		return nil, false
	}
	if !objectContainerDictionaryReceiverKnown(index, owner, receiver, scope, observationStatementID) {
		return nil, false
	}
	if scope == procedureir.ScopeParameter {
		if objectContainerParameterReceiverMutationUnsafe(index, owner, receiver, key, observationStatementID) {
			return nil, false
		}
		return objectContainerParameterArrayContractWithVisiting(index, owner, receiver, key, observationStatementID, visiting)
	}
	existsGuarded := objectContainerExistsGuarded(index, owner, receiver, key, observationStatementID)
	candidates := []sourceProcedure{owner}
	if scope == procedureir.ScopeModule {
		candidates = index.procedures
	}
	found := false
	var elements []objectContainerElement
	for _, proc := range candidates {
		if scope == procedureir.ScopeModule && proc.StartLine != owner.StartLine {
			calledBeforeObservation := objectContainerProcedureCalledBeforeObservation(index, owner, proc, observationStatementID)
			mayReachObservation := objectContainerProcedureMayReachObservation(index, owner, proc, observationStatementID)
			if !calledBeforeObservation && !objectContainerPublicFluentWriter(proc) {
				if !mayReachObservation {
					continue
				}
				if objectContainerProcedureMayMutateReceiver(index, proc, receiver, key) &&
					(!existsGuarded || !objectContainerModuleReceiverIsPrivate(index, receiver) ||
						objectContainerProcedureRemovesKey(index, proc, receiver, key, ownerObservationID(owner, proc, observationStatementID))) {
					return nil, false
				}
			}
		}
		declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
		_, candidateScope, candidateOK := objectDeclarationBinding(receiver, declarations)
		if !candidateOK || candidateScope != scope {
			continue
		}
		flowContext, _, contextOK := objectContainerGraphContext(index, proc)
		if !contextOK {
			continue
		}
		if objectContainerReceiverAliasUnsafe(index, owner, proc, receiver, observationStatementID) {
			return nil, false
		}
		safeWriteIDs := map[int]bool{}
		for statement := range proc.Statements.All() {
			if !objectContainerStatementBeforeObservation(flowContext.graph, statement.ID, ownerObservationID(owner, proc, observationStatementID)) {
				continue
			}
			if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
				continue
			}
			if objectContainerReceiverMemberAssignment(statement, receiver) {
				continue
			}
			target, targetOK := objectContainerSimpleTarget(statement)
			if !targetOK || !strings.EqualFold(target, receiver) || statement.Value == nil {
				continue
			}
			value := strings.ToLower(strings.TrimSpace(statement.Value.Text))
			if statement.Value.Kind != procedureir.ExpressionNew && !strings.HasPrefix(value, "new ") && !objectConstructorCallText(value) {
				return nil, false
			}
		}
		for call := range proc.Calls.All() {
			if !objectContainerStatementBeforeObservation(flowContext.graph, call.StatementID, ownerObservationID(owner, proc, observationStatementID)) {
				continue
			}
			statement, statementOK := objectContainerStatement(proc, call.StatementID)
			defaultReceiver, defaultKey, defaultLiteral, defaultTarget := objectContainerDefaultItemTarget(statement.Text)
			defaultCall := statementOK && defaultTarget && strings.EqualFold(defaultReceiver, receiver) && strings.EqualFold(cleanIdentifier(call.Callee.BaseName), defaultReceiver)
			if !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) && !defaultCall {
				continue
			}
			member := strings.ToLower(cleanIdentifier(call.Callee.Member))
			callKey, keyOK := objectContainerDictionaryKey(proc, call)
			if defaultCall {
				member = "item"
				callKey = defaultKey
				keyOK = defaultLiteral
			}
			switch member {
			case "add":
				if objectErrorResumeNextAt(proc, call.StatementID) {
					return nil, false
				}
				if !keyOK {
					return nil, false
				}
				if !objectContainerKeysEqual(callKey, key) {
					continue
				}
				found = true
				values, safe := objectContainerDirectArrayElements(index, proc, call, declarations)
				if !safe {
					return nil, false
				}
				safeWriteIDs[call.StatementID] = true
				elements = append(elements, values...)
			case "item":
				if objectErrorResumeNextAt(proc, call.StatementID) {
					return nil, false
				}
				statement, statementOK := objectContainerStatement(proc, call.StatementID)
				if !statementOK {
					return nil, false
				}
				matched, assignment := objectContainerDictionaryItemTarget(statement, receiver, key)
				if !keyOK {
					left, _, hasAssignment := strings.Cut(statement.Text, "=")
					if defaultReceiver, _, _, defaultTarget := objectContainerDefaultItemTarget(left); defaultTarget && strings.EqualFold(defaultReceiver, receiver) {
						return nil, false
					}
					leftReceiver, leftKey, literalTarget := objectContainerDictionaryItemExpression(left)
					if !literalTarget {
						if assignment && matched {
							return nil, false
						}
						continue
					}
					if !hasAssignment || !strings.EqualFold(leftReceiver, receiver) || !objectContainerKeysMayAlias(leftKey, key) {
						continue
					}
				} else if !objectContainerKeysMayAlias(callKey, key) {
					continue
				}
				if !assignment || !matched {
					continue
				}
				appendName, appendOK := objectContainerAppendCallName(statement)
				if !appendOK || strings.EqualFold(appendName, "array") {
					if found || statement.Value == nil {
						return nil, false
					}
					names, directOK := objectContainerArrayArgumentNames(statement.Value.Text)
					if !directOK {
						return nil, false
					}
					for _, name := range names {
						objectDeclaration, _, objectDeclared := objectDeclarationBinding(name, declarations)
						if !objectDeclared || !objectDeclaration.Object || !objectContainerDictionaryVariableConstructed(index, proc, name, statement.ID, declarations) {
							return nil, false
						}
						elements = append(elements, objectContainerElement{procedureStart: proc.StartLine, name: cleanIdentifier(name), storageStatementID: statement.ID})
					}
					safeWriteIDs[statement.ID] = true
					found = true
					continue
				}
				var appendCall *procedureir.CallSite
				for candidateCall := range proc.Calls.All() {
					if candidateCall.StatementID == statement.ID && candidateCall.Callee.Receiver == nil && strings.EqualFold(cleanIdentifier(candidateCall.Callee.BaseName), appendName) {
						candidateCopy := candidateCall
						appendCall = &candidateCopy
						break
					}
				}
				if appendCall == nil {
					return nil, false
				}
				appendInfo, appendSafe := objectContainerAppendInfoForResolvedCall(index, proc, *appendCall)
				if !appendSafe {
					return nil, false
				}
				appendArguments, argumentsOK := objectContainerCallArgumentsForProcedure(index, proc, *appendCall, appendName)
				if !argumentsOK {
					return nil, false
				}
				if appendInfo.arrayParameter >= len(appendArguments) || appendInfo.objectParameter >= len(appendArguments) {
					return nil, false
				}
				arrayArgument := cleanIdentifier(appendArguments[appendInfo.arrayParameter])
				objectArgument := cleanIdentifier(appendArguments[appendInfo.objectParameter])
				if arrayArgument == "" || objectArgument == "" || strings.ContainsAny(arrayArgument+objectArgument, ".()") {
					return nil, false
				}
				arrayReceiver, arrayKey, sourceOK := objectContainerArraySource(proc, arrayArgument, statement.ID, flowContext)
				if !sourceOK || !strings.EqualFold(arrayReceiver, receiver) || !objectContainerKeysEqual(arrayKey, key) || !objectContainerDictionaryVariableConstructed(index, proc, objectArgument, statement.ID, declarations) {
					return nil, false
				}
				objectDeclaration, _, objectDeclared := objectDeclarationBinding(objectArgument, declarations)
				if !objectDeclared || !objectDeclaration.Object {
					return nil, false
				}
				elements = append(elements, objectContainerElement{procedureStart: proc.StartLine, name: objectArgument, storageStatementID: statement.ID})
				safeWriteIDs[statement.ID] = true
				found = true
			case "remove":
				if !keyOK || objectContainerKeysMayAlias(callKey, key) {
					return nil, false
				}
			case "removeall":
				return nil, false
			case "exists", "count", "keys":
				// Read-only Dictionary members do not alter the key's value.
			default:
				// An unrecognized receiver call may mutate the Dictionary. Keep
				// the container proof fail-closed until its effects are modeled.
				return nil, false
			}
		}
		if objectContainerReceiverReconstructedAfterWrite(proc, receiver, safeWriteIDs, ownerObservationID(owner, proc, observationStatementID), flowContext.graph) {
			return nil, false
		}
		candidateObservationID := ownerObservationID(owner, proc, observationStatementID)
		if candidateObservationID > 0 && len(safeWriteIDs) > 0 && !objectContainerObservationCoveredByWrites(proc, candidateObservationID, safeWriteIDs) {
			return nil, false
		}
		if len(safeWriteIDs) > 0 && !objectContainerNormalExitCoveredByWrites(proc, safeWriteIDs) {
			return nil, false
		}
	}
	if !found || len(elements) == 0 {
		return nil, false
	}
	return elements, true
}

func objectContainerReceiverReconstructedAfterWrite(proc sourceProcedure, receiver string, writeIDs map[int]bool, observationStatementID int, view cfg.CFGView) bool {
	if len(writeIDs) == 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		if !objectContainerStatementBeforeObservation(view, statement.ID, observationStatementID) ||
			(statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment) || statement.Value == nil {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		if !targetOK || !strings.EqualFold(target, receiver) || !objectConstructorExpression(statement.Value) {
			continue
		}
		for writeID := range writeIDs {
			if statement.ID > writeID && objectContainerStatementCanReach(view, writeID, statement.ID) {
				return true
			}
		}
	}
	return false
}

func objectConstructorExpression(expression *procedureir.Expression) bool {
	if expression == nil {
		return false
	}
	value := strings.ToLower(strings.TrimSpace(expression.Text))
	return expression.Kind == procedureir.ExpressionNew || strings.HasPrefix(value, "new ") || objectConstructorCallText(value)
}

func ownerObservationID(owner, candidate sourceProcedure, observationStatementID int) int {
	if owner.StartLine == candidate.StartLine {
		return observationStatementID
	}
	return 0
}

func objectContainerParameterArrayContract(index *objectContainerIndex, proc sourceProcedure, receiver, key string, observationStatementID int) ([]objectContainerElement, bool) {
	return objectContainerParameterArrayContractWithVisiting(index, proc, receiver, key, observationStatementID, map[string]bool{})
}

func objectContainerParameterArrayContractWithVisiting(index *objectContainerIndex, proc sourceProcedure, receiver, key string, observationStatementID int, visiting map[string]bool) ([]objectContainerElement, bool) {
	visitKey := "parameter:" + strconv.Itoa(proc.StartLine) + ":" + strings.ToLower(cleanIdentifier(receiver)) + ":" + key
	if visiting[visitKey] {
		return nil, false
	}
	visiting[visitKey] = true
	defer delete(visiting, visitKey)
	declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
	_, scope, ok := objectDeclarationBinding(receiver, declarations)
	if !ok {
		return nil, false
	}
	if scope != procedureir.ScopeParameter {
		return objectContainerArrayContractForReceiverWithVisiting(index, proc, receiver, key, observationStatementID, visiting)
	}
	if !strings.EqualFold(strings.TrimSpace(proc.Visibility), "private") {
		return nil, false
	}
	if objectContainerParameterReceiverMutationUnsafe(index, proc, receiver, key, observationStatementID) {
		return nil, false
	}
	parameterIndex := -1
	for index, parameter := range proc.Params.AllIndexed() {
		if strings.EqualFold(cleanIdentifier(parameter.Name), receiver) {
			parameterIndex = index
			break
		}
	}
	if parameterIndex < 0 {
		return nil, false
	}
	var elements []objectContainerElement
	foundCaller := false
	for _, caller := range index.procedures {
		for call := range caller.Calls.All() {
			if call.Callee.Receiver != nil || !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), proc.Name) {
				continue
			}
			if !objectContainerCallResolutionUsable(call) ||
				!objectContainerResolvedCandidateMatches(proc, call.Resolution.Candidates[0]) {
				continue
			}
			callerFlow, _, callerContextOK := objectContainerGraphContext(index, caller)
			if !callerContextOK {
				continue
			}
			callBlock, callBlockOK := callerFlow.graph.BlockForStatement(call.StatementID)
			if !callBlockOK || !callerFlow.graph.IsReachable(callBlock.ID) {
				continue
			}
			parameterNames := make([]string, 0, proc.Params.Len())
			for _, parameter := range proc.Params.AllIndexed() {
				parameterNames = append(parameterNames, parameter.Name)
			}
			arguments, argumentsOK := objectContainerCallArgumentsForParameters(caller, call, parameterNames)
			if !argumentsOK {
				return nil, false
			}
			if parameterIndex >= len(arguments) {
				return nil, false
			}
			actual := cleanIdentifier(arguments[parameterIndex])
			if actual == "" || strings.ContainsAny(actual, ".()") {
				return nil, false
			}
			callerDeclarations := objectFlowDeclarations(index.file, caller, index.moduleDecls)
			actualDeclaration, _, actualOK := objectDeclarationBinding(actual, callerDeclarations)
			if !actualOK || !actualDeclaration.Object {
				return nil, false
			}
			actualElements, safe := objectContainerArrayContractForReceiverWithVisiting(index, caller, actual, key, call.StatementID, visiting)
			if !safe {
				return nil, false
			}
			foundCaller = true
			elements = append(elements, actualElements...)
		}
	}
	if !foundCaller || len(elements) == 0 {
		return nil, false
	}
	return elements, true
}

func objectContainerElementHasCollection(index *objectContainerIndex, element objectContainerElement, key string) bool {
	if index == nil {
		return false
	}
	for _, proc := range index.procedures {
		if proc.StartLine != element.procedureStart {
			continue
		}
		declarations := objectFlowDeclarations(index.file, proc, index.moduleDecls)
		flowContext, _, contextOK := objectContainerGraphContext(index, proc)
		if !contextOK {
			return false
		}
		if !objectContainerDictionaryVariableConstructed(index, proc, element.name, element.storageStatementID, declarations) {
			return false
		}
		addStatementID := 0
		for call := range proc.Calls.All() {
			if !strings.EqualFold(objectCallWithReceiverName(proc, call), element.name) || !strings.EqualFold(cleanIdentifier(call.Callee.Member), "add") {
				continue
			}
			if objectErrorResumeNextAt(proc, call.StatementID) {
				return false
			}
			callKey, ok := objectContainerDictionaryKey(proc, call)
			if !ok {
				return false
			}
			if !objectContainerKeysEqual(callKey, key) {
				continue
			}
			if addStatementID != 0 {
				return false
			}
			arguments := objectContainerCallArguments(proc, call)
			if len(arguments) < 2 {
				return false
			}
			valueName := cleanIdentifier(arguments[1])
			valueDeclaration, _, valueOK := objectDeclarationBinding(valueName, declarations)
			if !valueOK || dcKindFromType(valueDeclaration.Type) != dcCollection || !objectContainerVariableConstructed(index, proc, valueName, call.StatementID, declarations) {
				return false
			}
			addStatementID = call.StatementID
		}
		if addStatementID == 0 || element.storageStatementID <= 0 ||
			!objectContainerStatementDominates(flowContext.graph, addStatementID, element.storageStatementID) {
			return false
		}
		for statement := range proc.Statements.All() {
			if statement.ID == addStatementID || statement.ID == element.storageStatementID ||
				(statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment) {
				continue
			}
			target, targetOK := objectContainerSimpleTarget(statement)
			if targetOK && strings.EqualFold(target, element.name) &&
				objectContainerStatementCanReach(flowContext.graph, addStatementID, statement.ID) &&
				objectContainerStatementCanReach(flowContext.graph, statement.ID, element.storageStatementID) {
				return false
			}
		}
		for statement := range proc.Statements.All() {
			if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment || statement.Value == nil {
				continue
			}
			target, targetOK := objectContainerSimpleTarget(statement)
			if !objectContainerArgumentIsIdentifier(statement.Value.Text, element.name) {
				continue
			}
			if !targetOK || !strings.EqualFold(target, element.name) {
				return false
			}
		}
		for statement := range proc.Statements.All() {
			if statement.ID == addStatementID || !objectContainerStatementCanReach(flowContext.graph, addStatementID, statement.ID) {
				continue
			}
			if target, targetKey, ok := objectContainerDefaultItemExpression(statement.Text); ok &&
				strings.EqualFold(target, element.name) && objectContainerKeysMayAlias(targetKey, key) {
				return false
			}
			if matched, assignment := objectContainerDictionaryItemTarget(statement, element.name, key); assignment && matched {
				return false
			}
		}
		for call := range proc.Calls.All() {
			if call.StatementID == addStatementID || call.StatementID == element.storageStatementID {
				continue
			}
			for _, argument := range objectContainerCallArguments(proc, call) {
				if !objectContainerArgumentIsIdentifier(argument, element.name) {
					continue
				}
				if !objectContainerSafeAppendCallForElement(index, proc, call, element.name) {
					return false
				}
			}
		}
		for call := range proc.Calls.All() {
			if call.StatementID == addStatementID || !strings.EqualFold(objectCallWithReceiverName(proc, call), element.name) {
				continue
			}
			member := strings.ToLower(cleanIdentifier(call.Callee.Member))
			callKey, keyOK := objectContainerDictionaryKey(proc, call)
			switch member {
			case "add":
				if !keyOK || objectContainerKeysMayAlias(callKey, key) {
					return false
				}
			case "item":
				statement, statementOK := objectContainerStatement(proc, call.StatementID)
				if !statementOK {
					return false
				}
				matched, assignment := objectContainerDictionaryItemTarget(statement, element.name, key)
				if assignment && matched {
					return false
				}
			case "exists", "count", "keys":
				// Read-only Dictionary members do not change the stored Collection.
			case "remove":
				if !keyOK || objectContainerKeysMayAlias(callKey, key) {
					return false
				}
			case "removeall":
				return false
			default:
				return false
			}
		}
		return true
	}
	return false
}

func objectContainerConsumerAliasUnsafe(index *objectContainerIndex, proc sourceProcedure, elementName string, observationStatementID int) bool {
	flowContext, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return true
	}
	for statement := range proc.Statements.All() {
		if statement.ID == observationStatementID || !objectContainerStatementBeforeObservation(flowContext.graph, statement.ID, observationStatementID) {
			continue
		}
		if statement.Kind == procedureir.StatementForEach {
			if target, targetOK := objectContainerLoopTarget(statement); targetOK && strings.EqualFold(target, elementName) {
				return true
			}
			continue
		}
		if (statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment) || statement.Value == nil {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		arrayName, _, indexed := objectContainerDirectArrayElement(statement)
		if indexed {
			declarations := objectFlowDeclarations(flowContext.containerIndex.file, proc, flowContext.containerIndex.moduleDecls)
			arrayDeclaration, _, arrayOK := objectDeclarationBinding(arrayName, declarations)
			if arrayOK && arrayDeclaration.Array && (!targetOK || !strings.EqualFold(target, elementName)) {
				return true
			}
		}
		if !objectContainerArgumentIsIdentifier(statement.Value.Text, elementName) {
			continue
		}
		if !targetOK || !strings.EqualFold(target, elementName) {
			return true
		}
	}
	return false
}

func objectContainerConsumerElementPassedToCall(index *objectContainerIndex, proc sourceProcedure, elementName string, observationStatementID int) bool {
	flowContext, _, ok := objectContainerGraphContext(index, proc)
	if !ok {
		return true
	}
	callBlock, ok := flowContext.graph.BlockForStatement(observationStatementID)
	if !ok {
		return true
	}
	dominators := flowContext.graph.Dominators()
	assignmentStatementID := 0
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		if !targetOK || !strings.EqualFold(target, elementName) {
			continue
		}
		block, blockOK := flowContext.graph.BlockForStatement(statement.ID)
		if !blockOK || !objectBlockSetContains(dominators[callBlock.ID], block.ID) || block.ID == callBlock.ID && statement.ID >= observationStatementID {
			continue
		}
		if statement.ID > assignmentStatementID {
			assignmentStatementID = statement.ID
		}
	}
	if assignmentStatementID == 0 {
		return true
	}
	for call := range proc.Calls.All() {
		if call.StatementID == assignmentStatementID ||
			!objectContainerStatementCanReach(flowContext.graph, assignmentStatementID, call.StatementID) ||
			!objectContainerStatementCanReach(flowContext.graph, call.StatementID, observationStatementID) {
			continue
		}
		if objectContainerReflectiveCall(call) {
			return true
		}
		if call.Resolution.Status == procedureir.ResolutionBuiltinLike {
			continue
		}
		for _, argument := range objectContainerCallArguments(proc, call) {
			if objectContainerArgumentIsIdentifier(argument, elementName) {
				return true
			}
		}
	}
	return false
}

func objectContainerConsumerElementReassigned(proc sourceProcedure, receiver string, assignmentStatementID, observationStatementID int, view cfg.CFGView) bool {
	for statement := range proc.Statements.All() {
		if statement.ID == assignmentStatementID || !objectContainerStatementCanReach(view, assignmentStatementID, statement.ID) ||
			!objectContainerStatementCanReach(view, statement.ID, observationStatementID) ||
			(statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment) {
			continue
		}
		target, targetOK := objectContainerSimpleTarget(statement)
		if targetOK && strings.EqualFold(target, receiver) {
			return true
		}
	}
	return false
}

func objectContainerConsumerElementDictionaryMutation(proc sourceProcedure, receiver, key string, assignmentStatementID, observationStatementID int, view cfg.CFGView) bool {
	for statement := range proc.Statements.All() {
		if statement.ID == assignmentStatementID || !objectContainerStatementCanReach(view, assignmentStatementID, statement.ID) ||
			!objectContainerStatementCanReach(view, statement.ID, observationStatementID) {
			continue
		}
		if matched, assignment := objectContainerDictionaryItemTarget(statement, receiver, key); assignment && matched {
			return true
		}
		if target, targetKey, ok := objectContainerDefaultItemExpression(statement.Text); ok && strings.EqualFold(target, receiver) && objectContainerKeysMayAlias(targetKey, key) {
			return true
		}
	}
	for call := range proc.Calls.All() {
		if call.StatementID == assignmentStatementID || !objectContainerStatementCanReach(view, assignmentStatementID, call.StatementID) ||
			!objectContainerStatementCanReach(view, call.StatementID, observationStatementID) || !strings.EqualFold(objectCallWithReceiverName(proc, call), receiver) {
			continue
		}
		member := strings.ToLower(cleanIdentifier(call.Callee.Member))
		callKey, keyOK := objectContainerDictionaryKey(proc, call)
		switch member {
		case "removeall":
			return true
		case "add", "remove":
			if !keyOK || objectContainerKeysMayAlias(callKey, key) {
				return true
			}
		}
	}
	return false
}

func objectContainerConsumerElementDictionaryMutationFromAnyAssignment(proc sourceProcedure, receiver, key string, observationStatementID int, view cfg.CFGView, afterStatementID int) bool {
	for statement := range proc.Statements.All() {
		target, targetOK := objectContainerSimpleTarget(statement)
		if (statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment) || !targetOK ||
			!strings.EqualFold(target, receiver) || !objectContainerStatementBeforeObservation(view, statement.ID, observationStatementID) ||
			afterStatementID > 0 && !objectContainerStatementCanReach(view, afterStatementID, statement.ID) {
			continue
		}
		if objectContainerConsumerElementDictionaryMutation(proc, receiver, key, statement.ID, observationStatementID, view) {
			return true
		}
	}
	return false
}

func objectContainerElementsHaveCollection(index *objectContainerIndex, elements []objectContainerElement, key string, consumer sourceProcedure, receiver string, observationStatementID int) bool {
	if len(elements) == 0 {
		return false
	}
	seen := map[string]bool{}
	for _, element := range elements {
		identity := strconv.Itoa(element.procedureStart) + ":" + strconv.Itoa(element.storageStatementID) + ":" + strings.ToLower(cleanIdentifier(element.name))
		if seen[identity] {
			continue
		}
		seen[identity] = true
		consumerElementName := receiver
		if consumerElementName == "" {
			consumerElementName = element.name
		}
		if objectContainerConsumerAliasUnsafe(index, consumer, consumerElementName, observationStatementID) {
			return false
		}
		if objectContainerConsumerElementPassedToCall(index, consumer, consumerElementName, observationStatementID) {
			return false
		}
		flowContext, _, contextOK := objectContainerGraphContext(index, consumer)
		afterStatementID := 0
		if element.procedureStart == consumer.StartLine {
			afterStatementID = element.storageStatementID
		}
		if !contextOK || objectContainerConsumerElementDictionaryMutationFromAnyAssignment(consumer, consumerElementName, key, observationStatementID, flowContext.graph, afterStatementID) {
			return false
		}
		if !objectContainerElementHasCollection(index, element, key) {
			return false
		}
	}
	return true
}

func objectContainerArrayElementContract(proc sourceProcedure, arrayName string, statementID int, flowContext objectFlowContext, declarations declarationScope) ([]objectContainerElement, bool) {
	if flowContext.containerIndex == nil || objectErrorResumeNextAt(proc, statementID) {
		return nil, false
	}
	receiver, key, ok := objectContainerArraySource(proc, arrayName, statementID, flowContext)
	if !ok {
		return nil, false
	}
	elements, safe := objectContainerParameterArrayContract(flowContext.containerIndex, proc, receiver, key, statementID)
	if !safe || len(elements) == 0 {
		return nil, false
	}
	statement, statementOK := objectContainerStatement(proc, statementID)
	if !statementOK {
		return nil, false
	}
	indexExpression, direct := objectContainerDirectArrayElementIndex(statement, arrayName)
	indexSafe := objectContainerArrayIndexSafe(proc, arrayName, indexExpression, statementID, flowContext.graph, len(elements))
	if !direct || !indexSafe {
		return nil, false
	}
	return elements, true
}

func objectArrayElementAssigned(proc sourceProcedure, statementID int, call procedureir.CallSite, flowContext objectFlowContext, declarations declarationScope) bool {
	if call.Callee.Receiver != nil || call.Arguments.Count == 0 || flowContext.containerIndex == nil || objectErrorResumeNextAt(proc, statementID) {
		return false
	}
	arrayName := cleanIdentifier(call.Callee.BaseName)
	arrayDeclaration, _, ok := objectDeclarationBinding(arrayName, declarations)
	if !ok || !arrayDeclaration.Array {
		return false
	}
	statement, ok := objectContainerStatement(proc, statementID)
	if !ok || statement.Target == nil {
		return false
	}
	targetDeclaration, _, targetOK := objectDeclarationBinding(cleanIdentifier(statement.Target.Text), declarations)
	if !targetOK || !targetDeclaration.Object {
		return false
	}
	elements, safe := objectContainerArrayElementContract(proc, arrayName, statementID, flowContext, declarations)
	return safe && len(elements) > 0
}

func objectDictionaryMemberItemAssigned(proc sourceProcedure, statementID int, call procedureir.CallSite, flowContext objectFlowContext, declarations declarationScope) bool {
	if call.Callee.Receiver == nil || !strings.EqualFold(cleanIdentifier(call.Callee.Member), "item") || call.Arguments.Count == 0 || objectErrorResumeNextAt(proc, statementID) || flowContext.containerIndex == nil {
		return false
	}
	receiver := cleanIdentifier(strings.TrimSpace(*call.Callee.Receiver))
	if receiver == "" || strings.ContainsAny(receiver, ".()") {
		return false
	}
	receiverDeclaration, _, ok := objectDeclarationBinding(receiver, declarations)
	if !ok || !receiverDeclaration.Object || dcKindFromType(receiverDeclaration.Type) != dcDictionary && !strings.EqualFold(cleanIdentifier(receiverDeclaration.Type), "object") {
		return false
	}
	target, targetOK := objectContainerStatement(proc, statementID)
	if !targetOK || target.Target == nil {
		return false
	}
	targetDeclaration, _, targetDeclared := objectDeclarationBinding(cleanIdentifier(target.Target.Text), declarations)
	if !targetDeclared || dcKindFromType(targetDeclaration.Type) != dcCollection {
		return false
	}
	arguments := objectContainerCallArguments(proc, call)
	if len(arguments) != call.Arguments.Count || len(arguments) == 0 {
		return false
	}
	key, keyOK := objectContainerLiteral(arguments[0])
	if !keyOK {
		return false
	}
	elements, ok := objectContainerDictionaryElementSource(proc, receiver, key, statementID, flowContext, declarations)
	return ok && objectContainerElementsHaveCollection(flowContext.containerIndex, elements, key, proc, receiver, statementID)
}

func objectContainerDictionaryElementSource(proc sourceProcedure, receiver, key string, statementID int, flowContext objectFlowContext, declarations declarationScope) ([]objectContainerElement, bool) {
	if flowContext.containerIndex == nil || proc.Graph == nil {
		return nil, false
	}
	callBlock, ok := flowContext.graph.BlockForStatement(statementID)
	if !ok {
		return nil, false
	}
	dominators := flowContext.graph.Dominators()
	var selected procedureir.Statement
	selectedID := -1
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || !strings.EqualFold(objectContainerAssignmentTarget(statement), receiver) {
			continue
		}
		block, ok := flowContext.graph.BlockForStatement(statement.ID)
		if !ok || !objectBlockSetContains(dominators[callBlock.ID], block.ID) || block.ID == callBlock.ID && statement.ID >= statementID {
			continue
		}
		if statement.ID > selectedID {
			selected = statement
			selectedID = statement.ID
		}
	}
	if selectedID < 0 {
		return nil, false
	}
	if objectContainerConsumerElementReassigned(proc, receiver, selectedID, statementID, flowContext.graph) {
		return nil, false
	}
	arrayName, _, direct := objectContainerDirectArrayElement(selected)
	if !direct {
		return nil, false
	}
	arrayDeclaration, _, arrayOK := objectDeclarationBinding(arrayName, declarations)
	if !arrayOK || !arrayDeclaration.Array {
		return nil, false
	}
	result, safe := objectContainerArrayElementContract(proc, arrayName, selected.ID, flowContext, declarations)
	return result, safe && len(result) > 0
}
