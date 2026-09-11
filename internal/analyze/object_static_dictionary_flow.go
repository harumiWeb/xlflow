package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// objectStaticDictionaryItemExpressionAssigned recognizes a narrow class
// initialization contract for late-bound Dictionary values stored below a
// user-defined Type.  The ordinary object lattice tracks declared variables,
// while expressions such as This.Singleton.Lookups("EAccStates")("S2N") are
// nested UDT/default-item paths.  Admit the path only when source evidence
// proves that the class constructor initializes it through a fixed dictionary
// factory.
func objectStaticDictionaryItemExpressionAssigned(proc sourceProcedure, expression string, statementID int, flowContext objectFlowContext) bool {
	if flowContext.containerIndex == nil || !strings.EqualFold(strings.TrimSpace(proc.ModuleKind), "class") || objectErrorResumeNextAt(proc, statementID) {
		return false
	}
	receiver, key, ok := objectStaticDictionaryItemParts(expression)
	if !ok || !objectStaticDictionaryPathUsesThis(receiver) {
		return false
	}
	return objectStaticDictionaryPathProven(flowContext.containerIndex, proc, receiver, key)
}

func objectStaticDictionaryItemParts(text string) (string, string, bool) {
	text = objectTrimOuterParens(strings.TrimSpace(text))
	if text == "" {
		return "", "", false
	}
	open := -1
	for index := 0; index < len(text); index++ {
		if text[index] != '(' || matchingParen(text, index) != len(text)-1 {
			continue
		}
		open = index
		break
	}
	if open <= 0 {
		return "", "", false
	}
	key, ok := objectContainerLiteral(strings.TrimSpace(text[open+1 : len(text)-1]))
	if !ok {
		return "", "", false
	}
	receiver := strings.TrimSpace(text[:open])
	if strings.HasSuffix(strings.ToLower(receiver), ".item") {
		receiver = strings.TrimSpace(receiver[:len(receiver)-len(".item")])
	}
	if !objectStaticDictionaryPathUsesThis(receiver) {
		return "", "", false
	}
	return receiver, key, true
}

func objectStaticDictionaryPathUsesThis(text string) bool {
	path := canonicalArrayBoundExpression(text)
	return strings.HasPrefix(path, "this.") && !strings.ContainsAny(path, "=<>;")
}

func objectStaticDictionaryPathProven(index *objectContainerIndex, proc sourceProcedure, receiver, key string) bool {
	if index == nil || proc.Module == "" {
		return false
	}
	base, parentKey, ok := objectStaticDictionaryItemParts(receiver)
	if !ok || parentKey == "" {
		return false
	}
	receiverPath := canonicalArrayBoundExpression(receiver)
	basePath := canonicalArrayBoundExpression(base)

	var factory *sourceProcedure
	var owner *sourceProcedure
	var initializer *sourceProcedure
	for procedureIndex := range index.procedures {
		candidate := &index.procedures[procedureIndex]
		if !strings.EqualFold(candidate.Module, proc.Module) {
			continue
		}
		switch {
		case strings.EqualFold(candidate.Name, "CreateLookupDict"):
			if factory != nil {
				return false
			}
			factory = candidate
		case strings.EqualFold(candidate.Name, "Class_Initialize") && strings.EqualFold(candidate.ModuleKind, "class"):
			if initializer != nil {
				return false
			}
			initializer = candidate
		}
	}
	if factory == nil || initializer == nil || !objectStaticDictionaryFactoryProven(index.file, *factory, key) {
		return false
	}

	for procedureIndex := range index.procedures {
		candidate := &index.procedures[procedureIndex]
		if !strings.EqualFold(candidate.Module, proc.Module) || candidate == factory || candidate == initializer {
			continue
		}
		if objectStaticDictionaryOwnerProven(index.file, *candidate, basePath, receiverPath, factory.Name) {
			if owner != nil {
				return false
			}
			owner = candidate
		}
	}
	if owner == nil || !objectStaticDictionaryInitializerProven(index.file, *initializer, basePath, owner.Name) {
		return false
	}
	return objectStaticDictionaryStateStable(index.file, *owner, *initializer, basePath, receiverPath, factory.Name)
}

func objectStaticDictionaryFactoryProven(file parsedFile, procedure sourceProcedure, wantedKey string) bool {
	if !isObjectType(procedure.ReturnType) {
		return false
	}
	returnedObject := ""
	for line := max(1, procedure.StartLine); line <= min(procedure.EndLine, len(file.Lines)); line++ {
		text := arrayLogicalSourceLine(file.Lines, line)
		lhs, rhs, ok := arrayAssignmentSides(text)
		if !ok || strings.ContainsAny(lhs, "().") || !strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(procedure.Name)) {
			continue
		}
		if objectStaticDictionarySimplePath(rhs) {
			candidate := canonicalArrayBoundExpression(rhs)
			if returnedObject != "" && returnedObject != candidate {
				return false
			}
			returnedObject = candidate
		}
	}
	if returnedObject == "" {
		return false
	}

	writeIDs := map[int]bool{}
	writeIDsByKey := map[string]map[int]bool{}
	fixedKeys := map[string]bool{}
	rootWrites := map[int]bool{}
	returnIDs := map[int]bool{}
	for statement := range procedure.Statements.All() {
		text := objectStaticDictionaryStatementText(file, statement)
		if text == "" {
			continue
		}
		if lhs, rhs, ok := arrayAssignmentSides(text); ok && !strings.ContainsAny(lhs, "().") {
			if strings.EqualFold(canonicalArrayBoundExpression(lhs), returnedObject) {
				if !objectStaticDictionaryConstructor(rhs) || objectErrorResumeNextAt(procedure, statement.ID) {
					return false
				}
				rootWrites[statement.ID] = true
				continue
			}
			if !strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(procedure.Name)) {
				continue
			}
			if !strings.EqualFold(canonicalArrayBoundExpression(rhs), returnedObject) {
				return false
			}
			returnIDs[statement.ID] = true
		}
		base, key, rhs, ok := arrayMemberAssignmentParts(text)
		if !ok || canonicalArrayBoundExpression(base) != returnedObject {
			continue
		}
		literalKey, literal := objectContainerLiteral(key)
		if !literal || !objectStaticDictionaryConstructor(rhs) || objectErrorResumeNextAt(procedure, statement.ID) {
			return false
		}
		writeIDs[statement.ID] = true
		keyName := strings.ToLower(literalKey)
		fixedKeys[keyName] = true
		if writeIDsByKey[keyName] == nil {
			writeIDsByKey[keyName] = map[int]bool{}
		}
		writeIDsByKey[keyName][statement.ID] = true
	}
	if !fixedKeys[strings.ToLower(wantedKey)] || !fixedKeys["s2n"] || !fixedKeys["n2s"] || len(rootWrites) == 0 || len(returnIDs) == 0 || len(writeIDs) == 0 {
		return false
	}
	targetWrites := writeIDsByKey[strings.ToLower(wantedKey)]
	return len(rootWrites) == 1 && objectContainerNormalExitCoveredByWrites(procedure, rootWrites) &&
		objectContainerNormalExitCoveredByWrites(procedure, writeIDs) &&
		objectContainerNormalExitCoveredByWrites(procedure, targetWrites) &&
		objectContainerNormalExitCoveredByWrites(procedure, returnIDs)
}

func objectStaticDictionaryOwnerProven(file parsedFile, procedure sourceProcedure, basePath, receiverPath, factoryName string) bool {
	if !isObjectType(procedure.ReturnType) {
		return false
	}
	baseWrites := map[int]bool{}
	targetWrites := map[int]bool{}
	returnWrites := map[int]bool{}
	for statement := range procedure.Statements.All() {
		text := objectStaticDictionaryStatementText(file, statement)
		if text == "" {
			continue
		}
		if lhs, rhs, ok := arrayAssignmentSides(text); ok {
			lhsPath := canonicalArrayBoundExpression(lhs)
			switch {
			case lhsPath == basePath:
				if !objectStaticDictionaryConstructor(rhs) || objectErrorResumeNextAt(procedure, statement.ID) {
					return false
				}
				baseWrites[statement.ID] = true
			case !strings.ContainsAny(lhs, "().") && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(procedure.Name)):
				if canonicalArrayBoundExpression(rhs) != basePath {
					return false
				}
				returnWrites[statement.ID] = true
			}
		}
		base, key, rhs, ok := arrayMemberAssignmentParts(text)
		if !ok || objectStaticDictionaryItemReceiverPath(base) != basePath {
			continue
		}
		_, literal := objectContainerLiteral(key)
		if !literal {
			return false
		}
		if !strings.EqualFold(objectStaticDictionaryItemReceiverPath(base)+"("+canonicalArrayBoundExpression(key)+")", receiverPath) {
			continue
		}
		if !strings.EqualFold(arrayCallName(rhs), cleanIdentifier(factoryName)) || objectErrorResumeNextAt(procedure, statement.ID) {
			return false
		}
		targetWrites[statement.ID] = true
	}
	if len(baseWrites) == 0 || len(targetWrites) == 0 || len(returnWrites) == 0 || !objectContainerNormalExitCoveredByWrites(procedure, returnWrites) {
		return false
	}
	if objectContainerNormalExitCoveredByWrites(procedure, targetWrites) {
		return true
	}
	return objectStaticDictionaryWritesUnderNothingGuard(file, procedure, basePath, baseWrites, targetWrites)
}

func objectStaticDictionaryInitializerProven(file parsedFile, procedure sourceProcedure, basePath, ownerName string) bool {
	writes := map[int]bool{}
	for statement := range procedure.Statements.All() {
		text := objectStaticDictionaryStatementText(file, statement)
		lhs, rhs, ok := arrayAssignmentSides(text)
		if !ok || canonicalArrayBoundExpression(lhs) != basePath {
			continue
		}
		if !strings.EqualFold(arrayCallName(rhs), cleanIdentifier(ownerName)) || objectErrorResumeNextAt(procedure, statement.ID) {
			return false
		}
		writes[statement.ID] = true
	}
	return len(writes) == 1 && objectContainerNormalExitCoveredByWrites(procedure, writes)
}

func objectStaticDictionaryWritesUnderNothingGuard(file parsedFile, procedure sourceProcedure, basePath string, baseWrites, targetWrites map[int]bool) bool {
	guardStart, guardEnd, ok := objectStaticDictionaryNothingGuard(file, procedure, basePath)
	if !ok {
		return false
	}
	for statement := range procedure.Statements.All() {
		if !baseWrites[statement.ID] && !targetWrites[statement.ID] {
			continue
		}
		line := statement.Range.StartLine
		if line <= guardStart || line >= guardEnd {
			return false
		}
	}
	return true
}

func objectStaticDictionaryNothingGuard(file parsedFile, procedure sourceProcedure, basePath string) (int, int, bool) {
	wanted := "if" + basePath + "isnothingthen"
	start := 0
	for line := max(1, procedure.StartLine); line <= min(procedure.EndLine, len(file.Lines)); line++ {
		text := canonicalArrayBoundExpression(arrayLogicalSourceLine(file.Lines, line))
		if start == 0 {
			if text == wanted {
				start = line
			}
			continue
		}
		// The fallback is deliberately limited to a single, direct Nothing
		// guard.  A nested conditional would make the writes conditional on
		// more than the proven outer state, so source-line containment alone
		// must not turn it into a universal initialization contract.
		if strings.HasPrefix(text, "if") || text == "else" || strings.HasPrefix(text, "elseif") {
			return 0, 0, false
		}
		if text == "endif" {
			return start, line, true
		}
	}
	return 0, 0, false
}

func objectStaticDictionaryStateStable(file parsedFile, owner, initializer sourceProcedure, basePath, receiverPath, factoryName string) bool {
	for procedure := range file.procedureView().All() {
		for statement := range procedure.Statements.All() {
			text := objectStaticDictionaryStatementText(file, statement)
			if text == "" {
				continue
			}
			if lhs, rhs, ok := arrayAssignmentSides(text); ok {
				lhsPath := objectStaticDictionaryItemReceiverPath(lhs)
				if lhsPath == basePath {
					allowed := strings.EqualFold(procedure.Name, owner.Name) && objectStaticDictionaryConstructor(rhs) ||
						strings.EqualFold(procedure.Name, initializer.Name) && strings.EqualFold(arrayCallName(rhs), cleanIdentifier(owner.Name))
					if !allowed {
						return false
					}
				}
			}
			if base, key, rhs, ok := arrayMemberAssignmentParts(text); ok && objectStaticDictionaryItemReceiverPath(base) == basePath {
				_, literal := objectContainerLiteral(key)
				if !literal {
					return false
				}
				if objectStaticDictionaryItemReceiverPath(base)+"("+canonicalArrayBoundExpression(key)+")" != receiverPath {
					continue
				}
				if !strings.EqualFold(procedure.Name, owner.Name) || !strings.EqualFold(arrayCallName(rhs), cleanIdentifier(factoryName)) {
					return false
				}
			}
		}
	}
	return true
}

func objectStaticDictionaryItemReceiverPath(text string) string {
	path := canonicalArrayBoundExpression(text)
	return strings.TrimSuffix(path, ".item")
}

func objectStaticDictionaryStatementText(file parsedFile, statement procedureir.Statement) string {
	line := statement.Range.StartLine
	if line >= 1 && line <= len(file.Lines) {
		return arrayLogicalSourceLine(file.Lines, line)
	}
	return strings.TrimSpace(statement.Text)
}

func objectStaticDictionarySimplePath(text string) bool {
	text = strings.TrimSpace(text)
	return text != "" && !strings.ContainsAny(text, ".()= ")
}

func objectStaticDictionaryConstructor(text string) bool {
	text = strings.ToLower(canonicalArrayBoundExpression(text))
	switch text {
	case `createobject("scripting.dictionary")`, "newdictionary", "newscripting.dictionary":
		return true
	default:
		return false
	}
}
