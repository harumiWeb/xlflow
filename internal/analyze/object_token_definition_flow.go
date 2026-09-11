package analyze

import (
	"strings"
)

// objectCollectionShapeApplyKnownFactory records the object field established
// by the narrow TokenDefinition array factory used by stdLambda.  A function
// name alone is not enough evidence: the factory and its element constructor
// must both be unique in the module and have the expected unconditional source
// shape before the wildcard element path is admitted.
func objectCollectionShapeApplyKnownFactory(index *objectContainerIndex, proc sourceProcedure, state *objectCollectionShapeState, path, value string) bool {
	if !objectCollectionShapeCallWithoutArguments(value, "getTokenDefinitions") || !objectTokenDefinitionsFactoryProven(index, proc) {
		return false
	}
	objectCollectionShapeDelete(state, path)
	state.values[path] = true
	state.regexp[objectCollectionShapePath(path, "*", "regexobj")] = true
	return true
}

func objectCollectionShapeCallWithoutArguments(text, name string) bool {
	text = strings.TrimSpace(text)
	open := strings.IndexByte(text, '(')
	if open <= 0 || !strings.EqualFold(cleanIdentifier(text[:open]), cleanIdentifier(name)) {
		return false
	}
	inside, end, ok := dcBalancedContentSpan(text[open:])
	return ok && strings.TrimSpace(inside) == "" && strings.TrimSpace(text[open+end:]) == ""
}

func objectTokenDefinitionsFactoryProven(index *objectContainerIndex, caller sourceProcedure) bool {
	if index == nil {
		return false
	}
	definitions, ok := objectTokenDefinitionProcedure(index, caller.Module, "getTokenDefinitions")
	if !ok || !procedureReturnsArray(definitions) {
		return false
	}
	tokenDefinition, ok := objectTokenDefinitionProcedure(index, caller.Module, "getTokenDefinition")
	if !ok || !objectTokenDefinitionFactoryProven(index, tokenDefinition) {
		return false
	}

	arrayName := ""
	redimNames := map[string]bool{}
	returnIDs := map[int]bool{}
	elementWrites := make([]objectTokenDefinitionSourcePart, 0)
	for _, source := range objectTokenDefinitionSourceParts(index, definitions) {
		if objectTokenDefinitionUnsafeControl(source.text) {
			return false
		}
		lower := strings.ToLower(strings.TrimSpace(source.text))
		if strings.HasPrefix(lower, "redim ") {
			if name := objectTokenDefinitionRedimName(source.text); name != "" {
				redimNames[name] = true
			}
			continue
		}
		if lhs, rhs, assigned := arrayAssignmentSides(source.text); assigned && !strings.ContainsAny(lhs, "().") && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(definitions.Name)) {
			candidate := cleanIdentifier(rhs)
			if candidate == "" || strings.ContainsAny(candidate, "().") {
				return false
			}
			if arrayName != "" && !strings.EqualFold(arrayName, candidate) {
				return false
			}
			arrayName = candidate
			returnIDs[source.statementID] = true
		}
		if receiver, _, rhs, assigned := arrayMemberAssignmentParts(source.text); assigned {
			elementWrites = append(elementWrites, objectTokenDefinitionSourcePart{
				statementID: source.statementID,
				text:        source.text,
				receiver:    receiver,
				rhs:         rhs,
			})
		}
	}
	if arrayName == "" || !redimNames[strings.ToLower(arrayName)] || len(returnIDs) != 1 || len(elementWrites) == 0 {
		return false
	}
	writeIDs := map[int]bool{}
	for _, write := range elementWrites {
		if !strings.EqualFold(cleanIdentifier(write.receiver), cleanIdentifier(arrayName)) || arrayCallName(write.rhs) != "gettokendefinition" || objectErrorResumeNextAt(definitions, write.statementID) {
			return false
		}
		writeIDs[write.statementID] = true
	}
	return objectContainerNormalExitCoveredByWrites(definitions, writeIDs) && objectContainerNormalExitCoveredByWrites(definitions, returnIDs)
}

func objectTokenDefinitionFactoryProven(index *objectContainerIndex, procedure sourceProcedure) bool {
	if !strings.EqualFold(strings.ReplaceAll(strings.TrimSpace(procedure.ReturnType), " ", ""), "tokendefinition") {
		return false
	}
	regexWrites := map[int]bool{}
	for _, source := range objectTokenDefinitionSourceParts(index, procedure) {
		if objectTokenDefinitionUnsafeControl(source.text) {
			return false
		}
		target, value, assigned := objectCollectionShapeSetAssignment(source.text)
		if !assigned {
			continue
		}
		path, pathOK := objectCollectionShapePathText(target)
		if !pathOK || path != objectCollectionShapePath(procedure.Name, "regexobj") {
			continue
		}
		if objectErrorResumeNextAt(procedure, source.statementID) || !strings.HasPrefix(strings.ToLower(strings.TrimSpace(value)), "createobject(\"vbscript.regexp\"") {
			return false
		}
		regexWrites[source.statementID] = true
	}
	return len(regexWrites) == 1
}

type objectTokenDefinitionSourcePart struct {
	statementID int
	text        string
	receiver    string
	rhs         string
}

func objectTokenDefinitionProcedure(index *objectContainerIndex, module, name string) (sourceProcedure, bool) {
	var match sourceProcedure
	found := false
	for _, candidate := range index.procedures {
		if !strings.EqualFold(candidate.Module, module) || !strings.EqualFold(cleanIdentifier(candidate.Name), cleanIdentifier(name)) {
			continue
		}
		if found {
			return sourceProcedure{}, false
		}
		match = candidate
		found = true
	}
	return match, found
}

func objectTokenDefinitionSourceParts(index *objectContainerIndex, procedure sourceProcedure) []objectTokenDefinitionSourcePart {
	parts := make([]objectTokenDefinitionSourcePart, 0)
	for statement := range procedure.Statements.All() {
		text := objectCollectionShapeStatementSource(index, statement)
		text = arraySourceOrderStripComment(text)
		for _, part := range splitRangeValueSourceStatements(text) {
			part = strings.TrimSpace(part)
			if part != "" {
				parts = append(parts, objectTokenDefinitionSourcePart{statementID: statement.ID, text: part})
			}
		}
	}
	return parts
}

func objectTokenDefinitionRedimName(text string) string {
	text = strings.TrimSpace(text)
	text = strings.TrimSpace(text[len("ReDim "):])
	if strings.HasPrefix(strings.ToLower(text), "preserve ") {
		text = strings.TrimSpace(text[len("Preserve "):])
	}
	open := strings.IndexByte(text, '(')
	if open <= 0 {
		return ""
	}
	return strings.ToLower(cleanIdentifier(strings.TrimSpace(text[:open])))
}

func objectTokenDefinitionUnsafeControl(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{
		"if ", "elseif ", "else", "for ", "next", "do", "loop", "while ", "wend",
		"select case", "case ", "goto ", "on error", "exit ",
	} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}
