package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// objectDictionaryFactoryItemAssigned proves a Dictionary item returned from
// a module field when the key is guarded by Exists and every source write to
// that field stores a known object.  Exists alone is intentionally not enough:
// a Dictionary may contain Nothing or a scalar value, so local or externally
// mutable dictionaries remain conservative.
func objectDictionaryFactoryItemAssigned(proc sourceProcedure, statementID int, call procedureir.CallSite, flowContext objectFlowContext, declarations declarationScope) bool {
	if flowContext.containerIndex == nil || call.Callee.Receiver != nil || call.Arguments.Count == 0 {
		return false
	}
	receiver := cleanIdentifier(call.Callee.BaseName)
	if receiver == "" || !objectDictionaryItemGuarded(proc, statementID, receiver) {
		return false
	}
	declaration, scope, ok := objectDeclarationBinding(receiver, declarations)
	if !ok || scope != procedureir.ScopeModule || !declaration.Object || dcKindFromType(declaration.Type) != dcDictionary {
		return false
	}
	if !objectDictionaryModuleFieldIsPrivate(flowContext.containerIndex.file, declaration) {
		return false
	}
	return objectDictionaryModuleWritesAreObject(flowContext.containerIndex, receiver)
}

func objectDictionaryModuleWritesAreObject(index *objectContainerIndex, receiver string) bool {
	if index == nil || receiver == "" {
		return false
	}
	sawItemWrite := false
	for _, candidate := range index.procedures {
		if !strings.EqualFold(candidate.Module, index.file.Module) {
			continue
		}
		declarations := objectFlowDeclarations(index.file, candidate, index.moduleDecls)
		for statement := range candidate.Statements.All() {
			if !objectDictionaryStatementReachable(candidate, statement.ID) || objectErrorResumeNextAt(candidate, statement.ID) {
				continue
			}
			text := objectCollectionShapeStatementSource(index, statement)
			aliasTarget, aliasValue, aliasSet, aliasOK := objectDictionaryAssignmentText(text)
			if aliasOK && aliasSet {
				aliasTargetPath, targetOK := objectCollectionShapePathText(aliasTarget)
				aliasValuePath, valueOK := objectCollectionShapePathText(objectTrimOuterParens(aliasValue))
				if targetOK && valueOK {
					targetRoot := strings.Split(aliasTargetPath, "|")[0]
					valueRoot := strings.Split(aliasValuePath, "|")[0]
					// A private module field can be handed to another object
					// variable and mutated through that alias.  Reject the
					// factory contract rather than trying to infer all alias
					// lifetimes here.
					if targetRoot != receiver && (valueRoot == receiver || strings.HasPrefix(aliasValuePath, receiver+"|")) {
						return false
					}
					if targetRoot == receiver && valueRoot != receiver {
						return false
					}
				}
			}
			root, indexed, value, set, ok := objectDictionaryAssignmentParts(text)
			if !ok || !strings.EqualFold(root, receiver) {
				continue
			}
			if !indexed {
				if !set || !objectDictionaryConstructorValue(value) {
					return false
				}
				continue
			}
			if !set || !objectDictionaryStoredValueIsObject(candidate, value, declarations) {
				return false
			}
			sawItemWrite = true
		}
		for call := range candidate.Calls.All() {
			if !objectDictionaryStatementReachable(candidate, call.StatementID) || !objectDictionaryCallTouchesReceiver(candidate, call, receiver) {
				continue
			}
			if call.Callee.Receiver != nil {
				switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
				case "exists", "item", "count", "keys", "remove", "removeall":
					continue
				default:
					return false
				}
			}
			if call.Callee.Receiver == nil && call.Arguments.Count > 0 && strings.EqualFold(cleanIdentifier(call.Callee.BaseName), receiver) {
				// The default item read/write is checked from the statement text;
				// the call fact itself has no direction bit.
				continue
			}
			return false
		}
	}
	return sawItemWrite
}

func objectDictionaryModuleFieldIsPrivate(file parsedFile, declaration sourceDeclaration) bool {
	if declaration.Line <= 0 || declaration.Line > len(file.Lines) {
		return false
	}
	line := strings.ToLower(strings.TrimSpace(normalizedCodeLine(file.Lines[declaration.Line-1])))
	return strings.HasPrefix(line, "private ")
}

func objectDictionaryAssignmentParts(text string) (root string, indexed bool, value string, set bool, ok bool) {
	target, value, set, ok := objectDictionaryAssignmentText(text)
	if !ok {
		return "", false, "", false, false
	}
	path, pathOK := objectCollectionShapePathText(target)
	if !pathOK {
		return "", false, "", false, false
	}
	parts := strings.Split(path, "|")
	return parts[0], len(parts) > 1, strings.TrimSpace(value), set, true
}

func objectDictionaryAssignmentText(text string) (target, value string, set, ok bool) {
	text = strings.TrimSpace(text)
	set = strings.HasPrefix(strings.ToLower(text), "set ")
	if set {
		target, value, ok = objectCollectionShapeSetAssignment(text)
	} else {
		target, value, ok = objectCollectionShapeBareAssignment(text)
	}
	return target, value, set, ok
}

func objectDictionaryStoredValueIsObject(proc sourceProcedure, value string, declarations declarationScope) bool {
	lower := strings.ToLower(strings.TrimSpace(value))
	if strings.HasPrefix(lower, "new ") || strings.HasPrefix(lower, "createobject(") || strings.HasPrefix(lower, "getobject(") {
		return true
	}
	name := cleanIdentifier(value)
	if name == "" || strings.ContainsAny(strings.TrimSpace(value), ".()") {
		return false
	}
	declaration, _, ok := objectDeclarationBinding(name, declarations)
	return ok && declaration.Object && declaration.NewExpression
}

func objectDictionaryStatementReachable(proc sourceProcedure, statementID int) bool {
	if proc.Graph == nil {
		return false
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	block, ok := graph.BlockForStatement(statementID)
	return ok && graph.IsReachable(block.ID)
}

func objectDictionaryCallTouchesReceiver(proc sourceProcedure, call procedureir.CallSite, receiver string) bool {
	if call.Callee.Receiver != nil && strings.EqualFold(cleanIdentifier(strings.TrimSpace(*call.Callee.Receiver)), receiver) {
		return true
	}
	for _, actual := range objectCallActuals(call, proc.analysisFacts()) {
		if strings.EqualFold(cleanIdentifier(actual.text), receiver) {
			return true
		}
	}
	return call.Callee.Receiver == nil && strings.EqualFold(cleanIdentifier(call.Callee.BaseName), receiver) && call.Arguments.Count > 0
}

func objectDictionaryConstructorValue(value string) bool {
	lower := strings.ToLower(strings.TrimSpace(value))
	return strings.HasPrefix(lower, "new ") || strings.HasPrefix(lower, "createobject(\"scripting.dictionary\"")
}
