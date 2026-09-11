package analyze

import (
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// objectFlowApplyRepeatedSelectCaseObjectState restores an object proof that
// the ordinary must-analysis loses when two Select Case statements inspect an
// unchanged selector.  A no-match edge necessarily remains nullable, but a
// later matching Case is correlated with the earlier matching Case when the
// source proves an unconditional object factory on that earlier branch.
func objectFlowApplyRepeatedSelectCaseObjectState(proc sourceProcedure, state map[string]bool, flowContext objectFlowContext, edge vbacfg.Edge, declarations declarationScope) map[string]bool {
	if state == nil || edge.Kind != vbacfg.EdgeCase || flowContext.graph.BlockCount() == 0 {
		return state
	}
	caseBlock, ok := flowContext.graph.BlockByID(edge.To)
	if !ok || caseBlock.Statement == nil || caseBlock.Statement.Kind != procedureir.StatementCase {
		return state
	}
	currentCase := *caseBlock.Statement
	if currentCase.Control != nil && currentCase.Control.CaseElse {
		return state
	}
	currentSelect := procedureStatementByID(proc, currentCase.ParentID)
	if currentSelect.ID == 0 || currentSelect.Kind != procedureir.StatementSelect {
		return state
	}
	selector := selectCaseExpression(currentSelect.Text)
	selectorKey := objectRepeatedSelectCaseSelectorKey(selector)
	if selectorKey == "" || !objectRepeatedSelectCaseSelectorStable(proc, selector, selectorKey) {
		return state
	}
	caseLabel := objectRepeatedSelectCaseLabel(currentCase.Text)
	if caseLabel == "" {
		return state
	}
	previousSelect, previousCase, ok := objectRepeatedSelectCasePrevious(proc, currentSelect, selectorKey, caseLabel)
	if !ok {
		return state
	}
	proven := objectRepeatedSelectCaseFactoryAssignments(proc, previousCase, flowContext, declarations)
	if flowContext.containerIndex == nil || len(proven) == 0 || !objectRepeatedSelectCaseRegionStable(proc, flowContext.containerIndex.file, selectorKey, previousSelect, currentSelect, currentCase, previousCase, proven, declarations, flowContext.objectTypeNames) {
		return state
	}
	updated := state
	cloned := false
	for key := range proven {
		if updated[key] {
			continue
		}
		if !cloned {
			updated = cloneObjectState(state)
			cloned = true
		}
		updated[key] = true
	}
	return updated
}

func objectRepeatedSelectCaseSelectorKey(expression string) string {
	if path, ok := objectCollectionShapePathText(expression); ok {
		return path
	}
	return strings.ToLower(strings.Join(strings.Fields(strings.TrimSpace(expression)), ""))
}

func objectRepeatedSelectCaseSelectorStable(proc sourceProcedure, selector, selectorKey string) bool {
	selector = strings.TrimSpace(selector)
	if selector == "" || selectorKey == "" {
		return false
	}
	if isIdentifier(selector) {
		for parameter := range proc.Params.All() {
			if !strings.EqualFold(cleanIdentifier(parameter.Name), selector) {
				continue
			}
			return strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
		}
		return false
	}
	parts := objectMemberChainParts(selector)
	if len(parts) < 2 {
		return false
	}
	root := strings.ToLower(cleanIdentifier(parts[0]))
	return root == "this" || root == "me"
}

func objectRepeatedSelectCaseLabel(text string) string {
	first := strings.TrimSpace(strings.SplitN(text, "\n", 2)[0])
	if !strings.HasPrefix(strings.ToLower(first), "case ") {
		return ""
	}
	return strings.ToLower(strings.Join(strings.Fields(first), ""))
}

func objectRepeatedSelectCasePrevious(proc sourceProcedure, currentSelect procedureir.Statement, selectorKey, caseLabel string) (procedureir.Statement, procedureir.Statement, bool) {
	var previousSelect procedureir.Statement
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSelect || statement.ID == currentSelect.ID || statement.ParentID != currentSelect.ParentID || statement.Range.EndLine >= currentSelect.Range.StartLine {
			continue
		}
		if objectRepeatedSelectCaseSelectorKey(selectCaseExpression(statement.Text)) != selectorKey {
			continue
		}
		if previousSelect.ID == 0 || statement.Range.StartLine > previousSelect.Range.StartLine {
			previousSelect = statement
		}
	}
	if previousSelect.ID == 0 {
		return procedureir.Statement{}, procedureir.Statement{}, false
	}
	var previousCase procedureir.Statement
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementCase || statement.ParentID != previousSelect.ID || statement.Control != nil && statement.Control.CaseElse || objectRepeatedSelectCaseLabel(statement.Text) != caseLabel {
			continue
		}
		if previousCase.ID != 0 {
			return procedureir.Statement{}, procedureir.Statement{}, false
		}
		previousCase = statement
	}
	return previousSelect, previousCase, previousCase.ID != 0
}

func objectRepeatedSelectCaseRegionStable(proc sourceProcedure, file parsedFile, selectorKey string, previousSelect, currentSelect, currentCase, previousCase procedureir.Statement, proven map[string]bool, declarations declarationScope, objectTypeNames map[string]bool) bool {
	if !objectRepeatedSelectCaseFlowStable(proc, previousSelect, currentCase) {
		return false
	}
	if previousCase.ID == 0 || !objectRepeatedSelectCaseLinesStable(proc, selectorKey, previousCase.Range.StartLine+1, previousCase.Range.EndLine-1, nil, declarations, objectTypeNames) {
		return false
	}
	return objectRepeatedSelectCaseRegionStraightLine(file, previousSelect.Range.EndLine+1, currentSelect.Range.StartLine-1) &&
		objectRepeatedSelectCaseLinesStable(proc, selectorKey, previousSelect.Range.EndLine+1, currentSelect.Range.StartLine-1, proven, declarations, objectTypeNames)
}

func objectRepeatedSelectCaseFlowStable(proc sourceProcedure, previousSelect, currentCase procedureir.Statement) bool {
	if proc.Graph == nil || previousSelect.ID == 0 || currentCase.ID == 0 {
		return false
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	previousBlock, previousOK := graph.BlockForStatement(previousSelect.ID)
	currentBlock, currentOK := graph.BlockForStatement(currentCase.ID)
	if !previousOK || !currentOK {
		return false
	}
	// A Case arm is an alternative CFG entry, so the previous matching Case
	// itself cannot dominate the later arm. Require the previous Select block
	// to dominate the later matching Case instead; this rejects GoTo paths that
	// enter the later arm without evaluating the earlier selector.
	return graph.Dominates(previousBlock.ID, currentBlock.ID)
}

func objectRepeatedSelectCaseLinesStable(proc sourceProcedure, selectorKey string, start, end int, objectTargets map[string]bool, declarations declarationScope, objectTypeNames map[string]bool) bool {
	for statement := range proc.Statements.All() {
		if statement.Range.EndLine < start || statement.Range.StartLine > end {
			continue
		}
		if statement.Kind == procedureir.StatementCall || statement.Kind == procedureir.StatementOnError {
			return false
		}
		text := strings.TrimSpace(statement.Text)
		for _, assignment := range []func(string) (string, string, bool){objectCollectionShapeSetAssignment, objectCollectionShapeBareAssignment} {
			target, _, ok := assignment(text)
			if !ok {
				continue
			}
			if objectRepeatedSelectCaseSelectorWrite(target, selectorKey) {
				return false
			}
			if len(objectTargets) > 0 && isIdentifier(target) {
				declaration, scope, declared := objectDeclarationBinding(target, declarations)
				if declared && objectDeclarationIsKnownObject(declaration, objectTypeNames) && objectTargets[(objectVariable{Scope: scope, Name: target}).key()] {
					return false
				}
			}
		}
	}
	return true
}

func objectRepeatedSelectCaseSelectorWrite(target, selectorKey string) bool {
	targetKey := objectRepeatedSelectCaseSelectorKey(target)
	if targetKey == "" || targetKey == selectorKey {
		return targetKey != ""
	}
	return objectCollectionShapePathHasPrefix(selectorKey, targetKey)
}

func objectRepeatedSelectCaseFactoryAssignments(proc sourceProcedure, branch procedureir.Statement, flowContext objectFlowContext, declarations declarationScope) map[string]bool {
	if branch.ID == 0 || flowContext.containerIndex == nil || !objectRepeatedSelectCaseSourceStraightLine(flowContext.containerIndex.file, branch) {
		return nil
	}
	proven := map[string]bool{}
	invalid := map[string]bool{}
	for statement := range proc.Statements.All() {
		if statement.ParentID != branch.ID || statement.Kind == procedureir.StatementCase {
			continue
		}
		if statement.Kind == procedureir.StatementCall {
			return nil
		}
		text := objectCollectionShapeStatementSource(flowContext.containerIndex, statement)
		target, value, setAssignment := objectCollectionShapeSetAssignment(text)
		if !setAssignment {
			target, value, setAssignment = objectCollectionShapeBareAssignment(text)
		}
		if !setAssignment {
			continue
		}
		if !isIdentifier(target) {
			continue
		}
		declaration, scope, declared := objectDeclarationBinding(target, declarations)
		if !declared || !objectDeclarationIsKnownObject(declaration, flowContext.objectTypeNames) {
			continue
		}
		key := (objectVariable{Scope: scope, Name: target}).key()
		if objectRepeatedSelectCaseObjectFactoryProven(proc, value, flowContext) {
			if !invalid[key] {
				proven[key] = true
			}
			continue
		}
		invalid[key] = true
		delete(proven, key)
	}
	for key := range invalid {
		delete(proven, key)
	}
	return proven
}

func objectRepeatedSelectCaseObjectFactoryProven(proc sourceProcedure, value string, flowContext objectFlowContext) bool {
	value = strings.ToLower(strings.TrimSpace(value))
	if strings.HasPrefix(value, "new ") || strings.HasPrefix(value, "createobject(") || strings.HasPrefix(value, "getobject(") {
		return true
	}
	name, args, ok := objectCollectionShapeBareCall(value)
	if !ok || len(args) != 0 {
		return false
	}
	return objectBareObjectFunctionAssigned(proc, name, flowContext.summaries)
}

func objectRepeatedSelectCaseSourceStraightLine(file parsedFile, branch procedureir.Statement) bool {
	start := max(branch.Range.StartLine+1, 1)
	end := min(branch.Range.EndLine-1, len(file.Lines))
	for line := start; line <= end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		lower := strings.ToLower(text)
		if text == "" || strings.HasPrefix(text, "'") {
			continue
		}
		if strings.HasPrefix(text, "#") {
			return false
		}
		if objectRepeatedSelectCaseControlLine(lower) || strings.HasSuffix(lower, ":") {
			return false
		}
	}
	return true
}

func objectRepeatedSelectCaseRegionStraightLine(file parsedFile, start, end int) bool {
	start = max(start, 1)
	end = min(end, len(file.Lines))
	for line := start; line <= end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		lower := strings.ToLower(text)
		if text == "" || strings.HasPrefix(text, "'") || lower == "end select" || strings.HasPrefix(lower, "case ") {
			continue
		}
		if strings.HasPrefix(text, "#") || strings.HasSuffix(lower, ":") {
			return false
		}
		if objectRepeatedSelectCaseControlLine(lower) {
			return false
		}
	}
	return true
}

func objectRepeatedSelectCaseControlLine(lower string) bool {
	for _, keyword := range []string{
		"if", "elseif", "else", "end if", "for", "next", "do", "loop", "while", "wend",
		"select", "case", "goto", "exit", "on error", "resume", "with", "end with", "end select",
	} {
		if lower == keyword || strings.HasPrefix(lower, keyword+" ") {
			return true
		}
	}
	return false
}

func isIdentifier(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" || !isIdentifierStart(text[0]) {
		return false
	}
	for index := 1; index < len(text); index++ {
		if !isIdentifierPart(text[index]) {
			return false
		}
	}
	return true
}
