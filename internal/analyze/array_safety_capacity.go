package analyze

import (
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func arrayResumeNextCapacityGuards(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable) []arrayResumeNextCapacityGuard {
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), proc.EndLine)
	if start >= end {
		return nil
	}
	lineText := func(index int) string {
		return strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
	}
	nextNonEmpty := func(index int) (int, string, bool) {
		for index < end {
			text := lineText(index)
			if text != "" {
				return index, text, true
			}
			index++
		}
		return 0, "", false
	}

	var guards []arrayResumeNextCapacityGuard
	for index := start; index < end; index++ {
		if !arrayOnErrorResumeNextRe.MatchString(lineText(index)) {
			continue
		}
		probeIndex := -1
		var capacityName, targetName string
		for candidate := index + 1; candidate < end; candidate++ {
			text := lineText(candidate)
			if arrayOnErrorGotoZeroRe.MatchString(text) {
				break
			}
			if match := arrayCapacityProbeRe.FindStringSubmatch(text); len(match) == 3 {
				probeIndex = candidate
				capacityName = strings.ToLower(match[1])
				targetName = strings.ToLower(match[2])
				break
			}
			if match := arrayBoundsProbeRe.FindStringSubmatch(text); len(match) == 4 && strings.EqualFold(match[2], match[3]) {
				probeIndex = candidate
				capacityName = strings.ToLower(match[1])
				targetName = strings.ToLower(match[2])
				break
			}
		}
		if probeIndex < 0 {
			continue
		}
		variable, known := variables[targetName]
		if !known || !variable.isArray {
			continue
		}
		if arrayResumeNextCapacityStartsAtZero(file, proc, variables, index, capacityName) {
			if restoreIndex, restoreText, ok := nextNonEmpty(probeIndex + 1); ok &&
				(arrayOnErrorGotoZeroRe.MatchString(restoreText) || arrayOnErrorGotoRe.MatchString(restoreText)) {
				if checkIndex, checkText, ok := nextNonEmpty(restoreIndex + 1); ok {
					if match := arrayCheckedProbeExitRe.FindStringSubmatch(checkText); len(match) == 2 && strings.EqualFold(match[1], capacityName) {
						indexStartLine := checkIndex + 2
						indexEndLine := end
						for candidate := indexStartLine - 1; candidate < end; candidate++ {
							text := lineText(candidate)
							if erase := arrayEraseRe.FindStringSubmatch(text); len(erase) == 2 && strings.EqualFold(strings.TrimSpace(erase[1]), targetName) {
								indexEndLine = candidate
								break
							}
						}
						guards = append(guards, arrayResumeNextCapacityGuard{
							target:         targetName,
							probeLine:      probeIndex + 1,
							indexStartLine: indexStartLine,
							indexEndLine:   indexEndLine,
						})
						continue
					}
				}
			}
		}
		errIndex, errText, ok := nextNonEmpty(probeIndex + 1)
		if !ok || !arrayErrNumberFailureRe.MatchString(errText) {
			continue
		}
		errEnd := arraySourceIfEnd(file.Lines, errIndex, end)
		if errEnd < 0 {
			continue
		}
		capacityZero := false
		errCleared := false
		for candidate := errIndex + 1; candidate < errEnd; candidate++ {
			text := lineText(candidate)
			if strings.EqualFold(text, "err.clear") {
				errCleared = true
			}
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if assigned && !indexed && strings.EqualFold(lhs, capacityName) && strings.TrimSpace(rhs) == "0" {
				capacityZero = true
			}
		}
		if !capacityZero || !errCleared {
			continue
		}
		restoreIndex, restoreText, ok := nextNonEmpty(errEnd + 1)
		if !ok || !arrayOnErrorGotoZeroRe.MatchString(restoreText) {
			continue
		}
		capacityIfIndex, capacityIfText, ok := nextNonEmpty(restoreIndex + 1)
		if !ok {
			continue
		}
		capacityMatch := arrayCapacityIfRe.FindStringSubmatch(capacityIfText)
		if len(capacityMatch) != 2 || !strings.EqualFold(capacityMatch[1], capacityName) {
			continue
		}
		capacityIfEnd := arraySourceIfEnd(file.Lines, capacityIfIndex, end)
		if capacityIfEnd < 0 {
			continue
		}
		preserveTarget := false
		for candidate := capacityIfIndex + 1; candidate < capacityIfEnd; candidate++ {
			match := arrayRedimRe.FindStringSubmatch(lineText(candidate))
			if len(match) == 0 || strings.TrimSpace(match[1]) == "" {
				continue
			}
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct && strings.EqualFold(redim.name, targetName) {
					preserveTarget = true
					break
				}
			}
			if preserveTarget {
				break
			}
		}
		if !preserveTarget {
			continue
		}
		forIndex, forText, ok := nextNonEmpty(capacityIfEnd + 1)
		if !ok || !arrayForZeroToCountRe.MatchString(forText) {
			continue
		}
		targetIndexed := false
		nextIndex := -1
		for candidate := forIndex + 1; candidate < end; candidate++ {
			text := lineText(candidate)
			lower := strings.ToLower(text)
			if lower == "next" || strings.HasPrefix(lower, "next ") {
				nextIndex = candidate
				break
			}
			for _, use := range arrayIndexedUses(text, variables) {
				if strings.EqualFold(use.name, targetName) && len(use.args) > 0 {
					targetIndexed = true
					break
				}
			}
		}
		if !targetIndexed || nextIndex < 0 {
			continue
		}
		guards = append(guards, arrayResumeNextCapacityGuard{
			target:         targetName,
			probeLine:      probeIndex + 1,
			indexStartLine: forIndex + 2,
			indexEndLine:   nextIndex,
		})
	}
	return guards
}

// arrayResumeNextCapacityStartsAtZero proves the only state that makes a
// Resume Next bounds probe safe to use as a guard: a failed assignment must
// leave the capacity at zero. A stale positive value would make `If capacity
// <= 0 Then Exit ...` incorrectly accept an unallocated array.
func arrayResumeNextCapacityStartsAtZero(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, resumeIndex int, capacityName string) bool {
	variable, known := variables[strings.ToLower(cleanIdentifier(capacityName))]
	if known && variable.static {
		return false
	}
	start := max(0, proc.StartLine-1)
	lineText := func(index int) string {
		if index < 0 || index >= len(file.Lines) {
			return ""
		}
		return strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
	}
	for index := resumeIndex - 1; index >= start; index-- {
		text := lineText(index)
		if text == "" {
			continue
		}
		if rhs, assigned := arrayScalarAssignment(text, capacityName); assigned {
			return strings.TrimSpace(rhs) == "0"
		}
		break
	}

	if !known || variable.parameter || variable.isArray || variable.isVariant || !variable.knownScalar {
		return false
	}
	declarationLine := 0
	for declaration := range proc.Declarations.All() {
		if declaration.Scope != procedureir.ScopeLocal || !strings.EqualFold(cleanIdentifier(declaration.Name), cleanIdentifier(capacityName)) || declaration.IsArray || !arrayKnownScalarType(declaration.Type) {
			continue
		}
		declarationLine = declaration.Range.StartLine
		break
	}
	if declarationLine == 0 || declarationLine > resumeIndex+1 {
		return false
	}
	for index := start; index < resumeIndex; index++ {
		text := lineText(index)
		if text == "" || index+1 == declarationLine {
			continue
		}
		if _, assigned := arrayScalarAssignment(text, capacityName); assigned || strings.Contains(strings.ToLower(text), strings.ToLower(cleanIdentifier(capacityName))) {
			return false
		}
	}
	return true
}

func arrayScalarAssignment(text, name string) (string, bool) {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return "", false
	}
	statements := strings.Split(text, ":")
	for index := len(statements) - 1; index >= 0; index-- {
		statement := statements[index]
		if strings.TrimSpace(statement) == "" {
			continue
		}
		lhs, rhs, indexed, assigned := arrayAssignment(strings.TrimSpace(statement))
		if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
			return rhs, true
		}
		break
	}
	return "", false
}

func arraySourceIfEnd(lines []string, start, end int) int {
	depth := 0
	for index := start; index < end; index++ {
		text := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[index])))
		if text == "" {
			continue
		}
		if depth == 0 {
			depth = 1
			continue
		}
		if strings.HasPrefix(text, "if ") && strings.HasSuffix(text, " then") {
			depth++
			continue
		}
		if text == "end if" {
			depth--
			if depth == 0 {
				return index
			}
		}
	}
	return -1
}

func arrayResumeNextCapacityProbeApplies(guards []arrayResumeNextCapacityGuard, name string, line int) bool {
	for _, guard := range guards {
		if guard.probeLine == line && strings.EqualFold(guard.target, name) {
			return true
		}
	}
	return false
}

func arrayResumeNextCapacityIndexApplies(guards []arrayResumeNextCapacityGuard, name string, line int) bool {
	for _, guard := range guards {
		if line >= guard.indexStartLine && line <= guard.indexEndLine && strings.EqualFold(guard.target, name) {
			return true
		}
	}
	return false
}

func arrayResumeNextCapacityProofApplies(guards []arrayResumeNextCapacityGuard, name string, line int) bool {
	return arrayResumeNextCapacityProbeApplies(guards, name, line) || arrayResumeNextCapacityIndexApplies(guards, name, line)
}

type arrayModuleCapacityGuard struct {
	target       string
	required     string
	capacity     string
	redimLine    int
	capacityLine int
}

// applyArrayModuleCapacityGuardBranch recognizes the source-owned reusable
// buffer idiom:
//
//	If required > capacity Then
//	    ReDim buffer(0 To required - 1)
//	    capacity = required
//	End If
//
// On the false edge, the module scalar can only have its recorded capacity
// after the matching ReDim has completed. The source-owned and positive-input
// checks in arrayModuleCapacityGuardFor keep this refinement fail-closed.
func applyArrayModuleCapacityGuardBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	guard, ok := arrayModuleCapacityGuardFor(file, proc, ctx, statement, variables, moduleDecls)
	if !ok {
		return state
	}
	value, known := state[guard.target]
	variable, variableKnown := variables[guard.target]
	if !known || !variableKnown || !variable.isArray {
		return state
	}
	updated := cloneArrayState(state)
	value.kind = arrayAllocated
	value.knownArray = true
	value.mayBeEmpty = false
	updated[guard.target] = value
	return updated
}

func arrayModuleCapacityGuardFor(file parsedFile, proc sourceProcedure, ctx analysisContext, statement *procedureir.Statement, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) (arrayModuleCapacityGuard, bool) {
	if statement == nil || statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf {
		return arrayModuleCapacityGuard{}, false
	}
	required, capacity, ok := arrayModuleCapacityCondition(statement, variables, moduleDecls)
	if !ok || !arrayModuleCapacityRequiredPositive(file, proc, ctx, required, statement.Range.StartLine, variables) {
		return arrayModuleCapacityGuard{}, false
	}
	targets := make([]string, 0, 1)
	guard := arrayModuleCapacityGuard{required: required, capacity: capacity}
	for child := range proc.Statements.All() {
		if child.ParentID != statement.ID {
			continue
		}
		line := child.Range.StartLine
		text := strings.TrimSpace(child.Text)
		if text == "" && line >= 1 && line <= len(file.Lines) {
			text = strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		}
		switch child.Kind {
		case procedureir.StatementElse, procedureir.StatementElseIf:
			return arrayModuleCapacityGuard{}, false
		case procedureir.StatementReDim:
			if guard.redimLine != 0 {
				return arrayModuleCapacityGuard{}, false
			}
			match := arrayRedimRe.FindStringSubmatch(text)
			if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
				return arrayModuleCapacityGuard{}, false
			}
			clauses := splitArgs(match[2])
			if len(clauses) != 1 {
				return arrayModuleCapacityGuard{}, false
			}
			redim, direct := parseDirectArrayRedimClause(clauses[0])
			if !direct || !arrayCapacityRedimUsesRequired(redim.dimensions, required) {
				return arrayModuleCapacityGuard{}, false
			}
			name := strings.ToLower(cleanIdentifier(redim.name))
			if name == "" {
				return arrayModuleCapacityGuard{}, false
			}
			targets = append(targets, name)
			guard.target = name
			guard.redimLine = line
		case procedureir.StatementAssignment:
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), capacity) || !strings.EqualFold(strings.TrimSpace(rhs), required) {
				continue
			}
			if guard.capacityLine != 0 {
				return arrayModuleCapacityGuard{}, false
			}
			guard.capacityLine = line
		}
	}
	if len(targets) != 1 || guard.target == "" || guard.redimLine == 0 || guard.capacityLine == 0 || guard.redimLine >= guard.capacityLine {
		return arrayModuleCapacityGuard{}, false
	}
	targetDeclaration, targetDeclared := moduleDecls[guard.target]
	if !targetDeclared || !targetDeclaration.Array || targetDeclaration.Fixed || targetDeclaration.Parameter || !arrayModuleReadyGuardSourceOwned(file, targetDeclaration) {
		return arrayModuleCapacityGuard{}, false
	}
	capacityDeclaration, capacityDeclared := moduleDecls[guard.capacity]
	if !capacityDeclared || !arrayModuleReadyGuardSourceOwned(file, capacityDeclaration) {
		return arrayModuleCapacityGuard{}, false
	}
	targetVariable, targetKnown := variables[guard.target]
	if !targetKnown || !targetVariable.isArray || targetVariable.fixed {
		return arrayModuleCapacityGuard{}, false
	}
	if !arrayModuleCapacityLifecycleSafe(file, proc, guard, moduleDecls) {
		return arrayModuleCapacityGuard{}, false
	}
	return guard, true
}

func arrayModuleCapacityCondition(statement *procedureir.Statement, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) (required, capacity string, ok bool) {
	condition := statement.Text
	if statement.Condition != nil && strings.TrimSpace(statement.Condition.Text) != "" {
		condition = statement.Condition.Text
	}
	if parsed, _, parsedOK := arrayIfThenParts(condition); parsedOK {
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
	match := arrayCapacityComparisonRe.FindStringSubmatch(condition)
	if len(match) != 4 {
		return "", "", false
	}
	lhs := strings.ToLower(cleanIdentifier(match[1]))
	rhs := strings.ToLower(cleanIdentifier(match[3]))
	if match[2] == ">" {
		required, capacity = lhs, rhs
	} else {
		required, capacity = rhs, lhs
	}
	requiredVariable, requiredKnown := variables[required]
	capacityDeclaration, capacityDeclared := moduleDecls[capacity]
	if !requiredKnown || requiredVariable.parameter || requiredVariable.isArray || requiredVariable.isVariant || requiredVariable.isObject || !requiredVariable.knownScalar {
		return "", "", false
	}
	if _, isModule := moduleDecls[required]; isModule || !capacityDeclared || capacityDeclaration.Array || capacityDeclaration.Parameter || capacityDeclaration.Object || !arrayKnownScalarType(capacityDeclaration.Type) {
		return "", "", false
	}
	return required, capacity, true
}

func arrayCapacityRedimUsesRequired(dimensions, required string) bool {
	want := "0 to " + strings.ToLower(cleanIdentifier(required)) + " - 1"
	got := strings.ToLower(strings.Join(strings.Fields(dimensions), " "))
	return got == want
}

func arrayModuleCapacityRequiredPositive(file parsedFile, proc sourceProcedure, ctx analysisContext, required string, guardLine int, variables map[string]arrayVariable) bool {
	if guardLine <= proc.StartLine || guardLine > len(file.Lines) {
		return false
	}
	start := max(0, proc.StartLine-1)
	guardIndex := guardLine - 1
	assignmentLine := 0
	inputName := ""
	for index := start; index < guardIndex; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		lhs, rhs, indexed, assigned := arrayAssignment(text)
		if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), required) {
			continue
		}
		if assignmentLine != 0 {
			return false
		}
		match := arrayPositiveLengthExpressionRe.FindStringSubmatch(rhs)
		if len(match) != 3 {
			return false
		}
		factor, err := strconv.Atoi(match[2])
		if err != nil || factor <= 0 {
			return false
		}
		assignmentLine = index + 1
		inputName = strings.ToLower(cleanIdentifier(match[1]))
	}
	if assignmentLine == 0 || inputName == "" {
		return false
	}
	input, known := variables[inputName]
	if !known || input.isArray || input.isVariant || input.isObject || !input.knownScalar {
		return false
	}
	zeroGuardLine := 0
	for index := start; index < assignmentLine-1; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		condition, body, ok := arrayIfThenParts(text)
		if !ok {
			continue
		}
		guardInput, ok := arrayZeroLengthConditionInput(condition)
		if !ok || guardInput != inputName || !arrayModuleCapacityZeroLengthBody(file, proc, index+1, body, assignmentLine) {
			continue
		}
		statement := procedureStatementAtLine(proc, index+1)
		if statement.Kind != procedureir.StatementIf || statement.ParentID != 0 {
			continue
		}
		if zeroGuardLine != 0 {
			return false
		}
		zeroGuardLine = index + 1
	}
	if zeroGuardLine == 0 || proc.Graph == nil {
		return false
	}
	zeroStatement := procedureStatementAtLine(proc, zeroGuardLine)
	assignmentStatement := procedureStatementAtLine(proc, assignmentLine)
	if zeroStatement.Kind != procedureir.StatementIf || assignmentStatement.Kind != procedureir.StatementAssignment {
		return false
	}
	graph := arrayVBA227Graph(proc, ctx)
	zeroBlock, zeroOK := graph.BlockForStatement(zeroStatement.ID)
	assignmentBlock, assignmentOK := graph.BlockForStatement(assignmentStatement.ID)
	if !zeroOK || !assignmentOK {
		return false
	}
	for _, dominator := range graph.DominatorsOf(assignmentBlock.ID) {
		if dominator == zeroBlock.ID {
			return true
		}
	}
	return false
}

func arrayZeroLengthConditionInput(condition string) (string, bool) {
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	condition = strings.TrimSpace(condition)
	lower := strings.ToLower(condition)
	if strings.HasPrefix(lower, "if ") {
		condition = strings.TrimSpace(condition[len("if "):])
	}
	if then := arrayTopLevelKeywordIndex(condition, "then"); then >= 0 {
		condition = strings.TrimSpace(condition[:then])
	}
	match := arrayZeroLengthConditionRe.FindStringSubmatch(condition)
	if len(match) != 2 {
		return "", false
	}
	return strings.ToLower(cleanIdentifier(match[1])), true
}

func arrayModuleCapacityZeroLengthBody(file parsedFile, proc sourceProcedure, ifLine int, body string, assignmentLine int) bool {
	if strings.TrimSpace(body) != "" {
		thenBody, _, hasElse := arrayIfThenBodyParts(body)
		return !hasElse && strings.EqualFold(strings.TrimSpace(thenBody), "exit function")
	}
	end := arraySourceIfEnd(file.Lines, ifLine-1, min(len(file.Lines), proc.EndLine))
	if end < 0 || end+1 >= assignmentLine {
		return false
	}
	exits := 0
	for index := ifLine; index < end; index++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[index]))
		if text == "" {
			continue
		}
		if strings.EqualFold(text, "exit function") {
			exits++
			continue
		}
		lhs, _, indexed, assigned := arrayAssignment(text)
		if assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), proc.Name) {
			continue
		}
		return false
	}
	return exits == 1
}

func arrayModuleCapacityLifecycleSafe(file parsedFile, proc sourceProcedure, guard arrayModuleCapacityGuard, moduleDecls map[string]sourceDeclaration) bool {
	facts := file.moduleAnalysisFacts()
	if facts == nil {
		return false
	}
	sameProcedure := func(owner sourceProcedure) bool {
		return owner.StartByte == proc.StartByte && owner.StartLine == proc.StartLine && owner.EndLine == proc.EndLine
	}
	safe := true
	targetOperations := 0
	facts.forEachArrayOperationFor(guard.target, func(operation moduleArrayOperationFact) {
		if !safe {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
		if !ok {
			safe = false
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if scope.shadowsModule(guard.target) {
			return
		}
		targetOperations++
		if !sameProcedure(owner) || operation.Kind != moduleArrayDirectRedim || operation.Preserve || operation.Line+1 != guard.redimLine {
			safe = false
		}
	})
	if !safe || targetOperations != 1 {
		return false
	}

	capacityOperations := 0
	facts.forEachArrayOperationFor(guard.capacity, func(operation moduleArrayOperationFact) {
		if !safe {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
		if !ok {
			safe = false
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if scope.shadowsModule(guard.capacity) {
			return
		}
		capacityOperations++
		if !sameProcedure(owner) || operation.Kind != moduleArrayWholeAssignment || operation.Line+1 != guard.capacityLine || !strings.EqualFold(strings.TrimSpace(operation.RHS), guard.required) {
			safe = false
		}
	})
	return safe && capacityOperations == 1
}
