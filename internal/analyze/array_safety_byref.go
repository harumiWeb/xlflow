package analyze

import (
	"sort"
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func inferArrayByRefAllocationSummaries(files []parsedFile, ctx analysisContext, targets map[string]sourceProcedure) arrayByRefAllocationSummaries {
	summaries := arrayByRefAllocationSummaries{}
	procedures := make([]struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
	}, 0)
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !procedureHasByRefArrayParameter(proc) || !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			procedures = append(procedures, struct {
				file        parsedFile
				proc        sourceProcedure
				moduleDecls map[string]sourceDeclaration
			}{file: file, proc: proc, moduleDecls: moduleDecls})
		}
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		for call := range procedure.proc.Calls.All() {
			if targetKey, _, ok := arrayPrivateTargetForCall(ctx, targets, call); ok {
				dependents[targetKey] = append(dependents[targetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	contributions := make(arrayByRefAllocationSummaries, len(procedures))
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		procedure := procedures[index]
		key := arrayProcedureKey(procedure.proc)
		if !arrayProcedureIsParticipant(ctx, procedure.proc) {
			continue
		}
		if ctx.arrayStats != nil && procedure.proc.Graph != nil {
			ctx.arrayStats.addCFGWalk()
		}
		value := arrayByRefAllocationSummaryForProcedure(procedure.file, procedure.proc, summaries, ctx)
		old := arrayByRefAllocationSummaries{key: contributions[key]}
		fresh := arrayByRefAllocationSummaries{key: value}
		if arrayByRefAllocationSummariesEqual(old, fresh) {
			continue
		}
		if len(value) == 0 {
			delete(contributions, key)
			delete(summaries, key)
		} else {
			contributions[key] = value
			summaries[key] = value
		}
		for _, dependent := range dependents[key] {
			if !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
		}
	}
	return summaries
}

func arrayByRefAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, summaries arrayByRefAllocationSummaries, ctx analysisContext) map[int]bool {
	if proc.Graph == nil {
		return nil
	}
	parameters := map[string]int{}
	for index, parameter := range proc.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) {
			parameters[strings.ToLower(parameter.Name)] = index
		}
	}
	if len(parameters) == 0 {
		return nil
	}
	flowCtx := ctx
	flowCtx.arrayByRefAllocations = summaries
	moduleDecls := file.moduleDecls()
	return arrayByRefFlowAllocations(file, proc, flowCtx, moduleDecls)
}

// arrayByRefFlowAllocations proves ByRef array outputs at normal procedure
// exits. Unlike the direct-assignment summary above, this pass keeps the
// branch-local state established by an IsArray/type guard and meets it at the
// normal exit. It therefore recognizes helpers that fill the same output on
// multiple accepted input branches while excluding paths that terminate in a
// direct or project-local error raiser.
func arrayByRefFlowAllocations(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) map[int]bool {
	if proc.Graph == nil {
		return nil
	}
	variables := arrayVariables(file, proc, moduleDecls)
	initial := arrayInitialState(variables)
	graph := arrayVBA227Graph(proc, ctx)
	moduleArrays := arrayModuleNamesForProcedure(file, proc, moduleDecls)
	localGoSubAllocations := arrayLocalGoSubAllocationSummaries(proc, &graph, variables, ctx, arrayOptionBase(file), arrayIntegerConstants(file, proc, nil, nil), moduleArrays)
	parameterNames := make([]string, 0, proc.Params.Len())
	for _, parameter := range proc.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) {
			parameterNames = append(parameterNames, strings.ToLower(parameter.Name))
		}
	}
	var normalExit map[string]arrayValue
	hasNormalExit := false
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		out, _ := (Analyzer{}).arrayTransfer(file, proc, ctx, variables, in, text, line, nil, nil)
		out = applyArrayLocalGoSubStatementEffects(out, text, localGoSubAllocations)
		for _, call := range arrayCallsAtLine(proc.Calls, line) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
				out = applyArrayConditionalByRefCallEffects(out, proc, call, ctx)
			} else {
				out = applyArrayByRefCallEffects(out, proc, call, ctx)
			}
			out = applyArrayLocalGoSubEffects(out, proc, call, localGoSubAllocations)
		}
		return out
	}
	edgeState := func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
		out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
		if edge.To == graph.NormalExit() {
			if !hasNormalExit {
				normalExit = make(map[string]arrayValue, len(parameterNames))
				for _, name := range parameterNames {
					normalExit[name] = out[name]
				}
				hasNormalExit = true
			} else {
				for _, name := range parameterNames {
					normalExit[name] = meetArrayValue(normalExit[name], out[name])
				}
			}
		}
		return out
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, arrayAllocationTransferIsReliable, ctx.arrayStats)
	if !hasNormalExit {
		return nil
	}
	allocated := map[int]bool{}
	for index, parameter := range proc.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) {
			continue
		}
		value, ok := normalExit[strings.ToLower(parameter.Name)]
		if ok && value.kind == arrayAllocated && value.knownArray {
			allocated[index] = true
		}
	}
	return allocated
}

// inferArrayByRefConditionalAllocations recognizes the narrow, common output
// helper contract used by argument adapters:
//
//	If items.Count = 0 Then Exit Sub
//	ReDim values(0 To items.Count - 1)
//
// The guard must dominate the ReDim, and every normal path from the guard's
// non-exit branch must pass through that ReDim. This keeps the summary tied to
// the helper's actual control flow and avoids turning an arbitrary conditional
// ReDim into an allocation guarantee.
func inferArrayByRefConditionalAllocations(files []parsedFile) arrayByRefConditionalAllocations {
	summaries := arrayByRefConditionalAllocations{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if proc.Graph == nil {
				continue
			}
			parameters := map[string]int{}
			for index, parameter := range proc.Params.AllIndexed() {
				parameters[strings.ToLower(cleanIdentifier(parameter.Name))] = index
			}
			if len(parameters) == 0 {
				continue
			}
			guards := map[string]struct {
				statementID int
				line        int
				parameter   int
			}{}
			for statement := range proc.Statements.All() {
				match := arrayByRefCountExitRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
				if len(match) != 2 {
					continue
				}
				parameter, ok := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				if !ok {
					continue
				}
				guards[strings.ToLower(cleanIdentifier(match[1]))] = struct {
					statementID int
					line        int
					parameter   int
				}{statementID: statement.ID, line: statement.Range.StartLine, parameter: parameter}
			}
			if len(guards) == 0 {
				continue
			}
			for statement := range proc.Statements.All() {
				match := arrayByRefCountRedimRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
				if len(match) != 3 || arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
					continue
				}
				output, outputOK := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				guard, guardOK := guards[strings.ToLower(cleanIdentifier(match[2]))]
				if !outputOK || !guardOK || statement.Range.StartLine <= guard.line {
					continue
				}
				if output < 0 || output >= proc.Params.Len() || !parameterIsByRefArray(proc.Params.valueAt(output)) {
					continue
				}
				guardBlock, guardBlockOK := proc.Graph.BlockForStatement(guard.statementID)
				redimBlock, redimBlockOK := proc.Graph.BlockForStatement(statement.ID)
				if !guardBlockOK || !redimBlockOK {
					continue
				}
				guardDominatesRedim := false
				for _, candidate := range proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[redimBlock.ID] {
					if candidate == guardBlock.ID {
						guardDominatesRedim = true
						break
					}
				}
				if !guardDominatesRedim {
					continue
				}
				if !arrayFalseBranchRequiresBlock(*proc.Graph, guardBlock.ID, redimBlock.ID) {
					continue
				}
				key := arrayProcedureKey(proc)
				if summaries[key] == nil {
					summaries[key] = map[int]int{}
				}
				// Conflicting output contracts remain unknown rather than being
				// overwritten by declaration order.
				if previous, exists := summaries[key][output]; exists && previous != guard.parameter {
					delete(summaries[key], output)
					continue
				}
				summaries[key][output] = guard.parameter
			}
			if len(summaries[arrayProcedureKey(proc)]) == 0 {
				delete(summaries, arrayProcedureKey(proc))
			}
		}
	}
	return summaries
}

// inferArrayByRefLengthAllocations recognizes a helper that returns the
// successful length of a ByRef array through a paired ByRef scalar output:
//
// \tbyteLength = UBound(bytes) - LBound(bytes) + 1
//
// The assignment must dominate the procedure's normal exit, or be the normal
// branch of an explicit zero-length guard whose sibling assigns zero. A
// positive length at a caller then proves that the array-bound query completed
// successfully, without making the array unconditionally allocated on other
// paths.
func inferArrayByRefLengthAllocations(files []parsedFile) arrayByRefLengthAllocations {
	summaries := arrayByRefLengthAllocations{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if proc.Graph == nil {
				continue
			}
			parameters := map[string]int{}
			for index, parameter := range proc.Params.AllIndexed() {
				parameters[strings.ToLower(cleanIdentifier(parameter.Name))] = index
			}
			if len(parameters) == 0 {
				continue
			}
			dominators := arrayProcedureNormalExitDominators(proc)
			for statement := range proc.Statements.All() {
				text := strings.TrimSpace(normalizedCodeLine(statement.Text))
				match := arrayByRefLengthFullRe.FindStringSubmatch(text)
				if len(match) == 4 && !strings.EqualFold(cleanIdentifier(match[2]), cleanIdentifier(match[3])) {
					match = nil
				}
				if len(match) == 0 {
					upper := arrayByRefLengthUpperRe.FindStringSubmatch(text)
					if len(upper) == 3 {
						match = []string{upper[0], upper[1], upper[2], upper[2]}
					}
				}
				if len(match) != 4 {
					continue
				}
				lengthIndex, lengthOK := parameters[strings.ToLower(cleanIdentifier(match[1]))]
				arrayIndex, arrayOK := parameters[strings.ToLower(cleanIdentifier(match[2]))]
				if !lengthOK || !arrayOK || lengthIndex == arrayIndex || !parameterIsByRefScalar(proc.Params.valueAt(lengthIndex)) || !parameterIsByRefArray(proc.Params.valueAt(arrayIndex)) {
					continue
				}
				dominatesExit := arrayProcedureBlockDominatesNormalExit(proc, statement.ID, dominators)
				if !dominatesExit {
					parent, parentOK := arrayByRefLengthGuard(proc, statement.ID)
					if !parentOK || !arrayProcedureBlockDominatesNormalExit(proc, parent.ID, dominators) || !arrayByRefLengthHasZeroBranch(proc, parent.ID, statement.ID, match[1]) {
						continue
					}
				}
				key := arrayProcedureKey(proc)
				if summaries[key] == nil {
					summaries[key] = map[int]int{}
				}
				if previous, exists := summaries[key][arrayIndex]; exists && previous != lengthIndex {
					delete(summaries[key], arrayIndex)
					continue
				}
				summaries[key][arrayIndex] = lengthIndex
			}
			if len(summaries[arrayProcedureKey(proc)]) == 0 {
				delete(summaries, arrayProcedureKey(proc))
			}
		}
	}
	return summaries
}

func arrayProcedureStatementByID(proc sourceProcedure, id int) (procedureir.Statement, bool) {
	for statement := range proc.Statements.All() {
		if statement.ID == id {
			return statement, true
		}
	}
	return procedureir.Statement{}, false
}

func arrayByRefLengthGuard(proc sourceProcedure, statementID int) (procedureir.Statement, bool) {
	seen := map[int]bool{}
	for statementID != 0 && !seen[statementID] {
		seen[statementID] = true
		statement, ok := arrayProcedureStatementByID(proc, statementID)
		if !ok {
			return procedureir.Statement{}, false
		}
		if statement.Kind == procedureir.StatementIf {
			return statement, true
		}
		statementID = statement.ParentID
	}
	return procedureir.Statement{}, false
}

func arrayByRefLengthHasZeroBranch(proc sourceProcedure, parentID, formulaID int, lengthName string) bool {
	for statement := range proc.Statements.All() {
		if statement.ID == formulaID || statement.ParentID != parentID {
			continue
		}
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(statement.Text))
		if ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(lengthName)) && strings.TrimSpace(rhs) == "0" {
			return true
		}
	}
	return false
}

func arrayFalseBranchRequiresBlock(graph vbacfg.Graph, guardBlock, requiredBlock vbacfg.BlockID) bool {
	queue := make([]vbacfg.BlockID, 0, 1)
	for _, edge := range graph.Edges {
		if edge.From == guardBlock && edge.Kind == vbacfg.EdgeBranchFalse && edge.Class == vbacfg.EdgeNormal {
			queue = append(queue, edge.To)
		}
	}
	if len(queue) == 0 {
		return false
	}
	visited := map[vbacfg.BlockID]bool{}
	reachedRequired := false
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if visited[current] {
			continue
		}
		visited[current] = true
		if current == requiredBlock {
			reachedRequired = true
			continue
		}
		if current == graph.NormalExit {
			return false
		}
		for _, edge := range graph.Edges {
			if edge.From != current || edge.Class != vbacfg.EdgeNormal {
				continue
			}
			queue = append(queue, edge.To)
		}
	}
	return reachedRequired
}

func arrayByRefAllocationSummariesEqual(left, right arrayByRefAllocationSummaries) bool {
	if len(left) != len(right) {
		return false
	}
	for procedure, parameters := range left {
		other, ok := right[procedure]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index := range parameters {
			if !other[index] {
				return false
			}
		}
	}
	return true
}

// applyArrayByRefCallEffects carries the invalidating side of a private
// ByRef-array contract. The ordinary CFG walk records proven allocation
// outputs, but the recovered source-order fallback also has to account for a
// preceding helper that can Erase or otherwise replace the caller's array.
// Read-only helpers preserve the caller's allocation state; helpers whose
// parameter effect cannot be proven make it unknown.
func applyArrayByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	key, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return applyArrayUnknownByRefCallEffects(state, proc, call, ctx)
	}
	if !procedureHasByRefArrayParameter(target) {
		return state
	}
	bindings, ok := arrayCallArgumentBindings(proc, target, call)
	if !ok {
		return state
	}
	// The compact CFG solver may revisit sibling branch blocks with a state map
	// that shares storage with their predecessor. Keep this call's mutation
	// isolated so an invalidating true branch cannot poison its false sibling.
	updated := cloneArrayState(state)
	apply := func(index int, argument string) {
		if index < 0 || index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
			return
		}
		name := directArrayArgumentName(argument)
		if name == "" {
			return
		}
		if ctx.arrayByRefAllocations[key][index] {
			value := updated[name]
			value.kind = arrayAllocated
			value.knownArray = true
			updated[name] = value
			return
		}
		// Conditional and paired-length output contracts carry a stronger
		// caller-side fact than the generic invalidation scan. Their output may
		// be unallocated for the zero-length case, but the existing count/length
		// refinement will establish the successful branch without losing that
		// relation to an unknown state here.
		if _, conditional := ctx.arrayByRefConditionalAllocations[key][index]; conditional {
			return
		}
		if _, pairedLength := ctx.arrayByRefLengthAllocations[key][index]; pairedLength {
			return
		}
		if arrayByRefParameterMayInvalidate(target, index, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	for _, binding := range bindings {
		apply(binding.parameterIndex, binding.text)
	}
	return updated
}

func applyArrayUnknownByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	if _, _, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call); resolved {
		return state
	}
	updated := cloneArrayState(state)
	markUnknown := func(argument string) {
		name := directArrayArgumentName(argument)
		if name == "" {
			return
		}
		if _, tracked := updated[name]; tracked {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	for _, argument := range call.Arguments.Named {
		markUnknown(argument.ValueText)
	}
	for _, argument := range arrayCallArgumentTexts(proc, call) {
		markUnknown(argument)
	}
	return updated
}

// applyArrayConditionalByRefCallEffects joins the two possible outcomes of
// an inline conditional call without treating the call's allocation summary
// as unconditional. A helper that may invalidate its ByRef array makes the
// joined state unknown; a helper that only allocates leaves the prior state,
// which is the conservative meet of the conditional paths.
func applyArrayConditionalByRefCallEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext) arrayFlowState {
	_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return applyArrayUnknownByRefCallEffects(state, proc, call, ctx)
	}
	if !procedureHasByRefArrayParameter(target) {
		return state
	}
	bindings, ok := arrayCallArgumentBindings(proc, target, call)
	if !ok {
		return state
	}
	updated := cloneArrayState(state)
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		name := directArrayArgumentName(binding.text)
		if name == "" {
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

// arrayByRefParameterMayInvalidate reports whether a private ByRef-array
// parameter can lose its allocation while the procedure returns normally.
// Element writes and reads preserve an allocated input; Erase, unknown whole
// array replacement, and an unproven nested ByRef call do not. The recursion
// guard keeps recursive helper cycles from being treated as an additional
// mutation. Any direct mutation in the current cycle member is still found by
// its own scan.
func arrayByRefParameterMayInvalidate(proc sourceProcedure, parameterIndex int, ctx analysisContext, visiting map[string]bool) bool {
	if parameterIndex < 0 || parameterIndex >= proc.Params.Len() || !parameterIsByRefArray(proc.Params.valueAt(parameterIndex)) {
		return true
	}
	key := strings.ToLower(arrayProcedureKey(proc)) + "#" + strconv.Itoa(parameterIndex)
	if visiting[key] {
		return false
	}
	visiting[key] = true
	defer delete(visiting, key)
	name := strings.ToLower(cleanIdentifier(proc.Params.valueAt(parameterIndex).Name))
	names := map[string]bool{name: true}
	for statement := range proc.Statements.All() {
		if !arrayByRefStatementReachable(proc, statement) {
			continue
		}
		if statement.Recovered {
			return true
		}
		for _, part := range splitRangeValueSourceStatements(statement.Text) {
			part = strings.TrimSpace(part)
			if part == "" {
				continue
			}
			if arrayByRefParameterInlineMutation(part, names, ctx) {
				return true
			}
		}
		for _, nested := range arrayCallsAtLine(proc.Calls, statement.Range.StartLine) {
			if arrayByRefCallIsReadOnly(nested) {
				continue
			}
			if !arrayCallPassesDirectArrayArgument(proc, nested, name) {
				continue
			}
			nestedKey, nestedTarget, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, nested)
			if !resolved {
				return true
			}
			bindings, mapped := arrayCallArgumentBindings(proc, nestedTarget, nested)
			if !mapped {
				return true
			}
			for _, binding := range bindings {
				if directArrayArgumentName(binding.text) != name || binding.parameterIndex >= nestedTarget.Params.Len() || !parameterIsByRefArray(nestedTarget.Params.valueAt(binding.parameterIndex)) {
					continue
				}
				if ctx.arrayByRefAllocations[nestedKey][binding.parameterIndex] {
					continue
				}
				if arrayByRefParameterMayInvalidate(nestedTarget, binding.parameterIndex, ctx, visiting) {
					return true
				}
			}
		}
	}
	return false
}

func arrayByRefCallIsReadOnly(call procedureir.CallSite) bool {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return true
	}
	switch strings.ToLower(cleanIdentifier(call.Callee.BaseName)) {
	case "lbound", "ubound":
		return true
	default:
		return false
	}
}

func arrayByRefStatementReachable(proc sourceProcedure, statement procedureir.Statement) bool {
	if proc.Graph == nil {
		return true
	}
	block, ok := proc.Graph.BlockForStatement(statement.ID)
	if !ok {
		return true
	}
	return proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).IsReachable(block.ID)
}

func arrayByRefParameterInlineMutation(text string, names map[string]bool, ctx analysisContext) bool {
	if condition, body, ok := arrayIfThenParts(text); ok && strings.TrimSpace(body) != "" {
		condition = strings.TrimSpace(condition)
		lowerCondition := strings.ToLower(condition)
		switch {
		case strings.HasPrefix(lowerCondition, "if "):
			condition = strings.TrimSpace(condition[len("if "):])
		case strings.HasPrefix(lowerCondition, "elseif "):
			condition = strings.TrimSpace(condition[len("elseif "):])
		}
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		branches := []string{thenBody}
		if hasElse {
			branches = append(branches, elseBody)
		}
		if value, known := arraySourceOrderConstantBoolean(condition, nil); known {
			if value {
				branches = []string{thenBody}
			} else if hasElse {
				branches = []string{elseBody}
			} else {
				branches = nil
			}
		}
		for _, branch := range branches {
			for _, statement := range splitRangeValueSourceStatements(branch) {
				if arrayByRefParameterInlineMutation(strings.TrimSpace(statement), names, ctx) {
					return true
				}
			}
		}
		return false
	}
	text = strings.TrimSpace(text)
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
		for _, target := range splitArgs(match[1]) {
			if names[strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))] {
				return true
			}
		}
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			name := strings.ToLower(cleanIdentifier(redim.name))
			if !direct || !names[name] {
				continue
			}
			// ReDim Preserve retains an already allocated input array. Its shape
			// may change, but that is not an allocation invalidation at the
			// caller boundary.
			if strings.TrimSpace(match[1]) != "" {
				continue
			}
			if !arrayStatementAllocatesName(text, name, ctx) {
				return true
			}
		}
	}
	if lhs, _, indexed, ok := arrayAssignment(text); ok && !indexed {
		name := strings.ToLower(cleanIdentifier(lhs))
		return names[name] && !arrayStatementAllocatesName(text, name, ctx)
	}
	return false
}

func arrayLocalGoSubAllocationSummaries(proc sourceProcedure, graph *vbacfg.CFGView, variables map[string]arrayVariable, ctx analysisContext, base int, constants map[string]int, moduleArrays map[string]bool) arrayLocalGoSubAllocations {
	statements := make([]procedureir.Statement, 0, proc.Statements.Len())
	for statement := range proc.Statements.All() {
		statements = append(statements, statement)
	}
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartLine != statements[j].Range.StartLine {
			return statements[i].Range.StartLine < statements[j].Range.StartLine
		}
		return statements[i].ID < statements[j].ID
	})
	summaries := arrayLocalGoSubAllocations{}
	if graph == nil {
		return summaries
	}
	for index, label := range statements {
		if label.Kind != procedureir.StatementLabel {
			continue
		}
		labelName := arrayLocalGoSubLabelName(label)
		if labelName == "" {
			continue
		}
		end := len(statements)
		for cursor := index + 1; cursor < len(statements); cursor++ {
			if arrayLocalGoSubIsReturn(statements[cursor]) {
				end = cursor
				break
			}
		}
		if end == len(statements) {
			continue
		}
		summary := arrayLocalGoSubSummary{
			guaranteedAllocated: map[string]bool{},
			unknown:             map[string]bool{},
		}
		for name, variable := range variables {
			if variable.isArray && arrayLocalGoSubAllocationInvariant(proc, graph, statements, index, end, name, ctx, base, constants, moduleArrays) {
				summary.guaranteedAllocated[name] = true
			} else if variable.isArray && !variable.fixed && arrayLocalGoSubMayMutateName(proc, graph, statements, index, end, name, ctx, moduleArrays) {
				summary.unknown[name] = true
			}
		}
		summaries[labelName] = summary
	}
	return summaries
}

func arrayLocalGoSubAllocationInvariant(proc sourceProcedure, graph *vbacfg.CFGView, statements []procedureir.Statement, labelIndex, end int, name string, ctx analysisContext, base int, constants map[string]int, moduleArrays map[string]bool) bool {
	if graph == nil || labelIndex < 0 || labelIndex >= end || end > len(statements) {
		return false
	}
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	allowed := map[vbacfg.BlockID]procedureir.Statement{}
	var labelBlock vbacfg.Block
	labelFound := false
	for index := labelIndex; index < end; index++ {
		block, ok := graph.BlockForStatement(statements[index].ID)
		if !ok {
			return false
		}
		allowed[block.ID] = statements[index]
		if index == labelIndex {
			labelBlock = block
			labelFound = true
		}
	}
	if !labelFound {
		return false
	}
	returnBlock, returnOK := graph.BlockForStatement(statements[end].ID)
	if !returnOK {
		return false
	}

	type stateAtBlock struct {
		id        vbacfg.BlockID
		allocated bool
	}
	seenStates := map[vbacfg.BlockID]map[bool]bool{labelBlock.ID: {false: true}}
	queue := []stateAtBlock{{id: labelBlock.ID}}
	for len(queue) > 0 {
		currentState := queue[0]
		queue = queue[1:]
		currentID := currentState.id
		if _, ok := graph.BlockByID(currentID); !ok {
			return false
		}
		failed := false
		graph.ForEachOutgoing(currentID, func(edge vbacfg.Edge) bool {
			target, targetOK := graph.BlockByID(edge.To)
			if !targetOK {
				failed = true
				return true
			}
			// An unknown edge can leave the GoSub body through a dynamic or
			// recovered transfer. It is not evidence of a successful Return,
			// even when the predecessor happens to be allocated.
			if edge.Kind == vbacfg.EdgeUnknown || target.Kind == vbacfg.BlockUnknownExit {
				failed = true
				return true
			}
			statement, inside := allowed[target.ID]
			if !inside {
				// The first Return is deliberately outside the allowed body. Any
				// other statement target leaves the GoSub body without returning to
				// its caller, so it cannot be treated as a successful summary.
				if target.ID != returnBlock.ID && !arrayLocalGoSubIsTerminalBlock(target) {
					failed = true
					return true
				}
				if !currentState.allocated {
					failed = true
				}
				// Keep visiting sibling edges. A terminal edge is safe only when
				// every other edge from this block also preserves the summary.
				return true
			}
			nextAllocated := arrayLocalGoSubStateAfterStatement(statement.Text, name, currentState.allocated, ctx, base, constants)
			unknownCall, guaranteedCall := arrayLocalGoSubArrayCallEffect(proc, statement, name, ctx, moduleArrays)
			if unknownCall {
				nextAllocated = false
			} else if guaranteedCall {
				nextAllocated = true
			}
			if !seenStates[target.ID][nextAllocated] {
				if seenStates[target.ID] == nil {
					seenStates[target.ID] = map[bool]bool{}
				}
				seenStates[target.ID][nextAllocated] = true
				queue = append(queue, stateAtBlock{id: target.ID, allocated: nextAllocated})
			}
			return true
		})
		if failed {
			return false
		}
	}
	return true
}

func arrayLocalGoSubStateAfterStatement(text, name string, allocated bool, ctx analysisContext, base int, constants map[string]int) bool {
	names := map[string]bool{name: true}
	for _, statement := range splitRangeValueSourceStatements(text) {
		statement = strings.TrimSpace(statement)
		if statement == "" {
			continue
		}
		if arrayLocalGoSubAllocationIsReliable(statement, name, ctx, base, constants) {
			allocated = true
			continue
		}
		if arrayLocalGoSubPreserveKeepsAllocation(statement, name, allocated, base, constants) {
			continue
		}
		if arraySourceOrderInlineArrayMutation(statement, names, ctx) || arraySourceOrderMutatesArrayStatement(statement, names, ctx) {
			allocated = false
		}
	}
	return allocated
}

func arrayLocalGoSubPreserveKeepsAllocation(text, name string, allocated bool, base int, constants map[string]int) bool {
	if !allocated {
		return false
	}
	match := arrayRedimRe.FindStringSubmatch(text)
	if len(match) == 0 || strings.TrimSpace(match[1]) == "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		if direct && strings.EqualFold(cleanIdentifier(redim.name), name) && !arraySummaryStatementAlwaysFails(text, base, constants) {
			return true
		}
	}
	return false
}

func arrayLocalGoSubMayMutateName(proc sourceProcedure, graph *vbacfg.CFGView, statements []procedureir.Statement, labelIndex, end int, name string, ctx analysisContext, moduleArrays map[string]bool) bool {
	if graph == nil || labelIndex < 0 || labelIndex >= end || end > len(statements) {
		return true
	}
	for index := labelIndex; index < end; index++ {
		statement := statements[index]
		for _, part := range splitRangeValueSourceStatements(statement.Text) {
			part = strings.TrimSpace(part)
			if part == "" {
				continue
			}
			if arrayStatementAllocatesName(part, name, ctx) || arraySourceOrderInlineArrayMutation(part, map[string]bool{name: true}, ctx) || arraySourceOrderMutatesArrayStatement(part, map[string]bool{name: true}, ctx) {
				return true
			}
		}
		unknownCall, _ := arrayLocalGoSubArrayCallEffect(proc, statement, name, ctx, moduleArrays)
		if unknownCall {
			return true
		}
	}
	return false
}

// arrayLocalGoSubArrayCallEffect keeps a local GoSub summary fail-closed at
// calls that pass the tracked array directly. Only an already-proven private
// ByRef allocation contract is allowed to establish the post-call state;
// builtin-like calls are treated as non-mutating, while every other call may
// erase or otherwise invalidate the array behind the analyzer's view.
func arrayLocalGoSubArrayCallEffect(proc sourceProcedure, statement procedureir.Statement, name string, ctx analysisContext, moduleArrays map[string]bool) (unknown, guaranteedAllocated bool) {
	name = strings.ToLower(cleanIdentifier(name))
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
			continue
		}
		if call.StatementID != statement.ID && (call.StatementID != 0 || call.Range.StartLine != statement.Range.StartLine) {
			continue
		}
		arguments := arrayCallArgumentTexts(proc, call)
		relevantArgument := false
		key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if resolved {
			bindings, mapped := arrayCallArgumentBindings(proc, target, call)
			if mapped {
				for _, binding := range bindings {
					if directArrayArgumentName(binding.text) != name {
						continue
					}
					if binding.parameterIndex >= 0 && binding.parameterIndex < target.Params.Len() && target.Params.valueAt(binding.parameterIndex).ParamArray {
						continue
					}
					relevantArgument = true
					if binding.parameterIndex < target.Params.Len() && parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) && ctx.arrayByRefAllocations[key][binding.parameterIndex] {
						guaranteedAllocated = true
						continue
					}
					unknown = true
				}
			} else if call.Arguments.Count > 0 {
				// A resolved call with an incomplete argument projection cannot
				// prove which formal parameter receives the tracked array.
				unknown = true
			}
		} else {
			// Keep unresolved calls conservative by checking the raw actual
			// expressions. Named arguments are handled here only because there
			// is no resolved formal signature to bind them to.
			for _, argument := range call.Arguments.Named {
				if directArrayArgumentName(argument.ValueText) == name {
					relevantArgument = true
					unknown = true
				}
			}
			for _, argument := range arguments {
				if directArrayArgumentName(argument) == name {
					relevantArgument = true
					unknown = true
				}
			}
		}
		if call.Arguments.Count > 0 && len(arguments) != call.Arguments.Count {
			// An incomplete argument projection cannot prove that this call did
			// not receive the tracked array by reference.
			unknown = true
		}
		if !relevantArgument && moduleArrays[name] {
			if arrayLocalGoSubCallHasProvenModuleContract(ctx, call, name) {
				guaranteedAllocated = true
			} else {
				unknown = true
			}
		}
	}
	return unknown, guaranteedAllocated
}

func arrayLocalGoSubCallHasProvenModuleContract(ctx analysisContext, call procedureir.CallSite, name string) bool {
	key, _, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return false
	}
	return ctx.arrayModuleAllocations[key][strings.ToLower(cleanIdentifier(name))]
}

func arrayLocalGoSubIsTerminalBlock(block vbacfg.Block) bool {
	switch block.Kind {
	case vbacfg.BlockNormalExit, vbacfg.BlockExceptionalExit, vbacfg.BlockTerminationExit, vbacfg.BlockUnknownExit:
		return true
	default:
		return false
	}
}

func arrayLocalGoSubAllocationIsReliable(text, name string, ctx analysisContext, base int, constants map[string]int) bool {
	if !arrayStatementAllocatesName(text, name, ctx) {
		return false
	}
	return !arraySummaryStatementAlwaysFails(text, base, constants)
}

func arrayLocalGoSubLabelName(statement procedureir.Statement) string {
	label := statement.Label
	if label == "" {
		label = statement.Text
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(label, ":")))
}

func arrayLocalGoSubIsReturn(statement procedureir.Statement) bool {
	return strings.EqualFold(strings.TrimSpace(statement.Text), "return")
}

func arrayLocalGoSubTarget(proc sourceProcedure, call procedureir.CallSite) string {
	statement, ok := arrayProcedureStatementByID(proc, call.StatementID)
	if !ok || !strings.EqualFold(call.Callee.BaseName, "gosub") {
		return ""
	}
	text := strings.TrimSpace(statement.Text)
	if len(text) < len("gosub") || !strings.EqualFold(text[:len("gosub")], "gosub") {
		return ""
	}
	fields := strings.Fields(strings.TrimSpace(text[len("gosub"):]))
	if len(fields) == 0 {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(fields[0], ":")))
}

func applyArrayLocalGoSubEffects(state arrayFlowState, proc sourceProcedure, call procedureir.CallSite, summaries arrayLocalGoSubAllocations) arrayFlowState {
	target := arrayLocalGoSubTarget(proc, call)
	if target == "" {
		return state
	}
	return applyArrayLocalGoSubSummary(state, target, summaries)
}

func applyArrayLocalGoSubSummary(state arrayFlowState, target string, summaries arrayLocalGoSubAllocations) arrayFlowState {
	summary, known := summaries[target]
	if !known {
		// A local GoSub can mutate caller/module arrays without carrying an
		// explicit array argument. An absent summary therefore means unknown
		// state, not "no effect"; retaining allocated here would suppress
		// VBA227 after cleanup or another unmodelled side effect.
		updated := cloneArrayState(state)
		for name := range updated {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
		return updated
	}
	updated := cloneArrayState(state)
	for name := range summary.unknown {
		updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
	}
	for name := range summary.guaranteedAllocated {
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func applyArrayLocalGoSubStatementEffects(state arrayFlowState, text string, summaries arrayLocalGoSubAllocations) arrayFlowState {
	for _, statement := range splitRangeValueSourceStatements(text) {
		target, ok := arrayLocalGoSubTargetFromStatementText(statement)
		if !ok {
			continue
		}
		state = applyArrayLocalGoSubSummary(state, target, summaries)
	}
	return state
}

func arrayLocalGoSubTargetFromStatementText(text string) (string, bool) {
	fields := strings.Fields(strings.TrimSpace(text))
	if len(fields) < 2 || !strings.EqualFold(fields[0], "gosub") {
		return "", false
	}
	return strings.ToLower(strings.TrimSpace(strings.TrimSuffix(fields[1], ":"))), true
}

func inferArrayByRefEntryStates(a Analyzer, files []parsedFile, ctx analysisContext) (map[string]map[int]bool, map[string]map[int]string) {
	targets := ctx.arrayPrivateTargets
	if len(targets) == 0 {
		return map[string]map[int]bool{}, map[string]map[int]string{}
	}
	moduleAllocationSummaries := ctx.arrayModuleAllocations
	moduleInitializationStates := arrayModuleInitializationStates(files, moduleAllocationSummaries)

	type callerInfo struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
		variables   map[string]arrayVariable
		constants   map[string]int
	}
	callers := make([]callerInfo, 0)
	for _, file := range files {
		moduleDecls := file.moduleDecls()
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			caller := procedures.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, caller) {
				continue
			}
			eligibleCaller := false
			for call := range caller.Calls.All() {
				_, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
				if ok && procedureHasByRefArrayParameter(target) && arrayProcedureIsParticipant(ctx, target) {
					eligibleCaller = true
					break
				}
			}
			if !eligibleCaller {
				continue
			}
			callers = append(callers, callerInfo{
				file: file, proc: caller, moduleDecls: moduleDecls,
				variables: arrayVariables(file, caller, moduleDecls),
				constants: arrayIntegerConstants(file, caller, a.visibleConstantValues, a.visibleConstants),
			})
		}
	}
	sort.SliceStable(callers, func(i, j int) bool {
		return arrayProcedureLess(callers[i].proc, callers[j].proc)
	})

	evaluateCaller := func(caller callerInfo, entries map[string]map[int]bool, conditions map[string]map[int]string) map[string]map[int]arrayByRefEntryEvidence {
		file := caller.file
		proc := caller.proc
		moduleDecls := caller.moduleDecls
		variables := caller.variables
		constants := caller.constants
		localCtx := ctx
		localCtx.arrayByRefEntryConditions = conditions
		evidence := map[string]map[int]arrayByRefEntryEvidence{}
		initial := arrayInitialState(variables)
		initial = applyArrayModuleInitializationState(initial, file, proc, variables, moduleDecls, moduleInitializationStates)
		initial = applyArrayModuleReadyGuardState(initial, file, proc, variables, moduleDecls, ctx.arrayModuleReadyGuards)
		initial = applyArrayByRefEntryStates(initial, proc, variables, entries, conditions)
		initial = applyArrayModuleEntryState(initial, file, proc, variables, moduleDecls, ctx.arrayModuleEntryStates, ctx.arrayParticipantKeys)
		initial = applyArrayInternalStorageConfiguration(initial, file, proc, variables, moduleDecls, ctx.arrayModuleConfigurations[file.Path])
		var baseView vbacfg.CFGView
		var summaryGraph *vbacfg.CFGView
		var worklistReachable map[vbacfg.BlockID]bool
		worklistReachableLines := map[int]bool{}
		if proc.Graph != nil {
			baseView = proc.Graph.View(vbacfg.EdgeFilter{})
			summaryGraph = &baseView
			worklistReachable = arrayCFGWorklistReachable(&baseView)
			for statement := range proc.Statements.All() {
				line := statement.Range.StartLine
				owner, ownerOK := baseView.BlockForStatement(statement.ID)
				if line > 0 && ownerOK && worklistReachable[owner.ID] {
					worklistReachableLines[line] = true
				}
			}
		}
		moduleArrays := arrayModuleNamesForProcedure(file, proc, moduleDecls)
		localGoSubAllocations := arrayLocalGoSubAllocationSummaries(proc, summaryGraph, variables, localCtx, arrayOptionBase(file), constants, moduleArrays)
		visitForBlock := func(text string, line int, in arrayFlowState, ownerStatementID int, filterNestedCalls, skipNestedState bool) arrayFlowState {
			if skipNestedState {
				return in
			}
			var eligible []arrayByRefCallCandidate
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				// A physical line can own more than one CFG statement, such as
				// `If Not SendFrameFor(...) Then Exit Function`.  Only the
				// statement block that owns the call site may contribute entry
				// evidence; otherwise the nested Exit block re-records the same
				// call with its pre-call state and poisons the caller contract.
				if ownerStatementID > 0 && call.StatementID != ownerStatementID {
					if !filterNestedCalls {
						continue
					}
					owner, ownerOK := baseView.BlockForStatement(call.StatementID)
					if ownerOK && worklistReachable[owner.ID] {
						continue
					}
				}
				key, target, ok := arrayPrivateTargetForCall(localCtx, targets, call)
				if !ok || !procedureHasByRefArrayParameter(target) || !arrayProcedureIsParticipant(localCtx, target) {
					continue
				}
				eligible = append(eligible, arrayByRefCallCandidate{key: key, target: target, call: call})
			}
			if len(eligible) > 0 {
				allSameTarget := true
				for _, entry := range eligible[1:] {
					if entry.key != eligible[0].key {
						allSameTarget = false
						break
					}
				}
				// Record and apply each call in source order. This matters when
				// one physical line invokes the same ByRef helper repeatedly: the
				// second call must see an Erase or other invalidation from the
				// first call rather than the original pre-line state.
				recordState := cloneArrayState(in)
				for _, entry := range eligible {
					record := allSameTarget
					if !record {
						// Nested calls on one source line are normally kept
						// conservative because the pre-line state cannot describe
						// mutations from an earlier, different helper. An outer
						// ByRef call whose array argument is a proven allocated
						// expression is independent of that ordering, however.
						allProven, hasExpression := arrayByRefCallHasProvenArrayArguments(file, entry.target, proc, entry.call, recordState, localCtx)
						record = allProven && (hasExpression || arrayByRefCallIsInnermostNested(entry.call, eligible))
					}
					if record {
						arrayRecordByRefCall(evidence, entry.key, entry.target, proc, entry.call, file, recordState, localCtx)
					}
					recordState = applyArrayModuleCallEffects(recordState, file, proc, entry.call, localCtx, variables, moduleDecls)
					if arrayProcedureLineHasInlineConditional(file, entry.call.Range.StartLine) {
						recordState = applyArrayConditionalByRefCallEffects(recordState, proc, entry.call, localCtx)
					} else {
						recordState = applyArrayByRefCallEffects(recordState, proc, entry.call, localCtx)
					}
					recordState = applyArrayLocalGoSubEffects(recordState, proc, entry.call, localGoSubAllocations)
				}
			}
			// ByRef entry proofs must use the same logical-line normalization as
			// VBA227 itself; otherwise a continued Split assignment can make a
			// later call-site argument look unallocated.
			out, _ := a.arrayVBA227Transfer(file, proc, localCtx, variables, in, text, line, constants, nil, nil)
			out = applyArrayLocalGoSubStatementEffects(out, text, localGoSubAllocations)
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				if ownerStatementID > 0 && call.StatementID != ownerStatementID {
					if !filterNestedCalls {
						continue
					}
					owner, ownerOK := baseView.BlockForStatement(call.StatementID)
					if ownerOK && worklistReachable[owner.ID] {
						// This callback is visiting the container statement's
						// source range. Nested calls are visited again by their own
						// CFG block, so applying their post-call effect here would
						// leak one branch into its siblings.
						continue
					}
				}
				out = applyArrayModuleCallEffects(out, file, proc, call, localCtx, variables, moduleDecls)
				out = applyArrayUnknownModuleCallEffects(out, file, proc, call, localCtx, variables, moduleDecls)
				if arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
					out = applyArrayConditionalByRefCallEffects(out, proc, call, localCtx)
				} else {
					out = applyArrayByRefCallEffects(out, proc, call, localCtx)
				}
				out = applyArrayLocalGoSubEffects(out, proc, call, localGoSubAllocations)
			}
			return out
		}
		visit := func(text string, line int, in arrayFlowState) arrayFlowState {
			return visitForBlock(text, line, in, 0, false, false)
		}
		visitBlock := func(block vbacfg.Block, text string, line int, in arrayFlowState) arrayFlowState {
			filterNestedCalls := arrayCFGBlockOwnsNestedStatements(block)
			skipNestedState := false
			if filterNestedCalls && block.Statement != nil {
				start := block.Statement.Range.StartLine
				if start == 0 {
					start = block.Range.StartLine
				}
				skipNestedState = line > start && worklistReachableLines[line]
			}
			return visitForBlock(text, line, in, block.StatementID, filterNestedCalls, skipNestedState)
		}
		if proc.Graph == nil {
			state := initial
			for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
				state = visit(normalizedCodeLine(file.Lines[line-1]), line, state)
			}
			return evidence
		}
		fallbackFacts := buildArraySourceOrderFallbackFacts(file, proc, &baseView, variables, localCtx, constants)
		fallbackFacts.unknownFlow = len(proc.Graph.UnknownFlowSources) > 0
		if ctx.arrayStats != nil {
			ctx.arrayStats.addCFGWalk()
		}
		walkArrayCFGWithSourceLinesReliableStatsAndBlock(&baseView, file.Lines, initial, visit, visitBlock, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			out = applyArrayConditionalAllocationBranch(out, &baseView, block, edge)
			out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
		}, nil, ctx.arrayStats)
		// A recovered source construct can make a call block unreachable in the
		// CFG even though the call is valid VBA source.  The parser currently
		// represents some colon-separated single-line statements this way.  For
		// that narrow boundary, retain a call-site allocation proof from the
		// lexical source order; only direct array arguments with an allocation
		// invariant across all reachable branch alternatives are admitted.
		for call := range proc.Calls.All() {
			block, ok := proc.Graph.BlockForStatement(call.StatementID)
			if !ok || worklistReachable[block.ID] {
				continue
			}
			if !arrayByRefSourceOrderFallbackApplies(file, proc, &baseView, fallbackFacts, call) {
				continue
			}
			key, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
			if !ok || !procedureHasByRefArrayParameter(target) || !arrayProcedureIsParticipant(ctx, target) {
				continue
			}
			if state, proven := a.arrayByRefCallSourceOrderProof(file, fallbackFacts, localGoSubAllocations, proc, target, call, initial, ctx, variables, constants); proven {
				arrayRecordByRefCall(evidence, key, target, proc, call, file, state, ctx)
			}
		}
		return evidence
	}

	dependents := make(map[string][]int)
	indexByKey := make(map[string]int, len(callers))
	for index, caller := range callers {
		indexByKey[arrayProcedureKey(caller.proc)] = index
		for call := range caller.proc.Calls.All() {
			if key, target, ok := arrayPrivateTargetForCall(ctx, targets, call); ok && procedureHasByRefArrayParameter(target) && arrayProcedureIsParticipant(ctx, target) {
				dependents[key] = append(dependents[key], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}
	contributions := make(map[string]map[string]map[int]arrayByRefEntryEvidence, len(callers))
	entries := map[string]map[int]bool{}
	conditions := map[string]map[int]string{}
	queue := make([]int, len(callers))
	queued := make([]bool, len(callers))
	for index := range callers {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(callers) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		caller := callers[index]
		key := arrayProcedureKey(caller.proc)
		evidence := evaluateCaller(caller, entries, conditions)
		if arrayByRefEvidenceMapsEqual(contributions[key], evidence) {
			continue
		}
		contributions[key] = evidence
		merged := map[string]map[int]arrayByRefEntryEvidence{}
		for _, callerEvidence := range contributions {
			mergeArrayByRefEntryEvidence(merged, callerEvidence)
		}
		result := map[string]map[int]bool{}
		conditionalResult := map[string]map[int]string{}
		for targetKey, parameters := range merged {
			for parameterIndex, fact := range parameters {
				if !fact.seen {
					continue
				}
				if fact.allocated {
					if result[targetKey] == nil {
						result[targetKey] = map[int]bool{}
					}
					result[targetKey][parameterIndex] = true
				}
				if !fact.allocated && fact.conditionCompatible && fact.condition != "" {
					if conditionalResult[targetKey] == nil {
						conditionalResult[targetKey] = map[int]string{}
					}
					conditionalResult[targetKey][parameterIndex] = fact.condition
				}
			}
		}
		changedTargets := arrayByRefEntryChangedTargets(entries, conditions, result, conditionalResult)
		if len(changedTargets) == 0 {
			continue
		}
		sortArrayProcedureKeys(changedTargets, indexByKey)
		entries = result
		conditions = conditionalResult
		for _, target := range changedTargets {
			if dependent, ok := indexByKey[target]; ok && !queued[dependent] {
				queued[dependent] = true
				queue = append(queue, dependent)
			}
			for _, dependent := range dependents[target] {
				if !queued[dependent] {
					queued[dependent] = true
					queue = append(queue, dependent)
				}
			}
		}
	}
	return entries, conditions
}

func arrayByRefEntryStatesEqual(left, right map[string]map[int]bool) bool {
	if len(left) != len(right) {
		return false
	}
	for target, parameters := range left {
		other, ok := right[target]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index := range parameters {
			if !other[index] {
				return false
			}
		}
	}
	return true
}

func arrayByRefEntryConditionsEqual(left, right map[string]map[int]string) bool {
	if len(left) != len(right) {
		return false
	}
	for target, parameters := range left {
		other, ok := right[target]
		if !ok || len(parameters) != len(other) {
			return false
		}
		for index, condition := range parameters {
			if other[index] != condition {
				return false
			}
		}
	}
	return true
}

func arrayByRefEvidenceMapsEqual(left, right map[string]map[int]arrayByRefEntryEvidence) bool {
	if len(left) != len(right) {
		return false
	}
	for target, leftParameters := range left {
		rightParameters, ok := right[target]
		if !ok || len(leftParameters) != len(rightParameters) {
			return false
		}
		for index, leftFact := range leftParameters {
			if rightParameters[index] != leftFact {
				return false
			}
		}
	}
	return true
}

func mergeArrayByRefEntryEvidence(dst, src map[string]map[int]arrayByRefEntryEvidence) {
	for target, parameters := range src {
		for index, incoming := range parameters {
			if dst[target] == nil {
				dst[target] = map[int]arrayByRefEntryEvidence{}
			}
			current, exists := dst[target][index]
			if !exists {
				dst[target][index] = incoming
				continue
			}
			current.allocated = current.allocated && incoming.allocated
			if !incoming.allocated && incoming.condition == "" {
				current.conditionCompatible = false
			}
			if incoming.condition != "" {
				if current.condition == "" {
					current.condition = incoming.condition
				} else if !strings.EqualFold(current.condition, incoming.condition) {
					current.conditionCompatible = false
				}
			}
			current.conditionCompatible = current.conditionCompatible && incoming.conditionCompatible
			current.seen = current.seen || incoming.seen
			dst[target][index] = current
		}
	}
}

func arrayByRefEntryChangedTargets(oldEntries map[string]map[int]bool, oldConditions map[string]map[int]string, newEntries map[string]map[int]bool, newConditions map[string]map[int]string) []string {
	keys := map[string]bool{}
	for target := range oldEntries {
		keys[target] = true
	}
	for target := range oldConditions {
		keys[target] = true
	}
	for target := range newEntries {
		keys[target] = true
	}
	for target := range newConditions {
		keys[target] = true
	}
	changed := make([]string, 0, len(keys))
	for target := range keys {
		oldEntry := map[string]map[int]bool{}
		newEntry := map[string]map[int]bool{}
		if names := oldEntries[target]; len(names) > 0 {
			oldEntry[target] = names
		}
		if names := newEntries[target]; len(names) > 0 {
			newEntry[target] = names
		}
		oldCondition := map[string]map[int]string{}
		newCondition := map[string]map[int]string{}
		if conditions := oldConditions[target]; len(conditions) > 0 {
			oldCondition[target] = conditions
		}
		if conditions := newConditions[target]; len(conditions) > 0 {
			newCondition[target] = conditions
		}
		if !arrayByRefEntryStatesEqual(oldEntry, newEntry) || !arrayByRefEntryConditionsEqual(oldCondition, newCondition) {
			changed = append(changed, target)
		}
	}
	return changed
}

func sortArrayProcedureKeys(keys []string, indexByKey map[string]int) {
	sort.SliceStable(keys, func(i, j int) bool {
		left, leftOK := indexByKey[keys[i]]
		right, rightOK := indexByKey[keys[j]]
		if leftOK && rightOK && left != right {
			return left < right
		}
		if leftOK != rightOK {
			return leftOK
		}
		return keys[i] < keys[j]
	})
}

func arrayCallsAtLine(calls readOnlySpan[procedureir.CallSite], line int) []procedureir.CallSite {
	matched := make([]procedureir.CallSite, 0, 1)
	for call := range calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine != line {
			continue
		}
		matched = append(matched, call)
	}
	return matched
}

func arrayPrivateTargetForCall(ctx analysisContext, targets map[string]sourceProcedure, call procedureir.CallSite) (string, sourceProcedure, bool) {
	resolution := arrayCallResolution(ctx, call)
	if resolution.Status != procedureir.ResolutionMatched || len(resolution.Candidates) != 1 {
		return "", sourceProcedure{}, false
	}
	key := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
	target, ok := targets[key]
	return key, target, ok
}

// arrayCallResolution rechecks the one parser shape that currently loses an
// unparenthesized call's procedure name when its implicit member argument is
// written with a leading dot, for example `Consume .values`.  The ordinary
// resolver remains authoritative; the retry is deliberately limited to a
// single dotted argument so an unresolved object member is not reinterpreted
// as a project-local procedure call.
func arrayCallResolution(ctx analysisContext, call procedureir.CallSite) procedureir.CallResolution {
	resolution := call.Resolution
	if ctx.procedureResolver == nil {
		return resolution
	}
	resolution = ctx.procedureResolver.ResolveCall(call)
	if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
		return resolution
	}
	name := arrayImplicitMemberArgumentCallName(call.Callee.Text)
	if name == "" || strings.EqualFold(name, call.Callee.BaseName) {
		return resolution
	}
	retried := call
	retried.Callee.Text = name
	retried.Callee.BaseName = name
	retried.Callee.Member = name
	retried.Callee.Receiver = nil
	return ctx.procedureResolver.ResolveCall(retried)
}

func arrayImplicitMemberArgumentCallName(text string) string {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "call ") {
		text = strings.TrimSpace(text[len("call "):])
	}
	if text == "" || !isIdentifierStart(text[0]) {
		return ""
	}
	end := 1
	for end < len(text) && isIdentifierPart(text[end]) {
		end++
	}
	rest := strings.TrimSpace(text[end:])
	if len(rest) < 2 || rest[0] != '.' || !isIdentifierStart(rest[1]) {
		return ""
	}
	memberEnd := 2
	for memberEnd < len(rest) && isIdentifierPart(rest[memberEnd]) {
		memberEnd++
	}
	if strings.TrimSpace(rest[memberEnd:]) != "" {
		return ""
	}
	return cleanIdentifier(text[:end])
}

func procedureHasByRefArrayParameter(proc sourceProcedure) bool {
	for parameter := range proc.Params.All() {
		if parameterIsByRefArray(parameter) {
			return true
		}
	}
	return false
}

func arrayRecordByRefCall(evidence map[string]map[int]arrayByRefEntryEvidence, targetKey string, target, caller sourceProcedure, call procedureir.CallSite, file parsedFile, state arrayFlowState, ctx analysisContext) {
	// A self-recursive ByRef helper preserves the entry array state supplied by
	// its caller. Treating the recursive edge as an independent unknown entry
	// would poison the evidence from the allocated external call and keep the
	// callee conservative forever.
	if strings.EqualFold(targetKey, arrayProcedureKey(caller)) {
		return
	}
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return
	}
	arguments := make([]string, target.Params.Len())
	bound := make(map[int]bool, len(bindings))
	for _, binding := range bindings {
		arguments[binding.parameterIndex] = binding.text
		bound[binding.parameterIndex] = true
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) || !bound[index] {
			continue
		}
		name := ""
		name = directArrayArgumentName(arguments[index])
		value, known := state[name]
		allocated := known && value.kind == arrayAllocated && value.knownArray
		if !allocated && arrayQualifiedArgumentProvenAllocated(file, caller, call, arguments[index], ctx) {
			value = arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}
			known = true
			allocated = true
		}
		_, _, qualifiedMember := arrayQualifiedMemberParts(arguments[index])
		if !allocated && !qualifiedMember && index < len(arguments) {
			// A function returning a dynamic array is a valid ByRef array
			// argument in VBA.  The identifier-only path above cannot attach
			// that expression to a caller state entry, so consult the existing
			// array-return summaries before treating the callee entry as
			// unallocated.  Unknown or conditionally allocated returns remain
			// conservative because arrayExpressionState returns no allocated
			// proof for them.
			if returned, returnedKnown := arrayExpressionState(arguments[index], state, ctx); returnedKnown && returned.kind == arrayAllocated && returned.knownArray {
				value = returned
				known = true
				allocated = true
			}
		}
		condition := ""
		if known && !allocated {
			if source, positive := arrayVBA227PositiveLengthCondition(value.conditionalAllocationSource); positive {
				condition = arrayConditionalEntrySource(target, arguments, index, source)
			}
			if condition == "" && value.allocationCountSource != "" {
				condition = arrayConditionalEntrySource(target, arguments, index, value.allocationCountSource)
			}
		}
		if !allocated && condition == "" && arrayByRefCallArrayVacuouslyUnused(target, index, arguments) {
			// A call may intentionally pass an unallocated optional array (including
			// an array-return expression with no local state entry) to a helper whose
			// only uses are behind a false literal guard. It must not poison the
			// conditional-entry evidence collected from calls that do reach those
			// uses.
			continue
		}
		parameters := evidence[targetKey]
		if parameters == nil {
			parameters = map[int]arrayByRefEntryEvidence{}
		}
		fact := parameters[index]
		if !fact.seen {
			fact.allocated = allocated
			fact.conditionCompatible = allocated || condition != ""
			fact.condition = condition
		} else {
			fact.allocated = fact.allocated && allocated
			if !allocated && condition == "" {
				fact.conditionCompatible = false
			}
			if condition != "" {
				if fact.condition == "" {
					fact.condition = condition
				} else if !strings.EqualFold(fact.condition, condition) {
					fact.conditionCompatible = false
				}
			}
		}
		fact.seen = true
		parameters[index] = fact
		evidence[targetKey] = parameters
	}
}
