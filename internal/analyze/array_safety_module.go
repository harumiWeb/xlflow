package analyze

import (
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/gui"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func inferArrayModuleInvalidationSummaries(files []parsedFile, ctx analysisContext) arrayModuleInvalidationSummaries {
	summaries := arrayModuleInvalidationSummaries{}
	ctx.arrayModuleInvalidations = summaries
	ctx.arrayModuleInvalidationCacheWritable = true
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		moduleDecls := file.moduleDecls()
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			key := arrayProcedureKey(proc)
			if key == "" {
				continue
			}
			if _, cached := summaries[key]; cached {
				continue
			}
			summaries[key] = arrayPrivateModuleArrayInvalidationsWithVisiting(file, proc, moduleDecls, ctx, map[string]bool{})
		}
	}
	return summaries
}
func inferArrayModuleConfigurationStates(files []parsedFile, summaries arrayModuleAllocationSummaries) map[string]arrayModuleConfigurationState {
	states := map[string]arrayModuleConfigurationState{}
	for _, file := range files {
		state := arrayModuleConfigurationState{byProcedure: map[string]map[string]bool{}}
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			name := strings.ToLower(strings.TrimSpace(proc.Name))
			if !strings.HasPrefix(name, "configure") {
				continue
			}
			arrays := summaries[arrayProcedureKey(proc)]
			if len(arrays) == 0 {
				continue
			}
			state.byProcedure[name] = cloneArrayNameSet(arrays)
			switch name {
			case "configuredatatable":
				state.dataTable = mergeArrayNameSets(state.dataTable, arrays)
			case "configuregenericcollection":
				state.genericCollection = mergeArrayNameSets(state.genericCollection, arrays)
			}
		}
		if len(state.byProcedure) > 0 {
			states[file.Path] = state
		}
	}
	return states
}

func cloneArrayNameSet(values map[string]bool) map[string]bool {
	if len(values) == 0 {
		return nil
	}
	clone := make(map[string]bool, len(values))
	for name, allocated := range values {
		clone[name] = allocated
	}
	return clone
}

func mergeArrayNameSets(left, right map[string]bool) map[string]bool {
	if len(right) == 0 {
		return left
	}
	if left == nil {
		left = map[string]bool{}
	}
	for name := range right {
		left[name] = true
	}
	return left
}

func applyArrayModuleCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	conditional := arrayProcedureLineHasInlineConditional(file, call.Range.StartLine)
	if conditional && arrayProcedureLineInlineConditionIsFalse(file, call.Range.StartLine) {
		return state
	}
	key, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !ok {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	// Callers pass a block-local state to this transfer callback. The legacy
	// walkers keep predecessor input separate before invoking it, while the
	// compact cursor owns the map outright; update that private state in place
	// so module/ByRef effects do not reintroduce whole-state copies.
	updated := state
	if updated == nil {
		updated = arrayFlowState{}
	}
	markArgument := func(name string) {
		name = strings.ToLower(cleanIdentifier(name))
		variable, known := variables[name]
		if !known || !variable.isArray {
			return
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	markModule := func(name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			return
		}
		declaration, declared := moduleDecls[name]
		if !declared || !declaration.Array || declaration.Parameter {
			return
		}
		markArgument(name)
	}
	if !conditional {
		for name := range ctx.arrayModuleAllocations[key] {
			markModule(name)
		}
	}
	arguments, mapped := arrayCallFormalArguments(proc, target, call)
	if mapped && !conditional {
		for index := range ctx.arrayByRefAllocations[key] {
			if index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
				continue
			}
			markArgument(arguments[index])
		}
		for outputIndex, countIndex := range ctx.arrayByRefConditionalAllocations[key] {
			if outputIndex < 0 || outputIndex >= target.Params.Len() || countIndex < 0 || countIndex >= target.Params.Len() {
				continue
			}
			outputName := directArrayArgumentName(arguments[outputIndex])
			countSource := directArrayArgumentName(arguments[countIndex])
			if outputName == "" || countSource == "" {
				continue
			}
			name := strings.ToLower(outputName)
			variable, known := variables[name]
			if !known || !variable.isArray {
				continue
			}
			value := updated[name]
			if value.kind == arrayAllocated && value.knownArray {
				continue
			}
			if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, countSource) {
				continue
			}
			value.allocationCountSource = countSource
			updated[name] = value
		}
		for outputIndex, lengthIndex := range ctx.arrayByRefLengthAllocations[key] {
			if outputIndex >= len(arguments) || lengthIndex < 0 || lengthIndex >= len(arguments) {
				continue
			}
			outputName := directArrayArgumentName(arguments[outputIndex])
			lengthSource := directArrayArgumentName(arguments[lengthIndex])
			if outputName == "" || lengthSource == "" {
				continue
			}
			name := strings.ToLower(outputName)
			variable, known := variables[name]
			if !known || !variable.isArray {
				continue
			}
			value := updated[name]
			if value.kind == arrayAllocated && value.knownArray {
				continue
			}
			if value.allocationCountSource != "" && !strings.EqualFold(value.allocationCountSource, lengthSource) {
				continue
			}
			value.allocationCountSource = lengthSource
			updated[name] = value
		}
		for name := range arrayConfigurationArraysForGuard(file, target, arguments, ctx.arrayModuleConfigurations[file.Path]) {
			markModule(name)
		}
	}
	if !ctx.arraySkipModuleInvalidationEffects {
		invalidated := arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)
		for name := range invalidated {
			if declarations.shadowsModule(name) {
				continue
			}
			if _, tracked := updated[name]; tracked {
				updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
			}
		}
	}
	return updated
}

// arrayPrivateModuleArrayInvalidations returns the normal-return invalidation
// summary for a project-local helper. A precomputed summary is used by the
// production context; focused compatibility callers compute the same summary
// on demand.
func arrayPrivateModuleArrayInvalidations(file parsedFile, target sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	key := arrayProcedureKey(target)
	if ctx.arrayModuleInvalidations != nil {
		if summary, ok := ctx.arrayModuleInvalidations[key]; ok {
			return summary
		}
	}
	return arrayPrivateModuleArrayInvalidationsWithVisiting(file, target, moduleDecls, ctx, map[string]bool{})
}

// arrayPrivateModuleArrayInvalidationsWithVisiting identifies module arrays
// that are not proven allocated at a target's normal exit. The summary starts
// from allocated arrays so it models the effect on a caller that already has
// a valid module-array allocation. Direct operations are evaluated through the
// normal CFG, and resolved local calls are summarized recursively.
func arrayPrivateModuleArrayInvalidationsWithVisiting(file parsedFile, target sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, visiting map[string]bool) map[string]bool {
	key := arrayProcedureKey(target)
	if !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
		return nil
	}
	names := arrayModuleNamesForProcedure(file, target, moduleDecls)
	if len(names) == 0 {
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = nil
		}
		return nil
	}
	if visiting[key] {
		// A recursive effect cycle has no finite normal-return proof. Keep all
		// visible module arrays conservative rather than assuming the cycle is
		// read-only.
		return names
	}
	if ctx.arrayModuleInvalidations != nil {
		if summary, ok := ctx.arrayModuleInvalidations[key]; ok {
			return summary
		}
	}
	visiting[key] = true
	defer delete(visiting, key)

	variables := arrayVariables(file, target, moduleDecls)
	initial := arrayInitialState(variables)
	for name := range names {
		value := initial[name]
		value.kind = arrayAllocated
		value.knownArray = true
		initial[name] = value
	}
	if target.Graph == nil {
		state := initial
		constants := arrayIntegerConstants(file, target, nil, nil)
		if target.Statements.Len() > 0 {
			for statement := range target.Statements.All() {
				line := statement.Range.StartLine
				if line < 1 {
					line = target.StartLine
				}
				text := strings.TrimSpace(normalizedCodeLine(statement.Text))
				if text == "" && line >= 1 && line <= len(file.Lines) {
					text = normalizedCodeLine(file.Lines[line-1])
				}
				if text == "" {
					continue
				}
				state = arrayModuleSummaryTransfer(file, target, ctx, variables, state, text, line, constants, moduleDecls, names, visiting)
			}
		} else {
			for line := target.StartLine; line <= target.EndLine && line <= len(file.Lines); line++ {
				state = arrayModuleSummaryTransfer(file, target, ctx, variables, state, normalizedCodeLine(file.Lines[line-1]), line, constants, moduleDecls, names, visiting)
			}
		}
		result := arrayModuleInvalidationsFromState(names, state)
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = result
		}
		return result
	}

	graph := target.Graph.View(vbacfg.EdgeFilter{NormalOnly: true, WithoutNormalErrRaiseContinuation: true})
	constants := arrayIntegerConstants(file, target, nil, nil)
	var normalExit arrayFlowState
	hasNormalExit := false
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		return arrayModuleSummaryTransfer(file, target, ctx, variables, in, text, line, constants, moduleDecls, names, visiting)
	}
	edgeState := func(_ vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		if edge.To != graph.NormalExit() {
			return out
		}
		if !hasNormalExit {
			normalExit = cloneArrayState(out)
			hasNormalExit = true
		} else {
			normalExit = meetArrayState(normalExit, out)
		}
		return out
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, nil, ctx.arrayStats)
	if !hasNormalExit {
		if ctx.arrayModuleInvalidationCacheWritable && key != "" {
			ctx.arrayModuleInvalidations[key] = nil
		}
		return nil
	}
	result := arrayModuleInvalidationsFromState(names, normalExit)
	if ctx.arrayModuleInvalidationCacheWritable && key != "" {
		ctx.arrayModuleInvalidations[key] = result
	}
	return result
}

func arrayModuleInvalidationsFromState(names map[string]bool, state arrayFlowState) map[string]bool {
	invalidated := map[string]bool{}
	for name := range names {
		value, known := state[name]
		if !known || value.kind != arrayAllocated || !value.knownArray {
			invalidated[name] = true
		}
	}
	if len(invalidated) == 0 {
		return nil
	}
	return invalidated
}

func arrayModuleSummaryTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	if condition, body, ok := arrayIfThenParts(text); ok && strings.TrimSpace(body) != "" {
		thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
		condition = strings.TrimSpace(condition)
		lowerCondition := strings.ToLower(condition)
		switch {
		case strings.HasPrefix(lowerCondition, "if "):
			condition = strings.TrimSpace(condition[len("if "):])
		case strings.HasPrefix(lowerCondition, "elseif "):
			condition = strings.TrimSpace(condition[len("elseif "):])
		}
		thenState := arrayModuleSummaryTransferParts(file, proc, ctx, variables, cloneArrayState(state), thenBody, line, constants, moduleDecls, moduleArrays, visiting)
		elseState := cloneArrayState(state)
		if hasElse {
			elseState = arrayModuleSummaryTransferParts(file, proc, ctx, variables, elseState, elseBody, line, constants, moduleDecls, moduleArrays, visiting)
		}
		result := meetArrayState(thenState, elseState)
		applyCalls := func(value arrayFlowState, conditionValue, conditionKnown bool) arrayFlowState {
			for _, call := range arrayCallsAtLine(proc.Calls, line) {
				if conditionKnown && !arrayInlineConditionalCallIsReachable(file, call, conditionValue, hasElse) {
					continue
				}
				value = applyArrayModuleSummaryCallEffects(value, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
			}
			return value
		}
		if value, known := arraySourceOrderConstantBoolean(condition, constants); known {
			if value {
				result = applyCalls(thenState, true, true)
			} else if hasElse {
				result = applyCalls(elseState, false, true)
			} else {
				result = state
			}
		} else if !hasElse {
			result = applyCalls(meetArrayState(thenState, state), false, false)
		} else {
			result = applyCalls(result, false, false)
		}
		return result
	}
	state = arrayModuleSummaryTransferParts(file, proc, ctx, variables, state, text, line, constants, moduleDecls, moduleArrays, visiting)
	callsAtLine := arrayCallsAtLine(proc.Calls, line)
	for _, call := range callsAtLine {
		state = applyArrayModuleSummaryCallEffects(state, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
	}
	return state
}

func arrayModuleSummaryTransferParts(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	for _, part := range splitRangeValueSourceStatements(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		if condition, body, ok := arrayIfThenParts(part); ok && strings.TrimSpace(body) != "" {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			condition = strings.TrimSpace(condition)
			lowerCondition := strings.ToLower(condition)
			switch {
			case strings.HasPrefix(lowerCondition, "if "):
				condition = strings.TrimSpace(condition[len("if "):])
			case strings.HasPrefix(lowerCondition, "elseif "):
				condition = strings.TrimSpace(condition[len("elseif "):])
			}
			thenState := arrayModuleSummaryTransferParts(file, proc, ctx, variables, cloneArrayState(state), thenBody, line, constants, moduleDecls, moduleArrays, visiting)
			elseState := cloneArrayState(state)
			if hasElse {
				elseState = arrayModuleSummaryTransferParts(file, proc, ctx, variables, elseState, elseBody, line, constants, moduleDecls, moduleArrays, visiting)
			}
			if value, known := arraySourceOrderConstantBoolean(condition, constants); known {
				if value {
					state = thenState
				} else if hasElse {
					state = elseState
				}
			} else if hasElse {
				state = meetArrayState(thenState, elseState)
			} else {
				state = meetArrayState(thenState, state)
			}
			continue
		}
		state, _ = (Analyzer{}).arrayTransfer(file, proc, ctx, variables, state, part, line, constants, nil)
	}
	return state
}

func applyArrayModuleSummaryCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, moduleArrays map[string]bool, visiting map[string]bool) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	// The procedure IR also represents an indexed array expression such as
	// `result(index)` or `mZipWork(offset)` as a CallSite. Those expressions
	// are handled by arrayTransfer; they are not procedure calls whose module
	// effects belong in this summary. Without this guard, the resolver can
	// reinterpret the expression as a source-local target and recursively
	// invalidate every module array while summarizing an otherwise read-only
	// indexed assignment.
	if call.Callee.Receiver == nil {
		name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
		if variable, ok := variables[name]; ok && (variable.isArray || variable.isVariant) {
			return state
		}
	}
	key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
	if !resolved {
		target, resolved = arraySourceModuleTargetForCall(file, call, ctx)
		if !resolved {
			return state
		}
		key = arrayProcedureKey(target)
	}
	if !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
		return state
	}
	updated := cloneArrayState(state)
	if !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
		for name := range ctx.arrayModuleAllocations[key] {
			name = strings.ToLower(cleanIdentifier(name))
			if moduleArrays[name] {
				value := updated[name]
				value.kind = arrayAllocated
				value.knownArray = true
				updated[name] = value
			}
		}
	}
	for name := range arrayPrivateModuleArrayInvalidationsWithVisiting(file, target, moduleDecls, ctx, visiting) {
		if moduleArrays[name] {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	if !procedureHasByRefArrayParameter(target) {
		return updated
	}
	bindings, mapped := arrayCallArgumentBindings(proc, target, call)
	if !mapped {
		return updated
	}
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		name := strings.ToLower(directArrayArgumentName(binding.text))
		if !moduleArrays[name] {
			continue
		}
		if ctx.arrayByRefAllocations[key][binding.parameterIndex] {
			updated[name] = arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

func arrayPrivateCallMayInvalidateModuleArray(file parsedFile, caller sourceProcedure, target sourceProcedure, call procedureir.CallSite, moduleDecls map[string]sourceDeclaration, ctx analysisContext) bool {
	if len(arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)) > 0 {
		return true
	}
	if !procedureHasByRefArrayParameter(target) {
		return false
	}
	bindings, mapped := arrayCallArgumentBindings(caller, target, call)
	if !mapped {
		return true
	}
	for _, binding := range bindings {
		name := strings.ToLower(directArrayArgumentName(binding.text))
		declaration, declared := moduleDecls[name]
		if !declared || !declaration.Array || declaration.Parameter || binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(binding.parameterIndex)) {
			continue
		}
		if arrayByRefParameterMayInvalidate(target, binding.parameterIndex, ctx, map[string]bool{}) {
			return true
		}
	}
	return false
}

// applyArrayUnknownModuleCallEffects is used by the recovered source-order
// path, where a public call cannot be matched to a private module effect
// summary. A source-local public target can still be inspected for direct
// module-array effects, including when it receives no explicit array
// argument. Calls that do not identify a source-local target are left alone:
// treating every unresolved/external call as a mutation of every visible
// module array turns unrelated object construction and host calls into false
// positives.
func applyArrayUnknownModuleCallEffects(state arrayFlowState, file parsedFile, proc sourceProcedure, call procedureir.CallSite, ctx analysisContext, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
		return state
	}
	if ctx.arraySkipModuleInvalidationEffects {
		return state
	}
	if arrayProcedureLineInlineConditionIsFalse(file, call.Range.StartLine) {
		return state
	}
	if _, _, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call); resolved {
		return state
	}
	target, ok := arraySourceModuleTargetForCall(file, call, ctx)
	if !ok || !procedureUsesModuleArray(file, target, moduleDecls) {
		return state
	}
	invalidated := arrayPrivateModuleArrayInvalidations(file, target, moduleDecls, ctx)
	if len(invalidated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range invalidated {
		if declarations.shadowsModule(name) {
			continue
		}
		variable, known := variables[name]
		if !known || !variable.isArray {
			continue
		}
		if _, tracked := updated[name]; tracked {
			updated[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
		}
	}
	return updated
}

// arraySourceModuleTargetForCall recovers a source-local target for the
// source-order fallback. The project resolver may report a public target, but
// it may also be incomplete for a recovered call. In the latter case a unique
// same-module source procedure is sufficient evidence for inspecting that
// procedure's module-array accesses; an absent target remains an external or
// late-bound call and must not invalidate all module arrays.
func arraySourceModuleTargetForCall(file parsedFile, call procedureir.CallSite, ctx analysisContext) (sourceProcedure, bool) {
	procedures := file.procedureView()
	if procedures.Len() == 0 {
		return sourceProcedure{}, false
	}
	resolution := call.Resolution
	if ctx.procedureResolver != nil {
		resolution = ctx.procedureResolver.ResolveCall(call)
	}
	if resolution.Status == procedureir.ResolutionMatched && len(resolution.Candidates) == 1 {
		qualifiedName := strings.ToLower(strings.TrimSpace(resolution.Candidates[0].QualifiedName))
		for index := 0; index < procedures.Len(); index++ {
			target := procedures.valueAt(index)
			if strings.EqualFold(arrayProcedureKey(target), qualifiedName) {
				return target, true
			}
		}
	}
	// A receiver-bearing call that was not resolved to the exact source target
	// belongs to another object or to a late-bound member. Do not reinterpret a
	// same-module procedure with the same member name as its target.
	if call.Callee.Receiver != nil {
		return sourceProcedure{}, false
	}

	baseName := cleanIdentifier(strings.TrimPrefix(strings.TrimSpace(call.Callee.BaseName), "New "))
	if baseName == "" {
		baseName = cleanIdentifier(strings.TrimPrefix(strings.TrimSpace(call.Callee.Text), "New "))
	}
	callerModule := strings.TrimSpace(call.Caller.QualifiedName)
	if dot := strings.IndexByte(callerModule, '.'); dot >= 0 {
		callerModule = callerModule[:dot]
	}
	if callerModule == "" {
		callerModule = strings.TrimSpace(call.Module)
	}
	if callerModule == "" {
		callerModule = strings.TrimSpace(file.Module)
	}
	var match sourceProcedure
	matched := 0
	for index := 0; index < procedures.Len(); index++ {
		candidate := procedures.valueAt(index)
		if !strings.EqualFold(strings.TrimSpace(candidate.Name), baseName) ||
			!strings.EqualFold(strings.TrimSpace(candidate.Module), callerModule) {
			continue
		}
		visibility := strings.TrimSpace(candidate.Visibility)
		if strings.EqualFold(visibility, "Private") || strings.EqualFold(visibility, "Friend") {
			continue
		}
		match = candidate
		matched++
	}
	if matched == 1 {
		return match, true
	}
	return sourceProcedure{}, false
}

func arrayConfigurationArraysForGuard(file parsedFile, target sourceProcedure, arguments []string, configurations arrayModuleConfigurationState) map[string]bool {
	name := strings.ToLower(strings.TrimSpace(target.Name))
	if !strings.HasPrefix(name, "require") || !arrayGuardProcedureRejectsInvalidState(file, target) || !arrayGuardTargetsCurrentObject(target, arguments) {
		return nil
	}
	if arrays := configurations.byProcedure["configure"+strings.TrimPrefix(name, "require")]; len(arrays) > 0 {
		return arrays
	}
	if name == "requireerror" {
		if arrays := configurations.byProcedure["configureaggregateerror"]; len(arrays) > 0 {
			return arrays
		}
	}
	body := strings.ToLower(strings.Join(file.Lines[max(0, target.StartLine-1):min(len(file.Lines), target.EndLine)], "\n"))
	if strings.Contains(body, "role_data_table") {
		return configurations.dataTable
	}
	if arrayGuardUsesGenericCollectionConfiguration(body) {
		return configurations.genericCollection
	}
	return nil
}

// applyArrayInternalStorageConfiguration carries the class-instance array
// contract into Friend/Private storage members that are called through a
// configured receiver. These members intentionally do not repeat a public
// role guard: their callers have already established the owning collection,
// data-row, or aggregate-error configuration on that receiver.
func applyArrayInternalStorageConfiguration(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, configurations arrayModuleConfigurationState) arrayFlowState {
	if !strings.EqualFold(strings.TrimSpace(proc.ModuleKind), "class") {
		return state
	}
	arrays := arrayInternalStorageConfigurationArrays(proc, configurations)
	if len(arrays) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayInternalStorageConfigurationArrays(proc sourceProcedure, configurations arrayModuleConfigurationState) map[string]bool {
	name := strings.ToLower(strings.TrimSpace(proc.Name))
	switch name {
	case "internalcollectionitems", "internalcollectionkeys", "internalcollectionpriorities",
		"internaladdlookupgroup", "internalappendcollectionitem", "internalappendcollectionkey",
		"internalappendcollectionpriority", "internalqueuevalue", "internalpushvalue":
		return configurations.genericCollection
	case "internaladdwrapped":
		return mergeArrayNameSets(cloneArrayNameSet(configurations.byProcedure["configurelist"]), configurations.genericCollection)
	case "internaldatacolumns":
		return configurations.dataTable
	case "internaldatarows", "internalappendrowcell", "acceptrowchanges", "rejectrowchanges":
		return configurations.byProcedure["configuredatarow"]
	case "internalinnerexceptions":
		return configurations.byProcedure["configureaggregateerror"]
	default:
		return nil
	}
}

func arrayGuardUsesGenericCollectionConfiguration(body string) bool {
	for _, marker := range []string{
		"isgenericcollectionrole",
		"ispriorityqueuekind",
		"isdictionarycollection",
		"issetcollection",
		"mcollectionkind",
		"ROLE_DICTIONARY",
		"ROLE_HASH_SET",
		"ROLE_COLLECTION",
		"ROLE_IMMUTABLE",
		"ROLE_CONCURRENT",
	} {
		if strings.Contains(body, strings.ToLower(marker)) {
			return true
		}
	}
	return false
}

func arrayGuardProcedureRejectsInvalidState(file parsedFile, target sourceProcedure) bool {
	start := max(0, target.StartLine-1)
	end := min(len(file.Lines), target.EndLine)
	if start >= end {
		return false
	}
	body := strings.ToLower(strings.Join(file.Lines[start:end], "\n"))
	return strings.Contains(body, "err.raise") || strings.Contains(body, "raisecontracterror")
}

func arrayGuardTargetsCurrentObject(target sourceProcedure, arguments []string) bool {
	if target.Params.Len() <= 1 {
		return true
	}
	return len(arguments) > 0 && strings.EqualFold(strings.TrimSpace(arguments[0]), "me")
}

func applyArrayModuleConfigurationBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, configurations arrayModuleConfigurationState, variables map[string]arrayVariable, file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchTrue || statement.Condition == nil {
		return state
	}
	condition := strings.ToLower(strings.TrimSpace(statement.Condition.Text))
	var arrays map[string]bool
	switch {
	case arrayPositiveGenericCollectionKindBranch(condition):
		arrays = configurations.genericCollection
	case strings.Contains(condition, "role_immutable") && !strings.Contains(condition, "<> role_immutable"):
		arrays = configurations.genericCollection
	case strings.Contains(condition, "role_list") && !strings.Contains(condition, "<> role_list"):
		arrays = configurations.byProcedure["configurelist"]
	case strings.Contains(condition, "role_data_row") && !strings.Contains(condition, "<> role_data_row"):
		arrays = configurations.byProcedure["configuredatarow"]
	case strings.Contains(condition, "role_data_table") && !strings.Contains(condition, "<> role_data_table"):
		arrays = configurations.dataTable
	}
	if len(arrays) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayPositiveGenericCollectionKindBranch(condition string) bool {
	for _, marker := range []string{
		"isgenericcollectionrole",
		"ispriorityqueuekind",
		"issortedmapkind",
		"issortedsetkind",
	} {
		if strings.Contains(condition, marker) && !strings.Contains(condition, "not "+marker) {
			return true
		}
	}
	return false
}

func inferArrayModuleAllocationSummaries(files []parsedFile, ctx analysisContext, targets map[string]sourceProcedure, byRefSummaries arrayByRefAllocationSummaries) arrayModuleAllocationSummaries {
	summaries := arrayModuleAllocationSummaries{}
	procedures := make([]struct {
		file        parsedFile
		proc        sourceProcedure
		moduleDecls map[string]sourceDeclaration
	}, 0)
	// Package-local synthetic callers may omit ModuleFacts. Attach one local
	// immutable instance before expanding the procedure list so compatibility
	// paths do not rebuild and rescan module source once per procedure.
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			procedures = append(procedures, struct {
				file        parsedFile
				proc        sourceProcedure
				moduleDecls map[string]sourceDeclaration
			}{file: file, proc: proc, moduleDecls: moduleDecls})
		}
	}
	if len(procedures) == 0 {
		return summaries
	}
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	dominators := arrayProcedureDominators{}
	for _, procedure := range procedures {
		dominators[arrayProcedureKey(procedure.proc)] = arrayProcedureNormalExitDominators(procedure.proc)
	}

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
	contributions := make(arrayModuleAllocationSummaries, len(procedures))
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		procedure := procedures[index]
		if !arrayProcedureIsParticipant(ctx, procedure.proc) {
			continue
		}
		key := arrayProcedureKey(procedure.proc)
		value := arrayModuleAllocationSummaryForProcedure(procedure.file, procedure.proc, procedure.moduleDecls, targets, summaries, byRefSummaries, ctx, dominators[key])
		old := arrayModuleAllocationSummaries{key: contributions[key]}
		fresh := arrayModuleAllocationSummaries{key: value}
		if arrayModuleAllocationSummariesEqual(old, fresh) {
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

func arrayModuleAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, targets map[string]sourceProcedure, summaries arrayModuleAllocationSummaries, byRefSummaries arrayByRefAllocationSummaries, ctx analysisContext, dominators map[vbacfg.BlockID]bool) map[string]bool {
	moduleArrays := map[string]bool{}
	for name, declaration := range moduleDecls {
		if declaration.Array && !declaration.Parameter {
			moduleArrays[strings.ToLower(name)] = true
		}
	}
	if len(moduleArrays) == 0 || proc.Graph == nil {
		return nil
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	for name := range moduleArrays {
		if declarations.shadowsModule(name) {
			delete(moduleArrays, name)
		}
	}
	idempotentSetupArrays := arrayModuleIdempotentSetupArrays(file, proc, moduleDecls, ctx)
	allocated := map[string]bool{}
	addDirectAllocation := func(statementID int, name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if !moduleArrays[name] || (!arrayProcedureBlockDominatesNormalExit(proc, statementID, dominators) && !idempotentSetupArrays[name]) {
			return
		}
		allocated[name] = true
	}
	for statement := range proc.Statements.All() {
		text := strings.TrimSpace(statement.Text)
		if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if direct && !arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
					addDirectAllocation(statement.ID, redim.name)
				}
			}
		}
		if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed {
			name := strings.ToLower(cleanIdentifier(lhs))
			if moduleArrays[name] {
				if value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx); known && value.kind == arrayAllocated && value.knownArray {
					if !arrayProcedureLineHasInlineConditional(file, statement.Range.StartLine) {
						addDirectAllocation(statement.ID, name)
					}
				}
			}
		}
	}
	for call := range proc.Calls.All() {
		key, target, ok := arrayPrivateTargetForCall(ctx, targets, call)
		if !ok {
			continue
		}
		calleeArrays := summaries[key]
		calleeByRefArrays := byRefSummaries[key]
		if len(calleeArrays) == 0 && len(calleeByRefArrays) == 0 {
			continue
		}
		guaranteed := !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) && arrayProcedureBlockDominatesNormalExit(proc, call.StatementID, dominators)
		if !guaranteed && arrayProcedureHasIdempotentSetupGuard(file, proc, call.Range.StartLine, moduleDecls) {
			guaranteed = true
		}
		if !guaranteed {
			continue
		}
		for name := range calleeArrays {
			if moduleArrays[name] {
				allocated[name] = true
			}
		}
		if len(calleeByRefArrays) > 0 {
			arguments, mapped := arrayCallFormalArguments(proc, target, call)
			if mapped {
				for index := range calleeByRefArrays {
					if index >= target.Params.Len() || !parameterIsByRefArray(target.Params.valueAt(index)) {
						continue
					}
					name := strings.ToLower(directArrayArgumentName(arguments[index]))
					if moduleArrays[name] {
						allocated[name] = true
					}
				}
			}
		}
	}
	return allocated
}

// arrayModuleIdempotentSetupArrays recognizes the narrow one-time module
// initialization idiom used by private helper routines:
//
//	If ready Then Exit Sub
//	ReDim values(...)
//	ready = True
//
// The direct ReDim does not dominate the procedure's normal exit because the
// already-initialized branch exits early.  The summary can nevertheless carry
// the allocation when the Boolean guard is module-scoped, is written only to
// True by this procedure, is not written elsewhere in the module, and is the
// final executable statement.  These constraints keep an arbitrary Boolean
// branch from becoming an allocation proof.
func arrayModuleIdempotentSetupArrays(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	if proc.StartLine < 1 || proc.EndLine < proc.StartLine || proc.StartLine > len(file.Lines) {
		return nil
	}
	end := min(len(file.Lines), proc.EndLine)
	start := max(0, proc.StartLine-1)
	facts := file.moduleAnalysisFacts()
	type setupGuard struct {
		name    string
		checkAt int
	}
	guards := make([]setupGuard, 0)
	for index := start; index < end; index++ {
		name, ok := facts.sourceLineSetupGuard(index)
		if !ok {
			continue
		}
		declaration, ok := moduleDecls[name]
		if !ok || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			continue
		}
		guards = append(guards, setupGuard{name: name, checkAt: index})
	}
	if len(guards) == 0 {
		return nil
	}
	constants := arrayIntegerConstants(file, proc, nil, nil)

	lastExecutable := -1
	for index := start; index < end; index++ {
		if facts.sourceLineIsExecutable(index) {
			lastExecutable = index
		}
	}
	if lastExecutable < start {
		return nil
	}

	guardWrites := map[string][]struct {
		line int
		rhs  string
	}{}
	for _, guard := range guards {
		facts.forEachArrayOperationFor(guard.name, func(operation moduleArrayOperationFact) {
			if operation.Kind != moduleArrayWholeAssignment {
				return
			}
			guardWrites[guard.name] = append(guardWrites[guard.name], struct {
				line int
				rhs  string
			}{line: operation.Line, rhs: operation.RHS})
		})
	}

	result := map[string]bool{}
	for _, guard := range guards {
		writes := guardWrites[guard.name]
		if len(writes) != 1 || writes[0].line != lastExecutable || !strings.EqualFold(writes[0].rhs, "true") {
			continue
		}
		setAt := writes[0].line
		for index := guard.checkAt + 1; index < setAt; index++ {
			facts.forEachArrayOperationAt(index, func(operation moduleArrayOperationFact) {
				if operation.Kind != moduleArrayDirectRedim || operation.Preserve {
					return
				}
				name := operation.Name
				declaration, declared := moduleDecls[name]
				if !declared || !declaration.Array || declaration.Parameter {
					return
				}
				if moduleArrayOperationHasOtherWrite(facts, name, index) {
					return
				}
				if !arrayModuleSetupReDimIsReliable(file, proc, guard.checkAt, index, setAt, guard.name, name, constants, ctx, moduleDecls) {
					return
				}
				result[name] = true
			})
		}
	}
	if len(result) == 0 {
		return nil
	}
	return result
}

// inferArrayModuleReadyGuardStates recognizes the module-level lifecycle
// invariant used by consumers such as CSV readers:
//
//	moduleArray = Split(...)
//	ready = True
//
// and later:
//
//	If Not ready Then Exit Function
//	use moduleArray(...)
//
// The existing module allocation summary is caller-oriented and therefore
// cannot establish the state at an independently callable public procedure.
// This summary is deliberately source-owned and fail-closed so an arbitrary
// Boolean assignment never becomes an array allocation proof.
func inferArrayModuleReadyGuardStates(files []parsedFile, ctx analysisContext) arrayModuleReadyGuardStates {
	states := arrayModuleReadyGuardStates{}
	for _, file := range files {
		facts := file.moduleAnalysisFacts()
		if facts == nil {
			continue
		}
		moduleDecls := file.moduleDecls()
		for guardName, guardDeclaration := range moduleDecls {
			guardName = strings.ToLower(cleanIdentifier(guardName))
			if guardName == "" || guardDeclaration.Array || guardDeclaration.Parameter || !strings.EqualFold(strings.TrimSpace(guardDeclaration.Type), "Boolean") || !arrayModuleReadyGuardSourceOwned(file, guardDeclaration) {
				continue
			}

			var writes []moduleArrayOperationFact
			valid := true
			trueWrites := make([]moduleArrayOperationFact, 0, 1)
			facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
				if !valid {
					return
				}
				owner, owned := arrayModuleProcedureAtLine(file, operation.Line+1)
				if !owned {
					valid = false
					return
				}
				scope := newDeclarationScope(file, owner)
				scope.module = moduleDecls
				if scope.shadowsModule(guardName) {
					return
				}
				if operation.Kind != moduleArrayWholeAssignment {
					valid = false
					return
				}
				rhs := strings.ToLower(strings.TrimSpace(operation.RHS))
				if rhs != "true" && rhs != "false" {
					valid = false
					return
				}
				writes = append(writes, operation)
				if rhs == "true" {
					trueWrites = append(trueWrites, operation)
				}
			})
			if !valid || len(trueWrites) != 1 {
				continue
			}
			writer, ok := arrayModuleProcedureAtLine(file, trueWrites[0].Line+1)
			if !ok {
				continue
			}
			allocated := arrayModuleReadyGuardAllocationProof(file, writer, trueWrites[0].Line+1, moduleDecls, ctx)
			safeAllocated := map[string]bool{}
			for name := range allocated {
				if arrayModuleReadyGuardLifecycleSafe(file, guardName, map[string]bool{name: true}, facts, moduleDecls, ctx) {
					safeAllocated[name] = true
				}
			}
			if len(safeAllocated) == 0 {
				continue
			}
			if states[file.Path] == nil {
				states[file.Path] = map[string]map[string]bool{}
			}
			states[file.Path][guardName] = safeAllocated
		}
	}
	return states
}

func arrayModuleReadyGuardSourceOwned(file parsedFile, declaration sourceDeclaration) bool {
	if declaration.Line < 1 || declaration.Line > len(file.Lines) {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(file.Lines[declaration.Line-1])))
	return !strings.HasPrefix(lower, "public ") && !strings.HasPrefix(lower, "global ")
}

func arrayModuleProcedureAtLine(file parsedFile, line int) (sourceProcedure, bool) {
	if line < 1 {
		return sourceProcedure{}, false
	}
	procedures := file.procedureView()
	for index := 0; index < procedures.Len(); index++ {
		procedure := procedures.valueAt(index)
		if line >= procedure.StartLine && line <= procedure.EndLine {
			return procedure, true
		}
	}
	return sourceProcedure{}, false
}

func arrayModuleReadyGuardAllocationProof(file parsedFile, proc sourceProcedure, readyLine int, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	if proc.Graph == nil || readyLine < 1 || readyLine > len(file.Lines) {
		return nil
	}
	variables := arrayVariables(file, proc, moduleDecls)
	candidates := map[string]bool{}
	for name, declaration := range moduleDecls {
		name = strings.ToLower(cleanIdentifier(name))
		if name != "" && declaration.Array && !declaration.Parameter && arrayModuleReadyGuardSourceOwned(file, declaration) {
			if variable, known := variables[name]; known && variable.isArray {
				candidates[name] = true
			}
		}
	}
	if len(candidates) == 0 {
		return nil
	}

	graph := arrayVBA227Graph(proc, ctx)
	initial := arrayInitialState(variables)
	seenReady := false
	failed := map[string]bool{}
	visit := func(text string, line int, in arrayFlowState) arrayFlowState {
		if line == readyLine {
			seenReady = true
			for name := range candidates {
				value, known := in[name]
				if !known || value.kind != arrayAllocated || !value.knownArray {
					failed[name] = true
				}
			}
		}
		out, _ := (Analyzer{}).arrayVBA227Transfer(file, proc, ctx, variables, in, text, line, nil, nil, nil)
		for _, call := range arrayCallsAtLine(proc.Calls, line) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
		}
		return out
	}
	edgeState := func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
		out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
		return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
	}
	walkArrayCFGWithSourceLinesReliableStats(&graph, file.Lines, initial, visit, edgeState, arrayAllocationTransferIsReliable, nil)
	if !seenReady {
		return nil
	}
	result := map[string]bool{}
	for name := range candidates {
		if !failed[name] {
			result[name] = true
		}
	}
	return result
}

func arrayModuleReadyGuardLifecycleSafe(file parsedFile, guardName string, arrays map[string]bool, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration, ctx analysisContext) bool {
	if facts == nil {
		return false
	}
	for name := range arrays {
		name = strings.ToLower(cleanIdentifier(name))
		safe := true
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
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
			if scope.shadowsModule(name) {
				return
			}
			line := operation.Line + 1
			if line < 1 || line > len(file.Lines) {
				safe = false
				return
			}
			switch operation.Kind {
			case moduleArrayWholeAssignment:
				lhs, rhs, indexed, assigned := arrayAssignment(arrayLogicalCodeLine(file.Lines, line))
				if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), name) {
					safe = false
					return
				}
				value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx)
				if !known || value.kind != arrayAllocated || !value.knownArray {
					safe = false
				}
			case moduleArrayDirectRedim:
				if arraySummaryStatementAlwaysFails(arrayLogicalCodeLine(file.Lines, line), arrayOptionBase(file), arrayIntegerConstants(file, owner, nil, nil)) {
					safe = false
				}
			case moduleArrayErase:
				if !arrayModuleReadyGuardFalseWriteDominates(file, owner, guardName, line, facts, moduleDecls) {
					safe = false
				}
			default:
				safe = false
			}
		})
		if !safe {
			return false
		}
	}
	return true
}

func arrayModuleReadyGuardFalseWriteDominates(file parsedFile, proc sourceProcedure, guardName string, eraseLine int, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration) bool {
	if proc.Graph == nil || facts == nil || eraseLine <= proc.StartLine {
		return false
	}
	falseLines := make([]int, 0, 1)
	facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
		if operation.Kind != moduleArrayWholeAssignment || !strings.EqualFold(strings.TrimSpace(operation.RHS), "false") {
			return
		}
		line := operation.Line + 1
		owner, ok := arrayModuleProcedureAtLine(file, line)
		if !ok || owner.StartByte != proc.StartByte || owner.StartLine != proc.StartLine || owner.EndLine != proc.EndLine {
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if scope.shadowsModule(guardName) {
			return
		}
		if line < eraseLine {
			falseLines = append(falseLines, line)
		}
	})
	if len(falseLines) == 0 {
		return false
	}
	eraseStatement, eraseOK := arrayModuleStatementAtLine(proc, eraseLine)
	if !eraseOK {
		return false
	}
	normalGraph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	eraseBlock, eraseBlockOK := normalGraph.BlockForStatement(eraseStatement.ID)
	if !eraseBlockOK || !normalGraph.IsReachable(eraseBlock.ID) {
		return false
	}
	for _, falseLine := range falseLines {
		if arrayModuleReadyGuardHasWriteBetween(file, proc, guardName, falseLine, eraseLine, facts, moduleDecls) {
			continue
		}
		falseStatement, falseOK := arrayModuleStatementAtLine(proc, falseLine)
		if !falseOK {
			continue
		}
		falseBlock, falseBlockOK := normalGraph.BlockForStatement(falseStatement.ID)
		if !falseBlockOK {
			continue
		}
		for _, dominator := range normalGraph.DominatorsOf(eraseBlock.ID) {
			if dominator == falseBlock.ID {
				return true
			}
		}
	}
	return false
}

func arrayModuleReadyGuardHasWriteBetween(file parsedFile, proc sourceProcedure, guardName string, startLine, endLine int, facts *moduleAnalysisFacts, moduleDecls map[string]sourceDeclaration) bool {
	written := false
	facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
		if written || operation.Kind != moduleArrayWholeAssignment {
			return
		}
		line := operation.Line + 1
		if line <= startLine || line >= endLine {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, line)
		if !ok || owner.StartByte != proc.StartByte || owner.StartLine != proc.StartLine || owner.EndLine != proc.EndLine {
			return
		}
		scope := newDeclarationScope(file, owner)
		scope.module = moduleDecls
		if !scope.shadowsModule(guardName) {
			written = true
		}
	})
	return written
}

func arrayModuleStatementAtLine(proc sourceProcedure, line int) (procedureir.Statement, bool) {
	var found procedureir.Statement
	matched := false
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine != line {
			continue
		}
		if matched {
			return procedureir.Statement{}, false
		}
		found = statement
		matched = true
	}
	return found, matched
}

func arrayModuleSetupReDimIsReliable(file parsedFile, proc sourceProcedure, guardLine, redimLine, setLine int, guardName, name string, constants map[string]int, ctx analysisContext, moduleDecls map[string]sourceDeclaration) bool {
	if redimLine <= guardLine || redimLine >= setLine || redimLine < 0 || setLine >= len(file.Lines) {
		return false
	}
	sourceLine := redimLine + 1
	if sourceLine < 1 || sourceLine > len(file.Lines) || arraySummaryStatementAlwaysFails(normalizedCodeLine(file.Lines[redimLine]), arrayOptionBase(file), constants) {
		return false
	}
	variables := arrayVariables(file, proc, moduleDecls)
	// A call between the allocation and the ready flag can erase or replace the
	// module array without leaving a direct operation fact in this procedure.
	// A resolved private helper is admitted only when its direct and ByRef array
	// effects are known not to invalidate this module's arrays; public,
	// unresolved, and otherwise unmodelled calls remain conservative.
	for call := range proc.Calls.All() {
		line := call.Range.StartLine
		if line >= sourceLine && line < setLine+1 {
			if arrayCallIsIndexedArrayAccess(proc, call, variables) {
				continue
			}
			if call.IsRaiseEvent {
				return false
			}
			if call.Resolution.Status == procedureir.ResolutionBuiltinLike {
				continue
			}
			_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			if !resolved || arrayPrivateCallMayInvalidateModuleArray(file, proc, target, call, moduleDecls, ctx) {
				return false
			}
		}
	}

	if proc.Graph == nil {
		// Compatibility projections may not carry CFGs. Accept only a straight-
		// line source interval in that case; any visible control construct makes
		// the allocation conditional or permits an unmodelled bypass.
		for line := guardLine + 1; line < setLine; line++ {
			if line == redimLine {
				continue
			}
			if arrayModuleSetupLineHasControlFlow(file.Lines[line]) {
				return false
			}
		}
		return true
	}

	findStatement := func(line int, match func(procedureir.Statement) bool) (procedureir.Statement, bool) {
		var found procedureir.Statement
		matched := false
		for statement := range proc.Statements.All() {
			if statement.Range.StartLine != line || !match(statement) {
				continue
			}
			if matched {
				return procedureir.Statement{}, false
			}
			found = statement
			matched = true
		}
		return found, matched
	}
	guardStatement, guardOK := findStatement(guardLine+1, func(statement procedureir.Statement) bool {
		return len(arraySetupGuardRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))) == 2
	})
	redimStatement, redimOK := findStatement(sourceLine, func(statement procedureir.Statement) bool {
		match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(statement.Text)))
		if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
			return false
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), cleanIdentifier(name)) {
				return true
			}
		}
		return false
	})
	readyStatement, readyOK := findStatement(setLine+1, func(statement procedureir.Statement) bool {
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(statement.Text))
		return ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), cleanIdentifier(guardName)) && strings.EqualFold(strings.TrimSpace(rhs), "true")
	})
	if !guardOK || !redimOK || !readyOK {
		return false
	}
	guardBlock, guardBlockOK := proc.Graph.BlockForStatement(guardStatement.ID)
	redimBlock, redimBlockOK := proc.Graph.BlockForStatement(redimStatement.ID)
	readyBlock, readyBlockOK := proc.Graph.BlockForStatement(readyStatement.ID)
	if !guardBlockOK || !redimBlockOK || !readyBlockOK {
		return false
	}
	normalGraph := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	if !normalGraph.IsReachable(redimBlock.ID) || !normalGraph.IsReachable(readyBlock.ID) {
		return false
	}
	dominatesReady := false
	for _, dominator := range normalGraph.DominatorsOf(readyBlock.ID) {
		if dominator == redimBlock.ID {
			dominatesReady = true
			break
		}
	}
	if !dominatesReady || !arrayFalseBranchRequiresBlock(*proc.Graph, guardBlock.ID, redimBlock.ID) {
		return false
	}
	return arrayModuleSetupReachesNormalExit(normalGraph, readyBlock.ID)
}

// arrayCallIsIndexedArrayAccess filters the procedure IR's expression-call
// projection from actual procedure invocations. The VBA tree-sitter grammar
// represents an indexed array expression such as mPow2(index) as a CallSite;
// an unresolved array expression must not make an otherwise straight-line
// setup helper look like it contains an unknown side effect.
func arrayCallIsIndexedArrayAccess(proc sourceProcedure, call procedureir.CallSite, variables map[string]arrayVariable) bool {
	name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
	if name == "" || strings.Contains(name, ".") || call.StatementID <= 0 {
		return false
	}
	for statement := range proc.Statements.All() {
		if statement.ID != call.StatementID {
			continue
		}
		for _, use := range arrayIndexedUses(statement.Text, variables) {
			if strings.EqualFold(cleanIdentifier(use.name), name) {
				return true
			}
		}
		return false
	}
	return false
}

func arrayModuleSetupLineHasControlFlow(line string) bool {
	code := strings.TrimSpace(normalizedCodeLine(line))
	if code == "" {
		return false
	}
	lower := strings.ToLower(code)
	for _, prefix := range []string{
		"if ", "elseif ", "else", "end if", "for ", "for each ", "next", "do", "loop", "while ", "wend",
		"select ", "case ", "goto ", "on error ", "with ", "end with", "exit ",
	} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return strings.HasSuffix(code, ":")
}

func arrayModuleSetupReachesNormalExit(graph vbacfg.CFGView, from vbacfg.BlockID) bool {
	if from == graph.NormalExit() {
		return true
	}
	seen := map[vbacfg.BlockID]bool{from: true}
	queue := []vbacfg.BlockID{from}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		found := false
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if edge.To == graph.NormalExit() {
				found = true
				return false
			}
			if !seen[edge.To] {
				seen[edge.To] = true
				queue = append(queue, edge.To)
			}
			return true
		})
		if found {
			return true
		}
	}
	return false
}

func moduleArrayOperationHasOtherWrite(facts *moduleAnalysisFacts, name string, setupLine int) bool {
	name = strings.ToLower(cleanIdentifier(name))
	otherWrite := false
	facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
		if otherWrite {
			return
		}
		if operation.Kind == moduleArrayDirectRedim && operation.Line == setupLine {
			return
		}
		otherWrite = true
	})
	return otherWrite
}

func arrayProcedureLineHasInlineConditional(file parsedFile, line int) bool {
	if line < 1 || line > len(file.Lines) {
		return false
	}
	text := strings.ToLower(strings.TrimSpace(normalizedCodeLine(file.Lines[line-1])))
	return strings.HasPrefix(text, "if ") && strings.Contains(text, " then ")
}

func arrayProcedureLineInlineConditionIsFalse(file parsedFile, line int) bool {
	if !arrayProcedureLineHasInlineConditional(file, line) {
		return false
	}
	condition, body, ok := arrayIfThenParts(normalizedCodeLine(file.Lines[line-1]))
	if !ok || strings.TrimSpace(body) == "" {
		return false
	}
	condition = strings.TrimSpace(condition)
	lowerCondition := strings.ToLower(condition)
	if strings.HasPrefix(lowerCondition, "if ") {
		condition = strings.TrimSpace(condition[len("if "):])
	}
	value, known := arraySourceOrderConstantBoolean(condition, nil)
	return known && !value
}

func arrayInlineConditionalCallIsReachable(file parsedFile, call procedureir.CallSite, conditionValue, hasElse bool) bool {
	if conditionValue && !hasElse {
		return true
	}
	line := call.Range.StartLine
	if line < 1 || line > len(file.Lines) || call.Range.StartColumn <= 0 {
		return conditionValue || hasElse
	}
	raw := gui.StripComment(file.Lines[line-1])
	trimmed := strings.TrimSpace(raw)
	leading := len(raw) - len(strings.TrimLeft(raw, " \t"))
	_, body, ok := arrayIfThenParts(trimmed)
	if !ok || strings.TrimSpace(body) == "" {
		return true
	}
	prefixLength := 0
	switch {
	case strings.HasPrefix(strings.ToLower(trimmed), "if "):
		prefixLength = len("if ")
	case strings.HasPrefix(strings.ToLower(trimmed), "elseif "):
		prefixLength = len("elseif ")
	default:
		return true
	}
	rest := strings.TrimSpace(trimmed[prefixLength:])
	thenIndex := arrayTopLevelKeywordIndex(rest, "then")
	if thenIndex < 0 {
		return true
	}
	bodyStart := leading + prefixLength + (len(trimmed[prefixLength:]) - len(rest)) + thenIndex + len("then")
	elseIndex := arrayTopLevelKeywordIndex(body, "else")
	if elseIndex < 0 {
		return conditionValue
	}
	elseStart := bodyStart + elseIndex
	callColumn := call.Range.StartColumn - 1
	if callColumn < bodyStart {
		return true
	}
	inThen := callColumn < elseStart
	return inThen == conditionValue
}

func arrayProcedureNormalExitDominators(proc sourceProcedure) map[vbacfg.BlockID]bool {
	if proc.Graph == nil {
		return nil
	}
	dominators := proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[proc.Graph.NormalExit]
	result := make(map[vbacfg.BlockID]bool, len(dominators))
	for _, id := range dominators {
		result[id] = true
	}
	return result
}

func arrayProcedureBlockDominatesNormalExit(proc sourceProcedure, statementID int, dominators map[vbacfg.BlockID]bool) bool {
	if proc.Graph == nil || statementID <= 0 || len(dominators) == 0 {
		return false
	}
	block, ok := proc.Graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	return dominators[block.ID]
}

func arrayProcedureHasIdempotentSetupGuard(file parsedFile, proc sourceProcedure, candidateLine int, moduleDecls map[string]sourceDeclaration) bool {
	if candidateLine <= proc.StartLine {
		return false
	}
	guard := ""
	for line := proc.StartLine; line < candidateLine && line <= len(file.Lines); line++ {
		match := arraySetupGuardRe.FindStringSubmatch(normalizedCodeLine(file.Lines[line-1]))
		if len(match) != 2 {
			continue
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		declaration, ok := moduleDecls[name]
		if !ok || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			continue
		}
		guard = name
	}
	if guard == "" {
		return false
	}
	for line := candidateLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		lhs, rhs, indexed, ok := arrayAssignment(normalizedCodeLine(file.Lines[line-1]))
		if ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), guard) && strings.EqualFold(strings.TrimSpace(rhs), "true") {
			return true
		}
	}
	return false
}

func arrayModuleAllocationSummariesEqual(left, right arrayModuleAllocationSummaries) bool {
	if len(left) != len(right) {
		return false
	}
	for procedure, arrays := range left {
		other, ok := right[procedure]
		if !ok || len(arrays) != len(other) {
			return false
		}
		for name := range arrays {
			if !other[name] {
				return false
			}
		}
	}
	return true
}

func arrayModuleInitializationStates(files []parsedFile, summaries arrayModuleAllocationSummaries) map[string]map[string]bool {
	states := map[string]map[string]bool{}
	for _, file := range files {
		moduleKind := strings.ToLower(strings.TrimSpace(file.ModuleKind))
		if moduleKind != "form" && moduleKind != "class" {
			continue
		}
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		initializer := arrayModuleInitializerName(moduleKind)
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !strings.EqualFold(strings.TrimSpace(proc.Name), initializer) {
				continue
			}
			for name := range summaries[arrayProcedureKey(proc)] {
				if declaration, ok := moduleDecls[name]; ok && declaration.Array && !declaration.Parameter {
					if states[file.Path] == nil {
						states[file.Path] = map[string]bool{}
					}
					states[file.Path][name] = true
				}
			}
		}
	}
	return states
}

func applyArrayModuleInitializationState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, initializationStates map[string]map[string]bool) arrayFlowState {
	if len(initializationStates[file.Path]) == 0 {
		return state
	}
	moduleKind := strings.ToLower(strings.TrimSpace(file.ModuleKind))
	if moduleKind != "form" && moduleKind != "class" {
		return state
	}
	initializer := arrayModuleInitializerName(moduleKind)
	if strings.EqualFold(strings.TrimSpace(proc.Name), initializer) {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range initializationStates[file.Path] {
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, ok := moduleDecls[name]
		variable, known := variables[name]
		if !ok || !known || !declaration.Array || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayModuleInitializerName(moduleKind string) string {
	if moduleKind == "form" {
		return "userform" + "_initialize"
	}
	return "class" + "_initialize"
}

// inferArrayModuleEntryStates propagates a module-array allocation from a
// known caller to a project-local helper. Private procedures are analyzed as
// standalone procedures, so their initial state cannot otherwise reflect an
// assignment performed by the caller. A fact is retained only when every
// resolved call from the same module reaches the helper with that array
// allocated. The fixed point also covers chains of private helpers.
func inferArrayModuleEntryStates(a Analyzer, files []parsedFile, ctx analysisContext) arrayModuleEntryStates {
	if len(ctx.arrayPrivateTargets) == 0 {
		return arrayModuleEntryStates{}
	}

	type procedureInfo struct {
		file        parsedFile
		proc        sourceProcedure
		key         string
		moduleDecls map[string]sourceDeclaration
		variables   map[string]arrayVariable
	}
	procedures := make([]procedureInfo, 0)
	moduleArrays := map[string]map[string]bool{}
	moduleFiles := map[string]string{}
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			if !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			key := arrayParticipantLookupKey(proc, ctx.arrayParticipantKeys)
			procedures = append(procedures, procedureInfo{
				file: file, proc: proc, moduleDecls: moduleDecls,
				variables: arrayVariables(file, proc, moduleDecls),
				key:       key,
			})
			moduleArrays[key] = arrayModuleNamesForProcedure(file, proc, moduleDecls)
			moduleFiles[key] = file.Path
		}
	}
	if len(procedures) == 0 {
		return arrayModuleEntryStates{}
	}

	initializationStates := arrayModuleInitializationStates(files, ctx.arrayModuleAllocations)
	sort.SliceStable(procedures, func(i, j int) bool {
		return arrayProcedureLess(procedures[i].proc, procedures[j].proc)
	})
	indexByKey := make(map[string]int, len(procedures))
	dependents := make(map[string][]int)
	for index, procedure := range procedures {
		key := procedure.key
		indexByKey[key] = index
		for call := range procedure.proc.Calls.All() {
			_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			participantTargetKey := arrayParticipantLookupKey(target, ctx.arrayParticipantKeys)
			if ok && moduleFiles[participantTargetKey] == procedure.file.Path {
				dependents[participantTargetKey] = append(dependents[participantTargetKey], index)
			}
		}
	}
	for key := range dependents {
		sort.Ints(dependents[key])
	}

	evaluate := func(procedure procedureInfo, entries arrayModuleEntryStates) map[string]map[string]bool {
		variables := procedure.variables
		initial := arrayInitialState(variables)
		initial = applyArrayModuleInitializationState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, initializationStates)
		initial = applyArrayModuleReadyGuardState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, ctx.arrayModuleReadyGuards)
		initial = applyArrayModuleEntryState(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, entries, ctx.arrayParticipantKeys)
		initial = applyArrayInternalStorageConfiguration(initial, procedure.file, procedure.proc, variables, procedure.moduleDecls, ctx.arrayModuleConfigurations[procedure.file.Path])
		candidates := map[string]map[string]bool{}
		recordCall := func(call procedureir.CallSite, state arrayFlowState) {
			_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			key := arrayParticipantLookupKey(target, ctx.arrayParticipantKeys)
			if !ok || moduleFiles[key] != procedure.file.Path {
				return
			}
			names := moduleArrays[key]
			if len(names) == 0 {
				return
			}
			candidate := candidates[key]
			if candidate == nil {
				candidate = cloneArrayNameSet(names)
				candidates[key] = candidate
			}
			for name := range names {
				value, known := state[name]
				if !known || value.kind != arrayAllocated || !value.knownArray {
					candidate[name] = false
				}
			}
		}
		visit := func(text string, line int, in arrayFlowState) arrayFlowState {
			for _, call := range arrayCallsAtLine(procedure.proc.Calls, line) {
				recordCall(call, in)
			}
			out, _ := a.arrayTransfer(procedure.file, procedure.proc, ctx, variables, in, text, line, nil, nil)
			for _, call := range arrayCallsAtLine(procedure.proc.Calls, line) {
				out = applyArrayModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
				out = applyArrayUnknownModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
			}
			return out
		}
		if procedure.proc.Graph == nil {
			state := initial
			for line := procedure.proc.StartLine; line <= procedure.proc.EndLine && line <= len(procedure.file.Lines); line++ {
				state = visit(normalizedCodeLine(procedure.file.Lines[line-1]), line, state)
			}
			return candidates
		}
		graph := arrayVBA227Graph(procedure.proc, ctx)
		if ctx.arrayStats != nil {
			ctx.arrayStats.addCFGWalk()
		}
		walkArrayCFGWithEdgesStats(&graph, procedure.file.Lines, initial, visit, func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
			out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[procedure.file.Path], variables, procedure.file, procedure.proc, procedure.moduleDecls)
		}, ctx.arrayStats)
		return candidates
	}

	contributions := make(map[string]map[string]map[string]bool, len(procedures))
	entries := arrayModuleEntryStates{}
	queue := make([]int, len(procedures))
	queued := make([]bool, len(procedures))
	for index := range procedures {
		queue[index] = index
		queued[index] = true
	}
	for head := 0; head < len(queue); head++ {
		index := queue[head]
		queued[index] = false
		key := procedures[index].key
		if head >= len(procedures) && ctx.arrayStats != nil {
			ctx.arrayStats.addRevisit()
		}
		contribution := evaluate(procedures[index], entries)
		if arrayModuleEntryContributionsEqual(contributions[key], contribution) {
			continue
		}
		contributions[key] = contribution
		next := arrayModuleEntryStates{}
		for _, caller := range procedures {
			for target, names := range contributions[caller.key] {
				if next[target] == nil {
					next[target] = cloneArrayNameSet(names)
					continue
				}
				for name := range next[target] {
					if !names[name] {
						delete(next[target], name)
					}
				}
			}
		}
		changedTargetSet := make(map[string]bool, len(entries)+len(next))
		for target := range entries {
			changedTargetSet[target] = true
		}
		for target := range next {
			changedTargetSet[target] = true
		}
		changedTargets := make([]string, 0, len(changedTargetSet))
		for target := range changedTargetSet {
			if !arrayModuleEntryTargetEqual(entries, next, target) {
				changedTargets = append(changedTargets, target)
			}
		}
		if len(changedTargets) == 0 {
			continue
		}
		entries = next
		sortArrayProcedureKeys(changedTargets, indexByKey)
		for _, target := range changedTargets {
			for _, dependent := range dependents[target] {
				if !queued[dependent] {
					queued[dependent] = true
					queue = append(queue, dependent)
				}
			}
			if index, ok := indexByKey[target]; ok && !queued[index] {
				queued[index] = true
				queue = append(queue, index)
			}
		}
	}
	return entries
}

func arrayModuleEntryContributionsEqual(left, right map[string]map[string]bool) bool {
	if len(left) != len(right) {
		return false
	}
	for key, names := range left {
		other, ok := right[key]
		if !ok || len(names) != len(other) {
			return false
		}
		for name, allocated := range names {
			if other[name] != allocated {
				return false
			}
		}
	}
	return true
}

func arrayModuleEntryTargetEqual(left, right arrayModuleEntryStates, target string) bool {
	leftNames, leftOK := left[target]
	rightNames, rightOK := right[target]
	if !leftOK || !rightOK {
		return !leftOK && !rightOK
	}
	if len(leftNames) != len(rightNames) {
		return false
	}
	for name, allocated := range leftNames {
		if rightNames[name] != allocated {
			return false
		}
	}
	return true
}

func arrayModuleNamesForProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]bool {
	moduleArrays := map[string]bool{}
	for name, declaration := range moduleDecls {
		if declaration.Array && !declaration.Parameter {
			moduleArrays[strings.ToLower(name)] = true
		}
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	for name := range moduleArrays {
		if declarations.shadowsModule(name) {
			delete(moduleArrays, name)
		}
	}
	return moduleArrays
}

func applyArrayModuleEntryState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, entries arrayModuleEntryStates, participantKeys ...map[string]string) arrayFlowState {
	var keyIndex map[string]string
	if len(participantKeys) > 0 {
		keyIndex = participantKeys[0]
	}
	allocated := entries[arrayParticipantLookupKey(proc, keyIndex)]
	if len(allocated) == 0 && len(keyIndex) > 0 {
		// Focused compatibility callers may still provide the legacy base key.
		allocated = entries[arrayProcedureKey(proc)]
	}
	if len(allocated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name, isAllocated := range allocated {
		if !isAllocated {
			continue
		}
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || declaration.Parameter || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func applyArrayModuleReadyGuardState(state arrayFlowState, file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, guards arrayModuleReadyGuardStates) arrayFlowState {
	byGuard := guards[file.Path]
	if len(byGuard) == 0 {
		return state
	}
	guardName, ok := arrayModuleReadyGuardAtEntry(file, proc, moduleDecls)
	if !ok {
		return state
	}
	allocated := byGuard[guardName]
	if len(allocated) == 0 {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	updated := cloneArrayState(state)
	for name := range allocated {
		if declarations.shadowsModule(name) {
			continue
		}
		declaration, declared := moduleDecls[name]
		variable, known := variables[name]
		if !declared || !declaration.Array || declaration.Parameter || !known || !variable.isArray {
			continue
		}
		value := updated[name]
		value.kind = arrayAllocated
		value.knownArray = true
		updated[name] = value
	}
	return updated
}

func arrayModuleReadyGuardAtEntry(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) (string, bool) {
	if proc.StartLine < 1 || proc.EndLine <= proc.StartLine || proc.StartLine > len(file.Lines) {
		return "", false
	}
	end := min(len(file.Lines), proc.EndLine)
	for line := proc.StartLine + 1; line < end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" || strings.HasPrefix(text, "'") {
			continue
		}
		if declRe.MatchString(text) || strings.HasPrefix(strings.ToLower(text), "on error ") || strings.HasSuffix(text, ":") {
			continue
		}
		match := arrayModuleReadyGuardRe.FindStringSubmatch(text)
		if len(match) != 2 {
			return "", false
		}
		name := strings.ToLower(cleanIdentifier(match[1]))
		declaration, declared := moduleDecls[name]
		if !declared || declaration.Array || declaration.Parameter || !strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			return "", false
		}
		return name, true
	}
	return "", false
}

func arrayOptionPrivateModule(lines []string) bool {
	for _, line := range lines {
		if strings.EqualFold(strings.TrimSpace(normalizedCodeLine(line)), "option private module") {
			return true
		}
	}
	return false
}
