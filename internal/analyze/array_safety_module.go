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
		if !hasModuleDynamicArrayDeclaration(moduleDecls) {
			continue
		}
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if !arrayProcedureIsModuleEffectParticipant(ctx, proc) {
				continue
			}
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
	if arrays := arrayModuleStorageCallEstablishesGroup(file, target, ctx, ctx.arrayModuleStorageGuards[file.Path]); len(arrays) > 0 {
		updated = arrayModuleStorageAllocatedState(updated, arrays)
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
	if !arrayProcedureIsModuleEffectParticipant(ctx, target) {
		return nil
	}
	return arrayPrivateModuleArrayInvalidationsWithVisiting(file, target, moduleDecls, ctx, map[string]bool{})
}

func hasModuleDynamicArrayDeclaration(moduleDecls map[string]sourceDeclaration) bool {
	for _, declaration := range moduleDecls {
		if declaration.Array && !declaration.Fixed && !declaration.Parameter {
			return true
		}
	}
	return false
}

// arrayModuleStorageGuard describes a source-owned storage invariant. The
// named module arrays are resized with `0 To capacity - 1`, every erase of
// those arrays resets the capacity to zero, and the optional count scalars
// are changed only from zero by +/- one. A nonzero capacity (or count) is
// therefore enough to establish allocation for the group without trusting an
// arbitrary Boolean or numeric flag.
type arrayModuleStorageGuard struct {
	capacity string
	counts   map[string]bool
	arrays   map[string]bool
}

type arrayModuleStorageAssignment struct {
	rhs     string
	indexed bool
	line    int
	owner   sourceProcedure
	ownerOK bool
}

type arrayModuleStorageErase struct {
	line    int
	owner   sourceProcedure
	ownerOK bool
}

// arrayModuleStorageSourceFacts contains the source-level facts shared by all
// capacity candidates in one module. Keeping this index local to guard
// inference avoids reparsing every logical source line for every scalar.
type arrayModuleStorageSourceFacts struct {
	assignments map[string][]arrayModuleStorageAssignment
	erases      map[string][]arrayModuleStorageErase
}

func buildArrayModuleStorageSourceFacts(file parsedFile, assignmentNames, eraseNames map[string]bool) arrayModuleStorageSourceFacts {
	source := arrayModuleStorageSourceFacts{
		assignments: make(map[string][]arrayModuleStorageAssignment),
		erases:      make(map[string][]arrayModuleStorageErase),
	}
	for line := 1; line <= len(file.Lines); line++ {
		text := arrayLogicalCodeLine(file.Lines, line)
		if strings.TrimSpace(text) == "" {
			continue
		}
		owner, ownerOK := arrayModuleProcedureAtLine(file, line)
		for _, part := range splitRangeValueSourceStatements(text) {
			part = strings.TrimSpace(part)
			if lhs, rhs, indexed, assigned := arrayAssignment(part); assigned {
				name := strings.ToLower(cleanIdentifier(lhs))
				if assignmentNames[name] {
					source.assignments[name] = append(source.assignments[name], arrayModuleStorageAssignment{
						rhs: rhs, indexed: indexed, line: line, owner: owner, ownerOK: ownerOK,
					})
				}
			}
			if match := arrayEraseRe.FindStringSubmatch(part); len(match) == 2 {
				for _, target := range splitArgs(match[1]) {
					name := strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))
					if !eraseNames[name] {
						continue
					}
					source.erases[name] = append(source.erases[name], arrayModuleStorageErase{line: line, owner: owner, ownerOK: ownerOK})
				}
			}
		}
	}
	return source
}

func inferArrayModuleStorageGuards(files []parsedFile) map[string][]arrayModuleStorageGuard {
	guardsByFile := make(map[string][]arrayModuleStorageGuard)
	for index := range files {
		files[index].ensureModuleAnalysisFacts()
	}
	for _, file := range files {
		guards := arrayModuleStorageGuardsForFile(file, file.moduleDecls())
		if len(guards) > 0 {
			guardsByFile[file.Path] = guards
		}
	}
	return guardsByFile
}

func arrayModuleStorageGuardsForFile(file parsedFile, moduleDecls map[string]sourceDeclaration) []arrayModuleStorageGuard {
	facts := file.moduleAnalysisFacts()
	if facts == nil || len(moduleDecls) == 0 {
		return nil
	}
	arrays := make(map[string]bool)
	scalars := make(map[string]bool)
	for name, declaration := range moduleDecls {
		name = strings.ToLower(cleanIdentifier(name))
		if name == "" || declaration.Parameter {
			continue
		}
		if declaration.Array && !declaration.Fixed {
			arrays[name] = true
			continue
		}
		if !declaration.Array && !declaration.Object && arrayKnownScalarType(declaration.Type) {
			scalars[name] = true
		}
	}
	if len(arrays) == 0 || len(scalars) == 0 {
		return nil
	}

	// The capacity scalar must occur in the exact source-owned ReDim shape
	// before any of the more expensive lifecycle checks can succeed. Restrict
	// the candidate set up front; large modules commonly contain many scalar
	// declarations that are unrelated to their dynamic arrays.
	candidateCapacities := make(map[string]bool)
	for arrayName := range arrays {
		facts.forEachArrayOperationFor(arrayName, func(operation moduleArrayOperationFact) {
			if operation.Kind != moduleArrayDirectRedim || operation.Preserve {
				return
			}
			capacity, ok := arrayModuleStorageCapacityName(operation.Dimensions)
			if ok && scalars[capacity] {
				candidateCapacities[capacity] = true
			}
		})
	}
	if len(candidateCapacities) == 0 {
		return nil
	}
	source := buildArrayModuleStorageSourceFacts(file, scalars, arrays)

	guards := make([]arrayModuleStorageGuard, 0)
	for capacity := range candidateCapacities {
		storageArrays := make(map[string]bool)
		for name := range arrays {
			if arrayModuleStorageArrayMatches(file, facts, source, moduleDecls, name, capacity) {
				storageArrays[name] = true
			}
		}
		if len(storageArrays) == 0 || !arrayModuleStorageScalarWritesValid(file, facts, source, storageArrays, capacity) {
			continue
		}
		guard := arrayModuleStorageGuard{
			capacity: capacity,
			counts:   make(map[string]bool),
			arrays:   storageArrays,
		}
		for count := range scalars {
			if count == capacity || !arrayModuleStorageCountWritesValid(file, facts, source, storageArrays, count) {
				continue
			}
			guard.counts[count] = true
		}
		guards = append(guards, guard)
	}
	return guards
}

func arrayModuleStorageCapacityName(dimensions string) (string, bool) {
	compact := canonicalArrayBoundExpression(dimensions)
	if !strings.HasPrefix(compact, "0to") || !strings.HasSuffix(compact, "-1") {
		return "", false
	}
	name := strings.TrimSuffix(strings.TrimPrefix(compact, "0to"), "-1")
	if !isIdentifier(name) {
		return "", false
	}
	return strings.ToLower(cleanIdentifier(name)), true
}

func arrayModuleStorageArrayMatches(file parsedFile, facts *moduleAnalysisFacts, source arrayModuleStorageSourceFacts, moduleDecls map[string]sourceDeclaration, arrayName, capacity string) bool {
	declaration, declared := moduleDecls[arrayName]
	if !declared || !declaration.Array || declaration.Fixed || declaration.Parameter {
		return false
	}
	wantDimensions := canonicalArrayBoundExpression("0 To " + capacity + " - 1")
	redimCount := 0
	valid := true
	facts.forEachArrayOperationFor(arrayName, func(operation moduleArrayOperationFact) {
		if !valid {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
		if !ok || arrayProcedureHasErrorHandling(owner) {
			valid = false
			return
		}
		switch operation.Kind {
		case moduleArrayDirectRedim:
			if operation.Preserve || canonicalArrayBoundExpression(operation.Dimensions) != wantDimensions {
				valid = false
				return
			}
			redimCount++
		case moduleArrayErase:
			if !arrayModuleStorageResetWritesScalar(source, owner, capacity, arrayName) {
				valid = false
			}
		default:
			valid = false
		}
	})
	return valid && redimCount > 0
}

func arrayModuleStorageResetWritesScalar(source arrayModuleStorageSourceFacts, proc sourceProcedure, scalar, arrayName string) bool {
	zeroWritten := false
	erased := false
	for _, erase := range source.erases[strings.ToLower(cleanIdentifier(arrayName))] {
		if erase.ownerOK && arrayModuleStorageSameProcedure(erase.owner, proc) {
			erased = true
			break
		}
	}
	for _, assignment := range source.assignments[strings.ToLower(cleanIdentifier(scalar))] {
		if assignment.ownerOK && !assignment.indexed && arrayModuleStorageSameProcedure(assignment.owner, proc) {
			if value, ok := integerLiteral(assignment.rhs); ok && value == 0 {
				zeroWritten = true
				break
			}
		}
	}
	return zeroWritten && erased
}

func arrayModuleStorageScalarWritesValid(file parsedFile, facts *moduleAnalysisFacts, source arrayModuleStorageSourceFacts, arrays map[string]bool, scalar string) bool {
	seen := false
	for _, assignment := range source.assignments[strings.ToLower(cleanIdentifier(scalar))] {
		if assignment.indexed {
			continue
		}
		seen = true
		if !assignment.ownerOK {
			return false
		}
		if value, literal := integerLiteral(assignment.rhs); literal && value == 0 {
			if !arrayModuleStorageProcedureTouchesGroup(file, facts, assignment.owner, arrays) {
				return false
			}
			continue
		}
		if !arrayModuleStoragePositiveExpression(file, assignment.rhs) || !arrayModuleStorageHasRedimAfter(file, facts, assignment.owner, arrays, assignment.line) {
			return false
		}
	}
	return seen
}

func arrayModuleStorageCountWritesValid(file parsedFile, facts *moduleAnalysisFacts, source arrayModuleStorageSourceFacts, arrays map[string]bool, count string) bool {
	seenZero := false
	seenDelta := false
	for _, assignment := range source.assignments[strings.ToLower(cleanIdentifier(count))] {
		if assignment.indexed {
			continue
		}
		if !assignment.ownerOK {
			return false
		}
		if value, literal := integerLiteral(assignment.rhs); literal && value == 0 {
			if !arrayModuleStorageProcedureTouchesGroup(file, facts, assignment.owner, arrays) {
				return false
			}
			seenZero = true
			continue
		}
		compact := canonicalArrayBoundExpression(assignment.rhs)
		name := strings.ToLower(cleanIdentifier(count))
		if compact == name+"+1" || compact == name+"-1" {
			seenDelta = true
			continue
		}
		return false
	}
	return seenZero && seenDelta
}

func arrayModuleStorageSameProcedure(left, right sourceProcedure) bool {
	return left.StartByte == right.StartByte && left.StartLine == right.StartLine && left.EndLine == right.EndLine
}

// applyArrayModuleCoupledBoundsState carries a successful bounds query from
// one source-owned module array to another array whose complete lifecycle is
// identical. VBA code commonly keeps parallel key/value arrays in lockstep:
// both arrays are ReDim'ed and erased on the same source lines, and a
// successful UBound(keys) condition therefore proves that values is allocated
// too. Requiring identical operation locations and shapes keeps this
// refinement fail-closed for unrelated arrays.
func applyArrayModuleCoupledBoundsState(state arrayFlowState, file parsedFile, variables map[string]arrayVariable, moduleDecls map[string]sourceDeclaration, text string) arrayFlowState {
	facts := file.moduleAnalysisFacts()
	if facts == nil || len(state) == 0 || len(moduleDecls) == 0 || !arrayBoundCallRe.MatchString(text) {
		return state
	}
	boundNames := make(map[string]bool)
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(cleanIdentifier(bound[2]))
		if name == "" || !strings.EqualFold(bound[1], "ubound") && !strings.EqualFold(bound[1], "lbound") {
			continue
		}
		value, known := state[name]
		variable, declared := variables[name]
		declaration, moduleDeclared := moduleDecls[name]
		if !known || value.kind != arrayAllocated || !value.knownArray || !declared || !variable.isArray || !moduleDeclared || !declaration.Array || declaration.Fixed || declaration.Parameter || !arrayModuleReadyGuardSourceOwned(file, declaration) {
			continue
		}
		boundNames[name] = true
	}
	if len(boundNames) == 0 {
		return state
	}

	operations := make(map[string][]moduleArrayOperationFact)
	for name := range moduleDecls {
		name = strings.ToLower(cleanIdentifier(name))
		declaration := moduleDecls[name]
		variable, declared := variables[name]
		if name == "" || !declared || !declaration.Array || declaration.Fixed || declaration.Parameter || !variable.isArray || !arrayModuleReadyGuardSourceOwned(file, declaration) {
			continue
		}
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			operations[name] = append(operations[name], operation)
		})
	}

	paired := make(map[string]bool)
	for boundName := range boundNames {
		boundOperations := operations[boundName]
		if len(boundOperations) == 0 || !arrayModuleCoupledBoundsLifecycleComplete(boundOperations) {
			continue
		}
		for candidateName, candidateOperations := range operations {
			if candidateName == boundName || !arrayModuleCoupledBoundsLifecycleEqual(boundOperations, candidateOperations) {
				continue
			}
			paired[candidateName] = true
		}
	}
	if len(paired) == 0 {
		return state
	}

	updated := cloneArrayState(state)
	for name := range paired {
		value, known := updated[name]
		if !known {
			continue
		}
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeUnallocated = false
		updated[name] = value
	}
	return updated
}

func arrayModuleCoupledBoundsLifecycleComplete(operations []moduleArrayOperationFact) bool {
	hasRedim := false
	hasErase := false
	for _, operation := range operations {
		switch operation.Kind {
		case moduleArrayDirectRedim:
			hasRedim = true
		case moduleArrayErase:
			hasErase = true
		default:
			return false
		}
	}
	return hasRedim && hasErase
}

func arrayModuleCoupledBoundsLifecycleEqual(left, right []moduleArrayOperationFact) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		lineDistance := left[index].Line - right[index].Line
		if lineDistance < 0 {
			lineDistance = -lineDistance
		}
		if lineDistance > 4 || left[index].Kind != right[index].Kind || left[index].Preserve != right[index].Preserve || arrayModuleCoupledBoundsDimensions(left[index].Dimensions, left[index].Name) != arrayModuleCoupledBoundsDimensions(right[index].Dimensions, right[index].Name) {
			return false
		}
	}
	return true
}

func arrayModuleCoupledBoundsDimensions(dimensions, arrayName string) string {
	compact := canonicalArrayBoundExpression(dimensions)
	name := strings.ToLower(cleanIdentifier(arrayName))
	if name == "" {
		return compact
	}
	return strings.ReplaceAll(compact, name, "<array>")
}

func arrayModuleStorageProcedureTouchesGroup(file parsedFile, facts *moduleAnalysisFacts, proc sourceProcedure, arrays map[string]bool) bool {
	touched := false
	for name := range arrays {
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if touched {
				return
			}
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
			if ok && owner.StartByte == proc.StartByte && owner.StartLine == proc.StartLine && owner.EndLine == proc.EndLine {
				touched = true
			}
		})
		if touched {
			break
		}
	}
	return touched
}

func arrayModuleStorageHasRedimAfter(file parsedFile, facts *moduleAnalysisFacts, proc sourceProcedure, arrays map[string]bool, line int) bool {
	for name := range arrays {
		found := false
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if found || operation.Kind != moduleArrayDirectRedim || operation.Line+1 <= line {
				return
			}
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
			found = ok && owner.StartByte == proc.StartByte && owner.StartLine == proc.StartLine && owner.EndLine == proc.EndLine
		})
		if found {
			return true
		}
	}
	return false
}

func arrayModuleStoragePositiveExpression(file parsedFile, text string) bool {
	if value, ok := integerLiteral(text); ok {
		return value > 0
	}
	name := strings.ToLower(cleanIdentifier(arrayCallName(text)))
	if name == "" {
		return false
	}
	procedures := file.procedureView()
	for index := 0; index < procedures.Len(); index++ {
		proc := procedures.valueAt(index)
		if strings.EqualFold(proc.Name, name) {
			return arrayModuleStoragePositiveFunction(file, proc)
		}
	}
	return false
}

func arrayModuleStoragePositiveFunction(file parsedFile, proc sourceProcedure) bool {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet || arrayProcedureHasErrorHandling(proc) {
		return false
	}
	positive := map[string]bool{}
	returnSeen := false
	start := max(1, proc.StartLine)
	end := min(len(file.Lines), proc.EndLine)
	for line := start; line <= end; line++ {
		for _, part := range splitRangeValueSourceStatements(arrayLogicalCodeLine(file.Lines, line)) {
			lhs, rhs, indexed, assigned := arrayAssignment(part)
			if !assigned || indexed {
				continue
			}
			lhs = strings.ToLower(cleanIdentifier(lhs))
			if lhs == strings.ToLower(cleanIdentifier(proc.Name)) {
				returnSeen = true
				if !arrayModuleStoragePositiveRHS(rhs, positive) {
					return false
				}
				continue
			}
			if !arrayModuleStorageTracksPositiveScalar(lhs, rhs, positive) {
				if positive[lhs] {
					positive[lhs] = false
				}
				continue
			}
			positive[lhs] = true
		}
	}
	return returnSeen
}

func arrayModuleStoragePositiveRHS(rhs string, positive map[string]bool) bool {
	compact := canonicalArrayBoundExpression(rhs)
	if value, ok := integerLiteral(compact); ok {
		return value > 0
	}
	name := strings.ToLower(cleanIdentifier(compact))
	if positive[name] && name == compact {
		return true
	}
	for _, operator := range []string{"*", "+"} {
		index := strings.Index(compact, operator)
		if index <= 0 || index+len(operator) >= len(compact) {
			continue
		}
		base := strings.ToLower(cleanIdentifier(compact[:index]))
		if value, ok := integerLiteral(compact[index+len(operator):]); ok {
			return positive[base] && value > 0
		}
	}
	return false
}

func arrayModuleStorageTracksPositiveScalar(lhs, rhs string, positive map[string]bool) bool {
	if value, ok := integerLiteral(rhs); ok {
		return value > 0
	}
	compact := canonicalArrayBoundExpression(rhs)
	name := strings.ToLower(cleanIdentifier(lhs))
	if positive[strings.ToLower(cleanIdentifier(compact))] {
		return true
	}
	for _, operator := range []string{"*", "+"} {
		if !strings.HasPrefix(compact, name+operator) {
			continue
		}
		if value, ok := integerLiteral(strings.TrimPrefix(compact, name+operator)); ok {
			return positive[name] && value > 0
		}
	}
	return false
}

func applyArrayModuleStorageGuardBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, file parsedFile, proc sourceProcedure, ctx analysisContext, guards []arrayModuleStorageGuard) arrayFlowState {
	if statement == nil || (edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse) || statement.Condition == nil {
		return state
	}
	return applyArrayModuleStorageCondition(state, statement.Condition.Text, edge.Kind, statement.Range.StartLine, file, proc, ctx, guards)
}

func applyArrayModuleStorageCondition(state arrayFlowState, condition string, branch vbacfg.EdgeKind, guardLine int, file parsedFile, proc sourceProcedure, ctx analysisContext, guards []arrayModuleStorageGuard) arrayFlowState {
	if branch != vbacfg.EdgeBranchTrue && branch != vbacfg.EdgeBranchFalse {
		return state
	}
	state = applyArrayModuleAllocationCondition(state, condition, branch, guardLine, file, proc, ctx)
	if len(guards) == 0 {
		return state
	}
	lhs, operator, literal, ok := arrayCountComparison(condition)
	if !ok {
		callName, successBranch, callOK := arrayModuleStorageSuccessfulCallBranch(condition)
		if !callOK || branch != successBranch {
			return state
		}
		for _, guard := range guards {
			if arrayModuleStorageCallAllocatesGroup(file, proc, ctx, callName, guard, guardLine) {
				return arrayModuleStorageAllocatedState(state, guard.arrays)
			}
		}
		return state
	}
	lhs = strings.ToLower(cleanIdentifier(lhs))
	for _, guard := range guards {
		if lhs == guard.capacity && operator == "=" && literal == "0" && branch == vbacfg.EdgeBranchFalse {
			return arrayModuleStorageAllocatedState(state, guard.arrays)
		}
		if guard.counts[lhs] && operator == "=" && literal == "0" && branch == vbacfg.EdgeBranchFalse {
			return arrayModuleStorageAllocatedState(state, guard.arrays)
		}
		if arrayModuleStorageSuccessfulResultBranch(operator, literal, branch) && arrayModuleStorageCallProvesAllocation(file, proc, ctx, guard, lhs, guardLine) {
			return arrayModuleStorageAllocatedState(state, guard.arrays)
		}
	}
	return state
}

// applyArrayModuleAllocationCondition carries the allocation summary of a
// source-local Boolean helper onto the successful branch of its call. This is
// separate from the storage-capacity recognizer because a helper may establish
// a module array through another private setup routine without exposing the
// scalar capacity idiom at the call site. For example, `If Not ValidIndex(h)`
// reaches its caller's normal path only after ValidIndex has returned True.
func applyArrayModuleAllocationCondition(state arrayFlowState, condition string, branch vbacfg.EdgeKind, guardLine int, file parsedFile, proc sourceProcedure, ctx analysisContext) arrayFlowState {
	callName, successBranch, ok := arrayModuleStorageSuccessfulCallBranch(condition)
	if !ok || branch != successBranch || guardLine <= proc.StartLine {
		return state
	}
	updated := state
	cloned := false
	forEachArrayCallAtLineUncounted(proc, guardLine, func(call procedureir.CallSite) {
		if !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), callName) && !strings.EqualFold(cleanIdentifier(call.Callee.Text), callName) {
			return
		}
		key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !resolved {
			target, resolved = arraySourceModuleTargetForCall(file, call, ctx)
			if resolved {
				key = arrayProcedureKey(target)
			}
		}
		if !resolved || !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
			return
		}
		allocated := ctx.arrayModuleAllocations[key]
		if len(allocated) == 0 {
			return
		}
		if !cloned {
			updated = cloneArrayState(state)
			cloned = true
		}
		for name := range allocated {
			if !allocated[name] {
				continue
			}
			value, tracked := updated[name]
			if !tracked {
				continue
			}
			value.kind = arrayAllocated
			value.knownArray = true
			value.mayBeUnallocated = false
			updated[name] = value
		}
	})
	return updated
}

func arrayModuleStorageSuccessfulCallBranch(condition string) (string, vbacfg.EdgeKind, bool) {
	condition = strings.TrimSpace(condition)
	if condition == "" {
		return "", "", false
	}
	branch := vbacfg.EdgeBranchTrue
	lower := strings.ToLower(condition)
	if strings.HasPrefix(lower, "if ") {
		condition = strings.TrimSpace(condition[len("if "):])
		lower = strings.ToLower(condition)
	}
	if strings.HasPrefix(lower, "not ") {
		condition = strings.TrimSpace(condition[4:])
		branch = vbacfg.EdgeBranchFalse
	}
	if strings.HasPrefix(condition, "(") && strings.HasSuffix(condition, ")") {
		condition = strings.TrimSpace(condition[1 : len(condition)-1])
	}
	if !strings.Contains(condition, "(") || strings.ContainsAny(condition, ".!<>=") {
		return "", "", false
	}
	name := strings.TrimSpace(cleanIdentifier(arrayCallName(condition)))
	if name == "" {
		return "", "", false
	}
	return name, branch, true
}

func arrayModuleStorageCallAllocatesGroup(file parsedFile, proc sourceProcedure, ctx analysisContext, callName string, guard arrayModuleStorageGuard, guardLine int) bool {
	if guardLine <= proc.StartLine || callName == "" {
		return false
	}
	for call := range proc.Calls.All() {
		if call.IsRaiseEvent || call.Range.StartLine < proc.StartLine || call.Range.StartLine >= guardLine {
			continue
		}
		if !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), callName) && !strings.EqualFold(cleanIdentifier(call.Callee.Text), callName) {
			continue
		}
		key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !resolved {
			target, resolved = arraySourceModuleTargetForCall(file, call, ctx)
			if resolved {
				key = arrayProcedureKey(target)
			}
		}
		if !resolved || !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
			continue
		}
		allocation := ctx.arrayModuleAllocations[key]
		if len(allocation) == 0 {
			continue
		}
		proven := true
		for name := range guard.arrays {
			if !allocation[strings.ToLower(cleanIdentifier(name))] {
				proven = false
				break
			}
		}
		if proven {
			return true
		}
	}
	return false
}

// arrayModuleStorageTargetDirectlyAllocatesGroup recognizes a same-module
// helper that resizes every member of a storage group in the source. It is
// used by the higher-level preparation proof below; the ordinary allocation
// summary remains deliberately stricter because an early unsupported-state
// exit can otherwise make a direct ReDim fail to dominate that summary's
// broad normal-exit set.
func arrayModuleStorageTargetDirectlyAllocatesGroup(file parsedFile, target sourceProcedure, guard arrayModuleStorageGuard) bool {
	facts := file.moduleAnalysisFacts()
	if facts == nil || arrayProcedureHasErrorHandling(target) || len(guard.arrays) == 0 {
		return false
	}
	for name := range guard.arrays {
		name = strings.ToLower(cleanIdentifier(name))
		found := false
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if found || operation.Kind != moduleArrayDirectRedim || operation.Preserve {
				return
			}
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
			found = ok && arrayModuleStorageSameProcedure(owner, target)
		})
		if !found {
			return false
		}
	}
	return true
}

// arrayModuleStoragePreparationTarget recognizes the narrow capacity-based
// dispatcher used by hash indexes: an existing nonzero capacity is already a
// storage proof, while the zero-capacity branch calls a helper that directly
// ReDims the complete group. The target itself must not erase or rewrite the
// group; those effects would invalidate the post-call proof.
func arrayModuleStoragePreparationTarget(file parsedFile, target sourceProcedure, ctx analysisContext, guard arrayModuleStorageGuard) bool {
	if arrayProcedureHasErrorHandling(target) || target.StartLine < 1 || target.EndLine < target.StartLine {
		return false
	}
	capacityGuard := false
	for statement := range target.Statements.All() {
		if statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf || statement.Condition == nil {
			continue
		}
		lhs, operator, literal, ok := arrayCountComparison(statement.Condition.Text)
		if ok && strings.EqualFold(cleanIdentifier(lhs), guard.capacity) && operator == "=" && literal == "0" {
			capacityGuard = true
			break
		}
	}
	if !capacityGuard {
		return false
	}
	facts := file.moduleAnalysisFacts()
	for name := range guard.arrays {
		name = strings.ToLower(cleanIdentifier(name))
		invalid := false
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if invalid {
				return
			}
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
			if ok && arrayModuleStorageSameProcedure(owner, target) {
				invalid = true
			}
		})
		if invalid {
			return false
		}
	}
	for call := range target.Calls.All() {
		_, callee, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !resolved || !strings.EqualFold(strings.TrimSpace(callee.Module), strings.TrimSpace(file.Module)) {
			continue
		}
		if arrayModuleStorageTargetDirectlyAllocatesGroup(file, callee, guard) {
			return true
		}
	}
	return false
}

func arrayModuleStorageCallEstablishesGroup(file parsedFile, target sourceProcedure, ctx analysisContext, guards []arrayModuleStorageGuard) map[string]bool {
	for _, guard := range guards {
		if arrayModuleStoragePreparationTarget(file, target, ctx, guard) {
			return guard.arrays
		}
	}
	return nil
}

// applyArrayModuleStorageSourceGuardState mirrors a proven branch at the
// source-line boundary. A single-line `If capacity = 0 Then Exit Function`
// can share a CFG block with the first indexed statement, so the worklist may
// report that statement before the branch refinement is visible to the next
// source-line visit. This fallback is deliberately restricted to top-level
// early-exit guards and proven storage groups; it does not interpret an
// arbitrary numeric comparison as an allocation fact.
func applyArrayModuleStorageSourceGuardState(state arrayFlowState, file parsedFile, proc sourceProcedure, line int, ctx analysisContext, guards []arrayModuleStorageGuard) arrayFlowState {
	if line <= proc.StartLine {
		return state
	}
	state = applyArrayModuleAllocationSourceGuardState(state, file, proc, line, ctx)
	if len(guards) == 0 {
		return state
	}
	for _, guard := range guards {
		for sourceLine := max(proc.StartLine, 1); sourceLine <= line && sourceLine <= len(file.Lines); sourceLine++ {
			text := strings.TrimSpace(normalizedCodeLine(file.Lines[sourceLine-1]))
			condition, body, ok := arrayIfThenParts(text)
			if !ok {
				continue
			}
			statement := procedureStatementAtLine(proc, sourceLine)
			if statement.ID == 0 || !arrayModuleStorageSourceGuardDominates(proc, sourceLine, line) {
				continue
			}
			lhs, operator, literal, comparisonOK := arrayCountComparison(condition)
			if !comparisonOK {
				if callName, successBranch, callOK := arrayModuleStorageSuccessfulCallBranch(condition); callOK && successBranch == vbacfg.EdgeBranchFalse && strings.TrimSpace(body) != "" {
					thenBody, _, hasElse := arrayIfThenBodyParts(body)
					callAllocates := arrayModuleStorageCallAllocatesGroup(file, proc, ctx, callName, guard, sourceLine)
					if !hasElse && arrayModuleStorageExitBody(thenBody, proc) && !arrayModuleStorageGroupErasedBetween(file, guard, sourceLine, line) && callAllocates {
						state = arrayModuleStorageAllocatedState(state, guard.arrays)
					}
				}
				continue
			}
			if operator != "=" || literal != "0" {
				continue
			}
			lhs = strings.ToLower(cleanIdentifier(lhs))
			if lhs == guard.capacity {
				thenBody, _, hasElse := arrayIfThenBodyParts(body)
				if !hasElse && arrayModuleStorageExitBody(thenBody, proc) && !arrayModuleStorageGroupErasedBetween(file, guard, sourceLine, line) {
					state = arrayModuleStorageAllocatedState(state, guard.arrays)
				}
				continue
			}
			safeBranch, safeResult := vbacfg.EdgeBranchTrue, false
			if arrayModuleStorageSuccessfulResultBranch(operator, literal, vbacfg.EdgeBranchFalse) {
				safeBranch, safeResult = vbacfg.EdgeBranchFalse, true
			} else if arrayModuleStorageSuccessfulResultBranch(operator, literal, vbacfg.EdgeBranchTrue) {
				safeResult = true
			}
			if safeResult {
				state = applyArrayModuleStorageCondition(state, condition, safeBranch, sourceLine, file, proc, ctx, guards)
				continue
			}
			if !guard.counts[lhs] || strings.TrimSpace(body) != "" {
				continue
			}
			end := arraySourceIfEnd(file.Lines, sourceLine-1, min(len(file.Lines), proc.EndLine))
			if end < 0 || line <= end+1 || !arrayModuleStorageMultilineExitBody(file, sourceLine, end, proc) || arrayModuleStorageGroupErasedBetween(file, guard, end+1, line) {
				continue
			}
			state = arrayModuleStorageAllocatedState(state, guard.arrays)
		}
	}
	return state
}

// applyArrayModuleAllocationSourceGuardState covers a source-level early-exit
// guard whose CFG block can also contain the first normal-path access. The
// interprocedural summary is still required; the source-order fallback only
// mirrors `If Not Helper(...) Then Exit ...` and never treats an arbitrary
// Boolean as proof of allocation.
func applyArrayModuleAllocationSourceGuardState(state arrayFlowState, file parsedFile, proc sourceProcedure, line int, ctx analysisContext) arrayFlowState {
	if line <= proc.StartLine || len(ctx.arrayModuleAllocations) == 0 {
		return state
	}
	for sourceLine := max(proc.StartLine, 1); sourceLine < line && sourceLine <= len(file.Lines); sourceLine++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[sourceLine-1]))
		condition, body, ok := arrayIfThenParts(text)
		if !ok || strings.TrimSpace(body) == "" {
			continue
		}
		statement := procedureStatementAtLine(proc, sourceLine)
		if statement.ID == 0 || !arrayModuleStorageSourceGuardDominates(proc, sourceLine, line) {
			continue
		}
		_, successBranch, callOK := arrayModuleStorageSuccessfulCallBranch(condition)
		if !callOK || successBranch != vbacfg.EdgeBranchFalse {
			continue
		}
		thenBody, _, hasElse := arrayIfThenBodyParts(body)
		if hasElse || !arrayModuleStorageExitBody(thenBody, proc) {
			continue
		}
		valid := arrayModuleAllocationSourceGuardStillValid(file, proc, sourceLine, line, condition, ctx)
		if !valid {
			continue
		}
		state = applyArrayModuleAllocationCondition(state, condition, vbacfg.EdgeBranchFalse, sourceLine, file, proc, ctx)
	}
	return state
}

func arrayModuleAllocationSourceGuardStillValid(file parsedFile, proc sourceProcedure, startLine, endLine int, condition string, ctx analysisContext) bool {
	callName, _, ok := arrayModuleStorageSuccessfulCallBranch(condition)
	if !ok {
		return false
	}
	var allocated map[string]bool
	forEachArrayCallAtLineUncounted(proc, startLine, func(call procedureir.CallSite) {
		if allocated != nil || (!strings.EqualFold(cleanIdentifier(call.Callee.BaseName), callName) && !strings.EqualFold(cleanIdentifier(call.Callee.Text), callName)) {
			return
		}
		key, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !resolved {
			target, resolved = arraySourceModuleTargetForCall(file, call, ctx)
			if resolved {
				key = arrayProcedureKey(target)
			}
		}
		if !resolved || !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(file.Module)) {
			return
		}
		if summary := ctx.arrayModuleAllocations[key]; len(summary) > 0 {
			allocated = summary
		}
	})
	if len(allocated) == 0 {
		return false
	}
	for sourceLine := startLine + 1; sourceLine < endLine && sourceLine <= len(file.Lines); sourceLine++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[sourceLine-1]))
		if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
			for _, name := range splitArgs(match[1]) {
				if allocated[strings.ToLower(cleanIdentifier(name))] {
					return false
				}
			}
		}
		if lhs, _, indexed, assigned := arrayAssignment(text); assigned && !indexed && allocated[strings.ToLower(cleanIdentifier(lhs))] {
			return false
		}
	}
	return true
}

func arrayModuleStorageSourceGuardDominates(proc sourceProcedure, sourceLine, accessLine int) bool {
	if sourceLine == accessLine {
		return true
	}
	guard := procedureStatementAtLine(proc, sourceLine)
	access := procedureStatementAtLine(proc, accessLine)
	if guard.ID == 0 || access.ID == 0 {
		return false
	}
	if proc.Graph != nil {
		return arrayVBA227StatementLineDominates(proc, sourceLine, access)
	}
	return guard.ParentID == 0
}

func arrayModuleStorageExitBody(body string, proc sourceProcedure) bool {
	body = strings.ToLower(strings.TrimSpace(body))
	switch proc.ProcedureKind {
	case procedureir.ProcedureFunction, procedureir.ProcedurePropertyGet, procedureir.ProcedureProperty:
		return body == "exit function" || body == "exit property"
	default:
		return body == "exit sub"
	}
}

func arrayModuleStorageMultilineExitBody(file parsedFile, start, end int, proc sourceProcedure) bool {
	for index := start; index < end && index < len(file.Lines); index++ {
		if arrayModuleStorageExitBody(strings.TrimSpace(normalizedCodeLine(file.Lines[index])), proc) {
			return true
		}
	}
	return false
}

func arrayModuleStorageGroupErasedBetween(file parsedFile, guard arrayModuleStorageGuard, start, end int) bool {
	start = max(start, 1)
	end = min(end, len(file.Lines))
	for line := start; line < end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		match := arrayEraseRe.FindStringSubmatch(text)
		if len(match) != 2 {
			continue
		}
		for _, target := range splitArgs(match[1]) {
			if guard.arrays[strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))] {
				return true
			}
		}
	}
	return false
}

func arrayModuleStorageSuccessfulResultBranch(operator, literal string, branch vbacfg.EdgeKind) bool {
	if literal == "0" && operator == "<" {
		return branch == vbacfg.EdgeBranchFalse
	}
	if literal == "0" && operator == ">=" {
		return branch == vbacfg.EdgeBranchTrue
	}
	if literal == "-1" && operator == ">" {
		return branch == vbacfg.EdgeBranchTrue
	}
	if literal == "-1" && operator == "<=" {
		return branch == vbacfg.EdgeBranchFalse
	}
	return false
}

func arrayModuleStorageAllocatedState(state arrayFlowState, arrays map[string]bool) arrayFlowState {
	updated := cloneArrayState(state)
	changed := false
	for name := range arrays {
		value, ok := updated[name]
		if !ok {
			continue
		}
		value.kind = arrayAllocated
		value.knownArray = true
		value.mayBeUnallocated = false
		value.allocationCountSource = ""
		value.conditionalAllocationSource = ""
		updated[name] = value
		changed = true
	}
	if !changed {
		return state
	}
	return updated
}

// applyArrayModulePositiveCountGuardBranch carries a class-owned array
// allocation through the successful side of a positive count guard. A common
// class layout initializes one or more dynamic buffers in a loader, then
// exposes only methods that first reject token/index values outside a module
// count. The count is a valid witness only when the source proves that it is
// written positively after a plain ReDim and reset before the corresponding
// Erase; an arbitrary scalar comparison must remain conservative.
func applyArrayModulePositiveCountGuardBranch(state arrayFlowState, statement *procedureir.Statement, edge vbacfg.Edge, file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) arrayFlowState {
	if statement == nil || edge.Kind != vbacfg.EdgeBranchFalse || statement.Condition == nil {
		return state
	}
	count, ok := arrayModulePositiveCountGuardName(statement.Condition.Text)
	if !ok {
		return state
	}
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	if declarations.shadowsModule(count) {
		return state
	}
	return arrayModuleStorageAllocatedState(state, arrayModulePositiveCountGuardArrays(file, moduleDecls, count, ctx))
}

func arrayModulePositiveCountGuardName(condition string) (string, bool) {
	condition = strings.TrimSpace(condition)
	if parsed, _, ok := arrayIfThenParts(condition); ok {
		condition = parsed
	}
	lower := strings.ToLower(strings.TrimSpace(condition))
	for _, prefix := range []string{"if ", "elseif ", "else if "} {
		if strings.HasPrefix(lower, prefix) {
			condition = strings.TrimSpace(condition[len(prefix):])
			break
		}
	}
	if then := arrayTopLevelKeywordIndex(condition, "then"); then >= 0 {
		condition = strings.TrimSpace(condition[:then])
	}
	terms := arrayConditionOrRe.Split(condition, -1)
	if len(terms) == 2 {
		left, leftOK := parseArrayModuleCountGuardComparison(terms[0])
		right, rightOK := parseArrayModuleCountGuardComparison(terms[1])
		if !leftOK || !rightOK || left.lhs != right.lhs {
			return "", false
		}
		lowerBound := func(value arrayModuleCountGuardComparison) bool {
			return value.rhsIsLiteral && ((value.operator == "<" && value.rhs == 1) || (value.operator == "<=" && value.rhs == 0))
		}
		upperBound := func(value arrayModuleCountGuardComparison) (string, bool) {
			if value.rhsIsLiteral || value.operator != ">" || value.rhsName == "" {
				return "", false
			}
			return value.rhsName, true
		}
		if lowerBound(left) {
			return upperBound(right)
		}
		if lowerBound(right) {
			return upperBound(left)
		}
		return "", false
	}
	if len(terms) != 1 {
		return "", false
	}
	comparison, ok := parseArrayModuleCountGuardComparison(terms[0])
	if !ok || comparison.rhsIsLiteral && comparison.rhs != 0 {
		return "", false
	}
	if comparison.operator != "=" && comparison.operator != "<=" && comparison.operator != "<" {
		return "", false
	}
	if comparison.rhsIsLiteral {
		if comparison.rhs != 0 && (comparison.operator != "<" || comparison.rhs != 1) {
			return "", false
		}
		return comparison.lhs, true
	}
	return "", false
}

type arrayModuleCountGuardComparison struct {
	lhs          string
	operator     string
	rhs          int
	rhsName      string
	rhsIsLiteral bool
}

func parseArrayModuleCountGuardComparison(text string) (arrayModuleCountGuardComparison, bool) {
	match := arrayModuleCountComparisonRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) != 4 {
		return arrayModuleCountGuardComparison{}, false
	}
	comparison := arrayModuleCountGuardComparison{
		lhs:      strings.ToLower(cleanIdentifier(match[1])),
		operator: match[2],
	}
	if comparison.lhs == "" {
		return arrayModuleCountGuardComparison{}, false
	}
	if value, ok := integerLiteral(match[3]); ok {
		comparison.rhs = value
		comparison.rhsIsLiteral = true
	} else {
		comparison.rhsName = strings.ToLower(cleanIdentifier(match[3]))
		if comparison.rhsName == "" {
			return arrayModuleCountGuardComparison{}, false
		}
	}
	return comparison, true
}

func arrayModuleCountGuardCacheLookup(ctx analysisContext, key string) (map[string]bool, bool) {
	if ctx.arrayModuleCountGuardArrays == nil {
		return nil, false
	}
	if ctx.arrayModuleCountGuardArraysMu != nil {
		ctx.arrayModuleCountGuardArraysMu.RLock()
		defer ctx.arrayModuleCountGuardArraysMu.RUnlock()
	}
	arrays, ok := ctx.arrayModuleCountGuardArrays[key]
	return arrays, ok
}

func arrayModuleCountGuardCacheStore(ctx analysisContext, key string, arrays map[string]bool) {
	if ctx.arrayModuleCountGuardArrays == nil {
		return
	}
	if ctx.arrayModuleCountGuardArraysMu != nil {
		ctx.arrayModuleCountGuardArraysMu.Lock()
		defer ctx.arrayModuleCountGuardArraysMu.Unlock()
	}
	ctx.arrayModuleCountGuardArrays[key] = arrays
}

func arrayModulePositiveCountGuardArrays(file parsedFile, moduleDecls map[string]sourceDeclaration, count string, ctx analysisContext) map[string]bool {
	count = strings.ToLower(cleanIdentifier(count))
	cacheKey := strings.ToLower(strings.TrimSpace(file.Path)) + "\x00" + count
	if arrays, ok := arrayModuleCountGuardCacheLookup(ctx, cacheKey); ok {
		return arrays
	}
	result := map[string]bool{}
	if !strings.EqualFold(strings.TrimSpace(file.ModuleKind), "class") || count == "" {
		arrayModuleCountGuardCacheStore(ctx, cacheKey, result)
		return result
	}
	countDeclaration, countDeclared := moduleDeclarationForName(moduleDecls, count)
	if !countDeclared || countDeclaration.Array || countDeclaration.Object || countDeclaration.Parameter || !arrayModuleReadyGuardSourceOwned(file, countDeclaration) {
		arrayModuleCountGuardCacheStore(ctx, cacheKey, result)
		return result
	}
	facts := file.moduleAnalysisFacts()
	if facts == nil {
		arrayModuleCountGuardCacheStore(ctx, cacheKey, result)
		return result
	}

	var zeros, positives []arrayModuleCountOperation
	facts.forEachArrayOperationFor(count, func(operation moduleArrayOperationFact) {
		if operation.Kind != moduleArrayWholeAssignment {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
		if !ok || arrayProcedureHasErrorHandling(owner) {
			return
		}
		entry := arrayModuleCountOperation{fact: operation, owner: owner, ok: true}
		if value, literal := integerLiteral(operation.RHS); literal && value == 0 {
			zeros = append(zeros, entry)
		}
		if arrayModulePositiveCountAssignment(operation.RHS, count) {
			positives = append(positives, entry)
		}
	})
	if len(zeros) == 0 || len(positives) == 0 {
		arrayModuleCountGuardCacheStore(ctx, cacheKey, result)
		return result
	}

	arrays := make(map[string]bool)
	for name, declaration := range moduleDecls {
		name = strings.ToLower(cleanIdentifier(name))
		if name == "" || !declaration.Array || declaration.Fixed || declaration.Parameter {
			continue
		}
		var accepted bool
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if accepted || operation.Kind != moduleArrayDirectRedim || operation.Preserve {
				return
			}
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
			if !ok || arrayProcedureHasErrorHandling(owner) {
				return
			}
			for _, zero := range zeros {
				if !zero.ok || !arrayModuleStorageSameProcedure(zero.owner, owner) || zero.fact.Line >= operation.Line {
					continue
				}
				if !arrayModulePositiveCountWriteReachable(file, owner, operation.Line+1, positives, ctx) {
					continue
				}
				if !arrayModulePositiveCountErasesSafe(file, name, owner, zero.fact.Line, ctx) {
					continue
				}
				accepted = true
				break
			}
		})
		if accepted {
			arrays[name] = true
		}
	}
	for name := range arrays {
		result[name] = true
	}
	arrayModuleCountGuardCacheStore(ctx, cacheKey, result)
	return result
}

func arrayModulePositiveCountAssignment(rhs, count string) bool {
	if value, ok := integerLiteral(rhs); ok {
		return value > 0
	}
	compact := canonicalArrayBoundExpression(rhs)
	prefix := count + "+"
	if !strings.HasPrefix(compact, prefix) {
		return false
	}
	value, ok := integerLiteral(strings.TrimPrefix(compact, prefix))
	return ok && value > 0
}

type arrayModuleCountOperation struct {
	fact  moduleArrayOperationFact
	owner sourceProcedure
	ok    bool
}

func arrayModulePositiveCountWriteReachable(file parsedFile, setup sourceProcedure, redimLine int, positives []arrayModuleCountOperation, ctx analysisContext) bool {
	for _, positive := range positives {
		if !positive.ok {
			continue
		}
		if arrayModuleStorageSameProcedure(setup, positive.owner) {
			if positive.fact.Line+1 > redimLine {
				return true
			}
			continue
		}
		if arrayModuleProcedureReachesAfter(file, setup, positive.owner, redimLine, ctx, map[string]bool{}) {
			return true
		}
	}
	return false
}

func arrayModuleProcedureReachesAfter(file parsedFile, from, target sourceProcedure, afterLine int, ctx analysisContext, visiting map[string]bool) bool {
	fromKey := arrayProcedureKey(from)
	if visiting[fromKey] {
		return false
	}
	visiting[fromKey] = true
	defer delete(visiting, fromKey)
	for call := range from.Calls.All() {
		if afterLine > 0 && call.Range.StartLine <= afterLine {
			continue
		}
		_, callee, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !resolved {
			callee, resolved = arraySourceModuleTargetForCall(file, call, ctx)
		}
		if !resolved || !strings.EqualFold(strings.TrimSpace(callee.Module), strings.TrimSpace(from.Module)) {
			continue
		}
		if arrayModuleStorageSameProcedure(callee, target) || arrayModuleProcedureReachesAfter(file, callee, target, 0, ctx, visiting) {
			return true
		}
	}
	return false
}

func arrayModulePositiveCountErasesSafe(file parsedFile, arrayName string, setup sourceProcedure, zeroLine int, ctx analysisContext) bool {
	facts := file.moduleAnalysisFacts()
	if facts == nil {
		return false
	}
	safe := true
	facts.forEachArrayOperationFor(arrayName, func(operation moduleArrayOperationFact) {
		if !safe || operation.Kind != moduleArrayErase {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
		if !ok {
			safe = false
			return
		}
		if arrayModuleStorageSameProcedure(owner, setup) {
			if operation.Line > zeroLine {
				safe = false
			}
			return
		}
		if strings.EqualFold(strings.TrimSpace(owner.Name), "class_terminate") {
			return
		}
		calledBeforeReset := false
		for caller := range file.procedureView().All() {
			for call := range caller.Calls.All() {
				_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
				if !resolved || !arrayModuleStorageSameProcedure(target, owner) {
					continue
				}
				if arrayModuleStorageSameProcedure(caller, setup) && call.Range.StartLine <= zeroLine+1 {
					calledBeforeReset = true
					continue
				}
				if strings.EqualFold(strings.TrimSpace(caller.Name), "class_terminate") {
					calledBeforeReset = true
					continue
				}
				safe = false
			}
		}
		if !calledBeforeReset {
			safe = false
		}
	})
	return safe
}

func arrayModuleStorageCallProvesAllocation(file parsedFile, proc sourceProcedure, ctx analysisContext, guard arrayModuleStorageGuard, resultName string, guardLine int) bool {
	if guardLine <= proc.StartLine {
		return false
	}
	resultName = strings.ToLower(cleanIdentifier(resultName))
	for line := proc.StartLine; line < guardLine && line <= len(file.Lines); line++ {
		for _, part := range splitRangeValueSourceStatements(arrayLogicalCodeLine(file.Lines, line)) {
			lhs, rhs, indexed, assigned := arrayAssignment(part)
			if !assigned || indexed || !strings.EqualFold(cleanIdentifier(lhs), resultName) {
				continue
			}
			callName := strings.ToLower(cleanIdentifier(arrayCallName(rhs)))
			if callName == "" {
				continue
			}
			matched := false
			forEachArrayCallAtLineUncounted(proc, line, func(call procedureir.CallSite) {
				if matched || !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), callName) && !strings.EqualFold(cleanIdentifier(call.Callee.Text), callName) {
					return
				}
				target, resolved := arraySourceModuleTargetForCall(file, call, ctx)
				if key, privateTarget, privateResolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call); privateResolved {
					_ = key
					target, resolved = privateTarget, true
				}
				if resolved && arrayModuleStorageTargetHasCapacityGuard(file, target, guard.capacity) {
					matched = true
				}
			})
			if matched {
				return true
			}
		}
	}
	return false
}

func arrayModuleStorageTargetHasCapacityGuard(file parsedFile, proc sourceProcedure, capacity string) bool {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet || arrayProcedureHasErrorHandling(proc) {
		return false
	}
	defaultLine := 0
	guardLine := 0
	start := max(1, proc.StartLine)
	end := min(len(file.Lines), proc.EndLine)
	for line := start; line <= end; line++ {
		for _, part := range splitRangeValueSourceStatements(arrayLogicalCodeLine(file.Lines, line)) {
			if lhs, rhs, indexed, assigned := arrayAssignment(part); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), proc.Name) {
				if value, ok := integerLiteral(rhs); ok && value < 0 {
					defaultLine = line
				}
			}
			condition, body, ok := arrayIfThenParts(part)
			if !ok || strings.TrimSpace(body) == "" {
				continue
			}
			lhs, operator, literal, comparisonOK := arrayCountComparison(condition)
			thenBody, _, hasElse := arrayIfThenBodyParts(body)
			if comparisonOK && strings.EqualFold(cleanIdentifier(lhs), capacity) && operator == "=" && literal == "0" && !hasElse && strings.EqualFold(strings.TrimSpace(thenBody), "exit function") {
				guardLine = line
			}
		}
	}
	return defaultLine > 0 && guardLine > defaultLine
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
		return nil
	}
	if ctx.arrayModuleInvalidations != nil {
		if summary, ok := ctx.arrayModuleInvalidations[key]; ok {
			return summary
		}
	}
	ctx.arrayStats.addModuleInvalidationSummary(target.Graph != nil)
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
			forEachArrayCallAtLine(proc, line, func(call procedureir.CallSite) {
				if conditionKnown && !arrayInlineConditionalCallIsReachable(file, call, conditionValue, hasElse) {
					return
				}
				value = applyArrayModuleSummaryCallEffects(value, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
			}, ctx.arrayStats)
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
	forEachArrayCallAtLine(proc, line, func(call procedureir.CallSite) {
		state = applyArrayModuleSummaryCallEffects(state, file, proc, call, ctx, variables, moduleDecls, moduleArrays, visiting)
	}, ctx.arrayStats)
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
	if !arrayProcedureIsModuleEffectParticipant(ctx, target) {
		return updated
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

type arraySourceModuleTargetCacheKey struct {
	file        string
	module      string
	caller      string
	callee      string
	baseName    string
	receiver    string
	member      string
	id          int
	startByte   int
	endByte     int
	statementID int
}

type arraySourceModuleTargetCacheEntry struct {
	target   sourceProcedure
	resolved bool
}

func arraySourceModuleTargetCacheKeyForCall(file parsedFile, call procedureir.CallSite) (arraySourceModuleTargetCacheKey, bool) {
	fileKey := file.Path
	if fileKey == "" {
		fileKey = file.IR.Path
	}
	if fileKey == "" {
		return arraySourceModuleTargetCacheKey{}, false
	}
	receiver := ""
	if call.Callee.Receiver != nil {
		receiver = *call.Callee.Receiver
	}
	return arraySourceModuleTargetCacheKey{
		file:        fileKey,
		module:      call.Module,
		caller:      call.Caller.QualifiedName,
		callee:      call.Callee.Text,
		baseName:    call.Callee.BaseName,
		receiver:    receiver,
		member:      call.Callee.Member,
		id:          call.ID,
		startByte:   call.Range.StartByte,
		endByte:     call.Range.EndByte,
		statementID: call.StatementID,
	}, true
}

// arraySourceModuleTargetForCall recovers a source-local target for the
// source-order fallback. The project resolver may report a public target, but
// it may also be incomplete for a recovered call. In the latter case a unique
// same-module source procedure is sufficient evidence for inspecting that
// procedure's module-array accesses; an absent target remains an external or
// late-bound call and must not invalidate all module arrays.
func arraySourceModuleTargetForCall(file parsedFile, call procedureir.CallSite, ctx analysisContext) (sourceProcedure, bool) {
	// With an explicit project resolver the call target is determined by the
	// immutable project context. Contexts without one may rely on the call's
	// embedded Resolution, which is intentionally not part of this cache key.
	cacheable := ctx.procedureResolver != nil
	key := arraySourceModuleTargetCacheKey{}
	if cacheable {
		key, cacheable = arraySourceModuleTargetCacheKeyForCall(file, call)
	}
	if cacheable && ctx.arraySourceModuleTargetCache != nil {
		if ctx.arraySourceModuleTargetCacheMu != nil {
			ctx.arraySourceModuleTargetCacheMu.RLock()
			cached, ok := ctx.arraySourceModuleTargetCache[key]
			ctx.arraySourceModuleTargetCacheMu.RUnlock()
			if ok {
				return cached.target, cached.resolved
			}
		} else if cached, ok := ctx.arraySourceModuleTargetCache[key]; ok {
			return cached.target, cached.resolved
		}
	}

	target, resolved := arraySourceModuleTargetForCallUncached(file, call, ctx)
	if !cacheable || ctx.arraySourceModuleTargetCache == nil {
		return target, resolved
	}
	entry := arraySourceModuleTargetCacheEntry{target: target, resolved: resolved}
	if ctx.arraySourceModuleTargetCacheMu != nil {
		ctx.arraySourceModuleTargetCacheMu.Lock()
		if cached, ok := ctx.arraySourceModuleTargetCache[key]; ok {
			entry = cached
		} else {
			ctx.arraySourceModuleTargetCache[key] = entry
		}
		ctx.arraySourceModuleTargetCacheMu.Unlock()
	} else {
		ctx.arraySourceModuleTargetCache[key] = entry
	}
	return entry.target, entry.resolved
}

func arraySourceModuleTargetForCallUncached(file parsedFile, call procedureir.CallSite, ctx analysisContext) (sourceProcedure, bool) {
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
	case "rebuildprimarykeyindex", "indexdatarow":
		return configurations.dataTable
	case "createcollectionnode":
		return configurations.genericCollection
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
		value := arrayModuleAllocationSummaryForProcedure(procedure.file, procedure.proc, procedure.moduleDecls, targets, summaries, byRefSummaries, ctx, dominators)
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

func arrayModuleAllocationSummaryForProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, targets map[string]sourceProcedure, summaries arrayModuleAllocationSummaries, byRefSummaries arrayByRefAllocationSummaries, ctx analysisContext, dominators arrayProcedureDominators) map[string]bool {
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
	if len(moduleArrays) == 0 {
		return nil
	}
	key := arrayProcedureKey(proc)
	normalExitDominators, ok := dominators[key]
	if !ok {
		normalExitDominators = arrayProcedureNormalExitDominators(proc)
		dominators[key] = normalExitDominators
	}
	idempotentSetupArrays := arrayModuleIdempotentSetupArrays(file, proc, moduleDecls, ctx)
	countSetupArrays := arrayModuleCountSetupArrays(file, proc, moduleDecls, ctx)
	allocated := map[string]bool{}
	addDirectAllocation := func(statementID int, name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if !moduleArrays[name] || (!arrayProcedureBlockDominatesNormalExit(proc, statementID, normalExitDominators) && !idempotentSetupArrays[name] && !countSetupArrays[name]) {
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
				if value, known := arrayExpressionStateForProcedure(rhs, arrayFlowState{}, ctx, proc); known && value.kind == arrayAllocated && value.knownArray {
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
		guaranteed := !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) && arrayProcedureBlockDominatesNormalExit(proc, call.StatementID, normalExitDominators)
		if !guaranteed && arrayProcedureHasIdempotentSetupGuard(file, proc, call.Range.StartLine, moduleDecls) {
			guaranteed = true
		}
		if !guaranteed && arrayProcedureCallPrecedesSuccessfulReturn(file, proc, call) {
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

// arrayModuleCountSetupArrays recognizes the numeric ready-count variant of
// the one-time module initialization idiom:
//
//	If connectionCount > 0 Then Exit Sub
//	ReDim connections(0 To MAX_CONNECTIONS - 1)
//	connectionCount = MAX_CONNECTIONS
//
// A positive count is written only after every direct ReDim in the helper has
// completed. Since VBA leaves a module Long at zero initially and this helper
// has no error handler or erase/reset path, a normal return with a positive
// count proves the grouped arrays are allocated. The proof is intentionally
// limited to a source-owned scalar, a single final positive write, direct
// non-Preserve ReDim operations, and calls whose module effects are already
// modeled.
func arrayModuleCountSetupArrays(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext) map[string]bool {
	if proc.Graph == nil || arrayProcedureHasErrorHandling(proc) || proc.StartLine < 1 || proc.EndLine < proc.StartLine || proc.StartLine > len(file.Lines) {
		return nil
	}
	facts := file.moduleAnalysisFacts()
	start := max(0, proc.StartLine-1)
	end := min(len(file.Lines), proc.EndLine)
	guardLine := -1
	guardName := ""
	for line := start; line < end; line++ {
		condition, body, ok := arrayIfThenParts(normalizedCodeLine(file.Lines[line]))
		if !ok || !strings.EqualFold(strings.TrimSpace(body), "exit sub") {
			continue
		}
		lhs, operator, literal, comparisonOK := arrayCountComparison(condition)
		if !comparisonOK || operator != ">" || literal != "0" {
			continue
		}
		name := strings.ToLower(cleanIdentifier(lhs))
		declaration, declared := moduleDecls[name]
		if !declared || declaration.Array || declaration.Parameter || !arrayKnownScalarType(declaration.Type) {
			continue
		}
		if guardLine >= 0 {
			return nil
		}
		guardLine = line
		guardName = name
	}
	if guardLine < 0 {
		return nil
	}

	constants := arrayIntegerConstants(file, proc, nil, nil)
	var countWrite moduleArrayOperationFact
	countWrites := 0
	facts.forEachArrayOperationFor(guardName, func(operation moduleArrayOperationFact) {
		if operation.Kind != moduleArrayWholeAssignment {
			return
		}
		countWrites++
		countWrite = operation
	})
	if countWrites != 1 || countWrite.Line <= guardLine {
		return nil
	}
	positiveCount, err := constantIntegerExpression(strings.TrimSpace(countWrite.RHS), constants)
	if err != nil || positiveCount <= 0 {
		return nil
	}
	for line := countWrite.Line + 1; line < end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line]))
		lower := strings.ToLower(text)
		if text == "" || strings.HasPrefix(text, "'") || lower == "end sub" || lower == "end function" || lower == "end property" {
			continue
		}
		return nil
	}

	firstRedimLine := countWrite.Line
	for name, declaration := range moduleDecls {
		if name == "" || !declaration.Array || declaration.Fixed || declaration.Parameter {
			continue
		}
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if operation.Kind == moduleArrayDirectRedim && !operation.Preserve && operation.Line > guardLine && operation.Line < firstRedimLine {
				firstRedimLine = operation.Line
			}
		})
	}

	// Calls between the first allocation and the ready-count write must not be
	// able to erase or replace the grouped arrays before the invariant becomes
	// observable to callers. Calls before the first ReDim cannot invalidate the
	// proof: the guarded path would have exited, while the fresh path has not
	// established the grouped allocation yet.
	for call := range proc.Calls.All() {
		line := call.Range.StartLine - 1
		if line <= firstRedimLine || line > countWrite.Line {
			continue
		}
		if arrayCallIsIndexedArrayAccess(proc, call, arrayVariables(file, proc, moduleDecls)) || call.IsRaiseEvent || call.Resolution.Status == procedureir.ResolutionBuiltinLike {
			continue
		}
		_, target, resolved := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
		if !resolved || arrayPrivateCallMayInvalidateModuleArray(file, proc, target, call, moduleDecls, ctx) {
			return nil
		}
	}

	result := map[string]bool{}
	for name, declaration := range moduleDecls {
		name = strings.ToLower(cleanIdentifier(name))
		if name == "" || !declaration.Array || declaration.Fixed || declaration.Parameter {
			continue
		}
		redimCount := 0
		valid := true
		facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
			if !valid {
				return
			}
			if operation.Kind != moduleArrayDirectRedim || operation.Preserve || operation.Line <= guardLine || operation.Line >= countWrite.Line {
				valid = false
				return
			}
			owner, ownerOK := arrayModuleProcedureAtLine(file, operation.Line+1)
			if !ownerOK || arrayProcedureKey(owner) != arrayProcedureKey(proc) {
				valid = false
				return
			}
			redimCount++
		})
		if valid && redimCount > 0 {
			result[name] = true
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
		if !hasModuleDynamicArrayDeclaration(moduleDecls) {
			continue
		}
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
				owner, owned := arrayModuleProcedureAtLine(file, operation.Line+1, ctx.arrayStats)
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
			ctx.arrayStats.addModuleReadyGuardCandidate()
			writer, ok := arrayModuleProcedureAtLine(file, trueWrites[0].Line+1, ctx.arrayStats)
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

func arrayModuleProcedureAtLine(file parsedFile, line int, stats ...*arrayInterproceduralStats) (sourceProcedure, bool) {
	if line < 1 {
		return sourceProcedure{}, false
	}
	if facts := file.moduleAnalysisFacts(); facts != nil && len(facts.procedureRanges) > 0 && line < len(facts.procedureLineOwners) {
		owner := facts.procedureLineOwners[line]
		if owner >= 0 && owner < len(facts.procedureRanges) {
			if len(stats) > 0 {
				stats[0].addProcedureRangeIndexHit()
			}
			return facts.procedureRanges[owner], true
		}
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
	ctx.arrayStats.addModuleReadyGuardCFGWalk()

	graph := arrayVBA227Graph(proc, ctx)
	resumeNextEdges := arrayVBA227ResumeNextContinuationEdges(proc)
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
		out, _ := (Analyzer{}).arrayVBA227Transfer(file, proc, ctx, variables, in, text, line, nil, nil, nil, &graph, resumeNextEdges)
		forEachArrayCallAtLine(proc, line, func(call procedureir.CallSite) {
			out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
			out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
		}, ctx.arrayStats)
		return out
	}
	edgeState := func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
		out = applyArrayConditionalAllocationBranch(out, &graph, block, edge)
		out = applyArrayResumeNextFailureFlagBranch(out, block.Statement, edge)
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
	for name := range arrayModuleReadyGuardObjectSetupAllocationProof(file, proc, readyLine, candidates, moduleDecls) {
		result[name] = true
	}
	return result
}

// arrayModuleReadyGuardObjectSetupAllocationProof handles a two-stage module
// initializer whose object handle is the source-owned witness for an array
// allocation. A common UserForm pattern allocates module arrays in
// AttachForm, stores the form in a Private Object field, and later lets
// BuildScene set a Boolean ready flag only after `If mForm Is Nothing Then Exit
// Sub`. The Boolean guard's writer does not itself contain the ReDim, so the
// ordinary ready-writer CFG walk cannot see this invariant.
//
// The proof is deliberately narrow: the object field is Private, every
// non-Nothing assignment to it is a parameter assignment in one setup
// procedure, that procedure directly allocates every dynamic candidate array,
// and the object Nothing guard dominates the ready write.
func arrayModuleReadyGuardObjectSetupAllocationProof(file parsedFile, writer sourceProcedure, readyLine int, candidates map[string]bool, moduleDecls map[string]sourceDeclaration) map[string]bool {
	if writer.Graph == nil || readyLine <= writer.StartLine || readyLine > len(file.Lines) {
		return nil
	}
	objectName, guardLine, ok := arrayModuleObjectNothingExitGuard(file, writer, readyLine)
	if !ok {
		return nil
	}
	objectDeclaration, declared := moduleDecls[objectName]
	if !declared || !objectDeclaration.Object || objectDeclaration.Array || objectDeclaration.Parameter || !arrayModuleReadyGuardSourceOwned(file, objectDeclaration) {
		return nil
	}
	guardStatement, guardOK := arrayModuleStatementAtLine(writer, guardLine)
	readyStatement, readyOK := arrayModuleStatementAtLine(writer, readyLine)
	if !guardOK || !readyOK {
		return nil
	}
	normalGraph := writer.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	guardBlock, guardBlockOK := writer.Graph.BlockForStatement(guardStatement.ID)
	readyBlock, readyBlockOK := writer.Graph.BlockForStatement(readyStatement.ID)
	if !guardBlockOK || !readyBlockOK || !normalGraph.Dominates(guardBlock.ID, readyBlock.ID) {
		return nil
	}

	facts := file.moduleAnalysisFacts()
	if facts == nil {
		return nil
	}
	setup, ok := arrayModuleObjectSetupProcedure(file, objectName, candidates, moduleDecls, facts)
	if !ok {
		return nil
	}
	result := make(map[string]bool)
	for name := range candidates {
		declaration, declared := moduleDecls[name]
		if !declared || !declaration.Array || declaration.Fixed {
			continue
		}
		if arrayModuleSetupDirectlyAllocates(file, setup, name, facts) {
			result[name] = true
		}
	}
	return result
}

func arrayModuleObjectNothingExitGuard(file parsedFile, proc sourceProcedure, readyLine int) (string, int, bool) {
	for line := proc.StartLine + 1; line < readyLine && line <= proc.EndLine && line <= len(file.Lines); line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		match := arrayModuleObjectNothingExitRe.FindStringSubmatch(text)
		if len(match) != 2 {
			continue
		}
		objectName := strings.ToLower(cleanIdentifier(match[1]))
		for next := line + 1; next <= proc.EndLine && next <= len(file.Lines); next++ {
			body := strings.TrimSpace(strings.ToLower(normalizedCodeLine(file.Lines[next-1])))
			if body == "" || strings.HasPrefix(body, "'") {
				continue
			}
			if body == "exit sub" || body == "exit function" || body == "exit property" {
				return objectName, line, true
			}
			break
		}
	}
	return "", 0, false
}

func arrayModuleObjectSetupProcedure(file parsedFile, objectName string, candidates map[string]bool, moduleDecls map[string]sourceDeclaration, facts *moduleAnalysisFacts) (sourceProcedure, bool) {
	var match sourceProcedure
	matched := false
	procedures := file.procedureView()
	for index := 0; index < procedures.Len(); index++ {
		procedure := procedures.valueAt(index)
		if procedure.StartLine >= procedure.EndLine || arrayProcedureHasErrorHandling(procedure) || !arrayModuleProcedureSetsObjectParameter(file, procedure, objectName) {
			continue
		}
		allAllocated := true
		for name, declaration := range moduleDecls {
			name = strings.ToLower(cleanIdentifier(name))
			if name == "" || !candidates[name] || !declaration.Array || declaration.Fixed {
				continue
			}
			if !arrayModuleSetupDirectlyAllocates(file, procedure, name, facts) {
				allAllocated = false
				break
			}
		}
		if !allAllocated || !arrayModuleObjectAssignmentsBelongToSetup(file, objectName, procedure) {
			continue
		}
		if matched {
			return sourceProcedure{}, false
		}
		match = procedure
		matched = true
	}
	return match, matched
}

func arrayModuleProcedureSetsObjectParameter(file parsedFile, proc sourceProcedure, objectName string) bool {
	parameters := make(map[string]bool)
	for parameter := range proc.Params.All() {
		if isObjectType(parameter.Type) && !parameter.IsArray {
			parameters[strings.ToLower(cleanIdentifier(parameter.Name))] = true
		}
	}
	startLine := proc.StartLine
	if startLine < 1 {
		startLine = 1
	}
	for line := startLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		match := arrayModuleObjectSetRe.FindStringSubmatch(text)
		if len(match) == 3 && strings.EqualFold(cleanIdentifier(match[1]), objectName) && strings.ToLower(cleanIdentifier(match[2])) != "nothing" && parameters[strings.ToLower(cleanIdentifier(match[2]))] {
			return true
		}
	}
	return false
}

func arrayModuleObjectAssignmentsBelongToSetup(file parsedFile, objectName string, setup sourceProcedure) bool {
	seen := false
	for line := 1; line <= len(file.Lines); line++ {
		match := arrayModuleObjectSetRe.FindStringSubmatch(strings.TrimSpace(normalizedCodeLine(file.Lines[line-1])))
		if len(match) != 3 || !strings.EqualFold(cleanIdentifier(match[1]), objectName) || strings.EqualFold(cleanIdentifier(match[2]), "nothing") {
			continue
		}
		owner, ok := arrayModuleProcedureAtLine(file, line)
		if !ok || owner.StartByte != setup.StartByte || owner.StartLine != setup.StartLine || owner.EndLine != setup.EndLine || !arrayModuleProcedureSetsObjectParameter(file, setup, objectName) {
			return false
		}
		seen = true
	}
	return seen
}

func arrayModuleSetupDirectlyAllocates(file parsedFile, proc sourceProcedure, name string, facts *moduleAnalysisFacts) bool {
	allocated := false
	facts.forEachArrayOperationFor(name, func(operation moduleArrayOperationFact) {
		if allocated || operation.Kind != moduleArrayDirectRedim || operation.Preserve {
			return
		}
		owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1)
		allocated = ok && owner.StartByte == proc.StartByte && owner.StartLine == proc.StartLine && owner.EndLine == proc.EndLine
	})
	return allocated
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
			owner, ok := arrayModuleProcedureAtLine(file, operation.Line+1, ctx.arrayStats)
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
				value, known := arrayExpressionStateForProcedure(rhs, arrayFlowState{}, ctx, owner)
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
		if normalGraph.Dominates(falseBlock.ID, eraseBlock.ID) {
			return true
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
	dominatesReady := normalGraph.Dominates(redimBlock.ID, readyBlock.ID)
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
	_, body, ok := arrayIfThenParts(normalizedCodeLine(file.Lines[line-1]))
	return ok && strings.TrimSpace(body) != ""
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

// arrayInlineConditionalCallIsConditionExpression distinguishes a call made
// while evaluating an inline If condition from a call in the conditional body.
// The former is always evaluated before VBA chooses the body, so an
// unconditional module-allocation summary may be applied even when the line
// itself also contains an Exit statement.
func arrayInlineConditionalCallIsConditionExpression(file parsedFile, call procedureir.CallSite) bool {
	if !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) || call.Range.StartColumn <= 0 || call.Range.StartLine < 1 || call.Range.StartLine > len(file.Lines) {
		return false
	}
	raw := gui.StripComment(file.Lines[call.Range.StartLine-1])
	trimmed := strings.TrimSpace(raw)
	leading := len(raw) - len(strings.TrimLeft(raw, " \t"))
	lower := strings.ToLower(trimmed)
	prefixLength := 0
	switch {
	case strings.HasPrefix(lower, "if "):
		prefixLength = len("if ")
	case strings.HasPrefix(lower, "elseif "):
		prefixLength = len("elseif ")
	default:
		return false
	}
	rest := strings.TrimSpace(trimmed[prefixLength:])
	thenIndex := arrayTopLevelKeywordIndex(rest, "then")
	if thenIndex < 0 {
		return false
	}
	bodyStart := leading + prefixLength + (len(trimmed[prefixLength:]) - len(rest)) + thenIndex + len("then")
	return call.Range.StartColumn-1 < bodyStart
}

// arrayProcedureCallPrecedesSuccessfulReturn admits a helper allocation when
// the call is followed by the procedure's successful Boolean return on a
// straight-line path. This covers guards such as:
//
//	InitConnectionPool
//	ValidIndex = True
//
// An early invalid-input Exit before the call does not reach that return, so
// requiring the successful assignment after the call avoids treating that
// bypass as an allocation proof.
func arrayProcedureCallPrecedesSuccessfulReturn(file parsedFile, proc sourceProcedure, call procedureir.CallSite) bool {
	if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
		return false
	}
	start := max(proc.StartLine, call.Range.StartLine)
	end := min(len(file.Lines), proc.EndLine)
	for line := start; line <= end; line++ {
		text := strings.TrimSpace(normalizedCodeLine(file.Lines[line-1]))
		if text == "" || strings.HasPrefix(text, "'") {
			continue
		}
		if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed && strings.EqualFold(cleanIdentifier(lhs), proc.Name) && strings.EqualFold(strings.TrimSpace(rhs), "true") {
			return true
		}
		lower := strings.ToLower(text)
		if strings.HasPrefix(lower, "exit function") || strings.HasPrefix(lower, "exit property") || arrayModuleSetupLineHasControlFlow(text) {
			return false
		}
	}
	return false
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
	dominators := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).DominatorsOf(proc.Graph.NormalExit)
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
type arrayModuleProcedureInfo struct {
	file        parsedFile
	proc        sourceProcedure
	key         string
	moduleDecls map[string]sourceDeclaration
	variables   map[string]arrayVariable
}

func inferArrayModuleEntryStates(a Analyzer, files []parsedFile, ctx analysisContext) arrayModuleEntryStates {
	if len(ctx.arrayPrivateTargets) == 0 {
		return arrayModuleEntryStates{}
	}

	procedures := make([]arrayModuleProcedureInfo, 0)
	allProcedures := make([]arrayModuleProcedureInfo, 0)
	moduleArrays := map[string]map[string]bool{}
	moduleFiles := map[string]string{}
	for _, file := range files {
		procs := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procs.Len(); procedureIndex++ {
			proc := procs.valueAt(procedureIndex)
			info := arrayModuleProcedureInfo{
				file: file, proc: proc, moduleDecls: moduleDecls,
				variables: arrayVariables(file, proc, moduleDecls),
				key:       arrayParticipantLookupKey(proc, ctx.arrayParticipantKeys),
			}
			allProcedures = append(allProcedures, info)
			if !arrayProcedureIsParticipant(ctx, proc) {
				continue
			}
			procedures = append(procedures, info)
			moduleArrays[info.key] = arrayModuleNamesForProcedure(file, proc, moduleDecls)
			moduleFiles[info.key] = file.Path
		}
	}
	if len(procedures) == 0 {
		return arrayModuleEntryStates{}
	}
	initializationStates := ctx.arrayModuleInitializationStates
	if initializationStates == nil {
		initializationStates = arrayModuleInitializationStates(files, ctx.arrayModuleAllocations)
	}
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

	evaluate := func(procedure arrayModuleProcedureInfo, entries arrayModuleEntryStates) map[string]map[string]bool {
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
			in = applyArrayModuleStorageSourceGuardState(in, procedure.file, procedure.proc, line, ctx, ctx.arrayModuleStorageGuards[procedure.file.Path])
			forEachArrayCallAtLine(procedure.proc, line, func(call procedureir.CallSite) {
				recordCall(call, in)
			}, ctx.arrayStats)
			out, _ := a.arrayTransfer(procedure.file, procedure.proc, ctx, variables, in, text, line, nil, nil)
			forEachArrayCallAtLine(procedure.proc, line, func(call procedureir.CallSite) {
				out = applyArrayModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
				out = applyArrayUnknownModuleCallEffects(out, procedure.file, procedure.proc, call, ctx, variables, procedure.moduleDecls)
			}, ctx.arrayStats)
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
			out = applyArrayModuleStorageGuardBranch(out, block.Statement, edge, procedure.file, procedure.proc, ctx, ctx.arrayModuleStorageGuards[procedure.file.Path])
			out = applyArrayModulePositiveCountGuardBranch(out, block.Statement, edge, procedure.file, procedure.proc, ctx, procedure.moduleDecls)
			return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[procedure.file.Path], variables, procedure.file, procedure.proc, procedure.moduleDecls)
		}, ctx.arrayStats)
		return candidates
	}

	contributions := make(map[string]map[string]map[string]bool, len(procedures))
	entries := applyArrayModuleCallbackEntryStates(arrayModuleEntryStates{}, procedures, allProcedures, ctx)
	aggregateTargets := func(targets map[string]bool) arrayModuleEntryStates {
		aggregated := arrayModuleEntryStates{}
		for _, caller := range procedures {
			for target, names := range contributions[caller.key] {
				if !targets[target] {
					continue
				}
				if aggregated[target] == nil {
					aggregated[target] = cloneArrayNameSet(names)
					continue
				}
				for name := range aggregated[target] {
					if !names[name] {
						delete(aggregated[target], name)
					}
				}
			}
		}
		return aggregated
	}
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
		oldContribution := contributions[key]
		contributions[key] = contribution
		affectedTargetSet := make(map[string]bool, len(oldContribution)+len(contribution))
		for target := range oldContribution {
			affectedTargetSet[target] = true
		}
		for target := range contribution {
			affectedTargetSet[target] = true
		}
		next := aggregateTargets(affectedTargetSet)
		changedTargets := make([]string, 0, len(affectedTargetSet))
		updates := make(map[string]map[string]bool, len(affectedTargetSet))
		updatePresent := make(map[string]bool, len(affectedTargetSet))
		for target := range affectedTargetSet {
			names, present := next[target]
			oldNames, oldPresent := entries[target]
			if !arrayModuleEntryNamesEqual(oldNames, oldPresent, names, present) {
				changedTargets = append(changedTargets, target)
				updates[target] = names
				updatePresent[target] = present
			}
		}
		if len(changedTargets) == 0 {
			continue
		}
		sortArrayProcedureKeys(changedTargets, indexByKey)
		for _, target := range changedTargets {
			if updatePresent[target] {
				entries[target] = updates[target]
			} else {
				delete(entries, target)
			}
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

type arrayModuleCallbackCaller struct {
	procedure arrayModuleProcedureInfo
	call      procedureir.CallSite
}

// applyArrayModuleCallbackEntryStates binds a module-array allocation proof to
// a procedure installed as a callback. A callback has no ordinary VBA call
// edge, so the normal entry-state fixed point cannot see that it is registered
// only after its owner has initialized the module storage. The proof remains
// source-owned and conservative: every known same-module caller on the path
// to GetAddressOf(AddressOf callback) must have a dominating, successful
// allocation call before the registration call.
func applyArrayModuleCallbackEntryStates(entries arrayModuleEntryStates, procedures, allProcedures []arrayModuleProcedureInfo, ctx analysisContext) arrayModuleEntryStates {
	byKey := make(map[string]arrayModuleProcedureInfo, len(allProcedures))
	callers := make(map[string][]arrayModuleCallbackCaller)
	for _, procedure := range allProcedures {
		byKey[procedure.key] = procedure
	}
	for _, caller := range procedures {
		for call := range caller.proc.Calls.All() {
			_, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, call)
			if !ok {
				continue
			}
			targetKey := arrayParticipantLookupKey(target, ctx.arrayParticipantKeys)
			targetProcedure, ok := byKey[targetKey]
			if !ok || targetProcedure.file.Path != caller.file.Path {
				continue
			}
			callers[targetKey] = append(callers[targetKey], arrayModuleCallbackCaller{procedure: caller, call: call})
		}
	}

	type callbackProof struct {
		arrays map[string]bool
		ok     bool
	}
	memo := map[string]callbackProof{}
	var prove func(string, map[string]bool) callbackProof
	prove = func(key string, visiting map[string]bool) callbackProof {
		if cached, ok := memo[key]; ok {
			return cached
		}
		if visiting[key] {
			return callbackProof{}
		}
		visiting[key] = true
		defer delete(visiting, key)
		paths := callers[key]
		if len(paths) == 0 {
			return callbackProof{}
		}
		var common map[string]bool
		for _, path := range paths {
			arrays := arrayModuleAllocationBeforeCallback(path.procedure, path.call, ctx)
			pathOK := len(arrays) > 0
			if !pathOK {
				pathProof := prove(path.procedure.key, visiting)
				arrays, pathOK = pathProof.arrays, pathProof.ok
			}
			if !pathOK {
				return callbackProof{}
			}
			if common == nil {
				common = cloneArrayNameSet(arrays)
				continue
			}
			for name := range common {
				if !arrays[name] {
					delete(common, name)
				}
			}
		}
		result := callbackProof{arrays: common, ok: len(common) > 0}
		memo[key] = result
		return result
	}

	for _, registration := range procedures {
		for call := range registration.proc.Calls.All() {
			if !strings.EqualFold(call.Callee.BaseName, "GetAddressOf") {
				continue
			}
			targetName := arrayCallbackTargetName(registration.file, call)
			if targetName == "" {
				continue
			}
			proof := prove(registration.key, map[string]bool{})
			if !proof.ok {
				continue
			}
			for _, target := range allProcedures {
				if target.file.Path != registration.file.Path || !strings.EqualFold(target.proc.Name, targetName) {
					continue
				}
				targetModuleArrays := arrayModuleNamesForProcedure(target.file, target.proc, target.moduleDecls)
				for name := range proof.arrays {
					if !targetModuleArrays[name] {
						continue
					}
					if entries[target.key] == nil {
						entries[target.key] = map[string]bool{}
					}
					entries[target.key][name] = true
				}
			}
		}
	}
	return entries
}

func arrayModuleAllocationBeforeCallback(procedure arrayModuleProcedureInfo, registrationCall procedureir.CallSite, ctx analysisContext) map[string]bool {
	allocated := map[string]bool{}
	for candidate := range procedure.proc.Calls.All() {
		if candidate.Range.StartLine >= registrationCall.Range.StartLine || arrayProcedureLineHasInlineConditional(procedure.file, candidate.Range.StartLine) && (!arrayInlineConditionalCallIsConditionExpression(procedure.file, candidate) || !arrayInlineConditionalCallHasSuccessfulExitGuard(procedure.file, candidate)) {
			continue
		}
		key, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, candidate)
		if !ok || !strings.EqualFold(strings.TrimSpace(target.Module), strings.TrimSpace(procedure.proc.Module)) {
			continue
		}
		if !arrayModuleCallDominatesCallback(procedure.proc, candidate, registrationCall) {
			continue
		}
		for name, isAllocated := range ctx.arrayModuleAllocations[key] {
			if isAllocated {
				allocated[strings.ToLower(cleanIdentifier(name))] = true
			}
		}
	}
	return allocated
}

func arrayModuleCallDominatesCallback(proc sourceProcedure, allocationCall, registrationCall procedureir.CallSite) bool {
	if allocationCall.Range.StartLine >= registrationCall.Range.StartLine {
		return false
	}
	if proc.Graph == nil || allocationCall.StatementID <= 0 || registrationCall.StatementID <= 0 {
		return true
	}
	sourceBlock, sourceOK := proc.Graph.BlockForStatement(allocationCall.StatementID)
	targetBlock, targetOK := proc.Graph.BlockForStatement(registrationCall.StatementID)
	if !sourceOK || !targetOK {
		return false
	}
	for _, dominator := range proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).DominatorsOf(targetBlock.ID) {
		if dominator == sourceBlock.ID {
			return true
		}
	}
	return false
}

func arrayInlineConditionalCallHasSuccessfulExitGuard(file parsedFile, call procedureir.CallSite) bool {
	if !arrayProcedureLineHasInlineConditional(file, call.Range.StartLine) {
		return true
	}
	condition, body, ok := arrayIfThenParts(normalizedCodeLine(file.Lines[call.Range.StartLine-1]))
	if !ok || !strings.HasPrefix(strings.ToLower(strings.TrimSpace(body)), "exit ") {
		return false
	}
	condition = strings.TrimSpace(condition)
	if strings.HasPrefix(strings.ToLower(condition), "if ") {
		condition = strings.TrimSpace(condition[len("if "):])
	}
	return strings.Contains(strings.ToLower(condition), "not "+strings.ToLower(cleanIdentifier(call.Callee.BaseName)))
}

func arrayCallbackTargetName(file parsedFile, call procedureir.CallSite) string {
	if call.Range.StartLine < 1 || call.Range.StartLine > len(file.Lines) {
		return ""
	}
	text := strings.ToLower(gui.StripComment(file.Lines[call.Range.StartLine-1]))
	getAddressOf := strings.Index(text, "getaddressof")
	if getAddressOf < 0 {
		return ""
	}
	addressOf := strings.Index(text[getAddressOf+len("getaddressof"):], "addressof")
	if addressOf < 0 {
		return ""
	}
	start := getAddressOf + len("getaddressof") + addressOf + len("addressof")
	for start < len(text) && (text[start] == ' ' || text[start] == '\t' || text[start] == '(') {
		start++
	}
	end := start
	for end < len(text) && isIdentifierPart(text[end]) {
		end++
	}
	return cleanIdentifier(text[start:end])
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

func arrayModuleEntryNamesEqual(leftNames map[string]bool, leftOK bool, rightNames map[string]bool, rightOK bool) bool {
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
			blockMatch := arrayModuleReadyGuardBlockRe.FindStringSubmatch(text)
			if len(blockMatch) != 2 {
				return "", false
			}
			for next := line + 1; next < end; next++ {
				body := strings.TrimSpace(strings.ToLower(normalizedCodeLine(file.Lines[next-1])))
				if body == "" || strings.HasPrefix(body, "'") {
					continue
				}
				if body != "exit sub" && body != "exit function" && body != "exit property" {
					return "", false
				}
				match = blockMatch
				break
			}
			if len(match) != 2 || match[1] == "" {
				return "", false
			}
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
