package analyze

import (
	"maps"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// objectProcedureSummary records the object state a project-local procedure
// guarantees when it returns normally. A ByRef object parameter is useful to
// the caller only when every normal path assigns it a non-Nothing value;
// ByRefWritten separately records whether the callee can overwrite the alias
// at all, so a read-only ByRef call does not erase the caller's existing fact.
// The same rule is used for an object function result; an object-returning
// function with an omitted or Nothing result remains nullable.
type objectProcedureSummary struct {
	QualifiedName string
	File          string
	Module        string
	Kind          string
	Line          int
	Params        []objectParameterSummary
	ByRefAssigned map[int]bool
	ByRefWritten  map[int]bool
	ParamProgID   map[int]string
	// ParamNonNothing records a parameter that is guaranteed non-Nothing when
	// the procedure returns normally.  For ByVal parameters this is retained
	// only when the procedure never writes the local copy, so it represents a
	// precondition proven by a successful call rather than a local assignment.
	ParamNonNothing               map[int]bool
	ModuleAssigned                map[string]bool
	ModuleWritten                 map[string]bool
	ReturnCollection              bool
	ReturnAssigned                bool
	ReturnProgID                  string
	ReturnParameter               int
	ReturnParameterPreservesInput bool
	ReturnCollectionItemParameter int
}

type objectParameterSummary struct {
	Name   string
	Object bool
	ByRef  bool
}

type objectVariable struct {
	Scope procedureir.SymbolScope
	Name  string
}

func (v objectVariable) key() string {
	return strings.ToLower(string(v.Scope) + ":" + cleanIdentifier(v.Name))
}

// objectProcedurePlan contains the parts of one procedure's object-flow
// analysis that do not depend on the current interprocedural summaries or
// entry state.  The plan is owned by one batch analysis run and is never
// shared across runs.
type objectProcedurePlan struct {
	key            string
	file           parsedFile
	proc           sourceProcedure
	moduleDecls    map[string]sourceDeclaration
	declarations   declarationScope
	procedureDecls map[string]sourceDeclaration

	flowProc    sourceProcedure
	flowGraph   vbacfg.CFGView
	flowContext objectFlowContext
	reachable   map[vbacfg.BlockID]bool
	vars        map[string]objectVariable
	// receiverSummaryKeys and classInitializerKeys are immutable indexes built
	// once for the containing batch.  They keep summary lookups inside the
	// fixed-point iterations O(1) instead of rescanning every procedure.
	receiverSummaryKeys  map[string][]string
	classInitializerKeys []string
	classIndexBuilt      bool
	unknownFlow          bool
	relevant             bool
}

type objectEntryCall struct {
	caller  *objectProcedurePlan
	callee  *objectProcedurePlan
	call    procedureir.CallSite
	actuals []objectCallActual

	evaluated     bool
	contributions map[string]bool
}

// objectAnalysisContext is the run-local owner of all VBA202 state.  The
// summary and entry-state worklists update this context in place; no global
// mutable cache is used.
type objectAnalysisContext struct {
	plans map[string]*objectProcedurePlan
	order []string

	summaries map[string]objectProcedureSummary
	entries   map[string]map[string]bool

	summaryDependents   map[string][]string
	entryOutgoing       map[string][]*objectEntryCall
	entryIncoming       map[string][]*objectEntryCall
	moduleProcedureKeys map[string][]string
	callReachable       map[string]bool

	summaryEvaluations   int
	entryFlowEvaluations int
}

func buildObjectAnalysisPlans(files []parsedFile) *objectAnalysisContext {
	analysis := &objectAnalysisContext{
		plans:               map[string]*objectProcedurePlan{},
		summaries:           map[string]objectProcedureSummary{},
		entries:             map[string]map[string]bool{},
		summaryDependents:   map[string][]string{},
		entryOutgoing:       map[string][]*objectEntryCall{},
		entryIncoming:       map[string][]*objectEntryCall{},
		moduleProcedureKeys: map[string][]string{},
	}
	for _, file := range files {
		procedures := file.procedureView()
		moduleDecls := file.moduleDecls()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			key := objectSummaryKey(file.IR.Path, objectProcedureQualifiedName(proc), string(proc.ProcedureKind), proc.StartLine)
			plan := newObjectProcedurePlan(file, proc, moduleDecls, key)
			analysis.plans[key] = plan
			analysis.order = append(analysis.order, key)
		}
	}
	sort.Strings(analysis.order)
	analysis.initializeObjectSummaries()
	analysis.buildObjectIndexes()
	analysis.buildObjectCallReachability()
	analysis.buildObjectDependencies()
	analysis.prepareTerminalCallGraphs()
	return analysis
}

func (analysis *objectAnalysisContext) buildObjectIndexes() {
	receiverKeys := map[string][]string{}
	initializerKeys := map[string][]string{}
	moduleKeys := map[string][]string{}
	for _, key := range analysis.order {
		plan := analysis.plans[key]
		if plan == nil {
			continue
		}
		moduleKey := strings.ToLower(cleanIdentifier(plan.proc.Module))
		moduleKeys[moduleKey] = append(moduleKeys[moduleKey], key)
		receiverIndexKey := objectReceiverSummaryIndexKey(plan.proc.Module, plan.proc.Name)
		receiverKeys[receiverIndexKey] = append(receiverKeys[receiverIndexKey], key)
		if strings.EqualFold(plan.proc.ModuleKind, "class") && strings.EqualFold(plan.proc.Name, "Class_Initialize") {
			moduleKey := strings.ToLower(cleanIdentifier(plan.proc.Module))
			initializerKeys[moduleKey] = append(initializerKeys[moduleKey], key)
		}
	}
	for key, keys := range receiverKeys {
		sort.Strings(keys)
		receiverKeys[key] = uniqueStrings(keys)
	}
	for key, keys := range initializerKeys {
		sort.Strings(keys)
		initializerKeys[key] = uniqueStrings(keys)
	}
	for key, keys := range moduleKeys {
		sort.Strings(keys)
		moduleKeys[key] = uniqueStrings(keys)
	}
	predicateContracts := objectNonNothingPredicateContracts(analysis.plans)
	analysis.moduleProcedureKeys = moduleKeys
	for _, key := range analysis.order {
		plan := analysis.plans[key]
		if plan == nil {
			continue
		}
		plan.receiverSummaryKeys = receiverKeys
		plan.classInitializerKeys = append([]string(nil), initializerKeys[strings.ToLower(cleanIdentifier(plan.proc.Module))]...)
		plan.classIndexBuilt = true
		plan.flowContext.predicateContracts = predicateContracts
	}
}

// buildObjectCallReachability identifies procedures that can contribute an
// entry-state edge to the reachable project call graph.  Object entry states
// are contracts at private procedure boundaries, so an uncalled private
// helper must not poison the state of a shared callee merely because it has a
// nullable ByVal Object parameter.  Public procedures and host-invoked event
// procedures are roots; private procedures become reachable through resolved
// project-local calls from those roots.
func (analysis *objectAnalysisContext) buildObjectCallReachability() {
	analysis.callReachable = map[string]bool{}
	queue := make([]string, 0, len(analysis.order))
	for _, key := range analysis.order {
		plan := analysis.plans[key]
		if !objectProcedureIsCallRoot(plan) {
			continue
		}
		analysis.callReachable[key] = true
		queue = append(queue, key)
	}
	for len(queue) > 0 {
		key := queue[0]
		queue = queue[1:]
		plan := analysis.plans[key]
		if plan == nil {
			continue
		}
		for call := range plan.proc.Calls.All() {
			calleeKey, ok := analysis.objectCalleeKey(call)
			if !ok || analysis.callReachable[calleeKey] {
				continue
			}
			analysis.callReachable[calleeKey] = true
			queue = append(queue, calleeKey)
		}
	}
}

func objectProcedureIsCallRoot(plan *objectProcedurePlan) bool {
	if plan == nil {
		return false
	}
	if !strings.EqualFold(strings.TrimSpace(plan.proc.Visibility), "private") {
		return true
	}
	if plan.proc.IR != nil && plan.proc.IR.Symbol.IsEventHandler {
		return true
	}
	// Bare function calls used in expressions are not always represented as
	// CallSite facts by the VBA IR. An object-returning private function can
	// therefore be an implicit root even when no explicit call edge is
	// available to prove its reachability.
	if isObjectType(plan.proc.ReturnType) {
		return true
	}
	if !strings.EqualFold(strings.TrimSpace(plan.proc.ModuleKind), "class") {
		return false
	}
	return strings.EqualFold(strings.TrimSpace(plan.proc.Name), "Class_Initialize") ||
		strings.EqualFold(strings.TrimSpace(plan.proc.Name), "Class_Terminate")
}

func objectNonNothingPredicateContracts(plans map[string]*objectProcedurePlan) map[string]bool {
	contracts := map[string]bool{}
	qualifiedCounts := map[string]int{}
	qualifiedNames := map[string]bool{}
	nameCounts := map[string]int{}
	contractNames := map[string]bool{}
	for key, plan := range plans {
		if plan != nil {
			name := strings.ToLower(cleanIdentifier(plan.proc.Name))
			if name == "isexceltable" {
				nameCounts[name]++
			}
		}
		if objectProcedureNonNothingPredicate(plan) {
			contracts[key] = true
			qualified := strings.ToLower(objectProcedureQualifiedName(plan.proc))
			if qualified != "" {
				qualifiedCounts[qualified]++
				qualifiedNames[qualified] = true
			}
			contractNames[strings.ToLower(cleanIdentifier(plan.proc.Name))] = true
		}
	}
	for qualified := range qualifiedNames {
		if qualifiedCounts[qualified] == 1 {
			contracts[qualified] = true
		}
	}
	for name := range contractNames {
		if nameCounts[name] == 1 {
			contracts["name:"+name] = true
		}
	}
	return contracts
}

func objectProcedureNonNothingPredicate(plan *objectProcedurePlan) bool {
	if plan == nil || !strings.EqualFold(cleanIdentifier(plan.proc.Name), "isexceltable") || !strings.EqualFold(cleanIdentifier(plan.proc.ReturnType), "boolean") {
		return false
	}
	objectParameter := ""
	objectParameters := 0
	for parameter := range plan.proc.Params.All() {
		if !isObjectType(parameter.Type) {
			continue
		}
		objectParameters++
		objectParameter = parameter.Name
	}
	if objectParameters != 1 || objectParameter == "" || objectPredicateParameterWrites(plan, objectParameter) {
		return false
	}
	return objectPredicateHasNothingExit(plan, objectParameter)
}

func objectPredicateParameterWrites(plan *objectProcedurePlan, parameterName string) bool {
	if plan == nil {
		return true
	}
	for access := range plan.proc.Accesses.All() {
		if access.Scope != procedureir.ScopeParameter ||
			(access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) ||
			!strings.EqualFold(cleanIdentifier(access.Name), cleanIdentifier(parameterName)) {
			continue
		}
		if !objectMemberReceiver(plan.flowContext.facts, access) {
			return true
		}
	}
	for call := range plan.proc.Calls.All() {
		for _, actual := range objectCallActuals(call, plan.flowContext.facts) {
			if actual.parenthesized || !strings.EqualFold(cleanIdentifier(actual.text), cleanIdentifier(parameterName)) {
				continue
			}
			callee := cleanIdentifier(call.Callee.BaseName)
			if callee == "" {
				callee = objectBareCallName(call.Callee.Text)
			}
			if strings.EqualFold(callee, "typename") {
				continue
			}
			return true
		}
	}
	return false
}

func objectPredicateHasNothingExit(plan *objectProcedurePlan, parameterName string) bool {
	if plan == nil || plan.flowGraph.BlockCount() == 0 {
		return false
	}
	variable := objectVariable{Scope: procedureir.ScopeParameter, Name: parameterName}
	flow := objectPredicateNonNothingFlow(plan, variable)
	sawPotentialTrue := false
	for statement := range plan.proc.Statements.All() {
		if (statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet) || statement.Value == nil ||
			!objectPredicateWritesResult(statement, plan.proc.Name) {
			continue
		}
		value := strings.ToLower(strings.TrimSpace(statement.Value.Text))
		if value == "false" || value == "0" || value == "vbfalse" {
			continue
		}
		sawPotentialTrue = true
		block, ok := plan.flowGraph.BlockForStatement(statement.ID)
		if !ok || !flow[block.ID] {
			// A predicate can establish the contract only if every path that
			// can produce True has already proven the object non-Nothing.
			return false
		}
	}
	return sawPotentialTrue
}

func objectPredicateWritesResult(statement procedureir.Statement, procedureName string) bool {
	if statement.Target != nil && strings.EqualFold(cleanIdentifier(statement.Target.Text), cleanIdentifier(procedureName)) {
		return true
	}
	// Some Boolean expressions are canonicalized with the first nested call as
	// Statement.Target. The source-level assignment target remains unambiguous
	// at the beginning of the statement, so use it as the fallback binding.
	text := strings.TrimSpace(statement.Text)
	separator := strings.IndexByte(text, '=')
	if separator < 0 {
		return false
	}
	left := strings.TrimSpace(text[:separator])
	for _, prefix := range []string{"set ", "let "} {
		if len(left) >= len(prefix) && strings.EqualFold(left[:len(prefix)], prefix) {
			left = strings.TrimSpace(left[len(prefix):])
			break
		}
	}
	return strings.EqualFold(cleanIdentifier(left), cleanIdentifier(procedureName))
}

func objectPredicateNonNothingFlow(plan *objectProcedurePlan, variable objectVariable) map[vbacfg.BlockID]bool {
	result := map[vbacfg.BlockID]bool{}
	if plan == nil {
		return result
	}
	graph := plan.flowGraph
	reachable := map[vbacfg.BlockID]bool{}
	for _, id := range graph.Reachable() {
		reachable[id] = true
	}
	for _, id := range graph.Reachable() {
		result[id] = id != graph.Entry()
	}
	result[graph.Entry()] = false
	changed := true
	for changed {
		changed = false
		graph.ForEachBlock(func(block vbacfg.Block) bool {
			if !reachable[block.ID] || block.ID == graph.Entry() {
				return true
			}
			known := true
			seen := false
			graph.ForEachIncoming(block.ID, func(edge vbacfg.Edge) bool {
				if !reachable[edge.From] {
					return true
				}
				// The predicate contract describes normal function results. An
				// exceptional edge from the guard is not a normal path that can
				// assign the predicate result, so it must not manufacture a
				// false path at the assignment block.
				if edge.Class == vbacfg.EdgeExceptional {
					return true
				}
				state := result[edge.From]
				if edge.Uncertain {
					state = result[edge.From]
				}
				state = objectPredicateApplyGuard(plan, variable, edge, state)
				known = known && state
				seen = true
				return true
			})
			if seen && result[block.ID] != known {
				result[block.ID] = known
				changed = true
			}
			return true
		})
	}
	return result
}

func objectPredicateApplyGuard(plan *objectProcedurePlan, variable objectVariable, edge vbacfg.Edge, state bool) bool {
	if plan == nil || (edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse) || plan.flowContext.facts == nil {
		return state
	}
	statement, ok := plan.flowContext.facts.Statement(edge.StatementID)
	if !ok || statement.Condition == nil {
		return state
	}
	text := objectTrimOuterParens(strings.ToLower(strings.TrimSpace(statement.Condition.Text)))
	negated := false
	if strings.HasPrefix(text, "not ") {
		negated = true
		text = objectTrimOuterParens(strings.TrimSpace(strings.TrimPrefix(text, "not ")))
	}
	want := strings.ToLower(cleanIdentifier(variable.Name)) + " is nothing"
	if text != want {
		return state
	}
	provesNonNothing := edge.Kind == vbacfg.EdgeBranchFalse
	if negated {
		provesNonNothing = !provesNonNothing
	}
	if provesNonNothing {
		return true
	}
	return false
}

func newObjectProcedurePlan(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, key string) *objectProcedurePlan {
	plan := &objectProcedurePlan{
		key:            key,
		file:           file,
		proc:           proc,
		moduleDecls:    moduleDecls,
		declarations:   objectFlowDeclarations(file, proc, moduleDecls),
		procedureDecls: file.procedureDeclarationsFor(proc),
		reachable:      map[vbacfg.BlockID]bool{},
		vars:           map[string]objectVariable{},
	}
	plan.flowProc = proc
	if plan.proc.Graph != nil {
		plan.flowGraph = proc.Graph.WithoutNormalErrRaiseContinuationView()
		for _, id := range plan.flowGraph.Reachable() {
			plan.reachable[id] = true
		}
		plan.flowContext = newObjectFlowContext(plan.flowProc, plan.flowGraph)
	}
	addObjectVariables := func(scope procedureir.SymbolScope, declarations map[string]sourceDeclaration) {
		for _, declaration := range declarations {
			if !declaration.Object {
				continue
			}
			variable := objectVariable{Scope: scope, Name: declaration.Name}
			plan.vars[variable.key()] = variable
		}
	}
	// Keep the module layer shared. Local and parameter layers are small and
	// are enumerated independently so shadowed bindings still get distinct
	// object-flow variables without copying the module declaration map.
	addObjectVariables(procedureir.ScopeModule, plan.declarations.module)
	addObjectVariables(procedureir.ScopeLocal, plan.declarations.extra)
	addObjectVariables(procedureir.ScopeLocal, plan.declarations.local)
	addObjectVariables(procedureir.ScopeParameter, plan.declarations.parameters)
	if isObjectType(proc.ReturnType) {
		variable := objectVariable{Scope: procedureir.ScopeLocal, Name: proc.Name}
		plan.vars[variable.key()] = variable
	}
	plan.flowContext.vars = plan.vars
	plan.unknownFlow = proc.Graph == nil || len(proc.Graph.UnknownFlowSources) > 0
	for statement := range proc.Statements.All() {
		plan.unknownFlow = plan.unknownFlow || statement.Recovered
	}
	plan.relevant = objectProcedureRelevant(proc, moduleDecls)
	if !plan.relevant {
		for _, declaration := range plan.procedureDecls {
			if declaration.Object {
				plan.relevant = true
				break
			}
		}
	}
	if !plan.relevant {
		for parameter := range proc.Params.All() {
			if isObjectType(parameter.Type) {
				plan.relevant = true
				break
			}
		}
	}
	if !plan.relevant {
		plan.relevant = objectProcedureUsesModuleObject(proc, moduleDecls)
	}
	return plan
}

func (analysis *objectAnalysisContext) initializeObjectSummaries() {
	for _, key := range analysis.order {
		plan := analysis.plans[key]
		summary := objectProcedureSummary{
			QualifiedName:                 objectProcedureQualifiedName(plan.proc),
			File:                          plan.file.IR.Path,
			Module:                        plan.proc.Module,
			Kind:                          string(plan.proc.ProcedureKind),
			Line:                          plan.proc.StartLine,
			ByRefAssigned:                 map[int]bool{},
			ByRefWritten:                  map[int]bool{},
			ParamProgID:                   map[int]string{},
			ParamNonNothing:               map[int]bool{},
			ModuleAssigned:                map[string]bool{},
			ModuleWritten:                 map[string]bool{},
			ReturnCollection:              dcKindFromType(plan.proc.ReturnType) == dcCollection,
			ReturnParameter:               -1,
			ReturnCollectionItemParameter: -1,
		}
		for index, parameter := range plan.proc.Params.AllIndexed() {
			summary.Params = append(summary.Params, objectParameterSummary{
				Name:   parameter.Name,
				Object: isObjectType(parameter.Type),
				ByRef:  !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal"),
			})
			if summary.Params[index].Object && summary.Params[index].ByRef {
				summary.ByRefAssigned[index] = false
				summary.ByRefWritten[index] = false
			}
		}
		analysis.summaries[key] = summary
		analysis.entries[key] = objectInitialEntryState(plan)
	}
}

func objectInitialEntryState(plan *objectProcedurePlan) map[string]bool {
	state := map[string]bool{}
	for name, declaration := range plan.moduleDecls {
		if declaration.Object {
			state[(objectVariable{Scope: procedureir.ScopeModule, Name: name}).key()] = false
		}
	}
	for parameter := range plan.proc.Params.All() {
		if isObjectType(parameter.Type) {
			state[(objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}).key()] = false
		}
	}
	return state
}

func (analysis *objectAnalysisContext) buildObjectDependencies() {
	addSummaryDependency := func(calleeKey, callerKey string) {
		analysis.summaryDependents[calleeKey] = append(analysis.summaryDependents[calleeKey], callerKey)
	}
	for _, callerKey := range analysis.order {
		caller := analysis.plans[callerKey]
		for call := range caller.proc.Calls.All() {
			calleeKey, ok := analysis.objectCalleeKey(call)
			if ok {
				addSummaryDependency(calleeKey, callerKey)
			}
			for _, receiverKey := range analysis.objectReceiverCalleeKeys(caller, call) {
				addSummaryDependency(receiverKey, callerKey)
			}
			if !ok || calleeKey == callerKey || !analysis.callReachable[callerKey] {
				// Recursive entry-state calls are deliberately excluded. Summary
				// propagation keeps the self edge so recursive summaries converge.
				// Calls from unreachable private helpers likewise cannot establish a
				// runtime entry contract for the shared callee.
				continue
			}
			callee := analysis.plans[calleeKey]
			if callee == nil || caller.proc.Graph == nil {
				continue
			}
			block, found := caller.proc.Graph.BlockForStatement(call.StatementID)
			if !found || !caller.reachable[block.ID] {
				continue
			}
			entryCall := &objectEntryCall{
				caller:        caller,
				callee:        callee,
				call:          call,
				actuals:       objectCallActuals(call, caller.flowContext.facts),
				contributions: map[string]bool{},
			}
			analysis.entryOutgoing[callerKey] = append(analysis.entryOutgoing[callerKey], entryCall)
			analysis.entryIncoming[calleeKey] = append(analysis.entryIncoming[calleeKey], entryCall)
			calleeSummary := analysis.summaries[calleeKey]
			for actualIndex, actual := range entryCall.actuals {
				formalIndex := objectFormalIndex(entryCall.call, calleeSummary, actualIndex)
				if formalIndex < 0 || formalIndex >= len(calleeSummary.Params) || !calleeSummary.Params[formalIndex].Object {
					continue
				}
				statementIDs := map[int]bool{entryCall.call.StatementID: true}
				name := cleanIdentifier(actual.text)
				if name != "" {
					for statement := range caller.proc.Statements.All() {
						if statement.Kind == procedureir.StatementSet && statement.Target != nil && strings.EqualFold(cleanIdentifier(statement.Target.Text), name) {
							statementIDs[statement.ID] = true
						}
					}
				}
				for nestedCall := range caller.proc.Calls.All() {
					if !statementIDs[nestedCall.StatementID] {
						continue
					}
					factoryKey, factoryOK := analysis.objectCalleeKey(nestedCall)
					if factoryOK && factoryKey != calleeKey {
						addSummaryDependency(factoryKey, calleeKey)
					}
				}
			}
		}
	}
	for _, initializerKey := range analysis.order {
		initializer := analysis.plans[initializerKey]
		if initializer == nil || !strings.EqualFold(initializer.proc.ModuleKind, "class") || !strings.EqualFold(initializer.proc.Name, "Class_Initialize") {
			continue
		}
		moduleKey := strings.ToLower(cleanIdentifier(initializer.proc.Module))
		for _, dependentKey := range analysis.moduleProcedureKeys[moduleKey] {
			dependent := analysis.plans[dependentKey]
			if dependentKey != initializerKey && dependent != nil && strings.EqualFold(dependent.proc.Module, initializer.proc.Module) {
				addSummaryDependency(initializerKey, dependentKey)
			}
		}
	}
	for key, dependents := range analysis.summaryDependents {
		sort.Strings(dependents)
		analysis.summaryDependents[key] = uniqueStrings(dependents)
	}
	for key, calls := range analysis.entryOutgoing {
		sort.SliceStable(calls, func(i, j int) bool {
			if calls[i].callee.key != calls[j].callee.key {
				return calls[i].callee.key < calls[j].callee.key
			}
			return calls[i].call.ID < calls[j].call.ID
		})
		analysis.entryOutgoing[key] = calls
	}
	for key, calls := range analysis.entryIncoming {
		sort.SliceStable(calls, func(i, j int) bool {
			if calls[i].caller.key != calls[j].caller.key {
				return calls[i].caller.key < calls[j].caller.key
			}
			return calls[i].call.ID < calls[j].call.ID
		})
		analysis.entryIncoming[key] = calls
	}
}

// prepareTerminalCallGraphs removes normal-flow continuations after
// project-local helpers that cannot return normally.  A helper such as
// RaiseContractError is otherwise treated as an ordinary call, so the
// impossible fall-through edge keeps the pre-guard Nothing state alive in the
// caller and poisons both object use findings and return summaries.
func (analysis *objectAnalysisContext) prepareTerminalCallGraphs() {
	terminal := map[string]bool{}
	changed := true
	for changed {
		changed = false
		for _, key := range analysis.order {
			if terminal[key] {
				continue
			}
			plan := analysis.plans[key]
			if plan == nil || !analysis.objectPlanAlwaysTerminates(plan, terminal) {
				continue
			}
			terminal[key] = true
			changed = true
		}
	}

	for _, plan := range analysis.plans {
		if plan == nil || plan.proc.Graph == nil {
			continue
		}
		removed := map[vbacfg.BlockID]bool{}
		terminalCalls := cloneIntBoolMap(plan.flowContext.terminalCalls)
		for call := range plan.proc.Calls.All() {
			calleeKey, ok := analysis.objectCalleeKey(call)
			if !ok || !terminal[calleeKey] {
				continue
			}
			terminalCalls[call.ID] = true
			block, ok := plan.flowGraph.BlockForStatement(call.StatementID)
			if ok {
				removed[block.ID] = true
			}
		}
		if len(removed) == 0 {
			continue
		}
		plan.flowGraph = plan.flowGraph.WithoutNormalContinuationsFrom(removed)
		plan.reachable = map[vbacfg.BlockID]bool{}
		for _, id := range plan.flowGraph.Reachable() {
			plan.reachable[id] = true
		}
		receiverSummaryKeys := plan.flowContext.receiverSummaryKeys
		predicateContracts := plan.flowContext.predicateContracts
		plan.flowContext = newObjectFlowContext(plan.flowProc, plan.flowGraph)
		plan.flowContext.vars = plan.vars
		plan.flowContext.receiverSummaryKeys = receiverSummaryKeys
		plan.flowContext.predicateContracts = predicateContracts
		plan.flowContext.terminalCalls = terminalCalls
	}
}

func (analysis *objectAnalysisContext) objectPlanAlwaysTerminates(plan *objectProcedurePlan, terminal map[string]bool) bool {
	if plan == nil || plan.proc.Graph == nil {
		return false
	}
	directRaise := objectPlanHasDirectRaise(plan)
	if len(terminal) == 0 && !directRaise {
		return false
	}
	graph := plan.proc.Graph.WithoutNormalErrRaiseContinuationView()
	removed := map[vbacfg.BlockID]bool{}
	terminalCall := false
	for call := range plan.proc.Calls.All() {
		calleeKey, ok := analysis.objectCalleeKey(call)
		if !ok || !terminal[calleeKey] {
			continue
		}
		terminalCall = true
		block, ok := graph.BlockForStatement(call.StatementID)
		if ok {
			removed[block.ID] = true
		}
	}
	graph = graph.WithoutNormalContinuationsFrom(removed)
	if graph.IsReachable(graph.NormalExit()) || graph.IsReachable(graph.UnknownExit()) {
		return false
	}
	return terminalCall || directRaise
}

func objectPlanHasDirectRaise(plan *objectProcedurePlan) bool {
	if plan == nil {
		return false
	}
	for statement := range plan.proc.Statements.All() {
		text := strings.ToLower(strings.TrimSpace(statement.Text))
		if strings.HasPrefix(text, "err.raise") || strings.HasPrefix(text, "call err.raise") || strings.HasPrefix(text, "error ") {
			return true
		}
	}
	return false
}

func (analysis *objectAnalysisContext) objectReceiverCalleeKeys(caller *objectProcedurePlan, call procedureir.CallSite) []string {
	if caller == nil || call.Callee.Receiver == nil || call.Callee.Member == "" {
		return nil
	}
	receiver := strings.TrimSpace(*call.Callee.Receiver)
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(strings.SplitN(receiver, ".", 2)[0], "(", 2)[0]))
	declaration, ok := objectDeclarationByName(root, caller.declarations)
	if !ok || strings.TrimSpace(declaration.Type) == "" {
		return nil
	}
	typeName := strings.ToLower(cleanIdentifier(lastName(strings.TrimSpace(declaration.Type))))
	member := strings.ToLower(cleanIdentifier(call.Callee.Member))
	return append([]string(nil), caller.receiverSummaryKeys[objectReceiverSummaryIndexKey(typeName, member)]...)
}

func objectReceiverSummaryIndexKey(module, member string) string {
	return strings.ToLower(cleanIdentifier(module)) + "|" + strings.ToLower(cleanIdentifier(lastName(member)))
}

// uniqueStrings expects values to be sorted and removes adjacent duplicates
// in place. Callers must provide an owned slice because the input backing
// array is reused.
func uniqueStrings(values []string) []string {
	if len(values) < 2 {
		return values
	}
	out := values[:0]
	for _, value := range values {
		if len(out) == 0 || out[len(out)-1] != value {
			out = append(out, value)
		}
	}
	return out
}

func (analysis *objectAnalysisContext) objectCalleeKey(call procedureir.CallSite) (string, bool) {
	if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
		candidate := call.Resolution.Candidates[0]
		key := objectSummaryKey(candidate.File, candidate.QualifiedName, candidate.Kind, candidate.Line)
		if _, ok := analysis.plans[key]; ok {
			return key, true
		}
		var match string
		for candidateKey, plan := range analysis.plans {
			if !strings.EqualFold(plan.proc.Module+"."+plan.proc.Name, candidate.QualifiedName) && !strings.EqualFold(objectProcedureQualifiedName(plan.proc), candidate.QualifiedName) {
				continue
			}
			if candidate.Kind != "" && !strings.EqualFold(string(plan.proc.ProcedureKind), candidate.Kind) {
				continue
			}
			if candidate.Line > 0 && plan.proc.StartLine != candidate.Line {
				continue
			}
			if match != "" {
				return "", false
			}
			match = candidateKey
		}
		return match, match != ""
	}
	if call.Callee.Receiver != nil && !objectIsCurrentModuleReceiver(call) {
		return "", false
	}
	if !objectResolutionAllowsCalleeFallback(call.Resolution.Status) {
		return "", false
	}
	name := strings.TrimSpace(call.Callee.BaseName)
	if name == "" {
		name = strings.TrimSpace(call.Callee.Member)
	}
	if name == "" {
		name = objectBareCallName(call.Callee.Text)
	}
	if name == "" {
		return "", false
	}
	var match string
	for candidateKey, plan := range analysis.plans {
		if strings.TrimSpace(call.Module) != "" && !strings.EqualFold(strings.TrimSpace(plan.proc.Module), strings.TrimSpace(call.Module)) {
			continue
		}
		if !strings.EqualFold(cleanIdentifier(plan.proc.Name), cleanIdentifier(name)) {
			continue
		}
		if match != "" {
			return "", false
		}
		match = candidateKey
	}
	return match, match != ""
}

func objectResolutionAllowsDirectFallback(status procedureir.ResolutionStatus) bool {
	switch status {
	case "": // Synthetic CallSite values use the zero value for not attempted.
		return true
	case procedureir.ResolutionNotAttempted, procedureir.ResolutionUnresolved, procedureir.ResolutionIncomplete:
		return true
	default:
		return false
	}
}

func objectResolutionAllowsDirectSummary(status procedureir.ResolutionStatus) bool {
	return objectResolutionAllowsDirectFallback(status) || status == procedureir.ResolutionAmbiguous
}

func objectResolutionAllowsCalleeFallback(status procedureir.ResolutionStatus) bool {
	return objectResolutionAllowsDirectFallback(status) || status == procedureir.ResolutionAmbiguous
}

func objectBareCallName(text string) string {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(text), "call ") {
		text = strings.TrimSpace(text[len("call "):])
	}
	if text == "" {
		return ""
	}
	if dot := strings.IndexByte(text, '.'); dot >= 0 {
		text = text[:dot]
	}
	if open := strings.IndexByte(text, '('); open >= 0 {
		text = text[:open]
	}
	if space := strings.IndexAny(text, " \t"); space >= 0 {
		text = text[:space]
	}
	return cleanIdentifier(text)
}

func objectIsCurrentModuleReceiver(call procedureir.CallSite) bool {
	return call.Callee.Receiver != nil && strings.EqualFold(cleanIdentifier(strings.TrimSpace(*call.Callee.Receiver)), "me")
}

func (analysis *objectAnalysisContext) buildSummaries() map[string]objectProcedureSummary {
	for {
		queue := append([]string(nil), analysis.order...)
		queued := make(map[string]bool, len(queue))
		for _, key := range queue {
			queued[key] = true
		}
		for len(queue) > 0 {
			key := queue[0]
			queue = queue[1:]
			queued[key] = false
			plan := analysis.plans[key]
			if plan == nil || !plan.relevant {
				continue
			}
			analysis.summaryEvaluations++
			previous := analysis.summaries[key]
			flow := objectStateFlowPlan(plan, analysis.summaries, objectFlowOptions{})
			updated := previous
			updated.ByRefAssigned = cloneIntBoolMap(previous.ByRefAssigned)
			updated.ByRefWritten = cloneIntBoolMap(previous.ByRefWritten)
			updated.ParamProgID = analysis.objectParameterProgIDs(plan)
			updated.ParamNonNothing = cloneIntBoolMap(previous.ParamNonNothing)
			updated.ModuleAssigned = cloneBoolMap(previous.ModuleAssigned)
			updated.ModuleWritten = cloneBoolMap(previous.ModuleWritten)
			for index, parameter := range previous.Params {
				if !parameter.Object || !parameter.ByRef {
					continue
				}
				variable := objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}
				updated.ByRefAssigned[index] = !plan.unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
				updated.ByRefWritten[index] = plan.unknownFlow || objectProcedureWritesParameterIndexed(plan.proc, parameter.Name, analysis.summaries, plan.flowContext.facts)
				updated.ParamNonNothing[index] = !plan.unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
			}
			for index, parameter := range previous.Params {
				if !parameter.Object || parameter.ByRef {
					continue
				}
				variable := objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}
				updated.ParamNonNothing[index] = !plan.unknownFlow &&
					!objectProcedureWritesParameterIndexed(plan.proc, parameter.Name, analysis.summaries, plan.flowContext.facts) &&
					objectFlowExitDefinitelyAssigned(flow, variable)
			}
			for name, declaration := range plan.moduleDecls {
				if !declaration.Object {
					continue
				}
				variable := objectVariable{Scope: procedureir.ScopeModule, Name: declaration.Name}
				updated.ModuleAssigned[strings.ToLower(name)] = !plan.unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
				updated.ModuleWritten[strings.ToLower(name)] = objectProcedureWritesModuleFieldIndexed(plan.proc, declaration.Name, plan.flowContext.facts)
			}
			if isObjectType(plan.proc.ReturnType) {
				variable := objectVariable{Scope: procedureir.ScopeLocal, Name: plan.proc.Name}
				updated.ReturnAssigned = !plan.unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
			}
			updated.ReturnCollectionItemParameter = objectReturnCollectionItemParameter(plan)
			updated.ReturnParameter = objectReturnParameter(plan)
			updated.ReturnParameterPreservesInput = objectReturnParameterPreservesInput(plan, updated.ReturnParameter, analysis.summaries)
			updated.ReturnProgID = objectReturnProgID(plan, analysis.summaries)
			if !objectSummaryEqual(previous, updated) {
				analysis.summaries[key] = updated
				for _, dependent := range analysis.summaryDependents[key] {
					if !queued[dependent] {
						queue = append(queue, dependent)
						queued[dependent] = true
					}
				}
				if previous.ReturnAssigned != updated.ReturnAssigned && (previous.ReturnCollection || updated.ReturnCollection) {
					// Bare object-returning function calls are represented as
					// identifier expressions by the VBA IR and therefore have no
					// explicit summary dependency edge. Revisit the full object
					// worklist when a return contract changes so those implicit
					// callers observe the new fixed-point value as well.
					for _, dependent := range analysis.order {
						if !queued[dependent] {
							queue = append(queue, dependent)
							queued[dependent] = true
						}
					}
				}
			}
		}
		if !analysis.seedRecursiveCollectionSummaries() {
			return analysis.summaries
		}
	}
}

func (analysis *objectAnalysisContext) seedRecursiveCollectionSummaries() bool {
	optimistic := maps.Clone(analysis.summaries)
	for key, summary := range optimistic {
		if summary.ReturnCollection {
			summary.ReturnAssigned = true
			optimistic[key] = summary
		}
	}
	seeds := make([]string, 0)
	for _, key := range analysis.order {
		plan := analysis.plans[key]
		summary := analysis.summaries[key]
		if plan == nil || !plan.relevant || !summary.ReturnCollection || summary.ReturnAssigned || plan.unknownFlow || !objectPlanHasCollectionAllocation(plan) {
			continue
		}
		flow := objectStateFlowPlan(plan, optimistic, objectFlowOptions{})
		variable := objectVariable{Scope: procedureir.ScopeLocal, Name: plan.proc.Name}
		if objectFlowExitDefinitelyAssigned(flow, variable) {
			seeds = append(seeds, key)
		}
	}
	for _, key := range seeds {
		summary := analysis.summaries[key]
		summary.ReturnAssigned = true
		analysis.summaries[key] = summary
	}
	return len(seeds) > 0
}

func objectPlanHasCollectionAllocation(plan *objectProcedurePlan) bool {
	if plan == nil {
		return false
	}
	for expression := range plan.proc.Expressions.All() {
		if expression.Kind != procedureir.ExpressionNew {
			continue
		}
		text := strings.TrimSpace(expression.Text)
		newType, ok := strings.CutPrefix(strings.ToLower(text), "new ")
		if ok && dcKindFromType(strings.TrimSpace(newType)) == dcCollection {
			return true
		}
	}
	return false
}

func (analysis *objectAnalysisContext) objectParameterProgIDs(plan *objectProcedurePlan) map[int]string {
	result := map[int]string{}
	if plan == nil || !objectProcedureAllowsParameterEntry(plan.proc) {
		return result
	}
	incoming := analysis.entryIncoming[plan.key]
	if len(incoming) == 0 {
		return result
	}
	calleeSummary := analysis.summaries[plan.key]
	for parameterIndex, parameter := range calleeSummary.Params {
		if !parameter.Object {
			continue
		}
		candidate := ""
		valid := true
		for _, entryCall := range incoming {
			foundActual := false
			progid := ""
			for actualIndex, actual := range entryCall.actuals {
				if objectFormalIndex(entryCall.call, calleeSummary, actualIndex) != parameterIndex {
					continue
				}
				foundActual = true
				progid = objectActualFactoryProgID(entryCall.caller, entryCall.call, actual, analysis.summaries)
				break
			}
			if !foundActual || progid == "" {
				valid = false
				break
			}
			if candidate == "" {
				candidate = progid
				continue
			}
			if !strings.EqualFold(candidate, progid) {
				valid = false
				break
			}
		}
		if valid && candidate != "" {
			result[parameterIndex] = candidate
		}
	}
	return result
}

func objectActualFactoryProgID(caller *objectProcedurePlan, call procedureir.CallSite, actual objectCallActual, summaries map[string]objectProcedureSummary) string {
	if caller == nil {
		return ""
	}
	text := strings.TrimSpace(actual.text)
	if text == "" {
		return ""
	}
	if progid := objectCreateObjectProgID(text, caller.file.ConstantValues); progid != "" {
		return progid
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "new ") {
		switch dcKindFromType(strings.TrimSpace(text[len("new "):])) {
		case dcDictionary:
			return "scripting.dictionary"
		case dcCollection:
			return "vba.collection"
		}
	}
	name := cleanIdentifier(text)
	if declaration, _, ok := objectDeclarationBinding(name, caller.declarations); ok {
		switch dcKindFromType(declaration.Type) {
		case dcDictionary:
			return "scripting.dictionary"
		case dcCollection:
			return "vba.collection"
		}
	}
	return strings.ToLower(objectReceiverFactoryProgID(caller.proc, name, call.StatementID, summaries))
}

func (analysis *objectAnalysisContext) buildEntryStates() map[string]map[string]bool {
	queue := append([]string(nil), analysis.order...)
	queued := make(map[string]bool, len(queue))
	for _, key := range queue {
		queued[key] = true
	}
	for len(queue) > 0 {
		key := queue[0]
		queue = queue[1:]
		queued[key] = false
		caller := analysis.plans[key]
		if caller == nil {
			continue
		}
		analysis.entryFlowEvaluations++
		flow := objectStateFlowPlan(caller, analysis.summaries, objectFlowOptions{Entry: analysis.entries[key]})
		flowContext := caller.flowContext
		flowContext.receiverSummaryKeys = caller.receiverSummaryKeys
		for _, entryCall := range analysis.entryOutgoing[key] {
			contributions := map[string]bool{}
			callee := entryCall.callee
			state := flow.in
			if caller.proc.Graph == nil {
				entryCall.evaluated = true
				entryCall.contributions = map[string]bool{}
				continue
			}
			block, ok := caller.proc.Graph.BlockForStatement(entryCall.call.StatementID)
			if !ok || !caller.reachable[block.ID] {
				entryCall.evaluated = true
				entryCall.contributions = map[string]bool{}
				continue
			}
			blockState := state[block.ID]
			calleeSummary := analysis.summaries[callee.key]
			if objectProcedureAllowsParameterEntry(callee.proc) {
				for index, parameter := range callee.proc.Params.AllIndexed() {
					if !isObjectType(parameter.Type) {
						continue
					}
					parameterKey := (objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}).key()
					assigned, present := objectCallParameterAssigned(caller.proc, caller.declarations, entryCall.call, calleeSummary, index, entryCall.actuals, blockState, flow.vars, flowContext, analysis.summaries)
					contributions[parameterKey] = assigned && present
				}
			}
			if strings.EqualFold(callee.proc.Module, caller.proc.Module) {
				for name, declaration := range callee.moduleDecls {
					if !declaration.Object {
						continue
					}
					binding, scope, ok := objectDeclarationBinding(name, caller.declarations)
					if !ok || scope != procedureir.ScopeModule || !binding.Object {
						continue
					}
					fieldKey := (objectVariable{Scope: procedureir.ScopeModule, Name: name}).key()
					contributions[fieldKey] = blockState[fieldKey]
				}
			}
			if !entryCall.evaluated || !objectBoolMapEqual(entryCall.contributions, contributions) {
				entryCall.evaluated = true
				entryCall.contributions = contributions
				if analysis.recomputeEntry(entryCall.callee.key) && !queued[entryCall.callee.key] {
					queue = append(queue, entryCall.callee.key)
					queued[entryCall.callee.key] = true
				}
			}
		}
	}
	return analysis.entries
}

func (analysis *objectAnalysisContext) recomputeEntry(key string) bool {
	plan := analysis.plans[key]
	if plan == nil {
		return false
	}
	next := objectInitialEntryState(plan)
	incoming := analysis.entryIncoming[key]
	for fieldKey := range next {
		seen := false
		value := true
		for _, call := range incoming {
			if !call.evaluated {
				// An unprocessed incoming call is conservatively treated as a
				// failing contribution. Once it is evaluated, an inapplicable
				// field is removed from the meet by the `present` check below.
				seen = true
				value = false
				break
			}
			contribution, present := call.contributions[fieldKey]
			if !present {
				continue
			}
			seen = true
			if !contribution {
				value = false
				break
			}
		}
		if seen {
			next[fieldKey] = value
		}
	}
	if objectBoolMapEqual(analysis.entries[key], next) {
		return false
	}
	analysis.entries[key] = next
	return true
}

func objectBoolMapEqual(a, b map[string]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for key, value := range a {
		if b[key] != value {
			return false
		}
	}
	return true
}

func objectSummaryEqual(a, b objectProcedureSummary) bool {
	if a.ReturnAssigned != b.ReturnAssigned || a.ReturnProgID != b.ReturnProgID || a.ReturnParameter != b.ReturnParameter || a.ReturnParameterPreservesInput != b.ReturnParameterPreservesInput || a.ReturnCollectionItemParameter != b.ReturnCollectionItemParameter || len(a.ByRefAssigned) != len(b.ByRefAssigned) || len(a.ByRefWritten) != len(b.ByRefWritten) || len(a.ParamProgID) != len(b.ParamProgID) || len(a.ParamNonNothing) != len(b.ParamNonNothing) || len(a.ModuleAssigned) != len(b.ModuleAssigned) || len(a.ModuleWritten) != len(b.ModuleWritten) {
		return false
	}
	for index, value := range a.ByRefAssigned {
		if b.ByRefAssigned[index] != value {
			return false
		}
	}
	for index, value := range a.ByRefWritten {
		if b.ByRefWritten[index] != value {
			return false
		}
	}
	for index, value := range a.ParamProgID {
		if !strings.EqualFold(b.ParamProgID[index], value) {
			return false
		}
	}
	for index, value := range a.ParamNonNothing {
		if b.ParamNonNothing[index] != value {
			return false
		}
	}
	for name, value := range a.ModuleAssigned {
		if b.ModuleAssigned[name] != value {
			return false
		}
	}
	for name, value := range a.ModuleWritten {
		if b.ModuleWritten[name] != value {
			return false
		}
	}
	return true
}

func cloneBoolMap(in map[string]bool) map[string]bool {
	if in == nil {
		return map[string]bool{}
	}
	out := make(map[string]bool, len(in))
	for key, value := range in {
		out[key] = value
	}
	return out
}

func cloneIntBoolMap(in map[int]bool) map[int]bool {
	if in == nil {
		return map[int]bool{}
	}
	out := make(map[int]bool, len(in))
	for key, value := range in {
		out[key] = value
	}
	return out
}

func objectProcedureQualifiedName(proc sourceProcedure) string {
	if proc.Module == "" {
		return proc.Name
	}
	return proc.Module + "." + proc.Name
}

func objectSummaryKey(file, qualified, kind string, line int) string {
	return strings.ToLower(strings.TrimSpace(filepath.Clean(file)) + "|" + strings.TrimSpace(qualified) + "|" + strings.TrimSpace(kind) + "|" + strconv.Itoa(line))
}

func objectProcedureRelevant(proc sourceProcedure, moduleDecls map[string]sourceDeclaration) bool {
	if isObjectType(proc.ReturnType) {
		return true
	}
	for parameter := range proc.Params.All() {
		if isObjectType(parameter.Type) {
			return true
		}
	}
	return objectProcedureUsesModuleObject(proc, moduleDecls)
}

func objectProcedureUsesModuleObject(proc sourceProcedure, moduleDecls map[string]sourceDeclaration) bool {
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeModule {
			continue
		}
		name := strings.ToLower(cleanIdentifier(access.Name))
		if declaration, ok := moduleDecls[name]; ok && declaration.Object {
			return true
		}
	}
	return false
}

func objectProcedureWritesModuleFieldIndexed(proc sourceProcedure, name string, facts *procedureAnalysisFacts) bool {
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeModule || (access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) || !strings.EqualFold(cleanIdentifier(access.Name), cleanIdentifier(name)) {
			continue
		}
		if objectMemberReceiver(facts, access) {
			continue
		}
		return true
	}
	return false
}

func objectProcedureWritesParameterIndexed(proc sourceProcedure, name string, summaries map[string]objectProcedureSummary, facts *procedureAnalysisFacts) bool {
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeParameter ||
			(access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) ||
			!strings.EqualFold(cleanIdentifier(access.Name), cleanIdentifier(name)) {
			continue
		}
		if !objectMemberReceiver(facts, access) {
			return true
		}
	}
	for call := range proc.Calls.All() {
		actuals := objectCallActuals(call, facts)
		for actualIndex, actual := range actuals {
			if actual.parenthesized {
				// Parenthesized actuals are temporary ByVal expressions even
				// when the formal parameter is declared ByRef; they cannot
				// propagate a callee mutation back to this parameter.
				continue
			}
			if !strings.EqualFold(cleanIdentifier(actual.text), cleanIdentifier(name)) {
				continue
			}
			if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
				return true
			}
			summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries)
			if !ok {
				return true
			}
			formalIndex := objectFormalIndex(call, summary, actualIndex)
			if summary.ByRefWritten[formalIndex] {
				return true
			}
		}
	}
	return false
}

type objectFlowOptions struct {
	Entry map[string]bool
}

// objectUseBeforeSetIRFindingsPlan reports the first unsafe member/collection use
// of each object variable.  The state at a use comes from the CFG entry fact,
// not from source-line order, so branches, early exits, loops and error edges
// are all represented by the same must-analysis.
func (a Analyzer) objectUseBeforeSetIRFindingsPlan(plan *objectProcedurePlan, summaries map[string]objectProcedureSummary, entry map[string]bool) []Finding {
	if plan == nil || plan.proc.Graph == nil {
		return nil
	}
	if !plan.relevant {
		return nil
	}
	file, proc := plan.file, plan.proc
	declarations := plan.declarations
	flow := objectStateFlowPlan(plan, summaries, objectFlowOptions{Entry: entry})
	facts := plan.flowContext.facts

	reported := map[string]bool{}
	guardCache := &objectFlowGuardCache{}
	var findings []Finding
	// ProcedureIR emits accesses in source order. Keep that order directly;
	// sorting a copied collection (or an index permutation) only adds work.
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeLocal && access.Scope != procedureir.ScopeModule && access.Scope != procedureir.ScopeParameter {
			continue
		}
		// Indexed/default-property assignment still dereferences the receiver
		// before writing its item.  Keep plain `Set obj = ...` out via the
		// receiver-shape check below, but admit AccessWrite here so `dict(key) =`
		// and `collection(i) =` are diagnosed at the root object.
		if access.Mode != procedureir.AccessRead && access.Mode != procedureir.AccessReadWrite && access.Mode != procedureir.AccessWrite {
			continue
		}
		if !objectMemberReceiver(facts, access) {
			continue
		}
		declaration, scope, ok := objectDeclarationBinding(access.Name, declarations)
		// `As New` is guaranteed to produce a value on first use.  A normal
		// Static declaration, however, has module-lifetime state and may still
		// be Nothing on its first read; it therefore remains in the analysis.
		if !ok || !declaration.Object || declaration.NewExpression {
			continue
		}
		// An indexed object-array assignment constructs/replaces one element;
		// the array variable itself is not a receiver object that must already
		// be non-Nothing.  Reads of array elements remain eligible for the
		// ordinary object-state check.
		if declaration.Array && access.Mode == procedureir.AccessWrite {
			continue
		}
		variable := objectVariable{Scope: scope, Name: access.Name}
		key := variable.key()
		block, ok := plan.flowGraph.BlockForStatement(access.StatementID)
		if !ok || !plan.reachable[block.ID] {
			continue
		}
		if _, ok := flow.in[block.ID]; !ok {
			// The error-raising graph view may prove that a source-recovered
			// access has no normal path. Do not index a missing state as
			// MaybeNothing and manufacture a finding for that unreachable block.
			continue
		}
		if objectFlowAssigned(flow.in[block.ID], variable) || reported[key] {
			continue
		}
		if objectErrorResumeNextAt(proc, access.StatementID) {
			continue
		}
		if objectFlowGuardedByOpenFlag(proc, access.StatementID, access.Name, guardCache) {
			continue
		}
		findings = append(findings, a.simpleFinding(
			file, proc, access.Range.StartLine, "VBA202", "warning",
			declaration.Name+" may be dereferenced before a definitely non-Nothing value is proven.",
			"Every reachable path must establish a non-Nothing object before member, collection, or default-property access; otherwise runtime error 91 may occur.",
			"Assign `Set "+declaration.Name+" = ...` on every path before dereferencing it, or guard `If "+declaration.Name+" Is Nothing Then`.",
		))
		reported[key] = true
	}
	// The IR intentionally suppresses identifiers used as call callees from
	// VariableAccess (a procedure name is not normally a variable read).  An
	// object used as a default item/index call, however, is a receiver read:
	// `dict(key)`, `collection(i)`, and `obj.Property(...)` all dereference the
	// object before invoking the call.  Recover those roots from CallSite.
	// Calls are likewise source ordered in the canonical IR.
	for call := range proc.Calls.All() {
		name := objectCallReceiverName(call)
		if name == "" {
			continue
		}
		declaration, scope, ok := objectDeclarationBinding(name, declarations)
		if !ok || !declaration.Object || declaration.NewExpression || declaration.Array {
			continue
		}
		variable := objectVariable{Scope: scope, Name: name}
		key := variable.key()
		if reported[key] {
			continue
		}
		if _, exists := flow.vars[key]; !exists {
			continue
		}
		block, ok := plan.flowGraph.BlockForStatement(call.StatementID)
		if !ok || !plan.reachable[block.ID] {
			continue
		}
		if _, ok := flow.in[block.ID]; !ok {
			continue
		}
		if objectFlowAssigned(flow.in[block.ID], variable) {
			continue
		}
		if objectErrorResumeNextAt(proc, call.StatementID) {
			continue
		}
		if objectFlowGuardedByOpenFlag(proc, call.StatementID, name, guardCache) {
			continue
		}
		findings = append(findings, a.simpleFinding(
			file, proc, call.Range.StartLine, "VBA202", "warning",
			declaration.Name+" may be dereferenced before a definitely non-Nothing value is proven.",
			"Every reachable path must establish a non-Nothing object before member, collection, or default-property access; otherwise runtime error 91 may occur.",
			"Assign `Set "+declaration.Name+" = ...` on every path before dereferencing it, or guard `If "+declaration.Name+" Is Nothing Then`.",
		))
		reported[key] = true
	}
	return findings
}

func objectCallReceiverName(call procedureir.CallSite) string {
	if call.Callee.Receiver != nil {
		return cleanIdentifier(*call.Callee.Receiver)
	}
	if call.Arguments.Count > 0 && call.Callee.BaseName != "" {
		return cleanIdentifier(call.Callee.BaseName)
	}
	return ""
}

// On Error Resume Next deliberately turns an object dereference into a
// recoverable probe.  VBA202 should not report the probe itself; subsequent
// code is still analyzed normally after an explicit error-mode reset.
func objectErrorResumeNextAt(proc sourceProcedure, statementID int) bool {
	active := false
	validated := false
	for statement := range proc.Statements.All() {
		if statement.ID == statementID {
			return active || validated
		}
		text := strings.ToLower(strings.TrimSpace(statement.Text))
		if !active && strings.Contains(text, "raiseassertfailure") {
			// A common VBA probe pattern checks Err.Number and routes failures
			// through a terminating assertion after resetting the error mode.
			validated = true
		}
		if statement.Kind != procedureir.StatementOnError {
			continue
		}
		switch {
		case strings.Contains(text, "on error resume next"):
			active = true
		case strings.Contains(text, "on error goto 0"), strings.Contains(text, "on error goto -1"), strings.Contains(text, "on error goto"):
			active = false
		}
	}
	return active
}

type objectFlowGuardCache struct {
	graph      vbacfg.CFGView
	graphReady bool
	dominators map[vbacfg.BlockID][]vbacfg.BlockID
	successors map[vbacfg.BlockID][]vbacfg.BlockID
}

func objectFlowGuardedByOpenFlag(proc sourceProcedure, statementID int, objectName string, cache *objectFlowGuardCache) bool {
	if proc.Graph == nil {
		return false
	}
	if !cache.graphReady {
		cache.graph = proc.Graph.WithoutNormalErrRaiseContinuationView()
		cache.graphReady = true
		cache.dominators = cache.graph.Dominators()
	}
	flowProc := proc
	accessBlock, ok := cache.graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	dominators := cache.dominators
	for statement := range proc.Statements.All() {
		flag, negated, ok := objectBooleanGuard(statement)
		if !ok || !objectOpenFlagForObject(flag, objectName) {
			continue
		}
		conditionBlock, ok := cache.graph.BlockForStatement(statement.ID)
		if !ok || !objectBlockSetContains(dominators[accessBlock.ID], conditionBlock.ID) {
			continue
		}
		if !objectOpenFlagAssignmentDominates(flowProc, dominators, conditionBlock.ID, statement.ID, flag, objectName) {
			continue
		}
		safeBranch := vbacfg.EdgeBranchTrue
		if negated {
			safeBranch = vbacfg.EdgeBranchFalse
		}
		safeReachable := false
		unsafeReachable := false
		cache.buildSuccessors()
		cache.graph.ForEachOutgoing(conditionBlock.ID, func(edge vbacfg.Edge) bool {
			if edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
				return true
			}
			reachesAccess := objectFlowCanReach(cache.successors, edge.To, accessBlock.ID)
			if edge.Kind == safeBranch {
				safeReachable = safeReachable || reachesAccess
			} else {
				unsafeReachable = unsafeReachable || reachesAccess
			}
			return true
		})
		if safeReachable && !unsafeReachable {
			return true
		}
	}
	return false
}

func (cache *objectFlowGuardCache) buildSuccessors() {
	if cache.successors != nil {
		return
	}
	cache.successors = make(map[vbacfg.BlockID][]vbacfg.BlockID)
	cache.graph.ForEachEdge(func(edge vbacfg.Edge) bool {
		if edge.Class != vbacfg.EdgeExceptional {
			cache.successors[edge.From] = append(cache.successors[edge.From], edge.To)
		}
		return true
	})
}

func objectBooleanGuard(statement procedureir.Statement) (string, bool, bool) {
	if statement.Condition == nil {
		return "", false, false
	}
	text := strings.ToLower(strings.TrimSpace(statement.Condition.Text))
	if then := strings.Index(text, " then"); then >= 0 {
		text = strings.TrimSpace(text[:then])
	}
	negated := false
	if strings.HasPrefix(text, "not ") {
		negated = true
		text = strings.TrimSpace(strings.TrimPrefix(text, "not "))
	}
	name := cleanIdentifier(text)
	if name == "" || strings.ContainsAny(name, " .()=<>") {
		return "", false, false
	}
	return name, negated, true
}

func objectOpenFlagForObject(flag, objectName string) bool {
	flag = strings.ToLower(cleanIdentifier(flag))
	objectName = strings.ToLower(cleanIdentifier(objectName))
	return objectName != "" && strings.HasPrefix(flag, objectName) && strings.HasSuffix(flag, "opened")
}

func objectOpenFlagAssignmentDominates(proc sourceProcedure, dominators map[vbacfg.BlockID][]vbacfg.BlockID, conditionBlock vbacfg.BlockID, conditionStatementID int, flag, objectName string) bool {
	flag = strings.ToLower(cleanIdentifier(flag))
	objectName = strings.ToLower(cleanIdentifier(objectName))
	for flagStatement := range proc.Statements.All() {
		if strings.ToLower(compactStatement(flagStatement.Text)) != flag+"=true" {
			continue
		}
		flagBlock, ok := proc.Graph.BlockForStatement(flagStatement.ID)
		if !ok {
			continue
		}
		// A handler can be reached through an exceptional edge before the
		// assignment dominates the handler's block.  The guarded branch still
		// proves that the flag is True, so source order is a valid fallback for
		// the known flag/object pair after the object assignment dominates the
		// flag write itself.
		flagBeforeCondition := objectBlockSetContains(dominators[conditionBlock], flagBlock.ID) || flagStatement.ID < conditionStatementID
		if !flagBeforeCondition {
			continue
		}
		for objectStatement := range proc.Statements.All() {
			if !strings.HasPrefix(strings.ToLower(compactStatement(objectStatement.Text)), "set"+objectName+"=") {
				continue
			}
			objectBlock, ok := proc.Graph.BlockForStatement(objectStatement.ID)
			if ok && objectBlockSetContains(dominators[flagBlock.ID], objectBlock.ID) {
				return true
			}
		}
	}
	return false
}

func objectBlockSetContains(blocks []vbacfg.BlockID, candidate vbacfg.BlockID) bool {
	for _, block := range blocks {
		if block == candidate {
			return true
		}
	}
	return false
}

func objectFlowCanReach(successors map[vbacfg.BlockID][]vbacfg.BlockID, start, target vbacfg.BlockID) bool {
	if start == target {
		return true
	}
	seen := map[vbacfg.BlockID]bool{start: true}
	queue := []vbacfg.BlockID{start}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for _, successor := range successors[current] {
			if seen[successor] {
				continue
			}
			if successor == target {
				return true
			}
			seen[successor] = true
			queue = append(queue, successor)
		}
	}
	return false
}

type objectFlowResult struct {
	in         map[vbacfg.BlockID]map[string]bool
	out        map[vbacfg.BlockID]map[string]bool
	vars       map[string]objectVariable
	normalExit map[string]bool
}

type objectFlowContext struct {
	facts               *procedureAnalysisFacts
	predecessors        map[vbacfg.BlockID][]vbacfg.Edge
	vars                map[string]objectVariable
	receiverSummaryKeys map[string][]string
	valueState          map[string]bool
	predicateContracts  map[string]bool
	terminalCalls       map[int]bool
}

func newObjectFlowContext(proc sourceProcedure, graph vbacfg.CFGView) objectFlowContext {
	context := objectFlowContext{
		facts:        proc.analysisFacts(),
		predecessors: make(map[vbacfg.BlockID][]vbacfg.Edge),
		vars:         map[string]objectVariable{},
	}
	if proc.Graph != nil {
		graph.ForEachEdge(func(edge vbacfg.Edge) bool {
			context.predecessors[edge.To] = append(context.predecessors[edge.To], edge)
			return true
		})
	}
	return context
}

func objectFlowDeclarations(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) declarationScope {
	declarations := newDeclarationScope(file, proc)
	declarations.module = moduleDecls
	return declarations
}

func objectDeclarationFor(name string, scope procedureir.SymbolScope, declarations declarationScope) (sourceDeclaration, bool) {
	key := strings.ToLower(cleanIdentifier(name))
	switch scope {
	case procedureir.ScopeModule:
		declaration, ok := declarations.module[key]
		return declaration, ok
	case procedureir.ScopeParameter:
		if declaration, ok := declarations.parameters[key]; ok {
			return declaration, true
		}
	}
	return declarations.lookup(key)
}

func objectDeclarationBinding(name string, declarations declarationScope) (sourceDeclaration, procedureir.SymbolScope, bool) {
	key := strings.ToLower(cleanIdentifier(name))
	if key == "" {
		return sourceDeclaration{}, procedureir.ScopeUnresolved, false
	}
	if declaration, ok := declarations.parameters[key]; ok {
		return declaration, procedureir.ScopeParameter, true
	}
	if declaration, ok := declarations.local[key]; ok {
		return declaration, procedureir.ScopeLocal, true
	}
	if declaration, ok := declarations.extra[key]; ok {
		return declaration, procedureir.ScopeLocal, true
	}
	if declaration, ok := declarations.module[key]; ok {
		return declaration, procedureir.ScopeModule, true
	}
	return sourceDeclaration{}, procedureir.ScopeUnresolved, false
}

func objectStateFlowPlan(plan *objectProcedurePlan, summaries map[string]objectProcedureSummary, options objectFlowOptions) objectFlowResult {
	result := objectFlowResult{in: map[vbacfg.BlockID]map[string]bool{}, out: map[vbacfg.BlockID]map[string]bool{}, vars: map[string]objectVariable{}}
	if plan == nil || plan.proc.Graph == nil || len(plan.vars) == 0 {
		return result
	}
	result.vars = plan.vars
	flowProc := plan.flowProc
	flowGraph := plan.flowGraph
	flowContext := plan.flowContext
	flowContext.receiverSummaryKeys = plan.receiverSummaryKeys
	initial := map[string]bool{}
	for key, variable := range plan.vars {
		initial[key] = false
		declaration, _ := objectDeclarationFor(variable.Name, variable.Scope, plan.declarations)
		if declaration.NewExpression || objectClassLifecycleAssignedPlan(plan, variable, summaries) || (options.Entry != nil && options.Entry[variable.key()]) {
			initial[key] = true
		}
		// Parameters are intentionally not initialized at procedure entry.  A
		// caller can pass Nothing, and only a dominating Set or a proven ByRef
		// initializer establishes a safe value.
	}
	reachable := plan.reachable
	flowGraph.ForEachBlock(func(block vbacfg.Block) bool {
		if !reachable[block.ID] {
			return true
		}
		if block.ID == flowGraph.Entry() {
			result.in[block.ID] = cloneObjectState(initial)
		} else {
			result.in[block.ID] = objectStateAllTrue(plan.vars)
		}
		result.out[block.ID] = objectFlowTransfer(plan.file, flowProc, block, result.in[block.ID], plan.vars, plan.declarations, summaries, flowContext)
		return true
	})
	changed := true
	for changed {
		changed = false
		flowGraph.ForEachBlock(func(block vbacfg.Block) bool {
			if !reachable[block.ID] || block.ID == flowGraph.Entry() {
				return true
			}
			incoming := make([]map[string]bool, 0)
			for _, edge := range flowContext.predecessors[block.ID] {
				if !reachable[edge.From] {
					continue
				}
				state := result.out[edge.From]
				terminalExceptional := edge.Class == vbacfg.EdgeExceptional && objectFlowBlockHasTerminalCall(flowGraph, edge.From, flowContext)
				if edge.Class == vbacfg.EdgeExceptional {
					state = result.in[edge.From]
					if terminalExceptional {
						state = objectFlowExceptionalTerminalState(flowGraph, flowProc, edge.From, state, plan.vars, plan.declarations, flowContext, summaries)
					}
				}
				if (edge.Uncertain || objectFlowForEachZeroIteration(flowContext, edge)) && !terminalExceptional {
					state = result.in[edge.From]
				}
				state = objectFlowApplyGuard(state, flowContext, edge, plan.declarations)
				incoming = append(incoming, state)
			}
			if len(incoming) == 0 {
				return true
			}
			next := objectFlowIntersection(incoming, plan.vars)
			if !objectStateEqual(result.in[block.ID], next) {
				result.in[block.ID] = next
				changed = true
			}
			updated := objectFlowTransfer(plan.file, flowProc, block, next, plan.vars, plan.declarations, summaries, flowContext)
			if !objectStateEqual(result.out[block.ID], updated) {
				result.out[block.ID] = updated
				changed = true
			}
			return true
		})
	}
	if plan.proc.Graph != nil {
		result.normalExit = cloneObjectState(result.in[flowGraph.NormalExit()])
	}
	return result
}

func objectFlowExceptionalTerminalState(graph vbacfg.CFGView, proc sourceProcedure, blockID vbacfg.BlockID, input map[string]bool, vars map[string]objectVariable, declarations declarationScope, flowContext objectFlowContext, summaries map[string]objectProcedureSummary) map[string]bool {
	if !objectFlowBlockHasTerminalCall(graph, blockID, flowContext) {
		return input
	}
	block, ok := graph.BlockByID(blockID)
	if !ok || block.Statement == nil {
		return input
	}
	state := cloneObjectState(input)
	flowContext.facts.forEachCallForStatement(block.Statement.ID, func(call procedureir.CallSite) {
		if flowContext.terminalCalls[call.ID] {
			applyObjectCallEffects(proc, call, state, vars, declarations, flowContext.facts, summaries)
		}
	})
	return state
}

func objectFlowBlockHasTerminalCall(graph vbacfg.CFGView, blockID vbacfg.BlockID, flowContext objectFlowContext) bool {
	if len(flowContext.terminalCalls) == 0 || flowContext.facts == nil {
		return false
	}
	block, ok := graph.BlockByID(blockID)
	if !ok || block.Statement == nil {
		return false
	}
	found := false
	flowContext.facts.forEachCallForStatement(block.Statement.ID, func(call procedureir.CallSite) {
		found = found || flowContext.terminalCalls[call.ID]
	})
	return found
}

// objectClassLifecycleAssigned recognizes only VBA's language-guaranteed class
// construction hook.  Ordinary project procedures named Initialize, Setup,
// Load, and similar are callable code, not implicit constructors; their effects
// are applied only at resolved call sites.  The lifecycle summary itself is
// produced by the CFG flow and therefore requires a non-Nothing assignment on
// every reachable normal exit.
func objectClassLifecycleAssigned(proc sourceProcedure, variable objectVariable, summaries map[string]objectProcedureSummary) bool {
	if variable.Scope != procedureir.ScopeModule || !strings.EqualFold(proc.ModuleKind, "class") || strings.EqualFold(proc.Name, "Class_Initialize") {
		return false
	}
	for _, summary := range summaries {
		if !strings.EqualFold(summary.Module, proc.Module) || !strings.EqualFold(summary.QualifiedName, proc.Module+".Class_Initialize") {
			continue
		}
		if summary.ModuleAssigned[strings.ToLower(cleanIdentifier(variable.Name))] {
			return true
		}
	}
	return false
}

func objectClassLifecycleAssignedPlan(plan *objectProcedurePlan, variable objectVariable, summaries map[string]objectProcedureSummary) bool {
	if plan == nil || variable.Scope != procedureir.ScopeModule || !strings.EqualFold(plan.proc.ModuleKind, "class") || strings.EqualFold(plan.proc.Name, "Class_Initialize") {
		return false
	}
	if !plan.classIndexBuilt {
		// Compatibility callers may construct a standalone plan without the
		// batch index. Preserve the previous summary-scan behavior there.
		return objectClassLifecycleAssigned(plan.proc, variable, summaries)
	}
	for _, initializerKey := range plan.classInitializerKeys {
		if summary, ok := summaries[initializerKey]; ok && summary.ModuleAssigned[strings.ToLower(cleanIdentifier(variable.Name))] {
			return true
		}
	}
	return false
}

// objectFlowApplyGuard refines the state on the edge selected by a direct
// `obj Is Nothing`/`Not obj Is Nothing` condition.  Compound boolean
// expressions are deliberately ignored; VBA212 remains responsible for
// short-circuit/eager Boolean diagnostics and this analysis stays conservative.
func objectFlowApplyGuard(state map[string]bool, flowContext objectFlowContext, edge vbacfg.Edge, declarations declarationScope) map[string]bool {
	if edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	statement, ok := flowContext.facts.Statement(edge.StatementID)
	if !ok || statement.Condition == nil {
		return state
	}
	text := objectTrimOuterParens(strings.ToLower(strings.TrimSpace(statement.Condition.Text)))
	if strings.HasPrefix(text, "not ") {
		text = "not " + objectTrimOuterParens(strings.TrimSpace(strings.TrimPrefix(text, "not ")))
	}
	if comparison, ok := objectErrNumberGuard(text); ok && objectFlowExceptionalOnly(flowContext, edge.From) {
		possible := (comparison == "nonzero" && edge.Kind == vbacfg.EdgeBranchTrue) ||
			(comparison == "zero" && edge.Kind == vbacfg.EdgeBranchFalse)
		if !possible {
			// An error handler is entered only after Err.Number is populated. A
			// branch contradicting that fact is not a normal path; preserve the
			// lattice identity for the impossible edge so it cannot manufacture
			// a nullable return value at the procedure exit.
			return objectStateAllTrue(flowContext.vars)
		}
	}
	if name, expected, equals, ok := objectTypeNameGuard(text); ok {
		key := objectGuardVariableKey(name, state, declarations)
		if key == "" {
			return state
		}
		nonNothing := (equals && edge.Kind == vbacfg.EdgeBranchTrue) || (!equals && edge.Kind == vbacfg.EdgeBranchFalse)
		if strings.EqualFold(expected, "nothing") {
			nonNothing = !nonNothing
		}
		if nonNothing {
			updated := cloneObjectState(state)
			updated[key] = true
			return updated
		}
		return state
	}
	if name, predicateTrue, ok := objectNonNothingPredicateGuard(text); ok {
		if !objectFlowPredicateHasContract(flowContext, edge.StatementID, name) &&
			!flowContext.predicateContracts["name:"+strings.ToLower(cleanIdentifier(name))] {
			return state
		}
		key := objectGuardVariableKey(name, state, declarations)
		if key == "" {
			return state
		}
		// The contract proves that the predicate result itself is true only on
		// the matching edge.  In particular, `If Not IsExcelTable(x) Then`
		// must not refine the true edge: a Nothing argument makes the helper
		// return its default False value, so that branch can still dereference x.
		predicateResultTrue := (predicateTrue && edge.Kind == vbacfg.EdgeBranchTrue) ||
			(!predicateTrue && edge.Kind == vbacfg.EdgeBranchFalse)
		if !predicateResultTrue {
			return state
		}
		updated := cloneObjectState(state)
		updated[key] = true
		return updated
	}
	marker := " is nothing"
	index := strings.Index(text, marker)
	if index <= 0 || strings.Contains(text[index+len(marker):], " ") || strings.Contains(text[:index], " and ") || strings.Contains(text[:index], " or ") {
		return state
	}
	left := strings.TrimSpace(text[:index])
	not := false
	if strings.HasPrefix(left, "not ") {
		not = true
		left = strings.TrimSpace(strings.TrimPrefix(left, "not "))
	}
	left = objectTrimOuterParens(left)
	name := cleanIdentifier(left)
	if name == "" || strings.ContainsAny(name, " .()") {
		return state
	}
	key := objectGuardVariableKey(name, state, declarations)
	if key == "" {
		return state
	}
	nonNothingOnTrue := not
	if edge.Kind == vbacfg.EdgeBranchFalse {
		nonNothingOnTrue = !nonNothingOnTrue
	}
	if !not && edge.Kind == vbacfg.EdgeBranchTrue && objectFlowInlineGuardAssignment(statement, name) {
		nonNothingOnTrue = true
	}
	if _, ok := state[key]; ok {
		updated := cloneObjectState(state)
		updated[key] = nonNothingOnTrue
		return updated
	}
	return state
}

func objectFlowPredicateHasContract(flowContext objectFlowContext, statementID int, argumentName string) bool {
	if flowContext.facts == nil || len(flowContext.predicateContracts) == 0 {
		return false
	}
	matched := false
	flowContext.facts.forEachCallForStatement(statementID, func(call procedureir.CallSite) {
		if matched || call.Callee.Receiver != nil {
			return
		}
		calleeName := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
		if calleeName == "" {
			calleeName = strings.ToLower(objectBareCallName(call.Callee.Text))
		}
		contract := flowContext.predicateContracts["name:"+calleeName]
		if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
			candidate := call.Resolution.Candidates[0]
			key := objectSummaryKey(candidate.File, candidate.QualifiedName, candidate.Kind, candidate.Line)
			contract = contract || flowContext.predicateContracts[key] || flowContext.predicateContracts[strings.ToLower(candidate.QualifiedName)]
		}
		if !contract {
			return
		}
		for _, actual := range objectCallActuals(call, flowContext.facts) {
			if !actual.parenthesized && strings.EqualFold(cleanIdentifier(actual.text), cleanIdentifier(argumentName)) {
				matched = true
				return
			}
		}
	})
	return matched
}

func objectTrimOuterParens(text string) string {
	text = strings.TrimSpace(text)
	for len(text) >= 2 && text[0] == '(' && text[len(text)-1] == ')' {
		depth := 0
		wrapped := true
		for index, character := range text {
			switch character {
			case '(':
				depth++
			case ')':
				depth--
				if depth == 0 && index != len(text)-1 {
					wrapped = false
				}
			}
			if depth < 0 {
				wrapped = false
				break
			}
		}
		if !wrapped || depth != 0 {
			break
		}
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	return text
}

func objectErrNumberGuard(text string) (string, bool) {
	compact := compactStatement(strings.ToLower(strings.TrimSpace(text)))
	switch compact {
	case "err.number<>0":
		return "nonzero", true
	case "err.number=0":
		return "zero", true
	default:
		return "", false
	}
}

func objectFlowExceptionalOnly(flowContext objectFlowContext, blockID vbacfg.BlockID) bool {
	return objectFlowExceptionalOnlyFrom(flowContext, blockID, map[vbacfg.BlockID]bool{})
}

func objectFlowExceptionalOnlyFrom(flowContext objectFlowContext, blockID vbacfg.BlockID, seen map[vbacfg.BlockID]bool) bool {
	if seen[blockID] {
		return false
	}
	seen[blockID] = true
	pred := flowContext.predecessors[blockID]
	if len(pred) == 0 {
		return false
	}
	for _, edge := range pred {
		if edge.Class == vbacfg.EdgeExceptional {
			continue
		}
		if edge.Class != vbacfg.EdgeNormal || edge.Kind != vbacfg.EdgeFallthrough ||
			!objectFlowExceptionalOnlyFrom(flowContext, edge.From, seen) {
			return false
		}
	}
	return true
}

func objectGuardVariableKey(name string, state map[string]bool, declarations declarationScope) string {
	name = cleanIdentifier(strings.TrimSpace(name))
	if name == "" {
		return ""
	}
	if declaration, scope, ok := objectDeclarationBinding(name, declarations); ok {
		if !declaration.Object {
			return ""
		}
		key := (objectVariable{Scope: scope, Name: name}).key()
		if _, exists := state[key]; exists {
			return key
		}
		return ""
	}
	for _, scope := range []procedureir.SymbolScope{procedureir.ScopeLocal, procedureir.ScopeParameter, procedureir.ScopeModule} {
		key := (objectVariable{Scope: scope, Name: name}).key()
		if _, exists := state[key]; exists {
			return key
		}
	}
	return ""
}

func objectTypeNameGuard(text string) (string, string, bool, bool) {
	text = strings.TrimSpace(text)
	prefix := "typename("
	if !strings.HasPrefix(text, prefix) {
		return "", "", false, false
	}
	close := strings.Index(text[len(prefix):], ")")
	if close < 0 {
		return "", "", false, false
	}
	close += len(prefix)
	name := cleanIdentifier(strings.TrimSpace(text[len(prefix):close]))
	rest := strings.TrimSpace(text[close+1:])
	operator := ""
	switch {
	case strings.HasPrefix(rest, "<>"):
		operator = "<>"
	case strings.HasPrefix(rest, "="):
		operator = "="
	default:
		return "", "", false, false
	}
	rest = strings.TrimSpace(rest[len(operator):])
	if len(rest) < 2 || rest[0] != '"' {
		return "", "", false, false
	}
	end := strings.Index(rest[1:], "\"")
	if end < 0 {
		return "", "", false, false
	}
	end++
	expected := strings.TrimSpace(rest[1:end])
	if name == "" || expected == "" {
		return "", "", false, false
	}
	return name, expected, operator == "=", true
}

func objectNonNothingPredicateGuard(text string) (string, bool, bool) {
	text = objectTrimOuterParens(strings.TrimSpace(text))
	if then := strings.Index(text, " then"); then >= 0 {
		text = strings.TrimSpace(text[:then])
	}
	negated := false
	if strings.HasPrefix(text, "not ") {
		negated = true
		text = objectTrimOuterParens(strings.TrimSpace(strings.TrimPrefix(text, "not ")))
	}
	const prefix = "isexceltable("
	if !strings.HasPrefix(text, prefix) || !strings.HasSuffix(text, ")") {
		return "", false, false
	}
	argument := strings.TrimSpace(text[len(prefix) : len(text)-1])
	if argument == "" || strings.ContainsAny(argument, ".()") {
		return "", false, false
	}
	return cleanIdentifier(argument), !negated, true
}

func objectFlowInlineGuardAssignment(statement procedureir.Statement, name string) bool {
	text := strings.ToLower(strings.TrimSpace(statement.Text))
	then := strings.Index(text, " then ")
	if then < 0 {
		return false
	}
	tail := strings.TrimSpace(text[then+len(" then "):])
	prefix := "set " + strings.ToLower(cleanIdentifier(name)) + " ="
	if !strings.HasPrefix(tail, prefix) {
		return false
	}
	rhs := strings.TrimSpace(strings.TrimPrefix(tail, prefix))
	return strings.HasPrefix(rhs, "new ") || strings.HasPrefix(rhs, "createobject(") || strings.HasPrefix(rhs, "getobject(")
}

func objectFlowExitDefinitelyAssigned(flow objectFlowResult, variable objectVariable) bool {
	return objectFlowAssigned(flow.normalExit, variable)
}

func objectFlowAssigned(state map[string]bool, variable objectVariable) bool {
	return state[variable.key()]
}

func objectFlowTransfer(file parsedFile, proc sourceProcedure, block vbacfg.Block, input map[string]bool, vars map[string]objectVariable, declarations declarationScope, summaries map[string]objectProcedureSummary, flowContext objectFlowContext) map[string]bool {
	state := cloneObjectState(input)
	valueState := cloneObjectState(input)
	statement := block.Statement
	if statement == nil {
		return state
	}
	callSeen := false
	flowContext.facts.forEachCallForStatement(statement.ID, func(call procedureir.CallSite) {
		callSeen = true
		applyObjectCallEffects(proc, call, state, vars, declarations, flowContext.facts, summaries)
	})
	// The IR represents implicit call statements such as `Helper obj` as an
	// assignment-shaped statement. Their argument effects were applied above;
	// do not then treat the first argument as an object assignment target.
	if statement.Kind == procedureir.StatementAssignment && callSeen {
		return state
	}
	target, ok := objectFlowTarget(proc, *statement, declarations, flowContext)
	if !ok {
		return state
	}
	switch statement.Kind {
	case procedureir.StatementSet:
		flowContext.valueState = valueState
		state[target.key()] = objectFlowValueAssigned(proc, *statement, state, flowContext, declarations, summaries)
	case procedureir.StatementAssignment, procedureir.StatementReDim, procedureir.StatementFor:
		// A value assignment is not an object Set.  Treat it as unsafe even if
		// malformed VBA happens to compile through implicit coercion.
		state[target.key()] = false
	case procedureir.StatementForEach:
		state[target.key()] = true
	}
	_ = file
	return state
}

func objectFlowTarget(proc sourceProcedure, statement procedureir.Statement, declarations declarationScope, flowContext objectFlowContext) (objectVariable, bool) {
	var accessTarget objectVariable
	accessFound := false
	flowContext.facts.forEachAccessForStatement(statement.ID, func(access procedureir.VariableAccess) {
		if accessFound {
			return
		}
		if access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite {
			return
		}
		// A receiver in `obj.Member = value` or `dict(key) = value` is
		// dereferenced, but it is not the object reference being assigned.  Do
		// not reset its state merely because the member/index write is modeled
		// as AccessWrite by the IR.
		if objectMemberReceiver(flowContext.facts, access) {
			return
		}
		if objectIndexedReceiverWrite(flowContext.facts, statement.ID, access) {
			return
		}
		declaration, scope, ok := objectDeclarationBinding(access.Name, declarations)
		if !ok || !declaration.Object {
			return
		}
		accessTarget = objectVariable{Scope: scope, Name: access.Name}
		accessFound = true
	})
	if accessFound {
		return accessTarget, true
	}
	if statement.Target != nil && statement.Target.Kind != procedureir.ExpressionCall && statement.Target.Kind != procedureir.ExpressionMember && !objectIndexedTargetCall(flowContext.facts, statement.ID, statement.Target.Text) {
		name := cleanIdentifier(strings.TrimSpace(statement.Target.Text))
		if declaration, scope, ok := objectDeclarationBinding(name, declarations); ok && declaration.Object {
			return objectVariable{Scope: scope, Name: name}, true
		}
	}
	if proc.Name != "" && statement.Target != nil && strings.EqualFold(cleanIdentifier(statement.Target.Text), proc.Name) && isObjectType(proc.ReturnType) {
		return objectVariable{Scope: procedureir.ScopeLocal, Name: proc.Name}, true
	}
	return objectVariable{}, false
}

func objectIndexedReceiverWrite(facts *procedureAnalysisFacts, statementID int, access procedureir.VariableAccess) bool {
	found := false
	facts.forEachCallForStatement(statementID, func(call procedureir.CallSite) {
		if found {
			return
		}
		if call.Callee.Receiver == nil && call.Arguments.Count > 0 && strings.EqualFold(cleanIdentifier(call.Callee.BaseName), cleanIdentifier(access.Name)) {
			found = true
		}
	})
	return found
}

func objectIndexedTargetCall(facts *procedureAnalysisFacts, statementID int, targetText string) bool {
	name := cleanIdentifier(strings.TrimSpace(targetText))
	found := false
	facts.forEachCallForStatement(statementID, func(call procedureir.CallSite) {
		if found {
			return
		}
		if call.Callee.Receiver == nil && call.Arguments.Count > 0 && strings.EqualFold(cleanIdentifier(call.Callee.BaseName), name) {
			found = true
		}
	})
	return found
}

func objectFlowValueAssigned(proc sourceProcedure, statement procedureir.Statement, state map[string]bool, flowContext objectFlowContext, declarations declarationScope, summaries map[string]objectProcedureSummary) bool {
	value := statement.Value
	if value == nil {
		return false
	}
	return objectExpressionAssigned(proc, *value, state, flowContext, declarations, summaries, statement.ID)
}

func objectExpressionAssigned(proc sourceProcedure, expression procedureir.Expression, state map[string]bool, flowContext objectFlowContext, declarations declarationScope, summaries map[string]objectProcedureSummary, statementID int) bool {
	facts := flowContext.facts
	text := strings.TrimSpace(expression.Text)
	lower := strings.ToLower(text)
	if lower == "nothing" || strings.HasPrefix(lower, "nothing ") {
		return false
	}
	switch expression.Kind {
	case procedureir.ExpressionNew:
		return true
	case procedureir.ExpressionParentheses:
		for _, child := range expression.Children {
			if nested, ok := facts.Expression(child); ok {
				return objectExpressionAssigned(proc, nested, state, flowContext, declarations, summaries, statementID)
			}
		}
	case procedureir.ExpressionIdentifier:
		name := cleanIdentifier(text)
		if objectIntrinsicIdentifierAssigned(proc, name) {
			return true
		}
		declaration, scope, ok := objectDeclarationBinding(name, declarations)
		if !ok {
			if isObjectType(proc.ReturnType) && strings.EqualFold(name, cleanIdentifier(proc.Name)) {
				return state[(objectVariable{Scope: procedureir.ScopeLocal, Name: name}).key()]
			}
			return objectBareObjectFunctionAssigned(proc, name, summaries)
		}
		if !declaration.Object {
			return false
		}
		return state[(objectVariable{Scope: scope, Name: name}).key()]
	case procedureir.ExpressionCall:
		valueState := state
		if flowContext.valueState != nil {
			valueState = flowContext.valueState
		}
		for call := range proc.Calls.All() {
			if call.StatementID != statementID || (call.ExpressionID != 0 && call.ExpressionID != expression.ID) {
				continue
			}
			if objectCallReturnsAssigned(proc, statementID, call, valueState, flowContext, declarations, summaries) {
				return true
			}
		}
		return objectConstructorCallText(lower)
	case procedureir.ExpressionMember:
		if objectExcelMemberExpressionAssigned(expression.Text, proc, declarations) {
			return true
		}
		if objectMemberFunctionAssigned(proc, expression.Text, declarations, summaries) {
			return true
		}
		// A member rooted at an intrinsic workbook/application object is a
		// non-Nothing factory value.  For a user object, the member may itself be
		// Nothing; keep the assignment nullable and report a later dereference.
		for _, childID := range expression.Children {
			if child, ok := facts.Expression(childID); ok && child.Kind == procedureir.ExpressionIdentifier {
				root := strings.ToLower(cleanIdentifier(child.Text))
				if root == "thisworkbook" || root == "application" || (root == "me" && strings.EqualFold(proc.ModuleKind, "form")) {
					return true
				}
				break
			}
		}
		return false
	}
	if strings.HasPrefix(lower, "new ") {
		return true
	}
	return objectConstructorCallText(lower)
}

func objectBareObjectFunctionAssigned(proc sourceProcedure, name string, summaries map[string]objectProcedureSummary) bool {
	name = cleanIdentifier(name)
	if name == "" {
		return false
	}
	var match objectProcedureSummary
	found := false
	for _, summary := range summaries {
		if !summary.ReturnCollection || !strings.EqualFold(summary.Module, proc.Module) {
			continue
		}
		qualifiedName := strings.TrimSpace(summary.QualifiedName)
		if dot := strings.LastIndexByte(qualifiedName, '.'); dot >= 0 {
			qualifiedName = qualifiedName[dot+1:]
		}
		if !strings.EqualFold(cleanIdentifier(qualifiedName), name) {
			continue
		}
		if found {
			return false
		}
		match = summary
		found = true
	}
	return found && match.ReturnAssigned
}

func objectMemberFunctionAssigned(proc sourceProcedure, text string, declarations declarationScope, summaries map[string]objectProcedureSummary) bool {
	text = strings.TrimSpace(text)
	dot := strings.LastIndexByte(text, '.')
	if dot < 0 {
		return false
	}
	receiver := strings.TrimSpace(text[:dot])
	member := cleanIdentifier(strings.TrimSpace(text[dot+1:]))
	if receiver == "" || member == "" {
		return false
	}
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(strings.SplitN(receiver, ".", 2)[0], "(", 2)[0]))
	typeName := ""
	if strings.EqualFold(root, "me") && (strings.EqualFold(proc.ModuleKind, "class") || strings.EqualFold(proc.ModuleKind, "form")) {
		typeName = proc.Module
	} else if declaration, ok := objectDeclarationByName(root, declarations); ok {
		typeName = lastName(strings.TrimSpace(declaration.Type))
	}
	if typeName == "" {
		return false
	}
	var match objectProcedureSummary
	found := false
	for _, summary := range summaries {
		if !summary.ReturnCollection || !strings.EqualFold(cleanIdentifier(summary.Module), cleanIdentifier(typeName)) || !strings.EqualFold(lastName(summary.QualifiedName), member) {
			continue
		}
		if found {
			return false
		}
		match = summary
		found = true
	}
	return found && match.ReturnAssigned
}

func objectConstructorCallText(text string) bool {
	text = strings.TrimSpace(text)
	return strings.HasPrefix(text, "createobject(") || strings.HasPrefix(text, "getobject(")
}

func objectIntrinsicIdentifierAssigned(proc sourceProcedure, name string) bool {
	switch strings.ToLower(cleanIdentifier(name)) {
	case "thisworkbook", "application":
		// These VBA intrinsic object identifiers are always bound by the host.
		// Treating them as nullable when passed to a private object helper
		// loses the same proof that direct member access already has.
		return true
	case "me":
		// Me is the live instance for class, document, and UserForm
		// procedures. It is not available as an object in standard modules,
		// so keep the scope check explicit.
		return strings.EqualFold(proc.ModuleKind, "class") || strings.EqualFold(proc.ModuleKind, "form")
	default:
		return false
	}
}

func objectCallReturnsAssigned(proc sourceProcedure, statementID int, call procedureir.CallSite, state map[string]bool, flowContext objectFlowContext, declarations declarationScope, summaries map[string]objectProcedureSummary) bool {
	if objectConstructorCallText(strings.ToLower(call.Callee.Text)) || strings.EqualFold(call.Callee.BaseName, "CreateObject") || strings.EqualFold(call.Callee.BaseName, "GetObject") {
		return true
	}
	if objectCollectionItemAssigned(call, state, declarations) {
		return true
	}
	if objectCollectionMemberItemAssigned(call, state, declarations) {
		return true
	}
	if objectDictionaryItemAssigned(proc, statementID, call, state, declarations) {
		return true
	}
	if objectLateBoundFactoryMemberAssigned(proc, call) {
		return true
	}
	if call.Callee.Receiver == nil {
		if summary, ok := objectDirectCallSummary(proc, call, summaries); ok {
			return objectCallSummaryReturnsAssigned(proc, call, state, flowContext, declarations, summaries, summary)
		}
	}
	if call.Callee.Receiver != nil {
		// Excel object factories such as Worksheets(1) and Range("A1") are
		// non-Nothing object values for this state analysis.  Range.Find is the
		// notable nullable member and is handled by VBA201 as well.  Arbitrary
		// user-object members remain nullable until a procedure summary proves
		// otherwise.
		if strings.EqualFold(call.Callee.Member, "Find") {
			return false
		}
		receiver := strings.ToLower(cleanIdentifier(*call.Callee.Receiver))
		if strings.EqualFold(call.Callee.Member, "Add") && strings.HasSuffix(receiver, ".controls") {
			return true
		}
		if strings.EqualFold(call.Callee.Member, "Item") && strings.HasSuffix(receiver, ".controls") {
			// A valid Controls index returns the existing control object.  The
			// caller's loop bounds are responsible for proving the index; this
			// branch only models the non-Nothing result of the collection lookup.
			return true
		}
		if objectExcelMemberChainAssigned(call, state, declarations) {
			return true
		}
		if !objectErrorResumeNextAt(proc, statementID) && objectExcelMemberFactoryAssigned(call, declarations) {
			// A successful Excel member/factory call produces an object even
			// when its receiver came from a public boundary.  If the receiver
			// were Nothing, VBA would raise before the assignment's normal
			// continuation.  On Error Resume Next is excluded because that
			// mode can continue with an unchanged Nothing target.
			return true
		}
		if receiver == "thisworkbook" || receiver == "application" || receiver == "excel.application" ||
			strings.HasPrefix(receiver, "thisworkbook.") || strings.HasPrefix(receiver, "application.") || strings.HasPrefix(receiver, "excel.application.") {
			return true
		}
		if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
			summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries)
			return ok && objectCallSummaryReturnsAssigned(proc, call, state, flowContext, declarations, summaries, summary)
		}
		var summary objectProcedureSummary
		var ok bool
		if flowContext.receiverSummaryKeys != nil {
			summary, ok = objectReceiverReturnSummaryIndexed(call, declarations, summaries, flowContext.receiverSummaryKeys)
		} else {
			summary, ok = objectReceiverReturnSummary(call, declarations, summaries)
		}
		if ok {
			return objectCallSummaryReturnsAssigned(proc, call, state, flowContext, declarations, summaries, summary)
		}
		return false
	}
	if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
		return false
	}
	summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries)
	return ok && objectCallSummaryReturnsAssigned(proc, call, state, flowContext, declarations, summaries, summary)
}

func objectCallSummaryReturnsAssigned(proc sourceProcedure, call procedureir.CallSite, state map[string]bool, flowContext objectFlowContext, declarations declarationScope, summaries map[string]objectProcedureSummary, summary objectProcedureSummary) bool {
	if summary.ReturnAssigned {
		return true
	}
	if summary.ReturnParameter >= 0 && summary.ReturnParameter < len(summary.Params) {
		if summary.ParamNonNothing[summary.ReturnParameter] {
			return true
		}
		// A return-parameter alias can inherit the caller's pre-call state only
		// when the callee summary proves that its writes preserve that input.
		// A plain ByRefWritten bit is intentionally insufficient: recursive
		// aliases can preserve an initialized object, while `Set value =
		// Nothing` followed by `Set Result = value` must not.
		if summary.ReturnParameterPreservesInput {
			actuals := objectCallActuals(call, flowContext.facts)
			assigned, present := objectCallParameterAssigned(proc, declarations, call, summary, summary.ReturnParameter, actuals, state, flowContext.vars, flowContext, summaries)
			if present && assigned {
				return true
			}
		}
	}
	parameterIndex := summary.ReturnCollectionItemParameter
	if parameterIndex < 0 {
		return false
	}
	actuals := objectCallActuals(call, flowContext.facts)
	assigned, present := objectCallParameterAssigned(proc, declarations, call, summary, parameterIndex, actuals, state, flowContext.vars, flowContext, summaries)
	return present && assigned
}

// objectReturnParameter recognizes the common VBA helper shape where an
// object-returning function returns one of its object parameters on every
// normal path.  The alias matters for calls such as `Set dict = Helper(dict)`:
// the function result is as safe as the actual ByRef object even when the
// callee's standalone entry state must remain conservative.
func objectReturnParameter(plan *objectProcedurePlan) int {
	if plan == nil || !isObjectType(plan.proc.ReturnType) || plan.flowContext.facts == nil || plan.flowGraph.BlockCount() == 0 {
		return -1
	}
	normalExit := plan.flowGraph.NormalExit()
	if !plan.flowGraph.IsReachable(normalExit) {
		return -1
	}
	dominators := plan.flowGraph.Dominators()
	returnName := cleanIdentifier(plan.proc.Name)
	parameterIndex := -1
	sawReturnAssignment := false
	for statement := range plan.proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || statement.Target == nil || statement.Value == nil || !strings.EqualFold(cleanIdentifier(statement.Target.Text), returnName) {
			continue
		}
		valueName := cleanIdentifier(statement.Value.Text)
		if statement.Value.Kind != procedureir.ExpressionIdentifier || valueName == "" {
			return -1
		}
		currentIndex := -1
		for index, parameter := range plan.proc.Params.AllIndexed() {
			if isObjectType(parameter.Type) && strings.EqualFold(cleanIdentifier(parameter.Name), valueName) {
				currentIndex = index
				break
			}
		}
		if currentIndex < 0 {
			return -1
		}
		block, ok := plan.flowGraph.BlockForStatement(statement.ID)
		if !ok || !objectBlockSetContains(dominators[normalExit], block.ID) {
			return -1
		}
		if parameterIndex >= 0 && parameterIndex != currentIndex {
			return -1
		}
		parameterIndex = currentIndex
		sawReturnAssignment = true
	}
	if !sawReturnAssignment {
		return -1
	}
	return parameterIndex
}

func objectReturnParameterPreservesInput(plan *objectProcedurePlan, parameterIndex int, summaries map[string]objectProcedureSummary) bool {
	if plan == nil || parameterIndex < 0 || plan.flowContext.facts == nil {
		return false
	}
	parameterName := ""
	parameterType := ""
	parameterPassing := ""
	for index, parameter := range plan.proc.Params.AllIndexed() {
		if index != parameterIndex {
			continue
		}
		parameterName = parameter.Name
		parameterType = parameter.Type
		parameterPassing = parameter.Passing
		break
	}
	if parameterName == "" || !isObjectType(parameterType) || strings.EqualFold(strings.TrimSpace(parameterPassing), "ByVal") {
		return false
	}
	parameterName = cleanIdentifier(parameterName)
	parameterSummary := objectProcedureSummary{}
	for _, candidate := range plan.proc.Params.AllIndexed() {
		parameterSummary.Params = append(parameterSummary.Params, objectParameterSummary{
			Name:   candidate.Name,
			Object: isObjectType(candidate.Type),
			ByRef:  !strings.EqualFold(strings.TrimSpace(candidate.Passing), "ByVal"),
		})
	}
	preservingCalls := map[int]bool{}
	for statement := range plan.proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || statement.Target == nil || statement.Value == nil || !strings.EqualFold(cleanIdentifier(statement.Target.Text), parameterName) {
			continue
		}
		valueText := strings.TrimSpace(statement.Value.Text)
		valueLower := strings.ToLower(valueText)
		if valueLower == "nothing" {
			return false
		}
		if statement.Value.Kind == procedureir.ExpressionIdentifier && strings.EqualFold(cleanIdentifier(valueText), parameterName) {
			continue
		}
		if objectConstructorCallText(valueLower) || strings.HasPrefix(valueLower, "new ") {
			continue
		}
		preserved := false
		plan.flowContext.facts.forEachCallForStatement(statement.ID, func(call procedureir.CallSite) {
			if objectReturnParameterRecursiveAliasCall(plan.proc, call, parameterIndex, parameterName, parameterSummary, plan.flowContext.facts) {
				preservingCalls[call.ID] = true
				preserved = true
			}
		})
		if !preserved {
			return false
		}
	}
	for call := range plan.proc.Calls.All() {
		for actualIndex, actual := range objectCallActuals(call, plan.flowContext.facts) {
			if actual.parenthesized || !strings.EqualFold(cleanIdentifier(actual.text), parameterName) || preservingCalls[call.ID] {
				continue
			}
			if objectUnresolvedExpressionCallReadOnly(call) {
				continue
			}
			if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
				if summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries); ok {
					formalIndex := objectFormalIndex(call, summary, actualIndex)
					if formalIndex >= 0 && formalIndex < len(summary.Params) {
						formal := summary.Params[formalIndex]
						if !formal.Object || !formal.ByRef || !summary.ByRefWritten[formalIndex] {
							continue
						}
					}
				}
			}
			return false
		}
	}
	return true
}

func objectReturnParameterRecursiveAliasCall(proc sourceProcedure, call procedureir.CallSite, parameterIndex int, parameterName string, summary objectProcedureSummary, facts *procedureAnalysisFacts) bool {
	if call.Callee.Receiver != nil {
		return false
	}
	calleeName := cleanIdentifier(call.Callee.BaseName)
	if calleeName == "" {
		calleeName = objectBareCallName(call.Callee.Text)
	}
	if !strings.EqualFold(calleeName, cleanIdentifier(proc.Name)) {
		return false
	}
	if callerName := strings.TrimSpace(call.Caller.QualifiedName); callerName != "" && !strings.EqualFold(callerName, objectProcedureQualifiedName(proc)) {
		return false
	}
	for actualIndex, actual := range objectCallActuals(call, facts) {
		if actual.parenthesized || !strings.EqualFold(cleanIdentifier(actual.text), parameterName) {
			continue
		}
		if objectFormalIndex(call, summary, actualIndex) == parameterIndex {
			return true
		}
	}
	return false
}

func objectReturnProgID(plan *objectProcedurePlan, summaries map[string]objectProcedureSummary) string {
	if plan == nil || !isObjectType(plan.proc.ReturnType) {
		return ""
	}
	returnName := cleanIdentifier(plan.proc.Name)
	returnValues := make([]string, 0, 1)
	returnStatementIDs := make([]int, 0, 1)
	for statement := range plan.proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || statement.Target == nil || statement.Value == nil || !strings.EqualFold(cleanIdentifier(statement.Target.Text), returnName) {
			continue
		}
		returnValues = append(returnValues, statement.Value.Text)
		returnStatementIDs = append(returnStatementIDs, statement.ID)
	}
	if len(returnValues) == 0 {
		return ""
	}
	progid := ""
	for index, returnValue := range returnValues {
		progIDs := objectAssignedProgIDs(plan, returnValue, returnStatementIDs[index], summaries, map[string]bool{})
		if len(progIDs) != 1 || progid != "" && !strings.EqualFold(progid, progIDs[0]) {
			return ""
		}
		progid = progIDs[0]
	}
	return progid
}

func objectAssignedProgIDs(plan *objectProcedurePlan, valueText string, statementID int, summaries map[string]objectProcedureSummary, visiting map[string]bool) []string {
	if plan == nil {
		return nil
	}
	if progid := objectCreateObjectProgID(valueText, plan.file.ConstantValues); progid != "" {
		return []string{progid}
	}
	trimmed := strings.TrimSpace(valueText)
	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "new ") {
		switch dcKindFromType(strings.TrimSpace(trimmed[len("new "):])) {
		case dcDictionary:
			return []string{"scripting.dictionary"}
		case dcCollection:
			return []string{"vba.collection"}
		default:
			return nil
		}
	}
	name := cleanIdentifier(trimmed)
	if name == "" || strings.ContainsAny(trimmed, " .()") {
		return objectCallReturnProgIDs(plan, valueText, statementID, summaries)
	}
	key := strings.ToLower(name)
	if visiting[key] {
		return nil
	}
	visiting[key] = true
	defer delete(visiting, key)
	values := make([]string, 0, 1)
	for statement := range plan.proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || statement.Target == nil || statement.Value == nil || !strings.EqualFold(cleanIdentifier(statement.Target.Text), name) {
			continue
		}
		current := objectAssignedProgIDs(plan, statement.Value.Text, statement.ID, summaries, visiting)
		if len(current) == 0 {
			return nil
		}
		values = append(values, current...)
	}
	if len(values) == 0 {
		return nil
	}
	for _, value := range values[1:] {
		if !strings.EqualFold(value, values[0]) {
			return nil
		}
	}
	return []string{strings.ToLower(strings.TrimSpace(values[0]))}
}

func objectCallReturnProgIDs(plan *objectProcedurePlan, valueText string, statementID int, summaries map[string]objectProcedureSummary) []string {
	if plan == nil || strings.TrimSpace(valueText) == "" {
		return nil
	}
	name := objectBareCallName(valueText)
	if name == "" {
		return nil
	}
	var progid string
	matched := false
	for call := range plan.proc.Calls.All() {
		if call.StatementID != statementID || call.Callee.Receiver != nil || !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), name) {
			continue
		}
		summary, ok := objectDirectCallSummary(plan.proc, call, summaries)
		if !ok || summary.ReturnProgID == "" {
			return nil
		}
		if matched && !strings.EqualFold(progid, summary.ReturnProgID) {
			return nil
		}
		matched = true
		progid = summary.ReturnProgID
	}
	if !matched {
		return nil
	}
	return []string{strings.ToLower(strings.TrimSpace(progid))}
}

func objectCreateObjectProgID(text string, constants map[string]constexpr.Value) string {
	text = strings.TrimSpace(text)
	lower := strings.ToLower(text)
	const prefix = "createobject("
	if !strings.HasPrefix(lower, prefix) || !strings.HasSuffix(text, ")") {
		return ""
	}
	argument := strings.TrimSpace(text[len(prefix) : len(text)-1])
	if len(argument) >= 2 && argument[0] == '"' && argument[len(argument)-1] == '"' {
		return strings.ToLower(strings.TrimSpace(argument[1 : len(argument)-1]))
	}
	for name, value := range constants {
		if strings.EqualFold(strings.TrimSpace(name), argument) && value.Kind == constexpr.ValueString {
			return strings.ToLower(strings.TrimSpace(value.String))
		}
	}
	return ""
}

func objectReturnCollectionItemParameter(plan *objectProcedurePlan) int {
	if plan == nil || !isObjectType(plan.proc.ReturnType) || plan.flowContext.facts == nil {
		return -1
	}
	returnName := cleanIdentifier(plan.proc.Name)
	collectionParameters := map[string]int{}
	for index, parameter := range plan.proc.Params.AllIndexed() {
		if dcKindFromType(parameter.Type) == dcCollection {
			collectionParameters[strings.ToLower(cleanIdentifier(parameter.Name))] = index
		}
	}
	if len(collectionParameters) == 0 {
		return -1
	}
	if !plan.flowGraph.IsReachable(plan.flowGraph.NormalExit()) {
		return -1
	}
	collectionSources := map[string]int{}
	parameterIndex := -1
	sawReturnAssignment := false
	qualifyingBlocks := map[vbacfg.BlockID]bool{}
	for statement := range plan.proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || statement.Target == nil {
			continue
		}
		targetName := strings.ToLower(cleanIdentifier(statement.Target.Text))
		itemParameter, classified := objectCollectionItemSourceForStatement(plan, statement, collectionParameters, collectionSources)
		if targetName == strings.ToLower(returnName) {
			sawReturnAssignment = true
			if !classified {
				return -1
			}
			if itemParameter >= 0 {
				if parameterIndex >= 0 && parameterIndex != itemParameter {
					return -1
				}
				parameterIndex = itemParameter
				block, ok := plan.flowGraph.BlockForStatement(statement.ID)
				if !ok {
					return -1
				}
				qualifyingBlocks[block.ID] = true
			}
		}
		if itemParameter >= 0 {
			collectionSources[targetName] = itemParameter
		} else {
			delete(collectionSources, targetName)
		}
	}
	if !sawReturnAssignment {
		return -1
	}
	if parameterIndex < 0 || !objectNormalExitCoveredByBlocks(plan.flowGraph, qualifyingBlocks) {
		return -1
	}
	return parameterIndex
}

func objectNormalExitCoveredByBlocks(graph vbacfg.CFGView, qualifying map[vbacfg.BlockID]bool) bool {
	if len(qualifying) == 0 {
		return false
	}
	seen := map[vbacfg.BlockID]bool{graph.Entry(): true}
	queue := []vbacfg.BlockID{graph.Entry()}
	normalExit := graph.NormalExit()
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if current != graph.Entry() && qualifying[current] {
			continue
		}
		if current == normalExit {
			return false
		}
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if !seen[edge.To] {
				seen[edge.To] = true
				queue = append(queue, edge.To)
			}
			return true
		})
	}
	return true
}

func objectCollectionItemSourceForStatement(plan *objectProcedurePlan, statement procedureir.Statement, collectionParameters, collectionSources map[string]int) (int, bool) {
	itemParameter := -1
	plan.flowContext.facts.forEachCallForStatement(statement.ID, func(call procedureir.CallSite) {
		if itemParameter >= 0 || call.Callee.Receiver != nil || call.Arguments.Count == 0 {
			return
		}
		if index, ok := collectionParameters[strings.ToLower(cleanIdentifier(call.Callee.BaseName))]; ok {
			itemParameter = index
		}
	})
	if itemParameter >= 0 {
		return itemParameter, true
	}
	lower := ""
	if statement.Value != nil {
		lower = strings.ToLower(strings.TrimSpace(statement.Value.Text))
	}
	if objectConstructorCallText(lower) || strings.HasPrefix(lower, "new ") {
		return -1, true
	}
	if statement.Value != nil && statement.Value.Kind == procedureir.ExpressionIdentifier {
		if source, ok := collectionSources[strings.ToLower(cleanIdentifier(statement.Value.Text))]; ok {
			return source, true
		}
	}
	return -1, false
}

func objectCollectionItemAssigned(call procedureir.CallSite, state map[string]bool, declarations declarationScope) bool {
	if call.Callee.Receiver != nil || call.Arguments.Count == 0 || strings.TrimSpace(call.Callee.BaseName) == "" {
		return false
	}
	name := cleanIdentifier(call.Callee.BaseName)
	declaration, scope, ok := objectDeclarationBinding(name, declarations)
	if !ok || !declaration.Object || dcKindFromType(declaration.Type) != dcCollection {
		return false
	}
	return state[(objectVariable{Scope: scope, Name: name}).key()]
}

func objectCollectionMemberItemAssigned(call procedureir.CallSite, state map[string]bool, declarations declarationScope) bool {
	if call.Callee.Receiver == nil || !strings.EqualFold(cleanIdentifier(call.Callee.Member), "item") {
		return false
	}
	receiver := strings.TrimSpace(*call.Callee.Receiver)
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(strings.SplitN(receiver, ".", 2)[0], "(", 2)[0]))
	if root == "" {
		return false
	}
	declaration, scope, ok := objectDeclarationBinding(root, declarations)
	if !ok || !declaration.Object || dcKindFromType(declaration.Type) != dcCollection {
		return false
	}
	return state[(objectVariable{Scope: scope, Name: root}).key()]
}

func objectDictionaryItemAssigned(proc sourceProcedure, statementID int, call procedureir.CallSite, state map[string]bool, declarations declarationScope) bool {
	if call.Callee.Receiver != nil || call.Arguments.Count == 0 || strings.TrimSpace(call.Callee.BaseName) == "" {
		return false
	}
	name := cleanIdentifier(call.Callee.BaseName)
	declaration, scope, ok := objectDeclarationBinding(name, declarations)
	if !ok || !declaration.Object || dcKindFromType(declaration.Type) != dcDictionary && !strings.EqualFold(cleanIdentifier(declaration.Type), "object") {
		return false
	}
	if !state[(objectVariable{Scope: scope, Name: name}).key()] || !objectDictionaryItemGuarded(proc, statementID, name) {
		return false
	}
	return objectDictionaryItemTargetIsCollection(proc, statementID, declarations)
}

func objectDictionaryItemTargetIsCollection(proc sourceProcedure, statementID int, declarations declarationScope) bool {
	for statement := range proc.Statements.All() {
		if statement.ID != statementID || statement.Target == nil {
			continue
		}
		name := cleanIdentifier(statement.Target.Text)
		if declaration, _, ok := objectDeclarationBinding(name, declarations); ok && dcKindFromType(declaration.Type) == dcCollection {
			return true
		}
		return strings.EqualFold(name, cleanIdentifier(proc.Name)) && dcKindFromType(proc.ReturnType) == dcCollection
	}
	return false
}

func objectDictionaryItemGuarded(proc sourceProcedure, statementID int, receiver string) bool {
	for statement := range proc.Statements.All() {
		if statement.ID != statementID {
			continue
		}
		guardReceiver, keyText, ok := dcDefaultAccess(statement.Text)
		if !ok || !strings.EqualFold(guardReceiver, receiver) {
			return false
		}
		return dcAccessGuarded(proc, statement, receiver, keyText)
	}
	return false
}

func objectLateBoundFactoryMemberAssigned(proc sourceProcedure, call procedureir.CallSite) bool {
	receiver := objectCallWithReceiverName(proc, call)
	if receiver == "" || strings.Contains(receiver, ".") {
		return false
	}
	progid := objectCreateObjectReceiverProgID(proc, receiver, call.StatementID)
	if !strings.EqualFold(progid, "scripting.filesystemobject") {
		return false
	}
	if objectErrorResumeNextAt(proc, call.StatementID) {
		return false
	}
	switch strings.ToLower(cleanIdentifier(call.Callee.Member)) {
	case "opentextfile", "createtextfile", "getfile", "getfolder":
		return true
	default:
		return false
	}
}

func objectCreateObjectReceiverProgID(proc sourceProcedure, receiver string, statementID int) string {
	if proc.Graph == nil {
		return ""
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	callBlock, ok := graph.BlockForStatement(statementID)
	if !ok {
		return ""
	}
	dominators := graph.Dominators()
	assignmentCount := 0
	progid := ""
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || statement.Target == nil || !strings.EqualFold(cleanIdentifier(statement.Target.Text), receiver) || statement.Value == nil {
			continue
		}
		assignmentBlock, ok := graph.BlockForStatement(statement.ID)
		if !ok {
			return ""
		}
		if !objectBlockCanReach(graph, assignmentBlock.ID, callBlock.ID) {
			continue
		}
		if !objectBlockSetContains(dominators[callBlock.ID], assignmentBlock.ID) || objectErrorResumeNextAt(proc, statement.ID) {
			return ""
		}
		assignmentCount++
		text := strings.ToLower(strings.TrimSpace(statement.Value.Text))
		const prefix = "createobject("
		if !strings.HasPrefix(text, prefix) || !strings.HasSuffix(text, ")") {
			return ""
		}
		argument := strings.TrimSpace(text[len(prefix) : len(text)-1])
		if len(argument) < 2 || argument[0] != '"' || argument[len(argument)-1] != '"' {
			continue
		}
		current := strings.TrimSpace(argument[1 : len(argument)-1])
		if progid != "" && !strings.EqualFold(progid, current) {
			return ""
		}
		progid = current
	}
	if assignmentCount != 1 {
		return ""
	}
	return progid
}

func objectBlockCanReach(graph vbacfg.CFGView, from, target vbacfg.BlockID) bool {
	if from == target {
		return true
	}
	seen := map[vbacfg.BlockID]bool{from: true}
	queue := []vbacfg.BlockID{from}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		graph.ForEachOutgoing(current, func(edge vbacfg.Edge) bool {
			if edge.To == target {
				seen[target] = true
				return false
			}
			if !seen[edge.To] {
				seen[edge.To] = true
				queue = append(queue, edge.To)
			}
			return true
		})
		if seen[target] {
			return true
		}
	}
	return false
}

func objectReceiverFactoryProgID(proc sourceProcedure, receiver string, statementID int, summaries map[string]objectProcedureSummary) string {
	if direct := objectCreateObjectReceiverProgID(proc, receiver, statementID); direct != "" {
		return direct
	}
	if proc.Graph == nil {
		return ""
	}
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	callBlock, ok := graph.BlockForStatement(statementID)
	if !ok {
		return ""
	}
	dominators := graph.Dominators()
	assignmentCount := 0
	dominatingAssignmentCount := 0
	progid := ""
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementSet || statement.Target == nil || statement.Value == nil || !strings.EqualFold(cleanIdentifier(statement.Target.Text), receiver) {
			continue
		}
		assignmentBlock, ok := graph.BlockForStatement(statement.ID)
		if !ok {
			return ""
		}
		if !objectBlockCanReach(graph, assignmentBlock.ID, callBlock.ID) {
			continue
		}
		if objectErrorResumeNextAt(proc, statement.ID) {
			return ""
		}
		assignmentCount++
		if objectBlockSetContains(dominators[callBlock.ID], assignmentBlock.ID) {
			dominatingAssignmentCount++
		}
		current := ""
		text := strings.ToLower(strings.TrimSpace(statement.Value.Text))
		const prefix = "createobject("
		if strings.HasPrefix(text, prefix) && strings.HasSuffix(text, ")") {
			argument := strings.TrimSpace(text[len(prefix) : len(text)-1])
			if len(argument) >= 2 && argument[0] == '"' && argument[len(argument)-1] == '"' {
				current = strings.TrimSpace(argument[1 : len(argument)-1])
			}
		}
		if current == "" && strings.HasPrefix(text, "new ") && dcKindFromType(strings.TrimSpace(text[len("new "):])) == dcDictionary {
			current = "scripting.dictionary"
		}
		for call := range proc.Calls.All() {
			if call.StatementID != statement.ID || call.Callee.Receiver != nil {
				continue
			}
			summary, summaryOK := objectDirectCallSummary(proc, call, summaries)
			if summaryOK && summary.ReturnProgID != "" {
				if current != "" && !strings.EqualFold(current, summary.ReturnProgID) {
					return ""
				}
				current = summary.ReturnProgID
			}
		}
		if current == "" || progid != "" && !strings.EqualFold(progid, current) {
			return ""
		}
		progid = current
	}
	if assignmentCount == 0 || dominatingAssignmentCount == 0 {
		return ""
	}
	return progid
}

func objectDirectCallSummary(proc sourceProcedure, call procedureir.CallSite, summaries map[string]objectProcedureSummary) (objectProcedureSummary, bool) {
	if (call.Resolution.Status != procedureir.ResolutionMatched && !objectResolutionAllowsDirectSummary(call.Resolution.Status)) ||
		call.Callee.Receiver != nil || strings.TrimSpace(call.Callee.BaseName) == "" {
		return objectProcedureSummary{}, false
	}
	matches := objectSummaryCandidatesForDirectCall(proc.Module, call.Callee.BaseName, call.File, summaries)
	if len(matches) == 0 {
		return objectProcedureSummary{}, false
	}
	if len(matches) == 1 {
		return matches[0], true
	}
	returnAssigned := true
	for _, summary := range matches {
		returnAssigned = returnAssigned && summary.ReturnAssigned
	}
	if !returnAssigned {
		return objectProcedureSummary{}, false
	}
	return objectProcedureSummary{ReturnAssigned: true, ReturnParameter: -1, ReturnCollectionItemParameter: -1}, true
}

func objectSummaryCandidatesForDirectCall(module, name, callFile string, summaries map[string]objectProcedureSummary) []objectProcedureSummary {
	name = strings.TrimSpace(name)
	if strings.TrimSpace(module) == "" || name == "" {
		return nil
	}
	allMatches := make([]objectProcedureSummary, 0, 1)
	for _, summary := range summaries {
		if !strings.EqualFold(strings.TrimSpace(summary.Module), strings.TrimSpace(module)) || !strings.EqualFold(lastName(summary.QualifiedName), name) {
			continue
		}
		allMatches = append(allMatches, summary)
	}
	if len(allMatches) == 0 {
		return nil
	}
	// A project may contain a WIP copy of a module with the same VB_Name. The
	// call site file identifies the lexical function in that case; do not let
	// an unrelated duplicate affect object state. If the file is unavailable,
	// only a unique module/name match is safe to use.
	if normalizedCallFile := normalizedObjectSummaryFile(callFile); normalizedCallFile != "" {
		fileMatches := make([]objectProcedureSummary, 0, len(allMatches))
		for _, summary := range allMatches {
			if objectSummaryFilesMatch(normalizedCallFile, summary.File) {
				fileMatches = append(fileMatches, summary)
			}
		}
		if len(fileMatches) > 0 {
			return fileMatches
		}
	}
	if len(allMatches) == 1 {
		return allMatches
	}
	return nil
}

func normalizedObjectSummaryFile(file string) string {
	file = filepath.ToSlash(filepath.Clean(strings.TrimSpace(file)))
	file = strings.TrimPrefix(file, "./")
	if file == "." {
		return ""
	}
	return file
}

func objectSummaryFilesMatch(callFile, summaryFile string) bool {
	callFile = normalizedObjectSummaryFile(callFile)
	summaryFile = normalizedObjectSummaryFile(summaryFile)
	if callFile == "" || summaryFile == "" {
		return false
	}
	if strings.EqualFold(callFile, summaryFile) {
		return true
	}
	callFile = strings.ToLower(callFile)
	summaryFile = strings.ToLower(summaryFile)
	return strings.HasSuffix(callFile, "/"+summaryFile) || strings.HasSuffix(summaryFile, "/"+callFile)
}

func objectReceiverReturnSummary(call procedureir.CallSite, declarations declarationScope, summaries map[string]objectProcedureSummary) (objectProcedureSummary, bool) {
	return objectReceiverReturnSummaryIndexed(call, declarations, summaries, nil)
}

func objectReceiverReturnSummaryIndexed(call procedureir.CallSite, declarations declarationScope, summaries map[string]objectProcedureSummary, receiverSummaryKeys map[string][]string) (objectProcedureSummary, bool) {
	if call.Callee.Receiver == nil || call.Callee.Member == "" {
		return objectProcedureSummary{}, false
	}
	receiver := strings.TrimSpace(*call.Callee.Receiver)
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(strings.SplitN(receiver, ".", 2)[0], "(", 2)[0]))
	declaration, ok := objectDeclarationByName(root, declarations)
	if !ok || strings.TrimSpace(declaration.Type) == "" {
		return objectProcedureSummary{}, false
	}
	typeName := strings.ToLower(cleanIdentifier(lastName(strings.TrimSpace(declaration.Type))))
	member := strings.ToLower(cleanIdentifier(call.Callee.Member))
	if receiverSummaryKeys != nil {
		keys := receiverSummaryKeys[objectReceiverSummaryIndexKey(typeName, member)]
		if len(keys) != 1 {
			return objectProcedureSummary{}, false
		}
		match, ok := summaries[keys[0]]
		return match, ok
	}
	var match objectProcedureSummary
	found := false
	for _, summary := range summaries {
		if !strings.EqualFold(cleanIdentifier(summary.Module), typeName) || !strings.EqualFold(lastName(summary.QualifiedName), member) {
			continue
		}
		if found {
			return objectProcedureSummary{}, false
		}
		match, found = summary, true
	}
	return match, found
}

func objectExcelMemberChainAssigned(call procedureir.CallSite, state map[string]bool, declarations declarationScope) bool {
	if call.Callee.Receiver == nil {
		return false
	}
	receiver := strings.TrimSpace(*call.Callee.Receiver)
	parts := strings.Split(receiver, ".")
	if len(parts) == 0 {
		return false
	}
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(parts[0], "(", 2)[0]))
	if root == "" {
		return false
	}
	declaration, scope, ok := objectDeclarationBinding(root, declarations)
	if !ok || !excelObjectUseType(declaration.Type) {
		return false
	}
	if !state[(objectVariable{Scope: scope, Name: root}).key()] {
		return false
	}
	for _, part := range parts[1:] {
		if !excelObjectUseMember(part) {
			return false
		}
	}
	return excelObjectUseMember(call.Callee.Member)
}

func objectExcelMemberExpressionAssigned(text string, proc sourceProcedure, declarations declarationScope) bool {
	parts := strings.Split(strings.TrimSpace(text), ".")
	if len(parts) < 2 {
		return false
	}
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(parts[0], "(", 2)[0]))
	if root == "" {
		return false
	}
	declaration, _, ok := objectDeclarationBinding(root, declarations)
	if !ok || !excelObjectUseType(declaration.Type) {
		for irDeclaration := range proc.Declarations.All() {
			if strings.EqualFold(cleanIdentifier(irDeclaration.Name), root) && isObjectType(irDeclaration.Type) {
				declaration = sourceDeclaration{Type: irDeclaration.Type}
				ok = true
				break
			}
		}
	}
	if !ok || !excelObjectUseType(declaration.Type) {
		return false
	}
	for _, part := range parts[1:] {
		if !excelObjectFactoryMember(part) {
			return false
		}
	}
	return true
}

func objectExcelMemberFactoryAssigned(call procedureir.CallSite, declarations declarationScope) bool {
	if call.Callee.Receiver == nil {
		return false
	}
	receiver := strings.TrimSpace(*call.Callee.Receiver)
	parts := strings.Split(receiver, ".")
	if len(parts) == 0 {
		return false
	}
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(parts[0], "(", 2)[0]))
	declaration, ok := objectDeclarationByName(root, declarations)
	if !ok || !excelObjectUseType(declaration.Type) {
		return false
	}
	for _, part := range parts[1:] {
		if !excelObjectFactoryMember(part) {
			return false
		}
	}
	return excelObjectFactoryMember(call.Callee.Member)
}

func objectDeclarationByName(name string, declarations declarationScope) (sourceDeclaration, bool) {
	declaration, _, ok := objectDeclarationBinding(name, declarations)
	return declaration, ok
}

func excelObjectUseType(typ string) bool {
	typ = strings.ToLower(cleanIdentifier(strings.TrimSpace(typ)))
	switch typ {
	case "application", "workbook", "worksheet", "range", "chart", "chartobject", "series", "pivot table", "pivottable", "listobject", "window", "shape":
		return true
	default:
		return strings.HasSuffix(typ, ".application") || strings.HasSuffix(typ, ".workbook") || strings.HasSuffix(typ, ".worksheet") || strings.HasSuffix(typ, ".range") || strings.HasSuffix(typ, ".shape")
	}
}

func excelObjectUseMember(member string) bool {
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(strings.SplitN(member, "(", 2)[0])))
	switch member {
	case "application", "workbooks", "worksheets", "sheets", "range", "cells", "rows", "columns", "shapes", "chartobjects", "chart", "seriescollection", "newseries", "parent", "resize", "offset", "addshape", "selection", "usedrange", "interior", "borders", "font", "textframe", "characters", "fill", "line", "controls", "add":
		return true
	default:
		return false
	}
}

func excelObjectFactoryMember(member string) bool {
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(strings.SplitN(member, "(", 2)[0])))
	switch member {
	case "application", "workbooks", "worksheets", "sheets", "range", "cells", "rows", "columns", "shapes", "chartobjects", "chart", "seriescollection", "newseries", "usedrange", "interior", "borders", "font", "textframe", "characters", "fill", "line", "controls", "add", "addshape", "parent", "resize", "offset":
		return true
	default:
		return false
	}
}

func applyObjectCallEffects(proc sourceProcedure, call procedureir.CallSite, state map[string]bool, vars map[string]objectVariable, declarations declarationScope, facts *procedureAnalysisFacts, summaries map[string]objectProcedureSummary) {
	if objectExternalCallPreservesObjectArguments(proc, call, summaries) {
		// ADODB.Stream.CopyTo consumes its destination stream but does not
		// replace the caller's object variable.  Handle the late-bound With
		// form before the unresolved-expression guard below, because the IR may
		// retain the implicit receiver as an expression call.
		return
	}
	actuals := objectCallActuals(call, facts)
	if objectAddCallPreservesContainerArguments(proc, call, actuals, declarations, summaries) || objectExternalCallPreservesObjectArguments(proc, call, summaries) {
		// Collection/Dictionary Add accepts its item by value; the built-in
		// operation cannot replace a container argument with Nothing.  A late-
		// bound Object is included here because this is the normal VBA shape for
		// Scripting.Dictionary values.  Resolved project procedures remain
		// conservative and use their actual ByRef summaries.
		actuals = nil
	}
	var candidates []objectProcedureSummary
	if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
		if summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries); ok {
			if !strings.EqualFold(strings.TrimSpace(call.Caller.QualifiedName), strings.TrimSpace(summary.QualifiedName)) || objectRecursiveSummaryHasByValObjectActual(call, summary, actuals) {
				// Recursive summaries are useful only when they carry a ByVal
				// object reference.  A recursive ByRef edge remains an unknown
				// mutation boundary and must stay conservative.
				candidates = append(candidates, summary)
			}
		}
	} else if objectResolutionAllowsDirectSummary(call.Resolution.Status) && (call.Callee.Receiver == nil || objectIsCurrentModuleReceiver(call)) {
		name := call.Callee.BaseName
		if strings.TrimSpace(name) == "" {
			name = call.Callee.Member
		}
		matches := objectSummaryCandidatesForDirectCall(call.Module, name, call.File, summaries)
		if len(matches) != 1 {
			matches = nil
		}
		for _, summary := range matches {
			if !strings.EqualFold(strings.TrimSpace(call.Caller.QualifiedName), strings.TrimSpace(summary.QualifiedName)) || objectRecursiveSummaryHasByValObjectActual(call, summary, actuals) {
				// Preserve the recursive procedure summary so its ByVal object
				// parameters retain the caller's state across the recursive edge.
				candidates = append(candidates, summary)
			}
		}
	}
	// A small set of unresolved intrinsic expressions are read-only value
	// queries.  Other unresolved expressions may be external functions with
	// ByRef object arguments, so they must invalidate those arguments just like
	// an unresolved statement-level call.  A unique same-module summary was
	// applied above even when the resolver did not attach a candidate.
	if call.Resolution.Status != procedureir.ResolutionMatched && call.ExpressionID != 0 && len(candidates) == 0 && objectUnresolvedExpressionCallReadOnly(call) {
		return
	}
	// Module fields are represented without a module qualifier in VariableAccess.
	// Apply a callee's field summary only for calls within the same module, where
	// that unqualified binding is unambiguous; cross-module qualified fields stay
	// conservative rather than mutating an unrelated same-named field.
	for _, summary := range candidates {
		if !strings.EqualFold(strings.TrimSpace(summary.Module), strings.TrimSpace(call.Module)) {
			continue
		}
		for name, assigned := range summary.ModuleAssigned {
			if !summary.ModuleWritten[name] {
				continue
			}
			binding, scope, ok := objectDeclarationBinding(name, declarations)
			if !ok || scope != procedureir.ScopeModule || !binding.Object {
				continue
			}
			variable := objectVariable{Scope: procedureir.ScopeModule, Name: name}
			if _, exists := vars[variable.key()]; exists {
				state[variable.key()] = assigned
			}
		}
	}
	if len(actuals) == 0 {
		return
	}
	for actualIndex, actual := range actuals {
		name := cleanIdentifier(actual.text)
		if name == "" {
			continue
		}
		declaration, scope, ok := objectDeclarationBinding(name, declarations)
		if !ok || !declaration.Object {
			continue
		}
		variable := objectVariable{Scope: scope, Name: name}
		if _, exists := vars[variable.key()]; !exists {
			continue
		}
		if actual.parenthesized && len(candidates) == 0 {
			// An unresolved call receives a temporary ByVal value for a
			// parenthesized actual, so it cannot change the caller's object
			// reference.
			continue
		}
		assigned := false
		known := false
		for _, summary := range candidates {
			formalIndex := objectFormalIndex(call, summary, actualIndex)
			if formalIndex < 0 || formalIndex >= len(summary.Params) || !summary.Params[formalIndex].Object {
				continue
			}
			known = true
			if strings.EqualFold(strings.TrimSpace(call.Caller.QualifiedName), strings.TrimSpace(summary.QualifiedName)) && summary.Params[formalIndex].ByRef {
				// A recursive ByRef call may mutate the caller's object
				// reference, so do not use the recursive summary to prove it
				// remains assigned.
				assigned = false
				break
			}
			if actual.parenthesized {
				// Parenthesized actuals are passed through a temporary ByVal
				// value even when the formal parameter is ByRef.  The callee's
				// normal-return postcondition therefore cannot establish the
				// caller's original binding.
				assigned = state[variable.key()]
				continue
			}
			if summary.ParamNonNothing[formalIndex] {
				assigned = true
				continue
			}
			if !summary.Params[formalIndex].ByRef {
				// ByVal receives a copy and cannot alter the caller's object
				// reference; preserve its existing state.
				assigned = state[variable.key()]
				continue
			}
			if !summary.ByRefWritten[formalIndex] {
				// A ByRef parameter that is only read remains an alias to the
				// caller's object.  Lack of a guaranteed assignment is not the
				// same as evidence that the callee wrote Nothing into it.
				assigned = state[variable.key()]
				continue
			}
			if !summary.ByRefAssigned[formalIndex] {
				assigned = false
				break
			}
			assigned = true
		}
		if len(candidates) == 0 {
			// An unresolved/external call may mutate a ByRef object or leave it
			// Nothing; no definite state survives the call.
			state[variable.key()] = false
			continue
		}
		if !known {
			// A resolved non-object formal does not mutate this object reference.
			continue
		}
		state[variable.key()] = assigned
	}
}

func objectUnresolvedExpressionCallReadOnly(call procedureir.CallSite) bool {
	if call.Resolution.Status == procedureir.ResolutionMatched || call.ExpressionID == 0 {
		return false
	}
	name := strings.ToLower(cleanIdentifier(call.Callee.BaseName))
	if name == "" {
		name = strings.ToLower(objectBareCallName(call.Callee.Text))
	}
	switch name {
	case "typename", "strcomp", "isobject", "array":
		return true
	default:
		return false
	}
}

func objectAddCallPreservesContainerArguments(proc sourceProcedure, call procedureir.CallSite, actuals []objectCallActual, declarations declarationScope, summaries map[string]objectProcedureSummary) bool {
	if call.Callee.Receiver == nil || !strings.EqualFold(cleanIdentifier(call.Callee.Member), "add") {
		return false
	}
	receiver := objectCallWithReceiverName(proc, call)
	if receiver == "" || strings.Contains(receiver, ".") {
		return false
	}
	receiverDeclaration, _, ok := objectDeclarationBinding(receiver, declarations)
	if !ok || !receiverDeclaration.Object {
		return false
	}
	receiverKind := dcKindFromType(receiverDeclaration.Type)
	receiverIsContainer := receiverKind == dcCollection || receiverKind == dcDictionary
	if !receiverIsContainer && strings.EqualFold(cleanIdentifier(receiverDeclaration.Type), "object") {
		receiverProgID := objectReceiverFactoryProgID(proc, receiver, call.StatementID, summaries)
		receiverIsContainer = strings.EqualFold(receiverProgID, "scripting.dictionary")
		if !receiverIsContainer && receiverDeclaration.Parameter {
			if summary, ok := objectSummaryForProcedure(proc, summaries); ok {
				for index, parameter := range summary.Params {
					if strings.EqualFold(cleanIdentifier(parameter.Name), receiver) && strings.EqualFold(summary.ParamProgID[index], "scripting.dictionary") {
						receiverIsContainer = true
						break
					}
				}
			}
		}
	}
	if !receiverIsContainer {
		return false
	}
	objectArgument := false
	for _, actual := range actuals {
		name := cleanIdentifier(actual.text)
		if name == "" {
			continue
		}
		declaration, _, ok := objectDeclarationBinding(name, declarations)
		if !ok || !declaration.Object {
			continue
		}
		objectArgument = true
	}
	return objectArgument
}

func objectExternalCallPreservesObjectArguments(proc sourceProcedure, call procedureir.CallSite, summaries map[string]objectProcedureSummary) bool {
	// This Windows API declaration takes the object as ByVal.  The call can
	// inspect the COM interface, but it cannot clear the caller's object
	// reference, so preserve the definite-assignment state across it.
	if strings.EqualFold(cleanIdentifier(call.Callee.BaseName), "iunknown_getwindow") {
		return true
	}
	if !strings.EqualFold(cleanIdentifier(call.Callee.Member), "copyto") && !strings.Contains(strings.ToLower(call.Callee.Text), "copyto") {
		return false
	}
	receiver := objectCallWithReceiverName(proc, call)
	if receiver == "" || strings.Contains(receiver, ".") {
		return false
	}
	progid := objectReceiverFactoryProgID(proc, receiver, call.StatementID, summaries)
	return strings.EqualFold(progid, "adodb.stream")
}

func objectCallWithReceiverName(proc sourceProcedure, call procedureir.CallSite) string {
	if call.Callee.Receiver != nil {
		receiver := cleanIdentifier(strings.TrimSpace(*call.Callee.Receiver))
		if receiver != "" && receiver != "." {
			return receiver
		}
	}
	if !strings.HasPrefix(strings.TrimSpace(call.Callee.Text), ".") {
		return ""
	}
	withStack := []string{}
	for statement := range proc.Statements.All() {
		if statement.ID == call.StatementID {
			if len(withStack) == 0 {
				return ""
			}
			return cleanIdentifier(withStack[len(withStack)-1])
		}
		switch statement.Kind {
		case procedureir.StatementWith:
			text := strings.TrimSpace(statement.Text)
			if len(text) >= len("with ") && strings.EqualFold(text[:len("with ")], "with ") {
				expression := strings.TrimSpace(text[len("with "):])
				if newline := strings.IndexAny(expression, "\r\n"); newline >= 0 {
					expression = strings.TrimSpace(expression[:newline])
				}
				withStack = append(withStack, expression)
			}
		case procedureir.StatementEnd:
			if strings.HasPrefix(strings.ToLower(strings.TrimSpace(statement.Text)), "end with") && len(withStack) > 0 {
				withStack = withStack[:len(withStack)-1]
			}
		}
	}
	return ""
}

func objectRecursiveSummaryHasByValObjectActual(call procedureir.CallSite, summary objectProcedureSummary, actuals []objectCallActual) bool {
	if !strings.EqualFold(strings.TrimSpace(call.Caller.QualifiedName), strings.TrimSpace(summary.QualifiedName)) {
		return false
	}
	for actualIndex := range actuals {
		formalIndex := objectFormalIndex(call, summary, actualIndex)
		if formalIndex < 0 || formalIndex >= len(summary.Params) {
			continue
		}
		parameter := summary.Params[formalIndex]
		if parameter.Object && !parameter.ByRef {
			return true
		}
	}
	return false
}

type objectCallActual struct {
	expressionID  int
	text          string
	parenthesized bool
}

func objectProcedureAllowsParameterEntry(proc sourceProcedure) bool {
	return strings.EqualFold(strings.TrimSpace(proc.Visibility), "private")
}

func objectCallParameterAssigned(proc sourceProcedure, declarations declarationScope, call procedureir.CallSite, summary objectProcedureSummary, formalIndex int, actuals []objectCallActual, state map[string]bool, vars map[string]objectVariable, flowContext objectFlowContext, summaries map[string]objectProcedureSummary) (bool, bool) {
	for actualIndex, actual := range actuals {
		if objectFormalIndex(call, summary, actualIndex) != formalIndex {
			continue
		}
		if actual.expressionID != 0 {
			if expression, ok := flowContext.facts.Expression(actual.expressionID); ok && expression.Kind != procedureir.ExpressionIdentifier {
				return objectExpressionAssigned(proc, expression, state, flowContext, declarations, summaries, call.StatementID), true
			}
		}
		name := cleanIdentifier(actual.text)
		if name == "" {
			return false, true
		}
		if objectIntrinsicIdentifierAssigned(proc, name) {
			return true, true
		}
		declaration, scope, ok := objectDeclarationBinding(name, declarations)
		if !ok && proc.Name != "" && strings.EqualFold(name, cleanIdentifier(proc.Name)) && isObjectType(proc.ReturnType) {
			// A function's return slot is an object binding even though it is not
			// represented by a source Dim declaration. Keep it aligned with the
			// special return-slot handling in objectFlowTarget.
			variable := objectVariable{Scope: procedureir.ScopeLocal, Name: name}
			if _, exists := vars[variable.key()]; exists {
				return state[variable.key()], true
			}
		}
		if !ok || !declaration.Object {
			return false, false
		}
		variable := objectVariable{Scope: scope, Name: name}
		if _, exists := vars[variable.key()]; !exists {
			return false, false
		}
		return state[variable.key()], true
	}
	return false, false
}

func objectCallActuals(call procedureir.CallSite, facts *procedureAnalysisFacts) []objectCallActual {
	actuals := make([]objectCallActual, 0, len(call.Arguments.ExpressionIDs))
	if facts == nil {
		return actuals
	}
	for _, id := range call.Arguments.ExpressionIDs {
		expression, ok := facts.Expression(id)
		if !ok {
			// Preserve the positional slot.  Dropping a missing expression ID
			// would shift every following actual onto the wrong formal parameter.
			actuals = append(actuals, objectCallActual{})
			continue
		}
		parenthesized := false
		for expression.Kind == procedureir.ExpressionParentheses && len(expression.Children) == 1 {
			parenthesized = true
			nested, nestedOK := facts.Expression(expression.Children[0])
			if !nestedOK {
				break
			}
			expression = nested
		}
		actuals = append(actuals, objectCallActual{expressionID: expression.ID, text: expression.Text, parenthesized: parenthesized})
	}
	return actuals
}

func objectFormalIndex(call procedureir.CallSite, summary objectProcedureSummary, actualIndex int) int {
	// Named arguments carry their own expression IDs.  Positional arguments
	// retain source order in ExpressionIDs; map all non-named arguments first.
	for _, named := range call.Arguments.Named {
		if actualIndex < len(call.Arguments.ExpressionIDs) && named.ExpressionID != 0 && named.ExpressionID == call.Arguments.ExpressionIDs[actualIndex] {
			for index, parameter := range summary.Params {
				if strings.EqualFold(parameter.Name, named.Name) {
					return index
				}
			}
		}
	}
	unnamed := 0
	namedIDs := map[int]bool{}
	for _, named := range call.Arguments.Named {
		if named.ExpressionID != 0 {
			namedIDs[named.ExpressionID] = true
		}
	}
	for index, id := range call.Arguments.ExpressionIDs {
		if namedIDs[id] {
			continue
		}
		if index == actualIndex {
			return unnamed
		}
		unnamed++
	}
	return actualIndex
}

func objectSummaryForCandidate(candidate procedureir.Candidate, summaries map[string]objectProcedureSummary) (objectProcedureSummary, bool) {
	key := objectSummaryKey(candidate.File, candidate.QualifiedName, candidate.Kind, candidate.Line)
	if summary, ok := summaries[key]; ok {
		return summary, true
	}
	// Candidate paths can be relative while the parsed procedure path is
	// absolute (and synthetic IR may omit the path entirely).  Preserve the
	// uniqueness guarantees from resolution by falling back only when the
	// qualified name, kind, and declaration line identify exactly one summary.
	var match objectProcedureSummary
	found := false
	for _, summary := range summaries {
		if !strings.EqualFold(summary.QualifiedName, candidate.QualifiedName) {
			continue
		}
		if candidate.Kind != "" && !strings.EqualFold(summary.Kind, candidate.Kind) {
			continue
		}
		if candidate.Line > 0 && summary.Line != candidate.Line {
			continue
		}
		if found {
			return objectProcedureSummary{}, false
		}
		match, found = summary, true
	}
	if found {
		return match, true
	}
	return objectProcedureSummary{}, false
}

func objectSummaryForProcedure(proc sourceProcedure, summaries map[string]objectProcedureSummary) (objectProcedureSummary, bool) {
	qualified := objectProcedureQualifiedName(proc)
	var match objectProcedureSummary
	found := false
	for _, summary := range summaries {
		if !strings.EqualFold(summary.QualifiedName, qualified) {
			continue
		}
		if found {
			return objectProcedureSummary{}, false
		}
		match = summary
		found = true
	}
	return match, found
}

func objectFlowForEachZeroIteration(flowContext objectFlowContext, edge vbacfg.Edge) bool {
	if edge.Kind != vbacfg.EdgeLoopExit {
		return false
	}
	statement, ok := flowContext.facts.Statement(edge.StatementID)
	return ok && statement.Kind == procedureir.StatementForEach
}

func cloneObjectState(in map[string]bool) map[string]bool {
	out := make(map[string]bool, len(in))
	for key, value := range in {
		out[key] = value
	}
	return out
}

func objectStateAllTrue(vars map[string]objectVariable) map[string]bool {
	out := make(map[string]bool, len(vars))
	for key := range vars {
		out[key] = true
	}
	return out
}

func objectFlowIntersection(states []map[string]bool, vars map[string]objectVariable) map[string]bool {
	out := objectStateAllTrue(vars)
	for key := range out {
		for _, state := range states {
			if !state[key] {
				out[key] = false
				break
			}
		}
	}
	return out
}

func objectStateEqual(a, b map[string]bool) bool {
	if len(a) != len(b) {
		return false
	}
	for key, value := range a {
		if b[key] != value {
			return false
		}
	}
	return true
}

func objectMemberReceiver(facts *procedureAnalysisFacts, access procedureir.VariableAccess) bool {
	expression, ok := facts.Expression(access.ExpressionID)
	if !ok {
		return false
	}
	for expression.ParentID != 0 {
		parent, exists := facts.Expression(expression.ParentID)
		if !exists {
			break
		}
		if parent.Kind == procedureir.ExpressionMember {
			return true
		}
		if parent.Kind == procedureir.ExpressionCall {
			prefix := strings.ToLower(strings.TrimSpace(parent.Text))
			name := strings.ToLower(strings.TrimSpace(expression.Text))
			if strings.HasPrefix(prefix, name+"(") || strings.HasPrefix(prefix, name+".") {
				return true
			}
		}
		expression = parent
	}
	statement, ok := facts.Statement(access.StatementID)
	return ok && statement.Kind == procedureir.StatementWith &&
		(statement.TargetID == access.ExpressionID || statement.ValueID == access.ExpressionID)
}
