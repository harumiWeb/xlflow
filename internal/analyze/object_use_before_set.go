package analyze

import (
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
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
	QualifiedName  string
	File           string
	Module         string
	Kind           string
	Line           int
	Params         []objectParameterSummary
	ByRefAssigned  map[int]bool
	ByRefWritten   map[int]bool
	ModuleAssigned map[string]bool
	ModuleWritten  map[string]bool
	ReturnAssigned bool
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
	flowGraph   vbacfg.Graph
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
		procedures := file.procedureProjection()
		moduleDecls := file.moduleDecls()
		for _, proc := range procedures {
			key := objectSummaryKey(file.IR.Path, objectProcedureQualifiedName(proc), string(proc.ProcedureKind), proc.StartLine)
			plan := newObjectProcedurePlan(file, proc, moduleDecls, key)
			analysis.plans[key] = plan
			analysis.order = append(analysis.order, key)
		}
	}
	sort.Strings(analysis.order)
	analysis.initializeObjectSummaries()
	analysis.buildObjectIndexes()
	analysis.buildObjectDependencies()
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
	analysis.moduleProcedureKeys = moduleKeys
	for _, key := range analysis.order {
		plan := analysis.plans[key]
		if plan == nil {
			continue
		}
		plan.receiverSummaryKeys = receiverKeys
		plan.classInitializerKeys = append([]string(nil), initializerKeys[strings.ToLower(cleanIdentifier(plan.proc.Module))]...)
		plan.classIndexBuilt = true
	}
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
	if proc.Graph != nil {
		plan.flowGraph = proc.Graph.WithoutNormalErrRaiseContinuation()
		plan.flowProc.Graph = &plan.flowGraph
		for _, id := range plan.flowGraph.Reachable(vbacfg.EdgeFilter{}) {
			plan.reachable[id] = true
		}
		plan.flowContext = newObjectFlowContext(plan.flowProc)
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
	for _, statement := range proc.Statements {
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
		for _, parameter := range proc.Params {
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
			QualifiedName:  objectProcedureQualifiedName(plan.proc),
			File:           plan.file.IR.Path,
			Module:         plan.proc.Module,
			Kind:           string(plan.proc.ProcedureKind),
			Line:           plan.proc.StartLine,
			ByRefAssigned:  map[int]bool{},
			ByRefWritten:   map[int]bool{},
			ModuleAssigned: map[string]bool{},
			ModuleWritten:  map[string]bool{},
		}
		for index, parameter := range plan.proc.Params {
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
	for _, parameter := range plan.proc.Params {
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
		for _, call := range caller.proc.Calls {
			calleeKey, ok := analysis.objectCalleeKey(call)
			if ok {
				addSummaryDependency(calleeKey, callerKey)
			}
			for _, receiverKey := range analysis.objectReceiverCalleeKeys(caller, call) {
				addSummaryDependency(receiverKey, callerKey)
			}
			if !ok || calleeKey == callerKey {
				// Recursive entry-state calls are deliberately excluded. Summary
				// propagation keeps the self edge so recursive summaries converge.
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
	name := strings.TrimSpace(call.Callee.BaseName)
	if name == "" {
		name = strings.TrimSpace(call.Callee.Member)
	}
	if name == "" || strings.TrimSpace(call.Module) == "" {
		return "", false
	}
	var match string
	for candidateKey, plan := range analysis.plans {
		if !strings.EqualFold(strings.TrimSpace(plan.proc.Module), strings.TrimSpace(call.Module)) || !strings.EqualFold(cleanIdentifier(plan.proc.Name), cleanIdentifier(name)) {
			continue
		}
		if match != "" {
			return "", false
		}
		match = candidateKey
	}
	return match, match != ""
}

func objectIsCurrentModuleReceiver(call procedureir.CallSite) bool {
	return call.Callee.Receiver != nil && strings.EqualFold(cleanIdentifier(strings.TrimSpace(*call.Callee.Receiver)), "me")
}

func (analysis *objectAnalysisContext) buildSummaries() map[string]objectProcedureSummary {
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
		updated.ModuleAssigned = cloneBoolMap(previous.ModuleAssigned)
		updated.ModuleWritten = cloneBoolMap(previous.ModuleWritten)
		for index, parameter := range previous.Params {
			if !parameter.Object || !parameter.ByRef {
				continue
			}
			variable := objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}
			updated.ByRefAssigned[index] = !plan.unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
			updated.ByRefWritten[index] = plan.unknownFlow || objectProcedureWritesParameterIndexed(plan.proc, parameter.Name, analysis.summaries, plan.flowContext.facts)
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
		if !objectSummaryEqual(previous, updated) {
			analysis.summaries[key] = updated
			for _, dependent := range analysis.summaryDependents[key] {
				if !queued[dependent] {
					queue = append(queue, dependent)
					queued[dependent] = true
				}
			}
		}
	}
	return analysis.summaries
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
				for index, parameter := range callee.proc.Params {
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
	if a.ReturnAssigned != b.ReturnAssigned || len(a.ByRefAssigned) != len(b.ByRefAssigned) || len(a.ByRefWritten) != len(b.ByRefWritten) || len(a.ModuleAssigned) != len(b.ModuleAssigned) || len(a.ModuleWritten) != len(b.ModuleWritten) {
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
	for _, parameter := range proc.Params {
		if isObjectType(parameter.Type) {
			return true
		}
	}
	return objectProcedureUsesModuleObject(proc, moduleDecls)
}

func objectProcedureUsesModuleObject(proc sourceProcedure, moduleDecls map[string]sourceDeclaration) bool {
	for _, access := range proc.Accesses {
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
	for _, access := range proc.Accesses {
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
	for _, access := range proc.Accesses {
		if access.Scope != procedureir.ScopeParameter ||
			(access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) ||
			!strings.EqualFold(cleanIdentifier(access.Name), cleanIdentifier(name)) {
			continue
		}
		if !objectMemberReceiver(facts, access) {
			return true
		}
	}
	for _, call := range proc.Calls {
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
	for _, access := range proc.Accesses {
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
	for _, call := range proc.Calls {
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
	for _, statement := range proc.Statements {
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
	graph      *vbacfg.Graph
	dominators map[vbacfg.BlockID][]vbacfg.BlockID
	successors map[vbacfg.BlockID][]vbacfg.BlockID
}

func objectFlowGuardedByOpenFlag(proc sourceProcedure, statementID int, objectName string, cache *objectFlowGuardCache) bool {
	if proc.Graph == nil {
		return false
	}
	if cache.graph == nil {
		flowGraph := proc.Graph.WithoutNormalErrRaiseContinuation()
		cache.graph = &flowGraph
		cache.dominators = flowGraph.Dominators(vbacfg.EdgeFilter{})
	}
	flowProc := proc
	flowProc.Graph = cache.graph
	accessBlock, ok := flowProc.Graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	dominators := cache.dominators
	for _, statement := range proc.Statements {
		flag, negated, ok := objectBooleanGuard(statement)
		if !ok || !objectOpenFlagForObject(flag, objectName) {
			continue
		}
		conditionBlock, ok := flowProc.Graph.BlockForStatement(statement.ID)
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
		for _, edge := range flowProc.Graph.Edges {
			if edge.From != conditionBlock.ID || (edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse) {
				continue
			}
			reachesAccess := objectFlowCanReach(cache.successors, edge.To, accessBlock.ID)
			if edge.Kind == safeBranch {
				safeReachable = safeReachable || reachesAccess
			} else {
				unsafeReachable = unsafeReachable || reachesAccess
			}
		}
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
	cache.successors = make(map[vbacfg.BlockID][]vbacfg.BlockID, len(cache.graph.Blocks))
	for _, edge := range cache.graph.Edges {
		if edge.Class == vbacfg.EdgeExceptional {
			continue
		}
		cache.successors[edge.From] = append(cache.successors[edge.From], edge.To)
	}
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
	for _, flagStatement := range proc.Statements {
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
		for _, objectStatement := range proc.Statements {
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
}

func newObjectFlowContext(proc sourceProcedure) objectFlowContext {
	context := objectFlowContext{
		facts:        proc.analysisFacts(),
		predecessors: make(map[vbacfg.BlockID][]vbacfg.Edge),
		vars:         map[string]objectVariable{},
	}
	if proc.Graph != nil {
		for _, edge := range proc.Graph.Edges {
			context.predecessors[edge.To] = append(context.predecessors[edge.To], edge)
		}
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
	for _, block := range flowProc.Graph.Blocks {
		if !reachable[block.ID] {
			continue
		}
		if block.ID == flowProc.Graph.Entry {
			result.in[block.ID] = cloneObjectState(initial)
		} else {
			result.in[block.ID] = objectStateAllTrue(plan.vars)
		}
		result.out[block.ID] = objectFlowTransfer(plan.file, flowProc, block, result.in[block.ID], plan.vars, plan.declarations, summaries, flowContext)
	}
	changed := true
	for changed {
		changed = false
		for _, block := range flowProc.Graph.Blocks {
			if !reachable[block.ID] || block.ID == flowProc.Graph.Entry {
				continue
			}
			incoming := make([]map[string]bool, 0)
			for _, edge := range flowContext.predecessors[block.ID] {
				if !reachable[edge.From] {
					continue
				}
				state := result.out[edge.From]
				if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain || objectFlowForEachZeroIteration(flowContext, edge) {
					state = result.in[edge.From]
				}
				state = objectFlowApplyGuard(state, flowContext, edge, plan.declarations)
				incoming = append(incoming, state)
			}
			if len(incoming) == 0 {
				continue
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
		}
	}
	if flowProc.Graph != nil {
		result.normalExit = cloneObjectState(result.in[flowProc.Graph.NormalExit])
	}
	return result
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
	text := strings.ToLower(strings.TrimSpace(statement.Condition.Text))
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
	statement := block.Statement
	if statement == nil {
		return state
	}
	callSeen := false
	flowContext.facts.forEachCallForStatement(statement.ID, func(call procedureir.CallSite) {
		callSeen = true
		applyObjectCallEffects(call, state, vars, declarations, flowContext.facts, summaries)
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
		declaration, scope, ok := objectDeclarationBinding(name, declarations)
		if !ok || !declaration.Object {
			if isObjectType(proc.ReturnType) && strings.EqualFold(name, cleanIdentifier(proc.Name)) {
				return state[(objectVariable{Scope: procedureir.ScopeLocal, Name: name}).key()]
			}
			return false
		}
		return state[(objectVariable{Scope: scope, Name: name}).key()]
	case procedureir.ExpressionCall:
		for _, call := range proc.Calls {
			if call.StatementID != statementID || (call.ExpressionID != 0 && call.ExpressionID != expression.ID) {
				continue
			}
			if objectCallReturnsAssigned(proc, statementID, call, state, flowContext, declarations, summaries) {
				return true
			}
		}
		return objectConstructorCallText(lower)
	case procedureir.ExpressionMember:
		if objectExcelMemberExpressionAssigned(expression.Text, proc, declarations) {
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

func objectConstructorCallText(text string) bool {
	text = strings.TrimSpace(text)
	return strings.HasPrefix(text, "createobject(") || strings.HasPrefix(text, "getobject(")
}

func objectCallReturnsAssigned(proc sourceProcedure, statementID int, call procedureir.CallSite, state map[string]bool, flowContext objectFlowContext, declarations declarationScope, summaries map[string]objectProcedureSummary) bool {
	if objectConstructorCallText(strings.ToLower(call.Callee.Text)) || strings.EqualFold(call.Callee.BaseName, "CreateObject") || strings.EqualFold(call.Callee.BaseName, "GetObject") {
		return true
	}
	if call.Callee.Receiver == nil {
		if summary, ok := objectDirectCallSummary(proc, call, summaries); ok {
			return summary.ReturnAssigned
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
			return ok && summary.ReturnAssigned
		}
		var summary objectProcedureSummary
		var ok bool
		if flowContext.receiverSummaryKeys != nil {
			summary, ok = objectReceiverReturnSummaryIndexed(call, declarations, summaries, flowContext.receiverSummaryKeys)
		} else {
			summary, ok = objectReceiverReturnSummary(call, declarations, summaries)
		}
		if ok {
			return summary.ReturnAssigned
		}
		return false
	}
	if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
		return false
	}
	summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries)
	return ok && summary.ReturnAssigned
}

func objectDirectCallSummary(proc sourceProcedure, call procedureir.CallSite, summaries map[string]objectProcedureSummary) (objectProcedureSummary, bool) {
	if call.Callee.Receiver != nil || strings.TrimSpace(call.Callee.BaseName) == "" {
		return objectProcedureSummary{}, false
	}
	matches := objectSummaryCandidatesForDirectCall(proc.Module, call.Callee.BaseName, call.File, summaries)
	if len(matches) == 0 {
		return objectProcedureSummary{}, false
	}
	returnAssigned := true
	for _, summary := range matches {
		returnAssigned = returnAssigned && summary.ReturnAssigned
	}
	return objectProcedureSummary{ReturnAssigned: returnAssigned}, true
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
		for _, irDeclaration := range proc.Declarations {
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
	case "application", "workbook", "worksheet", "range", "chart", "pivot table", "pivottable", "listobject", "window", "shape":
		return true
	default:
		return strings.HasSuffix(typ, ".application") || strings.HasSuffix(typ, ".workbook") || strings.HasSuffix(typ, ".worksheet") || strings.HasSuffix(typ, ".range") || strings.HasSuffix(typ, ".shape")
	}
}

func excelObjectUseMember(member string) bool {
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(strings.SplitN(member, "(", 2)[0])))
	switch member {
	case "application", "workbooks", "worksheets", "sheets", "range", "cells", "rows", "columns", "shapes", "parent", "resize", "offset", "addshape", "selection", "usedrange", "interior", "borders", "font", "textframe", "characters", "fill", "line", "controls", "add":
		return true
	default:
		return false
	}
}

func excelObjectFactoryMember(member string) bool {
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(strings.SplitN(member, "(", 2)[0])))
	switch member {
	case "application", "workbooks", "worksheets", "sheets", "range", "cells", "rows", "columns", "shapes", "usedrange", "interior", "borders", "font", "textframe", "characters", "fill", "line", "controls", "add", "addshape", "parent", "resize", "offset":
		return true
	default:
		return false
	}
}

func applyObjectCallEffects(call procedureir.CallSite, state map[string]bool, vars map[string]objectVariable, declarations declarationScope, facts *procedureAnalysisFacts, summaries map[string]objectProcedureSummary) {
	// Unresolved calls embedded in expressions (for example TypeName(obj) or
	// StrComp(TypeName(obj), ...)) consume values but do not provide a reliable
	// ByRef mutation contract.  Only statement-level unresolved calls can affect
	// a direct object argument; resolved project calls retain full effect flow.
	if call.Resolution.Status != procedureir.ResolutionMatched && call.ExpressionID != 0 {
		return
	}
	actuals := objectCallActuals(call, facts)
	var candidates []objectProcedureSummary
	if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
		if summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries); ok {
			if !strings.EqualFold(strings.TrimSpace(call.Caller.QualifiedName), strings.TrimSpace(summary.QualifiedName)) {
				candidates = append(candidates, summary)
			}
		}
	} else if call.ExpressionID == 0 && (call.Callee.Receiver == nil || objectIsCurrentModuleReceiver(call)) {
		name := call.Callee.BaseName
		if strings.TrimSpace(name) == "" {
			name = call.Callee.Member
		}
		matches := objectSummaryCandidatesForDirectCall(call.Module, name, call.File, summaries)
		if len(matches) != 1 {
			matches = nil
		}
		for _, summary := range matches {
			if !strings.EqualFold(strings.TrimSpace(call.Caller.QualifiedName), strings.TrimSpace(summary.QualifiedName)) {
				candidates = append(candidates, summary)
			}
		}
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
			if !summary.Params[formalIndex].ByRef {
				// ByVal receives a copy and cannot alter the caller's object
				// reference; preserve its existing state.
				assigned = state[variable.key()]
				continue
			}
			if actual.parenthesized {
				// Parenthesized actuals are passed through a temporary ByVal
				// value even when the formal parameter is ByRef.
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
