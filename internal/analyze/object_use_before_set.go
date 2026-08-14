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

// buildObjectProcedureSummaries computes a small interprocedural contract for
// object use-before-set.  The fixed point starts with no callee guarantees,
// so recursive or unresolved calls never manufacture an initialization fact.
func buildObjectProcedureSummaries(files []parsedFile) map[string]objectProcedureSummary {
	type procedureInfo struct {
		file         parsedFile
		proc         sourceProcedure
		moduleDecls  map[string]sourceDeclaration
		declarations map[string]sourceDeclaration
	}
	infos := make([]procedureInfo, 0)
	summaries := map[string]objectProcedureSummary{}
	for _, file := range files {
		procedures := sourceProceduresFromIR(file.IR, file.CFG)
		moduleDecls := moduleDeclarations(file.Lines, procedures)
		for _, proc := range procedures {
			info := procedureInfo{
				file: file, proc: proc, moduleDecls: moduleDecls,
				declarations: objectFlowDeclarations(file.Lines, proc, moduleDecls),
			}
			infos = append(infos, info)
			qualified := objectProcedureQualifiedName(proc)
			summary := objectProcedureSummary{
				QualifiedName:  qualified,
				File:           file.IR.Path,
				Module:         proc.Module,
				Kind:           string(proc.ProcedureKind),
				Line:           proc.StartLine,
				ByRefAssigned:  map[int]bool{},
				ByRefWritten:   map[int]bool{},
				ModuleAssigned: map[string]bool{},
				ModuleWritten:  map[string]bool{},
			}
			for index, parameter := range proc.Params {
				summary.Params = append(summary.Params, objectParameterSummary{
					Name: parameter.Name, Object: isObjectType(parameter.Type),
					ByRef: !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal"),
				})
				if summary.Params[index].Object && summary.Params[index].ByRef {
					summary.ByRefAssigned[index] = false
					summary.ByRefWritten[index] = false
				}
			}
			summaries[objectSummaryKey(summary.File, summary.QualifiedName, summary.Kind, summary.Line)] = summary
		}
	}
	for iteration := 0; iteration < len(infos)+1; iteration++ {
		changed := false
		for _, info := range infos {
			if !objectProcedureRelevant(info.proc, info.moduleDecls) {
				continue
			}
			qualified := objectProcedureQualifiedName(info.proc)
			key := objectSummaryKey(info.file.IR.Path, qualified, string(info.proc.ProcedureKind), info.proc.StartLine)
			previous := summaries[key]
			flow := objectStateFlow(info.file, info.proc, info.declarations, summaries, objectFlowOptions{Summary: true})
			updated := previous
			updated.ByRefAssigned = cloneIntBoolMap(previous.ByRefAssigned)
			updated.ByRefWritten = cloneIntBoolMap(previous.ByRefWritten)
			updated.ModuleAssigned = cloneBoolMap(previous.ModuleAssigned)
			updated.ModuleWritten = cloneBoolMap(previous.ModuleWritten)
			unknownFlow := info.proc.Graph == nil || len(info.proc.Graph.UnknownFlowSources) > 0
			for _, statement := range info.proc.Statements {
				unknownFlow = unknownFlow || statement.Recovered
			}
			for index, parameter := range previous.Params {
				if !parameter.Object || !parameter.ByRef {
					continue
				}
				variable := objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}
				updated.ByRefAssigned[index] = !unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
				updated.ByRefWritten[index] = unknownFlow || objectProcedureWritesParameter(info.proc, parameter.Name, summaries)
			}
			for name, declaration := range info.moduleDecls {
				if !declaration.Object {
					continue
				}
				variable := objectVariable{Scope: procedureir.ScopeModule, Name: declaration.Name}
				updated.ModuleAssigned[strings.ToLower(name)] = !unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
				updated.ModuleWritten[strings.ToLower(name)] = objectProcedureWritesModuleField(info.proc, declaration.Name)
			}
			if isObjectType(info.proc.ReturnType) {
				variable := objectVariable{Scope: procedureir.ScopeLocal, Name: info.proc.Name}
				updated.ReturnAssigned = !unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
			}
			if !objectSummaryEqual(previous, updated) {
				summaries[key] = updated
				changed = true
			}
		}
		if !changed {
			break
		}
	}
	return summaries
}

// buildObjectProcedureEntryStates computes module-field and private-procedure
// parameter state at procedure entries from resolved call sites. A value is
// admitted only when every reachable call site into that procedure proves a
// definitely non-Nothing value immediately before the call. Uncalled,
// ambiguous, external, and recursive procedures therefore retain the
// conservative MaybeNothing entry state. Public procedures are excluded from
// parameter-entry inference because callers outside the analyzed project can
// still pass Nothing.
func buildObjectProcedureEntryStates(files []parsedFile, summaries map[string]objectProcedureSummary) map[string]map[string]bool {
	type procedureInfo struct {
		file         parsedFile
		proc         sourceProcedure
		moduleDecls  map[string]sourceDeclaration
		declarations map[string]sourceDeclaration
		key          string
	}
	infos := make([]procedureInfo, 0)
	byKey := map[string]procedureInfo{}
	entries := map[string]map[string]bool{}
	for _, file := range files {
		procedures := sourceProceduresFromIR(file.IR, file.CFG)
		moduleDecls := moduleDeclarations(file.Lines, procedures)
		for _, proc := range procedures {
			key := objectSummaryKey(file.IR.Path, objectProcedureQualifiedName(proc), string(proc.ProcedureKind), proc.StartLine)
			info := procedureInfo{file: file, proc: proc, moduleDecls: moduleDecls, declarations: objectFlowDeclarations(file.Lines, proc, moduleDecls), key: key}
			infos = append(infos, info)
			byKey[key] = info
			state := map[string]bool{}
			for name, declaration := range moduleDecls {
				if declaration.Object {
					state[(objectVariable{Scope: procedureir.ScopeModule, Name: name}).key()] = false
				}
			}
			for _, parameter := range proc.Params {
				if isObjectType(parameter.Type) {
					state[(objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}).key()] = false
				}
			}
			entries[key] = state
		}
	}
	for iteration := 0; iteration < len(infos)+1; iteration++ {
		next := map[string]map[string]bool{}
		seen := map[string]map[string]bool{}
		for key, state := range entries {
			next[key] = cloneBoolMap(state)
			seen[key] = map[string]bool{}
		}
		for _, caller := range infos {
			expressions := make(map[int]procedureir.Expression, len(caller.proc.Expressions))
			for _, expression := range caller.proc.Expressions {
				expressions[expression.ID] = expression
			}
			flow := objectStateFlow(caller.file, caller.proc, caller.declarations, summaries, objectFlowOptions{Entry: entries[caller.key]})
			for _, call := range caller.proc.Calls {
				if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
					continue
				}
				candidate := call.Resolution.Candidates[0]
				if strings.EqualFold(strings.TrimSpace(call.Caller.QualifiedName), strings.TrimSpace(candidate.QualifiedName)) {
					// A recursive call cannot establish a new entry guarantee for
					// itself; admitting its current module state would be circular.
					continue
				}
				calleeKey := objectSummaryKey(candidate.File, candidate.QualifiedName, candidate.Kind, candidate.Line)
				callee, ok := byKey[calleeKey]
				if !ok {
					continue
				}
				calleeSummary, ok := summaries[callee.key]
				if !ok {
					continue
				}
				if caller.proc.Graph == nil {
					continue
				}
				block, ok := caller.proc.Graph.BlockForStatement(call.StatementID)
				if !ok || !caller.proc.Graph.IsReachable(block.ID, vbacfg.EdgeFilter{}) {
					continue
				}
				state := flow.in[block.ID]
				if objectProcedureAllowsParameterEntry(callee.proc) {
					actuals := objectCallActuals(call, expressions)
					for index, parameter := range callee.proc.Params {
						if !isObjectType(parameter.Type) {
							continue
						}
						parameterKey := (objectVariable{Scope: procedureir.ScopeParameter, Name: parameter.Name}).key()
						assigned, present := objectCallParameterAssigned(caller.proc, caller.declarations, call, calleeSummary, index, actuals, state, flow.vars, expressions, summaries)
						if !seen[callee.key][parameterKey] {
							next[callee.key][parameterKey] = assigned && present
						}
						seen[callee.key][parameterKey] = true
						if !assigned || !present {
							next[callee.key][parameterKey] = false
						}
					}
				}
				if !strings.EqualFold(callee.proc.Module, caller.proc.Module) {
					continue
				}
				for name, declaration := range callee.moduleDecls {
					if !declaration.Object {
						continue
					}
					field := objectVariable{Scope: procedureir.ScopeModule, Name: name}
					fieldKey := field.key()
					if !seen[callee.key][fieldKey] {
						next[callee.key][fieldKey] = true
					}
					seen[callee.key][fieldKey] = true
					if !state[fieldKey] {
						next[callee.key][fieldKey] = false
					}
				}
			}
		}
		changed := false
		for key, state := range next {
			for fieldKey := range state {
				if !seen[key][fieldKey] {
					state[fieldKey] = false
				}
			}
			if !objectBoolMapEqual(entries[key], state) {
				changed = true
			}
		}
		entries = next
		if !changed {
			break
		}
	}
	return entries
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

func objectProcedureWritesModuleField(proc sourceProcedure, name string) bool {
	expressions := make(map[int]procedureir.Expression, len(proc.Expressions))
	for _, expression := range proc.Expressions {
		expressions[expression.ID] = expression
	}
	statements := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statements[statement.ID] = statement
	}
	for _, access := range proc.Accesses {
		if access.Scope != procedureir.ScopeModule || (access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) || !strings.EqualFold(cleanIdentifier(access.Name), cleanIdentifier(name)) {
			continue
		}
		if objectMemberReceiver(expressions, statements, access) {
			continue
		}
		return true
	}
	return false
}

func objectProcedureWritesParameter(proc sourceProcedure, name string, summaries map[string]objectProcedureSummary) bool {
	expressions := make(map[int]procedureir.Expression, len(proc.Expressions))
	for _, expression := range proc.Expressions {
		expressions[expression.ID] = expression
	}
	statements := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statements[statement.ID] = statement
	}
	for _, access := range proc.Accesses {
		if access.Scope != procedureir.ScopeParameter ||
			(access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) ||
			!strings.EqualFold(cleanIdentifier(access.Name), cleanIdentifier(name)) {
			continue
		}
		if !objectMemberReceiver(expressions, statements, access) {
			return true
		}
	}
	for _, call := range proc.Calls {
		actuals := objectCallActuals(call, expressions)
		for actualIndex, actual := range actuals {
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
	Summary bool
	Entry   map[string]bool
}

// objectUseBeforeSetIRFindings reports the first unsafe member/collection use
// of each object variable.  The state at a use comes from the CFG entry fact,
// not from source-line order, so branches, early exits, loops and error edges
// are all represented by the same must-analysis.
func (a Analyzer) objectUseBeforeSetIRFindings(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, summaries map[string]objectProcedureSummary, entry map[string]bool) []Finding {
	if proc.Graph == nil {
		return nil
	}
	declarations := objectFlowDeclarations(file.Lines, proc, moduleDecls)
	procedureDecls := procedureDeclarations(file.Lines, proc)
	relevant := isObjectType(proc.ReturnType)
	if !relevant {
		for _, declaration := range procedureDecls {
			if declaration.Object {
				relevant = true
				break
			}
		}
	}
	if !relevant {
		for _, parameter := range proc.Params {
			if isObjectType(parameter.Type) {
				relevant = true
				break
			}
		}
	}
	if !relevant {
		relevant = objectProcedureUsesModuleObject(proc, moduleDecls)
	}
	if !relevant {
		return nil
	}
	flow := objectStateFlow(file, proc, declarations, summaries, objectFlowOptions{Entry: entry})
	expressions := make(map[int]procedureir.Expression, len(proc.Expressions))
	for _, expression := range proc.Expressions {
		expressions[expression.ID] = expression
	}
	statements := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statements[statement.ID] = statement
	}

	accesses := append([]procedureir.VariableAccess(nil), proc.Accesses...)
	sort.SliceStable(accesses, func(i, j int) bool {
		if accesses[i].Range.StartByte != accesses[j].Range.StartByte {
			return accesses[i].Range.StartByte < accesses[j].Range.StartByte
		}
		return accesses[i].ExpressionID < accesses[j].ExpressionID
	})
	reported := map[string]bool{}
	guardCache := &objectFlowGuardCache{}
	var findings []Finding
	for _, access := range accesses {
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
		if !objectMemberReceiver(expressions, statements, access) {
			continue
		}
		declaration, ok := objectDeclarationFor(access.Name, access.Scope, declarations)
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
		variable := objectVariable{Scope: access.Scope, Name: access.Name}
		if access.Scope == procedureir.ScopeUnresolved {
			variable.Scope = procedureir.ScopeLocal
		}
		key := variable.key()
		block, ok := proc.Graph.BlockForStatement(access.StatementID)
		if !ok || !proc.Graph.IsReachable(block.ID, vbacfg.EdgeFilter{}) {
			continue
		}
		if _, ok := flow.in[block.ID]; !ok {
			// The error-raising graph view may prove that a source-recovered
			// access has no normal path. Do not index a missing state as
			// MaybeNothing and manufacture a finding for that unreachable block.
			continue
		}
		if objectErrorResumeNextAt(proc, access.StatementID) {
			continue
		}
		if objectFlowGuardedByOpenFlag(proc, access.StatementID, access.Name, guardCache) {
			continue
		}
		if objectFlowAssigned(flow.in[block.ID], variable) || reported[key] {
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
	calls := append([]procedureir.CallSite(nil), proc.Calls...)
	sort.SliceStable(calls, func(i, j int) bool { return calls[i].Range.StartByte < calls[j].Range.StartByte })
	for _, call := range calls {
		name := objectCallReceiverName(call)
		if name == "" {
			continue
		}
		variable := objectFlowVariableForName(name, flow.vars)
		key := variable.key()
		declaration, ok := objectDeclarationFor(name, variable.Scope, declarations)
		if !ok || !declaration.Object || declaration.NewExpression || declaration.Array || reported[key] {
			continue
		}
		block, ok := proc.Graph.BlockForStatement(call.StatementID)
		if !ok || !proc.Graph.IsReachable(block.ID, vbacfg.EdgeFilter{}) {
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
		for _, edge := range flowProc.Graph.Edges {
			if edge.From != conditionBlock.ID || (edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse) {
				continue
			}
			reachesAccess := objectFlowCanReach(flowProc.Graph, edge.To, accessBlock.ID)
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

func objectFlowCanReach(graph *vbacfg.Graph, start, target vbacfg.BlockID) bool {
	if start == target {
		return true
	}
	seen := map[vbacfg.BlockID]bool{start: true}
	queue := []vbacfg.BlockID{start}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		for _, edge := range graph.Edges {
			if edge.From != current || edge.Class == vbacfg.EdgeExceptional || seen[edge.To] {
				continue
			}
			if edge.To == target {
				return true
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
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
	expressions  map[int]procedureir.Expression
	statements   map[int]procedureir.Statement
	calls        map[int][]procedureir.CallSite
	accesses     map[int][]procedureir.VariableAccess
	predecessors map[vbacfg.BlockID][]vbacfg.Edge
	vars         map[string]objectVariable
}

func newObjectFlowContext(proc sourceProcedure) objectFlowContext {
	context := objectFlowContext{
		expressions:  make(map[int]procedureir.Expression, len(proc.Expressions)),
		statements:   make(map[int]procedureir.Statement, len(proc.Statements)),
		calls:        make(map[int][]procedureir.CallSite),
		accesses:     make(map[int][]procedureir.VariableAccess),
		predecessors: make(map[vbacfg.BlockID][]vbacfg.Edge),
		vars:         map[string]objectVariable{},
	}
	for _, expression := range proc.Expressions {
		context.expressions[expression.ID] = expression
	}
	for _, statement := range proc.Statements {
		context.statements[statement.ID] = statement
	}
	for _, call := range proc.Calls {
		context.calls[call.StatementID] = append(context.calls[call.StatementID], call)
	}
	for _, access := range proc.Accesses {
		context.accesses[access.StatementID] = append(context.accesses[access.StatementID], access)
	}
	if proc.Graph != nil {
		for _, edge := range proc.Graph.Edges {
			context.predecessors[edge.To] = append(context.predecessors[edge.To], edge)
		}
	}
	return context
}

func objectFlowDeclarations(lines []string, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]sourceDeclaration {
	declarations := map[string]sourceDeclaration{}
	localDecls := procedureDeclarations(lines, proc)
	for key, declaration := range moduleDecls {
		key = strings.ToLower(key)
		if _, shadowed := localDecls[key]; shadowed {
			declarations["@module:"+key] = declaration
			continue
		}
		declarations[key] = declaration
	}
	for key, declaration := range localDecls {
		declarations[strings.ToLower(key)] = declaration
	}
	for _, parameter := range proc.Params {
		declaration := sourceDeclaration{
			Name: parameter.Name, Type: parameter.Type, Object: isObjectType(parameter.Type), Parameter: true,
		}
		key := strings.ToLower(parameter.Name)
		if _, shadowed := declarations[key]; shadowed {
			declarations["@parameter:"+key] = declaration
		} else {
			declarations[key] = declaration
		}
	}
	return declarations
}

func objectDeclarationFor(name string, scope procedureir.SymbolScope, declarations map[string]sourceDeclaration) (sourceDeclaration, bool) {
	key := strings.ToLower(cleanIdentifier(name))
	if scope == procedureir.ScopeModule {
		if declaration, ok := declarations["@module:"+key]; ok {
			return declaration, true
		}
	}
	if scope == procedureir.ScopeParameter {
		if declaration, ok := declarations["@parameter:"+key]; ok {
			return declaration, true
		}
	}
	declaration, ok := declarations[key]
	return declaration, ok
}

func objectStateFlow(file parsedFile, proc sourceProcedure, declarations map[string]sourceDeclaration, summaries map[string]objectProcedureSummary, options objectFlowOptions) objectFlowResult {
	result := objectFlowResult{in: map[vbacfg.BlockID]map[string]bool{}, out: map[vbacfg.BlockID]map[string]bool{}, vars: map[string]objectVariable{}}
	if proc.Graph == nil {
		return result
	}
	flowGraph := proc.Graph.WithoutNormalErrRaiseContinuation()
	flowProc := proc
	flowProc.Graph = &flowGraph
	flowContext := newObjectFlowContext(flowProc)
	procedureDecls := procedureDeclarations(file.Lines, proc)
	for key, declaration := range declarations {
		if !declaration.Object {
			continue
		}
		scope := procedureir.ScopeModule
		if declaration.Parameter {
			scope = procedureir.ScopeParameter
		} else if _, local := procedureDecls[key]; local {
			scope = procedureir.ScopeLocal
		}
		variable := objectVariable{Scope: scope, Name: declaration.Name}
		result.vars[variable.key()] = variable
	}
	if isObjectType(proc.ReturnType) {
		variable := objectVariable{Scope: procedureir.ScopeLocal, Name: proc.Name}
		result.vars[variable.key()] = variable
	}
	if len(result.vars) == 0 {
		return result
	}
	flowContext.vars = result.vars
	initial := map[string]bool{}
	for key, variable := range result.vars {
		initial[key] = false
		declaration, _ := objectDeclarationFor(variable.Name, variable.Scope, declarations)
		if declaration.NewExpression || objectClassLifecycleAssigned(proc, variable, summaries) || (options.Entry != nil && options.Entry[variable.key()]) {
			initial[key] = true
		}
		// Parameters are intentionally not initialized at procedure entry.  A
		// caller can pass Nothing, and only a dominating Set or a proven ByRef
		// initializer establishes a safe value.
	}
	reachable := map[vbacfg.BlockID]bool{}
	for _, id := range flowProc.Graph.Reachable(vbacfg.EdgeFilter{}) {
		reachable[id] = true
	}
	for _, block := range flowProc.Graph.Blocks {
		if !reachable[block.ID] {
			continue
		}
		if block.ID == flowProc.Graph.Entry {
			result.in[block.ID] = cloneObjectState(initial)
		} else {
			result.in[block.ID] = objectStateAllTrue(result.vars)
		}
		result.out[block.ID] = objectFlowTransfer(file, flowProc, block, result.in[block.ID], result.vars, declarations, summaries, flowContext)
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
				state = objectFlowApplyGuard(state, flowContext, edge)
				incoming = append(incoming, state)
			}
			if len(incoming) == 0 {
				continue
			}
			next := objectFlowIntersection(incoming, result.vars)
			if !objectStateEqual(result.in[block.ID], next) {
				result.in[block.ID] = next
				changed = true
			}
			updated := objectFlowTransfer(file, flowProc, block, next, result.vars, declarations, summaries, flowContext)
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

// objectFlowApplyGuard refines the state on the edge selected by a direct
// `obj Is Nothing`/`Not obj Is Nothing` condition.  Compound boolean
// expressions are deliberately ignored; VBA212 remains responsible for
// short-circuit/eager Boolean diagnostics and this analysis stays conservative.
func objectFlowApplyGuard(state map[string]bool, flowContext objectFlowContext, edge vbacfg.Edge) map[string]bool {
	if edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	statement, ok := flowContext.statements[edge.StatementID]
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
		key := objectGuardVariableKey(name, state)
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
	key := objectGuardVariableKey(name, state)
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

func objectGuardVariableKey(name string, state map[string]bool) string {
	name = cleanIdentifier(strings.TrimSpace(name))
	if name == "" {
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

func objectFlowTransfer(file parsedFile, proc sourceProcedure, block vbacfg.Block, input map[string]bool, vars map[string]objectVariable, declarations map[string]sourceDeclaration, summaries map[string]objectProcedureSummary, flowContext objectFlowContext) map[string]bool {
	state := cloneObjectState(input)
	statement := block.Statement
	if statement == nil {
		return state
	}
	for _, call := range flowContext.calls[statement.ID] {
		applyObjectCallEffects(call, state, vars, declarations, flowContext.expressions, summaries)
	}
	// The IR represents implicit call statements such as `Helper obj` as an
	// assignment-shaped statement. Their argument effects were applied above;
	// do not then treat the first argument as an object assignment target.
	if statement.Kind == procedureir.StatementAssignment && len(flowContext.calls[statement.ID]) > 0 {
		return state
	}
	target, ok := objectFlowTarget(proc, *statement, declarations, flowContext)
	if !ok {
		return state
	}
	switch statement.Kind {
	case procedureir.StatementSet:
		state[target.key()] = objectFlowValueAssigned(proc, *statement, state, flowContext.expressions, declarations, summaries)
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

func objectFlowTarget(proc sourceProcedure, statement procedureir.Statement, declarations map[string]sourceDeclaration, flowContext objectFlowContext) (objectVariable, bool) {
	for _, access := range flowContext.accesses[statement.ID] {
		if access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite {
			continue
		}
		// A receiver in `obj.Member = value` or `dict(key) = value` is
		// dereferenced, but it is not the object reference being assigned.  Do
		// not reset its state merely because the member/index write is modeled
		// as AccessWrite by the IR.
		if objectMemberReceiver(flowContext.expressions, flowContext.statements, access) {
			continue
		}
		if objectIndexedReceiverWrite(flowContext.calls[statement.ID], access) {
			continue
		}
		declaration, ok := objectDeclarationFor(access.Name, access.Scope, declarations)
		if !ok || !declaration.Object {
			continue
		}
		scope := access.Scope
		if scope == procedureir.ScopeUnresolved {
			scope = procedureir.ScopeLocal
		}
		return objectVariable{Scope: scope, Name: access.Name}, true
	}
	if statement.Target != nil && statement.Target.Kind != procedureir.ExpressionCall && statement.Target.Kind != procedureir.ExpressionMember && !objectIndexedTargetCall(flowContext.calls[statement.ID], statement.Target.Text) {
		name := cleanIdentifier(strings.TrimSpace(statement.Target.Text))
		if declaration, ok := objectDeclarationFor(name, procedureir.ScopeLocal, declarations); ok && declaration.Object {
			scope := procedureir.ScopeLocal
			if declaration.Parameter {
				scope = procedureir.ScopeParameter
			}
			return objectVariable{Scope: scope, Name: name}, true
		}
	}
	if proc.Name != "" && statement.Target != nil && strings.EqualFold(cleanIdentifier(statement.Target.Text), proc.Name) && isObjectType(proc.ReturnType) {
		return objectVariable{Scope: procedureir.ScopeLocal, Name: proc.Name}, true
	}
	return objectVariable{}, false
}

func objectIndexedReceiverWrite(calls []procedureir.CallSite, access procedureir.VariableAccess) bool {
	for _, call := range calls {
		if call.Callee.Receiver == nil && call.Arguments.Count > 0 && strings.EqualFold(cleanIdentifier(call.Callee.BaseName), cleanIdentifier(access.Name)) {
			return true
		}
	}
	return false
}

func objectIndexedTargetCall(calls []procedureir.CallSite, targetText string) bool {
	name := cleanIdentifier(strings.TrimSpace(targetText))
	for _, call := range calls {
		if call.Callee.Receiver == nil && call.Arguments.Count > 0 && strings.EqualFold(cleanIdentifier(call.Callee.BaseName), name) {
			return true
		}
	}
	return false
}

func objectFlowValueAssigned(proc sourceProcedure, statement procedureir.Statement, state map[string]bool, expressions map[int]procedureir.Expression, declarations map[string]sourceDeclaration, summaries map[string]objectProcedureSummary) bool {
	value := statement.Value
	if value == nil {
		return false
	}
	return objectExpressionAssigned(proc, *value, state, expressions, declarations, summaries, statement.ID)
}

func objectExpressionAssigned(proc sourceProcedure, expression procedureir.Expression, state map[string]bool, expressions map[int]procedureir.Expression, declarations map[string]sourceDeclaration, summaries map[string]objectProcedureSummary, statementID int) bool {
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
			if nested, ok := expressions[child]; ok {
				return objectExpressionAssigned(proc, nested, state, expressions, declarations, summaries, statementID)
			}
		}
	case procedureir.ExpressionIdentifier:
		name := cleanIdentifier(text)
		for _, scope := range []procedureir.SymbolScope{procedureir.ScopeLocal, procedureir.ScopeParameter, procedureir.ScopeModule} {
			if value, exists := state[objectVariable{Scope: scope, Name: name}.key()]; exists {
				// Stop at the first lexical binding even when it is currently
				// Nothing; a same-named module field must not initialize a shadowing
				// local object.
				return value
			}
		}
	case procedureir.ExpressionCall:
		for _, call := range proc.Calls {
			if call.StatementID != statementID || (call.ExpressionID != 0 && call.ExpressionID != expression.ID) {
				continue
			}
			if objectCallReturnsAssigned(proc, statementID, call, state, declarations, summaries) {
				return true
			}
		}
		return objectConstructorCallText(lower)
	case procedureir.ExpressionMember:
		// A member rooted at an intrinsic workbook/application object is a
		// non-Nothing factory value.  For a user object, the member may itself be
		// Nothing; keep the assignment nullable and report a later dereference.
		for _, childID := range expression.Children {
			if child, ok := expressions[childID]; ok && child.Kind == procedureir.ExpressionIdentifier {
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

func objectCallReturnsAssigned(proc sourceProcedure, statementID int, call procedureir.CallSite, state map[string]bool, declarations map[string]sourceDeclaration, summaries map[string]objectProcedureSummary) bool {
	if objectConstructorCallText(strings.ToLower(call.Callee.Text)) || strings.EqualFold(call.Callee.BaseName, "CreateObject") || strings.EqualFold(call.Callee.BaseName, "GetObject") {
		return true
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
		if summary, ok := objectReceiverReturnSummary(call, declarations, summaries); ok {
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

func objectReceiverReturnSummary(call procedureir.CallSite, declarations map[string]sourceDeclaration, summaries map[string]objectProcedureSummary) (objectProcedureSummary, bool) {
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

func objectExcelMemberChainAssigned(call procedureir.CallSite, state map[string]bool, declarations map[string]sourceDeclaration) bool {
	if call.Callee.Receiver == nil {
		return false
	}
	receiver := strings.TrimSpace(*call.Callee.Receiver)
	parts := strings.Split(receiver, ".")
	if len(parts) == 0 {
		return false
	}
	root := cleanIdentifier(strings.TrimSpace(strings.SplitN(parts[0], "(", 2)[0]))
	if root == "" || !objectNameDefinitelyAssigned(root, state) {
		return false
	}
	declaration, ok := objectDeclarationByName(root, declarations)
	if !ok || !excelObjectUseType(declaration.Type) {
		return false
	}
	for _, part := range parts[1:] {
		if !excelObjectUseMember(part) {
			return false
		}
	}
	return excelObjectUseMember(call.Callee.Member)
}

func objectExcelMemberFactoryAssigned(call procedureir.CallSite, declarations map[string]sourceDeclaration) bool {
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

func objectNameDefinitelyAssigned(name string, state map[string]bool) bool {
	name = cleanIdentifier(strings.TrimSpace(name))
	for _, scope := range []procedureir.SymbolScope{procedureir.ScopeLocal, procedureir.ScopeParameter, procedureir.ScopeModule} {
		key := (objectVariable{Scope: scope, Name: name}).key()
		if value, exists := state[key]; exists {
			return value
		}
	}
	return false
}

func objectDeclarationByName(name string, declarations map[string]sourceDeclaration) (sourceDeclaration, bool) {
	for _, scope := range []procedureir.SymbolScope{procedureir.ScopeLocal, procedureir.ScopeParameter, procedureir.ScopeModule} {
		if declaration, ok := objectDeclarationFor(name, scope, declarations); ok {
			return declaration, true
		}
	}
	return sourceDeclaration{}, false
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
	case "application", "worksheets", "sheets", "range", "cells", "rows", "columns", "shapes", "parent", "resize", "offset", "addshape", "selection", "usedrange", "interior", "borders", "font", "textframe", "characters", "fill", "line", "controls", "add":
		return true
	default:
		return false
	}
}

func excelObjectFactoryMember(member string) bool {
	member = strings.ToLower(cleanIdentifier(strings.TrimSpace(strings.SplitN(member, "(", 2)[0])))
	switch member {
	case "application", "worksheets", "sheets", "range", "cells", "rows", "columns", "shapes", "usedrange", "interior", "borders", "font", "textframe", "characters", "fill", "line", "controls", "add", "addshape", "parent", "resize", "offset":
		return true
	default:
		return false
	}
}

func applyObjectCallEffects(call procedureir.CallSite, state map[string]bool, vars map[string]objectVariable, declarations map[string]sourceDeclaration, expressions map[int]procedureir.Expression, summaries map[string]objectProcedureSummary) {
	// Unresolved calls embedded in expressions (for example TypeName(obj) or
	// StrComp(TypeName(obj), ...)) consume values but do not provide a reliable
	// ByRef mutation contract.  Only statement-level unresolved calls can affect
	// a direct object argument; resolved project calls retain full effect flow.
	if call.Resolution.Status != procedureir.ResolutionMatched && call.ExpressionID != 0 {
		return
	}
	actuals := objectCallActuals(call, expressions)
	var candidates []objectProcedureSummary
	if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
		if summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries); ok {
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
		variable := objectVariable{Scope: actual.scope, Name: name}
		if variable.Scope == procedureir.ScopeUnresolved {
			variable = objectFlowVariableForName(name, vars)
		}
		declaration, ok := objectDeclarationFor(name, variable.Scope, declarations)
		if !ok || !declaration.Object {
			continue
		}
		if _, exists := vars[variable.key()]; !exists {
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
	expressionID int
	text         string
	scope        procedureir.SymbolScope
}

func objectProcedureAllowsParameterEntry(proc sourceProcedure) bool {
	return strings.EqualFold(strings.TrimSpace(proc.Visibility), "private")
}

func objectCallParameterAssigned(proc sourceProcedure, declarations map[string]sourceDeclaration, call procedureir.CallSite, summary objectProcedureSummary, formalIndex int, actuals []objectCallActual, state map[string]bool, vars map[string]objectVariable, expressions map[int]procedureir.Expression, summaries map[string]objectProcedureSummary) (bool, bool) {
	for actualIndex, actual := range actuals {
		if objectFormalIndex(call, summary, actualIndex) != formalIndex {
			continue
		}
		if actual.expressionID != 0 {
			if expression, ok := expressions[actual.expressionID]; ok && expression.Kind != procedureir.ExpressionIdentifier {
				return objectExpressionAssigned(proc, expression, state, expressions, declarations, summaries, call.StatementID), true
			}
		}
		name := cleanIdentifier(actual.text)
		if name == "" {
			return false, true
		}
		variable := objectFlowVariableForName(name, vars)
		if _, exists := vars[variable.key()]; !exists {
			return false, false
		}
		return state[variable.key()], true
	}
	return false, false
}

func objectCallActuals(call procedureir.CallSite, expressions map[int]procedureir.Expression) []objectCallActual {
	actuals := make([]objectCallActual, 0, len(call.Arguments.ExpressionIDs))
	for _, id := range call.Arguments.ExpressionIDs {
		expression, ok := expressions[id]
		if !ok {
			// Preserve the positional slot.  Dropping a missing expression ID
			// would shift every following actual onto the wrong formal parameter.
			actuals = append(actuals, objectCallActual{})
			continue
		}
		for expression.Kind == procedureir.ExpressionParentheses && len(expression.Children) == 1 {
			nested, nestedOK := expressions[expression.Children[0]]
			if !nestedOK {
				break
			}
			expression = nested
		}
		actuals = append(actuals, objectCallActual{expressionID: expression.ID, text: expression.Text, scope: procedureir.ScopeUnresolved})
	}
	return actuals
}

func objectFlowVariableForName(name string, vars map[string]objectVariable) objectVariable {
	for _, scope := range []procedureir.SymbolScope{procedureir.ScopeLocal, procedureir.ScopeParameter, procedureir.ScopeModule} {
		variable := objectVariable{Scope: scope, Name: name}
		if _, ok := vars[variable.key()]; ok {
			return variable
		}
	}
	return objectVariable{Scope: procedureir.ScopeLocal, Name: name}
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
	statement, ok := flowContext.statements[edge.StatementID]
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

func objectMemberReceiver(expressions map[int]procedureir.Expression, statements map[int]procedureir.Statement, access procedureir.VariableAccess) bool {
	expression, ok := expressions[access.ExpressionID]
	if !ok {
		return false
	}
	for expression.ParentID != 0 {
		parent, exists := expressions[expression.ParentID]
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
	statement, ok := statements[access.StatementID]
	return ok && statement.Kind == procedureir.StatementWith &&
		(statement.TargetID == access.ExpressionID || statement.ValueID == access.ExpressionID)
}
