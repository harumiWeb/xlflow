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
// guarantees when it returns normally.  A ByRef object parameter is useful to
// the caller only when every normal path assigns it a non-Nothing value.  The
// same rule is used for an object function result; an object-returning
// function with an omitted or Nothing result remains nullable.
type objectProcedureSummary struct {
	QualifiedName  string
	File           string
	Module         string
	Kind           string
	Line           int
	Params         []objectParameterSummary
	ByRefAssigned  map[int]bool
	ModuleAssigned map[string]bool
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
		file parsedFile
		proc sourceProcedure
	}
	infos := make([]procedureInfo, 0)
	summaries := map[string]objectProcedureSummary{}
	for _, file := range files {
		for _, proc := range sourceProceduresFromIR(file.IR, file.CFG) {
			info := procedureInfo{file: file, proc: proc}
			infos = append(infos, info)
			qualified := objectProcedureQualifiedName(proc)
			summary := objectProcedureSummary{
				QualifiedName:  qualified,
				File:           file.IR.Path,
				Module:         proc.Module,
				Kind:           string(proc.ProcedureKind),
				Line:           proc.StartLine,
				ByRefAssigned:  map[int]bool{},
				ModuleAssigned: map[string]bool{},
			}
			for index, parameter := range proc.Params {
				summary.Params = append(summary.Params, objectParameterSummary{
					Name: parameter.Name, Object: isObjectType(parameter.Type),
					ByRef: !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal"),
				})
				if summary.Params[index].Object && summary.Params[index].ByRef {
					summary.ByRefAssigned[index] = false
				}
			}
			summaries[objectSummaryKey(summary.File, summary.QualifiedName, summary.Kind, summary.Line)] = summary
		}
	}
	for iteration := 0; iteration < len(infos)+1; iteration++ {
		changed := false
		for _, info := range infos {
			qualified := objectProcedureQualifiedName(info.proc)
			key := objectSummaryKey(info.file.IR.Path, qualified, string(info.proc.ProcedureKind), info.proc.StartLine)
			previous := summaries[key]
			// Callee summaries must see the same module-level declarations as the
			// caller.  Without this, a procedure that initializes or returns a
			// module object is conservatively (and incorrectly) treated as
			// unrelated to that field.
			moduleDecls := moduleDeclarations(info.file.Lines, sourceProceduresFromIR(info.file.IR, info.file.CFG))
			declarations := objectFlowDeclarations(info.file.Lines, info.proc, moduleDecls)
			flow := objectStateFlow(info.file, info.proc, declarations, summaries, objectFlowOptions{Summary: true})
			updated := previous
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
			}
			for name, declaration := range moduleDecls {
				if !declaration.Object {
					continue
				}
				variable := objectVariable{Scope: procedureir.ScopeModule, Name: declaration.Name}
				updated.ModuleAssigned[strings.ToLower(name)] = !unknownFlow && objectFlowExitDefinitelyAssigned(flow, variable)
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

func objectSummaryEqual(a, b objectProcedureSummary) bool {
	if a.ReturnAssigned != b.ReturnAssigned || len(a.ByRefAssigned) != len(b.ByRefAssigned) || len(a.ModuleAssigned) != len(b.ModuleAssigned) {
		return false
	}
	for index, value := range a.ByRefAssigned {
		if b.ByRefAssigned[index] != value {
			return false
		}
	}
	for name, value := range a.ModuleAssigned {
		if b.ModuleAssigned[name] != value {
			return false
		}
	}
	return true
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

type objectFlowOptions struct {
	Summary bool
}

// objectUseBeforeSetIRFindings reports the first unsafe member/collection use
// of each object variable.  The state at a use comes from the CFG entry fact,
// not from source-line order, so branches, early exits, loops and error edges
// are all represented by the same must-analysis.
func (a Analyzer) objectUseBeforeSetIRFindings(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, summaries map[string]objectProcedureSummary) []Finding {
	if proc.Graph == nil {
		return nil
	}
	declarations := objectFlowDeclarations(file.Lines, proc, moduleDecls)
	flow := objectStateFlow(file, proc, declarations, summaries, objectFlowOptions{})
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
		key := strings.ToLower(cleanIdentifier(access.Name))
		declaration, ok := objectDeclarationFor(access.Name, access.Scope, declarations)
		// `As New` is guaranteed to produce a value on first use.  A normal
		// Static declaration, however, has module-lifetime state and may still
		// be Nothing on its first read; it therefore remains in the analysis.
		if !ok || !declaration.Object || declaration.NewExpression {
			continue
		}
		variable := objectVariable{Scope: access.Scope, Name: access.Name}
		if access.Scope == procedureir.ScopeUnresolved {
			variable.Scope = procedureir.ScopeLocal
		}
		block, ok := proc.Graph.BlockForStatement(access.StatementID)
		if !ok || !proc.Graph.IsReachable(block.ID, vbacfg.EdgeFilter{}) {
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
		key := strings.ToLower(cleanIdentifier(name))
		variable := objectFlowVariableForName(name, flow.vars)
		declaration, ok := objectDeclarationFor(name, variable.Scope, declarations)
		if !ok || !declaration.Object || declaration.NewExpression || reported[key] {
			continue
		}
		block, ok := proc.Graph.BlockForStatement(call.StatementID)
		if !ok || !proc.Graph.IsReachable(block.ID, vbacfg.EdgeFilter{}) || objectFlowAssigned(flow.in[block.ID], variable) {
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

type objectFlowResult struct {
	in         map[vbacfg.BlockID]map[string]bool
	out        map[vbacfg.BlockID]map[string]bool
	vars       map[string]objectVariable
	normalExit map[string]bool
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
	for key, declaration := range declarations {
		if !declaration.Object {
			continue
		}
		scope := procedureir.ScopeModule
		if declaration.Parameter {
			scope = procedureir.ScopeParameter
		} else if _, local := procedureDeclarations(file.Lines, proc)[key]; local {
			scope = procedureir.ScopeLocal
		}
		variable := objectVariable{Scope: scope, Name: declaration.Name}
		result.vars[variable.key()] = variable
	}
	if isObjectType(proc.ReturnType) {
		variable := objectVariable{Scope: procedureir.ScopeLocal, Name: proc.Name}
		result.vars[variable.key()] = variable
	}
	initial := map[string]bool{}
	for key, variable := range result.vars {
		initial[key] = false
		declaration, _ := objectDeclarationFor(variable.Name, variable.Scope, declarations)
		if declaration.NewExpression {
			initial[key] = true
		}
		// Parameters are intentionally not initialized at procedure entry.  A
		// caller can pass Nothing, and only a dominating Set or a proven ByRef
		// initializer establishes a safe value.
	}
	reachable := map[vbacfg.BlockID]bool{}
	for _, id := range proc.Graph.Reachable(vbacfg.EdgeFilter{}) {
		reachable[id] = true
	}
	for _, block := range proc.Graph.Blocks {
		if !reachable[block.ID] {
			continue
		}
		if block.ID == proc.Graph.Entry {
			result.in[block.ID] = cloneObjectState(initial)
		} else {
			result.in[block.ID] = objectStateAllTrue(result.vars)
		}
		result.out[block.ID] = objectFlowTransfer(file, proc, block, result.in[block.ID], result.vars, declarations, summaries)
	}
	changed := true
	for changed {
		changed = false
		for _, block := range proc.Graph.Blocks {
			if !reachable[block.ID] || block.ID == proc.Graph.Entry {
				continue
			}
			incoming := make([]map[string]bool, 0)
			for _, edge := range proc.Graph.Edges {
				if edge.To != block.ID || !reachable[edge.From] {
					continue
				}
				state := result.out[edge.From]
				if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain || objectFlowForEachZeroIteration(proc, edge) {
					state = result.in[edge.From]
				}
				state = objectFlowApplyGuard(state, proc, edge)
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
			updated := objectFlowTransfer(file, proc, block, next, result.vars, declarations, summaries)
			if !objectStateEqual(result.out[block.ID], updated) {
				result.out[block.ID] = updated
				changed = true
			}
		}
	}
	if proc.Graph != nil {
		result.normalExit = cloneObjectState(result.in[proc.Graph.NormalExit])
	}
	return result
}

// objectFlowApplyGuard refines the state on the edge selected by a direct
// `obj Is Nothing`/`Not obj Is Nothing` condition.  Compound boolean
// expressions are deliberately ignored; VBA212 remains responsible for
// short-circuit/eager Boolean diagnostics and this analysis stays conservative.
func objectFlowApplyGuard(state map[string]bool, proc sourceProcedure, edge vbacfg.Edge) map[string]bool {
	if edge.Kind != vbacfg.EdgeBranchTrue && edge.Kind != vbacfg.EdgeBranchFalse {
		return state
	}
	var condition *procedureir.Expression
	for _, statement := range proc.Statements {
		if statement.ID == edge.StatementID {
			condition = statement.Condition
			break
		}
	}
	if condition == nil {
		return state
	}
	text := strings.ToLower(strings.TrimSpace(condition.Text))
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
	ref := objectVariable{Scope: procedureir.ScopeLocal, Name: name}
	key := ref.key()
	if _, exists := state[key]; !exists {
		for _, scope := range []procedureir.SymbolScope{procedureir.ScopeParameter, procedureir.ScopeModule} {
			candidate := objectVariable{Scope: scope, Name: name}
			if _, ok := state[candidate.key()]; ok {
				key = candidate.key()
				break
			}
		}
	}
	nonNothingOnTrue := not
	if edge.Kind == vbacfg.EdgeBranchFalse {
		nonNothingOnTrue = !nonNothingOnTrue
	}
	if _, ok := state[key]; ok {
		updated := cloneObjectState(state)
		updated[key] = nonNothingOnTrue
		return updated
	}
	return state
}

func objectFlowExitDefinitelyAssigned(flow objectFlowResult, variable objectVariable) bool {
	return objectFlowAssigned(flow.normalExit, variable)
}

func objectFlowAssigned(state map[string]bool, variable objectVariable) bool {
	return state[variable.key()]
}

func objectFlowTransfer(file parsedFile, proc sourceProcedure, block vbacfg.Block, input map[string]bool, vars map[string]objectVariable, declarations map[string]sourceDeclaration, summaries map[string]objectProcedureSummary) map[string]bool {
	state := cloneObjectState(input)
	statement := block.Statement
	if statement == nil {
		return state
	}
	expressions := make(map[int]procedureir.Expression, len(proc.Expressions))
	for _, expression := range proc.Expressions {
		expressions[expression.ID] = expression
	}
	for _, call := range proc.Calls {
		if call.StatementID == statement.ID {
			applyObjectCallEffects(call, state, vars, declarations, expressions, summaries)
		}
	}
	target, ok := objectFlowTarget(proc, *statement, declarations)
	if !ok {
		return state
	}
	switch statement.Kind {
	case procedureir.StatementSet:
		state[target.key()] = objectFlowValueAssigned(proc, *statement, state, expressions, summaries)
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

func objectFlowTarget(proc sourceProcedure, statement procedureir.Statement, declarations map[string]sourceDeclaration) (objectVariable, bool) {
	for _, access := range proc.Accesses {
		if access.StatementID != statement.ID || (access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) {
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
	if statement.Target != nil {
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

func objectFlowValueAssigned(proc sourceProcedure, statement procedureir.Statement, state map[string]bool, expressions map[int]procedureir.Expression, summaries map[string]objectProcedureSummary) bool {
	value := statement.Value
	if value == nil {
		return false
	}
	return objectExpressionAssigned(proc, *value, state, expressions, summaries, statement.ID)
}

func objectExpressionAssigned(proc sourceProcedure, expression procedureir.Expression, state map[string]bool, expressions map[int]procedureir.Expression, summaries map[string]objectProcedureSummary, statementID int) bool {
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
				return objectExpressionAssigned(proc, nested, state, expressions, summaries, statementID)
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
			if objectCallReturnsAssigned(call, summaries) {
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
				if root == "thisworkbook" || root == "application" {
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

func objectCallReturnsAssigned(call procedureir.CallSite, summaries map[string]objectProcedureSummary) bool {
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
		return receiver == "thisworkbook" || receiver == "application" || receiver == "excel.application" ||
			strings.HasPrefix(receiver, "thisworkbook.") || strings.HasPrefix(receiver, "application.") || strings.HasPrefix(receiver, "excel.application.")
	}
	if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
		return false
	}
	summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries)
	return ok && summary.ReturnAssigned
}

func applyObjectCallEffects(call procedureir.CallSite, state map[string]bool, vars map[string]objectVariable, declarations map[string]sourceDeclaration, expressions map[int]procedureir.Expression, summaries map[string]objectProcedureSummary) {
	actuals := objectCallActuals(call, expressions)
	var candidates []objectProcedureSummary
	if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
		if summary, ok := objectSummaryForCandidate(call.Resolution.Candidates[0], summaries); ok {
			candidates = append(candidates, summary)
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
			if formalIndex < 0 || formalIndex >= len(summary.Params) || !summary.Params[formalIndex].Object || !summary.Params[formalIndex].ByRef {
				continue
			}
			known = true
			if !summary.ByRefAssigned[formalIndex] {
				assigned = false
				break
			}
			assigned = true
		}
		if !known || len(candidates) == 0 {
			// An unresolved/external call may mutate a ByRef object or leave it
			// Nothing; no definite state survives the call.
			state[variable.key()] = false
			continue
		}
		state[variable.key()] = assigned
	}
}

type objectCallActual struct {
	text  string
	scope procedureir.SymbolScope
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
		if expression.Kind != procedureir.ExpressionIdentifier {
			actuals = append(actuals, objectCallActual{})
			continue
		}
		actuals = append(actuals, objectCallActual{text: expression.Text, scope: procedureir.ScopeUnresolved})
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
	// Synthetic IR used by a few callers omits candidate identity fields.  Do
	// not let that weaken production resolution: accept the fallback only when
	// exactly one summary has the qualified name.
	var match objectProcedureSummary
	found := false
	for _, summary := range summaries {
		if !strings.EqualFold(summary.QualifiedName, candidate.QualifiedName) {
			continue
		}
		if found {
			return objectProcedureSummary{}, false
		}
		match, found = summary, true
	}
	if found && candidate.File == "" && candidate.Kind == "" && candidate.Line == 0 {
		return match, true
	}
	return objectProcedureSummary{}, false
}

func objectFlowForEachZeroIteration(proc sourceProcedure, edge vbacfg.Edge) bool {
	if edge.Kind != vbacfg.EdgeLoopExit {
		return false
	}
	for _, statement := range proc.Statements {
		if statement.ID == edge.StatementID && statement.Kind == procedureir.StatementForEach {
			return true
		}
	}
	return false
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
