package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type parameterMutationState uint8

const (
	parameterNotWritten parameterMutationState = iota
	parameterPossiblyWritten
	parameterWritten
)

type parameterMutationSummary map[string]parameterMutationState

type parameterMutationRecord struct {
	proc         sourceProcedure
	summary      parameterMutationSummary
	udtTypes     map[string]bool
	objectTypes  map[string]bool
	exprByID     map[int]procedureir.Expression
	withReceiver map[int]int
}

type parameterMutationWork struct {
	records     int
	evaluations int
}

func buildParameterMutationSummaries(files []parsedFile) map[string]parameterMutationSummary {
	summaries, _ := buildParameterMutationSummariesWithWork(files)
	return summaries
}

func buildParameterMutationSummariesWithWork(files []parsedFile) (map[string]parameterMutationSummary, parameterMutationWork) {
	index := parameterMutationIndex{
		byKey:  make(map[string]*parameterMutationRecord),
		byName: make(map[string][]*parameterMutationRecord),
	}
	ordered := make([]*parameterMutationRecord, 0)
	udtTypes, objectTypes := parameterCompositeTypes(files)
	for _, file := range files {
		for _, proc := range file.Procedures {
			if proc.IR == nil {
				continue
			}
			exprByID := parameterExpressionIndex(proc.IR)
			withReceiver := parameterWithReceivers(proc.IR)
			record := &parameterMutationRecord{
				proc:         proc,
				summary:      directParameterMutationSummary(proc, udtTypes, objectTypes, exprByID, withReceiver),
				udtTypes:     udtTypes,
				objectTypes:  objectTypes,
				exprByID:     exprByID,
				withReceiver: withReceiver,
			}
			registered := false
			for _, name := range parameterProcedureNames(proc) {
				index.byName[name] = append(index.byName[name], record)
			}
			for _, key := range parameterProcedureKeys(proc) {
				if _, exists := index.byKey[key]; !exists {
					index.byKey[key] = record
					registered = true
				}
			}
			if registered {
				ordered = append(ordered, record)
			}
		}
	}

	callers := parameterMutationCallers(ordered, index)
	queue := append([]*parameterMutationRecord(nil), ordered...)
	queued := make(map[*parameterMutationRecord]bool, len(ordered))
	for _, record := range queue {
		queued[record] = true
	}
	work := parameterMutationWork{records: len(ordered)}
	for len(queue) > 0 {
		record := queue[0]
		queue = queue[1:]
		queued[record] = false
		work.evaluations++
		if !propagateParameterCallMutations(record, index) {
			continue
		}
		for _, caller := range callers[record] {
			if !queued[caller] {
				queue = append(queue, caller)
				queued[caller] = true
			}
		}
	}

	out := make(map[string]parameterMutationSummary, len(index.byKey))
	for key, record := range index.byKey {
		out[key] = record.summary
	}
	return out, work
}

func parameterCompositeTypes(files []parsedFile) (map[string]bool, map[string]bool) {
	udtTypes := make(map[string]bool)
	objectTypes := make(map[string]bool)
	for _, file := range files {
		module := parameterName(file.IR.ModuleName)
		if module == "" {
			module = parameterName(file.Module)
		}
		moduleKind := file.ModuleKind
		if moduleKind == "" {
			moduleKind = file.IR.ModuleKind
		}
		if moduleKind != "" && !strings.EqualFold(moduleKind, "standard") {
			objectTypes[module] = true
		}
		for _, declaration := range file.IR.Declarations {
			if !strings.EqualFold(declaration.Kind, "type") {
				continue
			}
			name := parameterName(declaration.Name)
			udtTypes[name] = true
			if module != "" {
				udtTypes[module+"."+name] = true
			}
		}
	}
	return udtTypes, objectTypes
}

func parameterExpressionIndex(procedure *procedureir.ProcedureIR) map[int]procedureir.Expression {
	if procedure == nil {
		return nil
	}
	byID := make(map[int]procedureir.Expression, len(procedure.Expressions))
	for _, expression := range procedure.Expressions {
		byID[expression.ID] = expression
	}
	return byID
}

func parameterWithReceivers(procedure *procedureir.ProcedureIR) map[int]int {
	if procedure == nil {
		return nil
	}
	return withReceiverExpressions(*procedure)
}

// parameterMutationIndex stores mutation records under kind-qualified keys
// so same-named declarations with different procedure kinds (for example a
// Property Get and Property Let pair) keep separate summaries. The name
// index supports call candidates whose kind does not disambiguate: a unique
// name match is reused, while a collision resolves to possiblyWritten.
type parameterMutationIndex struct {
	byKey  map[string]*parameterMutationRecord
	byName map[string][]*parameterMutationRecord
}

func (index parameterMutationIndex) lookup(qualifiedName string, kind procedureir.ProcedureKind) *parameterMutationRecord {
	name := strings.ToLower(strings.TrimSpace(qualifiedName))
	if name == "" {
		return nil
	}
	if k := parameterKindPrefix(kind); k != "procedure" {
		if record := index.byKey[k+"|"+name]; record != nil {
			return record
		}
	}
	matches := index.byName[name]
	if len(matches) == 1 {
		return matches[0]
	}
	return nil
}

func parameterMutationCallers(ordered []*parameterMutationRecord, index parameterMutationIndex) map[*parameterMutationRecord][]*parameterMutationRecord {
	callers := make(map[*parameterMutationRecord][]*parameterMutationRecord)
	seen := make(map[*parameterMutationRecord]map[*parameterMutationRecord]bool)
	for _, caller := range ordered {
		for call := range caller.proc.Calls.All() {
			if call.Resolution.Status != procedureir.ResolutionMatched {
				continue
			}
			for _, candidate := range call.Resolution.Candidates {
				callee := index.lookup(candidate.QualifiedName, procedureir.ProcedureKind(candidate.Kind))
				if callee == nil {
					continue
				}
				if seen[callee] == nil {
					seen[callee] = make(map[*parameterMutationRecord]bool)
				}
				if !seen[callee][caller] {
					callers[callee] = append(callers[callee], caller)
					seen[callee][caller] = true
				}
			}
		}
	}
	return callers
}

func directParameterMutationSummary(proc sourceProcedure, udtTypes, objectTypes map[string]bool, exprByID map[int]procedureir.Expression, withReceiver map[int]int) parameterMutationSummary {
	summary := make(parameterMutationSummary, proc.Params.Len())
	for parameter := range proc.Params.All() {
		summary[parameterName(parameter.Name)] = parameterNotWritten
	}
	if proc.IR == nil || proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 {
		for name := range summary {
			summary[name] = parameterPossiblyWritten
		}
		return summary
	}
	accessesByExpression := make(map[int][]procedureir.VariableAccess)
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeParameter {
			continue
		}
		if access.ExpressionID != 0 {
			accessesByExpression[access.ExpressionID] = append(accessesByExpression[access.ExpressionID], access)
		}
		name := parameterName(access.Name)
		if _, ok := summary[name]; !ok {
			continue
		}
		if access.Mode == procedureir.AccessWrite || access.Mode == procedureir.AccessReadWrite {
			// v(0) = 1 writes through the binding to an element instead of
			// replacing it: not a VBA274 reassignment, but still possibly
			// caller-visible, so VBA275 must stay suppressed.
			if parameterElementWrite(exprByID, access.ExpressionID) {
				summary[name] = max(summary[name], parameterPossiblyWritten)
				continue
			}
			summary[name] = parameterWritten
		}
	}
	for _, statement := range proc.IR.Statements {
		if statement.Target != nil && strings.ContainsAny(statement.Target.Text, ".!") {
			root := parameterMemberTargetRoot(statement.Target.Text)
			if _, ok := summary[root]; !ok && statement.Target.Kind == procedureir.ExpressionMember {
				// Implicit member targets (.Left = 1) resolve through the
				// enclosing With receiver instead of the leading text. A
				// text root that matches no parameter (for example a
				// bracketed name truncated at an interior space) falls back
				// to expression resolution as well.
				root = parameterExprRootName(exprByID, withReceiver, statement.Target.ID, 0)
			}
			recordParameterMemberWrite(summary, proc.IR.Symbol.Parameters, root, udtTypes, objectTypes)
		}
		if statement.Kind != procedureir.StatementUnknown {
			continue
		}
		// Erase, Input #, Line Input #, and Get # write their variable
		// operands but surface as unknown statements with read accesses.
		switch statement.SyntaxKind {
		case "erase_statement", "input_statement", "line_input_statement":
			for _, expressionID := range statement.ExpressionIDs {
				expression, ok := exprByID[expressionID]
				if !ok || expression.SyntaxKind == "file_number_literal" {
					continue
				}
				recordParameterWriteTarget(summary, proc, udtTypes, objectTypes, exprByID, withReceiver, accessesByExpression, expressionID)
			}
		case "get_statement":
			if statement.TargetID != 0 {
				recordParameterWriteTarget(summary, proc, udtTypes, objectTypes, exprByID, withReceiver, accessesByExpression, statement.TargetID)
				continue
			}
			for _, expressionID := range statement.ExpressionIDs {
				expression, ok := exprByID[expressionID]
				if !ok || expression.SyntaxKind == "file_number_literal" {
					continue
				}
				markParameterSubtreePossiblyWritten(summary, proc.IR, accessesByExpression, expressionID)
			}
		}
	}
	return summary
}

// parameterElementWrite reports whether a write-mode access targets an
// element reached through the binding (v(0) = 1) rather than replacing the
// binding itself.
func parameterElementWrite(exprByID map[int]procedureir.Expression, expressionID int) bool {
	expression, ok := exprByID[expressionID]
	if !ok {
		return false
	}
	parent, ok := exprByID[expression.ParentID]
	if !ok {
		return false
	}
	return parent.Kind == procedureir.ExpressionCall || parent.Kind == procedureir.ExpressionMember
}

// parameterExprRootName resolves the variable an expression is rooted at.
// Qualified members recurse into the receiver; implicit members (.Left)
// resolve through the innermost enclosing With receiver.
func parameterExprRootName(exprByID map[int]procedureir.Expression, withReceiver map[int]int, expressionID, depth int) string {
	if depth > 16 {
		return ""
	}
	expression, ok := exprByID[expressionID]
	if !ok {
		return ""
	}
	switch expression.Kind {
	case procedureir.ExpressionIdentifier:
		return parameterName(expression.Text)
	case procedureir.ExpressionMember:
		_, receiver := udtMemberParts(expression, exprByID)
		if receiver != 0 {
			return parameterExprRootName(exprByID, withReceiver, receiver, depth+1)
		}
		if receiver, ok := withReceiver[expression.StatementID]; ok {
			return parameterExprRootName(exprByID, withReceiver, receiver, depth+1)
		}
		return ""
	case procedureir.ExpressionParentheses:
		if len(expression.Children) == 1 {
			return parameterExprRootName(exprByID, withReceiver, expression.Children[0], depth+1)
		}
		return ""
	case procedureir.ExpressionCall:
		// arr(0) is rooted at its callee: the callee identifier is metadata
		// recorded as the first child expression but emits no access.
		if len(expression.Children) > 0 {
			return parameterExprRootName(exprByID, withReceiver, expression.Children[0], depth+1)
		}
		return ""
	default:
		return ""
	}
}

func recordParameterMemberWrite(summary parameterMutationSummary, parameters []procedureir.Parameter, root string, udtTypes, objectTypes map[string]bool) {
	if _, ok := summary[root]; !ok {
		return
	}
	parameter, ok := parameterByName(parameters, root)
	if !ok {
		return
	}
	state := parameterPossiblyWritten
	switch {
	case parameterTypeIsUDT(parameter.Type, udtTypes):
		// A user-defined type shadows builtin object type names in module
		// scope, so the UDT check runs before the object check.
		state = parameterWritten
	case parameterTypeIsObject(parameter.Type, objectTypes):
		// Object member writes never rebind the reference.
		return
	}
	summary[root] = max(summary[root], state)
}

// recordParameterWriteTarget records a write to an expression used as a
// statement operand (Erase, Input #, Get #). Bare parameters are written;
// member targets follow the member-write rules; anything else fails open on
// every parameter read inside the target subtree.
func recordParameterWriteTarget(summary parameterMutationSummary, proc sourceProcedure, udtTypes, objectTypes map[string]bool, exprByID map[int]procedureir.Expression, withReceiver map[int]int, accessesByExpression map[int][]procedureir.VariableAccess, expressionID int) {
	expression, ok := exprByID[expressionID]
	if !ok {
		return
	}
	switch expression.Kind {
	case procedureir.ExpressionIdentifier:
		name := parameterName(expression.Text)
		if _, ok := summary[name]; ok {
			summary[name] = parameterWritten
		}
	case procedureir.ExpressionMember:
		root := parameterExprRootName(exprByID, withReceiver, expressionID, 0)
		recordParameterMemberWrite(summary, proc.IR.Symbol.Parameters, root, udtTypes, objectTypes)
	case procedureir.ExpressionCall:
		// Erase arr(0) or Get #1,,arr(0) writes through one element: the
		// binding survives, but the caller-visible contents may change.
		if name := parameterExprRootName(exprByID, withReceiver, expressionID, 0); name != "" {
			if _, ok := summary[name]; ok {
				summary[name] = max(summary[name], parameterPossiblyWritten)
			}
		}
	default:
		markParameterSubtreePossiblyWritten(summary, proc.IR, accessesByExpression, expressionID)
	}
}

func markParameterSubtreePossiblyWritten(summary parameterMutationSummary, procedure *procedureir.ProcedureIR, accessesByExpression map[int][]procedureir.VariableAccess, expressionID int) {
	for _, nestedID := range parameterArgumentExpressionIDs(procedure, expressionID) {
		for _, access := range accessesByExpression[nestedID] {
			name := parameterName(access.Name)
			if _, ok := summary[name]; ok {
				summary[name] = max(summary[name], parameterPossiblyWritten)
			}
		}
	}
}

func parameterMemberTargetRoot(text string) string {
	text = strings.TrimSpace(text)
	if strings.HasPrefix(text, "[") {
		// Bracketed identifiers ([My Field].Left) may contain any
		// characters up to the closing bracket; consume them whole.
		if close := strings.IndexByte(text, ']'); close >= 0 {
			return parameterName(text[:close+1])
		}
	}
	end := 0
	for end < len(text) {
		char := text[end]
		if (char >= 'a' && char <= 'z') || (char >= 'A' && char <= 'Z') ||
			(char >= '0' && char <= '9') || char == '_' || char == '[' || char == ']' {
			end++
			continue
		}
		break
	}
	return parameterName(text[:end])
}

func parameterByName(parameters []procedureir.Parameter, name string) (procedureir.Parameter, bool) {
	for _, parameter := range parameters {
		if parameterName(parameter.Name) == name {
			return parameter, true
		}
	}
	return procedureir.Parameter{}, false
}

func parameterTypeIsObject(typeName string, objectTypes map[string]bool) bool {
	normalized := parameterName(typeName)
	return isObjectType(typeName) || objectTypes[normalized] || objectTypes[parameterName(lastName(normalized))]
}

func parameterTypeIsUDT(typeName string, udtTypes map[string]bool) bool {
	normalized := parameterName(typeName)
	return udtTypes[normalized] || udtTypes[parameterName(lastName(normalized))]
}

func propagateParameterCallMutations(record *parameterMutationRecord, index parameterMutationIndex) bool {
	proc := record.proc
	accessesByExpression := make(map[int][]procedureir.VariableAccess)
	for access := range proc.Accesses.All() {
		if access.Scope == procedureir.ScopeParameter && access.ExpressionID != 0 {
			accessesByExpression[access.ExpressionID] = append(accessesByExpression[access.ExpressionID], access)
		}
	}
	changed := false
	for call := range proc.Calls.All() {
		for _, expressionID := range call.Arguments.ExpressionIDs {
			// A parenthesized argument ((x)) forces ByVal evaluation: the
			// callee receives a value and cannot write the caller binding.
			if expression, ok := record.exprByID[expressionID]; ok && expression.Kind == procedureir.ExpressionParentheses {
				continue
			}
			handled := make(map[string]bool)
			for _, nestedID := range parameterArgumentExpressionIDs(proc.IR, expressionID) {
				for _, access := range accessesByExpression[nestedID] {
					name := parameterName(access.Name)
					handled[name] = true
					bareName := parameterArgumentIsBareName(proc.IR, expressionID, name)
					if record.applyArgumentMutation(call, expressionID, name, bareName, index) {
						changed = true
					}
				}
			}
			// Arguments whose root emits no receiver access — an implicit
			// member (Call Replace(.Left)) inside a With block, or an
			// indexed call (arr(0)) whose callee identifier is metadata —
			// resolve the expression root directly.
			if name := parameterExprRootName(record.exprByID, record.withReceiver, expressionID, 0); name != "" && !handled[name] {
				if record.applyArgumentMutation(call, expressionID, name, false, index) {
					changed = true
				}
			}
		}
		// A space-separated implicit member argument (Replace .Left inside a
		// With block) parses into the callee expression instead of the
		// argument list. When the unresolved callee shows a spaced member
		// operator, fail open on the enclosing With receiver parameter.
		if call.Resolution.Status != procedureir.ResolutionMatched {
			if callee, ok := record.exprByID[call.ExpressionID]; ok &&
				callee.Kind == procedureir.ExpressionMember && strings.Contains(callee.Text, " .") {
				if receiverID, ok := record.withReceiver[call.StatementID]; ok {
					if name := parameterExprRootName(record.exprByID, record.withReceiver, receiverID, 0); name != "" {
						if _, ok := record.summary[name]; ok && record.summary[name] < parameterPossiblyWritten {
							record.summary[name] = parameterPossiblyWritten
							changed = true
						}
					}
				}
			}
		}
	}
	return changed
}

// applyArgumentMutation records how one argument expression may mutate the
// caller-visible parameter binding. bareName arguments can replace the
// binding outright; member or element arguments follow the member-write
// rules (object member writes do not rebind, UDT member writes propagate the
// callee state, unknown composites fail open).
func (record *parameterMutationRecord) applyArgumentMutation(call procedureir.CallSite, expressionID int, name string, bareName bool, index parameterMutationIndex) bool {
	if _, ok := record.summary[name]; !ok || record.summary[name] == parameterWritten {
		return false
	}
	state := calledArgumentMutation(call, expressionID, index)
	if !bareName {
		parameter, found := parameterByName(record.proc.IR.Symbol.Parameters, name)
		switch {
		case found && parameterTypeIsUDT(parameter.Type, record.udtTypes):
			// A user-defined type shadows builtin object type names in module
			// scope, so the UDT check runs before the object check and member
			// arguments propagate the callee state.
		case found && parameterTypeIsObject(parameter.Type, record.objectTypes):
			// Object member writes never rebind the reference.
			return false
		default:
			state = parameterPossiblyWritten
		}
	}
	if state > record.summary[name] {
		record.summary[name] = state
		return true
	}
	return false
}

func parameterArgumentIsBareName(procedure *procedureir.ProcedureIR, expressionID int, name string) bool {
	if procedure == nil {
		return false
	}
	for _, expression := range procedure.Expressions {
		if expression.ID == expressionID {
			return expression.Kind == procedureir.ExpressionIdentifier && parameterName(expression.Text) == name
		}
	}
	return false
}

func parameterArgumentExpressionIDs(procedure *procedureir.ProcedureIR, root int) []int {
	if procedure == nil || root == 0 {
		return nil
	}
	byID := make(map[int]procedureir.Expression, len(procedure.Expressions))
	for _, expression := range procedure.Expressions {
		byID[expression.ID] = expression
	}
	ids := make([]int, 0, 4)
	seen := make(map[int]bool)
	var visit func(int)
	visit = func(id int) {
		if id == 0 || seen[id] {
			return
		}
		seen[id] = true
		ids = append(ids, id)
		for _, child := range byID[id].Children {
			visit(child)
		}
	}
	visit(root)
	return ids
}

func calledArgumentMutation(call procedureir.CallSite, expressionID int, index parameterMutationIndex) parameterMutationState {
	if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) == 0 {
		return parameterPossiblyWritten
	}
	states := make([]parameterMutationState, 0, len(call.Resolution.Candidates))
	for _, candidate := range call.Resolution.Candidates {
		record := index.lookup(candidate.QualifiedName, procedureir.ProcedureKind(candidate.Kind))
		if record == nil {
			return parameterPossiblyWritten
		}
		index, ok := callArgumentParameterIndex(call, expressionID, record.proc.IR.Symbol.Parameters)
		if !ok {
			return parameterPossiblyWritten
		}
		parameter := record.proc.IR.Symbol.Parameters[index]
		if effectiveParameterPassing(record.proc.IR.Symbol, index) == "byval" || parameter.ParamArray {
			states = append(states, parameterNotWritten)
			continue
		}
		states = append(states, record.summary[parameterName(parameter.Name)])
	}
	if len(states) == 0 {
		return parameterPossiblyWritten
	}
	state := states[0]
	for _, candidateState := range states[1:] {
		if candidateState != state {
			return parameterPossiblyWritten
		}
	}
	return state
}

func callArgumentParameterIndex(call procedureir.CallSite, expressionID int, parameters []procedureir.Parameter) (int, bool) {
	for _, named := range call.Arguments.Named {
		if named.ExpressionID != expressionID {
			continue
		}
		for index, parameter := range parameters {
			if parameterName(parameter.Name) == parameterName(named.Name) {
				return index, true
			}
		}
		return 0, false
	}
	position := 0
	for _, id := range call.Arguments.ExpressionIDs {
		if id == expressionID {
			return position, position < len(parameters)
		}
		position++
	}
	return 0, false
}

func parameterKindPrefix(kind procedureir.ProcedureKind) string {
	normalized := strings.ToLower(strings.TrimSpace(string(kind)))
	if normalized == "" {
		return "procedure"
	}
	return normalized
}

func parameterProcedureNames(proc sourceProcedure) []string {
	qualified := strings.ToLower(strings.TrimSpace(proc.IR.Symbol.QualifiedName))
	moduleQualified := strings.ToLower(strings.TrimSpace(proc.Module + "." + proc.Name))
	if qualified == "" || qualified == moduleQualified {
		return []string{moduleQualified}
	}
	return []string{qualified, moduleQualified}
}

func parameterProcedureKeys(proc sourceProcedure) []string {
	kind := parameterKindPrefix(proc.IR.Symbol.Kind)
	names := parameterProcedureNames(proc)
	keys := make([]string, 0, len(names))
	for _, name := range names {
		keys = append(keys, kind+"|"+name)
	}
	return keys
}

func effectiveParameterPassing(symbol procedureir.ProcedureSymbol, index int) string {
	if propertyValueParameter(symbol, index) {
		return "byval"
	}
	if index < 0 || index >= len(symbol.Parameters) {
		return ""
	}
	return strings.ToLower(strings.TrimSpace(symbol.Parameters[index].Passing))
}

func propertyValueParameter(symbol procedureir.ProcedureSymbol, index int) bool {
	return (symbol.Kind == procedureir.ProcedurePropertyLet || symbol.Kind == procedureir.ProcedurePropertySet) &&
		index == len(symbol.Parameters)-1 && index >= 0
}

func (a Analyzer) parameterPassingFindings(file parsedFile, proc sourceProcedure, summaries map[string]parameterMutationSummary) []Finding {
	if proc.IR == nil || proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 {
		return nil
	}
	var summary parameterMutationSummary
	for _, key := range parameterProcedureKeys(proc) {
		if s := summaries[key]; s != nil {
			summary = s
			break
		}
	}
	if summary == nil {
		// Fail open: without a mutation record xlflow cannot prove the
		// parameter is never written, so no parameter-passing finding is
		// emitted for this procedure.
		return nil
	}
	constrained := parameterPassingSignatureConstrained(file, proc)
	var findings []Finding
	for index, parameter := range proc.IR.Symbol.Parameters {
		name := cleanIdentifier(parameter.Name)
		if name == "" || parameter.Recovered {
			continue
		}
		valueParameter := propertyValueParameter(proc.IR.Symbol, index)
		effective := effectiveParameterPassing(proc.IR.Symbol, index)
		state := summary[parameterName(name)]
		switch {
		case a.Config.Analyze.DetectMisleadingPropertyValueByRef && valueParameter && !constrained && parameter.PassingExplicit && strings.EqualFold(parameter.Passing, "ByRef"):
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA276", "warning",
				"Property value parameter "+name+" is declared ByRef but VBA always passes it ByVal.",
				"The final value parameter of Property Let and Property Set has ByVal runtime semantics even when ByRef is written.",
				"Declare the property value parameter ByVal so the signature matches its actual behavior."))
		case a.Config.Analyze.DetectImplicitByRefParameters && !constrained && !valueParameter && !parameter.ParamArray && effective == "byref" && !parameter.PassingExplicit:
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA273", "information",
				"Parameter "+name+" is implicitly passed ByRef.",
				"VBA defaults ordinary parameters to ByRef when no passing modifier is written.",
				"Add an explicit ByRef or ByVal modifier to document the intended API contract."))
		case a.Config.Analyze.DetectRedundantByRefModifiers && !constrained && !valueParameter && !parameter.ParamArray && effective == "byref" && parameter.PassingExplicit:
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA277", "information",
				"Explicit ByRef on parameter "+name+" repeats VBA's default.",
				"Ordinary VBA parameters are already ByRef when the modifier is omitted.",
				"Remove the ByRef modifier when the project style relies on VBA's default."))
		}
		if a.Config.Analyze.DetectAssignedByValParameters && effective == "byval" && state == parameterWritten {
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA274", "warning",
				"ByVal parameter "+name+" is reassigned inside the procedure.",
				"The assignment changes only the procedure-local copy and is not visible to the caller.",
				"Use a separate local variable, or change the API contract only when caller-visible mutation is intended."))
		}
		if a.Config.Analyze.DetectByRefParametersCanBeByVal && !constrained && !valueParameter && effective == "byref" &&
			!parameter.IsArray && !parameter.ParamArray && state == parameterNotWritten {
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA275", "information",
				"ByRef parameter "+name+" is never written and can be passed ByVal.",
				"No direct assignment or conservatively modeled ByRef call can replace the caller's argument.",
				"Declare the parameter ByVal to prevent unintended caller-visible reassignment."))
		}
	}
	return findings
}

func (a Analyzer) parameterFinding(file parsedFile, proc sourceProcedure, parameter procedureir.Parameter, code, severity, message, reason, suggestion string) Finding {
	rng := parameter.Range
	if parameter.PassingRange != nil {
		rng = *parameter.PassingRange
	}
	finding := a.simpleFinding(file, proc, rng.StartLine, code, severity, message, reason, suggestion)
	finding.Column = rng.StartColumn
	finding.EndLine = rng.EndLine
	finding.EndColumn = rng.EndColumn
	return finding
}

func parameterPassingSignatureConstrained(file parsedFile, proc sourceProcedure) bool {
	if proc.IR == nil || proc.IR.Symbol.IsEventHandler || eventHandlerKind(file, proc) != "" {
		return true
	}
	facts := file.moduleAnalysisFacts().unusedDeclarationFacts(file.Lines)
	if signatureConstrainedEvent(file, proc, facts) {
		return true
	}
	name := strings.ToLower(cleanIdentifier(proc.IR.Symbol.Name))
	for _, iface := range facts.implementsTargets {
		if strings.HasPrefix(name, iface+"_") {
			return true
		}
	}
	return false
}

func parameterName(name string) string {
	return strings.ToLower(strings.TrimSpace(cleanIdentifier(name)))
}
