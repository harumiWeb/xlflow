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
	proc        sourceProcedure
	summary     parameterMutationSummary
	udtTypes    map[string]bool
	objectTypes map[string]bool
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
	records := make(map[string]*parameterMutationRecord)
	ordered := make([]*parameterMutationRecord, 0)
	udtTypes, objectTypes := parameterCompositeTypes(files)
	for _, file := range files {
		for _, proc := range file.Procedures {
			if proc.IR == nil {
				continue
			}
			record := &parameterMutationRecord{
				proc: proc, summary: directParameterMutationSummary(proc, udtTypes, objectTypes),
				udtTypes: udtTypes, objectTypes: objectTypes,
			}
			registered := false
			for _, key := range parameterProcedureKeys(proc) {
				if _, exists := records[key]; !exists {
					records[key] = record
					registered = true
				}
			}
			if registered {
				ordered = append(ordered, record)
			}
		}
	}

	callers := parameterMutationCallers(ordered, records)
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
		if !propagateParameterCallMutations(record, records) {
			continue
		}
		for _, caller := range callers[record] {
			if !queued[caller] {
				queue = append(queue, caller)
				queued[caller] = true
			}
		}
	}

	out := make(map[string]parameterMutationSummary, len(records))
	for key, record := range records {
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

func parameterMutationCallers(ordered []*parameterMutationRecord, records map[string]*parameterMutationRecord) map[*parameterMutationRecord][]*parameterMutationRecord {
	callers := make(map[*parameterMutationRecord][]*parameterMutationRecord)
	seen := make(map[*parameterMutationRecord]map[*parameterMutationRecord]bool)
	for _, caller := range ordered {
		for call := range caller.proc.Calls.All() {
			if call.Resolution.Status != procedureir.ResolutionMatched {
				continue
			}
			for _, candidate := range call.Resolution.Candidates {
				callee := records[strings.ToLower(strings.TrimSpace(candidate.QualifiedName))]
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

func directParameterMutationSummary(proc sourceProcedure, udtTypes, objectTypes map[string]bool) parameterMutationSummary {
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
	for access := range proc.Accesses.All() {
		if access.Scope != procedureir.ScopeParameter {
			continue
		}
		name := parameterName(access.Name)
		if _, ok := summary[name]; !ok {
			continue
		}
		if access.Mode == procedureir.AccessWrite || access.Mode == procedureir.AccessReadWrite {
			summary[name] = parameterWritten
		}
	}
	for _, statement := range proc.IR.Statements {
		if statement.Target == nil || !strings.ContainsAny(statement.Target.Text, ".!") {
			continue
		}
		root := parameterMemberTargetRoot(statement.Target.Text)
		if _, ok := summary[root]; !ok {
			continue
		}
		parameter, ok := parameterByName(proc.IR.Symbol.Parameters, root)
		if !ok || parameterTypeIsObject(parameter.Type, objectTypes) {
			continue
		}
		state := parameterPossiblyWritten
		if parameterTypeIsUDT(parameter.Type, udtTypes) {
			state = parameterWritten
		}
		if state > summary[root] {
			summary[root] = state
		}
	}
	return summary
}

func parameterMemberTargetRoot(text string) string {
	text = strings.TrimSpace(text)
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

func propagateParameterCallMutations(record *parameterMutationRecord, records map[string]*parameterMutationRecord) bool {
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
			for _, nestedID := range parameterArgumentExpressionIDs(proc.IR, expressionID) {
				for _, access := range accessesByExpression[nestedID] {
					name := parameterName(access.Name)
					if _, ok := record.summary[name]; !ok || record.summary[name] == parameterWritten {
						continue
					}
					state := calledArgumentMutation(call, expressionID, records)
					if !parameterArgumentIsBareName(proc.IR, expressionID, name) {
						parameter, found := parameterByName(proc.IR.Symbol.Parameters, name)
						if found && parameterTypeIsObject(parameter.Type, record.objectTypes) {
							continue
						}
						if !found || !parameterTypeIsUDT(parameter.Type, record.udtTypes) {
							state = parameterPossiblyWritten
						}
					}
					if state > record.summary[name] {
						record.summary[name] = state
						changed = true
					}
				}
			}
		}
	}
	return changed
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

func calledArgumentMutation(call procedureir.CallSite, expressionID int, records map[string]*parameterMutationRecord) parameterMutationState {
	if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) == 0 {
		return parameterPossiblyWritten
	}
	states := make([]parameterMutationState, 0, len(call.Resolution.Candidates))
	for _, candidate := range call.Resolution.Candidates {
		record := records[strings.ToLower(strings.TrimSpace(candidate.QualifiedName))]
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

func parameterProcedureKeys(proc sourceProcedure) []string {
	qualified := strings.ToLower(strings.TrimSpace(proc.IR.Symbol.QualifiedName))
	moduleQualified := strings.ToLower(strings.TrimSpace(proc.Module + "." + proc.Name))
	if qualified == "" || qualified == moduleQualified {
		return []string{moduleQualified}
	}
	return []string{qualified, moduleQualified}
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
	summary := summaries[strings.ToLower(strings.TrimSpace(proc.IR.Symbol.QualifiedName))]
	if summary == nil {
		summary = summaries[strings.ToLower(strings.TrimSpace(proc.Module+"."+proc.Name))]
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
		case a.Config.Analyze.DetectMisleadingPropertyValueByRef && valueParameter && parameter.PassingExplicit && strings.EqualFold(parameter.Passing, "ByRef"):
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA273", "warning",
				"Property value parameter "+name+" is declared ByRef but VBA always passes it ByVal.",
				"The final value parameter of Property Let and Property Set has ByVal runtime semantics even when ByRef is written.",
				"Declare the property value parameter ByVal so the signature matches its actual behavior."))
		case a.Config.Analyze.DetectImplicitByRefParameters && !constrained && !valueParameter && effective == "byref" && !parameter.PassingExplicit:
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA270", "information",
				"Parameter "+name+" is implicitly passed ByRef.",
				"VBA defaults ordinary parameters to ByRef when no passing modifier is written.",
				"Add an explicit ByRef or ByVal modifier to document the intended API contract."))
		case a.Config.Analyze.DetectRedundantByRefModifiers && !constrained && !valueParameter && effective == "byref" && parameter.PassingExplicit:
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA274", "information",
				"Explicit ByRef on parameter "+name+" repeats VBA's default.",
				"Ordinary VBA parameters are already ByRef when the modifier is omitted.",
				"Remove the ByRef modifier when the project style relies on VBA's default."))
		}
		if a.Config.Analyze.DetectAssignedByValParameters && effective == "byval" && state == parameterWritten {
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA271", "warning",
				"ByVal parameter "+name+" is reassigned inside the procedure.",
				"The assignment changes only the procedure-local copy and is not visible to the caller.",
				"Use a separate local variable, or change the API contract only when caller-visible mutation is intended."))
		}
		if a.Config.Analyze.DetectByRefParametersCanBeByVal && !constrained && !valueParameter && effective == "byref" &&
			!parameter.IsArray && !parameter.ParamArray && state == parameterNotWritten {
			findings = append(findings, a.parameterFinding(file, proc, parameter, "VBA272", "information",
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
