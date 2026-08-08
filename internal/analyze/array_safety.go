package analyze

import (
	"regexp"
	"sort"
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// arrayAllocation is deliberately a three-point lattice.  In particular,
// unknown is not treated as allocated: a VBA runtime operation must be
// proven safe on every path before it is allowed through the analysis.
type arrayAllocation uint8

const (
	arrayUnknown arrayAllocation = iota
	arrayUnallocated
	arrayAllocated
)

type arrayBound struct {
	known bool
	value int
}

type arrayDimension struct {
	lower arrayBound
	upper arrayBound
}

type arrayOrigin uint8

const (
	arrayOriginUnknown arrayOrigin = iota
	arrayOriginLocal
	arrayOriginRangeValue
)

type arrayVariable struct {
	name       string
	typ        string
	isArray    bool
	isVariant  bool
	isObject   bool
	fixed      bool
	parameter  bool
	dimensions []arrayDimension
}

type arrayValue struct {
	kind       arrayAllocation
	knownArray bool
	dimensions []arrayDimension
	origin     arrayOrigin
}

type arrayFlowState map[string]arrayValue

type arrayUse struct {
	name string
	args []string
}

var (
	arrayRedimRe       = regexp.MustCompile(`(?i)^\s*redim\s+(preserve\s+)?(.+)$`)
	arrayRedimClauseRe = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?)?\s*$`)
	arrayEraseRe       = regexp.MustCompile(`(?i)^\s*erase\s+(.+)$`)
	arrayEraseNameRe   = regexp.MustCompile(`(?i)^[A-Za-z_]\w*$`)
	arrayBoundCallRe   = regexp.MustCompile(`(?i)\b(lbound|ubound)\s*\(\s*([A-Za-z_]\w*)\s*(?:,\s*([^)]*))?\)`)
	arrayForBoundRe    = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*([-+]?\d+)\s+to\s+(?:lbound|ubound)\s*\(\s*([A-Za-z_]\w*)`)
)

func (a Analyzer) arrayLifecycleFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []Finding {
	variables := arrayVariables(file, proc, moduleDecls)
	if len(variables) == 0 {
		return nil
	}
	objectArrayDiagnosticsApplicable := false
	for _, variable := range variables {
		if variable.isArray && variable.isObject {
			objectArrayDiagnosticsApplicable = true
			break
		}
	}
	if !a.Config.Analyze.DetectArrayLifecycleSafety && !a.Config.Analyze.DetectRedimPreserveDimension && !a.Config.Analyze.DetectObjectArrayComparison && !objectArrayDiagnosticsApplicable {
		return nil
	}
	initial := arrayInitialState(variables)
	if proc.Graph == nil {
		return a.arrayLifecycleLinearFindings(file, proc, ctx, variables, initial)
	}

	var findings []Finding
	seen := map[string]bool{}
	walkArrayCFG(proc.Graph, file.Lines, initial, func(text string, line int, in arrayFlowState) arrayFlowState {
		out, issues := a.arrayTransfer(file, proc, ctx, variables, in, text, line)
		for _, finding := range issues {
			key := finding.Code + ":" + strconv.Itoa(finding.Line) + ":" + finding.Message
			if !seen[key] {
				seen[key] = true
				findings = append(findings, finding)
			}
		}
		return out
	})
	sortFindings(findings)
	return findings
}

func (a Analyzer) arrayLifecycleLinearFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState) []Finding {
	var findings []Finding
	seen := map[string]bool{}
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		var issues []Finding
		state, issues = a.arrayTransfer(file, proc, ctx, variables, state, normalizedCodeLine(file.Lines[line-1]), line)
		for _, finding := range issues {
			key := finding.Code + ":" + strconv.Itoa(finding.Line) + ":" + finding.Message
			if !seen[key] {
				seen[key] = true
				findings = append(findings, finding)
			}
		}
	}
	sortFindings(findings)
	return findings
}

// walkArrayCFG owns the common allocation-state worklist used by both the
// procedure findings pass and the array-return summary pass. Exceptional and
// uncertain edges retain the predecessor's input state because the statement
// may not have completed before control leaves the block.
func walkArrayCFG(graph *vbacfg.Graph, lines []string, initial arrayFlowState, visit func(text string, line int, in arrayFlowState) arrayFlowState) {
	if graph == nil {
		return
	}
	blocks := make(map[vbacfg.BlockID]vbacfg.Block, len(graph.Blocks))
	outgoing := make(map[vbacfg.BlockID][]vbacfg.Edge)
	for _, block := range graph.Blocks {
		blocks[block.ID] = block
	}
	for _, edge := range graph.Edges {
		outgoing[edge.From] = append(outgoing[edge.From], edge)
	}
	inStates := map[vbacfg.BlockID]arrayFlowState{graph.Entry: initial}
	queued := map[vbacfg.BlockID]bool{graph.Entry: true}
	for len(queued) > 0 {
		// Block IDs are ordered so the worklist cannot inherit Go map iteration
		// order and change the fixed-point path through the array state lattice.
		var id vbacfg.BlockID
		first := true
		for candidate := range queued {
			if first || candidate < id {
				id = candidate
				first = false
			}
		}
		delete(queued, id)
		in := cloneArrayState(inStates[id])
		block, ok := blocks[id]
		if !ok {
			continue
		}
		out := in
		if block.Statement != nil {
			line := block.Statement.Range.StartLine
			if line == 0 {
				line = block.Range.StartLine
			}
			text := block.Statement.Text
			if strings.TrimSpace(text) == "" && line >= 1 && line <= len(lines) {
				text = normalizedCodeLine(lines[line-1])
			}
			out = visit(text, line, in)
		}
		for _, edge := range outgoing[id] {
			next := out
			if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				next = in
			}
			if mergeArrayState(inStates, edge.To, next) {
				queued[edge.To] = true
			}
		}
	}
}

func mergeArrayState(states map[vbacfg.BlockID]arrayFlowState, id vbacfg.BlockID, incoming arrayFlowState) bool {
	current, exists := states[id]
	if !exists {
		states[id] = cloneArrayState(incoming)
		return true
	}
	merged := meetArrayState(current, incoming)
	if arrayStateEqual(current, merged) {
		return false
	}
	states[id] = merged
	return true
}

func meetArrayState(left, right arrayFlowState) arrayFlowState {
	out := arrayFlowState{}
	keys := map[string]bool{}
	for key := range left {
		keys[key] = true
	}
	for key := range right {
		keys[key] = true
	}
	for key := range keys {
		l, lok := left[key]
		r, rok := right[key]
		if !lok || !rok {
			out[key] = arrayValue{kind: arrayUnknown}
			continue
		}
		out[key] = meetArrayValue(l, r)
	}
	return out
}

func meetArrayValue(left, right arrayValue) arrayValue {
	out := arrayValue{kind: left.kind, knownArray: left.knownArray, origin: left.origin, dimensions: append([]arrayDimension(nil), left.dimensions...)}
	if left.kind != right.kind {
		out.kind = arrayUnknown
	}
	if left.knownArray != right.knownArray {
		out.knownArray = false
	}
	if left.origin != right.origin {
		out.origin = arrayOriginUnknown
	}
	if !arrayDimensionsEqual(left.dimensions, right.dimensions) {
		out.dimensions = nil
	}
	return out
}

func cloneArrayState(state arrayFlowState) arrayFlowState {
	out := arrayFlowState{}
	for name, value := range state {
		value.dimensions = append([]arrayDimension(nil), value.dimensions...)
		out[name] = value
	}
	return out
}

func arrayStateEqual(left, right arrayFlowState) bool {
	if len(left) != len(right) {
		return false
	}
	for key, l := range left {
		r, ok := right[key]
		if !ok || l.kind != r.kind || l.knownArray != r.knownArray || l.origin != r.origin || !arrayDimensionsEqual(l.dimensions, r.dimensions) {
			return false
		}
	}
	return true
}

func arrayDimensionsEqual(left, right []arrayDimension) bool {
	if len(left) != len(right) {
		return false
	}
	for i := range left {
		if left[i] != right[i] {
			return false
		}
	}
	return true
}

func arrayInitialState(variables map[string]arrayVariable) arrayFlowState {
	state := arrayFlowState{}
	for name, variable := range variables {
		if variable.isVariant && !variable.isArray {
			state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
			continue
		}
		value := arrayValue{knownArray: variable.isArray, origin: arrayOriginLocal, dimensions: append([]arrayDimension(nil), variable.dimensions...)}
		if variable.parameter {
			value.kind = arrayUnknown
		} else if variable.fixed {
			value.kind = arrayAllocated
		} else {
			value.kind = arrayUnallocated
		}
		state[name] = value
	}
	return state
}

func (a Analyzer) arrayTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int) (arrayFlowState, []Finding) {
	state = cloneArrayState(state)
	var findings []Finding
	add := func(code, message, reason, suggestion string) {
		if code == "VBA227" && !a.Config.Analyze.DetectArrayLifecycleSafety {
			return
		}
		if code == "VBA208" && !a.Config.Analyze.DetectRedimPreserveDimension {
			return
		}
		if code == "VBA209" && !a.Config.Analyze.DetectObjectArrayComparison {
			return
		}
		findings = append(findings, a.simpleFinding(file, proc, line, code, "warning", message, reason, suggestion))
	}

	lower := strings.ToLower(strings.TrimSpace(text))
	if declRe.MatchString(text) || isProcedureHeaderLine(lower) {
		return state, findings
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		for _, clause := range splitArgs(match[2]) {
			clauseMatch := arrayRedimClauseRe.FindStringSubmatch(clause)
			if len(clauseMatch) == 0 {
				continue
			}
			name := strings.ToLower(clauseMatch[1])
			variable, known := variables[name]
			old := state[name]
			dimensions := parseArrayDimensions(clauseMatch[2], 0)
			if !known || !variable.isArray || variable.fixed {
				add("VBA227", clauseMatch[1]+" is not a dynamic array and cannot be resized with ReDim.", "ReDim requires a dynamic array; fixed-size arrays and scalar values have no resizable allocation state.", "Declare the value as a dynamic array, or remove ReDim and use its declared bounds.")
			} else if match[1] != "" && (len(old.dimensions) == 0 || !preserveDimensionsSafe(old.dimensions, dimensions)) {
				add("VBA208", "ReDim Preserve may change a non-final or unknown array dimension.", "VBA can only preserve an array while changing its final dimension, and that cannot be proven when the prior shape is unknown.", "Only change the final dimension, or copy values into a newly sized array explicitly.")
			}
			if known && variable.isArray && !variable.fixed {
				state[name] = arrayValue{kind: arrayAllocated, knownArray: true, dimensions: dimensions, origin: arrayOriginLocal}
			}
		}
		return state, findings
	}
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) > 0 {
		for _, target := range splitArgs(match[1]) {
			name := strings.ToLower(strings.TrimSpace(target))
			if !arrayEraseNameRe.MatchString(name) {
				continue
			}
			if variable, ok := variables[name]; ok {
				if variable.fixed {
					state[name] = arrayValue{kind: arrayAllocated, knownArray: true, dimensions: append([]arrayDimension(nil), variable.dimensions...), origin: arrayOriginLocal}
				} else if variable.isArray {
					state[name] = arrayValue{kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}
				} else if variable.isVariant {
					state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
				}
			}
		}
		return state, findings
	}
	// Inline `If ... Then ReDim ...` has a branch-specific state that is not
	// represented by one transfer result. Keep it conservative and avoid
	// treating the ReDim bounds themselves as an array access.
	if strings.Contains(lower, "redim ") {
		return state, findings
	}

	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		name := strings.ToLower(bound[2])
		value, ok := state[name]
		variable, known := variables[name]
		if !ok || !known || value.origin == arrayOriginRangeValue {
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			add("VBA227", bound[1]+" is used before "+variable.name+" is proven to be allocated.", "LBound and UBound raise a runtime error for an unallocated dynamic array and are unsafe for an unknown Variant.", "Allocate the array on every path before querying its bounds, or guard the operation explicitly.")
			continue
		}
		dimension := 1
		if argument := strings.TrimSpace(bound[3]); argument != "" {
			parsed, err := strconv.Atoi(argument)
			if err != nil {
				// A variable or expression is an unknown dimension, not an
				// invalid one. No contradiction can be proven statically.
				continue
			}
			dimension = parsed
		}
		if dimension < 1 || len(value.dimensions) > 0 && dimension > len(value.dimensions) {
			add("VBA227", bound[1]+" uses invalid dimension "+strconv.Itoa(dimension)+" for "+variable.name+".", "The requested dimension is outside the array dimensions known at this point.", "Use a valid dimension number for the array, or avoid assuming a shape that is not statically known.")
		}
	}

	for _, use := range arrayIndexedUses(text, variables) {
		value := state[strings.ToLower(use.name)]
		if value.origin == arrayOriginRangeValue {
			continue
		}
		if value.kind != arrayAllocated || !value.knownArray {
			add("VBA227", use.name+" is indexed before its array allocation is guaranteed.", "An array access can fail after Erase, before ReDim, or on a branch where allocation is not established.", "Allocate the array on every path before indexing it, or guard the access with a proven allocation check.")
			continue
		}
		if len(value.dimensions) > 0 && len(use.args) != len(value.dimensions) {
			add("VBA227", use.name+" is indexed with "+strconv.Itoa(len(use.args))+" dimension(s), but its known shape has "+strconv.Itoa(len(value.dimensions))+".", "The number of subscripts must match the array dimensions known to the analyzer.", "Use the correct number of subscripts or revise the declared array shape.")
			continue
		}
		for i, arg := range use.args {
			if i >= len(value.dimensions) {
				break
			}
			if literal, ok := integerLiteral(arg); ok {
				dimension := value.dimensions[i]
				if dimension.lower.known && literal < dimension.lower.value || dimension.upper.known && literal > dimension.upper.value {
					add("VBA227", use.name+" is indexed outside its known bounds.", "The subscript contradicts the lower or upper bound established by the declaration or ReDim.", "Use an index within the declared bounds, or establish the bounds dynamically before access.")
				}
			}
		}
	}

	if match := arrayForBoundRe.FindStringSubmatch(text); len(match) > 0 {
		name := strings.ToLower(match[2])
		if variable, ok := variables[name]; ok {
			value := state[name]
			if value.kind == arrayAllocated && value.knownArray && len(value.dimensions) > 0 && value.dimensions[0].lower.known {
				if start, ok := integerLiteral(match[1]); ok && start != value.dimensions[0].lower.value {
					add("VBA227", "The loop range assumes an inconsistent lower bound for "+variable.name+".", "The loop starts at a value different from the known lower bound of the array.", "Use LBound("+variable.name+") as the loop start.")
				}
			}
		}
	}

	if a.Config.Analyze.DetectObjectArrayComparison {
		names := make([]string, 0, len(variables))
		for name := range variables {
			names = append(names, name)
		}
		sort.Strings(names)
		for _, name := range names {
			variable := variables[name]
			if variable.isArray && identifierComparedAsOperand(strings.ToLower(text), name, proc) {
				add("VBA209", variable.name+" appears to be compared as a scalar value.", "VBA arrays cannot be compared directly to scalar values.", "Compare explicit elements or bounds instead of the array variable itself.")
			}
		}
	}

	if lhs, rhs, indexed, ok := arrayAssignment(text); ok {
		name := strings.ToLower(lhs)
		if variable, exists := variables[name]; exists && variable.isArray && variable.isObject && indexed && !strings.HasPrefix(lower, "set ") {
			code := "VBA101"
			typ := variable.typ
			callee := arrayCallName(rhs)
			if returnType := ctx.functionReturns[callee]; isObjectType(returnType) {
				code = "VBA102"
				typ = returnType
			}
			if code == "VBA101" || code == "VBA102" {
				// These are existing missing-Set rules; the lifecycle rule only
				// supplies the indexed target that the old text matcher missed.
				findings = append(findings, a.objectSetFinding(file, proc, line, code, strings.TrimSpace(lhs), typ))
			}
		}
		if !indexed {
			if value, known := arrayExpressionState(rhs, state, ctx); known {
				if variable, exists := variables[name]; exists && (variable.isArray || variable.isVariant) {
					state[name] = value
				}
			} else if variable, exists := variables[name]; exists && (variable.isArray || variable.isVariant) {
				state[name] = arrayValue{kind: arrayUnknown, knownArray: false, origin: arrayOriginUnknown}
			}
		}
	}
	return state, findings
}

func arrayVariables(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) map[string]arrayVariable {
	decls := cloneDeclarations(moduleDecls)
	for key, decl := range procedureDeclarations(file.Lines, proc) {
		decls[key] = decl
	}
	for _, param := range proc.Params {
		decls[strings.ToLower(param.Name)] = sourceDeclaration{Name: param.Name, Type: param.Type, Array: strings.Contains(param.Type, "()"), Object: isObjectType(param.Type), Parameter: true}
	}
	base := optionBase(file.Lines)
	variables := map[string]arrayVariable{}
	for key, decl := range decls {
		if !decl.Array && !strings.EqualFold(strings.TrimSpace(decl.Type), "Variant") {
			continue
		}
		variable := arrayVariable{name: decl.Name, typ: decl.Type, isArray: decl.Array, isVariant: strings.EqualFold(strings.TrimSpace(decl.Type), "Variant"), isObject: decl.Object, parameter: decl.Parameter}
		if decl.Array {
			variable.dimensions, variable.fixed = declarationDimensions(file.Lines, decl.Line, decl.Name, base)
			if len(variable.dimensions) > 0 {
				variable.fixed = true
			}
		}
		variables[key] = variable
	}
	return variables
}

func optionBase(lines []string) int {
	for _, line := range lines {
		fields := strings.Fields(strings.ToLower(normalizedCodeLine(line)))
		if len(fields) == 3 && fields[0] == "option" && fields[1] == "base" {
			if value, err := strconv.Atoi(fields[2]); err == nil && (value == 0 || value == 1) {
				return value
			}
		}
	}
	return 0
}

func declarationDimensions(lines []string, line int, name string, base int) ([]arrayDimension, bool) {
	if line < 1 || line > len(lines) {
		return nil, false
	}
	stmt := normalizedCodeLine(lines[line-1])
	m := declRe.FindStringSubmatch(stmt)
	if len(m) == 0 {
		return nil, false
	}
	for _, part := range splitArgs(m[1]) {
		candidate, _, array, _ := declarationNameAndType(part)
		if !array || !strings.EqualFold(candidate, name) {
			continue
		}
		start := strings.Index(part, "(")
		end := strings.LastIndex(part, ")")
		if start < 0 || end < start {
			return nil, false
		}
		raw := strings.TrimSpace(part[start+1 : end])
		if raw == "" {
			return nil, false
		}
		return parseArrayDimensions(raw, base), true
	}
	return nil, false
}

func parseArrayDimensions(text string, base int) []arrayDimension {
	var dimensions []arrayDimension
	for _, part := range splitArgs(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		pieces := strings.SplitN(strings.ToLower(part), " to ", 2)
		if len(pieces) == 2 {
			dimensions = append(dimensions, arrayDimension{lower: integerBound(pieces[0]), upper: integerBound(pieces[1])})
			continue
		}
		upper := integerBound(part)
		dimensions = append(dimensions, arrayDimension{lower: arrayBound{known: true, value: base}, upper: upper})
	}
	return dimensions
}

func integerBound(text string) arrayBound {
	value, ok := integerLiteral(text)
	return arrayBound{known: ok, value: value}
}

func integerLiteral(text string) (int, bool) {
	value, err := strconv.Atoi(strings.TrimSpace(text))
	return value, err == nil
}

func preserveDimensionsSafe(previous, next []arrayDimension) bool {
	if len(previous) != len(next) || len(previous) == 0 {
		return false
	}
	for i := 0; i < len(previous)-1; i++ {
		if !previous[i].lower.known || !previous[i].upper.known || !next[i].lower.known || !next[i].upper.known || previous[i] != next[i] {
			return false
		}
	}
	return true
}

func arrayIndexedUses(text string, variables map[string]arrayVariable) []arrayUse {
	var uses []arrayUse
	for i := 0; i < len(text); i++ {
		if !isIdentifierStart(text[i]) || (i > 0 && isIdentifierPart(text[i-1])) {
			continue
		}
		start := i
		i++
		for i < len(text) && isIdentifierPart(text[i]) {
			i++
		}
		name := text[start:i]
		key := strings.ToLower(name)
		if _, ok := variables[key]; !ok {
			continue
		}
		for i < len(text) && (text[i] == ' ' || text[i] == '\t') {
			i++
		}
		if i >= len(text) || text[i] != '(' {
			continue
		}
		end := matchingParen(text, i)
		if end < 0 {
			continue
		}
		uses = append(uses, arrayUse{name: name, args: splitArgs(text[i+1 : end])})
		i = end
	}
	return uses
}

func matchingParen(text string, start int) int {
	depth := 0
	inString := false
	for i := start; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString {
				depth--
				if depth == 0 {
					return i
				}
			}
		}
	}
	return -1
}

func isIdentifierStart(char byte) bool {
	return char == '_' || char >= 'A' && char <= 'Z' || char >= 'a' && char <= 'z'
}
func isIdentifierPart(char byte) bool { return isIdentifierStart(char) || char >= '0' && char <= '9' }

func arrayAssignment(text string) (lhs, rhs string, indexed, ok bool) {
	trimmed := strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(trimmed), "set ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	if strings.HasPrefix(strings.ToLower(trimmed), "let ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	for i := 0; i < len(trimmed); i++ {
		if trimmed[i] != '=' || i > 0 && (trimmed[i-1] == '<' || trimmed[i-1] == '>' || trimmed[i-1] == '=') {
			continue
		}
		lhs = strings.TrimSpace(trimmed[:i])
		rhs = strings.TrimSpace(trimmed[i+1:])
		if lhs == "" || strings.HasPrefix(strings.ToLower(lhs), "if ") {
			return "", "", false, false
		}
		if open := strings.Index(lhs, "("); open >= 0 && strings.HasSuffix(lhs, ")") {
			return cleanIdentifier(lhs[:open]), rhs, true, true
		}
		return cleanIdentifier(lhs), rhs, false, true
	}
	return "", "", false, false
}

func arrayExpressionState(rhs string, state arrayFlowState, ctx analysisContext) (arrayValue, bool) {
	lower := strings.ToLower(strings.TrimSpace(rhs))
	if strings.HasPrefix(lower, "array(") || lower == "array" {
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: []arrayDimension{{}}, origin: arrayOriginLocal}, true
	}
	if strings.Contains(lower, ".value") || strings.Contains(lower, ".value2") {
		return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginRangeValue}, true
	}
	name := arrayCallName(rhs)
	if name == "split" || name == "filter" {
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: []arrayDimension{{}}, origin: arrayOriginLocal}, true
	}
	if value, ok := state[name]; ok && value.kind == arrayAllocated && value.knownArray {
		return value, true
	}
	if value, ok := ctx.arrayReturns[name]; ok && value.kind == arrayAllocated && value.knownArray {
		return value, true
	}
	if name != "" && strings.Contains(rhs, "(") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	if strings.EqualFold(strings.TrimSpace(rhs), "empty") || strings.EqualFold(strings.TrimSpace(rhs), "nothing") {
		return arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}, true
	}
	return arrayValue{}, false
}

func arrayCallName(text string) string {
	name := strings.TrimSpace(lastName(strings.TrimSpace(text)))
	if open := strings.Index(name, "("); open >= 0 {
		name = name[:open]
	}
	return strings.ToLower(cleanIdentifier(name))
}

// inferArrayReturnSummaries intentionally summarizes only normal, directly
// observed return assignments.  A missing assignment, mixed assignment kinds,
// duplicate procedure names, and recursive/external calls remain unknown.
// This makes a project-wide batch pass useful without turning an unproved VBA
// function contract into an allocation guarantee.
func inferArrayReturnSummaries(files []parsedFile) map[string]arrayValue {
	type candidate struct {
		value arrayValue
		ok    bool
	}
	candidates := map[string][]candidate{}
	for _, file := range files {
		procedures := sourceProceduresFromIR(file.IR, file.CFG)
		moduleDecls := moduleDeclarations(file.Lines, procedures)
		for _, proc := range procedures {
			if proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet {
				continue
			}
			if proc.Name == "" {
				continue
			}
			variables := arrayVariables(file, proc, moduleDecls)
			ctx := analysisContext{arrayReturns: map[string]arrayValue{}}
			returnCandidates := map[int]candidate{}
			if proc.Graph == nil {
				// Without a CFG the scan cannot distinguish conditional from
				// unconditional return assignments, so leave the summary unknown.
				continue
			}
			walkArrayCFG(proc.Graph, file.Lines, arrayInitialState(variables), func(text string, line int, in arrayFlowState) arrayFlowState {
				if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed && strings.EqualFold(lhs, proc.Name) {
					value, known := arrayExpressionState(rhs, in, ctx)
					returnCandidates[line] = candidate{value: value, ok: known && value.kind == arrayAllocated && value.knownArray && value.origin != arrayOriginRangeValue}
				}
				out, _ := (Analyzer{}).arrayTransfer(file, proc, ctx, variables, in, text, line)
				return out
			})
			returnLines := make([]int, 0, len(returnCandidates))
			for line := range returnCandidates {
				returnLines = append(returnLines, line)
			}
			sort.Ints(returnLines)
			returns := make([]candidate, 0, len(returnLines))
			for _, line := range returnLines {
				returns = append(returns, returnCandidates[line])
			}
			if len(returns) == 0 {
				continue
			}
			if proc.Graph != nil && !proc.Graph.IsDefinitelyAssigned(proc.Graph.NormalExit, vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: proc.Name}, vbacfg.EdgeFilter{NormalOnly: true}) {
				continue
			}
			valid := returns[0].ok
			value := returns[0].value
			for _, returned := range returns[1:] {
				if !returned.ok || !arrayValueEqual(value, returned.value) {
					valid = false
					break
				}
			}
			candidates[strings.ToLower(proc.Name)] = append(candidates[strings.ToLower(proc.Name)], candidate{value: value, ok: valid})
		}
	}
	result := map[string]arrayValue{}
	for name, values := range candidates {
		if len(values) == 1 && values[0].ok {
			result[name] = values[0].value
		}
	}
	return result
}

func arrayValueEqual(left, right arrayValue) bool {
	return left.kind == right.kind && left.knownArray == right.knownArray && left.origin == right.origin && arrayDimensionsEqual(left.dimensions, right.dimensions)
}
