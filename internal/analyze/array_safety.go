package analyze

import (
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/lint"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
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
	known      bool
	value      int
	expression string
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
	name      string
	typ       string
	isArray   bool
	isVariant bool
	isObject  bool
	// knownScalar is intentionally narrower than "not an array".  Only
	// built-in scalar declarations are strong enough to prove that an array
	// operation or For Each source is invalid.  User-defined classes/UDTs and
	// unresolved external types remain unknown and therefore fail open.
	knownScalar bool
	fixed       bool
	parameter   bool
	dimensions  []arrayDimension
}

type arrayValue struct {
	kind          arrayAllocation
	knownArray    bool
	dimensions    []arrayDimension
	preserveShape []arrayDimension
	origin        arrayOrigin
}

type arrayFlowState map[string]arrayValue

type arrayUse struct {
	name string
	args []string
}

type directArrayRedimClause struct {
	name       string
	dimensions string
}

var (
	arrayRedimRe           = regexp.MustCompile(`(?i)^\s*redim\s+(preserve\s+)?(.+)$`)
	arrayRedimClauseRe     = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?)?\s*$`)
	arrayRedimTypeSuffixRe = regexp.MustCompile(`(?i)^as\s+[A-Za-z_][\w.]*(?:\s*\(\s*\))?$`)
	arrayEraseRe           = regexp.MustCompile(`(?i)^\s*erase\s+(.+)$`)
	arrayEraseNameRe       = regexp.MustCompile(`(?i)^[A-Za-z_]\w*$`)
	arrayBoundCallRe       = regexp.MustCompile(`(?i)\b(lbound|ubound)\s*\(\s*([^,)]*)\s*(?:,\s*([^)]*))?\)`)
	arrayForBoundRe        = regexp.MustCompile(`(?i)^\s*for\s+\w+\s*=\s*([-+]?\d+)\s+to\s+(?:lbound|ubound)\s*\(\s*([A-Za-z_]\w*)`)
	arrayForEachRe         = regexp.MustCompile(`(?i)^\s*for\s+each\s+[A-Za-z_]\w*\s+in\s+([^\r\n]+)`)
	arrayIndexedSourceRe   = regexp.MustCompile(`(?i)^\s*([A-Za-z_]\w*)\s*\(`)
)

func (a Analyzer) arrayLifecycleFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []Finding {
	// Array-shape enrichment is local to this rule. Do not mutate the shared
	// declaration map used by object, Excel, and dictionary analyses later in
	// the same module; doing so can change unrelated findings for subsequent
	// procedures.
	moduleDecls = cloneDeclarations(moduleDecls)
	// Keep module-level Const/Enum declarations visible to the shared shape
	// lattice; the historical text scanner only indexed Dim/Static/visibility
	// declarations.
	for _, declaration := range file.IR.Declarations {
		key := strings.ToLower(declaration.Name)
		if key == "" {
			continue
		}
		if _, exists := moduleDecls[key]; exists {
			continue
		}
		moduleDecls[key] = sourceDeclaration{
			Name: declaration.Name, Type: declaration.Type, Line: declaration.Range.StartLine,
			Object: declaration.IsObject, Array: declaration.IsArray,
		}
	}
	variables := arrayVariables(file, proc, moduleDecls)
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
	comparisonFindings := a.arrayComparisonFindings(file, proc, variables)
	comparisonFindings = append(comparisonFindings, a.arrayForEachFindings(file, proc, variables, ctx)...)
	initial := arrayInitialState(variables)
	// Constant bounds are scoped to the current procedure. Resolve them once
	// for this CFG walk so every ReDim transfer shares the same table without
	// repeatedly reparsing the module and procedure declarations.
	constants := arrayIntegerConstants(file, proc, a.visibleConstantValues, a.visibleConstants)
	if proc.Graph == nil {
		findings := append([]Finding(nil), comparisonFindings...)
		findings = append(findings, a.arrayLifecycleLinearFindings(file, proc, ctx, variables, initial, constants)...)
		return uniqueArrayFindings(findings)
	}

	findings := append([]Finding(nil), comparisonFindings...)
	seen := map[string]bool{}
	for _, finding := range findings {
		seen[arrayFindingKey(finding)] = true
	}
	walkArrayCFG(proc.Graph, file.Lines, initial, func(text string, line int, in arrayFlowState) arrayFlowState {
		out, issues := a.arrayTransfer(file, proc, ctx, variables, in, text, line, constants)
		for _, finding := range issues {
			key := arrayFindingKey(finding)
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

func (a Analyzer) arrayForEachFindings(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, ctx analysisContext) []Finding {
	if !a.Config.Analyze.DetectArrayLifecycleSafety {
		return nil
	}
	var findings []Finding
	for _, statement := range proc.Statements {
		if statement.Recovered || statement.Kind != procedureir.StatementForEach {
			continue
		}
		source := ""
		// The grammar has used both `value` and `collection` for the source
		// expression across parser versions.  Prefer the unambiguous statement
		// text when it matches the canonical For Each form, then fall back to
		// the IR value expression for recovered/alternate forms.
		if match := arrayForEachRe.FindStringSubmatch(statement.Text); len(match) > 0 {
			source = match[1]
		} else if statement.Value != nil {
			source = statement.Value.Text
		}
		source = strings.TrimSpace(strings.SplitN(source, "'", 2)[0])
		if !iterableSourceKnownInvalid(source, variables, arrayInitialState(variables), ctx) {
			continue
		}
		findings = append(findings, a.simpleFinding(file, proc, statement.Range.StartLine, "VBA227", "warning", strings.TrimSpace(source)+" is not a collection or array and cannot be used as a For Each source.", "For Each requires an iterable Collection or array value; this source is a known scalar.", "Iterate an array or Collection, or change the source expression to an iterable value."))
	}
	return findings
}

func uniqueArrayFindings(findings []Finding) []Finding {
	seen := map[string]bool{}
	out := make([]Finding, 0, len(findings))
	for _, finding := range findings {
		key := arrayFindingKey(finding)
		if seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, finding)
	}
	sortFindings(out)
	return out
}

func arrayFindingKey(finding Finding) string {
	return finding.Code + ":" + strconv.Itoa(finding.Line) + ":" + finding.Message
}

// arrayComparisonFindings inspects the shared expression IR instead of the
// CFG block text. A comparison's direct operand is the only position where
// an array is being used as a scalar; an array nested inside a call, member
// access, or indexed expression is a different semantic construct.
func (a Analyzer) arrayComparisonFindings(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable) []Finding {
	if !a.Config.Analyze.DetectObjectArrayComparison {
		return nil
	}
	expressions := make(map[int]procedureir.Expression, len(proc.Expressions))
	for _, expression := range proc.Expressions {
		expressions[expression.ID] = expression
	}
	statements := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statements[statement.ID] = statement
	}

	arrayNames := make([]string, 0, len(variables))
	nonArrayNames := procedureIRNonArrayNames(proc)
	for name, variable := range variables {
		if variable.isArray && !nonArrayNames[name] {
			arrayNames = append(arrayNames, name)
		}
	}
	sort.Strings(arrayNames)

	var findings []Finding
	for _, comparison := range proc.Expressions {
		if comparison.SyntaxKind != "comparison_expression" || comparison.Recovered {
			continue
		}
		statement, ok := statements[comparison.StatementID]
		if !ok || comparisonAssignmentCarrier(statement, comparison) {
			continue
		}
		matched := map[string]bool{}
		for _, childID := range comparison.Children {
			child, ok := directComparisonOperand(expressions, childID)
			if !ok || child.Recovered || child.Kind != procedureir.ExpressionIdentifier {
				continue
			}
			name := strings.ToLower(cleanIdentifier(child.Text))
			if variable, exists := variables[name]; exists && variable.isArray {
				matched[name] = true
			}
		}
		for _, name := range arrayNames {
			if !matched[name] {
				continue
			}
			variable := variables[name]
			findings = append(findings, a.simpleFinding(file, proc, comparison.Range.StartLine, "VBA209", "warning", variable.name+" appears to be compared as a scalar value.", "VBA arrays cannot be compared directly to scalar values.", "Compare explicit elements or bounds instead of the array variable itself."))
		}
	}
	sortFindings(findings)
	return findings
}

func procedureIRNonArrayNames(proc sourceProcedure) map[string]bool {
	names := map[string]bool{}
	for _, declaration := range proc.Declarations {
		if declaration.Recovered || declaration.Name == "" {
			continue
		}
		isArray := declaration.IsArray
		if declaration.Kind == "parameter" && strings.Contains(declaration.Type, "()") {
			isArray = true
		}
		if !isArray {
			names[strings.ToLower(declaration.Name)] = true
		}
	}
	return names
}

func directComparisonOperand(expressions map[int]procedureir.Expression, id int) (procedureir.Expression, bool) {
	child, ok := expressions[id]
	for ok && child.Kind == procedureir.ExpressionParentheses && len(child.Children) == 1 {
		child, ok = expressions[child.Children[0]]
	}
	return child, ok
}

func comparisonAssignmentCarrier(statement procedureir.Statement, expression procedureir.Expression) bool {
	if expression.ParentID != 0 {
		return false
	}
	return statement.Kind == procedureir.StatementAssignment || statement.Kind == procedureir.StatementSet
}

func (a Analyzer) arrayLifecycleLinearFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, constants map[string]int) []Finding {
	var findings []Finding
	seen := map[string]bool{}
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		var issues []Finding
		state, issues = a.arrayTransfer(file, proc, ctx, variables, state, normalizedCodeLine(file.Lines[line-1]), line, constants)
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
	out := arrayValue{
		kind:          left.kind,
		knownArray:    left.knownArray,
		origin:        left.origin,
		dimensions:    append([]arrayDimension(nil), left.dimensions...),
		preserveShape: append([]arrayDimension(nil), left.preserveShape...),
	}
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
	out.preserveShape = meetArrayDimensions(left.preserveShape, right.preserveShape)
	return out
}

func meetArrayDimensions(left, right []arrayDimension) []arrayDimension {
	if len(left) == 0 || len(left) != len(right) {
		return nil
	}
	out := make([]arrayDimension, len(left))
	for i := range left {
		out[i] = arrayDimension{
			lower: meetArrayBound(left[i].lower, right[i].lower),
			upper: meetArrayBound(left[i].upper, right[i].upper),
		}
	}
	return out
}

func meetArrayBound(left, right arrayBound) arrayBound {
	if arrayBoundsEquivalent(left, right) {
		return left
	}
	return arrayBound{}
}

func cloneArrayState(state arrayFlowState) arrayFlowState {
	out := arrayFlowState{}
	for name, value := range state {
		value.dimensions = append([]arrayDimension(nil), value.dimensions...)
		value.preserveShape = append([]arrayDimension(nil), value.preserveShape...)
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
		if !ok || l.kind != r.kind || l.knownArray != r.knownArray || l.origin != r.origin || !arrayDimensionsEqual(l.dimensions, r.dimensions) || !arrayDimensionsEqual(l.preserveShape, r.preserveShape) {
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
		if !variable.isArray {
			state[name] = arrayValue{kind: arrayUnknown, knownArray: false, origin: arrayOriginLocal}
			continue
		}
		value := arrayValue{
			knownArray:    variable.isArray,
			origin:        arrayOriginLocal,
			dimensions:    append([]arrayDimension(nil), variable.dimensions...),
			preserveShape: append([]arrayDimension(nil), variable.dimensions...),
		}
		if variable.fixed {
			value.kind = arrayAllocated
		} else if variable.parameter {
			value.kind = arrayUnknown
		} else {
			value.kind = arrayUnallocated
		}
		state[name] = value
	}
	return state
}

func (a Analyzer) arrayTransfer(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, text string, line int, constants map[string]int) (arrayFlowState, []Finding) {
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
		base := optionBase(file.Lines)
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if !direct {
				// Keep the existing receiver-array state transition for nested
				// member ReDim statements. VBA208 must not attribute the member
				// shape to that receiver, but changing the shared VBA227 state is
				// a separate analyzer contract.
				legacy := arrayRedimClauseRe.FindStringSubmatch(clause)
				if len(legacy) == 0 {
					continue
				}
				redim = directArrayRedimClause{name: legacy[1], dimensions: legacy[2]}
			}
			name := strings.ToLower(redim.name)
			variable, known := variables[name]
			// An unresolved target may be an implicit Variant or an external
			// member.  ReDim is valid for some of those runtime shapes, so only
			// report a target mismatch once the shared declaration facts prove a
			// scalar or fixed array.
			if !known {
				continue
			}
			old := state[name]
			dimensions := parseArrayDimensionsWithConstants(redim.dimensions, base, constants)
			// A Variant can be resized only after its array nature is proven by
			// an assignment such as Array(...).  An unproven Variant remains
			// unknown and must not be treated as a scalar ReDim misuse.
			variantArray := variable.isVariant && old.knownArray
			resizable := (variable.isArray || variantArray) && !variable.fixed
			if variable.isVariant && !old.knownArray {
				continue
			}
			if !resizable {
				// An unresolved object/UDT/external declaration is not a proven
				// scalar.  Leave it unknown rather than guessing that ReDim is
				// invalid; the shared shape contract is deliberately fail-open.
				if !variable.isArray && !variable.isVariant && !variable.knownScalar {
					continue
				}
				add("VBA227", redim.name+" is not a dynamic array and cannot be resized with ReDim.", "ReDim requires a dynamic array; fixed-size arrays and scalar values have no resizable allocation state.", "Declare the value as a dynamic array, or remove ReDim and use its declared bounds.")
			} else if impossibleArrayBounds(dimensions) {
				add("VBA227", redim.name+" has impossible constant ReDim bounds.", "A ReDim lower bound cannot be greater than its upper bound.", "Use bounds whose lower value is less than or equal to the upper value, or keep the bounds dynamic.")
			} else if direct && match[1] != "" && !preserveDimensionsSafe(old.preserveShape, dimensions) {
				add("VBA208", "ReDim Preserve may change a non-final or unknown array dimension.", "VBA can only preserve an array while changing its final dimension, and that cannot be proven when the prior shape is unknown.", "Only change the final dimension, or copy values into a newly sized array explicitly.")
			}
			if resizable && !impossibleArrayBounds(dimensions) {
				next := arrayValue{kind: arrayAllocated, knownArray: true, dimensions: dimensions, origin: arrayOriginLocal}
				if direct {
					next.preserveShape = dimensions
				} else {
					next.preserveShape = append([]arrayDimension(nil), old.preserveShape...)
				}
				state[name] = next
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
					state[name] = arrayValue{
						kind:          arrayAllocated,
						knownArray:    true,
						dimensions:    append([]arrayDimension(nil), variable.dimensions...),
						preserveShape: append([]arrayDimension(nil), variable.dimensions...),
						origin:        arrayOriginLocal,
					}
				} else if variable.isArray {
					state[name] = arrayValue{kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}
				} else if variable.isVariant {
					state[name] = arrayValue{kind: arrayUnknown, origin: arrayOriginUnknown}
				} else if variable.knownScalar || variable.isObject {
					add("VBA227", strings.TrimSpace(target)+" is not an array and cannot be erased as an array.", "Erase applies to arrays; a scalar value has no array allocation to clear.", "Erase an array variable or remove the Erase statement for this scalar value.")
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
		argument := strings.TrimSpace(bound[2])
		name := strings.ToLower(argument)
		value, ok := state[name]
		variable, known := variables[name]
		if !known {
			if scalarExpressionKnown(argument) {
				add("VBA227", bound[1]+" cannot be used on a known scalar expression.", "LBound and UBound require an array value; this argument is a statically known scalar.", "Pass an array value to the bound function or remove the bound query.")
			}
			continue
		}
		if !variable.isArray && !variable.isVariant {
			if !variable.knownScalar && !variable.isObject {
				continue
			}
			add("VBA227", bound[1]+" cannot be used on non-array "+variable.name+".", "LBound and UBound require an array value; this target is a known scalar.", "Pass an array variable to the bound function or remove the bound query.")
			continue
		}
		if !ok || value.origin == arrayOriginRangeValue {
			continue
		}
		// A Variant has no statically proven array nature.  Keep this path
		// fail-open; only a proven array (or a proven scalar handled above) is
		// actionable here.
		if variable.isVariant && !value.knownArray {
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

	if match := arrayForEachRe.FindStringSubmatch(text); len(match) > 0 {
		if iterableSourceKnownInvalid(match[1], variables, state, ctx) {
			add("VBA227", strings.TrimSpace(match[1])+" is not a collection or array and cannot be used as a For Each source.", "For Each requires an iterable Collection or array value; this source is a known scalar.", "Iterate an array or Collection, or change the source expression to an iterable value.")
		}
	}

	for _, use := range arrayIndexedUses(text, variables) {
		value := state[strings.ToLower(use.name)]
		if variable, ok := variables[strings.ToLower(use.name)]; ok && variable.isVariant && !value.knownArray {
			continue
		}
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
	base := optionBase(file.Lines)
	// The legacy line scanner intentionally ignores Const and Enum syntax.
	// Fill those gaps from the shared procedure IR so a known scalar constant
	// can still be classified as a non-iterable/non-array operation target.
	for _, declaration := range proc.Declarations {
		key := strings.ToLower(declaration.Name)
		if key == "" {
			continue
		}
		if _, exists := decls[key]; exists {
			continue
		}
		decls[key] = sourceDeclaration{
			Name: declaration.Name, Type: declaration.Type, Line: declaration.Range.StartLine,
			Object: declaration.IsObject, Array: declaration.IsArray,
			Fixed:      declaration.ValueShape == procedureir.ValueShapeFixedArray,
			Dimensions: parameterArrayDimensions(declaration.ArrayBounds, base),
		}
	}
	for _, param := range proc.Params {
		array := strings.Contains(param.Type, "()") || param.ValueShape == procedureir.ValueShapeFixedArray || param.ValueShape == procedureir.ValueShapeDynamicArray
		decls[strings.ToLower(param.Name)] = sourceDeclaration{
			Name: param.Name, Type: param.Type, Array: array,
			Fixed:      param.ValueShape == procedureir.ValueShapeFixedArray,
			Dimensions: parameterArrayDimensions(param.ArrayBounds, base),
			Object:     isObjectType(param.Type), Parameter: true,
		}
	}
	variables := map[string]arrayVariable{}
	for key, decl := range decls {
		typeName := strings.TrimSpace(decl.Type)
		// VBA declarations without an As clause are implicit Variant values.
		// Treat them exactly like an explicit Variant so array-sensitive rules
		// do not turn an unresolved value into a false scalar proof.
		isVariant := typeName == "" || strings.EqualFold(typeName, "Variant")
		variable := arrayVariable{name: decl.Name, typ: decl.Type, isArray: decl.Array, isVariant: isVariant, isObject: decl.Object, knownScalar: !isVariant && arrayKnownScalarType(typeName), parameter: decl.Parameter}
		if decl.Array {
			variable.dimensions, variable.fixed = declarationDimensions(file.Lines, decl.Line, decl.Name, base)
			if len(decl.Dimensions) > 0 {
				variable.dimensions = append([]arrayDimension(nil), decl.Dimensions...)
			}
			if decl.Fixed {
				variable.fixed = true
			}
			if len(variable.dimensions) > 0 {
				variable.fixed = true
			}
		}
		variables[key] = variable
	}
	return variables
}

func parameterArrayDimensions(bounds []procedureir.ArrayBound, base int) []arrayDimension {
	if len(bounds) == 0 {
		return nil
	}
	dimensions := make([]arrayDimension, 0, len(bounds))
	for _, bound := range bounds {
		lowerText, upperText := strings.TrimSpace(bound.Lower), strings.TrimSpace(bound.Upper)
		if lowerText == "" && upperText == "" {
			upperText = strings.TrimSpace(bound.Expression)
			lowerText = strconv.Itoa(base)
		}
		dimensions = append(dimensions, arrayDimension{
			lower: integerBound(lowerText), upper: integerBound(upperText),
		})
	}
	return dimensions
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
		dimensions = append(dimensions, arrayDimension{lower: integerBound(strconv.Itoa(base)), upper: upper})
	}
	return dimensions
}

func parseArrayDimensionsWithConstants(text string, base int, constants map[string]int) []arrayDimension {
	var dimensions []arrayDimension
	for _, part := range splitArgs(text) {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		pieces := strings.SplitN(strings.ToLower(part), " to ", 2)
		if len(pieces) == 2 {
			dimensions = append(dimensions, arrayDimension{lower: integerBoundWithConstants(pieces[0], constants), upper: integerBoundWithConstants(pieces[1], constants)})
			continue
		}
		dimensions = append(dimensions, arrayDimension{lower: integerBoundWithConstants(strconv.Itoa(base), constants), upper: integerBoundWithConstants(part, constants)})
	}
	return dimensions
}

// arrayIntegerConstants extends the shared range-value constant table with
// the Enum members that are valid integer bound names in VBA.  The parser is
// intentionally small and fail-open: unresolved expressions simply do not
// enter the table, so dynamic bounds remain unclassified.
func arrayIntegerConstants(file parsedFile, proc sourceProcedure, projectValues map[string]constexpr.Value, visibleNames map[string]bool) map[string]int {
	constants := rangeValueIntegerConstants(rangeValueModuleIntegerConstants(file.Lines, file.IR), proc)
	// Seed the ReDim evaluator from the shared typed projection so arithmetic,
	// Enum references, and qualified project constants follow the same rules as
	// declaration and Optional-default validation. Non-integral and unresolved
	// values intentionally stay out of this integer-only adapter.
	sharedValues := file.ConstantValues
	if sharedValues == nil {
		sharedValues = lint.ConstantValuesFromSource(string(file.Source), &file.IR, nil)
	}
	if len(projectValues) > 0 {
		merged := make(map[string]constexpr.Value, len(sharedValues)+len(projectValues))
		for name, value := range projectValues {
			key := strings.ToLower(cleanIdentifier(name))
			if visibleNames != nil && !visibleNames[key] {
				continue
			}
			merged[name] = value
		}
		for name, value := range sharedValues {
			merged[name] = value
		}
		sharedValues = merged
	}
	for name, value := range sharedValues {
		if value.Kind != constexpr.ValueInteger && value.Kind != constexpr.ValueLong && value.Kind != constexpr.ValueLongLong {
			continue
		}
		if int64(int(value.Integer)) != value.Integer {
			continue
		}
		constants[strings.ToLower(name)] = int(value.Integer)
	}
	inEnum := false
	var next *int
	conditionalDepth := 0
	for _, line := range file.Lines {
		code := strings.TrimSpace(normalizedCodeLine(line))
		lower := strings.ToLower(code)
		if strings.HasPrefix(lower, "#if ") {
			conditionalDepth++
			continue
		}
		if strings.HasPrefix(lower, "#end if") {
			if conditionalDepth > 0 {
				conditionalDepth--
			}
			continue
		}
		if conditionalDepth > 0 || strings.HasPrefix(lower, "#elseif ") || strings.HasPrefix(lower, "#else") {
			continue
		}
		if strings.HasPrefix(lower, "enum ") || strings.HasPrefix(lower, "public enum ") || strings.HasPrefix(lower, "private enum ") || strings.HasPrefix(lower, "friend enum ") {
			inEnum = true
			next = nil
			continue
		}
		if inEnum && strings.HasPrefix(lower, "end enum") {
			inEnum = false
			next = nil
			continue
		}
		if !inEnum || code == "" {
			continue
		}
		parts := strings.SplitN(code, "=", 2)
		fields := strings.Fields(parts[0])
		if len(fields) == 0 {
			continue
		}
		name := strings.TrimSpace(fields[0])
		if name == "" {
			continue
		}
		key := strings.ToLower(cleanIdentifier(name))
		if len(parts) == 2 {
			value, err := constantIntegerExpression(strings.TrimSpace(parts[1]), constants)
			if err != nil {
				next = nil
				continue
			}
			constants[key] = value
			n := value + 1
			next = &n
			continue
		}
		if next == nil {
			continue
		}
		constants[key] = *next
		n := *next + 1
		next = &n
	}
	return constants
}

func integerBound(text string) arrayBound {
	value, ok := integerLiteral(text)
	return arrayBound{known: ok, value: value, expression: canonicalArrayBoundExpression(text)}
}

func integerBoundWithConstants(text string, constants map[string]int) arrayBound {
	value, err := constantIntegerExpression(text, constants)
	if err != nil {
		return arrayBound{expression: canonicalArrayBoundExpression(text)}
	}
	return arrayBound{known: true, value: value, expression: canonicalArrayBoundExpression(text)}
}

func impossibleArrayBounds(dimensions []arrayDimension) bool {
	for _, dimension := range dimensions {
		if dimension.lower.known && dimension.upper.known && dimension.lower.value > dimension.upper.value {
			return true
		}
	}
	return false
}

func iterableSourceKnownInvalid(source string, variables map[string]arrayVariable, state arrayFlowState, ctx analysisContext) bool {
	source = strings.TrimSpace(source)
	if source == "" {
		return false
	}
	// A uniquely resolved scalar Function/Property Get return is as strong as
	// a scalar declaration.  Unknown, external, or duplicate call targets are
	// deliberately absent from functionShapes and therefore remain fail-open.
	if strings.Contains(source, "(") {
		callee := arrayCallName(source)
		if returnType, ok := ctx.functionReturns[callee]; ok && isObjectType(returnType) {
			// Known Collection/Dictionary/Object-compatible returns are valid
			// iterable sources even though their value shape is non-array.
			return false
		}
		if shape, ok := ctx.functionShapes[callee]; ok {
			switch shape {
			case procedureir.ValueShapeScalar:
				return true
			case procedureir.ValueShapeFixedArray, procedureir.ValueShapeDynamicArray:
				return false
			}
		}
	}
	// An indexed element of a statically typed scalar array is itself a
	// scalar, even though the base expression is an allocated array. Variant
	// and object arrays remain conservative because their elements may hold an
	// array or implement the Collection iteration contract.
	if match := arrayIndexedSourceRe.FindStringSubmatch(source); len(match) > 0 {
		if variable, ok := variables[strings.ToLower(match[1])]; ok && variable.isArray && !variable.isVariant && !variable.isObject && variable.knownScalar {
			return true
		}
	}
	if value, known := arrayExpressionState(source, state, ctx); known {
		if value.knownArray {
			return false
		}
		// Unknown calls and Variant values remain conservative.  A known scalar
		// expression is handled below where its declaration is available.
		if strings.Contains(source, "(") {
			return false
		}
	}
	if scalarExpressionKnown(source) {
		return true
	}
	name := source
	if dot := strings.IndexAny(name, ".("); dot >= 0 {
		name = name[:dot]
	}
	name = strings.TrimSpace(name)
	if !arrayEraseNameRe.MatchString(name) {
		// Numeric, Boolean, date, and string literals are definitely scalar; all other
		// expressions remain unknown rather than guessing their result type.
		return false
	}
	variable, ok := variables[strings.ToLower(name)]
	if !ok || variable.isArray || variable.isVariant || variable.isObject || !variable.knownScalar || dcKindFromType(variable.typ) != dcUnknown {
		return false
	}
	return true
}

func arrayKnownScalarType(typ string) bool {
	switch strings.ToLower(cleanIdentifier(strings.TrimSpace(typ))) {
	case "byte", "boolean", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal", "date", "string":
		return true
	default:
		return false
	}
}

func scalarExpressionKnown(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" {
		return false
	}
	lower := strings.ToLower(text)
	switch lower {
	case "true", "false", "nothing", "empty", "null":
		return true
	}
	if strings.HasPrefix(text, `"`) || (strings.HasPrefix(text, "#") && strings.HasSuffix(text, "#")) {
		return true
	}
	if _, ok := integerLiteral(text); ok {
		return true
	}
	if _, err := strconv.ParseFloat(strings.TrimRight(text, "%&!@$^"), 64); err == nil {
		return true
	}
	_, err := constantIntegerExpression(text, nil)
	return err == nil
}

func integerLiteral(text string) (int, bool) {
	value, err := strconv.Atoi(strings.TrimSpace(text))
	return value, err == nil
}

func preserveDimensionsSafe(previous, next []arrayDimension) bool {
	if len(next) == 1 {
		return len(previous) <= 1
	}
	if len(previous) != len(next) || len(previous) == 0 {
		return false
	}
	for i := 0; i < len(previous)-1; i++ {
		if !arrayBoundsEquivalent(previous[i].lower, next[i].lower) || !arrayBoundsEquivalent(previous[i].upper, next[i].upper) {
			return false
		}
	}
	return true
}

func arrayBoundsEquivalent(left, right arrayBound) bool {
	if left.known && right.known {
		return left.value == right.value
	}
	return left.expression != "" && left.expression == right.expression
}

func canonicalArrayBoundExpression(text string) string {
	var out strings.Builder
	inString := false
	for i := 0; i < len(text); i++ {
		ch := text[i]
		if ch == '"' {
			out.WriteByte(ch)
			if inString && i+1 < len(text) && text[i+1] == '"' {
				out.WriteByte(text[i+1])
				i++
				continue
			}
			inString = !inString
			continue
		}
		if !inString {
			if ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' {
				continue
			}
			if ch >= 'A' && ch <= 'Z' {
				ch += 'a' - 'A'
			}
		}
		out.WriteByte(ch)
	}
	return out.String()
}

func parseDirectArrayRedimClause(clause string) (directArrayRedimClause, bool) {
	clause = strings.TrimSpace(clause)
	if clause == "" || !isIdentifierStart(clause[0]) {
		return directArrayRedimClause{}, false
	}
	endName := 1
	for endName < len(clause) && isIdentifierPart(clause[endName]) {
		endName++
	}
	open := endName
	for open < len(clause) && (clause[open] == ' ' || clause[open] == '\t') {
		open++
	}
	if open >= len(clause) || clause[open] != '(' {
		return directArrayRedimClause{}, false
	}
	close := matchingParen(clause, open)
	if close < 0 {
		return directArrayRedimClause{}, false
	}
	remainder := strings.TrimSpace(clause[close+1:])
	if remainder != "" && !arrayRedimTypeSuffixRe.MatchString(remainder) {
		return directArrayRedimClause{}, false
	}
	dimensions := strings.TrimSpace(clause[open+1 : close])
	if dimensions == "" {
		return directArrayRedimClause{}, false
	}
	return directArrayRedimClause{name: clause[:endName], dimensions: dimensions}, true
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
		variable, ok := variables[key]
		if !ok || !variable.isArray && !variable.isVariant {
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
		shape := []arrayDimension{{}}
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: shape, preserveShape: shape, origin: arrayOriginLocal}, true
	}
	if strings.Contains(lower, ".value") || strings.Contains(lower, ".value2") {
		return arrayValue{kind: arrayAllocated, knownArray: true, origin: arrayOriginRangeValue}, true
	}
	name := arrayCallName(rhs)
	if name == "split" || name == "filter" {
		shape := []arrayDimension{{}}
		return arrayValue{kind: arrayAllocated, knownArray: true, dimensions: shape, preserveShape: shape, origin: arrayOriginLocal}, true
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
			constants := arrayIntegerConstants(file, proc, nil, nil)
			walkArrayCFG(proc.Graph, file.Lines, arrayInitialState(variables), func(text string, line int, in arrayFlowState) arrayFlowState {
				if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed && strings.EqualFold(lhs, proc.Name) {
					value, known := arrayExpressionState(rhs, in, ctx)
					returnCandidates[line] = candidate{value: value, ok: known && value.kind == arrayAllocated && value.knownArray && value.origin != arrayOriginRangeValue}
				}
				out, _ := (Analyzer{}).arrayTransfer(file, proc, ctx, variables, in, text, line, constants)
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
	return left.kind == right.kind && left.knownArray == right.knownArray && left.origin == right.origin && arrayDimensionsEqual(left.dimensions, right.dimensions) && arrayDimensionsEqual(left.preserveShape, right.preserveShape)
}
