package analyze

import (
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// variableAssignmentFindings projects procedure-local definite-assignment
// facts into VBA263 (read but never assigned) and VBA264 (read before a
// guaranteed assignment) findings. Both rules share one candidate pass so the
// conservative eligibility and fail-open decisions stay identical.
func (a Analyzer) variableAssignmentFindings(file parsedFile, proc sourceProcedure, signatures map[string]procedureSignature) []Finding {
	result := assignmentFacts(proc, signatures)
	if result == nil {
		return nil
	}
	var findings []Finding
	if a.Config.Analyze.DetectNeverAssignedVariables {
		for _, candidate := range result.neverAssigned {
			finding := a.simpleFinding(file, proc, candidate.decl.StartLine, "VBA263", "warning",
				"Local variable "+candidate.displayName+" is read but never assigned.",
				"No statement, ReDim, or resolvable ByRef call assigns a value to this variable, so every read observes the implicit default.",
				"Initialize the variable before use, or remove the declaration when the read was unintended.")
			finding.Column = candidate.decl.StartColumn
			finding.EndLine = candidate.decl.EndLine
			finding.EndColumn = candidate.decl.EndColumn
			findings = append(findings, finding)
		}
	}
	if a.Config.Analyze.DetectUnassignedVariableUsage {
		for _, candidate := range result.unassignedReads {
			finding := a.simpleFinding(file, proc, candidate.read.StartLine, "VBA264", "warning",
				"Variable "+candidate.displayName+" is read before any assignment is guaranteed to have executed.",
				"A reachable control-flow path enters this statement without a completed assignment to "+candidate.displayName+".",
				"Assign the variable on every path that reaches this read, or check the earlier assignment condition.")
			finding.Column = candidate.read.StartColumn
			finding.EndLine = candidate.read.EndLine
			finding.EndColumn = candidate.read.EndColumn
			findings = append(findings, finding)
		}
	}
	return findings
}

type assignmentNeverAssignedCandidate struct {
	name        string
	displayName string
	decl        vbaast.Range
}

type assignmentUnassignedReadCandidate struct {
	name        string
	displayName string
	read        vbaast.Range
}

type assignmentFactsResult struct {
	neverAssigned   []assignmentNeverAssignedCandidate
	unassignedReads []assignmentUnassignedReadCandidate
}

// assignmentFacts computes definite-assignment evidence for plain scalar local
// variables. It shares the dead-store eligibility and fail-open posture: only
// fully modeled procedures produce findings, and any ambiguous access removes
// the affected variable instead of the whole procedure where the variable
// identity is still provable.
func assignmentFacts(proc sourceProcedure, signatures map[string]procedureSignature) *assignmentFactsResult {
	if proc.Graph == nil || proc.Graph.Entry == 0 || proc.IR == nil || proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 {
		return nil
	}

	statementsByID := make(map[int]procedureir.Statement, proc.Statements.Len())
	siblingsByParent := make(map[int][]procedureir.Statement)
	for statement := range proc.Statements.All() {
		statementsByID[statement.ID] = statement
		siblingsByParent[statement.ParentID] = append(siblingsByParent[statement.ParentID], statement)
	}
	for _, siblings := range siblingsByParent {
		sort.Slice(siblings, func(i, j int) bool {
			return siblings[i].Range.StartByte < siblings[j].Range.StartByte
		})
		for i, statement := range siblings {
			if statement.SyntaxKind != "single_line_if_statement" || i+1 >= len(siblings) {
				continue
			}
			if siblings[i+1].Range.StartLine == statement.Range.EndLine {
				return nil
			}
		}
	}

	eligible := assignmentEligibleDeclarations(proc.Declarations)
	if len(eligible) == 0 {
		return nil
	}

	// VBA264 measures definite assignment over normal flow only. Exceptional
	// edges (On Error Resume Next resume paths, error-handler jumps) model
	// writes that may be interrupted, and treating them as ordinary
	// predecessors would empty the must-assignment sets for any procedure
	// that uses error handling, flooding the rule with false positives.
	view := proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true})
	reachable := make(map[vbacfg.BlockID]bool)
	for _, id := range view.Reachable() {
		reachable[id] = true
	}
	if len(reachable) == 0 {
		return nil
	}

	// Unknown flow can transfer control to an unmodeled statement, so neither
	// "never assigned" nor "read before assignment" can be proven. Fail open
	// for the whole procedure, matching the dead-store contract.
	for _, source := range proc.Graph.UnknownFlowSources {
		if reachable[source] {
			return nil
		}
	}
	for _, edge := range proc.Graph.Edges {
		if !reachable[edge.From] || edge.Class != vbacfg.EdgeNormal {
			continue
		}
		if edge.Uncertain || edge.Kind == vbacfg.EdgeUnknown {
			return nil
		}
	}

	// byRefWriteExpr records argument expressions that a callee may write
	// through a ByRef (or unresolved) parameter. Those accesses are potential
	// definitions, not ordinary reads.
	byRefWriteExpr := make(map[int]bool)
	accessesByStatement := make(map[int][]procedureir.VariableAccess)
	for access := range proc.Accesses.All() {
		accessesByStatement[access.StatementID] = append(accessesByStatement[access.StatementID], access)
	}
	for call := range proc.Calls.All() {
		writeIDs := callWritableArgumentExpressions(call, signatures)
		for _, access := range accessesByStatement[call.StatementID] {
			if writeIDs[access.ExpressionID] {
				byRefWriteExpr[access.ExpressionID] = true
			}
		}
	}

	type candidate struct {
		decl        procedureir.Declaration
		displayName string
	}
	candidates := make(map[string]candidate, len(eligible))
	declRanges := make(map[string]vbaast.Range, len(eligible))
	for declaration := range proc.Declarations.All() {
		name := assignmentCanonicalName(declaration.Name)
		if !eligible[name] {
			continue
		}
		candidates[name] = candidate{decl: declaration, displayName: cleanIdentifier(declaration.Name)}
		declRanges[name] = declaration.Range
	}

	hasWrite := make(map[string]bool, len(candidates))
	hasRead := make(map[string]bool, len(candidates))
	readsByStatement := make(map[int][]procedureir.VariableAccess)
	for access := range proc.Accesses.All() {
		name := assignmentCanonicalName(access.Name)
		if name == "" {
			continue
		}
		if _, ok := candidates[name]; !ok {
			continue
		}
		statement, ok := statementsByID[access.StatementID]
		if !ok {
			continue
		}
		if access.Scope != procedureir.ScopeLocal || len(statement.ConditionalBranches) > 0 {
			// An unresolved or shadowing same-name access cannot be attributed
			// to this local; conditional-compiled accesses are unmodeled.
			delete(candidates, name)
			continue
		}
		if assignmentAmbiguousStatement(statement) {
			delete(candidates, name)
			continue
		}
		if target, shaped := assignmentAmbiguousTarget(statement); shaped && (target == "" || target == name) {
			// LSet/RSet/Mid$ statements write a target the IR records only as
			// a read. Only the target variable is ambiguous; the remaining
			// accesses on the statement are ordinary reads.
			delete(candidates, name)
			continue
		}
		if nextVariableAccess(statement, access) {
			// `Next x` is loop bookkeeping, not a value read; the loop header
			// itself assigns the control variable.
			continue
		}
		if byRefWriteExpr[access.ExpressionID] {
			// A call may or may not write through the ByRef argument. Count it
			// as a potential definition so the variable is never provably
			// "never assigned", but do not treat the argument position itself
			// as a proven read: initializer calls only write their target.
			// Reads on statements before the call still feed VBA264.
			hasWrite[name] = true
			continue
		}
		switch access.Mode {
		case procedureir.AccessWrite, procedureir.AccessReadWrite:
			hasWrite[name] = true
		default:
			hasRead[name] = true
			readsByStatement[access.StatementID] = append(readsByStatement[access.StatementID], access)
		}
	}

	// Augment the CFG's per-statement assignment lists with potential ByRef
	// call writes so the definite-assignment fixpoint accounts for them. The
	// graph is shallow-copied; block slices are replaced before the view is
	// built, leaving the original untouched.
	byRefDefs := make(map[int][]vbacfg.Variable)
	for call := range proc.Calls.All() {
		writeIDs := callWritableArgumentExpressions(call, signatures)
		if len(writeIDs) == 0 {
			continue
		}
		for _, access := range accessesByStatement[call.StatementID] {
			name := assignmentCanonicalName(access.Name)
			if !writeIDs[access.ExpressionID] || name == "" {
				continue
			}
			byRefDefs[call.StatementID] = append(byRefDefs[call.StatementID], vbacfg.Variable{Scope: procedureir.ScopeLocal, Name: name})
		}
	}

	result := &assignmentFactsResult{}
	for name, candidate := range candidates {
		if hasRead[name] && !hasWrite[name] {
			result.neverAssigned = append(result.neverAssigned, assignmentNeverAssignedCandidate{
				name: name, displayName: candidate.displayName, decl: declRanges[name],
			})
		}
	}
	sort.Slice(result.neverAssigned, func(i, j int) bool {
		return result.neverAssigned[i].decl.StartByte < result.neverAssigned[j].decl.StartByte
	})

	if len(readsByStatement) == 0 {
		return result
	}
	augmented := *proc.Graph
	if len(byRefDefs) > 0 {
		blocks := make([]vbacfg.Block, len(augmented.Blocks))
		copy(blocks, augmented.Blocks)
		for i := range blocks {
			defs := byRefDefs[blocks[i].StatementID]
			if len(defs) == 0 {
				continue
			}
			merged := append([]vbacfg.Variable{}, blocks[i].Assignments...)
			merged = append(merged, defs...)
			blocks[i].Assignments = merged
		}
		augmented.Blocks = blocks
	}
	definite := augmented.View(vbacfg.EdgeFilter{NormalOnly: true}).DefiniteAssignments()
	definiteAt := make(map[vbacfg.BlockID]map[string]bool, len(definite))
	for blockID, variables := range definite {
		set := make(map[string]bool, len(variables))
		for _, variable := range variables {
			set[variable.Name] = true
		}
		definiteAt[blockID] = set
	}

	reported := make(map[string]bool)
	blocksByStatement := make(map[int]vbacfg.BlockID)
	for _, block := range augmented.Blocks {
		if block.Kind == vbacfg.BlockStatement && reachable[block.ID] {
			blocksByStatement[block.StatementID] = block.ID
		}
	}
	statementIDs := make([]int, 0, len(readsByStatement))
	for statementID := range readsByStatement {
		statementIDs = append(statementIDs, statementID)
	}
	sort.Ints(statementIDs)
	for _, statementID := range statementIDs {
		blockID, ok := blocksByStatement[statementID]
		if !ok {
			continue
		}
		// Blocks absent from definiteAt have an empty definite set; a nil map
		// read is safe and treats every candidate as unassigned at entry.
		entry := definiteAt[blockID]
		reads := append([]procedureir.VariableAccess{}, readsByStatement[statementID]...)
		sort.Slice(reads, func(i, j int) bool {
			return reads[i].Range.StartByte < reads[j].Range.StartByte
		})
		for _, access := range reads {
			name := assignmentCanonicalName(access.Name)
			candidate, ok := candidates[name]
			if !ok || reported[name] || entry[name] {
				continue
			}
			reported[name] = true
			result.unassignedReads = append(result.unassignedReads, assignmentUnassignedReadCandidate{
				name: name, displayName: candidate.displayName, read: access.Range,
			})
		}
	}
	sort.Slice(result.unassignedReads, func(i, j int) bool {
		return result.unassignedReads[i].read.StartByte < result.unassignedReads[j].read.StartByte
	})
	return result
}

// assignmentEligibleDeclarations restricts assignment analysis to plain
// scalar locals whose storage and initialization semantics are fully known.
// It is deliberately narrower than dead-store eligibility: a user-defined
// type variable is initialized by its declaration, and member reads must not
// be mistaken for whole-variable reads.
func assignmentEligibleDeclarations(declarations readOnlySpan[procedureir.Declaration]) map[string]bool {
	eligible := make(map[string]bool)
	for declaration := range declarations.All() {
		if declaration.Scope != procedureir.ScopeLocal || declaration.Kind == "return_slot" || declaration.IsStatic || declaration.IsArray ||
			declaration.IsObject || declaration.IsNew || declaration.IsConst || len(declaration.ConditionalBranches) > 0 ||
			!isScalarAssignmentType(declaration.Type) {
			continue
		}
		name := assignmentCanonicalName(declaration.Name)
		if name != "" {
			eligible[name] = true
		}
	}
	return eligible
}

func isScalarAssignmentType(typeName string) bool {
	switch strings.ToLower(strings.TrimSpace(typeName)) {
	case "byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal", "date", "boolean", "string":
		return true
	default:
		return false
	}
}

// assignmentAmbiguousStatement reports statements whose assignment semantics
// the IR does not model precisely enough to classify reads and writes. Input
// #/Line Input #/Get # parse as unknown statements but assign their targets.
// LSet/RSet/Mid$ statements are handled per variable by
// assignmentAmbiguousTarget instead.
func assignmentAmbiguousStatement(statement procedureir.Statement) bool {
	return statement.Kind == procedureir.StatementUnknown || statement.Kind == procedureir.StatementRecovered || statement.Recovered
}

// assignmentAmbiguousTarget isolates the write target of an LSet/RSet/Mid$
// assignment statement. These shapes parse as calls or assignments whose
// target is recorded only as a read, so the target variable cannot be
// classified. shaped reports whether the statement follows one of those
// forms; an empty target with shaped=true means the form was recognized but
// the target could not be isolated, so every candidate touched by the
// statement is ambiguous.
func assignmentAmbiguousTarget(statement procedureir.Statement) (target string, shaped bool) {
	if statement.Kind != procedureir.StatementCall && statement.Kind != procedureir.StatementAssignment {
		return "", false
	}
	text := strings.TrimSpace(statement.Text)
	keyword := strings.ToLower(text)
	// The keyword may attach directly to the target's argument list
	// (`Mid$(target, ...) = rhs`), so it ends at the first `(` or whitespace
	// rather than at a whitespace boundary alone.
	if index := strings.IndexAny(keyword, " (\t"); index >= 0 {
		keyword = keyword[:index]
	}
	rest := strings.TrimSpace(text[len(keyword):])
	switch strings.TrimRight(keyword, "$") {
	case "lset", "rset":
		return assignmentCanonicalName(leadingIdentifier(rest)), true
	case "mid":
		if !strings.HasPrefix(rest, "(") {
			return "", true
		}
		return assignmentCanonicalName(leadingIdentifier(strings.TrimSpace(rest[1:]))), true
	default:
		return "", false
	}
}

// leadingIdentifier returns the identifier at the start of text, stopping at
// the first character that cannot appear inside a VBA identifier. A bracketed
// or qualified target yields an empty result, which callers treat as an
// unresolvable ambiguous statement.
func leadingIdentifier(text string) string {
	end := 0
	for end < len(text) && isIdentifierByte(text[end]) {
		end++
	}
	return text[:end]
}

// nextVariableAccess reports whether an access is a For/For Each statement's
// own Next-variable occurrence (`Next x`). That position is loop syntax, not
// a value read; the loop header assigns the control variable, so the access
// must be excluded from both read tracking and unassigned-read reporting.
func nextVariableAccess(statement procedureir.Statement, access procedureir.VariableAccess) bool {
	control := statement.Control
	if control == nil {
		return false
	}
	for _, rng := range control.NextVariableRanges {
		if rng == access.Range {
			return true
		}
	}
	return false
}

// callWritableArgumentExpressions returns the argument expression IDs a call
// may write through. Only a matched call whose candidates all resolve to
// known signatures can exclude ByVal positions; every other status is
// treated as able to write any argument.
func callWritableArgumentExpressions(call procedureir.CallSite, signatures map[string]procedureSignature) map[int]bool {
	all := func() map[int]bool {
		ids := make(map[int]bool, len(call.Arguments.ExpressionIDs))
		for _, id := range call.Arguments.ExpressionIDs {
			ids[id] = true
		}
		return ids
	}
	if len(call.Arguments.ExpressionIDs) == 0 || call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) == 0 {
		return all()
	}
	writable := make(map[int]bool)
	named := make(map[int]string, len(call.Arguments.Named))
	for _, argument := range call.Arguments.Named {
		named[argument.ExpressionID] = argument.Name
	}
	complete := true
	for _, candidate := range call.Resolution.Candidates {
		signature, ok := signatures[strings.ToLower(strings.TrimSpace(candidate.QualifiedName))]
		if !ok {
			complete = false
			break
		}
		positional := 0
		for _, exprID := range call.Arguments.ExpressionIDs {
			if name, isNamed := named[exprID]; isNamed {
				if callParameterWritableByName(signature.Params, name) {
					writable[exprID] = true
				}
				continue
			}
			if callParameterWritableAt(signature.Params, positional) {
				writable[exprID] = true
			}
			positional++
		}
	}
	if !complete {
		return all()
	}
	return writable
}

func callParameterWritableAt(parameters readOnlySpan[procedureir.Parameter], index int) bool {
	for i := 0; i < parameters.Len(); i++ {
		parameter, ok := parameters.At(i)
		if !ok {
			return true
		}
		if i != index {
			continue
		}
		return !parameter.ParamArray && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "byval")
	}
	return true
}

func callParameterWritableByName(parameters readOnlySpan[procedureir.Parameter], name string) bool {
	key := assignmentCanonicalName(name)
	for parameter := range parameters.All() {
		if assignmentCanonicalName(parameter.Name) == key {
			return !parameter.ParamArray && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "byval")
		}
	}
	return true
}

func assignmentCanonicalName(name string) string {
	return strings.ToLower(strings.TrimSpace(cleanIdentifier(name)))
}
