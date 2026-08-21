package analyze

import (
	"fmt"
	"regexp"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var errorSuccessIdentifierRE = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]*$`)

// errorSuppressionFindings owns the interprocedural failure-information
// contract. Procedure summaries retain every origin, but findings are emitted
// only by the procedure that owns the loss site so transitive callers do not
// repeat the same root cause.
func (a Analyzer) errorSuppressionFindings(file parsedFile, proc sourceProcedure, project effects.ProjectSummary) []Finding {
	if !a.Config.Analyze.DetectErrorSuppressionPropagation || proc.Name == "" || proc.Effects == nil {
		return nil
	}
	var findings []Finding
	seen := map[string]bool{}
	procedureIdentities := project.Identities()
	resumeNextOwnedByVBA214 := map[int]bool{}
	if a.Config.Analyze.DetectLeakedOnErrorResumeNextScopes {
		for _, existing := range a.leakedOnErrorResumeNextFindings(file, proc) {
			resumeNextOwnedByVBA214[existing.Line] = true
		}
	}
	for _, evidence := range proc.Effects.Error.Direct {
		if evidence.Behavior != effects.ErrorSuppresses || evidence.Origin.Key() != proc.Effects.Identity.Key() {
			continue
		}
		line := evidence.Range.StartLine
		if line <= 0 {
			line = proc.StartLine
		}
		if evidence.Target == "resume_next" && resumeNextOwnedByVBA214[line] {
			continue
		}
		key := fmt.Sprintf("suppressed:%d:%d", line, evidence.StatementID)
		if seen[key] {
			continue
		}
		seen[key] = true
		message := proc.Name + " can catch a runtime error and return normally without signaling failure."
		reason := "An exceptional control-flow path reaches a normal procedure exit without an explicit rethrow, recovery, fallback return, or Boolean failure result."
		if errorEvidenceAt(proc.Effects.Error.Direct, effects.ErrorLogsAndContinues, evidence) {
			message = proc.Name + " logs a runtime error and then continues without signaling failure."
			reason = "The handler records error information, but the same exceptional path reaches a normal procedure exit without rethrowing or returning failure."
		}
		if chain := representativePublicErrorChain(procedureIdentities, project, evidence); chain != "" {
			reason += " Representative call chain: " + chain + "."
		}
		findings = append(findings, a.simpleFinding(
			file, proc, line, "VBA237", "warning", message, reason,
			"Rethrow the saved error after cleanup, return an explicit failure result, or resume only after the failure has been handled.",
		))
	}

	statements := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statements[statement.ID] = statement
	}
	for _, call := range proc.Calls {
		if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 || !applicationStateCallReachable(proc, call) {
			continue
		}
		callee, ok := project.LookupCandidateDirect(call.Resolution.Candidates[0])
		statement := statements[call.StatementID]
		if !ok || !directSuccessFlag(callee) || errorFailureOutputObserved(proc, call, callee) || errorSuccessResultUseUncertain(proc, call, statement) || errorSuccessResultChecked(proc, call, statement) {
			continue
		}
		line := call.Range.StartLine
		if line <= 0 {
			line = proc.StartLine
		}
		key := fmt.Sprintf("ignored:%d:%d", line, call.ID)
		if seen[key] {
			continue
		}
		seen[key] = true
		calleeName := callee.Identity.QualifiedName
		if calleeName == "" {
			calleeName = call.Callee.Text
		}
		reason := calleeName + " reports handled failure through its Boolean result, but " + proc.Name + " neither checks nor propagates that result."
		if isProjectVisibleProcedure(proc.Effects.Identity) {
			reason += " This call is itself on a public procedure boundary."
		}
		findings = append(findings, a.simpleFinding(
			file, proc, line, "VBA237", "warning",
			proc.Name+" ignores the Boolean success result from "+calleeName+".",
			reason,
			"Check the Boolean result in a branch, or return it from the caller so failure remains observable.",
		))
	}
	return findings
}

func errorFailureOutputObserved(proc sourceProcedure, call procedureir.CallSite, callee effects.ProcedureSummary) bool {
	for _, evidence := range callee.Error.Direct {
		if evidence.Behavior != effects.ErrorReturnsSuccess || evidence.Origin.Key() != callee.Identity.Key() {
			continue
		}
		for _, output := range evidence.FailureOutputs {
			expression, ok := callerArgumentExpression(proc, call, output)
			if !ok {
				continue
			}
			name := strings.TrimSpace(expression)
			if !errorSuccessIdentifierRE.MatchString(name) {
				continue
			}
			if strings.EqualFold(name, proc.Name) &&
				(proc.ProcedureKind == procedureir.ProcedureFunction || proc.ProcedureKind == procedureir.ProcedurePropertyGet) {
				return true
			}
			if procedureLocal(proc, name) && resultMayBeCheckedAfterStatement(proc, call.StatementID, name) {
				return true
			}
		}
	}
	return false
}

// resultMayBeCheckedAfterStatement is intentionally weaker than the Boolean
// success-result contract. A ByRef sentinel can be irrelevant on a path where
// an earlier sentinel already terminates the operation; it is sufficient that
// the output reaches a control decision on every path where it remains live.
func resultMayBeCheckedAfterStatement(proc sourceProcedure, statementID int, target string) bool {
	if proc.Graph == nil {
		return false
	}
	start, ok := proc.Graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	byID := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		byID[statement.ID] = statement
	}
	queue := []vbacfg.BlockID{}
	for _, edge := range proc.Graph.Edges {
		if edge.From == start.ID && edge.Class == vbacfg.EdgeNormal {
			queue = append(queue, edge.To)
		}
	}
	seen := map[vbacfg.BlockID]bool{}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		block, exists := errorResultBlock(*proc.Graph, current)
		if !exists || block.Statement == nil {
			continue
		}
		statement := byID[block.StatementID]
		if statementWritesName(proc, statement.ID, target) {
			continue
		}
		if errorResultCondition(statement.Kind) && statementReadsName(proc, statement.ID, target) {
			return true
		}
		for _, edge := range proc.Graph.Edges {
			if edge.From == current && edge.Class == vbacfg.EdgeNormal {
				queue = append(queue, edge.To)
			}
		}
	}
	return false
}

func callerArgumentExpression(proc sourceProcedure, call procedureir.CallSite, output effects.ErrorFailureOutput) (string, bool) {
	expressionID := 0
	if output.ParameterIndex >= 0 && output.ParameterIndex < len(call.Arguments.ExpressionIDs) {
		expressionID = call.Arguments.ExpressionIDs[output.ParameterIndex]
	} else {
		for _, named := range call.Arguments.Named {
			if strings.EqualFold(named.Name, output.ParameterName) {
				if named.ValueText != "" {
					return named.ValueText, true
				}
				expressionID = named.ExpressionID
				break
			}
		}
	}
	if expressionID == 0 {
		return "", false
	}
	for _, expression := range proc.Expressions {
		if expression.ID == expressionID {
			return expression.Text, true
		}
	}
	return "", false
}

func procedureLocal(proc sourceProcedure, name string) bool {
	for _, declaration := range proc.Declarations {
		if strings.EqualFold(declaration.Name, name) {
			return true
		}
	}
	return false
}

func directSuccessFlag(summary effects.ProcedureSummary) bool {
	for _, evidence := range summary.Error.Direct {
		if evidence.Behavior == effects.ErrorReturnsSuccess && evidence.Origin.Key() == summary.Identity.Key() {
			return true
		}
	}
	return false
}

func errorEvidenceAt(items []effects.ErrorEvidence, behavior effects.ErrorBehaviorKind, target effects.ErrorEvidence) bool {
	for _, item := range items {
		if item.Behavior == behavior && item.StatementID == target.StatementID && item.Range.StartLine == target.Range.StartLine {
			return true
		}
	}
	return false
}

func representativePublicErrorChain(identities []effects.ProcedureIdentity, project effects.ProjectSummary, loss effects.ErrorEvidence) string {
	for _, identity := range identities {
		if !errorEntryProcedure(identity) {
			continue
		}
		summary, ok := project.Lookup(identity)
		if !ok {
			continue
		}
		for _, evidence := range summary.Error.Propagated {
			if evidence.Behavior == loss.Behavior && evidence.Origin.Key() == loss.Origin.Key() &&
				evidence.StatementID == loss.StatementID && evidence.Range.StartLine == loss.Range.StartLine {
				parts := make([]string, 0, len(evidence.CallChain))
				for _, identity := range evidence.CallChain {
					name := identity.QualifiedName
					if name == "" {
						name = identity.Name
					}
					if name != "" {
						parts = append(parts, name)
					}
				}
				if len(parts) == 0 {
					parts = append(parts, summary.Identity.QualifiedName, loss.Origin.QualifiedName)
				}
				return strings.Join(parts, " -> ")
			}
		}
	}
	return ""
}

func errorEntryProcedure(identity effects.ProcedureIdentity) bool {
	return identity.IsEventHandler || isProjectVisibleProcedure(identity)
}

func errorSuccessResultUseUncertain(proc sourceProcedure, call procedureir.CallSite, statement procedureir.Statement) bool {
	for _, other := range proc.Calls {
		if other.ID != call.ID && other.StatementID == call.StatementID {
			return true
		}
	}
	if statement.Kind != procedureir.StatementAssignment {
		return !errorResultCondition(statement.Kind) && statement.Kind != procedureir.StatementCall
	}
	if statement.Target == nil || statement.Value == nil {
		return true
	}
	target := strings.TrimSpace(statement.Target.Text)
	if !errorSuccessIdentifierRE.MatchString(target) {
		return true
	}
	if strings.EqualFold(target, proc.Name) {
		return !strings.EqualFold(strings.TrimSpace(proc.ReturnType), "Boolean") ||
			(proc.ProcedureKind != procedureir.ProcedureFunction && proc.ProcedureKind != procedureir.ProcedurePropertyGet)
	}
	return !booleanProcedureLocal(proc, target)
}

func errorSuccessResultChecked(proc sourceProcedure, call procedureir.CallSite, statement procedureir.Statement) bool {
	switch statement.Kind {
	case procedureir.StatementIf, procedureir.StatementElseIf, procedureir.StatementSelect,
		procedureir.StatementCase, procedureir.StatementDo, procedureir.StatementWhile:
		return true
	}
	if statement.Kind != procedureir.StatementAssignment || statement.Target == nil || statement.Value == nil {
		return false
	}
	target := strings.TrimSpace(statement.Target.Text)
	if target == "" || !errorSuccessIdentifierRE.MatchString(target) {
		return false
	}
	if strings.EqualFold(target, proc.Name) {
		return strings.EqualFold(strings.TrimSpace(proc.ReturnType), "Boolean") &&
			(proc.ProcedureKind == procedureir.ProcedureFunction || proc.ProcedureKind == procedureir.ProcedurePropertyGet)
	}
	if !booleanProcedureLocal(proc, target) {
		return false
	}
	return booleanResultCheckedBeforeExit(proc, statement.ID, target)
}

func booleanResultCheckedBeforeExit(proc sourceProcedure, assignmentID int, target string) bool {
	if proc.Graph == nil {
		return false
	}
	assignment, ok := proc.Graph.BlockForStatement(assignmentID)
	if !ok {
		return false
	}
	byID := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		byID[statement.ID] = statement
	}
	queue := []vbacfg.BlockID{}
	for _, edge := range proc.Graph.Edges {
		if edge.From == assignment.ID && edge.Class == vbacfg.EdgeNormal {
			queue = append(queue, edge.To)
		}
	}
	seen := map[vbacfg.BlockID]bool{}
	foundCheck := false
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		if current == proc.Graph.NormalExit || current == proc.Graph.UnknownExit || current == proc.Graph.TerminationExit {
			return false
		}
		block, exists := errorResultBlock(*proc.Graph, current)
		if !exists || block.Statement == nil {
			continue
		}
		statement := byID[block.StatementID]
		if statementWritesName(proc, statement.ID, target) {
			return false
		}
		if errorResultCondition(statement.Kind) && statementReadsName(proc, statement.ID, target) {
			foundCheck = true
			continue
		}
		for _, edge := range proc.Graph.Edges {
			if edge.From == current && edge.Class == vbacfg.EdgeNormal {
				queue = append(queue, edge.To)
			}
		}
	}
	return foundCheck
}

func errorResultBlock(graph vbacfg.Graph, id vbacfg.BlockID) (vbacfg.Block, bool) {
	for _, block := range graph.Blocks {
		if block.ID == id {
			return block, true
		}
	}
	return vbacfg.Block{}, false
}

func errorResultCondition(kind procedureir.StatementKind) bool {
	switch kind {
	case procedureir.StatementIf, procedureir.StatementElseIf, procedureir.StatementSelect,
		procedureir.StatementCase, procedureir.StatementDo, procedureir.StatementWhile:
		return true
	default:
		return false
	}
}

func booleanProcedureLocal(proc sourceProcedure, name string) bool {
	for _, declaration := range proc.Declarations {
		if strings.EqualFold(declaration.Name, name) && strings.EqualFold(strings.TrimSpace(declaration.Type), "Boolean") {
			return true
		}
	}
	return false
}

func statementReadsName(proc sourceProcedure, statementID int, name string) bool {
	for _, access := range proc.Accesses {
		if access.StatementID == statementID && strings.EqualFold(access.Name, name) &&
			(access.Mode == procedureir.AccessRead || access.Mode == procedureir.AccessReadWrite) {
			return true
		}
	}
	return false
}

func statementWritesName(proc sourceProcedure, statementID int, name string) bool {
	for _, access := range proc.Accesses {
		if access.StatementID == statementID && strings.EqualFold(access.Name, name) &&
			(access.Mode == procedureir.AccessWrite || access.Mode == procedureir.AccessReadWrite) {
			return true
		}
	}
	return false
}
