package effects

import (
	"sort"
	"strconv"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

const (
	errorTargetHandler    = "handler"
	errorTargetCleanup    = "cleanup"
	errorTargetResumeNext = "resume_next"
)

type loggerContract map[int]bool

func extractErrorSummary(summary *ProcedureSummary, proc procedureir.ProcedureIR, graph cfg.Graph, reachable map[int]bool, candidateKeys map[string]string, loggerTargets map[string]loggerContract, rethrowTargets, terminalTargets map[string]bool) {
	statements := append([]procedureir.Statement(nil), proc.Statements...)
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartByte != statements[j].Range.StartByte {
			return statements[i].Range.StartByte < statements[j].Range.StartByte
		}
		return statements[i].ID < statements[j].ID
	})

	labels := map[string]procedureir.Statement{}
	checkedResumeNext := false
	for _, statement := range statements {
		if statement.Kind == procedureir.StatementLabel {
			labels[normalizedErrorLabel(statement.Label)] = statement
		}
	}

	for i, statement := range statements {
		if !reachable[statement.ID] || statement.Recovered {
			continue
		}
		if isRaiseStatement(proc, statement) {
			addErrorEvidence(summary, statement.Range, statement.ID, 0, ErrorMayRaise, "raise", strings.TrimSpace(statement.Text))
		}
		if statement.Kind != procedureir.StatementOnError || statement.Control == nil {
			continue
		}
		switch statement.Control.Transfer {
		case procedureir.TransferOnErrorResumeNext:
			addErrorEvidence(summary, statement.Range, statement.ID, 0, ErrorUsesResumeNext, errorTargetResumeNext, "")
			if checkedResumeNextProbe(statements, i, reachable) || explicitResumeNextFallback(proc, graph, statements, i, reachable) {
				checkedResumeNext = true
			} else {
				addErrorEvidence(summary, statement.Range, statement.ID, 0, ErrorSuppresses, errorTargetResumeNext, "unchecked or broad scope")
			}
		case procedureir.TransferOnErrorGoto:
			target := normalizedErrorLabel(statement.Control.Target)
			if target == "" {
				continue
			}
			label, ok := labels[target]
			if !ok {
				continue
			}
			addErrorEvidence(summary, statement.Range, statement.ID, 0, ErrorHasHandler, errorTargetHandler, target)
			outcome := inspectHandler(proc, graph, label.ID, candidateKeys, loggerTargets, rethrowTargets, terminalTargets)
			if outcome.swallows {
				fallback, ok := dominatingResultFallback(proc, graph, statement.ID)
				if !ok {
					fallback, ok = immediateLiteralResultFallback(proc, statement)
				}
				if ok {
					outcome.swallows = false
					if booleanFalseAssignment(proc, fallback) {
						outcome.returnsSuccess = true
						outcome.returnRange, outcome.returnStatementID = fallback.Range, fallback.ID
					}
				}
			}
			if outcome.rethrows {
				addErrorEvidence(summary, outcome.rethrowRange, outcome.rethrowStatementID, 0, ErrorRethrows, errorTargetHandler, target)
			}
			if outcome.returnsSuccess && explicitBooleanResults(proc, reachable) {
				addErrorEvidence(summary, outcome.returnRange, outcome.returnStatementID, 0, ErrorReturnsSuccess, "boolean", target)
				summary.Error.Direct[len(summary.Error.Direct)-1].FailureOutputs = append([]ErrorFailureOutput(nil), outcome.failureOutputs...)
			}
			if outcome.swallows {
				category := errorTargetHandler
				if strings.Contains(target, "clean") || strings.Contains(target, "finally") {
					category = errorTargetCleanup
				}
				addErrorEvidence(summary, label.Range, label.ID, 0, ErrorSuppresses, category, target)
				if outcome.logs {
					addErrorEvidence(summary, label.Range, label.ID, 0, ErrorLogsAndContinues, category, target)
				}
			}
		}
	}
	for _, edge := range graph.ExitTransitions(cfg.ExitSelection{Exceptional: true}, cfg.EdgeFilter{}) {
		if edge.Kind != cfg.EdgeError {
			continue
		}
		addErrorEvidence(summary, edge.Range, edge.StatementID, 0, ErrorMayRaise, "unhandled_fault", "exceptional_exit")
		break
	}

	// A Boolean Try-style helper is a stable caller contract only when both
	// success and failure values are explicit. This also covers helpers that
	// use a checked Resume Next probe rather than a label handler.
	if checkedResumeNext && explicitBooleanResults(proc, reachable) {
		statement := firstBooleanFailureAssignment(proc, reachable)
		if statement.ID != 0 {
			addErrorEvidence(summary, statement.Range, statement.ID, 0, ErrorReturnsSuccess, "boolean", "False")
		}
	}
}

func explicitResumeNextFallback(proc procedureir.ProcedureIR, graph cfg.Graph, statements []procedureir.Statement, index int, reachable map[int]bool) bool {
	if index < 0 || index >= len(statements) {
		return false
	}
	fallback, ok := dominatingResultFallback(proc, graph, statements[index].ID)
	if !ok || !booleanFalseAssignment(proc, fallback) {
		return false
	}
	probeAssignments := 0
	for _, statement := range statements[index+1:] {
		if !reachable[statement.ID] || statement.Recovered || statement.Kind == procedureir.StatementDeclaration || statement.Kind == procedureir.StatementLabel {
			continue
		}
		if statement.Kind == procedureir.StatementOnError && statement.Control != nil {
			return probeAssignments == 1 && (statement.Control.Transfer == procedureir.TransferOnErrorDisable || statement.Control.Transfer == procedureir.TransferOnErrorGoto)
		}
		if strings.EqualFold(assignmentProbeTarget(statement), proc.Symbol.Name) {
			probeAssignments++
			continue
		}
		return false
	}
	return false
}

func immediateLiteralResultFallback(proc procedureir.ProcedureIR, setup procedureir.Statement) (procedureir.Statement, bool) {
	for _, statement := range proc.Statements {
		if statement.Range.StartByte <= setup.Range.StartByte || statement.Kind == procedureir.StatementDeclaration || statement.Kind == procedureir.StatementLabel {
			continue
		}
		if !resultAssignment(proc, statement) || statement.Value == nil {
			return procedureir.Statement{}, false
		}
		if safeLiteralAssignment(statement) {
			return statement, true
		}
		return procedureir.Statement{}, false
	}
	return procedureir.Statement{}, false
}

func isNumericLiteral(value string) bool {
	if value == "" {
		return false
	}
	_, err := strconv.ParseFloat(strings.TrimRight(strings.TrimSpace(value), "%&^!#@"), 64)
	return err == nil
}

func safeLiteralAssignment(statement procedureir.Statement) bool {
	value := strings.TrimSpace(assignmentProbeValue(statement))
	lower := strings.ToLower(value)
	return lower == "true" || lower == "false" || lower == "nothing" || lower == "empty" || lower == "null" ||
		strings.HasPrefix(value, "\"") && strings.HasSuffix(value, "\"") || isNumericLiteral(value)
}

func dominatingResultFallback(proc procedureir.ProcedureIR, graph cfg.Graph, handlerSetupStatementID int) (procedureir.Statement, bool) {
	setup, ok := graph.BlockForStatement(handlerSetupStatementID)
	if !ok {
		return procedureir.Statement{}, false
	}
	dominators := graph.Dominators(cfg.EdgeFilter{NormalOnly: true})[setup.ID]
	dominates := map[cfg.BlockID]bool{}
	for _, id := range dominators {
		dominates[id] = true
	}
	var fallback procedureir.Statement
	for _, statement := range proc.Statements {
		if !resultAssignment(proc, statement) {
			continue
		}
		block, exists := graph.BlockForStatement(statement.ID)
		if exists && dominates[block.ID] && (fallback.ID == 0 || statement.Range.StartByte > fallback.Range.StartByte) {
			fallback = statement
		}
	}
	return fallback, fallback.ID != 0
}

type handlerOutcome struct {
	swallows           bool
	rethrows           bool
	returnsSuccess     bool
	logs               bool
	rethrowRange       vbaast.Range
	rethrowStatementID int
	returnRange        vbaast.Range
	returnStatementID  int
	failureOutputs     []ErrorFailureOutput
}

// inspectHandler walks the actual exceptional target subgraph. A path is
// intentionally handled when it explicitly raises, resumes, or returns a
// fallback value. Any remaining path to normal exit loses the original error.
func inspectHandler(proc procedureir.ProcedureIR, graph cfg.Graph, labelStatementID int, candidateKeys map[string]string, loggerTargets map[string]loggerContract, rethrowTargets, terminalTargets map[string]bool) handlerOutcome {
	graph = graph.WithoutNormalErrRaiseContinuation()
	start, ok := graph.BlockForStatement(labelStatementID)
	if !ok {
		return handlerOutcome{swallows: true}
	}
	byID := statementIndex(proc)
	type state struct {
		block       cfg.BlockID
		returned    bool
		recovered   bool
		errorValues string
		errorNumber string
	}
	queue := []state{{block: start.ID}}
	seen := map[state]bool{}
	out := handlerOutcome{}
	hasGuardedRethrow := handlerSubgraphHasRaise(proc, graph, start.ID)
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		if current.block == graph.NormalExit {
			if !current.returned && !current.recovered {
				out.swallows = true
			}
			continue
		}
		if current.block == graph.ExceptionalExit || current.block == graph.TerminationExit || current.block == graph.UnknownExit {
			continue
		}
		block := graphBlock(graph, current.block)
		errorGuard := errorGuardUnknown
		if block.Statement != nil {
			statement := byID[block.StatementID]
			if isRaiseStatement(proc, statement) || callsKnownTarget(proc, statement, candidateKeys, rethrowTargets) {
				out.rethrows = true
				if out.rethrowStatementID == 0 {
					out.rethrowRange, out.rethrowStatementID = statement.Range, statement.ID
				}
				continue
			}
			if callsKnownTarget(proc, statement, candidateKeys, terminalTargets) {
				continue
			}
			if statement.Kind == procedureir.StatementResume {
				continue
			}
			if resultAssignment(proc, statement) {
				current.returned = true
				if booleanFalseAssignment(proc, statement) {
					out.returnsSuccess = true
					if out.returnStatementID == 0 {
						out.returnRange, out.returnStatementID = statement.Range, statement.ID
					}
				}
			}
			if output, ok := byRefFailureOutput(proc, statement); ok {
				out.failureOutputs = appendUniqueFailureOutput(out.failureOutputs, output)
			}
			if statement.Kind == procedureir.StatementAssignment && statement.Target != nil && statement.Value != nil && containsErrorNumber(statement.Value.Text) {
				current.errorNumber = strings.ToLower(strings.TrimSpace(statement.Target.Text))
			}
			if errorLogStatement(statement, &current.errorValues) {
				out.logs = true
			}
			if callsKnownLogger(proc, statement, current.errorValues, candidateKeys, loggerTargets) {
				out.logs = true
			}
			errorGuard = errorPresentCondition(statement)
		}
		for _, edge := range graph.Edges {
			if edge.From != current.block {
				continue
			}
			// Fault transitions from statements in the handler describe a new
			// error, not an outcome of the original exceptional path.
			if edge.Kind == cfg.EdgeError && edge.Class == cfg.EdgeExceptional {
				continue
			}
			if errorGuard == errorGuardTrue && edge.Kind == cfg.EdgeBranchFalse {
				continue
			}
			if errorGuard == errorGuardFalse && edge.Kind == cfg.EdgeBranchTrue {
				continue
			}
			recovered := current.recovered
			if hasGuardedRethrow {
				if branch, ok := expectedErrorRecoveryBranch(block.Statement, current.errorNumber); ok && edge.Kind == branch {
					recovered = true
				}
			}
			queue = append(queue, state{block: edge.To, returned: current.returned, recovered: recovered, errorValues: current.errorValues, errorNumber: current.errorNumber})
		}
	}
	return out
}

// rethrowProcedureIndex recognizes project-local wrappers whose every normal
// path terminates in Err.Raise/Error or in another already-proven wrapper.
// The finite iteration also handles short wrapper chains without guessing
// about unresolved or ambiguous calls.
func rethrowProcedureIndex(inputs []procedureInput, candidateKeys map[string]string) map[string]bool {
	out := map[string]bool{}
	changed := true
	for changed {
		changed = false
		for _, input := range inputs {
			if out[input.id.Key()] || !procedureAlwaysRethrows(input.proc, input.graph, candidateKeys, out) {
				continue
			}
			out[input.id.Key()] = true
			changed = true
		}
	}
	return out
}

func terminalProcedureIndex(inputs []procedureInput, candidateKeys map[string]string) map[string]bool {
	out := map[string]bool{}
	changed := true
	for changed {
		changed = false
		for _, input := range inputs {
			if out[input.id.Key()] || !procedureAlwaysTerminates(input.proc, input.graph, candidateKeys, out) {
				continue
			}
			out[input.id.Key()] = true
			changed = true
		}
	}
	return out
}

func procedureAlwaysTerminates(proc procedureir.ProcedureIR, graph cfg.Graph, candidateKeys map[string]string, known map[string]bool) bool {
	if len(graph.Blocks) == 0 {
		return false
	}
	byID := statementIndex(proc)
	queue := []cfg.BlockID{graph.Entry}
	seen := map[cfg.BlockID]bool{}
	sawTerminal := false
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		if current == graph.NormalExit || current == graph.UnknownExit {
			return false
		}
		if current == graph.TerminationExit {
			sawTerminal = true
			continue
		}
		block := graphBlock(graph, current)
		if block.Statement != nil {
			statement := byID[block.StatementID]
			if isRaiseStatement(proc, statement) || callsKnownTarget(proc, statement, candidateKeys, known) {
				sawTerminal = true
				continue
			}
		}
		for _, edge := range graph.Edges {
			if edge.From == current && edge.Class == cfg.EdgeNormal {
				queue = append(queue, edge.To)
			}
		}
	}
	return sawTerminal
}

func procedureAlwaysRethrows(proc procedureir.ProcedureIR, graph cfg.Graph, candidateKeys map[string]string, known map[string]bool) bool {
	if len(graph.Blocks) == 0 {
		return false
	}
	byID := statementIndex(proc)
	queue := []cfg.BlockID{graph.Entry}
	seen := map[cfg.BlockID]bool{}
	sawRethrow := false
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		if current == graph.NormalExit || current == graph.UnknownExit || current == graph.TerminationExit {
			return false
		}
		block := graphBlock(graph, current)
		if block.Statement != nil {
			statement := byID[block.StatementID]
			if isRaiseStatement(proc, statement) || callsKnownTarget(proc, statement, candidateKeys, known) {
				sawRethrow = true
				continue
			}
		}
		for _, edge := range graph.Edges {
			if edge.From == current && edge.Class == cfg.EdgeNormal {
				queue = append(queue, edge.To)
			}
		}
	}
	return sawRethrow
}

func callsKnownTarget(proc procedureir.ProcedureIR, statement procedureir.Statement, candidateKeys map[string]string, targets map[string]bool) bool {
	for _, call := range proc.Calls {
		if call.StatementID != statement.ID || call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
			continue
		}
		if key, ok := candidateKeys[candidateKey(call.Resolution.Candidates[0])]; ok && targets[key] {
			return true
		}
	}
	return false
}

func handlerSubgraphHasRaise(proc procedureir.ProcedureIR, graph cfg.Graph, start cfg.BlockID) bool {
	byID := statementIndex(proc)
	queue := []cfg.BlockID{start}
	seen := map[cfg.BlockID]bool{}
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if seen[current] {
			continue
		}
		seen[current] = true
		block := graphBlock(graph, current)
		if block.Statement != nil && isRaiseStatement(proc, byID[block.StatementID]) {
			return true
		}
		for _, edge := range graph.Edges {
			if edge.From == current && edge.Kind != cfg.EdgeError {
				queue = append(queue, edge.To)
			}
		}
	}
	return false
}

func containsErrorNumber(text string) bool {
	return strings.Contains(strings.ToLower(stripVBStringLiterals(text)), "err.number")
}

func expectedErrorRecoveryBranch(statement *procedureir.Statement, copiedNumber string) (cfg.EdgeKind, bool) {
	if statement == nil || statement.Condition == nil || copiedNumber == "" {
		return "", false
	}
	condition := strings.ToLower(statement.Condition.Text)
	condition = strings.NewReplacer(" ", "", "\t", "", "(", "", ")", "").Replace(condition)
	for _, comparison := range []struct {
		op     string
		branch cfg.EdgeKind
	}{{"<>", cfg.EdgeBranchFalse}, {"=", cfg.EdgeBranchTrue}} {
		parts := strings.Split(condition, comparison.op)
		if len(parts) != 2 {
			continue
		}
		if strings.EqualFold(parts[0], copiedNumber) && nonZeroErrorCode(parts[1]) || strings.EqualFold(parts[1], copiedNumber) && nonZeroErrorCode(parts[0]) {
			return comparison.branch, true
		}
	}
	return "", false
}

func nonZeroErrorCode(text string) bool {
	value, err := strconv.ParseInt(strings.TrimSpace(text), 10, 64)
	return err == nil && value != 0
}

type errorGuardOutcome uint8

const (
	errorGuardUnknown errorGuardOutcome = iota
	errorGuardTrue
	errorGuardFalse
)

func errorPresentCondition(statement procedureir.Statement) errorGuardOutcome {
	if statement.Condition == nil {
		return errorGuardUnknown
	}
	condition := strings.ToLower(statement.Condition.Text)
	condition = strings.NewReplacer(" ", "", "\t", "", "(", "", ")", "").Replace(condition)
	switch {
	case strings.Contains(condition, "noterr.number=0"), strings.Contains(condition, "err.number<>0"), strings.Contains(condition, "0<>err.number"):
		return errorGuardTrue
	case strings.Contains(condition, "noterr.number<>0"), strings.Contains(condition, "err.number=0"), strings.Contains(condition, "0=err.number"):
		return errorGuardFalse
	default:
		return errorGuardUnknown
	}
}

// loggerProcedureIndex identifies wrappers from their bodies and uniquely
// resolved local calls, never from procedure names. A handler call still has
// to pass Err data before the wrapper is treated as error logging.
func loggerProcedureIndex(inputs []procedureInput, candidateKeys map[string]string) map[string]loggerContract {
	loggers := map[string]loggerContract{}
	procedures := map[string]procedureir.ProcedureIR{}
	for _, input := range inputs {
		procedures[input.id.Key()] = input.proc
	}
	for _, input := range inputs {
		reachable := reachableStatements(input.proc, input.graph)
		for _, statement := range input.proc.Statements {
			if reachable[statement.ID] && !statement.Recovered && recognizedLogSink(statement.Text) {
				contract := loggers[input.id.Key()]
				if contract == nil {
					contract = loggerContract{}
				}
				for index, parameter := range input.proc.Symbol.Parameters {
					if statementReadsParameter(input.proc, statement.ID, parameter.Name) {
						contract[index] = true
					}
				}
				if len(contract) > 0 {
					loggers[input.id.Key()] = contract
				}
			}
		}
	}
	changed := true
	for changed {
		changed = false
		for _, input := range inputs {
			if len(loggers[input.id.Key()]) > 0 {
				continue
			}
			reachable := reachableStatements(input.proc, input.graph)
			for _, call := range input.proc.Calls {
				if !reachable[call.StatementID] || call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
					continue
				}
				if target, ok := candidateKeys[candidateKey(call.Resolution.Candidates[0])]; ok && len(loggers[target]) > 0 {
					contract := forwardedLoggerParameters(input.proc, call, procedures[target], loggers[target])
					if len(contract) == 0 {
						continue
					}
					loggers[input.id.Key()] = contract
					changed = true
					break
				}
			}
		}
	}
	return loggers
}

func callsKnownLogger(proc procedureir.ProcedureIR, statement procedureir.Statement, copied string, candidateKeys map[string]string, loggerTargets map[string]loggerContract) bool {
	for _, call := range proc.Calls {
		if call.StatementID != statement.ID || call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
			continue
		}
		target, ok := candidateKeys[candidateKey(call.Resolution.Candidates[0])]
		if !ok {
			continue
		}
		for parameterIndex := range loggerTargets[target] {
			if expression, ok := callArgumentExpression(proc, call, procedureir.ProcedureIR{}, parameterIndex); ok && containsErrorValue(strings.ToLower(expression), copied) {
				return true
			}
		}
	}
	return false
}

func statementReadsParameter(proc procedureir.ProcedureIR, statementID int, name string) bool {
	for _, access := range proc.Accesses {
		if access.StatementID == statementID && access.Scope == procedureir.ScopeParameter && strings.EqualFold(access.Name, name) &&
			(access.Mode == procedureir.AccessRead || access.Mode == procedureir.AccessReadWrite) {
			return true
		}
	}
	return false
}

func forwardedLoggerParameters(caller procedureir.ProcedureIR, call procedureir.CallSite, callee procedureir.ProcedureIR, target loggerContract) loggerContract {
	out := loggerContract{}
	for parameterIndex := range target {
		expression, ok := callArgumentExpression(caller, call, callee, parameterIndex)
		if !ok {
			continue
		}
		for index, parameter := range caller.Symbol.Parameters {
			if identifierInExpression(expression, parameter.Name) {
				out[index] = true
			}
		}
	}
	return out
}

func callArgumentExpression(caller procedureir.ProcedureIR, call procedureir.CallSite, callee procedureir.ProcedureIR, parameterIndex int) (string, bool) {
	expressionID := 0
	if parameterIndex >= 0 && parameterIndex < len(call.Arguments.ExpressionIDs) {
		expressionID = call.Arguments.ExpressionIDs[parameterIndex]
	} else if parameterIndex >= 0 && parameterIndex < len(callee.Symbol.Parameters) {
		name := callee.Symbol.Parameters[parameterIndex].Name
		for _, named := range call.Arguments.Named {
			if strings.EqualFold(named.Name, name) {
				expressionID = named.ExpressionID
				if named.ValueText != "" {
					return named.ValueText, true
				}
				break
			}
		}
	}
	for _, expression := range caller.Expressions {
		if expression.ID == expressionID {
			return expression.Text, true
		}
	}
	return "", false
}

func identifierInExpression(expression, identifier string) bool {
	lower, name := strings.ToLower(expression), strings.ToLower(strings.TrimSpace(identifier))
	if name == "" {
		return false
	}
	for start := 0; start < len(lower); {
		relative := strings.Index(lower[start:], name)
		if relative < 0 {
			return false
		}
		index := start + relative
		beforeOK := index == 0 || !isIdentifierByte(lower[index-1])
		after := index + len(name)
		afterOK := after == len(lower) || !isIdentifierByte(lower[after])
		if beforeOK && afterOK {
			return true
		}
		start = after
	}
	return false
}

func isIdentifierByte(value byte) bool {
	return value == '_' || value >= 'a' && value <= 'z' || value >= '0' && value <= '9'
}

func checkedResumeNextProbe(statements []procedureir.Statement, index int, reachable map[int]bool) bool {
	faults, checked, restored := 0, false, false
	probeTarget := ""
	skipParent := 0
	parents := make(map[int]int, len(statements))
	for _, statement := range statements {
		parents[statement.ID] = statement.ParentID
	}
	for _, statement := range statements[index+1:] {
		if !reachable[statement.ID] || statement.Recovered {
			continue
		}
		if restored && skipParent != 0 && statementDescendsFrom(statement, skipParent, parents) {
			continue
		}
		if restored && statement.Kind == procedureir.StatementElse {
			skipParent = statement.ID
			continue
		}
		if statement.Kind == procedureir.StatementOnError && statement.Control != nil {
			if restored {
				break
			}
			restored = statement.Control.Transfer == procedureir.TransferOnErrorDisable || statement.Control.Transfer == procedureir.TransferOnErrorGoto
			if restored {
				continue
			}
			break
		}
		lower := strings.ToLower(statement.Text)
		conditionText := statement.Text
		if statement.Condition != nil {
			conditionText = statement.Condition.Text
		}
		probeObserved := errorProbeConditionKind(statement.Kind) &&
			(strings.Contains(lower, "err.") || probeTarget != "" && identifierInExpression(conditionText, probeTarget))
		assertsErr := strings.HasPrefix(strings.TrimSpace(lower), "debug.assert") && strings.Contains(lower, "err.")
		if probeObserved || assertsErr {
			checked = true
			if restored {
				break
			}
			continue
		}
		if (statement.Kind == procedureir.StatementAssignment || statement.Kind == procedureir.StatementSet) && containsErrorValue(lower, "") {
			// Capturing Err state or converting it into a Boolean flag is the
			// observation step, not a second compatibility probe.
			checked = true
			continue
		}
		if strings.Contains(lower, "err.clear") || statement.Kind == procedureir.StatementLabel || statement.Kind == procedureir.StatementDeclaration {
			continue
		}
		if restored {
			if probeTarget != "" && identifierInExpression(assignmentProbeValue(statement), probeTarget) {
				if derived := assignmentProbeTarget(statement); derived != "" {
					probeTarget = derived
					continue
				}
			}
			if safeLiteralAssignment(statement) {
				continue
			}
			break
		}
		if statement.Kind != procedureir.StatementElse && statement.Kind != procedureir.StatementExit {
			faults++
			if faults == 1 {
				probeTarget = assignmentProbeTarget(statement)
			}
		}
	}
	return faults == 1 && checked && restored
}

func statementDescendsFrom(statement procedureir.Statement, ancestor int, parents map[int]int) bool {
	for parent := statement.ParentID; parent != 0; parent = parents[parent] {
		if parent == ancestor {
			return true
		}
	}
	return false
}

func assignmentProbeValue(statement procedureir.Statement) string {
	if statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet {
		return ""
	}
	if index := strings.IndexByte(statement.Text, '='); index >= 0 && index+1 < len(statement.Text) {
		return strings.TrimSpace(statement.Text[index+1:])
	}
	if statement.Value != nil {
		return statement.Value.Text
	}
	return ""
}

func assignmentProbeTarget(statement procedureir.Statement) string {
	if statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet {
		return ""
	}
	text := strings.TrimSpace(statement.Text)
	if strings.HasPrefix(strings.ToLower(text), "set ") {
		text = strings.TrimSpace(text[4:])
	}
	index := strings.IndexByte(text, '=')
	if index <= 0 {
		return ""
	}
	target := strings.ToLower(strings.TrimSpace(text[:index]))
	for i := 0; i < len(target); i++ {
		if !isIdentifierByte(target[i]) {
			return ""
		}
	}
	return target
}

func errorProbeConditionKind(kind procedureir.StatementKind) bool {
	switch kind {
	case procedureir.StatementIf, procedureir.StatementElseIf, procedureir.StatementSelect,
		procedureir.StatementCase, procedureir.StatementDo, procedureir.StatementWhile:
		return true
	default:
		return false
	}
}

func explicitBooleanResults(proc procedureir.ProcedureIR, reachable map[int]bool) bool {
	if !strings.EqualFold(strings.TrimSpace(proc.Symbol.ReturnType), "Boolean") ||
		(proc.Symbol.Kind != procedureir.ProcedureFunction && proc.Symbol.Kind != procedureir.ProcedurePropertyGet && proc.Symbol.Kind != procedureir.ProcedureProperty) {
		return false
	}
	hasTrue, hasFalse := false, false
	for _, statement := range proc.Statements {
		if !reachable[statement.ID] || !resultAssignment(proc, statement) || statement.Value == nil {
			continue
		}
		switch strings.ToLower(strings.TrimSpace(statement.Value.Text)) {
		case "true":
			hasTrue = true
		case "false":
			hasFalse = true
		}
	}
	return hasTrue && hasFalse
}

func firstBooleanFailureAssignment(proc procedureir.ProcedureIR, reachable map[int]bool) procedureir.Statement {
	for _, statement := range proc.Statements {
		if reachable[statement.ID] && booleanFalseAssignment(proc, statement) {
			return statement
		}
	}
	return procedureir.Statement{}
}

func resultAssignment(proc procedureir.ProcedureIR, statement procedureir.Statement) bool {
	if (statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet) || statement.Target == nil {
		return false
	}
	target := strings.TrimSpace(statement.Target.Text)
	return strings.EqualFold(target, proc.Symbol.Name) ||
		len(target) > len(proc.Symbol.Name) && strings.EqualFold(target[:len(proc.Symbol.Name)], proc.Symbol.Name) && target[len(proc.Symbol.Name)] == '.'
}

func byRefFailureOutput(proc procedureir.ProcedureIR, statement procedureir.Statement) (ErrorFailureOutput, bool) {
	if (statement.Kind != procedureir.StatementAssignment && statement.Kind != procedureir.StatementSet) || statement.Target == nil || statement.Value == nil {
		return ErrorFailureOutput{}, false
	}
	for index, parameter := range proc.Symbol.Parameters {
		if !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByRef") || !strings.EqualFold(strings.TrimSpace(statement.Target.Text), parameter.Name) {
			continue
		}
		return ErrorFailureOutput{ParameterIndex: index, ParameterName: parameter.Name, Value: strings.TrimSpace(statement.Value.Text)}, true
	}
	return ErrorFailureOutput{}, false
}

func appendUniqueFailureOutput(items []ErrorFailureOutput, value ErrorFailureOutput) []ErrorFailureOutput {
	for _, item := range items {
		if item.ParameterIndex == value.ParameterIndex && strings.EqualFold(item.Value, value.Value) {
			return items
		}
	}
	return append(items, value)
}

func booleanFalseAssignment(proc procedureir.ProcedureIR, statement procedureir.Statement) bool {
	return strings.EqualFold(strings.TrimSpace(proc.Symbol.ReturnType), "Boolean") && resultAssignment(proc, statement) &&
		statement.Value != nil && strings.EqualFold(strings.TrimSpace(statement.Value.Text), "False")
}

func isRaiseStatement(proc procedureir.ProcedureIR, statement procedureir.Statement) bool {
	for _, call := range proc.Calls {
		if call.StatementID != statement.ID {
			continue
		}
		receiver := ""
		if call.Callee.Receiver != nil {
			receiver = strings.TrimSpace(*call.Callee.Receiver)
		}
		if strings.EqualFold(receiver, "Err") && strings.EqualFold(call.Callee.Member, "Raise") {
			return true
		}
		if statement.Kind == procedureir.StatementCall && strings.EqualFold(call.Callee.BaseName, "Error") {
			return true
		}
	}
	return false
}

func errorLogStatement(statement procedureir.Statement, copied *string) bool {
	lower := strings.ToLower(strings.TrimSpace(statement.Text))
	if statement.Kind == procedureir.StatementAssignment && statement.Target != nil && statement.Value != nil && containsErrorValue(strings.ToLower(statement.Value.Text), "") {
		*copied = strings.ToLower(strings.TrimSpace(statement.Target.Text))
		return false
	}
	return recognizedLogSink(lower) && containsErrorValue(lower, *copied)
}

func recognizedLogSink(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	return strings.HasPrefix(lower, "debug.print") || strings.HasPrefix(lower, "print #") || strings.HasPrefix(lower, "xlflowdebug.log")
}

func containsErrorValue(lower, copied string) bool {
	lower = stripVBStringLiterals(strings.ToLower(lower))
	if strings.Contains(lower, "err.number") || strings.Contains(lower, "err.description") || strings.Contains(lower, "err.source") || identifierInExpression(lower, "erl") {
		return true
	}
	return copied != "" && identifierInExpression(lower, copied)
}

func stripVBStringLiterals(text string) string {
	var out strings.Builder
	inString := false
	for i := 0; i < len(text); i++ {
		if text[i] != '"' {
			if !inString {
				out.WriteByte(text[i])
			}
			continue
		}
		if inString && i+1 < len(text) && text[i+1] == '"' {
			i++
			continue
		}
		inString = !inString
	}
	return out.String()
}

func graphBlock(graph cfg.Graph, id cfg.BlockID) cfg.Block {
	for _, block := range graph.Blocks {
		if block.ID == id {
			return block
		}
	}
	return cfg.Block{}
}

func normalizedErrorLabel(value string) string {
	return strings.ToLower(strings.Trim(strings.TrimSpace(value), "[]:"))
}

func addErrorEvidence(summary *ProcedureSummary, sourceRange vbaast.Range, statementID, callID int, behavior ErrorBehaviorKind, target, value string) {
	summary.Error.Direct = append(summary.Error.Direct, ErrorEvidence{
		Behavior: behavior, Origin: summary.Identity, CallChain: []ProcedureIdentity{summary.Identity}, Range: sourceRange,
		StatementID: statementID, CallID: callID, Target: target, Value: value,
	})
}

func refreshErrorFlags(summary *ErrorSummary) {
	*summary = ErrorSummary{
		Direct: append([]ErrorEvidence(nil), summary.Direct...), Propagated: append([]ErrorEvidence(nil), summary.Propagated...),
	}
	for _, facts := range [][]ErrorEvidence{summary.Direct, summary.Propagated} {
		for _, fact := range facts {
			switch fact.Behavior {
			case ErrorHasHandler:
				summary.HasErrorHandler = true
			case ErrorUsesResumeNext:
				summary.UsesResumeNext = true
			case ErrorSuppresses:
				summary.SuppressesErrors = true
			case ErrorRethrows:
				summary.RethrowsErrors = true
			case ErrorReturnsSuccess:
				summary.ReturnsSuccessFlag = true
			case ErrorMayRaise:
				summary.MayRaise = true
			case ErrorLogsAndContinues:
				summary.LogsAndContinues = true
			}
		}
	}
}
