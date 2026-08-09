package analyze

import (
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// objectUseBeforeSetIRFindings keeps VBA202 procedure-local: local object
// variables begin unassigned, while module state is unknown at procedure
// entry. The CFG supplies all-path assignment facts, including the assignment
// of a For Each iterator only on the loop-body edge.
func (a Analyzer) objectUseBeforeSetIRFindings(file parsedFile, proc sourceProcedure, maybeInitializedByCall map[string]int) []Finding {
	if proc.Graph == nil {
		return nil
	}

	declarations := procedureDeclarations(file.Lines, proc)
	flow := vbacfg.EdgeFilter{NormalOnly: true}
	definite := proc.Graph.DefiniteAssignments(flow)
	expressions := make(map[int]procedureir.Expression, len(proc.Expressions))
	for _, expression := range proc.Expressions {
		expressions[expression.ID] = expression
	}
	statements := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statements[statement.ID] = statement
	}

	reported := map[string]bool{}
	var findings []Finding
	for _, access := range proc.Accesses {
		if access.Scope != procedureir.ScopeLocal ||
			(access.Mode != procedureir.AccessRead && access.Mode != procedureir.AccessReadWrite) ||
			!objectMemberReceiver(expressions, statements, access) {
			continue
		}
		key := strings.ToLower(access.Name)
		declaration, ok := declarations[key]
		if !ok || !declaration.Object || declaration.NewExpression || declaration.Static || reported[key] {
			continue
		}
		block, ok := proc.Graph.BlockForStatement(access.StatementID)
		if !ok || !proc.Graph.IsReachable(block.ID, flow) {
			continue
		}
		if objectDefinitelyAssigned(definite[block.ID], key) {
			continue
		}
		// Preserve VBA202's established "obvious Set" boundary. The rule does
		// not attempt path correlation for arbitrary If/Select conditions; CFG
		// precision is used here for constructs such as For Each that the old
		// linear scan could not represent.
		if explicitObjectSetBefore(proc.Statements, key, access.Range.StartLine) {
			continue
		}
		if line := maybeInitializedByCall[key]; line > 0 && line < access.Range.StartLine {
			continue
		}

		findings = append(findings, a.simpleFinding(
			file, proc, access.Range.StartLine, "VBA202", "warning",
			declaration.Name+" may be used before it is assigned with Set.",
			"Object variables are Nothing until assigned with Set; member access before initialization can raise runtime error 91.",
			"Assign `Set "+declaration.Name+" = ...` before using members, or guard `If "+declaration.Name+" Is Nothing Then`.",
		))
		reported[key] = true
	}
	return findings
}

func explicitObjectSetBefore(statements []procedureir.Statement, name string, line int) bool {
	for _, statement := range statements {
		if statement.Range.StartLine >= line || statement.Kind != procedureir.StatementSet {
			continue
		}
		match := setAssignRe.FindStringSubmatch(statement.Text)
		if len(match) > 0 && strings.EqualFold(match[1], name) {
			return true
		}
	}
	return false
}

func objectMemberReceiver(expressions map[int]procedureir.Expression, statements map[int]procedureir.Statement, access procedureir.VariableAccess) bool {
	expression, ok := expressions[access.ExpressionID]
	if !ok {
		return false
	}
	if parent, ok := expressions[expression.ParentID]; expression.ParentID != 0 && ok && parent.Kind == procedureir.ExpressionMember {
		return true
	}
	statement, ok := statements[access.StatementID]
	return ok && statement.Kind == procedureir.StatementWith &&
		(statement.TargetID == access.ExpressionID || statement.ValueID == access.ExpressionID)
}

func objectDefinitelyAssigned(assigned []vbacfg.Variable, name string) bool {
	for _, variable := range assigned {
		if variable.Scope == procedureir.ScopeLocal && strings.EqualFold(variable.Name, name) {
			return true
		}
	}
	return false
}
