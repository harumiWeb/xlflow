package analyze

import (
	"path/filepath"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type bracketExpressionCandidate struct {
	Text  string
	Range vbaast.Range
}

// bracketExpressionFindings reports Excel host-evaluated bracket expressions
// from the procedure IR. It has no analyzer or configuration dependency so
// batch and realtime callers can use the same source fact.
func bracketExpressionFindings(rootDir string, file parsedFile, proc sourceProcedure) []Finding {
	if proc.IR == nil {
		return nil
	}
	candidates := bracketExpressionCandidates(proc.IR)
	if len(candidates) == 0 {
		return nil
	}

	relativePath, err := filepath.Rel(rootDir, file.Path)
	if err != nil {
		relativePath = file.Path
	}
	findings := make([]Finding, 0, len(candidates))
	for _, candidate := range candidates {
		line := candidate.Range.StartLine
		if line < 1 {
			continue
		}
		findings = append(findings, Finding{
			Code:         "VBA262",
			Severity:     "warning",
			File:         filepath.ToSlash(relativePath),
			Module:       file.Module,
			Procedure:    proc.Name,
			Line:         line,
			Column:       candidate.Range.StartColumn + 1,
			EndLine:      candidate.Range.EndLine,
			EndColumn:    candidate.Range.EndColumn + 1,
			ScopeEndLine: proc.EndLine,
			Message:      "Bracket expression " + candidate.Text + " is evaluated by the Excel host instead of normal VBA name and type checking.",
			Reason:       "Excel host-evaluated bracket expressions bypass normal VBA compile-time member and type validation.",
			Suggestion:   "Use an explicit Excel object-model expression such as Range(\"A1\") when practical.",
			NearbyCode:   nearby(file.Lines, line, 2),
		})
	}
	return findings
}

func bracketExpressionCandidates(proc *procedureir.ProcedureIR) []bracketExpressionCandidate {
	if proc == nil {
		return nil
	}
	candidates := make([]bracketExpressionCandidate, 0)
	for _, expression := range proc.Expressions {
		if !isBareBracketExpression(proc, expression) {
			continue
		}
		candidates = append(candidates, bracketExpressionCandidate{Text: strings.TrimSpace(expression.Text), Range: expression.Range})
	}
	return candidates
}

func isBareBracketExpression(proc *procedureir.ProcedureIR, expression procedureir.Expression) bool {
	if expression.Recovered || expression.Kind != procedureir.ExpressionIdentifier || expression.SyntaxKind != "identifier" {
		return false
	}
	text := strings.TrimSpace(expression.Text)
	if len(text) < 2 || text[0] != '[' || text[len(text)-1] != ']' {
		return false
	}
	if bracketExpressionInDeclaration(proc, expression) {
		return false
	}
	return !bracketExpressionInMember(proc, expression)
}

func bracketExpressionInDeclaration(proc *procedureir.ProcedureIR, expression procedureir.Expression) bool {
	if expression.StatementID <= 0 || expression.StatementID > len(proc.Statements) {
		return false
	}
	return proc.Statements[expression.StatementID-1].Kind == procedureir.StatementDeclaration
}

func bracketExpressionInMember(proc *procedureir.ProcedureIR, expression procedureir.Expression) bool {
	current := expression
	for current.ParentID > 0 && current.ParentID <= len(proc.Expressions) {
		parent := proc.Expressions[current.ParentID-1]
		switch parent.Kind {
		case procedureir.ExpressionParentheses:
			current = parent
			continue
		case procedureir.ExpressionMember:
			return true
		default:
			return false
		}
	}
	return false
}
