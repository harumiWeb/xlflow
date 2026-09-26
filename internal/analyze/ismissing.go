package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func (a Analyzer) invalidIsMissingUsageFindings(file parsedFile, proc sourceProcedure, resolver procedureir.Resolver) []Finding {
	if !a.Config.Analyze.DetectInvalidIsMissingUsage || proc.IR == nil || proc.Facts == nil ||
		proc.IR.Symbol.Recovered || len(proc.IR.Symbol.ConditionalBranches) > 0 {
		return nil
	}

	parameters := make(map[string]procedureir.Parameter, len(proc.IR.Symbol.Parameters))
	ambiguousParameters := make(map[string]struct{})
	for _, parameter := range proc.IR.Symbol.Parameters {
		name := parameterName(parameter.Name)
		if name == "" {
			continue
		}
		if _, exists := parameters[name]; exists {
			ambiguousParameters[name] = struct{}{}
			continue
		}
		parameters[name] = parameter
	}

	var findings []Finding
	for call := range proc.Facts.Calls().All() {
		if !isResolvedIsMissingIntrinsic(call, resolver) || procedureir.IsAssignmentTargetCall(call, *proc.IR) {
			continue
		}
		statement, ok := proc.Facts.Statement(call.StatementID)
		if !ok || statement.Recovered || statement.Kind == procedureir.StatementRecovered ||
			statement.Kind == procedureir.StatementUnknown || len(statement.ConditionalBranches) > 0 {
			continue
		}
		if call.ExpressionID > 0 {
			expression, ok := proc.Facts.Expression(call.ExpressionID)
			if !ok || expression.Recovered {
				continue
			}
		}
		if call.Arguments.Count != 1 || len(call.Arguments.ExpressionIDs) != 1 || call.Arguments.ExpressionIDs[0] <= 0 {
			continue
		}

		argument, ok := proc.Facts.Expression(call.Arguments.ExpressionIDs[0])
		if !ok || argument.Recovered || argument.Range.StartLine <= 0 || argument.Range.EndLine <= 0 {
			continue
		}
		parameterReference, ok := unparenthesizedExpression(proc.Facts, argument, len(proc.IR.Expressions))
		if !ok || parameterReference.Recovered {
			continue
		}

		var parameter procedureir.Parameter
		isProcedureParameter := false
		if parameterReference.Kind == procedureir.ExpressionIdentifier {
			name := parameterName(parameterReference.Text)
			if _, ambiguous := ambiguousParameters[name]; ambiguous {
				continue
			}
			parameter, isProcedureParameter = parameters[name]
		}
		if isProcedureParameter && effectiveIsMissingParameter(parameter) {
			continue
		}

		message, reason := invalidIsMissingArgumentMessage(argument, parameter, isProcedureParameter)
		finding := a.simpleFinding(file, proc, argument.Range.StartLine, "VBA283", "warning", message, reason,
			"Pass an Optional Variant parameter, or remove the IsMissing check.")
		finding.Column = argument.Range.StartColumn
		finding.EndLine = argument.Range.EndLine
		finding.EndColumn = argument.Range.EndColumn
		findings = append(findings, finding)
	}
	return findings
}

func isResolvedIsMissingIntrinsic(call procedureir.CallSite, resolver procedureir.Resolver) bool {
	resolution := call.Resolution
	if resolution.Status == procedureir.ResolutionNotAttempted && resolver != nil {
		resolution = resolver.ResolveCall(call)
	}
	if call.IsRaiseEvent || !strings.EqualFold(cleanIdentifier(call.Callee.BaseName), "IsMissing") ||
		resolution.Status != procedureir.ResolutionBuiltinLike {
		return false
	}
	if call.Callee.Receiver == nil {
		return !strings.HasPrefix(strings.TrimSpace(call.Callee.Text), ".")
	}
	return strings.EqualFold(cleanIdentifier(*call.Callee.Receiver), "VBA") &&
		call.MemberOperator == procedureir.MemberOperatorDot
}

func unparenthesizedExpression(facts *procedureAnalysisFacts, expression procedureir.Expression, expressionCount int) (procedureir.Expression, bool) {
	for depth := 0; expression.Kind == procedureir.ExpressionParentheses; depth++ {
		if expression.Recovered || len(expression.Children) != 1 || depth >= expressionCount {
			return procedureir.Expression{}, false
		}
		child, ok := facts.Expression(expression.Children[0])
		if !ok || child.Recovered {
			return procedureir.Expression{}, false
		}
		expression = child
	}
	return expression, true
}

func effectiveIsMissingParameter(parameter procedureir.Parameter) bool {
	return !parameter.Recovered && parameter.Optional && !parameter.ParamArray && !parameter.IsArray &&
		(strings.TrimSpace(parameter.Type) == "" || strings.EqualFold(strings.TrimSpace(parameter.Type), "Variant"))
}

func invalidIsMissingArgumentMessage(argument procedureir.Expression, parameter procedureir.Parameter, isProcedureParameter bool) (string, string) {
	name := strings.TrimSpace(argument.Text)
	if !isProcedureParameter {
		return "IsMissing argument " + name + " is not an Optional Variant parameter of this procedure.",
			"The argument does not directly name a parameter declared by the containing procedure."
	}
	parameterName := cleanIdentifier(parameter.Name)
	switch {
	case !parameter.Optional:
		return "IsMissing argument " + parameterName + " is not optional.",
			"The parameter is not declared Optional, so the caller cannot omit it."
	case parameter.ParamArray || parameter.IsArray:
		return "IsMissing argument " + parameterName + " is an array parameter.",
			"An array or ParamArray value does not carry the omitted-argument state IsMissing checks."
	case strings.TrimSpace(parameter.Type) != "" && !strings.EqualFold(strings.TrimSpace(parameter.Type), "Variant"):
		return "IsMissing argument " + parameterName + " is not Variant.",
			"IsMissing only detects omission for Optional Variant parameters; this parameter has type " + parameter.Type + "."
	default:
		return "IsMissing argument expression " + name + " cannot indicate an omitted Optional Variant parameter.",
			"Only a direct reference to an eligible parameter of the containing procedure can represent an omitted argument."
	}
}
