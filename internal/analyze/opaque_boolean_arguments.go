package analyze

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// OpaqueBooleanContext is the machine-readable evidence attached to VBA248.
// It intentionally contains only counts and resolved parameter names; source
// expressions remain in the normal nearby-code projection.
type OpaqueBooleanContext struct {
	PositionalLiteralCount        int      `json:"positional_literal_count"`
	NamedArgumentCount            int      `json:"named_argument_count"`
	ParameterNames                []string `json:"parameter_names,omitempty"`
	OptionalBooleanParameterCount *int     `json:"optional_boolean_parameter_count,omitempty"`
}

type opaqueBooleanLiteral struct {
	ParameterIndex int
	Value          string
}

func (a Analyzer) opaqueBooleanArgumentFindings(file parsedFile, proc sourceProcedure, signatures map[string]procedureSignature) []Finding {
	if !a.Config.Analyze.DetectOpaqueBooleanArguments {
		return nil
	}
	expressions := make(map[int]string, len(proc.Expressions))
	for _, expression := range proc.Expressions {
		if expression.ID > 0 {
			expressions[expression.ID] = expression.Text
		}
	}
	findings := make([]Finding, 0)
	for _, call := range proc.Calls {
		if call.Arguments.Count == 0 || len(call.Arguments.ExpressionIDs) == 0 {
			continue
		}
		namedIDs := make(map[int]bool, len(call.Arguments.Named))
		for _, named := range call.Arguments.Named {
			if named.ExpressionID > 0 {
				namedIDs[named.ExpressionID] = true
			}
		}
		positionalLiteralCount := 0
		positionalLiteralIndexes := make([]int, 0)
		positionalLiterals := make([]opaqueBooleanLiteral, 0)
		positionalIndex := 0
		for _, expressionID := range call.Arguments.ExpressionIDs {
			if expressionID == 0 {
				positionalIndex++
				continue
			}
			if namedIDs[expressionID] {
				continue
			}
			if value, ok := booleanLiteralValue(expressions[expressionID]); ok {
				positionalLiteralCount++
				positionalLiteralIndexes = append(positionalLiteralIndexes, positionalIndex)
				positionalLiterals = append(positionalLiterals, opaqueBooleanLiteral{ParameterIndex: positionalIndex, Value: value})
			}
			positionalIndex++
		}
		if positionalLiteralCount == 0 {
			continue
		}
		signature, signatureOK := opaqueBooleanCallSignature(call, signatures)
		optionalBooleanCount := 0
		var optionalBooleanParameterCount *int
		parameterNames := []string(nil)
		if signatureOK {
			for _, parameter := range signature.Params {
				if strings.EqualFold(strings.TrimSpace(parameter.Type), "Boolean") && parameter.Optional {
					optionalBooleanCount++
				}
			}
			optionalBooleanParameterCount = &optionalBooleanCount
			parameterNames = positionalParameterNames(signature.Params, positionalLiteralIndexes)
		}
		// A single positional Boolean is only ambiguous when the resolved
		// signature exposes multiple optional Boolean switches. Named
		// arguments make that case readable and therefore suppress it.
		if positionalLiteralCount < 2 && (len(call.Arguments.Named) > 0 || optionalBooleanCount < 2) {
			continue
		}
		callee := strings.TrimSpace(call.Callee.Text)
		if callee == "" {
			callee = "procedure"
		}
		message := fmt.Sprintf("Call to %s passes %d Boolean literal(s) positionally, which obscures the requested behavior.", callee, positionalLiteralCount)
		reason := "Positional True/False arguments hide which behavior each switch controls, especially when a call contains more than one Boolean literal."
		suggestion := "Use named arguments, an enum, or separate procedures for distinct behaviors."
		if len(parameterNames) > 0 {
			examples := make([]string, 0, len(parameterNames))
			for _, literal := range positionalLiterals {
				if literal.ParameterIndex < 0 || literal.ParameterIndex >= len(signature.Params) {
					continue
				}
				if name := strings.TrimSpace(signature.Params[literal.ParameterIndex].Name); name != "" {
					examples = append(examples, name+":="+literal.Value)
				}
			}
			if len(examples) > 0 {
				suggestion = "Use named arguments such as " + strings.Join(examples, ", ") + ", or replace the switches with an enum or separate procedures."
			}
		}
		finding := a.simpleFinding(file, proc, call.Range.StartLine, "VBA248", "warning", message, reason, suggestion)
		finding.Column = call.Range.StartColumn + 1
		finding.EndLine = call.Range.EndLine
		finding.EndColumn = call.Range.EndColumn + 1
		finding.OpaqueBoolean = &OpaqueBooleanContext{
			PositionalLiteralCount:        positionalLiteralCount,
			NamedArgumentCount:            len(call.Arguments.Named),
			ParameterNames:                parameterNames,
			OptionalBooleanParameterCount: optionalBooleanParameterCount,
		}
		findings = append(findings, finding)
	}
	return findings
}

func booleanLiteralValue(text string) (string, bool) {
	text = strings.TrimSpace(text)
	for len(text) >= 2 && strings.HasPrefix(text, "(") && strings.HasSuffix(text, ")") {
		text = strings.TrimSpace(text[1 : len(text)-1])
	}
	switch {
	case strings.EqualFold(text, "True"):
		return "True", true
	case strings.EqualFold(text, "False"):
		return "False", true
	default:
		return "", false
	}
}

func opaqueBooleanCallSignature(call procedureir.CallSite, signatures map[string]procedureSignature) (procedureSignature, bool) {
	if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
		return procedureSignature{}, false
	}
	candidate := call.Resolution.Candidates[0]
	keys := []string{candidate.QualifiedName, call.Callee.Text, call.Callee.BaseName}
	for _, key := range keys {
		key = strings.ToLower(strings.TrimSpace(key))
		if key == "" {
			continue
		}
		if signature, ok := signatures[key]; ok {
			return signature, true
		}
		if index := strings.LastIndexByte(key, '.'); index >= 0 {
			if signature, ok := signatures[key[index+1:]]; ok {
				return signature, true
			}
		}
	}
	return procedureSignature{}, false
}

func positionalParameterNames(parameters []parameterInfo, indexes []int) []string {
	if len(indexes) == 0 {
		return nil
	}
	result := make([]string, 0, len(indexes))
	for _, index := range indexes {
		if index < 0 || index >= len(parameters) {
			continue
		}
		if name := strings.TrimSpace(parameters[index].Name); name != "" {
			result = append(result, name)
		}
	}
	return result
}
