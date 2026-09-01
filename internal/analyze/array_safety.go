package analyze

import (
	"context"
	"sort"
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func (a Analyzer) arrayLifecycleFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration) []Finding {
	variables := arrayVariables(file, proc, moduleDecls)
	capacityGuards := arrayResumeNextCapacityGuards(file, proc, variables)
	constants := arrayIntegerConstants(file, proc, a.visibleConstantValues, a.visibleConstants)
	return a.arrayLifecycleFindingsPrepared(file, proc, ctx, moduleDecls, variables, constants, capacityGuards)
}

// arrayLifecycleFindingsPrepared is used by the shared procedure-local array
// result. Its preparation inputs are explicit so enabled projections do not
// rebuild the same variable catalog, constant environment, or capacity guard
// facts independently.
func (a Analyzer) arrayLifecycleFindingsPrepared(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) []Finding {
	return a.arrayLifecycleFindingsPreparedWithRuntime(file, proc, ctx, moduleDecls, variables, constants, capacityGuards, nil)
}

// arrayLifecycleFindingsPreparedWithRuntime optionally collects the
// deterministic array runtime projection while the shared array worklist is
// already visiting the procedure. A non-nil sink removes the historical
// second VBA249 array walk without changing the standalone helper contract.
func (a Analyzer) arrayLifecycleFindingsPreparedWithRuntime(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, runtimeSink *[]Finding) []Finding {
	return a.arrayLifecycleFindingsPreparedWithRuntimeEntry(file, proc, ctx, moduleDecls, variables, constants, capacityGuards, runtimeSink, nil)
}

func (a Analyzer) arrayLifecycleFindingsPreparedWithRuntimeEntry(file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, runtimeSink *[]Finding, preparedEntry arrayFlowState) []Finding {
	findings, _ := a.arrayLifecycleFindingsPreparedWithRuntimeEntryContext(context.Background(), file, proc, ctx, moduleDecls, variables, constants, capacityGuards, runtimeSink, preparedEntry)
	return findings
}

func (a Analyzer) arrayLifecycleFindingsPreparedWithRuntimeEntryContext(cancelCtx context.Context, file parsedFile, proc sourceProcedure, ctx analysisContext, moduleDecls map[string]sourceDeclaration, variables map[string]arrayVariable, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard, runtimeSink *[]Finding, preparedEntry arrayFlowState) ([]Finding, error) {
	if err := cancelCtx.Err(); err != nil {
		return nil, err
	}
	objectArrayDiagnosticsApplicable := false
	for _, variable := range variables {
		if variable.isArray && variable.isObject {
			objectArrayDiagnosticsApplicable = true
			break
		}
	}
	if !a.Config.Analyze.DetectArrayLifecycleSafety && !a.Config.Analyze.DetectRedimPreserveDimension && !a.Config.Analyze.DetectObjectArrayComparison && !objectArrayDiagnosticsApplicable && runtimeSink == nil {
		return nil, nil
	}
	vba227Variables := variables
	var vba227ResumeNextBefore []bool
	if a.Config.Analyze.DetectArrayLifecycleSafety {
		vba227Variables = arrayVBA227Variables(variables, file, proc)
		vba227ResumeNextBefore = arrayVBA227ResumeNextPrefixes(file, proc)
	}
	comparisonFindings := a.arrayComparisonFindings(file, proc, variables)
	comparisonFindings = append(comparisonFindings, a.arrayForEachFindings(file, proc, variables, ctx)...)
	baseLaneRequested := a.Config.Analyze.DetectArrayLifecycleSafety || a.Config.Analyze.DetectRedimPreserveDimension || objectArrayDiagnosticsApplicable
	initial := preparedEntry
	if initial == nil {
		initial = arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, variables)
	}
	runtimeBase := arrayOptionBase(file)
	if proc.Graph == nil {
		findings := append([]Finding(nil), comparisonFindings...)
		if !baseLaneRequested && runtimeSink == nil {
			findings = append(findings, a.arrayLifecycleLinearFindings(file, proc, ctx, variables, initial, constants, capacityGuards)...)
			return uniqueArrayFindings(findings), nil
		}
		seen := map[string]bool{}
		for _, finding := range findings {
			seen[arrayFindingKey(finding)] = true
		}
		vba227Initial := arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, vba227Variables)
		baseState := initial
		vba227State := vba227Initial
		runtimeState := arrayInitialState(variables)
		probe := a
		probe.Config.Analyze.DetectArrayLifecycleSafety = true
		runtimeSeen := map[string]bool{}
		for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
			if line&255 == 0 {
				if err := cancelCtx.Err(); err != nil {
					return nil, err
				}
			}
			text := normalizedCodeLine(file.Lines[line-1])
			if baseLaneRequested {
				var baseIssues []Finding
				baseState, baseIssues = a.arrayTransfer(file, proc, ctx, variables, baseState, text, line, constants, capacityGuards)
				for _, finding := range baseIssues {
					if a.Config.Analyze.DetectArrayLifecycleSafety && finding.Code == "VBA227" {
						continue
					}
					key := arrayFindingKey(finding)
					if !seen[key] {
						seen[key] = true
						findings = append(findings, finding)
					}
				}
				if a.Config.Analyze.DetectArrayLifecycleSafety {
					var lifecycleIssues []Finding
					vba227State, lifecycleIssues = a.arrayVBA227Transfer(file, proc, ctx, vba227Variables, vba227State, text, line, constants, capacityGuards, vba227ResumeNextBefore)
					for _, finding := range lifecycleIssues {
						if finding.Code != "VBA227" {
							continue
						}
						key := arrayFindingKey(finding)
						if !seen[key] {
							seen[key] = true
							findings = append(findings, finding)
						}
					}
				}
			}
			if runtimeSink != nil {
				for _, issue := range deterministicArrayRuntimeIssues(text, line, runtimeState, variables, constants, proc, runtimeBase) {
					key := strconv.Itoa(issue.line) + ":" + issue.kind + ":" + issue.operationKey
					if runtimeSeen[key] {
						continue
					}
					runtimeSeen[key] = true
					message, reason, suggestion := deterministicRuntimeFailureText(issue.kind)
					finding := a.simpleFinding(file, proc, issue.line, "VBA249", "error", message, reason, suggestion)
					finding.RuntimeError = &RuntimeErrorContext{Kind: issue.kind}
					finding.arrayOperationKey = issue.operationKey
					*runtimeSink = append(*runtimeSink, finding)
				}
				runtimeState, _ = probe.arrayTransfer(file, proc, ctx, variables, runtimeState, text, line, constants, nil)
			}
		}
		sortFindings(findings)
		if runtimeSink != nil {
			sortFindings(*runtimeSink)
		}
		return uniqueArrayFindings(findings), nil
	}

	findings := append([]Finding(nil), comparisonFindings...)
	seen := map[string]bool{}
	for _, finding := range findings {
		seen[arrayFindingKey(finding)] = true
	}
	baseView := proc.Graph.View(vbacfg.EdgeFilter{})
	lanes := make([]arrayCFGWorklistLane, 0, 3)
	if baseLaneRequested {
		lanes = append(lanes, arrayCFGWorklistLane{
			Graph: &baseView, Initial: initial, Stats: ctx.arrayStats,
			Visit: func(text string, line int, in arrayFlowState) arrayFlowState {
				out, issues := a.arrayTransfer(file, proc, ctx, variables, in, text, line, constants, capacityGuards)
				for _, call := range arrayCallsAtLine(proc.Calls, line) {
					out = applyArrayModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
					out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, variables, moduleDecls)
				}
				for _, finding := range issues {
					if finding.Code == "VBA227" {
						continue
					}
					key := arrayFindingKey(finding)
					if !seen[key] {
						seen[key] = true
						findings = append(findings, finding)
					}
				}
				return out
			},
			EdgeState: func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
				out = applyArrayConditionalAllocationBranch(out, &baseView, block, edge)
				out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, variables)
				return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], variables, file, proc, moduleDecls)
			},
		})
	}
	if a.Config.Analyze.DetectArrayLifecycleSafety {
		vba227Graph := arrayVBA227Graph(proc, ctx)
		vba227Initial := arrayEntryStateForProcedure(file, proc, ctx, moduleDecls, vba227Variables)
		lanes = append(lanes, arrayCFGWorklistLane{
			Graph: &vba227Graph, Initial: vba227Initial, Stats: ctx.arrayStats, SourceLines: true,
			ReliableExceptional: func(statement *procedureir.Statement, in, out arrayFlowState) bool {
				return arrayAllocationTransferIsReliable(statement, in, out)
			},
			Visit: func(text string, line int, in arrayFlowState) arrayFlowState {
				out, issues := a.arrayVBA227Transfer(file, proc, ctx, vba227Variables, in, text, line, constants, capacityGuards, vba227ResumeNextBefore)
				for _, call := range arrayCallsAtLine(proc.Calls, line) {
					out = applyArrayModuleCallEffects(out, file, proc, call, ctx, vba227Variables, moduleDecls)
					out = applyArrayUnknownModuleCallEffects(out, file, proc, call, ctx, vba227Variables, moduleDecls)
				}
				for _, finding := range issues {
					if finding.Code != "VBA227" {
						continue
					}
					key := arrayFindingKey(finding)
					if !seen[key] {
						seen[key] = true
						findings = append(findings, finding)
					}
				}
				return out
			},
			EdgeState: func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
				out = applyArrayConditionalAllocationBranch(out, &vba227Graph, block, edge)
				out = applyArrayVBA227ConditionalReDimBranch(out, proc, block.Statement, edge, vba227Variables)
				out = arraySuccessfulConditionState(out, block.Statement, vba227Variables, vba227ResumeNextBefore, proc)
				out = applyArrayModuleCapacityGuardBranch(out, block.Statement, edge, file, proc, ctx, vba227Variables, moduleDecls)
				out = applyArrayNotEmptyGuardBranch(out, block.Statement, edge, proc, vba227Variables)
				out = applyArrayAllocationFlagBranch(out, block.Statement, edge, vba227Variables)
				out = applyArrayAllocationGuard(out, block.Statement, edge, ctx.arrayAllocationGuards, vba227Variables)
				out = applyArraySafeBoundGuard(out, block.Statement, edge, ctx.arraySafeBoundGuards, vba227Variables)
				out = applyArrayForBoundState(out, block.Statement, edge, vba227Variables)
				return applyArrayModuleConfigurationBranch(out, block.Statement, edge, ctx.arrayModuleConfigurations[file.Path], vba227Variables, file, proc, moduleDecls)
			},
		})
	}
	if runtimeSink != nil {
		probe := a
		probe.Config.Analyze.DetectArrayLifecycleSafety = true
		runtimeSeen := map[string]bool{}
		runtimeState := arrayInitialState(variables)
		lanes = append(lanes, arrayCFGWorklistLane{
			Graph: &baseView, Initial: runtimeState, Stats: ctx.arrayStats,
			Visit: func(text string, line int, in arrayFlowState) arrayFlowState {
				for _, issue := range deterministicArrayRuntimeIssues(text, line, in, variables, constants, proc, runtimeBase) {
					key := strconv.Itoa(issue.line) + ":" + issue.kind + ":" + issue.operationKey
					if runtimeSeen[key] {
						continue
					}
					runtimeSeen[key] = true
					message, reason, suggestion := deterministicRuntimeFailureText(issue.kind)
					finding := a.simpleFinding(file, proc, issue.line, "VBA249", "error", message, reason, suggestion)
					finding.RuntimeError = &RuntimeErrorContext{Kind: issue.kind}
					finding.arrayOperationKey = issue.operationKey
					*runtimeSink = append(*runtimeSink, finding)
				}
				out, _ := probe.arrayTransfer(file, proc, ctx, variables, in, text, line, constants, nil)
				return out
			},
		})
	}
	if err := walkArrayCFGCombined(cancelCtx, file.Lines, lanes); err != nil {
		return nil, err
	}
	if runtimeSink != nil {
		sortFindings(*runtimeSink)
	}
	sortFindings(findings)
	return findings, nil
}

func (a Analyzer) arrayForEachFindings(file parsedFile, proc sourceProcedure, variables map[string]arrayVariable, ctx analysisContext) []Finding {
	if !a.Config.Analyze.DetectArrayLifecycleSafety {
		return nil
	}
	var findings []Finding
	for statement := range proc.Statements.All() {
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
	facts := proc.analysisFacts()

	arrayNames := make([]string, 0, len(variables))
	nonArrayNames := procedureIRNonArrayNames(proc)
	for name, variable := range variables {
		if variable.isArray && !nonArrayNames[name] {
			arrayNames = append(arrayNames, name)
		}
	}
	sort.Strings(arrayNames)

	var findings []Finding
	facts.forEachExpression(func(comparison procedureir.Expression) {
		if comparison.SyntaxKind != "comparison_expression" || comparison.Recovered {
			return
		}
		statement, ok := facts.Statement(comparison.StatementID)
		if !ok || comparisonAssignmentCarrier(statement, comparison) {
			return
		}
		matched := map[string]bool{}
		for _, childID := range comparison.Children {
			child, ok := directComparisonOperand(facts, childID)
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
	})
	sortFindings(findings)
	return findings
}

func procedureIRNonArrayNames(proc sourceProcedure) map[string]bool {
	names := map[string]bool{}
	for declaration := range proc.Declarations.All() {
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

func directComparisonOperand(facts *procedureAnalysisFacts, id int) (procedureir.Expression, bool) {
	child, ok := facts.Expression(id)
	for ok && child.Kind == procedureir.ExpressionParentheses && len(child.Children) == 1 {
		child, ok = facts.Expression(child.Children[0])
	}
	return child, ok
}

func comparisonAssignmentCarrier(statement procedureir.Statement, expression procedureir.Expression) bool {
	if expression.ParentID != 0 {
		return false
	}
	return statement.Kind == procedureir.StatementAssignment || statement.Kind == procedureir.StatementSet
}

func (a Analyzer) arrayLifecycleLinearFindings(file parsedFile, proc sourceProcedure, ctx analysisContext, variables map[string]arrayVariable, state arrayFlowState, constants map[string]int, capacityGuards []arrayResumeNextCapacityGuard) []Finding {
	var findings []Finding
	seen := map[string]bool{}
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		var issues []Finding
		state, issues = a.arrayTransfer(file, proc, ctx, variables, state, normalizedCodeLine(file.Lines[line-1]), line, constants, capacityGuards)
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
