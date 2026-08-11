package analyze

import (
	"context"
	"fmt"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
)

// CallCycleNode is one declaration in a confirmed procedure cycle. Path is
// closed by repeating its first node after the final procedure.
type CallCycleNode struct {
	QualifiedName string `json:"qualified_name"`
	Module        string `json:"module"`
	Kind          string `json:"kind"`
	ModuleKind    string `json:"module_kind,omitempty"`
	File          string `json:"file"`
	Line          int    `json:"line"`
}

type CallCycleEdge struct {
	Caller    string `json:"caller"`
	Callee    string `json:"callee"`
	File      string `json:"file"`
	Line      int    `json:"line"`
	Column    int    `json:"column,omitempty"`
	EndLine   int    `json:"end_line,omitempty"`
	EndColumn int    `json:"end_column,omitempty"`
}

type CallCycleEffect struct {
	Kind   string `json:"kind"`
	Origin string `json:"origin"`
	File   string `json:"file"`
	Line   int    `json:"line"`
	Target string `json:"target,omitempty"`
}

type CallCycleUncertainty struct {
	Kind   string `json:"kind"`
	Origin string `json:"origin"`
	Callee string `json:"callee,omitempty"`
	File   string `json:"file"`
	Line   int    `json:"line"`
	Column int    `json:"column,omitempty"`
}

// CallCycleContext is the additive JSON contract attached to VBA244.
type CallCycleContext struct {
	Path             []CallCycleNode        `json:"path"`
	Edges            []CallCycleEdge        `json:"edges"`
	CrossModule      bool                   `json:"cross_module"`
	EventHandlers    []string               `json:"event_handlers,omitempty"`
	DangerousEffects []CallCycleEffect      `json:"dangerous_effects,omitempty"`
	Uncertainty      []CallCycleUncertainty `json:"uncertainty,omitempty"`
}

func buildProcedureCallGraphSnapshot(files []parsedFile) callgraph.Snapshot {
	var snapshot callgraph.Snapshot
	for _, file := range files {
		for _, proc := range file.IR.Procedures {
			snapshot.Symbols = append(snapshot.Symbols, callgraph.Symbol{
				Name: proc.Symbol.Name, Kind: string(proc.Symbol.Kind), Module: file.IR.ModuleName,
				ModuleKind: file.IR.ModuleKind, File: file.IR.Path,
				Line: proc.Symbol.DeclarationRange.StartLine, Column: proc.Symbol.DeclarationRange.StartColumn,
				EndLine: proc.Symbol.DeclarationRange.EndLine, EndColumn: proc.Symbol.DeclarationRange.EndColumn,
				Visibility: proc.Symbol.Visibility,
			})
			for _, site := range proc.Calls {
				resolution := calls.Resolution{Status: string(site.Resolution.Status)}
				for _, candidate := range site.Resolution.Candidates {
					resolution.Candidates = append(resolution.Candidates, calls.Candidate{
						QualifiedName: candidate.QualifiedName, Kind: string(candidate.Kind), File: candidate.File, Line: candidate.Line,
					})
				}
				snapshot.Calls = append(snapshot.Calls, calls.Call{
					CallSite: calls.CallSite{
						File: site.File, Module: site.Module,
						Caller: &calls.Caller{Name: site.Caller.Name, Kind: string(site.Caller.Kind), QualifiedName: site.Caller.QualifiedName},
						Range:  site.Range,
					},
					Resolution: resolution,
				})
			}
		}
	}
	return snapshot
}

func (a Analyzer) procedureCallCycleFindings(ctx context.Context, files []parsedFile, project effects.ProjectSummary) ([]Finding, error) {
	cycles, err := callgraph.FindCyclesContext(ctx, buildProcedureCallGraphSnapshot(files))
	if err != nil {
		return nil, err
	}
	if len(cycles) == 0 {
		return nil, nil
	}
	byFile := make(map[string]parsedFile, len(files))
	for _, file := range files {
		byFile[canonicalCyclePath(file.Path)] = file
	}
	bySummary := make(map[string]effects.ProcedureSummary)
	for _, summary := range project.All() {
		identity := summary.Identity
		bySummary[cycleSummaryKey(identity.File, identity.QualifiedName, string(identity.Kind), identity.DeclarationLine)] = summary
	}
	out := make([]Finding, 0, len(cycles))
	for _, cycle := range cycles {
		if len(cycle.Nodes) == 0 || len(cycle.Edges) != len(cycle.Nodes) {
			continue
		}
		cycleContext, severity := buildCallCycleContext(cycle, bySummary)
		anchor := cycle.Edges[0].Location
		file, ok := byFile[canonicalCyclePath(anchor.File)]
		if !ok {
			paths := make([]string, 0, len(byFile))
			for path := range byFile {
				paths = append(paths, path)
			}
			sort.Strings(paths)
			target := strings.ToLower(filepath.ToSlash(anchor.File))
			for _, path := range paths {
				normalized := strings.ToLower(filepath.ToSlash(path))
				if normalized == target || strings.HasSuffix(normalized, "/"+target) {
					file, ok = byFile[path], true
					break
				}
			}
		}
		if !ok {
			continue
		}
		procedures := sourceProceduresFromIR(file.IR, file.CFG)
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		for _, candidate := range procedures {
			if strings.EqualFold(candidate.Name, cycleNodeName(cycle.Nodes[0].QualifiedName)) && candidate.StartLine == cycle.Nodes[0].Line {
				proc = candidate
				break
			}
		}
		pathText := make([]string, 0, len(cycleContext.Path))
		for _, node := range cycleContext.Path {
			pathText = append(pathText, node.QualifiedName)
		}
		message := fmt.Sprintf("Procedure call cycle detected: %s.", strings.Join(pathText, " -> "))
		reason := "Uniquely resolved project-local procedure calls form a recursive cycle."
		if severity == "warning" {
			reason += " The cycle contains a dangerous reachable effect."
		}
		if len(cycleContext.Uncertainty) > 0 {
			reason += fmt.Sprintf(" %d unresolved or dynamically bound call(s) remain uncertain.", len(cycleContext.Uncertainty))
		}
		finding := a.simpleFinding(file, proc, anchor.StartLine, "VBA244", severity, message, reason, "Break the cycle, add an explicit termination guard, or isolate the dangerous effect behind a non-recursive boundary.")
		finding.Column = anchor.StartColumn + 1
		finding.EndLine = anchor.EndLine
		finding.EndColumn = anchor.EndColumn + 1
		finding.ScopeEndLine = proc.EndLine
		finding.CallCycle = &cycleContext
		out = append(out, finding)
	}
	return out, nil
}

func buildCallCycleContext(cycle callgraph.Cycle, summaries map[string]effects.ProcedureSummary) (CallCycleContext, string) {
	ctx := CallCycleContext{Path: make([]CallCycleNode, 0, len(cycle.Nodes)+1), Edges: make([]CallCycleEdge, 0, len(cycle.Edges))}
	modules := map[string]bool{}
	eventSeen := map[string]bool{}
	dangerousSeen := map[string]bool{}
	uncertaintySeen := map[string]bool{}
	dangerous := false
	for _, node := range cycle.Nodes {
		summary, ok := summaryForCycleNode(node, summaries)
		moduleKind := ""
		if ok {
			moduleKind = summary.Identity.ModuleKind
		}
		ctx.Path = append(ctx.Path, CallCycleNode{QualifiedName: node.QualifiedName, Module: node.Module, Kind: node.Kind, ModuleKind: moduleKind, File: node.File, Line: node.Line})
		modules[strings.ToLower(node.Module)] = true
		if strings.EqualFold(node.Kind, "event") || (ok && summary.Identity.IsEventHandler) {
			identity := node.QualifiedName
			identityKey := identity
			if ok {
				identity = summary.Identity.QualifiedName
				identityKey = summary.Identity.Key()
			}
			if !eventSeen[identityKey] {
				eventSeen[identityKey] = true
				ctx.EventHandlers = append(ctx.EventHandlers, identity)
			}
			dangerous = true
		}
		if !ok {
			continue
		}
		processEvidence := func(evidenceItems []effects.Evidence) {
			for _, evidence := range evidenceItems {
				if !dangerousEffect(evidence.Effect) {
					continue
				}
				key := fmt.Sprintf("%s|%s|%d|%d|%s", evidence.Effect, evidence.Origin.Key(), evidence.Range.StartByte, evidence.Range.EndByte, evidence.Target)
				if dangerousSeen[key] {
					continue
				}
				dangerousSeen[key] = true
				ctx.DangerousEffects = append(ctx.DangerousEffects, CallCycleEffect{Kind: string(evidence.Effect), Origin: evidence.Origin.QualifiedName, File: evidence.Origin.File, Line: evidence.Range.StartLine, Target: evidence.Target})
				dangerous = true
			}
		}
		processEvidence(summary.Direct)
		processEvidence(summary.Propagated)
		processUncertainty := func(uncertaintyItems []effects.CallUncertainty) {
			for _, uncertainty := range uncertaintyItems {
				key := fmt.Sprintf("%s|%s|%d|%d|%s", uncertainty.Kind, uncertainty.Origin.Key(), uncertainty.Range.StartByte, uncertainty.Range.EndByte, uncertainty.Callee)
				if uncertaintySeen[key] {
					continue
				}
				uncertaintySeen[key] = true
				ctx.Uncertainty = append(ctx.Uncertainty, CallCycleUncertainty{Kind: string(uncertainty.Kind), Origin: uncertainty.Origin.QualifiedName, Callee: uncertainty.Callee, File: uncertainty.Origin.File, Line: uncertainty.Range.StartLine, Column: uncertainty.Range.StartColumn + 1})
			}
		}
		processUncertainty(summary.DirectUncertainty)
		processUncertainty(summary.PropagatedUncertainty)
	}
	ctx.CrossModule = len(modules) > 1
	if len(ctx.Path) > 0 {
		ctx.Path = append(ctx.Path, ctx.Path[0])
	}
	for _, edge := range cycle.Edges {
		ctx.Edges = append(ctx.Edges, CallCycleEdge{Caller: edge.Caller.QualifiedName, Callee: edge.Callee.QualifiedName, File: edge.Location.File, Line: edge.Location.StartLine, Column: edge.Location.StartColumn + 1, EndLine: edge.Location.EndLine, EndColumn: edge.Location.EndColumn + 1})
	}
	sort.Slice(ctx.EventHandlers, func(i, j int) bool {
		return strings.ToLower(ctx.EventHandlers[i]) < strings.ToLower(ctx.EventHandlers[j])
	})
	sort.Slice(ctx.DangerousEffects, func(i, j int) bool {
		return callCycleEffectKey(ctx.DangerousEffects[i]) < callCycleEffectKey(ctx.DangerousEffects[j])
	})
	sort.Slice(ctx.Uncertainty, func(i, j int) bool {
		return callCycleUncertaintyKey(ctx.Uncertainty[i]) < callCycleUncertaintyKey(ctx.Uncertainty[j])
	})
	if dangerous {
		return ctx, "warning"
	}
	return ctx, "information"
}

func summaryForCycleNode(node callgraph.ID, summaries map[string]effects.ProcedureSummary) (effects.ProcedureSummary, bool) {
	summary, ok := summaries[cycleSummaryKey(node.File, node.QualifiedName, node.Kind, node.Line)]
	return summary, ok
}

func cycleSummaryKey(file, qualifiedName, kind string, line int) string {
	return fmt.Sprintf("%s|%s|%s|%d", strings.ToLower(canonicalCyclePath(file)), strings.ToLower(qualifiedName), strings.ToLower(kind), line)
}

func dangerousEffect(kind effects.EffectKind) bool {
	switch kind {
	case effects.ChangesApplicationState, effects.SuppressesErrors, effects.OpensWorkbook, effects.OpensFile:
		return true
	default:
		return false
	}
}

func canonicalCyclePath(path string) string { return filepath.ToSlash(filepath.Clean(path)) }

func cycleNodeName(qualified string) string {
	if index := strings.LastIndex(qualified, "."); index >= 0 {
		return qualified[index+1:]
	}
	return qualified
}

func callCycleEffectKey(effect CallCycleEffect) string {
	return fmt.Sprintf("%s|%s|%s|%d|%s", effect.Kind, effect.Origin, effect.File, effect.Line, effect.Target)
}
func callCycleUncertaintyKey(item CallCycleUncertainty) string {
	return fmt.Sprintf("%s|%s|%s|%d|%d|%s", item.Kind, item.Origin, item.File, item.Line, item.Column, item.Callee)
}
