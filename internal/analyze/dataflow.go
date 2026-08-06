package analyze

import (
	"fmt"
	"strings"

	vbadf "github.com/harumiWeb/xlflow/internal/vba/dataflow"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// dataFlowFindings is the single adapter used by batch and realtime analysis.
// The dataflow package remains independent from analyzer and LSP protocols.
func (a Analyzer) dataFlowFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectUntrustedDataFlow || proc.Graph == nil {
		return nil
	}
	ir, ok := procedureIRForSource(file.IR, proc)
	if !ok {
		return nil
	}
	result := vbadf.AnalyzeProcedure(ir, *proc.Graph, vbadf.Options{Conservative: true})
	findings := make([]Finding, 0, len(result.Findings))
	for _, flow := range result.Findings {
		line := flow.Sink.Range.StartLine
		if line <= 0 {
			line = flow.Source.Range.StartLine
		}
		if line <= 0 {
			line = proc.StartLine
		}
		sourceLine := flow.Source.Range.StartLine
		if sourceLine <= 0 {
			sourceLine = line
		}
		sinkLine := flow.Sink.Range.StartLine
		if sinkLine <= 0 {
			sinkLine = line
		}
		path := formatDataFlowPath(flow.Path)
		message := fmt.Sprintf("Conservative analysis: %s flows to %s. Source: %s; sink: %s; path: %s.", flow.Source.Label, flow.Sink.Label, flow.Source.Label, flow.Sink.Label, path)
		reason := fmt.Sprintf("Conservative analysis keeps potentially untrusted data from %s through the propagation path %s until it reaches the sensitive sink %s.", flow.Source.Label, path, flow.Sink.Label)
		suggestion := "Validate against a narrow allowlist or use the sink-specific sanitizer before calling the sensitive API."
		finding := a.simpleFinding(file, proc, line, "VBA224", "warning", message, reason, suggestion)
		finding.DataFlow = &DataFlowContext{
			Source: DataFlowEndpoint{Kind: string(flow.Source.Kind), Label: flow.Source.Label, Line: sourceLine},
			Sink:   DataFlowEndpoint{Kind: string(flow.Sink.Kind), Label: flow.Sink.Label, Line: sinkLine},
			Path:   convertDataFlowPath(flow.Path, line),
		}
		findings = append(findings, finding)
	}
	return findings
}

func procedureIRForSource(document procedureir.DocumentIR, proc sourceProcedure) (procedureir.ProcedureIR, bool) {
	for _, candidate := range document.Procedures {
		if !strings.EqualFold(candidate.Symbol.Name, proc.Name) {
			continue
		}
		if proc.StartLine > 0 && candidate.Symbol.DeclarationRange.StartLine != proc.StartLine {
			continue
		}
		return candidate, true
	}
	return procedureir.ProcedureIR{}, false
}

func convertDataFlowPath(path []vbadf.PathStep, fallbackLine int) []DataFlowStep {
	steps := make([]DataFlowStep, 0, len(path))
	for _, step := range path {
		line := step.Range.StartLine
		if line <= 0 {
			line = fallbackLine
		}
		steps = append(steps, DataFlowStep{Kind: step.Kind, Label: step.Label, Line: line})
	}
	return steps
}

func formatDataFlowPath(path []vbadf.PathStep) string {
	if len(path) == 0 {
		return "source to sink"
	}
	labels := make([]string, 0, len(path))
	for _, step := range path {
		label := strings.TrimSpace(step.Label)
		if label == "" {
			label = step.Kind
		}
		labels = append(labels, label)
	}
	return strings.Join(labels, " -> ")
}
