package effects

import (
	"fmt"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// BenchmarkBuildLargeCallgraphs measures only effects.Build. The IR and CFG
// input is built before the timer starts so parser, resolver, and graph
// construction costs cannot hide fixed-point costs in the result.
//
// The workload matrix intentionally includes the graph shapes that tend to
// stress different parts of summary propagation. Keep the sizes here aligned
// with the large-callgraph performance contract; individual sub-benchmarks can
// still be selected with -bench when a full matrix is too expensive locally.
func BenchmarkBuildLargeCallgraphs(b *testing.B) {
	for _, workload := range effectBenchmarkWorkloads() {
		workload := workload
		b.Run(workload.name(), func(b *testing.B) {
			documents := buildEffectBenchmarkDocuments(workload.shape, workload.procedures)
			b.ReportAllocs()
			b.ReportMetric(float64(workload.procedures), "procedures/op")
			b.ResetTimer()
			var stats BuildStats
			for i := 0; i < b.N; i++ {
				project := Build(documents)
				stats = project.Stats()
			}
			b.StopTimer()
			b.ReportMetric(float64(stats.WorklistEvaluations), "worklist_evaluations/op")
			b.ReportMetric(float64(stats.MaxPropagatedFactsPerProcedure), "max_propagated_facts/procedure")
			b.ReportMetric(float64(stats.TotalPropagatedFacts), "propagated_facts/op")
			b.ReportMetric(float64(stats.ReusedDirectProcedures), "reused_direct_procedures/op")
			b.ReportMetric(float64(stats.RecomputedDirectProcedures), "recomputed_direct_procedures/op")
		})
	}
}

type effectBenchmarkShape string

const (
	effectBenchmarkChain       effectBenchmarkShape = "chain"
	effectBenchmarkFanIn       effectBenchmarkShape = "fan-in"
	effectBenchmarkFanOut      effectBenchmarkShape = "fan-out"
	effectBenchmarkDense       effectBenchmarkShape = "dense"
	effectBenchmarkUncertainty effectBenchmarkShape = "uncertainty"
	effectBenchmarkEffectHeavy effectBenchmarkShape = "effect-heavy"
)

type effectBenchmarkWorkload struct {
	shape      effectBenchmarkShape
	procedures int
}

func effectBenchmarkWorkloads() []effectBenchmarkWorkload {
	shapes := []effectBenchmarkShape{
		effectBenchmarkChain,
		effectBenchmarkFanIn,
		effectBenchmarkFanOut,
		effectBenchmarkDense,
		effectBenchmarkUncertainty,
		effectBenchmarkEffectHeavy,
	}
	workloads := make([]effectBenchmarkWorkload, 0, len(shapes)*3)
	for _, shape := range shapes {
		for _, procedures := range []int{500, 1000, 2000} {
			workloads = append(workloads, effectBenchmarkWorkload{shape: shape, procedures: procedures})
		}
	}
	return workloads
}

func (w effectBenchmarkWorkload) name() string {
	return fmt.Sprintf("%s/%d-procedures", w.shape, w.procedures)
}

// buildEffectBenchmarkDocuments constructs a single project document with
// already-resolved calls and a prebuilt CFG. Graph construction is deliberately
// outside the benchmark timer so effects.Build includes reachable-statement
// extraction and its block lookup, but not parser or CFG construction costs.
func buildEffectBenchmarkDocuments(shape effectBenchmarkShape, procedures int) []Document {
	if procedures < 1 {
		procedures = 1
	}
	const (
		path   = "effects-benchmark.bas"
		module = "EffectsBenchmark"
	)

	all := make([]procedureir.ProcedureIR, procedures)
	for index := range all {
		line := index + 1
		name := fmt.Sprintf("Proc%04d", index)
		proc := procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{
			Name: name, QualifiedName: module + "." + name,
			Kind: procedureir.ProcedureSub, Visibility: "Private",
			DeclarationRange: vbaast.Range{StartLine: line, EndLine: line},
		}}
		appendBenchmarkDirectFacts(&proc, shape, index, procedures)
		appendBenchmarkCalls(&proc, shape, index, procedures, path, module)
		all[index] = proc
	}

	ir := procedureir.DocumentIR{
		Path: path, ModuleName: module, ModuleKind: "standard", Procedures: all,
	}
	return []Document{{IR: ir, CFG: cfg.BuildDocument(ir)}}
}

func appendBenchmarkDirectFacts(proc *procedureir.ProcedureIR, shape effectBenchmarkShape, index, procedures int) {
	// Effect-heavy procedures provide several distinct direct facts. The
	// targets are unique per procedure so the benchmark retains provenance
	// pressure while remaining deterministic.
	factCount := 0
	switch shape {
	case effectBenchmarkEffectHeavy:
		if index == procedures-1 {
			factCount = 8
		}
	case effectBenchmarkChain, effectBenchmarkDense:
		// A single terminal origin is enough to exercise long transitive paths.
		if index == procedures-1 {
			// Keep the origin at the end of the call direction so callers
			// repeatedly receive it during fixed-point propagation.
			factCount = 1
		}
	case effectBenchmarkFanIn:
		if index == 0 {
			factCount = 1
		}
	case effectBenchmarkFanOut:
		if index == procedures-1 {
			factCount = 1
		}
	}
	for fact := 0; fact < factCount; fact++ {
		statementID := len(proc.Statements) + 1
		target := fmt.Sprintf("Range(\"A%d\").Value", index+1)
		value := fmt.Sprintf("%d", fact)
		if shape == effectBenchmarkEffectHeavy {
			stateTargets := []string{
				"Application.EnableEvents", "Application.Calculation", "Application.DisplayAlerts",
				"Application.ScreenUpdating", "Application.StatusBar", "Application.Interactive",
				"Application.AskToUpdateLinks", "Application.CutCopyMode",
			}
			target = stateTargets[fact]
			value = "False"
			if fact == 1 {
				value = "xlCalculationManual"
			}
		}
		proc.Statements = append(proc.Statements, procedureir.Statement{
			ID: statementID, Kind: procedureir.StatementAssignment,
			Text:   target + " = " + value,
			Range:  vbaast.Range{StartLine: index + 1, EndLine: index + 1},
			Target: &procedureir.Expression{Text: target},
			Value:  &procedureir.Expression{Text: value},
		})
	}
}

func appendBenchmarkCalls(proc *procedureir.ProcedureIR, shape effectBenchmarkShape, index, procedures int, path, module string) {
	const (
		fanout        = 8
		uncertainties = 16
	)

	appendMatched := func(target int) {
		if target < 0 || target >= procedures || target == index {
			return
		}
		statementID := len(proc.Statements) + 1
		name := fmt.Sprintf("Proc%04d", target)
		proc.Statements = append(proc.Statements, procedureir.Statement{
			ID: statementID, Kind: procedureir.StatementCall,
			Text: name, Range: vbaast.Range{StartLine: index + 1, EndLine: index + 1},
		})
		proc.Calls = append(proc.Calls, procedureir.CallSite{
			ID: statementID, StatementID: statementID,
			Callee: procedureir.Callee{Text: name, BaseName: name},
			Range:  vbaast.Range{StartLine: index + 1, EndLine: index + 1},
			Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{
				QualifiedName: module + "." + name, Kind: string(procedureir.ProcedureSub), File: path, Line: target + 1,
			}}},
		})
	}

	switch shape {
	case effectBenchmarkChain:
		appendMatched(index + 1)
	case effectBenchmarkFanIn:
		if index > 0 {
			appendMatched(0)
		}
	case effectBenchmarkFanOut:
		for offset := 1; offset <= fanout; offset++ {
			appendMatched(index + offset)
		}
	case effectBenchmarkDense:
		for offset := 1; offset <= fanout; offset++ {
			appendMatched(index + offset)
		}
		if index+fanout+1 < procedures {
			appendMatched(index + fanout + 1)
		}
	case effectBenchmarkUncertainty:
		appendMatched(index + 1)
		for uncertainty := 0; uncertainty < uncertainties; uncertainty++ {
			statementID := len(proc.Statements) + 1
			callee := fmt.Sprintf("Dynamic%04d_%02d", index, uncertainty)
			proc.Statements = append(proc.Statements, procedureir.Statement{
				ID: statementID, Kind: procedureir.StatementCall,
				Text: callee, Range: vbaast.Range{StartLine: index + 1, EndLine: index + 1},
			})
			proc.Calls = append(proc.Calls, procedureir.CallSite{
				ID: statementID, StatementID: statementID,
				Callee:     procedureir.Callee{Text: callee, BaseName: callee},
				Range:      vbaast.Range{StartLine: index + 1, EndLine: index + 1},
				Resolution: procedureir.CallResolution{Status: procedureir.ResolutionUnresolved},
			})
		}
	case effectBenchmarkEffectHeavy:
		appendMatched(index + 1)
	}
}

func TestEffectBenchmarkWorkloadScales(t *testing.T) {
	for _, workload := range effectBenchmarkWorkloads() {
		t.Run(workload.name(), func(t *testing.T) {
			documents := buildEffectBenchmarkDocuments(workload.shape, workload.procedures)
			if len(documents) != 1 || len(documents[0].IR.Procedures) != workload.procedures {
				t.Fatalf("workload has %d documents and %d procedures, want 1 and %d", len(documents), len(documents[0].IR.Procedures), workload.procedures)
			}
			for _, proc := range documents[0].IR.Procedures {
				if proc.Symbol.QualifiedName == "" {
					t.Fatal("benchmark procedure has no qualified name")
				}
				for _, call := range proc.Calls {
					if call.StatementID == 0 {
						t.Fatalf("%s procedure %s has a call without statement ID", workload.shape, proc.Symbol.QualifiedName)
					}
				}
			}
		})
	}
}
