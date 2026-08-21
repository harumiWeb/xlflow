package callgraph

import (
	"fmt"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/calls"
)

// benchmarkGraphSink keeps the graph produced by each iteration observable so
// the benchmark measures build rather than an optimized-away result.
var benchmarkGraphSink graph

func BenchmarkBuildProcedures(b *testing.B) {
	for _, procedures := range []int{500, 1000, 2000} {
		input := benchmarkSnapshot(procedures, 500, false)
		b.Run(fmt.Sprintf("procedures_%d_calls_500", procedures), func(b *testing.B) {
			benchmarkBuild(b, input)
		})
	}
}

func BenchmarkBuildMatchedCalls(b *testing.B) {
	const procedures = 2000
	for _, matchedCalls := range []int{500, 1000, 2000} {
		input := benchmarkSnapshot(procedures, matchedCalls, false)
		b.Run(fmt.Sprintf("procedures_%d_calls_%d", procedures, matchedCalls), func(b *testing.B) {
			benchmarkBuild(b, input)
		})
	}
}

func BenchmarkBuildSameNameModules(b *testing.B) {
	const procedures = 2000
	const matchedCalls = 2000
	input := benchmarkSnapshot(procedures, matchedCalls, true)
	benchmarkBuild(b, input)
}

func benchmarkBuild(b *testing.B, input Snapshot) {
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		benchmarkGraphSink = build(input)
	}
}

// benchmarkSnapshot builds all fixture data before the benchmark timer starts.
// When sameName is false, each procedure has a distinct name in one large
// module. When it is true, every module contains a procedure named Worker, so
// candidate identity must distinguish same-named procedures across files.
func benchmarkSnapshot(procedures, matchedCalls int, sameName bool) Snapshot {
	if procedures < 1 {
		procedures = 1
	}
	if matchedCalls < 0 {
		matchedCalls = 0
	}
	if matchedCalls > procedures {
		matchedCalls = procedures
	}

	snapshot := Snapshot{
		Symbols: make([]Symbol, procedures),
		Calls:   make([]calls.Call, matchedCalls),
	}
	for i := 0; i < procedures; i++ {
		module, file, name := "Main", "Main.bas", fmt.Sprintf("Procedure%04d", i)
		if sameName {
			module = fmt.Sprintf("Module%04d", i)
			file = module + ".bas"
			name = "Worker"
		}
		snapshot.Symbols[i] = Symbol{
			Name:       name,
			Kind:       "sub",
			Module:     module,
			ModuleKind: "standard",
			File:       file,
			Line:       i + 1,
			Column:     1,
		}
	}
	for i := 0; i < matchedCalls; i++ {
		callerIndex := i
		calleeIndex := (i + 1) % procedures
		callerModule, callerFile, callerName := benchmarkProcedureIdentity(callerIndex, sameName)
		calleeModule, calleeFile, calleeName := benchmarkProcedureIdentity(calleeIndex, sameName)
		snapshot.Calls[i] = calls.Call{
			CallSite: calls.CallSite{
				File:   callerFile,
				Module: callerModule,
				Caller: &calls.Caller{Name: callerName, Kind: "sub", QualifiedName: callerModule + "." + callerName},
			},
			Resolution: calls.Resolution{
				Status: "matched",
				Candidates: []calls.Candidate{{
					QualifiedName: calleeModule + "." + calleeName,
					Kind:          "sub",
					File:          calleeFile,
					Line:          calleeIndex + 1,
				}},
			},
		}
	}
	return snapshot
}

func benchmarkProcedureIdentity(index int, sameName bool) (module, file, name string) {
	module, file, name = "Main", "Main.bas", fmt.Sprintf("Procedure%04d", index)
	if sameName {
		module = fmt.Sprintf("Module%04d", index)
		file = module + ".bas"
		name = "Worker"
	}
	return module, file, name
}
