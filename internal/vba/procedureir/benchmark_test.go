package procedureir

import (
	"fmt"
	"runtime"
	"strconv"
	"strings"
	"testing"
)

// BenchmarkResolutionOverlay compares the allocation cost of the legacy
// materialized resolution path with the read-only overlay path.  The source
// and resolver are built before the timer starts so the measurements cover
// only project-dependent resolution and, for Resolve, the compatibility
// deep-copy work.
func BenchmarkResolutionOverlay(b *testing.B) {
	for _, procedureCount := range []int{500, 1000, 2000} {
		procedureCount := procedureCount
		b.Run(fmt.Sprintf("%d-procedures", procedureCount), func(b *testing.B) {
			document, resolver := resolutionOverlayBenchmarkFixture(b, procedureCount)
			calls, accesses := resolutionOverlayBenchmarkFactCounts(document)

			b.Run("materialized", func(b *testing.B) {
				b.ReportAllocs()
				b.ResetTimer()
				for i := 0; i < b.N; i++ {
					resolved := Resolve(document, resolver)
					runtime.KeepAlive(resolved)
				}
				b.StopTimer()
				b.ReportMetric(float64(len(document.Procedures)), "procedures/op")
				b.ReportMetric(float64(calls), "calls/op")
				b.ReportMetric(float64(accesses), "accesses/op")
			})

			b.Run("overlay", func(b *testing.B) {
				b.ReportAllocs()
				b.ResetTimer()
				for i := 0; i < b.N; i++ {
					resolved := ResolveView(document, resolver)
					runtime.KeepAlive(resolved)
				}
				b.StopTimer()
				b.ReportMetric(float64(len(document.Procedures)), "procedures/op")
				b.ReportMetric(float64(calls), "calls/op")
				b.ReportMetric(float64(accesses), "accesses/op")
			})
		})
	}
}

func resolutionOverlayBenchmarkFixture(tb testing.TB, procedureCount int) (DocumentIR, Resolver) {
	tb.Helper()
	var source strings.Builder
	// Each procedure intentionally contains two calls and several accesses.
	// This keeps the benchmark focused on resolution facts while retaining a
	// large amount of syntax-local IR for the materialized path to clone.
	for procedureIndex := 0; procedureIndex < procedureCount; procedureIndex++ {
		source.WriteString("Public Sub Procedure")
		source.WriteString(strconv.Itoa(procedureIndex))
		source.WriteString("()\n")
		source.WriteString("  Dim localValue As Long\n")
		source.WriteString("  Call Helper(localValue)\n")
		source.WriteString("  Call MissingProcedure(localValue)\n")
		source.WriteString("  SharedValue = SharedValue + ExternalValue\n")
		source.WriteString("  SharedValue = SharedValue + localValue\n")
		source.WriteString("End Sub\n\n")
	}

	document, err := BuildSource(BuildOptions{
		Path:       "ResolutionOverlayBenchmark.bas",
		ModuleName: "ResolutionOverlayBenchmark",
		ModuleKind: "standard",
	}, []byte(source.String()))
	if err != nil {
		tb.Fatalf("build %d-procedure resolution fixture: %v", procedureCount, err)
	}
	if len(document.Procedures) != procedureCount {
		tb.Fatalf("resolution fixture procedures = %d, want %d", len(document.Procedures), procedureCount)
	}

	resolver := NewResolver([]ResolverSymbol{
		{Name: "Helper", Module: "ProjectHelpers", Kind: "procedure", Visibility: "Public", File: "ProjectHelpers.bas", Line: 1},
		{Name: "SharedValue", Module: "ProjectState", Kind: "module_variable", Visibility: "Public", Type: "Long", File: "ProjectState.bas", Line: 1},
		{Name: "ExternalValue", Module: "ProjectState", Kind: "module_variable", Visibility: "Public", Type: "Long", File: "ProjectState.bas", Line: 2},
	})
	return document, resolver
}

func resolutionOverlayBenchmarkFactCounts(document DocumentIR) (calls, accesses int) {
	for _, procedure := range document.Procedures {
		calls += len(procedure.Calls)
		accesses += len(procedure.Accesses)
	}
	return calls, accesses
}
