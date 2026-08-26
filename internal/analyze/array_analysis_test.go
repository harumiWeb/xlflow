package analyze

import (
	"sync"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// ArrayAnalysisResult is handed to independent diagnostic projectors after
// materialization. This test exercises the read-only contract and, when run
// with -race, catches accidental lazy mutation of projection-owned slices.
func TestArrayAnalysisResultConcurrentRead(t *testing.T) {
	result := &ArrayAnalysisResult{
		variables: map[string]arrayVariable{
			"values": {name: "values", isArray: true},
		},
		lifecycleFindings: []Finding{{Code: "VBA227", Message: "array", NearbyCode: []string{"original nearby"}}},
		runtimeFindings:   []Finding{{Code: "VBA249", Message: "runtime", RuntimeError: &RuntimeErrorContext{Kind: "original runtime"}}},
		redimFindings:     []Finding{{Code: "VBA208", Message: "redim"}},
		rangeFindings:     []Finding{{Code: "VBA226", Message: "shape"}},
	}

	const readers = 8
	const iterations = 200
	var wait sync.WaitGroup
	wait.Add(readers)
	for i := 0; i < readers; i++ {
		go func() {
			defer wait.Done()
			for j := 0; j < iterations; j++ {
				for _, finding := range result.lifecycle() {
					if finding.Code != "VBA227" {
						t.Errorf("lifecycle code = %q", finding.Code)
					}
				}
				for _, finding := range result.runtime() {
					if finding.Code != "VBA249" {
						t.Errorf("runtime code = %q", finding.Code)
					}
				}
				if len(result.redim()) != 1 || len(result.rangeShape()) != 1 {
					t.Errorf("projection result was mutated")
				}
			}
		}()
	}
	wait.Wait()

	copyOfFindings := result.lifecycle()
	copyOfFindings[0].Code = "mutated-copy"
	copyOfFindings[0].NearbyCode[0] = "mutated-nearby-copy"
	if result.lifecycle()[0].Code != "VBA227" {
		t.Fatal("projector copy mutated the immutable result")
	}
	if result.lifecycle()[0].NearbyCode[0] != "original nearby" {
		t.Fatal("projector copy mutated nested nearby code in the immutable result")
	}
	runtimeCopy := result.runtime()
	runtimeCopy[0].RuntimeError.Kind = "mutated-runtime-copy"
	if result.runtime()[0].RuntimeError.Kind != "original runtime" {
		t.Fatal("projector copy mutated nested runtime error in the immutable result")
	}
}

func TestArrayLifecycleProjectionApplicableIncludesIndexedAccess(t *testing.T) {
	proc := sourceProcedure{Statements: newReadOnlySpan([]procedureir.Statement{{Text: "Debug.Print values(2)"}})}
	variables := map[string]arrayVariable{"values": {name: "values", isArray: true}}
	if !arrayLifecycleProjectionApplicable(parsedFile{}, proc, variables) {
		t.Fatal("indexed array access should make the VBA227 projection applicable")
	}
}

func TestArrayLifecycleProjectionApplicableIncludesRecoveredSource(t *testing.T) {
	proc := sourceProcedure{StartLine: 1, EndLine: 3}
	file := parsedFile{Lines: []string{"Sub Run()", "  Debug.Print values(2)", "End Sub"}}
	variables := map[string]arrayVariable{"values": {name: "values", isArray: true}}
	if !arrayLifecycleProjectionApplicable(file, proc, variables) {
		t.Fatal("recovered source indexed access should make the VBA227 projection applicable")
	}
}

func TestArrayVariablesFiltersCatalogScalarsByProcedureAccess(t *testing.T) {
	moduleDecls := map[string]sourceDeclaration{
		"usedscalar":   {Name: "UsedScalar", Type: "Long"},
		"unusedscalar": {Name: "UnusedScalar", Type: "Long"},
		"values":       {Name: "values", Type: "Long", Array: true},
	}
	file := parsedFile{
		Module:             "Main",
		Lines:              []string{"Sub Run()", "", "End Sub"},
		ModuleDeclarations: moduleDecls,
		ArrayVariableCatalog: map[string]arrayVariable{
			"usedscalar":   {name: "UsedScalar", typ: "Long", knownScalar: true},
			"unusedscalar": {name: "UnusedScalar", typ: "Long", knownScalar: true},
			"values":       {name: "values", typ: "Long", isArray: true},
		},
	}
	proc := sourceProcedure{StartLine: 1, EndLine: 3, Accesses: newReadOnlySpan([]procedureir.VariableAccess{
		{Name: "UsedScalar", Scope: procedureir.ScopeModule, Mode: procedureir.AccessRead},
		{Name: "values", Scope: procedureir.ScopeModule, Mode: procedureir.AccessRead},
	})}

	variables := arrayVariables(file, proc, moduleDecls)
	if _, ok := variables["usedscalar"]; !ok {
		t.Fatal("accessed module scalar was filtered out")
	}
	if _, ok := variables["values"]; !ok {
		t.Fatal("module array was filtered out")
	}
	if _, ok := variables["unusedscalar"]; ok {
		t.Fatal("unaccessed module scalar retained in procedure projection")
	}
}

func TestArrayVariablesFailsOpenWithoutAccessProjection(t *testing.T) {
	moduleDecls := map[string]sourceDeclaration{
		"modulevalue": {Name: "ModuleValue", Type: "Long"},
	}
	file := parsedFile{
		Lines:              []string{"Sub Run()", "", "End Sub"},
		ModuleDeclarations: moduleDecls,
		ArrayVariableCatalog: map[string]arrayVariable{
			"modulevalue": {name: "ModuleValue", typ: "Long", knownScalar: true},
		},
	}
	variables := arrayVariables(file, sourceProcedure{StartLine: 1, EndLine: 3}, moduleDecls)
	if _, ok := variables["modulevalue"]; !ok {
		t.Fatal("module declaration must be retained when access projection is unavailable")
	}
}
