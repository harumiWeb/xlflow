package procedureir

import (
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestBuildProjectResolverProjectsDocumentAndExternalSymbols(t *testing.T) {
	branch := ConditionalBranch{Condition: "VBA7", Branch: 0}
	documents := []DocumentIR{{
		Path:       "src/Project.bas",
		ModuleName: "  ProjectModule  ",
		ModuleKind: "standard",
		Declarations: []Declaration{{
			Name: "ProjectValue", Type: "Long", Visibility: "Private", Kind: "const",
			Parent: "Settings", Recovered: true, IsArray: true, IsConst: true,
			ValueShape: ValueShapeScalar, Range: ast.Range{StartLine: 12},
			ConditionalBranches: []ConditionalBranch{branch},
		}},
		Procedures: []ProcedureIR{{
			Symbol: ProcedureSymbol{
				Name: "ProjectFunction", Kind: ProcedureFunction, Visibility: "Public", ReturnType: "String",
				IsArray: true, ValueShape: ValueShapeFixedArray, Recovered: true,
				DeclarationRange:    ast.Range{StartLine: 24},
				ConditionalBranches: []ConditionalBranch{branch},
			},
		}},
	}}
	externalSymbols := []ResolverSymbol{{
		Name: "vbTextCompare", Module: "VBA", ModuleKind: "external", Kind: "enum_member",
		Visibility: "Public", Parent: "VbCompareMethod", File: "<typelib>VBA", IsConst: true,
		ValueShape: ValueShapeScalar,
	}}

	wantSymbols := []ResolverSymbol{
		{
			Name: "ProjectValue", Type: "Long", Module: "ProjectModule", ModuleKind: "standard",
			Kind: "const", Visibility: "Private", File: "src/Project.bas", Line: 12, Parent: "Settings",
			Recovered: true, IsArray: true, IsConst: true, ValueShape: ValueShapeScalar,
			ConditionalBranches: []ConditionalBranch{branch},
		},
		{
			Name: "ProjectFunction", Type: "String", Module: "ProjectModule", ModuleKind: "standard",
			Kind: "function", Visibility: "Public", File: "src/Project.bas", Line: 24,
			Recovered: true, IsArray: true, ValueShape: ValueShapeFixedArray,
			ConditionalBranches: []ConditionalBranch{branch},
		},
		externalSymbols[0],
	}

	got := BuildProjectResolver(documents, externalSymbols, false)
	want := NewResolverWithCompleteness(wantSymbols, false)
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("BuildProjectResolver() resolver differs from expected projection:\n got: %#v\nwant: %#v", got, want)
	}
	if gotResolver, ok := got.(SymbolResolver); !ok || gotResolver.complete {
		t.Fatalf("BuildProjectResolver() completeness = %#v, want incomplete resolver", got)
	}
}

func TestBuildProjectResolverDoesNotMutateInputs(t *testing.T) {
	branch := ConditionalBranch{Condition: "VBA7"}
	documents := []DocumentIR{{
		Path: "module.bas", ModuleName: "Module", ModuleKind: "standard",
		Declarations: []Declaration{{Name: "Value", ConditionalBranches: []ConditionalBranch{branch}}},
		Procedures:   []ProcedureIR{{Symbol: ProcedureSymbol{Name: "Run", ConditionalBranches: []ConditionalBranch{branch}}}},
	}}
	externalSymbols := []ResolverSymbol{{
		Name: "ExternalValue", Kind: "enum_member", ConditionalBranches: []ConditionalBranch{branch},
	}}
	documentsBefore := cloneProjectResolverTestDocuments(documents)
	externalBefore := append([]ResolverSymbol(nil), externalSymbols...)
	externalBefore[0].ConditionalBranches = append([]ConditionalBranch(nil), externalSymbols[0].ConditionalBranches...)

	_ = BuildProjectResolver(documents, externalSymbols, true)

	if !reflect.DeepEqual(documents, documentsBefore) {
		t.Fatal("BuildProjectResolver() mutated DocumentIR input")
	}
	if !reflect.DeepEqual(externalSymbols, externalBefore) {
		t.Fatal("BuildProjectResolver() mutated external symbols")
	}
}

func cloneProjectResolverTestDocuments(documents []DocumentIR) []DocumentIR {
	cloned := make([]DocumentIR, len(documents))
	for i, document := range documents {
		cloned[i] = document
		cloned[i].Declarations = append([]Declaration(nil), document.Declarations...)
		for j := range cloned[i].Declarations {
			cloned[i].Declarations[j].ConditionalBranches = append([]ConditionalBranch(nil), document.Declarations[j].ConditionalBranches...)
		}
		cloned[i].Procedures = append([]ProcedureIR(nil), document.Procedures...)
		for j := range cloned[i].Procedures {
			cloned[i].Procedures[j].Symbol.ConditionalBranches = append([]ConditionalBranch(nil), document.Procedures[j].Symbol.ConditionalBranches...)
		}
	}
	return cloned
}
