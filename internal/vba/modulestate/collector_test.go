package modulestate

import (
	"context"
	"errors"
	"path/filepath"
	"reflect"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestCollectIndexesFieldAccessAndCollectionMutators(t *testing.T) {
	path := filepath.Join("src", "Main.bas")
	items := "items"
	document := procedureir.DocumentIR{
		Path: path, ModuleName: "Main", ModuleKind: "standard",
		Declarations: []procedureir.Declaration{
			{Name: "count", Type: "Long", Scope: procedureir.ScopeProject, Kind: "variable", Range: procedureirRange(2)},
			{Name: "items", Type: "Scripting.Dictionary", Scope: procedureir.ScopeProject, Kind: "variable", IsObject: true, Range: procedureirRange(3)},
			{Name: "version", Type: "Long", Scope: procedureir.ScopeProject, Kind: "const", IsConst: true, Range: procedureirRange(4)},
		},
		Procedures: []procedureir.ProcedureIR{
			{
				Symbol:   procedureSymbol("Main", "ReadCount", 10),
				Accesses: []procedureir.VariableAccess{{Name: "count", Mode: procedureir.AccessRead, Scope: procedureir.ScopeProject}},
			},
			{
				Symbol:   procedureSymbol("Main", "WriteCount", 20),
				Accesses: []procedureir.VariableAccess{{Name: "count", Mode: procedureir.AccessWrite, Scope: procedureir.ScopeProject}},
			},
			{
				Symbol: procedureSymbol("Main", "Mutate", 30),
				Calls:  []procedureir.CallSite{{Callee: procedureir.Callee{Text: "items.Add", Receiver: &items, Member: "Add"}}},
			},
			{
				Symbol:     procedureSymbol("Main", "ReadItem", 40),
				Statements: []procedureir.Statement{{ID: 1, Kind: procedureir.StatementCall}},
				Calls:      []procedureir.CallSite{{StatementID: 1, ExpressionID: 2, Callee: procedureir.Callee{Text: "items.Item", Receiver: &items, Member: "Item"}}},
			},
			{
				Symbol:      procedureSymbol("Main", "SetItem", 50),
				Statements:  []procedureir.Statement{{ID: 1, Kind: procedureir.StatementSet, TargetID: 1}},
				Expressions: []procedureir.Expression{{ID: 1}, {ID: 2, ParentID: 1}},
				Calls:       []procedureir.CallSite{{StatementID: 1, ExpressionID: 2, Callee: procedureir.Callee{Text: "items.Item", Receiver: &items, Member: "Item"}}},
			},
		},
	}

	before := procedureir.DocumentIR{
		Path: document.Path, ModuleName: document.ModuleName, ModuleKind: document.ModuleKind,
		Declarations: append([]procedureir.Declaration(nil), document.Declarations...),
		Procedures:   append([]procedureir.ProcedureIR(nil), document.Procedures...),
	}
	facts := Collect([]procedureir.DocumentIR{document})
	if !reflect.DeepEqual(document, before) {
		t.Fatal("Collect mutated its DocumentIR input")
	}

	fieldByName := make(map[string]Field, len(facts.Fields))
	for _, field := range facts.Fields {
		fieldByName[field.Name] = field
	}
	count := fieldByName["count"]
	if !reflect.DeepEqual(count.Readers, []string{procedureID(path, "Main", procedureSymbol("Main", "ReadCount", 10))}) {
		t.Fatalf("count readers = %v", count.Readers)
	}
	if !reflect.DeepEqual(count.Writers, []string{procedureID(path, "Main", procedureSymbol("Main", "WriteCount", 20))}) {
		t.Fatalf("count writers = %v", count.Writers)
	}
	if field := fieldByName["version"]; field.Kind != "const" || field.IsCollection {
		t.Fatalf("constant field facts = %+v", field)
	}
	collection := fieldByName["items"]
	if !collection.IsCollection || !collection.IsObject {
		t.Fatalf("dictionary field classification = %+v", collection)
	}
	if len(collection.Mutators) != 2 || len(collection.Writers) != 2 {
		t.Fatalf("collection writer/mutator IDs = %v / %v", collection.Writers, collection.Mutators)
	}
	for _, procedure := range facts.Procedures {
		if procedure.Name == "ReadItem" && len(procedure.Mutators) != 0 {
			t.Fatalf("Item read was classified as mutation: %+v", procedure)
		}
		if procedure.Name == "SetItem" && !reflect.DeepEqual(procedure.Mutators, []string{collection.ID}) {
			t.Fatalf("Item assignment mutators = %v", procedure.Mutators)
		}
	}
}

func TestCollectRecognizesLateBoundInitializationBeforeMutatorsAndHonorsShadowing(t *testing.T) {
	path := filepath.Join("src", "LateBound.bas")
	items := "items"
	document := procedureir.DocumentIR{
		Path: path, ModuleName: "LateBound", ModuleKind: "standard",
		Declarations: []procedureir.Declaration{{Name: "items", Type: "Object", Scope: procedureir.ScopeProject, Kind: "variable", IsObject: true, Range: procedureirRange(2)}},
		Procedures: []procedureir.ProcedureIR{
			{
				Symbol: procedureSymbol("LateBound", "MutateFirst", 10),
				Calls:  []procedureir.CallSite{{Callee: procedureir.Callee{Text: "items.Add", Receiver: &items, Member: "Add"}}},
			},
			{
				Symbol:     procedureSymbol("LateBound", "InitializeLater", 20),
				Statements: []procedureir.Statement{{ID: 1, Kind: procedureir.StatementSet, Text: "Set items = New Collection"}},
				Accesses:   []procedureir.VariableAccess{{Name: "items", Mode: procedureir.AccessWrite, Scope: procedureir.ScopeProject, StatementID: 1}},
			},
			{
				Symbol: procedureSymbol("LateBound", "Shadowed", 30),
				Calls:  []procedureir.CallSite{{Callee: procedureir.Callee{Text: "items.Add", Receiver: &items, Member: "Add"}}},
			},
		},
	}
	document.Procedures[2].Symbol.Parameters = []procedureir.Parameter{{Name: "items"}}

	facts := Collect([]procedureir.DocumentIR{document})
	if len(facts.Fields) != 1 {
		t.Fatalf("field count = %d", len(facts.Fields))
	}
	field := facts.Fields[0]
	if !field.IsCollection {
		t.Fatalf("late-bound initialized field was not recognized: %+v", field)
	}
	mutatorID := findProcedure(t, facts, "LateBound.MutateFirst").ID
	if len(field.Mutators) != 1 || field.Mutators[0] != mutatorID {
		t.Fatalf("late-bound mutators = %v, procedures = %+v", field.Mutators, facts.Procedures)
	}
	if len(field.Writers) != 2 {
		t.Fatalf("late-bound writers = %v, want initializer and recognized mutator", field.Writers)
	}
	if len(facts.Procedures[2].Mutators) != 0 {
		t.Fatalf("local parameter shadow was treated as module mutation: %+v", facts.Procedures[2])
	}
}

func TestCollectResolvesProjectPathsAndOnlyUniqueLocalCalls(t *testing.T) {
	root := t.TempDir()
	sharedPath := filepath.Join("src", "Shared.bas")
	callerPath := filepath.Join("src", "Caller.bas")
	shared := procedureir.DocumentIR{
		Path: sharedPath, ModuleName: "Shared", ModuleKind: "standard",
		Declarations: []procedureir.Declaration{{Name: "state", Type: "Long", Scope: procedureir.ScopeProject, Kind: "variable", Range: procedureirRange(2)}},
		Procedures:   []procedureir.ProcedureIR{{Symbol: procedureSymbol("Shared", "Work", 10)}},
	}
	caller := procedureir.DocumentIR{
		Path: callerPath, ModuleName: "Caller", ModuleKind: "class",
		Procedures: []procedureir.ProcedureIR{
			{
				Symbol: procedureSymbol("Caller", "Run", 10),
				Accesses: []procedureir.VariableAccess{{
					Name: "state", Mode: procedureir.AccessRead, Scope: procedureir.ScopeProject,
					Resolution: procedureir.SymbolResolution{Scope: procedureir.ScopeProject, Candidates: []procedureir.Candidate{{File: sharedPath, QualifiedName: "Shared.state", Kind: "variable", Line: 2}}},
				}},
				Calls: []procedureir.CallSite{
					{Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{File: sharedPath, QualifiedName: "Shared.Work", Kind: string(procedureir.ProcedureSub), Line: 10}}}},
					{Resolution: procedureir.CallResolution{Status: procedureir.ResolutionAmbiguous, Candidates: []procedureir.Candidate{{File: sharedPath, QualifiedName: "Shared.Work", Kind: string(procedureir.ProcedureSub), Line: 10}, {File: callerPath, QualifiedName: "Caller.Work", Kind: string(procedureir.ProcedureSub), Line: 20}}}},
				},
			},
		},
	}

	facts := CollectWithOptions([]procedureir.DocumentIR{shared, caller}, Options{RootDir: root})
	if !reflect.DeepEqual(facts.Fields[0].Readers, []string{procedureID(callerPath, "Caller", caller.Procedures[0].Symbol)}) {
		t.Fatalf("resolved cross-module readers = %v", facts.Fields[0].Readers)
	}
	callerFacts := findProcedure(t, facts, "Caller.Run")
	if !reflect.DeepEqual(callerFacts.Callees, []string{procedureID(sharedPath, "Shared", shared.Procedures[0].Symbol)}) {
		t.Fatalf("resolved callees = %v", callerFacts.Callees)
	}
}

func TestCollectContextReturnsCancellationWithoutPartialFacts(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	facts, err := CollectContext(ctx, []procedureir.DocumentIR{{Path: "Main.bas"}}, Options{})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("CollectContext error = %v, want context.Canceled", err)
	}
	if len(facts.Fields) != 0 || len(facts.Procedures) != 0 {
		t.Fatalf("canceled collection returned partial facts: %+v", facts)
	}
}

func TestCollectPreservesRelativeCandidateIdentity(t *testing.T) {
	root := ".analyze-relative-vba240"
	alphaPath := filepath.Join(root, "src", "modules", "Alpha.bas")
	zuluPath := filepath.Join(root, "src", "modules", "Zulu.bas")
	declaration := procedureir.Declaration{Name: "sharedItems", Scope: procedureir.ScopeProject, Kind: "variable", Range: procedureirRange(2)}
	caller := procedureir.DocumentIR{Path: "Main.bas", ModuleName: "Main", ModuleKind: "standard", Procedures: []procedureir.ProcedureIR{{
		Symbol: procedureSymbol("Main", "Run", 1),
		Accesses: []procedureir.VariableAccess{{Name: "sharedItems", Mode: procedureir.AccessRead, Scope: procedureir.ScopeProject,
			Resolution: procedureir.SymbolResolution{Scope: procedureir.ScopeProject, Candidates: []procedureir.Candidate{{File: zuluPath, Line: 2}}}}},
	}}}
	facts := CollectWithOptions([]procedureir.DocumentIR{
		{Path: alphaPath, ModuleName: "Alpha", ModuleKind: "standard", Declarations: []procedureir.Declaration{declaration}},
		{Path: zuluPath, ModuleName: "Zulu", ModuleKind: "standard", Declarations: []procedureir.Declaration{declaration}},
		caller,
	}, Options{RootDir: root})
	for _, field := range facts.Fields {
		if (field.Module == "Zulu" && len(field.Readers) != 1) || (field.Module == "Alpha" && len(field.Readers) != 0) {
			t.Fatalf("relative candidate selected the wrong declaration: %+v", facts.Fields)
		}
	}
}

func findProcedure(t *testing.T, facts Facts, qualified string) Procedure {
	t.Helper()
	for _, procedure := range facts.Procedures {
		if procedure.QualifiedName == qualified {
			return procedure
		}
	}
	t.Fatalf("procedure %q not found in facts: %+v", qualified, facts.Procedures)
	return Procedure{}
}

func procedureSymbol(module, name string, line int) procedureir.ProcedureSymbol {
	return procedureir.ProcedureSymbol{
		Name: name, QualifiedName: module + "." + name,
		Kind: procedureir.ProcedureSub, DeclarationRange: procedureirRange(line),
	}
}

func procedureirRange(line int) vbaast.Range {
	return vbaast.Range{StartLine: line, EndLine: line}
}
