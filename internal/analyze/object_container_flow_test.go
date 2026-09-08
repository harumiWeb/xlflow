package analyze

import (
	"strings"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestObjectContainerProcedureCalledBeforeObservationRequiresResolvedCall(t *testing.T) {
	t.Parallel()
	graph := cfg.Build(procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{
			Name:          "Run",
			QualifiedName: "M.Run",
			Kind:          procedureir.ProcedureSub,
		},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementCall, Text: "Call Helper"},
			{ID: 2, Kind: procedureir.StatementCall, Text: "Debug.Print 1"},
		},
	})
	owner := sourceProcedure{
		Module:     "M",
		Name:       "Run",
		StartLine:  1,
		Graph:      &graph,
		Statements: newReadOnlySpan([]procedureir.Statement{{ID: 1, Kind: procedureir.StatementCall, Text: "Call Helper"}, {ID: 2, Kind: procedureir.StatementCall, Text: "Debug.Print 1"}}),
	}
	target := sourceProcedure{Module: "M", Name: "Helper", StartLine: 10}
	index := &objectContainerIndex{
		file:        parsedFile{ModuleDeclarations: map[string]sourceDeclaration{}},
		procedures:  []sourceProcedure{owner, target},
		moduleDecls: map[string]sourceDeclaration{},
	}

	t.Run("non-callable shadow", func(t *testing.T) {
		call := procedureir.CallSite{
			StatementID: 1,
			Callee:      procedureir.Callee{BaseName: "Helper"},
			Resolution:  procedureir.CallResolution{Status: procedureir.ResolutionNonCallable},
		}
		owner.Calls = newReadOnlySpan([]procedureir.CallSite{call})
		index.procedures[0] = owner
		if objectContainerProcedureCalledBeforeObservation(index, owner, target, 2) {
			t.Fatal("a non-callable shadowed helper must not establish module initialization")
		}
	})

	t.Run("unique resolved call", func(t *testing.T) {
		call := procedureir.CallSite{
			StatementID: 1,
			Callee:      procedureir.Callee{BaseName: "Helper"},
			Resolution: procedureir.CallResolution{
				Status:     procedureir.ResolutionMatched,
				Candidates: []procedureir.Candidate{{QualifiedName: "M.Helper", Kind: string(procedureir.ProcedureSub), Line: 10}},
			},
		}
		owner.Calls = newReadOnlySpan([]procedureir.CallSite{call})
		index.procedures[0] = owner
		if !objectContainerProcedureCalledBeforeObservation(index, owner, target, 2) {
			t.Fatal("a unique resolved helper must establish the reachable call edge")
		}
	})
}

func TestObjectContainerRejectsConstructorWithResumableHandler(t *testing.T) {
	t.Parallel()
	source := []byte(`Option Explicit

Public Sub Run()
  Dim data As Object
  On Error GoTo Handler
  Set data = CreateObject("Scripting.Dictionary")
  data.Add "actions", Array(data)
  Debug.Print data.Count
  Exit Sub
Handler:
  Resume Next
End Sub
`)
	doc, err := vbaast.ParseDocument("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, doc)
	if err != nil {
		t.Fatal(err)
	}
	controlFlow := cfg.BuildDocument(ir)
	procedures := sourceProceduresFromIRRef(&ir, controlFlow)
	if len(procedures) != 1 {
		t.Fatalf("procedures = %d, want 1", len(procedures))
	}
	file := parsedFile{
		Path:       "Main.bas",
		Lines:      normalizedSourceLines(string(source)),
		Module:     "Main",
		ModuleKind: "standard",
		Source:     source,
		IR:         ir,
		CFG:        controlFlow,
		Parsed:     doc,
		Procedures: procedures,
	}
	proc := procedures[0]
	observationID := 0
	for statement := range proc.Statements.All() {
		if strings.HasPrefix(statement.Text, "Debug.Print data.Count") {
			observationID = statement.ID
			break
		}
	}
	if observationID == 0 {
		t.Fatal("observation statement not found")
	}
	index := buildObjectContainerIndex(file)
	declarations := objectFlowDeclarations(file, proc, file.moduleDecls())
	if objectContainerVariableConstructed(index, proc, "data", observationID, declarations) {
		t.Fatal("a constructor under a handler that resumes into the observation must not be treated as definitely constructed")
	}
}
