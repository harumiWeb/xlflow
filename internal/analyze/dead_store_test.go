package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestAnalyzerReportsDeadStoreWhenOptedIn(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  Dim value As Long
  value = 1
  Debug.Print "done"
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectDeadStores = true
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	dead := findingsByCode(findings, "VBA256")
	if len(dead) != 1 || dead[0].Line != 4 {
		t.Fatalf("VBA256 findings = %+v, want one finding at assignment line", dead)
	}
}

func TestDeadStoreCandidatesReportsPlainScalarWrite(t *testing.T) {
	statements := []procedureir.Statement{deadStoreAssignment(1, "value", 2)}
	proc := deadStoreTestProcedure(
		[]procedureir.Declaration{deadStoreDeclaration("value", "Long")},
		statements,
		[]procedureir.VariableAccess{deadStoreAccess(1, "value", procedureir.AccessWrite)},
		[]vbacfg.Edge{{From: 1, To: 10, Class: vbacfg.EdgeNormal}, {From: 10, To: 90, Class: vbacfg.EdgeNormal}},
		nil,
	)

	got := deadStoreCandidates(proc)
	if len(got) != 1 || got[0].StatementID != 1 || got[0].Name != "value" || got[0].Range.StartLine != 2 {
		t.Fatalf("deadStoreCandidates = %+v, want one value assignment", got)
	}
}

func TestDeadStoreCandidatesRecognizesLaterReadsAndRHSReads(t *testing.T) {
	statements := []procedureir.Statement{
		deadStoreAssignment(1, "value", 2),
		deadStoreAssignment(2, "value", 3),
		deadStoreCall(3, 4),
		deadStoreAssignment(4, "value", 5),
	}
	accesses := []procedureir.VariableAccess{
		deadStoreAccess(1, "value", procedureir.AccessWrite),
		deadStoreAccess(2, "value", procedureir.AccessWrite),
		deadStoreAccess(3, "value", procedureir.AccessRead),
		deadStoreAccess(4, "value", procedureir.AccessWrite),
		deadStoreAccess(4, "value", procedureir.AccessRead),
	}
	proc := deadStoreTestProcedure(
		[]procedureir.Declaration{deadStoreDeclaration("value", "Long")},
		statements,
		accesses,
		[]vbacfg.Edge{
			{From: 1, To: 10, Class: vbacfg.EdgeNormal},
			{From: 10, To: 20, Class: vbacfg.EdgeNormal},
			{From: 20, To: 30, Class: vbacfg.EdgeNormal},
			{From: 30, To: 40, Class: vbacfg.EdgeNormal},
			{From: 40, To: 90, Class: vbacfg.EdgeNormal},
		},
		nil,
	)

	got := deadStoreCandidates(proc)
	if len(got) != 2 || got[0].StatementID != 1 || got[1].StatementID != 4 {
		t.Fatalf("deadStoreCandidates = %+v, want overwritten statement 1 and final statement 4", got)
	}
}

func TestDeadStoreCandidatesUsesBranchLoopAndGotoEdges(t *testing.T) {
	statements := []procedureir.Statement{
		deadStoreAssignment(1, "value", 2),
		deadStoreCall(2, 3),
	}
	proc := deadStoreTestProcedure(
		[]procedureir.Declaration{deadStoreDeclaration("value", "Long")},
		statements,
		[]procedureir.VariableAccess{
			deadStoreAccess(1, "value", procedureir.AccessWrite),
			deadStoreAccess(2, "value", procedureir.AccessRead),
		},
		[]vbacfg.Edge{
			{From: 1, To: 10, Class: vbacfg.EdgeNormal},
			{From: 10, To: 20, Kind: vbacfg.EdgeBranchTrue, Class: vbacfg.EdgeNormal},
			{From: 10, To: 90, Kind: vbacfg.EdgeBranchFalse, Class: vbacfg.EdgeNormal},
			{From: 20, To: 10, Kind: vbacfg.EdgeGoto, Class: vbacfg.EdgeNormal},
			{From: 20, To: 90, Class: vbacfg.EdgeNormal},
		},
		nil,
	)

	if got := deadStoreCandidates(proc); len(got) != 0 {
		t.Fatalf("deadStoreCandidates = %+v, want no finding because the loop branch reads value", got)
	}
}

func TestDeadStoreCandidatesDoesNotTreatExceptionalReadAsCompletedWriteUse(t *testing.T) {
	statements := []procedureir.Statement{
		deadStoreAssignment(1, "value", 2),
		deadStoreAssignment(2, "value", 3),
		deadStoreCall(3, 4),
	}
	proc := deadStoreTestProcedure(
		[]procedureir.Declaration{deadStoreDeclaration("value", "Long")},
		statements,
		[]procedureir.VariableAccess{
			deadStoreAccess(1, "value", procedureir.AccessWrite),
			deadStoreAccess(2, "value", procedureir.AccessWrite),
			deadStoreAccess(3, "value", procedureir.AccessRead),
		},
		[]vbacfg.Edge{
			{From: 1, To: 10, Class: vbacfg.EdgeNormal},
			{From: 10, To: 20, Class: vbacfg.EdgeNormal},
			{From: 20, To: 90, Class: vbacfg.EdgeNormal},
			{From: 20, To: 30, Class: vbacfg.EdgeExceptional, Uncertain: true},
			{From: 30, To: 90, Class: vbacfg.EdgeNormal},
		},
		nil,
	)

	got := deadStoreCandidates(proc)
	if len(got) != 1 || got[0].StatementID != 2 {
		t.Fatalf("deadStoreCandidates = %+v, want only completed write statement 2", got)
	}
}

func TestDeadStoreCandidatesFailsOpenForUnknownAndUncertainFlow(t *testing.T) {
	statements := []procedureir.Statement{deadStoreAssignment(1, "value", 2)}
	declaration := []procedureir.Declaration{deadStoreDeclaration("value", "Long")}
	access := []procedureir.VariableAccess{deadStoreAccess(1, "value", procedureir.AccessWrite)}

	unknown := deadStoreTestProcedure(
		declaration, statements, access,
		[]vbacfg.Edge{{From: 1, To: 10, Class: vbacfg.EdgeNormal}, {From: 10, To: 90, Class: vbacfg.EdgeNormal}},
		[]vbacfg.BlockID{10},
	)
	if got := deadStoreCandidates(unknown); len(got) != 0 {
		t.Fatalf("unknown-flow candidates = %+v, want none", got)
	}

	uncertain := deadStoreTestProcedure(
		declaration, statements, access,
		[]vbacfg.Edge{{From: 1, To: 10, Class: vbacfg.EdgeNormal}, {From: 10, To: 90, Class: vbacfg.EdgeNormal, Uncertain: true}},
		nil,
	)
	if got := deadStoreCandidates(uncertain); len(got) != 0 {
		t.Fatalf("uncertain-flow candidates = %+v, want none", got)
	}
}

func TestDeadStoreCandidatesTreatsReadWriteAsObservation(t *testing.T) {
	statements := []procedureir.Statement{
		deadStoreAssignment(1, "value", 2),
		deadStoreCall(2, 3),
	}
	proc := deadStoreTestProcedure(
		[]procedureir.Declaration{deadStoreDeclaration("value", "Long")},
		statements,
		[]procedureir.VariableAccess{
			deadStoreAccess(1, "value", procedureir.AccessWrite),
			deadStoreAccess(2, "value", procedureir.AccessReadWrite),
		},
		[]vbacfg.Edge{
			{From: 1, To: 10, Class: vbacfg.EdgeNormal},
			{From: 10, To: 20, Class: vbacfg.EdgeNormal},
			{From: 20, To: 90, Class: vbacfg.EdgeNormal},
		},
		nil,
	)
	if got := deadStoreCandidates(proc); len(got) != 0 {
		t.Fatalf("deadStoreCandidates = %+v, want no finding for ByRef/read-write observation", got)
	}
}

func TestDeadStoreCandidatesExcludesNonScalarAndNonLocalDeclarations(t *testing.T) {
	statements := []procedureir.Statement{
		deadStoreAssignment(1, "scalar", 2),
		deadStoreAssignment(2, "staticValue", 3),
		deadStoreAssignment(3, "arrayValue", 4),
		deadStoreAssignment(4, "objectValue", 5),
		deadStoreAssignment(5, "variantValue", 6),
	}
	declarations := []procedureir.Declaration{
		deadStoreDeclaration("scalar", "Long"),
		{ID: 2, Name: "staticValue", Type: "Long", Scope: procedureir.ScopeLocal, IsStatic: true},
		{ID: 3, Name: "arrayValue", Type: "Long", Scope: procedureir.ScopeLocal, IsArray: true},
		{ID: 4, Name: "objectValue", Type: "Object", Scope: procedureir.ScopeLocal, IsObject: true},
		{ID: 5, Name: "variantValue", Type: "Variant", Scope: procedureir.ScopeLocal, ValueShape: procedureir.ValueShapeVariant},
	}
	accesses := make([]procedureir.VariableAccess, 0, len(statements))
	for _, statement := range statements {
		accesses = append(accesses, deadStoreAccess(statement.ID, statement.Target.Text, procedureir.AccessWrite))
	}
	proc := deadStoreTestProcedure(
		declarations, statements, accesses,
		[]vbacfg.Edge{
			{From: 1, To: 10, Class: vbacfg.EdgeNormal},
			{From: 10, To: 20, Class: vbacfg.EdgeNormal},
			{From: 20, To: 30, Class: vbacfg.EdgeNormal},
			{From: 30, To: 40, Class: vbacfg.EdgeNormal},
			{From: 40, To: 50, Class: vbacfg.EdgeNormal},
			{From: 50, To: 90, Class: vbacfg.EdgeNormal},
		},
		nil,
	)
	got := deadStoreCandidates(proc)
	if len(got) != 1 || got[0].Name != "scalar" {
		t.Fatalf("deadStoreCandidates = %+v, want only scalar local", got)
	}
}

func TestDeadStoreCandidatesExcludesFunctionReturnSlot(t *testing.T) {
	statement := deadStoreAssignment(1, "Compute", 2)
	proc := deadStoreTestProcedure(
		[]procedureir.Declaration{{Name: "Compute", Type: "Long", Scope: procedureir.ScopeLocal, Kind: "return_slot"}},
		[]procedureir.Statement{statement},
		[]procedureir.VariableAccess{deadStoreAccess(1, "Compute", procedureir.AccessWrite)},
		[]vbacfg.Edge{{From: 1, To: 10, Class: vbacfg.EdgeNormal}, {From: 10, To: 90, Class: vbacfg.EdgeNormal}},
		nil,
	)
	if got := deadStoreCandidates(proc); len(got) != 0 {
		t.Fatalf("deadStoreCandidates = %+v, want no return-slot finding", got)
	}
}

func deadStoreTestProcedure(declarations []procedureir.Declaration, statements []procedureir.Statement, accesses []procedureir.VariableAccess, edges []vbacfg.Edge, unknown []vbacfg.BlockID) sourceProcedure {
	blocks := []vbacfg.Block{{ID: 1, Kind: vbacfg.BlockEntry}}
	for index := range statements {
		blocks = append(blocks, vbacfg.Block{
			ID: 10 + vbacfg.BlockID(index*10), Kind: vbacfg.BlockStatement,
			StatementID: statements[index].ID, Statement: &statements[index], Range: statements[index].Range,
		})
	}
	blocks = append(blocks,
		vbacfg.Block{ID: 90, Kind: vbacfg.BlockNormalExit},
		vbacfg.Block{ID: 91, Kind: vbacfg.BlockExceptionalExit},
		vbacfg.Block{ID: 92, Kind: vbacfg.BlockTerminationExit},
		vbacfg.Block{ID: 93, Kind: vbacfg.BlockUnknownExit},
	)
	graph := &vbacfg.Graph{
		Blocks: blocks, Edges: edges, UnknownFlowSources: unknown,
		Entry: 1, NormalExit: 90, ExceptionalExit: 91, TerminationExit: 92, UnknownExit: 93,
	}
	return sourceProcedure{
		Declarations: newReadOnlySpan(declarations), Statements: newReadOnlySpan(statements),
		Accesses: newReadOnlySpan(accesses), Graph: graph,
	}
}

func deadStoreDeclaration(name, typ string) procedureir.Declaration {
	return procedureir.Declaration{Name: name, Type: typ, Scope: procedureir.ScopeLocal}
}

func deadStoreAssignment(id int, name string, line int) procedureir.Statement {
	return procedureir.Statement{
		ID: id, Kind: procedureir.StatementAssignment, Text: name + " = 1",
		Range:  vbaast.Range{StartLine: line, EndLine: line},
		Target: &procedureir.Expression{Kind: procedureir.ExpressionIdentifier, Text: name},
		Value:  &procedureir.Expression{Kind: procedureir.ExpressionLiteral, Text: "1"},
	}
}

func deadStoreCall(id, line int) procedureir.Statement {
	return procedureir.Statement{
		ID: id, Kind: procedureir.StatementCall, Text: "Consume(value)",
		Range: vbaast.Range{StartLine: line, EndLine: line},
	}
}

func deadStoreAccess(statementID int, name string, mode procedureir.AccessMode) procedureir.VariableAccess {
	return procedureir.VariableAccess{
		Name: name, Mode: mode, Scope: procedureir.ScopeLocal, StatementID: statementID,
	}
}
