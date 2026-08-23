package analyze

import (
	"sync"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestSourceProceduresFromIRAttachesProcedureAnalysisFacts(t *testing.T) {
	document := procedureir.DocumentIR{
		ModuleName: "Sheet1",
		Procedures: []procedureir.ProcedureIR{{
			Symbol: procedureir.ProcedureSymbol{
				Name: "Run", Kind: procedureir.ProcedureSub,
				DeclarationRange: vbaast.Range{StartLine: 1, EndLine: 1, StartByte: 0, EndByte: 9},
			},
			Declarations: []procedureir.Declaration{{ID: 1, Name: "value", Scope: procedureir.ScopeLocal}},
			Statements: []procedureir.Statement{
				{ID: 10, Kind: procedureir.StatementAssignment},
				{ID: 20, Kind: procedureir.StatementCall},
			},
			Expressions: []procedureir.Expression{
				{ID: 101, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Sheet1.Range"},
				{ID: 102, StatementID: 10, Kind: procedureir.ExpressionIdentifier, Text: "value"},
				{ID: 103, StatementID: 20, Kind: procedureir.ExpressionMember, Text: "Range.Value"},
			},
			Calls: []procedureir.CallSite{
				{ID: 201, StatementID: 20, Callee: procedureir.Callee{BaseName: "First"}},
				{ID: 202, StatementID: 20, Callee: procedureir.Callee{BaseName: "Second"}},
			},
			Accesses: []procedureir.VariableAccess{
				{Name: "value", StatementID: 10, Mode: procedureir.AccessWrite},
				{Name: "value", StatementID: 20, Mode: procedureir.AccessRead},
			},
		}},
	}

	procedures := sourceProceduresFromIRRef(&document)
	if len(procedures) != 1 || procedures[0].Facts == nil {
		t.Fatalf("source procedures = %#v, want one procedure with facts", procedures)
	}
	if procedures[0].IR != &document.Procedures[0] {
		t.Fatalf("procedure IR = %p, want canonical %p", procedures[0].IR, &document.Procedures[0])
	}
	if procedures[0].Document != &document {
		t.Fatalf("procedure document = %p, want canonical %p", procedures[0].Document, &document)
	}
	if &procedures[0].IR.Declarations[0] != &document.Procedures[0].Declarations[0] ||
		&procedures[0].IR.Statements[0] != &document.Procedures[0].Statements[0] ||
		&procedures[0].IR.Expressions[0] != &document.Procedures[0].Expressions[0] ||
		&procedures[0].IR.Calls[0] != &document.Procedures[0].Calls[0] ||
		&procedures[0].IR.Accesses[0] != &document.Procedures[0].Accesses[0] {
		t.Fatal("procedure view does not use canonical IR collection storage")
	}
	facts := procedures[0].Facts
	if facts.procedure != procedures[0].IR {
		t.Fatalf("facts procedure = %p, want canonical %p", facts.procedure, procedures[0].IR)
	}
	if facts.declarations != nil || facts.statements != nil || facts.expressions != nil || facts.calls != nil || facts.accesses != nil {
		t.Fatal("production facts retained duplicate IR collection storage")
	}
	if got, ok := facts.Declaration(1); !ok || got.Name != "value" {
		t.Fatalf("declaration lookup = %#v, %v", got, ok)
	}
	if got, ok := facts.Statement(10); !ok || got.Kind != procedureir.StatementAssignment {
		t.Fatalf("statement lookup = %#v, %v", got, ok)
	}
	if got, ok := facts.Statements().At(1); !ok || got.ID != 20 {
		t.Fatalf("statement span lookup = %#v, %v", got, ok)
	}
	var statementIDs []int
	for statement := range facts.Statements().All() {
		statementIDs = append(statementIDs, statement.ID)
	}
	if len(statementIDs) != 2 || statementIDs[0] != 10 || statementIDs[1] != 20 {
		t.Fatalf("statement span iteration = %#v", statementIDs)
	}
	if _, ok := facts.Statement(999); ok {
		t.Fatalf("missing statement lookup unexpectedly succeeded")
	}
	if got, ok := facts.Expression(103); !ok || got.Text != "Range.Value" {
		t.Fatalf("expression lookup = %#v, %v", got, ok)
	}
	if got := facts.CallsForStatement(20); len(got) != 2 || got[0].ID != 201 || got[1].ID != 202 {
		t.Fatalf("calls by statement = %#v", got)
	}
	if got := facts.AccessesForStatement(10); len(got) != 1 || got[0].Name != "value" || got[0].Mode != procedureir.AccessWrite {
		t.Fatalf("accesses by statement = %#v", got)
	}
	if got := facts.MemberExpressionsForStatement(10); len(got) != 1 || got[0].ID != 101 {
		t.Fatalf("member expressions = %#v", got)
	}
	if got := facts.MemberExpressionsForStatement(20); len(got) != 1 || got[0].ID != 103 {
		t.Fatalf("member expressions for call = %#v", got)
	}

	// The canonical DocumentIR owns the view storage. Facts retain compact
	// indexes, so reading through an index observes the current canonical value
	// without retaining a copied IR object.
	document.Procedures[0].Statements[0].ID = 999
	if got, ok := facts.Statement(10); !ok || got.ID != 999 {
		t.Fatalf("facts did not read canonical IR after owner mutation = %#v, %v", got, ok)
	}
}

func TestProcedureAnalysisFactsConcurrentReads(t *testing.T) {
	statements := []procedureir.Statement{{ID: 10, ExpressionIDs: []int{1}}}
	expressions := []procedureir.Expression{{ID: 1, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Application.Range"}}
	calls := []procedureir.CallSite{{ID: 20, StatementID: 10}}
	accesses := []procedureir.VariableAccess{{Name: "value", StatementID: 10, Mode: procedureir.AccessRead}}
	facts := newProcedureAnalysisFacts(statements, expressions, calls, accesses)
	var wait sync.WaitGroup
	for worker := 0; worker < 8; worker++ {
		wait.Add(1)
		go func() {
			defer wait.Done()
			for iteration := 0; iteration < 1000; iteration++ {
				if _, ok := facts.Statement(10); !ok {
					t.Errorf("statement lookup failed")
					return
				}
				if _, ok := facts.Expression(1); !ok {
					t.Errorf("expression lookup failed")
					return
				}
				if len(facts.CallsForStatement(10)) != 1 || len(facts.AccessesForStatement(10)) != 1 || len(facts.MemberExpressionsForStatement(10)) != 1 {
					t.Errorf("grouped fact lookup returned incomplete data")
					return
				}
			}
		}()
	}
	wait.Wait()
}

func TestProcedureAnalysisFactsEmptyInputsRemainAllocationFree(t *testing.T) {
	facts := newProcedureAnalysisFacts(nil, nil, nil, nil)
	if facts == nil {
		t.Fatalf("empty facts = nil")
	}
	if facts.declarationIndex != nil || facts.statementIndex != nil || facts.expressionIndex != nil || facts.memberExpressionsByStatement != nil {
		t.Fatalf("empty facts allocated indexes: %#v", facts)
	}
	if _, ok := facts.Statement(1); ok {
		t.Fatalf("empty statement lookup unexpectedly succeeded")
	}
	if got := facts.CallsForStatement(1); got != nil {
		t.Fatalf("empty calls = %#v, want nil", got)
	}
	if got := facts.MemberExpressionsForStatement(1); got != nil {
		t.Fatalf("empty member expressions = %#v, want nil", got)
	}
}

func TestProcedureAnalysisFactsPreservesInterleavedStatementGroups(t *testing.T) {
	calls := []procedureir.CallSite{
		{ID: 1, StatementID: 10},
		{ID: 2, StatementID: 20},
		{ID: 3, StatementID: 10},
	}
	accesses := []procedureir.VariableAccess{
		{Name: "a", StatementID: 10},
		{Name: "b", StatementID: 20},
		{Name: "c", StatementID: 10},
	}
	facts := newProcedureAnalysisFacts(nil, nil, calls, accesses)
	if got := facts.CallsForStatement(10); len(got) != 2 || got[0].ID != 1 || got[1].ID != 3 {
		t.Fatalf("interleaved calls = %#v", got)
	}
	if got := facts.AccessesForStatement(10); len(got) != 2 || got[0].Name != "a" || got[1].Name != "c" {
		t.Fatalf("interleaved accesses = %#v", got)
	}
}

func TestProcedureAnalysisFactsMemberExpressionsPreserveRecoveryAndFallback(t *testing.T) {
	statements := []procedureir.Statement{{ID: 10, ExpressionIDs: []int{1}}}
	expressions := []procedureir.Expression{
		{ID: 1, Kind: procedureir.ExpressionMember, Text: "Application.Range", Children: []int{2, 3}},
		{ID: 2, Kind: procedureir.ExpressionMember, Text: "Range.Value2"},
		{ID: 3, Kind: procedureir.ExpressionMember, Text: "Recovered.Member", Recovered: true, StatementID: 10},
	}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	got := facts.MemberExpressionsForStatement(10)
	if len(got) != 2 || got[0].ID != 1 || got[1].ID != 2 {
		t.Fatalf("member expression fallback = %#v, want reachable non-recovered members", got)
	}
}

func TestProcedureAnalysisFactsMemberExpressionsMixedStatementIDFallback(t *testing.T) {
	statements := []procedureir.Statement{{ID: 10, ExpressionIDs: []int{1}}}
	expressions := []procedureir.Expression{
		{ID: 1, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Application.Range", Children: []int{2}},
		{ID: 2, Kind: procedureir.ExpressionMember, Text: "Range.Value2"},
	}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	got := facts.MemberExpressionsForStatement(10)
	if len(got) != 2 || got[0].ID != 1 || got[1].ID != 2 {
		t.Fatalf("mixed StatementID member expression fallback = %#v, want both reachable members", got)
	}
}

func TestProcedureAnalysisFactsMemberExpressionsReturnsOwnedCopy(t *testing.T) {
	statements := []procedureir.Statement{{ID: 10}}
	expressions := []procedureir.Expression{{ID: 1, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Range.Value"}}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	got := facts.MemberExpressionsForStatement(10)
	if len(got) != 1 {
		t.Fatalf("member expressions = %#v, want one expression", got)
	}
	got[0].ID = 99
	got = facts.MemberExpressionsForStatement(10)
	if len(got) != 1 || got[0].ID != 1 {
		t.Fatalf("member expressions were not returned as an owned copy: %#v", got)
	}
}

func TestProcedureAnalysisFactsMemberExpressionIteratorPreservesOrder(t *testing.T) {
	statements := []procedureir.Statement{{ID: 10}, {ID: 20}}
	expressions := []procedureir.Expression{
		{ID: 1, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Application.Range", Range: vbaast.Range{StartByte: 30}},
		{ID: 2, StatementID: 20, Kind: procedureir.ExpressionMember, Text: "Other.Range", Range: vbaast.Range{StartByte: 40}},
		{ID: 3, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Range.Value", Range: vbaast.Range{StartByte: 50}},
		{ID: 4, StatementID: 10, Kind: procedureir.ExpressionIdentifier, Text: "value", Range: vbaast.Range{StartByte: 60}},
	}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	var got []int
	facts.forEachMemberExpressionForStatement(10, func(expression procedureir.Expression) {
		got = append(got, expression.ID)
	})
	if len(got) != 2 || got[0] != 1 || got[1] != 3 {
		t.Fatalf("member expression iterator = %#v, want [1 3]", got)
	}
}

func TestProcedureAnalysisFactsMemberExpressionIteratorMatchesFallbackOrder(t *testing.T) {
	statements := []procedureir.Statement{{ID: 10, ExpressionIDs: []int{1}}}
	expressions := []procedureir.Expression{
		{ID: 1, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Application.Range", Children: []int{2}},
		{ID: 2, Kind: procedureir.ExpressionMember, Text: "Range.Value2"},
	}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	var got []int
	facts.forEachMemberExpressionForStatement(10, func(expression procedureir.Expression) {
		got = append(got, expression.ID)
	})
	if len(got) != 2 || got[0] != 1 || got[1] != 2 {
		t.Fatalf("fallback member expression iterator = %#v, want [1 2]", got)
	}
}

func TestProcedureAnalysisFactsMemberExpressionIteratorIsAllocationFree(t *testing.T) {
	statements := []procedureir.Statement{{ID: 10}}
	expressions := []procedureir.Expression{
		{ID: 1, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Application.Range"},
		{ID: 2, StatementID: 10, Kind: procedureir.ExpressionMember, Text: "Range.Value"},
	}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	allocs := testing.AllocsPerRun(100, func() {
		facts.forEachMemberExpressionForStatement(10, consumeMemberExpression)
	})
	if allocs != 0 {
		t.Fatalf("authoritative member expression iteration allocated %v times", allocs)
	}
}

func BenchmarkProcedureAnalysisFactsMemberExpressionIterator(b *testing.B) {
	statements := []procedureir.Statement{{ID: 10}}
	expressions := make([]procedureir.Expression, 32)
	for index := range expressions {
		expressions[index] = procedureir.Expression{
			ID: index + 1, StatementID: 10, Kind: procedureir.ExpressionMember,
			Text: "Application.Range", Range: vbaast.Range{StartByte: index},
		}
	}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	b.ReportAllocs()
	b.ResetTimer()
	for index := 0; index < b.N; index++ {
		facts.forEachMemberExpressionForStatement(10, consumeMemberExpression)
	}
}

func consumeMemberExpression(procedureir.Expression) {}

var benchmarkProcedureFactReadSink int

func BenchmarkProcedureAnalysisFactsReadOnlyViews(b *testing.B) {
	const factCount = 4096
	statements := make([]procedureir.Statement, factCount)
	expressions := make([]procedureir.Expression, factCount)
	calls := make([]procedureir.CallSite, factCount)
	accesses := make([]procedureir.VariableAccess, factCount)
	for index := 0; index < factCount; index++ {
		id := index + 1
		statements[index] = procedureir.Statement{ID: id}
		expressions[index] = procedureir.Expression{ID: id, StatementID: 1, Kind: procedureir.ExpressionIdentifier}
		calls[index] = procedureir.CallSite{ID: id, StatementID: 1}
		accesses[index] = procedureir.VariableAccess{Name: "value", StatementID: 1, Mode: procedureir.AccessRead}
	}
	facts := newProcedureAnalysisFacts(statements, expressions, calls, accesses)

	b.Run("statements", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			total := 0
			for statement := range facts.Statements().All() {
				total += statement.ID
			}
			benchmarkProcedureFactReadSink = total
		}
	})
	b.Run("expressions", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			total := 0
			for expression := range facts.Expressions().All() {
				total += expression.ID
			}
			benchmarkProcedureFactReadSink = total
		}
	})
	b.Run("calls", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			total := 0
			for call := range facts.Calls().All() {
				total += call.ID
			}
			benchmarkProcedureFactReadSink = total
		}
	})
	b.Run("accesses", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			total := 0
			for access := range facts.Accesses().All() {
				total += access.StatementID
			}
			benchmarkProcedureFactReadSink = total
		}
	})
	b.Run("grouped-calls", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			facts.forEachCallForStatement(1, consumeCallSite)
		}
	})
	b.Run("grouped-accesses", func(b *testing.B) {
		b.ReportAllocs()
		for b.Loop() {
			facts.forEachAccessForStatement(1, consumeVariableAccess)
		}
	})
}

func consumeCallSite(procedureir.CallSite)             {}
func consumeVariableAccess(procedureir.VariableAccess) {}

func TestProcedureAnalysisFactsEmptyAuthoritativeMemberGroupDoesNotTraverseOtherStatements(t *testing.T) {
	statements := []procedureir.Statement{
		{ID: 10, ExpressionIDs: []int{1}},
		{ID: 20},
	}
	expressions := []procedureir.Expression{
		{ID: 1, StatementID: 10, Kind: procedureir.ExpressionIdentifier, Children: []int{2}},
		{ID: 2, StatementID: 20, Kind: procedureir.ExpressionMember, Text: "Range.Value"},
	}
	facts := newProcedureAnalysisFacts(statements, expressions, nil, nil)
	if got := facts.MemberExpressionsForStatement(10); len(got) != 0 {
		t.Fatalf("authoritative empty member group = %#v, want nil", got)
	}
}

func TestIndexGroupsLookupExpandsContiguousSpan(t *testing.T) {
	groups := newIndexGroups([]int{10, 10})
	got, ok := groups.lookup(10)
	if !ok || len(got) != 2 || got[0] != 0 || got[1] != 1 {
		t.Fatalf("contiguous lookup = %#v, %v; want indexes [0 1]", got, ok)
	}
}
