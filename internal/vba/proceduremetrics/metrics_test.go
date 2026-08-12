package proceduremetrics

import (
	"reflect"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func testRange(start, end int) vbaast.Range {
	return vbaast.Range{StartLine: start, EndLine: end, StartByte: start * 10, EndByte: end * 10}
}

func TestCollectComputesAllProcedureMetrics(t *testing.T) {
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{
			Name: "Run", Kind: procedureir.ProcedureFunction,
			DeclarationRange: testRange(2, 20),
			Parameters: []procedureir.Parameter{
				{Name: "implicitRef"},
				{Name: "byVal", Passing: "ByVal"},
				{Name: "byRef", Passing: "ByRef"},
			},
		},
		Declarations: []procedureir.Declaration{
			{Scope: procedureir.ScopeLocal, Kind: "return_slot"},
			{Scope: procedureir.ScopeLocal, Kind: "variable"},
			{Scope: procedureir.ScopeLocal, Kind: "const"},
			{Scope: procedureir.ScopeParameter, Kind: "parameter"},
		},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementIf, Range: testRange(3, 3)},
			{ID: 2, ParentID: 1, Kind: procedureir.StatementElseIf, Range: testRange(4, 4)},
			{ID: 3, ParentID: 1, Kind: procedureir.StatementElse, Range: testRange(5, 5)},
			{ID: 4, ParentID: 1, Kind: procedureir.StatementSelect, Range: testRange(6, 6)},
			{ID: 5, ParentID: 4, Kind: procedureir.StatementCase, Range: testRange(7, 7)},
			{ID: 6, ParentID: 5, Kind: procedureir.StatementFor, Range: testRange(8, 8)},
			{ID: 7, ParentID: 6, Kind: procedureir.StatementForEach, Range: testRange(9, 9)},
			{ID: 8, ParentID: 7, Kind: procedureir.StatementWhile, Range: testRange(10, 10)},
			{ID: 9, ParentID: 8, Kind: procedureir.StatementDo, Range: testRange(11, 11)},
			{ID: 10, ParentID: 9, SyntaxKind: "do_condition", Kind: procedureir.StatementUnknown, Range: testRange(12, 12)},
			{ID: 11, Kind: procedureir.StatementGoTo, Range: testRange(13, 13)},
			{ID: 12, Kind: procedureir.StatementOnError, Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferOnErrorGoto}, Range: testRange(14, 14)},
			{ID: 13, Kind: procedureir.StatementExit, Text: "Exit Function", Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitFunction}, Range: testRange(15, 15)},
			{ID: 14, Kind: procedureir.StatementExit, Text: "Exit For", Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitFor}, Range: testRange(16, 16)},
			{ID: 15, Kind: procedureir.StatementEnd, Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferTerminate}, Range: testRange(17, 17)},
			{ID: 16, ParentID: 4, Kind: procedureir.StatementCase, Control: &procedureir.ControlFlowMetadata{CaseElse: true}, Range: testRange(18, 18)},
		},
		Calls: []procedureir.CallSite{
			{Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "Other.First"}}}},
			{Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "other.first"}}}},
			{Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "Other.Second"}}}},
			{Resolution: procedureir.CallResolution{Status: procedureir.ResolutionAmbiguous, Candidates: []procedureir.Candidate{{QualifiedName: "Other.A"}, {QualifiedName: "Other.B"}}}},
		},
	}
	got := Collect(Input{IR: procedure, File: "src/Main.bas", Module: "Main", ModuleKind: "standard"})
	want := ProcedureMetrics{
		File: "src/Main.bas", Module: "Main", ModuleKind: "standard", Name: "Run", Kind: procedureir.ProcedureFunction,
		Visibility:         "",
		ResolvedCallees:    []string{"other.first", "other.second"},
		ErrorHandlingCount: 1,
		AmbiguousCallCount: 1,
		DeclarationRange:   testRange(2, 20),
		Metrics: Metrics{
			CyclomaticComplexity: 8, MaxNestingDepth: 6, StatementCount: 15,
			SourceLineCount: 19, BranchCount: 3, LoopCount: 4, GotoCount: 1,
			ExitPointCount: 3, ParameterCount: 3, ByRefParameterCount: 2,
			LocalVariableCount: 2, CallFanOut: 2,
		},
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("metrics = %+v, want %+v", got, want)
	}
}

func TestEvaluateThresholdsUsesStrictPositiveThresholdsAndStableOrder(t *testing.T) {
	procedures := []ProcedureMetrics{
		{Name: "Second", File: "b.bas", DeclarationRange: testRange(4, 4), Metrics: Metrics{LoopCount: 3, GotoCount: 2}},
		{Name: "First", File: "a.bas", DeclarationRange: testRange(9, 9), Metrics: Metrics{LoopCount: 3, GotoCount: 1}},
	}
	violations, err := EvaluateThresholds(procedures, Thresholds{
		MetricLoopCount:      3, // equal is allowed
		MetricGotoCount:      1, // only Second exceeds
		MetricStatementCount: 0, // disabled
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(violations) != 1 || violations[0].Procedure != "Second" || violations[0].Metric != MetricGotoCount ||
		violations[0].Code != "MX001" || violations[0].Severity != "warning" {
		t.Fatalf("violations = %+v", violations)
	}
	if _, err := EvaluateThresholds(nil, Thresholds{MetricLoopCount: -1}); err == nil {
		t.Fatal("negative threshold was accepted")
	}
	if _, err := EvaluateThresholds(nil, Thresholds{MetricName("typo"): 1}); err == nil {
		t.Fatal("unknown threshold was accepted")
	}
}

func TestPropertyExitAndLoopExitAreCountedSeparately(t *testing.T) {
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{
			Name: "Value", Kind: procedureir.ProcedurePropertySet,
			DeclarationRange: testRange(1, 5),
		},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementExit, Text: "Exit Property", Range: testRange(3, 3)},
			{ID: 2, Kind: procedureir.StatementExit, Text: "Exit Do", Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitDo}, Range: testRange(4, 4)},
			{ID: 3, Kind: procedureir.StatementEnd, Range: testRange(5, 5)},
		},
	}
	got := Collect(Input{IR: procedure})
	if got.ExitPointCount != 3 {
		t.Fatalf("property exits = %d, want implicit + Exit Property + End = 3", got.ExitPointCount)
	}
}

func TestDoConditionDoesNotBreakNestingForRecoveredChildren(t *testing.T) {
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{Kind: procedureir.ProcedureSub, DeclarationRange: testRange(1, 4)},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementDo, Range: testRange(2, 2)},
			{ID: 2, ParentID: 1, Kind: procedureir.StatementDo, SyntaxKind: "do_condition", Range: testRange(3, 3)},
			{ID: 3, ParentID: 2, Kind: procedureir.StatementIf, Range: testRange(4, 4)},
		},
	}
	got := Collect(Input{IR: procedure})
	if got.MaxNestingDepth != 2 {
		t.Fatalf("nesting through do_condition = %d, want 2", got.MaxNestingDepth)
	}
}

func TestCollectDocumentSortsByFileAndDeclarationOffset(t *testing.T) {
	document := procedureir.DocumentIR{
		Path: "z.bas", ModuleName: "Z", ModuleKind: "standard",
		Procedures: []procedureir.ProcedureIR{
			{Symbol: procedureir.ProcedureSymbol{Name: "Later", DeclarationRange: testRange(5, 5)}},
			{Symbol: procedureir.ProcedureSymbol{Name: "Earlier", DeclarationRange: testRange(2, 2)}},
		},
	}
	got := CollectDocument(document, cfg.Document{})
	if len(got) != 2 || got[0].Name != "Earlier" || got[1].Name != "Later" {
		t.Fatalf("document metrics order = %+v", got)
	}
	if got[0].File != "z.bas" || got[0].Module != "Z" || got[0].ModuleKind != "standard" {
		t.Fatalf("document identity = %+v", got[0])
	}
}
