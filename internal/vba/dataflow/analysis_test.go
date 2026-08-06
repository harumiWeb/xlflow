package dataflow

import (
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestAnalyzeProcedureTracksParameterAliasAndConcatenation(t *testing.T) {
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{
			Name:             "Run",
			Kind:             procedureir.ProcedureSub,
			Parameters:       []procedureir.Parameter{{Name: "raw", Range: lineRange(1)}},
			DeclarationRange: lineRange(1), BodyRange: lineRange(5),
		},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementAssignment, Text: "aliasValue = raw", Range: lineRange(2), TargetID: 1, ValueID: 2, Target: expr(1, procedureir.ExpressionIdentifier, "aliasValue", 2), Value: expr(2, procedureir.ExpressionIdentifier, "raw", 2)},
			{ID: 2, Kind: procedureir.StatementAssignment, Text: `command = "cmd /c " & aliasValue`, Range: lineRange(3), TargetID: 3, ValueID: 4, Target: expr(3, procedureir.ExpressionIdentifier, "command", 3), Value: expr(4, procedureir.ExpressionBinary, `"cmd /c " & aliasValue`, 3)},
			{ID: 3, Kind: procedureir.StatementCall, Text: "Shell command", Range: lineRange(4), ExpressionIDs: []int{5}},
		},
		Expressions: []procedureir.Expression{
			*expr(1, procedureir.ExpressionIdentifier, "aliasValue", 2),
			*expr(2, procedureir.ExpressionIdentifier, "raw", 2),
			*expr(3, procedureir.ExpressionIdentifier, "command", 3),
			{ID: 4, Kind: procedureir.ExpressionBinary, Text: `"cmd /c " & aliasValue`, Range: lineRange(3), Children: []int{6, 7}, StatementID: 2},
			{ID: 5, Kind: procedureir.ExpressionCall, Text: "Shell command", Range: lineRange(4), StatementID: 3, Children: []int{3}},
			{ID: 6, Kind: procedureir.ExpressionLiteral, Text: `"cmd /c "`, Range: lineRange(3), StatementID: 2},
			{ID: 7, Kind: procedureir.ExpressionIdentifier, Text: "aliasValue", Range: lineRange(3), StatementID: 2},
		},
		Calls: []procedureir.CallSite{{ID: 1, Callee: procedureir.Callee{Text: "Shell", BaseName: "Shell"}, Arguments: procedureir.Arguments{ExpressionIDs: []int{3}}, Range: lineRange(4), StatementID: 3, ExpressionID: 5}},
	}
	result := AnalyzeProcedure(procedure, cfg.Build(procedure), Options{Conservative: true})
	if len(result.Findings) != 1 {
		t.Fatalf("findings = %+v", result.Findings)
	}
	finding := result.Findings[0]
	if finding.Source.Kind != SourceParameter || finding.Sink.Kind != SinkShell || finding.State != StateTainted {
		t.Fatalf("finding = %+v", finding)
	}
	if len(finding.Path) < 3 || finding.Path[1].Kind != "assignment" || !hasPathKind(finding.Path, "concatenation") {
		t.Fatalf("path = %+v", finding.Path)
	}
}

func hasPathKind(path []PathStep, kind string) bool {
	for _, step := range path {
		if step.Kind == kind {
			return true
		}
	}
	return false
}

func TestAnalyzeProcedureChoosesDeterministicRepresentativeAndUnknown(t *testing.T) {
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{Name: "Run", Kind: procedureir.ProcedureSub, DeclarationRange: lineRange(1), BodyRange: lineRange(6)},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementAssignment, Text: "raw = InputBox(\"x\")", Range: lineRange(2), Target: expr(1, procedureir.ExpressionIdentifier, "raw", 2), Value: expr(2, procedureir.ExpressionCall, `InputBox("x")`, 2)},
			{ID: 2, Kind: procedureir.StatementAssignment, Text: "raw = Custom(raw)", Range: lineRange(3), Target: expr(3, procedureir.ExpressionIdentifier, "raw", 3), Value: expr(4, procedureir.ExpressionCall, "Custom(raw)", 3)},
			{ID: 3, Kind: procedureir.StatementCall, Text: "Shell raw", Range: lineRange(4), ExpressionIDs: []int{5}},
		},
		Expressions: []procedureir.Expression{
			*expr(1, procedureir.ExpressionIdentifier, "raw", 2),
			{ID: 2, Kind: procedureir.ExpressionCall, Text: `InputBox("x")`, Range: lineRange(2), StatementID: 1},
			*expr(3, procedureir.ExpressionIdentifier, "raw", 3),
			{ID: 4, Kind: procedureir.ExpressionCall, Text: "Custom(raw)", Range: lineRange(3), StatementID: 2, Children: []int{6}},
			{ID: 5, Kind: procedureir.ExpressionCall, Text: "Shell raw", Range: lineRange(4), StatementID: 3, Children: []int{3}},
			{ID: 6, Kind: procedureir.ExpressionIdentifier, Text: "raw", Range: lineRange(3), StatementID: 2},
		},
		Calls: []procedureir.CallSite{
			{ID: 1, Callee: procedureir.Callee{Text: "InputBox", BaseName: "InputBox"}, Range: lineRange(2), StatementID: 1, ExpressionID: 2},
			{ID: 2, Callee: procedureir.Callee{Text: "Custom", BaseName: "Custom"}, Arguments: procedureir.Arguments{ExpressionIDs: []int{6}}, Range: lineRange(3), StatementID: 2, ExpressionID: 4},
			{ID: 3, Callee: procedureir.Callee{Text: "Shell", BaseName: "Shell"}, Arguments: procedureir.Arguments{ExpressionIDs: []int{3}}, Range: lineRange(4), StatementID: 3, ExpressionID: 5},
		},
	}
	result := AnalyzeProcedure(procedure, cfg.Build(procedure), Options{})
	if len(result.Findings) != 1 || result.Findings[0].State != StateUnknown {
		t.Fatalf("unknown findings = %+v", result.Findings)
	}
	if result.Findings[0].Source.Kind != SourceInputBox || len(result.Findings[0].Path) < 2 {
		t.Fatalf("unknown provenance = %+v", result.Findings[0])
	}
}

func TestJoinStateUsesConservativeStateAndShortestPath(t *testing.T) {
	source := Source{Kind: SourceInputBox, Label: "InputBox", Range: lineRange(2)}
	long := value{origins: map[string]provenance{sourceKey(source): {
		source: source, state: StateTainted,
		path: []PathStep{{Kind: "source", Label: "InputBox"}, {Kind: "assignment", Label: "a"}},
	}}}
	short := value{origins: map[string]provenance{sourceKey(source): {
		source: source, state: StateUnknown,
		path: []PathStep{{Kind: "source", Label: "InputBox"}},
	}}}
	joined, changed := joinState(abstractState{vars: map[string]value{"raw": long}}, abstractState{vars: map[string]value{"raw": short}}, true)
	if !changed || variableState(joined.vars["raw"]) != StateUnknown {
		t.Fatalf("joined state = %+v, changed=%v", joined, changed)
	}
	if len(joined.vars["raw"].origins[sourceKey(source)].path) != 1 {
		t.Fatalf("representative path = %+v", joined.vars["raw"].origins[sourceKey(source)].path)
	}
}

func expr(id int, kind procedureir.ExpressionKind, text string, line int) *procedureir.Expression {
	return &procedureir.Expression{ID: id, Kind: kind, Text: text, Range: lineRange(line), StatementID: line - 1}
}

func lineRange(line int) vbaast.Range {
	return vbaast.Range{StartLine: line, EndLine: line, StartByte: line * 10, EndByte: line*10 + 5}
}
