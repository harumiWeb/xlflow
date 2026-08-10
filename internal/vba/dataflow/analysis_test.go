package dataflow

import (
	"context"
	"reflect"
	"strings"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestAnalyzeProcedureContextReturnsCancellation(t *testing.T) {
	t.Parallel()
	ctx, cancel := context.WithCancel(context.Background())
	cancel()

	result, err := AnalyzeProcedureContext(ctx, procedureir.ProcedureIR{}, cfg.Graph{}, Options{})
	if err != context.Canceled {
		t.Fatalf("AnalyzeProcedureContext error = %v, want context.Canceled", err)
	}
	if !reflect.DeepEqual(result, Result{}) {
		t.Fatalf("canceled result = %#v, want zero result", result)
	}
}

func TestAnalyzeProcedureContextCancelsDuringLargeCFGTraversal(t *testing.T) {
	t.Parallel()
	const blockCount = 20000
	blocks := make([]cfg.Block, 0, blockCount)
	edges := make([]cfg.Edge, 0, blockCount-1)
	for id := 1; id <= blockCount; id++ {
		blocks = append(blocks, cfg.Block{ID: cfg.BlockID(id)})
		if id > 1 {
			edges = append(edges, cfg.Edge{ID: cfg.EdgeID(id - 1), From: cfg.BlockID(id - 1), To: cfg.BlockID(id)})
		}
	}

	ctx := &cancelAfterChecksContext{Context: context.Background(), remaining: 1000}
	result, err := AnalyzeProcedureContext(ctx, procedureir.ProcedureIR{}, cfg.Graph{Blocks: blocks, Edges: edges, Entry: 1}, Options{})
	if err != context.Canceled {
		t.Fatalf("large CFG error = %v, want context.Canceled", err)
	}
	if !reflect.DeepEqual(result, Result{}) {
		t.Fatalf("large CFG canceled result = %#v, want zero result", result)
	}
}

type cancelAfterChecksContext struct {
	context.Context
	remaining int
}

func (c *cancelAfterChecksContext) Err() error {
	if c.remaining <= 0 {
		return context.Canceled
	}
	c.remaining--
	return nil
}

func TestAnalyzeProcedureFindingsDoNotDependOnWorklistRank(t *testing.T) {
	t.Parallel()
	document, err := procedureir.BuildSource(procedureir.BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(`Option Explicit
Sub Run(ByVal raw As String)
    Dim command As String
    Dim index As Long
    command = raw
    For index = 1 To 3
        If index = 2 Then command = command & " suffix"
    Next index
    Shell command
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(document.Procedures) != 1 {
		t.Fatalf("procedures = %d, want 1", len(document.Procedures))
	}
	procedure := document.Procedures[0]
	graph := cfg.Build(procedure)
	reachable := make(map[cfg.BlockID]bool)
	for _, id := range graph.Reachable(cfg.EdgeFilter{}) {
		reachable[id] = true
	}

	normalAnalyzer := newProcedureAnalyzer(procedure, graph, Options{})
	normal, _ := normalAnalyzer.runWithStats()
	if len(normal.Findings) != 1 || normal.Findings[0].Source.Kind != SourceParameter || normal.Findings[0].Sink.Kind != SinkShell {
		t.Fatalf("normal findings = %+v, want one parameter-to-Shell flow", normal.Findings)
	}

	normalRank := normalAnalyzer.reversePostOrderRank(reachable)
	reversedRank := make(map[cfg.BlockID]int, len(normalRank))
	for id, rank := range normalRank {
		reversedRank[id] = len(normalRank) - 1 - rank
	}
	reversedAnalyzer := newProcedureAnalyzer(procedure, graph, Options{})
	// A finding observed from an intermediate state must not leak into the final
	// projection after the fixed point has converged.
	reversedAnalyzer.findings["transient"] = Finding{Source: Source{Kind: SourceUnknown}, Sink: Sink{Kind: SinkShell}}
	reversed, _ := reversedAnalyzer.runWithStatsAndRank(reversedRank)
	if !reflect.DeepEqual(reversed, normal) {
		t.Fatalf("analysis depends on worklist rank:\nnormal   = %#v\nreversed = %#v", normal, reversed)
	}
}

func TestAnalyzeProcedureCommandFindingsClassifyTaintAndStaticPathRisk(t *testing.T) {
	t.Parallel()
	document, err := procedureir.BuildSource(procedureir.BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(`Option Explicit
Sub Run(ByVal raw As String)
    Shell "cmd.exe /c echo " & raw
    Shell "C:\Program Files\Tool\tool.exe"
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	procedure := document.Procedures[0]
	result := AnalyzeProcedure(procedure, cfg.Build(procedure), Options{})
	if len(result.CommandFindings) != 2 {
		t.Fatalf("command findings = %+v, want taint and unquoted-path findings", result.CommandFindings)
	}
	seen := map[CommandRiskKind]bool{}
	for _, finding := range result.CommandFindings {
		seen[finding.RiskKind] = true
	}
	if !seen[CommandRiskTaintedCommandText] || !seen[CommandRiskUnquotedExecutablePath] {
		t.Fatalf("risk kinds = %v", seen)
	}
}

func TestAnalyzeProcedureTreatsKnownConstantsAsCleanAndRespectsLocalBindings(t *testing.T) {
	t.Parallel()
	document, err := procedureir.BuildSource(procedureir.BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(`Option Explicit
Sub SafeRun()
    Const LocalCommand As String = "fixed"
    Shell LocalCommand & ExternalCommand
End Sub

Sub ShadowedRun()
    Dim ExternalCommand As String
    Shell ExternalCommand
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	options := Options{IsKnownConstant: func(name string) bool {
		return strings.EqualFold(name, "ExternalCommand")
	}}

	safe := AnalyzeProcedure(document.Procedures[0], cfg.Build(document.Procedures[0]), options)
	if len(safe.Findings) != 0 {
		t.Fatalf("constant findings = %+v, want none", safe.Findings)
	}

	shadowed := AnalyzeProcedure(document.Procedures[1], cfg.Build(document.Procedures[1]), options)
	if len(shadowed.Findings) != 1 || shadowed.Findings[0].Source.Kind != SourceUnknown {
		t.Fatalf("shadowed constant findings = %+v, want one unknown local flow", shadowed.Findings)
	}
}

func TestAnalyzeProcedureTracksParameterAliasAndConcatenation(t *testing.T) {
	t.Parallel()
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{
			Name:             "Run",
			Kind:             procedureir.ProcedureSub,
			Parameters:       []procedureir.Parameter{{Name: "raw", Range: lineRange(1)}},
			DeclarationRange: lineRange(1), BodyRange: lineRange(5),
		},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementAssignment, Text: "aliasValue = raw", Range: lineRange(2), TargetID: 1, ValueID: 2, Target: expr(1, 1, procedureir.ExpressionIdentifier, "aliasValue", 2), Value: expr(2, 1, procedureir.ExpressionIdentifier, "raw", 2)},
			{ID: 2, Kind: procedureir.StatementAssignment, Text: `command = "cmd /c " & aliasValue`, Range: lineRange(3), TargetID: 3, ValueID: 4, Target: expr(3, 2, procedureir.ExpressionIdentifier, "command", 3), Value: expr(4, 2, procedureir.ExpressionBinary, `"cmd /c " & aliasValue`, 3)},
			{ID: 3, Kind: procedureir.StatementCall, Text: "Shell command", Range: lineRange(4), ExpressionIDs: []int{5}},
		},
		Expressions: []procedureir.Expression{
			*expr(1, 1, procedureir.ExpressionIdentifier, "aliasValue", 2),
			*expr(2, 1, procedureir.ExpressionIdentifier, "raw", 2),
			*expr(3, 2, procedureir.ExpressionIdentifier, "command", 3),
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
	t.Parallel()
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{Name: "Run", Kind: procedureir.ProcedureSub, DeclarationRange: lineRange(1), BodyRange: lineRange(6)},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementAssignment, Text: "raw = InputBox(\"x\")", Range: lineRange(2), Target: expr(1, 1, procedureir.ExpressionIdentifier, "raw", 2), Value: expr(2, 1, procedureir.ExpressionCall, `InputBox("x")`, 2)},
			{ID: 2, Kind: procedureir.StatementAssignment, Text: "raw = Custom(raw)", Range: lineRange(3), Target: expr(3, 2, procedureir.ExpressionIdentifier, "raw", 3), Value: expr(4, 2, procedureir.ExpressionCall, "Custom(raw)", 3)},
			{ID: 3, Kind: procedureir.StatementCall, Text: "Shell raw", Range: lineRange(4), ExpressionIDs: []int{5}},
		},
		Expressions: []procedureir.Expression{
			*expr(1, 1, procedureir.ExpressionIdentifier, "raw", 2),
			{ID: 2, Kind: procedureir.ExpressionCall, Text: `InputBox("x")`, Range: lineRange(2), StatementID: 1},
			*expr(3, 2, procedureir.ExpressionIdentifier, "raw", 3),
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
	t.Parallel()
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

func TestRecoveredAssignmentDoesNotRestoreACleanValue(t *testing.T) {
	t.Parallel()
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{Name: "Run", Kind: procedureir.ProcedureSub, DeclarationRange: lineRange(1), BodyRange: lineRange(5)},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementAssignment, Text: `raw = "initial"`, Range: lineRange(2), Target: expr(1, 1, procedureir.ExpressionIdentifier, "raw", 2), Value: expr(2, 1, procedureir.ExpressionLiteral, `"initial"`, 2)},
			{ID: 2, Kind: procedureir.StatementAssignment, Text: `raw = "fixed"`, Recovered: true, Range: lineRange(3), Target: expr(3, 2, procedureir.ExpressionIdentifier, "raw", 3), Value: expr(4, 2, procedureir.ExpressionLiteral, `"fixed"`, 3)},
			{ID: 3, Kind: procedureir.StatementCall, Text: "Shell raw", Range: lineRange(4), ExpressionIDs: []int{5}},
		},
		Expressions: []procedureir.Expression{
			*expr(1, 1, procedureir.ExpressionIdentifier, "raw", 2),
			*expr(2, 1, procedureir.ExpressionLiteral, `"initial"`, 2),
			*expr(3, 2, procedureir.ExpressionIdentifier, "raw", 3),
			*expr(4, 2, procedureir.ExpressionLiteral, `"fixed"`, 3),
			{ID: 5, Kind: procedureir.ExpressionCall, Text: "Shell raw", Range: lineRange(4), StatementID: 3, Children: []int{6}},
			{ID: 6, Kind: procedureir.ExpressionIdentifier, Text: "raw", Range: lineRange(4), StatementID: 3},
		},
		Calls: []procedureir.CallSite{{ID: 1, Callee: procedureir.Callee{Text: "Shell", BaseName: "Shell"}, Arguments: procedureir.Arguments{ExpressionIDs: []int{6}}, Range: lineRange(4), StatementID: 3, ExpressionID: 5}},
	}
	result := AnalyzeProcedure(procedure, cfg.Build(procedure), Options{})
	if len(result.Findings) != 1 || result.Findings[0].State != StateUnknown {
		t.Fatalf("recovered assignment findings = %+v", result.Findings)
	}
}

func TestJoinStateDoesNotReportChangeForEquivalentMissingValue(t *testing.T) {
	t.Parallel()
	state := abstractState{vars: map[string]value{"raw": unknownStandaloneValue()}}
	_, changed := joinState(state, abstractState{}, true)
	if changed {
		t.Fatal("joinState reported a change for an equivalent missing value")
	}
}

func TestLooksDatabaseReceiverUsesIdentifierBoundaries(t *testing.T) {
	t.Parallel()
	tests := []struct {
		receiver string
		full     string
		want     bool
	}{
		{receiver: "db", full: "db.execute", want: true},
		{receiver: "sql", full: "sql.execute", want: true},
		{receiver: "cn", full: "cn.execute", want: true},
		{receiver: "oldbook", full: "oldbook.execute", want: false},
		{receiver: "sandbox", full: "sandbox.execute", want: false},
		{receiver: "mydb", full: "mydb.execute", want: false},
		{receiver: "sqltable", full: "sqltable.execute", want: false},
	}
	for _, test := range tests {
		t.Run(test.receiver, func(t *testing.T) {
			if got := looksDatabaseReceiver(test.receiver, test.full); got != test.want {
				t.Fatalf("looksDatabaseReceiver(%q, %q) = %v, want %v", test.receiver, test.full, got, test.want)
			}
		})
	}
}

func expr(id, statementID int, kind procedureir.ExpressionKind, text string, line int) *procedureir.Expression {
	return &procedureir.Expression{ID: id, Kind: kind, Text: text, Range: lineRange(line), StatementID: statementID}
}

func lineRange(line int) vbaast.Range {
	return vbaast.Range{StartLine: line, EndLine: line, StartByte: line * 10, EndByte: line*10 + 5}
}
