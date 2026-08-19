package cfg

import (
	"context"
	"errors"
	"reflect"
	"strings"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestBuildDocumentContextReturnsCancellationWithoutPartialGraphs(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	result, err := BuildDocumentContext(ctx, procedureir.DocumentIR{Path: "Main.bas", Procedures: []procedureir.ProcedureIR{{}}})
	if !errors.Is(err, context.Canceled) || !reflect.DeepEqual(result, Document{}) {
		t.Fatalf("canceled result = (%+v, %v)", result, err)
	}
}

func TestWithoutNormalErrRaiseContinuationKeepsExceptionalFlow(t *testing.T) {
	t.Parallel()
	doc := buildIR(t, `Public Sub Work()
    On Error GoTo Handler
    Err.Raise 5
    Debug.Print "unreachable"
    Exit Sub
Handler:
    Resume Next
End Sub
`)
	graph := BuildDocument(doc).Graphs[0]
	raise := 0
	for _, statement := range doc.Procedures[0].Statements {
		if strings.HasPrefix(strings.ToLower(strings.TrimSpace(statement.Text)), "err.raise") {
			raise = statement.ID
			break
		}
	}
	block, ok := graph.BlockForStatement(raise)
	if !ok {
		t.Fatal("Err.Raise block not found")
	}
	filtered := graph.WithoutNormalErrRaiseContinuation()
	var normal, exceptional bool
	for _, edge := range filtered.Edges {
		if edge.From != block.ID {
			continue
		}
		normal = normal || edge.Class == EdgeNormal
		exceptional = exceptional || edge.Class == EdgeExceptional
	}
	if normal || !exceptional {
		t.Fatalf("Err.Raise outgoing flow: normal=%v exceptional=%v", normal, exceptional)
	}
	if len(graph.Edges) == len(filtered.Edges) {
		t.Fatal("original graph was not filtered")
	}
}

func TestWithoutNormalErrRaiseContinuationAlsoFiltersErrorStatement(t *testing.T) {
	t.Parallel()
	doc := buildIR(t, `Public Sub Work()
    Error 5
    Debug.Print "unreachable"
End Sub
`)
	graph := BuildDocument(doc).Graphs[0]
	statementID := doc.Procedures[0].Statements[0].ID
	block, ok := graph.BlockForStatement(statementID)
	if !ok {
		t.Fatal("Error statement block not found")
	}
	for _, edge := range graph.WithoutNormalErrRaiseContinuation().Edges {
		if edge.From == block.ID && edge.Class == EdgeNormal {
			t.Fatalf("Error statement retained normal edge: %+v", edge)
		}
	}
}

func TestBuildDocumentIsDeterministicAndCloneIsDefensive(t *testing.T) {
	t.Parallel()
	doc := buildIR(t, `Public Sub First()
    Call Work
End Sub
Public Sub Second()
End Sub
`)
	first := BuildDocument(doc)
	second := BuildDocument(doc)
	if !reflect.DeepEqual(first, second) {
		t.Fatal("identical IR produced different graphs")
	}
	if len(first.Graphs) != 2 || first.Graphs[0].Procedure.Name != "First" ||
		first.Graphs[1].Procedure.Name != "Second" {
		t.Fatalf("unexpected document graphs: %+v", first.Graphs)
	}
	doc.Procedures[0].Statements[0].Text = "mutated input"
	if first.Graphs[0].Blocks[5].Statement.Text == "mutated input" {
		t.Fatal("BuildDocument retains mutable input storage")
	}
	clone := CloneDocument(first)
	clone.Graphs[0].Blocks[5].Statement.Text = "changed"
	clone.Graphs[0].Edges[0].Kind = EdgeUnknown
	if first.Graphs[0].Blocks[5].Statement.Text == "changed" ||
		first.Graphs[0].Edges[0].Kind == EdgeUnknown {
		t.Fatal("CloneDocument shares graph storage")
	}
	doc.Procedures[0].Symbol.Parameters = append(doc.Procedures[0].Symbol.Parameters,
		procedureir.Parameter{Name: "late"})
	doc.Procedures[0].Statements[0].Text = "mutated input"
	if len(first.Graphs[0].Procedure.Parameters) != 0 ||
		first.Graphs[0].Blocks[5].Statement.Text == "mutated input" {
		t.Fatal("BuildDocument retained mutable input storage")
	}
}

func TestBlockForStatementIndexUsesSlicePositionAndFirstMatch(t *testing.T) {
	t.Parallel()
	graph := Graph{
		Blocks: []Block{
			{ID: 40, Kind: BlockEntry},
			{ID: 90, Kind: BlockStatement, StatementID: 10},
			{ID: 120, Kind: BlockStatement, StatementID: 10},
		},
	}
	graph.query = buildQueryIndex(graph)
	block, ok := graph.BlockForStatement(10)
	if !ok || block.ID != 90 {
		t.Fatalf("BlockForStatement(10) = (%+v, %v), want first sparse-ID block", block, ok)
	}
}

func TestBlockForStatementIndexPreservesFirstMatchAfterEarlierMutation(t *testing.T) {
	t.Parallel()
	graph := Graph{
		Blocks: []Block{
			{ID: 40, Kind: BlockEntry},
			{ID: 90, Kind: BlockStatement, StatementID: 10},
		},
	}
	graph.query = buildQueryIndex(graph)
	graph.Blocks = append([]Block{{ID: 5, Kind: BlockStatement, StatementID: 10}}, graph.Blocks...)
	block, ok := graph.BlockForStatement(10)
	if !ok || block.ID != 5 {
		t.Fatalf("BlockForStatement(10) after earlier mutation = (%+v, %v), want inserted block", block, ok)
	}
}

func TestQueryIndexInvalidatesReachabilityInputs(t *testing.T) {
	t.Parallel()
	graph := Graph{
		Blocks: []Block{
			{ID: 1, Kind: BlockEntry},
			{ID: 2, Kind: BlockStatement, StatementID: 1},
			{ID: 3, Kind: BlockUnknownExit},
			{ID: 4, Kind: BlockStatement, StatementID: 2},
			{ID: 5, Kind: BlockUnknownExit},
		},
		Edges: []Edge{{ID: 1, From: 1, To: 2, Class: EdgeNormal}},
		Entry: 1, UnknownExit: 3,
	}
	graph.query = buildQueryIndex(graph)
	if got, want := graph.Reachable(EdgeFilter{}), []BlockID{1, 2}; !reflect.DeepEqual(got, want) {
		t.Fatalf("initial Reachable() = %v, want %v", got, want)
	}

	graph.Entry = 4
	if got, want := graph.Reachable(EdgeFilter{}), []BlockID{4}; !reflect.DeepEqual(got, want) {
		t.Fatalf("Reachable() after Entry change = %v, want %v", got, want)
	}

	graph.UnknownFlowSources = []BlockID{4}
	if got, want := graph.Reachable(EdgeFilter{}), []BlockID{2, 3, 4}; !reflect.DeepEqual(got, want) {
		t.Fatalf("Reachable() after UnknownFlowSources change = %v, want %v", got, want)
	}

	graph.UnknownExit = 5
	if got, want := graph.Reachable(EdgeFilter{}), []BlockID{2, 4, 5}; !reflect.DeepEqual(got, want) {
		t.Fatalf("Reachable() after UnknownExit change = %v, want %v", got, want)
	}
}

func TestCloneDeepCopiesProcedureSignatureRanges(t *testing.T) {
	defaultRange := vbaast.Range{StartLine: 1, EndLine: 1, StartByte: 2, EndByte: 4}
	boundsRange := vbaast.Range{StartLine: 2, EndLine: 2, StartByte: 5, EndByte: 8}
	lowerRange := vbaast.Range{StartLine: 3, EndLine: 3, StartByte: 9, EndByte: 10}
	upperRange := vbaast.Range{StartLine: 3, EndLine: 3, StartByte: 11, EndByte: 12}
	procedureRange := vbaast.Range{StartLine: 1, EndLine: 4, StartByte: 0, EndByte: 20}
	graph := Graph{Procedure: procedureir.ProcedureSymbol{
		Parameters: []procedureir.Parameter{{
			DefaultRange: &defaultRange,
			BoundsRange:  &boundsRange,
			ArrayBounds:  []procedureir.ArrayBound{{LowerRange: &lowerRange, UpperRange: &upperRange}},
		}},
		ArrayBounds:      []procedureir.ArrayBound{{LowerRange: &lowerRange, UpperRange: &upperRange}},
		DeclarationRange: procedureRange,
	}}
	wantDefaultStart := graph.Procedure.Parameters[0].DefaultRange.StartByte
	wantBoundsEnd := graph.Procedure.Parameters[0].BoundsRange.EndByte
	wantParameterLowerStart := graph.Procedure.Parameters[0].ArrayBounds[0].LowerRange.StartByte
	wantParameterUpperEnd := graph.Procedure.Parameters[0].ArrayBounds[0].UpperRange.EndByte
	wantProcedureLowerStart := graph.Procedure.ArrayBounds[0].LowerRange.StartByte
	wantProcedureUpperEnd := graph.Procedure.ArrayBounds[0].UpperRange.EndByte
	clone := Clone(graph)
	clone.Procedure.Parameters[0].DefaultRange.StartByte++
	clone.Procedure.Parameters[0].BoundsRange.EndByte++
	clone.Procedure.Parameters[0].ArrayBounds[0].LowerRange.StartByte++
	clone.Procedure.Parameters[0].ArrayBounds[0].UpperRange.EndByte++
	clone.Procedure.ArrayBounds[0].LowerRange.StartByte++
	clone.Procedure.ArrayBounds[0].UpperRange.EndByte++
	if graph.Procedure.Parameters[0].DefaultRange.StartByte != wantDefaultStart ||
		graph.Procedure.Parameters[0].BoundsRange.EndByte != wantBoundsEnd ||
		graph.Procedure.Parameters[0].ArrayBounds[0].LowerRange.StartByte != wantParameterLowerStart ||
		graph.Procedure.Parameters[0].ArrayBounds[0].UpperRange.EndByte != wantParameterUpperEnd ||
		graph.Procedure.ArrayBounds[0].LowerRange.StartByte != wantProcedureLowerStart ||
		graph.Procedure.ArrayBounds[0].UpperRange.EndByte != wantProcedureUpperEnd {
		t.Fatal("Clone shares procedure signature range storage")
	}
}

func TestStructuredBranchesLoopsTransfersAndLabels(t *testing.T) {
	t.Parallel()
	procedure := buildProcedure(t, `Public Sub Run()
    If Ready Then
        For Each item In items
            If Skip Then GoTo Done
            If StopNow Then Exit For
            Call ContinueWork
        Next
    Else
        Call Other
    End If
    Do
        Call Again
    Loop While Ready
    Exit Sub
Done:
    End
End Sub
`)
	graph := Build(procedure)
	ifEdge := requireEdgeFromKind(t, graph, procedureir.StatementIf, EdgeBranchTrue)
	if ifEdge.Class != EdgeNormal || ifEdge.Uncertain {
		t.Fatalf("if edge = %+v", ifEdge)
	}
	requireEdge(t, graph, EdgeBranchFalse)
	requireEdge(t, graph, EdgeLoopBody)
	requireEdge(t, graph, EdgeLoopBack)
	requireEdge(t, graph, EdgeLoopExit)
	gotoEdge := requireEdge(t, graph, EdgeGoto)
	if graph.block(gotoEdge.To).Statement == nil ||
		graph.block(gotoEdge.To).Statement.Kind != procedureir.StatementLabel {
		t.Fatalf("goto did not resolve to label: %+v", gotoEdge)
	}
	requireTransitionTo(t, graph, EdgeProcedureExit, graph.NormalExit)
	requireTransitionTo(t, graph, EdgeTermination, graph.TerminationExit)
}

func TestSelectAndLoopFormsExposeConservativeAlternatives(t *testing.T) {
	t.Parallel()
	document := buildIR(t, `Public Sub Selection(ByVal value As Long)
    Select Case value
    Case 1
        Call One
    Case Else
        Call Other
    End Select
End Sub

Public Sub NoDefault(ByVal value As Long)
    Select Case value
    Case 1
        Call One
    End Select
End Sub

Public Sub Loops(ByVal ready As Boolean)
    Dim i As Long
    For i = 1 To 2
        Call Work
    Next
    While ready
        Call Work
    Wend
    Do While ready
        Call Work
    Loop
    Do Until ready
        Call Work
    Loop
    Do
        Call Work
    Loop While ready
    Do
        Call Work
    Loop Until ready
End Sub
`)
	if len(document.Procedures) != 3 {
		t.Fatalf("procedures = %d, want 3", len(document.Procedures))
	}
	selection := Build(document.Procedures[0])
	if got := countEdges(selection, EdgeCase); got != 2 {
		t.Fatalf("Select Case edges = %d, want 2; edges=%+v", got, selection.Edges)
	}
	if got := countEdges(selection, EdgeBranchFalse); got != 0 {
		t.Fatalf("Select Case Else should cover the default path, false edges=%d", got)
	}
	noDefault := Build(document.Procedures[1])
	if got := countEdges(noDefault, EdgeBranchFalse); got != 1 {
		t.Fatalf("Select without Case Else false edges = %d, want 1", got)
	}
	loops := Build(document.Procedures[2])
	if got := countEdges(loops, EdgeLoopBack); got < 6 {
		t.Fatalf("loop-back edges = %d, want at least one per loop; edges=%+v", got, loops.Edges)
	}
	if got := countEdges(loops, EdgeLoopExit); got < 6 {
		t.Fatalf("loop-exit edges = %d, want at least one per loop; edges=%+v", got, loops.Edges)
	}
	for _, block := range loops.Blocks {
		if block.Statement == nil || block.Statement.Kind != procedureir.StatementDo ||
			block.Statement.SyntaxKind == "do_condition" {
			continue
		}
		if block.Statement.Control == nil || block.Statement.Control.Loop == "" {
			t.Fatalf("Do block lacks normalized loop mode: %+v", block.Statement)
		}
	}
}

func TestElseIfTrueBranchSkipsFollowingElse(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run(ByVal value As Long)
    If value = 1 Then
        Call One
    ElseIf value = 2 Then
        Call Two
    ElseIf value = 3 Then
        Call Three
    Else
        Call Other
    End If
    Call Afterward
End Sub
`))
	var firstElseIf, secondElseIf, elseBlock Block
	for _, block := range graph.Blocks {
		if block.Statement == nil {
			continue
		}
		switch {
		case block.Statement.Kind == procedureir.StatementElseIf &&
			strings.Contains(block.Statement.Text, "value = 2"):
			firstElseIf = block
		case block.Statement.Kind == procedureir.StatementElseIf &&
			strings.Contains(block.Statement.Text, "value = 3"):
			secondElseIf = block
		case block.Statement.Kind == procedureir.StatementElse:
			elseBlock = block
		}
	}
	if firstElseIf.ID == 0 || secondElseIf.ID == 0 || elseBlock.ID == 0 {
		t.Fatalf("conditional chain blocks missing: first=%+v second=%+v else=%+v", firstElseIf, secondElseIf, elseBlock)
	}
	twoBlock := blockID(t, graph, statementIDContaining(t, graph, "Call Two"))
	threeBlock := blockID(t, graph, statementIDContaining(t, graph, "Call Three"))
	otherBlock := blockID(t, graph, statementIDContaining(t, graph, "Call Other"))
	afterBlock := blockID(t, graph, statementIDContaining(t, graph, "Call Afterward"))
	if !hasEdge(graph, firstElseIf.ID, twoBlock, EdgeBranchTrue) {
		t.Fatalf("ElseIf true edge does not enter its body: %+v", graph.Edges)
	}
	if !hasEdge(graph, firstElseIf.ID, secondElseIf.ID, EdgeBranchFalse) {
		t.Fatalf("first ElseIf false edge does not enter second ElseIf: %+v", graph.Edges)
	}
	if !hasEdge(graph, secondElseIf.ID, threeBlock, EdgeBranchTrue) {
		t.Fatalf("second ElseIf true edge does not enter its body: %+v", graph.Edges)
	}
	if !hasEdge(graph, secondElseIf.ID, elseBlock.ID, EdgeBranchFalse) {
		t.Fatalf("second ElseIf false edge does not enter Else: %+v", graph.Edges)
	}
	if !hasEdge(graph, twoBlock, afterBlock, EdgeFallthrough) {
		t.Fatalf("ElseIf body does not rejoin after the If: %+v", graph.Edges)
	}
	if !hasEdge(graph, threeBlock, afterBlock, EdgeFallthrough) {
		t.Fatalf("second ElseIf body does not rejoin after the If: %+v", graph.Edges)
	}
	if hasEdge(graph, twoBlock, otherBlock, EdgeFallthrough) {
		t.Fatalf("ElseIf true body incorrectly falls through into Else: %+v", graph.Edges)
	}
}

func TestEmptyLoopsRetainBackEdges(t *testing.T) {
	t.Parallel()
	document := buildIR(t, `Public Sub EmptyFor()
    For i = 1 To 2
    Next
End Sub

Public Sub EmptyWhile(ByVal ready As Boolean)
    While ready
    Wend
End Sub

Public Sub EmptyPreDo(ByVal ready As Boolean)
    Do While ready
    Loop
End Sub

Public Sub EmptyPostDo(ByVal ready As Boolean)
    Do
    Loop Until ready
End Sub

Public Sub EmptyInfiniteDo()
    Do
    Loop
End Sub
`)
	if len(document.Procedures) != 5 {
		t.Fatalf("procedures = %d, want 5", len(document.Procedures))
	}
	for _, procedure := range document.Procedures {
		graph := Build(procedure)
		if got := countEdges(graph, EdgeLoopBack); got == 0 {
			t.Fatalf("%s has no loop-back edge: %+v", procedure.Symbol.Name, graph.Edges)
		}
	}
}

func TestDuplicateAndMissingLabelsRemainConservative(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run()
    GoTo Duplicate
Duplicate:
    Call One
Duplicate:
    GoTo Missing
End Sub
`))
	var duplicateEdges int
	for _, edge := range graph.Edges {
		if edge.Kind == EdgeGoto && edge.Uncertain {
			duplicateEdges++
		}
	}
	if duplicateEdges != 2 {
		t.Fatalf("duplicate target edges = %d, want 2; edges=%+v", duplicateEdges, graph.Edges)
	}
	requireTransitionTo(t, graph, EdgeUnknown, graph.UnknownExit)
}

func TestInvalidExitTargetsUnknownFlow(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name      string
		procedure procedureir.ProcedureIR
	}{
		{
			name: "mismatched procedure exit",
			procedure: procedureir.ProcedureIR{
				Symbol: procedureir.ProcedureSymbol{Name: "Value", Kind: procedureir.ProcedureFunction},
				Statements: []procedureir.Statement{{
					ID: 1, Kind: procedureir.StatementExit,
					Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitSub},
				}},
			},
		},
		{
			name: "loop exit outside loop",
			procedure: procedureir.ProcedureIR{
				Symbol: procedureir.ProcedureSymbol{Name: "Run", Kind: procedureir.ProcedureSub},
				Statements: []procedureir.Statement{{
					ID: 1, Kind: procedureir.StatementExit,
					Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitFor},
				}},
			},
		},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			graph := Build(test.procedure)
			found := false
			for _, edge := range graph.Edges {
				if edge.From == blockID(t, graph, 1) && edge.To == graph.UnknownExit && edge.Uncertain {
					found = true
				}
			}
			if !found {
				t.Fatalf("invalid exit lacks uncertain unknown flow: %+v", graph.Edges)
			}
		})
	}
}

func TestForEachZeroIterationDoesNotAssignIterator(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run(ByVal items As Collection)
    Dim item As Variant
    For Each item In items
    Next
    Debug.Print item
End Sub
`))
	printBlock := blockID(t, graph, statementIDContaining(t, graph, "Debug.Print"))
	if graph.IsDefinitelyAssigned(
		printBlock,
		Variable{Scope: procedureir.ScopeLocal, Name: "item"},
		EdgeFilter{NormalOnly: true},
	) {
		t.Fatal("For Each iterator is not definitely assigned after a possible zero-iteration loop")
	}
}

func TestErrorModesProduceSeparateUncertainExceptionalEdges(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run()
    On Error GoTo Handler
    Call Work
    On Error Resume Next
    Call IgnoreFailure
    On Error GoTo 0
    Call Fail
    Exit Sub
Handler:
    Resume Next
End Sub
`))
	var handler, resumeNext, exceptional bool
	for _, edge := range graph.Edges {
		if edge.Kind != EdgeError {
			continue
		}
		if edge.Class != EdgeExceptional || !edge.Uncertain {
			t.Fatalf("error edge is not uncertain exceptional: %+v", edge)
		}
		source := graph.block(edge.From).Statement
		target := graph.block(edge.To).Statement
		if source != nil && strings.Contains(source.Text, "Work") && target != nil &&
			target.Kind == procedureir.StatementLabel {
			handler = true
		}
		if source != nil && strings.Contains(source.Text, "IgnoreFailure") && edge.To != graph.ExceptionalExit {
			resumeNext = true
		}
		if source != nil && strings.Contains(source.Text, "Fail") && edge.To == graph.ExceptionalExit {
			exceptional = true
		}
	}
	if !handler || !resumeNext || !exceptional {
		t.Fatalf("missing error mode edge: handler=%v next=%v exceptional=%v; edges=%+v",
			handler, resumeNext, exceptional, graph.Edges)
	}
	resume := requireEdge(t, graph, EdgeResume)
	if resume.Class != EdgeExceptional || !resume.Uncertain || resume.To != graph.UnknownExit {
		t.Fatalf("dynamic Resume Next edge = %+v, want uncertain exceptional unknown flow", resume)
	}
}

func TestMergedErrorModesAndBackwardGotoRemainExplicit(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run(ByVal guarded As Boolean, ByVal repeat As Boolean)
Start:
    If guarded Then
        On Error GoTo Handler
    Else
        On Error Resume Next
    End If
    Call Work
    If repeat Then GoTo Start
    Exit Sub
Handler:
    Resume Next
End Sub
`))
	workBlock := blockID(t, graph, statementIDContaining(t, graph, "Call Work"))
	handlerBlock := blockID(t, graph, statementIDContaining(t, graph, "Handler:"))
	var handlerMode, resumeNextMode bool
	for _, edge := range graph.Edges {
		if edge.From != workBlock || edge.Kind != EdgeError || edge.Class != EdgeExceptional || !edge.Uncertain {
			continue
		}
		handlerMode = handlerMode || edge.To == handlerBlock
		resumeNextMode = resumeNextMode || edge.To != handlerBlock
	}
	if !handlerMode || !resumeNextMode {
		t.Fatalf("merged error modes missing: handler=%v resume-next=%v edges=%+v",
			handlerMode, resumeNextMode, graph.Edges)
	}
	gotoEdge := requireEdge(t, graph, EdgeGoto)
	if gotoEdge.To != blockID(t, graph, statementIDContaining(t, graph, "Start:")) {
		t.Fatalf("backward GoTo target = %d, want Start label", gotoEdge.To)
	}
}

func TestRecoveredLabelTargetAlsoRetainsUnknownFlow(t *testing.T) {
	t.Parallel()
	graph := Build(procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{Name: "Run", Kind: procedureir.ProcedureSub},
		Statements: []procedureir.Statement{
			{
				ID: 1, Kind: procedureir.StatementGoTo,
				Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferGoto, Target: "Handler"},
			},
			{ID: 2, Kind: procedureir.StatementLabel, Label: "Handler", Recovered: true},
		},
	})
	gotoBlock := blockID(t, graph, 1)
	labelBlock := blockID(t, graph, 2)
	if !hasUncertainEdge(graph, gotoBlock, labelBlock) ||
		!hasUncertainEdge(graph, gotoBlock, graph.UnknownExit) {
		t.Fatalf("recovered target lacks candidate and unknown flow: %+v", graph.Edges)
	}
}

func TestValidUnclassifiedStatementKeepsOrdinaryFallthrough(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run(ByVal filePath As String)
    Open filePath For Input As #1
    GoTo Done
Handler:
    Call Failed
Done:
    Exit Sub
End Sub
`))
	unknownID := statementIDContaining(t, graph, "Open filePath")
	unknownBlock := graph.block(blockID(t, graph, unknownID))
	if unknownBlock.Statement == nil || unknownBlock.Statement.Kind != procedureir.StatementUnknown {
		t.Fatalf("Open statement kind = %+v, want valid unclassified statement", unknownBlock.Statement)
	}
	unknown := unknownBlock.ID
	gotoBlock := blockID(t, graph, statementIDContaining(t, graph, "GoTo Done"))
	handler := blockID(t, graph, statementIDContaining(t, graph, "Handler:"))
	if !hasEdge(graph, unknown, gotoBlock, EdgeFallthrough) {
		t.Fatalf("valid unclassified statement lost fallthrough: %+v", graph.Edges)
	}
	if len(graph.UnknownFlowSources) != 0 {
		t.Fatalf("valid unclassified statement became unknown flow: %+v", graph.UnknownFlowSources)
	}
	if graph.IsReachable(handler, EdgeFilter{NormalOnly: true}) {
		t.Fatal("valid unclassified statement made skipped handler reachable")
	}
}

func TestUnclassifiedStructuredChildrenRemainConservativelyReachable(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name       string
		statements []procedureir.Statement
		parentID   int
		childID    int
	}{
		{
			name: "if",
			statements: []procedureir.Statement{
				{ID: 1, Kind: procedureir.StatementIf},
				{
					ID: 2, ParentID: 1, Kind: procedureir.StatementCall,
					Control: &procedureir.ControlFlowMetadata{Branch: procedureir.BranchThen},
				},
				{
					ID: 3, ParentID: 1, Kind: procedureir.StatementCall,
					Control: &procedureir.ControlFlowMetadata{Branch: procedureir.BranchElse},
				},
				{ID: 4, ParentID: 1, Kind: procedureir.StatementRecovered, Recovered: true},
				{ID: 5, Kind: procedureir.StatementExit, Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitSub}},
			},
			parentID: 1,
			childID:  4,
		},
		{
			name: "select",
			statements: []procedureir.Statement{
				{ID: 1, Kind: procedureir.StatementSelect},
				{ID: 2, ParentID: 1, Kind: procedureir.StatementCase},
				{ID: 3, ParentID: 1, Kind: procedureir.StatementRecovered, Recovered: true},
				{ID: 4, Kind: procedureir.StatementExit, Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitSub}},
			},
			parentID: 1,
			childID:  3,
		},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			graph := Build(procedureir.ProcedureIR{
				Symbol:     procedureir.ProcedureSymbol{Name: "Run", Kind: procedureir.ProcedureSub},
				Statements: test.statements,
			})
			parent := blockID(t, graph, test.parentID)
			child := blockID(t, graph, test.childID)
			if !hasUncertainEdge(graph, parent, child) {
				t.Fatalf("unclassified child lacks local uncertain entry: %+v", graph.Edges)
			}
			if !graph.IsReachable(child, EdgeFilter{NormalOnly: true}) {
				t.Fatal("unclassified structured child is spuriously unreachable")
			}
		})
	}
}

func TestExactResumeTargetRetainsEnabledErrorMode(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run()
    On Error GoTo Handler
    Call First
    Exit Sub
Handler:
    Resume ContinueHere
ContinueHere:
    Call Second
    Exit Sub
End Sub
`))
	handler := blockID(t, graph, statementIDContaining(t, graph, "Handler:"))
	resume := blockID(t, graph, statementIDContaining(t, graph, "Resume ContinueHere"))
	continuation := blockID(t, graph, statementIDContaining(t, graph, "ContinueHere:"))
	second := blockID(t, graph, statementIDContaining(t, graph, "Call Second"))
	if !hasEdge(graph, resume, continuation, EdgeResume) {
		t.Fatalf("exact Resume target missing: %+v", graph.Edges)
	}
	found := false
	for _, edge := range graph.Edges {
		if edge.From == second && edge.To == handler && edge.Kind == EdgeError &&
			edge.Class == EdgeExceptional && edge.Uncertain {
			found = true
		}
	}
	if !found {
		t.Fatalf("post-Resume fault lost enabled handler mode: %+v", graph.Edges)
	}
}

func TestRecoveredFlowUsesLinearMarkersAndPreWriteAssignmentState(t *testing.T) {
	t.Parallel()
	const statementCount = 100
	statements := make([]procedureir.Statement, statementCount)
	for i := range statements {
		statements[i] = procedureir.Statement{
			ID: i + 1, Kind: procedureir.StatementRecovered, Recovered: true,
		}
	}
	procedure := procedureir.ProcedureIR{
		Symbol:     procedureir.ProcedureSymbol{Name: "Recovered", Kind: procedureir.ProcedureSub},
		Statements: statements,
		Accesses: []procedureir.VariableAccess{{
			Name: "value", Mode: procedureir.AccessWrite, Scope: procedureir.ScopeLocal, StatementID: 1,
		}},
	}
	graph := Build(procedure)
	if len(graph.UnknownFlowSources) != statementCount {
		t.Fatalf("unknown-flow sources = %d, want %d", len(graph.UnknownFlowSources), statementCount)
	}
	if got := countEdges(graph, EdgeUnknown); got != statementCount {
		t.Fatalf("unknown edges = %d, want linear count %d", got, statementCount)
	}
	if graph.IsDefinitelyAssigned(
		blockID(t, graph, 2),
		Variable{Scope: procedureir.ScopeLocal, Name: "value"},
		EdgeFilter{NormalOnly: true},
	) {
		t.Fatal("recovered assignment write propagated after unknown divergence")
	}
}

func TestUnknownFlowPreservesPreDivergenceFactsWithoutInventingSyntheticExits(t *testing.T) {
	t.Parallel()
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{Name: "Recovered", Kind: procedureir.ProcedureSub},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementAssignment},
			{ID: 2, Kind: procedureir.StatementRecovered, Recovered: true},
			{ID: 3, Kind: procedureir.StatementLabel, Label: "Cleanup"},
		},
		Accesses: []procedureir.VariableAccess{{
			Name: "value", Mode: procedureir.AccessWrite, Scope: procedureir.ScopeLocal, StatementID: 1,
		}},
	}
	graph := Build(procedure)
	filter := EdgeFilter{NormalOnly: true}
	assignment := blockID(t, graph, 1)
	cleanup := blockID(t, graph, 3)
	value := Variable{Scope: procedureir.ScopeLocal, Name: "value"}
	if !graph.IsDefinitelyAssigned(cleanup, value, filter) {
		t.Fatal("assignment completed before unknown divergence was discarded")
	}
	if !containsBlock(graph.Dominators(filter)[cleanup], assignment) {
		t.Fatal("pre-divergence assignment should still dominate the cleanup label")
	}
	if graph.IsReachable(graph.TerminationExit, filter) {
		t.Fatal("unknown flow invented a termination exit without an End statement")
	}
	if !graph.CleanupGuaranteed([]int{3}, ExitSelection{Normal: true}, filter) {
		t.Fatal("unknown flow invented a normal path around the terminal cleanup label")
	}
}

func TestReachabilityDominanceDefiniteAssignmentAndCleanup(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run(ByVal Ready As Boolean)
    Dim value As Long
    If Ready Then
        value = 1
    Else
        value = 2
    End If
Cleanup:
    Debug.Print value
    Exit Sub
Dead:
    value = 3
End Sub
`))
	dead := statementIDContaining(t, graph, "value = 3")
	if graph.IsReachable(blockID(t, graph, dead), EdgeFilter{NormalOnly: true}) {
		t.Fatal("dead statement is reachable")
	}
	if len(graph.Unreachable(EdgeFilter{NormalOnly: true})) == 0 {
		t.Fatal("expected unreachable statement")
	}
	cleanup := statementIDContaining(t, graph, "Cleanup:")
	printID := statementIDContaining(t, graph, "Debug.Print")
	printBlock := blockID(t, graph, printID)
	variable := Variable{Scope: procedureir.ScopeLocal, Name: "VALUE"}
	if !graph.IsDefinitelyAssigned(printBlock, variable, EdgeFilter{NormalOnly: true}) {
		t.Fatalf("value is not definitely assigned at print: %+v", graph.DefiniteAssignments(EdgeFilter{NormalOnly: true})[printBlock])
	}
	if !graph.IsDefinitelyAssigned(graph.Entry,
		Variable{Scope: procedureir.ScopeParameter, Name: "READY"}, EdgeFilter{}) {
		t.Fatal("procedure parameter is not seeded as definitely assigned")
	}
	dominators := graph.Dominators(EdgeFilter{NormalOnly: true})
	if !containsBlock(dominators[graph.NormalExit], blockID(t, graph, cleanup)) {
		t.Fatalf("cleanup does not dominate normal exit: %+v", dominators[graph.NormalExit])
	}
	if !graph.CleanupGuaranteed([]int{cleanup}, ExitSelection{Normal: true}, EdgeFilter{NormalOnly: true}) {
		t.Fatal("cleanup should cross every normal exit path")
	}
	if graph.CleanupGuaranteed([]int{cleanup}, ExitSelection{}, EdgeFilter{}) {
		t.Fatal("cleanup must not be guaranteed across conservative exceptional exits")
	}
	if len(graph.ExitTransitions(ExitSelection{Normal: true}, EdgeFilter{NormalOnly: true})) == 0 {
		t.Fatal("normal exit transitions missing")
	}
	if block, ok := graph.BlockForStatement(0); ok {
		t.Fatalf("statement ID 0 resolved to synthetic block: %+v", block)
	}
	if graph.CleanupGuaranteed([]int{0}, ExitSelection{Normal: true}, EdgeFilter{NormalOnly: true}) {
		t.Fatal("invalid cleanup statement ID proved a guarantee")
	}
}

func TestMultipleCleanupCandidatesCoverNormalExits(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run(ByVal first As Boolean)
    If first Then
CleanupA:
        Exit Sub
    Else
CleanupB:
        Exit Sub
    End If
End Sub
`))
	cleanupA := statementIDContaining(t, graph, "CleanupA:")
	cleanupB := statementIDContaining(t, graph, "CleanupB:")
	if !graph.CleanupGuaranteed(
		[]int{cleanupA, cleanupB},
		ExitSelection{Normal: true},
		EdgeFilter{NormalOnly: true},
	) {
		t.Fatal("one of the two cleanup labels crosses every normal exit path")
	}
}

func TestCFGRangesPreserveUTF8ByteColumnsAndCRLFOffsets(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, "Public Sub Run()\r\n"+
		"    Debug.Print \"日本語\": GoTo Done\r\n"+
		"Done:\r\n"+
		"End Sub\r\n"))
	gotoID := statementIDContaining(t, graph, "GoTo Done")
	gotoBlock, ok := graph.BlockForStatement(gotoID)
	if !ok || gotoBlock.Statement == nil {
		t.Fatal("GoTo block missing")
	}
	gotoEdge := requireEdge(t, graph, EdgeGoto)
	if gotoEdge.Range != gotoBlock.Statement.Range {
		t.Fatalf("edge range = %+v, statement range = %+v", gotoEdge.Range, gotoBlock.Statement.Range)
	}
	if gotoEdge.Range.StartLine != 2 || gotoEdge.Range.StartColumn <= 25 || gotoEdge.Range.StartByte <= 20 {
		t.Fatalf("UTF-8/CRLF range was not preserved: %+v", gotoEdge.Range)
	}
}

func TestExceptionalAssignmentEdgeUsesBlockInputSet(t *testing.T) {
	t.Parallel()
	graph := Build(buildProcedure(t, `Public Sub Run()
    Dim value As Long
    On Error GoTo Handler
    value = Work()
    Exit Sub
Handler:
    Debug.Print value
End Sub
`))
	handler := blockID(t, graph, statementIDContaining(t, graph, "Handler:"))
	if graph.IsDefinitelyAssigned(handler,
		Variable{Scope: procedureir.ScopeLocal, Name: "value"}, EdgeFilter{}) {
		t.Fatal("faulting assignment write propagated across exceptional edge")
	}
}

func TestUnknownRecoveryPreventsPositiveGuarantees(t *testing.T) {
	t.Parallel()
	procedure := procedureir.ProcedureIR{
		Symbol: procedureir.ProcedureSymbol{Name: "Recovered", Kind: procedureir.ProcedureSub},
		Statements: []procedureir.Statement{
			{ID: 1, Kind: procedureir.StatementRecovered, Recovered: true},
			{ID: 2, Kind: procedureir.StatementLabel, Label: "Cleanup"},
			{ID: 3, Kind: procedureir.StatementExit, Control: &procedureir.ControlFlowMetadata{Transfer: procedureir.TransferExitSub}},
		},
	}
	graph := Build(procedure)
	if graph.CleanupGuaranteed([]int{2}, ExitSelection{}, EdgeFilter{}) {
		t.Fatal("recovered control incorrectly proved cleanup")
	}
	unknown := false
	for _, edge := range graph.Edges {
		unknown = unknown || edge.From == blockID(t, graph, 1) && edge.Kind == EdgeUnknown &&
			edge.Uncertain && edge.To == graph.UnknownExit
	}
	if !unknown {
		t.Fatal("recovered statement lacks uncertain unknown transition")
	}
}

func buildIR(t *testing.T, source string) procedureir.DocumentIR {
	t.Helper()
	doc, err := procedureir.BuildSource(procedureir.BuildOptions{Path: "Module1.bas"}, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	return doc
}

func buildProcedure(t *testing.T, source string) procedureir.ProcedureIR {
	t.Helper()
	doc := buildIR(t, source)
	if len(doc.Procedures) != 1 {
		t.Fatalf("procedures = %d", len(doc.Procedures))
	}
	return doc.Procedures[0]
}

func requireEdge(t *testing.T, graph Graph, kind EdgeKind) Edge {
	t.Helper()
	for _, edge := range graph.Edges {
		if edge.Kind == kind {
			return edge
		}
	}
	t.Fatalf("missing %q edge: %+v", kind, graph.Edges)
	return Edge{}
}

func requireEdgeFromKind(t *testing.T, graph Graph, statementKind procedureir.StatementKind, edgeKind EdgeKind) Edge {
	t.Helper()
	for _, edge := range graph.Edges {
		block := graph.block(edge.From)
		if edge.Kind == edgeKind && block.Statement != nil && block.Statement.Kind == statementKind {
			return edge
		}
	}
	t.Fatalf("missing %q edge from %q", edgeKind, statementKind)
	return Edge{}
}

func requireTransitionTo(t *testing.T, graph Graph, kind EdgeKind, target BlockID) Edge {
	t.Helper()
	for _, edge := range graph.Edges {
		if edge.Kind == kind && edge.To == target {
			return edge
		}
	}
	t.Fatalf("missing %q edge to %d: %+v", kind, target, graph.Edges)
	return Edge{}
}

func statementIDContaining(t *testing.T, graph Graph, text string) int {
	t.Helper()
	for _, block := range graph.Blocks {
		if block.Statement != nil && strings.Contains(block.Statement.Text, text) {
			return block.StatementID
		}
	}
	t.Fatalf("missing statement containing %q", text)
	return 0
}

func blockID(t *testing.T, graph Graph, statementID int) BlockID {
	t.Helper()
	block, ok := graph.BlockForStatement(statementID)
	if !ok {
		t.Fatalf("missing statement block %d", statementID)
	}
	return block.ID
}

func containsBlock(ids []BlockID, target BlockID) bool {
	for _, id := range ids {
		if id == target {
			return true
		}
	}
	return false
}

func countEdges(graph Graph, kind EdgeKind) int {
	count := 0
	for _, edge := range graph.Edges {
		if edge.Kind == kind {
			count++
		}
	}
	return count
}

func hasEdge(graph Graph, from, to BlockID, kind EdgeKind) bool {
	for _, edge := range graph.Edges {
		if edge.From == from && edge.To == to && edge.Kind == kind {
			return true
		}
	}
	return false
}

func hasUncertainEdge(graph Graph, from, to BlockID) bool {
	for _, edge := range graph.Edges {
		if edge.From == from && edge.To == to && edge.Uncertain {
			return true
		}
	}
	return false
}
