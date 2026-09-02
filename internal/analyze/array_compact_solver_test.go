package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestArrayCompactAdapterDegradesAbsentSymbolsToUnknown(t *testing.T) {
	initial := arrayFlowState{
		"allocated":   {kind: arrayAllocated, knownArray: true},
		"unallocated": {kind: arrayUnallocated, knownArray: true},
	}
	adapter := newArrayCompactAdapter(initial)
	left := semanticstate.NewState[arrayValue](adapter.environment.Layout())
	incoming := semanticstate.NewState[arrayValue](adapter.environment.Layout())
	adapter.fromFlow(&left, initial)
	adapter.fromFlow(&incoming, arrayFlowState{
		"allocated": initial["allocated"],
	})

	changed := make([]semanticstate.SymbolID, 0, 1)
	if !left.JoinFrom(incoming.View(), arrayCompactLattice{}, &changed) {
		t.Fatal("missing incoming symbol should degrade the destination state")
	}
	flow := adapter.toFlow(left.View())
	if got := flow["unallocated"]; got.kind != arrayUnknown {
		t.Fatalf("absent incoming symbol = %#v, want unknown", got)
	}
	if got := flow["allocated"]; got.kind != arrayAllocated || !got.knownArray {
		t.Fatalf("unchanged symbol = %#v, want allocated", got)
	}

	destinationOnly := initial["allocated"]
	if !(arrayCompactLattice{}).Join(&destinationOnly, unknownArrayValue()) || destinationOnly.kind != arrayUnknown {
		t.Fatalf("unknown join = %#v, want unknown", destinationOnly)
	}
}

func TestArrayCompactAdapterIndexesOnlyCurrentCFGAssignments(t *testing.T) {
	statement := &procedureir.Statement{ID: 1, Kind: procedureir.StatementAssignment, Text: "values = Array(1)", Range: vbaast.Range{StartLine: 1, EndLine: 1}}
	graph := vbacfg.Graph{
		Blocks: []vbacfg.Block{
			{ID: 0, Kind: vbacfg.BlockEntry},
			{ID: 1, Kind: vbacfg.BlockStatement, StatementID: statement.ID, Statement: statement},
			{ID: 2, Kind: vbacfg.BlockNormalExit},
		},
		Edges: []vbacfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal},
			{ID: 1, From: 1, To: 2, Kind: vbacfg.EdgeProcedureExit, Class: vbacfg.EdgeNormal},
		},
		Entry: 0, NormalExit: 2, ExceptionalExit: 2, TerminationExit: 2, UnknownExit: 2,
	}
	view := graph.View(vbacfg.EdgeFilter{})
	adapter := newArrayCompactAdapterForLines(&view, []string{"values = Array(1)", "unrelated = Array(2)"}, arrayFlowState{
		"values": {kind: arrayUnknown, knownArray: true},
	})
	if _, ok := adapter.environment.Symbol("values"); !ok {
		t.Fatal("current CFG assignment target was not indexed")
	}
	if _, ok := adapter.environment.Symbol("unrelated"); ok {
		t.Fatal("assignment from an unrelated procedure was indexed")
	}
}

func TestArrayCompactAdvancedPreservesSourceLinesAndEdgeSemantics(t *testing.T) {
	redim := &procedureir.Statement{ID: 1, Kind: procedureir.StatementReDim, Text: "ReDim values(0)", Range: vbaast.Range{StartLine: 1, EndLine: 1}}
	normal := &procedureir.Statement{ID: 2, Kind: procedureir.StatementCall, Text: "normal", Range: vbaast.Range{StartLine: 2, EndLine: 2}}
	exceptional := &procedureir.Statement{ID: 3, Kind: procedureir.StatementCall, Text: "exceptional", Range: vbaast.Range{StartLine: 3, EndLine: 3}}
	uncertain := &procedureir.Statement{ID: 4, Kind: procedureir.StatementCall, Text: "uncertain", Range: vbaast.Range{StartLine: 4, EndLine: 4}}
	graph := vbacfg.Graph{
		Blocks: []vbacfg.Block{
			{ID: 0, Kind: vbacfg.BlockEntry},
			{ID: 1, Kind: vbacfg.BlockStatement, StatementID: redim.ID, Statement: redim},
			{ID: 2, Kind: vbacfg.BlockStatement, StatementID: normal.ID, Statement: normal},
			{ID: 3, Kind: vbacfg.BlockStatement, StatementID: exceptional.ID, Statement: exceptional},
			{ID: 4, Kind: vbacfg.BlockStatement, StatementID: uncertain.ID, Statement: uncertain},
			{ID: 5, Kind: vbacfg.BlockNormalExit},
		},
		Edges: []vbacfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal},
			{ID: 1, From: 1, To: 2, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal},
			{ID: 2, From: 1, To: 3, Kind: vbacfg.EdgeError, Class: vbacfg.EdgeExceptional},
			{ID: 3, From: 1, To: 4, Kind: vbacfg.EdgeBranchFalse, Class: vbacfg.EdgeNormal, Uncertain: true},
			{ID: 4, From: 2, To: 5, Kind: vbacfg.EdgeProcedureExit, Class: vbacfg.EdgeNormal},
			{ID: 5, From: 3, To: 5, Kind: vbacfg.EdgeProcedureExit, Class: vbacfg.EdgeNormal},
			{ID: 6, From: 4, To: 5, Kind: vbacfg.EdgeProcedureExit, Class: vbacfg.EdgeNormal},
		},
		Entry: 0, NormalExit: 5, ExceptionalExit: 5, TerminationExit: 5, UnknownExit: 5,
	}
	view := graph.View(vbacfg.EdgeFilter{})
	initial := arrayFlowState{"values": {kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}}
	seen := map[string]arrayValue{}
	var edgeCalls []vbacfg.Edge
	if err := walkArrayCFGCompactAdvanced(t.Context(), &view, []string{"ReDim values(0)", "normal", "exceptional", "uncertain"}, initial,
		func(text string, line int, in arrayFlowState) arrayFlowState {
			if line > 1 {
				seen[text] = in["values"]
			}
			if text == "ReDim values(0)" {
				value := in["values"]
				value.kind = arrayAllocated
				in["values"] = value
			}
			return in
		},
		func(block vbacfg.Block, edge vbacfg.Edge, out arrayFlowState) arrayFlowState {
			edgeCalls = append(edgeCalls, edge)
			if block.ID != 1 {
				return out
			}
			value := out["values"]
			value.kind = arrayAllocated
			value.knownArray = true
			out["values"] = value
			return out
		}, nil, func(_ *procedureir.Statement, _, _ arrayFlowState) bool { return true }, true); err != nil {
		t.Fatal(err)
	}
	if got := seen["normal"]; got.kind != arrayAllocated {
		t.Fatalf("normal edge state = %#v, want allocated", got)
	}
	if got := seen["exceptional"]; got.kind != arrayAllocated {
		t.Fatalf("reliable exceptional edge state = %#v, want transferred output, edges=%+v", got, edgeCalls)
	}
	if got := seen["uncertain"]; got.kind != arrayUnallocated {
		t.Fatalf("uncertain edge state = %#v, want predecessor input (unallocated), edges=%+v", got, edgeCalls)
	}
}

func TestArrayCompactAdvancedPreservesPhysicalSourceLineOrder(t *testing.T) {
	statement := &procedureir.Statement{ID: 1, Kind: procedureir.StatementCall, Text: "first\nsecond", Range: vbaast.Range{StartLine: 1, EndLine: 2}}
	graph := vbacfg.Graph{
		Blocks: []vbacfg.Block{{ID: 0, Kind: vbacfg.BlockEntry}, {ID: 1, Kind: vbacfg.BlockStatement, StatementID: statement.ID, Statement: statement}, {ID: 2, Kind: vbacfg.BlockNormalExit}},
		Edges:  []vbacfg.Edge{{ID: 0, From: 0, To: 1, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal}, {ID: 1, From: 1, To: 2, Kind: vbacfg.EdgeProcedureExit, Class: vbacfg.EdgeNormal}},
		Entry:  0, NormalExit: 2, ExceptionalExit: 2, TerminationExit: 2, UnknownExit: 2,
	}
	view := graph.View(vbacfg.EdgeFilter{})
	var lines []int
	initial := arrayFlowState{"values": {kind: arrayUnknown, knownArray: true, origin: arrayOriginLocal}}
	if err := walkArrayCFGCompactAdvanced(t.Context(), &view, []string{"first", "second"}, initial, func(_ string, line int, in arrayFlowState) arrayFlowState {
		lines = append(lines, line)
		return in
	}, nil, nil, nil, true); err != nil {
		t.Fatal(err)
	}
	if len(lines) != 2 || lines[0] != 1 || lines[1] != 2 {
		t.Fatalf("source-line visit order = %v, want [1 2]", lines)
	}
}

func TestArrayCompactReliableExceptionalRejectsReDimPreserve(t *testing.T) {
	in := arrayFlowState{"values": {kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}}
	out := arrayFlowState{"values": {kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal}}
	plain := &procedureir.Statement{Kind: procedureir.StatementReDim, Text: "ReDim values(0)"}
	preserve := &procedureir.Statement{Kind: procedureir.StatementReDim, Text: "ReDim Preserve values(0)"}
	if !arrayAllocationTransferIsReliable(plain, in, out) {
		t.Fatal("plain ReDim should be reliable across an exceptional edge")
	}
	if arrayAllocationTransferIsReliable(preserve, in, out) {
		t.Fatal("ReDim Preserve must not be treated as a reliable exceptional allocation")
	}
}

func TestArrayCompactAdvancedStopSuppressesSuccessors(t *testing.T) {
	first := &procedureir.Statement{ID: 1, Kind: procedureir.StatementCall, Text: "stop", Range: vbaast.Range{StartLine: 1, EndLine: 1}}
	second := &procedureir.Statement{ID: 2, Kind: procedureir.StatementCall, Text: "successor", Range: vbaast.Range{StartLine: 2, EndLine: 2}}
	graph := vbacfg.Graph{
		Blocks: []vbacfg.Block{{ID: 0, Kind: vbacfg.BlockEntry}, {ID: 1, Kind: vbacfg.BlockStatement, StatementID: first.ID, Statement: first}, {ID: 2, Kind: vbacfg.BlockStatement, StatementID: second.ID, Statement: second}, {ID: 3, Kind: vbacfg.BlockNormalExit}},
		Edges:  []vbacfg.Edge{{ID: 0, From: 0, To: 1, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal}, {ID: 1, From: 1, To: 2, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal}, {ID: 2, From: 2, To: 3, Kind: vbacfg.EdgeProcedureExit, Class: vbacfg.EdgeNormal}},
		Entry:  0, NormalExit: 3, ExceptionalExit: 3, TerminationExit: 3, UnknownExit: 3,
	}
	view := graph.View(vbacfg.EdgeFilter{})
	visited := map[string]bool{}
	initial := arrayFlowState{"values": {kind: arrayUnknown, knownArray: true, origin: arrayOriginLocal}}
	if err := walkArrayCFGCompactAdvanced(t.Context(), &view, []string{"stop", "successor"}, initial, func(text string, _ int, in arrayFlowState) arrayFlowState {
		visited[text] = true
		return in
	}, nil, func(text string, _ int) bool { return text == "stop" }, nil, true); err != nil {
		t.Fatal(err)
	}
	if !visited["stop"] || visited["successor"] {
		t.Fatalf("stop visited=%v, successor visited=%v; want successor suppressed", visited["stop"], visited["successor"])
	}
}

func TestArrayLegacyExceptionalEdgeKeepsPredecessorInput(t *testing.T) {
	call := &procedureir.Statement{ID: 1, Kind: procedureir.StatementCall, Text: "moduleCall", Range: vbaast.Range{StartLine: 1, EndLine: 1}}
	exceptional := &procedureir.Statement{ID: 2, Kind: procedureir.StatementCall, Text: "handler", Range: vbaast.Range{StartLine: 2, EndLine: 2}}
	graph := vbacfg.Graph{
		Blocks: []vbacfg.Block{
			{ID: 0, Kind: vbacfg.BlockEntry},
			{ID: 1, Kind: vbacfg.BlockStatement, StatementID: call.ID, Statement: call},
			{ID: 2, Kind: vbacfg.BlockStatement, StatementID: exceptional.ID, Statement: exceptional},
			{ID: 3, Kind: vbacfg.BlockNormalExit},
		},
		Edges: []vbacfg.Edge{
			{ID: 0, From: 0, To: 1, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal},
			{ID: 1, From: 1, To: 2, Kind: vbacfg.EdgeError, Class: vbacfg.EdgeExceptional},
			{ID: 2, From: 1, To: 3, Kind: vbacfg.EdgeFallthrough, Class: vbacfg.EdgeNormal},
			{ID: 3, From: 2, To: 3, Kind: vbacfg.EdgeProcedureExit, Class: vbacfg.EdgeNormal},
		},
		Entry: 0, NormalExit: 3, ExceptionalExit: 3, TerminationExit: 3, UnknownExit: 3,
	}
	view := graph.View(vbacfg.EdgeFilter{})
	initial := arrayFlowState{"values": {kind: arrayUnallocated, knownArray: true, origin: arrayOriginLocal}}
	var handlerState arrayValue
	walkArrayCFGWorklistLegacy(&view, []string{"moduleCall", "handler"}, initial,
		func(text string, _ int, in arrayFlowState) arrayFlowState {
			if text == "moduleCall" {
				value := in["values"]
				value.kind = arrayAllocated
				value.knownArray = true
				in["values"] = value
			}
			if text == "handler" {
				handlerState = in["values"]
			}
			return in
		}, nil, nil, false)
	if handlerState.kind != arrayUnallocated {
		t.Fatalf("exceptional handler state = %#v, want predecessor input %#v", handlerState, initial["values"])
	}
}

func TestArrayConditionalAllocationBranchDoesNotMutateInput(t *testing.T) {
	statement := &procedureir.Statement{
		Kind:      procedureir.StatementIf,
		Condition: &procedureir.Expression{Text: "count > 0"},
	}
	block := vbacfg.Block{Statement: statement}
	edge := vbacfg.Edge{Kind: vbacfg.EdgeBranchTrue}
	initial := arrayFlowState{
		"values": {kind: arrayUnknown, knownArray: true, allocationCountSource: "count"},
	}

	got := applyArrayConditionalAllocationBranch(initial, nil, block, edge)
	if got["values"].kind != arrayAllocated {
		t.Fatalf("positive branch state = %#v, want allocated", got["values"])
	}
	if initial["values"].kind == arrayAllocated {
		t.Fatalf("conditional refinement mutated its input state: %#v", initial["values"])
	}
}
