package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
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
