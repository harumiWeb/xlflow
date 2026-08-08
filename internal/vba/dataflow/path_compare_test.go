package dataflow

import (
	"fmt"
	"strings"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestComparePathKeysMatchesSerializedKeyOrdering(t *testing.T) {
	t.Parallel()
	paths := [][]PathStep{
		nil,
		{},
		{{Kind: "assignment", Label: "Alpha", Range: vbaast.Range{StartByte: 1, EndByte: 2, StartLine: 3}, StatementID: 4}},
		{{Kind: "assignment", Label: "alpha", Range: vbaast.Range{StartByte: 1, EndByte: 2, StartLine: 3}, StatementID: 4}},
		{{Kind: "assignment", Label: "ALPHA", Range: vbaast.Range{StartByte: 1, EndByte: 2, StartLine: 3}, StatementID: 4}},
		{{Kind: "assignment", Label: "ÄPFEL", Range: vbaast.Range{StartByte: 1, EndByte: 2, StartLine: 3}, StatementID: 4}},
		{{Kind: "assignment", Label: "äpfel", Range: vbaast.Range{StartByte: 1, EndByte: 2, StartLine: 3}, StatementID: 4}},
		{{Kind: "a\x00z", Label: "control\x01label", Range: vbaast.Range{StartByte: -10, EndByte: -2, StartLine: -1}, StatementID: -20}},
		{{Kind: "a", Label: "z\x00control", Range: vbaast.Range{StartByte: -2, EndByte: 10, StartLine: 11}, StatementID: 100}},
		{{Kind: "assignment", Label: "decimal", Range: vbaast.Range{StartByte: 1, EndByte: 10, StartLine: 100}, StatementID: 2}},
		{{Kind: "assignment", Label: "decimal", Range: vbaast.Range{StartByte: 10, EndByte: 1, StartLine: 2}, StatementID: 100}},
		{
			{Kind: "source", Label: "First", Range: vbaast.Range{StartByte: 1}, StatementID: 1},
			{Kind: "assignment", Label: "Second", Range: vbaast.Range{StartByte: 2}, StatementID: 2},
		},
	}

	for leftIndex, left := range paths {
		for rightIndex, right := range paths {
			want := strings.Compare(serializedPathKey(left), serializedPathKey(right))
			if got := comparePathKeys(left, right); got != want {
				t.Errorf("comparePathKeys(paths[%d], paths[%d]) = %d, want %d", leftIndex, rightIndex, got, want)
			}
		}
	}
}

func TestComparePathKeysASCIIAllocations(t *testing.T) {
	left := []PathStep{{Kind: "assignment", Label: "CustomerInput", Range: vbaast.Range{StartByte: 10, EndByte: 24, StartLine: 3}, StatementID: 7}}
	right := []PathStep{{Kind: "assignment", Label: "customerOutput", Range: vbaast.Range{StartByte: 10, EndByte: 25, StartLine: 3}, StatementID: 8}}
	if allocations := testing.AllocsPerRun(1000, func() {
		if comparePathKeys(left, right) >= 0 {
			t.Fatal("unexpected path ordering")
		}
	}); allocations != 0 {
		t.Fatalf("comparePathKeys allocated %.2f times for ASCII paths, want zero", allocations)
	}
}

func serializedPathKey(path []PathStep) string {
	var builder strings.Builder
	for _, step := range path {
		fmt.Fprintf(&builder, "%s\x00%s\x00%d\x00%d\x00%d\x00%d\x01", step.Kind, strings.ToLower(step.Label), step.Range.StartByte, step.Range.EndByte, step.Range.StartLine, step.StatementID)
	}
	return builder.String()
}
