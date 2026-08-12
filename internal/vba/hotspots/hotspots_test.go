package hotspots

import (
	"reflect"
	"testing"
)

func TestPercentilesAverageTiesAndConstantSignals(t *testing.T) {
	got := Percentiles([]int{1, 2, 2, 4})
	want := []float64{0, 50, 50, 100}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("percentiles = %#v, want %#v", got, want)
	}
	if got := Percentiles([]int{4, 4, 4}); !reflect.DeepEqual(got, []float64{0, 0, 0}) {
		t.Fatalf("constant percentiles = %#v", got)
	}
}

func TestRankIsDeterministicAndUsesEqualWeightSignals(t *testing.T) {
	inputs := []Input{
		{ID: "b", Kind: "procedure", File: "b.bas", Name: "B", RawSignals: map[string]int{"complexity": 2, "call_fan_out": 0}},
		{ID: "a", Kind: "procedure", File: "a.bas", Name: "A", RawSignals: map[string]int{"complexity": 1, "call_fan_out": 2}},
	}
	first := Rank(inputs)
	second := Rank([]Input{inputs[1], inputs[0]})
	if !reflect.DeepEqual(first, second) {
		t.Fatalf("rank changed with input order:\n%#v\n%#v", first, second)
	}
	if first[0].ID != "a" || first[1].ID != "b" || first[0].Score != 50 || first[1].Score != 50 {
		t.Fatalf("rank = %#v, want tied score with stable identity order", first)
	}
}

func TestSelectUsesUnionAndMarksThresholdPromotion(t *testing.T) {
	entities := []Entity{{ID: "a", Rank: 1, Score: 90}, {ID: "b", Rank: 2, Score: 40}, {ID: "c", Rank: 3, Score: 10}}
	got := Select(entities, 1, 40)
	if len(got) != 2 || !got[0].SelectedBy.TopN || !got[0].SelectedBy.Threshold || !got[1].SelectedBy.Threshold {
		t.Fatalf("selection = %#v", got)
	}
}
