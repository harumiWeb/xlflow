package preflight

import (
	"testing"

	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

func TestPolicyDecisions(t *testing.T) {
	policy := NewPolicy([]string{" vb052 "})
	if got := policy.Decide("VB052"); got != DecisionAllowed {
		t.Fatalf("VB052 decision = %v, want allowed", got)
	}
	if got := policy.Decide("VB054"); got != DecisionBlocking {
		t.Fatalf("VB054 decision = %v, want blocking", got)
	}
	if got := policy.Decide("VB001"); got != DecisionNonBlocking {
		t.Fatalf("VB001 decision = %v, want non-blocking", got)
	}
	if got := policy.Decide("UNKNOWN"); got != DecisionNonBlocking {
		t.Fatalf("unknown decision = %v, want non-blocking", got)
	}
}

func TestEveryRegistryBlockerUsesTheSamePolicy(t *testing.T) {
	for _, rule := range staticrules.All() {
		if !rule.PreflightBlocking {
			continue
		}
		policy := NewPolicy([]string{rule.ID})
		if got := policy.Decide(rule.ID); got != DecisionAllowed {
			t.Errorf("%s decision = %v, want allowed", rule.ID, got)
		}
	}
}

func TestPartitionIsStable(t *testing.T) {
	type item struct{ code string }
	items := []item{{"VB001"}, {"VB052"}, {"VB054"}, {"VB052"}}
	blocking, allowed := Partition(items, func(item item) string { return item.code }, NewPolicy([]string{"VB052"}))
	if len(blocking) != 1 || blocking[0].code != "VB054" {
		t.Fatalf("blocking = %#v", blocking)
	}
	if len(allowed) != 2 || allowed[0].code != "VB052" || allowed[1].code != "VB052" {
		t.Fatalf("allowed = %#v", allowed)
	}
}
