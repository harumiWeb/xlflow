package output

import (
	"fmt"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/architecture"
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
)

func TestArchitectureRendererLimitsListsAndShowsScopeContext(t *testing.T) {
	entries := make([]architecture.EntryPoint, 25)
	entities := make([]hotspots.Entity, 25)
	for i := range entries {
		entries[i] = architecture.EntryPoint{Target: fmt.Sprintf("Module.P%02d", i), Confidence: "possible", Reason: "public API"}
		entities[i] = hotspots.Entity{File: "src/modules/Module.bas", Module: "Module", Name: fmt.Sprintf("P%02d", i), Rank: i + 1, Score: float64(90 - i), RawSignals: map[string]int{"call_fan_in": 1, "call_fan_out": 2, "complexity": 3}}
	}
	env := New("architecture")
	env.Architecture = architecture.Report{
		Summary:        architecture.Summary{Modules: 1, Procedures: 25, ConfirmedCallEdges: 2, UncertainCallEdges: 3, DependencyCycles: 1, ExternalDependencies: 4},
		ProjectSummary: architecture.Summary{Modules: 8, Procedures: 50},
		Scope:          architecture.Scope{Module: "Module"},
		EntryPoints:    entries,
		Hotspots:       hotspots.Report{Procedures: entities},
		Uncertainty:    architecture.Uncertainty{UnresolvedCallCount: 3, AmbiguousCallCount: 2, DynamicCallCount: 1},
	}
	text := renderHuman(env, Options{})
	for _, want := range []string{"1 modules / 25 procedures", "Whole project: 8 modules / 50 procedures", "Module.P00 [possible]", "15 more entry point(s)", "15 more procedure hotspot(s)", "Unresolved calls", "(none)"} {
		if !strings.Contains(text, want) {
			t.Errorf("missing %q:\n%s", want, text)
		}
	}
	if strings.Contains(text, "P10") || strings.Contains(text, "P24") {
		t.Fatalf("renderer exceeded list limit:\n%s", text)
	}
}

func TestArchitectureRendererFailureDoesNotInventReport(t *testing.T) {
	env := Failure("architecture", Error{Code: "architecture_parse_failed", Message: "parse recovery"})
	text := renderHuman(env, Options{})
	if !strings.Contains(text, "architecture_parse_failed") || strings.Contains(text, "Project Architecture") || strings.Contains(text, "Top hotspots") {
		t.Fatalf("failure rendered a fabricated report: %s", text)
	}
}

func TestArchitectureRendererStateInventoryIncludesConstants(t *testing.T) {
	env := New("architecture")
	env.Architecture = architecture.Report{ModuleState: architecture.ModuleState{
		Fields:  []architecture.ModuleStateField{{Name: "Limit", Kind: "const"}, {Name: "state", Kind: "variable"}},
		Summary: architecture.ModuleStateSummary{MutableStateReads: 1, MutableStateWrites: 2},
	}}
	text := renderHuman(env, Options{})
	if !strings.Contains(text, "Module state inventory: 2 field(s) including constants; mutable reads 1 / writes 2 / mutations 0") || strings.Contains(text, "Mutable state: 2 field(s)") {
		t.Fatalf("constant inventory mislabeled as mutable fields: %s", text)
	}
}
