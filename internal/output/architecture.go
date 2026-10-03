package output

import (
	"fmt"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/architecture"
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
)

const architectureHumanLimit = 10

func (r renderer) renderArchitecture(env Envelope) string {
	report, ok := env.Architecture.(architecture.Report)
	if !ok {
		return ""
	}
	summary := report.Summary
	var b strings.Builder
	b.WriteString(r.section("Project Architecture"))
	fmt.Fprintf(&b, "%d modules / %d procedures\n", summary.Modules, summary.Procedures)
	if report.Scope.Path != "" || report.Scope.Module != "" {
		b.WriteString(r.kvRows(kvRow{"Path filter", report.Scope.Path}, kvRow{"Module filter", report.Scope.Module}))
		fmt.Fprintf(&b, "Whole project: %d modules / %d procedures; reachability and hotspot ranks use the whole project.\n", report.ProjectSummary.Modules, report.ProjectSummary.Procedures)
	}
	b.WriteString(r.section("Entry points"))
	for _, entry := range report.EntryPoints[:min(len(report.EntryPoints), architectureHumanLimit)] {
		fmt.Fprintf(&b, "  %s [%s] %s (%s)\n", entry.Target, entry.Confidence, entry.Reason, entry.Status)
	}
	writeArchitectureRemainder(&b, len(report.EntryPoints), "entry point")
	b.WriteString(r.section("Dependencies"))
	b.WriteString(r.kvRows(
		kvRow{"Confirmed calls", fmt.Sprint(summary.ConfirmedCallEdges)},
		kvRow{"Uncertain calls", fmt.Sprint(summary.UncertainCallEdges)},
		kvRow{"Cyclic components", fmt.Sprint(summary.DependencyCycles)},
		kvRow{"External references", fmt.Sprint(summary.ExternalDependencies)},
	))
	b.WriteString(r.section("Top hotspots"))
	for _, cohort := range []struct {
		label, kind string
		entities    []hotspots.Entity
	}{
		{"Procedures", "procedure hotspot", report.Hotspots.Procedures},
		{"Modules", "module hotspot", report.Hotspots.Modules},
	} {
		fmt.Fprintf(&b, "%s:\n", cohort.label)
		for _, entity := range cohort.entities[:min(len(cohort.entities), architectureHumanLimit)] {
			fmt.Fprintf(&b, "  #%d %s (%s:%d)  score %.2f  fan-in %d / fan-out %d  complexity %d\n", entity.Rank, entity.Name, entity.File, entity.Line, entity.Score, entity.RawSignals["call_fan_in"], entity.RawSignals["call_fan_out"], entity.RawSignals["complexity"])
		}
		writeArchitectureRemainder(&b, len(cohort.entities), cohort.kind)
	}
	b.WriteString("Hotspot scores are review leads; Excel-effect signals retain their metrics counting model.\n")
	b.WriteString(r.section("Unreachable internal procedures"))
	for _, node := range report.Unreachable[:min(len(report.Unreachable), architectureHumanLimit)] {
		fmt.Fprintf(&b, "  %s\n", node.ID.QualifiedName)
	}
	writeArchitectureRemainder(&b, len(report.Unreachable), "unreachable procedure")
	b.WriteString(r.section("State and effects"))
	state := report.ModuleState.Summary
	fmt.Fprintf(&b, "  Module state inventory: %d field(s) including constants; mutable reads %d / writes %d / mutations %d\n", len(report.ModuleState.Fields), state.MutableStateReads, state.MutableStateWrites, state.MutableStateMutations)
	fmt.Fprintf(&b, "  Excel effects: %d direct evidence item(s)\n", report.ExcelEffects.DirectEffectCount)
	b.WriteString(r.section("Uncertainty"))
	b.WriteString(r.kvRows(
		kvRow{"Unresolved calls", fmt.Sprint(report.Uncertainty.UnresolvedCallCount)},
		kvRow{"Ambiguous calls", fmt.Sprint(report.Uncertainty.AmbiguousCallCount)},
		kvRow{"Dynamic calls", fmt.Sprint(report.Uncertainty.DynamicCallCount)},
	))
	b.WriteString("Use --json for all evidence; inspect calls, graph dependencies, and impact for focused inspection.\n")
	return b.String()
}

func writeArchitectureRemainder(b *strings.Builder, count int, kind string) {
	if count == 0 {
		b.WriteString("  (none)\n")
	} else if count > architectureHumanLimit {
		fmt.Fprintf(b, "  %d more %s(s); use --json for the complete report.\n", count-architectureHumanLimit, kind)
	}
}
