package architecture

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
)

func TestReviewScopedEntryPointsPreserveGlobalReachability(t *testing.T) {
	root, cfg := contractProject(t, 0)
	writeCollectorSource(t, root, "src/modules/Main.bas", "Attribute VB_Name = \"Main\"\nPublic Sub Run()\nEnd Sub\nPublic Sub Internal()\nEnd Sub\n")
	writeCollectorSource(t, root, "src/modules/Worker.bas", "Attribute VB_Name = \"Worker\"\nPublic Sub Execute()\nMain.Internal\nEnd Sub\n")
	full, _, err := CollectContext(t.Context(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, options := range []Options{{Module: "Main"}, {Path: "src/modules/Main.bas"}} {
		scoped, _, err := CollectContext(t.Context(), root, cfg, options)
		if err != nil {
			t.Fatal(err)
		}
		if len(scoped.EntryPoints) == 0 || scoped.ProjectSummary != full.ProjectSummary {
			t.Fatalf("lost roots/project summary: %+v", scoped)
		}
		for _, entry := range scoped.EntryPoints {
			if entry.Target == "Worker.Execute" {
				t.Fatal("unrelated public root leaked into selected view")
			}
		}
		for _, procedure := range scoped.Procedures {
			index := slices.IndexFunc(full.Procedures, func(candidate Procedure) bool { return candidate.ID == procedure.ID })
			if index < 0 || !reflect.DeepEqual(procedure, full.Procedures[index]) {
				t.Fatal("display filter changed whole-project procedure facts")
			}
		}
	}
}

func TestReviewEntryPointFilterPreservesMissingAndAmbiguousEvidence(t *testing.T) {
	input := []EntryPoint{
		{Target: "Missing.Entry", Status: "unresolved", Confidence: "possible", Candidates: []string{}},
		{Target: "Run", Status: "ambiguous", Confidence: "possible", Candidates: []string{"a", "b"}},
		{Target: "Other.Run", Status: "resolved", NodeID: "b", Candidates: []string{"b"}},
	}
	got := filterEntryPoints(input, []Procedure{{ID: "a"}})
	if len(got) != 2 || got[0].Status != "unresolved" || got[1].Status != "ambiguous" || got[1].NodeID != "" || !reflect.DeepEqual(got[1].Candidates, []string{"a"}) {
		t.Fatalf("filtered root evidence=%+v", got)
	}
	if len(input[1].Candidates) != 2 {
		t.Fatal("filter mutated full root evidence")
	}
}

func TestReviewScopedCycleRetainsNonadjacentBoundaryMember(t *testing.T) {
	root, cfg := contractProject(t, 0)
	cfg.Project.Entry = "A.P"
	for index, name := range []string{"A", "B", "C", "D"} {
		next := []string{"B", "C", "D", "A"}[index]
		writeCollectorSource(t, root, "src/modules/"+name+".bas", "Attribute VB_Name = \""+name+"\"\nPublic Sub P()\n"+next+".P\nEnd Sub\n")
	}
	full, _, err := CollectContext(t.Context(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	scoped, _, err := CollectContext(t.Context(), root, cfg, Options{Module: "A"})
	if err != nil {
		t.Fatal(err)
	}
	if len(scoped.Cycles) != 1 || len(scoped.Cycles[0].Nodes) != 4 || scoped.Summary.Procedures != 1 || scoped.ProjectSummary != full.ProjectSummary {
		t.Fatalf("incomplete cycle/view counts: %+v", scoped)
	}
	for _, member := range scoped.Cycles[0].Nodes {
		id := "procedure|" + member.String()
		if !slices.ContainsFunc(scoped.Dependencies.Nodes, func(node callgraph.DependencyNode) bool { return node.ID == id }) {
			t.Errorf("SCC member lacks dependency node: %s", id)
		}
		if member.Module != "A" && !slices.Contains(scoped.DependencyBoundaryNodeIDs, id) {
			t.Errorf("SCC member lacks boundary marker: %s", id)
		}
	}
	for _, edge := range scoped.Dependencies.Edges {
		if !slices.ContainsFunc(full.Dependencies.Nodes, func(node callgraph.DependencyNode) bool {
			return node.File == "src/modules/A.bas" && (node.ID == edge.From || node.ID == edge.To)
		}) {
			t.Fatalf("SCC node retention expanded incident edge policy: %+v", edge)
		}
	}
}

func TestReviewConditionalDeclarationsKeepDistinctMetrics(t *testing.T) {
	root, cfg := contractProject(t, 0)
	writeCollectorSource(t, root, "src/modules/Main.bas", `Attribute VB_Name = "Main"
#If VBA7 Then
Public Sub Run()
If True Then
Debug.Print "branch"
End If
End Sub
#Else
Public Sub Run()
End Sub
#End If
`)
	report, _, err := CollectContext(t.Context(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if len(report.Procedures) != 2 || report.Procedures[0].Metrics.CyclomaticComplexity != 2 || report.Procedures[1].Metrics.CyclomaticComplexity != 1 {
		t.Fatalf("conditional declarations share metrics: %+v", report.Procedures)
	}
	if report.Procedures[0].DeclarationRange.StartByte == report.Procedures[1].DeclarationRange.StartByte {
		t.Fatal("fixture lost distinct declarations")
	}
	native, err := nativeProcedureMetrics(t, root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	for _, procedure := range report.Procedures {
		index := slices.IndexFunc(native, func(metric proceduremetrics.ProcedureMetrics) bool {
			return metric.DeclarationRange.StartByte == procedure.DeclarationRange.StartByte
		})
		if index < 0 || !reflect.DeepEqual(procedure.Metrics, native[index].Metrics) {
			t.Fatalf("conditional declaration does not match native metrics: %+v", procedure)
		}
	}
}

func TestReviewExistingEmptyDirectoriesAreValidScopes(t *testing.T) {
	root, cfg := contractProject(t, 1)
	full, _, err := CollectContext(t.Context(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{"src/modules/Empty", "artifacts"} {
		if err := os.MkdirAll(filepath.Join(root, path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(root, path, "ignored.frx"), []byte("not source"), 0o644); err != nil {
			t.Fatal(err)
		}
		report, _, err := CollectContext(t.Context(), root, cfg, Options{Path: path})
		if err != nil {
			t.Fatal(err)
		}
		if report.ProjectSummary != full.ProjectSummary || report.Summary != (Summary{}) || report.Modules == nil || report.Procedures == nil || report.EntryPoints == nil || report.Dependencies.Nodes == nil {
			t.Fatalf("empty directory lost initialized empty view: %+v", report)
		}
		_, _, err = CollectContext(t.Context(), root, cfg, Options{Path: path + "/ignored.frx"})
		if !errors.Is(err, ErrInvalidScope) {
			t.Fatalf("nondiscovered file scope error=%v", err)
		}
	}
}
