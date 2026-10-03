package architecture

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	cfggraph "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
	vbasymbols "github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestCollectorKeepsCanonicalModuleProcedureAndCallbackEvidence(t *testing.T) {
	root, cfg := collectorFixture(t)
	report, _, err := CollectContext(context.Background(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}

	kinds := make(map[string]bool)
	moduleIDs := make(map[string]string)
	for _, module := range report.Modules {
		kinds[module.Kind] = true
		moduleIDs[module.Name] = module.ID
		if !strings.HasPrefix(module.ID, "module|") {
			t.Errorf("module %q ID=%q is not a dependency node ID", module.Name, module.ID)
		}
	}
	for _, kind := range []string{"standard", "class", "form", "document"} {
		if !kinds[kind] {
			t.Errorf("report omitted %s module kind: %+v", kind, report.Modules)
		}
	}
	procedureIDs := make(map[string]bool)
	for _, procedure := range report.Procedures {
		procedureIDs[procedure.ID] = true
		if !strings.HasPrefix(procedure.ID, "procedure|") {
			t.Errorf("procedure %q ID=%q is not a dependency node ID", procedure.Name, procedure.ID)
		}
		if procedure.ModuleID == "" || procedure.ModuleID != moduleIDs[procedure.Module] {
			t.Errorf("procedure %s module join=%q, module ID=%q", procedure.QualifiedName, procedure.ModuleID, moduleIDs[procedure.Module])
		}
	}
	for _, node := range report.Dependencies.Nodes {
		if strings.HasPrefix(node.ID, "procedure|") && !procedureIDs[node.ID] {
			t.Errorf("dependency procedure %q has no report procedure", node.ID)
		}
	}
	if len(report.EntryPoints) == 0 || report.EntryPoints[0].NodeID == "" || !procedureIDs[report.EntryPoints[0].NodeID] {
		t.Fatalf("entry point does not join a report procedure: %+v", report.EntryPoints)
	}

	var callback *DynamicReference
	for index := range report.DynamicReferences {
		if report.DynamicReferences[index].API == "application.onkey" {
			callback = &report.DynamicReferences[index]
			break
		}
	}
	if callback == nil || callback.Target != "Worker.Callback" || callback.Kind != "static" {
		t.Fatalf("callback evidence was not retained: %+v", report.DynamicReferences)
	}
	if !procedureIDs[callback.CallerID] || callback.Range.StartByte <= 0 {
		t.Fatalf("callback evidence lacks canonical caller/range join: %+v", *callback)
	}
	encoded, err := json.Marshal(report)
	if err != nil {
		t.Fatal(err)
	}
	var wire map[string]any
	if err := json.Unmarshal(encoded, &wire); err != nil {
		t.Fatal(err)
	}
	refs, ok := wire["dynamic_references"].([]any)
	if !ok || len(refs) == 0 {
		t.Fatalf("dynamic_references JSON missing: %s", encoded)
	}
	ref := refs[0].(map[string]any)
	if _, ok := ref["caller_id"]; !ok {
		t.Errorf("callback JSON has no snake_case caller_id: %s", encoded)
	}
	if _, leaked := ref["CallerID"]; leaked {
		t.Errorf("callback JSON leaked Go field names: %s", encoded)
	}
	if report.ExcelEffects.DirectEffectCount == 0 {
		t.Fatal("recognized direct Excel effects were not projected")
	}
	for _, evidence := range report.ExcelEffects.Evidence {
		switch evidence.Kind {
		case "opens_file", "launches_process", "suppresses_errors", "raises_error":
			t.Errorf("non-Excel direct effect %q leaked into Excel effects", evidence.Kind)
		}
	}
}

func TestCollectorProcedureMetricsAndHotspotsMatchNativeProjection(t *testing.T) {
	root, cfg := collectorFixture(t)
	report, _, err := CollectContext(context.Background(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	nativeMetrics, err := nativeProcedureMetrics(t, root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	metricsByID := make(map[string]proceduremetrics.Metrics, len(nativeMetrics))
	for _, metric := range nativeMetrics {
		metricsByID[metricIdentityKey(metric.File, metric.Module, metric.Name, string(metric.Kind), metric.DeclarationRange.StartByte)] = metric.Metrics
	}
	for _, procedure := range report.Procedures {
		key := metricIdentityKey(procedure.File, procedure.Module, procedure.Name, procedure.Kind, procedure.DeclarationRange.StartByte)
		if got, ok := metricsByID[key]; !ok || !reflect.DeepEqual(procedure.Metrics, got) {
			t.Errorf("%s metrics=%+v native=%+v (present=%v)", procedure.QualifiedName, procedure.Metrics, got, ok)
		}
	}
	wantHotspots, _, err := hotspots.BuildFromMetrics(context.Background(), nativeMetrics)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(report.Hotspots, wantHotspots) {
		t.Fatalf("hotspot projection differs from native collector:\n got: %+v\nwant: %+v", report.Hotspots, wantHotspots)
	}
}

func TestCollectorScopeKeepsSCCBoundaryAndProjectRanks(t *testing.T) {
	root, cfg := collectorFixture(t)
	full, _, err := CollectContext(context.Background(), root, cfg, Options{})
	if err != nil {
		t.Fatal(err)
	}
	scoped, _, err := CollectContext(context.Background(), root, cfg, Options{Path: "src/modules/Main.bas"})
	if err != nil {
		t.Fatal(err)
	}
	if scoped.ProjectSummary != full.Summary {
		t.Fatalf("scoped project summary=%+v, full summary=%+v", scoped.ProjectSummary, full.Summary)
	}
	if len(full.Cycles) == 0 || len(scoped.Cycles) == 0 {
		t.Fatalf("cyclic SCC lost in scoped projection: full=%+v scoped=%+v", full.Cycles, scoped.Cycles)
	}
	if len(scoped.Cycles[0].Nodes) < 2 || len(scoped.Cycles[0].Witness.Edges) < 2 {
		t.Fatalf("scoped SCC is missing its full component or witness: %+v", scoped.Cycles[0])
	}
	if len(scoped.DependencyBoundaryNodeIDs) == 0 {
		t.Fatalf("cross-file dependency nodes were not marked as boundary nodes: %+v", scoped.Dependencies)
	}
	boundary := make(map[string]bool, len(scoped.DependencyBoundaryNodeIDs))
	for _, id := range scoped.DependencyBoundaryNodeIDs {
		boundary[id] = true
	}
	foundWorkerBoundary := false
	for _, node := range scoped.Dependencies.Nodes {
		if node.Module == "Worker" && boundary[node.ID] {
			foundWorkerBoundary = true
		}
	}
	if !foundWorkerBoundary {
		t.Fatalf("cycle partner Worker was not retained and marked as a boundary node: %+v", scoped.DependencyBoundaryNodeIDs)
	}

	fullRanks := make(map[string]int, len(full.Hotspots.Procedures))
	for _, hotspot := range full.Hotspots.Procedures {
		fullRanks[hotspot.ID] = hotspot.Rank
	}
	if len(scoped.Hotspots.Procedures) == 0 {
		t.Fatal("scoped report unexpectedly has no procedure hotspots")
	}
	for _, hotspot := range scoped.Hotspots.Procedures {
		if fullRanks[hotspot.ID] != hotspot.Rank {
			t.Errorf("scope reranked %s: scoped=%d full=%d", hotspot.ID, hotspot.Rank, fullRanks[hotspot.ID])
		}
	}
}

func TestCollectorReportsIncompleteTypeDatabaseWithoutFailing(t *testing.T) {
	root, cfg := collectorFixture(t)
	hooks := &collectionHooks{loadTypeDB: func() (typedb.LoadResult, error) {
		return typedb.LoadResult{DB: vbadb.New(), Complete: false, Warnings: []string{"generated TypeLib manifest is missing"}}, nil
	}}
	report, warnings, err := collectContextWithHooks(context.Background(), root, cfg, Options{}, hooks)
	if err != nil {
		t.Fatalf("source-only collection failed with incomplete optional type metadata: %v", err)
	}
	if !report.Uncertainty.TypeDatabaseLoaded || report.Uncertainty.TypeDatabaseComplete {
		t.Fatalf("TypeDB loaded/completeness flags=%v/%v", report.Uncertainty.TypeDatabaseLoaded, report.Uncertainty.TypeDatabaseComplete)
	}
	if len(report.Uncertainty.Reasons) == 0 || len(warnings) == 0 {
		t.Fatalf("incomplete TypeDB not made explicit: uncertainty=%+v warnings=%+v", report.Uncertainty, warnings)
	}
}

func collectorFixture(t testing.TB) (string, config.Config) {
	t.Helper()
	root := t.TempDir()
	cfg := config.Default()
	cfg.Project.Entry = "Main.Run"
	cfg.UserForm.CodeSource = "sidecar"
	writeCollectorSource(t, root, "src/modules/Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Call Main.Start
    Application.OnKey "{LEFT}", "Worker.Callback"
    Call Panel.FormAction
End Sub
Public Sub Start()
    Call Worker.LoopBack
    Application.Calculate
    Range("A1").Value = 1
    On Error Resume Next
    Open "output.txt" For Output As #1
    Shell "calc.exe"
End Sub
Private Sub Unused()
End Sub
`)
	writeCollectorSource(t, root, "src/classes/Worker.cls", `VERSION 1.0 CLASS
Attribute VB_Name = "Worker"
Option Explicit
Public Sub LoopBack()
    Call Main.Start
End Sub
Public Sub Callback()
End Sub
`)
	writeCollectorSource(t, root, "src/forms/Panel.frm", `VERSION 5.00
Begin VB.UserForm Panel
End
Attribute VB_Name = "Panel"
`)
	writeCollectorSource(t, root, "src/forms/code/Panel.bas", `Attribute VB_Name = "Panel"
Option Explicit
Public Sub FormAction()
End Sub
`)
	writeCollectorSource(t, root, "src/workbook/ThisWorkbook.cls", `VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Option Explicit
Public Sub Workbook_Open()
End Sub
`)
	return root, cfg
}

func writeCollectorSource(t testing.TB, root, relative, source string) {
	t.Helper()
	path := filepath.Join(root, filepath.FromSlash(relative))
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
}

func nativeProcedureMetrics(t testing.TB, root string, cfg config.Config) ([]proceduremetrics.ProcedureMetrics, error) {
	t.Helper()
	files, err := vbasymbols.DiscoverProjectSourceFilesContext(context.Background(), root, cfg)
	if err != nil {
		return nil, err
	}
	documents := make([]procedureir.DocumentIR, 0, len(files))
	graphs := make([]cfggraph.Document, 0, len(files))
	for _, file := range files {
		path, err := filepath.Rel(root, file.Path)
		if err != nil {
			return nil, err
		}
		path = filepath.ToSlash(path)
		source, err := os.ReadFile(file.Path)
		if err != nil {
			return nil, err
		}
		parsed, err := ast.ParseDocumentContext(context.Background(), path, source)
		if err != nil {
			return nil, err
		}
		symbolFile, err := vbasymbols.InspectParsedContext(context.Background(), vbasymbols.SourceOptions{
			RootDir: root, Path: path, ModuleKind: file.ModuleKind, IncludePrivate: true,
		}, parsed)
		if err != nil {
			parsed.Close()
			return nil, err
		}
		ir, err := procedureir.BuildParsedContext(context.Background(), procedureir.BuildOptions{
			RootDir: root, Path: path, ModuleName: symbolFile.ModuleName, ModuleKind: file.ModuleKind,
		}, parsed)
		parsed.Close()
		if err != nil {
			return nil, err
		}
		graph, err := cfggraph.BuildDocumentContext(context.Background(), ir)
		if err != nil {
			return nil, err
		}
		documents = append(documents, ir)
		graphs = append(graphs, graph)
	}
	resolver := procedureir.BuildProjectResolver(documents, nil, true)
	metrics := make([]proceduremetrics.ProcedureMetrics, 0)
	for index, document := range documents {
		resolved := procedureir.ResolveView(document, resolver).Materialize()
		metrics = append(metrics, proceduremetrics.CollectDocument(resolved, graphs[index])...)
	}
	return metrics, nil
}
