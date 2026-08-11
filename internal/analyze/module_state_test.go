package analyze

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA240ReportsCrossEntryMutableCollectionAndMetrics(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedItems As Collection

Public Sub Run()
  Set sharedItems = New Collection
  sharedItems.Add "run"
End Sub

Public Sub Other()
  sharedItems.Add "other"
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA240")
	if len(findings) != 1 || findings[0].Line != 2 {
		t.Fatalf("VBA240 findings = %+v, want one declaration finding", findings)
	}
	metrics, ok := result.AnalysisMetrics.(map[string]any)
	if !ok {
		t.Fatalf("analysis metrics type = %T, want map", result.AnalysisMetrics)
	}
	state, ok := metrics["module_state"].(map[string]any)
	if !ok {
		t.Fatalf("module state metrics = %#v", metrics["module_state"])
	}
	fields, ok := state["fields"].([]map[string]any)
	if !ok || len(fields) != 1 {
		t.Fatalf("module state fields = %#v", state["fields"])
	}
	if fields[0]["classification"] != "mutable" || fields[0]["reader_count"] != 2 || fields[0]["writer_count"] != 2 || filepath.IsAbs(fields[0]["file"].(string)) {
		t.Fatalf("unexpected field metrics = %#v", fields[0])
	}
	procedures, ok := state["procedures"].([]map[string]any)
	if !ok || len(procedures) == 0 {
		t.Fatalf("procedure access metrics = %#v", state["procedures"])
	}
	seenRun := false
	for _, procedure := range procedures {
		if procedure["qualified"] == "Main.Run" {
			seenRun = true
			reads, _ := procedure["reads"].([]string)
			writes, _ := procedure["writes"].([]string)
			if !containsString(reads, "Main.sharedItems") || !containsString(writes, "Main.sharedItems") {
				t.Fatalf("Run access metrics = %#v", procedure)
			}
		}
	}
	if !seenRun {
		t.Fatalf("procedure metrics missing Main.Run: %#v", procedures)
	}
}

func containsString(items []string, want string) bool {
	for _, item := range items {
		if item == want {
			return true
		}
	}
	return false
}

func TestVBA240DoesNotWarnForSingleRootButReportsReadOnlyConfiguration(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private configuredName As String

Public Sub Run()
  Debug.Print configuredName
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA240"); len(findings) != 0 {
		t.Fatalf("single-root read-only state should not warn: %+v", findings)
	}
	metrics := result.AnalysisMetrics.(map[string]any)
	state := metrics["module_state"].(map[string]any)
	fields := state["fields"].([]map[string]any)
	if len(fields) != 1 || fields[0]["classification"] != "read_only_configuration" {
		t.Fatalf("unexpected read-only metrics: %#v", fields)
	}
}

func TestVBA240LeavesConstantsOutOfMutableRisk(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Const VersionName As String = "v1"

Public Sub Run()
  Debug.Print VersionName
End Sub

Public Sub Other()
  Debug.Print VersionName
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA240"); len(findings) != 0 {
		t.Fatalf("constant state should not warn: %+v", findings)
	}
}

func TestVBA240DoesNotTreatCollectionItemReadsAsMutations(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedItems As Collection

Public Sub Run()
  Debug.Print sharedItems.Item(1)
End Sub

Public Sub Other()
  Debug.Print sharedItems.Item(1)
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA240"); len(findings) != 0 {
		t.Fatalf("collection Item reads should not be treated as mutation: %+v", findings)
	}
	metrics := result.AnalysisMetrics.(map[string]any)
	state := metrics["module_state"].(map[string]any)
	fields := state["fields"].([]map[string]any)
	if len(fields) != 1 || fields[0]["writer_count"] != 0 || fields[0]["mutator_count"] != 0 {
		t.Fatalf("unexpected Item read metrics: %#v", fields)
	}
}

func TestVBA240IgnoresLocalStateThatShadowsModuleField(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedItems As Collection

Public Sub Run()
  Dim sharedItems As Collection
  Set sharedItems = New Collection
  sharedItems.Add "local"
End Sub

Public Sub Other()
  sharedItems.Add "module"
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA240"); len(findings) != 0 {
		t.Fatalf("shadowed local collection should not create module-state finding: %+v", findings)
	}
	state := result.AnalysisMetrics.(map[string]any)["module_state"].(map[string]any)
	field := state["fields"].([]map[string]any)[0]
	if field["writer_count"] != 1 || field["mutator_count"] != 1 || field["root_count"] != 1 {
		t.Fatalf("shadowed local was attributed to module field: %#v", field)
	}
}

func TestVBA240DoesNotResolvePrivateFieldFromAnotherModule(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Other.bas", `Option Explicit
Private sharedItems As Collection

Public Sub Initialize()
  Set sharedItems = New Collection
End Sub
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Sub Run()
  sharedItems.Add "run"
End Sub

Public Sub Other()
  sharedItems.Add "other"
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA240"); len(findings) != 0 {
		t.Fatalf("private field from another module should not be attributed: %+v", findings)
	}
	state := result.AnalysisMetrics.(map[string]any)["module_state"].(map[string]any)
	field := state["fields"].([]map[string]any)[0]
	if field["writer_count"] != 1 || field["mutator_count"] != 0 {
		t.Fatalf("unexpected private field metrics: %#v", field)
	}
}

func TestVBA240MetricsSortDeclarationLinesNumerically(t *testing.T) {
	metrics := moduleStateMetricsProjection([]*moduleStateField{
		{DisplayFile: "Main.bas", Name: "lineTen", Line: 10},
		{DisplayFile: "Main.bas", Name: "lineTwo", Line: 2},
	}, map[string]*moduleStateProcedureAccess{})
	fields := metrics["module_state"].(map[string]any)["fields"].([]map[string]any)
	if got := fields[0]["name"]; got != "lineTwo" {
		t.Fatalf("numeric field ordering = %#v, want lineTwo first", fields)
	}
}

func TestModuleStateBuiltInTypeMatchingUsesQualifiedExactNames(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name      string
		typ       string
		wantExcel bool
		wantStore bool
	}{
		{name: "qualified Excel type", typ: " Excel.Application ", wantExcel: true},
		{name: "qualified VBA type", typ: "VBA.Range", wantExcel: true},
		{name: "qualified scripting type", typ: "Scripting.Dictionary", wantStore: true},
		{name: "qualified collection type", typ: "VBA.Collection", wantStore: true},
		{name: "Excel substring", typ: "RangeHelper"},
		{name: "application substring", typ: "ApplicationConfig"},
		{name: "collection substring", typ: "MyCollection"},
		{name: "unrecognized qualifier", typ: "Custom.Range"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := moduleStateExcelType(tt.typ); got != tt.wantExcel {
				t.Fatalf("moduleStateExcelType(%q) = %t, want %t", tt.typ, got, tt.wantExcel)
			}
			if got := moduleStateCollectionType(tt.typ); got != tt.wantStore {
				t.Fatalf("moduleStateCollectionType(%q) = %t, want %t", tt.typ, got, tt.wantStore)
			}
		})
	}
}

func TestModuleStateProcedureMetricsSortsSameNamedProceduresByKey(t *testing.T) {
	t.Parallel()
	accesses := map[string]*moduleStateProcedureAccess{
		"z": {
			Procedure: moduleStateProcedure{Key: "z", File: "Main.bas", DisplayFile: "Main.bas", Module: "Main", Name: "Value"},
			Writes:    map[string]string{"z": "Main.z"},
		},
		"a": {
			Procedure: moduleStateProcedure{Key: "a", File: "Main.bas", DisplayFile: "Main.bas", Module: "Main", Name: "Value"},
			Writes:    map[string]string{"a": "Main.a"},
		},
	}
	for attempt := 0; attempt < 20; attempt++ {
		items := moduleStateProcedureMetrics(accesses)
		if got := items[0]["writes"].([]string)[0]; got != "Main.a" {
			t.Fatalf("same-named procedure metrics = %#v, want key-sorted order", items)
		}
	}
}

func TestVBA240RecognizesLateBoundDictionaryInitialization(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedItems As Object

Public Sub Run()
  Set sharedItems = CreateObject("Scripting.Dictionary")
  sharedItems.Add "run", 1
End Sub

Public Sub Other()
  sharedItems.Add "other", 2
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA240")
	if len(findings) != 1 || !strings.Contains(findings[0].Reason, "Collection or Dictionary") {
		t.Fatalf("late-bound dictionary findings = %+v", findings)
	}
	state := result.AnalysisMetrics.(map[string]any)["module_state"].(map[string]any)
	field := state["fields"].([]map[string]any)[0]
	if field["collection_or_dictionary"] != true {
		t.Fatalf("late-bound dictionary metrics = %#v", field)
	}
}

func TestVBA240RecognizesLateBoundDictionaryWhenMutatorPrecedesInitializer(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private sharedItems As Object

Public Sub Other()
  sharedItems.Add "other", 2
End Sub

Public Sub Run()
  Set sharedItems = CreateObject("Scripting.Dictionary")
  sharedItems.Add "run", 1
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA240")
	if len(findings) != 1 || !strings.Contains(findings[0].Reason, "Collection or Dictionary") {
		t.Fatalf("late-bound dictionary findings with mutator first = %+v", findings)
	}
	state := result.AnalysisMetrics.(map[string]any)["module_state"].(map[string]any)
	field := state["fields"].([]map[string]any)[0]
	if field["collection_or_dictionary"] != true || field["mutator_count"] != 2 {
		t.Fatalf("late-bound dictionary metrics with mutator first = %#v", field)
	}
}

func TestVBA240ProvidesCachedExcelReferenceGuidance(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private cachedBook As Workbook

Public Sub Run()
  Debug.Print cachedBook.Name
End Sub

Public Sub Reset()
  Set cachedBook = Nothing
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA240")
	if len(findings) != 1 || !strings.Contains(findings[0].Reason, "cached Excel object") || !strings.Contains(findings[0].Suggestion, "Reacquire") {
		t.Fatalf("cached Excel guidance = %+v", findings)
	}
	state := result.AnalysisMetrics.(map[string]any)["module_state"].(map[string]any)
	field := state["fields"].([]map[string]any)[0]
	if field["cached_excel_reference"] != true {
		t.Fatalf("cached Excel metrics = %#v", field)
	}
}

func TestVBA240ReportsStateRetainedAcrossDocumentEvents(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	workbookDir := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbookDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbookDir, "ThisWorkbook.cls"), []byte(`Attribute VB_Name = "ThisWorkbook"
Option Explicit
Private eventState As Long

Private Sub Workbook_Open()
  InitializeEventState
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  ReportEventState
End Sub

Private Sub InitializeEventState()
  eventState = 1
End Sub

Private Sub ReportEventState()
  Debug.Print eventState
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA240")
	if len(findings) != 1 || !strings.Contains(findings[0].Reason, "retained across event") {
		t.Fatalf("event state findings = %+v", findings)
	}
}

func TestVBA240ReportsMutableStateInsideCallCycle(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Private cycleState As Long

Public Sub Run()
  cycleState = cycleState + 1
  Helper
End Sub

Private Sub Helper()
  cycleState = cycleState + 1
  Run
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	findings := findingsByCode(result.Findings, "VBA240")
	if len(findings) != 1 || !strings.Contains(findings[0].Reason, "call cycle") {
		t.Fatalf("cycle state findings = %+v", findings)
	}
}

func TestVBA240IndexesStandardClassDocumentAndFormModules(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Standard.bas", `Option Explicit
Private standardState As Long
Public Sub Run()
  standardState = 1
End Sub
`)
	writeClass(t, dir, "State.cls", `Option Explicit
Private classState As Long
Public Sub Touch()
  classState = 1
End Sub
`)
	workbookDir := filepath.Join(dir, "src", "workbook")
	if err := os.MkdirAll(workbookDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(workbookDir, "ThisWorkbook.cls"), []byte(`Attribute VB_Name = "ThisWorkbook"
Option Explicit
Private workbookState As Long
Private Sub Workbook_Open()
  workbookState = 1
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	writeFormSidecar(t, dir, "UserForm1.bas", `Option Explicit
Private formState As Long
Private Sub UserForm_Activate()
  formState = 1
End Sub
`)

	cfg := config.Default()
	cfg.Analyze.DetectRiskyModuleState = true
	result, err := (Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	metrics := result.AnalysisMetrics.(map[string]any)
	state := metrics["module_state"].(map[string]any)
	fields := state["fields"].([]map[string]any)
	kinds := map[string]bool{}
	for _, field := range fields {
		kinds[field["module_kind"].(string)] = true
	}
	for _, kind := range []string{"standard", "class", "document", "form"} {
		if !kinds[kind] {
			t.Fatalf("module kinds = %#v, missing %s", kinds, kind)
		}
	}
}
