package corpus

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"
)

func TestBuildInventoryCountsSourcesAndProfilesDeterministically(t *testing.T) {
	root := t.TempDir()
	manifest := inventoryFixtureManifest()
	writeInventorySources(t, root, manifest.Projects[0], map[string]string{
		"Main.bas":         "Attribute VB_Name = \"Main\"\n",
		"Nested/Thing.cls": "Attribute VB_Name = \"Thing\"\n",
		"Dialog.frm":       "VERSION 5.00\n",
		"LICENSE":          "MIT\n",
	})
	writeInventorySources(t, root, manifest.Projects[1], map[string]string{
		"A.bas": "Attribute VB_Name = \"A\"\n",
		"B.BAS": "Attribute VB_Name = \"B\"\n",
	})

	first, err := BuildInventory(root, manifest)
	if err != nil {
		t.Fatal(err)
	}
	second, err := BuildCorpusInventory(root, manifest)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first, second) {
		t.Fatalf("inventory changed between runs:\n%#v\n%#v", first, second)
	}
	if first.ProjectCount != 2 || first.EnabledCount != 1 || first.DisabledCount != 1 {
		t.Fatalf("project counts = %#v", first)
	}
	if !reflect.DeepEqual(first.Sources, SourceInventory{Bas: 3, Cls: 1, Frm: 1}) {
		t.Fatalf("source counts = %#v", first.Sources)
	}
	if got := []string{first.Profiles[0].Profile, first.Profiles[1].Profile}; !reflect.DeepEqual(got, []string{ProfileExcel, ProfileGenericVBA}) {
		t.Fatalf("profiles = %v", got)
	}
	wantLog := "corpus inventory: projects=2 enabled=1 disabled=1 bas=3 cls=1 frm=1 total=5\n" +
		"  profile=excel projects=1\n" +
		"  profile=generic-vba projects=1\n" +
		"  project=one profile=generic-vba enabled=true bas=1 cls=1 frm=1 total=3\n" +
		"  project=two profile=excel enabled=false bas=2 cls=0 frm=0 total=2\n"
	if got := FormatInventorySummary(first); got != wantLog {
		t.Fatalf("inventory summary = %q, want %q", got, wantLog)
	}
}

func TestValidateInventoryReportsCountAndStaleProjectFailures(t *testing.T) {
	root := t.TempDir()
	manifest := inventoryFixtureManifest()
	writeInventorySources(t, root, manifest.Projects[0], map[string]string{"Main.bas": ""})
	writeInventorySources(t, root, manifest.Projects[1], map[string]string{"Main.bas": "", "Thing.cls": ""})
	if err := os.MkdirAll(filepath.Join(root, "projects", "third_party", "stale"), 0o755); err != nil {
		t.Fatal(err)
	}
	report, err := ValidateInventory(root, manifest)
	if err == nil {
		t.Fatal("ValidateInventory accepted mismatched source counts")
	}
	var validation *InventoryValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("validation error = %T %v", err, err)
	}
	if len(validation.Mismatches) != 2 || !reflect.DeepEqual(validation.Unexpected, []string{"stale"}) {
		t.Fatalf("validation details = %#v", validation)
	}
	if report.ProjectCount != len(manifest.Projects) {
		t.Fatalf("report was not retained on validation error: %#v", report)
	}
}

func TestValidateInventoryReportsMissingProjectWithCompleteReport(t *testing.T) {
	root := t.TempDir()
	manifest := inventoryFixtureManifest()
	writeInventorySources(t, root, manifest.Projects[0], map[string]string{"Main.bas": ""})
	report, err := ValidateInventory(root, manifest)
	if err == nil {
		t.Fatal("ValidateInventory accepted missing project")
	}
	var validation *InventoryValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("validation error = %T %v", err, err)
	}
	if !reflect.DeepEqual(validation.Missing, []string{"two"}) {
		t.Fatalf("missing projects = %v", validation.Missing)
	}
	if report.ProjectCount != 2 || !reflect.DeepEqual(report.Missing, []string{"two"}) {
		t.Fatalf("incomplete report = %#v", report)
	}
}

func TestSummarizeDiagnosticsSortsRulesAndProjectsAndPreservesDuplicates(t *testing.T) {
	report := Report{Diagnostics: []Diagnostic{
		{Project: "self/z", Surface: SurfaceAnalyze, Code: "VBA209"},
		{Project: "self/a", Surface: SurfaceLint, Code: "VBA209"},
		{Project: "self/a", Surface: SurfaceLint, Code: "VBA209"},
		{Project: "third_party/vba", Surface: SurfaceAnalyze, Code: "VBA201"},
	}}
	summary := SummarizeDiagnostics(report)
	if summary.Total != 4 || summary.Lint != 2 || summary.Analyze != 2 {
		t.Fatalf("summary totals = %#v", summary)
	}
	if len(summary.Rules) != 2 || summary.Rules[0].Code != "VBA201" || summary.Rules[1].Code != "VBA209" {
		t.Fatalf("rule order = %#v", summary.Rules)
	}
	if summary.Rules[1].Total != 3 || summary.Rules[1].Lint != 2 || summary.Rules[1].Analyze != 1 {
		t.Fatalf("VBA209 aggregate = %#v", summary.Rules[1])
	}
	if !reflect.DeepEqual(summary.Rules[1].Projects, []ProjectDiagnosticCount{{Project: "self/a", Count: 2}, {Project: "self/z", Count: 1}}) {
		t.Fatalf("VBA209 projects = %#v", summary.Rules[1].Projects)
	}
	wantLog := "corpus diagnostics: total=4 lint=2 analyze=2 rules=2\n" +
		"  rule=VBA201 total=1 lint=0 analyze=1 projects=third_party/vba:1\n" +
		"  rule=VBA209 total=3 lint=2 analyze=1 projects=self/a:2,self/z:1\n"
	if got := FormatDiagnosticSummary(summary); got != wantLog {
		t.Fatalf("diagnostic summary = %q, want %q", got, wantLog)
	}
	if got := FormatDiagnosticSummary(BuildDiagnosticSummary(report)); got != wantLog {
		t.Fatalf("BuildDiagnosticSummary output = %q", got)
	}
}

func inventoryFixtureManifest() Manifest {
	return Manifest{
		SchemaVersion: SchemaVersion,
		Projects: []Project{
			{ID: "one", Path: "projects/third_party/one", Profile: ProfileGenericVBA, Enabled: true, SourceCounts: SourceCounts{Bas: 1, Cls: 1, Frm: 1}},
			{ID: "two", Path: "projects/third_party/two", Profile: ProfileExcel, Enabled: false, SourceCounts: SourceCounts{Bas: 2}},
		},
	}
}

func writeInventorySources(t *testing.T, root string, project Project, files map[string]string) {
	t.Helper()
	projectRoot := filepath.Join(root, filepath.FromSlash(project.Path))
	for name, contents := range files {
		path := filepath.Join(projectRoot, filepath.FromSlash(name))
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(contents), 0o644); err != nil {
			t.Fatal(err)
		}
	}
}
