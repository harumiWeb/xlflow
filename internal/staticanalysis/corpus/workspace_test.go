package corpus

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestMaterializeThirdPartyProjectsPreservesSourcesAndClassifications(t *testing.T) {
	manifest, corpusRoot, err := LoadManifest(filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	tempRoot := t.TempDir()
	for _, project := range manifest.Projects {
		if !project.Enabled {
			continue
		}
		projectRoot := filepath.Join(corpusRoot, filepath.FromSlash(project.Path))
		before, err := TreeDigest(projectRoot)
		if err != nil {
			t.Fatal(err)
		}
		workspace, err := MaterializeThirdPartyProject(corpusRoot, project, MaterializeOptions{TempRoot: tempRoot})
		if err != nil {
			t.Fatalf("materialize %s: %v", project.ID, err)
		}
		loaded, err := config.Load(workspace.Root)
		if err != nil {
			t.Fatalf("load generated config for %s: %v", project.ID, err)
		}
		if loaded.UserForm.CodeSource != "frm" {
			t.Fatalf("generated userform code source = %q", loaded.UserForm.CodeSource)
		}
		if project.Profile == ProfileExcel && len(loaded.Analyze.DisabledRules) != 0 {
			t.Fatalf("excel profile unexpectedly disabled rules: %v", loaded.Analyze.DisabledRules)
		}
		if project.Profile != ProfileExcel && len(loaded.Analyze.DisabledRules) == 0 {
			t.Fatalf("%s profile did not apply non-Excel policy", project.Profile)
		}
		if len(workspace.Mappings) == 0 {
			t.Fatalf("project %s produced no source mappings", project.ID)
		}
		for workspacePath, sourcePath := range workspace.Mappings {
			body, err := os.ReadFile(filepath.Join(workspace.Root, filepath.FromSlash(workspacePath)))
			if err != nil {
				t.Fatal(err)
			}
			original, err := os.ReadFile(filepath.Join(projectRoot, filepath.FromSlash(sourcePath)))
			if err != nil {
				t.Fatal(err)
			}
			if string(body) != string(original) {
				t.Fatalf("source bytes changed for %s: %s", project.ID, sourcePath)
			}
		}
		if err := workspace.Close(); err != nil {
			t.Fatal(err)
		}
		if _, err := os.Stat(workspace.Root); !errors.Is(err, os.ErrNotExist) {
			t.Fatalf("workspace was not removed, stat error = %v", err)
		}
		after, err := TreeDigest(projectRoot)
		if err != nil {
			t.Fatal(err)
		}
		if before != after {
			t.Fatalf("materialization changed source tree for %s", project.ID)
		}
	}
	entries, err := os.ReadDir(tempRoot)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 0 {
		t.Fatalf("materializer left temporary children: %v", entries)
	}
}

func TestMaterializeThirdPartyProjectUsesDocumentOverrides(t *testing.T) {
	manifest, corpusRoot, err := LoadManifest(filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	tempRoot := t.TempDir()
	for _, want := range []struct {
		id, source, destination string
	}{
		{"access-examples", "Form_001_About_frm.cls", "src/workbook/Form_001_About_frm.cls"},
		{"vba-web", "examples/analytics/AnalyticsSheet.cls", "src/workbook/examples/analytics/AnalyticsSheet.cls"},
	} {
		var project Project
		for _, candidate := range manifest.Projects {
			if candidate.ID == want.id {
				project = candidate
			}
		}
		workspace, err := MaterializeThirdPartyProject(corpusRoot, project, MaterializeOptions{TempRoot: tempRoot})
		if err != nil {
			t.Fatal(err)
		}
		found := false
		for workspacePath, sourcePath := range workspace.Mappings {
			if sourcePath == want.source {
				found = workspacePath == want.destination
			}
		}
		if !found {
			t.Fatalf("%s was not mapped to %s: %v", want.source, want.destination, workspace.Mappings)
		}
		if err := workspace.Close(); err != nil {
			t.Fatal(err)
		}
	}
}

func TestMaterializeThirdPartyProjectRejectsMissingClassificationAndCleansUp(t *testing.T) {
	manifest, corpusRoot, err := LoadManifest(filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	project := manifest.Projects[1]
	project.Classifications = []Classification{{Path: "missing.cls", Kind: ModuleKindClass}}
	tempRoot := t.TempDir()
	if _, err := MaterializeThirdPartyProject(corpusRoot, project, MaterializeOptions{TempRoot: tempRoot}); err == nil || !errors.Is(err, ErrUnsupportedProjectLayout) {
		t.Fatalf("missing classification error = %v", err)
	}
	entries, err := os.ReadDir(tempRoot)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 0 {
		t.Fatalf("failed materialization left temporary children: %v", entries)
	}
}

func TestMaterializeThirdPartyProjectRejectsSymlinkSource(t *testing.T) {
	corpusRoot := t.TempDir()
	sourceRoot := filepath.Join(corpusRoot, "projects", "third_party", "fixture")
	if err := os.MkdirAll(sourceRoot, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(sourceRoot, "Main.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(filepath.Join(sourceRoot, "Main.bas"), filepath.Join(sourceRoot, "Link.bas")); err != nil {
		t.Skipf("symlink creation is unavailable: %v", err)
	}
	project := Project{ID: "fixture", Path: "projects/third_party/fixture", Profile: ProfileGenericVBA, Enabled: true}
	if _, err := MaterializeThirdPartyProject(corpusRoot, project, MaterializeOptions{TempRoot: t.TempDir()}); err == nil || !errors.Is(err, ErrUnsupportedProjectLayout) {
		t.Fatalf("symlink source error = %v", err)
	}
}

func TestThirdPartyProfileFilteringAndStableRemap(t *testing.T) {
	mappings := map[string]string{"src/modules/Main.bas": "nested/Main.bas"}
	lintDiagnostics, err := normalizeThirdPartyLintDiagnostics("third_party/access", ProfileAccess, mappings, []lint.Issue{
		{Code: "VB002", Severity: "warning", File: "src/modules/Main.bas", Line: 1},
		{Code: "VB001", Severity: "warning", File: "src/modules/Main.bas", Line: 2},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(lintDiagnostics) != 1 || lintDiagnostics[0].Code != "VB001" || lintDiagnostics[0].File != "nested/Main.bas" {
		t.Fatalf("unexpected filtered lint diagnostics: %+v", lintDiagnostics)
	}
	analyzeDiagnostics, err := normalizeThirdPartyAnalyzeDiagnostics("third_party/access", ProfileAccess, mappings, []analyze.Finding{
		{Code: "VBA104", Severity: "error", File: "src/modules/Main.bas", Line: 1},
		{Code: "VBA219", Severity: "warning", File: "src/modules/Main.bas", Line: 3},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(analyzeDiagnostics) != 1 || analyzeDiagnostics[0].Code != "VBA219" {
		t.Fatalf("unexpected filtered analyze diagnostics: %+v", analyzeDiagnostics)
	}
	if _, err := normalizeThirdPartyLintDiagnostics("third_party/access", ProfileAccess, mappings, []lint.Issue{{Code: "VB001", File: "temporary/path.bas"}}); err == nil || !strings.Contains(err.Error(), "not in the materialized source map") {
		t.Fatalf("unmapped path error = %v", err)
	}
}

func TestRunThirdPartyProjectsReportsSkipsAndRemapsDiagnostics(t *testing.T) {
	manifest, corpusRoot, err := LoadManifest(filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	projects := []Project{manifest.Projects[0], manifest.Projects[1]}
	projects[0].Enabled = false
	projects[0].Notes = "not included by this test"
	report, err := RunThirdPartyProjects(corpusRoot, projects, MaterializeOptions{TempRoot: t.TempDir()})
	if err != nil {
		t.Fatal(err)
	}
	if len(report.Skipped) != 1 || report.Skipped[0].Project != "third_party/access-examples" {
		t.Fatalf("unexpected skips: %+v", report.Skipped)
	}
	if len(report.Surfaces) != 2 {
		t.Fatalf("surfaces = %+v", report.Surfaces)
	}
	for _, diagnostic := range report.Diagnostics {
		if strings.Contains(diagnostic.File, "xlflow-corpus-") || filepath.IsAbs(filepath.FromSlash(diagnostic.File)) {
			t.Fatalf("diagnostic leaked temporary path: %+v", diagnostic)
		}
	}
}

func TestRunThirdPartyProjectsPreservesPartialResultsAfterWorkspaceFailure(t *testing.T) {
	manifest, corpusRoot, err := LoadManifest(filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	broken := manifest.Projects[1]
	broken.ID = "broken"
	broken.Path = "projects/third_party/missing"
	report, err := RunThirdPartyProjects(corpusRoot, []Project{broken, manifest.Projects[1]}, MaterializeOptions{TempRoot: t.TempDir()})
	if err == nil {
		t.Fatal("expected workspace failure")
	}
	if len(report.Failures) != 1 || report.Failures[0].Kind != FailureWorkspace {
		t.Fatalf("unexpected failures: %+v", report.Failures)
	}
	if len(report.Diagnostics) == 0 || len(report.Surfaces) != 2 {
		t.Fatalf("successful project results were not preserved: %+v", report)
	}
}

func TestMaterializedWorkspaceClassifiesFilesForProductionSymbols(t *testing.T) {
	manifest, corpusRoot, err := LoadManifest(filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	project := manifest.Projects[1]
	workspace, err := MaterializeThirdPartyProject(corpusRoot, project, MaterializeOptions{TempRoot: t.TempDir()})
	if err != nil {
		t.Fatal(err)
	}
	defer func() {
		if err := workspace.Close(); err != nil {
			t.Errorf("close materialized workspace: %v", err)
		}
	}()
	cfg, err := config.Load(workspace.Root)
	if err != nil {
		t.Fatal(err)
	}
	files, err := symbols.DiscoverSourceFiles(symbols.Options{RootDir: workspace.Root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range files {
		if file.ModuleKind == "" {
			t.Fatalf("source file has no module kind: %+v", file)
		}
	}
}
