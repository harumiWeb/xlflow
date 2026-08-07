package corpus

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
)

func TestDiscoverExampleProjectsAndSelectByID(t *testing.T) {
	repo := t.TempDir()
	for _, name := range []string{"zeta", "alpha"} {
		root := filepath.Join(repo, "example", name)
		if err := os.MkdirAll(root, 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(root, config.FileName), []byte("[project]\nentry = \"Main.Run\"\n"), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	if err := os.MkdirAll(filepath.Join(repo, "example", "not-a-project"), 0o755); err != nil {
		t.Fatal(err)
	}
	projects, err := DiscoverExampleProjects(repo)
	if err != nil {
		t.Fatal(err)
	}
	if len(projects) != 2 {
		t.Fatalf("discovered %d projects, want 2: %+v", len(projects), projects)
	}
	if got := []string{projects[0].ID, projects[1].ID}; !reflect.DeepEqual(got, []string{"self/alpha", "self/zeta"}) {
		t.Fatalf("discovered IDs = %v", got)
	}
	selected, err := SelectProjects(projects, []string{"self/zeta"})
	if err != nil || len(selected) != 1 || selected[0].ID != "self/zeta" {
		t.Fatalf("selected = %+v, err = %v", selected, err)
	}
	if _, err := SelectProjects(projects, []string{"self/missing"}); err == nil {
		t.Fatal("unknown project ID was accepted")
	}
	if _, err := SelectProjects(projects, []string{"self/alpha", "self/alpha"}); err == nil {
		t.Fatal("duplicate selected project ID was accepted")
	}
}

func TestRunProjectsClassifiesFailuresAndKeepsPartialResults(t *testing.T) {
	projects := []NativeProject{
		{ID: "self/config", Root: "config"},
		{ID: "self/parser", Root: "parser"},
		{ID: "self/execution", Root: "execution"},
		{ID: "self/ok", Root: "ok"},
	}
	exec := executor{
		loadConfig: func(root string) (config.Config, error) {
			if root == "config" {
				return config.Config{}, errors.New("invalid xlflow.toml")
			}
			return config.Default(), nil
		},
		runLint: func(root string, _ config.Config) (lint.Result, error) {
			if root == "execution" {
				return lint.Result{}, errors.New("lint filesystem failure")
			}
			return lint.Result{Issues: []lint.Issue{{Code: "VB001", Severity: "warning", File: "src/modules/Main.bas", Line: 2}}}, nil
		},
		runAnalyze: func(root string, _ config.Config) (analyze.Result, error) {
			if root == "parser" {
				return analyze.Result{}, &analyze.ParseError{Path: "src/modules/Main.bas", HasError: true}
			}
			return analyze.Result{Findings: []analyze.Finding{{Code: "VBA201", Severity: "warning", File: "src/modules/Main.bas", Line: 3}}}, nil
		},
	}
	report, err := runProjects(projects, exec)
	if err == nil {
		t.Fatal("expected aggregate RunError")
	}
	var runErr *RunError
	if !errors.As(err, &runErr) || len(runErr.Failures) != 3 {
		t.Fatalf("aggregate error = %T %+v", err, err)
	}
	if len(report.Diagnostics) != 4 {
		t.Fatalf("diagnostics = %+v", report.Diagnostics)
	}
	if len(report.Surfaces) != 6 {
		t.Fatalf("surface results = %+v", report.Surfaces)
	}
	var parser, execution, configFailure bool
	for _, failure := range report.Failures {
		switch failure.Kind {
		case FailureParser:
			parser = true
		case FailureExecution:
			execution = true
		case FailureInvalidProjectConfig:
			configFailure = true
		}
	}
	if !parser || !execution || !configFailure {
		t.Fatalf("failure kinds = %+v", report.Failures)
	}
	if report.Diagnostics[0].Project != "self/execution" || report.Diagnostics[0].Surface != SurfaceAnalyze || report.Diagnostics[0].File != "src/modules/Main.bas" {
		t.Fatalf("diagnostics are not normalized/sorted: %+v", report.Diagnostics)
	}
}

func TestRunProjectsDiagnosticsAreDeterministic(t *testing.T) {
	projects := []NativeProject{{ID: "self/sample", Root: "sample"}}
	exec := executor{
		loadConfig: func(string) (config.Config, error) { return config.Default(), nil },
		runLint: func(string, config.Config) (lint.Result, error) {
			return lint.Result{Issues: []lint.Issue{
				{Code: "VB002", Severity: "warning", File: "b.bas", Line: 4},
				{Code: "VB001", Severity: "warning", File: "a.bas", Line: 2},
			}}, nil
		},
		runAnalyze: func(string, config.Config) (analyze.Result, error) {
			return analyze.Result{Findings: []analyze.Finding{{Code: "VBA201", Severity: "warning", File: "a.bas", Line: 3}}}, nil
		},
	}
	first, err := runProjects(projects, exec)
	if err != nil {
		t.Fatal(err)
	}
	second, err := runProjects(projects, exec)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first.Diagnostics, second.Diagnostics) {
		t.Fatalf("diagnostics changed between runs:\n%+v\n%+v", first.Diagnostics, second.Diagnostics)
	}
	if first.Diagnostics[0].Surface != SurfaceLint || first.Diagnostics[0].File != "a.bas" || first.Diagnostics[2].Surface != SurfaceAnalyze {
		t.Fatalf("unexpected deterministic order: %+v", first.Diagnostics)
	}
}

func TestRunProjectsUsesEachProjectConfig(t *testing.T) {
	root := t.TempDir()
	projects := make([]NativeProject, 0, 2)
	for _, name := range []string{"enabled", "disabled"} {
		projectRoot := filepath.Join(root, name)
		if err := os.MkdirAll(filepath.Join(projectRoot, "src", "modules"), 0o755); err != nil {
			t.Fatal(err)
		}
		cfg := config.Default()
		if name == "disabled" {
			cfg.Lint.DisabledRules = []string{"VB002"}
		}
		if err := config.Write(filepath.Join(projectRoot, config.FileName), cfg); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(projectRoot, "src", "modules", "Main.bas"), []byte("Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n  Range(\"A1\").Select\nEnd Sub\n"), 0o644); err != nil {
			t.Fatal(err)
		}
		projects = append(projects, NativeProject{ID: "self/" + name, Root: projectRoot})
	}
	report, err := RunProjects(projects)
	if err != nil {
		t.Fatal(err)
	}
	counts := map[string]int{}
	for _, diagnostic := range report.Diagnostics {
		if diagnostic.Code == "VB002" {
			counts[diagnostic.Project]++
		}
	}
	if counts["self/enabled"] == 0 || counts["self/disabled"] != 0 {
		t.Fatalf("per-project config was not respected: %+v", counts)
	}
}

func TestRealWorldCorpusExampleGenQRCode(t *testing.T) {
	repoRoot, err := filepath.Abs(filepath.Join("..", "..", ".."))
	if err != nil {
		t.Fatal(err)
	}
	projects, err := DiscoverExampleProjects(repoRoot)
	if err != nil {
		t.Fatal(err)
	}
	selected, err := SelectProjects(projects, []string{"self/gen-qrcode"})
	if err != nil {
		t.Fatal(err)
	}
	projectTree := filepath.Join(repoRoot, "example", "gen-qrcode")
	before, err := TreeDigest(projectTree)
	if err != nil {
		t.Fatal(err)
	}
	first, err := RunProjects(selected)
	if err != nil {
		t.Fatal(err)
	}
	after, err := TreeDigest(projectTree)
	if err != nil {
		t.Fatal(err)
	}
	if before != after {
		t.Fatal("corpus runner mutated example/gen-qrcode")
	}
	if len(first.Surfaces) != 2 || first.Surfaces[0].Surface != SurfaceLint || first.Surfaces[1].Surface != SurfaceAnalyze {
		t.Fatalf("surfaces = %+v", first.Surfaces)
	}
	if len(first.Diagnostics) == 0 {
		t.Fatal("expected real-world analyze diagnostics")
	}
	second, err := RunProjects(selected)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first.Diagnostics, second.Diagnostics) {
		t.Fatal("real-world diagnostics were not deterministic")
	}
}
