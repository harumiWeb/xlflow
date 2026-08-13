package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

func TestRootCommandIncludesMetricsCommand(t *testing.T) {
	root := (&app{}).rootCommand()
	cmd, _, err := root.Find([]string{"metrics"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "metrics" {
		t.Fatalf("metrics command = %#v", cmd)
	}
}

func TestCollectProcedureMetricsResolvesProjectFanOutAndSorts(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Main.bas"), []byte(`Attribute VB_Name = "Main"
Public Sub Run()
    If True Then
        Helper
    Else
        Helper
    End If
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Helper.bas"), []byte(`Attribute VB_Name = "Helper"
Public Sub Helper()
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Src.Modules = filepath.ToSlash(filepath.Join("src", "modules"))
	got, warnings, err := collectProcedureMetrics(context.Background(), root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	if len(warnings) != 0 || len(got) != 2 {
		t.Fatalf("metrics = %+v, warnings = %+v", got, warnings)
	}
	if got[0].Name != "Helper" || got[1].Name != "Run" {
		t.Fatalf("procedure order = %q, %q", got[0].Name, got[1].Name)
	}
	if got[1].CallFanOut != 1 || got[1].BranchCount != 1 || got[1].CyclomaticComplexity != 2 {
		t.Fatalf("Run metrics = %+v", got[1].Metrics)
	}
}

func TestDiscoverMetricsSourceFilesRespectsUserFormCodeSourceForTests(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "tests", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	frmPath := filepath.Join(formsDir, "Fixture.frm")
	sidecarPath := filepath.Join(codeDir, "Fixture.bas")
	if err := os.WriteFile(frmPath, []byte("VERSION 5.00\nBegin VB.Form Fixture\nEnd\nAttribute VB_Name = \"Fixture\"\nPrivate Sub Embedded()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(sidecarPath, []byte("Option Explicit\nPrivate Sub Sidecar()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	standardPath := filepath.Join(root, "tests", "Keep.bas")
	classPath := filepath.Join(root, "tests", "Keep.cls")
	for path, source := range map[string]string{
		standardPath: "Public Sub KeepStandard()\nEnd Sub\n",
		classPath:    "Public Sub KeepClass()\nEnd Sub\n",
	} {
		if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}

	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	files, err := discoverMetricsSourceFiles(root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	got := metricsSourceKinds(files)
	if _, ok := got[filepath.Clean(frmPath)]; ok {
		t.Fatalf("sidecar mode included generated form: %+v", got)
	}
	if got[filepath.Clean(sidecarPath)] != "form" {
		t.Fatalf("sidecar mode classified form code as %q: %+v", got[filepath.Clean(sidecarPath)], got)
	}
	if got[filepath.Clean(standardPath)] != "standard" || got[filepath.Clean(classPath)] != "class" {
		t.Fatalf("sidecar mode changed ordinary test sources: %+v", got)
	}

	cfg.UserForm.CodeSource = "frm"
	files, err = discoverMetricsSourceFiles(root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	got = metricsSourceKinds(files)
	if got[filepath.Clean(frmPath)] != "form" {
		t.Fatalf("frm mode did not include embedded form: %+v", got)
	}
	if _, ok := got[filepath.Clean(sidecarPath)]; ok {
		t.Fatalf("frm mode included sidecar form code: %+v", got)
	}
	if got[filepath.Clean(standardPath)] != "standard" || got[filepath.Clean(classPath)] != "class" {
		t.Fatalf("frm mode changed ordinary test sources: %+v", got)
	}
}

func metricsSourceKinds(files []symbols.SourceFile) map[string]string {
	result := make(map[string]string, len(files))
	for _, file := range files {
		result[filepath.Clean(file.Path)] = file.ModuleKind
	}
	return result
}

func TestMetricsCommandJSONThresholdFailureRetainsMetrics(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Main.bas"), []byte("Attribute VB_Name = \"Main\"\nPublic Sub Run()\n    If True Then\n    End If\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(root, config.FileName), []byte(`[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"
[metrics.thresholds]
branch_count = 0
cyclomatic_complexity = 1
`), 0o644); err != nil {
		t.Fatal(err)
	}
	var stdout, stderr strings.Builder
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	rootCmd := a.rootCommand()
	rootCmd.SetArgs([]string{"--json", "metrics"})
	if err := rootCmd.Execute(); output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("metrics exit = %v (%v), stdout=%s", output.ExitCode(err), err, stdout.String())
	}
	var payload map[string]any
	if err := json.Unmarshal([]byte(stdout.String()), &payload); err != nil {
		t.Fatalf("JSON = %v\n%s", err, stdout.String())
	}
	if payload["metrics"] == nil || payload["diagnostics"] == nil {
		t.Fatalf("threshold failure omitted metrics/diagnostics: %s", stdout.String())
	}
	if payload["status"] != output.StatusFailed || payload["error"].(map[string]any)["code"] != "metrics_threshold_exceeded" {
		t.Fatalf("payload status/error = %#v", payload)
	}
}

func TestMetricsCommandHotspotsReportAndSelectors(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	sources := map[string]string{
		"Main.bas":   "Attribute VB_Name = \"Main\"\nPublic Sub Run()\n If True Then\n End If\nEnd Sub\n",
		"Helper.bas": "Attribute VB_Name = \"Helper\"\nPublic Sub Help()\nEnd Sub\n",
	}
	for name, source := range sources {
		if err := os.WriteFile(filepath.Join(moduleDir, name), []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	if err := os.WriteFile(filepath.Join(root, config.FileName), []byte(`[project]
entry = "Main.Run"
[excel]
path = "build/Book.xlsm"
[src]
modules = "src/modules"
[metrics.hotspots]
procedure_top_n = 1
`), 0o644); err != nil {
		t.Fatal(err)
	}
	var stdout, stderr strings.Builder
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "metrics"})
	if err := cmd.Execute(); err != nil {
		t.Fatalf("top-n metrics failed: %v\n%s", err, stdout.String())
	}
	var payload map[string]any
	if err := json.Unmarshal([]byte(stdout.String()), &payload); err != nil {
		t.Fatal(err)
	}
	metrics := payload["metrics"].(map[string]any)
	hotspots := metrics["hotspots"].(map[string]any)
	if len(hotspots["procedures"].([]any)) != 2 || len(payload["diagnostics"].([]any)) != 1 {
		t.Fatalf("hotspot payload = %#v", payload)
	}
	diagnostic := payload["diagnostics"].([]any)[0].(map[string]any)
	if diagnostic["code"] != "MX002" || diagnostic["severity"] != "information" {
		t.Fatalf("hotspot diagnostic = %#v", diagnostic)
	}
}

func TestMetricsCommandHotspotThresholdFailure(t *testing.T) {
	root := writeHotspotMetricsProject(t, map[string]string{
		"Main.bas":   "Attribute VB_Name = \"Main\"\nPublic Sub Run()\n If True Then\n End If\nEnd Sub\n",
		"Helper.bas": "Attribute VB_Name = \"Helper\"\nPublic Sub Help()\nEnd Sub\n",
	}, "procedure_score_threshold = 1")
	var stdout, stderr strings.Builder
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "metrics"})
	if err := cmd.Execute(); output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("threshold metrics exit = %v (%v)\n%s", output.ExitCode(err), err, stdout.String())
	}
	var payload map[string]any
	if err := json.Unmarshal([]byte(stdout.String()), &payload); err != nil {
		t.Fatal(err)
	}
	if payload["status"] != output.StatusFailed || payload["error"].(map[string]any)["code"] != "metrics_hotspot_threshold_exceeded" {
		t.Fatalf("threshold payload = %#v", payload)
	}
	diagnostics := payload["diagnostics"].([]any)
	if len(diagnostics) == 0 || diagnostics[0].(map[string]any)["severity"] != "warning" {
		t.Fatalf("threshold diagnostics = %#v", diagnostics)
	}
	selectedBy := diagnostics[0].(map[string]any)["selected_by"].(map[string]any)
	if selectedBy["threshold"] != true {
		t.Fatalf("threshold selection = %#v", selectedBy)
	}
}

func TestMetricsCommandUnmetHotspotThresholdWithTopNIsSuccessful(t *testing.T) {
	root := writeHotspotMetricsProject(t, map[string]string{
		"Main.bas": "Attribute VB_Name = \"Main\"\nPublic Sub Run()\nEnd Sub\n",
	}, "procedure_top_n = 1\nprocedure_score_threshold = 1")
	var stdout, stderr strings.Builder
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "metrics"})
	if err := cmd.Execute(); err != nil {
		t.Fatalf("top-n with unmet threshold failed: %v\n%s", err, stdout.String())
	}
	var payload map[string]any
	if err := json.Unmarshal([]byte(stdout.String()), &payload); err != nil {
		t.Fatal(err)
	}
	if payload["status"] != output.StatusOK || payload["error"] != nil {
		t.Fatalf("unmet threshold payload = %#v", payload)
	}
	diagnostics := payload["diagnostics"].([]any)
	if len(diagnostics) != 1 || diagnostics[0].(map[string]any)["severity"] != "information" {
		t.Fatalf("top-n diagnostics = %#v", diagnostics)
	}
}

func TestMetricsCommandPrefersComplexityThresholdErrorWhenCombined(t *testing.T) {
	root := writeHotspotMetricsProject(t, map[string]string{
		"Main.bas":   "Attribute VB_Name = \"Main\"\nPublic Sub Run()\n If True Then\n End If\nEnd Sub\n",
		"Helper.bas": "Attribute VB_Name = \"Helper\"\nPublic Sub Help()\nEnd Sub\n",
	}, "procedure_score_threshold = 1")
	path := filepath.Join(root, config.FileName)
	contents, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	contents = []byte(strings.Replace(string(contents), "[metrics.hotspots]", "[metrics.thresholds]\ncyclomatic_complexity = 1\n[metrics.hotspots]", 1))
	if err := os.WriteFile(path, contents, 0o644); err != nil {
		t.Fatal(err)
	}
	var stdout, stderr strings.Builder
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "metrics"})
	if err := cmd.Execute(); output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("combined thresholds exit = %v (%v)\n%s", output.ExitCode(err), err, stdout.String())
	}
	var payload map[string]any
	if err := json.Unmarshal([]byte(stdout.String()), &payload); err != nil {
		t.Fatal(err)
	}
	if payload["error"].(map[string]any)["code"] != "metrics_threshold_exceeded" {
		t.Fatalf("combined threshold error = %#v", payload["error"])
	}
	diagnostics := payload["diagnostics"].([]any)
	if len(diagnostics) < 2 || diagnostics[0].(map[string]any)["code"] != "MX001" {
		t.Fatalf("combined diagnostics ordering = %#v", diagnostics)
	}
}

func writeHotspotMetricsProject(t *testing.T, sources map[string]string, hotspotConfig string) string {
	t.Helper()
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	for name, source := range sources {
		if err := os.WriteFile(filepath.Join(moduleDir, name), []byte(source), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	configText := fmt.Sprintf(`[project]
entry = "Main.Run"
[excel]
path = "build/Book.xlsm"
[src]
modules = "src/modules"
[metrics.hotspots]
%s
`, hotspotConfig)
	if err := os.WriteFile(filepath.Join(root, config.FileName), []byte(configText), 0o644); err != nil {
		t.Fatal(err)
	}
	return root
}

func TestBuildHotspotReportAggregatesGraphSignals(t *testing.T) {
	metrics := []proceduremetrics.ProcedureMetrics{
		{File: "main.bas", Module: "Main", Name: "A", Kind: procedureir.ProcedureSub, ResolvedCallees: []string{"helper.b"}, Metrics: proceduremetrics.Metrics{CyclomaticComplexity: 2, CallFanOut: 1}},
		{File: "helper.bas", Module: "Helper", Name: "B", Kind: procedureir.ProcedureSub, ResolvedCallees: []string{"main.a"}, Metrics: proceduremetrics.Metrics{CyclomaticComplexity: 3, CallFanOut: 1}},
		{File: "other.bas", Module: "Other", Name: "C", Kind: procedureir.ProcedureSub, ResolvedCallees: []string{"helper.b"}, Metrics: proceduremetrics.Metrics{CyclomaticComplexity: 1, CallFanOut: 1}},
	}
	report := buildHotspotReport(metrics)
	byProcedure := map[string]hotspots.Entity{}
	for _, entity := range report.Procedures {
		byProcedure[entity.Name] = entity
	}
	if byProcedure["A"].RawSignals["call_fan_in"] != 1 || byProcedure["A"].RawSignals["cycle_count"] != 1 {
		t.Fatalf("procedure A graph signals = %#v", byProcedure["A"].RawSignals)
	}
	if byProcedure["A"].RawSignals["affected_module_count"] != 2 {
		t.Fatalf("procedure A affected modules = %#v", byProcedure["A"].RawSignals)
	}
	byModule := map[string]hotspots.Entity{}
	for _, entity := range report.Modules {
		byModule[entity.Name] = entity
	}
	if byModule["Helper"].RawSignals["call_fan_in"] != 2 || byModule["Helper"].RawSignals["call_fan_out"] != 1 {
		t.Fatalf("Helper module graph signals = %#v", byModule["Helper"].RawSignals)
	}
	if byModule["Main"].RawSignals["affected_module_count"] != 2 {
		t.Fatalf("Main module affected modules = %#v", byModule["Main"].RawSignals)
	}
}

func TestCycleParticipationCountsIsDeterministicAndBounded(t *testing.T) {
	build := func(reverse bool) map[string]map[string]bool {
		graph := make(map[string]map[string]bool)
		for offset := 0; offset < 10; offset++ {
			i := offset
			if reverse {
				i = 9 - offset
			}
			from := fmt.Sprintf("n%02d", i)
			graph[from] = map[string]bool{}
			for targetOffset := 0; targetOffset < 10; targetOffset++ {
				j := targetOffset
				if reverse {
					j = 9 - targetOffset
				}
				if i == j {
					continue
				}
				to := fmt.Sprintf("n%02d", j)
				graph[from][to] = true
			}
		}
		return graph
	}
	first := cycleParticipationCounts(build(false))
	second := cycleParticipationCounts(build(true))
	if !reflect.DeepEqual(first, second) {
		t.Fatalf("cycle counts changed with input order: %#v vs %#v", first, second)
	}
	for i := 0; i < 10; i++ {
		if first[fmt.Sprintf("n%02d", i)] == 0 {
			t.Fatalf("node n%02d did not participate in a cycle: %#v", i, first)
		}
	}
}
