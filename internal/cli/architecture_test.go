package cli

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/vba/architecture"
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
)

func TestArchitectureCommandSourceOnlyPolicy(t *testing.T) {
	cmd, _, err := (&app{}).rootCommand().Find([]string{"architecture"})
	if err != nil || cmd.Name() != "architecture" {
		t.Fatalf("architecture registration: %v, %v", cmd, err)
	}
	descriptor, err := coordination.LookupCLI("xlflow architecture")
	if err != nil || descriptor.RequiresExcel || descriptor.Policy.ResourceScope != coordination.ResourceNone || !descriptor.Policy.ParallelSafe {
		t.Fatalf("architecture policy: %+v, %v", descriptor, err)
	}
}

func TestArchitectureCommandIgnoresDiagnosticAndMetricsPolicies(t *testing.T) {
	root := architectureCLIProject(t, 2)
	configPath := filepath.Join(root, config.FileName)
	configText, err := os.ReadFile(configPath)
	if err != nil {
		t.Fatal(err)
	}
	configText = append(configText, []byte(`
[metrics]
exclude = ["src/modules/*.bas"]
[metrics.thresholds]
cyclomatic_complexity = 1
[metrics.hotspots]
procedure_score_threshold = 1
module_score_threshold = 1
`)...)
	if err := os.WriteFile(configPath, configText, 0o644); err != nil {
		t.Fatal(err)
	}
	first, code := runArchitectureCLI(t, root, "--json", "architecture")
	if code != output.ExitSuccess {
		t.Fatalf("exit=%d: %s", code, first)
	}
	var payload map[string]any
	if err := json.Unmarshal([]byte(first), &payload); err != nil {
		t.Fatal(err)
	}
	if payload["command"] != "architecture" || payload["status"] != output.StatusOK || len(payload["diagnostics"].([]any)) != 0 {
		t.Fatalf("envelope: %#v", payload)
	}
	report := payload["architecture"].(map[string]any)
	if report["schema_version"] != float64(1) || len(report["procedures"].([]any)) != 4 {
		t.Fatalf("report lost metrics-excluded source: %#v", report)
	}
	summary := report["summary"].(map[string]any)
	if summary["unreachable"] != float64(len(report["unreachable"].([]any))) || summary["confirmed_reachable"] == float64(0) {
		t.Fatalf("reachability summary disagrees with evidence: %#v", report)
	}
	if _, err := os.Stat(filepath.Join(root, ".xlflow")); !os.IsNotExist(err) {
		t.Fatalf("source-only command changed project state: %v", err)
	}
	second, code := runArchitectureCLI(t, root, "architecture", "--json")
	if code != output.ExitSuccess || second != first {
		t.Fatalf("repeated JSON differs: exit=%d\nfirst=%s\nsecond=%s", code, first, second)
	}
}

func TestArchitectureCommandFilterPreservesBoundaryAndRank(t *testing.T) {
	root := architectureCLIProject(t, 2)
	fullText, fullCode := runArchitectureCLI(t, root, "architecture", "--json")
	filteredText, filteredCode := runArchitectureCLI(t, root, "architecture", "--module", "main", "--path", "src/modules/Main.bas", "--json")
	if fullCode != 0 || filteredCode != 0 {
		t.Fatalf("full/filtered exits=%d/%d: %s\n%s", fullCode, filteredCode, fullText, filteredText)
	}
	decode := func(text string) map[string]any {
		t.Helper()
		var envelope map[string]any
		if err := json.Unmarshal([]byte(text), &envelope); err != nil {
			t.Fatal(err)
		}
		return envelope["architecture"].(map[string]any)
	}
	full, filtered := decode(fullText), decode(filteredText)
	if len(filtered["procedures"].([]any)) != 2 {
		t.Fatalf("display filter procedures: %#v", filtered["procedures"])
	}
	if len(filtered["unreachable"].([]any)) != 1 || filtered["summary"].(map[string]any)["unreachable"] != float64(1) {
		t.Fatalf("display filter lost unreachable evidence: %#v", filtered)
	}
	if !reflect.DeepEqual(full["summary"], filtered["project_summary"]) {
		t.Fatalf("filtered context lost whole-project summary")
	}
	dependencies := filtered["dependencies"].(map[string]any)
	if len(dependencies["edges"].([]any)) == 0 {
		t.Fatalf("filtered report lost boundary dependency: %#v", dependencies)
	}
	fullHotspots := full["hotspots"].(map[string]any)["procedures"].([]any)
	for _, scoped := range filtered["hotspots"].(map[string]any)["procedures"].([]any) {
		scoped := scoped.(map[string]any)
		found := false
		for _, item := range fullHotspots {
			candidate := item.(map[string]any)
			if candidate["id"] == scoped["id"] {
				found = true
				if candidate["rank"] != scoped["rank"] || candidate["score"] != scoped["score"] {
					t.Fatalf("filtered hotspot reranked: full=%#v scoped=%#v", candidate, scoped)
				}
			}
		}
		if !found {
			t.Fatalf("unknown filtered hotspot: %#v", scoped)
		}
	}
}

func TestArchitectureCommandErrorsOmitPartialReport(t *testing.T) {
	for _, tc := range []struct {
		name         string
		args         []string
		parseFailure bool
		code         int
		errorCode    string
	}{
		{name: "module", args: []string{"--module", "Missing"}, code: output.ExitConfig, errorCode: "architecture_scope_invalid"},
		{name: "path", args: []string{"--path", "does-not-exist"}, code: output.ExitConfig, errorCode: "architecture_scope_invalid"},
		{name: "parse", parseFailure: true, code: output.ExitValidation, errorCode: "architecture_parse_failed"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			root := architectureCLIProject(t, 1)
			if tc.parseFailure {
				if err := os.WriteFile(filepath.Join(root, "src/modules/Broken.bas"), []byte("Public Sub Broken(\n"), 0o644); err != nil {
					t.Fatal(err)
				}
			}
			args := append([]string{"architecture", "--json"}, tc.args...)
			text, code := runArchitectureCLI(t, root, args...)
			var payload map[string]any
			if err := json.Unmarshal([]byte(text), &payload); err != nil {
				t.Fatalf("JSON: %v: %s", err, text)
			}
			if code != tc.code || payload["architecture"] != nil || payload["error"].(map[string]any)["code"] != tc.errorCode {
				t.Fatalf("exit=%d payload=%#v", code, payload)
			}
		})
	}
}

func TestArchitectureCommandValidEmptyFilterIntersection(t *testing.T) {
	root := architectureCLIProject(t, 2)
	text, code := runArchitectureCLI(t, root, "architecture", "--module", "Main", "--path", "src/modules/Worker01.bas", "--json")
	var payload struct {
		Architecture struct {
			Summary struct {
				Procedures int `json:"procedures"`
			} `json:"summary"`
			ProjectSummary struct {
				Procedures int `json:"procedures"`
			} `json:"project_summary"`
		} `json:"architecture"`
	}
	if err := json.Unmarshal([]byte(text), &payload); err != nil {
		t.Fatalf("JSON: %v: %s", err, text)
	}
	if code != output.ExitSuccess || payload.Architecture.Summary.Procedures != 0 || payload.Architecture.ProjectSummary.Procedures != 4 {
		t.Fatalf("empty intersection: exit=%d: %s", code, text)
	}
}

func TestArchitectureHumanOutputIsBounded(t *testing.T) {
	root := architectureCLIProject(t, 24)
	text, code := runArchitectureCLI(t, root, "architecture")
	for _, want := range []string{"Project Architecture", "Entry points", "Dependencies", "Top hotspots", "Unreachable internal procedures", "Uncertainty", "use --json"} {
		if !strings.Contains(text, want) {
			t.Errorf("missing %q: %s", want, text)
		}
	}
	if code != 0 || len(strings.Split(text, "\n")) > 150 {
		t.Fatalf("unbounded/failed human output: exit=%d: %s", code, text)
	}
}

func TestArchitectureHotspotsMatchMetricsCommandCollector(t *testing.T) {
	root := architectureCLIProject(t, 2)
	cfg := config.Default()
	metrics, _, err := collectProcedureMetrics(t.Context(), root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	want, _, err := hotspots.BuildFromMetrics(t.Context(), metrics)
	if err != nil {
		t.Fatal(err)
	}
	report, _, err := architecture.CollectContext(t.Context(), root, cfg, architecture.Options{})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(report.Hotspots, want) {
		t.Fatalf("architecture changed native metrics hotspot projection:\ngot=%+v\nwant=%+v", report.Hotspots, want)
	}
}

func runArchitectureCLI(t *testing.T, root string, args ...string) (string, int) {
	t.Helper()
	var stdout, stderr strings.Builder
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs(args)
	returnErr := cmd.ExecuteContext(t.Context())
	return stdout.String(), output.ExitCode(returnErr)
}

func architectureCLIProject(t *testing.T, modules int) string {
	t.Helper()
	root := t.TempDir()
	if err := os.MkdirAll(filepath.Join(root, "src/modules"), 0o755); err != nil {
		t.Fatal(err)
	}
	configText := "[project]\nentry = \"Main.Run\"\n[excel]\npath = \"absent.xlsm\"\n[src]\nmodules = \"src/modules\"\n"
	if err := os.WriteFile(filepath.Join(root, config.FileName), []byte(configText), 0o644); err != nil {
		t.Fatal(err)
	}
	for i := range modules {
		name := fmt.Sprintf("Worker%02d", i)
		body := "Public Sub Execute()\nEnd Sub\nPrivate Sub Unused()\nEnd Sub\n"
		if i == 0 {
			name = "Main"
			body = "Public Sub Run()\n    If True Then\n        Worker01.Execute\n    End If\nEnd Sub\nPrivate Sub Unused()\nEnd Sub\n"
		}
		if err := os.WriteFile(filepath.Join(root, "src/modules", name+".bas"), []byte("Attribute VB_Name = \""+name+"\"\n"+body), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	return root
}
