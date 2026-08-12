package cli

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
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
