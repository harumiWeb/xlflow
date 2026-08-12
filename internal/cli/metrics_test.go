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
