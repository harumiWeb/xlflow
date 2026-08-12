package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/output"
)

func writePreflightWaiverModule(t *testing.T, root, relativePath, source string) {
	t.Helper()
	path := filepath.Join(root, relativePath)
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
}

func warningForRule(t *testing.T, warnings []map[string]any, rule string) map[string]any {
	t.Helper()
	for _, warning := range warnings {
		if warning["code"] == "preflight_diagnostic_allowed" && warning["rule"] == rule {
			return warning
		}
	}
	t.Fatalf("warning for %s not found: %#v", rule, warnings)
	return nil
}

func TestRunSourcePreflightAllowsAndAggregatesLintBlocker(t *testing.T) {
	dir := t.TempDir()
	writePreflightWaiverModule(t, dir, filepath.Join("src", "modules", "Main.bas"), `Attribute VB_Name = "Main"
Option Explicit
Private Const Scalar As Long = 1
Public Sub Run()
    Scalar()
    Scalar()
End Sub
`)
	cfg := config.Default()
	cfg.Preflight.AllowedDiagnostics = []string{"VB052"}
	a := &app{cwd: dir, stdout: &bytes.Buffer{}, stderr: &bytes.Buffer{}}
	if err := a.runSourcePreflight(context.Background(), "push", cfg, "pushing to Excel", nil, nil); err != nil {
		t.Fatalf("preflight error = %v", err)
	}
	// A repeated preflight in run --push must not double-count the same source
	// locations.
	if err := a.runSourcePreflight(context.Background(), "run", cfg, "running macros", nil, nil); err != nil {
		t.Fatalf("second preflight error = %v", err)
	}
	warning := warningForRule(t, a.preflightWaiverWarnings(), "VB052")
	if warning["count"] != 2 {
		t.Fatalf("warning = %#v, want count 2; waivers = %#v", warning, a.preflightWaivers)
	}
}

func TestRunSourcePreflightAllowsAnalyzeBlocker(t *testing.T) {
	dir := t.TempDir()
	writePreflightWaiverModule(t, dir, filepath.Join("src", "modules", "Main.bas"), `Attribute VB_Name = "Main"
Option Explicit
Public Sub MutateLong(ByRef value As Long)
End Sub
Public Sub Run()
    Dim value As String
    MutateLong value
End Sub
`)
	cfg := config.Default()
	cfg.Preflight.AllowedDiagnostics = []string{"VBA228"}
	a := &app{cwd: dir, stdout: &bytes.Buffer{}, stderr: &bytes.Buffer{}}
	if err := a.runSourcePreflight(context.Background(), "run", cfg, "running macros", nil, nil); err != nil {
		t.Fatalf("preflight error = %v", err)
	}
	warning := warningForRule(t, a.preflightWaiverWarnings(), "VBA228")
	if warning["count"] != 1 {
		t.Fatalf("warning = %#v, want count 1", warning)
	}
}

func TestRunSourcePreflightStillBlocksOtherDiagnosticsAndReportsWaiver(t *testing.T) {
	dir := t.TempDir()
	writePreflightWaiverModule(t, dir, filepath.Join("src", "classes", "Widget.cls"), `Attribute VB_Name = "Widget"
Option Explicit
Private Const Scalar As Long = 1
Public Sub Run()
    Scalar()
    RaiseEvent Missing
End Sub
`)
	cfg := config.Default()
	cfg.Preflight.AllowedDiagnostics = []string{"VB052"}
	var stdout bytes.Buffer
	a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}, json: true}
	err := a.runSourcePreflight(context.Background(), "push", cfg, "pushing to Excel", nil, nil)
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("error = %v, exit = %d", err, output.ExitCode(err))
	}
	var envelope struct {
		Issues   []map[string]any `json:"issues"`
		Warnings []map[string]any `json:"warnings"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &envelope); err != nil {
		t.Fatal(err)
	}
	if len(envelope.Issues) != 1 || envelope.Issues[0]["code"] != "VB054" {
		t.Fatalf("issues = %#v", envelope.Issues)
	}
	if warning := warningForRule(t, envelope.Warnings, "VB052"); warning["count"] != float64(1) {
		t.Fatalf("warning = %#v", warning)
	}
}

func TestPreflightWaiverWarningAttachesToLaterFailure(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{stdout: &stdout, stderr: &bytes.Buffer{}, json: true}
	a.recordPreflightWaiver(preflightWaiver{Rule: "VB052", Family: "lint", File: "Main.bas", Line: 4})
	env := output.Failure("push", output.Error{Code: "vba_compile_failed", Message: "compile failed", Phase: "compile_vba"})
	err := a.write(env, output.ExitValidation)
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("error = %v, exit = %d", err, output.ExitCode(err))
	}
	var got struct {
		Warnings []map[string]any `json:"warnings"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	warning := warningForRule(t, got.Warnings, "VB052")
	if warning["count"] != float64(1) {
		t.Fatalf("warning = %#v", warning)
	}
}

func TestPreflightWaiverWarningAttachesToProcessingFailure(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{stdout: &stdout, stderr: &bytes.Buffer{}, json: true}
	a.recordPreflightWaiver(preflightWaiver{Rule: "VB052", Family: "lint", File: "Main.bas", Line: 4})
	err := a.writeFailure("push", output.ExitEnvironment, "analyze_failed", context.Canceled)
	if err == nil || output.ExitCode(err) != output.ExitEnvironment {
		t.Fatalf("error = %v", err)
	}
	var got struct {
		Warnings []map[string]any `json:"warnings"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatal(err)
	}
	if warning := warningForRule(t, got.Warnings, "VB052"); warning["count"] != float64(1) {
		t.Fatalf("warning = %#v", warning)
	}
}

func TestPreflightWaiverWarningRendersInHumanOutput(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{stdout: &stdout, stderr: &bytes.Buffer{}}
	a.recordPreflightWaiver(preflightWaiver{Rule: "VB052", Family: "lint", File: "Main.bas", Line: 4})
	env := output.Failure("push", output.Error{Code: "vba_compile_failed", Message: "compile failed", Phase: "compile_vba"})
	if err := a.write(env, output.ExitValidation); err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("error = %v", err)
	}
	for _, want := range []string{
		"Warnings:",
		"[preflight_diagnostic_allowed]",
		"VB052",
		"Excel/VBE compilation may still fail.",
	} {
		if !strings.Contains(stdout.String(), want) {
			t.Fatalf("output missing %q:\n%s", want, stdout.String())
		}
	}
}

func TestPreflightWaiverWarningsAreSortedByRule(t *testing.T) {
	a := &app{}
	a.recordPreflightWaiver(preflightWaiver{Rule: "VBA228", Family: "analyze", File: "Main.bas", Line: 6})
	a.recordPreflightWaiver(preflightWaiver{Rule: "VB052", Family: "lint", File: "Main.bas", Line: 4})
	warnings := a.preflightWaiverWarnings()
	if len(warnings) != 2 || warnings[0]["rule"] != "VB052" || warnings[1]["rule"] != "VBA228" {
		t.Fatalf("warnings = %#v", warnings)
	}
}

func TestConfiguredButUndetectedWaiverIsQuiet(t *testing.T) {
	dir := t.TempDir()
	writePreflightWaiverModule(t, dir, filepath.Join("src", "modules", "Main.bas"), `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
End Sub
`)
	cfg := config.Default()
	cfg.Preflight.AllowedDiagnostics = []string{"VB052"}
	a := &app{cwd: dir, stdout: &bytes.Buffer{}, stderr: &bytes.Buffer{}}
	if err := a.runSourcePreflight(context.Background(), "push", cfg, "pushing to Excel", nil, nil); err != nil {
		t.Fatal(err)
	}
	if warnings := a.preflightWaiverWarnings(); len(warnings) != 0 {
		t.Fatalf("warnings = %#v", warnings)
	}
}

func TestPreflightWaiverDoesNotSuppressBatchDiagnostics(t *testing.T) {
	dir := t.TempDir()
	writePreflightWaiverModule(t, dir, filepath.Join("src", "modules", "Main.bas"), `Attribute VB_Name = "Main"
Option Explicit
Private Const Scalar As Long = 1
Public Sub MutateLong(ByRef value As Long)
End Sub
Public Sub Run()
    Dim value As String
    Scalar()
    MutateLong value
End Sub
`)
	cfg := config.Default()
	cfg.Preflight.AllowedDiagnostics = []string{"VB052", "VBA228"}
	lintResult, err := (lint.Linter{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	foundLint := false
	for _, issue := range lintResult.Issues {
		if issue.Code == "VB052" && issue.Severity == "error" {
			foundLint = true
		}
	}
	if !foundLint {
		t.Fatalf("VB052 error missing: %#v", lintResult.Issues)
	}
	analyzeResult, err := (analyze.Analyzer{RootDir: dir, Config: cfg}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	foundAnalyze := false
	for _, finding := range analyzeResult.Findings {
		if finding.Code == "VBA228" && finding.Severity == "error" {
			foundAnalyze = true
		}
	}
	if !foundAnalyze {
		t.Fatalf("VBA228 error missing: %#v", analyzeResult.Findings)
	}
}

func TestIgnoredAnalysisFindingDoesNotCreateWaiverWarning(t *testing.T) {
	dir := t.TempDir()
	writePreflightWaiverModule(t, dir, filepath.Join("src", "modules", "Main.bas"), `Attribute VB_Name = "Main"
Option Explicit
Public Sub MutateLong(ByRef value As Long)
End Sub
Public Sub Run()
    Dim value As String
    MutateLong value
End Sub
`)
	cfg := config.Default()
	cfg.Preflight.AllowedDiagnostics = []string{"VBA228"}
	a := &app{cwd: dir, stdout: &bytes.Buffer{}, stderr: &bytes.Buffer{}}
	if err := a.runSourcePreflight(context.Background(), "run", cfg, "running macros", map[string]bool{"VBA228": true}, nil); err != nil {
		t.Fatal(err)
	}
	if warnings := a.preflightWaiverWarnings(); len(warnings) != 0 {
		t.Fatalf("warnings = %#v", warnings)
	}
}
