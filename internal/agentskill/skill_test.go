package agentskill

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

var referenceFiles = []string{
	"testing.md",
	"debugging.md",
	"formulas.md",
	"forms.md",
	"xlflow-ui.md",
	"recovery.md",
	"code-analysis.md",
}

func requireContains(t *testing.T, body string, wants ...string) {
	t.Helper()
	for _, want := range wants {
		if !strings.Contains(body, want) {
			t.Errorf("missing invariant %q", want)
		}
	}
}

func requireReferenceFiles(t *testing.T, skillDir string) {
	t.Helper()
	for _, name := range referenceFiles {
		path := filepath.Join(skillDir, "references", name)
		if _, err := os.Stat(path); err != nil {
			t.Errorf("missing installed reference %s: %v", name, err)
		}
	}
}

func requireBefore(t *testing.T, body, first, second string) {
	t.Helper()
	if strings.Index(body, first) >= strings.Index(body, second) {
		t.Errorf("expected %q before %q", first, second)
	}
}

func TestInstallUsesProviderDefaultTarget(t *testing.T) {
	dir := t.TempDir()
	result, err := Install(InstallOptions{RootDir: dir, Agent: "codex"})
	if err != nil {
		t.Fatal(err)
	}
	if result.Path != ".codex/skills/xlflow" {
		t.Fatalf("path = %q", result.Path)
	}
	if result.Agent != "codex" {
		t.Fatalf("agent = %q", result.Agent)
	}
	skillDir := filepath.Join(dir, ".codex", "skills", "xlflow")
	body, err := os.ReadFile(filepath.Join(skillDir, "SKILL.md"))
	if err != nil {
		t.Fatal(err)
	}
	installed := string(body)
	requireContains(t, installed,
		"name: xlflow",
		"## Core Invariants",
		"## Canonical Development Loop",
		"Analyze structural impact when relevant.",
		"Use structured JSON output for agent-facing commands.",
		"If `recovery.required=true`",
		"Do not save or blindly retry.",
		"## Quick Orientation",
		"xlflow --help",
		"xlflow macros --session --json",
		"xlflow inspect range --sheet Result --address A1:F20 --include-style --session --json",
		"xlflow export-image --sheet Sheet1 --range A1:M21 --out preview.png --session --json",
		"xlflow inspect-gui --json",
		"export-image` as the required proof",
		"## Quick Start: Normal Source Change",
		"xlflow status --json",
		"xlflow session start --json",
		"xlflow session attach --json",
		"xlflow pull --session --json",
		"xlflow lint --json",
		"xlflow analyze --json",
		"xlflow push --fast --session --no-save --json",
		"xlflow test --filter Module.TestName --session --no-save --json",
		"xlflow run --diagnostic --headless --session --json",
		"xlflow inspect workbook --session --json",
		"xlflow test --session --no-save --json",
		"xlflow save --session --json",
		"xlflow session stop --json",
		"If its structured result reports recovery required, stop here",
		"## Dispatch to Specialized References",
		"[recovery.md](references/recovery.md)",
		"[code-analysis.md](references/code-analysis.md)",
		"## Evidence of Completion",
	)
	requireBefore(t, installed, "xlflow lint --json", "xlflow push --fast --session --no-save --json")
	requireBefore(t, installed, "xlflow analyze --json", "xlflow push --fast --session --no-save --json")
	requireBefore(t, installed, "xlflow push --fast --session --no-save --json", "xlflow test --filter Module.TestName --session --no-save --json")
	for _, removed := range []string{"xlflow attach --active", "--ui-stream", "VB014", "## Progress Rules"} {
		if strings.Contains(installed, removed) {
			t.Errorf("installed orchestration skill still contains removed command detail %q", removed)
		}
	}
	requireReferenceFiles(t, skillDir)
	recovery, err := os.ReadFile(filepath.Join(skillDir, "references", "recovery.md"))
	if err != nil {
		t.Fatal(err)
	}
	requireContains(t, string(recovery),
		"Do not retry the failed operation",
		"Do not save the quarantined workbook",
		"xlflow session stop --discard --json",
		"external user-owned session",
		"only after manual Excel recovery is complete",
		"user explicitly accepts",
	)
	analysis, err := os.ReadFile(filepath.Join(skillDir, "references", "code-analysis.md"))
	if err != nil {
		t.Fatal(err)
	}
	requireContains(t, string(analysis),
		"Start with the smallest question",
		"Do not dump the full project graph",
		"An empty resolved caller list",
		"Re-run structural analysis",
	)
	ui, err := os.ReadFile(filepath.Join(skillDir, "references", "xlflow-ui.md"))
	if err != nil {
		t.Fatal(err)
	}
	requireContains(t, string(ui), "## Common Headless Failure Triage", "Application.OnKey", "xlflow inspect-gui --json")
	testing, err := os.ReadFile(filepath.Join(skillDir, "references", "testing.md"))
	if err != nil {
		t.Fatal(err)
	}
	requireContains(t, string(testing), "xlflow test --visible", "Application.OnKey", "Do not remove, bypass, or weaken production behavior")
	debugging, err := os.ReadFile(filepath.Join(skillDir, "references", "debugging.md"))
	if err != nil {
		t.Fatal(err)
	}
	requireContains(t, string(debugging), "## Common Diagnostic Triage", "existing baseline findings")
}

func TestInstallSupportsAllProviders(t *testing.T) {
	for _, provider := range Providers() {
		t.Run(provider.Name, func(t *testing.T) {
			dir := t.TempDir()
			if _, err := Install(InstallOptions{RootDir: dir, Agent: provider.Name}); err != nil {
				t.Fatal(err)
			}
			skillDir := filepath.Join(dir, filepath.FromSlash(provider.Dir), "xlflow")
			body, err := os.ReadFile(filepath.Join(skillDir, "SKILL.md"))
			if err != nil {
				t.Fatal(err)
			}
			requireContains(t, string(body), "## Canonical Development Loop", "## Dispatch to Specialized References")
			requireReferenceFiles(t, skillDir)
		})
	}
}

func TestInstallUsesExplicitTarget(t *testing.T) {
	dir := t.TempDir()
	result, err := Install(InstallOptions{RootDir: dir, Target: "skills"})
	if err != nil {
		t.Fatal(err)
	}
	if result.Path != "skills/xlflow" {
		t.Fatalf("path = %q", result.Path)
	}
	skillDir := filepath.Join(dir, "skills", "xlflow")
	if _, err := os.Stat(filepath.Join(skillDir, "SKILL.md")); err != nil {
		t.Fatal(err)
	}
	requireReferenceFiles(t, skillDir)
}

func TestInstallRefusesOverwriteUnlessForced(t *testing.T) {
	dir := t.TempDir()
	if _, err := Install(InstallOptions{RootDir: dir, Agent: "codex"}); err != nil {
		t.Fatal(err)
	}
	if _, err := Install(InstallOptions{RootDir: dir, Agent: "codex"}); err == nil {
		t.Fatal("expected overwrite refusal")
	}
	if _, err := Install(InstallOptions{RootDir: dir, Agent: "codex", Force: true}); err != nil {
		t.Fatal(err)
	}
}

func TestInstallRejectsUnknownProvider(t *testing.T) {
	_, err := Install(InstallOptions{RootDir: t.TempDir(), Agent: "unknown"})
	if err == nil || !strings.Contains(err.Error(), "unsupported skill agent") {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestInstallRejectsCopilotProvider(t *testing.T) {
	_, err := Install(InstallOptions{RootDir: t.TempDir(), Agent: "copilot"})
	if err == nil || !strings.Contains(err.Error(), "unsupported skill agent") {
		t.Fatalf("unexpected error: %v", err)
	}
}
