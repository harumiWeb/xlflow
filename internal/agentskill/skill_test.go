package agentskill

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

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
	body, err := os.ReadFile(filepath.Join(dir, ".codex", "skills", "xlflow", "SKILL.md"))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(body), "name: xlflow") {
		t.Fatalf("unexpected skill body:\n%s", body)
	}
	for _, want := range []string{
		"Treat the configured source directories as authoritative",
		"Use a listed `qualified_name` from `xlflow macros --session --json`",
		"Run `xlflow doctor --keepalive --json` for setup phases",
		"xlflow session start",
		"xlflow pull --session --keepalive --json",
		"xlflow save --json",
		"xlflow push --fast --session --no-save --keepalive --json",
		"When the macro argument is omitted, `xlflow run` uses `project.entry` from `xlflow.toml`.",
		"Matching sessions are auto-reused for `list forms`, `inspect form`, `form snapshot`, `pull`, `push`, `macros`, `run`, `export-image`, `test`, `trace`, and `save`",
		"Use `xlflow list forms --session --keepalive --json` when you need workbook UserForm names",
		"Use `xlflow form snapshot <FormName> --out <path> --session --keepalive --json`",
		"Use `xlflow export-image` when verification depends on rendered appearance",
		"--gui-compile-errors",
		"xlflow session stop",
		"XLFLOW_DONE status=success command=pull",
	} {
		if !strings.Contains(string(body), want) {
			t.Fatalf("installed skill is missing %q:\n%s", want, body)
		}
	}
}

func TestInstallSupportsAllProviders(t *testing.T) {
	for _, provider := range Providers() {
		t.Run(provider.Name, func(t *testing.T) {
			dir := t.TempDir()
			if _, err := Install(InstallOptions{RootDir: dir, Agent: provider.Name}); err != nil {
				t.Fatal(err)
			}
			if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(provider.Dir), "xlflow", "SKILL.md")); err != nil {
				t.Fatalf("expected skill for %s: %v", provider.Name, err)
			}
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
	if _, err := os.Stat(filepath.Join(dir, "skills", "xlflow", "SKILL.md")); err != nil {
		t.Fatal(err)
	}
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
