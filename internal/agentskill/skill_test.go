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
		"Use a listed `qualified_name` from `xlflow macros --json`",
		"Run `xlflow doctor --json` for setup phases",
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
