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
		"Run `xlflow doctor --json` for setup phases",
		"xlflow session start",
		"xlflow pull --session --json",
		"xlflow save --json",
		"xlflow push --fast --session --no-save --json",
		"When the macro argument is omitted, `xlflow run` uses `project.entry` from `xlflow.toml`.",
		"Matching sessions are auto-reused for `list forms`, `inspect form`, `form snapshot`, `form build`, `form export-image`, `pull`, `push`, `macros`, `run`, `export-image`, `test`, `save`, `ui button add`, `ui button list`, and `ui button remove`",
		"Use `xlflow list forms --session --json` when you need workbook UserForm names",
		"Use `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --json`",
		"Use `xlflow form build <spec> --session --json`",
		"Use `xlflow form build <spec> --session --overwrite --json`",
		"[references/formulas.md](references/formulas.md)",
		"xlflow formulas pull --json",
		"[references/forms.md](references/forms.md)",
		"[references/xlflow-ui.md](references/xlflow-ui.md)",
		"XlflowUI.MsgBox",
		"XlflowUI.GetOpenFilename",
		"DefaultResponse",
		"DefaultValue",
		"--msgbox",
		"--inputbox",
		"--filedialog",
		"--ui-stream",
		"xlflow: ui kind=msgbox id=confirm-save source=default result=yes",
		"xlflow: ui kind=file-open id=source-files source=scripted value=C:\\temp\\a.txt | C:\\temp\\b.txt",
		"ui.events",
		"UI section in human output",
		"Use `xlflow form export-image <FormName> --out <path> --session --json`",
		"Treat it as secondary visual confirmation because the capture path is experimental",
		"Use `xlflow export-image` when verification depends on rendered appearance",
		"--gui-compile-errors",
		"[lint].forbid_interactive_input = false",
		"xlflow session stop",
		"Interactive terminals show a spinner, non-interactive or --json runs fall back to a single stderr progress line",
		"`ui button add`",
		"`ui button list`",
		"`ui button remove`",
		"process cleanup --all",
		"process list",
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
	if _, err := os.Stat(filepath.Join(dir, "skills", "xlflow", "references", "forms.md")); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(dir, "skills", "xlflow", "references", "formulas.md")); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(dir, "skills", "xlflow", "references", "xlflow-ui.md")); err != nil {
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
