package config

import (
	"os"
	"path/filepath"
	"testing"
)

func TestLoadDefaultsAndConfiguredValues(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
name = "sales"
entry = "Main.Run"

[excel]
path = "build/Sales.xlsm"
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Project.Name != "sales" || cfg.Project.Entry != "Main.Run" {
		t.Fatalf("project config = %+v", cfg.Project)
	}
	if cfg.Src.Modules != "src/modules" {
		t.Fatalf("default modules dir = %q", cfg.Src.Modules)
	}
	if !cfg.Lint.RequireOptionExplicit {
		t.Fatalf("lint defaults were not applied")
	}
	if !cfg.Lint.ForbidInteractiveInput {
		t.Fatalf("interactive input lint default was not applied")
	}
}

func TestLoadAllowsDisablingInteractiveInputLint(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
forbid_interactive_input = false
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Lint.ForbidInteractiveInput {
		t.Fatal("expected forbid_interactive_input=false to be honored")
	}
	if !cfg.Lint.RequireOptionExplicit {
		t.Fatal("expected other lint defaults to remain enabled")
	}
}

func TestLoadMissingConfig(t *testing.T) {
	_, err := Load(t.TempDir())
	if err == nil {
		t.Fatal("expected missing config error")
	}
}
