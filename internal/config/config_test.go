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
	if !cfg.VBA.Folders || cfg.VBA.FolderAnnotation != "update" || !cfg.VBA.DefaultComponentFolders {
		t.Fatalf("unexpected vba defaults: %+v", cfg.VBA)
	}
	if cfg.UserForm.CodeSource != "sidecar" {
		t.Fatalf("unexpected userform defaults: %+v", cfg.UserForm)
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

func TestLoadRejectsInvalidFolderAnnotation(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[vba]
folder_annotation = "broken"
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil {
		t.Fatal("expected invalid folder annotation error")
	}
}

func TestLoadRejectsInvalidUserFormCodeSource(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "broken"
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil {
		t.Fatal("expected invalid userform code source error")
	}
}

func TestWriteProducesReadableConfig(t *testing.T) {
	dir := t.TempDir()
	cfg := Default()
	cfg.Project.Name = "write-test"
	cfg.UserForm.CodeSource = "frm"
	cfg.Lint.ForbidInteractiveInput = false

	p := filepath.Join(dir, FileName)
	if err := Write(p, cfg); err != nil {
		t.Fatalf("Write failed: %v", err)
	}

	loaded, err := Load(dir)
	if err != nil {
		t.Fatalf("Load failed after Write: %v", err)
	}
	if loaded.Project.Name != "write-test" {
		t.Fatalf("name mismatch: got %q, want write-test", loaded.Project.Name)
	}
	if loaded.UserForm.CodeSource != "frm" {
		t.Fatalf("userform.code_source mismatch: got %q, want frm", loaded.UserForm.CodeSource)
	}
	if loaded.Lint.ForbidInteractiveInput {
		t.Fatal("expected forbid_interactive_input=false")
	}
	if !loaded.Lint.RequireOptionExplicit {
		t.Fatal("expected require_option_explicit=true")
	}
}
