package config

import (
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

func writePreflightConfig(t *testing.T, body string) string {
	t.Helper()
	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, FileName), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	return dir
}

func TestLoadNormalizesPreflightAllowedDiagnostics(t *testing.T) {
	dir := writePreflightConfig(t, `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[preflight]
allowed_diagnostics = [" vb052 ", "VBA228", "", "VB052"]
`)
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"VB052", "VBA228"}
	if !reflect.DeepEqual(cfg.Preflight.AllowedDiagnostics, want) {
		t.Fatalf("allowed diagnostics = %#v, want %#v", cfg.Preflight.AllowedDiagnostics, want)
	}
}

func TestLoadPreflightDefaultsToEmpty(t *testing.T) {
	dir := writePreflightConfig(t, `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
`)
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(cfg.Preflight.AllowedDiagnostics) != 0 {
		t.Fatalf("allowed diagnostics = %#v", cfg.Preflight.AllowedDiagnostics)
	}
}

func TestEveryRegistryPreflightBlockerCanBeAllowed(t *testing.T) {
	for _, rule := range staticrules.All() {
		if !rule.PreflightBlocking {
			continue
		}
		cfg := PreflightConfig{AllowedDiagnostics: []string{rule.ID}}
		if err := normalizePreflightConfig(&cfg); err != nil {
			t.Errorf("%s: %v", rule.ID, err)
		}
	}
}

func TestLoadRejectsInvalidPreflightAllowedDiagnostics(t *testing.T) {
	for name, id := range map[string]string{
		"unknown":                "VB999",
		"non-registry-integrity": "FRM202",
		"non-blocking":           "VB001",
	} {
		t.Run(name, func(t *testing.T) {
			dir := writePreflightConfig(t, `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[preflight]
allowed_diagnostics = ["`+id+`"]
`)
			_, err := Load(dir)
			if err == nil || !strings.Contains(err.Error(), "[preflight].allowed_diagnostics") {
				t.Fatalf("error = %v", err)
			}
		})
	}
}

func TestLoadRejectsUnknownPreflightKey(t *testing.T) {
	dir := writePreflightConfig(t, `[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[preflight]
allow_all = true
`)
	_, err := Load(dir)
	if err == nil || !strings.Contains(err.Error(), "unknown preflight configuration key: preflight.allow_all") {
		t.Fatalf("error = %v", err)
	}
}

func TestWritePreflightConfigRoundTrip(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, FileName)
	cfg := Default()
	cfg.Preflight.AllowedDiagnostics = []string{" vba228 ", "VB052", "VBA228"}
	if err := Write(path, cfg); err != nil {
		t.Fatal(err)
	}
	loaded, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"VBA228", "VB052"}
	if !reflect.DeepEqual(loaded.Preflight.AllowedDiagnostics, want) {
		t.Fatalf("allowed diagnostics = %#v, want %#v", loaded.Preflight.AllowedDiagnostics, want)
	}
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(body), "[preflight]") || !strings.Contains(string(body), "allowed_diagnostics = [") {
		t.Fatalf("generated config missing preflight section:\n%s", body)
	}
}
