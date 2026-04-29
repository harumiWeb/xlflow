package project

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestInitScaffold(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Init(dir, workbook)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/Input.xlsm" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
	for _, path := range []string{
		config.FileName,
		"prompts/agent.md",
		"src/modules",
		"src/classes",
		"src/forms",
		"src/workbook",
		"tests",
		".xlflow",
	} {
		if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(path))); err != nil {
			t.Fatalf("expected %s: %v", path, err)
		}
	}
}

func TestInitRefusesOverwrite(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := Init(dir, workbook); err != nil {
		t.Fatal(err)
	}
	if _, err := Init(dir, workbook); err == nil {
		t.Fatal("expected overwrite refusal")
	}
}
