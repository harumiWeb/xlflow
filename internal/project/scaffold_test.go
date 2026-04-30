package project

import (
	"os"
	"path/filepath"
	"strings"
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
		"src/modules/XlflowAssert.bas",
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
	if _, err := os.Stat(filepath.Join(dir, "prompts", "agent.md")); !os.IsNotExist(err) {
		t.Fatalf("expected prompts/agent.md not to be scaffolded, got %v", err)
	}
}

func TestScaffoldCreatesAssertHelper(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "XlflowAssert.bas"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	if !strings.Contains(got, `Attribute VB_Name = "XlflowAssert"`) {
		t.Fatalf("Assert helper should preserve the imported module name:\n%s", got)
	}
	if !strings.Contains(got, "Public Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = \"\")") {
		t.Fatalf("AssertEquals signature missing from helper:\n%s", got)
	}
	if !strings.Contains(got, "Err.Raise vbObjectError + 513") {
		t.Fatalf("AssertEquals should raise a VBA error on mismatch:\n%s", got)
	}
	for _, want := range []string{
		"IsObject(expected) Or IsObject(actual)",
		"IsArray(expected) Or IsArray(actual)",
		"IsNull(expected) Or IsNull(actual)",
		"AssertEquals supports scalar values only",
		"Private Function DescribeAssertValue",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("Assert helper should contain %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "AssertTrue") {
		t.Fatalf("AssertTrue is out of scope for v1 helper:\n%s", got)
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

func TestNewScaffoldDefaultWorkbook(t *testing.T) {
	dir := t.TempDir()
	result, err := New(dir, "", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/Book.xlsm" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Project.Name != "Book" {
		t.Fatalf("project name = %q", cfg.Project.Name)
	}
	if cfg.Excel.Path != "build/Book.xlsm" {
		t.Fatalf("excel path = %q", cfg.Excel.Path)
	}
}

func TestNewScaffoldAppendsXlsmExtension(t *testing.T) {
	dir := t.TempDir()
	result, err := New(dir, "Sales", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/Sales.xlsm" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Project.Name != "Sales" {
		t.Fatalf("project name = %q", cfg.Project.Name)
	}
}

func TestNewScaffoldKeepsXlsmExtension(t *testing.T) {
	dir := t.TempDir()
	result, err := New(dir, "Sales.xlsm", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/Sales.xlsm" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
}

func TestNewRejectsNonXlsmExtension(t *testing.T) {
	dir := t.TempDir()
	_, err := New(dir, "Sales.xlsx", fakeWorkbookCreator)
	if err == nil {
		t.Fatal("expected extension validation error")
	}
	if !strings.Contains(err.Error(), ".xlsm") {
		t.Fatalf("unexpected error: %v", err)
	}
	if _, statErr := os.Stat(filepath.Join(dir, config.FileName)); !os.IsNotExist(statErr) {
		t.Fatalf("expected config not to be created, got %v", statErr)
	}
}

func TestNewRefusesOverwrite(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Sales", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	if _, err := New(dir, "Sales", fakeWorkbookCreator); err == nil {
		t.Fatal("expected overwrite refusal")
	}
}

func fakeWorkbookCreator(path string) error {
	return os.WriteFile(path, []byte("fake workbook"), 0o644)
}
