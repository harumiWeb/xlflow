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
		".gitignore",
	} {
		if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(path))); err != nil {
			t.Fatalf("expected %s: %v", path, err)
		}
	}
	if _, err := os.Stat(filepath.Join(dir, "prompts", "agent.md")); !os.IsNotExist(err) {
		t.Fatalf("expected prompts/agent.md not to be scaffolded, got %v", err)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.UserForm.CodeSource != "frm" {
		t.Fatalf("init userform code source = %q, want frm", cfg.UserForm.CodeSource)
	}
}

func TestNewScaffoldCreatesGitignore(t *testing.T) {
	dir := t.TempDir()
	result, err := New(dir, "Book", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, ".gitignore"))
	if err != nil {
		t.Fatal(err)
	}
	if got := string(body); got != defaultGitignore {
		t.Fatalf(".gitignore =\n%s", got)
	}
	if !containsString(result.Created, ".gitignore") {
		t.Fatalf("expected .gitignore in created paths: %v", result.Created)
	}
}

func TestScaffoldAppendsMissingGitignoreEntries(t *testing.T) {
	dir := t.TempDir()
	existing := "# User\n.env\n"
	if err := os.WriteFile(filepath.Join(dir, ".gitignore"), []byte(existing), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := New(dir, "Book", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, ".gitignore"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	if !strings.HasPrefix(got, existing) {
		t.Fatalf("existing .gitignore content should be preserved:\n%s", got)
	}
	for _, want := range []string{"~$*.xls*", "*.tmp", ".xlflow/", "build/"} {
		if strings.Count(got, want) != 1 {
			t.Fatalf("expected one %q entry in .gitignore:\n%s", want, got)
		}
	}
	if !containsString(result.Created, ".gitignore") {
		t.Fatalf("expected .gitignore in created paths: %v", result.Created)
	}
}

func TestScaffoldDoesNotDuplicateGitignoreEntries(t *testing.T) {
	dir := t.TempDir()
	existing := "# Existing\nbuild/\n*.tmp\n.xlflow/\n~$*.xls*\n"
	if err := os.WriteFile(filepath.Join(dir, ".gitignore"), []byte(existing), 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := New(dir, "Book", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, ".gitignore"))
	if err != nil {
		t.Fatal(err)
	}
	if got := string(body); got != existing {
		t.Fatalf(".gitignore should be unchanged when entries exist:\n%s", got)
	}
	if containsString(result.Created, ".gitignore") {
		t.Fatalf("did not expect unchanged .gitignore in created paths: %v", result.Created)
	}
}

func TestExistingGitignoreDoesNotBlockOverwriteRefusal(t *testing.T) {
	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, ".gitignore"), []byte("# User\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := New(dir, "Sales", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	if _, err := New(dir, "Sales", fakeWorkbookCreator); err == nil {
		t.Fatal("expected protected file overwrite refusal")
	} else if strings.Contains(err.Error(), ".gitignore") {
		t.Fatalf("expected .gitignore not to cause overwrite refusal: %v", err)
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
	if !cfg.VBA.Folders || cfg.VBA.FolderAnnotation != "update" || !cfg.VBA.DefaultComponentFolders {
		t.Fatalf("unexpected scaffold vba config: %+v", cfg.VBA)
	}
	if !cfg.Lint.ForbidInteractiveInput {
		t.Fatal("expected scaffolded config to enable interactive input lint")
	}
	if cfg.UserForm.CodeSource != "sidecar" {
		t.Fatalf("new userform code source = %q, want sidecar", cfg.UserForm.CodeSource)
	}
	for _, path := range []string{"src/workbook/ThisWorkbook.bas", "src/workbook/Sheet1.bas"} {
		body, err := os.ReadFile(filepath.Join(dir, filepath.FromSlash(path)))
		if err != nil {
			t.Fatalf("expected %s: %v", path, err)
		}
		if string(body) != defaultDocumentModule {
			t.Fatalf("%s = %q, want %q", path, string(body), defaultDocumentModule)
		}
	}
}

func TestInitScaffoldDoesNotCreatePlaceholderWorkbookModules(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := Init(dir, workbook); err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{"src/workbook/ThisWorkbook.bas", "src/workbook/Sheet1.bas"} {
		if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(path))); !os.IsNotExist(err) {
			t.Fatalf("expected %s not to be scaffolded for init, got %v", path, err)
		}
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

func containsString(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}
