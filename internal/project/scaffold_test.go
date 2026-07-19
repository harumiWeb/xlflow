package project

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/testdiscover"
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
		"src/modules/Xlflow/XlflowAssert.bas",
		"src/modules/Tests/SampleTests.bas",
		"src/modules",
		"src/classes",
		"src/forms",
		"src/workbook",
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

func TestInitScaffoldAcceptsSidecarCodeSource(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := InitWithOptions(dir, workbook, InitOptions{UserFormCodeSource: "sidecar"}); err != nil {
		t.Fatal(err)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.UserForm.CodeSource != "sidecar" {
		t.Fatalf("init userform code source = %q, want sidecar", cfg.UserForm.CodeSource)
	}
}

func TestInitScaffoldRejectsInvalidCodeSource(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "Input.xlsm")
	if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := InitWithOptions(dir, workbook, InitOptions{UserFormCodeSource: "broken"}); err == nil || !strings.Contains(err.Error(), "frm, sidecar") {
		t.Fatalf("expected invalid code source error, got %v", err)
	}
}

func TestInitScaffoldPreservesXlamWorkbook(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "ExistingAddin.xlam")
	wantBody := []byte("fake add-in workbook")
	if err := os.WriteFile(workbook, wantBody, 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Init(dir, workbook)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/ExistingAddin.xlam" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
	gotBody, err := os.ReadFile(filepath.Join(dir, "build", "ExistingAddin.xlam"))
	if err != nil {
		t.Fatal(err)
	}
	if string(gotBody) != string(wantBody) {
		t.Fatalf("copied workbook body = %q, want %q", string(gotBody), string(wantBody))
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Excel.Path != "build/ExistingAddin.xlam" {
		t.Fatalf("excel path = %q", cfg.Excel.Path)
	}
	for _, path := range []string{
		"src/modules",
		"src/classes",
		"src/forms",
		"src/workbook",
	} {
		if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(path))); err != nil {
			t.Fatalf("expected %s: %v", path, err)
		}
	}
}

func TestInitScaffoldPreservesXlsbWorkbook(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "ExistingModel.xlsb")
	wantBody := []byte("fake binary workbook")
	if err := os.WriteFile(workbook, wantBody, 0o644); err != nil {
		t.Fatal(err)
	}
	result, err := Init(dir, workbook)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/ExistingModel.xlsb" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
	gotBody, err := os.ReadFile(filepath.Join(dir, "build", "ExistingModel.xlsb"))
	if err != nil {
		t.Fatal(err)
	}
	if string(gotBody) != string(wantBody) {
		t.Fatalf("copied workbook body = %q, want %q", string(gotBody), string(wantBody))
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Excel.Path != "build/ExistingModel.xlsb" {
		t.Fatalf("excel path = %q", cfg.Excel.Path)
	}
}

func TestInitRejectsUnsupportedWorkbookExtensions(t *testing.T) {
	for _, name := range []string{"Existing.xlsx", "Existing.txt"} {
		t.Run(name, func(t *testing.T) {
			dir := t.TempDir()
			workbook := filepath.Join(dir, name)
			if err := os.WriteFile(workbook, []byte("fake workbook"), 0o644); err != nil {
				t.Fatal(err)
			}
			_, err := Init(dir, workbook)
			if err == nil {
				t.Fatal("expected extension validation error")
			}
			if !strings.Contains(err.Error(), ".xlsm, .xlam, or .xlsb") {
				t.Fatalf("unexpected error: %v", err)
			}
			if _, statErr := os.Stat(filepath.Join(dir, config.FileName)); !os.IsNotExist(statErr) {
				t.Fatalf("expected config not to be created, got %v", statErr)
			}
		})
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

func TestNewScaffoldCreatesSampleTests(t *testing.T) {
	dir := t.TempDir()
	_, err := New(dir, "Book", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "src", "modules", "Tests", "SampleTests.bas")
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("expected sample test module: %v", err)
	}
	content := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "SampleTests"`,
		"Option Explicit",
		"xlflow tests are public parameterless Sub procedures",
		"xlflow test --rerun-failed 1",
		"'@Tag(\"smoke\")",
		"Public Sub Test_Sample_Pass()",
		"'@TestCase(\"adds positives\"; 1, 2, 3)",
		"Public Sub Test_Adds_Numbers(ByVal leftValue As Long, ByVal rightValue As Long, ByVal expected As Long)",
		"'@ExpectedError(5)",
		"Public Sub Test_Expected_Error()",
		"'@Todo(\"not implemented yet\")",
		"Public Sub Test_Sample_Todo()",
	} {
		if !strings.Contains(content, want) {
			t.Fatalf("sample test module missing %q:\n%s", want, content)
		}
	}

	discovered, err := testdiscover.Discover(testdiscover.Options{RootDir: dir, Config: config.Default()})
	if err != nil {
		t.Fatalf("sample test module should be discoverable: %v", err)
	}
	if discovered.Summary.Tests != 5 {
		t.Fatalf("sample test count = %d, want 5: %+v", discovered.Summary.Tests, discovered.Items)
	}
	byID := map[string]testdiscover.Test{}
	for _, item := range discovered.Items {
		byID[item.ID] = item
	}
	if pass := byID["SampleTests.Test_Sample_Pass"]; len(pass.Tags) != 1 || pass.Tags[0] != "smoke" {
		t.Fatalf("sample pass tags = %v, want [smoke]", pass.Tags)
	}
	if positive := byID["SampleTests.Test_Adds_Numbers[adds positives]"]; len(positive.Arguments) != 3 {
		t.Fatalf("parameterized sample arguments = %v, want 3 args", positive.Arguments)
	}
	if expected := byID["SampleTests.Test_Expected_Error"]; expected.ExpectedError == nil || expected.ExpectedError.Number != 5 {
		t.Fatalf("expected-error sample metadata = %+v, want error 5", expected.ExpectedError)
	}
	if todo := byID["SampleTests.Test_Sample_Todo"]; todo.StatusHint != "todo" || todo.Todo == nil {
		t.Fatalf("todo sample metadata = status %q todo %+v, want todo", todo.StatusHint, todo.Todo)
	}
}

func TestNewScaffoldCreatesAppModuleWithDocComment(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "App.bas"))
	if err != nil {
		t.Fatalf("expected App module: %v", err)
	}
	content := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "App"`,
		"''' Runs the workbook automation entry point.",
		"''' Args:",
		"'''     wb: Workbook that xlflow passes from Main.Run.",
		"Public Sub RunCore(ByVal wb As Workbook)",
	} {
		if !strings.Contains(content, want) {
			t.Fatalf("App module missing %q:\n%s", want, content)
		}
	}
}

func TestInstallHelperModulesUsesConfiguredModuleRoot(t *testing.T) {
	dir := t.TempDir()
	result, err := InstallHelperModules(dir, config.SourceConfig{Modules: filepath.ToSlash(filepath.Join("custom", "modules"))})
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{
		"custom/modules/Xlflow/XlflowAssert.bas",
		"custom/modules/Xlflow/XlflowRuntime.bas",
		"custom/modules/Xlflow/XlflowUI.bas",
		"custom/modules/Xlflow/XlflowDebug.bas",
	} {
		if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(path))); err != nil {
			t.Fatalf("expected %s: %v", path, err)
		}
		if !containsString(result.Created, path) {
			t.Fatalf("expected %s in created paths: %v", path, result.Created)
		}
	}
	if _, err := os.Stat(filepath.Join(dir, "custom", "modules", "Main.bas")); !os.IsNotExist(err) {
		t.Fatalf("expected Main.bas not to be installed for helper bundle, got %v", err)
	}
	body, err := os.ReadFile(filepath.Join(dir, "custom", "modules", "Xlflow", "XlflowAssert.bas"))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(body), "Public Sub AssertStrictEquals") || !strings.Contains(string(body), "Public Sub AssertRangeEquals") {
		t.Fatalf("installed helper should include expanded assertions:\n%s", string(body))
	}
	if len(result.Created) != 4 {
		t.Fatalf("created count = %d, want 4", len(result.Created))
	}
}

func TestInstallHelperModulesRefusesOverwrite(t *testing.T) {
	dir := t.TempDir()
	moduleDir := filepath.Join(dir, "src", "modules", "Xlflow")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "XlflowUI.bas"), []byte("existing"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := InstallHelperModules(dir, config.SourceConfig{}); err == nil {
		t.Fatal("expected overwrite refusal")
	} else if !strings.Contains(err.Error(), "refusing to overwrite existing file") {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestInstallHelperModulesRefusesLegacyRootHelper(t *testing.T) {
	dir := t.TempDir()
	legacyPath := filepath.Join(dir, "src", "modules", "XlflowUI.bas")
	if err := os.MkdirAll(filepath.Dir(legacyPath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(legacyPath, []byte("existing"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := InstallHelperModules(dir, config.SourceConfig{}); err == nil {
		t.Fatal("expected legacy helper collision refusal")
	} else if !strings.Contains(err.Error(), legacyPath) {
		t.Fatalf("expected legacy helper path in error, got %v", err)
	}
}

func TestNewModuleCreatesStandardModule(t *testing.T) {
	dir := t.TempDir()
	result, err := NewModule(dir, "InvoiceProcessor", "standard", config.SourceConfig{Modules: filepath.ToSlash(filepath.Join("custom", "modules"))})
	if err != nil {
		t.Fatalf("NewModule() error = %v", err)
	}
	if result.Kind != "standard" || result.Name != "InvoiceProcessor" || result.Path != "custom/modules/InvoiceProcessor.bas" {
		t.Fatalf("unexpected result: %+v", result)
	}
	body, err := os.ReadFile(filepath.Join(dir, "custom", "modules", "InvoiceProcessor.bas"))
	if err != nil {
		t.Fatal(err)
	}
	want := "Attribute VB_Name = \"InvoiceProcessor\"\nOption Explicit\n"
	if string(body) != want {
		t.Fatalf("module body = %q, want %q", string(body), want)
	}
}

func TestNewModuleCreatesClassModule(t *testing.T) {
	dir := t.TempDir()
	result, err := NewModule(dir, "InvoiceService", "class", config.SourceConfig{Classes: filepath.ToSlash(filepath.Join("custom", "classes"))})
	if err != nil {
		t.Fatalf("NewModule() error = %v", err)
	}
	if result.Kind != "class" || result.Name != "InvoiceService" || result.Path != "custom/classes/InvoiceService.cls" {
		t.Fatalf("unexpected result: %+v", result)
	}
	body, err := os.ReadFile(filepath.Join(dir, "custom", "classes", "InvoiceService.cls"))
	if err != nil {
		t.Fatal(err)
	}
	want := "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nAttribute VB_Name = \"InvoiceService\"\nOption Explicit\n"
	if string(body) != want {
		t.Fatalf("class body = %q, want %q", string(body), want)
	}
}

func TestNewModuleRejectsInvalidTypeNameAndOverwrite(t *testing.T) {
	dir := t.TempDir()
	if _, err := NewModule(dir, "InvoiceProcessor", "form", config.SourceConfig{}); err == nil {
		t.Fatal("expected invalid type error")
	}
	for _, name := range []string{"", "../Escape", "foo/bar", "foo\\bar", "1Bad", "Bad Name"} {
		if _, err := NewModule(dir, name, "standard", config.SourceConfig{}); err == nil {
			t.Fatalf("expected invalid name error for %q", name)
		}
	}
	if _, err := NewModule(dir, "Duplicate", "standard", config.SourceConfig{}); err != nil {
		t.Fatalf("first NewModule() error = %v", err)
	}
	if _, err := NewModule(dir, "Duplicate", "standard", config.SourceConfig{}); err == nil {
		t.Fatal("expected overwrite error")
	} else if !errors.Is(err, ErrScaffoldExists) {
		t.Fatalf("expected ErrScaffoldExists, got %v", err)
	}
}

func TestNewModuleRejectsDuplicateNameAcrossRoots(t *testing.T) {
	dir := t.TempDir()
	if _, err := NewModule(dir, "Duplicate", "standard", config.SourceConfig{}); err != nil {
		t.Fatalf("first NewModule() error = %v", err)
	}
	if _, err := NewModule(dir, "Duplicate", "class", config.SourceConfig{}); err == nil {
		t.Fatal("expected duplicate component name error")
	} else if !errors.Is(err, ErrScaffoldExists) {
		t.Fatalf("expected ErrScaffoldExists, got %v", err)
	}
}

func TestNewUserFormCreatesSidecarArtifacts(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	cfg.Src.Forms = filepath.ToSlash(filepath.Join("custom", "forms"))
	result, err := NewUserForm(dir, "CustomerForm", cfg)
	if err != nil {
		t.Fatalf("NewUserForm() error = %v", err)
	}
	if result.Name != "CustomerForm" || result.CodeSource != "sidecar" {
		t.Fatalf("unexpected result: %+v", result)
	}
	if result.CodePath != "custom/forms/code/CustomerForm.bas" || result.SpecPath != "custom/forms/specs/CustomerForm.yaml" {
		t.Fatalf("unexpected paths: %+v", result)
	}
	codePath := filepath.Join(dir, "custom", "forms", "code", "CustomerForm.bas")
	codeBody, err := os.ReadFile(codePath)
	if err != nil {
		t.Fatal(err)
	}
	if string(codeBody) != "Option Explicit\n" {
		t.Fatalf("code body = %q", string(codeBody))
	}
	specPath := filepath.Join(dir, "custom", "forms", "specs", "CustomerForm.yaml")
	specBody, err := os.ReadFile(specPath)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(specBody), "warnings:") {
		t.Fatalf("new authoring spec must not include snapshot-only warnings: %s", specBody)
	}
	input, err := forms.ResolveSpecInput(dir, specPath)
	if err != nil {
		t.Fatal(err)
	}
	spec, err := forms.LoadFormSpec(input)
	if err != nil {
		t.Fatalf("generated spec should load: %v", err)
	}
	if spec.Form.Name != "CustomerForm" || spec.Kind != "xlflow.userform" || spec.Basis != "designer" {
		t.Fatalf("unexpected spec: %+v", spec)
	}
	if _, err := os.Stat(filepath.Join(dir, "custom", "forms", "CustomerForm.frm")); !os.IsNotExist(err) {
		t.Fatalf("form new should not create .frm, stat err = %v", err)
	}
}

func TestNewUserFormRejectsFrmModeInvalidNameAndOverwrite(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	cfg.UserForm.CodeSource = "frm"
	if _, err := NewUserForm(dir, "CustomerForm", cfg); err == nil {
		t.Fatal("expected frm mode error")
	}
	cfg.UserForm.CodeSource = "sidecar"
	if _, err := NewUserForm(dir, "Bad Name", cfg); err == nil {
		t.Fatal("expected invalid name error")
	}
	if _, err := NewUserForm(dir, "CustomerForm", cfg); err != nil {
		t.Fatalf("first NewUserForm() error = %v", err)
	}
	if _, err := NewUserForm(dir, "CustomerForm", cfg); err == nil {
		t.Fatal("expected overwrite error")
	} else if !errors.Is(err, ErrScaffoldExists) {
		t.Fatalf("expected ErrScaffoldExists, got %v", err)
	}
}

func TestNewUserFormRejectsDuplicateNameAcrossRoots(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if _, err := NewModule(dir, "CustomerForm", "class", cfg.Src); err != nil {
		t.Fatalf("NewModule() error = %v", err)
	}
	if _, err := NewUserForm(dir, "CustomerForm", cfg); err == nil {
		t.Fatal("expected duplicate component name error")
	} else if !errors.Is(err, ErrScaffoldExists) {
		t.Fatalf("expected ErrScaffoldExists, got %v", err)
	}
}

func TestNewModuleRejectsExistingSidecarUserFormName(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	if _, err := NewUserForm(dir, "CustomerForm", cfg); err != nil {
		t.Fatalf("NewUserForm() error = %v", err)
	}
	if _, err := NewModule(dir, "CustomerForm", "standard", cfg.Src); err == nil {
		t.Fatal("expected duplicate sidecar form name error")
	} else if !errors.Is(err, ErrScaffoldExists) {
		t.Fatalf("expected ErrScaffoldExists, got %v", err)
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
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "Xlflow", "XlflowAssert.bas"))
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
	if !strings.Contains(got, "''' Asserts that two scalar values are equal.") {
		t.Fatalf("Assert helper should explain its intended use:\n%s", got)
	}
	if !strings.Contains(got, "Private Const assertFailureNumber As Long = vbObjectError + 513") {
		t.Fatalf("AssertEquals should raise a VBA error on mismatch:\n%s", got)
	}
	for _, want := range []string{
		"IsObject(expected) Or IsObject(actual)",
		"IsArray(expected) Or IsArray(actual)",
		"IsNull(expected) Or IsNull(actual)",
		"AssertEquals supports scalar values only",
		"Private Function FormatAssertValue",
		"Private Function EscapeAssertString",
		"Private Function AssertValueTypeName",
		"Private Const assertFailureNumber As Long = vbObjectError + 513",
		"Err.Raise assertFailureNumber, source, detail",
		"FormatAssertValue = \"<Null>\"",
		"FormatAssertValue = \"<Empty>\"",
		"FormatAssertValue = \"<Nothing>\"",
		"FormatAssertValue = \"<String: \"\"\"",
		"AssertValueTypeName = \"Long\"",
		"AssertValueTypeName = \"Double\"",
		"Public Sub AssertTrue(ByVal condition As Boolean",
		"Public Sub AssertFalse(ByVal condition As Boolean",
		"Public Sub AssertNotEqual(ByVal expected As Variant, ByVal actual As Variant",
		"Public Sub AssertStrictEquals(ByVal expected As Variant, ByVal actual As Variant",
		"Public Sub AssertNull(ByVal value As Variant",
		"Public Sub AssertNotNull(ByVal value As Variant",
		"Public Sub AssertEmpty(ByVal value As Variant",
		"Public Sub AssertNotEmpty(ByVal value As Variant",
		"Public Sub AssertNear(ByVal expected As Variant, ByVal actual As Variant, ByVal tolerance As Double",
		"Public Sub AssertContains(ByVal expectedSubstring As Variant, ByVal actual As Variant",
		"Public Sub AssertStartsWith(ByVal prefix As Variant, ByVal actual As Variant",
		"Public Sub AssertEndsWith(ByVal suffix As Variant, ByVal actual As Variant",
		"Public Sub AssertMatches(ByVal pattern As Variant, ByVal actual As Variant",
		`Set regex = CreateObject("VBScript.RegExp")`,
		"regex.IgnoreCase = False",
		"regex.MultiLine = False",
		"Public Sub AssertArrayEquals(ByVal expected As Variant, ByVal actual As Variant",
		"array bounds differ on dimension",
		"array mismatch at (",
		"Private Function ArrayDimensionCount",
		"Public Sub AssertRangeEquals(ByVal expected As Variant, ByVal actualRange As Object",
		"actualValues = actualRange.Value2",
		"range mismatch at ",
		"Private Function RangeCellLabel",
		"Public Sub AssertSame(ByVal expected As Variant, ByVal actual As Variant",
		"Public Sub AssertNotSame(ByVal expected As Variant, ByVal actual As Variant",
		"Public Sub AssertFail(Optional ByVal message As String = \"\")",
		"Public Sub AssertInconclusive(Optional ByVal message As String = \"\")",
		"Public Sub AssertIsNothing(ByVal value As Variant",
		"Public Sub AssertIsNotNothing(ByVal value As Variant",
		"'''     expected: Expected scalar value.",
		"'''     condition: Boolean condition to verify.",
		"''' Marks the current test as inconclusive.",
		"Err.Raise vbObjectError + 516",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("Assert helper should contain %q:\n%s", want, got)
		}
	}
}

func TestNewScaffoldCreatesRuntimeHelper(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "Xlflow", "XlflowRuntime.bas"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "XlflowRuntime"`,
		"XlflowRuntime exposes the execution mode",
		"''' Returns the current xlflow runtime mode as a stable numeric value.",
		"''' Returns the normalized runtime mode name injected by xlflow.",
		"''' Indicates whether the workbook is running without direct human interaction.",
		"Private Const xlflowInteractive As Long = 0",
		"Public Function Mode() As Long",
		"Public Function ModeName() As String",
		"Public Function IsHeadless() As Boolean",
		`ThisWorkbook.Names("__XLFLOW_MODE__").RefersTo`,
		`Environ$("XLFLOW_MODE")`,
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("runtime helper should contain %q:\n%s", want, got)
		}
	}
}

func TestNewScaffoldCreatesUIHelper(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "Xlflow", "XlflowUI.bas"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "XlflowUI"`,
		"XlflowUI keeps one workbook-side dialog API",
		"''' Shows a message box or resolves a scripted response in headless mode.",
		"'''     Id: Stable dialog id used by xlflow --msgbox.",
		"''' Shows an input box or resolves a scripted value in headless mode.",
		"''' Opens Excel's file picker or resolves scripted file paths in headless mode.",
		"''' Opens an Office file dialog or resolves scripted file paths in headless mode.",
		"''' Opens Excel's Save As picker or resolves a scripted path in headless mode.",
		"''' Opens a folder picker or resolves a scripted folder path in headless mode.",
		"Public Function MsgBox(ByVal Id As String, ByVal Prompt As String",
		"Optional ByVal DefaultResponse As String = \"\"",
		"Public Function InputBox(ByVal Id As String, ByVal Prompt As String",
		"Optional ByVal DefaultValue As String = \"\"",
		"Public Function GetOpenFilename(ByVal Id As String",
		"Public Function FileDialogOpen(ByVal Id As String",
		"Public Function GetSaveAsFilename(ByVal Id As String",
		"Public Function FolderPicker(ByVal Id As String",
		"ValidateDialogId Id, \"XlflowUI.MsgBox\"",
		"ValidateDialogId Id, \"XlflowUI.InputBox\"",
		"ValidateDialogId Id, \"XlflowUI.GetOpenFilename\"",
		"ValidateDialogId Id, \"XlflowUI.FileDialogOpen\"",
		"ValidateDialogId Id, \"XlflowUI.GetSaveAsFilename\"",
		"ValidateDialogId Id, \"XlflowUI.FolderPicker\"",
		"Private Const xlflowErrInvalidDialogId As Long",
		"Private Const xlflowErrMissingFileDialogResponse As Long",
		"Private Const xlflowFileDialogCancelToken As String = \"@cancel\"",
		"Private Sub ValidateDialogId(ByVal Id As String, ByVal SourceName As String)",
		"If XlflowRuntime.IsHeadless() Then",
		"EmitHeadlessUIEvent \"msgbox\"",
		"EmitHeadlessUIEvent \"inputbox\"",
		"EmitHeadlessUIEvent \"get-open\"",
		"EmitHeadlessUIEvent \"file-open\"",
		"EmitHeadlessUIEvent \"save-as\"",
		"EmitHeadlessUIEvent \"folder\"",
		"ResolveFileDialogResponse(\"get-open\"",
		"ResolveFileDialogResponse(\"file-open\"",
		"ResolveFileDialogResponse(\"save-as\"",
		"ResolveFileDialogResponse(\"folder\"",
		"If Len(DefaultValue) = 0 Then",
		"Err.Raise xlflowErrMissingInputResponse, \"XlflowUI.InputBox\", \"Missing scripted InputBox response",
		"BuildFileDialogResponseName(ByVal Kind As String, ByVal Id As String)",
		"SelectedItemsToVariantArray(ByVal SelectedItems As Office.FileDialogSelectedItems)",
		"Private Function ResolveStreamHelperMacro() As String",
		"Private Function JsonEscape(ByVal value As String) As String",
		`ThisWorkbook.Names(BuildResponseName(Kind, Id)).RefersTo`,
		`ThisWorkbook.Names(BuildFileDialogResponseName(Kind, Id)).RefersTo`,
		"__XLFLOW_UI_",
		"VBA.Interaction.MsgBox",
		"VBA.Interaction.InputBox",
		"Application.GetOpenFilename",
		"Application.GetSaveAsFilename",
		"Application.FileDialog(msoFileDialogOpen)",
		"Application.FileDialog(msoFileDialogFolderPicker)",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("UI helper should contain %q:\n%s", want, got)
		}
	}
}

func TestNewScaffoldCreatesDebugHelper(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "Xlflow", "XlflowDebug.bas"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "XlflowDebug"`,
		"XlflowDebug mirrors workbook-side debug output to the terminal during xlflow runs",
		"''' Writes workbook-side debug output to Debug.Print and xlflow run/test output when available.",
		"'''     Parts: Values to stringify and join with spaces.",
		"Public Sub Log(ParamArray Parts() As Variant)",
		"Debug.Print message",
		"Debug.Print",
		"EmitDebugEvent message",
		`Private Const xlflowDebugPipeName As String = "__XLFLOW_DEBUG_PIPE__"`,
		"JsonProperty(\"event\", \"debug_log\")",
		"JsonProperty(\"source\", \"XlflowDebug.Log\")",
		"JsonProperty(\"runtime_mode\", XlflowRuntime.ModeName())",
		"ResolveDebugPipeName()",
		`ThisWorkbook.Names(Name).RefersTo`,
		"CreateFileW",
		"WriteFile",
		"CloseHandle",
		"StringifyValue(ByVal Value As Variant)",
		"[Object ",
		"[Array]",
		"[Empty]",
		"[Null]",
		"[Unsupported Variant]",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("debug helper should contain %q:\n%s", want, got)
		}
	}
}

func TestNewScaffoldRuntimeHelperLintsCleanly(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	runtimePath := filepath.Join(dir, "src", "modules", "Xlflow", "XlflowRuntime.bas")
	issues, err := lint.Linter{
		RootDir: dir,
		Config:  config.Default(),
		PathFilter: func(path string) bool {
			return filepath.Clean(path) == filepath.Clean(runtimePath)
		},
	}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("runtime helper should lint cleanly: %+v", issues)
	}
}

func TestNewScaffoldUIHelperLintsCleanly(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	uiPath := filepath.Join(dir, "src", "modules", "Xlflow", "XlflowUI.bas")
	issues, err := lint.Linter{
		RootDir: dir,
		Config:  config.Default(),
		PathFilter: func(path string) bool {
			return filepath.Clean(path) == filepath.Clean(uiPath)
		},
	}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("UI helper should lint cleanly: %+v", issues)
	}
}

func TestNewScaffoldDebugHelperLintsCleanly(t *testing.T) {
	dir := t.TempDir()
	if _, err := New(dir, "Book", fakeWorkbookCreator); err != nil {
		t.Fatal(err)
	}
	debugPath := filepath.Join(dir, "src", "modules", "Xlflow", "XlflowDebug.bas")
	issues, err := lint.Linter{
		RootDir: dir,
		Config:  config.Default(),
		PathFilter: func(path string) bool {
			return filepath.Clean(path) == filepath.Clean(debugPath)
		},
	}.Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("debug helper should lint cleanly: %+v", issues)
	}
}

func TestXlflowDebugLogDoesNotForwardParamArrayToHelper(t *testing.T) {
	body := defaultDebugRuntimeModule
	if !strings.Contains(body, "Public Sub Log(ParamArray Parts() As Variant)") {
		t.Fatal("template should contain ParamArray Log entry")
	}
	if strings.Contains(body, "JoinLogMessage(Parts)") {
		t.Fatalf("Log must not forward ParamArray Parts() to another helper; build the message inline instead:\n%s", body)
	}
	if strings.Contains(body, "Private Function JoinLogMessage") {
		t.Fatalf("template should not declare JoinLogMessage; ParamArray forwarding is not portable across VBA hosts:\n%s", body)
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
	if _, err := os.Stat(filepath.Join(dir, "src", "modules", "Xlflow", "XlflowRuntime.bas")); !os.IsNotExist(err) {
		t.Fatalf("expected runtime helper not to be scaffolded for init, got %v", err)
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

func TestNewScaffoldKeepsXlamExtension(t *testing.T) {
	dir := t.TempDir()
	result, err := New(dir, "MyAddin.xlam", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/MyAddin.xlam" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Excel.Path != "build/MyAddin.xlam" {
		t.Fatalf("excel path = %q", cfg.Excel.Path)
	}
	if cfg.Project.Name != "MyAddin" {
		t.Fatalf("project name = %q", cfg.Project.Name)
	}
}

func TestNewScaffoldKeepsXlsbExtension(t *testing.T) {
	dir := t.TempDir()
	result, err := New(dir, "LargeModel.xlsb", fakeWorkbookCreator)
	if err != nil {
		t.Fatal(err)
	}
	if result.Workbook != "build/LargeModel.xlsb" {
		t.Fatalf("workbook path = %q", result.Workbook)
	}
	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Excel.Path != "build/LargeModel.xlsb" {
		t.Fatalf("excel path = %q", cfg.Excel.Path)
	}
	if cfg.Project.Name != "LargeModel" {
		t.Fatalf("project name = %q", cfg.Project.Name)
	}
}

func TestNormalizeWorkbookName(t *testing.T) {
	tests := []struct {
		name    string
		want    string
		wantErr bool
	}{
		{name: "Book", want: "Book.xlsm"},
		{name: "Book.xlsm", want: "Book.xlsm"},
		{name: "Addin.xlam", want: "Addin.xlam"},
		{name: "LargeModel.xlsb", want: "LargeModel.xlsb"},
		{name: "Book.xlsx", wantErr: true},
		{name: "Book.txt", wantErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := normalizeWorkbookName(tt.name)
			if tt.wantErr {
				if err == nil {
					t.Fatal("expected error")
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if got != tt.want {
				t.Fatalf("normalizeWorkbookName(%q) = %q, want %q", tt.name, got, tt.want)
			}
		})
	}
}

func TestNewRejectsUnsupportedWorkbookExtensions(t *testing.T) {
	for _, name := range []string{"Sales.xlsx", "Sales.txt"} {
		t.Run(name, func(t *testing.T) {
			dir := t.TempDir()
			_, err := New(dir, name, fakeWorkbookCreator)
			if err == nil {
				t.Fatal("expected extension validation error")
			}
			if !strings.Contains(err.Error(), ".xlsm, .xlam, or .xlsb") {
				t.Fatalf("unexpected error: %v", err)
			}
			if _, statErr := os.Stat(filepath.Join(dir, config.FileName)); !os.IsNotExist(statErr) {
				t.Fatalf("expected config not to be created, got %v", statErr)
			}
		})
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

func TestNewRefusesLegacyRootHelper(t *testing.T) {
	dir := t.TempDir()
	legacyPath := filepath.Join(dir, "src", "modules", "XlflowAssert.bas")
	if err := os.MkdirAll(filepath.Dir(legacyPath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(legacyPath, []byte("existing"), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := New(dir, "Sales", fakeWorkbookCreator); err == nil {
		t.Fatal("expected legacy helper collision refusal")
	} else if !strings.Contains(err.Error(), legacyPath) {
		t.Fatalf("expected legacy helper path in error, got %v", err)
	}
}

func fakeWorkbookCreator(path string) error {
	return os.WriteFile(path, []byte("fake workbook"), 0o644)
}

func TestGenerateTestModule(t *testing.T) {
	dir := t.TempDir()
	src := config.SourceConfig{Modules: "src/modules"}
	result, err := GenerateTestModule(dir, "OrderServiceTests", src)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !result.Created {
		t.Fatal("expected Created=true")
	}
	expectedPath := filepath.Join(dir, "src", "modules", "OrderServiceTests.bas")
	if result.Path != expectedPath {
		t.Fatalf("path = %q, want %q", result.Path, expectedPath)
	}

	content, err := os.ReadFile(expectedPath)
	if err != nil {
		t.Fatalf("failed to read generated file: %v", err)
	}
	body := string(content)
	for _, want := range []string{
		`Attribute VB_Name = "OrderServiceTests"`,
		"Option Explicit",
		"Public Sub BeforeAll()",
		"Public Sub AfterAll()",
		"Public Sub BeforeEach()",
		"Public Sub AfterEach()",
		"Public Sub Test_Sample()",
		"XlflowAssert.AssertTrue True",
	} {
		if !strings.Contains(body, want) {
			t.Fatalf("generated body missing %q:\n%s", want, body)
		}
	}
}

func TestGenerateTestModuleAlreadyExists(t *testing.T) {
	dir := t.TempDir()
	src := config.SourceConfig{Modules: "src/modules"}
	expectedPath := filepath.Join(dir, "src", "modules", "DuplicateTests.bas")
	if err := os.MkdirAll(filepath.Dir(expectedPath), 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(expectedPath, []byte("existing"), 0644); err != nil {
		t.Fatal(err)
	}
	_, err := GenerateTestModule(dir, "DuplicateTests", src)
	if err == nil {
		t.Fatal("expected error for existing file")
	}
}

func TestGenerateTestModuleEmptyName(t *testing.T) {
	dir := t.TempDir()
	src := config.SourceConfig{Modules: "src/modules"}
	_, err := GenerateTestModule(dir, "", src)
	if err == nil {
		t.Fatal("expected error for empty name")
	}
}

func TestGenerateTestModuleRejectsPathTraversal(t *testing.T) {
	dir := t.TempDir()
	src := config.SourceConfig{Modules: "src/modules"}
	for _, name := range []string{"../Escape", "foo/bar", "foo\\bar", "..\\Escape"} {
		_, err := GenerateTestModule(dir, name, src)
		if err == nil {
			t.Fatalf("expected error for invalid module name %q", name)
		}
	}
}

func TestGenerateTestModuleDefaultsToTestsDirWhenModulesEmpty(t *testing.T) {
	dir := t.TempDir()
	src := config.SourceConfig{Modules: ""}
	result, err := GenerateTestModule(dir, "SampleTests", src)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	expectedPath := filepath.Join(dir, "src", "modules", "Tests", "SampleTests.bas")
	if result.Path != expectedPath {
		t.Fatalf("path = %q, want %q", result.Path, expectedPath)
	}
}

func containsString(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}
