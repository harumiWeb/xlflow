package project

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
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
		"'@Tag(\"smoke\")",
		"Public Sub BeforeAll()",
		"Public Sub AfterAll()",
		"Public Sub BeforeEach()",
		"Public Sub AfterEach()",
		"Public Sub Test_Sample_Pass()",
		"Public Sub Test_Sample_Inconclusive()",
		"XlflowAssert.AssertInconclusive",
	} {
		if !strings.Contains(content, want) {
			t.Fatalf("sample test module missing %q:\n%s", want, content)
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
		"custom/modules/XlflowAssert.bas",
		"custom/modules/XlflowRuntime.bas",
		"custom/modules/XlflowUI.bas",
		"custom/modules/XlflowDebug.bas",
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
	if len(result.Created) != 4 {
		t.Fatalf("created count = %d, want 4", len(result.Created))
	}
}

func TestInstallHelperModulesRefusesOverwrite(t *testing.T) {
	dir := t.TempDir()
	moduleDir := filepath.Join(dir, "src", "modules")
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
	if !strings.Contains(got, "Minimal assertion helpers for workbook-side tests") {
		t.Fatalf("Assert helper should explain its intended use:\n%s", got)
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
		"Public Sub AssertTrue(ByVal condition As Boolean",
		"Public Sub AssertFalse(ByVal condition As Boolean",
		"Public Sub AssertNotEqual(ByVal expected As Variant, ByVal actual As Variant",
		"Public Sub AssertFail(Optional ByVal message As String = \"\")",
		"Public Sub AssertInconclusive(Optional ByVal message As String = \"\")",
		"Public Sub AssertIsNothing(ByVal value As Variant",
		"Public Sub AssertIsNotNothing(ByVal value As Variant",
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
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "XlflowRuntime.bas"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "XlflowRuntime"`,
		"XlflowRuntime exposes the execution mode",
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
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "XlflowUI.bas"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "XlflowUI"`,
		"XlflowUI keeps one workbook-side dialog API",
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
	body, err := os.ReadFile(filepath.Join(dir, "src", "modules", "XlflowDebug.bas"))
	if err != nil {
		t.Fatal(err)
	}
	got := string(body)
	for _, want := range []string{
		`Attribute VB_Name = "XlflowDebug"`,
		"XlflowDebug mirrors workbook-side debug output to the terminal during xlflow runs",
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
	runtimePath := filepath.Join(dir, "src", "modules", "XlflowRuntime.bas")
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
	uiPath := filepath.Join(dir, "src", "modules", "XlflowUI.bas")
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
	debugPath := filepath.Join(dir, "src", "modules", "XlflowDebug.bas")
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
	if _, err := os.Stat(filepath.Join(dir, "src", "modules", "XlflowRuntime.bas")); !os.IsNotExist(err) {
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
