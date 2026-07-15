package project

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/workbookformat"
)

type InitResult struct {
	ConfigPath string   `json:"config_path"`
	Workbook   string   `json:"workbook"`
	Created    []string `json:"created"`
}

type InstallModulesResult struct {
	Created []string `json:"created"`
}

type NewModuleResult struct {
	Kind    string   `json:"kind"`
	Name    string   `json:"name"`
	Path    string   `json:"path"`
	Created []string `json:"created"`
}

type NewFormResult struct {
	Name       string   `json:"name"`
	CodePath   string   `json:"code_path"`
	SpecPath   string   `json:"spec_path"`
	Created    []string `json:"created"`
	CodeSource string   `json:"code_source"`
}

var (
	ErrInvalidComponentName    = errors.New("invalid component name")
	ErrInvalidModuleType       = errors.New("invalid module type")
	ErrScaffoldExists          = errors.New("scaffold target already exists")
	ErrUserFormRequiresSidecar = errors.New(`form new requires [userform].code_source = "sidecar"`)
)

type bundledModuleTemplate struct {
	fileName string
	body     string
}

type WorkbookCreator func(path string) error

type InitOptions struct {
	UserFormCodeSource string
}

func Init(cwd, workbookPath string) (InitResult, error) {
	return InitWithOptions(cwd, workbookPath, InitOptions{UserFormCodeSource: "frm"})
}

func InitWithOptions(cwd, workbookPath string, opts InitOptions) (InitResult, error) {
	var result InitResult
	if workbookPath == "" {
		return result, errors.New("workbook path is required")
	}
	codeSource := strings.TrimSpace(opts.UserFormCodeSource)
	if codeSource == "" {
		codeSource = "frm"
	}
	if codeSource != "frm" && codeSource != "sidecar" {
		return result, fmt.Errorf("userform code source must be one of frm, sidecar")
	}
	if err := workbookformat.ValidateProjectWorkbookPath(workbookPath); err != nil {
		return result, err
	}
	srcInfo, err := os.Stat(workbookPath)
	if err != nil {
		return result, fmt.Errorf("cannot read workbook: %w", err)
	}
	if srcInfo.IsDir() {
		return result, fmt.Errorf("workbook path is a directory: %s", workbookPath)
	}
	destPath := filepath.Join(cwd, "build", filepath.Base(workbookPath))
	result, err = createScaffold(cwd, destPath, projectName(workbookPath), func(path string) error {
		return copyFile(workbookPath, path)
	}, codeSource, false, false)
	if err != nil {
		return InitResult{}, err
	}
	return result, nil
}

func New(cwd, workbookName string, createWorkbook WorkbookCreator) (InitResult, error) {
	if createWorkbook == nil {
		return InitResult{}, errors.New("workbook creator is required")
	}
	name, err := normalizeWorkbookName(workbookName)
	if err != nil {
		return InitResult{}, err
	}
	destPath := filepath.Join(cwd, "build", name)
	return createScaffold(cwd, destPath, projectName(name), createWorkbook, "", true, true)
}

func createScaffold(cwd, destPath, name string, createWorkbook WorkbookCreator, userFormCodeSource string, scaffoldWorkbookModules bool, scaffoldRuntimeHelper bool) (InitResult, error) {
	var result InitResult
	configPath := filepath.Join(cwd, config.FileName)
	moduleDir := filepath.Join(cwd, "src", "modules")
	thisWorkbookPath := filepath.Join(cwd, "src", "workbook", "ThisWorkbook.bas")
	sheet1Path := filepath.Join(cwd, "src", "workbook", "Sheet1.bas")
	sampleTestPath := filepath.Join(cwd, "src", "modules", "Tests", "SampleTests.bas")
	protectedPaths := []string{destPath, configPath}
	protectedPaths = append(protectedPaths, bundledModuleConflictPaths(moduleDir, scaffoldBaseModuleTemplates())...)
	if scaffoldRuntimeHelper {
		protectedPaths = append(protectedPaths, bundledModuleConflictPaths(moduleDir, scaffoldRuntimeHelperModuleTemplates())...)
	}
	if scaffoldWorkbookModules {
		protectedPaths = append(protectedPaths, thisWorkbookPath, sheet1Path)
	}
	protectedPaths = append(protectedPaths, sampleTestPath)
	for _, path := range protectedPaths {
		if _, err := os.Stat(path); err == nil {
			return result, fmt.Errorf("refusing to overwrite existing file: %s", path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return result, err
		}
	}

	dirs := []string{
		filepath.Join(cwd, "src", "modules"),
		filepath.Join(cwd, "src", "modules", "Tests"),
		filepath.Join(cwd, "src", "classes"),
		filepath.Join(cwd, "src", "forms"),
		filepath.Join(cwd, "src", "workbook"),
		filepath.Join(cwd, "build"),
		filepath.Join(cwd, ".xlflow"),
	}
	for _, dir := range dirs {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return result, err
		}
		result.Created = append(result.Created, filepath.ToSlash(rel(cwd, dir)))
	}

	if err := createWorkbook(destPath); err != nil {
		return result, err
	}
	result.Workbook = filepath.ToSlash(rel(cwd, destPath))
	result.Created = append(result.Created, result.Workbook)

	cfg := config.Default()
	cfg.Project.Name = name
	cfg.Excel.Path = result.Workbook
	if strings.TrimSpace(userFormCodeSource) != "" {
		cfg.UserForm.CodeSource = userFormCodeSource
	}
	if err := config.Write(configPath, cfg); err != nil {
		return result, err
	}
	result.ConfigPath = config.FileName
	result.Created = append(result.Created, config.FileName)

	installedDefaults, err := installBundledModules(cwd, moduleDir, scaffoldBaseModuleTemplates())
	if err != nil {
		return result, err
	}
	result.Created = append(result.Created, installedDefaults.Created...)
	if scaffoldRuntimeHelper {
		installedHelpers, err := installBundledModules(cwd, moduleDir, scaffoldRuntimeHelperModuleTemplates())
		if err != nil {
			return result, err
		}
		result.Created = append(result.Created, installedHelpers.Created...)
	}
	if scaffoldWorkbookModules {
		for _, item := range []struct {
			path string
			body string
		}{
			{thisWorkbookPath, defaultDocumentModule},
			{sheet1Path, defaultDocumentModule},
		} {
			if err := writeExclusive(item.path, item.body); err != nil {
				return result, err
			}
			result.Created = append(result.Created, filepath.ToSlash(rel(cwd, item.path)))
		}
	}
	if err := writeExclusive(sampleTestPath, sampleTestModule); err != nil {
		return result, err
	}
	result.Created = append(result.Created, filepath.ToSlash(rel(cwd, sampleTestPath)))

	gitignorePath := filepath.Join(cwd, ".gitignore")
	updatedGitignore, err := ensureGitignore(gitignorePath)
	if err != nil {
		return result, err
	}
	if updatedGitignore {
		result.Created = append(result.Created, ".gitignore")
	}
	return result, nil
}

func InstallHelperModules(cwd string, src config.SourceConfig) (InstallModulesResult, error) {
	moduleRoot := ResolveModuleRoot(cwd, src)
	return installBundledModules(cwd, moduleRoot, installHelperModuleTemplates())
}

func InstallRuntimeHelperModules(cwd string, src config.SourceConfig) (InstallModulesResult, error) {
	moduleRoot := ResolveModuleRoot(cwd, src)
	return installBundledModules(cwd, moduleRoot, scaffoldRuntimeHelperModuleTemplates())
}

func ResolveModuleRoot(cwd string, src config.SourceConfig) string {
	moduleRoot := strings.TrimSpace(src.Modules)
	if moduleRoot == "" {
		moduleRoot = config.Default().Src.Modules
	}
	if !filepath.IsAbs(moduleRoot) {
		moduleRoot = filepath.Join(cwd, filepath.FromSlash(moduleRoot))
	}
	return moduleRoot
}

func ResolveClassRoot(cwd string, src config.SourceConfig) string {
	classRoot := strings.TrimSpace(src.Classes)
	if classRoot == "" {
		classRoot = config.Default().Src.Classes
	}
	if !filepath.IsAbs(classRoot) {
		classRoot = filepath.Join(cwd, filepath.FromSlash(classRoot))
	}
	return classRoot
}

func ResolveFormRoot(cwd string, src config.SourceConfig) string {
	formRoot := strings.TrimSpace(src.Forms)
	if formRoot == "" {
		formRoot = config.Default().Src.Forms
	}
	if !filepath.IsAbs(formRoot) {
		formRoot = filepath.Join(cwd, filepath.FromSlash(formRoot))
	}
	return formRoot
}

func NewModule(cwd, name, kind string, src config.SourceConfig) (NewModuleResult, error) {
	var result NewModuleResult
	cleanName, err := cleanComponentName(name)
	if err != nil {
		return result, err
	}
	kind = strings.TrimSpace(strings.ToLower(kind))

	var root string
	var ext string
	var body string
	switch kind {
	case "standard":
		root = ResolveModuleRoot(cwd, src)
		ext = ".bas"
		body = standardModuleTemplate(cleanName)
	case "class":
		root = ResolveClassRoot(cwd, src)
		ext = ".cls"
		body = classModuleTemplate(cleanName)
	default:
		return result, fmt.Errorf("%w: module type must be one of standard, class", ErrInvalidModuleType)
	}

	if err := rejectComponentNameCollision(cwd, src, cleanName); err != nil {
		return result, err
	}
	path := filepath.Join(root, cleanName+ext)
	if _, err := os.Stat(path); err == nil {
		return result, fmt.Errorf("%w: refusing to overwrite existing file: %s", ErrScaffoldExists, path)
	} else if !errors.Is(err, os.ErrNotExist) {
		return result, err
	}
	if err := os.MkdirAll(root, 0o755); err != nil {
		return result, err
	}
	if err := writeExclusive(path, body); err != nil {
		return result, err
	}

	result.Kind = kind
	result.Name = cleanName
	result.Path = filepath.ToSlash(rel(cwd, path))
	result.Created = []string{result.Path}
	return result, nil
}

func NewUserForm(cwd, name string, cfg config.Config) (NewFormResult, error) {
	var result NewFormResult
	if cfg.UserForm.CodeSource != "sidecar" {
		return result, ErrUserFormRequiresSidecar
	}
	cleanName, err := cleanComponentName(name)
	if err != nil {
		return result, err
	}
	if err := rejectComponentNameCollision(cwd, cfg.Src, cleanName); err != nil {
		return result, err
	}
	formRoot := ResolveFormRoot(cwd, cfg.Src)
	codeDir := filepath.Join(formRoot, "code")
	specDir := filepath.Join(formRoot, "specs")
	codePath := filepath.Join(codeDir, cleanName+".bas")
	specPath := filepath.Join(specDir, cleanName+".yaml")
	for _, path := range []string{codePath, specPath} {
		if _, err := os.Stat(path); err == nil {
			return result, fmt.Errorf("%w: refusing to overwrite existing file: %s", ErrScaffoldExists, path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return result, err
		}
	}
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		return result, err
	}
	if err := os.MkdirAll(specDir, 0o755); err != nil {
		return result, err
	}
	if err := writeExclusive(codePath, userFormCodeTemplate()); err != nil {
		return result, err
	}
	if err := writeExclusive(specPath, userFormSpecTemplate(cleanName)); err != nil {
		_ = os.Remove(codePath)
		return result, err
	}

	result.Name = cleanName
	result.CodePath = filepath.ToSlash(rel(cwd, codePath))
	result.SpecPath = filepath.ToSlash(rel(cwd, specPath))
	result.Created = []string{result.CodePath, result.SpecPath}
	result.CodeSource = "sidecar"
	return result, nil
}

func rejectComponentNameCollision(cwd string, src config.SourceConfig, name string) error {
	if strings.TrimSpace(name) == "" {
		return nil
	}
	roots := []struct {
		root             string
		exts             map[string]bool
		skipFormSidecars bool
	}{
		{root: ResolveModuleRoot(cwd, src), exts: map[string]bool{".bas": true}},
		{root: ResolveClassRoot(cwd, src), exts: map[string]bool{".cls": true}},
		{root: ResolveFormRoot(cwd, src), exts: map[string]bool{".frm": true}, skipFormSidecars: true},
	}
	for _, entry := range roots {
		if err := filepath.WalkDir(entry.root, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				if errors.Is(err, os.ErrNotExist) {
					return nil
				}
				return err
			}
			if d.IsDir() {
				base := strings.ToLower(d.Name())
				if entry.skipFormSidecars && (base == "code" || base == "specs") {
					return filepath.SkipDir
				}
				return nil
			}
			ext := strings.ToLower(filepath.Ext(path))
			if !entry.exts[ext] {
				return nil
			}
			if strings.EqualFold(strings.TrimSuffix(filepath.Base(path), filepath.Ext(path)), name) {
				return fmt.Errorf("%w: VBA component name %q already exists at %s", ErrScaffoldExists, name, path)
			}
			return nil
		}); err != nil {
			if errors.Is(err, os.ErrNotExist) {
				continue
			}
			return err
		}
	}
	formRoot := ResolveFormRoot(cwd, src)
	for _, path := range []string{
		filepath.Join(formRoot, "code", name+".bas"),
		filepath.Join(formRoot, "specs", name+".yaml"),
		filepath.Join(formRoot, "specs", name+".yml"),
	} {
		if _, err := os.Stat(path); err == nil {
			return fmt.Errorf("%w: VBA component name %q already exists at %s", ErrScaffoldExists, name, path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return err
		}
	}
	return nil
}

func cleanComponentName(name string) (string, error) {
	cleanName := strings.TrimSpace(name)
	cleanName = strings.TrimSuffix(cleanName, ".bas")
	cleanName = strings.TrimSuffix(cleanName, ".cls")
	cleanName = strings.TrimSuffix(cleanName, ".frm")
	if cleanName == "" {
		return "", fmt.Errorf("%w: component name is required", ErrInvalidComponentName)
	}
	if cleanName != filepath.Base(cleanName) ||
		strings.Contains(cleanName, "..") ||
		strings.Contains(cleanName, "/") ||
		strings.Contains(cleanName, `\`) {
		return "", ErrInvalidComponentName
	}
	for i, r := range cleanName {
		if i == 0 {
			if (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z') || r == '_' {
				continue
			}
			return "", fmt.Errorf("%w: component name must start with an ASCII letter or underscore", ErrInvalidComponentName)
		}
		if (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z') || (r >= '0' && r <= '9') || r == '_' {
			continue
		}
		return "", fmt.Errorf("%w: component name may contain only ASCII letters, digits, and underscores", ErrInvalidComponentName)
	}
	return cleanName, nil
}

func standardModuleTemplate(name string) string {
	return fmt.Sprintf("Attribute VB_Name = %q\nOption Explicit\n", name)
}

func classModuleTemplate(name string) string {
	return fmt.Sprintf("VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nAttribute VB_Name = %q\nOption Explicit\n", name)
}

func userFormCodeTemplate() string {
	return "Option Explicit\n"
}

func userFormSpecTemplate(name string) string {
	return fmt.Sprintf(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: %s
  caption: %s
controls: []
warnings: []

# Example controls:
# controls:
#   - id: lblTitle
#     name: lblTitle
#     type: Label
#     caption: Title
#     left: 12
#     top: 12
#     width: 120
#     height: 18
`, name, name)
}

func installBundledModules(cwd, moduleDir string, templates []bundledModuleTemplate) (InstallModulesResult, error) {
	var result InstallModulesResult
	paths := bundledModulePaths(moduleDir, templates)
	for _, path := range bundledModuleConflictPaths(moduleDir, templates) {
		if _, err := os.Stat(path); err == nil {
			return result, fmt.Errorf("refusing to overwrite existing file: %s", path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return result, err
		}
	}
	for index, template := range templates {
		path := paths[index]
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			return result, err
		}
		if err := writeExclusive(path, template.body); err != nil {
			return result, err
		}
		result.Created = append(result.Created, filepath.ToSlash(rel(cwd, path)))
	}
	return result, nil
}

func bundledModulePaths(moduleDir string, templates []bundledModuleTemplate) []string {
	paths := make([]string, 0, len(templates))
	for _, template := range templates {
		paths = append(paths, filepath.Join(moduleDir, template.fileName))
	}
	return paths
}

// bundledModuleConflictPaths includes the legacy root location for bundled
// helpers. This prevents new scaffolds and module install from creating two
// source files for the same VBA component in existing projects.
func bundledModuleConflictPaths(moduleDir string, templates []bundledModuleTemplate) []string {
	paths := bundledModulePaths(moduleDir, templates)
	for _, template := range templates {
		if filepath.Dir(template.fileName) == "Xlflow" {
			paths = append(paths, filepath.Join(moduleDir, filepath.Base(template.fileName)))
		}
	}
	return paths
}

func installHelperModuleTemplates() []bundledModuleTemplate {
	return []bundledModuleTemplate{
		{fileName: filepath.Join("Xlflow", "XlflowAssert.bas"), body: defaultAssertModule},
		{fileName: filepath.Join("Xlflow", "XlflowRuntime.bas"), body: defaultRuntimeModule},
		{fileName: filepath.Join("Xlflow", "XlflowUI.bas"), body: defaultUIRuntimeModule},
		{fileName: filepath.Join("Xlflow", "XlflowDebug.bas"), body: defaultDebugRuntimeModule},
	}
}

func scaffoldRuntimeHelperModuleTemplates() []bundledModuleTemplate {
	return []bundledModuleTemplate{
		{fileName: filepath.Join("Xlflow", "XlflowRuntime.bas"), body: defaultRuntimeModule},
		{fileName: filepath.Join("Xlflow", "XlflowUI.bas"), body: defaultUIRuntimeModule},
		{fileName: filepath.Join("Xlflow", "XlflowDebug.bas"), body: defaultDebugRuntimeModule},
	}
}

func scaffoldBaseModuleTemplates() []bundledModuleTemplate {
	return []bundledModuleTemplate{
		{fileName: filepath.Join("Xlflow", "XlflowAssert.bas"), body: defaultAssertModule},
		{fileName: "Main.bas", body: defaultMainModule},
		{fileName: "App.bas", body: defaultAppModule},
		{fileName: "Ui.bas", body: defaultUiModule},
	}
}

func copyFile(src, dest string) (err error) {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := in.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	out, err := os.OpenFile(dest, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0o644)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := out.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	if _, err := io.Copy(out, in); err != nil {
		return err
	}
	return nil
}

func writeExclusive(path, body string) (err error) {
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0o644)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := f.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	_, err = f.WriteString(body)
	return err
}

func ensureGitignore(path string) (bool, error) {
	body, err := os.ReadFile(path)
	if errors.Is(err, os.ErrNotExist) {
		if err := writeExclusive(path, defaultGitignore); err != nil {
			return false, err
		}
		return true, nil
	}
	if err != nil {
		return false, err
	}

	missing := missingGitignoreSections(string(body))
	if len(missing) == 0 {
		return false, nil
	}

	appendBody := strings.Join(missing, "\n\n") + "\n"
	if len(body) > 0 {
		lineEnding := "\n"
		if strings.Contains(string(body), "\r\n") {
			lineEnding = "\r\n"
			appendBody = strings.ReplaceAll(appendBody, "\n", "\r\n")
		}
		if !strings.HasSuffix(string(body), "\n") {
			appendBody = lineEnding + lineEnding + appendBody
		} else if !strings.HasSuffix(string(body), "\n\n") {
			appendBody = lineEnding + appendBody
		}
	}

	f, err := os.OpenFile(path, os.O_WRONLY|os.O_APPEND, 0o644)
	if err != nil {
		return false, err
	}
	if _, err := f.WriteString(appendBody); err != nil {
		_ = f.Close()
		return false, err
	}
	if err := f.Close(); err != nil {
		return false, err
	}
	return true, nil
}

func missingGitignoreSections(body string) []string {
	lines := map[string]bool{}
	for _, line := range strings.FieldsFunc(body, func(r rune) bool {
		return r == '\n' || r == '\r'
	}) {
		lines[strings.TrimSpace(line)] = true
	}

	var sections []string
	if !lines["~$*.xls*"] || !lines["*.tmp"] {
		var entries []string
		if !lines["~$*.xls*"] {
			entries = append(entries, "~$*.xls*")
		}
		if !lines["*.tmp"] {
			entries = append(entries, "*.tmp")
		}
		sections = append(sections, "# Excel\n"+strings.Join(entries, "\n"))
	}
	if !lines[".xlflow/"] || !lines["build/"] {
		var entries []string
		if !lines[".xlflow/"] {
			entries = append(entries, ".xlflow/")
		}
		if !lines["build/"] {
			entries = append(entries, "build/")
		}
		sections = append(sections, "# xlflow\n"+strings.Join(entries, "\n"))
	}
	return sections
}

func rel(base, path string) string {
	r, err := filepath.Rel(base, path)
	if err != nil {
		return path
	}
	return r
}

func projectName(path string) string {
	name := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	name = strings.TrimSpace(name)
	if name == "" {
		return "sample"
	}
	return name
}

func normalizeWorkbookName(name string) (string, error) {
	return workbookformat.NormalizeProjectWorkbookName(name)
}

const defaultGitignore = `# Excel
~$*.xls*
*.tmp

# xlflow
.xlflow/
build/
`

const defaultAssertModule = `Attribute VB_Name = "XlflowAssert"
Option Explicit

Private Const assertFailureNumber As Long = vbObjectError + 513

''' Asserts that two scalar values are equal.
'''
''' Args:
'''     expected: Expected scalar value.
'''     actual: Actual scalar value.
'''     message: Optional failure message prefix.
Public Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If IsObject(expected) Or IsObject(actual) Then
    Err.Raise vbObjectError + 514, "XlflowAssert.AssertEquals", "AssertEquals supports scalar values only. Compare object properties such as Range.Value2."
  End If

  If IsArray(expected) Or IsArray(actual) Then
    Err.Raise vbObjectError + 515, "XlflowAssert.AssertEquals", "AssertEquals supports scalar values only. Array comparison is not supported."
  End If

  If IsNull(expected) Or IsNull(actual) Then
    If IsNull(expected) And IsNull(actual) Then
      Exit Sub
    End If
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertEquals"
  End If

  If expected <> actual Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertEquals"
  End If
End Sub

''' Asserts that two scalar values have the same VarType and value.
'''
''' Args:
'''     expected: Expected scalar value.
'''     actual: Actual scalar value.
'''     message: Optional failure message prefix.
Public Sub AssertStrictEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If IsObject(expected) Or IsObject(actual) Then
    RaiseAssertFailure message, "AssertStrictEquals supports scalar values only. Object comparison is not supported.", "XlflowAssert.AssertStrictEquals"
  End If

  If IsArray(expected) Or IsArray(actual) Then
    RaiseAssertFailure message, "AssertStrictEquals supports scalar values only. Array comparison is not supported.", "XlflowAssert.AssertStrictEquals"
  End If

  If VarType(expected) <> VarType(actual) Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertStrictEquals"
  End If

  If IsNull(expected) Or IsEmpty(expected) Then
    Exit Sub
  End If

  If expected <> actual Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " but got " & FormatAssertValue(actual), "XlflowAssert.AssertStrictEquals"
  End If
End Sub

''' Asserts that two scalar values are different.
'''
''' Args:
'''     expected: Value that should differ from actual.
'''     actual: Value that should differ from expected.
'''     message: Optional failure message prefix.
Public Sub AssertNotEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If IsObject(expected) Or IsObject(actual) Then
    Err.Raise vbObjectError + 514, "XlflowAssert.AssertNotEqual", "AssertNotEqual supports scalar values only."
  End If

  If IsArray(expected) Or IsArray(actual) Then
    Err.Raise vbObjectError + 515, "XlflowAssert.AssertNotEqual", "AssertNotEqual supports scalar values only."
  End If

  If IsNull(expected) And IsNull(actual) Then
    RaiseAssertFailure message, "expected values to differ, but both are Null", "XlflowAssert.AssertNotEqual"
    Exit Sub
  End If

  If IsNull(expected) Or IsNull(actual) Then
    Exit Sub
  End If

  If expected = actual Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " to differ from " & FormatAssertValue(actual), "XlflowAssert.AssertNotEqual"
  End If
End Sub

''' Asserts that a value is Null.
Public Sub AssertNull(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsNull(value) Then
    RaiseAssertFailure message, "expected <Null> but got " & FormatAssertValue(value), "XlflowAssert.AssertNull"
  End If
End Sub

''' Asserts that a value is not Null.
Public Sub AssertNotNull(ByVal value As Variant, Optional ByVal message As String = "")
  If IsNull(value) Then
    RaiseAssertFailure message, "expected a non-Null value but got <Null>", "XlflowAssert.AssertNotNull"
  End If
End Sub

''' Asserts that a value is Empty according to VBA IsEmpty.
Public Sub AssertEmpty(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsEmpty(value) Then
    RaiseAssertFailure message, "expected <Empty> but got " & FormatAssertValue(value), "XlflowAssert.AssertEmpty"
  End If
End Sub

''' Asserts that a value is not Empty according to VBA IsEmpty.
Public Sub AssertNotEmpty(ByVal value As Variant, Optional ByVal message As String = "")
  If IsEmpty(value) Then
    RaiseAssertFailure message, "expected a non-Empty value but got <Empty>", "XlflowAssert.AssertNotEmpty"
  End If
End Sub

''' Asserts that two numeric values are within a non-negative tolerance.
Public Sub AssertNear(ByVal expected As Variant, ByVal actual As Variant, ByVal tolerance As Double, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertNear"

  If tolerance < 0 Then
    RaiseAssertFailure message, "tolerance must be non-negative but got " & FormatAssertValue(tolerance), source
  End If

  If Not IsAssertNumeric(expected) Then
    RaiseAssertFailure message, "expected value must be numeric but got " & FormatAssertValue(expected), source
  End If

  If Not IsAssertNumeric(actual) Then
    RaiseAssertFailure message, "actual value must be numeric but got " & FormatAssertValue(actual), source
  End If

  Dim difference As Double
  difference = Abs(CDbl(expected) - CDbl(actual))
  If difference > tolerance Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(expected) & " +/- " & FormatAssertValue(tolerance) & " but got " & FormatAssertValue(actual) & "; difference was " & FormatAssertValue(difference), source
  End If
End Sub

''' Asserts that a string contains an expected substring using binary comparison.
Public Sub AssertContains(ByVal expectedSubstring As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsAssertString(expectedSubstring) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertContains expects string values; expected substring was " & FormatAssertValue(expectedSubstring) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertContains"
  End If

  If InStr(1, CStr(actual), CStr(expectedSubstring), vbBinaryCompare) = 0 Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to contain " & FormatAssertValue(expectedSubstring), "XlflowAssert.AssertContains"
  End If
End Sub

''' Asserts that a string starts with an expected prefix using binary comparison.
Public Sub AssertStartsWith(ByVal prefix As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsAssertString(prefix) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertStartsWith expects string values; prefix was " & FormatAssertValue(prefix) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertStartsWith"
  End If

  If Left$(CStr(actual), Len(CStr(prefix))) <> CStr(prefix) Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to start with " & FormatAssertValue(prefix), "XlflowAssert.AssertStartsWith"
  End If
End Sub

''' Asserts that a string ends with an expected suffix using binary comparison.
Public Sub AssertEndsWith(ByVal suffix As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsAssertString(suffix) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertEndsWith expects string values; suffix was " & FormatAssertValue(suffix) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertEndsWith"
  End If

  If Right$(CStr(actual), Len(CStr(suffix))) <> CStr(suffix) Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to end with " & FormatAssertValue(suffix), "XlflowAssert.AssertEndsWith"
  End If
End Sub

''' Asserts that a string matches a VBScript.RegExp pattern.
Public Sub AssertMatches(ByVal pattern As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertMatches"

  If Not IsAssertString(pattern) Or Not IsAssertString(actual) Then
    RaiseAssertFailure message, "AssertMatches expects string values; pattern was " & FormatAssertValue(pattern) & " and actual was " & FormatAssertValue(actual), source
  End If

  Dim regex As Object
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  Set regex = CreateObject("VBScript.RegExp")
  If Err.Number <> 0 Then
    Dim createError As String
    createError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "VBScript.RegExp is not available: " & createError, source
  End If
  regex.Pattern = CStr(pattern)
  regex.IgnoreCase = False
  regex.Global = False
  regex.MultiLine = False
  If Err.Number <> 0 Then
    Dim patternError As String
    patternError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "invalid regex pattern " & FormatAssertValue(pattern) & ": " & patternError, source
  End If

  Dim matched As Boolean
  matched = regex.Test(CStr(actual))
  If Err.Number <> 0 Then
    Dim testError As String
    testError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "invalid regex pattern " & FormatAssertValue(pattern) & ": " & testError, source
  End If
  On Error GoTo 0

  If Not matched Then
    RaiseAssertFailure message, "expected " & FormatAssertValue(actual) & " to match pattern " & FormatAssertValue(pattern), source
  End If
End Sub

''' Asserts that one- or two-dimensional scalar arrays are equal.
Public Sub AssertArrayEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertArrayEquals"

  If Not IsArray(expected) Or Not IsArray(actual) Then
    RaiseAssertFailure message, "AssertArrayEquals expects arrays; expected was " & FormatAssertValue(expected) & " and actual was " & FormatAssertValue(actual), source
  End If

  Dim expectedDims As Long
  Dim actualDims As Long
  expectedDims = ArrayDimensionCount(expected)
  actualDims = ArrayDimensionCount(actual)
  If expectedDims <> actualDims Then
    RaiseAssertFailure message, "array dimensions differ; expected " & CStr(expectedDims) & " but got " & CStr(actualDims), source
  End If
  If expectedDims < 1 Or expectedDims > 2 Then
    RaiseAssertFailure message, "AssertArrayEquals supports one- and two-dimensional arrays only; got " & CStr(expectedDims) & " dimensions", source
  End If

  Dim dimension As Long
  For dimension = 1 To expectedDims
    If LBound(expected, dimension) <> LBound(actual, dimension) Or UBound(expected, dimension) <> UBound(actual, dimension) Then
      RaiseAssertFailure message, "array bounds differ on dimension " & CStr(dimension) & vbCrLf & "expected <" & CStr(LBound(expected, dimension)) & " To " & CStr(UBound(expected, dimension)) & ">" & vbCrLf & "actual   <" & CStr(LBound(actual, dimension)) & " To " & CStr(UBound(actual, dimension)) & ">", source
    End If
  Next dimension

  Dim i As Long
  Dim j As Long
  If expectedDims = 1 Then
    For i = LBound(expected, 1) To UBound(expected, 1)
      If Not ScalarArrayValuesEqual(expected(i), actual(i)) Then
        RaiseAssertFailure message, "array mismatch at (" & CStr(i) & ")" & vbCrLf & "expected " & FormatAssertValue(expected(i)) & vbCrLf & "actual   " & FormatAssertValue(actual(i)), source
      End If
    Next i
  Else
    For i = LBound(expected, 1) To UBound(expected, 1)
      For j = LBound(expected, 2) To UBound(expected, 2)
        If Not ScalarArrayValuesEqual(expected(i, j), actual(i, j)) Then
          RaiseAssertFailure message, "array mismatch at (" & CStr(i) & ", " & CStr(j) & ")" & vbCrLf & "expected " & FormatAssertValue(expected(i, j)) & vbCrLf & "actual   " & FormatAssertValue(actual(i, j)), source
        End If
      Next j
    Next i
  End If
End Sub

''' Asserts that expected scalar or two-dimensional array values equal Range.Value2.
Public Sub AssertRangeEquals(ByVal expected As Variant, ByVal actualRange As Object, Optional ByVal message As String = "")
  Dim source As String
  source = "XlflowAssert.AssertRangeEquals"

  If actualRange Is Nothing Then
    RaiseAssertFailure message, "actualRange must be an Excel Range object but got <Nothing>", source
  End If

  Dim rowCount As Long
  Dim columnCount As Long
  Dim actualValues As Variant
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  rowCount = CLng(actualRange.Rows.Count)
  columnCount = CLng(actualRange.Columns.Count)
  actualValues = actualRange.Value2
  If Err.Number <> 0 Then
    Dim rangeError As String
    rangeError = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseAssertFailure message, "actualRange must expose Range members Rows, Columns, Cells, and Value2: " & rangeError, source
  End If
  On Error GoTo 0

  If rowCount = 1 And columnCount = 1 Then
    If IsArray(expected) Then
      RaiseAssertFailure message, "scalar expected value is required for a single-cell range but got " & FormatAssertValue(expected), source
    End If
    If Not ScalarArrayValuesEqual(expected, actualValues) Then
      RaiseAssertFailure message, "range mismatch at " & RangeCellLabel(actualRange, 1, 1) & vbCrLf & "expected " & FormatAssertValue(expected) & vbCrLf & "actual   " & FormatAssertValue(actualValues), source
    End If
    Exit Sub
  End If

  If Not IsArray(expected) Then
    RaiseAssertFailure message, "two-dimensional expected array is required for a multi-cell range but got " & FormatAssertValue(expected), source
  End If

  If ArrayDimensionCount(expected) <> 2 Then
    RaiseAssertFailure message, "expected array for a multi-cell range must be two-dimensional", source
  End If

  Dim expectedRows As Long
  Dim expectedColumns As Long
  expectedRows = UBound(expected, 1) - LBound(expected, 1) + 1
  expectedColumns = UBound(expected, 2) - LBound(expected, 2) + 1
  If expectedRows <> rowCount Or expectedColumns <> columnCount Then
    RaiseAssertFailure message, "range size differs; expected <" & CStr(expectedRows) & " x " & CStr(expectedColumns) & "> but got <" & CStr(rowCount) & " x " & CStr(columnCount) & ">", source
  End If

  Dim rowIndex As Long
  Dim columnIndex As Long
  Dim expectedRow As Long
  Dim expectedColumn As Long
  For rowIndex = 1 To rowCount
    expectedRow = LBound(expected, 1) + rowIndex - 1
    For columnIndex = 1 To columnCount
      expectedColumn = LBound(expected, 2) + columnIndex - 1
      If Not ScalarArrayValuesEqual(expected(expectedRow, expectedColumn), actualValues(rowIndex, columnIndex)) Then
        RaiseAssertFailure message, "range mismatch at " & RangeCellLabel(actualRange, rowIndex, columnIndex) & vbCrLf & "expected " & FormatAssertValue(expected(expectedRow, expectedColumn)) & vbCrLf & "actual   " & FormatAssertValue(actualValues(rowIndex, columnIndex)), source
      End If
    Next columnIndex
  Next rowIndex
End Sub

''' Asserts that two object references are the same reference.
Public Sub AssertSame(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsObject(expected) Or Not IsObject(actual) Then
    RaiseAssertFailure message, "AssertSame expects object references; expected was " & FormatAssertValue(expected) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertSame"
  End If

  If Not (expected Is actual) Then
    RaiseAssertFailure message, "expected same object reference but got different references", "XlflowAssert.AssertSame"
  End If
End Sub

''' Asserts that two object references are not the same reference.
Public Sub AssertNotSame(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  If Not IsObject(expected) Or Not IsObject(actual) Then
    RaiseAssertFailure message, "AssertNotSame expects object references; expected was " & FormatAssertValue(expected) & " and actual was " & FormatAssertValue(actual), "XlflowAssert.AssertNotSame"
  End If

  If expected Is actual Then
    RaiseAssertFailure message, "expected different object references but both were the same", "XlflowAssert.AssertNotSame"
  End If
End Sub

''' Asserts that a condition is True.
'''
''' Args:
'''     condition: Boolean condition to verify.
'''     message: Optional failure message.
Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
  If Not condition Then
    RaiseAssertFailure message, "expected True but got False", "XlflowAssert.AssertTrue"
  End If
End Sub

''' Asserts that a condition is False.
'''
''' Args:
'''     condition: Boolean condition to verify.
'''     message: Optional failure message.
Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
  If condition Then
    RaiseAssertFailure message, "expected False but got True", "XlflowAssert.AssertFalse"
  End If
End Sub

''' Fails the current test immediately.
'''
''' Args:
'''     message: Optional failure message.
Public Sub AssertFail(Optional ByVal message As String = "")
  RaiseAssertFailure message, "assertion failed", "XlflowAssert.AssertFail"
End Sub

''' Marks the current test as inconclusive.
'''
''' Args:
'''     message: Optional inconclusive reason.
Public Sub AssertInconclusive(Optional ByVal message As String = "")
  Dim detail As String
  detail = "inconclusive"
  If Len(message) > 0 Then
    detail = message
  End If
  Err.Raise vbObjectError + 516, "XlflowAssert.AssertInconclusive", detail
End Sub

''' Asserts that an object reference is Nothing.
'''
''' Args:
'''     value: Object reference to verify.
'''     message: Optional failure message.
Public Sub AssertIsNothing(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsObject(value) Then
    RaiseAssertFailure message, "expected an object but got a non-object", "XlflowAssert.AssertIsNothing"
    Exit Sub
  End If
  If Not (value Is Nothing) Then
    RaiseAssertFailure message, "expected Nothing but got an object reference", "XlflowAssert.AssertIsNothing"
  End If
End Sub

''' Asserts that an object reference is not Nothing.
'''
''' Args:
'''     value: Object reference to verify.
'''     message: Optional failure message.
Public Sub AssertIsNotNothing(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsObject(value) Then
    RaiseAssertFailure message, "expected an object but got a non-object", "XlflowAssert.AssertIsNotNothing"
    Exit Sub
  End If
  If value Is Nothing Then
    RaiseAssertFailure message, "expected an object reference but got Nothing", "XlflowAssert.AssertIsNotNothing"
  End If
End Sub

Private Sub RaiseAssertFailure(ByVal message As String, ByVal detail As String, ByVal source As String)
  If Len(message) > 0 Then
    detail = message & ": " & detail
  End If
  Err.Raise assertFailureNumber, source, detail
End Sub

Private Function FormatAssertValue(ByVal value As Variant) As String
  If IsObject(value) Then
    If value Is Nothing Then
      FormatAssertValue = "<Nothing>"
    Else
      FormatAssertValue = "<Object: " & TypeName(value) & ">"
    End If
    Exit Function
  End If

  If IsArray(value) Then
    FormatAssertValue = "<Array>"
    Exit Function
  End If

  If IsNull(value) Then
    FormatAssertValue = "<Null>"
  ElseIf IsEmpty(value) Then
    FormatAssertValue = "<Empty>"
  ElseIf VarType(value) = vbString Then
    FormatAssertValue = "<String: """ & EscapeAssertString(CStr(value)) & """>"
  ElseIf VarType(value) = vbBoolean Then
    If CBool(value) Then
      FormatAssertValue = "<Boolean: True>"
    Else
      FormatAssertValue = "<Boolean: False>"
    End If
  ElseIf VarType(value) = vbDate Then
    FormatAssertValue = "<Date: " & Format$(CDate(value), "yyyy-mm-dd hh:nn:ss") & ">"
  Else
    FormatAssertValue = "<" & AssertValueTypeName(value) & ": " & CStr(value) & ">"
  End If
End Function

Private Function EscapeAssertString(ByVal value As String) As String
  EscapeAssertString = Replace(value, """", """""")
  EscapeAssertString = Replace(EscapeAssertString, vbCrLf, "\r\n")
  EscapeAssertString = Replace(EscapeAssertString, vbCr, "\r")
  EscapeAssertString = Replace(EscapeAssertString, vbLf, "\n")
  EscapeAssertString = Replace(EscapeAssertString, vbTab, "\t")
End Function

Private Function AssertValueTypeName(ByVal value As Variant) As String
  Select Case VarType(value)
    Case vbByte
      AssertValueTypeName = "Byte"
    Case vbInteger
      AssertValueTypeName = "Integer"
    Case vbLong
      AssertValueTypeName = "Long"
    Case vbSingle
      AssertValueTypeName = "Single"
    Case vbDouble
      AssertValueTypeName = "Double"
    Case vbCurrency
      AssertValueTypeName = "Currency"
    Case vbDecimal
      AssertValueTypeName = "Decimal"
    Case Else
      AssertValueTypeName = TypeName(value)
  End Select
End Function

Private Function IsAssertNumeric(ByVal value As Variant) As Boolean
  If IsObject(value) Or IsArray(value) Or IsNull(value) Or IsEmpty(value) Then
    IsAssertNumeric = False
  Else
    IsAssertNumeric = IsNumeric(value)
  End If
End Function

Private Function IsAssertString(ByVal value As Variant) As Boolean
  IsAssertString = Not IsObject(value) And Not IsArray(value) And VarType(value) = vbString
End Function

Private Function ScalarArrayValuesEqual(ByVal expected As Variant, ByVal actual As Variant) As Boolean
  If IsObject(expected) Or IsObject(actual) Or IsArray(expected) Or IsArray(actual) Then
    ScalarArrayValuesEqual = False
  ElseIf IsNull(expected) Or IsNull(actual) Then
    ScalarArrayValuesEqual = IsNull(expected) And IsNull(actual)
  Else
    ScalarArrayValuesEqual = (expected = actual)
  End If
End Function

Private Function ArrayDimensionCount(ByVal value As Variant) As Long
  Dim dimension As Long
  Dim lowerBound As Long ' xlflow:disable-line VB020
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  For dimension = 1 To 60
    Err.Clear
    lowerBound = LBound(value, dimension)
    If Err.Number <> 0 Then
      Exit For
    End If
  Next dimension
  On Error GoTo 0
  ArrayDimensionCount = dimension - 1
End Function

Private Function RangeCellLabel(ByVal actualRange As Object, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
  Dim address As String
  Dim sheetName As String
  ' xlflow:disable-next-line VB004
  On Error Resume Next
  address = actualRange.Cells(rowIndex, columnIndex).Address(False, False)
  sheetName = actualRange.Worksheet.Name
  If Err.Number <> 0 Then
    Err.Clear
    RangeCellLabel = "row " & CStr(rowIndex) & ", column " & CStr(columnIndex)
  ElseIf Len(sheetName) > 0 Then
    RangeCellLabel = sheetName & "!" & address
  Else
    RangeCellLabel = address
  End If
  On Error GoTo 0
End Function
`

const defaultRuntimeModule = `Attribute VB_Name = "XlflowRuntime"
Option Explicit

' XlflowRuntime exposes the execution mode that xlflow injected before user VBA started.
' Use these helpers when workbook code must branch between interactive and unattended flows.
Private Const xlflowInteractive As Long = 0
Private Const xlflowHeadless As Long = 1
Private Const xlflowCI As Long = 2
Private Const xlflowAgent As Long = 3
Private Const xlflowTest As Long = 4

''' Returns the current xlflow runtime mode as a stable numeric value.
'''
''' Returns:
'''     One of the internal xlflow mode constants.
Public Function Mode() As Long
	Select Case ModeName()
		Case "headless"
			Mode = xlflowHeadless
		Case "ci"
			Mode = xlflowCI
		Case "agent"
			Mode = xlflowAgent
		Case "test"
			Mode = xlflowTest
		Case Else
			Mode = xlflowInteractive
	End Select
End Function

''' Returns the normalized runtime mode name injected by xlflow.
'''
''' Returns:
'''     interactive, headless, ci, agent, or test.
Public Function ModeName() As String
	Dim raw As String
	raw = ReadWorkbookModeName()
	If Len(raw) = 0 Then
		raw = Environ$("XLFLOW_MODE")
	End If
	raw = LCase$(Trim$(raw))

	Select Case raw
		Case "headless", "ci", "agent", "test"
			ModeName = raw
		Case Else
			ModeName = "interactive"
	End Select
End Function

''' Indicates whether the workbook is running in normal human-driven Excel usage.
'''
''' Returns:
'''     True when the runtime mode is interactive.
Public Function IsInteractive() As Boolean
	IsInteractive = (Mode() = xlflowInteractive)
End Function

''' Indicates whether the workbook is running without direct human interaction.
'''
''' Returns:
'''     True for headless, CI, agent, and test modes.
Public Function IsHeadless() As Boolean
	Select Case Mode()
		Case xlflowHeadless, xlflowCI, xlflowAgent, xlflowTest
			IsHeadless = True
		Case Else
			IsHeadless = False
	End Select
End Function

''' Indicates whether the workbook is running in CI mode.
'''
''' Returns:
'''     True when the runtime mode is ci.
Public Function IsCI() As Boolean
	IsCI = (Mode() = xlflowCI)
End Function

''' Indicates whether the workbook is running in agent mode.
'''
''' Returns:
'''     True when the runtime mode is agent.
Public Function IsAgent() As Boolean
	IsAgent = (Mode() = xlflowAgent)
End Function

''' Indicates whether the workbook is running under xlflow test.
'''
''' Returns:
'''     True when the runtime mode is test.
Public Function IsTest() As Boolean
	IsTest = (Mode() = xlflowTest)
End Function

Private Function ReadWorkbookModeName() As String
	On Error GoTo Missing
	ReadWorkbookModeName = DecodeWorkbookDefinedName(ThisWorkbook.Names("__XLFLOW_MODE__").RefersTo)
	Exit Function

Missing:
	ReadWorkbookModeName = ""
End Function

Private Function DecodeWorkbookDefinedName(ByVal refersTo As String) As String
	If Len(refersTo) = 0 Then
		DecodeWorkbookDefinedName = ""
		Exit Function
	End If
	If Left$(refersTo, 1) = "=" Then
		refersTo = Mid$(refersTo, 2)
	End If
	If Len(refersTo) >= 2 Then
		If Left$(refersTo, 1) = Chr$(34) And Right$(refersTo, 1) = Chr$(34) Then
			refersTo = Mid$(refersTo, 2, Len(refersTo) - 2)
		End If
	End If
	DecodeWorkbookDefinedName = Replace$(refersTo, Chr$(34) & Chr$(34), Chr$(34))
End Function
`

const defaultUIRuntimeModule = `Attribute VB_Name = "XlflowUI"
Option Explicit

' XlflowUI keeps one workbook-side dialog API that works in both interactive Excel and xlflow headless runs.
' Use stable dialog ids because xlflow maps scripted responses to those ids during unattended execution.
Private Const xlflowResponseErrorBase As Long = vbObjectError + 520
Private Const xlflowErrInvalidDialogId As Long = xlflowResponseErrorBase + 1
Private Const xlflowErrInvalidMsgBoxResponse As Long = xlflowResponseErrorBase + 2
Private Const xlflowErrMissingInputResponse As Long = xlflowResponseErrorBase + 3
Private Const xlflowErrMissingFileDialogResponse As Long = xlflowResponseErrorBase + 4
Private Const xlflowErrInvalidFileDialogResponse As Long = xlflowResponseErrorBase + 5
Private Const xlflowFileDialogCancelToken As String = "@cancel"
Private Const xlflowStreamHelperName As String = "__XLFLOW_UI_STREAM_HELPER__"
Private Const xlflowStreamRedactInputName As String = "__XLFLOW_UI_STREAM_REDACT_INPUT__"

''' Shows a message box or resolves a scripted response in headless mode.
'''
''' Args:
'''     Id: Stable dialog id used by xlflow --msgbox.
'''     Prompt: Message shown to the user.
'''     Buttons: VBA MsgBox button and icon flags.
'''     Title: Optional dialog title.
'''     DefaultResponse: Optional headless fallback response token.
'''
''' Returns:
'''     The resolved VbMsgBoxResult value.
Public Function MsgBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal DefaultResponse As String = "") As VbMsgBoxResult
	Dim responseSource As String
	Dim responseToken As String
	ValidateDialogId Id, "XlflowUI.MsgBox"
	If XlflowRuntime.IsHeadless() Then
		On Error GoTo HandleHeadlessFailure
		MsgBox = ResolveMsgBoxResponse(Id, DefaultResponse, responseSource, responseToken)
		EmitHeadlessUIEvent "msgbox", Id, Prompt, Title, responseSource, responseToken, "", False, ""
		Exit Function
	End If

	MsgBox = VBA.Interaction.MsgBox(Prompt, Buttons, Title)
	Exit Function

HandleHeadlessFailure:
	EmitHeadlessUIEvent "msgbox", Id, Prompt, Title, "error", "", "", False, Err.Description
	Err.Raise Err.Number, Err.Source, Err.Description
End Function

''' Shows an input box or resolves a scripted value in headless mode.
'''
''' Args:
'''     Id: Stable dialog id used by xlflow --inputbox.
'''     Prompt: Prompt shown to the user.
'''     Title: Optional dialog title.
'''     Default: Interactive Excel default text.
'''     DefaultValue: Optional headless fallback value.
'''
''' Returns:
'''     The entered or scripted input value.
Public Function InputBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "", Optional ByVal DefaultValue As String = "") As String
	Dim responseSource As String
	Dim responseValue As String
	Dim displayValue As String
	Dim redacted As Boolean
	ValidateDialogId Id, "XlflowUI.InputBox"
	If XlflowRuntime.IsHeadless() Then
		On Error GoTo HandleHeadlessFailure
		responseValue = ResolveInputResponse(Id, DefaultValue, responseSource)
		redacted = ShouldRedactInputStream()
		displayValue = responseValue
		If redacted Then
			displayValue = "[redacted]"
		End If
		EmitHeadlessUIEvent "inputbox", Id, Prompt, Title, responseSource, "", displayValue, redacted, ""
		InputBox = responseValue
		Exit Function
	End If

	InputBox = VBA.Interaction.InputBox(Prompt, Title, Default)
	Exit Function

HandleHeadlessFailure:
	redacted = ShouldRedactInputStream()
	displayValue = ""
	If redacted Then
		displayValue = "[redacted]"
	End If
	EmitHeadlessUIEvent "inputbox", Id, Prompt, Title, "error", "", displayValue, redacted, Err.Description
	Err.Raise Err.Number, Err.Source, Err.Description
End Function

''' Opens Excel's file picker or resolves scripted file paths in headless mode.
'''
''' Args:
'''     Id: Stable dialog id used by xlflow --filedialog get-open.
'''     FileFilter: Excel file filter string.
'''     FilterIndex: Initial filter index.
'''     Title: Optional dialog title.
'''     ButtonText: Optional button text.
'''     MultiSelect: True to allow multiple selected files.
'''     DefaultValue: Optional headless fallback path or path array.
'''
''' Returns:
'''     False for cancel, a string path, or an array of string paths.
Public Function GetOpenFilename(ByVal Id As String, Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
	Dim responseSource As String
	Dim displayValue As String
	ValidateDialogId Id, "XlflowUI.GetOpenFilename"
	If XlflowRuntime.IsHeadless() Then
		On Error GoTo HandleHeadlessFailure
		GetOpenFilename = ResolveFileDialogResponse("get-open", Id, MultiSelect, DefaultValue, responseSource, displayValue)
		EmitHeadlessUIEvent "get-open", Id, "", Title, responseSource, "", displayValue, False, ""
		Exit Function
	End If

	GetOpenFilename = Application.GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
	Exit Function

HandleHeadlessFailure:
	EmitHeadlessUIEvent "get-open", Id, "", Title, "error", "", "", False, Err.Description
	Err.Raise Err.Number, Err.Source, Err.Description
End Function

''' Opens an Office file dialog or resolves scripted file paths in headless mode.
'''
''' Args:
'''     Id: Stable dialog id used by xlflow --filedialog file-open.
'''     Title: Optional dialog title.
'''     ButtonText: Optional button text.
'''     MultiSelect: True to allow multiple selected files.
'''     DefaultValue: Optional headless fallback path or path array.
'''
''' Returns:
'''     False for cancel, a string path, or an array of string paths.
Public Function FileDialogOpen(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal DefaultValue As Variant) As Variant
	Dim responseSource As String
	Dim displayValue As String
	Dim dialog As FileDialog
	ValidateDialogId Id, "XlflowUI.FileDialogOpen"
	If XlflowRuntime.IsHeadless() Then
		On Error GoTo HandleHeadlessFailure
		FileDialogOpen = ResolveFileDialogResponse("file-open", Id, MultiSelect, DefaultValue, responseSource, displayValue)
		EmitHeadlessUIEvent "file-open", Id, "", Title, responseSource, "", displayValue, False, ""
		Exit Function
	End If

	Set dialog = Application.FileDialog(msoFileDialogOpen)
	With dialog
		If Len(Title) > 0 Then .Title = Title
		If Len(ButtonText) > 0 Then .ButtonName = ButtonText
		.AllowMultiSelect = MultiSelect
		If .Show <> -1 Then
			FileDialogOpen = False
		ElseIf MultiSelect Then
			FileDialogOpen = SelectedItemsToVariantArray(.SelectedItems)
		Else
			FileDialogOpen = CStr(.SelectedItems(1))
		End If
	End With
	Exit Function

HandleHeadlessFailure:
	EmitHeadlessUIEvent "file-open", Id, "", Title, "error", "", "", False, Err.Description
	Err.Raise Err.Number, Err.Source, Err.Description
End Function

''' Opens Excel's Save As picker or resolves a scripted path in headless mode.
'''
''' Args:
'''     Id: Stable dialog id used by xlflow --filedialog save-as.
'''     InitialFileName: Suggested file name for interactive Excel.
'''     FileFilter: Excel file filter string.
'''     FilterIndex: Initial filter index.
'''     Title: Optional dialog title.
'''     ButtonText: Optional button text.
'''     DefaultValue: Optional headless fallback path or False for cancel.
'''
''' Returns:
'''     False for cancel or the selected file path.
Public Function GetSaveAsFilename(ByVal Id As String, Optional ByVal InitialFileName As String = "", Optional ByVal FileFilter As String = "", Optional ByVal FilterIndex As Long = 1, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal DefaultValue As Variant) As Variant
	Dim responseSource As String
	Dim displayValue As String
	ValidateDialogId Id, "XlflowUI.GetSaveAsFilename"
	If XlflowRuntime.IsHeadless() Then
		On Error GoTo HandleHeadlessFailure
		GetSaveAsFilename = ResolveFileDialogResponse("save-as", Id, False, DefaultValue, responseSource, displayValue)
		EmitHeadlessUIEvent "save-as", Id, "", Title, responseSource, "", displayValue, False, ""
		Exit Function
	End If

	GetSaveAsFilename = Application.GetSaveAsFilename(InitialFileName, FileFilter, FilterIndex, Title, ButtonText)
	Exit Function

HandleHeadlessFailure:
	EmitHeadlessUIEvent "save-as", Id, "", Title, "error", "", "", False, Err.Description
	Err.Raise Err.Number, Err.Source, Err.Description
End Function

''' Opens a folder picker or resolves a scripted folder path in headless mode.
'''
''' Args:
'''     Id: Stable dialog id used by xlflow --filedialog folder.
'''     Title: Optional dialog title.
'''     ButtonText: Optional button text.
'''     InitialPath: Initial folder for interactive Excel.
'''     DefaultValue: Optional headless fallback path or False for cancel.
'''
''' Returns:
'''     False for cancel or the selected folder path.
Public Function FolderPicker(ByVal Id As String, Optional ByVal Title As String = "", Optional ByVal ButtonText As String = "", Optional ByVal InitialPath As String = "", Optional ByVal DefaultValue As Variant) As Variant
	Dim responseSource As String
	Dim displayValue As String
	Dim dialog As FileDialog
	ValidateDialogId Id, "XlflowUI.FolderPicker"
	If XlflowRuntime.IsHeadless() Then
		On Error GoTo HandleHeadlessFailure
		FolderPicker = ResolveFileDialogResponse("folder", Id, False, DefaultValue, responseSource, displayValue)
		EmitHeadlessUIEvent "folder", Id, "", Title, responseSource, "", displayValue, False, ""
		Exit Function
	End If

	Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
	With dialog
		If Len(Title) > 0 Then .Title = Title
		If Len(ButtonText) > 0 Then .ButtonName = ButtonText
		If Len(InitialPath) > 0 Then .InitialFileName = InitialPath
		If .Show <> -1 Then
			FolderPicker = False
		Else
			FolderPicker = CStr(.SelectedItems(1))
		End If
	End With
	Exit Function

HandleHeadlessFailure:
	EmitHeadlessUIEvent "folder", Id, "", Title, "error", "", "", False, Err.Description
	Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Resolves a headless MsgBox response from xlflow markers or workbook defaults.
Private Function ResolveMsgBoxResponse(ByVal Id As String, Optional ByVal DefaultResponse As String = "", Optional ByRef ResponseSource As String = "", Optional ByRef ResponseToken As String = "") As VbMsgBoxResult
	Dim response As String
	On Error GoTo UseDefault
	response = LCase$(Trim$(ReadResponseValue("msgbox", Id)))
	ResponseSource = "scripted"
	GoTo Resolve

UseDefault:
	If Len(Trim$(DefaultResponse)) = 0 Then
		ResponseSource = "error"
		Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Missing scripted MsgBox response for dialog id '" & Id & "'."
	End If
	response = LCase$(Trim$(DefaultResponse))
	ResponseSource = "default"

Resolve:
	ResponseToken = response

	Select Case response
		Case "abort"
			ResolveMsgBoxResponse = vbAbort
		Case "cancel"
			ResolveMsgBoxResponse = vbCancel
		Case "ignore"
			ResolveMsgBoxResponse = vbIgnore
		Case "no"
			ResolveMsgBoxResponse = vbNo
		Case "ok"
			ResolveMsgBoxResponse = vbOK
		Case "retry"
			ResolveMsgBoxResponse = vbRetry
		Case "yes"
			ResolveMsgBoxResponse = vbYes
		Case Else
			If ResponseSource = "default" Then
				Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Invalid default MsgBox response for dialog id '" & Id & "'."
			End If
			ResponseSource = "error"
			Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Invalid scripted MsgBox response for dialog id '" & Id & "'."
	End Select
End Function

Private Function ResolveInputResponse(ByVal Id As String, Optional ByVal DefaultValue As String = "", Optional ByRef ResponseSource As String = "") As String
	On Error GoTo UseDefault
	ResponseSource = "scripted"
	ResolveInputResponse = ReadResponseValue("input", Id)
	Exit Function

UseDefault:
	If Len(DefaultValue) = 0 Then
		ResponseSource = "error"
		Err.Raise xlflowErrMissingInputResponse, "XlflowUI.InputBox", "Missing scripted InputBox response for dialog id '" & Id & "'."
	End If
	ResponseSource = "default"
	ResolveInputResponse = DefaultValue
End Function

' Resolves file dialog responses from injected markers.
' Cancel is represented as False, and multi-select responses return a Variant string array.
Private Function ResolveFileDialogResponse(ByVal Kind As String, ByVal Id As String, ByVal MultiSelect As Boolean, Optional ByVal DefaultValue As Variant, Optional ByRef ResponseSource As String = "", Optional ByRef DisplayValue As String = "") As Variant
	Dim markerValue As String

	On Error GoTo UseDefault
	markerValue = ReadFileDialogResponseValue(Kind, Id)
	On Error GoTo InvalidScripted
	ResponseSource = "scripted"
	ResolveFileDialogResponse = NormalizeFileDialogResponse(Kind, Id, MultiSelect, markerValue, DisplayValue)
	Exit Function

UseDefault:
	If IsEmpty(DefaultValue) Then
		ResponseSource = "error"
		Err.Raise xlflowErrMissingFileDialogResponse, FileDialogSourceName(Kind), "Missing scripted file dialog response for dialog id '" & Id & "'."
	End If
	ResponseSource = "default"
	ResolveFileDialogResponse = NormalizeFileDialogDefault(Kind, Id, MultiSelect, DefaultValue, DisplayValue)
	Exit Function

InvalidScripted:
	ResponseSource = "error"
	Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Converts injected newline-delimited marker values into the same Variant shape Excel would normally return.
Private Function NormalizeFileDialogResponse(ByVal Kind As String, ByVal Id As String, ByVal MultiSelect As Boolean, ByVal MarkerValue As String, Optional ByRef DisplayValue As String = "") As Variant
	Dim values As Variant
	Dim count As Long

	If LCase$(Trim$(MarkerValue)) = xlflowFileDialogCancelToken Then
		DisplayValue = xlflowFileDialogCancelToken
		NormalizeFileDialogResponse = False
		Exit Function
	End If

	values = SplitFileDialogValues(MarkerValue)
	count = FileDialogValueCount(values)
	If MultiSelect Then
		DisplayValue = JoinStringArray(values, " | ")
		NormalizeFileDialogResponse = values
		Exit Function
	End If
	If count <> 1 Then
		Err.Raise xlflowErrInvalidFileDialogResponse, FileDialogSourceName(Kind), "File dialog '" & Id & "' expected one scripted path but received " & CStr(count) & "."
	End If
	DisplayValue = CStr(values(LBound(values)))
	NormalizeFileDialogResponse = CStr(values(LBound(values)))
End Function

Private Function NormalizeFileDialogDefault(ByVal Kind As String, ByVal Id As String, ByVal MultiSelect As Boolean, ByVal DefaultValue As Variant, Optional ByRef DisplayValue As String = "") As Variant
	Dim values As Variant
	Dim count As Long

	If VarType(DefaultValue) = vbBoolean Then
		If CBool(DefaultValue) = False Then
			DisplayValue = xlflowFileDialogCancelToken
			NormalizeFileDialogDefault = False
			Exit Function
		End If
	End If

	If IsArray(DefaultValue) Then
		values = VariantArrayToStringArray(DefaultValue)
		count = FileDialogValueCount(values)
		If MultiSelect Then
			DisplayValue = JoinStringArray(values, " | ")
			NormalizeFileDialogDefault = values
			Exit Function
		End If
		If count <> 1 Then
			Err.Raise xlflowErrInvalidFileDialogResponse, FileDialogSourceName(Kind), "File dialog '" & Id & "' default value must contain exactly one path."
		End If
		DisplayValue = CStr(values(LBound(values)))
		NormalizeFileDialogDefault = CStr(values(LBound(values)))
		Exit Function
	End If

	If MultiSelect Then
		values = Array(CStr(DefaultValue))
		DisplayValue = JoinStringArray(values, " | ")
		NormalizeFileDialogDefault = values
		Exit Function
	End If

	DisplayValue = CStr(DefaultValue)
	NormalizeFileDialogDefault = DisplayValue
End Function

' Reads the workbook-defined name that xlflow injected for one file dialog wrapper call.
Private Function ReadFileDialogResponseValue(ByVal Kind As String, ByVal Id As String) As String
	On Error GoTo Missing
	ReadFileDialogResponseValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(BuildFileDialogResponseName(Kind, Id)).RefersTo)
	Exit Function

Missing:
	Err.Raise xlflowErrMissingFileDialogResponse, FileDialogSourceName(Kind), "Missing scripted file dialog response for dialog id '" & Id & "'."
End Function

Private Function BuildFileDialogResponseName(ByVal Kind As String, ByVal Id As String) As String
	Dim normalizedKind As String

	normalizedKind = LCase$(Trim$(Kind))
	Select Case normalizedKind
		Case "get-open"
			BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_GET_OPEN_" & NormalizeResponseId(Id) & "__"
		Case "file-open"
			BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_FILE_OPEN_" & NormalizeResponseId(Id) & "__"
		Case "save-as"
			BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_SAVE_AS_" & NormalizeResponseId(Id) & "__"
		Case "folder"
			BuildFileDialogResponseName = "__XLFLOW_UI_FILEDIALOG_FOLDER_" & NormalizeResponseId(Id) & "__"
		Case Else
			Err.Raise xlflowErrInvalidFileDialogResponse, FileDialogSourceName(Kind), "Unsupported file dialog kind '" & Kind & "'."
	End Select
End Function

Private Function FileDialogSourceName(ByVal Kind As String) As String
	Select Case LCase$(Trim$(Kind))
		Case "get-open"
			FileDialogSourceName = "XlflowUI.GetOpenFilename"
		Case "file-open"
			FileDialogSourceName = "XlflowUI.FileDialogOpen"
		Case "save-as"
			FileDialogSourceName = "XlflowUI.GetSaveAsFilename"
		Case "folder"
			FileDialogSourceName = "XlflowUI.FolderPicker"
		Case Else
			FileDialogSourceName = "XlflowUI.FileDialog"
	End Select
End Function

' Headless multi-select responses are stored as newline-delimited paths.
Private Function SplitFileDialogValues(ByVal MarkerValue As String) As Variant
	Dim normalizedValue As String
	Dim values As Variant
	Dim i As Long

	normalizedValue = Replace$(MarkerValue, vbCrLf, vbLf)
	normalizedValue = Replace$(normalizedValue, vbCr, vbLf)
	If Len(normalizedValue) = 0 Then
		Err.Raise xlflowErrInvalidFileDialogResponse, "XlflowUI.FileDialog", "Scripted file dialog response must contain at least one path."
	End If
	values = Split(normalizedValue, vbLf)
	For i = LBound(values) To UBound(values)
		If Len(CStr(values(i))) = 0 Then
			Err.Raise xlflowErrInvalidFileDialogResponse, "XlflowUI.FileDialog", "Scripted file dialog response contains an empty path entry."
		End If
	Next i
	SplitFileDialogValues = values
End Function

Private Function VariantArrayToStringArray(ByVal Values As Variant) As Variant
	Dim result() As String
	Dim i As Long

	ReDim result(LBound(Values) To UBound(Values))
	For i = LBound(Values) To UBound(Values)
		result(i) = CStr(Values(i))
		If Len(result(i)) = 0 Then
			Err.Raise xlflowErrInvalidFileDialogResponse, "XlflowUI.FileDialog", "File dialog default values cannot contain empty paths."
		End If
	Next i
	VariantArrayToStringArray = result
End Function

Private Function FileDialogValueCount(ByVal Values As Variant) As Long
	FileDialogValueCount = UBound(Values) - LBound(Values) + 1
End Function

Private Function JoinStringArray(ByVal Values As Variant, ByVal Separator As String) As String
	Dim i As Long
	For i = LBound(Values) To UBound(Values)
		If i > LBound(Values) Then
			JoinStringArray = JoinStringArray & Separator
		End If
		JoinStringArray = JoinStringArray & CStr(Values(i))
	Next i
End Function

Private Function SelectedItemsToVariantArray(ByVal SelectedItems As Office.FileDialogSelectedItems) As Variant
	Dim values() As String
	Dim i As Long

	ReDim values(0 To SelectedItems.Count - 1)
	For i = 1 To SelectedItems.Count
		values(i - 1) = CStr(SelectedItems(i))
	Next i
	SelectedItemsToVariantArray = values
End Function

Private Sub EmitHeadlessUIEvent(ByVal Kind As String, ByVal Id As String, ByVal Prompt As String, ByVal Title As String, ByVal ResponseSource As String, Optional ByVal ResolvedResult As String = "", Optional ByVal ResolvedValue As String = "", Optional ByVal Redacted As Boolean = False, Optional ByVal FailureMessage As String = "")
	Dim helperMacro As String
	Dim payload As String

	helperMacro = ResolveStreamHelperMacro()
	If Len(helperMacro) = 0 Then
		Exit Sub
	End If

	payload = "{" & _
		JsonProperty("event", "ui_dialog") & "," & _
		JsonProperty("kind", Kind) & "," & _
		JsonProperty("dialog_id", Id) & "," & _
		JsonProperty("prompt", Prompt) & "," & _
		JsonProperty("title", Title) & "," & _
		JsonProperty("response_source", ResponseSource) & "," & _
		JsonProperty("resolved_result", ResolvedResult) & "," & _
		JsonProperty("resolved_value", ResolvedValue) & "," & _
		JsonBooleanProperty("redacted", Redacted) & "," & _
		JsonProperty("runtime_mode", XlflowRuntime.ModeName()) & "," & _
		JsonProperty("error", FailureMessage) & "}"

	InvokeStreamHelper helperMacro, payload
End Sub

Private Sub InvokeStreamHelper(ByVal helperMacro As String, ByVal payload As String)
	On Error GoTo Swallow
	Application.Run "'" & ThisWorkbook.Name & "'!" & helperMacro, payload
	Exit Sub

Swallow:
End Sub

Private Function ResolveStreamHelperMacro() As String
	ResolveStreamHelperMacro = ReadOptionalDefinedNameValue(xlflowStreamHelperName)
End Function

Private Function ShouldRedactInputStream() As Boolean
	Dim configured As String

	configured = LCase$(Trim$(ReadOptionalDefinedNameValue(xlflowStreamRedactInputName)))
	ShouldRedactInputStream = (configured <> "false")
End Function

Private Function ReadOptionalDefinedNameValue(ByVal name As String) As String
	On Error GoTo Missing
	ReadOptionalDefinedNameValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(name).RefersTo)
	Exit Function

Missing:
	ReadOptionalDefinedNameValue = ""
End Function

Private Function JsonEscape(ByVal value As String) As String
	value = Replace$(value, "\", "\\")
	value = Replace$(value, Chr$(34), Chr$(92) & Chr$(34))
	value = Replace$(value, vbCrLf, "\n")
	value = Replace$(value, vbCr, "\n")
	value = Replace$(value, vbLf, "\n")
	value = Replace$(value, vbTab, "\t")
	JsonEscape = value
End Function

Private Function JsonProperty(ByVal name As String, ByVal value As String) As String
	JsonProperty = Chr$(34) & JsonEscape(name) & Chr$(34) & ":" & Chr$(34) & JsonEscape(value) & Chr$(34)
End Function

Private Function JsonBooleanProperty(ByVal name As String, ByVal value As Boolean) As String
	JsonBooleanProperty = Chr$(34) & JsonEscape(name) & Chr$(34) & ":" & LCase$(CStr(value))
End Function

Private Function ReadResponseValue(ByVal Kind As String, ByVal Id As String) As String
	On Error GoTo Missing
	ReadResponseValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(BuildResponseName(Kind, Id)).RefersTo)
	Exit Function

Missing:
	If LCase$(Trim$(Kind)) = "msgbox" Then
		Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Missing scripted MsgBox response for dialog id '" & Id & "'."
	End If
	Err.Raise xlflowErrMissingInputResponse, "XlflowUI.InputBox", "Missing scripted InputBox response for dialog id '" & Id & "'."
End Function

Private Function BuildResponseName(ByVal Kind As String, ByVal Id As String) As String
	BuildResponseName = "__XLFLOW_UI_" & UCase$(Trim$(Kind)) & "_" & NormalizeResponseId(Id) & "__"
End Function

Private Function NormalizeResponseId(ByVal value As String) As String
	Dim i As Long
	Dim ch As String
	Dim normalized As String
	Dim lastSeparator As Boolean

	value = LCase$(Trim$(value))
	For i = 1 To Len(value)
		ch = Mid$(value, i, 1)
		If ch Like "[a-z0-9]" Then
			normalized = normalized & ch
			lastSeparator = False
		ElseIf Len(normalized) > 0 And Not lastSeparator Then
			normalized = normalized & "_"
			lastSeparator = True
		End If
	Next i

	Do While Len(normalized) > 0 And Right$(normalized, 1) = "_"
		normalized = Left$(normalized, Len(normalized) - 1)
	Loop

	NormalizeResponseId = normalized
End Function

Private Sub ValidateDialogId(ByVal Id As String, ByVal SourceName As String)
	If Len(NormalizeResponseId(Id)) = 0 Then
		Err.Raise xlflowErrInvalidDialogId, SourceName, "Dialog id must contain at least one ASCII letter or digit."
	End If
End Sub

Private Function DecodeWorkbookDefinedName(ByVal refersTo As String) As String
	If Len(refersTo) = 0 Then
		DecodeWorkbookDefinedName = ""
		Exit Function
	End If
	If Left$(refersTo, 1) = "=" Then
		refersTo = Mid$(refersTo, 2)
	End If
	If Len(refersTo) >= 2 Then
		If Left$(refersTo, 1) = Chr$(34) And Right$(refersTo, 1) = Chr$(34) Then
			refersTo = Mid$(refersTo, 2, Len(refersTo) - 2)
		End If
	End If
	DecodeWorkbookDefinedName = Replace$(refersTo, Chr$(34) & Chr$(34), Chr$(34))
End Function
`

const defaultMainModule = `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
  App.RunCore ThisWorkbook
End Sub
`

const defaultAppModule = `Attribute VB_Name = "App"
Option Explicit

''' Runs the workbook automation entry point.
'''
''' Args:
'''     wb: Workbook that xlflow passes from Main.Run.
Public Sub RunCore(ByVal wb As Workbook)
  ' Put workbook automation here.
End Sub
`

const defaultUiModule = `Attribute VB_Name = "Ui"
Option Explicit

Public Sub RunFromButton()
  Main.Run
End Sub
`

const defaultDocumentModule = `Option Explicit
`

const defaultDebugRuntimeModule = `Attribute VB_Name = "XlflowDebug"
Option Explicit

' XlflowDebug mirrors workbook-side debug output to the terminal during xlflow runs.
' Use XlflowDebug.Log instead of raw Debug.Print when terminal-visible logs are desired.
Private Const xlflowDebugPipeName As String = "__XLFLOW_DEBUG_PIPE__"

#If VBA7 Then
	Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
	Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
	Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
	Private Const xlflowInvalidHandleValue As LongPtr = -1
#Else
	Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
	Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
	Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
	Private Const xlflowInvalidHandleValue As Long = -1
#End If

Private Const xlflowGenericWrite As Long = &H40000000
Private Const xlflowOpenExisting As Long = 3

''' Writes workbook-side debug output to Debug.Print and xlflow run/test output when available.
'''
''' Args:
'''     Parts: Values to stringify and join with spaces.
Public Sub Log(ParamArray Parts() As Variant)
	Dim index As Long
	Dim lowerBound As Long
	Dim upperBound As Long
	Dim message As String
	Dim errorNumber As Long
	Dim errorSource As String
	Dim errorDescription As String

	On Error GoTo EmptyParts
	lowerBound = LBound(Parts)
	upperBound = UBound(Parts)
	On Error GoTo 0

	For index = lowerBound To upperBound
		If index > lowerBound Then
			message = message & " "
		End If
		message = message & StringifyValue(Parts(index))
	Next index

GoTo PrintMessage

EmptyParts:
	errorNumber = Err.Number
	errorSource = Err.Source
	errorDescription = Err.Description
	Err.Clear
	On Error GoTo 0
	If errorNumber <> 9 Then
		Err.Raise errorNumber, errorSource, errorDescription
	End If

PrintMessage:
	If Len(message) = 0 Then
		Debug.Print
	Else
		Debug.Print message
	End If

	EmitDebugEvent message
End Sub

Private Sub EmitDebugEvent(ByVal Message As String)
	Dim pipeName As String
	Dim payload As String

	pipeName = ResolveDebugPipeName()
	If Len(pipeName) = 0 Then
		Exit Sub
	End If

	payload = "{" & _
		JsonProperty("event", "debug_log") & "," & _
		JsonProperty("message", Message) & "," & _
		JsonProperty("runtime_mode", XlflowRuntime.ModeName()) & "," & _
		JsonProperty("source", "XlflowDebug.Log") & "}"

	SendPipeText pipeName, payload & vbLf
End Sub

Private Function StringifyValue(ByVal Value As Variant) As String
	If IsObject(Value) Then
		On Error GoTo ObjectFallback
		StringifyValue = "[Object " & TypeName(Value) & "]"
		On Error GoTo 0
		Exit Function
ObjectFallback:
		Err.Clear
		On Error GoTo 0
		StringifyValue = "[Object]"
		Exit Function
	End If

	If IsArray(Value) Then
		StringifyValue = "[Array]"
		Exit Function
	End If

	If IsEmpty(Value) Then
		StringifyValue = "[Empty]"
		Exit Function
	End If

	If IsNull(Value) Then
		StringifyValue = "[Null]"
		Exit Function
	End If

	Select Case VarType(Value)
		Case vbBoolean
			If CBool(Value) Then
				StringifyValue = "True"
			Else
				StringifyValue = "False"
			End If
		Case vbError
			StringifyValue = "[Error " & CStr(CLng(Value)) & "]"
		Case Else
			On Error GoTo UnsupportedValue
			StringifyValue = CStr(Value)
			On Error GoTo 0
			Exit Function
UnsupportedValue:
			Err.Clear
			On Error GoTo 0
			StringifyValue = "[Unsupported Variant]"
	End Select
End Function

Private Function ResolveDebugPipeName() As String
	ResolveDebugPipeName = ReadOptionalDefinedNameValue(xlflowDebugPipeName)
End Function

Private Function ReadOptionalDefinedNameValue(ByVal Name As String) As String
	On Error GoTo Missing
	ReadOptionalDefinedNameValue = DecodeWorkbookDefinedName(ThisWorkbook.Names(Name).RefersTo)
	Exit Function

Missing:
	ReadOptionalDefinedNameValue = ""
End Function

Private Function DecodeWorkbookDefinedName(ByVal RefersTo As String) As String
	If Len(RefersTo) = 0 Then
		DecodeWorkbookDefinedName = ""
		Exit Function
	End If
	If Left$(RefersTo, 1) = "=" Then
		RefersTo = Mid$(RefersTo, 2)
	End If
	If Len(RefersTo) >= 2 Then
		If Left$(RefersTo, 1) = Chr$(34) And Right$(RefersTo, 1) = Chr$(34) Then
			RefersTo = Mid$(RefersTo, 2, Len(RefersTo) - 2)
		End If
	End If
	DecodeWorkbookDefinedName = Replace$(RefersTo, Chr$(34) & Chr$(34), Chr$(34))
End Function

Private Function JsonEscape(ByVal Value As String) As String
	Value = Replace$(Value, "\", "\\")
	Value = Replace$(Value, Chr$(34), Chr$(92) & Chr$(34))
	Value = Replace$(Value, vbCrLf, "\n")
	Value = Replace$(Value, vbCr, "\n")
	Value = Replace$(Value, vbLf, "\n")
	Value = Replace$(Value, vbTab, "\t")
	JsonEscape = Value
End Function

Private Function JsonProperty(ByVal Name As String, ByVal Value As String) As String
	JsonProperty = Chr$(34) & JsonEscape(Name) & Chr$(34) & ":" & Chr$(34) & JsonEscape(Value) & Chr$(34)
End Function

Private Sub SendPipeText(ByVal PipeName As String, ByVal Payload As String)
	Dim bytesWritten As Long
#If VBA7 Then
	Dim pipeHandle As LongPtr
#Else
	Dim pipeHandle As Long
#End If

	pipeHandle = CreateFileW(StrPtr(PipeName), xlflowGenericWrite, 0, 0, xlflowOpenExisting, 0, 0)
	If pipeHandle = xlflowInvalidHandleValue Then
		Exit Sub
	End If

	On Error GoTo Cleanup
	Call WriteFile(pipeHandle, StrPtr(Payload), Len(Payload) * 2, bytesWritten, 0)

Cleanup:
	If pipeHandle <> xlflowInvalidHandleValue Then
		Call CloseHandle(pipeHandle)
	End If
End Sub
`

const sampleTestModule = `Attribute VB_Name = "SampleTests"
Option Explicit

' xlflow tests are public parameterless Sub procedures whose names match
' Test* or *_Test.  Parameterized tests use ByVal scalar arguments plus
' @TestCase(...) comments.
'
' Use XlflowAssert helpers to raise clear, JSON-friendly failures.
'
' Useful commands:
'   xlflow test
'   xlflow test --json
'   xlflow test --fail-fast
'   xlflow test --max-failures 3
'   xlflow test --rerun-failed 1
'
' Optional hooks named BeforeAll / AfterAll / BeforeEach / AfterEach run
' around tests in this module.  Keep tests independent; use hooks only
' when setup or cleanup is actually needed.

'@Tag("smoke")
Public Sub Test_Sample_Pass()
    XlflowAssert.AssertEquals 2, 1 + 1, "basic arithmetic should work"
    XlflowAssert.AssertTrue Len("xlflow") > 0, "strings should have length"
End Sub

'@TestCase("adds positives"; 1, 2, 3)
'@TestCase("adds negatives"; -1, -2, -3)
Public Sub Test_Adds_Numbers(ByVal leftValue As Long, ByVal rightValue As Long, ByVal expected As Long)
    XlflowAssert.AssertEquals expected, leftValue + rightValue, "sum should match"
End Sub

'@ExpectedError(5)
Public Sub Test_Expected_Error()
    Err.Raise 5, "SampleTests", "Invalid procedure call or argument"
End Sub

'@Todo("not implemented yet")
Public Sub Test_Sample_Todo()
    ' Keep planned tests visible without executing them.
End Sub
`

// GenerateTestModuleResult is returned by GenerateTestModule.
type GenerateTestModuleResult struct {
	Path    string `json:"path"`
	Created bool   `json:"created"`
}

// GenerateTestModule creates a new test module file at the configured modules directory.
// The moduleName should not include the .bas extension.
func GenerateTestModule(cwd, moduleName string, src config.SourceConfig) (GenerateTestModuleResult, error) {
	var result GenerateTestModuleResult
	if moduleName == "" {
		return result, errors.New("module name is required")
	}
	cleanName := strings.TrimSpace(moduleName)
	cleanName = strings.TrimSuffix(cleanName, ".bas")
	if cleanName == "" {
		return result, errors.New("module name is empty after cleaning")
	}
	if cleanName != filepath.Base(cleanName) ||
		strings.Contains(cleanName, "..") ||
		strings.Contains(cleanName, "/") ||
		strings.Contains(cleanName, `\`) {
		return result, errors.New("invalid module name")
	}

	modulesRoot := src.Modules
	if modulesRoot == "" {
		modulesRoot = filepath.Join("src", "modules", "Tests")
	}
	moduleDir := filepath.Join(cwd, filepath.Clean(modulesRoot))
	destPath := filepath.Join(moduleDir, cleanName+".bas")

	if _, err := os.Stat(destPath); err == nil {
		return result, fmt.Errorf("test module already exists: %s", destPath)
	}

	body := fmt.Sprintf(`Attribute VB_Name = "%s"
Option Explicit

Public Sub BeforeAll()
End Sub

Public Sub AfterAll()
End Sub

Public Sub BeforeEach()
End Sub

Public Sub AfterEach()
End Sub

Public Sub Test_Sample()
    XlflowAssert.AssertTrue True, "replace with real assertions"
End Sub
`, cleanName)

	if err := os.MkdirAll(moduleDir, 0755); err != nil {
		return result, fmt.Errorf("failed to create modules directory: %w", err)
	}
	if err := os.WriteFile(destPath, []byte(body), 0644); err != nil {
		return result, fmt.Errorf("failed to write test module: %w", err)
	}

	result.Path = destPath
	result.Created = true
	return result, nil
}
