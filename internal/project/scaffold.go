package project

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

type InitResult struct {
	ConfigPath string   `json:"config_path"`
	Workbook   string   `json:"workbook"`
	Created    []string `json:"created"`
}

type InstallModulesResult struct {
	Created []string `json:"created"`
}

type bundledModuleTemplate struct {
	fileName string
	body     string
}

type WorkbookCreator func(path string) error

func Init(cwd, workbookPath string) (InitResult, error) {
	var result InitResult
	if workbookPath == "" {
		return result, errors.New("workbook path is required")
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
	}, "frm", false, false)
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
	protectedPaths = append(protectedPaths, bundledModulePaths(moduleDir, scaffoldBaseModuleTemplates())...)
	if scaffoldRuntimeHelper {
		protectedPaths = append(protectedPaths, bundledModulePaths(moduleDir, scaffoldRuntimeHelperModuleTemplates())...)
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

func installBundledModules(cwd, moduleDir string, templates []bundledModuleTemplate) (InstallModulesResult, error) {
	var result InstallModulesResult
	paths := bundledModulePaths(moduleDir, templates)
	for _, path := range paths {
		if _, err := os.Stat(path); err == nil {
			return result, fmt.Errorf("refusing to overwrite existing file: %s", path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return result, err
		}
	}
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		return result, err
	}
	for index, template := range templates {
		path := paths[index]
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

func installHelperModuleTemplates() []bundledModuleTemplate {
	return []bundledModuleTemplate{
		{fileName: "XlflowAssert.bas", body: defaultAssertModule},
		{fileName: "XlflowRuntime.bas", body: defaultRuntimeModule},
		{fileName: "XlflowUI.bas", body: defaultUIRuntimeModule},
		{fileName: "XlflowDebug.bas", body: defaultDebugRuntimeModule},
	}
}

func scaffoldRuntimeHelperModuleTemplates() []bundledModuleTemplate {
	return []bundledModuleTemplate{
		{fileName: "XlflowRuntime.bas", body: defaultRuntimeModule},
		{fileName: "XlflowUI.bas", body: defaultUIRuntimeModule},
		{fileName: "XlflowDebug.bas", body: defaultDebugRuntimeModule},
	}
}

func scaffoldBaseModuleTemplates() []bundledModuleTemplate {
	return []bundledModuleTemplate{
		{fileName: "XlflowAssert.bas", body: defaultAssertModule},
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
	name = strings.TrimSpace(name)
	if name == "" {
		return "Book.xlsm", nil
	}
	name = filepath.Base(name)
	ext := strings.ToLower(filepath.Ext(name))
	if ext == "" {
		return name + ".xlsm", nil
	}
	if ext != ".xlsm" {
		return "", fmt.Errorf("workbook name must use .xlsm extension: %s", name)
	}
	return name, nil
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

' Minimal assertion helpers for workbook-side tests.
' Keep assertions scalar so failures stay easy to read from xlflow JSON and terminal output.
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
    RaiseAssertFailure message, "expected <" & DescribeAssertValue(expected) & "> but got <" & DescribeAssertValue(actual) & ">", "XlflowAssert.AssertEquals"
  End If

  If expected <> actual Then
    RaiseAssertFailure message, "expected <" & DescribeAssertValue(expected) & "> but got <" & DescribeAssertValue(actual) & ">", "XlflowAssert.AssertEquals"
  End If
End Sub

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

  If expected = actual Then
    RaiseAssertFailure message, "expected <" & DescribeAssertValue(expected) & "> to differ from <" & DescribeAssertValue(actual) & ">", "XlflowAssert.AssertNotEqual"
  End If
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
  If Not condition Then
    RaiseAssertFailure message, "expected True but got False", "XlflowAssert.AssertTrue"
  End If
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
  If condition Then
    RaiseAssertFailure message, "expected False but got True", "XlflowAssert.AssertFalse"
  End If
End Sub

Public Sub AssertFail(Optional ByVal message As String = "")
  RaiseAssertFailure message, "assertion failed", "XlflowAssert.AssertFail"
End Sub

Public Sub AssertInconclusive(Optional ByVal message As String = "")
  Dim detail As String
  detail = "inconclusive"
  If Len(message) > 0 Then
    detail = message
  End If
  Err.Raise vbObjectError + 516, "XlflowAssert.AssertInconclusive", detail
End Sub

Public Sub AssertIsNothing(ByVal value As Variant, Optional ByVal message As String = "")
  If Not IsObject(value) Then
    RaiseAssertFailure message, "expected an object but got a non-object", "XlflowAssert.AssertIsNothing"
    Exit Sub
  End If
  If Not value Is Nothing Then
    RaiseAssertFailure message, "expected Nothing but got an object reference", "XlflowAssert.AssertIsNothing"
  End If
End Sub

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
  Err.Raise vbObjectError + 513, source, detail
End Sub

Private Function DescribeAssertValue(ByVal value As Variant) As String
  If IsNull(value) Then
    DescribeAssertValue = "Null"
  ElseIf IsEmpty(value) Then
    DescribeAssertValue = "Empty"
  Else
    DescribeAssertValue = CStr(value)
  End If
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

' Returns a stable numeric mode value for lightweight branching.
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

' Returns the normalized mode name injected by xlflow.
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

' True only for normal human-driven Excel usage.
Public Function IsInteractive() As Boolean
	IsInteractive = (Mode() = xlflowInteractive)
End Function

' True for all unattended-style modes such as headless, CI, agent, and test.
Public Function IsHeadless() As Boolean
	Select Case Mode()
		Case xlflowHeadless, xlflowCI, xlflowAgent, xlflowTest
			IsHeadless = True
		Case Else
			IsHeadless = False
	End Select
End Function

Public Function IsCI() As Boolean
	IsCI = (Mode() = xlflowCI)
End Function

Public Function IsAgent() As Boolean
	IsAgent = (Mode() = xlflowAgent)
End Function

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

' Wrapper for VBA.MsgBox. In headless-like modes xlflow resolves the response from --msgbox.
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

' Wrapper for VBA.InputBox. Default is interactive-only UI text, while DefaultValue is the headless fallback.
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

' Wrapper for Application.GetOpenFilename.
' In headless mode, repeat --filedialog get-open:<id>=<path> to simulate multi-select results.
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

' Wrapper for Application.FileDialog(msoFileDialogOpen).
' In headless mode, this uses the same --filedialog file-open:<id>=<path> contract as other xlflow runs.
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

' Wrapper for Application.GetSaveAsFilename.
' Use --filedialog save-as:<id>=<path> in unattended runs, or @cancel to simulate Cancel.
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

' Wrapper for Application.FileDialog(msoFileDialogFolderPicker).
' Use --filedialog folder:<id>=<path> in unattended runs, or @cancel to simulate Cancel.
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

Public Sub Log(ParamArray Parts() As Variant)
	Dim message As String
	message = JoinLogMessage(Parts)
	If Len(message) = 0 Then
		Debug.Print
	Else
		Debug.Print message
	End If

	EmitDebugEvent message
End Sub

Private Function JoinLogMessage(ByVal Parts() As Variant) As String
	Dim index As Long
	Dim lowerBound As Long
	Dim upperBound As Long
	Dim errorNumber As Long
	Dim errorSource As String
	Dim errorDescription As String

	On Error GoTo EmptyParts
	lowerBound = LBound(Parts)
	upperBound = UBound(Parts)
	On Error GoTo 0

	For index = lowerBound To upperBound
		If index > lowerBound Then
			JoinLogMessage = JoinLogMessage & " "
		End If
		JoinLogMessage = JoinLogMessage & StringifyValue(Parts(index))
	Next index
	Exit Function

EmptyParts:
	errorNumber = Err.Number
	errorSource = Err.Source
	errorDescription = Err.Description
	Err.Clear
	On Error GoTo 0
	If errorNumber = 9 Then
		Exit Function
	End If
	Err.Raise errorNumber, errorSource, errorDescription
End Function

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
' Test* or *_Test.  xlflow discovers them automatically at run time.
'
' Use XlflowAssert helpers to raise clear, JSON-friendly failures.
'
' Tags: add '@Tag("name")' comment lines directly above a test sub
' and run only matching tests with  xlflow test --tag <name>.
'
' Hooks: BeforeAll / AfterAll / BeforeEach / AfterEach are optional
' reserved names.  They must be public parameterless Subs and they
' affect only tests in the same module.
'
' Keep tests independent.  Use BeforeEach / AfterEach for isolation
' and BeforeAll for expensive one-time setup.

Public Sub BeforeAll()
    ' Runs once before the first test in this module.
End Sub

Public Sub AfterAll()
    ' Runs once after the last test in this module.
End Sub

Public Sub BeforeEach()
    ' Runs before every test in this module.
End Sub

Public Sub AfterEach()
    ' Runs after every test in this module.
End Sub

'@Tag("smoke")
Public Sub Test_Sample_Pass()
    ' A passing test demonstrates the AssertEquals helper.
    XlflowAssert.AssertEquals 1 + 1, 2, "basic arithmetic should work"
End Sub

Public Sub Test_Sample_Inconclusive()
    ' Mark not-yet-implemented tests as inconclusive instead of
    ' commenting them out.  They appear as [?] in terminal output.
    XlflowAssert.AssertInconclusive "not implemented yet"
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
	if cleanName != filepath.Base(cleanName) || strings.Contains(cleanName, "..") {
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
