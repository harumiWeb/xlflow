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
	assertPath := filepath.Join(cwd, "src", "modules", "XlflowAssert.bas")
	runtimePath := filepath.Join(cwd, "src", "modules", "XlflowRuntime.bas")
	uiHelperPath := filepath.Join(cwd, "src", "modules", "XlflowUI.bas")
	mainPath := filepath.Join(cwd, "src", "modules", "Main.bas")
	appPath := filepath.Join(cwd, "src", "modules", "App.bas")
	uiPath := filepath.Join(cwd, "src", "modules", "Ui.bas")
	thisWorkbookPath := filepath.Join(cwd, "src", "workbook", "ThisWorkbook.bas")
	sheet1Path := filepath.Join(cwd, "src", "workbook", "Sheet1.bas")
	protectedPaths := []string{destPath, configPath, assertPath, mainPath, appPath, uiPath}
	if scaffoldRuntimeHelper {
		protectedPaths = append(protectedPaths, runtimePath, uiHelperPath)
	}
	if scaffoldWorkbookModules {
		protectedPaths = append(protectedPaths, thisWorkbookPath, sheet1Path)
	}
	for _, path := range protectedPaths {
		if _, err := os.Stat(path); err == nil {
			return result, fmt.Errorf("refusing to overwrite existing file: %s", path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return result, err
		}
	}

	dirs := []string{
		filepath.Join(cwd, "src", "modules"),
		filepath.Join(cwd, "src", "classes"),
		filepath.Join(cwd, "src", "forms"),
		filepath.Join(cwd, "src", "workbook"),
		filepath.Join(cwd, "tests"),
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

	if err := writeExclusive(assertPath, defaultAssertModule); err != nil {
		return result, err
	}
	result.Created = append(result.Created, filepath.ToSlash(rel(cwd, assertPath)))
	if scaffoldRuntimeHelper {
		if err := writeExclusive(runtimePath, defaultRuntimeModule); err != nil {
			return result, err
		}
		result.Created = append(result.Created, filepath.ToSlash(rel(cwd, runtimePath)))
		if err := writeExclusive(uiHelperPath, defaultUIRuntimeModule); err != nil {
			return result, err
		}
		result.Created = append(result.Created, filepath.ToSlash(rel(cwd, uiHelperPath)))
	}
	for _, item := range []struct {
		path string
		body string
	}{
		{mainPath, defaultMainModule},
		{appPath, defaultAppModule},
		{uiPath, defaultUiModule},
	} {
		if err := writeExclusive(item.path, item.body); err != nil {
			return result, err
		}
		result.Created = append(result.Created, filepath.ToSlash(rel(cwd, item.path)))
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
    RaiseAssertEqualsFailure expected, actual, message
  End If

  If expected <> actual Then
    RaiseAssertEqualsFailure expected, actual, message
  End If
End Sub

Private Sub RaiseAssertEqualsFailure(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
  Dim detail As String
  detail = "expected <" & DescribeAssertValue(expected) & "> but got <" & DescribeAssertValue(actual) & ">"
  If Len(message) > 0 Then
    detail = message & ": " & detail
  End If
  Err.Raise vbObjectError + 513, "XlflowAssert.AssertEquals", detail
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

Private Const xlflowInteractive As Long = 0
Private Const xlflowHeadless As Long = 1
Private Const xlflowCI As Long = 2
Private Const xlflowAgent As Long = 3
Private Const xlflowTest As Long = 4

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

Public Function IsInteractive() As Boolean
	IsInteractive = (Mode() = xlflowInteractive)
End Function

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

Private Const xlflowResponseErrorBase As Long = vbObjectError + 520
Private Const xlflowErrInvalidDialogId As Long = xlflowResponseErrorBase + 1
Private Const xlflowErrInvalidMsgBoxResponse As Long = xlflowResponseErrorBase + 2
Private Const xlflowErrMissingInputResponse As Long = xlflowResponseErrorBase + 3

Public Function MsgBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal DefaultResponse As String = "") As VbMsgBoxResult
	ValidateDialogId Id, "XlflowUI.MsgBox"
	If XlflowRuntime.IsHeadless() Then
		MsgBox = ResolveMsgBoxResponse(Id, DefaultResponse)
		Exit Function
	End If

	MsgBox = VBA.Interaction.MsgBox(Prompt, Buttons, Title)
End Function

Public Function InputBox(ByVal Id As String, ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "", Optional ByVal DefaultValue As String = "") As String
	ValidateDialogId Id, "XlflowUI.InputBox"
	If XlflowRuntime.IsHeadless() Then
		InputBox = ResolveInputResponse(Id, DefaultValue)
		Exit Function
	End If

	InputBox = VBA.Interaction.InputBox(Prompt, Title, Default)
End Function

Private Function ResolveMsgBoxResponse(ByVal Id As String, Optional ByVal DefaultResponse As String = "") As VbMsgBoxResult
	Dim response As String
	On Error GoTo UseDefault
	response = LCase$(Trim$(ReadResponseValue("msgbox", Id)))
	GoTo Resolve

UseDefault:
	If Len(Trim$(DefaultResponse)) = 0 Then
		Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Missing scripted MsgBox response for dialog id '" & Id & "'."
	End If
	response = LCase$(Trim$(DefaultResponse))

Resolve:

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
			Err.Raise xlflowErrInvalidMsgBoxResponse, "XlflowUI.MsgBox", "Missing or invalid scripted MsgBox response for dialog id '" & Id & "'."
	End Select
End Function

Private Function ResolveInputResponse(ByVal Id As String, Optional ByVal DefaultValue As String = "") As String
	On Error GoTo UseDefault
	ResolveInputResponse = ReadResponseValue("input", Id)
	Exit Function

UseDefault:
	ResolveInputResponse = DefaultValue
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
