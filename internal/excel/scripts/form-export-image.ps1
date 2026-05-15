param(
  [string]$WorkbookPath,
  [string]$FormName,
  [string]$OutputPath,
  [string]$Initializer = "",
  [string]$Visible = "false",
  [string]$Overwrite = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "form export-image"
$excel = $null
$workbook = $null
$runtimeExcel = $null
$runtimeWorkbook = $null
$runtimeVBProject = $null
$runtimeWorkbookPath = ""
$sessionAttached = $false
$sessionMode = "none"
$saveState = [ordered]@{ dirty = $false; needs_save = $false }
$tempModuleName = ""
$tempModuleInstalled = $false
$tempModuleRemoved = $false
$createdParentDirs = $false
$temporaryExportPath = ""
$phase = "validate_args"

function Set-FormExportImageValidationError {
  param(
    [string]$Code,
    [string]$Message
  )

  Set-XlflowError -Result $result -Code $Code -Message $Message -Source "xlflow"
}

function Get-XlflowImageDimensions {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    return $null
  }

  $image = $null
  try {
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
    $image = [System.Drawing.Image]::FromFile($Path)
    return [ordered]@{
      width_px = [int]$image.Width
      height_px = [int]$image.Height
    }
  } catch {
    Write-Verbose ("failed to read exported form image dimensions: " + $_.Exception.Message)
    return $null
  } finally {
    if ($null -ne $image) {
      try { $image.Dispose() } catch { Write-Verbose ("failed to dispose form image handle: " + $_.Exception.Message) }
    }
  }
}

function New-XlflowFormExportImageModuleName {
  $suffix = [Guid]::NewGuid().ToString("N").Substring(0, 20)
  return "XlflowCap_" + $suffix
}

function New-XlflowFormExportImageModuleCode {
  return @'
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
#Else
Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
#End If

Private xlflowCapturedForm As Object
Private xlflowCaptureFormName As String
Private xlflowCaptureToken As String
Private xlflowCaptureInitializer As String
Private xlflowCaptureScheduledAt As Date
Private xlflowLastErrorSource As String
Private xlflowLastErrorMessage As String
Private xlflowCaptureReady As Boolean
Private xlflowCaptureWindowCaption As String
Private xlflowCaptureWindowHandle As String

Private Function XlflowFindFormWindowHandle(ByVal caption As String) As String
#If VBA7 Then
  Dim hwnd As LongPtr
#Else
  Dim hwnd As Long
#End If

  hwnd = 0
  If Len(caption) > 0 Then
    hwnd = FindWindowW(0, StrPtr(caption))
  End If
  XlflowFindFormWindowHandle = CStr(hwnd)
End Function

Public Sub XlflowPrepareFormImageCapture(ByVal formName As String, ByVal token As String, Optional ByVal initializer As String = "")
  xlflowCaptureFormName = formName
  xlflowCaptureToken = token
  xlflowCaptureInitializer = Trim$(initializer)
  xlflowLastErrorSource = ""
  xlflowLastErrorMessage = ""
  xlflowCaptureReady = False
  xlflowCaptureWindowCaption = ""
  xlflowCaptureWindowHandle = "0"
  xlflowCaptureScheduledAt = Now + TimeSerial(0, 0, 1)
  Application.OnTime xlflowCaptureScheduledAt, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!XlflowExecuteFormImageCapture"
End Sub

Public Sub XlflowExecuteFormImageCapture()
  Dim loaded As Boolean
  Dim initializerRan As Boolean
  Dim caption As String
  On Error GoTo ErrHandler

  Set xlflowCapturedForm = UserForms.Add(xlflowCaptureFormName)
  loaded = True

  If Len(xlflowCaptureInitializer) > 0 Then
    CallByName xlflowCapturedForm, xlflowCaptureInitializer, VbMethod, ThisWorkbook
    initializerRan = True
  End If

  caption = ""
  On Error Resume Next
  caption = CStr(xlflowCapturedForm.Caption)
  On Error GoTo ErrHandler
  If Len(caption) = 0 Then
    caption = xlflowCaptureFormName
  End If

  xlflowCapturedForm.Caption = caption & " [xlflow-capture-" & xlflowCaptureToken & "]"
  xlflowCapturedForm.Show vbModeless
  DoEvents
  xlflowCaptureWindowCaption = CStr(xlflowCapturedForm.Caption)
  xlflowCaptureWindowHandle = XlflowFindFormWindowHandle(xlflowCaptureWindowCaption)
  xlflowCaptureReady = True
  Exit Sub

ErrHandler:
  If Not loaded Then
    xlflowLastErrorSource = "XlflowFormImageCapture.runtime_load"
  ElseIf Len(xlflowCaptureInitializer) > 0 And Not initializerRan Then
    xlflowLastErrorSource = "XlflowFormImageCapture.initializer"
  Else
    xlflowLastErrorSource = "XlflowFormImageCapture.capture_prepare"
  End If
  xlflowLastErrorMessage = Err.Description
End Sub

Public Sub XlflowCleanupFormImageCapture()
  On Error Resume Next
  If xlflowCaptureScheduledAt <> 0 Then
    Application.OnTime xlflowCaptureScheduledAt, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!XlflowExecuteFormImageCapture", , False
  End If
  If Not xlflowCapturedForm Is Nothing Then
    Unload xlflowCapturedForm
  End If
  Set xlflowCapturedForm = Nothing
  xlflowCaptureScheduledAt = 0
  xlflowCaptureReady = False
  xlflowCaptureWindowCaption = ""
  xlflowCaptureWindowHandle = "0"
  xlflowLastErrorSource = ""
  xlflowLastErrorMessage = ""
  On Error GoTo 0
End Sub

Public Function XlflowReadFormImageCaptureStatus() As String
  XlflowReadFormImageCaptureStatus = xlflowLastErrorSource & vbTab & xlflowLastErrorMessage & vbTab & CStr(xlflowCaptureReady) & vbTab & Replace(xlflowCaptureWindowCaption, vbTab, " ") & vbTab & xlflowCaptureWindowHandle
End Function

'@
}

function New-XlflowFormExportRuntimeWorkbookCopy {
  param($SourceWorkbook)

  $extension = ".xlsm"
  $tempExcel = $null
  $tempWorkbook = $null
  try {
    $fullName = [string]$SourceWorkbook.FullName
    $candidateExtension = [System.IO.Path]::GetExtension($fullName)
    if (-not [string]::IsNullOrWhiteSpace($candidateExtension)) {
      $extension = $candidateExtension
    }
  } catch {
    Write-Verbose ("failed to resolve workbook extension for form-export-image temp copy: " + $_.Exception.Message)
  }

  $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("xlflow-form-export-image-" + [Guid]::NewGuid().ToString("N") + $extension)
  try {
    $SourceWorkbook.SaveCopyAs($tempPath)

    $tempExcel = New-Object -ComObject Excel.Application
    $tempExcel.Visible = $true
    $tempExcel.DisplayAlerts = $false
    $tempWorkbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $tempExcel -WorkbookPath $tempPath -DisplayAlerts $false -DisableAutomationMacros $false
    return [pscustomobject][ordered]@{
      excel = $tempExcel
      workbook = $tempWorkbook
      path = $tempPath
    }
  } catch {
    if ($null -ne $tempWorkbook -or $null -ne $tempExcel) {
      Close-XlflowCom -Workbook $tempWorkbook -Excel $tempExcel -Save $false
    }
    if (Test-Path -LiteralPath $tempPath) {
      Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
    throw
  }
}

function Invoke-XlflowPrepareFormImageCapture {
  param(
    $Excel,
    $Workbook,
    [string]$TargetFormName,
    [string]$Token,
    [string]$InitializerName = ""
  )

  $workbookName = ([string]$Workbook.Name).Replace("'", "''")
  $macroName = "'" + $workbookName + "'!XlflowPrepareFormImageCapture"
  $Excel.Run($macroName, $TargetFormName, $Token, $InitializerName) | Out-Null
}

function Get-XlflowFormImageCaptureStatus {
  param(
    $Excel,
    $Workbook
  )

  $workbookName = ([string]$Workbook.Name).Replace("'", "''")
  $macroName = "'" + $workbookName + "'!XlflowReadFormImageCaptureStatus"
  $value = [string]$Excel.Run($macroName)
  $parts = $value -split "`t", 5
  $ready = $false
  if ($parts.Length -ge 3) {
    $ready = [string]$parts[2] -eq "True"
  }
  $hwnd = [int64]0
  if ($parts.Length -ge 5) {
    [void][int64]::TryParse([string]$parts[4], [ref]$hwnd)
  }
  return [pscustomobject][ordered]@{
    source = $(if ($parts.Length -ge 1) { $parts[0] } else { "" })
    message = $(if ($parts.Length -ge 2) { $parts[1] } else { "" })
    ready = $ready
    caption = $(if ($parts.Length -ge 4) { $parts[3] } else { "" })
    hwnd = $hwnd
  }
}

function Invoke-XlflowCleanupFormImageCapture {
  param(
    $Excel,
    $Workbook
  )

  $workbookName = ([string]$Workbook.Name).Replace("'", "''")
  $macroName = "'" + $workbookName + "'!XlflowCleanupFormImageCapture"
  try {
    $Excel.Run($macroName) | Out-Null
  } catch {
    Write-Verbose ("failed to unload temporary captured form: " + $_.Exception.Message)
  }
}

function Get-XlflowFormExportImageErrorCode {
  param([System.Exception]$Exception)

  if ($null -eq $Exception) {
    return "form_export_image_failed"
  }
  $source = [string]$Exception.Source
  $message = [string]$Exception.Message
  if ($source -like "*XlflowFormImageCapture.initializer*" -or $message -like "*XlflowFormImageCapture.initializer*") {
    return "form_initializer_failed"
  }
  if ($source -like "*XlflowFormImageCapture.runtime_load*" -or $message -like "*XlflowFormImageCapture.runtime_load*") {
    return "runtime_form_load_failed"
  }
  if ($source -like "*XlflowFormImageCapture.capture_prepare*" -or $message -like "*XlflowFormImageCapture.capture_prepare*") {
    return "runtime_form_load_failed"
  }
  if ($source -like "*XlflowFormImageCapture.compile*" -or $message -like "*XlflowFormImageCapture.compile*") {
    return "vba_compile_failed"
  }
  if ($message -like "*window_not_found*") {
    return "window_not_found"
  }
  if ($message -like "*image_capture_failed*") {
    return "image_capture_failed"
  }
  return "form_export_image_failed"
}

function Get-XlflowWindowCaptureInfo {
  param(
    [System.IntPtr]$Hwnd
  )

  Add-XlflowNativeMethods
  if ($Hwnd -eq [System.IntPtr]::Zero) {
    return $null
  }
  if (-not [XlflowNativeMethods]::IsWindowVisible($Hwnd)) {
    return $null
  }
  $rect = New-Object XlflowNativeMethods+RECT
  if (-not [XlflowNativeMethods]::GetWindowRect($Hwnd, [ref]$rect)) {
    return $null
  }
  $width = [int]($rect.Right - $rect.Left)
  $height = [int]($rect.Bottom - $rect.Top)
  if ($width -le 0 -or $height -le 0) {
    return $null
  }
  return [pscustomobject][ordered]@{
    hwnd = [int64]$Hwnd
    title = [XlflowNativeMethods]::GetWindowTextString($Hwnd)
    class_name = [XlflowNativeMethods]::GetClassNameString($Hwnd)
    left = [int]$rect.Left
    top = [int]$rect.Top
    width = $width
    height = $height
  }
}

function Find-XlflowWindowByTitle {
  param(
    [int]$ProcessId,
    [string]$Title,
    [switch]$ExactMatch
  )

  if ([string]::IsNullOrWhiteSpace($Title)) {
    return $null
  }

  Add-XlflowNativeMethods
  $candidateSets = @()
  if ($ProcessId -gt 0) {
    $candidateSets += ,([XlflowNativeMethods]::GetWindowsForProcess([uint32]$ProcessId))
  }
  $candidateSets += ,([XlflowNativeMethods]::GetTopLevelWindows())
  foreach ($candidateSet in $candidateSets) {
    foreach ($hwnd in $candidateSet) {
      $info = Get-XlflowWindowCaptureInfo -Hwnd $hwnd
      if ($null -eq $info) {
        continue
      }
      if ($ExactMatch) {
        if ([string]$info.title -ceq $Title) {
          return $info
        }
        continue
      }
      if (-not [string]::IsNullOrWhiteSpace([string]$info.title) -and [string]$info.title.IndexOf($Title, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        return $info
      }
    }
  }
  return $null
}

function Find-XlflowWindowByCaptionToken {
  param(
    [int]$ProcessId,
    [string]$Token,
    [int]$TimeoutMilliseconds = 5000,
    [int]$PollMilliseconds = 100
  )

  $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
  while ([DateTime]::UtcNow -lt $deadline) {
    $window = Find-XlflowWindowByTitle -ProcessId $ProcessId -Title $Token
    if ($null -ne $window) {
      return $window
    }
    Start-Sleep -Milliseconds $PollMilliseconds
  }
  return $null
}

function Resolve-XlflowFormImageCaptureWindow {
  param(
    [int]$ProcessId,
    [string]$Token,
    $CaptureStatus
  )

  if ($null -ne $CaptureStatus -and [int64]$CaptureStatus.hwnd -ne 0) {
    $window = Get-XlflowWindowCaptureInfo -Hwnd ([System.IntPtr]([int64]$CaptureStatus.hwnd))
    if ($null -ne $window) {
      return $window
    }
  }
  if ($null -ne $CaptureStatus -and -not [string]::IsNullOrWhiteSpace([string]$CaptureStatus.caption)) {
    $window = Find-XlflowWindowByTitle -ProcessId $ProcessId -Title ([string]$CaptureStatus.caption) -ExactMatch
    if ($null -ne $window) {
      return $window
    }
  }
  return Find-XlflowWindowByCaptionToken -ProcessId $ProcessId -Token $Token -TimeoutMilliseconds 1 -PollMilliseconds 1
}

function Wait-XlflowFormImageCaptureWindow {
  param(
    $Excel,
    $Workbook,
    [int]$ProcessId,
    [string]$Token,
    [int]$TimeoutMilliseconds = 7000,
    [int]$PollMilliseconds = 100
  )

  $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
  while ([DateTime]::UtcNow -lt $deadline) {
    $captureStatus = Get-XlflowFormImageCaptureStatus -Excel $Excel -Workbook $Workbook
    if (-not [string]::IsNullOrWhiteSpace($captureStatus.source) -or -not [string]::IsNullOrWhiteSpace($captureStatus.message)) {
      $source = $captureStatus.source
      if ([string]::IsNullOrWhiteSpace($source)) {
        $source = "XlflowFormImageCapture.capture_prepare"
      }
      throw ($source + ": " + $captureStatus.message)
    }
    if (-not [bool]$captureStatus.ready -and [int64]$captureStatus.hwnd -eq 0 -and [string]::IsNullOrWhiteSpace([string]$captureStatus.caption)) {
      Start-Sleep -Milliseconds $PollMilliseconds
      continue
    }
    $window = Resolve-XlflowFormImageCaptureWindow -ProcessId $ProcessId -Token $Token -CaptureStatus $captureStatus
    if ($null -ne $window) {
      return [pscustomobject][ordered]@{
        status = $captureStatus
        window = $window
      }
    }
    Start-Sleep -Milliseconds $PollMilliseconds
  }
  throw "window_not_found: could not find a visible UserForm window for capture token " + $Token
}

function Set-XlflowFormExportDialogFailure {
  param(
    $Dialog,
    $Selection
  )

  $messageLines = @(Get-XlflowExcelDialogMessageLines -Dialog $Dialog)
  $message = ($messageLines -join [Environment]::NewLine)
  if ([string]::IsNullOrWhiteSpace($message)) {
    $message = "VBA dialog was shown while preparing the runtime UserForm capture."
  }

  $source = "VBA"
  $line = 0
  if ($null -ne $Selection -and $null -ne $Selection.location) {
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.location.module)) {
      $source = [string]$Selection.location.module
    }
    if ([int]$Selection.location.line -gt 0) {
      $line = [int]$Selection.location.line
    }
  }

  $code = "runtime_form_load_failed"
  $number = 0
  if ([string]$Dialog.kind -eq "compile") {
    $code = "vba_compile_failed"
  } else {
    $number = Get-XlflowVBARuntimeDialogErrorNumber -Dialog $Dialog
  }

  Set-XlflowError -Result $result -Code $code -Message $message -Source $source -Number $number -Line $line -Phase $phase
}

function Save-XlflowWindowImage {
  param(
    [int64]$Hwnd,
    [string]$Path
  )

  Add-XlflowNativeMethods
  Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

  $rect = New-Object XlflowNativeMethods+RECT
  if (-not [XlflowNativeMethods]::GetWindowRect([IntPtr]$Hwnd, [ref]$rect)) {
    throw "window_not_found: failed to resolve window bounds"
  }
  $width = [int]($rect.Right - $rect.Left)
  $height = [int]($rect.Bottom - $rect.Top)
  if ($width -le 0 -or $height -le 0) {
    throw "image_capture_failed: target window has non-positive bounds"
  }

  $bitmap = $null
  $graphics = $null
  $hdc = [IntPtr]::Zero
  try {
    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $hdc = $graphics.GetHdc()
    $printOk = [XlflowNativeMethods]::PrintWindow([IntPtr]$Hwnd, $hdc, 0)
    $graphics.ReleaseHdc($hdc)
    $hdc = [IntPtr]::Zero

    if (-not $printOk) {
      $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size($width, $height)))
    }

    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    return [ordered]@{
      width_px = $width
      height_px = $height
    }
  } catch {
    throw "image_capture_failed: " + $_.Exception.Message
  } finally {
    if ($hdc -ne [IntPtr]::Zero -and $null -ne $graphics) {
      try { $graphics.ReleaseHdc($hdc) } catch { Write-Verbose ("failed to release capture hdc: " + $_.Exception.Message) }
    }
    if ($null -ne $graphics) {
      try { $graphics.Dispose() } catch { Write-Verbose ("failed to dispose capture graphics: " + $_.Exception.Message) }
    }
    if ($null -ne $bitmap) {
      try { $bitmap.Dispose() } catch { Write-Verbose ("failed to dispose capture bitmap: " + $_.Exception.Message) }
    }
  }
}

try {
  if ([string]::IsNullOrWhiteSpace($FormName)) {
    Set-FormExportImageValidationError -Code "form_export_image_args_invalid" -Message "form name is required"
    Write-XlflowJson -Result $result
    exit
  }
  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    Set-FormExportImageValidationError -Code "form_export_image_args_invalid" -Message "-OutputPath is required."
    Write-XlflowJson -Result $result
    exit
  }

  $resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
  if (Test-Path -LiteralPath $resolvedOutputPath -PathType Container) {
    Set-FormExportImageValidationError -Code "form_export_image_args_invalid" -Message ("Output path '" + $resolvedOutputPath + "' is a directory.")
    Write-XlflowJson -Result $result
    exit
  }

  $extension = [System.IO.Path]::GetExtension($resolvedOutputPath)
  if ([string]::IsNullOrWhiteSpace($extension) -or $extension.ToLowerInvariant() -ne ".png") {
    Set-XlflowError -Result $result -Code "unsupported_image_format" -Message ("Image format '" + $extension.TrimStart(".") + "' is not supported. Supported formats: png.") -Source "xlflow"
    Write-XlflowJson -Result $result
    exit
  }

  $outputParent = Split-Path -Parent $resolvedOutputPath
  if (-not [string]::IsNullOrWhiteSpace($outputParent) -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
    $createdParentDirs = $true
  }

  if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not (ConvertTo-XlflowBool $Overwrite)) {
    Set-XlflowError -Result $result -Code "output_file_exists" -Message ("Output file '" + $resolvedOutputPath + "' already exists. Use --overwrite to replace it.") -Source "xlflow"
    Write-XlflowJson -Result $result
    exit
  }
  $exportPath = $resolvedOutputPath
  if ((Test-Path -LiteralPath $resolvedOutputPath -PathType Leaf) -and (ConvertTo-XlflowBool $Overwrite)) {
    $tempParent = $outputParent
    if ([string]::IsNullOrWhiteSpace($tempParent)) {
      $tempParent = (Get-Location).ProviderPath
    }
    $temporaryExportPath = Join-Path $tempParent ("xlflow-form-export-" + [Guid]::NewGuid().ToString("N") + ".png")
    $exportPath = $temporaryExportPath
  }

  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $phase = "open_source_workbook"
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached

  try {
    $null = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbproject_access_denied" -Message "VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center." -Source "Excel"
    throw
  }

  $userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)
  if ($FormName -notin $userFormNames) {
    Set-XlflowError -Result $result -Code "form_not_found" -Message ("UserForm '" + $FormName + "' was not found in the workbook.") -Source "xlflow"
    throw "form not found"
  }

  $runtimeOpenResult = New-XlflowFormExportRuntimeWorkbookCopy -SourceWorkbook $workbook
  $phase = "open_runtime_copy"
  $runtimeExcel = $runtimeOpenResult.excel
  $runtimeWorkbook = $runtimeOpenResult.workbook
  $runtimeWorkbookPath = $runtimeOpenResult.path
  $runtimeVBProject = $runtimeWorkbook.VBProject
  $tempModuleName = New-XlflowFormExportImageModuleName
  $phase = "install_helper_module"
  $null = Install-XlflowVBComponentFromCode -VBProject $runtimeVBProject -Name $tempModuleName -Code (New-XlflowFormExportImageModuleCode)
  $tempModuleInstalled = $true

  $processId = Get-XlflowExcelProcessId -Excel $runtimeExcel
  $captureToken = "xlflow-capture-" + [Guid]::NewGuid().ToString("N")
  $phase = "schedule_form_capture"
  $captureResult = Invoke-XlflowExcelCallWithDialogWatch -Excel $runtimeExcel -Workbook $runtimeWorkbook -DialogKind "any_vba" -CaptureDialogs $true -WaitMilliseconds 250 -Invocation {
    Invoke-XlflowPrepareFormImageCapture -Excel $runtimeExcel -Workbook $runtimeWorkbook -TargetFormName $FormName -Token $captureToken -InitializerName $Initializer
    Wait-XlflowFormImageCaptureWindow -Excel $runtimeExcel -Workbook $runtimeWorkbook -ProcessId $processId -Token $captureToken -TimeoutMilliseconds 7000 -PollMilliseconds 100
  }
  if ([bool]$captureResult.dialog.found) {
    Set-XlflowFormExportDialogFailure -Dialog $captureResult.dialog -Selection $captureResult.selection
    throw "runtime dialog shown"
  }
  if ($null -ne $captureResult.exception) {
    throw $captureResult.exception.Exception
  }
  $phase = "find_form_window"
  $window = $captureResult.value.window

  $phase = "capture_window_image"
  $dimensions = Save-XlflowWindowImage -Hwnd $window.hwnd -Path $exportPath
  if (-not [string]::IsNullOrWhiteSpace($temporaryExportPath)) {
    Move-Item -LiteralPath $temporaryExportPath -Destination $resolvedOutputPath -Force
    $temporaryExportPath = ""
  }

  $phase = "build_result"
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = [ordered]@{
    kind = $(if ($sessionAttached) { "live_session" } else { "file" })
    path = $WorkbookPath
    description = $(Get-XlflowTargetDescription -Kind $(if ($sessionAttached) { "live_session" } else { "file" }))
    form = $FormName
    capture_state = "temporary_copy"
    note = "Runtime export used a temporary workbook copy."
  }
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  $result.forms = [ordered]@{
    name = $FormName
    basis = "runtime"
  }
  if (-not [string]::IsNullOrWhiteSpace($Initializer)) {
    $result.forms.initializer = $Initializer
  }
  $result.output = [ordered]@{
    path = $resolvedOutputPath
    format = "png"
  }
  if ($createdParentDirs) {
    $result.output.created_parent_dirs = $true
  }
  if ($null -ne $dimensions) {
    $result.output.width_px = $dimensions.width_px
    $result.output.height_px = $dimensions.height_px
  } else {
    $actualDimensions = Get-XlflowImageDimensions -Path $resolvedOutputPath
    if ($null -ne $actualDimensions) {
      $result.output.width_px = $actualDimensions.width_px
      $result.output.height_px = $actualDimensions.height_px
    }
  }

  Add-XlflowWarning -Result $result -Code "runtime_form_loads_initialize" -Message "Form image export loads the form at runtime and executes UserForm_Initialize."
  Add-XlflowWarning -Result $result -Code "runtime_form_temp_copy" -Message "Form image export executed against a temporary workbook copy so the source workbook and live session are not mutated."
  if (-not [string]::IsNullOrWhiteSpace($Initializer)) {
    Add-XlflowWarning -Result $result -Code "runtime_form_initializer_invoked" -Message ("Form image export also invoked " + $Initializer + "(ThisWorkbook).")
  }
  Add-XlflowWarning -Result $result -Code "userform_image_export_experimental" -Message "UserForm image export is experimental and currently supports Windows desktop Excel only."
  if ($saveState.needs_save) {
    Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. `form export-image` used the live workbook state, not the saved workbook file."
  }
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "exported runtime UserForm " + $FormName + " to " + $resolvedOutputPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  if ($null -eq $result.error) {
    $code = Get-XlflowFormExportImageErrorCode -Exception $_.Exception
    Set-XlflowError -Result $result -Code $code -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult -Phase $phase
  }
  if ($null -eq $result.workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  }
  if ($null -eq $result.target -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.target = [ordered]@{
      kind = $(if ($sessionAttached) { "live_session" } else { "file" })
      path = $WorkbookPath
      form = $FormName
      capture_state = "temporary_copy"
      note = "Runtime export used a temporary workbook copy."
    }
  }
  if ($null -eq $result.session -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  }
} finally {
  if ($null -ne $runtimeExcel -and $null -ne $runtimeWorkbook) {
    Invoke-XlflowCleanupFormImageCapture -Excel $runtimeExcel -Workbook $runtimeWorkbook
  }
  if ($tempModuleInstalled -and $null -ne $runtimeVBProject) {
    $tempModuleRemoved = Remove-XlflowVBComponentByName -VBProject $runtimeVBProject -Name $tempModuleName
    if (-not $tempModuleRemoved) {
      Add-XlflowWarning -Result $result -Code "temporary_component_cleanup_failed" -Message ("Temporary helper module '" + $tempModuleName + "' could not be removed automatically.")
    }
  }
  if ($null -ne $runtimeWorkbook -or $null -ne $runtimeExcel) {
    Close-XlflowCom -Workbook $runtimeWorkbook -Excel $runtimeExcel -Save $false
  }
  if (-not [string]::IsNullOrWhiteSpace($runtimeWorkbookPath) -and (Test-Path -LiteralPath $runtimeWorkbookPath)) {
    Remove-Item -LiteralPath $runtimeWorkbookPath -Force -ErrorAction SilentlyContinue
  }
  if (-not [string]::IsNullOrWhiteSpace($temporaryExportPath) -and (Test-Path -LiteralPath $temporaryExportPath -PathType Leaf)) {
    try {
      Remove-Item -LiteralPath $temporaryExportPath -Force
    } catch {
      Write-Verbose ("failed to remove temporary form export image: " + $_.Exception.Message)
    }
  }
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
