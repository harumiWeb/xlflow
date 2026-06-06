try {
  $script:XlflowConsoleUtf8 = New-Object System.Text.UTF8Encoding($false)
  [Console]::InputEncoding = $script:XlflowConsoleUtf8
  [Console]::OutputEncoding = $script:XlflowConsoleUtf8
  $OutputEncoding = $script:XlflowConsoleUtf8
} catch {
  Write-Verbose ("failed to configure UTF-8 console encoding: " + $_.Exception.Message)
}

function Get-XlflowPowerShellBridgeInfo {
  $hostName = ""
  $edition = ""
  $version = ""

  try {
    $process = Get-Process -Id $PID -ErrorAction Stop
    if ($null -ne $process -and -not [string]::IsNullOrWhiteSpace($process.ProcessName)) {
      $hostName = $process.ProcessName
      if (-not $hostName.EndsWith(".exe")) {
        $hostName += ".exe"
      }
    }
  } catch {
    Write-Verbose ("failed to resolve PowerShell bridge process name: " + $_.Exception.Message)
  }

  try {
    $edition = [string]$PSVersionTable.PSEdition
  } catch {
    Write-Verbose ("failed to resolve PowerShell edition: " + $_.Exception.Message)
  }

  try {
    if ($null -ne $PSVersionTable -and $null -ne $PSVersionTable.PSVersion) {
      $version = $PSVersionTable.PSVersion.ToString()
    }
  } catch {
    Write-Verbose ("failed to resolve PowerShell version: " + $_.Exception.Message)
  }

  if ([string]::IsNullOrWhiteSpace($hostName)) {
    if ($edition -eq "Core") {
      $hostName = "pwsh.exe"
    } else {
      $hostName = "powershell.exe"
    }
  }

  return [ordered]@{
    host = $hostName
    edition = $edition
    version = $version
  }
}

function New-XlflowResult {
  param([string]$Command)
  return [ordered]@{
    status = "ok"
    command = $Command
    error = $null
    logs = @()
    bridge = (Get-XlflowPowerShellBridgeInfo)
  }
}

function Set-XlflowError {
  param(
    $Result,
    [string]$Code,
    [string]$Message,
    [string]$Source = "",
    [int]$Number = 0,
    [int]$Line = 0,
    [string]$Phase = ""
  )
  $Result.status = "failed"
  $Result.error = [ordered]@{
    code = $Code
    message = $Message
    source = $Source
    number = $Number
    line = $Line
    phase = $Phase
  }
}

function Add-XlflowWarning {
  param(
    $Result,
    [string]$Code,
    [string]$Message
  )

  if (-not $Result.Contains("warnings") -or $null -eq $Result["warnings"]) {
    $Result["warnings"] = @()
  }
  $Result["warnings"] += [ordered]@{
    code = $Code
    message = $Message
  }
}

function Add-XlflowHint {
  param(
    $Result,
    [string]$Code,
    [string]$Message
  )

  if (-not $Result.Contains("hints") -or $null -eq $Result["hints"]) {
    $Result["hints"] = @()
  }
  $Result["hints"] += [ordered]@{
    code = $Code
    message = $Message
  }
}

function Add-XlflowStateWarning {
  param(
    $Result,
    [string]$Code,
    [string]$Message
  )

  Add-XlflowWarning -Result $Result -Code $Code -Message $Message
}

function Get-XlflowRuntimeModeName {
  param([string]$Mode)

  $normalized = ([string]$Mode).Trim().ToLowerInvariant()
  switch ($normalized) {
    "headless" { return "headless" }
    "ci" { return "ci" }
    "agent" { return "agent" }
    "test" { return "test" }
    default { return "interactive" }
  }
}

function Get-XlflowRuntimeModeRefersTo {
  param([string]$Mode)

  return '="' + (Get-XlflowRuntimeModeName -Mode $Mode).Replace('"', '""') + '"'
}

function Get-XlflowRuntimeVersionRefersTo {
  return '="1"'
}

function Get-XlflowStringRefersTo {
  param([string]$Value)

  return '="' + ([string]$Value).Replace('"', '""') + '"'
}

function DecodeWorkbookDefinedName {
  param([string]$RefersTo)

  $value = [string]$RefersTo
  if ([string]::IsNullOrWhiteSpace($value)) {
    return ""
  }
  if ($value.StartsWith('="') -and $value.EndsWith('"') -and $value.Length -ge 3) {
    return $value.Substring(2, $value.Length - 3).Replace('""', '"')
  }
  if ($value.StartsWith('=')) {
    return $value.Substring(1)
  }
  return $value
}

function ConvertTo-XlflowUIResponseId {
  param([string]$Value)

  $value = ([string]$Value).Trim().ToLowerInvariant()
  $builder = New-Object System.Text.StringBuilder
  $lastSeparator = $false
  foreach ($char in $value.ToCharArray()) {
    $isValid = (($char -ge [char]'a' -and $char -le [char]'z') -or ($char -ge [char]'0' -and $char -le [char]'9'))
    if ($isValid) {
      [void]$builder.Append($char)
      $lastSeparator = $false
      continue
    }
    if (-not $lastSeparator -and $builder.Length -gt 0) {
      [void]$builder.Append("_")
      $lastSeparator = $true
    }
  }
  return $builder.ToString().Trim("_")
}

function Get-XlflowUIResponseDefinedName {
  param(
    [string]$Kind,
    [string]$Id
  )

  $normalizedKind = ([string]$Kind).Trim().ToLowerInvariant()
  if ($normalizedKind -ne "msgbox" -and $normalizedKind -ne "input") {
    throw "unsupported xlflow UI response kind: $Kind"
  }
  $normalizedId = ConvertTo-XlflowUIResponseId -Value $Id
  if ([string]::IsNullOrWhiteSpace($normalizedId)) {
    throw "xlflow UI response id cannot be empty"
  }
  return "__XLFLOW_UI_" + $normalizedKind.ToUpperInvariant() + "_" + $normalizedId + "__"
}

function Get-XlflowFileDialogResponseDefinedName {
  param(
    [string]$Kind,
    [string]$Id
  )

  $normalizedKind = ([string]$Kind).Trim().ToLowerInvariant()
  switch ($normalizedKind) {
    "get-open" { $kindToken = "GET_OPEN" }
    "file-open" { $kindToken = "FILE_OPEN" }
    "save-as" { $kindToken = "SAVE_AS" }
    "folder" { $kindToken = "FOLDER" }
    default { throw "unsupported xlflow file dialog kind: $Kind" }
  }
  $normalizedId = ConvertTo-XlflowUIResponseId -Value $Id
  if ([string]::IsNullOrWhiteSpace($normalizedId)) {
    throw "xlflow file dialog response id cannot be empty"
  }
  return "__XLFLOW_UI_FILEDIALOG_" + $kindToken + "_" + $normalizedId + "__"
}

function ConvertFrom-XlflowUIResponsesJson {
  param([string]$Json)

  $responses = [ordered]@{}
  if ([string]::IsNullOrWhiteSpace([string]$Json)) {
    return $responses
  }
  $decodedBytes = [System.Convert]::FromBase64String([string]$Json)
  $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
  $parsed = ConvertFrom-Json -InputObject $decodedJson
  if ($null -eq $parsed) {
    return $responses
  }
  foreach ($property in $parsed.PSObject.Properties) {
    $responses[[string]$property.Name] = [string]$property.Value
  }
  return $responses
}

function ConvertFrom-XlflowFileDialogResponsesJson {
  param([string]$Json)

  if ([string]::IsNullOrWhiteSpace([string]$Json)) {
    return @()
  }
  $decodedBytes = [System.Convert]::FromBase64String([string]$Json)
  $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
  $parsed = ConvertFrom-Json -InputObject $decodedJson
  if ($null -eq $parsed) {
    return @()
  }
  return @($parsed)
}

function ConvertTo-XlflowFileDialogMarkerValue {
  param($Response)

  if ($null -eq $Response) {
    return ""
  }
  if ([bool]$Response.cancelled) {
    return "@cancel"
  }
  $values = New-Object System.Collections.Generic.List[string]
  foreach ($value in @($Response.values)) {
    if ($null -ne $value) {
      $values.Add([string]$value) | Out-Null
    }
  }
  return [string]::Join("`n", $values.ToArray())
}

function Get-XlflowWorkbookDefinedNameState {
  param(
    $Workbook,
    [string]$Name
  )

  $state = [ordered]@{
    exists = $false
    refers_to = ""
    visible = $false
  }
  if ($null -eq $Workbook -or [string]::IsNullOrWhiteSpace($Name)) {
    return $state
  }
  try {
    $definedName = $Workbook.Names.Item($Name)
    if ($null -eq $definedName) {
      return $state
    }
    $state.exists = $true
    try {
      $state.refers_to = [string]$definedName.RefersTo
    } catch {
      Write-Verbose ("failed to read workbook defined name RefersTo: " + $_.Exception.Message)
    }
    try {
      $state.visible = [bool]$definedName.Visible
    } catch {
      $state.visible = $false
    }
    return $state
  } catch {
    return $state
  }
}

function Set-XlflowWorkbookDefinedName {
  param(
    $Workbook,
    [string]$Name,
    [string]$RefersTo,
    [bool]$Visible = $false
  )

  if ($null -eq $Workbook -or [string]::IsNullOrWhiteSpace($Name) -or [string]::IsNullOrWhiteSpace($RefersTo)) {
    return
  }
  $definedName = $null
  try {
    $definedName = $Workbook.Names.Item($Name)
  } catch {
    $definedName = $null
  }
  if ($null -eq $definedName) {
    $definedName = $Workbook.Names.Add($Name, $RefersTo)
  } else {
    $definedName.RefersTo = $RefersTo
  }
  try {
    $definedName.Visible = $Visible
  } catch {
    Write-Verbose ("failed to set workbook defined name visibility: " + $_.Exception.Message)
  }
}

function Remove-XlflowWorkbookDefinedName {
  param(
    $Workbook,
    [string]$Name
  )

  if ($null -eq $Workbook -or [string]::IsNullOrWhiteSpace($Name)) {
    return
  }
  try {
    $definedName = $Workbook.Names.Item($Name)
    if ($null -ne $definedName) {
      $definedName.Delete()
    }
  } catch {
    Write-Verbose ("failed to delete workbook defined name: " + $_.Exception.Message)
  }
}

function Get-XlflowWorkbookPersistedStateHash {
  param($Workbook)

  if ($null -eq $Workbook) {
    return ""
  }
  $extension = [System.IO.Path]::GetExtension([string]$Workbook.FullName)
  if ([string]::IsNullOrWhiteSpace($extension)) {
    $extension = ".xlsm"
  }
  $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("xlflow-runtime-" + [guid]::NewGuid().ToString("N") + $extension)
  try {
    $Workbook.SaveCopyAs($tempPath)
    return Get-XlflowFileHash -Path $tempPath
  } catch {
    Write-Verbose ("failed to capture workbook persisted-state hash: " + $_.Exception.Message)
    return ""
  } finally {
    if (Test-Path -LiteralPath $tempPath) {
      Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
  }
}

function Start-XlflowRuntimeInjection {
  param(
    $Workbook,
    $Result,
    [string]$Mode = "",
    [string]$Source = "default",
    [string]$MsgBoxResponsesJSON = "",
    [string]$InputResponsesJSON = "",
    [string]$FileDialogResponsesJSON = "",
    [string]$DebugStreamEnabled = "false",
    [string]$DebugStreamPipeName = "",
    [string]$UIStreamEnabled = "false",
    [string]$UIStreamPipeName = "",
    [string]$UIStreamRedactInput = "true"
  )

  $modeName = Get-XlflowRuntimeModeName -Mode $Mode
  $msgBoxResponses = ConvertFrom-XlflowUIResponsesJson -Json $MsgBoxResponsesJSON
  $inputResponses = ConvertFrom-XlflowUIResponsesJson -Json $InputResponsesJSON
  $fileDialogResponses = @(ConvertFrom-XlflowFileDialogResponsesJson -Json $FileDialogResponsesJSON)
  $debugStreamEnabled = (ConvertTo-XlflowBool $DebugStreamEnabled) -and -not [string]::IsNullOrWhiteSpace([string]$DebugStreamPipeName)
  $uiStreamEnabled = (ConvertTo-XlflowBool $UIStreamEnabled) -and -not [string]::IsNullOrWhiteSpace([string]$UIStreamPipeName)
  $state = [ordered]@{
    applied = $false
    saved_before = $null
    baseline_hash = ""
    debug_stream_enabled = $debugStreamEnabled
    debug_stream_pipe_name = [string]$DebugStreamPipeName
    ui_stream_enabled = $uiStreamEnabled
    ui_stream_pipe_name = [string]$UIStreamPipeName
    ui_stream_redact_input = (ConvertTo-XlflowBool $UIStreamRedactInput)
    ui_stream_module = ""
    names = [ordered]@{
      __XLFLOW_MODE__ = (Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name "__XLFLOW_MODE__")
      __XLFLOW_RUNTIME_VERSION__ = (Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name "__XLFLOW_RUNTIME_VERSION__")
    }
  }

  foreach ($entry in $msgBoxResponses.GetEnumerator()) {
    $name = Get-XlflowUIResponseDefinedName -Kind "msgbox" -Id ([string]$entry.Key)
    $state.names[$name] = Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name $name
  }
  foreach ($entry in $inputResponses.GetEnumerator()) {
    $name = Get-XlflowUIResponseDefinedName -Kind "input" -Id ([string]$entry.Key)
    $state.names[$name] = Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name $name
  }
  foreach ($entry in $fileDialogResponses) {
    $name = Get-XlflowFileDialogResponseDefinedName -Kind ([string]$entry.kind) -Id ([string]$entry.dialog_id)
    $state.names[$name] = Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name $name
  }
  $state.names["__XLFLOW_DEBUG_PIPE__"] = Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name "__XLFLOW_DEBUG_PIPE__"
  $state.names["__XLFLOW_UI_STREAM_HELPER__"] = Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name "__XLFLOW_UI_STREAM_HELPER__"
  $state.names["__XLFLOW_UI_STREAM_REDACT_INPUT__"] = Get-XlflowWorkbookDefinedNameState -Workbook $Workbook -Name "__XLFLOW_UI_STREAM_REDACT_INPUT__"
  if ($null -eq $Workbook) {
    if ($null -ne $Result) {
      $Result.runtime = [ordered]@{
        mode = $modeName
        mode_name = $modeName
        source = ([string]$Source).Trim().ToLowerInvariant()
        injected = $false
      }
    }
    return $state
  }

  try {
    $state.saved_before = [bool]$Workbook.Saved
  } catch {
    $state.saved_before = $null
  }
  if ($state.saved_before -eq $true) {
    $state.baseline_hash = Get-XlflowWorkbookPersistedStateHash -Workbook $Workbook
  }

  Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name "__XLFLOW_MODE__" -RefersTo (Get-XlflowRuntimeModeRefersTo -Mode $modeName) -Visible $false
  Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name "__XLFLOW_RUNTIME_VERSION__" -RefersTo (Get-XlflowRuntimeVersionRefersTo) -Visible $false
  foreach ($entry in $msgBoxResponses.GetEnumerator()) {
    Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name (Get-XlflowUIResponseDefinedName -Kind "msgbox" -Id ([string]$entry.Key)) -RefersTo (Get-XlflowStringRefersTo -Value ([string]$entry.Value)) -Visible $false
  }
  foreach ($entry in $inputResponses.GetEnumerator()) {
    Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name (Get-XlflowUIResponseDefinedName -Kind "input" -Id ([string]$entry.Key)) -RefersTo (Get-XlflowStringRefersTo -Value ([string]$entry.Value)) -Visible $false
  }
  foreach ($entry in $fileDialogResponses) {
    Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name (Get-XlflowFileDialogResponseDefinedName -Kind ([string]$entry.kind) -Id ([string]$entry.dialog_id)) -RefersTo (Get-XlflowStringRefersTo -Value (ConvertTo-XlflowFileDialogMarkerValue -Response $entry)) -Visible $false
  }
  if ($debugStreamEnabled) {
    Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name "__XLFLOW_DEBUG_PIPE__" -RefersTo (Get-XlflowStringRefersTo -Value ([string]$DebugStreamPipeName)) -Visible $false
  }
  $state.applied = $true

  if ($null -ne $Result) {
    $Result.runtime = [ordered]@{
      mode = $modeName
      mode_name = $modeName
      source = ([string]$Source).Trim().ToLowerInvariant()
      injected = $true
    }
  }

  return $state
}

function Restore-XlflowRuntimeInjection {
  param(
    $Workbook,
    $State
  )

  if ($null -eq $Workbook -or $null -eq $State -or -not [bool]$State.applied) {
    return
  }

  foreach ($name in @($State.names.Keys)) {
    $nameState = $State.names[$name]
    if ($null -eq $nameState) {
      Remove-XlflowWorkbookDefinedName -Workbook $Workbook -Name $name
      continue
    }
    if ([bool]$nameState.exists) {
      Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name $name -RefersTo ([string]$nameState.refers_to) -Visible ([bool]$nameState.visible)
    } else {
      Remove-XlflowWorkbookDefinedName -Workbook $Workbook -Name $name
    }
  }

  if (-not [string]::IsNullOrWhiteSpace([string]$State.ui_stream_module)) {
    try {
      $project = $Workbook.VBProject
      if ($null -ne $project) {
        Remove-XlflowUIStreamModule -VBProject $project -ModuleName ([string]$State.ui_stream_module) | Out-Null
      }
    } catch {
      Write-Verbose ("failed to remove UI stream module during runtime cleanup: " + $_.Exception.Message)
    }
  }

  if ($State.saved_before -eq $true -and -not [string]::IsNullOrWhiteSpace([string]$State.baseline_hash)) {
    $currentHash = Get-XlflowWorkbookPersistedStateHash -Workbook $Workbook
    if (-not [string]::IsNullOrWhiteSpace($currentHash) -and $currentHash -eq [string]$State.baseline_hash) {
      try {
        $Workbook.Saved = $true
      } catch {
        Write-Verbose ("failed to restore workbook saved state after runtime cleanup: " + $_.Exception.Message)
      }
    }
  }
}

function Enable-XlflowUIStreamRuntimeInjection {
  param(
    $Workbook,
    $State,
    $VBProject
  )

  if ($null -eq $Workbook -or $null -eq $State -or -not [bool]$State.applied -or -not [bool]$State.ui_stream_enabled) {
    return $false
  }
  if ($null -eq $VBProject -or [string]::IsNullOrWhiteSpace([string]$State.ui_stream_pipe_name)) {
    return $false
  }

  $State.ui_stream_module = "XlflowUIStream_" + [guid]::NewGuid().ToString("N").Substring(0, 8)
  Install-XlflowUIStreamModule -VBProject $VBProject -ModuleName $State.ui_stream_module -PipeName ([string]$State.ui_stream_pipe_name)
  Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name "__XLFLOW_UI_STREAM_HELPER__" -RefersTo (Get-XlflowStringRefersTo -Value ($State.ui_stream_module + ".EmitEvent")) -Visible $false
  Set-XlflowWorkbookDefinedName -Workbook $Workbook -Name "__XLFLOW_UI_STREAM_REDACT_INPUT__" -RefersTo (Get-XlflowStringRefersTo -Value (([bool]$State.ui_stream_redact_input).ToString().ToLowerInvariant())) -Visible $false
  return $true
}

function New-XlflowUIStreamModuleCode {
  param([string]$PipeName)

  $builder = New-Object System.Text.StringBuilder
  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("#If VBA7 Then")
  [void]$builder.AppendLine('  Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr')
  [void]$builder.AppendLine('  Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long')
  [void]$builder.AppendLine('  Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long')
  [void]$builder.AppendLine('  Private Const INVALID_HANDLE_VALUE As LongPtr = -1')
  [void]$builder.AppendLine("#Else")
  [void]$builder.AppendLine('  Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long')
  [void]$builder.AppendLine('  Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long')
  [void]$builder.AppendLine('  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long')
  [void]$builder.AppendLine('  Private Const INVALID_HANDLE_VALUE As Long = -1')
  [void]$builder.AppendLine("#End If")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Private Const GENERIC_WRITE As Long = &H40000000")
  [void]$builder.AppendLine("Private Const OPEN_EXISTING As Long = 3")
  [void]$builder.AppendLine('Private Const mPipeName As String = "' + $PipeName.Replace('"', '""') + '"')
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Sub EmitEvent(ByVal jsonText As String)")
  [void]$builder.AppendLine("  SendPipeText jsonText & vbLf")
  [void]$builder.AppendLine("End Sub")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Private Sub SendPipeText(ByVal payload As String)")
  [void]$builder.AppendLine("  Dim bytesWritten As Long")
  [void]$builder.AppendLine("#If VBA7 Then")
  [void]$builder.AppendLine("  Dim pipeHandle As LongPtr")
  [void]$builder.AppendLine("#Else")
  [void]$builder.AppendLine("  Dim pipeHandle As Long")
  [void]$builder.AppendLine("#End If")
  [void]$builder.AppendLine("  pipeHandle = CreateFileW(StrPtr(mPipeName), GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)")
  [void]$builder.AppendLine("  If pipeHandle = INVALID_HANDLE_VALUE Then")
  [void]$builder.AppendLine("    Exit Sub")
  [void]$builder.AppendLine("  End If")
  [void]$builder.AppendLine("  On Error GoTo Cleanup")
  [void]$builder.AppendLine("  Call WriteFile(pipeHandle, StrPtr(payload), Len(payload) * 2, bytesWritten, 0)")
  [void]$builder.AppendLine("Cleanup:")
  [void]$builder.AppendLine("  On Error Resume Next")
  [void]$builder.AppendLine("  If pipeHandle <> INVALID_HANDLE_VALUE Then CloseHandle pipeHandle")
  [void]$builder.AppendLine("  On Error GoTo 0")
  [void]$builder.AppendLine("End Sub")
  return $builder.ToString()
}

function Install-XlflowUIStreamModule {
  param(
    $VBProject,
    [string]$ModuleName,
    [string]$PipeName
  )

  $component = $VBProject.VBComponents.Add(1)
  $component.Name = $ModuleName
  $component.CodeModule.AddFromString((New-XlflowUIStreamModuleCode -PipeName $PipeName))
  return $ModuleName
}

function Remove-XlflowUIStreamModule {
  param(
    $VBProject,
    [string]$ModuleName
  )

  if ($null -eq $VBProject -or [string]::IsNullOrWhiteSpace($ModuleName)) {
    return $false
  }
  try {
    $existing = $VBProject.VBComponents.Item($ModuleName)
    $VBProject.VBComponents.Remove($existing)
    return $true
  } catch {
    return $false
  }
}

function Get-XlflowUserFormComponents {
  param($Workbook)

  $forms = New-Object System.Collections.Generic.List[object]
  if ($null -eq $Workbook) {
    return @()
  }
  try {
    foreach ($component in @($Workbook.VBProject.VBComponents)) {
      if ($null -ne $component -and [int]$component.Type -eq 3) {
        $forms.Add($component)
      }
    }
  } catch {
    throw
  }
  return @($forms.ToArray())
}

function Get-XlflowUserFormNames {
  param($Workbook)

  $names = New-Object System.Collections.Generic.List[string]
  foreach ($component in @(Get-XlflowUserFormComponents -Workbook $Workbook)) {
    try {
      $name = [string]$component.Name
      if (-not [string]::IsNullOrWhiteSpace($name)) {
        $names.Add($name)
      }
    } catch {
      Write-Verbose ("failed to read UserForm component name: " + $_.Exception.Message)
    }
  }
  return @($names.ToArray())
}

function Add-XlflowUserFormDiscoveryMessages {
  param(
    $Result,
    [string[]]$Names
  )

  $normalized = @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
  if ($normalized.Count -eq 0) {
    return
  }
  Add-XlflowWarning -Result $Result -Code "userform_state_partial" -Message ("UserForms detected: " + ($normalized -join ", ") + ". `.frm` text may not fully represent layout, binary `.frx` state, or VBIDE Designer-backed properties.")
  Add-XlflowHint -Result $Result -Code "userform_planned_commands" -Message "UserForm workflow: `xlflow pull --json`, `xlflow inspect form <name> --designer --json`, `xlflow form snapshot <name> --out src/forms/specs/<name>.yaml`, edit spec/code artifacts, then `xlflow form build src/forms/specs/<name>.yaml --overwrite` and verify with `xlflow form export-image <name> --out <path>`."
}

function Add-XlflowUserFormSessionStaleWarning {
  param(
    $Result,
    [string[]]$Names
  )

  $normalized = @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
  if ($normalized.Count -eq 0) {
    return
  }
  Add-XlflowStateWarning -Result $Result -Code "userform_unsaved_session_state" -Message ("Workbook contains UserForms (" + ($normalized -join ", ") + ") and the live workbook is newer than disk. Run `xlflow save --session` and `xlflow pull` before reviewing `.frm`/`.frx` or disk-backed inspect output.")
}

function Get-XlflowSourceUserFormNames {
  param([string]$FormsDir)

  if ([string]::IsNullOrWhiteSpace($FormsDir) -or -not (Test-Path -LiteralPath $FormsDir)) {
    return @()
  }

  $names = New-Object System.Collections.Generic.List[string]
  foreach ($file in Get-ChildItem -LiteralPath $FormsDir -Recurse -File -Filter *.frm | Sort-Object FullName) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    if (-not [string]::IsNullOrWhiteSpace($name)) {
      $names.Add($name)
    }
  }
  return @($names.ToArray() | Sort-Object -Unique)
}

function Use-XlflowUserFormCodeSidecar {
  param([string]$CodeSource)

  return [string]::Equals(([string]$CodeSource).Trim(), "sidecar", [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-XlflowUserFormCodeDir {
  param([string]$FormsDir)

  if ([string]::IsNullOrWhiteSpace($FormsDir)) {
    return ""
  }
  return Join-Path $FormsDir "code"
}

function Get-XlflowUserFormCodePath {
  param(
    [string]$FormsDir,
    [string]$FormName
  )

  $codeDir = Get-XlflowUserFormCodeDir -FormsDir $FormsDir
  if ([string]::IsNullOrWhiteSpace($codeDir) -or [string]::IsNullOrWhiteSpace($FormName)) {
    return ""
  }
  return Join-Path $codeDir ($FormName + ".bas")
}

function Get-XlflowUserFormCodeFiles {
  param([string]$FormsDir)

  $codeDir = Get-XlflowUserFormCodeDir -FormsDir $FormsDir
  if ([string]::IsNullOrWhiteSpace($codeDir) -or -not (Test-Path -LiteralPath $codeDir)) {
    return @()
  }

  $files = New-Object System.Collections.Generic.List[object]
  foreach ($file in Get-ChildItem -LiteralPath $codeDir -Recurse -File -Filter *.bas | Sort-Object FullName) {
    $formName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $files.Add([pscustomobject][ordered]@{
      kind = "form_code"
      root_dir = $codeDir
      full_name = $file.FullName
      relative_path = (Get-XlflowRelativePath -BasePath $codeDir -TargetPath $file.FullName).Replace("\", "/")
      extension = ".bas"
      module_name = $formName
      form_name = $formName
    })
  }
  return @($files.ToArray())
}

function ConvertTo-XlflowBool {
  param([string]$Value)
  return [bool](($Value -eq "true") -or ($Value -eq "True") -or ($Value -eq "1"))
}

function Get-XlflowRelativePath {
  param(
    [string]$BasePath,
    [string]$TargetPath
  )

  if ([string]::IsNullOrWhiteSpace($BasePath) -or [string]::IsNullOrWhiteSpace($TargetPath)) {
    return ""
  }

  $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
  $targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)
  $directorySeparator = [System.IO.Path]::DirectorySeparatorChar
  $altDirectorySeparator = [System.IO.Path]::AltDirectorySeparatorChar

  if (-not $baseFullPath.EndsWith([string]$directorySeparator) -and -not $baseFullPath.EndsWith([string]$altDirectorySeparator)) {
    $baseFullPath += $directorySeparator
  }

  $baseUri = New-Object System.Uri($baseFullPath)
  $targetUri = New-Object System.Uri($targetFullPath)
  $relativeUri = $baseUri.MakeRelativeUri($targetUri)
  $relativeUriText = [System.Uri]::UnescapeDataString($relativeUri.ToString())
  if ($relativeUri.IsAbsoluteUri -or $relativeUriText -match '^[A-Za-z]+:') {
    return $targetFullPath
  }
  $relativePath = $relativeUriText.Replace("/", "\")

  if ([string]::IsNullOrWhiteSpace($relativePath)) {
    return "."
  }

  return $relativePath
}

function Get-XlflowTargetDescription {
  param([string]$Kind)

  switch ($Kind) {
    "source" { return "VBA source files in the project directory" }
    "file" { return "Saved workbook file on disk" }
    "live_session" { return "Workbook currently open through xlflow session" }
    default { return "" }
  }
}

function New-XlflowTargetResult {
  param(
    [string]$Kind,
    [string]$Path = "",
    [string]$Description = "",
    [string]$Note = "",
    [string]$Sheet = "",
    [string]$Range = ""
  )

  $target = [ordered]@{
    kind = $Kind
  }
  if (-not [string]::IsNullOrWhiteSpace($Path)) {
    $target.path = $Path
  }
  if ([string]::IsNullOrWhiteSpace($Description)) {
    $Description = Get-XlflowTargetDescription -Kind $Kind
  }
  if (-not [string]::IsNullOrWhiteSpace($Description)) {
    $target.description = $Description
  }
  if (-not [string]::IsNullOrWhiteSpace($Note)) {
    $target.note = $Note
  }
  if (-not [string]::IsNullOrWhiteSpace($Sheet)) {
    $target.sheet = $Sheet
  }
  if (-not [string]::IsNullOrWhiteSpace($Range)) {
    $target.range = $Range
  }
  return $target
}

function New-XlflowSessionResult {
  param(
    [bool]$Active = $false,
    [string]$WorkbookPath = "",
    [AllowNull()]$Dirty = $null,
    [AllowNull()]$SaveRequired = $null,
    [string]$Mode = "",
    [string]$SourceOfTruth = "",
    [AllowNull()]$UserFormsPresent = $null,
    [AllowNull()]$UserFormCount = $null
  )

  $session = [ordered]@{
    active = $Active
  }
  if (-not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $session.workbook_path = $WorkbookPath
  }
  if ($PSBoundParameters.ContainsKey("Dirty") -and $null -ne $Dirty) {
    $session.dirty = [bool]$Dirty
  }
  if ($PSBoundParameters.ContainsKey("SaveRequired") -and $null -ne $SaveRequired) {
    $session.save_required = [bool]$SaveRequired
    $session.live_newer_than_disk = [bool]$SaveRequired
  }
  if (-not [string]::IsNullOrWhiteSpace($Mode)) {
    $session.mode = $Mode
  }
  if ([string]::IsNullOrWhiteSpace($SourceOfTruth)) {
    if ($session.Contains("live_newer_than_disk") -and [bool]$session.live_newer_than_disk) {
      $SourceOfTruth = "live_workbook"
    } else {
      $SourceOfTruth = "saved_workbook"
    }
  }
  if (-not [string]::IsNullOrWhiteSpace($SourceOfTruth)) {
    $session.source_of_truth = $SourceOfTruth
  }
  if ($PSBoundParameters.ContainsKey("UserFormsPresent") -and $null -ne $UserFormsPresent) {
    $session.userforms_present = [bool]$UserFormsPresent
  }
  if ($PSBoundParameters.ContainsKey("UserFormCount") -and $null -ne $UserFormCount) {
    $session.userform_count = [int]$UserFormCount
  }
  return $session
}

function Close-XlflowCom {
  param($Workbook, $Excel, [bool]$Save)
  $excelProcessId = 0
  $workbookCloseFailed = $false
  $quitFailed = $false
  if ($null -ne $Excel) {
    $excelProcessId = Get-XlflowExcelProcessId -Excel $Excel
  }
  if ($null -ne $Workbook) {
    try { $Workbook.Close($Save) | Out-Null } catch { $workbookCloseFailed = $true; Write-Verbose ("failed to close workbook: " + $_.Exception.Message) }
  }
  if ($null -ne $Excel) {
    try { $Excel.Quit() | Out-Null } catch { $quitFailed = $true; Write-Verbose ("failed to quit Excel: " + $_.Exception.Message) }
  }
  Release-XlflowComObject -Object $Workbook -Name "workbook COM object"
  Release-XlflowComObject -Object $Excel -Name "Excel COM object"
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  if ($excelProcessId -gt 0) {
    $stillRunning = $false
    for ($i = 0; $i -lt 20; $i++) {
      if (-not (Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue)) {
        $stillRunning = $false
        break
      }
      $stillRunning = $true
      Start-Sleep -Milliseconds 100
    }
    if ($stillRunning) {
      if ($workbookCloseFailed -or $quitFailed) {
        Write-Verbose ("skipping force-stop for lingering Excel process " + $excelProcessId + " after graceful close failure")
      } else {
        try {
          Stop-Process -Id $excelProcessId -Force
        } catch {
          Write-Verbose ("failed to stop lingering Excel process " + $excelProcessId + ": " + $_.Exception.Message)
        }
      }
    }
  }
}

function Release-XlflowComObject {
  param(
    $Object,
    [string]$Name = "COM object"
  )

  if ($null -eq $Object) {
    return
  }

  try {
    if (-not [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
      return
    }
    while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) -gt 0) {
    }
  } catch {
    Write-Verbose ("failed to release " + $Name + ": " + $_.Exception.Message)
  }
}

function Release-XlflowComReferences {
  param($Workbook, $Excel)
  $sessionWorkbook = Get-Variable -Name "XlflowSessionWorkbook" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  $sessionExcel = Get-Variable -Name "XlflowSessionExcel" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  if ($null -ne $Workbook) {
    if ($null -ne $sessionWorkbook -and [object]::ReferenceEquals($Workbook, $sessionWorkbook)) {
      Write-Verbose "leaving in-process xlflow session workbook reference open"
    } else {
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch { Write-Verbose ("failed to release workbook COM object: " + $_.Exception.Message) }
    }
  }
  if ($null -ne $Excel) {
    if ($null -ne $sessionExcel -and [object]::ReferenceEquals($Excel, $sessionExcel)) {
      Write-Verbose "leaving in-process xlflow session Excel reference open"
    } else {
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch { Write-Verbose ("failed to release Excel COM object: " + $_.Exception.Message) }
    }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Set-XlflowExcelAutomationDefaults {
  param($Excel, [bool]$DisplayAlerts = $false)

  if ($null -eq $Excel) {
    return
  }
  try { $Excel.DisplayAlerts = $DisplayAlerts } catch { Write-Verbose ("failed to set DisplayAlerts: " + $_.Exception.Message) }
  try { $Excel.EnableEvents = $false } catch { Write-Verbose ("failed to disable Excel events: " + $_.Exception.Message) }
}

function Disable-XlflowExcelAutomationMacros {
  param($Excel)

  if ($null -eq $Excel) {
    return
  }
  try { $Excel.AutomationSecurity = 3 } catch { Write-Verbose ("failed to force-disable automation macros: " + $_.Exception.Message) }
}

function Open-XlflowWorkbookWithXlflowDefaults {
  param(
    $Excel,
    [string]$WorkbookPath,
    [bool]$DisplayAlerts = $false,
    [bool]$DisableAutomationMacros = $true
  )

  Set-XlflowExcelAutomationDefaults -Excel $Excel -DisplayAlerts $DisplayAlerts
  if ($DisableAutomationMacros) {
    Disable-XlflowExcelAutomationMacros -Excel $Excel
  }
  return $Excel.Workbooks.Open($WorkbookPath)
}

function Get-XlflowActiveExcel {
  $sessionExcel = Get-Variable -Name "XlflowSessionExcel" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  if ($null -ne $sessionExcel) {
    return $sessionExcel
  }
  try {
    return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  } catch {
    throw "xlflow session is not running"
  }
}

function Get-XlflowOpenWorkbook {
  param($Excel, [string]$WorkbookPath)

  $target = [System.IO.Path]::GetFullPath($WorkbookPath)
  $sessionWorkbook = Get-Variable -Name "XlflowSessionWorkbook" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  if ($null -ne $sessionWorkbook) {
    try {
      if ([System.IO.Path]::GetFullPath([string]$sessionWorkbook.FullName) -ieq $target) {
        return $sessionWorkbook
      }
    } catch {
      Write-Verbose ("failed to inspect in-process session workbook: " + $_.Exception.Message)
    }
  }
  try {
    $bound = [System.Runtime.InteropServices.Marshal]::BindToMoniker($target)
    if ($null -ne $bound) {
      try {
        if ([System.IO.Path]::GetFullPath([string]$bound.FullName) -ieq $target) {
          return $bound
        }
      } catch {
        Write-Verbose ("failed to inspect moniker-bound workbook: " + $_.Exception.Message)
      }
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($bound) | Out-Null } catch { Write-Verbose ("failed to release moniker-bound workbook: " + $_.Exception.Message) }
    }
  } catch {
    Write-Verbose ("failed to bind open workbook by path: " + $_.Exception.Message)
  }
  foreach ($candidate in @($Excel.Workbooks)) {
    try {
      if ([System.IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $target) {
        return $candidate
      }
    } catch {
      Write-Verbose ("failed to inspect open workbook: " + $_.Exception.Message)
    }
  }
  throw "xlflow session workbook is not open: $WorkbookPath"
}

function Add-XlflowNativeMethods {
  if ("XlflowNativeMethods" -as [type]) {
    return
  }
  Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class XlflowNativeMethods {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }

  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")]
  public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

  [DllImport("oleacc.dll")]
  public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectId, ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

  public static IntPtr[] GetWindowsForProcess(uint targetProcessId) {
    List<IntPtr> windows = new List<IntPtr>();
    EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
      uint processId = 0;
      GetWindowThreadProcessId(hWnd, out processId);
      if (processId == targetProcessId) {
        windows.Add(hWnd);
      }
      return true;
    }, IntPtr.Zero);
    return windows.ToArray();
  }

  public static IntPtr[] GetTopLevelWindows() {
    List<IntPtr> windows = new List<IntPtr>();
    EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
      windows.Add(hWnd);
      return true;
    }, IntPtr.Zero);
    return windows.ToArray();
  }

  public static IntPtr[] GetChildWindows(IntPtr parentHwnd) {
    List<IntPtr> windows = new List<IntPtr>();
    EnumWindowsProc callback = delegate(IntPtr hWnd, IntPtr lParam) {
      windows.Add(hWnd);
      return true;
    };
    EnumChildWindows(parentHwnd, callback, IntPtr.Zero);
    return windows.ToArray();
  }

  [DllImport("user32.dll")]
  public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc enumProc, IntPtr lParam);

  [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  public static extern int GetWindowTextW(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern int GetWindowTextLengthW(IntPtr hWnd);

  [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  public static extern int GetClassNameW(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

  [DllImport("user32.dll")]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

  [DllImport("user32.dll")]
  public static extern uint GetDpiForWindow(IntPtr hWnd);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern IntPtr SendMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool PostMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

  public static string GetWindowTextString(IntPtr hWnd) {
    int length = GetWindowTextLengthW(hWnd);
    System.Text.StringBuilder builder = new System.Text.StringBuilder(Math.Max(length + 1, 256));
    GetWindowTextW(hWnd, builder, builder.Capacity);
    return builder.ToString();
  }

  public static string GetClassNameString(IntPtr hWnd) {
    System.Text.StringBuilder builder = new System.Text.StringBuilder(256);
    GetClassNameW(hWnd, builder, builder.Capacity);
    return builder.ToString();
  }
}
"@
}

function Get-XlflowExcelProcessId {
  param($Excel)

  if ($null -eq $Excel) {
    return 0
  }
  try {
    Add-XlflowNativeMethods
    $hwnd = [IntPtr]([int64]$Excel.Hwnd)
    $processId = [uint32]0
    [void][XlflowNativeMethods]::GetWindowThreadProcessId($hwnd, [ref]$processId)
    return [int]$processId
  } catch {
    Write-Verbose ("failed to resolve Excel process id: " + $_.Exception.Message)
    return 0
  }
}

function New-XlflowExcelDialogWatcherResult {
  return [pscustomobject][ordered]@{
    found = $false
    kind = ""
    hwnd = 0
    title = ""
    class_name = ""
    text = @()
    buttons = @()
    children = @()
    clicked_button = ""
    action = ""
    selection = [ordered]@{
      location = [ordered]@{
        module = ""
        line = 0
        column = 0
        end_line = 0
        end_column = 0
        token = ""
      }
      nearby_code = @()
    }
    break_mode_reset = $false
  }
}

function Get-XlflowExcelDialogMessageLines {
  param($Dialog)

  if ($null -eq $Dialog) {
    return @()
  }

  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($line in @($Dialog.text)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
      $lines.Add([string]$line)
    }
  }
  if ($lines.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$Dialog.title)) {
    $lines.Add([string]$Dialog.title)
  }
  return @($lines.ToArray())
}

function Get-XlflowVBARuntimeDialogErrorNumber {
  param($Dialog)

  $message = ((Get-XlflowExcelDialogMessageLines -Dialog $Dialog) -join [Environment]::NewLine)
  if ([string]::IsNullOrWhiteSpace($message)) {
    return 0
  }

  $patterns = @(
    "(?i)(?:run-?time error|runtime error)\s*'?(?<number>-?\d+)'?",
    "実行時エラー\s*'?(?<number>-?\d+)'?"
  )
  foreach ($pattern in $patterns) {
    $match = [regex]::Match($message, $pattern)
    if (-not $match.Success) {
      continue
    }
    $parsed = 0
    if ([int]::TryParse($match.Groups["number"].Value, [ref]$parsed)) {
      return $parsed
    }
  }
  return 0
}

function Test-XlflowCompileDialogSignals {
  param(
    [string]$Title,
    [string]$StaticText,
    [string]$ButtonText
  )

  return (
    $StaticText -match "(?i)(compile|syntax error|expected)" -or
    $StaticText -match "コンパイル|構文エラー|必要です" -or
    $ButtonText -match "(?i)(Debug|Compile)" -or
    $ButtonText -match "(デバッグ|コンパイル)" -or
    $Title -match "(?i)(compile|syntax error)" -or
    $Title -match "コンパイル|構文エラー"
  )
}

function Test-XlflowAllowDialogFirstButtonFallback {
  param(
    [string]$DialogKind,
    [string]$MatchedKind = ""
  )

  if ($MatchedKind -eq "compile") {
    return $false
  }
  return $DialogKind -ne "compile"
}

function Start-XlflowExcelDialogWatcher {
  param(
    [int]$ProcessId,
    [string]$Kind = "compile",
    [int]$TimeoutMilliseconds = 10000,
    [int]$PollMilliseconds = 50
  )

  Add-XlflowNativeMethods
  if ($ProcessId -le 0) {
    return $null
  }

  $commonScriptPath = Join-Path $PSScriptRoot "common.ps1"
  $ps = [PowerShell]::Create()
  $null = $ps.AddScript({
    param([int]$TargetProcessId, [string]$DialogKind, [int]$TimeoutMs, [int]$PollMs, [string]$CommonScriptPath)

    $bmClick = [uint32]0x00F5
    $wmClose = [uint32]0x0010
    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)

    while ([DateTime]::UtcNow -lt $deadline) {
      foreach ($hwnd in [XlflowNativeMethods]::GetWindowsForProcess([uint32]$TargetProcessId)) {
        if (-not [XlflowNativeMethods]::IsWindowVisible($hwnd)) {
          continue
        }

        $title = [XlflowNativeMethods]::GetWindowTextString($hwnd)
        $className = [XlflowNativeMethods]::GetClassNameString($hwnd)
        $childInfos = New-Object System.Collections.Generic.List[object]
        $staticTexts = New-Object System.Collections.Generic.List[string]
        $buttons = New-Object System.Collections.Generic.List[object]

        foreach ($child in [XlflowNativeMethods]::GetChildWindows($hwnd)) {
          $childText = [XlflowNativeMethods]::GetWindowTextString($child)
          $childClass = [XlflowNativeMethods]::GetClassNameString($child)
          $info = [pscustomobject][ordered]@{
            hwnd = [int64]$child
            class_name = $childClass
            text = $childText
          }
          $childInfos.Add($info)
          if ($childClass -eq "Static" -and -not [string]::IsNullOrWhiteSpace($childText)) {
            $staticTexts.Add($childText)
          }
          if ($childClass -eq "Button" -and -not [string]::IsNullOrWhiteSpace($childText)) {
            $buttons.Add($info)
          }
        }

        $buttonTexts = @($buttons | ForEach-Object { [string]$_.text })
        $joinedButtonText = ($buttonTexts -join [Environment]::NewLine)
        $joinedStaticText = (@($staticTexts.ToArray()) -join [Environment]::NewLine)
        $looksLikeVBAHostDialog = $className -eq "#32770" -and $buttons.Count -gt 0 -and ($staticTexts.Count -gt 0 -or $title -match "(?i)(visual basic|excel|vba)")
        $looksLikeRuntimeDialog = $looksLikeVBAHostDialog -and (
          $joinedStaticText -match "(?i)(run-?time error|runtime error)" -or
          $joinedStaticText -match "実行時エラー" -or
          $joinedButtonText -match "(?i)(Debug|End|Continue)" -or
          $joinedButtonText -match "(デバッグ|終了|継続)"
        )
        $looksLikeCompileDialog = $looksLikeVBAHostDialog -and (
          Test-XlflowCompileDialogSignals -Title $title -StaticText $joinedStaticText -ButtonText $joinedButtonText
        )

        $matchedKind = ""
        switch ($DialogKind) {
          "runtime" {
            if (-not $looksLikeRuntimeDialog) {
              continue
            }
            $matchedKind = "runtime"
          }
          "any_vba" {
            if ($looksLikeRuntimeDialog) {
              $matchedKind = "runtime"
            } elseif ($looksLikeCompileDialog) {
              $matchedKind = "compile"
            } else {
              continue
            }
          }
          default {
            if (-not $looksLikeCompileDialog) {
              continue
            }
            $matchedKind = "compile"
          }
        }

        $buttonToClick = $null
        $action = ""
        switch ($matchedKind) {
          "runtime" {
            foreach ($button in $buttons) {
              if ($button.text -match "(?i)Debug" -or $button.text -match "デバッグ") {
                $buttonToClick = $button
                $action = "runtime_debug"
                break
              }
            }
            if ($null -eq $buttonToClick) {
              foreach ($button in $buttons) {
              if ($button.text -match "(?i)End" -or $button.text -match "終了") {
                $buttonToClick = $button
                $action = "runtime_end"
                break
              }
            }
            }
            if ($null -eq $buttonToClick) {
              foreach ($button in $buttons) {
                if ($button.text -match "(?i)^(OK|Close)$" -or $button.text -match "はい|閉じる") {
                  $buttonToClick = $button
                  $action = "runtime_close"
                  break
                }
              }
            }
          }
          default {
            foreach ($button in $buttons) {
              if ($button.text -match "(?i)^(OK|Close)$" -or $button.text -match "はい|閉じる") {
                $buttonToClick = $button
                $action = "compile_close"
                break
              }
            }
          }
        }

        if (-not $looksLikeCompileDialog -and -not $looksLikeRuntimeDialog) {
          continue
        }
        if ((Test-XlflowAllowDialogFirstButtonFallback -DialogKind $DialogKind -MatchedKind $matchedKind) -and $null -eq $buttonToClick -and $buttons.Count -gt 0) {
          $buttonToClick = $buttons[0]
          if ([string]::IsNullOrWhiteSpace($action)) {
            $action = $matchedKind + "_first_button"
          }
        }

        if ($null -ne $buttonToClick) {
          [void][XlflowNativeMethods]::SendMessageW([IntPtr]([int64]$buttonToClick.hwnd), $bmClick, [IntPtr]::Zero, [IntPtr]::Zero)
        } else {
          [void][XlflowNativeMethods]::PostMessageW($hwnd, $wmClose, [IntPtr]::Zero, [IntPtr]::Zero)
          if ([string]::IsNullOrWhiteSpace($action)) {
            $action = $matchedKind + "_close"
          }
        }

        $selection = [ordered]@{
          location = [ordered]@{
            module = ""
            line = 0
            column = 0
            end_line = 0
            end_column = 0
            token = ""
          }
          nearby_code = @()
        }
        $breakModeReset = $false
        if ($matchedKind -eq "runtime" -and $action -eq "runtime_debug") {
          Start-Sleep -Milliseconds ([Math]::Max(50, $PollMs))
          try {
            $debugCapture = Invoke-XlflowRuntimeDebugSelectionCaptureProcess -CommonScriptPath $CommonScriptPath -ProcessId $TargetProcessId -WaitMilliseconds 5000 -PollMilliseconds ([Math]::Max(10, $PollMs))
            $selection = $debugCapture.selection
            $breakModeReset = [bool]$debugCapture.break_mode_reset
          } catch {
            Write-Verbose ("failed to capture runtime debug selection in watcher: " + $_.Exception.Message)
          }
        }

        return [pscustomobject][ordered]@{
          found = $true
          kind = $matchedKind
          hwnd = [int64]$hwnd
          title = $title
          class_name = $className
          text = @($staticTexts.ToArray())
          buttons = @($buttons | ForEach-Object { $_.text })
          children = @($childInfos.ToArray())
          clicked_button = $(if ($null -ne $buttonToClick) { [string]$buttonToClick.text } else { "" })
          action = $action
          selection = $selection
          break_mode_reset = $breakModeReset
        }
      }
      Start-Sleep -Milliseconds $PollMs
    }

    return [pscustomobject][ordered]@{
      found = $false
      kind = $DialogKind
      hwnd = 0
      title = ""
      class_name = ""
      text = @()
      buttons = @()
      children = @()
      clicked_button = ""
      action = ""
      selection = [ordered]@{
        location = [ordered]@{
          module = ""
          line = 0
          column = 0
          end_line = 0
          end_column = 0
          token = ""
        }
        nearby_code = @()
      }
      break_mode_reset = $false
    }
  })
  $null = $ps.AddArgument($ProcessId)
  $null = $ps.AddArgument($Kind)
  $null = $ps.AddArgument($TimeoutMilliseconds)
  $null = $ps.AddArgument($PollMilliseconds)
  $null = $ps.AddArgument($commonScriptPath)

  return [pscustomobject][ordered]@{
    powershell = $ps
    async = $ps.BeginInvoke()
  }
}

function Receive-XlflowExcelDialogWatcher {
  param(
    $Watcher,
    [int]$WaitMilliseconds = 250
  )

  if ($null -eq $Watcher) {
    return (New-XlflowExcelDialogWatcherResult)
  }

  try {
    if ($WaitMilliseconds -gt 0 -and -not $Watcher.async.AsyncWaitHandle.WaitOne($WaitMilliseconds)) {
      try { $Watcher.powershell.Stop() } catch { Write-Verbose ("failed to stop Excel dialog watcher: " + $_.Exception.Message) }
      return (New-XlflowExcelDialogWatcherResult)
    }
    $items = @($Watcher.powershell.EndInvoke($Watcher.async))
    if ($items.Count -gt 0) {
      return $items[0]
    }
  } catch {
    Write-Verbose ("failed to receive Excel dialog watcher result: " + $_.Exception.Message)
  } finally {
    try { $Watcher.powershell.Dispose() } catch { Write-Verbose ("failed to dispose Excel dialog watcher: " + $_.Exception.Message) }
  }

  return (New-XlflowExcelDialogWatcherResult)
}

function Start-XlflowVBEDialogWatcher {
  param(
    [int]$ProcessId,
    [int]$TimeoutMilliseconds = 10000,
    [int]$PollMilliseconds = 50
  )

  return Start-XlflowExcelDialogWatcher -ProcessId $ProcessId -Kind "compile" -TimeoutMilliseconds $TimeoutMilliseconds -PollMilliseconds $PollMilliseconds
}

function Receive-XlflowVBEDialogWatcher {
  param(
    $Watcher,
    [int]$WaitMilliseconds = 250
  )

  return Receive-XlflowExcelDialogWatcher -Watcher $Watcher -WaitMilliseconds $WaitMilliseconds
}

function Invoke-XlflowExcelCallWithDialogWatch {
  param(
    $Excel,
    $Workbook,
    [scriptblock]$Invocation,
    [string]$DialogKind = "runtime",
    [bool]$CaptureDialogs = $true,
    [int]$WaitMilliseconds = 250
  )

  $watcher = $null
  $caught = $null
  $value = $null
  $dialog = New-XlflowExcelDialogWatcherResult
  $selection = [ordered]@{
    location = [ordered]@{
      module = ""
      line = 0
      column = 0
      end_line = 0
      end_column = 0
      token = ""
    }
    nearby_code = @()
  }

  try {
    if ($CaptureDialogs) {
      $processId = Get-XlflowExcelProcessId -Excel $Excel
      $watcher = Start-XlflowExcelDialogWatcher -ProcessId $processId -Kind $DialogKind
    }
    try {
      $value = & $Invocation
    } catch {
      $caught = $_
    }
  } finally {
    $dialog = Receive-XlflowExcelDialogWatcher -Watcher $watcher -WaitMilliseconds $WaitMilliseconds
    if ($null -ne $Workbook -and [bool]$dialog.found) {
      try {
        if ([string]$dialog.kind -eq "runtime" -and [string]$dialog.action -eq "runtime_debug") {
          if ($null -ne $dialog.selection -and (Test-XlflowSelectionDiagnosticHasMeaningfulData -Selection $dialog.selection)) {
            $selection = $dialog.selection
          } else {
            $selection = Get-XlflowVBERuntimeSelectionDiagnostic -VBE $Workbook.VBProject.VBE
          }
          if (-not [bool]$dialog.break_mode_reset) {
            [void](Exit-XlflowVBEBreakMode -VBE $Workbook.VBProject.VBE)
          }
        } elseif ($null -ne $dialog.selection -and (Test-XlflowSelectionDiagnosticHasMeaningfulData -Selection $dialog.selection)) {
          $selection = $dialog.selection
        } else {
          $selection = Get-XlflowVBESelectionDiagnostic -VBE $Workbook.VBProject.VBE
        }
      } catch {
        Write-Verbose ("failed to capture VBE selection after Excel dialog: " + $_.Exception.Message)
      }
    }
  }

  return [pscustomobject][ordered]@{
    value = $value
    exception = $caught
    dialog = $dialog
    selection = $selection
  }
}

function Get-XlflowVBECompileControl {
  param($VBE)

  $captions = @("Compile VBAProject", "Compile Project", "コンパイル", "VBAProject のコンパイル")
  $compileControlId = 578
  try {
    $commandBars = $VBE.CommandBars

    try {
      $control = $commandBars.FindControl($null, $compileControlId)
      if ($null -ne $control) {
        return $control
      }
    } catch {
      Write-Verbose ("failed to resolve VBE compile control by id: " + $_.Exception.Message)
    }

    try {
      $menuBar = $commandBars.Item("Menu Bar")
      foreach ($menu in @($menuBar.Controls)) {
        try {
          $menuCaption = ([string]$menu.Caption).Replace("&", "")
          if ($menuCaption -notmatch "^(?i:debug)$" -and $menuCaption -notmatch "^デバッグ$") {
            continue
          }
          foreach ($control in @($menu.Controls)) {
            try {
              if ([int]$control.Id -eq $compileControlId) {
                return $control
              }
            } catch {
              Write-Verbose ("failed to inspect VBE compile control id in menu bar: " + $_.Exception.Message)
            }
            try {
              $caption = ([string]$control.Caption).Replace("&", "")
              if ($caption -match "(?i)^compile\b" -or $caption -match "コンパイル") {
                return $control
              }
            } catch {
              Write-Verbose ("failed to inspect VBE compile control caption in menu bar: " + $_.Exception.Message)
            }
          }
        } catch {
          Write-Verbose ("failed to inspect VBE menu bar popup: " + $_.Exception.Message)
        }
      }
    } catch {
      Write-Verbose ("VBE menu bar was not found: " + $_.Exception.Message)
    }

    foreach ($barName in @("Debug", "デバッグ")) {
      try {
        $bar = $commandBars.Item($barName)
        foreach ($caption in $captions) {
          try {
            return $bar.Controls.Item($caption)
          } catch {
            Write-Verbose ("compile control was not found by caption " + $caption + ": " + $_.Exception.Message)
          }
        }
        foreach ($control in @($bar.Controls)) {
          try {
            $caption = ([string]$control.Caption).Replace("&", "")
            if ([int]$control.Id -eq $compileControlId -or $caption -match "(?i)\bcompile\b" -or $caption -match "コンパイル") {
              return $control
            }
          } catch {
            Write-Verbose ("failed to inspect VBE Debug control: " + $_.Exception.Message)
          }
        }
      } catch {
        Write-Verbose ("VBE command bar was not found: " + $_.Exception.Message)
      }
    }
  } catch {
    Write-Verbose ("failed to inspect VBE command bars: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowVBEResetControl {
  param($VBE)

  try {
    $commandBars = $VBE.CommandBars
    foreach ($barName in @("Run", "実行")) {
      try {
        $bar = $commandBars.Item($barName)
        foreach ($caption in @("Reset", "リセット")) {
          try {
            return $bar.Controls.Item($caption)
          } catch {
            Write-Verbose ("reset control was not found by caption " + $caption + ": " + $_.Exception.Message)
          }
        }
        foreach ($control in @($bar.Controls)) {
          try {
            $caption = ([string]$control.Caption).Replace("&", "")
            if ($caption -match "(?i)reset" -or $caption -match "リセット") {
              return $control
            }
          } catch {
            Write-Verbose ("failed to inspect VBE Run control: " + $_.Exception.Message)
          }
        }
      } catch {
        Write-Verbose ("VBE run command bar was not found: " + $_.Exception.Message)
      }
    }
  } catch {
    Write-Verbose ("failed to inspect VBE run command bars: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowVBESelectionDiagnostic {
  param(
    $VBE,
    [bool]$PreferUserCode = $false
  )

  function New-XlflowEmptySelectionDiagnostic {
    return [ordered]@{
      location = [ordered]@{
        module = ""
        line = 0
        column = 0
        end_line = 0
        end_column = 0
        token = ""
      }
      nearby_code = @()
    }
  }

  function Get-XlflowCodePaneSelectionDiagnostic {
    param($Pane)

    $selection = New-XlflowEmptySelectionDiagnostic
    if ($null -eq $Pane) {
      return $selection
    }

    $location = $selection.location
    $nearby = @()

    try {
      $module = $Pane.CodeModule
      if ($null -ne $module) {
        $location.module = [string]$module.Name
      }
      $startLine = 0
      $startColumn = 0
      $endLine = 0
      $endColumn = 0
      $Pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
      $location.line = [int]$startLine
      $location.column = [int]$startColumn
      $location.end_line = [int]$endLine
      $location.end_column = [int]$endColumn

      if ($null -ne $module -and $startLine -gt 0 -and $startLine -le $module.CountOfLines) {
        $lineText = [string]$module.Lines($startLine, 1)
        if ($startColumn -gt 0 -and $startColumn -le $lineText.Length) {
          $tokenStart = $startColumn - 1
          $tokenLength = 1
          if ($endLine -eq $startLine -and $endColumn -gt $startColumn) {
            $tokenLength = [Math]::Min($endColumn - $startColumn, $lineText.Length - $tokenStart)
          } else {
            while ($tokenStart -gt 0 -and $lineText[$tokenStart - 1] -match "[A-Za-z0-9_]") {
              $tokenStart--
            }
            $tokenEnd = $startColumn - 1
            while ($tokenEnd -lt $lineText.Length -and $lineText[$tokenEnd] -match "[A-Za-z0-9_]") {
              $tokenEnd++
            }
            $tokenLength = [Math]::Max(1, $tokenEnd - $tokenStart)
          }
          $location.token = $lineText.Substring($tokenStart, $tokenLength).Trim()
        }

        $first = [Math]::Max(1, $startLine - 2)
        $last = [Math]::Min($module.CountOfLines, $startLine + 2)
        $items = New-Object System.Collections.Generic.List[string]
        for ($lineNo = $first; $lineNo -le $last; $lineNo++) {
          $prefix = "  "
          if ($lineNo -eq $startLine) {
            $prefix = "> "
          }
          $items.Add($prefix + $lineNo + " | " + [string]$module.Lines($lineNo, 1))
          if ($lineNo -eq $startLine -and $startColumn -gt 0) {
            $caretWidth = 1
            if ($endLine -eq $startLine -and $endColumn -gt $startColumn) {
              $caretWidth = [Math]::Max(1, $endColumn - $startColumn)
            } elseif (-not [string]::IsNullOrWhiteSpace($location.token)) {
              $caretWidth = $location.token.Length
            }
            $items.Add("    | " + (" " * [Math]::Max(0, $startColumn - 1)) + ("^" * $caretWidth))
          }
        }
        $nearby = @($items.ToArray())
      }
    } catch {
      Write-Verbose ("failed to read code pane selection diagnostic: " + $_.Exception.Message)
    }

    return [ordered]@{
      location = $location
      nearby_code = @($nearby)
    }
  }

  $location = [ordered]@{
    module = ""
    line = 0
    column = 0
    end_line = 0
    end_column = 0
    token = ""
  }
  $nearby = @()

  try {
    $bestSelection = New-XlflowEmptySelectionDiagnostic
    $bestScore = Get-XlflowSelectionDiagnosticScore -Selection $bestSelection

    $activePane = $null
    try {
      $activePane = $VBE.ActiveCodePane
    } catch {
      $activePane = $null
    }

    foreach ($candidate in @([pscustomobject]@{ pane = $activePane; active = $true }) + @($VBE.CodePanes | ForEach-Object { [pscustomobject]@{ pane = $_; active = $false } })) {
      $selection = Get-XlflowCodePaneSelectionDiagnostic -Pane $candidate.pane
      $score = Get-XlflowSelectionDiagnosticScore -Selection $selection -ActivePane ([bool]$candidate.active) -PreferUserCode $PreferUserCode
      if ($score -gt $bestScore) {
        $bestSelection = $selection
        $bestScore = $score
      }
    }
    if (Test-XlflowSelectionDiagnosticHasMeaningfulData -Selection $bestSelection) {
      return $bestSelection
    }
  } catch {
    Write-Verbose ("failed to read VBE selection diagnostic: " + $_.Exception.Message)
  }

  return [ordered]@{
    location = $location
    nearby_code = @($nearby)
  }
}

function Test-XlflowSelectionDiagnosticHasMeaningfulData {
  param($Selection)

  if ($null -eq $Selection) {
    return $false
  }
  if ($null -ne $Selection.location) {
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.location.module)) {
      return $true
    }
    if ([int]$Selection.location.line -gt 0) {
      return $true
    }
    if ([int]$Selection.location.column -gt 0) {
      return $true
    }
  }
  return @($Selection.nearby_code).Count -gt 0
}

function Test-XlflowSelectionTargetsTemporaryRunHarness {
  param($Selection)

  if ($null -eq $Selection -or $null -eq $Selection.location) {
    return $false
  }
  $module = [string]$Selection.location.module
  return $module -like "XlflowRun_*"
}

function Get-XlflowSelectionDiagnosticCurrentLineText {
  param($Selection)

  if ($null -eq $Selection) {
    return ""
  }
  foreach ($line in @($Selection.nearby_code)) {
    $match = [regex]::Match([string]$line, '^\>\s+\d+\s+\|\s?(?<text>.*)$')
    if ($match.Success) {
      return [string]$match.Groups["text"].Value
    }
  }
  return ""
}

function Test-XlflowSelectionDiagnosticLooksStructural {
  param([string]$LineText)

  if ([string]::IsNullOrWhiteSpace($LineText)) {
    return $false
  }

  return (
    $LineText -match '^\s*(Attribute|Option)\b' -or
    $LineText -match '^\s*(Public|Private|Friend)\s+(Sub|Function|Property)\b' -or
    $LineText -match '^\s*End\s+(Sub|Function|Property|If|With|Select|Type|Enum)\b' -or
    $LineText -match '^\s*Exit\s+(Sub|Function|Property)\b'
  )
}

function Get-XlflowSelectionDiagnosticScore {
  param(
    $Selection,
    [bool]$ActivePane = $false,
    [bool]$PreferUserCode = $false
  )

  if (-not (Test-XlflowSelectionDiagnosticHasMeaningfulData -Selection $Selection)) {
    return -10000
  }

  $score = 0
  if ($ActivePane) {
    $score += 10
  }
  if (Test-XlflowSelectionTargetsTemporaryRunHarness -Selection $Selection) {
    $score -= 500
  }
  if ($null -ne $Selection.location) {
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.location.module)) {
      $score += 50
    }
    if ([int]$Selection.location.line -gt 0) {
      $score += 20
    }
    if ([int]$Selection.location.column -gt 0) {
      $score += 10
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.location.token)) {
      $score += 15
    }
  }
  if (@($Selection.nearby_code).Count -gt 0) {
    $score += 20
  }

  $lineText = Get-XlflowSelectionDiagnosticCurrentLineText -Selection $Selection
  if (-not [string]::IsNullOrWhiteSpace($lineText)) {
    $score += 25
    if ($PreferUserCode) {
      if (Test-XlflowSelectionDiagnosticLooksStructural -LineText $lineText) {
        $score -= 60
      } else {
        $score += 60
      }
      if ($lineText -match "^\s*('|Rem\b)") {
        $score -= 25
      }
    }
  }

  return $score
}

function Get-XlflowVBERuntimeSelectionDiagnostic {
  param(
    $VBE,
    [int]$WaitMilliseconds = 1500,
    [int]$PollMilliseconds = 50
  )

  $bestSelection = Get-XlflowVBESelectionDiagnostic -VBE $VBE -PreferUserCode $true
  $bestScore = Get-XlflowSelectionDiagnosticScore -Selection $bestSelection -PreferUserCode $true
  if ($bestScore -ge 120 -and -not (Test-XlflowSelectionTargetsTemporaryRunHarness -Selection $bestSelection)) {
    return $bestSelection
  }
  $deadline = (Get-Date).AddMilliseconds([Math]::Max(0, $WaitMilliseconds))
  while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds ([Math]::Max(1, $PollMilliseconds))
    $selection = Get-XlflowVBESelectionDiagnostic -VBE $VBE -PreferUserCode $true
    $score = Get-XlflowSelectionDiagnosticScore -Selection $selection -PreferUserCode $true
    if ($score -gt $bestScore) {
      $bestSelection = $selection
      $bestScore = $score
    }
    if ($score -ge 120 -and -not (Test-XlflowSelectionTargetsTemporaryRunHarness -Selection $selection)) {
      return $selection
    }
  }

  return $bestSelection
}

function Exit-XlflowVBEBreakMode {
  param($VBE)

  try {
    $control = Get-XlflowVBEResetControl -VBE $VBE
    if ($null -eq $control) {
      return $false
    }
    $control.Execute()
    return $true
  } catch {
    Write-Verbose ("failed to reset VBE break mode: " + $_.Exception.Message)
  }
  return $false
}

function Get-XlflowRuntimeDebugSelectionByProcessId {
  param(
    [int]$ProcessId,
    [int]$WaitMilliseconds = 5000,
    [int]$PollMilliseconds = 50
  )

  $result = [ordered]@{
    selection = [ordered]@{
      location = [ordered]@{
        module = ""
        line = 0
        column = 0
        end_line = 0
        end_column = 0
        token = ""
      }
      nearby_code = @()
    }
    break_mode_reset = $false
  }

  if ($ProcessId -le 0) {
    return $result
  }

  try {
    Add-XlflowNativeMethods
    $iid = [Guid]"00020400-0000-0000-C000-000000000046"
    foreach ($hwnd in [XlflowNativeMethods]::GetWindowsForProcess([uint32]$ProcessId)) {
      foreach ($candidateHwnd in @($hwnd) + @([XlflowNativeMethods]::GetChildWindows($hwnd))) {
        $dispatch = $null
        $hr = [XlflowNativeMethods]::AccessibleObjectFromWindow($candidateHwnd, 4294967280, [ref]$iid, [ref]$dispatch)
        if ($hr -ne 0 -or $null -eq $dispatch) {
          continue
        }

        $vbe = $null
        try {
          $vbe = $dispatch.Application.VBE
        } catch {
          try {
            $vbe = $dispatch.VBE
          } catch {
            $vbe = $null
          }
        }
        if ($null -eq $vbe) {
          continue
        }

        try {
          $selection = Get-XlflowVBERuntimeSelectionDiagnostic -VBE $vbe -WaitMilliseconds $WaitMilliseconds -PollMilliseconds $PollMilliseconds
          if (Test-XlflowSelectionDiagnosticHasMeaningfulData -Selection $selection) {
            $result.selection = $selection
            $result.break_mode_reset = Exit-XlflowVBEBreakMode -VBE $vbe
            return $result
          }
        } catch {
          Write-Verbose ("failed to collect runtime debug selection for process " + $ProcessId + ": " + $_.Exception.Message)
        }
      }
    }
  } catch {
    Write-Verbose ("failed to inspect runtime debug selection for process " + $ProcessId + ": " + $_.Exception.Message)
  }

  return $result
}

function Invoke-XlflowRuntimeDebugSelectionCaptureProcess {
  param(
    [string]$CommonScriptPath,
    [int]$ProcessId,
    [int]$WaitMilliseconds = 5000,
    [int]$PollMilliseconds = 50
  )

  $empty = [ordered]@{
    selection = [ordered]@{
      location = [ordered]@{
        module = ""
        line = 0
        column = 0
        end_line = 0
        end_column = 0
        token = ""
      }
      nearby_code = @()
    }
    break_mode_reset = $false
  }

  if ([string]::IsNullOrWhiteSpace($CommonScriptPath) -or $ProcessId -le 0) {
    return $empty
  }

  $escapedPath = $CommonScriptPath.Replace("'", "''")
  $command = "& { . '" + $escapedPath + "'; `$capture = Get-XlflowRuntimeDebugSelectionByProcessId -ProcessId " + $ProcessId + " -WaitMilliseconds " + $WaitMilliseconds + " -PollMilliseconds " + $PollMilliseconds + "; `$capture | ConvertTo-Json -Depth 6 -Compress }"
  try {
    $json = & powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -Command $command
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace([string]$json)) {
      return $empty
    }
    return ($json | ConvertFrom-Json)
  } catch {
    Write-Verbose ("failed to capture runtime debug selection through external PowerShell: " + $_.Exception.Message)
  }

  return $empty
}

function New-XlflowExcelMacroRunnerResult {
  return [ordered]@{
    completed = $false
    ok = $false
    value = $null
    error = $null
  }
}

function Start-XlflowExcelMacroRunnerProcess {
  param(
    [string]$CommonScriptPath,
    [int]$ProcessId,
    [string]$MacroReference
  )

  if ([string]::IsNullOrWhiteSpace($CommonScriptPath) -or $ProcessId -le 0 -or [string]::IsNullOrWhiteSpace($MacroReference)) {
    return $null
  }

  $outputPath = Join-Path ([System.IO.Path]::GetTempPath()) ("xlflow-macro-run-" + [guid]::NewGuid().ToString("N") + ".json")
  $escapedPath = $CommonScriptPath.Replace("'", "''")
  $escapedOutputPath = $outputPath.Replace("'", "''")
  $escapedMacroReference = $MacroReference.Replace("'", "''")
  $command = "& { . '" + $escapedPath + "'; `$result = [ordered]@{ ok = `$true; value = `$null; error = `$null }; try { `$excel = Get-XlflowExcelByProcessId -ProcessId " + $ProcessId + "; if (`$null -eq `$excel) { throw 'xlflow could not reconnect to the target Excel instance.' }; `$result.value = `$excel.Run('" + $escapedMacroReference + "') } catch { `$result.ok = `$false; `$result.error = [ordered]@{ message = [string]`$PSItem.Exception.Message; source = [string]`$PSItem.Exception.Source; hresult = [int]`$PSItem.Exception.HResult } }; `$result | ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath '" + $escapedOutputPath + "' -Encoding UTF8 }"

  try {
    $process = Start-Process -FilePath "powershell.exe" -ArgumentList @("-STA", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $command) -WindowStyle Hidden -PassThru
    return [pscustomobject][ordered]@{
      process = $process
      output_path = $outputPath
    }
  } catch {
    Write-Verbose ("failed to start macro runner process: " + $_.Exception.Message)
    if (Test-Path -LiteralPath $outputPath) {
      Remove-Item -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue
    }
  }

  return $null
}

function Test-XlflowExcelMacroRunnerProcessExited {
  param($Runner)

  if ($null -eq $Runner -or $null -eq $Runner.process) {
    return $true
  }
  try {
    return [bool]$Runner.process.HasExited
  } catch {
    Write-Verbose ("failed to inspect macro runner process exit state: " + $_.Exception.Message)
    return $true
  }
}

function Receive-XlflowExcelMacroRunnerProcess {
  param(
    $Runner,
    [int]$WaitMilliseconds = 0
  )

  $result = New-XlflowExcelMacroRunnerResult
  if ($null -eq $Runner) {
    return $result
  }

  try {
    if ($null -ne $Runner.process) {
      if ($WaitMilliseconds -gt 0) {
        [void]$Runner.process.WaitForExit($WaitMilliseconds)
      } elseif (-not [bool]$Runner.process.HasExited) {
        return $result
      }
    }

    if ($null -ne $Runner.process -and -not [bool]$Runner.process.HasExited) {
      return $result
    }

    $result.completed = $true
    if (-not [string]::IsNullOrWhiteSpace([string]$Runner.output_path) -and (Test-Path -LiteralPath $Runner.output_path)) {
      $payload = Get-Content -LiteralPath $Runner.output_path -Raw | ConvertFrom-Json
      if ($null -ne $payload) {
        $result.ok = [bool]$payload.ok
        $result.value = $payload.value
        $result.error = $payload.error
      }
    }
  } catch {
    Write-Verbose ("failed to receive macro runner process result: " + $_.Exception.Message)
  } finally {
    if ($result.completed) {
      if ($null -ne $Runner.process) {
        try { $Runner.process.Dispose() } catch { Write-Verbose ("failed to dispose macro runner process: " + $_.Exception.Message) }
      }
      if (-not [string]::IsNullOrWhiteSpace([string]$Runner.output_path) -and (Test-Path -LiteralPath $Runner.output_path)) {
        Remove-Item -LiteralPath $Runner.output_path -Force -ErrorAction SilentlyContinue
      }
    }
  }

  return $result
}

function Stop-XlflowExcelMacroRunnerProcess {
  param($Runner)

  if ($null -eq $Runner) {
    return
  }
  try {
    if ($null -ne $Runner.process -and -not [bool]$Runner.process.HasExited) {
      Stop-Process -Id $Runner.process.Id -Force -ErrorAction SilentlyContinue
      [void]$Runner.process.WaitForExit(1000)
    }
  } catch {
    Write-Verbose ("failed to stop macro runner process: " + $_.Exception.Message)
  }
}

function Invoke-XlflowExcelMacroRunWithDialogWatch {
  param(
    $Excel,
    $Workbook,
    [string]$MacroReference,
    [string]$DialogKind = "runtime",
    [bool]$CaptureDialogs = $true,
    [int]$WaitMilliseconds = 250
  )

  $dialog = New-XlflowExcelDialogWatcherResult
  $selection = [ordered]@{
    location = [ordered]@{
      module = ""
      line = 0
      column = 0
      end_line = 0
      end_column = 0
      token = ""
    }
    nearby_code = @()
  }
  $run = New-XlflowExcelMacroRunnerResult
  $watcher = $null
  $runner = $null

  try {
    $processId = Get-XlflowExcelProcessId -Excel $Excel
    if ($CaptureDialogs) {
      $watcher = Start-XlflowExcelDialogWatcher -ProcessId $processId -Kind $DialogKind
    }
    $runner = Start-XlflowExcelMacroRunnerProcess -CommonScriptPath (Join-Path $PSScriptRoot "common.ps1") -ProcessId $processId -MacroReference $MacroReference
    if ($null -eq $runner) {
      return [pscustomobject][ordered]@{
        value = $null
        error = [ordered]@{
          message = "xlflow could not start the macro runner process."
          source = "xlflow"
          hresult = 0
        }
        dialog = $dialog
        selection = $selection
      }
    }

    $dialogReady = $false
    while ($true) {
      if (-not $dialogReady -and $null -ne $watcher -and $watcher.async.AsyncWaitHandle.WaitOne(25)) {
        $dialog = Receive-XlflowExcelDialogWatcher -Watcher $watcher -WaitMilliseconds 0
        $dialogReady = $true
      }
      if (Test-XlflowExcelMacroRunnerProcessExited -Runner $runner) {
        break
      }
      if ($dialogReady -and [bool]$dialog.found) {
        Start-Sleep -Milliseconds ([Math]::Max(50, $WaitMilliseconds))
        if (-not (Test-XlflowExcelMacroRunnerProcessExited -Runner $runner)) {
          Stop-XlflowExcelMacroRunnerProcess -Runner $runner
        }
        break
      }
      Start-Sleep -Milliseconds 25
    }

    if (-not $dialogReady) {
      $dialog = Receive-XlflowExcelDialogWatcher -Watcher $watcher -WaitMilliseconds $WaitMilliseconds
      $dialogReady = $true
    }
    $run = Receive-XlflowExcelMacroRunnerProcess -Runner $runner -WaitMilliseconds $WaitMilliseconds

    if ($null -ne $Workbook -and [bool]$dialog.found) {
      try {
        if ([string]$dialog.kind -eq "runtime" -and [string]$dialog.action -eq "runtime_debug") {
          if ($null -ne $dialog.selection -and (Test-XlflowSelectionDiagnosticHasMeaningfulData -Selection $dialog.selection)) {
            $selection = $dialog.selection
          } else {
            $selection = Get-XlflowVBERuntimeSelectionDiagnostic -VBE $Workbook.VBProject.VBE
          }
          if (-not [bool]$dialog.break_mode_reset) {
            [void](Exit-XlflowVBEBreakMode -VBE $Workbook.VBProject.VBE)
          }
        } elseif ($null -ne $dialog.selection -and (Test-XlflowSelectionDiagnosticHasMeaningfulData -Selection $dialog.selection)) {
          $selection = $dialog.selection
        } else {
          $selection = Get-XlflowVBESelectionDiagnostic -VBE $Workbook.VBProject.VBE
        }
      } catch {
        Write-Verbose ("failed to capture VBE selection after Excel macro runner dialog: " + $_.Exception.Message)
      }
    }
  } finally {
    if ($null -ne $watcher) {
      try {
        $null = Receive-XlflowExcelDialogWatcher -Watcher $watcher -WaitMilliseconds 0
      } catch {
        Write-Verbose ("failed to finalize Excel dialog watcher: " + $_.Exception.Message)
      }
    }
    if ($null -ne $runner -and -not $run.completed) {
      Stop-XlflowExcelMacroRunnerProcess -Runner $runner
      $run = Receive-XlflowExcelMacroRunnerProcess -Runner $runner -WaitMilliseconds 250
    }
  }

  return [pscustomobject][ordered]@{
    value = $run.value
    error = $run.error
    dialog = $dialog
    selection = $selection
  }
}

function New-XlflowErrorPayloadException {
  param($ErrorPayload)

  if ($null -eq $ErrorPayload) {
    return $null
  }

  $message = [string]$ErrorPayload.message
  $hresult = 0
  if ($null -ne $ErrorPayload.hresult) {
    $hresult = [int]$ErrorPayload.hresult
  }
  $exception = New-Object System.Runtime.InteropServices.COMException($message, $hresult)
  if (-not [string]::IsNullOrWhiteSpace([string]$ErrorPayload.source)) {
    $exception.Source = [string]$ErrorPayload.source
  }
  return $exception
}

function Invoke-XlflowVBECompile {
  param($Excel, $Workbook)

  $result = [ordered]@{
    ok = $true
    dialog = [ordered]@{ found = $false; text = @(); buttons = @(); children = @() }
    selection = [ordered]@{ location = [ordered]@{}; nearby_code = @() }
    error = ""
  }

  $watcher = $null
  try {
    $control = Get-XlflowVBECompileControl -VBE $Workbook.VBProject.VBE
    if ($null -eq $control) {
      throw "VBE Compile command was not found."
    }

    $compileEnabled = $true
    try {
      $compileEnabled = [bool]$control.Enabled
    } catch {
      $compileEnabled = $true
    }
    if (-not $compileEnabled) {
      return $result
    }

    $processId = Get-XlflowExcelProcessId -Excel $Excel
    $watcher = Start-XlflowVBEDialogWatcher -ProcessId $processId
    $null = $control.Execute()
  } catch {
    $result.error = $_.Exception.Message
    $result.ok = $false
  } finally {
    $dialog = Receive-XlflowVBEDialogWatcher -Watcher $watcher -WaitMilliseconds 3000
    if ($null -ne $dialog) {
      $result.dialog = $dialog
    }
  }

  if ($null -ne $result.dialog -and $result.dialog.found) {
    $result.ok = $false
    try {
      $result.selection = Get-XlflowVBESelectionDiagnostic -VBE $Workbook.VBProject.VBE
    } catch {
      Write-Verbose ("failed to collect compile selection diagnostic: " + $_.Exception.Message)
    }
  }

  return $result
}

function Get-XlflowExcelByProcessId {
  param([int]$ProcessId)

  if ($ProcessId -le 0) {
    return $null
  }
  try {
    Add-XlflowNativeMethods
    $iid = [Guid]"00020400-0000-0000-C000-000000000046"
    foreach ($hwnd in [XlflowNativeMethods]::GetWindowsForProcess([uint32]$ProcessId)) {
      foreach ($candidateHwnd in @($hwnd) + @([XlflowNativeMethods]::GetChildWindows($hwnd))) {
        $dispatch = $null
        $hr = [XlflowNativeMethods]::AccessibleObjectFromWindow($candidateHwnd, 4294967280, [ref]$iid, [ref]$dispatch)
        if ($hr -ne 0 -or $null -eq $dispatch) {
          continue
        }
        $candidate = $dispatch
        try {
          $candidate = $dispatch.Application
        } catch {
          $candidate = $dispatch
        }
        try {
          if ($candidate.Workbooks.Count -gt 0) {
            return ,$candidate
          }
        } catch {
          Write-Verbose ("accessible object is not an Excel application: " + $_.Exception.Message)
        }
      }
    }
  } catch {
    Write-Verbose ("failed to resolve Excel by process id: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowWorkbookStateByProcessId {
  param([int]$ProcessId)

  if ($ProcessId -le 0) {
    return $null
  }
  $sawWorkbookFreeState = $false
  try {
    Add-XlflowNativeMethods
    $iid = [Guid]"00020400-0000-0000-C000-000000000046"
    foreach ($hwnd in [XlflowNativeMethods]::GetWindowsForProcess([uint32]$ProcessId)) {
      foreach ($candidateHwnd in @($hwnd) + @([XlflowNativeMethods]::GetChildWindows($hwnd))) {
        $dispatch = $null
        $hr = [XlflowNativeMethods]::AccessibleObjectFromWindow($candidateHwnd, 4294967280, [ref]$iid, [ref]$dispatch)
        if ($hr -ne 0 -or $null -eq $dispatch) {
          continue
        }
        $candidate = $dispatch
        try {
          $candidate = $dispatch.Application
        } catch {
          Release-XlflowComObject -Object $dispatch -Name "dispatch COM object (non-Excel)"
          continue
        }
        try {
          if ($candidate.Workbooks.Count -gt 0) {
            Release-XlflowComObject -Object $candidate -Name "candidate Excel COM object"
            Release-XlflowComObject -Object $dispatch -Name "dispatch COM object"
            return $true
          }
          $sawWorkbookFreeState = $true
          Release-XlflowComObject -Object $candidate -Name "candidate Excel COM object"
          Release-XlflowComObject -Object $dispatch -Name "dispatch COM object"
          continue
        } catch {
          Release-XlflowComObject -Object $candidate -Name "candidate COM object (non-Excel)"
          Release-XlflowComObject -Object $dispatch -Name "dispatch COM object (non-Excel)"
          Write-Verbose ("accessible object is not an Excel application: " + $_.Exception.Message)
        }
      }
    }
  } catch {
    Write-Verbose ("failed to resolve workbook state by process id: " + $_.Exception.Message)
  }
  if ($sawWorkbookFreeState) {
    return $false
  }
  return $null
}

function Get-XlflowVBEByProcessId {
  param([int]$ProcessId)

  if ($ProcessId -le 0) {
    return $null
  }
  try {
    Add-XlflowNativeMethods
    $iid = [Guid]"00020400-0000-0000-C000-000000000046"
    foreach ($hwnd in [XlflowNativeMethods]::GetWindowsForProcess([uint32]$ProcessId)) {
      foreach ($candidateHwnd in @($hwnd) + @([XlflowNativeMethods]::GetChildWindows($hwnd))) {
        $dispatch = $null
        $hr = [XlflowNativeMethods]::AccessibleObjectFromWindow($candidateHwnd, 4294967280, [ref]$iid, [ref]$dispatch)
        if ($hr -ne 0 -or $null -eq $dispatch) {
          continue
        }
        $candidate = $dispatch
        try {
          $candidate = $dispatch.Application.VBE
        } catch {
          try {
            $candidate = $dispatch.VBE
          } catch {
            $candidate = $dispatch
          }
        }
        try {
          $null = $candidate.ActiveCodePane
          $null = $candidate.CodePanes
          $null = $candidate.CommandBars
          return ,$candidate
        } catch {
          Write-Verbose ("accessible object is not a VBE automation object: " + $_.Exception.Message)
        }
      }
    }
  } catch {
    Write-Verbose ("failed to resolve VBE by process id: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowExcelByHwnd {
  param([int64]$Hwnd)

  if ($Hwnd -eq 0) {
    return $null
  }
  try {
    Add-XlflowNativeMethods
    $iid = [Guid]"00020400-0000-0000-C000-000000000046"
    foreach ($candidateHwnd in @([IntPtr]$Hwnd) + @([XlflowNativeMethods]::GetChildWindows([IntPtr]$Hwnd))) {
      $dispatch = $null
      $hr = [XlflowNativeMethods]::AccessibleObjectFromWindow($candidateHwnd, 4294967280, [ref]$iid, [ref]$dispatch)
      if ($hr -ne 0 -or $null -eq $dispatch) {
        continue
      }
      $candidate = $dispatch
      try {
        $candidate = $dispatch.Application
      } catch {
        $candidate = $dispatch
      }
      try {
        if ($candidate.Workbooks.Count -gt 0) {
          return ,$candidate
        }
      } catch {
        Write-Verbose ("accessible object is not an Excel application: " + $_.Exception.Message)
      }
    }
  } catch {
    Write-Verbose ("failed to resolve Excel by hwnd: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowExcelFromSessionMetadata {
  param([string]$MetadataPath)

  if ([string]::IsNullOrWhiteSpace($MetadataPath) -or -not (Test-Path -LiteralPath $MetadataPath)) {
    return $null
  }
  try {
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
    if ($null -ne $metadata -and $null -ne $metadata.hwnd -and $metadata.hwnd -ne 0) {
      $excel = Get-XlflowExcelByHwnd -Hwnd ([int64]$metadata.hwnd)
      if ($null -ne $excel) {
        return $excel
      }
    }
    if ($null -ne $metadata -and $metadata.pid -gt 0) {
      return Get-XlflowExcelByProcessId -ProcessId ([int]$metadata.pid)
    }
  } catch {
    Write-Verbose ("failed to read session metadata for Excel lookup: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowSessionExcel {
  param([string]$MetadataPath)

  $excel = Get-XlflowExcelFromSessionMetadata -MetadataPath $MetadataPath
  if ($null -ne $excel) {
    return $excel
  }
  return Get-XlflowActiveExcel
}

function Test-XlflowSessionMetadataMatchesWorkbook {
  param([string]$MetadataPath, [string]$WorkbookPath)

  if ([string]::IsNullOrWhiteSpace($MetadataPath) -or [string]::IsNullOrWhiteSpace($WorkbookPath) -or -not (Test-Path -LiteralPath $MetadataPath)) {
    return $false
  }
  try {
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
    if ($null -eq $metadata -or [string]::IsNullOrWhiteSpace([string]$metadata.workbook_path)) {
      return $false
    }
    return [System.IO.Path]::GetFullPath([string]$metadata.workbook_path) -ieq [System.IO.Path]::GetFullPath($WorkbookPath)
  } catch {
    Write-Verbose ("failed to compare session metadata workbook path: " + $_.Exception.Message)
    return $false
  }
}

function Get-XlflowWorkbookDirtyState {
  param($Workbook)

  if ($null -eq $Workbook) {
    return $null
  }
  try {
    return -not [bool]$Workbook.Saved
  } catch {
    Write-Verbose ("failed to inspect workbook dirty state: " + $_.Exception.Message)
    return $null
  }
}

function Get-XlflowWorkbookSaveState {
  param($Workbook, [bool]$SessionAttached = $false)

  $dirty = $false
  $needsSave = $false
  if ($SessionAttached) {
    $dirtyState = Get-XlflowWorkbookDirtyState -Workbook $Workbook
    if ($null -ne $dirtyState) {
      $dirty = [bool]$dirtyState
      $needsSave = [bool]$dirtyState
    } else {
      $dirty = $true
      $needsSave = $true
    }
  }

  return [ordered]@{
    dirty = $dirty
    needs_save = $needsSave
  }
}

function New-XlflowWorkbookResult {
  param(
    [string]$WorkbookPath,
    [bool]$SessionAttached = $false,
    [string]$SessionMode = "none",
    [AllowNull()]$Saved = $null,
    [string]$SaveAsPath = "",
    [AllowNull()]$Dirty = $null,
    [AllowNull()]$NeedsSave = $null,
    [AllowNull()]$DirtyBeforeStop = $null,
    [AllowNull()]$AutoSavedOnStop = $null
  )

  $workbook = [ordered]@{
    path = $WorkbookPath
    session = $SessionAttached
    session_mode = $SessionMode
    session_requested = ($SessionMode -eq "explicit")
    auto_session = ($SessionMode -eq "auto")
  }
  if ($PSBoundParameters.ContainsKey("Saved") -and $null -ne $Saved) {
    $workbook.saved = [bool]$Saved
  }
  if ($PSBoundParameters.ContainsKey("SaveAsPath")) {
    if ([string]::IsNullOrWhiteSpace($SaveAsPath)) {
      $workbook.save_as = $null
    } else {
      $workbook.save_as = $SaveAsPath
    }
  }
  if ($PSBoundParameters.ContainsKey("Dirty") -and $null -ne $Dirty) {
    $workbook.dirty = [bool]$Dirty
  }
  if ($PSBoundParameters.ContainsKey("NeedsSave") -and $null -ne $NeedsSave) {
    $workbook.needs_save = [bool]$NeedsSave
  }
  if ($PSBoundParameters.ContainsKey("DirtyBeforeStop") -and $null -ne $DirtyBeforeStop) {
    $workbook.dirty_before_stop = [bool]$DirtyBeforeStop
  }
  if ($PSBoundParameters.ContainsKey("AutoSavedOnStop") -and $null -ne $AutoSavedOnStop) {
    $workbook.auto_saved_on_stop = [bool]$AutoSavedOnStop
  }
  return $workbook
}

function Get-XlflowSessionUsageLog {
  param([string]$SessionMode)

  switch ($SessionMode) {
    "explicit" { return "using xlflow session workbook (--session)" }
    "auto" { return "auto-reused matching xlflow session workbook" }
    "managed" { return "using managed xlflow session workbook" }
    default { return $null }
  }
}

function Open-XlflowWorkbookForCommand {
  param(
    [string]$WorkbookPath,
    [string]$Visible = "false",
    [string]$DisplayAlerts = "false",
    [string]$DisableAutomationMacros = "true",
    [string]$CaptureOpenVBADialogs = "false",
    [string]$UseSession = "false",
    [string]$MetadataPath = "",
    [bool]$AllowIsolatedOpen = $true,
    [int]$OpenDialogWaitMilliseconds = 1500
  )

  if (ConvertTo-XlflowBool $UseSession) {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
    return [pscustomobject][ordered]@{
      excel = $excel
      workbook = $workbook
      session_attached = $true
      session_mode = "explicit"
    }
  }

  if (Test-XlflowSessionMetadataMatchesWorkbook -MetadataPath $MetadataPath -WorkbookPath $WorkbookPath) {
    $excel = Get-XlflowExcelFromSessionMetadata -MetadataPath $MetadataPath
    if ($null -ne $excel) {
      try {
        $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
        return [pscustomobject][ordered]@{
          excel = $excel
          workbook = $workbook
          session_attached = $true
          session_mode = "auto"
        }
      } catch {
        Write-Verbose ("failed to attach to matching xlflow session workbook: " + $_.Exception.Message)
      }
    }
  }

  if (-not $AllowIsolatedOpen) {
    throw "no matching xlflow session workbook is running; run xlflow session start or use the configured workbook session"
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $workbook = $null
  $openDialog = New-XlflowExcelDialogWatcherResult
  $openSelection = [ordered]@{
    location = [ordered]@{
      module = ""
      line = 0
      column = 0
      end_line = 0
      end_column = 0
      token = ""
    }
    nearby_code = @()
  }
  if (ConvertTo-XlflowBool $CaptureOpenVBADialogs) {
    $openResult = Invoke-XlflowExcelCallWithDialogWatch -Excel $excel -Workbook $null -Invocation {
      Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts (ConvertTo-XlflowBool $DisplayAlerts) -DisableAutomationMacros (ConvertTo-XlflowBool $DisableAutomationMacros)
    } -DialogKind "any_vba" -CaptureDialogs $true -WaitMilliseconds $OpenDialogWaitMilliseconds
    $workbook = $openResult.value
    $openDialog = $openResult.dialog
    $openSelection = $openResult.selection
    if ($null -ne $openResult.exception) {
      throw $openResult.exception.Exception
    }
  } else {
    $workbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts (ConvertTo-XlflowBool $DisplayAlerts) -DisableAutomationMacros (ConvertTo-XlflowBool $DisableAutomationMacros)
  }
  return [pscustomobject][ordered]@{
    excel = $excel
    workbook = $workbook
    session_attached = $false
    session_mode = "none"
    open_dialog = $openDialog
    open_selection = $openSelection
  }
}

function Close-XlflowSessionWorkbook {
  param([string]$WorkbookPath, [string]$MetadataPath, [bool]$Save)

  $workbook = $null
  $excel = $null
  try {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
    try { $excel = $workbook.Application } catch { $excel = $null }
    $workbook.Close($Save) | Out-Null
    if ($null -ne $excel) {
      $excel.Quit() | Out-Null
    }
  } catch {
    Write-Verbose ("failed to close xlflow session workbook: " + $_.Exception.Message)
  } finally {
    if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and (Test-Path -LiteralPath $MetadataPath)) {
      Remove-Item -LiteralPath $MetadataPath -Force -ErrorAction SilentlyContinue
    }
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  }
}

function Get-XlflowFileHash {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    return ""
  }
  $stream = [System.IO.File]::OpenRead($Path)
  try {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
      $bytes = $sha.ComputeHash($stream)
      return ([System.BitConverter]::ToString($bytes) -replace "-", "").ToLowerInvariant()
    } finally {
      $sha.Dispose()
    }
  } finally {
    $stream.Dispose()
  }
}

function Get-XlflowFolderConfig {
  param(
    [string]$Folders = "true",
    [string]$FolderAnnotation = "update",
    [string]$DefaultComponentFolders = "true"
  )

  return [pscustomobject][ordered]@{
    folders = (ConvertTo-XlflowBool $Folders)
    folder_annotation = [string]$FolderAnnotation
    default_component_folders = (ConvertTo-XlflowBool $DefaultComponentFolders)
  }
}

function Get-XlflowComponentExtension {
  param([int]$ComponentType)

  switch ($ComponentType) {
    1 { return ".bas" }
    2 { return ".cls" }
    3 { return ".frm" }
    100 { return ".bas" }
    default { return "" }
  }
}

function Get-XlflowComponentTypeName {
  param([int]$ComponentType)

  switch ($ComponentType) {
    1 { return "standard_module" }
    2 { return "class_module" }
    3 { return "userform" }
    100 { return "document_module" }
    default { return "unknown" }
  }
}

function Get-XlflowComponentRootDir {
  param(
    [int]$ComponentType,
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir
  )

  switch ($ComponentType) {
    1 { return $ModulesDir }
    2 { return $ClassesDir }
    3 { return $FormsDir }
    100 { return $WorkbookDir }
    default { return "" }
  }
}

function Get-XlflowContentNewline {
  param([string]$Text)

  if ($Text -match "`r`n") {
    return "`r`n"
  }
  return "`n"
}

function ConvertTo-XlflowFolderPathSegment {
  param([string]$Segment)

  $clean = [string]$Segment
  $clean = $clean.Trim()
  if ([string]::IsNullOrWhiteSpace($clean)) {
    return ""
  }
  $clean = [regex]::Replace($clean, '[<>:"/\\|?*\x00-\x1f]', "_")
  $clean = $clean.Trim(" ", ".")
  if ($clean -eq "." -or $clean -eq "..") {
    return ""
  }
  return $clean
}

function ConvertFrom-XlflowFolderAnnotation {
  param([string]$Annotation)

  if ([string]::IsNullOrWhiteSpace($Annotation)) {
    return @()
  }

  $segments = New-Object System.Collections.Generic.List[string]
  foreach ($part in ([string]$Annotation).Split(".")) {
    $segment = ConvertTo-XlflowFolderPathSegment -Segment $part
    if ([string]::IsNullOrWhiteSpace($segment)) {
      return $null
    }
    $segments.Add($segment)
  }
  return @($segments.ToArray())
}

function ConvertTo-XlflowFolderAnnotation {
  param([string[]]$Segments)

  if ($null -eq $Segments -or $Segments.Count -eq 0) {
    return ""
  }

  $clean = New-Object System.Collections.Generic.List[string]
  foreach ($segment in $Segments) {
    $part = ConvertTo-XlflowFolderPathSegment -Segment $segment
    if ([string]::IsNullOrWhiteSpace($part)) {
      continue
    }
    $clean.Add($part)
  }
  if ($clean.Count -eq 0) {
    return ""
  }
  return ($clean -join ".")
}

function Test-XlflowFolderAnnotationHeaderLine {
  param([string]$TrimmedLine)

  if ([string]::IsNullOrWhiteSpace($TrimmedLine)) {
    return $true
  }
  if ($TrimmedLine.StartsWith("'")) {
    return $true
  }
  if ($TrimmedLine -match '^Attribute\s+VB_') {
    return $true
  }
  if ($TrimmedLine -match '^(Option\s+Explicit|VERSION\s+1\.0\s+CLASS|BEGIN\b|END\b|MultiUse\b)') {
    return $true
  }
  return $false
}

function Find-XlflowFolderAnnotation {
  param(
    [string]$Text,
    [int]$MaxLines = 50
  )

  $lines = $Text -split "`r?`n"
  $max = [Math]::Min($lines.Count, $MaxLines)
  for ($i = 0; $i -lt $max; $i++) {
    $line = [string]$lines[$i]
    $trimmed = $line.Trim()
    if ($trimmed -match "^'\s*@Folder\(""([^""]+)""\)\s*$") {
      $segments = ConvertFrom-XlflowFolderAnnotation -Annotation $matches[1]
      return [pscustomobject][ordered]@{
        found = $true
        valid = ($null -ne $segments)
        malformed = ($null -eq $segments)
        annotation = $(if ($null -ne $segments) { $matches[1] } else { "" })
        line_index = $i
      }
    }
    if ($trimmed -match "@Folder") {
      return [pscustomobject][ordered]@{
        found = $true
        valid = $false
        malformed = $true
        annotation = ""
        line_index = $i
      }
    }
    if (-not (Test-XlflowFolderAnnotationHeaderLine -TrimmedLine $trimmed)) {
      break
    }
  }

  return [pscustomobject][ordered]@{
    found = $false
    valid = $false
    malformed = $false
    annotation = ""
    line_index = -1
  }
}

function Get-XlflowFolderAnnotationInsertIndex {
  param([string[]]$Lines)

  $index = 0
  if ($Lines.Count -gt 0 -and $Lines[0].Trim() -eq "VERSION 1.0 CLASS") {
    $index = 1
    while ($index -lt $Lines.Count) {
      $trimmed = $Lines[$index].Trim()
      $index++
      if ($trimmed -eq "END") {
        break
      }
    }
  }
  while ($index -lt $Lines.Count -and $Lines[$index].Trim() -match '^Attribute\s+VB_') {
    $index++
  }
  return $index
}

function Update-XlflowFolderAnnotationText {
  param(
    [string]$Text,
    [string]$FolderAnnotationMode = "update",
    [string]$DesiredAnnotation = ""
  )

  if ($FolderAnnotationMode -eq "ignore" -or $FolderAnnotationMode -eq "preserve") {
    return $Text
  }

  $newline = Get-XlflowContentNewline -Text $Text
  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($line in ($Text -split "`r?`n")) {
    $lines.Add([string]$line)
  }
  $annotationInfo = Find-XlflowFolderAnnotation -Text $Text
  $annotationLine = ""
  if (-not [string]::IsNullOrWhiteSpace($DesiredAnnotation)) {
    $annotationLine = "'@Folder(""$DesiredAnnotation"")"
  }

  if ($annotationInfo.found) {
    if ([string]::IsNullOrWhiteSpace($annotationLine)) {
      $lines.RemoveAt([int]$annotationInfo.line_index)
    } else {
      $lines[[int]$annotationInfo.line_index] = $annotationLine
    }
  } elseif (-not [string]::IsNullOrWhiteSpace($annotationLine)) {
    $insertIndex = Get-XlflowFolderAnnotationInsertIndex -Lines @($lines.ToArray())
    $lines.Insert([int]$insertIndex, $annotationLine)
  }

  return (@($lines.ToArray()) -join $newline)
}

function Get-XlflowRelativePathSegments {
  param(
    [string]$RootDir,
    [string]$Path
  )

  if ([string]::IsNullOrWhiteSpace($RootDir) -or [string]::IsNullOrWhiteSpace($Path)) {
    return @()
  }

  $relativeDir = Split-Path -Parent (Get-XlflowRelativePath -BasePath $RootDir -TargetPath $Path)
  if ([string]::IsNullOrWhiteSpace($relativeDir) -or $relativeDir -eq ".") {
    return @()
  }

  $segments = New-Object System.Collections.Generic.List[string]
  foreach ($part in ($relativeDir -split '[\\/]')) {
    if ($part -eq "..") {
      throw "path '$Path' resolves outside root '$RootDir' (relative: '$relativeDir')"
    }
    $segment = ConvertTo-XlflowFolderPathSegment -Segment $part
    if ([string]::IsNullOrWhiteSpace($segment)) {
      continue
    }
    $segments.Add($segment)
  }
  return @($segments.ToArray())
}

function Get-XlflowFolderAnnotationForPath {
  param(
    [string]$RootDir,
    [string]$Path
  )

  return ConvertTo-XlflowFolderAnnotation -Segments (Get-XlflowRelativePathSegments -RootDir $RootDir -Path $Path)
}

function Get-XlflowFolderAnnotationForComponent {
  param(
    $Component,
    [string]$FolderAnnotationMode = "update"
  )

  if ($FolderAnnotationMode -eq "ignore") {
    return ""
  }

  try {
    $text = Get-XlflowCodeModuleText -CodeModule $Component.CodeModule
    $annotation = Find-XlflowFolderAnnotation -Text $text
    if ($annotation.valid) {
      return [string]$annotation.annotation
    }
  } catch {
    Write-Verbose ("failed to inspect folder annotation from code module: " + $_.Exception.Message)
  }
  return ""
}

function Get-XlflowSourceComponentFiles {
  param(
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir,
    [string]$CodeSource = "frm"
  )

  $files = New-Object System.Collections.Generic.List[object]
  foreach ($entry in @(
    @{ kind = "module"; dir = $ModulesDir },
    @{ kind = "class"; dir = $ClassesDir },
    @{ kind = "form"; dir = $FormsDir },
    @{ kind = "document"; dir = $WorkbookDir }
  )) {
    $dir = [string]$entry.dir
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) {
      continue
    }
    $excludedDir = ""
    if ($entry.kind -eq "form") {
      $excludedDir = Get-XlflowUserFormCodeDir -FormsDir $FormsDir
    }
    foreach ($file in Get-ChildItem -LiteralPath $dir -Recurse -File | Sort-Object FullName) {
      if (-not [string]::IsNullOrWhiteSpace($excludedDir)) {
        $excludedFullPath = [System.IO.Path]::GetFullPath($excludedDir)
        $directorySeparator = [string][System.IO.Path]::DirectorySeparatorChar
        if (-not $excludedFullPath.EndsWith($directorySeparator)) {
          $excludedFullPath += $directorySeparator
        }
        $fileFullPath = [System.IO.Path]::GetFullPath($file.FullName)
        if ($fileFullPath.StartsWith($excludedFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
          continue
        }
      }
      if ($file.Extension -notin @(".bas", ".cls", ".frm", ".frx")) {
        continue
      }
      $files.Add([pscustomobject][ordered]@{
        kind = [string]$entry.kind
        root_dir = $dir
        full_name = $file.FullName
        relative_path = (Get-XlflowRelativePath -BasePath $dir -TargetPath $file.FullName).Replace("\", "/")
        extension = [string]$file.Extension.ToLowerInvariant()
        module_name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
      })
    }
  }
  if (Use-XlflowUserFormCodeSidecar -CodeSource $CodeSource) {
    foreach ($file in @(Get-XlflowUserFormCodeFiles -FormsDir $FormsDir)) {
      $files.Add($file)
    }
  }
  return @($files.ToArray())
}

function Clear-XlflowSourceComponentFiles {
  param(
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir,
    [string]$CodeSource = "frm"
  )

  foreach ($file in @(Get-XlflowSourceComponentFiles -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir -CodeSource $CodeSource)) {
    Remove-Item -LiteralPath $file.full_name -Force -ErrorAction SilentlyContinue
  }
}

function Get-XlflowSourceFingerprint {
  param(
    [string]$WorkbookPath,
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir,
    [string]$CodeSource = "frm"
  )

  $files = New-Object System.Collections.Generic.List[object]
  foreach ($file in @(Get-XlflowSourceComponentFiles -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir -CodeSource $CodeSource)) {
    $files.Add([ordered]@{
      kind = [string]$file.kind
      path = [string]$file.relative_path
      hash = Get-XlflowFileHash -Path $file.full_name
    })
  }
  return [pscustomobject][ordered]@{
    workbook_path = [System.IO.Path]::GetFullPath($WorkbookPath)
    files = @($files.ToArray())
  }
}

function Test-XlflowFingerprintMatchesState {
  param($Fingerprint, [string]$StatePath)

  if ([string]::IsNullOrWhiteSpace($StatePath) -or -not (Test-Path -LiteralPath $StatePath)) {
    return $false
  }
  try {
    $existing = Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json
    $currentJson = $Fingerprint | ConvertTo-Json -Depth 10 -Compress
    $existingJson = $existing | ConvertTo-Json -Depth 10 -Compress
    return $currentJson -eq $existingJson
  } catch {
    return $false
  }
}

function Write-XlflowFingerprintState {
  param($Fingerprint, [string]$StatePath)

  if ([string]::IsNullOrWhiteSpace($StatePath)) {
    return
  }
  $parent = Split-Path -Parent $StatePath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }
  $Fingerprint | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StatePath -Encoding UTF8
}

function Get-XlflowComponentPath {
  param(
    $Component,
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir,
    [string]$Folders = "true",
    [string]$FolderAnnotation = "update",
    [string]$DefaultComponentFolders = "true"
  )

  $rootDir = Get-XlflowComponentRootDir -ComponentType $Component.Type -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir
  $extension = Get-XlflowComponentExtension -ComponentType $Component.Type
  if ([string]::IsNullOrWhiteSpace($rootDir) -or [string]::IsNullOrWhiteSpace($extension)) {
    return $null
  }

  $segments = @()
  $folderConfig = Get-XlflowFolderConfig -Folders $Folders -FolderAnnotation $FolderAnnotation -DefaultComponentFolders $DefaultComponentFolders
  if ($folderConfig.folders -and $folderConfig.folder_annotation -ne "ignore") {
    $annotation = Get-XlflowFolderAnnotationForComponent -Component $Component -FolderAnnotationMode $folderConfig.folder_annotation
    if (-not [string]::IsNullOrWhiteSpace($annotation)) {
      $parsed = ConvertFrom-XlflowFolderAnnotation -Annotation $annotation
      if ($null -ne $parsed) {
        $segments = $parsed
      }
    }
  }

  $path = $rootDir
  foreach ($segment in $segments) {
    $path = Join-Path $path $segment
  }
  return Join-Path $path ($Component.Name + $extension)
}

function Get-XlflowUtf8Encoding {
  return (New-Object System.Text.UTF8Encoding -ArgumentList $false)
}

function Get-XlflowCp932Encoding {
  try {
    $providerType = [type]::GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages")
    if ($null -ne $providerType) {
      $provider = $providerType.GetProperty("Instance").GetValue($null, $null)
      [System.Text.Encoding]::RegisterProvider($provider)
    }
  } catch {
    Write-Verbose ("failed to register code page provider: " + $_.Exception.Message)
  }
  return [System.Text.Encoding]::GetEncoding(932)
}

function Get-XlflowUtf8Text {
  param([string]$Path)
  return [System.IO.File]::ReadAllText($Path, (Get-XlflowUtf8Encoding))
}

function Set-XlflowUtf8Text {
  param([string]$Path, [string]$Text)
  [System.IO.File]::WriteAllText($Path, $Text, (Get-XlflowUtf8Encoding))
}

function Get-XlflowCp932Text {
  param([string]$Path)
  return [System.IO.File]::ReadAllText($Path, (Get-XlflowCp932Encoding))
}

function Set-XlflowCp932Text {
  param([string]$Path, [string]$Text)
  [System.IO.File]::WriteAllText($Path, $Text, (Get-XlflowCp932Encoding))
}

function Convert-XlflowExportedSourceToUtf8 {
  param([string]$Path)
  $content = Get-XlflowCp932Text -Path $Path
  Set-XlflowUtf8Text -Path $Path -Text $content
}

function ConvertTo-XlflowFRMStringLiteral {
  param([string]$Value)

  if ($null -eq $Value) {
    $Value = ""
  }
  return '"' + ([string]$Value).Replace('"', '""') + '"'
}

function Get-XlflowUserFormDesignerCaption {
  param($Component)

  if ($null -eq $Component) {
    return $null
  }
  try {
    if ([int]$Component.Type -ne 3) {
      return $null
    }
  } catch {
    return $null
  }
  try {
    return [string]$Component.Designer.Caption
  } catch {
    Write-Verbose ("failed to inspect UserForm designer caption: " + $_.Exception.Message)
    return $null
  }
}

function Normalize-XlflowUserFormArtifactFile {
  param(
    [string]$Path,
    [AllowNull()]
    $Caption
  )

  if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
    return
  }
  if ($null -eq $Caption) {
    return
  }
  if ([System.IO.Path]::GetExtension($Path) -ine ".frm") {
    return
  }

  $text = Get-XlflowUtf8Text -Path $Path
  if ([string]::IsNullOrEmpty($text)) {
    return
  }

  $newline = Get-XlflowContentNewline -Text $text
  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($line in ($text -split "`r?`n")) {
    $lines.Add([string]$line)
  }

  $beginIndex = -1
  $endIndex = -1
  $captionIndex = -1
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = [string]$lines[$i]
    if ($beginIndex -lt 0) {
      if ($line -match '^\s*Begin\b') {
        $beginIndex = $i
      }
      continue
    }
    if ($line -match '^\s*End\s*$') {
      $endIndex = $i
      break
    }
    if ($line -match '^\s*Caption\s*=') {
      $captionIndex = $i
    }
  }
  if ($beginIndex -lt 0 -or $endIndex -lt 0) {
    return
  }

  $captionLine = '   Caption         =   ' + (ConvertTo-XlflowFRMStringLiteral -Value $Caption)
  if ($captionIndex -ge 0) {
    $lines[$captionIndex] = $captionLine
  } else {
    $lines.Insert(($beginIndex + 1), $captionLine)
  }

  Set-XlflowUtf8Text -Path $Path -Text ((@($lines.ToArray()) -join $newline) + $newline)
}

function Copy-XlflowSourceForImport {
  param(
    [string]$SourcePath,
    [string]$DestinationPath,
    [string]$RootDir = "",
    [string]$FolderAnnotationMode = "update"
  )

  $parent = Split-Path -Parent $DestinationPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }

  if ([System.IO.Path]::GetExtension($SourcePath) -ieq ".frx") {
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    return
  }

  $content = Get-XlflowUtf8Text -Path $SourcePath
  if (-not [string]::IsNullOrWhiteSpace($RootDir)) {
    $desiredAnnotation = Get-XlflowFolderAnnotationForPath -RootDir $RootDir -Path $SourcePath
    $content = Update-XlflowFolderAnnotationText -Text $content -FolderAnnotationMode $FolderAnnotationMode -DesiredAnnotation $desiredAnnotation
  }
  Set-XlflowCp932Text -Path $DestinationPath -Text $content
}

function Write-XlflowJson {
  param([hashtable]$Result)
  $Result | ConvertTo-Json -Depth 10
}

function Find-XlflowTestProcedures {
  param([string]$ModuleName, [string]$Code)

  $tests = New-Object System.Collections.Generic.List[object]
  if ([string]::IsNullOrEmpty($Code)) {
    return $tests
  }

  $lines = $Code -split "`r?`n"
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    $match = [regex]::Match($line, '^(?:Public\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(\s*\))?\s*(?:''.*)?$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
      continue
    }
    $name = $match.Groups[1].Value
    if ($name -like "Test*" -or $name -like "*_Test") {
      $tags = New-Object System.Collections.Generic.List[string]
      for ($j = $i - 1; $j -ge 0; $j--) {
        $prev = $lines[$j].Trim()
        if ([string]::IsNullOrWhiteSpace($prev)) {
          continue
        }
        $tagMatch = [regex]::Match($prev, "^'\s*@Tag\s*\(""([^""]+)""\)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($tagMatch.Success) {
          $tags.Add($tagMatch.Groups[1].Value) | Out-Null
          continue
        }
        if ($prev -like "''*") {
          continue
        }
        break
      }
      $tests.Add([pscustomobject][ordered]@{
        name = $name
        module = $ModuleName
        line = $i + 1
        tags = @($tags.ToArray())
      })
    }
  }

  foreach ($test in $tests) {
    Write-Output $test
  }
}

function Find-XlflowModuleHooks {
  param([string]$ModuleName, [string]$Code)

  $hooks = [ordered]@{ BeforeAll = $null; AfterAll = $null; BeforeEach = $null; AfterEach = $null }
  if ([string]::IsNullOrEmpty($Code)) {
    return $hooks
  }

  $lines = $Code -split "`r?`n"
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    $match = [regex]::Match($line, '^(?:Public\s+)?Sub\s+(BeforeAll|AfterAll|BeforeEach|AfterEach)\s*(?:\(\s*\))?\s*(?:''.*)?$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
      continue
    }
    $name = $match.Groups[1].Value
    switch ($name.ToLowerInvariant()) {
      "beforeall"  { $hooks.BeforeAll  = [pscustomobject][ordered]@{ name = $name; module = $ModuleName; line = $i + 1 } }
      "afterall"   { $hooks.AfterAll   = [pscustomobject][ordered]@{ name = $name; module = $ModuleName; line = $i + 1 } }
      "beforeeach" { $hooks.BeforeEach = [pscustomobject][ordered]@{ name = $name; module = $ModuleName; line = $i + 1 } }
      "aftereach"  { $hooks.AfterEach  = [pscustomobject][ordered]@{ name = $name; module = $ModuleName; line = $i + 1 } }
    }
  }

  return $hooks
}

function Select-XlflowTests {
  param($Tests, [string]$Filter = "", [string]$ModuleFilter = "", [string]$TagFilter = "")

  $selected = New-Object System.Collections.Generic.List[object]
  foreach ($test in $Tests) {
    $include = $true
    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
      if ($test.name -ne $Filter) { $include = $false }
    }
    if (-not [string]::IsNullOrWhiteSpace($ModuleFilter)) {
      if ($test.module -ne $ModuleFilter) { $include = $false }
    }
    if (-not [string]::IsNullOrWhiteSpace($TagFilter)) {
      $tagFound = $false
      foreach ($tag in $test.tags) {
        if ($tag -ieq $TagFilter) {
          $tagFound = $true
          break
        }
      }
      if (-not $tagFound) { $include = $false }
    }
    if ($include) {
      $selected.Add($test) | Out-Null
    }
  }
  foreach ($test in $selected) {
    Write-Output $test
  }
}

function Get-XlflowCodeModuleText {
  param($CodeModule)

  if ($null -eq $CodeModule -or $CodeModule.CountOfLines -le 0) {
    return ""
  }
  return $CodeModule.Lines(1, $CodeModule.CountOfLines)
}

function Set-XlflowCodeModuleText {
  param(
    $CodeModule,
    [string]$Text
  )

  if ($null -eq $CodeModule) {
    return
  }
  if ($CodeModule.CountOfLines -gt 0) {
    $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
  }
  if (-not [string]::IsNullOrWhiteSpace($Text)) {
    $CodeModule.AddFromString($Text)
  }
}

function Get-XlflowUserFormCodeTextFromSource {
  param(
    [string]$FormsDir,
    [string]$FormName
  )

  $path = Get-XlflowUserFormCodePath -FormsDir $FormsDir -FormName $FormName
  if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
    return $null
  }
  return (Get-XlflowUtf8Text -Path $path)
}

function Export-XlflowUserFormCodeBehind {
  param(
    $Component,
    [string]$FormsDir
  )

  if ($null -eq $Component -or [int]$Component.Type -ne 3) {
    return ""
  }

  $path = Get-XlflowUserFormCodePath -FormsDir $FormsDir -FormName ([string]$Component.Name)
  if ([string]::IsNullOrWhiteSpace($path)) {
    return ""
  }

  $text = Get-XlflowCodeModuleText -CodeModule $Component.CodeModule
  if ([string]::IsNullOrWhiteSpace($text)) {
    if (Test-Path -LiteralPath $path) {
      Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
    }
    return ""
  }

  $parent = Split-Path -Parent $path
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }
  Set-XlflowUtf8Text -Path $path -Text $text
  return $path
}

function Sync-XlflowUserFormCodeBehind {
  param(
    $Component,
    [string]$FormsDir,
    [string]$FallbackText = $null
  )

  if ($null -eq $Component -or [int]$Component.Type -ne 3) {
    return $false
  }

  $text = Get-XlflowUserFormCodeTextFromSource -FormsDir $FormsDir -FormName ([string]$Component.Name)
  if ($null -eq $text) {
    $text = $FallbackText
  }
  if ($null -eq $text) {
    return $false
  }

  Set-XlflowCodeModuleText -CodeModule $Component.CodeModule -Text $text
  return $true
}

function New-XlflowTestRunnerCode {
  param($Tests, $HooksByModule)

  $builder = New-Object System.Text.StringBuilder
  $dq = '"'
  [void]$builder.AppendLine('Option Explicit')
  [void]$builder.AppendLine('')

  # Generate dedicated BeforeAll/AfterAll wrappers per module so On Error Resume Next
  # works correctly without nesting Application.Run inside another COM invocation.
  $moduleNames = $Tests | ForEach-Object { $_.module } | Select-Object -Unique
  foreach ($mod in $moduleNames) {
    $hooks = $HooksByModule[$mod]
    if ($null -ne $hooks.BeforeAll) {
      [void]$builder.AppendLine("Public Function RunBeforeAll_$mod() As Variant")
      [void]$builder.AppendLine('  On Error GoTo Handler')
      [void]$builder.AppendLine("  $mod.$($hooks.BeforeAll.name)")
      [void]$builder.AppendLine("  RunBeforeAll_$mod = Array(True, CLng(0), $dq$dq, $dq$dq, $dq$dq)")
      [void]$builder.AppendLine('  Exit Function')
      [void]$builder.AppendLine('Handler:')
      [void]$builder.AppendLine("  RunBeforeAll_$mod = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description), $dq$dq)")
      [void]$builder.AppendLine('  Err.Clear')
      [void]$builder.AppendLine('End Function')
      [void]$builder.AppendLine('')
    }
    if ($null -ne $hooks.AfterAll) {
      [void]$builder.AppendLine("Public Function RunAfterAll_$mod() As Variant")
      [void]$builder.AppendLine('  On Error GoTo Handler')
      [void]$builder.AppendLine("  $mod.$($hooks.AfterAll.name)")
      [void]$builder.AppendLine("  RunAfterAll_$mod = Array(True, CLng(0), $dq$dq, $dq$dq, $dq$dq)")
      [void]$builder.AppendLine('  Exit Function')
      [void]$builder.AppendLine('Handler:')
      [void]$builder.AppendLine("  RunAfterAll_$mod = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description), $dq$dq)")
      [void]$builder.AppendLine('  Err.Clear')
      [void]$builder.AppendLine('End Function')
      [void]$builder.AppendLine('')
    }
  }

  [void]$builder.AppendLine('Public Function RunTest(ByVal testIndex As Long) As Variant')
  [void]$builder.AppendLine('  On Error Resume Next')
  [void]$builder.AppendLine('  Err.Clear')
  [void]$builder.AppendLine('  Dim beforeEachErr As Variant')
  [void]$builder.AppendLine('  Dim testErr As Variant')
  [void]$builder.AppendLine('  Dim afterEachErr As Variant')
  [void]$builder.AppendLine('  Dim statusHint As String')
  [void]$builder.AppendLine('  Dim phaseHint As String')
  [void]$builder.AppendLine('  Select Case testIndex')

  $index = 0
  foreach ($test in $Tests) {
    $hooks = $HooksByModule[$test.module]
    $beforeEachName = ""
    $afterEachName = ""
    if ($null -ne $hooks) {
      if ($null -ne $hooks.BeforeEach) { $beforeEachName = $hooks.BeforeEach.name }
      if ($null -ne $hooks.AfterEach) { $afterEachName = $hooks.AfterEach.name }
    }

    [void]$builder.AppendLine("    Case $index")
    [void]$builder.AppendLine('      statusHint = ""')
    [void]$builder.AppendLine('      phaseHint = ""')
    if ($beforeEachName -ne "") {
      [void]$builder.AppendLine("      $($test.module).$beforeEachName")
      [void]$builder.AppendLine('      If Err.Number <> 0 Then')
      [void]$builder.AppendLine('        beforeEachErr = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description))')
      [void]$builder.AppendLine('        phaseHint = "before_each"')
      [void]$builder.AppendLine('        Err.Clear')
      [void]$builder.AppendLine('      End If')
    }
    [void]$builder.AppendLine('      If IsEmpty(beforeEachErr) Then')
    [void]$builder.AppendLine("        $($test.module).$($test.name)")
    [void]$builder.AppendLine('        If Err.Number <> 0 Then')
    [void]$builder.AppendLine('          If Err.Number = vbObjectError + 516 Then')
    [void]$builder.AppendLine('            statusHint = "inconclusive"')
    [void]$builder.AppendLine('          End If')
    [void]$builder.AppendLine('          testErr = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description))')
    [void]$builder.AppendLine('          phaseHint = "test"')
    [void]$builder.AppendLine('          Err.Clear')
    [void]$builder.AppendLine('        End If')
    [void]$builder.AppendLine('      End If')
    if ($afterEachName -ne "") {
      [void]$builder.AppendLine("      $($test.module).$afterEachName")
      [void]$builder.AppendLine('      If Err.Number <> 0 Then')
      [void]$builder.AppendLine('        afterEachErr = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description))')
      [void]$builder.AppendLine('        If phaseHint = "" Then')
      [void]$builder.AppendLine('          phaseHint = "after_each"')
      [void]$builder.AppendLine('        End If')
      [void]$builder.AppendLine('        Err.Clear')
      [void]$builder.AppendLine('      End If')
    }
    [void]$builder.AppendLine('      If Not IsEmpty(afterEachErr) Then')
    [void]$builder.AppendLine('        phaseHint = "after_each"')
    [void]$builder.AppendLine('        statusHint = "failed"')
    [void]$builder.AppendLine('        RunTest = Array(False, afterEachErr(1), afterEachErr(2), afterEachErr(3), statusHint, phaseHint)')
    [void]$builder.AppendLine('      ElseIf Not IsEmpty(testErr) Then')
    [void]$builder.AppendLine('        RunTest = Array(False, testErr(1), testErr(2), testErr(3), statusHint, phaseHint)')
    [void]$builder.AppendLine('      ElseIf Not IsEmpty(beforeEachErr) Then')
    [void]$builder.AppendLine('        RunTest = Array(False, beforeEachErr(1), beforeEachErr(2), beforeEachErr(3), statusHint, phaseHint)')
    [void]$builder.AppendLine('      Else')
    [void]$builder.AppendLine('        RunTest = Array(True, CLng(0), "", "", statusHint, phaseHint)')
    [void]$builder.AppendLine('      End If')
    $index++
  }

  [void]$builder.AppendLine('  End Select')
  [void]$builder.AppendLine('  Err.Clear')
  [void]$builder.AppendLine('End Function')
  return $builder.ToString()
}

function Get-XlflowDocumentModuleContent {
  param([string]$Path)

  $lines = (Get-XlflowUtf8Text -Path $Path) -split "`r?`n"
  $filtered = New-Object System.Collections.Generic.List[string]
  $inClassHeader = $false
  $classHeaderBuffer = New-Object System.Collections.Generic.List[string]

  foreach ($line in $lines) {
    $trimmed = $line.Trim()
    if ($trimmed -eq "VERSION 1.0 CLASS") {
      $inClassHeader = $true
      $classHeaderBuffer.Clear()
      $classHeaderBuffer.Add($line)
      continue
    }
    if ($inClassHeader) {
      $classHeaderBuffer.Add($line)
      if ($trimmed -eq "END") {
        $inClassHeader = $false
        $classHeaderBuffer.Clear()
      }
      continue
    }
    if ($trimmed -match '^Attribute\s+VB_') {
      continue
    }
    $filtered.Add($line)
  }

  if ($inClassHeader -and $classHeaderBuffer.Count -gt 0) {
    foreach ($headerLine in $classHeaderBuffer) {
      $filtered.Add($headerLine)
    }
  }

  $hasOptionExplicit = $false
  $hasNonHeaderCode = $false
  foreach ($line in $filtered) {
    $trimmed = $line.Trim()
    if ($trimmed -eq "") {
      continue
    }
    if ($trimmed -ieq "Option Explicit") {
      $hasOptionExplicit = $true
      continue
    }
    $hasNonHeaderCode = $true
  }

  if (-not $hasOptionExplicit -and -not $hasNonHeaderCode) {
    $filtered.Add("")
    $filtered.Add("Option Explicit")
  }

  return ($filtered -join [Environment]::NewLine)
}

function Normalize-XlflowDocumentModuleFile {
  param(
    [string]$Path,
    [string]$RootDir = "",
    [string]$FolderAnnotationMode = "update"
  )

  $content = Get-XlflowDocumentModuleContent -Path $Path
  if (-not [string]::IsNullOrWhiteSpace($RootDir)) {
    $desiredAnnotation = Get-XlflowFolderAnnotationForPath -RootDir $RootDir -Path $Path
    $content = Update-XlflowFolderAnnotationText -Text $content -FolderAnnotationMode $FolderAnnotationMode -DesiredAnnotation $desiredAnnotation
  }
  Set-XlflowUtf8Text -Path $Path -Text $content
}

function Sync-XlflowDocumentModule {
  param(
    $Component,
    [string]$Path,
    [string]$RootDir = "",
    [string]$FolderAnnotationMode = "update"
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    return $false
  }

  $code = Get-XlflowDocumentModuleContent -Path $Path
  if (-not [string]::IsNullOrWhiteSpace($RootDir)) {
    $desiredAnnotation = Get-XlflowFolderAnnotationForPath -RootDir $RootDir -Path $Path
    $code = Update-XlflowFolderAnnotationText -Text $code -FolderAnnotationMode $FolderAnnotationMode -DesiredAnnotation $desiredAnnotation
  }
  $module = $Component.CodeModule
  $lineCount = $module.CountOfLines

  if ($lineCount -gt 0) {
    $module.DeleteLines(1, $lineCount)
  }

  if (-not [string]::IsNullOrWhiteSpace($code)) {
    $module.AddFromString($code)
  }

  return $true
}

function Find-XlflowDuplicateModuleNames {
  param($Files)

  $seen = @{}
  $originalNames = @{}
  foreach ($file in @($Files)) {
    if ($file.extension -eq ".frx") {
      continue
    }
    $key = ([string]$file.module_name).ToLowerInvariant()
    if (-not $seen.ContainsKey($key)) {
      $seen[$key] = New-Object System.Collections.Generic.List[string]
      $originalNames[$key] = [string]$file.module_name
    }
    $seen[$key].Add([string]$file.relative_path)
  }

  $duplicates = New-Object System.Collections.Generic.List[object]
  foreach ($key in $seen.Keys | Sort-Object) {
    if ($seen[$key].Count -lt 2) {
      continue
    }
    $duplicates.Add([pscustomobject][ordered]@{
      module_name = [string]$originalNames[$key]
      paths = @($seen[$key].ToArray())
    })
  }
  return @($duplicates.ToArray())
}

function Find-XlflowDocumentModulePath {
  param(
    [string]$WorkbookDir,
    [string]$ComponentName
  )

  if ([string]::IsNullOrWhiteSpace($WorkbookDir) -or -not (Test-Path -LiteralPath $WorkbookDir)) {
    return ""
  }

  foreach ($file in @(Get-XlflowSourceComponentFiles -ModulesDir "" -ClassesDir "" -FormsDir "" -WorkbookDir $WorkbookDir)) {
    if ($file.extension -eq ".bas" -and $file.module_name -ieq $ComponentName) {
      return [string]$file.full_name
    }
  }
  return ""
}

function ConvertFrom-XlflowRunArgumentsJson {
  param([string]$Json)

  if ([string]::IsNullOrWhiteSpace($Json)) {
    return @()
  }
  # Decode base64 JSON
  $decodedBytes = [System.Convert]::FromBase64String($Json)
  $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)

  $specs = ConvertFrom-Json -InputObject $decodedJson
  $values = New-Object System.Collections.Generic.List[object]
  foreach ($spec in $specs) {
    switch ([string]$spec.type) {
      "string" {
        $values.Add([string]$spec.value)
      }
      "int" {
        $parsed = 0
        if (-not [int]::TryParse([string]$spec.value, [ref]$parsed)) {
          throw "invalid int run argument: $($spec.value)"
        }
        $values.Add($parsed)
      }
      "bool" {
        if ($spec.value -ne "true" -and $spec.value -ne "false") {
          throw "invalid bool run argument: $($spec.value)"
        }
        $values.Add((ConvertTo-XlflowBool ([string]$spec.value)))
      }
      default {
        throw "unsupported run argument type: $($spec.type)"
      }
    }
  }
  return $values.ToArray()
}

function ConvertTo-XlflowVBALiteral {
  param([string]$Type, [string]$Value)

  switch ($Type) {
    "string" { return '"' + $Value.Replace('"', '""') + '"' }
    "int" { return "CLng(" + $Value + ")" }
    "double" { return "CDbl(" + $Value + ")" }
    "bool" {
      if ($Value -eq "true") {
        return "CBool(True)"
      }
      return "CBool(False)"
    }
    default { throw "unsupported run argument type: $Type" }
  }
}

function Get-XlflowMacroModuleName {
  param([string]$MacroName)

  $parts = $MacroName.Split(".")
  if ($parts.Count -lt 2) {
    return $MacroName
  }
  return ($parts[0..($parts.Count - 2)] -join ".")
}

function Assert-XlflowSaveAsExtension {
  param([string]$WorkbookPath, [string]$SaveAsPath)

  if ([string]::IsNullOrWhiteSpace($SaveAsPath)) {
    return
  }
  $workbookExtension = [System.IO.Path]::GetExtension($WorkbookPath)
  $saveAsExtension = [System.IO.Path]::GetExtension($SaveAsPath)
  if ($workbookExtension -ne $saveAsExtension) {
    throw "save-as extension $saveAsExtension does not match workbook extension $workbookExtension"
  }
}

function Format-XlflowMacroFailureMessage {
  param(
    [string]$ModuleName,
    [int]$Line,
    [int]$Number,
    [string]$Description
  )

  $parts = New-Object System.Collections.Generic.List[string]
  if (-not [string]::IsNullOrWhiteSpace($ModuleName)) {
    $parts.Add($ModuleName)
  }
  if ($Line -gt 0) {
    $parts.Add("line " + $Line)
  }
  if ($Number -ne 0) {
    $parts.Add("Err " + $Number)
  }
  if ([string]::IsNullOrWhiteSpace($Description)) {
    return ($parts -join " ")
  }
  if ($parts.Count -eq 0) {
    return $Description
  }
  return (($parts -join " ") + ": " + $Description)
}

function New-XlflowRunnerModuleCode {
  $builder = New-Object System.Text.StringBuilder
  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("' Persistent marker module for xlflow fast run workflows.")
  [void]$builder.AppendLine("Public Function XlflowRunnerVersion() As String")
  [void]$builder.AppendLine('  XlflowRunnerVersion = "1"')
  [void]$builder.AppendLine("End Function")
  return $builder.ToString()
}

function Test-XlflowRunnerModuleInstalled {
  param($VBProject)
  try {
    $null = $VBProject.VBComponents.Item("XlflowRunner")
    return $true
  } catch {
    return $false
  }
}

function Install-XlflowRunnerModule {
  param($VBProject)
  try {
    $existing = $VBProject.VBComponents.Item("XlflowRunner")
    $VBProject.VBComponents.Remove($existing)
  } catch {
    Write-Verbose ("XlflowRunner was not installed before install: " + $_.Exception.Message)
  }
  $component = $VBProject.VBComponents.Add(1)
  $component.Name = "XlflowRunner"
  $component.CodeModule.AddFromString((New-XlflowRunnerModuleCode))
}

function Remove-XlflowRunnerModule {
  param($VBProject)
  try {
    $existing = $VBProject.VBComponents.Item("XlflowRunner")
    $VBProject.VBComponents.Remove($existing)
    return $true
  } catch {
    return $false
  }
}

function Remove-XlflowVBComponentByName {
  param(
    $VBProject,
    [string]$Name
  )

  try {
    $existing = $VBProject.VBComponents.Item($Name)
    $VBProject.VBComponents.Remove($existing)
    return $true
  } catch {
    return $false
  }
}

function Get-XlflowVBComponentByName {
  param(
    $VBProject,
    [string]$Name
  )

  try {
    return $VBProject.VBComponents.Item($Name)
  } catch {
    return $null
  }
}

function Install-XlflowVBComponentFromCode {
  param(
    $VBProject,
    [string]$Name,
    [string]$Code
  )

  $existing = Get-XlflowVBComponentByName -VBProject $VBProject -Name $Name
  if ($null -ne $existing) {
    throw ("VBA component '" + $Name + "' already exists.")
  }
  $component = $VBProject.VBComponents.Add(1)
  $component.Name = $Name
  $component.CodeModule.AddFromString($Code)
  return $component
}

function New-XlflowInspectFormModuleCode {
  return @'
Option Explicit

Private Const xlflowBasisDesigner As String = "designer"
Private Const xlflowBasisRuntime As String = "runtime"
Private Const xlflowCoordinateSystem As String = "parent-relative"

Public Function XlflowInspectFormJson(ByVal formName As String, ByVal basis As String, Optional ByVal initializer As String = "") As String
  Dim normalizedBasis As String
  normalizedBasis = LCase$(Trim$(basis))

  Select Case normalizedBasis
    Case xlflowBasisDesigner
      XlflowInspectFormJson = InspectDesignerFormJson(formName)
    Case xlflowBasisRuntime
      XlflowInspectFormJson = InspectRuntimeFormJson(formName, initializer)
    Case Else
      Err.Raise vbObjectError + 7300, "XlflowInspectFormJson.args", "Unsupported inspect basis: " & basis
  End Select
End Function

Private Function InspectDesignerFormJson(ByVal formName As String) As String
  On Error GoTo ErrHandler

  Dim component As Object
  Dim designer As Object

  Set component = ThisWorkbook.VBProject.VBComponents.Item(formName)
  Set designer = component.Designer
  InspectDesignerFormJson = SerializeFormSnapshot(formName, xlflowBasisDesigner, designer)
  Exit Function

ErrHandler:
  Err.Raise Err.Number, "XlflowInspectFormJson.designer", Err.Description
End Function

Private Function InspectRuntimeFormJson(ByVal formName As String, ByVal initializer As String) As String
  Dim formInstance As Object
  Dim loaded As Boolean
  Dim initializerRan As Boolean
  Dim errorNumber As Long
  Dim errorDescription As String
  Dim errorSource As String

  On Error GoTo ErrHandler

  Set formInstance = UserForms.Add(formName)
  loaded = True

  If Len(Trim$(initializer)) > 0 Then
    CallByName formInstance, Trim$(initializer), VbMethod, ThisWorkbook
    initializerRan = True
  End If

  InspectRuntimeFormJson = SerializeFormSnapshot(formName, xlflowBasisRuntime, formInstance)

Cleanup:
  On Error Resume Next
  If loaded Then
    Unload formInstance
  End If
  Set formInstance = Nothing
  On Error GoTo 0

  If errorNumber <> 0 Then
    Err.Raise errorNumber, errorSource, errorDescription
  End If
  Exit Function

ErrHandler:
  errorNumber = Err.Number
  errorDescription = Err.Description
  If Not loaded Then
    errorSource = "XlflowInspectFormJson.runtime_load"
  ElseIf Len(Trim$(initializer)) > 0 And Not initializerRan Then
    errorSource = "XlflowInspectFormJson.initializer"
  Else
    errorSource = "XlflowInspectFormJson.enumerate"
  End If
  Resume Cleanup
End Function

Private Function SerializeFormSnapshot(ByVal formName As String, ByVal basis As String, ByVal formObject As Object) As String
  Dim json As String
  Dim hasFields As Boolean
  Dim controls As Object

  json = "{"

  JsonAddString json, "name", formName, hasFields
  JsonAddString json, "basis", basis, hasFields
  JsonAddStringFromMember json, formObject, "Caption", "caption", hasFields
  JsonAddNumberFromMember json, formObject, "Width", "width", hasFields
  JsonAddNumberFromMember json, formObject, "Height", "height", hasFields
  JsonAddString json, "coordinate_system", xlflowCoordinateSystem, hasFields

  Set controls = GetObjectControls(formObject)
  JsonAddRaw json, "controls", SerializeControls(controls, formName), hasFields

  json = json & "}"
  SerializeFormSnapshot = json
End Function

Private Function SerializeControls(ByVal controls As Object, ByVal expectedParentName As String) As String
  Dim json As String
  Dim first As Boolean
  Dim control As Object

  json = "["
  first = True

  If Not controls Is Nothing Then
    For Each control In controls
      If Not ControlHasExpectedParent(control, expectedParentName) Then
        GoTo ContinueLoop
      End If
      If Not first Then
        json = json & ","
      End If
      json = json & SerializeControl(control)
      first = False
ContinueLoop:
    Next control
  End If

  json = json & "]"
  SerializeControls = json
End Function

Private Function SerializeControl(ByVal control As Object) As String
  Dim json As String
  Dim hasFields As Boolean
  Dim children As Object
  Dim listCount As Long
  Dim selectedIndex As Long
  Dim typeNameValue As String

  json = "{"

  JsonAddStringFromMember json, control, "Name", "name", hasFields
  typeNameValue = TypeName(control)
  JsonAddString json, "type", typeNameValue, hasFields
  JsonAddStringFromMember json, control, "ProgId", "prog_id", hasFields
  JsonAddStringFromMember json, control, "Caption", "caption", hasFields
  JsonAddStringFromMember json, control, "Text", "text", hasFields
  JsonAddStringFromMember json, control, "Value", "value", hasFields
  JsonAddNumberFromMember json, control, "Left", "left", hasFields
  JsonAddNumberFromMember json, control, "Top", "top", hasFields
  JsonAddNumberFromMember json, control, "Width", "width", hasFields
  JsonAddNumberFromMember json, control, "Height", "height", hasFields
  JsonAddLongFromMember json, control, "TabIndex", "tab_index", hasFields
  JsonAddBoolFromMember json, control, "Enabled", "enabled", hasFields
  JsonAddBoolFromMember json, control, "Visible", "visible", hasFields

  If TryGetLongMember(control, "ListIndex", selectedIndex) Then
    JsonAddLong json, "selected_index", selectedIndex, hasFields
  End If
  If TryGetLongMember(control, "ListCount", listCount) Then
    JsonAddRaw json, "list", SerializeControlList(control, listCount), hasFields
  End If

  If ControlCanContainChildren(typeNameValue) Then
    Set children = GetObjectControls(control)
  End If
  If Not children Is Nothing Then
    JsonAddRaw json, "controls", SerializeControls(children, SafeControlName(control)), hasFields
  End If

  json = json & "}"
  SerializeControl = json
End Function

Private Function ControlHasExpectedParent(ByVal control As Object, ByVal expectedParentName As String) As Boolean
  On Error GoTo Missing

  Dim parentObject As Object
  Dim parentName As String
  Set parentObject = CallByName(control, "Parent", VbGet)
  If parentObject Is Nothing Then
    GoTo Missing
  End If

  parentName = CallByName(parentObject, "Name", VbGet)
  ControlHasExpectedParent = (StrComp(Trim$(parentName), Trim$(expectedParentName), vbTextCompare) = 0)
  Exit Function

Missing:
  ControlHasExpectedParent = False
End Function

Private Function SafeControlName(ByVal control As Object) As String
  On Error GoTo Missing

  SafeControlName = CStr(CallByName(control, "Name", VbGet))
  Exit Function

Missing:
  SafeControlName = vbNullString
End Function

Private Function ControlCanContainChildren(ByVal controlType As String) As Boolean
  Select Case LCase$(Trim$(controlType))
    Case "frame", "multipage", "page", "tabstrip"
      ControlCanContainChildren = True
    Case Else
      ControlCanContainChildren = False
  End Select
End Function

Private Function SerializeControlList(ByVal control As Object, ByVal listCount As Long) As String
  Dim json As String
  Dim i As Long
  Dim first As Boolean
  Dim itemValue As String

  json = "["
  first = True

  If listCount > 0 Then
    For i = 0 To listCount - 1
      If TryGetListItem(control, i, itemValue) Then
        If Not first Then
          json = json & ","
        End If
        json = json & JsonQuote(itemValue)
        first = False
      End If
    Next i
  End If

  json = json & "]"
  SerializeControlList = json
End Function

Private Function TryGetListItem(ByVal control As Object, ByVal index As Long, ByRef valueOut As String) As Boolean
  On Error GoTo Missing

  Dim itemValue As Variant
  itemValue = control.List(index)
  If IsNull(itemValue) Or IsEmpty(itemValue) Then
    valueOut = vbNullString
  Else
    valueOut = CStr(itemValue)
  End If
  TryGetListItem = True
  Exit Function

Missing:
  valueOut = vbNullString
  TryGetListItem = False
End Function

Private Function GetObjectControls(ByVal target As Object) As Object
  On Error Resume Next
  Set GetObjectControls = target.Controls
  On Error GoTo 0
End Function

Private Sub JsonAddStringFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As String
  If TryGetStringMember(target, memberName, value) Then
    JsonAddString json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddNumberFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Double
  If TryGetNumberMember(target, memberName, value) Then
    JsonAddNumber json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddLongFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Long
  If TryGetLongMember(target, memberName, value) Then
    JsonAddLong json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddBoolFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Boolean
  If TryGetBoolMember(target, memberName, value) Then
    JsonAddBool json, jsonKey, value, hasFields
  End If
End Sub

Private Function TryGetStringMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As String) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CStr(rawValue)
  TryGetStringMember = True
  Exit Function

Missing:
  valueOut = vbNullString
  TryGetStringMember = False
End Function

Private Function TryGetNumberMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Double) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CDbl(rawValue)
  TryGetNumberMember = True
  Exit Function

Missing:
  valueOut = 0
  TryGetNumberMember = False
End Function

Private Function TryGetLongMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Long) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CLng(rawValue)
  TryGetLongMember = True
  Exit Function

Missing:
  valueOut = 0
  TryGetLongMember = False
End Function

Private Function TryGetBoolMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Boolean) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CBool(rawValue)
  TryGetBoolMember = True
  Exit Function

Missing:
  valueOut = False
  TryGetBoolMember = False
End Function

Private Sub JsonAddRaw(ByRef json As String, ByVal key As String, ByVal rawValue As String, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & rawValue
  hasFields = True
End Sub

Private Sub JsonAddString(ByRef json As String, ByVal key As String, ByVal value As String, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & JsonQuote(value)
  hasFields = True
End Sub

Private Sub JsonAddNumber(ByRef json As String, ByVal key As String, ByVal value As Double, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & Trim$(Str$(value))
  hasFields = True
End Sub

Private Sub JsonAddLong(ByRef json As String, ByVal key As String, ByVal value As Long, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & CStr(value)
  hasFields = True
End Sub

Private Sub JsonAddBool(ByRef json As String, ByVal key As String, ByVal value As Boolean, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & IIf(value, "true", "false")
  hasFields = True
End Sub

Private Function JsonQuote(ByVal value As String) As String
  JsonQuote = """" & JsonEscape(value) & """"
End Function

Private Function JsonEscape(ByVal value As String) As String
  Dim text As String
  text = value
  text = Replace(text, "\", "\\")
  text = Replace(text, """", Chr$(92) & Chr$(34))
  text = Replace(text, vbCrLf, "\n")
  text = Replace(text, vbCr, "\n")
  text = Replace(text, vbLf, "\n")
  text = Replace(text, vbTab, "\t")
  JsonEscape = text
End Function
'@
}

function Test-XlflowEventProcedureName {
  param([string]$ProcedureName)

  if ($ProcedureName -match '^(?:Workbook|Worksheet)_') {
    return $true
  }
  if ($ProcedureName -match '^(?:Auto_Open|Auto_Close)$') {
    return $true
  }
  return $false
}

function Find-XlflowMacroProcedures {
  param([string]$ModuleName, [int]$ComponentType = 0, [string]$Code)

  $macros = New-Object System.Collections.Generic.List[object]
  if ([string]::IsNullOrEmpty($Code)) {
    return $macros
  }

  $componentTypeName = Get-XlflowComponentTypeName -ComponentType $ComponentType

  $lines = $Code -split "`r?`n"
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    if ($line -match '^(?i)(Private|Friend)\s+(Sub|Function)\b') {
      continue
    }
    $match = [regex]::Match($line, '^(?:(Public)\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(([^)]*)\))?', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
      continue
    }
    $name = $match.Groups[3].Value
    if ([string]::IsNullOrWhiteSpace($name)) {
      continue
    }
    $visibility = if ($match.Groups[1].Success) { "Public" } else { "Public" }
    $argText = $match.Groups[4].Value.Trim()
    $macroArgs = @()
    $hasParams = $false
    if (-not [string]::IsNullOrWhiteSpace($argText)) {
      $macroArgs = @($argText -split "," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
      $hasParams = ($macroArgs.Count -gt 0)
    }

    $reason = $null
    $runnable = $false

    if ($hasParams) {
      $reason = "has_parameters"
    } elseif (Test-XlflowEventProcedureName -ProcedureName $name) {
      $reason = "event_procedure"
    } elseif ($componentTypeName -eq "userform" -or $componentTypeName -eq "document_module" -or $componentTypeName -eq "unknown") {
      $reason = "unsupported_component_type"
    } else {
      $runnable = $true
    }

    $qualifiedName = $ModuleName + "." + $name

    $macros.Add([pscustomobject][ordered]@{
      module = $ModuleName
      name = $name
      qualified_name = $qualifiedName
      kind = $match.Groups[2].Value.ToLowerInvariant()
      args = @($macroArgs)
      line = $i + 1
      component_type = $componentTypeName
      visibility = $visibility
      has_parameters = $hasParams
      runnable = $runnable
      reason_not_runnable = $reason
    })
  }

  foreach ($macro in $macros) {
    Write-Output $macro
  }
}

function Test-XlflowMacroTargetFailure {
  param(
    [int]$Number,
    [string]$Description
  )

  if ($Description -match '(?i)(cannot run the macro|sub or function not defined|macro may not be available|unable to run)') {
    return $true
  }
  if ($Description -match 'マクロ.*(実行できません|使用できない|利用できない)' -or $Description -match 'Sub または Function が定義されていません') {
    return $true
  }
  if ($Number -eq 1004 -and $Description -match '(?i)macro') {
    return $true
  }
  return $false
}

function Test-XlflowMacroDisabledFailure {
  param(
    [int]$Number,
    [string]$Description
  )

  if ([string]::IsNullOrWhiteSpace($Description)) {
    return $false
  }
  if ($Description -match '(?i)(security settings.*macro|macros? (?:have been|were|are) disabled|disable all macros|because of your security settings|security warning)') {
    return $true
  }
  if ($Description -match 'セキュリティ.*マクロ.*無効' -or $Description -match 'マクロ.*無効.*セキュリティ') {
    return $true
  }
  if ($Number -eq 1004 -and $Description -match 'セキュリティ') {
    return $true
  }
  return $false
}

function ConvertTo-XlflowUIButtonId {
  param([string]$Value)

  $value = ([string]$Value).Trim().ToLowerInvariant()
  $builder = New-Object System.Text.StringBuilder
  $lastDash = $false
  foreach ($char in $value.ToCharArray()) {
    $isValid = (($char -ge [char]'a' -and $char -le [char]'z') -or ($char -ge [char]'0' -and $char -le [char]'9'))
    if ($isValid) {
      [void]$builder.Append($char)
      $lastDash = $false
      continue
    }
    if (-not $lastDash -and $builder.Length -gt 0) {
      [void]$builder.Append("-")
      $lastDash = $true
    }
  }
  return $builder.ToString().Trim("-")
}

function ConvertTo-XlflowUIButtonName {
  param([string]$Id)
  return "xlflow.button." + (ConvertTo-XlflowUIButtonId -Value $Id)
}

function Get-XlflowWorksheet {
  param($Workbook, [string]$Sheet)

  try {
    return $Workbook.Worksheets.Item($Sheet)
  } catch {
    return $null
  }
}

function Get-XlflowUIButton {
  param($Worksheet, [string]$Name)

  $buttons = $Worksheet.Buttons()
  for ($i = 1; $i -le $buttons.Count; $i++) {
    $button = $buttons.Item($i)
    if ($button.Name -eq $Name) {
      return $button
    }
  }
  return $null
}

function ConvertTo-XlflowUIButtonInfo {
  param($Button, [string]$Sheet, [string]$Id, [bool]$Updated = $false)

  $cell = ""
  try {
    $cell = $Button.TopLeftCell.Address($false, $false)
  } catch {
    Write-Verbose ("failed to read button top-left cell: " + $_.Exception.Message)
  }
  return [ordered]@{
    id = $Id
    name = $Button.Name
    sheet = $Sheet
    text = $Button.Caption
    macro = $Button.OnAction
    cell = $cell
    left = [double]$Button.Left
    top = [double]$Button.Top
    width = [double]$Button.Width
    height = [double]$Button.Height
    updated = $Updated
  }
}

function Test-XlflowMacroExists {
  param($Workbook, [string]$MacroName)

  $project = $Workbook.VBProject
  foreach ($component in @($project.VBComponents)) {
    if ($component.Name -like "Xlflow*") {
      continue
    }
    $code = Get-XlflowCodeModuleText -CodeModule $component.CodeModule
    $macros = Find-XlflowMacroProcedures -ModuleName $component.Name -Code $code
    foreach ($macro in @($macros)) {
      if ($macro.qualified_name -eq $MacroName -or $macro.name -eq $MacroName) {
        return $true
      }
    }
  }
  return $false
}

function New-XlflowRunHarnessModuleName {
  $suffix = [Guid]::NewGuid().ToString("N").Substring(0, 20)
  return "XlflowRun_" + $suffix
}

function New-XlflowRunHarnessCode {
  param(
    [string]$MacroName,
    [object[]]$Arguments
  )

  $builder = New-Object System.Text.StringBuilder
  $moduleName = Get-XlflowMacroModuleName -MacroName $MacroName
  $literals = New-Object System.Collections.Generic.List[string]
  foreach ($argument in $Arguments) {
    $literals.Add((ConvertTo-XlflowVBALiteral -Type ([string]$argument.type) -Value ([string]$argument.value)))
  }
  $macroLiteral = ConvertTo-XlflowVBALiteral -Type "string" -Value $MacroName
  $invocation = "Application.Run targetMacro"
  if ($literals.Count -gt 0) {
    $invocation += ", " + ($literals -join ", ")
  }

  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Function RunMacro() As Variant")
  [void]$builder.AppendLine("  Dim startedAt As Double")
  [void]$builder.AppendLine("  Dim targetMacro As String")
  [void]$builder.AppendLine("  startedAt = Timer")
  [void]$builder.AppendLine("  targetMacro = ""'"" & ThisWorkbook.Name & ""'!"" & " + $macroLiteral)
  [void]$builder.AppendLine("  On Error GoTo Handler")
  [void]$builder.AppendLine("  " + $invocation)
  [void]$builder.AppendLine('  RunMacro = Array(True, "' + $moduleName + '", CLng(0), "", CLng(0), CLng((Timer - startedAt) * 1000))')
  [void]$builder.AppendLine("  Exit Function")
  [void]$builder.AppendLine("Handler:")
  [void]$builder.AppendLine('  RunMacro = Array(False, "' + $moduleName + '", CLng(Err.Number), CStr(Err.Description), CLng(Erl), CLng((Timer - startedAt) * 1000))')
  [void]$builder.AppendLine("  Err.Clear")
  [void]$builder.AppendLine("End Function")
  return $builder.ToString()
}
