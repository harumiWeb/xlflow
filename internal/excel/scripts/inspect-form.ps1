param(
  [string]$WorkbookPath,
  [string]$FormName,
  [string]$Basis = "runtime",
  [string]$Initializer = "",
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "inspect"
$excel = $null
$workbook = $null
$vbProject = $null
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

function Add-XlflowInspectFormSaveRequiredWarning {
  param(
    $Result,
    $SaveState
  )

  if ($null -ne $SaveState -and [bool]$SaveState.needs_save) {
    Add-XlflowStateWarning -Result $Result -Code "save_required" -Message "The live session workbook differs from disk. Run `xlflow save --session` to persist workbook changes."
  }
}

function Set-InspectFormValidationError {
  param(
    $Result,
    [string]$Code,
    [string]$Message
  )

  Set-XlflowError -Result $Result -Code $Code -Message $Message -Source "xlflow"
}

function Get-XlflowInspectFormErrorCode {
  param(
    [System.Exception]$Exception
  )

  if ($null -eq $Exception) {
    return "inspect_form_failed"
  }
  $source = [string]$Exception.Source
  $message = [string]$Exception.Message
  if ($source -like "*XlflowInspectFormJson.initializer*" -or $message -like "*XlflowInspectFormJson.initializer*") {
    return "form_initializer_failed"
  }
  if ($source -like "*XlflowInspectFormJson.runtime_load*" -or $message -like "*XlflowInspectFormJson.runtime_load*") {
    return "runtime_form_load_failed"
  }
  if ($source -like "*XlflowInspectFormJson.enumerate*" -or $message -like "*XlflowInspectFormJson.enumerate*") {
    return "control_enumeration_failed"
  }
  if ($source -like "*XlflowInspectFormJson.designer*" -or $message -like "*XlflowInspectFormJson.designer*") {
    return "designer_access_failed"
  }
  return "inspect_form_failed"
}

function Invoke-XlflowInspectFormSnapshot {
  param(
    $Excel,
    $Workbook,
    [string]$TargetFormName,
    [string]$TargetBasis,
    [string]$InitializerName = ""
  )

  $workbookName = ([string]$Workbook.Name).Replace("'", "''")
  $macroName = "'" + $workbookName + "'!XlflowInspectFormJson"
  $json = $Excel.Run($macroName, $TargetFormName, $TargetBasis, $InitializerName)
  if ([string]::IsNullOrWhiteSpace([string]$json)) {
    throw "inspect form helper returned no JSON"
  }
  return ($json | ConvertFrom-Json)
}

function New-XlflowInspectFormModuleName {
  $suffix = [Guid]::NewGuid().ToString("N").Substring(0, 20)
  return "XlflowForm_" + $suffix
}

function Update-XlflowInspectFormResultSaveState {
  param(
    $Result,
    $Workbook,
    [string]$WorkbookPath,
    [bool]$SessionAttached,
    [string]$SessionMode
  )

  if ($null -eq $Workbook -or [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    return [ordered]@{ dirty = $false; needs_save = $false }
  }

  $currentSaveState = Get-XlflowWorkbookSaveState -Workbook $Workbook -SessionAttached $SessionAttached
  $Result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $SessionAttached -SessionMode $SessionMode -Saved $false -Dirty $currentSaveState.dirty -NeedsSave $currentSaveState.needs_save
  $Result.target = New-XlflowTargetResult -Kind $(if ($SessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $Result.session = New-XlflowSessionResult -Active $SessionAttached -WorkbookPath $WorkbookPath -Dirty $currentSaveState.dirty -SaveRequired $currentSaveState.needs_save -Mode $SessionMode
  return $currentSaveState
}

function New-XlflowInspectRuntimeWorkbookCopy {
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
    Write-Verbose ("failed to resolve workbook extension for inspect-form temp copy: " + $_.Exception.Message)
  }

  $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("xlflow-inspect-form-" + [Guid]::NewGuid().ToString("N") + $extension)
  try {
    $SourceWorkbook.SaveCopyAs($tempPath)

    $tempExcel = New-Object -ComObject Excel.Application
    $tempExcel.Visible = $false
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

function Get-XlflowSafeMemberValue {
  param(
    $Target,
    [string]$Name,
    $Default = $null
  )

  try {
    return $Target.$Name
  } catch {
    return $Default
  }
}

function Get-XlflowSafeControls {
  param($Target)

  try {
    return @($Target.Controls)
  } catch {
    return @()
  }
}

function Test-XlflowControlCanContainChildren {
  param([string]$ControlType)

  return $ControlType -in @("Frame", "MultiPage", "Page", "TabStrip")
}

function Get-XlflowSafeControlList {
  param($Control)

  $listCount = Get-XlflowSafeMemberValue -Target $Control -Name "ListCount" -Default $null
  if ($null -eq $listCount) {
    return @(), $null
  }
  $items = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt [int]$listCount; $i++) {
    try {
      $value = $Control.List($i)
      if ($null -eq $value) {
        $items.Add("")
      } else {
        $items.Add([string]$value)
      }
    } catch {
      break
    }
  }
  return @($items.ToArray()), [int](Get-XlflowSafeMemberValue -Target $Control -Name "ListIndex" -Default -1)
}

function Get-XlflowDesignerControlSnapshot {
  param($Control)

  $progId = [string](Get-XlflowSafeMemberValue -Target $Control -Name "ProgId" -Default "")
  $controlType = "Control"
  if (-not [string]::IsNullOrWhiteSpace($progId)) {
    $segments = $progId.Split(".")
    if ($segments.Length -ge 2 -and -not [string]::IsNullOrWhiteSpace($segments[1])) {
      $controlType = $segments[1]
    }
  }

  $snapshot = [ordered]@{
    name = [string](Get-XlflowSafeMemberValue -Target $Control -Name "Name" -Default "")
    type = $controlType
  }

  foreach ($field in @(
    @{ name = "Caption"; key = "caption" },
    @{ name = "Text"; key = "text" },
    @{ name = "Value"; key = "value" },
    @{ name = "Left"; key = "left" },
    @{ name = "Top"; key = "top" },
    @{ name = "Width"; key = "width" },
    @{ name = "Height"; key = "height" },
    @{ name = "TabIndex"; key = "tab_index" },
    @{ name = "Enabled"; key = "enabled" },
    @{ name = "Visible"; key = "visible" }
  )) {
    $value = Get-XlflowSafeMemberValue -Target $Control -Name $field.name -Default $null
    if ($null -ne $value) {
      $snapshot[$field.key] = $value
    }
  }
  if (-not [string]::IsNullOrWhiteSpace($progId)) {
    $snapshot.prog_id = $progId
  }

  $listData = Get-XlflowSafeControlList -Control $Control
  if ($listData.Count -ge 1 -and $null -ne $listData[0]) {
    $snapshot.list = @($listData[0])
  }
  if ($listData.Count -ge 2 -and $null -ne $listData[1] -and [int]$listData[1] -ge -1) {
    $snapshot.selected_index = [int]$listData[1]
  }

  $children = @()
  if (Test-XlflowControlCanContainChildren -ControlType $controlType) {
    $children = @(Get-XlflowSafeControls -Target $Control)
  }
  if ($children.Count -gt 0) {
    $snapshot.controls = @($children | ForEach-Object { Get-XlflowDesignerControlSnapshot -Control $_ })
  }

  return [pscustomobject]$snapshot
}

function Get-XlflowDesignerFormSnapshot {
  param(
    $VBProject,
    [string]$TargetFormName
  )

  try {
    $component = $VBProject.VBComponents.Item($TargetFormName)
    $designer = $component.Designer
  } catch {
    throw
  }

  $controls = @(Get-XlflowSafeControls -Target $designer)
  return [pscustomobject][ordered]@{
    name = $TargetFormName
    basis = "designer"
    caption = [string](Get-XlflowSafeMemberValue -Target $designer -Name "Caption" -Default "")
    width = Get-XlflowSafeMemberValue -Target $designer -Name "Width" -Default $null
    height = Get-XlflowSafeMemberValue -Target $designer -Name "Height" -Default $null
    coordinate_system = "parent-relative"
    controls = @($controls | ForEach-Object { Get-XlflowDesignerControlSnapshot -Control $_ })
  }
}

try {
  if ([string]::IsNullOrWhiteSpace($Basis)) {
    $normalizedBasis = "runtime"
  } else {
    $normalizedBasis = $Basis.Trim().ToLowerInvariant()
  }
  if ([string]::IsNullOrWhiteSpace($FormName)) {
    Set-InspectFormValidationError -Result $result -Code "inspect_form_args_invalid" -Message "form name is required"
    Write-XlflowJson -Result $result
    exit
  }
  if ($normalizedBasis -notin @("runtime", "designer", "both")) {
    Set-InspectFormValidationError -Result $result -Code "inspect_form_args_invalid" -Message ("unsupported inspect form basis: " + $Basis)
    Write-XlflowJson -Result $result
    exit
  }
  if (-not [string]::IsNullOrWhiteSpace($Initializer) -and $normalizedBasis -eq "designer") {
    Set-InspectFormValidationError -Result $result -Code "inspect_form_args_invalid" -Message "--initializer can only be used with runtime or both inspection"
    Write-XlflowJson -Result $result
    exit
  }

  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached

  try {
    $vbProject = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbproject_access_denied" -Message "VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center." -Source "Excel"
    throw
  }

  $userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)
  if ($FormName -notin $userFormNames) {
    Set-XlflowError -Result $result -Code "form_not_found" -Message ("UserForm '" + $FormName + "' was not found in the workbook.") -Source "xlflow"
    throw "form not found"
  }

  if ($normalizedBasis -eq "both") {
    $runtimeOpenResult = New-XlflowInspectRuntimeWorkbookCopy -SourceWorkbook $workbook
    $runtimeExcel = $runtimeOpenResult.excel
    $runtimeWorkbook = $runtimeOpenResult.workbook
    $runtimeWorkbookPath = $runtimeOpenResult.path
    $runtimeVBProject = $runtimeWorkbook.VBProject
    $tempModuleName = New-XlflowInspectFormModuleName
    $null = Install-XlflowVBComponentFromCode -VBProject $runtimeVBProject -Name $tempModuleName -Code (New-XlflowInspectFormModuleCode)
    $tempModuleInstalled = $true
    $designerSnapshot = Invoke-XlflowInspectFormSnapshot -Excel $runtimeExcel -Workbook $runtimeWorkbook -TargetFormName $FormName -TargetBasis "designer"
    $runtimeSnapshot = Invoke-XlflowInspectFormSnapshot -Excel $runtimeExcel -Workbook $runtimeWorkbook -TargetFormName $FormName -TargetBasis "runtime" -InitializerName $Initializer
    $result.forms = [ordered]@{
      runtime = $runtimeSnapshot
      designer = $designerSnapshot
    }
  } elseif ($normalizedBasis -eq "runtime") {
    $runtimeOpenResult = New-XlflowInspectRuntimeWorkbookCopy -SourceWorkbook $workbook
    $runtimeExcel = $runtimeOpenResult.excel
    $runtimeWorkbook = $runtimeOpenResult.workbook
    $runtimeWorkbookPath = $runtimeOpenResult.path
    $runtimeVBProject = $runtimeWorkbook.VBProject
    $tempModuleName = New-XlflowInspectFormModuleName
    $null = Install-XlflowVBComponentFromCode -VBProject $runtimeVBProject -Name $tempModuleName -Code (New-XlflowInspectFormModuleCode)
    $tempModuleInstalled = $true
    $result.forms = Invoke-XlflowInspectFormSnapshot -Excel $runtimeExcel -Workbook $runtimeWorkbook -TargetFormName $FormName -TargetBasis "runtime" -InitializerName $Initializer
  } else {
    $runtimeOpenResult = New-XlflowInspectRuntimeWorkbookCopy -SourceWorkbook $workbook
    $runtimeExcel = $runtimeOpenResult.excel
    $runtimeWorkbook = $runtimeOpenResult.workbook
    $runtimeWorkbookPath = $runtimeOpenResult.path
    $runtimeVBProject = $runtimeWorkbook.VBProject
    $tempModuleName = New-XlflowInspectFormModuleName
    $null = Install-XlflowVBComponentFromCode -VBProject $runtimeVBProject -Name $tempModuleName -Code (New-XlflowInspectFormModuleCode)
    $tempModuleInstalled = $true
    $result.forms = Invoke-XlflowInspectFormSnapshot -Excel $runtimeExcel -Workbook $runtimeWorkbook -TargetFormName $FormName -TargetBasis "designer"
  }

  $saveState = Update-XlflowInspectFormResultSaveState -Result $result -Workbook $workbook -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
  if ($normalizedBasis -in @("runtime", "both")) {
    Add-XlflowWarning -Result $result -Code "runtime_form_loads_initialize" -Message "Runtime inspection loads the form and executes UserForm_Initialize."
    Add-XlflowWarning -Result $result -Code "runtime_form_temp_copy" -Message "Runtime inspection executed against a temporary workbook copy so the source workbook and live session are not mutated."
  }
  if (-not [string]::IsNullOrWhiteSpace($Initializer) -and $normalizedBasis -in @("runtime", "both")) {
    Add-XlflowWarning -Result $result -Code "runtime_form_initializer_invoked" -Message ("Runtime inspection also invoked " + $Initializer + "(ThisWorkbook).")
  }
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "inspected " + $normalizedBasis + " UserForm " + $FormName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  if ($null -eq $result.error) {
    $code = Get-XlflowInspectFormErrorCode -Exception $_.Exception
    Set-XlflowError -Result $result -Code $code -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  if ($null -eq $result.workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  }
  if ($null -eq $result.target -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  }
  if ($null -eq $result.session -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  }
} finally {
  if ($tempModuleInstalled -and $null -ne $runtimeVBProject) {
    $tempModuleRemoved = Remove-XlflowVBComponentByName -VBProject $runtimeVBProject -Name $tempModuleName
    if (-not $tempModuleRemoved) {
      Add-XlflowWarning -Result $result -Code "temporary_component_cleanup_failed" -Message ("Temporary helper module '" + $tempModuleName + "' could not be removed automatically.")
    }
  }
  if ($null -ne $workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $saveState = Update-XlflowInspectFormResultSaveState -Result $result -Workbook $workbook -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
    Add-XlflowInspectFormSaveRequiredWarning -Result $result -SaveState $saveState
  }
  if ($null -ne $runtimeWorkbook -or $null -ne $runtimeExcel) {
    Close-XlflowCom -Workbook $runtimeWorkbook -Excel $runtimeExcel -Save $false
  }
  if (-not [string]::IsNullOrWhiteSpace($runtimeWorkbookPath) -and (Test-Path -LiteralPath $runtimeWorkbookPath)) {
    Remove-Item -LiteralPath $runtimeWorkbookPath -Force -ErrorAction SilentlyContinue
  }
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
