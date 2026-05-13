param(
  [string]$Action,
  [string]$WorkbookPath,
  [string]$FormsDir,
  [string]$ModulesDir,
  [string]$ClassesDir,
  [string]$WorkbookDir,
  [string]$ProjectRoot,
  [string]$Folders = "true",
  [string]$FolderAnnotation = "update",
  [string]$DefaultComponentFolders = "true",
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "list"
$excel = $null
$workbook = $null
$sessionAttached = $false
$sessionMode = "none"
$saveState = [ordered]@{ dirty = $false; needs_save = $false }

function Add-XlflowListSaveRequiredWarning {
  param(
    $Result,
    $SaveState
  )

  if ($null -ne $SaveState -and [bool]$SaveState.needs_save) {
    Add-XlflowStateWarning -Result $Result -Code "save_required" -Message "The live session workbook differs from disk. Run `xlflow save --session` to persist workbook changes."
  }
}

function ConvertTo-XlflowPortableRelativePath {
  param(
    [string]$BasePath,
    [string]$TargetPath
  )

  $relativePath = Get-XlflowRelativePath -BasePath $BasePath -TargetPath $TargetPath
  if ([string]::IsNullOrWhiteSpace($relativePath)) {
    return ""
  }
  return $relativePath.Replace("\", "/")
}

try {
  if ($Action -ne "forms") {
    Set-XlflowError -Result $result -Code "list_args_invalid" -Message "-Action must be forms." -Source "xlflow"
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
    $null = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbproject_access_denied" -Message "VBProject access is denied. Enable 'Trust access to the VBA project object model' in Excel Trust Center." -Source "Excel"
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
    $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
    Add-XlflowListSaveRequiredWarning -Result $result -SaveState $saveState
    Write-XlflowJson -Result $result
    exit
  }

  $forms = New-Object System.Collections.Generic.List[object]
  foreach ($component in @(Get-XlflowUserFormComponents -Workbook $workbook | Sort-Object Name)) {
    $sourcePath = Get-XlflowComponentPath -Component $component -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir -Folders $Folders -FolderAnnotation $FolderAnnotation -DefaultComponentFolders $DefaultComponentFolders
    $frxAbsolutePath = ""
    $sourceRelativePath = ""
    $frxRelativePath = ""
    $hasFrx = $false

    if (-not [string]::IsNullOrWhiteSpace($sourcePath)) {
      $sourceRelativePath = ConvertTo-XlflowPortableRelativePath -BasePath $ProjectRoot -TargetPath $sourcePath
      $frxAbsolutePath = [System.IO.Path]::ChangeExtension($sourcePath, ".frx")
      $hasFrx = Test-Path -LiteralPath $frxAbsolutePath
      if ($hasFrx) {
        $frxRelativePath = ConvertTo-XlflowPortableRelativePath -BasePath $ProjectRoot -TargetPath $frxAbsolutePath
      }
    }

    $form = [ordered]@{
      name = [string]$component.Name
      component_type = "MSForm"
      has_frx = $hasFrx
      source_path = $sourceRelativePath
    }
    if (-not [string]::IsNullOrWhiteSpace($frxRelativePath)) {
      $form.frx_path = $frxRelativePath
    }
    $forms.Add($form) | Out-Null
  }

  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  $result.forms = $forms.ToArray()
  Add-XlflowListSaveRequiredWarning -Result $result -SaveState $saveState
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "listed $($forms.Count) UserForm component(s)") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  Set-XlflowError -Result $result -Code "form_list_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  Add-XlflowListSaveRequiredWarning -Result $result -SaveState $saveState
} finally {
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
