param(
  [string]$WorkbookPath,
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = "",
  [string]$Entry = "",
  [string]$RunnableOnly = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "macros"
$excel = $null
$workbook = $null
$sessionAttached = $false
$sessionMode = "none"
$saveState = [ordered]@{ dirty = $false; needs_save = $false }

try {
  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached

  try {
    $project = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel"
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
    $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
    Write-XlflowJson -Result $result
    exit
  }

  $macros = New-Object System.Collections.Generic.List[object]
  foreach ($component in @($project.VBComponents)) {
    if ($component.Name -like "Xlflow*") {
      continue
    }
    $code = Get-XlflowCodeModuleText -CodeModule $component.CodeModule
    $componentMacros = Find-XlflowMacroProcedures -ModuleName $component.Name -ComponentType $component.Type -Code $code
    foreach ($macro in @($componentMacros)) {
      if ($null -ne $macro) {
        $macros.Add($macro) | Out-Null
      }
    }
  }

  $runnableOnly = $RunnableOnly -eq "true"
  if ($runnableOnly) {
    $macros = [System.Collections.Generic.List[object]]($macros | Where-Object { $_.runnable })
  }

  $defaultEntry = ""
  if (-not [string]::IsNullOrWhiteSpace($Entry)) {
    $entryMatch = $macros | Where-Object { $_.qualified_name -eq $Entry -and $_.runnable }
    if ($entryMatch) {
      $defaultEntry = $Entry
    }
  }

  $suggestions = @()
  if ($defaultEntry -ne "") {
    $suggestions += @([ordered]@{ title = "Run the default entrypoint"; command = "xlflow run $defaultEntry --session --json" })
  } else {
    $firstRunnable = $macros | Where-Object { $_.runnable } | Select-Object -First 1
    if ($firstRunnable) {
      $suggestions += @([ordered]@{ title = "Run the first runnable macro"; command = $firstRunnable.run_command })
    }
  }

  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  $result.macros = $macros.ToArray()
  if ($defaultEntry -ne "") {
    $result.default_entry = $defaultEntry
  }
  $result.suggestions = $suggestions
  if ($saveState.needs_save) {
    Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes."
  }
  if ($macros.Count -eq 0) {
    Add-XlflowHint -Result $result -Code "macros_empty_before_push" -Message "If you edited source files, run `xlflow push --session` before `xlflow macros --session`."
    Add-XlflowHint -Result $result -Code "macros_read_from_workbook" -Message "`macros` discovers procedures from the workbook, not directly from source files."
  }
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "discovered $($macros.Count) macro entrypoint(s)") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  Set-XlflowError -Result $result -Code "macro_discovery_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
} finally {
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
