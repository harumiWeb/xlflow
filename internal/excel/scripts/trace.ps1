param(
  [string]$Action = "enable",
  [string]$WorkbookPath,
  [string]$ModulesDir = "",
  [string]$Visible = "false",
  [string]$Force = "false",
  [string]$TraceDir = "",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "trace"
$excel = $null
$workbook = $null
$forceRemove = ConvertTo-XlflowBool $Force
$sessionAttached = $false
$sessionMode = "none"

try {
  if ($Action -eq "inject") {
    $Action = "enable"
  }

  if ($Action -eq "clean") {
    if ([string]::IsNullOrWhiteSpace($TraceDir)) {
      $TraceDir = Join-Path (Join-Path (Get-Location) ".xlflow") "traces"
    }
    $removed = 0
    if (Test-Path -LiteralPath $TraceDir) {
      $files = @(Get-ChildItem -LiteralPath $TraceDir -File -ErrorAction SilentlyContinue)
      $removed = $files.Count
      Remove-Item -LiteralPath $TraceDir -Recurse -Force
    }
    $result.trace = [ordered]@{ cleaned = $true; path = $TraceDir; files_removed = $removed }
    $result.logs = @("cleaned trace logs from " + $TraceDir)
    Write-XlflowJson -Result $result
    exit 0
  }

  if ($Action -notin @("enable", "disable", "status")) {
    Set-XlflowError -Result $result -Code "trace_args_invalid" -Message ("unsupported trace action: " + $Action) -Source "xlflow"
    throw "unsupported trace action"
  }

  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode

  try {
    $vbProject = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message $_.Exception.Message -Source "vbide"
    throw
  }

  if ($Action -eq "enable") {
    [void](Remove-XlflowTraceModule -VBProject $vbProject)
    $component = $vbProject.VBComponents.Add(1)
    $component.Name = "XlflowTrace"
    $component.CodeModule.AddFromString((New-XlflowTraceModuleCode))
    $workbook.Save()

    if (-not [string]::IsNullOrWhiteSpace($ModulesDir)) {
      $sourcePath = Write-XlflowTraceModuleSource -ModulesDir $ModulesDir
      $result.source = [ordered]@{ path = $sourcePath; updated = $true }
    }
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $true -NeedsSave $false -Dirty $false
    $result.trace = [ordered]@{ lifecycle = "enabled"; workbook_injected = $true; log_dir = $TraceDir }
    $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "enabled XlflowTrace in " + $WorkbookPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($null -ne $result.source) {
      $result.logs += ("wrote " + $result.source.path)
    }
  } elseif ($Action -eq "disable") {
    $removedWorkbook = Remove-XlflowTraceModule -VBProject $vbProject
    $sourceRemoved = $false
    if (-not [string]::IsNullOrWhiteSpace($ModulesDir)) {
      $sourcePath = Join-Path $ModulesDir "XlflowTrace.bas"
      if (Test-Path -LiteralPath $sourcePath) {
        if ((Test-XlflowTraceModuleSourceMatches -ModulesDir $ModulesDir) -or $forceRemove) {
          Remove-Item -LiteralPath $sourcePath -Force
          $sourceRemoved = $true
          $result.source = [ordered]@{ path = $sourcePath; updated = $true; removed = $true }
        } else {
          Set-XlflowError -Result $result -Code "trace_source_modified" -Message "XlflowTrace.bas differs from the bundled helper. Use --force to remove it." -Source "xlflow"
          throw "trace source modified"
        }
      }
    }
    if ($removedWorkbook) {
      $workbook.Save()
    }
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $removedWorkbook -NeedsSave $false -Dirty $false
    $result.trace = [ordered]@{ lifecycle = "disabled"; workbook_removed = $removedWorkbook; source_removed = $sourceRemoved; log_dir = $TraceDir }
    $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "disabled XlflowTrace in " + $WorkbookPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  } else {
    $sourcePath = $null
    $sourceExists = $false
    $sourceMatches = $false
    if (-not [string]::IsNullOrWhiteSpace($ModulesDir)) {
      $sourcePath = Join-Path $ModulesDir "XlflowTrace.bas"
      $sourceExists = Test-Path -LiteralPath $sourcePath
      $sourceMatches = Test-XlflowTraceModuleSourceMatches -ModulesDir $ModulesDir
    }
    $workbookInjected = Test-XlflowTraceModuleInjected -VBProject $vbProject
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -NeedsSave $false -Dirty $false
    $result.source = [ordered]@{ path = $sourcePath; exists = $sourceExists; matches_bundled = $sourceMatches }
    $result.trace = [ordered]@{ status = "ok"; workbook_injected = $workbookInjected; source_exists = $sourceExists; source_matches_bundled = $sourceMatches; log_dir = $TraceDir }
    $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "reported XlflowTrace status for " + $WorkbookPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  }
} catch {
  if ($null -eq $result.error) {
    Set-XlflowError -Result $result -Code "trace_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  if ($null -eq $result.workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -NeedsSave $false -Dirty $false
  }
} finally {
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
