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

  if (ConvertTo-XlflowBool $UseSession) {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
    $sessionAttached = $true
  } else {
    $resolvedWorkbookPath = $null
    $sessionWorkbookPath = $null
    if (-not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
      $resolvedWorkbookPath = [System.IO.Path]::GetFullPath($WorkbookPath)
    }
    if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and (Test-Path -LiteralPath $MetadataPath)) {
      try {
        $sessionMetadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
        if ($null -ne $sessionMetadata -and -not [string]::IsNullOrWhiteSpace([string]$sessionMetadata.workbook_path)) {
          $sessionWorkbookPath = [System.IO.Path]::GetFullPath([string]$sessionMetadata.workbook_path)
        }
      } catch {
        Write-Verbose ("failed to read trace session metadata: " + $_.Exception.Message)
      }
    }
    if ($null -ne $resolvedWorkbookPath -and $null -ne $sessionWorkbookPath -and $resolvedWorkbookPath -ieq $sessionWorkbookPath) {
      $excel = Get-XlflowExcelFromSessionMetadata -MetadataPath $MetadataPath
      if ($null -ne $excel) {
        try {
          $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
          $sessionAttached = $true
        } catch {
          Write-Verbose ("failed to attach trace command to active session workbook: " + $_.Exception.Message)
          $workbook = $null
          $excel = $null
        }
      }
    }
    if (-not $sessionAttached) {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = ConvertTo-XlflowBool $Visible
      $workbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts $false -DisableAutomationMacros $true
    }
  }

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
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true; session = $sessionAttached }
    $result.trace = [ordered]@{ lifecycle = "enabled"; workbook_injected = $true; log_dir = $TraceDir }
    $result.logs = @("enabled XlflowTrace in " + $WorkbookPath)
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
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $removedWorkbook; session = $sessionAttached }
    $result.trace = [ordered]@{ lifecycle = "disabled"; workbook_removed = $removedWorkbook; source_removed = $sourceRemoved; log_dir = $TraceDir }
    $result.logs = @("disabled XlflowTrace in " + $WorkbookPath)
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
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; session = $sessionAttached }
    $result.source = [ordered]@{ path = $sourcePath; exists = $sourceExists; matches_bundled = $sourceMatches }
    $result.trace = [ordered]@{ status = "ok"; workbook_injected = $workbookInjected; source_exists = $sourceExists; source_matches_bundled = $sourceMatches; log_dir = $TraceDir }
    $result.logs = @("reported XlflowTrace status for " + $WorkbookPath)
  }
} catch {
  if ($null -eq $result.error) {
    Set-XlflowError -Result $result -Code "trace_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  if ($null -eq $result.workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; session = $sessionAttached }
  }
} finally {
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
