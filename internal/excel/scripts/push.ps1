param(
  [string]$WorkbookPath,
  [string]$ModulesDir,
  [string]$ClassesDir,
  [string]$FormsDir,
  [string]$WorkbookDir,
  [string]$BackupRoot,
  [string]$StatePath = "",
  [string]$Visible = "false",
  [string]$BackupMode = "always",
  [string]$ChangedOnly = "false",
  [string]$UseSession = "false",
  [string]$NoSave = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "push"
$excel = $null
$workbook = $null
$tmpImportDir = $null
$sessionAttached = $false
$sessionMode = "none"
$sessionInvalidated = $false

try {
  if ($BackupMode -ne "always" -and $BackupMode -ne "never") {
    Set-XlflowError -Result $result -Code "push_args_invalid" -Message "-BackupMode must be always or never." -Source "xlflow"
    throw "invalid backup mode"
  }
  $fingerprint = Get-XlflowSourceFingerprint -WorkbookPath $WorkbookPath -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir
  if ((ConvertTo-XlflowBool $ChangedOnly) -and (Test-XlflowFingerprintMatchesState -Fingerprint $fingerprint -StatePath $StatePath)) {
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $false -SessionMode "none" -Saved $false -NeedsSave $false -Dirty $false
    $result.backup = [ordered]@{ path = $null; mode = $BackupMode }
    $result.source = [ordered]@{ changed_only = $true; changed = $false; state = $StatePath }
    $result.logs = @("source state unchanged; skipped workbook import")
    Write-XlflowJson -Result $result
    exit 0
  }

  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $backupRootFull = [System.IO.Path]::GetFullPath($BackupRoot)
  $xlflowRoot = Split-Path -Parent $backupRootFull
  if ([string]::IsNullOrWhiteSpace($xlflowRoot)) {
    throw "backup root parent could not be resolved: $BackupRoot"
  }
  $backupDir = Join-Path $backupRootFull $timestamp
  $tmpRoot = Join-Path $xlflowRoot "tmp"
  $tmpImportRoot = Join-Path $tmpRoot "import"
  $tmpImportDir = Join-Path $tmpImportRoot $timestamp
  if ($BackupMode -eq "always") {
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
  }
  New-Item -ItemType Directory -Force -Path $tmpImportDir | Out-Null

  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  if ($sessionAttached) {
    try { $excel = $workbook.Application } catch { Write-Verbose ("failed to resolve workbook application: " + $_.Exception.Message) }
  }

  foreach ($component in @($workbook.VBProject.VBComponents)) {
    if ($BackupMode -eq "always") {
      $backupPath = Get-XlflowComponentPath -Component $component -ModulesDir $backupDir -ClassesDir $backupDir -FormsDir $backupDir -WorkbookDir $backupDir
      if ($null -ne $backupPath) {
        $component.Export($backupPath)
      }
    }
    if ($component.Type -ne 100) {
      $workbook.VBProject.VBComponents.Remove($component)
    }
  }

  $imported = @()
  $updatedDocumentModules = @()
  foreach ($dir in @($ModulesDir, $ClassesDir, $FormsDir)) {
    if (Test-Path -LiteralPath $dir) {
      foreach ($file in Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Extension -in @(".bas", ".cls", ".frm") }) {
        $importPath = Join-Path $tmpImportDir $file.Name
        Copy-XlflowSourceForImport -SourcePath $file.FullName -DestinationPath $importPath
        if ($file.Extension -ieq ".frm") {
          $frxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".frx")
          if (Test-Path -LiteralPath $frxPath) {
            $importFrxPath = [System.IO.Path]::ChangeExtension($importPath, ".frx")
            Copy-XlflowSourceForImport -SourcePath $frxPath -DestinationPath $importFrxPath
          }
        }
        $workbook.VBProject.VBComponents.Import($importPath) | Out-Null
        $imported += $file.FullName
      }
    }
  }
  if (Test-Path -LiteralPath $WorkbookDir) {
    foreach ($component in @($workbook.VBProject.VBComponents | Where-Object { $_.Type -eq 100 })) {
      $sourcePath = Join-Path $WorkbookDir ($component.Name + ".bas")
      if (Sync-XlflowDocumentModule -Component $component -Path $sourcePath) {
        $updatedDocumentModules += $sourcePath
      }
    }
  }

  $saved = $false
  if (-not (ConvertTo-XlflowBool $NoSave)) {
    $workbook.Save()
    $saved = $true
  }
  Write-XlflowFingerprintState -Fingerprint $fingerprint -StatePath $StatePath
  $needsSave = $sessionAttached -and -not $saved
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $saved -NeedsSave $needsSave -Dirty $needsSave
  $result.backup = [ordered]@{ path = $(if ($BackupMode -eq "always") { $backupDir } else { $null }); mode = $BackupMode }
  $result.source = [ordered]@{ changed_only = (ConvertTo-XlflowBool $ChangedOnly); changed = $true; state = $StatePath }
  $result.logs = @(
    @(
      $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
      $(if ($BackupMode -eq "always") { "backed up existing VBA components" } else { "skipped VBA backup" }),
      "imported $($imported.Count) source file(s)",
      "updated $($updatedDocumentModules.Count) workbook module(s)",
      $(if ($saved) {
        "saved workbook in place"
      } elseif ($sessionAttached) {
        "SAVE REQUIRED: live session workbook differs from disk; run xlflow save before session stop"
      } else {
        "left workbook unchanged on disk"
      })
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  )
} catch {
  Set-XlflowError -Result $result -Code "excel_import_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  if ($sessionAttached) {
    try {
      if ($null -ne $workbook) {
        $workbook.Close($false) | Out-Null
      }
      if ($null -ne $excel) {
        $excel.Quit() | Out-Null
      }
      if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and (Test-Path -LiteralPath $MetadataPath)) {
        Remove-Item -LiteralPath $MetadataPath -Force -ErrorAction SilentlyContinue
      }
      $sessionInvalidated = $true
      $result.logs = @("invalidated xlflow session after failed workbook import")
    } catch {
      $result.logs = @("failed to invalidate xlflow session after import failure: " + $_.Exception.Message)
    }
  }
} finally {
  if ($sessionAttached -and -not $sessionInvalidated) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
  if (-not [string]::IsNullOrWhiteSpace($tmpImportDir) -and (Test-Path -LiteralPath $tmpImportDir)) {
    Remove-Item -LiteralPath $tmpImportDir -Recurse -Force -ErrorAction SilentlyContinue
  }
}

Write-XlflowJson -Result $result
