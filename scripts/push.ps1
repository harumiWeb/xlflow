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
$sessionInvalidated = $false

try {
  if ($BackupMode -ne "always" -and $BackupMode -ne "never") {
    Set-XlflowError -Result $result -Code "push_args_invalid" -Message "-BackupMode must be always or never." -Source "xlflow"
    throw "invalid backup mode"
  }
  $fingerprint = Get-XlflowSourceFingerprint -WorkbookPath $WorkbookPath -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir
  if ((ConvertTo-XlflowBool $ChangedOnly) -and (Test-XlflowFingerprintMatchesState -Fingerprint $fingerprint -StatePath $StatePath)) {
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; session = (ConvertTo-XlflowBool $UseSession) }
    $result.backup = [ordered]@{ path = $null; mode = $BackupMode }
    $result.source = [ordered]@{ changed_only = $true; changed = $false; state = $StatePath }
    $result.logs = @("source state unchanged; skipped workbook import")
    Write-XlflowJson -Result $result
    exit 0
  }

  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $backupDir = Join-Path $BackupRoot $timestamp
  $tmpImportDir = Join-Path (Join-Path (Split-Path -Parent $BackupRoot) "tmp") (Join-Path "import" $timestamp)
  if ($BackupMode -eq "always") {
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
  }
  New-Item -ItemType Directory -Force -Path $tmpImportDir | Out-Null

  if (ConvertTo-XlflowBool $UseSession) {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
    try { $excel = $workbook.Application } catch { Write-Verbose ("failed to resolve workbook application: " + $_.Exception.Message) }
    $sessionAttached = $true
  } else {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = ConvertTo-XlflowBool $Visible
    $workbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts $false -DisableAutomationMacros $true
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
  $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $saved; session = (ConvertTo-XlflowBool $UseSession) }
  $result.backup = [ordered]@{ path = $(if ($BackupMode -eq "always") { $backupDir } else { $null }); mode = $BackupMode }
  $result.source = [ordered]@{ changed_only = (ConvertTo-XlflowBool $ChangedOnly); changed = $true; state = $StatePath }
  $result.logs = @(
    $(if ($BackupMode -eq "always") { "backed up existing VBA components" } else { "skipped VBA backup" }),
    "imported $($imported.Count) source file(s)",
    "updated $($updatedDocumentModules.Count) workbook module(s)",
    $(if ($saved) { "saved workbook in place" } else { "left session workbook unsaved" })
  )
} catch {
  Set-XlflowError -Result $result -Code "excel_import_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  if ((ConvertTo-XlflowBool $UseSession) -and $sessionAttached) {
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
  if ((ConvertTo-XlflowBool $UseSession) -and -not $sessionInvalidated) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
  if (-not [string]::IsNullOrWhiteSpace($tmpImportDir) -and (Test-Path -LiteralPath $tmpImportDir)) {
    Remove-Item -LiteralPath $tmpImportDir -Recurse -Force -ErrorAction SilentlyContinue
  }
}

Write-XlflowJson -Result $result
