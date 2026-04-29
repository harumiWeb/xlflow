param(
  [string]$WorkbookPath,
  [string]$ModulesDir,
  [string]$ClassesDir,
  [string]$FormsDir,
  [string]$WorkbookDir,
  [string]$BackupRoot,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "push"
$excel = $null
$workbook = $null

try {
  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $backupDir = Join-Path $BackupRoot $timestamp
  New-Item -ItemType Directory -Force -Path $backupDir | Out-Null

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($WorkbookPath)

  foreach ($component in @($workbook.VBProject.VBComponents)) {
    $backupPath = Get-XlflowComponentPath -Component $component -ModulesDir $backupDir -ClassesDir $backupDir -FormsDir $backupDir -WorkbookDir $backupDir
    if ($null -ne $backupPath) {
      $component.Export($backupPath)
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
        $workbook.VBProject.VBComponents.Import($file.FullName) | Out-Null
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

  $workbook.Save()
  $result.workbook = [ordered]@{ path = $WorkbookPath }
  $result.backup = [ordered]@{ path = $backupDir }
  $result.logs = @(
    "backed up existing VBA components",
    "imported $($imported.Count) source file(s)",
    "updated $($updatedDocumentModules.Count) workbook module(s)"
  )
} catch {
  Set-XlflowError -Result $result -Code "excel_import_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
