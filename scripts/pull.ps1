param(
  [string]$WorkbookPath,
  [string]$ModulesDir,
  [string]$ClassesDir,
  [string]$FormsDir,
  [string]$WorkbookDir,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "pull"
$excel = $null
$workbook = $null

try {
  New-Item -ItemType Directory -Force -Path $ModulesDir, $ClassesDir, $FormsDir, $WorkbookDir | Out-Null
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($WorkbookPath)

  $exported = @()
  foreach ($component in $workbook.VBProject.VBComponents) {
    $path = Get-XlflowComponentPath -Component $component -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir
    if ($null -ne $path) {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      $component.Export($path)
      Convert-XlflowExportedSourceToUtf8 -Path $path
      if ($component.Type -eq 100) {
        Normalize-XlflowDocumentModuleFile -Path $path
      }
      $exported += $path
    }
  }

  $result.workbook = [ordered]@{ path = $WorkbookPath }
  $result.logs = @("exported $($exported.Count) VBA component(s)")
} catch {
  Set-XlflowError -Result $result -Code "excel_export_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
