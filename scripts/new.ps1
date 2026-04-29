param(
  [string]$WorkbookPath
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "new"
$excel = $null
$workbook = $null

try {
  $parent = Split-Path -Parent $WorkbookPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Add()
  $workbook.SaveAs($WorkbookPath, 52)

  $result.workbook = [ordered]@{ path = $WorkbookPath }
  $result.logs = @("created workbook $WorkbookPath")
} catch {
  Set-XlflowError -Result $result -Code "excel_create_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
