param(
  [string]$WorkbookPath,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "doctor"
$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = $false

  $diagnostics = [ordered]@{
    excel_installed = $true
    workbook_openable = $false
    vbide_access = $false
    fix = $null
  }

  if ($WorkbookPath -and (Test-Path -LiteralPath $WorkbookPath)) {
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $diagnostics.workbook_openable = $true
    try {
      $null = $workbook.VBProject.VBComponents.Count
      $diagnostics.vbide_access = $true
    } catch {
      $diagnostics.fix = "Enable 'Trust access to the VBA project object model' in Excel Trust Center."
    }
  }

  $result.diagnostics = $diagnostics
  $result.workbook = [ordered]@{ path = $WorkbookPath }
  if (-not $diagnostics.vbide_access) {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel"
  }
} catch {
  Set-XlflowError -Result $result -Code "excel_unavailable" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
