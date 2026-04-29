param(
  [string]$WorkbookPath,
  [string]$MacroName,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "run"
$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($WorkbookPath)
  $excel.Run($MacroName)
  $workbook.Save()
  $result.macro = $MacroName
  $result.workbook = [ordered]@{ path = $WorkbookPath }
  $result.logs = @("macro executed successfully")
} catch {
  Set-XlflowError -Result $result -Code "macro_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result.macro = $MacroName
  $result.workbook = [ordered]@{ path = $WorkbookPath }
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
