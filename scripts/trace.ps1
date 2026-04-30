param(
  [string]$Action = "inject",
  [string]$WorkbookPath,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "trace"
$excel = $null
$workbook = $null

try {
  if ($Action -ne "inject") {
    Set-XlflowError -Result $result -Code "trace_args_invalid" -Message ("unsupported trace action: " + $Action) -Source "xlflow"
    throw "unsupported trace action"
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($WorkbookPath)

  try {
    $vbProject = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message $_.Exception.Message -Source "vbide"
    throw
  }

  try {
    $existing = $vbProject.VBComponents.Item("XlflowTrace")
    $vbProject.VBComponents.Remove($existing)
  } catch {}

  $component = $vbProject.VBComponents.Add(1)
  $component.Name = "XlflowTrace"
  $component.CodeModule.AddFromString((New-XlflowTraceModuleCode))
  $workbook.Save()

  $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true }
  $result.logs = @("injected XlflowTrace into " + $WorkbookPath)
} catch {
  if ($result.error -eq $null) {
    Set-XlflowError -Result $result -Code "trace_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  if ($null -eq $result.workbook) {
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false }
  }
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
