param(
  [string]$WorkbookPath,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "macros"
$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($WorkbookPath)

  try {
    $project = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel"
    $result.workbook = [ordered]@{ path = $WorkbookPath }
    Write-XlflowJson -Result $result
    exit
  }

  $macros = New-Object System.Collections.Generic.List[object]
  foreach ($component in @($project.VBComponents)) {
    if ($component.Name -like "Xlflow*") {
      continue
    }
    $code = Get-XlflowCodeModuleText -CodeModule $component.CodeModule
    $componentMacros = Find-XlflowMacroProcedures -ModuleName $component.Name -Code $code
    foreach ($macro in @($componentMacros)) {
      if ($null -ne $macro) {
        $macros.Add($macro) | Out-Null
      }
    }
  }

  $result.workbook = [ordered]@{ path = $WorkbookPath }
  $result.macros = $macros.ToArray()
  $result.logs = @("discovered $($macros.Count) macro entrypoint(s)")
} catch {
  Set-XlflowError -Result $result -Code "macro_discovery_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result.workbook = [ordered]@{ path = $WorkbookPath }
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
