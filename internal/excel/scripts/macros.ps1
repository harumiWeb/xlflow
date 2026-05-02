param(
  [string]$WorkbookPath,
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "macros"
$excel = $null
$workbook = $null

try {
  if (ConvertTo-XlflowBool $UseSession) {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
  } else {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = ConvertTo-XlflowBool $Visible
    $workbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts $false -DisableAutomationMacros $true
  }

  try {
    $project = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel"
    $result.workbook = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
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

  $result.workbook = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
  $result.macros = $macros.ToArray()
  $result.logs = @("discovered $($macros.Count) macro entrypoint(s)")
} catch {
  Set-XlflowError -Result $result -Code "macro_discovery_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result.workbook = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
} finally {
  if (ConvertTo-XlflowBool $UseSession) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
