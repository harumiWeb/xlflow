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
$sessionAttached = $false
$sessionMode = "none"

try {
  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode

  try {
    $project = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel"
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
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

  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
  $result.macros = $macros.ToArray()
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "discovered $($macros.Count) macro entrypoint(s)") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  Set-XlflowError -Result $result -Code "macro_discovery_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
} finally {
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
