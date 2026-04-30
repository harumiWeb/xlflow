param(
  [string]$WorkbookPath,
  [string]$MacroName,
  [string]$MacroArgsJson = "[]",
  [string]$Visible = "false",
  [string]$DisplayAlerts = "false",
  [string]$SaveWorkbook = "false",
  [string]$SaveAsPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "run"
$excel = $null
$workbook = $null
$vbProject = $null
$runnerComponent = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = ConvertTo-XlflowBool $DisplayAlerts
  $workbook = $excel.Workbooks.Open($WorkbookPath)

  $typedValues = @(ConvertFrom-XlflowRunArgumentsJson -Json $MacroArgsJson)
  $argumentSpecs = @()
  if (-not [string]::IsNullOrWhiteSpace($MacroArgsJson)) {
    # Decode base64 JSON to get specs for harness code generation
    $decodedBytes = [System.Convert]::FromBase64String($MacroArgsJson)
    $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
    $argumentSpecs = ConvertFrom-Json -InputObject $decodedJson
  }

  try {
    $vbProject = $workbook.VBProject
    $runnerComponent = $vbProject.VBComponents.Add(1)
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message $_.Exception.Message -Source "vbide"
    throw
  }

  $runnerName = New-XlflowRunHarnessModuleName
  $runnerComponent.Name = $runnerName
  $runnerComponent.CodeModule.AddFromString((New-XlflowRunHarnessCode -MacroName $MacroName -Arguments $argumentSpecs))

  $runResult = $excel.Run($runnerName + ".RunMacro")
  $successLog = "ran " + $MacroName + " in " + ([int]$runResult[5]) + "ms"
  if ($null -ne $runnerComponent) {
    $vbProject.VBComponents.Remove($runnerComponent)
    $runnerComponent = $null
  }
  $result.macro = [ordered]@{
    name = $MacroName
    args = @($typedValues)
    duration_ms = [int]$runResult[5]
  }

  if (-not [bool]$runResult[0]) {
    $failureMessage = Format-XlflowMacroFailureMessage -ModuleName ([string]$runResult[1]) -Line ([int]$runResult[4]) -Number ([int]$runResult[2]) -Description ([string]$runResult[3])
    Set-XlflowError -Result $result -Code "macro_failed" -Message $failureMessage -Source ([string]$runResult[1]) -Number ([int]$runResult[2]) -Line ([int]$runResult[4])
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $null }
  } elseif (ConvertTo-XlflowBool $SaveWorkbook) {
    $workbook.Save()
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true; save_as = $null }
    $result.logs = @($successLog, "saved workbook in place")
  } elseif (-not [string]::IsNullOrWhiteSpace($SaveAsPath)) {
    Assert-XlflowSaveAsExtension -WorkbookPath $WorkbookPath -SaveAsPath $SaveAsPath
    $targetDir = Split-Path -Parent $SaveAsPath
    if (-not [string]::IsNullOrWhiteSpace($targetDir)) {
      New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
    }
    $workbook.SaveCopyAs($SaveAsPath)
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $SaveAsPath }
    $result.logs = @($successLog, ("wrote workbook copy to " + $SaveAsPath))
  } else {
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $null }
    $result.logs = @($successLog, "left workbook unchanged on disk")
  }
} catch {
  if ($result.error -eq $null) {
    Set-XlflowError -Result $result -Code "macro_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  if ($null -eq $result.macro) {
    $result.macro = [ordered]@{ name = $MacroName; args = @(); duration_ms = 0 }
  }
  $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $null }
} finally {
  if ($null -ne $runnerComponent -and $null -ne $vbProject) {
    try { $vbProject.VBComponents.Remove($runnerComponent) } catch {}
  }
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
