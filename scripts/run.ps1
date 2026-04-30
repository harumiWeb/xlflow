param(
  [string]$WorkbookPath,
  [string]$MacroName,
  [string]$MacroArgsJson = "[]",
  [string]$Visible = "false",
  [string]$DisplayAlerts = "false",
  [string]$SaveWorkbook = "false",
  [string]$SaveAsPath = "",
  [string]$TraceEnabled = "false",
  [string]$TraceFile = "",
  [int]$TimeoutSeconds = 0
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "run"
$excel = $null
$workbook = $null
$vbProject = $null
$runnerComponent = $null
$traceRequested = ConvertTo-XlflowBool $TraceEnabled
$currentPhase = "initialize"

try {
  if ($traceRequested) {
    if ([string]::IsNullOrWhiteSpace($TraceFile)) {
      $TraceFile = Join-Path (Join-Path ([System.IO.Path]::GetTempPath()) "xlflow") ("trace-" + [guid]::NewGuid().ToString("N") + ".log")
    }
    $traceDir = Split-Path -Parent $TraceFile
    if (-not [string]::IsNullOrWhiteSpace($traceDir)) {
      New-Item -ItemType Directory -Force -Path $traceDir | Out-Null
    }
    Set-Content -LiteralPath $TraceFile -Value "" -NoNewline
    $result.trace = [ordered]@{
      enabled = $true
      path = $TraceFile
      events = @()
      read_error = $null
    }
  }

  $currentPhase = "open_workbook"
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
    $currentPhase = "prepare_vbide"
    $vbProject = $workbook.VBProject
    if ($traceRequested -and -not (Test-XlflowTraceModuleInjected -VBProject $vbProject)) {
      Set-XlflowError -Result $result -Code "trace_not_injected" -Message "XlflowTrace module is not injected. Run xlflow trace inject before xlflow run --trace." -Source "xlflow" -Phase $currentPhase
      throw "XlflowTrace module is not injected"
    }
    $currentPhase = "inject_harness"
    $runnerComponent = $vbProject.VBComponents.Add(1)
  } catch {
    if ($result.error -eq $null) {
      Set-XlflowError -Result $result -Code "vbide_access_denied" -Message $_.Exception.Message -Source "vbide" -Phase $currentPhase
    }
    throw
  }

  $runnerName = New-XlflowRunHarnessModuleName
  $runnerComponent.Name = $runnerName
  $runnerComponent.CodeModule.AddFromString((New-XlflowRunHarnessCode -MacroName $MacroName -Arguments $argumentSpecs -TraceEnabled $traceRequested -TraceFile $TraceFile))

  $currentPhase = "invoke_macro"
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
    $errorCode = "macro_failed"
    if (Test-XlflowMacroTargetFailure -Number ([int]$runResult[2]) -Description ([string]$runResult[3])) {
      $errorCode = "macro_not_found"
    }
    Set-XlflowError -Result $result -Code $errorCode -Message $failureMessage -Source ([string]$runResult[1]) -Number ([int]$runResult[2]) -Line ([int]$runResult[4]) -Phase $currentPhase
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $null }
  } elseif (ConvertTo-XlflowBool $SaveWorkbook) {
    $currentPhase = "save_result"
    $workbook.Save()
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true; save_as = $null }
    $result.logs = @($successLog, "saved workbook in place")
  } elseif (-not [string]::IsNullOrWhiteSpace($SaveAsPath)) {
    $currentPhase = "save_result"
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
    $errorCode = "macro_failed"
    if ($currentPhase -eq "invoke_macro" -and (Test-XlflowMacroTargetFailure -Number ([int]$_.Exception.HResult) -Description $_.Exception.Message)) {
      $errorCode = "macro_not_found"
    }
    Set-XlflowError -Result $result -Code $errorCode -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult -Phase $currentPhase
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
  if ($traceRequested) {
    $currentPhase = "read_trace"
    if ($null -eq $result.trace) {
      $result.trace = [ordered]@{
        enabled = $true
        path = $TraceFile
        events = @()
        read_error = $null
      }
    }
    try {
      $events = @(Read-XlflowTraceEvents -Path $TraceFile)
      $result.trace.events = @($events)
      if ($result.status -eq "failed" -and $events.Count -eq 0) {
        $result.trace.hint = "no trace events were written; execution may have failed before reaching user XlflowLog calls"
        $result.logs += $result.trace.hint
      }
      foreach ($event in $events) {
        $result.logs += ("[" + $event.timestamp + "] " + $event.message)
      }
    } catch {
      $result.trace.read_error = $_.Exception.Message
    }
  }
}

Write-XlflowJson -Result $result
