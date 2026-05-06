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
  [string]$Direct = "false",
  [string]$Diagnostic = "false",
  [string]$SuppressModalErrors = "true",
  [string]$UseSession = "false",
  [string]$MetadataPath = "",
  [int]$TimeoutSeconds = 0
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "run"
$excel = $null
$workbook = $null
$vbProject = $null
$runnerComponent = $null
$traceRequested = ConvertTo-XlflowBool $TraceEnabled
$traceTempInjected = $false
$currentPhase = "initialize"
$sessionAttached = $false
$sessionMode = "none"
$suppressModalErrors = ConvertTo-XlflowBool $SuppressModalErrors

function Get-XlflowRunFailureCode {
  param(
    [int]$Number,
    [string]$Description
  )

  if (Test-XlflowMacroDisabledFailure -Number $Number -Description $Description) {
    return "macro_disabled"
  }
  if (Test-XlflowMacroTargetFailure -Number $Number -Description $Description) {
    return "macro_not_found"
  }
  return "macro_failed"
}

function New-XlflowRuntimeDialogDiagnostic {
  param($Dialog, $Selection)

  $diag = [ordered]@{
    kind = "runtime"
    dialog = $Dialog
  }
  $messageLines = @(Get-XlflowExcelDialogMessageLines -Dialog $Dialog)
  if ($messageLines.Count -gt 0) {
    $diag.message = @($messageLines)
  }
  if ($null -ne $Selection) {
    $diag.location = $Selection.location
    $diag.nearby_code = @($Selection.nearby_code)
  }
  return $diag
}

function Set-XlflowRuntimeDialogFailure {
  param(
    [string]$ErrorCode,
    [string]$FallbackSource,
    [int]$FallbackNumber,
    [int]$FallbackLine,
    $Dialog,
    $Selection
  )

  $messageLines = @(Get-XlflowExcelDialogMessageLines -Dialog $Dialog)
  $message = ($messageLines -join [Environment]::NewLine)
  if ([string]::IsNullOrWhiteSpace($message)) {
    $message = "VBA runtime error dialog was shown."
  }

  $source = $FallbackSource
  $line = $FallbackLine
  if ($null -ne $Selection -and $null -ne $Selection.location) {
    if (-not [string]::IsNullOrWhiteSpace([string]$Selection.location.module)) {
      $source = [string]$Selection.location.module
    }
    if ([int]$Selection.location.line -gt 0) {
      $line = [int]$Selection.location.line
    }
  }

  $number = Get-XlflowVBARuntimeDialogErrorNumber -Dialog $Dialog
  if ($number -eq 0) {
    $number = $FallbackNumber
  }

  Set-XlflowError -Result $result -Code $ErrorCode -Message $message -Source $source -Number $number -Line $line -Phase $currentPhase
  $result.run_diagnostic = New-XlflowRuntimeDialogDiagnostic -Dialog $Dialog -Selection $Selection
}

try {
  if ($TimeoutSeconds -lt 0) {
    Set-XlflowError -Result $result -Code "run_args_invalid" -Message "-TimeoutSeconds must be greater than or equal to 0." -Source "xlflow"
    throw "invalid timeout"
  }
  if ((ConvertTo-XlflowBool $Direct) -and (ConvertTo-XlflowBool $Diagnostic)) {
    Set-XlflowError -Result $result -Code "run_args_invalid" -Message "-Direct cannot be used with diagnostic mode." -Source "xlflow" -Phase $currentPhase
    throw "invalid direct run"
  }
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
  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts $DisplayAlerts -DisableAutomationMacros "false" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode

  $typedValues = @(ConvertFrom-XlflowRunArgumentsJson -Json $MacroArgsJson)
  $argumentSpecs = @()
  if (-not [string]::IsNullOrWhiteSpace($MacroArgsJson)) {
    # Decode base64 JSON to get specs for harness code generation
    $decodedBytes = [System.Convert]::FromBase64String($MacroArgsJson)
    $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
    $argumentSpecs = ConvertFrom-Json -InputObject $decodedJson
  }

  if (ConvertTo-XlflowBool $Direct) {
    if ($traceRequested) {
      Set-XlflowError -Result $result -Code "run_args_invalid" -Message "-Direct cannot be used with trace." -Source "xlflow" -Phase $currentPhase
      throw "invalid direct run"
    }
    if ($typedValues.Count -gt 0) {
      Set-XlflowError -Result $result -Code "run_args_invalid" -Message "-Direct cannot be used with macro arguments." -Source "xlflow" -Phase $currentPhase
      throw "invalid direct run"
    }
    $currentPhase = "invoke_macro"
    $startedAt = Get-Date
    $invokeResult = Invoke-XlflowExcelCallWithDialogWatch -Excel $excel -Workbook $workbook -Invocation { $excel.Run($MacroName) } -DialogKind "runtime" -CaptureDialogs ([bool]$suppressModalErrors)
    $durationMs = [int]((Get-Date) - $startedAt).TotalMilliseconds
    if ([bool]$invokeResult.dialog.found) {
      $errorCode = "macro_failed"
      $failureNumber = 0
      $failureDescription = ""
      if ($null -ne $invokeResult.exception -and $null -ne $invokeResult.exception.Exception) {
        $failureNumber = [int]$invokeResult.exception.Exception.HResult
        $failureDescription = [string]$invokeResult.exception.Exception.Message
        $errorCode = Get-XlflowRunFailureCode -Number $failureNumber -Description $failureDescription
      }
      Set-XlflowRuntimeDialogFailure -ErrorCode $errorCode -FallbackSource (Get-XlflowMacroModuleName -MacroName $MacroName) -FallbackNumber $failureNumber -FallbackLine 0 -Dialog $invokeResult.dialog -Selection $invokeResult.selection
      $result.macro = [ordered]@{
        name = $MacroName
        args = @($typedValues)
        duration_ms = $durationMs
        direct = $true
      }
      $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath "" -NeedsSave $saveState.needs_save -Dirty $saveState.dirty
      throw "runtime dialog shown"
    }
    if ($null -ne $invokeResult.exception) {
      throw $invokeResult.exception.Exception
    }
    $successLog = "ran " + $MacroName + " directly in " + $durationMs + "ms"
    $result.macro = [ordered]@{
      name = $MacroName
      args = @($typedValues)
      duration_ms = $durationMs
      direct = $true
    }
    if (ConvertTo-XlflowBool $SaveWorkbook) {
      $currentPhase = "save_result"
      $workbook.Save()
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $true -SaveAsPath ""
      $result.logs = @(
        @(
          $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
          $successLog,
          "saved workbook in place"
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      )
    } elseif (-not [string]::IsNullOrWhiteSpace($SaveAsPath)) {
      $currentPhase = "save_result"
      Assert-XlflowSaveAsExtension -WorkbookPath $WorkbookPath -SaveAsPath $SaveAsPath
      $targetDir = Split-Path -Parent $SaveAsPath
      if (-not [string]::IsNullOrWhiteSpace($targetDir)) {
        New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
      }
      $workbook.SaveCopyAs($SaveAsPath)
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath $SaveAsPath -NeedsSave $sessionAttached -Dirty $sessionAttached
      $result.logs = @(
        @(
          $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
          $successLog,
          ("wrote workbook copy to " + $SaveAsPath)
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      )
    } else {
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath "" -NeedsSave $sessionAttached -Dirty $sessionAttached
      $result.logs = @(
        @(
          $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
          $successLog,
          $(if ($sessionAttached) {
            "SAVE REQUIRED: live session workbook differs from disk; run xlflow save before session stop"
          } else {
            "left workbook unchanged on disk"
          })
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      )
    }
    Write-XlflowJson -Result $result
    return
  }

  try {
    $currentPhase = "prepare_vbide"
    $vbProject = $workbook.VBProject
    if ($traceRequested -and -not (Test-XlflowTraceModuleInjected -VBProject $vbProject)) {
      $traceComponent = $vbProject.VBComponents.Add(1)
      $traceComponent.Name = "XlflowTrace"
      $traceComponent.CodeModule.AddFromString((New-XlflowTraceModuleCode))
      $traceTempInjected = $true
      $result.trace.lifecycle = "temporary"
      $result.trace.temporary_injected = $true
    } elseif ($traceRequested) {
      $result.trace.lifecycle = "existing"
    }
    if (ConvertTo-XlflowBool $Diagnostic) {
      $currentPhase = "compile_vba"
      $compileResult = Invoke-XlflowVBECompile -Excel $excel -Workbook $workbook
      if (-not [bool]$compileResult.ok) {
        $location = $compileResult.selection.location
        $messageLines = @($compileResult.dialog.text)
        if ($messageLines.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($compileResult.dialog.title)) {
          $messageLines = @([string]$compileResult.dialog.title)
        }
        $message = ($messageLines -join [Environment]::NewLine)
        if ([string]::IsNullOrWhiteSpace($message)) {
          $message = "VBA compile failed."
        }
        Set-XlflowError -Result $result -Code "vba_compile_failed" -Message $message -Source ([string]$location.module) -Line ([int]$location.line) -Phase $currentPhase
        $result.run_diagnostic = [ordered]@{
          kind = "compile"
          message = @($messageLines)
          location = $location
          nearby_code = @($compileResult.selection.nearby_code)
          dialog = $compileResult.dialog
        }
        $result.macro = [ordered]@{ name = $MacroName; args = @($typedValues); duration_ms = 0; direct = $false }
        $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath "" -NeedsSave $false -Dirty $false
        throw "vba compile failed"
      }
    }
    $currentPhase = "verify_macro"
    if (-not (Test-XlflowMacroExists -Workbook $workbook -MacroName $MacroName)) {
      Set-XlflowError -Result $result -Code "macro_not_found" -Message ("Macro not found: " + $MacroName) -Source "Excel" -Phase $currentPhase
      throw "macro target missing"
    }
    $currentPhase = "inject_harness"
    $runnerComponent = $vbProject.VBComponents.Add(1)
  } catch {
    if ($null -eq $result.error) {
      Set-XlflowError -Result $result -Code "vbide_access_denied" -Message $_.Exception.Message -Source "vbide" -Phase $currentPhase
    }
    throw
  }

  $runnerName = New-XlflowRunHarnessModuleName
  $runnerComponent.Name = $runnerName
  $runnerComponent.CodeModule.AddFromString((New-XlflowRunHarnessCode -MacroName $MacroName -Arguments $argumentSpecs -TraceEnabled $traceRequested -TraceFile $TraceFile))

  $currentPhase = "invoke_macro"
  $startedAt = Get-Date
  $invokeResult = Invoke-XlflowExcelCallWithDialogWatch -Excel $excel -Workbook $workbook -Invocation { $excel.Run($runnerName + ".RunMacro") } -DialogKind "runtime" -CaptureDialogs ([bool]$suppressModalErrors)
  $runResult = $invokeResult.value
  $durationMs = [int]((Get-Date) - $startedAt).TotalMilliseconds
  if ($null -ne $runResult -and $runResult.Count -gt 5) {
    $durationMs = [int]$runResult[5]
  }
  if ([bool]$invokeResult.dialog.found) {
    $errorCode = "macro_failed"
    $failureSource = Get-XlflowMacroModuleName -MacroName $MacroName
    $failureNumber = 0
    $failureLine = 0
    $failureDescription = ""
    if ($null -ne $runResult -and $runResult.Count -gt 4) {
      $failureSource = [string]$runResult[1]
      $failureNumber = [int]$runResult[2]
      $failureDescription = [string]$runResult[3]
      $failureLine = [int]$runResult[4]
      $errorCode = Get-XlflowRunFailureCode -Number $failureNumber -Description $failureDescription
    } elseif ($null -ne $invokeResult.exception -and $null -ne $invokeResult.exception.Exception) {
      $failureSource = [string]$invokeResult.exception.Exception.Source
      $failureNumber = [int]$invokeResult.exception.Exception.HResult
      $failureDescription = [string]$invokeResult.exception.Exception.Message
      $errorCode = Get-XlflowRunFailureCode -Number $failureNumber -Description $failureDescription
    }
    Set-XlflowRuntimeDialogFailure -ErrorCode $errorCode -FallbackSource $failureSource -FallbackNumber $failureNumber -FallbackLine $failureLine -Dialog $invokeResult.dialog -Selection $invokeResult.selection
    $result.macro = [ordered]@{
      name = $MacroName
      args = @($typedValues)
      duration_ms = $durationMs
      direct = $false
    }
    $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath "" -NeedsSave $saveState.needs_save -Dirty $saveState.dirty
    throw "runtime dialog shown"
  }
  if ($null -ne $invokeResult.exception) {
    throw $invokeResult.exception.Exception
  }
  $successLog = "ran " + $MacroName + " in " + $durationMs + "ms"
  if ($null -ne $runnerComponent) {
    $vbProject.VBComponents.Remove($runnerComponent)
    $runnerComponent = $null
  }
  if ($traceTempInjected -and $null -ne $vbProject) {
    [void](Remove-XlflowTraceModule -VBProject $vbProject)
    $traceTempInjected = $false
    if ($null -ne $result.trace) {
      $result.trace.temporary_reverted = $true
    }
  }
  $result.macro = [ordered]@{
    name = $MacroName
    args = @($typedValues)
    duration_ms = $durationMs
    direct = $false
  }

  if (-not [bool]$runResult[0]) {
    $failureMessage = Format-XlflowMacroFailureMessage -ModuleName ([string]$runResult[1]) -Line ([int]$runResult[4]) -Number ([int]$runResult[2]) -Description ([string]$runResult[3])
    $errorCode = "macro_failed"
    if (Test-XlflowMacroTargetFailure -Number ([int]$runResult[2]) -Description ([string]$runResult[3])) {
      $errorCode = "macro_not_found"
    }
    Set-XlflowError -Result $result -Code $errorCode -Message $failureMessage -Source ([string]$runResult[1]) -Number ([int]$runResult[2]) -Line ([int]$runResult[4]) -Phase $currentPhase
    $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath "" -NeedsSave $saveState.needs_save -Dirty $saveState.dirty
  } elseif (ConvertTo-XlflowBool $SaveWorkbook) {
    $currentPhase = "save_result"
    $workbook.Save()
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $true -SaveAsPath ""
    $result.logs = @(
      @(
        $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
        $successLog,
        "saved workbook in place"
      ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
  } elseif (-not [string]::IsNullOrWhiteSpace($SaveAsPath)) {
    $currentPhase = "save_result"
    Assert-XlflowSaveAsExtension -WorkbookPath $WorkbookPath -SaveAsPath $SaveAsPath
    $targetDir = Split-Path -Parent $SaveAsPath
    if (-not [string]::IsNullOrWhiteSpace($targetDir)) {
      New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
    }
    $workbook.SaveCopyAs($SaveAsPath)
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath $SaveAsPath -NeedsSave $sessionAttached -Dirty $sessionAttached
    $result.logs = @(
      @(
        $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
        $successLog,
        ("wrote workbook copy to " + $SaveAsPath)
      ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
  } else {
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath "" -NeedsSave $sessionAttached -Dirty $sessionAttached
    $result.logs = @(
      @(
        $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
        $successLog,
        $(if ($sessionAttached) {
          "SAVE REQUIRED: live session workbook differs from disk; run xlflow save before session stop"
        } else {
          "left workbook unchanged on disk"
        })
      ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
  }
} catch {
  if ($null -eq $result.error) {
    $errorCode = "macro_failed"
    if ($currentPhase -eq "invoke_macro") {
      if (Test-XlflowMacroDisabledFailure -Number ([int]$_.Exception.HResult) -Description $_.Exception.Message) {
        $errorCode = "macro_disabled"
      } elseif (Test-XlflowMacroTargetFailure -Number ([int]$_.Exception.HResult) -Description $_.Exception.Message) {
        $errorCode = "macro_not_found"
      }
    }
    Set-XlflowError -Result $result -Code $errorCode -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult -Phase $currentPhase
  }
  if ($null -eq $result.macro) {
    $result.macro = [ordered]@{ name = $MacroName; args = @(); duration_ms = 0; direct = (ConvertTo-XlflowBool $Direct) }
  }
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -SaveAsPath "" -NeedsSave $saveState.needs_save -Dirty $saveState.dirty
} finally {
  if ($null -ne $runnerComponent -and $null -ne $vbProject) {
    try { $vbProject.VBComponents.Remove($runnerComponent) } catch { Write-Verbose ("failed to remove run harness module: " + $_.Exception.Message) }
  }
  if ($traceTempInjected -and $null -ne $vbProject) {
    try {
      [void](Remove-XlflowTraceModule -VBProject $vbProject)
      if ($null -ne $result.trace) {
        $result.trace.temporary_reverted = $true
      }
    } catch {
      Write-Verbose ("failed to remove temporary trace module: " + $_.Exception.Message)
    }
  }
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
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
      foreach ($traceEvent in $events) {
        $result.logs += ("[" + $traceEvent.timestamp + "] " + $traceEvent.message)
      }
    } catch {
      $result.trace.read_error = $_.Exception.Message
    }
  }
}

Write-XlflowJson -Result $result
