param(
  [string]$WorkbookPath,
  [string]$Filter = "",
  [string]$Visible = "false",
  [string]$RuntimeMode = "test",
  [string]$RuntimeSource = "command",
  [string]$MsgBoxResponsesJSON = "",
  [string]$InputResponsesJSON = "",
  [string]$FileDialogResponsesJSON = "",
  [string]$DebugStreamEnabled = "false",
  [string]$DebugStreamPipeName = "",
  [string]$UIStreamEnabled = "false",
  [string]$UIStreamRedactInput = "true",
  [string]$UIStreamPipeName = "",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "test"
$excel = $null
$workbook = $null
$runnerComponent = $null
$sessionAttached = $false
$sessionMode = "none"
$runtimeState = $null

try {
  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "false" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $runtimeState = Start-XlflowRuntimeInjection -Workbook $workbook -Result $result -Mode $RuntimeMode -Source $RuntimeSource -MsgBoxResponsesJSON $MsgBoxResponsesJSON -InputResponsesJSON $InputResponsesJSON -FileDialogResponsesJSON $FileDialogResponsesJSON -DebugStreamEnabled $DebugStreamEnabled -DebugStreamPipeName $DebugStreamPipeName -UIStreamEnabled $UIStreamEnabled -UIStreamPipeName $UIStreamPipeName -UIStreamRedactInput $UIStreamRedactInput

  try {
    $project = $workbook.VBProject
    [void](Enable-XlflowUIStreamRuntimeInjection -Workbook $workbook -State $runtimeState -VBProject $project)
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel"
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
    Write-XlflowJson -Result $result
    exit
  }

  $discovered = New-Object System.Collections.Generic.List[object]
  foreach ($component in @($project.VBComponents)) {
    $code = Get-XlflowCodeModuleText -CodeModule $component.CodeModule
    $moduleTests = Find-XlflowTestProcedures -ModuleName $component.Name -Code $code
    foreach ($test in @($moduleTests)) {
      if ($null -ne $test) {
        $discovered.Add($test) | Out-Null
      }
    }
  }

  if ($discovered.Count -eq 0) {
    Set-XlflowError -Result $result -Code "no_tests_found" -Message "no VBA tests found"
    $result["workbook"] = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
    $result["tests"] = @()
    Write-XlflowJson -Result $result
    exit
  }

  $duplicates = @($discovered | Group-Object -Property name | Where-Object { $_.Count -gt 1 })
  if ($duplicates.Count -gt 0) {
    $names = ($duplicates | ForEach-Object { $_.Name }) -join ", "
    Set-XlflowError -Result $result -Code "duplicate_test_name" -Message ("duplicate VBA test name(s): " + $names)
    $result["workbook"] = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
    $result["tests"] = $discovered.ToArray()
    Write-XlflowJson -Result $result
    exit
  }

  $selected = @(Select-XlflowTests -Tests $discovered -Filter $Filter)
  if ($selected.Count -eq 0) {
    Set-XlflowError -Result $result -Code "test_not_found" -Message ("test not found: " + $Filter)
    $result["workbook"] = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
    $result["tests"] = @()
    Write-XlflowJson -Result $result
    exit
  }

  $results = New-Object System.Collections.Generic.List[object]
  $logs = New-Object System.Collections.Generic.List[string]
  $failed = 0
  $runnerName = "XlflowTestRunner" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
  $runnerComponent = $project.VBComponents.Add(1)
  $runnerComponent.Name = $runnerName
  $runnerComponent.CodeModule.AddFromString((New-XlflowTestRunnerCode -Tests $selected))

  $testIndex = 0
  foreach ($test in $selected) {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
      $runResult = $excel.Run(($runnerName + ".RunTest"), $testIndex)
      $stopwatch.Stop()
      if ([bool]$runResult[0]) {
        $results.Add([pscustomobject][ordered]@{
          name = $test.name
          module = $test.module
          status = "passed"
          duration_ms = [int]$stopwatch.ElapsedMilliseconds
        }) | Out-Null
        $logs.Add("PASS " + $test.name) | Out-Null
      } else {
        $failed++
        $message = [string]$runResult[3]
        $results.Add([pscustomobject][ordered]@{
          name = $test.name
          module = $test.module
          status = "failed"
          duration_ms = [int]$stopwatch.ElapsedMilliseconds
          error = [ordered]@{
            code = "test_failed"
            message = $message
            source = [string]$runResult[2]
            number = [int]$runResult[1]
          }
        }) | Out-Null
        $logs.Add("FAIL " + $test.name + ": " + $message) | Out-Null
      }
    } catch {
      $stopwatch.Stop()
      $failed++
      $message = $_.Exception.Message
      $results.Add([pscustomobject][ordered]@{
        name = $test.name
        module = $test.module
        status = "failed"
        duration_ms = [int]$stopwatch.ElapsedMilliseconds
        error = [ordered]@{
          code = "test_failed"
          message = $message
          source = $_.Exception.Source
          number = $_.Exception.HResult
        }
      }) | Out-Null
      $logs.Add("FAIL " + $test.name + ": " + $message) | Out-Null
    }
    $testIndex++
  }

  $project.VBComponents.Remove($runnerComponent)
  $runnerComponent = $null
  if ($null -ne $runtimeState) {
    Restore-XlflowRuntimeInjection -Workbook $workbook -State $runtimeState
    $runtimeState = $null
  }
  $workbook.Save()
  $result["workbook"] = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $true -NeedsSave $false -Dirty $false
  $result["tests"] = $results.ToArray()
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode)) + $logs.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  if ($failed -gt 0) {
    Set-XlflowError -Result $result -Code "test_failed" -Message ("$failed of $($selected.Count) test(s) failed")
  }
} catch {
  Set-XlflowError -Result $result -Code "test_environment_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result["workbook"] = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
} finally {
  if ($null -ne $runnerComponent) {
    try { $workbook.VBProject.VBComponents.Remove($runnerComponent) | Out-Null } catch { Write-Verbose ("failed to remove test harness module: " + $_.Exception.Message) }
  }
  if ($null -ne $runtimeState) {
    try {
      Restore-XlflowRuntimeInjection -Workbook $workbook -State $runtimeState
    } catch {
      Write-Verbose ("failed to restore runtime injection state: " + $_.Exception.Message)
    }
  }
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
