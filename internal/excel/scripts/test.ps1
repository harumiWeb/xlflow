param(
  [string]$WorkbookPath,
  [string]$Filter = "",
  [string]$ModuleFilter = "",
  [string]$TagFilter = "",
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

  $selected = @(Select-XlflowTests -Tests $discovered -Filter $Filter -ModuleFilter $ModuleFilter -TagFilter $TagFilter)
  if ($selected.Count -eq 0) {
    $filterDesc = $Filter
    if ([string]::IsNullOrWhiteSpace($filterDesc)) { $filterDesc = $ModuleFilter }
    if ([string]::IsNullOrWhiteSpace($filterDesc)) { $filterDesc = $TagFilter }
    if ([string]::IsNullOrWhiteSpace($filterDesc)) { $filterDesc = "(no filter)" }
    Set-XlflowError -Result $result -Code "test_not_found" -Message ("test not found: " + $filterDesc)
    $result["workbook"] = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
    $result["tests"] = @()
    Write-XlflowJson -Result $result
    exit
  }

  $results = New-Object System.Collections.Generic.List[object]
  $logs = New-Object System.Collections.Generic.List[string]
  $failed = 0
  $inconclusiveCount = 0
  $runnerName = "XlflowTestRunner" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
  $runnerComponent = $project.VBComponents.Add(1)
  $runnerComponent.Name = $runnerName

  # Build hooks map before generating runner so each test case can embed its hooks directly
  $hooksByModule = @{}
  foreach ($moduleGroup in @($selected | Group-Object -Property module)) {
    $moduleName = $moduleGroup.Name
    $moduleComponent = $null
    foreach ($c in @($project.VBComponents)) {
      if ($c.Name -eq $moduleName) {
        $moduleComponent = $c
        break
      }
    }
    $moduleCode = ""
    if ($null -ne $moduleComponent) {
      $moduleCode = Get-XlflowCodeModuleText -CodeModule $moduleComponent.CodeModule
    }
    $hooksByModule[$moduleName] = Find-XlflowModuleHooks -ModuleName $moduleName -Code $moduleCode
  }

  # Assign sequential index to each selected test so the runner can dispatch via Select Case
  $testIndex = 0
  foreach ($test in $selected) {
    $test | Add-Member -MemberType NoteProperty -Name "index" -Value $testIndex -Force
    $testIndex++
  }

  $runnerComponent.CodeModule.AddFromString((New-XlflowTestRunnerCode -Tests $selected -HooksByModule $hooksByModule))

  foreach ($moduleGroup in @($selected | Group-Object -Property module)) {
    $moduleName = $moduleGroup.Name
    $hooks = $hooksByModule[$moduleName]

    # BeforeAll
    $beforeAllFailed = $false
    if ($null -ne $hooks.BeforeAll) {
      $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
      try {
        $runResult = $excel.Run(($runnerName + ".RunBeforeAll_" + $moduleName))
        $stopwatch.Stop()
        if (-not [bool]$runResult[0]) {
          $beforeAllFailed = $true
          $message = [string]$runResult[3]
          foreach ($test in $moduleGroup.Group) {
            $failed++
            $results.Add([pscustomobject][ordered]@{
              name = $test.name
              module = $test.module
              status = "failed"
              duration_ms = [int]$stopwatch.ElapsedMilliseconds
              error = [ordered]@{
                code = "before_all_failed"
                message = $message
                source = [string]$runResult[2]
                number = [int]$runResult[1]
              }
            }) | Out-Null
            $logs.Add("FAIL " + $test.name + ": before_all_failed: " + $message) | Out-Null
          }
        }
      } catch {
        $stopwatch.Stop()
        $beforeAllFailed = $true
        $message = $_.Exception.Message
        foreach ($test in $moduleGroup.Group) {
          $failed++
          $results.Add([pscustomobject][ordered]@{
            name = $test.name
            module = $test.module
            status = "failed"
            duration_ms = [int]$stopwatch.ElapsedMilliseconds
            error = [ordered]@{
              code = "before_all_failed"
              message = $message
              source = $_.Exception.Source
              number = $_.Exception.HResult
            }
          }) | Out-Null
          $logs.Add("FAIL " + $test.name + ": before_all_failed: " + $message) | Out-Null
        }
      }
    }

    if (-not $beforeAllFailed) {
      foreach ($test in $moduleGroup.Group) {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        try {
          $runResult = $excel.Run(($runnerName + ".RunTest"), $test.index)
          $stopwatch.Stop()
          $statusHint = [string]$runResult[4]
          $phaseHint = [string]$runResult[5]
          $status = "passed"
          $errorCode = ""
          $errorMessage = ""
          $errorSource = ""
          $errorNumber = 0
          if (-not [bool]$runResult[0]) {
            $status = "failed"
            $errorCode = "test_failed"
            $errorMessage = [string]$runResult[3]
            $errorSource = [string]$runResult[2]
            $errorNumber = [int]$runResult[1]
            if ($statusHint -eq "inconclusive") {
              $status = "inconclusive"
              $errorCode = "test_inconclusive"
            } else {
              switch ($phaseHint) {
                "before_each" { $errorCode = "before_each_failed" }
                "after_each"  { $errorCode = "after_each_failed" }
              }
            }
          }
          $results.Add([pscustomobject][ordered]@{
            name = $test.name
            module = $test.module
            status = $status
            duration_ms = [int]$stopwatch.ElapsedMilliseconds
            tags = $test.tags
            error = [ordered]@{
              code = $errorCode
              message = $errorMessage
              source = $errorSource
              number = $errorNumber
            }
          }) | Out-Null
          if ($status -eq "passed") {
            $logs.Add("PASS " + $test.name) | Out-Null
          } elseif ($status -eq "inconclusive") {
            $inconclusiveCount++
            $logs.Add("? " + $test.name + ": inconclusive") | Out-Null
          } else {
            $failed++
            $logs.Add("FAIL " + $test.name + ": " + $errorMessage) | Out-Null
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
            tags = $test.tags
            error = [ordered]@{
              code = "test_failed"
              message = $message
              source = $_.Exception.Source
              number = $_.Exception.HResult
            }
          }) | Out-Null
          $logs.Add("FAIL " + $test.name + ": " + $message) | Out-Null
        }
      }
    }

    # AfterAll
    if ($null -ne $hooks.AfterAll -and -not $beforeAllFailed) {
      $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
      try {
        $runResult = $excel.Run(($runnerName + ".RunAfterAll_" + $moduleName))
        $stopwatch.Stop()
        if (-not [bool]$runResult[0]) {
          $message = [string]$runResult[3]
          # Mark all tests in this module as failed (overwrite existing results)
          $moduleTestNames = @($moduleGroup.Group | ForEach-Object { $_.name })
          for ($i = 0; $i -lt $results.Count; $i++) {
            if ($results[$i].module -eq $moduleName -and $moduleTestNames -contains $results[$i].name) {
              if ($results[$i].status -eq "passed" -or $results[$i].status -eq "inconclusive") {
                if ($results[$i].status -eq "passed") { $failed++ }
                if ($results[$i].status -eq "inconclusive") { $inconclusiveCount-- }
              }
              $results[$i] = [pscustomobject][ordered]@{
                name = $results[$i].name
                module = $results[$i].module
                status = "failed"
                duration_ms = [int]$results[$i].duration_ms
                error = [ordered]@{
                  code = "after_all_failed"
                  message = $message
                  source = [string]$runResult[2]
                  number = [int]$runResult[1]
                }
              }
              $logs.Add("FAIL " + $results[$i].name + ": after_all_failed: " + $message) | Out-Null
            }
          }
        }
      } catch {
        $stopwatch.Stop()
        $message = $_.Exception.Message
        $moduleTestNames = @($moduleGroup.Group | ForEach-Object { $_.name })
        for ($i = 0; $i -lt $results.Count; $i++) {
          if ($results[$i].module -eq $moduleName -and $moduleTestNames -contains $results[$i].name) {
            if ($results[$i].status -eq "passed" -or $results[$i].status -eq "inconclusive") {
              if ($results[$i].status -eq "passed") { $failed++ }
              if ($results[$i].status -eq "inconclusive") { $inconclusiveCount-- }
            }
            $results[$i] = [pscustomobject][ordered]@{
              name = $results[$i].name
              module = $results[$i].module
              status = "failed"
              duration_ms = [int]$results[$i].duration_ms
              error = [ordered]@{
                code = "after_all_failed"
                message = $message
                source = $_.Exception.Source
                number = $_.Exception.HResult
              }
            }
            $logs.Add("FAIL " + $results[$i].name + ": after_all_failed: " + $message) | Out-Null
          }
        }
      }
    }
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
