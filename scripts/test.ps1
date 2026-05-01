param(
  [string]$WorkbookPath,
  [string]$Filter = "",
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "test"
$excel = $null
$workbook = $null
$runnerComponent = $null

try {
  if (ConvertTo-XlflowBool $UseSession) {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
  } else {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = ConvertTo-XlflowBool $Visible
    $workbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts $false -DisableAutomationMacros $false
  }

  try {
    $project = $workbook.VBProject
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel"
    $result.workbook = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
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
    $result["workbook"] = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
    $result["tests"] = @()
    Write-XlflowJson -Result $result
    exit
  }

  $duplicates = @($discovered | Group-Object -Property name | Where-Object { $_.Count -gt 1 })
  if ($duplicates.Count -gt 0) {
    $names = ($duplicates | ForEach-Object { $_.Name }) -join ", "
    Set-XlflowError -Result $result -Code "duplicate_test_name" -Message ("duplicate VBA test name(s): " + $names)
    $result["workbook"] = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
    $result["tests"] = $discovered.ToArray()
    Write-XlflowJson -Result $result
    exit
  }

  $selected = @(Select-XlflowTests -Tests $discovered -Filter $Filter)
  if ($selected.Count -eq 0) {
    Set-XlflowError -Result $result -Code "test_not_found" -Message ("test not found: " + $Filter)
    $result["workbook"] = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
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
  $workbook.Save()
  $result["workbook"] = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
  $result["tests"] = $results.ToArray()
  $result.logs = $logs.ToArray()
  if ($failed -gt 0) {
    Set-XlflowError -Result $result -Code "test_failed" -Message ("$failed of $($selected.Count) test(s) failed")
  }
} catch {
  Set-XlflowError -Result $result -Code "test_environment_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result["workbook"] = [ordered]@{ path = $WorkbookPath; session = (ConvertTo-XlflowBool $UseSession) }
} finally {
  if ($null -ne $runnerComponent) {
    try { $workbook.VBProject.VBComponents.Remove($runnerComponent) | Out-Null } catch { Write-Verbose ("failed to remove test harness module: " + $_.Exception.Message) }
  }
  if (ConvertTo-XlflowBool $UseSession) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
