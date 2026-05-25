param(
  [string]$Action,
  [string]$TargetPid = "",
  [string]$Auto = "false",
  [string]$All = "false"
)

. "$PSScriptRoot/common.ps1"

function Get-XlflowExcelProcesses {
  $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
  $results = New-Object System.Collections.Generic.List[object]
  foreach ($proc in $processes) {
    $hasWorkbook = Get-XlflowWorkbookStateByProcessId -ProcessId ([int]$proc.Id)
    $results.Add([ordered]@{
      pid = [int]$proc.Id
      has_workbook = $hasWorkbook
    })
  }
  return @($results.ToArray())
}

function Stop-XlflowExcelProcess {
  param([int]$ProcessId)

  try {
    $proc = Get-Process -Id $ProcessId -ErrorAction Stop
  } catch {
    return @{ terminated = $true; method = "unknown" }
  }
  $method = ""
  $terminated = $false
  try {
    [void]$proc.CloseMainWindow()
    $stopped = $proc.WaitForExit(3000)
    if ($stopped) {
      return @{ terminated = $true; method = "graceful" }
    }
  } catch {
    Write-Verbose ("graceful close failed for PID ${ProcessId}: " + $_.Exception.Message)
  }

  $stillAlive = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
  if ($stillAlive) {
    try {
      Stop-Process -Id $ProcessId -Force -ErrorAction Stop
      $terminated = $true
      $method = "force"
    } catch {
      $method = "none"
      Write-Verbose ("force stop failed for PID ${ProcessId}: " + $_.Exception.Message)
    }
  } else {
    $terminated = $true
    $method = "unknown"
  }

  return @{ terminated = $terminated; method = $method }
}

$result = New-XlflowResult -Command "process"

try {
  if ($Action -eq "list") {
    $processes = Get-XlflowExcelProcesses
    $result.process = @($processes)
    if ($processes.Count -eq 0) {
      $result.logs = @("0 Excel processes found")
    } else {
      $result.logs = @("found $($processes.Count) Excel process(es)")
    }
    Write-XlflowJson -Result $result
    exit
  }

  if ($Action -eq "cleanup") {
    $isAuto = ConvertTo-XlflowBool $Auto
    $isAll = ConvertTo-XlflowBool $All
    $pidInt = 0
    $mode = ""
    $targets = New-Object System.Collections.Generic.List[int]

    if ([int]::TryParse($TargetPid, [ref]$pidInt) -and $pidInt -gt 0) {
      $mode = "pid"
      $proc = Get-Process -Id $pidInt -ErrorAction SilentlyContinue
      if (-not $proc -or $proc.ProcessName -ne "EXCEL") {
        Set-XlflowError -Result $result -Code "process_not_found" -Message "no Excel process found with PID $TargetPid" -Source "xlflow"
        Write-XlflowJson -Result $result
        exit
      }
      $targets.Add($pidInt) | Out-Null
    } elseif ($isAuto) {
      $mode = "auto"
      $processes = Get-XlflowExcelProcesses
      foreach ($proc in $processes) {
        if ($proc.has_workbook -eq $false) {
          $targets.Add([int]$proc.pid) | Out-Null
        }
      }
    } elseif ($isAll) {
      $mode = "all"
      $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
      foreach ($proc in $processes) {
        $targets.Add([int]$proc.Id) | Out-Null
      }
    } else {
      Set-XlflowError -Result $result -Code "process_args_invalid" -Message "process cleanup requires a PID, --auto, or --all" -Source "xlflow"
      Write-XlflowJson -Result $result
      exit
    }

    if ($targets.Count -eq 0) {
        $result.process = [ordered]@{
          action = "cleanup"
          mode = $mode
          total = 0
          results = @()
        }
        $result.logs = @("0 Excel processes to clean up")
        Write-XlflowJson -Result $result
        exit
    }

    $cleanupResults = New-Object System.Collections.Generic.List[object]
    try {
      foreach ($targetPid in $targets.ToArray()) {
        try {
          if ($isAll) {
            $stillExists = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
            if (-not $stillExists) {
              $cleanupResults.Add([ordered]@{
                pid = $targetPid
                terminated = $true
                method = "unknown"
              }) | Out-Null
              continue
            }
            Stop-Process -Id $targetPid -Force -ErrorAction Stop
            $cleanupResults.Add([ordered]@{
              pid = $targetPid
              terminated = $true
              method = "force"
            }) | Out-Null
          } else {
            $outcome = Stop-XlflowExcelProcess -ProcessId $targetPid
            $cleanupResults.Add([ordered]@{
              pid = $targetPid
              terminated = [bool]$outcome.terminated
              method = [string]$outcome.method
            }) | Out-Null
          }
        } catch {
          $cleanupResults.Add([ordered]@{
            pid = $targetPid
            terminated = $false
            method = "none"
          }) | Out-Null
        }
      }

      $terminatedCount = (@($cleanupResults.ToArray() | Where-Object { [bool]$_.terminated }).Count)
      $failedCount = $targets.Count - $terminatedCount

      $result.process = [ordered]@{
        action = "cleanup"
        mode = $mode
        total = $targets.Count
        results = @($cleanupResults.ToArray())
      }

      if ($failedCount -gt 0) {
        Set-XlflowError -Result $result -Code "process_termination_failed" -Message "$failedCount of $($targets.Count) Excel process(es) failed to terminate" -Source "xlflow"
      }

      if ($failedCount -eq 0) {
        $result.logs = @("terminated $terminatedCount Excel process(es)")
      }
      Write-XlflowJson -Result $result
      exit
    } catch {
      Set-XlflowError -Result $result -Code "process_cleanup_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
      Write-XlflowJson -Result $result
      exit
    }
  }

  Set-XlflowError -Result $result -Code "process_args_invalid" -Message "Action must be list or cleanup" -Source "xlflow"
  Write-XlflowJson -Result $result
  exit
} catch {
  Set-XlflowError -Result $result -Code "process_enumeration_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  Write-XlflowJson -Result $result
  exit
}
