param(
  [string]$Action,
  [string]$WorkbookPath,
  [string]$MetadataPath,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "session"
$excel = $null
$workbook = $null
$sessionStarted = $false

function Release-XlflowSessionComReferences {
  param($Workbook, $Excel)

  if ($null -ne $Workbook) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch { Write-Verbose ("failed to release workbook COM object: " + $_.Exception.Message) }
  }
  if ($null -ne $Excel) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch { Write-Verbose ("failed to release Excel COM object: " + $_.Exception.Message) }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Write-XlflowSessionMetadata {
  param($Excel, [string]$WorkbookPath, [string]$MetadataPath)

  $parent = Split-Path -Parent $MetadataPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }
  $pidValue = 0
  try {
    $hwnd = [int64]$Excel.Hwnd
    $proc = Get-Process | Where-Object { $_.MainWindowHandle -eq $hwnd } | Select-Object -First 1
    if ($null -ne $proc) {
      $pidValue = [int]$proc.Id
    }
  } catch {
    Write-Verbose ("failed to resolve Excel process id: " + $_.Exception.Message)
  }
  [ordered]@{
    pid = $pidValue
    workbook_path = [System.IO.Path]::GetFullPath($WorkbookPath)
    port = 0
    token = [guid]::NewGuid().ToString("N")
    started_at = (Get-Date).ToString("o")
  } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $MetadataPath -Encoding UTF8
}

function Read-XlflowSessionMetadata {
  param([string]$MetadataPath)
  if ([string]::IsNullOrWhiteSpace($MetadataPath) -or -not (Test-Path -LiteralPath $MetadataPath)) {
    return $null
  }
  return Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
}

try {
  switch ($Action) {
    "start" {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = ConvertTo-XlflowBool $Visible
      $excel.DisplayAlerts = $false
      $workbook = $excel.Workbooks.Open($WorkbookPath)
      Write-XlflowSessionMetadata -Excel $excel -WorkbookPath $WorkbookPath -MetadataPath $MetadataPath
      $sessionStarted = $true
      $result.workbook = [ordered]@{ path = $WorkbookPath; session = $true }
      $result.logs = @("started xlflow Excel session")
    }
    "status" {
      $metadata = Read-XlflowSessionMetadata -MetadataPath $MetadataPath
      $running = $false
      if ($null -ne $metadata -and $metadata.pid -gt 0) {
        $running = $null -ne (Get-Process -Id ([int]$metadata.pid) -ErrorAction SilentlyContinue)
      }
      $open = $false
      if ($running -or $null -ne $metadata) {
        try {
          $excel = Get-XlflowActiveExcel
          $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
          $open = $true
          $running = $true
        } catch {
          $open = $false
        }
      }
      $result.session = [ordered]@{ running = $running; workbook_open = $open; metadata = $metadata }
      $result.workbook = [ordered]@{ path = $WorkbookPath; session = $running }
      $result.logs = @($(if ($running -and $open) { "xlflow session is running" } else { "xlflow session is not running" }))
    }
    "save" {
      $excel = Get-XlflowActiveExcel
      $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
      $workbook.Save()
      $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true; session = $true }
      $result.logs = @("saved xlflow session workbook")
    }
    "stop" {
      $excel = Get-XlflowActiveExcel
      $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
      $workbook.Close($true) | Out-Null
      $excel.Quit() | Out-Null
      if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and (Test-Path -LiteralPath $MetadataPath)) {
        Remove-Item -LiteralPath $MetadataPath -Force -ErrorAction SilentlyContinue
      }
      $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true; session = $false }
      $result.logs = @("stopped xlflow Excel session")
    }
    default {
      Set-XlflowError -Result $result -Code "session_args_invalid" -Message "-Action must be start, status, stop, or save." -Source "xlflow"
    }
  }
} catch {
  if ($null -eq $result.error) {
    Set-XlflowError -Result $result -Code "session_failed" -Message $_.Exception.Message -Source $_.Exception.Source
  }
} finally {
  if ($Action -eq "start") {
    if (-not $sessionStarted) {
      Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
    }
  } elseif ($Action -eq "status" -or $Action -eq "save" -or $Action -eq "stop") {
    Release-XlflowSessionComReferences -Workbook $workbook -Excel $excel
  }
}

Write-XlflowJson -Result $result
