param(
  [string]$Action,
  [string]$WorkbookPath,
  [string]$MetadataPath,
  [string]$UseSession = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "session"
$excel = $null
$workbook = $null
$sessionStarted = $false
$sessionMode = "managed"

function Write-XlflowSessionMetadata {
  param($Excel, [string]$WorkbookPath, [string]$MetadataPath)

  $parent = Split-Path -Parent $MetadataPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }
  $pidValue = 0
  $hwndValue = 0
  try { $hwndValue = [int64]$Excel.Hwnd } catch { Write-Verbose ("failed to read Excel hwnd: " + $_.Exception.Message) }
  $pidValue = Get-XlflowExcelProcessId -Excel $Excel
  [ordered]@{
    pid = $pidValue
    hwnd = $hwndValue
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
      if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and (Test-Path -LiteralPath $MetadataPath)) {
        Close-XlflowSessionWorkbook -WorkbookPath $WorkbookPath -MetadataPath $MetadataPath -Save $false
      }
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $true
      try { $excel.UserControl = $true } catch { Write-Verbose ("failed to set Excel UserControl: " + $_.Exception.Message) }
      $workbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts $false -DisableAutomationMacros $false
      $userFormNames = @()
      try {
        $userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)
      } catch {
        Write-Verbose ("failed to inspect UserForms during session start: " + $_.Exception.Message)
      }
      Write-XlflowSessionMetadata -Excel $excel -WorkbookPath $WorkbookPath -MetadataPath $MetadataPath
      $sessionStarted = $true
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $true -SessionMode $sessionMode
      $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath
      $result.session = New-XlflowSessionResult -Active $true -WorkbookPath $WorkbookPath -Dirty $false -SaveRequired $false -Mode $sessionMode
      Add-XlflowUserFormDiscoveryMessages -Result $result -Names $userFormNames
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
          $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
          $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
          $open = $true
          $running = $true
        } catch {
          $open = $false
        }
      }
      $dirty = Get-XlflowWorkbookDirtyState -Workbook $workbook
      $needsSave = $running -and $open -and (($null -eq $dirty) -or [bool]$dirty)
      $userFormNames = @()
      $userFormNamesKnown = $false
      if ($running -and $open) {
        try {
          $userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)
          $userFormNamesKnown = $true
        } catch {
          Write-Verbose ("failed to inspect UserForms during session status: " + $_.Exception.Message)
          Add-XlflowStateWarning -Result $result -Code "userform_detection_unavailable" -Message "xlflow could not determine whether the live workbook contains UserForms during session status. Treat disk-backed inspect and source review as potentially incomplete until the workbook is saved and reviewed explicitly."
        }
      }
      $result.session = [ordered]@{
        running = $running
        workbook_open = $open
        metadata = $metadata
        active = ($running -and $open)
        workbook_path = $WorkbookPath
        dirty = $dirty
        needs_save = $needsSave
        save_required = $needsSave
        live_newer_than_disk = $needsSave
        source_of_truth = $(if ($needsSave) { "live_workbook" } else { "saved_workbook" })
        mode = $(if ($running -and $open) { $sessionMode } else { "none" })
      }
      if ($running -and $open) {
        $result.session.userforms_known = $userFormNamesKnown
        if ($userFormNamesKnown) {
          $result.session.userforms_present = ($userFormNames.Count -gt 0)
          $result.session.userform_count = $userFormNames.Count
        }
      }
      if ($running -and $open) {
        $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $true -SessionMode $sessionMode -Dirty $dirty -NeedsSave $needsSave
        $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath
        Add-XlflowUserFormDiscoveryMessages -Result $result -Names $userFormNames
      } else {
        $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $false -SessionMode "none"
        $result.target = New-XlflowTargetResult -Kind "file" -Path $WorkbookPath
      }
      if ($needsSave) {
        Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes."
        Add-XlflowUserFormSessionStaleWarning -Result $result -Names $userFormNames
      }
      $result.logs = @($(if ($running -and $open) { "xlflow session is running" } else { "xlflow session is not running" }))
    }
    "save" {
      $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -UseSession $UseSession -MetadataPath $MetadataPath -AllowIsolatedOpen $false
      $excel = $openResult.excel
      $workbook = $openResult.workbook
      $sessionMode = [string]$openResult.session_mode
      $userFormNames = @()
      try {
        $userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)
      } catch {
        Write-Verbose ("failed to inspect UserForms during session save: " + $_.Exception.Message)
      }
      $workbook.Save()
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $true -SessionMode $sessionMode -Saved $true -Dirty $false -NeedsSave $false
      $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath
      $result.session = New-XlflowSessionResult -Active $true -WorkbookPath $WorkbookPath -Dirty $false -SaveRequired $false -Mode $sessionMode
      Add-XlflowUserFormDiscoveryMessages -Result $result -Names $userFormNames
      $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "saved xlflow session workbook") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    "stop" {
      $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
      $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
      $userFormNames = @()
      try {
        $userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)
      } catch {
        Write-Verbose ("failed to inspect UserForms during session stop: " + $_.Exception.Message)
      }
      $wasDirty = Get-XlflowWorkbookDirtyState -Workbook $workbook
      if ($null -eq $wasDirty) {
        $wasDirty = $false
      }
      try { $excel = $workbook.Application } catch { Write-Verbose ("failed to resolve workbook application: " + $_.Exception.Message) }
      $workbook.Close($true) | Out-Null
      $excel.Quit() | Out-Null
      if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and (Test-Path -LiteralPath $MetadataPath)) {
        Remove-Item -LiteralPath $MetadataPath -Force -ErrorAction SilentlyContinue
      }
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $false -SessionMode "none" -Saved $true -DirtyBeforeStop $wasDirty -AutoSavedOnStop $wasDirty
      $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath
      $result.session = New-XlflowSessionResult -Active $false -WorkbookPath $WorkbookPath -Dirty $false -SaveRequired $false -Mode "none"
      Add-XlflowUserFormDiscoveryMessages -Result $result -Names $userFormNames
      $result.logs = @(
        @(
          $(if ($wasDirty) { "warning: session workbook had unsaved changes before stop" } else { $null }),
          $(if ($wasDirty) { "auto-saved workbook while stopping xlflow session; prefer xlflow save before stop" } else { $null }),
          "stopped xlflow Excel session"
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      )
      if ($wasDirty) {
        Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live session workbook had unsaved changes and was saved while stopping the session."
      }
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
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  }
}

Write-XlflowJson -Result $result
