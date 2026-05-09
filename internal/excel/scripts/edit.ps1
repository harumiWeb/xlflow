param(
  [string]$Action,
  [string]$WorkbookPath,
  [string]$Visible = "false",
  [string]$Sheet = "",
  [string]$Cell = "",
  [string]$RangeAddress = "",
  [string]$Rows = "",
  [string]$Columns = "",
  [string]$Value = "",
  [string]$Formula = "",
  [string]$Fill = "",
  [string]$Clear = "",
  [string]$Height = "",
  [string]$Width = "",
  [string]$Events = "keep",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "edit"
$excel = $null
$workbook = $null
$worksheet = $null
$targetRange = $null
$sessionAttached = $false
$sessionMode = "none"
$saveState = [ordered]@{ dirty = $false; needs_save = $false }
$eventMode = "keep"
$eventStateCaptured = $false
$eventStateChanged = $false
$enableEventsBefore = $null
$enableEventsAfter = $null
$eventsRestored = $null
$saveHint = 'Run `xlflow save --session` to persist changes to disk.'

function Set-EditValidationError {
  param([string]$Code, [string]$Message)
  Set-XlflowError -Result $result -Code $Code -Message $Message -Source "xlflow"
}

function ConvertTo-XlflowNormalizedColor {
  param([string]$Value)

  $trimmed = ([string]$Value).Trim()
  if ($trimmed -notmatch '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$') {
    throw "Color '$Value' is not supported in MVP. Use #RGB or #RRGGBB."
  }
  $hex = $trimmed.Substring(1).ToUpperInvariant()
  if ($hex.Length -eq 3) {
    $hex = ([string]$hex[0] * 2) + ([string]$hex[1] * 2) + ([string]$hex[2] * 2)
  }
  return "#" + $hex
}

function ConvertTo-XlflowOleColor {
  param([string]$Value)

  $normalized = ConvertTo-XlflowNormalizedColor -Value $Value
  $hex = $normalized.Substring(1)
  $red = [Convert]::ToInt32($hex.Substring(0, 2), 16)
  $green = [Convert]::ToInt32($hex.Substring(2, 2), 16)
  $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16)
  return ($red + ($green * 256) + ($blue * 65536))
}

function ConvertFrom-XlflowOleColor {
  param($ColorValue)

  if ($null -eq $ColorValue) {
    return $null
  }
  try {
    $color = [int64]$ColorValue
    $red = $color -band 255
    $green = ($color -shr 8) -band 255
    $blue = ($color -shr 16) -band 255
    return ('#{0:X2}{1:X2}{2:X2}' -f $red, $green, $blue)
  } catch {
    return $null
  }
}

function Get-XlflowCellValueSnapshot {
  param($Range)

  try {
    $value = $Range.Value2
    if ($null -eq $value) {
      return ""
    }
    return $value
  } catch {
    return ""
  }
}

function Get-XlflowCellFormulaSnapshot {
  param($Range)

  try {
    $formulaValue = [string]$Range.Formula
    if ([string]::IsNullOrWhiteSpace($formulaValue)) {
      return $null
    }
    if (-not $formulaValue.StartsWith("=")) {
      return $null
    }
    return $formulaValue
  } catch {
    return $null
  }
}

function Get-XlflowRangeFillSummary {
  param($Range)

  $count = [int]$Range.Count
  if ($count -le 0) {
    return $null
  }
  $first = $null
  for ($i = 1; $i -le $count; $i++) {
    $cell = $null
    try {
      $cell = $Range.Item($i)
      $current = ConvertFrom-XlflowOleColor -ColorValue $cell.Interior.Color
      if ($null -eq $first) {
        $first = $current
        continue
      }
      if ($first -ne $current) {
        return "mixed"
      }
    } finally {
      Release-XlflowComObject -Object $cell -Name "range cell COM object"
    }
  }
  return $first
}

function Get-XlflowRangeCellCount {
  param($Range)

  try {
    return ([int]$Range.Rows.Count * [int]$Range.Columns.Count)
  } catch {
    return [int]$Range.Count
  }
}

function Get-XlflowRowHeightSummary {
  param($RowRange)

  $count = [int]$RowRange.Rows.Count
  if ($count -le 0) {
    return $null
  }
  $first = $null
  for ($i = 1; $i -le $count; $i++) {
    $row = $null
    try {
      $row = $RowRange.Rows.Item($i)
      $current = [double]$row.RowHeight
      if ($null -eq $first) {
        $first = $current
        continue
      }
      if ([Math]::Abs($first - $current) -gt 0.0001) {
        return "mixed"
      }
    } finally {
      Release-XlflowComObject -Object $row -Name "row COM object"
    }
  }
  return $first
}

function Get-XlflowColumnWidthSummary {
  param($ColumnRange)

  $count = [int]$ColumnRange.Columns.Count
  if ($count -le 0) {
    return $null
  }
  $first = $null
  for ($i = 1; $i -le $count; $i++) {
    $column = $null
    try {
      $column = $ColumnRange.Columns.Item($i)
      $current = [double]$column.ColumnWidth
      if ($null -eq $first) {
        $first = $current
        continue
      }
      if ([Math]::Abs($first - $current) -gt 0.0001) {
        return "mixed"
      }
    } finally {
      Release-XlflowComObject -Object $column -Name "column COM object"
    }
  }
  return $first
}

function Set-XlflowEditResultContext {
  param([string]$Kind, [string]$SelectorName, [string]$SelectorValue)

  $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $edit = [ordered]@{
    kind = $Kind
    sheet = $worksheet.Name
  }
  if (-not [string]::IsNullOrWhiteSpace($SelectorName) -and -not [string]::IsNullOrWhiteSpace($SelectorValue)) {
    $edit[$SelectorName] = $SelectorValue
  }
  $result.edit = $edit
}

function Update-XlflowEditResultSaveState {
  param($Workbook)

  $script:saveState = Get-XlflowWorkbookSaveState -Workbook $Workbook -SessionAttached $sessionAttached
  if ($null -ne $result.session) {
    $result.session.dirty = $saveState.dirty
    $result.session.save_required = $saveState.needs_save
  }
  if ($null -ne $result.workbook) {
    $result.workbook.dirty = $saveState.dirty
    $result.workbook.needs_save = $saveState.needs_save
  }
}

try {
  if ([string]::IsNullOrWhiteSpace($Action)) {
    Set-EditValidationError -Code "edit_args_invalid" -Message "-Action is required."
    Write-XlflowJson -Result $result
    exit
  }
  if ($Action -ne "cell" -and $Action -ne "range" -and $Action -ne "rows" -and $Action -ne "columns") {
    Set-EditValidationError -Code "edit_args_invalid" -Message ("Unsupported edit action: " + $Action)
    Write-XlflowJson -Result $result
    exit
  }
  if (-not (ConvertTo-XlflowBool $UseSession)) {
    Set-EditValidationError -Code "session_required" -Message "`xlflow edit` requires an active session. Run `xlflow session start` first."
    Write-XlflowJson -Result $result
    exit
  }

  try {
    $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "false" -UseSession $UseSession -MetadataPath $MetadataPath -AllowIsolatedOpen $false
  } catch {
    Set-EditValidationError -Code "session_required" -Message "`xlflow edit` requires an active session. Run `xlflow session start` first."
    Write-XlflowJson -Result $result
    exit
  }

  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached

  $worksheet = Get-XlflowWorksheet -Workbook $workbook -Sheet $Sheet
  if ($null -eq $worksheet) {
    Set-XlflowError -Result $result -Code "sheet_not_found" -Message ("Sheet '" + $Sheet + "' was not found.") -Source "Excel"
    throw "sheet_not_found"
  }

  switch ($Action) {
    "cell" {
      try {
        $targetRange = $worksheet.Range($Cell)
        $cellAddress = [string]$targetRange.Address($false, $false)
      } catch {
        Set-XlflowError -Result $result -Code "invalid_cell_address" -Message ("Cell address '" + $Cell + "' is invalid.") -Source "Excel"
        throw "invalid_cell_address"
      }

      Set-XlflowEditResultContext -Kind "cell" -SelectorName "cell" -SelectorValue $cellAddress
      $valueRequested = $PSBoundParameters.ContainsKey("Value")
      $formulaRequested = $PSBoundParameters.ContainsKey("Formula")
      $fillRequested = -not [string]::IsNullOrWhiteSpace($Fill)

      if ($fillRequested) {
        try {
          $normalizedFill = ConvertTo-XlflowNormalizedColor -Value $Fill
        } catch {
          Set-XlflowError -Result $result -Code "invalid_color" -Message $_.Exception.Message -Source "xlflow"
          throw "invalid_color"
        }
        $beforeFill = Get-XlflowRangeFillSummary -Range $targetRange
        $targetRange.Interior.Pattern = 1
        $targetRange.Interior.Color = ConvertTo-XlflowOleColor -Value $normalizedFill
        $result.edit.mutation = [ordered]@{
          style = [ordered]@{
            fill = [ordered]@{
              before = $beforeFill
              after = $normalizedFill
            }
            changed = ($beforeFill -ne $normalizedFill)
          }
        }
        Update-XlflowEditResultSaveState -Workbook $workbook
        $result.logs = @($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), ("edited {0}!{1} fill in the live Excel session" -f $worksheet.Name, $cellAddress), $saveHint) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        break
      }

      $eventMode = ([string]$Events).Trim().ToLowerInvariant()
      if ([string]::IsNullOrWhiteSpace($eventMode)) {
        $eventMode = "keep"
      } elseif ($eventMode -ne "keep" -and $eventMode -ne "on" -and $eventMode -ne "off") {
        Set-EditValidationError -Code "edit_args_invalid" -Message "-Events must be keep, on, or off."
        throw "edit_args_invalid"
      }
      try {
        $enableEventsBefore = [bool]$excel.EnableEvents
        $eventStateCaptured = $true
        if ($eventMode -eq "on") {
          $excel.EnableEvents = $true
          $eventStateChanged = $true
        } elseif ($eventMode -eq "off") {
          $excel.EnableEvents = $false
          $eventStateChanged = $true
        }
      } catch {
        Write-Verbose ("failed to inspect or set Application.EnableEvents: " + $_.Exception.Message)
      }

      try {
        $beforeValue = Get-XlflowCellValueSnapshot -Range $targetRange
        $beforeFormula = Get-XlflowCellFormulaSnapshot -Range $targetRange
        if ($valueRequested) {
          $targetRange.Value2 = $Value
        } elseif ($formulaRequested) {
          $targetRange.Formula = $Formula
        } else {
          Set-EditValidationError -Code "edit_args_invalid" -Message "One of -Value, -Formula, or -Fill is required."
          throw "edit_args_invalid"
        }
      } catch {
        $result.edit.events = [ordered]@{
          mode = $eventMode
          enable_events_before = $enableEventsBefore
        }
        if ($null -eq $result.error) {
          Set-XlflowError -Result $result -Code "edit_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult -Phase "edit_cell"
          if ($eventMode -eq "on") {
            Add-XlflowHint -Result $result -Code "possible_event_handler_failure" -Message "This edit ran with events enabled. If a Worksheet_Change handler also ran, inspect that VBA path as well."
          }
        }
        throw
      } finally {
        if ($eventStateCaptured) {
          try {
            $enableEventsAfter = [bool]$excel.EnableEvents
          } catch {
            Write-Verbose ("failed to read Application.EnableEvents after edit: " + $_.Exception.Message)
          }
          if ($eventStateChanged) {
            try {
              $excel.EnableEvents = $enableEventsBefore
              $eventsRestored = $true
            } catch {
              $eventsRestored = $false
              Add-XlflowWarning -Result $result -Code "events_restore_failed" -Message "xlflow could not restore the previous Application.EnableEvents state."
            }
          } else {
            $eventsRestored = $true
          }
        }
      }

      if ($eventStateCaptured) {
        try {
          $enableEventsAfter = [bool]$excel.EnableEvents
        } catch {
          Write-Verbose ("failed to read restored Application.EnableEvents state: " + $_.Exception.Message)
        }
      }

      $afterValue = Get-XlflowCellValueSnapshot -Range $targetRange
      $afterFormula = Get-XlflowCellFormulaSnapshot -Range $targetRange
      $result.edit.events = [ordered]@{
        mode = $eventMode
        enable_events_before = $enableEventsBefore
        enable_events_after = $enableEventsAfter
        restored = $eventsRestored
      }
      if ($valueRequested) {
        $result.edit.mutation = [ordered]@{
          value = [ordered]@{
            before = $beforeValue
            after = $afterValue
          }
        }
      } else {
        $result.edit.mutation = [ordered]@{
          formula = [ordered]@{
            before = $beforeFormula
            after = $afterFormula
          }
        }
      }
      Update-XlflowEditResultSaveState -Workbook $workbook
      $mutationLabel = $(if ($valueRequested) { "value" } else { "formula" })
      $result.logs = @($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), ("edited {0}!{1} {2} in the live Excel session" -f $worksheet.Name, $cellAddress, $mutationLabel), $saveHint) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
    "range" {
      try {
        $targetRange = $worksheet.Range($RangeAddress)
        $normalizedRange = [string]$targetRange.Address($false, $false)
      } catch {
        Set-XlflowError -Result $result -Code "invalid_range" -Message ("Range '" + $RangeAddress + "' is invalid for sheet '" + $Sheet + "'.") -Source "Excel"
        throw "invalid_range"
      }

      Set-XlflowEditResultContext -Kind "range" -SelectorName "range" -SelectorValue $normalizedRange
      $affectedCells = Get-XlflowRangeCellCount -Range $targetRange

      if (-not [string]::IsNullOrWhiteSpace($Fill)) {
        try {
          $normalizedFill = ConvertTo-XlflowNormalizedColor -Value $Fill
        } catch {
          Set-XlflowError -Result $result -Code "invalid_color" -Message $_.Exception.Message -Source "xlflow"
          throw "invalid_color"
        }
        $beforeFill = Get-XlflowRangeFillSummary -Range $targetRange
        $targetRange.Interior.Pattern = 1
        $targetRange.Interior.Color = ConvertTo-XlflowOleColor -Value $normalizedFill
        $result.edit.mutation = [ordered]@{
          style = [ordered]@{
            fill = [ordered]@{
              before = $beforeFill
              after = $normalizedFill
            }
          }
          affected_cells = $affectedCells
        }
        Update-XlflowEditResultSaveState -Workbook $workbook
        $result.logs = @($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), ("edited {0}!{1} fill in the live Excel session" -f $worksheet.Name, $normalizedRange), $saveHint) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      } else {
        $clearMode = ([string]$Clear).Trim().ToLowerInvariant()
        switch ($clearMode) {
          "contents" { $targetRange.ClearContents() | Out-Null }
          "formats" { $targetRange.ClearFormats() | Out-Null }
          "all" { $targetRange.Clear() | Out-Null }
          default {
            Set-EditValidationError -Code "edit_args_invalid" -Message "-Clear must be contents, formats, or all."
            throw "edit_args_invalid"
          }
        }
        $result.edit.mutation = [ordered]@{
          clear = [ordered]@{
            mode = $clearMode
          }
          affected_cells = $affectedCells
        }
        Update-XlflowEditResultSaveState -Workbook $workbook
        $result.logs = @($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), ("cleared {0} on {1}!{2} in the live Excel session" -f $clearMode, $worksheet.Name, $normalizedRange), $saveHint) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      }
    }
    "rows" {
      try {
        $targetRange = $worksheet.Rows($Rows)
        $normalizedRows = [string]$Rows
      } catch {
        Set-XlflowError -Result $result -Code "invalid_row_selector" -Message ("Row selector '" + $Rows + "' is invalid.") -Source "Excel"
        throw "invalid_row_selector"
      }

      Set-XlflowEditResultContext -Kind "rows" -SelectorName "rows" -SelectorValue $normalizedRows
      $beforeHeight = Get-XlflowRowHeightSummary -RowRange $targetRange
      $targetRange.RowHeight = [double]$Height
      $afterHeight = Get-XlflowRowHeightSummary -RowRange $targetRange
      $result.edit.mutation = [ordered]@{
        row_height = [ordered]@{
          before = $beforeHeight
          after = $afterHeight
        }
        affected_rows = [int]$targetRange.Rows.Count
      }
      Update-XlflowEditResultSaveState -Workbook $workbook
      $result.logs = @($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), ("edited row height for {0}!{1} in the live Excel session" -f $worksheet.Name, $normalizedRows), $saveHint) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
    "columns" {
      try {
        $targetRange = $worksheet.Columns($Columns)
        $normalizedColumns = [string]$Columns
      } catch {
        Set-XlflowError -Result $result -Code "invalid_column_selector" -Message ("Column selector '" + $Columns + "' is invalid.") -Source "Excel"
        throw "invalid_column_selector"
      }

      Set-XlflowEditResultContext -Kind "columns" -SelectorName "columns" -SelectorValue $normalizedColumns
      $beforeWidth = Get-XlflowColumnWidthSummary -ColumnRange $targetRange
      $targetRange.ColumnWidth = [double]$Width
      $afterWidth = Get-XlflowColumnWidthSummary -ColumnRange $targetRange
      $result.edit.mutation = [ordered]@{
        column_width = [ordered]@{
          before = $beforeWidth
          after = $afterWidth
        }
        affected_columns = [int]$targetRange.Columns.Count
      }
      Update-XlflowEditResultSaveState -Workbook $workbook
      $result.logs = @($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), ("edited column width for {0}!{1} in the live Excel session" -f $worksheet.Name, $normalizedColumns), $saveHint) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
  }
} catch {
  if ($null -eq $result.error) {
    Set-XlflowError -Result $result -Code "edit_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  if ($null -eq $result.target -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath
  }
  if ($null -eq $result.workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  }
  if ($null -eq $result.session -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  }
} finally {
  Release-XlflowComObject -Object $targetRange -Name "edit target range COM object"
  Release-XlflowComObject -Object $worksheet -Name "worksheet COM object"
  Release-XlflowComReferences -Workbook $workbook -Excel $excel
}

Write-XlflowJson -Result $result
