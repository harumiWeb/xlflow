param(
  [string]$WorkbookPath,
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = "",
  [string]$Target = "",
  [string]$Sheet = "",
  [string]$Address = "",
  [string]$IncludeStyle = "false",
  [string]$MaxRows = "0",
  [string]$MaxCols = "0"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "inspect"
$excel = $null
$workbook = $null
$sessionAttached = $false
$sessionMode = "none"
$saveState = [ordered]@{ dirty = $false; needs_save = $false }

function Set-InspectValidationError {
  param([string]$Message)
  Set-XlflowError -Result $result -Code "inspect_args_invalid" -Message $Message -Source "xlflow"
}

function ConvertTo-XlflowCellTextValue {
  param($Cell)

  if ($null -eq $Cell) {
    return $null
  }
  try {
    $text = [string]$Cell.Text
    if ([string]::IsNullOrWhiteSpace($text)) {
      return $null
    }
    return $text
  } catch {
    return $null
  }
}

function ConvertTo-XlflowFormulaValue {
  param($Cell)

  if ($null -eq $Cell) {
    return $null
  }
  try {
    $formula = [string]$Cell.Formula
    if ([string]::IsNullOrWhiteSpace($formula)) {
      return $null
    }
    if (-not $formula.StartsWith("=")) {
      return $null
    }
    return $formula
  } catch {
    return $null
  }
}

function ConvertTo-XlflowVisibleBool {
  param($Value)

  try {
    return ([int]$Value -eq -1)
  } catch {
    return $false
  }
}

function ConvertTo-XlflowColorHex {
  param($ColorValue)

  if ($null -eq $ColorValue) {
    return $null
  }
  try {
    $color = [int64]$ColorValue
    if ($color -lt 0) {
      return $null
    }
    $red = $color -band 255
    $green = ($color -shr 8) -band 255
    $blue = ($color -shr 16) -band 255
    return ('#{0:X2}{1:X2}{2:X2}' -f $red, $green, $blue)
  } catch {
    return $null
  }
}

function ConvertTo-XlflowBorderStyleName {
  param(
    $LineStyle,
    $Weight
  )

  try {
    $line = [int]$LineStyle
    if ($line -eq -4142) {
      return "none"
    }
    switch ($line) {
      -4119 { return "double" }
      -4115 { return "dashed" }
      -4118 { return "dotted" }
      4 { return "dashDot" }
      5 { return "dashDotDot" }
      13 { return "slantDashDot" }
    }
    switch ([int]$Weight) {
      1 { return "hair" }
      2 { return "thin" }
      -4138 { return "medium" }
      4 { return "thick" }
      default { return "thin" }
    }
  } catch {
    return "none"
  }
}

function ConvertTo-XlflowAlignmentName {
  param(
    $Value,
    [string]$Axis
  )

  try {
    switch ([int]$Value) {
      -4131 { return "left" }
      -4108 { return "center" }
      -4152 { return "right" }
      -4130 { return "justify" }
      -4117 { return "distributed" }
      -4160 { return "top" }
      -4107 { return "bottom" }
      -4108 { return "center" }
      default {
        if ($Axis -eq "horizontal" -and [int]$Value -eq -4108) { return "center" }
        if ($Axis -eq "vertical" -and [int]$Value -eq -4108) { return "center" }
        return $null
      }
    }
  } catch {
    return $null
  }
}

function ConvertTo-XlflowFillType {
  param($Pattern)

  try {
    switch ([int]$Pattern) {
      -4142 { return "none" }
      1 { return "solid" }
      default { return ("pattern:" + [int]$Pattern) }
    }
  } catch {
    return "none"
  }
}

function Get-XlflowRangeAddress {
  param($Range)

  if ($null -eq $Range) {
    return ""
  }
  try {
    return [string]$Range.Address($false, $false, 1, $false)
  } catch {
    return ""
  }
}

function New-XlflowLiveInspectTargetInfo {
  param([string]$Path)

  return [ordered]@{
    kind = "live_session"
    path = $Path
    note = "This command inspected the live workbook currently open in Excel through xlflow session."
  }
}

function Get-XlflowWorksheetByName {
  param(
    $Workbook,
    [string]$SheetName
  )

  foreach ($worksheet in @($Workbook.Worksheets)) {
    try {
      if ([string]$worksheet.Name -eq $SheetName) {
        return $worksheet
      }
    } catch {
      Write-Verbose ("failed to inspect worksheet while resolving " + $SheetName + ": " + $_.Exception.Message)
    }
  }
  throw "sheet '$SheetName' not found"
}

function Get-XlflowUsedRangeInfo {
  param($Worksheet)

  $usedRange = $null
  try {
    $usedRange = $Worksheet.UsedRange
    if ($null -eq $usedRange) {
      return [ordered]@{
        range = $null
        address = ""
        row_count = 0
        column_count = 0
      }
    }
    $address = Get-XlflowRangeAddress -Range $usedRange
    $rowCount = [int]$usedRange.Rows.Count
    $columnCount = [int]$usedRange.Columns.Count
    if ($rowCount -eq 1 -and $columnCount -eq 1) {
      $probe = $null
      try {
        $probe = $usedRange.Cells.Item(1, 1)
        if ($null -eq (ConvertTo-XlflowCellTextValue -Cell $probe) -and $null -eq (ConvertTo-XlflowFormulaValue -Cell $probe)) {
          return [ordered]@{
            range = $null
            address = ""
            row_count = 0
            column_count = 0
          }
        }
      } finally {
        Release-XlflowComObject -Object $probe -Name "used-range probe cell COM object"
      }
    }
    return [ordered]@{
      range = $usedRange
      address = $address
      row_count = $rowCount
      column_count = $columnCount
    }
  } catch {
    throw "failed to inspect used range: $($_.Exception.Message)"
  }
}

function New-XlflowSheetSummary {
  param($Worksheet)

  $usedInfo = Get-XlflowUsedRangeInfo -Worksheet $Worksheet
  $summary = [ordered]@{
    name = [string]$Worksheet.Name
    index = [int]$Worksheet.Index
    visible = (ConvertTo-XlflowVisibleBool $Worksheet.Visible)
    used_range = [string]$usedInfo.address
    row_count = [int]$usedInfo.row_count
    column_count = [int]$usedInfo.column_count
  }
  if ($null -ne $usedInfo.range) {
    Release-XlflowComObject -Object $usedInfo.range -Name "worksheet used range COM object"
  }
  return $summary
}

function New-XlflowBorderEdgeSnapshot {
  param($Border)

  $color = $null
  try {
    if ([int]$Border.LineStyle -ne -4142) {
      $color = ConvertTo-XlflowColorHex -ColorValue $Border.Color
    }
  } catch {
    $color = $null
  }
  return [ordered]@{
    style = (ConvertTo-XlflowBorderStyleName -LineStyle $Border.LineStyle -Weight $Border.Weight)
    color = $color
  }
}

function New-XlflowStyleSnapshot {
  param($Cell)

  $fillType = $null
  $fillColor = $null
  $font = $null
  $numberFormat = $null
  $horizontal = $null
  $vertical = $null
  try {
    $fillType = ConvertTo-XlflowFillType -Pattern $Cell.Interior.Pattern
    if ($fillType -ne "none") {
      $fillColor = ConvertTo-XlflowColorHex -ColorValue $Cell.Interior.Color
    }
  } catch {
    $fillType = "none"
    $fillColor = $null
  }
  try {
    $font = [ordered]@{
      name = [string]$Cell.Font.Name
      size = [double]$Cell.Font.Size
      bold = [bool]$Cell.Font.Bold
      italic = [bool]$Cell.Font.Italic
      color = (ConvertTo-XlflowColorHex -ColorValue $Cell.Font.Color)
    }
  } catch {
    $font = $null
  }
  try {
    $numberFormat = [string]$Cell.NumberFormat
    if ([string]::IsNullOrWhiteSpace($numberFormat)) {
      $numberFormat = $null
    }
  } catch {
    $numberFormat = $null
  }
  try {
    $horizontal = ConvertTo-XlflowAlignmentName -Value $Cell.HorizontalAlignment -Axis "horizontal"
  } catch {
    $horizontal = $null
  }
  try {
    $vertical = ConvertTo-XlflowAlignmentName -Value $Cell.VerticalAlignment -Axis "vertical"
  } catch {
    $vertical = $null
  }
  return [ordered]@{
    fill = $(if ($fillType -eq "none" -and $null -eq $fillColor) { $null } else { [ordered]@{ type = $fillType; color = $fillColor } })
    font = $font
    border = [ordered]@{
      top = (New-XlflowBorderEdgeSnapshot -Border $Cell.Borders.Item(8))
      right = (New-XlflowBorderEdgeSnapshot -Border $Cell.Borders.Item(10))
      bottom = (New-XlflowBorderEdgeSnapshot -Border $Cell.Borders.Item(9))
      left = (New-XlflowBorderEdgeSnapshot -Border $Cell.Borders.Item(7))
    }
    number_format = $numberFormat
    horizontal_alignment = $horizontal
    vertical_alignment = $vertical
  }
}

function New-XlflowRangeSnapshot {
  param(
    $Worksheet,
    $Range,
    [string]$RangeAddress,
    [string]$UsedRangeAddress,
    [int]$MaxRowsValue,
    [int]$MaxColsValue,
    [bool]$WithStyle
  )

  $rowCount = [int]$Range.Rows.Count
  $columnCount = [int]$Range.Columns.Count
  $returnRows = $rowCount
  $returnCols = $columnCount
  $truncated = $false
  if ($MaxRowsValue -gt 0 -and $rowCount -gt $MaxRowsValue) {
    $returnRows = $MaxRowsValue
    $truncated = $true
  }
  if ($MaxColsValue -gt 0 -and $columnCount -gt $MaxColsValue) {
    $returnCols = $MaxColsValue
    $truncated = $true
  }

  $values = New-Object System.Collections.Generic.List[object]
  $cells = New-Object System.Collections.Generic.List[object]
  $rows = New-Object System.Collections.Generic.List[object]
  $columns = New-Object System.Collections.Generic.List[object]
  $warnings = New-Object System.Collections.Generic.List[string]
  $mergedRanges = New-Object System.Collections.Generic.List[string]
  $mergedLookup = @{}
  $mergedSeen = @{}

  $startRow = [int]$Range.Row
  $startCol = [int]$Range.Column

  for ($rowOffset = 0; $rowOffset -lt $returnRows; $rowOffset++) {
    $line = New-Object System.Collections.Generic.List[object]
    for ($colOffset = 0; $colOffset -lt $returnCols; $colOffset++) {
      $cell = $null
      try {
        $cell = $Worksheet.Cells.Item($startRow + $rowOffset, $startCol + $colOffset)
        $line.Add((ConvertTo-XlflowCellTextValue -Cell $cell))
        if ($WithStyle) {
          $cellAddress = Get-XlflowRangeAddress -Range $cell
          $mergeAddress = $null
          try {
            if ([bool]$cell.MergeCells) {
              $mergeAddress = Get-XlflowRangeAddress -Range $cell.MergeArea
              if (-not [string]::IsNullOrWhiteSpace($mergeAddress)) {
                $mergedLookup[$cellAddress] = $mergeAddress
                if (-not $mergedSeen.ContainsKey($mergeAddress)) {
                  $mergedSeen[$mergeAddress] = $true
                  $mergedRanges.Add($mergeAddress)
                }
              }
            }
          } catch {
            $mergeAddress = $null
          }
          $style = New-XlflowStyleSnapshot -Cell $cell
          $cells.Add([ordered]@{
            address = $cellAddress
            row = [int]($startRow + $rowOffset)
            column = [int]($startCol + $colOffset)
            value = (ConvertTo-XlflowCellTextValue -Cell $cell)
            formula = (ConvertTo-XlflowFormulaValue -Cell $cell)
            fill = $style.fill
            font = $style.font
            border = $style.border
            number_format = $style.number_format
            horizontal_alignment = $style.horizontal_alignment
            vertical_alignment = $style.vertical_alignment
            merged = ($null -ne $mergeAddress)
            merge_range = $mergeAddress
          })
        }
      } finally {
        Release-XlflowComObject -Object $cell -Name "range cell COM object"
      }
    }
    $values.Add(@($line.ToArray()))
  }

  if ($WithStyle) {
    for ($rowOffset = 0; $rowOffset -lt $returnRows; $rowOffset++) {
      $rowRef = $null
      try {
        $rowIndex = $startRow + $rowOffset
        $rowRef = $Worksheet.Rows.Item($rowIndex)
        $rows.Add([ordered]@{
          row = [int]$rowIndex
          height = [double]$rowRef.RowHeight
          hidden = [bool]$rowRef.Hidden
        })
      } finally {
        Release-XlflowComObject -Object $rowRef -Name "worksheet row COM object"
      }
    }
    for ($colOffset = 0; $colOffset -lt $returnCols; $colOffset++) {
      $columnRef = $null
      $columnHeaderCell = $null
      try {
        $colIndex = $startCol + $colOffset
        $columnRef = $Worksheet.Columns.Item($colIndex)
        $columnHeaderCell = $Worksheet.Cells.Item(1, $colIndex)
        $columnName = Get-XlflowRangeAddress -Range $columnHeaderCell
        $columnName = $columnName -replace '\d+$', ''
        $columns.Add([ordered]@{
          column = $columnName
          index = [int]$colIndex
          width = [double]$columnRef.ColumnWidth
          hidden = [bool]$columnRef.Hidden
        })
      } finally {
        Release-XlflowComObject -Object $columnHeaderCell -Name "worksheet column header cell COM object"
        Release-XlflowComObject -Object $columnRef -Name "worksheet column COM object"
      }
    }
  }

  $returnedRange = ""
  if ($returnRows -gt 0 -and $returnCols -gt 0) {
    $topLeft = $null
    $bottomRight = $null
    try {
      $topLeft = $Worksheet.Cells.Item($startRow, $startCol)
      $bottomRight = $Worksheet.Cells.Item($startRow + $returnRows - 1, $startCol + $returnCols - 1)
      $returnedRange = (Get-XlflowRangeAddress -Range $Worksheet.Range($topLeft, $bottomRight))
    } finally {
      Release-XlflowComObject -Object $topLeft -Name "returned range start cell COM object"
      Release-XlflowComObject -Object $bottomRight -Name "returned range end cell COM object"
    }
  }

  if ($truncated) {
    $warnings.Add(("Output was truncated: selection has " + $rowCount + " row(s) x " + $columnCount + " column(s), returned " + $returnRows + " row(s) x " + $returnCols + " column(s)."))
  }

  $snapshot = [ordered]@{
    sheet = [string]$Worksheet.Name
    row_count = [int]$rowCount
    column_count = [int]$columnCount
    values = @($values.ToArray())
    truncated = [bool]$truncated
    max_rows = [int]$MaxRowsValue
    max_cols = [int]$MaxColsValue
    returned_range = $returnedRange
  }
  if (-not [string]::IsNullOrWhiteSpace($RangeAddress)) {
    $snapshot.range = $RangeAddress
  }
  if (-not [string]::IsNullOrWhiteSpace($UsedRangeAddress)) {
    $snapshot.used_range = $UsedRangeAddress
  }
  if ($warnings.Count -gt 0) {
    $snapshot.warnings = @($warnings.ToArray())
  }
  if ($WithStyle) {
    $snapshot.style_included = $true
    $snapshot.cells = @($cells.ToArray())
    $snapshot.columns = @($columns.ToArray())
    $snapshot.rows = @($rows.ToArray())
    $snapshot.merged_ranges = @($mergedRanges.ToArray())
  }
  return $snapshot
}

try {
  $normalizedTarget = ([string]$Target).Trim().ToLowerInvariant()
  if ($normalizedTarget -notin @("workbook", "sheets", "range", "used-range", "cell")) {
    Set-InspectValidationError -Message ("unsupported inspect target '" + $Target + "'")
    Write-XlflowJson -Result $result
    exit
  }

  $maxRowsValue = 0
  $maxColsValue = 0
  try {
    $maxRowsValue = [int]$MaxRows
    $maxColsValue = [int]$MaxCols
  } catch {
    Set-InspectValidationError -Message "max row/column limits must be integers"
    Write-XlflowJson -Result $result
    exit
  }

  if (($normalizedTarget -eq "range" -or $normalizedTarget -eq "cell") -and ([string]::IsNullOrWhiteSpace($Sheet) -or [string]::IsNullOrWhiteSpace($Address))) {
    Set-InspectValidationError -Message "sheet and address are required for range/cell inspect"
    Write-XlflowJson -Result $result
    exit
  }
  if ($normalizedTarget -eq "used-range" -and [string]::IsNullOrWhiteSpace($Sheet)) {
    Set-InspectValidationError -Message "sheet is required for used-range inspect"
    Write-XlflowJson -Result $result
    exit
  }

  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath -AllowIsolatedOpen $false
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached

  $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save

  switch ($normalizedTarget) {
    "workbook" {
      $sheets = New-Object System.Collections.Generic.List[object]
      foreach ($worksheet in @($workbook.Worksheets)) {
        $sheets.Add((New-XlflowSheetSummary -Worksheet $worksheet))
      }
      $workbookSummary = [ordered]@{
        path = $WorkbookPath
        name = [System.IO.Path]::GetFileName($WorkbookPath)
        sheets = @($sheets.ToArray())
      }
      try {
        $workbookSummary.active_sheet = [string]$workbook.ActiveSheet.Name
      } catch {
        Write-Verbose ("failed to resolve active sheet: " + $_.Exception.Message)
      }
      $result.inspect = [ordered]@{
        target = "workbook"
        source = "excel_com"
        target_info = (New-XlflowLiveInspectTargetInfo -Path $WorkbookPath)
        workbook = $workbookSummary
      }
      $result.logs = @("inspected live workbook " + $WorkbookPath)
    }
    "sheets" {
      $sheets = New-Object System.Collections.Generic.List[object]
      foreach ($worksheet in @($workbook.Worksheets)) {
        $sheets.Add((New-XlflowSheetSummary -Worksheet $worksheet))
      }
      $result.inspect = [ordered]@{
        target = "sheets"
        source = "excel_com"
        target_info = (New-XlflowLiveInspectTargetInfo -Path $WorkbookPath)
        sheets = @($sheets.ToArray())
      }
      $result.logs = @("inspected live workbook worksheets")
    }
    "range" {
      $worksheet = $null
      $range = $null
      try {
        $worksheet = Get-XlflowWorksheetByName -Workbook $workbook -SheetName $Sheet
        $range = $worksheet.Range($Address)
        $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath -Sheet $Sheet -Range (Get-XlflowRangeAddress -Range $range)
        $result.inspect = [ordered]@{
          target = "range"
          source = "excel_com"
          target_info = (New-XlflowLiveInspectTargetInfo -Path $WorkbookPath)
          range = (New-XlflowRangeSnapshot -Worksheet $worksheet -Range $range -RangeAddress (Get-XlflowRangeAddress -Range $range) -UsedRangeAddress "" -MaxRowsValue $maxRowsValue -MaxColsValue $maxColsValue -WithStyle (ConvertTo-XlflowBool $IncludeStyle))
        }
        $result.logs = @("inspected live range " + $Sheet + "!" + (Get-XlflowRangeAddress -Range $range))
      } finally {
        Release-XlflowComObject -Object $range -Name "inspect range COM object"
        Release-XlflowComObject -Object $worksheet -Name "inspect range worksheet COM object"
      }
    }
    "used-range" {
      $worksheet = $null
      try {
        $worksheet = Get-XlflowWorksheetByName -Workbook $workbook -SheetName $Sheet
        $usedInfo = Get-XlflowUsedRangeInfo -Worksheet $worksheet
        $snapshot = $null
        if ($null -eq $usedInfo.range) {
          $snapshot = [ordered]@{
            sheet = [string]$worksheet.Name
            used_range = ""
            row_count = 0
            column_count = 0
            values = @()
          }
          if (ConvertTo-XlflowBool $IncludeStyle) {
            $snapshot.style_included = $true
            $snapshot.cells = @()
            $snapshot.columns = @()
            $snapshot.rows = @()
            $snapshot.merged_ranges = @()
          }
          $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath -Sheet $Sheet
        } else {
          $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath -Sheet $Sheet -Range $usedInfo.address
          $snapshot = New-XlflowRangeSnapshot -Worksheet $worksheet -Range $usedInfo.range -RangeAddress "" -UsedRangeAddress $usedInfo.address -MaxRowsValue $maxRowsValue -MaxColsValue $maxColsValue -WithStyle (ConvertTo-XlflowBool $IncludeStyle)
        }
        $result.inspect = [ordered]@{
          target = "used-range"
          source = "excel_com"
          target_info = (New-XlflowLiveInspectTargetInfo -Path $WorkbookPath)
          range = $snapshot
        }
        $result.logs = @("inspected live used range for " + $Sheet)
        Release-XlflowComObject -Object $usedInfo.range -Name "inspect used range COM object"
      } finally {
        Release-XlflowComObject -Object $worksheet -Name "inspect used-range worksheet COM object"
      }
    }
    "cell" {
      $worksheet = $null
      $cell = $null
      try {
        $worksheet = Get-XlflowWorksheetByName -Workbook $workbook -SheetName $Sheet
        $cell = $worksheet.Range($Address)
        $normalizedAddress = Get-XlflowRangeAddress -Range $cell
        $result.target = New-XlflowTargetResult -Kind "live_session" -Path $WorkbookPath -Sheet $Sheet -Range $normalizedAddress
        $result.inspect = [ordered]@{
          target = "cell"
          source = "excel_com"
          target_info = (New-XlflowLiveInspectTargetInfo -Path $WorkbookPath)
          cell = [ordered]@{
            sheet = [string]$worksheet.Name
            address = $normalizedAddress
            value = (ConvertTo-XlflowCellTextValue -Cell $cell)
          }
        }
        $result.logs = @("inspected live cell " + $Sheet + "!" + $normalizedAddress)
      } finally {
        Release-XlflowComObject -Object $cell -Name "inspect cell COM object"
        Release-XlflowComObject -Object $worksheet -Name "inspect cell worksheet COM object"
      }
    }
  }
} catch {
  $message = $_.Exception.Message
  if ($message -like "*xlflow session*") {
    Set-XlflowError -Result $result -Code "session_required" -Message $message -Source "xlflow"
  } elseif ($message -like "sheet '*' not found") {
    Set-XlflowError -Result $result -Code "sheet_not_found" -Message $message -Source "xlflow"
  } else {
    Set-XlflowError -Result $result -Code "inspect_failed" -Message $message -Source "powershell"
  }
} finally {
  Release-XlflowComReferences -Workbook $workbook -Excel $excel
}

Write-XlflowJson -Result $result
