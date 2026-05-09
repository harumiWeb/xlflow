param(
  [string]$WorkbookPath,
  [string]$Visible = "false",
  [string]$Sheet = "",
  [string]$RangeAddress = "",
  [string]$OutputPath = "",
  [string]$OutputIsDefault = "false",
  [string]$ImageFormat = "png",
  [string]$Overwrite = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "export-image"
$excel = $null
$workbook = $null
$worksheet = $null
$range = $null
$chartObjects = $null
$chartObject = $null
$chart = $null
$sessionAttached = $false
$sessionMode = "none"
$originalVisible = $false
$restoreVisible = $false
$saveState = [ordered]@{ dirty = $false; needs_save = $false }
$savedSheetName = ""
$savedSelectionAddress = ""
$createdParentDirs = $false
$exported = $false
$temporaryExportPath = ""

function Set-ExportImageValidationError {
  param([string]$Code, [string]$Message)
  Set-XlflowError -Result $result -Code $Code -Message $Message -Source "xlflow"
}

function Get-XlflowImageDimensions {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    return $null
  }

  $image = $null
  try {
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
    $image = [System.Drawing.Image]::FromFile($Path)
    return [ordered]@{
      width_px = [int]$image.Width
      height_px = [int]$image.Height
    }
  } catch {
    Write-Verbose ("failed to read exported image dimensions: " + $_.Exception.Message)
    return $null
  } finally {
    if ($null -ne $image) {
      try { $image.Dispose() } catch { Write-Verbose ("failed to dispose exported image handle: " + $_.Exception.Message) }
    }
  }
}

try {
  $normalizedFormat = [string]$ImageFormat
  if ([string]::IsNullOrWhiteSpace($normalizedFormat)) {
    $normalizedFormat = "png"
  }
  $normalizedFormat = $normalizedFormat.Trim().ToLowerInvariant()
  if ($normalizedFormat -ne "png") {
    Set-ExportImageValidationError -Code "unsupported_image_format" -Message ("Image format '" + $ImageFormat + "' is not supported. Supported formats: png.")
    Write-XlflowJson -Result $result
    exit
  }

  if ([string]::IsNullOrWhiteSpace($Sheet)) {
    Set-ExportImageValidationError -Code "export_image_args_invalid" -Message "-Sheet is required."
    Write-XlflowJson -Result $result
    exit
  }
  if ([string]::IsNullOrWhiteSpace($RangeAddress)) {
    Set-ExportImageValidationError -Code "export_image_args_invalid" -Message "-RangeAddress is required."
    Write-XlflowJson -Result $result
    exit
  }
  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    Set-ExportImageValidationError -Code "export_image_args_invalid" -Message "-OutputPath is required."
    Write-XlflowJson -Result $result
    exit
  }

  $resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
  if (Test-Path -LiteralPath $resolvedOutputPath -PathType Container) {
    Set-XlflowError -Result $result -Code "export_image_args_invalid" -Message ("Output path '" + $resolvedOutputPath + "' is a directory.") -Source "xlflow"
    Write-XlflowJson -Result $result
    exit
  }
  $extension = [System.IO.Path]::GetExtension($resolvedOutputPath)
  if (-not [string]::IsNullOrWhiteSpace($extension) -and $extension.ToLowerInvariant() -ne ".png") {
    Set-ExportImageValidationError -Code "unsupported_image_format" -Message ("Image format '" + $extension.TrimStart(".") + "' is not supported. Supported formats: png.")
    Write-XlflowJson -Result $result
    exit
  }

  $outputParent = Split-Path -Parent $resolvedOutputPath
  if (-not [string]::IsNullOrWhiteSpace($outputParent) -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
    $createdParentDirs = $true
  }

  if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not (ConvertTo-XlflowBool $Overwrite)) {
    Set-XlflowError -Result $result -Code "output_file_exists" -Message ("Output file '" + $resolvedOutputPath + "' already exists. Use --overwrite to replace it.") -Source "xlflow"
    Write-XlflowJson -Result $result
    exit
  }
  $exportPath = $resolvedOutputPath
  if ((Test-Path -LiteralPath $resolvedOutputPath -PathType Leaf) -and (ConvertTo-XlflowBool $Overwrite)) {
    $temporaryExportPath = Join-Path $outputParent ("xlflow-export-" + [Guid]::NewGuid().ToString("N") + ".png")
    $exportPath = $temporaryExportPath
  }

  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
  try {
    $originalVisible = [bool]$excel.Visible
  } catch {
    Write-Verbose ("failed to capture Excel visibility before export: " + $_.Exception.Message)
  }

  if ($sessionAttached) {
    $activeSheet = $null
    $selection = $null
    try {
      $activeSheet = $excel.ActiveSheet
      if ($null -ne $activeSheet) {
        $savedSheetName = [string]$activeSheet.Name
      }
    } catch {
      Write-Verbose ("failed to capture active sheet before export: " + $_.Exception.Message)
    } finally {
      Release-XlflowComObject -Object $activeSheet -Name "active sheet COM object"
    }
    try {
      $selection = $excel.Selection
      if ($null -ne $selection) {
        $savedSelectionAddress = [string]$selection.Address($false, $false)
      }
    } catch {
      Write-Verbose ("failed to capture active selection before export: " + $_.Exception.Message)
    } finally {
      Release-XlflowComObject -Object $selection -Name "selection COM object"
    }
  }

  $worksheet = Get-XlflowWorksheet -Workbook $workbook -Sheet $Sheet
  if ($null -eq $worksheet) {
    Set-XlflowError -Result $result -Code "sheet_not_found" -Message ("Sheet '" + $Sheet + "' was not found.") -Source "Excel"
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
    throw "sheet_not_found"
  }

  try {
    $range = $worksheet.Range($RangeAddress)
    $null = $range.Address($false, $false)
  } catch {
    Set-XlflowError -Result $result -Code "invalid_range" -Message ("Range '" + $RangeAddress + "' is invalid for sheet '" + $Sheet + "'.") -Source "Excel"
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
    throw "invalid_range"
  }

  try {
    $worksheet.Activate() | Out-Null
  } catch {
    Write-Verbose ("failed to activate worksheet before export: " + $_.Exception.Message)
  }
  if (-not $originalVisible) {
    try {
      $excel.Visible = $true
      $restoreVisible = $true
    } catch {
      Write-Verbose ("failed to temporarily show Excel before export: " + $_.Exception.Message)
    }
  }
  try {
    $range.Select() | Out-Null
  } catch {
    Write-Verbose ("failed to select export range before export: " + $_.Exception.Message)
  }
  Start-Sleep -Milliseconds 300

  $range.CopyPicture(2, -4147) | Out-Null
  Start-Sleep -Milliseconds 200
  $chartObjects = $worksheet.ChartObjects()
  $chartObject = $chartObjects.Add([double]$range.Left, [double]$range.Top, [Math]::Max([double]$range.Width, 1), [Math]::Max([double]$range.Height, 1))
  $chartObject.Name = "xlflow.export." + [Guid]::NewGuid().ToString("N")
  $chart = $chartObject.Chart
  $chart.Paste() | Out-Null
  Start-Sleep -Milliseconds 200
  if ($restoreVisible) {
    try {
      $excel.Visible = $originalVisible
      $restoreVisible = $false
    } catch {
      Write-Verbose ("failed to restore Excel visibility after capture: " + $_.Exception.Message)
    }
  }

  $exportOk = $chart.Export($exportPath, "PNG")
  if (-not $exportOk -and -not (Test-Path -LiteralPath $exportPath)) {
    throw "Excel did not create the requested image file."
  }
  if (-not [string]::IsNullOrWhiteSpace($temporaryExportPath)) {
    Move-Item -LiteralPath $temporaryExportPath -Destination $resolvedOutputPath -Force
    $temporaryExportPath = ""
  }

  $output = [ordered]@{
    path = $resolvedOutputPath
    format = $normalizedFormat
    default = (ConvertTo-XlflowBool $OutputIsDefault)
  }
  if ($createdParentDirs) {
    $output.created_parent_dirs = $true
  }
  $dimensions = Get-XlflowImageDimensions -Path $resolvedOutputPath
  if ($null -ne $dimensions) {
    $output.width_px = $dimensions.width_px
    $output.height_px = $dimensions.height_px
  }

  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = [ordered]@{
    kind = $(if ($sessionAttached) { "live_session" } else { "file" })
    path = $WorkbookPath
    sheet = $worksheet.Name
    range = [string]$range.Address($false, $false)
  }
  $result.output = $output
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "exported " + $worksheet.Name + "!" + [string]$range.Address($false, $false) + " to " + $resolvedOutputPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  $exported = $true
} catch {
  if ($null -eq $result.error) {
    Set-XlflowError -Result $result -Code "export_image_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  if ($null -eq $result.workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  }
} finally {
  if ($restoreVisible) {
    try {
      $excel.Visible = $originalVisible
    } catch {
      Write-Verbose ("failed to restore Excel visibility after export: " + $_.Exception.Message)
    }
  }
  if ($sessionAttached -and -not [string]::IsNullOrWhiteSpace($savedSheetName)) {
    $savedSheet = $null
    $savedSelectionRange = $null
    try {
      $savedSheet = Get-XlflowWorksheet -Workbook $workbook -Sheet $savedSheetName
      if ($null -ne $savedSheet) {
        $savedSheet.Activate() | Out-Null
        if (-not [string]::IsNullOrWhiteSpace($savedSelectionAddress)) {
          try {
            $savedSelectionRange = $savedSheet.Range($savedSelectionAddress)
            $savedSelectionRange.Select() | Out-Null
          } catch {
            Write-Verbose ("failed to restore selection after export: " + $_.Exception.Message)
          }
        }
      }
    } catch {
      Write-Verbose ("failed to restore active sheet after export: " + $_.Exception.Message)
    } finally {
      Release-XlflowComObject -Object $savedSelectionRange -Name "saved selection range COM object"
      Release-XlflowComObject -Object $savedSheet -Name "saved sheet COM object"
    }
  }

  if ($null -ne $chartObject) {
    try {
      $chartObject.Delete() | Out-Null
    } catch {
      if ($exported) {
        Add-XlflowWarning -Result $result -Code "temporary_object_cleanup_failed" -Message "The image was exported, but xlflow could not delete a temporary chart object."
      } else {
        Write-Verbose ("failed to remove temporary chart object: " + $_.Exception.Message)
      }
    }
  }

  Release-XlflowComObject -Object $chart -Name "chart COM object"
  Release-XlflowComObject -Object $chartObject -Name "temporary chart COM object"
  Release-XlflowComObject -Object $chartObjects -Name "chart objects collection COM object"
  Release-XlflowComObject -Object $range -Name "range COM object"
  Release-XlflowComObject -Object $worksheet -Name "worksheet COM object"
  if (-not [string]::IsNullOrWhiteSpace($temporaryExportPath) -and (Test-Path -LiteralPath $temporaryExportPath -PathType Leaf)) {
    try {
      Remove-Item -LiteralPath $temporaryExportPath -Force
    } catch {
      Write-Verbose ("failed to remove temporary export file: " + $_.Exception.Message)
    }
  }

  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
