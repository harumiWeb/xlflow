param(
  [string]$Action,
  [string]$WorkbookPath,
  [string]$Visible = "false",
  [string]$Sheet = "",
  [string]$Cell = "",
  [string]$Text = "",
  [string]$Macro = "",
  [string]$Id = "",
  [int]$Width = 160,
  [int]$Height = 40,
  [string]$CreateSheet = "false",
  [string]$VerifyMacro = "false",
  [string]$MetadataPath = "",
  [string]$UseSession = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "ui"
$excel = $null
$workbook = $null
$saveWorkbook = $false
$sessionAttached = $false
$sessionMode = "none"
$saveState = [ordered]@{ dirty = $false; needs_save = $false }

function Set-UIButtonValidationError {
  param([string]$Message)
  Set-XlflowError -Result $result -Code "ui_button_args_invalid" -Message $Message -Source "xlflow"
}

try {
  if ([string]::IsNullOrWhiteSpace($Action)) {
    Set-UIButtonValidationError -Message "-Action is required."
    Write-XlflowJson -Result $result
    exit
  }
  if ($Action -ne "add" -and $Action -ne "list" -and $Action -ne "remove") {
    Set-UIButtonValidationError -Message ("Unsupported action: " + $Action)
    Write-XlflowJson -Result $result
    exit
  }

  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -MetadataPath $MetadataPath -UseSession $UseSession
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save

  switch ($Action) {
    "add" {
      if ([string]::IsNullOrWhiteSpace($Sheet)) { Set-UIButtonValidationError -Message "-Sheet is required."; break }
      if ([string]::IsNullOrWhiteSpace($Cell)) { Set-UIButtonValidationError -Message "-Cell is required."; break }
      if ([string]::IsNullOrWhiteSpace($Text)) { Set-UIButtonValidationError -Message "-Text is required."; break }
      if ([string]::IsNullOrWhiteSpace($Macro)) { Set-UIButtonValidationError -Message "-Macro is required."; break }
      if ($Width -le 0) { Set-UIButtonValidationError -Message "-Width must be greater than 0."; break }
      if ($Height -le 0) { Set-UIButtonValidationError -Message "-Height must be greater than 0."; break }

      $buttonId = ConvertTo-XlflowUIButtonId -Value $Id
      if ([string]::IsNullOrWhiteSpace($buttonId)) {
        $buttonId = ConvertTo-XlflowUIButtonId -Value $Macro
      }
      if ([string]::IsNullOrWhiteSpace($buttonId)) {
        Set-UIButtonValidationError -Message "-Id could not be derived from -Macro."
        break
      }

      $worksheet = Get-XlflowWorksheet -Workbook $workbook -Sheet $Sheet
      if ($null -eq $worksheet) {
        if (ConvertTo-XlflowBool $CreateSheet) {
          $worksheet = $workbook.Worksheets.Add()
          $worksheet.Name = $Sheet
        } else {
          Set-XlflowError -Result $result -Code "sheet_not_found" -Message ("Worksheet not found: " + $Sheet) -Source "Excel"
          break
        }
      }

      if (ConvertTo-XlflowBool $VerifyMacro) {
        try {
          if (-not (Test-XlflowMacroExists -Workbook $workbook -MacroName $Macro)) {
            Set-XlflowError -Result $result -Code "macro_not_found" -Message ("Macro not found: " + $Macro) -Source "Excel" -Phase "verify_macro"
            break
          }
        } catch {
          Set-XlflowError -Result $result -Code "vbide_access_denied" -Message "VBIDE access is not available." -Source "Excel" -Phase "verify_macro"
          break
        }
      }

      try {
        $range = $worksheet.Range($Cell)
      } catch {
        Set-XlflowError -Result $result -Code "ui_button_args_invalid" -Message ("Invalid cell address: " + $Cell) -Source "Excel"
        break
      }

      $buttonName = ConvertTo-XlflowUIButtonName -Id $buttonId
      $button = Get-XlflowUIButton -Worksheet $worksheet -Name $buttonName
      $updated = $true
      if ($null -eq $button) {
        $button = $worksheet.Buttons().Add($range.Left, $range.Top, $Width, $Height)
        $button.Name = $buttonName
        $updated = $false
      }

      $button.Caption = $Text
      $button.OnAction = $Macro
      $button.Left = $range.Left
      $button.Top = $range.Top
      $button.Width = $Width
      $button.Height = $Height
      $saveWorkbook = $true
      $result.ui = [ordered]@{
        button = (ConvertTo-XlflowUIButtonInfo -Button $button -Sheet $worksheet.Name -Id $buttonId -Updated $updated)
      }
      if ($updated) {
        $result.logs = @("updated workbook button " + $buttonName)
      } else {
        $result.logs = @("added workbook button " + $buttonName)
      }
    }
    "list" {
      $buttons = New-Object System.Collections.Generic.List[object]
      foreach ($worksheet in @($workbook.Worksheets)) {
        if (-not [string]::IsNullOrWhiteSpace($Sheet) -and $worksheet.Name -ne $Sheet) {
          continue
        }
        $buttonCollection = $worksheet.Buttons()
        for ($i = 1; $i -le $buttonCollection.Count; $i++) {
          $button = $buttonCollection.Item($i)
          if ($button.Name -like "xlflow.button.*") {
            $buttonId = $button.Name.Substring("xlflow.button.".Length)
            $buttons.Add((ConvertTo-XlflowUIButtonInfo -Button $button -Sheet $worksheet.Name -Id $buttonId)) | Out-Null
          }
        }
      }
      if (-not [string]::IsNullOrWhiteSpace($Sheet) -and $buttons.Count -eq 0 -and $null -eq (Get-XlflowWorksheet -Workbook $workbook -Sheet $Sheet)) {
        Set-XlflowError -Result $result -Code "sheet_not_found" -Message ("Worksheet not found: " + $Sheet) -Source "Excel"
        break
      }
      $result.ui = [ordered]@{ buttons = $buttons.ToArray() }
      $result.logs = @("found $($buttons.Count) xlflow-managed button(s)")
    }
    "remove" {
      $buttonId = ConvertTo-XlflowUIButtonId -Value $Id
      if ([string]::IsNullOrWhiteSpace($buttonId)) {
        Set-UIButtonValidationError -Message "-Id is required."
        break
      }
      $buttonName = ConvertTo-XlflowUIButtonName -Id $buttonId
      $removed = $null
      foreach ($worksheet in @($workbook.Worksheets)) {
        if (-not [string]::IsNullOrWhiteSpace($Sheet) -and $worksheet.Name -ne $Sheet) {
          continue
        }
        $button = Get-XlflowUIButton -Worksheet $worksheet -Name $buttonName
        if ($null -ne $button) {
          $removed = ConvertTo-XlflowUIButtonInfo -Button $button -Sheet $worksheet.Name -Id $buttonId
          $button.Delete() | Out-Null
          break
        }
      }
      if ($null -eq $removed) {
        if (-not [string]::IsNullOrWhiteSpace($Sheet) -and $null -eq (Get-XlflowWorksheet -Workbook $workbook -Sheet $Sheet)) {
          Set-XlflowError -Result $result -Code "sheet_not_found" -Message ("Worksheet not found: " + $Sheet) -Source "Excel"
          break
        }
        Set-XlflowError -Result $result -Code "button_not_found" -Message ("Button not found: " + $buttonName) -Source "Excel"
        break
      }
      $saveWorkbook = $true
      $result.ui = [ordered]@{ button = $removed }
      $result.logs = @("removed workbook button " + $buttonName)
    }
  }
  if ($null -ne $workbook) {
    $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  }
  if ($result.status -eq "ok") {
    if ($saveState.needs_save) {
      Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes."
    }
  }
} catch {
  Set-XlflowError -Result $result -Code "ui_button_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
  if ($null -ne $WorkbookPath) {
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
    $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
    if ($saveState.needs_save) {
      Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes."
    }
  }
} finally {
  if ($result.status -ne "ok") {
    $saveWorkbook = $false
  }
  if ($saveWorkbook -and $null -ne $workbook) {
    try {
      $workbook.Save()
      $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
      $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
      $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
      if (-not $result.Contains("warnings") -or $null -eq $result["warnings"]) {
        $result["warnings"] = @()
      }
      $result.warnings = @($result.warnings | Where-Object { $_.code -ne "save_required" })
      if ($saveState.needs_save) {
        Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes."
      }
    } catch {
      Set-XlflowError -Result $result -Code "save_failed" -Message ("Failed to save workbook: " + $_.Exception.Message) -Source "Excel"
      $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $true -NeedsSave $true
      $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
      $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $true -SaveRequired $true -Mode $sessionMode
      if (-not $result.Contains("warnings") -or $null -eq $result["warnings"]) {
        $result["warnings"] = @()
      }
      $result.warnings = @($result.warnings | Where-Object { $_.code -ne "save_required" })
      Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. Run `xlflow save --session` to persist workbook changes."
    }
  }
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
