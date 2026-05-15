param(
  [string]$Action = "",
  [string]$WorkbookPath = "",
  [string]$SpecPath = "",
  [string]$SpecJson64 = "",
  [string]$Overwrite = "false",
  [string]$NoSave = "false",
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "form write"
$excel = $null
$workbook = $null
$sessionAttached = $false
$sessionMode = "none"
$saved = $false
$jsonWritten = $false
$phase = "validate_args"

function Set-FormWriteValidationError {
  param(
    [string]$Code,
    [string]$Message
  )

  Set-XlflowError -Result $result -Code $Code -Message $Message -Source "xlflow" -Phase "validate_args"
}

function Get-XlflowFormWriteArgsCode {
  param([string]$CurrentAction)

  if ([string]::Equals($CurrentAction, "apply", [System.StringComparison]::OrdinalIgnoreCase)) {
    return "form_apply_args_invalid"
  }
  return "form_build_args_invalid"
}

function Get-XlflowFormWriteErrorCode {
  param(
    [string]$CurrentAction,
    [System.Exception]$Exception
  )

  $message = [string]$Exception.Message
  if ($message -like "form_already_exists:*") {
    return "form_already_exists"
  }
  if ($message -like "form_not_found:*") {
    return "form_not_found"
  }
  if ($message -like "unsupported_form_control:*") {
    return "unsupported_form_control"
  }
  if ($message -like "designer_write_failed:*") {
    return "designer_write_failed"
  }
  return $(if ($CurrentAction -eq "apply") { "form_apply_failed" } else { "form_build_failed" })
}

function ConvertFrom-XlflowFormSpecJson64 {
  param([string]$Encoded)

  try {
    $bytes = [System.Convert]::FromBase64String($Encoded)
    $json = [System.Text.Encoding]::UTF8.GetString($bytes)
    if ([string]::IsNullOrWhiteSpace($json)) {
      throw "decoded form spec was empty"
    }
    return ($json | ConvertFrom-Json)
  } catch {
    throw ("invalid form spec payload: " + $_.Exception.Message)
  }
}

function Get-XlflowFormControlProgId {
  param($ControlSpec)

  $progId = [string]$ControlSpec.progId
  if (-not [string]::IsNullOrWhiteSpace($progId)) {
    return $progId
  }

  switch (([string]$ControlSpec.type).ToLowerInvariant()) {
    "label" { return "Forms.Label.1" }
    "textbox" { return "Forms.TextBox.1" }
    "combobox" { return "Forms.ComboBox.1" }
    "listbox" { return "Forms.ListBox.1" }
    "commandbutton" { return "Forms.CommandButton.1" }
    "checkbox" { return "Forms.CheckBox.1" }
    "optionbutton" { return "Forms.OptionButton.1" }
    "frame" { return "Forms.Frame.1" }
    default { throw ("unsupported_form_control: unsupported control type '" + [string]$ControlSpec.type + "'") }
  }
}

function Get-XlflowControlSpecId {
  param($ControlSpec)

  $identifier = [string]$ControlSpec.id
  if (-not [string]::IsNullOrWhiteSpace($identifier)) {
    return $identifier
  }
  return [string]$ControlSpec.name
}

function Get-XlflowSpecMemberValue {
  param(
    $SpecObject,
    [string]$DirectName,
    [string]$ObservedName = ""
  )

  if ($null -ne $SpecObject.PSObject.Properties[$DirectName]) {
    return $SpecObject.$DirectName
  }
  if (-not [string]::IsNullOrWhiteSpace($ObservedName) -and
      $null -ne $SpecObject.PSObject.Properties["observed"] -and
      $null -ne $SpecObject.observed -and
      $null -ne $SpecObject.observed.PSObject.Properties[$ObservedName]) {
    return $SpecObject.observed.$ObservedName
  }
  return $null
}

function Get-XlflowBuildFormValue {
  param(
    $FormSpec,
    [string]$BuildName,
    [string]$ObservedName
  )

  if ($null -ne $FormSpec.PSObject.Properties["build"] -and
      $null -ne $FormSpec.build -and
      $null -ne $FormSpec.build.PSObject.Properties[$BuildName]) {
    return $FormSpec.build.$BuildName
  }
  if ($null -ne $FormSpec.PSObject.Properties[$BuildName]) {
    return $FormSpec.$BuildName
  }
  if ($null -ne $FormSpec.PSObject.Properties["observed"] -and
      $null -ne $FormSpec.observed -and
      $null -ne $FormSpec.observed.PSObject.Properties[$ObservedName]) {
    return $FormSpec.observed.$ObservedName
  }
  return $null
}

function Get-XlflowControlSpecChildren {
  param(
    $AllControls,
    [string]$ParentId,
    $FallbackChildren = $null
  )

  $children = @($AllControls | Where-Object {
    $candidateParentId = [string]$_.parentId
    -not [string]::IsNullOrWhiteSpace($candidateParentId) -and
      [string]::Equals($candidateParentId, $ParentId, [System.StringComparison]::Ordinal)
  })
  if ($children.Count -gt 0) {
    return $children | Sort-Object -Property @{ Expression = {
      if ($null -ne $_.PSObject.Properties["zIndex"]) { [int]$_.zIndex } else { [int]::MaxValue }
    } }
  }
  if ($null -ne $FallbackChildren) {
    return @($FallbackChildren | Where-Object { $null -ne $_ })
  }
  return @()
}

function Get-XlflowRootControlSpecs {
  param($Spec)

  $allControls = @($Spec.controls | Where-Object { $null -ne $_ })
  $roots = @($allControls | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.parentId) })
  if ($roots.Count -gt 0) {
    return $roots | Sort-Object -Property @{ Expression = {
      if ($null -ne $_.PSObject.Properties["zIndex"]) { [int]$_.zIndex } else { [int]::MaxValue }
    } }
  }
  return $allControls
}

function Set-XlflowVBComponentProperty {
  param(
    $Component,
    [string]$PropertyName,
    $Value
  )

  try {
    foreach ($property in @($Component.Properties)) {
      if ([string]::Equals([string]$property.Name, $PropertyName, [System.StringComparison]::OrdinalIgnoreCase)) {
        $property.Value = $Value
        return $true
      }
    }
  } catch {
    return $false
  }
  return $false
}

function Add-XlflowFormWriteWarning {
  param(
    [string]$Code,
    [string]$Message,
    [string]$ControlName = "",
    [string]$FieldPath = ""
  )

  if (-not $result.Contains("warnings") -or $null -eq $result["warnings"]) {
    $result["warnings"] = @()
  }
  $warning = [ordered]@{
    code = $Code
    message = $Message
  }
  if (-not [string]::IsNullOrWhiteSpace($ControlName)) {
    $warning.control = $ControlName
  }
  if (-not [string]::IsNullOrWhiteSpace($FieldPath)) {
    $warning.field_path = $FieldPath
  }
  $result["warnings"] += $warning
}

function Add-XlflowFormContractWarnings {
  param($Spec)

  $formSpec = $Spec.form
  $hasFormSizeExpectation = $false
  foreach ($property in @("width", "height")) {
    if ($null -ne $formSpec.PSObject.Properties["build"] -and
        $null -ne $formSpec.build -and
        $null -ne $formSpec.build.PSObject.Properties[$property]) {
      $hasFormSizeExpectation = $true
      break
    }
    if ($null -ne $formSpec.PSObject.Properties[$property]) {
      $hasFormSizeExpectation = $true
      break
    }
    if ($null -ne $formSpec.PSObject.Properties["observed"] -and
        $null -ne $formSpec.observed -and
        $null -ne $formSpec.observed.PSObject.Properties[$property]) {
      $hasFormSizeExpectation = $true
      break
    }
  }
  if ($hasFormSizeExpectation) {
    Add-XlflowFormWriteWarning -Code "best_effort_form_size" -Message "Form-level width/height are best-effort in Designer build and may not round-trip through Excel VBIDE Designer APIs. Field scope: form.observed.width, form.observed.height, form.build.width, form.build.height." -FieldPath "form.observed.width"
  }

  $listStateControls = New-Object System.Collections.Generic.List[string]
  foreach ($controlSpec in @($Spec.controls | Where-Object { $null -ne $_ })) {
    $controlType = ([string]$controlSpec.type).Trim().ToLowerInvariant()
    if ($controlType -notin @("combobox", "listbox")) {
      continue
    }
    $hasListState = $false
    if ($null -ne $controlSpec.PSObject.Properties["list"] -or $null -ne $controlSpec.PSObject.Properties["selectedIndex"]) {
      $hasListState = $true
    }
    if (-not $hasListState -and
        $null -ne $controlSpec.PSObject.Properties["observed"] -and
        $null -ne $controlSpec.observed -and
        ($null -ne $controlSpec.observed.PSObject.Properties["list"] -or $null -ne $controlSpec.observed.PSObject.Properties["selectedIndex"])) {
      $hasListState = $true
    }
    if ($hasListState) {
      $listStateControls.Add([string]$controlSpec.name)
    }
  }
  if ($listStateControls.Count -gt 0) {
    $controlNames = ($listStateControls | Select-Object -Unique) -join ", "
    Add-XlflowFormWriteWarning -Code "best_effort_list_state" -Message ("Design-time ComboBox/ListBox list and selectedIndex are best-effort during build and should be treated as observed-only for round-trip expectations. Field scope includes controls[*].list and controls[*].selectedIndex. Controls: " + $controlNames + ".") -FieldPath "controls[*].selectedIndex"
  }
}

function Set-XlflowFormProperty {
  param(
    $Target,
    [string]$PropertyName,
    $Value,
    [string]$ControlName = ""
  )

  try {
    $Target.$PropertyName = $Value
    return $true
  } catch {
    Add-XlflowFormWriteWarning -Code "unsupported_property" -Message ("Skipping unsupported property '" + $PropertyName + "'. " + $_.Exception.Message) -ControlName $ControlName
    return $false
  }
}

function Clear-XlflowDesignerControls {
  param($Container)

  try {
    while ($Container.Controls.Count -gt 0) {
      $control = $Container.Controls.Item($Container.Controls.Count - 1)
      $Container.Controls.Remove([string]$control.Name)
    }
  } catch {
    throw ("designer_write_failed: failed to clear existing controls. " + $_.Exception.Message)
  }
}

function Set-XlflowControlListItems {
  param(
    $Control,
    $ControlSpec
  )

  if ($null -eq $ControlSpec.PSObject.Properties["list"]) {
    return
  }
  try {
    if ($null -ne $Control.ListCount) {
      if ($Control.ListCount -gt 0) {
        $Control.Clear()
      }
    }
  } catch {
    Add-XlflowFormWriteWarning -Code "unsupported_property" -Message ("Skipping unsupported list reset. " + $_.Exception.Message) -ControlName ([string]$ControlSpec.name)
    return
  }

  foreach ($item in @($ControlSpec.list)) {
    try {
      $Control.AddItem([string]$item)
    } catch {
      Add-XlflowFormWriteWarning -Code "unsupported_property" -Message ("Skipping unsupported list item append. " + $_.Exception.Message) -ControlName ([string]$ControlSpec.name)
      return
    }
  }
}

function Set-XlflowDesignerControlProperties {
  param(
    $Control,
    $ControlSpec
  )

  $controlName = [string]$ControlSpec.name
  foreach ($property in @("caption", "left", "top", "width", "height", "tabIndex", "enabled", "visible")) {
    $observedName = $property
    if ($property -eq "tabIndex") {
      $observedName = "tabIndex"
    }
    $value = Get-XlflowSpecMemberValue -SpecObject $ControlSpec -DirectName $property -ObservedName $observedName
    if ($null -eq $value) {
      continue
    }
    switch ($property) {
      "caption" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "Caption" -Value ([string]$value) -ControlName $controlName }
      "left" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "Left" -Value ([double]$value) -ControlName $controlName }
      "top" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "Top" -Value ([double]$value) -ControlName $controlName }
      "width" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "Width" -Value ([double]$value) -ControlName $controlName }
      "height" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "Height" -Value ([double]$value) -ControlName $controlName }
      "tabIndex" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "TabIndex" -Value ([int]$value) -ControlName $controlName }
      "enabled" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "Enabled" -Value ([bool]$value) -ControlName $controlName }
      "visible" { $null = Set-XlflowFormProperty -Target $Control -PropertyName "Visible" -Value ([bool]$value) -ControlName $controlName }
    }
  }
  $value = Get-XlflowSpecMemberValue -SpecObject $ControlSpec -DirectName "value" -ObservedName "value"
  if ($null -ne $value) {
    $null = Set-XlflowFormProperty -Target $Control -PropertyName "Value" -Value $value -ControlName $controlName
  }
  $textValue = Get-XlflowSpecMemberValue -SpecObject $ControlSpec -DirectName "text" -ObservedName "text"
  if ($null -ne $textValue) {
    $null = Set-XlflowFormProperty -Target $Control -PropertyName "Text" -Value ([string]$textValue) -ControlName $controlName
  }
  Set-XlflowControlListItems -Control $Control -ControlSpec $ControlSpec
  $selectedIndex = Get-XlflowSpecMemberValue -SpecObject $ControlSpec -DirectName "selectedIndex" -ObservedName "selectedIndex"
  if ($null -ne $selectedIndex) {
    $null = Set-XlflowFormProperty -Target $Control -PropertyName "ListIndex" -Value ([int]$selectedIndex) -ControlName $controlName
  }
}

function Add-XlflowDesignerControl {
  param(
    $Parent,
    $ControlSpec,
    $AllControls
  )

  $progId = Get-XlflowFormControlProgId -ControlSpec $ControlSpec
  $controlName = [string]$ControlSpec.name
  try {
    $control = $Parent.Controls.Add($progId, $controlName, $true)
  } catch {
    throw ("designer_write_failed: failed to add control '" + $controlName + "' with ProgID '" + $progId + "'. " + $_.Exception.Message)
  }

  Set-XlflowDesignerControlProperties -Control $control -ControlSpec $ControlSpec
  foreach ($child in @(Get-XlflowControlSpecChildren -AllControls $AllControls -ParentId (Get-XlflowControlSpecId -ControlSpec $ControlSpec) -FallbackChildren $ControlSpec.controls)) {
    Add-XlflowDesignerControl -Parent $control -ControlSpec $child -AllControls $AllControls
  }
}

function Set-XlflowDesignerFormProperties {
  param(
    $Designer,
    $Component,
    $Spec
  )

  $caption = Get-XlflowBuildFormValue -FormSpec $Spec.form -BuildName "caption" -ObservedName "caption"
  if ($null -ne $caption) {
    $null = Set-XlflowFormProperty -Target $Designer -PropertyName "Caption" -Value ([string]$caption)
  }
  $width = Get-XlflowBuildFormValue -FormSpec $Spec.form -BuildName "width" -ObservedName "width"
  if ($null -ne $width) {
    if (-not (Set-XlflowVBComponentProperty -Component $Component -PropertyName "Width" -Value ([double]$width))) {
      Add-XlflowFormWriteWarning -Code "unsupported_property" -Message "Skipping unsupported property 'Width'."
    }
  }
  $height = Get-XlflowBuildFormValue -FormSpec $Spec.form -BuildName "height" -ObservedName "height"
  if ($null -ne $height) {
    if (-not (Set-XlflowVBComponentProperty -Component $Component -PropertyName "Height" -Value ([double]$height))) {
      Add-XlflowFormWriteWarning -Code "unsupported_property" -Message "Skipping unsupported property 'Height'."
    }
  }
}

function Get-XlflowVBComponentByName {
  param(
    $VBProject,
    [string]$Name
  )

  try {
    return $VBProject.VBComponents.Item($Name)
  } catch {
    return $null
  }
}

function New-XlflowUserFormComponent {
  param(
    $VBProject,
    [string]$Name
  )

  try {
    $component = $VBProject.VBComponents.Add(3)
    $component.Name = $Name
    return $component
  } catch {
    throw ("designer_write_failed: failed to create UserForm '" + $Name + "'. " + $_.Exception.Message)
  }
}

function New-XlflowFormRestoreDirectory {
  $path = Join-Path ([System.IO.Path]::GetTempPath()) ("xlflow-form-restore-" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $path -Force | Out-Null
  return $path
}

function Export-XlflowVBComponentBackup {
  param(
    $Component,
    [string]$Directory
  )

  $exportPath = Join-Path $Directory ([string]$Component.Name + ".frm")
  try {
    $Component.Export($exportPath)
    return $exportPath
  } catch {
    throw ("designer_write_failed: failed to export existing UserForm '" + [string]$Component.Name + "' before overwrite. " + $_.Exception.Message)
  }
}

function Import-XlflowVBComponentBackup {
  param(
    $VBProject,
    [string]$ExportPath,
    [string]$ExpectedName
  )

  try {
    $restored = $VBProject.VBComponents.Import($ExportPath)
    if ($null -eq $restored) {
      throw "VBComponents.Import returned no component."
    }
    if (-not [string]::Equals([string]$restored.Name, $ExpectedName, [System.StringComparison]::OrdinalIgnoreCase)) {
      try {
        $restored.Name = $ExpectedName
      } catch {
        throw ("restored component name '" + [string]$restored.Name + "' did not match expected '" + $ExpectedName + "'. " + $_.Exception.Message)
      }
    }
    return $restored
  } catch {
    throw ("designer_write_failed: failed to restore original UserForm '" + $ExpectedName + "' after overwrite failure. " + $_.Exception.Message)
  }
}

function Remove-XlflowVBComponentChecked {
  param(
    $VBProject,
    [string]$Name
  )

  try {
    $component = $VBProject.VBComponents.Item($Name)
    if ([int]$component.Type -ne 3) {
      throw ("form_already_exists: component '" + $Name + "' exists but is not a UserForm.")
    }
    $VBProject.VBComponents.Remove($component)
  } catch {
    if ([string]$_.Exception.Message -like "form_already_exists:*") {
      throw $_.Exception.Message
    }
    throw ("designer_write_failed: failed to remove existing component '" + $Name + "'. " + $_.Exception.Message)
  }
}

function Invoke-XlflowFormBuild {
  param(
    $VBProject,
    $Workbook,
    $Spec,
    [bool]$AllowOverwrite,
    [bool]$CanCheckpointSave
  )

  $formName = [string]$Spec.form.name
  $existing = Get-XlflowVBComponentByName -VBProject $VBProject -Name $formName
  $restorePath = $null
  $restoreDirectory = $null
  $removedExisting = $false
  $component = $null
  if ($null -ne $existing) {
    if (-not $AllowOverwrite) {
      throw ("form_already_exists: UserForm '" + $formName + "' already exists.")
    }
    $restoreDirectory = New-XlflowFormRestoreDirectory
    try {
      $restorePath = Export-XlflowVBComponentBackup -Component $existing -Directory $restoreDirectory
      Remove-XlflowVBComponentChecked -VBProject $VBProject -Name $formName
      $removedExisting = $true
      if (-not $CanCheckpointSave) {
        throw "designer_write_failed: overwrite requires an intermediate workbook save, but save is disabled for this command."
      }
      try {
        $Workbook.Save()
      } catch {
        throw ("designer_write_failed: failed to save workbook after removing existing UserForm '" + $formName + "'. " + $_.Exception.Message)
      }

      $component = New-XlflowUserFormComponent -VBProject $VBProject -Name $formName
      try {
        $designer = $component.Designer
      } catch {
        throw ("designer_write_failed: failed to access Designer for '" + $formName + "'. " + $_.Exception.Message)
      }
      Set-XlflowDesignerFormProperties -Designer $designer -Component $component -Spec $Spec
      $allControls = @($Spec.controls | Where-Object { $null -ne $_ })
      foreach ($controlSpec in @(Get-XlflowRootControlSpecs -Spec $Spec)) {
        Add-XlflowDesignerControl -Parent $designer -ControlSpec $controlSpec -AllControls $allControls
      }
    } catch {
      if ($removedExisting) {
        try {
          $partial = Get-XlflowVBComponentByName -VBProject $VBProject -Name $formName
          if ($null -ne $partial) {
            [void](Remove-XlflowVBComponentChecked -VBProject $VBProject -Name $formName)
          }
        } catch {
          throw ("designer_write_failed: overwrite failed for UserForm '" + $formName + "' and cleanup of the partial replacement also failed. " + $_.Exception.Message)
        }
        Import-XlflowVBComponentBackup -VBProject $VBProject -ExportPath $restorePath -ExpectedName $formName | Out-Null
        try {
          $Workbook.Save()
        } catch {
          throw ("designer_write_failed: restored original UserForm '" + $formName + "' after overwrite failure, but saving the restoration failed. " + $_.Exception.Message)
        }
      }
      throw
    } finally {
      if (-not [string]::IsNullOrWhiteSpace($restoreDirectory) -and (Test-Path -LiteralPath $restoreDirectory)) {
        Remove-Item -LiteralPath $restoreDirectory -Recurse -Force -ErrorAction SilentlyContinue
      }
    }
    return
  }
  $component = New-XlflowUserFormComponent -VBProject $VBProject -Name $formName
  try {
    $designer = $component.Designer
  } catch {
    throw ("designer_write_failed: failed to access Designer for '" + $formName + "'. " + $_.Exception.Message)
  }
  Set-XlflowDesignerFormProperties -Designer $designer -Component $component -Spec $Spec
  $allControls = @($Spec.controls | Where-Object { $null -ne $_ })
  foreach ($controlSpec in @(Get-XlflowRootControlSpecs -Spec $Spec)) {
    Add-XlflowDesignerControl -Parent $designer -ControlSpec $controlSpec -AllControls $allControls
  }
}

function Invoke-XlflowFormApply {
  param(
    $VBProject,
    $Spec
  )

  $formName = [string]$Spec.form.name
  $component = Get-XlflowVBComponentByName -VBProject $VBProject -Name $formName
  if ($null -eq $component -or [int]$component.Type -ne 3) {
    throw ("form_not_found: UserForm '" + $formName + "' was not found in the workbook.")
  }
  try {
    $designer = $component.Designer
  } catch {
    throw ("designer_write_failed: failed to access Designer for '" + $formName + "'. " + $_.Exception.Message)
  }
  Clear-XlflowDesignerControls -Container $designer
  Set-XlflowDesignerFormProperties -Designer $designer -Component $component -Spec $Spec
  $allControls = @($Spec.controls | Where-Object { $null -ne $_ })
  foreach ($controlSpec in @(Get-XlflowRootControlSpecs -Spec $Spec)) {
    Add-XlflowDesignerControl -Parent $designer -ControlSpec $controlSpec -AllControls $allControls
  }
}

try {
  $normalizedAction = ([string]$Action).Trim().ToLowerInvariant()
  if ($normalizedAction -notin @("build", "apply")) {
    Set-FormWriteValidationError -Code (Get-XlflowFormWriteArgsCode -CurrentAction $normalizedAction) -Message ("unsupported form action: " + $Action)
    $jsonWritten = $true
    Write-XlflowJson -Result $result
    exit
  }
  if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    Set-FormWriteValidationError -Code (Get-XlflowFormWriteArgsCode -CurrentAction $normalizedAction) -Message "WorkbookPath is required."
    $jsonWritten = $true
    Write-XlflowJson -Result $result
    exit
  }
  if ([string]::IsNullOrWhiteSpace($SpecJson64)) {
    Set-FormWriteValidationError -Code (Get-XlflowFormWriteArgsCode -CurrentAction $normalizedAction) -Message "SpecJson64 is required."
    $jsonWritten = $true
    Write-XlflowJson -Result $result
    exit
  }
  if ((ConvertTo-XlflowBool $NoSave) -and -not (ConvertTo-XlflowBool $UseSession)) {
    Set-FormWriteValidationError -Code (Get-XlflowFormWriteArgsCode -CurrentAction $normalizedAction) -Message "--NoSave requires --UseSession."
    $jsonWritten = $true
    Write-XlflowJson -Result $result
    exit
  }
  if ($normalizedAction -eq "build" -and (ConvertTo-XlflowBool $Overwrite) -and (ConvertTo-XlflowBool $NoSave)) {
    Set-FormWriteValidationError -Code (Get-XlflowFormWriteArgsCode -CurrentAction $normalizedAction) -Message "--overwrite cannot be combined with --NoSave because Excel requires an intermediate save before recreating the UserForm."
    $jsonWritten = $true
    Write-XlflowJson -Result $result
    exit
  }

  $spec = ConvertFrom-XlflowFormSpecJson64 -Encoded $SpecJson64
  Add-XlflowFormContractWarnings -Spec $spec
  $phase = "open_workbook"
  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode

  $phase = "write_designer"
  if ($normalizedAction -eq "build") {
    Invoke-XlflowFormBuild -VBProject $workbook.VBProject -Workbook $workbook -Spec $spec -AllowOverwrite (ConvertTo-XlflowBool $Overwrite) -CanCheckpointSave (-not (ConvertTo-XlflowBool $NoSave))
  } else {
    Invoke-XlflowFormApply -VBProject $workbook.VBProject -Spec $spec
  }

  $phase = "save_workbook"
  if ($sessionAttached) {
    if (-not (ConvertTo-XlflowBool $NoSave)) {
      $workbook.Save()
      $saved = $true
    }
  } else {
    $workbook.Save()
    $saved = $true
  }

  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $saved -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  $result.forms = [ordered]@{
    name = [string]$spec.form.name
    basis = [string]$spec.basis
    action = $normalizedAction
    coordinate_system = [string]$spec.coordinateSystem
    control_count = @($spec.controls).Count
    spec_path = $SpecPath
    overwrite = (ConvertTo-XlflowBool $Overwrite)
  }
  if ($null -ne $spec.form.PSObject.Properties["caption"]) {
    $result.forms.caption = [string]$spec.form.caption
  }
  Add-XlflowHint -Result $result -Code "userform_review_commands" -Message ("Review the result with `xlflow inspect form " + [string]$spec.form.name + " --designer --json` or `xlflow form export-image " + [string]$spec.form.name + " --out <path>`.")
  if ($saveState.needs_save) {
    Add-XlflowStateWarning -Result $result -Code "save_required" -Message ("The live workbook is newer than disk after `form " + $normalizedAction + "`. Run `xlflow save --session` before relying on disk-backed inspect, pull, or source review.")
  }
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), ($normalizedAction + " form " + [string]$spec.form.name + " from " + $SpecPath)) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  if ($null -eq $result.error) {
    $code = Get-XlflowFormWriteErrorCode -CurrentAction $normalizedAction -Exception $_.Exception
    $message = [string]$_.Exception.Message
    foreach ($prefix in @("form_already_exists: ", "form_not_found: ", "unsupported_form_control: ", "designer_write_failed: ")) {
      if ($message.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        $message = $message.Substring($prefix.Length)
        break
      }
    }
    Set-XlflowError -Result $result -Code $code -Message $message -Source $_.Exception.Source -Number $_.Exception.HResult -Phase $phase
  }
  if ($null -ne $workbook -and -not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
    $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Saved $false -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
    $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
    $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  }
} finally {
  if ($null -ne $workbook -and -not $sessionAttached) {
    try {
      $workbook.Close($false) | Out-Null
    } catch {
      Write-Verbose ("failed to close workbook after form write: " + $_.Exception.Message)
    }
  }
  if ($null -ne $excel -and -not $sessionAttached) {
    try {
      $excel.Quit() | Out-Null
    } catch {
      Write-Verbose ("failed to quit Excel after form write: " + $_.Exception.Message)
    }
  }
  Release-XlflowComReferences -Workbook $workbook -Excel $excel
  if (-not $jsonWritten) {
    Write-XlflowJson -Result $result
  }
}
