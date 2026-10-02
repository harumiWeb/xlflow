[CmdletBinding()]
param(
    [ValidateSet('create', 'verify')][string]$Phase = 'create',
    [string]$WorkspacePath = '',
    [string]$WorkbookPath = '',
    [string]$ExpectedPath = '',
    [string]$NormalizedWorkbookPath = '',
    [string]$FixtureDirectory = ''
)

# Local developer harness for the UserForm generation gate. This script is
# intentionally separate from test-userform-mutation-e2e.ps1. It requires
# trusted VBIDE access and must never run from ordinary tests or CI.
# Verify executes workbook VBA; use only a trusted generated artifact.
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$utf8 = [Text.UTF8Encoding]::new($false)
$script:comRefs = [Collections.Generic.List[object]]::new()
$script:excel = $null
$script:workbook = $null
$script:excelPID = 0
$script:cleanupErrors = [Collections.Generic.List[string]]::new()
$script:formNames = @('GenerationBaseline', 'EmptyForm')
$script:generatedFormNames = @('GeneratedForm', 'GeneratedEmptyForm')
$script:controlSpecs = @(
    [pscustomobject]@{ type = 'Label'; progId = 'Forms.Label.1'; name = 'LabelMain'; left = 6; top = 6; width = 72; height = 18 },
    [pscustomobject]@{ type = 'TextBox'; progId = 'Forms.TextBox.1'; name = 'TextBoxMain'; left = 114; top = 6; width = 120; height = 18 },
    [pscustomobject]@{ type = 'CommandButton'; progId = 'Forms.CommandButton.1'; name = 'CommandButtonMain'; left = 6; top = 30; width = 72; height = 24 },
    [pscustomobject]@{ type = 'CheckBox'; progId = 'Forms.CheckBox.1'; name = 'CheckBoxMain'; left = 114; top = 30; width = 72; height = 18 },
    [pscustomobject]@{ type = 'OptionButton'; progId = 'Forms.OptionButton.1'; name = 'OptionButtonMain'; left = 6; top = 54; width = 72; height = 18 },
    [pscustomobject]@{ type = 'ToggleButton'; progId = 'Forms.ToggleButton.1'; name = 'ToggleButtonMain'; left = 114; top = 54; width = 72; height = 18 },
    [pscustomobject]@{ type = 'ComboBox'; progId = 'Forms.ComboBox.1'; name = 'ComboBoxMain'; left = 6; top = 78; width = 120; height = 18 },
    [pscustomobject]@{ type = 'ListBox'; progId = 'Forms.ListBox.1'; name = 'ListBoxMain'; left = 6; top = 102; width = 120; height = 72 },
    [pscustomobject]@{ type = 'SpinButton'; progId = 'Forms.SpinButton.1'; name = 'SpinButtonMain'; left = 216; top = 6; width = 18; height = 36 },
    [pscustomobject]@{ type = 'ScrollBar'; progId = 'Forms.ScrollBar.1'; name = 'ScrollBarMain'; left = 114; top = 78; width = 120; height = 18 },
    [pscustomobject]@{ type = 'Image'; progId = 'Forms.Image.1'; name = 'ImageMain'; left = 162; top = 108; width = 72; height = 72 }
)
$script:newControlNames = @('ToggleButtonMain', 'SpinButtonMain', 'ScrollBarMain', 'ImageMain')

function Hold-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $script:comRefs.Add($Value)
    }
    return ,$Value
}

function Release-Children {
    for ($i = $script:comRefs.Count - 1; $i -ge 0; $i--) {
        $value = $script:comRefs[$i]
        if ([Runtime.InteropServices.Marshal]::IsComObject($value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
        }
    }
    $script:comRefs.Clear()
}

function Write-Json([string]$Path, [object]$Value) {
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 30), $utf8)
}

function Export-VbaProject([string]$Source, [string]$Destination) {
    if (Test-Path -LiteralPath $Destination) {
        throw "Refusing to overwrite existing binary evidence: $Destination"
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression
    # Excel may still hold the saved workbook in this session. Read the saved
    # package with sharing rather than reopening Excel for every observation.
    $packageStream = [IO.File]::Open($Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $archive = [IO.Compression.ZipArchive]::new($packageStream, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $entry = $archive.GetEntry('xl/vbaProject.bin')
        if ($null -eq $entry) { throw "VBA project missing from $Source" }
        $inputStream = $entry.Open()
        try {
            $outputStream = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew)
            try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose() }
        } finally { $inputStream.Dispose() }
    } finally { $archive.Dispose(); $packageStream.Dispose() }
}

function Get-BinaryEvidence([string]$Path) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [IO.File]::ReadAllBytes($Path)
        $digest = ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    } finally { $sha.Dispose() }
    return [pscustomobject]@{
        entry = 'xl/vbaProject.bin'
        path = $Path
        length = [int64]$bytes.Length
        sha256 = $digest
    }
}

function Get-OptionalProperty([object]$Object, [string]$Name) {
    try {
        $flags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::GetProperty
        $value = $Object.GetType().InvokeMember($Name, $flags, $null, $Object, @(), [Globalization.CultureInfo]::InvariantCulture)
        return [pscustomobject]@{ found = $true; value = $value }
    } catch {
        return [pscustomobject]@{ found = $false; value = $null }
    }
}

function Convert-ObservedValue([object]$Value) {
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $null }
    if ([Runtime.InteropServices.Marshal]::IsComObject($Value)) { return $null }
    if ($Value -is [System.Array]) {
        $items = @()
        foreach ($item in $Value) { $items += Convert-ObservedValue $item }
        return ,$items
    }
    if ($Value -is [bool] -or $Value -is [string] -or $Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        return $Value
    }
    return [string]$Value
}

function Add-ObservedField([System.Collections.IDictionary]$Row, [object]$Control, [string]$PropertyName, [string]$OutputName = '') {
    $property = Get-OptionalProperty $Control $PropertyName
    if ($property.found) {
        if (-not $OutputName) { $OutputName = $PropertyName.Substring(0, 1).ToLowerInvariant() + $PropertyName.Substring(1) }
        $Row[$OutputName] = Convert-ObservedValue $property.value
    }
}

function Get-ControlType([object]$Control) {
    $progId = Get-OptionalProperty $Control 'ProgID'
    if ($progId.found -and $null -ne $progId.value) {
        $text = [string]$progId.value
        if ($text -match '^Forms\.(?<type>[^.]+)\.\d+$') { return $Matches.type }
    }
    $className = Get-OptionalProperty $Control 'ClassName'
    if ($className.found -and $className.value) { return [string]$className.value }
    $known = @($script:controlSpecs | Where-Object { $_.name -ceq [string]$Control.Name })
    if ($known.Count -eq 1) { return [string]$known[0].type }
    return ''
}

function Get-ControlSnapshot([object]$Control) {
    $row = [ordered]@{
        name = [string]$Control.Name
        type = Get-ControlType $Control
        left = [double]$Control.Left
        top = [double]$Control.Top
        width = [double]$Control.Width
        height = [double]$Control.Height
        enabled = [bool]$Control.Enabled
        visible = [bool]$Control.Visible
    }
    Add-ObservedField $row $Control 'TabIndex' 'tabIndex'

    Add-ObservedField $row $Control 'Caption' 'caption'
    Add-ObservedField $row $Control 'Text' 'text'
    Add-ObservedField $row $Control 'Value' 'value'
    if (-not $row.Contains('text') -and $row.Contains('value')) { $row.text = $row.value }
    foreach ($property in @('Min', 'Max', 'SmallChange', 'LargeChange', 'Orientation', 'Delay', 'ProportionalThumb', 'ColumnCount', 'BoundColumn', 'ListRows', 'MatchEntry', 'Style', 'PictureSizeMode', 'PictureAlignment', 'SpecialEffect', 'BorderStyle', 'AutoSize', 'BackStyle')) {
        Add-ObservedField $row $Control $property
    }
    $properties = [ordered]@{}
    foreach ($property in @('MaxLength', 'Tag', 'ControlTipText', 'GroupName', 'BackColor', 'ForeColor', 'BorderColor')) {
        $propertyValue = Get-OptionalProperty $Control $property
        if ($propertyValue.found) { $properties[$property] = Convert-ObservedValue $propertyValue.value }
    }
    if ($properties.Count -gt 0) { $row.properties = [pscustomobject]$properties }
    $row.type = [string]$row.type
    return [pscustomobject]$row
}

function Get-FormComponent([string]$FormName) {
    $project = Hold-Com $script:workbook.VBProject
    $components = Hold-Com $project.VBComponents
    return (Hold-Com ($components.Item($FormName)))
}

function Get-Designer([string]$FormName) {
    $component = Get-FormComponent $FormName
    return (Hold-Com $component.Designer)
}

function Get-FormSnapshot([string]$FormName) {
    $designer = Get-Designer $FormName
    $caption = Get-OptionalProperty $designer 'Caption'
    $root = [ordered]@{}
    foreach ($property in @('Width', 'Height', 'InsideWidth', 'InsideHeight')) {
        $value = Get-OptionalProperty $designer $property
        if ($value.found -and $null -ne $value.value) {
            $root[$property.Substring(0, 1).ToLowerInvariant() + $property.Substring(1)] = [double]$value.value
        }
    }
    if ($Phase -eq 'verify') {
        # ClientHeight in persisted VBFrame includes space excluded by the
        # runtime MSForms InsideHeight property; compare matching representations.
        $exportPath = Join-Path ([IO.Path]::GetDirectoryName($WorkbookPath)) ($FormName + '-' + [Guid]::NewGuid().ToString('N') + '.frm')
        $component = Get-FormComponent $FormName
        $component.Export($exportPath)
        $exported = [IO.File]::ReadAllText($exportPath)
        foreach ($dimension in @('Width', 'Height')) {
            $match = [regex]::Match($exported, '(?m)^\s*Client' + $dimension + '\s*=\s*(\d+)')
            if (-not $match.Success) { throw "Missing exported Client$dimension for $FormName" }
            $root['client' + $dimension] = [double]$match.Groups[1].Value / 20.0
        }
    }

    $controls = Hold-Com $designer.Controls
    $rows = @()
    for ($i = 0; $i -lt [int]$controls.Count; $i++) {
        $rows += Get-ControlSnapshot (Hold-Com ($controls.Item($i)))
    }
    $formWidth = if ($root.Contains('clientWidth')) { $root.clientWidth } elseif ($root.Contains('width')) { $root.width } else { $null }
    $formHeight = if ($root.Contains('clientHeight')) { $root.clientHeight } elseif ($root.Contains('height')) { $root.height } else { $null }
    return [pscustomobject]@{
        name = $FormName
        form = [pscustomobject]@{
            name = $FormName
            caption = if ($caption.found) { [string]$caption.value } else { $null }
            width = $formWidth
            height = $formHeight
        }
        root = [pscustomobject]$root
        controls = @($rows)
    }
}

function Get-WorkbookObservation([string[]]$Names) {
    $forms = @()
    foreach ($name in $Names) { $forms += Get-FormSnapshot $name }
    return [pscustomobject]@{ forms = @($forms) }
}

function Set-FormClientSize([string]$FormName, [object]$Designer) {
    $insideWidthSet = $false
    $insideHeightSet = $false
    try { $Designer.InsideWidth = 240; $insideWidthSet = $true } catch { $insideWidthSet = $false }
    try { $Designer.InsideHeight = 180; $insideHeightSet = $true } catch { $insideHeightSet = $false }
    if (-not ($insideWidthSet -and $insideHeightSet)) {
        $component = Get-FormComponent $FormName
        $properties = Hold-Com $component.Properties
        $width = Hold-Com ($properties.Item('Width'))
        $height = Hold-Com ($properties.Item('Height'))
        $width.Value = 240
        $height.Value = 180
    }
}

function New-Form([string]$FormName, [string]$Caption, [bool]$Populate) {
    $project = Hold-Com $script:workbook.VBProject
    $components = Hold-Com $project.VBComponents
    $component = Hold-Com ($components.Add(3))
    $designer = Hold-Com $component.Designer
    $component.Name = $FormName
    $designer.Caption = $Caption
    Set-FormClientSize $FormName $designer
    if ($Populate) {
        $controls = Hold-Com $designer.Controls
        foreach ($spec in $script:controlSpecs) {
            $control = Hold-Com ($controls.Add($spec.progId, $spec.name, $true))
            $control.Left = $spec.left
            $control.Top = $spec.top
            $control.Width = $spec.width
            $control.Height = $spec.height
        }
    }
}

function Add-MainModule {
    $project = Hold-Com $script:workbook.VBProject
    $components = Hold-Com $project.VBComponents
    $module = Hold-Com ($components.Add(1))
    $module.Name = 'Main'
    $code = Hold-Com $module.CodeModule
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString(@'
Option Explicit

Public Sub RunGenerationSentinel()
    On Error GoTo Failed
    Dim generated As Object
    Dim blankInstance As Object
    Dim verified As Boolean
    Set generated = VBA.UserForms.Add("GeneratedForm")
    Set blankInstance = VBA.UserForms.Add("GeneratedEmptyForm")
    verified = CallByName(generated, "VerifyGeneration", VbMethod)
    If Not verified Then Err.Raise vbObjectError + 883, "GeneratedForm.VerifyGeneration", "returned False"
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-883-ok"
    Unload generated
    Unload blankInstance
    Exit Sub
Failed:
    On Error Resume Next
    Unload generated
    Unload blankInstance
    On Error GoTo 0
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-883-failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
'@)
}

function Save-Stage([string]$Name, [string[]]$Names) {
    $path = Join-Path $WorkspacePath "$Name.xlsm"
    $before = Get-WorkbookObservation $Names
    Release-Children
    $script:workbook.SaveAs($path, 52)
    $script:workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook)
    $script:workbook = $null
    $books = Hold-Com $script:excel.Workbooks
    $script:workbook = $books.Open($path, 0, $false)
    $after = Get-WorkbookObservation $Names
    Release-Children
    $observation = [pscustomobject]@{
        stage = $Name
        workbook = $path
        beforeSave = $before
        reopened = $after
    }
    Write-Json (Join-Path $WorkspacePath "$Name.json") $observation
    Export-VbaProject $path (Join-Path $WorkspacePath "$Name.bin")
    return [pscustomobject]@{ path = $path; observation = $observation }
}

function Export-FormEvidence([string]$FormName, [string]$Destination) {
    if (Test-Path -LiteralPath $Destination) { throw "Refusing to overwrite form evidence: $Destination" }
    $component = Get-FormComponent $FormName
    $component.Export($Destination)
    if (-not (Test-Path -LiteralPath $Destination)) { throw "Excel did not export $FormName to $Destination" }
    $source = [IO.File]::ReadAllText($Destination)
    $attributes = @(
        $source -split "`r?`n" |
            Where-Object { $_ -match '^\s*Attribute\s+VB_Base\s*=' } |
            ForEach-Object { $_.Trim() }
    )
    return [pscustomobject]@{
        form = $FormName
        path = $Destination
        vbBaseAttributes = @($attributes)
    }
}

function Assert-Value([string]$Path, [object]$Actual, [object]$Expected) {
    # Projection represents two-state control values as True/False strings;
    # Excel exposes the same persisted values as VARIANT_BOOL.
    if ($Path -match '\.value$' -and $Actual -is [bool] -and $Expected -is [string] -and $Expected -in @('True', 'False')) {
        $Expected = [bool]::Parse($Expected)
    }
    if ($Actual -is [bool] -and $Expected -is [string] -and $Expected -match '^(True|False)$') {
        $Expected = [bool]::Parse($Expected)
    }
    if ($Path -match '\.(left|top|width|height|insideWidth|insideHeight|clientWidth|clientHeight)$' -and $null -ne $Actual -and $null -ne $Expected) {
        if ([Math]::Abs([double]$Actual - [double]$Expected) -gt 0.05) {
            throw "${Path}: expected=$Expected actual=$Actual"
        }
        return
    }
    $actualJson = ConvertTo-Json -InputObject @{ value = $Actual } -Compress -Depth 10
    $expectedJson = ConvertTo-Json -InputObject @{ value = $Expected } -Compress -Depth 10
    if ($actualJson -cne $expectedJson) { throw "${Path}: expected=$expectedJson actual=$actualJson" }
}

function Get-PropertyOrNull([object]$Object, [string]$Name) {
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-ExpectedFormMap([object]$Expected) {
    $map = [ordered]@{}
    $add = {
        param([object]$Candidate, [string]$FallbackName)
        $form = Get-PropertyOrNull $Candidate 'form'
        $name = if ($null -ne $form) { [string](Get-PropertyOrNull $form 'name') } else { [string](Get-PropertyOrNull $Candidate 'name') }
        if (-not $name) { $name = $FallbackName }
        if ($name) { $map[$name] = $Candidate }
    }
    if ($Expected -is [System.Array]) {
        foreach ($candidate in $Expected) { & $add $candidate '' }
        return $map
    }
    $forms = Get-PropertyOrNull $Expected 'forms'
    if ($null -ne $forms) {
        if ($forms.PSObject.Properties.Count -gt 0 -and $forms -isnot [System.Array] -and $null -eq (Get-PropertyOrNull $forms 'form')) {
            foreach ($property in $forms.PSObject.Properties) { & $add $property.Value $property.Name }
        } else {
            foreach ($candidate in @($forms)) { & $add $candidate '' }
        }
        return $map
    }
    foreach ($name in $script:generatedFormNames) {
        $candidate = Get-PropertyOrNull $Expected $name
        if ($null -ne $candidate) { & $add $candidate $name }
    }
    if ($map.Count -eq 0) { & $add $Expected '' }
    return $map
}

function Assert-ExpectedFields([string]$Path, [object]$Actual, [object]$Expected, [string[]]$Fields) {
    foreach ($field in $Fields) {
        $expectedValue = Get-PropertyOrNull $Expected $field
        if ($null -ne $Expected.PSObject.Properties[$field]) {
            # MSForms.Image has no public TabIndex property. Its persisted site
            # TabIndex is asserted by the normalized pure-Go readback gate.
            if ($field -eq 'tabIndex' -and $Actual.type -eq 'Image' -and $null -eq $Actual.PSObject.Properties[$field]) { continue }
            Assert-Value "$Path.$field" (Get-PropertyOrNull $Actual $field) $expectedValue
        }
    }
}

function Assert-ExpectedForm([object]$Snapshot, [object]$Expected) {
    $expectedForm = Get-PropertyOrNull $Expected 'form'
    if ($null -eq $expectedForm) { $expectedForm = $Expected }
    $expectedName = Get-PropertyOrNull $expectedForm 'name'
    if ($expectedName -and [string]$expectedName -cne [string]$Snapshot.name) {
        throw "Expected form name differs: expected=$expectedName actual=$($Snapshot.name)"
    }
    Assert-ExpectedFields "$($Snapshot.name).form" $Snapshot.form $expectedForm @('caption')

    $expectedObservedForm = Get-PropertyOrNull $expectedForm 'observed'
    if ($null -ne $expectedObservedForm) {
        Assert-ExpectedFields "$($Snapshot.name).form.observed" $Snapshot.root $expectedObservedForm @('clientWidth', 'clientHeight')
    }
    $expectedBuildForm = Get-PropertyOrNull $expectedForm 'build'
    if ($null -eq $expectedBuildForm) { $expectedBuildForm = Get-PropertyOrNull $Expected 'build' }
    if ($null -eq $expectedBuildForm) { throw "$($Snapshot.name): ExpectedPath must provide form.build.clientWidth/clientHeight" }
    foreach ($dimension in @('clientWidth', 'clientHeight')) {
        if ($null -eq $expectedBuildForm.PSObject.Properties[$dimension]) {
            throw "$($Snapshot.name): ExpectedPath must provide form.build.$dimension"
        }
    }
    Assert-ExpectedFields "$($Snapshot.name).form.build" $Snapshot.form $expectedBuildForm @('caption')
    Assert-ExpectedFields "$($Snapshot.name).form.build" $Snapshot.root $expectedBuildForm @('clientWidth', 'clientHeight')

    $expectedControls = Get-PropertyOrNull $Expected 'controls'
    if ($null -eq $expectedControls) { $expectedControls = Get-PropertyOrNull $expectedForm 'controls' }
    if ($null -eq $expectedControls) { return }
    foreach ($expectedControl in @($expectedControls)) {
        $expectedName = [string](Get-PropertyOrNull $expectedControl 'name')
        if (-not $expectedName) { throw "Expected control has no name in $($Snapshot.name)" }
        $actual = @($Snapshot.controls | Where-Object { $_.name -ceq $expectedName })
        if ($actual.Count -ne 1) { throw "Missing or duplicate control $($Snapshot.name).$expectedName" }
        Assert-ExpectedFields "$($Snapshot.name).$expectedName" $actual[0] $expectedControl @('type', 'caption', 'text', 'value', 'left', 'top', 'width', 'height', 'tabIndex', 'enabled', 'visible', 'min', 'max', 'smallChange', 'largeChange', 'orientation', 'columnCount', 'boundColumn', 'listRows', 'matchEntry', 'style', 'pictureSizeMode', 'pictureAlignment', 'specialEffect', 'borderStyle', 'autoSize', 'backStyle')
        $observed = Get-PropertyOrNull $expectedControl 'observed'
        if ($null -ne $observed) {
            Assert-ExpectedFields "$($Snapshot.name).$expectedName.observed" $actual[0] $observed @('caption', 'text', 'value', 'left', 'top', 'width', 'height', 'tabIndex', 'enabled', 'visible', 'min', 'max', 'smallChange', 'largeChange', 'orientation', 'columnCount', 'boundColumn', 'listRows', 'matchEntry', 'style', 'pictureSizeMode', 'pictureAlignment', 'specialEffect', 'borderStyle', 'autoSize', 'backStyle')
            $observedProperties = Get-PropertyOrNull $observed 'properties'
            if ($null -ne $observedProperties) {
                $actualProperties = Get-PropertyOrNull $actual[0] 'properties'
                foreach ($property in $observedProperties.PSObject.Properties) {
                    Assert-Value "$($Snapshot.name).$expectedName.observed.properties.$($property.Name)" (Get-PropertyOrNull $actualProperties $property.Name) $property.Value
                }
            }
        }
        $expectedProperties = Get-PropertyOrNull $expectedControl 'properties'
        if ($null -ne $expectedProperties) {
            $actualProperties = Get-PropertyOrNull $actual[0] 'properties'
            foreach ($property in $expectedProperties.PSObject.Properties) {
                Assert-Value "$($Snapshot.name).$expectedName.properties.$($property.Name)" (Get-PropertyOrNull $actualProperties $property.Name) $property.Value
            }
        }
    }
}

function Assert-GeneratedFormShape([object]$Generated, [object]$Empty) {
    if ($Generated.name -cne 'GeneratedForm') { throw "Expected GeneratedForm, got $($Generated.name)" }
    if ($Empty.name -cne 'GeneratedEmptyForm') { throw "Expected GeneratedEmptyForm, got $($Empty.name)" }
    $expectedNames = @($script:controlSpecs | ForEach-Object { $_.name })
    if ($Generated.controls.Count -ne $expectedNames.Count) {
        throw "GeneratedForm has $($Generated.controls.Count) controls; expected $($expectedNames.Count)"
    }
    foreach ($name in $expectedNames) {
        $controlMatches = @($Generated.controls | Where-Object { $_.name -ceq $name })
        if ($controlMatches.Count -ne 1) {
            throw "GeneratedForm is missing control $name"
        }
        $expectedType = [string]($script:controlSpecs | Where-Object { $_.name -ceq $name } | Select-Object -First 1).type
        if ($controlMatches[0].type -and $controlMatches[0].type -cne $expectedType) {
            throw "GeneratedForm.$name type differs: expected=$expectedType actual=$($controlMatches[0].type)"
        }
    }
    if ($Empty.controls.Count -ne 0) { throw "EmptyForm has $($Empty.controls.Count) controls" }
}

function Publish-Fixture {
    if (-not $FixtureDirectory) { return }
    $FixtureDirectory = [IO.Path]::GetFullPath($FixtureDirectory)
    if (Test-Path -LiteralPath $FixtureDirectory) { throw "Refusing to overwrite fixture directory: $FixtureDirectory" }
    [void][IO.Directory]::CreateDirectory($FixtureDirectory)
    $artifacts = @('baseline.xlsm', 'baseline.bin', 'baseline.json', 'environment.json')
    foreach ($control in @('ToggleButtonMain', 'SpinButtonMain', 'ScrollBarMain', 'ImageMain')) {
        $artifacts += "enabled-$control-False.bin", "enabled-$control-False.json"
    }
    foreach ($name in $artifacts) {
        $source = Join-Path $WorkspacePath $name
        if (-not (Test-Path -LiteralPath $source)) { throw "Missing successful create artifact: $source" }
        $destination = Join-Path $FixtureDirectory $name
        if (Test-Path -LiteralPath $destination) { throw "Refusing to overwrite fixture artifact: $destination" }
        [IO.File]::Copy($source, $destination, $false)
    }
    Write-Output "fixture published after Excel cleanup: $FixtureDirectory"
}

function Invoke-GenerationSentinel {
    $macroName = "'$($script:workbook.Name.Replace("'", "''"))'!Main.RunGenerationSentinel"
    try {
        [void]$script:excel.Run($macroName)
    } catch {
        Write-Output "Sentinel failed; workbook-qualified macro=$macroName"
        throw
    }
    $sheets = Hold-Com $script:workbook.Worksheets
    $sheet = Hold-Com ($sheets.Item(1))
    $cell = Hold-Com ($sheet.Range('A1'))
    $value = $cell.Value2
    if ([string]$value -cne 'issue-883-ok') { throw "Sentinel differs: $value" }
    $widthCell = Hold-Com ($sheet.Range('B1'))
    $heightCell = Hold-Com ($sheet.Range('B2'))
    Write-Output "runtime observation only (not persisted client dimensions): InsideWidth=$($widthCell.Value2) InsideHeight=$($heightCell.Value2)"
    Release-Children
}

if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Excel is already running. This isolated harness requires no concurrent Excel session.'
}

if ($Phase -eq 'create') {
    if (-not $WorkspacePath) {
        $WorkspacePath = Join-Path $repoRoot ('tmp_workspaces/issue-883-generation-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
    }
    $WorkspacePath = [IO.Path]::GetFullPath($WorkspacePath)
    $allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces')) + [IO.Path]::DirectorySeparatorChar
    if (-not $WorkspacePath.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'create workspace must be below repo/tmp_workspaces' }
    if (Test-Path -LiteralPath $WorkspacePath) { throw 'create requires a fresh workspace; existing artifacts are preserved' }
    [void][IO.Directory]::CreateDirectory($WorkspacePath)
    if ($FixtureDirectory) {
        $FixtureDirectory = [IO.Path]::GetFullPath($FixtureDirectory)
        if (Test-Path -LiteralPath $FixtureDirectory) { throw 'FixtureDirectory must be new; existing evidence is never overwritten' }
    }
} else {
    if (-not $WorkbookPath -or -not (Test-Path -LiteralPath $WorkbookPath)) { throw 'verify requires an existing -WorkbookPath' }
    $WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
    if ($ExpectedPath -and -not (Test-Path -LiteralPath $ExpectedPath)) { throw 'ExpectedPath does not exist' }
    if ($NormalizedWorkbookPath) {
        $NormalizedWorkbookPath = [IO.Path]::GetFullPath($NormalizedWorkbookPath)
        if ((Test-Path -LiteralPath $NormalizedWorkbookPath) -or $NormalizedWorkbookPath -eq $WorkbookPath) { throw 'NormalizedWorkbookPath must be a new output path' }
    }
    $binaryOutput = if ($NormalizedWorkbookPath) { "$NormalizedWorkbookPath.bin" } else { "$WorkbookPath.bin" }
    if (Test-Path -LiteralPath $binaryOutput) { throw "Refusing to overwrite binary evidence: $binaryOutput" }
}

$failure = $null
$environment = $null
$baselinePath = $null
$baselineBinaryPath = $null
$phaseOutput = $null

try {
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false
    $script:excel.DisplayAlerts = $false
    $script:excel.EnableEvents = $false
    $script:excel.AutomationSecurity = if ($Phase -eq 'verify') { 1 } else { 3 }
    $owned = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($owned.Count -ne 1) { throw 'Cannot establish a unique owned Excel process' }
    $script:excelPID = $owned[0].Id
    $environment = [pscustomobject]@{
        harness = 'test-userform-generation-e2e.ps1'
        issue = 883
        phase = $Phase
        version = [string]$script:excel.Version
        build = [string]$script:excel.Build
        operatingSystem = [string]$script:excel.OperatingSystem
        processId = $script:excelPID
        trustedVbideRequired = $true
    }
    $books = Hold-Com $script:excel.Workbooks

    if ($Phase -eq 'create') {
        $script:workbook = $books.Add()
        New-Form 'GenerationBaseline' 'issue883-generation-baseline' $true
        New-Form 'EmptyForm' 'issue883-empty-form' $false
        Add-MainModule
        Release-Children

        $stage = Save-Stage 'baseline' $script:formNames
        $baselinePath = $stage.path
        $baselineBinaryPath = Join-Path $WorkspacePath 'baseline.bin'
        $sourceEvidence = @(
            Export-FormEvidence 'GenerationBaseline' (Join-Path $WorkspacePath 'GenerationBaseline.frm')
            Export-FormEvidence 'EmptyForm' (Join-Path $WorkspacePath 'EmptyForm.frm')
        )
        Release-Children
        $baselineObservation = $stage.observation | Add-Member -PassThru -NotePropertyName forms -NotePropertyValue (Get-WorkbookObservation $script:formNames).forms
        $baselineObservation | Add-Member -PassThru -NotePropertyName vbBase -NotePropertyValue $sourceEvidence
        $baselineObservation | Add-Member -PassThru -NotePropertyName binary -NotePropertyValue (Get-BinaryEvidence $baselineBinaryPath)
        Write-Json (Join-Path $WorkspacePath 'baseline.json') $baselineObservation

        foreach ($controlName in $script:newControlNames) {
            foreach ($enabled in @($false, $true)) {
                $designer = Get-Designer 'GenerationBaseline'
                $controls = Hold-Com $designer.Controls
                $control = Hold-Com ($controls.Item($controlName))
                $control.Enabled = $enabled
                Release-Children
                [void](Save-Stage "enabled-$controlName-$enabled" $script:formNames)
            }
        }
        $phaseOutput = "create complete: $WorkspacePath"
    } else {
        $script:excel.AutomationSecurity = 1
        $script:workbook = $books.Open($WorkbookPath, 0, $true)
        $generated = Get-FormSnapshot 'GeneratedForm'
        $empty = Get-FormSnapshot 'GeneratedEmptyForm'
        Assert-GeneratedFormShape $generated $empty
        Release-Children
        if ($ExpectedPath) {
            $expected = Get-Content -LiteralPath $ExpectedPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $expectedMap = Get-ExpectedFormMap $expected
            foreach ($name in $script:generatedFormNames) {
                if ($expectedMap.Contains($name)) {
                    $snapshot = if ($name -eq 'GeneratedForm') { $generated } else { $empty }
                    Assert-ExpectedForm $snapshot $expectedMap[$name]
                }
            }
        }
        Invoke-GenerationSentinel
        $sourceForBinary = $WorkbookPath
        if ($NormalizedWorkbookPath) {
            $script:workbook.SaveAs($NormalizedWorkbookPath, 52)
            $script:workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook)
            $script:workbook = $null
            $books = Hold-Com $script:excel.Workbooks
            $script:workbook = $books.Open($NormalizedWorkbookPath, 0, $true)
            $generated = Get-FormSnapshot 'GeneratedForm'
            $empty = Get-FormSnapshot 'GeneratedEmptyForm'
            Assert-GeneratedFormShape $generated $empty
            Release-Children
            if ($ExpectedPath) {
                foreach ($name in $script:generatedFormNames) {
                    if ($expectedMap.Contains($name)) {
                        $snapshot = if ($name -eq 'GeneratedForm') { $generated } else { $empty }
                        Assert-ExpectedForm $snapshot $expectedMap[$name]
                    }
                }
            }
            $sourceForBinary = $NormalizedWorkbookPath
            $phaseOutput = "verify complete: $WorkbookPath normalized=$NormalizedWorkbookPath sentinel=issue-883-ok"
        } else {
            $phaseOutput = "verify complete: $WorkbookPath sentinel=issue-883-ok"
        }
        $script:workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook)
        $script:workbook = $null
        Export-VbaProject $sourceForBinary $binaryOutput
        $phaseOutput = "$phaseOutput binary=$binaryOutput"
    }
} catch {
    $failure = $_
} finally {
    try { Release-Children } catch { $script:cleanupErrors.Add($_.Exception.Message) }
    if ($null -ne $script:workbook) {
        try { $script:workbook.Close($false) } catch { $script:cleanupErrors.Add($_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook) } catch { $script:cleanupErrors.Add($_.Exception.Message) }
        $script:workbook = $null
    }
    if ($null -ne $script:excel) {
        try { $script:excel.Quit() } catch { $script:cleanupErrors.Add($_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel) } catch { $script:cleanupErrors.Add($_.Exception.Message) }
        $script:excel = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($script:excelPID -ne 0) {
        $process = Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue
        if ($null -ne $process) {
            if (-not $process.WaitForExit(10000)) {
                $script:cleanupErrors.Add("Owned Excel PID $($script:excelPID) did not exit after Quit")
            }
        }
        if (Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue) {
            $script:cleanupErrors.Add("Owned Excel PID $($script:excelPID) is still running")
        }
    }
    if ($script:cleanupErrors.Count -ne 0) {
        $failure = [System.Exception]::new("Excel cleanup was not clean: $($script:cleanupErrors -join '; ')", $failure)
    }
}

if ($null -ne $failure) { throw $failure }

if ($Phase -eq 'create') {
    $environment | Add-Member -PassThru -NotePropertyName cleanupConfirmed -NotePropertyValue $true | Out-Null
    $environment | Add-Member -PassThru -NotePropertyName baselineWorkbook -NotePropertyValue $baselinePath | Out-Null
    $environment | Add-Member -PassThru -NotePropertyName baselineBinary -NotePropertyValue $baselineBinaryPath | Out-Null
    Write-Json (Join-Path $WorkspacePath 'environment.json') $environment
    Publish-Fixture
}

Write-Output $phaseOutput
Write-Output "cleanup confirmed: owned Excel PID=$($script:excelPID) exited"
