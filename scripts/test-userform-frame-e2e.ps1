[CmdletBinding()]
param(
    [ValidateSet('create', 'inspect', 'verify')][string]$Phase = 'create',
    [ValidateSet('generated', 'template', 'cli-blank', 'cli-template')][string]$Variant = 'generated',
    [string]$WorkspacePath = '',
    [string]$WorkbookPath = '',
    [string]$ExpectedPath = '',
    [string]$FixtureDirectory = '',
    [switch]$ObserveOnly
)

# Developer-only Excel oracle for Issue #884. Requires trusted VBIDE access.
# It creates an Excel-authored nested Frame fixture or verifies a compiler/pack
# artifact by inspecting its runtime Designer before and after save/reopen.
# Never run this harness from ordinary tests or CI.
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$authoredFixtureDirectory = Join-Path $repoRoot 'internal/vba/userforms/compiler/testdata/frame-excel-authored'
$generatedFixtureDirectory = Join-Path $repoRoot 'internal/vba/userforms/compiler/testdata/frame-excel-generated'
$utf8 = [Text.UTF8Encoding]::new($false)
$script:comRefs = [Collections.Generic.List[object]]::new()
$script:excel = $null
$script:workbook = $null
$script:excelPID = 0
$script:cleanupErrors = [Collections.Generic.List[string]]::new()
$script:knownTypes = @{
    RootCommon = 'Label'
    EmptyFrame = 'Frame'
    ParentFrame = 'Frame'
    RootText = 'TextBox'
    SiblingText = 'TextBox'
    NestedFrame = 'Frame'
    NestedLabel = 'Label'
}
$script:expectedTypes = @{}

function Hold-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $identity = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($Value)
        try {
            $known = $false
            foreach ($prior in $script:comRefs) {
                $priorIdentity = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($prior)
                try { if ($priorIdentity -eq $identity) { $known = $true } }
                finally { [void][Runtime.InteropServices.Marshal]::Release($priorIdentity) }
                if ($known) { break }
            }
            if (-not $known) { $script:comRefs.Add($Value) }
        } finally { [void][Runtime.InteropServices.Marshal]::Release($identity) }
    }
    return ,$Value
}

function Release-Children([switch]$Final) {
    if (-not $Final) {
        # Keep every RCW until final cleanup, including references obtained
        # before closing/reopening a workbook. Clearing loses release ownership.
        return
    }
    $releasedIdentities = [Collections.Generic.HashSet[IntPtr]]::new()
    for ($i = $script:comRefs.Count - 1; $i -ge 0; $i--) {
        $value = $script:comRefs[$i]
        $identity = [IntPtr]::Zero
        try {
            if (-not [Runtime.InteropServices.Marshal]::IsComObject($value)) { continue }
            $identity = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($value)
            if ($releasedIdentities.Add($identity)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) }
        } catch {
            $script:cleanupErrors.Add("COM reference release failed: $($_.Exception.Message)")
        } finally {
            if ($identity -ne [IntPtr]::Zero) { [void][Runtime.InteropServices.Marshal]::Release($identity) }
        }
    }
    $releasedIdentities.Clear()
    $script:comRefs.Clear()
}

function Write-Json([string]$Path, [object]$Value) {
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 40), $utf8)
}

function Get-ComProcessIds {
    return @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id }))
}

function Get-OwnedExcelPid([int[]]$Before) {
    if (-not ('FrameExcelProcessIdentity' -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class FrameExcelProcessIdentity {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
    }
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        $owned = [uint32]0
        [void][FrameExcelProcessIdentity]::GetWindowThreadProcessId([IntPtr]$script:excel.Hwnd, [ref]$owned)
        if ($owned -ne 0) {
            if ([int]$owned -in $Before) { throw 'Excel COM attached to a pre-existing instance' }
            return [int]$owned
        }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    throw 'Excel COM did not create a uniquely identifiable owned process'
}

function Export-VbaProject([string]$Source, [string]$Destination) {
    if (Test-Path -LiteralPath $Destination) { throw "Refusing to overwrite binary evidence: $Destination" }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression
    $stream = [IO.File]::Open($Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $entry = $archive.GetEntry('xl/vbaProject.bin')
        if ($null -eq $entry) { throw "VBA project missing from $Source" }
        $entryStream = $entry.Open()
        try {
            $output = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try { $entryStream.CopyTo($output) } finally { $output.Dispose() }
        } finally { $entryStream.Dispose() }
    } finally { $archive.Dispose(); $stream.Dispose() }
}

function Get-BinaryEvidence([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
    return [pscustomobject]@{ entry = 'xl/vbaProject.bin'; path = $Path; length = [int64]$bytes.Length; sha256 = $hash }
}

function Get-OptionalProperty([object]$Object, [string]$Name) {
    try {
        $flags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::GetProperty
        $value = $Object.GetType().InvokeMember($Name, $flags, $null, $Object, @(), [Globalization.CultureInfo]::InvariantCulture)
        if ($value -is [DBNull]) { $value = $null }
        if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { $value = '[COM object]' }
        if ($value -is [single] -or $value -is [double] -or $value -is [decimal]) { $value = [double]$value }
        if ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64]) { $value = [long]$value }
        return [pscustomobject]@{ available = $true; value = $value }
    } catch {
        return [pscustomobject]@{ available = $false }
    }
}

function Get-RequiredDouble([object]$Object, [string]$Name) {
    $property = Get-OptionalProperty $Object $Name
    if (-not $property.available -or $null -eq $property.value) { throw "Required COM property $Name is unavailable" }
    return [double]$property.value
}

function Get-Designer {
    $project = Hold-Com $script:workbook.VBProject
    $components = Hold-Com $project.VBComponents
    $component = Hold-Com ($components.Item('FrameTopologyForm'))
    return (Hold-Com $component.Designer)
}

function Get-ControlParentName([object]$Control) {
    $parent = Hold-Com $Control.Parent
    if ($null -eq $parent) { return '' }
    return [string]$parent.Name
}

function Get-ControlsByIndex([object]$Controls, [string]$ExpectedParentName) {
    $count = [int]$Controls.Count
    if ($count -eq 0) { return @() }
    $start = 0
    try { $first = Hold-Com ($Controls.Item(0)) }
    catch { $start = 1; $first = Hold-Com ($Controls.Item(1)) }
    $result = [Collections.Generic.List[object]]::new()
    if ([string]::Equals((Get-ControlParentName $first), $ExpectedParentName, [StringComparison]::OrdinalIgnoreCase)) { $result.Add($first) }
    for ($index = 1; $index -lt $count; $index++) {
        $control = Hold-Com ($Controls.Item($start + $index))
        if ([string]::Equals((Get-ControlParentName $control), $ExpectedParentName, [StringComparison]::OrdinalIgnoreCase)) { $result.Add($control) }
    }
    return $result.ToArray()
}

function Add-ControlSnapshot([object]$Control, [string]$Parent, [int]$SiblingIndex, [int]$Depth, [Collections.Generic.List[object]]$Rows, [Collections.Generic.List[object]]$Collections) {
    $progId = Get-OptionalProperty $Control 'ProgID'
    $className = Get-OptionalProperty $Control 'ClassName'
    $kind = ''
    if ($progId.available -and $progId.value -and ([string]$progId.value -match '^Forms\.(?<type>[^.]+)\.\d+$')) { $kind = $Matches.type }
    if (-not $kind -and $className.available -and $className.value) { $kind = [string]$className.value }
    if ($script:expectedTypes.ContainsKey([string]$Control.Name)) { $kind = [string]$script:expectedTypes[[string]$Control.Name] }
    if (-not $kind -and $script:knownTypes.ContainsKey([string]$Control.Name)) { $kind = $script:knownTypes[[string]$Control.Name] }
    if (-not $kind) { throw "Cannot identify MSForms class for control $($Control.Name)" }
    $progIdValue = if ($progId.available -and $progId.value) { [string]$progId.value } else { "Forms.$kind.1" }
    $properties = [ordered]@{}
    foreach ($propertyName in @('Caption', 'Text', 'Value', 'Left', 'Top', 'Width', 'Height', 'TabIndex', 'Enabled', 'Visible', 'Tag', 'ControlTipText', 'BackColor', 'ForeColor', 'BorderColor', 'BorderStyle', 'SpecialEffect', 'AutoSize', 'BackStyle')) {
        $properties[$propertyName] = (Get-OptionalProperty $Control $propertyName)
    }
    $row = [ordered]@{
        name = [string]$Control.Name
        type = $kind
        parent = $(if ($Parent) { $Parent } else { $null })
        depth = $Depth
        siblingIndex = $SiblingIndex
        progId = $progIdValue
        properties = $properties
    }
    $Rows.Add([pscustomobject]$row)
    if ($kind -eq 'Frame') {
        $children = Hold-Com $Control.Controls
        $childControls = @(Get-ControlsByIndex $children ([string]$Control.Name))
        $names = @($childControls | ForEach-Object { [string]$_.Name })
        $Collections.Add([pscustomobject]@{ parent = [string]$Control.Name; count = $childControls.Count; order = $names })
        for ($i = 0; $i -lt $childControls.Count; $i++) {
            Add-ControlSnapshot $childControls[$i] ([string]$Control.Name) $i ($Depth + 1) $Rows $Collections
        }
    }
}

function Get-Snapshot {
    $designer = Get-Designer
    $controls = Hold-Com $designer.Controls
    $rootControls = @(Get-ControlsByIndex $controls 'FrameTopologyForm')
    $rows = [Collections.Generic.List[object]]::new()
    $collections = [Collections.Generic.List[object]]::new()
    $rootNames = @($rootControls | ForEach-Object { [string]$_.Name })
    $collections.Add([pscustomobject]@{ parent = $null; count = $rootControls.Count; order = $rootNames })
    for ($i = 0; $i -lt $rootControls.Count; $i++) {
        Add-ControlSnapshot $rootControls[$i] '' $i 0 $rows $collections
    }
    $formProperties = [ordered]@{}
    foreach ($name in @('Caption', 'Width', 'Height', 'InsideWidth', 'InsideHeight')) {
        $formProperties[$name] = Get-OptionalProperty $designer $name
    }
    return [pscustomobject]@{
        form = [pscustomobject]@{ name = 'FrameTopologyForm'; properties = $formProperties }
        collections = @($collections.ToArray())
        controls = @($rows.ToArray())
    }
}

function Get-TopologySignature([object]$Snapshot) {
    $parts = foreach ($row in $Snapshot.controls) {
        $p = $row.properties
        $parent = if ($row.parent) { [string]$row.parent } else { '<root>' }
        $geometry = foreach ($name in @('Left', 'Top', 'Width', 'Height')) {
            if ($p[$name].available) { ([double]$p[$name].value).ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { '?' }
        }
        "$parent/$($row.siblingIndex):$($row.name):$($row.type):$($geometry -join ',')"
    }
    return ($parts -join ';')
}

function Assert-FrameFixture([object]$Snapshot) {
    $byName = @{}
    foreach ($row in $Snapshot.controls) { $byName[[string]$row.name] = $row }
    foreach ($name in @('RootCommon', 'EmptyFrame', 'ParentFrame', 'RootText', 'SiblingText', 'NestedFrame', 'NestedLabel')) {
        if (-not $byName.ContainsKey($name)) { throw "Expected UserForm control missing: $name" }
    }
    if ($byName.EmptyFrame.type -ne 'Frame' -or $byName.ParentFrame.type -ne 'Frame' -or $byName.NestedFrame.type -ne 'Frame') { throw 'Expected Frame control type differs' }
    if ($null -ne $byName.RootCommon.parent -or $null -ne $byName.EmptyFrame.parent -or $null -ne $byName.ParentFrame.parent -or $null -ne $byName.RootText.parent) { throw 'Root control parent metadata differs' }
    if ($byName.SiblingText.parent -cne 'ParentFrame' -or $byName.NestedFrame.parent -cne 'ParentFrame' -or $byName.NestedLabel.parent -cne 'NestedFrame') {
        $actualParents = ($Snapshot.controls | ForEach-Object { "$($_.name)<-$($_.parent)" }) -join ', '
        $actualCollections = ($Snapshot.collections | ForEach-Object { "$($_.parent):$($_.order -join ',')" }) -join '; '
        throw "Nested control parent metadata differs: $actualParents; collections=$actualCollections"
    }
    $collectionsByParent = @{}
    foreach ($collection in $Snapshot.collections) { $key = if ($collection.parent) { [string]$collection.parent } else { '<root>' }; $collectionsByParent[$key] = $collection }
    if ($collectionsByParent.EmptyFrame.count -ne 0) { throw 'EmptyFrame must have no child controls' }
    if (($collectionsByParent['<root>'].order -join ',') -cne 'RootCommon,EmptyFrame,ParentFrame,RootText') { throw 'Root control order differs' }
    if (($collectionsByParent.ParentFrame.order -join ',') -cne 'SiblingText,NestedFrame') { throw 'ParentFrame child order differs' }
    if (($collectionsByParent.NestedFrame.order -join ',') -cne 'NestedLabel') { throw 'NestedFrame child order differs' }
}

function New-AuthoredWorkbook {
    $books = Hold-Com $script:excel.Workbooks
    $script:workbook = $books.Add()
    $project = Hold-Com $script:workbook.VBProject
    $components = Hold-Com $project.VBComponents
    $formComponent = Hold-Com ($components.Add(3))
    $designer = Hold-Com $formComponent.Designer
    $formComponent.Name = 'FrameTopologyForm'
    $designer.Caption = 'issue884-frame-topology'
    $formProperties = Hold-Com $formComponent.Properties
    $formWidth = Hold-Com ($formProperties.Item('Width'))
    $formHeight = Hold-Com ($formProperties.Item('Height'))
    $formWidth.Value = 340
    $formHeight.Value = 260

    $controls = Hold-Com $designer.Controls
    $rootCommon = Hold-Com ($controls.Add('Forms.Label.1', 'RootCommon', $true))
    $rootCommon.Caption = 'root-common'
    $rootCommon.Left = 12; $rootCommon.Top = 10; $rootCommon.Width = 90; $rootCommon.Height = 18
    $emptyFrame = Hold-Com ($controls.Add('Forms.Frame.1', 'EmptyFrame', $true))
    $emptyFrame.Caption = 'empty-frame'
    $emptyFrame.Left = 12; $emptyFrame.Top = 40; $emptyFrame.Width = 100; $emptyFrame.Height = 58
    $parentFrame = Hold-Com ($controls.Add('Forms.Frame.1', 'ParentFrame', $true))
    $parentFrame.Caption = 'parent-frame'
    $parentFrame.Left = 130; $parentFrame.Top = 40; $parentFrame.Width = 180; $parentFrame.Height = 136
    $rootText = Hold-Com ($controls.Add('Forms.TextBox.1', 'RootText', $true))
    $rootText.Value = 'root-text'
    $rootText.Left = 12; $rootText.Top = 202; $rootText.Width = 180; $rootText.Height = 20

    $parentControls = Hold-Com $parentFrame.Controls
    $siblingText = Hold-Com ($parentControls.Add('Forms.TextBox.1', 'SiblingText', $true))
    $siblingText.Value = 'sibling-text'
    $siblingText.Left = 10; $siblingText.Top = 22; $siblingText.Width = 84; $siblingText.Height = 18
    $nestedFrame = Hold-Com ($parentControls.Add('Forms.Frame.1', 'NestedFrame', $true))
    $nestedFrame.Caption = 'nested-frame'
    $nestedFrame.Left = 12; $nestedFrame.Top = 54; $nestedFrame.Width = 144; $nestedFrame.Height = 64
    $nestedControls = Hold-Com $nestedFrame.Controls
    $nestedLabel = Hold-Com ($nestedControls.Add('Forms.Label.1', 'NestedLabel', $true))
    $nestedLabel.Caption = 'nested-label'
    $nestedLabel.Left = 8; $nestedLabel.Top = 20; $nestedLabel.Width = 108; $nestedLabel.Height = 18

    $module = Hold-Com ($components.Add(1))
    $module.Name = 'Main'
    $code = Hold-Com $module.CodeModule
    $code.AddFromString(@'
Option Explicit

Private Function DumpFrameControls(ByVal controls As Object, ByVal parentPath As String) As String
    Dim control As Object
    Dim index As Long
    Dim row As String
    For Each control In controls
        row = parentPath & CStr(index) & ":" & CStr(control.Name) & ":" & TypeName(control) & ":" & _
            CStr(control.Left) & "," & CStr(control.Top) & "," & CStr(control.Width) & "," & CStr(control.Height)
        DumpFrameControls = DumpFrameControls & row & ";"
        If TypeName(control) = "Frame" Then
            DumpFrameControls = DumpFrameControls & DumpFrameControls(control.Controls, parentPath & CStr(control.Name) & "/")
        End If
        index = index + 1
    Next control
End Function

Public Sub RunFrameSentinel()
    On Error GoTo Failed
    Dim form As FrameTopologyForm
    Set form = New FrameTopologyForm
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-884-ok|" & DumpFrameControls(form.Controls, "")
    Unload form
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-884-failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
'@)
    Release-Children
}

function Get-ExcelEnvironment([string]$Phase) {
    return [pscustomobject]@{
        harness = 'scripts/test-userform-frame-e2e.ps1'
        issue = 884
        phase = $Phase
        version = [string]$script:excel.Version
        build = [string]$script:excel.Build
        operatingSystem = [string]$script:excel.OperatingSystem
        processId = $script:excelPID
        processArchitecture = [Environment]::Is64BitProcess.ToString()
        powershell = [string]$PSVersionTable.PSVersion
        trustedVbideRequired = $true
    }
}

function Get-ExpectedRuntimeSignature([object]$Expected) {
    $byId = @{}
    foreach ($control in $Expected.controls) { $byId[[string]$control.id] = $control }
    $segments = [Collections.Generic.List[string]]::new()
    function Add-RuntimeSignatureChildren([string]$ParentName, [string]$ParentId) {
        $siblings = @($Expected.controls | Where-Object { ([string]$_.parentId) -ceq $ParentId } | Sort-Object { [int]$_.zIndex })
        $index = 0
        foreach ($item in $siblings) {
            $segments.Add("$ParentName/$index`:$($item.name):$($item.type);")
            if ($item.type -ceq 'Frame') { Add-RuntimeSignatureChildren ([string]$item.name) ([string]$item.id) }
            $index++
        }
    }
    Add-RuntimeSignatureChildren ([string]$Expected.form.name) ''
    return ($segments -join '')
}

function Invoke-FrameSentinel([string]$ExpectedSignature = '') {
    $escapedWorkbookName = ([string]$script:workbook.Name).Replace("'", "''")
    $macro = "'$escapedWorkbookName'!Main.RunFrameSentinel"
    try { [void]$script:excel.Run($macro) }
    catch { throw "Application.Run failed: macro=[$macro] AutomationSecurity=$($script:excel.AutomationSecurity): $($_.Exception.Message)" }
    $sheets = Hold-Com $script:workbook.Worksheets
    $sheet = Hold-Com ($sheets.Item(1))
    $cell = Hold-Com ($sheet.Range('A1'))
    $value = [string]$cell.Value2
    if (-not $value.StartsWith('issue-884-ok|', [StringComparison]::Ordinal)) { throw "Frame runtime sentinel failed: $value" }
    if ($ExpectedSignature) {
        $actualEdges = @(([regex]::Replace($value.Substring('issue-884-ok|'.Length), '/\d+:', '/:') -split ';') | Sort-Object)
        $expectedEdges = @(([regex]::Replace($ExpectedSignature.Substring('issue-884-ok|'.Length), '/\d+:', '/:') -split ';') | Sort-Object)
        if (($actualEdges -join ';') -cne ($expectedEdges -join ';')) { throw "Frame runtime hierarchy differs: expected=[$ExpectedSignature] actual=[$value]" }
    }
    Release-Children
    return $value
}

function Assert-ExpectedSnapshot([object]$Snapshot, [object]$Expected) {
    if ($Expected.kind -cne 'xlflow.userform' -or $Expected.basis -cne 'designer') { throw 'ExpectedPath must contain a canonical designer FormSpec' }
    if ($Snapshot.form.name -cne $Expected.form.name) { throw "Form identity differs: expected=$($Expected.form.name) actual=$($Snapshot.form.name)" }
    $expectedById = @{}
    $expectedByName = @{}
    foreach ($item in $Expected.controls) {
        $expectedById[[string]$item.id] = $item
        $expectedByName[[string]$item.name] = $item
        $script:expectedTypes[[string]$item.name] = [string]$item.type
    }
    if ($Snapshot.controls.Count -ne $Expected.controls.Count) { throw "Control count differs: expected=$($Expected.controls.Count) actual=$($Snapshot.controls.Count)" }
    $actualByName = @{}
    foreach ($actual in $Snapshot.controls) { $actualByName[[string]$actual.name] = $actual }
    foreach ($item in $Expected.controls) {
        $name = [string]$item.name
        if (-not $actualByName.ContainsKey($name)) { throw "Expected control missing: $name" }
        $actual = $actualByName[$name]
        if ($actual.type -cne [string]$item.type) { throw "${name}.type differs: expected=$($item.type) actual=$($actual.type)" }
        if ($actual.progId -cne [string]$item.progId) { throw "${name}.progId differs: expected=$($item.progId) actual=$($actual.progId)" }
        $expectedParent = ''
        if ($item.parentId) {
            if (-not $expectedById.ContainsKey([string]$item.parentId)) { throw "${name}.parentId is dangling: $($item.parentId)" }
            $expectedParent = [string]$expectedById[[string]$item.parentId].name
        }
        $actualParent = if ($actual.parent) { [string]$actual.parent } else { '' }
        if ($actualParent -cne $expectedParent) { throw "${name}.parent differs: expected=[$expectedParent] actual=[$actualParent]" }
        $properties = $actual.properties
        foreach ($field in @('caption', 'text', 'value', 'left', 'top', 'width', 'height', 'tabIndex', 'enabled', 'visible')) {
            $expectedProperty = $item.PSObject.Properties[$field]
            if ($null -eq $expectedProperty -or $null -eq $expectedProperty.Value) { continue }
            $propertyName = switch ($field) { 'caption' { 'Caption' } 'text' { 'Text' } 'value' { 'Value' } 'tabIndex' { 'TabIndex' } default { $field.Substring(0,1).ToUpperInvariant() + $field.Substring(1) } }
            $observed = $properties[$propertyName]
            if (-not $observed.available -and $field -eq 'text') { $observed = $properties.Value }
            if (-not $observed.available) { throw "${name}.${field} was not observable through Excel COM" }
            if ($field -in @('left', 'top', 'width', 'height')) {
                if ([Math]::Abs([double]$observed.value - [double]$expectedProperty.Value) -gt 0.05) { throw "${name}.${field} differs: expected=$($expectedProperty.Value) actual=$($observed.value)" }
            } elseif ($field -in @('enabled', 'visible')) {
                if ([bool]$observed.value -ne [bool]$expectedProperty.Value) { throw "${name}.${field} differs" }
            } elseif ($field -eq 'tabIndex') {
                if ([int]$observed.value -ne [int]$expectedProperty.Value) { throw "${name}.tabIndex differs" }
            } elseif ([string]$observed.value -cne [string]$expectedProperty.Value) { throw "${name}.${field} differs: expected=[$($expectedProperty.Value)] actual=[$($observed.value)]" }
        }
    }
    $expectedParents = @('') + @($Expected.controls | Where-Object parentId | ForEach-Object { [string]$expectedById[[string]$_.parentId].name } | Sort-Object -Unique)
    foreach ($parentName in $expectedParents) {
        $expectedNames = @($Expected.controls | Where-Object { if ($parentName) { [string]$expectedById[[string]$_.parentId].name -ceq $parentName } else { -not $_.parentId } } | Sort-Object { [int]$_.zIndex } | ForEach-Object { [string]$_.name })
        $collection = @($Snapshot.collections | Where-Object { ([string]$_.parent) -ceq $parentName })
        if ($collection.Count -ne 1) { throw "COM child collection for parent [$parentName] missing or duplicated" }
        # Controls enumeration follows host collection identity rather than
        # persisted z-order. Verify membership here; the pure-Go normalized
        # fixture test checks saved site order independently.
        if ((@($collection[0].order | Sort-Object) -join "`n") -cne (@($expectedNames | Sort-Object) -join "`n")) { throw "Child membership differs for parent [$parentName]" }
    }
}

function Assert-ObservedSnapshot([object]$Actual, [object]$Expected) {
    if (($Actual.form | ConvertTo-Json -Depth 20 -Compress) -cne ($Expected.form | ConvertTo-Json -Depth 20 -Compress)) { throw 'Saved/reopened form COM properties differ from the pre-save observation' }
    if (($Actual.controls | ConvertTo-Json -Depth 20 -Compress) -cne ($Expected.controls | ConvertTo-Json -Depth 20 -Compress)) { throw 'Saved/reopened control COM properties differ from the pre-save observation' }
    $actualCollections = @($Actual.collections | Select-Object parent, order)
    $expectedCollections = @($Expected.collections | Select-Object parent, order)
    if (($actualCollections | ConvertTo-Json -Depth 20 -Compress) -cne ($expectedCollections | ConvertTo-Json -Depth 20 -Compress)) { throw 'Saved/reopened immediate child order differs from the pre-save observation' }
}

if (@(Get-ComProcessIds).Count -ne 0) { throw 'Excel is already running; refusing to interfere with another Excel or VBE oracle session' }

if ($Phase -eq 'create') {
    if (-not $WorkspacePath) {
        $WorkspacePath = Join-Path $repoRoot 'tmp_workspaces/issue-884-frame-e2e'
        if (Test-Path -LiteralPath $WorkspacePath) {
            $WorkspacePath = Join-Path $repoRoot ('tmp_workspaces/issue-884-frame-e2e-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
        }
    }
    $WorkspacePath = [IO.Path]::GetFullPath($WorkspacePath)
    $tmpRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces')) + [IO.Path]::DirectorySeparatorChar
    if (-not $WorkspacePath.StartsWith($tmpRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Create workspace must be below the repository tmp_workspaces directory' }
    if (Test-Path -LiteralPath $WorkspacePath) { throw "Workspace already exists; preserving it: $WorkspacePath" }
    if (-not $FixtureDirectory) { $FixtureDirectory = $authoredFixtureDirectory }
    $FixtureDirectory = [IO.Path]::GetFullPath($FixtureDirectory)
    if ($FixtureDirectory -cne [IO.Path]::GetFullPath($authoredFixtureDirectory)) { throw 'Create can write only to frame-excel-authored' }
    if (Test-Path -LiteralPath $FixtureDirectory) { throw "Fixture directory already exists; preserving it: $FixtureDirectory" }
    [void][IO.Directory]::CreateDirectory($WorkspacePath)
} elseif ($Phase -eq 'inspect') {
    if (-not $WorkbookPath -or -not (Test-Path -LiteralPath $WorkbookPath)) { throw 'inspect requires an existing -WorkbookPath' }
    if (-not $ExpectedPath -or -not (Test-Path -LiteralPath $ExpectedPath)) { throw 'inspect requires a pre-save snapshot -ExpectedPath' }
    if (-not $WorkspacePath) { throw 'inspect requires a fresh -WorkspacePath below tmp_workspaces' }
    $WorkspacePath = [IO.Path]::GetFullPath($WorkspacePath)
    $tmpRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces')) + [IO.Path]::DirectorySeparatorChar
    if (-not $WorkspacePath.StartsWith($tmpRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Inspect workspace must be below the repository tmp_workspaces directory' }
    if (Test-Path -LiteralPath $WorkspacePath) { throw "Workspace already exists; preserving it: $WorkspacePath" }
    [void][IO.Directory]::CreateDirectory($WorkspacePath)
    if (-not $FixtureDirectory) { $FixtureDirectory = $authoredFixtureDirectory }
    $FixtureDirectory = [IO.Path]::GetFullPath($FixtureDirectory)
    if ($FixtureDirectory -cne [IO.Path]::GetFullPath($authoredFixtureDirectory)) { throw 'Inspect can write only to frame-excel-authored' }
    if (Test-Path -LiteralPath $FixtureDirectory) { throw "Fixture directory already exists; preserving it: $FixtureDirectory" }
    $WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
    $ExpectedPath = [IO.Path]::GetFullPath($ExpectedPath)
} else {
    if (-not $WorkbookPath -or -not (Test-Path -LiteralPath $WorkbookPath)) { throw 'verify requires an existing -WorkbookPath' }
    if (-not $ExpectedPath -or -not (Test-Path -LiteralPath $ExpectedPath)) { throw 'verify requires an existing -ExpectedPath' }
    if (-not $WorkspacePath) { throw 'verify requires a fresh -WorkspacePath below tmp_workspaces' }
    $WorkspacePath = [IO.Path]::GetFullPath($WorkspacePath)
    $tmpRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces')) + [IO.Path]::DirectorySeparatorChar
    if (-not $WorkspacePath.StartsWith($tmpRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Verify workspace must be below the repository tmp_workspaces directory' }
    if (Test-Path -LiteralPath $WorkspacePath) { throw "Workspace already exists; preserving it: $WorkspacePath" }
    [void][IO.Directory]::CreateDirectory($WorkspacePath)
    if (-not $FixtureDirectory) { $FixtureDirectory = $generatedFixtureDirectory }
    $FixtureDirectory = [IO.Path]::GetFullPath($FixtureDirectory)
    if ($FixtureDirectory -cne [IO.Path]::GetFullPath($generatedFixtureDirectory)) { throw 'Verify can write only to frame-excel-generated' }
    $variantBase = switch ($Variant) { 'generated' { 'generated' } 'template' { 'template-edited' } 'cli-blank' { 'cli-blank' } 'cli-template' { 'cli-template' } }
    foreach ($leaf in @("$variantBase.bin", "$variantBase.json", "$variantBase-environment.json", "$variantBase-expected.json")) {
        if ($ObserveOnly) { continue }
        if (Test-Path -LiteralPath (Join-Path $FixtureDirectory $leaf)) { throw "Generated evidence already exists; preserving it: $(Join-Path $FixtureDirectory $leaf)" }
    }
    $WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
    $ExpectedPath = [IO.Path]::GetFullPath($ExpectedPath)
}

$beforePids = @(Get-ComProcessIds)
try {
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false
    $script:excel.DisplayAlerts = $false
    $script:excel.EnableEvents = $false
    $script:excel.AutomationSecurity = if ($Phase -eq 'verify') { 1 } else { 3 }
    $script:excelPID = Get-OwnedExcelPid $beforePids
    $environment = Get-ExcelEnvironment $Phase
    if ($Phase -eq 'create') {
        New-AuthoredWorkbook
        $baselinePath = Join-Path $WorkspacePath 'baseline.xlsm'
        $before = Get-Snapshot
        Write-Json (Join-Path $WorkspacePath 'observed-before-save.json') $before
        Assert-FrameFixture $before
        $signatureBefore = Get-TopologySignature $before
        Release-Children
        $script:workbook.SaveAs($baselinePath, 52)
        $script:workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook)
        $script:workbook = $null
        $books = Hold-Com $script:excel.Workbooks
        $script:workbook = $books.Open($baselinePath, 0, $false)
        $reopened = Get-Snapshot
        Assert-FrameFixture $reopened
        $signatureAfter = Get-TopologySignature $reopened
        if ($signatureBefore -cne $signatureAfter) { throw 'Excel-authored topology changed after save/reopen' }
        Release-Children
        $observation = [pscustomobject]@{
            issue = 884
            stage = 'excel-authored-frame-topology'
            workbook = $baselinePath
            beforeSave = $before
            reopened = $reopened
            topologySignature = $signatureAfter
            runtimeSentinel = 'not-run for Excel-authored source fixture'
            environment = $environment
        }
        $script:pendingPublish = [pscustomobject]@{ FixtureDirectory = $FixtureDirectory; WorkspacePath = $WorkspacePath; Source = $baselinePath; BinaryPath = (Join-Path $WorkspacePath 'baseline.bin'); JsonName = 'baseline.json'; Files = @('baseline.xlsm', 'baseline.bin', 'baseline.json'); Observation = $observation; Environment = $environment }
    } elseif ($Phase -eq 'inspect') {
        $expected = Get-Content -LiteralPath $ExpectedPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $books = Hold-Com $script:excel.Workbooks
        $script:workbook = $books.Open($WorkbookPath, 0, $true)
        $reopened = Get-Snapshot
        Write-Json (Join-Path $WorkspacePath 'observed-reopened.json') $reopened
        Assert-ObservedSnapshot $reopened $expected
        $signature = Get-TopologySignature $reopened
        Release-Children
        $script:workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook)
        $script:workbook = $null
        $observation = [pscustomobject]@{
            issue = 884
            stage = 'excel-authored-frame-topology-reopened-inspection'
            workbook = $WorkbookPath
            expectedPreSaveSnapshot = $ExpectedPath
            beforeSave = $expected
            reopened = $reopened
            topologySignature = $signature
            runtimeSentinel = 'not-run during source fixture inspection'
            environment = $environment
        }
        $script:pendingPublish = [pscustomobject]@{ FixtureDirectory = $FixtureDirectory; WorkspacePath = $WorkspacePath; Source = $WorkbookPath; WorkbookSource = $WorkbookPath; BinaryPath = (Join-Path $WorkspacePath 'baseline.bin'); JsonName = 'baseline.json'; Files = @('baseline.bin', 'baseline.json'); Observation = $observation; Environment = $environment }
    } else {
        $expected = Get-Content -LiteralPath $ExpectedPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $script:expectedTypes = @{}
        foreach ($item in $expected.controls) { $script:expectedTypes[[string]$item.name] = [string]$item.type }
        $books = Hold-Com $script:excel.Workbooks
        $script:workbook = $books.Open($WorkbookPath, 0, $false)
        $before = Get-Snapshot
        Write-Json (Join-Path $WorkspacePath 'observed-before-verify.json') $before
        Assert-ExpectedSnapshot $before $expected
        $expectedRuntime = 'issue-884-ok|' + (Get-ExpectedRuntimeSignature $expected)
        $sentinelBefore = Invoke-FrameSentinel $expectedRuntime
        $signatureBefore = Get-TopologySignature $before
        $script:workbook.Save()
        $script:workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook)
        $script:workbook = $null
        $books = Hold-Com $script:excel.Workbooks
        $script:workbook = $books.Open($WorkbookPath, 0, $false)
        $reopened = Get-Snapshot
        Assert-ExpectedSnapshot $reopened $expected
        $signatureAfter = Get-TopologySignature $reopened
        if ($signatureBefore -cne $signatureAfter) { throw 'Runtime Designer topology changed after save/reopen' }
        $sentinelAfter = Invoke-FrameSentinel $sentinelBefore
        $script:workbook.Save()
        $script:workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook)
        $script:workbook = $null
        $observation = [pscustomobject]@{
            issue = 884
            stage = 'excel-verified-generated-frame-topology'
            workbook = $WorkbookPath
            expected = $ExpectedPath
            beforeSave = $before
            reopened = $reopened
            topologySignature = $signatureAfter
            runtimeSentinelBeforeSave = $sentinelBefore
            runtimeSentinelReopened = $sentinelAfter
            environment = $environment
        }
        $script:pendingPublish = [pscustomobject]@{ FixtureDirectory = $FixtureDirectory; WorkspacePath = $WorkspacePath; Source = $WorkbookPath; BinaryPath = (Join-Path $WorkspacePath "$variantBase.bin"); JsonName = "$variantBase.json"; EnvironmentName = "$variantBase-environment.json"; Files = @("$variantBase.bin", "$variantBase.json"); Observation = $observation; Environment = $environment }
    }
} catch {
    $failure = $_
} finally {
    try { Release-Children -Final } catch { $script:cleanupErrors.Add($_.Exception.Message) }
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
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($script:excelPID -ne 0) {
        $process = Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue
        if ($null -ne $process -and -not $process.WaitForExit(45000)) {
            $script:cleanupErrors.Add("Owned Excel PID $script:excelPID did not exit after Quit")
        }
        if (Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue) { $script:cleanupErrors.Add("Owned Excel PID $script:excelPID is still running") }
    }
}

if ($script:cleanupErrors.Count -ne 0) {
    $cleanupFailure = "Excel cleanup was not confirmed: $($script:cleanupErrors -join '; ')"
    if ($null -ne $failure) { throw "$($failure.Exception.Message); $cleanupFailure" }
    throw $cleanupFailure
}
if ($null -ne $failure) { throw $failure }
if ($null -ne $script:pendingPublish) {
    $publish = $script:pendingPublish
    $environment = $publish.Environment
    Export-VbaProject $publish.Source $publish.BinaryPath
    $observation = $publish.Observation
    $observation | Add-Member -NotePropertyName binary -NotePropertyValue (Get-BinaryEvidence $publish.BinaryPath)
    Write-Json (Join-Path $publish.WorkspacePath $publish.JsonName) $observation
    $environment | Add-Member -NotePropertyName cleanupConfirmed -NotePropertyValue $true
    $environment | Add-Member -NotePropertyName binarySha256 -NotePropertyValue $observation.binary.sha256
    $environment | Add-Member -NotePropertyName ownedExcelPidExited -NotePropertyValue $true
    $environmentName = if ($publish.EnvironmentName) { $publish.EnvironmentName } else { 'environment.json' }
    $environmentPath = Join-Path $publish.WorkspacePath $environmentName
    Write-Json $environmentPath $environment
    if ($Phase -eq 'verify') {
        $expectedName = "$variantBase-expected.json"
        Write-Json (Join-Path $publish.WorkspacePath $expectedName) $expected
        $publish.Files += $expectedName
    }
    if (-not $ObserveOnly) {
    [void][IO.Directory]::CreateDirectory($publish.FixtureDirectory)
    if ($publish.WorkbookSource) { [IO.File]::Copy($publish.WorkbookSource, (Join-Path $publish.FixtureDirectory 'baseline.xlsm'), $false) }
    foreach ($file in $publish.Files) {
        [IO.File]::Copy((Join-Path $publish.WorkspacePath $file), (Join-Path $publish.FixtureDirectory $file), $false)
    }
    [IO.File]::Copy($environmentPath, (Join-Path $publish.FixtureDirectory $environmentName), $false)
    }
    if ($Phase -eq 'create' -or $Phase -eq 'inspect') {
        Write-Output "create complete: workspace=$WorkspacePath fixture=$($publish.FixtureDirectory) Excel=$($environment.version) build=$($environment.build) PID=$script:excelPID binaryBytes=$($observation.binary.length) sha256=$($observation.binary.sha256)"
        Write-Output "topology: $($observation.topologySignature)"
    } else {
        Write-Output "verify complete: workbook=$WorkbookPath workspace=$WorkspacePath fixture=$($publish.FixtureDirectory) Excel=$($environment.version) build=$($environment.build) PID=$script:excelPID sha256=$($observation.binary.sha256)"
        Write-Output "topology: $($observation.topologySignature)"
        Write-Output "sentinel reopened: $($observation.runtimeSentinelReopened)"
    }
}
Write-Output "cleanup confirmed: owned Excel PID=$script:excelPID exited"
