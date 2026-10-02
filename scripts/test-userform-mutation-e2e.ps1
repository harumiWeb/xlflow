[CmdletBinding()]
param(
    [ValidateSet('create', 'verify')][string]$Phase = 'create',
    [string]$WorkspacePath = '',
    [string]$WorkbookPath = '',
    [string]$ExpectedPath = '',
    [string]$NormalizedWorkbookPath = '',
    [string]$FixtureDirectory = ''
)

# Local developer harness. Never run from ordinary tests/CI; requires trusted VBIDE.
# create retains every workbook and before/reopened observation in a fresh directory.
# verify reads a compiler-produced workbook and optionally compares a snapshot JSON.
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$utf8 = [Text.UTF8Encoding]::new($false)
$script:comRefs = [Collections.Generic.List[object]]::new()
$excel = $null
$workbook = $null
$excelPID = 0
$cleanupErrors = [Collections.Generic.List[string]]::new()

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
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 15), $utf8)
}

function Export-VbaProject([string]$Source, [string]$Destination) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Source)
    try {
        $entry = $archive.GetEntry('xl/vbaProject.bin')
        if ($null -eq $entry) { throw "VBA project missing from $Source" }
        $inputStream = $entry.Open()
        $outputStream = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew)
        try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose(); $inputStream.Dispose() }
    } finally { $archive.Dispose() }
}

function Get-Designer {
    $project = Hold-Com $workbook.VBProject
    $components = Hold-Com $project.VBComponents
    $component = Hold-Com ($components.Item('MutationForm'))
    return (Hold-Com $component.Designer)
}

function Get-Control([object]$Designer, [string]$Name) {
    $controls = Hold-Com $Designer.Controls
    if ($Name -eq 'NestedText') {
        $frame = Hold-Com ($controls.Item('FrameMain'))
        $children = Hold-Com $frame.Controls
        return (Hold-Com ($children.Item($Name)))
    }
    return (Hold-Com ($controls.Item($Name)))
}

function Get-Snapshot {
    $designer = Get-Designer
    $rows = @()
    foreach ($name in @('LabelMain', 'TextMain', 'ButtonMain', 'CheckMain', 'OptionMain', 'ComboMain', 'ListMain', 'FrameMain', 'NestedText')) {
        $control = Get-Control $designer $name
        $typeNames = @{ LabelMain = 'Label'; TextMain = 'TextBox'; ButtonMain = 'CommandButton'; CheckMain = 'CheckBox'; OptionMain = 'OptionButton'; ComboMain = 'ComboBox'; ListMain = 'ListBox'; FrameMain = 'Frame'; NestedText = 'TextBox' }
        $row = [ordered]@{
            name = $name; enabled = [bool]$control.Enabled; visible = [bool]$control.Visible
            type = $typeNames[$name]
            left = [double]$control.Left; top = [double]$control.Top
            width = [double]$control.Width; height = [double]$control.Height
            tabIndex = [int]$control.TabIndex
        }
        if ($name -in @('LabelMain', 'ButtonMain', 'CheckMain', 'OptionMain', 'FrameMain')) {
            $row.caption = [string]$control.Caption
        }
        if ($name -in @('TextMain', 'NestedText', 'ComboMain', 'ListMain', 'CheckMain', 'OptionMain')) {
            $row.value = $control.Value
        }
        if ($name -in @('TextMain', 'NestedText')) { $row.text = [string]$control.Value }
        if ($name -in @('ComboMain', 'ListMain')) {
            $items = @()
            for ($i = 0; $i -lt $control.ListCount; $i++) { $items += [string]$control.List($i, 0) }
            $row.list = @($items)
            $row.listCount = [int]$control.ListCount
            $row.selectedIndex = [int]$control.ListIndex
            $row.rowSource = [string]$control.RowSource
        }
        $rows += [pscustomobject]$row
    }
    return [pscustomobject]@{ caption = [string]$designer.Caption; controls = $rows }
}

function Assert-Value([string]$Path, [object]$Actual, [object]$Expected) {
    if ($Path -match '\.(left|top|width|height)$') {
        # Excel exposes twips; binary HIMETRIC rounding can differ by < one twip.
        if ([Math]::Abs([double]$Actual - [double]$Expected) -gt 0.05) { throw "${Path}: expected=$Expected actual=$Actual" }
        return
    }
    if ($Path -match '\.(backColor|foreColor|borderColor)$') {
        if (([long]$Actual -band 0xffffffffL) -ne ([long]$Expected -band 0xffffffffL)) { throw "${Path}: color differs" }
        return
    }
    if ($Actual -is [bool] -and $Expected -is [string] -and $Expected -match '^(True|False)$') {
        $Expected = [bool]::Parse($Expected)
    }
    $actualJson = ConvertTo-Json -InputObject @{ value = $Actual } -Compress -Depth 5
    $expectedJson = ConvertTo-Json -InputObject @{ value = $Expected } -Compress -Depth 5
    if ($actualJson -cne $expectedJson) { throw "${Path}: expected=$expectedJson actual=$actualJson" }
}

function Assert-ExpectedSnapshot([object]$Snapshot, [object]$Expected) {
    if ($Expected.PSObject.Properties.Name -contains 'reopened') { $Expected = $Expected.reopened }
    if ($Expected.PSObject.Properties.Name -contains 'form') {
        if ($Expected.form.name -cne 'MutationForm') { throw 'Expected form name differs from MutationForm' }
        if ($Expected.form.PSObject.Properties.Name -contains 'caption') { Assert-Value 'form.caption' $Snapshot.caption $Expected.form.caption }
        if ($Expected.form.build -and $Expected.form.build.PSObject.Properties.Name -contains 'caption') { Assert-Value 'form.build.caption' $Snapshot.caption $Expected.form.build.caption }
    } elseif ($Expected.PSObject.Properties.Name -contains 'caption') { Assert-Value 'form.caption' $Snapshot.caption $Expected.caption }
    $designer = Get-Designer
    foreach ($row in $Expected.controls) {
        $actual = @($Snapshot.controls | Where-Object name -eq $row.name)
        if ($actual.Count -ne 1) { throw "Missing control $($row.name)" }
        foreach ($field in $row.PSObject.Properties) {
            if ($field.Name -in @('name', 'type', 'caption', 'text', 'value', 'left', 'top', 'width', 'height', 'tabIndex', 'enabled', 'visible', 'list', 'listCount', 'selectedIndex', 'rowSource')) {
                Assert-Value "$($row.name).$($field.Name)" $actual[0].($field.Name) $field.Value
            } elseif ($field.Name -eq 'properties') {
                $control = Get-Control $designer $row.name
                foreach ($property in $field.Value.PSObject.Properties) {
                    Assert-Value "$($row.name).$($property.Name)" $control.($property.Name) $property.Value
                }
            }
        }
    }
    Release-Children
}

function Save-Stage([string]$Name) {
    $path = Join-Path $WorkspacePath "$Name.xlsm"
    $before = Get-Snapshot
    Release-Children
    $workbook.SaveAs($path, 52)
    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    $script:workbook = $null
    $books = Hold-Com $excel.Workbooks
    $script:workbook = $books.Open($path, 0, $false)
    $after = Get-Snapshot
    Release-Children
    $observation = [pscustomobject]@{ stage = $Name; workbook = $path; beforeSave = $before; reopened = $after }
    Write-Json (Join-Path $WorkspacePath "$Name.json") $observation
    Write-Output "stage=$Name saved-and-reopened=$path"
    return
}

if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Excel is already running. This isolated harness requires no concurrent Excel/oracle session.'
}
if ($Phase -eq 'create') {
    if (-not $WorkspacePath) {
        $WorkspacePath = Join-Path $repoRoot ('tmp_workspaces/issue-882-observe-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
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
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = $(if ($Phase -eq 'verify') { 1 } else { 3 })
    # There were no Excel processes before creation: only this process is owned.
    $owned = @(Get-Process EXCEL)
    if ($owned.Count -ne 1) { throw 'Cannot establish unique owned Excel process' }
    $excelPID = $owned[0].Id
    $version = [pscustomobject]@{ version = [string]$excel.Version; build = [string]$excel.Build; operatingSystem = [string]$excel.OperatingSystem; processId = $excelPID }
    $books = Hold-Com $excel.Workbooks
    if ($Phase -eq 'create') {
        Write-Json (Join-Path $WorkspacePath 'environment.json') $version
        $workbook = $books.Add()
        $project = Hold-Com $workbook.VBProject
        $components = Hold-Com $project.VBComponents
        $form = Hold-Com ($components.Add(3))
        # Access Designer before renaming; some Excel versions reject earlier rename.
        $designer = Hold-Com $form.Designer
        $form.Name = 'MutationForm'
        $designer.Caption = 'issue882-before'
        $properties = Hold-Com $form.Properties
        $widthProperty = Hold-Com ($properties.Item('Width'))
        $heightProperty = Hold-Com ($properties.Item('Height'))
        $widthProperty.Value = 360
        $heightProperty.Value = 300
        $mainModule = Hold-Com ($components.Add(1))
        $mainModule.Name = 'Main'
        $code = Hold-Com $mainModule.CodeModule
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString(@'
Option Explicit
Public Sub RunMutationSentinel()
    On Error GoTo Failed
    Dim form As MutationForm
    Set form = New MutationForm
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-882-ok"
    Unload form
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-882-failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
'@)
        $controls = Hold-Com $designer.Controls
        $types = @('Label', 'TextBox', 'CommandButton', 'CheckBox', 'OptionButton', 'ComboBox', 'ListBox', 'Frame')
        $names = @('LabelMain', 'TextMain', 'ButtonMain', 'CheckMain', 'OptionMain', 'ComboMain', 'ListMain', 'FrameMain')
        for ($i = 0; $i -lt $types.Count; $i++) {
            $control = Hold-Com ($controls.Add("Forms.$($types[$i]).1", $names[$i], $true))
            $control.Left = 12 + ($i % 2) * 160
            $control.Top = 12 + [Math]::Floor($i / 2) * 54
            $control.Width = 144
            $control.Height = 36
            if ($types[$i] -in @('Label', 'CommandButton', 'CheckBox', 'OptionButton', 'Frame')) { $control.Caption = "before-$($types[$i])" }
            if ($types[$i] -eq 'TextBox') { $control.Value = 'before-TextBox' }
            if ($types[$i] -eq 'Frame') {
                $children = Hold-Com $control.Controls
                $child = Hold-Com ($children.Add('Forms.TextBox.1', 'NestedText', $true))
                $child.Left = 6; $child.Top = 12; $child.Width = 96; $child.Height = 18
                $child.Value = 'before-NestedText'
            }
        }
        Release-Children
        Save-Stage '00-baseline'
        foreach ($name in @('ComboMain', 'ListMain')) {
            $designer = Get-Designer
            $control = Get-Control $designer $name
            $control.AddItem('alpha'); $control.AddItem('beta'); $control.AddItem('gamma')
            Release-Children
            Save-Stage "01-items-$name"
            # Recreate items after reopening, then select, to observe live/persisted boundaries.
            $designer = Get-Designer
            $control = Get-Control $designer $name
            if ($control.ListCount -eq 0) { $control.AddItem('alpha'); $control.AddItem('beta'); $control.AddItem('gamma') }
            $control.ListIndex = 1
            Release-Children
            Save-Stage "02-selection-$name"
        }
        foreach ($name in $names + @('NestedText')) {
            foreach ($enabled in @($false, $true)) {
                $designer = Get-Designer
                $control = Get-Control $designer $name
                $control.Enabled = $enabled
                Release-Children
                Save-Stage "03-enabled-$name-$enabled"
            }
        }
        foreach ($name in $names + @('NestedText')) {
            $designer = Get-Designer
            $control = Get-Control $designer $name
            $control.Enabled = $false
            Release-Children
        }
        Save-Stage '04-all-disabled'
        if ($FixtureDirectory) {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            $workbook = $null
            [void][IO.Directory]::CreateDirectory($FixtureDirectory)
            [IO.File]::Copy((Join-Path $WorkspacePath '00-baseline.xlsm'), (Join-Path $FixtureDirectory 'baseline.xlsm'), $false)
            foreach ($stage in @('00-baseline', '02-selection-ComboMain', '02-selection-ListMain', '04-all-disabled')) {
                Export-VbaProject (Join-Path $WorkspacePath "$stage.xlsm") (Join-Path $FixtureDirectory "$stage.bin")
                [IO.File]::Copy((Join-Path $WorkspacePath "$stage.json"), (Join-Path $FixtureDirectory "$stage.json"), $false)
            }
            [IO.File]::Copy((Join-Path $WorkspacePath 'environment.json'), (Join-Path $FixtureDirectory 'environment.json'), $false)
        }
        Write-Output "create complete: $WorkspacePath Excel=$($version.version) build=$($version.build) OS=$($version.operatingSystem)"
    } else {
        $excel.AutomationSecurity = 1
        $workbook = $books.Open($WorkbookPath, 0, $true)
        $snapshot = Get-Snapshot
        Release-Children
        if ($ExpectedPath) {
            $expected = Get-Content -LiteralPath $ExpectedPath -Raw -Encoding UTF8 | ConvertFrom-Json
            Assert-ExpectedSnapshot $snapshot $expected
        }
        try {
            [void]$excel.Run("'$($workbook.Name.Replace("'", "''"))'!Main.RunMutationSentinel")
        } catch {
            $project = Hold-Com $workbook.VBProject
            $components = Hold-Com $project.VBComponents
            $main = Hold-Com ($components.Item('Main'))
            $code = Hold-Com $main.CodeModule
            Write-Output "Sentinel failed; AutomationSecurity=$($excel.AutomationSecurity) Main lines=$($code.CountOfLines)"
            if ($code.CountOfLines -gt 0) { Write-Output ($code.Lines(1, $code.CountOfLines)) }
            throw
        }
        $sheets = Hold-Com $workbook.Worksheets
        $sheet = Hold-Com ($sheets.Item(1))
        $cell = Hold-Com ($sheet.Range('A1'))
        if ($cell.Value2 -cne 'issue-882-ok') { throw "Sentinel differs: $($cell.Value2)" }
        Release-Children
        if ($NormalizedWorkbookPath) {
            $workbook.SaveAs($NormalizedWorkbookPath, 52)
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            $workbook = $null
            $books = Hold-Com $excel.Workbooks
            $workbook = $books.Open($NormalizedWorkbookPath, 0, $true)
            $snapshot = Get-Snapshot
            Release-Children
            if ($ExpectedPath) { Assert-ExpectedSnapshot $snapshot $expected }
            Export-VbaProject $NormalizedWorkbookPath ($NormalizedWorkbookPath + '.bin')
            Write-Output "Excel save/reopen normalization verified: $NormalizedWorkbookPath"
        }
        $snapshot | ConvertTo-Json -Depth 10
        Write-Output "verify complete: $WorkbookPath sentinel=issue-882-ok Excel=$($version.version) build=$($version.build)"
    }
} finally {
    try { Release-Children } catch { $cleanupErrors.Add($_.Exception.Message) }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { $cleanupErrors.Add($_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch { $cleanupErrors.Add($_.Exception.Message) }
        $workbook = $null
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { $cleanupErrors.Add($_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { $cleanupErrors.Add($_.Exception.Message) }
        $excel = $null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excelPID -ne 0) {
        $process = Get-Process -Id $excelPID -ErrorAction SilentlyContinue
        if ($null -ne $process -and -not $process.WaitForExit(10000)) {
            $cleanupErrors.Add("Owned Excel PID $excelPID did not exit after Quit; terminating only this owned process")
            Stop-Process -Id $excelPID -Force
        }
    }
    if ($cleanupErrors.Count -ne 0) { throw "Excel cleanup was not clean: $($cleanupErrors -join '; ')" }
    Write-Output "cleanup confirmed: owned Excel PID=$excelPID exited; artifacts retained"
}
