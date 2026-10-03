[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$WorkspacePath,
    [string]$WorkbookPath = '',
    [string]$ExpectedSentinel = ''
)

# Developer-only, sequential Excel evidence. Never called by ordinary tests/CI.
$ErrorActionPreference = 'Stop'
$script:utf8 = [Text.UTF8Encoding]::new($false)
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces'))
$workspacePath = [IO.Path]::GetFullPath($WorkspacePath)
if (-not $workspacePath.StartsWith($workspaceRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Workspace must be inside tmp_workspaces' }
if (Test-Path -LiteralPath $workspacePath) { throw 'Workspace must be fresh' }
if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Close existing Excel processes before running this gate' }
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'test-userform-frame-e2e.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -ne 0) { throw 'Cannot parse shared COM helpers' }
$helpers = @('Hold-Com', 'Release-Children', 'Get-ComProcessIds', 'Get-OwnedExcelPid', 'Export-VbaProject', 'Write-Json', 'Get-OptionalProperty')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    if ($node.Name -in $helpers) { . ([scriptblock]::Create($node.Extent.Text)) }
}
[void][IO.Directory]::CreateDirectory($workspacePath)
$script:comRefs = [Collections.Generic.List[object]]::new()
$script:cleanupErrors = [Collections.Generic.List[string]]::new()
$script:excel = $null; $script:workbook = $null; $script:excelPID = 0
$stages = [Collections.Generic.List[object]]::new()
$failure = $null

function Get-TabEvidence([object]$Collection) {
    $rows = @()
    for ($index = 0; $index -lt $Collection.Count; $index++) {
        $item = Hold-Com ($Collection.Item($index))
        $row = [ordered]@{ name = [string]$item.Name }
        foreach ($key in @('Caption', 'ControlTipText', 'Tag', 'Accelerator', 'Enabled', 'Visible', 'Left', 'Top', 'Width', 'Height')) { $row[$key] = Get-OptionalProperty $item $key }
        $rows += [pscustomobject]$row
    }
    return ,$rows
}
function Get-MultiEvidence {
    $components = Hold-Com (Hold-Com $script:workbook.VBProject).VBComponents
    $component = Hold-Com ($components.Item('MultiTopologyForm'))
    $designer = Hold-Com $component.Designer
    $controls = Hold-Com $designer.Controls
    $multi = Hold-Com ($controls.Item('MultiMain'))
    $strip = Hold-Com ($controls.Item('StripMain'))
    return [pscustomobject]@{
        pageValue = Get-OptionalProperty $multi 'Value'
        tabValue = Get-OptionalProperty $strip 'Value'
        enabled = [bool]$multi.Enabled; tabEnabled = [bool]$strip.Enabled
        width = $multi.Width; height = $multi.Height
        pages = Get-TabEvidence (Hold-Com $multi.Pages)
        tabs = Get-TabEvidence (Hold-Com $strip.Tabs)
    }
}
function Save-Stage([string]$Name) {
    Write-Output "Saving stage $Name"
    $path = Join-Path $workspacePath ($Name + '.xlsm')
    $before = Get-MultiEvidence
    $script:workbook.SaveCopyAs($path)
    $stages.Add([pscustomobject]@{ name = $Name; workbook = $path; before = $before; runtimeSentinel = $ExpectedSentinel })
}
function Test-Runtime {
    if (-not $ExpectedSentinel) { return }
    $sheet = Hold-Com ((Hold-Com $script:workbook.Worksheets).Item(1))
    $cell = Hold-Com ($sheet.Range('A1'))
    $cell.Value2 = ''
    $qualifiedMacro = "'" + $script:workbook.Name.Replace("'", "''") + "'!Main.RunMultiSentinel"
    [void]$script:excel.Run($qualifiedMacro)
    if ([string]$cell.Value2 -ne $ExpectedSentinel) { throw "Runtime sentinel differs: $($cell.Value2)" }
    Write-Output "Runtime confirmed: $($cell.Value2)"
}
try {
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false; $script:excel.DisplayAlerts = $false
    $script:excel.EnableEvents = $false; $script:excel.AutomationSecurity = 3
    if ($ExpectedSentinel) { $script:excel.AutomationSecurity = 1 }
    $script:excelPID = Get-OwnedExcelPid @()
    $books = Hold-Com $script:excel.Workbooks
    if ($WorkbookPath) {
        $script:workbook = $books.Open([IO.Path]::GetFullPath($WorkbookPath), 0, $false)
        Test-Runtime
        Save-Stage 'normalized'
    } else {
        $script:workbook = $books.Add()
        $project = Hold-Com $script:workbook.VBProject
        $components = Hold-Com $project.VBComponents
        $component = Hold-Com ($components.Add(3))
        $component.Name = 'MultiTopologyForm'
        $designer = Hold-Com $component.Designer
        $props = Hold-Com $component.Properties
        (Hold-Com ($props.Item('Width'))).Value = 360
        (Hold-Com ($props.Item('Height'))).Value = 300
        $controls = Hold-Com $designer.Controls
        $multi = Hold-Com ($controls.Add('Forms.MultiPage.1', 'MultiMain', $true))
        $multi.Left = 12; $multi.Top = 12; $multi.Width = 240; $multi.Height = 180
        $pages = Hold-Com $multi.Pages
        $first = Hold-Com ($pages.Item(0)); $first.Name = 'PageAlpha'; $first.Caption = 'Alpha'
        $second = Hold-Com ($pages.Item(1)); $second.Name = 'PageBeta'; $second.Caption = 'Beta'
        $first.Tag = 'alpha-tag'; $first.ControlTipText = 'alpha-tip'; $first.Accelerator = 'A'
        $inner = Hold-Com $first.Controls
        $text = Hold-Com ($inner.Add('Forms.TextBox.1', 'PageText', $true))
        $text.Left = 8; $text.Top = 8; $text.Width = 100; $text.Height = 18; $text.Value = 'page text'
        $frame = Hold-Com ($inner.Add('Forms.Frame.1', 'PageFrame', $true))
        $frame.Left = 8; $frame.Top = 36; $frame.Width = 150; $frame.Height = 80
        $label = Hold-Com ((Hold-Com $frame.Controls).Add('Forms.Label.1', 'NestedLabel', $true)); $label.Caption = 'nested'
        $strip = Hold-Com ($controls.Add('Forms.TabStrip.1', 'StripMain', $true))
        $strip.Left = 12; $strip.Top = 206; $strip.Width = 240; $strip.Height = 48
        $tabs = Hold-Com $strip.Tabs
        $tab = Hold-Com ($tabs.Item(0)); $tab.Name = 'TabAlpha'; $tab.Caption = 'Alpha'; $tab.Tag = 'tab-tag'; $tab.ControlTipText = 'tab-tip'; $tab.Accelerator = 'T'
        (Hold-Com ($tabs.Item(1))).Name = 'TabBeta'
        $multi.Value = 1; $strip.Value = 1
        $script:workbook.SaveAs((Join-Path $workspacePath 'working.xlsm'), 52)
        Save-Stage 'baseline'
        $first.Caption = -join @([char]0x65e5, [char]0x672c, [char]0x8a9e); $first.Enabled = $false; $first.Visible = $false
        $multi.Enabled = $false; $strip.Enabled = $false
        Save-Stage 'disabled'
        $first.Enabled = $true; $first.Visible = $true; $multi.Enabled = $true; $strip.Enabled = $true
        $first.Caption = 'Alpha'
        $third = Hold-Com ($pages.Add('PageGamma', 'Gamma')); $third.Tag = 'gamma-tag'
        [void]$tabs.Add('TabGamma', 'Gamma')
        Save-Stage 'added'
        $multi.Width = 264; $multi.Height = 204
        Save-Stage 'resized'
        $first.Index = 2
        (Hold-Com ($tabs.Item('TabAlpha'))).Index = 2
        Save-Stage 'reordered'
        # Reacquire collections after reordering, rather than borrowing stale members.
        $pages = Hold-Com $multi.Pages; $tabs = Hold-Com $strip.Tabs
        [void]$pages.Remove('PageGamma'); [void]$tabs.Remove(1)
        Save-Stage 'removed'
        for ($index = [int]$pages.Count - 1; $index -ge 0; $index--) {
            $pageName = [string](Hold-Com ($pages.Item($index))).Name
            Write-Output "Removing page $pageName"
            [void]$pages.Remove($pageName)
        }
        for ($index = [int]$tabs.Count - 1; $index -ge 0; $index--) { Write-Output "Removing tab $index"; [void]$tabs.Remove($index) }
        Save-Stage 'empty'
    }
    $script:workbook.Close($false)
    foreach ($stage in $stages) {
        Write-Output "Reopening stage $($stage.name)"
        $script:workbook = $books.Open($stage.workbook, 0, $true)
        $stage | Add-Member -NotePropertyName after -NotePropertyValue (Get-MultiEvidence)
        Test-Runtime
        $script:workbook.Close($false)
    }
    $environment = [pscustomobject]@{ issue = 885; harness = 'scripts/test-userform-multipage-e2e.ps1'; capturedUtc = [DateTime]::UtcNow.ToString('o'); version = $script:excel.Version; build = $script:excel.Build; processId = $script:excelPID; cleanupConfirmed = $false; runtimeVerified = [bool]$ExpectedSentinel }
} catch { $failure = $_; Write-Output "Capture failed: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" } finally {
    if ($null -ne $script:workbook) { try { $script:workbook.Close($false) } catch { Write-Verbose "Workbook was already closed: $($_.Exception.Message)" } }
    try { Release-Children -Final } catch { $script:cleanupErrors.Add($_.Exception.Message) }
    if ($null -ne $script:workbook) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:workbook) } catch { $script:cleanupErrors.Add($_.Exception.Message) } }
    if ($null -ne $script:excel) {
        try { $script:excel.Quit() } catch { $script:cleanupErrors.Add($_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel) } catch { $script:cleanupErrors.Add($_.Exception.Message) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($script:excelPID -ne 0) {
        $owned = Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue
        if ($null -ne $owned -and -not $owned.WaitForExit(15000)) { $script:cleanupErrors.Add('Owned Excel process did not exit') }
    }
}
if ($script:cleanupErrors.Count -gt 0) { throw (($script:cleanupErrors -join '; ') + $(if ($failure) { ' Original failure: ' + $failure.Exception.Message })) }
if ($null -ne $failure) { throw $failure }
$environment.cleanupConfirmed = $true
foreach ($stage in $stages) {
    Export-VbaProject $stage.workbook (Join-Path $workspacePath ($stage.name + '.bin'))
    $hashAlgorithm = [Security.Cryptography.SHA256]::Create()
    try { $binaryHash = $hashAlgorithm.ComputeHash([IO.File]::ReadAllBytes((Join-Path $workspacePath ($stage.name + '.bin')))) } finally { $hashAlgorithm.Dispose() }
    $stage | Add-Member -NotePropertyName binarySha256 -NotePropertyValue ([BitConverter]::ToString($binaryHash).Replace('-', '').ToLowerInvariant())
    Write-Json (Join-Path $workspacePath ($stage.name + '.json')) $stage
}
Write-Json (Join-Path $workspacePath 'environment.json') $environment
Write-Output "MultiPage evidence captured with confirmed cleanup: $workspacePath"
