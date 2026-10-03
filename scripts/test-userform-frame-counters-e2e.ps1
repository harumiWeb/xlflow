[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspacePath,
    [string]$WorkbookPath = ''
)

# Developer-only Issue #884 counter evidence capture. This script creates three
# SaveCopyAs workbooks in a fresh tmp_workspaces directory, then exports their
# VBA projects only after Excel cleanup and owned-process exit are confirmed.
$ErrorActionPreference = 'Stop'
$utf8 = [Text.UTF8Encoding]::new($false)
if ($utf8.GetPreamble().Length -ne 0) {
    throw 'Counter environment JSON must use UTF-8 without a BOM'
}
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces'))
$authoredBaseline = [IO.Path]::GetFullPath((Join-Path $repoRoot 'internal/vba/userforms/compiler/testdata/frame-excel-authored/baseline.xlsm'))

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = $authoredBaseline
} else {
    $WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
}
if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "WorkbookPath does not exist: $WorkbookPath"
}

$workspacePath = [IO.Path]::GetFullPath($WorkspacePath)
$workspacePrefix = $workspaceRoot.TrimEnd([char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)) + [IO.Path]::DirectorySeparatorChar
if (-not $workspacePath.StartsWith($workspacePrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "WorkspacePath must be a child of $workspaceRoot"
}
if (Test-Path -LiteralPath $workspacePath) {
    throw "WorkspacePath must be fresh and must not already exist: $workspacePath"
}

$harnessPath = Join-Path $PSScriptRoot 'test-userform-frame-e2e.ps1'
$tokens = $null
$parseErrors = $null
$harnessAst = [Management.Automation.Language.Parser]::ParseFile($harnessPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -ne 0) {
    throw "Could not parse shared frame harness: $($parseErrors[0].Message)"
}
$sharedFunctions = @(
    'Hold-Com',
    'Release-Children',
    'Get-ComProcessIds',
    'Get-OwnedExcelPid',
    'Export-VbaProject',
    'Get-BinaryEvidence',
    'Write-Json',
    'Get-Designer',
    'Get-ExcelEnvironment'
)
$functionAsts = $harnessAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)
$loadedFunctions = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($functionAst in $functionAsts) {
    if ($functionAst.Name -in $sharedFunctions) {
        $functionScriptBlock = [scriptblock]::Create($functionAst.Extent.Text)
        . $functionScriptBlock
        [void]$loadedFunctions.Add($functionAst.Name)
    }
}
$missingFunctions = @($sharedFunctions | Where-Object { -not $loadedFunctions.Contains($_) })
if ($missingFunctions.Count -ne 0) {
    throw "Shared frame harness functions are missing: $($missingFunctions -join ', ')"
}

[void][IO.Directory]::CreateDirectory($workspacePath)
$script:comRefs = [Collections.Generic.List[object]]::new()
$script:cleanupErrors = [Collections.Generic.List[string]]::new()
$script:excel = $null
$script:workbook = $null
$script:excelPID = 0
$script:stagesReady = $false
$failure = $null
$environment = $null
$stageDefinitions = @(
    [pscustomobject]@{ Name = 'added'; Workbook = (Join-Path $workspacePath 'added.xlsm'); Binary = (Join-Path $workspacePath 'added.bin') },
    [pscustomobject]@{ Name = 'removed'; Workbook = (Join-Path $workspacePath 'removed.xlsm'); Binary = (Join-Path $workspacePath 'removed.bin') },
    [pscustomobject]@{ Name = 'reordered'; Workbook = (Join-Path $workspacePath 'reordered.xlsm'); Binary = (Join-Path $workspacePath 'reordered.bin') }
)

$beforePids = @(Get-ComProcessIds)
if ($beforePids.Count -ne 0) {
    throw "Refusing to start Excel while another Excel process exists: $($beforePids -join ', ')"
}

try {
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false
    $script:excel.DisplayAlerts = $false
    $script:excel.EnableEvents = $false
    $script:excel.AutomationSecurity = 3
    $script:excelPID = Get-OwnedExcelPid $beforePids

    $books = Hold-Com $script:excel.Workbooks
    $script:workbook = $books.Open($WorkbookPath, 0, $true)
    $designer = Get-Designer
    $controls = Hold-Com $designer.Controls
    $parentFrame = Hold-Com ($controls.Item('ParentFrame'))
    $children = Hold-Com $parentFrame.Controls
    $nestedFrame = Hold-Com ($children.Item('NestedFrame'))
    $nestedChildren = Hold-Com $nestedFrame.Controls

    $diffLabel = Hold-Com ($nestedChildren.Add('Forms.Label.1', 'DiffLabel', $true))
    $diffLabel.Caption = 'diff'
    $script:workbook.SaveCopyAs($stageDefinitions[0].Workbook)

    [void]$nestedChildren.Remove('DiffLabel')
    $script:workbook.SaveCopyAs($stageDefinitions[1].Workbook)

    $siblingText = Hold-Com ($children.Item('SiblingText'))
    [void]$siblingText.ZOrder(0)
    $script:workbook.SaveCopyAs($stageDefinitions[2].Workbook)

    $environment = Get-ExcelEnvironment 'counter-differential'
    $script:stagesReady = $true
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
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($script:excelPID -ne 0) {
        $process = Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue
        if ($null -ne $process -and -not $process.WaitForExit(45000)) {
            $script:cleanupErrors.Add("Owned Excel PID $script:excelPID did not exit after Quit")
        }
        if (Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue) {
            $script:cleanupErrors.Add("Owned Excel PID $script:excelPID is still running")
        }
    }
}

if ($script:cleanupErrors.Count -ne 0) {
    $cleanupFailure = "Excel cleanup was not confirmed: $($script:cleanupErrors -join '; ')"
    if ($null -ne $failure) { throw "$($failure.Exception.Message); $cleanupFailure" }
    throw $cleanupFailure
}
if ($null -ne $failure) { throw $failure }
if (-not $script:stagesReady) { throw 'Counter capture did not finish all SaveCopyAs stages' }

# Excel is closed and its uniquely owned PID has exited before any VBA binary
# extraction or environment evidence is written.
$binaryEvidence = [Collections.Generic.List[object]]::new()
foreach ($stage in $stageDefinitions) {
    if (-not (Test-Path -LiteralPath $stage.Workbook -PathType Leaf)) {
        throw "SaveCopyAs output is missing: $($stage.Workbook)"
    }
    Export-VbaProject $stage.Workbook $stage.Binary
    $binaryEvidence.Add((Get-BinaryEvidence $stage.Binary))
}
$environment | Add-Member -NotePropertyName inputWorkbook -NotePropertyValue $WorkbookPath
$environment.harness = 'scripts/test-userform-frame-counters-e2e.ps1'
$environment | Add-Member -NotePropertyName workspacePath -NotePropertyValue $workspacePath
$environment | Add-Member -NotePropertyName cleanupConfirmed -NotePropertyValue $true
$environment | Add-Member -NotePropertyName ownedExcelPidExited -NotePropertyValue $true
$environment | Add-Member -NotePropertyName stages -NotePropertyValue @($stageDefinitions | ForEach-Object { $_.Name })
$environment | Add-Member -NotePropertyName binaries -NotePropertyValue @($binaryEvidence.ToArray())
Write-Json (Join-Path $workspacePath 'environment.json') $environment

Write-Output "Counter differential workbooks and VBA projects saved; cleanup confirmed for owned Excel PID $($environment.processId)."
Write-Output "Workspace: $workspacePath"
