# Developer-only regression checks; no Excel or xlflow invocation.
$ErrorActionPreference = 'Stop'
$harness = Join-Path $PSScriptRoot 'test-pack-userforms-e2e.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($harness, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors -join "`n") }
foreach ($name in @('Get-ProcessIdentity', 'Stop-OwnedExcel')) {
    $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    . ([ScriptBlock]::Create($function.Extent.Text))
}

function Get-Process {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Isolated cleanup test stub prevents querying real processes.')]
    param([int]$Id, [object]$ErrorAction)
    if ($Id -ne 12345 -or $ErrorAction -ne 'SilentlyContinue') { throw 'Unexpected process query' }
    $script:processLookups++
    if ($script:processLookups -gt 1) { throw 'Cleanup looked up a potentially reused PID' }
    return $script:fakeProcess
}

function Stop-Process {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Isolated cleanup test trap prevents termination of real processes.')]
    param()
    throw 'Cleanup must never terminate by PID'
}
function Release-ComRefs { }

$startTime = [DateTime]::UtcNow
$script:comRefs = [Collections.Generic.List[object]]::new()
foreach ($case in @('normal exit', 'forced exit', 'reused PID', 'unknown identity')) {
    $script:processLookups = 0
    $script:currentCleanupErrors = [Collections.Generic.List[string]]::new()
    $script:currentWorkbook = $null
    $script:excel = $null
    $script:excelPID = 12345
    $script:excelStartTime = $startTime.ToUniversalTime().ToString('o')
    $script:excelRecord = @{ cleanup = $null }
    $record = $script:excelRecord
    $script:fakeProcess = [PSCustomObject]@{
        Id = 12345
        ProcessName = 'EXCEL'
        StartTime = $(if ($case -eq 'reused PID') { $startTime.AddSeconds(1) } else { $startTime })
        Pinned = $false
        Killed = $false
        Disposed = $false
        WaitCount = 0
        ExitImmediately = ($case -eq 'normal exit')
    }
    if ($case -eq 'unknown identity') {
        $script:fakeProcess | Add-Member ScriptProperty StartTime { throw 'Unknown start time' } -Force
        $script:excelStartTime = '<unavailable>'
    }
    $script:fakeProcess | Add-Member ScriptProperty Handle { $this.Pinned = $true; return 1 }
    $script:fakeProcess | Add-Member ScriptMethod WaitForExit {
        param([int]$Timeout)
        if ($Timeout -ne 10000) { throw 'Unexpected cleanup timeout' }
        $this.WaitCount++
        return ($this.ExitImmediately -or $this.WaitCount -gt 1)
    }
    $script:fakeProcess | Add-Member ScriptMethod Kill {
        if (-not $this.Pinned) { throw 'Process handle was not pinned' }
        $this.Killed = $true
    }
    $script:fakeProcess | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
    Stop-OwnedExcel
    $expectedExit = $case -in @('normal exit', 'forced exit')
    $expectedKill = $case -eq 'forced exit'
    if ($record.cleanup.exited -ne $expectedExit -or $script:fakeProcess.Killed -ne $expectedKill -or $record.cleanup.forced_termination -ne $expectedKill) {
        throw "Unexpected cleanup result for $case"
    }
    if (-not $script:fakeProcess.Pinned -or -not $script:fakeProcess.Disposed -or $script:processLookups -ne 1) {
        throw "Process ownership was not maintained for $case"
    }
    if ($expectedExit -and $record.cleanup.errors.Count) { throw ($record.cleanup.errors -join "`n") }
    if (-not $expectedExit -and $record.cleanup.errors.Count -ne 1) { throw "Unverified identity was not rejected for $case" }
}

# Exercise the actual script's early validation. Keep all test artifacts under
# tmp_workspaces, including the junction target; do not invoke Excel on failure.
$testRoot = Join-Path (Join-Path $PSScriptRoot '../tmp_workspaces') ('issue-887-harness-' + [Guid]::NewGuid().ToString('N'))
$testRoot = [IO.Path]::GetFullPath($testRoot)
$target = Join-Path $testRoot 'target'
[void][IO.Directory]::CreateDirectory($target)
$link = Join-Path $testRoot 'link'
[void](New-Item -ItemType Junction -Path $link -Target $target)
$workspace = Join-Path $link 'must-not-be-created'
$rejected = $false
try { & $harness -WorkspacePath $workspace } catch {
    if ($_.Exception.Message -notlike '*must not traverse a reparse point*') { throw }
    $rejected = $true
}
if (-not $rejected -or (Test-Path -LiteralPath (Join-Path $target 'must-not-be-created'))) {
    throw 'Junction workspace was not rejected before writing'
}
Write-Output "Harness ownership and junction regressions passed; retained workspace: $testRoot"
