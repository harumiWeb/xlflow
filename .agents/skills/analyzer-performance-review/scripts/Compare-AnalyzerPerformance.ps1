[CmdletBinding(DefaultParameterSetName = "Run")]
param(
    [Parameter(ParameterSetName = "Run")]
    [string] $BaseRef = "origin/main",

    [Parameter(ParameterSetName = "Run")]
    [ValidateRange(1, 20)]
    [int] $Count = 3,

    [Parameter(ParameterSetName = "Run")]
    [ValidateSet("ronecone", "std-vba")]
    [string[]] $Projects = @("ronecone", "std-vba"),

    [Parameter(ParameterSetName = "Run")]
    [ValidateSet("cold", "warm", "local-edit", "dependency-edit")]
    [string[]] $Modes = @("cold", "warm", "local-edit", "dependency-edit"),

    [Parameter(Mandatory = $true, ParameterSetName = "Compare")]
    [string] $BaseLog,

    [Parameter(Mandatory = $true, ParameterSetName = "Compare")]
    [string] $HeadLog,

    [string] $OutputDirectory,

    [switch] $FailOnRegression
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-GitText {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Repository,

        [Parameter(Mandatory = $true)]
        [string[]] $Arguments
    )

    $previousErrorActionPreference = $ErrorActionPreference
    try {
        # Native tools can write progress to stderr even when they succeed.
        # Judge them by the exit code instead of promoting stderr to a
        # terminating PowerShell error.
        $ErrorActionPreference = "Continue"
        $output = & rtk proxy git -C $Repository @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
    if ($exitCode -ne 0) {
        throw "git $($Arguments -join ' ') failed:`n$($output -join [Environment]::NewLine)"
    }
    return (($output | Out-String).Trim())
}

function Get-Median {
    param([double[]] $Values)

    if ($Values.Count -eq 0) {
        return $null
    }
    $ordered = @($Values | Sort-Object)
    $middle = [int][Math]::Floor($ordered.Count / 2)
    if ($ordered.Count % 2 -eq 1) {
        return [double]$ordered[$middle]
    }
    return ([double]$ordered[$middle - 1] + [double]$ordered[$middle]) / 2.0
}

function Read-BenchmarkRows {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Side
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($line in [System.IO.File]::ReadLines((Resolve-Path -LiteralPath $Path))) {
        if ($line -notmatch '^BenchmarkRealWorldCorpus/(?<project>[^/]+)/analyze-only/(?<mode>cold|warm|local-edit|dependency-edit)-\d+\s+(?<rest>.+)$') {
            continue
        }
        $tokens = @($Matches.rest -split '\s+')
        if ($tokens.Count -lt 3) {
            continue
        }
        $metrics = [ordered]@{}
        for ($index = 1; $index + 1 -lt $tokens.Count; $index += 2) {
            $value = 0.0
            if ([double]::TryParse($tokens[$index], [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
                $metrics[$tokens[$index + 1]] = $value
            }
        }
        $rows.Add([pscustomobject]@{
                side = $Side
                project = $Matches.project
                mode = $Matches.mode
                metrics = $metrics
            })
    }
    if ($rows.Count -eq 0) {
        throw "No analyzer corpus benchmark rows found in '$Path'."
    }
    return @($rows)
}

function Get-Aggregates {
    param([object[]] $Rows)

    $aggregates = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $Rows | Group-Object side, project, mode) {
        $first = $group.Group[0]
        $metricNames = @($group.Group | ForEach-Object { $_.metrics.Keys } | Sort-Object -Unique)
        $metrics = [ordered]@{}
        foreach ($metricName in $metricNames) {
            $values = @($group.Group | ForEach-Object {
                    if ($_.metrics.Contains($metricName)) {
                        [double]$_.metrics[$metricName]
                    }
                })
            if ($values.Count -eq $group.Count) {
                $metrics[$metricName] = Get-Median -Values $values
            }
        }
        $aggregates.Add([pscustomobject]@{
                side = $first.side
                project = $first.project
                mode = $first.mode
                samples = $group.Count
                metrics = $metrics
            })
    }
    return @($aggregates)
}

function Compare-Aggregates {
    param([object[]] $Aggregates)

    $groups = @($Aggregates | Group-Object project, mode)
    foreach ($group in $groups) {
        $baseRows = @($group.Group | Where-Object side -eq "base")
        $headRows = @($group.Group | Where-Object side -eq "head")
        if ($group.Group.Count -ne 2 -or $baseRows.Count -ne 1 -or $headRows.Count -ne 1) {
            throw "Aggregate group $($group.Name) must contain exactly one base and one head row (base=$($baseRows.Count), head=$($headRows.Count))."
        }
        $baseMetricNames = @($baseRows[0].metrics.Keys | Sort-Object)
        $headMetricNames = @($headRows[0].metrics.Keys | Sort-Object)
        $baseOnly = @($baseMetricNames | Where-Object { $_ -notin $headMetricNames })
        $headOnly = @($headMetricNames | Where-Object { $_ -notin $baseMetricNames })
        if ($baseOnly.Count -gt 0 -or $headOnly.Count -gt 0) {
            $baseOnlyText = [string]::Join(", ", $baseOnly)
            $headOnlyText = [string]::Join(", ", $headOnly)
            throw "Metric sets differ for $($group.Name) (base-only: [$baseOnlyText], head-only: [$headOnlyText])."
        }
    }

    $comparisons = [System.Collections.Generic.List[object]]::new()
    $headRows = @($Aggregates | Where-Object side -eq "head")
    foreach ($head in $headRows) {
        $base = $Aggregates | Where-Object { $_.side -eq "base" -and $_.project -eq $head.project -and $_.mode -eq $head.mode } | Select-Object -First 1
        if ($null -eq $base) {
            throw "Missing base samples for $($head.project)/$($head.mode)."
        }
        foreach ($metricName in @($head.metrics.Keys)) {
            $baseValue = [double]$base.metrics[$metricName]
            $headValue = [double]$head.metrics[$metricName]
            $deltaPercent = if ($baseValue -eq 0.0) {
                if ($headValue -eq 0.0) { 0.0 } else { $null }
            } else {
                (($headValue - $baseValue) / $baseValue) * 100.0
            }
            $deterministic = $metricName.StartsWith("counter_", [StringComparison]::Ordinal)
            $suspicious = if ($deterministic) {
                $headValue -gt $baseValue
            } elseif ($metricName -in @("B/op", "allocs/op")) {
                $null -eq $deltaPercent -or $deltaPercent -gt 5.0
            } elseif ($metricName -eq "ns/op") {
                $null -eq $deltaPercent -or $deltaPercent -gt 10.0
            } else {
                $false
            }
            $comparisons.Add([pscustomobject]@{
                    project = $head.project
                    mode = $head.mode
                    metric = $metricName
                    base = $baseValue
                    head = $headValue
                    delta_percent = $deltaPercent
                    deterministic = $deterministic
                    suspicious = $suspicious
                })
        }
    }
    return @($comparisons)
}

function Find-UnstableCounters {
    param([object[]] $Rows)

    $unstable = [System.Collections.Generic.List[object]]::new()
    foreach ($group in $Rows | Group-Object side, project, mode) {
        $first = $group.Group[0]
        $counterNames = @($group.Group | ForEach-Object { $_.metrics.Keys } | Where-Object { $_.StartsWith("counter_", [StringComparison]::Ordinal) } | Sort-Object -Unique)
        foreach ($counterName in $counterNames) {
            $values = @($group.Group | ForEach-Object {
                    if ($_.metrics.Contains($counterName)) {
                        [double]$_.metrics[$counterName]
                    }
                } | Sort-Object -Unique)
            if ($values.Count -gt 1) {
                $unstable.Add([pscustomobject]@{
                        side = $first.side
                        project = $first.project
                        mode = $first.mode
                        metric = $counterName
                        values = $values
                    })
            }
        }
    }
    return @($unstable)
}

function Format-MetricValue {
    param([double] $Value)
    if ([Math]::Abs($Value) -ge 1000) {
        return $Value.ToString("N0", [Globalization.CultureInfo]::InvariantCulture)
    }
    return $Value.ToString("0.###", [Globalization.CultureInfo]::InvariantCulture)
}

function Write-Report {
    param(
        [object[]] $Aggregates,
        [object[]] $Comparisons,
        [object[]] $UnstableCounters,
        [string] $Directory,
        [hashtable] $Metadata
    )

    [System.IO.Directory]::CreateDirectory($Directory) | Out-Null
    $payload = [ordered]@{
        metadata = $Metadata
        aggregates = $Aggregates
        comparisons = $Comparisons
        unstable_counters = $UnstableCounters
    }
    $jsonPath = Join-Path $Directory "comparison.json"
    [System.IO.File]::WriteAllText($jsonPath, ($payload | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# Analyzer performance comparison")
    $lines.Add("")
    $lines.Add("- Base: ``$($Metadata.base)``")
    $lines.Add("- Head: ``$($Metadata.head)``")
    if ($Metadata.ContainsKey("head_has_uncommitted_changes")) {
        $lines.Add("- Head included uncommitted changes: $($Metadata.head_has_uncommitted_changes)")
    }
    $lines.Add("- Samples: $($Metadata.count)")
    $lines.Add("- Command: ``$($Metadata.command)``")
    $lines.Add("")
    $lines.Add("## Primary metrics (median)")
    $lines.Add("")
    $lines.Add("| Project | Mode | Metric | Base | Head | Delta |")
    $lines.Add("| --- | --- | --- | ---: | ---: | ---: |")
    foreach ($item in $Comparisons | Where-Object metric -in @("ns/op", "B/op", "allocs/op") | Sort-Object project, mode, metric) {
        $delta = if ($null -eq $item.delta_percent) { "n/a" } else { "{0:+0.0;-0.0;0.0}%" -f $item.delta_percent }
        $lines.Add("| $($item.project) | $($item.mode) | $($item.metric) | $(Format-MetricValue $item.base) | $(Format-MetricValue $item.head) | $delta |")
    }
    $lines.Add("")
    $lines.Add("## Changed deterministic counters")
    $lines.Add("")
    $changedCounters = @($Comparisons | Where-Object { $_.deterministic -and $_.base -ne $_.head } | Sort-Object project, mode, metric)
    if ($changedCounters.Count -eq 0) {
        $lines.Add("No deterministic counter changed.")
    } else {
        $lines.Add("| Project | Mode | Counter | Base | Head | Delta | Suspicious |")
        $lines.Add("| --- | --- | --- | ---: | ---: | ---: | :---: |")
        foreach ($item in $changedCounters) {
            $lines.Add("| $($item.project) | $($item.mode) | $($item.metric) | $(Format-MetricValue $item.base) | $(Format-MetricValue $item.head) | $(Format-MetricValue ($item.head - $item.base)) | $($item.suspicious) |")
        }
    }
    $lines.Add("")
    $lines.Add("## Unstable deterministic counters")
    $lines.Add("")
    if ($UnstableCounters.Count -eq 0) {
        $lines.Add("No deterministic counter varied between repeated samples.")
    } else {
        foreach ($item in $UnstableCounters | Sort-Object side, project, mode, metric) {
            $lines.Add("- $($item.side) $($item.project)/$($item.mode) $($item.metric): $([string]::Join(', ', $item.values))")
        }
    }
    $lines.Add("")
    $suspiciousCount = @($Comparisons | Where-Object suspicious).Count + $UnstableCounters.Count
    $lines.Add("Suspicious metric increases: **$suspiciousCount**. Each requires explanation; elapsed-time noise alone is not a structural regression verdict.")

    $reportPath = Join-Path $Directory "report.md"
    [System.IO.File]::WriteAllLines($reportPath, $lines, [Text.UTF8Encoding]::new($false))
    return [pscustomobject]@{ Json = $jsonPath; Markdown = $reportPath; Suspicious = $suspiciousCount }
}

$repoRoot = Invoke-GitText -Repository $PSScriptRoot -Arguments @("rev-parse", "--show-toplevel")
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $OutputDirectory = Join-Path $repoRoot ".tmp-analyzer-performance-$stamp"
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
[System.IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null

$metadata = @{
    base = "log:$BaseLog"
    head = "log:$HeadLog"
    count = 0
    command = "compare existing logs"
}

if ($PSCmdlet.ParameterSetName -eq "Run") {
    $baseSha = Invoke-GitText -Repository $repoRoot -Arguments @("rev-parse", "$BaseRef^{commit}")
    $headSha = Invoke-GitText -Repository $repoRoot -Arguments @("rev-parse", "HEAD^{commit}")
    $projectPattern = [string]::Join("|", $Projects)
    $modePattern = [string]::Join("|", $Modes)
    $benchmarkPattern = "^BenchmarkRealWorldCorpus/($projectPattern)/analyze-only/($modePattern)$"
    $goArguments = @(
        "test", "./internal/staticanalysis/corpus", "-run", "^$", "-bench", $benchmarkPattern,
        "-benchmem", "-benchtime=1x", "-count=$Count", "-timeout=25m"
    )
    $commandText = "scripts/dev/go.ps1 $($goArguments -join ' ')"
    $basePath = Join-Path $OutputDirectory "base.txt"
    $headPath = Join-Path $OutputDirectory "head.txt"
    $tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
    $baseWorktree = Join-Path $tempRoot ("xlflow-analyzer-perf-" + [Guid]::NewGuid().ToString("N"))
    $baseWorktree = [System.IO.Path]::GetFullPath($baseWorktree)
    if (-not $baseWorktree.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [System.IO.Path]::GetFileName($baseWorktree) -notlike "xlflow-analyzer-perf-*") {
        throw "Refusing unsafe temporary worktree path '$baseWorktree'."
    }
    try {
        Invoke-GitText -Repository $repoRoot -Arguments @("worktree", "add", "--detach", $baseWorktree, $baseSha) | Out-Null
        foreach ($run in @(
                @{ Name = "base"; Root = $baseWorktree; Log = $basePath },
                @{ Name = "head"; Root = $repoRoot; Log = $headPath }
            )) {
            Write-Output "Running $($run.Name) analyzer benchmark..."
            $goScript = Join-Path $run.Root "scripts\dev\go.ps1"
            Push-Location $run.Root
            try {
                $previousErrorActionPreference = $ErrorActionPreference
                try {
                    $ErrorActionPreference = "Continue"
                    $output = & rtk proxy powershell -NoProfile -ExecutionPolicy Bypass -File $goScript @goArguments 2>&1
                    $exitCode = $LASTEXITCODE
                } finally {
                    $ErrorActionPreference = $previousErrorActionPreference
                }
            } finally {
                Pop-Location
            }
            [System.IO.File]::WriteAllLines($run.Log, [string[]]$output, [Text.UTF8Encoding]::new($false))
            if ($exitCode -ne 0) {
                throw "$($run.Name) benchmark failed with exit code $exitCode. See '$($run.Log)'."
            }
        }
    } finally {
        if (Test-Path -LiteralPath $baseWorktree) {
            $previousErrorActionPreference = $ErrorActionPreference
            try {
                $ErrorActionPreference = "Continue"
                & rtk proxy git -C $repoRoot worktree remove --force $baseWorktree 2>&1 | Out-Null
                $removeExitCode = $LASTEXITCODE
            } finally {
                $ErrorActionPreference = $previousErrorActionPreference
            }
            if ($removeExitCode -ne 0) {
                Write-Warning "Could not remove temporary worktree '$baseWorktree'."
            }
        }
    }
    $BaseLog = $basePath
    $HeadLog = $headPath
    $metadata = @{
        base = $baseSha
        head = $headSha
        head_has_uncommitted_changes = -not [string]::IsNullOrWhiteSpace((Invoke-GitText -Repository $repoRoot -Arguments @("status", "--short")))
        count = $Count
        projects = $Projects
        modes = $Modes
        command = $commandText
    }
}

$rows = @(
    Read-BenchmarkRows -Path $BaseLog -Side "base"
    Read-BenchmarkRows -Path $HeadLog -Side "head"
)
$aggregates = Get-Aggregates -Rows $rows
$sampleCounts = @($aggregates | ForEach-Object samples | Sort-Object -Unique)
if ($sampleCounts.Count -ne 1) {
    throw "Base/head benchmark sample counts differ: $([string]::Join(', ', $sampleCounts))."
}
if ($metadata.count -eq 0) {
    $metadata.count = $sampleCounts[0]
} elseif ($metadata.count -ne $sampleCounts[0]) {
    throw "Benchmark produced $($sampleCounts[0]) samples per case, expected $($metadata.count)."
}
if ($PSCmdlet.ParameterSetName -eq "Run") {
    $expectedAggregateCount = $Projects.Count * $Modes.Count * 2
    if ($aggregates.Count -ne $expectedAggregateCount) {
        throw "Benchmark produced $($aggregates.Count) base/head project-mode groups, expected $expectedAggregateCount."
    }
}
$comparisons = Compare-Aggregates -Aggregates $aggregates
$unstableCounters = @(Find-UnstableCounters -Rows $rows)
$result = Write-Report -Aggregates $aggregates -Comparisons $comparisons -UnstableCounters $unstableCounters -Directory $OutputDirectory -Metadata $metadata

Write-Output "Analyzer performance report: $($result.Markdown)"
Write-Output "Machine-readable comparison: $($result.Json)"
Write-Output "Suspicious metric increases: $($result.Suspicious)"
if ($FailOnRegression -and $result.Suspicious -gt 0) {
    exit 2
}
