param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$ExpectedHeading,

    [string]$ReportDirectory = ""
)

$ErrorActionPreference = 'Stop'

function Resolve-TaktReportDirectory {
    param([string]$ExplicitReportDirectory)

    if (-not [string]::IsNullOrWhiteSpace($ExplicitReportDirectory)) {
        return $ExplicitReportDirectory
    }

    foreach ($envName in @('TAKT_REPORT_DIR', 'REPORT_DIR')) {
        $value = [Environment]::GetEnvironmentVariable($envName)
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    $runsRoot = Join-Path $PSScriptRoot 'runs'
    if (-not (Test-Path -LiteralPath $runsRoot)) {
        return $null
    }

    $runDirectories = Get-ChildItem -LiteralPath $runsRoot -Directory -ErrorAction SilentlyContinue
    if ($null -eq $runDirectories) {
        return $null
    }

    $runningRun = $runDirectories |
        ForEach-Object {
            $metaPath = Join-Path $_.FullName 'meta.json'
            if (-not (Test-Path -LiteralPath $metaPath)) {
                return
            }

            $meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
            if ($meta.status -ne 'running') {
                return
            }

            [pscustomobject]@{
                Meta = $meta
                MetaPath = $metaPath
            }
        } |
        Sort-Object { [datetime]$_.Meta.updatedAt } -Descending |
        Select-Object -First 1

    if ($null -eq $runningRun) {
        return $null
    }

    return Join-Path (Split-Path -Parent $PSScriptRoot) $runningRun.Meta.reportDirectory
}

function Resolve-ReportPath {
    param(
        [string]$RawPath,
        [string]$ExplicitReportDirectory
    )

    foreach ($prefix in @('{report_dir}/', '｛report_dir｝/')) {
        if ($RawPath.StartsWith($prefix, [System.StringComparison]::Ordinal)) {
            $reportDirectory = Resolve-TaktReportDirectory -ExplicitReportDirectory $ExplicitReportDirectory
            if ([string]::IsNullOrWhiteSpace($reportDirectory)) {
                return $null
            }

            $relativeReportPath = $RawPath.Substring($prefix.Length)
            return Join-Path $reportDirectory $relativeReportPath
        }
    }

    return $RawPath
}

function Test-IsLeadingMetaLine {
    param([string]$Line)

    $trimmed = $Line.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $true
    }

    foreach ($prefix in @(
        'The user wants me',
        'Let me',
        'I need to',
        'Now I',
        'I''m in',
        'Looking back',
        'Got it',
        'Understood',
        'Here is',
        'I will',
        'First, I'
    )) {
        if ($trimmed.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Remove-LeadingNoise {
    param([string[]]$Lines)

    $result = New-Object System.Collections.Generic.List[string]
    $seenContent = $false
    $inFence = $false

    foreach ($line in $Lines) {
        $trimmed = $line.Trim()

        if (-not $seenContent) {
            if ($trimmed -match '^```') {
                $inFence = -not $inFence
                continue
            }

            if ($inFence) {
                continue
            }

            if (Test-IsLeadingMetaLine -Line $line) {
                continue
            }
        }

        if (-not $seenContent -and [string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        $seenContent = $true
        $result.Add($line)
    }

    while ($result.Count -gt 0 -and [string]::IsNullOrWhiteSpace($result[$result.Count - 1])) {
        $result.RemoveAt($result.Count - 1)
    }

    return ,$result.ToArray()
}

$resolvedPath = Resolve-ReportPath -RawPath $Path -ExplicitReportDirectory $ReportDirectory
if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
    Write-Output "sanitize-report: skipped unresolved placeholder path '$Path'"
    exit 0
}

if (-not (Test-Path -LiteralPath $resolvedPath)) {
    Write-Output "sanitize-report: skipped missing report '$resolvedPath'"
    exit 0
}

$lines = Get-Content -LiteralPath $resolvedPath
$headingIndex = -1
$hadFenceBeforeHeading = $false
for ($index = 0; $index -lt $lines.Count; $index++) {
    if ($headingIndex -lt 0 -and $lines[$index].Trim() -match '^```') {
        $hadFenceBeforeHeading = $true
    }

    if ($lines[$index].Trim() -eq $ExpectedHeading) {
        $headingIndex = $index
        break
    }
}

if ($headingIndex -ge 0) {
    $bodyLines = @($lines[$headingIndex..($lines.Count - 1)])
}
else {
    $bodyLines = Remove-LeadingNoise -Lines $lines
    $bodyLines = @($ExpectedHeading, '') + $bodyLines
}

$sanitized = Remove-LeadingNoise -Lines $bodyLines
if ($sanitized.Count -eq 0 -or $sanitized[0].Trim() -ne $ExpectedHeading) {
    $tail = if ($sanitized.Count -gt 0) { @('') + $sanitized } else { @() }
    $sanitized = @($ExpectedHeading) + $tail
}

if ($hadFenceBeforeHeading) {
    while ($sanitized.Count -gt 0 -and $sanitized[$sanitized.Count - 1].Trim() -match '^```') {
        $sanitized = if ($sanitized.Count -eq 1) { @() } else { @($sanitized[0..($sanitized.Count - 2)]) }
    }

    while ($sanitized.Count -gt 0 -and [string]::IsNullOrWhiteSpace($sanitized[$sanitized.Count - 1])) {
        $sanitized = if ($sanitized.Count -eq 1) { @() } else { @($sanitized[0..($sanitized.Count - 2)]) }
    }
}

$content = [string]::Join([Environment]::NewLine, $sanitized)
if (-not $content.EndsWith([Environment]::NewLine, [System.StringComparison]::Ordinal)) {
    $content += [Environment]::NewLine
}

[System.IO.File]::WriteAllText($resolvedPath, $content, [System.Text.UTF8Encoding]::new($false))
Write-Output "sanitize-report: normalized '$resolvedPath'"
exit 0
