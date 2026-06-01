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
        throw "TAKT runs directory not found: $runsRoot"
    }

    $runningRun = Get-ChildItem -LiteralPath $runsRoot -Directory |
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
        throw 'No running TAKT workflow run found to resolve {report_dir}.'
    }

    return Join-Path (Split-Path -Parent $PSScriptRoot) $runningRun.Meta.reportDirectory
}

foreach ($prefix in @('{report_dir}/', '｛report_dir｝/')) {
    if ($Path.StartsWith($prefix, [System.StringComparison]::Ordinal)) {
        $reportDirectory = Resolve-TaktReportDirectory -ExplicitReportDirectory $ReportDirectory
        $relativeReportPath = $Path.Substring($prefix.Length)
        $Path = Join-Path $reportDirectory $relativeReportPath
        break
    }
}

if (-not (Test-Path -LiteralPath $Path)) {
    throw "Report file not found: $Path"
}

$lines = Get-Content -LiteralPath $Path
if ($lines.Count -eq 0) {
    throw "Report file is empty: $Path"
}

$firstNonEmptyIndex = -1
for ($index = 0; $index -lt $lines.Count; $index++) {
    if (-not [string]::IsNullOrWhiteSpace($lines[$index])) {
        $firstNonEmptyIndex = $index
        break
    }
}

if ($firstNonEmptyIndex -lt 0) {
    throw "Report file contains only whitespace: $Path"
}

$firstNonEmptyLine = $lines[$firstNonEmptyIndex].Trim()
if ($firstNonEmptyLine -ne $ExpectedHeading) {
    throw "Report must start with '$ExpectedHeading' but found '$firstNonEmptyLine' in $Path"
}

$bannedPrefixes = @(
    'The user wants me',
    'Let me',
    'I need to',
    'Now I',
    'I''m in',
    'Looking back',
    'Got it',
    'Understood'
)

$inspectionLimit = [Math]::Min($lines.Count - 1, $firstNonEmptyIndex + 24)
for ($index = $firstNonEmptyIndex + 1; $index -le $inspectionLimit; $index++) {
    $trimmed = $lines[$index].Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        continue
    }

    foreach ($prefix in $bannedPrefixes) {
        if ($trimmed.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Report contains meta commentary near the top: '$trimmed' in $Path"
        }
    }
}
