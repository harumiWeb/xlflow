param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$ExpectedHeading
)

$ErrorActionPreference = 'Stop'

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
