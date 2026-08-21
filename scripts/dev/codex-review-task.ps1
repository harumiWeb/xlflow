[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$reviewScript = Join-Path $PSScriptRoot "codex-review.ps1"
$reviewArgumentsJson = $env:XLFLOW_CODEX_REVIEW_ARGS_JSON
$reviewArguments = @()

if (-not [string]::IsNullOrWhiteSpace($reviewArgumentsJson)) {
    $parsedArguments = ConvertFrom-Json -InputObject $reviewArgumentsJson

    if ($null -ne $parsedArguments -and $parsedArguments -isnot [array]) {
        throw "Codex review task arguments must be a JSON array."
    }

    foreach ($argument in $parsedArguments) {
        $reviewArguments += [string]$argument
    }
}

$valueParameters = @(
    "Base",
    "Model",
    "Effort",
    "HeartbeatSeconds",
    "TimeoutSeconds",
    "ReviewMode"
)
$reviewParameterMap = @{}

for ($index = 0; $index -lt $reviewArguments.Count; $index++) {
    $token = $reviewArguments[$index]

    if ($token -notmatch "^-([^=]+)(?:=(.*))?$") {
        throw "Unsupported Codex review task argument '$token'."
    }

    $parameterName = $Matches[1].TrimStart("-")
    $hasInlineValue = $null -ne $Matches[2]

    if ($parameterName -eq "SkipFetch") {
        if ($hasInlineValue) {
            $reviewParameterMap["SkipFetch"] = [System.Convert]::ToBoolean($Matches[2])
        } else {
            $reviewParameterMap["SkipFetch"] = $true
        }

        continue
    }

    if ($parameterName -notin $valueParameters) {
        throw "Unsupported Codex review task parameter '-$parameterName'."
    }

    if ($hasInlineValue) {
        $reviewParameterMap[$parameterName] = $Matches[2]
        continue
    }

    if ($index + 1 -ge $reviewArguments.Count) {
        throw "Missing value for Codex review task parameter '-$parameterName'."
    }

    $index++
    $reviewParameterMap[$parameterName] = $reviewArguments[$index]
}

& $reviewScript @reviewParameterMap
$exitCode = if ($null -ne $LASTEXITCODE) {
    [int]$LASTEXITCODE
} else {
    0
}

exit $exitCode
