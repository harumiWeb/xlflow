[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($env:OS -ne "Windows_NT") {
    throw "The Codex review argument regression test is Windows-only."
}

. (Join-Path $PSScriptRoot "codex-review-arguments.ps1")

$percentVariableName = "XLFLOW_CODEX_REVIEW_TEST_LITERAL_PERCENT"
$percentEnvironmentPath = "Env:{0}" -f $percentVariableName
$previousEnvironmentItem = Get-Item `
    -Path $percentEnvironmentPath `
    -ErrorAction SilentlyContinue
$hadPreviousEnvironment = $null -ne $previousEnvironmentItem
$previousEnvironmentValue = if ($hadPreviousEnvironment) {
    [string]$previousEnvironmentItem.Value
} else {
    $null
}

$tempRoot = Join-Path `
    ([System.IO.Path]::GetTempPath()) `
    ("xlflow-codex-review-argument-test-" + [Guid]::NewGuid().ToString("N"))
$probeScript = Join-Path $tempRoot "probe.ps1"
$shimPath = Join-Path $tempRoot "codex.cmd"
$outputFile = Join-Path $tempRoot "arguments.txt"

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    Set-Content -LiteralPath $probeScript -Encoding utf8 -Value @'
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

Set-Content `
    -LiteralPath $env:XLFLOW_CODEX_REVIEW_TEST_OUTPUT `
    -Value ($Arguments -join "`n") `
    -Encoding utf8
'@

    Set-Content -LiteralPath $shimPath -Encoding ascii -Value @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "$probeScript" %*
"@

    Set-Item -Path $percentEnvironmentPath -Value "%"
    $env:XLFLOW_CODEX_REVIEW_TEST_OUTPUT = $outputFile

    $cases = @(
        "origin/feature&review",
        "origin/feature%PATH%review",
        "origin/feature!review",
        "origin/feature|review",
        "origin/feature<review",
        "origin/feature>review",
        "origin/feature(review)",
        "origin/feature^review"
    )

    foreach ($case in $cases) {
        $encoded = ConvertTo-CmdProcessArgument `
            -Value $case `
            -PercentVariableName $percentVariableName

        $process = Start-Process `
            -FilePath $shimPath `
            -ArgumentList $encoded `
            -WindowStyle Hidden `
            -PassThru `
            -Wait

        if ($process.ExitCode -ne 0) {
            throw "cmd.exe argument probe failed for '$case' with exit code $($process.ExitCode)."
        }

        $actual = @(Get-Content -LiteralPath $outputFile -Encoding utf8)

        if ($actual.Count -ne 1 -or $actual[0] -ne $case) {
            throw "cmd.exe changed '$case' into '$($actual -join "|")'."
        }
    }

    Write-Output ("Codex review argument encoding passed for {0} case(s)." -f $cases.Count)
} finally {
    if ($hadPreviousEnvironment) {
        Set-Item -Path $percentEnvironmentPath -Value $previousEnvironmentValue
    } else {
        Remove-Item `
            -Path $percentEnvironmentPath `
            -Force `
            -ErrorAction SilentlyContinue
    }

    Remove-Item Env:XLFLOW_CODEX_REVIEW_TEST_OUTPUT -ErrorAction SilentlyContinue
    Remove-Item `
        -LiteralPath $tempRoot `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue
}
