[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($env:OS -ne "Windows_NT") {
    throw "The Codex review argument regression test is Windows-only."
}

. (Join-Path $PSScriptRoot "codex-review-arguments.ps1")

function Test-TaskArgumentTransport {
    param(
        [Parameter(Mandatory)]
        [string]$TempRoot
    )

    $taskCommand = Get-Command task -ErrorAction Stop |
        Select-Object -First 1
    $taskFile = Join-Path $TempRoot "Taskfile.transport.yml"
    $probeScript = Join-Path $TempRoot "task-transport-probe.ps1"
    $outputFile = Join-Path $TempRoot "task-transport.json"
    $outputEnvironmentPath = "Env:XLFLOW_CODEX_REVIEW_TASK_TEST_OUTPUT"
    $expected = @(
        "-Base",
        "origin/feature&review",
        "-Model",
        "gpt%PATH%model",
        "-ReviewMode",
        "uncommitted"
    )

    Set-Content -LiteralPath $probeScript -Encoding utf8 -Value @'
Set-Content `
    -LiteralPath $env:XLFLOW_CODEX_REVIEW_TASK_TEST_OUTPUT `
    -Value $env:XLFLOW_CODEX_REVIEW_ARGS_JSON `
    -Encoding utf8
'@

    Set-Content -LiteralPath $taskFile -Encoding utf8 -Value @"
version: "3"
tasks:
  probe:
    silent: true
    env:
      XLFLOW_CODEX_REVIEW_ARGS_JSON: '{{.CLI_ARGS_LIST | toJson}}'
    cmds:
      - cmd: powershell -NoProfile -ExecutionPolicy Bypass -File '$probeScript'
        platforms: [windows]
"@

    $previousOutputEnvironmentItem = Get-Item `
        -Path $outputEnvironmentPath `
        -ErrorAction SilentlyContinue
    $hadPreviousOutputEnvironment = $null -ne $previousOutputEnvironmentItem
    $previousOutputEnvironmentValue = if ($hadPreviousOutputEnvironment) {
        [string]$previousOutputEnvironmentItem.Value
    } else {
        $null
    }

    try {
        Set-Item -Path $outputEnvironmentPath -Value $outputFile
        $taskArguments = @("-t", $taskFile, "probe", "--") + $expected
        & $taskCommand.Source @taskArguments *> $null
        $exitCode = $LASTEXITCODE

        if ($exitCode -ne 0) {
            throw "Task argument transport probe failed with exit code $exitCode."
        }

        $raw = Get-Content -LiteralPath $outputFile -Raw -Encoding utf8
        $parsed = ConvertFrom-Json -InputObject $raw
        $actual = @()

        foreach ($argument in $parsed) {
            $actual += [string]$argument
        }

        if ($actual.Count -ne $expected.Count) {
            throw (
                "Task transported {0} arguments instead of {1}. Raw JSON: {2}" -f
                    $actual.Count,
                    $expected.Count,
                    $raw
            )
        }

        foreach ($index in 0..($expected.Count - 1)) {
            if ($actual[$index] -ne $expected[$index]) {
                throw "Task changed argument '$($expected[$index])' into '$($actual[$index])'."
            }
        }
    } finally {
        if ($hadPreviousOutputEnvironment) {
            Set-Item `
                -Path $outputEnvironmentPath `
                -Value $previousOutputEnvironmentValue
        } else {
            Remove-Item `
                -Path $outputEnvironmentPath `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
}

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
$outputEnvironmentPath = "Env:XLFLOW_CODEX_REVIEW_TEST_OUTPUT"
$previousOutputEnvironmentItem = Get-Item `
    -Path $outputEnvironmentPath `
    -ErrorAction SilentlyContinue
$hadPreviousOutputEnvironment = $null -ne $previousOutputEnvironmentItem
$previousOutputEnvironmentValue = if ($hadPreviousOutputEnvironment) {
    [string]$previousOutputEnvironmentItem.Value
} else {
    $null
}

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

    Test-TaskArgumentTransport -TempRoot $tempRoot

    Write-Output (
        "Codex review argument encoding and Task transport passed for {0} case(s)." -f
            $cases.Count
    )
} finally {
    if ($hadPreviousEnvironment) {
        Set-Item -Path $percentEnvironmentPath -Value $previousEnvironmentValue
    } else {
        Remove-Item `
            -Path $percentEnvironmentPath `
            -Force `
            -ErrorAction SilentlyContinue
    }

    if ($hadPreviousOutputEnvironment) {
        Set-Item `
            -Path $outputEnvironmentPath `
            -Value $previousOutputEnvironmentValue
    } else {
        Remove-Item `
            -Path $outputEnvironmentPath `
            -Force `
            -ErrorAction SilentlyContinue
    }
    Remove-Item `
        -LiteralPath $tempRoot `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue
}
