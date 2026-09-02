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
        "uncommitted",
        "-Commit",
        "HEAD"
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

function Test-ReviewScopeSelection {
    param(
        [Parameter(Mandatory)]
        [string]$TempRoot
    )

    $reviewScript = Join-Path $PSScriptRoot "codex-review.ps1"
    $probeScript = Join-Path $TempRoot "review-scope-probe.ps1"
    $shimPath = Join-Path $TempRoot "codex.cmd"
    $previousPath = $env:PATH
    $outputEnvironmentPath = "Env:XLFLOW_CODEX_REVIEW_SCOPE_OUTPUT"
    $previousOutputEnvironmentItem = Get-Item `
        -Path $outputEnvironmentPath `
        -ErrorAction SilentlyContinue
    $hadPreviousOutputEnvironment = $null -ne $previousOutputEnvironmentItem
    $previousOutputEnvironmentValue = if ($hadPreviousOutputEnvironment) {
        [string]$previousOutputEnvironmentItem.Value
    } else {
        $null
    }

    Set-Content -LiteralPath $probeScript -Encoding utf8 -Value @'
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

Set-Content `
    -LiteralPath $env:XLFLOW_CODEX_REVIEW_SCOPE_OUTPUT `
    -Value ($Arguments -join "`n") `
    -Encoding utf8

for ($index = 0; $index + 1 -lt $Arguments.Count; $index++) {
    if ($Arguments[$index] -eq "--output-last-message") {
        Set-Content `
            -LiteralPath $Arguments[$index + 1] `
            -Value "review scope probe" `
            -Encoding utf8
        break
    }
}
'@

    Set-Content -LiteralPath $shimPath -Encoding ascii -Value @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "$probeScript" %*
"@

    $powershellCommand = Get-Command powershell -ErrorAction Stop |
        Select-Object -First 1
    $repoRoot = (& git rev-parse --show-toplevel 2>$null).Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($repoRoot)) {
        throw "Review scope probe is not running inside a Git repository."
    }

    function Invoke-ReviewScopeProbe {
        param(
            [Parameter(Mandatory)]
            [string]$OutputFile,
            [Parameter(Mandatory)]
            [string[]]$ReviewArguments
        )

        $stdoutFile = Join-Path $TempRoot "review-scope.stdout.log"
        $stderrFile = Join-Path $TempRoot "review-scope.stderr.log"
        $processArguments = @(
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            $reviewScript
        ) + $ReviewArguments
        $processArgumentString = (
            $processArguments |
                ForEach-Object {
                    ConvertTo-ProcessArgument -Value ([string]$_)
                }
        ) -join " "

        Set-Item -Path $outputEnvironmentPath -Value $OutputFile
        $process = Start-Process `
            -FilePath $powershellCommand.Source `
            -ArgumentList $processArgumentString `
            -WorkingDirectory $repoRoot `
            -RedirectStandardOutput $stdoutFile `
            -RedirectStandardError $stderrFile `
            -WindowStyle Hidden `
            -PassThru `
            -Wait

        if ($process.ExitCode -ne 0) {
            $stderr = if (Test-Path -LiteralPath $stderrFile) {
                Get-Content -LiteralPath $stderrFile -Raw
            } else {
                ""
            }
            throw (
                "Review scope probe failed with exit code {0}: {1}" -f
                    $process.ExitCode,
                    $stderr
            )
        }

        $reviewOutput = Get-Content -LiteralPath $stdoutFile -Raw
        if ($reviewOutput -notmatch "review scope probe") {
            throw "Review scope probe did not produce a final review message."
        }

        return @(Get-Content -LiteralPath $OutputFile -Encoding utf8)
    }

    try {
        $env:PATH = "{0};{1}" -f $TempRoot, $previousPath

        $baseOutputFile = Join-Path $TempRoot "review-base.args"
        $baseReviewArguments = @(Invoke-ReviewScopeProbe `
            -OutputFile $baseOutputFile `
            -ReviewArguments @(
                "-SkipFetch",
                "-HeartbeatSeconds",
                "5",
                "-TimeoutSeconds",
                "30",
                "-ReviewMode",
                "base",
                "-Base",
                "HEAD~1"
            ))
        $baseIndex = -1
        for ($index = 0; $index -lt $baseReviewArguments.Count; $index++) {
            if ($baseReviewArguments[$index] -eq "--base") {
                $baseIndex = $index
                break
            }
        }
        if ($baseIndex -lt 0 -or $baseReviewArguments[$baseIndex + 1] -ne "HEAD~1") {
            throw (
                "Base review arguments did not select HEAD~1: {0}" -f
                    ($baseReviewArguments -join "|")
            )
        }

        $head = (& git rev-parse --verify "HEAD^{commit}").Trim()
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($head)) {
            throw "Review scope probe could not resolve HEAD."
        }
        $commitOutputFile = Join-Path $TempRoot "review-commit.args"
        $commitReviewArguments = @(Invoke-ReviewScopeProbe `
            -OutputFile $commitOutputFile `
            -ReviewArguments @(
                "-SkipFetch",
                "-HeartbeatSeconds",
                "5",
                "-TimeoutSeconds",
                "30",
                "-ReviewMode",
                "commit",
                "-Commit",
                "HEAD"
            ))
        $commitIndex = -1
        for ($index = 0; $index -lt $commitReviewArguments.Count; $index++) {
            if ($commitReviewArguments[$index] -eq "--commit") {
                $commitIndex = $index
                break
            }
        }
        if ($commitIndex -lt 0 -or $commitReviewArguments[$commitIndex + 1] -ne $head) {
            throw (
                "Commit review arguments did not select HEAD ({0}): {1}" -f
                    $head,
                    ($commitReviewArguments -join "|")
            )
        }
        if ($commitReviewArguments -contains "--base" -or
            $commitReviewArguments -contains "--uncommitted") {
            throw "Commit review unexpectedly included another review scope."
        }
    } finally {
        $env:PATH = $previousPath
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
    Test-ReviewScopeSelection -TempRoot $tempRoot

    Write-Output (
        "Codex review argument encoding, Task transport, and scope selection passed for {0} case(s)." -f
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
