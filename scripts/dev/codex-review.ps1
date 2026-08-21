[CmdletBinding()]
param(
    [string]$Base = "origin/main",
    [string]$Model = "gpt-5.6-luna",

    [ValidateSet("low", "medium", "high", "xhigh", "max")]
    [string]$Effort = "xhigh",

    [ValidateRange(5, 300)]
    [int]$HeartbeatSeconds = 15,

    [ValidateRange(30, 86400)]
    [int]$TimeoutSeconds = 1800,

    [switch]$SkipFetch,

    [ValidateSet("auto", "base", "uncommitted")]
    [string]$ReviewMode = "auto"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

. (Join-Path $PSScriptRoot "codex-review-arguments.ps1")

function Write-Heartbeat {
    param(
        [string]$Glyph,
        [TimeSpan]$Elapsed
    )

    $text = "{0} Codex review running... {1:hh\:mm\:ss}" -f $Glyph, $Elapsed
    [Console]::Error.WriteLine($text)
    [Console]::Error.Flush()
}

function Write-LogTail {
    param(
        [string[]]$Path,
        [int]$Lines = 80
    )

    foreach ($logPath in $Path) {
        if (-not (Test-Path -LiteralPath $logPath)) {
            continue
        }

        [Console]::Error.WriteLine("")
        [Console]::Error.WriteLine(
            "--- Codex diagnostic log (tail): {0} ---" -f $logPath
        )

        foreach ($line in Get-Content -LiteralPath $logPath -Tail $Lines) {
            [Console]::Error.WriteLine($line)
        }
    }
}

function Stop-CodexProcess {
    param(
        [int]$ProcessId
    )

    if ($ProcessId -le 0) {
        return
    }

    $taskKillCommand = Get-Command taskkill.exe -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if ($null -ne $taskKillCommand) {
        $taskKillPath = $taskKillCommand.Source
        & $taskKillPath /PID $processId /T /F *> $null
        return
    }

    Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
}

$repoRoot = (& git rev-parse --show-toplevel 2>$null)

if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($repoRoot)) {
    throw "Not inside a Git repository."
}

$repoRoot = $repoRoot.Trim()
Set-Location -LiteralPath $repoRoot

$worktreeStatus = @(git status --porcelain=v1 --untracked-files=all)

if ($LASTEXITCODE -ne 0) {
    throw "Failed to inspect the Git worktree."
}

# Keep origin/main current before reviewing. Do this quietly because successful
# setup output is not useful to the calling agent. An explicit uncommitted-only
# review does not need the base reference.
if (
    $ReviewMode -ne "uncommitted" -and
    -not $SkipFetch -and
    $Base -eq "origin/main"
) {
    & git fetch --quiet origin main

    if ($LASTEXITCODE -ne 0) {
        throw "Failed to update origin/main before Codex review."
    }
}

$hasUncommittedChanges = $worktreeStatus.Count -gt 0

if (-not $hasUncommittedChanges) {
    if ($ReviewMode -eq "uncommitted") {
        $reviewModeArguments = @("--uncommitted")
    } else {
        $reviewModeArguments = @("--base", $Base)
    }
} elseif ($ReviewMode -eq "uncommitted") {
    $reviewModeArguments = @("--uncommitted")
} elseif ($ReviewMode -eq "base") {
    [Console]::Error.WriteLine(
        "Codex review is using base mode; uncommitted worktree changes are excluded."
    )
    $reviewModeArguments = @("--base", $Base)
} else {
    & git diff --quiet ("{0}...HEAD" -f $Base) --
    $baseDiffExitCode = $LASTEXITCODE

    if ($baseDiffExitCode -gt 1) {
        throw "Failed to compare HEAD with review base '$Base'."
    }

    if ($baseDiffExitCode -eq 1) {
        throw (
            "The worktree has uncommitted changes and committed changes relative to " +
            "'$Base'. Commit or stash one scope first, or select " +
            "-ReviewMode base or -ReviewMode uncommitted explicitly."
        )
    }

    $reviewModeArguments = @("--uncommitted")
}

# Prefer concrete command shims so the background worker does not have to
# rediscover Codex through a different shell environment.
$codexPath = $null

foreach ($commandName in @(
    "codex.cmd",
    "codex.exe",
    "codex.ps1",
    "codex"
)) {
    $command = Get-Command $commandName -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
        $codexPath = $command.Source
        break
    }
}

if ([string]::IsNullOrWhiteSpace($codexPath)) {
    throw "Codex CLI was not found on PATH."
}

$cmdLauncher = [System.IO.Path]::GetExtension($codexPath) -ieq ".cmd"
$cmdLiteralPercentVariableName = "XLFLOW_CODEX_REVIEW_LITERAL_PERCENT"

$tempRoot = Join-Path `
    ([System.IO.Path]::GetTempPath()) `
    ("xlflow-codex-review-" + [Guid]::NewGuid().ToString("N"))

$resultFile = Join-Path $tempRoot "final-review.md"
$logFile = Join-Path $tempRoot "codex.log"
$errorLogFile = Join-Path $tempRoot "codex.stderr.log"
$logPaths = @($logFile, $errorLogFile)

$launcherPath = $codexPath
$launcherArguments = @()

if ([System.IO.Path]::GetExtension($codexPath) -ieq ".ps1") {
    $launcherCommand = $null

    foreach ($launcherName in @("pwsh", "powershell")) {
        $launcherCommand = Get-Command $launcherName -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if ($null -ne $launcherCommand) {
            break
        }
    }

    if ($null -eq $launcherCommand) {
        throw "A PowerShell launcher is required to run codex.ps1."
    }

    $launcherPath = $launcherCommand.Source
    $launcherArguments = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        $codexPath
    )
}

$codexReviewArguments = @(
    "exec",
    "--ephemeral",
    "-m",
    $Model,
    "-s",
    "read-only",
    "-c",
    ('model_reasoning_effort="' + $Effort + '"'),
    "--output-last-message",
    $resultFile,
    "review"
)
$codexReviewArguments += $reviewModeArguments

$processArguments = @($launcherArguments) + @($codexReviewArguments)
$processArgumentString = (
    $processArguments |
        ForEach-Object {
            if ($cmdLauncher) {
                ConvertTo-CmdProcessArgument `
                    -Value ([string]$_) `
                    -PercentVariableName $cmdLiteralPercentVariableName
            } else {
                ConvertTo-ProcessArgument -Value ([string]$_)
            }
        }
) -join " "

New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

$reviewProcess = $null
$reviewProcessId = 0
$timedOut = $false
$cmdPercentEnvironmentSet = $false
$cmdPercentEnvironmentPath = "Env:{0}" -f $cmdLiteralPercentVariableName
$hadPreviousCmdPercentEnvironment = $false
$previousCmdPercentEnvironmentValue = $null

try {
    if ($cmdLauncher) {
        $previousEnvironmentItem = Get-Item `
            -Path $cmdPercentEnvironmentPath `
            -ErrorAction SilentlyContinue

        if ($null -ne $previousEnvironmentItem) {
            $hadPreviousCmdPercentEnvironment = $true
            $previousCmdPercentEnvironmentValue = [string]$previousEnvironmentItem.Value
        }

        Set-Item -Path $cmdPercentEnvironmentPath -Value "%"
        $cmdPercentEnvironmentSet = $true
    }

    # Run Codex off-screen. All normal Codex output, including reasoning
    # summaries, tool calls and diagnostics, is redirected to temporary logs.
    $reviewProcess = Start-Process `
        -FilePath $launcherPath `
        -ArgumentList $processArgumentString `
        -WorkingDirectory $repoRoot `
        -RedirectStandardOutput $logFile `
        -RedirectStandardError $errorLogFile `
        -WindowStyle Hidden `
        -PassThru
    $reviewProcessId = $reviewProcess.Id

    $spinner = @(
        "⠋",
        "⠙",
        "⠹",
        "⠸",
        "⠼",
        "⠴",
        "⠦",
        "⠧",
        "⠇",
        "⠏"
    )

    $spinnerIndex = 0
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $nextHeartbeat = [TimeSpan]::Zero
    $reviewTimeout = [TimeSpan]::FromSeconds($TimeoutSeconds)

    while (-not $reviewProcess.HasExited) {
        if ($watch.Elapsed -ge $reviewTimeout) {
            $timedOut = $true
            break
        }

        if ($watch.Elapsed -ge $nextHeartbeat) {
            Write-Heartbeat `
                -Glyph $spinner[$spinnerIndex % $spinner.Count] `
                -Elapsed $watch.Elapsed

            $spinnerIndex++
            $nextHeartbeat = $watch.Elapsed.Add(
                [TimeSpan]::FromSeconds($HeartbeatSeconds)
            )
        }

        Start-Sleep -Seconds 1
        $reviewProcess.Refresh()
    }

    $watch.Stop()

    if ($timedOut) {
        [Console]::Error.WriteLine(
            "Codex review timed out after {0:hh\:mm\:ss}." -f $watch.Elapsed
        )

        Stop-CodexProcess -ProcessId $reviewProcessId
        $null = $reviewProcess.WaitForExit(5000)
        Write-LogTail -Path $logPaths
        exit 124
    }

    $exitCode = [int]$reviewProcess.ExitCode

    if ($exitCode -ne 0) {
        [Console]::Error.WriteLine(
            "Codex review failed with exit code {0} after {1:hh\:mm\:ss}." -f `
                $exitCode,
                $watch.Elapsed
        )

        Write-LogTail -Path $logPaths

        exit $exitCode
    }

    if (-not (Test-Path -LiteralPath $resultFile)) {
        [Console]::Error.WriteLine(
            "Codex review succeeded but produced no final-message file."
        )

        Write-LogTail -Path $logPaths
        exit 1
    }

    $review = [System.IO.File]::ReadAllText($resultFile)

    if ([string]::IsNullOrWhiteSpace($review)) {
        [Console]::Error.WriteLine(
            "Codex review succeeded but the final review was empty."
        )

        Write-LogTail -Path $logPaths
        exit 1
    }

    [Console]::Error.WriteLine(
        "✓ Codex review finished in {0:hh\:mm\:ss}" -f $watch.Elapsed
    )

    # The ONLY semantic success output.
    [Console]::Out.WriteLine($review.TrimEnd())
}
finally {
    if ($null -ne $reviewProcess) {
        $reviewProcess.Refresh()

        if (-not $reviewProcess.HasExited) {
            Stop-CodexProcess -ProcessId $reviewProcessId
            $null = $reviewProcess.WaitForExit(5000)
        }
    }

    if ($cmdPercentEnvironmentSet) {
        if ($hadPreviousCmdPercentEnvironment) {
            Set-Item `
                -Path $cmdPercentEnvironmentPath `
                -Value $previousCmdPercentEnvironmentValue
        } else {
            Remove-Item `
                -Path $cmdPercentEnvironmentPath `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }

    Remove-Item `
        -LiteralPath $tempRoot `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue
}
