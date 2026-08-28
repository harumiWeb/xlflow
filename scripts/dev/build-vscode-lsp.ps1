[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$commitOutput = (& git -C $repoRoot rev-parse HEAD | Out-String)
if ($LASTEXITCODE -ne 0) {
    throw "Could not determine the current Git commit for the debug LSP build."
}
$commit = $commitOutput.Trim()
if ([string]::IsNullOrWhiteSpace($commit)) {
    throw "Could not determine the current Git commit for the debug LSP build."
}
$commitDateOutput = (& git -C $repoRoot log -1 --format=%cI HEAD | Out-String)
if ($LASTEXITCODE -ne 0) {
    throw "Could not determine the current Git commit date for the debug LSP build."
}
$commitDate = $commitDateOutput.Trim()
if ([string]::IsNullOrWhiteSpace($commitDate)) {
    throw "Could not determine the current Git commit date for the debug LSP build."
}
$workingTree = (& git -C $repoRoot status --porcelain | Out-String).Trim()
$commitLabel = if ([string]::IsNullOrWhiteSpace($workingTree)) { $commit } else { "$commit-dirty" }
$outputCandidate = if ([IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath
}
else {
    Join-Path $repoRoot $OutputPath
}
$resolvedOutputPath = [IO.Path]::GetFullPath($outputCandidate)
$outputDirectory = Split-Path -Parent $resolvedOutputPath
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$goScript = Join-Path $PSScriptRoot "go.ps1"
$ldFlags = "-X main.commit=$commitLabel -X main.date=$commitDate"
$goArgs = @("build", "-ldflags", $ldFlags, "-o", $resolvedOutputPath, "./cmd/xlflow")
Push-Location $repoRoot
try {
    & $goScript -GoArgs $goArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Go LSP build failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}

if (-not (Test-Path -LiteralPath $resolvedOutputPath -PathType Leaf)) {
    throw "Go LSP build did not produce '$resolvedOutputPath'."
}

$treeState = if ([string]::IsNullOrWhiteSpace($workingTree)) { "clean" } else { "with local changes" }
Write-Output "Built xlflow LSP for Extension Host: $resolvedOutputPath (commit $commitLabel, $treeState)"
