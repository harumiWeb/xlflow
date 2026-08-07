param(
    [string] $Manifest = "",
    [string] $CorpusRoot = "",
    [string] $UpstreamCheckout = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Resolve defaults from the repository root so this script is safe to invoke
# from any working directory. The manifest remains the only source of the
# upstream repository and pinned commit; this wrapper only forwards paths.
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
if ([string]::IsNullOrWhiteSpace($Manifest)) {
    $Manifest = Join-Path $repoRoot "testdata\static-analysis-corpus\manifest.json"
}
if ([string]::IsNullOrWhiteSpace($CorpusRoot)) {
    $CorpusRoot = Split-Path -Parent $Manifest
}

$manifestPath = (Resolve-Path -LiteralPath $Manifest).Path
$corpusPath = (Resolve-Path -LiteralPath $CorpusRoot).Path
$goScript = Join-Path $repoRoot "scripts\dev\go.ps1"
if (-not (Test-Path -LiteralPath $goScript)) {
    throw "Cannot find the repository Go wrapper: $goScript"
}

$goArgs = @(
    "run",
    "./cmd/xlflow-static-analysis-corpus",
    "sync",
    "--manifest", $manifestPath,
    "--corpus-root", $corpusPath
)
if (-not [string]::IsNullOrWhiteSpace($UpstreamCheckout)) {
    $checkoutPath = (Resolve-Path -LiteralPath $UpstreamCheckout).Path
    $goArgs += @("--upstream-checkout", $checkoutPath)
}

$global:LASTEXITCODE = 0
$exitCode = 1
Push-Location -LiteralPath $repoRoot
try {
    & $goScript @goArgs
    $exitCode = if ($null -ne $global:LASTEXITCODE) { [int]$global:LASTEXITCODE } else { 0 }
}
finally {
    Pop-Location
}
exit $exitCode
