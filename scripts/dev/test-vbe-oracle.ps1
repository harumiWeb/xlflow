param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $OracleArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($env:OS -ne "Windows_NT") {
    throw "The VBE oracle is Windows-only and requires a local Excel installation."
}

$goScript = Join-Path $PSScriptRoot "go.ps1"
if (-not (Test-Path -LiteralPath $goScript)) {
    throw "Cannot find the repository Go wrapper: $goScript"
}

# Keep the oracle outside the public xlflow command tree. One invocation owns
# one sequential Excel/VBE run; do not start this script concurrently.
& $goScript run ./cmd/xlflow-vbe-oracle @OracleArgs
exit $LASTEXITCODE
