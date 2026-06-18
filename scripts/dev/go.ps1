param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $GoArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ucrtBin = "C:\msys64\ucrt64\bin"
$gcc = Join-Path $ucrtBin "gcc.exe"
$gxx = Join-Path $ucrtBin "g++.exe"

if (Test-Path -LiteralPath $gcc) {
    $pathParts = @($ucrtBin)
    if (-not [string]::IsNullOrWhiteSpace($env:PATH)) {
        $pathParts += $env:PATH
    }
    $env:PATH = [string]::Join([IO.Path]::PathSeparator, $pathParts)
    $env:CC = $gcc
    if (Test-Path -LiteralPath $gxx) {
        $env:CXX = $gxx
    }
}

& go @GoArgs
exit $LASTEXITCODE
