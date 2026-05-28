Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $root 'bridge/dotnet/Xlflow.ExcelBridge.sln'

dotnet test $solution --configuration Release
