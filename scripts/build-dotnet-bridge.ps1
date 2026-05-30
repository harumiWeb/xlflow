param(
	[switch]$InstallToGoBin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root 'bridge/dotnet/src/Xlflow.ExcelBridge/Xlflow.ExcelBridge.csproj'
$bridgeExe = Join-Path $root 'bridge/dotnet/src/Xlflow.ExcelBridge/bin/Release/net8.0/xlflow-excel-bridge.exe'

dotnet build $project --configuration Release

if (-not $InstallToGoBin) {
	return
}

if (-not (Test-Path $bridgeExe)) {
	throw "Bridge executable was not produced: $bridgeExe"
}

$goBin = (go env GOBIN).Trim()
if ([string]::IsNullOrWhiteSpace($goBin)) {
	$goPath = (go env GOPATH).Trim()
	if ([string]::IsNullOrWhiteSpace($goPath)) {
		throw 'Unable to resolve GOBIN/GOPATH from go env.'
	}

	$primaryGoPath = $goPath.Split([IO.Path]::PathSeparator)[0]
	$goBin = Join-Path $primaryGoPath 'bin'
}

New-Item -ItemType Directory -Path $goBin -Force | Out-Null
$destination = Join-Path $goBin 'xlflow-excel-bridge.exe'
Copy-Item -Path $bridgeExe -Destination $destination -Force
Write-Output "Installed xlflow-excel-bridge.exe to $destination"
