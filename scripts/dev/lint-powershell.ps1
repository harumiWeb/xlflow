param()

$ErrorActionPreference = "Stop"

function Get-XlflowPSScriptAnalyzerModulePath {
  $module = Get-Module -ListAvailable PSScriptAnalyzer | Sort-Object Version -Descending | Select-Object -First 1
  if ($null -ne $module) {
    return $module.Path
  }

  $searchBases = New-Object System.Collections.Generic.List[string]
  $documentsDir = [Environment]::GetFolderPath("MyDocuments")
  if (-not [string]::IsNullOrWhiteSpace($documentsDir)) {
    $searchBases.Add((Join-Path $documentsDir "PowerShell\\Modules")) | Out-Null
    $searchBases.Add((Join-Path $documentsDir "WindowsPowerShell\\Modules")) | Out-Null
  }
  if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
    $searchBases.Add((Join-Path $env:ProgramFiles "PowerShell\\Modules")) | Out-Null
    $searchBases.Add((Join-Path $env:ProgramFiles "WindowsPowerShell\\Modules")) | Out-Null
  }

  $manifests = New-Object System.Collections.Generic.List[object]
  foreach ($base in ($searchBases | Select-Object -Unique)) {
    if ([string]::IsNullOrWhiteSpace($base)) {
      continue
    }
    $moduleRoot = Join-Path $base "PSScriptAnalyzer"
    if (-not (Test-Path -LiteralPath $moduleRoot)) {
      continue
    }
    foreach ($manifest in @(Get-ChildItem -LiteralPath $moduleRoot -Filter "PSScriptAnalyzer.psd1" -Recurse -File -ErrorAction SilentlyContinue)) {
      $manifests.Add($manifest) | Out-Null
    }
  }

  if ($manifests.Count -eq 0) {
    return $null
  }

  return (
    $manifests |
      Sort-Object @{ Expression = {
        try {
          [version]$_.Directory.Name
        } catch {
          [version]"0.0"
        }
      }; Descending = $true }, FullName |
      Select-Object -ExpandProperty FullName -First 1
  )
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$settingsPath = Join-Path $repoRoot "PSScriptAnalyzerSettings.psd1"

Set-Location $repoRoot

$analyzerModulePath = Get-XlflowPSScriptAnalyzerModulePath
if ([string]::IsNullOrWhiteSpace($analyzerModulePath)) {
  throw "PSScriptAnalyzer is not installed or not visible to this PowerShell session."
}

Import-Module $analyzerModulePath -ErrorAction Stop

$trackedFiles = @(git --no-pager ls-files -- "*.ps1")
if ($LASTEXITCODE -ne 0) {
  throw "failed to enumerate PowerShell sources with git ls-files"
}

$files = @(
  $trackedFiles |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    ForEach-Object { [System.IO.Path]::GetFullPath((Join-Path $repoRoot $_)) }
)

if ($files.Count -eq 0) {
  Write-Output "No PowerShell sources found."
  exit 0
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($file in $files) {
  foreach ($finding in @(Invoke-ScriptAnalyzer -Path $file -Settings $settingsPath)) {
    $results.Add($finding) | Out-Null
  }
}

if ($results.Count -eq 0) {
  Write-Output ("PSScriptAnalyzer passed for " + $files.Count + " file(s).")
  exit 0
}

$results |
  Sort-Object ScriptName, Line, RuleName |
  Select-Object ScriptName, RuleName, Severity, Line, Message |
  Format-Table -AutoSize |
  Out-String -Width 220 |
  Write-Output

exit 1
