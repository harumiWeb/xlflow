param(
  [string]$LicencesPath = ".\THIRD_PARTY_LICENCES.md",
  [string]$Package = "./cmd/xlflow"
)

$ErrorActionPreference = "Stop"

# This check intentionally validates the dependency closure of ./cmd/xlflow,
# which matches the published licence inventory contract in THIRD_PARTY_LICENCES.md.

function Get-TrackedModules {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "licence inventory not found: $Path"
  }

  $pattern = '^\|\s*`(?<module>[^`]+)`\s*\|\s*`(?<version>[^`]+)`\s*\|'
  $tracked = @{}
  foreach ($line in Get-Content -LiteralPath $Path) {
    if ($line -match $pattern) {
      $module = $Matches["module"].Trim()
      $version = $Matches["version"].Trim()
      if ([string]::IsNullOrWhiteSpace($module) -or [string]::IsNullOrWhiteSpace($version)) {
        continue
      }
      if ($tracked.ContainsKey($module)) {
        throw "duplicate module entry in licence inventory: $module"
      }
      $tracked[$module] = $version
    }
  }
  return $tracked
}

function Get-ExpectedModules {
  param([string]$TargetPackage)

  $lines = go list -deps -f "{{if and .Module (not .Module.Main)}}{{.Module.Path}}`t{{.Module.Version}}{{end}}" $TargetPackage
  if ($LASTEXITCODE -ne 0) {
    throw "go list failed for $TargetPackage"
  }

  $expected = @{}
  foreach ($line in $lines) {
    if ([string]::IsNullOrWhiteSpace($line)) {
      continue
    }
    $parts = $line -split "`t", 2
    if ($parts.Count -ne 2) {
      throw "unexpected go list output: $line"
    }
    $module = $parts[0].Trim()
    $version = $parts[1].Trim()
    if ([string]::IsNullOrWhiteSpace($module) -or [string]::IsNullOrWhiteSpace($version)) {
      continue
    }
    $expected[$module] = $version
  }
  return $expected
}

$tracked = Get-TrackedModules -Path $LicencesPath
$expected = Get-ExpectedModules -TargetPackage $Package

$missing = New-Object System.Collections.Generic.List[string]
$mismatched = New-Object System.Collections.Generic.List[string]
$extra = New-Object System.Collections.Generic.List[string]

foreach ($module in ($expected.Keys | Sort-Object)) {
  if (-not $tracked.ContainsKey($module)) {
    $missing.Add("$module $($expected[$module])")
    continue
  }
  if ($tracked[$module] -ne $expected[$module]) {
    $mismatched.Add("$module expected=$($expected[$module]) actual=$($tracked[$module])")
  }
}

foreach ($module in ($tracked.Keys | Sort-Object)) {
  if (-not $expected.ContainsKey($module)) {
    $extra.Add("$module $($tracked[$module])")
  }
}

if ($missing.Count -gt 0 -or $mismatched.Count -gt 0 -or $extra.Count -gt 0) {
  if ($missing.Count -gt 0) {
    Write-Error ("missing licence inventory entries:`n" + ($missing -join "`n"))
  }
  if ($mismatched.Count -gt 0) {
    Write-Error ("mismatched licence inventory versions:`n" + ($mismatched -join "`n"))
  }
  if ($extra.Count -gt 0) {
    Write-Error ("stale licence inventory entries:`n" + ($extra -join "`n"))
  }
  exit 1
}

Write-Output ("licence inventory matches " + $expected.Count + " modules for " + $Package)
