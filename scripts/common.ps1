function New-XlflowResult {
  param([string]$Command)
  return [ordered]@{
    status = "ok"
    command = $Command
    error = $null
    logs = @()
  }
}

function Set-XlflowError {
  param(
    [hashtable]$Result,
    [string]$Code,
    [string]$Message,
    [string]$Source = "",
    [int]$Number = 0
  )
  $Result.status = "failed"
  $Result.error = [ordered]@{
    code = $Code
    message = $Message
    source = $Source
    number = $Number
  }
}

function ConvertTo-XlflowBool {
  param([string]$Value)
  return $Value -eq "true" -or $Value -eq "True" -or $Value -eq "1"
}

function Close-XlflowCom {
  param($Workbook, $Excel, [bool]$Save)
  if ($null -ne $Workbook) {
    try { $Workbook.Close($Save) | Out-Null } catch {}
  }
  if ($null -ne $Excel) {
    try { $Excel.Quit() | Out-Null } catch {}
  }
  if ($null -ne $Workbook) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch {}
  }
  if ($null -ne $Excel) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Get-XlflowComponentPath {
  param($Component, [string]$ModulesDir, [string]$ClassesDir, [string]$FormsDir, [string]$WorkbookDir)
  $name = $Component.Name
  switch ($Component.Type) {
    1 { return Join-Path $ModulesDir ($name + ".bas") }
    2 { return Join-Path $ClassesDir ($name + ".cls") }
    3 { return Join-Path $FormsDir ($name + ".frm") }
    100 { return Join-Path $WorkbookDir ($name + ".bas") }
    default { return $null }
  }
}

function Write-XlflowJson {
  param([hashtable]$Result)
  $Result | ConvertTo-Json -Depth 8
}

function Get-XlflowDocumentModuleContent {
  param([string]$Path)

  $lines = Get-Content -LiteralPath $Path
  $filtered = New-Object System.Collections.Generic.List[string]
  $inClassHeader = $false
  $classHeaderBuffer = New-Object System.Collections.Generic.List[string]

  foreach ($line in $lines) {
    $trimmed = $line.Trim()
    if ($trimmed -eq "VERSION 1.0 CLASS") {
      $inClassHeader = $true
      $classHeaderBuffer.Clear()
      $classHeaderBuffer.Add($line)
      continue
    }
    if ($inClassHeader) {
      $classHeaderBuffer.Add($line)
      if ($trimmed -eq "END") {
        $inClassHeader = $false
        $classHeaderBuffer.Clear()
      }
      continue
    }
    if ($trimmed -match '^Attribute\s+VB_') {
      continue
    }
    $filtered.Add($line)
  }

  if ($inClassHeader -and $classHeaderBuffer.Count -gt 0) {
    foreach ($headerLine in $classHeaderBuffer) {
      $filtered.Add($headerLine)
    }
  }

  $hasOptionExplicit = $false
  $hasNonHeaderCode = $false
  foreach ($line in $filtered) {
    $trimmed = $line.Trim()
    if ($trimmed -eq "") {
      continue
    }
    if ($trimmed -ieq "Option Explicit") {
      $hasOptionExplicit = $true
      continue
    }
    $hasNonHeaderCode = $true
  }

  if (-not $hasOptionExplicit -and -not $hasNonHeaderCode) {
    $filtered.Add("")
    $filtered.Add("Option Explicit")
  }

  return ($filtered -join [Environment]::NewLine)
}

function Normalize-XlflowDocumentModuleFile {
  param([string]$Path)

  $content = Get-XlflowDocumentModuleContent -Path $Path
  Set-Content -LiteralPath $Path -Value $content
}

function Sync-XlflowDocumentModule {
  param($Component, [string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) {
    return $false
  }

  $code = Get-XlflowDocumentModuleContent -Path $Path
  $module = $Component.CodeModule
  $lineCount = $module.CountOfLines

  if ($lineCount -gt 0) {
    $module.DeleteLines(1, $lineCount)
  }

  if (-not [string]::IsNullOrWhiteSpace($code)) {
    $module.AddFromString($code)
  }

  return $true
}
