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
    $Result,
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

function Find-XlflowTestProcedures {
  param([string]$ModuleName, [string]$Code)

  $tests = New-Object System.Collections.Generic.List[object]
  if ([string]::IsNullOrEmpty($Code)) {
    return $tests
  }

  $lines = $Code -split "`r?`n"
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    $match = [regex]::Match($line, '^(?:Public\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(\s*\))?\s*(?:''.*)?$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
      continue
    }
    $name = $match.Groups[1].Value
    if ($name -like "Test*" -or $name -like "*_Test") {
      $tests.Add([pscustomobject][ordered]@{
        name = $name
        module = $ModuleName
        line = $i + 1
      })
    }
  }

  foreach ($test in $tests) {
    Write-Output $test
  }
}

function Select-XlflowTests {
  param($Tests, [string]$Filter = "")

  $selected = New-Object System.Collections.Generic.List[object]
  foreach ($test in $Tests) {
    if ([string]::IsNullOrWhiteSpace($Filter) -or $test.name -eq $Filter) {
      $selected.Add($test)
    }
  }
  foreach ($test in $selected) {
    Write-Output $test
  }
}

function Get-XlflowCodeModuleText {
  param($CodeModule)

  if ($null -eq $CodeModule -or $CodeModule.CountOfLines -le 0) {
    return ""
  }
  return $CodeModule.Lines(1, $CodeModule.CountOfLines)
}

function New-XlflowTestRunnerCode {
  param($Tests)

  $builder = New-Object System.Text.StringBuilder
  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Function RunTest(ByVal testIndex As Long) As Variant")
  [void]$builder.AppendLine("  On Error Resume Next")
  [void]$builder.AppendLine("  Err.Clear")
  [void]$builder.AppendLine("  Select Case testIndex")
  $index = 0
  foreach ($test in $Tests) {
    [void]$builder.AppendLine("    Case $index")
    [void]$builder.AppendLine("      " + $test.module + "." + $test.name)
    $index++
  }
  [void]$builder.AppendLine("  End Select")
  [void]$builder.AppendLine("  If Err.Number <> 0 Then")
  [void]$builder.AppendLine("    RunTest = Array(False, CLng(Err.Number), CStr(Err.Source), CStr(Err.Description))")
  [void]$builder.AppendLine("  Else")
  [void]$builder.AppendLine("    RunTest = Array(True, CLng(0), """", """")")
  [void]$builder.AppendLine("  End If")
  [void]$builder.AppendLine("  Err.Clear")
  [void]$builder.AppendLine("End Function")
  return $builder.ToString()
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
