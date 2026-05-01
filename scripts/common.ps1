try {
  $script:XlflowConsoleUtf8 = New-Object System.Text.UTF8Encoding($false)
  [Console]::InputEncoding = $script:XlflowConsoleUtf8
  [Console]::OutputEncoding = $script:XlflowConsoleUtf8
  $OutputEncoding = $script:XlflowConsoleUtf8
} catch {
  Write-Verbose ("failed to configure UTF-8 console encoding: " + $_.Exception.Message)
}

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
    [int]$Number = 0,
    [int]$Line = 0,
    [string]$Phase = ""
  )
  $Result.status = "failed"
  $Result.error = [ordered]@{
    code = $Code
    message = $Message
    source = $Source
    number = $Number
    line = $Line
    phase = $Phase
  }
}

function ConvertTo-XlflowBool {
  param([string]$Value)
  return $Value -eq "true" -or $Value -eq "True" -or $Value -eq "1"
}

function Close-XlflowCom {
  param($Workbook, $Excel, [bool]$Save)
  if ($null -ne $Workbook) {
    try { $Workbook.Close($Save) | Out-Null } catch { Write-Verbose ("failed to close workbook: " + $_.Exception.Message) }
  }
  if ($null -ne $Excel) {
    try { $Excel.Quit() | Out-Null } catch { Write-Verbose ("failed to quit Excel: " + $_.Exception.Message) }
  }
  if ($null -ne $Workbook) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch { Write-Verbose ("failed to release workbook COM object: " + $_.Exception.Message) }
  }
  if ($null -ne $Excel) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch { Write-Verbose ("failed to release Excel COM object: " + $_.Exception.Message) }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Release-XlflowComReferences {
  param($Workbook, $Excel)
  $sessionWorkbook = Get-Variable -Name "XlflowSessionWorkbook" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  $sessionExcel = Get-Variable -Name "XlflowSessionExcel" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  if ($null -ne $Workbook) {
    if ($null -ne $sessionWorkbook -and [object]::ReferenceEquals($Workbook, $sessionWorkbook)) {
      Write-Verbose "leaving in-process xlflow session workbook reference open"
    } else {
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch { Write-Verbose ("failed to release workbook COM object: " + $_.Exception.Message) }
    }
  }
  if ($null -ne $Excel) {
    if ($null -ne $sessionExcel -and [object]::ReferenceEquals($Excel, $sessionExcel)) {
      Write-Verbose "leaving in-process xlflow session Excel reference open"
    } else {
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch { Write-Verbose ("failed to release Excel COM object: " + $_.Exception.Message) }
    }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Set-XlflowExcelAutomationDefaults {
  param($Excel, [bool]$DisplayAlerts = $false)

  if ($null -eq $Excel) {
    return
  }
  try { $Excel.DisplayAlerts = $DisplayAlerts } catch { Write-Verbose ("failed to set DisplayAlerts: " + $_.Exception.Message) }
  try { $Excel.EnableEvents = $false } catch { Write-Verbose ("failed to disable Excel events: " + $_.Exception.Message) }
}

function Disable-XlflowExcelAutomationMacros {
  param($Excel)

  if ($null -eq $Excel) {
    return
  }
  try { $Excel.AutomationSecurity = 3 } catch { Write-Verbose ("failed to force-disable automation macros: " + $_.Exception.Message) }
}

function Open-XlflowWorkbookWithXlflowDefaults {
  param(
    $Excel,
    [string]$WorkbookPath,
    [bool]$DisplayAlerts = $false,
    [bool]$DisableAutomationMacros = $true
  )

  Set-XlflowExcelAutomationDefaults -Excel $Excel -DisplayAlerts $DisplayAlerts
  if ($DisableAutomationMacros) {
    Disable-XlflowExcelAutomationMacros -Excel $Excel
  }
  return $Excel.Workbooks.Open($WorkbookPath)
}

function Get-XlflowActiveExcel {
  $sessionExcel = Get-Variable -Name "XlflowSessionExcel" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  if ($null -ne $sessionExcel) {
    return $sessionExcel
  }
  try {
    return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  } catch {
    throw "xlflow session is not running"
  }
}

function Get-XlflowOpenWorkbook {
  param($Excel, [string]$WorkbookPath)

  $target = [System.IO.Path]::GetFullPath($WorkbookPath)
  $sessionWorkbook = Get-Variable -Name "XlflowSessionWorkbook" -Scope Global -ValueOnly -ErrorAction SilentlyContinue
  if ($null -ne $sessionWorkbook) {
    try {
      if ([System.IO.Path]::GetFullPath([string]$sessionWorkbook.FullName) -ieq $target) {
        return $sessionWorkbook
      }
    } catch {
      Write-Verbose ("failed to inspect in-process session workbook: " + $_.Exception.Message)
    }
  }
  try {
    $bound = [System.Runtime.InteropServices.Marshal]::BindToMoniker($target)
    if ($null -ne $bound) {
      try {
        if ([System.IO.Path]::GetFullPath([string]$bound.FullName) -ieq $target) {
          return $bound
        }
      } catch {
        Write-Verbose ("failed to inspect moniker-bound workbook: " + $_.Exception.Message)
      }
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($bound) | Out-Null } catch { Write-Verbose ("failed to release moniker-bound workbook: " + $_.Exception.Message) }
    }
  } catch {
    Write-Verbose ("failed to bind open workbook by path: " + $_.Exception.Message)
  }
  foreach ($candidate in @($Excel.Workbooks)) {
    try {
      if ([System.IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $target) {
        return $candidate
      }
    } catch {
      Write-Verbose ("failed to inspect open workbook: " + $_.Exception.Message)
    }
  }
  throw "xlflow session workbook is not open: $WorkbookPath"
}

function Add-XlflowNativeMethods {
  if ("XlflowNativeMethods" -as [type]) {
    return
  }
  Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class XlflowNativeMethods {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")]
  public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

  [DllImport("oleacc.dll")]
  public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectId, ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

  public static IntPtr[] GetWindowsForProcess(uint targetProcessId) {
    List<IntPtr> windows = new List<IntPtr>();
    EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
      uint processId = 0;
      GetWindowThreadProcessId(hWnd, out processId);
      if (processId == targetProcessId) {
        windows.Add(hWnd);
      }
      return true;
    }, IntPtr.Zero);
    return windows.ToArray();
  }

  public static IntPtr[] GetChildWindows(IntPtr parentHwnd) {
    List<IntPtr> windows = new List<IntPtr>();
    EnumWindowsProc callback = delegate(IntPtr hWnd, IntPtr lParam) {
      windows.Add(hWnd);
      return true;
    };
    EnumChildWindows(parentHwnd, callback, IntPtr.Zero);
    return windows.ToArray();
  }

  [DllImport("user32.dll")]
  public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc enumProc, IntPtr lParam);
}
"@
}

function Get-XlflowExcelProcessId {
  param($Excel)

  if ($null -eq $Excel) {
    return 0
  }
  try {
    Add-XlflowNativeMethods
    $hwnd = [IntPtr]([int64]$Excel.Hwnd)
    $processId = [uint32]0
    [void][XlflowNativeMethods]::GetWindowThreadProcessId($hwnd, [ref]$processId)
    return [int]$processId
  } catch {
    Write-Verbose ("failed to resolve Excel process id: " + $_.Exception.Message)
    return 0
  }
}

function Get-XlflowExcelByProcessId {
  param([int]$ProcessId)

  if ($ProcessId -le 0) {
    return $null
  }
  try {
    Add-XlflowNativeMethods
    $iid = [Guid]"00020400-0000-0000-C000-000000000046"
    foreach ($hwnd in [XlflowNativeMethods]::GetWindowsForProcess([uint32]$ProcessId)) {
      $dispatch = $null
      $hr = [XlflowNativeMethods]::AccessibleObjectFromWindow($hwnd, 4294967280, [ref]$iid, [ref]$dispatch)
      if ($hr -ne 0 -or $null -eq $dispatch) {
        continue
      }
      $candidate = $dispatch
      try {
        $candidate = $dispatch.Application
      } catch {
        $candidate = $dispatch
      }
      try {
        if ($candidate.Workbooks.Count -gt 0) {
          return $candidate
        }
      } catch {
        Write-Verbose ("accessible object is not an Excel application: " + $_.Exception.Message)
      }
    }
  } catch {
    Write-Verbose ("failed to resolve Excel by process id: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowExcelByHwnd {
  param([int64]$Hwnd)

  if ($Hwnd -eq 0) {
    return $null
  }
  try {
    Add-XlflowNativeMethods
    $iid = [Guid]"00020400-0000-0000-C000-000000000046"
    foreach ($candidateHwnd in @([IntPtr]$Hwnd) + @([XlflowNativeMethods]::GetChildWindows([IntPtr]$Hwnd))) {
      $dispatch = $null
      $hr = [XlflowNativeMethods]::AccessibleObjectFromWindow($candidateHwnd, 4294967280, [ref]$iid, [ref]$dispatch)
      if ($hr -ne 0 -or $null -eq $dispatch) {
        continue
      }
      $candidate = $dispatch
      try {
        $candidate = $dispatch.Application
      } catch {
        $candidate = $dispatch
      }
      try {
        if ($candidate.Workbooks.Count -gt 0) {
          return $candidate
        }
      } catch {
        Write-Verbose ("accessible object is not an Excel application: " + $_.Exception.Message)
      }
    }
  } catch {
    Write-Verbose ("failed to resolve Excel by hwnd: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowExcelFromSessionMetadata {
  param([string]$MetadataPath)

  if ([string]::IsNullOrWhiteSpace($MetadataPath) -or -not (Test-Path -LiteralPath $MetadataPath)) {
    return $null
  }
  try {
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
    if ($null -ne $metadata -and $null -ne $metadata.hwnd -and $metadata.hwnd -ne 0) {
      $excel = Get-XlflowExcelByHwnd -Hwnd ([int64]$metadata.hwnd)
      if ($null -ne $excel) {
        return $excel
      }
    }
    if ($null -ne $metadata -and $metadata.pid -gt 0) {
      return Get-XlflowExcelByProcessId -ProcessId ([int]$metadata.pid)
    }
  } catch {
    Write-Verbose ("failed to read session metadata for Excel lookup: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowSessionExcel {
  param([string]$MetadataPath)

  $excel = Get-XlflowExcelFromSessionMetadata -MetadataPath $MetadataPath
  if ($null -ne $excel) {
    return $excel
  }
  return Get-XlflowActiveExcel
}

function Close-XlflowSessionWorkbook {
  param([string]$WorkbookPath, [string]$MetadataPath, [bool]$Save)

  $workbook = $null
  $excel = $null
  try {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
    try { $excel = $workbook.Application } catch { $excel = $null }
    $workbook.Close($Save) | Out-Null
    if ($null -ne $excel) {
      $excel.Quit() | Out-Null
    }
  } catch {
    Write-Verbose ("failed to close xlflow session workbook: " + $_.Exception.Message)
  } finally {
    if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and (Test-Path -LiteralPath $MetadataPath)) {
      Remove-Item -LiteralPath $MetadataPath -Force -ErrorAction SilentlyContinue
    }
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  }
}

function Get-XlflowFileHash {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    return ""
  }
  $stream = [System.IO.File]::OpenRead($Path)
  try {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
      $bytes = $sha.ComputeHash($stream)
      return ([System.BitConverter]::ToString($bytes) -replace "-", "").ToLowerInvariant()
    } finally {
      $sha.Dispose()
    }
  } finally {
    $stream.Dispose()
  }
}

function Get-XlflowSourceFingerprint {
  param(
    [string]$WorkbookPath,
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir
  )

  $files = New-Object System.Collections.Generic.List[object]
  foreach ($entry in @(
    @{ kind = "module"; dir = $ModulesDir },
    @{ kind = "class"; dir = $ClassesDir },
    @{ kind = "form"; dir = $FormsDir },
    @{ kind = "workbook"; dir = $WorkbookDir }
  )) {
    $dir = [string]$entry.dir
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) {
      continue
    }
    foreach ($file in Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Extension -in @(".bas", ".cls", ".frm", ".frx") } | Sort-Object FullName) {
      $base = [System.IO.Path]::GetFullPath($dir).TrimEnd("\", "/")
      $full = [System.IO.Path]::GetFullPath($file.FullName)
      $relative = $full
      if ($full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relative = $full.Substring($base.Length).TrimStart("\", "/")
      }
      $files.Add([ordered]@{
        kind = [string]$entry.kind
        path = $relative.Replace("\", "/")
        hash = Get-XlflowFileHash -Path $file.FullName
      })
    }
  }
  return [pscustomobject][ordered]@{
    workbook_path = [System.IO.Path]::GetFullPath($WorkbookPath)
    files = @($files.ToArray())
  }
}

function Test-XlflowFingerprintMatchesState {
  param($Fingerprint, [string]$StatePath)

  if ([string]::IsNullOrWhiteSpace($StatePath) -or -not (Test-Path -LiteralPath $StatePath)) {
    return $false
  }
  try {
    $existing = Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json
    $currentJson = $Fingerprint | ConvertTo-Json -Depth 10 -Compress
    $existingJson = $existing | ConvertTo-Json -Depth 10 -Compress
    return $currentJson -eq $existingJson
  } catch {
    return $false
  }
}

function Write-XlflowFingerprintState {
  param($Fingerprint, [string]$StatePath)

  if ([string]::IsNullOrWhiteSpace($StatePath)) {
    return
  }
  $parent = Split-Path -Parent $StatePath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }
  $Fingerprint | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StatePath -Encoding UTF8
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

function Get-XlflowUtf8Encoding {
  return (New-Object System.Text.UTF8Encoding -ArgumentList $false)
}

function Get-XlflowCp932Encoding {
  try {
    $providerType = [type]::GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages")
    if ($null -ne $providerType) {
      $provider = $providerType.GetProperty("Instance").GetValue($null, $null)
      [System.Text.Encoding]::RegisterProvider($provider)
    }
  } catch {
    Write-Verbose ("failed to register code page provider: " + $_.Exception.Message)
  }
  return [System.Text.Encoding]::GetEncoding(932)
}

function Get-XlflowUtf8Text {
  param([string]$Path)
  return [System.IO.File]::ReadAllText($Path, (Get-XlflowUtf8Encoding))
}

function Set-XlflowUtf8Text {
  param([string]$Path, [string]$Text)
  [System.IO.File]::WriteAllText($Path, $Text, (Get-XlflowUtf8Encoding))
}

function Get-XlflowCp932Text {
  param([string]$Path)
  return [System.IO.File]::ReadAllText($Path, (Get-XlflowCp932Encoding))
}

function Set-XlflowCp932Text {
  param([string]$Path, [string]$Text)
  [System.IO.File]::WriteAllText($Path, $Text, (Get-XlflowCp932Encoding))
}

function Convert-XlflowExportedSourceToUtf8 {
  param([string]$Path)
  $content = Get-XlflowCp932Text -Path $Path
  Set-XlflowUtf8Text -Path $Path -Text $content
}

function Copy-XlflowSourceForImport {
  param([string]$SourcePath, [string]$DestinationPath)

  $parent = Split-Path -Parent $DestinationPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }

  if ([System.IO.Path]::GetExtension($SourcePath) -ieq ".frx") {
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    return
  }

  $content = Get-XlflowUtf8Text -Path $SourcePath
  Set-XlflowCp932Text -Path $DestinationPath -Text $content
}

function Write-XlflowJson {
  param([hashtable]$Result)
  $Result | ConvertTo-Json -Depth 10
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

  $lines = (Get-XlflowUtf8Text -Path $Path) -split "`r?`n"
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
  Set-XlflowUtf8Text -Path $Path -Text $content
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

function ConvertFrom-XlflowRunArgumentsJson {
  param([string]$Json)

  if ([string]::IsNullOrWhiteSpace($Json)) {
    return @()
  }
  # Decode base64 JSON
  $decodedBytes = [System.Convert]::FromBase64String($Json)
  $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)

  $specs = ConvertFrom-Json -InputObject $decodedJson
  $values = New-Object System.Collections.Generic.List[object]
  foreach ($spec in $specs) {
    switch ([string]$spec.type) {
      "string" {
        $values.Add([string]$spec.value)
      }
      "int" {
        $parsed = 0
        if (-not [int]::TryParse([string]$spec.value, [ref]$parsed)) {
          throw "invalid int run argument: $($spec.value)"
        }
        $values.Add($parsed)
      }
      "bool" {
        if ($spec.value -ne "true" -and $spec.value -ne "false") {
          throw "invalid bool run argument: $($spec.value)"
        }
        $values.Add((ConvertTo-XlflowBool ([string]$spec.value)))
      }
      default {
        throw "unsupported run argument type: $($spec.type)"
      }
    }
  }
  return $values.ToArray()
}

function ConvertTo-XlflowVBALiteral {
  param([string]$Type, [string]$Value)

  switch ($Type) {
    "string" { return '"' + $Value.Replace('"', '""') + '"' }
    "int" { return "CLng(" + $Value + ")" }
    "bool" {
      if ($Value -eq "true") {
        return "CBool(True)"
      }
      return "CBool(False)"
    }
    default { throw "unsupported run argument type: $Type" }
  }
}

function Get-XlflowMacroModuleName {
  param([string]$MacroName)

  $parts = $MacroName.Split(".")
  if ($parts.Count -lt 2) {
    return $MacroName
  }
  return ($parts[0..($parts.Count - 2)] -join ".")
}

function Assert-XlflowSaveAsExtension {
  param([string]$WorkbookPath, [string]$SaveAsPath)

  if ([string]::IsNullOrWhiteSpace($SaveAsPath)) {
    return
  }
  $workbookExtension = [System.IO.Path]::GetExtension($WorkbookPath)
  $saveAsExtension = [System.IO.Path]::GetExtension($SaveAsPath)
  if ($workbookExtension -ne $saveAsExtension) {
    throw "save-as extension $saveAsExtension does not match workbook extension $workbookExtension"
  }
}

function Format-XlflowMacroFailureMessage {
  param(
    [string]$ModuleName,
    [int]$Line,
    [int]$Number,
    [string]$Description
  )

  $parts = New-Object System.Collections.Generic.List[string]
  if (-not [string]::IsNullOrWhiteSpace($ModuleName)) {
    $parts.Add($ModuleName)
  }
  if ($Line -gt 0) {
    $parts.Add("line " + $Line)
  }
  if ($Number -ne 0) {
    $parts.Add("Err " + $Number)
  }
  if ([string]::IsNullOrWhiteSpace($Description)) {
    return ($parts -join " ")
  }
  if ($parts.Count -eq 0) {
    return $Description
  }
  return (($parts -join " ") + ": " + $Description)
}

function New-XlflowTraceModuleCode {
  $builder = New-Object System.Text.StringBuilder
  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Private mTraceFile As String")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Sub XlflowSetTraceFile(ByVal path As String)")
  [void]$builder.AppendLine("  mTraceFile = path")
  [void]$builder.AppendLine("End Sub")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Sub XlflowLog(ByVal message As String)")
  [void]$builder.AppendLine("  If Len(mTraceFile) = 0 Then")
  [void]$builder.AppendLine('    Err.Raise vbObjectError + 900, "XlflowTrace.XlflowLog", "trace file is not configured. Run the macro with xlflow run --trace."')
  [void]$builder.AppendLine("  End If")
  [void]$builder.AppendLine("  Dim f As Integer")
  [void]$builder.AppendLine("  Dim opened As Boolean")
  [void]$builder.AppendLine("  On Error GoTo Handler")
  [void]$builder.AppendLine("  f = FreeFile")
  [void]$builder.AppendLine("  Open mTraceFile For Append As #f")
  [void]$builder.AppendLine("  opened = True")
  [void]$builder.AppendLine('  Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & message')
  [void]$builder.AppendLine("  Close #f")
  [void]$builder.AppendLine("  Exit Sub")
  [void]$builder.AppendLine("Handler:")
  [void]$builder.AppendLine("  Dim errNumber As Long")
  [void]$builder.AppendLine("  Dim errSource As String")
  [void]$builder.AppendLine("  Dim errDescription As String")
  [void]$builder.AppendLine("  errNumber = Err.Number")
  [void]$builder.AppendLine("  errSource = Err.Source")
  [void]$builder.AppendLine("  errDescription = Err.Description")
  [void]$builder.AppendLine("  On Error Resume Next")
  [void]$builder.AppendLine("  If opened Then Close #f")
  [void]$builder.AppendLine("  On Error GoTo 0")
  [void]$builder.AppendLine("  Err.Raise errNumber, errSource, errDescription")
  [void]$builder.AppendLine("End Sub")
  return $builder.ToString()
}

function Write-XlflowTraceModuleSource {
  param([string]$ModulesDir)

  if ([string]::IsNullOrWhiteSpace($ModulesDir)) {
    return $null
  }

  New-Item -ItemType Directory -Force -Path $ModulesDir | Out-Null
  $path = Join-Path $ModulesDir "XlflowTrace.bas"
  Set-XlflowUtf8Text -Path $path -Text (Get-XlflowTraceModuleSourceText)
  return $path
}

function Get-XlflowTraceModuleSourceText {
  return 'Attribute VB_Name = "XlflowTrace"' + [Environment]::NewLine + (New-XlflowTraceModuleCode)
}

function Test-XlflowTraceModuleSourceMatches {
  param([string]$ModulesDir)

  if ([string]::IsNullOrWhiteSpace($ModulesDir)) {
    return $false
  }
  $path = Join-Path $ModulesDir "XlflowTrace.bas"
  if (-not (Test-Path -LiteralPath $path)) {
    return $false
  }
  $existing = (Get-XlflowUtf8Text -Path $path).Trim()
  $expected = (Get-XlflowTraceModuleSourceText).Trim()
  return $existing -eq $expected
}

function Remove-XlflowTraceModule {
  param($VBProject)

  try {
    $existing = $VBProject.VBComponents.Item("XlflowTrace")
    $VBProject.VBComponents.Remove($existing)
    return $true
  } catch {
    return $false
  }
}

function Test-XlflowTraceModuleInjected {
  param($VBProject)

  try {
    $null = $VBProject.VBComponents.Item("XlflowTrace")
    return $true
  } catch {
    return $false
  }
}

function New-XlflowRunnerModuleCode {
  $builder = New-Object System.Text.StringBuilder
  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("' Persistent marker module for xlflow fast run workflows.")
  [void]$builder.AppendLine("Public Function XlflowRunnerVersion() As String")
  [void]$builder.AppendLine('  XlflowRunnerVersion = "1"')
  [void]$builder.AppendLine("End Function")
  return $builder.ToString()
}

function Test-XlflowRunnerModuleInstalled {
  param($VBProject)
  try {
    $null = $VBProject.VBComponents.Item("XlflowRunner")
    return $true
  } catch {
    return $false
  }
}

function Install-XlflowRunnerModule {
  param($VBProject)
  try {
    $existing = $VBProject.VBComponents.Item("XlflowRunner")
    $VBProject.VBComponents.Remove($existing)
  } catch {
    Write-Verbose ("XlflowRunner was not installed before install: " + $_.Exception.Message)
  }
  $component = $VBProject.VBComponents.Add(1)
  $component.Name = "XlflowRunner"
  $component.CodeModule.AddFromString((New-XlflowRunnerModuleCode))
}

function Remove-XlflowRunnerModule {
  param($VBProject)
  try {
    $existing = $VBProject.VBComponents.Item("XlflowRunner")
    $VBProject.VBComponents.Remove($existing)
    return $true
  } catch {
    return $false
  }
}

function ConvertTo-XlflowTraceEvent {
  param([string]$Line)

  $timestamp = ""
  $message = $Line
  $tab = $Line.IndexOf("`t")
  if ($tab -ge 0) {
    $timestamp = $Line.Substring(0, $tab)
    if ($tab + 1 -lt $Line.Length) {
      $message = $Line.Substring($tab + 1)
    } else {
      $message = ""
    }
  }
  return [ordered]@{
    timestamp = $timestamp
    message = $message
    raw = $Line
  }
}

function Read-XlflowTraceEvents {
  param([string]$Path)

  $events = New-Object System.Collections.Generic.List[object]
  if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
    return $events
  }
  $lines = Get-Content -LiteralPath $Path
  foreach ($line in $lines) {
    if ([string]::IsNullOrWhiteSpace($line)) {
      continue
    }
    $events.Add((ConvertTo-XlflowTraceEvent -Line $line))
  }
  foreach ($traceEvent in $events) {
    Write-Output $traceEvent
  }
}

function Find-XlflowMacroProcedures {
  param([string]$ModuleName, [string]$Code)

  $macros = New-Object System.Collections.Generic.List[object]
  if ([string]::IsNullOrEmpty($Code)) {
    return $macros
  }

  $lines = $Code -split "`r?`n"
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i].Trim()
    if ($line -match '^(?i)(Private|Friend)\s+(Sub|Function)\b') {
      continue
    }
    $match = [regex]::Match($line, '^(?:(Public)\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(([^)]*)\))?', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
      continue
    }
    $name = $match.Groups[3].Value
    if ([string]::IsNullOrWhiteSpace($name)) {
      continue
    }
    $argText = $match.Groups[4].Value.Trim()
    $macroArgs = @()
    if (-not [string]::IsNullOrWhiteSpace($argText)) {
      $macroArgs = @($argText -split "," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    $macros.Add([pscustomobject][ordered]@{
      module = $ModuleName
      name = $name
      qualified_name = ($ModuleName + "." + $name)
      kind = $match.Groups[2].Value.ToLowerInvariant()
      args = @($macroArgs)
      line = $i + 1
    })
  }

  foreach ($macro in $macros) {
    Write-Output $macro
  }
}

function Test-XlflowMacroTargetFailure {
  param(
    [int]$Number,
    [string]$Description
  )

  if ($Description -match '(?i)(cannot run the macro|sub or function not defined|macro may not be available|unable to run)') {
    return $true
  }
  if ($Description -match 'マクロ.*(実行できません|使用できない|利用できない)' -or $Description -match 'Sub または Function が定義されていません') {
    return $true
  }
  if ($Number -eq 1004 -and $Description -match '(?i)macro') {
    return $true
  }
  return $false
}

function Test-XlflowMacroDisabledFailure {
  param(
    [int]$Number,
    [string]$Description
  )

  if ([string]::IsNullOrWhiteSpace($Description)) {
    return $false
  }
  if ($Description -match '(?i)(security settings.*macro|macros? (?:have been|were|are) disabled|disable all macros|because of your security settings|security warning)') {
    return $true
  }
  if ($Description -match 'セキュリティ.*マクロ.*無効' -or $Description -match 'マクロ.*無効.*セキュリティ') {
    return $true
  }
  if ($Number -eq 1004 -and $Description -match 'セキュリティ') {
    return $true
  }
  return $false
}

function ConvertTo-XlflowUIButtonId {
  param([string]$Value)

  $value = ([string]$Value).Trim().ToLowerInvariant()
  $builder = New-Object System.Text.StringBuilder
  $lastDash = $false
  foreach ($char in $value.ToCharArray()) {
    $isValid = (($char -ge [char]'a' -and $char -le [char]'z') -or ($char -ge [char]'0' -and $char -le [char]'9'))
    if ($isValid) {
      [void]$builder.Append($char)
      $lastDash = $false
      continue
    }
    if (-not $lastDash -and $builder.Length -gt 0) {
      [void]$builder.Append("-")
      $lastDash = $true
    }
  }
  return $builder.ToString().Trim("-")
}

function ConvertTo-XlflowUIButtonName {
  param([string]$Id)
  return "xlflow.button." + (ConvertTo-XlflowUIButtonId -Value $Id)
}

function Get-XlflowWorksheet {
  param($Workbook, [string]$Sheet)

  foreach ($worksheet in @($Workbook.Worksheets)) {
    if ($worksheet.Name -eq $Sheet) {
      return $worksheet
    }
  }
  return $null
}

function Get-XlflowUIButton {
  param($Worksheet, [string]$Name)

  $buttons = $Worksheet.Buttons()
  for ($i = 1; $i -le $buttons.Count; $i++) {
    $button = $buttons.Item($i)
    if ($button.Name -eq $Name) {
      return $button
    }
  }
  return $null
}

function ConvertTo-XlflowUIButtonInfo {
  param($Button, [string]$Sheet, [string]$Id, [bool]$Updated = $false)

  $cell = ""
  try {
    $cell = $Button.TopLeftCell.Address($false, $false)
  } catch {
    Write-Verbose ("failed to read button top-left cell: " + $_.Exception.Message)
  }
  return [ordered]@{
    id = $Id
    name = $Button.Name
    sheet = $Sheet
    text = $Button.Caption
    macro = $Button.OnAction
    cell = $cell
    left = [double]$Button.Left
    top = [double]$Button.Top
    width = [double]$Button.Width
    height = [double]$Button.Height
    updated = $Updated
  }
}

function Test-XlflowMacroExists {
  param($Workbook, [string]$MacroName)

  $project = $Workbook.VBProject
  foreach ($component in @($project.VBComponents)) {
    if ($component.Name -like "Xlflow*") {
      continue
    }
    $code = Get-XlflowCodeModuleText -CodeModule $component.CodeModule
    $macros = Find-XlflowMacroProcedures -ModuleName $component.Name -Code $code
    foreach ($macro in @($macros)) {
      if ($macro.qualified_name -eq $MacroName -or $macro.name -eq $MacroName) {
        return $true
      }
    }
  }
  return $false
}

function New-XlflowRunHarnessModuleName {
  $suffix = [Guid]::NewGuid().ToString("N").Substring(0, 20)
  return "XlflowRun_" + $suffix
}

function New-XlflowRunHarnessCode {
  param(
    [string]$MacroName,
    [object[]]$Arguments,
    [bool]$TraceEnabled = $false,
    [string]$TraceFile = ""
  )

  $builder = New-Object System.Text.StringBuilder
  $moduleName = Get-XlflowMacroModuleName -MacroName $MacroName
  $literals = New-Object System.Collections.Generic.List[string]
  foreach ($argument in $Arguments) {
    $literals.Add((ConvertTo-XlflowVBALiteral -Type ([string]$argument.type) -Value ([string]$argument.value)))
  }
  $macroLiteral = ConvertTo-XlflowVBALiteral -Type "string" -Value $MacroName
  $invocation = "Application.Run targetMacro"
  if ($literals.Count -gt 0) {
    $invocation += ", " + ($literals -join ", ")
  }

  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Function RunMacro() As Variant")
  [void]$builder.AppendLine("  Dim startedAt As Double")
  [void]$builder.AppendLine("  Dim targetMacro As String")
  [void]$builder.AppendLine("  startedAt = Timer")
  [void]$builder.AppendLine("  targetMacro = ""'"" & ThisWorkbook.Name & ""'!"" & " + $macroLiteral)
  [void]$builder.AppendLine("  On Error GoTo Handler")
  if ($TraceEnabled) {
    [void]$builder.AppendLine("  XlflowTrace.XlflowSetTraceFile " + (ConvertTo-XlflowVBALiteral -Type "string" -Value $TraceFile))
  }
  [void]$builder.AppendLine("  " + $invocation)
  [void]$builder.AppendLine('  RunMacro = Array(True, "' + $moduleName + '", CLng(0), "", CLng(0), CLng((Timer - startedAt) * 1000))')
  [void]$builder.AppendLine("  Exit Function")
  [void]$builder.AppendLine("Handler:")
  [void]$builder.AppendLine('  RunMacro = Array(False, "' + $moduleName + '", CLng(Err.Number), CStr(Err.Description), CLng(Erl), CLng((Timer - startedAt) * 1000))')
  [void]$builder.AppendLine("  Err.Clear")
  [void]$builder.AppendLine("End Function")
  return $builder.ToString()
}
