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
  if ($Number -eq 1004 -and $Description -match '(?i)macro') {
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
  $invocation = $MacroName
  if ($literals.Count -gt 0) {
    $invocation += " " + ($literals -join ", ")
  }

  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Function RunMacro() As Variant")
  [void]$builder.AppendLine("  Dim startedAt As Double")
  [void]$builder.AppendLine("  startedAt = Timer")
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
