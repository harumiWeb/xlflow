try {
  $script:XlflowConsoleUtf8 = New-Object System.Text.UTF8Encoding($false)
  [Console]::InputEncoding = $script:XlflowConsoleUtf8
  [Console]::OutputEncoding = $script:XlflowConsoleUtf8
  $OutputEncoding = $script:XlflowConsoleUtf8
} catch {
  Write-Verbose ("failed to configure UTF-8 console encoding: " + $_.Exception.Message)
}

function Get-XlflowPowerShellBridgeInfo {
  $hostName = ""
  $edition = ""
  $version = ""

  try {
    $process = Get-Process -Id $PID -ErrorAction Stop
    if ($null -ne $process -and -not [string]::IsNullOrWhiteSpace($process.ProcessName)) {
      $hostName = $process.ProcessName
      if (-not $hostName.EndsWith(".exe")) {
        $hostName += ".exe"
      }
    }
  } catch {
    Write-Verbose ("failed to resolve PowerShell bridge process name: " + $_.Exception.Message)
  }

  try {
    $edition = [string]$PSVersionTable.PSEdition
  } catch {
    Write-Verbose ("failed to resolve PowerShell edition: " + $_.Exception.Message)
  }

  try {
    if ($null -ne $PSVersionTable -and $null -ne $PSVersionTable.PSVersion) {
      $version = $PSVersionTable.PSVersion.ToString()
    }
  } catch {
    Write-Verbose ("failed to resolve PowerShell version: " + $_.Exception.Message)
  }

  if ([string]::IsNullOrWhiteSpace($hostName)) {
    if ($edition -eq "Core") {
      $hostName = "pwsh.exe"
    } else {
      $hostName = "powershell.exe"
    }
  }

  return [ordered]@{
    host = $hostName
    edition = $edition
    version = $version
  }
}

function New-XlflowResult {
  param([string]$Command)
  return [ordered]@{
    status = "ok"
    command = $Command
    error = $null
    logs = @()
    bridge = (Get-XlflowPowerShellBridgeInfo)
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

function Add-XlflowWarning {
  param(
    $Result,
    [string]$Code,
    [string]$Message
  )

  if (-not $Result.Contains("warnings") -or $null -eq $Result["warnings"]) {
    $Result["warnings"] = @()
  }
  $Result["warnings"] += [ordered]@{
    code = $Code
    message = $Message
  }
}

function Add-XlflowHint {
  param(
    $Result,
    [string]$Code,
    [string]$Message
  )

  if (-not $Result.Contains("hints") -or $null -eq $Result["hints"]) {
    $Result["hints"] = @()
  }
  $Result["hints"] += [ordered]@{
    code = $Code
    message = $Message
  }
}

function Add-XlflowStateWarning {
  param(
    $Result,
    [string]$Code,
    [string]$Message
  )

  Add-XlflowWarning -Result $Result -Code $Code -Message $Message
}

function Get-XlflowUserFormComponents {
  param($Workbook)

  $forms = New-Object System.Collections.Generic.List[object]
  if ($null -eq $Workbook) {
    return @()
  }
  try {
    foreach ($component in @($Workbook.VBProject.VBComponents)) {
      if ($null -ne $component -and [int]$component.Type -eq 3) {
        $forms.Add($component)
      }
    }
  } catch {
    throw
  }
  return @($forms.ToArray())
}

function Get-XlflowUserFormNames {
  param($Workbook)

  $names = New-Object System.Collections.Generic.List[string]
  foreach ($component in @(Get-XlflowUserFormComponents -Workbook $Workbook)) {
    try {
      $name = [string]$component.Name
      if (-not [string]::IsNullOrWhiteSpace($name)) {
        $names.Add($name)
      }
    } catch {
      Write-Verbose ("failed to read UserForm component name: " + $_.Exception.Message)
    }
  }
  return @($names.ToArray())
}

function Add-XlflowUserFormDiscoveryMessages {
  param(
    $Result,
    [string[]]$Names
  )

  $normalized = @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
  if ($normalized.Count -eq 0) {
    return
  }
  Add-XlflowWarning -Result $Result -Code "userform_state_partial" -Message ("UserForms detected: " + ($normalized -join ", ") + ". `.frm` text may not fully represent layout, binary `.frx` state, or VBIDE Designer-backed properties.")
  Add-XlflowHint -Result $Result -Code "userform_planned_commands" -Message "Related commands for deeper UserForm inspection include `xlflow form snapshot <name> --out <path>`, `xlflow inspect form <name> --runtime --json`, and `xlflow export-form-image <name>`."
}

function Add-XlflowUserFormSessionStaleWarning {
  param(
    $Result,
    [string[]]$Names
  )

  $normalized = @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
  if ($normalized.Count -eq 0) {
    return
  }
  Add-XlflowStateWarning -Result $Result -Code "userform_unsaved_session_state" -Message ("Workbook contains UserForms (" + ($normalized -join ", ") + ") and the current session changes are not saved. Disk `.frm`/`.frx` state may differ from the live workbook. Run `xlflow save --session` and `xlflow pull` before reviewing UserForm source differences.")
}

function Get-XlflowSourceUserFormNames {
  param([string]$FormsDir)

  if ([string]::IsNullOrWhiteSpace($FormsDir) -or -not (Test-Path -LiteralPath $FormsDir)) {
    return @()
  }

  $names = New-Object System.Collections.Generic.List[string]
  foreach ($file in Get-ChildItem -LiteralPath $FormsDir -Recurse -File -Filter *.frm | Sort-Object FullName) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    if (-not [string]::IsNullOrWhiteSpace($name)) {
      $names.Add($name)
    }
  }
  return @($names.ToArray() | Sort-Object -Unique)
}

function ConvertTo-XlflowBool {
  param([string]$Value)
  return $Value -eq "true" -or $Value -eq "True" -or $Value -eq "1"
}

function Get-XlflowRelativePath {
  param(
    [string]$BasePath,
    [string]$TargetPath
  )

  if ([string]::IsNullOrWhiteSpace($BasePath) -or [string]::IsNullOrWhiteSpace($TargetPath)) {
    return ""
  }

  $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
  $targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)
  $directorySeparator = [System.IO.Path]::DirectorySeparatorChar
  $altDirectorySeparator = [System.IO.Path]::AltDirectorySeparatorChar

  if (-not $baseFullPath.EndsWith([string]$directorySeparator) -and -not $baseFullPath.EndsWith([string]$altDirectorySeparator)) {
    $baseFullPath += $directorySeparator
  }

  $baseUri = New-Object System.Uri($baseFullPath)
  $targetUri = New-Object System.Uri($targetFullPath)
  $relativeUri = $baseUri.MakeRelativeUri($targetUri)
  $relativeUriText = [System.Uri]::UnescapeDataString($relativeUri.ToString())
  if ($relativeUri.IsAbsoluteUri -or $relativeUriText -match '^[A-Za-z]+:') {
    return $targetFullPath
  }
  $relativePath = $relativeUriText.Replace("/", "\")

  if ([string]::IsNullOrWhiteSpace($relativePath)) {
    return "."
  }

  return $relativePath
}

function Get-XlflowTargetDescription {
  param([string]$Kind)

  switch ($Kind) {
    "source" { return "VBA source files in the project directory" }
    "file" { return "Saved workbook file on disk" }
    "live_session" { return "Workbook currently open through xlflow session" }
    default { return "" }
  }
}

function New-XlflowTargetResult {
  param(
    [string]$Kind,
    [string]$Path = "",
    [string]$Description = "",
    [string]$Note = "",
    [string]$Sheet = "",
    [string]$Range = ""
  )

  $target = [ordered]@{
    kind = $Kind
  }
  if (-not [string]::IsNullOrWhiteSpace($Path)) {
    $target.path = $Path
  }
  if ([string]::IsNullOrWhiteSpace($Description)) {
    $Description = Get-XlflowTargetDescription -Kind $Kind
  }
  if (-not [string]::IsNullOrWhiteSpace($Description)) {
    $target.description = $Description
  }
  if (-not [string]::IsNullOrWhiteSpace($Note)) {
    $target.note = $Note
  }
  if (-not [string]::IsNullOrWhiteSpace($Sheet)) {
    $target.sheet = $Sheet
  }
  if (-not [string]::IsNullOrWhiteSpace($Range)) {
    $target.range = $Range
  }
  return $target
}

function New-XlflowSessionResult {
  param(
    [bool]$Active = $false,
    [string]$WorkbookPath = "",
    [AllowNull()]$Dirty = $null,
    [AllowNull()]$SaveRequired = $null,
    [string]$Mode = ""
  )

  $session = [ordered]@{
    active = $Active
  }
  if (-not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $session.workbook_path = $WorkbookPath
  }
  if ($PSBoundParameters.ContainsKey("Dirty") -and $null -ne $Dirty) {
    $session.dirty = [bool]$Dirty
  }
  if ($PSBoundParameters.ContainsKey("SaveRequired") -and $null -ne $SaveRequired) {
    $session.save_required = [bool]$SaveRequired
  }
  if (-not [string]::IsNullOrWhiteSpace($Mode)) {
    $session.mode = $Mode
  }
  return $session
}

function Close-XlflowCom {
  param($Workbook, $Excel, [bool]$Save)
  $excelProcessId = 0
  $workbookCloseFailed = $false
  $quitFailed = $false
  if ($null -ne $Excel) {
    $excelProcessId = Get-XlflowExcelProcessId -Excel $Excel
  }
  if ($null -ne $Workbook) {
    try { $Workbook.Close($Save) | Out-Null } catch { $workbookCloseFailed = $true; Write-Verbose ("failed to close workbook: " + $_.Exception.Message) }
  }
  if ($null -ne $Excel) {
    try { $Excel.Quit() | Out-Null } catch { $quitFailed = $true; Write-Verbose ("failed to quit Excel: " + $_.Exception.Message) }
  }
  Release-XlflowComObject -Object $Workbook -Name "workbook COM object"
  Release-XlflowComObject -Object $Excel -Name "Excel COM object"
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  if ($excelProcessId -gt 0) {
    $stillRunning = $false
    for ($i = 0; $i -lt 20; $i++) {
      if (-not (Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue)) {
        $stillRunning = $false
        break
      }
      $stillRunning = $true
      Start-Sleep -Milliseconds 100
    }
    if ($stillRunning) {
      if ($workbookCloseFailed -or $quitFailed) {
        Write-Verbose ("skipping force-stop for lingering Excel process " + $excelProcessId + " after graceful close failure")
      } else {
        try {
          Stop-Process -Id $excelProcessId -Force
        } catch {
          Write-Verbose ("failed to stop lingering Excel process " + $excelProcessId + ": " + $_.Exception.Message)
        }
      }
    }
  }
}

function Release-XlflowComObject {
  param(
    $Object,
    [string]$Name = "COM object"
  )

  if ($null -eq $Object) {
    return
  }

  try {
    if (-not [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
      return
    }
    while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) -gt 0) {
    }
  } catch {
    Write-Verbose ("failed to release " + $Name + ": " + $_.Exception.Message)
  }
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

  [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  public static extern int GetWindowTextW(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern int GetWindowTextLengthW(IntPtr hWnd);

  [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  public static extern int GetClassNameW(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

  [DllImport("user32.dll")]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern IntPtr SendMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool PostMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

  public static string GetWindowTextString(IntPtr hWnd) {
    int length = GetWindowTextLengthW(hWnd);
    System.Text.StringBuilder builder = new System.Text.StringBuilder(Math.Max(length + 1, 256));
    GetWindowTextW(hWnd, builder, builder.Capacity);
    return builder.ToString();
  }

  public static string GetClassNameString(IntPtr hWnd) {
    System.Text.StringBuilder builder = new System.Text.StringBuilder(256);
    GetClassNameW(hWnd, builder, builder.Capacity);
    return builder.ToString();
  }
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

function New-XlflowExcelDialogWatcherResult {
  return [pscustomobject][ordered]@{
    found = $false
    kind = ""
    hwnd = 0
    title = ""
    class_name = ""
    text = @()
    buttons = @()
    children = @()
    clicked_button = ""
    action = ""
  }
}

function Get-XlflowExcelDialogMessageLines {
  param($Dialog)

  if ($null -eq $Dialog) {
    return @()
  }

  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($line in @($Dialog.text)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
      $lines.Add([string]$line)
    }
  }
  if ($lines.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$Dialog.title)) {
    $lines.Add([string]$Dialog.title)
  }
  return @($lines.ToArray())
}

function Get-XlflowVBARuntimeDialogErrorNumber {
  param($Dialog)

  $message = ((Get-XlflowExcelDialogMessageLines -Dialog $Dialog) -join [Environment]::NewLine)
  if ([string]::IsNullOrWhiteSpace($message)) {
    return 0
  }

  $patterns = @(
    "(?i)(?:run-?time error|runtime error)\s*'?(?<number>-?\d+)'?",
    "実行時エラー\s*'?(?<number>-?\d+)'?"
  )
  foreach ($pattern in $patterns) {
    $match = [regex]::Match($message, $pattern)
    if (-not $match.Success) {
      continue
    }
    $parsed = 0
    if ([int]::TryParse($match.Groups["number"].Value, [ref]$parsed)) {
      return $parsed
    }
  }
  return 0
}

function Test-XlflowCompileDialogSignals {
  param(
    [string]$Title,
    [string]$StaticText,
    [string]$ButtonText
  )

  return (
    $StaticText -match "(?i)(compile|syntax error|expected)" -or
    $StaticText -match "コンパイル|構文エラー|必要です" -or
    $ButtonText -match "(?i)(Debug|Compile)" -or
    $ButtonText -match "(デバッグ|コンパイル)" -or
    $Title -match "(?i)(compile|syntax error)" -or
    $Title -match "コンパイル|構文エラー"
  )
}

function Test-XlflowAllowDialogFirstButtonFallback {
  param([string]$DialogKind)

  return $DialogKind -ne "compile"
}

function Start-XlflowExcelDialogWatcher {
  param(
    [int]$ProcessId,
    [string]$Kind = "compile",
    [int]$TimeoutMilliseconds = 10000,
    [int]$PollMilliseconds = 50
  )

  Add-XlflowNativeMethods
  if ($ProcessId -le 0) {
    return $null
  }

  $ps = [PowerShell]::Create()
  $null = $ps.AddScript({
    param([int]$TargetProcessId, [string]$DialogKind, [int]$TimeoutMs, [int]$PollMs)

    $bmClick = [uint32]0x00F5
    $wmClose = [uint32]0x0010
    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)

    while ([DateTime]::UtcNow -lt $deadline) {
      foreach ($hwnd in [XlflowNativeMethods]::GetWindowsForProcess([uint32]$TargetProcessId)) {
        if (-not [XlflowNativeMethods]::IsWindowVisible($hwnd)) {
          continue
        }

        $title = [XlflowNativeMethods]::GetWindowTextString($hwnd)
        $className = [XlflowNativeMethods]::GetClassNameString($hwnd)
        $childInfos = New-Object System.Collections.Generic.List[object]
        $staticTexts = New-Object System.Collections.Generic.List[string]
        $buttons = New-Object System.Collections.Generic.List[object]

        foreach ($child in [XlflowNativeMethods]::GetChildWindows($hwnd)) {
          $childText = [XlflowNativeMethods]::GetWindowTextString($child)
          $childClass = [XlflowNativeMethods]::GetClassNameString($child)
          $info = [pscustomobject][ordered]@{
            hwnd = [int64]$child
            class_name = $childClass
            text = $childText
          }
          $childInfos.Add($info)
          if ($childClass -eq "Static" -and -not [string]::IsNullOrWhiteSpace($childText)) {
            $staticTexts.Add($childText)
          }
          if ($childClass -eq "Button" -and -not [string]::IsNullOrWhiteSpace($childText)) {
            $buttons.Add($info)
          }
        }

        $buttonTexts = @($buttons | ForEach-Object { [string]$_.text })
        $joinedButtonText = ($buttonTexts -join [Environment]::NewLine)
        $joinedStaticText = (@($staticTexts.ToArray()) -join [Environment]::NewLine)
        $looksLikeVBAHostDialog = $className -eq "#32770" -and $buttons.Count -gt 0 -and ($staticTexts.Count -gt 0 -or $title -match "(?i)(visual basic|excel|vba)")
        $looksLikeRuntimeDialog = $looksLikeVBAHostDialog -and (
          $joinedStaticText -match "(?i)(run-?time error|runtime error)" -or
          $joinedStaticText -match "実行時エラー" -or
          $joinedButtonText -match "(?i)(Debug|End|Continue)" -or
          $joinedButtonText -match "(デバッグ|終了|継続)"
        )
        $looksLikeCompileDialog = $looksLikeVBAHostDialog -and (
          Test-XlflowCompileDialogSignals -Title $title -StaticText $joinedStaticText -ButtonText $joinedButtonText
        )

        $buttonToClick = $null
        $action = ""
        switch ($DialogKind) {
          "runtime" {
            if (-not $looksLikeRuntimeDialog) {
              continue
            }
            foreach ($button in $buttons) {
              if ($button.text -match "(?i)End" -or $button.text -match "終了") {
                $buttonToClick = $button
                $action = "runtime_end"
                break
              }
            }
            if ($null -eq $buttonToClick) {
              foreach ($button in $buttons) {
                if ($button.text -match "(?i)^(OK|Close)$" -or $button.text -match "はい|閉じる") {
                  $buttonToClick = $button
                  $action = "runtime_close"
                  break
                }
              }
            }
          }
          default {
            if (-not $looksLikeCompileDialog) {
              continue
            }
            foreach ($button in $buttons) {
              if ($button.text -match "(?i)^(OK|Close)$" -or $button.text -match "はい|閉じる") {
                $buttonToClick = $button
                $action = "compile_close"
                break
              }
            }
          }
        }

        if (-not $looksLikeCompileDialog -and -not $looksLikeRuntimeDialog) {
          continue
        }
        if ((Test-XlflowAllowDialogFirstButtonFallback -DialogKind $DialogKind) -and $null -eq $buttonToClick -and $buttons.Count -gt 0) {
          $buttonToClick = $buttons[0]
          if ([string]::IsNullOrWhiteSpace($action)) {
            $action = $DialogKind + "_first_button"
          }
        }

        if ($null -ne $buttonToClick) {
          [void][XlflowNativeMethods]::SendMessageW([IntPtr]([int64]$buttonToClick.hwnd), $bmClick, [IntPtr]::Zero, [IntPtr]::Zero)
        } else {
          [void][XlflowNativeMethods]::PostMessageW($hwnd, $wmClose, [IntPtr]::Zero, [IntPtr]::Zero)
          if ([string]::IsNullOrWhiteSpace($action)) {
            $action = $DialogKind + "_close"
          }
        }

        return [pscustomobject][ordered]@{
          found = $true
          kind = $DialogKind
          hwnd = [int64]$hwnd
          title = $title
          class_name = $className
          text = @($staticTexts.ToArray())
          buttons = @($buttons | ForEach-Object { $_.text })
          children = @($childInfos.ToArray())
          clicked_button = $(if ($null -ne $buttonToClick) { [string]$buttonToClick.text } else { "" })
          action = $action
        }
      }
      Start-Sleep -Milliseconds $PollMs
    }

    return [pscustomobject][ordered]@{
      found = $false
      kind = $DialogKind
      hwnd = 0
      title = ""
      class_name = ""
      text = @()
      buttons = @()
      children = @()
      clicked_button = ""
      action = ""
    }
  })
  $null = $ps.AddArgument($ProcessId)
  $null = $ps.AddArgument($Kind)
  $null = $ps.AddArgument($TimeoutMilliseconds)
  $null = $ps.AddArgument($PollMilliseconds)

  return [pscustomobject][ordered]@{
    powershell = $ps
    async = $ps.BeginInvoke()
  }
}

function Receive-XlflowExcelDialogWatcher {
  param(
    $Watcher,
    [int]$WaitMilliseconds = 250
  )

  if ($null -eq $Watcher) {
    return (New-XlflowExcelDialogWatcherResult)
  }

  try {
    if ($WaitMilliseconds -gt 0 -and -not $Watcher.async.AsyncWaitHandle.WaitOne($WaitMilliseconds)) {
      try { $Watcher.powershell.Stop() } catch { Write-Verbose ("failed to stop Excel dialog watcher: " + $_.Exception.Message) }
      return (New-XlflowExcelDialogWatcherResult)
    }
    $items = @($Watcher.powershell.EndInvoke($Watcher.async))
    if ($items.Count -gt 0) {
      return $items[0]
    }
  } catch {
    Write-Verbose ("failed to receive Excel dialog watcher result: " + $_.Exception.Message)
  } finally {
    try { $Watcher.powershell.Dispose() } catch { Write-Verbose ("failed to dispose Excel dialog watcher: " + $_.Exception.Message) }
  }

  return (New-XlflowExcelDialogWatcherResult)
}

function Start-XlflowVBEDialogWatcher {
  param(
    [int]$ProcessId,
    [int]$TimeoutMilliseconds = 10000,
    [int]$PollMilliseconds = 50
  )

  return Start-XlflowExcelDialogWatcher -ProcessId $ProcessId -Kind "compile" -TimeoutMilliseconds $TimeoutMilliseconds -PollMilliseconds $PollMilliseconds
}

function Receive-XlflowVBEDialogWatcher {
  param(
    $Watcher,
    [int]$WaitMilliseconds = 250
  )

  return Receive-XlflowExcelDialogWatcher -Watcher $Watcher -WaitMilliseconds $WaitMilliseconds
}

function Invoke-XlflowExcelCallWithDialogWatch {
  param(
    $Excel,
    $Workbook,
    [scriptblock]$Invocation,
    [string]$DialogKind = "runtime",
    [bool]$CaptureDialogs = $true,
    [int]$WaitMilliseconds = 250
  )

  $watcher = $null
  $caught = $null
  $value = $null
  $dialog = New-XlflowExcelDialogWatcherResult
  $selection = [ordered]@{
    location = [ordered]@{
      module = ""
      line = 0
      column = 0
      end_line = 0
      end_column = 0
      token = ""
    }
    nearby_code = @()
  }

  try {
    if ($CaptureDialogs) {
      $processId = Get-XlflowExcelProcessId -Excel $Excel
      $watcher = Start-XlflowExcelDialogWatcher -ProcessId $processId -Kind $DialogKind
    }
    try {
      $value = & $Invocation
    } catch {
      $caught = $_
    }
  } finally {
    $dialog = Receive-XlflowExcelDialogWatcher -Watcher $watcher -WaitMilliseconds $WaitMilliseconds
    if ($null -ne $Workbook -and [bool]$dialog.found) {
      try {
        $selection = Get-XlflowVBESelectionDiagnostic -VBE $Workbook.VBProject.VBE
      } catch {
        Write-Verbose ("failed to capture VBE selection after Excel dialog: " + $_.Exception.Message)
      }
    }
  }

  return [pscustomobject][ordered]@{
    value = $value
    exception = $caught
    dialog = $dialog
    selection = $selection
  }
}

function Get-XlflowVBECompileControl {
  param($VBE)

  $captions = @("Compile VBAProject", "Compile Project", "コンパイル", "VBAProject のコンパイル")
  try {
    $commandBars = $VBE.CommandBars
    foreach ($barName in @("Debug", "デバッグ")) {
      try {
        $bar = $commandBars.Item($barName)
        foreach ($caption in $captions) {
          try {
            return $bar.Controls.Item($caption)
          } catch {
            Write-Verbose ("compile control was not found by caption " + $caption + ": " + $_.Exception.Message)
          }
        }
        foreach ($control in @($bar.Controls)) {
          try {
            $caption = ([string]$control.Caption).Replace("&", "")
            if ($caption -match "(?i)compile" -or $caption -match "コンパイル") {
              return $control
            }
          } catch {
            Write-Verbose ("failed to inspect VBE Debug control: " + $_.Exception.Message)
          }
        }
      } catch {
        Write-Verbose ("VBE command bar was not found: " + $_.Exception.Message)
      }
    }
  } catch {
    Write-Verbose ("failed to inspect VBE command bars: " + $_.Exception.Message)
  }
  return $null
}

function Get-XlflowVBESelectionDiagnostic {
  param($VBE)

  $location = [ordered]@{
    module = ""
    line = 0
    column = 0
    end_line = 0
    end_column = 0
    token = ""
  }
  $nearby = @()

  try {
    $pane = $VBE.ActiveCodePane
    if ($null -eq $pane) {
      return [ordered]@{ location = $location; nearby_code = @() }
    }
    $module = $pane.CodeModule
    if ($null -ne $module) {
      $location.module = [string]$module.Name
    }
    $startLine = 0
    $startColumn = 0
    $endLine = 0
    $endColumn = 0
    $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
    $location.line = [int]$startLine
    $location.column = [int]$startColumn
    $location.end_line = [int]$endLine
    $location.end_column = [int]$endColumn

    if ($null -ne $module -and $startLine -gt 0 -and $startLine -le $module.CountOfLines) {
      $lineText = [string]$module.Lines($startLine, 1)
      if ($startColumn -gt 0 -and $startColumn -le $lineText.Length) {
        $tokenStart = $startColumn - 1
        $tokenLength = 1
        if ($endLine -eq $startLine -and $endColumn -gt $startColumn) {
          $tokenLength = [Math]::Min($endColumn - $startColumn, $lineText.Length - $tokenStart)
        } else {
          while ($tokenStart -gt 0 -and $lineText[$tokenStart - 1] -match "[A-Za-z0-9_]") {
            $tokenStart--
          }
          $tokenEnd = $startColumn - 1
          while ($tokenEnd -lt $lineText.Length -and $lineText[$tokenEnd] -match "[A-Za-z0-9_]") {
            $tokenEnd++
          }
          $tokenLength = [Math]::Max(1, $tokenEnd - $tokenStart)
        }
        $location.token = $lineText.Substring($tokenStart, $tokenLength).Trim()
      }

      $first = [Math]::Max(1, $startLine - 2)
      $last = [Math]::Min($module.CountOfLines, $startLine + 2)
      $items = New-Object System.Collections.Generic.List[string]
      for ($lineNo = $first; $lineNo -le $last; $lineNo++) {
        $prefix = "  "
        if ($lineNo -eq $startLine) {
          $prefix = "> "
        }
        $items.Add($prefix + $lineNo + " | " + [string]$module.Lines($lineNo, 1))
        if ($lineNo -eq $startLine -and $startColumn -gt 0) {
          $caretWidth = 1
          if ($endLine -eq $startLine -and $endColumn -gt $startColumn) {
            $caretWidth = [Math]::Max(1, $endColumn - $startColumn)
          } elseif (-not [string]::IsNullOrWhiteSpace($location.token)) {
            $caretWidth = $location.token.Length
          }
          $items.Add("    | " + (" " * [Math]::Max(0, $startColumn - 1)) + ("^" * $caretWidth))
        }
      }
      $nearby = @($items.ToArray())
    }
  } catch {
    Write-Verbose ("failed to read VBE selection diagnostic: " + $_.Exception.Message)
  }

  return [ordered]@{
    location = $location
    nearby_code = @($nearby)
  }
}

function Invoke-XlflowVBECompile {
  param($Excel, $Workbook)

  $result = [ordered]@{
    ok = $true
    dialog = [ordered]@{ found = $false; text = @(); buttons = @(); children = @() }
    selection = [ordered]@{ location = [ordered]@{}; nearby_code = @() }
    error = ""
  }

  $watcher = $null
  try {
    $processId = Get-XlflowExcelProcessId -Excel $Excel
    $watcher = Start-XlflowVBEDialogWatcher -ProcessId $processId
    $control = Get-XlflowVBECompileControl -VBE $Workbook.VBProject.VBE
    if ($null -eq $control) {
      throw "VBE Compile command was not found."
    }
    $null = $control.Execute()
  } catch {
    $result.error = $_.Exception.Message
  } finally {
    $dialog = Receive-XlflowVBEDialogWatcher -Watcher $watcher -WaitMilliseconds 3000
    if ($null -ne $dialog) {
      $result.dialog = $dialog
    }
  }

  if ($null -ne $result.dialog -and $result.dialog.found) {
    $result.ok = $false
    try {
      $result.selection = Get-XlflowVBESelectionDiagnostic -VBE $Workbook.VBProject.VBE
    } catch {
      Write-Verbose ("failed to collect compile selection diagnostic: " + $_.Exception.Message)
    }
  } elseif (-not [string]::IsNullOrWhiteSpace($result.error)) {
    throw $result.error
  }

  return $result
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

function Test-XlflowSessionMetadataMatchesWorkbook {
  param([string]$MetadataPath, [string]$WorkbookPath)

  if ([string]::IsNullOrWhiteSpace($MetadataPath) -or [string]::IsNullOrWhiteSpace($WorkbookPath) -or -not (Test-Path -LiteralPath $MetadataPath)) {
    return $false
  }
  try {
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
    if ($null -eq $metadata -or [string]::IsNullOrWhiteSpace([string]$metadata.workbook_path)) {
      return $false
    }
    return [System.IO.Path]::GetFullPath([string]$metadata.workbook_path) -ieq [System.IO.Path]::GetFullPath($WorkbookPath)
  } catch {
    Write-Verbose ("failed to compare session metadata workbook path: " + $_.Exception.Message)
    return $false
  }
}

function Get-XlflowWorkbookDirtyState {
  param($Workbook)

  if ($null -eq $Workbook) {
    return $null
  }
  try {
    return -not [bool]$Workbook.Saved
  } catch {
    Write-Verbose ("failed to inspect workbook dirty state: " + $_.Exception.Message)
    return $null
  }
}

function Get-XlflowWorkbookSaveState {
  param($Workbook, [bool]$SessionAttached = $false)

  $dirty = $false
  $needsSave = $false
  if ($SessionAttached) {
    $dirtyState = Get-XlflowWorkbookDirtyState -Workbook $Workbook
    if ($null -ne $dirtyState) {
      $dirty = [bool]$dirtyState
      $needsSave = [bool]$dirtyState
    } else {
      $dirty = $true
      $needsSave = $true
    }
  }

  return [ordered]@{
    dirty = $dirty
    needs_save = $needsSave
  }
}

function New-XlflowWorkbookResult {
  param(
    [string]$WorkbookPath,
    [bool]$SessionAttached = $false,
    [string]$SessionMode = "none",
    [AllowNull()]$Saved = $null,
    [string]$SaveAsPath = "",
    [AllowNull()]$Dirty = $null,
    [AllowNull()]$NeedsSave = $null,
    [AllowNull()]$DirtyBeforeStop = $null,
    [AllowNull()]$AutoSavedOnStop = $null
  )

  $workbook = [ordered]@{
    path = $WorkbookPath
    session = $SessionAttached
    session_mode = $SessionMode
    session_requested = ($SessionMode -eq "explicit")
    auto_session = ($SessionMode -eq "auto")
  }
  if ($PSBoundParameters.ContainsKey("Saved") -and $null -ne $Saved) {
    $workbook.saved = [bool]$Saved
  }
  if ($PSBoundParameters.ContainsKey("SaveAsPath")) {
    if ([string]::IsNullOrWhiteSpace($SaveAsPath)) {
      $workbook.save_as = $null
    } else {
      $workbook.save_as = $SaveAsPath
    }
  }
  if ($PSBoundParameters.ContainsKey("Dirty") -and $null -ne $Dirty) {
    $workbook.dirty = [bool]$Dirty
  }
  if ($PSBoundParameters.ContainsKey("NeedsSave") -and $null -ne $NeedsSave) {
    $workbook.needs_save = [bool]$NeedsSave
  }
  if ($PSBoundParameters.ContainsKey("DirtyBeforeStop") -and $null -ne $DirtyBeforeStop) {
    $workbook.dirty_before_stop = [bool]$DirtyBeforeStop
  }
  if ($PSBoundParameters.ContainsKey("AutoSavedOnStop") -and $null -ne $AutoSavedOnStop) {
    $workbook.auto_saved_on_stop = [bool]$AutoSavedOnStop
  }
  return $workbook
}

function Get-XlflowSessionUsageLog {
  param([string]$SessionMode)

  switch ($SessionMode) {
    "explicit" { return "using xlflow session workbook (--session)" }
    "auto" { return "auto-reused matching xlflow session workbook" }
    "managed" { return "using managed xlflow session workbook" }
    default { return $null }
  }
}

function Open-XlflowWorkbookForCommand {
  param(
    [string]$WorkbookPath,
    [string]$Visible = "false",
    [string]$DisplayAlerts = "false",
    [string]$DisableAutomationMacros = "true",
    [string]$UseSession = "false",
    [string]$MetadataPath = "",
    [bool]$AllowIsolatedOpen = $true
  )

  if (ConvertTo-XlflowBool $UseSession) {
    $excel = Get-XlflowSessionExcel -MetadataPath $MetadataPath
    $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
    return [pscustomobject][ordered]@{
      excel = $excel
      workbook = $workbook
      session_attached = $true
      session_mode = "explicit"
    }
  }

  if (Test-XlflowSessionMetadataMatchesWorkbook -MetadataPath $MetadataPath -WorkbookPath $WorkbookPath) {
    $excel = Get-XlflowExcelFromSessionMetadata -MetadataPath $MetadataPath
    if ($null -ne $excel) {
      try {
        $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $WorkbookPath
        return [pscustomobject][ordered]@{
          excel = $excel
          workbook = $workbook
          session_attached = $true
          session_mode = "auto"
        }
      } catch {
        Write-Verbose ("failed to attach to matching xlflow session workbook: " + $_.Exception.Message)
      }
    }
  }

  if (-not $AllowIsolatedOpen) {
    throw "no matching xlflow session workbook is running; run xlflow session start or use the configured workbook session"
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $workbook = Open-XlflowWorkbookWithXlflowDefaults -Excel $excel -WorkbookPath $WorkbookPath -DisplayAlerts (ConvertTo-XlflowBool $DisplayAlerts) -DisableAutomationMacros (ConvertTo-XlflowBool $DisableAutomationMacros)
  return [pscustomobject][ordered]@{
    excel = $excel
    workbook = $workbook
    session_attached = $false
    session_mode = "none"
  }
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

function Get-XlflowFolderConfig {
  param(
    [string]$Folders = "true",
    [string]$FolderAnnotation = "update",
    [string]$DefaultComponentFolders = "true"
  )

  return [pscustomobject][ordered]@{
    folders = (ConvertTo-XlflowBool $Folders)
    folder_annotation = [string]$FolderAnnotation
    default_component_folders = (ConvertTo-XlflowBool $DefaultComponentFolders)
  }
}

function Get-XlflowComponentExtension {
  param([int]$ComponentType)

  switch ($ComponentType) {
    1 { return ".bas" }
    2 { return ".cls" }
    3 { return ".frm" }
    100 { return ".bas" }
    default { return "" }
  }
}

function Get-XlflowComponentRootDir {
  param(
    [int]$ComponentType,
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir
  )

  switch ($ComponentType) {
    1 { return $ModulesDir }
    2 { return $ClassesDir }
    3 { return $FormsDir }
    100 { return $WorkbookDir }
    default { return "" }
  }
}

function Get-XlflowContentNewline {
  param([string]$Text)

  if ($Text -match "`r`n") {
    return "`r`n"
  }
  return "`n"
}

function ConvertTo-XlflowFolderPathSegment {
  param([string]$Segment)

  $clean = [string]$Segment
  $clean = $clean.Trim()
  if ([string]::IsNullOrWhiteSpace($clean)) {
    return ""
  }
  $clean = [regex]::Replace($clean, '[<>:"/\\|?*\x00-\x1f]', "_")
  $clean = $clean.Trim(" ", ".")
  if ($clean -eq "." -or $clean -eq "..") {
    return ""
  }
  return $clean
}

function ConvertFrom-XlflowFolderAnnotation {
  param([string]$Annotation)

  if ([string]::IsNullOrWhiteSpace($Annotation)) {
    return @()
  }

  $segments = New-Object System.Collections.Generic.List[string]
  foreach ($part in ([string]$Annotation).Split(".")) {
    $segment = ConvertTo-XlflowFolderPathSegment -Segment $part
    if ([string]::IsNullOrWhiteSpace($segment)) {
      return $null
    }
    $segments.Add($segment)
  }
  return @($segments.ToArray())
}

function ConvertTo-XlflowFolderAnnotation {
  param([string[]]$Segments)

  if ($null -eq $Segments -or $Segments.Count -eq 0) {
    return ""
  }

  $clean = New-Object System.Collections.Generic.List[string]
  foreach ($segment in $Segments) {
    $part = ConvertTo-XlflowFolderPathSegment -Segment $segment
    if ([string]::IsNullOrWhiteSpace($part)) {
      continue
    }
    $clean.Add($part)
  }
  if ($clean.Count -eq 0) {
    return ""
  }
  return ($clean -join ".")
}

function Test-XlflowFolderAnnotationHeaderLine {
  param([string]$TrimmedLine)

  if ([string]::IsNullOrWhiteSpace($TrimmedLine)) {
    return $true
  }
  if ($TrimmedLine.StartsWith("'")) {
    return $true
  }
  if ($TrimmedLine -match '^Attribute\s+VB_') {
    return $true
  }
  if ($TrimmedLine -match '^(Option\s+Explicit|VERSION\s+1\.0\s+CLASS|BEGIN\b|END\b|MultiUse\b)') {
    return $true
  }
  return $false
}

function Find-XlflowFolderAnnotation {
  param(
    [string]$Text,
    [int]$MaxLines = 50
  )

  $lines = $Text -split "`r?`n"
  $max = [Math]::Min($lines.Count, $MaxLines)
  for ($i = 0; $i -lt $max; $i++) {
    $line = [string]$lines[$i]
    $trimmed = $line.Trim()
    if ($trimmed -match "^'\s*@Folder\(""([^""]+)""\)\s*$") {
      $segments = ConvertFrom-XlflowFolderAnnotation -Annotation $matches[1]
      return [pscustomobject][ordered]@{
        found = $true
        valid = ($null -ne $segments)
        malformed = ($null -eq $segments)
        annotation = $(if ($null -ne $segments) { $matches[1] } else { "" })
        line_index = $i
      }
    }
    if ($trimmed -match "@Folder") {
      return [pscustomobject][ordered]@{
        found = $true
        valid = $false
        malformed = $true
        annotation = ""
        line_index = $i
      }
    }
    if (-not (Test-XlflowFolderAnnotationHeaderLine -TrimmedLine $trimmed)) {
      break
    }
  }

  return [pscustomobject][ordered]@{
    found = $false
    valid = $false
    malformed = $false
    annotation = ""
    line_index = -1
  }
}

function Get-XlflowFolderAnnotationInsertIndex {
  param([string[]]$Lines)

  $index = 0
  if ($Lines.Count -gt 0 -and $Lines[0].Trim() -eq "VERSION 1.0 CLASS") {
    $index = 1
    while ($index -lt $Lines.Count) {
      $trimmed = $Lines[$index].Trim()
      $index++
      if ($trimmed -eq "END") {
        break
      }
    }
  }
  while ($index -lt $Lines.Count -and $Lines[$index].Trim() -match '^Attribute\s+VB_') {
    $index++
  }
  return $index
}

function Update-XlflowFolderAnnotationText {
  param(
    [string]$Text,
    [string]$FolderAnnotationMode = "update",
    [string]$DesiredAnnotation = ""
  )

  if ($FolderAnnotationMode -eq "ignore" -or $FolderAnnotationMode -eq "preserve") {
    return $Text
  }

  $newline = Get-XlflowContentNewline -Text $Text
  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($line in ($Text -split "`r?`n")) {
    $lines.Add([string]$line)
  }
  $annotationInfo = Find-XlflowFolderAnnotation -Text $Text
  $annotationLine = ""
  if (-not [string]::IsNullOrWhiteSpace($DesiredAnnotation)) {
    $annotationLine = "'@Folder(""$DesiredAnnotation"")"
  }

  if ($annotationInfo.found) {
    if ([string]::IsNullOrWhiteSpace($annotationLine)) {
      $lines.RemoveAt([int]$annotationInfo.line_index)
    } else {
      $lines[[int]$annotationInfo.line_index] = $annotationLine
    }
  } elseif (-not [string]::IsNullOrWhiteSpace($annotationLine)) {
    $insertIndex = Get-XlflowFolderAnnotationInsertIndex -Lines @($lines.ToArray())
    $lines.Insert([int]$insertIndex, $annotationLine)
  }

  return (@($lines.ToArray()) -join $newline)
}

function Get-XlflowRelativePathSegments {
  param(
    [string]$RootDir,
    [string]$Path
  )

  if ([string]::IsNullOrWhiteSpace($RootDir) -or [string]::IsNullOrWhiteSpace($Path)) {
    return @()
  }

  $relativeDir = Split-Path -Parent (Get-XlflowRelativePath -BasePath $RootDir -TargetPath $Path)
  if ([string]::IsNullOrWhiteSpace($relativeDir) -or $relativeDir -eq ".") {
    return @()
  }

  $segments = New-Object System.Collections.Generic.List[string]
  foreach ($part in ($relativeDir -split '[\\/]')) {
    if ($part -eq "..") {
      throw "path '$Path' resolves outside root '$RootDir' (relative: '$relativeDir')"
    }
    $segment = ConvertTo-XlflowFolderPathSegment -Segment $part
    if ([string]::IsNullOrWhiteSpace($segment)) {
      continue
    }
    $segments.Add($segment)
  }
  return @($segments.ToArray())
}

function Get-XlflowFolderAnnotationForPath {
  param(
    [string]$RootDir,
    [string]$Path
  )

  return ConvertTo-XlflowFolderAnnotation -Segments (Get-XlflowRelativePathSegments -RootDir $RootDir -Path $Path)
}

function Get-XlflowFolderAnnotationForComponent {
  param(
    $Component,
    [string]$FolderAnnotationMode = "update"
  )

  if ($FolderAnnotationMode -eq "ignore") {
    return ""
  }

  try {
    $text = Get-XlflowCodeModuleText -CodeModule $Component.CodeModule
    $annotation = Find-XlflowFolderAnnotation -Text $text
    if ($annotation.valid) {
      return [string]$annotation.annotation
    }
  } catch {
    Write-Verbose ("failed to inspect folder annotation from code module: " + $_.Exception.Message)
  }
  return ""
}

function Get-XlflowSourceComponentFiles {
  param(
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
    @{ kind = "document"; dir = $WorkbookDir }
  )) {
    $dir = [string]$entry.dir
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) {
      continue
    }
    foreach ($file in Get-ChildItem -LiteralPath $dir -Recurse -File | Sort-Object FullName) {
      if ($file.Extension -notin @(".bas", ".cls", ".frm", ".frx")) {
        continue
      }
      $files.Add([pscustomobject][ordered]@{
        kind = [string]$entry.kind
        root_dir = $dir
        full_name = $file.FullName
        relative_path = (Get-XlflowRelativePath -BasePath $dir -TargetPath $file.FullName).Replace("\", "/")
        extension = [string]$file.Extension.ToLowerInvariant()
        module_name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
      })
    }
  }
  return @($files.ToArray())
}

function Clear-XlflowSourceComponentFiles {
  param(
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir
  )

  foreach ($file in @(Get-XlflowSourceComponentFiles -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir)) {
    Remove-Item -LiteralPath $file.full_name -Force -ErrorAction SilentlyContinue
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
  foreach ($file in @(Get-XlflowSourceComponentFiles -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir)) {
    $files.Add([ordered]@{
      kind = [string]$file.kind
      path = [string]$file.relative_path
      hash = Get-XlflowFileHash -Path $file.full_name
    })
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
  param(
    $Component,
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$WorkbookDir,
    [string]$Folders = "true",
    [string]$FolderAnnotation = "update",
    [string]$DefaultComponentFolders = "true"
  )

  $rootDir = Get-XlflowComponentRootDir -ComponentType $Component.Type -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir
  $extension = Get-XlflowComponentExtension -ComponentType $Component.Type
  if ([string]::IsNullOrWhiteSpace($rootDir) -or [string]::IsNullOrWhiteSpace($extension)) {
    return $null
  }

  $segments = @()
  $folderConfig = Get-XlflowFolderConfig -Folders $Folders -FolderAnnotation $FolderAnnotation -DefaultComponentFolders $DefaultComponentFolders
  if ($folderConfig.folders -and $folderConfig.folder_annotation -ne "ignore") {
    $annotation = Get-XlflowFolderAnnotationForComponent -Component $Component -FolderAnnotationMode $folderConfig.folder_annotation
    if (-not [string]::IsNullOrWhiteSpace($annotation)) {
      $parsed = ConvertFrom-XlflowFolderAnnotation -Annotation $annotation
      if ($null -ne $parsed) {
        $segments = $parsed
      }
    }
  }

  $path = $rootDir
  foreach ($segment in $segments) {
    $path = Join-Path $path $segment
  }
  return Join-Path $path ($Component.Name + $extension)
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
  param(
    [string]$SourcePath,
    [string]$DestinationPath,
    [string]$RootDir = "",
    [string]$FolderAnnotationMode = "update"
  )

  $parent = Split-Path -Parent $DestinationPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
  }

  if ([System.IO.Path]::GetExtension($SourcePath) -ieq ".frx") {
    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    return
  }

  $content = Get-XlflowUtf8Text -Path $SourcePath
  if (-not [string]::IsNullOrWhiteSpace($RootDir)) {
    $desiredAnnotation = Get-XlflowFolderAnnotationForPath -RootDir $RootDir -Path $SourcePath
    $content = Update-XlflowFolderAnnotationText -Text $content -FolderAnnotationMode $FolderAnnotationMode -DesiredAnnotation $desiredAnnotation
  }
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
  param(
    [string]$Path,
    [string]$RootDir = "",
    [string]$FolderAnnotationMode = "update"
  )

  $content = Get-XlflowDocumentModuleContent -Path $Path
  if (-not [string]::IsNullOrWhiteSpace($RootDir)) {
    $desiredAnnotation = Get-XlflowFolderAnnotationForPath -RootDir $RootDir -Path $Path
    $content = Update-XlflowFolderAnnotationText -Text $content -FolderAnnotationMode $FolderAnnotationMode -DesiredAnnotation $desiredAnnotation
  }
  Set-XlflowUtf8Text -Path $Path -Text $content
}

function Sync-XlflowDocumentModule {
  param(
    $Component,
    [string]$Path,
    [string]$RootDir = "",
    [string]$FolderAnnotationMode = "update"
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    return $false
  }

  $code = Get-XlflowDocumentModuleContent -Path $Path
  if (-not [string]::IsNullOrWhiteSpace($RootDir)) {
    $desiredAnnotation = Get-XlflowFolderAnnotationForPath -RootDir $RootDir -Path $Path
    $code = Update-XlflowFolderAnnotationText -Text $code -FolderAnnotationMode $FolderAnnotationMode -DesiredAnnotation $desiredAnnotation
  }
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

function Find-XlflowDuplicateModuleNames {
  param($Files)

  $seen = @{}
  $originalNames = @{}
  foreach ($file in @($Files)) {
    if ($file.extension -eq ".frx") {
      continue
    }
    $key = ([string]$file.module_name).ToLowerInvariant()
    if (-not $seen.ContainsKey($key)) {
      $seen[$key] = New-Object System.Collections.Generic.List[string]
      $originalNames[$key] = [string]$file.module_name
    }
    $seen[$key].Add([string]$file.relative_path)
  }

  $duplicates = New-Object System.Collections.Generic.List[object]
  foreach ($key in $seen.Keys | Sort-Object) {
    if ($seen[$key].Count -lt 2) {
      continue
    }
    $duplicates.Add([pscustomobject][ordered]@{
      module_name = [string]$originalNames[$key]
      paths = @($seen[$key].ToArray())
    })
  }
  return @($duplicates.ToArray())
}

function Find-XlflowDocumentModulePath {
  param(
    [string]$WorkbookDir,
    [string]$ComponentName
  )

  if ([string]::IsNullOrWhiteSpace($WorkbookDir) -or -not (Test-Path -LiteralPath $WorkbookDir)) {
    return ""
  }

  foreach ($file in @(Get-XlflowSourceComponentFiles -ModulesDir "" -ClassesDir "" -FormsDir "" -WorkbookDir $WorkbookDir)) {
    if ($file.extension -eq ".bas" -and $file.module_name -ieq $ComponentName) {
      return [string]$file.full_name
    }
  }
  return ""
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

function Remove-XlflowVBComponentByName {
  param(
    $VBProject,
    [string]$Name
  )

  try {
    $existing = $VBProject.VBComponents.Item($Name)
    $VBProject.VBComponents.Remove($existing)
    return $true
  } catch {
    return $false
  }
}

function Get-XlflowVBComponentByName {
  param(
    $VBProject,
    [string]$Name
  )

  try {
    return $VBProject.VBComponents.Item($Name)
  } catch {
    return $null
  }
}

function Install-XlflowVBComponentFromCode {
  param(
    $VBProject,
    [string]$Name,
    [string]$Code
  )

  $existing = Get-XlflowVBComponentByName -VBProject $VBProject -Name $Name
  if ($null -ne $existing) {
    throw ("VBA component '" + $Name + "' already exists.")
  }
  $component = $VBProject.VBComponents.Add(1)
  $component.Name = $Name
  $component.CodeModule.AddFromString($Code)
  return $component
}

function New-XlflowInspectFormModuleCode {
  return @'
Option Explicit

Private Const xlflowBasisDesigner As String = "designer"
Private Const xlflowBasisRuntime As String = "runtime"
Private Const xlflowCoordinateSystem As String = "parent-relative"

Public Function XlflowInspectFormJson(ByVal formName As String, ByVal basis As String, Optional ByVal initializer As String = "") As String
  Dim normalizedBasis As String
  normalizedBasis = LCase$(Trim$(basis))

  Select Case normalizedBasis
    Case xlflowBasisDesigner
      XlflowInspectFormJson = InspectDesignerFormJson(formName)
    Case xlflowBasisRuntime
      XlflowInspectFormJson = InspectRuntimeFormJson(formName, initializer)
    Case Else
      Err.Raise vbObjectError + 7300, "XlflowInspectFormJson.args", "Unsupported inspect basis: " & basis
  End Select
End Function

Private Function InspectDesignerFormJson(ByVal formName As String) As String
  On Error GoTo ErrHandler

  Dim component As Object
  Dim designer As Object

  Set component = ThisWorkbook.VBProject.VBComponents.Item(formName)
  Set designer = component.Designer
  InspectDesignerFormJson = SerializeFormSnapshot(formName, xlflowBasisDesigner, designer)
  Exit Function

ErrHandler:
  Err.Raise Err.Number, "XlflowInspectFormJson.designer", Err.Description
End Function

Private Function InspectRuntimeFormJson(ByVal formName As String, ByVal initializer As String) As String
  Dim formInstance As Object
  Dim loaded As Boolean
  Dim initializerRan As Boolean
  Dim errorNumber As Long
  Dim errorDescription As String
  Dim errorSource As String

  On Error GoTo ErrHandler

  Set formInstance = UserForms.Add(formName)
  loaded = True

  If Len(Trim$(initializer)) > 0 Then
    CallByName formInstance, Trim$(initializer), VbMethod, ThisWorkbook
    initializerRan = True
  End If

  InspectRuntimeFormJson = SerializeFormSnapshot(formName, xlflowBasisRuntime, formInstance)

Cleanup:
  On Error Resume Next
  If loaded Then
    Unload formInstance
  End If
  Set formInstance = Nothing
  On Error GoTo 0

  If errorNumber <> 0 Then
    Err.Raise errorNumber, errorSource, errorDescription
  End If
  Exit Function

ErrHandler:
  errorNumber = Err.Number
  errorDescription = Err.Description
  If Not loaded Then
    errorSource = "XlflowInspectFormJson.runtime_load"
  ElseIf Len(Trim$(initializer)) > 0 And Not initializerRan Then
    errorSource = "XlflowInspectFormJson.initializer"
  Else
    errorSource = "XlflowInspectFormJson.enumerate"
  End If
  Resume Cleanup
End Function

Private Function SerializeFormSnapshot(ByVal formName As String, ByVal basis As String, ByVal formObject As Object) As String
  Dim json As String
  Dim hasFields As Boolean
  Dim controls As Object

  json = "{"

  JsonAddString json, "name", formName, hasFields
  JsonAddString json, "basis", basis, hasFields
  JsonAddStringFromMember json, formObject, "Caption", "caption", hasFields
  JsonAddNumberFromMember json, formObject, "Width", "width", hasFields
  JsonAddNumberFromMember json, formObject, "Height", "height", hasFields
  JsonAddString json, "coordinate_system", xlflowCoordinateSystem, hasFields

  Set controls = GetObjectControls(formObject)
  JsonAddRaw json, "controls", SerializeControls(controls), hasFields

  json = json & "}"
  SerializeFormSnapshot = json
End Function

Private Function SerializeControls(ByVal controls As Object) As String
  Dim json As String
  Dim first As Boolean
  Dim control As Object

  json = "["
  first = True

  If Not controls Is Nothing Then
    For Each control In controls
      If Not first Then
        json = json & ","
      End If
      json = json & SerializeControl(control)
      first = False
    Next control
  End If

  json = json & "]"
  SerializeControls = json
End Function

Private Function SerializeControl(ByVal control As Object) As String
  Dim json As String
  Dim hasFields As Boolean
  Dim children As Object
  Dim listCount As Long
  Dim selectedIndex As Long
  Dim typeNameValue As String

  json = "{"

  JsonAddStringFromMember json, control, "Name", "name", hasFields
  typeNameValue = TypeName(control)
  JsonAddString json, "type", typeNameValue, hasFields
  JsonAddStringFromMember json, control, "ProgId", "prog_id", hasFields
  JsonAddStringFromMember json, control, "Caption", "caption", hasFields
  JsonAddStringFromMember json, control, "Text", "text", hasFields
  JsonAddStringFromMember json, control, "Value", "value", hasFields
  JsonAddNumberFromMember json, control, "Left", "left", hasFields
  JsonAddNumberFromMember json, control, "Top", "top", hasFields
  JsonAddNumberFromMember json, control, "Width", "width", hasFields
  JsonAddNumberFromMember json, control, "Height", "height", hasFields
  JsonAddLongFromMember json, control, "TabIndex", "tab_index", hasFields
  JsonAddBoolFromMember json, control, "Enabled", "enabled", hasFields
  JsonAddBoolFromMember json, control, "Visible", "visible", hasFields

  If TryGetLongMember(control, "ListIndex", selectedIndex) Then
    JsonAddLong json, "selected_index", selectedIndex, hasFields
  End If
  If TryGetLongMember(control, "ListCount", listCount) Then
    JsonAddRaw json, "list", SerializeControlList(control, listCount), hasFields
  End If

  If ControlCanContainChildren(typeNameValue) Then
    Set children = GetObjectControls(control)
  End If
  If Not children Is Nothing Then
    JsonAddRaw json, "controls", SerializeControls(children), hasFields
  End If

  json = json & "}"
  SerializeControl = json
End Function

Private Function ControlCanContainChildren(ByVal controlType As String) As Boolean
  Select Case LCase$(Trim$(controlType))
    Case "frame", "multipage", "page", "tabstrip"
      ControlCanContainChildren = True
    Case Else
      ControlCanContainChildren = False
  End Select
End Function

Private Function SerializeControlList(ByVal control As Object, ByVal listCount As Long) As String
  Dim json As String
  Dim i As Long
  Dim first As Boolean
  Dim itemValue As String

  json = "["
  first = True

  If listCount > 0 Then
    For i = 0 To listCount - 1
      If TryGetListItem(control, i, itemValue) Then
        If Not first Then
          json = json & ","
        End If
        json = json & JsonQuote(itemValue)
        first = False
      End If
    Next i
  End If

  json = json & "]"
  SerializeControlList = json
End Function

Private Function TryGetListItem(ByVal control As Object, ByVal index As Long, ByRef valueOut As String) As Boolean
  On Error GoTo Missing

  Dim itemValue As Variant
  itemValue = control.List(index)
  If IsNull(itemValue) Or IsEmpty(itemValue) Then
    valueOut = vbNullString
  Else
    valueOut = CStr(itemValue)
  End If
  TryGetListItem = True
  Exit Function

Missing:
  valueOut = vbNullString
  TryGetListItem = False
End Function

Private Function GetObjectControls(ByVal target As Object) As Object
  On Error Resume Next
  Set GetObjectControls = target.Controls
  On Error GoTo 0
End Function

Private Sub JsonAddStringFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As String
  If TryGetStringMember(target, memberName, value) Then
    JsonAddString json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddNumberFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Double
  If TryGetNumberMember(target, memberName, value) Then
    JsonAddNumber json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddLongFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Long
  If TryGetLongMember(target, memberName, value) Then
    JsonAddLong json, jsonKey, value, hasFields
  End If
End Sub

Private Sub JsonAddBoolFromMember(ByRef json As String, ByVal target As Object, ByVal memberName As String, ByVal jsonKey As String, ByRef hasFields As Boolean)
  Dim value As Boolean
  If TryGetBoolMember(target, memberName, value) Then
    JsonAddBool json, jsonKey, value, hasFields
  End If
End Sub

Private Function TryGetStringMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As String) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CStr(rawValue)
  TryGetStringMember = True
  Exit Function

Missing:
  valueOut = vbNullString
  TryGetStringMember = False
End Function

Private Function TryGetNumberMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Double) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CDbl(rawValue)
  TryGetNumberMember = True
  Exit Function

Missing:
  valueOut = 0
  TryGetNumberMember = False
End Function

Private Function TryGetLongMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Long) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CLng(rawValue)
  TryGetLongMember = True
  Exit Function

Missing:
  valueOut = 0
  TryGetLongMember = False
End Function

Private Function TryGetBoolMember(ByVal target As Object, ByVal memberName As String, ByRef valueOut As Boolean) As Boolean
  On Error GoTo Missing

  Dim rawValue As Variant
  rawValue = CallByName(target, memberName, VbGet)
  If IsObject(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
    GoTo Missing
  End If

  valueOut = CBool(rawValue)
  TryGetBoolMember = True
  Exit Function

Missing:
  valueOut = False
  TryGetBoolMember = False
End Function

Private Sub JsonAddRaw(ByRef json As String, ByVal key As String, ByVal rawValue As String, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & rawValue
  hasFields = True
End Sub

Private Sub JsonAddString(ByRef json As String, ByVal key As String, ByVal value As String, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & JsonQuote(value)
  hasFields = True
End Sub

Private Sub JsonAddNumber(ByRef json As String, ByVal key As String, ByVal value As Double, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & Trim$(Str$(value))
  hasFields = True
End Sub

Private Sub JsonAddLong(ByRef json As String, ByVal key As String, ByVal value As Long, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & CStr(value)
  hasFields = True
End Sub

Private Sub JsonAddBool(ByRef json As String, ByVal key As String, ByVal value As Boolean, ByRef hasFields As Boolean)
  If hasFields Then
    json = json & ","
  End If
  json = json & JsonQuote(key) & ":" & IIf(value, "true", "false")
  hasFields = True
End Sub

Private Function JsonQuote(ByVal value As String) As String
  JsonQuote = """" & JsonEscape(value) & """"
End Function

Private Function JsonEscape(ByVal value As String) As String
  Dim text As String
  text = value
  text = Replace(text, "\", "\\")
  text = Replace(text, """", Chr$(92) & Chr$(34))
  text = Replace(text, vbCrLf, "\n")
  text = Replace(text, vbCr, "\n")
  text = Replace(text, vbLf, "\n")
  text = Replace(text, vbTab, "\t")
  JsonEscape = text
End Function
'@
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

  try {
    return $Workbook.Worksheets.Item($Sheet)
  } catch {
    return $null
  }
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
