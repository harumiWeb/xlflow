[CmdletBinding()]
param(
    [string]$WorkspacePath = ''
)

# Developer-only Issue #887 gate. The caller must install xlflow first. This
# harness deliberately keeps its fresh workspace and every generated artifact.
# It uses pure-Go pack for publication and one owned Excel instance for the
# sequential save/reopen/runtime checks. Trusted VBIDE access is required only
# while creating the template fixture.
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces'))
$utf8NoBom = [Text.UTF8Encoding]::new($false, $true)
$script:comRefs = [Collections.Generic.List[object]]::new()
$script:excel = $null
$script:excelPID = 0
$script:excelStartTime = ''
$script:excelRecord = $null
$script:currentWorkbook = $null
$script:lastExcelCleanupErrors = @()
$script:lastExcelCleanupForced = $false
$script:xlflowPath = ''

if ([string]::IsNullOrWhiteSpace($WorkspacePath)) {
    $WorkspacePath = Join-Path $workspaceRoot ('issue-887-pack-userforms-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
}
$WorkspacePath = [IO.Path]::GetFullPath($WorkspacePath)
$allowedRoot = $workspaceRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
if (-not $WorkspacePath.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw "WorkspacePath must be a child of tmp_workspaces: $WorkspacePath"
}
if (Test-Path -LiteralPath $WorkspacePath) {
    throw "WorkspacePath must be fresh; refusing to delete or reuse $WorkspacePath"
}
# Reject aliases before any write, including a junction at tmp_workspaces itself.
for ($ancestor = [IO.DirectoryInfo]::new($WorkspacePath).Parent; $null -ne $ancestor; $ancestor = $ancestor.Parent) {
    if ($ancestor.Exists -and ($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw "WorkspacePath must not traverse a reparse point: $($ancestor.FullName)"
    }
}
[void][IO.Directory]::CreateDirectory($WorkspacePath)

$blankWorkspace = Join-Path $WorkspacePath 'blank'
$templateWorkspace = Join-Path $WorkspacePath 'template'
$sourceWorkspace = Join-Path $WorkspacePath 'template-source'
$emptySourceWorkspace = Join-Path $WorkspacePath 'template-empty-source'
$script:evidence = [ordered]@{
    schema_version = 1
    issue = 887
    harness = 'scripts/test-pack-userforms-e2e.ps1'
    workspace = $WorkspacePath
    invocation = [ordered]@{
        command_line = [Environment]::CommandLine
        invocation_line = $MyInvocation.Line
        script_path = [IO.Path]::GetFullPath($PSCommandPath)
        parameters = [ordered]@{ WorkspacePath = $WorkspacePath }
    }
    xlflow = [ordered]@{ command = 'xlflow'; resolved_path = ''; version = '' }
    commands = @()
    excel_instances = @()
    artifacts = @()
    cleanup_confirmed = $false
    success = $false
}

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content
    )

    $parent = Split-Path -Parent $Path
    if ($parent) { [void][IO.Directory]::CreateDirectory($parent) }
    [IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Write-Json {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$Value
    )

    $parent = Split-Path -Parent $Path
    if ($parent) { [void][IO.Directory]::CreateDirectory($parent) }
    Write-Utf8NoBom $Path (($Value | ConvertTo-Json -Depth 30) + "`n")
}

function Hold-Com {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $script:comRefs.Add($Value)
    }
    return ,$Value
}

function Release-ComRefs {
    for ($index = $script:comRefs.Count - 1; $index -ge 0; $index--) {
        $value = $script:comRefs[$index]
        try {
            if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
            }
        } catch {
            $script:currentCleanupErrors.Add($_.Exception.Message)
        }
    }
    $script:comRefs.Clear()
}

function Get-ProcessIdentity {
    param([Parameter(Mandatory)][object]$Process)

    $startTime = ''
    try { $startTime = $Process.StartTime.ToUniversalTime().ToString('o') } catch { $startTime = '<unavailable>' }
    return [ordered]@{
        pid = [int]$Process.Id
        name = [string]$Process.ProcessName
        start_time_utc = $startTime
    }
}

function Start-OwnedExcel {
    param(
        [Parameter(Mandatory)][string]$Purpose,
        [Parameter(Mandatory)][int]$AutomationSecurity
    )

    $existing = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($existing.Count -ne 0) {
        throw "Excel is already running; isolated Issue #887 harness requires no concurrent Excel instance. PIDs=$($existing.Id -join ',')"
    }

    $script:currentCleanupErrors = [Collections.Generic.List[string]]::new()
    $script:comRefs.Clear()
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false
    $script:excel.DisplayAlerts = $false
    $script:excel.EnableEvents = $false
    $script:excel.AutomationSecurity = $AutomationSecurity

    $owned = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($owned.Count -ne 1) {
        throw "Cannot establish one owned Excel process for $Purpose; observed PIDs=$($owned.Id -join ',')"
    }
    $processIdentity = Get-ProcessIdentity $owned[0]
    $script:excelPID = $processIdentity.pid
    $script:excelStartTime = $processIdentity.start_time_utc
    $script:excelRecord = [ordered]@{
        purpose = $Purpose
        process = $processIdentity
        version = [string]$script:excel.Version
        build = [string]$script:excel.Build
        operating_system = [string]$script:excel.OperatingSystem
        automation_security = $AutomationSecurity
        cleanup = $null
    }
    $script:evidence.excel_instances += $script:excelRecord
}

function Stop-OwnedExcel {
    if ($null -eq $script:excelRecord -and $null -eq $script:excel) { return }

    $cleanup = [ordered]@{
        quit_requested = $false
        exited = $false
        forced_termination = $false
        verified_process = $null
        errors = @()
    }
    try {
        if ($null -ne $script:currentWorkbook) {
            try { $script:currentWorkbook.Close($false) } catch { $script:currentCleanupErrors.Add($_.Exception.Message) }
            $script:currentWorkbook = $null
        }
        try { Release-ComRefs } catch { $script:currentCleanupErrors.Add($_.Exception.Message) }
        if ($null -ne $script:excel) {
            try {
                $script:excel.Quit()
                $cleanup.quit_requested = $true
            } catch {
                $script:currentCleanupErrors.Add($_.Exception.Message)
            }
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel) } catch { $script:currentCleanupErrors.Add($_.Exception.Message) }
            $script:excel = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if ($script:excelPID -ne 0) {
            $process = Get-Process -Id $script:excelPID -ErrorAction SilentlyContinue
            if ($null -eq $process) {
                $cleanup.exited = $true
            } else {
                # Pin the handle before checking identity. Keep this same process
                # object through waiting and termination: a PID may be reused.
                [void]$process.Handle
                $identity = Get-ProcessIdentity $process
                $cleanup.verified_process = $identity
                if ($identity.name -ne 'EXCEL' -or $identity.start_time_utc -eq '<unavailable>' -or $identity.start_time_utc -ne $script:excelStartTime) {
                    $script:currentCleanupErrors.Add("Excel PID $($script:excelPID) was reused by a different process; it was not terminated")
                } else {
                    try {
                        $cleanup.exited = $process.WaitForExit(10000)
                        if (-not $cleanup.exited) {
                            $process.Kill()
                            $cleanup.forced_termination = $true
                            $cleanup.exited = $process.WaitForExit(10000)
                        }
                        if (-not $cleanup.exited) {
                            $script:currentCleanupErrors.Add("Owned Excel PID $($script:excelPID) remained after safe termination")
                        }
                    } catch { $script:currentCleanupErrors.Add($_.Exception.Message) }
                }
                $process.Dispose()
            }
        }
    } finally {
        $cleanup.errors = @($script:currentCleanupErrors)
        if ($null -ne $script:excelRecord) { $script:excelRecord.cleanup = $cleanup }
        $script:lastExcelCleanupErrors = @($script:currentCleanupErrors)
        $script:lastExcelCleanupForced = [bool]$cleanup.forced_termination
        $script:excelPID = 0
        $script:excelStartTime = ''
        $script:excelRecord = $null
        $script:currentWorkbook = $null
        $script:comRefs.Clear()
    }
}

function Invoke-WithOwnedExcel {
    param(
        [Parameter(Mandatory)][string]$Purpose,
        [Parameter(Mandatory)][int]$AutomationSecurity,
        [Parameter(Mandatory)][scriptblock]$Body
    )

    $bodyError = $null
    $result = $null
    Start-OwnedExcel -Purpose $Purpose -AutomationSecurity $AutomationSecurity
    try {
        $result = & $Body
    } catch {
        $bodyError = $_
    } finally {
        Stop-OwnedExcel
    }
    if ($null -ne $bodyError) {
        if ($script:lastExcelCleanupErrors.Count -gt 0) {
            throw "$($bodyError.Exception.Message); Excel cleanup: $($script:lastExcelCleanupErrors -join '; ')"
        }
        throw $bodyError
    }
    if ($script:lastExcelCleanupErrors.Count -gt 0) {
        throw "Excel cleanup was not clean: $($script:lastExcelCleanupErrors -join '; ')"
    }
    if ($script:lastExcelCleanupForced) {
        throw 'Excel cleanup required forced termination; gate is failed even though the owned PID exited'
    }
    return $result
}

function Invoke-XlflowJson {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string[]]$Arguments,
        [switch]$AllowFailure
    )

    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $stdoutPath = [IO.Path]::GetTempFileName()
    $stderrPath = [IO.Path]::GetTempFileName()
    try {
        Push-Location $Root
        try {
            & xlflow @Arguments 1>$stdoutPath 2>$stderrPath
            $exitCode = $LASTEXITCODE
        } finally {
            Pop-Location
        }
        $raw = [IO.File]::ReadAllText($stdoutPath)
        $stderr = [IO.File]::ReadAllText($stderrPath)
    } finally {
        $ErrorActionPreference = $previous
        if (Test-Path -LiteralPath $stdoutPath) { Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $stderrPath) { Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue }
    }
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "xlflow $($Arguments -join ' ') produced no JSON (exit $exitCode): $stderr"
    }
    $json = $raw | ConvertFrom-Json
    $record = [ordered]@{
        working_directory = [IO.Path]::GetFullPath($Root)
        executable = $script:xlflowPath
        arguments = @($Arguments)
        invocation = 'xlflow ' + (($Arguments | ForEach-Object { if ($_ -match '\s') { '"' + $_.Replace('"', '\"') + '"' } else { $_ } }) -join ' ')
        exit_code = $exitCode
        status = [string]$json.status
        stdout = $raw
        stderr = $stderr
    }
    $script:evidence.commands += $record
    if (-not $AllowFailure -and ($exitCode -ne 0 -or $json.status -ne 'ok')) {
        throw "xlflow $($Arguments -join ' ') failed: $raw$stderr"
    }
    return [pscustomobject]@{ ExitCode = $exitCode; Json = $json; Raw = $raw; Stderr = $stderr }
}

function Set-PackTopology {
    param([Parameter(Mandatory)][string]$ConfigPath, [Parameter(Mandatory)][ValidateSet('template', 'source')][string]$Topology)

    $text = [IO.File]::ReadAllText($ConfigPath)
    if ($text -notmatch '(?m)^\s*userform_topology\s*=') {
        $text = $text.TrimEnd() + "`r`n`r`n[pack]`r`nuserform_topology = `"$Topology`"`r`n"
    } else {
        $text = [regex]::Replace($text, '(?m)^\s*userform_topology\s*=.*$', "userform_topology = `"$Topology`"")
    }
    Write-Utf8NoBom $ConfigPath $text
}

function Set-ExcelPath {
    param([Parameter(Mandatory)][string]$ConfigPath, [Parameter(Mandatory)][string]$RelativePath)

    $text = [IO.File]::ReadAllText($ConfigPath)
    $lines = [regex]::Split($text, "\r?\n")
    $section = ''
    $changed = $false
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if ($lines[$index] -match '^\s*\[([^]]+)\]') {
            $section = $matches[1].ToLowerInvariant()
            continue
        }
        if ($section -eq 'excel' -and $lines[$index] -match '^\s*path\s*=') {
            $lines[$index] = 'path = "' + $RelativePath.Replace('\', '/') + '"'
            $changed = $true
            break
        }
    }
    if (-not $changed) { throw "xlflow.toml has no [excel].path entry: $ConfigPath" }
    Write-Utf8NoBom $ConfigPath ($lines -join "`r`n")
}

function Ensure-SourceTree {
    param([Parameter(Mandatory)][string]$Root)

    foreach ($relative in @('src\modules', 'src\classes', 'src\workbook', 'src\forms', 'src\forms\specs', 'src\forms\code')) {
        [void][IO.Directory]::CreateDirectory((Join-Path $Root $relative))
    }
}

function Get-Sha256 {
    param([Parameter(Mandatory)][string]$Path)

    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($Path))).Replace('-', '').ToLowerInvariant()) }
    finally { $sha.Dispose() }
}

function Get-ReferenceSnapshot {
    param([Parameter(Mandatory)][object]$Project)

    $items = @()
    $references = Hold-Com $Project.References
    for ($index = 1; $index -le $references.Count; $index++) {
        $reference = Hold-Com ($references.Item($index))
        $items += [pscustomobject]@{
            name = [string]$reference.Name
            guid = [string]$reference.Guid
            major = [int]$reference.Major
            minor = [int]$reference.Minor
            broken = [bool]$reference.IsBroken
        }
    }
    return @($items | Sort-Object name, guid)
}

function Get-ProjectFormNames {
    param([Parameter(Mandatory)][object]$Project)

    $names = @()
    $components = Hold-Com $Project.VBComponents
    for ($index = 1; $index -le $components.Count; $index++) {
        $component = Hold-Com ($components.Item($index))
        if ([int]$component.Type -eq 3) { $names += [string]$component.Name }
    }
    return @($names | Sort-Object)
}

function Get-FormSnapshot {
    param(
        [Parameter(Mandatory)][object]$Project,
        [Parameter(Mandatory)][string]$FormName,
        [string]$ControlName = ''
    )

    $components = Hold-Com $Project.VBComponents
    $component = Hold-Com ($components.Item($FormName))
    if ([int]$component.Type -ne 3) { throw "$FormName is not a UserForm component" }
    $designer = Hold-Com $component.Designer
    $properties = Hold-Com $component.Properties
    $width = Hold-Com ($properties.Item('Width'))
    $height = Hold-Com ($properties.Item('Height'))
    $snapshot = [ordered]@{
        name = $FormName
        caption = [string]$designer.Caption
        width = [double]$width.Value
        height = [double]$height.Value
        controls = @()
    }
    if (-not [string]::IsNullOrWhiteSpace($ControlName)) {
        $controls = Hold-Com $designer.Controls
        $control = Hold-Com ($controls.Item($ControlName))
        $progId = ''
        try { $progId = [string]$control.ProgID } catch { $progId = '' }
        $snapshot.controls = @([ordered]@{
            name = [string]$control.Name
            type = $progId
            value = [string]$control.Value
            left = [double]$control.Left
            top = [double]$control.Top
            width = [double]$control.Width
            height = [double]$control.Height
        })
    }
    return [pscustomobject]$snapshot
}

function Assert-NumericEqual {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][double]$Actual, [Parameter(Mandatory)][double]$Expected)
    if ([Math]::Abs($Actual - $Expected) -gt 0.05) { throw "$Path differs: actual=$Actual expected=$Expected" }
}

function Assert-FormSnapshot {
    param([Parameter(Mandatory)][object]$Project, [Parameter(Mandatory)][object]$Expected)

    $actual = Get-FormSnapshot -Project $Project -FormName ([string]$Expected.name) -ControlName ([string]$Expected.control_name)
    if ($actual.caption -cne [string]$Expected.caption) { throw "$($Expected.name).caption differs: actual=$($actual.caption) expected=$($Expected.caption)" }
    if ($null -ne $Expected.width) { Assert-NumericEqual "$($Expected.name).width" $actual.width ([double]$Expected.width) }
    if ($null -ne $Expected.height) { Assert-NumericEqual "$($Expected.name).height" $actual.height ([double]$Expected.height) }
    if ($null -ne $Expected.control) {
        $control = @($actual.controls)[0]
        if ($control.name -cne [string]$Expected.control.name) { throw "$($Expected.name) control differs" }
        if ($null -ne $Expected.control.value -and $control.value -cne [string]$Expected.control.value) { throw "$($Expected.name).$($control.name).value differs" }
        foreach ($property in @('left', 'top', 'width', 'height')) {
            $expectedValue = $Expected.control[$property]
            if ($null -ne $expectedValue) {
                Assert-NumericEqual "$($Expected.name).$($control.name).$property" ([double]$control.$property) ([double]$expectedValue)
            }
        }
    }
}

function Assert-References {
    param([Parameter(Mandatory)][object]$Project, [object[]]$Expected)

    $actual = @(Get-ReferenceSnapshot $Project)
    if ($actual.Count -eq 0) { throw 'packed workbook has no VBA references' }
    $broken = @($actual | Where-Object broken)
    if ($broken.Count -ne 0) { throw "packed workbook has broken references: $($broken.name -join ', ')" }
    if ($null -ne $Expected) {
        $actualKeys = @($actual | ForEach-Object { "$($_.name)|$($_.guid)|$($_.major)|$($_.minor)" } | Sort-Object)
        $expectedKeys = @($Expected | ForEach-Object { "$($_.name)|$($_.guid)|$($_.major)|$($_.minor)" } | Sort-Object)
        if (($actualKeys -join "`n") -cne ($expectedKeys -join "`n")) {
            throw "packed workbook reference set differs: actual=$($actualKeys -join ', ') expected=$($expectedKeys -join ', ')"
        }
    }
    return $actual
}

function Add-Or-ReplaceMainModule {
    param([Parameter(Mandatory)][object]$Project, [Parameter(Mandatory)][string]$Source)

    $components = Hold-Com $Project.VBComponents
    $module = $null
    for ($index = 1; $index -le $components.Count; $index++) {
        $candidate = Hold-Com ($components.Item($index))
        if ([int]$candidate.Type -eq 1 -and [string]$candidate.Name -eq 'Main') {
            $module = $candidate
            break
        }
    }
    if ($null -eq $module) {
        $module = Hold-Com ($components.Add(1))
        $module.Name = 'Main'
    }
    $code = Hold-Com $module.CodeModule
    if ([int]$code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString($Source)
}

function Add-TemplateForm {
    param(
        [Parameter(Mandatory)][object]$Project,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Caption,
        [switch]$WithTextBox
    )

    $components = Hold-Com $Project.VBComponents
    $component = Hold-Com ($components.Add(3))
    $designer = Hold-Com $component.Designer
    $component.Name = $Name
    $designer.Caption = $Caption
    $properties = Hold-Com $component.Properties
    $captionProperty = Hold-Com ($properties.Item('Caption'))
    $captionProperty.Value = $Caption
    $width = Hold-Com ($properties.Item('Width'))
    $height = Hold-Com ($properties.Item('Height'))
    $width.Value = 360
    $height.Value = 240
    if ($WithTextBox) {
        $controls = Hold-Com $designer.Controls
        $control = Hold-Com ($controls.Add('Forms.TextBox.1', 'TextMain', $true))
        $control.Left = 18
        $control.Top = 24
        $control.Width = 144
        $control.Height = 24
        $control.Value = 'template text'
    }
}

function New-TemplateWorkbook {
    param([Parameter(Mandatory)][string]$WorkbookPath)

    $fixtureWorkbookPath = $WorkbookPath
    Invoke-WithOwnedExcel -Purpose 'template fixture creation' -AutomationSecurity 3 -Body {
        $books = Hold-Com $script:excel.Workbooks
        $workbook = Hold-Com ($books.Open($fixtureWorkbookPath, 0, $false))
        $script:currentWorkbook = $workbook
        $project = Hold-Com $workbook.VBProject
        Add-TemplateForm -Project $project -Name 'Login' -Caption 'Template Login' -WithTextBox
        Add-TemplateForm -Project $project -Name 'KeepForm' -Caption 'Keep this omitted template form'
        Add-Or-ReplaceMainModule -Project $project -Source @'
Option Explicit
Public Sub Run()
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "template fixture"
End Sub
'@
        $workbook.Save()
        $metadata = [ordered]@{
            workbook = $fixtureWorkbookPath
            forms = [ordered]@{
                Login = Get-FormSnapshot -Project $project -FormName 'Login' -ControlName 'TextMain'
                KeepForm = Get-FormSnapshot -Project $project -FormName 'KeepForm'
            }
            references = @(Get-ReferenceSnapshot $project)
            form_names = @(Get-ProjectFormNames $project)
            excel = [ordered]@{
                version = [string]$script:excel.Version
                build = [string]$script:excel.Build
                operating_system = [string]$script:excel.OperatingSystem
            }
        }
        $workbook.Close($false)
        $script:currentWorkbook = $null
        Release-ComRefs
        return [pscustomobject]$metadata
    }
}

function Write-BlankSources {
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][bool]$IncludeEmptyForm)

    Ensure-SourceTree $Root
    Write-Utf8NoBom (Join-Path $Root 'src\workbook\ThisWorkbook.bas') "Option Explicit`r`n"
    Write-Utf8NoBom (Join-Path $Root 'src\workbook\Sheet1.bas') "Option Explicit`r`n"
    $emptyCheck = if ($IncludeEmptyForm) {
        @'
    Dim emptyForm As Object
    Set emptyForm = VBA.UserForms.Add("EmptyForm")
    If emptyForm.Controls.Count <> 0 Then Err.Raise 5, , "EmptyForm is not empty"
'@
    } else { '' }
    $unloadEmpty = if ($IncludeEmptyForm) { '    Unload emptyForm' } else { '' }
    $sentinel = if ($IncludeEmptyForm) { 'pack blank userforms initial' } else { 'pack blank userforms removed' }
    $main = @"
Attribute VB_Name = `"Main`"
Option Explicit
Public Sub Run()
    On Error GoTo Failed
    Dim login As Object
    Set login = VBA.UserForms.Add(`"Login`")
    If login.Caption <> `"Blank Login`" Then Err.Raise 5, , `"Login caption mismatch`"
    If CStr(login.Controls(`"TextBoxLogin`").Value) <> `"blank text`" Then Err.Raise 5, , `"Login text mismatch`"
    If Abs(CDbl(login.Controls(`"TextBoxLogin`").Left) - 18#) > 0.1 Then Err.Raise 5, , `"Login left mismatch`"
__EMPTY_CHECK__
    ThisWorkbook.Worksheets(1).Range(`"A1`").Value2 = `"$sentinel`"
    Unload login
__UNLOAD_EMPTY__
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range(`"A1`").Value2 = `"pack blank userforms failed:`" & CStr(Err.Number) & `":"` & Err.Description
End Sub
"@
    $main = $main.Replace('__EMPTY_CHECK__', $emptyCheck).Replace('__UNLOAD_EMPTY__', $unloadEmpty)
    Write-Utf8NoBom (Join-Path $Root 'src\modules\Main.bas') $main
    $loginSpec = @'
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: Login
  caption: "Blank Login"
controls:
  - id: login_text
    name: TextBoxLogin
    type: TextBox
    text: "blank text"
    left: 18
    top: 24
    width: 144
    height: 24
'@
    Write-Utf8NoBom (Join-Path $Root 'src\forms\specs\Login.yaml') $loginSpec
    Write-Utf8NoBom (Join-Path $Root 'src\forms\code\Login.bas') "Option Explicit`r`n"
    if ($IncludeEmptyForm) {
        Write-Utf8NoBom (Join-Path $Root 'src\forms\specs\EmptyForm.yaml') @'
schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: EmptyForm
  caption: "Empty Form"
controls: []
'@
        Write-Utf8NoBom (Join-Path $Root 'src\forms\code\EmptyForm.bas') "Option Explicit`r`n"
    } else {
        # Exercise source removal while retaining the fresh workspace and all
        # generated artifacts. Only these harness-owned files are removed.
        foreach ($relative in @('src\forms\specs\EmptyForm.yaml', 'src\forms\code\EmptyForm.bas')) {
            $path = Join-Path $Root $relative
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        }
    }
}

function Write-TemplateSources {
    param([Parameter(Mandatory)][string]$Root)

    Ensure-SourceTree $Root
    Write-Utf8NoBom (Join-Path $Root 'src\workbook\ThisWorkbook.bas') "Option Explicit`r`n"
    Write-Utf8NoBom (Join-Path $Root 'src\workbook\Sheet1.bas') "Option Explicit`r`n"
    Write-Utf8NoBom (Join-Path $Root 'src\modules\Main.bas') @'
Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    On Error GoTo Failed
    Dim login As Object, added As Object, kept As Object
    Set login = VBA.UserForms.Add("Login")
    Set added = VBA.UserForms.Add("NewForm")
    Set kept = VBA.UserForms.Add("KeepForm")
    If login.Caption <> "Packed Login" Then Err.Raise 5, , "Login caption mismatch"
    If CStr(login.Controls("TextMain").Value) <> "packed text" Then Err.Raise 5, , "Login text mismatch"
    If Abs(CDbl(login.Controls("TextMain").Left) - 42#) > 0.1 Then Err.Raise 5, , "Login left mismatch"
    If CStr(added.Caption) <> "Added Form" Then Err.Raise 5, , "NewForm caption mismatch"
    If CStr(added.Controls("NewText").Value) <> "new form text" Then Err.Raise 5, , "NewForm text mismatch"
    If CStr(kept.Caption) <> "Keep this omitted template form" Then Err.Raise 5, , "KeepForm was not preserved"
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "pack template default"
    Unload login
    Unload added
    Unload kept
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "pack template default failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
'@
    Write-Utf8NoBom (Join-Path $Root 'src\forms\specs\Login.yaml') @'
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: Login
  caption: "Packed Login"
controls:
  - id: source_login_text
    name: TextMain
    type: TextBox
    text: "packed text"
    left: 42
'@
    Write-Utf8NoBom (Join-Path $Root 'src\forms\code\Login.bas') "Option Explicit`r`n"
    Write-Utf8NoBom (Join-Path $Root 'src\forms\specs\NewForm.yaml') @'
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: NewForm
  caption: "Added Form"
controls:
  - id: new_form_text
    name: NewText
    type: TextBox
    text: "new form text"
    left: 12
'@
    Write-Utf8NoBom (Join-Path $Root 'src\forms\code\NewForm.bas') "Option Explicit`r`n"
    Write-Utf8NoBom (Join-Path $Root 'src\forms\Login.frm') @'
VERSION 5.00
Begin VB.UserForm Login
   Caption = "Template Login"
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@
    if (@(Get-ChildItem -LiteralPath (Join-Path $Root 'src\forms') -Filter '*.frx' -File -ErrorAction SilentlyContinue).Count -ne 0) {
        throw 'Issue #887 source fixture unexpectedly requires a .frx artifact'
    }
}

function New-ScenarioProject {
    param(
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$TemplateRoot,
        [Parameter(Mandatory)][ValidateSet('login', 'empty')][string]$FormSet
    )

    [void][IO.Directory]::CreateDirectory($Destination)
    foreach ($relative in @('build', 'src\modules', 'src\classes', 'src\workbook', 'src\forms\specs', 'src\forms\code')) {
        [void][IO.Directory]::CreateDirectory((Join-Path $Destination $relative))
    }
    Copy-Item -LiteralPath (Join-Path $TemplateRoot 'xlflow.toml') -Destination (Join-Path $Destination 'xlflow.toml')
    Copy-Item -LiteralPath (Join-Path $TemplateRoot 'build\TemplateSource.xlsm') -Destination (Join-Path $Destination 'build\TemplateSource.xlsm')
    foreach ($relative in @('src\classes', 'src\workbook')) {
        $source = Join-Path $TemplateRoot $relative
        if (Test-Path -LiteralPath $source) {
            foreach ($item in Get-ChildItem -LiteralPath $source) {
                Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $Destination $relative) -Recurse
            }
        }
    }
    Set-PackTopology -ConfigPath (Join-Path $Destination 'xlflow.toml') -Topology source
    if ($FormSet -eq 'login') {
        foreach ($relative in @('src\forms\specs\Login.yaml', 'src\forms\code\Login.bas', 'src\forms\Login.frm')) {
            Copy-Item -LiteralPath (Join-Path $TemplateRoot $relative) -Destination (Join-Path $Destination $relative)
        }
        Write-Utf8NoBom (Join-Path $Destination 'src\modules\Main.bas') @'
Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    On Error GoTo Failed
    Dim login As Object
    Set login = VBA.UserForms.Add("Login")
    If login.Caption <> "Packed Login" Then Err.Raise 5, , "Login caption mismatch"
    If CStr(login.Controls("TextMain").Value) <> "packed text" Then Err.Raise 5, , "Login text mismatch"
    If Abs(CDbl(login.Controls("TextMain").Left) - 42#) > 0.1 Then Err.Raise 5, , "Login left mismatch"
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "pack template source login"
    Unload login
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "pack template source login failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
'@
    } else {
        Write-Utf8NoBom (Join-Path $Destination 'src\modules\Main.bas') @'
Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    On Error GoTo Failed
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "pack template source empty"
    Exit Sub
Failed:
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "pack template source empty failed:" & CStr(Err.Number) & ":" & Err.Description
End Sub
'@
    }
}

function Assert-PackSuccess {
    param(
        [Parameter(Mandatory)][object]$Result,
        [Parameter(Mandatory)][ValidateSet('blank', 'template')][string]$Base,
        [Parameter(Mandatory)][string]$ArtifactPath
    )

    if ($Result.Json.pack.base -ne $Base -or $Result.Json.pack.backend -ne 'pure-go' -or $Result.Json.pack.vbe_validation -ne 'not_performed') {
        throw "pack JSON contract failed for ${ArtifactPath}: $($Result.Raw)"
    }
    if ($Result.Json.pack.PSObject.Properties.Name -contains 'experimental') {
        throw "pack JSON retained the deprecated experimental marker for ${ArtifactPath}: $($Result.Raw)"
    }
    if (-not (Test-Path -LiteralPath $ArtifactPath)) { throw "pack did not publish $ArtifactPath" }
}

function Pack-Artifact {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$RelativeOutput,
        [Parameter(Mandatory)][ValidateSet('blank', 'template')][string]$Base
    )

    $result = if ($Base -eq 'blank') {
        Invoke-XlflowJson -Root $Root -Arguments @('pack', '--blank', '--out', $RelativeOutput, '--json')
    } else {
        Invoke-XlflowJson -Root $Root -Arguments @('pack', '--out', $RelativeOutput, '--json')
    }
    $artifactPath = [IO.Path]::GetFullPath((Join-Path $Root $RelativeOutput))
    Assert-PackSuccess -Result $result -Base $Base -ArtifactPath $artifactPath
    $jsonPath = [IO.Path]::ChangeExtension($artifactPath, '.pack.json')
    Write-Utf8NoBom $jsonPath $result.Raw
    return [pscustomobject]@{ Result = $result; Path = $artifactPath; JsonPath = $jsonPath }
}

function Invoke-FilePullReadback {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][string]$ArtifactPath,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$ExpectedForms
    )

    $root = Join-Path $WorkspacePath ('readback-' + $Label)
    [void][IO.Directory]::CreateDirectory($root)
    foreach ($relative in @('build', 'src\modules', 'src\classes', 'src\workbook', 'src\forms')) {
        [void][IO.Directory]::CreateDirectory((Join-Path $root $relative))
    }
    Copy-Item -LiteralPath (Join-Path $templateWorkspace 'xlflow.toml') -Destination (Join-Path $root 'xlflow.toml')
    $readbackArtifact = Join-Path $root ('build\' + [IO.Path]::GetFileName($ArtifactPath))
    Copy-Item -LiteralPath $ArtifactPath -Destination $readbackArtifact
    Set-ExcelPath -ConfigPath (Join-Path $root 'xlflow.toml') -RelativePath ('build/' + [IO.Path]::GetFileName($ArtifactPath))
    $beforePids = @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id | Sort-Object)
    $result = Invoke-XlflowJson -Root $root -Arguments @('pull', '--backend', 'file', '--json')
    $afterPids = @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id | Sort-Object)
    $newPids = @($afterPids | Where-Object { $beforePids -notcontains $_ })
    if ($newPids.Count -ne 0) { throw "pull --backend file started Excel for ${Label}: $($newPids -join ',')" }
    if ($result.Json.pull.backend -ne 'file' -or $result.Json.pull.source -ne 'saved_workbook' -or $result.Json.pull.backend_selection -ne 'explicit') {
        throw "file pull JSON contract failed for ${Label}: $($result.Raw)"
    }
    $specDir = Join-Path $root 'src\forms\specs'
    $specFiles = if (Test-Path -LiteralPath $specDir) { @(Get-ChildItem -LiteralPath $specDir -File | Where-Object Extension -in @('.yaml', '.yml', '.json')) } else { @() }
    $actualForms = @($specFiles | ForEach-Object BaseName | Sort-Object)
    $expected = @($ExpectedForms | Sort-Object)
    if (($actualForms -join "`n") -cne ($expected -join "`n")) {
        throw "file pull form topology differs for ${Label}: actual=$($actualForms -join ',') expected=$($expected -join ',')"
    }
    foreach ($formName in $expected) {
        $specFile = @($specFiles | Where-Object BaseName -eq $formName)
        if ($specFile.Count -ne 1) { throw "file pull did not produce one canonical spec for $formName in $Label" }
        $specText = [IO.File]::ReadAllText($specFile[0].FullName)
        if ($specText -notmatch "(?m)^\s*name:\s*$([regex]::Escape($formName))\s*$") {
            throw "file pull spec identity differs for $formName in $Label"
        }
    }
    $readbackJson = Join-Path $root 'pull-file.json'
    Write-Utf8NoBom $readbackJson $result.Raw
    $expectedJson = [ordered]@{
        label = $Label
        config_excel_path = ('build/' + [IO.Path]::GetFileName($ArtifactPath))
        artifact = $ArtifactPath
        backend = 'file'
        source = 'saved_workbook'
        forms = $actualForms
        no_form_specs = ($expected.Count -eq 0)
        excel_pids_before = $beforePids
        excel_pids_after = $afterPids
    }
    Write-Json (Join-Path $root 'expected.json') $expectedJson
    return [pscustomobject]@{ Root = $root; Result = $result; Forms = $actualForms }
}

function Assert-ArtifactWorkbook {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][string]$ArtifactPath,
        [Parameter(Mandatory)][string]$Sentinel,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Forms,
        [object[]]$FormExpectations = @(),
        [object[]]$ReferenceExpectations = $null
    )

    $artifactRecord = [ordered]@{
        label = $Label
        path = $ArtifactPath
        sentinel = $Sentinel
        expected_forms = @($Forms | Sort-Object)
        before_run_forms = $null
        after_reopen_forms = $null
        references = $null
        reopened_sentinel = $null
        excel = $null
    }
    $books = Hold-Com $script:excel.Workbooks
    $workbook = Hold-Com ($books.Open($ArtifactPath, 0, $false))
    $script:currentWorkbook = $workbook
    $project = Hold-Com $workbook.VBProject
    $beforeNames = @(Get-ProjectFormNames $project)
    $artifactRecord.before_run_forms = $beforeNames
    if (($beforeNames -join "`n") -cne ((@($Forms) | Sort-Object) -join "`n")) {
        throw "$Label form topology before run differs: actual=$($beforeNames -join ',') expected=$($Forms -join ',')"
    }
    foreach ($expected in $FormExpectations) { Assert-FormSnapshot -Project $project -Expected $expected }
    $beforeReferences = Assert-References -Project $project -Expected $ReferenceExpectations
    $artifactRecord.references = $beforeReferences
    $macro = "'$($workbook.Name.Replace("'", "''"))'!Main.Run"
    try {
        [void]$script:excel.Run($macro)
    } catch {
        throw "$Label macro invocation failed ($macro): $($_.Exception.Message)"
    }
    $sheet = Hold-Com ($workbook.Worksheets.Item(1))
    $cell = Hold-Com ($sheet.Range('A1'))
    if ([string]$cell.Value2 -cne $Sentinel) { throw "$Label sentinel before save was '$($cell.Value2)', expected '$Sentinel'" }
    $workbook.Save()
    $workbook.Close($false)
    $script:currentWorkbook = $null
    Release-ComRefs

    $books = Hold-Com $script:excel.Workbooks
    $workbook = Hold-Com ($books.Open($ArtifactPath, 0, $true))
    $script:currentWorkbook = $workbook
    $project = Hold-Com $workbook.VBProject
    $afterNames = @(Get-ProjectFormNames $project)
    $artifactRecord.after_reopen_forms = $afterNames
    if (($afterNames -join "`n") -cne ((@($Forms) | Sort-Object) -join "`n")) {
        throw "$Label form topology after reopen differs: actual=$($afterNames -join ',') expected=$($Forms -join ',')"
    }
    foreach ($expected in $FormExpectations) { Assert-FormSnapshot -Project $project -Expected $expected }
    $afterReferences = Assert-References -Project $project -Expected $ReferenceExpectations
    $artifactRecord.references_after_reopen = $afterReferences
    $sheet = Hold-Com ($workbook.Worksheets.Item(1))
    $cell = Hold-Com ($sheet.Range('A1'))
    $artifactRecord.reopened_sentinel = [string]$cell.Value2
    if ($artifactRecord.reopened_sentinel -cne $Sentinel) { throw "$Label sentinel after reopen was '$($artifactRecord.reopened_sentinel)', expected '$Sentinel'" }
    $artifactRecord.excel = [ordered]@{
        version = [string]$script:excel.Version
        build = [string]$script:excel.Build
        operating_system = [string]$script:excel.OperatingSystem
        process_id = $script:excelPID
    }
    $workbook.Close($false)
    $script:currentWorkbook = $null
    Release-ComRefs
    $jsonPath = Join-Path (Split-Path -Parent $ArtifactPath) ($Label + '.excel.json')
    Write-Json $jsonPath $artifactRecord
    $script:evidence.artifacts += $artifactRecord
    return [pscustomobject]$artifactRecord
}

function Assert-UnsupportedFramePicturePublication {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$SpecPath,
        [Parameter(Mandatory)][string]$ExistingArtifactPath
    )

    $beforeHash = Get-Sha256 $ExistingArtifactPath
    Write-Utf8NoBom $SpecPath @'
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: NewForm
  caption: "Unsupported Frame picture"
controls:
  - id: frame_new
    name: FrameNew
    type: Frame
    properties:
      picture: unsupported
'@
    $rejected = Invoke-XlflowJson -Root $Root -Arguments @('pack', '--out', 'dist/TemplateDefault.xlsm', '--json') -AllowFailure
    Write-Utf8NoBom (Join-Path $Root 'dist\unsupported-frame-picture-rejection.json') $rejected.Raw
    if ($rejected.ExitCode -ne 1 -or $rejected.Json.error.code -ne 'pack_userform_generation_unsupported') {
        throw "unsupported Frame picture rejection contract failed: $($rejected.Raw)$($rejected.Stderr)"
    }
    $afterHash = Get-Sha256 $ExistingArtifactPath
    if ($beforeHash -cne $afterHash) { throw 'unsupported Frame picture rejection changed the existing published artifact' }
    Write-Json (Join-Path $Root 'dist\unsupported-frame-picture-expected.json') ([ordered]@{
        error_code = 'pack_userform_generation_unsupported'
        output = $ExistingArtifactPath
        output_sha256_before = $beforeHash
        output_sha256_after = $afterHash
        unchanged = $true
    })
}

$failure = $null
try {
    $xlflowCommand = Get-Command xlflow -ErrorAction SilentlyContinue
    if ($null -eq $xlflowCommand) { throw 'xlflow was not found on PATH. Run task install before this gate.' }
    $script:xlflowPath = [string]$xlflowCommand.Source
    $script:evidence.xlflow.resolved_path = $script:xlflowPath
    $version = Invoke-XlflowJson -Root $repoRoot -Arguments @('version', '--json')
    $script:evidence.xlflow.version = $version.Raw

    [void][IO.Directory]::CreateDirectory($blankWorkspace)
    $blankNew = Invoke-XlflowJson -Root $blankWorkspace -Arguments @('new', 'BlankSource.xlsm', '--no-update-check', '--json')
    Write-Utf8NoBom (Join-Path $blankWorkspace 'new.json') $blankNew.Raw
    # Blank mode is source-authoritative even when the config is explicit.
    Set-PackTopology -ConfigPath (Join-Path $blankWorkspace 'xlflow.toml') -Topology source
    Write-BlankSources -Root $blankWorkspace -IncludeEmptyForm $true
    $blankInitial = Pack-Artifact -Root $blankWorkspace -RelativeOutput 'dist/BlankUserFormsInitial.xlsm' -Base blank
    $blankInitialExpected = [ordered]@{
        label = 'blank-initial'
        artifact = $blankInitial.Path
        sentinel = 'pack blank userforms initial'
        forms = @('EmptyForm', 'Login')
        login = [ordered]@{ name = 'Login'; caption = 'Blank Login'; control_name = 'TextBoxLogin'; control = [ordered]@{ name = 'TextBoxLogin'; value = 'blank text'; left = 18; top = 24; width = 144; height = 24 } }
        empty_form = [ordered]@{ name = 'EmptyForm'; caption = 'Empty Form' }
    }
    Write-Json (Join-Path $blankWorkspace 'dist\BlankUserFormsInitial.expected.json') $blankInitialExpected

    Write-BlankSources -Root $blankWorkspace -IncludeEmptyForm $false
    $blankRemoved = Pack-Artifact -Root $blankWorkspace -RelativeOutput 'dist/BlankUserFormsRemoved.xlsm' -Base blank
    $blankRemovedExpected = [ordered]@{
        label = 'blank-removed'
        artifact = $blankRemoved.Path
        sentinel = 'pack blank userforms removed'
        forms = @('Login')
        login = $blankInitialExpected.login
    }
    Write-Json (Join-Path $blankWorkspace 'dist\BlankUserFormsRemoved.expected.json') $blankRemovedExpected

    [void][IO.Directory]::CreateDirectory($templateWorkspace)
    $templateNew = Invoke-XlflowJson -Root $templateWorkspace -Arguments @('new', 'TemplateSource.xlsm', '--no-update-check', '--json')
    Write-Utf8NoBom (Join-Path $templateWorkspace 'new.json') $templateNew.Raw
    $templateWorkbookPath = Join-Path $templateWorkspace 'build\TemplateSource.xlsm'
    $templateMetadata = New-TemplateWorkbook -WorkbookPath $templateWorkbookPath
    Write-Json (Join-Path $templateWorkspace 'template-template.json') $templateMetadata
    Set-PackTopology -ConfigPath (Join-Path $templateWorkspace 'xlflow.toml') -Topology template
    Write-TemplateSources -Root $templateWorkspace

    $templateDefault = Pack-Artifact -Root $templateWorkspace -RelativeOutput 'dist/TemplateDefault.xlsm' -Base template
    $loginTemplate = $templateMetadata.forms.Login
    $keepTemplate = $templateMetadata.forms.KeepForm
    $templateDefaultExpected = [ordered]@{
        label = 'template-default'
        artifact = $templateDefault.Path
        sentinel = 'pack template default'
        forms = @('KeepForm', 'Login', 'NewForm')
        login = [ordered]@{
            name = 'Login'
            caption = 'Packed Login'
            width = $loginTemplate.width
            height = $loginTemplate.height
            control_name = 'TextMain'
            control = [ordered]@{
                name = 'TextMain'
                value = 'packed text'
                left = 42
                top = $loginTemplate.controls[0].top
                width = $loginTemplate.controls[0].width
                height = $loginTemplate.controls[0].height
            }
        }
        keep = [ordered]@{ name = 'KeepForm'; caption = $keepTemplate.caption; width = $keepTemplate.width; height = $keepTemplate.height }
        added = [ordered]@{ name = 'NewForm'; caption = 'Added Form'; control_name = 'NewText'; control = [ordered]@{ name = 'NewText'; value = 'new form text'; left = 12 } }
        references = @($templateMetadata.references)
    }
    Write-Json (Join-Path $templateWorkspace 'dist\TemplateDefault.expected.json') $templateDefaultExpected

    Assert-UnsupportedFramePicturePublication -Root $templateWorkspace -SpecPath (Join-Path $templateWorkspace 'src\forms\specs\NewForm.yaml') -ExistingArtifactPath $templateDefault.Path
    Write-TemplateSources -Root $templateWorkspace

    New-ScenarioProject -Destination $sourceWorkspace -TemplateRoot $templateWorkspace -FormSet login
    $sourceArtifact = Pack-Artifact -Root $sourceWorkspace -RelativeOutput 'dist/TemplateSource.xlsm' -Base template
    $sourceExpected = [ordered]@{
        label = 'template-source'
        artifact = $sourceArtifact.Path
        sentinel = 'pack template source login'
        forms = @('Login')
        login = $templateDefaultExpected.login
        references = @($templateMetadata.references)
    }
    Write-Json (Join-Path $sourceWorkspace 'dist\TemplateSource.expected.json') $sourceExpected

    New-ScenarioProject -Destination $emptySourceWorkspace -TemplateRoot $templateWorkspace -FormSet empty
    $emptySourceArtifact = Pack-Artifact -Root $emptySourceWorkspace -RelativeOutput 'dist/TemplateSourceEmpty.xlsm' -Base template
    $emptySourceExpected = [ordered]@{
        label = 'template-source-empty'
        artifact = $emptySourceArtifact.Path
        sentinel = 'pack template source empty'
        forms = @()
        references = @($templateMetadata.references)
        no_form_specs = $true
    }
    Write-Json (Join-Path $emptySourceWorkspace 'dist\TemplateSourceEmpty.expected.json') $emptySourceExpected

    Invoke-WithOwnedExcel -Purpose 'Issue #887 artifact verification' -AutomationSecurity 1 -Body {
        $null = Assert-ArtifactWorkbook -Label 'blank-initial' -ArtifactPath $blankInitial.Path -Sentinel 'pack blank userforms initial' -Forms @('EmptyForm', 'Login') -FormExpectations @(
            [ordered]@{ name = 'Login'; caption = 'Blank Login'; control_name = 'TextBoxLogin'; control = [ordered]@{ name = 'TextBoxLogin'; value = 'blank text'; left = 18; top = 24; width = 144; height = 24 } },
            [ordered]@{ name = 'EmptyForm'; caption = 'Empty Form' }
        )
        $null = Assert-ArtifactWorkbook -Label 'blank-removed' -ArtifactPath $blankRemoved.Path -Sentinel 'pack blank userforms removed' -Forms @('Login') -FormExpectations @(
            [ordered]@{ name = 'Login'; caption = 'Blank Login'; control_name = 'TextBoxLogin'; control = [ordered]@{ name = 'TextBoxLogin'; value = 'blank text'; left = 18; top = 24; width = 144; height = 24 } }
        )
        $null = Assert-ArtifactWorkbook -Label 'template-default' -ArtifactPath $templateDefault.Path -Sentinel 'pack template default' -Forms @('KeepForm', 'Login', 'NewForm') -ReferenceExpectations @($templateMetadata.references) -FormExpectations @(
            [ordered]@{ name = 'Login'; caption = 'Packed Login'; width = $loginTemplate.width; height = $loginTemplate.height; control_name = 'TextMain'; control = [ordered]@{ name = 'TextMain'; value = 'packed text'; left = 42; top = $loginTemplate.controls[0].top; width = $loginTemplate.controls[0].width; height = $loginTemplate.controls[0].height } },
            [ordered]@{ name = 'KeepForm'; caption = $keepTemplate.caption; width = $keepTemplate.width; height = $keepTemplate.height },
            [ordered]@{ name = 'NewForm'; caption = 'Added Form'; control_name = 'NewText'; control = [ordered]@{ name = 'NewText'; value = 'new form text'; left = 12 } }
        )
        $null = Assert-ArtifactWorkbook -Label 'template-source' -ArtifactPath $sourceArtifact.Path -Sentinel 'pack template source login' -Forms @('Login') -ReferenceExpectations @($templateMetadata.references) -FormExpectations @(
            [ordered]@{ name = 'Login'; caption = 'Packed Login'; width = $loginTemplate.width; height = $loginTemplate.height; control_name = 'TextMain'; control = [ordered]@{ name = 'TextMain'; value = 'packed text'; left = 42; top = $loginTemplate.controls[0].top; width = $loginTemplate.controls[0].width; height = $loginTemplate.controls[0].height } }
        )
        $null = Assert-ArtifactWorkbook -Label 'template-source-empty' -ArtifactPath $emptySourceArtifact.Path -Sentinel 'pack template source empty' -Forms @() -ReferenceExpectations @($templateMetadata.references)
    }

    [void](Invoke-FilePullReadback -Label 'blank-initial' -ArtifactPath $blankInitial.Path -ExpectedForms @('EmptyForm', 'Login'))
    [void](Invoke-FilePullReadback -Label 'blank-removed' -ArtifactPath $blankRemoved.Path -ExpectedForms @('Login'))
    [void](Invoke-FilePullReadback -Label 'template-default' -ArtifactPath $templateDefault.Path -ExpectedForms @('KeepForm', 'Login', 'NewForm'))
    [void](Invoke-FilePullReadback -Label 'template-source' -ArtifactPath $sourceArtifact.Path -ExpectedForms @('Login'))
    [void](Invoke-FilePullReadback -Label 'template-source-empty' -ArtifactPath $emptySourceArtifact.Path -ExpectedForms @())
} catch {
    $failure = $_
} finally {
    if ($null -ne $script:excel -or $null -ne $script:excelRecord) {
        Stop-OwnedExcel
    }
    $script:evidence.cleanup_confirmed = @($script:evidence.excel_instances | Where-Object {
        $null -ne $_.cleanup -and $_.cleanup.exited -and -not $_.cleanup.forced_termination -and @($_.cleanup.errors).Count -eq 0
    }).Count -eq @($script:evidence.excel_instances).Count
    $script:evidence.success = ($null -eq $failure -and $script:evidence.cleanup_confirmed)
    $script:evidence.finished_utc = [DateTime]::UtcNow.ToString('o')
    try { Write-Json (Join-Path $WorkspacePath 'run-evidence.json') $script:evidence } catch { $script:evidence.evidence_write_error = $_.Exception.Message }
}

if ($null -ne $failure) { throw $failure }
if (-not $script:evidence.cleanup_confirmed) { throw 'Issue #887 gate completed without confirmed owned Excel cleanup' }
Write-Output "pack UserForm E2E passed: workspace=$WorkspacePath evidence=$(Join-Path $WorkspacePath 'run-evidence.json')"
