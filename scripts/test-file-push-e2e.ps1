[CmdletBinding()]
param(
    [switch]$KeepWorkspace,
    [string]$WorkspaceSuffix = ''
)

$ErrorActionPreference = 'Stop'
$directorySeparators = [char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
if ($WorkspaceSuffix.IndexOfAny($directorySeparators) -ge 0) {
    throw 'WorkspaceSuffix must not contain directory separators.'
}
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces'))
$workspace = [IO.Path]::GetFullPath((Join-Path $workspaceRoot ("file-push-release-e2e$WorkspaceSuffix")))
$utf8NoBom = [Text.UTF8Encoding]::new($false, $true)
$japaneseModuleName = (-join [char[]](0x65E5, 0x672C, 0x8A9E)) + 'Module'
$japaneseMarker = -join [char[]](0x65E5, 0x672C, 0x8A9E, 0x78BA, 0x8A8D)
$japaneseComment = -join [char[]](0x65E5, 0x672C, 0x8A9E, 0x30B3, 0x30E1, 0x30F3, 0x30C8)
$sentinel = "file push ok|class ok|sheet ok|workbook ok|$japaneseMarker"

function Invoke-XlflowJson {
    param([Parameter(Mandatory)][string[]]$Arguments, [switch]$AllowFailure)

    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $stdoutPath = [IO.Path]::GetTempFileName()
    $stderrPath = [IO.Path]::GetTempFileName()
    try {
        & xlflow @Arguments 1>$stdoutPath 2>$stderrPath
        $exitCode = $LASTEXITCODE
        $raw = Get-Content -LiteralPath $stdoutPath -Raw
        $stderr = Get-Content -LiteralPath $stderrPath -Raw
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "xlflow $($Arguments -join ' ') produced no JSON (exit $exitCode): $stderr"
    }
    $json = $raw | ConvertFrom-Json
    if (-not $AllowFailure -and ($exitCode -ne 0 -or $json.status -ne 'ok')) {
        throw "xlflow $($Arguments -join ' ') failed: $raw$stderr"
    }
    return @{ ExitCode = $exitCode; Json = $json; Raw = $raw; Stderr = $stderr }
}

function Write-Utf8NoBom {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Content)

    $parent = Split-Path -Parent $Path
    if ($parent) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    [IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Assert-FilePushedSentinel {
    param([Parameter(Mandatory)][string]$WorkbookPath)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $null
    $sheet = $null
    $project = $null
    try {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        # Force VBE compile of the rebuilt project before executing: a file
        # push is structurally valid only, so this is the real-Excel proof.
        $project = $workbook.VBProject
        $compile = $null
        try {
            $compile = $excel.VBE.CommandBars.FindControl(1, 578)
            if ($null -ne $compile) { $compile.Execute() }
        } catch {
            # FindControl is best-effort; Application.Run below still forces
            # on-demand compilation before the macro body executes.
            Write-Verbose "VBE compile menu not available: $_"
        } finally {
            Release-ComObject $compile
        }
        [void]$excel.Run("'$($workbook.Name)'!Main.Run")
        $sheet = $workbook.Worksheets.Item(1)
        $actual = [string]$sheet.Range('A1').Value2
        if ($actual -cne $sentinel) {
            throw "file-pushed sentinel was '$actual', want '$sentinel'"
        }
        return [pscustomobject]@{
            Sentinel = $actual
            ExcelVersion = $excel.Version
            ExcelOperatingSystem = $excel.OperatingSystem
        }
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        $excel.Quit()
        Release-ComObject $sheet
        Release-ComObject $project
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if (-not (Get-Command xlflow -ErrorAction SilentlyContinue)) {
    throw 'xlflow was not found on PATH. Run task install before scripts/test-file-push-e2e.ps1.'
}

New-Item -ItemType Directory -Force -Path $workspaceRoot | Out-Null
$workspaceBoundary = $workspaceRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
if (-not $workspace.StartsWith($workspaceBoundary, [StringComparison]::OrdinalIgnoreCase)) {
    throw "unsafe release-gate workspace path: $workspace"
}
if (Test-Path -LiteralPath $workspace) {
    Remove-Item -LiteralPath $workspace -Recurse -Force
}

try {
    New-Item -ItemType Directory -Path $workspace | Out-Null
    Push-Location $workspace
    try {
        Invoke-XlflowJson @('new', 'FilePushGate.xlsm', '--no-update-check', '--json') | Out-Null
        $workbookPath = Join-Path $workspace 'build\FilePushGate.xlsm'
        $workbookBytesBefore = [Convert]::ToBase64String([IO.File]::ReadAllBytes($workbookPath))

        # Replace the scaffolded source with the gate's full module tree:
        # nested folders, a class, document modules, and non-ASCII names.
        $scaffoldMain = Join-Path $workspace 'src\modules\Main.bas'
        if (Test-Path -LiteralPath $scaffoldMain) { Remove-Item -LiteralPath $scaffoldMain -Force }

        Write-Utf8NoBom (Join-Path $workspace 'src\modules\Domain\Core\Main.bas') @"
Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim service As Greeting
    Set service = New Greeting
    ThisWorkbook.Worksheets(1).Range("A1").Value = TextHelpers.Prefix() & "|" & service.Message() & "|" & Sheet1.DocumentMarker() & "|" & ThisWorkbook.WorkbookMarker() & "|" & $japaneseModuleName.JapaneseMarker()
End Sub
"@
        Write-Utf8NoBom (Join-Path $workspace 'src\modules\Shared\TextHelpers.bas') @'
Attribute VB_Name = "TextHelpers"
Option Explicit

Public Function Prefix() As String
    Prefix = "file push ok"
End Function
'@
        Write-Utf8NoBom (Join-Path $workspace "src\modules\International\$japaneseModuleName.bas") @"
Attribute VB_Name = "$japaneseModuleName"
Option Explicit
' $japaneseComment

Public Function JapaneseMarker() As String
    JapaneseMarker = "$japaneseMarker"
End Function
"@
        Write-Utf8NoBom (Join-Path $workspace 'src\classes\Domain\Services\Greeting.cls') @'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Greeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Message() As String
    Message = "class ok"
End Function
'@
        Write-Utf8NoBom (Join-Path $workspace 'src\workbook\Sheet1.bas') @'
Option Explicit

Public Function DocumentMarker() As String
    DocumentMarker = "sheet ok"
End Function
'@
        Write-Utf8NoBom (Join-Path $workspace 'src\workbook\ThisWorkbook.bas') @'
Option Explicit

Public Function WorkbookMarker() As String
    WorkbookMarker = "workbook ok"
End Function
'@

        # Safety gate: an Office lock file must block the push before mutation.
        $lockPath = Join-Path $workspace 'build\~$FilePushGate.xlsm'
        Write-Utf8NoBom $lockPath 'lock'
        $locked = Invoke-XlflowJson @('push', '--backend', 'file', '--json') -AllowFailure
        if ($locked.ExitCode -eq 0 -or $locked.Json.error.code -ne 'push_workbook_open') {
            throw "lock-file rejection contract failed: $($locked.Raw)$($locked.Stderr)"
        }
        if ([Convert]::ToBase64String([IO.File]::ReadAllBytes($workbookPath)) -ne $workbookBytesBefore) {
            throw 'file push modified the workbook despite the lock file'
        }
        Remove-Item -LiteralPath $lockPath -Force

        # Safety gate: a recorded matching session must block the push even
        # when the recorded PID is not alive.
        New-Item -ItemType Directory -Force -Path (Join-Path $workspace '.xlflow') | Out-Null
        Write-Utf8NoBom (Join-Path $workspace '.xlflow\session.json') '{"pid":424242,"workbook_path":"build/FilePushGate.xlsm"}'
        $sessionBlocked = Invoke-XlflowJson @('push', '--backend', 'file', '--json') -AllowFailure
        if ($sessionBlocked.ExitCode -eq 0 -or $sessionBlocked.Json.error.code -ne 'push_active_session') {
            throw "session-record rejection contract failed: $($sessionBlocked.Raw)$($sessionBlocked.Stderr)"
        }
        Remove-Item -LiteralPath (Join-Path $workspace '.xlflow\session.json') -Force

        # The actual file push: headless rebuild + atomic replace.
        $filePush = Invoke-XlflowJson @('push', '--backend', 'file', '--json')
        if ($filePush.Json.push.backend -ne 'file' -or
            $filePush.Json.push.target -ne 'saved_workbook' -or
            $filePush.Json.push.vbe_validation -ne 'not_performed') {
            throw "file push JSON lost backend authority: $($filePush.Raw)"
        }
        $skipWarning = @($filePush.Json.warnings | Where-Object { $_.code -eq 'vbe_validation_skipped' })
        if ($skipWarning.Count -eq 0) {
            throw "file push did not report vbe_validation_skipped: $($filePush.Raw)"
        }
        if ($null -eq $filePush.Json.backup -or $filePush.Json.backup.mode -ne 'always') {
            throw "default file push did not record a backup: $($filePush.Raw)"
        }
        if (-not (Test-Path -LiteralPath (Join-Path $workspace '.xlflow\state\push.json'))) {
            throw 'file push did not write .xlflow/state/push.json'
        }

        # Real Excel proof: open the rebuilt workbook, compile, and run the
        # sentinel macro. This is the boundary the file backend cannot check.
        $excelResult = Assert-FilePushedSentinel $workbookPath

        # changed-only must skip when nothing changed.
        $skipped = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json')
        if ($skipped.Json.source.changed -ne $false) {
            throw "unchanged file push did not skip: $($skipped.Raw)"
        }
        if ($null -ne $skipped.Json.backup) {
            throw "skipped file push created a backup: $($skipped.Raw)"
        }

        # A source edit must defeat the skip and re-publish.
        Write-Utf8NoBom (Join-Path $workspace 'src\modules\Shared\TextHelpers.bas') @'
Attribute VB_Name = "TextHelpers"
Option Explicit

Public Function Prefix() As String
    Prefix = "file push ok"
End Function

Public Function Added() As String
    Added = "added"
End Function
'@
        $repushed = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json')
        if ($repushed.Json.source.changed -ne $true) {
            throw "changed file push reported a skip: $($repushed.Raw)"
        }

        # Cross-backend interop: an Excel push --changed-only inside a live
        # session must read the file backend's push.json schema without
        # crashing. (Whether it re-imports depends on transform parity; the
        # release gate asserts the command succeeds, then save --session and
        # session stop leave the workbook closed for the next file push.)
        Invoke-XlflowJson @('session', 'start', '--json') | Out-Null
        try {
            Invoke-XlflowJson @('push', '--session', '--changed-only', '--json') | Out-Null
            Invoke-XlflowJson @('save', '--session', '--json') | Out-Null
        } finally {
            Invoke-XlflowJson @('session', 'stop', '--json') | Out-Null
        }

        # The session file must be gone again and a file push must run.
        $afterSession = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json')
        if ($afterSession.Json.push.backend -ne 'file') {
            throw "file push after Excel session lost backend authority: $($afterSession.Raw)"
        }

        # Argument contract.
        $badCombo = Invoke-XlflowJson @('push', '--backend', 'file', '--session', '--json') -AllowFailure
        if ($badCombo.ExitCode -eq 0 -or $badCombo.Json.error.code -ne 'push_args_invalid') {
            throw "--backend file --session was not rejected: $($badCombo.Raw)"
        }
    } finally {
        Pop-Location
    }

    Write-Output "file push release gate passed: workspace=$workspace sentinel=$($excelResult.Sentinel) Excel=$($excelResult.ExcelVersion) OS=$($excelResult.ExcelOperatingSystem)"
} finally {
    if (-not $KeepWorkspace -and (Test-Path -LiteralPath $workspace)) {
        Remove-Item -LiteralPath $workspace -Recurse -Force
    }
}
