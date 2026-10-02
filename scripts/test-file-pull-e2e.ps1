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
$workspace = [IO.Path]::GetFullPath((Join-Path $workspaceRoot ("file-pull-release-e2e$WorkspaceSuffix")))
$userFormWorkspace = [IO.Path]::GetFullPath((Join-Path $workspaceRoot ("file-pull-userform-e2e$WorkspaceSuffix")))
$utf8NoBom = [Text.UTF8Encoding]::new($false, $true)
$japaneseModuleName = (-join [char[]](0x65E5, 0x672C, 0x8A9E)) + 'Module'
$japaneseMarker = -join [char[]](0x65E5, 0x672C, 0x8A9E, 0x78BA, 0x8A8D)
$japaneseComment = -join [char[]](0x65E5, 0x672C, 0x8A9E, 0x30B3, 0x30E1, 0x30F3, 0x30C8)
$sentinel = "file pull ok|class ok|sheet ok|workbook ok|$japaneseMarker"

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

function Get-SourceSnapshot {
    param([Parameter(Mandatory)][string]$Root)

    $snapshot = [ordered]@{}
    foreach ($sourceRoot in @('src\modules', 'src\classes', 'src\workbook')) {
        $absoluteRoot = Join-Path $Root $sourceRoot
        if (-not (Test-Path -LiteralPath $absoluteRoot)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $absoluteRoot -File -Recurse | Sort-Object FullName) {
            $bytes = [IO.File]::ReadAllBytes($file.FullName)
            if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) {
                throw "tracked source has a UTF-8 BOM: $($file.FullName)"
            }
            $text = $utf8NoBom.GetString($bytes)
            if (-not ($text.EndsWith("`n") -or $text.EndsWith("`r"))) {
                throw "tracked source has no final newline: $($file.FullName)"
            }
            $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
            if (-not $file.FullName.StartsWith($fullRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
                throw "source escaped snapshot root: $($file.FullName)"
            }
            $relative = $file.FullName.Substring($fullRoot.Length).TrimStart('\', '/').Replace('\', '/')
            $snapshot[$relative] = $text.Replace("`r`n", "`n").Replace("`r", "`n")
        }
    }
    return $snapshot
}

function Get-SourceByteSnapshot {
    param([Parameter(Mandatory)][string]$Root)

    $snapshot = [ordered]@{}
    $sourceRoot = [IO.Path]::GetFullPath((Join-Path $Root 'src')).TrimEnd('\', '/')
    if (-not (Test-Path -LiteralPath $sourceRoot)) { return $snapshot }
    foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -File -Recurse | Sort-Object FullName) {
        if (-not $file.FullName.StartsWith($sourceRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            throw "source escaped byte snapshot root: $($file.FullName)"
        }
        $relative = $file.FullName.Substring($sourceRoot.Length).TrimStart('\', '/').Replace('\', '/')
        $snapshot[$relative] = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file.FullName))
    }
    return $snapshot
}

function Assert-SnapshotsEqual {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Expected,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Actual
    )

    $expectedKeys = @($Expected.Keys | Sort-Object)
    $actualKeys = @($Actual.Keys | Sort-Object)
    if (($expectedKeys -join "`n") -ne ($actualKeys -join "`n")) {
        throw "source topology differs:`nexpected=$($expectedKeys -join ', ')`nactual=$($actualKeys -join ', ')"
    }
    foreach ($key in $expectedKeys) {
        if ($Expected[$key] -cne $Actual[$key]) {
            throw "source content differs: $key"
        }
    }
}

function Add-Worksheet {
    param([Parameter(Mandatory)][string]$WorkbookPath)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $null
    $sheet = $null
    try {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $sheet = $workbook.Worksheets.Add()
        $sheet.Name = 'Data'
        $workbook.Save()
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        $excel.Quit()
        Release-ComObject $sheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Add-UserForm {
    param([Parameter(Mandatory)][string]$WorkbookPath)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $null
    $project = $null
    $component = $null
    $codeModule = $null
    try {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $project = $workbook.VBProject
        $component = $project.VBComponents.Add(3)
        $component.Name = 'FilePullForm'
        $codeModule = $component.CodeModule
        $codeModule.AddFromString("Private Sub UserForm_Initialize()`r`n    Debug.Print `"FILE_PULL_USERFORM`"`r`nEnd Sub")
        $workbook.Save()
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        $excel.Quit()
        Release-ComObject $codeModule
        Release-ComObject $component
        Release-ComObject $project
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Assert-PackedSentinel {
    param([Parameter(Mandatory)][string]$WorkbookPath)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $null
    $sheet = $null
    try {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        [void]$excel.Run("'$($workbook.Name)'!Main.Run")
        $sheet = $workbook.Worksheets.Item(1)
        $actual = [string]$sheet.Range('A1').Value2
        if ($actual -cne $sentinel) {
            throw "packed sentinel was '$actual', want '$sentinel'"
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
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if (-not (Get-Command xlflow -ErrorAction SilentlyContinue)) {
    throw 'xlflow was not found on PATH. Run task install before scripts/test-file-pull-e2e.ps1.'
}

New-Item -ItemType Directory -Force -Path $workspaceRoot | Out-Null
$workspaceBoundary = $workspaceRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
foreach ($candidate in @($workspace, $userFormWorkspace)) {
    if (-not $candidate.StartsWith($workspaceBoundary, [StringComparison]::OrdinalIgnoreCase)) {
        throw "unsafe release-gate workspace path: $candidate"
    }
    if (Test-Path -LiteralPath $candidate) {
        Remove-Item -LiteralPath $candidate -Recurse -Force
    }
}

try {
    New-Item -ItemType Directory -Path $workspace | Out-Null
    Push-Location $workspace
    try {
        Invoke-XlflowJson @('new', 'FilePullGate.xlsm', '--no-update-check', '--json') | Out-Null
        $workbookPath = Join-Path $workspace 'build\FilePullGate.xlsm'
        Add-Worksheet $workbookPath
        Invoke-XlflowJson @('pull', '--backend', 'excel', '--json') | Out-Null
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
    Prefix = "file pull ok"
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
        Write-Utf8NoBom (Join-Path $workspace 'src\classes\Domain\Services\Counter.cls') @'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Value As Long
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

        Invoke-XlflowJson @('session', 'start', '--json') | Out-Null
        try {
            Invoke-XlflowJson @('push', '--fast', '--session', '--no-save', '--json') | Out-Null
            Invoke-XlflowJson @('save', '--session', '--json') | Out-Null
        } finally {
            Invoke-XlflowJson @('session', 'stop', '--json') | Out-Null
        }

        Invoke-XlflowJson @('pull', '--backend', 'excel', '--json') | Out-Null
        $baseline = Get-SourceSnapshot $workspace
        foreach ($sourceRoot in @('src\modules', 'src\classes', 'src\workbook')) {
            $path = Join-Path $workspace $sourceRoot
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force }
        }
        $filePull = Invoke-XlflowJson @('pull', '--backend', 'file', '--json')
        if ($filePull.Json.pull.backend -ne 'file' -or $filePull.Json.pull.source -ne 'saved_workbook') {
            throw "file pull JSON lost backend authority: $($filePull.Raw)"
        }
        $fileSnapshot = Get-SourceSnapshot $workspace
        Assert-SnapshotsEqual $baseline $fileSnapshot

        foreach ($sourceRoot in @('src\modules', 'src\classes', 'src\workbook')) {
            $path = Join-Path $workspace $sourceRoot
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force }
        }
        $autoFilePull = Invoke-XlflowJson @('pull', '--json')
        if ($autoFilePull.Json.pull.backend -ne 'file' -or
            $autoFilePull.Json.pull.backend_selection -ne 'auto' -or
            $autoFilePull.Json.pull.selection_reason -ne 'file_backend_supported' -or
            $autoFilePull.Json.pull.source -ne 'saved_workbook') {
            throw "closed-workbook auto pull did not select file: $($autoFilePull.Raw)"
        }
        Assert-SnapshotsEqual $baseline (Get-SourceSnapshot $workspace)

        $liveExcel = New-Object -ComObject Excel.Application
        $liveExcel.Visible = $true
        $liveExcel.DisplayAlerts = $false
        $liveWorkbook = $null
        $liveProject = $null
        $liveComponent = $null
        $liveCode = $null
        try {
            $liveWorkbook = $liveExcel.Workbooks.Open((Join-Path $workspace 'build\FilePullGate.xlsm'))
            $liveProject = $liveWorkbook.VBProject
            $liveComponent = $liveProject.VBComponents.Item('Main')
            $liveCode = $liveComponent.CodeModule
            $liveCode.InsertLines($liveCode.CountOfLines + 1, "`r`nPublic Function AutoLiveMarker() As String`r`n    AutoLiveMarker = `"unsaved live marker`"`r`nEnd Function")
            $autoLivePull = Invoke-XlflowJson @('pull', '--json')
            if ($autoLivePull.Json.pull.backend -ne 'excel' -or
                $autoLivePull.Json.pull.backend_selection -ne 'auto' -or
                $autoLivePull.Json.pull.selection_reason -ne 'workbook_open_in_excel' -or
                $autoLivePull.Json.pull.source -ne 'live_workbook') {
                throw "open-workbook auto pull did not preserve live authority: $($autoLivePull.Raw)"
            }
            $liveSource = [IO.File]::ReadAllText((Join-Path $workspace 'src\modules\Domain\Core\Main.bas'))
            if (-not $liveSource.Contains('AutoLiveMarker')) {
                throw 'auto Excel pull did not export the unsaved live VBA marker'
            }
        } finally {
            if ($null -ne $liveWorkbook) { $liveWorkbook.Close($false) }
            $liveExcel.Quit()
            Release-ComObject $liveCode
            Release-ComObject $liveComponent
            Release-ComObject $liveProject
            Release-ComObject $liveWorkbook
            Release-ComObject $liveExcel
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
        $restoredFilePull = Invoke-XlflowJson @('pull', '--backend', 'file', '--json')
        if ($restoredFilePull.Json.pull.backend_selection -ne 'explicit') {
            throw "explicit file pull selection metadata was lost: $($restoredFilePull.Raw)"
        }
        Assert-SnapshotsEqual $baseline (Get-SourceSnapshot $workspace)

        $pack = Invoke-XlflowJson @('pack', '--out', 'dist/FilePullRoundTrip.xlsm', '--json')
        if ($pack.Json.pack.backend -ne 'pure-go' -or $pack.Json.pack.vbe_validation -ne 'not_performed') {
            throw "pack JSON contract failed: $($pack.Raw)"
        }
        $excelResult = Assert-PackedSentinel (Join-Path $workspace 'dist\FilePullRoundTrip.xlsm')
    } finally {
        Pop-Location
    }

    New-Item -ItemType Directory -Path $userFormWorkspace | Out-Null
    Copy-Item -LiteralPath (Join-Path $workspace 'xlflow.toml') -Destination $userFormWorkspace
    Copy-Item -LiteralPath (Join-Path $workspace 'src') -Destination $userFormWorkspace -Recurse
    New-Item -ItemType Directory -Path (Join-Path $userFormWorkspace 'build') | Out-Null
    $userFormWorkbook = Join-Path $userFormWorkspace 'build\FilePullGate.xlsm'
    Copy-Item -LiteralPath (Join-Path $workspace 'build\FilePullGate.xlsm') -Destination $userFormWorkbook
    Add-UserForm $userFormWorkbook
    Push-Location $userFormWorkspace
    try {
        $excelPidsBefore = @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id | Sort-Object)
        $autoFormPull = Invoke-XlflowJson @('pull', '--json')
        $excelPidsAfter = @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id | Sort-Object)
        if ($autoFormPull.Json.pull.backend -ne 'file' -or
            $autoFormPull.Json.pull.selection_reason -ne 'file_backend_supported') {
            throw "UserForm auto pull did not select file: $($autoFormPull.Raw)"
        }
        $newExcelPids = @($excelPidsAfter | Where-Object { $excelPidsBefore -notcontains $_ })
        if ($newExcelPids.Count -gt 0) {
            throw "UserForm file pull started Excel processes: new=$($newExcelPids -join ',') before=$($excelPidsBefore -join ',') after=$($excelPidsAfter -join ',')"
        }
        $formsCanary = Join-Path $userFormWorkspace 'src\forms\release-gate-canary.frx'
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $formsCanary) | Out-Null
        [IO.File]::WriteAllBytes($formsCanary, [byte[]](0, 13, 10, 255, 128, 1))
        $formSpecPath = Join-Path $userFormWorkspace 'src\forms\specs\FilePullForm.yaml'
        $formCodePath = Join-Path $userFormWorkspace 'src\forms\code\FilePullForm.bas'
        if (-not (Test-Path -LiteralPath $formSpecPath) -or -not (Test-Path -LiteralPath $formCodePath)) {
            throw "UserForm file pull did not publish canonical artifacts: spec=$formSpecPath code=$formCodePath"
        }
        if (-not ([IO.File]::ReadAllText($formCodePath).Contains('FILE_PULL_USERFORM'))) {
            throw 'UserForm code-behind sidecar did not contain the workbook marker'
        }
        if (-not ([IO.File]::ReadAllText($formSpecPath).Contains('name: FilePullForm'))) {
            throw 'UserForm Designer YAML did not contain the workbook form identity'
        }
        $beforeRepeat = Get-SourceByteSnapshot $userFormWorkspace
        $explicitFormPull = Invoke-XlflowJson @('pull', '--backend', 'file', '--json')
        if ($explicitFormPull.Json.pull.backend_selection -ne 'explicit') {
            throw "explicit UserForm file pull selection metadata was lost: $($explicitFormPull.Raw)"
        }
        $afterRepeat = Get-SourceByteSnapshot $userFormWorkspace
        Assert-SnapshotsEqual $beforeRepeat $afterRepeat
        if (-not (Test-Path -LiteralPath $formsCanary)) {
            throw 'file pull removed an unmanaged compatibility .frx artifact'
        }
    } finally {
        Pop-Location
    }

    Write-Output "file pull release gate passed: workspace=$workspace userform_workspace=$userFormWorkspace sentinel=$($excelResult.Sentinel) Excel=$($excelResult.ExcelVersion) OS=$($excelResult.ExcelOperatingSystem)"
} finally {
    if (-not $KeepWorkspace) {
        foreach ($candidate in @($workspace, $userFormWorkspace)) {
            if (Test-Path -LiteralPath $candidate) { Remove-Item -LiteralPath $candidate -Recurse -Force }
        }
    }
}
