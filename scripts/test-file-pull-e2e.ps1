[CmdletBinding()]
param(
    [switch]$KeepWorkspace,
    [string]$WorkspaceSuffix = ''
)

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces'))
$workspace = [IO.Path]::GetFullPath((Join-Path $workspaceRoot ("file-pull-release-e2e$WorkspaceSuffix")))
$rejectionWorkspace = [IO.Path]::GetFullPath((Join-Path $workspaceRoot ("file-pull-userform-rejection-e2e$WorkspaceSuffix")))
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
    try {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $project = $workbook.VBProject
        $component = $project.VBComponents.Add(3)
        $component.Name = 'RejectedForm'
        $workbook.Save()
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        $excel.Quit()
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
foreach ($candidate in @($workspace, $rejectionWorkspace)) {
    if (-not $candidate.StartsWith($workspaceRoot, [StringComparison]::OrdinalIgnoreCase)) {
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

        $pack = Invoke-XlflowJson @('pack', '--out', 'dist/FilePullRoundTrip.xlsm', '--json')
        if ($pack.Json.pack.backend -ne 'pure-go' -or $pack.Json.pack.vbe_validation -ne 'not_performed') {
            throw "pack JSON contract failed: $($pack.Raw)"
        }
        $excelResult = Assert-PackedSentinel (Join-Path $workspace 'dist\FilePullRoundTrip.xlsm')
    } finally {
        Pop-Location
    }

    New-Item -ItemType Directory -Path $rejectionWorkspace | Out-Null
    Copy-Item -LiteralPath (Join-Path $workspace 'xlflow.toml') -Destination $rejectionWorkspace
    Copy-Item -LiteralPath (Join-Path $workspace 'src') -Destination $rejectionWorkspace -Recurse
    New-Item -ItemType Directory -Path (Join-Path $rejectionWorkspace 'build') | Out-Null
    $rejectionWorkbook = Join-Path $rejectionWorkspace 'build\FilePullGate.xlsm'
    Copy-Item -LiteralPath (Join-Path $workspace 'build\FilePullGate.xlsm') -Destination $rejectionWorkbook
    Add-UserForm $rejectionWorkbook
    Push-Location $rejectionWorkspace
    try {
        $beforeRejection = Get-SourceSnapshot $rejectionWorkspace
        $rejected = Invoke-XlflowJson @('pull', '--backend', 'file', '--json') -AllowFailure
        if ($rejected.ExitCode -eq 0 -or $rejected.Json.error.code -ne 'pull_userform_unsupported') {
            throw "UserForm rejection contract failed: $($rejected.Raw)$($rejected.Stderr)"
        }
        $afterRejection = Get-SourceSnapshot $rejectionWorkspace
        Assert-SnapshotsEqual $beforeRejection $afterRejection
    } finally {
        Pop-Location
    }

    Write-Output "file pull release gate passed: workspace=$workspace rejection_workspace=$rejectionWorkspace sentinel=$($excelResult.Sentinel) Excel=$($excelResult.ExcelVersion) OS=$($excelResult.ExcelOperatingSystem)"
} finally {
    if (-not $KeepWorkspace) {
        foreach ($candidate in @($workspace, $rejectionWorkspace)) {
            if (Test-Path -LiteralPath $candidate) { Remove-Item -LiteralPath $candidate -Recurse -Force }
        }
    }
}
