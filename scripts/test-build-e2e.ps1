[CmdletBinding()]
param(
    [switch]$KeepWorkspace,
    [ValidateSet('.xlsm', '.xlam', '.xlsb')]
    [string[]]$WorkbookExtensions = @('.xlsm', '.xlam', '.xlsb'),
    [string]$WorkspaceSuffix = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = Join-Path $repoRoot 'tmp_workspaces'

function Invoke-XlflowJson {
    param([Parameter(Mandatory)][string[]]$Arguments, [switch]$AllowFailure)

    # xlflow emits normal progress on stderr. Keep stdout clean for the JSON
    # envelope while preserving stderr for failure diagnostics.
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
    $json = $raw | ConvertFrom-Json
    if (-not $AllowFailure -and ($exitCode -ne 0 -or $json.status -ne 'ok')) {
        throw "xlflow $($Arguments -join ' ') failed: $raw$stderr"
    }
    return @{ ExitCode = $exitCode; Json = $json; Raw = $raw; Stderr = $stderr }
}

function Write-ReleaseSources {
    param([Parameter(Mandatory)][string]$Path)

    Set-Content -LiteralPath (Join-Path $Path 'src\modules\Main.bas') -Encoding ascii @'
Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    ThisWorkbook.Worksheets(1).Range("A1").Value = "build e2e ok"
End Sub
'@
    Set-Content -LiteralPath (Join-Path $Path 'src\classes\ReleaseService.cls') -Encoding ascii @'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReleaseService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DescribeRelease() As String
    DescribeRelease = "release"
End Function
'@
    New-Item -ItemType Directory -Force -Path (Join-Path $Path 'src\modules\Tests') | Out-Null
    Set-Content -LiteralPath (Join-Path $Path 'src\modules\Tests\DevOnly.bas') -Encoding ascii @'
Attribute VB_Name = "DevOnly"
Option Explicit
Public Sub ShouldNotShip()
End Sub
'@
    Add-Content -LiteralPath (Join-Path $Path 'xlflow.toml') -Encoding ascii @'

[build]
exclude = ["src/modules/Tests/**"]
'@
}

function Assert-BuildArtifact {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Extension, [switch]$ExpectUserForm)

    $dry = Invoke-XlflowJson @('build', '--dry-run', '--json')
    if ($dry.Json.build.schema_version -ne 1 -or $dry.Json.build.validation.vbe_compile -ne 'not_run') {
        throw "dry-run did not return the v1 build manifest shape: $($dry.Raw)"
    }
    $output = Join-Path $Path (Join-Path 'build\Release' ("Release$Extension"))
    if (Test-Path -LiteralPath $output) { throw "dry-run unexpectedly created $output" }

    $result = Invoke-XlflowJson @('build', '--base', ("build/Release$Extension"), '--out', ("build/Release/Release$Extension"), '--json')
    if (-not (Test-Path -LiteralPath $output)) { throw "build did not create $output" }
    $manifestPath = "$output.build.json"
    if (-not (Test-Path -LiteralPath $manifestPath)) { throw "build did not create $manifestPath" }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    if ($manifest.schema_version -ne 1 -or $manifest.validation.vbe_compile -ne 'passed' -or -not $result.Json.build.manifest.published) {
        throw "build manifest did not record a successful validation: $($result.Raw)"
    }
    if ($manifest.included_components.name -notcontains 'Main' -or $manifest.excluded_components.name -notcontains 'DevOnly') {
        throw "build manifest did not report the expected included/excluded components"
    }

    # Prove replacement and the published workbook's VBA project, not only CLI output.
    $second = Invoke-XlflowJson @('build', '--base', ("build/Release$Extension"), '--out', ("build/Release/Release$Extension"), '--json')
    if (-not $second.Json.output.replaced_existing) { throw "second build did not report replacement" }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($output)
        try {
            $components = @($wb.VBProject.VBComponents | ForEach-Object { $_.Name })
            foreach ($required in @('Main', 'ReleaseService', 'ThisWorkbook')) {
                if ($components -notcontains $required) { throw "published $Extension workbook is missing $required" }
            }
            if ($ExpectUserForm -and $components -notcontains 'ReleaseForm') { throw "published $Extension workbook is missing ReleaseForm" }
            if ($components -contains 'DevOnly') { throw "published $Extension workbook retained excluded DevOnly" }
        } finally {
            $wb.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
        }
    } finally {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if (-not (Get-Command xlflow -ErrorAction SilentlyContinue)) {
    throw 'xlflow was not found on PATH. Run task install before scripts/test-build-e2e.ps1.'
}

New-Item -ItemType Directory -Force -Path $workspaceRoot | Out-Null
foreach ($extension in $WorkbookExtensions) {
    $workspace = Join-Path $workspaceRoot ("build-e2e-" + $extension.TrimStart('.') + $WorkspaceSuffix)
    if (Test-Path -LiteralPath $workspace) { Remove-Item -LiteralPath $workspace -Recurse -Force }
    New-Item -ItemType Directory -Path $workspace | Out-Null
    try {
        Push-Location $workspace
        Invoke-XlflowJson @('new', ("Release$extension"), '--json') | Out-Null
        Write-ReleaseSources $workspace
        $expectUserForm = $extension -eq '.xlsm'
        if ($expectUserForm) {
            Invoke-XlflowJson @('form', 'new', 'ReleaseForm', '--json') | Out-Null
            Invoke-XlflowJson @('form', 'build', 'src/forms/specs/ReleaseForm.yaml', '--json') | Out-Null
        }
        # Canonicalize the hand-authored fixture through Excel's own export
        # format before the release build. This also proves that push ignores
        # [build].exclude: DevOnly is imported into the development workbook
        # and then excluded only from the build artifact below.
        Invoke-XlflowJson @('push', '--json') | Out-Null
        Invoke-XlflowJson @('pull', '--json') | Out-Null
        Assert-BuildArtifact $workspace $extension -ExpectUserForm:$expectUserForm
        Write-Output "build E2E passed: $workspace"
    } finally {
        Pop-Location
        if (-not $KeepWorkspace -and (Test-Path -LiteralPath $workspace)) { Remove-Item -LiteralPath $workspace -Recurse -Force }
    }
}
