[CmdletBinding()]
param([Parameter(Mandatory)][string]$WorkspacePath)

# Developer-only gate: production pack and ordinary tests never launch Excel.
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$WorkspacePath = [IO.Path]::GetFullPath($WorkspacePath)
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces')) + [IO.Path]::DirectorySeparatorChar
if (-not $WorkspacePath.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Workspace must be below repo/tmp_workspaces' }
if (Test-Path -LiteralPath $WorkspacePath) { throw 'A fresh workspace is required' }
$utf8 = [Text.UTF8Encoding]::new($false)
foreach ($dir in @('src/modules', 'src/classes', 'src/workbook', 'src/forms/specs', 'src/forms/code', 'dist')) {
    [void][IO.Directory]::CreateDirectory((Join-Path $WorkspacePath $dir))
}
$configText = @'
[project]
name = "Issue886"
[workbook]
path = "build/Book.xlsm"
[userform]
code_source = "sidecar"
'@
[IO.File]::WriteAllText((Join-Path $WorkspacePath 'xlflow.toml'), $configText, $utf8)
foreach ($name in @('ThisWorkbook', 'Sheet1')) {
    [IO.File]::WriteAllText((Join-Path $WorkspacePath "src/workbook/$name.bas"), "Option Explicit`n", $utf8)
}
$expectedPath = Join-Path $WorkspacePath 'expected.json'
$fixture = Join-Path $repoRoot 'internal/vba/userforms/compiler/testdata/generation-excel-generated/expected.json'
[IO.File]::Copy($fixture, $expectedPath)
$expected = [IO.File]::ReadAllText($fixture) | ConvertFrom-Json
foreach ($form in $expected.forms) {
    # Observations remain in expected.json; only supported authoring intent
    # enters the generated source tree.
    $authoring = ($form | ConvertTo-Json -Depth 60 | ConvertFrom-Json)
    foreach ($control in $authoring.controls) {
        foreach ($field in @('observed', 'unsupported', 'properties', 'list', 'selectedIndex')) { $control.PSObject.Properties.Remove($field) }
    }
    $authoring.form.PSObject.Properties.Remove('observed')
    $authoring.PSObject.Properties.Remove('warnings')
    $specPath = Join-Path $WorkspacePath ("src/forms/specs/" + $form.form.name + '.json')
    [IO.File]::WriteAllText($specPath, ($authoring | ConvertTo-Json -Depth 60), $utf8)
}
$code = @'
Option Explicit
Public Function VerifyGeneration() As Boolean
    Dim textControl As MSForms.TextBox
    Set textControl = Me.TextBoxMain
    VerifyGeneration = Me.Controls.Count = 11 And CBool(Me.ToggleButtonMain.Value) And CLng(Me.SpinButtonMain.Value) = 12 And CLng(Me.ScrollBarMain.Value) = 34 And TypeName(textControl) = "TextBox"
End Function
'@
[IO.File]::WriteAllText((Join-Path $WorkspacePath 'src/forms/code/GeneratedForm.bas'), $code, $utf8)
$main = @'
Attribute VB_Name = "Main"
Option Explicit
Public Sub RunGenerationSentinel()
    Dim generated As GeneratedForm, emptyForm As GeneratedEmptyForm
    Set generated = New GeneratedForm
    Set emptyForm = New GeneratedEmptyForm
    If Not generated.VerifyGeneration() Or emptyForm.Controls.Count <> 0 Then Err.Raise 5
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-883-ok"
    ThisWorkbook.Worksheets(1).Range("B1").Value2 = generated.InsideWidth
    ThisWorkbook.Worksheets(1).Range("B2").Value2 = generated.InsideHeight
    Unload generated
    Unload emptyForm
End Sub
'@
[IO.File]::WriteAllText((Join-Path $WorkspacePath 'src/modules/Main.bas'), $main, $utf8)
Push-Location $WorkspacePath
try {
    $raw = & xlflow --json pack --blank --out dist/Generated.xlsm
    if ($LASTEXITCODE -ne 0) { throw "blank pack failed: $raw" }
    [IO.File]::WriteAllText((Join-Path $WorkspacePath 'pack.json'), ($raw -join "`n"), $utf8)
    $result = ($raw -join "`n") | ConvertFrom-Json
    if ($result.pack.backend -ne 'pure-go' -or $result.pack.vbe_validation -ne 'not_performed' -or $result.pack.modules.form -ne 2) { throw 'Unexpected pack contract' }
} finally { Pop-Location }
& (Join-Path $PSScriptRoot 'test-userform-generation-e2e.ps1') -Phase verify -WorkbookPath (Join-Path $WorkspacePath 'dist/Generated.xlsm') -ExpectedPath $expectedPath -NormalizedWorkbookPath (Join-Path $WorkspacePath 'normalized.xlsm')
if (-not $?) { throw 'Excel reference/generation verification failed' }
Write-Output "Forms reference and blank pack gate passed: $WorkspacePath"
