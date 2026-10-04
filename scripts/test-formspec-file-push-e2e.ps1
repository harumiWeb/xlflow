[CmdletBinding()]
param([string]$BaselineWorkbook = '', [switch]$KeepWorkspace)

# Developer-only canonical pull/push/pack gate; never run in PR CI.
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$utf8 = [Text.UTF8Encoding]::new($false)
if (-not $BaselineWorkbook -or -not (Test-Path -LiteralPath $BaselineWorkbook)) { throw 'Pass the Excel-authored baseline from test-userform-pictures-e2e.ps1.' }
$BaselineWorkbook = [IO.Path]::GetFullPath($BaselineWorkbook)
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close Excel before the file-operation gate.' }
$workspace = Join-Path $repoRoot ('tmp_workspaces/issue-912-file-push-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
[void][IO.Directory]::CreateDirectory($workspace)
$evidence = [Collections.Generic.List[object]]::new()
function Invoke-XlflowJson([string[]]$Arguments, [switch]$AllowFailure) {
    $stderrPath = Join-Path $workspace 'last-stderr.txt'
    $stdoutPath = Join-Path $workspace 'last-stdout.json'
    $previousPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        & xlflow @Arguments 1> $stdoutPath 2> $stderrPath
        $exitCode = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previousPreference }
    $raw = [IO.File]::ReadAllText($stdoutPath)
    if ([string]::IsNullOrWhiteSpace($raw)) { throw "No JSON from xlflow $($Arguments -join ' '): $([IO.File]::ReadAllText($stderrPath))" }
    $json = $raw | ConvertFrom-Json
    $evidence.Add([pscustomobject]@{ arguments = $Arguments; exitCode = $exitCode; result = $json })
    if ($exitCode -ne 0 -and -not $AllowFailure) { throw "xlflow $($Arguments -join ' ') failed: $raw $([IO.File]::ReadAllText($stderrPath))" }
    return [pscustomobject]@{ Json = $json; ExitCode = $exitCode }
}
function Get-ContentHash([string]$Path) {
    $hashAlgorithm = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hashAlgorithm.ComputeHash([IO.File]::ReadAllBytes($Path))).Replace('-', '') }
    finally { $hashAlgorithm.Dispose() }
}
function Get-AssetSnapshot {
    $result = [ordered]@{}
    foreach ($file in Get-ChildItem -LiteralPath 'src/forms/assets' -File) { $result[$file.Name] = Get-ContentHash $file.FullName }
    return $result
}
Push-Location $workspace
try {
    [void](Invoke-XlflowJson @('init', $BaselineWorkbook, '--json'))
    $configPath = Join-Path $workspace 'xlflow.toml'
    $config = [IO.File]::ReadAllText($configPath).Replace('code_source = "frm"', 'code_source = "sidecar"')
    [IO.File]::WriteAllText($configPath, $config, $utf8)
    # init exports compatibility files; this fresh workspace deliberately tests
    # canonical authority without those artifacts.
    $formsRoot = [IO.Path]::GetFullPath((Join-Path $workspace 'src/forms'))
    foreach ($compatibilityFile in Get-ChildItem -LiteralPath $formsRoot -File -Recurse | Where-Object Extension -in '.frm', '.frx') {
        if (-not $compatibilityFile.FullName.StartsWith($formsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Compatibility file escaped the fresh forms root.' }
        Remove-Item -LiteralPath $compatibilityFile.FullName
    }
    [void](Invoke-XlflowJson @('pull', '--backend', 'file', '--json'))
    if (@(Get-ChildItem 'src/forms' -Filter '*.frm' -Recurse).Count -or @(Get-ChildItem 'src/forms' -Filter '*.frx' -Recurse).Count) { throw 'Canonical gate must contain no compatibility artifacts.' }
    $assetsBefore = Get-AssetSnapshot
    if ($assetsBefore.Count -ne 2) { throw "Expected two extracted image assets, got $($assetsBefore.Count)." }
    $workbook = @(Get-ChildItem 'build' -Filter '*.xlsm')[0].FullName
    # Build a form-free saved workbook, then add its first form through file push.
    $specsRoot = Join-Path $workspace 'src/forms/specs'
    $codeRoot = Join-Path $workspace 'src/forms/code'
    $savedSpecs = Join-Path $workspace 'preserved-specs'
    $savedCode = Join-Path $workspace 'preserved-code'
    Move-Item -LiteralPath $specsRoot -Destination $savedSpecs
    $hasCodeRoot = Test-Path -LiteralPath $codeRoot
    if ($hasCodeRoot) { Move-Item -LiteralPath $codeRoot -Destination $savedCode }
    try {
        [void](Invoke-XlflowJson @('pack', '--blank', '--out', 'dist/FormFree.xlsm', '--json'))
    } finally {
        Move-Item -LiteralPath $savedSpecs -Destination $specsRoot
        if ($hasCodeRoot) { Move-Item -LiteralPath $savedCode -Destination $codeRoot }
    }
    Copy-Item -LiteralPath (Join-Path $workspace 'dist/FormFree.xlsm') -Destination $workbook -Force
    $first = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json')
    if ($first.Json.push.backend -ne 'file' -or $first.Json.source.changed -ne $true) { throw 'First canonical push was not applied.' }
    $skip = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json')
    if ($skip.Json.source.changed -ne $false) { throw 'Unchanged canonical push did not skip.' }
    [void](Invoke-XlflowJson @('pull', '--backend', 'file', '--json'))
    if ((Get-AssetSnapshot | ConvertTo-Json -Compress) -cne ($assetsBefore | ConvertTo-Json -Compress)) { throw 'Image content changed during file pull/push/pull.' }
    $specPath = Join-Path $workspace 'src/forms/specs/PictureForm.yaml'
    $spec = [IO.File]::ReadAllText($specPath)
    # Change only a spec while preserving its YAML indentation.
    $captionPattern = [regex]::new('(?m)^(\s*)caption:.*$')
    $edited = $captionPattern.Replace($spec, '$1caption: File-push pictures', 1)
    $edited = [regex]::Replace($edited, '(?m)(^\s*build:\r?\n\s*)caption:.*$', '$1caption: File-push pictures')
    if ($edited -ceq $spec) { throw 'PictureForm snapshot has no editable caption.' }
    [IO.File]::WriteAllText($specPath, $edited, $utf8)
    $specPush = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json')
    if ($specPush.Json.source.changed -ne $true) { throw 'Spec-only change did not invalidate changed-only.' }
    # Excel normalizes LoadPicture JPEGs to BMP during authoring. Explicitly
    # exercise a raw JPEG source through our generator rather than infer support.
    $imagePaths = [regex]::Matches($edited, '(?m)^\s*path:\s*(src/forms/assets/[^\r\n]+)')
    if ($imagePaths.Count -ne 2) { throw 'Expected two canonical picture paths.' }
    [void][IO.Directory]::CreateDirectory((Join-Path $workspace 'assets'))
    Copy-Item -LiteralPath (Join-Path $repoRoot 'internal/vba/userforms/compiler/testdata/pictures-excel-authored/logo.jpg') -Destination (Join-Path $workspace 'assets/logo.jpg')
    $jpegPathMatch = $imagePaths[1].Groups[1]
    $edited = $edited.Remove($jpegPathMatch.Index, $jpegPathMatch.Length).Insert($jpegPathMatch.Index, 'assets/logo.jpg')
    [IO.File]::WriteAllText($specPath, $edited, $utf8)
    [void](Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json'))
    [void](Invoke-XlflowJson @('pack', '--blank', '--out', 'dist/BlankPictures.xlsm', '--json'))
    [void](Invoke-XlflowJson @('pack', '--out', 'dist/TemplatePictures.xlsm', '--json'))
    $asset = Get-Item -LiteralPath (Join-Path $workspace $imagePaths[0].Groups[1].Value.Trim())
    $beforeRejected = Get-ContentHash $workbook
    [IO.File]::WriteAllBytes($asset.FullName, [byte[]](0, 1, 2, 3))
    $rejected = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json') -AllowFailure
    if ($rejected.ExitCode -eq 0 -or (Get-ContentHash $workbook) -cne $beforeRejected) { throw 'Malformed asset did not reject without workbook mutation.' }
    Add-Type -AssemblyName System.Drawing
    $bitmap = [Drawing.Bitmap]::new(24, 16)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([Drawing.Color]::Green)
        $bitmap.Save($asset.FullName, [Drawing.Imaging.ImageFormat]::Bmp)
    } finally { $graphics.Dispose(); $bitmap.Dispose() }
    $assetPush = Invoke-XlflowJson @('push', '--backend', 'file', '--changed-only', '--json')
    if ($assetPush.Json.source.changed -ne $true) { throw 'Asset-only change did not invalidate changed-only.' }
    [void](Invoke-XlflowJson @('pull', '--backend', 'file', '--json'))
    $assetsAfter = Get-AssetSnapshot
    if (($assetsAfter | ConvertTo-Json -Compress) -ceq ($assetsBefore | ConvertTo-Json -Compress)) { throw 'Asset replacement was not extracted.' }
    Move-Item -LiteralPath (Join-Path $workspace 'src/forms/assets') -Destination (Join-Path $workspace 'unavailable-assets')
    Move-Item -LiteralPath (Join-Path $workspace 'assets') -Destination (Join-Path $workspace 'unavailable-external-assets')
    foreach ($artifact in @($workbook, (Join-Path $workspace 'dist/BlankPictures.xlsm'), (Join-Path $workspace 'dist/TemplatePictures.xlsm'))) {
        & (Join-Path $PSScriptRoot 'test-userform-pictures-e2e.ps1') -Phase verify -WorkbookPath $artifact
        if (-not $?) { throw "Excel picture verification failed: $artifact" }
    }
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Excel cleanup incomplete.' }
    [IO.File]::WriteAllText((Join-Path $workspace 'gate-results.json'), ($evidence | ConvertTo-Json -Depth 30), $utf8)
    Write-Output "canonical FormSpec gate passed: workspace=$workspace workbook=$workbook source-assets-unavailable=True"
} finally {
    Pop-Location
    [IO.File]::WriteAllText((Join-Path $workspace 'gate-results.json'), ($evidence | ConvertTo-Json -Depth 30), $utf8)
    # Evidence is retained, including on failure. KeepWorkspace documents intent.
    if (-not $KeepWorkspace) { Write-Output "evidence retained: $workspace" }
}
