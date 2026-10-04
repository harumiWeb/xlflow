[CmdletBinding()]
param(
    [ValidateSet('create', 'verify')][string]$Phase = 'create',
    [string]$WorkspacePath = '',
    [string]$WorkbookPath = '',
    [string]$FixtureDirectory = ''
)

# Developer-only Excel compatibility evidence. Never run from ordinary tests/CI.
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$utf8 = [Text.UTF8Encoding]::new($false)
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close Excel before running this isolated gate.' }
if (-not $WorkspacePath) {
    $WorkspacePath = Join-Path $repoRoot ('tmp_workspaces/issue-912-pictures-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
}
$WorkspacePath = [IO.Path]::GetFullPath($WorkspacePath)
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot 'tmp_workspaces')) + [IO.Path]::DirectorySeparatorChar
if (-not $WorkspacePath.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Workspace must be below repo/tmp_workspaces.' }
if (Test-Path -LiteralPath $WorkspacePath) { throw 'Use a fresh workspace; previous evidence is preserved.' }
[void][IO.Directory]::CreateDirectory($WorkspacePath)
if ($FixtureDirectory -and (Test-Path -LiteralPath $FixtureDirectory)) { throw 'Fixture directory must be new.' }
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class PictureGateWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
'@
$refs = [Collections.Generic.List[object]]::new()
function Hold-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { $refs.Add($Value) }
    return ,$Value
}
function Release-Children {
    for ($i = $refs.Count - 1; $i -ge 0; $i--) {
        if ([Runtime.InteropServices.Marshal]::IsComObject($refs[$i])) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($refs[$i]) }
    }
    $refs.Clear()
}
function Get-PictureObservation([object]$Book) {
    $raw = [string]$excel.Run("'$($Book.Name.Replace("'", "''"))'!Main.ObservePictures")
    $rows = @()
    foreach ($entry in $raw.Split(';')) {
        $parts = $entry.Split('|')
        if ($parts.Count -ne 6) { throw "Invalid picture observation: $raw" }
        if ([int]$parts[1] -ne 1 -or [int]$parts[2] -ne 508 -or [int]$parts[3] -ne 339) { throw "Picture does not match the 24x16 fixture dimensions: $entry" }
        if ([int]$parts[4] -ne 3 -or [int]$parts[5] -ne 3) { throw "Picture layout settings changed: $entry" }
        $rows += [pscustomobject]@{ name = $parts[0]; type = [int]$parts[1]; width = [int]$parts[2]; height = [int]$parts[3]; pictureSizeMode = [int]$parts[4]; pictureAlignment = [int]$parts[5] }
    }
    Release-Children
    return ,$rows
}
function Export-VbaProject([string]$Source, [string]$Target) {
    $archive = [IO.Compression.ZipFile]::OpenRead($Source)
    try {
        $entry = $archive.GetEntry('xl/vbaProject.bin')
        if ($null -eq $entry) { throw 'Missing VBA project.' }
        [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $Target, $false)
    } finally { $archive.Dispose() }
}
$excel = $null
$book = $null
$excelProcessId = [uint32]0
$success = $false
$cleanupConfirmed = $false
$environment = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    [void][PictureGateWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
    $environment = [ordered]@{ issue = 912; phase = $Phase; version = [string]$excel.Version; build = [string]$excel.Build; os = [Environment]::OSVersion.VersionString; pid = $excelProcessId; cleanupConfirmed = $false }
    if ($Phase -eq 'create') {
        $bitmap = [Drawing.Bitmap]::new(24, 16)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([Drawing.Color]::Navy)
            $graphics.FillRectangle([Drawing.Brushes]::Gold, 2, 3, 12, 8)
            $bitmap.Save((Join-Path $WorkspacePath 'logo.bmp'), [Drawing.Imaging.ImageFormat]::Bmp)
            $bitmap.Save((Join-Path $WorkspacePath 'logo.jpg'), [Drawing.Imaging.ImageFormat]::Jpeg)
        } finally { $graphics.Dispose(); $bitmap.Dispose() }
        $books = Hold-Com $excel.Workbooks
        $book = $books.Add()
        $project = Hold-Com $book.VBProject
        $components = Hold-Com $project.VBComponents
        $form = Hold-Com ($components.Add(3))
        $designer = Hold-Com $form.Designer
        $form.Name = 'PictureForm'
        $controls = Hold-Com $designer.Controls
        $bmp = Hold-Com ($controls.Add('Forms.Image.1', 'ImageBmp', $true))
        $bmp.Left = 6; $bmp.Top = 6; $bmp.Width = 48; $bmp.Height = 32
        $bmp.PictureSizeMode = 3; $bmp.PictureAlignment = 3
        $frame = Hold-Com ($controls.Add('Forms.Frame.1', 'FrameMain', $true))
        $frame.Left = 6; $frame.Top = 48; $frame.Width = 90; $frame.Height = 72
        $children = Hold-Com $frame.Controls
        $jpeg = Hold-Com ($children.Add('Forms.Image.1', 'ImageJpeg', $true))
        $jpeg.Left = 6; $jpeg.Top = 6; $jpeg.Width = 48; $jpeg.Height = 32
        $jpeg.PictureSizeMode = 3; $jpeg.PictureAlignment = 3
        $module = Hold-Com ($components.Add(1))
        $module.Name = 'Main'
        $code = Hold-Com $module.CodeModule
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $bmpPath = (Join-Path $WorkspacePath 'logo.bmp').Replace('"', '""')
        $jpegPath = (Join-Path $WorkspacePath 'logo.jpg').Replace('"', '""')
        $authorCode = @"
Option Explicit
Public Sub AuthorPictures()
    Dim d As Object
    Set d = ThisWorkbook.VBProject.VBComponents("PictureForm").Designer
    Set d.Controls("ImageBmp").Picture = LoadPicture("$bmpPath")
    Set d.Controls("FrameMain").Controls("ImageJpeg").Picture = LoadPicture("$jpegPath")
End Sub
Public Function ObservePictures() As String
    Dim f As Object, b As Object, j As Object
    Set f = VBA.UserForms.Add("PictureForm")
    Set b = f.Controls("ImageBmp")
    Set j = f.Controls("FrameMain").Controls("ImageJpeg")
    ObservePictures = b.Name & "|" & b.Picture.Type & "|" & b.Picture.Width & "|" & b.Picture.Height & "|" & b.PictureSizeMode & "|" & b.PictureAlignment
    ObservePictures = ObservePictures & ";" & j.Name & "|" & j.Picture.Type & "|" & j.Picture.Width & "|" & j.Picture.Height & "|" & j.PictureSizeMode & "|" & j.PictureAlignment
    Unload f
End Function
Public Sub RunPictureSentinel()
    Dim f As Object
    Set f = VBA.UserForms.Add("PictureForm")
    If f.Controls("ImageBmp").Picture.Type <> 1 Then Err.Raise 5
    If f.Controls("FrameMain").Controls("ImageJpeg").Picture.Type <> 1 Then Err.Raise 5
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "issue-912-ok"
    Unload f
End Sub
"@
        $code.AddFromString($authorCode)
        $WorkbookPath = Join-Path $WorkspacePath 'baseline.xlsm'
        $book.SaveAs($WorkbookPath, 52)
        Release-Children
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $book = $null
        $books = Hold-Com $excel.Workbooks
        $book = $books.Open($WorkbookPath)
        [void]$excel.Run("'$($book.Name.Replace("'", "''"))'!Main.AuthorPictures")
        $before = Get-PictureObservation $book
        $book.Save()
    } else {
        if (-not $WorkbookPath -or -not (Test-Path -LiteralPath $WorkbookPath)) { throw 'verify requires -WorkbookPath.' }
        $WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
        $books = Hold-Com $excel.Workbooks
        $book = $books.Open($WorkbookPath)
        $before = Get-PictureObservation $book
    }
    [void]$excel.Run("'$($book.Name.Replace("'", "''"))'!Main.RunPictureSentinel")
    $sheets = Hold-Com $book.Worksheets
    $sheet = Hold-Com ($sheets.Item(1))
    $cell = Hold-Com ($sheet.Range('A1'))
    if ([string]$cell.Value2 -cne 'issue-912-ok') { throw 'Sentinel differs.' }
    Release-Children
    $normalized = Join-Path $WorkspacePath 'normalized.xlsm'
    $book.SaveAs($normalized, 52)
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
    $books = Hold-Com $excel.Workbooks
    $book = $books.Open($normalized)
    $after = Get-PictureObservation $book
    if (($before | ConvertTo-Json -Depth 8 -Compress) -cne ($after | ConvertTo-Json -Depth 8 -Compress)) { throw 'Pictures changed after save/reopen.' }
    [void]$excel.Run("'$($book.Name.Replace("'", "''"))'!Main.RunPictureSentinel")
    $environment.beforeSave = $before
    $environment.reopened = $after
    $environment.sentinel = 'issue-912-ok'
    $success = $true
} finally {
    Release-Children
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excelProcessId) {
        for ($i = 0; $i -lt 30; $i++) {
            if (-not (Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue)) { $cleanupConfirmed = $true; break }
            Start-Sleep -Milliseconds 500
        }
    }
    if ($null -ne $environment) {
        $environment.cleanupConfirmed = $cleanupConfirmed
        [IO.File]::WriteAllText((Join-Path $WorkspacePath 'environment.json'), ($environment | ConvertTo-Json -Depth 12), $utf8)
    }
    if (-not $cleanupConfirmed) { throw "Excel cleanup not confirmed: PID=$excelProcessId. No fixture may be promoted." }
}
if ($success -and $Phase -eq 'create') {
    Export-VbaProject $WorkbookPath (Join-Path $WorkspacePath 'baseline.bin')
    Export-VbaProject (Join-Path $WorkspacePath 'normalized.xlsm') (Join-Path $WorkspacePath 'normalized.bin')
    if ($FixtureDirectory) {
        [void][IO.Directory]::CreateDirectory($FixtureDirectory)
        foreach ($name in @('baseline.bin', 'normalized.bin', 'logo.bmp', 'logo.jpg', 'environment.json')) {
            [IO.File]::Copy((Join-Path $WorkspacePath $name), (Join-Path $FixtureDirectory $name), $false)
        }
    }
}
Write-Output "picture gate passed: phase=$Phase workspace=$WorkspacePath workbook=$WorkbookPath cleanup=$cleanupConfirmed"
