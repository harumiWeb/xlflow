[CmdletBinding()]
param(
    [switch]$KeepWorkspace,
    [string]$WorkspaceSuffix = ''
)

$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = Join-Path $repoRoot 'tmp_workspaces'
$templateWorkspace = Join-Path $workspaceRoot ("pack-stable-e2e$WorkspaceSuffix")
$blankWorkspace = Join-Path $workspaceRoot ("pack-stable-blank-e2e$WorkspaceSuffix")
$utf8NoBom = [Text.UTF8Encoding]::new($false)
$japaneseSheetName = -join [char[]](0x8868, 0x793A, 0x540D)
$japaneseModuleName = (-join [char[]](0x65E5, 0x672C, 0x8A9E)) + 'Module'
$japaneseAddedMarker = -join [char[]](0x6A19, 0x6E96, 0x8FFD, 0x52A0)
$japaneseModuleMarker = -join [char[]](0x65E5, 0x672C, 0x8A9E, 0x30E2, 0x30B8, 0x30E5, 0x30FC, 0x30EB)
$japaneseFormCaption = -join [char[]](0x5B89, 0x5B9A, 0x7248, 0x30D5, 0x30A9, 0x30FC, 0x30E0)
$japanesePageCaption = (-join [char[]](0x30DA, 0x30FC, 0x30B8)) + '1'

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

function Get-ReferenceSnapshot {
    param([Parameter(Mandatory)][object]$Project)

    $result = @()
    $references = $Project.References
    try {
        for ($index = 1; $index -le $references.Count; $index++) {
            $reference = $references.Item($index)
            try {
                $result += "$($reference.Name)|$($reference.Guid)|$($reference.Major)|$($reference.Minor)"
            } finally {
                Release-ComObject $reference
            }
        }
    } finally {
        Release-ComObject $references
    }
    return @($result | Sort-Object)
}

function Add-CodeComponent {
    param(
        [Parameter(Mandatory)][object]$Project,
        [Parameter(Mandatory)][int]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Code
    )

    $component = $Project.VBComponents.Add($Type)
    try {
        $component.Name = $Name
        $component.CodeModule.AddFromString($Code)
    } finally {
        Release-ComObject $component
    }
}

function Get-VBComponentByName {
    param([Parameter(Mandatory)][object]$Project, [Parameter(Mandatory)][string]$Name)

    for ($index = 1; $index -le $Project.VBComponents.Count; $index++) {
        $component = $Project.VBComponents.Item($index)
        if ($component.Name -eq $Name) { return $component }
        Release-ComObject $component
    }
    throw "VBA component not found: $Name"
}

function New-PackTemplate {
    param([Parameter(Mandatory)][string]$WorkbookPath)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $workbook = $null
    $project = $null
    try {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        Write-Verbose 'Configuring VBA project metadata'
        $project = $workbook.VBProject

        Write-Verbose 'Adding second worksheet and assigning DataSheet CodeName'
        $sheet = $workbook.Worksheets.Add()
        try {
            $sheet.Name = $japaneseSheetName
            $workbook.Save()
            Release-ComObject $project
            $project = $workbook.VBProject
            $sheetComponent = $null
            for ($index = 1; $index -le $project.VBComponents.Count; $index++) {
                $candidate = $project.VBComponents.Item($index)
                if ($candidate.Type -eq 100 -and $candidate.Name -notin @('ThisWorkbook', 'Sheet1')) {
                    $sheetComponent = $candidate
                    break
                }
                Release-ComObject $candidate
            }
            if ($null -eq $sheetComponent) { throw 'new worksheet document component was not found' }
            try { $sheetComponent.Name = 'DataSheet' } finally { Release-ComObject $sheetComponent }
        } finally {
            Release-ComObject $sheet
        }

        Write-Verbose 'Adding standard and class components'
        Add-CodeComponent $project 1 'UpdateStandard' "Option Explicit`r`nPublic Function Marker() As String`r`n  Marker = `"standard-before`"`r`nEnd Function`r`n"
        Add-CodeComponent $project 1 'RemoveStandard' "Option Explicit`r`n"
        Add-CodeComponent $project 1 'RenameStandard' "Option Explicit`r`n"
        Add-CodeComponent $project 2 'UpdateClass' "Option Explicit`r`nPublic Function Marker() As String`r`n  Marker = `"class-before`"`r`nEnd Function`r`n"
        Add-CodeComponent $project 2 'RemoveClass' "Option Explicit`r`n"
        Add-CodeComponent $project 2 'RenameClass' "Option Explicit`r`n"

        Write-Verbose 'Adding Scripting Runtime reference'
        try {
            [void]$project.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}', 1, 0)
        } catch {
            if ($_.Exception.Message -notmatch 'already in use|already referenced') { throw }
        }

        Write-Verbose 'Adding nested UserForm designer fixture'
        $form = $project.VBComponents.Add(3)
        try {
            $form.Name = 'ReleaseForm'
            $form.CodeModule.AddFromString("Option Explicit`r`nPublic Function PackMarker() As String`r`n  PackMarker = `"form-before`"`r`nEnd Function`r`n")
            $designer = $form.Designer
            try {
                $designer.Caption = $japaneseFormCaption
                $frame = $designer.Controls.Add('Forms.Frame.1', 'FrameMain', $true)
                try {
                    $frame.Caption = 'frame-preserved'
                    $nestedLabel = $frame.Controls.Add('Forms.Label.1', 'NestedLabel', $true)
                    try { $nestedLabel.Caption = 'nested-preserved' } finally { Release-ComObject $nestedLabel }
                } finally { Release-ComObject $frame }
                $pages = $designer.Controls.Add('Forms.MultiPage.1', 'PagesMain', $true)
                try {
                    $page = $pages.Pages.Item(0)
                    try {
                        $page.Caption = $japanesePageCaption
                        $pageLabel = $page.Controls.Add('Forms.Label.1', 'PageLabel', $true)
                        try { $pageLabel.Caption = 'page-preserved' } finally { Release-ComObject $pageLabel }
                    } finally { Release-ComObject $page }
                } finally { Release-ComObject $pages }
            } finally { Release-ComObject $designer }
        } finally { Release-ComObject $form }

        Write-Verbose 'Saving Excel-created template'
        $workbook.Save()
        return @{
            ProjectName = $project.Name
            References = @(Get-ReferenceSnapshot $project)
            ExcelVersion = $excel.Version
            ExcelOperatingSystem = $excel.OperatingSystem
        }
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        $excel.Quit()
        Release-ComObject $project
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Write-ClassSource {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Marker)

    Write-Utf8NoBom $Path @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "$Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Marker() As String
    Marker = "$Marker"
End Function
"@
}

function Write-PackSources {
    param([Parameter(Mandatory)][string]$Path)

    Get-ChildItem -LiteralPath (Join-Path $Path 'src\modules') -Filter '*.bas' -File | Remove-Item -Force
    Remove-Item -LiteralPath (Join-Path $Path 'src\classes\RemoveClass.cls'), (Join-Path $Path 'src\classes\RenameClass.cls') -Force

    Write-Utf8NoBom (Join-Path $Path 'src\modules\UpdateStandard.bas') @'
Attribute VB_Name = "UpdateStandard"
Option Explicit
Public Function Marker() As String
    Marker = "standard-updated"
End Function
'@
    Write-Utf8NoBom (Join-Path $Path 'src\modules\RenamedStandard.bas') @'
Attribute VB_Name = "RenamedStandard"
Option Explicit
Public Function Marker() As String
    Marker = "standard-renamed"
End Function
'@
    Write-Utf8NoBom (Join-Path $Path 'src\modules\AddedStandard.bas') @"
Attribute VB_Name = "AddedStandard"
Option Explicit
Public Function Marker() As String
    ' Non-ASCII comment and string literal
    Marker = "$japaneseAddedMarker"
End Function
"@
    Write-Utf8NoBom (Join-Path $Path "src\modules\$japaneseModuleName.bas") @"
Attribute VB_Name = "$japaneseModuleName"
Option Explicit
Public Function Marker() As String
    Marker = "$japaneseModuleMarker"
End Function
"@

    Write-ClassSource (Join-Path $Path 'src\classes\UpdateClass.cls') 'UpdateClass' 'class-updated'
    Write-ClassSource (Join-Path $Path 'src\classes\RenamedClass.cls') 'RenamedClass' 'class-renamed'
    Write-ClassSource (Join-Path $Path 'src\classes\AddedClass.cls') 'AddedClass' 'class-added'

    Write-Utf8NoBom (Join-Path $Path 'src\workbook\ThisWorkbook.bas') @'
Option Explicit
Public Function WorkbookMarker() As String
    WorkbookMarker = "workbook-updated"
End Function
'@
    Write-Utf8NoBom (Join-Path $Path 'src\workbook\Sheet1.bas') @'
Option Explicit
Public Function SheetMarker() As String
    SheetMarker = "sheet1-updated"
End Function
'@
    Write-Utf8NoBom (Join-Path $Path 'src\workbook\DataSheet.bas') @'
Option Explicit
Public Function SheetMarker() As String
    SheetMarker = "datasheet-updated"
End Function
'@
    Write-Utf8NoBom (Join-Path $Path 'src\forms\code\ReleaseForm.bas') @'
Option Explicit
Public Function PackMarker() As String
    PackMarker = Me.FrameMain.Caption & "|" & Me.PagesMain.Pages(0).Controls("PageLabel").Caption
End Function
'@
    Write-Utf8NoBom (Join-Path $Path 'src\modules\Main.bas') @"
Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim updated As New UpdateClass
    Dim renamed As New RenamedClass
    Dim added As New AddedClass
    Dim form As Object
    Set form = VBA.UserForms.Add("ReleaseForm")
    If UpdateStandard.Marker() <> "standard-updated" Then Err.Raise 5
    If RenamedStandard.Marker() <> "standard-renamed" Then Err.Raise 5
    If AddedStandard.Marker() <> "$japaneseAddedMarker" Then Err.Raise 5
    If $japaneseModuleName.Marker() <> "$japaneseModuleMarker" Then Err.Raise 5
    If updated.Marker() <> "class-updated" Then Err.Raise 5
    If renamed.Marker() <> "class-renamed" Then Err.Raise 5
    If added.Marker() <> "class-added" Then Err.Raise 5
    If ThisWorkbook.WorkbookMarker() <> "workbook-updated" Then Err.Raise 5
    If Sheet1.SheetMarker() <> "sheet1-updated" Then Err.Raise 5
    If DataSheet.SheetMarker() <> "datasheet-updated" Then Err.Raise 5
    If form.PackMarker() <> "frame-preserved|page-preserved" Then Err.Raise 5
    Unload form
    ThisWorkbook.Worksheets(1).Range("A1").Value = "pack stable ok"
End Sub
"@
}

function Assert-PackedWorkbook {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$TemplateMetadata,
        [Parameter(Mandatory)][string[]]$RequiredComponents,
        [string[]]$ForbiddenComponents = @()
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $workbook = $null
    $project = $null
    try {
        $workbook = $excel.Workbooks.Open($Path)
        $project = $workbook.VBProject
        $components = @()
        for ($index = 1; $index -le $project.VBComponents.Count; $index++) {
            $component = $project.VBComponents.Item($index)
            try { $components += $component.Name } finally { Release-ComObject $component }
        }
        foreach ($name in $RequiredComponents) {
            if ($components -notcontains $name) { throw "packed workbook is missing component $name; got $($components -join ', ')" }
        }
        foreach ($name in $ForbiddenComponents) {
            if ($components -contains $name) { throw "packed workbook retained removed component $name" }
        }
        $dataSheet = $workbook.Worksheets.Item($japaneseSheetName)
        try {
            if ($dataSheet.CodeName -ne 'DataSheet') { throw "worksheet CodeName was $($dataSheet.CodeName), expected DataSheet" }
        } finally { Release-ComObject $dataSheet }
        if ($project.Name -ne $TemplateMetadata.ProjectName) {
            throw 'packed workbook did not preserve the VBA project name'
        }
        $references = @(Get-ReferenceSnapshot $project)
        if (Compare-Object $TemplateMetadata.References $references) {
            throw 'packed workbook did not preserve the VBA reference set'
        }
        $form = Get-VBComponentByName $project 'ReleaseForm'
        try {
            $designer = $form.Designer
            try {
                $frame = $designer.Controls.Item('FrameMain')
                $pages = $designer.Controls.Item('PagesMain')
                try {
                    if ($frame.Caption -ne 'frame-preserved' -or $frame.Controls.Item('NestedLabel').Caption -ne 'nested-preserved') {
                        throw 'packed UserForm did not preserve nested Frame state'
                    }
                    if ($pages.Pages.Item(0).Controls.Item('PageLabel').Caption -ne 'page-preserved') {
                        throw 'packed UserForm did not preserve MultiPage state'
                    }
                } finally {
                    Release-ComObject $pages
                    Release-ComObject $frame
                }
            } finally { Release-ComObject $designer }
        } finally { Release-ComObject $form }
        [void]$excel.Run("'$($workbook.Name)'!Main.Run")
        $sentinel = $workbook.Worksheets.Item(1).Range('A1').Value2
        if ($sentinel -ne 'pack stable ok') { throw "pack sentinel was '$sentinel'" }
        return $sentinel
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        $excel.Quit()
        Release-ComObject $project
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Test-TemplatePack {
    New-Item -ItemType Directory -Path $templateWorkspace | Out-Null
    Push-Location $templateWorkspace
    try {
        Invoke-XlflowJson @('new', 'StablePack.xlsm', '--json') | Out-Null
        $workbookPath = Join-Path $templateWorkspace 'build\StablePack.xlsm'
        $metadata = New-PackTemplate $workbookPath
        Invoke-XlflowJson @('pull', '--json') | Out-Null
        if (-not (Test-Path -LiteralPath (Join-Path $templateWorkspace 'src\forms\ReleaseForm.frx'))) {
            throw 'Excel did not export the expected ReleaseForm.frx resource'
        }
        Write-PackSources $templateWorkspace
        $result = Invoke-XlflowJson @('pack', '--out', 'dist/StablePack.xlsm', '--json')
        Write-Utf8NoBom (Join-Path $templateWorkspace 'dist\pack-result.json') $result.Raw
        if ($result.Json.pack.backend -ne 'pure-go' -or $result.Json.pack.vbe_validation -ne 'not_performed') {
            throw "stable pack JSON lost backend/validation contract: $($result.Raw)"
        }
        if ($result.Json.pack.PSObject.Properties.Name -contains 'experimental') {
            throw "stable pack JSON retained experimental marker: $($result.Raw)"
        }
        $required = @('ThisWorkbook', 'Sheet1', 'DataSheet', 'Main', 'UpdateStandard', 'RenamedStandard', 'AddedStandard', $japaneseModuleName, 'UpdateClass', 'RenamedClass', 'AddedClass', 'ReleaseForm')
        $forbidden = @('RemoveStandard', 'RenameStandard', 'RemoveClass', 'RenameClass')
        $sentinel = Assert-PackedWorkbook (Join-Path $templateWorkspace 'dist\StablePack.xlsm') $metadata $required $forbidden

        Write-Utf8NoBom (Join-Path $templateWorkspace 'src\forms\NewForm.frm') @'
VERSION 5.00
Begin VB.UserForm NewForm
   Caption = "NewForm"
End
Attribute VB_Name = "NewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@
        $rejected = Invoke-XlflowJson @('pack', '--out', 'dist/Rejected.xlsm', '--json') -AllowFailure
        Write-Utf8NoBom (Join-Path $templateWorkspace 'dist\new-userform-rejection.json') $rejected.Raw
        if ($rejected.ExitCode -ne 1 -or $rejected.Json.error.code -ne 'pack_userform_generation_unsupported') {
            throw "new UserForm rejection contract failed: $($rejected.Raw)$($rejected.Stderr)"
        }
        if (Test-Path -LiteralPath (Join-Path $templateWorkspace 'dist\Rejected.xlsm')) {
            throw 'rejected pack published an artifact'
        }
        Write-Output "pack template E2E passed: workspace=$templateWorkspace sentinel=$sentinel Excel=$($metadata.ExcelVersion) OS=$($metadata.ExcelOperatingSystem)"
    } finally {
        Pop-Location
    }
}

function Test-BlankPack {
    New-Item -ItemType Directory -Path $blankWorkspace | Out-Null
    Push-Location $blankWorkspace
    try {
        Invoke-XlflowJson @('new', 'BlankSource.xlsm', '--json') | Out-Null
        Write-Utf8NoBom (Join-Path $blankWorkspace 'src\modules\Main.bas') @'
Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    ThisWorkbook.Worksheets(1).Range("A1").Value = "pack blank ok"
End Sub
'@
        $result = Invoke-XlflowJson @('pack', '--blank', '--out', 'dist/Blank.xlsm', '--json')
        Write-Utf8NoBom (Join-Path $blankWorkspace 'dist\pack-result.json') $result.Raw
        if ($result.Json.pack.base -ne 'blank' -or $result.Json.pack.backend -ne 'pure-go' -or $result.Json.pack.vbe_validation -ne 'not_performed') {
            throw "blank pack JSON contract failed: $($result.Raw)"
        }
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1
        $workbook = $null
        try {
            $workbook = $excel.Workbooks.Open((Join-Path $blankWorkspace 'dist\Blank.xlsm'))
            [void]$excel.Run("'$($workbook.Name)'!Main.Run")
            $sentinel = $workbook.Worksheets.Item(1).Range('A1').Value2
            if ($sentinel -ne 'pack blank ok') { throw "blank pack sentinel was '$sentinel'" }
            Write-Output "pack blank E2E passed: workspace=$blankWorkspace sentinel=$sentinel"
        } finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
            $excel.Quit()
            Release-ComObject $workbook
            Release-ComObject $excel
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    } finally {
        Pop-Location
    }
}

if (-not (Get-Command xlflow -ErrorAction SilentlyContinue)) {
    throw 'xlflow was not found on PATH. Run task install before scripts/test-pack-e2e.ps1.'
}

New-Item -ItemType Directory -Force -Path $workspaceRoot | Out-Null
foreach ($workspace in @($templateWorkspace, $blankWorkspace)) {
    if (Test-Path -LiteralPath $workspace) { Remove-Item -LiteralPath $workspace -Recurse -Force }
}
try {
    Test-TemplatePack
    Test-BlankPack
} finally {
    if (-not $KeepWorkspace) {
        foreach ($workspace in @($templateWorkspace, $blankWorkspace)) {
            if (Test-Path -LiteralPath $workspace) { Remove-Item -LiteralPath $workspace -Recurse -Force }
        }
    }
}
