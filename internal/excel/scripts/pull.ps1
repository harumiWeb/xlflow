param(
  [string]$WorkbookPath,
  [string]$ModulesDir,
  [string]$ClassesDir,
  [string]$FormsDir,
  [string]$WorkbookDir,
  [string]$CodeSource = "frm",
  [string]$Folders = "true",
  [string]$FolderAnnotation = "update",
  [string]$DefaultComponentFolders = "true",
  [string]$Visible = "false",
  [string]$UseSession = "false",
  [string]$MetadataPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "pull"
$excel = $null
$workbook = $null
$sessionAttached = $false
$sessionMode = "none"
$saveState = [ordered]@{ dirty = $false; needs_save = $false }

try {
  New-Item -ItemType Directory -Force -Path $ModulesDir, $ClassesDir, $FormsDir, $WorkbookDir | Out-Null
  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode
  $saveState = Get-XlflowWorkbookSaveState -Workbook $workbook -SessionAttached $sessionAttached
  $userFormNames = @()
  try {
    $userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)
  } catch {
    Write-Verbose ("failed to inspect UserForms during pull: " + $_.Exception.Message)
  }

  Clear-XlflowSourceComponentFiles -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir -CodeSource $CodeSource

  $exported = @()
  $exportedFormCode = @()
  foreach ($component in $workbook.VBProject.VBComponents) {
    $path = Get-XlflowComponentPath -Component $component -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir -Folders $Folders -FolderAnnotation $FolderAnnotation -DefaultComponentFolders $DefaultComponentFolders
    if ($null -ne $path) {
      $parent = Split-Path -Parent $path
      if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
      }
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      $component.Export($path)
      Convert-XlflowExportedSourceToUtf8 -Path $path
      if ($component.Type -eq 100) {
        Normalize-XlflowDocumentModuleFile -Path $path -RootDir $WorkbookDir -FolderAnnotationMode $FolderAnnotation
      } elseif ($FolderAnnotation -eq "update") {
        $rootDir = Get-XlflowComponentRootDir -ComponentType $component.Type -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir
        $desiredAnnotation = Get-XlflowFolderAnnotationForPath -RootDir $rootDir -Path $path
        $content = Get-XlflowUtf8Text -Path $path
        $content = Update-XlflowFolderAnnotationText -Text $content -FolderAnnotationMode $FolderAnnotation -DesiredAnnotation $desiredAnnotation
        Set-XlflowUtf8Text -Path $path -Text $content
      }
      $exported += $path
    }
    if ((Use-XlflowUserFormCodeSidecar -CodeSource $CodeSource) -and $component.Type -eq 3) {
      $codePath = Export-XlflowUserFormCodeBehind -Component $component -FormsDir $FormsDir
      if (-not [string]::IsNullOrWhiteSpace($codePath)) {
        $exportedFormCode += $codePath
      }
    }
  }

  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode -Dirty $saveState.dirty -NeedsSave $saveState.needs_save
  $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
  Add-XlflowUserFormDiscoveryMessages -Result $result -Names $userFormNames
  if ($saveState.needs_save) {
    Add-XlflowStateWarning -Result $result -Code "save_required" -Message "The live workbook is newer than disk. `pull` exported from the live workbook rather than the saved workbook file."
  }
  $result.logs = @(@(
      $(Get-XlflowSessionUsageLog -SessionMode $sessionMode),
      "exported $($exported.Count) VBA component(s)",
      "exported $($exportedFormCode.Count) UserForm code-behind sidecar(s)"
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  Set-XlflowError -Result $result -Code "excel_export_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  $result.target = New-XlflowTargetResult -Kind $(if ($sessionAttached) { "live_session" } else { "file" }) -Path $WorkbookPath
  $result.session = New-XlflowSessionResult -Active $sessionAttached -WorkbookPath $WorkbookPath -Dirty $saveState.dirty -SaveRequired $saveState.needs_save -Mode $sessionMode
} finally {
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
