param(
  [string]$WorkbookPath,
  [string]$ModulesDir,
  [string]$ClassesDir,
  [string]$FormsDir,
  [string]$WorkbookDir,
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

try {
  New-Item -ItemType Directory -Force -Path $ModulesDir, $ClassesDir, $FormsDir, $WorkbookDir | Out-Null
  $openResult = Open-XlflowWorkbookForCommand -WorkbookPath $WorkbookPath -Visible $Visible -DisplayAlerts "false" -DisableAutomationMacros "true" -UseSession $UseSession -MetadataPath $MetadataPath
  $excel = $openResult.excel
  $workbook = $openResult.workbook
  $sessionAttached = [bool]$openResult.session_attached
  $sessionMode = [string]$openResult.session_mode

  Clear-XlflowSourceComponentFiles -ModulesDir $ModulesDir -ClassesDir $ClassesDir -FormsDir $FormsDir -WorkbookDir $WorkbookDir

  $exported = @()
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
  }

  $result.workbook = New-XlflowWorkbookResult -WorkbookPath $WorkbookPath -SessionAttached $sessionAttached -SessionMode $sessionMode
  $result.logs = @(@($(Get-XlflowSessionUsageLog -SessionMode $sessionMode), "exported $($exported.Count) VBA component(s)") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
} catch {
  Set-XlflowError -Result $result -Code "excel_export_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
} finally {
  if ($sessionAttached) {
    Release-XlflowComReferences -Workbook $workbook -Excel $excel
  } else {
    Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
  }
}

Write-XlflowJson -Result $result
