param(
  [string]$WorkbookPath,
  [string]$Active = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "attach"

try {
  if (-not (ConvertTo-XlflowBool $Active)) {
    Set-XlflowError -Result $result -Code "attach_args_invalid" -Message "--active is required for attach in this version." -Source "xlflow"
    Write-XlflowJson -Result $result
    exit
  }

  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  if ($null -eq $excel -or $null -eq $excel.ActiveWorkbook) {
    Set-XlflowError -Result $result -Code "active_workbook_not_found" -Message "No active Excel workbook is available." -Source "Excel"
    Write-XlflowJson -Result $result
    exit
  }

  $activeWorkbook = $excel.ActiveWorkbook
  $activePath = $activeWorkbook.FullName
  $configuredPath = [System.IO.Path]::GetFullPath($WorkbookPath)
  $pathsMatch = ([string]::Equals($activePath, $configuredPath, [System.StringComparison]::OrdinalIgnoreCase))

  $result.workbook = [ordered]@{
    path = $activePath
    configured_path = $configuredPath
    active = $true
    matches_config = $pathsMatch
  }

  if (-not $pathsMatch) {
    Set-XlflowError -Result $result -Code "active_workbook_mismatch" -Message ("Active workbook does not match configured workbook: " + $activePath) -Source "Excel"
  } else {
    $result.logs = @("attached to active workbook " + $activePath)
  }
} catch {
  Set-XlflowError -Result $result -Code "active_workbook_not_found" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
}

Write-XlflowJson -Result $result
