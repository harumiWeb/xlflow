param(
  [string]$Action,
  [string]$WorkbookPath,
  [string]$Visible = "false"
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "runner"
$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($WorkbookPath)
  $project = $workbook.VBProject

  switch ($Action) {
    "install" {
      Install-XlflowRunnerModule -VBProject $project
      $workbook.Save()
      $result.runner = [ordered]@{ installed = $true; module = "XlflowRunner" }
      $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true }
      $result.logs = @("installed XlflowRunner")
    }
    "remove" {
      $removed = Remove-XlflowRunnerModule -VBProject $project
      if ($removed) {
        $workbook.Save()
      }
      $result.runner = [ordered]@{ installed = $false; removed = $removed; module = "XlflowRunner" }
      $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $removed }
      $result.logs = @($(if ($removed) { "removed XlflowRunner" } else { "XlflowRunner was not installed" }))
    }
    "status" {
      $installed = Test-XlflowRunnerModuleInstalled -VBProject $project
      $result.runner = [ordered]@{ installed = $installed; module = "XlflowRunner" }
      $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false }
      $result.logs = @($(if ($installed) { "XlflowRunner is installed" } else { "XlflowRunner is not installed" }))
    }
    default {
      Set-XlflowError -Result $result -Code "runner_args_invalid" -Message "-Action must be install, remove, or status." -Source "xlflow"
    }
  }
} catch {
  if ($null -eq $result.error) {
    Set-XlflowError -Result $result -Code "runner_failed" -Message $_.Exception.Message -Source $_.Exception.Source
  }
} finally {
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
