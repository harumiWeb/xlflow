package scripts_test

import (
	"encoding/json"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func TestPowerShellScriptsParse(t *testing.T) {
	scripts := []string{"common.ps1", "doctor.ps1", "new.ps1", "pull.ps1", "push.ps1", "run.ps1", "test.ps1"}
	for _, script := range scripts {
		script := script
		t.Run(script, func(t *testing.T) {
			path := filepath.Join(".", script)
			cmd := exec.Command("pwsh", "-NoProfile", "-Command", "try { [scriptblock]::Create((Get-Content -Raw -LiteralPath '"+path+"')) | Out-Null } catch { Write-Error $_; exit 1 }")
			out, err := cmd.CombinedOutput()
			if err != nil {
				t.Fatalf("script parse failed: %v\n%s", err, out)
			}
		})
	}
}

func TestTestProcedureDiscoveryRules(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $body = @('Option Explicit','Public Sub TestCreateReport()','End Sub','Sub Totals_Test()','End Sub','Private Sub TestPrivate()','End Sub','Public Sub TestWithArg(value As Variant)','End Sub','Public Sub Helper()','End Sub') -join [Environment]::NewLine; Find-XlflowTestProcedures -ModuleName 'ReportTests' -Code $body | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("test discovery failed: %v\n%s", err, out)
	}
	var got []struct {
		Name   string `json:"name"`
		Module string `json:"module"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse discovery output: %v\n%s", err, out)
	}
	if len(got) != 2 {
		t.Fatalf("expected 2 discovered tests, got %d: %+v", len(got), got)
	}
	if got[0].Name != "TestCreateReport" || got[0].Module != "ReportTests" {
		t.Fatalf("unexpected first test: %+v", got[0])
	}
	if got[1].Name != "Totals_Test" || got[1].Module != "ReportTests" {
		t.Fatalf("unexpected second test: %+v", got[1])
	}
}

func TestTestProcedureFilterUsesExactMatch(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $tests = @([ordered]@{ name = 'TestCreateReport'; module = 'ReportTests' }, [ordered]@{ name = 'TestCreateReportSlow'; module = 'ReportTests' }); $selected = @(Select-XlflowTests -Tests $tests -Filter 'TestCreateReport'); ConvertTo-Json -InputObject $selected -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("test filter failed: %v\n%s", err, out)
	}
	var got []struct {
		Name string `json:"name"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse filter output: %v\n%s", err, out)
	}
	if len(got) != 1 || got[0].Name != "TestCreateReport" {
		t.Fatalf("expected exact filter match only, got %+v", got)
	}
}

func TestTestRunnerCodeCatchesVBAErrors(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $tests = @([pscustomobject]@{ module = 'ReportTests'; name = 'TestFailure' }); New-XlflowTestRunnerCode -Tests $tests",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("runner code generation failed: %v\n%s", err, out)
	}
	got := string(out)
	for _, want := range []string{"Public Function RunTest", "On Error Resume Next", "Case 0", "ReportTests.TestFailure", "Err.Number", "Err.Description"} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected runner code to contain %q:\n%s", want, got)
		}
	}
}

func TestSetXlflowErrorMutatesResultEnvelope(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $result = New-XlflowResult -Command 'test'; Set-XlflowError -Result $result -Code 'test_failed' -Message 'boom'; Write-XlflowJson -Result $result",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Set-XlflowError failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse envelope: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "test_failed" || got.Error.Message != "boom" {
		t.Fatalf("expected failed envelope, got %+v", got)
	}
}

func TestDocumentModuleContentNormalization(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; Set-Content -LiteralPath $tmp -Value @('Attribute VB_Name = \"ThisWorkbook\"','Attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"','Option Explicit','Private Sub Workbook_Open()','End Sub'); . ./common.ps1; $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("normalization failed: %v\n%s", err, out)
	}
	got := string(out)
	if got == "" {
		t.Fatal("expected normalized content")
	}
	if strings.Contains(got, "Attribute VB_") {
		t.Fatalf("attribute lines were not removed: %q", got)
	}
	for _, marker := range []string{"VERSION 1.0 CLASS", "BEGIN", "MultiUse = -1", "END"} {
		if strings.Contains(got, marker) {
			t.Fatalf("class header lines were not removed: %q", got)
		}
	}
	if !strings.Contains(got, "Option Explicit") || !strings.Contains(got, "Workbook_Open") {
		t.Fatalf("expected VBA body to remain: %q", got)
	}
}

func TestDocumentModuleContentAddsOptionExplicitForEmptyDocumentModule(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; Set-Content -LiteralPath $tmp -Value @('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','END','Attribute VB_Name = \"Sheet1\"','Attribute VB_GlobalNameSpace = False','Attribute VB_Creatable = False','Attribute VB_PredeclaredId = True','Attribute VB_Exposed = True'); . ./common.ps1; $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("normalization failed: %v\n%s", err, out)
	}
	got := string(out)
	if !strings.Contains(got, "Option Explicit") {
		t.Fatalf("expected Option Explicit to be added for empty document module: %q", got)
	}
	for _, marker := range []string{"VERSION 1.0 CLASS", "BEGIN", "MultiUse = -1", "END"} {
		if strings.Contains(got, marker) {
			t.Fatalf("expected empty document module normalization to drop class header lines: %q", got)
		}
	}
}

func TestGetXlflowComponentPathMapsClassUserFormAndDocumentModules(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $modules = 'modules'; $classes = 'classes'; $forms = 'forms'; $workbook = 'workbook'; $class = [pscustomobject]@{ Name = 'Class1'; Type = 2 }; $form = [pscustomobject]@{ Name = 'UserForm1'; Type = 3 }; $document = [pscustomobject]@{ Name = 'ThisWorkbook'; Type = 100 }; Get-XlflowComponentPath -Component $class -ModulesDir $modules -ClassesDir $classes -FormsDir $forms -WorkbookDir $workbook; Get-XlflowComponentPath -Component $form -ModulesDir $modules -ClassesDir $classes -FormsDir $forms -WorkbookDir $workbook; Get-XlflowComponentPath -Component $document -ModulesDir $modules -ClassesDir $classes -FormsDir $forms -WorkbookDir $workbook",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("component path mapping failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	lines := strings.FieldsFunc(got, func(r rune) bool {
		return r == '\r' || r == '\n'
	})
	if len(lines) != 3 {
		t.Fatalf("expected 3 mapped paths, got %d: %q", len(lines), got)
	}
	normalizedClassPath := strings.ReplaceAll(lines[0], "\\", "/")
	normalizedFormPath := strings.ReplaceAll(lines[1], "\\", "/")
	normalizedDocumentPath := strings.ReplaceAll(lines[2], "\\", "/")
	if !strings.HasSuffix(normalizedClassPath, "classes/Class1.cls") {
		t.Fatalf("expected class module path to end with classes\\Class1.cls: %q", lines[0])
	}
	if !strings.HasSuffix(normalizedFormPath, "forms/UserForm1.frm") {
		t.Fatalf("expected userform path to end with forms\\UserForm1.frm: %q", lines[1])
	}
	if !strings.HasSuffix(normalizedDocumentPath, "workbook/ThisWorkbook.bas") {
		t.Fatalf("expected document module path to end with workbook\\ThisWorkbook.bas: %q", lines[2])
	}
}

func TestDocumentModuleContentKeepsExecutableEndStatement(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; Set-Content -LiteralPath $tmp -Value @('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','END','Attribute VB_Name = \"ThisWorkbook\"','Option Explicit','Public Sub StopAll()','  End','End Sub'); . ./common.ps1; $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("normalization failed: %v\n%s", err, out)
	}
	got := strings.ReplaceAll(string(out), "\r\n", "\n")
	if !strings.Contains(got, "\n  End\n") {
		t.Fatalf("expected executable End statement to remain in normalized document module: %q", got)
	}
	for _, marker := range []string{"VERSION 1.0 CLASS", "BEGIN", "MultiUse = -1"} {
		if strings.Contains(got, marker) {
			t.Fatalf("expected class header lines to be removed: %q", got)
		}
	}
}

func TestDocumentModuleContentDropsAdditionalClassHeaderProperties(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; Set-Content -LiteralPath $tmp -Value @('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','  Persistable = 0  ''NotPersistable','END','Attribute VB_Name = \"ThisWorkbook\"','Option Explicit','Public Sub Hello()','End Sub'); . ./common.ps1; $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("normalization failed: %v\n%s", err, out)
	}
	got := string(out)
	if strings.Contains(got, "Persistable = 0") {
		t.Fatalf("expected class header property lines to be removed: %q", got)
	}
	if !strings.Contains(got, "Public Sub Hello()") {
		t.Fatalf("expected executable VBA body to remain: %q", got)
	}
}

func TestDocumentModuleContentDoesNotTruncateBodyWhenHeaderEndMissing(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; Set-Content -LiteralPath $tmp -Value @('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','Option Explicit','Public Sub Recover()','End Sub'); . ./common.ps1; $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("normalization failed: %v\n%s", err, out)
	}
	got := string(out)
	if !strings.Contains(got, "Option Explicit") || !strings.Contains(got, "Public Sub Recover()") || !strings.Contains(got, "End Sub") {
		t.Fatalf("expected malformed header input to keep executable VBA body text: %q", got)
	}
}

func TestNormalizeDocumentModuleFileRewritesExportedWorkbookModule(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; Set-Content -LiteralPath $tmp -Value @('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','END','Attribute VB_Name = \"Sheet1\"','Attribute VB_GlobalNameSpace = False','Attribute VB_Creatable = False','Attribute VB_PredeclaredId = True','Attribute VB_Exposed = True'); . ./common.ps1; Normalize-XlflowDocumentModuleFile -Path $tmp; Get-Content -Raw -LiteralPath $tmp; Remove-Item -LiteralPath $tmp -Force",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("normalization failed: %v\n%s", err, out)
	}
	got := string(out)
	if !strings.Contains(got, "Option Explicit") {
		t.Fatalf("expected rewritten workbook module to include Option Explicit: %q", got)
	}
	if strings.Contains(got, "Attribute VB_") {
		t.Fatalf("expected rewritten workbook module to drop attribute lines: %q", got)
	}
	for _, marker := range []string{"VERSION 1.0 CLASS", "BEGIN", "MultiUse = -1", "END"} {
		if strings.Contains(got, marker) {
			t.Fatalf("expected rewritten workbook module to drop class header lines: %q", got)
		}
	}
}

func TestUserFormFrxCompanionSiblingPathAndReferenceArePreserved(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
$result = [ordered]@{
  skip = $false
  skipReason = ''
  frmPath = ''
  frxPath = ''
  frxExists = $false
  frmReferencesFrx = $false
  frxIsSibling = $false
  initialPullStatus = ''
  pushStatus = ''
  roundtripPullStatus = ''
}
$excel = $null
$workbook = $null
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $wbPath = Join-Path $root 'userform-frx-regression.xlsm'
  $modulesDir1 = Join-Path $root 'src1/modules'
  $classesDir1 = Join-Path $root 'src1/classes'
  $formsDir1 = Join-Path $root 'src1/forms'
  $workbookDir1 = Join-Path $root 'src1/workbook'
  $modulesDir2 = Join-Path $root 'src2/modules'
  $classesDir2 = Join-Path $root 'src2/classes'
  $formsDir2 = Join-Path $root 'src2/forms'
  $workbookDir2 = Join-Path $root 'src2/workbook'
  $backupRoot = Join-Path $root 'backups'
  $formName = 'UserFormFrxRegression'

  try {
    $excel = New-Object -ComObject Excel.Application
  } catch {
    $result.skip = $true
    $result.skipReason = 'Excel COM is unavailable: ' + $_.Exception.Message
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Add()

  try {
    $null = $workbook.VBProject
  } catch {
    $result.skip = $true
    $result.skipReason = 'VBProject access is unavailable (trust access to VBA project object model may be disabled): ' + $_.Exception.Message
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $component = $workbook.VBProject.VBComponents.Add(3)
  $component.Name = $formName
  $designer = $component.Designer
  $label = $designer.Controls.Add('Forms.Label.1')
  $label.Caption = 'frx-regression'
  $workbook.SaveAs($wbPath, 52)

  $workbook.Close($true) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  $workbook = $null
  $excel.Quit() | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  $excel = $null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  $pull1 = & ./pull.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir1 -ClassesDir $classesDir1 -FormsDir $formsDir1 -WorkbookDir $workbookDir1 -Visible false | ConvertFrom-Json
  $result.initialPullStatus = $pull1.status
  if ($pull1.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $push = & ./push.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir1 -ClassesDir $classesDir1 -FormsDir $formsDir1 -WorkbookDir $workbookDir1 -BackupRoot $backupRoot -Visible false | ConvertFrom-Json
  $result.pushStatus = $push.status
  if ($push.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $pull2 = & ./pull.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir2 -ClassesDir $classesDir2 -FormsDir $formsDir2 -WorkbookDir $workbookDir2 -Visible false | ConvertFrom-Json
  $result.roundtripPullStatus = $pull2.status
  $result.frmPath = Join-Path $formsDir2 ($formName + '.frm')
  $result.frxPath = Join-Path $formsDir2 ($formName + '.frx')
  $result.frxExists = Test-Path -LiteralPath $result.frxPath
  if (Test-Path -LiteralPath $result.frmPath) {
    $frmContent = Get-Content -Raw -LiteralPath $result.frmPath
    $frxName = [System.IO.Path]::GetFileName($result.frxPath)
    $result.frmReferencesFrx = $frmContent -match ('OleObjectBlob\s*=\s*".*' + [regex]::Escape($frxName) + '.*":')
  }
  $result.frxIsSibling = ((Split-Path -Parent $result.frmPath) -eq (Split-Path -Parent $result.frxPath))
  $result | ConvertTo-Json -Compress
} catch {
  $result.skip = $false
  $result.skipReason = ''
  $result.pullStatus = 'error: ' + $_.Exception.Message
  $result | ConvertTo-Json -Compress
  exit 1
} finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($false) | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
  }
  if ($null -ne $excel) {
    try { $excel.Quit() | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("userform companion check failed: %v\n%s", err, out)
	}

	var got struct {
		Skip             bool   `json:"skip"`
		SkipReason       string `json:"skipReason"`
		FrmPath          string `json:"frmPath"`
		FrxPath          string `json:"frxPath"`
		FrxExists        bool   `json:"frxExists"`
		FrmReferencesFrx bool   `json:"frmReferencesFrx"`
		FrxIsSibling     bool   `json:"frxIsSibling"`
		InitialPullStatus string `json:"initialPullStatus"`
		PushStatus        string `json:"pushStatus"`
		RoundtripPullStatus string `json:"roundtripPullStatus"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse powershell output as json: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skipf("skipped: %s", got.SkipReason)
	}
	if got.InitialPullStatus != "ok" {
		t.Fatalf("expected initial pull.ps1 to succeed, got status=%q output=%s", got.InitialPullStatus, out)
	}
	if got.PushStatus != "ok" {
		t.Fatalf("expected push.ps1 to succeed, got status=%q output=%s", got.PushStatus, out)
	}
	if got.RoundtripPullStatus != "ok" {
		t.Fatalf("expected roundtrip pull.ps1 to succeed, got status=%q output=%s", got.RoundtripPullStatus, out)
	}

	normalizedFrmPath := strings.ReplaceAll(got.FrmPath, "\\", "/")
	normalizedFrxPath := strings.ReplaceAll(got.FrxPath, "\\", "/")
	if !strings.HasSuffix(normalizedFrmPath, "/forms/UserFormFrxRegression.frm") {
		t.Fatalf("expected userform export path to end with forms/UserFormFrxRegression.frm: %q", got.FrmPath)
	}
	if !strings.HasSuffix(normalizedFrxPath, "/forms/UserFormFrxRegression.frx") {
		t.Fatalf("expected userform companion path to end with forms/UserFormFrxRegression.frx: %q", got.FrxPath)
	}
	if !got.FrxExists {
		t.Fatal("expected .frx companion file to exist beside .frm")
	}
	if !got.FrmReferencesFrx {
		t.Fatal("expected .frm content to reference its .frx companion")
	}
	if !got.FrxIsSibling {
		t.Fatal("expected .frx companion to be created in the same directory as .frm")
	}
}
