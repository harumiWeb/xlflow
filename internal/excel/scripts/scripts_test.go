package scripts_test

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"
)

func TestPowerShellScriptsParse(t *testing.T) {
	scripts := []string{"attach.ps1", "common.ps1", "doctor.ps1", "edit.ps1", "export-image.ps1", "form-export-image.ps1", "form-write.ps1", "inspect-form.ps1", "list.ps1", "macros.ps1", "new.ps1", "pull.ps1", "push.ps1", "run.ps1", "runner.ps1", "session.ps1", "test.ps1", "trace.ps1", "ui.ps1"}
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

func TestFormWriteScriptValidatesArgsBeforeWorkbookOpen(t *testing.T) {
	specJSON := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"UserForm1"},"controls":[],"warnings":[]}`
	specJSON64 := base64.StdEncoding.EncodeToString([]byte(specJSON))
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./form-write.ps1 -Action build -SpecJson64 '"+specJSON64+"' -NoSave true -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-write validation command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse form-write output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "form_build_args_invalid" {
		t.Fatalf("expected form_build_args_invalid failure, got %+v", got)
	}
}

func TestFormWriteScriptRejectsOverwriteWithNoSaveBeforeWorkbookOpen(t *testing.T) {
	specJSON := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"UserForm1"},"controls":[],"warnings":[]}`
	specJSON64 := base64.StdEncoding.EncodeToString([]byte(specJSON))
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./form-write.ps1 -Action build -SpecJson64 '"+specJSON64+"' -FormsDir 'C:\\forms' -Overwrite true -NoSave true -UseSession true -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-write overwrite/no-save validation command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse form-write output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "form_build_args_invalid" {
		t.Fatalf("expected form_build_args_invalid failure, got %+v", got)
	}
	if !strings.Contains(got.Error.Message, "--overwrite cannot be combined with --NoSave") {
		t.Fatalf("unexpected validation message: %+v", got)
	}
}

func TestFormWriteScriptRequiresFormsDirBeforeWorkbookOpen(t *testing.T) {
	specJSON := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"UserForm1"},"controls":[],"warnings":[]}`
	specJSON64 := base64.StdEncoding.EncodeToString([]byte(specJSON))
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./form-write.ps1 -Action build -SpecJson64 '"+specJSON64+"' -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-write FormsDir validation command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse form-write output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "form_build_args_invalid" {
		t.Fatalf("expected form_build_args_invalid failure, got %+v", got)
	}
	if !strings.Contains(got.Error.Message, "FormsDir is required.") {
		t.Fatalf("unexpected validation message: %+v", got)
	}
}

func TestFormWriteScriptUsesDesignerApiAndSessionSaveWarnings(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "form-write.ps1"))
	if err != nil {
		t.Fatalf("failed to read form-write.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"VBComponents.Add(3)",
		"Controls.Add($progId, $controlName, $true)",
		"Clear-XlflowDesignerControls",
		"Controls.Item($Container.Controls.Count - 1)",
		"Where-Object { $null -ne $_ }",
		"Get-XlflowRootControlSpecs",
		"Get-XlflowControlSpecChildren",
		"Set-XlflowVBComponentProperty",
		"Export-XlflowVBComponentBackup",
		"Import-XlflowVBComponentBackup",
		"Sync-XlflowUserFormCodeBehind",
		"Export-XlflowBuiltUserFormArtifacts",
		"failed to remove partially created UserForm after name assignment failure",
		"Get-XlflowCodeModuleText -CodeModule $existing.CodeModule",
		"Add-XlflowFormContractWarnings",
		"best_effort_form_size",
		"best_effort_list_state",
		"field_path",
		"component '\" + $Name + \"' exists but is not a UserForm",
		"save_required",
		"userform_review_commands",
		"synchronized UserForm source artifacts for",
		"FormsDir is required.",
		"Invoke-XlflowFormApply -VBProject $workbook.VBProject -Spec $spec -FormsDir $FormsDir -CodeSource $CodeSource",
		"[void](Sync-XlflowUserFormCodeBehind -Component $component -FormsDir $FormsDir)",
		"--NoSave requires --UseSession",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-write.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestNormalizeXlflowUserFormArtifactFileSkipsUnreadableCaption(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.frm')
try {
  $before = @('VERSION 5.00','Begin {GUID} RegistrationForm','   Caption         =   "KeepMe"','   ClientHeight    =   3036','End','Attribute VB_Name = "RegistrationForm"') -join [Environment]::NewLine
  Set-XlflowUtf8Text -Path $tmp -Text $before
  Normalize-XlflowUserFormArtifactFile -Path $tmp -Caption $null
  Get-XlflowUtf8Text -Path $tmp
} finally {
  if (Test-Path -LiteralPath $tmp) {
    Remove-Item -LiteralPath $tmp -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Normalize-XlflowUserFormArtifactFile null caption failed: %v\n%s", err, out)
	}
	got := strings.ReplaceAll(string(out), "\r\n", "\n")
	if !strings.Contains(got, `Caption         =   "KeepMe"`) {
		t.Fatalf("expected caption to remain unchanged, got %q", got)
	}
}

func TestCommonScriptTreatsUserFormCodeSidecarsSeparately(t *testing.T) {
	root := t.TempDir()
	modulesDir := filepath.Join(root, "src", "modules")
	formsDir := filepath.Join(root, "src", "forms")
	workbookDir := filepath.Join(root, "src", "workbook")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{
		filepath.Join(modulesDir, "Main.bas"),
		filepath.Join(formsDir, "CalendarPicker.frm"),
		filepath.Join(formsDir, "CalendarPicker.frx"),
		filepath.Join(formsDir, "code", "CalendarPicker.bas"),
		filepath.Join(workbookDir, "ThisWorkbook.bas"),
	} {
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte("test"), 0o644); err != nil {
			t.Fatal(err)
		}
	}

	command := fmt.Sprintf(`. ./common.ps1; $files = @(Get-XlflowSourceComponentFiles -ModulesDir '%s' -ClassesDir '' -FormsDir '%s' -WorkbookDir '%s' -CodeSource 'sidecar'); $files | ConvertTo-Json -Depth 5 -Compress`,
		modulesDir, formsDir, workbookDir)
	cmd := exec.Command("pwsh", "-NoProfile", "-Command", command)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Get-XlflowSourceComponentFiles command failed: %v\n%s", err, out)
	}

	var got []struct {
		Kind         string `json:"kind"`
		RelativePath string `json:"relative_path"`
		ModuleName   string `json:"module_name"`
		FormName     string `json:"form_name"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse source component files: %v\n%s", err, out)
	}

	var kinds []string
	for _, file := range got {
		kinds = append(kinds, file.Kind+":"+file.RelativePath)
	}
	sort.Strings(kinds)
	want := []string{
		"document:ThisWorkbook.bas",
		"form:CalendarPicker.frm",
		"form:CalendarPicker.frx",
		"form_code:CalendarPicker.bas",
		"module:Main.bas",
	}
	if !reflect.DeepEqual(kinds, want) {
		t.Fatalf("source component files = %#v, want %#v", kinds, want)
	}
}

func TestCommonScriptOmitsUserFormCodeSidecarsInFRMMode(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "CustomerForm.frm"), []byte("test"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "CustomerForm.bas"), []byte("test"), 0o644); err != nil {
		t.Fatal(err)
	}
	command := fmt.Sprintf(`. ./common.ps1; $files = @(Get-XlflowSourceComponentFiles -ModulesDir '' -ClassesDir '' -FormsDir '%s' -WorkbookDir '' -CodeSource 'frm'); $files | ConvertTo-Json -Depth 5 -Compress`, formsDir)
	cmd := exec.Command("pwsh", "-NoProfile", "-Command", command)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Get-XlflowSourceComponentFiles frm-mode command failed: %v\n%s", err, out)
	}
	type sourceFile struct {
		Kind string `json:"kind"`
	}
	var got []sourceFile
	if err := json.Unmarshal(out, &got); err != nil {
		var single sourceFile
		if errSingle := json.Unmarshal(out, &single); errSingle != nil {
			t.Fatalf("failed to parse frm-mode source component files: %v\n%s", err, out)
		}
		got = []sourceFile{single}
	}
	if len(got) != 1 || got[0].Kind != "form" {
		t.Fatalf("got %#v, want only form entries", got)
	}
}

func TestFormWriteScriptCommunicatesWeakDesignerContractFields(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "form-write.ps1"))
	if err != nil {
		t.Fatalf("failed to read form-write.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`Add-XlflowFormWriteWarning -Code "best_effort_form_size"`,
		`Form-level width/height are best-effort`,
		`form.observed.width`,
		`Add-XlflowFormWriteWarning -Code "best_effort_list_state"`,
		`observed-only for round-trip expectations`,
		`controls[*].selectedIndex`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("expected weak-field contract warning %q in form-write.ps1:\n%s", want, text)
		}
	}
}

func TestFormWriteScriptUsesSnapshotDimensionsWithoutOffset(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$script:XlflowLoadFunctionsOnly = $true; . ./form-write.ps1; "+
			"$cases = [ordered]@{ "+
			"observedOnly = (Get-XlflowUserFormBuildDimensions -FormSpec ([pscustomobject]@{ observed = [pscustomobject]@{ width = 300; height = 262 } })); "+
			"buildPreferred = (Get-XlflowUserFormBuildDimensions -FormSpec ([pscustomobject]@{ build = [pscustomobject]@{ width = 301; height = 263 }; observed = [pscustomobject]@{ width = 300; height = 262 } })) "+
			"}; $cases | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-write dimension command failed: %v\n%s", err, out)
	}
	var got struct {
		ObservedOnly struct {
			Width  float64 `json:"width"`
			Height float64 `json:"height"`
		} `json:"observedOnly"`
		BuildPreferred struct {
			Width  float64 `json:"width"`
			Height float64 `json:"height"`
		} `json:"buildPreferred"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse form-write dimensions: %v\n%s", err, out)
	}
	if got.ObservedOnly.Width != 300 || got.ObservedOnly.Height != 262 {
		t.Fatalf("observed fallback = %+v, want width=300 height=262", got.ObservedOnly)
	}
	if got.BuildPreferred.Width != 301 || got.BuildPreferred.Height != 263 {
		t.Fatalf("build preference = %+v, want width=301 height=263", got.BuildPreferred)
	}
}

func TestFormExportImageScriptRepairsGenericRuntimeCaptionFromSourceDesigner(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "form-export-image.ps1"))
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`Private xlflowExpectedCaption As String`,
		`Optional ByVal expectedCaption As String = ""`,
		`xlflowExpectedCaption = Trim$(expectedCaption)`,
		`If Len(xlflowExpectedCaption) > 0 Then`,
		`Or LCase$(Left$(caption, 8)) = "userform" Then`,
		`function Get-XlflowFormExportSourceDesignerCaption`,
		`return [string]$component.Designer.Caption`,
		`-ExpectedCaption $sourceDesignerCaption`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("expected runtime caption repair %q in form-export-image.ps1:\n%s", want, text)
		}
	}
}

func TestFormExportImageScriptKeepsVisibleWindowOnContainingScreen(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$script:XlflowLoadFunctionsOnly = $true; . ./form-export-image.ps1; "+
			"$window = [pscustomobject]@{ left = -1500; top = 120; width = 400; height = 300 }; "+
			"$areas = @([pscustomobject]@{ left = -1920; top = 0; right = 0; bottom = 1080; width = 1920; height = 1080 }, [pscustomobject]@{ left = 0; top = 0; right = 1920; bottom = 1080; width = 1920; height = 1080 }); "+
			"$area = Get-XlflowBestWorkingAreaForWindowInfo -WindowInfo $window -WorkingAreas $areas; "+
			"$plan = Get-XlflowWindowCaptureRepositionPlan -WindowInfo $window -WorkArea $area -Margin 16; "+
			"[pscustomobject]@{ area = $area; plan = $plan } | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-export-image containing-screen command failed: %v\n%s", err, out)
	}
	var got struct {
		Area struct {
			Left int `json:"left"`
		} `json:"area"`
		Plan struct {
			Left  int  `json:"left"`
			Top   int  `json:"top"`
			Moved bool `json:"moved"`
		} `json:"plan"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse containing-screen output: %v\n%s", err, out)
	}
	if got.Area.Left != -1920 {
		t.Fatalf("best work area = %+v, want left=-1920", got.Area)
	}
	if got.Plan.Moved || got.Plan.Left != -1500 || got.Plan.Top != 120 {
		t.Fatalf("reposition plan = %+v, want unchanged visible window", got.Plan)
	}
}

func TestFormExportImageScriptRepositionsOnlyWithinContainingScreen(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$script:XlflowLoadFunctionsOnly = $true; . ./form-export-image.ps1; "+
			"$window = [pscustomobject]@{ left = -1918; top = -12; width = 400; height = 300 }; "+
			"$areas = @([pscustomobject]@{ left = -1920; top = 0; right = 0; bottom = 1080; width = 1920; height = 1080 }, [pscustomobject]@{ left = 0; top = 0; right = 1920; bottom = 1080; width = 1920; height = 1080 }); "+
			"$area = Get-XlflowBestWorkingAreaForWindowInfo -WindowInfo $window -WorkingAreas $areas; "+
			"Get-XlflowWindowCaptureRepositionPlan -WindowInfo $window -WorkArea $area -Margin 16 | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-export-image reposition command failed: %v\n%s", err, out)
	}
	var got struct {
		Left  int  `json:"left"`
		Top   int  `json:"top"`
		Moved bool `json:"moved"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse reposition output: %v\n%s", err, out)
	}
	if !got.Moved || got.Left != -1904 || got.Top != 16 {
		t.Fatalf("reposition plan = %+v, want left=-1904 top=16 moved=true", got)
	}
	if got.Left >= 0 {
		t.Fatalf("window should stay on the negative-coordinate monitor, got %+v", got)
	}
}

func TestFormExportImageScriptClampsCaptureScaleFromDPI(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$script:XlflowLoadFunctionsOnly = $true; . ./form-export-image.ps1; "+
			"[ordered]@{ low = (Get-XlflowClampedCaptureScale -Dpi 72); normal = (Get-XlflowClampedCaptureScale -Dpi 96); scaled = (Get-XlflowClampedCaptureScale -Dpi 144); capped = (Get-XlflowClampedCaptureScale -Dpi 600) } | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-export-image dpi scale command failed: %v\n%s", err, out)
	}
	var got struct {
		Low    float64 `json:"low"`
		Normal float64 `json:"normal"`
		Scaled float64 `json:"scaled"`
		Capped float64 `json:"capped"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse dpi scale output: %v\n%s", err, out)
	}
	if got.Low != 1.0 || got.Normal != 1.0 || got.Scaled != 1.5 || got.Capped != 4.0 {
		t.Fatalf("unexpected capture scales: %+v", got)
	}
}

func TestFormExportImageScriptTrimsBlackEdgesAfterCapture(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$script:XlflowLoadFunctionsOnly = $true; . ./form-export-image.ps1; "+
			"Add-Type -AssemblyName System.Drawing; "+
			"$bitmap = New-Object System.Drawing.Bitmap(64, 64); "+
			"for ($x = 0; $x -lt 64; $x++) { for ($y = 0; $y -lt 64; $y++) { $bitmap.SetPixel($x, $y, [System.Drawing.Color]::Black) } }; "+
			"for ($x = 10; $x -lt 54; $x++) { for ($y = 10; $y -lt 54; $y++) { $bitmap.SetPixel($x, $y, [System.Drawing.Color]::White) } }; "+
			"$trimmed = Trim-XlflowBitmapBlackEdges -Bitmap $bitmap; "+
			"[pscustomobject]@{ width = $trimmed.Width; height = $trimmed.Height } | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-export-image trim command failed: %v\n%s", err, out)
	}
	var got struct {
		Width  int `json:"width"`
		Height int `json:"height"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse trim output: %v\n%s", err, out)
	}
	if got.Width != 44 || got.Height != 44 {
		t.Fatalf("trimmed size = %+v, want 44x44", got)
	}
}

func TestFormWriteScriptOverwritePathBacksUpAndRestoresOnFailure(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "form-write.ps1"))
	if err != nil {
		t.Fatalf("failed to read form-write.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"New-XlflowFormRestoreDirectory",
		"Export-XlflowVBComponentBackup -Component $existing -Directory $restoreDirectory",
		"Import-XlflowVBComponentBackup -VBProject $VBProject -ExportPath $restorePath -ExpectedName $formName",
		"Remove-XlflowVBComponentInstance -VBProject $VBProject -Component $component",
		"restored original UserForm '\" + $formName + \"' after overwrite failure",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-write.ps1 missing overwrite restore handling %q:\n%s", want, text)
		}
	}
}

func TestPullAndPushScriptsHandleUserFormCodeSidecars(t *testing.T) {
	checks := map[string][]string{
		"pull.ps1": {
			"Export-XlflowUserFormCodeBehind -Component $component -FormsDir $FormsDir",
			"Use-XlflowUserFormCodeSidecar -CodeSource $CodeSource",
			"Normalize-XlflowUserFormArtifactFile -Path $path -Caption (Get-XlflowUserFormDesignerCaption -Component $component)",
			"exported $($exportedFormCode.Count) UserForm code-behind sidecar(s)",
		},
		"push.ps1": {
			`Where-Object { $_.kind -ne "form_code" }`,
			"Sync-XlflowUserFormCodeBehind -Component $importedComponent -FormsDir $FormsDir",
			"Use-XlflowUserFormCodeSidecar -CodeSource $CodeSource",
			"synced $($syncedFormCode.Count) UserForm code-behind sidecar(s)",
		},
	}
	for script, wants := range checks {
		data, err := os.ReadFile(filepath.Join(".", script))
		if err != nil {
			t.Fatalf("failed to read %s: %v", script, err)
		}
		text := string(data)
		for _, want := range wants {
			if !strings.Contains(text, want) {
				t.Fatalf("%s missing %q:\n%s", script, want, text)
			}
		}
	}
}

func TestNormalizeXlflowUserFormArtifactFileUpdatesCaptionLine(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.frm')
try {
  Set-XlflowUtf8Text -Path $tmp -Text (@('VERSION 5.00','Begin {GUID} RegistrationForm','   Caption         =   "UserForm1"','   ClientHeight    =   3036','End','Attribute VB_Name = "RegistrationForm"') -join [Environment]::NewLine)
  Normalize-XlflowUserFormArtifactFile -Path $tmp -Caption 'RegistrationForm'
  Get-XlflowUtf8Text -Path $tmp
} finally {
  if (Test-Path -LiteralPath $tmp) {
    Remove-Item -LiteralPath $tmp -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Normalize-XlflowUserFormArtifactFile failed: %v\n%s", err, out)
	}
	got := strings.ReplaceAll(string(out), "\r\n", "\n")
	if !strings.Contains(got, `Caption         =   "RegistrationForm"`) {
		t.Fatalf("expected normalized caption line, got %q", got)
	}
	if strings.Contains(got, `Caption         =   "UserForm1"`) {
		t.Fatalf("expected stale caption line to be removed, got %q", got)
	}
}

func TestNormalizeXlflowUserFormArtifactFileInsertsMissingCaptionLine(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.frm')
try {
  Set-XlflowUtf8Text -Path $tmp -Text (@('VERSION 5.00','Begin {GUID} RegistrationForm','   ClientHeight    =   3036','End','Attribute VB_Name = "RegistrationForm"') -join [Environment]::NewLine)
  Normalize-XlflowUserFormArtifactFile -Path $tmp -Caption 'RegistrationForm'
  Get-XlflowUtf8Text -Path $tmp
} finally {
  if (Test-Path -LiteralPath $tmp) {
    Remove-Item -LiteralPath $tmp -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Normalize-XlflowUserFormArtifactFile insert failed: %v\n%s", err, out)
	}
	got := strings.ReplaceAll(string(out), "\r\n", "\n")
	if !strings.Contains(got, `Begin {GUID} RegistrationForm`+"\n"+`   Caption         =   "RegistrationForm"`) {
		t.Fatalf("expected inserted caption line after Begin, got %q", got)
	}
}

func TestFormWriteScriptMatchesParentIDsCaseSensitively(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "form-write.ps1"))
	if err != nil {
		t.Fatalf("failed to read form-write.ps1: %v", err)
	}
	text := string(data)
	start := strings.Index(text, "function Get-XlflowControlSpecChildren")
	if start < 0 {
		t.Fatalf("Get-XlflowControlSpecChildren not found:\n%s", text)
	}
	end := strings.Index(text[start:], "function Get-XlflowRootControlSpecs")
	if end < 0 {
		t.Fatalf("Get-XlflowRootControlSpecs boundary not found:\n%s", text)
	}
	section := text[start : start+end]
	if !strings.Contains(section, "[System.StringComparison]::Ordinal") {
		t.Fatalf("expected case-sensitive parentId matching in Get-XlflowControlSpecChildren:\n%s", section)
	}
	if strings.Contains(section, "[System.StringComparison]::OrdinalIgnoreCase") {
		t.Fatalf("unexpected case-insensitive parentId matching in Get-XlflowControlSpecChildren:\n%s", section)
	}
}

func TestFormWriteScriptAppliesSelectedIndexAfterListPopulation(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "form-write.ps1"))
	if err != nil {
		t.Fatalf("failed to read form-write.ps1: %v", err)
	}
	text := string(data)
	listIndex := strings.Index(text, "Set-XlflowControlListItems -Control $Control -ControlSpec $ControlSpec")
	selectedIndex := strings.Index(text, `Set-XlflowFormProperty -Target $Control -PropertyName "ListIndex"`)
	if listIndex < 0 || selectedIndex < 0 {
		t.Fatalf("expected list population and selected index assignment in form-write.ps1:\n%s", text)
	}
	if selectedIndex < listIndex {
		t.Fatalf("expected selected index assignment after list population")
	}
}

func TestFormWriteScriptDecodesSpecInWindowsPowerShell(t *testing.T) {
	if _, err := exec.LookPath("powershell"); err != nil {
		t.Skip("powershell is not available")
	}

	specJSON := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"UserForm1"},"controls":[{"name":"Label1","type":"Label"}],"warnings":[]}`
	specJSON64 := base64.StdEncoding.EncodeToString([]byte(specJSON))
	cmd := exec.Command(
		"powershell",
		"-NoProfile",
		"-Command",
		"$r = ./form-write.ps1 -Action build -SpecJson64 '"+specJSON64+"' -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-write decode check failed in Windows PowerShell: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse form-write decode output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil {
		t.Fatalf("unexpected decoded form spec result: %+v", got)
	}
	if strings.Contains(got.Error.Message, "invalid form spec payload") {
		t.Fatalf("unexpected decoded form spec result: %+v", got)
	}
}

func TestFormExportImageScriptValidatesArgsBeforeWorkbookOpen(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./form-export-image.ps1 -FormName 'UserForm1' -OutputPath 'C:\\temp\\UserForm1.webp' -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("form-export-image validation command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse form-export-image output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "unsupported_image_format" {
		t.Fatalf("expected unsupported_image_format failure, got %+v", got)
	}
}

func TestFormExportImageScriptWaitsForCaptureStatusBeforeFallback(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "form-export-image.ps1"))
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	want := "if (-not [bool]$captureStatus.ready -and [int64]$captureStatus.hwnd -eq 0 -and [string]::IsNullOrWhiteSpace([string]$captureStatus.caption))"
	if !strings.Contains(text, want) {
		t.Fatalf("form-export-image.ps1 missing empty-status wait guard %q:\n%s", want, text)
	}
	for _, want := range []string{
		"function Wait-XlflowStableWindowCaptureInfo",
		"$stableCount -ge $StableSamples",
		"Wait-XlflowStableWindowCaptureInfo -Hwnd",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-export-image.ps1 missing stable-window guard %q:\n%s", want, text)
		}
	}
}

func trailingJSONLine(out []byte) []byte {
	lines := strings.Split(strings.TrimSpace(string(out)), "\n")
	for index := len(lines) - 1; index >= 0; index-- {
		line := strings.TrimSpace(lines[index])
		if strings.HasPrefix(line, "{") && strings.HasSuffix(line, "}") {
			return []byte(line)
		}
	}
	return out
}

func TestInspectFormScriptValidatesBasisBeforeWorkbookOpen(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./inspect-form.ps1 -Basis nope -FormName 'UserForm1' -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("inspect-form basis validation command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse inspect-form output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "inspect_form_args_invalid" {
		t.Fatalf("expected inspect_form_args_invalid failure, got %+v", got)
	}
}

func TestInspectFormScriptUsesTemporaryHelperModuleAndWarnings(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "inspect-form.ps1"))
	if err != nil {
		t.Fatalf("failed to read inspect-form.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"Install-XlflowVBComponentFromCode",
		"New-XlflowInspectFormModuleName",
		"New-XlflowInspectRuntimeWorkbookCopy",
		"SaveCopyAs($tempPath)",
		"New-XlflowInspectFormModuleCode",
		"runtime_form_loads_initialize",
		"runtime_form_temp_copy",
		"temporary_component_cleanup_failed",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("inspect-form.ps1 missing %q:\n%s", want, text)
		}
	}
	for _, want := range []string{
		`Get-XlflowSafeMemberValue -Target $_ -Name "Parent"`,
		"Get-XlflowDesignerControlSnapshot -Control $_",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("inspect-form.ps1 missing nested-control parent filter %q:\n%s", want, text)
		}
	}
	if strings.Contains(text, "runtime_inspect_session_dirty") {
		t.Fatalf("inspect-form.ps1 should no longer report live-session dirty mutation for runtime temp-copy inspection:\n%s", text)
	}
}

func TestCommonScriptStrictDesignerFiltersControlsByParentName(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "common.ps1"))
	if err != nil {
		t.Fatalf("failed to read common.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"SerializeControls(controls, formName)",
		"ControlHasExpectedParent",
		"SafeControlName(control)",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("common.ps1 missing strict designer parent filtering %q:\n%s", want, text)
		}
	}
}

func TestInspectFormScriptDesignerDoesNotRequireRunnableVBA(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`
$result = [ordered]@{
	skip = $false
	skipReason = ''
	status = ''
	errorCode = ''
	types = @()
}

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('xlflow-inspect-form-designer-' + [guid]::NewGuid().ToString())
$wbPath = Join-Path $root 'DesignerTypes.xlsm'
$excel = $null
$workbook = $null

try {
	New-Item -ItemType Directory -Path $root -Force | Out-Null

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
		$result.skipReason = 'VBProject access is unavailable: ' + $_.Exception.Message
		$result | ConvertTo-Json -Compress
		exit 0
	}

	$component = $workbook.VBProject.VBComponents.Add(3)
	$component.Name = 'DesignerTypesForm'
	$designer = $component.Designer

	$brokenModule = $workbook.VBProject.VBComponents.Add(1)
	$brokenModule.Name = 'BrokenCompileModule'
	$brokenModule.CodeModule.AddFromString("Option Explicit" + [Environment]::NewLine + "Public Sub Broken()" + [Environment]::NewLine + "  Dim missing As" + [Environment]::NewLine + "End Sub")

	$label = $designer.Controls.Add('Forms.Label.1')
	$label.Name = 'lblCaption'
	$label.Caption = 'Designer label'

	$textBox = $designer.Controls.Add('Forms.TextBox.1')
	$textBox.Name = 'txtValue'
	$textBox.Text = 'Designer text'

	$workbook.SaveAs($wbPath, 52)

	$inspect = & ./inspect-form.ps1 -Basis designer -FormName 'DesignerTypesForm' -WorkbookPath $wbPath | ConvertFrom-Json
	$result.status = $inspect.status
	if ($null -ne $inspect.error) {
		$result.errorCode = $inspect.error.code
	}
	if ($null -ne $inspect.forms -and $null -ne $inspect.forms.controls) {
		$result.types = @($inspect.forms.controls | ForEach-Object { [string]$_.name })
	}

	$result | ConvertTo-Json -Compress
} catch {
	$result.skip = $false
	$result.skipReason = ''
	$result.status = 'command_failed'
	$result.errorCode = $_.Exception.Message
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
		t.Fatalf("inspect-form designer compile-tolerance check failed: %v\n%s", err, out)
	}

	var got struct {
		Skip       bool     `json:"skip"`
		SkipReason string   `json:"skipReason"`
		Status     string   `json:"status"`
		ErrorCode  string   `json:"errorCode"`
		Types      []string `json:"types"`
	}
	if err := json.Unmarshal(trailingJSONLine(out), &got); err != nil {
		t.Fatalf("failed to parse inspect-form designer output: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skip(got.SkipReason)
	}
	if got.Status != "ok" {
		t.Fatalf("expected inspect-form designer status ok, got %+v", got)
	}
	want := []string{"lblCaption", "txtValue"}
	if !reflect.DeepEqual(got.Types, want) {
		t.Fatalf("designer control names = %#v, want %#v", got.Types, want)
	}
}

func TestInspectFormScriptStrictDesignerReturnsConcreteControlTypes(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`
$result = [ordered]@{
	skip = $false
	skipReason = ''
	status = ''
	errorCode = ''
	types = @()
}

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('xlflow-inspect-form-strict-designer-' + [guid]::NewGuid().ToString())
$wbPath = Join-Path $root 'DesignerTypes.xlsm'
$excel = $null
$workbook = $null

try {
	New-Item -ItemType Directory -Path $root -Force | Out-Null

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
		$result.skipReason = 'VBProject access is unavailable: ' + $_.Exception.Message
		$result | ConvertTo-Json -Compress
		exit 0
	}

	$component = $workbook.VBProject.VBComponents.Add(3)
	$component.Name = 'DesignerTypesForm'
	$designer = $component.Designer

	$label = $designer.Controls.Add('Forms.Label.1')
	$label.Name = 'lblCaption'
	$label.Caption = 'Designer label'

	$textBox = $designer.Controls.Add('Forms.TextBox.1')
	$textBox.Name = 'txtValue'
	$textBox.Text = 'Designer text'

	$workbook.SaveAs($wbPath, 52)

	$inspect = & ./inspect-form.ps1 -Basis designer -StrictDesigner true -FormName 'DesignerTypesForm' -WorkbookPath $wbPath | ConvertFrom-Json
	$result.status = $inspect.status
	if ($null -ne $inspect.error) {
		$result.errorCode = $inspect.error.code
	}
	if ($null -ne $inspect.forms -and $null -ne $inspect.forms.controls) {
		$result.types = @($inspect.forms.controls | ForEach-Object { [string]$_.type })
	}

	$result | ConvertTo-Json -Compress
} catch {
	$result.skip = $false
	$result.skipReason = ''
	$result.status = 'command_failed'
	$result.errorCode = $_.Exception.Message
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
		t.Fatalf("inspect-form strict designer concrete type check failed: %v\n%s", err, out)
	}

	var got struct {
		Skip       bool     `json:"skip"`
		SkipReason string   `json:"skipReason"`
		Status     string   `json:"status"`
		ErrorCode  string   `json:"errorCode"`
		Types      []string `json:"types"`
	}
	if err := json.Unmarshal(trailingJSONLine(out), &got); err != nil {
		t.Fatalf("failed to parse strict designer output: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skip(got.SkipReason)
	}
	if got.Status != "ok" {
		t.Fatalf("expected strict designer status ok, got %+v", got)
	}
	want := []string{"Label", "TextBox"}
	if !reflect.DeepEqual(got.Types, want) {
		t.Fatalf("strict designer control types = %#v, want %#v", got.Types, want)
	}
}

func TestCommonScriptInstallVBComponentRefusesToReplaceExistingComponent(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "common.ps1"))
	if err != nil {
		t.Fatalf("failed to read common.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"function Get-XlflowVBComponentByName",
		"VBA component '\" + $Name + \"' already exists.",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("common.ps1 missing %q:\n%s", want, text)
		}
	}
	if strings.Contains(text, "[void](Remove-XlflowVBComponentByName -VBProject $VBProject -Name $Name)") {
		t.Fatalf("Install-XlflowVBComponentFromCode should not blindly remove existing components:\n%s", text)
	}
}

func TestListScriptValidatesActionBeforeWorkbookOpen(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./list.ps1 -Action nope -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("list action validation command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse list output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "list_args_invalid" {
		t.Fatalf("expected list_args_invalid failure, got %+v", got)
	}
}

func TestListScriptUsesFormComponentPathAndPortableRelativePaths(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "list.ps1"))
	if err != nil {
		t.Fatalf("failed to read list.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"Get-XlflowComponentPath -Component $component",
		"ConvertTo-XlflowPortableRelativePath",
		"component_type = \"MSForm\"",
		"Get-XlflowUserFormComponents -Workbook $workbook",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("list.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestListScriptPreservesSaveRequiredWarningOnFailurePaths(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "list.ps1"))
	if err != nil {
		t.Fatalf("failed to read list.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"function Add-XlflowListSaveRequiredWarning",
		"Add-XlflowListSaveRequiredWarning -Result $result -SaveState $saveState",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("list.ps1 missing %q:\n%s", want, text)
		}
	}
	if count := strings.Count(text, "Add-XlflowListSaveRequiredWarning -Result $result -SaveState $saveState"); count < 3 {
		t.Fatalf("expected save-required warning helper on success and failure paths, found %d:\n%s", count, text)
	}
}

func TestEditScriptRequiresActiveSessionBeforeExcelMutation(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./edit.ps1 -Action cell -WorkbookPath 'C:\\missing.xlsm' -Sheet 'Input' -Cell 'B2' -Value 'ABC123' -UseSession false | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("edit session-required command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse edit output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "session_required" {
		t.Fatalf("expected session_required failure, got %+v", got)
	}
}

func TestEditScriptPreservesStructuredErrorBeforeFallbackCatch(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "edit.ps1"))
	if err != nil {
		t.Fatalf("failed to read edit.ps1: %v", err)
	}
	text := string(data)
	if !strings.Contains(text, "if ($null -eq $result.error)") {
		t.Fatalf("expected edit.ps1 to preserve pre-classified errors before assigning a fallback:\n%s", text)
	}
	if !strings.Contains(text, "Add-XlflowHint -Result $result -Code \"possible_event_handler_failure\"") {
		t.Fatalf("expected edit.ps1 to keep event-handler context as a hint instead of forcing a new error code:\n%s", text)
	}
}

func TestEditScriptValidatesEventsAndUpdatesSaveStateAfterMutation(t *testing.T) {
	data, err := os.ReadFile(filepath.Join(".", "edit.ps1"))
	if err != nil {
		t.Fatalf("failed to read edit.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"Set-EditValidationError -Code \"edit_args_invalid\" -Message \"-Events must be keep, on, or off.\"",
		"function Update-XlflowEditResultSaveState",
		"Update-XlflowEditResultSaveState -Workbook $workbook",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("expected edit.ps1 to contain %q:\n%s", want, text)
		}
	}
}

func TestCommonScriptRelativePathHelperWorksInWindowsPowerShell(t *testing.T) {
	if _, err := exec.LookPath("powershell"); err != nil {
		t.Skip("powershell is not available")
	}

	cmd := exec.Command(
		"powershell",
		"-NoProfile",
		"-Command",
		". ./common.ps1; Get-XlflowRelativePath -BasePath 'C:\\repo\\src\\modules' -TargetPath 'C:\\repo\\src\\modules\\Domain\\Services\\Main.bas'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("relative path helper failed in Windows PowerShell: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "Domain\\Services\\Main.bas" {
		t.Fatalf("relative path = %q, want %q", out, "Domain\\Services\\Main.bas")
	}
}

func TestCommonScriptRelativePathHelperPreservesAbsoluteTargetAcrossDrives(t *testing.T) {
	if _, err := exec.LookPath("powershell"); err != nil {
		t.Skip("powershell is not available")
	}

	cmd := exec.Command(
		"powershell",
		"-NoProfile",
		"-Command",
		". ./common.ps1; Get-XlflowRelativePath -BasePath 'C:\\repo\\src\\modules' -TargetPath 'D:\\shared\\Main.bas'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("cross-drive relative path helper failed in Windows PowerShell: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "D:\\shared\\Main.bas" {
		t.Fatalf("relative path = %q, want %q", out, "D:\\shared\\Main.bas")
	}
}

func TestCommonScriptExposesReleaseComObjectHelper(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; (Get-Command Release-XlflowComObject).CommandType",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Release-XlflowComObject helper check failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "Function" {
		t.Fatalf("expected Release-XlflowComObject helper, got %q", out)
	}
}

func TestUIButtonIdAndNameNormalization(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; ConvertTo-XlflowUIButtonId -Value 'Main.Run Aggregation'; ConvertTo-XlflowUIButtonName -Id 'Main.Run Aggregation'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("button id normalization failed: %v\n%s", err, out)
	}
	lines := strings.Fields(strings.TrimSpace(string(out)))
	if len(lines) != 2 {
		t.Fatalf("unexpected output: %q", out)
	}
	if lines[0] != "main-run-aggregation" {
		t.Fatalf("id = %q, want main-run-aggregation", lines[0])
	}
	if lines[1] != "xlflow.button.main-run-aggregation" {
		t.Fatalf("name = %q, want xlflow.button.main-run-aggregation", lines[1])
	}
}

func TestNewXlflowResultIncludesBridgeMetadata(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; New-XlflowResult -Command 'run' | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("bridge metadata command failed: %v\n%s", err, out)
	}
	var got struct {
		Bridge *struct {
			Host    string `json:"host"`
			Edition string `json:"edition"`
			Version string `json:"version"`
		} `json:"bridge"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse bridge metadata output: %v\n%s", err, out)
	}
	if got.Bridge == nil || got.Bridge.Host == "" || got.Bridge.Edition == "" || got.Bridge.Version == "" {
		t.Fatalf("expected bridge metadata, got %+v", got)
	}
}

func TestAddXlflowWarningAppendsToOrderedResult(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $r = New-XlflowResult -Command 'export-image'; Add-XlflowWarning -Result $r -Code 'cleanup_failed' -Message 'warning'; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Add-XlflowWarning command failed: %v\n%s", err, out)
	}
	var got struct {
		Warnings []struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"warnings"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse warning output: %v\n%s", err, out)
	}
	if len(got.Warnings) != 1 || got.Warnings[0].Code != "cleanup_failed" || got.Warnings[0].Message != "warning" {
		t.Fatalf("unexpected warnings: %+v", got.Warnings)
	}
}

func TestAddXlflowHintAppendsToOrderedResult(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $r = New-XlflowResult -Command 'macros'; Add-XlflowHint -Result $r -Code 'macros_empty_before_push' -Message 'push first'; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Add-XlflowHint command failed: %v\n%s", err, out)
	}
	var got struct {
		Hints []struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"hints"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse hint output: %v\n%s", err, out)
	}
	if len(got.Hints) != 1 || got.Hints[0].Code != "macros_empty_before_push" || got.Hints[0].Message != "push first" {
		t.Fatalf("unexpected hints: %+v", got.Hints)
	}
}

func TestNewXlflowTargetAndSessionResults(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; [ordered]@{ target = (New-XlflowTargetResult -Kind 'live_session' -Path 'build\\Book.xlsm'); session = (New-XlflowSessionResult -Active $true -WorkbookPath 'build\\Book.xlsm' -Dirty $true -SaveRequired $true -Mode 'explicit') } | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("target/session helper command failed: %v\n%s", err, out)
	}
	var got struct {
		Target struct {
			Kind        string `json:"kind"`
			Path        string `json:"path"`
			Description string `json:"description"`
		} `json:"target"`
		Session struct {
			Active       bool   `json:"active"`
			WorkbookPath string `json:"workbook_path"`
			Dirty        bool   `json:"dirty"`
			SaveRequired bool   `json:"save_required"`
			Mode         string `json:"mode"`
		} `json:"session"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse target/session output: %v\n%s", err, out)
	}
	if got.Target.Kind != "live_session" || got.Target.Path != "build\\Book.xlsm" || got.Target.Description == "" {
		t.Fatalf("unexpected target: %+v", got.Target)
	}
	if !got.Session.Active || !got.Session.Dirty || !got.Session.SaveRequired || got.Session.Mode != "explicit" {
		t.Fatalf("unexpected session: %+v", got.Session)
	}
}

func TestCloseXlflowComSkipsForceKillAfterGracefulCloseFailure(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; "+
			"$script:stopCalled = $false; "+
			"function Get-XlflowExcelProcessId { param($Excel) return 123 }; "+
			"function Get-Process { param([int]$Id) [pscustomobject]@{ Id = $Id } }; "+
			"function Stop-Process { param([int]$Id, [switch]$Force) $script:stopCalled = $true }; "+
			"function Start-Sleep { param([int]$Milliseconds) }; "+
			"function Release-XlflowComObject { param($Object, [string]$Name = 'COM object') }; "+
			"$workbook = New-Object psobject; "+
			"$workbook | Add-Member -MemberType ScriptMethod -Name Close -Value { param($Save) throw 'close failed' }; "+
			"$excel = New-Object psobject; "+
			"$excel | Add-Member -MemberType ScriptMethod -Name Quit -Value { return $null }; "+
			"Close-XlflowCom -Workbook $workbook -Excel $excel -Save $true; "+
			"$script:stopCalled | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Close-XlflowCom failure handling command failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "false" {
		t.Fatalf("expected Close-XlflowCom to skip force kill after graceful failure, got %q", out)
	}
}

func TestUIScriptRejectsUnsupportedActionAsStructuredFailure(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"./ui.ps1 -Action nope -WorkbookPath 'C:\\missing.xlsm' | ConvertFrom-Json | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("ui action command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse ui output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "ui_button_args_invalid" {
		t.Fatalf("expected structured ui failure, got %+v", got)
	}
}

func TestRunScriptAcceptsTimeoutParameter(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$command = Get-Command ./run.ps1; $command.Parameters.ContainsKey('TimeoutSeconds')",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run script timeout parameter check failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "True" {
		t.Fatalf("expected run.ps1 to expose TimeoutSeconds, got %q", out)
	}
}

func TestRunScriptAcceptsDiagnosticParameter(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$command = Get-Command ./run.ps1; $command.Parameters.ContainsKey('Diagnostic')",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run script diagnostic parameter check failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "True" {
		t.Fatalf("expected run.ps1 to expose Diagnostic, got %q", out)
	}
}

func TestRunScriptAcceptsSuppressModalErrorsParameter(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$command = Get-Command ./run.ps1; $command.Parameters.ContainsKey('SuppressModalErrors')",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run script suppress modal parameter check failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "True" {
		t.Fatalf("expected run.ps1 to expose SuppressModalErrors, got %q", out)
	}
}

func TestRunScriptWatchesAnyVBADialogDuringInvoke(t *testing.T) {
	data, err := os.ReadFile("run.ps1")
	if err != nil {
		t.Fatalf("failed to read run.ps1: %v", err)
	}
	text := string(data)
	if count := strings.Count(text, `DialogKind "any_vba"`); count < 2 {
		t.Fatalf("expected run.ps1 to watch any_vba dialogs during both direct and harness invoke paths, found %d:\n%s", count, text)
	}
	for _, want := range []string{
		"function Set-XlflowVBADialogFailure",
		`Set-XlflowError -Result $result -Code "vba_compile_failed"`,
		"function Find-XlflowPendingVBADialog",
		`CaptureOpenVBADialogs $SuppressModalErrors`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("run.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestOpenWorkbookHelperCanCaptureVBADialogsDuringIsolatedOpen(t *testing.T) {
	data, err := os.ReadFile("common.ps1")
	if err != nil {
		t.Fatalf("failed to read common.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`[string]$CaptureOpenVBADialogs = "false"`,
		`[int]$OpenDialogWaitMilliseconds = 1500`,
		`Invoke-XlflowExcelCallWithDialogWatch -Excel $excel -Workbook $null`,
		`open_dialog = $openDialog`,
		`open_selection = $openSelection`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("common.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestExportImageScriptUsesPrinterPictureCopyMode(t *testing.T) {
	data, err := os.ReadFile("export-image.ps1")
	if err != nil {
		t.Fatalf("failed to read export-image.ps1: %v", err)
	}
	text := string(data)
	if !strings.Contains(text, "$range.CopyPicture(2, -4147) | Out-Null") {
		t.Fatalf("expected export-image.ps1 to use printer-picture CopyPicture mode and suppress pipeline output")
	}
	for _, needle := range []string{
		"$range.Select() | Out-Null",
		"$excel.Visible = $true",
		"Test-Path -LiteralPath $resolvedOutputPath -PathType Container",
		"Move-Item -LiteralPath $temporaryExportPath -Destination $resolvedOutputPath -Force",
		"Release-XlflowComObject -Object $chart",
		"Release-XlflowComObject -Object $chartObject",
		"Release-XlflowComObject -Object $chartObjects",
		"Release-XlflowComObject -Object $range",
		"Release-XlflowComObject -Object $worksheet",
		"Release-XlflowComObject -Object $savedSheet",
	} {
		if !strings.Contains(text, needle) {
			t.Fatalf("expected export-image.ps1 to release %q", needle)
		}
	}
	if strings.Contains(text, "Remove-Item -LiteralPath $resolvedOutputPath -Force") {
		t.Fatalf("expected export-image.ps1 to avoid deleting the destination before export succeeds")
	}
}

func TestFormExportImageScriptUsesTemporaryHelperAndWindowCapture(t *testing.T) {
	data, err := os.ReadFile("form-export-image.ps1")
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		"Install-XlflowVBComponentFromCode",
		"New-XlflowFormExportRuntimeWorkbookCopy",
		"SaveCopyAs($tempPath)",
		"XlflowPrepareFormImageCapture",
		"Invoke-XlflowExcelCallWithDialogWatch",
		`DialogKind "any_vba"`,
		"XlflowFindFormWindowHandle",
		"Resolve-XlflowFormImageCaptureWindow",
		"Wait-XlflowFormImageCaptureWindow",
		"Wait-XlflowStableWindowCaptureInfo",
		"[XlflowNativeMethods]::PrintWindow",
		"CopyFromScreen",
		"$paddingRight = 0",
		"$paddingBottom = 0",
		"runtime_form_loads_initialize",
		"runtime_form_temp_copy",
		"userform_image_export_experimental",
		"temporary_component_cleanup_failed",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-export-image.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestFormExportImageScriptPrefersPrintWindowBeforeScreenFallback(t *testing.T) {
	data, err := os.ReadFile("form-export-image.ps1")
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	copyIndex := strings.Index(text, "$graphics.CopyFromScreen")
	printIndex := strings.Index(text, "[XlflowNativeMethods]::PrintWindow")
	if copyIndex == -1 || printIndex == -1 {
		t.Fatalf("expected both CopyFromScreen and PrintWindow in form-export-image.ps1:\n%s", text)
	}
	if printIndex > copyIndex {
		t.Fatalf("expected PrintWindow to be attempted before CopyFromScreen fallback:\n%s", text)
	}
	for _, want := range []string{
		`$printOk = $false`,
		`[XlflowNativeMethods]::PrintWindow([IntPtr]$Hwnd, $hdc, 2)`,
		`if (-not $printOk) {`,
		`$printOk = [XlflowNativeMethods]::PrintWindow([IntPtr]$Hwnd, $hdc, 0)`,
		`PrintWindow failed; falling back to CopyFromScreen`,
		`CopyFromScreen fallback failed after PrintWindow`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-export-image.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestFormExportImageScriptRejectsXLMAINWindowFallback(t *testing.T) {
	data, err := os.ReadFile("form-export-image.ps1")
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`function Test-XlflowLikelyUserFormWindow`,
		`[string]::Equals($className, "XLMAIN", [System.StringComparison]::OrdinalIgnoreCase)`,
		`$className -match "(?i)^Thunder"`,
		`Test-XlflowLikelyUserFormWindow -WindowInfo $window`,
		`Test-XlflowLikelyUserFormWindow -WindowInfo $stableWindow`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-export-image.ps1 missing XLMAIN guard %q:\n%s", want, text)
		}
	}
}

func TestFormExportImageScriptReportsChosenCaptureWindowMetadata(t *testing.T) {
	data, err := os.ReadFile("form-export-image.ps1")
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`$result.target.capture_window = [ordered]@{`,
		`class_name = $window.class_name`,
		`width = $window.width`,
		`height = $window.height`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-export-image.ps1 missing capture-window metadata %q:\n%s", want, text)
		}
	}
}

func TestFormExportImageScriptMapsPrepareFailuresToValidationCodes(t *testing.T) {
	data, err := os.ReadFile("form-export-image.ps1")
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`XlflowFormImageCapture.capture_prepare`,
		`return "runtime_form_load_failed"`,
		`return "vba_compile_failed"`,
		`Set-XlflowFormExportDialogFailure`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-export-image.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestFormExportImageScriptNormalizesOverwriteTempParentForCurrentDirectory(t *testing.T) {
	data, err := os.ReadFile("form-export-image.ps1")
	if err != nil {
		t.Fatalf("failed to read form-export-image.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`$tempParent = $outputParent`,
		`if ([string]::IsNullOrWhiteSpace($tempParent))`,
		`$tempParent = (Get-Location).ProviderPath`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("form-export-image.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestAnyVbaDialogWatcherDoesNotFallbackToFirstButtonForCompileDialogs(t *testing.T) {
	data, err := os.ReadFile("common.ps1")
	if err != nil {
		t.Fatalf("failed to read common.ps1: %v", err)
	}
	text := string(data)
	for _, want := range []string{
		`param(`,
		`[string]$MatchedKind = ""`,
		`if ($MatchedKind -eq "compile")`,
		`Test-XlflowAllowDialogFirstButtonFallback -DialogKind $DialogKind -MatchedKind $matchedKind`,
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("common.ps1 missing %q:\n%s", want, text)
		}
	}
}

func TestInvokeXlflowExcelCallWithDialogWatchUsesShortPostInvokeWait(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; "+
			"$script:waitMs = -1; "+
			"function Get-XlflowExcelProcessId { param($Excel) return 123 }; "+
			"function Start-XlflowExcelDialogWatcher { param([int]$ProcessId, [string]$Kind = 'compile', [int]$TimeoutMilliseconds = 10000, [int]$PollMilliseconds = 50) return [pscustomobject]@{ powershell = $null; async = $null } }; "+
			"function Receive-XlflowExcelDialogWatcher { param($Watcher, [int]$WaitMilliseconds = 250) $script:waitMs = $WaitMilliseconds; return (New-XlflowExcelDialogWatcherResult) }; "+
			"$r = Invoke-XlflowExcelCallWithDialogWatch -Excel ([pscustomobject]@{}) -Workbook $null -Invocation { 'ok' }; "+
			"[pscustomobject]@{ wait = $script:waitMs; value = $r.value } | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("invoke dialog watch command failed: %v\n%s", err, out)
	}
	var got struct {
		Wait  int    `json:"wait"`
		Value string `json:"value"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse invoke dialog watch output: %v\n%s", err, out)
	}
	if got.Wait != 250 {
		t.Fatalf("wait = %d, want 250", got.Wait)
	}
	if got.Value != "ok" {
		t.Fatalf("value = %q, want ok", got.Value)
	}
}

func TestCommonScriptCompileDialogSafetyHelpers(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; "+
			"$result = [pscustomobject]@{ "+
			"compileSignal = (Test-XlflowCompileDialogSignals -Title 'Microsoft Visual Basic' -StaticText \"Compile error:`nExpected: expression\" -ButtonText 'OK'); "+
			"saveSignal = (Test-XlflowCompileDialogSignals -Title 'Microsoft Excel' -StaticText 'Do you want to save the changes?' -ButtonText \"Yes`nNo`nCancel\"); "+
			"compileFallback = (Test-XlflowAllowDialogFirstButtonFallback -DialogKind 'compile'); "+
			"runtimeFallback = (Test-XlflowAllowDialogFirstButtonFallback -DialogKind 'runtime') "+
			"}; $result | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("compile dialog helper command failed: %v\n%s", err, out)
	}
	var got struct {
		CompileSignal   bool `json:"compileSignal"`
		SaveSignal      bool `json:"saveSignal"`
		CompileFallback bool `json:"compileFallback"`
		RuntimeFallback bool `json:"runtimeFallback"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse compile helper output: %v\n%s", err, out)
	}
	if !got.CompileSignal {
		t.Fatalf("expected compile-specific dialog text to be detected, got %+v", got)
	}
	if got.SaveSignal {
		t.Fatalf("expected generic Excel save dialog text to be ignored, got %+v", got)
	}
	if got.CompileFallback {
		t.Fatalf("compile watcher should not use first-button fallback, got %+v", got)
	}
	if !got.RuntimeFallback {
		t.Fatalf("runtime watcher should keep first-button fallback, got %+v", got)
	}
}

func TestRunScriptRejectsDirectDiagnosticBeforeOpeningWorkbook(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./run.ps1 -WorkbookPath 'C:\\missing.xlsm' -MacroName 'Main.Run' -Direct true -Diagnostic true | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run direct diagnostic command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code  string `json:"code"`
			Phase string `json:"phase"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse run output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code != "run_args_invalid" || got.Error.Phase != "initialize" {
		t.Fatalf("expected direct diagnostic argument failure, got %+v", got)
	}
}

func TestRunScriptAllowsDirectWhenDiagnosticFalse(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$r = ./run.ps1 -WorkbookPath 'C:\\missing.xlsm' -MacroName 'Main.Run' -MacroArgsJson 'W10=' -Direct true -Diagnostic false | ConvertFrom-Json; $r | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run direct diagnostic=false command failed: %v\n%s", err, out)
	}
	var got struct {
		Error *struct {
			Code  string `json:"code"`
			Phase string `json:"phase"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse run output: %v\n%s", err, out)
	}
	if got.Error == nil || got.Error.Code == "run_args_invalid" {
		t.Fatalf("expected direct run to pass diagnostic=false validation, got %+v", got)
	}
}

func TestPushScriptScopesSaveSessionWarningToSessionRuns(t *testing.T) {
	body, err := os.ReadFile(filepath.Join(".", "push.ps1"))
	if err != nil {
		t.Fatal(err)
	}
	text := string(body)
	if !strings.Contains(text, "Open-XlflowWorkbookForCommand") {
		t.Fatalf("push.ps1 should use the shared workbook-open helper for session reuse:\n%s", text)
	}
	if !strings.Contains(text, "\"SAVE REQUIRED: live workbook is newer than disk; run xlflow save before session stop\"") {
		t.Fatalf("push.ps1 should emit the strengthened save-required guidance:\n%s", text)
	}
	if !strings.Contains(text, "\"left workbook unchanged on disk\"") {
		t.Fatalf("push.ps1 should preserve the non-session unchanged-disk log:\n%s", text)
	}
	if !strings.Contains(text, "Get-XlflowSourceUserFormNames") {
		t.Fatalf("push.ps1 should inspect source UserForms during Phase 1 warnings:\n%s", text)
	}
	if !strings.Contains(text, "Add-XlflowUserFormSessionStaleWarning") {
		t.Fatalf("push.ps1 should emit the UserForm stale-session warning for no-save pushes:\n%s", text)
	}
}

func TestPullScriptClearsSourcesOnlyAfterWorkbookOpen(t *testing.T) {
	body, err := os.ReadFile(filepath.Join(".", "pull.ps1"))
	if err != nil {
		t.Fatal(err)
	}
	text := string(body)
	openIdx := strings.Index(text, "Open-XlflowWorkbookForCommand")
	clearIdx := strings.Index(text, "Clear-XlflowSourceComponentFiles")
	if openIdx < 0 || clearIdx < 0 {
		t.Fatalf("expected pull.ps1 to call workbook open and source cleanup:\n%s", text)
	}
	if clearIdx < openIdx {
		t.Fatalf("pull.ps1 clears exported sources before opening the workbook, which can destroy the source tree on open failure:\n%s", text)
	}
}

func TestPullScriptTreatsUserFormInspectionAsBestEffort(t *testing.T) {
	body, err := os.ReadFile(filepath.Join(".", "pull.ps1"))
	if err != nil {
		t.Fatal(err)
	}
	text := string(body)
	if !strings.Contains(text, "failed to inspect UserForms during pull") {
		t.Fatalf("pull.ps1 should swallow auxiliary UserForm inspection failures:\n%s", text)
	}
	if !strings.Contains(text, "$userFormNames = @(Get-XlflowUserFormNames -Workbook $workbook)") {
		t.Fatalf("pull.ps1 should still collect UserForm names when available:\n%s", text)
	}
}

func TestRunScriptAllowsFastDirectWhenDiagnosticFalse(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; (ConvertTo-XlflowBool 'true') -and (ConvertTo-XlflowBool 'false')",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("bool expression command failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "False" {
		t.Fatalf("expected explicit bool expression to be false, got %q", out)
	}
}

func TestVBESelectionDiagnosticHandlesMissingPane(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; Get-XlflowVBESelectionDiagnostic -VBE ([pscustomobject]@{}) | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("selection diagnostic command failed: %v\n%s", err, out)
	}
	var got struct {
		Location struct {
			Line int `json:"line"`
		} `json:"location"`
		NearbyCode []string `json:"nearby_code"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse selection diagnostic output: %v\n%s", err, out)
	}
	if got.Location.Line != 0 || len(got.NearbyCode) != 0 {
		t.Fatalf("expected empty selection diagnostic, got %+v", got)
	}
}

func TestExcelDialogMessageLinesPreferDialogText(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $dialog = [pscustomobject]@{ title = 'Microsoft Visual Basic'; text = @('Run-time error ''438'':', 'Object does not support this property or method.'); buttons = @(); children = @() }; Get-XlflowExcelDialogMessageLines -Dialog $dialog | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("dialog message extraction failed: %v\n%s", err, out)
	}
	var got []string
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse dialog message lines: %v\n%s", err, out)
	}
	if len(got) != 2 || got[0] != "Run-time error '438':" {
		t.Fatalf("unexpected dialog message lines: %+v", got)
	}
}

func TestVBARuntimeDialogErrorNumberRecognizesLocalizedText(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $dialog = [pscustomobject]@{ title = 'Microsoft Visual Basic'; text = @('実行時エラー ''438'':', 'オブジェクトは、このプロパティまたはメソッドをサポートしていません。') }; Get-XlflowVBARuntimeDialogErrorNumber -Dialog $dialog",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("runtime dialog error number parsing failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "438" {
		t.Fatalf("expected runtime dialog error number 438, got %q", out)
	}
}

func TestAttachActiveWithoutWorkbookReturnsStructuredFailure(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"./attach.ps1 -WorkbookPath 'C:\\missing.xlsm' -Active true | ConvertFrom-Json | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("attach active command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Error  *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse attach output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Error == nil || got.Error.Code == "" {
		t.Fatalf("expected structured attach failure, got %+v", got)
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

func TestMacroProcedureDiscoveryRules(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $body = @('Option Explicit','Public Sub Run()','End Sub','Sub Generate(path As String, count As Long)','End Sub','Public Function Build() As Boolean','End Function','Private Sub Hidden()','End Sub') -join [Environment]::NewLine; Find-XlflowMacroProcedures -ModuleName 'Main' -Code $body | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("macro discovery failed: %v\n%s", err, out)
	}
	var got []struct {
		Module        string   `json:"module"`
		Name          string   `json:"name"`
		QualifiedName string   `json:"qualified_name"`
		Kind          string   `json:"kind"`
		Args          []string `json:"args"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse discovery output: %v\n%s", err, out)
	}
	if len(got) != 3 {
		t.Fatalf("expected 3 discovered macros, got %d: %+v", len(got), got)
	}
	if got[0].QualifiedName != "Main.Run" || got[0].Kind != "sub" {
		t.Fatalf("unexpected first macro: %+v", got[0])
	}
	if got[1].Name != "Generate" || len(got[1].Args) != 2 || got[1].Args[0] != "path As String" {
		t.Fatalf("unexpected argument discovery: %+v", got[1])
	}
	if got[2].Name != "Build" || got[2].Kind != "function" {
		t.Fatalf("unexpected function discovery: %+v", got[2])
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

func TestSetXlflowErrorIncludesPhase(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $result = New-XlflowResult -Command 'run'; Set-XlflowError -Result $result -Code 'macro_failed' -Message 'boom' -Phase 'invoke_macro'; Write-XlflowJson -Result $result",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Set-XlflowError failed: %v\n%s", err, out)
	}
	var got struct {
		Error *struct {
			Phase string `json:"phase"`
		} `json:"error"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse envelope: %v\n%s", err, out)
	}
	if got.Error == nil || got.Error.Phase != "invoke_macro" {
		t.Fatalf("expected phase metadata, got %+v", got)
	}
}

func TestSessionStopSingleLogSerializesAsArray(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $wasDirty = $false; $result = New-XlflowResult -Command 'session'; $result.logs = @(@($(if ($wasDirty) { 'warning: session workbook had unsaved changes before stop' } else { $null }), $(if ($wasDirty) { 'auto-saved workbook while stopping xlflow session; prefer xlflow save before stop' } else { $null }), 'stopped xlflow Excel session') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }); Write-XlflowJson -Result $result",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("session stop log serialization failed: %v\n%s", err, out)
	}
	var got struct {
		Logs []string `json:"logs"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse serialized logs: %v\n%s", err, out)
	}
	if len(got.Logs) != 1 || got.Logs[0] != "stopped xlflow Excel session" {
		t.Fatalf("expected single-item logs array, got %+v", got.Logs)
	}
}

func TestSessionStatusTreatsUnknownDirtyStateAsSaveRequired(t *testing.T) {
	body, err := os.ReadFile(filepath.Join(".", "session.ps1"))
	if err != nil {
		t.Fatalf("read session.ps1: %v", err)
	}
	text := string(body)
	if !strings.Contains(text, "$needsSave = $running -and $open -and (($null -eq $dirty) -or [bool]$dirty)") {
		t.Fatalf("session.ps1 should conservatively treat unknown dirty state as save-required:\n%s", text)
	}
	if !strings.Contains(text, "Get-XlflowUserFormNames -Workbook $workbook") {
		t.Fatalf("session.ps1 should probe workbook UserForms on a best-effort basis:\n%s", text)
	}
	if !strings.Contains(text, "userforms_known = $userFormNamesKnown") {
		t.Fatalf("session.ps1 should distinguish unknown UserForm detection from false:\n%s", text)
	}
	if !strings.Contains(text, "userform_detection_unavailable") {
		t.Fatalf("session.ps1 should warn when UserForm detection is unavailable:\n%s", text)
	}
}

func TestMacrosScriptVBIDEAccessDenialStillIncludesTargetAndSession(t *testing.T) {
	body, err := os.ReadFile(filepath.Join(".", "macros.ps1"))
	if err != nil {
		t.Fatalf("read macros.ps1: %v", err)
	}
	text := string(body)
	for _, want := range []string{
		"$result.target = New-XlflowTargetResult",
		"$result.session = New-XlflowSessionResult",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("macros.ps1 should preserve target/session on VBIDE denial, missing %q:\n%s", want, text)
		}
	}
}

func TestSetXlflowExcelAutomationDefaultsLeavesAutomationSecurityUnchanged(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $excel = [pscustomobject]@{ DisplayAlerts = $true; EnableEvents = $true; AutomationSecurity = 2 }; Set-XlflowExcelAutomationDefaults -Excel $excel -DisplayAlerts $false; [ordered]@{ DisplayAlerts = $excel.DisplayAlerts; EnableEvents = $excel.EnableEvents; AutomationSecurity = $excel.AutomationSecurity } | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Set-XlflowExcelAutomationDefaults failed: %v\n%s", err, out)
	}
	var got struct {
		DisplayAlerts      bool `json:"DisplayAlerts"`
		EnableEvents       bool `json:"EnableEvents"`
		AutomationSecurity int  `json:"AutomationSecurity"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse helper state: %v\n%s", err, out)
	}
	if got.DisplayAlerts || got.EnableEvents || got.AutomationSecurity != 2 {
		t.Fatalf("unexpected helper state: %+v", got)
	}
}

func TestDisableXlflowExcelAutomationMacrosForcesDisable(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $excel = [pscustomobject]@{ AutomationSecurity = 2 }; Disable-XlflowExcelAutomationMacros -Excel $excel; $excel.AutomationSecurity",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Disable-XlflowExcelAutomationMacros failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "3" {
		t.Fatalf("automation security = %q, want 3", out)
	}
}

func TestMacroDisabledFailureDetectionRecognizesJapaneseSecurityMessage(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; Test-XlflowMacroDisabledFailure -Number 1004 -Description 'セキュリティの設定により、マクロが無効になりました。マクロを実行するには、このブックを開き直してマクロを有効にしてください。'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("macro disabled classification failed: %v\n%s", err, out)
	}
	if strings.TrimSpace(string(out)) != "True" {
		t.Fatalf("expected localized macro-disabled detection, got %q", out)
	}
}

func TestSourceTextEncodingHelpersRoundTripJapaneseViaCp932(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $sourcePath = Join-Path $root 'Main.bas'
  $importPath = Join-Path $root 'import\Main.bas'
  $text = 'Public Sub Run()' + [Environment]::NewLine + '  MsgBox "処理が完了しました"' + [Environment]::NewLine + 'End Sub'
  Set-XlflowUtf8Text -Path $sourcePath -Text $text
  Copy-XlflowSourceForImport -SourcePath $sourcePath -DestinationPath $importPath
  $roundtrip = Get-XlflowCp932Text -Path $importPath
  $cp932Base64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($importPath))
  $utf8Base64 = [Convert]::ToBase64String((Get-XlflowUtf8Encoding).GetBytes($text))
  [ordered]@{
    roundtrip = $roundtrip
    cp932DiffersFromUtf8 = $cp932Base64 -ne $utf8Base64
  } | ConvertTo-Json -Compress
} finally {
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("encoding helper round trip failed: %v\n%s", err, out)
	}
	var got struct {
		Roundtrip            string `json:"roundtrip"`
		Cp932DiffersFromUTF8 bool   `json:"cp932DiffersFromUtf8"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse encoding helper output: %v\n%s", err, out)
	}
	if !strings.Contains(got.Roundtrip, "処理が完了しました") {
		t.Fatalf("expected Japanese text to survive CP932 round trip: %q", got.Roundtrip)
	}
	if !got.Cp932DiffersFromUTF8 {
		t.Fatalf("expected import file bytes to be CP932, not UTF-8: %s", out)
	}
}

func TestXlflowFileHashDoesNotDependOnGetFileHashCmdlet(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$tmp = New-TemporaryFile
try {
  [System.IO.File]::WriteAllText($tmp, 'abc', [System.Text.Encoding]::ASCII)
  function Get-FileHash { throw 'Get-FileHash should not be called' }
  Get-XlflowFileHash -Path $tmp
} finally {
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("file hash helper failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	const want = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
	if got != want {
		t.Fatalf("hash = %q, want %q", got, want)
	}
}

func TestCopyXlflowSourceForImportPreservesFrxBytes(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $sourcePath = Join-Path $root 'UserForm1.frx'
  $importPath = Join-Path $root 'import\UserForm1.frx'
  $bytes = [byte[]](0, 255, 130, 160, 13, 10)
  [System.IO.File]::WriteAllBytes($sourcePath, $bytes)
  Copy-XlflowSourceForImport -SourcePath $sourcePath -DestinationPath $importPath
  [ordered]@{
    source = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($sourcePath))
    copied = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($importPath))
  } | ConvertTo-Json -Compress
} finally {
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("frx copy failed: %v\n%s", err, out)
	}
	var got struct {
		Source string `json:"source"`
		Copied string `json:"copied"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse frx copy output: %v\n%s", err, out)
	}
	if got.Source != got.Copied {
		t.Fatalf("expected .frx bytes to be copied unchanged, got source=%q copied=%q", got.Source, got.Copied)
	}
}

func TestGetXlflowSourceUserFormNamesFindsRecursiveFrmFiles(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  $formsDir = Join-Path $root 'src\forms'
  $nestedDir = Join-Path $formsDir 'Nested'
  New-Item -ItemType Directory -Force -Path $nestedDir | Out-Null
  Set-Content -LiteralPath (Join-Path $formsDir 'UserForm1.frm') -Value 'VERSION 5.00' -Encoding UTF8
  Set-Content -LiteralPath (Join-Path $nestedDir 'UserForm2.frm') -Value 'VERSION 5.00' -Encoding UTF8
  [ordered]@{ names = @(Get-XlflowSourceUserFormNames -FormsDir $formsDir) } | ConvertTo-Json -Compress
} finally {
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("Get-XlflowSourceUserFormNames failed: %v\n%s", err, out)
	}
	var got struct {
		Names []string `json:"names"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse source UserForm names: %v\n%s", err, out)
	}
	sort.Strings(got.Names)
	if want := []string{"UserForm1", "UserForm2"}; !reflect.DeepEqual(got.Names, want) {
		t.Fatalf("names = %#v, want %#v", got.Names, want)
	}
}

func TestAddXlflowUserFormMessagesAddsDiscoveryAndStaleWarnings(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$result = New-XlflowResult -Command 'push'
Add-XlflowUserFormDiscoveryMessages -Result $result -Names @('CustomerForm', 'OrderForm')
Add-XlflowUserFormSessionStaleWarning -Result $result -Names @('CustomerForm', 'OrderForm')
$result | ConvertTo-Json -Depth 6 -Compress`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("userform message helpers failed: %v\n%s", err, out)
	}
	var got struct {
		Warnings []struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"warnings"`
		Hints []struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"hints"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse userform message helper output: %v\n%s", err, out)
	}
	if len(got.Warnings) != 2 {
		t.Fatalf("warnings = %#v", got.Warnings)
	}
	if got.Warnings[0].Code != "userform_state_partial" || !strings.Contains(got.Warnings[0].Message, "CustomerForm, OrderForm") {
		t.Fatalf("discovery warning = %#v", got.Warnings[0])
	}
	if got.Warnings[1].Code != "userform_unsaved_session_state" {
		t.Fatalf("stale warning = %#v", got.Warnings[1])
	}
	if len(got.Hints) != 1 || got.Hints[0].Code != "userform_planned_commands" {
		t.Fatalf("hints = %#v", got.Hints)
	}
}

func TestDocumentModuleContentNormalization(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; . ./common.ps1; Set-XlflowUtf8Text -Path $tmp -Text (@('Attribute VB_Name = \"ThisWorkbook\"','Attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"','Option Explicit','Private Sub Workbook_Open()','  MsgBox \"\"起動しました\"\"','End Sub') -join [Environment]::NewLine); $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
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
	if !strings.Contains(got, "Option Explicit") || !strings.Contains(got, "Workbook_Open") || !strings.Contains(got, "起動しました") {
		t.Fatalf("expected VBA body to remain: %q", got)
	}
}

func TestDocumentModuleContentAddsOptionExplicitForEmptyDocumentModule(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; . ./common.ps1; Set-XlflowUtf8Text -Path $tmp -Text (@('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','END','Attribute VB_Name = \"Sheet1\"','Attribute VB_GlobalNameSpace = False','Attribute VB_Creatable = False','Attribute VB_PredeclaredId = True','Attribute VB_Exposed = True') -join [Environment]::NewLine); $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
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

func TestGetXlflowComponentPathUsesFolderAnnotation(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`. ./common.ps1
$codeModule = New-Object PSObject -Property @{ CountOfLines = 2 }
$scriptBlock = { param($start, $count) "'@Folder(""Domain.Services"")" + [Environment]::NewLine + 'Option Explicit' }
$codeModule | Add-Member -MemberType ScriptMethod -Name Lines -Value $scriptBlock
$component = [pscustomobject]@{ Name = 'StockService'; Type = 1; CodeModule = $codeModule }
Get-XlflowComponentPath -Component $component -ModulesDir 'src/modules' -ClassesDir 'src/classes' -FormsDir 'src/forms' -WorkbookDir 'src/workbook'`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("component path mapping with folder annotation failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(strings.ReplaceAll(string(out), "\\", "/"))
	if !strings.HasSuffix(got, "src/modules/Domain/Services/StockService.bas") {
		t.Fatalf("expected nested path, got %q", got)
	}
}

func TestCopyXlflowSourceForImportUpdatesFolderAnnotationFromPath(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  $sourceDir = Join-Path $root 'src\modules\Domain\Services'
  $importPath = Join-Path $root 'import\StockService.bas'
  New-Item -ItemType Directory -Force -Path $sourceDir | Out-Null
  $sourcePath = Join-Path $sourceDir 'StockService.bas'
  Set-XlflowUtf8Text -Path $sourcePath -Text (@('Attribute VB_Name = "StockService"', '''@Folder("Legacy")', 'Option Explicit') -join [Environment]::NewLine)
  Copy-XlflowSourceForImport -SourcePath $sourcePath -DestinationPath $importPath -RootDir (Join-Path $root 'src\modules') -FolderAnnotationMode 'update'
  Get-XlflowCp932Text -Path $importPath
} finally {
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("folder annotation import update failed: %v\n%s", err, out)
	}
	got := strings.ReplaceAll(string(out), "\r\n", "\n")
	if !strings.Contains(got, `'@Folder("Domain.Services")`) {
		t.Fatalf("expected updated folder annotation, got %q", got)
	}
	if strings.Contains(got, `'@Folder("Legacy")`) {
		t.Fatalf("expected legacy annotation to be replaced, got %q", got)
	}
}

func TestGetXlflowFolderAnnotationForPathRejectsOutOfRootPaths(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
try {
  Get-XlflowFolderAnnotationForPath -RootDir 'C:\repo\src\modules' -Path 'C:\repo\helpers\Util.bas'
  throw 'expected out-of-root path to fail'
} catch {
  $_.Exception.Message
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("out-of-root folder annotation check failed: %v\n%s", err, out)
	}
	if !strings.Contains(string(out), "resolves outside root") {
		t.Fatalf("expected out-of-root failure message, got %q", out)
	}
}

func TestFindXlflowDuplicateModuleNamesDetectsRecursiveConflicts(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  $modulesDir = Join-Path $root 'src\modules'
  New-Item -ItemType Directory -Force -Path (Join-Path $modulesDir 'Domain'), (Join-Path $modulesDir 'Infrastructure') | Out-Null
  Set-XlflowUtf8Text -Path (Join-Path $modulesDir 'Domain\User.bas') -Text 'Attribute VB_Name = "User"'
  Set-XlflowUtf8Text -Path (Join-Path $modulesDir 'Infrastructure\User.bas') -Text 'Attribute VB_Name = "User"'
  $files = Get-XlflowSourceComponentFiles -ModulesDir $modulesDir -ClassesDir '' -FormsDir '' -WorkbookDir ''
  Find-XlflowDuplicateModuleNames -Files $files | ConvertTo-Json -Compress
} finally {
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("duplicate detection failed: %v\n%s", err, out)
	}
	type duplicateResult struct {
		ModuleName string   `json:"module_name"`
		Paths      []string `json:"paths"`
	}
	var got []duplicateResult
	if err := json.Unmarshal(out, &got); err != nil {
		var single duplicateResult
		if singleErr := json.Unmarshal(out, &single); singleErr != nil {
			t.Fatalf("failed to parse duplicate detection output: %v\n%s", err, out)
		}
		got = []duplicateResult{single}
	}
	if len(got) != 1 || got[0].ModuleName != "User" || len(got[0].Paths) != 2 {
		t.Fatalf("unexpected duplicate detection result: %+v", got)
	}
}

func TestDocumentModuleContentKeepsExecutableEndStatement(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; . ./common.ps1; Set-XlflowUtf8Text -Path $tmp -Text (@('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','END','Attribute VB_Name = \"ThisWorkbook\"','Option Explicit','Public Sub StopAll()','  End','End Sub') -join [Environment]::NewLine); $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
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
		"$tmp = New-TemporaryFile; . ./common.ps1; Set-XlflowUtf8Text -Path $tmp -Text (@('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','  Persistable = 0  ''NotPersistable','END','Attribute VB_Name = \"ThisWorkbook\"','Option Explicit','Public Sub Hello()','End Sub') -join [Environment]::NewLine); $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
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
		"$tmp = New-TemporaryFile; . ./common.ps1; Set-XlflowUtf8Text -Path $tmp -Text (@('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','Option Explicit','Public Sub Recover()','End Sub') -join [Environment]::NewLine); $content = Get-XlflowDocumentModuleContent -Path $tmp; Remove-Item -LiteralPath $tmp -Force; $content",
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
		"$tmp = New-TemporaryFile; . ./common.ps1; Set-XlflowUtf8Text -Path $tmp -Text (@('VERSION 1.0 CLASS','BEGIN','  MultiUse = -1  ''True','END','Attribute VB_Name = \"Sheet1\"','Attribute VB_GlobalNameSpace = False','Attribute VB_Creatable = False','Attribute VB_PredeclaredId = True','Attribute VB_Exposed = True') -join [Environment]::NewLine); Normalize-XlflowDocumentModuleFile -Path $tmp; Get-XlflowUtf8Text -Path $tmp; Remove-Item -LiteralPath $tmp -Force",
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

func TestRunArgumentConversionSupportsExplicitTypes(t *testing.T) {
	// Base64-encode the JSON since that's what the function now expects
	json := `[{"type":"string","value":"hello"},{"type":"int","value":"7"},{"type":"bool","value":"true"}]`
	json64 := base64.StdEncoding.EncodeToString([]byte(json))
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		fmt.Sprintf(". ./common.ps1; $json64 = '%s'; $values = ConvertFrom-XlflowRunArgumentsJson -Json $json64; ConvertTo-Json -InputObject $values -Compress", json64),
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run argument conversion failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if got != "[\"hello\",7,true]" {
		t.Fatalf("converted values = %s", got)
	}
}

func TestRunArgumentConversionSupportsExplicitTypesInWindowsPowerShell(t *testing.T) {
	if _, err := exec.LookPath("powershell"); err != nil {
		t.Skip("powershell is not available in this environment")
	}
	json := `[{"type":"string","value":"hello"},{"type":"int","value":"7"},{"type":"bool","value":"true"}]`
	json64 := base64.StdEncoding.EncodeToString([]byte(json))
	cmd := exec.Command(
		"powershell",
		"-NoProfile",
		"-Command",
		fmt.Sprintf(". ./common.ps1; $json64 = '%s'; $values = ConvertFrom-XlflowRunArgumentsJson -Json $json64; ConvertTo-Json -InputObject $values -Compress", json64),
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run argument conversion failed in powershell: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if got != "[\"hello\",7,true]" {
		t.Fatalf("converted values = %s", got)
	}
}

func TestRunHarnessCodeIncludesMacroInvocationAndErrorLine(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $args = @([pscustomobject]@{ type = 'string'; value = 'fixtures\\sample.xlsx' }, [pscustomobject]@{ type = 'int'; value = '3' }, [pscustomobject]@{ type = 'bool'; value = 'true' }); New-XlflowRunHarnessCode -MacroName 'Report.Generate' -Arguments $args",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run harness code generation failed: %v\n%s", err, out)
	}
	got := string(out)
	for _, want := range []string{"Dim targetMacro As String", `targetMacro = "'" & ThisWorkbook.Name & "'!" & "Report.Generate"`, "Application.Run targetMacro, \"fixtures\\sample.xlsx\", CLng(3), CBool(True)", "\"fixtures\\sample.xlsx\"", "CLng(3)", "CBool(True)", "Err.Description", "Erl"} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected run harness code to contain %q:\n%s", want, got)
		}
	}
}

func TestTraceModuleCodeProvidesPublicLoggerAPI(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; New-XlflowTraceModuleCode",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("trace module code generation failed: %v\n%s", err, out)
	}
	got := string(out)
	for _, want := range []string{"Public Sub XlflowLog(ByVal message As String)", "Public Sub XlflowSetTraceFile(ByVal path As String)", "On Error GoTo Handler", "Open mTraceFile For Append", "If opened Then Close #f", "Err.Raise errNumber, errSource, errDescription", "Format$(Now"} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected trace module code to contain %q:\n%s", want, got)
		}
	}
}

func TestWriteTraceModuleSourceWritesUtf8BasFile(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  $path = Write-XlflowTraceModuleSource -ModulesDir $root
  [ordered]@{
    path = $path
    content = Get-XlflowUtf8Text -Path $path
    bom = ([System.IO.File]::ReadAllBytes($path)[0] -eq 239)
  } | ConvertTo-Json -Compress
} finally {
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("trace module source write failed: %v\n%s", err, out)
	}
	var got struct {
		Path    string `json:"path"`
		Content string `json:"content"`
		BOM     bool   `json:"bom"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse source write output: %v\n%s", err, out)
	}
	if !strings.HasSuffix(strings.ReplaceAll(got.Path, "\\", "/"), "/XlflowTrace.bas") {
		t.Fatalf("unexpected trace source path: %q", got.Path)
	}
	if !strings.Contains(got.Content, `Attribute VB_Name = "XlflowTrace"`) || !strings.Contains(got.Content, "Public Sub XlflowLog") {
		t.Fatalf("unexpected trace source content: %q", got.Content)
	}
	if got.BOM {
		t.Fatal("expected UTF-8 without BOM")
	}
}

func TestTraceModuleSourceMatchDetectsModifiedHelper(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $path = Write-XlflowTraceModuleSource -ModulesDir $root
  $before = Test-XlflowTraceModuleSourceMatches -ModulesDir $root
  Add-Content -LiteralPath $path -Value "' user edit"
  $after = Test-XlflowTraceModuleSourceMatches -ModulesDir $root
  [ordered]@{ before = $before; after = $after } | ConvertTo-Json -Compress
} finally {
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
}`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("trace source match check failed: %v\n%s", err, out)
	}
	var got struct {
		Before bool `json:"before"`
		After  bool `json:"after"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse trace source match output: %v\n%s", err, out)
	}
	if !got.Before || got.After {
		t.Fatalf("expected bundled source to match before modification only: %+v", got)
	}
}

func TestTraceInjectThenPushPreservesTraceModule(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$result = [ordered]@{
  skip = $false
  skipReason = ''
  traceStatus = ''
  pushStatus = ''
  sourceExists = $false
  traceStillInjected = $false
}
$excel = $null
$workbook = $null
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $wbPath = Join-Path $root 'trace-persist.xlsm'
  $modulesDir = Join-Path $root 'src/modules'
  $classesDir = Join-Path $root 'src/classes'
  $formsDir = Join-Path $root 'src/forms'
  $workbookDir = Join-Path $root 'src/workbook'
  $backupRoot = Join-Path $root '.xlflow/backups'

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
    $result.skipReason = 'VBProject access is unavailable: ' + $_.Exception.Message
    $result | ConvertTo-Json -Compress
    exit 0
  }
  $workbook.SaveAs($wbPath, 52)
  $global:XlflowSessionExcel = $excel
  $global:XlflowSessionWorkbook = $workbook

  $trace = & ./trace.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir -Visible false -UseSession true | ConvertFrom-Json
  $result.traceStatus = $trace.status
  $sourcePath = Join-Path $modulesDir 'XlflowTrace.bas'
  $result.sourceExists = Test-Path -LiteralPath $sourcePath
  if ($trace.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $push = & ./push.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir -ClassesDir $classesDir -FormsDir $formsDir -WorkbookDir $workbookDir -BackupRoot $backupRoot -Visible false -UseSession true | ConvertFrom-Json
  $result.pushStatus = $push.status
  if ($push.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $excel = Get-XlflowActiveExcel
  $workbook = Get-XlflowOpenWorkbook -Excel $excel -WorkbookPath $wbPath
  $result.traceStillInjected = Test-XlflowTraceModuleInjected -VBProject $workbook.VBProject
  $result | ConvertTo-Json -Compress
} catch {
  $result.skip = $false
  $result.skipReason = ''
  $result.error = $_.Exception.Message
  $result | ConvertTo-Json -Compress
  exit 1
} finally {
  $global:XlflowSessionWorkbook = $null
  $global:XlflowSessionExcel = $null
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
		t.Fatalf("trace persistence check failed: %v\n%s", err, out)
	}
	var got struct {
		Skip               bool   `json:"skip"`
		SkipReason         string `json:"skipReason"`
		TraceStatus        string `json:"traceStatus"`
		PushStatus         string `json:"pushStatus"`
		SourceExists       bool   `json:"sourceExists"`
		TraceStillInjected bool   `json:"traceStillInjected"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse trace persistence output: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skipf("skipped: %s", got.SkipReason)
	}
	if got.TraceStatus != "ok" || !got.SourceExists {
		t.Fatalf("expected trace inject to create source, got %+v output=%s", got, out)
	}
	if got.PushStatus != "ok" {
		t.Fatalf("expected push to succeed, got %+v output=%s", got, out)
	}
	if !got.TraceStillInjected {
		t.Fatalf("expected XlflowTrace to remain after push, got %+v output=%s", got, out)
	}
}

func TestTraceEnableAutoAttachesToMatchingSessionWorkbook(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
. ./common.ps1
$result = [ordered]@{
  skip = $false
  skipReason = ''
  enableStatus = ''
  statusWorkbookInjected = $false
  statusSession = $false
  sourceExists = $false
}
$excel = $null
$workbook = $null
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $wbPath = Join-Path $root 'trace-session.xlsm'
  $modulesDir = Join-Path $root 'src/modules'
  $metadataPath = Join-Path $root '.xlflow/session.json'
  New-Item -ItemType Directory -Force -Path $modulesDir | Out-Null
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $metadataPath) | Out-Null

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
    $result.skipReason = 'VBProject access is unavailable: ' + $_.Exception.Message
    $result | ConvertTo-Json -Compress
    exit 0
  }
  $workbook.SaveAs($wbPath, 52)
  $processId = Get-XlflowExcelProcessId -Excel $excel
  $hwnd = [int64]$excel.Hwnd
  [ordered]@{
    pid = $processId
    hwnd = $hwnd
    workbook_path = [System.IO.Path]::GetFullPath($wbPath)
    port = 0
    token = [guid]::NewGuid().ToString('N')
    started_at = (Get-Date).ToString('o')
  } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $metadataPath -Encoding UTF8

  $enable = & ./trace.ps1 -Action enable -WorkbookPath $wbPath -ModulesDir $modulesDir -MetadataPath $metadataPath | ConvertFrom-Json
  $status = & ./trace.ps1 -Action status -WorkbookPath $wbPath -ModulesDir $modulesDir -MetadataPath $metadataPath | ConvertFrom-Json
  $result.enableStatus = $enable.status
  $result.statusWorkbookInjected = [bool]$status.trace.workbook_injected
  $result.statusSession = [bool]$status.workbook.session
  $result.sourceExists = Test-Path -LiteralPath (Join-Path $modulesDir 'XlflowTrace.bas')
  $result | ConvertTo-Json -Compress
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
		t.Fatalf("trace auto-session attach failed: %v\n%s", err, out)
	}
	var got struct {
		Skip                   bool   `json:"skip"`
		SkipReason             string `json:"skipReason"`
		EnableStatus           string `json:"enableStatus"`
		StatusWorkbookInjected bool   `json:"statusWorkbookInjected"`
		StatusSession          bool   `json:"statusSession"`
		SourceExists           bool   `json:"sourceExists"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse trace auto-session attach output: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skipf("skipped: %s", got.SkipReason)
	}
	if got.EnableStatus != "ok" || !got.StatusWorkbookInjected || !got.StatusSession || !got.SourceExists {
		t.Fatalf("expected trace command to auto-attach to matching session workbook, got %+v output=%s", got, out)
	}
}

func TestRunHarnessCodeConfiguresTraceBeforeMacro(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; New-XlflowRunHarnessCode -MacroName 'Report.Generate' -Arguments @() -TraceEnabled $true -TraceFile 'C:\\Temp\\xlflow\\trace.log'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run harness trace generation failed: %v\n%s", err, out)
	}
	got := string(out)
	setup := `XlflowTrace.XlflowSetTraceFile "C:\Temp\xlflow\trace.log"`
	invocation := "Application.Run targetMacro"
	if !strings.Contains(got, setup) {
		t.Fatalf("expected trace setup %q:\n%s", setup, got)
	}
	if strings.Index(got, setup) > strings.Index(got, invocation) {
		t.Fatalf("expected trace setup before macro invocation:\n%s", got)
	}
}

func TestReadTraceEventsParsesTimestampMessageAndRawLine(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$tmp = New-TemporaryFile; Set-Content -LiteralPath $tmp -Value @(\"2026-04-29 21:12:03`tstart GenerateReport\",\"2026-04-29 21:12:04`tlastRow=128\"); . ./common.ps1; $events = @(Read-XlflowTraceEvents -Path $tmp); Remove-Item -LiteralPath $tmp -Force; ConvertTo-Json -InputObject $events -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("trace event parsing failed: %v\n%s", err, out)
	}
	var got []struct {
		Timestamp string `json:"timestamp"`
		Message   string `json:"message"`
		Raw       string `json:"raw"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse trace events: %v\n%s", err, out)
	}
	if len(got) != 2 {
		t.Fatalf("expected 2 trace events, got %d: %+v", len(got), got)
	}
	if got[0].Timestamp != "2026-04-29 21:12:03" || got[0].Message != "start GenerateReport" || got[0].Raw != "2026-04-29 21:12:03\tstart GenerateReport" {
		t.Fatalf("unexpected first trace event: %+v", got[0])
	}
}

func TestRunTraceFailureWithNoEventsIncludesHint(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		"$missing = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.xlsm'); ./run.ps1 -WorkbookPath $missing -MacroName 'Main.Run' -TraceEnabled true | ConvertFrom-Json | ConvertTo-Json -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run trace failure command failed: %v\n%s", err, out)
	}
	var got struct {
		Status string `json:"status"`
		Trace  *struct {
			Events []any  `json:"events"`
			Hint   string `json:"hint"`
		} `json:"trace"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse run output: %v\n%s", err, out)
	}
	if got.Status != "failed" || got.Trace == nil || len(got.Trace.Events) != 0 || !strings.Contains(got.Trace.Hint, "no trace events") {
		t.Fatalf("expected empty trace hint, got %+v", got)
	}
}

func TestRunTraceBlankWorkbookReturnsMacroNotFound(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
$result = [ordered]@{
  skip = $false
  skipReason = ''
  status = ''
  errorCode = ''
  phase = ''
}
$excel = $null
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
  } catch {
    $result.skip = $true
    $result.skipReason = 'Excel COM is unavailable: ' + $_.Exception.Message
    $result | ConvertTo-Json -Compress
    exit 0
  }
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $wbPath = Join-Path $root 'blank.xlsm'
  ./new.ps1 -WorkbookPath $wbPath | Out-Null
  $run = ./run.ps1 -WorkbookPath $wbPath -MacroName 'Main.Run' -MacroArgsJson 'W10=' -TraceEnabled true | ConvertFrom-Json
  $result.status = $run.status
  $result.errorCode = $run.error.code
  $result.phase = $run.error.phase
} finally {
  if ($null -ne $excel) {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
  if (Test-Path -LiteralPath $root) {
    Remove-Item -LiteralPath $root -Recurse -Force
  }
  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()
}
$result | ConvertTo-Json -Compress`,
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("blank workbook run command failed: %v\n%s", err, out)
	}
	var got struct {
		Skip       bool   `json:"skip"`
		SkipReason string `json:"skipReason"`
		Status     string `json:"status"`
		ErrorCode  string `json:"errorCode"`
		Phase      string `json:"phase"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse blank workbook run output: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skip(got.SkipReason)
	}
	if got.Status != "failed" || got.ErrorCode != "macro_not_found" || got.Phase != "verify_macro" {
		t.Fatalf("expected macro_not_found during verify_macro, got %+v", got)
	}
}

func TestUIButtonAddListRemoveEndToEnd(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
$result = [ordered]@{
  skip = $false
  skipReason = ''
  addStatus = ''
  updateStatus = ''
  listStatus = ''
  removeStatus = ''
  finalListStatus = ''
  buttonCountAfterUpdate = 0
  buttonCountAfterRemove = 0
  updated = $false
  text = ''
  macro = ''
}
$excel = $null
$workbook = $null
$sessionStarted = $false
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $wbPath = Join-Path $root 'ui-button.xlsm'

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
    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = 'Main'
    $module.CodeModule.AddFromString(('Option Explicit', 'Public Sub Run()', 'End Sub') -join [Environment]::NewLine)
  } catch {
    $result.skip = $true
    $result.skipReason = 'VBProject access is unavailable: ' + $_.Exception.Message
    $result | ConvertTo-Json -Compress
    exit 0
  }
  $workbook.SaveAs($wbPath, 52)
  $workbook.Close($true) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  $workbook = $null
  $excel.Quit() | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  $excel = $null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  $add = & ./ui.ps1 -Action add -WorkbookPath $wbPath -Sheet 'Menu' -Cell 'B2' -Text 'Run' -Macro 'Main.Run' -Id 'run' -CreateSheet true -VerifyMacro true | ConvertFrom-Json
  $result.addStatus = $add.status
  if ($add.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $update = & ./ui.ps1 -Action add -WorkbookPath $wbPath -Sheet 'Menu' -Cell 'B3' -Text 'Run Updated' -Macro 'Main.Run' -Id 'run' -VerifyMacro true | ConvertFrom-Json
  $result.updateStatus = $update.status
  $result.updated = $update.ui.button.updated
  if ($update.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $list = & ./ui.ps1 -Action list -WorkbookPath $wbPath -Sheet 'Menu' | ConvertFrom-Json
  $result.listStatus = $list.status
  $buttons = @($list.ui.buttons)
  $result.buttonCountAfterUpdate = $buttons.Count
  if ($buttons.Count -gt 0) {
    $result.text = $buttons[0].text
    $result.macro = $buttons[0].macro
  }

  $remove = & ./ui.ps1 -Action remove -WorkbookPath $wbPath -Sheet 'Menu' -Id 'run' | ConvertFrom-Json
  $result.removeStatus = $remove.status
  if ($remove.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $finalList = & ./ui.ps1 -Action list -WorkbookPath $wbPath -Sheet 'Menu' | ConvertFrom-Json
  $result.finalListStatus = $finalList.status
  $result.buttonCountAfterRemove = @($finalList.ui.buttons).Count
  $result | ConvertTo-Json -Compress
} catch {
  $result.error = $_.Exception.Message
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
		t.Fatalf("ui button e2e failed: %v\n%s", err, out)
	}
	var got struct {
		Skip                   bool   `json:"skip"`
		SkipReason             string `json:"skipReason"`
		AddStatus              string `json:"addStatus"`
		UpdateStatus           string `json:"updateStatus"`
		ListStatus             string `json:"listStatus"`
		RemoveStatus           string `json:"removeStatus"`
		FinalListStatus        string `json:"finalListStatus"`
		ButtonCountAfterUpdate int    `json:"buttonCountAfterUpdate"`
		ButtonCountAfterRemove int    `json:"buttonCountAfterRemove"`
		Updated                bool   `json:"updated"`
		Text                   string `json:"text"`
		Macro                  string `json:"macro"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse ui button e2e output: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skipf("skipped: %s", got.SkipReason)
	}
	if got.AddStatus != "ok" || got.UpdateStatus != "ok" || got.ListStatus != "ok" || got.RemoveStatus != "ok" || got.FinalListStatus != "ok" {
		t.Fatalf("unexpected ui statuses: %+v output=%s", got, out)
	}
	if !got.Updated {
		t.Fatalf("expected second add to update existing button: %+v", got)
	}
	if got.ButtonCountAfterUpdate != 1 {
		t.Fatalf("expected exactly one button after idempotent update, got %+v", got)
	}
	if got.Text != "Run Updated" || got.Macro != "Main.Run" {
		t.Fatalf("unexpected button metadata after update: %+v", got)
	}
	if got.ButtonCountAfterRemove != 0 {
		t.Fatalf("expected no buttons after remove, got %+v", got)
	}
}

func TestRunHarnessCodeAcceptsDecodedJSONArgumentArrays(t *testing.T) {
	json := `[{"type":"string","value":"fixtures\\sample.xlsx"},{"type":"int","value":"3"},{"type":"bool","value":"true"}]`
	json64 := base64.StdEncoding.EncodeToString([]byte(json))
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		fmt.Sprintf(". ./common.ps1; $json64 = '%s'; $decodedJson = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($json64)); $args = ConvertFrom-Json -InputObject $decodedJson; New-XlflowRunHarnessCode -MacroName 'Report.Generate' -Arguments $args", json64),
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run harness code generation from decoded JSON failed: %v\n%s", err, out)
	}
	got := string(out)
	for _, want := range []string{`targetMacro = "'" & ThisWorkbook.Name & "'!" & "Report.Generate"`, "Application.Run targetMacro, \"fixtures\\sample.xlsx\", CLng(3), CBool(True)", "\"fixtures\\sample.xlsx\"", "CLng(3)", "CBool(True)"} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected run harness code to contain %q:\n%s", want, got)
		}
	}
}

func TestRunHarnessCodeEscapesEmbeddedQuotesInStringArguments(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $args = @([pscustomobject]@{ type = 'string'; value = 'say \"hi\"' }); New-XlflowRunHarnessCode -MacroName 'M.Sub' -Arguments $args",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run harness code generation failed: %v\n%s", err, out)
	}
	got := string(out)
	want := `"say ""hi"""`
	if !strings.Contains(got, want) {
		t.Fatalf("expected run harness code to contain escaped string literal %q:\n%s", want, got)
	}
}

func TestSaveAsExtensionValidationRejectsMismatchedTargets(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; try { Assert-XlflowSaveAsExtension -WorkbookPath 'build\\Book.xlsm' -SaveAsPath 'build\\Book.xlsx'; 'unexpected success' } catch { $_.Exception.Message }",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("save-as validation command failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if !strings.Contains(got, "does not match workbook extension") {
		t.Fatalf("validation output = %q", got)
	}
}

func TestFormatMacroFailureMessageIncludesLineAndErrNumber(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; Format-XlflowMacroFailureMessage -ModuleName 'Main' -Line 10 -Number 5 -Description 'inputPath is required'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("macro failure message formatting failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if got != "Main line 10 Err 5: inputPath is required" {
		t.Fatalf("failure message = %q", got)
	}
}

func TestRunHarnessModuleNameFitsVBAModuleLimit(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $name = New-XlflowRunHarnessModuleName; Write-Output $name; Write-Output $name.Length",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run harness module name generation failed: %v\n%s", err, out)
	}
	lines := strings.Fields(strings.TrimSpace(string(out)))
	if len(lines) != 2 {
		t.Fatalf("unexpected output: %q", out)
	}
	if !strings.HasPrefix(lines[0], "XlflowRun_") {
		t.Fatalf("module name = %q", lines[0])
	}
	if lines[1] != "30" {
		t.Fatalf("module name length = %q, want 30", lines[1])
	}
}

func TestFormatMacroFailureMessageDescriptionOnlyNoLeadingColon(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; Format-XlflowMacroFailureMessage -ModuleName '' -Line 0 -Number 0 -Description 'inputPath is required'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("macro failure message formatting failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if got != "inputPath is required" {
		t.Fatalf("failure message = %q, expected no leading colon", got)
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
  $global:XlflowSessionExcel = $excel
  $global:XlflowSessionWorkbook = $workbook

  $pull1 = & ./pull.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir1 -ClassesDir $classesDir1 -FormsDir $formsDir1 -WorkbookDir $workbookDir1 -Visible false -UseSession true | ConvertFrom-Json
  $result.initialPullStatus = $pull1.status
  if ($pull1.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $push = & ./push.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir1 -ClassesDir $classesDir1 -FormsDir $formsDir1 -WorkbookDir $workbookDir1 -BackupRoot $backupRoot -Visible false -UseSession true | ConvertFrom-Json
  $result.pushStatus = $push.status
  if ($push.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $pull2 = & ./pull.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir2 -ClassesDir $classesDir2 -FormsDir $formsDir2 -WorkbookDir $workbookDir2 -Visible false -UseSession true | ConvertFrom-Json
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
  $global:XlflowSessionWorkbook = $null
  $global:XlflowSessionExcel = $null
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
		Skip                bool   `json:"skip"`
		SkipReason          string `json:"skipReason"`
		FrmPath             string `json:"frmPath"`
		FrxPath             string `json:"frxPath"`
		FrxExists           bool   `json:"frxExists"`
		FrmReferencesFrx    bool   `json:"frmReferencesFrx"`
		FrxIsSibling        bool   `json:"frxIsSibling"`
		InitialPullStatus   string `json:"initialPullStatus"`
		PushStatus          string `json:"pushStatus"`
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

func TestUserFormCodeSidecarRoundTripInSidecarMode(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		`$ErrorActionPreference = 'Stop'
$result = [ordered]@{
  skip = $false
  skipReason = ''
  initialPullStatus = ''
  pushStatus = ''
  roundtripPullStatus = ''
  initialSidecarHasA = $false
  finalSidecarHasB = $false
  finalFrmHasB = $false
}
$excel = $null
$workbook = $null
$component = $null
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
  $wbPath = Join-Path $root 'userform-sidecar-roundtrip.xlsm'
  $modulesDir1 = Join-Path $root 'src1/modules'
  $classesDir1 = Join-Path $root 'src1/classes'
  $formsDir1 = Join-Path $root 'src1/forms'
  $workbookDir1 = Join-Path $root 'src1/workbook'
  $modulesDir2 = Join-Path $root 'src2/modules'
  $classesDir2 = Join-Path $root 'src2/classes'
  $formsDir2 = Join-Path $root 'src2/forms'
  $workbookDir2 = Join-Path $root 'src2/workbook'
  $backupRoot = Join-Path $root 'backups'
  $formName = 'UserFormSidecar'

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
  $component.CodeModule.AddFromString(@'
Option Explicit

Private versionTag As String

Private Sub UserForm_Initialize()
    versionTag = "A"
End Sub
'@)
  $workbook.SaveAs($wbPath, 52)
  $global:XlflowSessionExcel = $excel
  $global:XlflowSessionWorkbook = $workbook

  $pull1 = & ./pull.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir1 -ClassesDir $classesDir1 -FormsDir $formsDir1 -WorkbookDir $workbookDir1 -CodeSource sidecar -Visible false -UseSession true | ConvertFrom-Json
  $result.initialPullStatus = $pull1.status
  if ($pull1.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $sidecar1 = Join-Path $formsDir1 'code\UserFormSidecar.bas'
  if (Test-Path -LiteralPath $sidecar1) {
    $result.initialSidecarHasA = ((Get-Content -Raw -LiteralPath $sidecar1) -like '*versionTag = "A"*')
  }

  @'
Option Explicit

Private versionTag As String

Private Sub UserForm_Initialize()
    versionTag = "B"
End Sub
'@ | Set-Content -LiteralPath $sidecar1 -Encoding UTF8

  $push = & ./push.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir1 -ClassesDir $classesDir1 -FormsDir $formsDir1 -WorkbookDir $workbookDir1 -CodeSource sidecar -BackupRoot $backupRoot -Visible false -UseSession true | ConvertFrom-Json
  $result.pushStatus = $push.status
  if ($push.status -ne 'ok') {
    $result | ConvertTo-Json -Compress
    exit 0
  }

  $pull2 = & ./pull.ps1 -WorkbookPath $wbPath -ModulesDir $modulesDir2 -ClassesDir $classesDir2 -FormsDir $formsDir2 -WorkbookDir $workbookDir2 -CodeSource sidecar -Visible false -UseSession true | ConvertFrom-Json
  $result.roundtripPullStatus = $pull2.status
  if ($pull2.status -eq 'ok') {
    $sidecar2 = Join-Path $formsDir2 'code\UserFormSidecar.bas'
    $frm2 = Join-Path $formsDir2 'UserFormSidecar.frm'
    if (Test-Path -LiteralPath $sidecar2) {
      $result.finalSidecarHasB = ((Get-Content -Raw -LiteralPath $sidecar2) -like '*versionTag = "B"*')
    }
    if (Test-Path -LiteralPath $frm2) {
      $result.finalFrmHasB = ((Get-Content -Raw -LiteralPath $frm2) -like '*versionTag = "B"*')
    }
  }

  $result | ConvertTo-Json -Compress
} catch {
  $result.skip = $false
  $result.skipReason = ''
  $result.error = $_.Exception.Message
  $result | ConvertTo-Json -Compress
  exit 1
} finally {
  $global:XlflowSessionWorkbook = $null
  $global:XlflowSessionExcel = $null
  if ($null -ne $component) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($component) | Out-Null } catch {}
  }
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
		t.Fatalf("userform code sidecar roundtrip failed: %v\n%s", err, out)
	}

	var got struct {
		Skip                bool   `json:"skip"`
		SkipReason          string `json:"skipReason"`
		InitialPullStatus   string `json:"initialPullStatus"`
		PushStatus          string `json:"pushStatus"`
		RoundtripPullStatus string `json:"roundtripPullStatus"`
		InitialSidecarHasA  bool   `json:"initialSidecarHasA"`
		FinalSidecarHasB    bool   `json:"finalSidecarHasB"`
		FinalFrmHasB        bool   `json:"finalFrmHasB"`
	}
	if err := json.Unmarshal(out, &got); err != nil {
		t.Fatalf("failed to parse sidecar roundtrip output as json: %v\n%s", err, out)
	}
	if got.Skip {
		t.Skipf("skipped: %s", got.SkipReason)
	}
	if got.InitialPullStatus != "ok" || got.PushStatus != "ok" || got.RoundtripPullStatus != "ok" {
		t.Fatalf("unexpected sidecar roundtrip statuses: %+v output=%s", got, out)
	}
	if !got.InitialSidecarHasA {
		t.Fatalf("expected initial sidecar export to capture code-behind A: %+v", got)
	}
	if !got.FinalSidecarHasB || !got.FinalFrmHasB {
		t.Fatalf("expected sidecar push/pull roundtrip to preserve B code-behind: %+v", got)
	}
}
