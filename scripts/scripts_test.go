package scripts_test

import (
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func TestPowerShellScriptsParse(t *testing.T) {
	scripts := []string{"common.ps1", "doctor.ps1", "new.ps1", "pull.ps1", "push.ps1", "run.ps1"}
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
