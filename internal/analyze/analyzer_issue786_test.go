package analyze

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestIssue786ClassProcedureShadowsExternalStandardProcedure(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "CDPHelpers.bas", `Attribute VB_Name = "CDPHelpers"
Option Explicit

Public Sub printMsg(ByVal level As Long, ByVal text As String, ByVal fromProcedure As String, ByVal logFileExtension As String)
End Sub
`)
	writeClass(t, dir, "PrintMsgOverrideRepro.cls", `VERSION 1.0 CLASS
Attribute VB_Name = "PrintMsgOverrideRepro"
Option Explicit

Private Sub printMsg(ByVal level As Long, ByVal text As String, ByVal fromProcedure As String, Optional ByVal isHeader As Boolean = False)
End Sub

Public Sub CallIt()
    printMsg 1, "hello", "CallIt"
End Sub
`)
	writeModule(t, dir, "ExternalCaller.bas", `Attribute VB_Name = "ExternalCaller"
Option Explicit

Public Sub Run()
    printMsg 1, "hello", "Run"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VB045")
	if len(got) != 1 || !strings.Contains(got[0].File, "ExternalCaller.bas") || !strings.Contains(got[0].Message, "expects at least 4 argument") {
		t.Fatalf("Issue #786 VB045 findings = %+v, want only external fallback mismatch", got)
	}
}
