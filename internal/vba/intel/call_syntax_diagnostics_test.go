package intel

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestDiagnosticsExposeVB059FromSharedLint(t *testing.T) {
	analyzer := Analyzer{RootDir: t.TempDir(), Config: config.Default()}
	doc := Document{
		Path: t.TempDir() + "\\Main.bas",
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Probe()
    WritePair (1, 2)
End Sub
Private Sub WritePair(ByVal first As Long, ByVal second As Long)
End Sub
`,
	}
	diagnostics := analyzer.Diagnostics(doc)
	got := diagnosticsByCode(diagnostics, "VB059")
	if len(got) != 1 {
		t.Fatalf("VB059 diagnostics = %#v, want one", diagnostics)
	}
	if got[0].Severity != "error" || got[0].Range.Start.Line != 3 {
		t.Fatalf("VB059 diagnostic = %#v, want error on source line 4 (zero-based 3)", got[0])
	}
	if len(diagnosticsByCode(diagnostics, "VB014")) != 0 {
		t.Fatalf("VB059 syntax evidence should replace generic parser recovery: %#v", diagnostics)
	}
}
