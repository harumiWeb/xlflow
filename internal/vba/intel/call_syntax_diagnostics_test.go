package intel

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestDiagnosticsExposeVB059FromSharedLint(t *testing.T) {
	root := t.TempDir()
	analyzer := Analyzer{RootDir: root, Config: config.Default()}
	doc := Document{
		Path: filepath.Join(root, "Main.bas"),
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

func TestCompileEquivalentDiagnosticsExposeArrayShapeRules(t *testing.T) {
	root := t.TempDir()
	doc := Document{
		Path: filepath.Join(root, "Main.bas"),
		Source: `Option Explicit
Public Const Limit As Long = 2
Public Sub Probe()
    Dim bad(3 To 1) As Long
    Limit = 3
End Sub
`,
	}
	diagnostics := (Analyzer{RootDir: root, Config: config.Default()}).CompileEquivalentDiagnosticsContext(context.Background(), doc)
	if got := diagnosticsByCode(diagnostics, "VB060"); len(got) != 1 || got[0].Severity != "error" {
		t.Fatalf("VB060 diagnostics = %#v", diagnostics)
	}
	if got := diagnosticsByCode(diagnostics, "VB061"); len(got) != 1 || got[0].Severity != "error" {
		t.Fatalf("VB061 diagnostics = %#v", diagnostics)
	}
}

func TestCompileEquivalentDiagnosticsUseVisibleProjectConstants(t *testing.T) {
	root := t.TempDir()
	doc := Document{
		Path: filepath.Join(root, "Main.bas"),
		Source: `Option Explicit
Public Sub Probe()
    ProjectLimit = 3
End Sub
`,
	}
	diagnostics := (Analyzer{RootDir: root, Config: config.Default(), VisibleConstants: map[string]bool{"projectlimit": true}}).CompileEquivalentDiagnosticsContext(context.Background(), doc)
	if got := diagnosticsByCode(diagnostics, "VB060"); len(got) != 1 || got[0].Severity != "error" {
		t.Fatalf("project VB060 diagnostics = %#v", diagnostics)
	}
}
