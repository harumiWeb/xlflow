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

func TestCompileEquivalentDiagnosticsIgnoreWritableMemberAssignments(t *testing.T) {
	root := t.TempDir()
	doc := Document{
		Path: filepath.Join(root, "Main.bas"),
		Source: `Option Explicit
Public Const Hidden As Long = 1
Public Sub Probe(ByVal ws As Worksheet)
    ws.Rows(1).Hidden = True
    ws.Columns(1).Hidden = False
    With ws.Rows(1)
        .Hidden = False
    End With
End Sub
`,
	}
	diagnostics := (Analyzer{RootDir: root, Config: config.Default()}).CompileEquivalentDiagnosticsContext(context.Background(), doc)
	if got := diagnosticsByCode(diagnostics, "VB060"); len(got) != 0 {
		t.Fatalf("writable member assignments produced VB060 diagnostics: %#v", got)
	}
}

func TestDiagnosticsExposeIssue594SyntaxRules(t *testing.T) {
	root := t.TempDir()
	cases := []struct {
		name   string
		code   string
		source string
	}{
		{
			name:   "conditional branch",
			code:   "VB062",
			source: "Option Explicit\nPublic Sub Probe()\n    If ready\n        Debug.Print \"x\"\n    End If\nEnd Sub\n",
		},
		{
			name:   "select case",
			code:   "VB063",
			source: "Option Explicit\nPublic Sub Probe()\n    Select Case value\n    Case Else\n    Case Else\n    End Select\nEnd Sub\n",
		},
		{
			name:   "open mode",
			code:   "VB064",
			source: "Option Explicit\nPublic Sub Probe()\n    Open path For As #1\nEnd Sub\n",
		},
		{
			name:   "TypeOf trailing token",
			code:   "VB065",
			source: "Option Explicit\nPublic Sub Probe()\n    Dim value As Object\n    TypeOf value Is Collection Extra\nEnd Sub\n",
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			doc := Document{Path: filepath.Join(root, tc.name+".bas"), Source: tc.source}
			diagnostics := (Analyzer{RootDir: root, Config: config.Default()}).Diagnostics(doc)
			got := diagnosticsByCode(diagnostics, tc.code)
			if len(got) != 1 || got[0].Severity != "error" {
				t.Fatalf("%s diagnostics = %#v", tc.code, diagnostics)
			}
			if len(diagnosticsByCode(diagnostics, "VB014")) != 0 {
				t.Fatalf("%s should suppress overlapping VB014: %#v", tc.code, diagnostics)
			}
		})
	}
}
