package intel

import (
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestDiagnosticsExposeVB066ForVBEAcceptedPropertyGetTerminators(t *testing.T) {
	root := t.TempDir()
	analyzer := Analyzer{RootDir: root, Config: config.Default()}
	for _, tc := range []struct {
		name string
		end  string
	}{
		{name: "end sub", end: "End Sub"},
		{name: "end function", end: "End Function"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			doc := Document{
				Path:       filepath.Join(root, tc.name+".bas"),
				ModuleKind: "standard",
				Source:     "Attribute VB_Name = \"Probe\"\nOption Explicit\nPublic Property Get Value() As Long\n    Value = 1\n" + tc.end + "\n",
			}
			diagnostics := analyzer.Diagnostics(doc)
			got := diagnosticsByCode(diagnostics, "VB066")
			if len(got) != 1 || got[0].Severity != "warning" {
				t.Fatalf("VB066 diagnostics = %#v, want one warning", diagnostics)
			}
			if len(diagnosticsByCode(diagnostics, "VB012")) != 0 {
				t.Fatalf("VBE-accepted mismatch must not emit VB012: %#v", diagnostics)
			}
			if len(diagnosticsByCode(diagnostics, "VB014")) != 0 {
				t.Fatalf("VBE-accepted mismatch must not retain VB014 recovery: %#v", diagnostics)
			}
		})
	}
}
