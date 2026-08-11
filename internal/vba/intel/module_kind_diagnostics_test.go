package intel

import (
	"context"
	"path/filepath"
	"testing"
)

func TestDiagnosticsUsesExplicitDocumentModuleKind(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := "Option Explicit\nPublic Event Changed()\n"
	for _, tc := range []struct {
		kind string
		want bool
	}{
		{kind: "standard", want: true},
		{kind: "class", want: false},
	} {
		t.Run(tc.kind, func(t *testing.T) {
			diagnostics := analyzer.DiagnosticsContext(context.Background(), Document{
				Path:       filepath.Join(t.TempDir(), "Module.bas"),
				Source:     source,
				ModuleKind: tc.kind,
			})
			got := false
			for _, diagnostic := range diagnostics {
				if diagnostic.Code == "VB050" {
					got = true
				}
			}
			if got != tc.want {
				t.Fatalf("VB050 present=%t, want %t; diagnostics=%+v", got, tc.want, diagnostics)
			}
		})
	}
}
