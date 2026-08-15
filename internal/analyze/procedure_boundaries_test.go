package analyze

import (
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

const acceptedPropertyGetTerminatorSource = `Attribute VB_Name = "Probe"
Option Explicit

Public Property Get VB012Probe() As Long
    VB012Probe = 1
End Sub
`

func TestAnalyzerAcceptsVBEAcceptedPropertyGetTerminatorRecovery(t *testing.T) {
	cases := []struct {
		name string
		rel  string
	}{
		{name: "standard", rel: filepath.Join("src", "modules", "Probe.bas")},
		{name: "class", rel: filepath.Join("src", "classes", "Probe.cls")},
		{name: "workbook document", rel: filepath.Join("src", "workbook", "ThisWorkbook.bas")},
		{name: "worksheet document", rel: filepath.Join("src", "workbook", "Sheet1.bas")},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			path := filepath.Join(dir, tc.rel)
			if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
				t.Fatal(err)
			}
			if err := os.WriteFile(path, []byte(acceptedPropertyGetTerminatorSource), 0o644); err != nil {
				t.Fatal(err)
			}

			_, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
			if err != nil {
				var parseErr *ParseError
				if errors.As(err, &parseErr) {
					t.Fatalf("VBE-accepted Property Get terminator must not fail analyzer parse gate: %v", err)
				}
				t.Fatal(err)
			}
		})
	}
}

func TestAnalyzerRejectsVBEInvalidProcedureTerminator(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Probe.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(`Attribute VB_Name = "Probe"
Option Explicit

Public Function VB012Probe() As Long
    VB012Probe = 1
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}

	_, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	var parseErr *ParseError
	if !errors.As(err, &parseErr) {
		t.Fatalf("VBE-rejected terminator error = %v, want ParseError", err)
	}
}
