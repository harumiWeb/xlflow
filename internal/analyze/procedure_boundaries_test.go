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
		name         string
		rel          string
		relativeRoot bool
	}{
		{name: "standard", rel: filepath.Join("src", "modules", "Probe.bas")},
		{name: "class", rel: filepath.Join("src", "classes", "Probe.cls")},
		{name: "workbook document", rel: filepath.Join("src", "workbook", "ThisWorkbook.bas")},
		{name: "worksheet document", rel: filepath.Join("src", "workbook", "Sheet1.bas")},
		{name: "worksheet document with relative root", rel: filepath.Join("src", "workbook", "Sheet1.bas"), relativeRoot: true},
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

			rootDir := dir
			if tc.relativeRoot {
				cwd, cwdErr := os.Getwd()
				if cwdErr != nil {
					t.Fatal(cwdErr)
				}
				relativeRoot, relErr := filepath.Rel(cwd, dir)
				if relErr != nil {
					t.Fatal(relErr)
				}
				rootDir = relativeRoot
				if filepath.IsAbs(rootDir) {
					t.Fatalf("relative test root = %q, want relative path", rootDir)
				}
			}
			result, err := (Analyzer{RootDir: rootDir, Config: config.Default()}).RunResult()
			if err != nil {
				var parseErr *ParseError
				if errors.As(err, &parseErr) {
					t.Fatalf("VBE-accepted Property Get terminator must not fail analyzer parse gate: %v", err)
				}
				t.Fatal(err)
			}
			if result.AnalyzedFiles != 1 {
				t.Fatalf("analyzed files = %d, want 1 for %s", result.AnalyzedFiles, path)
			}
		})
	}
}

func TestSourceRealtimeClassifiesWorkbookAndWorksheetDocuments(t *testing.T) {
	cfg := config.Default()
	for _, tc := range []struct {
		name string
	}{
		{name: "ThisWorkbook.bas"},
		{name: "Sheet1.bas"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			path := filepath.Join(dir, cfg.Src.Workbook, tc.name)
			kind, err := realtimeModuleKind(dir, cfg, path)
			if err != nil {
				t.Fatal(err)
			}
			if kind != "document" {
				t.Fatalf("realtime module kind for %s = %q, want document", tc.name, kind)
			}
			if _, err := SourceRealtimeFindings(dir, path, cfg, []byte(`Option Explicit
Public Sub Run()
End Sub
`)); err != nil {
				t.Fatalf("realtime analysis for %s: %v", tc.name, err)
			}
		})
	}
}

func TestAnalyzerRejectsAcceptedTerminatorWithMalformedColonStatement(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "src", "modules", "Probe.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	source := `Attribute VB_Name = "Probe"
Option Explicit

Public Property Get VB012Probe() As Long
    VB012Probe = 1
End Sub: Dim
`
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}

	_, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	var parseErr *ParseError
	if !errors.As(err, &parseErr) {
		t.Fatalf("malformed same-line recovery error = %v, want ParseError", err)
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
