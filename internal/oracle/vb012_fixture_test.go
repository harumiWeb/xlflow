package oracle

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// TestVB012TerminatorFixtureMatrix keeps the committed VB012 audit evidence
// complete. The expected VBE result and the VB012 lint/LSP contract are
// promoted from the sequential local Excel run; parser shape and style
// diagnostics remain separate production tests.
func TestVB012TerminatorFixtureMatrix(t *testing.T) {
	manifestPath := filepath.Join("..", "..", "testdata", "vbe-oracle", "manifest.json")
	manifest, root, err := LoadManifest(manifestPath)
	if err != nil {
		t.Fatal(err)
	}

	type moduleSpec struct {
		suffix string
		kind   string
		name   string
		target string
	}
	openers := map[string]string{
		"sub":          "Public Sub VB012Probe()",
		"function":     "Public Function VB012Probe() As Long",
		"property-get": "Public Property Get VB012Probe() As Long",
		"property-let": "Public Property Let VB012Probe(ByVal value As Long)",
		"property-set": "Public Property Set VB012Probe(ByVal value As Object)",
	}
	ends := []string{"end-sub", "end-function", "end-property"}
	endSource := map[string]string{
		"end-sub":      "End Sub",
		"end-function": "End Function",
		"end-property": "End Property",
	}
	accepted := map[string]bool{
		"function-end-function":     true,
		"property-get-end-function": true,
		"property-get-end-property": true,
		"property-get-end-sub":      true,
		"property-let-end-property": true,
		"property-set-end-property": true,
		"sub-end-sub":               true,
	}
	modules := []moduleSpec{
		{suffix: "standard", kind: "standard", name: "Main"},
		{suffix: "class", kind: "class", name: "VB012Class"},
		{suffix: "document-workbook", kind: "document", name: "ThisWorkbook", target: "workbook"},
		{suffix: "document-worksheet", kind: "document", name: "Sheet1", target: "worksheet"},
	}

	entries := make(map[string]ManifestEntry)
	for _, entry := range manifest.Cases {
		if strings.HasPrefix(entry.ID, "vb012-terminator-") {
			if _, exists := entries[entry.ID]; exists {
				t.Fatalf("duplicate VB012 fixture %q", entry.ID)
			}
			entries[entry.ID] = entry
		}
	}
	const wantCount = 5 * 3 * 4
	if len(entries) != wantCount {
		t.Fatalf("VB012 fixture count = %d, want %d", len(entries), wantCount)
	}

	for opener, declaration := range openers {
		for _, end := range ends {
			for _, module := range modules {
				id := "vb012-terminator-" + opener + "-" + end + "-" + module.suffix
				entry, ok := entries[id]
				if !ok {
					t.Errorf("missing VB012 matrix fixture %q", id)
					continue
				}
				fixture, sources, err := LoadCase(root, entry)
				if err != nil {
					t.Errorf("load %s: %v", id, err)
					continue
				}
				target := opener + "-" + end
				wantExpected := ExpectedRejected
				if accepted[target] {
					wantExpected = ExpectedAccepted
				}
				if fixture.VBE.Expected != wantExpected {
					t.Errorf("%s vbe.expected = %q, want %q", id, fixture.VBE.Expected, wantExpected)
				}
				if fixture.VBE.EvidencePhase != EvidenceCompile || fixture.Provenance.Status != "asserted" || len(fixture.Provenance.VerifiedOn) == 0 {
					t.Errorf("%s lacks promoted compile provenance: vbe=%+v provenance=%+v", id, fixture.VBE, fixture.Provenance)
				}
				if fixture.Analysis.BindingStatus != BindingBound || !containsString(fixture.Analysis.RuleCodes, "VB012") {
					t.Errorf("%s analysis binding = %+v, want bound VB012", id, fixture.Analysis)
				}
				if accepted[target] {
					if len(fixture.Analysis.ForbiddenDiagnostics) != 1 || fixture.Analysis.ForbiddenDiagnostics[0].Code != "VB012" {
						t.Errorf("%s accepted contract = %+v, want VB012 forbidden", id, fixture.Analysis.ForbiddenDiagnostics)
					}
				} else if len(fixture.Analysis.ExpectedDiagnostics) != 1 || fixture.Analysis.ExpectedDiagnostics[0].Code != "VB012" || fixture.Analysis.ExpectedDiagnostics[0].Severity != "error" {
					t.Errorf("%s rejected contract = %+v, want VB012/error", id, fixture.Analysis.ExpectedDiagnostics)
				}
				if len(fixture.Modules) != 1 {
					t.Errorf("%s module count = %d, want 1", id, len(fixture.Modules))
					continue
				}
				gotModule := fixture.Modules[0]
				if gotModule.Kind != module.kind || gotModule.Name != module.name {
					t.Errorf("%s module = (%q, %q), want (%q, %q)", id, gotModule.Kind, gotModule.Name, module.kind, module.name)
				}
				if gotModule.DocumentTarget != module.target {
					t.Errorf("%s document_target = %q, want %q", id, gotModule.DocumentTarget, module.target)
				}
				source := string(sources[gotModule.Name])
				if !strings.Contains(source, declaration) {
					t.Errorf("%s source is missing opener %q", id, declaration)
				}
				if !strings.Contains(source, endSource[end]) {
					t.Errorf("%s source is missing terminator %q", id, endSource[end])
				}

				// Keep the raw tree-sitter recovery result separate from the
				// promoted VBE result. Every non-matching matrix cell currently
				// recovers in the CST; the analyzer parse gate admits only the two
				// VBE-accepted Property Get mismatches through the explicit
				// compatibility exception.
				projectRoot := t.TempDir()
				cfg := config.Default()
				moduleRoot := cfg.Src.Modules
				switch gotModule.Kind {
				case "class":
					moduleRoot = cfg.Src.Classes
				case "document":
					moduleRoot = cfg.Src.Workbook
				}
				path := filepath.Join(projectRoot, moduleRoot, gotModule.Name+".bas")
				if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
					t.Fatal(err)
				}
				if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
					t.Fatal(err)
				}
				ir, err := procedureir.BuildSource(procedureir.BuildOptions{
					RootDir: projectRoot, Path: path, ModuleName: gotModule.Name, ModuleKind: gotModule.Kind,
				}, []byte(source))
				if err != nil {
					t.Fatal(err)
				}
				openerKind := strings.SplitN(opener, "-", 2)[0]
				wantRecovery := openerKind != strings.TrimPrefix(end, "end-")
				gotRecovery := ir.Parse.HasError || ir.Parse.HasMissing
				if gotRecovery != wantRecovery {
					t.Errorf("%s CST recovery = %v, want %v (parse=%+v)", id, gotRecovery, wantRecovery, ir.Parse)
				}

				issues, err := (lint.Linter{RootDir: projectRoot, Config: cfg, ModuleKind: gotModule.Kind}).LintSource(path, []byte(source))
				if err != nil {
					t.Fatal(err)
				}
				boundaryCodes := make(map[string][]lint.Issue)
				for _, issue := range issues {
					switch issue.Code {
					case "VB010", "VB011", "VB012", "VB014", "VB066":
						boundaryCodes[issue.Code] = append(boundaryCodes[issue.Code], issue)
					}
				}
				expectedBoundaryCode := ""
				if wantRecovery {
					if accepted[target] {
						expectedBoundaryCode = "VB066"
					} else {
						expectedBoundaryCode = "VB012"
					}
				}
				for _, code := range []string{"VB010", "VB011", "VB012", "VB014", "VB066"} {
					want := 0
					if code == expectedBoundaryCode {
						want = 1
					}
					if got := len(boundaryCodes[code]); got != want {
						t.Errorf("%s %s count = %d, want %d (issues=%+v)", id, code, got, want, boundaryCodes)
					}
				}
				if expectedBoundaryCode == "VB012" && len(boundaryCodes["VB012"]) > 0 && boundaryCodes["VB012"][0].Severity != "error" {
					t.Errorf("%s VB012 severity = %q, want error", id, boundaryCodes["VB012"][0].Severity)
				}
				if expectedBoundaryCode == "VB066" && len(boundaryCodes["VB066"]) > 0 && boundaryCodes["VB066"][0].Severity != "warning" {
					t.Errorf("%s VB066 severity = %q, want warning", id, boundaryCodes["VB066"][0].Severity)
				}

				_, analyzeErr := (analyze.Analyzer{RootDir: projectRoot, Config: cfg}).RunResult()
				var parseErr *analyze.ParseError
				wantParseError := wantRecovery && !accepted[target]
				if wantParseError {
					if !errors.As(analyzeErr, &parseErr) {
						t.Errorf("%s rejected recovery error = %v, want ParseError", id, analyzeErr)
					}
				} else if analyzeErr != nil {
					t.Errorf("%s accepted/matching analyzer error = %v", id, analyzeErr)
				}
			}
		}
	}
}
