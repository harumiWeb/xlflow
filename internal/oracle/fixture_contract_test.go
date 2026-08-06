package oracle

import (
	"context"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

// The committed fixture contract is deliberately Excel-free. This test keeps
// the same sources usable by batch lint/analyze and the realtime/LSP analyzer,
// while the VBE observations remain an optional local developer step.
func TestCommittedFixtureContractsWithoutExcel(t *testing.T) {
	manifestPath := filepath.Join("..", "..", "testdata", "vbe-oracle", "manifest.json")
	manifest, root, err := LoadManifest(manifestPath)
	if err != nil {
		t.Fatal(err)
	}
	corpus := make([]Case, 0, len(manifest.Cases))
	for _, entry := range manifest.Cases {
		c, _, err := LoadCase(root, entry)
		if err != nil {
			t.Fatal(err)
		}
		corpus = append(corpus, c)
	}
	if _, err := ValidateBindingCoverage(corpus); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}
	for _, entry := range manifest.Cases {
		entry := entry
		t.Run(entry.ID, func(t *testing.T) {
			c, sources, err := LoadCase(root, entry)
			if err != nil {
				t.Fatal(err)
			}
			if entry.ID == manifest.Controls.Accept || entry.ID == manifest.Controls.Reject {
				if c.Analysis.BindingStatus != BindingNotApplicable {
					t.Fatalf("control fixture binding_status = %q, want %q", c.Analysis.BindingStatus, BindingNotApplicable)
				}
			}
			projectRoot := t.TempDir()
			modulesRoot := filepath.Join(projectRoot, "src", "modules")
			if err := os.MkdirAll(modulesRoot, 0o755); err != nil {
				t.Fatal(err)
			}
			projections := map[string][]Diagnostic{}
			for _, module := range c.Modules {
				path := filepath.Join(modulesRoot, module.Name+".bas")
				if err := os.WriteFile(path, sources[module.Name], 0o644); err != nil {
					t.Fatal(err)
				}
			}

			lintResult, err := (lint.Linter{RootDir: projectRoot, Config: cfg}).RunResult()
			if err != nil {
				t.Fatal(err)
			}
			for _, issue := range lintResult.Issues {
				projections["lint"] = append(projections["lint"], Diagnostic{
					Code: issue.Code, Severity: issue.Severity,
					Range: &Range{StartLine: issue.Line, StartColumn: issue.Column, EndLine: issue.Line, EndColumn: issue.Column},
				})
			}

			analyzeResult, err := (analyze.Analyzer{RootDir: projectRoot, Config: cfg}).RunResult()
			if err != nil {
				t.Fatal(err)
			}
			for _, finding := range analyzeResult.Findings {
				projections["analyze"] = append(projections["analyze"], Diagnostic{
					Code: finding.Code, Severity: finding.Severity,
					Range: &Range{StartLine: finding.Line, StartColumn: finding.Column, EndLine: finding.EndLine, EndColumn: finding.EndColumn},
				})
			}

			lspAnalyzer := intel.Analyzer{RootDir: projectRoot, Config: cfg, DB: db}
			for _, module := range c.Modules {
				path := filepath.Join(modulesRoot, module.Name+".bas")
				for _, diagnostic := range lspAnalyzer.DiagnosticsContext(context.Background(), intel.Document{
					URI: path, Path: path, Source: string(sources[module.Name]), ModuleKind: "standard",
				}) {
					projections["lsp"] = append(projections["lsp"], Diagnostic{
						Code: diagnostic.Code, Severity: diagnostic.Severity,
						Range: &Range{
							StartLine: diagnostic.Range.Start.Line + 1, StartColumn: diagnostic.Range.Start.Character + 1,
							EndLine: diagnostic.Range.End.Line + 1, EndColumn: diagnostic.Range.End.Character + 1,
						},
					})
				}
			}
			if err := CheckDiagnosticProjections(c.Analysis, projections); err != nil {
				t.Fatal(err)
			}
		})
	}
}

func TestOracleBindingCoverage(t *testing.T) {
	manifestPath := filepath.Join("..", "..", "testdata", "vbe-oracle", "manifest.json")
	manifest, root, err := LoadManifest(manifestPath)
	if err != nil {
		t.Fatal(err)
	}
	corpus := make([]Case, 0, len(manifest.Cases))
	for _, entry := range manifest.Cases {
		c, _, err := LoadCase(root, entry)
		if err != nil {
			t.Fatal(err)
		}
		corpus = append(corpus, c)
	}
	report, err := ValidateBindingCoverage(corpus)
	t.Log("\n" + report.String())
	if err != nil {
		t.Fatal(err)
	}
	if report.AssertedFixtures != 23 || report.UnboundFixtures != 21 || report.NotApplicable != 2 {
		t.Fatalf("unexpected current corpus coverage: %+v", report)
	}
}
