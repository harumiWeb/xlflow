package oracle

import (
	"context"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/typedb"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

// The committed fixture contract is deliberately Excel-free. This test keeps
// the same sources usable by batch lint/analyze and the realtime/LSP analyzer,
// while the VBE observations remain an optional local developer step.
func TestCommittedFixtureContractsWithoutExcel(t *testing.T) {
	typeDBDir := t.TempDir()
	const typeDBOutput = "oracle.generated.json"
	if err := os.WriteFile(filepath.Join(typeDBDir, typeDBOutput), []byte(`{"types":[]}`), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := typedb.WriteManifest(typeDBDir, typedb.Manifest{
		GeneratorVersion: "test",
		Libraries:        []typedb.ManifestLibrary{{Name: "Oracle", Output: typeDBOutput}},
	}); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, typeDBDir)

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
			modulePath := func(module Module) string {
				root := cfg.Src.Modules
				switch module.Kind {
				case "class":
					root = cfg.Src.Classes
				case "form":
					root = filepath.Join(cfg.Src.Forms, "code")
				case "document":
					root = cfg.Src.Workbook
				}
				return filepath.Join(projectRoot, root, module.Name+".bas")
			}
			projections := map[string][]Diagnostic{}
			for _, module := range c.Modules {
				path := modulePath(module)
				if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
					t.Fatal(err)
				}
				if err := os.WriteFile(path, sources[module.Name], 0o644); err != nil {
					t.Fatal(err)
				}
			}

			lintResult, err := (lint.Linter{RootDir: projectRoot, Config: cfg}).RunResult()
			if err != nil {
				t.Fatal(err)
			}
			for _, issue := range lintResult.Issues {
				startColumn := issue.Column
				if startColumn < 1 {
					startColumn = 1
				}
				endLine := issue.EndLine
				if endLine < 1 {
					endLine = issue.Line
				}
				endColumn := issue.EndColumn
				if endColumn < 1 {
					endColumn = startColumn + 1
				}
				projections["lint"] = append(projections["lint"], Diagnostic{
					Code: issue.Code, Severity: issue.Severity,
					Range: &Range{StartLine: issue.Line, StartColumn: startColumn, EndLine: endLine, EndColumn: endColumn},
				})
			}

			if analysisUsesSurface(c.Analysis, "analyze") {
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
			}

			lspAnalyzer := intel.Analyzer{RootDir: projectRoot, Config: cfg, DB: db}
			// The direct intel analyzer normally runs file-local diagnostics. For
			// this Excel-free contract test, install the same project resolver
			// projection used by LSP Full so interprocedural VB052–VB054 contracts
			// are checked on the LSP surface as well.
			if analysisUsesSurface(c.Analysis, "analyze") {
				resolverSymbols := make([]procedureir.ResolverSymbol, 0)
				for _, module := range c.Modules {
					ir, buildErr := procedureir.BuildSource(procedureir.BuildOptions{
						RootDir: projectRoot, Path: modulePath(module), ModuleName: module.Name, ModuleKind: module.Kind,
					}, sources[module.Name])
					if buildErr != nil {
						t.Fatal(buildErr)
					}
					for _, declaration := range ir.Declarations {
						resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
							Name: declaration.Name, Type: declaration.Type, Module: ir.ModuleName, ModuleKind: ir.ModuleKind,
							Kind: declaration.Kind, Visibility: declaration.Visibility, File: ir.Path, Line: declaration.Range.StartLine,
							Parent: declaration.Parent, IsArray: declaration.IsArray, IsConst: declaration.IsConst,
							ValueShape: declaration.ValueShape,
						})
					}
					for _, procedure := range ir.Procedures {
						resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
							Name: procedure.Symbol.Name, Type: procedure.Symbol.ReturnType, Module: ir.ModuleName, ModuleKind: ir.ModuleKind,
							Kind: string(procedure.Symbol.Kind), Visibility: procedure.Symbol.Visibility, File: ir.Path,
							Line:    procedure.Symbol.DeclarationRange.StartLine,
							IsArray: procedure.Symbol.IsArray, ValueShape: procedure.Symbol.ValueShape,
						})
					}
				}
				projectResolver := procedureir.NewResolverWithCompleteness(resolverSymbols, true)
				lspAnalyzer.RealtimeFindingsFunc = func(_ context.Context, _ string, _ config.Config, _ *vbaast.ParsedDocument, ir procedureir.DocumentIR, _ vbacfg.Document) ([]intel.RealtimeFinding, error) {
					resolved := procedureir.Resolve(ir, projectResolver)
					findings := make([]intel.RealtimeFinding, 0)
					for _, diagnostic := range procedureir.Diagnostics(resolved, true) {
						findings = append(findings, intel.RealtimeFinding{
							Code: diagnostic.Code, Severity: "error", Line: diagnostic.Range.StartLine,
							Column: diagnostic.Range.StartColumn, EndLine: diagnostic.Range.EndLine,
							EndColumn: diagnostic.Range.EndColumn, Message: diagnostic.Message,
						})
					}
					return findings, nil
				}
			}
			for _, module := range c.Modules {
				path := modulePath(module)
				for _, diagnostic := range lspAnalyzer.DiagnosticsContext(context.Background(), intel.Document{
					URI: path, Path: path, Source: string(sources[module.Name]), ModuleKind: module.Kind,
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

func analysisUsesSurface(analysis AnalysisExpectation, surface string) bool {
	expectations := append(append([]DiagnosticExpectation(nil), analysis.ExpectedDiagnostics...), analysis.ForbiddenDiagnostics...)
	for _, expectation := range expectations {
		if containsString(effectiveSurfaces(expectation.Code, expectation.Surfaces), surface) {
			return true
		}
	}
	return false
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
	if report.AssertedFixtures != 103 || report.BoundFixtures != 87 || report.PartialFixtures != 0 || report.UnboundFixtures != 15 || report.NotApplicable != 2 {
		t.Fatalf("unexpected current corpus coverage: %+v", report)
	}
	assertIDs := func(name string, got, want []string) {
		t.Helper()
		if !reflect.DeepEqual(got, want) {
			t.Fatalf("%s fixture IDs = %v, want %v", name, got, want)
		}
	}
	assertIDs("bound", report.BoundIDs, []string{
		"byref-bare-variable",
		"byref-compatible",
		"byref-incompatible",
		"byref-literal-temporaries",
		"const-assignment-accepted",
		"const-assignment-rejected",
		"declaration-keyword-valid-controls",
		"declaration-redim-after-comma",
		"declaration-repeated-dim-after-comma",
		"duplicate-declaration",
		"duplicate-declaration-valid-controls",
		"duplicate-named-argument",
		"fixed-array-reversed-bound",
		"fixed-array-valid-bound",
		"invalid-declaration-placement",
		"invalid-declaration-placement-valid-controls",
		"invalid-event-standard",
		"invalid-friend-standard",
		"invalid-implements-standard",
		"invalid-me-standard",
		"invalid-public-object-member",
		"invalid-withevents-shape",
		"invalid-withevents-standard",
		"issue590-ambiguous-enum",
		"issue590-ambiguous-enum-valid",
		"issue590-missing-project-procedure",
		"issue590-missing-project-procedure-valid",
		"issue590-non-callable",
		"issue590-non-callable-valid",
		"issue590-undeclared-raiseevent",
		"issue590-undeclared-raiseevent-valid",
		"issue591-duplicate-label",
		"issue591-duplicate-label-valid",
		"issue591-invalid-exit",
		"issue591-invalid-exit-valid",
		"issue591-next-mismatch",
		"issue591-next-mismatch-valid",
		"issue591-undefined-label",
		"issue591-undefined-label-valid",
		"issue592-accepted-bare",
		"issue592-accepted-call",
		"issue592-accepted-function",
		"issue592-accepted-multi-idiom",
		"issue592-accepted-zero",
		"issue592-call-unparenthesized",
		"issue592-function-bare-expression",
		"issue592-invalid-call-target",
		"issue592-standalone-empty",
		"issue592-standalone-space-multi",
		"issue594-conditional-elseif-after-else",
		"issue594-conditional-missing-then",
		"issue594-conditional-orphan-else",
		"issue594-conditional-orphan-elseif",
		"issue594-conditional-valid",
		"issue594-open-missing-for-mode",
		"issue594-open-valid",
		"issue594-select-case-duplicate-else",
		"issue594-select-case-outside",
		"issue594-select-case-valid",
		"issue594-typeof-trailing-token",
		"issue594-typeof-valid",
		"known-as-type",
		"known-enum-as-type",
		"known-named-argument",
		"missing-required-argument",
		"option-in-procedure",
		"option-in-procedure-valid-controls",
		"optional-argument-omitted",
		"procedure-optional-numeric-literals-valid",
		"procedure-parameter-limit-60",
		"procedure-parameter-limit-61",
		"procedure-signature-invalid",
		"procedure-signature-valid",
		"property-signature-invalid",
		"property-signature-valid",
		"scalar-assignment",
		"set-object-target",
		"set-scalar-target",
		"unknown-as-type",
		"unknown-named-argument",
		"valid-module-class",
		"valid-module-document-workbook",
		"valid-module-document-worksheet",
		"valid-module-form",
		"valid-module-standard",
		"vb009-backslash-doubled-quotes-valid",
		"vb009-c-style-quote-escape",
	})
	assertIDs("partially-bound", report.PartialIDs, nil)
	assertIDs("unbound", report.UnboundIDs, []string{
		"byref-parenthesized-variable",
		"erase-scalar",
		"foreach-scalar",
		"function-bare-call",
		"function-parenthesized-call",
		"lbound-scalar",
		"missing-set-object-assignment",
		"object-to-object-parameter",
		"redim-fixed-array",
		"redim-reversed-bound",
		"redim-scalar",
		"scalar-to-object-parameter",
		"sub-bare-call",
		"sub-parenthesized-call",
		"vb010-conditional-procedure-valid",
	})
	assertIDs("not-applicable", report.NotApplicableIDs, []string{"known-compile-accept", "known-compile-reject"})
}
