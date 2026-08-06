package oracle

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func writeFixture(t *testing.T, expected string) (string, Manifest) {
	t.Helper()
	root := t.TempDir()
	caseDir := filepath.Join(root, "cases", "sample")
	if err := os.MkdirAll(caseDir, 0o755); err != nil {
		t.Fatal(err)
	}
	source := "Attribute VB_Name = \"Main\"\nPublic Sub Main()\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(caseDir, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	caseJSON := Case{SchemaVersion: SchemaVersion, ID: "sample", Modules: []Module{{Name: "Main", Kind: "standard", Path: "Main.bas", Entry: true}}, Probe: Probe{Mode: ProbeCompile}, VBE: VBEExpectation{Expected: expected, EvidencePhase: EvidenceUnknown, DiagnosticMeaning: MeaningObservation}, Analysis: AnalysisExpectation{BindingStatus: BindingNotApplicable}, Provenance: Provenance{Status: "pending"}}
	body, _ := json.Marshal(caseJSON)
	if err := os.WriteFile(filepath.Join(caseDir, "case.json"), body, 0o644); err != nil {
		t.Fatal(err)
	}
	manifest := Manifest{SchemaVersion: SchemaVersion, Cases: []ManifestEntry{{ID: "sample"}}, Controls: Controls{Accept: "sample", Reject: "sample-reject"}}
	manifest.Cases = append(manifest.Cases, ManifestEntry{ID: "sample-reject", Path: "cases/sample-reject/case.json"})
	rejectDir := filepath.Join(root, "cases", "sample-reject")
	if err := os.MkdirAll(rejectDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(rejectDir, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	reject := caseJSON
	reject.ID = "sample-reject"
	reject.VBE.Expected = ExpectedObserve
	rejectBody, _ := json.Marshal(reject)
	if err := os.WriteFile(filepath.Join(rejectDir, "case.json"), rejectBody, 0o644); err != nil {
		t.Fatal(err)
	}
	strictDir := filepath.Join(root, "cases", "sample-strict")
	if err := os.MkdirAll(strictDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(strictDir, "Main.bas"), []byte(source), 0o644); err != nil {
		t.Fatal(err)
	}
	strict := caseJSON
	strict.ID = "sample-strict"
	strictBody, _ := json.Marshal(strict)
	if err := os.WriteFile(filepath.Join(strictDir, "case.json"), strictBody, 0o644); err != nil {
		t.Fatal(err)
	}
	manifest.Cases = append(manifest.Cases, ManifestEntry{ID: "sample-strict", Path: "cases/sample-strict/case.json"})
	manifestBody, _ := json.Marshal(manifest)
	manifestPath := filepath.Join(root, "manifest.json")
	if err := os.WriteFile(manifestPath, manifestBody, 0o644); err != nil {
		t.Fatal(err)
	}
	return manifestPath, manifest
}

func TestValidateManifestAndCaseContainment(t *testing.T) {
	manifestPath, _ := writeFixture(t, ExpectedObserve)
	manifest, root, err := LoadManifest(manifestPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadCase(root, manifest.Cases[0]); err != nil {
		t.Fatal(err)
	}
	bad := manifest.Cases[0]
	bad.Path = "../case.json"
	if _, _, err := LoadCase(root, bad); err == nil {
		t.Fatal("expected path containment error")
	}
}

func TestValidateManifestRequiresCaseDirectoryAndEntryAgreement(t *testing.T) {
	manifestPath, _ := writeFixture(t, ExpectedObserve)
	manifest, _, err := LoadManifest(manifestPath)
	if err != nil {
		t.Fatal(err)
	}
	manifest.Cases[0].Path = "cases/other/case.json"
	if err := ValidateManifest(manifest); err == nil {
		t.Fatal("expected manifest case path mismatch")
	}
}

func TestValidateCaseRequiresAssertedProvenance(t *testing.T) {
	c := Case{SchemaVersion: SchemaVersion, ID: "x", Modules: []Module{{Name: "Main", Kind: "standard", Path: "Main.bas"}}, Probe: Probe{Mode: ProbeCompile}, VBE: VBEExpectation{Expected: ExpectedAccepted, EvidencePhase: EvidenceCompile}, Analysis: AnalysisExpectation{BindingStatus: BindingUnbound}}
	if err := ValidateCase(c, "x", t.TempDir()); err == nil {
		t.Fatal("expected provenance validation error")
	}
}

func analysisNote(value string) *string {
	return &value
}

func validAssertedCase(expected string) Case {
	meaning := MeaningSpecification
	if expected == ExpectedRejected {
		meaning = MeaningCompileError
	}
	return Case{
		SchemaVersion: SchemaVersion,
		ID:            "x",
		Modules:       []Module{{Name: "Main", Kind: "standard", Path: "Main.bas"}},
		Probe:         Probe{Mode: ProbeCompile},
		VBE:           VBEExpectation{Expected: expected, EvidencePhase: EvidenceCompile, DiagnosticMeaning: meaning},
		Analysis:      AnalysisExpectation{BindingStatus: BindingUnbound},
		Provenance: Provenance{
			Status: "asserted",
			VerifiedOn: []VerificationMetadata{{
				ExcelVersion: "16.0", ExcelBuild: "17932", Bitness: "x64", Locale: "ja-JP", VerifiedAt: "2026-08-06T00:00:00Z",
			}},
		},
	}
}

func TestValidateCaseBindingMetadata(t *testing.T) {
	tests := []struct {
		name    string
		prepare func(*Case)
		wantErr bool
	}{
		{name: "unbound", prepare: func(c *Case) {}},
		{name: "partially-bound", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingPartiallyBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.BindingNote = analysisNote("positive contract is pending")
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
		}},
		{name: "bound rejected", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101", Severity: "warning"}}
		}},
		{name: "bound accepted", prepare: func(c *Case) {
			c.VBE.Expected = ExpectedAccepted
			c.VBE.DiagnosticMeaning = MeaningSpecification
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ForbiddenDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
		}},
		{name: "not-applicable", prepare: func(c *Case) {
			c.VBE.Expected = ExpectedAccepted
			c.VBE.DiagnosticMeaning = MeaningSpecification
			c.Analysis.BindingStatus = BindingNotApplicable
		}},
		{name: "missing status", prepare: func(c *Case) { c.Analysis.BindingStatus = "" }, wantErr: true},
		{name: "invalid status", prepare: func(c *Case) { c.Analysis.BindingStatus = "future" }, wantErr: true},
		{name: "unbound rule code needs note", prepare: func(c *Case) {
			c.Analysis.RuleCodes = []string{"VBA101"}
		}, wantErr: true},
		{name: "partially-bound needs rule code", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingPartiallyBound
			c.Analysis.BindingNote = analysisNote("pending")
		}, wantErr: true},
		{name: "partially-bound needs note", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingPartiallyBound
			c.Analysis.RuleCodes = []string{"VBA101"}
		}, wantErr: true},
		{name: "bound needs rule code", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
		}, wantErr: true},
		{name: "bound code needs contract", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VB002"}}
		}, wantErr: true},
		{name: "rejected bound code cannot be forbidden-only", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VB002"}}
			c.Analysis.ForbiddenDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
		}, wantErr: true},
		{name: "bound rejected needs expected", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
		}, wantErr: true},
		{name: "bound accepted needs forbidden", prepare: func(c *Case) {
			c.VBE.Expected = ExpectedAccepted
			c.VBE.DiagnosticMeaning = MeaningSpecification
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
		}, wantErr: true},
		{name: "accepted bound code cannot be expected-only", prepare: func(c *Case) {
			c.VBE.Expected = ExpectedAccepted
			c.VBE.DiagnosticMeaning = MeaningSpecification
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
			c.Analysis.ForbiddenDiagnostics = []DiagnosticExpectation{{Code: "VB002"}}
		}, wantErr: true},
		{name: "bound observe rejected", prepare: func(c *Case) {
			c.VBE.Expected = ExpectedObserve
			c.VBE.EvidencePhase = EvidenceUnknown
			c.VBE.DiagnosticMeaning = MeaningObservation
			c.Provenance = Provenance{Status: "pending"}
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
		}, wantErr: true},
		{name: "not-applicable rule code", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingNotApplicable
			c.Analysis.RuleCodes = []string{"VBA101"}
		}, wantErr: true},
		{name: "not-applicable expected", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingNotApplicable
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
		}, wantErr: true},
		{name: "not-applicable forbidden", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingNotApplicable
			c.Analysis.ForbiddenDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
		}, wantErr: true},
		{name: "empty rule code", prepare: func(c *Case) {
			c.Analysis.RuleCodes = []string{""}
		}, wantErr: true},
		{name: "padded rule code", prepare: func(c *Case) {
			c.Analysis.RuleCodes = []string{" VBA101"}
		}, wantErr: true},
		{name: "duplicate rule code", prepare: func(c *Case) {
			c.Analysis.RuleCodes = []string{"VBA101", "VBA101"}
		}, wantErr: true},
		{name: "empty binding note", prepare: func(c *Case) {
			c.Analysis.BindingNote = analysisNote("  ")
		}, wantErr: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			c := validAssertedCase(ExpectedRejected)
			tt.prepare(&c)
			err := ValidateCase(c, c.ID, t.TempDir())
			if (err != nil) != tt.wantErr {
				t.Fatalf("ValidateCase() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestValidateCaseRuleRegistryBindings(t *testing.T) {
	tests := []struct {
		name      string
		prepare   func(*Case)
		wantErr   bool
		errSubstr string
	}{
		{name: "canonical registry diagnostic", prepare: func(c *Case) {
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101", Severity: "warning", Surfaces: []string{"analyze"}}}
		}},
		{name: "dynamic severity", prepare: func(c *Case) {
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA214", Severity: "error"}}
		}},
		{name: "unknown diagnostic code", prepare: func(c *Case) {
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA999"}}
		}, wantErr: true, errSubstr: "not in the static-analysis rule registry"},
		{name: "non-canonical diagnostic code", prepare: func(c *Case) {
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "vba101"}}
		}, wantErr: true, errSubstr: "canonical registry ID"},
		{name: "family-incompatible surface", prepare: func(c *Case) {
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101", Surfaces: []string{"lint"}}}
		}, wantErr: true, errSubstr: "unsupported surface"},
		{name: "non-realtime lsp surface", prepare: func(c *Case) {
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101", Surfaces: []string{"lsp"}}}
		}, wantErr: true, errSubstr: "unsupported surface"},
		{name: "unsupported severity", prepare: func(c *Case) {
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101", Severity: "error"}}
		}, wantErr: true, errSubstr: "unsupported severity"},
		{name: "partially-bound rule code needs contract", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingPartiallyBound
			c.Analysis.BindingNote = analysisNote("pending")
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VB002"}}
		}, wantErr: true, errSubstr: "expected or forbidden diagnostic contract"},
		{name: "bound contract needs rule code", prepare: func(c *Case) {
			c.Analysis.BindingStatus = BindingBound
			c.Analysis.RuleCodes = []string{"VBA101"}
			c.Analysis.ExpectedDiagnostics = []DiagnosticExpectation{{Code: "VBA101"}}
			c.Analysis.ForbiddenDiagnostics = []DiagnosticExpectation{{Code: "VB002"}}
		}, wantErr: true, errSubstr: "not declared by analysis.rule_codes"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			c := validAssertedCase(ExpectedRejected)
			tt.prepare(&c)
			err := ValidateCase(c, c.ID, t.TempDir())
			if (err != nil) != tt.wantErr {
				t.Fatalf("ValidateCase() error = %v, wantErr %v", err, tt.wantErr)
			}
			if tt.errSubstr != "" && !strings.Contains(err.Error(), tt.errSubstr) {
				t.Fatalf("ValidateCase() error = %q, want substring %q", err, tt.errSubstr)
			}
		})
	}
}

func TestRepositoryManifestLoadsAllCases(t *testing.T) {
	manifestPath := filepath.Join("..", "..", "testdata", "vbe-oracle", "manifest.json")
	manifest, root, err := LoadManifest(manifestPath)
	if err != nil {
		t.Skipf("repository fixtures unavailable: %v", err)
	}
	for _, entry := range manifest.Cases {
		if _, _, err := LoadCase(root, entry); err != nil {
			t.Fatalf("case %s: %v", entry.ID, err)
		}
	}
}
