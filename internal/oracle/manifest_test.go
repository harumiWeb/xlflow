package oracle

import (
	"encoding/json"
	"os"
	"path/filepath"
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
	caseJSON := Case{SchemaVersion: SchemaVersion, ID: "sample", Modules: []Module{{Name: "Main", Kind: "standard", Path: "Main.bas", Entry: true}}, Probe: Probe{Mode: ProbeCompile}, VBE: VBEExpectation{Expected: expected, EvidencePhase: EvidenceUnknown, DiagnosticMeaning: MeaningObservation}, Provenance: Provenance{Status: "pending"}}
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
	c := Case{SchemaVersion: SchemaVersion, ID: "x", Modules: []Module{{Name: "Main", Kind: "standard", Path: "Main.bas"}}, Probe: Probe{Mode: ProbeCompile}, VBE: VBEExpectation{Expected: ExpectedAccepted, EvidencePhase: EvidenceCompile}}
	if err := ValidateCase(c, "x", t.TempDir()); err == nil {
		t.Fatal("expected provenance validation error")
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
