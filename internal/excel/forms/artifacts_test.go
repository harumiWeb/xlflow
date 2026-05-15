package forms

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestValidateUserFormArtifactsAgainstSpecsRejectsMissingFRM(t *testing.T) {
	formsDir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(formsDir, "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	spec := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: RegistrationForm\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(filepath.Join(formsDir, "specs", "RegistrationForm.yaml"), []byte(spec), 0o644); err != nil {
		t.Fatal(err)
	}

	issues, err := ValidateUserFormArtifactsAgainstSpecs(formsDir, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 1 {
		t.Fatalf("issues = %#v, want 1 missing artifact issue", issues)
	}
	if !strings.Contains(issues[0].Message, "no matching .frm artifact") {
		t.Fatalf("unexpected issue: %+v", issues[0])
	}
}

func TestValidateUserFormArtifactsAgainstSpecsRejectsVBNameMismatch(t *testing.T) {
	formsDir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(formsDir, "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	spec := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: RegistrationForm\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(filepath.Join(formsDir, "specs", "RegistrationForm.yaml"), []byte(spec), 0o644); err != nil {
		t.Fatal(err)
	}
	frm := "VERSION 5.00\nBegin {GUID} RegistrationForm\nEnd\nAttribute VB_Name = \"UserForm1\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(formsDir, "RegistrationForm.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}

	issues, err := ValidateUserFormArtifactsAgainstSpecs(formsDir, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 1 {
		t.Fatalf("issues = %#v, want 1 VB_Name mismatch issue", issues)
	}
	if issues[0].Line != 4 || !strings.Contains(issues[0].Message, "Attribute VB_Name") {
		t.Fatalf("unexpected issue: %+v", issues[0])
	}
}

func TestValidateUserFormArtifactsAgainstSpecsRejectsSpecFilenameMismatch(t *testing.T) {
	formsDir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(formsDir, "specs"), 0o755); err != nil {
		t.Fatal(err)
	}
	spec := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: RegistrationForm\ncontrols: []\nwarnings: []\n"
	if err := os.WriteFile(filepath.Join(formsDir, "specs", "UserForm1.yaml"), []byte(spec), 0o644); err != nil {
		t.Fatal(err)
	}
	frm := "VERSION 5.00\nBegin {GUID} RegistrationForm\nEnd\nAttribute VB_Name = \"RegistrationForm\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(formsDir, "RegistrationForm.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}

	issues, err := ValidateUserFormArtifactsAgainstSpecs(formsDir, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 1 {
		t.Fatalf("issues = %#v, want 1 filename mismatch issue", issues)
	}
	if !strings.Contains(issues[0].Message, "spec file") || !strings.Contains(issues[0].Message, "form.name") {
		t.Fatalf("unexpected issue: %+v", issues[0])
	}
}
