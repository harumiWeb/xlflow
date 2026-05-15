package forms

import (
	"os"
	"path/filepath"
	"testing"
)

func TestExtractUserFormCodeFromFRM(t *testing.T) {
	body := "VERSION 5.00\r\nBegin {GUID} UserForm1\r\nEnd\r\nAttribute VB_Name = \"UserForm1\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n\r\nOption Explicit\r\n\r\nPrivate Sub UserForm_Initialize()\r\nEnd Sub\r\n"
	got := ExtractUserFormCodeFromFRM(body)
	want := "Option Explicit\n\nPrivate Sub UserForm_Initialize()\nEnd Sub\n"
	if got != want {
		t.Fatalf("ExtractUserFormCodeFromFRM() = %q, want %q", got, want)
	}
}

func TestSyncUserFormCodeSidecars(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	frmBody := "VERSION 5.00\nBegin {GUID} CustomerForm\nEnd\nAttribute VB_Name = \"CustomerForm\"\nAttribute VB_GlobalNameSpace = False\n\nOption Explicit\n\nPrivate Sub UserForm_Initialize()\n    version = \"frm\"\nEnd Sub\n"
	sidecarBody := "Option Explicit\n\nPrivate Sub UserForm_Initialize()\n    version = \"sidecar\"\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(formsDir, "CustomerForm.frm"), []byte(frmBody), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "CustomerForm.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}
	updated, err := SyncUserFormCodeSidecars(formsDir, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(updated) != 1 || updated[0].FormName != "CustomerForm" {
		t.Fatalf("updated = %#v", updated)
	}
	rewritten, err := os.ReadFile(filepath.Join(formsDir, "CustomerForm.frm"))
	if err != nil {
		t.Fatal(err)
	}
	if got := NormalizeUserFormCodeText(ExtractUserFormCodeFromFRM(string(rewritten))); got != NormalizeUserFormCodeText(sidecarBody) {
		t.Fatalf("rewritten frm code = %q, want %q", got, NormalizeUserFormCodeText(sidecarBody))
	}
}

func TestSyncUserFormCodeSidecarsHonorsTargetFilter(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "A.frm"), []byte("Attribute VB_Name = \"A\"\n\nOption Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "A.bas"), []byte("Option Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "B.frm"), []byte("Attribute VB_Name = \"B\"\n\nOption Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "B.bas"), []byte("Option Explicit\n\nPrivate Sub Test()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	updated, err := SyncUserFormCodeSidecars(formsDir, map[string]bool{"A": true})
	if err != nil {
		t.Fatal(err)
	}
	if len(updated) != 0 {
		t.Fatalf("updated = %#v, want none for filtered forms", updated)
	}
	frmBody, err := os.ReadFile(filepath.Join(formsDir, "B.frm"))
	if err != nil {
		t.Fatal(err)
	}
	if got := NormalizeUserFormCodeText(ExtractUserFormCodeFromFRM(string(frmBody))); got != "Option Explicit" {
		t.Fatalf("unfiltered form should remain unchanged, got %q", got)
	}
}

func TestValidateUserFormCodeSidecarsRejectsAttributeHeaders(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "CustomerForm.frm"), []byte("Attribute VB_Name = \"CustomerForm\"\n\nOption Explicit\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecarBody := "Attribute VB_Name = \"CustomerForm\"\nOption Explicit\n"
	if err := os.WriteFile(filepath.Join(formsDir, "code", "CustomerForm.bas"), []byte(sidecarBody), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := ValidateUserFormCodeSidecars(formsDir, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 1 {
		t.Fatalf("issues = %#v, want one", issues)
	}
	if issues[0].FormName != "CustomerForm" || issues[0].Line != 1 {
		t.Fatalf("unexpected issue = %#v", issues[0])
	}
}
