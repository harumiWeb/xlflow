package coordination

import (
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

func TestNewWorkbookIdentityEquivalentNativePaths(t *testing.T) {
	baseDir := t.TempDir()
	direct, err := NewWorkbookIdentity(baseDir, filepath.Join("project", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(direct): %v", err)
	}
	withSegments, err := NewWorkbookIdentity(baseDir, filepath.Join("project", ".", "nested", "..", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(with segments): %v", err)
	}

	if direct != withSegments {
		t.Fatalf("equivalent paths produced different identities:\n direct: %#v\nsegments: %#v", direct, withSegments)
	}
	if !filepath.IsAbs(direct.CanonicalPath) {
		t.Fatalf("CanonicalPath = %q, want absolute path", direct.CanonicalPath)
	}
}

func TestNewWorkbookIdentityAbsolutePathIgnoresBaseForResolution(t *testing.T) {
	firstBase := t.TempDir()
	secondBase := t.TempDir()
	workbookPath := filepath.Join(firstBase, "book.xlsm")

	first, err := NewWorkbookIdentity(firstBase, workbookPath)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(first base): %v", err)
	}
	second, err := NewWorkbookIdentity(secondBase, workbookPath)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(second base): %v", err)
	}
	if first != second {
		t.Fatalf("absolute path identities differ: %#v != %#v", first, second)
	}
}

func TestNewWorkbookIdentityDifferentPaths(t *testing.T) {
	baseDir := t.TempDir()
	first, err := NewWorkbookIdentity(baseDir, "first.xlsm")
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(first): %v", err)
	}
	second, err := NewWorkbookIdentity(baseDir, "second.xlsm")
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(second): %v", err)
	}
	if first.LockID == second.LockID {
		t.Fatalf("different workbook paths share LockID %q", first.LockID)
	}
}

func TestNewWorkbookIdentityDoesNotRequireWorkbook(t *testing.T) {
	baseDir := t.TempDir()
	wantPath := filepath.Join(baseDir, "missing", "book.xlsm")

	identity, err := NewWorkbookIdentity(baseDir, filepath.Join("missing", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity: %v", err)
	}
	if identity.CanonicalPath != filepath.Clean(wantPath) {
		t.Fatalf("CanonicalPath = %q, want %q", identity.CanonicalPath, filepath.Clean(wantPath))
	}
}

func TestNewWorkbookIdentityResolvesExistingParentAliasForMissingWorkbook(t *testing.T) {
	baseDir := t.TempDir()
	realParent := filepath.Join(baseDir, "real-project")
	if err := os.Mkdir(realParent, 0o700); err != nil {
		t.Fatalf("mkdir real parent: %v", err)
	}
	aliasParent := filepath.Join(baseDir, "alias-project")
	if err := os.Symlink(realParent, aliasParent); err != nil {
		t.Skipf("creating a directory symlink is unavailable: %v", err)
	}

	realIdentity, err := NewWorkbookIdentity(baseDir, filepath.Join(realParent, "missing", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(real): %v", err)
	}
	aliasIdentity, err := NewWorkbookIdentity(baseDir, filepath.Join(aliasParent, "missing", "book.xlsm"))
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(alias): %v", err)
	}
	if aliasIdentity != realIdentity {
		t.Fatalf("missing workbook under parent alias = %#v, want %#v", aliasIdentity, realIdentity)
	}
}

func TestNewWorkbookIdentityPreservesUnicodeAndSpacesForDiagnostics(t *testing.T) {
	baseDir := t.TempDir()
	workbookPath := filepath.Join("顧客 ワークブック", "売上 2026.xlsm")

	identity, err := NewWorkbookIdentity(baseDir, workbookPath)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity: %v", err)
	}
	wantPath := filepath.Join(baseDir, workbookPath)
	if identity.CanonicalPath != wantPath {
		t.Fatalf("CanonicalPath = %q, want %q", identity.CanonicalPath, wantPath)
	}
}

func TestNewWorkbookIdentityLockIDIsOpaqueAndStable(t *testing.T) {
	baseDir := t.TempDir()
	workbookPath := filepath.Join("private-customer-name", "book.xlsm")

	first, err := NewWorkbookIdentity(baseDir, workbookPath)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(first): %v", err)
	}
	second, err := NewWorkbookIdentity(baseDir, workbookPath)
	if err != nil {
		t.Fatalf("NewWorkbookIdentity(second): %v", err)
	}
	if first.LockID != second.LockID {
		t.Fatalf("LockID is not deterministic: %q != %q", first.LockID, second.LockID)
	}
	if !regexp.MustCompile(`^xlflow-workbook-v1-[0-9a-f]{64}$`).MatchString(first.LockID) {
		t.Fatalf("LockID = %q, want fixed ASCII prefix and SHA-256 hex", first.LockID)
	}
	if strings.Contains(strings.ToLower(first.LockID), "private-customer-name") ||
		strings.Contains(strings.ToLower(first.LockID), "book.xlsm") {
		t.Fatalf("LockID exposes workbook path: %q", first.LockID)
	}
}

func TestNewWorkbookIdentityRejectsInvalidInput(t *testing.T) {
	absoluteBase := t.TempDir()
	tests := []struct {
		name         string
		baseDir      string
		workbookPath string
	}{
		{name: "empty base", workbookPath: "book.xlsm"},
		{name: "relative base", baseDir: "relative", workbookPath: "book.xlsm"},
		{name: "empty workbook", baseDir: absoluteBase},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if _, err := NewWorkbookIdentity(tt.baseDir, tt.workbookPath); err == nil {
				t.Fatal("NewWorkbookIdentity() error = nil, want error")
			}
		})
	}
}
