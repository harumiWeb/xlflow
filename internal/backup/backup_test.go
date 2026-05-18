package backup

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func TestCreateListLatestAndRestore(t *testing.T) {
	root := t.TempDir()
	workbookDir := filepath.Join(root, "build")
	if err := os.MkdirAll(workbookDir, 0o755); err != nil {
		t.Fatal(err)
	}
	workbookPath := filepath.Join(workbookDir, "Book.xlsm")
	if err := os.WriteFile(workbookPath, []byte("original"), 0o644); err != nil {
		t.Fatal(err)
	}

	first, err := Create(root, workbookPath, "before-push", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(workbookPath, []byte("changed"), 0o644); err != nil {
		t.Fatal(err)
	}
	second, err := Create(root, workbookPath, "pre-rollback", time.Date(2026, 5, 18, 11, 0, 0, 0, time.UTC))
	if err != nil {
		t.Fatal(err)
	}

	records, err := List(root, workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(records) != 2 {
		t.Fatalf("records = %d, want 2", len(records))
	}
	if records[0].ID != second.ID || records[1].ID != first.ID {
		t.Fatalf("records order = %#v", records)
	}
	latest, err := Latest(root, workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	if latest.ID != second.ID {
		t.Fatalf("latest = %q, want %q", latest.ID, second.ID)
	}
	if err := os.WriteFile(workbookPath, []byte("broken"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := Restore(workbookPath, first); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	if string(body) != "original" {
		t.Fatalf("restored body = %q, want original", string(body))
	}
}

func TestListIgnoresLegacyBackupDirectoriesWithoutMetadata(t *testing.T) {
	root := t.TempDir()
	workbookPath := filepath.Join(root, "build", "Book.xlsm")
	if err := os.MkdirAll(filepath.Dir(workbookPath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(workbookPath, []byte("book"), 0o644); err != nil {
		t.Fatal(err)
	}
	legacyDir := filepath.Join(Root(root), "20260518-100000")
	if err := os.MkdirAll(legacyDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(legacyDir, "Module1.bas"), []byte("Attribute VB_Name = \"Module1\""), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := Create(root, workbookPath, "before-push", time.Date(2026, 5, 18, 10, 30, 0, 0, time.UTC)); err != nil {
		t.Fatal(err)
	}

	records, err := List(root, workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(records) != 1 {
		t.Fatalf("records = %d, want 1", len(records))
	}
}

func TestFindFiltersByWorkbookPath(t *testing.T) {
	root := t.TempDir()
	bookA := filepath.Join(root, "build", "A.xlsm")
	bookB := filepath.Join(root, "build", "B.xlsm")
	if err := os.MkdirAll(filepath.Dir(bookA), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(bookA, []byte("A"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(bookB, []byte("B"), 0o644); err != nil {
		t.Fatal(err)
	}
	record, err := Create(root, bookA, "before-push", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC))
	if err != nil {
		t.Fatal(err)
	}
	if _, err := Find(root, bookB, record.ID); err == nil {
		t.Fatal("expected missing backup for other workbook")
	}
}

func TestCreateAddsNumericSuffixOnCollision(t *testing.T) {
	root := t.TempDir()
	workbookPath := filepath.Join(root, "build", "Book.xlsm")
	if err := os.MkdirAll(filepath.Dir(workbookPath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(workbookPath, []byte("book"), 0o644); err != nil {
		t.Fatal(err)
	}
	now := time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC)
	first, err := Create(root, workbookPath, "before push", now)
	if err != nil {
		t.Fatal(err)
	}
	second, err := Create(root, workbookPath, "before push", now)
	if err != nil {
		t.Fatal(err)
	}
	if first.ID == second.ID {
		t.Fatalf("expected unique IDs, got %q", first.ID)
	}
	if !strings.HasPrefix(second.ID, first.ID+"-") {
		t.Fatalf("second ID = %q, want prefix %q-", second.ID, first.ID)
	}
}
