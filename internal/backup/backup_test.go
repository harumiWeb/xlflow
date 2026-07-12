package backup

import (
	"errors"
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

func TestScanReturnsValidRecordsAndClassifiesInvalidAndLegacyEntries(t *testing.T) {
	root := t.TempDir()
	workbookPath := writeWorkbook(t, root, "Book.xlsm", "book")
	valid := createBackupEntry(t, root, "valid", workbookPath, "Book.xlsm", "valid book", time.Date(2026, 5, 18, 12, 0, 0, 0, time.UTC))

	invalidJSONDir := filepath.Join(Root(root), "invalid-json")
	if err := os.MkdirAll(invalidJSONDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(invalidJSONDir, metadataFileName), []byte("{"), 0o644); err != nil {
		t.Fatal(err)
	}

	missingFileDir := filepath.Join(Root(root), "missing-file")
	if err := os.MkdirAll(missingFileDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := writeMetadata(filepath.Join(missingFileDir, metadataFileName), Metadata{
		ID:                   "missing-file",
		CreatedAt:            time.Date(2026, 5, 18, 13, 0, 0, 0, time.UTC),
		Reason:               "before-push",
		OriginalWorkbookPath: workbookPath,
		BackupFilePath:       "missing.xlsm",
	}); err != nil {
		t.Fatal(err)
	}

	missingFieldDir := filepath.Join(Root(root), "missing-field")
	if err := os.MkdirAll(missingFieldDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := writeMetadata(filepath.Join(missingFieldDir, metadataFileName), Metadata{
		ID:                   "missing-field",
		CreatedAt:            time.Date(2026, 5, 18, 14, 0, 0, 0, time.UTC),
		OriginalWorkbookPath: workbookPath,
		BackupFilePath:       "Book.xlsm",
	}); err != nil {
		t.Fatal(err)
	}

	invalidTimestampDir := filepath.Join(Root(root), "invalid-timestamp")
	if err := os.MkdirAll(invalidTimestampDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := writeMetadata(filepath.Join(invalidTimestampDir, metadataFileName), Metadata{
		ID:                   "invalid-timestamp",
		Reason:               "before-push",
		OriginalWorkbookPath: workbookPath,
		BackupFilePath:       "Book.xlsm",
	}); err != nil {
		t.Fatal(err)
	}

	unsafeDir := filepath.Join(Root(root), "unsafe-path")
	if err := os.MkdirAll(unsafeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := writeMetadata(filepath.Join(unsafeDir, metadataFileName), Metadata{
		ID:                   "unsafe-path",
		CreatedAt:            time.Date(2026, 5, 18, 15, 0, 0, 0, time.UTC),
		Reason:               "before-push",
		OriginalWorkbookPath: workbookPath,
		BackupFilePath:       filepath.Join("..", "..", "outside.xlsm"),
	}); err != nil {
		t.Fatal(err)
	}

	legacyDir := filepath.Join(Root(root), "legacy")
	if err := os.MkdirAll(legacyDir, 0o755); err != nil {
		t.Fatal(err)
	}

	scan, err := Scan(root, workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(scan.Records) != 1 || scan.Records[0].ID != valid.ID {
		t.Fatalf("records = %#v, want valid only", scan.Records)
	}
	if scan.Records[0].SizeBytes != int64(len("valid book")) {
		t.Fatalf("size = %d, want %d", scan.Records[0].SizeBytes, len("valid book"))
	}
	if len(scan.Invalid) != 5 {
		t.Fatalf("invalid = %#v, want 5", scan.Invalid)
	}
	for _, code := range []string{"invalid_metadata_json", "missing_backup_file", "missing_required_field", "invalid_created_at", "unsafe_backup_file_path"} {
		if !hasInvalidCode(scan.Invalid, code) {
			t.Fatalf("invalid entries missing code %q: %#v", code, scan.Invalid)
		}
	}
	if len(scan.Legacy) != 1 || !samePath(scan.Legacy[0].Directory, legacyDir) {
		t.Fatalf("legacy = %#v, want %s", scan.Legacy, legacyDir)
	}
}

func TestScanSortsByMetadataTimestampAndFiltersWorkbook(t *testing.T) {
	root := t.TempDir()
	bookA := writeWorkbook(t, root, "A.xlsm", "A")
	bookB := writeWorkbook(t, root, "B.xlsm", "B")

	older := createBackupEntry(t, root, "20991231-235959-push", bookA, "A.xlsm", "older", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC))
	newer := createBackupEntry(t, root, "20000101-000000-push", bookA, "A.xlsm", "newer", time.Date(2026, 5, 18, 11, 0, 0, 0, time.UTC))
	createBackupEntry(t, root, "other-workbook", bookB, "B.xlsm", "other", time.Date(2026, 5, 18, 12, 0, 0, 0, time.UTC))

	records, err := List(root, bookA)
	if err != nil {
		t.Fatal(err)
	}
	if len(records) != 2 {
		t.Fatalf("records = %#v, want 2", records)
	}
	if records[0].ID != newer.ID || records[1].ID != older.ID {
		t.Fatalf("record order = %#v, want metadata timestamp order", records)
	}
}

func TestParseRetentionDurationAndSize(t *testing.T) {
	duration, err := ParseRetentionDuration("2w")
	if err != nil {
		t.Fatal(err)
	}
	if duration != 14*24*time.Hour {
		t.Fatalf("duration = %s, want 336h", duration)
	}
	if _, err := ParseRetentionDuration("-1d"); err == nil {
		t.Fatal("expected negative duration to fail")
	}
	size, err := ParseSize("2GB")
	if err != nil {
		t.Fatal(err)
	}
	if size != 2_000_000_000 {
		t.Fatalf("size = %d, want decimal 2GB", size)
	}
	if _, err := ParseSize("2GiB"); err == nil {
		t.Fatal("expected unsupported size unit to fail")
	}
}

func TestPruneKeepLastProtectsNewestAcrossConditions(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, "Book.xlsm", "book")
	base := time.Date(2026, 7, 12, 12, 0, 0, 0, time.UTC)
	createBackupEntry(t, root, "old", workbook, "Book.xlsm", strings.Repeat("a", 10), base.Add(-72*time.Hour))
	createBackupEntry(t, root, "middle", workbook, "Book.xlsm", strings.Repeat("b", 10), base.Add(-48*time.Hour))
	createBackupEntry(t, root, "new", workbook, "Book.xlsm", strings.Repeat("c", 10), base.Add(-24*time.Hour))
	keep := 2

	result, err := Prune(root, workbook, PruneOptions{
		KeepLast:     &keep,
		OlderThan:    time.Hour,
		MaxTotalSize: 1,
		DryRun:       true,
		Now:          base,
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Candidates) != 1 || result.Candidates[0].ID != "old" {
		t.Fatalf("candidates = %#v, want only old", result.Candidates)
	}
	if got := strings.Join(result.Candidates[0].Reasons, ","); got != "exceeds_keep_last,exceeds_max_total_size,older_than" {
		t.Fatalf("reasons = %q", got)
	}
}

func TestPruneUsesMetadataTimestampNotBackupID(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, "Book.xlsm", "book")
	now := time.Date(2026, 7, 12, 12, 0, 0, 0, time.UTC)
	createBackupEntry(t, root, "19990101-000000-push", workbook, "Book.xlsm", "new", now.Add(-time.Hour))
	createBackupEntry(t, root, "20990101-000000-push", workbook, "Book.xlsm", "old", now.Add(-48*time.Hour))

	result, err := Prune(root, workbook, PruneOptions{OlderThan: 24 * time.Hour, DryRun: true, Now: now})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Candidates) != 1 || result.Candidates[0].ID != "20990101-000000-push" {
		t.Fatalf("candidates = %#v, want metadata-old backup", result.Candidates)
	}
}

func TestPruneScopeAndAllWorkbooks(t *testing.T) {
	root := t.TempDir()
	bookA := writeWorkbook(t, root, "A.xlsm", "A")
	bookB := writeWorkbook(t, root, "B.xlsm", "B")
	now := time.Date(2026, 7, 12, 12, 0, 0, 0, time.UTC)
	createBackupEntry(t, root, "a-old", bookA, "A.xlsm", "A", now.Add(-48*time.Hour))
	createBackupEntry(t, root, "b-old", bookB, "B.xlsm", "B", now.Add(-48*time.Hour))

	scoped, err := Prune(root, bookA, PruneOptions{OlderThan: time.Hour, DryRun: true, Now: now})
	if err != nil {
		t.Fatal(err)
	}
	if len(scoped.Candidates) != 1 || scoped.Candidates[0].ID != "a-old" {
		t.Fatalf("scoped candidates = %#v", scoped.Candidates)
	}
	all, err := Prune(root, bookA, PruneOptions{OlderThan: time.Hour, DryRun: true, AllWorkbooks: true, Now: now})
	if err != nil {
		t.Fatal(err)
	}
	if len(all.Candidates) != 2 {
		t.Fatalf("all candidates = %#v, want 2", all.Candidates)
	}
}

func TestPruneDryRunDoesNotDeleteAndRealPruneDeletesSameCandidate(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, "Book.xlsm", "book")
	now := time.Date(2026, 7, 12, 12, 0, 0, 0, time.UTC)
	old := createBackupEntry(t, root, "old", workbook, "Book.xlsm", "old", now.Add(-48*time.Hour))
	createBackupEntry(t, root, "new", workbook, "Book.xlsm", "new", now.Add(-time.Hour))
	keep := 1

	dryRun, err := Prune(root, workbook, PruneOptions{KeepLast: &keep, DryRun: true, Now: now})
	if err != nil {
		t.Fatal(err)
	}
	if len(dryRun.Candidates) != 1 || dryRun.Candidates[0].ID != old.ID {
		t.Fatalf("dry-run candidates = %#v", dryRun.Candidates)
	}
	if _, err := os.Stat(old.Directory); err != nil {
		t.Fatalf("dry-run deleted directory: %v", err)
	}
	actual, err := Prune(root, workbook, PruneOptions{KeepLast: &keep, Now: now})
	if err != nil {
		t.Fatal(err)
	}
	if actual.Deleted != 1 || actual.FreedBytes <= int64(len("old")) {
		t.Fatalf("actual result = %#v", actual)
	}
	if _, err := os.Stat(old.Directory); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("old directory still exists or unexpected error: %v", err)
	}
}

func TestPruneIncludeInvalidAndLegacyOnlyWhenExplicit(t *testing.T) {
	root := t.TempDir()
	workbook := writeWorkbook(t, root, "Book.xlsm", "book")
	invalidDir := filepath.Join(Root(root), "invalid")
	if err := os.MkdirAll(invalidDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(invalidDir, metadataFileName), []byte("{"), 0o644); err != nil {
		t.Fatal(err)
	}
	legacyDir := filepath.Join(Root(root), "legacy")
	if err := os.MkdirAll(legacyDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(legacyDir, "legacy.txt"), []byte("legacy"), 0o644); err != nil {
		t.Fatal(err)
	}

	keep := 99
	excluded, err := Prune(root, workbook, PruneOptions{KeepLast: &keep, DryRun: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(excluded.Candidates) != 0 {
		t.Fatalf("candidates = %#v, want invalid/legacy excluded", excluded.Candidates)
	}
	included, err := Prune(root, workbook, PruneOptions{IncludeInvalid: true, IncludeLegacy: true, DryRun: true})
	if err != nil {
		t.Fatal(err)
	}
	if len(included.Candidates) != 2 {
		t.Fatalf("included candidates = %#v, want invalid and legacy", included.Candidates)
	}
}

func TestDeleteRejectsScopeMismatchAndUnsafeManagedDirectory(t *testing.T) {
	root := t.TempDir()
	bookA := writeWorkbook(t, root, "A.xlsm", "A")
	bookB := writeWorkbook(t, root, "B.xlsm", "B")
	if _, err := Delete(root, bookA, " "); err == nil || !backupErrorCodeIs(err, ErrDeleteArgsInvalid) {
		t.Fatalf("Delete empty ID err = %v, want %s", err, ErrDeleteArgsInvalid)
	}
	record := createBackupEntry(t, root, "backup-a", bookA, "A.xlsm", "A", time.Now())
	if _, err := Delete(root, bookB, record.ID); err == nil || !backupErrorCodeIs(err, ErrDeleteScopeMismatch) {
		t.Fatalf("Delete scope err = %v, want %s", err, ErrDeleteScopeMismatch)
	}

	record.Directory = Root(root)
	if _, err := deleteManagedDirectory(root, record.ID, record.Directory); err == nil || !backupErrorCodeIs(err, ErrDeleteUnsafePath) {
		t.Fatalf("deleteManagedDirectory err = %v, want %s", err, ErrDeleteUnsafePath)
	}
}

func TestCreateCleansCreatedDirectoryOnCopyFailure(t *testing.T) {
	root := t.TempDir()
	missingWorkbook := filepath.Join(root, "missing", "Book.xlsm")
	if _, err := Create(root, missingWorkbook, "before-push", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC)); err == nil {
		t.Fatal("expected copy failure")
	}
	entries, err := os.ReadDir(Root(root))
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 0 {
		t.Fatalf("backup entries = %d, want cleanup", len(entries))
	}
}

func TestCreateCleansCreatedDirectoryOnMetadataFailure(t *testing.T) {
	root := t.TempDir()
	workbookPath := writeWorkbook(t, root, "Book.xlsm", "book")
	originalWriteFile := writeFile
	writeFile = func(path string, data []byte, perm os.FileMode) error {
		if filepath.Base(path) == metadataFileName {
			return errors.New("metadata boom")
		}
		return originalWriteFile(path, data, perm)
	}
	t.Cleanup(func() { writeFile = originalWriteFile })

	if _, err := Create(root, workbookPath, "before-push", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC)); err == nil || !strings.Contains(err.Error(), "metadata boom") {
		t.Fatalf("Create error = %v, want metadata boom", err)
	}
	entries, err := os.ReadDir(Root(root))
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 0 {
		t.Fatalf("backup entries = %d, want cleanup", len(entries))
	}
}

func TestCreatePreservesOriginalErrorWhenCleanupFails(t *testing.T) {
	root := t.TempDir()
	workbookPath := writeWorkbook(t, root, "Book.xlsm", "book")
	originalWriteFile := writeFile
	writeFile = func(path string, data []byte, perm os.FileMode) error {
		if filepath.Base(path) == metadataFileName {
			return errors.New("metadata boom")
		}
		return originalWriteFile(path, data, perm)
	}
	t.Cleanup(func() { writeFile = originalWriteFile })
	originalRemoveAll := removeAll
	removeAll = func(path string) error { return errors.New("cleanup boom") }
	t.Cleanup(func() { removeAll = originalRemoveAll })

	_, err := Create(root, workbookPath, "before-push", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC))
	if err == nil {
		t.Fatal("expected error")
	}
	if !strings.Contains(err.Error(), "metadata boom") || !strings.Contains(err.Error(), "backup_cleanup_failed") {
		t.Fatalf("Create error = %v, want original and cleanup context", err)
	}
}

func TestCreateKeepsSuccessfulBackupDirectory(t *testing.T) {
	root := t.TempDir()
	workbookPath := writeWorkbook(t, root, "Book.xlsm", "book")
	record, err := Create(root, workbookPath, "before-push", time.Date(2026, 5, 18, 10, 0, 0, 0, time.UTC))
	if err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(record.Directory); err != nil {
		t.Fatalf("backup directory missing after success: %v", err)
	}
	if record.SizeBytes != int64(len("book")) {
		t.Fatalf("size = %d, want %d", record.SizeBytes, len("book"))
	}
}

func writeWorkbook(t *testing.T, root, name, body string) string {
	t.Helper()
	path := filepath.Join(root, "build", name)
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	return path
}

func createBackupEntry(t *testing.T, root, id, workbookPath, backupName, body string, createdAt time.Time) Record {
	t.Helper()
	dir := filepath.Join(Root(root), id)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	backupPath := filepath.Join(dir, backupName)
	if err := os.WriteFile(backupPath, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	metadata := Metadata{
		ID:                   id,
		CreatedAt:            createdAt,
		Reason:               "before-push",
		OriginalWorkbookPath: workbookPath,
		BackupFilePath:       backupName,
	}
	if err := writeMetadata(filepath.Join(dir, metadataFileName), metadata); err != nil {
		t.Fatal(err)
	}
	return Record{Metadata: metadata, Directory: dir, BackupFileAbsPath: backupPath, SizeBytes: int64(len(body))}
}

func hasInvalidCode(entries []InvalidEntry, code string) bool {
	for _, entry := range entries {
		if entry.Code == code {
			return true
		}
	}
	return false
}

func backupErrorCodeIs(err error, code string) bool {
	var backupErr *Error
	return errors.As(err, &backupErr) && backupErr.Code == code
}
