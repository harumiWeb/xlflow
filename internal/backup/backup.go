package backup

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"
)

const metadataFileName = "metadata.json"

type Metadata struct {
	ID                   string    `json:"id"`
	CreatedAt            time.Time `json:"created_at"`
	Reason               string    `json:"reason"`
	OriginalWorkbookPath string    `json:"original_workbook_path"`
	BackupFilePath       string    `json:"backup_file_path"`
}

type Record struct {
	Metadata
	Directory         string
	BackupFileAbsPath string
}

func Root(rootDir string) string {
	return filepath.Join(rootDir, ".xlflow", "backups")
}

func List(rootDir, workbookPath string) ([]Record, error) {
	workbookAbs, err := filepath.Abs(workbookPath)
	if err != nil {
		return nil, err
	}
	backupRoot := Root(rootDir)
	entries, err := os.ReadDir(backupRoot)
	if errors.Is(err, os.ErrNotExist) {
		return []Record{}, nil
	}
	if err != nil {
		return nil, err
	}

	records := make([]Record, 0, len(entries))
	for _, entry := range entries {
		if !entry.IsDir() {
			continue
		}
		record, ok, err := readRecord(filepath.Join(backupRoot, entry.Name()))
		if err != nil {
			return nil, err
		}
		if !ok {
			continue
		}
		if !samePath(record.OriginalWorkbookPath, workbookAbs) {
			continue
		}
		records = append(records, record)
	}
	sort.Slice(records, func(i, j int) bool {
		if records[i].CreatedAt.Equal(records[j].CreatedAt) {
			return records[i].ID > records[j].ID
		}
		return records[i].CreatedAt.After(records[j].CreatedAt)
	})
	return records, nil
}

func Find(rootDir, workbookPath, backupID string) (Record, error) {
	records, err := List(rootDir, workbookPath)
	if err != nil {
		return Record{}, err
	}
	for _, record := range records {
		if record.ID == backupID {
			return record, nil
		}
	}
	return Record{}, fmt.Errorf("backup %q was not found for %s", backupID, workbookPath)
}

func Latest(rootDir, workbookPath string) (Record, error) {
	records, err := List(rootDir, workbookPath)
	if err != nil {
		return Record{}, err
	}
	if len(records) == 0 {
		return Record{}, fmt.Errorf("no backups found for %s", workbookPath)
	}
	return records[0], nil
}

func Create(rootDir, workbookPath, reason string, now time.Time) (Record, error) {
	if strings.TrimSpace(reason) == "" {
		return Record{}, errors.New("backup reason is required")
	}
	workbookAbs, err := filepath.Abs(workbookPath)
	if err != nil {
		return Record{}, err
	}
	if now.IsZero() {
		now = time.Now()
	}
	id := uniqueID(rootDir, now, reason)
	dir := filepath.Join(Root(rootDir), id)
	if err := os.MkdirAll(dir, 0o755); err != nil {
		return Record{}, err
	}
	backupName := filepath.Base(workbookAbs)
	backupAbs := filepath.Join(dir, backupName)
	if err := copyFile(workbookAbs, backupAbs); err != nil {
		return Record{}, err
	}
	metadata := Metadata{
		ID:                   id,
		CreatedAt:            now,
		Reason:               reason,
		OriginalWorkbookPath: workbookAbs,
		BackupFilePath:       backupName,
	}
	if err := writeMetadata(filepath.Join(dir, metadataFileName), metadata); err != nil {
		return Record{}, err
	}
	return Record{
		Metadata:          metadata,
		Directory:         dir,
		BackupFileAbsPath: backupAbs,
	}, nil
}

func Restore(targetWorkbookPath string, record Record) error {
	targetAbs, err := filepath.Abs(targetWorkbookPath)
	if err != nil {
		return err
	}
	if !samePath(record.OriginalWorkbookPath, targetAbs) {
		return fmt.Errorf("backup %q does not belong to %s", record.ID, targetWorkbookPath)
	}
	parent := filepath.Dir(targetAbs)
	if err := os.MkdirAll(parent, 0o755); err != nil {
		return err
	}
	tempPath := filepath.Join(parent, fmt.Sprintf(".xlflow-rollback-%d.tmp", time.Now().UnixNano()))
	if err := copyFile(record.BackupFileAbsPath, tempPath); err != nil {
		return err
	}
	if err := os.Remove(targetAbs); err != nil {
		if !errors.Is(err, os.ErrNotExist) {
			_ = os.Remove(tempPath)
			return err
		}
	}
	if err := os.Rename(tempPath, targetAbs); err != nil {
		_ = os.Remove(tempPath)
		return err
	}
	return nil
}

func readRecord(dir string) (Record, bool, error) {
	metadataPath := filepath.Join(dir, metadataFileName)
	body, err := os.ReadFile(metadataPath)
	if errors.Is(err, os.ErrNotExist) {
		return Record{}, false, nil
	}
	if err != nil {
		return Record{}, false, err
	}
	body = bytes.TrimPrefix(body, []byte{0xEF, 0xBB, 0xBF})
	var metadata Metadata
	if err := json.Unmarshal(body, &metadata); err != nil {
		return Record{}, false, fmt.Errorf("parse %s: %w", metadataPath, err)
	}
	if strings.TrimSpace(metadata.ID) == "" ||
		metadata.CreatedAt.IsZero() ||
		strings.TrimSpace(metadata.Reason) == "" ||
		strings.TrimSpace(metadata.OriginalWorkbookPath) == "" ||
		strings.TrimSpace(metadata.BackupFilePath) == "" {
		return Record{}, false, nil
	}
	backupAbs := metadata.BackupFilePath
	if !filepath.IsAbs(backupAbs) {
		backupAbs = filepath.Join(dir, metadata.BackupFilePath)
	}
	if _, err := os.Stat(backupAbs); err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return Record{}, false, nil
		}
		return Record{}, false, err
	}
	originalAbs, err := filepath.Abs(metadata.OriginalWorkbookPath)
	if err != nil {
		return Record{}, false, err
	}
	metadata.OriginalWorkbookPath = originalAbs
	return Record{
		Metadata:          metadata,
		Directory:         dir,
		BackupFileAbsPath: filepath.Clean(backupAbs),
	}, true, nil
}

func writeMetadata(path string, metadata Metadata) error {
	body, err := json.MarshalIndent(metadata, "", "  ")
	if err != nil {
		return err
	}
	body = append(body, '\n')
	return os.WriteFile(path, body, 0o644)
}

func uniqueID(rootDir string, now time.Time, reason string) string {
	base := now.Format("20060102-150405") + "-" + sanitizeReason(reason)
	id := base
	for i := 1; ; i++ {
		if _, err := os.Stat(filepath.Join(Root(rootDir), id)); errors.Is(err, os.ErrNotExist) {
			return id
		}
		id = fmt.Sprintf("%s-%d", base, i)
	}
}

func sanitizeReason(reason string) string {
	reason = strings.TrimSpace(strings.ToLower(reason))
	var b strings.Builder
	lastDash := false
	for _, r := range reason {
		switch {
		case r >= 'a' && r <= 'z':
			b.WriteRune(r)
			lastDash = false
		case r >= '0' && r <= '9':
			b.WriteRune(r)
			lastDash = false
		default:
			if b.Len() > 0 && !lastDash {
				b.WriteByte('-')
				lastDash = true
			}
		}
	}
	out := strings.Trim(b.String(), "-")
	if out == "" {
		return "backup"
	}
	return out
}

func copyFile(src, dst string) (err error) {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := in.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	out, err := os.OpenFile(dst, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0o644)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := out.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	if _, err := io.Copy(out, in); err != nil {
		return err
	}
	return out.Sync()
}

func samePath(a, b string) bool {
	aa := filepath.Clean(a)
	bb := filepath.Clean(b)
	return strings.EqualFold(aa, bb)
}
