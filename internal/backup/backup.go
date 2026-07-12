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

var (
	removeAll = os.RemoveAll
	writeFile = os.WriteFile
)

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
	SizeBytes         int64
}

type ScanResult struct {
	Records []Record
	Invalid []InvalidEntry
	Legacy  []LegacyEntry
}

type InvalidEntry struct {
	Directory string
	Code      string
	Message   string
}

type LegacyEntry struct {
	Directory string
}

func Root(rootDir string) string {
	return filepath.Join(rootDir, ".xlflow", "backups")
}

func Scan(rootDir, workbookPath string) (ScanResult, error) {
	workbookAbs, err := filepath.Abs(workbookPath)
	if err != nil {
		return ScanResult{}, err
	}
	backupRoot := Root(rootDir)
	entries, err := os.ReadDir(backupRoot)
	if errors.Is(err, os.ErrNotExist) {
		return ScanResult{}, nil
	}
	if err != nil {
		return ScanResult{}, err
	}

	result := ScanResult{
		Records: make([]Record, 0, len(entries)),
		Invalid: []InvalidEntry{},
		Legacy:  []LegacyEntry{},
	}
	for _, entry := range entries {
		if !entry.IsDir() {
			continue
		}
		dir := filepath.Join(backupRoot, entry.Name())
		record, state := scanRecord(backupRoot, dir)
		switch state.kind {
		case scanEntryValid:
			if !samePath(record.OriginalWorkbookPath, workbookAbs) {
				continue
			}
			result.Records = append(result.Records, record)
		case scanEntryInvalid:
			result.Invalid = append(result.Invalid, InvalidEntry{
				Directory: dir,
				Code:      state.code,
				Message:   state.message,
			})
		case scanEntryLegacy:
			result.Legacy = append(result.Legacy, LegacyEntry{Directory: dir})
		}
	}
	sort.Slice(result.Records, func(i, j int) bool {
		if result.Records[i].CreatedAt.Equal(result.Records[j].CreatedAt) {
			return result.Records[i].ID > result.Records[j].ID
		}
		return result.Records[i].CreatedAt.After(result.Records[j].CreatedAt)
	})
	return result, nil
}

func List(rootDir, workbookPath string) ([]Record, error) {
	result, err := Scan(rootDir, workbookPath)
	if err != nil {
		return nil, err
	}
	return result.Records, nil
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
		return Record{}, cleanupCreateDir(dir, err)
	}
	metadata := Metadata{
		ID:                   id,
		CreatedAt:            now,
		Reason:               reason,
		OriginalWorkbookPath: workbookAbs,
		BackupFilePath:       backupName,
	}
	if err := writeMetadata(filepath.Join(dir, metadataFileName), metadata); err != nil {
		return Record{}, cleanupCreateDir(dir, err)
	}
	info, err := os.Stat(backupAbs)
	if err != nil {
		return Record{}, cleanupCreateDir(dir, err)
	}
	return Record{
		Metadata:          metadata,
		Directory:         dir,
		BackupFileAbsPath: backupAbs,
		SizeBytes:         info.Size(),
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

type scanEntryKind int

const (
	scanEntryInvalid scanEntryKind = iota
	scanEntryLegacy
	scanEntryValid
)

type scanEntryState struct {
	kind    scanEntryKind
	code    string
	message string
}

func scanRecord(backupRoot, dir string) (Record, scanEntryState) {
	metadataPath := filepath.Join(dir, metadataFileName)
	body, err := os.ReadFile(metadataPath)
	if errors.Is(err, os.ErrNotExist) {
		return Record{}, scanEntryState{kind: scanEntryLegacy}
	}
	if err != nil {
		return Record{}, invalidState("metadata_read_failed", err)
	}
	body = bytes.TrimPrefix(body, []byte{0xEF, 0xBB, 0xBF})
	var metadata Metadata
	if err := json.Unmarshal(body, &metadata); err != nil {
		return Record{}, invalidState("invalid_metadata_json", err)
	}
	if strings.TrimSpace(metadata.ID) == "" {
		return Record{}, invalidMessage("missing_required_field", "metadata field id is required")
	}
	if metadata.CreatedAt.IsZero() {
		return Record{}, invalidMessage("invalid_created_at", "metadata field created_at is required and must be a valid timestamp")
	}
	if strings.TrimSpace(metadata.Reason) == "" {
		return Record{}, invalidMessage("missing_required_field", "metadata field reason is required")
	}
	if strings.TrimSpace(metadata.OriginalWorkbookPath) == "" {
		return Record{}, invalidMessage("missing_required_field", "metadata field original_workbook_path is required")
	}
	if strings.TrimSpace(metadata.BackupFilePath) == "" {
		return Record{}, invalidMessage("missing_required_field", "metadata field backup_file_path is required")
	}
	backupAbs := metadata.BackupFilePath
	if !filepath.IsAbs(backupAbs) {
		backupAbs = filepath.Join(dir, metadata.BackupFilePath)
	}
	backupAbs = filepath.Clean(backupAbs)
	backupRootAbs, err := filepath.Abs(backupRoot)
	if err != nil {
		return Record{}, invalidState("backup_root_invalid", err)
	}
	backupAbsFull, err := filepath.Abs(backupAbs)
	if err != nil {
		return Record{}, invalidState("unsafe_backup_file_path", err)
	}
	if !pathWithin(backupRootAbs, backupAbsFull) {
		return Record{}, invalidMessage("unsafe_backup_file_path", "backup file path must stay inside the managed backup directory")
	}
	info, err := os.Stat(backupAbsFull)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return Record{}, invalidMessage("missing_backup_file", "referenced backup file does not exist")
		}
		return Record{}, invalidState("backup_file_stat_failed", err)
	}
	if info.IsDir() {
		return Record{}, invalidMessage("missing_backup_file", "referenced backup file is a directory")
	}
	originalAbs, err := filepath.Abs(metadata.OriginalWorkbookPath)
	if err != nil {
		return Record{}, invalidState("unsafe_original_workbook_path", err)
	}
	metadata.OriginalWorkbookPath = originalAbs
	return Record{
		Metadata:          metadata,
		Directory:         dir,
		BackupFileAbsPath: filepath.Clean(backupAbsFull),
		SizeBytes:         info.Size(),
	}, scanEntryState{kind: scanEntryValid}
}

func invalidState(code string, err error) scanEntryState {
	return invalidMessage(code, err.Error())
}

func invalidMessage(code, message string) scanEntryState {
	return scanEntryState{kind: scanEntryInvalid, code: code, message: message}
}

func cleanupCreateDir(dir string, primary error) error {
	if cleanupErr := removeAll(dir); cleanupErr != nil {
		return errors.Join(primary, fmt.Errorf("backup_cleanup_failed: %w", cleanupErr))
	}
	return primary
}

func writeMetadata(path string, metadata Metadata) error {
	body, err := json.MarshalIndent(metadata, "", "  ")
	if err != nil {
		return err
	}
	body = append(body, '\n')
	return writeFile(path, body, 0o644)
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

func pathWithin(parent, child string) bool {
	parent = filepath.Clean(parent)
	child = filepath.Clean(child)
	if samePath(parent, child) {
		return true
	}
	rel, err := filepath.Rel(parent, child)
	if err != nil {
		return false
	}
	return rel != "." && !strings.HasPrefix(rel, ".."+string(filepath.Separator)) && rel != ".." && !filepath.IsAbs(rel)
}
