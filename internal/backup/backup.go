package backup

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"
)

const metadataFileName = "metadata.json"

var (
	removeAll = os.RemoveAll
	writeFile = os.WriteFile
)

const (
	ErrPruneArgsInvalid    = "backup_prune_args_invalid"
	ErrPruneFailed         = "backup_prune_failed"
	ErrDeleteArgsInvalid   = "backup_delete_args_invalid"
	ErrNotFound            = "backup_not_found"
	ErrDeleteFailed        = "backup_delete_failed"
	ErrDeleteUnsafePath    = "backup_delete_unsafe_path"
	ErrDeleteScopeMismatch = "backup_delete_scope_mismatch"
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

type Error struct {
	Code string
	Err  error
}

func (e *Error) Error() string {
	if e == nil || e.Err == nil {
		return ""
	}
	return e.Err.Error()
}

func (e *Error) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Err
}

type PruneOptions struct {
	KeepLast        *int
	OlderThan       time.Duration
	OlderThanSet    bool
	MaxTotalSize    int64
	MaxTotalSizeSet bool
	DryRun          bool
	AllWorkbooks    bool
	IncludeInvalid  bool
	IncludeLegacy   bool
	Now             time.Time
}

type PruneResult struct {
	DryRun         bool
	Matched        int
	Deleted        int
	Failed         int
	FreedBytes     int64
	Candidates     []CandidateEntry
	DeletedEntries []DeletedEntry
	FailedEntries  []FailedEntry
}

type DeleteResult struct {
	ID         string
	Path       string
	FreedBytes int64
}

type CandidateEntry struct {
	ID        string
	Directory string
	CreatedAt time.Time
	Reason    string
	SizeBytes int64
	Reasons   []string
	Status    string
	Code      string
	Message   string
}

type DeletedEntry struct {
	ID         string
	Directory  string
	FreedBytes int64
}

type FailedEntry struct {
	ID        string
	Directory string
	Code      string
	Message   string
}

func Root(rootDir string) string {
	return filepath.Join(rootDir, ".xlflow", "backups")
}

func ParseRetentionDuration(value string) (time.Duration, error) {
	value = strings.TrimSpace(value)
	if value == "" {
		return 0, fmt.Errorf("duration is required")
	}
	unit := value[len(value)-1:]
	rawNumber := strings.TrimSpace(value[:len(value)-1])
	rawNumber = strings.TrimPrefix(rawNumber, "+")
	n, err := strconv.ParseInt(rawNumber, 10, 64)
	if err != nil || n < 0 {
		return 0, fmt.Errorf("duration must be a non-negative integer followed by h, d, or w")
	}
	switch unit {
	case "h":
		return time.Duration(n) * time.Hour, nil
	case "d":
		return time.Duration(n) * 24 * time.Hour, nil
	case "w":
		return time.Duration(n) * 7 * 24 * time.Hour, nil
	default:
		return 0, fmt.Errorf("duration unit must be h, d, or w")
	}
}

func ParseSize(value string) (int64, error) {
	value = strings.TrimSpace(value)
	if value == "" {
		return 0, fmt.Errorf("size is required")
	}
	i := 0
	for ; i < len(value); i++ {
		ch := value[i]
		if ch < '0' || ch > '9' {
			break
		}
	}
	if i == 0 {
		return 0, fmt.Errorf("size must be a non-negative integer followed by KB, MB, or GB")
	}
	rawNumber := value[:i]
	unit := strings.ToUpper(strings.TrimSpace(value[i:]))
	n, err := strconv.ParseInt(rawNumber, 10, 64)
	if err != nil || n < 0 {
		return 0, fmt.Errorf("size must be a non-negative integer followed by KB, MB, or GB")
	}
	multiplier := int64(0)
	switch unit {
	case "KB":
		multiplier = 1000
	case "MB":
		multiplier = 1000 * 1000
	case "GB":
		multiplier = 1000 * 1000 * 1000
	default:
		return 0, fmt.Errorf("size unit must be KB, MB, or GB")
	}
	if n > math.MaxInt64/multiplier {
		return 0, fmt.Errorf("size is too large")
	}
	return n * multiplier, nil
}

func Scan(rootDir, workbookPath string) (ScanResult, error) {
	workbookAbs, err := filepath.Abs(workbookPath)
	if err != nil {
		return ScanResult{}, err
	}
	return scan(rootDir, func(record Record) bool {
		return samePath(record.OriginalWorkbookPath, workbookAbs)
	})
}

func ScanAll(rootDir string) (ScanResult, error) {
	return scan(rootDir, nil)
}

func scan(rootDir string, includeRecord func(Record) bool) (ScanResult, error) {
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
			if includeRecord != nil && !includeRecord(record) {
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

func Prune(rootDir, workbookPath string, opts PruneOptions) (PruneResult, error) {
	if err := validatePruneOptions(opts); err != nil {
		return PruneResult{}, &Error{Code: ErrPruneArgsInvalid, Err: err}
	}
	scan, err := scanForPrune(rootDir, workbookPath, opts.AllWorkbooks)
	if err != nil {
		return PruneResult{}, &Error{Code: ErrPruneFailed, Err: err}
	}
	now := opts.Now
	if now.IsZero() {
		now = time.Now()
	}
	candidates, err := selectPruneCandidates(rootDir, scan, opts, now)
	if err != nil {
		return PruneResult{}, err
	}
	result := PruneResult{
		DryRun:     opts.DryRun,
		Matched:    len(candidates),
		Candidates: candidates,
	}
	if opts.DryRun {
		return result, nil
	}
	for _, candidate := range candidates {
		deleted, err := deleteManagedDirectory(rootDir, candidate.ID, candidate.Directory)
		if err != nil {
			result.FailedEntries = append(result.FailedEntries, FailedEntry{
				ID:        candidate.ID,
				Directory: candidate.Directory,
				Code:      codeForDeleteError(err),
				Message:   err.Error(),
			})
			continue
		}
		result.DeletedEntries = append(result.DeletedEntries, deleted)
		result.FreedBytes += deleted.FreedBytes
	}
	result.Deleted = len(result.DeletedEntries)
	result.Failed = len(result.FailedEntries)
	result.Candidates = nil
	if result.Failed > 0 {
		return result, &Error{Code: ErrPruneFailed, Err: fmt.Errorf("%d backup deletion(s) failed", result.Failed)}
	}
	return result, nil
}

func Delete(rootDir, workbookPath, backupID string) (DeleteResult, error) {
	backupID = strings.TrimSpace(backupID)
	if backupID == "" {
		return DeleteResult{}, &Error{Code: ErrDeleteArgsInvalid, Err: fmt.Errorf("--backup is required")}
	}
	scan, err := ScanAll(rootDir)
	if err != nil {
		return DeleteResult{}, &Error{Code: ErrDeleteFailed, Err: err}
	}
	workbookAbs, err := filepath.Abs(workbookPath)
	if err != nil {
		return DeleteResult{}, &Error{Code: ErrDeleteFailed, Err: err}
	}
	var matches []Record
	for _, record := range scan.Records {
		if record.ID == backupID {
			matches = append(matches, record)
		}
	}
	if len(matches) == 0 {
		return DeleteResult{}, &Error{Code: ErrNotFound, Err: fmt.Errorf("backup %q was not found", backupID)}
	}
	if len(matches) > 1 {
		return DeleteResult{}, &Error{Code: ErrDeleteUnsafePath, Err: fmt.Errorf("backup %q is ambiguous", backupID)}
	}
	record := matches[0]
	if !samePath(record.OriginalWorkbookPath, workbookAbs) {
		return DeleteResult{}, &Error{Code: ErrDeleteScopeMismatch, Err: fmt.Errorf("backup %q does not belong to %s", backupID, workbookPath)}
	}
	deleted, err := deleteManagedDirectory(rootDir, record.ID, record.Directory)
	if err != nil {
		return DeleteResult{}, &Error{Code: codeForDeleteError(err), Err: err}
	}
	return DeleteResult{ID: deleted.ID, Path: deleted.Directory, FreedBytes: deleted.FreedBytes}, nil
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

func validatePruneOptions(opts PruneOptions) error {
	olderThanSet := opts.OlderThanSet || opts.OlderThan > 0
	maxTotalSizeSet := opts.MaxTotalSizeSet || opts.MaxTotalSize > 0
	if opts.KeepLast != nil && *opts.KeepLast < 0 {
		return fmt.Errorf("--keep-last must be a non-negative integer")
	}
	if opts.OlderThan < 0 {
		return fmt.Errorf("--older-than must be non-negative")
	}
	if opts.MaxTotalSize < 0 {
		return fmt.Errorf("--max-total-size must be non-negative")
	}
	if opts.KeepLast == nil && !olderThanSet && !maxTotalSizeSet && !opts.IncludeInvalid && !opts.IncludeLegacy {
		return fmt.Errorf("at least one pruning condition or include flag is required")
	}
	return nil
}

func scanForPrune(rootDir, workbookPath string, allWorkbooks bool) (ScanResult, error) {
	if allWorkbooks {
		return ScanAll(rootDir)
	}
	return Scan(rootDir, workbookPath)
}

func selectPruneCandidates(rootDir string, scan ScanResult, opts PruneOptions, now time.Time) ([]CandidateEntry, error) {
	records := append([]Record{}, scan.Records...)
	sortRecordsOldestFirst(records)
	protected := map[string]bool{}
	if opts.KeepLast != nil {
		newest := append([]Record{}, scan.Records...)
		sortRecordsNewestFirst(newest)
		for i, record := range newest {
			if i >= *opts.KeepLast {
				break
			}
			protected[record.ID] = true
		}
	}

	candidateByID := map[string]*CandidateEntry{}
	if opts.KeepLast != nil {
		for _, record := range records {
			if protected[record.ID] {
				continue
			}
			addRecordCandidate(candidateByID, record, "exceeds_keep_last")
		}
	}
	if opts.OlderThanSet || opts.OlderThan > 0 {
		cutoff := now.Add(-opts.OlderThan)
		for _, record := range records {
			if protected[record.ID] {
				continue
			}
			if record.CreatedAt.Before(cutoff) {
				addRecordCandidate(candidateByID, record, "older_than")
			}
		}
	}
	if opts.MaxTotalSizeSet || opts.MaxTotalSize > 0 {
		total := int64(0)
		for _, record := range records {
			total += record.SizeBytes
		}
		for _, record := range records {
			if total <= opts.MaxTotalSize {
				break
			}
			if protected[record.ID] {
				continue
			}
			addRecordCandidate(candidateByID, record, "exceeds_max_total_size")
			total -= record.SizeBytes
		}
	}

	candidates := make([]CandidateEntry, 0, len(candidateByID)+len(scan.Invalid)+len(scan.Legacy))
	for _, candidate := range candidateByID {
		candidates = append(candidates, *candidate)
	}
	if opts.IncludeInvalid {
		for _, entry := range scan.Invalid {
			if err := ensureManagedChild(rootDir, entry.Directory); err != nil {
				return nil, &Error{Code: ErrDeleteUnsafePath, Err: err}
			}
			size, err := directorySize(entry.Directory)
			if err != nil {
				return nil, &Error{Code: ErrPruneFailed, Err: err}
			}
			candidates = append(candidates, CandidateEntry{
				ID:        filepath.Base(entry.Directory),
				Directory: entry.Directory,
				SizeBytes: size,
				Reasons:   []string{"invalid_entry"},
				Status:    "invalid",
				Code:      entry.Code,
				Message:   entry.Message,
			})
		}
	}
	if opts.IncludeLegacy {
		for _, entry := range scan.Legacy {
			if err := ensureManagedChild(rootDir, entry.Directory); err != nil {
				return nil, &Error{Code: ErrDeleteUnsafePath, Err: err}
			}
			size, err := directorySize(entry.Directory)
			if err != nil {
				return nil, &Error{Code: ErrPruneFailed, Err: err}
			}
			candidates = append(candidates, CandidateEntry{
				ID:        filepath.Base(entry.Directory),
				Directory: entry.Directory,
				SizeBytes: size,
				Reasons:   []string{"legacy_entry"},
				Status:    "legacy",
			})
		}
	}
	sort.Slice(candidates, func(i, j int) bool {
		a := candidates[i]
		b := candidates[j]
		if !a.CreatedAt.Equal(b.CreatedAt) {
			if a.CreatedAt.IsZero() {
				return false
			}
			if b.CreatedAt.IsZero() {
				return true
			}
			return a.CreatedAt.Before(b.CreatedAt)
		}
		if a.Status != b.Status {
			return a.Status < b.Status
		}
		return a.ID < b.ID
	})
	return candidates, nil
}

func addRecordCandidate(candidates map[string]*CandidateEntry, record Record, reason string) {
	candidate := candidates[record.ID]
	if candidate == nil {
		candidate = &CandidateEntry{
			ID:        record.ID,
			Directory: record.Directory,
			CreatedAt: record.CreatedAt,
			Reason:    record.Reason,
			SizeBytes: record.SizeBytes,
			Status:    "valid",
		}
		candidates[record.ID] = candidate
	}
	for _, existing := range candidate.Reasons {
		if existing == reason {
			return
		}
	}
	candidate.Reasons = append(candidate.Reasons, reason)
	sort.Strings(candidate.Reasons)
}

func sortRecordsNewestFirst(records []Record) {
	sort.Slice(records, func(i, j int) bool {
		if records[i].CreatedAt.Equal(records[j].CreatedAt) {
			return records[i].ID > records[j].ID
		}
		return records[i].CreatedAt.After(records[j].CreatedAt)
	})
}

func sortRecordsOldestFirst(records []Record) {
	sort.Slice(records, func(i, j int) bool {
		if records[i].CreatedAt.Equal(records[j].CreatedAt) {
			return records[i].ID < records[j].ID
		}
		return records[i].CreatedAt.Before(records[j].CreatedAt)
	})
}

func deleteManagedDirectory(rootDir, id, dir string) (DeletedEntry, error) {
	if err := ensureManagedChild(rootDir, dir); err != nil {
		return DeletedEntry{}, err
	}
	size, err := directorySize(dir)
	if err != nil {
		return DeletedEntry{}, err
	}
	if err := removeAll(dir); err != nil {
		return DeletedEntry{}, &Error{Code: ErrDeleteFailed, Err: err}
	}
	return DeletedEntry{ID: id, Directory: dir, FreedBytes: size}, nil
}

func ensureManagedChild(rootDir, dir string) error {
	backupRootAbs, err := filepath.Abs(Root(rootDir))
	if err != nil {
		return &Error{Code: ErrDeleteUnsafePath, Err: err}
	}
	backupRootAbs = filepath.Clean(backupRootAbs)
	dirAbs, err := filepath.Abs(dir)
	if err != nil {
		return &Error{Code: ErrDeleteUnsafePath, Err: err}
	}
	dirAbs = filepath.Clean(dirAbs)
	if samePath(backupRootAbs, dirAbs) {
		return &Error{Code: ErrDeleteUnsafePath, Err: fmt.Errorf("backup root cannot be deleted")}
	}
	if !samePath(filepath.Dir(dirAbs), backupRootAbs) {
		return &Error{Code: ErrDeleteUnsafePath, Err: fmt.Errorf("backup target must be a direct child of the managed backup root")}
	}
	if !pathWithin(backupRootAbs, dirAbs) {
		return &Error{Code: ErrDeleteUnsafePath, Err: fmt.Errorf("backup target must stay inside the managed backup root")}
	}
	info, err := os.Lstat(dirAbs)
	if err != nil {
		return &Error{Code: ErrDeleteUnsafePath, Err: err}
	}
	if !info.IsDir() {
		return &Error{Code: ErrDeleteUnsafePath, Err: fmt.Errorf("backup target is not a directory")}
	}
	rootEval, err := filepath.EvalSymlinks(backupRootAbs)
	if err != nil {
		return &Error{Code: ErrDeleteUnsafePath, Err: err}
	}
	dirEval, err := filepath.EvalSymlinks(dirAbs)
	if err != nil {
		return &Error{Code: ErrDeleteUnsafePath, Err: err}
	}
	if !pathWithin(rootEval, dirEval) || samePath(rootEval, dirEval) {
		return &Error{Code: ErrDeleteUnsafePath, Err: fmt.Errorf("backup target resolves outside the managed backup root")}
	}
	return nil
}

func directorySize(dir string) (int64, error) {
	var total int64
	err := filepath.WalkDir(dir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if path != dir && d.Type()&os.ModeSymlink != 0 {
			return nil
		}
		if d.IsDir() {
			return nil
		}
		info, err := d.Info()
		if err != nil {
			return err
		}
		total += info.Size()
		return nil
	})
	return total, err
}

func codeForDeleteError(err error) string {
	var backupErr *Error
	if errors.As(err, &backupErr) && backupErr.Code != "" {
		return backupErr.Code
	}
	return ErrDeleteFailed
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
