// Package sourceencoding defines the encoding contract for tracked VBA source
// files and provides the source-only encoding commands.
package sourceencoding

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"unicode/utf8"

	"golang.org/x/text/encoding/japanese"
)

const (
	UTF8   = "utf-8"
	CP932  = "cp932"
	ToUTF8 = UTF8
)

type Status string

const (
	StatusValid        Status = "valid"
	StatusInvalidUTF8  Status = "invalid_utf8"
	StatusUTF8BOM      Status = "utf8_bom"
	StatusInvalidCP932 Status = "invalid_cp932"
	StatusUTF16BOM     Status = "utf16_bom"
	StatusConverted    Status = "converted"
	StatusUnchanged    Status = "unchanged"
)

// Position identifies a byte in a source file. Offset and ByteColumn are
// zero- and one-based respectively, matching the conventions used by the
// CLI's source diagnostics.
type Position struct {
	Offset     int `json:"offset"`
	Line       int `json:"line"`
	ByteColumn int `json:"byte_column"`
}

// Error is a validation error for one source file. It is intentionally typed
// so CLI callers can preserve the source-encoding error contract while still
// wrapping the original cause where one exists.
type Error struct {
	Path     string
	Status   Status
	Reason   string
	Position Position
	Cause    error
}

func (e *Error) Error() string {
	if e == nil {
		return "invalid VBA source encoding"
	}
	return fmt.Sprintf("%s: %s at byte offset %d (line %d, byte column %d)", e.Path, e.Reason, e.Position.Offset, e.Position.Line, e.Position.ByteColumn)
}

func (e *Error) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Cause
}

// Details returns the stable JSON-compatible details for the CLI envelope.
func (e *Error) Details() map[string]any {
	if e == nil {
		return nil
	}
	return map[string]any{
		"status":      string(e.Status),
		"reason":      e.Reason,
		"offset":      e.Position.Offset,
		"line":        e.Position.Line,
		"byte_column": e.Position.ByteColumn,
	}
}

// ValidationErrors retains every invalid source found by a project scan.
// The first error is also exposed through Error so errors.AsType can project a
// useful source path into the common CLI envelope.
type ValidationErrors struct {
	Errors []*Error
}

func (e *ValidationErrors) Error() string {
	if e == nil || len(e.Errors) == 0 {
		return "invalid VBA source encoding"
	}
	if len(e.Errors) == 1 {
		return e.Errors[0].Error()
	}
	return fmt.Sprintf("%d VBA source files have invalid encoding: %s", len(e.Errors), e.Errors[0].Error())
}

func (e *ValidationErrors) Unwrap() error {
	if e == nil || len(e.Errors) == 0 {
		return nil
	}
	return e.Errors[0]
}

type FileResult struct {
	Path   string `json:"path"`
	Status Status `json:"status"`
	Reason string `json:"reason,omitempty"`
	Position
}

type Summary struct {
	Total     int `json:"total"`
	Valid     int `json:"valid"`
	Invalid   int `json:"invalid"`
	Converted int `json:"converted"`
	Unchanged int `json:"unchanged"`
}

type Result struct {
	Expected string       `json:"expected"`
	From     string       `json:"from,omitempty"`
	To       string       `json:"to,omitempty"`
	Files    []FileResult `json:"files"`
	Summary  Summary      `json:"summary"`
}

// Roots is the set of project-managed source roots. The legacy tests root is
// always included by DiscoverFiles in addition to these configured roots.
type Roots struct {
	Modules  string
	Classes  string
	Forms    string
	Workbook string
}

type Options struct {
	RootDir string
	Roots   Roots
	Paths   []string
}

// Validate accepts only UTF-8 without a BOM. Unknown path extensions are
// allowed here because parser callers also use this function for in-memory
// snippets; filesystem discovery limits the command to VBA source extensions.
func Validate(path string, source []byte) error {
	if hasPrefix(source, []byte{0xef, 0xbb, 0xbf}) {
		return newEncodingError(path, StatusUTF8BOM, "UTF-8 BOM is not allowed", source, 0)
	}
	if hasPrefix(source, []byte{0xff, 0xfe}) || hasPrefix(source, []byte{0xfe, 0xff}) {
		return newEncodingError(path, StatusUTF16BOM, "UTF-16 BOM is not supported", source, 0)
	}
	if utf8.Valid(source) {
		return nil
	}
	return newEncodingError(path, StatusInvalidUTF8, "source is not valid UTF-8", source, firstInvalidUTF8(source))
}

// Inspect reports the required source status without attempting conversion.
func Inspect(path string, source []byte) FileResult {
	if hasPrefix(source, []byte{0xef, 0xbb, 0xbf}) {
		return resultForError(path, newEncodingError(path, StatusUTF8BOM, "UTF-8 BOM is not allowed", source, 0))
	}
	if hasPrefix(source, []byte{0xff, 0xfe}) || hasPrefix(source, []byte{0xfe, 0xff}) {
		return resultForError(path, newEncodingError(path, StatusUTF16BOM, "UTF-16 BOM is not supported", source, 0))
	}
	if offset := firstInvalidUTF8(source); offset >= 0 {
		return resultForError(path, newEncodingError(path, StatusInvalidUTF8, "source is not valid UTF-8", source, offset))
	}
	return FileResult{Path: path, Status: StatusValid}
}

// DecodeCP932 strictly decodes Windows Code Page 932. The standard x/text
// decoder replaces malformed sequences with U+FFFD, so a small validation
// pass rejects those sequences before the decoder is allowed to produce the
// converted output.
func DecodeCP932(path string, source []byte) ([]byte, error) {
	if hasPrefix(source, []byte{0xef, 0xbb, 0xbf}) {
		return nil, newEncodingError(path, StatusUTF8BOM, "UTF-8 BOM is not allowed", source, 0)
	}
	if hasPrefix(source, []byte{0xff, 0xfe}) || hasPrefix(source, []byte{0xfe, 0xff}) {
		return nil, newEncodingError(path, StatusUTF16BOM, "UTF-16 BOM is not supported", source, 0)
	}
	if offset := firstInvalidCP932(source); offset >= 0 {
		return nil, newEncodingError(path, StatusInvalidCP932, "source is not valid CP932", source, offset)
	}
	decoded, err := japanese.ShiftJIS.NewDecoder().Bytes(source)
	if err != nil {
		return nil, newEncodingError(path, StatusInvalidCP932, "source is not valid CP932", source, 0).withCause(err)
	}
	return decoded, nil
}

// Check scans all selected files and returns every encoding violation. File
// paths in the result are project-relative and sorted deterministically.
func Check(ctx context.Context, opts Options) (Result, error) {
	files, err := DiscoverFiles(ctx, opts)
	if err != nil {
		return Result{Expected: UTF8, Files: []FileResult{}}, err
	}
	result := Result{Expected: UTF8, Files: make([]FileResult, 0, len(files))}
	var invalid []*Error
	for _, file := range files {
		if err := contextError(ctx); err != nil {
			return Result{}, err
		}
		body, readErr := os.ReadFile(file.absolutePath)
		if readErr != nil {
			return result, &FileIOError{Path: file.Path, Operation: "read", Cause: readErr}
		}
		item := Inspect(file.Path, body)
		result.Files = append(result.Files, item)
		if item.Status != StatusValid {
			invalid = append(invalid, newEncodingError(file.Path, item.Status, item.Reason, body, item.Offset))
		}
	}
	result.Summary = summarize(result.Files)
	if len(invalid) > 0 {
		return result, &ValidationErrors{Errors: invalid}
	}
	return result, nil
}

// Convert converts only non-UTF-8 source files. It stages every converted
// file before changing any original and retains temporary originals until all
// replacements have succeeded, allowing a failed transaction to roll back.
func Convert(ctx context.Context, opts Options, from string) (Result, error) {
	from = strings.ToLower(strings.TrimSpace(from))
	result := Result{Expected: UTF8, From: from, To: UTF8, Files: []FileResult{}}
	if from != CP932 {
		return result, &ScopeError{Path: from, Reason: "--from must be cp932"}
	}
	files, err := DiscoverFiles(ctx, opts)
	if err != nil {
		return result, err
	}
	pending := make([]*stagedFile, 0, len(files))
	var invalid []*Error
	for _, file := range files {
		if err := contextError(ctx); err != nil {
			return Result{}, err
		}
		body, readErr := os.ReadFile(file.absolutePath)
		if readErr != nil {
			return result, &FileIOError{Path: file.Path, Operation: "read", Cause: readErr}
		}
		info, statErr := os.Stat(file.absolutePath)
		if statErr != nil {
			return result, &FileIOError{Path: file.Path, Operation: "stat", Cause: statErr}
		}
		item := &stagedFile{file: file, path: file.canonical, body: body, mode: info.Mode()}
		if hasPrefix(body, []byte{0xef, 0xbb, 0xbf}) {
			item.result = resultForError(file.Path, newEncodingError(file.Path, StatusUTF8BOM, "UTF-8 BOM is not allowed", body, 0))
			invalid = append(invalid, newEncodingError(file.Path, item.result.Status, item.result.Reason, body, 0))
		} else if hasPrefix(body, []byte{0xff, 0xfe}) || hasPrefix(body, []byte{0xfe, 0xff}) {
			item.result = resultForError(file.Path, newEncodingError(file.Path, StatusUTF16BOM, "UTF-16 BOM is not supported", body, 0))
			invalid = append(invalid, newEncodingError(file.Path, item.result.Status, item.result.Reason, body, 0))
		} else if utf8.Valid(body) {
			item.result = FileResult{Path: file.Path, Status: StatusUnchanged}
		} else {
			decoded, decodeErr := DecodeCP932(file.Path, body)
			if decodeErr != nil {
				item.result = resultForError(file.Path, decodeErr)
				var encodingErr *Error
				if encodingErr, _ = errors.AsType[*Error](decodeErr); encodingErr != nil {
					invalid = append(invalid, encodingErr)
				}
			} else {
				item.result = FileResult{Path: file.Path, Status: StatusConverted}
				item.data = decoded
			}
		}
		pending = append(pending, item)
		result.Files = append(result.Files, item.result)
	}
	result.Summary = summarize(result.Files)
	if len(invalid) > 0 {
		return result, &ValidationErrors{Errors: invalid}
	}

	// Stage all replacements and probe each containing directory before the
	// first original is renamed.
	for _, item := range pending {
		if item.result.Status != StatusConverted {
			continue
		}
		if err := ensureDirectoryWritable(item.path); err != nil {
			cleanupPending(pending)
			return result, &FileIOError{Path: item.file.Path, Operation: "write", Cause: err}
		}
		temp, stageErr := stageFile(item.path, item.data, item.mode)
		if stageErr != nil {
			cleanupPending(pending)
			return result, &FileIOError{Path: item.file.Path, Operation: "write", Cause: stageErr}
		}
		item.temp = temp
	}

	for _, item := range pending {
		if item.result.Status != StatusConverted {
			continue
		}
		backup, commitErr := commitFile(item)
		item.backup = backup
		if commitErr != nil {
			rollbackErr := rollbackPending(pending)
			return result, &TransactionError{Path: item.file.Path, Cause: commitErr, Rollback: rollbackErr}
		}
		item.applied = true
	}
	var cleanupErrs []error
	for _, item := range pending {
		if item.temp != "" {
			if removeErr := os.Remove(item.temp); removeErr != nil && !errors.Is(removeErr, fs.ErrNotExist) {
				cleanupErrs = append(cleanupErrs, fmt.Errorf("remove temporary file %s: %w", item.temp, removeErr))
			}
		}
		if item.backup != "" {
			if removeErr := os.Remove(item.backup); removeErr != nil && !errors.Is(removeErr, fs.ErrNotExist) {
				cleanupErrs = append(cleanupErrs, fmt.Errorf("remove conversion backup %s: %w", item.backup, removeErr))
			}
		}
	}
	if len(cleanupErrs) > 0 {
		return result, &TransactionError{Path: pending[0].file.Path, Cause: errors.Join(cleanupErrs...)}
	}
	return result, nil
}

// File identifies one discovered project-relative VBA source file.
type File struct {
	absolutePath string
	canonical    string
	Path         string
}

type stagedFile struct {
	file       File
	path       string
	body       []byte
	result     FileResult
	data       []byte
	temp       string
	backupTemp string
	backup     string
	mode       fs.FileMode
	applied    bool
}

// DiscoverFiles resolves configured roots and optional explicit paths. It
// follows file symlinks only when their resolved target remains within a
// project-managed root, and rejects all escaping links/junctions.
func DiscoverFiles(ctx context.Context, opts Options) ([]File, error) {
	if err := contextError(ctx); err != nil {
		return nil, err
	}
	rootDir := opts.RootDir
	if strings.TrimSpace(rootDir) == "" {
		rootDir = "."
	}
	rootAbs, err := filepath.Abs(rootDir)
	if err != nil {
		return nil, err
	}
	rootAbs = filepath.Clean(rootAbs)
	rootCanonical, err := filepath.EvalSymlinks(rootAbs)
	if err != nil {
		return nil, &ScopeError{Path: rootAbs, Reason: "project root could not be resolved", Cause: err}
	}
	rootCanonical = filepath.Clean(rootCanonical)

	configured := []string{opts.Roots.Modules, opts.Roots.Classes, opts.Roots.Forms, opts.Roots.Workbook, "tests"}
	managed := make([]string, 0, len(configured))
	for _, raw := range configured {
		if strings.TrimSpace(raw) == "" {
			continue
		}
		path := raw
		if !filepath.IsAbs(path) {
			path = filepath.Join(rootAbs, path)
		}
		path = filepath.Clean(path)
		info, statErr := os.Stat(path)
		if statErr != nil {
			if errors.Is(statErr, fs.ErrNotExist) {
				continue
			}
			return nil, &FileIOError{Path: path, Operation: "discover", Cause: statErr}
		}
		canonical, evalErr := filepath.EvalSymlinks(path)
		if evalErr != nil {
			return nil, &ScopeError{Path: path, Reason: "managed source root could not be resolved", Cause: evalErr}
		}
		if !pathInside(rootCanonical, canonical) {
			return nil, &ScopeError{Path: path, Reason: "managed source root escapes the project root"}
		}
		if info.IsDir() {
			managed = append(managed, path)
		} else if isSourceExtension(path) {
			managed = append(managed, path)
		} else {
			return nil, &ScopeError{Path: path, Reason: "managed source root is not a VBA source file or directory"}
		}
	}

	requested := opts.Paths
	if len(requested) == 0 {
		requested = managed
	}
	candidates := make([]File, 0)
	for _, raw := range requested {
		if err := contextError(ctx); err != nil {
			return nil, err
		}
		path := raw
		if !filepath.IsAbs(path) {
			path = filepath.Join(rootAbs, path)
		}
		path = filepath.Clean(path)
		if !pathLexicallyManaged(path, managed) {
			return nil, &ScopeError{Path: raw, Reason: "path is outside the configured project source roots"}
		}
		info, statErr := os.Stat(path)
		if statErr != nil {
			return nil, &FileIOError{Path: raw, Operation: "discover", Cause: statErr}
		}
		if !pathManaged(path, managed) {
			return nil, &ScopeError{Path: raw, Reason: "path is outside the configured project source roots"}
		}
		if info.IsDir() {
			if walkErr := filepath.WalkDir(path, func(candidate string, entry os.DirEntry, err error) error {
				if err != nil {
					return err
				}
				if contextErr := contextError(ctx); contextErr != nil {
					return contextErr
				}
				if entry.Type()&os.ModeSymlink != 0 {
					canonical, evalErr := filepath.EvalSymlinks(candidate)
					if evalErr != nil {
						return &ScopeError{Path: candidate, Reason: "source link could not be resolved", Cause: evalErr}
					}
					if !pathInside(rootCanonical, canonical) || !pathManagedCanonical(canonical, managed) {
						return &ScopeError{Path: candidate, Reason: "source link escapes the configured project source roots"}
					}
				}
				if entry.IsDir() {
					return nil
				}
				if !isSourceExtension(candidate) {
					return nil
				}
				item, itemErr := discovered(rootAbs, rootCanonical, candidate, managed)
				if itemErr != nil {
					return itemErr
				}
				candidates = append(candidates, item)
				return nil
			}); walkErr != nil {
				return nil, &FileIOError{Path: raw, Operation: "discover", Cause: walkErr}
			}
			continue
		}
		if !isSourceExtension(path) {
			return nil, &ScopeError{Path: raw, Reason: "path must have a .bas, .cls, or .frm extension"}
		}
		item, itemErr := discovered(rootAbs, rootCanonical, path, managed)
		if itemErr != nil {
			return nil, itemErr
		}
		candidates = append(candidates, item)
	}

	sort.Slice(candidates, func(i, j int) bool {
		left, right := candidates[i].Path, candidates[j].Path
		if runtime.GOOS == "windows" {
			return strings.ToLower(left) < strings.ToLower(right)
		}
		return left < right
	})
	deduped := candidates[:0]
	seen := make(map[string]struct{}, len(candidates))
	for _, item := range candidates {
		key := item.canonical
		if runtime.GOOS == "windows" {
			key = strings.ToLower(key)
		}
		if _, exists := seen[key]; exists {
			continue
		}
		seen[key] = struct{}{}
		deduped = append(deduped, item)
	}
	return deduped, nil
}

type ScopeError struct {
	Path   string
	Reason string
	Cause  error
}

func (e *ScopeError) Error() string {
	if e == nil {
		return "invalid source path"
	}
	if e.Path == "" {
		return e.Reason
	}
	return fmt.Sprintf("%s: %s", e.Path, e.Reason)
}

func (e *ScopeError) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Cause
}

type FileIOError struct {
	Path      string
	Operation string
	Cause     error
}

func (e *FileIOError) Error() string {
	if e == nil {
		return "source file I/O failed"
	}
	return fmt.Sprintf("%s %s failed: %v", e.Operation, e.Path, e.Cause)
}

func (e *FileIOError) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Cause
}

type TransactionError struct {
	Path     string
	Cause    error
	Rollback error
}

func (e *TransactionError) Error() string {
	if e == nil {
		return "source conversion transaction failed"
	}
	if e.Rollback != nil {
		return fmt.Sprintf("conversion of %s failed: %v; rollback failed: %v", e.Path, e.Cause, e.Rollback)
	}
	return fmt.Sprintf("conversion of %s failed: %v", e.Path, e.Cause)
}

func (e *TransactionError) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Cause
}

func newEncodingError(path string, status Status, reason string, source []byte, offset int) *Error {
	return &Error{Path: path, Status: status, Reason: reason, Position: sourcePosition(source, offset)}
}

func (e *Error) withCause(cause error) *Error {
	e.Cause = cause
	return e
}

func resultForError(path string, err error) FileResult {
	result := FileResult{Path: path, Status: StatusInvalidUTF8}
	encodingErr, ok := errors.AsType[*Error](err)
	if !ok || encodingErr == nil {
		result.Reason = err.Error()
		return result
	}
	result.Status = encodingErr.Status
	result.Reason = encodingErr.Reason
	result.Position = encodingErr.Position
	return result
}

func summarize(files []FileResult) Summary {
	result := Summary{Total: len(files)}
	for _, file := range files {
		switch file.Status {
		case StatusValid:
			result.Valid++
		case StatusConverted:
			result.Converted++
		case StatusUnchanged:
			result.Unchanged++
		default:
			result.Invalid++
		}
	}
	return result
}

func contextError(ctx context.Context) error {
	if ctx == nil {
		return nil
	}
	return ctx.Err()
}

func hasPrefix(source, prefix []byte) bool {
	return bytes.HasPrefix(source, prefix)
}

func firstInvalidUTF8(source []byte) int {
	for offset := 0; offset < len(source); {
		r, size := utf8.DecodeRune(source[offset:])
		if r == utf8.RuneError && size == 1 {
			return offset
		}
		offset += size
	}
	return -1
}

func firstInvalidCP932(source []byte) int {
	for offset := 0; offset < len(source); {
		size := 1
		if isCP932Lead(source[offset]) {
			size = 2
		}
		end := min(offset+size, len(source))
		decoded, err := japanese.ShiftJIS.NewDecoder().Bytes(source[offset:end])
		if err != nil || strings.ContainsRune(string(decoded), utf8.RuneError) {
			return offset
		}
		offset = end
	}
	return -1
}

func isCP932Lead(value byte) bool {
	return (0x81 <= value && value <= 0x9f) || (0xe0 <= value && value <= 0xfc)
}

func sourcePosition(source []byte, offset int) Position {
	if offset < 0 {
		offset = 0
	}
	if offset > len(source) {
		offset = len(source)
	}
	line, column := 1, 1
	for index := 0; index < offset; index++ {
		switch source[index] {
		case '\r':
			line++
			column = 1
			if index+1 < offset && source[index+1] == '\n' {
				index++
			}
		case '\n':
			line++
			column = 1
		default:
			column++
		}
	}
	return Position{Offset: offset, Line: line, ByteColumn: column}
}

func isSourceExtension(path string) bool {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".bas", ".cls", ".frm":
		return true
	default:
		return false
	}
}

func pathInside(root, path string) bool {
	relative, err := filepath.Rel(root, path)
	if err != nil {
		return false
	}
	if relative == "." {
		return true
	}
	if runtime.GOOS == "windows" {
		relative = strings.ToLower(relative)
	}
	return relative != ".." && !strings.HasPrefix(relative, ".."+string(os.PathSeparator)) && !filepath.IsAbs(relative)
}

func pathManaged(path string, managed []string) bool {
	canonical, err := filepath.EvalSymlinks(path)
	if err != nil {
		return false
	}
	for _, root := range managed {
		rootCanonical, rootErr := filepath.EvalSymlinks(root)
		if rootErr == nil && pathInside(rootCanonical, canonical) {
			return true
		}
	}
	return false
}

func pathLexicallyManaged(path string, managed []string) bool {
	for _, root := range managed {
		if pathInside(root, path) {
			return true
		}
	}
	return false
}

func discovered(rootAbs, rootCanonical, path string, managed []string) (File, error) {
	canonical, err := filepath.EvalSymlinks(path)
	if err != nil {
		return File{}, &ScopeError{Path: path, Reason: "source path could not be resolved", Cause: err}
	}
	if !pathInside(rootCanonical, canonical) || !pathManagedCanonical(canonical, managed) {
		return File{}, &ScopeError{Path: path, Reason: "source path escapes the configured project source roots"}
	}
	relative, err := filepath.Rel(rootAbs, path)
	if err != nil || filepath.IsAbs(relative) || strings.HasPrefix(relative, ".."+string(os.PathSeparator)) || relative == ".." {
		return File{}, &ScopeError{Path: path, Reason: "source path is outside the project root"}
	}
	return File{absolutePath: path, Path: filepath.ToSlash(relative), canonical: filepath.Clean(canonical)}, nil
}

func pathManagedCanonical(path string, managed []string) bool {
	for _, root := range managed {
		rootCanonical, err := filepath.EvalSymlinks(root)
		if err == nil && pathInside(rootCanonical, path) {
			return true
		}
	}
	return false
}

func ensureDirectoryWritable(path string) error {
	directory := filepath.Dir(path)
	temp, err := os.CreateTemp(directory, ".xlflow-encoding-probe-*")
	if err != nil {
		return err
	}
	tempPath := temp.Name()
	if closeErr := temp.Close(); closeErr != nil {
		_ = os.Remove(tempPath)
		return closeErr
	}
	return os.Remove(tempPath)
}

func stageFile(path string, data []byte, mode fs.FileMode) (string, error) {
	temp, err := os.CreateTemp(filepath.Dir(path), "."+filepath.Base(path)+".xlflow-encoding-*")
	if err != nil {
		return "", err
	}
	tempPath := temp.Name()
	defer func() {
		if err != nil {
			_ = os.Remove(tempPath)
		}
	}()
	if err = temp.Chmod(mode.Perm()); err != nil {
		_ = temp.Close()
		return "", err
	}
	if _, err = temp.Write(data); err != nil {
		_ = temp.Close()
		return "", err
	}
	if err = temp.Sync(); err != nil {
		_ = temp.Close()
		return "", err
	}
	if err = temp.Close(); err != nil {
		return "", err
	}
	return tempPath, nil
}

func commitFile(item *stagedFile) (string, error) {
	backupFile, err := os.CreateTemp(filepath.Dir(item.path), "."+filepath.Base(item.path)+".xlflow-encoding-backup-*")
	if err != nil {
		return "", err
	}
	backup := backupFile.Name()
	item.backupTemp = backup
	if err = backupFile.Close(); err != nil {
		return "", err
	}
	if err = os.Remove(backup); err != nil {
		_ = os.Remove(backup)
		return "", err
	}
	item.backupTemp = ""
	if err = os.Rename(item.path, backup); err != nil {
		return "", err
	}
	if err = os.Rename(item.temp, item.path); err != nil {
		return backup, err
	}
	return backup, nil
}

func rollbackPending(pending []*stagedFile) error {
	var rollbackErrors []error
	for index := len(pending) - 1; index >= 0; index-- {
		item := pending[index]
		if item.backup == "" {
			if item.backupTemp != "" {
				if err := os.Remove(item.backupTemp); err != nil && !errors.Is(err, fs.ErrNotExist) {
					rollbackErrors = append(rollbackErrors, err)
				}
			}
			if item.temp != "" {
				if err := os.Remove(item.temp); err != nil && !errors.Is(err, fs.ErrNotExist) {
					rollbackErrors = append(rollbackErrors, err)
				}
			}
			continue
		}
		if item.applied {
			if err := os.Remove(item.path); err != nil && !errors.Is(err, fs.ErrNotExist) {
				rollbackErrors = append(rollbackErrors, err)
				continue
			}
		}
		if err := os.Rename(item.backup, item.path); err != nil {
			rollbackErrors = append(rollbackErrors, err)
		}
		if item.temp != "" {
			_ = os.Remove(item.temp)
		}
		if item.backupTemp != "" {
			_ = os.Remove(item.backupTemp)
		}
	}
	return errors.Join(rollbackErrors...)
}

func cleanupPending(pending []*stagedFile) {
	for _, item := range pending {
		if item.temp != "" {
			_ = os.Remove(item.temp)
		}
		if item.backup != "" {
			_ = os.Remove(item.backup)
		}
		if item.backupTemp != "" {
			_ = os.Remove(item.backupTemp)
		}
	}
}
