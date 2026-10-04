package filepush

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/wsl"
)

// dotnetTicksPerUnixNano converts a Unix-nanosecond timestamp into .NET
// DateTime ticks: 621355968000000000 is 1970-01-01T00:00:00Z in ticks.
const dotnetTicksEpochOffset = 621355968000000000

// The push-state schema is shared with the .NET Excel bridge so a state file
// written by either backend can be consumed by the other. Field order and the
// snake_case names match System.Text.Json's serialized output.
type sourceFileEntry struct {
	Kind string `json:"kind"`
	Path string `json:"path"`
	Hash string `json:"hash"`
}

type sourceFingerprint struct {
	WorkbookPath       string            `json:"workbook_path"`
	Files              []sourceFileEntry `json:"files"`
	LineNumbersEnabled bool              `json:"line_numbers_enabled"`
	// FolderAnnotation records the effective [vba.folder_annotation] mode so
	// a mode change without source edits invalidates --changed-only. States
	// written before this field existed compare unequal and re-push, which
	// is the safe direction.
	FolderAnnotation string `json:"folder_annotation"`
}

type pushSavedFile struct {
	Path                  string `json:"path"`
	LastWriteTimeUTCTicks int64  `json:"last_write_time_utc_ticks"`
	Length                int64  `json:"length"`
}

type pushAppliedTo struct {
	SessionID   string         `json:"session_id"`
	SessionPid  int            `json:"session_pid"`
	SessionHwnd int64          `json:"session_hwnd"`
	SavedFile   *pushSavedFile `json:"saved_file"`
}

type pushState struct {
	Fingerprint sourceFingerprint `json:"fingerprint"`
	AppliedTo   *pushAppliedTo    `json:"applied_to"`
}

// runningOnWSL is stubbed by tests.
var runningOnWSL = wsl.IsWSL

// normalizeFingerprintPath canonicalizes a workbook path for push state so a
// fingerprint or saved-file stamp written on one OS still matches after the
// other backend sees the same file across a WSL/Windows boundary: a WSL
// absolute path under /mnt/<drive> and the Windows absolute path for the same
// file both normalize to the Windows form (`D:\...`). Windows-form input is
// recognized before host absolutization because filepath.Abs("C:\\x") on a
// Unix host would produce garbage. Non-WSL hosts never rewrite /mnt/ paths —
// a bare Linux mount cannot be proven to be the same file Windows sees.
func normalizeFingerprintPath(path string) string {
	if windowsPath, ok := windowsDrivePath(path); ok {
		return windowsPath
	}
	// Check the raw absolute /mnt/<drive>/... form before host absolutization:
	// under WSL filepath.Abs preserves it, and checking first keeps the
	// mapping testable on non-WSL hosts too.
	if runningOnWSL() {
		if windowsPath, ok := wslMountToWindowsPath(path); ok {
			return windowsPath
		}
	}
	abs, err := filepath.Abs(path)
	if err != nil {
		return filepath.Clean(path)
	}
	abs = filepath.Clean(abs)
	if runningOnWSL() {
		if windowsPath, ok := wslMountToWindowsPath(abs); ok {
			return windowsPath
		}
	}
	return abs
}

// windowsDrivePath reports the canonical `D:\...` form when path already looks
// like a Windows drive-rooted path (X:\ or X:/), or false otherwise.
func windowsDrivePath(path string) (string, bool) {
	p := strings.TrimSpace(path)
	if len(p) < 3 || p[1] != ':' || !isASCIILetter(p[0]) || (p[2] != '\\' && p[2] != '/') {
		return "", false
	}
	cleaned := filepath.ToSlash(filepath.Clean(strings.ReplaceAll(p, "\\", "/")))
	rest := ""
	if len(cleaned) > 2 {
		rest = strings.ReplaceAll(cleaned[2:], "/", "\\")
	}
	return strings.ToUpper(p[:1]) + ":" + rest, true
}

// wslMountToWindowsPath maps an absolute /mnt/<drive>/... path to its Windows
// form `D:\...`. Non-/mnt paths report false so Linux-only project paths keep
// their native form in state — Windows cannot see them anyway.
func wslMountToWindowsPath(abs string) (string, bool) {
	slashed := filepath.ToSlash(abs)
	if !strings.HasPrefix(slashed, "/mnt/") || len(slashed) < len("/mnt/c") {
		return "", false
	}
	rest := slashed[len("/mnt/"):]
	if !isASCIILetter(rest[0]) || (len(rest) > 1 && rest[1] != '/') {
		return "", false
	}
	drive := strings.ToUpper(rest[:1])
	if len(rest) == 1 {
		return drive + `:\`, true
	}
	return drive + ":" + strings.ReplaceAll(rest[1:], "/", "\\"), true
}

func isASCIILetter(b byte) bool {
	return (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z')
}

// computeFingerprint ports ComputeFingerprint: one entry per discovered file,
// the normalized workbook path, and the transform toggles (line numbers,
// folder annotation mode) that change the effective imported text.
func computeFingerprint(workbookPath string, files []discoveredFile, lineNumbers bool, folderAnnotation string) (sourceFingerprint, error) {
	entries := make([]sourceFileEntry, 0, len(files))
	for _, file := range files {
		body := file.Body
		if !file.HasBody {
			var err error
			body, err = os.ReadFile(file.FullPath)
			if err != nil {
				return sourceFingerprint{}, &SourceReadError{Path: file.FullPath, Err: err}
			}
		}
		sum := sha256.Sum256(body)
		entries = append(entries, sourceFileEntry{
			Kind: file.Kind,
			Path: file.RelativePath,
			Hash: hex.EncodeToString(sum[:]),
		})
	}
	return sourceFingerprint{
		WorkbookPath:       normalizeFingerprintPath(workbookPath),
		Files:              entries,
		LineNumbersEnabled: lineNumbers,
		FolderAnnotation:   folderAnnotation,
	}, nil
}

// fingerprintEquals compares fingerprints as multisets of (kind, path, hash)
// entries. The .NET comparer re-serializes both fingerprints to compact JSON,
// which is order-sensitive; this comparer is deliberately order-insensitive so
// a state file produced by the bridge's filesystem enumeration order still
// matches a Go enumeration. When the sets differ only in order the bridge
// would re-push, which is always safe.
func fingerprintEquals(left, right sourceFingerprint) bool {
	if normalizeFingerprintPath(left.WorkbookPath) != normalizeFingerprintPath(right.WorkbookPath) {
		return false
	}
	if left.LineNumbersEnabled != right.LineNumbersEnabled {
		return false
	}
	if left.FolderAnnotation != right.FolderAnnotation {
		return false
	}
	if len(left.Files) != len(right.Files) {
		return false
	}
	key := func(f sourceFileEntry) string { return f.Kind + "\x00" + f.Path + "\x00" + f.Hash }
	counts := make(map[string]int, len(left.Files))
	for _, file := range left.Files {
		counts[key(file)]++
	}
	for _, file := range right.Files {
		k := key(file)
		if counts[k] == 0 {
			return false
		}
		counts[k]--
	}
	return true
}

// readPushState ports TryReadPushState, including the legacy bare-fingerprint
// shape, which can never justify a changed-only skip because it carries no
// delivery evidence.
func readPushState(statePath string) (*pushState, error) {
	body, err := os.ReadFile(statePath)
	if errors.Is(err, os.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}
	body = bytes.TrimPrefix(body, []byte{0xEF, 0xBB, 0xBF})
	var raw map[string]json.RawMessage
	if err := json.Unmarshal(body, &raw); err != nil {
		return nil, err
	}
	state := &pushState{}
	fingerprintJSON, hasFingerprint := raw["fingerprint"]
	if !hasFingerprint {
		// Legacy shape: the root object is a bare SourceFingerprint.
		fingerprintJSON = body
	}
	if err := json.Unmarshal(fingerprintJSON, &state.Fingerprint); err != nil {
		return nil, err
	}
	if appliedJSON, ok := raw["applied_to"]; ok && len(appliedJSON) > 0 && string(appliedJSON) != "null" {
		var applied pushAppliedTo
		if err := json.Unmarshal(appliedJSON, &applied); err != nil {
			return nil, err
		}
		state.AppliedTo = &applied
	}
	return state, nil
}

// savedFileStampMatches ports SavedFileStampMatches: the recorded absolute
// path, last-write ticks, and length must all still describe the workbook.
func savedFileStampMatches(saved *pushSavedFile, workbookPath string) bool {
	if saved == nil || saved.Path == "" || !pathsEqual(saved.Path, workbookPath) {
		return false
	}
	info, err := os.Stat(workbookPath)
	if err != nil {
		return false
	}
	return fileWriteTicks(info) == saved.LastWriteTimeUTCTicks && info.Size() == saved.Length
}

func fileWriteTicks(info os.FileInfo) int64 {
	mod := info.ModTime().UTC()
	return dotnetTicksEpochOffset + mod.UnixNano()/100
}

func pathsEqual(left, right string) bool {
	return strings.EqualFold(normalizeFingerprintPath(left), normalizeFingerprintPath(right))
}

// shouldSkipUnchanged ports TrySkipUnchangedImport for the file backend: the
// recorded fingerprint must match and applied_to.saved_file must still stamp
// the workbook on disk. Session coverage is never claimed here because the
// file backend refuses to run while a matching session is live — that guard
// runs in the CLI before mutation, while a pure skip carries no write risk.
func shouldSkipUnchanged(statePath string, fingerprint sourceFingerprint, workbookPath string) bool {
	state, err := readPushState(statePath)
	if err != nil || state == nil || state.AppliedTo == nil {
		return false
	}
	if !fingerprintEquals(fingerprint, state.Fingerprint) {
		return false
	}
	return savedFileStampMatches(state.AppliedTo.SavedFile, workbookPath)
}

// writePushState writes push.json through the same atomic publish primitive
// used for the workbook so a crash cannot leave a torn state file.
func writePushState(statePath string, fingerprint sourceFingerprint, applied *pushAppliedTo) error {
	if err := os.MkdirAll(filepath.Dir(statePath), 0o755); err != nil {
		return err
	}
	body, err := json.MarshalIndent(pushState{Fingerprint: fingerprint, AppliedTo: applied}, "", "  ")
	if err != nil {
		return err
	}
	_, err = coordination.PublishFile(statePath, body, nil)
	return err
}

// buildFilePushAppliedTo mirrors BuildAppliedTo for a saved-file push: no
// session identity is claimed, and the saved-file stamp is read from the
// published artifact so the next changed-only run can verify coverage.
func buildFilePushAppliedTo(workbookPath string) *pushAppliedTo {
	applied := &pushAppliedTo{}
	info, err := os.Stat(workbookPath)
	if err != nil {
		return applied
	}
	applied.SavedFile = &pushSavedFile{
		Path:                  normalizeFingerprintPath(workbookPath),
		LastWriteTimeUTCTicks: fileWriteTicks(info),
		Length:                info.Size(),
	}
	return applied
}
