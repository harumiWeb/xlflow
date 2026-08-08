package corpus

import (
	"bufio"
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	pathpkg "path"
	"path/filepath"
	"sort"
	"strings"
)

// SnapshotDiagnostic is the stable, prose-free identity persisted in a
// corpus snapshot. Line and column are one-based source positions; column 0
// means that the analyzer did not provide a column.
type SnapshotDiagnostic struct {
	Project  string  `json:"project"`
	File     string  `json:"file"`
	Surface  Surface `json:"surface"`
	Code     string  `json:"code"`
	Severity string  `json:"severity"`
	Line     int     `json:"line"`
	Column   int     `json:"column"`
}

// SnapshotID identifies one project/surface snapshot file.
type SnapshotID struct {
	Project string
	Surface Surface
}

// SnapshotSet contains all normalized diagnostics keyed by project/surface.
// Each slice is sorted canonically and retains duplicate rows.
type SnapshotSet map[SnapshotID][]SnapshotDiagnostic

// SnapshotDiff is a multiset difference. Repeated equal rows are represented
// repeatedly in Added or Removed, so multiplicity changes are observable.
type SnapshotDiff struct {
	Added   []SnapshotDiagnostic
	Removed []SnapshotDiagnostic
}

func (d SnapshotDiff) Empty() bool { return len(d.Added) == 0 && len(d.Removed) == 0 }

func snapshotIDLess(a, b SnapshotID) bool {
	if a.Project != b.Project {
		return a.Project < b.Project
	}
	return surfaceRank(a.Surface) < surfaceRank(b.Surface)
}

func snapshotDiagnosticLess(a, b SnapshotDiagnostic) bool {
	if a.Project != b.Project {
		return a.Project < b.Project
	}
	if surfaceRank(a.Surface) != surfaceRank(b.Surface) {
		return surfaceRank(a.Surface) < surfaceRank(b.Surface)
	}
	if a.File != b.File {
		return a.File < b.File
	}
	if a.Line != b.Line {
		return a.Line < b.Line
	}
	if a.Column != b.Column {
		return a.Column < b.Column
	}
	if a.Code != b.Code {
		return a.Code < b.Code
	}
	return a.Severity < b.Severity
}

func validateSnapshotID(id SnapshotID) error {
	if id.Surface != SurfaceLint && id.Surface != SurfaceAnalyze {
		return fmt.Errorf("unsupported snapshot surface %q", id.Surface)
	}
	project := id.Project
	if project == "" || strings.TrimSpace(project) != project || strings.Contains(project, "\\") || strings.Contains(project, ":") {
		return fmt.Errorf("invalid snapshot project %q", project)
	}
	clean := pathpkg.Clean(project)
	if clean != project || pathpkg.IsAbs(project) || project == "." || strings.HasPrefix(project, "../") || project == ".." {
		return fmt.Errorf("invalid snapshot project path %q", project)
	}
	for _, segment := range strings.Split(project, "/") {
		if segment == "" || segment == "." || segment == ".." {
			return fmt.Errorf("invalid snapshot project path %q", project)
		}
	}
	return nil
}

func validateSnapshotDiagnostic(d SnapshotDiagnostic) error {
	if err := validateSnapshotID(SnapshotID{Project: d.Project, Surface: d.Surface}); err != nil {
		return err
	}
	if d.File == "" || strings.TrimSpace(d.File) != d.File || strings.Contains(d.File, "\\") || strings.Contains(d.File, ":") {
		return fmt.Errorf("invalid snapshot file %q", d.File)
	}
	file := pathpkg.Clean(d.File)
	if file != d.File || pathpkg.IsAbs(d.File) || file == "." || file == ".." || strings.HasPrefix(file, "../") {
		return fmt.Errorf("invalid snapshot file path %q", d.File)
	}
	for _, segment := range strings.Split(d.File, "/") {
		if segment == "" || segment == "." || segment == ".." {
			return fmt.Errorf("invalid snapshot file path %q", d.File)
		}
	}
	if d.Code == "" || d.Severity == "" {
		return errors.New("snapshot diagnostic code and severity are required")
	}
	if d.Line < 1 {
		return fmt.Errorf("snapshot diagnostic line must be positive, got %d", d.Line)
	}
	if d.Column < 0 {
		return fmt.Errorf("snapshot diagnostic column must not be negative, got %d", d.Column)
	}
	return nil
}

// normalizeSnapshotDiagnostic converts a runner diagnostic while rejecting
// paths that could leak a temporary or absolute workspace location.
func normalizeSnapshotDiagnostic(d Diagnostic) (SnapshotDiagnostic, error) {
	file := filepath.ToSlash(filepath.Clean(d.File))
	normalized := SnapshotDiagnostic{
		Project:  d.Project,
		File:     file,
		Surface:  d.Surface,
		Code:     d.Code,
		Severity: d.Severity,
		Line:     d.Line,
		Column:   d.Column,
	}
	if err := validateSnapshotDiagnostic(normalized); err != nil {
		return SnapshotDiagnostic{}, err
	}
	return normalized, nil
}

// SnapshotSetFromReport normalizes every successful surface in a report.
// Callers must check report failures before publishing an update.
func SnapshotSetFromReport(report Report) (SnapshotSet, error) {
	set := make(SnapshotSet, len(report.Surfaces))
	for _, surface := range report.Surfaces {
		id := SnapshotID{Project: surface.Project, Surface: surface.Surface}
		if err := validateSnapshotID(id); err != nil {
			return nil, err
		}
		if _, exists := set[id]; exists {
			return nil, fmt.Errorf("duplicate snapshot surface %s/%s", surface.Project, surface.Surface)
		}
		rows := make([]SnapshotDiagnostic, 0, len(surface.Diagnostics))
		for _, diagnostic := range surface.Diagnostics {
			row, err := normalizeSnapshotDiagnostic(diagnostic)
			if err != nil {
				return nil, fmt.Errorf("normalize %s/%s diagnostic: %w", surface.Project, surface.Surface, err)
			}
			rows = append(rows, row)
		}
		sort.SliceStable(rows, func(i, j int) bool { return snapshotDiagnosticLess(rows[i], rows[j]) })
		set[id] = rows
	}
	return set, nil
}

func (set SnapshotSet) IDs() []SnapshotID {
	ids := make([]SnapshotID, 0, len(set))
	for id := range set {
		ids = append(ids, id)
	}
	sort.Slice(ids, func(i, j int) bool { return snapshotIDLess(ids[i], ids[j]) })
	return ids
}

func snapshotPath(root string, id SnapshotID) (string, error) {
	if err := validateSnapshotID(id); err != nil {
		return "", err
	}
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return "", fmt.Errorf("resolve snapshot root: %w", err)
	}
	rel := filepath.Join(filepath.FromSlash(string(id.Surface)), filepath.FromSlash(id.Project)+".jsonl")
	destination := filepath.Join(rootAbs, rel)
	relCheck, err := filepath.Rel(rootAbs, destination)
	if err != nil || relCheck == ".." || strings.HasPrefix(relCheck, ".."+string(os.PathSeparator)) {
		return "", fmt.Errorf("snapshot path escapes root for %q", id.Project)
	}
	return destination, nil
}

func encodeSnapshotRow(row SnapshotDiagnostic) ([]byte, error) {
	if err := validateSnapshotDiagnostic(row); err != nil {
		return nil, err
	}
	body, err := json.Marshal(row)
	if err != nil {
		return nil, err
	}
	return append(body, '\n'), nil
}

func decodeSnapshotRow(line []byte) (SnapshotDiagnostic, error) {
	if len(line) == 0 {
		return SnapshotDiagnostic{}, errors.New("blank snapshot line")
	}
	dec := json.NewDecoder(bytes.NewReader(line))
	dec.DisallowUnknownFields()
	var row SnapshotDiagnostic
	if err := dec.Decode(&row); err != nil {
		return SnapshotDiagnostic{}, fmt.Errorf("decode snapshot row: %w", err)
	}
	var extra any
	if err := dec.Decode(&extra); err != io.EOF {
		if err == nil {
			return SnapshotDiagnostic{}, errors.New("snapshot row contains trailing JSON")
		}
		return SnapshotDiagnostic{}, fmt.Errorf("snapshot row contains trailing data: %w", err)
	}
	canonical, err := encodeSnapshotRow(row)
	if err != nil {
		return SnapshotDiagnostic{}, err
	}
	if !bytes.Equal(line, bytes.TrimSuffix(canonical, []byte{'\n'})) {
		return SnapshotDiagnostic{}, errors.New("snapshot row is not in canonical JSON form")
	}
	return row, nil
}

func loadSnapshotFile(path string, id SnapshotID) ([]SnapshotDiagnostic, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = file.Close() }()
	rows := make([]SnapshotDiagnostic, 0)
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		row, decodeErr := decodeSnapshotRow(scanner.Bytes())
		if decodeErr != nil {
			return nil, fmt.Errorf("%s: %w", path, decodeErr)
		}
		if row.Project != id.Project || row.Surface != id.Surface {
			return nil, fmt.Errorf("%s: row identity does not match snapshot file", path)
		}
		rows = append(rows, row)
	}
	if err := scanner.Err(); err != nil {
		return nil, fmt.Errorf("read snapshot %s: %w", path, err)
	}
	if !sort.SliceIsSorted(rows, func(i, j int) bool { return snapshotDiagnosticLess(rows[i], rows[j]) }) {
		return nil, fmt.Errorf("snapshot %s is not deterministically sorted", path)
	}
	return rows, nil
}

// LoadSnapshotSet loads exactly the expected project/surface files and rejects
// missing or stale files below root.
func LoadSnapshotSet(root string, expected []SnapshotID) (SnapshotSet, error) {
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return nil, fmt.Errorf("resolve snapshot root: %w", err)
	}
	if info, statErr := os.Stat(rootAbs); statErr != nil {
		return nil, fmt.Errorf("stat snapshot root: %w", statErr)
	} else if !info.IsDir() {
		return nil, fmt.Errorf("snapshot root %q is not a directory", rootAbs)
	}
	set := make(SnapshotSet, len(expected))
	expectedPaths := make(map[string]SnapshotID, len(expected))
	for _, id := range expected {
		if _, exists := set[id]; exists {
			return nil, fmt.Errorf("duplicate expected snapshot %s/%s", id.Project, id.Surface)
		}
		path, pathErr := snapshotPath(rootAbs, id)
		if pathErr != nil {
			return nil, pathErr
		}
		rows, loadErr := loadSnapshotFile(path, id)
		if loadErr != nil {
			return nil, fmt.Errorf("load snapshot %s/%s: %w", id.Project, id.Surface, loadErr)
		}
		set[id] = rows
		expectedPaths[filepath.Clean(path)] = id
	}
	err = filepath.WalkDir(rootAbs, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			return nil
		}
		if _, ok := expectedPaths[filepath.Clean(path)]; !ok {
			return fmt.Errorf("unexpected file below snapshot root: %s", path)
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return set, nil
}

func writeSnapshotTree(root string, set SnapshotSet) error {
	ids := set.IDs()
	for _, id := range ids {
		rows := append([]SnapshotDiagnostic(nil), set[id]...)
		sort.SliceStable(rows, func(i, j int) bool { return snapshotDiagnosticLess(rows[i], rows[j]) })
		path, err := snapshotPath(root, id)
		if err != nil {
			return err
		}
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			return fmt.Errorf("create snapshot directory: %w", err)
		}
		file, err := os.Create(path)
		if err != nil {
			return fmt.Errorf("create snapshot %s: %w", path, err)
		}
		writer := bufio.NewWriter(file)
		for _, row := range rows {
			encoded, encodeErr := encodeSnapshotRow(row)
			if encodeErr != nil {
				_ = file.Close()
				return fmt.Errorf("encode snapshot %s: %w", path, encodeErr)
			}
			if _, writeErr := writer.Write(encoded); writeErr != nil {
				_ = file.Close()
				return fmt.Errorf("write snapshot %s: %w", path, writeErr)
			}
		}
		if err := writer.Flush(); err != nil {
			_ = file.Close()
			return fmt.Errorf("flush snapshot %s: %w", path, err)
		}
		if err := file.Close(); err != nil {
			return fmt.Errorf("close snapshot %s: %w", path, err)
		}
	}
	return nil
}

// WriteSnapshotSet atomically publishes a complete snapshot tree. The caller
// uses this only for explicit developer updates; normal verification only
// calls LoadSnapshotSet.
func WriteSnapshotSet(root string, set SnapshotSet) error {
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return fmt.Errorf("resolve snapshot root: %w", err)
	}
	parent := filepath.Dir(rootAbs)
	if err := os.MkdirAll(parent, 0o755); err != nil {
		return fmt.Errorf("create snapshot parent: %w", err)
	}
	stage, err := os.MkdirTemp(parent, ".static-analysis-snapshots-")
	if err != nil {
		return fmt.Errorf("create snapshot staging root: %w", err)
	}
	defer func() { _ = os.RemoveAll(stage) }()
	if err := writeSnapshotTree(stage, set); err != nil {
		return err
	}
	backup, err := os.MkdirTemp(parent, ".static-analysis-snapshots-backup-")
	if err != nil {
		return fmt.Errorf("create snapshot backup path: %w", err)
	}
	if err := os.Remove(backup); err != nil {
		_ = os.RemoveAll(backup)
		return fmt.Errorf("prepare snapshot backup path: %w", err)
	}
	backupExists := false
	preserveBackup := false
	defer func() {
		if backupExists && !preserveBackup {
			_ = os.RemoveAll(backup)
		}
	}()
	hadRoot := false
	if _, statErr := os.Stat(rootAbs); statErr == nil {
		if err := os.Rename(rootAbs, backup); err != nil {
			return fmt.Errorf("stage existing snapshot root: %w", err)
		}
		hadRoot = true
		backupExists = true
	} else if !os.IsNotExist(statErr) {
		return fmt.Errorf("inspect existing snapshot root: %w", statErr)
	}
	if err := os.Rename(stage, rootAbs); err != nil {
		if hadRoot {
			if restoreErr := os.Rename(backup, rootAbs); restoreErr == nil {
				backupExists = false
			} else {
				preserveBackup = true
				return fmt.Errorf("publish snapshot root: %w (restore existing root: %v)", err, restoreErr)
			}
		}
		return fmt.Errorf("publish snapshot root: %w", err)
	}
	if hadRoot {
		if err := os.RemoveAll(backup); err == nil {
			backupExists = false
		}
	}
	return nil
}

// CompareSnapshotSets compares rows as multisets and reports additions and
// removals in deterministic order. Missing IDs are represented as removals and
// unexpected IDs as additions.
func CompareSnapshotSets(want, got SnapshotSet) SnapshotDiff {
	diff := SnapshotDiff{}
	ids := make(map[SnapshotID]struct{}, len(want)+len(got))
	for id := range want {
		ids[id] = struct{}{}
	}
	for id := range got {
		ids[id] = struct{}{}
	}
	ordered := make([]SnapshotID, 0, len(ids))
	for id := range ids {
		ordered = append(ordered, id)
	}
	sort.Slice(ordered, func(i, j int) bool { return snapshotIDLess(ordered[i], ordered[j]) })
	for _, id := range ordered {
		wantCounts := countSnapshotRows(want[id])
		gotCounts := countSnapshotRows(got[id])
		rowKeys := make([]SnapshotDiagnostic, 0, len(wantCounts)+len(gotCounts))
		seen := make(map[SnapshotDiagnostic]struct{}, len(wantCounts)+len(gotCounts))
		for row := range wantCounts {
			seen[row] = struct{}{}
			rowKeys = append(rowKeys, row)
		}
		for row := range gotCounts {
			if _, exists := seen[row]; !exists {
				rowKeys = append(rowKeys, row)
			}
		}
		sort.Slice(rowKeys, func(i, j int) bool { return snapshotDiagnosticLess(rowKeys[i], rowKeys[j]) })
		for _, row := range rowKeys {
			if delta := gotCounts[row] - wantCounts[row]; delta > 0 {
				for i := 0; i < delta; i++ {
					diff.Added = append(diff.Added, row)
				}
			} else if delta < 0 {
				for i := 0; i < -delta; i++ {
					diff.Removed = append(diff.Removed, row)
				}
			}
		}
	}
	return diff
}

func countSnapshotRows(rows []SnapshotDiagnostic) map[SnapshotDiagnostic]int {
	counts := make(map[SnapshotDiagnostic]int, len(rows))
	for _, row := range rows {
		counts[row]++
	}
	return counts
}
