package filepush

import (
	"fmt"
	"maps"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

// discoveredFile mirrors the .NET VbaSourceHelper.DiscoveredSourceFile shape.
// It is produced by the raw filesystem enumeration used for source
// fingerprinting, duplicate detection, and line-number preflight; it is
// intentionally separate from sourceinventory.Discover so the fingerprint
// covers exactly the same file set the Excel bridge records.
type discoveredFile struct {
	Kind         string // "module" | "class" | "form" | "document" | "form_code"
	RootDir      string
	FullPath     string
	RelativePath string // relative to RootDir, '/' separators
	Extension    string // lowercase
	ModuleName   string
}

// sourceRoots resolves the configured source directories against the project
// root. Relative config paths are anchored under root; absolute paths are used
// as-is.
type sourceRoots struct {
	modules  string
	classes  string
	forms    string
	workbook string
}

func resolvedRoots(root string, cfg config.Config) sourceRoots {
	resolve := func(path string) string {
		if filepath.IsAbs(path) {
			return filepath.Clean(path)
		}
		return filepath.Join(root, filepath.FromSlash(strings.ReplaceAll(path, "\\", "/")))
	}
	return sourceRoots{
		modules:  resolve(cfg.Src.Modules),
		classes:  resolve(cfg.Src.Classes),
		forms:    resolve(cfg.Src.Forms),
		workbook: resolve(cfg.Src.Workbook),
	}
}

func isSidecarMode(codeSource string) bool {
	return strings.EqualFold(strings.TrimSpace(codeSource), "sidecar")
}

// discoverSourceFiles ports VbaSourceHelper.DiscoverSourceFiles: plain
// recursive extension matching under each configured root, with the reserved
// src/forms/code directory excluded from the "form" enumeration only in
// sidecar mode. Missing roots contribute no files, matching the .NET
// Directory.Exists guard. Within each root group files are sorted by relative
// path so the fingerprint is deterministic.
func discoverSourceFiles(roots sourceRoots, codeSource string) []discoveredFile {
	var files []discoveredFile
	addFiles := func(dir, kind, ext, excludedDir string) {
		files = append(files, filesFromDir(dir, kind, ext, excludedDir)...)
	}

	addFiles(roots.modules, "module", ".bas", "")
	addFiles(roots.classes, "class", ".cls", "")

	formsCodeDir := ""
	if isSidecarMode(codeSource) {
		formsCodeDir = filepath.Join(roots.forms, "code")
	}
	addFiles(roots.forms, "form", ".bas", formsCodeDir)
	addFiles(roots.forms, "form", ".cls", formsCodeDir)
	addFiles(roots.forms, "form", ".frm", formsCodeDir)
	addFiles(roots.forms, "form", ".frx", formsCodeDir)

	addFiles(roots.workbook, "document", ".bas", "")

	if isSidecarMode(codeSource) && strings.TrimSpace(roots.forms) != "" {
		codeDir := filepath.Join(roots.forms, "code")
		if info, err := os.Stat(codeDir); err == nil && info.IsDir() {
			files = append(files, filesFromDir(codeDir, "form_code", ".bas", "")...)
		}
	}
	return files
}

func filesFromDir(dir, kind, ext, excludedDir string) []discoveredFile {
	if strings.TrimSpace(dir) == "" {
		return nil
	}
	info, err := os.Stat(dir)
	if err != nil || !info.IsDir() {
		return nil
	}
	excludedFull := ""
	if strings.TrimSpace(excludedDir) != "" {
		if abs, absErr := filepath.Abs(excludedDir); absErr == nil {
			excludedFull = filepath.Clean(abs) + string(filepath.Separator)
		}
	}
	var found []discoveredFile
	_ = filepath.WalkDir(dir, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil || entry.IsDir() {
			return walkErr
		}
		if !strings.EqualFold(filepath.Ext(entry.Name()), ext) {
			return nil
		}
		abs, absErr := filepath.Abs(path)
		if absErr != nil {
			return nil
		}
		abs = filepath.Clean(abs)
		if excludedFull != "" && strings.HasPrefix(strings.ToLower(abs), strings.ToLower(excludedFull)) {
			return nil
		}
		rel, relErr := filepath.Rel(dir, abs)
		if relErr != nil {
			return nil
		}
		found = append(found, discoveredFile{
			Kind:         kind,
			RootDir:      dir,
			FullPath:     abs,
			RelativePath: filepath.ToSlash(rel),
			Extension:    ext,
			ModuleName:   strings.TrimSuffix(entry.Name(), filepath.Ext(entry.Name())),
		})
		return nil
	})
	slices.SortFunc(found, func(a, b discoveredFile) int {
		return strings.Compare(a.RelativePath, b.RelativePath)
	})
	return found
}

// findDuplicateModuleNames ports FindDuplicateModuleNames: a
// case-insensitive basename collision across module/class/form/document
// entries is fatal. .frx companions and form-code sidecars never collide
// because they are not standalone components.
func findDuplicateModuleNames(files []discoveredFile) [][]string {
	seen := map[string][]string{}
	for _, file := range files {
		if file.Extension == ".frx" || file.Kind == "form_code" {
			continue
		}
		key := strings.ToLower(file.ModuleName)
		seen[key] = append(seen[key], file.RelativePath)
	}
	var duplicates [][]string
	for _, key := range slices.Sorted(maps.Keys(seen)) {
		if len(seen[key]) > 1 {
			duplicates = append(duplicates, seen[key])
		}
	}
	return duplicates
}

// validateLineNumberSources ports TryValidateLineNumberSources: before any
// workbook mutation, every discovered text file must prove safe for Erl
// instrumentation so the push cannot fail halfway through.
func validateLineNumberSources(files []discoveredFile) error {
	for _, file := range files {
		if file.Extension == ".frx" {
			continue
		}
		body, err := os.ReadFile(file.FullPath)
		if err != nil {
			continue
		}
		if _, issue := tryAddLineNumbers(string(body)); issue != nil {
			return fmt.Errorf("%w: %s:%d: %s", ErrLineNumberSafety, file.RelativePath, issue.Line, issue.Message)
		}
	}
	return nil
}
