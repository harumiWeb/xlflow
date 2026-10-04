package filepush

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
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
	Body         []byte
	HasBody      bool
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

func canonicalPathKey(path string) string {
	return strings.ToLower(filepath.Clean(path))
}

func samePath(left, right string) bool {
	return strings.EqualFold(filepath.Clean(left), filepath.Clean(right))
}

// discoverSourceFiles ports VbaSourceHelper.DiscoverSourceFiles: plain
// recursive extension matching under each configured root, with the reserved
// src/forms/code directory excluded from the "form" enumeration only in
// sidecar mode. Missing roots contribute no files, matching the .NET
// Directory.Exists guard. Within each root group files are sorted by relative
// path so the fingerprint is deterministic.
func discoverSourceFiles(roots sourceRoots, codeSource string) ([]discoveredFile, error) {
	var files []discoveredFile
	addFiles := func(dir, kind, ext, excludedDir string) error {
		found, err := filesFromDir(dir, kind, ext, excludedDir)
		if err != nil {
			return err
		}
		files = append(files, found...)
		return nil
	}

	for _, root := range []struct{ dir, kind, ext, excluded string }{
		{roots.modules, "module", ".bas", ""},
		{roots.classes, "class", ".cls", ""},
	} {
		if err := addFiles(root.dir, root.kind, root.ext, root.excluded); err != nil {
			return nil, err
		}
	}

	formsCodeDir := ""
	if isSidecarMode(codeSource) {
		formsCodeDir = filepath.Join(roots.forms, "code")
	}
	for _, ext := range []string{".bas", ".cls", ".frm", ".frx"} {
		if err := addFiles(roots.forms, "form", ext, formsCodeDir); err != nil {
			return nil, err
		}
	}
	if err := addFiles(roots.workbook, "document", ".bas", ""); err != nil {
		return nil, err
	}

	if isSidecarMode(codeSource) && strings.TrimSpace(roots.forms) != "" {
		codeDir := filepath.Join(roots.forms, "code")
		if found, err := filesFromDir(codeDir, "form_code", ".bas", ""); err != nil {
			return nil, err
		} else {
			files = append(files, found...)
		}
	}
	return files, nil
}

func filesFromDir(dir, kind, ext, excludedDir string) ([]discoveredFile, error) {
	if strings.TrimSpace(dir) == "" {
		return nil, nil
	}
	info, err := os.Stat(dir)
	if errors.Is(err, os.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, &SourceReadError{Path: dir, Err: err}
	}
	if !info.IsDir() {
		return nil, fmt.Errorf("%w: source root %s is not a directory", packpkg.ErrAmbiguousLayout, dir)
	}
	excludedFull := ""
	if strings.TrimSpace(excludedDir) != "" {
		if abs, absErr := filepath.Abs(excludedDir); absErr == nil {
			excludedFull = filepath.Clean(abs) + string(filepath.Separator)
		}
	}
	var found []discoveredFile
	walkErr := filepath.WalkDir(dir, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return &SourceReadError{Path: path, Err: walkErr}
		}
		if entry.IsDir() {
			if kind == "form" && samePath(path, filepath.Join(dir, "assets")) {
				return filepath.SkipDir
			}
			return nil
		}
		if !strings.EqualFold(filepath.Ext(entry.Name()), ext) {
			return nil
		}
		abs, absErr := filepath.Abs(path)
		if absErr != nil {
			return &SourceReadError{Path: path, Err: absErr}
		}
		abs = filepath.Clean(abs)
		if excludedFull != "" && strings.HasPrefix(strings.ToLower(abs), strings.ToLower(excludedFull)) {
			return nil
		}
		rel, relErr := filepath.Rel(dir, abs)
		if relErr != nil {
			return &SourceReadError{Path: path, Err: relErr}
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
	if walkErr != nil {
		return nil, walkErr
	}
	slices.SortFunc(found, func(a, b discoveredFile) int {
		return strings.Compare(a.RelativePath, b.RelativePath)
	})
	return found, nil
}
