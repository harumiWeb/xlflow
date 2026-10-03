package symbols

import (
	"cmp"
	"context"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

// DiscoverProjectSourceFilesContext extends the shared configured-root discovery with
// the repository's literal tests/ tree. The latter is intentionally not part
// of symbols.DiscoverSourceFiles because symbols inspection only reports
// configured production roots.
func DiscoverProjectSourceFilesContext(ctx context.Context, root string, projectCfg config.Config) ([]SourceFile, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	files, err := DiscoverSourceFilesContext(ctx, Options{RootDir: root, Config: projectCfg})
	if err != nil {
		return nil, err
	}
	seen := make(map[string]bool, len(files))
	for _, file := range files {
		abs, absErr := filepath.Abs(file.Path)
		if absErr == nil {
			seen[strings.ToLower(filepath.Clean(abs))] = true
		}
	}
	testsRoot := filepath.Join(root, "tests")
	if _, statErr := os.Stat(testsRoot); statErr != nil {
		if os.IsNotExist(statErr) {
			return files, nil
		}
		return nil, statErr
	}
	err = filepath.WalkDir(testsRoot, func(path string, entry os.DirEntry, walkErr error) error {
		if err := ctx.Err(); err != nil {
			return err
		}
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			return nil
		}
		ext := strings.ToLower(filepath.Ext(path))
		if ext != ".bas" && ext != ".cls" && ext != ".frm" {
			return nil
		}
		abs, absErr := filepath.Abs(path)
		if absErr != nil {
			return absErr
		}
		key := strings.ToLower(filepath.Clean(abs))
		if seen[key] {
			return nil
		}
		moduleKind, include := projectTestSourceKind(projectCfg, abs)
		if !include {
			return nil
		}
		seen[key] = true
		files = append(files, SourceFile{Path: abs, ModuleKind: moduleKind})
		return nil
	})
	if err != nil {
		return nil, err
	}
	slices.SortFunc(files, func(a, b SourceFile) int { return cmp.Compare(a.Path, b.Path) })
	return files, nil
}

// projectTestSourceKind applies the same UserForm source-of-truth rule to the
// literal tests/ tree that symbols.DiscoverSourceFiles applies to configured
// source roots. A tests form pair is conventionally laid out as
// tests/<forms-dir>/<Name>.frm and tests/<forms-dir>/code/<Name>.bas. Ordinary
// .bas/.cls files retain their existing standard/class classification.
func projectTestSourceKind(cfg config.Config, path string) (string, bool) {
	ext := strings.ToLower(filepath.Ext(path))
	switch ext {
	case ".cls":
		return "class", true
	case ".frm":
		sidecar := filepath.Join(filepath.Dir(path), "code", strings.TrimSuffix(filepath.Base(path), ext)+".bas")
		if strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
			if _, err := os.Stat(sidecar); err == nil {
				return "", false
			}
		}
		return "form", true
	case ".bas":
		if strings.EqualFold(filepath.Base(filepath.Dir(path)), "code") {
			formsDir := filepath.Dir(filepath.Dir(path))
			form := filepath.Join(formsDir, strings.TrimSuffix(filepath.Base(path), ext)+".frm")
			if _, err := os.Stat(form); err == nil {
				if strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
					return "form", true
				}
				return "", false
			}
		}
		return "standard", true
	default:
		return "", false
	}
}
