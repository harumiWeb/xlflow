// Package sourceprojectfs adapts a configured xlflow source tree to the
// filesystem-independent sourceproject model.
package sourceprojectfs

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

// Options controls loading a source project from a configured filesystem
// workspace. RootDir is the project root used to resolve configured source
// directories and the legacy tests directory. PathFilter, when present, is
// called with the source's logical path before the file is read. Logical paths
// preserve the absolute or relative form implied by RootDir.
type Options struct {
	RootDir    string
	Config     config.Config
	PathFilter func(string) bool
}

type candidate struct {
	physicalPath string
	logicalPath  string
	kind         sourceproject.ModuleKind
	isTest       bool
}

// Load loads a source project using a background context.
func Load(opts Options) (sourceproject.SourceProject, error) {
	return LoadContext(context.Background(), opts)
}

// LoadContext is the cancellable variant of Load. It discovers configured
// production roots and the legacy tests directory, reads selected files, and
// returns their exact source bytes in a sourceproject.SourceProject.
func LoadContext(ctx context.Context, opts Options) (project sourceproject.SourceProject, err error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return sourceproject.SourceProject{}, err
	}
	rootDir := opts.RootDir
	if rootDir == "" {
		rootDir = "."
	}
	absoluteRoot, err := absoluteCleanPath(rootDir)
	if err != nil {
		return sourceproject.SourceProject{}, err
	}

	finishDiscovery := analysisstats.Measure(ctx, "source_discovery")
	discoveryFinished := false
	defer func() {
		if !discoveryFinished {
			finishDiscovery(0, err)
		}
	}()
	production, discoveryErr := symbols.DiscoverSourceFilesContext(ctx, symbols.Options{
		RootDir: opts.RootDir,
		Config:  opts.Config,
	})
	if discoveryErr != nil {
		return sourceproject.SourceProject{}, discoveryErr
	}
	tests, discoveryErr := symbols.DiscoverSourceFilesContext(ctx, symbols.Options{
		RootDir: opts.RootDir,
		Config:  opts.Config,
		Path:    "tests",
	})
	if discoveryErr != nil {
		return sourceproject.SourceProject{}, discoveryErr
	}

	candidates := make([]candidate, 0, len(production)+len(tests))
	seen := make(map[string]struct{}, cap(candidates))
	appendCandidates := func(files []symbols.SourceFile, isTest bool) error {
		for _, file := range files {
			if err := ctx.Err(); err != nil {
				return err
			}
			physicalPath, err := absoluteCleanPath(file.Path)
			if err != nil {
				return err
			}
			logicalPath := logicalSourcePath(opts.RootDir, absoluteRoot, physicalPath)
			key := dedupeKey(physicalPath)
			if _, ok := seen[key]; ok {
				// Production candidates are appended first, so they retain
				// precedence if a configured root overlaps tests.
				continue
			}
			kind, err := convertModuleKind(file.ModuleKind)
			if err != nil {
				return err
			}
			seen[key] = struct{}{}
			candidates = append(candidates, candidate{
				physicalPath: physicalPath,
				logicalPath:  logicalPath,
				kind:         kind,
				isTest:       isTest && kind == sourceproject.ModuleKindStandard,
			})
		}
		return nil
	}
	if err := appendCandidates(production, false); err != nil {
		return sourceproject.SourceProject{}, err
	}
	if err := appendCandidates(tests, true); err != nil {
		return sourceproject.SourceProject{}, err
	}
	sort.Slice(candidates, func(i, j int) bool {
		if candidates[i].logicalPath == candidates[j].logicalPath {
			return !candidates[i].isTest && candidates[j].isTest
		}
		return candidates[i].logicalPath < candidates[j].logicalPath
	})

	selected := candidates[:0]
	for _, file := range candidates {
		if err := ctx.Err(); err != nil {
			return sourceproject.SourceProject{}, err
		}
		if opts.PathFilter != nil && !opts.PathFilter(file.logicalPath) {
			continue
		}
		selected = append(selected, file)
	}
	finishDiscovery(len(selected), nil)
	discoveryFinished = true

	files := make([]sourceproject.SourceFile, 0, len(selected))
	for _, file := range selected {
		if err := ctx.Err(); err != nil {
			return sourceproject.SourceProject{}, err
		}
		finishRead := analysisstats.Measure(ctx, "file_read")
		source, readErr := os.ReadFile(file.physicalPath)
		finishRead(len(source), readErr)
		if readErr != nil {
			return sourceproject.SourceProject{}, readErr
		}
		files = append(files, sourceproject.SourceFile{
			Path:       file.logicalPath,
			Source:     source,
			ModuleKind: file.kind,
			IsTest:     file.isTest,
		})
	}
	if err := ctx.Err(); err != nil {
		return sourceproject.SourceProject{}, err
	}
	return sourceproject.SourceProject{Files: files}, nil
}

func logicalSourcePath(rootDir, absoluteRoot, physicalPath string) string {
	relativePath, err := filepath.Rel(absoluteRoot, physicalPath)
	if err != nil {
		// A configured absolute source root may be on another Windows volume.
		// filepath.WalkDir historically returned an absolute path in this case.
		return physicalPath
	}
	return filepath.Clean(filepath.Join(rootDir, relativePath))
}

func absoluteCleanPath(path string) (string, error) {
	abs, err := filepath.Abs(path)
	if err != nil {
		return "", err
	}
	return filepath.Clean(abs), nil
}

func dedupeKey(path string) string {
	if runtime.GOOS == "windows" {
		return strings.ToLower(path)
	}
	return path
}

func convertModuleKind(kind string) (sourceproject.ModuleKind, error) {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case string(sourceproject.ModuleKindStandard):
		return sourceproject.ModuleKindStandard, nil
	case string(sourceproject.ModuleKindClass):
		return sourceproject.ModuleKindClass, nil
	case string(sourceproject.ModuleKindForm):
		return sourceproject.ModuleKindForm, nil
	case string(sourceproject.ModuleKindDocument):
		return sourceproject.ModuleKindDocument, nil
	default:
		return "", fmt.Errorf("unsupported VBA module kind %q", kind)
	}
}
