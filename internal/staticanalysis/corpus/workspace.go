package corpus

import (
	"errors"
	"fmt"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

var ErrUnsupportedProjectLayout = errors.New("unsupported third-party project layout")

// MaterializeOptions controls the temporary parent used by a third-party
// workspace. TempRoot is useful for tests; the caller owns that directory and
// the materializer only removes the unique child it creates.
type MaterializeOptions struct {
	TempRoot string
}

// MaterializedWorkspace is an isolated xlflow-compatible view of one
// vendored project. Mappings use slash-separated paths relative to Root.
type MaterializedWorkspace struct {
	Root      string
	ProjectID string
	Profile   string
	Mappings  map[string]string
	closed    bool
}

// Close removes only this workspace. It is safe to call more than once.
func (w *MaterializedWorkspace) Close() error {
	if w == nil || w.Root == "" || w.closed {
		return nil
	}
	if err := os.RemoveAll(w.Root); err != nil {
		return fmt.Errorf("remove materialized workspace %q: %w", w.Root, err)
	}
	w.closed = true
	return nil
}

// MaterializeThirdPartyProject copies one checked-in project into a temporary
// xlflow workspace and generates a config that preserves the project's host
// profile policy.
func MaterializeThirdPartyProject(corpusRoot string, project Project, opts MaterializeOptions) (workspace MaterializedWorkspace, err error) {
	if err := validateProjectForMaterialization(project); err != nil {
		return workspace, err
	}
	root, err := filepath.Abs(corpusRoot)
	if err != nil {
		return workspace, fmt.Errorf("resolve corpus root: %w", err)
	}
	sourceRoot, err := containedPath(root, project.Path)
	if err != nil {
		return workspace, fmt.Errorf("project %q source: %w", project.ID, err)
	}
	if info, statErr := os.Lstat(sourceRoot); statErr != nil {
		return workspace, fmt.Errorf("project %q source: %w", project.ID, statErr)
	} else if !info.IsDir() || info.Mode()&os.ModeSymlink != 0 {
		return workspace, fmt.Errorf("%w: project %q source is not a regular directory", ErrUnsupportedProjectLayout, project.ID)
	}
	resolvedRoot, err := filepath.EvalSymlinks(root)
	if err != nil {
		return workspace, fmt.Errorf("resolve corpus root: %w", err)
	}
	resolvedSource, err := filepath.EvalSymlinks(sourceRoot)
	if err != nil {
		return workspace, fmt.Errorf("resolve project %q source: %w", project.ID, err)
	}
	if err := ensureContained(resolvedRoot, resolvedSource); err != nil {
		return workspace, fmt.Errorf("project %q source: %w", project.ID, err)
	}

	tempParent := opts.TempRoot
	if tempParent == "" {
		tempParent = os.TempDir()
	}
	tempParent, err = filepath.Abs(tempParent)
	if err != nil {
		return workspace, fmt.Errorf("resolve workspace temp root: %w", err)
	}
	workspace.Root, err = os.MkdirTemp(tempParent, "xlflow-corpus-")
	if err != nil {
		return workspace, fmt.Errorf("create materialized workspace: %w", err)
	}
	workspace.ProjectID = "third_party/" + project.ID
	workspace.Profile = project.Profile
	workspace.Mappings = make(map[string]string)
	defer func() {
		if err != nil {
			if closeErr := workspace.Close(); closeErr != nil {
				err = errors.Join(err, closeErr)
			}
		}
	}()

	for _, dir := range []string{"src/modules", "src/classes", "src/forms", "src/workbook"} {
		if err := os.MkdirAll(filepath.Join(workspace.Root, filepath.FromSlash(dir)), 0o755); err != nil {
			return workspace, fmt.Errorf("create workspace source root %s: %w", dir, err)
		}
	}
	overrides := make(map[string]string, len(project.Classifications))
	for _, classification := range project.Classifications {
		overrides[strings.ToLower(classification.Path)] = classification.Kind
	}
	seenOverrides := make(map[string]bool, len(overrides))
	seenDestinations := make(map[string]string)
	err = filepath.WalkDir(sourceRoot, func(sourcePath string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.Type()&os.ModeSymlink != 0 {
			return fmt.Errorf("%w: source contains symlink or reparse point: %s", ErrUnsupportedProjectLayout, filepath.ToSlash(sourcePath))
		}
		info, infoErr := entry.Info()
		if infoErr != nil {
			return infoErr
		}
		if entry.IsDir() {
			return nil
		}
		if !info.Mode().IsRegular() {
			return fmt.Errorf("%w: source contains special file: %s", ErrUnsupportedProjectLayout, filepath.ToSlash(sourcePath))
		}
		rel, relErr := filepath.Rel(sourceRoot, sourcePath)
		if relErr != nil {
			return relErr
		}
		relSlash := filepath.ToSlash(rel)
		ext := strings.ToLower(filepath.Ext(rel))
		if ext != ".bas" && ext != ".cls" && ext != ".frm" {
			return nil
		}
		kind := defaultModuleKind(ext)
		if override, ok := overrides[strings.ToLower(relSlash)]; ok {
			kind = override
			seenOverrides[strings.ToLower(relSlash)] = true
		}
		destDir := moduleKindDirectory(kind)
		destRel := path.Join(destDir, relSlash)
		key := strings.ToLower(destRel)
		if prior, exists := seenDestinations[key]; exists {
			return fmt.Errorf("%w: source files %q and %q collide in workspace destination %q", ErrUnsupportedProjectLayout, prior, relSlash, destRel)
		}
		seenDestinations[key] = relSlash
		destination := filepath.Join(workspace.Root, filepath.FromSlash(destRel))
		if err := ensureContained(workspace.Root, destination); err != nil {
			return err
		}
		if err := os.MkdirAll(filepath.Dir(destination), 0o755); err != nil {
			return err
		}
		body, readErr := os.ReadFile(sourcePath)
		if readErr != nil {
			return readErr
		}
		if writeErr := os.WriteFile(destination, body, 0o644); writeErr != nil {
			return writeErr
		}
		workspaceRel, relErr := filepath.Rel(workspace.Root, destination)
		if relErr != nil {
			return relErr
		}
		workspace.Mappings[filepath.ToSlash(workspaceRel)] = relSlash
		return nil
	})
	if err != nil {
		return workspace, fmt.Errorf("materialize project %q: %w", project.ID, err)
	}
	for overridePath := range overrides {
		if !seenOverrides[overridePath] {
			return workspace, fmt.Errorf("%w: project %q classification targets missing source %q", ErrUnsupportedProjectLayout, project.ID, overridePath)
		}
	}

	cfg := config.Default()
	cfg.Project.Name = project.ID
	cfg.Project.Entry = "Corpus.Run"
	cfg.UserForm.CodeSource = "frm"
	applyProfilePolicy(&cfg, project.Profile)
	if err := config.Write(filepath.Join(workspace.Root, config.FileName), cfg); err != nil {
		return workspace, fmt.Errorf("write materialized project config: %w", err)
	}
	return workspace, nil
}

func validateProjectForMaterialization(project Project) error {
	if !identifierPattern.MatchString(project.ID) {
		return fmt.Errorf("invalid project id %q", project.ID)
	}
	if err := validateRelativePath(project.Path); err != nil {
		return fmt.Errorf("invalid project path: %w", err)
	}
	if !strings.HasPrefix(project.Path, "projects/third_party/") || project.Path == "projects/third_party/" {
		return fmt.Errorf("invalid project path %q: must be below projects/third_party", project.Path)
	}
	if err := validateClassifications(project.ID, project.Classifications); err != nil {
		return err
	}
	return nil
}

func defaultModuleKind(ext string) string {
	switch strings.ToLower(ext) {
	case ".bas":
		return ModuleKindStandard
	case ".cls":
		return ModuleKindClass
	case ".frm":
		return ModuleKindForm
	default:
		return ""
	}
}

func moduleKindDirectory(kind string) string {
	switch kind {
	case ModuleKindStandard:
		return "src/modules"
	case ModuleKindClass:
		return "src/classes"
	case ModuleKindForm:
		return "src/forms"
	case ModuleKindDocument:
		return "src/workbook"
	default:
		return ""
	}
}

func containedPath(root, relative string) (string, error) {
	if err := validateRelativePath(relative); err != nil {
		return "", err
	}
	candidate := filepath.Join(root, filepath.FromSlash(relative))
	if err := ensureContained(root, candidate); err != nil {
		return "", err
	}
	return candidate, nil
}

func ensureContained(root, candidate string) error {
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return err
	}
	candidateAbs, err := filepath.Abs(candidate)
	if err != nil {
		return err
	}
	rel, err := filepath.Rel(rootAbs, candidateAbs)
	if err != nil {
		return err
	}
	if rel == ".." || strings.HasPrefix(rel, ".."+string(os.PathSeparator)) {
		return errors.New("path escapes its root")
	}
	return nil
}
