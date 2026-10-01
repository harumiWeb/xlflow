// Package filepull extracts tracked VBA source from a saved .xlsm workbook
// without starting Excel or using COM/VBIDE.
package filepull

import (
	"archive/zip"
	"context"
	"errors"
	"fmt"
	"io"
	"io/fs"
	"maps"
	"os"
	"path/filepath"
	"regexp"
	"slices"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
)

var (
	ErrMissingVBAProject   = errors.New("file pull: workbook has no xl/vbaProject.bin")
	ErrMalformedVBAProject = errors.New("file pull: malformed VBA project")
	ErrProtectedProject    = errors.New("file pull: protected VBA project")
	ErrUserFormUnsupported = errors.New("file pull: UserForm extraction unsupported")
	ErrUnsafeSourcePath    = errors.New("file pull: unsafe source path")
	ErrPublish             = errors.New("file pull: source publication failed")
)

type ModuleCounts struct {
	Standard int
	Class    int
	Document int
	Form     int
}

type Result struct {
	WorkbookPath string
	ModulesDir   string
	ClassesDir   string
	FormsDir     string
	WorkbookDir  string
	Written      []string
	Removed      []string
	Modules      ModuleCounts
	CodePage     uint16
}

type plannedFile struct {
	path string
	body []byte
}

type plan struct {
	result Result
	files  []plannedFile
	stale  []string
}

type PullOptions struct {
	Coordination *coordination.Manager
	Wait         bool
	WaitTimeout  time.Duration
}

var folderAnnotationPattern = regexp.MustCompile(`(?i)^'?@Folder\(\s*"([^"]*)"\s*\)`)

// Pull parses, validates, and publishes one complete saved-workbook snapshot.
func Pull(root string, cfg config.Config, workbookPath string) (Result, error) {
	return PullContext(context.Background(), root, cfg, workbookPath, PullOptions{})
}

// PullContext parses, validates, and publishes one complete saved-workbook
// snapshot while holding leases for every managed source root.
func PullContext(ctx context.Context, root string, cfg config.Config, workbookPath string, opts PullOptions) (Result, error) {
	release, err := acquireSourceTrees(ctx, root, cfg, opts)
	if err != nil {
		return Result{}, err
	}
	defer release()

	p, err := buildPlan(root, cfg, workbookPath)
	if err != nil {
		return Result{}, err
	}
	if err := publish(p); err != nil {
		return Result{}, err
	}
	return p.result, nil
}

func acquireSourceTrees(ctx context.Context, root string, cfg config.Config, opts PullOptions) (func(), error) {
	if ctx == nil {
		ctx = context.Background()
	}
	manager := opts.Coordination
	if manager == nil {
		var err error
		manager, err = coordination.NewDefaultManager()
		if err != nil {
			return nil, fmt.Errorf("initialize source-tree coordination: %w", err)
		}
	}
	type lockTarget struct {
		identity coordination.ResourceIdentity
		shared   bool
	}
	roots := resolvedRoots(root, cfg)
	targets := map[string]lockTarget{}
	for _, path := range []string{roots.modules, roots.classes, roots.workbook} {
		identity, err := coordination.NewSourceTreeIdentity(root, path)
		if err != nil {
			return nil, fmt.Errorf("resolve source-tree identity %s: %w", path, err)
		}
		for index, hierarchyIdentity := range coordination.SourceTreeLockHierarchy(identity) {
			target := lockTarget{identity: hierarchyIdentity, shared: index != 0}
			if current, exists := targets[hierarchyIdentity.LockID]; !exists || current.shared {
				targets[hierarchyIdentity.LockID] = target
			}
		}
	}
	lockIDs := slices.Sorted(maps.Keys(targets))
	leases := make([]*coordination.Lease, 0, len(lockIDs))
	release := func() {
		for i := len(leases) - 1; i >= 0; i-- {
			_ = leases[i].Release()
		}
	}
	acquireCtx := ctx
	waitStarted := false
	for _, lockID := range lockIDs {
		target := targets[lockID]
		request := coordination.AcquireRequest{
			Identity:      target.identity,
			Command:       "pull",
			OperationKind: coordination.OperationMutate,
			ResourceScope: coordination.ResourceSourceTree,
			Shared:        target.shared,
		}
		lease, err := manager.Acquire(acquireCtx, request)
		if opts.Wait && errors.Is(err, coordination.ErrSourceTreeBusy) {
			if !waitStarted && opts.WaitTimeout > 0 {
				var cancelWait context.CancelFunc
				acquireCtx, cancelWait = context.WithTimeout(ctx, opts.WaitTimeout)
				defer cancelWait()
				waitStarted = true
			}
			request.Wait = true
			lease, err = manager.Acquire(acquireCtx, request)
		}
		if err == nil && waitStarted && acquireCtx.Err() != nil {
			_ = lease.Release()
			err = acquireCtx.Err()
		}
		if err != nil {
			release()
			return nil, err
		}
		leases = append(leases, lease)
	}
	return release, nil
}

func buildPlan(root string, cfg config.Config, workbookPath string) (plan, error) {
	projectBytes, err := readVBAProject(workbookPath)
	if err != nil {
		return plan{}, err
	}
	project, err := vbaproject.Read(projectBytes)
	if err != nil {
		return plan{}, fmt.Errorf("%w: %v", ErrMalformedVBAProject, err)
	}
	if project.Protection.IsProtected {
		return plan{}, ErrProtectedProject
	}
	for _, module := range project.Modules {
		if module.Type == vbaproject.ModuleForm {
			return plan{}, fmt.Errorf("%w: %s", ErrUserFormUnsupported, module.Name)
		}
	}

	roots := resolvedRoots(root, cfg)
	if err := validateFormsRootSeparation(root, roots); err != nil {
		return plan{}, err
	}
	result := Result{
		WorkbookPath: workbookPath,
		ModulesDir:   roots.modules,
		ClassesDir:   roots.classes,
		FormsDir:     roots.forms,
		WorkbookDir:  roots.workbook,
		CodePage:     project.Props.CodePage,
	}
	p := plan{result: result}
	seenNames := map[string]string{}
	seenPaths := map[string]string{}
	desired := map[string]string{}
	for _, module := range project.Modules {
		if !sourceinventory.ValidComponentName(module.Name) {
			return plan{}, fmt.Errorf("%w: invalid component name %q", ErrMalformedVBAProject, module.Name)
		}
		nameKey := strings.ToLower(module.Name)
		if prior, ok := seenNames[nameKey]; ok {
			return plan{}, fmt.Errorf("%w: duplicate components %q and %q", ErrMalformedVBAProject, prior, module.Name)
		}
		seenNames[nameKey] = module.Name

		disk, err := vbaproject.ExportModuleSource(module)
		if err != nil {
			return plan{}, fmt.Errorf("%w: %v", ErrMalformedVBAProject, err)
		}
		if cfg.VBA.LineNumbers.Enabled {
			disk, err = removeGeneratedLineNumbers(disk)
			if err != nil {
				return plan{}, err
			}
		}
		if !utf8.ValidString(disk) {
			return plan{}, fmt.Errorf("%w: module %q is not valid UTF-8", ErrMalformedVBAProject, module.Name)
		}

		targetRoot, extension, err := moduleDestination(roots, module.Type)
		if err != nil {
			return plan{}, fmt.Errorf("%w: module %q: %v", ErrMalformedVBAProject, module.Name, err)
		}
		relative := module.Name + extension
		if cfg.VBA.Folders && cfg.VBA.FolderAnnotation == "update" {
			segments, err := folderSegments(disk)
			if err != nil {
				return plan{}, fmt.Errorf("%w: module %q: %v", ErrUnsafeSourcePath, module.Name, err)
			}
			parts := append(segments, relative)
			relative = filepath.Join(parts...)
		}
		target, err := safeTarget(targetRoot, relative)
		if err != nil {
			return plan{}, err
		}
		pathKey := strings.ToLower(filepath.Clean(target))
		if prior, ok := seenPaths[pathKey]; ok {
			return plan{}, fmt.Errorf("%w: %s and %s resolve to the same source path", ErrUnsafeSourcePath, prior, module.Name)
		}
		seenPaths[pathKey] = module.Name
		desired[pathKey] = target
		p.files = append(p.files, plannedFile{path: target, body: []byte(disk)})
		switch module.Type {
		case vbaproject.ModuleStd:
			p.result.Modules.Standard++
		case vbaproject.ModuleClass:
			p.result.Modules.Class++
		case vbaproject.ModuleDocument:
			p.result.Modules.Document++
		}
	}

	for _, managedRoot := range []struct {
		path string
		exts map[string]bool
	}{
		{roots.modules, map[string]bool{".bas": true}},
		{roots.classes, map[string]bool{".cls": true}},
		{roots.workbook, map[string]bool{".bas": true, ".cls": true}},
	} {
		existing, err := managedFiles(managedRoot.path, managedRoot.exts)
		if err != nil {
			return plan{}, err
		}
		for _, path := range existing {
			desiredPath, ok := desired[strings.ToLower(filepath.Clean(path))]
			if !ok {
				p.stale = append(p.stale, path)
				continue
			}
			matches, err := matchesManagedTarget(managedRoot.path, path, desiredPath)
			if err != nil {
				return plan{}, err
			}
			if !matches {
				p.stale = append(p.stale, path)
			}
		}
	}
	slices.SortFunc(p.files, func(a, b plannedFile) int { return strings.Compare(a.path, b.path) })
	slices.Sort(p.stale)
	for _, file := range p.files {
		p.result.Written = append(p.result.Written, file.path)
	}
	p.result.Removed = append(p.result.Removed, p.stale...)
	return p, nil
}

func readVBAProject(workbookPath string) ([]byte, error) {
	reader, err := zip.OpenReader(workbookPath)
	if err != nil {
		var pathErr *os.PathError
		if errors.As(err, &pathErr) {
			return nil, err
		}
		return nil, fmt.Errorf("%w: open workbook package: %v", ErrMalformedVBAProject, err)
	}
	defer func() { _ = reader.Close() }()
	var matches []*zip.File
	for _, entry := range reader.File {
		if filepath.ToSlash(entry.Name) == "xl/vbaProject.bin" {
			matches = append(matches, entry)
		}
	}
	if len(matches) == 0 {
		return nil, ErrMissingVBAProject
	}
	if len(matches) != 1 {
		return nil, fmt.Errorf("%w: duplicate xl/vbaProject.bin entries", ErrMalformedVBAProject)
	}
	rc, err := matches[0].Open()
	if err != nil {
		return nil, fmt.Errorf("%w: open xl/vbaProject.bin: %v", ErrMalformedVBAProject, err)
	}
	body, readErr := io.ReadAll(rc)
	closeErr := rc.Close()
	if readErr != nil {
		return nil, fmt.Errorf("%w: read xl/vbaProject.bin: %v", ErrMalformedVBAProject, readErr)
	}
	if closeErr != nil {
		return nil, fmt.Errorf("%w: close xl/vbaProject.bin: %v", ErrMalformedVBAProject, closeErr)
	}
	return body, nil
}

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
		modules: resolve(cfg.Src.Modules), classes: resolve(cfg.Src.Classes),
		forms: resolve(cfg.Src.Forms), workbook: resolve(cfg.Src.Workbook),
	}
}

func validateFormsRootSeparation(projectRoot string, roots sourceRoots) error {
	forms, err := coordination.NewWorkbookIdentity(projectRoot, roots.forms)
	if err != nil {
		return fmt.Errorf("%w: resolve forms root %s: %v", ErrUnsafeSourcePath, roots.forms, err)
	}
	for _, managed := range []string{roots.modules, roots.classes, roots.workbook} {
		identity, err := coordination.NewWorkbookIdentity(projectRoot, managed)
		if err != nil {
			return fmt.Errorf("%w: resolve managed root %s: %v", ErrUnsafeSourcePath, managed, err)
		}
		if pathsOverlap(forms.CanonicalPath, identity.CanonicalPath) {
			return fmt.Errorf("%w: forms root %s overlaps managed root %s", ErrUnsafeSourcePath, roots.forms, managed)
		}
	}
	return nil
}

func pathsOverlap(left, right string) bool {
	return pathWithin(left, right) || pathWithin(right, left)
}

func pathWithin(root, target string) bool {
	rel, err := filepath.Rel(root, target)
	return err == nil && rel != ".." && !strings.HasPrefix(rel, ".."+string(filepath.Separator)) && !filepath.IsAbs(rel)
}

func matchesManagedTarget(root, existing, desired string) (bool, error) {
	if filepath.Clean(existing) == filepath.Clean(desired) {
		return true, nil
	}
	exact, err := hasExactRelativeSpelling(root, desired)
	if err != nil || exact {
		return false, err
	}
	existingInfo, err := os.Stat(existing)
	if err != nil {
		return false, err
	}
	desiredInfo, err := os.Stat(desired)
	if errors.Is(err, os.ErrNotExist) {
		return false, nil
	}
	if err != nil {
		return false, err
	}
	return os.SameFile(existingInfo, desiredInfo), nil
}

func hasExactRelativeSpelling(root, target string) (bool, error) {
	rel, err := filepath.Rel(root, target)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) || filepath.IsAbs(rel) {
		return false, fmt.Errorf("%w: target %s is outside managed root %s", ErrUnsafeSourcePath, target, root)
	}
	current := root
	for _, segment := range strings.Split(rel, string(filepath.Separator)) {
		entries, err := os.ReadDir(current)
		if errors.Is(err, os.ErrNotExist) {
			return false, nil
		}
		if err != nil {
			return false, err
		}
		found := false
		for _, entry := range entries {
			if entry.Name() == segment {
				found = true
				current = filepath.Join(current, entry.Name())
				break
			}
		}
		if !found {
			return false, nil
		}
	}
	return true, nil
}

func moduleDestination(roots sourceRoots, typ vbaproject.ModuleType) (string, string, error) {
	switch typ {
	case vbaproject.ModuleStd:
		return roots.modules, ".bas", nil
	case vbaproject.ModuleClass:
		return roots.classes, ".cls", nil
	case vbaproject.ModuleDocument:
		return roots.workbook, ".bas", nil
	default:
		return "", "", fmt.Errorf("unsupported module type %d", typ)
	}
}

func folderSegments(source string) ([]string, error) {
	for line := range strings.Lines(source) {
		match := folderAnnotationPattern.FindStringSubmatch(strings.TrimSpace(line))
		if len(match) != 2 {
			continue
		}
		var segments []string
		for segment := range strings.SplitSeq(match[1], ".") {
			cleaned, err := cleanFolderSegment(segment)
			if err != nil {
				return nil, err
			}
			if cleaned != "" {
				segments = append(segments, cleaned)
			}
		}
		return segments, nil
	}
	return nil, nil
}

func cleanFolderSegment(segment string) (string, error) {
	segment = strings.TrimSpace(segment)
	if segment == "" {
		return "", nil
	}
	if segment == "." || segment == ".." || strings.ContainsAny(segment, `/\\`) {
		return "", fmt.Errorf("unsafe folder segment %q", segment)
	}
	var b strings.Builder
	for _, r := range segment {
		if r < 32 || strings.ContainsRune(`<>:"|?*`, r) {
			b.WriteByte('_')
			continue
		}
		b.WriteRune(r)
	}
	cleaned := strings.TrimSpace(b.String())
	if cleaned == "" || cleaned == "." || cleaned == ".." {
		return "", fmt.Errorf("unsafe folder segment %q", segment)
	}
	return cleaned, nil
}

func safeTarget(root, relative string) (string, error) {
	target := filepath.Join(root, relative)
	rel, err := filepath.Rel(root, target)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) || filepath.IsAbs(rel) {
		return "", fmt.Errorf("%w: %s", ErrUnsafeSourcePath, target)
	}
	rootIdentity, err := coordination.NewWorkbookIdentity(root, root)
	if err != nil {
		return "", fmt.Errorf("%w: resolve source root %s: %v", ErrUnsafeSourcePath, root, err)
	}
	targetIdentity, err := coordination.NewWorkbookIdentity(root, target)
	if err != nil {
		return "", fmt.Errorf("%w: resolve target %s: %v", ErrUnsafeSourcePath, target, err)
	}
	canonicalRel, err := filepath.Rel(rootIdentity.CanonicalPath, targetIdentity.CanonicalPath)
	if err != nil || canonicalRel == ".." || strings.HasPrefix(canonicalRel, ".."+string(filepath.Separator)) || filepath.IsAbs(canonicalRel) {
		return "", fmt.Errorf("%w: target %s escapes source root", ErrUnsafeSourcePath, target)
	}
	return target, nil
}

func managedFiles(root string, extensions map[string]bool) ([]string, error) {
	var files []string
	err := filepath.WalkDir(root, func(path string, entry fs.DirEntry, walkErr error) error {
		if errors.Is(walkErr, os.ErrNotExist) {
			return nil
		}
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			return nil
		}
		if extensions[strings.ToLower(filepath.Ext(entry.Name()))] {
			files = append(files, path)
		}
		return nil
	})
	if errors.Is(err, os.ErrNotExist) {
		return nil, nil
	}
	return files, err
}
