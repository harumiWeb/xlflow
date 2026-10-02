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
	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

var (
	ErrMissingVBAProject             = errors.New("file pull: workbook has no xl/vbaProject.bin")
	ErrMalformedVBAProject           = errors.New("file pull: malformed VBA project")
	ErrProtectedProject              = errors.New("file pull: protected VBA project")
	ErrUserFormCodeSourceUnsupported = errors.New("file pull: UserForm code source unsupported")
	ErrUserFormDesignerMalformed     = errors.New("file pull: malformed UserForm Designer storage")
	ErrUserFormDesignerUnsupported   = errors.New("file pull: unsupported UserForm Designer structure")
	ErrUserFormIdentityMismatch      = errors.New("file pull: inconsistent UserForm module and storage identity")
	ErrUnsafeSourcePath              = errors.New("file pull: unsafe source path")
	ErrPublish                       = errors.New("file pull: source publication failed")
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

// ProbeResult describes whether the saved-workbook backend can produce a
// complete source snapshot without publishing any files.
type ProbeResult struct {
	Supported bool
	Reason    string
	CodePage  uint16
	HasForms  bool
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

type inspectedProject struct {
	project   *vbaproject.Project
	container *cfb.Container
}

type extractedForm struct {
	name string
	code string
	spec forms.FormSpec
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

// Probe validates the saved-workbook snapshot needed for backend selection
// without reading or publishing the source tree. Source-tree validation remains
// in PullContext, where leases and wait policy protect managed source roots.
func Probe(workbookPath string, cfg config.Config) (ProbeResult, error) {
	inspection, err := inspectProject(workbookPath)
	if err != nil {
		return ProbeResult{}, err
	}
	result := ProbeResult{Supported: true, Reason: "supported", CodePage: inspection.project.Props.CodePage}
	for _, module := range inspection.project.Modules {
		if module.Type == vbaproject.ModuleForm {
			result.HasForms = true
			break
		}
	}
	if result.HasForms && !strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
		result.Supported = false
		result.Reason = "userform_code_source"
		return result, nil
	}
	if _, err := extractForms(inspection, cfg); err != nil {
		return ProbeResult{}, err
	}
	return result, nil
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
	managedRoots := []string{roots.modules, roots.classes, roots.workbook}
	if strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
		managedRoots = append(managedRoots, roots.forms)
	}
	for _, path := range managedRoots {
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
	inspection, err := inspectProject(workbookPath)
	if err != nil {
		return plan{}, err
	}
	return buildProjectPlan(root, cfg, workbookPath, inspection)
}

func inspectProject(workbookPath string) (*inspectedProject, error) {
	projectBytes, err := readVBAProject(workbookPath)
	if err != nil {
		return nil, err
	}
	project, err := vbaproject.Read(projectBytes)
	if err != nil {
		return nil, fmt.Errorf("%w: %v", ErrMalformedVBAProject, err)
	}
	if project.Protection.IsProtected {
		return nil, ErrProtectedProject
	}
	container, err := cfb.Open(projectBytes)
	if err != nil {
		return nil, fmt.Errorf("%w: reopen CFB container: %v", ErrMalformedVBAProject, err)
	}
	return &inspectedProject{project: project, container: container}, nil
}

func buildProjectPlan(root string, cfg config.Config, workbookPath string, inspection *inspectedProject) (plan, error) {
	project := inspection.project
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
	extractedForms, err := extractForms(inspection, cfg)
	if err != nil {
		return plan{}, err
	}
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

		if module.Type == vbaproject.ModuleForm {
			continue
		}

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

	for _, form := range extractedForms {
		specBody, err := forms.MarshalSnapshot("yaml", form.spec)
		if err != nil {
			return plan{}, fmt.Errorf("%w: form %q: %v", ErrUserFormDesignerUnsupported, form.name, err)
		}
		specTarget, err := safeTarget(roots.forms, filepath.Join("specs", form.name+".yaml"))
		if err != nil {
			return plan{}, err
		}
		if err := addPlannedFile(&p, seenPaths, desired, form.name, specTarget, specBody); err != nil {
			return plan{}, err
		}
		if form.code != "" {
			codeTarget, err := safeTarget(roots.forms, filepath.Join("code", form.name+".bas"))
			if err != nil {
				return plan{}, err
			}
			if err := addPlannedFile(&p, seenPaths, desired, form.name, codeTarget, []byte(form.code)); err != nil {
				return plan{}, err
			}
		}
		p.result.Modules.Form++
	}

	managedRoots := []struct {
		path string
		exts map[string]bool
	}{
		{roots.modules, map[string]bool{".bas": true}},
		{roots.classes, map[string]bool{".cls": true}},
		{roots.workbook, map[string]bool{".bas": true, ".cls": true}},
	}
	if strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
		managedRoots = append(managedRoots,
			struct {
				path string
				exts map[string]bool
			}{filepath.Join(roots.forms, "specs"), map[string]bool{".yaml": true, ".yml": true, ".json": true}},
			struct {
				path string
				exts map[string]bool
			}{filepath.Join(roots.forms, "code"), map[string]bool{".bas": true}},
		)
	}
	for _, managedRoot := range managedRoots {
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

func addPlannedFile(p *plan, seenPaths, desired map[string]string, component, target string, body []byte) error {
	pathKey := strings.ToLower(filepath.Clean(target))
	if prior, ok := seenPaths[pathKey]; ok {
		return fmt.Errorf("%w: %s and %s resolve to the same source path", ErrUnsafeSourcePath, prior, component)
	}
	seenPaths[pathKey] = component
	desired[pathKey] = target
	p.files = append(p.files, plannedFile{path: target, body: body})
	return nil
}

func extractForms(inspection *inspectedProject, cfg config.Config) ([]extractedForm, error) {
	formModules := make([]vbaproject.Module, 0)
	for _, module := range inspection.project.Modules {
		if module.Type == vbaproject.ModuleForm {
			formModules = append(formModules, module)
		}
	}
	storages := oforms.DiscoverForms(inspection.container)
	if len(formModules) == 0 && len(storages) == 0 {
		return nil, nil
	}
	if !strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") {
		return nil, fmt.Errorf("%w: configured code_source %q", ErrUserFormCodeSourceUnsupported, cfg.UserForm.CodeSource)
	}

	storageByName := make(map[string]string, len(storages))
	for _, storage := range storages {
		key := strings.ToLower(storage)
		if prior, exists := storageByName[key]; exists {
			return nil, fmt.Errorf("%w: duplicate Designer storages %q and %q", ErrUserFormIdentityMismatch, prior, storage)
		}
		storageByName[key] = storage
	}

	formsByName := make(map[string]extractedForm, len(formModules))
	for _, module := range formModules {
		storageIdentity := module.StreamName
		if storageIdentity == "" {
			storageIdentity = module.Name
		}
		storage, ok := storageByName[strings.ToLower(storageIdentity)]
		if !ok {
			return nil, fmt.Errorf("%w: form module %q has no Designer storage %q", ErrUserFormIdentityMismatch, module.Name, storageIdentity)
		}
		if !strings.EqualFold(module.Name, storage) {
			return nil, fmt.Errorf("%w: form module %q maps to Designer storage %q", ErrUserFormIdentityMismatch, module.Name, storage)
		}
		delete(storageByName, strings.ToLower(storageIdentity))

		parsed, err := oforms.ReadForm(inspection.container, storage, inspection.project.Props.CodePage)
		if err != nil {
			return nil, fmt.Errorf("%w: form %q: %v", ErrUserFormDesignerMalformed, module.Name, err)
		}
		spec, err := projection.Project(parsed)
		if err != nil {
			return nil, fmt.Errorf("%w: form %q: %v", ErrUserFormDesignerUnsupported, module.Name, err)
		}
		spec.Form.Name = module.Name
		spec = forms.NormalizeFormSpec(spec)
		code, err := exportFormCode(module)
		if err != nil {
			return nil, fmt.Errorf("%w: form %q code: %v", ErrUserFormIdentityMismatch, module.Name, err)
		}
		if cfg.VBA.LineNumbers.Enabled && code != "" {
			code, err = removeGeneratedLineNumbers(code)
			if err != nil {
				return nil, err
			}
		}
		key := strings.ToLower(module.Name)
		if _, exists := formsByName[key]; exists {
			return nil, fmt.Errorf("%w: duplicate form module %q", ErrUserFormIdentityMismatch, module.Name)
		}
		formsByName[key] = extractedForm{name: module.Name, code: code, spec: spec}
	}
	if len(storageByName) > 0 {
		return nil, fmt.Errorf("%w: Designer storage %q has no form module", ErrUserFormIdentityMismatch, slices.Sorted(maps.Values(storageByName))[0])
	}
	result := make([]extractedForm, 0, len(formsByName))
	for _, key := range slices.Sorted(maps.Keys(formsByName)) {
		result = append(result, formsByName[key])
	}
	return result, nil
}

func exportFormCode(module vbaproject.Module) (string, error) {
	source := strings.ReplaceAll(strings.ReplaceAll(module.Source, "\r\n", "\n"), "\r", "\n")
	if err := vbaproject.ValidateModuleIdentity(module.Name, source); err != nil {
		return "", err
	}
	lines := strings.Split(source, "\n")
	start := 0
	for start < len(lines) && strings.HasPrefix(strings.TrimSpace(lines[start]), "Attribute VB_") {
		start++
	}
	for start < len(lines) && strings.TrimSpace(lines[start]) == "" {
		start++
	}
	code := strings.TrimRight(strings.Join(lines[start:], "\n"), "\n")
	if code == "" {
		return "", nil
	}
	if !utf8.ValidString(code) {
		return "", fmt.Errorf("code-behind is not valid UTF-8")
	}
	return code + "\n", nil
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
