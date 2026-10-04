package filepush

import (
	"errors"
	"fmt"
	"maps"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

type sourceSnapshot struct {
	modules   []packpkg.SourceModule
	files     []discoveredFile
	assetDirs []string
}

func captureSourceSnapshot(root string, cfg config.Config) (sourceSnapshot, error) {
	rawFiles, err := discoverSourceFiles(resolvedRoots(root, cfg), cfg.UserForm.CodeSource)
	if err != nil {
		return sourceSnapshot{}, err
	}
	return collectSources(root, cfg, rawFiles)
}

func samePathLists(left, right []string) bool {
	if len(left) != len(right) {
		return false
	}
	left = slices.Clone(left)
	right = slices.Clone(right)
	slices.SortFunc(left, func(a, b string) int { return strings.Compare(strings.ToLower(a), strings.ToLower(b)) })
	slices.SortFunc(right, func(a, b string) int { return strings.Compare(strings.ToLower(a), strings.ToLower(b)) })
	for i := range left {
		if !strings.EqualFold(filepath.Clean(left[i]), filepath.Clean(right[i])) {
			return false
		}
	}
	return true
}

func findDuplicateSourceModules(modules []packpkg.SourceModule) [][]string {
	seen := map[string][]string{}
	for _, module := range modules {
		key := strings.ToLower(module.Name)
		path := module.SourcePath
		if path == "" {
			path = module.Name
		}
		seen[key] = append(seen[key], path)
	}
	var duplicates [][]string
	for _, key := range slices.Sorted(maps.Keys(seen)) {
		if len(seen[key]) > 1 {
			paths := slices.Clone(seen[key])
			slices.Sort(paths)
			duplicates = append(duplicates, paths)
		}
	}
	return duplicates
}

// collectSources resolves one canonical source snapshot for packing and
// fingerprinting. Captured artifact bytes are reused by both operations.
func collectSources(root string, cfg config.Config, rawFiles []discoveredFile) (sourceSnapshot, error) {
	components, err := sourceinventory.Discover(sourceinventory.Options{
		Root: root, Config: cfg, ValidateFormArtifacts: true,
		CanonicalFormSpecs: true, AllowLegacyFormArtifacts: true,
		IgnoreCanonicalCompatibilityArtifacts: true, PreserveMissingFormCode: true,
		AllowLooseFormModules: true, AllowMissingRoots: true,
	})
	if err != nil {
		var layoutErr *sourceinventory.LayoutError
		if errors.As(err, &layoutErr) {
			return sourceSnapshot{}, fmt.Errorf("%w: %v", packpkg.ErrAmbiguousLayout, err)
		}
		return sourceSnapshot{}, err
	}
	roots := resolvedRoots(root, cfg)
	sidecar := isSidecarMode(cfg.UserForm.CodeSource)
	lineNumbers := cfg.VBA.LineNumbers.Enabled
	sources := make([]packpkg.SourceModule, 0, len(components))
	inputs := map[string]discoveredFile{}
	assetDirs := map[string]string{}
	addInput := func(kind, path, fullPath string, body []byte) {
		abs := fullPath
		if abs == "" {
			abs = filepath.Join(root, filepath.FromSlash(path))
		}
		abs, _ = filepath.Abs(abs)
		abs = filepath.Clean(abs)
		inputPath := abs
		if kind == "form_asset" {
			// Discovery enumerates the logical reference, while the asset's
			// FullPath retains its resolved target for reads and coordination.
			inputPath = filepath.FromSlash(path)
			if !filepath.IsAbs(inputPath) {
				inputPath = filepath.Join(root, inputPath)
			}
			inputPath, _ = filepath.Abs(inputPath)
		}
		relative := filepath.ToSlash(path)
		inputs[canonicalPathKey(inputPath)] = discoveredFile{
			Kind: kind, RootDir: filepath.Dir(abs), FullPath: abs,
			RelativePath: relative, Extension: strings.ToLower(filepath.Ext(path)),
			ModuleName: strings.TrimSuffix(filepath.Base(path), filepath.Ext(path)),
			Body:       slices.Clone(body), HasBody: true,
		}
	}
	for _, component := range components {
		if component.Type == sourceinventory.ComponentDocument && !strings.EqualFold(filepath.Ext(component.AbsolutePath), ".bas") {
			// The Excel bridge only applies document modules discovered as
			// <name>.bas under the workbook root; a .cls there is never
			// fingerprinted or imported. Skip it for the same outcome.
			continue
		}
		source := string(component.Source)
		sidecarText, hasSidecar := "", false
		if component.Type == sourceinventory.ComponentForm && sidecar {
			for _, artifact := range component.Related {
				if artifact.Role != sourceinventory.ArtifactRolePicture && strings.EqualFold(filepath.Ext(artifact.Path), ".bas") && strings.EqualFold(filepath.Base(artifact.Path), component.Name+".bas") {
					sidecarText, hasSidecar = string(artifact.Source), true
					break
				}
			}
		}
		var transformed string
		if component.Type == sourceinventory.ComponentForm && component.FormSpec != nil {
			// Canonical sidecar bytes already contain code only. They are
			// instrumented directly; stale .frm compatibility text is never
			// merged into the code-authoritative source.
			if component.PreserveExistingCode {
				transformed = ""
			} else if sidecar {
				transformed, err = instrumentFormSidecar(component.SourcePath, source, lineNumbers)
			} else {
				transformed, err = transformComponentSource(component, source, "", false, componentRootDir(component.Type, roots), cfg, lineNumbers)
			}
		} else {
			transformed, err = transformComponentSource(component, source, sidecarText, hasSidecar, componentRootDir(component.Type, roots), cfg, lineNumbers)
		}
		if err != nil {
			return sourceSnapshot{}, err
		}
		sources = append(sources, packpkg.SourceModule{
			SourcePath:           component.SourcePath,
			RelatedPaths:         component.RelatedPaths(),
			Name:                 component.Name,
			Type:                 packpkg.ModuleType(component.Type),
			Source:               transformed,
			FormSpec:             component.FormSpec,
			PreserveExistingCode: component.PreserveExistingCode,
		})
		if component.Type != sourceinventory.ComponentForm {
			addInput(sourceKind(component.Type), component.SourcePath, component.AbsolutePath, component.Source)
		}
		for _, artifact := range component.Related {
			kind := relatedSourceKind(artifact, component.Type, sidecar)
			addInput(kind, artifact.Path, artifact.AbsolutePath, artifact.Source)
			if kind == "form_asset" {
				dir := filepath.Dir(artifact.AbsolutePath)
				assetDirs[canonicalPathKey(dir)] = dir
			}
		}
	}
	// Loose .bas/.cls files under the forms root are fingerprinted as "form"
	// entries and imported by the Excel bridge as standard/class modules via
	// VBIDE, so submit them to pack with the matching component types. The
	// inventory skips them because they carry no UserForm artifact role.
	for _, file := range rawFiles {
		if file.Kind != "form" || (file.Extension != ".bas" && file.Extension != ".cls") {
			continue
		}
		if input, ok := inputs[canonicalPathKey(file.FullPath)]; ok && input.Kind == "form_asset" {
			// A FormSpec reference gives this path picture authority even when
			// its suffix would otherwise make bridge discovery call it a module.
			continue
		}
		body, err := os.ReadFile(file.FullPath)
		if err != nil {
			return sourceSnapshot{}, &SourceReadError{Path: file.FullPath, Err: err}
		}
		loose := sourceinventory.Component{
			SourcePath:   looseFormSourcePath(root, file.FullPath),
			AbsolutePath: file.FullPath,
			Name:         file.ModuleName,
			Type:         sourceinventory.ComponentStandard,
			Source:       body,
		}
		if file.Extension == ".cls" {
			loose.Type = sourceinventory.ComponentClass
		}
		transformed, err := transformComponentSource(loose, string(body), "", false, roots.forms, cfg, lineNumbers)
		if err != nil {
			return sourceSnapshot{}, err
		}
		sources = append(sources, packpkg.SourceModule{
			SourcePath: loose.SourcePath,
			Name:       loose.Name,
			Type:       packpkg.ModuleType(loose.Type),
			Source:     transformed,
		})
		addInput(file.Kind, file.RelativePath, file.FullPath, body)
	}

	files := make([]discoveredFile, 0, len(rawFiles)+len(inputs))
	seen := map[string]bool{}
	for _, raw := range rawFiles {
		key := canonicalPathKey(raw.FullPath)
		if input, ok := inputs[key]; ok {
			if input.Kind != "form_asset" {
				input.Kind = raw.Kind
				input.RootDir = raw.RootDir
				input.RelativePath = raw.RelativePath
			}
			files = append(files, input)
			seen[key] = true
			continue
		}
		if raw.Kind == "form" && (raw.Extension == ".frm" || raw.Extension == ".frx") && sidecar {
			// Canonical filepush ignores stale compatibility artifacts. Legacy
			// forms remain represented by the canonical inventory above.
			continue
		}
		if raw.Kind == "form" && (raw.Extension == ".bas" || raw.Extension == ".cls") {
			body, readErr := os.ReadFile(raw.FullPath)
			if readErr != nil {
				return sourceSnapshot{}, &SourceReadError{Path: raw.FullPath, Err: readErr}
			}
			files = append(files, discoveredFile{
				Kind: raw.Kind, RootDir: raw.RootDir, FullPath: raw.FullPath,
				RelativePath: raw.RelativePath, Extension: raw.Extension,
				ModuleName: raw.ModuleName, Body: body, HasBody: true,
			})
			seen[key] = true
			continue
		}
		return sourceSnapshot{}, fmt.Errorf("%w: discovered source %s is absent from canonical inventory", packpkg.ErrAmbiguousLayout, raw.RelativePath)
	}
	for key, input := range inputs {
		if !seen[key] {
			files = append(files, input)
		}
	}
	slices.SortFunc(files, func(a, b discoveredFile) int {
		if cmp := strings.Compare(a.Kind, b.Kind); cmp != 0 {
			return cmp
		}
		return strings.Compare(a.RelativePath, b.RelativePath)
	})
	assetDirList := make([]string, 0, len(assetDirs))
	for _, dir := range assetDirs {
		assetDirList = append(assetDirList, dir)
	}
	slices.Sort(assetDirList)
	return sourceSnapshot{modules: sources, files: files, assetDirs: assetDirList}, nil
}

func sourceKind(typ sourceinventory.ComponentType) string {
	switch typ {
	case sourceinventory.ComponentStandard:
		return "module"
	case sourceinventory.ComponentClass:
		return "class"
	case sourceinventory.ComponentDocument:
		return "document"
	default:
		return "form"
	}
}

func relatedSourceKind(artifact sourceinventory.Artifact, typ sourceinventory.ComponentType, sidecar bool) string {
	if artifact.Role == sourceinventory.ArtifactRolePicture {
		return "form_asset"
	}
	ext := strings.ToLower(filepath.Ext(artifact.Path))
	switch {
	case ext == ".json" || ext == ".yaml" || ext == ".yml":
		return "form_spec"
	case ext == ".frm" || ext == ".frx":
		return "form"
	case ext == ".bas" && typ == sourceinventory.ComponentForm && sidecar:
		return "form_code"
	default:
		return sourceKind(typ)
	}
}

func instrumentFormSidecar(path, source string, enabled bool) (string, error) {
	if !enabled {
		return source, nil
	}
	numbered, issue := tryAddLineNumbers(source)
	if issue != nil {
		return "", fmt.Errorf("%w: %s sidecar: %d: %s", ErrLineNumberSafety, path, issue.Line, issue.Message)
	}
	return numbered, nil
}

func looseFormSourcePath(root, fullPath string) string {
	rel, err := filepath.Rel(root, fullPath)
	if err != nil {
		return fullPath
	}
	return filepath.ToSlash(rel)
}

// transformComponentSource applies push-time text transforms in the same order
// the bridge applies them: for document modules the exported-attribute
// normalization runs first (UpdateDocumentModules), then the folder annotation
// update, then Erl instrumentation (PrepareSourceForImport). In sidecar mode a
// UserForm's code-behind comes from its code/<Name>.bas file verbatim (the
// bridge replaces the imported code module with the instrumented sidecar
// text), so the sidecar is numbered but never re-annotated.
func transformComponentSource(component sourceinventory.Component, source, sidecarText string, hasSidecar bool, annotationRoot string, cfg config.Config, lineNumbers bool) (string, error) {
	if component.Type == sourceinventory.ComponentDocument {
		source = normalizeDocumentModuleContent(source)
	}
	if annotationRoot != "" {
		source = updateFolderAnnotationText(source, cfg.VBA.FolderAnnotation, folderAnnotationForPath(annotationRoot, component.AbsolutePath))
	}
	if lineNumbers {
		numbered, issue := tryAddLineNumbers(source)
		if issue != nil {
			return "", fmt.Errorf("%w: %s:%d: %s", ErrLineNumberSafety, component.SourcePath, issue.Line, issue.Message)
		}
		source = numbered
	}
	if !hasSidecar {
		return source, nil
	}
	code := sidecarText
	if lineNumbers {
		numbered, issue := tryAddLineNumbers(code)
		if issue != nil {
			return "", fmt.Errorf("%w: %s sidecar: %d: %s", ErrLineNumberSafety, component.SourcePath, issue.Line, issue.Message)
		}
		code = numbered
	}
	return forms.MergeUserFormCodeIntoFRM(source, code), nil
}

func componentRootDir(typ sourceinventory.ComponentType, roots sourceRoots) string {
	switch typ {
	case sourceinventory.ComponentStandard:
		return roots.modules
	case sourceinventory.ComponentClass:
		return roots.classes
	case sourceinventory.ComponentForm:
		return roots.forms
	case sourceinventory.ComponentDocument:
		return roots.workbook
	default:
		return ""
	}
}

// normalizeDocumentModuleContent ports VbaSourceHelper.NormalizeDocumentModuleContent:
// document-module disk sources are pure code, but hand-authored files may carry
// a class header or Attribute VB_* lines; both are dropped, and an empty body
// becomes the canonical "Option Explicit" module the bridge would produce.
func normalizeDocumentModuleContent(text string) string {
	lines := splitLinesKeepEmpty(text)
	filtered := make([]string, 0, len(lines))
	inClassHeader := false
	var classHeaderBuffer []string
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "VERSION 1.0 CLASS" {
			inClassHeader = true
			classHeaderBuffer = []string{line}
			continue
		}
		if inClassHeader {
			classHeaderBuffer = append(classHeaderBuffer, line)
			if trimmed == "END" {
				inClassHeader = false
				classHeaderBuffer = nil
			}
			continue
		}
		if attributeVBPattern.MatchString(trimmed) {
			continue
		}
		filtered = append(filtered, line)
	}
	if inClassHeader && len(classHeaderBuffer) > 0 {
		filtered = append(filtered, classHeaderBuffer...)
	}

	hasOptionExplicit := false
	hasNonHeaderCode := false
	for _, line := range filtered {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" {
			continue
		}
		if strings.EqualFold(trimmed, "Option Explicit") {
			hasOptionExplicit = true
			continue
		}
		hasNonHeaderCode = true
	}
	if !hasOptionExplicit && !hasNonHeaderCode {
		filtered = []string{"Option Explicit"}
	}
	return strings.Join(filtered, "\r\n")
}
