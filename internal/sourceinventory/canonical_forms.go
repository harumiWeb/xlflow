package sourceinventory

import (
	"encoding/json"
	"errors"
	"fmt"
	"io/fs"
	"maps"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/sourcepath"
	formpicture "github.com/harumiWeb/xlflow/internal/vba/userforms/picture"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

type canonicalFormFile struct {
	path string
	rel  string
	name string
	ext  string
}

// DiscoverPictureArtifacts resolves only canonical picture references so
// source preflights can exclude binary assets before inspecting VBA text.
// It does not classify code or compatibility artifacts.
func DiscoverPictureArtifacts(root string, cfg config.Config) ([]Artifact, error) {
	base, err := resolveRoot(root, cfg.Src.Forms, false)
	if err != nil {
		return nil, err
	}
	specDir := filepath.Join(base, "specs")
	entries, err := os.ReadDir(specDir)
	if errors.Is(err, fs.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}
	var result []Artifact
	for _, entry := range entries {
		ext := strings.ToLower(filepath.Ext(entry.Name()))
		if entry.IsDir() || ext != ".yaml" && ext != ".yml" && ext != ".json" {
			continue // The complete inventory validates layout separately.
		}
		path := filepath.Join(specDir, entry.Name())
		body, err := os.ReadFile(path)
		if err != nil {
			return nil, err
		}
		format := "yaml"
		if ext == ".json" {
			format = "json"
		}
		snapshot, err := loadAuthoredFormSpec(forms.SpecInput{Path: path, DisplayPath: displayPath(root, path), Format: format}, body)
		if err != nil {
			return nil, layoutError("%s: %w", path, err)
		}
		artifacts, _, err := loadFormPictureAssets(root, &snapshot)
		if err != nil {
			return nil, err
		}
		result = append(result, artifacts...)
	}
	return result, nil
}

func collectCanonicalForms(root, base string, opts Options) ([]Component, error) {
	groups := map[string]map[string]string{}
	specs := map[string]*forms.FormSpec{}
	specBodies := map[string][]byte{}
	legacyDirs := map[string]bool{}
	legacyDirsWithForms := map[string]bool{}
	pictureDirs := map[string]bool{}
	pictureFiles := map[string]bool{}
	var otherFiles []canonicalFormFile
	err := filepath.WalkDir(base, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			if path != base {
				// Pull deliberately retains assets after a spec stops referencing
				// them. The reserved directory is not a source-layout namespace;
				// referenced files are loaded directly from their FormSpecs.
				if samePath(path, filepath.Join(base, "assets")) {
					return filepath.SkipDir
				}
				if samePath(path, filepath.Join(base, "specs")) || samePath(path, filepath.Join(base, "code")) {
					return nil
				}
				parent := filepath.Dir(path)
				if samePath(parent, filepath.Join(base, "specs")) || samePath(parent, filepath.Join(base, "code")) {
					return layoutError("unsupported canonical form directory %s", displayPath(root, path))
				}
				legacyDirs[path] = true
			}
			return nil
		}
		rel, err := filepath.Rel(base, path)
		if err != nil {
			return err
		}
		parts := strings.Split(filepath.ToSlash(rel), "/")
		ext := strings.ToLower(filepath.Ext(path))
		name := strings.TrimSuffix(entry.Name(), filepath.Ext(path))
		if len(parts) > 1 && parts[0] == "specs" {
			if len(parts) != 2 || ext != ".yaml" && ext != ".yml" && ext != ".json" {
				return layoutError("unsupported canonical form source %s", displayPath(root, path))
			}
			if !ValidComponentName(name) {
				return layoutError("invalid form component %s", displayPath(root, path))
			}
			groupKey := canonicalFormGroupKey("", name)
			artifacts := groups[groupKey]
			if artifacts == nil {
				artifacts = map[string]string{}
				groups[groupKey] = artifacts
			}
			if prior := artifacts["spec"]; prior != "" {
				return layoutError("duplicate form artifact %s and %s", prior, path)
			}
			artifacts["spec"] = path
			format := "yaml"
			if ext == ".json" {
				format = "json"
			}
			input := forms.SpecInput{Path: path, DisplayPath: displayPath(root, path), Format: format}
			body, err := os.ReadFile(path)
			if err != nil {
				return fmt.Errorf("read form spec %s: %w", displayPath(root, path), err)
			}
			snapshot, err := loadAuthoredFormSpec(input, body)
			if err != nil {
				return layoutError("%s: %w", path, err)
			}
			if snapshot.Form.Name != name {
				return layoutError("spec filename %s differs from form.name %s", name, snapshot.Form.Name)
			}
			specs[groupKey] = &snapshot
			specBodies[groupKey] = body
			return nil
		}
		otherFiles = append(otherFiles, canonicalFormFile{path: path, rel: rel, name: name, ext: ext})
		return nil
	})
	if err != nil {
		return nil, err
	}

	// Resolve picture references before classifying other files. A picture's
	// FormSpec reference, not its filename suffix, determines its role; this
	// keeps e.g. a BMP named .bas out of sidecar and module discovery.
	pictureArtifactsByGroup := map[string][]Artifact{}
	referencedPictureFiles := map[string]bool{}
	for _, groupKey := range slices.Sorted(maps.Keys(specs)) {
		assets, referenced, err := loadFormPictureAssets(root, specs[groupKey])
		if err != nil {
			return nil, err
		}
		pictureArtifactsByGroup[groupKey] = assets
		for path := range referenced {
			referencedPictureFiles[path] = true
		}
	}
	for _, file := range otherFiles {
		pathKey := canonicalPathKey(file.path)
		if referencedPictureFiles[pathKey] {
			for dir := filepath.Dir(file.path); dir != base && dir != "."; dir = filepath.Dir(dir) {
				pictureDirs[dir] = true
			}
			continue
		}

		parts := strings.Split(filepath.ToSlash(file.rel), "/")
		if len(parts) > 1 && parts[0] == "code" {
			if len(parts) != 2 || file.ext != ".bas" {
				return nil, layoutError("unsupported canonical form source %s", displayPath(root, file.path))
			}
			if !ValidComponentName(file.name) {
				return nil, layoutError("invalid form component %s", displayPath(root, file.path))
			}
			groupKey := canonicalFormGroupKey("", file.name)
			artifacts := groups[groupKey]
			if artifacts == nil {
				artifacts = map[string]string{}
				groups[groupKey] = artifacts
			}
			if prior := artifacts["code"]; prior != "" {
				return nil, layoutError("duplicate form artifact %s and %s", prior, file.path)
			}
			artifacts["code"] = file.path
			continue
		}
		if file.ext == ".frm" || file.ext == ".frx" {
			if !ValidComponentName(file.name) {
				return nil, layoutError("invalid form component %s", displayPath(root, file.path))
			}
			relDir := filepath.ToSlash(filepath.Dir(file.rel))
			if relDir == "." {
				relDir = ""
			}
			groupKey := canonicalFormGroupKey(relDir, file.name)
			artifacts := groups[groupKey]
			if artifacts == nil {
				artifacts = map[string]string{}
				groups[groupKey] = artifacts
			}
			if prior := artifacts[file.ext]; prior != "" {
				return nil, layoutError("duplicate form artifact %s and %s", prior, file.path)
			}
			artifacts[file.ext] = file.path
			for dir := filepath.Dir(file.path); dir != base && dir != "."; dir = filepath.Dir(dir) {
				legacyDirsWithForms[dir] = true
			}
			continue
		}
		if (file.ext == ".bas" || file.ext == ".cls") && opts.AllowLooseFormModules {
			// Loose modules under the forms root are submitted separately by
			// filepush; they do not represent UserForm Designer artifacts.
			continue
		}
		// Any other unreferenced file is treated as a possible picture asset;
		// it is rejected below unless a FormSpec actually references it.
		pictureFiles[pathKey] = true
		for dir := filepath.Dir(file.path); dir != base && dir != "."; dir = filepath.Dir(dir) {
			pictureDirs[dir] = true
		}
	}

	for _, dir := range slices.Sorted(maps.Keys(legacyDirs)) {
		if !legacyDirsWithForms[dir] && !pictureDirs[dir] {
			return nil, layoutError("unsupported canonical form directory %s", displayPath(root, dir))
		}
	}

	result := make([]Component, 0, len(groups))
	if err := reconcileCanonicalLegacyArtifacts(groups); err != nil {
		return nil, err
	}
	for _, groupKey := range slices.Sorted(maps.Keys(groups)) {
		artifacts := maps.Clone(groups[groupKey])
		snapshot := specs[groupKey]
		if snapshot == nil && !opts.AllowLegacyFormArtifacts {
			return nil, layoutError("form %s has artifacts without canonical spec", formArtifactName(artifacts))
		}
		if snapshot == nil && artifacts[".frm"] == "" {
			return nil, layoutError("form %s has artifacts without canonical spec", formArtifactName(artifacts))
		}

		name := formArtifactName(artifacts)
		component := Component{
			Name:         name,
			Type:         ComponentForm,
			SourcePath:   displayPath(root, firstFormArtifact(artifacts)),
			AbsolutePath: artifacts[".frm"],
			FormSpec:     snapshot,
		}
		if snapshot != nil {
			name = snapshot.Form.Name
			component.Name = name
			component.SourcePath = displayPath(root, artifacts["spec"])
			component.AbsolutePath = artifacts["spec"]
		}
		if component.SourcePath == "" {
			component.SourcePath = displayPath(root, artifacts["code"])
			component.AbsolutePath = artifacts["code"]
		}

		useSidecar := strings.EqualFold(opts.Config.UserForm.CodeSource, "sidecar")
		ignoreCompatibility := opts.IgnoreCanonicalCompatibilityArtifacts && snapshot != nil && useSidecar
		for _, kind := range []string{"spec", "code", ".frm", ".frx"} {
			path := artifacts[kind]
			if path != "" && strings.TrimSuffix(filepath.Base(path), filepath.Ext(path)) != name {
				return nil, layoutError("form artifact identity differs from %s: %s", name, path)
			}
		}
		if ignoreCompatibility {
			// Stale compatibility files are outside both import authority and
			// dependency fingerprints. Validate their identity above without
			// opening their contents.
			artifacts[".frm"] = ""
			artifacts[".frx"] = ""
		}

		bodies := map[string][]byte{}
		for _, kind := range []string{"spec", "code", ".frm", ".frx"} {
			path := artifacts[kind]
			if path == "" {
				continue
			}
			body := specBodies[groupKey]
			if kind != "spec" {
				var err error
				body, err = os.ReadFile(path)
				if err != nil {
					return nil, fmt.Errorf("read form artifact %s: %w", displayPath(root, path), err)
				}
			}
			bodies[path] = body
			component.Related = append(component.Related, Artifact{
				Path:         displayPath(root, path),
				AbsolutePath: path,
				Source:       body,
			})
		}

		switch {
		case useSidecar && artifacts["code"] != "":
			body := bodies[artifacts["code"]]
			for line := range strings.SplitSeq(string(body), "\n") {
				if strings.HasPrefix(strings.TrimSpace(line), "Attribute VB_") {
					return nil, layoutError("sidecar contains Attribute VB_ header: %s", artifacts["code"])
				}
			}
			if snapshot != nil {
				component.Source = body
			} else {
				component.Source = bodies[artifacts[".frm"]]
			}
		case opts.PreserveMissingFormCode && useSidecar:
			component.PreserveExistingCode = true
		case artifacts[".frm"] != "":
			if snapshot != nil && formSpecHasUnsynchronizedCompatibilityArtifact(snapshot) {
				return nil, layoutError("FRM201: %s.frm is marked compatibility_artifact_unsynchronized and cannot provide UserForm code", name)
			}
			body := bodies[artifacts[".frm"]]
			if snapshot != nil {
				if err := vbaproject.ValidateModuleIdentity(name, string(body)); err != nil {
					return nil, layoutError("frm identity does not match %s: %w", name, err)
				}
				component.Source = []byte(forms.ExtractUserFormCodeFromFRM(string(body)))
			} else {
				component.Source = body
			}
		case snapshot != nil && !useSidecar:
			return nil, layoutError("frm code mode requires %s.frm", name)
		}
		component.Related = append(component.Related, pictureArtifactsByGroup[groupKey]...)
		result = append(result, component)
	}
	if len(pictureFiles) > 0 {
		for _, path := range slices.Sorted(maps.Keys(pictureFiles)) {
			return nil, layoutError("unreferenced UserForm picture asset %s", displayPath(root, path))
		}
	}
	return result, nil
}

func formSpecHasUnsynchronizedCompatibilityArtifact(snapshot *forms.FormSpec) bool {
	if snapshot == nil {
		return false
	}
	for _, warning := range snapshot.Warnings {
		if warning.Code == forms.CompatibilityArtifactUnsynchronizedWarningCode {
			return true
		}
	}
	return false
}

func canonicalPathKey(path string) string {
	return sourcepath.Key(path)
}

func loadFormPictureAssets(root string, snapshot *forms.FormSpec) ([]Artifact, map[string]bool, error) {
	assets := map[string]Artifact{}
	referencedFiles := map[string]bool{}
	var visit func([]forms.FormSpecControl) error
	visit = func(controls []forms.FormSpecControl) error {
		for i := range controls {
			control := &controls[i]
			if control.Picture != nil {
				pictureSpec := control.Picture
				if pictureSpec.Remove {
					if strings.TrimSpace(pictureSpec.Path) != "" {
						return layoutError("UserForm control %s picture removal cannot include a path", control.Name)
					}
					pictureSpec.Data = nil
				} else {
					if strings.TrimSpace(pictureSpec.Path) == "" {
						return layoutError("UserForm control %s picture path is empty", control.Name)
					}
					asset, err := formpicture.Load(root, pictureSpec.Path)
					if err != nil {
						if errors.Is(err, fs.ErrNotExist) || errors.Is(err, fs.ErrPermission) {
							return fmt.Errorf("load UserForm picture %q: %w", pictureSpec.Path, err)
						}
						return layoutError("load UserForm picture %q: %v", pictureSpec.Path, err)
					}
					pictureSpec.Data = slices.Clone(asset.Data)
					logicalPath := picturePathFromProjectRoot(root, pictureSpec.Path)
					logicalAbsolute, err := filepath.Abs(logicalPath)
					if err != nil {
						return fmt.Errorf("resolve UserForm picture %q: %w", pictureSpec.Path, err)
					}
					physicalPath, err := filepath.EvalSymlinks(logicalAbsolute)
					if err != nil {
						return fmt.Errorf("resolve UserForm picture %q: %w", pictureSpec.Path, err)
					}
					display := displayPath(root, logicalAbsolute)
					key := canonicalPathKey(logicalAbsolute)
					assets[key] = Artifact{Path: display, AbsolutePath: filepath.Clean(physicalPath), Source: slices.Clone(asset.Data), Role: ArtifactRolePicture}
					referencedFiles[canonicalPathKey(logicalAbsolute)] = true
				}
			}
			if err := visit(control.Controls); err != nil {
				return err
			}
		}
		return nil
	}
	if err := visit(snapshot.Controls); err != nil {
		return nil, nil, err
	}
	keys := slices.Sorted(maps.Keys(assets))
	result := make([]Artifact, 0, len(keys))
	for _, key := range keys {
		result = append(result, assets[key])
	}
	return result, referencedFiles, nil
}

func picturePathFromProjectRoot(root, path string) string {
	return filepath.Clean(filepath.Join(root, filepath.FromSlash(strings.ReplaceAll(path, "\\", "/"))))
}

func canonicalFormGroupKey(directory, name string) string {
	return strings.ToLower(filepath.ToSlash(directory)) + "\x00" + strings.ToLower(name)
}

func formArtifactName(artifacts map[string]string) string {
	for _, kind := range []string{"spec", "code", ".frm", ".frx"} {
		if path := artifacts[kind]; path != "" {
			return strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
		}
	}
	return ""
}

func firstFormArtifact(artifacts map[string]string) string {
	for _, kind := range []string{".frm", "spec", "code", ".frx"} {
		if path := artifacts[kind]; path != "" {
			return path
		}
	}
	return ""
}

// reconcileCanonicalLegacyArtifacts preserves collectForms' name-based
// sidecar association for legacy forms while keeping canonical specs and
// sidecars flat under their reserved roots. A nested legacy form is eligible
// only when its name is unique; otherwise silently choosing a Designer would
// make the source layout ambiguous.
func reconcileCanonicalLegacyArtifacts(groups map[string]map[string]string) error {
	byName := map[string][]string{}
	for _, key := range slices.Sorted(maps.Keys(groups)) {
		artifacts := groups[key]
		if artifacts[".frm"] == "" {
			continue
		}
		byName[strings.ToLower(formArtifactName(artifacts))] = append(byName[strings.ToLower(formArtifactName(artifacts))], key)
	}
	for _, key := range slices.Sorted(maps.Keys(groups)) {
		artifacts := groups[key]
		if artifacts[".frm"] != "" || (artifacts["spec"] == "" && artifacts["code"] == "") {
			continue
		}
		name := strings.ToLower(formArtifactName(artifacts))
		candidates := byName[name]
		if len(candidates) == 0 {
			continue
		}
		if len(candidates) > 1 {
			return layoutError("ambiguous legacy form %s: multiple nested .frm artifacts match canonical artifacts", formArtifactName(artifacts))
		}
		legacyKey := candidates[0]
		legacy := groups[legacyKey]
		for _, kind := range []string{".frm", ".frx"} {
			if artifacts[kind] == "" {
				artifacts[kind] = legacy[kind]
				legacy[kind] = ""
			}
		}
		if legacy[".frm"] == "" && legacy[".frx"] == "" && legacy["spec"] == "" && legacy["code"] == "" {
			delete(groups, legacyKey)
		}
	}
	return nil
}

// loadAuthoredFormSpec retains normalized topology while restoring the
// author's explicit form fields. LoadFormSpec synthesizes compatibility
// observed/build fields for general consumers; pack must not turn an omitted
// property into a Designer edit.
func loadAuthoredFormSpec(input forms.SpecInput, body []byte) (forms.FormSpec, error) {
	sourceIssues, err := forms.ValidateFormSpecSource(input, body)
	if err != nil {
		return forms.FormSpec{}, err
	}
	for _, issue := range sourceIssues {
		if issue.Severity == forms.SeverityError {
			return forms.FormSpec{}, fmt.Errorf("%s: %s", issue.Code, issue.Message)
		}
	}
	var raw forms.FormSpec
	switch input.Format {
	case "json":
		err = json.Unmarshal(body, &raw)
	case "yaml":
		err = yaml.Unmarshal(body, &raw)
	default:
		err = fmt.Errorf("unsupported form spec format %q", input.Format)
	}
	if err != nil {
		return forms.FormSpec{}, err
	}
	normalized := forms.NormalizeFormSpec(raw)
	if err := forms.ValidateFormSpec(normalized); err != nil {
		return forms.FormSpec{}, err
	}
	normalized.Form.Caption = raw.Form.Caption
	normalized.Form.Width = raw.Form.Width
	normalized.Form.Height = raw.Form.Height
	normalized.Form.Observed = raw.Form.Observed
	normalized.Form.Build = raw.Form.Build
	for i := range normalized.Controls {
		// Control fields authored at the top level are retained by normalization;
		// only its synthesized observed compatibility view is discarded.
		normalized.Controls[i].Observed = nil
	}
	for _, issue := range sourceIssues {
		if issue.Severity == forms.SeverityWarning {
			normalized.ValidationWarnings = append(normalized.ValidationWarnings, issue)
		}
	}
	return normalized, nil
}
