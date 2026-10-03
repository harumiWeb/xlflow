package sourceinventory

import (
	"encoding/json"
	"fmt"
	"io/fs"
	"maps"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

func collectCanonicalForms(root, base string, opts Options) ([]Component, error) {
	groups := map[string]map[string]string{}
	specs := map[string]*forms.FormSpec{}
	legacyDirs := map[string]bool{}
	legacyDirsWithForms := map[string]bool{}
	err := filepath.WalkDir(base, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			if path != base {
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
		if !ValidComponentName(name) {
			return layoutError("invalid form component %s", displayPath(root, path))
		}

		kind := ""
		groupKey := ""
		switch {
		case len(parts) == 2 && parts[0] == "specs" && (ext == ".yaml" || ext == ".yml" || ext == ".json"):
			kind = "spec"
			groupKey = canonicalFormGroupKey("", name)
		case len(parts) == 2 && parts[0] == "code" && ext == ".bas":
			kind = "code"
			groupKey = canonicalFormGroupKey("", name)
		case ext == ".frm" || ext == ".frx":
			if len(parts) > 1 && (parts[0] == "specs" || parts[0] == "code") {
				return layoutError("unsupported canonical form source %s", displayPath(root, path))
			}
			relDir := filepath.ToSlash(filepath.Dir(rel))
			if relDir == "." {
				relDir = ""
			}
			kind = ext
			groupKey = canonicalFormGroupKey(relDir, name)
			for dir := filepath.Dir(path); dir != base && dir != "."; dir = filepath.Dir(dir) {
				legacyDirsWithForms[dir] = true
			}
		default:
			return layoutError("unsupported canonical form source %s", displayPath(root, path))
		}

		artifacts := groups[groupKey]
		if artifacts == nil {
			artifacts = map[string]string{}
			groups[groupKey] = artifacts
		}
		if prior := artifacts[kind]; prior != "" {
			return layoutError("duplicate form artifact %s and %s", prior, path)
		}
		artifacts[kind] = path
		if kind != "spec" {
			return nil
		}

		format := "yaml"
		if ext == ".json" {
			format = "json"
		}
		input := forms.SpecInput{Path: path, DisplayPath: displayPath(root, path), Format: format}
		snapshot, err := loadAuthoredFormSpec(input)
		if err != nil {
			return layoutError("%s: %w", path, err)
		}
		if snapshot.Form.Name != name {
			return layoutError("spec filename %s differs from form.name %s", name, snapshot.Form.Name)
		}
		specs[groupKey] = &snapshot
		return nil
	})
	if err != nil {
		return nil, err
	}
	for _, dir := range slices.Sorted(maps.Keys(legacyDirs)) {
		if !legacyDirsWithForms[dir] {
			return nil, layoutError("unsupported canonical form directory %s", displayPath(root, dir))
		}
	}

	result := make([]Component, 0, len(groups))
	if err := reconcileCanonicalLegacyArtifacts(groups); err != nil {
		return nil, err
	}
	for _, groupKey := range slices.Sorted(maps.Keys(groups)) {
		artifacts := groups[groupKey]
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

		for _, kind := range []string{"spec", "code", ".frm", ".frx"} {
			path := artifacts[kind]
			if path == "" {
				continue
			}
			if strings.TrimSuffix(filepath.Base(path), filepath.Ext(path)) != name {
				return nil, layoutError("form artifact identity differs from %s: %s", name, path)
			}
			body, err := os.ReadFile(path)
			if err != nil {
				return nil, err
			}
			component.Related = append(component.Related, Artifact{
				Path:         displayPath(root, path),
				AbsolutePath: path,
				Source:       body,
			})
		}

		useSidecar := strings.EqualFold(opts.Config.UserForm.CodeSource, "sidecar")
		switch {
		case useSidecar && artifacts["code"] != "":
			body, err := os.ReadFile(artifacts["code"])
			if err != nil {
				return nil, err
			}
			for line := range strings.SplitSeq(string(body), "\n") {
				if strings.HasPrefix(strings.TrimSpace(line), "Attribute VB_") {
					return nil, layoutError("sidecar contains Attribute VB_ header: %s", artifacts["code"])
				}
			}
			if snapshot != nil {
				component.Source = body
			} else {
				component.Source, err = os.ReadFile(artifacts[".frm"])
				if err != nil {
					return nil, err
				}
			}
		case artifacts[".frm"] != "":
			body, err := os.ReadFile(artifacts[".frm"])
			if err != nil {
				return nil, err
			}
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
		result = append(result, component)
	}
	return result, nil
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
func loadAuthoredFormSpec(input forms.SpecInput) (forms.FormSpec, error) {
	normalized, err := forms.LoadFormSpec(input)
	if err != nil {
		return forms.FormSpec{}, err
	}
	body, err := os.ReadFile(input.Path)
	if err != nil {
		return forms.FormSpec{}, err
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
	return normalized, nil
}
