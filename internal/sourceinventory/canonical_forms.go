package sourceinventory

import (
	"io/fs"
	"maps"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"

	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func collectCanonicalForms(root, base string, opts Options) ([]Component, error) {
	byName := map[string]*Component{}
	paths := map[string]map[string]string{}
	err := filepath.WalkDir(base, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			if path != base && path != filepath.Join(base, "specs") && path != filepath.Join(base, "code") {
				return layoutError("unsupported canonical form directory %s", displayPath(root, path))
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
		switch {
		case len(parts) == 2 && parts[0] == "specs" && (ext == ".yaml" || ext == ".yml" || ext == ".json"):
			kind = "spec"
		case len(parts) == 2 && parts[0] == "code" && ext == ".bas":
			kind = "code"
		case len(parts) == 1 && (ext == ".frm" || ext == ".frx"):
			kind = ext
		default:
			return layoutError("unsupported canonical form source %s", displayPath(root, path))
		}
		key := strings.ToLower(name)
		if paths[key] == nil {
			paths[key] = map[string]string{}
		}
		if prior := paths[key][kind]; prior != "" {
			return layoutError("duplicate form artifact %s and %s", prior, path)
		}
		paths[key][kind] = path
		if kind != "spec" {
			return nil
		}
		format := "yaml"
		if ext == ".json" {
			format = "json"
		}
		input := forms.SpecInput{Path: path, DisplayPath: displayPath(root, path), Format: format}
		snapshot, err := forms.LoadFormSpec(input)
		if err != nil {
			return layoutError("%s: %w", path, err)
		}
		if snapshot.Form.Name != name {
			return layoutError("spec filename %s differs from form.name %s", name, snapshot.Form.Name)
		}
		byName[key] = &Component{Name: name, Type: ComponentForm, SourcePath: displayPath(root, path), AbsolutePath: path, FormSpec: &snapshot}
		return nil
	})
	if err != nil {
		return nil, err
	}
	var result []Component
	for _, key := range slices.Sorted(maps.Keys(paths)) {
		component := byName[key]
		if component == nil {
			return nil, layoutError("form %s has artifacts without canonical spec", key)
		}
		artifacts := paths[key]
		for _, kind := range []string{"code", ".frm", ".frx"} {
			path := artifacts[kind]
			if path == "" {
				continue
			}
			if strings.TrimSuffix(filepath.Base(path), filepath.Ext(path)) != component.Name {
				return nil, layoutError("form artifact identity differs from %s: %s", component.Name, path)
			}
			body, err := os.ReadFile(path)
			if err != nil {
				return nil, err
			}
			component.Related = append(component.Related, Artifact{Path: displayPath(root, path), AbsolutePath: path, Source: body})
		}
		useSidecar := strings.EqualFold(opts.Config.UserForm.CodeSource, "sidecar")
		if useSidecar && artifacts["code"] != "" {
			body, err := os.ReadFile(artifacts["code"])
			if err != nil {
				return nil, err
			}
			for line := range strings.SplitSeq(string(body), "\n") {
				if strings.HasPrefix(strings.TrimSpace(line), "Attribute VB_") {
					return nil, layoutError("sidecar contains Attribute VB_ header: %s", artifacts["code"])
				}
			}
			component.Source = body
		} else if artifacts[".frm"] != "" {
			body, err := os.ReadFile(artifacts[".frm"])
			if err != nil {
				return nil, err
			}
			if err := vbaproject.ValidateModuleIdentity(component.Name, string(body)); err != nil {
				return nil, layoutError("frm identity does not match %s: %w", component.Name, err)
			}
			component.Source = []byte(forms.ExtractUserFormCodeFromFRM(string(body)))
		} else if !useSidecar {
			return nil, layoutError("frm code mode requires %s.frm", component.Name)
		}
		result = append(result, *component)
	}
	return result, nil
}
