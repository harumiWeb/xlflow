// Package sourceinventory discovers the complete, read-only VBA source tree.
// It is shared by build and pack so both commands validate the same source
// layout before any workbook-backed work begins.
package sourceinventory

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
)

type ComponentType string

const (
	ComponentStandard ComponentType = "standard"
	ComponentClass    ComponentType = "class"
	ComponentDocument ComponentType = "document"
	ComponentForm     ComponentType = "form"
)

type Artifact struct {
	Path         string
	AbsolutePath string
	Source       []byte
}

type Component struct {
	SourcePath   string
	AbsolutePath string
	Name         string
	Type         ComponentType
	Source       []byte
	Related      []Artifact
}

func (c Component) RelatedPaths() []string {
	paths := make([]string, len(c.Related))
	for i, artifact := range c.Related {
		paths[i] = artifact.Path
	}
	return paths
}

func (c Component) Artifact(path string) ([]byte, bool) {
	for _, artifact := range c.Related {
		if strings.EqualFold(artifact.Path, path) {
			return artifact.Source, true
		}
	}
	return nil, false
}

type Options struct {
	Root                  string
	Config                config.Config
	RestrictToRoot        bool
	ValidateFormArtifacts bool
	// PrimaryExcluded is used by build to resolve flat UserForm sidecars only
	// against forms whose primary .frm survives primary-path exclusions.
	PrimaryExcluded func(path string) bool
}

type formFiles struct {
	name    string
	frm     string
	frx     string
	related []string
}

func Discover(opts Options) ([]Component, error) {
	root, err := filepath.Abs(opts.Root)
	if err != nil {
		return nil, fmt.Errorf("resolve project root: %w", err)
	}
	root = filepath.Clean(root)
	var components []Component
	for _, source := range []struct {
		dir  string
		typ  ComponentType
		exts map[string]bool
	}{
		{opts.Config.Src.Modules, ComponentStandard, map[string]bool{".bas": true}},
		{opts.Config.Src.Classes, ComponentClass, map[string]bool{".cls": true}},
		{opts.Config.Src.Workbook, ComponentDocument, map[string]bool{".bas": true, ".cls": true}},
	} {
		items, collectErr := collectCode(root, source.dir, source.typ, source.exts, opts.RestrictToRoot)
		if collectErr != nil {
			return nil, collectErr
		}
		components = append(components, items...)
	}
	formComponents, err := collectForms(root, opts)
	if err != nil {
		return nil, err
	}
	components = append(components, formComponents...)
	slices.SortFunc(components, func(a, b Component) int {
		if cmp := strings.Compare(a.SourcePath, b.SourcePath); cmp != 0 {
			return cmp
		}
		return strings.Compare(string(a.Type), string(b.Type))
	})
	return components, nil
}

func collectCode(root, configured string, typ ComponentType, allowed map[string]bool, restrict bool) ([]Component, error) {
	base, err := resolveRoot(root, configured, restrict)
	if err != nil {
		return nil, fmt.Errorf("resolve %s source root: %w", typ, err)
	}
	info, err := os.Stat(base)
	if err != nil {
		return nil, fmt.Errorf("read %s source root %s: %w", typ, displayPath(root, base), err)
	}
	if !info.IsDir() {
		return nil, fmt.Errorf("%s source root %s is not a directory", typ, displayPath(root, base))
	}
	var out []Component
	err = filepath.WalkDir(base, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			return nil
		}
		ext := strings.ToLower(filepath.Ext(d.Name()))
		if !allowed[ext] {
			return fmt.Errorf("unsupported %s source file %s", typ, displayPath(root, path))
		}
		body, readErr := os.ReadFile(path)
		if readErr != nil {
			return fmt.Errorf("read source %s: %w", displayPath(root, path), readErr)
		}
		name := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		if !ValidComponentName(name) {
			return fmt.Errorf("invalid VBA component name %q in %s", name, displayPath(root, path))
		}
		out = append(out, Component{SourcePath: displayPath(root, path), AbsolutePath: path, Name: name, Type: typ, Source: body})
		return nil
	})
	return out, err
}

func collectForms(root string, opts Options) ([]Component, error) {
	base, err := resolveRoot(root, opts.Config.Src.Forms, opts.RestrictToRoot)
	if err != nil {
		return nil, fmt.Errorf("resolve form source root: %w", err)
	}
	if info, statErr := os.Stat(base); statErr != nil {
		return nil, fmt.Errorf("read form source root %s: %w", displayPath(root, base), statErr)
	} else if !info.IsDir() {
		return nil, fmt.Errorf("form source root %s is not a directory", displayPath(root, base))
	}

	sidecar := strings.EqualFold(opts.Config.UserForm.CodeSource, "sidecar")
	if sidecar {
		issues, validateErr := forms.ValidateUserFormCodeSidecars(base, nil)
		if validateErr != nil {
			return nil, validateErr
		}
		if len(issues) > 0 {
			return nil, fmt.Errorf("invalid UserForm sidecar: %s", issues[0].Error())
		}
	}
	if sidecar || opts.ValidateFormArtifacts {
		artifactIssues, validateErr := forms.ValidateUserFormArtifactsAgainstSpecs(base, nil)
		if validateErr != nil {
			return nil, validateErr
		}
		if len(artifactIssues) > 0 {
			return nil, fmt.Errorf("invalid UserForm artifact: %s", artifactIssues[0].Message)
		}
	}

	byLocation := map[string]*formFiles{}
	byName := map[string][]*formFiles{}
	codeDir := filepath.Join(base, "code")
	specDir := filepath.Join(base, "specs")
	err = filepath.WalkDir(base, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			if samePath(path, codeDir) || samePath(path, specDir) {
				return filepath.SkipDir
			}
			return nil
		}
		ext := strings.ToLower(filepath.Ext(d.Name()))
		if ext != ".frm" && ext != ".frx" {
			return fmt.Errorf("unsupported UserForm source file %s", displayPath(root, path))
		}
		if _, readErr := os.ReadFile(path); readErr != nil {
			return fmt.Errorf("read source %s: %w", displayPath(root, path), readErr)
		}
		name := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		if !ValidComponentName(name) {
			return fmt.Errorf("invalid VBA component name %q in %s", name, displayPath(root, path))
		}
		key := strings.ToLower(filepath.Clean(filepath.Dir(path))) + "\x00" + strings.ToLower(name)
		entry := byLocation[key]
		if entry == nil {
			entry = &formFiles{name: name}
			byLocation[key] = entry
			byName[strings.ToLower(name)] = append(byName[strings.ToLower(name)], entry)
		}
		switch ext {
		case ".frm":
			if entry.frm != "" {
				return fmt.Errorf("ambiguous UserForm source %q: %s and %s", name, displayPath(root, entry.frm), displayPath(root, path))
			}
			entry.frm = path
		case ".frx":
			if entry.frx != "" {
				return fmt.Errorf("ambiguous UserForm companion for %q", name)
			}
			entry.frx = path
		}
		return nil
	})
	if err != nil {
		return nil, err
	}

	eligible := make(map[string][]*formFiles, len(byName))
	for name, entries := range byName {
		for _, entry := range entries {
			if entry.frm == "" || opts.PrimaryExcluded == nil || !opts.PrimaryExcluded(displayPath(root, entry.frm)) {
				eligible[name] = append(eligible[name], entry)
			}
		}
	}
	if err := addFormArtifacts(root, base, eligible, byName, sidecar); err != nil {
		return nil, err
	}

	var out []Component
	for _, entries := range byName {
		for _, entry := range entries {
			if entry.frm == "" {
				return nil, fmt.Errorf("incomplete UserForm %q: .frx has no matching .frm", entry.name)
			}
			body, readErr := os.ReadFile(entry.frm)
			if readErr != nil {
				return nil, fmt.Errorf("read source %s: %w", displayPath(root, entry.frm), readErr)
			}
			paths := append([]string{}, entry.related...)
			if entry.frx != "" {
				paths = append(paths, entry.frx)
			}
			slices.Sort(paths)
			related := make([]Artifact, 0, len(paths))
			for _, path := range paths {
				artifactBody, artifactErr := os.ReadFile(path)
				if artifactErr != nil {
					return nil, fmt.Errorf("read source %s: %w", displayPath(root, path), artifactErr)
				}
				related = append(related, Artifact{Path: displayPath(root, path), AbsolutePath: path, Source: artifactBody})
			}
			out = append(out, Component{SourcePath: displayPath(root, entry.frm), AbsolutePath: entry.frm, Name: entry.name, Type: ComponentForm, Source: body, Related: related})
		}
	}
	return out, nil
}

func addFormArtifacts(root, formsDir string, byName, allForms map[string][]*formFiles, sidecar bool) error {
	locations := []struct {
		dir     string
		allowed map[string]bool
		unique  bool
	}{
		{filepath.Join(formsDir, "specs"), map[string]bool{".yaml": true, ".yml": true, ".json": true}, false},
	}
	if sidecar {
		locations = append(locations, struct {
			dir     string
			allowed map[string]bool
			unique  bool
		}{filepath.Join(formsDir, "code"), map[string]bool{".bas": true}, true})
	}
	for _, location := range locations {
		entries, err := os.ReadDir(location.dir)
		if errors.Is(err, os.ErrNotExist) {
			continue
		}
		if err != nil {
			return err
		}
		for _, entry := range entries {
			path := filepath.Join(location.dir, entry.Name())
			if entry.IsDir() {
				return fmt.Errorf("unsupported UserForm sidecar directory %s", displayPath(root, path))
			}
			ext := strings.ToLower(filepath.Ext(entry.Name()))
			if !location.allowed[ext] {
				return fmt.Errorf("unsupported UserForm sidecar file %s", displayPath(root, path))
			}
			name := strings.TrimSuffix(entry.Name(), filepath.Ext(entry.Name()))
			matches := byName[strings.ToLower(name)]
			if len(matches) == 0 {
				matches = allForms[strings.ToLower(name)]
				if len(matches) == 0 {
					return fmt.Errorf("orphan UserForm sidecar %s has no matching .frm", displayPath(root, path))
				}
			}
			if location.unique && len(matches) != 1 {
				return fmt.Errorf("ambiguous UserForm sidecar %s matches multiple .frm artifacts", displayPath(root, path))
			}
			for _, form := range matches {
				if form.frm == "" {
					return fmt.Errorf("orphan UserForm sidecar %s has no matching .frm", displayPath(root, path))
				}
				form.related = append(form.related, path)
			}
		}
	}
	return nil
}

func resolveRoot(root, configured string, restrict bool) (string, error) {
	if strings.TrimSpace(configured) == "" {
		return "", errors.New("path is required")
	}
	path := configured
	if !filepath.IsAbs(path) {
		path = filepath.Join(root, filepath.FromSlash(strings.ReplaceAll(path, "\\", "/")))
	}
	path = filepath.Clean(path)
	if restrict {
		rel, err := filepath.Rel(root, path)
		if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) || filepath.IsAbs(rel) {
			return "", fmt.Errorf("path %q must be inside the project root", configured)
		}
	}
	return path, nil
}

func ValidComponentName(name string) bool {
	runes := []rune(name)
	if len(runes) == 0 || len(runes) > 255 || !unicode.IsLetter(runes[0]) {
		return false
	}
	for _, r := range runes[1:] {
		if r != '_' && !unicode.IsLetter(r) && !unicode.IsDigit(r) {
			return false
		}
	}
	return true
}

func displayPath(root, path string) string {
	rel, err := filepath.Rel(root, path)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return filepath.ToSlash(path)
	}
	return filepath.ToSlash(rel)
}

func samePath(a, b string) bool { return strings.EqualFold(filepath.Clean(a), filepath.Clean(b)) }
