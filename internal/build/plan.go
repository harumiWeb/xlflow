// Package build resolves the read-only source plan used by the Excel-backed
// release build command. It deliberately has no dependency on CLI or Excel.
package build

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"unicode"

	"github.com/bmatcuk/doublestar/v4"
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

// BuildComponent is one VBA project component and every tracked UserForm
// artifact that belongs to it. Paths are normalized relative to the project
// root and use forward slashes.
type BuildComponent struct {
	SourcePath   string        `json:"source_path"`
	Name         string        `json:"name"`
	Type         ComponentType `json:"type"`
	Reason       string        `json:"reason"`
	RelatedPaths []string      `json:"related_paths,omitempty"`
}

type BuildWarning struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Pattern string `json:"pattern,omitempty"`
}

// BuildPlan is deterministic and entirely read-only. BaseWorkbook and
// OutputPath are normalized project-root-relative paths.
type BuildPlan struct {
	BaseWorkbook string           `json:"base_workbook"`
	OutputPath   string           `json:"output_path"`
	Included     []BuildComponent `json:"included"`
	Excluded     []BuildComponent `json:"excluded"`
	Warnings     []BuildWarning   `json:"warnings,omitempty"`
}

type Options struct {
	Root         string
	Config       config.Config
	BaseWorkbook string
	OutputPath   string
}

type formFiles struct {
	name    string
	frm     string
	frx     string
	related []string
}

// Plan resolves source inputs without opening Excel or modifying any source
// or workbook. Empty BaseWorkbook uses excel.path; empty OutputPath uses
// build/Release/<base filename>.
func Plan(opts Options) (BuildPlan, error) {
	root, err := filepath.Abs(opts.Root)
	if err != nil {
		return BuildPlan{}, fmt.Errorf("resolve project root: %w", err)
	}
	root = filepath.Clean(root)
	if opts.BaseWorkbook == "" {
		opts.BaseWorkbook = opts.Config.Excel.Path
	}
	base, err := projectPath(root, opts.BaseWorkbook)
	if err != nil {
		return BuildPlan{}, fmt.Errorf("resolve build base: %w", err)
	}
	if opts.OutputPath == "" {
		opts.OutputPath = filepath.Join("build", "Release", filepath.Base(base.absolute))
	}
	output, err := projectPath(root, opts.OutputPath)
	if err != nil {
		return BuildPlan{}, fmt.Errorf("resolve build output: %w", err)
	}
	if sameFileIdentity(base.absolute, output.absolute) {
		return BuildPlan{}, errors.New("build base and output must refer to different files")
	}

	patterns, err := normalizePatterns(opts.Config.Build.Exclude)
	if err != nil {
		return BuildPlan{}, err
	}
	components, err := collectComponents(root, opts.Config)
	if err != nil {
		return BuildPlan{}, err
	}

	matched := make(map[string]bool, len(patterns))
	plan := BuildPlan{BaseWorkbook: base.relative, OutputPath: output.relative}
	for _, component := range components {
		componentPatterns := matchingPatterns(component, patterns)
		if len(componentPatterns) == 0 {
			component.Reason = "included"
			plan.Included = append(plan.Included, component)
			continue
		}
		for _, pattern := range componentPatterns {
			matched[pattern] = true
		}
		component.Reason = "excluded by " + componentPatterns[0]
		plan.Excluded = append(plan.Excluded, component)
	}
	if err := validateIncludedNames(plan.Included); err != nil {
		return BuildPlan{}, err
	}
	for _, pattern := range patterns {
		if !matched[pattern] {
			plan.Warnings = append(plan.Warnings, BuildWarning{
				Code: "build_exclude_unmatched", Pattern: pattern,
				Message: fmt.Sprintf("build exclusion pattern %q did not match any source component", pattern),
			})
		}
	}
	return plan, nil
}

type resolvedPath struct{ absolute, relative string }

func projectPath(root, value string) (resolvedPath, error) {
	if strings.TrimSpace(value) == "" {
		return resolvedPath{}, errors.New("path is required")
	}
	path := value
	if !filepath.IsAbs(path) {
		path = filepath.Join(root, filepath.FromSlash(strings.ReplaceAll(path, "\\", "/")))
	}
	path = filepath.Clean(path)
	rel, err := filepath.Rel(root, path)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) || filepath.IsAbs(rel) {
		return resolvedPath{}, fmt.Errorf("path %q must be inside the project root", value)
	}
	return resolvedPath{absolute: path, relative: filepath.ToSlash(rel)}, nil
}

func sameFileIdentity(a, b string) bool {
	a = resolveExistingAncestor(a)
	b = resolveExistingAncestor(b)
	return strings.EqualFold(filepath.Clean(a), filepath.Clean(b))
}

func resolveExistingAncestor(path string) string {
	path = filepath.Clean(path)
	var suffix []string
	for {
		resolved, err := filepath.EvalSymlinks(path)
		if err == nil {
			for i := len(suffix) - 1; i >= 0; i-- {
				resolved = filepath.Join(resolved, suffix[i])
			}
			return resolved
		}
		parent := filepath.Dir(path)
		if parent == path {
			return path
		}
		suffix = append(suffix, filepath.Base(path))
		path = parent
	}
}

func normalizePatterns(raw []string) ([]string, error) {
	seen := map[string]bool{}
	patterns := make([]string, 0, len(raw))
	for _, value := range raw {
		pattern := filepath.ToSlash(strings.TrimSpace(strings.ReplaceAll(value, "\\", "/")))
		if pattern == "" {
			return nil, errors.New("build.exclude must not contain an empty pattern")
		}
		if strings.HasPrefix(pattern, "/") || isDriveAbsolute(pattern) || pattern == ".." || strings.HasPrefix(pattern, "../") || strings.Contains(pattern, "/../") {
			return nil, fmt.Errorf("build exclusion pattern %q must be project-root-relative", value)
		}
		if !doublestar.ValidatePattern(pattern) {
			return nil, fmt.Errorf("invalid build exclusion pattern %q", value)
		}
		if !seen[pattern] {
			seen[pattern] = true
			patterns = append(patterns, pattern)
		}
	}
	sort.Strings(patterns)
	return patterns, nil
}

func isDriveAbsolute(path string) bool {
	return len(path) >= 2 && path[1] == ':' && ((path[0] >= 'a' && path[0] <= 'z') || (path[0] >= 'A' && path[0] <= 'Z'))
}

func collectComponents(root string, cfg config.Config) ([]BuildComponent, error) {
	components := make([]BuildComponent, 0)
	for _, source := range []struct {
		dir  string
		typ  ComponentType
		exts map[string]bool
	}{
		{cfg.Src.Modules, ComponentStandard, map[string]bool{".bas": true}},
		{cfg.Src.Classes, ComponentClass, map[string]bool{".cls": true}},
		{cfg.Src.Workbook, ComponentDocument, map[string]bool{".bas": true, ".cls": true}},
	} {
		items, err := collectCodeComponents(root, source.dir, source.typ, source.exts)
		if err != nil {
			return nil, err
		}
		components = append(components, items...)
	}
	formComponents, err := collectFormComponents(root, cfg)
	if err != nil {
		return nil, err
	}
	components = append(components, formComponents...)
	sort.Slice(components, func(i, j int) bool {
		if components[i].SourcePath != components[j].SourcePath {
			return components[i].SourcePath < components[j].SourcePath
		}
		return components[i].Type < components[j].Type
	})
	return components, nil
}

func collectCodeComponents(root, configured string, typ ComponentType, allowed map[string]bool) ([]BuildComponent, error) {
	base, err := projectPath(root, configured)
	if err != nil {
		return nil, fmt.Errorf("resolve %s source root: %w", typ, err)
	}
	info, err := os.Stat(base.absolute)
	if err != nil {
		return nil, fmt.Errorf("read %s source root %s: %w", typ, base.relative, err)
	}
	if !info.IsDir() {
		return nil, fmt.Errorf("%s source root %s is not a directory", typ, base.relative)
	}
	var out []BuildComponent
	err = filepath.WalkDir(base.absolute, func(path string, d os.DirEntry, walkErr error) error {
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
		if _, err := os.ReadFile(path); err != nil {
			return fmt.Errorf("read source %s: %w", displayPath(root, path), err)
		}
		name := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		if !validComponentName(name) {
			return fmt.Errorf("invalid VBA component name %q in %s", name, displayPath(root, path))
		}
		out = append(out, BuildComponent{SourcePath: displayPath(root, path), Name: name, Type: typ})
		return nil
	})
	if err != nil {
		return nil, err
	}
	return out, nil
}

func collectFormComponents(root string, cfg config.Config) ([]BuildComponent, error) {
	base, err := projectPath(root, cfg.Src.Forms)
	if err != nil {
		return nil, fmt.Errorf("resolve form source root: %w", err)
	}
	if info, statErr := os.Stat(base.absolute); statErr != nil {
		return nil, fmt.Errorf("read form source root %s: %w", base.relative, statErr)
	} else if !info.IsDir() {
		return nil, fmt.Errorf("form source root %s is not a directory", base.relative)
	}

	if cfg.UserForm.CodeSource == "sidecar" {
		issues, validateErr := forms.ValidateUserFormCodeSidecars(base.absolute, nil)
		if validateErr != nil {
			return nil, validateErr
		}
		if len(issues) > 0 {
			return nil, fmt.Errorf("invalid UserForm sidecar: %s", issues[0].Error())
		}
		artifactIssues, validateErr := forms.ValidateUserFormArtifactsAgainstSpecs(base.absolute, nil)
		if validateErr != nil {
			return nil, validateErr
		}
		if len(artifactIssues) > 0 {
			return nil, fmt.Errorf("invalid UserForm artifact: %s", artifactIssues[0].Message)
		}
	}

	byLocation := map[string]*formFiles{}
	byName := map[string][]*formFiles{}
	codeDir := filepath.Join(base.absolute, "code")
	specDir := filepath.Join(base.absolute, "specs")
	err = filepath.WalkDir(base.absolute, func(path string, d os.DirEntry, walkErr error) error {
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
		if _, err := os.ReadFile(path); err != nil {
			return fmt.Errorf("read source %s: %w", displayPath(root, path), err)
		}
		name := strings.TrimSuffix(d.Name(), filepath.Ext(d.Name()))
		if !validComponentName(name) {
			return fmt.Errorf("invalid VBA component name %q in %s", name, displayPath(root, path))
		}
		key := formLocationKey(filepath.Dir(path), name)
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
	if err := addFormArtifacts(root, base.absolute, byName, cfg.UserForm.CodeSource == "sidecar"); err != nil {
		return nil, err
	}
	var out []BuildComponent
	for _, entries := range byName {
		for _, entry := range entries {
			if entry.frm == "" {
				return nil, fmt.Errorf("incomplete UserForm %q: .frx has no matching .frm", entry.name)
			}
			name := strings.TrimSuffix(filepath.Base(entry.frm), filepath.Ext(entry.frm))
			related := append([]string{}, entry.related...)
			if entry.frx != "" {
				related = append(related, displayPath(root, entry.frx))
			}
			sort.Strings(related)
			out = append(out, BuildComponent{SourcePath: displayPath(root, entry.frm), Name: name, Type: ComponentForm, RelatedPaths: related})
		}
	}
	return out, nil
}

func addFormArtifacts(root, formsDir string, byName map[string][]*formFiles, sidecar bool) error {
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
			if entry.IsDir() {
				return fmt.Errorf("unsupported UserForm sidecar directory %s", displayPath(root, filepath.Join(location.dir, entry.Name())))
			}
			ext := strings.ToLower(filepath.Ext(entry.Name()))
			if !location.allowed[ext] {
				return fmt.Errorf("unsupported UserForm sidecar file %s", displayPath(root, filepath.Join(location.dir, entry.Name())))
			}
			path := filepath.Join(location.dir, entry.Name())
			if _, err := os.ReadFile(path); err != nil {
				return fmt.Errorf("read source %s: %w", displayPath(root, path), err)
			}
			name := strings.TrimSuffix(entry.Name(), filepath.Ext(entry.Name()))
			forms := byName[strings.ToLower(name)]
			if len(forms) == 0 {
				return fmt.Errorf("orphan UserForm sidecar %s has no matching .frm", displayPath(root, path))
			}
			if location.unique && len(forms) != 1 {
				return fmt.Errorf("ambiguous UserForm sidecar %s matches multiple .frm artifacts", displayPath(root, path))
			}
			for _, form := range forms {
				if form.frm == "" {
					return fmt.Errorf("orphan UserForm sidecar %s has no matching .frm", displayPath(root, path))
				}
				form.related = append(form.related, displayPath(root, path))
			}
		}
	}
	return nil
}

func formLocationKey(dir, name string) string {
	return strings.ToLower(filepath.Clean(dir)) + "\x00" + strings.ToLower(name)
}

func matchingPatterns(component BuildComponent, patterns []string) []string {
	paths := append([]string{component.SourcePath}, component.RelatedPaths...)
	var matches []string
	for _, pattern := range patterns {
		for _, path := range paths {
			matched, err := doublestar.Match(pattern, path)
			if err == nil && matched {
				matches = append(matches, pattern)
				break
			}
		}
	}
	return matches
}

func validateIncludedNames(components []BuildComponent) error {
	seen := map[string]BuildComponent{}
	for _, component := range components {
		key := strings.ToLower(component.Name)
		if previous, ok := seen[key]; ok {
			return fmt.Errorf("duplicate included VBA component name %q: %s and %s", component.Name, previous.SourcePath, component.SourcePath)
		}
		seen[key] = component
	}
	return nil
}

func validComponentName(name string) bool {
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
	if err != nil {
		return filepath.ToSlash(path)
	}
	return filepath.ToSlash(rel)
}

func samePath(a, b string) bool { return strings.EqualFold(filepath.Clean(a), filepath.Clean(b)) }
