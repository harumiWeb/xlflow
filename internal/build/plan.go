// Package build resolves the read-only source plan used by the Excel-backed
// release build command. It deliberately has no dependency on CLI or Excel.
package build

import (
	"errors"
	"fmt"
	"path/filepath"
	"sort"
	"strings"

	"github.com/bmatcuk/doublestar/v4"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/sourceinventory"
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
	components, err := collectComponents(root, opts.Config, patterns)
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

func collectComponents(root string, cfg config.Config, patterns []string) ([]BuildComponent, error) {
	discovered, err := sourceinventory.Discover(sourceinventory.Options{
		Root:           root,
		Config:         cfg,
		RestrictToRoot: true,
		PrimaryExcluded: func(path string) bool {
			return len(matchingPatterns(BuildComponent{SourcePath: path}, patterns)) > 0
		},
	})
	if err != nil {
		return nil, err
	}
	components := make([]BuildComponent, 0, len(discovered))
	for _, component := range discovered {
		components = append(components, BuildComponent{
			SourcePath: component.SourcePath, Name: component.Name,
			Type: ComponentType(component.Type), RelatedPaths: component.RelatedPaths(),
		})
	}
	return components, nil
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
