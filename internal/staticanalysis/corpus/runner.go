package corpus

import (
	"errors"
	"fmt"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
)

// NativeProject identifies a checked-in xlflow project under example/.
// Native projects are deliberately separate from the third-party manifest's
// Project type, whose Path and provenance fields have a different contract.
type NativeProject struct {
	ID   string
	Root string
}

type Surface string

const (
	SurfaceConfig  Surface = "config"
	SurfaceLint    Surface = "lint"
	SurfaceAnalyze Surface = "analyze"
)

type FailureKind string

const (
	FailureInvalidProjectConfig FailureKind = "invalid_project_config"
	FailureParser               FailureKind = "parser_failure"
	FailureExecution            FailureKind = "execution_failure"
	FailureWorkspace            FailureKind = "workspace_failure"
	FailureUnsupportedLayout    FailureKind = "unsupported_project_layout"
)

type Diagnostic struct {
	Project  string  `json:"project"`
	Surface  Surface `json:"surface"`
	Code     string  `json:"code"`
	Severity string  `json:"severity"`
	File     string  `json:"file"`
	Line     int     `json:"line"`
	Column   int     `json:"column,omitempty"`
}

type Failure struct {
	Project string      `json:"project"`
	Surface Surface     `json:"surface"`
	Kind    FailureKind `json:"kind"`
	Err     error       `json:"-"`
}

type Skip struct {
	Project string `json:"project"`
	Reason  string `json:"reason"`
}

type SurfaceResult struct {
	Project     string       `json:"project"`
	Surface     Surface      `json:"surface"`
	Diagnostics []Diagnostic `json:"diagnostics"`
	Failure     *Failure     `json:"failure,omitempty"`
}

type Report struct {
	Surfaces    []SurfaceResult `json:"surfaces"`
	Diagnostics []Diagnostic    `json:"diagnostics"`
	Failures    []Failure       `json:"failures"`
	Skipped     []Skip          `json:"skipped,omitempty"`
}

// RunError reports project-level failures while preserving all successful and
// partial results in the accompanying Report.
type RunError struct {
	Failures []Failure
}

func (e *RunError) Error() string {
	if e == nil {
		return "static-analysis corpus execution failed"
	}
	return fmt.Sprintf("static-analysis corpus execution failed for %d project surface(s)", len(e.Failures))
}

// DiscoverExampleProjects finds configured native projects without recursing
// into their source trees. Results are identified and ordered deterministically.
func DiscoverExampleProjects(repoRoot string) ([]NativeProject, error) {
	root, err := filepath.Abs(repoRoot)
	if err != nil {
		return nil, fmt.Errorf("resolve repository root: %w", err)
	}
	examplesRoot := filepath.Join(root, "example")
	entries, err := os.ReadDir(examplesRoot)
	if err != nil {
		return nil, fmt.Errorf("read example directory: %w", err)
	}
	projects := make([]NativeProject, 0)
	for _, entry := range entries {
		if !entry.IsDir() || entry.Type()&os.ModeSymlink != 0 {
			continue
		}
		projectRoot := filepath.Join(examplesRoot, entry.Name())
		configPath := filepath.Join(projectRoot, config.FileName)
		if _, err := os.Stat(configPath); err != nil {
			if os.IsNotExist(err) {
				continue
			}
			return nil, fmt.Errorf("inspect project %s: %w", entry.Name(), err)
		}
		projects = append(projects, NativeProject{
			ID:   "self/" + entry.Name(),
			Root: projectRoot,
		})
	}
	sort.Slice(projects, func(i, j int) bool { return projects[i].ID < projects[j].ID })
	return projects, nil
}

// SelectProjects filters projects by ID while retaining discovery order.
// Empty ids selects every project.
func SelectProjects(projects []NativeProject, ids []string) ([]NativeProject, error) {
	if len(ids) == 0 {
		return append([]NativeProject(nil), projects...), nil
	}
	known := make(map[string]NativeProject, len(projects))
	for _, project := range projects {
		if _, exists := known[project.ID]; exists {
			return nil, fmt.Errorf("duplicate project ID %q", project.ID)
		}
		known[project.ID] = project
	}
	wanted := make(map[string]struct{}, len(ids))
	for _, id := range ids {
		if _, exists := wanted[id]; exists {
			return nil, fmt.Errorf("duplicate selected project ID %q", id)
		}
		if _, exists := known[id]; !exists {
			return nil, fmt.Errorf("unknown project ID %q", id)
		}
		wanted[id] = struct{}{}
	}
	selected := make([]NativeProject, 0, len(ids))
	for _, project := range projects {
		if _, ok := wanted[project.ID]; ok {
			selected = append(selected, project)
		}
	}
	return selected, nil
}

type executor struct {
	loadConfig func(string) (config.Config, error)
	runLint    func(string, config.Config) (lint.Result, error)
	runAnalyze func(string, config.Config) (analyze.Result, error)
}

func productionExecutor() executor {
	return executor{
		loadConfig: config.Load,
		runLint: func(root string, cfg config.Config) (lint.Result, error) {
			return (lint.Linter{RootDir: root, Config: cfg}).RunResult()
		},
		runAnalyze: func(root string, cfg config.Config) (analyze.Result, error) {
			return (analyze.Analyzer{RootDir: root, Config: cfg}).RunResult()
		},
	}
}

// RunProjects executes lint and analyze independently for each project. A
// diagnostic is data, not a failure; only configuration or execution errors
// populate the returned RunError.
func RunProjects(projects []NativeProject) (Report, error) {
	return runProjects(projects, productionExecutor())
}

// SelectThirdPartyProjects filters manifest projects by their stable manifest
// IDs. Empty ids selects every project in manifest order, including disabled
// entries so callers can report their documented skip reasons.
func SelectThirdPartyProjects(projects []Project, ids []string) ([]Project, error) {
	if len(ids) == 0 {
		return append([]Project(nil), projects...), nil
	}
	known := make(map[string]Project, len(projects))
	for _, project := range projects {
		if _, exists := known[project.ID]; exists {
			return nil, fmt.Errorf("duplicate project ID %q", project.ID)
		}
		known[project.ID] = project
	}
	wanted := make(map[string]struct{}, len(ids))
	for _, id := range ids {
		if _, exists := wanted[id]; exists {
			return nil, fmt.Errorf("duplicate selected project ID %q", id)
		}
		if _, exists := known[id]; !exists {
			return nil, fmt.Errorf("unknown project ID %q", id)
		}
		wanted[id] = struct{}{}
	}
	selected := make([]Project, 0, len(ids))
	for _, project := range projects {
		if _, ok := wanted[project.ID]; ok {
			selected = append(selected, project)
		}
	}
	return selected, nil
}

// RunThirdPartyProjects materializes and analyzes each vendored project in an
// isolated workspace. The caller supplies the manifest directory as
// corpusRoot, not the project destination itself.
func RunThirdPartyProjects(corpusRoot string, projects []Project, opts MaterializeOptions) (Report, error) {
	return runThirdPartyProjects(corpusRoot, projects, opts, productionExecutor())
}

func runThirdPartyProjects(corpusRoot string, projects []Project, opts MaterializeOptions, exec executor) (Report, error) {
	report := Report{
		Surfaces:    make([]SurfaceResult, 0, len(projects)*2),
		Diagnostics: make([]Diagnostic, 0),
		Failures:    make([]Failure, 0),
		Skipped:     make([]Skip, 0),
	}
	for _, project := range projects {
		projectID := "third_party/" + project.ID
		if !project.Enabled {
			report.Skipped = append(report.Skipped, Skip{Project: projectID, Reason: project.Notes})
			continue
		}
		workspace, err := MaterializeThirdPartyProject(corpusRoot, project, opts)
		if err != nil {
			failure := Failure{Project: projectID, Surface: SurfaceConfig, Kind: classifyThirdPartyFailure(err), Err: err}
			report.Failures = append(report.Failures, failure)
			continue
		}
		cfg, err := exec.loadConfig(workspace.Root)
		if err != nil {
			failure := Failure{Project: projectID, Surface: SurfaceConfig, Kind: FailureInvalidProjectConfig, Err: err}
			report.Failures = append(report.Failures, failure)
			if closeErr := workspace.Close(); closeErr != nil {
				closeFailure := Failure{Project: projectID, Surface: SurfaceConfig, Kind: FailureWorkspace, Err: closeErr}
				report.Failures = append(report.Failures, closeFailure)
			}
			continue
		}

		lintSurface := SurfaceResult{Project: projectID, Surface: SurfaceLint, Diagnostics: make([]Diagnostic, 0)}
		lintResult, lintErr := exec.runLint(workspace.Root, cfg)
		if lintErr != nil {
			failure := Failure{Project: projectID, Surface: SurfaceLint, Kind: classifyFailure(lintErr), Err: lintErr}
			lintSurface.Failure = &failure
			report.Failures = append(report.Failures, failure)
		} else if diagnostics, normalizeErr := normalizeThirdPartyLintDiagnostics(projectID, project.Profile, workspace.Mappings, lintResult.Issues); normalizeErr != nil {
			failure := Failure{Project: projectID, Surface: SurfaceLint, Kind: FailureWorkspace, Err: normalizeErr}
			lintSurface.Failure = &failure
			report.Failures = append(report.Failures, failure)
		} else {
			lintSurface.Diagnostics = diagnostics
			report.Diagnostics = append(report.Diagnostics, diagnostics...)
		}
		report.Surfaces = append(report.Surfaces, lintSurface)

		analyzeSurface := SurfaceResult{Project: projectID, Surface: SurfaceAnalyze, Diagnostics: make([]Diagnostic, 0)}
		analyzeResult, analyzeErr := exec.runAnalyze(workspace.Root, cfg)
		if analyzeErr != nil {
			failure := Failure{Project: projectID, Surface: SurfaceAnalyze, Kind: classifyFailure(analyzeErr), Err: analyzeErr}
			analyzeSurface.Failure = &failure
			report.Failures = append(report.Failures, failure)
		} else if diagnostics, normalizeErr := normalizeThirdPartyAnalyzeDiagnostics(projectID, project.Profile, workspace.Mappings, analyzeResult.Findings); normalizeErr != nil {
			failure := Failure{Project: projectID, Surface: SurfaceAnalyze, Kind: FailureWorkspace, Err: normalizeErr}
			analyzeSurface.Failure = &failure
			report.Failures = append(report.Failures, failure)
		} else {
			analyzeSurface.Diagnostics = diagnostics
			report.Diagnostics = append(report.Diagnostics, diagnostics...)
		}
		report.Surfaces = append(report.Surfaces, analyzeSurface)
		if closeErr := workspace.Close(); closeErr != nil {
			failure := Failure{Project: projectID, Surface: SurfaceConfig, Kind: FailureWorkspace, Err: closeErr}
			report.Failures = append(report.Failures, failure)
		}
	}
	sortDiagnostics(report.Diagnostics)
	sortThirdPartyFailuresAndSkips(&report)
	if len(report.Failures) > 0 {
		return report, &RunError{Failures: append([]Failure(nil), report.Failures...)}
	}
	return report, nil
}

func runProjects(projects []NativeProject, exec executor) (Report, error) {
	report := Report{
		Surfaces:    make([]SurfaceResult, 0, len(projects)*2),
		Diagnostics: make([]Diagnostic, 0),
		Failures:    make([]Failure, 0),
	}
	for _, project := range projects {
		cfg, err := exec.loadConfig(project.Root)
		if err != nil {
			report.Failures = append(report.Failures, Failure{
				Project: project.ID, Surface: SurfaceConfig, Kind: FailureInvalidProjectConfig, Err: err,
			})
			continue
		}

		lintResult, err := exec.runLint(project.Root, cfg)
		lintSurface := SurfaceResult{Project: project.ID, Surface: SurfaceLint, Diagnostics: make([]Diagnostic, 0)}
		if err != nil {
			failure := Failure{Project: project.ID, Surface: SurfaceLint, Kind: classifyFailure(err), Err: err}
			lintSurface.Failure = &failure
			report.Failures = append(report.Failures, failure)
		} else {
			lintSurface.Diagnostics = normalizeLintDiagnostics(project.ID, lintResult.Issues)
			report.Diagnostics = append(report.Diagnostics, lintSurface.Diagnostics...)
		}
		report.Surfaces = append(report.Surfaces, lintSurface)

		analyzeResult, err := exec.runAnalyze(project.Root, cfg)
		analyzeSurface := SurfaceResult{Project: project.ID, Surface: SurfaceAnalyze, Diagnostics: make([]Diagnostic, 0)}
		if err != nil {
			failure := Failure{Project: project.ID, Surface: SurfaceAnalyze, Kind: classifyFailure(err), Err: err}
			analyzeSurface.Failure = &failure
			report.Failures = append(report.Failures, failure)
		} else {
			analyzeSurface.Diagnostics = normalizeAnalyzeDiagnostics(project.ID, analyzeResult.Findings)
			report.Diagnostics = append(report.Diagnostics, analyzeSurface.Diagnostics...)
		}
		report.Surfaces = append(report.Surfaces, analyzeSurface)
	}
	sortDiagnostics(report.Diagnostics)
	sort.SliceStable(report.Failures, func(i, j int) bool {
		a, b := report.Failures[i], report.Failures[j]
		if a.Project != b.Project {
			return a.Project < b.Project
		}
		if a.Surface != b.Surface {
			return surfaceRank(a.Surface) < surfaceRank(b.Surface)
		}
		return a.Kind < b.Kind
	})
	if len(report.Failures) > 0 {
		return report, &RunError{Failures: append([]Failure(nil), report.Failures...)}
	}
	return report, nil
}

func classifyFailure(err error) FailureKind {
	var parseErr *analyze.ParseError
	if errors.As(err, &parseErr) {
		return FailureParser
	}
	return FailureExecution
}

func classifyThirdPartyFailure(err error) FailureKind {
	if errors.Is(err, ErrUnsupportedProjectLayout) {
		return FailureUnsupportedLayout
	}
	return FailureWorkspace
}

func normalizeLintDiagnostics(project string, issues []lint.Issue) []Diagnostic {
	result := make([]Diagnostic, 0, len(issues))
	for _, issue := range issues {
		result = append(result, Diagnostic{Project: project, Surface: SurfaceLint, Code: issue.Code, Severity: issue.Severity, File: filepath.ToSlash(filepath.Clean(issue.File)), Line: issue.Line, Column: issue.Column})
	}
	return result
}

func normalizeAnalyzeDiagnostics(project string, findings []analyze.Finding) []Diagnostic {
	result := make([]Diagnostic, 0, len(findings))
	for _, finding := range findings {
		result = append(result, Diagnostic{Project: project, Surface: SurfaceAnalyze, Code: finding.Code, Severity: finding.Severity, File: filepath.ToSlash(filepath.Clean(finding.File)), Line: finding.Line, Column: finding.Column})
	}
	return result
}

func normalizeThirdPartyLintDiagnostics(project, profile string, mappings map[string]string, issues []lint.Issue) ([]Diagnostic, error) {
	result := make([]Diagnostic, 0, len(issues))
	for _, issue := range issues {
		if profileExcludes(profile, issue.Code) {
			continue
		}
		file, err := remapThirdPartyPath(mappings, issue.File)
		if err != nil {
			return nil, err
		}
		result = append(result, Diagnostic{Project: project, Surface: SurfaceLint, Code: issue.Code, Severity: issue.Severity, File: file, Line: issue.Line, Column: issue.Column})
	}
	return result, nil
}

func normalizeThirdPartyAnalyzeDiagnostics(project, profile string, mappings map[string]string, findings []analyze.Finding) ([]Diagnostic, error) {
	result := make([]Diagnostic, 0, len(findings))
	for _, finding := range findings {
		if profileExcludes(profile, finding.Code) {
			continue
		}
		file, err := remapThirdPartyPath(mappings, finding.File)
		if err != nil {
			return nil, err
		}
		result = append(result, Diagnostic{Project: project, Surface: SurfaceAnalyze, Code: finding.Code, Severity: finding.Severity, File: file, Line: finding.Line, Column: finding.Column})
	}
	return result, nil
}

func remapThirdPartyPath(mappings map[string]string, file string) (string, error) {
	key := path.Clean(filepath.ToSlash(strings.TrimSpace(file)))
	if key == "." || key == "" {
		return "", fmt.Errorf("third-party diagnostic has empty source path")
	}
	if source, ok := mappings[key]; ok {
		return source, nil
	}
	return "", fmt.Errorf("third-party diagnostic path %q is not in the materialized source map", file)
}

func sortThirdPartyFailuresAndSkips(report *Report) {
	sort.SliceStable(report.Failures, func(i, j int) bool {
		a, b := report.Failures[i], report.Failures[j]
		if a.Project != b.Project {
			return a.Project < b.Project
		}
		if a.Surface != b.Surface {
			return surfaceRank(a.Surface) < surfaceRank(b.Surface)
		}
		return a.Kind < b.Kind
	})
	sort.SliceStable(report.Skipped, func(i, j int) bool {
		return report.Skipped[i].Project < report.Skipped[j].Project
	})
}

func sortDiagnostics(diagnostics []Diagnostic) {
	sort.SliceStable(diagnostics, func(i, j int) bool {
		a, b := diagnostics[i], diagnostics[j]
		if a.Project != b.Project {
			return a.Project < b.Project
		}
		if a.Surface != b.Surface {
			return surfaceRank(a.Surface) < surfaceRank(b.Surface)
		}
		if a.File != b.File {
			return a.File < b.File
		}
		if a.Line != b.Line {
			return a.Line < b.Line
		}
		if a.Column != b.Column {
			return a.Column < b.Column
		}
		if a.Code != b.Code {
			return a.Code < b.Code
		}
		return a.Severity < b.Severity
	})
}

func surfaceRank(surface Surface) int {
	switch surface {
	case SurfaceConfig:
		return 0
	case SurfaceLint:
		return 1
	case SurfaceAnalyze:
		return 2
	default:
		return 3
	}
}
