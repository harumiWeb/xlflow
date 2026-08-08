package corpus

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

// SourceInventory is the observed source-file inventory for one project. It
// aliases the manifest's SourceCounts so one canonical count contract is used
// by synchronization, inventory, and callers.
type SourceInventory = SourceCounts

func addSourceCounts(a, b SourceCounts) SourceCounts {
	return SourceCounts{Bas: a.Bas + b.Bas, Cls: a.Cls + b.Cls, Frm: a.Frm + b.Frm}
}

func equalSourceCounts(a, b SourceCounts) bool {
	return a.Bas == b.Bas && a.Cls == b.Cls && a.Frm == b.Frm
}

// ProjectInventory records one manifest entry and its observed source tree.
// Expected is copied from the manifest and Actual is derived from disk.
type ProjectInventory struct {
	ID       string          `json:"id"`
	Profile  string          `json:"profile"`
	Enabled  bool            `json:"enabled"`
	Path     string          `json:"path"`
	Expected SourceInventory `json:"expected"`
	Actual   SourceInventory `json:"actual"`
}

// InventoryReport is a deterministic summary of the manifest-backed source
// tree. Project rows and profile rows are sorted lexically by their stable
// IDs/names. The aggregate counts include all manifest entries, including
// explicitly disabled projects.
type InventoryReport struct {
	Projects      []ProjectInventory `json:"projects"`
	ProjectCount  int                `json:"project_count"`
	EnabledCount  int                `json:"enabled_count"`
	DisabledCount int                `json:"disabled_count"`
	Profiles      []ProfileInventory `json:"profiles"`
	Sources       SourceInventory    `json:"sources"`
	Missing       []string           `json:"missing,omitempty"`
	Unexpected    []string           `json:"unexpected,omitempty"`
}

// ProfileInventory is a deterministic profile count in an InventoryReport.
type ProfileInventory struct {
	Profile string `json:"profile"`
	Count   int    `json:"count"`
}

// InventoryMismatch identifies a project whose observed source counts differ
// from its manifest SourceCounts entry.
type InventoryMismatch struct {
	Project  string          `json:"project"`
	Expected SourceInventory `json:"expected"`
	Actual   SourceInventory `json:"actual"`
}

// InventoryValidationError reports all inventory failures in stable project
// order rather than stopping at the first mismatch.
type InventoryValidationError struct {
	Missing    []string            `json:"missing,omitempty"`
	Unexpected []string            `json:"unexpected,omitempty"`
	Mismatches []InventoryMismatch `json:"mismatches,omitempty"`
}

func (e *InventoryValidationError) Error() string {
	if e == nil {
		return "static-analysis corpus inventory validation failed"
	}
	parts := make([]string, 0, 3)
	if len(e.Missing) > 0 {
		parts = append(parts, fmt.Sprintf("missing projects: %s", strings.Join(e.Missing, ", ")))
	}
	if len(e.Unexpected) > 0 {
		parts = append(parts, fmt.Sprintf("unexpected projects: %s", strings.Join(e.Unexpected, ", ")))
	}
	if len(e.Mismatches) > 0 {
		items := make([]string, 0, len(e.Mismatches))
		for _, mismatch := range e.Mismatches {
			items = append(items, fmt.Sprintf("%s expected bas=%d cls=%d frm=%d, got bas=%d cls=%d frm=%d", mismatch.Project, mismatch.Expected.Bas, mismatch.Expected.Cls, mismatch.Expected.Frm, mismatch.Actual.Bas, mismatch.Actual.Cls, mismatch.Actual.Frm))
		}
		parts = append(parts, "source_counts mismatches: "+strings.Join(items, "; "))
	}
	if len(parts) == 0 {
		return "static-analysis corpus inventory validation failed"
	}
	return "static-analysis corpus inventory validation failed: " + strings.Join(parts, "; ")
}

// BuildInventory scans the manifest destination tree below corpusRoot and
// returns deterministic project/profile/source counts. It does not compare
// observed counts with the manifest; use ValidateInventory for that check.
func BuildInventory(corpusRoot string, manifest Manifest) (InventoryReport, error) {
	root, err := filepath.Abs(corpusRoot)
	if err != nil {
		return InventoryReport{}, fmt.Errorf("resolve corpus root: %w", err)
	}
	if info, statErr := os.Stat(root); statErr != nil {
		return InventoryReport{}, fmt.Errorf("stat corpus root: %w", statErr)
	} else if !info.IsDir() {
		return InventoryReport{}, fmt.Errorf("corpus root %q is not a directory", root)
	}

	report := InventoryReport{
		Projects: make([]ProjectInventory, 0, len(manifest.Projects)),
		Profiles: make([]ProfileInventory, 0),
	}
	profileCounts := make(map[string]int)
	for _, project := range manifest.Projects {
		observed, scanErr := scanInventoryProject(root, project)
		if scanErr != nil {
			var missing *InventoryValidationError
			if !errors.As(scanErr, &missing) || len(missing.Missing) == 0 {
				return InventoryReport{}, scanErr
			}
			report.Missing = append(report.Missing, missing.Missing...)
		}
		expected := SourceInventory{Bas: project.SourceCounts.Bas, Cls: project.SourceCounts.Cls, Frm: project.SourceCounts.Frm}
		report.Projects = append(report.Projects, ProjectInventory{
			ID: project.ID, Profile: project.Profile, Enabled: project.Enabled,
			Path: project.Path, Expected: expected, Actual: observed,
		})
		report.Sources = addSourceCounts(report.Sources, observed)
		profileCounts[project.Profile]++
		if project.Enabled {
			report.EnabledCount++
		} else {
			report.DisabledCount++
		}
	}
	sort.SliceStable(report.Projects, func(i, j int) bool { return report.Projects[i].ID < report.Projects[j].ID })
	sort.Strings(report.Missing)
	report.ProjectCount = len(report.Projects)
	profiles := make([]string, 0, len(profileCounts))
	for profile := range profileCounts {
		profiles = append(profiles, profile)
	}
	sort.Strings(profiles)
	for _, profile := range profiles {
		report.Profiles = append(report.Profiles, ProfileInventory{Profile: profile, Count: profileCounts[profile]})
	}
	return report, nil
}

// BuildCorpusInventory is a compatibility alias that makes the corpus scope
// explicit at call sites which also build other report types.
func BuildCorpusInventory(corpusRoot string, manifest Manifest) (InventoryReport, error) {
	return BuildInventory(corpusRoot, manifest)
}

// ValidateInventory verifies every manifest destination and source count. It
// rejects missing or extra project directories and returns the complete report
// even when count mismatches are found, making failures easy to inspect.
func ValidateInventory(corpusRoot string, manifest Manifest) (InventoryReport, error) {
	report, err := BuildInventory(corpusRoot, manifest)
	if err != nil {
		return InventoryReport{}, err
	}
	err = validateInventoryReport(corpusRoot, manifest, &report)
	return report, err
}

// ValidateCorpusInventory is the corresponding explicit-name alias.
func ValidateCorpusInventory(corpusRoot string, manifest Manifest) (InventoryReport, error) {
	return ValidateInventory(corpusRoot, manifest)
}

func validateInventoryReport(corpusRoot string, manifest Manifest, report *InventoryReport) error {
	if report == nil {
		return errors.New("nil inventory report")
	}
	expected := make(map[string]Project, len(manifest.Projects))
	for _, project := range manifest.Projects {
		expected[project.ID] = project
	}
	observed := make(map[string]struct{}, len(report.Projects))
	validation := &InventoryValidationError{}
	missing := make(map[string]struct{}, len(report.Missing))
	for _, project := range report.Missing {
		validation.Missing = append(validation.Missing, project)
		missing[project] = struct{}{}
	}
	for _, project := range report.Projects {
		if _, isMissing := missing[project.ID]; isMissing {
			continue
		}
		observed[project.ID] = struct{}{}
		manifestProject, ok := expected[project.ID]
		if !ok {
			validation.Unexpected = append(validation.Unexpected, project.ID)
			continue
		}
		want := SourceInventory{Bas: manifestProject.SourceCounts.Bas, Cls: manifestProject.SourceCounts.Cls, Frm: manifestProject.SourceCounts.Frm}
		if !equalSourceCounts(project.Actual, want) {
			validation.Mismatches = append(validation.Mismatches, InventoryMismatch{Project: project.ID, Expected: want, Actual: project.Actual})
		}
	}
	for _, project := range manifest.Projects {
		if _, ok := observed[project.ID]; !ok {
			if _, alreadyMissing := missing[project.ID]; !alreadyMissing {
				validation.Missing = append(validation.Missing, project.ID)
			}
		}
	}
	// BuildInventory visits manifest entries only, so inspect the managed
	// destination to detect stale project directories as well.
	if extra, err := inventoryUnexpectedProjects(corpusRoot, manifest); err != nil {
		return err
	} else {
		report.Unexpected = append(report.Unexpected, extra...)
		validation.Unexpected = append(validation.Unexpected, extra...)
	}
	sort.Strings(validation.Missing)
	sort.Strings(validation.Unexpected)
	sort.SliceStable(validation.Mismatches, func(i, j int) bool { return validation.Mismatches[i].Project < validation.Mismatches[j].Project })
	if len(validation.Missing) == 0 && len(validation.Unexpected) == 0 && len(validation.Mismatches) == 0 {
		return nil
	}
	return validation
}

func scanInventoryProject(root string, project Project) (SourceInventory, error) {
	projectRoot, err := containedPath(root, project.Path)
	if err != nil {
		return SourceInventory{}, fmt.Errorf("project %q inventory path: %w", project.ID, err)
	}
	info, err := os.Lstat(projectRoot)
	if err != nil {
		if os.IsNotExist(err) {
			return SourceInventory{}, &InventoryValidationError{Missing: []string{project.ID}}
		}
		return SourceInventory{}, fmt.Errorf("stat project %q inventory path: %w", project.ID, err)
	}
	if !info.IsDir() || info.Mode()&os.ModeSymlink != 0 {
		return SourceInventory{}, fmt.Errorf("project %q inventory path is not a regular directory", project.ID)
	}
	counts, err := CountSourceFiles(projectRoot)
	if err != nil {
		return SourceInventory{}, fmt.Errorf("scan project %q inventory: %w", project.ID, err)
	}
	return counts, nil
}

func inventoryUnexpectedProjects(root string, manifest Manifest) ([]string, error) {
	managed := filepath.Join(root, "projects", "third_party")
	entries, err := os.ReadDir(managed)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, fmt.Errorf("read managed corpus projects: %w", err)
	}
	expected := make(map[string]struct{}, len(manifest.Projects))
	for _, project := range manifest.Projects {
		expected[strings.ToLower(filepath.Base(filepath.FromSlash(project.Path)))] = struct{}{}
	}
	matched := make(map[string]struct{}, len(expected))
	var extra []string
	for _, entry := range entries {
		key := strings.ToLower(entry.Name())
		if _, ok := expected[key]; !ok {
			extra = append(extra, entry.Name())
			continue
		}
		if _, duplicate := matched[key]; duplicate {
			extra = append(extra, entry.Name())
			continue
		}
		matched[key] = struct{}{}
	}
	sort.Strings(extra)
	return extra, nil
}

// RuleDiagnosticSummary contains deterministic aggregate and per-project
// counts for one diagnostic code.
type RuleDiagnosticSummary struct {
	Code     string                   `json:"code"`
	Total    int                      `json:"total"`
	Lint     int                      `json:"lint"`
	Analyze  int                      `json:"analyze"`
	Projects []ProjectDiagnosticCount `json:"projects"`
}

type ProjectDiagnosticCount struct {
	Project string `json:"project"`
	Count   int    `json:"count"`
}

// DiagnosticSummary is a stable projection of a runner Report. Rules are
// ordered lexically and project rows are ordered by project ID.
type DiagnosticSummary struct {
	Total   int                     `json:"total"`
	Lint    int                     `json:"lint"`
	Analyze int                     `json:"analyze"`
	Rules   []RuleDiagnosticSummary `json:"rules"`
}

// SummarizeDiagnostics groups report diagnostics by rule, surface, and
// project. It does not mutate the report or deduplicate repeated rows.
func SummarizeDiagnostics(report Report) DiagnosticSummary {
	type aggregate struct {
		RuleDiagnosticSummary
		projects map[string]int
	}
	aggregates := make(map[string]*aggregate)
	summary := DiagnosticSummary{Rules: make([]RuleDiagnosticSummary, 0)}
	for _, diagnostic := range report.Diagnostics {
		summary.Total++
		switch diagnostic.Surface {
		case SurfaceLint:
			summary.Lint++
		case SurfaceAnalyze:
			summary.Analyze++
		}
		group := aggregates[diagnostic.Code]
		if group == nil {
			group = &aggregate{RuleDiagnosticSummary: RuleDiagnosticSummary{Code: diagnostic.Code}, projects: make(map[string]int)}
			aggregates[diagnostic.Code] = group
		}
		group.Total++
		switch diagnostic.Surface {
		case SurfaceLint:
			group.Lint++
		case SurfaceAnalyze:
			group.Analyze++
		}
		group.projects[diagnostic.Project]++
	}
	codes := make([]string, 0, len(aggregates))
	for code := range aggregates {
		codes = append(codes, code)
	}
	sort.Strings(codes)
	for _, code := range codes {
		group := aggregates[code]
		projects := make([]string, 0, len(group.projects))
		for project := range group.projects {
			projects = append(projects, project)
		}
		sort.Strings(projects)
		for _, project := range projects {
			group.Projects = append(group.Projects, ProjectDiagnosticCount{Project: project, Count: group.projects[project]})
		}
		summary.Rules = append(summary.Rules, group.RuleDiagnosticSummary)
	}
	return summary
}

// BuildDiagnosticSummary is a descriptive alias for callers that prefer the
// Build* naming used by corpus inventory helpers.
func BuildDiagnosticSummary(report Report) DiagnosticSummary { return SummarizeDiagnostics(report) }

// ByRule returns a defensive copy of the deterministic rule rows. The name
// mirrors SnapshotDiff.ByRule for callers that consume either projection.
func (s DiagnosticSummary) ByRule() []RuleDiagnosticSummary {
	rows := make([]RuleDiagnosticSummary, len(s.Rules))
	for i, rule := range s.Rules {
		rows[i] = rule
		rows[i].Projects = append([]ProjectDiagnosticCount(nil), rule.Projects...)
	}
	return rows
}

// FormatInventorySummary renders the stable fields useful in corpus test
// logs. It intentionally avoids map iteration and filesystem paths.
func FormatInventorySummary(report InventoryReport) string {
	var builder strings.Builder
	fmt.Fprintf(&builder, "corpus inventory: projects=%d enabled=%d disabled=%d bas=%d cls=%d frm=%d total=%d\n", report.ProjectCount, report.EnabledCount, report.DisabledCount, report.Sources.Bas, report.Sources.Cls, report.Sources.Frm, report.Sources.Total())
	for _, profile := range report.Profiles {
		fmt.Fprintf(&builder, "  profile=%s projects=%d\n", profile.Profile, profile.Count)
	}
	for _, project := range report.Projects {
		fmt.Fprintf(&builder, "  project=%s profile=%s enabled=%t bas=%d cls=%d frm=%d total=%d\n", project.ID, project.Profile, project.Enabled, project.Actual.Bas, project.Actual.Cls, project.Actual.Frm, project.Actual.Total())
	}
	return builder.String()
}

// FormatDiagnosticSummary renders one deterministic line per rule and keeps
// project-level counts visible for review logs.
func FormatDiagnosticSummary(summary DiagnosticSummary) string {
	var builder strings.Builder
	fmt.Fprintf(&builder, "corpus diagnostics: total=%d lint=%d analyze=%d rules=%d\n", summary.Total, summary.Lint, summary.Analyze, len(summary.Rules))
	for _, rule := range summary.Rules {
		fmt.Fprintf(&builder, "  rule=%s total=%d lint=%d analyze=%d", rule.Code, rule.Total, rule.Lint, rule.Analyze)
		if len(rule.Projects) > 0 {
			builder.WriteString(" projects=")
			for i, project := range rule.Projects {
				if i > 0 {
					builder.WriteByte(',')
				}
				fmt.Fprintf(&builder, "%s:%d", project.Project, project.Count)
			}
		}
		builder.WriteByte('\n')
	}
	return builder.String()
}
