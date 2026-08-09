package corpus

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

// ReviewDetail is the full normalized diagnostic identity used to prepare
// exact review-ledger evidence. Unlike snapshots, it retains the end position.
type ReviewDetail struct {
	Project   string  `json:"project"`
	File      string  `json:"file"`
	Surface   Surface `json:"surface"`
	Code      string  `json:"code"`
	Severity  string  `json:"severity"`
	Line      int     `json:"start_line"`
	Column    int     `json:"start_column"`
	EndLine   int     `json:"end_line"`
	EndColumn int     `json:"end_column"`
}

// DiscoverSnapshotIDs returns every committed project/surface snapshot ID.
func DiscoverSnapshotIDs(root string) ([]SnapshotID, error) {
	ids := make([]SnapshotID, 0)
	err := filepath.WalkDir(root, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() || filepath.Ext(path) != ".jsonl" {
			return nil
		}
		relative, err := filepath.Rel(root, path)
		if err != nil {
			return err
		}
		parts := strings.SplitN(filepath.ToSlash(relative), "/", 2)
		if len(parts) != 2 {
			return fmt.Errorf("invalid snapshot path %q", relative)
		}
		id := SnapshotID{Project: strings.TrimSuffix(parts[1], ".jsonl"), Surface: Surface(parts[0])}
		if err := validateSnapshotID(id); err != nil {
			return fmt.Errorf("snapshot path %q: %w", relative, err)
		}
		ids = append(ids, id)
		return nil
	})
	if err != nil {
		return nil, err
	}
	sort.Slice(ids, func(i, j int) bool { return snapshotIDLess(ids[i], ids[j]) })
	return ids, nil
}

// LoadSnapshotSubset loads only the requested files. It intentionally does
// not reject unrelated committed snapshots, which makes focused verification
// read-only and independent of the rest of the corpus.
func LoadSnapshotSubset(root string, ids []SnapshotID) (SnapshotSet, error) {
	set := make(SnapshotSet, len(ids))
	for _, id := range ids {
		if _, exists := set[id]; exists {
			return nil, fmt.Errorf("duplicate expected snapshot %s/%s", id.Project, id.Surface)
		}
		path, err := snapshotPath(root, id)
		if err != nil {
			return nil, err
		}
		rows, err := loadSnapshotFile(path, id)
		if err != nil {
			return nil, fmt.Errorf("load snapshot %s/%s: %w", id.Project, id.Surface, err)
		}
		set[id] = rows
	}
	return set, nil
}

// RunSelectedCorpus executes all corpus projects, or only the named stable
// project IDs. The returned report retains the ordinary production adapters.
func RunSelectedCorpus(repoRoot string, projectIDs []string) (Report, error) {
	native, err := DiscoverExampleProjects(repoRoot)
	if err != nil {
		return Report{}, err
	}
	manifestPath := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "manifest.json")
	manifest, corpusRoot, err := LoadManifest(manifestPath)
	if err != nil {
		return Report{}, err
	}
	if _, err := ValidateInventory(corpusRoot, manifest); err != nil {
		return Report{}, err
	}

	nativeIDs, thirdIDs := splitProjectIDs(projectIDs)
	selectedNative, nativeErr := SelectProjects(native, nativeIDs)
	thirdManifestIDs := make([]string, len(thirdIDs))
	for i, id := range thirdIDs {
		thirdManifestIDs[i] = strings.TrimPrefix(id, "third_party/")
	}
	selectedThird, thirdErr := SelectThirdPartyProjects(manifest.Projects, thirdManifestIDs)
	if len(projectIDs) > 0 && len(nativeIDs) == 0 {
		selectedNative = nil
	}
	if len(projectIDs) > 0 && len(thirdIDs) == 0 {
		selectedThird = nil
	}
	if nativeErr != nil || thirdErr != nil {
		return Report{}, errors.Join(nativeErr, thirdErr)
	}

	nativeReport, nativeRunErr := RunProjects(selectedNative)
	thirdReport, thirdRunErr := RunThirdPartyProjects(corpusRoot, selectedThird, MaterializeOptions{})
	combined := Report{}
	appendReport(&combined, nativeReport)
	appendReport(&combined, thirdReport)
	sortDiagnostics(combined.Diagnostics)
	sortThirdPartyFailuresAndSkips(&combined)
	return combined, errors.Join(nativeRunErr, thirdRunErr)
}

func splitProjectIDs(ids []string) (native, third []string) {
	for _, id := range ids {
		if strings.HasPrefix(id, "third_party/") {
			third = append(third, id)
		} else {
			native = append(native, id)
		}
	}
	return native, third
}

func appendReport(dst *Report, src Report) {
	dst.Surfaces = append(dst.Surfaces, src.Surfaces...)
	dst.Diagnostics = append(dst.Diagnostics, src.Diagnostics...)
	dst.Failures = append(dst.Failures, src.Failures...)
	dst.Skipped = append(dst.Skipped, src.Skipped...)
}

// FilterReport and FilterSnapshots retain only the selected rule. Project
// selection is performed before execution, so filtering by project here is
// unnecessary and cannot conceal an unexpected project result.
func FilterReport(report Report, rule string) Report {
	if rule == "" {
		return report
	}
	filtered := Report{Failures: report.Failures, Skipped: report.Skipped}
	for _, surface := range report.Surfaces {
		copySurface := surface
		copySurface.Diagnostics = filterDiagnostics(surface.Diagnostics, rule)
		filtered.Surfaces = append(filtered.Surfaces, copySurface)
		filtered.Diagnostics = append(filtered.Diagnostics, copySurface.Diagnostics...)
	}
	return filtered
}

func filterDiagnostics(rows []Diagnostic, rule string) []Diagnostic {
	result := make([]Diagnostic, 0)
	for _, row := range rows {
		if row.Code == rule {
			result = append(result, row)
		}
	}
	return result
}

func FilterSnapshots(set SnapshotSet, rule string) SnapshotSet {
	if rule == "" {
		return set
	}
	filtered := make(SnapshotSet, len(set))
	for id, rows := range set {
		for _, row := range rows {
			if row.Code == rule {
				filtered[id] = append(filtered[id], row)
			}
		}
		if filtered[id] == nil {
			filtered[id] = []SnapshotDiagnostic{}
		}
	}
	return filtered
}

func FilterReviews(reviews []DiagnosticReview, projects []string, rule string) []DiagnosticReview {
	wanted := make(map[string]struct{}, len(projects))
	for _, project := range projects {
		wanted[project] = struct{}{}
	}
	result := make([]DiagnosticReview, 0)
	for _, review := range reviews {
		if len(wanted) > 0 {
			if _, ok := wanted[review.Project]; !ok {
				continue
			}
		}
		if rule != "" && review.Diagnostic.Code != rule {
			continue
		}
		result = append(result, review)
	}
	return result
}

// ResolveReviewDetails joins start-only committed candidates with a fresh
// analyzer run and returns exact ranges while retaining multiplicity.
func ResolveReviewDetails(candidates SnapshotReviewCandidateReport, report Report) ([]ReviewDetail, error) {
	available := make(map[string][]Diagnostic)
	for _, diagnostic := range report.Diagnostics {
		available[startIdentity(diagnostic)] = append(available[startIdentity(diagnostic)], diagnostic)
	}
	details := make([]ReviewDetail, 0, len(candidates.Rows))
	for _, candidate := range candidates.Rows {
		key := snapshotStartIdentity(candidate)
		matches := available[key]
		if len(matches) == 0 {
			return nil, fmt.Errorf("fresh corpus run did not reproduce %s/%s:%d:%d [%s]", candidate.Project, candidate.File, candidate.Line, candidate.Column, candidate.Code)
		}
		diagnostic := matches[0]
		available[key] = matches[1:]
		details = append(details, ReviewDetail{Project: diagnostic.Project, File: diagnostic.File, Surface: diagnostic.Surface, Code: diagnostic.Code, Severity: diagnostic.Severity, Line: diagnostic.Line, Column: diagnostic.Column, EndLine: diagnostic.EndLine, EndColumn: diagnostic.EndColumn})
	}
	return details, nil
}

func startIdentity(d Diagnostic) string {
	return fmt.Sprintf("%s\x00%s\x00%s\x00%s\x00%s\x00%d\x00%d", d.Project, d.File, d.Surface, d.Code, d.Severity, d.Line, d.Column)
}

func snapshotStartIdentity(d SnapshotDiagnostic) string {
	return fmt.Sprintf("%s\x00%s\x00%s\x00%s\x00%s\x00%d\x00%d", d.Project, d.File, d.Surface, d.Code, d.Severity, d.Line, d.Column)
}

func FormatReviewDetails(details []ReviewDetail) string {
	var builder strings.Builder
	builder.WriteString("project\tfile\tsurface\tcode\tseverity\trange\n")
	for _, detail := range details {
		fmt.Fprintf(&builder, "%s\t%s\t%s\t%s\t%s\t%d:%d-%d:%d\n", detail.Project, detail.File, detail.Surface, detail.Code, detail.Severity, detail.Line, detail.Column, detail.EndLine, detail.EndColumn)
	}
	return builder.String()
}

// BuildReviewDrafts creates canonical ledger rows but never writes the ledger.
func BuildReviewDrafts(details []ReviewDetail, classification ReviewClassification, rationale, regressionTest, regressionException string) ([]DiagnosticReview, error) {
	grouped := make(map[string]DiagnosticReview)
	for _, detail := range details {
		review := DiagnosticReview{SchemaVersion: ReviewSchemaVersion, Project: detail.Project, File: detail.File, Classification: classification, Diagnostic: ReviewedDiagnostic{Code: detail.Code, Severity: detail.Severity, Surface: string(detail.Surface), Range: &ReviewedRange{StartLine: detail.Line, StartColumn: detail.Column, EndLine: detail.EndLine, EndColumn: detail.EndColumn}, Count: 1}, Rationale: rationale, RegressionTest: regressionTest, RegressionException: regressionException}
		key := diagnosticReviewIdentity(review)
		if existing, ok := grouped[key]; ok {
			existing.Diagnostic.Count++
			grouped[key] = existing
		} else {
			grouped[key] = review
		}
	}
	result := make([]DiagnosticReview, 0, len(grouped))
	for _, review := range grouped {
		if err := validateDiagnosticReview(review); err != nil {
			return nil, err
		}
		result = append(result, review)
	}
	sort.Slice(result, func(i, j int) bool { return diagnosticReviewSortKey(result[i]) < diagnosticReviewSortKey(result[j]) })
	return result, nil
}

func EncodeReviewDrafts(reviews []DiagnosticReview) ([]byte, error) {
	reviews = append([]DiagnosticReview(nil), reviews...)
	sort.Slice(reviews, func(i, j int) bool { return diagnosticReviewSortKey(reviews[i]) < diagnosticReviewSortKey(reviews[j]) })
	var output bytes.Buffer
	previousIdentity := ""
	for index, review := range reviews {
		if err := validateDiagnosticReview(review); err != nil {
			return nil, err
		}
		identity := diagnosticReviewIdentity(review)
		if index > 0 && identity == previousIdentity {
			return nil, fmt.Errorf("review draft repeats diagnostic identity %q", identity)
		}
		previousIdentity = identity
		row, err := json.Marshal(review)
		if err != nil {
			return nil, err
		}
		output.Write(row)
		output.WriteByte('\n')
	}
	return output.Bytes(), nil
}
