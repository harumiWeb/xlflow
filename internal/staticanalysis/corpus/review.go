package corpus

import (
	"bufio"
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"go/ast"
	"go/parser"
	"go/token"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/staticanalysis/contract"
)

const ReviewSchemaVersion = 1

type ReviewClassification string

const (
	ReviewTruePositive  ReviewClassification = "true-positive"
	ReviewFalsePositive ReviewClassification = "false-positive"
)

type ReviewedDiagnostic struct {
	Code               string         `json:"code"`
	Severity           string         `json:"severity"`
	Surface            string         `json:"surface"`
	Range              *ReviewedRange `json:"range"`
	Count              int            `json:"count"`
	AllowedOccurrences int            `json:"allowed_occurrences,omitempty"`
}

type ReviewedRange struct {
	StartLine   int `json:"start_line"`
	StartColumn int `json:"start_column"`
	EndLine     int `json:"end_line"`
	EndColumn   int `json:"end_column"`
}

// DiagnosticReview records a human classification for one stable corpus
// location. Diagnostics absent from this artifact remain unreviewed.
type DiagnosticReview struct {
	SchemaVersion       int                  `json:"schema_version"`
	Project             string               `json:"project"`
	File                string               `json:"file"`
	Classification      ReviewClassification `json:"classification"`
	Diagnostic          ReviewedDiagnostic   `json:"diagnostic"`
	Rationale           string               `json:"rationale"`
	RegressionTest      string               `json:"regression_test,omitempty"`
	RegressionException string               `json:"regression_exception,omitempty"`
}

type RuleReviewMetrics struct {
	Rule         string
	Reviewed     int
	Unreviewed   int
	TP           int
	FP           int
	Precision    float64
	HasPrecision bool
}

type ReviewMetrics struct {
	Reviewed   int
	Unreviewed int
	TP         int
	FP         int
	Rules      []RuleReviewMetrics
}

// LoadDiagnosticReviews loads strict, canonical JSON Lines review evidence.
func LoadDiagnosticReviews(path string) (reviews []DiagnosticReview, resultErr error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("open diagnostic reviews: %w", err)
	}
	defer func() {
		if err := file.Close(); err != nil {
			resultErr = errors.Join(resultErr, fmt.Errorf("close diagnostic reviews: %w", err))
		}
	}()

	reviews = make([]DiagnosticReview, 0)
	seen := make(map[string]struct{})
	previousKey := ""
	scanner := bufio.NewScanner(file)
	scanner.Buffer(make([]byte, 64*1024), 1024*1024)
	for lineNumber := 1; scanner.Scan(); lineNumber++ {
		line := append([]byte(nil), scanner.Bytes()...)
		if len(bytes.TrimSpace(line)) == 0 {
			return nil, fmt.Errorf("diagnostic reviews line %d is blank", lineNumber)
		}
		var review DiagnosticReview
		decoder := json.NewDecoder(bytes.NewReader(line))
		decoder.DisallowUnknownFields()
		if err := decoder.Decode(&review); err != nil {
			return nil, fmt.Errorf("decode diagnostic reviews line %d: %w", lineNumber, err)
		}
		if err := rejectTrailingJSON(decoder); err != nil {
			return nil, fmt.Errorf("decode diagnostic reviews line %d: %w", lineNumber, err)
		}
		if err := validateDiagnosticReview(review); err != nil {
			return nil, fmt.Errorf("diagnostic reviews line %d: %w", lineNumber, err)
		}
		canonical, err := json.Marshal(review)
		if err != nil {
			return nil, fmt.Errorf("encode diagnostic reviews line %d: %w", lineNumber, err)
		}
		if !bytes.Equal(line, canonical) {
			return nil, fmt.Errorf("diagnostic reviews line %d is not in canonical JSON form", lineNumber)
		}
		identity := diagnosticReviewIdentity(review)
		if _, duplicate := seen[identity]; duplicate {
			return nil, fmt.Errorf("diagnostic reviews line %d repeats an earlier diagnostic identity", lineNumber)
		}
		seen[identity] = struct{}{}
		key := diagnosticReviewSortKey(review)
		if previousKey != "" && key < previousKey {
			return nil, fmt.Errorf("diagnostic reviews line %d is out of canonical order", lineNumber)
		}
		previousKey = key
		reviews = append(reviews, review)
	}
	if err := scanner.Err(); err != nil {
		return nil, fmt.Errorf("read diagnostic reviews: %w", err)
	}
	return reviews, nil
}

func diagnosticReviewSortKey(review DiagnosticReview) string {
	return diagnosticReviewIdentity(review) + "\x00" + string(review.Classification)
}

func diagnosticReviewIdentity(review DiagnosticReview) string {
	rng := review.Diagnostic.Range
	return fmt.Sprintf("%s\x00%s\x00%s\x00%s\x00%s\x00%08d:%08d:%08d:%08d", review.Project, review.File, review.Diagnostic.Surface, review.Diagnostic.Code, review.Diagnostic.Severity, rng.StartLine, rng.StartColumn, rng.EndLine, rng.EndColumn)
}

func LoadReviews(path string) ([]DiagnosticReview, error) { return LoadDiagnosticReviews(path) }

// ValidateReviewSources ensures committed evidence still names a source in
// the native or vendored corpus tree.
func ValidateReviewSources(repoRoot, corpusRoot string, reviews []DiagnosticReview) error {
	for _, review := range reviews {
		var source string
		if project, ok := strings.CutPrefix(review.Project, "self/"); ok {
			source = filepath.Join(repoRoot, "example", filepath.FromSlash(project), filepath.FromSlash(review.File))
		} else if project, ok := strings.CutPrefix(review.Project, "third_party/"); ok {
			source = filepath.Join(corpusRoot, "projects", "third_party", filepath.FromSlash(project), filepath.FromSlash(review.File))
		} else {
			return fmt.Errorf("review project %q is outside the corpus", review.Project)
		}
		info, err := os.Stat(source)
		if err != nil {
			return fmt.Errorf("review source %s/%s: %w", review.Project, review.File, err)
		}
		if info.IsDir() {
			return fmt.Errorf("review source %s/%s is not a file", review.Project, review.File)
		}
		if review.RegressionTest != "" {
			parts := strings.Split(review.RegressionTest, "::")
			if len(parts) != 2 || strings.TrimSpace(parts[1]) == "" {
				return fmt.Errorf("review regression_test %q must be <path>::<test-name>", review.RegressionTest)
			}
			testPath := filepath.ToSlash(filepath.Clean(parts[0]))
			if testPath != parts[0] || filepath.IsAbs(parts[0]) || testPath == ".." || strings.HasPrefix(testPath, "../") || !strings.HasSuffix(testPath, "_test.go") {
				return fmt.Errorf("review regression_test %q has an invalid path", review.RegressionTest)
			}
			absoluteTestPath := filepath.Join(repoRoot, filepath.FromSlash(testPath))
			if info, statErr := os.Stat(absoluteTestPath); statErr != nil {
				return fmt.Errorf("review regression_test %q: %w", review.RegressionTest, statErr)
			} else if info.IsDir() {
				return fmt.Errorf("review regression_test %q is not a file", review.RegressionTest)
			}
			parsed, parseErr := parser.ParseFile(token.NewFileSet(), absoluteTestPath, nil, 0)
			if parseErr != nil {
				return fmt.Errorf("parse review regression_test %q: %w", review.RegressionTest, parseErr)
			}
			found := false
			for _, declaration := range parsed.Decls {
				function, ok := declaration.(*ast.FuncDecl)
				if ok && function.Recv == nil && function.Name.Name == parts[1] {
					found = true
					break
				}
			}
			if !found {
				return fmt.Errorf("review regression_test %q does not name a function in the test file", review.RegressionTest)
			}
		}
	}
	return nil
}

func rejectTrailingJSON(decoder *json.Decoder) error {
	var trailing any
	if err := decoder.Decode(&trailing); err == io.EOF {
		return nil
	} else if err == nil {
		return errors.New("review row contains trailing JSON")
	} else {
		return fmt.Errorf("review row contains trailing data: %w", err)
	}
}

func validateDiagnosticReview(review DiagnosticReview) error {
	if review.SchemaVersion != ReviewSchemaVersion {
		return fmt.Errorf("unsupported schema_version %d (want %d)", review.SchemaVersion, ReviewSchemaVersion)
	}
	if err := validateSnapshotDiagnosticIdentity(review.Project, review.File, review.Diagnostic.Code); err != nil {
		return err
	}
	if review.Classification != ReviewTruePositive && review.Classification != ReviewFalsePositive {
		return fmt.Errorf("unsupported classification %q", review.Classification)
	}
	if err := contract.ValidateExpectations([]contract.DiagnosticExpectation{review.expectation()}); err != nil {
		return err
	}
	rng := review.Diagnostic.Range
	if rng == nil || rng.StartLine < 1 {
		return errors.New("diagnostic review requires a valid exact range")
	}
	if review.Diagnostic.Count < 1 {
		return errors.New("diagnostic review count must be positive")
	}
	if review.Diagnostic.AllowedOccurrences < 0 {
		return errors.New("diagnostic review allowed_occurrences must not be negative")
	}
	if strings.TrimSpace(review.Rationale) == "" {
		return errors.New("diagnostic review rationale is required")
	}
	hasTest := strings.TrimSpace(review.RegressionTest) != ""
	hasException := strings.TrimSpace(review.RegressionException) != ""
	if review.Classification == ReviewFalsePositive && hasTest == hasException {
		return errors.New("false-positive review requires exactly one of regression_test or regression_exception")
	}
	if review.Classification == ReviewTruePositive && (hasTest || hasException) {
		return errors.New("true-positive review must not declare a regression reference")
	}
	if review.Classification == ReviewTruePositive && review.Diagnostic.AllowedOccurrences != 0 {
		return errors.New("true-positive review must not declare allowed_occurrences")
	}
	return nil
}

func (review DiagnosticReview) expectation() contract.DiagnosticExpectation {
	var rng *contract.Range
	if review.Diagnostic.Range != nil {
		value := review.Diagnostic.Range
		rng = &contract.Range{StartLine: value.StartLine, StartColumn: value.StartColumn, EndLine: value.EndLine, EndColumn: value.EndColumn}
	}
	return contract.DiagnosticExpectation{
		Code: review.Diagnostic.Code, Severity: review.Diagnostic.Severity,
		Range: rng, Surfaces: []string{review.Diagnostic.Surface},
	}
}

type reviewLocation struct{ project, file string }

// EvaluateDiagnosticReviews verifies required and forbidden contracts, then
// derives reviewed-only precision and the count of current unreviewed output.
func EvaluateDiagnosticReviews(reviews []DiagnosticReview, report Report) (ReviewMetrics, error) {
	byLocation := make(map[reviewLocation]contract.ExpectationSet)
	actualByLocation := make(map[reviewLocation][]contract.Diagnostic)
	for _, diagnostic := range report.Diagnostics {
		location := reviewLocation{project: diagnostic.Project, file: diagnostic.File}
		actualByLocation[location] = append(actualByLocation[location], contractDiagnostic(diagnostic))
	}
	for index, review := range reviews {
		if err := validateDiagnosticReview(review); err != nil {
			return ReviewMetrics{}, fmt.Errorf("review %d: %w", index+1, err)
		}
		location := reviewLocation{project: review.Project, file: review.File}
		expectations := byLocation[location]
		if review.Classification == ReviewTruePositive {
			for range review.Diagnostic.Count {
				expectations.ExpectedDiagnostics = append(expectations.ExpectedDiagnostics, review.expectation())
			}
		} else if review.Diagnostic.AllowedOccurrences == 0 {
			expectations.ForbiddenDiagnostics = append(expectations.ForbiddenDiagnostics, review.expectation())
		}
		byLocation[location] = expectations
	}
	locations := make([]reviewLocation, 0, len(byLocation))
	for location := range byLocation {
		locations = append(locations, location)
	}
	sort.Slice(locations, func(i, j int) bool {
		if locations[i].project != locations[j].project {
			return locations[i].project < locations[j].project
		}
		return locations[i].file < locations[j].file
	})
	violations := make([]string, 0)
	for _, location := range locations {
		expectations := byLocation[location]
		if err := contract.Check(expectations, actualByLocation[location]); err != nil {
			violations = append(violations, fmt.Sprintf("%s/%s: %v", location.project, location.file, err))
		}
	}
	for _, review := range reviews {
		if review.Classification != ReviewFalsePositive || review.Diagnostic.AllowedOccurrences == 0 {
			continue
		}
		location := reviewLocation{project: review.Project, file: review.File}
		count := 0
		for _, diagnostic := range actualByLocation[location] {
			if contract.Matches(review.expectation(), diagnostic) {
				count++
			}
		}
		if count > review.Diagnostic.AllowedOccurrences {
			violations = append(violations, fmt.Sprintf("%s/%s: forbidden diagnostic %s exceeded allowed occurrences (%d > %d)", review.Project, review.File, review.Diagnostic.Code, count, review.Diagnostic.AllowedOccurrences))
		}
	}
	if len(violations) > 0 {
		return ReviewMetrics{}, fmt.Errorf("review contract violations: %s", strings.Join(violations, "; "))
	}
	return buildReviewMetrics(reviews, report.Diagnostics), nil
}

func EvaluateReviews(reviews []DiagnosticReview, report Report) (ReviewMetrics, error) {
	return EvaluateDiagnosticReviews(reviews, report)
}

func contractDiagnostic(diagnostic Diagnostic) contract.Diagnostic {
	endLine, endColumn := normalizeDiagnosticEnd(diagnostic.Line, diagnostic.Column, diagnostic.EndLine, diagnostic.EndColumn)
	return contract.Diagnostic{
		Code: diagnostic.Code, Severity: diagnostic.Severity, Surface: string(diagnostic.Surface),
		Range: &contract.Range{StartLine: diagnostic.Line, StartColumn: diagnostic.Column, EndLine: endLine, EndColumn: endColumn},
	}
}

func buildReviewMetrics(reviews []DiagnosticReview, diagnostics []Diagnostic) ReviewMetrics {
	used := make([]bool, len(diagnostics))
	for _, review := range reviews {
		if review.Classification != ReviewTruePositive {
			continue
		}
		for range review.Diagnostic.Count {
			for index, diagnostic := range diagnostics {
				if used[index] || review.Project != diagnostic.Project || review.File != diagnostic.File || !contract.Matches(review.expectation(), contractDiagnostic(diagnostic)) {
					continue
				}
				used[index] = true
				break
			}
		}
	}
	codes := make([]string, len(diagnostics))
	for index := range diagnostics {
		codes[index] = diagnostics[index].Code
	}
	return buildReviewMetricsFromMatches(reviews, codes, used)
}

// EvaluateSnapshotReviews provides the fast committed-evidence view used by
// corpus:metrics. Snapshot rows retain only start positions, so this verifies
// true-positive multiplicity and computes coverage; the full real-world run
// remains the authority for exact expected and forbidden end-range contracts.
func EvaluateSnapshotReviews(reviews []DiagnosticReview, snapshots SnapshotSet) (ReviewMetrics, error) {
	rows := make([]SnapshotDiagnostic, 0)
	for _, id := range snapshots.IDs() {
		rows = append(rows, snapshots[id]...)
	}
	used, err := consumeSnapshotTruePositiveReviews(reviews, rows)
	if err != nil {
		return ReviewMetrics{}, err
	}
	codes := make([]string, len(rows))
	for index := range rows {
		codes[index] = rows[index].Code
	}
	return buildReviewMetricsFromMatches(reviews, codes, used), nil
}

func consumeSnapshotTruePositiveReviews(reviews []DiagnosticReview, rows []SnapshotDiagnostic) ([]bool, error) {
	used := make([]bool, len(rows))
	violations := make([]string, 0)
	for reviewIndex, review := range reviews {
		if err := validateDiagnosticReview(review); err != nil {
			return nil, fmt.Errorf("review %d: %w", reviewIndex+1, err)
		}
		if review.Classification == ReviewFalsePositive {
			continue
		}
		available := make([]int, 0)
		for index, row := range rows {
			if !used[index] && snapshotReviewMatches(review, row) {
				available = append(available, index)
			}
		}
		if len(available) < review.Diagnostic.Count {
			violations = append(violations, fmt.Sprintf("%s/%s: expected %d %s diagnostic(s), found %d", review.Project, review.File, review.Diagnostic.Count, review.Diagnostic.Code, len(available)))
			continue
		}
		for _, index := range available[:review.Diagnostic.Count] {
			used[index] = true
		}
	}
	if len(violations) > 0 {
		return nil, fmt.Errorf("committed review contract violations: %s", strings.Join(violations, "; "))
	}
	return used, nil
}

func snapshotReviewMatches(review DiagnosticReview, row SnapshotDiagnostic) bool {
	rng := review.Diagnostic.Range
	return review.Project == row.Project && review.File == row.File && review.Diagnostic.Surface == string(row.Surface) &&
		review.Diagnostic.Code == row.Code && review.Diagnostic.Severity == row.Severity && rng != nil &&
		rng.StartLine == row.Line && rng.StartColumn == row.Column
}

func buildReviewMetricsFromMatches(reviews []DiagnosticReview, codes []string, used []bool) ReviewMetrics {
	type counts struct{ tp, fp int }
	perRule := make(map[string]counts)
	metrics := ReviewMetrics{}
	for _, review := range reviews {
		count := perRule[review.Diagnostic.Code]
		if review.Classification == ReviewTruePositive {
			metrics.TP += review.Diagnostic.Count
			count.tp += review.Diagnostic.Count
		} else {
			metrics.FP += review.Diagnostic.Count
			count.fp += review.Diagnostic.Count
		}
		perRule[review.Diagnostic.Code] = count
	}
	metrics.Reviewed = metrics.TP + metrics.FP
	unreviewedByRule := make(map[string]int)
	for index, matched := range used {
		if !matched {
			metrics.Unreviewed++
			unreviewedByRule[codes[index]]++
			if _, exists := perRule[codes[index]]; !exists {
				perRule[codes[index]] = counts{}
			}
		}
	}

	rules := make([]string, 0, len(perRule))
	for code := range perRule {
		rules = append(rules, code)
	}
	sort.Strings(rules)
	for _, code := range rules {
		count := perRule[code]
		reviewed := count.tp + count.fp
		precision := 0.0
		if reviewed > 0 {
			precision = float64(count.tp) / float64(reviewed)
		}
		metrics.Rules = append(metrics.Rules, RuleReviewMetrics{Rule: code, Reviewed: reviewed, Unreviewed: unreviewedByRule[code], TP: count.tp, FP: count.fp, Precision: precision, HasPrecision: reviewed > 0})
	}
	return metrics
}

func FormatReviewMetrics(metrics ReviewMetrics) string {
	var builder strings.Builder
	fmt.Fprintf(&builder, "corpus reviews: reviewed=%d tp=%d fp=%d unreviewed=%d\n", metrics.Reviewed, metrics.TP, metrics.FP, metrics.Unreviewed)
	builder.WriteString("rule reviewed unreviewed tp fp precision\n")
	rules := append([]RuleReviewMetrics(nil), metrics.Rules...)
	sort.SliceStable(rules, func(i, j int) bool { return rules[i].Rule < rules[j].Rule })
	for _, rule := range rules {
		precision := "n/a"
		if rule.HasPrecision {
			precision = fmt.Sprintf("%.1f%%", rule.Precision*100)
		}
		fmt.Fprintf(&builder, "%s %d %d %d %d %s\n", rule.Rule, rule.Reviewed, rule.Unreviewed, rule.TP, rule.FP, precision)
	}
	return builder.String()
}
