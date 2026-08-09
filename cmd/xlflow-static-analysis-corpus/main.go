package main

import (
	"context"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"os"
	"os/signal"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/staticanalysis/corpus"
	"github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

func main() {
	ctx, stop := signal.NotifyContext(context.Background(), os.Interrupt)
	defer stop()
	os.Exit(run(ctx, os.Args[1:], os.Stdout, os.Stderr))
}

func run(ctx context.Context, args []string, stdout, stderr io.Writer) int {
	if len(args) == 0 {
		usage(stderr)
		return 2
	}
	var err error
	switch args[0] {
	case "sync":
		err = runSync(ctx, args[1:], stderr)
	case "review-details":
		err = runReviewDetails(args[1:], stdout, stderr, false)
	case "review-draft":
		err = runReviewDetails(args[1:], stdout, stderr, true)
	case "verify":
		err = runVerify(args[1:], stdout, stderr)
	default:
		usage(stderr)
		return 2
	}
	if err != nil {
		_, _ = fmt.Fprintln(stderr, err)
		return 1
	}
	return 0
}

func usage(w io.Writer) {
	_, _ = fmt.Fprintln(w, "usage: xlflow-static-analysis-corpus <sync|review-details|review-draft|verify> [options]")
}

func runSync(ctx context.Context, args []string, stderr io.Writer) error {
	flags := flag.NewFlagSet("sync", flag.ContinueOnError)
	flags.SetOutput(stderr)
	manifest := flags.String("manifest", "testdata/static-analysis-corpus/manifest.json", "corpus manifest")
	corpusRoot := flags.String("corpus-root", "testdata/static-analysis-corpus", "corpus root")
	checkout := flags.String("upstream-checkout", "", "local checkout at the pinned upstream commit")
	if err := flags.Parse(args); err != nil {
		return err
	}
	if flags.NArg() != 0 {
		return fmt.Errorf("sync does not accept positional arguments")
	}
	return corpus.Sync(ctx, corpus.SyncOptions{ManifestPath: *manifest, CorpusRoot: *corpusRoot, UpstreamCheckout: *checkout})
}

type reviewOptions struct {
	repoRoot            string
	rule                string
	project             string
	limit               int
	jsonOutput          bool
	classification      string
	rationale           string
	regressionTest      string
	regressionException string
}

func parseReviewOptions(name string, args []string, stderr io.Writer, draft bool) (reviewOptions, error) {
	flags := flag.NewFlagSet(name, flag.ContinueOnError)
	flags.SetOutput(stderr)
	var opts reviewOptions
	flags.StringVar(&opts.repoRoot, "repo-root", ".", "repository root")
	flags.StringVar(&opts.rule, "rule", "", "diagnostic rule ID")
	flags.StringVar(&opts.project, "project", "", "stable corpus project ID")
	flags.IntVar(&opts.limit, "limit", 20, "maximum candidate occurrences")
	flags.BoolVar(&opts.jsonOutput, "json", false, "emit JSON instead of TSV")
	if draft {
		flags.StringVar(&opts.classification, "classification", "", "true-positive or false-positive")
		flags.StringVar(&opts.rationale, "rationale", "", "human review rationale")
		flags.StringVar(&opts.regressionTest, "regression-test", "", "FP regression test path and name")
		flags.StringVar(&opts.regressionException, "regression-exception", "", "FP regression exception")
	}
	if err := flags.Parse(args); err != nil {
		return opts, err
	}
	if flags.NArg() != 0 {
		return opts, fmt.Errorf("%s does not accept positional arguments", name)
	}
	if opts.rule == "" {
		return opts, fmt.Errorf("--rule is required")
	}
	metadata, ok := rules.Lookup(opts.rule)
	if !ok {
		return opts, fmt.Errorf("unknown diagnostic rule %q", opts.rule)
	}
	opts.rule = metadata.ID
	if draft && opts.classification == "" {
		return opts, fmt.Errorf("--classification is required")
	}
	if draft {
		classification := corpus.ReviewClassification(opts.classification)
		if classification != corpus.ReviewTruePositive && classification != corpus.ReviewFalsePositive {
			return opts, fmt.Errorf("unsupported classification %q", opts.classification)
		}
		if strings.TrimSpace(opts.rationale) == "" {
			return opts, fmt.Errorf("--rationale is required")
		}
		hasTest := strings.TrimSpace(opts.regressionTest) != ""
		hasException := strings.TrimSpace(opts.regressionException) != ""
		if classification == corpus.ReviewFalsePositive && hasTest == hasException {
			return opts, fmt.Errorf("false-positive draft requires exactly one of --regression-test or --regression-exception")
		}
		if classification == corpus.ReviewTruePositive && (hasTest || hasException) {
			return opts, fmt.Errorf("true-positive draft does not accept regression evidence")
		}
	}
	return opts, nil
}

func runReviewDetails(args []string, stdout, stderr io.Writer, draft bool) error {
	opts, err := parseReviewOptions(map[bool]string{false: "review-details", true: "review-draft"}[draft], args, stderr, draft)
	if err != nil {
		return err
	}
	details, err := collectReviewDetails(opts)
	if err != nil {
		return err
	}
	if !draft {
		if opts.jsonOutput {
			encoder := json.NewEncoder(stdout)
			encoder.SetIndent("", "  ")
			return encoder.Encode(details)
		}
		_, err = io.WriteString(stdout, corpus.FormatReviewDetails(details))
		return err
	}
	drafts, err := corpus.BuildReviewDrafts(details, corpus.ReviewClassification(opts.classification), opts.rationale, opts.regressionTest, opts.regressionException)
	if err != nil {
		return err
	}
	encoded, err := corpus.EncodeReviewDrafts(drafts)
	if err != nil {
		return err
	}
	_, err = stdout.Write(encoded)
	return err
}

func collectReviewDetails(opts reviewOptions) ([]corpus.ReviewDetail, error) {
	repoRoot, err := filepath.Abs(opts.repoRoot)
	if err != nil {
		return nil, err
	}
	corpusRoot := filepath.Join(repoRoot, "testdata", "static-analysis-corpus")
	ids, err := corpus.DiscoverSnapshotIDs(filepath.Join(corpusRoot, "snapshots"))
	if err != nil {
		return nil, err
	}
	snapshots, err := corpus.LoadSnapshotSubset(filepath.Join(corpusRoot, "snapshots"), ids)
	if err != nil {
		return nil, err
	}
	reviews, err := corpus.LoadDiagnosticReviews(filepath.Join(corpusRoot, "reviews", "diagnostics.jsonl"))
	if err != nil {
		return nil, err
	}
	if opts.project != "" {
		snapshots = filterSnapshotProject(snapshots, opts.project)
		reviews = corpus.FilterReviews(reviews, []string{opts.project}, "")
		if len(snapshots) == 0 {
			return nil, fmt.Errorf("unknown corpus project %q", opts.project)
		}
	}
	candidates, err := corpus.FindSnapshotReviewCandidates(reviews, snapshots, opts.rule, opts.limit)
	if err != nil {
		return nil, err
	}
	projects := uniqueCandidateProjects(candidates.Rows)
	if len(projects) == 0 {
		return []corpus.ReviewDetail{}, nil
	}
	report, err := corpus.RunSelectedCorpus(repoRoot, projects)
	if err != nil {
		return nil, err
	}
	return corpus.ResolveReviewDetails(candidates, report)
}

func filterSnapshotProject(set corpus.SnapshotSet, project string) corpus.SnapshotSet {
	filtered := make(corpus.SnapshotSet)
	for id, rows := range set {
		if id.Project == project {
			filtered[id] = rows
		}
	}
	return filtered
}

func uniqueCandidateProjects(rows []corpus.SnapshotDiagnostic) []string {
	seen := make(map[string]struct{})
	for _, row := range rows {
		seen[row.Project] = struct{}{}
	}
	result := make([]string, 0, len(seen))
	for project := range seen {
		result = append(result, project)
	}
	sort.Strings(result)
	return result
}

func runVerify(args []string, stdout, stderr io.Writer) error {
	flags := flag.NewFlagSet("verify", flag.ContinueOnError)
	flags.SetOutput(stderr)
	repoRootFlag := flags.String("repo-root", ".", "repository root")
	project := flags.String("project", "", "stable corpus project ID")
	rule := flags.String("rule", "", "diagnostic rule ID")
	if err := flags.Parse(args); err != nil {
		return err
	}
	if flags.NArg() != 0 {
		return fmt.Errorf("verify does not accept positional arguments")
	}
	if *project == "" && *rule == "" {
		return fmt.Errorf("focused verify requires --project, --rule, or both; use corpus:test for the full gate")
	}
	if *rule != "" {
		metadata, ok := rules.Lookup(*rule)
		if !ok {
			return fmt.Errorf("unknown diagnostic rule %q", *rule)
		}
		*rule = metadata.ID
	}
	repoRoot, err := filepath.Abs(*repoRootFlag)
	if err != nil {
		return err
	}
	projects := []string(nil)
	if *project != "" {
		projects = []string{*project}
	}
	report, err := corpus.RunSelectedCorpus(repoRoot, projects)
	if err != nil {
		return err
	}
	report = corpus.FilterReport(report, *rule)
	actual, err := corpus.SnapshotSetFromReport(report)
	if err != nil {
		return err
	}
	snapshotRoot := filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "snapshots")
	committedIDs, err := corpus.DiscoverSnapshotIDs(snapshotRoot)
	if err != nil {
		return err
	}
	expectedIDs := selectSnapshotIDs(committedIDs, *project)
	want, err := corpus.LoadSnapshotSubset(snapshotRoot, expectedIDs)
	if err != nil {
		return err
	}
	for _, id := range actual.IDs() {
		if _, exists := want[id]; !exists {
			want[id] = []corpus.SnapshotDiagnostic{}
		}
	}
	want = corpus.FilterSnapshots(want, *rule)
	if diff := corpus.CompareSnapshotSets(want, actual); !diff.Empty() {
		return fmt.Errorf("%s", strings.TrimSpace(corpus.FormatSnapshotDiff(diff)))
	}
	reviews, err := corpus.LoadDiagnosticReviews(filepath.Join(repoRoot, "testdata", "static-analysis-corpus", "reviews", "diagnostics.jsonl"))
	if err != nil {
		return err
	}
	reviews = corpus.FilterReviews(reviews, projects, *rule)
	metrics, err := corpus.EvaluateDiagnosticReviews(reviews, report)
	if err != nil {
		return err
	}
	_, err = fmt.Fprint(stdout, corpus.FormatReviewMetrics(metrics))
	return err
}

func selectSnapshotIDs(ids []corpus.SnapshotID, project string) []corpus.SnapshotID {
	if project == "" {
		return append([]corpus.SnapshotID(nil), ids...)
	}
	selected := make([]corpus.SnapshotID, 0, 2)
	for _, id := range ids {
		if id.Project == project {
			selected = append(selected, id)
		}
	}
	return selected
}
