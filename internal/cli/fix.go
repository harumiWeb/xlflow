package cli

import (
	"cmp"
	"context"
	"crypto/sha256"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/sourceedit"
	"github.com/harumiWeb/xlflow/internal/sourcepath"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/sourceencoding"
	"github.com/harumiWeb/xlflow/internal/vbafmt"
)

type fixPosition struct {
	Line      int `json:"line"`
	Character int `json:"character"`
}

type fixRange struct {
	Start fixPosition `json:"start"`
	End   fixPosition `json:"end"`
}

type fixEdit struct {
	Range   fixRange `json:"range"`
	NewText string   `json:"new_text"`
}

type fixDiagnostic struct {
	Code       string   `json:"code"`
	Severity   string   `json:"severity"`
	Source     string   `json:"source"`
	Message    string   `json:"message"`
	Range      fixRange `json:"range"`
	Rule       string   `json:"rule"`
	Confidence string   `json:"confidence"`
}

type fixResult struct {
	Mode         string          `json:"mode"`
	Path         string          `json:"path"`
	Rule         string          `json:"rule"`
	Changed      bool            `json:"changed"`
	Applied      bool            `json:"applied"`
	Edits        []fixEdit       `json:"edits"`
	Diagnostics  []fixDiagnostic `json:"diagnostics"`
	Diff         string          `json:"diff"`
	SourceSHA256 string          `json:"source_sha256"`
	ResultSHA256 string          `json:"result_sha256"`
}

type fixPreviewSummary struct {
	FileCount        int `json:"file_count"`
	ChangedFileCount int `json:"changed_file_count"`
	EditCount        int `json:"edit_count"`
}

type fixPreviewResult struct {
	Mode    string            `json:"mode"`
	Rule    string            `json:"rule"`
	Files   []fixResult       `json:"files"`
	Summary fixPreviewSummary `json:"summary"`
}

func (a *app) fixCommand() *cobra.Command {
	var rules []string
	var apply bool
	cmd := &cobra.Command{
		Use:   "fix <path> [<path>...] --rule <code>",
		Short: "Preview or apply one rule's diagnostic fixes in explicit source files",
		Args: func(_ *cobra.Command, args []string) error {
			if len(args) == 0 {
				return a.writeFailure("fix", output.ExitConfig, "fix_args_invalid", errors.New("at least one source file is required"))
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			if len(rules) != 1 || strings.TrimSpace(rules[0]) == "" {
				return a.writeFailure("fix", output.ExitConfig, "fix_args_invalid", errors.New("exactly one --rule is required"))
			}
			rule := strings.ToUpper(strings.TrimSpace(rules[0]))
			kind := "source.fixAll.xlflow." + strings.ToLower(rule)
			if !slices.Contains(intel.FixAllCodeActionKinds(), kind) {
				return a.writeFailure("fix", output.ExitConfig, "fix_args_invalid", fmt.Errorf("rule %q is not eligible for CLI application", rule))
			}
			cfg, err := a.loadConfig("fix")
			if err != nil {
				return err
			}
			if len(args) > 1 {
				return a.runFixMulti(cmd.Context(), cfg, args, rule, kind, apply)
			}
			return a.runFix(cmd.Context(), cfg, args[0], rule, kind, apply)
		},
	}
	cmd.Flags().StringSliceVar(&rules, "rule", nil, "one eligible rule: VB044, VB080, VB084 or VB086 (does not enable disabled rules)")
	cmd.Flags().BoolVar(&apply, "apply", false, "write the proposed result after validating the captured source")
	return cmd
}

type fixSource struct {
	path     string
	relative string
}

func (a *app) resolveFixSource(ctx context.Context, cfg config.Config, arg string) (fixSource, error) {
	if err := ctx.Err(); err != nil {
		return fixSource{}, a.writeFixFailure(err)
	}
	path := filepath.FromSlash(strings.ReplaceAll(arg, `\`, "/"))
	if strings.TrimSpace(path) == "" || (strings.ToLower(filepath.Ext(path)) != ".bas" && strings.ToLower(filepath.Ext(path)) != ".cls") {
		return fixSource{}, a.writeFailure("fix", output.ExitConfig, "fix_args_invalid", errors.New("path must name one existing .bas or .cls source file"))
	}
	if !filepath.IsAbs(path) {
		path = filepath.Join(a.cwd, path)
	}
	path = filepath.Clean(path)
	// Reject directories before scoped discovery can walk them. Validate ownership
	// before reading the selected file's contents.
	info, err := os.Lstat(path)
	if err != nil {
		return fixSource{}, a.writeFixFailure(err)
	}
	if !info.Mode().IsRegular() {
		return fixSource{}, a.writeFailure("fix", output.ExitConfig, "fix_args_invalid", errors.New("path must name a regular source file without symlinks"))
	}
	files, err := sourceencoding.DiscoverFiles(ctx, sourceEncodingOptions(a.cwd, cfg, []string{arg}))
	if err != nil {
		if _, ok := errors.AsType[*sourceencoding.ScopeError](err); ok {
			return fixSource{}, a.writeFailure("fix", output.ExitConfig, "fix_args_invalid", err)
		}
		return fixSource{}, a.writeFixFailure(err)
	}
	if len(files) != 1 || !fixCanonicalFormCode(a.cwd, cfg, path) {
		return fixSource{}, a.writeFailure("fix", output.ExitConfig, "fix_args_invalid", errors.New("path must name one active managed source file"))
	}
	return fixSource{path: path, relative: files[0].Path}, nil
}

func (a *app) prepareFix(ctx context.Context, cfg config.Config, source fixSource, rule, kind string) (fixResult, *sourceedit.Snapshot, string, error) {
	path := source.path
	captured, err := sourceedit.Capture(ctx, path)
	if err != nil {
		return fixResult{}, nil, "", a.writeFixFailure(err)
	}
	if err := sourceencoding.Validate(source.relative, captured.Source); err != nil {
		return fixResult{}, nil, "", a.writeSourceEncodingResult(output.New("fix"), output.ExitValidation, "source_encoding_invalid", err)
	}
	doc := intel.Document{Path: captured.Path, Source: string(captured.Source), ModuleKind: "standard"}
	if strings.EqualFold(filepath.Ext(path), ".cls") {
		doc.ModuleKind = "class"
	}
	snapshot := intel.NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := intel.Analyzer{RootDir: a.cwd, Config: cfg, SourceOnly: true}
	actions, err := analyzer.CodeActionsContext(ctx, doc, intel.Range{}, []string{kind})
	if err != nil {
		return fixResult{}, nil, "", a.writeFixFailure(err)
	}
	result := fixResult{Mode: "preview", Path: source.relative, Rule: rule, Edits: []fixEdit{}, Diagnostics: []fixDiagnostic{}}
	edits := []intel.CodeActionEdit{}
	for _, action := range actions {
		if action.Kind != kind {
			continue
		}
		edits = append(edits, action.Edits...)
		for _, diagnostic := range action.Diagnostics {
			result.Diagnostics = append(result.Diagnostics, fixDiagnostic{Code: diagnostic.Code, Severity: diagnostic.Severity, Source: diagnostic.Source, Message: diagnostic.Message, Range: projectFixRange(diagnostic.Range), Rule: diagnostic.Rule, Confidence: diagnostic.Confidence})
		}
	}
	replacement, err := intel.ApplyCodeActionEdits(ctx, doc, edits)
	if err != nil {
		if errors.Is(err, intel.ErrInvalidCodeActionEdits) {
			return fixResult{}, nil, "", a.writeFailure("fix", output.ExitValidation, "fix_edits_invalid", err)
		}
		return fixResult{}, nil, "", a.writeFixFailure(err)
	}
	if err := sourceencoding.Validate(source.relative, []byte(replacement)); err != nil {
		return fixResult{}, nil, "", a.writeSourceEncodingResult(output.New("fix"), output.ExitValidation, "source_encoding_invalid", err)
	}
	for _, edit := range edits {
		result.Edits = append(result.Edits, fixEdit{Range: projectFixRange(edit.Range), NewText: edit.NewText})
	}
	result.Changed = replacement != doc.Source
	result.SourceSHA256 = fmt.Sprintf("%x", sha256.Sum256(captured.Source))
	result.ResultSHA256 = fmt.Sprintf("%x", sha256.Sum256([]byte(replacement)))
	if result.Changed {
		result.Diff = vbafmt.Diff(result.Path, doc.Source, replacement)
	}
	return result, captured, replacement, nil
}

func (a *app) runFix(ctx context.Context, cfg config.Config, arg, rule, kind string, apply bool) error {
	source, err := a.resolveFixSource(ctx, cfg, arg)
	if err != nil {
		return err
	}
	result, captured, replacement, err := a.prepareFix(ctx, cfg, source, rule, kind)
	if err != nil {
		return err
	}
	if apply {
		result.Mode = "apply"
		if result.Changed {
			if err := captured.Apply(ctx, []byte(replacement)); err != nil {
				return a.writeFixFailure(err)
			}
			result.Applied = true
		}
	}
	if !result.Applied {
		if err := ctx.Err(); err != nil {
			return a.writeFixFailure(err)
		}
	}
	env := output.New("fix")
	env.Target = map[string]any{"path": result.Path, "rule": rule}
	env.Output = result
	env.Logs = []string{fmt.Sprintf("%s %s: %d edit(s), changed=%t, applied=%t", result.Mode, result.Path, len(result.Edits), result.Changed, result.Applied)}
	if result.Diff != "" {
		env.Logs = append(env.Logs, result.Diff)
	}
	return a.write(env, output.ExitSuccess)
}

func (a *app) runFixMulti(ctx context.Context, cfg config.Config, args []string, rule, kind string, apply bool) error {
	unique := make(map[string]fixSource, len(args))
	for _, arg := range args {
		source, err := a.resolveFixSource(ctx, cfg, arg)
		if err != nil {
			return err
		}
		key := sourcepath.Key(source.path)
		if previous, exists := unique[key]; !exists || source.relative < previous.relative {
			unique[key] = source
		}
	}
	sources := make([]fixSource, 0, len(unique))
	for _, source := range unique {
		sources = append(sources, source)
	}
	slices.SortFunc(sources, func(a, b fixSource) int { return cmp.Compare(a.relative, b.relative) })
	mode := "preview"
	if apply {
		mode = "apply"
	}
	result := fixPreviewResult{Mode: mode, Rule: rule, Files: make([]fixResult, 0, len(sources))}
	env := output.New("fix")
	paths := make([]string, 0, len(sources))
	changes := make([]sourceedit.Change, 0, len(sources))
	for _, source := range sources {
		file, snapshot, replacement, err := a.prepareFix(ctx, cfg, source, rule, kind)
		if err != nil {
			return err
		}
		if apply {
			changes = append(changes, sourceedit.Change{Snapshot: snapshot, Replacement: []byte(replacement)})
			file.Mode = "apply"
		}
		result.Files = append(result.Files, file)
		paths = append(paths, file.Path)
		result.Summary.FileCount++
		result.Summary.EditCount += len(file.Edits)
		if file.Changed {
			result.Summary.ChangedFileCount++
		}
	}
	if err := ctx.Err(); err != nil {
		return a.writeFixFailure(err)
	}
	if apply {
		batch, err := sourceedit.ApplyBatch(ctx, changes)
		if err != nil {
			return a.writeFixBatchFailure(err, batch)
		}
		for i := range result.Files {
			result.Files[i].Applied = result.Files[i].Changed
		}
	}
	for _, file := range result.Files {
		env.Logs = append(env.Logs, fmt.Sprintf("%s %s: %d edit(s), changed=%t, applied=%t", mode, file.Path, len(file.Edits), file.Changed, file.Applied))
		if file.Diff != "" {
			env.Logs = append(env.Logs, file.Diff)
		}
	}
	env.Logs = append(env.Logs, fmt.Sprintf("%s: %d file(s), %d changed, %d edit(s)", mode, result.Summary.FileCount, result.Summary.ChangedFileCount, result.Summary.EditCount))
	env.Target = map[string]any{"paths": paths, "rule": rule}
	env.Output = result
	return a.write(env, output.ExitSuccess)
}

func projectFixRange(r intel.Range) fixRange {
	return fixRange{Start: fixPosition{Line: r.Start.Line, Character: r.Start.Character}, End: fixPosition{Line: r.End.Line, Character: r.End.Character}}
}

// Designer resources and inactive form sidecars are not source merely because
// their filenames happen to carry a VBA extension.
func fixCanonicalFormCode(root string, cfg config.Config, path string) bool {
	forms := cfg.Src.Forms
	if !filepath.IsAbs(forms) {
		forms = filepath.Join(root, forms)
	}
	rel, err := filepath.Rel(forms, path)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return true
	}
	parts := strings.Split(filepath.ToSlash(rel), "/")
	return len(parts) >= 2 && strings.EqualFold(parts[0], "code") && strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") && strings.EqualFold(filepath.Ext(path), ".bas")
}

func (a *app) writeFixFailure(err error) error {
	code, exit := "fix_failed", output.ExitEnvironment
	if errors.Is(err, sourceedit.ErrStale) {
		code, exit = "fix_source_changed", output.ExitValidation
	}
	if errors.Is(err, context.Canceled) || errors.Is(err, context.DeadlineExceeded) {
		code = "fix_cancelled"
	}
	return a.writeFailure("fix", exit, code, err)
}

type fixBatchFailureDetails struct {
	Committed bool                  `json:"committed"`
	Files     []fixBatchFailureFile `json:"files"`
	Artifacts []string              `json:"artifacts"`
}

type fixBatchFailureFile struct {
	Path         string `json:"path"`
	State        string `json:"state"`
	RecoveryPath string `json:"recovery_path,omitempty"`
}

func (a *app) writeFixBatchFailure(err error, batch sourceedit.BatchResult) error {
	code, exit := "fix_failed", output.ExitEnvironment
	batchErr, _ := errors.AsType[*sourceedit.BatchError](err)
	switch {
	case batchErr != nil && batchErr.Rollback != nil:
		code, exit = "fix_rollback_failed", output.ExitEnvironment
	case batchErr != nil && batch.Committed && batchErr.Cleanup != nil:
		code, exit = "fix_cleanup_failed", output.ExitEnvironment
	case errors.Is(err, context.Canceled), errors.Is(err, context.DeadlineExceeded):
		code, exit = "fix_cancelled", output.ExitEnvironment
	case errors.Is(err, sourceedit.ErrStale):
		code, exit = "fix_source_changed", output.ExitValidation
	}

	details := fixBatchFailureDetails{
		Committed: batch.Committed,
		Files:     make([]fixBatchFailureFile, 0, len(batch.Files)),
		Artifacts: make([]string, 0, len(batch.Artifacts)),
	}
	logs := []string{fmt.Sprintf("apply transaction: committed=%t", batch.Committed)}
	for _, file := range batch.Files {
		path := fixProjectPath(a.cwd, file.Path)
		recoveryPath := fixProjectPath(a.cwd, file.RecoveryPath)
		details.Files = append(details.Files, fixBatchFailureFile{Path: path, State: file.State, RecoveryPath: recoveryPath})
		logs = append(logs, fmt.Sprintf("apply %s: %s", path, file.State))
		if recoveryPath != "" {
			logs = append(logs, fmt.Sprintf("recovery %s: %s", path, recoveryPath))
		}
	}
	for _, artifact := range batch.Artifacts {
		artifact = fixProjectPath(a.cwd, artifact)
		details.Artifacts = append(details.Artifacts, artifact)
		logs = append(logs, "artifact: "+artifact)
	}

	env := output.Failure("fix", output.Error{Code: code, Message: err.Error(), Details: details})
	env.Logs = logs
	a.addConfigWarnings(&env)
	a.addPreflightWaiverWarnings(&env)
	if writeErr := output.WriteWithOptions(a.stdoutWriter(), env, a.outputOptions()); writeErr != nil {
		return output.WithExitCode(exit, writeErr)
	}
	return output.WithExitCode(exit, err)
}

func fixProjectPath(root, path string) string {
	if path == "" {
		return ""
	}
	path = filepath.FromSlash(path)
	if !filepath.IsAbs(path) {
		return filepath.ToSlash(filepath.Clean(path))
	}
	relative, err := filepath.Rel(root, path)
	if err == nil && relative != ".." && !strings.HasPrefix(relative, ".."+string(filepath.Separator)) {
		return filepath.ToSlash(relative)
	}
	return filepath.ToSlash(path)
}
