package cli

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/sourceedit"
)

func runFixMultiApplyTest(t *testing.T, root string, ctx context.Context, rule string, paths ...string) (fixPreviewResult, output.Envelope, error) {
	t.Helper()
	args := append(append([]string(nil), paths...), "--rule", rule, "--apply")
	_, env, err := runFixTest(t, root, ctx, args...)
	if env.Output == nil {
		return fixPreviewResult{}, env, err
	}
	data, marshalErr := json.Marshal(env.Output)
	if marshalErr != nil {
		t.Fatal(marshalErr)
	}
	var result fixPreviewResult
	if decodeErr := json.Unmarshal(data, &result); decodeErr != nil {
		t.Fatalf("decode multi-file apply: %v; output=%s", decodeErr, data)
	}
	return result, env, err
}

func assertFixMultiApplyCounts(t *testing.T, result fixPreviewResult, files, changed, edits int) {
	t.Helper()
	if result.Summary.FileCount != files || result.Summary.ChangedFileCount != changed || result.Summary.EditCount != edits {
		t.Fatalf("summary=%+v, want files=%d changed=%d edits=%d", result.Summary, files, changed, edits)
	}
	if len(result.Files) != files {
		t.Fatalf("got %d result files, want %d", len(result.Files), files)
	}
}

func TestFixMultiApplyRulesMatchPreviewAndHashes(t *testing.T) {
	for _, tc := range []struct {
		rule, source, replacement, extra string
	}{
		{"VB044", "Sub Alpha()\n Const PROCEDURE_NAME As String = \"旧\" ' 日本語😀\nEnd Sub\n", "Sub Alpha()\n Const PROCEDURE_NAME As String = \"Alpha\" ' 日本語😀\nEnd Sub\n", "[lint.procedure_name_constant]\nenabled = true\nconstant_name = \"PROCEDURE_NAME\"\n"},
		{"VB080", "Sub Alpha()\n Dim x As Long\n Let x = 1 ' 日本語😀\n Let x = 2\nEnd Sub\n", "Sub Alpha()\n Dim x As Long\n x = 1 ' 日本語😀\n x = 2\nEnd Sub\n", ""},
		{"VB084", "Sub Alpha()\n Dim i As Long\n For i = 1 To 3 Step 1 ' 日本語😀\n Next i\nEnd Sub\n", "Sub Alpha()\n Dim i As Long\n For i = 1 To 3 ' 日本語😀\n Next i\nEnd Sub\n", ""},
		{"VB086", "Option Base 0 ' 日本語😀\nSub Alpha()\nEnd Sub\n", " ' 日本語😀\nSub Alpha()\nEnd Sub\n", ""},
	} {
		t.Run(tc.rule, func(t *testing.T) {
			root := t.TempDir()
			writeFixConfig(t, root, tc.extra)
			paths := []string{"src/modules/Alpha.bas", "src/modules/Beta.bas"}
			eols := []string{"\n", "\r\n"}
			wantSources := make([]string, len(paths))
			for i, path := range paths {
				source := strings.ReplaceAll(tc.source, "\n", eols[i])
				wantSources[i] = strings.ReplaceAll(tc.replacement, "\n", eols[i])
				writeEncodingFile(t, root, path, []byte(source))
			}
			writeEncodingFile(t, root, "src/modules/Unselected.bas", []byte{0xff})

			preview, previewEnv, previewErr := runFixMultiPreviewTest(t, root, t.Context(), tc.rule, paths[1], paths[0])
			if previewErr != nil || previewEnv.Status != output.StatusOK || preview.Mode != "preview" {
				t.Fatalf("preview=%+v envelope=%+v error=%v", preview, previewEnv, previewErr)
			}
			applied, applyEnv, applyErr := runFixMultiApplyTest(t, root, t.Context(), tc.rule, paths[1], paths[0])
			if applyErr != nil || applyEnv.Status != output.StatusOK || applied.Mode != "apply" || applied.Rule != tc.rule {
				t.Fatalf("apply=%+v envelope=%+v error=%v", applied, applyEnv, applyErr)
			}
			assertFixMultiApplyCounts(t, applied, len(paths), len(paths), preview.Summary.EditCount)
			if applied.Summary.FileCount != preview.Summary.FileCount || applied.Summary.ChangedFileCount != preview.Summary.ChangedFileCount || applied.Summary.EditCount != preview.Summary.EditCount || len(applied.Files) != len(preview.Files) {
				t.Fatalf("apply summary/files differ from preview: apply=%+v preview=%+v", applied, preview)
			}
			for i, path := range paths {
				want := preview.Files[i]
				want.Mode = "apply"
				want.Applied = true
				if !reflect.DeepEqual(applied.Files[i], want) {
					t.Fatalf("apply file %s differs from preview:\n apply=%+v\n want=%+v", path, applied.Files[i], want)
				}
				if applied.Files[i].SourceSHA256 != fmt.Sprintf("%x", sha256.Sum256([]byte(strings.ReplaceAll(tc.source, "\n", eols[i])))) ||
					applied.Files[i].ResultSHA256 != fmt.Sprintf("%x", sha256.Sum256([]byte(wantSources[i]))) {
					t.Fatalf("incorrect hashes for %s: %+v", path, applied.Files[i])
				}
				if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || string(got) != wantSources[i] {
					t.Fatalf("applied source %s=%q error=%v, want %q", path, got, readErr, wantSources[i])
				}
			}

			again, againEnv, againErr := runFixMultiApplyTest(t, root, t.Context(), tc.rule, paths...)
			if againErr != nil || againEnv.Status != output.StatusOK || again.Mode != "apply" || again.Summary.ChangedFileCount != 0 {
				t.Fatalf("repeat apply=%+v envelope=%+v error=%v", again, againEnv, againErr)
			}
			for _, file := range again.Files {
				if file.Changed || file.Applied || len(file.Edits) != 0 || file.Edits == nil || file.Diagnostics == nil {
					t.Fatalf("repeat apply produced a change: %+v", file)
				}
			}
		})
	}

}

func TestFixMultiApplyMixedChangesAndNoOps(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	paths := []string{"src/modules/Changed.bas", "src/modules/Suppressed.bas", "src/modules/Unchanged.bas"}
	sources := []string{
		"Sub A()\n Let x = 1\nEnd Sub\n",
		"Sub A()\n Let x = 1 ' xlflow:disable-line VB080\nEnd Sub\n",
		"Sub A()\nEnd Sub\n",
	}
	for i, path := range paths {
		writeEncodingFile(t, root, path, []byte(sources[i]))
	}

	result, env, err := runFixMultiApplyTest(t, root, t.Context(), "VB080", paths...)
	if err != nil || env.Status != output.StatusOK || result.Mode != "apply" {
		t.Fatalf("apply=%+v envelope=%+v error=%v", result, env, err)
	}
	assertFixMultiApplyCounts(t, result, 3, 1, 1)
	for i, file := range result.Files {
		if file.Path != paths[i] || file.Mode != "apply" || file.Applied != (i == 0) || file.Changed != (i == 0) {
			t.Fatalf("file[%d]=%+v", i, file)
		}
		if i > 0 && (len(file.Edits) != 0 || file.SourceSHA256 != file.ResultSHA256) {
			t.Fatalf("no-op file differs from its captured source: %+v", file)
		}
		want := sources[i]
		if i == 0 {
			want = "Sub A()\n x = 1\nEnd Sub\n"
		}
		if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(paths[i]))); readErr != nil || got == nil || string(got) != want {
			t.Fatalf("source %s=%q error=%v, want %q", paths[i], got, readErr, want)
		}
	}
}

func TestFixMultiApplyAllNoOpsAllowReadOnlyFiles(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "disabled_rules = [\"VB080\"]\n")
	paths := []string{"src/modules/First.bas", "src/modules/Second.bas"}
	source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
	for _, path := range paths {
		writeEncodingFile(t, root, path, source)
		if err := os.Chmod(filepath.Join(root, filepath.FromSlash(path)), 0o444); err != nil {
			t.Skipf("cannot mark source read-only: %v", err)
		}
	}
	t.Cleanup(func() {
		for _, path := range paths {
			_ = os.Chmod(filepath.Join(root, filepath.FromSlash(path)), 0o644)
		}
	})

	result, env, err := runFixMultiApplyTest(t, root, t.Context(), "VB080", paths...)
	if err != nil || env.Status != output.StatusOK || result.Mode != "apply" {
		t.Fatalf("apply=%+v envelope=%+v error=%v", result, env, err)
	}
	assertFixMultiApplyCounts(t, result, 2, 0, 0)
	for i, file := range result.Files {
		if file.Path != paths[i] || file.Changed || file.Applied || file.Mode != "apply" || len(file.Edits) != 0 {
			t.Fatalf("disabled-rule file=%+v", file)
		}
		if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(paths[i]))); readErr != nil || !bytes.Equal(got, source) {
			t.Fatalf("read-only source changed: %x error=%v", got, readErr)
		}
	}
}

func TestFixMultiApplyReadOnlyFailureDoesNotWriteAndProjectsDetails(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	paths := []string{"src/modules/ReadOnly.bas", "src/modules/Writable.bas"}
	source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
	for _, path := range paths {
		writeEncodingFile(t, root, path, source)
	}
	readOnlyPath := filepath.Join(root, filepath.FromSlash(paths[0]))
	if err := os.Chmod(readOnlyPath, 0o444); err != nil {
		t.Skipf("cannot mark source read-only: %v", err)
	}
	t.Cleanup(func() { _ = os.Chmod(readOnlyPath, 0o644) })

	_, env, err := runFixMultiApplyTest(t, root, t.Context(), "VB080", paths...)
	if err == nil || output.ExitCode(err) != output.ExitEnvironment || env.Error == nil || env.Error.Code != "fix_failed" || env.Output != nil {
		t.Fatalf("envelope=%+v error=%v, want atomic apply failure", env, err)
	}
	data, marshalErr := json.Marshal(env.Error.Details)
	if marshalErr != nil {
		t.Fatal(marshalErr)
	}
	var details struct {
		Committed bool `json:"committed"`
		Files     []struct {
			Path         string `json:"path"`
			State        string `json:"state"`
			RecoveryPath string `json:"recovery_path"`
		} `json:"files"`
		Artifacts []string `json:"artifacts"`
	}
	if decodeErr := json.Unmarshal(data, &details); decodeErr != nil {
		t.Fatalf("decode batch failure details %s: %v", data, decodeErr)
	}
	if details.Committed || len(details.Files) != 2 || len(details.Artifacts) != 0 {
		t.Fatalf("unexpected failure details: %+v", details)
	}
	for i, file := range details.Files {
		if file.Path != paths[i] || file.State != "not_applied" || file.RecoveryPath != "" {
			t.Fatalf("failure file[%d]=%+v", i, file)
		}
	}
	for _, path := range paths {
		if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || !bytes.Equal(got, source) {
			t.Fatalf("failed batch modified %s: %x error=%v", path, got, readErr)
		}
	}
}

func TestFixMultiApplyLateInvalidEncodingWritesNothing(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	firstPath := "src/modules/AValid.bas"
	firstSource := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
	invalidPath := "src/modules/ZInvalid.bas"
	invalidSource := []byte{0xff}
	writeEncodingFile(t, root, firstPath, firstSource)
	writeEncodingFile(t, root, invalidPath, invalidSource)

	_, env, err := runFixMultiApplyTest(t, root, t.Context(), "VB080", firstPath, invalidPath)
	if err == nil || env.Error == nil || env.Error.Code != "source_encoding_invalid" || env.Output != nil {
		t.Fatalf("envelope=%+v error=%v, want late encoding failure", env, err)
	}
	for path, want := range map[string][]byte{firstPath: firstSource, invalidPath: invalidSource} {
		if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || !bytes.Equal(got, want) {
			t.Fatalf("failed preparation modified %s: %x error=%v", path, got, readErr)
		}
	}
}

func TestFixMultiApplyDeduplicatesAliasesAndKeepsMultiFileSchema(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	path := "src/modules/Main.bas"
	source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
	writeEncodingFile(t, root, path, source)
	abs := filepath.Join(root, filepath.FromSlash(path))
	result, env, err := runFixMultiApplyTest(t, root, t.Context(), "VB080", path, abs, filepath.Join(".", filepath.FromSlash(path)))
	if err != nil || env.Status != output.StatusOK || result.Mode != "apply" {
		t.Fatalf("apply=%+v envelope=%+v error=%v", result, env, err)
	}
	assertFixMultiApplyCounts(t, result, 1, 1, 1)
	if result.Files[0].Path != path || result.Files[0].Mode != "apply" || !result.Files[0].Applied {
		t.Fatalf("deduplicated file result=%+v", result.Files[0])
	}
	data, marshalErr := json.Marshal(env.Output)
	if marshalErr != nil {
		t.Fatal(marshalErr)
	}
	var shape map[string]json.RawMessage
	if decodeErr := json.Unmarshal(data, &shape); decodeErr != nil {
		t.Fatal(decodeErr)
	}
	if _, hasFiles := shape["files"]; !hasFiles {
		t.Fatalf("multi-argument apply did not preserve multi-file schema: %s", data)
	}
	if _, hasPath := shape["path"]; hasPath {
		t.Fatalf("multi-file apply unexpectedly used single-file schema: %s", data)
	}
	if got, readErr := os.ReadFile(abs); readErr != nil || string(got) != "Sub A()\n x = 1\nEnd Sub\n" {
		t.Fatalf("deduplicated source=%q error=%v", got, readErr)
	}
}

func TestFixMultiApplyHumanOutputReportsAppliedFiles(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	paths := []string{"src/modules/First.bas", "src/modules/Second.bas"}
	for _, path := range paths {
		writeEncodingFile(t, root, path, []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
	}
	var stdout, stderr bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"fix", paths[0], paths[1], "--rule", "VB080", "--apply"})
	if err := cmd.ExecuteContext(t.Context()); err != nil {
		t.Fatalf("text apply failed: %v; stdout=%q stderr=%q", err, stdout.String(), stderr.String())
	}
	for _, want := range []string{
		"apply src/modules/First.bas: 1 edit(s), changed=true, applied=true",
		"apply src/modules/Second.bas: 1 edit(s), changed=true, applied=true",
		"apply: 2 file(s), 2 changed, 2 edit(s)",
	} {
		if !strings.Contains(stdout.String(), want) {
			t.Errorf("text output does not contain %q:\n%s", want, stdout.String())
		}
	}
}

func TestFixMultiApplyCancellationWritesNothing(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	paths := []string{"src/modules/First.bas", "src/modules/Second.bas"}
	source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
	for _, path := range paths {
		writeEncodingFile(t, root, path, source)
	}
	ctx, cancel := context.WithCancel(t.Context())
	cancel()

	_, env, err := runFixMultiApplyTest(t, root, ctx, "VB080", paths...)
	if err == nil || output.ExitCode(err) != output.ExitEnvironment || env.Error == nil || env.Error.Code != "fix_cancelled" || env.Output != nil {
		t.Fatalf("envelope=%+v error=%v, want cancellation", env, err)
	}
	for _, path := range paths {
		if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || !bytes.Equal(got, source) {
			t.Fatalf("cancelled apply modified %s: %x error=%v", path, got, readErr)
		}
	}
}

func TestFixMultiBatchFailureProjectionAndCategories(t *testing.T) {
	root := t.TempDir()
	sourcePath := filepath.Join(root, "src", "modules", "Main.bas")
	recoveryPath := filepath.Join(root, ".xlflow", "recovery", "Main.bas.bak")
	artifactPath := filepath.Join(root, ".xlflow", "recovery", "Main.bas.tmp")
	tests := []struct {
		name       string
		err        error
		result     sourceedit.BatchResult
		wantCode   string
		wantExit   int
		wantLogSub []string
		wantHuman  []string
	}{
		{
			name: "rollback failure",
			err:  &sourceedit.BatchError{Cause: context.Canceled, Rollback: errors.New("restore failed")},
			result: sourceedit.BatchResult{
				Files:     []sourceedit.BatchFileResult{{Path: sourcePath, State: "recovery_required", RecoveryPath: recoveryPath}},
				Artifacts: []string{artifactPath},
			},
			wantCode: "fix_rollback_failed", wantExit: output.ExitEnvironment,
			wantLogSub: []string{"apply src/modules/Main.bas: recovery_required", "recovery src/modules/Main.bas: .xlflow/recovery/Main.bas.bak", "artifact: .xlflow/recovery/Main.bas.tmp"},
			wantHuman:  []string{"apply src/modules/Main.bas: recovery_required", "recovery src/modules/Main.bas: .xlflow/recovery/Main.bas.bak", "artifact: .xlflow/recovery/Main.bas.tmp"},
		},
		{
			name:     "committed cleanup failure",
			err:      &sourceedit.BatchError{Cause: errors.New("cleanup failed"), Cleanup: errors.New("cleanup failed")},
			result:   sourceedit.BatchResult{Committed: true, Files: []sourceedit.BatchFileResult{{Path: sourcePath, State: "applied"}}},
			wantCode: "fix_cleanup_failed", wantExit: output.ExitEnvironment,
			wantLogSub: []string{"apply src/modules/Main.bas: applied"},
			wantHuman:  []string{"apply transaction: committed=true", "apply src/modules/Main.bas: applied"},
		},
		{
			name:     "stale source",
			err:      &sourceedit.BatchError{Cause: sourceedit.ErrStale},
			result:   sourceedit.BatchResult{Files: []sourceedit.BatchFileResult{{Path: sourcePath, State: "not_applied"}}},
			wantCode: "fix_source_changed", wantExit: output.ExitValidation,
		},
		{
			name:     "cancellation",
			err:      &sourceedit.BatchError{Cause: context.Canceled},
			result:   sourceedit.BatchResult{Files: []sourceedit.BatchFileResult{{Path: sourcePath, State: "not_applied"}}},
			wantCode: "fix_cancelled", wantExit: output.ExitEnvironment,
		},
		{
			name:     "cancellation precedes stale source in joined error",
			err:      &sourceedit.BatchError{Cause: errors.Join(sourceedit.ErrStale, context.Canceled)},
			result:   sourceedit.BatchResult{Files: []sourceedit.BatchFileResult{{Path: sourcePath, State: "not_applied"}}},
			wantCode: "fix_cancelled", wantExit: output.ExitEnvironment,
		},
		{
			name:     "duplicate physical identity",
			err:      &sourceedit.BatchError{Cause: errors.New("sourceedit: duplicate source identity")},
			result:   sourceedit.BatchResult{Files: []sourceedit.BatchFileResult{{Path: sourcePath, State: "not_applied"}}},
			wantCode: "fix_failed", wantExit: output.ExitEnvironment,
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			var stdout, stderr bytes.Buffer
			a := &app{cwd: root, stdout: &stdout, stderr: &stderr, json: true}
			gotErr := a.writeFixBatchFailure(tc.err, tc.result)
			if output.ExitCode(gotErr) != tc.wantExit {
				t.Fatalf("exit=%d error=%v, want %d", output.ExitCode(gotErr), gotErr, tc.wantExit)
			}
			var env output.Envelope
			if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
				t.Fatalf("decode error envelope %q: %v", stdout.String(), decodeErr)
			}
			if env.Status != output.StatusFailed || env.Error == nil || env.Error.Code != tc.wantCode || env.Output != nil {
				t.Fatalf("failure envelope=%+v", env)
			}
			data, marshalErr := json.Marshal(env.Error.Details)
			if marshalErr != nil {
				t.Fatal(marshalErr)
			}
			var details map[string]json.RawMessage
			if decodeErr := json.Unmarshal(data, &details); decodeErr != nil {
				t.Fatal(decodeErr)
			}
			for _, key := range []string{"committed", "files", "artifacts"} {
				if _, ok := details[key]; !ok {
					t.Errorf("error.details missing snake_case %q: %s", key, data)
				}
			}
			var committed bool
			if decodeErr := json.Unmarshal(details["committed"], &committed); decodeErr != nil || committed != tc.result.Committed {
				t.Fatalf("committed=%t error=%v, want %t", committed, decodeErr, tc.result.Committed)
			}
			for _, want := range tc.wantLogSub {
				if !strings.Contains(strings.Join(env.Logs, "\n"), want) {
					t.Errorf("logs do not contain %q: %+v", want, env.Logs)
				}
			}
			if len(tc.wantHuman) > 0 {
				var humanStdout, humanStderr bytes.Buffer
				humanApp := &app{cwd: root, stdout: &humanStdout, stderr: &humanStderr}
				humanErr := humanApp.writeFixBatchFailure(tc.err, tc.result)
				if output.ExitCode(humanErr) != tc.wantExit {
					t.Fatalf("human output exit=%d, want %d", output.ExitCode(humanErr), tc.wantExit)
				}
				for _, want := range tc.wantHuman {
					if !strings.Contains(humanStdout.String(), want) {
						t.Errorf("human output does not contain %q:\n%s", want, humanStdout.String())
					}
				}
			}
		})
	}
}
