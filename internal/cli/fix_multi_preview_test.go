package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"runtime"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

type fixMultiPreviewTestResult struct {
	Mode    string                     `json:"mode"`
	Rule    string                     `json:"rule"`
	Files   []fixResult                `json:"files"`
	Summary fixMultiPreviewTestSummary `json:"summary"`
}

type fixMultiPreviewTestSummary struct {
	FileCount        int `json:"file_count"`
	ChangedFileCount int `json:"changed_file_count"`
	EditCount        int `json:"edit_count"`
}

func runFixMultiPreviewTest(t *testing.T, root string, ctx context.Context, rule string, paths ...string) (fixMultiPreviewTestResult, output.Envelope, error) {
	t.Helper()
	args := append(append([]string(nil), paths...), "--rule", rule)
	_, env, err := runFixTest(t, root, ctx, args...)
	if env.Output == nil {
		return fixMultiPreviewTestResult{}, env, err
	}
	data, marshalErr := json.Marshal(env.Output)
	if marshalErr != nil {
		t.Fatal(marshalErr)
	}
	var result fixMultiPreviewTestResult
	if decodeErr := json.Unmarshal(data, &result); decodeErr != nil {
		t.Fatalf("decode multi-file preview: %v; output=%s", decodeErr, data)
	}
	return result, env, err
}

func assertFixMultiPreviewCounts(t *testing.T, result fixMultiPreviewTestResult, files, changed, edits int) {
	t.Helper()
	if result.Summary.FileCount != files || result.Summary.ChangedFileCount != changed || result.Summary.EditCount != edits {
		t.Fatalf("summary=%+v, want files=%d changed=%d edits=%d", result.Summary, files, changed, edits)
	}
	if len(result.Files) != files {
		t.Fatalf("got %d result files, want %d", len(result.Files), files)
	}
}

func TestFixMultiPreviewRulesMatchSingleFilePreviews(t *testing.T) {
	for _, tc := range []struct {
		rule, source, extra string
	}{
		{"VB044", "Sub Alpha()\n Const PROCEDURE_NAME As String = \"旧\" ' 日本語😀\nEnd Sub\n", "[lint.procedure_name_constant]\nenabled = true\nconstant_name = \"PROCEDURE_NAME\"\n"},
		{"VB080", "Sub Alpha()\n Dim x As Long\n Let x = 1 ' 日本語😀\nEnd Sub\n", ""},
		{"VB084", "Sub Alpha()\n Dim i As Long\n For i = 1 To 3 Step 1 ' 日本語😀\n Next i\nEnd Sub\n", ""},
		{"VB086", "Option Base 0 ' 日本語😀\nSub Alpha()\nEnd Sub\n", ""},
	} {
		t.Run(tc.rule, func(t *testing.T) {
			root := t.TempDir()
			writeFixConfig(t, root, tc.extra)
			paths := []string{"src/modules/Alpha.bas", "src/modules/Beta.bas"}
			eols := []string{"\n", "\r\n"}
			for i, path := range paths {
				source := strings.ReplaceAll(tc.source, "\n", eols[i])
				writeEncodingFile(t, root, path, []byte(source))
			}
			// Discovery must stay scoped to the selected paths.
			writeEncodingFile(t, root, "src/modules/Unselected.bas", []byte{0xff})

			got, env, err := runFixMultiPreviewTest(t, root, t.Context(), tc.rule, paths[1], paths[0])
			if err != nil || env.Status != output.StatusOK || got.Mode != "preview" || got.Rule != tc.rule {
				t.Fatalf("preview=%+v envelope=%+v error=%v", got, env, err)
			}
			assertFixMultiPreviewCounts(t, got, 2, 2, 2)
			if got.Files[0].Path != paths[0] || got.Files[1].Path != paths[1] {
				t.Fatalf("paths are not normalized and sorted: %+v", got.Files)
			}
			for i, path := range paths {
				single, singleEnv, singleErr := runFixTest(t, root, t.Context(), path, "--rule", tc.rule)
				if singleErr != nil || singleEnv.Status != output.StatusOK {
					t.Fatalf("single preview path=%s result=%+v envelope=%+v error=%v", path, single, singleEnv, singleErr)
				}
				if !reflect.DeepEqual(got.Files[i], single) {
					t.Fatalf("multi result differs from single preview for %s:\n multi=%+v\n single=%+v", path, got.Files[i], single)
				}
				if got.Files[i].Applied || got.Files[i].Mode != "preview" {
					t.Fatalf("file result is not preview-only: %+v", got.Files[i])
				}
				data, marshalErr := json.Marshal(singleEnv.Output)
				if marshalErr != nil {
					t.Fatal(marshalErr)
				}
				var shape map[string]json.RawMessage
				if decodeErr := json.Unmarshal(data, &shape); decodeErr != nil {
					t.Fatal(decodeErr)
				}
				if _, hasFiles := shape["files"]; hasFiles {
					t.Fatalf("single-file output unexpectedly uses multi-file schema: %s", data)
				}
				wantSource := strings.ReplaceAll(tc.source, "\n", eols[i])
				if source, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || string(source) != wantSource {
					t.Fatalf("preview modified %s: source=%q error=%v", path, source, readErr)
				}
			}
		})
	}
}

func TestFixMultiPreviewDeduplicatesAliasesAndSortsPaths(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	paths := []string{"src/modules/Zed.bas", "src/modules/apple.bas", "src/modules/mango.bas"}
	for _, path := range paths {
		writeEncodingFile(t, root, path, []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
	}
	absUpper := filepath.Join(root, filepath.FromSlash(paths[0]))
	args := []string{
		paths[2], absUpper, filepath.Join(".", filepath.FromSlash(paths[1])),
		filepath.Join(".", filepath.FromSlash(paths[0])), paths[1], paths[2],
	}
	first, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", args...)
	if err != nil || env.Status != output.StatusOK || first.Mode != "preview" || first.Rule != "VB080" {
		t.Fatalf("preview=%+v envelope=%+v error=%v", first, env, err)
	}
	assertFixMultiPreviewCounts(t, first, 3, 3, 3)
	wantPaths := []string{"src/modules/Zed.bas", "src/modules/apple.bas", "src/modules/mango.bas"}
	for i, path := range wantPaths {
		if first.Files[i].Path != path {
			t.Fatalf("file[%d].Path=%q, want %q", i, first.Files[i].Path, path)
		}
	}

	reversed := append([]string(nil), args...)
	for i, j := 0, len(reversed)-1; i < j; i, j = i+1, j-1 {
		reversed[i], reversed[j] = reversed[j], reversed[i]
	}
	second, secondEnv, secondErr := runFixMultiPreviewTest(t, root, t.Context(), "VB080", reversed...)
	if secondErr != nil || secondEnv.Status != output.StatusOK || !reflect.DeepEqual(first, second) {
		t.Fatalf("preview ordering depends on input order:\n first=%+v\n second=%+v\n envelope=%+v error=%v", first, second, secondEnv, secondErr)
	}

	// Multiple arguments still select the batch schema when aliases deduplicate to one file.
	single, singleEnv, singleErr := runFixMultiPreviewTest(t, root, t.Context(), "VB080", absUpper, paths[0], filepath.Join(".", filepath.FromSlash(paths[0])))
	if singleErr != nil || singleEnv.Status != output.StatusOK || single.Mode != "preview" {
		t.Fatalf("deduplicated preview=%+v envelope=%+v error=%v", single, singleEnv, singleErr)
	}
	assertFixMultiPreviewCounts(t, single, 1, 1, 1)
	if single.Files[0].Path != paths[0] {
		t.Fatalf("deduplicated path=%q, want %q", single.Files[0].Path, paths[0])
	}
}

func TestFixMultiPreviewWindowsCaseOnlyAliases(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	upperPath := "src/modules/Case.bas"
	writeEncodingFile(t, root, upperPath, []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
	paths := []string{upperPath, "src/modules/case.bas"}
	wantFiles := 2
	if runtime.GOOS == "windows" {
		wantFiles = 1
	} else {
		writeEncodingFile(t, root, paths[1], []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
	}

	result, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", paths...)
	if err != nil || env.Status != output.StatusOK {
		t.Fatalf("preview=%+v envelope=%+v error=%v", result, env, err)
	}
	assertFixMultiPreviewCounts(t, result, wantFiles, wantFiles, wantFiles)
	for _, path := range paths {
		got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path)))
		if readErr != nil || string(got) != "Sub A()\n Let x = 1\nEnd Sub\n" {
			t.Fatalf("preview modified %s: source=%q error=%v", path, got, readErr)
		}
	}
}

func TestFixMultiPreviewIncludesUnchangedSuppressedAndDisabledFiles(t *testing.T) {
	t.Run("changed suppressed and unchanged", func(t *testing.T) {
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
		result, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", paths...)
		if err != nil || env.Status != output.StatusOK {
			t.Fatalf("preview=%+v envelope=%+v error=%v", result, env, err)
		}
		assertFixMultiPreviewCounts(t, result, 3, 1, 1)
		for i, file := range result.Files {
			if file.Path != paths[i] || file.Applied || file.Mode != "preview" {
				t.Fatalf("file[%d]=%+v", i, file)
			}
			if i == 0 && (!file.Changed || len(file.Edits) != 1) {
				t.Fatalf("changed result=%+v", file)
			}
			if i > 0 && (file.Changed || len(file.Edits) != 0 || file.Diff != "" || file.SourceSHA256 != file.ResultSHA256) {
				t.Fatalf("no-change result=%+v", file)
			}
			if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(paths[i]))); readErr != nil || string(got) != sources[i] {
				t.Fatalf("preview modified %s: source=%q error=%v", paths[i], got, readErr)
			}
		}
	})

	t.Run("disabled rule remains included", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "disabled_rules = [\"VB080\"]\n")
		paths := []string{"src/modules/First.bas", "src/modules/Second.bas"}
		for _, path := range paths {
			writeEncodingFile(t, root, path, []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
		}
		result, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", paths...)
		if err != nil || env.Status != output.StatusOK {
			t.Fatalf("preview=%+v envelope=%+v error=%v", result, env, err)
		}
		assertFixMultiPreviewCounts(t, result, 2, 0, 0)
		for _, file := range result.Files {
			if file.Changed || file.Applied || len(file.Edits) != 0 || len(file.Diagnostics) != 0 || file.Diff != "" {
				t.Fatalf("disabled rule produced a correction: %+v", file)
			}
		}
	})
}

func TestFixMultiPreviewAcceptsManagedConfiguredRoots(t *testing.T) {
	root := t.TempDir()
	writeEncodingFile(t, root, config.FileName, []byte(`[src]
modules = "custom/modules"
classes = "custom/classes"
forms = "custom/forms"
workbook = "custom/workbook"
[lint]
require_option_explicit = false
detect_let_assignment = true
[userform]
code_source = "sidecar"
`))
	paths := []string{
		"custom/modules/Module.bas",
		"custom/classes/Class.cls",
		"custom/forms/code/Form.bas",
		"custom/workbook/ThisWorkbook.cls",
		"tests/Case.cls",
	}
	for _, path := range paths {
		writeEncodingFile(t, root, path, []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
	}
	result, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", paths...)
	if err != nil || env.Status != output.StatusOK {
		t.Fatalf("preview=%+v envelope=%+v error=%v", result, env, err)
	}
	assertFixMultiPreviewCounts(t, result, len(paths), len(paths), len(paths))
	for _, file := range result.Files {
		if !file.Changed || file.Applied || len(file.Edits) != 1 {
			t.Fatalf("managed path was not previewed: %+v", file)
		}
	}
}

func TestFixMultiPreviewFailuresDoNotPublishPartialResults(t *testing.T) {
	for _, tc := range []struct {
		name, secondPath, errorCode string
		secondBytes                 []byte
	}{
		{name: "unsupported file", secondPath: "src/modules/Unsupported.frm", errorCode: "fix_args_invalid", secondBytes: []byte("form")},
		{name: "late invalid encoding", secondPath: "src/modules/ZInvalid.bas", errorCode: "source_encoding_invalid", secondBytes: []byte{0xff}},
		{name: "outside source root", secondPath: "outside/Other.bas", errorCode: "fix_args_invalid", secondBytes: []byte("Sub A()\n Let x = 1\nEnd Sub\n")},
	} {
		t.Run(tc.name, func(t *testing.T) {
			root := t.TempDir()
			writeFixConfig(t, root, "")
			firstPath := "src/modules/Valid.bas"
			firstSource := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
			writeEncodingFile(t, root, firstPath, firstSource)
			writeEncodingFile(t, root, tc.secondPath, tc.secondBytes)
			_, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", firstPath, tc.secondPath)
			if err == nil || env.Status == output.StatusOK || env.Error == nil || env.Error.Code != tc.errorCode || env.Output != nil {
				t.Fatalf("envelope=%+v error=%v, want error %s and no preview payload", env, err, tc.errorCode)
			}
			if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(firstPath))); readErr != nil || !bytes.Equal(got, firstSource) {
				t.Fatalf("valid source changed after failed batch: %x error=%v", got, readErr)
			}
			if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(tc.secondPath))); readErr != nil || !bytes.Equal(got, tc.secondBytes) {
				t.Fatalf("invalid source changed after failed batch: %x error=%v", got, readErr)
			}
		})
	}

	t.Run("inactive form sidecar", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "[userform]\ncode_source = \"frm\"\n")
		firstPath := "src/modules/Valid.bas"
		firstSource := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
		sidecarPath := "src/forms/code/Form.bas"
		sidecarSource := []byte("Sub FormCode()\n Let x = 1\nEnd Sub\n")
		writeEncodingFile(t, root, firstPath, firstSource)
		writeEncodingFile(t, root, sidecarPath, sidecarSource)
		_, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", firstPath, sidecarPath)
		if err == nil || env.Error == nil || env.Error.Code != "fix_args_invalid" || env.Output != nil {
			t.Fatalf("envelope=%+v error=%v, want inactive sidecar rejection without preview payload", env, err)
		}
		for path, want := range map[string][]byte{firstPath: firstSource, sidecarPath: sidecarSource} {
			if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || !bytes.Equal(got, want) {
				t.Fatalf("failed batch modified %s: source=%x error=%v", path, got, readErr)
			}
		}
	})

	t.Run("directory with source suffix", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "")
		firstPath := "src/modules/Valid.bas"
		firstSource := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
		writeEncodingFile(t, root, firstPath, firstSource)
		if err := os.MkdirAll(filepath.Join(root, "src/modules/Directory.bas"), 0o755); err != nil {
			t.Fatal(err)
		}
		_, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", firstPath, "src/modules/Directory.bas")
		if err == nil || env.Error == nil || env.Error.Code != "fix_args_invalid" || env.Output != nil {
			t.Fatalf("envelope=%+v error=%v, want directory rejection without preview payload", env, err)
		}
		if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(firstPath))); readErr != nil || !bytes.Equal(got, firstSource) {
			t.Fatalf("valid source changed after directory rejection: %x error=%v", got, readErr)
		}
	})

	t.Run("symlink", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "")
		path := filepath.Join(root, "src/modules/Target.bas")
		writeEncodingFile(t, root, "src/modules/Target.bas", []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
		alias := filepath.Join(root, "src/modules/Alias.bas")
		if err := os.Symlink(path, alias); err != nil {
			t.Skipf("symlink creation is unavailable: %v", err)
		}
		_, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", "src/modules/Target.bas", "src/modules/Alias.bas")
		if err == nil || env.Error == nil || env.Error.Code != "fix_args_invalid" || env.Output != nil {
			t.Fatalf("envelope=%+v error=%v, want symlink rejection without preview payload", env, err)
		}
		if got, readErr := os.ReadFile(path); readErr != nil || string(got) != "Sub A()\n Let x = 1\nEnd Sub\n" {
			t.Fatalf("target changed after symlink rejection: %q error=%v", got, readErr)
		}
	})
}

func TestFixMultiPreviewTextShowsDiffsAndAggregateCounts(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	paths := []string{"src/modules/First.bas", "src/modules/Second.bas"}
	source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
	for _, path := range paths {
		writeEncodingFile(t, root, path, source)
	}

	var stdout, stderr bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"fix", paths[0], paths[1], "--rule", "VB080"})
	err := cmd.ExecuteContext(t.Context())
	if err != nil {
		t.Fatalf("text preview failed: %v; stdout=%q stderr=%q", err, stdout.String(), stderr.String())
	}
	text := stdout.String()
	for _, want := range []string{
		"preview src/modules/First.bas: 1 edit(s), changed=true, applied=false",
		"--- a/src/modules/First.bas",
		"+++ b/src/modules/First.bas",
		"preview src/modules/Second.bas: 1 edit(s), changed=true, applied=false",
		"--- a/src/modules/Second.bas",
		"+++ b/src/modules/Second.bas",
		"preview: 2 file(s), 2 changed, 2 edit(s)",
	} {
		if !strings.Contains(text, want) {
			t.Errorf("text output does not contain %q:\n%s", want, text)
		}
	}
	for _, path := range paths {
		if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || !bytes.Equal(got, source) {
			t.Fatalf("text preview modified %s: %q error=%v", path, got, readErr)
		}
	}
}

func TestFixMultiApplyRequiresConfig(t *testing.T) {
	root := t.TempDir()
	path := "src/modules/Main.bas"
	source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
	writeEncodingFile(t, root, path, source)
	abs := filepath.Join(root, filepath.FromSlash(path))
	_, env, err := runFixTest(t, root, t.Context(), path, abs, "--rule", "VB080", "--apply")
	if err == nil || env.Error == nil || env.Error.Code != "config_error" || env.Output != nil {
		t.Fatalf("envelope=%+v error=%v, want missing-config failure", env, err)
	}
	if got, readErr := os.ReadFile(abs); readErr != nil || !bytes.Equal(got, source) {
		t.Fatalf("source changed after missing-config failure: %q error=%v", got, readErr)
	}
}

func TestFixMultiPreviewReadOnlyAndCancellation(t *testing.T) {
	t.Run("read-only source", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "")
		paths := []string{"src/modules/ReadOnly.bas", "src/modules/Writable.bas"}
		source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
		for _, path := range paths {
			writeEncodingFile(t, root, path, source)
		}
		readOnlyPath := filepath.Join(root, filepath.FromSlash(paths[0]))
		if err := os.Chmod(readOnlyPath, 0o444); err != nil {
			t.Skipf("cannot mark file read-only: %v", err)
		}
		t.Cleanup(func() { _ = os.Chmod(readOnlyPath, 0o644) })
		result, env, err := runFixMultiPreviewTest(t, root, t.Context(), "VB080", paths...)
		if err != nil || env.Status != output.StatusOK {
			t.Fatalf("preview=%+v envelope=%+v error=%v", result, env, err)
		}
		assertFixMultiPreviewCounts(t, result, 2, 2, 2)
		for _, path := range paths {
			if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || !bytes.Equal(got, source) {
				t.Fatalf("read-only preview modified %s: %q error=%v", path, got, readErr)
			}
		}
	})

	t.Run("cancelled before result publication", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "")
		paths := []string{"src/modules/First.bas", "src/modules/Second.bas"}
		source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
		for _, path := range paths {
			writeEncodingFile(t, root, path, source)
		}
		ctx, cancel := context.WithCancel(t.Context())
		cancel()
		_, env, err := runFixMultiPreviewTest(t, root, ctx, "VB080", paths...)
		if err == nil || env.Error == nil || env.Error.Code != "fix_cancelled" || env.Output != nil {
			t.Fatalf("envelope=%+v error=%v, want cancellation without partial preview", env, err)
		}
		for _, path := range paths {
			if got, readErr := os.ReadFile(filepath.Join(root, filepath.FromSlash(path))); readErr != nil || !bytes.Equal(got, source) {
				t.Fatalf("cancelled preview modified %s: %q error=%v", path, got, readErr)
			}
		}
	})
}
