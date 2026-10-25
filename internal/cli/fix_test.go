package cli

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

func runFixTest(t *testing.T, root string, ctx context.Context, args ...string) (fixResult, output.Envelope, error) {
	t.Helper()
	var stdout, stderr bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &stderr}
	cmd := a.rootCommand()
	cmd.SetArgs(append([]string{"--json", "fix"}, args...))
	err := cmd.ExecuteContext(ctx)
	var env output.Envelope
	if decodeErr := json.Unmarshal(stdout.Bytes(), &env); decodeErr != nil {
		t.Fatalf("decode output: %v; stdout=%s stderr=%s error=%v", decodeErr, stdout.String(), stderr.String(), err)
	}
	var result fixResult
	if env.Output != nil {
		data, marshalErr := json.Marshal(env.Output)
		if marshalErr != nil {
			t.Fatal(marshalErr)
		}
		if decodeErr := json.Unmarshal(data, &result); decodeErr != nil {
			t.Fatal(decodeErr)
		}
		if strings.Contains(string(data), `"Range"`) || strings.Contains(string(data), `"Character"`) {
			t.Fatalf("unprojected protocol-neutral output: %s", data)
		}
	}
	return result, env, err
}

func writeFixConfig(t *testing.T, root, extra string) {
	t.Helper()
	writeEncodingFile(t, root, config.FileName, []byte(`[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"
[lint]
require_option_explicit = false
detect_let_assignment = true
detect_redundant_step_one = true
detect_redundant_option_base_zero = true
`+extra))
}

func TestFixPreviewApplyRules(t *testing.T) {
	for _, eol := range []string{"\n", "\r\n"} {
		for _, tc := range []struct{ rule, source, want, extra string }{
			{"VB044", "Sub Alpha()\n Const PROCEDURE_NAME As String = \"旧\" ' 日本語😀\nEnd Sub\n", "Sub Alpha()\n Const PROCEDURE_NAME As String = \"Alpha\" ' 日本語😀\nEnd Sub\n", "[lint.procedure_name_constant]\nenabled = true\nconstant_name = \"PROCEDURE_NAME\"\n"},
			{"VB080", "Sub Alpha()\n Dim x As Long\n Let x = 1 ' 日本語😀\n Let x = 2\nEnd Sub\n", "Sub Alpha()\n Dim x As Long\n x = 1 ' 日本語😀\n x = 2\nEnd Sub\n", ""},
			{"VB084", "Sub Alpha()\n Dim i As Long\n For i = 1 To 3 Step 1 ' 日本語😀\n Next i\nEnd Sub\n", "Sub Alpha()\n Dim i As Long\n For i = 1 To 3 ' 日本語😀\n Next i\nEnd Sub\n", ""},
			{"VB086", "Option Base 0 ' 日本語😀\nSub Alpha()\nEnd Sub\n", " ' 日本語😀\nSub Alpha()\nEnd Sub\n", ""},
		} {
			t.Run(tc.rule+fmt.Sprintf("-%q", eol), func(t *testing.T) {
				root := t.TempDir()
				writeFixConfig(t, root, tc.extra)
				path := "src/modules/Main.bas"
				source, want := strings.ReplaceAll(tc.source, "\n", eol), strings.ReplaceAll(tc.want, "\n", eol)
				writeEncodingFile(t, root, path, []byte(source))
				// An unrelated invalid source must not cause a whole-project preflight.
				writeEncodingFile(t, root, "src/modules/Unrelated.bas", []byte{0xff})
				preview, env, err := runFixTest(t, root, t.Context(), path, "--rule", tc.rule)
				if err != nil || env.Status != output.StatusOK || !preview.Changed || preview.Applied || preview.Mode != "preview" || len(preview.Edits) == 0 || len(preview.Diagnostics) == 0 {
					t.Fatalf("preview=%+v envelope=%+v error=%v", preview, env, err)
				}
				if got, _ := os.ReadFile(filepath.Join(root, path)); string(got) != source {
					t.Fatalf("preview wrote source: %q", got)
				}
				if preview.SourceSHA256 != fmt.Sprintf("%x", sha256.Sum256([]byte(source))) || preview.ResultSHA256 != fmt.Sprintf("%x", sha256.Sum256([]byte(want))) || !strings.Contains(preview.Diff, "--- ") {
					t.Fatalf("incorrect hashes/diff: %+v", preview)
				}
				for _, diagnostic := range preview.Diagnostics {
					if diagnostic.Code != tc.rule || diagnostic.Message == "" {
						t.Fatalf("incorrect diagnostics: %+v", diagnostic)
					}
				}
				applied, _, err := runFixTest(t, root, t.Context(), path, "--rule", tc.rule, "--apply")
				if err != nil || !applied.Applied || applied.Mode != "apply" || applied.ResultSHA256 != preview.ResultSHA256 || applied.Diff != preview.Diff {
					t.Fatalf("applied=%+v error=%v", applied, err)
				}
				if got, _ := os.ReadFile(filepath.Join(root, path)); string(got) != want {
					t.Fatalf("apply source=%q want=%q", got, want)
				}
				again, _, err := runFixTest(t, root, t.Context(), path, "--rule", tc.rule, "--apply")
				if err != nil || again.Applied || again.Changed || len(again.Edits) != 0 || again.Edits == nil || again.Diagnostics == nil || again.Diff != "" {
					t.Fatalf("repeat=%+v error=%v", again, err)
				}
			})
		}
	}
}

func TestFixWithholdsDisabledSuppressedUnsafe(t *testing.T) {
	for _, tc := range []struct{ name, source, extra string }{
		{"suppressed", "Sub A()\n Let x = 1 ' xlflow:disable-line VB080\nEnd Sub\n", ""},
		{"disabled rule", "Sub A()\n Let x = 1\nEnd Sub\n", "disabled_rules = [\"VB080\"]\n"},
		{"unsafe", "Public Property Let Name(ByVal value As String)\nEnd Property\n", ""},
	} {
		t.Run(tc.name, func(t *testing.T) {
			root := t.TempDir()
			writeFixConfig(t, root, tc.extra)
			path := "src/classes/Main.cls"
			writeEncodingFile(t, root, path, []byte(tc.source))
			result, _, err := runFixTest(t, root, t.Context(), path, "--rule", "VB080", "--apply")
			if err != nil || result.Changed || result.Applied || len(result.Edits) != 0 {
				t.Fatalf("result=%+v error=%v", result, err)
			}
			if got, _ := os.ReadFile(filepath.Join(root, path)); string(got) != tc.source {
				t.Fatalf("source changed: %q", got)
			}
		})
	}
	root := t.TempDir()
	writeEncodingConfig(t, root)
	writeEncodingFile(t, root, "tests/Disabled.bas", []byte("Option Base 0\n"))
	result, _, err := runFixTest(t, root, t.Context(), "tests/Disabled.bas", "--rule", "VB086", "--apply")
	if err != nil || result.Changed {
		t.Fatalf("default-disabled rule was enabled: %+v error=%v", result, err)
	}
}

func TestFixExportedClassMetadataAndUTF16(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	path := "src/classes/Example.cls"
	source := "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Example\"\r\nOption Explicit\r\nSub Demo()\r\n Debug.Print \"😀\": Let value = 1 ' 日本語\r\nEnd Sub\r\n"
	writeEncodingFile(t, root, path, []byte(source))
	result, _, err := runFixTest(t, root, t.Context(), path, "--rule", "VB080", "--apply")
	if err != nil || !result.Applied || len(result.Edits) != 1 {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	if result.Edits[0].Range.Start.Line != 7 || result.Edits[0].Range.Start.Character != 19 {
		t.Fatalf("UTF-16 range = %+v", result.Edits[0].Range)
	}
	want := strings.Replace(source, " Let value", " value", 1)
	if got, readErr := os.ReadFile(filepath.Join(root, path)); readErr != nil || string(got) != want {
		t.Fatalf("class metadata/bytes changed: %q error=%v", got, readErr)
	}
}

func TestFixArgumentsAndOwnership(t *testing.T) {
	root := t.TempDir()
	writeFixConfig(t, root, "")
	if err := os.MkdirAll(filepath.Join(root, "src/modules/Directory.bas"), 0o755); err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{"src/modules/Main.bas", "outside.bas", "src/forms/assets/Resource.bas", "src/forms/specs/Resource.cls", "src/forms/code/Form.bas", "src/forms/Form.frm"} {
		writeEncodingFile(t, root, path, []byte("Option Base 0\n"))
	}
	for _, args := range [][]string{
		{}, {"src/modules/Main.bas"}, {"src/modules/Main.bas", "--rule", "VB001"}, {"src/modules/Main.bas", "--rule", "VB020"},
		{"src/modules/Main.bas", "--rule", "VBA275"},
		{"src/modules/Main.bas", "--rule", "VB080,VB084"}, {"src/modules/Main.bas", "--rule", "VB080", "--rule", "VB084"},
		{"src/modules/Main.bas", "outside.bas", "--rule", "VB086"},
		{"outside.bas", "--rule", "VB086"}, {"src/forms/assets/Resource.bas", "--rule", "VB086"}, {"src/forms/specs/Resource.cls", "--rule", "VB086"}, {"src/forms/Form.frm", "--rule", "VB086"},
		{"src/modules/Directory.bas", "--rule", "VB080"},
	} {
		t.Run(strings.Join(args, " "), func(t *testing.T) {
			_, env, err := runFixTest(t, root, t.Context(), args...)
			if err == nil || env.Error == nil || env.Error.Code != "fix_args_invalid" {
				t.Fatalf("envelope=%+v error=%v", env, err)
			}
		})
	}
	writeFixConfig(t, root, "[userform]\ncode_source = \"frm\"\n")
	_, inactive, inactiveErr := runFixTest(t, root, t.Context(), "src/forms/code/Form.bas", "--rule", "VB086")
	if inactiveErr == nil || inactive.Error == nil || inactive.Error.Code != "fix_args_invalid" {
		t.Fatalf("inactive sidecar: %+v error=%v", inactive, inactiveErr)
	}
	// Active form code and document/test class sources are source-only inputs.
	writeFixConfig(t, root, "[userform]\ncode_source = \"sidecar\"\n")
	for _, path := range []string{"src/forms/code/Form.bas", "src/workbook/ThisWorkbook.cls", "tests/Case.cls"} {
		writeEncodingFile(t, root, path, []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
		result, _, err := runFixTest(t, root, t.Context(), path, "--rule", "vb080")
		if err != nil || !result.Changed || result.Rule != "VB080" {
			t.Fatalf("path=%s result=%+v error=%v", path, result, err)
		}
	}
}

func TestFixFailureContracts(t *testing.T) {
	t.Run("missing config", func(t *testing.T) {
		_, env, err := runFixTest(t, t.TempDir(), t.Context(), "src/modules/Main.bas", "--rule", "VB080")
		if err == nil || env.Error == nil || env.Error.Code != "config_error" {
			t.Fatalf("envelope=%+v error=%v", env, err)
		}
	})
	for _, source := range [][]byte{{0xef, 0xbb, 0xbf, 'x'}, {0x81}, {0xff, 0xfe, 'x', 0}} {
		root := t.TempDir()
		writeFixConfig(t, root, "")
		path := "src/modules/Main.bas"
		writeEncodingFile(t, root, path, source)
		_, env, err := runFixTest(t, root, t.Context(), path, "--rule", "VB080", "--apply")
		if err == nil || env.Error == nil || env.Error.Code != "source_encoding_invalid" {
			t.Fatalf("envelope=%+v error=%v", env, err)
		}
		if got, _ := os.ReadFile(filepath.Join(root, path)); !bytes.Equal(got, source) {
			t.Fatalf("invalid encoding overwritten: %x", got)
		}
	}
	t.Run("readonly", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "")
		path := "src/modules/Main.bas"
		source := []byte("Sub A()\n Let x = 1\nEnd Sub\n")
		writeEncodingFile(t, root, path, source)
		abs := filepath.Join(root, path)
		if err := os.Chmod(abs, 0o444); err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() { _ = os.Chmod(abs, 0o644) })
		if result, _, err := runFixTest(t, root, t.Context(), path, "--rule", "VB080"); err != nil || !result.Changed {
			t.Fatalf("readonly preview=%+v error=%v", result, err)
		}
		_, env, err := runFixTest(t, root, t.Context(), path, "--rule", "VB080", "--apply")
		if err == nil || env.Error == nil || env.Error.Code != "fix_failed" {
			t.Fatalf("envelope=%+v error=%v", env, err)
		}
		if got, _ := os.ReadFile(abs); !bytes.Equal(got, source) {
			t.Fatalf("readonly source changed: %q", got)
		}
	})
	t.Run("cancelled", func(t *testing.T) {
		root := t.TempDir()
		writeFixConfig(t, root, "")
		writeEncodingFile(t, root, "src/modules/Main.bas", []byte("Sub A()\n Let x = 1\nEnd Sub\n"))
		ctx, cancel := context.WithCancel(t.Context())
		cancel()
		_, env, err := runFixTest(t, root, ctx, "src/modules/Main.bas", "--rule", "VB080", "--apply")
		if err == nil || env.Error == nil || env.Error.Code != "fix_cancelled" {
			t.Fatalf("envelope=%+v error=%v", env, err)
		}
	})
}
