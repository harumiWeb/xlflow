package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"golang.org/x/text/encoding/japanese"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/vba/sourceencoding"
)

func TestEncodingCheckCommandReportsAllViolationsInStableJSON(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)
	writeEncodingFile(t, root, "src/modules/00-valid.bas", []byte("Option Explicit\r\n"))
	writeEncodingFile(t, root, "src/modules/10-bom.cls", []byte{0xef, 0xbb, 0xbf, 'S'})
	writeEncodingFile(t, root, "src/forms/20-invalid.frm", []byte("Begin\n\xff\nEnd\n"))
	writeEncodingFile(t, root, "tests/30-valid.cls", []byte("Sub Test()\nEnd Sub\n"))
	writeEncodingFile(t, root, "src/forms/ignored.frx", []byte{0xff})

	var stdout bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "encoding", "check"})

	err := cmd.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("encoding check error = %v, exit = %d", err, output.ExitCode(err))
	}
	var got struct {
		Status string `json:"status"`
		Source struct {
			Expected string `json:"expected"`
			Files    []struct {
				Path   string `json:"path"`
				Status string `json:"status"`
			} `json:"files"`
			Summary sourceencoding.Summary `json:"summary"`
		} `json:"source"`
		Error *struct {
			Code        string         `json:"code"`
			Source      string         `json:"source"`
			Details     map[string]any `json:"details"`
			Suggestions []string       `json:"suggestions"`
		} `json:"error"`
	}
	if decodeErr := json.Unmarshal(stdout.Bytes(), &got); decodeErr != nil {
		t.Fatalf("decode encoding check JSON: %v\n%s", decodeErr, stdout.String())
	}
	if got.Status != output.StatusFailed || got.Source.Expected != sourceencoding.UTF8 || got.Error == nil {
		t.Fatalf("envelope = %#v", got)
	}
	if got.Error.Code != "source_encoding_invalid" || got.Error.Source != "src/forms/20-invalid.frm" {
		t.Fatalf("error = %#v", got.Error)
	}
	if len(got.Error.Suggestions) != 2 || got.Error.Suggestions[0] != "encoding check" || got.Error.Suggestions[1] != "save the file as UTF-8 without BOM" {
		t.Fatalf("suggestions = %#v", got.Error.Suggestions)
	}
	wantPaths := []string{"src/forms/20-invalid.frm", "src/modules/00-valid.bas", "src/modules/10-bom.cls", "tests/30-valid.cls"}
	if len(got.Source.Files) != len(wantPaths) {
		t.Fatalf("files = %#v", got.Source.Files)
	}
	for i, want := range wantPaths {
		if got.Source.Files[i].Path != want {
			t.Fatalf("files[%d] = %#v, want path %q", i, got.Source.Files[i], want)
		}
	}
	if got.Source.Summary.Total != 4 || got.Source.Summary.Valid != 2 || got.Source.Summary.Invalid != 2 {
		t.Fatalf("summary = %+v", got.Source.Summary)
	}
	if _, ok := got.Error.Details["files"]; !ok {
		t.Fatalf("multiple violations were not included in error.details: %#v", got.Error.Details)
	}
}

func TestEncodingCheckScopeFailureUsesConfigExitCode(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)

	var stdout bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "encoding", "check", "src/modules/ignored.frx"})
	err := cmd.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("encoding check error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout.String())
	}
	if got := errorCodeFromJSON(t, stdout.String()); got != "encoding_args_invalid" {
		t.Fatalf("error code = %q, output=%s", got, stdout.String())
	}
}

func TestEncodingBOMErrorDoesNotSuggestCP932Conversion(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)
	writeEncodingFile(t, root, "src/modules/Bom.bas", []byte{0xef, 0xbb, 0xbf, 'x'})

	var stdout bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "analyze"})
	err := cmd.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("analyze error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout.String())
	}
	var got struct {
		Error struct {
			Code        string   `json:"code"`
			Suggestions []string `json:"suggestions"`
		} `json:"error"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("decode BOM JSON: %v\n%s", err, stdout.String())
	}
	if got.Error.Code != "source_encoding_invalid" || len(got.Error.Suggestions) != 2 || got.Error.Suggestions[1] == "encoding convert --from cp932" {
		t.Fatalf("BOM error = %#v", got.Error)
	}
}

func TestEncodingConvertCommandConvertsCP932AndLeavesUTF8ByteIdentical(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)
	originalText := "' 日本語コメント\r\nvalue = \"代入\"\r\n"
	original, err := japanese.ShiftJIS.NewEncoder().Bytes([]byte(originalText))
	if err != nil {
		t.Fatal(err)
	}
	writeEncodingFile(t, root, "src/modules/Japanese.cls", original)
	unchanged := []byte("Sub Run()\r\nEnd Sub\r\n")
	writeEncodingFile(t, root, "src/modules/Keep.bas", unchanged)

	var stdout bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "encoding", "convert", "--from", "cp932"})
	if err := cmd.Execute(); err != nil {
		t.Fatalf("encoding convert error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout.String())
	}
	converted, err := os.ReadFile(filepath.Join(root, "src", "modules", "Japanese.cls"))
	if err != nil {
		t.Fatal(err)
	}
	if string(converted) != originalText {
		t.Fatalf("converted = %q, want %q", converted, originalText)
	}
	kept, err := os.ReadFile(filepath.Join(root, "src", "modules", "Keep.bas"))
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(kept, unchanged) {
		t.Fatalf("UTF-8 file changed from %q to %q", unchanged, kept)
	}
	var got struct {
		Source struct {
			From    string                 `json:"from"`
			To      string                 `json:"to"`
			Summary sourceencoding.Summary `json:"summary"`
		} `json:"source"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("decode convert JSON: %v\n%s", err, stdout.String())
	}
	if got.Source.From != sourceencoding.CP932 || got.Source.To != sourceencoding.UTF8 || got.Source.Summary.Converted != 1 || got.Source.Summary.Unchanged != 1 {
		t.Fatalf("conversion source = %+v", got.Source)
	}
}

func TestEncodingConvertFailureDoesNotSuggestRetryingConversion(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)
	writeEncodingFile(t, root, "src/modules/Bad.cls", []byte{0x81})

	var stdout bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "encoding", "convert", "--from", "cp932"})
	err := cmd.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("convert error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout.String())
	}
	var got struct {
		Error struct {
			Code        string   `json:"code"`
			Suggestions []string `json:"suggestions"`
		} `json:"error"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("decode conversion failure JSON: %v\n%s", err, stdout.String())
	}
	if got.Error.Code != "source_encoding_invalid" || len(got.Error.Suggestions) != 1 || got.Error.Suggestions[0] != "encoding check" {
		t.Fatalf("conversion failure = %#v", got.Error)
	}
}

func TestEncodingConvertRequiresFrom(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{cwd: t.TempDir(), stdout: &stdout, stderr: &bytes.Buffer{}}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "encoding", "convert"})
	err := cmd.Execute()
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("error = %v, exit = %d", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout.String()); got != "encoding_args_invalid" {
		t.Fatalf("error code = %q", got)
	}
}

func TestCheckEncodingPreflightStopsBeforeDoctor(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)
	writeEncodingFile(t, root, "src/modules/Bad.bas", []byte("Sub Bad()\n\xff\nEnd Sub\n"))
	called := false
	var stdout bytes.Buffer
	a := &app{
		cwd:    root,
		stdout: &stdout,
		stderr: &bytes.Buffer{},
		checkDoctor: func(config.Config, excel.CommandOptions) (output.Envelope, int, error) {
			called = true
			return output.Envelope{}, output.ExitSuccess, nil
		},
	}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "check"})
	err := cmd.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("check error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout.String())
	}
	if called {
		t.Fatal("doctor runner was called after source encoding validation failed")
	}
	if got := errorCodeFromJSON(t, stdout.String()); got != "source_encoding_invalid" {
		t.Fatalf("error code = %q, output=%s", got, stdout.String())
	}
}

func TestAnalyzeRejectsCP932BeforeRuleConfigurationCanMatter(t *testing.T) {
	for _, disabled := range []bool{false, true} {
		t.Run(map[bool]string{false: "VBA223-enabled", true: "VBA223-disabled"}[disabled], func(t *testing.T) {
			root := t.TempDir()
			writeEncodingConfig(t, root)
			if disabled {
				configText := `[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"

[analyze]
disabled_rules = ["VBA223"]
`
				writeEncodingFile(t, root, config.FileName, []byte(configText))
			}
			encoded, err := japanese.ShiftJIS.NewEncoder().Bytes([]byte("' 日本語コメント\r\nvalue = \"代入\"\r\n"))
			if err != nil {
				t.Fatal(err)
			}
			writeEncodingFile(t, root, "src/classes/Japanese.cls", encoded)

			var stdout bytes.Buffer
			a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
			cmd := a.rootCommand()
			cmd.SetArgs([]string{"--json", "analyze"})
			err = cmd.Execute()
			if err == nil || output.ExitCode(err) != output.ExitValidation {
				t.Fatalf("analyze error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout.String())
			}
			if got := errorCodeFromJSON(t, stdout.String()); got != "source_encoding_invalid" {
				t.Fatalf("error code = %q, output=%s", got, stdout.String())
			}
		})
	}
}

func TestLintAndPushRejectInvalidSourceBeforeMutation(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)
	badPath := filepath.Join(root, "src", "modules", "Bad.bas")
	bad := []byte("Sub Bad()\n  value = \xff\nEnd Sub\n")
	writeEncodingFile(t, root, "src/modules/Bad.bas", bad)

	var lintOutput bytes.Buffer
	lintApp := &app{cwd: root, stdout: &lintOutput, stderr: &bytes.Buffer{}}
	lintCmd := lintApp.rootCommand()
	lintCmd.SetArgs([]string{"--json", "lint"})
	if err := lintCmd.Execute(); err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("lint error = %v, exit = %d\n%s", err, output.ExitCode(err), lintOutput.String())
	}
	if got := errorCodeFromJSON(t, lintOutput.String()); got != "source_encoding_invalid" {
		t.Fatalf("lint error code = %q, output=%s", got, lintOutput.String())
	}

	cfg, err := config.Load(root)
	if err != nil {
		t.Fatal(err)
	}
	var pushOutput bytes.Buffer
	pushApp := &app{cwd: root, stdout: &pushOutput, stderr: &bytes.Buffer{}, json: true}
	_, _, err = pushApp.pushSource(context.Background(), "push", cfg, excel.PushOptions{}, "")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("push preflight error = %v, exit = %d\n%s", err, output.ExitCode(err), pushOutput.String())
	}
	if got := errorCodeFromJSON(t, pushOutput.String()); got != "source_encoding_invalid" {
		t.Fatalf("push error code = %q, output=%s", got, pushOutput.String())
	}
	after, err := os.ReadFile(badPath)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(after, bad) {
		t.Fatalf("invalid source changed during rejected commands: %x -> %x", bad, after)
	}
}

func TestDirectLintAndAnalyzeRejectSidecarFormEncoding(t *testing.T) {
	for _, command := range []string{"lint", "analyze"} {
		t.Run(command, func(t *testing.T) {
			root := t.TempDir()
			writeEncodingConfig(t, root)
			writeEncodingFile(t, root, "src/forms/CustomerForm.frm", []byte{0xef, 0xbb, 0xbf, 'x'})
			writeEncodingFile(t, root, "src/forms/code/CustomerForm.bas", []byte("Option Explicit\n"))

			var stdout bytes.Buffer
			a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
			cmd := a.rootCommand()
			cmd.SetArgs([]string{"--json", command})
			err := cmd.Execute()
			if err == nil || output.ExitCode(err) != output.ExitValidation {
				t.Fatalf("%s error = %v, exit = %d\n%s", command, err, output.ExitCode(err), stdout.String())
			}
			if got := errorCodeFromJSON(t, stdout.String()); got != "source_encoding_invalid" {
				t.Fatalf("%s error code = %q, output=%s", command, got, stdout.String())
			}
		})
	}
}

func TestFormMigrateSidecarRejectsInvalidFRMBeforeExcelInspection(t *testing.T) {
	root := t.TempDir()
	writeEncodingConfig(t, root)
	writeEncodingFile(t, root, config.FileName, []byte(`[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"

[userform]
code_source = "frm"
`))
	writeEncodingFile(t, root, "src/forms/CustomerForm.frm", []byte{0xef, 0xbb, 0xbf, 'x'})

	var stdout bytes.Buffer
	a := &app{cwd: root, stdout: &stdout, stderr: &bytes.Buffer{}}
	cmd := a.rootCommand()
	cmd.SetArgs([]string{"--json", "form", "migrate", "sidecar", "CustomerForm"})
	err := cmd.Execute()
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("form migration error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout.String())
	}
	if got := errorCodeFromJSON(t, stdout.String()); got != "source_encoding_invalid" {
		t.Fatalf("error code = %q, output=%s", got, stdout.String())
	}
}

func writeEncodingConfig(t *testing.T, root string) {
	t.Helper()
	configText := `[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"
`
	writeEncodingFile(t, root, config.FileName, []byte(configText))
}

func writeEncodingFile(t *testing.T, root, relative string, body []byte) {
	t.Helper()
	path := filepath.Join(root, filepath.FromSlash(relative))
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, body, 0o644); err != nil {
		t.Fatal(err)
	}
}
