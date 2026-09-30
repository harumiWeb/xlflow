package cli

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"io"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"

	"golang.org/x/text/encoding/japanese"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
)

func TestRootCommandIncludesPackCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"pack"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "pack" {
		t.Fatalf("expected pack command, got %#v", cmd)
	}
	for _, name := range []string{"out", "template", "blank"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected pack command to define --%s", name)
		}
	}
	experimental := cmd.Flags().Lookup("experimental")
	if experimental == nil || !experimental.Hidden || experimental.Deprecated == "" {
		t.Fatalf("deprecated --experimental compatibility flag = %#v", experimental)
	}
}

func TestPackCommandBlankCreatesWorkbookWithoutTemplate(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackSourceTree(t, dir, false)

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--blank", "--out", "dist/Fresh.xlsm")
	if err != nil {
		t.Fatalf("pack --blank: %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	packOutput := cliObjectMap(env.Pack)
	if packOutput["base"] != "blank" {
		t.Fatalf("pack output = %+v", packOutput)
	}
	if _, ok := packOutput["template"]; ok {
		t.Fatalf("blank output must omit template: %+v", packOutput)
	}
	body, err := os.ReadFile(filepath.Join(dir, "dist", "Fresh.xlsm"))
	if err != nil {
		t.Fatal(err)
	}
	if _, err := vbaproject.Read(zipEntryBytes(t, body, "xl/vbaProject.bin")); err != nil {
		t.Fatalf("fresh vbaProject.bin: %v", err)
	}
}

func TestPackCommandRejectsBlankWithTemplate(t *testing.T) {
	dir := t.TempDir()
	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--blank", "--template", "base.xlsm", "--out", "dist/Fresh.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_args_invalid" {
		t.Fatalf("error code = %q", got)
	}
}

func TestPackCommandRejectsBlankWithEmptyTemplateFlag(t *testing.T) {
	dir := t.TempDir()
	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--blank", "--template", "", "--out", "dist/Fresh.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_args_invalid" {
		t.Fatalf("error code = %q, want pack_args_invalid\n%s", got, stdout)
	}
}

func TestPackCommandBlankIgnoresUnrelatedWorkbookLock(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)
	writePackSourceTree(t, dir, false)
	// Blank mode never reads the configured workbook, so a stale Excel lock
	// beside it must not block packing an unrelated destination.
	if err := os.WriteFile(filepath.Join(dir, "build", "~$Book.xlsm"), []byte("lock"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--blank", "--out", "dist/Fresh.xlsm")
	if err != nil {
		t.Fatalf("pack --blank: %v\n%s", err, stdout)
	}
	if _, err := os.Stat(filepath.Join(dir, "dist", "Fresh.xlsm")); err != nil {
		t.Fatalf("expected blank output: %v", err)
	}
}

func TestPackCommandExplicitTemplateIgnoresUnrelatedWorkbookLock(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)
	writePackSourceTree(t, dir, false)
	writePackSourceModule(t, dir, filepath.Join("build", "Template.xlsm"), buildPackWorkbookFixture(t, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin")))
	// With an explicit --template the configured workbook is never read or
	// replaced, so its lock file must not block packing another destination.
	if err := os.WriteFile(filepath.Join(dir, "build", "~$Book.xlsm"), []byte("lock"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--template", "build/Template.xlsm", "--out", "dist/Fresh.xlsm")
	if err != nil {
		t.Fatalf("pack --template: %v\n%s", err, stdout)
	}
	if _, err := os.Stat(filepath.Join(dir, "dist", "Fresh.xlsm")); err != nil {
		t.Fatalf("expected template output: %v", err)
	}
}

func TestPackCommandBlankRejectsUserFormArtifacts(t *testing.T) {
	cases := []struct {
		name string
		path string
	}{
		{name: "frm", path: filepath.Join("src", "forms", "Login.frm")},
		{name: "orphan frx", path: filepath.Join("src", "forms", "Login.frx")},
		{name: "orphan sidecar code", path: filepath.Join("src", "forms", "code", "Login.bas")},
		{name: "orphan form spec", path: filepath.Join("src", "forms", "specs", "Login.yaml")},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			writePackConfig(t, dir)
			writePackSourceTree(t, dir, false)
			writePackSourceModule(t, dir, tc.path, []byte("artifact"))

			stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--blank", "--out", "dist/Fresh.xlsm")
			if err == nil || output.ExitCode(err) != output.ExitValidation {
				t.Fatalf("err=%v exit=%d, want validation failure", err, output.ExitCode(err))
			}
			if got := errorCodeFromJSON(t, stdout); got != "pack_blank_userform_unsupported" {
				t.Fatalf("error code = %q, want pack_blank_userform_unsupported\n%s", got, stdout)
			}
		})
	}
}

func TestPackCommandValidationFailures(t *testing.T) {
	cases := []struct {
		name string
		args []string
		code int
		err  string
	}{
		{name: "missing out", args: []string{"--json", "pack"}, code: output.ExitConfig, err: "pack_args_invalid"},
		{name: "bad out extension", args: []string{"--json", "pack", "--out", "dist/Book.xlsx"}, code: output.ExitConfig, err: "pack_args_invalid"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			stdout, err := runPackCommandForTest(dir, tc.args...)
			if err == nil || output.ExitCode(err) != tc.code {
				t.Fatalf("err=%v exit=%d, want exit=%d", err, output.ExitCode(err), tc.code)
			}
			if got := errorCodeFromJSON(t, stdout); got != tc.err {
				t.Fatalf("error code = %q, want %q\n%s", got, tc.err, stdout)
			}
		})
	}
}

func TestPackCommandTemplateNotFound(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_template_not_found" {
		t.Fatalf("error code = %q, want pack_template_not_found\n%s", got, stdout)
	}
}

func TestPackCommandRejectsXlsbConfiguredWorkbook(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	cfg.Excel.Path = filepath.ToSlash(filepath.Join("build", "Model.xlsb"))
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "workbook_format_unsupported" {
		t.Fatalf("error code = %q, want workbook_format_unsupported\n%s", got, stdout)
	}
	var env output.Envelope
	if decodeErr := json.Unmarshal([]byte(stdout), &env); decodeErr != nil {
		t.Fatalf("json output should be valid: %v\n%s", decodeErr, stdout)
	}
	workbook := cliObjectMap(env.Workbook)
	if workbook["format"] != "xlsb" || workbook["capability"] != "pack" {
		t.Fatalf("workbook payload = %+v", workbook)
	}
}

func TestPackCommandInPlaceGuard(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "build/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_in_place_overwrite" {
		t.Fatalf("error code = %q, want pack_in_place_overwrite\n%s", got, stdout)
	}
}

func TestPackCommandActiveSessionLockFile(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)
	if err := os.WriteFile(filepath.Join(dir, "build", "~$Book.xlsm"), []byte("lock"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_active_session" {
		t.Fatalf("error code = %q, want pack_active_session\n%s", got, stdout)
	}
}

func TestPackCommandActiveSessionOutputLockFile(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)
	if err := os.MkdirAll(filepath.Join(dir, "dist"), 0o755); err != nil {
		t.Fatal(err)
	}
	// A lock file for the --out target means the destination workbook is open.
	if err := os.WriteFile(filepath.Join(dir, "dist", "~$Book.xlsm"), []byte("lock"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_active_session" {
		t.Fatalf("error code = %q, want pack_active_session\n%s", got, stdout)
	}
}

func TestPackCommandActiveSessionAliasLockFile(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)
	alias := filepath.Join(dir, "build", "Alias.xlsm")
	if err := os.Symlink(filepath.Join(dir, "build", "Book.xlsm"), alias); err != nil {
		t.Skipf("symlink creation unavailable: %v", err)
	}
	if err := os.WriteFile(filepath.Join(dir, "build", "~$Alias.xlsm"), []byte("lock"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--template", "build/Alias.xlsm", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_active_session" {
		t.Fatalf("error code = %q, want pack_active_session\n%s", got, stdout)
	}
}

func TestPackCommandActiveSessionMetadata(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)
	// An xlflow session recorded for the template, with no Office lock file present.
	meta, err := json.Marshal(map[string]any{
		"workbook_path": filepath.Join(dir, "build", "Book.xlsm"),
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), meta, 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_active_session" {
		t.Fatalf("error code = %q, want pack_active_session\n%s", got, stdout)
	}
}

func TestCollectPackSourceModulesIncludesForms(t *testing.T) {
	dir := t.TempDir()
	writePackSourceTree(t, dir, true)
	cfg := config.Default()
	cfg.Build.Exclude = []string{"src/modules/**", "src/forms/**"}

	sources, err := collectPackSourceModules(dir, cfg)
	if err != nil {
		t.Fatal(err)
	}
	counts := map[packpkg.ModuleType]int{}
	var sawForm bool
	for _, source := range sources {
		counts[source.Type]++
		if source.Name == "UserForm1" && source.Type == packpkg.ModuleTypeForm {
			sawForm = true
		}
	}
	if !sawForm {
		t.Fatal("UserForm1.frm should be collected as a ModuleTypeForm source")
	}
	if counts[packpkg.ModuleTypeStandard] != 1 || counts[packpkg.ModuleTypeClass] != 1 || counts[packpkg.ModuleTypeDocument] != 2 || counts[packpkg.ModuleTypeForm] != 1 {
		t.Fatalf("source counts = %v", counts)
	}
}

func TestCollectPackSourceModulesRejectsMissingConfiguredRoot(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	for _, sourceDir := range []string{cfg.Src.Classes, cfg.Src.Forms, cfg.Src.Workbook} {
		if err := os.MkdirAll(filepath.Join(dir, filepath.FromSlash(sourceDir)), 0o755); err != nil {
			t.Fatal(err)
		}
	}
	_, err := collectPackSourceModules(dir, cfg)
	if !errors.Is(err, packpkg.ErrAmbiguousLayout) || !strings.Contains(err.Error(), "read standard source root") {
		t.Fatalf("err = %v", err)
	}
}

func TestPackCommandEndToEndJSONAndWorkbook(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	writePackSourceTree(t, dir, false) // std/class/document only; the form path has its own E2E

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err != nil {
		t.Fatalf("pack command error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout)
	}
	outPath := filepath.Join(dir, "dist", "Book.xlsm")
	if _, err := os.Stat(outPath); err != nil {
		t.Fatal(err)
	}
	bin := zipEntryBytes(t, readFileForTest(t, outPath), "xl/vbaProject.bin")
	project, err := vbaproject.Read(bin)
	if err != nil {
		t.Fatalf("output vbaProject.bin is not readable: %v", err)
	}
	if got := moduleSourceForTest(project, "Module1"); !bytes.Contains([]byte(got), []byte(`Debug.Print "Hello, world"`)) {
		t.Fatalf("Module1 was not packed from source:\n%s", got)
	}

	var env struct {
		Status   string         `json:"status"`
		Command  string         `json:"command"`
		Output   map[string]any `json:"output"`
		Pack     map[string]any `json:"pack"`
		Warnings []struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"warnings"`
	}
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout)
	}
	if env.Status != output.StatusOK || env.Command != "pack" {
		t.Fatalf("unexpected envelope status/command: %#v", env)
	}
	if env.Output["path"] != "dist/Book.xlsm" || env.Output["format"] != "xlsm" || env.Output["created_parent_dirs"] != true {
		t.Fatalf("unexpected output payload: %#v", env.Output)
	}
	if env.Pack["backend"] != "pure-go" || env.Pack["vbe_validation"] != "not_performed" {
		t.Fatalf("unexpected pack payload: %#v", env.Pack)
	}
	if _, ok := env.Pack["experimental"]; ok {
		t.Fatalf("stable pack payload retained experimental marker: %#v", env.Pack)
	}
	modules, ok := env.Pack["modules"].(map[string]any)
	if !ok {
		t.Fatalf("pack.modules missing: %#v", env.Pack)
	}
	if modules["standard"] != float64(1) || modules["class"] != float64(1) || modules["document"] != float64(2) {
		t.Fatalf("unexpected module counts: %#v", modules)
	}
	if len(env.Warnings) != 1 || env.Warnings[0].Code != "vbe_validation_skipped" {
		t.Fatalf("unexpected warnings: %#v", env.Warnings)
	}

	if env.Output["publication"] != "atomic_create" || env.Output["replaced_existing"] != false {
		t.Fatalf("unexpected publication metadata for new output: %#v", env.Output)
	}
	cleanup, ok := env.Output["temporary_cleanup"].(map[string]any)
	if !ok || cleanup["status"] != "clean" {
		t.Fatalf("unexpected temporary_cleanup payload: %#v", env.Output["temporary_cleanup"])
	}
	assertNoPackStagingResidual(t, filepath.Join(dir, "dist"))

	// A second pack over the same destination must publish through the atomic
	// replace path and report it.
	stdout, err = runPackCommandForTest(dir, "--json", "pack", "--out", "dist/Book.xlsm")
	if err != nil {
		t.Fatalf("second pack command error = %v\n%s", err, stdout)
	}
	var second map[string]any
	if err := json.Unmarshal([]byte(stdout), &second); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout)
	}
	secondOutput, _ := second["output"].(map[string]any)
	if secondOutput["publication"] != "atomic_replace" || secondOutput["replaced_existing"] != true {
		t.Fatalf("unexpected publication metadata for replaced output: %#v", secondOutput)
	}
	assertNoPackStagingResidual(t, filepath.Join(dir, "dist"))
}

func TestPackCommandAcceptsUTF8JapaneseSource(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	writePackSourceTree(t, dir, false)

	modulePath := filepath.Join(dir, "src", "modules", "Module1.bas")
	module := append(readFileForTest(t, modulePath), []byte("' 日本語コメント\r\n")...)
	if err := os.WriteFile(modulePath, module, 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err != nil {
		t.Fatalf("pack command error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout)
	}
	bin := zipEntryBytes(t, readFileForTest(t, filepath.Join(dir, "dist", "Book.xlsm")), "xl/vbaProject.bin")
	project, err := vbaproject.Read(bin)
	if err != nil {
		t.Fatal(err)
	}
	if got := moduleSourceForTest(project, "Module1"); !strings.Contains(got, "日本語コメント") {
		t.Fatalf("packed Module1 lost UTF-8 Japanese source:\n%s", got)
	}
}

func TestPackCommandRunsSourceEncodingPreflightBeforeGeneration(t *testing.T) {
	cp932, err := japanese.ShiftJIS.NewEncoder().Bytes([]byte("Attribute VB_Name = \"Japanese\"\r\n' 日本語\r\n"))
	if err != nil {
		t.Fatal(err)
	}
	cases := []struct {
		name         string
		path         string
		status       string
		suggestion   string
		writeInvalid func(*testing.T, string)
	}{
		{
			name:       "CP932 module",
			path:       "src/modules/Japanese.bas",
			status:     "invalid_utf8",
			suggestion: "encoding convert --from cp932",
			writeInvalid: func(t *testing.T, dir string) {
				writePackSourceModule(t, dir, filepath.Join("src", "modules", "Japanese.bas"), cp932)
			},
		},
		{
			name:       "UTF-8 BOM form",
			path:       "src/forms/UserForm1.frm",
			status:     "utf8_bom",
			suggestion: "save the file as UTF-8 without BOM",
			writeInvalid: func(t *testing.T, dir string) {
				writePackSourceModule(t, dir, filepath.Join("src", "forms", "UserForm1.frm"), []byte{0xef, 0xbb, 0xbf, 'x'})
			},
		},
		{
			name:       "invalid UTF-8 form sidecar",
			path:       "src/forms/code/UserForm1.bas",
			status:     "invalid_utf8",
			suggestion: "encoding check",
			writeInvalid: func(t *testing.T, dir string) {
				writePackSourceModule(t, dir, filepath.Join("src", "forms", "UserForm1.frm"), []byte("VERSION 5.00\r\nBegin VB.UserForm UserForm1\r\nEnd\r\n"))
				writePackSourceModule(t, dir, filepath.Join("src", "forms", "code", "UserForm1.bas"), []byte("Private Sub Broken()\r\n\x81\r\nEnd Sub\r\n"))
			},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			writePackConfig(t, dir)
			// A protected template proves encoding validation wins before the
			// pack engine can parse or regenerate vbaProject.bin.
			writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p3_protected.bin"))
			tc.writeInvalid(t, dir)

			sentinel := []byte("previous-valid-output")
			outPath := filepath.Join(dir, "dist", "Book.xlsm")
			if err := os.MkdirAll(filepath.Dir(outPath), 0o755); err != nil {
				t.Fatal(err)
			}
			if err := os.WriteFile(outPath, sentinel, 0o644); err != nil {
				t.Fatal(err)
			}

			stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
			if err == nil || output.ExitCode(err) != output.ExitValidation {
				t.Fatalf("pack error = %v, exit = %d, want validation failure\n%s", err, output.ExitCode(err), stdout)
			}
			var env struct {
				Command string `json:"command"`
				Source  struct {
					Expected string `json:"expected"`
					Files    []struct {
						Path   string `json:"path"`
						Status string `json:"status"`
					} `json:"files"`
					Summary struct {
						Invalid int `json:"invalid"`
					} `json:"summary"`
				} `json:"source"`
				Error *struct {
					Code        string         `json:"code"`
					Source      string         `json:"source"`
					Details     map[string]any `json:"details"`
					Suggestions []string       `json:"suggestions"`
				} `json:"error"`
			}
			if err := json.Unmarshal([]byte(stdout), &env); err != nil {
				t.Fatalf("decode pack failure JSON: %v\n%s", err, stdout)
			}
			if env.Command != "pack" || env.Error == nil || env.Error.Code != "source_encoding_invalid" || env.Error.Source != tc.path {
				t.Fatalf("unexpected error envelope: %+v\n%s", env.Error, stdout)
			}
			if env.Error.Details["status"] != tc.status || !slices.Contains(env.Error.Suggestions, tc.suggestion) {
				t.Fatalf("unexpected encoding details/suggestions: %+v %+v", env.Error.Details, env.Error.Suggestions)
			}
			for _, key := range []string{"reason", "offset", "line", "byte_column"} {
				if _, ok := env.Error.Details[key]; !ok {
					t.Fatalf("encoding details missing %q: %+v", key, env.Error.Details)
				}
			}
			fileIndex := slices.IndexFunc(env.Source.Files, func(file struct {
				Path   string `json:"path"`
				Status string `json:"status"`
			}) bool {
				return file.Path == tc.path && file.Status == tc.status
			})
			if env.Source.Expected != "utf-8" || env.Source.Summary.Invalid == 0 || fileIndex < 0 {
				t.Fatalf("unexpected source payload: %+v", env.Source)
			}
			if got := readFileForTest(t, outPath); !bytes.Equal(got, sentinel) {
				t.Fatalf("pack modified existing output after encoding failure: %q", got)
			}
		})
	}
}

func TestPackCommandSupportsExternalAbsoluteSourceRoot(t *testing.T) {
	t.Run("valid UTF-8", func(t *testing.T) {
		dir := t.TempDir()
		external := t.TempDir()
		writePackConfig(t, dir)
		writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
		writePackSourceTree(t, dir, false)

		cfg, err := config.Load(dir)
		if err != nil {
			t.Fatal(err)
		}
		cfg.Src.Modules = filepath.Join(external, "modules")
		if err := os.Remove(filepath.Join(dir, config.FileName)); err != nil {
			t.Fatal(err)
		}
		if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
			t.Fatal(err)
		}
		writePackSourceModule(t, "", filepath.Join(cfg.Src.Modules, "Module1.bas"), readPackFixture(t, "testdata", "disk", "p1", "modules", "Module1.bas"))

		stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
		if err != nil {
			t.Fatalf("pack external source error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout)
		}
	})

	t.Run("invalid encoding", func(t *testing.T) {
		dir := t.TempDir()
		external := t.TempDir()
		writePackConfig(t, dir)
		writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p3_protected.bin"))

		cfg, err := config.Load(dir)
		if err != nil {
			t.Fatal(err)
		}
		cfg.Src.Modules = filepath.Join(external, "modules")
		if err := os.Remove(filepath.Join(dir, config.FileName)); err != nil {
			t.Fatal(err)
		}
		if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
			t.Fatal(err)
		}
		path := filepath.Join(cfg.Src.Modules, "Bad.bas")
		writePackSourceModule(t, "", path, []byte("Sub Broken()\r\n\xff\r\nEnd Sub\r\n"))

		stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
		if err == nil || output.ExitCode(err) != output.ExitValidation {
			t.Fatalf("pack external source error = %v, exit = %d, want validation failure\n%s", err, output.ExitCode(err), stdout)
		}
		var env output.Envelope
		if err := json.Unmarshal([]byte(stdout), &env); err != nil {
			t.Fatal(err)
		}
		if env.Error == nil || env.Error.Code != "source_encoding_invalid" || env.Error.Source != filepath.ToSlash(path) {
			t.Fatalf("unexpected external-root failure: %+v", env.Error)
		}
	})
}

func TestPackCommandValidatesBackslashConfiguredRoot(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p3_protected.bin"))

	cfg, err := config.Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	cfg.Src.Modules = `src\modules`
	if err := os.Remove(filepath.Join(dir, config.FileName)); err != nil {
		t.Fatal(err)
	}
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	writePackSourceModule(t, dir, filepath.Join("src", "modules", "Bad.bas"), []byte("Sub Broken()\r\n\xff\r\nEnd Sub\r\n"))

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("pack backslash source error = %v, exit = %d, want validation failure\n%s", err, output.ExitCode(err), stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	if env.Error == nil || env.Error.Code != "source_encoding_invalid" || env.Error.Source != "src/modules/Bad.bas" {
		t.Fatalf("unexpected backslash-root failure: %+v", env.Error)
	}
}

func TestPackCommandPublishesThroughOutputSymlink(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	writePackSourceTree(t, dir, false)

	referent := filepath.Join(dir, "shared", "Release.xlsm")
	if err := os.MkdirAll(filepath.Dir(referent), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(referent, []byte("previous-output"), 0o644); err != nil {
		t.Fatal(err)
	}
	alias := filepath.Join(dir, "dist", "Book.xlsm")
	if err := os.MkdirAll(filepath.Dir(alias), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(referent, alias); err != nil {
		t.Skipf("symlink creation unavailable: %v", err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err != nil {
		t.Fatalf("pack command error = %v\n%s", err, stdout)
	}
	info, err := os.Lstat(alias)
	if err != nil {
		t.Fatal(err)
	}
	if info.Mode()&os.ModeSymlink == 0 {
		t.Fatal("output symlink was replaced instead of publishing through it")
	}
	bin := zipEntryBytes(t, readFileForTest(t, referent), "xl/vbaProject.bin")
	if _, err := vbaproject.Read(bin); err != nil {
		t.Fatalf("referent vbaProject.bin is not readable: %v", err)
	}
	var env struct {
		Output map[string]any `json:"output"`
	}
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout)
	}
	if env.Output["path"] != "dist/Book.xlsm" || env.Output["publication"] != "atomic_replace" {
		t.Fatalf("unexpected output payload: %#v", env.Output)
	}
}

func TestValidatePackArtifactReadsVBAPayload(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "corrupt.xlsm")
	payload := []byte("unique-vba-payload-for-crc")
	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	header := &zip.FileHeader{Name: "xl/vbaProject.bin", Method: zip.Store}
	entry, err := writer.CreateHeader(header)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := entry.Write(payload); err != nil {
		t.Fatal(err)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	artifact := buf.Bytes()
	offset := bytes.Index(artifact, payload)
	if offset < 0 {
		t.Fatal("stored VBA payload not found in fixture")
	}
	artifact[offset] ^= 0xff
	if err := os.WriteFile(path, artifact, 0o644); err != nil {
		t.Fatal(err)
	}
	if err := validatePackArtifact(path); err == nil || !strings.Contains(err.Error(), "read staged xl/vbaProject.bin") {
		t.Fatalf("err = %v, want staged VBA payload read failure", err)
	}
}

func TestPackCommandEndToEndUpdatesFormCodeBehind(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p4_form.bin"))
	frm := "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1\r\n" +
		"   Caption = \"x\"\r\nEnd\r\n" +
		"Attribute VB_Name = \"UserForm1\"\r\n" +
		"Attribute VB_GlobalNameSpace = False\r\n" +
		"Attribute VB_Creatable = False\r\n" +
		"Attribute VB_PredeclaredId = True\r\n" +
		"Attribute VB_Exposed = False\r\n" +
		"Private Sub CommandButton1_Click()\r\n    Debug.Print \"CLI_PACKED_FORM\"\r\nEnd Sub\r\n"
	writePackSourceModule(t, dir, filepath.Join("src", "forms", "UserForm1.frm"), []byte(frm))

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err != nil {
		t.Fatalf("pack command error = %v, exit = %d\n%s", err, output.ExitCode(err), stdout)
	}
	bin := zipEntryBytes(t, readFileForTest(t, filepath.Join(dir, "dist", "Book.xlsm")), "xl/vbaProject.bin")
	project, err := vbaproject.Read(bin)
	if err != nil {
		t.Fatalf("output vbaProject.bin is not readable: %v", err)
	}
	if got := moduleSourceForTest(project, "UserForm1"); !bytes.Contains([]byte(got), []byte(`Debug.Print "CLI_PACKED_FORM"`)) {
		t.Fatalf("UserForm1 code-behind was not packed from source:\n%s", got)
	}
	var env struct {
		Pack map[string]any `json:"pack"`
	}
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout)
	}
	modules, _ := env.Pack["modules"].(map[string]any)
	if modules["form"] != float64(1) {
		t.Fatalf("expected form count 1, got %#v", modules)
	}
}

func TestPackCommandMapsProtectedProjectEngineError(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p3_protected.bin"))

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("err=%v exit=%d, want validation failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_protected_project" {
		t.Fatalf("error code = %q, want pack_protected_project\n%s", got, stdout)
	}
}

func TestPackCommandMapsSignedProjectEngineError(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	project, err := vbaproject.Read(readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	project.RawStreams["_VBA_PROJECT_CUR/VBAProjectSignature"] = []byte("signature sentinel")
	signed, err := vbaproject.Write(project)
	if err != nil {
		t.Fatal(err)
	}
	writePackTemplate(t, dir, signed)

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("err=%v exit=%d, want validation failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_signed_project" {
		t.Fatalf("error code = %q, want pack_signed_project\n%s", got, stdout)
	}
	if _, statErr := os.Stat(filepath.Join(dir, "dist", "Book.xlsm")); !errors.Is(statErr, os.ErrNotExist) {
		t.Fatalf("signed project rejection published an artifact, stat error = %v", statErr)
	}
}

func TestPackCommandAddsModuleAbsentFromTemplate(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	writePackSourceModule(t, dir, filepath.Join("src", "modules", "Module99.bas"), []byte("Attribute VB_Name = \"Module99\"\r\nOption Explicit\r\n"))

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err != nil {
		t.Fatalf("pack command error = %v\n%s", err, stdout)
	}
	bin := zipEntryBytes(t, readFileForTest(t, filepath.Join(dir, "dist", "Book.xlsm")), "xl/vbaProject.bin")
	project, err := vbaproject.Read(bin)
	if err != nil {
		t.Fatal(err)
	}
	if got := moduleSourceForTest(project, "Module99"); !strings.Contains(got, `Attribute VB_Name = "Module99"`) {
		t.Fatalf("Module99 was not packed: %q", got)
	}
	if got := moduleSourceForTest(project, "Module1"); got != "" {
		t.Fatalf("template-only Module1 survived source-authoritative pack: %q", got)
	}
}

func TestPackCommandRejectsAddedModuleNameMismatch(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	writePackSourceModule(t, dir, filepath.Join("src", "modules", "NewModule.bas"), []byte("Attribute VB_Name = \"OldName\"\r\nOption Explicit\r\n"))

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("err=%v exit=%d, want validation failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_ambiguous_layout" {
		t.Fatalf("error code = %q, want pack_ambiguous_layout\n%s", got, stdout)
	}
	if _, statErr := os.Stat(filepath.Join(dir, "dist", "Book.xlsm")); !errors.Is(statErr, os.ErrNotExist) {
		t.Fatalf("output artifact should not be published, stat error = %v", statErr)
	}
}

func TestPackCommandInPlaceGuardRejectsAliasedOutput(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)

	link := filepath.Join(dir, "build-alias")
	if err := os.Symlink(filepath.Join(dir, "build"), link); err != nil {
		t.Skipf("symlink creation unavailable: %v", err)
	}
	defer func() { _ = os.Remove(link) }()

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "build-alias/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_in_place_overwrite" {
		t.Fatalf("error code = %q, want pack_in_place_overwrite\n%s", got, stdout)
	}
}

func TestPackCommandPreservesExistingOutputOnEngineFailure(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p3_protected.bin"))
	sentinel := []byte("previous-valid-output")
	if err := os.MkdirAll(filepath.Join(dir, "dist"), 0o755); err != nil {
		t.Fatal(err)
	}
	outPath := filepath.Join(dir, "dist", "Book.xlsm")
	if err := os.WriteFile(outPath, sentinel, 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("err=%v exit=%d, want validation failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_protected_project" {
		t.Fatalf("error code = %q, want pack_protected_project\n%s", got, stdout)
	}
	body, readErr := os.ReadFile(outPath)
	if readErr != nil {
		t.Fatal(readErr)
	}
	if !bytes.Equal(body, sentinel) {
		t.Fatalf("existing output was modified after failed pack: %q", body)
	}
}

func TestPackCommandActiveSessionAlias(t *testing.T) {
	dir := t.TempDir()
	writePackProject(t, dir, false)

	// A session recorded through a symlinked project alias must still block pack.
	link := filepath.Join(dir, "alias-project")
	if err := os.Symlink(dir, link); err != nil {
		t.Skipf("symlink creation unavailable: %v", err)
	}
	defer func() { _ = os.Remove(link) }()
	meta, err := json.Marshal(map[string]any{
		"workbook_path": filepath.Join(link, "build", "Book.xlsm"),
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), meta, 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
	if err == nil || output.ExitCode(err) != output.ExitConfig {
		t.Fatalf("err=%v exit=%d, want config failure", err, output.ExitCode(err))
	}
	if got := errorCodeFromJSON(t, stdout); got != "pack_active_session" {
		t.Fatalf("error code = %q, want pack_active_session\n%s", got, stdout)
	}
}

func runPackCommandForTest(dir string, args ...string) (string, error) {
	var stdout bytes.Buffer
	a := &app{
		cwd:            dir,
		stdout:         &stdout,
		stderr:         &bytes.Buffer{},
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}
	root := a.rootCommand()
	root.SetArgs(args)
	err := root.Execute()
	return stdout.String(), err
}

func writePackProject(t *testing.T, dir string, withSources bool) {
	t.Helper()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	if withSources {
		writePackSourceTree(t, dir, true)
	}
}

func writePackConfig(t *testing.T, dir string) {
	t.Helper()
	cfg := config.Default()
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	for _, sourceDir := range []string{cfg.Src.Modules, cfg.Src.Classes, cfg.Src.Forms, cfg.Src.Workbook} {
		if err := os.MkdirAll(filepath.Join(dir, filepath.FromSlash(sourceDir)), 0o755); err != nil {
			t.Fatal(err)
		}
	}
}

func writePackSourceTree(t *testing.T, dir string, includeForm bool) {
	t.Helper()
	writePackSourceModule(t, dir, filepath.Join("src", "modules", "Module1.bas"), readPackFixture(t, "testdata", "disk", "p1", "modules", "Module1.bas"))
	writePackSourceModule(t, dir, filepath.Join("src", "classes", "Class1.cls"), readPackFixture(t, "testdata", "disk", "p1", "classes", "Class1.cls"))
	writePackSourceModule(t, dir, filepath.Join("src", "workbook", "Sheet1.bas"), readPackFixture(t, "testdata", "disk", "p1", "workbook", "Sheet1.bas"))
	writePackSourceModule(t, dir, filepath.Join("src", "workbook", "ThisWorkbook.bas"), readPackFixture(t, "testdata", "disk", "p1", "workbook", "ThisWorkbook.bas"))
	if includeForm {
		writePackSourceModule(t, dir, filepath.Join("src", "forms", "UserForm1.frm"), []byte("VERSION 5.00\r\nBegin VB.UserForm UserForm1\r\nEnd\r\n"))
	}
}

func writePackTemplate(t *testing.T, dir string, vbaProject []byte) {
	t.Helper()
	template := buildPackWorkbookFixture(t, vbaProject)
	path := filepath.Join(dir, "build", "Book.xlsm")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, template, 0o644); err != nil {
		t.Fatal(err)
	}
}

func writePackSourceModule(t *testing.T, dir string, path string, body []byte) {
	t.Helper()
	full := filepath.Join(dir, path)
	if err := os.MkdirAll(filepath.Dir(full), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(full, body, 0o644); err != nil {
		t.Fatal(err)
	}
}

func readPackFixture(t *testing.T, parts ...string) []byte {
	t.Helper()
	all := append([]string{"..", "pack"}, parts...)
	return readFileForTest(t, filepath.Join(all...))
}

func assertNoPackStagingResidual(t *testing.T, dir string) {
	t.Helper()
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatalf("read output dir: %v", err)
	}
	for _, entry := range entries {
		if strings.HasPrefix(entry.Name(), ".xlflow-publish-") {
			t.Fatalf("temporary publish artifact was not cleaned: %s", entry.Name())
		}
	}
}

func readFileForTest(t *testing.T, path string) []byte {
	t.Helper()
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return body
}

func buildPackWorkbookFixture(t *testing.T, vbaProject []byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	add := func(name string, method uint16, body []byte) {
		t.Helper()
		header := &zip.FileHeader{Name: name, Method: method}
		w, err := writer.CreateHeader(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(body); err != nil {
			t.Fatal(err)
		}
	}
	add("[Content_Types].xml", zip.Deflate, []byte(`<Types></Types>`))
	add("xl/workbook.xml", zip.Store, []byte(`<workbook/>`))
	add("xl/vbaProject.bin", zip.Deflate, vbaProject)
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func zipEntryBytes(t *testing.T, data []byte, name string) []byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, entry := range reader.File {
		if entry.Name != name {
			continue
		}
		rc, err := entry.Open()
		if err != nil {
			t.Fatal(err)
		}
		body, readErr := io.ReadAll(rc)
		closeErr := rc.Close()
		if readErr != nil {
			t.Fatal(readErr)
		}
		if closeErr != nil {
			t.Fatal(closeErr)
		}
		return body
	}
	t.Fatalf("zip entry %s not found", name)
	return nil
}

func moduleSourceForTest(project *vbaproject.Project, name string) string {
	for _, module := range project.Modules {
		if module.Name == name {
			return module.Source
		}
	}
	return ""
}

func errorCodeFromJSON(t *testing.T, stdout string) string {
	t.Helper()
	var env struct {
		Error *struct {
			Code string `json:"code"`
		} `json:"error"`
	}
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatalf("json output should be valid: %v\n%s", err, stdout)
	}
	if env.Error == nil {
		t.Fatalf("expected error payload in %s", stdout)
	}
	return env.Error.Code
}

// makePackSourceRoots creates the four configured source roots so shared
// discovery does not fail on a missing directory before exercising UserForm
// collection behavior.
func makePackSourceRoots(t *testing.T, root string, cfg config.Config) {
	t.Helper()
	for _, sourceDir := range []string{cfg.Src.Modules, cfg.Src.Classes, cfg.Src.Forms, cfg.Src.Workbook} {
		if err := os.MkdirAll(filepath.Join(root, filepath.FromSlash(sourceDir)), 0o755); err != nil {
			t.Fatal(err)
		}
	}
}

func packFormSources(t *testing.T, root string, cfg config.Config) []packpkg.SourceModule {
	t.Helper()
	sources, err := collectPackSourceModules(root, cfg)
	if err != nil {
		t.Fatalf("collectPackSourceModules: %v", err)
	}
	var forms []packpkg.SourceModule
	for _, source := range sources {
		if source.Type == packpkg.ModuleTypeForm {
			forms = append(forms, source)
		}
	}
	return forms
}

func TestCollectPackSourceModulesSidecarMergesBasIntoSource(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default() // code_source defaults to "sidecar"
	makePackSourceRoots(t, root, cfg)
	formsDir := filepath.Join(root, filepath.FromSlash(cfg.Src.Forms))
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	frm := "VERSION 5.00\r\nBegin {GUID} UserForm1\r\n   Caption = \"x\"\r\nEnd\r\n" +
		"Attribute VB_Name = \"UserForm1\"\r\n" +
		"Private Sub a()\r\n    Debug.Print \"OLD\"\r\nEnd Sub\r\n"
	bas := "Private Sub a()\r\n    Debug.Print \"NEW\"\r\nEnd Sub\r\n"
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "UserForm1.bas"), []byte(bas), 0o644); err != nil {
		t.Fatal(err)
	}
	got := packFormSources(t, root, cfg)
	if len(got) != 1 || got[0].Name != "UserForm1" || got[0].Type != packpkg.ModuleTypeForm {
		t.Fatalf("unexpected sources: %+v", got)
	}
	if !strings.Contains(got[0].Source, "NEW") {
		t.Errorf("sidecar code not merged into source: %q", got[0].Source)
	}
	after, _ := os.ReadFile(filepath.Join(formsDir, "UserForm1.frm"))
	if string(after) != frm {
		t.Error("pack must not write the .frm on disk")
	}
}

func TestCollectPackSourceModulesFrmModeReadsFrmVerbatim(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.UserForm.CodeSource = "frm"
	makePackSourceRoots(t, root, cfg)
	formsDir := filepath.Join(root, filepath.FromSlash(cfg.Src.Forms))
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	frm := "Attribute VB_Name = \"UserForm1\"\r\nPrivate Sub a()\r\n    Debug.Print \"FRM\"\r\nEnd Sub\r\n"
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}
	// A stray sidecar must be ignored in frm mode.
	if err := os.WriteFile(filepath.Join(formsDir, "code", "UserForm1.bas"), []byte("Private Sub a()\r\n    Debug.Print \"SIDE\"\r\nEnd Sub\r\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	got := packFormSources(t, root, cfg)
	if len(got) != 1 || got[0].Source != frm {
		t.Errorf("frm mode should read .frm verbatim, got %+v", got)
	}
}

func TestCollectPackSourceModulesSidecarFallsBackToFrmWhenNoBas(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default() // sidecar
	makePackSourceRoots(t, root, cfg)
	formsDir := filepath.Join(root, filepath.FromSlash(cfg.Src.Forms))
	frm := "Attribute VB_Name = \"UserForm1\"\r\nPrivate Sub a()\r\nEnd Sub\r\n"
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}
	got := packFormSources(t, root, cfg)
	if len(got) != 1 || got[0].Source != frm {
		t.Errorf("missing sidecar should fall back to .frm, got %+v", got)
	}
}

func TestCollectPackSourceModulesSidecarWithAttributeHeaderFailsLoud(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	makePackSourceRoots(t, root, cfg)
	formsDir := filepath.Join(root, filepath.FromSlash(cfg.Src.Forms))
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte("Attribute VB_Name = \"UserForm1\"\r\ncode\r\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	bad := "Attribute VB_Name = \"UserForm1\"\r\nPrivate Sub a()\r\nEnd Sub\r\n"
	if err := os.WriteFile(filepath.Join(formsDir, "code", "UserForm1.bas"), []byte(bad), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := collectPackSourceModules(root, cfg)
	if !errors.Is(err, packpkg.ErrAmbiguousLayout) {
		t.Fatalf("want ErrAmbiguousLayout for attribute-bearing sidecar, got %v", err)
	}
}

func TestCollectPackSourceModulesOrphanSidecarFailsLoud(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	makePackSourceRoots(t, root, cfg)
	formsDir := filepath.Join(root, filepath.FromSlash(cfg.Src.Forms))
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	// sidecar with no matching .frm
	if err := os.WriteFile(filepath.Join(formsDir, "code", "Ghost.bas"), []byte("Private Sub a()\r\nEnd Sub\r\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := collectPackSourceModules(root, cfg)
	if !errors.Is(err, packpkg.ErrAmbiguousLayout) {
		t.Fatalf("want ErrAmbiguousLayout for orphan sidecar, got %v", err)
	}
}

func TestCollectPackSourceModulesNestedSidecarFailsLoud(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	makePackSourceRoots(t, root, cfg)
	formsDir := filepath.Join(root, filepath.FromSlash(cfg.Src.Forms))
	if err := os.MkdirAll(filepath.Join(formsDir, "code", "nested"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte("Attribute VB_Name = \"UserForm1\"\r\ncode\r\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "nested", "Ghost.bas"), []byte("Option Explicit\r\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := collectPackSourceModules(root, cfg)
	if !errors.Is(err, packpkg.ErrAmbiguousLayout) {
		t.Fatalf("want ErrAmbiguousLayout for nested sidecar directory, got %v", err)
	}
}

func TestCollectPackSourceModulesNestedFrmWithSidecarNotOrphan(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default() // sidecar
	makePackSourceRoots(t, root, cfg)
	formsDir := filepath.Join(root, filepath.FromSlash(cfg.Src.Forms))
	if err := os.MkdirAll(filepath.Join(formsDir, "sub"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	// A .frm kept in a subdirectory is discovered recursively by shared source
	// inventory, so its sidecar must not be rejected as an orphan.
	frm := "Attribute VB_Name = \"Nested\"\r\nPrivate Sub a()\r\nEnd Sub\r\n"
	if err := os.WriteFile(filepath.Join(formsDir, "sub", "Nested.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "code", "Nested.bas"), []byte("Private Sub a()\r\nEnd Sub\r\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	got := packFormSources(t, root, cfg)
	var found bool
	for _, m := range got {
		if m.Name == "Nested" && m.Type == packpkg.ModuleTypeForm {
			found = true
		}
	}
	if !found {
		t.Errorf("nested Nested.frm should be collected as a form, got %+v", got)
	}
}
