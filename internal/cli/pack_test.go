package cli

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

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
	for _, name := range []string{"out", "template", "experimental"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected pack command to define --%s", name)
		}
	}
}

func TestPackCommandValidationFailures(t *testing.T) {
	cases := []struct {
		name string
		args []string
		code int
		err  string
	}{
		{name: "experimental gate", args: []string{"--json", "pack", "--out", "dist/Book.xlsm"}, code: output.ExitConfig, err: "pack_experimental_required"},
		{name: "missing out", args: []string{"--json", "pack", "--experimental"}, code: output.ExitConfig, err: "pack_args_invalid"},
		{name: "bad out extension", args: []string{"--json", "pack", "--experimental", "--out", "dist/Book.xlsx"}, code: output.ExitConfig, err: "pack_args_invalid"},
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

	stdout, err := runPackCommandForTest(dir, "--json", "pack", "--experimental", "--out", "dist/Book.xlsm")
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
	if env.Pack["backend"] != "pure-go" || env.Pack["vbe_validation"] != "not_performed" || env.Pack["experimental"] != true {
		t.Fatalf("unexpected pack payload: %#v", env.Pack)
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
