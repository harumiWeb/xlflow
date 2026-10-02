package filepush

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
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
)

func TestPushAppliesSourceTreeAndWritesState(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Public Sub Run()\n    Debug.Print \"pushed\"\nEnd Sub\n"))
	writeTestFile(t, filepath.Join(root, "src", "classes", "Class1.cls"), classSourceText(t, "Class1"))
	writeTestFile(t, filepath.Join(root, "src", "workbook", "ThisWorkbook.bas"), "Option Explicit\n")
	writeTestFile(t, filepath.Join(root, "src", "workbook", "Sheet1.bas"), "Option Explicit\n")

	cfg := testConfig()
	result, err := Push(root, cfg, workbook, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if result.Skipped || result.Backup == nil || result.Backup.Reason != "before-push" || result.Backup.Backend != "file" {
		t.Fatalf("unexpected result: %+v", result)
	}
	if result.Meta.Standard != 1 || result.Meta.Class != 1 || result.Meta.Document != 2 {
		t.Fatalf("meta = %+v", result.Meta)
	}
	project := readProject(t, workbook)
	if got := moduleSourceOf(project, "Module1"); !strings.Contains(got, `Debug.Print "pushed"`) {
		t.Fatalf("Module1 source = %q", got)
	}
	if _, err := os.Stat(result.Backup.BackupFileAbsPath); err != nil {
		t.Fatalf("backup not created: %v", err)
	}
	state := readState(t, result.StatePath)
	current := computeFingerprint(workbook, discoverSourceFiles(resolvedRoots(root, cfg), cfg.UserForm.CodeSource), false)
	if !fingerprintEquals(state.Fingerprint, current) {
		t.Fatal("recorded fingerprint does not match the pushed source tree")
	}
	if !savedFileStampMatches(state.AppliedTo.SavedFile, workbook) {
		t.Fatal("saved_file stamp does not describe the published workbook")
	}
}

func TestPushChangedOnlySkipsUnchangedWorkbook(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	cfg := testConfig()
	first, err := Push(root, cfg, workbook, Options{ChangedOnly: true})
	if err != nil || first.Skipped {
		t.Fatalf("first push = %+v, %v", first, err)
	}
	second, err := Push(root, cfg, workbook, Options{ChangedOnly: true})
	if err != nil {
		t.Fatal(err)
	}
	if !second.Skipped {
		t.Fatal("unchanged source did not skip")
	}
	if second.Backup != nil {
		t.Fatal("skip must not create a backup")
	}
}

func TestPushChangedOnlyDetectsSourceAndWorkbookDrift(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	module := filepath.Join(root, "src", "modules", "Module1.bas")
	writeTestFile(t, module, moduleSource("Module1", "Option Explicit\n"))
	cfg := testConfig()
	if _, err := Push(root, cfg, workbook, Options{ChangedOnly: true}); err != nil {
		t.Fatal(err)
	}

	writeTestFile(t, module, moduleSource("Module1", "Option Explicit\nPublic Sub Changed()\nEnd Sub\n"))
	drift, err := Push(root, cfg, workbook, Options{ChangedOnly: true})
	if err != nil || drift.Skipped {
		t.Fatalf("source drift did not re-push: %+v, %v", drift, err)
	}

	// A workbook saved elsewhere invalidates the saved_file stamp even when the
	// source fingerprint is unchanged.
	writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	drift, err = Push(root, cfg, workbook, Options{ChangedOnly: true})
	if err != nil || drift.Skipped {
		t.Fatalf("workbook drift did not re-push: %+v, %v", drift, err)
	}
}

func TestPushRejectsDuplicateModuleNames(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Main.bas"), moduleSource("Main", "Option Explicit\n"))
	writeTestFile(t, filepath.Join(root, "src", "classes", "main.cls"), classSourceText(t, "main"))
	if _, err := Push(root, testConfig(), workbook, Options{}); !errors.Is(err, ErrDuplicateModule) {
		t.Fatalf("error = %v, want ErrDuplicateModule", err)
	}
}

func TestPushGuardRejectsBeforeMutation(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	before := mustRead(t, workbook)
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	guardErr := errors.New("workbook is open")
	_, err := Push(root, testConfig(), workbook, Options{
		Guard: func(string) error { return guardErr },
	})
	if !errors.Is(err, guardErr) {
		t.Fatalf("error = %v, want guard error", err)
	}
	if !bytes.Equal(mustRead(t, workbook), before) {
		t.Fatal("workbook changed despite guard rejection")
	}
	if _, statErr := os.Stat(filepath.Join(root, ".xlflow", "state", "push.json")); !errors.Is(statErr, os.ErrNotExist) {
		t.Fatal("push state written despite guard rejection")
	}
}

func TestPushBackupNeverSkipsBackup(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	result, err := Push(root, testConfig(), workbook, Options{BackupMode: "never"})
	if err != nil {
		t.Fatal(err)
	}
	if result.Backup != nil {
		t.Fatal("backup created under --backup never")
	}
	if entries, err := os.ReadDir(filepath.Join(root, ".xlflow", "backups")); err == nil && len(entries) > 0 {
		t.Fatal("backup directory populated under --backup never")
	}
}

func TestPushRejectsProtectedProject(t *testing.T) {
	root := newSourceTree(t)
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	protected := writeWorkbook(t, root, readFixture(t, "p3_protected.bin"))
	if _, err := Push(root, testConfig(), protected, Options{}); err == nil || !strings.Contains(err.Error(), "protected") {
		t.Fatalf("protected workbook error = %v", err)
	}
}

func TestPushLeavesWorkbookUntouchedWhenSourcePreflightFails(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	before := mustRead(t, workbook)
	cfg := testConfig()
	cfg.VBA.LineNumbers.Enabled = true
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Public Sub Run()\n    GoTo 10\nEnd Sub\n"))
	if _, err := Push(root, cfg, workbook, Options{}); !errors.Is(err, ErrLineNumberSafety) {
		t.Fatalf("error = %v, want ErrLineNumberSafety", err)
	}
	if !bytes.Equal(mustRead(t, workbook), before) {
		t.Fatal("workbook changed despite line-number preflight failure")
	}
}

func TestPushAppliesFolderAnnotationUpdate(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Domain", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	cfg := testConfig()
	cfg.VBA.FolderAnnotation = "update"
	if _, err := Push(root, cfg, workbook, Options{}); err != nil {
		t.Fatal(err)
	}
	project := readProject(t, workbook)
	if got := moduleSourceOf(project, "Module1"); !strings.Contains(got, `'@Folder("Domain")`) {
		t.Fatalf("Module1 source lacks folder annotation: %q", got)
	}
}

func TestPushUpdatesUserFormCodeBehindFromSidecar(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p4_form.bin"))
	frm := "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1\r\n" +
		"   Caption = \"x\"\r\nEnd\r\n" +
		"Attribute VB_Name = \"UserForm1\"\r\n" +
		"Attribute VB_GlobalNameSpace = False\r\n" +
		"Attribute VB_Creatable = False\r\n" +
		"Attribute VB_PredeclaredId = True\r\n" +
		"Attribute VB_Exposed = False\r\n"
	writeTestFile(t, filepath.Join(root, "src", "forms", "UserForm1.frm"), frm)
	writeTestFile(t, filepath.Join(root, "src", "forms", "code", "UserForm1.bas"), "Private Sub UserForm_Click()\n    Debug.Print \"SIDECAR\"\nEnd Sub\n")
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	result, err := Push(root, cfg, workbook, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if result.Meta.Form != 1 {
		t.Fatalf("meta = %+v, want one form", result.Meta)
	}
	project := readProject(t, workbook)
	if got := moduleSourceOf(project, "UserForm1"); !strings.Contains(got, "SIDECAR") {
		t.Fatalf("UserForm1 code-behind = %q", got)
	}
	// The tracked .frm keeps its embedded (empty) code: sidecar merge happens
	// in memory only.
	if body := mustRead(t, filepath.Join(root, "src", "forms", "UserForm1.frm")); bytes.Contains(body, []byte("SIDECAR")) {
		t.Fatal("file push rewrote the tracked .frm")
	}
}

func TestPushStateUsesDotNetShape(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	if _, err := Push(root, testConfig(), workbook, Options{}); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(filepath.Join(root, ".xlflow", "state", "push.json"))
	if err != nil {
		t.Fatal(err)
	}
	var raw map[string]any
	if err := json.Unmarshal(body, &raw); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"fingerprint", "applied_to"} {
		if _, ok := raw[key]; !ok {
			t.Fatalf("push state missing %q: %s", key, body)
		}
	}
	fingerprint := raw["fingerprint"].(map[string]any)
	for _, key := range []string{"workbook_path", "files", "line_numbers_enabled"} {
		if _, ok := fingerprint[key]; !ok {
			t.Fatalf("fingerprint missing %q", key)
		}
	}
	saved := raw["applied_to"].(map[string]any)["saved_file"].(map[string]any)
	for _, key := range []string{"path", "last_write_time_utc_ticks", "length"} {
		if _, ok := saved[key]; !ok {
			t.Fatalf("saved_file missing %q", key)
		}
	}
}

func testConfig() config.Config {
	cfg := config.Default()
	cfg.Excel.Path = "Book.xlsm"
	return cfg
}

// newSourceTree creates every configured source root because
// sourceinventory.Discover rejects a missing root.
func newSourceTree(t testing.TB) string {
	t.Helper()
	root := t.TempDir()
	for _, dir := range []string{"modules", "classes", "forms", "workbook"} {
		if err := os.MkdirAll(filepath.Join(root, "src", dir), 0o755); err != nil {
			t.Fatal(err)
		}
	}
	return root
}

func moduleSource(name, body string) string {
	return `Attribute VB_Name = "` + name + `"` + "\n" + body
}

func classSourceText(t testing.TB, name string) string {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "pack", "testdata", "disk", "p1", "classes", "Class1.cls"))
	if err != nil {
		t.Fatal(err)
	}
	return strings.ReplaceAll(string(body), "Class1", name)
}

func readFixture(t testing.TB, name string) []byte {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "pack", "vbaproject", "testdata", "corpus", name))
	if err != nil {
		t.Fatal(err)
	}
	return body
}

func writeWorkbook(t testing.TB, root string, project []byte) string {
	t.Helper()
	path := filepath.Join(root, "Book.xlsm")
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	add := func(name string, body []byte) {
		entry, err := w.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write(body); err != nil {
			t.Fatal(err)
		}
	}
	add("[Content_Types].xml", []byte(`<Types></Types>`))
	add("xl/workbook.xml", []byte(`<workbook/>`))
	if project != nil {
		add("xl/vbaProject.bin", project)
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	writeTestFile(t, path, buf.String())
	return path
}

func writeTestFile(t testing.TB, path, body string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func mustRead(t testing.TB, path string) []byte {
	t.Helper()
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return body
}

func readProject(t testing.TB, workbookPath string) *vbaproject.Project {
	t.Helper()
	reader, err := zip.OpenReader(workbookPath)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = reader.Close() }()
	for _, entry := range reader.File {
		if entry.Name != "xl/vbaProject.bin" {
			continue
		}
		rc, err := entry.Open()
		if err != nil {
			t.Fatal(err)
		}
		bin, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		project, err := vbaproject.Read(bin)
		if err != nil {
			t.Fatal(err)
		}
		return project
	}
	t.Fatal("workbook has no xl/vbaProject.bin")
	return nil
}

func moduleSourceOf(project *vbaproject.Project, name string) string {
	for _, module := range project.Modules {
		if module.Name == name {
			return module.Source
		}
	}
	return ""
}

func readState(t testing.TB, path string) pushState {
	t.Helper()
	state, err := readPushState(path)
	if err != nil || state == nil || state.AppliedTo == nil {
		t.Fatalf("readPushState = %+v, %v", state, err)
	}
	return *state
}
