package filepush

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
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
	snapshot, err := captureSourceSnapshot(root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	current, err := computeFingerprint(workbook, snapshot.files, false, cfg.VBA.FolderAnnotation)
	if err != nil {
		t.Fatal(err)
	}
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
		Guard: func(context.Context, string) error { return guardErr },
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

func TestPushCaseDistinctAssetsDirectory(t *testing.T) {
	root := newSourceTree(t)
	formsDir := filepath.Join(root, "src", "forms")
	writeTestFile(t, filepath.Join(formsDir, "assets", "ignored.frm"), "retained opaque asset")
	if _, err := os.Stat(filepath.Join(formsDir, "Assets")); err == nil {
		t.Skip("requires a case-sensitive filesystem")
	} else if !os.IsNotExist(err) {
		t.Fatal(err)
	}
	workbook := writeWorkbook(t, root, readFixture(t, "p4_form.bin"))
	frm := "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\n"
	writeTestFile(t, filepath.Join(formsDir, "Assets", "UserForm1.frm"), frm)
	writeTestFile(t, filepath.Join(formsDir, "code", "UserForm1.bas"), "Private Sub KeepMe()\n    Debug.Print \"case-sensitive-directory\"\nEnd Sub\n")
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	files, err := discoverSourceFiles(resolvedRoots(root, cfg), cfg.UserForm.CodeSource)
	if err != nil {
		t.Fatal(err)
	}
	var foundForm bool
	for _, file := range files {
		if file.Kind == "form" {
			if file.RelativePath != "Assets/UserForm1.frm" {
				t.Fatalf("reserved asset entered discovery: %+v", file)
			}
			foundForm = true
		}
	}
	if !foundForm {
		t.Fatal("case-distinct nested form missing from source discovery")
	}
	result, err := Push(root, cfg, workbook, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if result.Meta.Form != 1 || !strings.Contains(moduleSourceOf(readProject(t, workbook), "UserForm1"), "KeepMe") {
		t.Fatal("case-distinct nested form lost during push")
	}
	var trackedForm bool
	for _, entry := range readState(t, result.StatePath).Fingerprint.Files {
		if entry.Kind == "form" {
			if entry.Path != "Assets/UserForm1.frm" {
				t.Fatalf("reserved asset entered fingerprint: %+v", entry)
			}
			trackedForm = true
		}
	}
	if !trackedForm {
		t.Fatal("case-distinct nested form missing from fingerprint")
	}
}

func TestPushChangedOnlyTracksFormSpecAndReferencedPictureBytes(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	assetDir := filepath.Join(root, "src", "forms", "assets")
	if err := os.MkdirAll(assetDir, 0o755); err != nil {
		t.Fatal(err)
	}
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	assetPath := filepath.Join(assetDir, "123abc456def.bmp")
	if err := os.WriteFile(assetPath, bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	specPath := filepath.Join(root, "src", "forms", "specs", "Login.json")
	spec := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login","caption":"Before"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/assets/123abc456def.bmp"}}]}`
	writeTestFile(t, specPath, spec)

	first, err := Push(root, cfg, workbook, Options{ChangedOnly: true, BackupMode: "never"})
	if err != nil || first.Skipped {
		t.Fatalf("first push = %+v, %v", first, err)
	}
	// Old SHA-addressed files and other user files are retained without
	// becoming source inputs or invalidating changed-only state.
	if err := os.WriteFile(filepath.Join(assetDir, "deadbeef.bmp"), bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(assetDir, "manifest.txt"), []byte("retained"), 0o644); err != nil {
		t.Fatal(err)
	}
	unchanged, err := Push(root, cfg, workbook, Options{ChangedOnly: true, BackupMode: "never"})
	if err != nil || !unchanged.Skipped {
		t.Fatalf("retained assets invalidated unchanged skip: %+v, %v", unchanged, err)
	}

	writeTestFile(t, specPath, strings.Replace(spec, "Before", "After", 1))
	changed, err := Push(root, cfg, workbook, Options{ChangedOnly: true, BackupMode: "never"})
	if err != nil || changed.Skipped {
		t.Fatalf("FormSpec-only change did not invalidate fingerprint: %+v, %v", changed, err)
	}
}

func TestPushCaseDistinctCodeDirectory(t *testing.T) {
	root := newSourceTree(t)
	formsDir := filepath.Join(root, "src", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(filepath.Join(formsDir, "Code")); err == nil {
		t.Skip("requires a case-sensitive filesystem")
	} else if !os.IsNotExist(err) {
		t.Fatal(err)
	}
	writeTestFile(t, filepath.Join(formsDir, "Code", "Utility.bas"), moduleSource("Utility", "Kept"))
	writeTestFile(t, filepath.Join(formsDir, "Code", "Helper.cls"), classSourceText(t, "Helper"))
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	first, err := Push(root, cfg, workbook, Options{BackupMode: "never", ChangedOnly: true})
	if err != nil || first.Skipped {
		t.Fatalf("push = %+v, %v", first, err)
	}
	project := readProject(t, workbook)
	for _, name := range []string{"Utility", "Helper"} {
		if moduleSourceOf(project, name) == "" {
			t.Fatalf("case-distinct loose module %s was omitted", name)
		}
	}
	unchanged, err := Push(root, cfg, workbook, Options{BackupMode: "never", ChangedOnly: true})
	if err != nil || !unchanged.Skipped {
		t.Fatalf("unchanged push = %+v, %v", unchanged, err)
	}
	for _, path := range []string{"Utility.bas", "Helper.cls"} {
		file := filepath.Join(formsDir, "Code", path)
		body, err := os.ReadFile(file)
		if err != nil {
			t.Fatal(err)
		}
		writeTestFile(t, file, string(body)+"\n' changed\n")
		changed, err := Push(root, cfg, workbook, Options{BackupMode: "never", ChangedOnly: true})
		if err != nil || changed.Skipped {
			t.Fatalf("%s change omitted from fingerprint: %+v, %v", path, changed, err)
		}
	}
}

func TestPushLocksReferencedPictureDirectory(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	assetDir := filepath.Join(root, "assets")
	if err := os.MkdirAll(assetDir, 0o755); err != nil {
		t.Fatal(err)
	}
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(assetDir, "logo.bas"), bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestFile(t, filepath.Join(root, "src", "forms", "specs", "Login.json"), `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"assets/logo.bas"}}]}`)
	snapshot, err := captureSourceSnapshot(root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	var pictureFound bool
	for _, input := range snapshot.files {
		if canonicalPathKey(input.FullPath) == canonicalPathKey(filepath.Join(assetDir, "logo.bas")) {
			pictureFound = input.Kind == "form_asset"
		}
	}
	if !pictureFound {
		t.Fatal("referenced arbitrary-suffix picture was not inventoried as a form asset")
	}
	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewSourceTreeIdentity(root, assetDir)
	if err != nil {
		t.Fatal(err)
	}
	held, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity: identity, Command: "pull", OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceSourceTree,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = held.Release() }()
	if _, err := Push(root, cfg, workbook, Options{Coordination: manager, BackupMode: "never"}); !errors.Is(err, coordination.ErrSourceTreeBusy) {
		t.Fatalf("push error = %v, want referenced asset directory lock contention", err)
	}
}

func TestPushReferencedPictureWithBasSuffixIsNotImportedAsModule(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	imagePath := filepath.Join(root, "src", "forms", "images", "logo.bas")
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Dir(imagePath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(imagePath, bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	writeTestFile(t, filepath.Join(root, "src", "forms", "specs", "Login.json"), `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/images/logo.bas"}}]}`)
	writeTestFile(t, filepath.Join(root, "src", "forms", "code", "Login.bas"), "Private Sub FormCode()\nEnd Sub\n")

	snapshot, err := captureSourceSnapshot(root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	var pictureFound bool
	for _, file := range snapshot.files {
		if canonicalPathKey(file.FullPath) == canonicalPathKey(imagePath) && file.Kind == "form_asset" {
			pictureFound = true
		}
	}
	if !pictureFound {
		t.Fatal("picture with module suffix was not inventoried as an asset")
	}
	for _, module := range snapshot.modules {
		if strings.EqualFold(module.Name, "logo") {
			t.Fatalf("picture bytes were submitted as VBA module source: %+v", module)
		}
	}
	if _, err := Push(root, cfg, workbook, Options{BackupMode: "never"}); err != nil {
		t.Fatal(err)
	}
	if got := moduleSourceOf(readProject(t, workbook), "logo"); got != "" {
		t.Fatalf("picture was imported as a VBA module: %q", got)
	}
}

func TestPushReferencedPictureSymlinkWithBasSuffixIsNotImportedAsModule(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	imagePath := filepath.Join(root, "src", "forms", "assets", "picture.bmp")
	bmp, err := os.ReadFile(filepath.Join("..", "vba", "userforms", "compiler", "testdata", "pictures-excel-authored", "logo.bmp"))
	if err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Dir(imagePath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(imagePath, bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	linkPath := filepath.Join(root, "src", "forms", "Linked.bas")
	if err := os.Symlink(filepath.Join("assets", "picture.bmp"), linkPath); err != nil {
		t.Skipf("cannot create an image symlink: %v", err)
	}
	writeTestFile(t, filepath.Join(root, "src", "forms", "specs", "Login.json"), `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login"},"controls":[{"id":"image1","name":"Image1","type":"Image","picture":{"path":"src/forms/Linked.bas"}}]}`)
	snapshot, err := captureSourceSnapshot(root, cfg)
	if err != nil {
		t.Fatal(err)
	}
	var pictureFound bool
	for _, file := range snapshot.files {
		if file.Kind == "form_asset" && file.RelativePath == "src/forms/Linked.bas" && canonicalPathKey(file.FullPath) == canonicalPathKey(imagePath) {
			pictureFound = true
		}
	}
	if !pictureFound {
		t.Fatal("symlink picture did not retain its logical reference and physical target")
	}
	first, err := Push(root, cfg, workbook, Options{BackupMode: "never", ChangedOnly: true})
	if err != nil || first.Skipped {
		t.Fatalf("initial symlink picture push = %+v, %v", first, err)
	}
	if got := moduleSourceOf(readProject(t, workbook), "Linked"); got != "" {
		t.Fatalf("symlink picture was imported as a VBA module: %q", got)
	}
	unchanged, err := Push(root, cfg, workbook, Options{BackupMode: "never", ChangedOnly: true})
	if err != nil || !unchanged.Skipped {
		t.Fatalf("unchanged symlink picture push = %+v, %v", unchanged, err)
	}
	bmp[len(bmp)-1] ^= 1
	if err := os.WriteFile(imagePath, bmp, 0o644); err != nil {
		t.Fatal(err)
	}
	changed, err := Push(root, cfg, workbook, Options{BackupMode: "never", ChangedOnly: true})
	if err != nil || changed.Skipped {
		t.Fatalf("symlink target change did not invalidate the fingerprint: %+v, %v", changed, err)
	}
}

func TestPushCanonicalFormMissingAndEmptySidecarSemantics(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	cfg := testConfig()
	cfg.UserForm.CodeSource = "sidecar"
	specPath := filepath.Join(root, "src", "forms", "specs", "Login.json")
	sidecarPath := filepath.Join(root, "src", "forms", "code", "Login.bas")
	spec := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Login","caption":"Initial"},"controls":[]}`
	writeTestFile(t, specPath, spec)
	writeTestFile(t, sidecarPath, "Private Sub KeepMe()\n    Debug.Print \"keep\"\nEnd Sub\n")
	if _, err := Push(root, cfg, workbook, Options{BackupMode: "never"}); err != nil {
		t.Fatalf("initial form push: %v", err)
	}
	initialCode := moduleSourceOf(readProject(t, workbook), "Login")

	if err := os.Remove(sidecarPath); err != nil {
		t.Fatal(err)
	}
	cfg.Pack.UserFormTopology = "source"
	cfg.VBA.LineNumbers.Enabled = true
	writeTestFile(t, specPath, strings.Replace(spec, "Initial", "Updated", 1))
	if _, err := Push(root, cfg, workbook, Options{BackupMode: "never"}); err != nil {
		t.Fatalf("push with omitted sidecar: %v", err)
	}
	project := readProject(t, workbook)
	if got := moduleSourceOf(project, "Login"); got != initialCode {
		t.Fatalf("omitted sidecar changed existing code:\n got: %q\nwant: %q", got, initialCode)
	}
	if len(project.Forms) != 1 {
		t.Fatalf("omitted forms were removed under source topology: %d forms", len(project.Forms))
	}

	if err := os.WriteFile(sidecarPath, []byte{}, 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := Push(root, cfg, workbook, Options{BackupMode: "never"}); err != nil {
		t.Fatalf("push with explicit empty sidecar: %v", err)
	}
	project = readProject(t, workbook)
	if got := moduleSourceOf(project, "Login"); strings.Contains(got, "KeepMe") || strings.Contains(got, `Debug.Print "keep"`) {
		t.Fatalf("explicit empty sidecar did not clear code: %q", got)
	}

	if err := os.Remove(specPath); err != nil {
		t.Fatal(err)
	}
	if err := os.Remove(sidecarPath); err != nil {
		t.Fatal(err)
	}
	if _, err := Push(root, cfg, workbook, Options{BackupMode: "never"}); err != nil {
		t.Fatalf("push with omitted form under source topology: %v", err)
	}
	project = readProject(t, workbook)
	if len(project.Forms) != 1 {
		t.Fatalf("filepush removed an omitted form from the template topology: %d forms", len(project.Forms))
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
	for _, key := range []string{"workbook_path", "files", "line_numbers_enabled", "folder_annotation"} {
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

// Loose .bas/.cls files under src/forms are valid bridge input: the Excel
// backend imports them as standard/class modules, so the file backend must
// push them instead of rejecting the layout.
func TestPushImportsLooseFormModules(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	writeTestFile(t, filepath.Join(root, "src", "forms", "FormHelper.bas"), moduleSource("FormHelper", "Public Function Loose() As String\n    Loose = \"loose-bas\"\nEnd Function\n"))
	writeTestFile(t, filepath.Join(root, "src", "forms", "FormUtil.cls"), classSourceText(t, "FormUtil"))

	result, err := Push(root, testConfig(), workbook, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if result.Meta.Standard != 2 || result.Meta.Class != 1 {
		t.Fatalf("meta = %+v, want 2 standard + 1 class", result.Meta)
	}
	project := readProject(t, workbook)
	if got := moduleSourceOf(project, "FormHelper"); !strings.Contains(got, "loose-bas") {
		t.Fatalf("loose form .bas not imported: %q", got)
	}
	if got := moduleSourceOf(project, "FormUtil"); !strings.Contains(got, "VB_Name = \"FormUtil\"") {
		t.Fatalf("loose form .cls not imported: %q", got)
	}
}

// The same file seen as /mnt/<drive>/... under WSL and <drive>:\... on Windows
// must produce one canonical state path, or --changed-only can never skip
// across the backend/host boundary the bridge schema is shared across.
func TestPushStatePathsCanonicalAcrossHosts(t *testing.T) {
	previous := runningOnWSL
	runningOnWSL = func() bool { return true }
	t.Cleanup(func() { runningOnWSL = previous })

	if got := normalizeFingerprintPath(`C:\proj\build\Book.xlsm`); got != `C:\proj\build\Book.xlsm` {
		t.Fatalf("windows path = %q", got)
	}
	if got := normalizeFingerprintPath(`/mnt/c/proj/build/Book.xlsm`); got != `C:\proj\build\Book.xlsm` {
		t.Fatalf("wsl mount path = %q, want windows form", got)
	}
	if got := normalizeFingerprintPath(`d:\proj\book.xlsm`); got != `D:\proj\book.xlsm` {
		t.Fatalf("drive letter not canonicalized: %q", got)
	}
	// A non-/mnt Linux path has no Windows counterpart.
	for _, p := range []string{`/home/dev/proj/Book.xlsm`, `/mnt/`, `/mnt/1x/Book.xlsm`} {
		if _, ok := wslMountToWindowsPath(p); ok {
			t.Fatalf("wslMountToWindowsPath(%q) unexpectedly mapped", p)
		}
	}

	left, err := computeFingerprint(`/mnt/c/proj/build/Book.xlsm`, nil, false, "")
	if err != nil {
		t.Fatal(err)
	}
	right := sourceFingerprint{WorkbookPath: `C:\proj\build\Book.xlsm`}
	if !fingerprintEquals(left, right) {
		t.Fatal("cross-host workbook paths broke fingerprint equality")
	}
	if !pathsEqual(`C:\proj\build\Book.xlsm`, `/mnt/c/proj/build/Book.xlsm`) {
		t.Fatal("cross-host saved-file paths did not compare equal")
	}

	runningOnWSL = func() bool { return false }
	if got := normalizeFingerprintPath(`/mnt/c/proj/Book.xlsm`); got == `C:\proj\Book.xlsm` {
		t.Fatalf("non-WSL /mnt path was rewritten: %q", got)
	}
}

// A configured source root that does not exist contributes no files: the
// Excel bridge skips missing directories, so the file backend must tolerate
// them too instead of failing the push with a layout error.
func TestPushToleratesMissingSourceRoots(t *testing.T) {
	root := t.TempDir()
	if err := os.MkdirAll(filepath.Join(root, "src", "modules"), 0o755); err != nil {
		t.Fatal(err)
	}
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))

	result, err := Push(root, testConfig(), workbook, Options{})
	if err != nil {
		t.Fatalf("push with missing source roots: %v", err)
	}
	if result.Meta.Standard != 1 {
		t.Fatalf("meta = %+v, want 1 standard module", result.Meta)
	}
}

// The mutation window holds the workbook lease so a concurrent rollback, pack,
// or other mutator cannot race the atomic replace. Contention fails fast with
// the shared busy contract; a changed-only skip still never touches the lease.
func TestPushFailsBusyWhenWorkbookLeaseHeld(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	before := mustRead(t, workbook)

	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, workbook)
	if err != nil {
		t.Fatal(err)
	}
	held, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity: identity, Command: "rollback", OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = held.Release() }()

	_, err = Push(root, testConfig(), workbook, Options{Coordination: manager})
	var busy *coordination.BusyError
	if !errors.Is(err, ErrWorkbookLease) || !errors.As(err, &busy) {
		t.Fatalf("error = %v, want ErrWorkbookLease wrapping BusyError", err)
	}
	if !bytes.Equal(mustRead(t, workbook), before) {
		t.Fatal("workbook changed despite lease contention")
	}
}

// A recovery marker records an indeterminate Excel operation; the file backend
// must honor it even when the lock file and session record have cleared, just
// like lease-coordinated commands do.
func TestPushRejectsWorkbookRequiringRecovery(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	before := mustRead(t, workbook)

	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	identity, err := coordination.NewWorkbookIdentity(root, workbook)
	if err != nil {
		t.Fatal(err)
	}
	lease, err := manager.Acquire(context.Background(), coordination.AcquireRequest{
		Identity: identity, Command: "push", OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := lease.PublishRecovery(coordination.RecoveryPublication{
		Reason: "save timed out", Operation: "save",
	}); err != nil {
		t.Fatal(err)
	}
	if err := lease.Release(); err != nil {
		t.Fatal(err)
	}

	_, err = Push(root, testConfig(), workbook, Options{Coordination: manager})
	var required *coordination.RecoveryRequiredError
	if !errors.Is(err, ErrRecoveryCheck) || !errors.As(err, &required) {
		t.Fatalf("error = %v, want ErrRecoveryCheck wrapping RecoveryRequiredError", err)
	}
	if !bytes.Equal(mustRead(t, workbook), before) {
		t.Fatal("workbook changed despite recovery requirement")
	}
}

// Folder-annotation mode participates in the fingerprint: changing the mode
// without touching sources must invalidate a recorded changed-only state.
func TestPushChangedOnlyDetectsFolderAnnotationModeChange(t *testing.T) {
	root := newSourceTree(t)
	workbook := writeWorkbook(t, root, readFixture(t, "p1_compiled.bin"))
	writeTestFile(t, filepath.Join(root, "src", "modules", "Module1.bas"), moduleSource("Module1", "Option Explicit\n"))
	cfg := testConfig()
	cfg.VBA.FolderAnnotation = "ignore"
	if _, err := Push(root, cfg, workbook, Options{ChangedOnly: true}); err != nil {
		t.Fatal(err)
	}
	cfg.VBA.FolderAnnotation = "update"
	result, err := Push(root, cfg, workbook, Options{ChangedOnly: true})
	if err != nil {
		t.Fatal(err)
	}
	if result.Skipped {
		t.Fatal("annotation mode change did not invalidate changed-only state")
	}
}
