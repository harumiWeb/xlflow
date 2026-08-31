package sourceprojectfs

import (
	"bytes"
	"context"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func TestLoadClassifiesConfiguredRootsAndLegacyTests(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.UserForm.CodeSource = "frm"

	writeSource(t, root, filepath.Join("src", "modules", "nested", "Standard.bas"), "standard")
	writeSource(t, root, filepath.Join("src", "classes", "Service.cls"), "class")
	writeSource(t, root, filepath.Join("src", "forms", "UserForm1.frm"), "form")
	writeSource(t, root, filepath.Join("src", "workbook", "ThisWorkbook.cls"), "document")
	writeSource(t, root, filepath.Join("tests", "StandardTest.bas"), "test")

	project, err := Load(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 5 {
		t.Fatalf("loaded %d files, want 5: %+v", len(project.Files), project.Files)
	}

	wantPaths := []string{
		filepath.Join(root, "src", "classes", "Service.cls"),
		filepath.Join(root, "src", "forms", "UserForm1.frm"),
		filepath.Join(root, "src", "modules", "nested", "Standard.bas"),
		filepath.Join(root, "src", "workbook", "ThisWorkbook.cls"),
		filepath.Join(root, "tests", "StandardTest.bas"),
	}
	sort.Strings(wantPaths)
	gotPaths := make([]string, len(project.Files))
	for i, file := range project.Files {
		gotPaths[i] = file.Path
	}
	if !reflect.DeepEqual(gotPaths, wantPaths) {
		t.Fatalf("paths = %v, want sorted %v", gotPaths, wantPaths)
	}

	wantKinds := map[string]sourceproject.ModuleKind{
		"Service.cls":      sourceproject.ModuleKindClass,
		"UserForm1.frm":    sourceproject.ModuleKindForm,
		"Standard.bas":     sourceproject.ModuleKindStandard,
		"ThisWorkbook.cls": sourceproject.ModuleKindDocument,
		"StandardTest.bas": sourceproject.ModuleKindStandard,
	}
	for _, file := range project.Files {
		wantKind, ok := wantKinds[filepath.Base(file.Path)]
		if !ok {
			t.Fatalf("unexpected file %q", file.Path)
		}
		wantTest := filepath.Base(file.Path) == "StandardTest.bas"
		if file.ModuleKind != wantKind || file.IsTest != wantTest {
			t.Fatalf("file %q = kind %q, isTest=%v; want kind %q, isTest=%v", file.Path, file.ModuleKind, file.IsTest, wantKind, wantTest)
		}
	}
}

func TestLoadUsesSidecarFormSourceAndSkipsDesignerArtifact(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	writeSource(t, root, filepath.Join("src", "forms", "UserForm1.frm"), "designer")
	sidecarPath := filepath.Join("src", "forms", "code", "UserForm1.bas")
	sidecarSource := []byte("Option Explicit\r\nPublic Sub Initialize_日本語()\r\nEnd Sub\r\n")
	writeBytes(t, root, sidecarPath, sidecarSource)

	project, err := Load(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 1 {
		t.Fatalf("loaded %d files, want one sidecar: %+v", len(project.Files), project.Files)
	}
	file := project.Files[0]
	if file.Path != filepath.Join(root, sidecarPath) || file.ModuleKind != sourceproject.ModuleKindForm || file.IsTest {
		t.Fatalf("sidecar file = %+v", file)
	}
	if !bytes.Equal(file.Source, sidecarSource) {
		t.Fatalf("sidecar source bytes changed: %q", file.Source)
	}
}

func TestLoadAppliesPathFilterBeforeReading(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	keepPath := filepath.Join(root, "src", "modules", "Keep.bas")
	dropPath := filepath.Join(root, "src", "modules", "Drop.bas")
	writeBytes(t, root, filepath.Join("src", "modules", "Keep.bas"), []byte("keep"))
	writeBytes(t, root, filepath.Join("src", "modules", "Drop.bas"), []byte("drop"))

	var filtered []string
	recorder := analysisstats.NewRecorder()
	ctx := analysisstats.WithRecorder(context.Background(), recorder)
	project, err := LoadContext(ctx, Options{
		RootDir: root,
		Config:  cfg,
		PathFilter: func(path string) bool {
			filtered = append(filtered, path)
			return filepath.Clean(path) == filepath.Clean(keepPath)
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 1 || project.Files[0].Path != keepPath {
		t.Fatalf("filtered project = %+v, want only %q", project.Files, keepPath)
	}
	if !reflect.DeepEqual(filtered, []string{dropPath, keepPath}) {
		t.Fatalf("filter paths = %v, want sorted absolute paths %v", filtered, []string{dropPath, keepPath})
	}
	stages, _ := recorder.Totals()
	foundDiscovery := false
	for _, stage := range stages {
		if stage.Name == "source_discovery" {
			foundDiscovery = true
			if stage.ResultCount != 1 {
				t.Fatalf("source discovery result count = %d, want one selected file", stage.ResultCount)
			}
		}
	}
	if !foundDiscovery {
		t.Fatalf("missing source_discovery stage: %+v", stages)
	}
}

func TestLoadPreservesRelativeLogicalPaths(t *testing.T) {
	cwd, err := os.Getwd()
	if err != nil {
		t.Fatal(err)
	}
	root, err := os.MkdirTemp(cwd, ".sourceprojectfs-relative-root-")
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = os.RemoveAll(root) })
	relativeRoot, err := filepath.Rel(cwd, root)
	if err != nil {
		t.Fatal(err)
	}
	relativeSource := filepath.Join(relativeRoot, "src", "modules", "Relative.bas")
	writeSource(t, root, filepath.Join("src", "modules", "Relative.bas"), "relative")

	var filteredPath string
	project, err := Load(Options{
		RootDir: relativeRoot,
		Config:  config.Default(),
		PathFilter: func(path string) bool {
			filteredPath = path
			return true
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 1 || project.Files[0].Path != relativeSource {
		t.Fatalf("relative project files = %+v, want logical path %q", project.Files, relativeSource)
	}
	if filteredPath != relativeSource {
		t.Fatalf("path filter received %q, want logical path %q", filteredPath, relativeSource)
	}
}

func TestLoadWithEmptyRootUsesWorkingDirectoryRelativeLogicalPaths(t *testing.T) {
	root := t.TempDir()
	t.Chdir(root)
	relativeSource := filepath.Join("src", "modules", "WorkingDirectory.bas")
	writeSource(t, root, relativeSource, "working directory")

	project, err := Load(Options{Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 1 || project.Files[0].Path != relativeSource {
		t.Fatalf("empty-root project files = %+v, want logical path %q", project.Files, relativeSource)
	}
}

func TestLoadDeduplicatesWithProductionPrecedenceAndSorts(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	// Pointing a configured production root at tests deliberately creates an
	// overlap. The production candidate must win and therefore must not be
	// marked IsTest.
	cfg.Src.Modules = "tests"
	writeSource(t, root, filepath.Join("tests", "z.bas"), "z")
	writeSource(t, root, filepath.Join("tests", "a.bas"), "a")

	project, err := Load(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 2 {
		t.Fatalf("loaded %d files, want 2 after dedupe: %+v", len(project.Files), project.Files)
	}
	if project.Files[0].Path >= project.Files[1].Path {
		t.Fatalf("files are not sorted: %q then %q", project.Files[0].Path, project.Files[1].Path)
	}
	for _, file := range project.Files {
		if file.IsTest {
			t.Fatalf("overlapping production file marked as test: %+v", file)
		}
	}
}

func TestLoadPreservesSourceBytesAndRecordsStages(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	want := []byte("Attribute VB_Name = \"Bytes\"\r\n' Unicode: 日本語\x00\r\n")
	path := filepath.Join("src", "modules", "Bytes.bas")
	writeBytes(t, root, path, want)

	recorder := analysisstats.NewRecorder()
	ctx := analysisstats.WithRecorder(context.Background(), recorder)
	project, err := LoadContext(ctx, Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 1 || !bytes.Equal(project.Files[0].Source, want) {
		t.Fatalf("source bytes = %q, want exact %q", project.Files[0].Source, want)
	}
	stages, _ := recorder.Totals()
	byName := make(map[string]analysisstats.Stage, len(stages))
	for _, stage := range stages {
		byName[stage.Name] = stage
	}
	if byName["source_discovery"].ResultCount != 1 || byName["file_read"].Calls != 1 || byName["file_read"].ResultCount != len(want) {
		t.Fatalf("loader stages = %+v, want one discovered file and one byte-counted read", byName)
	}
}

func TestLoadMissingRootsProduceEmptyProject(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	cfg.Src.Modules = "missing/modules"
	cfg.Src.Classes = "missing/classes"
	cfg.Src.Forms = "missing/forms"
	cfg.Src.Workbook = "missing/workbook"

	project, err := Load(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if len(project.Files) != 0 {
		t.Fatalf("missing roots project = %+v, want empty project", project)
	}
}

func TestLoadPropagatesCancellation(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	project, err := LoadContext(ctx, Options{RootDir: t.TempDir(), Config: config.Default()})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("error = %v, want context.Canceled", err)
	}
	if len(project.Files) != 0 {
		t.Fatalf("project = %+v, want zero value on cancellation", project)
	}
}

func TestLoadPropagatesReadError(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	path := filepath.Join(root, "src", "modules", "ReadError.bas")
	writeBytes(t, root, filepath.Join("src", "modules", "ReadError.bas"), []byte("source"))
	project, err := Load(Options{
		RootDir: root,
		Config:  cfg,
		PathFilter: func(got string) bool {
			if strings.EqualFold(filepath.Clean(got), filepath.Clean(path)) {
				return os.Remove(got) == nil
			}
			return true
		},
	})
	if len(project.Files) != 0 {
		t.Fatalf("project = %+v, want zero value on read error", project)
	}
	if err == nil {
		t.Fatal("Load returned nil error after filtered file was removed")
	}
	if !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("error = %v, want file read error", err)
	}
}

func writeSource(t *testing.T, root, relative, marker string) {
	t.Helper()
	writeBytes(t, root, relative, []byte("Option Explicit\n' "+marker+"\n"))
}

func writeBytes(t *testing.T, root, relative string, source []byte) {
	t.Helper()
	path := filepath.Join(root, relative)
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, source, 0o644); err != nil {
		t.Fatal(err)
	}
}
