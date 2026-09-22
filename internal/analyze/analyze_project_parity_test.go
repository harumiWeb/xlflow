package analyze

import (
	"cmp"
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

type analysisParityFixture struct {
	name      string
	files     []analysisParityFile
	required  []analysisParityExpectation
	forbidden []analysisParityExpectation
}

type analysisParityFile struct {
	path string
	kind sourceproject.ModuleKind
}

type analysisParityExpectation struct {
	code      string
	file      string
	module    string
	procedure string
	line      int
}

type analysisParityDiagnostic struct {
	Code      string
	Severity  string
	File      string
	Module    string
	Procedure string
	Line      int
	Column    int
	Message   string
	Reason    string
}

func TestAnalyzerFilesystemAndInMemoryParity(t *testing.T) {
	fixtures := []analysisParityFixture{
		{
			name: "single-standard",
			files: []analysisParityFile{
				{path: "src/modules/Main.bas", kind: sourceproject.ModuleKindStandard},
			},
			required: []analysisParityExpectation{
				{code: "VBA202", file: "src/modules/Main.bas", module: "Main", procedure: "Run", line: 7},
				{code: "VBA205", file: "src/modules/Main.bas", module: "Main", procedure: "Run", line: 10},
			},
			forbidden: []analysisParityExpectation{
				{code: "VBA205", file: "src/modules/Main.bas", module: "Main", procedure: "Run", line: 9},
			},
		},
		{
			name: "module-kinds",
			files: []analysisParityFile{
				{path: "src/modules/Main.bas", kind: sourceproject.ModuleKindStandard},
				{path: "src/classes/Worker.cls", kind: sourceproject.ModuleKindClass},
				{path: "src/workbook/ThisWorkbook.cls", kind: sourceproject.ModuleKindDocument},
				{path: "src/workbook/Sheet1.cls", kind: sourceproject.ModuleKindDocument},
			},
			required: []analysisParityExpectation{
				{code: "VBA101", file: "src/modules/Main.bas", module: "Main", procedure: "StandardRun", line: 4},
				{code: "VBA101", file: "src/classes/Worker.cls", module: "Worker", procedure: "ClassRun", line: 4},
				{code: "VBA101", file: "src/workbook/ThisWorkbook.cls", module: "ThisWorkbook", procedure: "WorkbookRun", line: 4},
				{code: "VBA101", file: "src/workbook/Sheet1.cls", module: "Sheet1", procedure: "WorksheetRun", line: 4},
			},
		},
		{
			name: "cross-module",
			files: []analysisParityFile{
				{path: "src/modules/Main.bas", kind: sourceproject.ModuleKindStandard},
				{path: "src/modules/Helpers.bas", kind: sourceproject.ModuleKindStandard},
			},
			required: []analysisParityExpectation{
				{code: "VBA228", file: "src/modules/Main.bas", module: "Main", procedure: "Validate", line: 7},
				{code: "VBA251", file: "src/modules/Main.bas", module: "Main", procedure: "Lookup", line: 12},
				{code: "VBA244", file: "src/modules/Helpers.bas", module: "Helpers", procedure: "Beta", line: 3},
			},
		},
	}

	for _, fixture := range fixtures {
		t.Run(fixture.name, func(t *testing.T) {
			fixtureRoot := filepath.Join("testdata", "parity", fixture.name)
			project, aliases := loadAnalysisParityProject(t, fixtureRoot, fixture.files)
			db, err := vbadb.LoadBuiltin()
			if err != nil {
				t.Fatalf("load built-in type database: %v", err)
			}
			typeDB := &TypeDatabase{DB: db, Complete: false}
			cfg := config.Default()

			filesystemResult, err := (Analyzer{
				RootDir: fixtureRoot,
				Config:  cfg,
				TypeDB:  typeDB,
			}).RunResultContext(t.Context())
			if err != nil {
				t.Fatalf("filesystem-backed analysis: %v", err)
			}

			inMemoryResult, err := (Analyzer{
				RootDir: t.TempDir(),
				Config:  cfg,
				TypeDB:  typeDB,
			}).AnalyzeProject(t.Context(), project)
			if err != nil {
				t.Fatalf("in-memory analysis: %v", err)
			}

			if filesystemResult.AnalyzedFiles != len(fixture.files) {
				t.Fatalf("filesystem analyzed files = %d, want %d", filesystemResult.AnalyzedFiles, len(fixture.files))
			}
			if inMemoryResult.AnalyzedFiles != len(fixture.files) {
				t.Fatalf("in-memory analyzed files = %d, want %d", inMemoryResult.AnalyzedFiles, len(fixture.files))
			}

			filesystemFindings := normalizeAnalysisParityDiagnostics(filesystemResult.Findings, aliases)
			inMemoryFindings := normalizeAnalysisParityDiagnostics(inMemoryResult.Findings, aliases)
			if !reflect.DeepEqual(filesystemFindings, inMemoryFindings) {
				t.Fatalf("normalized findings differ:\nfilesystem: %#v\nin-memory: %#v", filesystemFindings, inMemoryFindings)
			}

			filesystemPreflight := normalizeAnalysisParityDiagnostics(filesystemResult.PreflightFindings, aliases)
			inMemoryPreflight := normalizeAnalysisParityDiagnostics(inMemoryResult.PreflightFindings, aliases)
			if !reflect.DeepEqual(filesystemPreflight, inMemoryPreflight) {
				t.Fatalf("normalized preflight findings differ:\nfilesystem: %#v\nin-memory: %#v", filesystemPreflight, inMemoryPreflight)
			}
			for _, expectation := range fixture.required {
				assertAnalysisParityFinding(t, filesystemFindings, expectation, true)
			}
			for _, expectation := range fixture.forbidden {
				assertAnalysisParityFinding(t, filesystemFindings, expectation, false)
			}
		})
	}
}

func loadAnalysisParityProject(t *testing.T, fixtureRoot string, files []analysisParityFile) (sourceproject.SourceProject, map[string]string) {
	t.Helper()
	project := sourceproject.SourceProject{Files: make([]sourceproject.SourceFile, 0, len(files))}
	aliases := make(map[string]string, len(files)*2)
	for _, file := range files {
		physicalPath := filepath.Join(fixtureRoot, filepath.FromSlash(file.path))
		source, err := os.ReadFile(physicalPath)
		if err != nil {
			t.Fatalf("read parity fixture %s: %v", file.path, err)
		}
		absolutePath, err := filepath.Abs(physicalPath)
		if err != nil {
			t.Fatalf("resolve parity fixture %s: %v", file.path, err)
		}
		virtualPath := filepath.ToSlash(filepath.Join("virtual", filepath.FromSlash(file.path)))
		aliases[analysisParityPathKey(absolutePath)] = file.path
		aliases[analysisParityPathKey(virtualPath)] = file.path
		project.Files = append(project.Files, sourceproject.SourceFile{
			Path:       virtualPath,
			Source:     source,
			ModuleKind: file.kind,
		})
	}
	return project, aliases
}

func normalizeAnalysisParityDiagnostics(findings []Finding, aliases map[string]string) []analysisParityDiagnostic {
	normalized := make([]analysisParityDiagnostic, 0, len(findings))
	for _, finding := range findings {
		file := aliases[analysisParityPathKey(finding.File)]
		if file == "" {
			file = filepath.ToSlash(finding.File)
		}
		normalized = append(normalized, analysisParityDiagnostic{
			Code:      finding.Code,
			Severity:  finding.Severity,
			File:      file,
			Module:    finding.Module,
			Procedure: finding.Procedure,
			Line:      finding.Line,
			Column:    finding.Column,
			Message:   finding.Message,
			Reason:    finding.Reason,
		})
	}
	slices.SortFunc(normalized, func(a, b analysisParityDiagnostic) int {
		return cmp.Or(
			cmp.Compare(a.File, b.File),
			cmp.Compare(a.Line, b.Line),
			cmp.Compare(a.Column, b.Column),
			cmp.Compare(a.Code, b.Code),
			cmp.Compare(a.Module, b.Module),
			cmp.Compare(a.Procedure, b.Procedure),
			cmp.Compare(a.Message, b.Message),
			cmp.Compare(a.Reason, b.Reason),
		)
	})
	return normalized
}

func analysisParityPathKey(path string) string {
	key := filepath.ToSlash(filepath.Clean(path))
	if filepath.Separator == '\\' {
		key = strings.ToLower(key)
	}
	return key
}

func assertAnalysisParityFinding(t *testing.T, findings []analysisParityDiagnostic, want analysisParityExpectation, present bool) {
	t.Helper()
	for _, finding := range findings {
		if finding.Code == want.code && finding.File == want.file && finding.Module == want.module && finding.Procedure == want.procedure && finding.Line == want.line {
			if !present {
				t.Fatalf("unexpected diagnostic %+v found in %#v", want, findings)
			}
			return
		}
	}
	if present {
		t.Fatalf("required diagnostic %+v missing from %#v", want, findings)
	}
}
