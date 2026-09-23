package analyze

import (
	"archive/zip"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

type testWorksheetCodeNameCatalog map[string]string

func (catalog testWorksheetCodeNameCatalog) Lookup(name string) (string, bool) {
	codeName, ok := catalog[name]
	return codeName, ok
}

func TestAnalyzeProjectExcelSemanticInspections(t *testing.T) {
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}
	if err := db.MergeJSON([]byte(`{"types":[{"name":"Excel.WorksheetFunction","library":"Excel","kind":"interface","source":"typelib","confidence":"generated","members":[{"name":"Sum","kind":"method","return_type":"Double"}]}]}`)); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectWorksheetStringAccess = true
	cfg.Analyze.DetectApplicationWorksheetFunction = true
	cfg.Analyze.DetectHostBracketExpressions = true
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       "virtual/Main.bas",
			ModuleKind: sourceproject.ModuleKindStandard,
			Source:     []byte("Attribute VB_Name = \"Main\"\nPublic Sub Run()\n  Dim value As Variant\n  value = Application.Sum(1, 2)\n  value = [A1]\n  value = ThisWorkbook.Worksheets(\"Report\").Range(\"A1\").Value\nEnd Sub\n"),
		},
		{
			Path:       "virtual/SheetReport.cls",
			ModuleKind: sourceproject.ModuleKindDocument,
			Source:     []byte("Attribute VB_Name = \"SheetReport\"\nOption Explicit\n"),
		},
	}}
	result, err := (Analyzer{
		Config:             cfg,
		TypeDB:             &TypeDatabase{DB: db, Complete: true},
		WorksheetCodeNames: testWorksheetCodeNameCatalog{"Report": "SheetReport"},
	}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"VBA260", "VBA261", "VBA262"} {
		if findings := findingsByCode(result.Findings, code); len(findings) != 1 {
			t.Errorf("%s findings = %+v, want one", code, findings)
		}
	}
}

func TestAnalyzeProjectWorksheetCatalogMissingFailsOpenWithoutFilesystemRead(t *testing.T) {
	cfg := config.Default()
	cfg.Excel.Path = "must-not-be-opened.xlsm"
	cfg.Analyze.DetectWorksheetStringAccess = true
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source:     []byte("Public Sub Run()\n  Debug.Print ThisWorkbook.Worksheets(\"Report\").Name\nEnd Sub\n"),
	}}}
	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatal(err)
	}
	if findings := findingsByCode(result.Findings, "VBA260"); len(findings) != 0 {
		t.Fatalf("VBA260 findings without catalog = %+v", findings)
	}
	if !resultHasWarningCode(result, "analysis_capability_unavailable") {
		t.Fatalf("warnings = %+v, want worksheet catalog capability warning", result.Warnings)
	}
}

func TestAnalyzeProjectWorksheetCatalogMissingDoesNotWarnWithoutCandidate(t *testing.T) {
	cfg := config.Default()
	cfg.Excel.Path = "must-not-be-opened.xlsm"
	cfg.Analyze.DetectWorksheetStringAccess = true
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source:     []byte("Public Sub Run()\n  Debug.Print ThisWorkbook.Name\nEnd Sub\n"),
	}}}
	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatal(err)
	}
	if resultHasWarningCode(result, "analysis_capability_unavailable") {
		t.Fatalf("warnings = %+v, want no worksheet catalog capability warning", result.Warnings)
	}
}

func TestLoadWorksheetCodeNameCatalogWarnsWhenEntriesAreExcluded(t *testing.T) {
	workbookPath := writeWorksheetCatalogFixture(t, map[string]string{
		"xl/workbook.xml":            `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Report" sheetId="1" r:id="missing"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships/>`,
	})
	catalog, warning := loadWorksheetCodeNameCatalog("", workbookPath)
	if catalog == nil {
		t.Fatal("catalog = nil, want partial catalog")
	}
	if warning == nil || warning["code"] != "analysis_capability_unavailable" || warning["capability"] != "worksheet_codename_catalog" {
		t.Fatalf("warning = %+v, want worksheet catalog capability warning", warning)
	}
}

func writeWorksheetCatalogFixture(t *testing.T, files map[string]string) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "book.xlsx")
	output, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	archive := zip.NewWriter(output)
	for name, body := range files {
		entry, err := archive.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write([]byte(body)); err != nil {
			t.Fatal(err)
		}
	}
	if err := archive.Close(); err != nil {
		t.Fatal(err)
	}
	if err := output.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}

func resultHasWarningCode(result Result, code string) bool {
	for _, warning := range result.Warnings {
		if warning["code"] == code {
			return true
		}
	}
	return false
}
