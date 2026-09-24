package analyze

import (
	"path/filepath"
	"strings"
	"testing"
)

type worksheetCodeNameCatalogFixture map[string]string

func (c worksheetCodeNameCatalogFixture) Lookup(name string) (string, bool) {
	for visibleName, codeName := range c {
		if strings.EqualFold(visibleName, name) {
			return codeName, true
		}
	}
	return "", false
}

func TestWorksheetCodeNameFindingsResolveOnlyThisWorkbookLiteralAccess(t *testing.T) {
	catalog := worksheetCodeNameCatalogFixture{"Report": "SheetReport"}
	documentModules := map[string]string{"sheetreport": "SheetReport"}
	for _, test := range []struct {
		name      string
		statement string
		want      int
	}{
		{name: "worksheets item", statement: `Set target = ThisWorkbook.Worksheets("Report").Range("A1")`, want: 1},
		{name: "explicit item member", statement: `Set target = ThisWorkbook.Worksheets.Item("Report").Range("A1")`, want: 1},
		{name: "case insensitive lookup", statement: `Set target = ThisWorkbook.Worksheets("report").Range("A1")`, want: 1},
		{name: "member qualified thisworkbook", statement: `Set target = provider.ThisWorkbook.Worksheets("Report").Range("A1")`, want: 0},
		{name: "spaced member qualified thisworkbook", statement: `Set target = provider . ThisWorkbook.Worksheets("Report").Range("A1")`, want: 0},
		{name: "external workbook", statement: `Set target = Workbooks("Book.xlsx").Worksheets("Report").Range("A1")`, want: 0},
		{name: "sheets collection", statement: `Set target = ThisWorkbook.Sheets("Report").Range("A1")`, want: 0},
		{name: "dynamic selector", statement: `Set target = ThisWorkbook.Worksheets(sheetName).Range("A1")`, want: 0},
		{name: "index selector", statement: `Set target = ThisWorkbook.Worksheets(1).Range("A1")`, want: 0},
		{name: "direct codename", statement: `Set target = SheetReport.Range("A1")`, want: 0},
	} {
		t.Run(test.name, func(t *testing.T) {
			got := worksheetCodeNameFindings(test.statement, catalog, documentModules)
			if len(got) != test.want {
				t.Fatalf("worksheet CodeName findings = %#v, want %d", got, test.want)
			}
			if test.want == 1 && (got[0].Code != "VBA260" || got[0].CodeName != "SheetReport") {
				t.Fatalf("finding = %#v", got[0])
			}
		})
	}
}

func TestWorksheetCodeNameStatementFindingsUsePhysicalContinuationRange(t *testing.T) {
	root := t.TempDir()
	lines := []string{
		"Public Sub Run()",
		"    Set target = ThisWorkbook _",
		"        .Worksheets(\"Report\").Range(\"A1\")",
		"End Sub",
	}
	statement, positions, ok := worksheetLogicalStatementWithPositions(lines, 1, 2)
	if !ok {
		t.Fatal("logical worksheet statement was not built")
	}
	file := parsedFile{Path: filepath.Join(root, "Main.bas"), Lines: lines, Module: "Main"}
	proc := sourceProcedure{Name: "Run", StartLine: 1, EndLine: 4}
	analyzer := Analyzer{RootDir: root, WorksheetCodeNames: worksheetCodeNameCatalogFixture{"Report": "SheetReport"}}
	findings := analyzer.worksheetCodeNameStatementFindings(file, proc, 2, statement, positions, false, map[string]string{"sheetreport": "SheetReport"})
	if len(findings) != 1 {
		t.Fatalf("VBA260 findings = %+v, want one", findings)
	}
	finding := findings[0]
	wantStart := strings.Index(lines[2], "Worksheets") + 1
	wantEnd := strings.Index(lines[2], ")") + 2
	if finding.Line != 3 || finding.Column != wantStart || finding.EndLine != 3 || finding.EndColumn != wantEnd {
		t.Fatalf("VBA260 range = %d:%d-%d:%d, want 3:%d-3:%d", finding.Line, finding.Column, finding.EndLine, finding.EndColumn, wantStart, wantEnd)
	}

	if shadowed := analyzer.worksheetCodeNameStatementFindings(file, proc, 2, statement, positions, true, map[string]string{"sheetreport": "SheetReport"}); len(shadowed) != 0 {
		t.Fatalf("shadowed ThisWorkbook findings = %+v, want none", shadowed)
	}
}

func TestWorksheetCodeNameFindingsRequireUniqueMappingAndDocumentModule(t *testing.T) {
	statement := `Set target = ThisWorkbook.Worksheets("Report").Range("A1")`
	catalog := worksheetCodeNameCatalogFixture{"Report": "SheetReport"}
	for _, test := range []struct {
		name            string
		documentModules map[string]string
		want            int
	}{
		{name: "missing module", documentModules: map[string]string{"sheetother": "SheetOther"}},
		{name: "ambiguous module", documentModules: map[string]string{"first": "SheetReport", "second": "SheetReport"}},
		{name: "module key fallback", documentModules: map[string]string{"SheetReport": ""}, want: 1},
	} {
		t.Run(test.name, func(t *testing.T) {
			got := worksheetCodeNameFindings(statement, catalog, test.documentModules)
			if len(got) != test.want {
				t.Fatalf("worksheet CodeName findings = %#v, want %d", got, test.want)
			}
		})
	}
}

func TestWorksheetCodeNameFindingsIgnoreStringsAndComments(t *testing.T) {
	catalog := worksheetCodeNameCatalogFixture{"Report": "SheetReport"}
	documentModules := map[string]string{"sheetreport": "SheetReport"}
	for _, statement := range []string{
		`Debug.Print "ThisWorkbook.Worksheets(""Report"")" ' ThisWorkbook.Worksheets("Report")`,
		`Rem ThisWorkbook.Worksheets("Report")`,
		`Debug.Print 1: Rem ThisWorkbook.Worksheets("Report")`,
	} {
		if got := worksheetCodeNameFindings(statement, catalog, documentModules); len(got) != 0 {
			t.Fatalf("worksheet CodeName findings for %q = %#v, want none", statement, got)
		}
	}
}
