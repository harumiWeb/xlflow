package analyze

import (
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
