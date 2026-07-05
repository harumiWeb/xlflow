package ooxml

import (
	"archive/zip"
	"os"
	"path/filepath"
	"testing"
)

func TestReadWorkbookResolvesRelationshipsAndDefinedNames(t *testing.T) {
	path := writeZipFixture(t, map[string]string{
		"xl/workbook.xml": `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Invoice" sheetId="1" r:id="rId1"/>
    <sheet name="Summary" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="TaxRate">Config!$B$2</definedName>
    <definedName name="InvoiceTotal" localSheetId="0">Invoice!$G$12</definedName>
  </definedNames>
</workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships>
  <Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Target="/xl/worksheets/sheet2.xml"/>
</Relationships>`,
		"xl/worksheets/sheet1.xml": `<worksheet/>`,
		"xl/worksheets/sheet2.xml": `<worksheet/>`,
	})
	pkg, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = pkg.Close() }()

	wb, err := pkg.ReadWorkbook()
	if err != nil {
		t.Fatal(err)
	}
	if wb.Sheets[0].Path != "xl/worksheets/sheet1.xml" || wb.Sheets[1].Path != "xl/worksheets/sheet2.xml" {
		t.Fatalf("sheet paths = %#v", wb.Sheets)
	}
	if got := DefinedNameScope(wb.DefinedNames[0], wb.Sheets); got != "workbook" {
		t.Fatalf("first name scope = %q", got)
	}
	if wb.DefinedNames[0].RefersTo != "=Config!$B$2" {
		t.Fatalf("refers_to = %q", wb.DefinedNames[0].RefersTo)
	}
	if got := DefinedNameScope(wb.DefinedNames[1], wb.Sheets); got != "Invoice" {
		t.Fatalf("second name scope = %q", got)
	}
}

func TestReadWorksheetFormulasStreamsFormulaElementsOnly(t *testing.T) {
	path := writeZipFixture(t, map[string]string{
		"xl/worksheets/sheet1.xml": `<worksheet><sheetData>
  <row r="1"><c r="A1"><f>SUM(B1:C1)</f><v>12</v></c></row>
  <row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2*C2</f></c></row>
  <row r="3"><c r="A3"><f t="shared" si="0"/></c></row>
</sheetData></worksheet>`,
	})
	pkg, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = pkg.Close() }()

	formulas, err := pkg.ReadWorksheetFormulas("xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatal(err)
	}
	if len(formulas) != 3 {
		t.Fatalf("formula count = %d", len(formulas))
	}
	if formulas[0].Text != "=SUM(B1:C1)" {
		t.Fatalf("formula text = %q", formulas[0].Text)
	}
	if formulas[1].Type != "shared" || formulas[1].Ref != "A2:A3" || formulas[1].SharedIndex != "0" {
		t.Fatalf("shared formula = %#v", formulas[1])
	}
	if formulas[2].Text != "" {
		t.Fatalf("shared child text = %q", formulas[2].Text)
	}
}

func writeZipFixture(t *testing.T, files map[string]string) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "book.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	zw := zip.NewWriter(out)
	for name, body := range files {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(body)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	if err := out.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}
