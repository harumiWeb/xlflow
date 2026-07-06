package formulas

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/formula"
)

func TestInspectSummarySheetCellAndRange(t *testing.T) {
	dir := writeInspectSnapshotFixture(t)

	summary, err := InspectSummary(dir)
	if err != nil {
		t.Fatal(err)
	}
	if summary.Workbook != "Book.xlsm" || len(summary.Sheets) != 3 {
		t.Fatalf("summary = %#v", summary)
	}
	invoice := summary.Sheets[0]
	if invoice.Name != "Invoice" || invoice.FormulaRegionCount != 3 || invoice.FormulaCellCount != 7 {
		t.Fatalf("invoice summary = %#v", invoice)
	}
	if invoice.ParseStatus.OK != 2 || invoice.ParseStatus.Partial != 1 || len(invoice.Features) != 1 || invoice.Features[0] != "structured_reference" {
		t.Fatalf("invoice parse/features = %#v", invoice)
	}
	if len(invoice.DependsOnSheets) != 1 || invoice.DependsOnSheets[0] != "Config" {
		t.Fatalf("invoice deps = %#v", invoice.DependsOnSheets)
	}
	if len(summary.DefinedNames) != 1 || summary.DefinedNames[0].Name != "TaxRate" {
		t.Fatalf("defined names = %#v", summary.DefinedNames)
	}

	sheet, err := InspectSheet(dir, "Invoice")
	if err != nil {
		t.Fatal(err)
	}
	if len(sheet.Regions) != 3 || sheet.Regions[0].Range != "D2:D4" || sheet.Regions[2].ParseStatus != "partial" {
		t.Fatalf("sheet regions = %#v", sheet.Regions)
	}

	cell, err := InspectCell(dir, "Invoice!E3")
	if err != nil {
		t.Fatal(err)
	}
	if cell.Region == nil || cell.Region.Range != "E2:E4" {
		t.Fatalf("cell region = %#v", cell.Region)
	}
	if cell.ExpandedFormula != "=D3*Config!$B$2" {
		t.Fatalf("expanded formula = %q", cell.ExpandedFormula)
	}

	empty, err := InspectCell(dir, "Invoice!A10")
	if err != nil {
		t.Fatal(err)
	}
	if empty.Region != nil || empty.ExpandedFormula != "" {
		t.Fatalf("empty cell result = %#v", empty)
	}

	overlap, err := InspectRange(dir, "Invoice!D3:E10")
	if err != nil {
		t.Fatal(err)
	}
	if len(overlap.Regions) != 2 || overlap.Regions[0].Range != "D2:D4" || overlap.Regions[1].Range != "E2:E4" {
		t.Fatalf("overlap regions = %#v", overlap.Regions)
	}
}

func TestInspectQuotedSheetAndR1C1Expansion(t *testing.T) {
	dir := writeInspectSnapshotFixture(t)

	cell, err := InspectCell(dir, "'Sales Data'!B5")
	if err != nil {
		t.Fatal(err)
	}
	if cell.Region == nil || cell.Region.Range != "B5" {
		t.Fatalf("cell region = %#v", cell.Region)
	}
	if cell.ExpandedFormula != "='Config Sheet'!$B$2+A5" {
		t.Fatalf("expanded formula = %q", cell.ExpandedFormula)
	}

	cases := []struct {
		name    string
		pattern string
		base    formula.CellRef
		want    string
	}{
		{name: "copied formula", pattern: "=RC[-2]*RC[-1]", base: formula.CellRef{Row: 500, Col: 4}, want: "=B500*C500"},
		{name: "mixed absolute relative", pattern: "=R1C1+R[1]C1+R1C[-1]", base: formula.CellRef{Row: 2, Col: 3}, want: "=$A$1+$A3+B$1"},
		{name: "range", pattern: "=SUM(R[-6]C[-2]:R[-1]C[-2])", base: formula.CellRef{Row: 8, Col: 3}, want: "=SUM(A2:A7)"},
		{name: "quoted sheet", pattern: "='売上 集計'!R2C2+RC[-1]", base: formula.CellRef{Row: 5, Col: 2}, want: "='売上 集計'!$B$2+A5"},
	}
	for _, tt := range cases {
		t.Run(tt.name, func(t *testing.T) {
			got, ok := ExpandR1C1Formula(tt.pattern, tt.base)
			if !ok {
				t.Fatalf("ExpandR1C1Formula returned ok=false")
			}
			if got != tt.want {
				t.Fatalf("expanded = %q, want %q", got, tt.want)
			}
		})
	}
}

func TestInspectSnapshotErrors(t *testing.T) {
	dir := t.TempDir()
	if _, err := InspectSummary(dir); err == nil || !strings.Contains(err.Error(), "manifest not found") {
		t.Fatalf("missing manifest error = %v", err)
	}

	dir = t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, "manifest.json"), []byte(`{"version":1,"sheets":[{"name":"Missing","path":"sheets/missing.regions.jsonl"}]}`), 0o644); err != nil {
		t.Fatal(err)
	}
	if _, err := InspectSummary(dir); err == nil || !strings.Contains(err.Error(), "file not found") {
		t.Fatalf("missing sheet file error = %v", err)
	}

	dir = writeInspectSnapshotFixture(t)
	if _, err := InspectSheet(dir, "Missing"); err == nil || !strings.Contains(err.Error(), "sheet not found") {
		t.Fatalf("missing sheet error = %v", err)
	}
	if _, err := InspectCell(dir, "Invoice!D2:E2"); err == nil || !strings.Contains(err.Error(), "invalid cell address") {
		t.Fatalf("invalid cell error = %v", err)
	}
}

func writeInspectSnapshotFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	mustWriteInspectFile(t, filepath.Join(dir, "manifest.json"), `{
  "version": 1,
  "workbook": "Book.xlsm",
  "parse_status_summary": {"ok": 4, "partial": 1, "failed": 0},
  "sheets": [
    {"index": 1, "name": "Invoice", "sheet_id": "1", "path": "sheets/001-Invoice.regions.jsonl", "formula_region_count": 3, "parse_status_summary": {"ok": 2, "partial": 1, "failed": 0}},
    {"index": 2, "name": "Summary", "sheet_id": "2", "path": "sheets/002-Summary.regions.jsonl", "formula_region_count": 1, "parse_status_summary": {"ok": 1, "partial": 0, "failed": 0}},
    {"index": 3, "name": "Sales Data", "sheet_id": "3", "path": "sheets/003-Sales-Data.regions.jsonl", "formula_region_count": 1, "parse_status_summary": {"ok": 1, "partial": 0, "failed": 0}}
  ]
}`)
	mustWriteInspectFile(t, filepath.Join(dir, "names.jsonl"), "{\"name\":\"TaxRate\",\"scope\":\"workbook\",\"refers_to\":\"=Config!$B$2\",\"kind\":\"formula\"}\n")
	mustWriteInspectFile(t, filepath.Join(dir, "sheets", "001-Invoice.regions.jsonl"), strings.Join([]string{
		`{"range":"D2:D4","formula_r1c1":"=RC[-2]*RC[-1]","example_cell":"D2","example_formula":"=B2*C2","count":3,"parse_status":"ok","refs":["B2:B4","C2:C4"]}`,
		`{"range":"E2:E4","formula_r1c1":"=RC[-1]*Config!R2C2","example_cell":"E2","example_formula":"=D2*Config!$B$2","count":3,"parse_status":"ok","depends_on_sheets":["Config"]}`,
		`{"range":"G4","formula":"=SUM(SalesTable[Amount])","example_cell":"G4","example_formula":"=SUM(SalesTable[Amount])","count":1,"parse_status":"partial","features":["structured_reference"],"functions":["SUM"]}`,
	}, "\n")+"\n")
	mustWriteInspectFile(t, filepath.Join(dir, "sheets", "002-Summary.regions.jsonl"), `{"range":"A1","formula_r1c1":"=SUM(Invoice!R[1]C[6]:R[2]C[6])","example_cell":"A1","example_formula":"=SUM(Invoice!G2:G3)","count":1,"parse_status":"ok","depends_on_sheets":["Invoice"],"functions":["SUM"]}`+"\n")
	mustWriteInspectFile(t, filepath.Join(dir, "sheets", "003-Sales-Data.regions.jsonl"), `{"range":"B5","formula_r1c1":"='Config Sheet'!R2C2+RC[-1]","example_cell":"B5","example_formula":"='Config Sheet'!$B$2+A5","count":1,"parse_status":"ok","depends_on_sheets":["'Config Sheet'"]}`+"\n")
	return dir
}

func mustWriteInspectFile(t *testing.T, path, body string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}
