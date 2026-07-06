package cli

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/xuri/excelize/v2"
)

func TestRootCommandIncludesFormulasPullCommand(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"formulas", "pull"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "pull" {
		t.Fatalf("expected formulas pull command, got %#v", cmd)
	}
	cmd, _, err = root.Find([]string{"formulas", "inspect"})
	if err != nil {
		t.Fatal(err)
	}
	if cmd == nil || cmd.Name() != "inspect" {
		t.Fatalf("expected formulas inspect command, got %#v", cmd)
	}
}

func TestFormulasPullWritesStableSnapshotsAndRemovesStaleOutput(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	cfg.Excel.Path = filepath.ToSlash(filepath.Join("build", "Book.xlsm"))
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	writeFormulaWorkbookFixture(t, filepath.Join(dir, "build", "Book.xlsm"))
	stale := filepath.Join(dir, "formulas", "sheets", "999-Stale.regions.jsonl")
	if err := os.MkdirAll(filepath.Dir(stale), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(stale, []byte("{}\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	manual := filepath.Join(dir, "formulas", "README.txt")
	if err := os.WriteFile(manual, []byte("keep me\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runFormulasCommandForTest(dir, "--json", "formulas", "pull")
	if err != nil {
		t.Fatalf("formulas pull error = %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatalf("invalid JSON: %v\n%s", err, stdout)
	}
	if env.Command != "formulas pull" {
		t.Fatalf("command = %q", env.Command)
	}
	outputPayload, ok := env.Output.(map[string]any)
	if !ok {
		t.Fatalf("output payload = %T", env.Output)
	}
	parseSummary, ok := outputPayload["parse_status_summary"].(map[string]any)
	if !ok {
		t.Fatalf("parse status summary payload = %T", outputPayload["parse_status_summary"])
	}
	if parseSummary["ok"] != float64(3) || parseSummary["partial"] != float64(1) || parseSummary["failed"] != float64(0) {
		t.Fatalf("parse status summary = %#v", parseSummary)
	}
	if _, err := os.Stat(stale); !os.IsNotExist(err) {
		t.Fatalf("stale output still exists or stat failed: %v", err)
	}
	if got := readText(t, manual); got != "keep me\n" {
		t.Fatalf("manual file = %q", got)
	}

	manifest := readText(t, filepath.Join(dir, "formulas", "manifest.json"))
	for _, want := range []string{
		`"workbook": "Book.xlsm"`,
		`"parse_status_summary": {`,
		`"ok": 3`,
		`"partial": 1`,
		`"failed": 0`,
		`"path": "sheets/001-Invoice.regions.jsonl"`,
		`"formula_region_count": 3`,
		`"path": "sheets/002-Summary.regions.jsonl"`,
	} {
		if !strings.Contains(manifest, want) {
			t.Fatalf("manifest missing %q:\n%s", want, manifest)
		}
	}
	names := readText(t, filepath.Join(dir, "formulas", "names.jsonl"))
	if !strings.Contains(names, `{"name":"TaxRate","scope":"workbook","refers_to":"=Config!$B$2","kind":"formula"}`) {
		t.Fatalf("names missing workbook name:\n%s", names)
	}
	if !strings.Contains(names, `{"name":"InvoiceTotal","scope":"Invoice","refers_to":"=Invoice!$G$12","kind":"formula"}`) {
		t.Fatalf("names missing sheet name:\n%s", names)
	}
	invoice := readText(t, filepath.Join(dir, "formulas", "sheets", "001-Invoice.regions.jsonl"))
	for _, want := range []string{
		`"range":"C2:C4","formula_r1c1":"=RC[-2]*RC[-1]","example_cell":"C2","example_formula":"=A2*B2","count":3,"parse_status":"ok","refs":["A2:A4","B2:B4"],"storage_kinds":["shared"]`,
		`"range":"G2:G3","formula_r1c1":"=RC[-2]*RC[-1]"`,
		`"range":"G4","formula":"=Table1[Amount]","example_cell":"G4","example_formula":"=Table1[Amount]","count":1,"parse_status":"partial","features":["structured_reference"]`,
	} {
		if !strings.Contains(invoice, want) {
			t.Fatalf("invoice regions missing %q:\n%s", want, invoice)
		}
	}
	assertJSONLRegionsCoverDistinctCells(t, filepath.Join(dir, "formulas", "sheets", "001-Invoice.regions.jsonl"), 6)
	assertJSONLRegionsCoverDistinctCells(t, filepath.Join(dir, "formulas", "sheets", "002-Summary.regions.jsonl"), 1)
}

func TestFormulasPullSupportsStandaloneSourceAndOutput(t *testing.T) {
	dir := t.TempDir()
	workbook := filepath.Join(dir, "fixtures", "Standalone.xlsx")
	outputDir := filepath.Join(dir, "snapshots")
	writeFormulaWorkbookFixture(t, workbook)

	stdout, err := runFormulasCommandForTest(dir, "--json", "formulas", "pull", "--src", workbook, "--out", outputDir)
	if err != nil {
		t.Fatalf("formulas pull error = %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatalf("invalid JSON: %v\n%s", err, stdout)
	}
	output, ok := env.Output.(map[string]any)
	if !ok {
		t.Fatalf("output payload = %T", env.Output)
	}
	if output["dir"] != "snapshots" {
		t.Fatalf("output dir = %#v, want snapshots", output["dir"])
	}
	if _, err := os.Stat(filepath.Join(outputDir, "manifest.json")); err != nil {
		t.Fatalf("manifest not written: %v", err)
	}
	if _, err := os.Stat(filepath.Join(dir, config.FileName)); !os.IsNotExist(err) {
		t.Fatalf("test should not create or require xlflow.toml: %v", err)
	}
}

func TestFormulasInspectViewsAndJSON(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	cfg.Excel.Path = filepath.ToSlash(filepath.Join("build", "Book.xlsm"))
	if err := config.Write(filepath.Join(dir, config.FileName), cfg); err != nil {
		t.Fatal(err)
	}
	writeFormulaWorkbookFixture(t, filepath.Join(dir, "build", "Book.xlsm"))
	if stdout, err := runFormulasCommandForTest(dir, "--json", "formulas", "pull"); err != nil {
		t.Fatalf("formulas pull error = %v\n%s", err, stdout)
	}

	stdout, err := runFormulasCommandForTest(dir, "formulas", "inspect")
	if err != nil {
		t.Fatalf("formulas inspect summary error = %v\n%s", err, stdout)
	}
	for _, want := range []string{"Formula summary", "Invoice", "formula regions: 3", "formula cells: 6", "Defined names:", "TaxRate -> =Config!$B$2"} {
		if !strings.Contains(stdout, want) {
			t.Fatalf("summary output missing %q:\n%s", want, stdout)
		}
	}

	stdout, err = runFormulasCommandForTest(dir, "formulas", "inspect", "--sheet", "Invoice")
	if err != nil {
		t.Fatalf("formulas inspect sheet error = %v\n%s", err, stdout)
	}
	for _, want := range []string{"Invoice formulas", "C2:C4", "pattern: =RC[-2]*RC[-1]", "G4", "parse: partial", "structured_reference"} {
		if !strings.Contains(stdout, want) {
			t.Fatalf("sheet output missing %q:\n%s", want, stdout)
		}
	}

	stdout, err = runFormulasCommandForTest(dir, "formulas", "inspect", "--cell", "Invoice!C3")
	if err != nil {
		t.Fatalf("formulas inspect cell error = %v\n%s", err, stdout)
	}
	for _, want := range []string{"Invoice!C3", "Region:", "C2:C4", "Expanded formula at Invoice!C3:", "=A3*B3"} {
		if !strings.Contains(stdout, want) {
			t.Fatalf("cell output missing %q:\n%s", want, stdout)
		}
	}

	stdout, err = runFormulasCommandForTest(dir, "formulas", "inspect", "--range", "Invoice!C2:G4")
	if err != nil {
		t.Fatalf("formulas inspect range error = %v\n%s", err, stdout)
	}
	for _, want := range []string{"Formula regions overlapping Invoice!C2:G4", "C2:C4", "G2:G3", "G4"} {
		if !strings.Contains(stdout, want) {
			t.Fatalf("range output missing %q:\n%s", want, stdout)
		}
	}

	stdout, err = runFormulasCommandForTest(dir, "--json", "formulas", "inspect", "--cell", "Invoice!C3")
	if err != nil {
		t.Fatalf("formulas inspect json error = %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatalf("invalid JSON: %v\n%s", err, stdout)
	}
	outputPayload := cliObjectMap(env.Output)
	inspectPayload := cliObjectMap(outputPayload["formulas_inspect"])
	if inspectPayload["view"] != "cell" || inspectPayload["cell"] != "Invoice!C3" || inspectPayload["expanded_formula"] != "=A3*B3" {
		t.Fatalf("inspect payload = %#v", inspectPayload)
	}
	region := cliObjectMap(inspectPayload["region"])
	if region["range"] != "C2:C4" || region["formula_r1c1"] != "=RC[-2]*RC[-1]" {
		t.Fatalf("region payload = %#v", region)
	}
}

func TestFormulasInspectErrors(t *testing.T) {
	dir := t.TempDir()
	stdout, err := runFormulasCommandForTest(dir, "--json", "formulas", "inspect", "--summary", "--sheet", "Invoice")
	if err == nil {
		t.Fatalf("expected conflicting selector error\n%s", stdout)
	}
	if !strings.Contains(stdout, "formulas_inspect_args_invalid") {
		t.Fatalf("conflicting selector output = %s", stdout)
	}

	stdout, err = runFormulasCommandForTest(dir, "--json", "formulas", "inspect", "--dir", "missing")
	if err == nil {
		t.Fatalf("expected missing snapshot error\n%s", stdout)
	}
	if !strings.Contains(stdout, "formulas_inspect_failed") || !strings.Contains(stdout, "manifest not found") {
		t.Fatalf("missing snapshot output = %s", stdout)
	}

	badDir := filepath.Join(dir, "bad")
	if err := os.MkdirAll(badDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(badDir, "manifest.json"), []byte(`{"version":1,"sheets":[{"name":"Invoice","path":"sheets/001-Invoice.regions.jsonl"}]}`), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(badDir, "sheets"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(badDir, "sheets", "001-Invoice.regions.jsonl"), []byte("{bad json}\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	stdout, err = runFormulasCommandForTest(dir, "--json", "formulas", "inspect", "--dir", badDir)
	if err == nil {
		t.Fatalf("expected malformed snapshot error\n%s", stdout)
	}
	if !strings.Contains(stdout, "formulas_inspect_failed") || strings.Contains(stdout, "formulas_inspect_args_invalid") {
		t.Fatalf("malformed snapshot output = %s", stdout)
	}
}

func runFormulasCommandForTest(dir string, args ...string) (string, error) {
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

func writeFormulaWorkbookFixture(t *testing.T, path string) {
	t.Helper()
	files := map[string]string{
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
  <Relationship Id="rId2" Target="worksheets/sheet2.xml"/>
</Relationships>`,
		"xl/worksheets/sheet1.xml": `<worksheet><sheetData>
  <row r="2"><c r="C2"><f t="shared" ref="C2:C4" si="0">A2*B2</f></c><c r="G2"><f>E2*F2</f><v>10</v></c></row>
  <row r="3"><c r="C3"><f t="shared" si="0"/></c><c r="G3"><f>E3*F3</f><v>20</v></c></row>
  <row r="4"><c r="C4"><f t="shared" si="0"/></c><c r="G4"><f>Table1[Amount]</f><v>30</v></c></row>
</sheetData></worksheet>`,
		"xl/worksheets/sheet2.xml": `<worksheet><sheetData>
  <row r="1"><c r="A1"><f>SUM(Invoice!G2:G3)</f></c></row>
</sheetData></worksheet>`,
	}
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
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
}

func readText(t *testing.T, path string) string {
	t.Helper()
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return string(body)
}

func assertJSONLRegionsCoverDistinctCells(t *testing.T, path string, wantCells int) {
	t.Helper()
	type regionRow struct {
		Range string `json:"range"`
		Count int    `json:"count"`
	}
	lines := strings.Split(strings.TrimSpace(readText(t, path)), "\n")
	seen := map[string]string{}
	gotCells := 0
	var previous string
	for _, line := range lines {
		if strings.TrimSpace(line) == "" {
			continue
		}
		var row regionRow
		if err := json.Unmarshal([]byte(line), &row); err != nil {
			t.Fatalf("invalid region JSON in %s: %v\n%s", path, err, line)
		}
		if previous != "" && compareRangeTopLeft(previous, row.Range) > 0 {
			t.Fatalf("regions are not sorted in %s: %s before %s", path, previous, row.Range)
		}
		previous = row.Range
		cells := expandA1RangeForTest(t, row.Range)
		if row.Count != len(cells) {
			t.Fatalf("region %s count = %d, expanded cells = %d", row.Range, row.Count, len(cells))
		}
		gotCells += len(cells)
		for _, cell := range cells {
			if prior, ok := seen[cell]; ok {
				t.Fatalf("cell %s appears in multiple regions in %s: %s and %s", cell, path, prior, row.Range)
			}
			seen[cell] = row.Range
		}
	}
	if gotCells != wantCells {
		t.Fatalf("expanded region cell count in %s = %d, want %d", path, gotCells, wantCells)
	}
}

func compareRangeTopLeft(left, right string) int {
	leftCell := strings.Split(left, ":")[0]
	rightCell := strings.Split(right, ":")[0]
	leftCol, leftRow, _ := excelize.CellNameToCoordinates(leftCell)
	rightCol, rightRow, _ := excelize.CellNameToCoordinates(rightCell)
	switch {
	case leftRow < rightRow:
		return -1
	case leftRow > rightRow:
		return 1
	case leftCol < rightCol:
		return -1
	case leftCol > rightCol:
		return 1
	default:
		return strings.Compare(left, right)
	}
}

func expandA1RangeForTest(t *testing.T, rangeText string) []string {
	t.Helper()
	parts := strings.Split(rangeText, ":")
	if len(parts) == 0 || len(parts) > 2 {
		t.Fatalf("invalid range %q", rangeText)
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(parts[0])
	if err != nil {
		t.Fatalf("invalid range start %q: %v", rangeText, err)
	}
	endCol, endRow := startCol, startRow
	if len(parts) == 2 {
		endCol, endRow, err = excelize.CellNameToCoordinates(parts[1])
		if err != nil {
			t.Fatalf("invalid range end %q: %v", rangeText, err)
		}
	}
	var cells []string
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				t.Fatalf("invalid coordinates %d,%d: %v", col, row, err)
			}
			cells = append(cells, cell)
		}
	}
	sort.Strings(cells)
	return cells
}
