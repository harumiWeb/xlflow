package cli

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
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
	if _, err := os.Stat(stale); !os.IsNotExist(err) {
		t.Fatalf("stale output still exists or stat failed: %v", err)
	}

	manifest := readText(t, filepath.Join(dir, "formulas", "manifest.json"))
	for _, want := range []string{
		`"workbook": "Book.xlsm"`,
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
		`"range":"C2:C4","formula_r1c1":"=RC[-2]*RC[-1]","example_cell":"C2","example_formula":"=A2*B2","count":3,"parse_status":"ok","storage_kinds":["shared"]`,
		`"range":"G2:G3","formula_r1c1":"=RC[-2]*RC[-1]"`,
		`"range":"G4","formula":"=Table1[Amount]","example_cell":"G4","example_formula":"=Table1[Amount]","count":1,"parse_status":"partial","features":["structured_reference"]`,
	} {
		if !strings.Contains(invoice, want) {
			t.Fatalf("invoice regions missing %q:\n%s", want, invoice)
		}
	}
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
