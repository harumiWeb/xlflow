package inspect

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestWorkbookSummarizesSheets(t *testing.T) {
	path := createInspectWorkbook(t)

	got, err := Workbook(path)
	if err != nil {
		t.Fatal(err)
	}

	if got.Path != path {
		t.Fatalf("path = %q, want %q", got.Path, path)
	}
	if got.Name != filepath.Base(path) {
		t.Fatalf("name = %q, want %q", got.Name, filepath.Base(path))
	}
	if got.ActiveSheet != "Visible" {
		t.Fatalf("active sheet = %q, want Visible", got.ActiveSheet)
	}
	if len(got.Sheets) != 2 {
		t.Fatalf("sheet count = %d, want 2", len(got.Sheets))
	}
	if got.Sheets[0].UsedRange != "A1:C2" || got.Sheets[0].RowCount != 2 || got.Sheets[0].ColumnCount != 3 {
		t.Fatalf("unexpected visible sheet summary: %#v", got.Sheets[0])
	}
	if got.Sheets[1].Name != "Hidden" || got.Sheets[1].Visible {
		t.Fatalf("unexpected hidden sheet summary: %#v", got.Sheets[1])
	}
}

func TestRangeReturnsMatrixValues(t *testing.T) {
	path := createInspectWorkbook(t)

	got, err := Range(path, "Visible", "A1:C2", Limits{MaxRows: 10, MaxCols: 10})
	if err != nil {
		t.Fatal(err)
	}

	if got.Range != "A1:C2" {
		t.Fatalf("range = %q, want A1:C2", got.Range)
	}
	if got.ReturnedRange != "A1:C2" {
		t.Fatalf("returned range = %q, want A1:C2", got.ReturnedRange)
	}
	if got.Truncated {
		t.Fatalf("truncated = true, want false")
	}
	want := [][]any{
		{"A1", nil, "C1"},
		{nil, "B2", nil},
	}
	if len(got.Values) != len(want) {
		t.Fatalf("row count = %d, want %d", len(got.Values), len(want))
	}
	for i := range want {
		for j := range want[i] {
			if got.Values[i][j] != want[i][j] {
				t.Fatalf("values[%d][%d] = %#v, want %#v", i, j, got.Values[i][j], want[i][j])
			}
		}
	}
}

func TestUsedRangeAppliesLimits(t *testing.T) {
	path := createInspectWorkbook(t)

	got, err := UsedRange(path, "Visible", Limits{MaxRows: 1, MaxCols: 2})
	if err != nil {
		t.Fatal(err)
	}

	if got.UsedRange != "A1:C2" {
		t.Fatalf("used range = %q, want A1:C2", got.UsedRange)
	}
	if got.ReturnedRange != "A1:B1" {
		t.Fatalf("returned range = %q, want A1:B1", got.ReturnedRange)
	}
	if !got.Truncated {
		t.Fatal("expected truncated used range")
	}
	if len(got.Warnings) != 1 || !strings.Contains(got.Warnings[0], "truncated") {
		t.Fatalf("warnings = %#v, want truncation warning", got.Warnings)
	}
}

func TestUsedRangeUsesSparseBounds(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Sparse.xlsx")
	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "D10", "value"); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}

	got, err := UsedRange(path, "Sheet1", Limits{MaxRows: 10, MaxCols: 10})
	if err != nil {
		t.Fatal(err)
	}
	if got.UsedRange != "D10" {
		t.Fatalf("used range = %q, want D10", got.UsedRange)
	}
	if got.ReturnedRange != "D10" {
		t.Fatalf("returned range = %q, want D10", got.ReturnedRange)
	}
	if got.RowCount != 1 || got.ColumnCount != 1 {
		t.Fatalf("size = %d x %d, want 1 x 1", got.RowCount, got.ColumnCount)
	}
}

func TestUsedRangeEmptySheetEmitsEmptyValues(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Empty.xlsx")
	f := excelize.NewFile()
	if _, err := f.NewSheet("Blank"); err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteSheet("Sheet1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}

	got, err := UsedRange(path, "Blank", Limits{MaxRows: 10, MaxCols: 10})
	if err != nil {
		t.Fatal(err)
	}
	data, err := json.Marshal(got)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), `"values":[]`) {
		t.Fatalf("json = %s, want values array", data)
	}
}

func TestCellReturnsNilForBlankCells(t *testing.T) {
	path := createInspectWorkbook(t)

	got, err := Cell(path, "Visible", "B1")
	if err != nil {
		t.Fatal(err)
	}
	if got.Address != "B1" {
		t.Fatalf("address = %q, want B1", got.Address)
	}
	if got.Value != nil {
		t.Fatalf("value = %#v, want nil", got.Value)
	}
}

func TestRangeRejectsMissingSheet(t *testing.T) {
	path := createInspectWorkbook(t)

	_, err := Range(path, "Missing", "A1", Limits{MaxRows: 10, MaxCols: 10})
	if err == nil || !strings.Contains(err.Error(), "not found") {
		t.Fatalf("expected missing sheet error, got %v", err)
	}
}

func createInspectWorkbook(t *testing.T) string {
	t.Helper()

	dir := t.TempDir()
	path := filepath.Join(dir, "Inspect.xlsx")
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "Visible"); err != nil {
		t.Fatal(err)
	}
	if _, err := f.NewSheet("Hidden"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Visible", "A1", "A1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Visible", "C1", "C1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Visible", "B2", "B2"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Hidden", "A1", "secret"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSheetVisible("Hidden", false); err != nil {
		t.Fatal(err)
	}
	f.SetActiveSheet(0)
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}
