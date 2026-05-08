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

	got, err := Range(path, "Visible", "A1:C2", RangeOptions{Limits: Limits{MaxRows: 10, MaxCols: 10}})
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
	if got.StyleIncluded {
		t.Fatal("style should not be included without --include-style")
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

func TestRangeIncludeStyleReturnsCellRowColumnAndMergeMetadata(t *testing.T) {
	path := createStyledInspectWorkbook(t)

	got, err := Range(path, "Styled", "A1:C3", RangeOptions{
		Limits:       Limits{MaxRows: 10, MaxCols: 10},
		IncludeStyle: true,
	})
	if err != nil {
		t.Fatal(err)
	}

	if !got.StyleIncluded {
		t.Fatal("expected style_included")
	}
	if len(got.Cells) != 9 {
		t.Fatalf("cell count = %d, want 9", len(got.Cells))
	}
	if len(got.Rows) != 3 {
		t.Fatalf("row count = %d, want 3", len(got.Rows))
	}
	if len(got.Columns) != 3 {
		t.Fatalf("column count = %d, want 3", len(got.Columns))
	}
	if len(got.MergedRanges) != 1 || got.MergedRanges[0] != "B2:C3" {
		t.Fatalf("merged ranges = %#v, want [B2:C3]", got.MergedRanges)
	}

	cellA1 := findCellSnapshot(t, got.Cells, "A1")
	if cellA1.Formula == nil || *cellA1.Formula != "=SUM(B1:C1)" {
		t.Fatalf("A1 formula = %#v, want =SUM(B1:C1)", cellA1.Formula)
	}
	if cellA1.Fill == nil || cellA1.Fill.Type != "solid" || cellA1.Fill.Color == nil || *cellA1.Fill.Color != "#000000" {
		t.Fatalf("A1 fill = %#v, want solid #000000", cellA1.Fill)
	}
	if cellA1.Font == nil || cellA1.Font.Color == nil || *cellA1.Font.Color != "#FFFFFF" || !cellA1.Font.Bold || !cellA1.Font.Italic {
		t.Fatalf("A1 font = %#v, want white bold italic font", cellA1.Font)
	}
	if cellA1.Border.Right.Style != "thin" || cellA1.Border.Right.Color == nil || *cellA1.Border.Right.Color != "#D9D9D9" {
		t.Fatalf("A1 right border = %#v, want thin #D9D9D9", cellA1.Border.Right)
	}
	if cellA1.NumberFormat == nil || *cellA1.NumberFormat != "0.00%" {
		t.Fatalf("A1 number format = %#v, want 0.00%%", cellA1.NumberFormat)
	}
	if cellA1.HorizontalAlignment == nil || *cellA1.HorizontalAlignment != "center" {
		t.Fatalf("A1 horizontal alignment = %#v, want center", cellA1.HorizontalAlignment)
	}
	if cellA1.VerticalAlignment == nil || *cellA1.VerticalAlignment != "center" {
		t.Fatalf("A1 vertical alignment = %#v, want center", cellA1.VerticalAlignment)
	}

	cellB2 := findCellSnapshot(t, got.Cells, "B2")
	if !cellB2.Merged || cellB2.MergeRange == nil || *cellB2.MergeRange != "B2:C3" {
		t.Fatalf("B2 merge = merged:%t range:%#v, want B2:C3", cellB2.Merged, cellB2.MergeRange)
	}

	cellC3 := findCellSnapshot(t, got.Cells, "C3")
	if !cellC3.Merged || cellC3.MergeRange == nil || *cellC3.MergeRange != "B2:C3" {
		t.Fatalf("C3 merge = merged:%t range:%#v, want B2:C3", cellC3.Merged, cellC3.MergeRange)
	}

	if got.Rows[1].Row != 2 || got.Rows[1].Height != 25 || !got.Rows[1].Hidden {
		t.Fatalf("row 2 metadata = %#v, want height 25 hidden true", got.Rows[1])
	}
	if got.Columns[1].Column != "B" || got.Columns[1].Width != 20 || !got.Columns[1].Hidden {
		t.Fatalf("column B metadata = %#v, want width 20 hidden true", got.Columns[1])
	}
}

func TestUsedRangeAppliesLimits(t *testing.T) {
	path := createInspectWorkbook(t)

	got, err := UsedRange(path, "Visible", RangeOptions{Limits: Limits{MaxRows: 1, MaxCols: 2}})
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

	got, err := UsedRange(path, "Sheet1", RangeOptions{Limits: Limits{MaxRows: 10, MaxCols: 10}})
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

	got, err := UsedRange(path, "Blank", RangeOptions{Limits: Limits{MaxRows: 10, MaxCols: 10}, IncludeStyle: true})
	if err != nil {
		t.Fatal(err)
	}
	data, err := json.Marshal(got)
	if err != nil {
		t.Fatal(err)
	}
	text := string(data)
	if !strings.Contains(text, `"values":[]`) {
		t.Fatalf("json = %s, want values array", data)
	}
	if !strings.Contains(text, `"cells":[]`) || !strings.Contains(text, `"merged_ranges":[]`) {
		t.Fatalf("json = %s, want empty style arrays", data)
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

	_, err := Range(path, "Missing", "A1", RangeOptions{Limits: Limits{MaxRows: 10, MaxCols: 10}})
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

func createStyledInspectWorkbook(t *testing.T) string {
	t.Helper()

	dir := t.TempDir()
	path := filepath.Join(dir, "StyledInspect.xlsx")
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "Styled"); err != nil {
		t.Fatal(err)
	}
	styleID, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"000000"},
		},
		Font: &excelize.Font{
			Family: "Calibri",
			Size:   11,
			Bold:   true,
			Italic: true,
			Color:  "FFFFFF",
		},
		Border: []excelize.Border{
			{Type: "right", Style: 1, Color: "D9D9D9"},
			{Type: "bottom", Style: 1, Color: "D9D9D9"},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		NumFmt: 10,
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellFormula("Styled", "A1", "=SUM(B1:C1)"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Styled", "B1", 0.25); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Styled", "C1", "label"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle("Styled", "A1", "A1", styleID); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle("Styled", "B2", "C3", styleID); err != nil {
		t.Fatal(err)
	}
	if err := f.MergeCell("Styled", "B2", "C3"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Styled", "B2", "merged"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetRowHeight("Styled", 2, 25); err != nil {
		t.Fatal(err)
	}
	if err := f.SetRowVisible("Styled", 2, false); err != nil {
		t.Fatal(err)
	}
	if err := f.SetColWidth("Styled", "B", "B", 20); err != nil {
		t.Fatal(err)
	}
	if err := f.SetColVisible("Styled", "B", false); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}

func findCellSnapshot(t *testing.T, cells []StyledCellSnapshot, address string) StyledCellSnapshot {
	t.Helper()
	for _, cell := range cells {
		if cell.Address == address {
			return cell
		}
	}
	t.Fatalf("cell %s not found in %#v", address, cells)
	return StyledCellSnapshot{}
}
