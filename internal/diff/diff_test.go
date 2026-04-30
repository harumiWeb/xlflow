package diff

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCompareDetectsWorkbookDiffs(t *testing.T) {
	dir := t.TempDir()
	beforePath := filepath.Join(dir, "before.xlsx")
	afterPath := filepath.Join(dir, "after.xlsx")

	before := excelize.NewFile()
	mustSetCellValue(t, before, "Sheet1", "A1", "old")
	mustSetCellFormula(t, before, "Sheet1", "B1", "SUM(1,1)")
	if err := before.SaveAs(beforePath); err != nil {
		t.Fatal(err)
	}

	after := excelize.NewFile()
	mustSetCellValue(t, after, "Sheet1", "A1", "new")
	mustSetCellFormula(t, after, "Sheet1", "B1", "SUM(2,2)")
	if _, err := after.NewSheet("Added"); err != nil {
		t.Fatal(err)
	}
	if err := after.SaveAs(afterPath); err != nil {
		t.Fatal(err)
	}

	got, err := Compare(Options{BeforeWorkbook: beforePath, AfterWorkbook: afterPath})
	if err != nil {
		t.Fatal(err)
	}
	if got.Summary.SheetDiffs != 1 || got.Sheets[0] != (SheetDiff{Name: "Added", Kind: "added"}) {
		t.Fatalf("sheet diffs = %#v", got.Sheets)
	}
	if !hasCellDiff(got.Cells, "Sheet1", "A1", "value", "old", "new") {
		t.Fatalf("missing A1 value diff: %#v", got.Cells)
	}
	if !hasCellDiff(got.Cells, "Sheet1", "B1", "formula", "SUM(1,1)", "SUM(2,2)") {
		t.Fatalf("missing B1 formula diff: %#v", got.Cells)
	}
}

func TestCompareIgnoresTrailingBlankCells(t *testing.T) {
	dir := t.TempDir()
	beforePath := filepath.Join(dir, "before.xlsx")
	afterPath := filepath.Join(dir, "after.xlsx")

	before := excelize.NewFile()
	mustSetCellValue(t, before, "Sheet1", "A1", "same")
	if err := before.SaveAs(beforePath); err != nil {
		t.Fatal(err)
	}
	after := excelize.NewFile()
	mustSetCellValue(t, after, "Sheet1", "A1", "same")
	if err := after.SaveAs(afterPath); err != nil {
		t.Fatal(err)
	}

	got, err := Compare(Options{BeforeWorkbook: beforePath, AfterWorkbook: afterPath})
	if err != nil {
		t.Fatal(err)
	}
	if got.Summary.TotalDiffs != 0 {
		t.Fatalf("diffs = %#v", got)
	}
}

func TestCompareDetectsVBADiffsAndIgnoresNonSource(t *testing.T) {
	dir := t.TempDir()
	beforeWorkbook := filepath.Join(dir, "before.xlsx")
	afterWorkbook := filepath.Join(dir, "after.xlsx")
	saveEmptyWorkbook(t, beforeWorkbook)
	saveEmptyWorkbook(t, afterWorkbook)
	beforeDir := filepath.Join(dir, "before-src")
	afterDir := filepath.Join(dir, "after-src")
	mustWrite(t, filepath.Join(beforeDir, "Module1.bas"), "Sub Main()\r\nEnd Sub\r\n")
	mustWrite(t, filepath.Join(afterDir, "Module1.bas"), "Sub Main()\nDebug.Print 1\nEnd Sub\n")
	mustWrite(t, filepath.Join(afterDir, "Class1.cls"), "VERSION 1.0 CLASS\n")
	mustWrite(t, filepath.Join(afterDir, "UserForm1.frx"), "ignored")

	got, err := Compare(Options{
		BeforeWorkbook: beforeWorkbook,
		AfterWorkbook:  afterWorkbook,
		VBABeforeDir:   beforeDir,
		VBAAfterDir:    afterDir,
	})
	if err != nil {
		t.Fatal(err)
	}
	if got.Summary.VBADiffs != 2 {
		t.Fatalf("vba diffs = %#v", got.VBA)
	}
	if got.VBA[0].File != "Class1.cls" || got.VBA[0].Kind != "added" {
		t.Fatalf("first vba diff = %#v", got.VBA[0])
	}
	if got.VBA[1].File != "Module1.bas" || got.VBA[1].Kind != "changed" {
		t.Fatalf("second vba diff = %#v", got.VBA[1])
	}
	if len(got.VBA[1].Changes) == 0 {
		t.Fatalf("expected line changes: %#v", got.VBA[1])
	}
}

func TestCompareNormalizesVBALineEndings(t *testing.T) {
	dir := t.TempDir()
	beforeWorkbook := filepath.Join(dir, "before.xlsx")
	afterWorkbook := filepath.Join(dir, "after.xlsx")
	saveEmptyWorkbook(t, beforeWorkbook)
	saveEmptyWorkbook(t, afterWorkbook)
	beforeDir := filepath.Join(dir, "before-src")
	afterDir := filepath.Join(dir, "after-src")
	mustWrite(t, filepath.Join(beforeDir, "Module1.bas"), "Sub Main()\r\nEnd Sub\r\n")
	mustWrite(t, filepath.Join(afterDir, "Module1.bas"), "Sub Main()\nEnd Sub\n")

	got, err := Compare(Options{
		BeforeWorkbook: beforeWorkbook,
		AfterWorkbook:  afterWorkbook,
		VBABeforeDir:   beforeDir,
		VBAAfterDir:    afterDir,
	})
	if err != nil {
		t.Fatal(err)
	}
	if got.Summary.VBADiffs != 0 {
		t.Fatalf("vba diffs = %#v", got.VBA)
	}
}

func TestLogs(t *testing.T) {
	got := WorkbookDiff{}.Logs()
	if len(got) != 1 || got[0] != "no differences found" {
		t.Fatalf("logs = %#v", got)
	}
}

func mustSetCellValue(t *testing.T, f *excelize.File, sheet, cell string, value any) {
	t.Helper()
	if err := f.SetCellValue(sheet, cell, value); err != nil {
		t.Fatal(err)
	}
}

func mustSetCellFormula(t *testing.T, f *excelize.File, sheet, cell, formula string) {
	t.Helper()
	if err := f.SetCellFormula(sheet, cell, formula); err != nil {
		t.Fatal(err)
	}
}

func saveEmptyWorkbook(t *testing.T, path string) {
	t.Helper()
	f := excelize.NewFile()
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
}

func mustWrite(t *testing.T, path, content string) {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(content), 0o644); err != nil {
		t.Fatal(err)
	}
}

func hasCellDiff(diffs []CellDiff, sheet, address, kind, before, after string) bool {
	for _, diff := range diffs {
		if diff.Sheet == sheet && diff.Address == address && diff.Kind == kind && diff.Before == before && diff.After == after {
			return true
		}
	}
	return false
}
