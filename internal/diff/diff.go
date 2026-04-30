package diff

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Options struct {
	BeforeWorkbook string
	AfterWorkbook  string
	VBABeforeDir   string
	VBAAfterDir    string
}

type WorkbookDiff struct {
	Summary Summary     `json:"summary"`
	Sheets  []SheetDiff `json:"sheets"`
	Cells   []CellDiff  `json:"cells"`
	VBA     []VBADiff   `json:"vba"`
}

type Summary struct {
	SheetDiffs int `json:"sheet_diffs"`
	CellDiffs  int `json:"cell_diffs"`
	VBADiffs   int `json:"vba_diffs"`
	TotalDiffs int `json:"total_diffs"`
}

type SheetDiff struct {
	Name string `json:"name"`
	Kind string `json:"kind"`
}

type CellDiff struct {
	Sheet   string `json:"sheet"`
	Address string `json:"address"`
	Kind    string `json:"kind"`
	Before  string `json:"before"`
	After   string `json:"after"`
}

type VBADiff struct {
	File    string         `json:"file"`
	Kind    string         `json:"kind"`
	Changes []TextLineDiff `json:"changes,omitempty"`
}

type TextLineDiff struct {
	Line   int    `json:"line"`
	Before string `json:"before"`
	After  string `json:"after"`
}

func Compare(opts Options) (WorkbookDiff, error) {
	result := WorkbookDiff{
		Sheets: []SheetDiff{},
		Cells:  []CellDiff{},
		VBA:    []VBADiff{},
	}
	workbookDiff, err := compareWorkbooks(opts.BeforeWorkbook, opts.AfterWorkbook)
	if err != nil {
		return result, err
	}
	result.Sheets = workbookDiff.Sheets
	result.Cells = workbookDiff.Cells

	if opts.VBABeforeDir != "" || opts.VBAAfterDir != "" {
		vbaDiffs, err := compareVBADirs(opts.VBABeforeDir, opts.VBAAfterDir)
		if err != nil {
			return result, err
		}
		result.VBA = vbaDiffs
	}
	result.Summary = Summary{
		SheetDiffs: len(result.Sheets),
		CellDiffs:  len(result.Cells),
		VBADiffs:   len(result.VBA),
	}
	result.Summary.TotalDiffs = result.Summary.SheetDiffs + result.Summary.CellDiffs + result.Summary.VBADiffs
	return result, nil
}

func (d WorkbookDiff) Logs() []string {
	if d.Summary.TotalDiffs == 0 {
		return []string{"no differences found"}
	}
	logs := make([]string, 0, d.Summary.TotalDiffs)
	for _, sheet := range d.Sheets {
		prefix := "+"
		if sheet.Kind == "removed" {
			prefix = "-"
		}
		logs = append(logs, fmt.Sprintf("Sheet: %s %s", prefix, sheet.Name))
	}
	currentSheet := ""
	for _, cell := range d.Cells {
		if cell.Sheet != currentSheet {
			currentSheet = cell.Sheet
			logs = append(logs, "Sheet: "+currentSheet)
		}
		logs = append(logs, fmt.Sprintf("%s %s: %q -> %q", cell.Address, cell.Kind, cell.Before, cell.After))
	}
	for _, vba := range d.VBA {
		logs = append(logs, fmt.Sprintf("%s %s", vba.File, vba.Kind))
	}
	return logs
}

type partialWorkbookDiff struct {
	Sheets []SheetDiff
	Cells  []CellDiff
}

func compareWorkbooks(beforePath, afterPath string) (partialWorkbookDiff, error) {
	result := partialWorkbookDiff{
		Sheets: []SheetDiff{},
		Cells:  []CellDiff{},
	}
	before, err := excelize.OpenFile(beforePath)
	if err != nil {
		return result, fmt.Errorf("open before workbook: %w", err)
	}
	defer func() { _ = before.Close() }()
	after, err := excelize.OpenFile(afterPath)
	if err != nil {
		return result, fmt.Errorf("open after workbook: %w", err)
	}
	defer func() { _ = after.Close() }()

	beforeSheets := sheetSet(before.GetSheetList())
	afterSheets := sheetSet(after.GetSheetList())
	for _, name := range sortedKeys(beforeSheets) {
		if !afterSheets[name] {
			result.Sheets = append(result.Sheets, SheetDiff{Name: name, Kind: "removed"})
		}
	}
	for _, name := range sortedKeys(afterSheets) {
		if !beforeSheets[name] {
			result.Sheets = append(result.Sheets, SheetDiff{Name: name, Kind: "added"})
		}
	}
	for _, name := range sortedKeys(beforeSheets) {
		if afterSheets[name] {
			cells, err := compareSheetCells(before, after, name)
			if err != nil {
				return result, err
			}
			result.Cells = append(result.Cells, cells...)
		}
	}
	return result, nil
}

func compareSheetCells(before, after *excelize.File, sheet string) ([]CellDiff, error) {
	beforeRows, err := before.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("read before sheet %q: %w", sheet, err)
	}
	afterRows, err := after.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("read after sheet %q: %w", sheet, err)
	}
	maxRows := max(len(beforeRows), len(afterRows))
	maxCols := 0
	for _, row := range beforeRows {
		maxCols = max(maxCols, len(row))
	}
	for _, row := range afterRows {
		maxCols = max(maxCols, len(row))
	}

	diffs := []CellDiff{}
	for row := 1; row <= maxRows; row++ {
		for col := 1; col <= maxCols; col++ {
			addr, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return nil, err
			}
			beforeFormula, err := before.GetCellFormula(sheet, addr)
			if err != nil {
				return nil, fmt.Errorf("read before formula %s!%s: %w", sheet, addr, err)
			}
			afterFormula, err := after.GetCellFormula(sheet, addr)
			if err != nil {
				return nil, fmt.Errorf("read after formula %s!%s: %w", sheet, addr, err)
			}
			if beforeFormula != afterFormula {
				diffs = append(diffs, CellDiff{Sheet: sheet, Address: addr, Kind: "formula", Before: beforeFormula, After: afterFormula})
			}
			beforeValue, err := before.GetCellValue(sheet, addr)
			if err != nil {
				return nil, fmt.Errorf("read before value %s!%s: %w", sheet, addr, err)
			}
			afterValue, err := after.GetCellValue(sheet, addr)
			if err != nil {
				return nil, fmt.Errorf("read after value %s!%s: %w", sheet, addr, err)
			}
			if beforeValue != afterValue {
				diffs = append(diffs, CellDiff{Sheet: sheet, Address: addr, Kind: "value", Before: beforeValue, After: afterValue})
			}
		}
	}
	return diffs, nil
}

func compareVBADirs(beforeDir, afterDir string) ([]VBADiff, error) {
	beforeFiles, err := vbaFiles(beforeDir)
	if err != nil {
		return nil, fmt.Errorf("read before VBA dir: %w", err)
	}
	afterFiles, err := vbaFiles(afterDir)
	if err != nil {
		return nil, fmt.Errorf("read after VBA dir: %w", err)
	}
	seen := map[string]bool{}
	for file := range beforeFiles {
		seen[file] = true
	}
	for file := range afterFiles {
		seen[file] = true
	}
	files := sortedKeys(seen)
	diffs := make([]VBADiff, 0)
	for _, file := range files {
		beforeText, beforeOK := beforeFiles[file]
		afterText, afterOK := afterFiles[file]
		switch {
		case !beforeOK:
			diffs = append(diffs, VBADiff{File: file, Kind: "added"})
		case !afterOK:
			diffs = append(diffs, VBADiff{File: file, Kind: "removed"})
		case beforeText != afterText:
			diffs = append(diffs, VBADiff{File: file, Kind: "changed", Changes: lineDiff(beforeText, afterText)})
		}
	}
	return diffs, nil
}

func vbaFiles(root string) (map[string]string, error) {
	files := map[string]string{}
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		switch strings.ToLower(filepath.Ext(path)) {
		case ".bas", ".cls", ".frm":
		default:
			return nil
		}
		rel, err := filepath.Rel(root, path)
		if err != nil {
			return err
		}
		data, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		files[filepath.ToSlash(rel)] = normalizeText(string(data))
		return nil
	})
	return files, err
}

func lineDiff(beforeText, afterText string) []TextLineDiff {
	beforeLines := splitLines(beforeText)
	afterLines := splitLines(afterText)
	maxLines := max(len(beforeLines), len(afterLines))
	changes := make([]TextLineDiff, 0)
	for i := 0; i < maxLines; i++ {
		beforeLine := ""
		if i < len(beforeLines) {
			beforeLine = beforeLines[i]
		}
		afterLine := ""
		if i < len(afterLines) {
			afterLine = afterLines[i]
		}
		if beforeLine != afterLine {
			changes = append(changes, TextLineDiff{Line: i + 1, Before: beforeLine, After: afterLine})
		}
	}
	return changes
}

func normalizeText(s string) string {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	s = strings.ReplaceAll(s, "\r", "\n")
	return s
}

func splitLines(s string) []string {
	s = strings.TrimSuffix(normalizeText(s), "\n")
	if s == "" {
		return []string{}
	}
	return strings.Split(s, "\n")
}

func sheetSet(names []string) map[string]bool {
	set := make(map[string]bool, len(names))
	for _, name := range names {
		set[name] = true
	}
	return set
}

func sortedKeys[T any](m map[string]T) []string {
	keys := make([]string, 0, len(m))
	for key := range m {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}
