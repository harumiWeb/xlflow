package inspect

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Payload struct {
	Target   string           `json:"target"`
	Format   string           `json:"format,omitempty"`
	Source   string           `json:"source,omitempty"`
	Workbook *WorkbookSummary `json:"workbook,omitempty"`
	Sheets   []SheetSummary   `json:"sheets,omitempty"`
	Range    *RangeSnapshot   `json:"range,omitempty"`
	Cell     *CellSnapshot    `json:"cell,omitempty"`
}

type WorkbookSummary struct {
	Path        string         `json:"path"`
	Name        string         `json:"name"`
	Sheets      []SheetSummary `json:"sheets"`
	ActiveSheet string         `json:"active_sheet,omitempty"`
}

type SheetSummary struct {
	Name        string `json:"name"`
	Index       int    `json:"index"`
	Visible     bool   `json:"visible"`
	UsedRange   string `json:"used_range,omitempty"`
	RowCount    int    `json:"row_count"`
	ColumnCount int    `json:"column_count"`
}

type RangeSnapshot struct {
	Sheet         string   `json:"sheet"`
	Range         string   `json:"range,omitempty"`
	UsedRange     string   `json:"used_range,omitempty"`
	ReturnedRange string   `json:"returned_range,omitempty"`
	RowCount      int      `json:"row_count"`
	ColumnCount   int      `json:"column_count"`
	Values        [][]any  `json:"values,omitempty"`
	Truncated     bool     `json:"truncated,omitempty"`
	MaxRows       int      `json:"max_rows,omitempty"`
	MaxCols       int      `json:"max_cols,omitempty"`
	Warnings      []string `json:"warnings,omitempty"`
}

type CellSnapshot struct {
	Sheet   string `json:"sheet"`
	Address string `json:"address"`
	Value   any    `json:"value"`
}

type Limits struct {
	MaxRows int
	MaxCols int
}

func Workbook(path string) (WorkbookSummary, error) {
	result := WorkbookSummary{
		Path:   path,
		Name:   filepath.Base(path),
		Sheets: []SheetSummary{},
	}
	err := withWorkbook(path, func(f *excelize.File) error {
		sheets, err := sheetSummaries(f)
		if err != nil {
			return err
		}
		result.Sheets = sheets
		result.ActiveSheet = activeSheetName(f, sheets)
		return nil
	})
	return result, err
}

func Sheets(path string) ([]SheetSummary, error) {
	var result []SheetSummary
	err := withWorkbook(path, func(f *excelize.File) error {
		sheets, err := sheetSummaries(f)
		if err != nil {
			return err
		}
		result = sheets
		return nil
	})
	return result, err
}

func Range(path, sheet, address string, limits Limits) (RangeSnapshot, error) {
	var result RangeSnapshot
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := requireSheet(f, sheet); err != nil {
			return err
		}
		startCol, startRow, endCol, endRow, normalized, err := parseAddressRange(address)
		if err != nil {
			return err
		}
		result, err = readRangeSnapshot(f, sheet, startCol, startRow, endCol, endRow, normalized, "", limits)
		return err
	})
	return result, err
}

func UsedRange(path, sheet string, limits Limits) (RangeSnapshot, error) {
	var result RangeSnapshot
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := requireSheet(f, sheet); err != nil {
			return err
		}
		info, err := usedRangeInfo(f, sheet)
		if err != nil {
			return err
		}
		if info.RowCount == 0 || info.ColumnCount == 0 {
			result = RangeSnapshot{
				Sheet:       sheet,
				UsedRange:   "",
				RowCount:    0,
				ColumnCount: 0,
				Values:      [][]any{},
			}
			return nil
		}
		result, err = readRangeSnapshot(f, sheet, 1, 1, info.ColumnCount, info.RowCount, "", info.Address, limits)
		return err
	})
	return result, err
}

func Cell(path, sheet, address string) (CellSnapshot, error) {
	var result CellSnapshot
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := requireSheet(f, sheet); err != nil {
			return err
		}
		normalized, err := parseSingleCell(address)
		if err != nil {
			return err
		}
		value, err := f.GetCellValue(sheet, normalized)
		if err != nil {
			return fmt.Errorf("read cell %s!%s: %w", sheet, normalized, err)
		}
		result = CellSnapshot{
			Sheet:   sheet,
			Address: normalized,
			Value:   nullableString(value),
		}
		return nil
	})
	return result, err
}

func withWorkbook(path string, fn func(*excelize.File) error) error {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return fmt.Errorf("open workbook: %w", err)
	}
	defer func() { _ = f.Close() }()
	return fn(f)
}

func sheetSummaries(f *excelize.File) ([]SheetSummary, error) {
	names := f.GetSheetList()
	sheets := make([]SheetSummary, 0, len(names))
	for index, name := range names {
		info, err := usedRangeInfo(f, name)
		if err != nil {
			return nil, err
		}
		visible, err := f.GetSheetVisible(name)
		if err != nil {
			return nil, fmt.Errorf("read worksheet visibility %q: %w", name, err)
		}
		sheets = append(sheets, SheetSummary{
			Name:        name,
			Index:       index + 1,
			Visible:     visible,
			UsedRange:   info.Address,
			RowCount:    info.RowCount,
			ColumnCount: info.ColumnCount,
		})
	}
	return sheets, nil
}

func activeSheetName(f *excelize.File, sheets []SheetSummary) string {
	index := f.GetActiveSheetIndex()
	if name := f.GetSheetName(index); name != "" {
		return name
	}
	if index >= 0 && index < len(sheets) {
		return sheets[index].Name
	}
	if len(sheets) > 0 {
		return sheets[0].Name
	}
	return ""
}

type rangeInfo struct {
	Address     string
	RowCount    int
	ColumnCount int
}

func usedRangeInfo(f *excelize.File, sheet string) (rangeInfo, error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return rangeInfo{}, fmt.Errorf("read worksheet %q: %w", sheet, err)
	}
	rowCount := len(rows)
	colCount := 0
	for _, row := range rows {
		if len(row) > colCount {
			colCount = len(row)
		}
	}
	info := rangeInfo{
		RowCount:    rowCount,
		ColumnCount: colCount,
	}
	if rowCount == 0 || colCount == 0 {
		return info, nil
	}
	address, err := addressFromBounds(1, 1, colCount, rowCount)
	if err != nil {
		return rangeInfo{}, err
	}
	info.Address = address
	return info, nil
}

func requireSheet(f *excelize.File, sheet string) error {
	index, err := f.GetSheetIndex(sheet)
	if err != nil {
		return fmt.Errorf("read worksheet %q: %w", sheet, err)
	}
	if index < 0 {
		return fmt.Errorf("sheet %q not found", sheet)
	}
	return nil
}

func parseAddressRange(address string) (startCol, startRow, endCol, endRow int, normalized string, err error) {
	parts := strings.SplitN(strings.TrimSpace(address), ":", 2)
	if len(parts) == 0 || strings.TrimSpace(parts[0]) == "" {
		return 0, 0, 0, 0, "", fmt.Errorf("address is required")
	}
	first, err := parseSingleCell(parts[0])
	if err != nil {
		return 0, 0, 0, 0, "", err
	}
	last := first
	if len(parts) == 2 {
		last, err = parseSingleCell(parts[1])
		if err != nil {
			return 0, 0, 0, 0, "", err
		}
	}
	startCol, startRow, err = excelize.CellNameToCoordinates(first)
	if err != nil {
		return 0, 0, 0, 0, "", fmt.Errorf("invalid address %q: %w", address, err)
	}
	endCol, endRow, err = excelize.CellNameToCoordinates(last)
	if err != nil {
		return 0, 0, 0, 0, "", fmt.Errorf("invalid address %q: %w", address, err)
	}
	if endCol < startCol {
		startCol, endCol = endCol, startCol
	}
	if endRow < startRow {
		startRow, endRow = endRow, startRow
	}
	normalized, err = addressFromBounds(startCol, startRow, endCol, endRow)
	if err != nil {
		return 0, 0, 0, 0, "", err
	}
	return startCol, startRow, endCol, endRow, normalized, nil
}

func parseSingleCell(address string) (string, error) {
	clean := strings.ToUpper(strings.TrimSpace(strings.ReplaceAll(address, "$", "")))
	if clean == "" {
		return "", fmt.Errorf("address is required")
	}
	if strings.Contains(clean, ":") {
		return "", fmt.Errorf("expected a single cell address, got %q", address)
	}
	if _, _, err := excelize.CellNameToCoordinates(clean); err != nil {
		return "", fmt.Errorf("invalid address %q: %w", address, err)
	}
	return clean, nil
}

func readRangeSnapshot(f *excelize.File, sheet string, startCol, startRow, endCol, endRow int, normalizedRange, usedRange string, limits Limits) (RangeSnapshot, error) {
	fullRows := endRow - startRow + 1
	fullCols := endCol - startCol + 1
	returnEndCol := endCol
	returnEndRow := endRow
	truncated := false
	if limits.MaxRows > 0 && fullRows > limits.MaxRows {
		returnEndRow = startRow + limits.MaxRows - 1
		truncated = true
	}
	if limits.MaxCols > 0 && fullCols > limits.MaxCols {
		returnEndCol = startCol + limits.MaxCols - 1
		truncated = true
	}
	values := make([][]any, 0, returnEndRow-startRow+1)
	for row := startRow; row <= returnEndRow; row++ {
		line := make([]any, 0, returnEndCol-startCol+1)
		for col := startCol; col <= returnEndCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return RangeSnapshot{}, err
			}
			value, err := f.GetCellValue(sheet, cell)
			if err != nil {
				return RangeSnapshot{}, fmt.Errorf("read cell %s!%s: %w", sheet, cell, err)
			}
			line = append(line, nullableString(value))
		}
		values = append(values, line)
	}
	returnedRange := ""
	if len(values) > 0 && len(values[0]) > 0 {
		rangeAddress, err := addressFromBounds(startCol, startRow, returnEndCol, returnEndRow)
		if err != nil {
			return RangeSnapshot{}, err
		}
		returnedRange = rangeAddress
	}
	snapshot := RangeSnapshot{
		Sheet:         sheet,
		Range:         normalizedRange,
		UsedRange:     usedRange,
		ReturnedRange: returnedRange,
		RowCount:      fullRows,
		ColumnCount:   fullCols,
		Values:        values,
		Truncated:     truncated,
		MaxRows:       limits.MaxRows,
		MaxCols:       limits.MaxCols,
	}
	if truncated {
		snapshot.Warnings = []string{
			fmt.Sprintf(
				"Output was truncated: selection has %d row(s) x %d column(s), returned %d row(s) x %d column(s).",
				fullRows,
				fullCols,
				returnEndRow-startRow+1,
				returnEndCol-startCol+1,
			),
		}
	}
	return snapshot, nil
}

func addressFromBounds(startCol, startRow, endCol, endRow int) (string, error) {
	start, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return "", err
	}
	end, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return "", err
	}
	if start == end {
		return start, nil
	}
	return start + ":" + end, nil
}

func nullableString(value string) any {
	if value == "" {
		return nil
	}
	return value
}
