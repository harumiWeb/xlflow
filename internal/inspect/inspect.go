package inspect

import (
	"archive/zip"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Payload struct {
	Target     string           `json:"target"`
	TargetInfo *TargetInfo      `json:"target_info,omitempty"`
	Format     string           `json:"format,omitempty"`
	Source     string           `json:"source,omitempty"`
	Workbook   *WorkbookSummary `json:"workbook,omitempty"`
	Sheets     []SheetSummary   `json:"sheets,omitempty"`
	Range      *RangeSnapshot   `json:"range,omitempty"`
	Cell       *CellSnapshot    `json:"cell,omitempty"`
	Form       any              `json:"form,omitempty"`
	Forms      any              `json:"forms,omitempty"`
}

type TargetInfo struct {
	Kind string `json:"kind"`
	Path string `json:"path,omitempty"`
	Note string `json:"note,omitempty"`
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
	Sheet         string               `json:"sheet"`
	Range         string               `json:"range,omitempty"`
	UsedRange     string               `json:"used_range,omitempty"`
	ReturnedRange string               `json:"returned_range,omitempty"`
	RowCount      int                  `json:"row_count"`
	ColumnCount   int                  `json:"column_count"`
	Values        [][]any              `json:"values"`
	Truncated     bool                 `json:"truncated,omitempty"`
	MaxRows       int                  `json:"max_rows,omitempty"`
	MaxCols       int                  `json:"max_cols,omitempty"`
	Warnings      []string             `json:"warnings,omitempty"`
	StyleIncluded bool                 `json:"style_included,omitempty"`
	Cells         []StyledCellSnapshot `json:"cells,omitempty"`
	Columns       []ColumnSnapshot     `json:"columns,omitempty"`
	Rows          []RowSnapshot        `json:"rows,omitempty"`
	MergedRanges  []string             `json:"merged_ranges,omitempty"`
}

func (snapshot RangeSnapshot) MarshalJSON() ([]byte, error) {
	type alias RangeSnapshot
	if !snapshot.StyleIncluded {
		return json.Marshal(alias(snapshot))
	}
	data, err := json.Marshal(alias(snapshot))
	if err != nil {
		return nil, err
	}
	var payload map[string]any
	if err := json.Unmarshal(data, &payload); err != nil {
		return nil, err
	}
	if _, ok := payload["cells"]; !ok {
		payload["cells"] = []StyledCellSnapshot{}
	}
	if _, ok := payload["columns"]; !ok {
		payload["columns"] = []ColumnSnapshot{}
	}
	if _, ok := payload["rows"]; !ok {
		payload["rows"] = []RowSnapshot{}
	}
	if _, ok := payload["merged_ranges"]; !ok {
		payload["merged_ranges"] = []string{}
	}
	return json.Marshal(payload)
}

type StyledCellSnapshot struct {
	Address             string             `json:"address"`
	Row                 int                `json:"row"`
	Column              int                `json:"column"`
	Value               any                `json:"value"`
	Formula             *string            `json:"formula"`
	Fill                *CellFillSnapshot  `json:"fill"`
	Font                *CellFontSnapshot  `json:"font"`
	Border              CellBorderSnapshot `json:"border"`
	NumberFormat        *string            `json:"number_format"`
	HorizontalAlignment *string            `json:"horizontal_alignment"`
	VerticalAlignment   *string            `json:"vertical_alignment"`
	Merged              bool               `json:"merged"`
	MergeRange          *string            `json:"merge_range"`
}

type CellFillSnapshot struct {
	Type  string  `json:"type"`
	Color *string `json:"color"`
}

type CellFontSnapshot struct {
	Name   string  `json:"name"`
	Size   float64 `json:"size"`
	Bold   bool    `json:"bold"`
	Italic bool    `json:"italic"`
	Color  *string `json:"color"`
}

type CellBorderSnapshot struct {
	Top    BorderEdgeSnapshot `json:"top"`
	Right  BorderEdgeSnapshot `json:"right"`
	Bottom BorderEdgeSnapshot `json:"bottom"`
	Left   BorderEdgeSnapshot `json:"left"`
}

type BorderEdgeSnapshot struct {
	Style string  `json:"style"`
	Color *string `json:"color"`
}

type ColumnSnapshot struct {
	Column string  `json:"column"`
	Index  int     `json:"index"`
	Width  float64 `json:"width"`
	Hidden bool    `json:"hidden"`
}

type RowSnapshot struct {
	Row    int     `json:"row"`
	Height float64 `json:"height"`
	Hidden bool    `json:"hidden"`
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

type RangeOptions struct {
	Limits       Limits
	IncludeStyle bool
}

func SavedFileTargetInfo(path string) *TargetInfo {
	return &TargetInfo{
		Kind: "file",
		Path: path,
		Note: "This command inspected the saved workbook file on disk, not an unsaved live Excel session.",
	}
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

func Range(path, sheet, address string, opts RangeOptions) (RangeSnapshot, error) {
	var result RangeSnapshot
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := requireSheet(f, sheet); err != nil {
			return err
		}
		startCol, startRow, endCol, endRow, normalized, err := parseAddressRange(address)
		if err != nil {
			return err
		}
		result, err = readRangeSnapshot(f, sheet, startCol, startRow, endCol, endRow, normalized, "", opts)
		return err
	})
	return result, err
}

func UsedRange(path, sheet string, opts RangeOptions) (RangeSnapshot, error) {
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
				Sheet:         sheet,
				UsedRange:     "",
				RowCount:      0,
				ColumnCount:   0,
				Values:        [][]any{},
				StyleIncluded: opts.IncludeStyle,
			}
			if opts.IncludeStyle {
				result.Cells = []StyledCellSnapshot{}
				result.Columns = []ColumnSnapshot{}
				result.Rows = []RowSnapshot{}
				result.MergedRanges = []string{}
			}
			return nil
		}
		result, err = readRangeSnapshot(f, sheet, info.StartCol, info.StartRow, info.EndCol, info.EndRow, "", info.Address, opts)
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
	StartCol    int
	StartRow    int
	EndCol      int
	EndRow      int
	RowCount    int
	ColumnCount int
}

func usedRangeInfo(f *excelize.File, sheet string) (rangeInfo, error) {
	xmlPath, err := worksheetXMLPath(f, sheet)
	if err != nil {
		return rangeInfo{}, err
	}
	reader, err := zip.OpenReader(f.Path)
	if err != nil {
		return rangeInfo{}, fmt.Errorf("open workbook archive: %w", err)
	}
	defer func() { _ = reader.Close() }()
	var ws *zip.File
	for _, file := range reader.File {
		if file.Name == xmlPath {
			ws = file
			break
		}
	}
	if ws == nil {
		return rangeInfo{}, fmt.Errorf("worksheet xml %q not found", xmlPath)
	}
	rc, err := ws.Open()
	if err != nil {
		return rangeInfo{}, fmt.Errorf("open worksheet xml %q: %w", xmlPath, err)
	}
	defer func() { _ = rc.Close() }()
	decoder := xml.NewDecoder(rc)
	seen := false
	minRow, minCol := 0, 0
	maxRow, maxCol := 0, 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return rangeInfo{}, fmt.Errorf("scan worksheet %q: %w", sheet, err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "c" {
			continue
		}
		ref := ""
		for _, attr := range start.Attr {
			if attr.Name.Local == "r" {
				ref = attr.Value
				break
			}
		}
		if ref == "" {
			continue
		}
		col, row, err := excelize.CellNameToCoordinates(ref)
		if err != nil {
			return rangeInfo{}, fmt.Errorf("scan worksheet %q cell %q: %w", sheet, ref, err)
		}
		if !seen {
			minRow, maxRow = row, row
			minCol, maxCol = col, col
			seen = true
			continue
		}
		if row < minRow {
			minRow = row
		}
		if row > maxRow {
			maxRow = row
		}
		if col < minCol {
			minCol = col
		}
		if col > maxCol {
			maxCol = col
		}
	}
	if !seen {
		return rangeInfo{}, nil
	}
	address, err := addressFromBounds(minCol, minRow, maxCol, maxRow)
	if err != nil {
		return rangeInfo{}, err
	}
	return rangeInfo{
		Address:     address,
		StartCol:    minCol,
		StartRow:    minRow,
		EndCol:      maxCol,
		EndRow:      maxRow,
		RowCount:    maxRow - minRow + 1,
		ColumnCount: maxCol - minCol + 1,
	}, nil
}

func worksheetXMLPath(f *excelize.File, sheet string) (string, error) {
	if f == nil {
		return "", fmt.Errorf("workbook is nil")
	}
	value := reflect.ValueOf(f).Elem().FieldByName("sheetMap")
	if !value.IsValid() || value.Kind() != reflect.Map {
		return "", fmt.Errorf("workbook does not expose sheet map")
	}
	for _, key := range value.MapKeys() {
		if key.Kind() != reflect.String {
			continue
		}
		if key.String() != sheet {
			continue
		}
		return value.MapIndex(key).String(), nil
	}
	return "", fmt.Errorf("sheet %q not found", sheet)
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

func readRangeSnapshot(f *excelize.File, sheet string, startCol, startRow, endCol, endRow int, normalizedRange, usedRange string, opts RangeOptions) (RangeSnapshot, error) {
	fullRows := endRow - startRow + 1
	fullCols := endCol - startCol + 1
	returnEndCol := endCol
	returnEndRow := endRow
	truncated := false
	if opts.Limits.MaxRows > 0 && fullRows > opts.Limits.MaxRows {
		returnEndRow = startRow + opts.Limits.MaxRows - 1
		truncated = true
	}
	if opts.Limits.MaxCols > 0 && fullCols > opts.Limits.MaxCols {
		returnEndCol = startCol + opts.Limits.MaxCols - 1
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
		MaxRows:       opts.Limits.MaxRows,
		MaxCols:       opts.Limits.MaxCols,
		StyleIncluded: opts.IncludeStyle,
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
	if !opts.IncludeStyle || returnedRange == "" {
		return snapshot, nil
	}
	if err := populateStyleSnapshot(f, sheet, startCol, startRow, returnEndCol, returnEndRow, &snapshot); err != nil {
		return RangeSnapshot{}, err
	}
	return snapshot, nil
}

func populateStyleSnapshot(f *excelize.File, sheet string, startCol, startRow, endCol, endRow int, snapshot *RangeSnapshot) error {
	mergeLookup, mergedRanges, err := mergedCellLookup(f, sheet, startCol, startRow, endCol, endRow)
	if err != nil {
		return err
	}
	snapshot.MergedRanges = mergedRanges
	snapshot.Rows = make([]RowSnapshot, 0, endRow-startRow+1)
	snapshot.Columns = make([]ColumnSnapshot, 0, endCol-startCol+1)
	snapshot.Cells = make([]StyledCellSnapshot, 0, (endRow-startRow+1)*(endCol-startCol+1))

	for row := startRow; row <= endRow; row++ {
		height, err := f.GetRowHeight(sheet, row)
		if err != nil {
			return fmt.Errorf("read row height %s!%d: %w", sheet, row, err)
		}
		hidden, err := f.GetRowVisible(sheet, row)
		if err != nil {
			return fmt.Errorf("read row visibility %s!%d: %w", sheet, row, err)
		}
		snapshot.Rows = append(snapshot.Rows, RowSnapshot{
			Row:    row,
			Height: height,
			Hidden: !hidden,
		})
	}
	for col := startCol; col <= endCol; col++ {
		columnName, err := excelize.ColumnNumberToName(col)
		if err != nil {
			return err
		}
		width, err := f.GetColWidth(sheet, columnName)
		if err != nil {
			return fmt.Errorf("read column width %s!%s: %w", sheet, columnName, err)
		}
		visible, err := f.GetColVisible(sheet, columnName)
		if err != nil {
			return fmt.Errorf("read column visibility %s!%s: %w", sheet, columnName, err)
		}
		snapshot.Columns = append(snapshot.Columns, ColumnSnapshot{
			Column: columnName,
			Index:  col,
			Width:  width,
			Hidden: !visible,
		})
	}
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			cellAddress, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return err
			}
			value, err := f.GetCellValue(sheet, cellAddress)
			if err != nil {
				return fmt.Errorf("read cell %s!%s: %w", sheet, cellAddress, err)
			}
			formula, err := f.GetCellFormula(sheet, cellAddress)
			if err != nil {
				return fmt.Errorf("read formula %s!%s: %w", sheet, cellAddress, err)
			}
			styleID, err := f.GetCellStyle(sheet, cellAddress)
			if err != nil {
				return fmt.Errorf("read style id %s!%s: %w", sheet, cellAddress, err)
			}
			style, err := f.GetStyle(styleID)
			if err != nil {
				return fmt.Errorf("read style %s!%s: %w", sheet, cellAddress, err)
			}
			mergeRange := mergeLookup[cellAddress]
			snapshot.Cells = append(snapshot.Cells, StyledCellSnapshot{
				Address:             cellAddress,
				Row:                 row,
				Column:              col,
				Value:               nullableString(value),
				Formula:             nullableStringPtr(formula),
				Fill:                buildFillSnapshot(style),
				Font:                buildFontSnapshot(style),
				Border:              buildBorderSnapshot(style),
				NumberFormat:        numberFormatString(style),
				HorizontalAlignment: alignmentValue(style, true),
				VerticalAlignment:   alignmentValue(style, false),
				Merged:              mergeRange != "",
				MergeRange:          nullableStringPtr(mergeRange),
			})
		}
	}
	return nil
}

func mergedCellLookup(f *excelize.File, sheet string, startCol, startRow, endCol, endRow int) (map[string]string, []string, error) {
	mergedCells, err := f.GetMergeCells(sheet, true)
	if err != nil {
		return nil, nil, fmt.Errorf("read merged cells for %q: %w", sheet, err)
	}
	lookup := map[string]string{}
	mergedRanges := make([]string, 0, len(mergedCells))
	for _, mergedCell := range mergedCells {
		if len(mergedCell) == 0 {
			continue
		}
		ref := mergedCell[0]
		minCol, minRow, maxCol, maxRow, _, err := parseAddressRange(ref)
		if err != nil {
			return nil, nil, fmt.Errorf("parse merged range %q: %w", ref, err)
		}
		if !boundsOverlap(startCol, startRow, endCol, endRow, minCol, minRow, maxCol, maxRow) {
			continue
		}
		mergedRanges = append(mergedRanges, ref)
		for row := maxInt(startRow, minRow); row <= minInt(endRow, maxRow); row++ {
			for col := maxInt(startCol, minCol); col <= minInt(endCol, maxCol); col++ {
				cellAddress, err := excelize.CoordinatesToCellName(col, row)
				if err != nil {
					return nil, nil, err
				}
				lookup[cellAddress] = ref
			}
		}
	}
	return lookup, mergedRanges, nil
}

func buildFillSnapshot(style *excelize.Style) *CellFillSnapshot {
	if style == nil {
		return nil
	}
	fillType := strings.TrimSpace(style.Fill.Type)
	if strings.EqualFold(fillType, "pattern") {
		fillType = fillTypeFromPattern(style.Fill.Pattern)
	}
	color := firstColor(style.Fill.Color)
	if fillType == "" && color == nil {
		return nil
	}
	if fillType == "" {
		fillType = fillTypeFromPattern(style.Fill.Pattern)
	}
	return &CellFillSnapshot{
		Type:  fillType,
		Color: color,
	}
}

func buildFontSnapshot(style *excelize.Style) *CellFontSnapshot {
	if style == nil || style.Font == nil {
		return nil
	}
	return &CellFontSnapshot{
		Name:   style.Font.Family,
		Size:   style.Font.Size,
		Bold:   style.Font.Bold,
		Italic: style.Font.Italic,
		Color:  normalizeColor(style.Font.Color),
	}
}

func buildBorderSnapshot(style *excelize.Style) CellBorderSnapshot {
	snapshot := CellBorderSnapshot{
		Top:    BorderEdgeSnapshot{Style: "none"},
		Right:  BorderEdgeSnapshot{Style: "none"},
		Bottom: BorderEdgeSnapshot{Style: "none"},
		Left:   BorderEdgeSnapshot{Style: "none"},
	}
	if style == nil {
		return snapshot
	}
	for _, border := range style.Border {
		edge := BorderEdgeSnapshot{
			Style: borderStyleName(border.Style),
			Color: normalizeColor(border.Color),
		}
		switch strings.ToLower(strings.TrimSpace(border.Type)) {
		case "top":
			snapshot.Top = edge
		case "right":
			snapshot.Right = edge
		case "bottom":
			snapshot.Bottom = edge
		case "left":
			snapshot.Left = edge
		}
	}
	return snapshot
}

func numberFormatString(style *excelize.Style) *string {
	if style == nil {
		return nil
	}
	if style.CustomNumFmt != nil {
		return nullableStringPtr(*style.CustomNumFmt)
	}
	if format, ok := builtInNumberFormats[style.NumFmt]; ok {
		return strPtr(format)
	}
	if style.NumFmt == 0 {
		return strPtr("General")
	}
	if style.NumFmt > 0 {
		return strPtr(fmt.Sprintf("builtin:%d", style.NumFmt))
	}
	return nil
}

func alignmentValue(style *excelize.Style, horizontal bool) *string {
	if style == nil || style.Alignment == nil {
		return nil
	}
	if horizontal {
		return nullableStringPtr(style.Alignment.Horizontal)
	}
	return nullableStringPtr(style.Alignment.Vertical)
}

func fillTypeFromPattern(pattern int) string {
	switch pattern {
	case 0:
		return "none"
	case 1:
		return "solid"
	default:
		return fmt.Sprintf("pattern:%d", pattern)
	}
}

func firstColor(colors []string) *string {
	for _, color := range colors {
		if normalized := normalizeColor(color); normalized != nil {
			return normalized
		}
	}
	return nil
}

func normalizeColor(raw string) *string {
	value := strings.TrimSpace(raw)
	if value == "" {
		return nil
	}
	value = strings.TrimPrefix(value, "#")
	value = strings.ToUpper(value)
	switch len(value) {
	case 8:
		value = value[2:]
	case 6:
	default:
		return strPtr("#" + value)
	}
	return strPtr("#" + value)
}

func borderStyleName(style int) string {
	switch style {
	case 0:
		return "none"
	case 1:
		return "thin"
	case 2:
		return "medium"
	case 3:
		return "dashed"
	case 4:
		return "dotted"
	case 5:
		return "thick"
	case 6:
		return "double"
	case 7:
		return "hair"
	case 8:
		return "mediumDashed"
	case 9:
		return "dashDot"
	case 10:
		return "mediumDashDot"
	case 11:
		return "dashDotDot"
	case 12:
		return "mediumDashDotDot"
	case 13:
		return "slantDashDot"
	default:
		return fmt.Sprintf("unknown:%d", style)
	}
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

func boundsOverlap(startColA, startRowA, endColA, endRowA, startColB, startRowB, endColB, endRowB int) bool {
	return startColA <= endColB && endColA >= startColB && startRowA <= endRowB && endRowA >= startRowB
}

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func nullableString(value string) any {
	if value == "" {
		return nil
	}
	return value
}

func nullableStringPtr(value string) *string {
	value = strings.TrimSpace(value)
	if value == "" {
		return nil
	}
	return &value
}

func strPtr(value string) *string {
	return &value
}

var builtInNumberFormats = map[int]string{
	0:  "General",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "mm-dd-yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "hh:mm",
	21: "hh:mm:ss",
	22: "m/d/yy hh:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[Red](#,##0)",
	39: "#,##0.00 ;(#,##0.00)",
	40: "#,##0.00 ;[Red](#,##0.00)",
	41: "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)",
	42: "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
	43: "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
	44: "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mm:ss.0",
	48: "##0.0E+0",
	49: "@",
}
