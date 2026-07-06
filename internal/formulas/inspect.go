package formulas

import (
	"bufio"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/formula"
	"github.com/xuri/excelize/v2"
)

type InspectArgumentError struct {
	Err error
}

func (e InspectArgumentError) Error() string {
	return e.Err.Error()
}

func (e InspectArgumentError) Unwrap() error {
	return e.Err
}

func IsInspectArgumentError(err error) bool {
	var argErr InspectArgumentError
	return errors.As(err, &argErr)
}

func inspectArgumentErrorf(format string, args ...any) error {
	return InspectArgumentError{Err: fmt.Errorf(format, args...)}
}

type InspectResult struct {
	View            string          `json:"view"`
	Dir             string          `json:"dir"`
	Workbook        string          `json:"workbook,omitempty"`
	Sheets          []SheetSummary  `json:"sheets,omitempty"`
	DefinedNames    []DefinedName   `json:"defined_names,omitempty"`
	Sheet           string          `json:"sheet,omitempty"`
	Regions         []InspectRegion `json:"regions,omitempty"`
	Cell            string          `json:"cell,omitempty"`
	Region          *InspectRegion  `json:"region,omitempty"`
	ExpandedFormula string          `json:"expanded_formula,omitempty"`
	Range           string          `json:"range,omitempty"`
}

type SheetSummary struct {
	Name               string             `json:"name"`
	FormulaRegionCount int                `json:"formula_region_count"`
	FormulaCellCount   int                `json:"formula_cell_count"`
	ParseStatus        ParseStatusSummary `json:"parse_status"`
	Features           []string           `json:"features,omitempty"`
	DependsOnSheets    []string           `json:"depends_on_sheets,omitempty"`
}

type InspectRegion struct {
	Sheet             string   `json:"sheet"`
	Range             string   `json:"range"`
	Kind              string   `json:"kind"`
	FormulaR1C1       string   `json:"formula_r1c1,omitempty"`
	Formula           string   `json:"formula,omitempty"`
	ExampleCell       string   `json:"example_cell,omitempty"`
	ExampleFormula    string   `json:"example_formula,omitempty"`
	Count             int      `json:"count"`
	ParseStatus       string   `json:"parse_status"`
	Features          []string `json:"features,omitempty"`
	Refs              []string `json:"refs,omitempty"`
	DependsOnSheets   []string `json:"depends_on_sheets,omitempty"`
	Functions         []string `json:"functions,omitempty"`
	StorageKinds      []string `json:"storage_kinds,omitempty"`
	StorageGroupCount int      `json:"storage_group_count,omitempty"`

	start cellPoint
	end   cellPoint
}

type snapshot struct {
	manifest Manifest
	names    []DefinedName
	sheets   []snapshotSheet
}

type snapshotSheet struct {
	manifest SheetManifest
	regions  []InspectRegion
}

func InspectSummary(dir string) (InspectResult, error) {
	snap, err := loadSnapshot(dir)
	if err != nil {
		return InspectResult{}, err
	}
	return InspectResult{
		View:         "summary",
		Dir:          dir,
		Workbook:     snap.manifest.Workbook,
		Sheets:       snap.summaries(),
		DefinedNames: snap.names,
	}, nil
}

func InspectSheet(dir, sheetName string) (InspectResult, error) {
	snap, err := loadSnapshot(dir)
	if err != nil {
		return InspectResult{}, err
	}
	sheet, ok := snap.sheetByName(sheetName)
	if !ok {
		return InspectResult{}, fmt.Errorf("formula snapshot sheet not found: %s", sheetName)
	}
	return InspectResult{
		View:     "sheet",
		Dir:      dir,
		Workbook: snap.manifest.Workbook,
		Sheet:    sheet.manifest.Name,
		Regions:  sheet.regions,
	}, nil
}

func InspectCell(dir, selector string) (InspectResult, error) {
	sheetName, address, err := parseSheetAddress(selector, false)
	if err != nil {
		return InspectResult{}, err
	}
	col, row, err := excelize.CellNameToCoordinates(address)
	if err != nil {
		return InspectResult{}, fmt.Errorf("invalid cell address %q: %w", address, err)
	}
	snap, err := loadSnapshot(dir)
	if err != nil {
		return InspectResult{}, err
	}
	sheet, ok := snap.sheetByName(sheetName)
	if !ok {
		return InspectResult{}, fmt.Errorf("formula snapshot sheet not found: %s", sheetName)
	}
	cell := cellPoint{row: row, col: col}
	result := InspectResult{
		View:     "cell",
		Dir:      dir,
		Workbook: snap.manifest.Workbook,
		Cell:     sheet.manifest.Name + "!" + address,
	}
	for _, region := range sheet.regions {
		if !pointInRegion(cell, region) {
			continue
		}
		copy := region
		result.Region = &copy
		if region.FormulaR1C1 != "" {
			if expanded, ok := ExpandR1C1Formula(region.FormulaR1C1, formula.CellRef{Row: row, Col: col}); ok {
				result.ExpandedFormula = expanded
			}
		}
		break
	}
	return result, nil
}

func InspectRange(dir, selector string) (InspectResult, error) {
	sheetName, address, err := parseSheetAddress(selector, true)
	if err != nil {
		return InspectResult{}, err
	}
	start, end, _, ok := rangeInfo(address)
	if !ok {
		return InspectResult{}, fmt.Errorf("invalid range address %q", address)
	}
	snap, err := loadSnapshot(dir)
	if err != nil {
		return InspectResult{}, err
	}
	sheet, ok := snap.sheetByName(sheetName)
	if !ok {
		return InspectResult{}, fmt.Errorf("formula snapshot sheet not found: %s", sheetName)
	}
	result := InspectResult{
		View:     "range",
		Dir:      dir,
		Workbook: snap.manifest.Workbook,
		Range:    sheet.manifest.Name + "!" + address,
	}
	for _, region := range sheet.regions {
		if regionsOverlap(start, end, region.start, region.end) {
			result.Regions = append(result.Regions, region)
		}
	}
	return result, nil
}

func loadSnapshot(dir string) (snapshot, error) {
	manifestPath := filepath.Join(dir, "manifest.json")
	body, err := os.ReadFile(manifestPath)
	if err != nil {
		if os.IsNotExist(err) {
			return snapshot{}, fmt.Errorf("formula snapshot manifest not found: %s", manifestPath)
		}
		return snapshot{}, err
	}
	var manifest Manifest
	if err := json.Unmarshal(body, &manifest); err != nil {
		return snapshot{}, fmt.Errorf("invalid formula snapshot manifest %s: %w", manifestPath, err)
	}
	if manifest.Version != 1 {
		return snapshot{}, fmt.Errorf("unsupported formula snapshot manifest version %d", manifest.Version)
	}
	snap := snapshot{manifest: manifest}
	names, err := readJSONL[DefinedName](filepath.Join(dir, "names.jsonl"), true)
	if err != nil {
		return snapshot{}, err
	}
	snap.names = names
	for _, sheet := range manifest.Sheets {
		path := filepath.Join(dir, filepath.FromSlash(sheet.Path))
		regions, err := readJSONL[FormulaRegion](path, false)
		if err != nil {
			return snapshot{}, err
		}
		inspectRegions := make([]InspectRegion, 0, len(regions))
		for _, region := range regions {
			inspect, err := inspectRegion(sheet.Name, region)
			if err != nil {
				return snapshot{}, fmt.Errorf("invalid formula region in %s: %w", path, err)
			}
			inspectRegions = append(inspectRegions, inspect)
		}
		sort.SliceStable(inspectRegions, func(i, j int) bool {
			if inspectRegions[i].start.row != inspectRegions[j].start.row {
				return inspectRegions[i].start.row < inspectRegions[j].start.row
			}
			if inspectRegions[i].start.col != inspectRegions[j].start.col {
				return inspectRegions[i].start.col < inspectRegions[j].start.col
			}
			return inspectRegions[i].Range < inspectRegions[j].Range
		})
		snap.sheets = append(snap.sheets, snapshotSheet{manifest: sheet, regions: inspectRegions})
	}
	return snap, nil
}

func readJSONL[T any](path string, optional bool) (values []T, err error) {
	file, err := os.Open(path)
	if err != nil {
		if optional && os.IsNotExist(err) {
			return nil, nil
		}
		if os.IsNotExist(err) {
			return nil, fmt.Errorf("formula snapshot file not found: %s", path)
		}
		return nil, err
	}
	defer func() {
		if closeErr := file.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	scanner := bufio.NewScanner(file)
	scanner.Buffer(make([]byte, 64*1024), 8*1024*1024)
	lineNo := 0
	for scanner.Scan() {
		lineNo++
		line := strings.TrimSpace(scanner.Text())
		if line == "" {
			continue
		}
		var value T
		if err := json.Unmarshal([]byte(line), &value); err != nil {
			return nil, fmt.Errorf("invalid JSONL in %s line %d: %w", path, lineNo, err)
		}
		values = append(values, value)
	}
	if err := scanner.Err(); err != nil {
		return nil, err
	}
	return values, nil
}

func inspectRegion(sheet string, region FormulaRegion) (InspectRegion, error) {
	start, end, _, ok := rangeInfo(region.Range)
	if !ok {
		return InspectRegion{}, fmt.Errorf("invalid region range %q", region.Range)
	}
	return InspectRegion{
		Sheet:             sheet,
		Range:             region.Range,
		Kind:              "formula",
		FormulaR1C1:       region.FormulaR1C1,
		Formula:           region.Formula,
		ExampleCell:       region.ExampleCell,
		ExampleFormula:    region.ExampleFormula,
		Count:             region.Count,
		ParseStatus:       region.ParseStatus,
		Features:          region.Features,
		Refs:              region.Refs,
		DependsOnSheets:   region.DependsOnSheets,
		Functions:         region.Functions,
		StorageKinds:      region.StorageKinds,
		StorageGroupCount: region.StorageGroupCount,
		start:             start,
		end:               end,
	}, nil
}

func (s snapshot) summaries() []SheetSummary {
	summaries := make([]SheetSummary, 0, len(s.sheets))
	for _, sheet := range s.sheets {
		features := map[string]bool{}
		deps := map[string]bool{}
		cellCount := 0
		for _, region := range sheet.regions {
			cellCount += region.Count
			for _, feature := range region.Features {
				features[feature] = true
			}
			for _, dep := range region.DependsOnSheets {
				deps[dep] = true
			}
		}
		summaries = append(summaries, SheetSummary{
			Name:               sheet.manifest.Name,
			FormulaRegionCount: len(sheet.regions),
			FormulaCellCount:   cellCount,
			ParseStatus:        sheet.manifest.ParseStatusSummary,
			Features:           sortedKeys(features),
			DependsOnSheets:    sortedKeys(deps),
		})
	}
	return summaries
}

func (s snapshot) sheetByName(name string) (snapshotSheet, bool) {
	for _, sheet := range s.sheets {
		if sheet.manifest.Name == name {
			return sheet, true
		}
	}
	return snapshotSheet{}, false
}

func sortedKeys(values map[string]bool) []string {
	if len(values) == 0 {
		return nil
	}
	result := make([]string, 0, len(values))
	for value := range values {
		if value != "" {
			result = append(result, value)
		}
	}
	sort.Strings(result)
	return result
}

func parseSheetAddress(selector string, allowRange bool) (string, string, error) {
	selector = strings.TrimSpace(selector)
	if selector == "" {
		return "", "", inspectArgumentErrorf("selector is required")
	}
	bang := sheetAddressBang(selector)
	if bang < 0 {
		return "", "", inspectArgumentErrorf("expected selector in the form Sheet!A1")
	}
	sheet := strings.TrimSpace(selector[:bang])
	address := strings.TrimSpace(selector[bang+1:])
	if strings.HasPrefix(sheet, "'") && strings.HasSuffix(sheet, "'") && len(sheet) >= 2 {
		sheet = strings.ReplaceAll(sheet[1:len(sheet)-1], "''", "'")
	}
	if sheet == "" || address == "" {
		return "", "", inspectArgumentErrorf("expected selector in the form Sheet!A1")
	}
	if allowRange {
		if _, _, _, ok := rangeInfo(address); !ok {
			return "", "", inspectArgumentErrorf("invalid range address %q", address)
		}
		return sheet, address, nil
	}
	if strings.Contains(address, ":") {
		return "", "", inspectArgumentErrorf("invalid cell address %q", address)
	}
	if _, _, err := excelize.CellNameToCoordinates(address); err != nil {
		return "", "", inspectArgumentErrorf("invalid cell address %q: %w", address, err)
	}
	return sheet, address, nil
}

func sheetAddressBang(selector string) int {
	inQuote := false
	for i := 0; i < len(selector); i++ {
		switch selector[i] {
		case '\'':
			if inQuote && i+1 < len(selector) && selector[i+1] == '\'' {
				i++
				continue
			}
			inQuote = !inQuote
		case '!':
			if !inQuote {
				return i
			}
		}
	}
	return -1
}

func pointInRegion(point cellPoint, region InspectRegion) bool {
	return point.row >= region.start.row && point.row <= region.end.row &&
		point.col >= region.start.col && point.col <= region.end.col
}

func regionsOverlap(aStart, aEnd, bStart, bEnd cellPoint) bool {
	return aStart.row <= bEnd.row && aEnd.row >= bStart.row &&
		aStart.col <= bEnd.col && aEnd.col >= bStart.col
}

var (
	r1c1CombinedPattern = regexp.MustCompile(`^R(\[[-+]?\d+\]|\d*)C(\[[-+]?\d+\]|\d*)$`)
	r1c1RowThenCPattern = regexp.MustCompile(`^R(\d+)C$`)
	r1c1ColumnPattern   = regexp.MustCompile(`^C(\d+)$`)
	r1c1RowColumnOnly   = regexp.MustCompile(`^[RC]\d+$`)
)

func ExpandR1C1Formula(value string, base formula.CellRef) (string, bool) {
	if base.Row <= 0 || base.Col <= 0 {
		return "", false
	}
	tokens, diagnostics := formula.Lex(value)
	if len(diagnostics) > 0 {
		return "", false
	}
	var b strings.Builder
	for i := 0; i < len(tokens); {
		if replacement, consumed, ok := parseR1C1Reference(tokens, i, base); ok {
			b.WriteString(replacement)
			i += consumed
			continue
		}
		if looksLikeUnsupportedR1C1(tokens, i) {
			return "", false
		}
		b.WriteString(tokens[i].Text)
		i++
	}
	return b.String(), true
}

func parseR1C1Reference(tokens []formula.Token, pos int, base formula.CellRef) (string, int, bool) {
	start := pos
	prefix := ""
	if pos+1 < len(tokens) && tokens[pos+1].Text == "!" &&
		(tokens[pos].Kind == formula.TokenIdentifier || tokens[pos].Kind == formula.TokenQuotedName) {
		prefix = tokens[pos].Text + "!"
		pos += 2
	}
	first, consumed, ok := parseR1C1Endpoint(tokens, pos, base)
	if !ok {
		return "", 0, false
	}
	pos += consumed
	if pos < len(tokens) && tokens[pos].Text == ":" {
		second, secondConsumed, ok := parseR1C1Endpoint(tokens, pos+1, base)
		if !ok {
			return "", 0, false
		}
		renderedFirst, ok := renderExpandedR1C1(first)
		if !ok {
			return "", 0, false
		}
		renderedSecond, ok := renderExpandedR1C1(second)
		if !ok {
			return "", 0, false
		}
		return prefix + renderedFirst + ":" + renderedSecond, pos + 1 + secondConsumed - start, true
	}
	rendered, ok := renderExpandedR1C1(first)
	if !ok {
		return "", 0, false
	}
	return prefix + rendered, pos - start, true
}

type expandedR1C1Endpoint struct {
	row    int
	col    int
	rowAbs bool
	colAbs bool
}

const (
	maxExcelRows = 1048576
	maxExcelCols = 16384
)

func parseR1C1Endpoint(tokens []formula.Token, pos int, base formula.CellRef) (expandedR1C1Endpoint, int, bool) {
	if pos >= len(tokens) {
		return expandedR1C1Endpoint{}, 0, false
	}
	if tokens[pos].Kind == formula.TokenIdentifier {
		if strings.EqualFold(tokens[pos].Text, "RC") {
			col, colAbs, colConsumed, ok := parseSplitR1C1Part(tokens, pos+1, base.Col, maxExcelCols)
			if !ok {
				return expandedR1C1Endpoint{}, 0, false
			}
			return expandedR1C1Endpoint{row: base.Row, col: col, colAbs: colAbs}, colConsumed + 1, true
		}
		if match := r1c1RowThenCPattern.FindStringSubmatch(strings.ToUpper(tokens[pos].Text)); match != nil {
			row, rowAbs, ok := resolveR1C1Part(match[1], base.Row, maxExcelRows)
			if !ok {
				return expandedR1C1Endpoint{}, 0, false
			}
			col, colAbs, colConsumed, ok := parseSplitR1C1Part(tokens, pos+1, base.Col, maxExcelCols)
			if !ok {
				return expandedR1C1Endpoint{}, 0, false
			}
			return expandedR1C1Endpoint{row: row, col: col, rowAbs: rowAbs, colAbs: colAbs}, colConsumed + 1, true
		}
		if match := r1c1CombinedPattern.FindStringSubmatch(strings.ToUpper(tokens[pos].Text)); match != nil {
			row, rowAbs, ok := resolveR1C1Part(match[1], base.Row, maxExcelRows)
			if !ok {
				return expandedR1C1Endpoint{}, 0, false
			}
			col, colAbs, ok := resolveR1C1Part(match[2], base.Col, maxExcelCols)
			if !ok {
				return expandedR1C1Endpoint{}, 0, false
			}
			return expandedR1C1Endpoint{row: row, col: col, rowAbs: rowAbs, colAbs: colAbs}, 1, true
		}
		if strings.EqualFold(tokens[pos].Text, "R") {
			row, rowAbs, consumed, ok := parseSplitR1C1Part(tokens, pos+1, base.Row, maxExcelRows)
			if !ok {
				return expandedR1C1Endpoint{}, 0, false
			}
			col, colAbs, colConsumed, ok := parseR1C1Column(tokens, pos+consumed+1, base.Col)
			if !ok {
				return expandedR1C1Endpoint{}, 0, false
			}
			return expandedR1C1Endpoint{row: row, col: col, rowAbs: rowAbs, colAbs: colAbs}, consumed + colConsumed + 1, true
		}
	}
	return expandedR1C1Endpoint{}, 0, false
}

func parseR1C1Column(tokens []formula.Token, pos int, base int) (int, bool, int, bool) {
	if pos >= len(tokens) || tokens[pos].Kind != formula.TokenIdentifier {
		return 0, false, 0, false
	}
	if strings.EqualFold(tokens[pos].Text, "C") {
		col, colAbs, consumed, ok := parseSplitR1C1Part(tokens, pos+1, base, maxExcelCols)
		return col, colAbs, consumed + 1, ok
	}
	if match := r1c1ColumnPattern.FindStringSubmatch(strings.ToUpper(tokens[pos].Text)); match != nil {
		col, colAbs, ok := resolveR1C1Part(match[1], base, maxExcelCols)
		return col, colAbs, 1, ok
	}
	return 0, false, 0, false
}

func parseSplitR1C1Part(tokens []formula.Token, pos int, base int, maxValue int) (int, bool, int, bool) {
	if pos >= len(tokens) || tokens[pos].Text == ":" || tokens[pos].Text == ")" || tokens[pos].Text == "," {
		return base, false, 0, true
	}
	if tokens[pos].Kind == formula.TokenNumber {
		value, err := strconv.Atoi(tokens[pos].Text)
		if err != nil || value <= 0 || value > maxValue {
			return 0, false, 0, false
		}
		return value, true, 1, true
	}
	if tokens[pos].Text != "[" {
		return base, false, 0, true
	}
	if pos+2 >= len(tokens) {
		return 0, false, 0, false
	}
	offsetPos := pos + 1
	offsetText := tokens[offsetPos].Text
	consumed := 3
	if (offsetText == "-" || offsetText == "+") && pos+3 < len(tokens) {
		offsetText += tokens[pos+2].Text
		consumed = 4
	}
	if pos+consumed-1 >= len(tokens) || tokens[pos+consumed-1].Text != "]" {
		return 0, false, 0, false
	}
	offset, err := strconv.Atoi(offsetText)
	if err != nil {
		return 0, false, 0, false
	}
	value := base + offset
	if value <= 0 || value > maxValue {
		return 0, false, 0, false
	}
	return value, false, consumed, true
}

func resolveR1C1Part(text string, base int, maxValue int) (int, bool, bool) {
	if text == "" {
		return base, false, true
	}
	if strings.HasPrefix(text, "[") && strings.HasSuffix(text, "]") {
		offset, err := strconv.Atoi(text[1 : len(text)-1])
		if err != nil {
			return 0, false, false
		}
		value := base + offset
		if value <= 0 || value > maxValue {
			return 0, false, false
		}
		return value, false, true
	}
	value, err := strconv.Atoi(text)
	if err != nil || value <= 0 || value > maxValue {
		return 0, false, false
	}
	return value, true, true
}

func looksLikeUnsupportedR1C1(tokens []formula.Token, pos int) bool {
	if pos >= len(tokens) || tokens[pos].Kind != formula.TokenIdentifier {
		return false
	}
	text := strings.ToUpper(tokens[pos].Text)
	switch {
	case text == "R" || text == "C":
		return nextTokenStartsR1C1Part(tokens, pos+1)
	case strings.HasPrefix(text, "R[") || strings.HasPrefix(text, "C["):
		return true
	case r1c1CombinedPattern.MatchString(text) || r1c1RowThenCPattern.MatchString(text):
		return true
	case r1c1RowColumnOnly.MatchString(text):
		return true
	default:
		return false
	}
}

func nextTokenStartsR1C1Part(tokens []formula.Token, pos int) bool {
	if pos >= len(tokens) {
		return false
	}
	return tokens[pos].Text == "[" || tokens[pos].Kind == formula.TokenNumber || tokens[pos].Text == ":"
}

func renderExpandedR1C1(endpoint expandedR1C1Endpoint) (string, bool) {
	if endpoint.row <= 0 || endpoint.row > maxExcelRows || endpoint.col <= 0 || endpoint.col > maxExcelCols {
		return "", false
	}
	col, err := excelize.ColumnNumberToName(endpoint.col)
	if err != nil {
		return "", false
	}
	if endpoint.colAbs {
		col = "$" + col
	}
	row := strconv.Itoa(endpoint.row)
	if endpoint.rowAbs {
		row = "$" + row
	}
	return col + row, true
}
