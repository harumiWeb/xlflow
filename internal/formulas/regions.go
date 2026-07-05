package formulas

import (
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/formula"
	"github.com/xuri/excelize/v2"
)

func BuildRegions(cells []FormulaCell) []FormulaRegion {
	if len(cells) == 0 {
		return nil
	}
	sortCells(cells)
	sharedAnchors := map[string]FormulaCell{}
	sharedCovered := map[string]bool{}
	var regions []FormulaRegion
	var normal []FormulaCell

	for _, cell := range cells {
		if cell.Kind == "shared" && cell.Formula != "" && cell.SharedIndex != "" {
			sharedAnchors[cell.SharedIndex] = cell
		}
	}

	for _, cell := range cells {
		if cell.Kind != "shared" {
			normal = append(normal, cell)
			continue
		}
		if cell.Formula != "" && cell.SharedRef != "" {
			region := buildSharedRegion(cell)
			regions = append(regions, region)
			if region.ParseStatus == string(formula.ParseStatusOK) || region.ParseStatus == string(formula.ParseStatusPartial) {
				sharedCovered[cell.SharedIndex] = true
			}
			continue
		}
		if cell.Formula != "" && cell.SharedRef == "" {
			regions = append(regions, failedSharedAnchorRegion(cell, "shared_formula_malformed_ref"))
			continue
		}
		if cell.SharedIndex != "" && sharedCovered[cell.SharedIndex] {
			continue
		}
		if cell.Formula == "" {
			if _, ok := sharedAnchors[cell.SharedIndex]; ok {
				continue
			}
			regions = append(regions, failedSingleCellRegion(cell, "shared_formula_missing_anchor"))
			continue
		}
		normal = append(normal, cell)
	}

	regions = append(regions, buildNormalRegions(normal)...)
	sortRegions(regions)
	return regions
}

func buildSharedRegion(cell FormulaCell) FormulaRegion {
	start, end, count, ok := rangeInfo(cell.SharedRef)
	if !ok {
		r := failedSingleCellRegion(cell, "shared_formula_malformed_ref")
		r.Kind = "shared"
		r.SharedIndex = cell.SharedIndex
		r.Anchor = cell.Cell
		return r
	}
	summary := normalizeFormula(cell.Formula, cell.Row, cell.Col)
	region := FormulaRegion{
		Range:          cell.SharedRef,
		Kind:           "shared",
		SharedIndex:    cell.SharedIndex,
		Anchor:         cell.Cell,
		ExampleCell:    cell.Cell,
		ExampleFormula: cell.Formula,
		Count:          count,
		ParseStatus:    string(summary.status),
		Features:       summary.features,
		startRow:       start.row,
		endRow:         end.row,
		col:            start.col,
	}
	if summary.status == formula.ParseStatusOK {
		region.FormulaR1C1 = summary.formulaR1C1
	} else {
		region.Formula = summary.rawFormula
	}
	region.key = regionKey{Kind: region.Kind, FormulaR1C1: region.FormulaR1C1, Formula: region.Formula, ParseStatus: region.ParseStatus}
	return region
}

func buildNormalRegions(cells []FormulaCell) []FormulaRegion {
	if len(cells) == 0 {
		return nil
	}
	sortCellsByColumn(cells)
	regions := []FormulaRegion{}
	var current *FormulaRegion
	for _, cell := range cells {
		region := singleNormalRegion(cell)
		if current != nil && canExtend(*current, region) {
			current.endRow = region.endRow
			current.Count++
			current.Range = renderRange(current.col, current.startRow, current.col, current.endRow)
			continue
		}
		if current != nil {
			regions = append(regions, *current)
		}
		current = &region
	}
	if current != nil {
		regions = append(regions, *current)
	}
	markOutliers(regions)
	return regions
}

func singleNormalRegion(cell FormulaCell) FormulaRegion {
	summary := normalizeFormula(cell.Formula, cell.Row, cell.Col)
	region := FormulaRegion{
		Range:          cell.Cell,
		Kind:           "normal",
		ExampleCell:    cell.Cell,
		ExampleFormula: cell.Formula,
		Count:          1,
		ParseStatus:    string(summary.status),
		Features:       summary.features,
		startRow:       cell.Row,
		endRow:         cell.Row,
		col:            cell.Col,
	}
	if summary.status == formula.ParseStatusOK {
		region.FormulaR1C1 = summary.formulaR1C1
	} else {
		region.Formula = summary.rawFormula
	}
	region.key = regionKey{Kind: region.Kind, FormulaR1C1: region.FormulaR1C1, Formula: region.Formula, ParseStatus: region.ParseStatus}
	return region
}

func failedSingleCellRegion(cell FormulaCell, feature string) FormulaRegion {
	features := []string{feature}
	return FormulaRegion{
		Range:          cell.Cell,
		Kind:           cell.Kind,
		Formula:        cell.Formula,
		ExampleCell:    cell.Cell,
		ExampleFormula: cell.Formula,
		Count:          1,
		ParseStatus:    string(formula.ParseStatusFailed),
		Features:       features,
		startRow:       cell.Row,
		endRow:         cell.Row,
		col:            cell.Col,
		key:            regionKey{Kind: cell.Kind, Formula: cell.Formula, ParseStatus: string(formula.ParseStatusFailed)},
	}
}

func failedSharedAnchorRegion(cell FormulaCell, feature string) FormulaRegion {
	region := failedSingleCellRegion(cell, feature)
	region.Kind = "shared"
	region.SharedIndex = cell.SharedIndex
	region.Anchor = cell.Cell
	region.key.Kind = "shared"
	return region
}

func normalizeFormula(value string, row, col int) normalizeSummary {
	if row <= 0 || col <= 0 {
		return normalizeSummary{
			status:     formula.ParseStatusFailed,
			rawFormula: value,
		}
	}
	result := formula.NormalizeA1ToR1C1Pattern(value, formula.NormalizeOptions{BaseCell: formula.CellRef{Row: row, Col: col}})
	return normalizeSummary{
		status:      result.Status,
		formulaR1C1: result.Formula,
		rawFormula:  value,
		features:    featureCodes(result.Features),
	}
}

func featureCodes(features []formula.Feature) []string {
	if len(features) == 0 {
		return nil
	}
	seen := map[string]bool{}
	var codes []string
	for _, feature := range features {
		if feature.Code == "" || seen[feature.Code] {
			continue
		}
		seen[feature.Code] = true
		codes = append(codes, feature.Code)
	}
	sort.Strings(codes)
	return codes
}

func canExtend(left, right FormulaRegion) bool {
	return left.Kind == "normal" &&
		right.Kind == "normal" &&
		left.col == right.col &&
		left.endRow+1 == right.startRow &&
		left.key == right.key
}

func markOutliers(regions []FormulaRegion) {
	for i := 1; i+1 < len(regions); i++ {
		prev, current, next := regions[i-1], &regions[i], regions[i+1]
		if current.Count != 1 || current.Kind != "normal" || prev.Kind != "normal" || next.Kind != "normal" {
			continue
		}
		if prev.col != current.col || next.col != current.col || prev.Count <= 1 || next.Count <= 1 {
			continue
		}
		if prev.key == next.key && current.key != prev.key {
			current.Features = appendUnique(current.Features, "outlier")
			sort.Strings(current.Features)
		}
	}
}

func appendUnique(values []string, value string) []string {
	for _, existing := range values {
		if existing == value {
			return values
		}
	}
	return append(values, value)
}

type cellPoint struct {
	row int
	col int
}

func rangeInfo(value string) (cellPoint, cellPoint, int, bool) {
	parts := strings.Split(value, ":")
	if len(parts) == 1 {
		col, row, err := excelize.CellNameToCoordinates(parts[0])
		if err != nil {
			return cellPoint{}, cellPoint{}, 0, false
		}
		return cellPoint{row: row, col: col}, cellPoint{row: row, col: col}, 1, true
	}
	if len(parts) != 2 {
		return cellPoint{}, cellPoint{}, 0, false
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(parts[0])
	if err != nil {
		return cellPoint{}, cellPoint{}, 0, false
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(parts[1])
	if err != nil {
		return cellPoint{}, cellPoint{}, 0, false
	}
	if endCol < startCol || endRow < startRow {
		return cellPoint{}, cellPoint{}, 0, false
	}
	count := (endCol - startCol + 1) * (endRow - startRow + 1)
	return cellPoint{row: startRow, col: startCol}, cellPoint{row: endRow, col: endCol}, count, true
}

func renderRange(startCol, startRow, endCol, endRow int) string {
	start, _ := excelize.CoordinatesToCellName(startCol, startRow)
	end, _ := excelize.CoordinatesToCellName(endCol, endRow)
	if start == end {
		return start
	}
	return start + ":" + end
}

func sortRegions(regions []FormulaRegion) {
	sort.SliceStable(regions, func(i, j int) bool {
		if regions[i].startRow != regions[j].startRow {
			return regions[i].startRow < regions[j].startRow
		}
		if regions[i].col != regions[j].col {
			return regions[i].col < regions[j].col
		}
		return regions[i].Range < regions[j].Range
	})
}
