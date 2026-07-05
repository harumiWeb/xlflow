package formulas

import (
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/formula"
	"github.com/xuri/excelize/v2"
)

func BuildRegions(cells []FormulaCell) []FormulaRegion {
	logical := logicalFormulaCells(cells)
	if len(logical) == 0 {
		return nil
	}
	regions := buildFormulaRegions(logical)
	sortRegionsByColumn(regions)
	markOutliers(regions)
	sortRegions(regions)
	finalizeRegionIntelligence(regions)
	finalizeStorageGroupCounts(regions)
	return regions
}

func logicalFormulaCells(cells []FormulaCell) []FormulaCell {
	sortCells(cells)
	byCell := map[string]FormulaCell{}
	sharedAnchors := map[string]FormulaCell{}
	sharedAnchorSeen := map[string]bool{}

	for _, cell := range cells {
		if cell.Kind == "shared" && cell.Formula != "" && cell.SharedIndex != "" {
			sharedAnchors[cell.SharedIndex] = cell
		}
	}

	for _, cell := range cells {
		if cell.Kind != "shared" {
			cell.StorageKind = "normal"
			cell.priority = 30
			putFormulaCell(byCell, cell)
			continue
		}
		switch {
		case cell.Formula != "" && cell.SharedRef != "":
			expandSharedFormula(byCell, cell)
			sharedAnchorSeen[cell.SharedIndex] = true
		case cell.Formula != "":
			putFormulaCell(byCell, failedFormulaCell(cell, "shared_formula_malformed_ref", 30))
		case cell.SharedIndex != "" && sharedAnchorSeen[cell.SharedIndex]:
			continue
		default:
			if _, ok := sharedAnchors[cell.SharedIndex]; ok {
				continue
			}
			putFormulaCell(byCell, failedFormulaCell(cell, "shared_formula_missing_anchor", 10))
		}
	}

	logical := make([]FormulaCell, 0, len(byCell))
	for _, cell := range byCell {
		logical = append(logical, cell)
	}
	sortCellsByColumn(logical)
	return logical
}

func putFormulaCell(cells map[string]FormulaCell, cell FormulaCell) {
	if cell.Cell == "" {
		return
	}
	if cell.StorageKind == "" {
		cell.StorageKind = formulaStorageKind(cell)
	}
	if existing, ok := cells[cell.Cell]; ok && existing.priority > cell.priority {
		return
	}
	cells[cell.Cell] = cell
}

func expandSharedFormula(cells map[string]FormulaCell, anchor FormulaCell) {
	start, end, _, ok := rangeInfo(anchor.SharedRef)
	if !ok {
		putFormulaCell(cells, failedFormulaCell(anchor, "shared_formula_malformed_ref", 30))
		return
	}
	summary := normalizeFormula(anchor.Formula, anchor.Row, anchor.Col)
	groupID := "shared:" + anchor.SharedIndex + ":" + anchor.SharedRef
	for row := start.row; row <= end.row; row++ {
		for col := start.col; col <= end.col; col++ {
			cellName, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				continue
			}
			formulaText := anchor.Formula
			if row != anchor.Row || col != anchor.Col {
				formulaText = translateSharedFormula(anchor.Formula, formula.CellRef{Row: anchor.Row, Col: anchor.Col}, formula.CellRef{Row: row, Col: col})
			}
			priority := 20
			if row == anchor.Row && col == anchor.Col {
				priority = 25
			}
			putFormulaCell(cells, FormulaCell{
				Cell:         cellName,
				Row:          row,
				Col:          col,
				Kind:         "formula",
				Formula:      formulaText,
				FormulaR1C1:  summary.formulaR1C1,
				ParseStatus:  string(summary.status),
				Features:     summary.features,
				StorageKind:  "shared",
				StorageGroup: groupID,
				priority:     priority,
			})
		}
	}
}

func failedFormulaCell(cell FormulaCell, feature string, priority int) FormulaCell {
	cell.Kind = "formula"
	cell.ParseStatus = string(formula.ParseStatusFailed)
	cell.Features = []string{feature}
	cell.StorageKind = formulaStorageKind(cell)
	cell.priority = priority
	return cell
}

func formulaStorageKind(cell FormulaCell) string {
	if cell.StorageKind != "" {
		return cell.StorageKind
	}
	if cell.Kind == "shared" || cell.SharedIndex != "" {
		return "shared"
	}
	if cell.Kind == "" || cell.Kind == "formula" {
		return "normal"
	}
	return cell.Kind
}

func buildFormulaRegions(cells []FormulaCell) []FormulaRegion {
	sortCellsByColumn(cells)
	regions := []FormulaRegion{}
	var current *FormulaRegion
	for _, cell := range cells {
		region := singleFormulaRegion(cell)
		if current != nil && canExtend(*current, region) {
			extendRegion(current, region)
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
	return regions
}

func singleFormulaRegion(cell FormulaCell) FormulaRegion {
	summary := normalizeCell(cell)
	region := FormulaRegion{
		Range:          cell.Cell,
		ExampleCell:    cell.Cell,
		ExampleFormula: cell.Formula,
		Count:          1,
		ParseStatus:    string(summary.status),
		Features:       summary.features,
		StorageKinds:   []string{formulaStorageKind(cell)},
		startRow:       cell.Row,
		endRow:         cell.Row,
		col:            cell.Col,
		storageGroups:  1,
	}
	if cell.StorageGroup != "" {
		region.storageGroupIDs = map[string]bool{cell.StorageGroup: true}
	}
	if summary.status == formula.ParseStatusOK {
		region.FormulaR1C1 = summary.formulaR1C1
	} else {
		region.Formula = summary.rawFormula
	}
	region.key = buildRegionKey(region)
	return region
}

func extendRegion(current *FormulaRegion, next FormulaRegion) {
	current.endRow = next.endRow
	current.Count += next.Count
	current.Range = renderRange(current.col, current.startRow, current.col, current.endRow)
	current.StorageKinds = mergeStringSets(current.StorageKinds, next.StorageKinds)
	if len(current.storageGroupIDs) == 0 && len(next.storageGroupIDs) == 0 {
		current.storageGroups = 1
		return
	}
	if current.storageGroupIDs == nil {
		current.storageGroupIDs = map[string]bool{}
	}
	for group := range next.storageGroupIDs {
		current.storageGroupIDs[group] = true
	}
	current.storageGroups = len(current.storageGroupIDs)
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

func normalizeCell(cell FormulaCell) normalizeSummary {
	if cell.ParseStatus != "" {
		return normalizeSummary{
			status:      formula.ParseStatus(cell.ParseStatus),
			formulaR1C1: cell.FormulaR1C1,
			rawFormula:  cell.Formula,
			features:    cell.Features,
		}
	}
	return normalizeFormula(cell.Formula, cell.Row, cell.Col)
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
	return left.col == right.col &&
		left.endRow+1 == right.startRow &&
		left.key == right.key
}

func markOutliers(regions []FormulaRegion) {
	for i := 1; i+1 < len(regions); i++ {
		prev, current, next := regions[i-1], &regions[i], regions[i+1]
		if current.Count != 1 {
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

func buildRegionKey(region FormulaRegion) regionKey {
	features := append([]string{}, region.Features...)
	sort.Strings(features)
	return regionKey{
		FormulaR1C1: region.FormulaR1C1,
		Formula:     region.Formula,
		ParseStatus: region.ParseStatus,
		Features:    strings.Join(features, "\x00"),
	}
}

func mergeStringSets(left, right []string) []string {
	seen := map[string]bool{}
	var merged []string
	for _, value := range append(left, right...) {
		if value == "" || seen[value] {
			continue
		}
		seen[value] = true
		merged = append(merged, value)
	}
	sort.Strings(merged)
	return merged
}

func finalizeStorageGroupCounts(regions []FormulaRegion) {
	for i := range regions {
		if len(regions[i].storageGroupIDs) > 0 {
			regions[i].storageGroups = len(regions[i].storageGroupIDs)
		}
		if regions[i].storageGroups > 1 {
			regions[i].StorageGroupCount = regions[i].storageGroups
		}
		if regions[i].storageGroups == 1 && len(regions[i].StorageKinds) == 1 && regions[i].StorageKinds[0] == "normal" {
			regions[i].StorageKinds = nil
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

func translateSharedFormula(value string, from, to formula.CellRef) string {
	result := formula.NormalizeA1ToR1C1Pattern(value, formula.NormalizeOptions{BaseCell: from})
	if len(result.References) == 0 {
		return value
	}
	shifted := value
	deltaRow := to.Row - from.Row
	deltaCol := to.Col - from.Col
	refs := append([]formula.Reference{}, result.References...)
	sort.Slice(refs, func(i, j int) bool {
		return refs[i].Span.Start > refs[j].Span.Start
	})
	for _, ref := range refs {
		replacement, ok := renderShiftedReference(ref, deltaRow, deltaCol)
		if !ok || ref.Span.Start < 0 || ref.Span.End > len(shifted) || ref.Span.Start > ref.Span.End {
			continue
		}
		shifted = shifted[:ref.Span.Start] + replacement + shifted[ref.Span.End:]
	}
	return shifted
}

func renderShiftedReference(ref formula.Reference, deltaRow, deltaCol int) (string, bool) {
	prefix := ""
	if ref.Sheet != "" {
		prefix = ref.Sheet + "!"
	}
	start := shiftEndpoint(ref.Start, deltaRow, deltaCol)
	switch ref.Kind {
	case formula.ReferenceKindCell:
		return prefix + renderA1Endpoint(start, true, true), true
	case formula.ReferenceKindRange:
		if ref.End == nil {
			return "", false
		}
		end := shiftEndpoint(*ref.End, deltaRow, deltaCol)
		return prefix + renderA1Endpoint(start, true, true) + ":" + renderA1Endpoint(end, true, true), true
	case formula.ReferenceKindColumnRange:
		if ref.End == nil {
			return "", false
		}
		end := shiftEndpoint(*ref.End, deltaRow, deltaCol)
		return prefix + renderA1Endpoint(start, false, true) + ":" + renderA1Endpoint(end, false, true), true
	case formula.ReferenceKindRowRange:
		if ref.End == nil {
			return "", false
		}
		end := shiftEndpoint(*ref.End, deltaRow, deltaCol)
		return prefix + renderA1Endpoint(start, true, false) + ":" + renderA1Endpoint(end, true, false), true
	default:
		return "", false
	}
}

func shiftEndpoint(endpoint formula.RefEndpoint, deltaRow, deltaCol int) formula.RefEndpoint {
	if endpoint.Row > 0 && !endpoint.RowAbs {
		endpoint.Row += deltaRow
	}
	if endpoint.Col > 0 && !endpoint.ColAbs {
		endpoint.Col += deltaCol
	}
	return endpoint
}

func renderA1Endpoint(endpoint formula.RefEndpoint, includeRow, includeCol bool) string {
	var b strings.Builder
	if includeCol {
		if endpoint.ColAbs {
			b.WriteByte('$')
		}
		name, _ := excelize.ColumnNumberToName(endpoint.Col)
		b.WriteString(name)
	}
	if includeRow {
		if endpoint.RowAbs {
			b.WriteByte('$')
		}
		b.WriteString(strconv.Itoa(endpoint.Row))
	}
	return b.String()
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

func sortRegionsByColumn(regions []FormulaRegion) {
	sort.SliceStable(regions, func(i, j int) bool {
		if regions[i].col != regions[j].col {
			return regions[i].col < regions[j].col
		}
		if regions[i].startRow != regions[j].startRow {
			return regions[i].startRow < regions[j].startRow
		}
		return regions[i].Range < regions[j].Range
	})
}
